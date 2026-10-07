"""Utilidades de mantenimiento del buzón por IMAP: liberar espacio borrando
definitivamente correos ya procesados, y aplicar etiquetas de Gmail.

Por qué "copiar a Papelera + purgar desde ahí" y no un simple
STORE(\\Deleted) + EXPUNGE en la carpeta original: en Gmail, hacer eso en
INBOX o en Enviados NO borra el correo — solo le quita esa etiqueta, y el
mensaje sigue existiendo (y ocupando espacio) en "Todos los mensajes". Para
liberar espacio de verdad hay que moverlo a la Papelera y purgarlo también
ahí, igual que "Eliminar para siempre" en la interfaz web.

Por qué se purga solo lo que el propio servidor confirma que acabamos de
mover (vía la respuesta COPYUID, RFC 4315): la Papelera puede tener otros
correos del usuario que no tienen nada que ver con esta limpieza. Si no
podemos confirmar con certeza el UID que le tocó al mensaje dentro de la
Papelera, lo dejamos ahí sin purgar (se autoborra solo en ~30 días) en vez
de arriesgarnos a purgar algo que no corresponde.
"""

import base64
import email
import imaplib
import json
import re
from email.utils import getaddresses
from pathlib import Path

import pandas as pd

from src.config import Config

BATCH_SIZE = 200

_LIST_RE = re.compile(r'^\((?P<flags>[^)]*)\)\s+"(?P<delim>[^"]*)"\s+(?P<name>.+)$')


def imap_utf7_encode(s: str) -> str:
    """Codifica en UTF-7 modificado (RFC 3501), necesario para nombres de
    buzón/etiquetas de IMAP con caracteres no ASCII (ej. 'Respondió')."""
    if s.isascii():
        return s.replace("&", "&-")

    result: list[str] = []
    b64_chunk: list[str] = []

    def flush_b64chunk():
        if not b64_chunk:
            return
        raw = "".join(b64_chunk).encode("utf-16-be")
        b64 = base64.b64encode(raw).decode("ascii").rstrip("=").replace("/", ",")
        result.append("&" + b64 + "-")
        b64_chunk.clear()

    for ch in s:
        if ch == "&":
            flush_b64chunk()
            result.append("&-")
        elif 0x20 <= ord(ch) <= 0x7E:
            flush_b64chunk()
            result.append(ch)
        else:
            b64_chunk.append(ch)
    flush_b64chunk()
    return "".join(result)


def find_special_folder(
    mail: imaplib.IMAP4_SSL, special_use: str, fallback_names: list[str]
) -> str:
    """Busca el buzón marcado con el atributo especial (\\Trash, \\Sent,
    RFC 6154). Si el servidor no lo informa, prueba nombres comunes."""
    status, folders = mail.list()
    if status == "OK" and folders:
        for raw in folders:
            line = raw.decode(errors="ignore") if isinstance(raw, bytes) else raw
            match = _LIST_RE.match(line)
            if not match:
                continue
            if special_use.lower() in match.group("flags").lower():
                name = match.group("name").strip()
                if name.startswith('"') and name.endswith('"'):
                    name = name[1:-1]
                return name

    for name in fallback_names:
        try:
            status, _ = mail.select(f'"{name}"', readonly=True)
        except imaplib.IMAP4.error:
            status = "NO"
        if status == "OK":
            return name

    raise RuntimeError(
        f"No se pudo encontrar el buzón especial {special_use}. Probé: {fallback_names}"
    )


def _uid_batches(uids: list[int], size: int = BATCH_SIZE):
    uids = sorted(set(int(u) for u in uids))
    for i in range(0, len(uids), size):
        yield uids[i : i + size]


def _parse_copyuid(raw: bytes | None) -> list[int]:
    """Convierte la respuesta COPYUID cruda (ej. b'2 1,2 4537,4538': UIDVALIDITY,
    UIDs de origen, UIDs de destino) en la lista de UIDs de destino.

    imaplib NO devuelve este dato en el resultado de `mail.uid('COPY', ...)`
    (ahí llega `[None]`) — lo guarda aparte, en
    `mail.untagged_responses['COPYUID']`. Hay que leerlo de ahí."""
    dest_uids: list[int] = []
    if not raw:
        return dest_uids
    parts = raw.decode().split()
    if len(parts) < 3:
        return dest_uids
    for part in parts[2].split(","):
        if ":" in part:
            start, end = part.split(":")
            dest_uids.extend(range(int(start), int(end) + 1))
        else:
            dest_uids.append(int(part))
    return dest_uids


def permanently_delete(
    mail: imaplib.IMAP4_SSL, folder: str, uids: list[int], trash_folder: str
) -> tuple[int, int]:
    """Borra definitivamente los mensajes (por UID) del buzón `folder`.

    Devuelve (movidos_a_papelera, purgados_definitivamente). El segundo
    número puede ser menor que el primero si el servidor no confirmó el UID
    de destino para alguno de los lotes: esos quedan a salvo en la Papelera
    en vez de arriesgarse a purgar de más."""
    if not uids:
        return 0, 0

    status, _ = mail.select(f'"{folder}"')
    if status != "OK":
        raise RuntimeError(f"No se pudo abrir el buzón {folder!r}")

    moved = 0
    purge_targets: list[int] = []

    for batch in _uid_batches(uids):
        uid_set = ",".join(str(u) for u in batch)

        mail.untagged_responses.pop("COPYUID", None)
        status, _ = mail.uid("COPY", uid_set, f'"{trash_folder}"')
        if status != "OK":
            continue

        copyuid_raw = mail.untagged_responses.get("COPYUID")
        dest_uids = _parse_copyuid(copyuid_raw[-1] if copyuid_raw else None)

        mail.uid("STORE", uid_set, "+FLAGS", "(\\Deleted)")
        mail.expunge()
        moved += len(batch)

        if len(dest_uids) == len(batch):
            purge_targets.extend(dest_uids)

    purged = 0
    if purge_targets:
        mail.select(f'"{trash_folder}"')
        for i in range(0, len(purge_targets), BATCH_SIZE):
            chunk = ",".join(str(u) for u in purge_targets[i : i + BATCH_SIZE])
            mail.uid("STORE", chunk, "+FLAGS", "(\\Deleted)")
        mail.expunge()
        purged = len(purge_targets)

    return moved, purged


def label_message(mail: imaplib.IMAP4_SSL, uid, label: str) -> None:
    """Aplica una etiqueta de Gmail a un mensaje del buzón ya seleccionado."""
    encoded = imap_utf7_encode(label)
    mail.uid("STORE", str(uid), "+X-GM-LABELS", f'("{encoded}")')


def extract_recipients_from_header_blob(header_blob: bytes) -> set[str]:
    """Recibe el bloque crudo de un FETCH BODY[HEADER.FIELDS (TO)] (ej.
    b'To: a@b.com\\r\\n\\r\\n') y devuelve el conjunto de correos normalizados
    que aparecen en el campo To."""
    header_msg = email.message_from_bytes(header_blob)
    to_values = header_msg.get_all("To") or []
    return {addr.strip().lower() for _, addr in getaddresses(to_values) if addr}


def _sent_state_file(config: Config) -> Path:
    return Path(config.log_file).parent / "sent_cleanup_state.json"


def cleanup_sent_folder(config: Config, auto_confirm: bool = False) -> tuple[int, int]:
    """Borra definitivamente, de la carpeta de Enviados, las copias de los
    correos que ya están confirmados como 'enviado' en send_log.xlsx.

    Por qué es seguro borrarlos ahí mismo, sin esperar nada: tu copia en
    Enviados no tiene ninguna influencia sobre si el correo llegó o rebotó
    — eso se decide aparte, leyendo los avisos que llegan a la bandeja de
    entrada (`check_bounces`). Es solo tu copia local del envío, y suele ser
    la que más espacio ocupa por llevar el folleto adjunto.

    Recuerda el UID hasta donde ya revisó (en `sent_cleanup_state.json`)
    para no releer toda la carpeta en cada corrida; los mensajes ya borrados
    nunca vuelven a aparecer, así que no hace falta revisarlos de nuevo."""
    log_df = pd.read_excel(config.log_file, sheet_name="Sheet1", engine="openpyxl")
    sent_emails = set(
        log_df.loc[log_df["Estado"] == "enviado", "Correo"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
    )
    if not sent_emails:
        print("No hay correos marcados como 'enviado' en el log todavía.")
        return 0, 0

    state_path = _sent_state_file(config)
    last_uid = 0
    if state_path.exists():
        try:
            last_uid = int(json.loads(state_path.read_text()).get("last_uid", 0))
        except Exception:
            last_uid = 0

    with imaplib.IMAP4_SSL(config.imap_host) as mail:
        mail.login(config.smtp_user, config.smtp_password)

        sent_folder = find_special_folder(
            mail, "\\Sent", ["[Gmail]/Sent Mail", "[Gmail]/Enviados", "Sent"]
        )
        trash_folder = find_special_folder(
            mail, "\\Trash", ["[Gmail]/Trash", "[Gmail]/Papelera", "Trash"]
        )

        mail.select(f'"{sent_folder}"')
        criteria = f"UID {last_uid + 1}:*" if last_uid else "ALL"
        status, data = mail.uid("search", None, criteria)
        if status != "OK" or not data or not data[0]:
            print("No hay correos nuevos en Enviados para revisar.")
            return 0, 0

        all_uids = sorted(int(u) for u in data[0].split() if int(u) > last_uid)
        if not all_uids:
            print("No hay correos nuevos en Enviados para revisar.")
            return 0, 0

        max_uid_seen = max(all_uids)
        to_delete: list[int] = []

        for i in range(0, len(all_uids), BATCH_SIZE):
            chunk = all_uids[i : i + BATCH_SIZE]
            uid_set = ",".join(str(u) for u in chunk)
            status, msg_data = mail.uid(
                "fetch", uid_set, "(BODY.PEEK[HEADER.FIELDS (TO)])"
            )
            if status != "OK" or not msg_data:
                continue

            current_uid = None
            for part in msg_data:
                if isinstance(part, tuple):
                    header_line = part[0].decode(errors="ignore")
                    uid_match = re.search(r"UID (\d+)", header_line)
                    if uid_match:
                        current_uid = int(uid_match.group(1))
                    recipients = extract_recipients_from_header_blob(part[1])
                    if current_uid and recipients & sent_emails:
                        to_delete.append(current_uid)

        moved = purged = 0
        if to_delete:
            if not auto_confirm:
                respuesta = (
                    input(
                        f"Se van a borrar definitivamente {len(to_delete)} correo(s) de "
                        f"'{sent_folder}' (copias tuyas ya confirmadas como enviadas). "
                        f"¿Continuar? (s/n): "
                    )
                    .strip()
                    .lower()
                )
                if respuesta not in {"s", "si", "y", "yes"}:
                    print(
                        "Cancelado. No se borró nada; la próxima corrida volverá a "
                        "considerar estos mismos correos."
                    )
                    return 0, 0
            moved, purged = permanently_delete(mail, sent_folder, to_delete, trash_folder)

        state_path.parent.mkdir(parents=True, exist_ok=True)
        state_path.write_text(json.dumps({"last_uid": max_uid_seen}))

    print(f"Revisados {len(all_uids)} correo(s) nuevos en '{sent_folder}'.")
    print(f"Movidos a Papelera: {moved}, purgados definitivamente: {purged}.")
    return moved, purged
