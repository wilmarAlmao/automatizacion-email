"""Módulo que detecta rebotes (bounces), revisando por IMAP la bandeja de
entrada de la cuenta que envía los correos.

Por qué por IMAP y no en el momento del envío: Gmail casi siempre ACEPTA el
correo apenas lo mandamos (por eso `src/main.py` imprime "enviado"), y recién
minutos u horas después el servidor del destinatario lo rechaza. Ese rechazo
llega como un correo de aviso ("Delivery Status Notification (Failure)",
"Mail Delivery Subsystem", etc.) a la bandeja de la propia cuenta que mandó el
correo original. Este módulo busca esos avisos y deja el resultado en dos
lugares:

- `data/bounces_log.xlsx`: historial de cada rebote confirmado.
- `data/send_log.xlsx`: la fila de ese cliente pasa a Estado = "rebotado".

Cómo identifica al destinatario que rebotó: los avisos de rebote siguen un
formato estándar (RFC 3464) que trae una parte "message/delivery-status" con
campos de máquina como `Final-Recipient` y `Action`. Leemos esos campos
directamente en vez de "adivinar" buscando cualquier texto con forma de
correo en el cuerpo del mensaje — así evitamos confundir un correo real con,
por ejemplo, un identificador interno de Google que también tiene forma de
correo (algo así como `6a95a78f...@mx.google.com`).

También distinguimos el tipo de aviso:
- `Action: failed` (o asuntos "Undelivered.../Undeliverable...") = rebote
  confirmado, se procesa.
- `Action: delayed` (asunto con "(Delay)") = aviso de RETRASO temporal, no un
  rebote — Gmail sigue intentando entregarlo. Se ignora: no es un rebote real
  todavía (puede que después llegue o que después sí rebote de verdad).
"""

import email
import imaplib
import json
import re
from datetime import datetime
from email.header import decode_header, make_header
from email.utils import parseaddr
from pathlib import Path
from typing import Optional

import pandas as pd

from src.config import Config
from src.mailbox_cleanup import find_special_folder, label_message, permanently_delete

REPLY_LABEL = "Respondió"

BOUNCE_SENDER_KEYWORDS = ("mailer-daemon", "postmaster", "mail delivery subsystem")
BOUNCE_SUBJECT_KEYWORDS = (
    "delivery status notification",
    "undelivered mail",
    "undeliverable",
    "returned to sender",
    "failure notice",
    "mail delivery failed",
    "no pudo ser entregado",
    "no se pudo entregar",
)
DELAY_KEYWORDS = ("delay", "retraso", "retrasado")

# Direcciones que no son destinatarios reales: identificadores internos que
# Google incluye en el cuerpo de sus propios avisos.
IGNORED_DOMAINS = ("mx.google.com",)

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")


def _state_file(config: Config) -> Path:
    return Path(config.log_file).parent / "bounce_checker_state.json"


def _load_last_uid(config: Config) -> int:
    path = _state_file(config)
    if path.exists():
        try:
            return int(json.loads(path.read_text()).get("last_uid", 0))
        except Exception:
            return 0
    return 0


def _save_last_uid(config: Config, uid: int) -> None:
    path = _state_file(config)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"last_uid": uid}))


def _decode_subject(msg: "email.message.Message") -> str:
    """Decodifica un asunto que puede venir codificado en MIME
    (ej. '=?iso-8859-1?Q?Undeliverable...?='), para que se lea normal."""
    raw = msg.get("Subject", "")
    try:
        return str(make_header(decode_header(raw)))
    except Exception:
        return raw


def _is_real_recipient(address: str) -> bool:
    address = address.strip().rstrip(".").lower()
    if not address or "@" not in address:
        return False
    domain = address.split("@", 1)[1]
    return domain not in IGNORED_DOMAINS


def _looks_like_bounce(msg: "email.message.Message", subject: str) -> bool:
    from_header = (msg.get("From") or "").lower()
    subject_lower = subject.lower()
    if any(keyword in from_header for keyword in BOUNCE_SENDER_KEYWORDS):
        return True
    if any(keyword in subject_lower for keyword in BOUNCE_SUBJECT_KEYWORDS):
        return True
    return False


def _parse_delivery_status(msg: "email.message.Message") -> list[dict]:
    """Busca la parte 'message/delivery-status' (formato estándar RFC 3464)
    y devuelve, por cada destinatario reportado, su acción (failed, delayed,
    delivered, relayed...) y su dirección. Este es el método confiable —
    lee campos de máquina en vez de adivinar texto libre."""
    results = []
    if not msg.is_multipart():
        return results

    for part in msg.walk():
        if part.get_content_type() != "message/delivery-status":
            continue

        payload = part.get_payload()
        blocks = payload if isinstance(payload, list) else [payload]

        for block in blocks:
            action = (block.get("Action") or "").strip().lower()
            recipient_raw = block.get("Final-Recipient") or block.get("Original-Recipient") or ""
            # Formato típico: "rfc822; alguien@dominio.com"
            recipient = recipient_raw.split(";", 1)[-1].strip() if recipient_raw else ""
            recipient = recipient.strip().rstrip(".").lower()

            if recipient:
                results.append({"action": action, "recipient": recipient})

    return results


def _extract_failed_recipient_fallback(
    msg: "email.message.Message", known_emails: set, sender_email: str
) -> Optional[str]:
    """Heurístico de respaldo, solo para cuando el aviso NO trae la parte
    'message/delivery-status' estructurada (poco común, pero pasa con
    algunos proveedores). Busca en el texto un correo que sí reconozcamos
    de nuestra propia lista de envíos."""
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            try:
                payload = part.get_payload(decode=True)
                if payload:
                    parts.append(payload.decode(errors="ignore"))
            except Exception:
                continue
    else:
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                parts.append(payload.decode(errors="ignore"))
        except Exception:
            pass

    body_text = "\n".join(parts)
    found = {e.lower().rstrip(".") for e in EMAIL_REGEX.findall(body_text)}
    found.discard(sender_email.lower())
    found = {e for e in found if _is_real_recipient(e)}

    matches = found & known_emails
    return next(iter(matches), None)


def check_bounces(config: Config) -> tuple[list[dict], int]:
    """Revisa la bandeja de entrada por IMAP y:

    1. Devuelve la lista de rebotes CONFIRMADOS (no avisos de retraso)
       encontrados desde la última vez que se corrió esta función.
    2. Borra definitivamente de la bandeja el correo de aviso de cada rebote
       confirmado (ya quedó guardado en `bounced`, así que no se pierde el
       dato — ver `log_bounces`).
    3. A cualquier otro correo entrante que NO sea un aviso de rebote, y cuyo
       remitente sea alguien a quien le mandamos la campaña, le pone la
       etiqueta `REPLY_LABEL` en Gmail (es una respuesta real de un cliente).

    Devuelve (rebotes_confirmados, cantidad_de_respuestas_etiquetadas)."""
    log_df = pd.read_excel(config.log_file, sheet_name="Sheet1", engine="openpyxl")
    known_emails = set(log_df["Correo"].dropna().astype(str).str.strip().str.lower())

    bounced: list[dict] = []
    bounce_uids: list[int] = []
    replied_count = 0
    last_uid = _load_last_uid(config)
    max_uid_seen = last_uid

    with imaplib.IMAP4_SSL(config.imap_host) as mail:
        mail.login(config.smtp_user, config.smtp_password)
        mail.select("INBOX")

        criteria = f"UID {last_uid + 1}:*" if last_uid else "ALL"
        status, data = mail.uid("search", None, criteria)
        if status != "OK" or not data or not data[0]:
            return bounced, replied_count

        for uid_bytes in data[0].split():
            uid_int = int(uid_bytes)
            if uid_int <= last_uid:
                continue
            max_uid_seen = max(max_uid_seen, uid_int)

            status, msg_data = mail.uid("fetch", uid_bytes, "(RFC822)")
            if status != "OK" or not msg_data or msg_data[0] is None:
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            subject = _decode_subject(msg)

            if not _looks_like_bounce(msg, subject):
                sender_email = parseaddr(msg.get("From", ""))[1].strip().lower()
                if sender_email in known_emails:
                    label_message(mail, uid_int, REPLY_LABEL)
                    replied_count += 1
                continue

            dsn_entries = _parse_delivery_status(msg)
            found_confirmed_bounce = False

            if dsn_entries:
                for entry in dsn_entries:
                    if entry["action"] != "failed":
                        continue  # "delayed", "delivered", "relayed": no es un rebote confirmado
                    if not _is_real_recipient(entry["recipient"]):
                        continue
                    bounced.append(
                        {
                            "Correo": entry["recipient"],
                            "Fecha_deteccion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Asunto_rebote": subject,
                        }
                    )
                    found_confirmed_bounce = True
            else:
                # Sin parte delivery-status estructurada. Si el asunto indica
                # claramente un retraso (no un rebote), lo ignoramos.
                if any(keyword in subject.lower() for keyword in DELAY_KEYWORDS):
                    continue
                failed_email = _extract_failed_recipient_fallback(
                    msg, known_emails, config.sender_email
                )
                if failed_email:
                    bounced.append(
                        {
                            "Correo": failed_email,
                            "Fecha_deteccion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Asunto_rebote": subject,
                        }
                    )
                    found_confirmed_bounce = True

            if found_confirmed_bounce:
                bounce_uids.append(uid_int)

        if bounce_uids:
            trash_folder = find_special_folder(
                mail, "\\Trash", ["[Gmail]/Trash", "[Gmail]/Papelera", "Trash"]
            )
            moved, purged = permanently_delete(mail, "INBOX", bounce_uids, trash_folder)
            print(
                f"Avisos de rebote borrados de la bandeja: {moved} "
                f"(purgados definitivamente: {purged})."
            )

    _save_last_uid(config, max_uid_seen)
    return bounced, replied_count


def cleanup_old_bounces(config: Config, auto_confirm: bool = False) -> list[dict]:
    """Barrido de UNA SOLA VEZ para ponerse al día: revisa los correos
    VIEJOS de la bandeja (los que `check_bounces` ya procesó antes de que
    este programa supiera borrarlos) y borra definitivamente los que sean
    avisos de rebote confirmados. De aquí en adelante `check_bounces` ya los
    borra solo en cada corrida — esto no vuelve a hacer falta después."""
    log_df = pd.read_excel(config.log_file, sheet_name="Sheet1", engine="openpyxl")
    known_emails = set(log_df["Correo"].dropna().astype(str).str.strip().str.lower())

    last_uid = _load_last_uid(config)
    found: list[dict] = []
    bounce_uids: list[int] = []

    with imaplib.IMAP4_SSL(config.imap_host) as mail:
        mail.login(config.smtp_user, config.smtp_password)
        mail.select("INBOX")

        criteria = f"UID 1:{last_uid}" if last_uid else "ALL"
        status, data = mail.uid("search", None, criteria)
        if status != "OK" or not data or not data[0]:
            print("No hay correos antiguos que revisar.")
            return found

        uid_list = [int(u) for u in data[0].split()]
        print(f"Revisando {len(uid_list)} correo(s) antiguo(s) de la bandeja...")

        for uid_int in uid_list:
            status, msg_data = mail.uid("fetch", str(uid_int), "(RFC822)")
            if status != "OK" or not msg_data or msg_data[0] is None:
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            subject = _decode_subject(msg)

            if not _looks_like_bounce(msg, subject):
                continue

            dsn_entries = _parse_delivery_status(msg)
            confirmed = False

            if dsn_entries:
                for entry in dsn_entries:
                    if entry["action"] != "failed":
                        continue
                    if not _is_real_recipient(entry["recipient"]):
                        continue
                    found.append(
                        {
                            "Correo": entry["recipient"],
                            "Fecha_deteccion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Asunto_rebote": subject,
                        }
                    )
                    confirmed = True
            else:
                if any(keyword in subject.lower() for keyword in DELAY_KEYWORDS):
                    continue
                failed_email = _extract_failed_recipient_fallback(
                    msg, known_emails, config.sender_email
                )
                if failed_email:
                    found.append(
                        {
                            "Correo": failed_email,
                            "Fecha_deteccion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Asunto_rebote": subject,
                        }
                    )
                    confirmed = True

            if confirmed:
                bounce_uids.append(uid_int)

        if not bounce_uids:
            print("No se encontraron avisos de rebote viejos pendientes de borrar.")
            return found

        print(
            f"Se encontraron {len(bounce_uids)} aviso(s) de rebote viejo(s) para borrar definitivamente."
        )
        if not auto_confirm:
            respuesta = input("¿Continuar con el borrado? (s/n): ").strip().lower()
            if respuesta not in {"s", "si", "y", "yes"}:
                print("Cancelado. No se borró nada.")
                return found

        trash_folder = find_special_folder(
            mail, "\\Trash", ["[Gmail]/Trash", "[Gmail]/Papelera", "Trash"]
        )
        moved, purged = permanently_delete(mail, "INBOX", bounce_uids, trash_folder)
        print(f"Avisos de rebote viejos borrados: {moved} (purgados definitivamente: {purged}).")

    return found


def log_bounces(bounced: list[dict], file_path: str) -> None:
    """Guarda cada rebote confirmado en un Excel aparte, como historial.
    Evita duplicar una dirección que ya esté registrada."""
    if not bounced:
        return

    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    if path.exists():
        existing = pd.read_excel(path, engine="openpyxl")
    else:
        existing = pd.DataFrame(columns=["Correo", "Fecha_deteccion", "Asunto_rebote"])

    combined = pd.concat([existing, pd.DataFrame(bounced)], ignore_index=True)
    combined["Correo"] = combined["Correo"].astype(str).str.strip().str.rstrip(".").str.lower()
    combined = combined.sort_values("Fecha_deteccion").drop_duplicates(
        subset=["Correo"], keep="first"
    )
    combined.to_excel(path, index=False, engine="openpyxl")


def update_send_log(bounced: list[dict], log_file: str) -> int:
    """Marca como 'rebotado' en data/send_log.xlsx las filas de los correos
    detectados. Devuelve cuántas filas se actualizaron."""
    if not bounced:
        return 0

    df = pd.read_excel(log_file, sheet_name="Sheet1", engine="openpyxl")
    bounced_emails = {b["Correo"].strip().rstrip(".").lower() for b in bounced}
    if not bounced_emails:
        return 0

    correo_normalizado = df["Correo"].astype(str).str.strip().str.rstrip(".").str.lower()
    mask = correo_normalizado.isin(bounced_emails)
    df.loc[mask, "Estado"] = "rebotado"
    with pd.ExcelWriter(log_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
    return int(mask.sum())


SENT_SUBJECT_PREFIX = "Novedades de Tu Empresa para "


def purge_trash_backlog(config: Config, auto_confirm: bool = False) -> tuple[int, int]:
    """Barrido de UNA SOLA VEZ para arreglar el atasco que dejó un bug
    anterior: `cleanup_sent.py` y `cleanup_old_bounces.py` movían los
    correos a la Papelera, pero por un error nunca los purgaban de verdad
    (ya corregido). Esos correos siguen ahí ocupando espacio.

    En vez de vaciar toda la Papelera (podrías tener ahí algo tuyo, propio,
    sin relación con la campaña), recorre lo que hay y purga SOLO lo que
    reconoce como propio de esta herramienta: copias de correos enviados de
    la campaña (mismo remitente y mismo prefijo de asunto que usa
    `src/main.py`) o avisos de rebote. Todo lo demás queda intacto.

    Devuelve (purgados, total_revisados_en_papelera)."""
    with imaplib.IMAP4_SSL(config.imap_host) as mail:
        mail.login(config.smtp_user, config.smtp_password)

        trash_folder = find_special_folder(
            mail, "\\Trash", ["[Gmail]/Trash", "[Gmail]/Papelera", "Trash"]
        )
        mail.select(f'"{trash_folder}"')

        status, data = mail.uid("search", None, "ALL")
        if status != "OK" or not data or not data[0]:
            print("La Papelera está vacía.")
            return 0, 0

        all_uids = [int(u) for u in data[0].split()]
        print(f"Revisando {len(all_uids)} mensaje(s) en la Papelera...")

        to_purge: list[int] = []
        sender_self = config.sender_email.strip().lower()

        for i in range(0, len(all_uids), 200):
            chunk = all_uids[i : i + 200]
            uid_set = ",".join(str(u) for u in chunk)
            status, msg_data = mail.uid(
                "fetch", uid_set, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])"
            )
            if status != "OK" or not msg_data:
                continue

            current_uid = None
            for part in msg_data:
                if not isinstance(part, tuple):
                    continue
                header_line = part[0].decode(errors="ignore")
                uid_match = re.search(r"UID (\d+)", header_line)
                if uid_match:
                    current_uid = int(uid_match.group(1))

                header_msg = email.message_from_bytes(part[1])
                subject = _decode_subject(header_msg)
                from_addr = parseaddr(header_msg.get("From", ""))[1].strip().lower()

                is_bounce = _looks_like_bounce(header_msg, subject)
                is_own_sent_copy = from_addr == sender_self and subject.startswith(
                    SENT_SUBJECT_PREFIX
                )

                if current_uid and (is_bounce or is_own_sent_copy):
                    to_purge.append(current_uid)

        if not to_purge:
            print("No se encontró nada reconocible como propio de la campaña en la Papelera.")
            return 0, len(all_uids)

        print(
            f"Se identificaron {len(to_purge)} de {len(all_uids)} mensaje(s) en la Papelera "
            f"como copias de envíos o avisos de rebote de esta campaña."
        )
        if not auto_confirm:
            respuesta = (
                input(
                    "Estos se purgarán DEFINITIVAMENTE (ya están en la Papelera, esto los "
                    "quita también de ahí). El resto de la Papelera no se toca. "
                    "¿Continuar? (s/n): "
                )
                .strip()
                .lower()
            )
            if respuesta not in {"s", "si", "y", "yes"}:
                print("Cancelado. No se borró nada.")
                return 0, len(all_uids)

        uid_set_str = ",".join(str(u) for u in to_purge)
        mail.uid("STORE", uid_set_str, "+FLAGS", "(\\Deleted)")
        mail.expunge()

    print(f"Purgados definitivamente {len(to_purge)} mensaje(s) de la Papelera.")
    return len(to_purge), len(all_uids)
