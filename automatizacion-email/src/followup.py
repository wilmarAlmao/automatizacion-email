"""Correo de SEGUIMIENTO (campaña cierre de año).

Se le manda solo a quien:
- ya recibió el correo de presentación (Estado = "enviado" en los logs),
- hace al menos FOLLOWUP_DAYS días,
- NO respondió (se revisa la cuenta de Gmail por IMAP),
- no rebotó y no recibió ya este seguimiento.

Uso, desde la carpeta raíz del proyecto:
    python -m src.followup          (pide confirmación s/n)
    python -m src.followup --yes    (sin preguntar)

Con SEND_EMAILS=false es un ensayo: muestra a quién le mandaría y guarda una
vista previa del correo en data/followup_preview.html, sin enviar nada.
"""

import email
import imaplib
import sys
import time
import urllib.request
import zlib
from datetime import datetime, timedelta
from email.utils import parseaddr
from pathlib import Path
from urllib.parse import quote

import pandas as pd

from src.bounce_checker import BOUNCE_SENDER_KEYWORDS, _decode_subject
from src.config import Config, load_config
from src.email_sender import EmailSender
from src.logger import ExcelLogger
from src.mailbox_cleanup import find_special_folder
from src.main import confirm_send, count_sent_today, is_business_day, is_within_send_window
from src.template_renderer import render_followup_body, render_followup_text
from src.validators import is_valid_email

WHATSAPP_NUMBER = "00000000000"
WHATSAPP_MESSAGE = "Hola, quiero más información sobre sus servicios."

# Estados del log de seguimiento que significan "ya no hay que volver a mandarle".
FOLLOWUP_DONE_STATES = {"enviado", "rebotado", "correo inválido"}

# Dominios de correo gratuito: aquí una respuesta de alguien NO significa que
# respondió la empresa entera (cada dirección es una persona distinta).
FREE_MAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "hotmail.com", "hotmail.es", "outlook.com",
    "outlook.es", "live.com", "msn.com", "yahoo.com", "yahoo.es", "icloud.com",
    "me.com", "aol.com", "protonmail.com", "proton.me",
}

# Las respuestas automáticas ("fuera de la oficina") no cuentan como respuesta.
AUTO_REPLY_SUBJECT_KEYWORDS = (
    "respuesta automática", "respuesta automatica", "automatic reply",
    "auto reply", "autoreply", "auto-reply", "out of office", "fuera de la oficina",
    "fuera de oficina", "ausencia", "vacaciones",
)


def _normalize(address: str) -> str:
    return str(address).strip().rstrip(".").lower()


def _domain(address: str) -> str:
    return address.split("@", 1)[1] if "@" in address else ""


def _is_auto_reply(header_msg: "email.message.Message", subject: str) -> bool:
    auto_submitted = (header_msg.get("Auto-Submitted") or "").strip().lower()
    if auto_submitted and auto_submitted != "no":
        return True
    if header_msg.get("X-Autoreply") or header_msg.get("X-Autorespond"):
        return True
    subject_lower = subject.lower()
    return any(keyword in subject_lower for keyword in AUTO_REPLY_SUBJECT_KEYWORDS)


def load_replied_addresses(config: Config, since: datetime) -> set[str]:
    """Remitentes de todos los correos recibidos desde `since` (carpeta
    "Todos los mensajes" de Gmail, así se cuentan también los archivados).
    Se ignoran: la propia cuenta, avisos de rebote y respuestas automáticas."""
    sender_self = _normalize(config.sender_email)
    smtp_self = _normalize(config.smtp_user)
    replied: set[str] = set()

    with imaplib.IMAP4_SSL(config.imap_host) as mail:
        mail.login(config.smtp_user, config.smtp_password)
        all_mail = find_special_folder(
            mail, "\\All", ["[Gmail]/All Mail", "[Gmail]/Todos", "INBOX"]
        )
        mail.select(f'"{all_mail}"', readonly=True)

        # IMAP exige el mes en inglés (strftime %b depende del idioma del sistema)
        months_en = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        since_imap = f"{since.day:02d}-{months_en[since.month - 1]}-{since.year}"
        status, data = mail.uid("search", None, f"SINCE {since_imap}")
        if status != "OK" or not data or not data[0]:
            return replied

        uids = [int(u) for u in data[0].split()]
        print(f"Revisando {len(uids)} correo(s) de la cuenta desde {since:%Y-%m-%d} para detectar respuestas...")

        for i in range(0, len(uids), 200):
            uid_set = ",".join(str(u) for u in uids[i : i + 200])
            status, msg_data = mail.uid(
                "fetch",
                uid_set,
                "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT AUTO-SUBMITTED X-AUTOREPLY X-AUTORESPOND)])",
            )
            if status != "OK" or not msg_data:
                continue

            for part in msg_data:
                if not isinstance(part, tuple):
                    continue
                header_msg = email.message_from_bytes(part[1])
                from_addr = _normalize(parseaddr(header_msg.get("From", ""))[1])
                if not from_addr or from_addr in {sender_self, smtp_self}:
                    continue
                if any(k in from_addr for k in BOUNCE_SENDER_KEYWORDS):
                    continue
                subject = _decode_subject(header_msg)
                if _is_auto_reply(header_msg, subject):
                    continue
                replied.add(from_addr)

    return replied


def load_bounced_addresses(config: Config) -> set[str]:
    bounced: set[str] = set()
    path = Path(config.bounce_log_file)
    if path.exists():
        df = pd.read_excel(path, engine="openpyxl")
        bounced |= set(df["Correo"].dropna().map(_normalize))
    return bounced


def load_sent_clients(config: Config) -> pd.DataFrame:
    """Une los logs de envío y devuelve, por cada correo, el ÚLTIMO envío
    confirmado. También descarta a quien figure como rebotado en algún log."""
    frames = []
    for log_path in config.followup_source_logs:
        path = Path(log_path)
        if not path.exists():
            print(f"Aviso: no existe {log_path}, se omite.")
            continue
        df = pd.read_excel(path, sheet_name="Sheet1", engine="openpyxl")
        frames.append(df[["Fecha", "Cliente", "Correo", "Estado"]])
    if not frames:
        return pd.DataFrame(columns=["Fecha", "Cliente", "Correo"])

    df = pd.concat(frames, ignore_index=True)
    df["Correo"] = df["Correo"].astype(str).map(_normalize)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")

    bounced_in_logs = set(df.loc[df["Estado"] == "rebotado", "Correo"])
    sent = df[(df["Estado"] == "enviado") & ~df["Correo"].isin(bounced_in_logs)]
    sent = sent.dropna(subset=["Fecha"]).sort_values("Fecha")
    return sent.drop_duplicates(subset=["Correo"], keep="last")[["Fecha", "Cliente", "Correo"]]


def load_followup_done(config: Config) -> set[str]:
    path = Path(config.followup_log_file)
    if not path.exists():
        return set()
    df = pd.read_excel(path, sheet_name="Sheet1", engine="openpyxl")
    return set(df.loc[df["Estado"].isin(FOLLOWUP_DONE_STATES), "Correo"].dropna().map(_normalize))


def select_followup_candidates(
    sent: pd.DataFrame,
    followup_days: int,
    replied: set[str],
    bounced: set[str],
    already_done: set[str],
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Filtra los envíos y devuelve (candidatos, conteo_de_descartes)."""
    cutoff = datetime.now() - timedelta(days=followup_days)
    replied_domains = {_domain(a) for a in replied} - FREE_MAIL_DOMAINS - {""}

    def replied_mask(correo: str) -> bool:
        return correo in replied or _domain(correo) in replied_domains

    reasons = {
        "muy reciente": ~(sent["Fecha"] <= cutoff),
        "respondió (él o alguien de su empresa)": sent["Correo"].map(replied_mask),
        "rebotó": sent["Correo"].isin(bounced),
        "ya recibió el seguimiento": sent["Correo"].isin(already_done),
        "correo inválido": ~sent["Correo"].map(is_valid_email),
    }
    keep = pd.Series(True, index=sent.index)
    counts: dict[str, int] = {}
    for reason, mask in reasons.items():
        counts[reason] = int((keep & mask).sum())
        keep &= ~mask
    return sent[keep], counts


def image_is_online(url: str) -> bool:
    """Evita mandar miles de correos con la imagen rota."""
    if not url:
        return False
    try:
        with urllib.request.urlopen(url, timeout=15) as response:
            return response.status == 200 and response.headers.get_content_maintype() == "image"
    except Exception:
        return False


def whatsapp_url() -> str:
    return f"https://wa.me/{WHATSAPP_NUMBER}?text={quote(WHATSAPP_MESSAGE)}"


# Tres asuntos para probar cuál engancha más. Cada correo recibe siempre la
# misma variante (se elige por el propio correo), y queda anotada en el log
# de seguimiento para comparar respuestas por variante.
SUBJECT_VARIANTS = {
    "A": "Asunto de prueba A",
    "B": "Asunto de prueba B",
    "C": "Asunto de prueba C",
}


def pick_subject(address: str) -> tuple[str, str]:
    keys = sorted(SUBJECT_VARIANTS)
    variant = keys[zlib.crc32(address.encode("utf-8")) % len(keys)]
    return variant, SUBJECT_VARIANTS[variant]


def main():
    print("Correo de seguimiento — campaña cierre de año")
    auto_confirm = "--yes" in sys.argv[1:]
    config = load_config()
    print("send_emails:", config.send_emails)

    if config.send_emails:
        if not image_is_online(config.followup_image_url):
            print(
                f"La imagen del flyer no está disponible en {config.followup_image_url}. "
                "Súbela (assets/imagen-seguimiento.jpg) o corrige FOLLOWUP_IMAGE_URL. "
                "No se envía nada."
            )
            return
        if not is_business_day():
            print("Hoy es sábado o domingo. No se envía nada.")
            return
        if not is_within_send_window(config):
            print(
                f"Fuera del horario de envío permitido "
                f"({config.send_window_start:02d}:00–{config.send_window_end:02d}:00)."
            )
            return

    sent = load_sent_clients(config)
    print(f"Clientes con correo de presentación enviado: {len(sent)}")
    if sent.empty:
        return

    # Si no se puede revisar la bandeja, NO se manda nada: sería arriesgarse a
    # escribirle "¿ya tienen fecha?" a alguien que ya respondió.
    try:
        replied = load_replied_addresses(config, sent["Fecha"].min().to_pydatetime())
    except Exception as error:
        print("No se pudo revisar la bandeja por IMAP:", error)
        print("No se envía nada sin saber quién respondió.")
        return
    print(f"Remitentes distintos que escribieron a la cuenta: {len(replied)}")

    candidates, discarded = select_followup_candidates(
        sent,
        config.followup_days,
        replied,
        load_bounced_addresses(config),
        load_followup_done(config),
    )
    for reason, count in discarded.items():
        if count:
            print(f"  - Se omiten {count}: {reason}")
    print(f"Clientes para seguimiento: {len(candidates)}")
    if candidates.empty:
        print("No hay nadie pendiente de seguimiento hoy.")
        return

    wa_url = whatsapp_url()
    if not config.send_emails:
        first = candidates.iloc[0]
        preview = Path("data/followup_preview.html")
        # La vista previa usa la imagen local, para verla aunque aún no esté subida.
        local_image = "../assets/imagen-seguimiento.jpg"
        preview.write_text(
            render_followup_body(first["Cliente"], local_image, wa_url), encoding="utf-8"
        )
        online = "sí" if image_is_online(config.followup_image_url) else "NO (falta subirla)"
        print(f"Imagen del flyer en línea: {online}")
        variants = candidates["Correo"].map(lambda c: pick_subject(c)[0]).value_counts()
        print("Reparto de asuntos:", ", ".join(f"{k}={v}" for k, v in sorted(variants.items())))
        days = -(-len(candidates) // config.followup_daily_limit)
        print(f"Con lotes de {config.followup_daily_limit}/día: ~{days} día(s) hábiles.")
        print("DRY RUN: no se envía nada. Primeros 15 destinatarios:")
        for _, row in candidates.head(15).iterrows():
            print(f"  [{pick_subject(row['Correo'])[0]}] {row['Fecha']:%Y-%m-%d}  {row['Cliente']} <{row['Correo']}>")
        print(f"Vista previa del correo guardada en {preview}.")
        return

    # Dos topes: el de la cuenta (suma ambos envíos) y el propio del
    # seguimiento, más bajo, para cuidar la entregabilidad (envío por lotes).
    sent_today = count_sent_today(config.log_file) + count_sent_today(config.followup_log_file)
    followups_today = count_sent_today(config.followup_log_file)
    allowed = min(
        config.daily_limit - sent_today, config.followup_daily_limit - followups_today
    )
    print(
        f"Lote de hoy: hasta {config.followup_daily_limit} seguimientos "
        f"(ya enviados hoy: {followups_today}). Total de la cuenta hoy: {sent_today}/{config.daily_limit}."
    )
    if allowed <= 0:
        print("Ya se completó el lote de hoy. Continúa el próximo día hábil.")
        return

    if not auto_confirm and not confirm_send():
        print("Envío cancelado por el usuario.")
        return

    email_sender = EmailSender(config)
    logger = ExcelLogger(config.followup_log_file)

    for index, (_, row) in enumerate(candidates.iterrows()):
        remaining = len(candidates) - index
        if not is_within_send_window(config):
            print(f"Se llegó al fin del horario. Quedan {remaining}: continúa el próximo día hábil.")
            break
        if allowed <= 0:
            print(f"Lote de hoy completo. Quedan {remaining}: continúa el próximo día hábil.")
            break

        name, address = str(row["Cliente"]), row["Correo"]
        variant, subject = pick_subject(address)
        status = email_sender.send_email(
            address,
            subject,
            render_followup_body(name, config.followup_image_url, wa_url),
            render_followup_text(name, wa_url),
        )
        logger.log_send(name, address, status, extra={"Variante": variant, "Asunto": subject})
        print(f"[{variant}] {name} <{address}>: {status}")
        allowed -= 1
        time.sleep(2)  # pausa de 2 segundos para no saturar el servidor SMTP

    print(f"Listo. Resultado de cada envío en {config.followup_log_file}.")


if __name__ == "__main__":
    main()
