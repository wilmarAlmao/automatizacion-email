import os
from dataclasses import dataclass
from dotenv import load_dotenv


@dataclass
class Config:
    smtp_host: str
    smtp_port: int
    smtp_user: str
    smtp_password: str
    sender_name: str
    sender_email: str
    send_emails: bool
    clients_file: str
    log_file: str
    brochure_url: str
    imap_host: str
    bounce_log_file: str
    send_window_start: int
    send_window_end: int
    daily_limit: int
    followup_days: int
    followup_daily_limit: int
    followup_log_file: str
    followup_source_logs: list[str]
    followup_image_url: str


def _parse_bool(value: str, default: bool = False) -> bool:
    return str(value).strip().lower() in {"1", "true", "yes", "si", "y"}


def load_config() -> Config:
    load_dotenv()

    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    log_file = os.getenv("LOG_FILE", "data/send_log.xlsx")

    # Logs de donde salen los clientes para el correo de seguimiento. Por
    # defecto: el log viejo (data/send_log.xlsx) + el log actual.
    default_sources = ",".join(dict.fromkeys(["data/send_log.xlsx", log_file]))
    followup_source_logs = [
        p.strip()
        for p in os.getenv("FOLLOWUP_SOURCE_LOGS", default_sources).split(",")
        if p.strip()
    ]

    return Config(
        smtp_host=smtp_host,
        smtp_port=int(os.getenv("SMTP_PORT", 587)),
        smtp_user=os.getenv("SMTP_USER", ""),
        smtp_password=os.getenv("SMTP_PASSWORD", ""),
        sender_name=os.getenv("SENDER_NAME", "Ventas"),
        sender_email=os.getenv("SENDER_EMAIL", ""),
        send_emails=_parse_bool(os.getenv("SEND_EMAILS", "false")),
        clients_file=os.getenv("CLIENTS_FILE", "data/clients.xlsx"),
        log_file=log_file,
        brochure_url=os.getenv("BROCHURE_URL", ""),
        # Si no se define IMAP_HOST en .env, lo deducimos del SMTP_HOST
        # (smtp.gmail.com -> imap.gmail.com). Se usa para revisar rebotes.
        imap_host=os.getenv("IMAP_HOST", "") or smtp_host.replace("smtp.", "imap.", 1),
        bounce_log_file=os.getenv("BOUNCE_LOG_FILE", "data/bounces_log.xlsx"),
        # Ventana horaria en la que está permitido enviar correos reales.
        # Fuera de este rango, main.py no manda nada aunque se lo pidan.
        send_window_start=int(os.getenv("SEND_WINDOW_START", 9)),
        send_window_end=int(os.getenv("SEND_WINDOW_END", 17)),
        # Máximo de correos reales por día (cuenta también los de corridas
        # anteriores del mismo día). Lo que sobre se manda al día siguiente.
        daily_limit=int(os.getenv("DAILY_LIMIT", 700)),
        # Seguimiento: días que deben pasar desde el primer correo antes de
        # mandar el de seguimiento a quien no respondió.
        followup_days=int(os.getenv("FOLLOWUP_DAYS", 5)),
        # Lote diario del seguimiento (recomendación de la campaña: 40–60 por
        # día por remitente). Además respeta DAILY_LIMIT de la cuenta.
        followup_daily_limit=int(os.getenv("FOLLOWUP_DAILY_LIMIT", 60)),
        followup_log_file=os.getenv("FOLLOWUP_LOG_FILE", "data/followup_log.xlsx"),
        followup_source_logs=followup_source_logs,
        # URL pública (enlace directo) de la imagen del correo de seguimiento.
        followup_image_url=os.getenv("FOLLOWUP_IMAGE_URL", ""),
    )
