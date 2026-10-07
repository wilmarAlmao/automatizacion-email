from datetime import datetime, time as dt_time, timedelta
from pathlib import Path

import pandas as pd

from src.appointments import Appointment, AppointmentScheduler
from src.config import Config, load_config
from src.email_sender import EmailSender
from src.excel_reader import read_clients
from src.logger import ExcelLogger
from src.template_renderer import render_email_body, render_email_text
from src.validators import is_valid_email
import sys
import time

HANDLED_STATES = {"enviado", "correo inválido", "rebotado", "fallido"}
# Estados que implican un intento real contra el servidor SMTP; son los que
# cuentan para el límite diario (un "correo inválido" nunca llega a enviarse).
SMTP_ATTEMPT_STATES = {"enviado", "rebotado", "fallido"}


def confirm_send() -> bool:
    answer = input("¿Deseas enviar los correos ahora? (s/n): ").strip().lower()
    return answer in {"s", "si", "y", "yes"}


def is_within_send_window(config: Config) -> bool:
    now = datetime.now().time()
    return dt_time(config.send_window_start, 0) <= now < dt_time(config.send_window_end, 0)


def is_business_day() -> bool:
    """Lunes=0 ... domingo=6. Las empresas (sobre todo RRHH) no trabajan
    sábado ni domingo, así que no tiene sentido mandar la campaña esos días."""
    return datetime.now().weekday() < 5


def load_already_handled_emails(log_file: str) -> set[str]:
    """Correos que ya tienen un resultado real registrado en corridas
    anteriores (enviado, rebotado, correo inválido o fallido). Se usa para
    no volver a mandarles nada si el programa se corta a mitad de camino y
    se vuelve a correr al día siguiente."""
    path = Path(log_file)
    if not path.exists():
        return set()

    df = pd.read_excel(path, sheet_name="Sheet1", engine="openpyxl")
    return set(
        df.loc[df["Estado"].isin(HANDLED_STATES), "Correo"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
    )


def count_sent_today(log_file: str) -> int:
    """Cuántos correos ya se intentaron enviar hoy según el log, sumando
    todas las corridas del día. Así el límite diario se respeta aunque el
    programa se ejecute varias veces el mismo día."""
    path = Path(log_file)
    if not path.exists():
        return 0

    df = pd.read_excel(path, sheet_name="Sheet1", engine="openpyxl")
    today = datetime.now().strftime("%Y-%m-%d")
    is_today = df["Fecha"].astype(str).str.startswith(today)
    return int((is_today & df["Estado"].isin(SMTP_ATTEMPT_STATES)).sum())


def create_follow_up_appointment(client: dict, followup_days: int) -> Appointment:
    """Fecha en la que toca el seguimiento: `followup_days` días después del
    envío. Ese día `python -m src.followup` ya le manda el segundo correo si
    el cliente no respondió."""
    scheduled_at = (datetime.now() + timedelta(days=followup_days)).replace(
        hour=10, minute=0, second=0, microsecond=0
    )
    return Appointment(
        client_name=client["name"],
        client_email=client["email"],
        scheduled_at=scheduled_at,
        subject=f"Seguimiento para {client['name']}",
        notes="Si no respondió: correo de seguimiento (python -m src.followup) y llamada.",
    )


def main():
    print("Iniciando el proyecto de envío de correos de ventas...")

    auto_confirm = "--yes" in sys.argv[1:]
    config = load_config()
    print("send_emails:", config.send_emails)
    print("Configuración cargada correctamente.")

    if config.send_emails and not config.brochure_url:
        print("Falta BROCHURE_URL en .env. El botón 'Conócenos' quedaría roto. No se envía nada.")
        return

    if config.send_emails and not is_business_day():
        print("Hoy es sábado o domingo. No se envía nada — se reintenta el próximo día hábil.")
        return

    if config.send_emails and not is_within_send_window(config):
        print(
            f"Fuera del horario de envío permitido "
            f"({config.send_window_start:02d}:00–{config.send_window_end:02d}:00). "
            f"No se enviará nada ahora."
        )
        return

    clients = read_clients(config.clients_file)
    print(f"Clientes leídos: {len(clients)}")

    already_handled = load_already_handled_emails(config.log_file)
    pending_clients = [
        c for c in clients if c["email"].strip().lower() not in already_handled
    ]
    skipped = len(clients) - len(pending_clients)
    if skipped:
        print(f"Se omiten {skipped} cliente(s) ya procesados en corridas anteriores.")

    if not pending_clients:
        print("No quedan clientes pendientes. Proceso completado.")
        return
    print(f"Clientes pendientes por procesar: {len(pending_clients)}")

    # El límite diario es de la cuenta de Gmail: suma también los correos de
    # seguimiento (src/followup.py) que se hayan mandado hoy.
    sent_today = (
        count_sent_today(config.log_file) + count_sent_today(config.followup_log_file)
        if config.send_emails
        else 0
    )
    if config.send_emails:
        print(f"Límite diario: {config.daily_limit}. Ya enviados hoy: {sent_today}.")
        if sent_today >= config.daily_limit:
            print("Ya se alcanzó el límite diario. No se envía nada — continúa mañana.")
            return

    email_sender = EmailSender(config)
    logger = ExcelLogger(config.log_file)
    scheduler = AppointmentScheduler()

    if config.send_emails:
        print("Modo envío activado. Se enviarán correos al finalizar la confirmación.")
        if not auto_confirm and not confirm_send():
            print("Envío cancelado por el usuario. No se enviaron correos.")
            return
        if auto_confirm:
            print("Confirmación automática (--yes): se procede sin preguntar.")
    else:
        print("Modo de prueba activado. No se enviarán correos reales.")

    cut_off_early = False
    for index, client in enumerate(pending_clients):
        if config.send_emails and not is_within_send_window(config):
            remaining = len(pending_clients) - index
            print(
                f"Se llegó al límite de horario ({config.send_window_end:02d}:00). "
                f"Quedan {remaining} cliente(s) pendientes: continuará automáticamente "
                f"mañana en el mismo horario."
            )
            cut_off_early = True
            break
        if config.send_emails and sent_today >= config.daily_limit:
            remaining = len(pending_clients) - index
            print(
                f"Se llegó al límite diario de {config.daily_limit} correos. "
                f"Quedan {remaining} cliente(s) pendientes: continuará mañana."
            )
            cut_off_early = True
            break
        body_html = render_email_body(client["name"], config.brochure_url)
        body_text = render_email_text(client["name"])
        subject = f"Novedades de Tu Empresa para {client['name']}"

        print(f"Preparando correo para: {client['name']} <{client['email']}>")

        if not is_valid_email(client["email"]):
            print(f"Correo inválido: {client['email']}. Se omite este cliente.")
            logger.log_send(client["name"], client["email"], "correo inválido")
            continue

        if config.send_emails:
            status = email_sender.send_email(
                client["email"],
                subject,
                body_html,
                body_text,
            )
            print(f"Resultado del envío: {status}")
            sent_today += 1
            time.sleep(2)  # pausa de 2 segundos para no saturar el servidor SMTP
        else:
            print("DRY RUN: correo preparado pero no enviado.")
            status = "pendiente"

        logger.log_send(client["name"], client["email"], status)

        if status == "enviado":
            appointment = create_follow_up_appointment(client, config.followup_days)
            scheduler.add_appointment(appointment)
            print(f"Cita de seguimiento registrada para {client['name']} el {appointment.scheduled_at}.")

    if not cut_off_early:
        print("Proceso completado. Verifica `data/send_log.xlsx` y `data/appointments.xlsx`.")
        print(
            "Tip: los rebotes reales suelen llegar horas después. Corre `python "
            "check_bounces.py` más tarde para detectarlos."
        )


if __name__ == "__main__":
    main()
