"""Pruebas rápidas y seguras para src/bounce_checker.py.

No se conecta a Gmail ni toca tus archivos reales de data/ — arma mensajes de
rebote de ejemplo (parecidos a los que ya viste en la terminal) y confirma
que el programa los clasifica bien. Todo lo que crea (Excel de prueba) vive
en una carpeta temporal que se borra sola al terminar.

Cómo correrlo, desde la carpeta raíz del proyecto:

    python test_bounce_checker.py

Si todo sale bien, vas a ver "TODO OK" al final. Si algo falla, te va a
decir exactamente cuál prueba y por qué — eso significa que hay que revisar
el código antes de confiar en él con datos reales.
"""

import email
import sys
import tempfile
from email.message import Message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd

from src.bounce_checker import (
    _decode_subject,
    _extract_failed_recipient_fallback,
    _is_real_recipient,
    _looks_like_bounce,
    _parse_delivery_status,
    log_bounces,
    update_send_log,
)

fallos = []


def check(nombre, condicion):
    estado = "OK" if condicion else "FALLÓ"
    print(f"[{estado}] {nombre}")
    if not condicion:
        fallos.append(nombre)


def construir_rebote_con_dsn(destinatario: str, accion: str, asunto: str):
    """Arma un correo de rebote 'de verdad', con la parte estructurada que
    usan Gmail y la mayoría de proveedores (RFC 3464)."""
    outer = MIMEMultipart("report", report_type="delivery-status")
    outer["From"] = "Mail Delivery Subsystem <mailer-daemon@googlemail.com>"
    outer["Subject"] = asunto
    outer.attach(MIMEText("Aviso de entrega.", "plain"))

    dsn = Message()
    dsn["Content-Type"] = "message/delivery-status"
    bloque_mensaje = email.message_from_string("Reporting-MTA: dns; googlemail.com\n")
    bloque_destinatario = email.message_from_string(
        f"Final-Recipient: rfc822; {destinatario}\n"
        f"Action: {accion}\n"
        "Status: 5.1.1\n"
    )
    dsn.set_payload([bloque_mensaje, bloque_destinatario])
    outer.attach(dsn)

    return email.message_from_bytes(outer.as_bytes())


def construir_rebote_sin_dsn(asunto: str, cuerpo: str):
    """Simula un proveedor que NO manda la parte estructurada — solo texto
    libre (el caso menos común, donde usamos el heurístico de respaldo)."""
    outer = MIMEMultipart()
    outer["From"] = "Mail Delivery Subsystem <mailer-daemon@example.com>"
    outer["Subject"] = asunto
    outer.attach(MIMEText(cuerpo, "plain"))
    return email.message_from_bytes(outer.as_bytes())


print("=== Prueba 1: rebote real confirmado (Action: failed) ===")
msg1 = construir_rebote_con_dsn(
    "directorio@empresa-uno.com", "failed", "Delivery Status Notification (Failure)"
)
subject1 = _decode_subject(msg1)
entradas1 = _parse_delivery_status(msg1)
check("se detecta como aviso de rebote", _looks_like_bounce(msg1, subject1))
check("se encuentra 1 entrada DSN", len(entradas1) == 1)
check("la acción es 'failed'", bool(entradas1) and entradas1[0]["action"] == "failed")
check(
    "el destinatario es el correcto",
    bool(entradas1) and entradas1[0]["recipient"] == "directorio@empresa-uno.com",
)

print()
print("=== Prueba 2: aviso de RETRASO, no debe contarse como rebote ===")
msg2 = construir_rebote_con_dsn(
    "contacto@empresa-dos.com", "delayed", "Delivery Status Notification (Delay)"
)
entradas2 = _parse_delivery_status(msg2)
check("la acción detectada es 'delayed'", bool(entradas2) and entradas2[0]["action"] == "delayed")
check(
    "check_bounces.py debe ignorarlo (la acción no es 'failed')",
    bool(entradas2) and entradas2[0]["action"] != "failed",
)

print()
print("=== Prueba 3: dirección interna de Google, no es un cliente real ===")
check(
    "una dirección @mx.google.com se reconoce como NO real",
    not _is_real_recipient("6a95a78f.797575bc.7ba30.98dc@mx.google.com"),
)
check(
    "una dirección normal de cliente sí se reconoce como real",
    _is_real_recipient("ventas@empresa-tres.com"),
)

print()
print("=== Prueba 4: rebote sin parte técnica (heurístico de respaldo) ===")
msg4 = construir_rebote_sin_dsn(
    "Undelivered Mail Returned to Sender",
    "El mensaje para sac@empresa-cuatro.com no pudo entregarse.",
)
subject4 = _decode_subject(msg4)
recuperado = _extract_failed_recipient_fallback(
    msg4, known_emails={"sac@empresa-cuatro.com"}, sender_email="ventas@tuempresa.com"
)
check("se detecta como aviso de rebote", _looks_like_bounce(msg4, subject4))
check("se identifica el correo correcto por el heurístico de respaldo", recuperado == "sac@empresa-cuatro.com")

print()
print("=== Prueba 5: guardar en bounces_log.xlsx sin duplicar ===")
with tempfile.TemporaryDirectory() as tmp:
    log_path = Path(tmp) / "bounces_log.xlsx"
    rebotes = [
        {"Correo": "ventas@empresa-tres.com.", "Fecha_deteccion": "2026-01-01 10:00:00", "Asunto_rebote": "x"},
        {"Correo": "VENTAS@empresa-tres.com", "Fecha_deteccion": "2026-01-01 10:00:05", "Asunto_rebote": "x"},
    ]
    log_bounces(rebotes, str(log_path))
    resultado = pd.read_excel(log_path, engine="openpyxl")
    check("no se duplica la misma dirección (con mayúsculas/punto distintos)", len(resultado) == 1)
    check(
        "la dirección quedó normalizada en minúsculas y sin punto final",
        resultado["Correo"].iloc[0] == "ventas@empresa-tres.com",
    )

print()
print("=== Prueba 6: marcar como 'rebotado' en send_log.xlsx ===")
with tempfile.TemporaryDirectory() as tmp:
    send_log_path = Path(tmp) / "send_log.xlsx"
    pd.DataFrame(
        [
            {"Fecha": "2026-01-01", "Cliente": "Empresa Tres", "Correo": "ventas@empresa-tres.com", "Estado": "enviado"},
            {"Fecha": "2026-01-01", "Cliente": "Otro", "Correo": "otro@ejemplo.com", "Estado": "enviado"},
        ]
    ).to_excel(send_log_path, index=False, engine="openpyxl")

    actualizadas = update_send_log(
        [{"Correo": "ventas@empresa-tres.com.", "Fecha_deteccion": "x", "Asunto_rebote": "x"}],
        str(send_log_path),
    )
    resultado = pd.read_excel(send_log_path, engine="openpyxl")
    estado_tres = resultado.loc[resultado["Correo"] == "ventas@empresa-tres.com", "Estado"].iloc[0]
    estado_otro = resultado.loc[resultado["Correo"] == "otro@ejemplo.com", "Estado"].iloc[0]

    check("se actualizó exactamente 1 fila", actualizadas == 1)
    check("la fila correcta quedó como 'rebotado'", estado_tres == "rebotado")
    check("la otra fila NO se tocó", estado_otro == "enviado")

print()
if fallos:
    print(f"❌ {len(fallos)} prueba(s) fallaron: {', '.join(fallos)}")
    sys.exit(1)
else:
    print("✅ TODO OK — las 6 pruebas pasaron.")
