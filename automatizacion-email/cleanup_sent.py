"""Script independiente: borra definitivamente, de tu carpeta de Enviados,
las copias de los correos que ya están confirmados como 'enviado' en
data/send_log.xlsx. Libera espacio del buzón (esas copias suelen llevar el
folleto PDF adjunto, que es lo que más pesa).

No afecta si un correo ya enviado luego rebota o no: eso se detecta aparte,
leyendo la bandeja de entrada (`check_bounces.py`).

Uso, desde la carpeta raíz del proyecto:
    python cleanup_sent.py            # pide confirmación antes de borrar
    python cleanup_sent.py --yes      # no pregunta (para correrlo programado)
"""

import sys

from src.config import load_config
from src.mailbox_cleanup import cleanup_sent_folder


def main():
    auto_confirm = "--yes" in sys.argv[1:]

    config = load_config()
    print("Revisando la carpeta de Enviados en busca de correos ya confirmados...")

    cleanup_sent_folder(config, auto_confirm=auto_confirm)


if __name__ == "__main__":
    main()
