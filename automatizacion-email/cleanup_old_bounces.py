"""Script independiente, de uso ÚNICO (una sola vez): borra definitivamente
los avisos de rebote VIEJOS que ya estaban en tu bandeja de entrada desde
antes de que el programa supiera borrarlos automáticamente.

De aquí en adelante ya no hace falta correr esto: `check_bounces.py` borra
los avisos de rebote nuevos solo, en cada corrida.

Uso, desde la carpeta raíz del proyecto:
    python cleanup_old_bounces.py            # pide confirmación antes de borrar
    python cleanup_old_bounces.py --yes      # no pregunta
"""

import sys

from src.bounce_checker import cleanup_old_bounces, log_bounces, update_send_log
from src.config import load_config


def main():
    auto_confirm = "--yes" in sys.argv[1:]
    config = load_config()

    found = cleanup_old_bounces(config, auto_confirm=auto_confirm)

    if found:
        log_bounces(found, config.bounce_log_file)
        updated_rows = update_send_log(found, config.log_file)
        print(f"Confirmados en el log: {len(found)}. Filas marcadas 'rebotado': {updated_rows}.")


if __name__ == "__main__":
    main()
