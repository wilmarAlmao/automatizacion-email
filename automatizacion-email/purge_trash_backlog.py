"""Script independiente, de uso ÚNICO (una sola vez): arregla el atasco que
dejó un bug ya corregido en `cleanup_sent.py` / `cleanup_old_bounces.py` —
esos correos se movían a la Papelera pero nunca se purgaban de verdad, así
que se quedaron ahí ocupando espacio.

Revisa lo que hay en la Papelera y purga SOLO lo que reconoce como propio de
esta herramienta (copias de correos enviados de la campaña, o avisos de
rebote). No toca nada más que tengas en la Papelera por tu cuenta.

Después de correr esto una vez, no vuelve a hacer falta: los demás scripts
ya purgan bien de ahora en adelante.

Uso, desde la carpeta raíz del proyecto:
    python purge_trash_backlog.py            # pide confirmación antes de purgar
    python purge_trash_backlog.py --yes      # no pregunta
"""

import sys

from src.bounce_checker import purge_trash_backlog
from src.config import load_config


def main():
    auto_confirm = "--yes" in sys.argv[1:]
    config = load_config()

    purge_trash_backlog(config, auto_confirm=auto_confirm)


if __name__ == "__main__":
    main()
