"""Script independiente: revisa la bandeja de entrada en busca de rebotes y
actualiza los registros. Correlo un rato después de una campaña (los rebotes
no llegan al instante) — por ejemplo, al día siguiente de correr `python -m
src.main` con SEND_EMAILS=true.

Uso, desde la carpeta raíz del proyecto:
    python check_bounces.py
"""

from src.bounce_checker import check_bounces, log_bounces, update_send_log
from src.config import load_config


def main():
    config = load_config()
    print("Revisando la bandeja de entrada en busca de rebotes...")

    bounced, replied_count = check_bounces(config)

    if replied_count:
        print(f"Se etiquetaron {replied_count} correo(s) entrante(s) como respuesta de cliente.")

    if not bounced:
        print("No se encontraron rebotes nuevos.")
        return

    print(f"Se detectaron {len(bounced)} rebote(s):")
    for item in bounced:
        print(f"  - {item['Correo']}  (\"{item['Asunto_rebote']}\")")

    log_bounces(bounced, config.bounce_log_file)
    updated_rows = update_send_log(bounced, config.log_file)

    print(f"Guardado en {config.bounce_log_file}.")
    print(f"Se marcaron {updated_rows} fila(s) como 'rebotado' en {config.log_file}.")


if __name__ == "__main__":
    main()
