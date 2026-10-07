import pandas as pd
from pathlib import Path


def generate_sample_clients():
    path = Path("data")
    path.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(
        [
            {"Nombre": "Ana Pérez", "Correo": "ana@example.com"},
            {"Nombre": "Luis García", "Correo": "luis@example.com"},
            {"Nombre": "María López", "Correo": "maria@example.com"},
        ]
    )
    df.to_excel(path / "datapractics.xlsx", index=False, engine="openpyxl")
    print(f"Archivo de clientes creado en {path / 'data/datapractics.xlsx'}")


if __name__ == "__main__":
    generate_sample_clients()
