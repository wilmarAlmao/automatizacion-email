import pandas as pd
from datetime import datetime
from pathlib import Path


class ExcelLogger:
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.file_path.parent.mkdir(parents=True, exist_ok=True)

        if not self.file_path.exists():
            df = pd.DataFrame(columns=["Fecha", "Cliente", "Correo", "Estado"])
            df.to_excel(self.file_path, sheet_name="Sheet1", index=False, engine="openpyxl")

    def log_send(
        self, client_name: str, client_email: str, status: str, extra: dict | None = None
    ) -> None:
        """`extra`: columnas adicionales opcionales (ej. la variante de asunto
        del correo de seguimiento)."""
        df = pd.read_excel(self.file_path, sheet_name="Sheet1", engine="openpyxl")
        new_row = {
            "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Cliente": client_name,
            "Correo": client_email,
            "Estado": status,
            **(extra or {}),
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        with pd.ExcelWriter(
            self.file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)
