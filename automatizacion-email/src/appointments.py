from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import pandas as pd


@dataclass
class Appointment:
    client_name: str
    client_email: str
    scheduled_at: datetime
    subject: str
    notes: str = ""


class AppointmentScheduler:
    def __init__(self, file_path: str = "data/appointments.xlsx"):
        self.file_path = Path(file_path)
        self.file_path.parent.mkdir(parents=True, exist_ok=True)

        if not self.file_path.exists():
            df = pd.DataFrame(
                columns=["Cliente", "Correo", "FechaHora", "Asunto", "Notas"]
            )
            df.to_excel(self.file_path, index=False, engine="openpyxl")

    def add_appointment(self, appointment: Appointment) -> None:
        df = pd.read_excel(self.file_path, engine="openpyxl")
        new_row = {
            "Cliente": appointment.client_name,
            "Correo": appointment.client_email,
            "FechaHora": appointment.scheduled_at.strftime("%Y-%m-%d %H:%M:%S"),
            "Asunto": appointment.subject,
            "Notas": appointment.notes,
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(self.file_path, index=False, engine="openpyxl")

    def list_appointments(self) -> list[Appointment]:
        df = pd.read_excel(self.file_path, engine="openpyxl")
        appointments = []
        for _, row in df.iterrows():
            appointments.append(
                Appointment(
                    client_name=str(row["Cliente"]),
                    client_email=str(row["Correo"]),
                    scheduled_at=datetime.strptime(
                        str(row["FechaHora"]), "%Y-%m-%d %H:%M:%S"
                    ),
                    subject=str(row["Asunto"]),
                    notes=str(row["Notas"]),
                )
            )
        return appointments
