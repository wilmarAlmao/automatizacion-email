import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from src.config import Config


class EmailSender:
    def __init__(self, config: Config):
        self.config = config

    def send_email(
        self,
        recipient_email: str,
        subject: str,
        html_body: str,
        text_body: str,
    ) -> str:
        """Intenta enviar el correo. Devuelve uno de estos estados como texto:

        - "enviado": el servidor SMTP aceptó el mensaje para su entrega.
        - "rebotado": el servidor SMTP rechazó al destinatario de inmediato
          (rebote "duro" detectado en el acto, ej. dominio que no existe).
        - "fallido": otro error (conexión, autenticación, adjunto, etc.).

        Importante: la mayoría de rebotes NO se detectan aquí. Gmail casi
        siempre acepta el mensaje en el momento (por eso decimos "enviado"),
        y el rebote real llega minutos u horas después como un aviso a la
        bandeja de entrada de la cuenta que envía. Para esos casos usa
        `check_bounces.py` (ver `src/bounce_checker.py`).
        """
        message = MIMEMultipart("mixed")
        message["From"] = f"{self.config.sender_name} <{self.config.sender_email}>"
        message["To"] = recipient_email
        message["Subject"] = subject
        message["Reply-To"] = self.config.sender_email

        # Encabezado que activa el boton "Cancelar suscripcion" que Gmail/Outlook
        # muestran junto al remitente. Un correo sin esto, enviado a muchos
        # desconocidos, se parece mucho mas a spam.
        message["List-Unsubscribe"] = f"<mailto:{self.config.sender_email}?subject=BAJA>"
        message["List-Unsubscribe-Post"] = "List-Unsubscribe=One-Click"

        # Parte "alternative": el programa de correo del destinatario elige mostrar
        # la version en texto o la version en HTML.
        alternative = MIMEMultipart("alternative")
        alternative.attach(MIMEText(text_body, "plain", "utf-8"))
        alternative.attach(MIMEText(html_body, "html", "utf-8"))
        message.attach(alternative)

        try:
            with smtplib.SMTP(self.config.smtp_host, self.config.smtp_port) as server:
                server.starttls()
                server.login(self.config.smtp_user, self.config.smtp_password)
                server.sendmail(self.config.sender_email, recipient_email, message.as_string())
            return "enviado"
        except smtplib.SMTPRecipientsRefused as error:
            # El servidor rechazó al destinatario en el acto (rebote inmediato).
            print(f"Rebote inmediato para {recipient_email}:", error)
            return "rebotado"
        except Exception as error:
            print("Error al enviar correo:", error)
            return "fallido"
