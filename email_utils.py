import smtplib
from email.message import EmailMessage
from pathlib import Path

def send_email(recipient, subject, body, sender_email, sender_password, smtp_server="smtp.example.com", smtp_port=587, attachments=None):
    """Sendet eine E-Mail mit optionalen Anhängen."""
    try:
        email = EmailMessage()
        email["To"] = recipient
        email["Subject"] = subject
        email["From"] = sender_email
        email.set_content(body)

        # Falls Anhänge vorhanden sind, hinzufügen
        if attachments:
            for file_path in attachments:
                try:
                    with open(file_path, "rb") as f:
                        email.add_attachment(
                            f.read(),
                            maintype="application",
                            subtype="octet-stream",
                            filename=Path(file_path).name,
                        )
                except FileNotFoundError:
                    print(f"Fehler: Datei {file_path} nicht gefunden.")

        # E-Mail senden
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(email)
            print(f"E-Mail erfolgreich an {recipient} gesendet.")
    except smtplib.SMTPException as e:
        print(f"Fehler beim Senden der E-Mail: {e}")