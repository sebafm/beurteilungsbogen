import smtplib
from email.message import EmailMessage
from pathlib import Path

def create_email():
    """Erstellt ein neues E-Mail-Objekt."""
    return EmailMessage()

def set_recipient(email, recipient):
    """Setzt den Empfänger der E-Mail."""
    email["To"] = recipient

def set_subject(email, subject):
    """Setzt den Betreff der E-Mail."""
    email["Subject"] = subject

def set_body(email, body):
    """Fügt den Nachrichtentext zur E-Mail hinzu."""
    email.set_content(body)

def add_attachment(email, file_path):
    """Hängt eine Datei an die E-Mail an."""
    try:
        with open(file_path, "rb") as f:
            email.add_attachment(
                f.read(), maintype="application", subtype="octet-stream", filename=Path(file_path).name
            )
    except FileNotFoundError:
        print(f"Fehler: Datei {file_path} nicht gefunden.")
    except Exception as e:
        print(f"Fehler beim Anhängen der Datei {file_path}: {e}")

def send_email(email, smtp_server, smtp_port, sender_email, sender_password):
    """Sendet die E-Mail über den angegebenen SMTP-Server."""
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            email["From"] = sender_email
            server.send_message(email)
            print(f"E-Mail erfolgreich an {email['To']} gesendet.")
    except smtplib.SMTPException as e:
        print(f"Fehler beim Senden der E-Mail: {e}")