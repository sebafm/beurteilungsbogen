import os
import pandas as pd
import unittest
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from word_utils import load_word_template, replace_placeholders, save_word_document  # Funktionen aus word_utils.py importieren
from email_utils import create_email, set_recipient, set_subject, set_body, add_attachment, send_email  # Platzhalter für E-Mail-Bibliothek

# Funktion zum Laden der Excel-Datei mit Mitarbeiterdaten
def lade_excel_datei(dateipfad):
    try:
        return pd.read_excel(dateipfad).to_dict(orient="records")  # Konvertiert Excel in eine Liste von Dictionaries
    except FileNotFoundError:
        print(f"Fehler: Datei {dateipfad} nicht gefunden.")
        return None
    except Exception as e:
        print(f"Fehler beim Laden der Excel-Datei: {e}")
        return None

# Funktion zur Gruppierung der Mitarbeitenden nach Vorgesetzten
def gruppiere_nach_vorgesetzten(daten):
    vorgesetzten_gruppen = {}
    for eintrag in daten:
        vorgesetzter = eintrag["Vorgesetzter_Email"]
        mitarbeiter = {k: eintrag[k] for k in eintrag if k != "Vorgesetzter_Email"}  # Entfernt den Vorgesetzten-Schlüssel
        vorgesetzten_gruppen.setdefault(vorgesetzter, []).append(mitarbeiter)
    return vorgesetzten_gruppen

# Funktion zur Erstellung eines Word-Dokuments für einen Mitarbeiter
def erstelle_word_dokument(mitarbeiter, vorlage_pfad, speicher_pfad):
    try:
        dokument = load_word_template(vorlage_pfad)  # Lädt die Word-Vorlage
        replace_placeholders(dokument, mitarbeiter)  # Ersetzt Platzhalter mit Mitarbeiterdaten
        save_word_document(dokument, speicher_pfad)  # Speichert das ausgefüllte Dokument
    except Exception as e:
        print(f"Fehler beim Erstellen des Dokuments für {mitarbeiter['Name']}: {e}")

# Funktion zum Senden einer E-Mail mit angehängten Dokumenten
def sende_email_mit_anhang(vorgesetzter_email, dokumenten_liste):
    try:
        email = create_email()
        set_recipient(email, vorgesetzter_email)  # Setzt den Empfänger
        set_subject(email, "Leistungsbeurteilungen Ihrer Mitarbeitenden")  # Setzt den Betreff
        set_body(email, "Anbei finden Sie die Beurteilungsbögen Ihrer Mitarbeitenden.")  # Fügt den Nachrichtentext hinzu
        for dokument in dokumenten_liste:
            add_attachment(email, dokument)  # Hängt die Dokumente an
        send_email(email)  # Sendet die E-Mail
        print(f"E-Mail an {vorgesetzter_email} gesendet.")
    except Exception as e:
        print(f"Fehler beim Senden der E-Mail an {vorgesetzter_email}: {e}")

# Funktion zur Verarbeitung der Gruppe eines Vorgesetzten
def verarbeite_vorgesetzten_gruppe(vorgesetzter, mitarbeiter_liste, word_vorlage, speicher_ordner):
    dokumente = []
    for mitarbeiter in mitarbeiter_liste:
        dateiname = Path(speicher_ordner) / f"Beurteilungsbogen_{mitarbeiter['Name']}.docx"
        erstelle_word_dokument(mitarbeiter, word_vorlage, dateiname)  # Erstellt das Dokument
        dokumente.append(dateiname)
    sende_email_mit_anhang(vorgesetzter, dokumente)  # Sendet die E-Mail mit den Dokumenten

# Hauptfunktion des Skripts
def main():
    excel_datei = "Mitarbeiterdaten.xlsx"
    word_vorlage = "Beurteilungsbogen_Vorlage.docx"
    speicher_ordner = "Beurteilungsbögen/"
    os.makedirs(speicher_ordner, exist_ok=True)  # Erstellt den Speicherordner, falls er nicht existiert

    daten = lade_excel_datei(excel_datei)
    if daten is None:
        return  # Falls die Datei nicht geladen werden konnte, wird das Skript beendet

    vorgesetzten_gruppen = gruppiere_nach_vorgesetzten(daten)
    
    # Nutzt Multi-Threading für parallele Verarbeitung der Vorgesetzten-Gruppen
    with ThreadPoolExecutor() as executor:
        for vorgesetzter, mitarbeiter_liste in vorgesetzten_gruppen.items():
            executor.submit(verarbeite_vorgesetzten_gruppe, vorgesetzter, mitarbeiter_liste, word_vorlage, speicher_ordner)

    print("Prozess abgeschlossen.")

# Unit Tests für die Funktionen
class TestFunctions(unittest.TestCase):
  
    #Unit-Test für die Funktion gruppiere_nach_vorgesetzten:
    def test_gruppiere_nach_vorgesetzten(self):
        daten = [
            {"Vorgesetzter_Email": "chef@example.com", "Name": "Mitarbeiter1"},
            {"Vorgesetzter_Email": "chef@example.com", "Name": "Mitarbeiter2"},
            {"Vorgesetzter_Email": "boss@example.com", "Name": "Mitarbeiter3"}
        ]
        ergebnis = gruppiere_nach_vorgesetzten(daten)
        self.assertEqual(len(ergebnis), 2)  # Es sollten 2 Gruppen geben
        self.assertEqual(len(ergebnis["chef@example.com"]), 2)  # Eine Gruppe mit 2 Mitarbeitenden
        self.assertEqual(len(ergebnis["boss@example.com"]), 1)  # Eine Gruppe mit 1 Mitarbeitenden

if __name__ == "__main__":
    main()
    unittest.main()
