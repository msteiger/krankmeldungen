Mailplanung – lokale Windows-Anwendung
========================================

Enthaltene Dateien
- mailplanung.py: Tkinter-Anwendung für Suche, Gruppierung, Vorschau und Versand
- smtp_setup.py: interaktive SMTP-Ersteinrichtung mit verdeckter Passworteingabe
- Mailplanung_Vorlage_inhalt.xlsx: bereitgestellte Excel-Datei
- requirements.txt: benötigte Python-Pakete (openpyxl, cryptography)
- Start_Mailplanung.bat und SMTP_Einrichten.bat: Windows-Startdateien
- test_mailplanung.py: lokale Funktionstests ohne echte SMTP-Verbindungen

Einrichtung
1. ZIP in einen eigenen Ordner entpacken.
2. Python 3 für Windows installieren und beim Setup „Add Python to PATH“ aktivieren.
3. SMTP_Einrichten.bat starten. Server, Port, Absender und Benutzername eingeben; das Passwort bleibt bei der Eingabe unsichtbar.
4. Start_Mailplanung.bat starten. Die Pakete werden installiert und die Anwendung geöffnet.

SMTP-Einrichtung erzeugt mail_config.json und smtp_key.key. Das Passwort steht nur Fernet-verschlüsselt in der JSON-Datei. Der Schlüssel liegt separat. Schützen und sichern Sie beide Dateien gemeinsam; wer beide Dateien erhält, kann das Passwort entschlüsseln. Die Dateien sind in .gitignore eingetragen und werden nicht in das bereitgestellte ZIP übernommen. Legen Sie sie niemals in öffentlichen oder ungeschützten Ablagen ab.

Arbeitsablauf
1. Die Personen werden aus dem Tabellenblatt „Personen“ ab Zeile 2 geladen. Suche findet Vor- und Nachnamen sowie Jahrgang; Mehrfachauswahl ist möglich.
2. Nach dem Hinzufügen werden Personen automatisch nach Jahrgang gruppiert.
3. Je Jahrgang wird genau eine E-Mail vorbereitet. Betreff und Nachricht kommen aus der ersten ausgefüllten Datenzeile in „Inhalt“. Unterstützte Nachrichtsplatzhalter: {PERSONEN} und {JAHRGANG}.
4. „An“ wird aus „Jahrgangsempfaenger“ nach Jahrgang gelesen. „CC“ kommt aus „Empfaenger“ passend zu lokalem aktuellem Wochentag und Jahrgang. Ist die CC-Adresse identisch mit der An-Adresse, wird sie aus CC entfernt und nur einmal als An-Empfänger verwendet. Der Wochentag ist nicht in der Oberfläche auswählbar und wird beim Anzeigen bzw. Senden neu ermittelt.
5. Tragen Sie die festen An-Adressen im neuen Blatt „Jahrgangsempfaenger“ ein. Das Blatt wird beim Laden ergänzt. Neue Jahrgänge werden hinzugefügt; vorhandene Adressen bleiben unangetastet.
6. Vor dem Versand müssen für alle Gruppen An- und CC-Adressen vorhanden sein. Die Bestätigungsabfrage zeigt Anzahl und Empfänger jeder E-Mail.
7. Versandresultate stehen im Blatt „Protokoll“ einschließlich Datum, Uhrzeit, Wochentag, Jahrgang, Anzahl, An, CC, Status und Fehlerhinweis. Zugangsdaten und rohe Serverfehler werden nicht protokolliert.

Wichtige Hinweise
- Die Excel-Datei wird beim Laden um „Jahrgangsempfaenger“ ergänzt und beim Protokollieren geändert. Erstellen Sie regelmäßige Sicherungskopien.
- Prüfen Sie vor dem echten Versand alle Empfängeradressen und den Nachrichtentext.
- Für SMTP wird STARTTLS verwendet. Nutzen Sie bei aktivierter Mehr-Faktor-Anmeldung ggf. ein App-Passwort.
- SMTP-Werte werden nicht in der Oberfläche angezeigt oder bearbeitet. Zur Aktualisierung SMTP_Einrichten.bat ausführen; vorhandene Konfiguration wird erst nach ausdrücklicher Bestätigung ersetzt.

Funktionstests
Mit installiertem Python und Abhängigkeiten im Programmordner ausführen:
  python -m unittest test_mailplanung.py
Die Tests prüfen Gruppierung, Platzhalter, Empfängerzuordnung, Wochentag und fehlende Adressen. Sie verwenden keinen SMTP-Server und senden keine E-Mails.
