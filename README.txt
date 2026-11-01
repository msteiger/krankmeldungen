# Mailplanung – lokale Windows-Anwendung

## Inhalt
- `mailplanung.py`: grafische Anwendung
- `Mailplanung_Vorlage_inhalt.xlsx`: deine Excel-Datei
- `Start_Mailplanung.bat`: Startdatei für Windows
- `requirements.txt`: benötigtes Python-Paket

## Start
1. ZIP-Datei in einen eigenen Ordner entpacken.
2. Python 3 installieren (bei der Installation „Add Python to PATH“ aktivieren).
3. `Start_Mailplanung.bat` doppelklicken. Beim ersten Start wird `openpyxl` installiert.
4. Die Excel-Datei wird automatisch geladen. Falls nötig, über „Durchsuchen“ auswählen.

## Ablauf
1. Personen über Vor- oder Nachnamen suchen.
2. Eine oder mehrere Personen markieren und „Markierte hinzufügen“ klicken.
3. Die Anwendung gruppiert die Auswahl automatisch nach Jahrgang.
4. Für jeden Jahrgang wird genau eine E-Mail vorbereitet.
5. Betreff und Nachricht kommen aus dem Blatt `Inhalt`. Im Nachrichtentext wird `{PERSONEN}` durch die ausgewählten Namen ersetzt. `{JAHRGANG}` kann ebenfalls verwendet werden.
6. Der CC-Empfänger wird aus `Empfaenger` anhand von Wochentag und Jahrgang gelesen.
7. Vorschau prüfen und erst danach „E-Mails senden“ wählen.
8. Versandresultate werden im Blatt `Protokoll` eingetragen.

## SMTP-Einstellungen
Für GMX sind üblicherweise `mail.gmx.net` und Port `587` mit STARTTLS passend. Bei anderen Anbietern die SMTP-Daten des Anbieters eintragen.
- **An-Adresse:** Pflichtfeld. Da die Empfänger laut Vorgabe im CC stehen, wird diese Adresse in das Feld „An“ eingetragen.
- **Absenderadresse / SMTP-Benutzername:** meist die eigene E-Mail-Adresse.
- **SMTP-Passwort:** Bei Anbietern mit Zwei-Faktor-Anmeldung ein separates App-Passwort verwenden. Das Passwort wird nicht in der Konfigurationsdatei gespeichert.

## Sicherheit und Hinweise
- Es wird nichts automatisch versendet. Vor dem Versand erscheint eine Bestätigungsabfrage.
- Prüfe die Beispieladressen in `Empfaenger`: In der bereitgestellten Vorlage stehen offenbar Test-/Platzhalteradressen. Ersetze sie vor einem echten Versand durch die korrekten Adressen.
- Die Excel-Datei wird beim Protokollieren geändert. Lege regelmäßig eine Sicherungskopie an.
