#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lokale Mailplanung: Personen auswählen, nach Jahrgang gruppieren und je Jahrgang eine Mail senden."""
import json
import smtplib
import ssl
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from email.message import EmailMessage
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    raise SystemExit("Bitte zuerst installieren: python -m pip install openpyxl")

APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "mail_config.json"
WEEKDAYS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]

DEFAULT_CONFIG = {
    "smtp_server": "mail.gmx.net",
    "smtp_port": 587,
    "sender": "",
    "username": "",
    "to_address": "",
    "use_starttls": True
}

class MailPlanApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mailplanung – Krankmeldungen")
        self.root.geometry("1120x760")
        self.root.minsize(900, 650)
        self.people = []
        self.selected = {}
        self.workbook_path = tk.StringVar(value=str(APP_DIR / "Mailplanung_Vorlage_inhalt.xlsx"))
        self.weekday = tk.StringVar(value=WEEKDAYS[datetime.now().weekday()] if datetime.now().weekday() < 5 else "Montag")
        self.search = tk.StringVar()
        self.smtp = self.load_config()
        self.status = tk.StringVar(value="Bitte Excel-Datei laden.")
        self.build_ui()
        self.load_workbook()

    def load_config(self):
        cfg = DEFAULT_CONFIG.copy()
        if CONFIG_FILE.exists():
            try:
                cfg.update(json.loads(CONFIG_FILE.read_text(encoding="utf-8")))
            except Exception:
                pass
        return cfg

    def save_config(self):
        cfg = {
            "smtp_server": self.smtp_server.get().strip(),
            "smtp_port": int(self.smtp_port.get().strip()),
            "sender": self.sender.get().strip(),
            "username": self.username.get().strip(),
            "to_address": self.to_address.get().strip(),
            "use_starttls": True
        }
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        self.smtp = cfg

    def build_ui(self):
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        filebar = ttk.Frame(outer)
        filebar.pack(fill="x")
        ttk.Label(filebar, text="Excel-Datei:").pack(side="left")
        ttk.Entry(filebar, textvariable=self.workbook_path).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(filebar, text="Durchsuchen…", command=self.choose_file).pack(side="left")
        ttk.Button(filebar, text="Neu laden", command=self.load_workbook).pack(side="left", padx=(6,0))

        settings = ttk.LabelFrame(outer, text="Versand")
        settings.pack(fill="x", pady=10)
        self.smtp_server = tk.StringVar(value=self.smtp.get("smtp_server", "mail.gmx.net"))
        self.smtp_port = tk.StringVar(value=str(self.smtp.get("smtp_port", 587)))
        self.sender = tk.StringVar(value=self.smtp.get("sender", ""))
        self.username = tk.StringVar(value=self.smtp.get("username", ""))
        self.to_address = tk.StringVar(value=self.smtp.get("to_address", ""))
        self.password = tk.StringVar()
        fields = [
            ("SMTP-Server", self.smtp_server, 22), ("Port", self.smtp_port, 7),
            ("Absenderadresse", self.sender, 30), ("SMTP-Benutzername", self.username, 24),
            ("An (Pflichtfeld)", self.to_address, 30), ("SMTP-Passwort", self.password, 20)
        ]
        for i, (label, var, width) in enumerate(fields):
            r, c = divmod(i, 3)
            cell = ttk.Frame(settings)
            cell.grid(row=r, column=c, sticky="ew", padx=6, pady=5)
            ttk.Label(cell, text=label).pack(anchor="w")
            ent = ttk.Entry(cell, textvariable=var, width=width, show="*" if label == "SMTP-Passwort" else "")
            ent.pack(fill="x")
        for c in range(3):
            settings.columnconfigure(c, weight=1)
        ttk.Label(settings, text="Das Passwort wird nicht gespeichert; es bleibt nur bis zum Schließen der Anwendung im Arbeitsspeicher.").grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(0,5))
        ttk.Button(settings, text="Einstellungen speichern (ohne Passwort)", command=self.save_settings).grid(row=2, column=2, sticky="e", padx=6, pady=(0,5))

        main = ttk.Frame(outer)
        main.pack(fill="both", expand=True)
        left = ttk.LabelFrame(main, text="Personen suchen und auswählen")
        left.pack(side="left", fill="both", expand=True, padx=(0,6))
        right = ttk.LabelFrame(main, text="Ausgewählte Personen")
        right.pack(side="left", fill="both", expand=True, padx=(6,0))

        top = ttk.Frame(left); top.pack(fill="x", padx=8, pady=8)
        ttk.Label(top, text="Suche (Vorname oder Nachname):").pack(anchor="w")
        search_entry = ttk.Entry(top, textvariable=self.search)
        search_entry.pack(fill="x")
        self.search.trace_add("write", lambda *_: self.refresh_people())
        self.people_list = tk.Listbox(left, selectmode="extended", height=18, exportselection=False)
        self.people_list.pack(fill="both", expand=True, padx=8, pady=(0,6))
        ttk.Button(left, text="Markierte hinzufügen →", command=self.add_people).pack(anchor="e", padx=8, pady=8)

        daybar = ttk.Frame(right); daybar.pack(fill="x", padx=8, pady=8)
        ttk.Label(daybar, text="Wochentag für CC-Empfänger:").pack(side="left")
        daybox = ttk.Combobox(daybar, textvariable=self.weekday, values=WEEKDAYS, state="readonly", width=14)
        daybox.pack(side="left", padx=6)
        daybox.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        self.selected_list = tk.Listbox(right, height=18, exportselection=False)
        self.selected_list.pack(fill="both", expand=True, padx=8, pady=(0,6))
        ttk.Button(right, text="Markierte entfernen", command=self.remove_people).pack(anchor="e", padx=8, pady=8)

        preview_frame = ttk.LabelFrame(outer, text="E-Mail-Vorschau (eine E-Mail pro Jahrgang)")
        preview_frame.pack(fill="both", expand=True, pady=10)
        self.preview = tk.Text(preview_frame, height=9, wrap="word")
        self.preview.pack(fill="both", expand=True, padx=8, pady=8)
        actions = ttk.Frame(outer); actions.pack(fill="x")
        ttk.Label(actions, textvariable=self.status).pack(side="left", fill="x", expand=True)
        ttk.Button(actions, text="Vorschau aktualisieren", command=self.update_preview).pack(side="right", padx=4)
        ttk.Button(actions, text="E-Mails senden…", command=self.send_emails).pack(side="right", padx=4)

    def choose_file(self):
        p = filedialog.askopenfilename(title="Excel-Datei auswählen", filetypes=[("Excel-Dateien", "*.xlsx")])
        if p:
            self.workbook_path.set(p)
            self.load_workbook()

    def load_workbook(self):
        p = Path(self.workbook_path.get())
        if not p.exists():
            self.status.set("Excel-Datei nicht gefunden. Bitte Datei auswählen.")
            return
        try:
            wb = load_workbook(p, data_only=True)
            for sheet in ("Personen", "Inhalt", "Empfaenger", "Protokoll"):
                if sheet not in wb.sheetnames:
                    raise ValueError(f"Das Tabellenblatt '{sheet}' fehlt.")
            ws = wb["Personen"]
            self.people = []
            for row_no, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                first, last, year = (list(row) + [None, None, None])[:3]
                if first and last and year:
                    self.people.append({"id": row_no, "first": str(first).strip(), "last": str(last).strip(), "year": str(year).strip()})
            self.selected.clear()
            self.refresh_people()
            self.refresh_selected()
            self.status.set(f"{len(self.people)} Personen geladen.")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Excel-Fehler", str(e))

    def refresh_people(self):
        q = self.search.get().strip().casefold()
        self.people_list.delete(0, tk.END)
        self.visible_people = []
        for p in self.people:
            label = f'{p["first"]} {p["last"]}  –  {p["year"]}'
            if q in label.casefold() and p["id"] not in self.selected:
                self.visible_people.append(p)
                self.people_list.insert(tk.END, label)

    def add_people(self):
        for idx in self.people_list.curselection():
            p = self.visible_people[idx]
            self.selected[p["id"]] = p
        self.refresh_people()
        self.refresh_selected()
        self.update_preview()

    def refresh_selected(self):
        self.selected_list.delete(0, tk.END)
        self.visible_selected = sorted(self.selected.values(), key=lambda p: (p["year"].casefold(), p["last"].casefold(), p["first"].casefold()))
        for p in self.visible_selected:
            self.selected_list.insert(tk.END, f'{p["year"]}: {p["first"]} {p["last"]}')

    def remove_people(self):
        for idx in reversed(self.selected_list.curselection()):
            p = self.visible_selected[idx]
            self.selected.pop(p["id"], None)
        self.refresh_people()
        self.refresh_selected()
        self.update_preview()

    def get_mail_data(self):
        wb = load_workbook(self.workbook_path.get(), data_only=True)
        content_ws = wb["Inhalt"]
        subject = body = None
        for row in content_ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:
                subject, body = str(row[0]), str(row[1])
                break
        if not subject or not body:
            raise ValueError("Im Tabellenblatt 'Inhalt' fehlt eine Zeile mit Betreff und Nachricht.")
        recipients = {}
        for row in wb["Empfaenger"].iter_rows(min_row=2, values_only=True):
            day, year, email = (list(row) + [None, None, None])[:3]
            if day and year and email:
                recipients[(str(day).strip(), str(year).strip())] = str(email).strip()
        return subject, body, recipients

    def build_messages(self):
        subject, template, recipients = self.get_mail_data()
        groups = {}
        for p in self.selected.values():
            groups.setdefault(p["year"], []).append(p)
        result = []
        for year in sorted(groups, key=str.casefold):
            names = ", ".join(f'{p["first"]} {p["last"]}' for p in sorted(groups[year], key=lambda p: (p["last"].casefold(), p["first"].casefold())))
            body = template.replace("{PERSONEN}", names).replace("{JAHRGANG}", year)
            cc = recipients.get((self.weekday.get(), year), "")
            result.append({"year": year, "subject": subject, "body": body, "cc": cc, "count": len(groups[year])})
        return result

    def update_preview(self):
        self.preview.delete("1.0", tk.END)
        if not self.selected:
            self.preview.insert(tk.END, "Noch keine Personen ausgewählt.\n")
            return
        try:
            messages = self.build_messages()
            for m in messages:
                self.preview.insert(tk.END, f'Jahrgang {m["year"]} – {m["count"]} Person(en)\n')
                self.preview.insert(tk.END, f'An: {self.to_address.get().strip() or "(bitte An-Adresse eintragen)"}\n')
                self.preview.insert(tk.END, f'CC: {m["cc"] or "(kein CC-Empfänger in Excel hinterlegt)"}\n')
                self.preview.insert(tk.END, f'Betreff: {m["subject"]}\n{m["body"]}\n')
                self.preview.insert(tk.END, "—" * 55 + "\n")
        except Exception as e:
            self.preview.insert(tk.END, f"Vorschau nicht möglich: {e}\n")

    def save_settings(self):
        try:
            self.save_config()
            self.status.set("Einstellungen gespeichert (Passwort nicht gespeichert).")
        except Exception as e:
            messagebox.showerror("Einstellungen", str(e))

    def log_result(self, entries):
        p = Path(self.workbook_path.get())
        wb = load_workbook(p)
        ws = wb["Protokoll"]
        for entry in entries:
            row = ws.max_row + 1
            ws.cell(row, 1, datetime.now().strftime("%d.%m.%Y"))
            ws.cell(row, 2, datetime.now().strftime("%H:%M:%S"))
            ws.cell(row, 3, self.weekday.get())
            ws.cell(row, 4, entry["year"])
            ws.cell(row, 5, entry["count"])
            ws.cell(row, 6, entry["cc"])
            ws.cell(row, 7, entry["status"])
            ws.cell(row, 8, entry["error"])
        wb.save(p)

    def send_emails(self):
        if not self.selected:
            messagebox.showwarning("Keine Auswahl", "Bitte zuerst Personen auswählen.")
            return
        try:
            messages = self.build_messages()
            to_addr = self.to_address.get().strip()
            sender = self.sender.get().strip()
            username = self.username.get().strip()
            password = self.password.get()
            server = self.smtp_server.get().strip()
            port = int(self.smtp_port.get().strip())
            if not all([to_addr, sender, username, password, server]):
                raise ValueError("Bitte An-Adresse, Absenderadresse, SMTP-Benutzername, SMTP-Passwort und SMTP-Server ausfüllen.")
            for m in messages:
                if not m["cc"]:
                    raise ValueError(f"Für Jahrgang {m['year']} ist am {self.weekday.get()} keine CC-Adresse im Blatt 'Empfaenger' hinterlegt.")
            summary = "\n".join(f'{m["year"]}: {m["count"]} Person(en), CC: {m["cc"]}' for m in messages)
            if not messagebox.askyesno("Versand bestätigen", f"Es werden {len(messages)} E-Mail(s) versendet:\n\n{summary}\n\nMöchtest du wirklich senden?"):
                return
            log = []
            context = ssl.create_default_context()
            with smtplib.SMTP(server, port, timeout=30) as smtp:
                smtp.ehlo()
                smtp.starttls(context=context)
                smtp.ehlo()
                smtp.login(username, password)
                for m in messages:
                    msg = EmailMessage()
                    msg["From"] = sender
                    msg["To"] = to_addr
                    msg["Cc"] = m["cc"]
                    msg["Subject"] = m["subject"]
                    msg.set_content(m["body"])
                    try:
                        smtp.send_message(msg, from_addr=sender, to_addrs=[to_addr, m["cc"]])
                        log.append({**m, "status": "Gesendet", "error": ""})
                    except Exception as e:
                        log.append({**m, "status": "Fehler", "error": str(e)})
                        raise
            self.log_result(log)
            self.status.set(f"{len(log)} E-Mail(s) erfolgreich versendet und protokolliert.")
            messagebox.showinfo("Versand abgeschlossen", f"{len(log)} E-Mail(s) wurden versendet.")
        except Exception as e:
            try:
                if 'log' in locals() and log:
                    self.log_result(log)
            except Exception:
                pass
            self.status.set("Versand nicht vollständig abgeschlossen.")
            messagebox.showerror("Versandfehler", str(e))

def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    MailPlanApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
