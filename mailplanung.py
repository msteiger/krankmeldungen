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

from cryptography.fernet import Fernet, InvalidToken
from openpyxl import load_workbook

APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "mail_config.json"
KEY_FILE = APP_DIR / "smtp_key.key"
DEFAULT_BOOK = APP_DIR / "Mailplanung_Vorlage_inhalt.xlsx"
WEEKDAYS = ("Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag")
LOG_HEADERS = ("Datum", "Uhrzeit", "Wochentag", "Jahrgang", "Anzahl Personen", "Hauptempfänger (An)", "CC-Empfänger", "Versandstatus", "Fehlermeldung")


def local_weekday(now=None):
    weekday = (now or datetime.now()).weekday()
    return "Montag" if weekday >= 5 else WEEKDAYS[weekday]


def ensure_recipient_sheet(workbook):
    """Create the fixed To-address sheet and add newly seen years without touching addresses."""
    if "Jahrgangsempfaenger" not in workbook.sheetnames:
        ws = workbook.create_sheet("Jahrgangsempfaenger")
        ws.append(("Jahrgang", "E-Mail-Adresse"))
    ws = workbook["Jahrgangsempfaenger"]
    if ws.max_row == 0 or ws.cell(1, 1).value is None:
        ws.cell(1, 1, "Jahrgang")
        ws.cell(1, 2, "E-Mail-Adresse")
    existing = {str(ws.cell(r, 1).value).strip() for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value is not None}
    people = workbook["Personen"]
    for row in people.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True):
        year = str(row[0]).strip() if row[0] is not None else ""
        if year and year not in existing:
            ws.append((year, None))
            existing.add(year)
    return ws


def read_mail_data(path, years=()):
    path = Path(path)
    if not path.is_file():
        raise ValueError("Die ausgewählte Excel-Datei wurde nicht gefunden.")
    wb = load_workbook(path)
    for sheet in ("Personen", "Inhalt", "Empfaenger", "Protokoll"):
        if sheet not in wb.sheetnames:
            raise ValueError(f"Das Tabellenblatt '{sheet}' fehlt in der Excel-Datei.")
    to_ws = ensure_recipient_sheet(wb)
    if wb["Protokoll"].max_row == 0 or wb["Protokoll"].cell(1, 1).value is None:
        for col, header in enumerate(LOG_HEADERS, 1):
            wb["Protokoll"].cell(1, col, header)
    people = []
    for row_no, row in enumerate(wb["Personen"].iter_rows(min_row=2, min_col=1, max_col=3, values_only=True), 2):
        first, last, year = row
        if first is None and last is None and year is None:
            continue
        if not all(v is not None and str(v).strip() for v in (first, last, year)):
            raise ValueError(f"Ungültiger Personendatensatz in 'Personen', Zeile {row_no}: Vorname, Nachname und Jahrgang sind erforderlich.")
        people.append({"id": row_no, "first": str(first).strip(), "last": str(last).strip(), "year": str(year).strip()})
    content = next(((str(a).strip(), str(b)) for a, b in wb["Inhalt"].iter_rows(min_row=2, min_col=1, max_col=2, values_only=True) if a and b), None)
    if not content:
        raise ValueError("Im Tabellenblatt 'Inhalt' fehlt eine Zeile mit Betreff und Nachricht.")
    cc = {}
    for day, year, address in wb["Empfaenger"].iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
        if day and year and address:
            cc[(str(day).strip(), str(year).strip())] = str(address).strip()
    to = {str(to_ws.cell(r, 1).value).strip(): str(to_ws.cell(r, 2).value).strip()
          for r in range(2, to_ws.max_row + 1) if to_ws.cell(r, 1).value and to_ws.cell(r, 2).value}
    wb.save(path)
    return people, content[0], content[1], to, cc


def load_smtp_config():
    if not CONFIG_FILE.is_file():
        raise ValueError("SMTP-Konfiguration fehlt. Bitte zuerst SMTP_Einrichten.bat ausführen.")
    try:
        cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8-sig"))
        required = ("smtp_server", "smtp_port", "sender", "username", "password_encrypted")
        if any(not cfg.get(k) for k in required):
            raise ValueError
        if not KEY_FILE.is_file():
            raise ValueError
        password = Fernet(KEY_FILE.read_bytes()).decrypt(cfg["password_encrypted"].encode()).decode()
        return cfg, password
    except (OSError, ValueError, KeyError, InvalidToken, json.JSONDecodeError, TypeError):
        raise ValueError("SMTP-Konfiguration oder Schlüsseldatei fehlt oder ist ungültig. Bitte SMTP_Einrichten.bat erneut ausführen.") from None


def grouped_messages(selected, subject, template, to, cc, day):
    groups = {}
    for person in selected:
        groups.setdefault(person["year"], []).append(person)
    result = []
    for year in sorted(groups, key=str.casefold):
        people = sorted(groups[year], key=lambda p: (p["last"].casefold(), p["first"].casefold()))
        to_address = to.get(year, "").strip()
        cc_address = cc.get((day, year), "").strip()
        # Do not deliver the same mailbox twice through To and Cc.
        effective_cc = "" if to_address and cc_address and to_address.casefold() == cc_address.casefold() else cc_address
        result.append({"year": year, "count": len(people), "to": to_address,
                       "cc": effective_cc, "cc_configured": bool(cc_address), "subject": subject,
                       "body": template.replace("{PERSONEN}", ", ".join(p["first"] + " " + p["last"] for p in people)).replace("{JAHRGANG}", year)})
    return result


def validate_messages(messages):
    missing = []
    for m in messages:
        if not m["to"]:
            missing.append(f"Jahrgang {m['year']}: Hauptempfänger (An) fehlt.")
        if not m.get("cc_configured", bool(m["cc"])):
            missing.append(f"Jahrgang {m['year']}: CC-Adresse für den aktuellen Wochentag fehlt.")
    if missing:
        raise ValueError("Versand nicht möglich:\n" + "\n".join(missing))


class MailPlanApp:
    def __init__(self, root):
        self.root = root
        root.title("Mailplanung – Krankmeldungen")
        root.geometry("1100x760")
        root.minsize(850, 620)
        self.people, self.selected, self.visible_people = [], {}, []
        self.path = tk.StringVar(value=str(DEFAULT_BOOK))
        self.search = tk.StringVar()
        self.status = tk.StringVar(value="Excel-Datei wird geladen …")
        self.build_ui()
        self.load_workbook()

    def build_ui(self):
        outer = ttk.Frame(self.root, padding=12); outer.pack(fill="both", expand=True)
        bar = ttk.Frame(outer); bar.pack(fill="x")
        ttk.Label(bar, text="Excel-Datei:").pack(side="left")
        ttk.Entry(bar, textvariable=self.path, state="readonly").pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(bar, text="Datei auswählen", command=self.choose_file).pack(side="left")
        ttk.Button(bar, text="Neu laden", command=self.load_workbook).pack(side="left", padx=(6, 0))
        main = ttk.Frame(outer); main.pack(fill="both", expand=True, pady=10)
        left = ttk.LabelFrame(main, text="Personensuche"); left.pack(side="left", fill="both", expand=True, padx=(0, 6))
        right = ttk.LabelFrame(main, text="Ausgewählte Personen"); right.pack(side="left", fill="both", expand=True, padx=(6, 0))
        ttk.Label(left, text="Vorname oder Nachname:").pack(anchor="w", padx=8, pady=(8, 0))
        ttk.Entry(left, textvariable=self.search).pack(fill="x", padx=8)
        self.search.trace_add("write", lambda *_: self.refresh_people())
        self.people_list = ttk.Treeview(left, columns=("first", "last", "year"), show="headings", selectmode="extended")
        for col, title, width in (("first", "Vorname", 120), ("last", "Nachname", 150), ("year", "Jahrgang", 85)):
            self.people_list.heading(col, text=title); self.people_list.column(col, width=width, minwidth=60, stretch=True)
        self.people_list.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Button(left, text="Hinzufügen →", command=self.add_people).pack(anchor="e", padx=8, pady=(0, 8))
        self.selected_list = ttk.Treeview(right, columns=("first", "last", "year"), show="headings", selectmode="extended")
        for col, title, width in (("first", "Vorname", 120), ("last", "Nachname", 150), ("year", "Jahrgang", 85)):
            self.selected_list.heading(col, text=title); self.selected_list.column(col, width=width, minwidth=60, stretch=True)
        self.selected_list.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Button(right, text="Markierte entfernen", command=self.remove_people).pack(anchor="e", padx=8, pady=(0, 8))
        preview_frame = ttk.LabelFrame(outer, text="E-Mail-Vorschau (eine E-Mail pro Jahrgang)"); preview_frame.pack(fill="both", expand=True)
        self.preview = tk.Text(preview_frame, height=12, wrap="word"); self.preview.pack(fill="both", expand=True, padx=8, pady=8)
        actions = ttk.Frame(outer); actions.pack(fill="x", pady=(8, 0))
        ttk.Label(actions, textvariable=self.status).pack(side="left", fill="x", expand=True)
        ttk.Button(actions, text="Vorschau aktualisieren", command=self.update_preview).pack(side="right", padx=4)
        ttk.Button(actions, text="E-Mails senden …", command=self.send_emails).pack(side="right", padx=4)

    def choose_file(self):
        p = filedialog.askopenfilename(title="Excel-Datei auswählen", filetypes=[("Excel-Dateien", "*.xlsx")])
        if p:
            self.path.set(p); self.load_workbook()

    def load_workbook(self):
        try:
            self.people, self.subject, self.template, self.to, self.cc = read_mail_data(self.path.get())
            self.selected.clear(); self.refresh_people(); self.refresh_selected(); self.update_preview()
            self.status.set(f"{len(self.people)} Personen geladen. Wochentag: {local_weekday()}.")
        except Exception as e:
            self.status.set("Excel-Datei konnte nicht geladen werden.")
            messagebox.showerror("Excel-Fehler", str(e))

    def refresh_people(self):
        if not hasattr(self, "people_list"): return
        q = self.search.get().strip().casefold()
        for item in self.people_list.get_children(): self.people_list.delete(item)
        self.visible_people = []
        for p in self.people:
            if p["id"] not in self.selected and q in (p["first"] + " " + p["last"] + " " + p["year"]).casefold():
                self.visible_people.append(p); self.people_list.insert("", tk.END, values=(p["first"], p["last"], p["year"]))

    def refresh_selected(self):
        if not hasattr(self, "selected_list"): return
        for item in self.selected_list.get_children(): self.selected_list.delete(item)
        self.visible_selected = sorted(self.selected.values(), key=lambda p: (p["year"].casefold(), p["last"].casefold(), p["first"].casefold()))
        for p in self.visible_selected: self.selected_list.insert("", tk.END, values=(p["first"], p["last"], p["year"]))

    def add_people(self):
        rows = self.people_list.selection()
        for row in rows:
            p = self.visible_people[self.people_list.index(row)]; self.selected[p["id"]] = p
        self.refresh_people(); self.refresh_selected(); self.update_preview()

    def remove_people(self):
        for row in self.selected_list.selection():
            i = self.selected_list.index(row); self.selected.pop(self.visible_selected[i]["id"], None)
        self.refresh_people(); self.refresh_selected(); self.update_preview()

    def messages(self):
        # Determine weekday at preview/send time from the local system clock.
        return grouped_messages(self.selected.values(), self.subject, self.template, self.to, self.cc, local_weekday())

    def update_preview(self):
        if not hasattr(self, "preview"): return
        self.preview.delete("1.0", tk.END)
        if not self.selected:
            self.preview.insert(tk.END, "Noch keine Personen ausgewählt.\n"); return
        try:
            for m in self.messages():
                cc_line = f'CC: {m["cc"]}\n' if m["cc"] else ""
                self.preview.insert(tk.END, f'Jahrgang {m["year"]} – {m["count"]} Person(en)\nAn: {m["to"] or "(fehlt)"}\n' + cc_line + f'Betreff: {m["subject"]}\n{m["body"]}\n' + "─" * 55 + "\n")
        except Exception as e: self.preview.insert(tk.END, f"Vorschau nicht möglich: {e}\n")

    def log_result(self, entries):
        wb = load_workbook(self.path.get()); ws = wb["Protokoll"]
        if ws.cell(1, 1).value is None:
            for col, header in enumerate(LOG_HEADERS, 1): ws.cell(1, col, header)
        for entry in entries:
            now = datetime.now(); row = ws.max_row + 1
            values = (now.strftime("%d.%m.%Y"), now.strftime("%H:%M:%S"), entry["day"], entry["year"], entry["count"], entry["to"], entry["cc"], entry["status"], entry["error"])
            for col, value in enumerate(values, 1): ws.cell(row, col, value)
        wb.save(self.path.get())

    def send_emails(self):
        if not self.selected:
            messagebox.showwarning("Keine Auswahl", "Bitte zuerst Personen auswählen."); return
        log = []
        try:
            messages = self.messages(); validate_messages(messages)
            cfg, password = load_smtp_config()
            summary = "\n".join(f'{m["year"]}: An {m["to"]}; CC {m["cc"] or "(keiner; identisch mit An)"}' for m in messages)
            if not messagebox.askyesno("Versand bestätigen", f"Es werden {len(messages)} E-Mail(s) versendet:\n\n{summary}\n\nMöchtest du wirklich senden?"):
                return
            day = local_weekday(); context = ssl.create_default_context()
            attempted = 0
            try:
                with smtplib.SMTP(cfg["smtp_server"], int(cfg["smtp_port"]), timeout=30) as smtp:
                    smtp.ehlo()
                    if cfg.get("use_starttls", True): smtp.starttls(context=context); smtp.ehlo()
                    smtp.login(cfg["username"], password)
                    for m in messages:
                        attempted += 1
                        record = {**m, "day": day, "status": "Fehler", "error": "Diese E-Mail konnte nicht versendet werden."}
                        try:
                            msg = EmailMessage(); msg["From"] = cfg["sender"]; msg["To"] = m["to"]
                            if m["cc"]: msg["Cc"] = m["cc"]
                            msg["Subject"] = m["subject"]; msg.set_content(m["body"])
                            smtp.send_message(msg, from_addr=cfg["sender"], to_addrs=[m["to"]] + ([m["cc"]] if m["cc"] else []))
                            record.update(status="Gesendet", error="")
                        except Exception:
                            record["error"] = "Diese E-Mail konnte nicht versendet werden. Serverdetails werden aus Sicherheitsgründen nicht protokolliert."
                        log.append(record)
            except Exception:
                # Preserve per-message outcomes, including connection/login failures, without logging raw SMTP errors.
                for m in messages[attempted:]:
                    log.append({**m, "day": day, "status": "Fehler", "error": "Verbindung oder Anmeldung beim SMTP-Server fehlgeschlagen."})
            self.log_result(log)
            failed = sum(e["status"] != "Gesendet" for e in log)
            self.status.set(f"Versand abgeschlossen: {len(log)-failed} erfolgreich, {failed} fehlgeschlagen; protokolliert.")
            messagebox.showinfo("Versand abgeschlossen", self.status.get())
        except Exception as e:
            # Do not expose SMTP/library exception text: it may contain sensitive details.
            message = str(e) if isinstance(e, ValueError) else "SMTP-Verbindung oder Versand fehlgeschlagen. Zugangsdaten werden nicht angezeigt."
            self.status.set("Versand nicht vollständig abgeschlossen.")
            messagebox.showerror("Versandfehler", message)


def main():
    root = tk.Tk()
    try: ttk.Style().theme_use("vista")
    except tk.TclError: pass
    MailPlanApp(root); root.mainloop()


if __name__ == "__main__": main()
