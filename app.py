import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter.font as tkFont
import os
import platform
from datetime import datetime
from PyPDF2 import PdfReader
from pyhanko.pdf_utils.reader import PdfFileReader
from pyhanko.sign import validation

from PyPDF2.errors import PdfReadError

class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF-Metadaten-Extraktor")
        self.geometry("1200x600")
        self.configure(bg="#0C0A09")

        # Style für Dark Mode
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Toplevel", background="#0C0A09")
        style.configure("TFrame", background="#0C0A09")
        style.configure("TLabel", background="#0C0A09", foreground="white", font=('Arial', 13))
        style.configure("Header.TLabel", font=('Arial', 14, 'bold'), foreground="#FFC000")
        style.configure("Error.TLabel", foreground="#FF0000")

        # Haupt-Frame
        self.main_frame = ttk.Frame(self, style="TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Canvas und Scrollbar für scrollbaren Inhalt
        self.canvas = tk.Canvas(self.main_frame, bg="#0C0A09", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style="TFrame")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        # Nur die erste Datei aus der Liste nehmen
        file_path = self.tk.splitlist(event.data)[0]

        # Vorherige Inhalte löschen
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        metadata, is_error = self.process_file(file_path)

        row_index = 0
        for key, value in metadata.items():
            if value: # Nur anzeigen, wenn ein Wert vorhanden ist
                # Key Label (Überschrift)
                key_label = ttk.Label(self.scrollable_frame, text=key, style="Header.TLabel")
                key_label.grid(row=row_index, column=0, sticky="nw", padx=5, pady=2)

                # Value Label
                val_style = "Error.TLabel" if is_error and key == "Datei" else "TLabel"
                value_label = ttk.Label(self.scrollable_frame, text=value, wraplength=self.scrollable_frame.winfo_width() - 150, justify="left", style=val_style)
                value_label.grid(row=row_index, column=1, sticky="nw", padx=5, pady=2)
                row_index += 1

    def process_file(self, file_path):
        file_name = os.path.basename(file_path)
        is_error = False
        metadata = {
            "Datei": file_name,
            "/Title": "",
            "/Author": "",
            "/CreationDate": "",
            "/ModDate": "",
            "xmp:CreateDate": "",
            "xmp:ModifyDate": "",
            "xmp:MetadataDate": "",
            "Dateisystem-Datum": "",
            "Signaturdaten": ""
        }

        if not file_path.lower().endswith('.pdf'):
            is_error = True
            metadata["/Title"] = "Keine PDF-Datei"
            return metadata, is_error

        try:
            # PyPDF2 für allgemeine Metadaten
            with open(file_path, 'rb') as f:
                reader = PdfReader(f)
                info = reader.metadata

                if info:
                    metadata["/Title"] = info.get('/Title', '')
                    metadata["/Author"] = info.get('/Author', '')
                    metadata["/CreationDate"] = self.format_pdf_date(info.get('/CreationDate'))
                    metadata["/ModDate"] = self.format_pdf_date(info.get('/ModDate'))

                try:
                    xmp = reader.xmp_metadata
                    if xmp:
                        metadata["xmp:CreateDate"] = str(xmp.get('xmp:CreateDate', ''))
                        metadata["xmp:ModifyDate"] = str(xmp.get('xmp:ModifyDate', ''))
                        metadata["xmp:MetadataDate"] = str(xmp.get('xmp:MetadataDate', ''))
                except Exception:
                    metadata["xmp:CreateDate"] = "XMP-Daten fehlerhaft"

                if not metadata["/CreationDate"]:
                     fs_date = self.get_filesystem_creation_date(file_path)
                     if fs_date and fs_date != "N/A":
                        metadata["Dateisystem-Datum"] = fs_date

            # pyHanko für Signaturdaten
            metadata["Signaturdaten"] = self.get_signature_dates(file_path)

        except PdfReadError:
            metadata["/Title"] = "Fehler beim Lesen der PDF"
            is_error = True
        except Exception as e:
            metadata["/Title"] = f"Unerwarteter Fehler: {e}"
            is_error = True

        return metadata, is_error

    def format_pdf_date(self, pdf_date):
        if not pdf_date:
            return ""
        try:
            date_str = str(pdf_date).replace("D:", "").split('+')[0].split('-')[0].split('Z')[0]
            dt = datetime.strptime(date_str[:14], "%Y%m%d%H%M%S")
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except (ValueError, IndexError):
            return str(pdf_date)

    def get_filesystem_creation_date(self, file_path):
        try:
            timestamp = 0
            if platform.system() == "Windows":
                timestamp = os.path.getctime(file_path)
            else:
                try:
                    timestamp = os.stat(file_path).st_birthtime
                except AttributeError:
                    timestamp = os.path.getctime(file_path)
            return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return "N/A"

    def get_signature_dates(self, file_path):
        try:
            with open(file_path, 'rb') as f:
                r = PdfFileReader(f)
                if not r.embedded_signatures:
                    return ""

                dates = []
                for sig in r.embedded_signatures:
                    try:
                        # Direkten Leseversuch des Zeitstempels, um Validierungsfehler zu umgehen
                        signing_time = sig.get_signing_time()
                        # Zeitstempel ist timezone-aware (UTC), in lokale Zeit umwandeln
                        local_dt = signing_time.astimezone()
                        dates.append(local_dt.strftime("%Y-%m-%d %H:%M:%S"))
                    except Exception:
                        # Fängt Fehler ab, falls der Zeitstempel selbst fehlt oder defekt ist
                        dates.append("Zeitstempel nicht lesbar")

                return "\n".join(dates) if dates else "Keine Zeitstempel gefunden"
        except Exception:
            # Fängt Fehler ab, falls die Datei von pyHanko nicht gelesen werden kann
            return "Datei nicht lesbar/Signatur-Fehler"

if __name__ == "__main__":
    app = App()
    app.mainloop()