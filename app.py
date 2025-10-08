import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter.font as tkFont
import os
import platform
from datetime import datetime
import re
from PyPDF2 import PdfReader
from pyhanko.pdf_utils.reader import PdfFileReader
from PyPDF2.errors import PdfReadError


class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF-Metadaten-Extraktor")
        self.geometry("1350x600")
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
            "Signatur-Name": "",
            "Signatur-Datum": ""
        }

        if not file_path.lower().endswith('.pdf'):
            is_error = True
            metadata["/Title"] = "Keine PDF-Datei"
            return metadata, is_error

        try:
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

            # Signaturdaten (echte e-Signatur oder Fallback)
            sig_name, sig_date = self.get_signature_info(file_path)
            metadata["Signatur-Name"] = sig_name
            metadata["Signatur-Datum"] = sig_date

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

    def get_signature_info(self, file_path):
        """Liest Signaturname + Datum (echte e-Signatur oder Fallback)."""
        # 1) Versuch mit pyHanko (echte Signatur)
        try:
            with open(file_path, 'rb') as f:
                reader = PdfFileReader(f)
                signatures = reader.embedded_signatures

                if signatures:
                    for sig in signatures:
                        try:
                            signing_time = sig.get_signing_time()
                            signer = ""
                            if hasattr(sig, "signer_cert") and sig.signer_cert:
                                signer = sig.signer_cert.subject.human_friendly
                            if signing_time:
                                local_dt = signing_time.astimezone()
                                return signer, local_dt.strftime("%Y-%m-%d %H:%M:%S")
                        except Exception:
                            continue
        except Exception:
            pass

        # 2) Fallback: Suche Signaturfeld im AcroForm (echte digitale Signatur)
        try:
            reader = PdfReader(file_path)
            if "/AcroForm" in reader.trailer["/Root"]:
                acroform = reader.trailer["/Root"]["/AcroForm"]
                if "/Fields" in acroform:
                    for field in acroform["/Fields"]:
                        obj = field.get_object()
                        if obj.get("/FT") == "/Sig" and "/V" in obj:
                            sig_dict = obj["/V"]
                            name = sig_dict.get("/Name", "")
                            date = sig_dict.get("/M", "")
                            date = self.format_pdf_date(date)
                            return name, date
        except Exception:
            pass

        # 3) Fallback: Suche Signaturbox im PDF (sichtbare Unterschrift als Text)
        try:
            reader = PdfReader(file_path)
            for page in reader.pages:
                if "/Annots" in page:
                    for annot in page["/Annots"]:
                        obj = annot.get_object()
                        if "/Contents" in obj:
                            text = str(obj["/Contents"]).strip()

                            # Falls Name und Datum in getrennten Zeilen stehen
                            if "\n" in text:
                                parts = text.splitlines()
                                if len(parts) >= 2:
                                    name = parts[0].strip()
                                    date_str = parts[1].strip()

                                    # ISO-Format
                                    try:
                                        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
                                        return name, dt.strftime("%Y-%m-%d %H:%M:%S")
                                    except Exception:
                                        pass

                                    # deutsches Format
                                    try:
                                        dt = datetime.strptime(date_str, "%d.%m.%Y %H:%M")
                                        return name, dt.strftime("%Y-%m-%d %H:%M:%S")
                                    except Exception:
                                        try:
                                            dt = datetime.strptime(date_str, "%d.%m.%Y")
                                            return name, dt.strftime("%Y-%m-%d")
                                        except Exception:
                                            pass

                                    # Slash-Format
                                    try:
                                        dt = datetime.strptime(date_str, "%d/%m/%Y %H:%M")
                                        return name, dt.strftime("%Y-%m-%d %H:%M:%S")
                                    except Exception:
                                        try:
                                            dt = datetime.strptime(date_str, "%d/%m/%Y")
                                            return name, dt.strftime("%Y-%m-%d")
                                        except Exception:
                                            pass

                                    return name, date_str  # ungeparst zurückgeben

                            # Falls doch in einer Zeile: Regex-Muster
                            m = re.search(r"(?P<name>.+?),\s*(?P<date>.+)", text)
                            if m:
                                return m.group("name").strip(), m.group("date").strip()

                            return text, ""  # Fallback: nur roher Text

            return "", ""
        except Exception:
            return "", ""


if __name__ == "__main__":
    app = App()
    app.mainloop()