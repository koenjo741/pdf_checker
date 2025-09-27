import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter.font as tkFont
import os
import platform
from datetime import datetime
from PyPDF2 import PdfReader

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
        style.configure("Treeview",
                        background="#292524",
                        foreground="white",
                        fieldbackground="#292524",
                        font=('Arial', 13),
                        rowheight=35)
        style.map('Treeview',
                  background=[('selected', '#555555')],
                  foreground=[('selected', 'white')])
        style.configure("Treeview.Heading",
                        background="#000000",
                        foreground="#FFC000",
                        font=('Arial', 14, 'bold'))

        columns = ("File", "CreationDate", "ModDate", "XMPCreateDate", "XMPModifyDate", "XMPMetadataDate", "FileSystemDate")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")

        self.tree.heading("File", text="Datei")
        self.tree.heading("CreationDate", text="/CreationDate")
        self.tree.heading("ModDate", text="/ModDate")
        self.tree.heading("XMPCreateDate", text="xmp:CreateDate")
        self.tree.heading("XMPModifyDate", text="xmp:ModifyDate")
        self.tree.heading("XMPMetadataDate", text="xmp:MetadataDate")
        self.tree.heading("FileSystemDate", text="Dateisystem-Datum")

        # Tags für Zeilenfarben und Schriftfarben
        self.tree.tag_configure('oddrow', background='#1D1A19') # 30% dunkler
        self.tree.tag_configure('evenrow', background='#2D2A29') # 30% dunkler
        self.tree.tag_configure('error', foreground='#FF0000')      # Rot
        self.tree.tag_configure('creation_date', foreground='#00FF7F') # Grün
        self.tree.tag_configure('fs_date', foreground='#00FFFF')      # Blau
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)
        self.row_count = 0

    def on_drop(self, event):
        files = self.tk.splitlist(event.data)
        for file_path in files:
            data_row, color_tag = self.process_file(file_path)

            # Tag für abwechselnde Zeilenfarbe hinzufügen
            row_style_tag = 'evenrow' if self.row_count % 2 == 0 else 'oddrow'

            self.tree.insert("", "end", values=data_row, tags=(color_tag, row_style_tag))
            self.row_count += 1

        self.adjust_column_widths()

    def adjust_column_widths(self):
        """Passt die Spaltenbreiten basierend auf dem Inhalt an."""
        font = tkFont.Font(family="Arial", size=13)  # Schriftart aus dem Style

        for col in self.tree['columns']:
            # Beginne mit der Breite des Spaltentitels
            max_width = font.measure(self.tree.heading(col, 'text'))

            # Überprüfe die Breite jeder Zelle in der Spalte
            for item in self.tree.get_children():
                try:
                    cell_value = self.tree.item(item, 'values')[self.tree['columns'].index(col)]
                    cell_width = font.measure(str(cell_value))
                    if cell_width > max_width:
                        max_width = cell_width
                except (IndexError, TypeError):
                    pass

            # Setze die Spaltenbreite mit etwas Puffer
            self.tree.column(col, width=max_width + 30)

    def process_file(self, file_path):
        file_name = os.path.basename(file_path)
        tag = ''
        creation_date, mod_date, xmp_create_date, xmp_modify_date, xmp_metadata_date, fs_date = "", "", "", "", "", ""

        if not file_path.lower().endswith('.pdf'):
            tag = 'error'
            row = (file_name, "Keine PDF-Datei", "", "", "", "", "")
            return row, tag

        try:
            with open(file_path, 'rb') as f:
                reader = PdfReader(f)
                info = reader.metadata

                if info:
                    creation_date = self.format_pdf_date(info.get('/CreationDate'))
                    mod_date = self.format_pdf_date(info.get('/ModDate'))

                # XMP-Verarbeitung isolieren, um Fehler abzufangen
                try:
                    xmp = reader.xmp_metadata
                    if xmp:
                        xmp_create_date = str(xmp.get('xmp:CreateDate', ''))
                        xmp_modify_date = str(xmp.get('xmp:ModifyDate', ''))
                        xmp_metadata_date = str(xmp.get('xmp:MetadataDate', ''))
                except Exception:
                    xmp_create_date = "XMP-Daten fehlerhaft"
                    xmp_modify_date = ""
                    xmp_metadata_date = ""

                if creation_date:
                    tag = 'creation_date'
                else:
                    fs_date = self.get_filesystem_creation_date(file_path)
                    if fs_date and fs_date != "N/A":
                        tag = 'fs_date'

        except PdfReadError:
            creation_date = "Fehler beim Lesen der PDF"
            tag = 'error'
        except Exception as e:
            creation_date = f"Unerwarteter Fehler: {e}"
            tag = 'error'

        row = (file_name, creation_date, mod_date, xmp_create_date, xmp_modify_date, xmp_metadata_date, fs_date)
        return row, tag

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

if __name__ == "__main__":
    app = App()
    app.mainloop()