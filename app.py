import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
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
        self.configure(bg="#2E2E2E")

        # Style für Dark Mode
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Toplevel", background="#2E2E2E")
        style.configure("Treeview",
                        background="#3C3C3C",
                        foreground="white",
                        fieldbackground="#3C3C3C",
                        rowheight=25)
        style.map('Treeview', background=[('selected', '#555555')])
        style.configure("Treeview.Heading",
                        background="#555555",
                        foreground="white",
                        font=('Arial', 10, 'bold'))

        columns = ("File", "CreationDate", "ModDate", "XMPCreateDate", "XMPModifyDate", "XMPMetadataDate", "FileSystemDate")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")

        self.tree.heading("File", text="Datei")
        self.tree.heading("CreationDate", text="/CreationDate")
        self.tree.heading("ModDate", text="/ModDate")
        self.tree.heading("XMPCreateDate", text="xmp:CreateDate")
        self.tree.heading("XMPModifyDate", text="xmp:ModifyDate")
        self.tree.heading("XMPMetadataDate", text="xmp:MetadataDate")
        self.tree.heading("FileSystemDate", text="Dateisystem-Datum")

        self.tree.column("File", width=250)
        self.tree.column("CreationDate", width=150)
        self.tree.column("ModDate", width=150)
        self.tree.column("XMPCreateDate", width=150)
        self.tree.column("XMPModifyDate", width=150)
        self.tree.column("XMPMetadataDate", width=150)
        self.tree.column("FileSystemDate", width=150)

        self.tree.tag_configure('error', foreground='#FF6B6B') # Helleres Rot
        self.tree.tag_configure('creation_date', foreground='#81C784') # Grün
        self.tree.tag_configure('fs_date', foreground='#64B5F6') # Hellblau
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        files = self.tk.splitlist(event.data)
        for file_path in files:
            data_row, tag = self.process_file(file_path)
            self.tree.insert("", "end", values=data_row, tags=(tag,))

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

                xmp = reader.xmp_metadata
                if xmp:
                    xmp_create_date = str(xmp.get('xmp:CreateDate', ''))
                    xmp_modify_date = str(xmp.get('xmp:ModifyDate', ''))
                    xmp_metadata_date = str(xmp.get('xmp:MetadataDate', ''))

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