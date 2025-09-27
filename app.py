import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDND
import os
import platform
from datetime import datetime
from PyPDF2 import PdfReader

from PyPDF2.errors import PdfReadError

class App(TkinterDND.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF-Metadaten-Extraktor")
        self.geometry("1200x600")

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

        self.tree.tag_configure('error', foreground='red')
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        files = self.tk.splitlist(event.data)
        for file_path in files:
            data_row, is_error = self.process_file(file_path)
            tags = ('error',) if is_error else ()
            self.tree.insert("", "end", values=data_row, tags=tags)

    def process_file(self, file_path):
        file_name = os.path.basename(file_path)

        creation_date, mod_date, xmp_create_date, xmp_modify_date, xmp_metadata_date, fs_date = "", "", "", "", "", ""
        is_error = False

        if not file_path.lower().endswith('.pdf'):
            is_error = True
            row = (file_name, "Keine PDF-Datei", "", "", "", "", "")
            return row, is_error

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

                if not creation_date:
                    fs_date = self.get_filesystem_creation_date(file_path)
        except PdfReadError:
            creation_date = "Fehler beim Lesen der PDF"
            is_error = True
        except Exception as e:
            creation_date = f"Unerwarteter Fehler: {e}"
            is_error = True

        row = (file_name, creation_date, mod_date, xmp_create_date, xmp_modify_date, xmp_metadata_date, fs_date)
        return row, is_error

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