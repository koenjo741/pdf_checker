# PDF-Metadaten-Extraktor

Dieses Programm extrahiert Metadaten aus PDF-Dateien. Es bietet eine grafische Benutzeroberfläche, auf die Dateien per Drag & Drop gezogen werden können.

## Funktionalität

- Extrahiert `/CreationDate`, `/ModDate`, `xmp:CreateDate`, `xmp:ModifyDate` und `xmp:MetadataDate`.
- Falls `/CreationDate` fehlt, wird der Erstellungszeitstempel des Dateisystems verwendet.
- Nicht-PDF-Dateien werden erkannt und in der Liste rot markiert.

## Ausführung

1.  Installieren Sie die Abhängigkeiten:
    ```bash
    pip install -r requirements.txt
    ```
2.  Führen Sie das Skript aus:
    ```bash
    python app.py
    ```

## Erstellen einer `.exe`-Datei (für Windows)

Um eine eigenständige `.exe`-Datei zu erstellen, die ohne installierte Python-Umgebung ausgeführt werden kann, können Sie `PyInstaller` verwenden.

1.  Installieren Sie `PyInstaller`:
    ```bash
    pip install pyinstaller
    ```
2.  Führen Sie den folgenden Befehl im Terminal aus, um die Anwendung zu bündeln:
    ```bash
    pyinstaller --onefile --windowed --add-data "C:\Users\josef\AppData\Local\Programs\Python\Python313\Lib\site-packages\tkinterdnd2;tkinterdnd2" app.py
    ```
    *Hinweis: Der obige Pfad zu `tkinterdnd2` ist spezifisch für Ihr System. Wenn Sie Python aktualisieren oder an einem anderen Ort installieren, muss er möglicherweise angepasst werden.*

3.  Die fertige `.exe`-Datei befindet sich anschließend im `dist`-Verzeichnis.