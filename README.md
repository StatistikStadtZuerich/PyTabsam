# PyTabsam

## Zweck und grober Ablauf
PyTabsam ist ein Python-Skript zur Erstellung von Excel Tabellensammlungen, welche die bestehende Visual Basic Lösung von Statistik Stadt Zürich ablöst. 
PyTabsam erstellt aus einzelnen standardisierten Excel-Dateien in einem Verzeichnis eine zusammengefasste Excel-Datei mit mehreren Tabellenblättern.

In wenigen Worten kann der Prozessablauf wie folgt beschrieben werden.

### Einlesen der Steuerungsdatei
Die Steuerungsdatei config.json wird eingelesen und alle Verzeichnisse welche als "collection" angegeben wurden, werden verarbeitet. 

### Einlesen der Metadaten pro Tabellensammlung
Für jede Tabellensammlung werden die Metadaten aller Input-Dateien eingelesen. Falls Erläuterungen vorhanden sind, werden diese ebenfalls aus dem Tabellenblatt "Internet" eingelesen.

### Einlesen der Fussnoten
Für jede Tabellensammlung werden die Fussnoten aller Input-Dateien gelesen. Fehlerhafte Fussnoten werden mit einem Log-Eintrag ausgewiesen und müssen in den Quelldateien korrigiert werden. Mehr Details dazu finden sich im Kapitel "Warnungen".

### Kopieren der Vorlage und Befüllen der Inhalte aller Tabellenblätter pro Tabellensammlung
Die Vorlage "VorlageTabsam.xlsx" für eine Tabellensammlung wird für die Verarbeitung kopiert. Aufgrund der Informationen aus den eingelesenen Informationen werden alle Tabellenblätter inklusive Inhaltsverzeichnis und falls vorhanden der Erläuterungen generiert. Im gleichen Durchlauf pro collection-Eintrag werden die Daten aus dem Tabellenblatt "Internet" der Quelldateien gelesen und in die Tabellenblätter der Zieldatei eingefügt. Dabei werden auch die Formatierungen pro Tabellenblatt aus der Quelldatei übernommen und bei Bedarf an das Ziel angepasst (z.B. Setzen der Schriftfarbe schwarz). Die Fussnoten werden ebenfalls pro Tabellenblatt aufbereitet und mit hochgestellten Zahlen geschrieben.

## Bedienungsanleitung

### Steuerungsdatei config.json
Die Steuerungsdatei config.json enthält die Konfiguration der Input- und Output-Dateien. Die Steuerungsdatei muss jeweils vor der Verarbeitung für die zur erstellenden Tabellensammlungen konfiguriert werden.

#### title
Der "title" wird für die Metadaten der Excel-Datei verwendet. Dieser Titel ist ersichtlich in der Excel-Datei unter dem Menüpunkt "Datei-Informationen-Eigenschaften-Titel".

#### input_subpath
Der "input_subpath" setzt das Unterverzeichnis für das Quellverzeichnis der entsprechenden Tabellensammlung. Der Pfad wird ohne Schrägstriche angegeben. Er wird mit dem untenstehenden "path_input" und einem zusätzlichen Schrägstrich zusammengesetzt. In diesem Verzeichnis liegen die mit P-Transform erstellten Input-Dateien. Der "path_input" und "input_subpath" können auf die gewohnten Verzeichnisse z.B. "O:/Output/JB/Kap_Tabellensammlung/2021" und "07_03" gesetzt werden.

#### output_filename
Der "output_filename" setzt den Namen der Output-Datei.

#### path_input
Der "path_input" setzt den Basispfad auf das Quellverzeichnis für alle Tabellensammlungen. Der Pfad muss ohne abschliessenden Schrägstrich angegeben werden. 

#### path_output
Der "path_output" gibt das Ziel-Verzeichnis an, wohin die Output-Dateien geschrieben werden. Dabei ist zu beachten, dass bestehende Output-Dateien überschrieben werden. Es wird deshalb empfohlen eine temporäre Ablage zu verwenden.


### Warnungen
Es muss sichergestellt werden, dass keine Warnungen im Log vorhanden sind.

#### Fehlender Text zur Fussnote
Wenn in einer eingelesenen Tabelle im Tabellenblatt "Internet" ein Rautezeichen (#) gefolgt von einer Zahl gefunden wird (z.B. #1), dann wird dies als Fussnote interpretiert. PyTabsam prüft dann, ob im Tabellenblatt "Fussnoten" in der Spalte B ein Eintrag mit dem gleichen Code gefunden wird. Es werden nur Fussnoten geprüft, welche in der Spalte A den Tag "\<web\>" enthalten. Falls in der Spalte A nicht "\<web\>" steht bzw. der Fussnoten-Code in der Spalte B nicht gefunden wird, dann wird folgende Warnung ausgegeben:
```
WARNING (2022-03-08 14:29:20): The footnote #1 was referenced, but not found in the worksheet Fussnoten of T_02.02.03.xlsm
```
**Behebung:** Prüfen Sie in der Quelldatei die Übereinstimmung von Fussnoten in den Tabellenblättern "Fussnoten" und "Internet" und stellen sie sicher, dass "\<web\>" in der Spalte A des Tabellenblattes "Fussnoten" als Tag gesetzt ist.
       
#### Fehlende Nutzung einer Fussnotendefinition
Wenn eine Fussnote gemäss obiger Beschreibung korrekt im Tabellenblatt "Fussnoten" erfasst wurde, dieser Code aber weder in einem der Titel noch im Tabellenblatt "Internet" verwendet wird, dann wird folgende Warnung ausgegeben:
```
WARNING (2022-03-08 14:29:27): The footnote #4 is not used, but was found in the worksheet Fussnoten of T_02.02.05.2019.xlsm
```
**Behebung:** Prüfen Sie in der Quelldatei die Übereinstimmung von Fussnoten in den Tabellenblättern "Fussnoten" und "Internet". 

## Installation

### Pillow 
Pillow wird benötigt um Bilder zu verarbeiten. Ohne Pillow kann PyTabsam das SSZ-Logo aus der Vorlage nicht ins Ziel-Excel übernehmen.
Die Installation von Pillow erfolgt direkt im Python-Interpreter (z.B. im Miniconda Prompt):
```
conda install pillow
```
       
### Python Packages
Folgende Packages müssen zur Verfügung stehen:
       
| Package        | Zweck                         | 
| -------------- | ----------------------------- | 
| json           | Einlesen der JSON-Konfigurationsdatei config.json 
| pandas         | Wird wie eine Datenbank genutzt. Siehe Kapitel "Tabellen" in der "Technischen Dokumentation"
| os             | Betriebsystemfunktionen um Verzeichnisse mit Quell-Exceldateien einzulesen
| re             | Regular Expressions zum Filtern von Dateinamen und Inhalten
| openpyxl       | Excel lesen, verändern und schreiben
| shutil         | Betriebsystemfunktionen zur Kopie der Vorlagedatei
| datetime       | Zeitstempel für Output und Log
| copy           | Generische Kopierfunktion um Excel-Inhalte und -Formate zu übertragen
       
## Technische Dokumentation

### Funktionen
Die folgende Darstellung gibt eine Übersicht über die Funktionsweise und die im Code enthaltenen Funktionen:
       
```
                                     +---------------------------------------------------------+
                                     |                     PyTabsam main                       |
+-------------------------------+    |                                                         |    +---------------------------+
| Ordner A                      |    | INPUT LESEN                  OUTPUT SCHREIBEN           |    | Excel A                   |
|                               |    |                                                         |    |     Sheet "Erläuterungen" |
|   +------------------------+  |    | read_coll_dir                create_tabsam              | -> |     Sheet "T_1"           |
|   | Excel T_1              |  |    | read_all_md_fn                +- create_worksheet_expl  |    |     Sheet "T_2"           |
|   |     Sheet "Metadaten"  |  |    |  +- read_xls_metadata         +- create_worksheet       |    +---------------------------+
|   |     Sheet "Internet"   |  |    |  +- read_xls_footnote            +- read_write_data     |
|   |     Sheet "Fussnoten"  |  | -> |                                     +- convert_footnote |    +---------------------------+
|   +------------------------+  |    |                                  +- prepare_footnotes   |    | Excel B                   |
|                               |    |                                  +- write_footnotes     |    .                           .
|   +------------------------+  |    | HILFSFUNKTIONEN                                         | -> .                           .
|   | Excel T_2              |  |    |                                                         |
|   |     Sheet "Metadaten"  |  |    | tolog, read_config, to_superscript                      |
|   |     Sheet "Internet"   |  |    |                                                         |
|   |     Sheet "Fussnoten"  |  |    +---------------------------------------------------------+
|   +------------------------+  |
|                               |
|   +------------------------+  |
|   | Erläuterungen          |  |
|   |     Sheet "Internet"   |  |
|   +------------------------+  |
|                               |
+-------------------------------+ 

+-------------------------------+ 
| Ordner B                      |
.                               .
.                               .
```
       
#### main()
Hauptprogramm, ruft alle anderen Funktionen auf

#### tolog(level, text)
Gibt Log-Informationen aus. Level ist INFO, WARNING oder ERROR

#### read_config()
Liest die Konfigurationsdatei config.json ein

#### read_coll_dir()
Öffnet jedes in der Konfiguration angegebene Collection-Verzeichnis und prüft ob relevante Dateinamen für Sheets und Explanation vorkommen. 

#### read_all_md_fn()
Loop über alle Metadaten und Fussnoten. Ruft read_xls_metadata und read_xls_footnote auf.
       
#### read_xls_metadata(sheet_id)
Liest die Metadaten aus den Quell-Exceldateien ein

#### read_xls_footnote(sheet_id, list_foot)
Liest die Fussnoten aus den Quell-Exceldateien ein

#### convert_footnote(sheet_id, input_string)
Prüft ob der input_string eine Referenz zu einer Fussnote enthält. Nutzt to_superscript um Zahlen hochzustellen.
       
#### to_superscript(matchobj)
Konvertiert einen Fussnotencode (z.B. #1) zu seinem hochgestellten Äquivalent. 
Hinweis: Diese Funktion ist begrenzt auf die Zahlen von 1 bis 25. 

#### prepare_footnotes(sheet_id)
Überprüft, ob alle Fussnotendefinitionen genutzt wurden und erstellt eine Liste für den später generierten Output.

#### write_footnotes(dest_ws, list_wsfn, row_start)
Fügt verwendete Fussnotentexte ans Ende des aktuellen Excel-Tabellenblattes hinzu

#### create_tabsam()
Erstellt und generiert die Zieldatei mit allen Tabellenblättern.
Hinweis: Die Zieldatei wird überschrieben, falls sie bereits exisitiert.

#### create_worksheet_expl(coll_ID, dest_file)
Erstellt das Tabellenblatt "Erläuterungen" falls dieses in der Quelldatei vorhanden ist.
Liest die Daten von der Quelldatei im Tabellenblatt "Internet" und erstellt das Tabellenblatt in der Zieldatei.
Die Formatierungen aus dem Tabellenblatt "Internet" der Quelldatei werden in die Zieldatei übernommen.
Der Titel "Erläuterungen" wird im Inhaltsverzeichnis gesetzt.

#### create_worksheets(coll_ID, dest_file)
Erstellt alle Tabellenblätter aufgrund der vorhandenen Quelldateien.
Zuerst wird der Header im Tabellenblatt aufgrund der Angaben im DataFrame data_sheet geschrieben.
Liest die Daten von der Quelldatei im Tabellenblatt "Internet" und erstellt das Tabellenblatt in der Zieldatei.
Die Formatierungen aus dem Tabellenblatt "Internet" der Quelldatei werden in die Zieldatei übernommen.
Die Zeilenhöhe wird auf die einheitliche Höhe von 12.75 festgelegt.
Die Angaben im Inhaltsverzeichnis werden ergänzt.

#### read_write_data(source_ws, dest_ws, row_start, sheet_id):
Die Daten aus dem Tabellenblatt "Internet" in der Quelldatei werden zeilenweise gelesen und in das Tabellenblatt
der Zieldatei geschrieben.
Die Formatierungen aus dem Tabellenblatt "Internet" der Quelldatei werden in die Zieldatei übernommen.

       
### Tabellen 

```
+-------------+  1   n  +-------------+  1   n  +-------------+
| Collection  |---------| Sheet       |---------| Footnote    |
+-------------+         +-------------+         +-------------+
       |
       |  1          1  +-------------+
       +----------------| Explanation |
                        +-------------+
```


#### Collection
Pandas DataFrame data_coll

| Feldname        | Beispiel                      | Beschreibung |
| --------------- | ----------------------------- | ------------ |
| id              | 1                             | Schlüsselfeld 
| title           | Meteorologie                  | Titel der Tabellensammlung. Wird in den Metadaten-Titel der Datei gespeichert 
| input_path      | O:/Output/JAMA/Tabellen/02_02 | Pfad welcher die einzelnen Excel-Datein enthält 
| output_filename | Meteorologie.xlsx             | Name der zu erstellenden Datei 

#### Sheet
Pandas DataFrame data_sheet

| Feldname        | Beispiel                      | Beschreibung |
| --------------- | ----------------------------- | ------------ |
| ID              | 1                             | Primärschlüssel
| FK_collection   | 1                             | Fremdschlüssel zu Collection
| filename        | T_02.02.01.2017.xlsm          | Dateiname der Quelldatei 
| directory       | O:/Output/JAMA/Tabellen/02_02 | Verzeichnis der Quelldatei
| sheet_name      | T02.02.01.2017                | Generierter Name des Arbeitsblattes
| code            | T_2.2.1.2017                  | Tabellencode aus Blatt Metadaten
| title           | Wetterrekorde #1              | Tabellentitel aus Blatt Metadaten
| title_wfn       | Wetterrekorde                 | Tabellentitel ohne Fussnote (wfn=without footnote)
| subtitle1       | Station Zürich Fluntern #2    | Tabellenuntertitel1 aus Blatt Metadaten
| subtitle1_wfn   | Station Zürich Fluntern       | Tabellenuntertitel1 ohne Fussnote
| subtitle2       | historisch und 2017 #3        | Tabellenuntertitel2 aus Blatt Metadaten
| subtitle2_wfn   | historisch und 2017           | Tabellenuntertitel2 ohne Fussnote
| source          | MeteoSchweiz                  | Quelle aus Blatt Metadaten
| order           | 1                             | Sortierreihenfolge. Wird aktuell nicht verwendet (Reihenfolge der Quelldateinamen) 

#### Explanation
Pandas DataFrame data_expl

| Feldname        | Beispiel                      | Beschreibung |
| --------------- | ----------------------------- | ------------ |
| ID              | 1                             | Primärschlüssel
| FK_collection   | 1                             | Fremdschlüssel zu Collection
| filename        | T_02.02.01.2017.xlsm          | Dateiname der Quelldatei 
| directory       | O:/Output/JAMA/Tabellen/02_02 | Verzeichnis der Quelldatei

#### Footnote
Pandas DataFrame data_foot

| Feldname        | Beispiel                      | Beschreibung |
| --------------- | ----------------------------- | ------------ |
| ID              | 1                             | Primärschlüssel
| FK_sheet        | 1                             | Fremdschlüssel zu Sheet
| code            | #1                            | Nummerierung gemäss Input 
| text            | Neue Grenzwerte ab...         | Text der eigentlichen Fussnote
| used            | true                          | Wird die Fussnote verwendet (true/false)
