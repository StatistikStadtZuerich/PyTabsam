# PyTabsam
Python Excel Tabellensammlung

## Bedienungsanleitung



### Warnungen
Es muss sichergestellt werden, dass keine Warnungen im Log vorhanden sind.

#### Fehlender Text zur Fussnote
Wenn in einer eingelesenen Tabelle im Tabellenblatt "Internet" ein Rautezeichen (#) gefolgt von einer Zahl gefunden wird (z.B. #1), dann wird dies als Fussnote interpretiert. PyTabsam prüft dann, ob im Tabellenblatt "Fussnoten" in der Spalte B mit dem gleichen Code gefunden wird. Es werden nur Fussnoten geprüft, welche in der Spalte A den Tag "<web>" enthalten. Falls in der Spalte A nicht "<web>" steht bzw. der Fussnoten-Code in der Spalte B nicht gefunden wird, dann wird folgende Warnung ausgegeben:
```
WARNING (2022-03-08 14:29:20): The footnote #1 was referenced, but not found in the worksheet Fussnoten of T_02.02.03.xlsm
```
**Behebung:** Prüfen Sie in der Quelldatei die Übereinstimmung von Fussnoten in den Tabellenblättern "Fussnoten" und "Internet" und stellen sie sicher, dass "<web>" in der Spalte A des Tabellenblattes "Fussnoten" als Tag gesetzt ist.
       
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
|   +------------------------+  |    | read_coll_dir                create_tabsam              |    |     Sheet "T_1"           |
|   | Excel T_1              |  |    | read_all_md_fn                +- create_worksheet_expl  |    |     Sheet "T_2"           |
|   |     Sheet "Metadaten"  |  |    |  +- read_xls_metadata         +- create_worksheet       |    +---------------------------+
|   |     Sheet "Internet"   |  |    |  +- read_xls_footnote            +- read_write_data     |
|   |     Sheet "Fussnoten"  |  |    |                                     +- convert_footnote |    +---------------------------+
|   +------------------------+  |    |                                  +- prepare_footnotes   |    | Excel B                   |
|                               |    |                                  +- write_footnotes     |    .                           .
|   +------------------------+  |    | HILFSFUNKTIONEN                                         |    .                           .
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
Create and generate the destination excel files
If the destination files already exists, they will be overwritten

#### create_worksheet_expl(coll_ID, dest_file)
Read the data from the source worksheet "Internet" and write it to the destination worksheet
Copy all the format of the source worksheet "Internet"
Write the titel in the table of content in the worksheet "Inhalt"
Save the xlsx file

#### create_worksheets(coll_ID, dest_file)
Write the header of the sheet
Read the data from the source worksheet "Internet" and write it to the destination worksheet
Copy all the format of the source worksheet "Internet"
Set the uniform row height to 12.75 for the common worksheet
Write the titel in the table of content in the worksheet "Inhalt"
Save the xlsx file

#### read_write_data(source_ws, dest_ws, row_start, sheet_id):

       
       
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
