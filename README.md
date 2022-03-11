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
