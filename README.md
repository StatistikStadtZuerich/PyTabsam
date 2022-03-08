# PyTabsam
Python Excel Tabellensammlung

## Bedienungsanleitung


## Installation


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
