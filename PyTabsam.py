# -*- coding: utf-8 -*-
# PyTabsam
# @author: sszsth, sszgrm

import json
import pandas as pd

# Leere Listen vorbereiten
list_coll  = []
list_sheet = []
list_expl  = []

# Konfiguration einlesen
with open('config.json', 'r', encoding="utf-8") as f:
    config = json.load(f)
    
# Beispiele, wie die Konfiguration genutzt werden kann:
print(config)
print(len(config["collection"]))
path_output = config['path_output']
print(path_output)



# Generieren von Testdaten
# COLLECTION
elem_list_coll = [1, "Meteorologie", "O:/Projekte/PyTabsam/Testfaelle-Input/02_02", "Meteorologie.xlsx"]
list_coll.append(elem_list_coll)
elem_list_coll = [2, "Luftqualität", "O:/Projekte/PyTabsam/Testfaelle-Input/07_03", "Luftqualitaet.xlsx"]
list_coll.append(elem_list_coll)
elem_list_coll = [3, "Abfallentsorgung", "O:/Projekte/PyTabsam/Testfaelle-Input/07_02", "Abfallentsorgung.xlsx"]
list_coll.append(elem_list_coll)
data_coll = pd.DataFrame(list_coll, columns = ['id', 'title' , 'input_path', 'output_path'])

# SHEET
elem_list_sheet = [1, 1, "T_02.02.01.2017.xlsm", "T02.02.01.2017", "T_2.2.1.2017", "Wetterrekorde", "Station Zürich Fluntern", "historisch und 2017", "MeteoSchweiz", 1]
list_sheet.append(elem_list_sheet)
data_sheet = pd.DataFrame(list_sheet, columns = ['ID', 'FK_collection', 'filename', 'sheet_name', 'code', 'title', 'subtitle1', 'subtitle2', 'source', 'order'])

# EXPLANATION
elem_list_expl = [1, 1, "T_02.02.0.Erläuterung.xlsm"]
list_expl.append(elem_list_expl)
data_expl = pd.DataFrame(list_expl, columns = ['ID', 'FK_collection', 'filename'])


print(data_coll)
print(data_sheet)
print(data_expl)




