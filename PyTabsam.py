# -*- coding: utf-8 -*-
# PyTabsam
# @author: sszsth, sszgrm

import json
import pandas as pd

# Leere Listen vorbereiten
data_coll  = pd.DataFrame()
print(data_coll)
data_sheet = pd.DataFrame()
data_expl  = pd.DataFrame()

# Funktion read_config
# Konfiguration einlesen
def read_config():
  with open('config.json', 'r', encoding="utf-8") as f:
    config = json.load(f)
    list_coll = []
    global data_coll

    for key in config:
      conf_value = config[key]
      print("The key and value are ({}) = ({})".format(key, conf_value))
      if key == "collection":
        print("collection gefunden!")
        for i in range(len(conf_value)):
          coll_elem = conf_value[i]
          pk = i+1
          elem_list_coll = [pk, coll_elem["title"], coll_elem["input_path"], coll_elem["output_filename"]]
          list_coll.append(elem_list_coll)
        data_coll = pd.DataFrame(list_coll, columns = ['id', 'title' , 'input_path', 'output_filename'])

# Funktion create_sampledata
# Generieren von Testdaten
def create_sampledata():
  global data_coll, data_sheet, data_expl
  # COLLECTION
  # 06.02.2022 grm Testdaten auskommentiert, da der identische DataFrame über die Konfiguration erstellt wird
  # list_coll = []
  # elem_list_coll = [1, "Meteorologie", "O:/Projekte/PyTabsam/Testfaelle-Input/02_02", "Meteorologie.xlsx"]
  # list_coll.append(elem_list_coll)
  # elem_list_coll = [2, "Luftqualität", "O:/Projekte/PyTabsam/Testfaelle-Input/07_03", "Luftqualitaet.xlsx"]
  # list_coll.append(elem_list_coll)
  # elem_list_coll = [3, "Abfallentsorgung", "O:/Projekte/PyTabsam/Testfaelle-Input/07_02", "Abfallentsorgung.xlsx"]
  # list_coll.append(elem_list_coll)
  # data_coll = pd.DataFrame(list_coll, columns = ['id', 'title' , 'input_path', 'output_path'])
  
  # SHEET
  list_sheet = []
  elem_list_sheet = [1, 1, "T_02.02.01.2017.xlsm", "T02.02.01.2017", "T_2.2.1.2017", "Wetterrekorde", "Station Zürich Fluntern", "historisch und 2017", "MeteoSchweiz", 1]
  list_sheet.append(elem_list_sheet)
  data_sheet = pd.DataFrame(list_sheet, columns = ['ID', 'FK_collection', 'filename', 'sheet_name', 'code', 'title', 'subtitle1', 'subtitle2', 'source', 'order'])
  
  # EXPLANATION
  list_expl = []
  elem_list_expl = [1, 1, "T_02.02.0.Erläuterung.xlsm"]
  list_expl.append(elem_list_expl)
  data_expl = pd.DataFrame(list_expl, columns = ['ID', 'FK_collection', 'filename'])

def read_collection():
  global data_coll
  data_coll = data_coll.reset_index()  # make sure indexes pair with number of rows
  for index, row in data_coll.iterrows():
    print(row['input_path'])

# Hauptprogramm
# Define `main()` function
def main():
  global data_coll
  print("This is the main function")
  print("Read the configuration")
  read_config()
  print("Create sample data")
  create_sampledata()
  print("Loop trough all collections collections")
  read_collection()
  
  print(data_coll)
  print(data_sheet)
  print(data_expl)

# PyTabsam ausführen
# Execute `main()` function 
if __name__ == '__main__':
  main()





