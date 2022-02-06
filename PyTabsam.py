# -*- coding: utf-8 -*-
# PyTabsam
# @author: sszsth, sszgrm

import json
import pandas as pd
import os # operating system
import re # regular expressions

# Leere Listen vorbereiten
data_coll  = pd.DataFrame()
print(data_coll)
data_sheet = pd.DataFrame()
data_expl  = pd.DataFrame()

# Function log
# Write logging information
def log(level, text):
  print(level + ": " + text)

# Function read_config
# Read the configuration file
def read_config():
  with open('config.json', 'r', encoding="utf-8") as f:
    config = json.load(f)
    list_coll = []
    global data_coll

    for key in config:
      conf_value = config[key]
      if key == "collection":
        # The collections are provided as a list. Read them an add them do the Dataframe data_coll
        for i in range(len(conf_value)):
          coll_elem = conf_value[i]
          pk = i+1
          elem_list_coll = [pk, coll_elem["title"], coll_elem["input_path"], coll_elem["output_filename"]]
          list_coll.append(elem_list_coll)
        data_coll = pd.DataFrame(list_coll, columns = ['id', 'title' , 'input_path', 'output_filename'])

# Function create_sampledata
# Generate sample data
def create_sampledata():
  global data_coll, data_sheet, data_expl
  # COLLECTION
  list_coll = []
  elem_list_coll = [1, "Meteorologie", "O:/Projekte/PyTabsam/Testfaelle-Input/02_02", "Meteorologie.xlsx"]
  list_coll.append(elem_list_coll)
  elem_list_coll = [2, "Luftqualität", "O:/Projekte/PyTabsam/Testfaelle-Input/07_03", "Luftqualitaet.xlsx"]
  list_coll.append(elem_list_coll)
  elem_list_coll = [3, "Abfallentsorgung", "O:/Projekte/PyTabsam/Testfaelle-Input/07_02", "Abfallentsorgung.xlsx"]
  list_coll.append(elem_list_coll)
  data_coll = pd.DataFrame(list_coll, columns = ['id', 'title' , 'input_path', 'output_path'])
  
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

# Function read_coll_dir
# Open each collection path and scan for all files in it. Match names to expl. and sheets.
def read_coll_dir():
  global data_coll, data_sheet, data_expl
  list_sheet = []
  count_sheet = 0
  list_expl = []
  count_expl = 0
  data_coll = data_coll.reset_index()  # make sure indexes pair with number of rows
  for index, row in data_coll.iterrows():
    collection_id   = row['id']
    collection_path = row['input_path']
    # get a list of all the files in the directory
    files = [file for file in os.listdir(collection_path)]
    for file in files:
      # Regular Expressen for an excel file containing an explanation
      # Example: T_02.02.0.Erläuterungen.xlsm
      is_expl  = re.match("T_\d+.\d+.\d+.Erläuterungen.xlsm", file)
      # Regular Expressen for an excel file containing data
      # Examples: T_02.02.03.xlsm, T_02.02.01.2017.xlsm, T_G03.03.01.xlsm, T_G07.03.05a
      is_sheet = re.match("(T_|T_G)\d+.\d+.\d+(.xlsm|.\d+.xlsm|\w.xlsm)", file)

      # The filename matches an explanation      
      if is_expl:
        count_expl += 1
        elem_list_expl = [count_expl, collection_id, file]
        list_expl.append(elem_list_expl)
      # The filename matches a worksheet with data      
      elif is_sheet:
        count_sheet += 1
        elem_list_sheet = [count_sheet, collection_id, file]
        list_sheet.append(elem_list_sheet)
      else:
        log("WARNING", "File " + file + " in " + collection_path + " has an invalid filename. It will be ignored.")

    data_expl  = pd.DataFrame(list_expl, columns = ['ID', 'FK_collection', 'filename'])
    data_sheet = pd.DataFrame(list_sheet, columns = ['ID', 'FK_collection', 'filename'])

# Main progam
def main():
  global data_coll
  print("Read the configuration")
  read_config()
  print("Loop trough all collection directories")
  read_coll_dir()

  # The sample data will overwrite the data gathered by read_collection
  print("Create sample data")
  create_sampledata()
  
  print(data_coll)
  print(data_sheet)
  print(data_expl)

# Execute main of PyTabsam
if __name__ == '__main__':
  main()



