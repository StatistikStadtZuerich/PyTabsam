# -*- coding: utf-8 -*-
# PyTabsam
# @author: sszsth, sszgrm

import json
import pandas as pd
import os # operating system
import re # regular expressions
from openpyxl import load_workbook
# @Hansj�rg: Diese Seite k�nnte interessant sein um Inhalte zu kopieren:
# https://stackoverflow.com/questions/42344041/how-to-copy-worksheet-from-one-workbook-to-another-one-using-openpyxl

# Leere Listen vorbereiten
data_coll  = pd.DataFrame([],dtype=pd.StringDtype())
count_sheet = 0
data_sheet = pd.DataFrame([],dtype=pd.StringDtype())
count_expl = 0
data_expl  = pd.DataFrame([],dtype=pd.StringDtype())

# Function tolog
# Write logging information
def tolog(level, text):
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
  global data_coll, count_sheet, data_sheet, count_expl, data_expl
  # COLLECTION
  list_coll = []
  elem_list_coll = [1, "Meteorologie", "O:/Projekte/PyTabsam/Testfaelle-Input/02_02", "Meteorologie.xlsx"]
  list_coll.append(elem_list_coll)
  elem_list_coll = [2, "Luftqualit�t", "O:/Projekte/PyTabsam/Testfaelle-Input/07_03", "Luftqualitaet.xlsx"]
  list_coll.append(elem_list_coll)
  elem_list_coll = [3, "Abfallentsorgung", "O:/Projekte/PyTabsam/Testfaelle-Input/07_02", "Abfallentsorgung.xlsx"]
  list_coll.append(elem_list_coll)
  data_coll = pd.DataFrame(list_coll, columns = ['id', 'title' , 'input_path', 'output_path'])
  
  # SHEET
  list_sheet = []
  elem_list_sheet = [1, 1, "T_02.02.01.2017.xlsm", "O:/Projekte/PyTabsam/Testfaelle-Input/02_02", "T02.02.01.2017", "T_2.2.1.2017", "Wetterrekorde", "Station Z�rich Fluntern", "historisch und 2017", "MeteoSchweiz", 1]
  list_sheet.append(elem_list_sheet)
  data_sheet = pd.DataFrame(list_sheet, columns = ['ID', 'FK_collection', 'filename', 'directory', 'sheet_name', 'code', 'title', 'subtitle1', 'subtitle2', 'source', 'order'])
  count_sheet = 1
  
  # EXPLANATION
  list_expl = []
  elem_list_expl = [1, 1, "T_02.02.0.Erl�uterung.xlsm", "O:/Projekte/PyTabsam/Testfaelle-Input/02_02"]
  list_expl.append(elem_list_expl)
  data_expl = pd.DataFrame(list_expl, columns = ['ID', 'FK_collection', 'filename', 'directory'])
  count_expl = 1

# Function read_coll_dir
# Open each collection path and scan for all files in it. Match names to expl. and sheets.
def read_coll_dir():
  global data_coll, count_sheet, data_sheet, count_expl, data_expl
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
    for filename in files:
      is_expl = re.match(r'T_\d+.\d+.\d+.Erl�uterungen.xlsm', filename)
      # Regular Expressen for an excel file containing an explanation
      # Example: T_02.02.0.Erl�uterungen.xlsm
      is_sheet = re.match(r'(T_|T_G)\d+.\d+.\d+(.xlsm|.\d+.xlsm|\w.xlsm)', filename)
      # Regular Expressen for an excel file containing data
      # Examples: T_02.02.03.xlsm, T_02.02.01.2017.xlsm, T_G03.03.01.xlsm, T_G07.03.05a

      # The filename matches an explanation
      if is_expl:
        count_expl += 1
        elem_list_expl = [count_expl, collection_id, filename, collection_path]
        list_expl.append(elem_list_expl)
      # The filename matches a worksheet with data
      elif is_sheet:
        count_sheet += 1
        elem_list_sheet = [count_sheet, collection_id, filename,collection_path]
        list_sheet.append(elem_list_sheet)
      else:
        tolog("WARNING", "File " + filename + " in " + collection_path + " has an invalid filename. It will be ignored.")
        
    data_expl  = pd.DataFrame(list_expl, columns = ['ID', 'FK_collection', 'filename', 'directory'])
    data_sheet = pd.DataFrame(list_sheet, columns = ['ID', 'FK_collection', 'filename', 'directory'])

# Function read_xls_expl
# Read excel that contains the explanation
def read_xls_expl():
  # Test to read single cells from one sheet
  wb = load_workbook(filename = r'O:\Projekte\PyTabsam\Testfaelle-Input\07_03\T_07.03.0.Erl�uterungen.xlsm', read_only=True)
  sheet_ranges = wb['Internet']
  print(sheet_ranges['A1'].value)
  print(sheet_ranges['B1'].value)
  print(sheet_ranges['A2'].value)
  print(sheet_ranges['B2'].value)

# Function read_xls_metadata
# Read metadata out of excel 
def read_xls_metadata(sheet_id):
  global data_sheet
  xls_filename = data_sheet.loc[data_sheet['ID'] == sheet_id, 'filename'].values[0]
  xls_directory = data_sheet.loc[data_sheet['ID'] == sheet_id, 'directory'].values[0]
  xls_path = xls_directory + "/" + xls_filename
  #print("Excel: "+ xls_path)
  wb = load_workbook(filename = xls_path)
  sheet_md = wb['Metadaten']
  colkey = ""
  for (row, col), source_cell in sheet_md._cells.items():
    #print(str(row) + " - "+ str(col) + ": " + str(source_cell._value))
    if col==1:
      colkey = str(source_cell._value).lower()
    # add values from the metadata sheet directly to the dataframe by ID
    if col==2:
      if colkey=="tabellencode":
        #print("Tabellencode: " + str(source_cell._value))
        data_sheet.loc[data_sheet['ID'] == sheet_id, ['code']] = str(source_cell._value)
      if colkey=="tabellentitel":
        #print("Tabellentitel: " + str(source_cell._value))
        data_sheet.loc[data_sheet['ID'] == sheet_id, ['title']] = str(source_cell._value)
      if colkey=="tabellenuntertitel1":
        #print("Tabellenuntertitel1: " + str(source_cell._value))
        data_sheet.loc[data_sheet['ID'] == sheet_id, ['subtitle1']] = str(source_cell._value)
      if colkey=="tabellenuntertitel2":
        #print("Tabellenuntertitel2: " + str(source_cell._value))
        data_sheet.loc[data_sheet['ID'] == sheet_id, ['subtitle2']] = str(source_cell._value)
      if colkey=="quelle":
        #print("Quelle: " + str(source_cell._value))
        data_sheet.loc[data_sheet['ID'] == sheet_id, ['source']] = str(source_cell._value)

# Function read_all_metadata
# Read metadata of all sheets in dataframe 
def read_all_metadata():
  global count_sheet, data_sheet
  #print(count_sheet)
  sheet_id = 0
  while sheet_id < count_sheet:
    sheet_id += 1
    read_xls_metadata(sheet_id)


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
  
  #read_xls_expl()
  print("Open all excel sheets, read metadata and add to dataframe")
  read_all_metadata()
  
  #print(data_coll)
  #print(data_sheet)
  #print(data_expl)

# Execute main of PyTabsam
if __name__ == '__main__':
  main()



