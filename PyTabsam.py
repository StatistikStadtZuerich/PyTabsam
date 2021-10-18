# -*- coding: utf-8 -*-
"""
PyTabsam
Created on Mon Oct 18 10:08:56 2021

@author: sszsth, sszgrm
"""

import json

# Konfiguration einlesen
with open('config.json', 'r', encoding="utf-8") as f:
    config = json.load(f)
# Beispiele, wie die Konfiguration genutzt werden kann:
print(config)
print(len(config["list_input"]))
path_output = config['path_output']
print(path_output)


