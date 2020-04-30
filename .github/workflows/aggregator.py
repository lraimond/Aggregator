#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Apr 26 16:40:33 2020

@author: lelioraimondi
"""
import os as os
import pandas as pd

basecwd = os.getcwd()
srccwd = basecwd + "/source_templates"
os.chdir(srccwd)
rfiles = os.listdir()
files = list(filter(lambda x: x.endswith("xlsx"), rfiles))

# Read data model and produce list of sheets
os.chdir(basecwd)
dm = pd.read_excel("datamodel.xlsx")
sheets = list(dict.fromkeys(dm['Sheet']) )

# Generate object database

obj = pd.DataFrame (columns=["File","Sheet","Row","Column","Value"])

# Read all datapoints required in the data model and put them in the
# object database

os.chdir(srccwd)
for sheet in sheets:
    subdm = dm[dm.Sheet == sheet]
    for file in files:
        cf = pd.read_excel(file, sheet_name=sheet, header=None) 
        for i in range (0,subdm.shape[0]): 
            X = subdm.iloc[i,1]
            Y = subdm.iloc[i,2]
            D = cf.iloc[X-1,Y-1]
            obj = obj.append({'File' : file , 'Sheet' : sheet, 
                      'Row' : X,'Column' : Y,'Value' : D,} , ignore_index=True)
        
# Write the object database to a file

os.chdir(basecwd)    
obj.to_excel("output.xlsx", index=False)
    