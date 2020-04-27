#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Apr 26 16:40:33 2020

@author: lelioraimondi
"""

import pandas as pd

files = ["templ1.xlsx","templ2.xlsx","templ3.xlsx"]

# Read data model 
dm = pd.read_excel("datamodel.xlsx")
sheets = list(dict.fromkeys(dm['Sheet']) )

# Generate object database

obj = pd.DataFrame (columns=["File","Sheet","Row","Column","Value"])

# Read all datapoints required in the data model and put them in the
# object database
for sheet in sheets:
    for file in files:
        cf = pd.read_excel(file, sheet_name=sheet, header=None) 
        for i in range (0,dm.shape[0]):
            Z = dm.iloc[i,0]
            X = dm.iloc[i,1]
            Y = dm.iloc[i,2]
            if (Z==sheet):
                D = cf.iloc[X-1,Y-1]
                obj = obj.append({'File' : file , 'Sheet' : Z, 
                      'Row' : X,'Column' : Y,'Value' : D,} , ignore_index=True)
        
# Write the object database to a file
    
obj.to_excel("output.xlsx", index=False)
    