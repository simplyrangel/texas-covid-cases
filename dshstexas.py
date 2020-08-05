"""Functions to read and format Texas COVID-19 data provided by the 
Texas Department of State Health Services.

Data available online at: 
https://dshs.texas.gov/coronavirus/additionaldata.aspx"""
import numpy as np
import pandas as pd
from datetime import datetime

def read_testing(fi):
	df = pd.read_excel(
		fi, 
		index_col=[0], 
		skiprows=[0], 
		nrows=254,
		na_values=["--", "-"])
	newcols = []
	for col in df.columns:
		x = col.split("Through ")[-1].replace("*", "")		
		x = "%s 2020" %x
		newcols.append(datetime.strptime(x, "%B %d %Y"))
	df.columns = newcols
	return df	

def read_cases(fi):
	df = pd.read_excel(
		fi, 
		index_col=[0], 
		skiprows=[0,1], 
		nrows=254)
	newcols = []
	for col in df.columns:
		x = col.split("\n")[-1]
		x = x.split("Cases")[-1]
		x = x.replace(" ", "").replace("*", "")
		if x == "Population":
			newcols.append(x)
		else:
			x = "%s 2020" %x
			x = datetime.strptime(x, "%m-%d %Y")
			newcols.append(x.strftime("%Y-%m-%d"))
	df.columns = newcols
	return df

def read_deaths(fi):
    df = pd.read_excel(
        fi, 
        index_col=[0],
        skiprows=[0,1],
        na_values=["."])
    df = df.dropna(axis=0, how="all") #remove metadata or empty rows
    
    # only some column headers are broken:
    new_cols = []
    for col in df.columns:
        x = col
        if type(col) is str:
            x = col.replace("`", "")
            x = datetime.strptime(x, "%m/%d/%Y")
        x = x.strftime("%Y-%m-%d")
        new_cols.append(x)
    df.columns = new_cols
    
    # massage index:
    new_index = df.index.tolist()
    new_index = [x.lower().capitalize() for x in new_index]
    df.index = new_index
    """
    newcols = []
    for col in df.columns:
        x = col.replace("Fatalities", "")
        x = x.replace(" ", "")
        if x == "Population":
            newcols.append(x)
        else:
            x = x.replace("`", "")
            x = datetime.strptime(x, "%m/%d/%Y")
            newcols.append(x.strftime("%Y-%m-%d"))
    df.columns = newcols
    """
    return df
	
	
	
	
	
	
	
