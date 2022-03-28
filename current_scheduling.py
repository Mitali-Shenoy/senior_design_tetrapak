
import pandas as pd
import numpy as np
import tkinter as tk 

from tkinter import filedialog

tk.Tk().withdraw() # prevents an empty tkinter window from appearing

#asking the user to select their input file
file_path = filedialog.askopenfilename()

print(file_path)

#converting input excel file into a dataframe 
df = pd.read_csv(file_path)

print(df)

#cleaning up input file 
df = df.dropna(axis=0)
df.columns = df.iloc[0]
df = df.drop(df.index[0])

#print(df)

# we only care about WIP's that are in the warehouse after the lamination stage ('WIP Printing-Lamina')
df = df[df['Material']=='WIP Printing-Lamina']
#print(df)

# filtering the dataframe to only show values which we want for current rules implementation (Order, Lanes, No of Rolls, POrder Due Date)
df = df[[" Order", " Lanes", " No Of Rolls", " POrder Due Date"]]
df = df.astype({" Order": str, " Lanes": int, " No Of Rolls": int})
df[" POrder Due Date"] = pd.to_datetime(df[" POrder Due Date"])
print(df)
#print(df.dtypes)


# Implementation of current 5 scheduling rules
