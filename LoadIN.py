import tkinter as tk
import pandas as pd
import xlrd # not in use
import openpyxl 
import numpy as np
from tkinter import filedialog
from itertools import islice

root = tk.Tk()
root.withdraw()

file_path = " "
file_path=filedialog.askopenfilename()
sheet1 = 0
sheet2 = 1
sheet3 = 2

mtg_data = openpyxl.load_workbook(file_path)
card_values = mtg_data.active

mtg_data.active = sheet3
coeffcients = mtg_data.active

mtg_data.active = sheet2
allcards = mtg_data.active

allcards_data = allcards.values
coefficients_data = coeffcients.values
card_values_data = card_values.values

#get workbook headers and reformat worksheet
#will need to make a fuction for other dtaframes
cols = next(allcards_data)[1:]
allcards_data = list(allcards_data)
allcards_data = (islice(data, 1, None) for data in allcards_data)


cols_coeffcients = next(coefficients_data)[1:]
coefficient_data = list(coefficients_data)
coefficient_data = (islice(data, 1, None) for data in coefficient_data)

cols_card_values = next(card_values_data)[1:]
print(cols_card_values)
card_values_data = list(card_values_data)
card_values_data = (islice(data, 1, None) for data in card_values_data)


#Create data frames
allcards_dataframe = pd.DataFrame(allcards_data, columns = cols)
coefficients_dataframe = pd.DataFrame(coefficient_data, columns = cols_coeffcients)
card_values_dataframe = pd.DataFrame(card_values_data, columns = cols_card_values)

#filter out unneeded columns from allcards_dataframe
allcards_dataframe = allcards_dataframe[allcards_dataframe.columns.difference(['index', 'id', 'name', 'types', 'convertedManaCost', 'keyword1', 'keyword2', 'intrinsic Value'])]

coefficients_dataframe = coefficients_dataframe[coefficients_dataframe.columns.difference(['Keyword'])]


#compute matrix multplication on coffcient values and keyword values
coeffcient_values = coefficients_dataframe.to_numpy()

result = allcards_dataframe.dot(coeffcient_values)

card_values_dataframe['TrueValue'] =  result

#card_values_dataframe.to_excel("D:\Documents\MagicData\Results.xlsx")

#print(card_values_dataframe.head(10))

#result = result.to_numpy()

#result = np.reshape(result,len(result))

#print(result)
#print("this is the result length " + str(len(result)) )
#print(allcards_dataframe.head(2))