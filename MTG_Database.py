import tkinter as tk
import openpyxl
import pandas as pd
import numpy as np
from tkinter import filedialog
from openpyxl import Workbook
from itertools import islice

class MTGDatabase:
    mtg_data = Workbook.active()
    card_values = Workbook.active()
    coeffcients = Workbook.active()
    allcards = Workbook.active()

    allcards_dataframe = pd.DataFrame()
    coefficients_dataframe = pd.DataFrame()
    card_values_dataframe =  pd.DataFrame()

    sheet1 = 0
    sheet2 = 1
    sheet3 = 2

    #initialize database
    def __init__(self):
        self._loadData()
        #clean data
        

    def _loadData (self):
        root = tk.Tk()
        root.withdraw()
        file_path = " "
        file_path=filedialog.askopenfilename()

        MTGDatabase.mtg_data = openpyxl.load_workbook(file_path)
        MTGDatabase.card_values = self._setSheet(MTGDatabase.sheet1)
        MTGDatabase.coeffcients = self._setSheet(MTGDatabase.sheet3)
        MTGDatabase.allcards = self._setSheet(MTGDatabase.sheet2)

        MTGDatabase.allcards_dataframe = self._convertToDataframe(MTGDatabase.allcards)
        MTGDatabase.coefficients_dataframe = self._convertToDataframe(MTGDatabase.coeffcients)
        MTGDatabase.card_values_dataframe = self._convertToDataframe(MTGDatabase.card_values)



    def _setSheet (self,sheet):
        MTGDatabase.mtg_data.active = sheet
        return MTGDatabase.mtg_data.active


    def _convertToDataframe (self, sheet_data):
        data = sheet_data.values

        cols = iter(data)[1:]
        data = list(data)
        data = (islice(data_iter, 1, None) for data_iter in data)

        data_dataframe = pd.DataFrame(data, columns = cols)

        return data_dataframe


    def _cleanData (self):
        
