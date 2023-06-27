# Import Python libraries
import xlwings as xw
import pandas as pd
import numpy as np
from Correlaciones.Model.funciones_de_correlaciones import boStanding,boGlaso,RsStan,RsGla,uo_beal,uo_beggs_ro,pbStanding,pbGlaso
import matplotlib.pyplot as plt
from matplotlib import ticker
import seaborn as sns
####


# Define sheets
SHEET_SUMMARY = "sheet1"

# Define column names
MUESTRA = "muestras"
T = "temperaturas"
Pb = "pburbuja"
Rs = "rs_medido"
API = "grad_api"
Yg = "gs_gas"

# Index of Muestras Parameters
M1, M2, M3, M4, M5, M6 = 0, 1, 2, 3, 4, 5




def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[SHEET_SUMMARY]

    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("correlaciones.xlsm").set_mock_caller()
    main()
