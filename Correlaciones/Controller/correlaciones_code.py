# Import Python libraries
import xlwings as xw
import pandas as pd
import numpy as np
from Correlaciones.Model.Funciones_de_correlaciones import boStanding,boGlaso,RsStan,RsGla,uo_beal,uo_beggs_ro,pbStanding,pbGlaso
import matplotlib.pyplot as plt
from matplotlib import ticker
import seaborn as sns

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
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
