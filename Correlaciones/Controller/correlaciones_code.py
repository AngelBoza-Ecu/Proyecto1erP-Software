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

# Data
DF_DATOS = "df_datos"
VALORES = "valores"
# Result cells # Call range cells from MS Excel
BO_STANDING = "Bo_Stan"
BO_GLASO = "Bo_Glas"
RS_STANDING = "Rs_Stan"
RS_GLASO = "Rs_Glas"
PB_STANDING = "Pb_Stan"
PB_GLASO = "Pb_Glas"
UO_BEAL = "uo_beal"
UO_BEGGS = "uo_beggs"


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[SHEET_SUMMARY]

    df_DATA = sheet[DF_DATOS].options(pd.DataFrame, index=False, expand="table").value
    params = sheet[VALORES].options(np.array, transpose=True).value





if __name__ == "__main__":
    xw.Book("correlaciones.xlsm").set_mock_caller()
    main()
