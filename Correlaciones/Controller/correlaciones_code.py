# Import Python libraries
import xlwings as xw
import pandas as pd
import numpy as np
from Correlaciones.Model.funciones_de_correlaciones import boStanding
from Correlaciones.Model.funciones_de_correlaciones import boGlaso
from Correlaciones.Model.funciones_de_correlaciones import RsStan
from Correlaciones.Model.funciones_de_correlaciones import RsGla
from Correlaciones.Model.funciones_de_correlaciones import uo_beal
from Correlaciones.Model.funciones_de_correlaciones import uo_beggs_ro
from Correlaciones.Model.funciones_de_correlaciones import pbStanding
from Correlaciones.Model.funciones_de_correlaciones import pbGlaso
import matplotlib.pyplot as plt
from matplotlib import ticker
import seaborn as sns


# Define sheets
SHEET_SUMMARY = "Datos"

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
DF_RESULTADOS1 = "df_resultados1"
DF_RESUL1_MUES1 = "res1_muestra1"
DF_RESUL1_MUES2 = "res1_muestra2"
DF_RESUL1_MUES3 = "res1_muestra3"
DF_RESUL1_MUES4 = "res1_muestra4"
DF_RESUL1_MUES5 = "res1_muestra5"
DF_RESUL1_MUES6 = "res1_muestra6"
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
    df_resul1_mues1 = sheet[DF_RESUL1_MUES1].options(pd.DataFrame, index=False,
                                                   expand="table").value
    df_resul2_mues2 = sheet[DF_RESUL1_MUES2].options(pd.DataFrame, index=False,
                                                   expand="table").value
    df_RESULTADOS1 = sheet[DF_RESULTADOS1].options(pd.DataFrame, index=False,
                                                   expand="table").value
    for index, row in df_DATA.iterrows():
        for index1, row1 in df_RESULTADOS1.iterrows():
            temperatura = int(row[0])
            pburbuja = int(row[1])
            Rs_medido = int(row[2])
            grados_api = int(row[3])
            sg_gas = int(row[4])

            Rs1 = RsStan(grados_api, sg_gas, pburbuja, temperatura)
            Pb1 = pbStanding(Rs_medido, sg_gas, grados_api, temperatura)
            Bo1 = boStanding(Rs_medido, sg_gas, grados_api, temperatura)
            Uo1 = uo_beal(grados_api, temperatura)
            sheet[DF_RESULTADOS1].value = Rs1, Pb1, Bo1, Uo1



if __name__ == "__main__":
    xw.Book("correlaciones.xlsm").set_mock_caller()
    main()
