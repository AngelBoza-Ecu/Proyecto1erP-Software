# Import Python libraries
import xlwings as xw
import pandas as pd
import numpy as np
from Correlaciones.Model.funciones_de_correlaciones import boStanding
from Correlaciones.Model.funciones_de_correlaciones import boGlaso
from Correlaciones.Model.funciones_de_correlaciones import boAlMarhoun
from Correlaciones.Model.funciones_de_correlaciones import RsStan
from Correlaciones.Model.funciones_de_correlaciones import RsGla
from Correlaciones.Model.funciones_de_correlaciones import uo_beal
from Correlaciones.Model.funciones_de_correlaciones import uo_beggs_ro
from Correlaciones.Model.funciones_de_correlaciones import pbStanding
from Correlaciones.Model.funciones_de_correlaciones import pbGlaso
from Correlaciones.Model.funciones_de_correlaciones import pbAlMarhoun

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
#M1, M2, M3, M4, M5, M6 = 0, 1, 2, 3, 4, 5

# Data
DF_DATOS = "df_datos"
MUESTRA1 = "muestra1"
DF_DATOS1 = "valores"
DF_DATOS2 = "valores2"
DF_RSSTAN1 = "Rs_Stan1"
DF_PBSTAN1 = "Pb_Stan1"
DF_BOSTAN1 = "Bo_Stan1"
DF_UOBEAL1 = "Uo_Beal1"
DF_RSGLA11 = "Rs_Gla11"
DF_PBALM = "Pb_AlM"
DF_BOALM = "Bo_AlM"
DF_UOBEGGS = "Uo_Beggs2"


# DF_PBGLA11 = "Pb_Gla11"
#
# DF_RSSTANM1 = "Rs_StanM1"
# DF_PBSTANM1 = "Pb_StanM1"
# DF_BOSTANM1 = "Bo_StanM1"
# DF_UOBEALM1 = "Uo_BealM1"
#
# DF_RESULTADOS1 = "df_resultados1"
# DF_RESUL1_MUES1 = "res1_muestra1"
# DF_RESUL1_MUES2 = "res1_muestra2"
# DF_RESUL1_MUES3 = "res1_muestra3"
# DF_RESUL1_MUES4 = "res1_muestra4"
# DF_RESUL1_MUES5 = "res1_muestra5"
# DF_RESUL1_MUES6 = "res1_muestra6"
#VALORES = "valores"


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

    df_DATA = sheet[DF_DATOS1].options(np.array,transpose=True).value

    #Muestra 1
    sheet[DF_RSSTAN1].options(np.array,transpose=True).value = RsStan(*df_DATA[0:4])
    sheet[DF_PBSTAN1].options(np.array,transpose=True).value = pbStanding(*df_DATA[1:5])
    sheet[DF_BOSTAN1].options(np.array,transpose=True).value = boStanding(*df_DATA[1:5])
    sheet[DF_UOBEAL1].options(np.array,transpose=True).value = uo_beal(*df_DATA[2:4])

    #Muestra 2
    df_DATA1 = sheet[DF_DATOS2].options(np.array, transpose=True).value
    sheet[DF_RSGLA11].options(np.array,transpose=True).value = RsGla(*df_DATA1[0:4])
    sheet[DF_PBALM].options(np.array,transpose=True).value = pbAlMarhoun(*df_DATA1[1:5])
    sheet[DF_BOALM].options(np.array,transpose=True).value = boAlMarhoun(*df_DATA1[1:5])
    sheet[DF_UOBEGGS].options(np.array,transpose=True).value = uo_beggs_ro(*df_DATA1[2:4])


    # sheet[DF_RESUL1_MUES1].value = pbStanding(*df_DATA[1])
    # sheet[DF_RESUL1_MUES1].value = boStanding(*df_DATA[2])
    # sheet[DF_RESUL1_MUES1].value = uo_beal(*df_DATA[3])

    # params = sheet[VALORES].options(np.array, transpose=True).value
    # 3df_resul1_mues1 = sheet[DF_RESUL1_MUES1].options(pd.DataFrame, index=False,
    #                                                expand="table").value
    # df_resul2_mues2 = sheet[DF_RESUL1_MUES2].options(pd.DataFrame, index=False,
    #                                                expand="table").value
    # df_RESULTADOS1 = sheet[DF_RESULTADOS1].options(pd.DataFrame, index=False,
    #                                                expand="table").value
    # for index, row in df_DATA.iterrows():
    #     for index1, row1 in df_RESULTADOS1.iterrows():
    #         temperatura = int(row[0])
    #         pburbuja = int(row[1])
    #         Rs_medido = int(row[2])
    #         grados_api = int(row[3])
    #         sg_gas = int(row[4])
    #
    #         Rs1 = RsStan(grados_api, sg_gas, pburbuja, temperatura)
    #         Pb1 = pbStanding(Rs_medido, sg_gas, grados_api, temperatura)
    #         Bo1 = boStanding(Rs_medido, sg_gas, grados_api, temperatura)
    #         Uo1 = uo_beal(grados_api, temperatura)
    #         sheet[DF_RESULTADOS1].value = Rs1, Pb1, Bo1, Uo1



if __name__ == "__main__":
    xw.Book("correlaciones.xlsm").set_mock_caller()
    main()
