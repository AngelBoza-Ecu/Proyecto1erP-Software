# Import Python libraries
import xlwings as xw
import pandas as pd
import numpy as np
from Correlaciones.Model.funciones_de_correlaciones import boStanding
from Correlaciones.Model.funciones_de_correlaciones import boAlMarhoun
from Correlaciones.Model.funciones_de_correlaciones import RsStan
from Correlaciones.Model.funciones_de_correlaciones import RsGla
from Correlaciones.Model.funciones_de_correlaciones import uo_beal
from Correlaciones.Model.funciones_de_correlaciones import uo_beggs_ro
from Correlaciones.Model.funciones_de_correlaciones import pbStanding
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
M1, M2, M3, M4, M5, M6 = 0, 1, 2, 3, 4, 5

# Data
DF_DATOS = "df_datos"
MUESTRA1 = "muestra1"
DF_DATOS1 = "valores"
DF_DATOS2 = "valores2"
DF_RSSTAN1 = "Rs_Stan1"
DF_PBSTAN1 = "Pb_Stan1"
DF_BOSTAN1 = "Bo_Stan1"
DF_UOBEAL1 = "Uo_Beal1"
DF_RSGLA_gra = "Rs_Glas1"
DF_PBALM = "Pb_AlM"
DF_BOALM = "Bo_AlM"
DF_UOBEGGS = "Uo_Beggs2"

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

    df_DATA = sheet[DF_DATOS1].options(np.array, transpose=True).value

    # Muestra 1
    sheet[DF_RSSTAN1].options(np.array, transpose=True).value = RsStan(*df_DATA[0:4])
    sheet[DF_PBSTAN1].options(np.array, transpose=True).value = pbStanding(
        *df_DATA[1:5])
    sheet[DF_BOSTAN1].options(np.array, transpose=True).value = boStanding(
        *df_DATA[1:5])
    sheet[DF_UOBEAL1].options(np.array, transpose=True).value = uo_beal(*df_DATA[2:4])

    # Muestra 2
    df_DATA1 = sheet[DF_DATOS2].options(np.array, transpose=True).value
    sheet[DF_RSGLA_gra].options(np.array, transpose=True).value = RsGla(*df_DATA1[0:4])
    sheet[DF_PBALM].options(np.array, transpose=True).value = pbAlMarhoun(
        *df_DATA1[1:5])
    sheet[DF_BOALM].options(np.array, transpose=True).value = boAlMarhoun(
        *df_DATA1[1:5])
    sheet[DF_UOBEGGS].options(np.array, transpose=True).value = uo_beggs_ro(
        *df_DATA1[2:4])

    # Generación de graficos
    # Crear la figura y el gráfico Pb vs Rs_Stan
    x = sheet[Pb].value
    y = sheet[DF_RSGLA_gra].value
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)

    # Asignar valores a las variables en el gráfico
    ax.scatter(x, y, color='red', label='Valores obtenidos')

    # Personalizar el gráfico
    ax.set_xlabel('Presión de burbuja')
    ax.set_ylabel('Rs_Standing')
    ax.set_title('Presión de burbuja vs Rs_Stan')
    # Mostrar la leyenda
    ax.legend()
    z = np.polyfit(x, y,
                   1)  # Ajuste de una línea de primer grado (puede ajustarse a otro
    # grado si se desea)
    p = np.poly1d(z)
    ax.plot(x, p(x), "r--")  # Trazar la línea de tendencia en rojo discontinuo

    # Mostrar el gráfico
    sheet.pictures.add(fig, name='Pb vs Rs_STan', update=True,
                       left=sheet.range("B20").left, top=sheet.range("B20").top)

    # Gráfica Pb_Stan vs T
    # Crear la figura y el gráfico Pb vs Rs_Stan
    x = sheet[T].value
    y = sheet[DF_PBSTAN1].value
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)

    # Asignar valores a las variables en el gráfico
    ax.scatter(x, y, color='red', label='Valores obtenidos')

    # Personalizar el gráfico
    ax.set_xlabel('Temperatura')
    ax.set_ylabel('Pb_Sta')
    ax.set_title('Pb_Stan vs Temperatura')
    # Mostrar la leyenda
    ax.legend()
    z = np.polyfit(x, y,
                   1)  # Ajuste de una línea de primer grado (puede ajustarse a otro
    # grado si se desea)
    p = np.poly1d(z)
    ax.plot(x, p(x), "r--")  # Trazar la línea de tendencia en rojo discontinuo

    # Mostrar el gráfico
    sheet.pictures.add(fig, name='T vs Pb_STan', update=True,
                       left=sheet.range("B40").left, top=sheet.range("B40").top)

    # Gráfica Bo_Stan vs T
    x = sheet[T].value
    y = sheet[DF_BOSTAN1].value
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)

    # Asignar valores a las variables en el gráfico
    ax.scatter(x, y, color='red', label='Valores obtenidos')

    # Personalizar el gráfico
    ax.set_xlabel('Temperatura')
    ax.set_ylabel('Bo_Stan')
    ax.set_title('Bo_Stan vs Temperatura')
    # Mostrar la leyenda
    ax.legend()
    z = np.polyfit(x, y,
                   1)  # Ajuste de una línea de primer grado (puede ajustarse a otro
    # grado si se desea)
    p = np.poly1d(z)
    ax.plot(x, p(x), "r--")  # Trazar la línea de tendencia en rojo discontinuo

    # Mostrar el gráfico
    sheet.pictures.add(fig, name='Bo_STan vs T', update=True,
                       left=sheet.range("B60").left, top=sheet.range("B60").top)

    # Gráfica Uo_Beal vs T
    x = sheet[T].value
    y = sheet[DF_UOBEAL1].value
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)

    # Asignar valores a las variables en el gráfico
    ax.scatter(x, y, color='red', label='Valores obtenidos')

    # Personalizar el gráfico
    ax.set_xlabel('Temperatura')
    ax.set_ylabel('Uo_Beal')
    ax.set_title('Uo_Beal vs Temperatura')
    # Mostrar la leyenda
    ax.legend()
    z = np.polyfit(x, y,
                   1)  # Ajuste de una línea de primer grado (puede ajustarse a otro
    # grado si se desea)
    p = np.poly1d(z)
    ax.plot(x, p(x), "r--")  # Trazar la línea de tendencia en rojo discontinuo

    # Mostrar el gráfico
    sheet.pictures.add(fig, name='Uo_Beal vs T', update=True,
                       left=sheet.range("B80").left, top=sheet.range("B80").top)

    # Gráfica Rs_Glasso vs Pburbuja
    x = sheet[Pb].value
    y = sheet[DF_RSGLA_gra].value
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)

    # Asignar valores a las variables en el gráfico
    ax.scatter(x, y, color='red', label='Valores obtenidos')

    # Personalizar el gráfico
    ax.set_xlabel('Pburbuja')
    ax.set_ylabel('Rs_Glasso')
    ax.set_title('Rs_Glasso vs Pburbuja')
    # Mostrar la leyenda
    ax.legend()
    z = np.polyfit(x, y,
                   1)  # Ajuste de una línea de primer grado (puede ajustarse a otro
    # grado si se desea)
    p = np.poly1d(z)
    ax.plot(x, p(x), "r--")  # Trazar la línea de tendencia en rojo discontinuo

    # Mostrar el gráfico
    sheet.pictures.add(fig, name='Rs_Glasso vs Pburbuja', update=True,
                       left=sheet.range("I20").left, top=sheet.range("I20").top)

    # Gráfica Pb_AlM vs T
    x = sheet[T].value
    y = sheet[DF_PBALM].value
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)

    # Asignar valores a las variables en el gráfico
    ax.scatter(x, y, color='red', label='Valores obtenidos')

    # Personalizar el gráfico
    ax.set_xlabel('Temperatura')
    ax.set_ylabel('Pb_AlM')
    ax.set_title('Pb_AlM vs Temperatura')
    # Mostrar la leyenda
    ax.legend()
    z = np.polyfit(x, y,
                   1)  # Ajuste de una línea de primer grado (puede ajustarse a otro
    # grado si se desea)
    p = np.poly1d(z)
    ax.plot(x, p(x), "r--")  # Trazar la línea de tendencia en rojo discontinuo

    # Mostrar el gráfico
    sheet.pictures.add(fig, name='Pb_AlM vs T', update=True,
                       left=sheet.range("I40").left, top=sheet.range("I40").top)

    # Gráfica Bo_AlM vs T
    x = sheet[T].value
    y = sheet[DF_BOALM].value
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)

    # Asignar valores a las variables en el gráfico
    ax.scatter(x, y, color='red', label='Valores obtenidos')

    # Personalizar el gráfico
    ax.set_xlabel('Temperatura')
    ax.set_ylabel('Bo_AlM')
    ax.set_title('Bo_AlM vs Temperatura')
    # Mostrar la leyenda
    ax.legend()
    z = np.polyfit(x, y,
                   1)  # Ajuste de una línea de primer grado (puede ajustarse a otro
    # grado si se desea)
    p = np.poly1d(z)
    ax.plot(x, p(x), "r--")  # Trazar la línea de tendencia en rojo discontinuo

    # Mostrar el gráfico
    sheet.pictures.add(fig, name='Bo_AlM vs T', update=True,
                       left=sheet.range("I60").left, top=sheet.range("I60").top)

    # Gráfica Uo_Beggs vs T
    x = sheet[T].value
    y = sheet[DF_UOBEGGS].value
    fig = plt.figure()
    ax = fig.add_subplot(1, 1, 1)

    # Asignar valores a las variables en el gráfico
    ax.scatter(x, y, color='red', label='Valores obtenidos')

    # Personalizar el gráfico
    ax.set_xlabel('Temperatura')
    ax.set_ylabel('Uo_Beggs')
    ax.set_title('Uo_Beggs vs Temperatura')
    # Mostrar la leyenda
    ax.legend()
    z = np.polyfit(x, y,
                   1)  # Ajuste de una línea de primer grado (puede ajustarse a otro
    # grado si se desea)
    p = np.poly1d(z)
    ax.plot(x, p(x), "r--")  # Trazar la línea de tendencia en rojo discontinuo

    # Mostrar el gráfico
    sheet.pictures.add(fig, name='Uo_Beggs vs T', update=True,
                       left=sheet.range("I80").left, top=sheet.range("I80").top)


if __name__ == "__main__":
    xw.Book("correlaciones.xlsm").set_mock_caller()
    main()
