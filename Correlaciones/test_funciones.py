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

#%% #%% Call workbook
xb = xw.Book("Correlaciones/Controller/correlaciones.xlsm")
sheet = xb.sheets["Datos"]

#%% Call dataframe
