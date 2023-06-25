
from math import log
# %%
# Funcion para la correlación de Standing
def boStanding(rs,gsg,gsoil,temp):
  """Esta función generada en base a la correlación de Standing permite
  calcular el factor volumétrico de formación en función a Rs, gsg, gsoil y   temperatura en R"""
  bo=0.9759+0.000120*(rs*((gsg/gsoil)**0.5)+1.25*(temp+460)-460)**1.2
  return bo

# %%
# Funcion para la correlación de Glaso
def boGlaso(rs,gsg,gsoil,temp):
  """Esta función generada en base a la correlación de Glaso permite
  calcular el factor volumétrico de formación en función a los mismos
  parámetros utilizados en Standing"""
  Bob=rs*(gsg/gsoil)**0.526+0.968*((temp+460)-460)
  A=-6.58511+2.91329*log(Bob,10)-0.27683*(log(Bob,10)**2)
  bo=1.0+10**A
  return bo

#%%
import numpy as np

#Correlación de Standing

def RsStan(API,yg,P,T):
  Plpca = P+14.7
  varx = 0.0125*API - 0.0009*(T)
  #No es necesario restarle 460 porque al sumarle el mismo valor para convertirlo en °R, se elimina y el usuario puede ingresar el valor de la Temp en °F, sin problemas
  rs = yg * (((Plpca/18.2) + 1.4) * 10**varx)**1.2048

  return rs

#Correlación Glasso

def RsGla(yg,API,T,p):
  plpca = p+14.7
  varx = 2.8869 - (14.1811 - 3.3093*np.log10(plpca))**0.5
  pb = 10**varx
  rs = yg * (((API**0.989)/(T**0.172)) * (pb))**1.2255

  return rs

# %%
# Funcion para la correlación de Beal
def uo_beal(api,temperatura):
    """Función que calcula la viscosidad del petróleo
    la viscosidad se obtiene en unidades de centipoise
    Parámetros:
    api: que son los grados api del petróleo
    temperatura: es la temperatura en unidades Ranking"""
    valor_a = (10**(0.43+(8.33/api)))
    viscosidad = (0.32 + ((1.8*10**7)/(api**4.53))*(360/(temperatura-260))**valor_a)
    print("La viscosidad es: ", viscosidad)


#%%
#Función para la correlación de Beggs-Robinson

def uo_beggs_ro(api, temperatura):
    """
    Función que calcula la viscosidad usando la correlación de Beggs-Robinson
    Parámetros:
    api: que son los grados api del petróleo
    temperatura: es la temperatura en unidades Ranking
    """
    valor_z = 3.0324 - (0.02023*api)
    valor_y = 10**valor_z
    valor_x = valor_y * ((temperatura-460)**(-1.163))
    viscosidad = 10**valor_x - 1
    print("La viscosidad es: ", viscosidad)