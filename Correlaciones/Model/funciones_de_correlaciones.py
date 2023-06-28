
import numpy as np
from math import log

# %%
# Funcion para la correlación de Standing
def boStanding(gsg,temp,API,rs):
  """Esta función generada en base a la correlación de Standing permite
  calcular el factor volumétrico de formación en función a Rs, gsg, gsoil y   temperatura en R"""
  bo=0.9759+0.000120*(rs*((gsg/(141.5/131.5+API))**0.5)+1.25*(temp+460)-460)**1.2
  return bo

# %%
# Funcion para la correlación de Glaso
def boGlaso(gsg,temp,API,rs):
  """Esta función generada en base a la correlación de Glaso permite
  calcular el factor volumétrico de formación en función a los mismos
  parámetros utilizados en Standing"""
  Bob=rs*(gsg/(141.5/131.5+API))**0.526+0.968*((temp+460)-460)
  A=-6.58511+2.91329*np.log10(Bob)-0.27683*(np.log10(Bob)**2)
  bo=1.0+10**A
  return bo

#%%
#Correlación de Standing

def RsStan(P,yg,T,API):
  """Esta función pormite calcular la solubilidad del gas mediante uso de la correlación de Standing,
  dicho función depende de los siguientes parámetros que deben ser ingresados por el usuario:
  yg: gravedad específica del gas
  API: grados API del oil
  T: temperatura en °F, porque la correlación mismo hace la conversión eliminándola
  p: presión en lpc ya que la función la convierte en absoluta sumando 14.7"""
  Plpca = P+14.7
  varx = 0.0125*API - 0.0009*(T)
  #No es necesario restarle 460 porque al sumarle el mismo valor para convertirlo en °R, se elimina y el usuario puede ingresar el valor de la Temp en °F, sin problemas
  rs = yg * (((Plpca/18.2) + 1.4) * 10**varx)**1.2048

  return rs

#%%
#Correlación Glasso

def RsGla(P,yg,T,API):
  """Esta función pormite calcular la solubilidad del gas mediante uso de la correlación de Glasso,
  dicho función depende de los siguientes parámetros que deben ser ingresados por el usuario:
  yg: gravedad específica del gas
  API: grados API del oil
  T: temperatura en °F, porque la correlación mismo hace la conversión eliminándola
  p: presión en lpc ya que la función la convierte en absoluta sumando 14.7"""
  plpca = P+14.7
  varx = 2.8869 - (14.1811 - 3.3093*np.log10(plpca))**0.5
  pb = 10**varx
  rs = yg * (((API**0.989)/(T**0.172)) * (pb))**1.2255

  return rs

# %%
# Funcion para la correlación de Beal
def uo_beal(temperatura,api):
    """Función que calcula la viscosidad del petróleo
    la viscosidad se obtiene en unidades de centipoise
    Parámetros:
    api: que son los grados api del petróleo
    temperatura: es la temperatura en unidades Ranking"""
    valor_a = (10**(0.43+(8.33/api)))
    viscosidad = (0.32 + ((1.8*10**7)/(api**4.53))*(360/(temperatura-200))**valor_a)
    return viscosidad

    #print("La viscosidad es: ", viscosidad)


#%%
#Función para la correlación de Beggs-Robinson

def uo_beggs_ro(temperatura,api):
    """
    Función que calcula la viscosidad usando la correlación de Beggs-Robinson
    Parámetros:
    api: que son los grados api del petróleo
    temperatura: es la temperatura en unidades Ranking
    """
    valor_z = 3.0324 - (0.02023*api)
    valor_y = 10**valor_z
    valor_x = valor_y * ((temperatura)**(-1.163))

    viscosidad = 10**valor_x - 1

    return viscosidad
    #print("La viscosidad es: ", viscosidad)
# %%
# Funcion para la correlación de Standing para el punto de burbuja
def pbStanding(gsg,temp,API,rs):
  """Función utilizada para calcular el pb según la correlación de Standing,usando como parametros RS medido, gsg, API y la temperatura en R"""
  a=0.00091*((temp+460)-460)-0.0125*(API)
  pb=18.2*(((rs/gsg)**0.83)*((10)**a)-1.4)
  return pb
# %%
# Funcion para la correlación de Glaso para el punto de burbuja
def pbGlaso(gsg,API,temp,rs):
  """Función utilizada para calcular el pb según la correlación de Glaso
  ,usando como parametros RS medido, gsg, API y la temperatura en °F"""
  pB=(rs/gsg)**0.816*(temp**0.172)*(API**-0.989)
  lpb = np.log10(pB)
  exp=1.7669+1.7447*lpb-0.30218*(lpb**2)
  pb=10**exp
  return pb

# %%
def boAlMarhoun(yg,T,API,rs):
    """Se creó est función adicional para la obtención del factor volumétrico puesto
    que la función de Glasso presentó problemas con el logaritmo que esta incluye.
    Esta correlación utiliza las constantes a, b y c, así como también F la cual es una
    nueva variable
    La temperatura debe estar °R."""

    a = 0.742390
    b = 0.323294
    c = -1.202040
    F = (rs**a) * (yg**b) * ((141.5/(API+131.5))**c)
    bo = 0.497069 + (0.862963*(10**-3)*(T+460)) + (0.182594*(10**-2)*(F)) + 0.318099*(10**-5)*(F**2)

    return bo
# %%
def pbAlMarhoun(yg,T,API,rs):
    """Esta función también fue creada como una acción mediática hacia la correlación
    de Glasso que presentó problemas con el logaritmo.
    Pose 5 constantes (a, b, c ,d y e)
    La temperatura debe estar en °R."""

    a = 5.38088*(10**-3)
    b = 0.715082
    c = -1.87784
    d = 3.1437
    e = 1.32657
    TR = T+460
    pb = a * ((rs)**b) * ((yg)**c) * ((141.5/(API+131.5))**d) * ((TR)**e)
    return pb
