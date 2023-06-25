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