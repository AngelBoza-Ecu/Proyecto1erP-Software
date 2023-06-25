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


