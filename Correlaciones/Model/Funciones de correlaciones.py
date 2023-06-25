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