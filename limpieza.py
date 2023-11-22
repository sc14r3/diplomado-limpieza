import pandas as pd
import numpy as np
from unidecode import unidecode
import re
import spell

def limpiar_detalle_dano(texto):
    # Convertir a minúsculas y quitar caracteres especiales y acentos
    texto = unidecode(str(texto).lower())

    # Separar por ';'
    partes = texto.split(';')

    # Eliminar números y puntos en cada parte
    partes = [re.sub(r'\d+|-|,|\.', '', parte) for parte in partes]

    # Corregir cada palabra en todas las partes
    partes_corregidas = [[spell.correction(palabra) for palabra in parte.split()] for parte in partes]

    # Unir las partes corregidas en cadenas
    partes_corregidas = [' '.join(parte) for parte in partes_corregidas]

    resultado = ', '.join(partes_corregidas)

    return resultado

df = pd.read_excel("CD_limpieza.xlsx")

# Renombre de columnas
df1 = df.rename(columns = {'TIPO_DA„O': 'TIPO_DANO', 'DETALLE_DA„O': 'DETALLE_DANO','MONTO_TOTAL_DE_ATENCIîN_PRE': 'MONTO_TOTAL_DE_ATENCION_PRE', 'MONTO_TOTAL_DE_ATENCIîN_CIEN': 'MONTO_TOTAL_DE_ATENCION_CIEN'})

df2 = df1.drop_duplicates(subset=['CCT'])

# Fecha de forma correcta
df2['FECHA_EVENTO'] = df2['FECHA_EVENTO'].dt.date

# Imputamos la variable MATRICULA con la media
df2['MATRICULA'].fillna(df2['MATRICULA'].mean(), inplace=True)

# Tratamiento de columna TIPO_DAÑO
tipos = df2["TIPO_DANO"].str.split(expand=True)
df2['TIPO_DANO'] = tipos[0]
df2['TIPO_DANO'].fillna(method ='ffill', inplace = True)

# Tratamiento de columna DETALLE_DANO
#df2['DETALLE_DANO'] = df2['DETALLE_DANO'].apply(lambda x: unidecode(str(x)))
df2['DETALLE_DANO'] = df2['DETALLE_DANO'].apply(limpiar_detalle_dano)

# Tratatmiendo de columna COSTO_ORIGINAL_PRE
df2['COSTO_ORIGINAL_PRE'] = df2['COSTO_ORIGINAL_PRE'].replace('no disponible', 0)
df2['COSTO_ORIGINAL_PRE'] = df2['COSTO_ORIGINAL_PRE'].astype('float64')
df2['COSTO_ORIGINAL_PRE'] = df2['COSTO_ORIGINAL_PRE'].fillna(df2['MONTO_EJERCIDO_PRE'].fillna(df2['MONTO_TOTAL_DE_ATENCION_PRE']))

# Tratamiento de columna MONTO_TOTAL_DE_ATENCION_PRE
df2['MONTO_TOTAL_DE_ATENCION_PRE'] = df2['MONTO_TOTAL_DE_ATENCION_PRE'].replace('no disponible', 0)
df2['MONTO_TOTAL_DE_ATENCION_PRE'] = df2['MONTO_TOTAL_DE_ATENCION_PRE'].astype('float64')
df2['MONTO_TOTAL_DE_ATENCION_PRE'] = df2['MONTO_TOTAL_DE_ATENCION_PRE'].fillna(df2['COSTO_ORIGINAL_PRE'].fillna(df2['MONTO_EJERCIDO_PRE']))

# Tratamiendo de columna MONTO_EJERCIDO_PRE
df2['MONTO_EJERCIDO_PRE'] = df2['MONTO_EJERCIDO_PRE'].replace('no disponible', 0)
df2['MONTO_EJERCIDO_PRE'] = df2['MONTO_EJERCIDO_PRE'].astype('float64')
df2['MONTO_EJERCIDO_PRE'] = df2['MONTO_EJERCIDO_PRE'].fillna(df2['COSTO_ORIGINAL_PRE'].fillna(df2['MONTO_TOTAL_DE_ATENCION_PRE']))

# Tratamiento de columna MONTO_TOTAL_DE_ATENCION_CIEN
df2['MONTO_TOTAL_DE_ATENCION_CIEN'] = df2['MONTO_TOTAL_DE_ATENCION_CIEN'].replace(-1, 0)

df2.to_excel('CD_limpios.xlsx')
