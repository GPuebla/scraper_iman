import pandas as pd
from openpyxl import load_workbook
import re

def extraer_dato(patron, texto):
    match = re.search(patron, texto, re.DOTALL)
    return match.group(1).strip() if match else None

# Cargar datos
df = pd.read_excel("datos.xlsx")

# Nuevas columnas
df["Tela"] = df["columna_original"].apply(lambda x: extraer_dato(r'Tela:\s*(.+?)(?=Colección|Descripción|$)', x))
df["Colección"] = df["columna_original"].apply(lambda x: extraer_dato(r'Colección\s*(.+?)(?=\n|$)', x))
df["Largo por Talle"] = df["columna_original"].apply(lambda x: extraer_dato(r'LARGO POR TALLE.*?(?=TALLE|$)', x))
df["Tabla TALLE"] = df["columna_original"].apply(lambda x: extraer_dato(r'TALLE\s?\d{2}-[A-Z]+.*?(?=Contorno|$)', x))
df["Contorno busto"] = df["columna_original"].apply(lambda x: extraer_dato(r'Contorno de busto\s.*', x))
df["Tiro alto"] = df["columna_original"].apply(lambda x: extraer_dato(r'Tiro alto.*', x))
df["Tiro medio"] = df["columna_original"].apply(lambda x: extraer_dato(r'Tiro medio.*', x))
df["Tiro bajo"] = df["columna_original"].apply(lambda x: extraer_dato(r'Tiro bajo.*', x))
df["Contorno cadera"] = df["columna_original"].apply(lambda x: extraer_dato(r'Contorno de cadera.*', x))
df["Altura"] = df["columna_original"].apply(lambda x: extraer_dato(r'Altura.*', x))

# Nombre del archivo Excel (puede ser nuevo o existente)
archivo_excel = "datos.xlsx"
nombre_hoja = "Procesados"

try:
    # Intentar abrir el archivo si ya existe
    writer = pd.ExcelWriter(archivo_excel, engine='openpyxl', mode='a', if_sheet_exists='replace')
except FileNotFoundError:
    # Si no existe, se crea uno nuevo
    writer = pd.ExcelWriter(archivo_excel, engine='openpyxl', mode='w')

# Guardar en la hoja especificada
df.to_excel(writer, sheet_name=nombre_hoja, index=False)

# Guardar el archivo
writer.close()