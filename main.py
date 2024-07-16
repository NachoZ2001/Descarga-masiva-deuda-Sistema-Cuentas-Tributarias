from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, NamedStyle
import pandas as pd
import time
import pyautogui
import os
import glob

# Crear el archivo de resultados
resultados = []

# Recorrer los archivos en la carpeta Deudas
deudas_folder = r'C:\Program Files\Sublime Merge\Descarga-masiva-deuda-Sistema-de-Cuentas-Tributarias\data\Deudas'
for file_path in glob.glob(os.path.join(deudas_folder, '*.xlsx')):
    # Obtener el nombre del archivo sin la extensión y sin " - Deudas"
    file_name = os.path.basename(file_path).replace(' - Deudas.xlsx', '')
    cuit, cliente = file_name.split(' - ')
    
    # Leer el archivo Excel
    df = pd.read_excel(file_path)
        # Agregar la información al archivo de resultados
    df.insert(0, 'CUIT', cuit)
    df.insert(1, 'Cliente', cliente)
    df['Cantidad de faltas de presentación'] = 0

    resultados.append(df)

# Combinar todos los dataframes en uno solo
df_resultados = pd.concat(resultados, ignore_index=True)

# Guardar el dataframe combinado en un nuevo archivo Excel
resultados_file_path = r'C:\Program Files\Sublime Merge\Descarga-masiva-deuda-Sistema-de-Cuentas-Tributarias\data\output\resultados.xlsx'
df_resultados.to_excel(resultados_file_path, index=False)

print(f"Archivo de resultados guardado en: {resultados_file_path}")

