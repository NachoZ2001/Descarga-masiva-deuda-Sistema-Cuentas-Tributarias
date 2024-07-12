from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

# Lee el archivo Excel
df = pd.read_excel(r'C:\Program Files\Sublime Merge\Descarga-masiva-deuda-Sistema-de-Cuentas-Tributarias\data\input\clientes.xlsx')

# Supongamos que las columnas son 'CUIT' y 'Contraseña'
cuit_login_list = df['CUIT para ingresar'].tolist()
cuit_represent_list = df['CUIT representado'].tolist()
password_list = df['Contraseña'].tolist()
download_list = df['Ubicacion Descarga'].tolist()

# Configura el driver de Selenium    
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")

# Inicializar driver
driver = webdriver.Chrome(options=options)

# Función para iniciar sesión y extraer datos
def extraer_datos(cuit_ingresar, cuit_representado, password, ubicacion_descarga):
    driver.get('https://auth.afip.gob.ar/contribuyente_/login.xhtml')
    
    # Iniciar sesión
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'F1:username'))).send_keys(cuit_ingresar)
    driver.find_element(By.ID, 'F1:btnSiguiente').click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'F1:password'))).send_keys(password)
    driver.find_element(By.ID, 'F1:btnIngresar').click()

    time.sleep(5)
    
    # Navegar y seleccionar el CUIT
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, "Ver todos"))).click()
    time.sleep(5)
    driver.find_element(By.ID, 'buscadorInput').send_keys('SISTEMA DE CUENTAS TRIBUTARIAS')
    time.sleep(5)
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'rbt-menu-item-0'))).click()
    time.sleep(10)

    # Cambiar a la pestaña del modulo SCT
    window_handles = driver.window_handles
    driver.switch_to.window(window_handles[-1])
    
    # Seleccionar el CUIT representado del dropdown
    try:
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.NAME, "$PropertySelection"))).click()
        select_element = Select(driver.find_element(By.NAME, "$PropertySelection"))

        # Seleccionar la opción por su texto
        desired_option_text = str(cuit_representado)
        select_element.select_by_visible_text(desired_option_text)       
    except Exception as e:
        print(f"Error al seleccionar CUIT: {e}")
        return 0

    # Descargar archivo XLSX
    driver.find_element(By.XPATH, "//a[@class='dt-button buttons-excel buttons-html5' and span[text()='Excel']]").click()

    time.sleep(5)
    return 1

# Iterar sobre cada cliente
for cuit_ingresar, cuit_representado, password, download in zip(cuit_login_list, cuit_represent_list, password_list, download_list):
    faltas = extraer_datos(cuit_ingresar, cuit_representado, password, download)
    print(f'CUIT: {cuit_representado}, Faltas de presentación: {faltas}')
