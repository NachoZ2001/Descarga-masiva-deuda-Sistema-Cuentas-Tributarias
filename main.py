from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import pyautogui

# Lee el archivo Excel
df = pd.read_excel(r'C:\Program Files\Sublime Merge\Descarga-masiva-deuda-Sistema-de-Cuentas-Tributarias\data\input\clientes.xlsx')

# Supongamos que las columnas son 'CUIT' y 'Contraseña'
cuit_login_list = df['CUIT para ingresar'].tolist()
cuit_represent_list = df['CUIT representado'].tolist()
password_list = df['Contraseña'].tolist()
download_list = df['Ubicacion Descarga'].tolist()
posterior_list = df['Posterior'].tolist()
anterior_list = df['Anterior'].tolist()

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")

# Configurar preferencias para que siempre pregunte por la ubicación de descarga
prefs = {
        "download.prompt_for_download": True,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
        }
options.add_experimental_option("prefs", prefs)

# Inicializar driver
driver = webdriver.Chrome(options=options)

driver.get('https://auth.afip.gob.ar/contribuyente_/login.xhtml')

def iniciar_sesion(cuit_ingresar, password):
    # Iniciar sesión
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'F1:username'))).send_keys(cuit_ingresar)
    driver.find_element(By.ID, 'F1:btnSiguiente').click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'F1:password'))).send_keys(password)
    driver.find_element(By.ID, 'F1:btnIngresar').click()

    time.sleep(5)

def ingresar_modulo():
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

def seleccionar_cuit_representado(cuit_representado):
    # Seleccionar el CUIT representado del dropdown
    try:
        current_selection = Select(driver.find_element(By.NAME, "$PropertySelection")).first_selected_option.text           
        if current_selection != str(cuit_representado):
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "$PropertySelection"))).click()
            select_element = Select(driver.find_element(By.NAME, "$PropertySelection"))
            select_element.select_by_visible_text(str(cuit_representado))
    except Exception as e:
        print(f"Error al seleccionar CUIT: {e}")
        return 0

def exportar_excel(ubicacion_descarga, cuit_representado):
    # Descargar archivo XLSX
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='dt-button buttons-excel buttons-html5' and span[text()='Excel']]"))).click()

    # Esperar a que aparezca el cuadro de diálogo de guardar archivo
    time.sleep(5) 

    # Simular la interacción con el cuadro de diálogo de guardar archivo usando pyautogui
    ubicacion = ubicacion_descarga
    nombre_archivo = f"{cuit_representado} - Deudas.xlsx"

    # Escribir el nombre del archivo 
    pyautogui.write(nombre_archivo)
    time.sleep(1)

    # Cambiar a ubicacion       
    pyautogui.hotkey('alt', 'd')
    time.sleep(0.5)

    # Escribir la ubicación de descarga
    pyautogui.write(ubicacion)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    
    # Activar el botón "Guardar"
    pyautogui.hotkey('alt', 't')
    time.sleep(1)

    # Confirmar el guardado
    pyautogui.press('enter')
   
    time.sleep(1)

def cerrar_sesion():
    driver.close()
    driver.find_element(By.ID, "ContenedorContribuyente").click()
    driver.find_element(By.CLASS_NAME, "media p-a-1").click()

    time.sleep(20)

# Función para iniciar sesión y extraer datos
def extraer_datos_nuevo(cuit_ingresar, cuit_representado, password, ubicacion_descarga, posterior):
    iniciar_sesion(cuit_ingresar, password)
    ingresar_modulo()
    seleccionar_cuit_representado(cuit_representado)
    exportar_excel(ubicacion_descarga, cuit_representado)

    # Obtener cantidad de faltas de presentacion
    cantidad_faltas_presentacion = driver.find_element(By.NAME, "functor$1").get_attribute('value')

    # Controlar si cerrar sesion o no en base al posterior
    if posterior == 0:
        cerrar_sesion()
    
    return cantidad_faltas_presentacion

def extraer_datos(cuit_representado, ubicacion_descarga, posterior):
    seleccionar_cuit_representado(cuit_representado)
    exportar_excel(ubicacion_descarga, cuit_representado)

    # Controlar si cerrar sesion o no en base al posterior
    if posterior == 0:
        cerrar_sesion()

# Iterar sobre cada cliente
for cuit_ingresar, cuit_representado, password, download, posterior, anterior in zip(cuit_login_list, cuit_represent_list, password_list, download_list, posterior_list, anterior_list):
    if anterior == 0:
        extraer_datos_nuevo(cuit_ingresar, cuit_representado, password, download, posterior)
    else: 
        extraer_datos(cuit_representado, download, posterior)

