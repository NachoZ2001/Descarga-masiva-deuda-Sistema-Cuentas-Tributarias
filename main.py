from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Lee el archivo Excel
df = pd.read_excel('C:\Program Files\Sublime Merge\Descarga-masiva-deuda-Sistema-de-Cuentas-Tributarias\data\input\clientes.xlsx')

# Supongamos que las columnas son 'CUIT' y 'Contraseña'
cuit_login_list = df['CUIT para ingresar'].tolist()
cuit_represent_list = df['CUIT representado'].tolist()
password_list = df['Contraseña'].tolist()

# Configura el driver de Selenium    
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

# Función para iniciar sesión y extraer datos
def extraer_datos(cuit_ingresar, cuit_representado, password):

    driver.get('https://auth.afip.gob.ar/contribuyente_/login.xhtml')
    
    # Encuentra los campos de CUIT y contraseña e inicia sesión
    driver.find_element(By.ID, 'F1:username').send_keys(cuit_ingresar)
    driver.find_element(By.ID, 'F1:btnSiguiente').click()
    driver.find_element(By.ID, 'F1:password').send_keys(password)
    driver.find_element(By.ID, 'F1:btnIngresar').click()

    # Agrega un tiempo de espera para cargar la página
    time.sleep(5)
    
    # Espera a que el botón "Ver todos" sea visible y haz clic en él
    wait = WebDriverWait(driver, 10)
    ver_todos_button = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Ver todos")))
    ver_todos_button.click()
    
    # Agrega un tiempo de espera para cargar la página
    time.sleep(5)

    # Escribir en el buscador el nombre del modulo
    driver.find_element(By.ID, 'buscadorInput').send_keys('SISTEMA DE CUENTAS TRIBUTARIAS')

    # Agrega un tiempo de espera para cargar el modulo
    time.sleep(5)

    # Hacer click en modulo desplegado
    wait = WebDriverWait(driver, 10)
    modulo_button = wait.until(EC.element_to_be_clickable((By.ID, 'rbt-menu-item-0')))
    modulo_button.click()

    time.sleep(25)
    
    # Hacer clic en el select para desplegar las opciones
    select_element = driver.find_element(By.NAME, '$PropertySelection')
    select_element.click()

    return 1

# Itera sobre cada cliente
for cuit_ingresar, cuit_representado, password in zip(cuit_login_list, cuit_represent_list,password_list):
    faltas = extraer_datos(cuit_ingresar, cuit_representado, password)
    print(f'CUIT: {cuit_representado}, Faltas de presentación: {faltas}')

