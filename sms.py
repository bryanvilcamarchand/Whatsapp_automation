# Importamos las librerias necesarias

import selenium

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

import time
import pandas as pd

# Nos conectamos a Whatsapp Web usando Chromedriver. El cual debemos tener en la carpeta del proyecto.
driver_path = '/Users/bryanvimar/Documents/webscrapping/chromedriver'
driver = webdriver.Chrome(driver_path)
driver.get("https://web.whatsapp.com") 

# Si estas usando 2 pantallas, puedes añadir el siguiente código para abrir Chrome en la segunda.
driver.set_window_position(1650,0)

# Colocamos un timer de 12 segundos para esperar que cargue la página. Si tu conexion es buena podrías evaluar bajar el time a 5 o 6.

time.sleep(10)


# En un archivo debemos contener los datos telefonicos a buscar y los datos que queremos comunicar. Para este caso, son asesores de venta, y comunicaremos el detalle de sus incentivos del mes.

data = pd.read_excel('Datos.xlsx', sheet_name='liquidacion_enero')

# Colocamos un timer de 2 segundos para esperar que cargue

time.sleep(2)

# Empezamos el bucle

for i in range(len(data)):
    
# Inspeccionamos la página para ubicar la zona en la web donde se buscan los contactos.

    search = driver.find_element_by_xpath("//div[@role='textbox']")
    search.click()
    
# Ahora ingresamos el número del contacto en format texto. Añadimos un timer de 2 segundo para que cargue.

    seller = data.loc[i, 'Número'].astype(str) 
    seller1 = search.send_keys(seller)
    time.sleep(2)
    
# Damos click al contacto encontrado. De acuerdo a los numeros ingresados en la tabla Datos cargada previamente.

    selected_seller = driver.find_element_by_xpath("//span[@class='_3q9s6']")
    time.sleep(1)
    selected_seller.click()

    
# Generamos el mensaje para el contacto seleccionado. Primer definimos como variables cada una de las columnas de la tabla. Luego las añadimos al mensaje. Convertimos a string todos los tipos de datos int o date, para que puedan aparecer en el mensaje.

    name = data.loc[i,'Asesor']
    date_message = data.loc[i,'Mes Incentivo'].strftime("%d/%m/%Y") 
    month_sales = data.loc[i,'Mes']
    cumpl = data.loc[i,'Cumplimiento'].astype(str) 
    sueldo_base = data.loc[i,'Sueldo base'].astype(str) 
    incentivo = data.loc[i,'Incentivo'].astype(str) 
    total = data.loc[i,'Total'].astype(str) 

    mensaje = ("Hola " + name + ", tu cumplimiento del mes de  " + month_sales  + " fue de " + cumpl + "%. Tu sueldo total es de S/ " + total + u'\u263A' + ". \n " + "Siendo tu básico de S/ " + sueldo_base + " y comisiones de S/ " + incentivo + ". \n " + "Muchas vibras y ánimos para el próximo mes.")
            
# Ubicamos la caja para ingresar los mensajes y la seleccionamos.

    input = '//div[@title="Type a message"]'
    input_box = driver.find_element_by_xpath(input)
    time.sleep(2)
    
# Seleccionamos enviar

    input_box.send_keys(mensaje + Keys.ENTER)
    time.sleep(2) 
