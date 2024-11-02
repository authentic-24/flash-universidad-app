import time
from flask import Flask, render_template, request, jsonify
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from fuzzywuzzy import process, fuzz
import logging
import os

app = Flask(__name__)

# Configuración del logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
handler = logging.StreamHandler()
handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

excel_path = './ofertas_de_empleo.xlsx'  # Reemplaza con la ruta a tu archivo de Excel
df = pd.read_excel(excel_path)

# Leer el archivo Excel en un DataFrame al inicio
df = pd.read_excel(excel_path, engine='openpyxl')

logging.basicConfig(level=logging.INFO)  # Configurar el nivel de registro
logger = logging.getLogger(__name__)  # Crear un logger específico para tu aplicación

universidades = [
    {
        'nombre': 'Areandina',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/fundacion-universitaria-area-andina/Home.aspx?ekp=4ylL1XxasTfttAUBYbQ2O8Q53xyOoS9yOGA4sBzeaMoGTMRtQ9JhAEv3Z3lWrzc5gDV2smXs/PVcbucdgGmsOiGO66RslP0fEIA/WDN3NFJAdEVRavBug38qTuAtrz2Jnrjzi/NovzGYOPNLWj2nI3e17ItOgCZkSFwi0/lCXZsUDROE8lDL94JuX4DH/2/EJ9djiMf/FoHjJrqf7GccjssjmRcfxKKbrOdNmHtmcD30y1MI2tlGYCl3SDXs3CqRFYroNjvjmOXpp1TBim4KVOBwF3gjPu54VelnxEceB7ddRPsqRBe1H0NL6ipawIbv2a12OkbG4o4Wc3imIT9jO8Qd/BtcW+LLJz7GEB2+jh0DoUvjL9YOR8gum7EfkS5mCsLlT4j4jcJtCpIgpoyBLg==',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'U catolica',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-catolica-oriente/Home.aspx?ekp=4ylL1XxasTfttAUBYbQ2O8Q53xyOoS9yOGA4sBzeaMoGTMRtQ9JhAEv3Z3lWrzc5gDV2smXs/PX5nG1LskBC5PvagFAdsEwhlJc1icLyDM+IKFN8iH1S1hhAwZIfdf05Wke5wCGYC1pqLOJ+1ooHFIFoO2uP56VLR13sZy/jWuCPFpq1F0+tcg==',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'Simonbolivar',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-simon-bolivar/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'Antonio Nariño',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/uan/Home.aspx?ekp=4ylL1XxasTfttAUBYbQ2O8Q53xyOoS9yOGA4sBzeaMoGTMRtQ9JhAEv3Z3lWrzc5gDV2smXs/PX5nG1LskBC5PvagFAdsEwhlJc1icLyDM+IKFN8iH1S1kyIkrHtI8WwWCN4cH45ovHBtwkpwIGF7w==',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'cooperativa',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-coperativa-colombia/Home.aspx?ekp=4ylL1XxasTfttAUBYbQ2O8Q53xyOoS9yOGA4sBzeaMoGTMRtQ9JhAEv3Z3lWrzc5gDV2smXs/PX5nG1LskBC5PvagFAdsEwhlJc1icLyDM+IKFN8iH1S1hhAwZIfdf05Yb7PobaEWuUuaV07YmoXXpMNaXFYRw7gOk5w0QAXA507kyijiig6s/uLYzKzhp0b',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'santiago',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/usc/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'San buenaventura',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-san-buenaventura-cali/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'icesi',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/icesi/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'Elbosque',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-bosque/Home.aspx?ekp=Q9XEO1kGCJWxeptS7fVbM1wwr5cmn//G5geoKwcQIPvp3R/lt+w6WaAy+rsStNYeTBIgtabW6VpjP8yzHBalbGOlAM/TorZ1jhEYgH9W5nQiLECX8JMIvMTOHfpbUt1odGX/bl6tddgEaUPbel+983JN5xARsd4R7iRP6a7KHGPve/kuLFWV3VNpb3fndPh81kCyrIhVXCeTkZU4OL/xeMeT81ojVos+',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'Uao',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/UAO/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'unisangil',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/corporacion-universitaria-san-gil/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'unad',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-nacional-abierta-distancia/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'unisanitas',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/fundacion-universitaria-sanitas/Home.aspx?ekp=4ylL1XxasTfttAUBYbQ2O8Q53xyOoS9yOGA4sBzeaMoGTMRtQ9JhAEv3Z3lWrzc5gDV2smXs/PX5nG1LskBC5PvagFAdsEwhlJc1icLyDM+IKFN8iH1S1n62Og4lOdEIWDoMdx4W5Jl+d5otbKwgcr7/Rj5MsNznnjQc4d9CgwZ4SlP4dB+gSC+dGjxZ30YU',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'lasalle',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-salle/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'militar',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/u_militar/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'catolica',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-catolica-pereira/Home.aspx?ekp=4ylL1XxasTfttAUBYbQ2O8Q53xyOoS9yOGA4sBzeaMoGTMRtQ9JhAEv3Z3lWrzc5gDV2smXs/PX5nG1LskBC5PvagFAdsEwhlJc1icLyDM+IKFN8iH1S1hhAwZIfdf05Wke5wCGYC1px+aq9aWFlkcHlFEwpv4j4vyNe+WfGNlr9whOL5BniyQ==x',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'Antioquia',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-antioquia/Home.aspx?ekp=4ylL1XxasTfttAUBYbQ2O8Q53xyOoS9yOGA4sBzeaMoGTMRtQ9JhAEv3Z3lWrzc5gDV2smXs/PX5nG1LskBC5PvagFAdsEwhlJc1icLyDM+IKFN8iH1S1hhAwZIfdf05+iN5EARiD9/CMOPqmqDU8yR1TvcG9JitgIrKdGOK5vJr+H/TfuOpxg==',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'Piloto',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/universidad-piloto/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    {
        'nombre': 'santander',
        'url': 'https://www.elempleo.com/colombia/Files/BasesUniversitarias/udes/Home.aspx',
        'username': 'desarrollador-junior@yo-soy.co',
        'password': '900299250'
    },
    
]
# Función para guardar el archivo Excel actualizado
def guardar_excel_actualizado():
    global df  # Acceder al DataFrame global
    df.to_excel(excel_path, index=False, engine='openpyxl')
    
    # Abrir el archivo Excel después de guardarlo
    try:
        os.system(f'start excel "{excel_path}"')
        print(f"Archivo Excel '{excel_path}' abierto correctamente.")
    except Exception as e:
        print(f"No se pudo abrir el archivo Excel: {e}")


# Función para cargar datos desde el archivo Excel
def cargar_datos_excel(ruta):
    try:
        return pd.read_excel(ruta)
    except Exception as e:
        logger.error(f"Error al cargar el archivo Excel: {str(e)}")
        return None

# Función para inicializar el driver de Selenium
def iniciar_driver():
    try:
        return webdriver.Chrome()
    except Exception as e:
        logger.error(f"Error al iniciar el driver de Selenium: {str(e)}")
        return None

# Función para procesar cada universidad
def procesar_universidad(nombre_universidad, oferta, fila, driver):
    
    # Buscar la configuración de la universidad por su nombre
    universidad = next((u for u in universidades if u['nombre'] == nombre_universidad), None)
    
    if universidad is None:
        logger.error(f"No se encontró la universidad '{nombre_universidad}' en la configuración.")
        return f"No se encontró la universidad '{nombre_universidad}' en la configuración."
    url = universidad['url']
    username = universidad['username']
    password = universidad['password']
    try:
        driver = webdriver.Chrome()
        driver.get(url)

        # Realiza el inicio de sesión si es necesario
        if username and password:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ctl00_Header1_UniversityCompaniesControl_headerLoginControl_txtUserName'))
            ).send_keys(username)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ctl00_Header1_UniversityCompaniesControl_headerLoginControl_txtPassword'))
            ).send_keys(password)
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'ctl00_Header1_UniversityCompaniesControl_headerLoginControl_btnLogin'))
            ).click()

    

        for index, row in df.iterrows():
                
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_content_PublishJobOfferButton"))
                ).click()
                # Completa el título
                campo1 = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'ctl00_content_ctl73_title'))
                )
                campo1.clear()
                campo1.send_keys(row.iloc[0])

                # Define la palabra clave y otros parámetros iniciales
                palabra_clave = row.iloc[1].strip().lower()
                umbral_similitud = 80
                intentos = 0
                opcion_seleccionada = None

                # Continúa intentando mientras no se encuentre la opción y los intentos sean menores a 5
                while opcion_seleccionada is None and intentos < 5:
                    try:
                        # Espera a que los contenedores de opciones estén disponibles
                        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'checkboxlist-container')))
                        contenedores = driver.find_elements(By.CLASS_NAME, 'checkboxlist-container')

                        for contenedor in contenedores:
                            opciones_input = contenedor.find_elements(By.XPATH, './/input[@type="checkbox"]')
                            for input_checkbox in opciones_input:
                                texto_opcion = input_checkbox.find_element(By.XPATH, './following-sibling::label').text.strip().lower()
                                similitud = fuzz.partial_ratio(palabra_clave, texto_opcion)
                                if similitud >= umbral_similitud:
                                    input_checkbox.click()
                                    opcion_seleccionada = texto_opcion
                                    print(f"Opción seleccionada: {opcion_seleccionada} para la palabra clave {palabra_clave}")
                                    break
                            if opcion_seleccionada:
                                break
                    except Exception as e:
                        print(f"Error al buscar opciones: {e}")

                    intentos += 1
                    if opcion_seleccionada is None:
                        # Espera antes de intentar nuevamente

                
                        print(f"No se encontró una opción adecuada para '{palabra_clave}' tras {intentos} intentos.")
                # Verifica si el valor de la columna C-2 es 'sí' utilizando iloc para acceder a la posición 2

                palabra_clave = row.iloc[2].strip().lower()
                umbral_similitud = 80
                intentos = 0
                opcion_seleccionada = None

                while opcion_seleccionada is None and intentos < 5:
                    try:
                        # Espera a que los contenedores de opciones estén disponibles
                        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'checkboxlist-container')))
                        contenedores = driver.find_elements(By.CLASS_NAME, 'checkboxlist-container')

                        for contenedor in contenedores:
                            opciones_input = contenedor.find_elements(By.XPATH, './/input[@type="checkbox"]')
                            for input_checkbox in opciones_input:
                                texto_opcion = input_checkbox.find_element(By.XPATH, './following-sibling::label').text.strip().lower()
                                similitud = fuzz.partial_ratio(palabra_clave, texto_opcion)
                                if similitud >= umbral_similitud:
                                    input_checkbox.click()
                                    opcion_seleccionada = texto_opcion
                                    print(f"Opción seleccionada: {opcion_seleccionada} para la palabra clave {palabra_clave}")
                                    break
                            if opcion_seleccionada:
                                break
                    except Exception as e:
                        print(f"Error al buscar opciones: {e}")

                    intentos += 1
                    if opcion_seleccionada is None:
                        

            
                        print(f"No se encontró una opción adecuada para '{palabra_clave}' tras {intentos} intentos.")

                palabra_clave_select = row.iloc[3].strip()
                print(f"Palabra clave del select obtenida de la hoja de cálculo: '{palabra_clave_select}'")
                select = driver.find_element(By.ID, 'ctl00_content_ctl73_positionLevel')
                select.click()  
                
                opciones_select = select.find_elements(By.TAG_NAME, 'option')
                for opcion in opciones_select:
                    opcion_texto = opcion.text.strip()
                    print(f"Valor de la opción encontrada: '{opcion_texto}'")
                    if opcion_texto == palabra_clave_select:
                            opcion.click()  # Seleccionar la opción
                            print(f"Valor del select '{palabra_clave_select}' seleccionado.")
                            break
                time.sleep(3)     
                palabra_clave_select = row.iloc[4].strip()
                print(f"Palabra clave del select obtenida de la hoja de cálculo: '{palabra_clave_select}'")
                select = driver.find_element(By.ID, 'ctl00_content_ctl73_positionSubLevel')
                select.click()  
                
                opciones_select = select.find_elements(By.TAG_NAME, 'option')
                for opcion in opciones_select:
                    opcion_texto = opcion.text.strip()
                    print(f"Valor de la opción encontrada: '{opcion_texto}'")
                    if opcion_texto == palabra_clave_select:
                        opcion.click()  # Seleccionar la opción
                        print(f"Valor del select '{palabra_clave_select}' seleccionado.")
                        break
                    
                
                
                
                # Completa el título
                campo2 = driver.find_element(By.ID, 'ctl00_content_ctl73_field_JobOffer_AdditionalInfo_VacancyQuantity_box')
                campo2.clear()
                campo2.send_keys(row.iloc[5])

                palabra_clave_select = row.iloc[6].strip()
                print(f"Palabra clave del select obtenida de la hoja de cálculo: '{palabra_clave_select}'")
                select = driver.find_element(By.ID, 'ctl00_content_ctl73_field_JobOffer_Salary_SalaryInfo_box')
                select.click()
                                
                opciones_select = select.find_elements(By.TAG_NAME, 'option')
                for opcion in opciones_select:
                    opcion_texto = opcion.text.strip()
                    print(f"Valor de la opción encontrada: '{opcion_texto}'")
                    if opcion_texto == palabra_clave_select:
                        opcion.click()  # Seleccionar la opción
                        print(f"Valor del select '{palabra_clave_select}' seleccionado.")




                for i, field_id in enumerate(['ctl00_content_ctl74_descriptionBox_text', 'ctl00_content_ctl74_requirements_text', 'ctl00_content_ctl78_fromExperienceYears', 'ctl00_content_ctl78_toExperienceYeras'], start=7):
                    campo = driver.find_element(By.ID, field_id)
                    campo.clear()
                    campo.send_keys(row.iloc[i])

                # Verifica las opciones de radio y checkbox
                for i, field_id in enumerate(['ctl00_content_ctl77_academyState_2', 'ctl00_content_ctl73_PublishCompanyNameNo'], start=9):
                    palabra_clave = row.iloc[i].strip().lower()
                    # umbral_similitud = 80
                    # intentos = 0
                    # opcion_seleccionada = None

                palabra_clave = str(row.iloc[9]).lower()
                umbral_similitud = 80
                intentos = 0
                opcion_seleccionada = None

                while opcion_seleccionada is None and intentos < 5:
                    try:
                        # Espera a que los contenedores de opciones estén disponibles
                        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'checkboxlist-container')))
                        contenedores = driver.find_elements(By.CLASS_NAME, 'checkboxlist-container')

                        for contenedor in contenedores:
                            # Busca los spans que contienen los inputs de tipo radio
                            opciones_radio = contenedor.find_elements(By.XPATH, ".//span[contains(@id, 'ctl')]/input[@type='radio']")
                            for input_radio in opciones_radio:
                                # Puedes obtener el texto del label asociado si es necesario
                                label = input_radio.find_element(By.XPATH, './following-sibling::label')
                                texto_opcion = label.text.strip().lower()
                                similitud = fuzz.partial_ratio(palabra_clave, texto_opcion)
                                if similitud >= umbral_similitud:
                                    input_radio.click()  # Hacemos clic en el radio button
                                    opcion_seleccionada = texto_opcion
                                    print(f"Opción seleccionada: {opcion_seleccionada} para la palabra clave {palabra_clave}")
                                    break
                            if opcion_seleccionada:
                                break
                    except Exception as e:
                        print(f"Error al buscar opciones: {e}")

                    intentos += 1
                    if opcion_seleccionada is None:
                    

            
                        print(f"No se encontró una opción adecuada para '{palabra_clave}' tras {intentos} intentos.")



                palabra_clave = row.iloc[10].strip().lower()
                umbral_similitud = 80
                intentos = 0
                opcion_seleccionada = None

                while opcion_seleccionada is None and intentos < 5:
                    try:
                        # Espera a que los contenedores de opciones estén disponibles
                        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'checkboxlist-container')))
                        contenedores = driver.find_elements(By.CLASS_NAME, 'checkboxlist-container')

                        for contenedor in contenedores:
                            opciones_input = contenedor.find_elements(By.XPATH, './/input[@type="checkbox"]')
                            for input_checkbox in opciones_input:
                                texto_opcion = input_checkbox.find_element(By.XPATH, './following-sibling::label').text.strip().lower()
                                similitud = fuzz.partial_ratio(palabra_clave, texto_opcion)
                                if similitud >= umbral_similitud:
                                    input_checkbox.click()
                                    opcion_seleccionada = texto_opcion
                                    print(f"Opción seleccionada: {opcion_seleccionada} para la palabra clave {palabra_clave}")
                                    break
                            if opcion_seleccionada:
                                break
                    except Exception as e:
                        print(f"Error al buscar opciones: {e}")

                    intentos += 1
                    if opcion_seleccionada is None:
                    

            
                        print(f"No se encontró una opción adecuada para '{palabra_clave}' tras {intentos} intentos.")


                palabra_clave = str(row.iloc[11]).lower()

                umbral_similitud = 80
                intentos = 0
                opcion_seleccionada = None

                while opcion_seleccionada is None and intentos < 5:
                    try:
                        # Espera a que los contenedores de opciones estén disponibles
                        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'checkboxlist-container')))
                        contenedores = driver.find_elements(By.CLASS_NAME, 'checkboxlist-container')

                        for contenedor in contenedores:
                            # Busca los spans que contienen los inputs de tipo radio
                            opciones_radio = contenedor.find_elements(By.XPATH, ".//span[contains(@id, 'ctl')]/input[@type='radio']")
                            for input_radio in opciones_radio:
                                # Puedes obtener el texto del label asociado si es necesario
                                label = input_radio.find_element(By.XPATH, './following-sibling::label')
                                texto_opcion = label.text.strip().lower()
                                similitud = fuzz.partial_ratio(palabra_clave, texto_opcion)
                                if similitud >= umbral_similitud:
                                    input_radio.click()  # Hacemos clic en el radio button
                                    opcion_seleccionada = texto_opcion
                                    print(f"Opción seleccionada: {opcion_seleccionada} para la palabra clave {palabra_clave}")
                                    break
                            if opcion_seleccionada:
                                break
                    except Exception as e:
                        print(f"Error al buscar opciones: {e}")

                    intentos += 1
                    if opcion_seleccionada is None:
                        

                        print(f"No se encontró una opción adecuada para '{palabra_clave}' tras {intentos} intentos.")


                palabra_clave = row.iloc[12].strip().lower()
                umbral_similitud = 80
                intentos = 0
                opcion_seleccionada = None

                while opcion_seleccionada is None and intentos < 5:
                        try:
                            # Espera a que los contenedores de opciones estén disponibles
                            WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.ID, 'ctl00_content_ctl76_sectors_childList')))
                            contenedores = driver.find_elements(By.ID, 'ctl00_content_ctl76_sectors_childList')

                            for contenedor in contenedores:
                                opciones_input = contenedor.find_elements(By.XPATH, './/input[@type="checkbox"]')
                                for input_checkbox in opciones_input:
                                    texto_opcion = input_checkbox.find_element(By.XPATH, './following-sibling::label').text.strip().lower()
                                    similitud = fuzz.partial_ratio(palabra_clave, texto_opcion)
                                    if similitud >= umbral_similitud:
                                        input_checkbox.click()
                                        opcion_seleccionada = texto_opcion
                                        print(f"Opción seleccionada: {opcion_seleccionada} para la palabra clave {palabra_clave}")
                                        break
                                if opcion_seleccionada:
                                    break
                        except Exception as e:
                            print(f"Error al buscar opciones: {e}")

                        intentos += 1
                        if opcion_seleccionada is None:
                            

            
                            print(f"No se encontró una opción adecuada para '{palabra_clave}' tras {intentos} intentos.")

                palabra_clave = str(row.iloc[13]).lower()

                umbral_similitud = 80
                intentos = 0
                opcion_seleccionada = None

                while opcion_seleccionada is None and intentos < 5:
                    try:
                        # Espera a que los contenedores de opciones estén disponibles
                        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'checkboxlist-container')))
                        contenedores = driver.find_elements(By.CLASS_NAME, 'checkboxlist-container')

                        for contenedor in contenedores:
                            # Busca los spans que contienen los inputs de tipo radio
                            opciones_radio = contenedor.find_elements(By.XPATH, ".//span[contains(@id, 'ctl')]/input[@type='radio']")
                            for input_radio in opciones_radio:
                                # Puedes obtener el texto del label asociado si es necesario
                                label = input_radio.find_element(By.XPATH, './following-sibling::label')
                                texto_opcion = label.text.strip().lower()
                                similitud = fuzz.partial_ratio(palabra_clave, texto_opcion)
                                if similitud >= umbral_similitud:
                                    input_radio.click()  # Hacemos clic en el radio button
                                    opcion_seleccionada = texto_opcion
                                    print(f"Opción seleccionada: {opcion_seleccionada} para la palabra clave {palabra_clave}")
                                    break
                            if opcion_seleccionada:
                                break
                    except Exception as e:
                        print(f"Error al buscar opciones: {e}")

                    intentos += 1
                    if opcion_seleccionada is None:
                    

                
                        print(f"No se encontró una opción adecuada para '{palabra_clave}' tras {intentos} intentos.")



                palabra_clave = row.iloc[14].strip().lower()
                umbral_similitud = 100
                intentos = 0
                opcion_seleccionada = None

                while opcion_seleccionada is None and intentos < 5:
                        try:
                            # Espera a que los contenedores de opciones estén disponibles
                            WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.ID, 'ctl00_content_ctl77_checkBoxList')))
                            contenedores = driver.find_elements(By.ID, 'ctl00_content_ctl77_checkBoxList')

                            for contenedor in contenedores:
                                opciones_input = contenedor.find_elements(By.XPATH, './/input[@type="checkbox"]')
                                for input_checkbox in opciones_input:
                                    texto_opcion = input_checkbox.find_element(By.XPATH, './following-sibling::label').text.strip().lower()
                                    similitud = fuzz.partial_ratio(palabra_clave, texto_opcion)
                                    if similitud >= umbral_similitud:
                                        input_checkbox.click()
                                        opcion_seleccionada = texto_opcion
                                        print(f"Opción seleccionada: {opcion_seleccionada} para la palabra clave {palabra_clave}")
                                        break
                                if opcion_seleccionada:
                                    break
                        except Exception as e:
                            print(f"Error al buscar opciones: {e}")

                        intentos += 1
                        if opcion_seleccionada is None:
                            

                
                            print(f"No se encontró una opción adecuada para '{palabra_clave}' tras {intentos} intentos.")



                campo_from = driver.find_element(By.ID, 'ctl00_content_ctl78_fromExperienceYears')
                campo_from.clear()
                campo_from.send_keys(row.iloc[15])

                # Completa el campo 'To Experience Years'
                campo_to = driver.find_element(By.ID, 'ctl00_content_ctl78_toExperienceYeras')  # Corregido 'Years' a 'Yeras'
                campo_to.clear()
                campo_to.send_keys(row.iloc[16])

                palabra_clave_select = str(row.iloc[17]).strip()
                print(f"Tipo de dato: {type(row.iloc[17])}, Valor: {row.iloc[17]}")
                select = driver.find_element(By.ID, 'ctl00_content_ctl78_field_JobOffer_AdditionalInfo_RequiredExperience_box')
                select.click()  
                
                opciones_select = select.find_elements(By.TAG_NAME, 'option')
                for opcion in opciones_select:
                    opcion_texto = opcion.text.strip()
                    print(f"Valor de la opción encontrada: '{opcion_texto}'")
                    if opcion_texto == palabra_clave_select:
                        opcion.click()  # Seleccionar la opción
                        print(f"Valor del select '{palabra_clave_select}' seleccionado.")

                palabra_clave_select = row.iloc[18].strip()
                print(f"Palabra clave del select obtenida de la hoja de cálculo: '{palabra_clave_select}'")
                select = driver.find_element(By.ID, 'ctl00_content_ctl78_field_JobOffer_AdditionalInfo_ContractType_box')
                select.click()  
                
                opciones_select = select.find_elements(By.TAG_NAME, 'option')
                for opcion in opciones_select:
                    opcion_texto = opcion.text.strip()
                    print(f"Valor de la opción encontrada: '{opcion_texto}'")
                    if opcion_texto == palabra_clave_select:
                        opcion.click()  # Seleccionar la opción
                        print(f"Valor del select '{palabra_clave_select}' seleccionado.")

                palabra_clave_select = row.iloc[19].strip()
                print(f"Palabra clave del select obtenida de la hoja de cálculo: '{palabra_clave_select}'")
                select = driver.find_element(By.ID, 'ctl00_content_ctl78_field_JobOffer_AdditionalInfo_DedicatedTime_box')
                select.click()  
                
                opciones_select = select.find_elements(By.TAG_NAME, 'option')
                for opcion in opciones_select:
                    opcion_texto = opcion.text.strip()
                    print(f"Valor de la opción encontrada: '{opcion_texto}'")
                    if opcion_texto == palabra_clave_select:
                        opcion.click()  # Seleccionar la opción
                        print(f"Valor del select '{palabra_clave_select}' seleccionado.")

        # fi    la_actual += 1
                

                boton = driver.find_element(By.ID, 'ctl00_content_ctl73_PublishCompanyNameNo')  # Reemplazar 'id_del_boton' con el ID real

                # Hacer clic en el botón
                boton.click()

                boton = driver.find_element(By.ID, 'ctl00_content_ctl73_field_JobOffer_Salary_PublishSalary_box_list_1')  # Reemplazar 'id_del_boton' con el ID real

                # Hacer clic en el botón
                boton.click() 

                boton = driver.find_element(By.ID, 'ctl00_content_ctl77_academyState_2')  # Reemplazar 'id_del_boton' con el ID real

                # Hacer clic en el botón
                boton.click()

                boton = driver.find_element(By.ID, 'ctl00_content_Save')  # Boton de guardado

                # Hacer clic en el botón
                boton.click()

                boton = driver.find_element(By.ID, 'ctl00_content_Save_dialogContainerok_button')  # Boton que acepta los cambios 

            # Hacer clic en el botón
                boton.click()

                

        driver.quit()

        return 'Processed successfully'
    except Exception as e:
        logger.error(f"Error procesando universidad en {url}: {str(e)}")
        return f"Error: {str(e)}"
    finally:   
        driver.quit()
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        data = request.get_json()
        logger.debug(f"Datos recibidos: {data}")

        nombre_universidad = data.get('universidad')
        logger.debug(f"Nombre de universidad recibido: {nombre_universidad}")

        if not nombre_universidad:
            return jsonify({'message': 'Nombre de universidad no proporcionado.', 'status': 'error'})

        # Cargar el archivo Excel usando la ruta definida globalmente
        df = cargar_datos_excel(excel_path)
        if df is None:
            return jsonify({"status": "Error al cargar el archivo Excel"})

        # Filtrar las filas correspondientes a la universidad recibida
        filas_universidad = df[df['universidades'].str.contains(nombre_universidad, case=False, na=False)]

        if filas_universidad.empty:
            return jsonify({"status": f"No se encontraron ofertas para la universidad {nombre_universidad}"}), 404

        # Inicializar el driver de Selenium
        driver = iniciar_driver()
        if driver is None:
            return jsonify({"status": "Error al iniciar el driver de Selenium"})

        resultados = []

        for index, fila in filas_universidad.iterrows():
            # Construir la oferta para procesar
            oferta = {
                'titulo': fila['título de la oferta'],  # Ajustar nombres de columnas según tu archivo Excel
                'area': fila['área'],
                'cargo_equivalente': fila['cargo equivalente'],
                'nivel_educativo': fila['nivel educativo'],
                'subnivel': fila['subnivel'],
                'cantidad_vacantes': fila['cantidad de vacantes'],
                'rango_salario': fila['rango del salario en millones'],
                'descripcion': fila['descripción'],
                'requisitos': fila['requisitos'],
                'departamento': fila['departamento'],
                'ciudad': fila['ciudad'],
                'sector': fila['sector'],
                'subsectores': fila['subsectores'],
                'nivel_estudio': fila['nivel de estudio'],
                'profesion': fila['profecion'],
                'experiencia_minima': fila['años totales de experiencia entre'],
                'experiencia_maxima': fila['años totales de experiencia'],
                'experiencia_requerida': fila['Experiencia requerida'],
                'tipo_contrato': fila['tipo de contrato'],
                'tiempo_dedicado': fila['tiempo dedicado']
            }
            # Procesar la universidad con la oferta construida
            resultado = procesar_universidad(nombre_universidad, oferta, fila, driver)
            resultados.append({'nombre': nombre_universidad, 'resultado': resultado})

        # Cerrar el driver de Selenium al finalizar
        driver.quit()

        logger.debug(f"Resultado del procesamiento: {resultados}")

        return jsonify({
            'message': f"Procesando universidad {nombre_universidad} correctamente.",
            'status': 'success',
            'resultados': resultados
        })
    except Exception as e:
        logger.error(f"Error en la solicitud de procesamiento: {e}")
        return jsonify({'message': 'Error interno del servidor.', 'status': 'error'}), 500


@app.route('/actualizar_excel', methods=['POST'])
def actualizar_excel():
    try:
        data = request.get_json()
        logger.debug(f"Datos recibidos para actualizar Excel: {data}")

        actualizar = data.get('actualizar')

        if actualizar:
            guardar_excel_actualizado()
            return jsonify({'message': 'Archivo Excel actualizado correctamente.', 'status': 'success'})
        else:
            return jsonify({'message': 'No se especificó la acción para actualizar el archivo Excel.', 'status': 'error'})
    except Exception as e:
        logger.error(f"Error al actualizar el archivo Excel: {e}")
        return jsonify({'message': 'Error interno del servidor.', 'status': 'error'}), 500

def guardar_excel_actualizado():
    try:
        # Lee el archivo Excel existente
        logger.debug(f"Abriendo el archivo Excel desde {excel_path}")
        df = pd.read_excel(excel_path)
        logger.debug("Archivo Excel leído correctamente.")

        # Realiza las actualizaciones necesarias en el DataFrame
        # Aquí puedes agregar la lógica para actualizar el DataFrame según sea necesario
        # Ejemplo: df['nueva_columna'] = 'valor'

        # Guarda el DataFrame actualizado de vuelta al archivo Excel
        logger.debug("Guardando el archivo Excel actualizado.")
        df.to_excel(excel_path, index=False)
        logger.debug("Archivo Excel guardado correctamente.")
    except Exception as e:
        logger.error(f"Error al guardar el archivo Excel actualizado: {e}")
        raise e


@app.route('/abrir_excel', methods=['POST'])
def abrir_excel():
    try:
        # Verifica que se haya enviado un archivo
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'No se ha recibido ningún archivo.'})

        file = request.files['file']

        # Verifica que el archivo tenga una extensión válida
        if file.filename == '' or not file.filename.endswith('.xlsx'):
            return jsonify({'status': 'error', 'message': 'El archivo no es válido. Debe ser un archivo .xlsx.'})

        # Guardar el archivo en una ruta temporal
        excel_path = './temp_' + file.filename
        file.save(excel_path)

        # Abre el archivo Excel
        os.system(f'start EXCEL.EXE "{os.path.abspath(excel_path)}"')

        return jsonify({'status': 'success', 'message': 'Archivo Excel abierto correctamente.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

    
@app.route('/automatizar_universidades', methods=['POST'])
def automatizar_universidades():
    try:
        # Cargar el archivo Excel
        df = cargar_datos_excel(excel_path)
        if df is None:
            return jsonify({"status": "Error al cargar el archivo Excel"})

        # Inicializar el driver de Selenium
        driver = iniciar_driver()
        if driver is None:
            return jsonify({"status": "Error al iniciar el driver de Selenium"})

        # Obtener las universidades únicas del DataFrame
        universidades = df['universidades'].dropna().unique()
        resultados = []

        for universidad in universidades:
            logger.debug(f"Procesando universidad: {universidad}")

            # Encontrar coincidencias parciales de nombres de universidades
            coincidencias = process.extractBests(universidad, df['universidades'].dropna().astype(str).tolist(), scorer=fuzz.partial_ratio, score_cutoff=80)
            
            if not coincidencias:
                logger.error(f"No se encontró la universidad '{universidad}' en la configuración.")
                continue

            for nombre_df, _ in coincidencias:
                # Filtrar las filas correspondientes a la universidad encontrada
                filas_universidad = df[df['universidades'] == nombre_df]

                for index, fila in filas_universidad.iterrows():
                    # Construir la oferta para procesar
                    oferta = {
                        'titulo': fila['título de la oferta'],  # Ajustar nombres de columnas según tu archivo Excel
                        'area': fila['área'],
                        'cargo_equivalente': fila['cargo equivalente'],
                        'nivel_educativo': fila['nivel educativo'],
                        'subnivel': fila['subnivel'],
                        'cantidad_vacantes': fila['cantidad de vacantes'],
                        'rango_salario': fila['rango del salario en millones'],
                        'descripcion': fila['descripción'],
                        'requisitos': fila['requisitos'],
                        'departamento': fila['departamento'],
                        'ciudad': fila['ciudad'],
                        'sector': fila['sector'],
                        'subsectores': fila['subsectores'],
                        'nivel_estudio': fila['nivel de estudio'],
                        'profesion': fila['profecion'],
                        'experiencia_minima': fila['años totales de experiencia entre'],
                        'experiencia_maxima': fila['años totales de experiencia'],
                        'experiencia_requerida': fila['Experiencia requerida'],
                        'tipo_contrato': fila['tipo de contrato'],
                        'tiempo_dedicado': fila['tiempo dedicado']
                    }
                    # Procesar la universidad con la oferta construida
                    resultado = procesar_universidad(universidad, oferta, fila, driver)
                    resultados.append({'nombre': universidad, 'resultado': resultado})

        # Cerrar el driver de Selenium al finalizar
        driver.quit()
        return jsonify({"status": "Automatización completada", "resultados": resultados})

    except Exception as e:
        logger.error(f"Error en la automatización de universidades: {str(e)}")
        return jsonify({"status": f"Error: {str(e)}"}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=int(os.environ.get("PORT", 8180)))