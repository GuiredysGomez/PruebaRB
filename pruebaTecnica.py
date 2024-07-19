import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.keys import Keys # Para enviar palabras
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import ssl
import smtplib
from email.message import EmailMessage
from confi import notification_email
from confi import passw

# INICIO Subrutina encargada de enviar el mail
def enviar(dest, asunto, cuerpo):
    # Mensaje del correo electrónico
    mail = EmailMessage()
    mail['From'] = notification_email
    mail['To'] = dest.strip().lower()
    mail['Subject'] = asunto
    mail.set_content(cuerpo)
    #Se agrega SSL
    context= ssl.create_default_context()
    
    #Inicia sesion y envia el mail
    # Como no se especifica, decidi usar como host gmail

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(notification_email, passw)
        smtp.sendmail(notification_email, dest.strip().lower(), mail.as_string())

# FIN Subrutina encargada de enviar el mail

# INICIO  el Main

# Ruta del archivo de Excel
excel_file = 'Base Seguimiento Observ Auditoría al_30042021.xlsx'  
app = xw.App(visible=False)  # Abre el Excel en modo invisible usando Xlwings
wb = app.books.open(excel_file)  # Abre el archivo de Excel
ws = wb.sheets['Hoja1']  # los datos se encuentran en la hoja "Hoja1"
driver = webdriver.Chrome()  # Inicializa el controlador de Chrome
audit_form_url = 'https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG'  # URL del formulario de auditoría

# Recorre cada fila en la hoja de Excel (los datos comienzan en la fila 2)
for row_num in range(2, ws.range(f'A{ws.cells.last_cell.row}').end('up').row + 1):  
    process = ws.cells(row_num, 1).value  # Obtiene el proceso de la columna A
    process_dict={
        'Operaciones': 'operaciones',
        'Cuentas por Cobrar':'cuentas',
        'Riesgo': 'riesgo',
        'TI': 'ti',
        'Financiero': 'financiero',
        'Continuidad Operacional': 'continuidad',
        'Contabilidad': 'contabilidad',
        'Gobierno Corp': 'gobierno'
    }
    process_status = ws.cells(row_num, 10).value  # Obtiene el estado del proceso de la columna J
    observation = ws.cells(row_num, 2).value  # Obtiene la observación de la columna B
    commitment_date = ws.cells(row_num, 6).value  # Obtiene la fecha de compromiso de la columna F    
    tipo_riesgo = ws.cells(row_num, 3).value  # Obtiene el tipo de riesgo de la columna C
    severidad = ws.cells(row_num, 4).value  # Obtiene la severidad de la columna D
    severidad_dict={
        'Alto': '2',
        'Bajo':'0',
        'Medio': '1'
    }
    responsable = ws.cells(row_num, 7).value  # Obtiene el responsable de la columna G
    # Obtiene el correo del responsable de la columna I
    dest = ws.cells(row_num, 9).value 
    # Procesa la informacion segun los estados
    if process_status == 'Regularizado':
        
        driver.get(audit_form_url)  # Abre la URL del formulario de auditoría

        process_input = Select(driver.find_element(By.ID,'process'))  # Busca el campo del proceso
        process_input.select_by_value(process_dict[process.strip()])       
        #time.sleep(3)

        tipo_riesgo_input = driver.find_element(By.ID,'tipo_riesgo')  # Busca el campo del tipo de riesgo
        tipo_riesgo_input.send_keys(tipo_riesgo)  # Ingresa el proceso en el campo
        #time.sleep(3)

        process_input = Select(driver.find_element(By.ID,'severidad'))  # Busca el campo severidad
        process_input.select_by_value(severidad_dict[severidad.strip()])
        #time.sleep(3)

        responsable_input = driver.find_element(By.ID,'res')  # Busca el campo del responsable
        responsable_input.send_keys(responsable)  # Ingresa el proceso en el campo     
        #time.sleep(3)   

        observation_input = driver.find_element(By.ID,'obs') #Busca el campo de entrada para la observación
        observation_input.send_keys(observation)  # Ingresa la observación en el campo
        #time.sleep(3)

        commitment_date_input = driver.find_element(By.ID,'date')  # Busca el campo de entrada para la fecha de compromiso
        commitment_date_input.send_keys(commitment_date.strftime('%d-%m-%Y'))  # Ingresa la fecha de compromiso en el campo con el formato 'DD-MM-YYYY'
        #time.sleep(3)

        submit_button = driver.find_element(By.ID,'submit')  # Busca el botón de enviar
        submit_button.click()  # Hace clic en el botón de enviar            

    elif process_status == 'Atrasado':         
        # Arma el mensaje para la persona responsable
        asunto = f'Proceso Atrasado: {process}'  
        cuerpo = f"""
            Proceso: {process}
            Estado: Atrasado
            Observación: {observation}
            Fecha de Compromiso: {commitment_date.strftime('%d-%m-%Y')}          
        """  
       # Llama subrutina de envio de mail
        enviar(dest, asunto, cuerpo)  # Envía el correo electrónico

    elif process_status == 'Pendientes':
        continue  # Ignora los procesos con estado "Pendientes"

wb.close()

# FIN Main