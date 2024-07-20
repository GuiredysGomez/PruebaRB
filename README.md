# PruebaRB
Prueba Técnica Guiredys Gómez
Descripción del código:
Este código Python automatiza el proceso de seguimiento de observaciones en auditorías, utilizando datos de un archivo de Excel y enviando correos electrónicos a los responsables de los procesos o Rellenando un formulario web segun corresponda.

# Funcionalidades:
Lee datos de un archivo de Excel con información sobre las observaciones de la auditoría.
Clasifica los datos según su estado ("Regularizado", "Atrasado" o "Pendiente").
Para los datos "Regularizado", completa un formulario web con la información correspondiente.
Para los datos "Atrasado", envía un correo electrónico al responsable del proceso informando del retraso.
Ignora los datos "Pendiente".

# Estructura del código:
# Importación de bibliotecas:

xlwings: Para interactuar con el archivo de Excel.
selenium: Para controlar el navegador web y completar el formulario.
smtplib, email: Para enviar correos electrónicos.
datetime: Para formatear fechas.
ssl: Para establecer una conexión segura con el servidor de correo electrónico.
confi: Módulo personalizado que contiene las credenciales de correo electrónico.

# Subrutina enviar:

Define la función enviar que recibe como parámetros el destinatario, el asunto y el cuerpo del correo electrónico.
Utiliza smtplib para establecer una conexión segura con el servidor de correo electrónico (Gmail en este caso).
Se autentica con las credenciales almacenadas en el módulo confi.
Envía el correo electrónico utilizando la información proporcionada.

# Main:

Define la ruta del archivo de Excel, abre la aplicación de Excel en modo invisible y obtiene la hoja de trabajo con los datos.
Inicializa el controlador de Chrome para la automatización web.
Obtiene la URL del formulario de auditoría.
Recorre cada fila en la hoja de Excel, comenzando desde la fila 2.

Para cada fila:
Obtiene el proceso, el estado del proceso, la observación, la fecha de compromiso, el tipo de riesgo, la severidad y el responsable.

Clasifica los datos según su estado:
Regularizado:
Abre la URL del formulario de auditoría.
Selecciona el proceso en el campo correspondiente.
Ingresa el tipo de riesgo, la severidad, el responsable, la observación y la fecha de compromiso.
Envía el formulario.

Atrasado:
Arma un mensaje de correo electrónico con la información del proceso atrasado.
Llama a la subrutina enviar para enviar el correo electrónico al responsable.

Pendiente:
Ignora la observación.

Cierra el archivo de Excel.

Requisitos:
Tener instaladas las siguientes bibliotecas:
xlwings
selenium
smtplib
email
datetime
ssl
Tener un controlador de Chrome instalado y configurado correctamente.
Tener un archivo de Excel con la información de las observaciones de la auditoría, con la siguiente estructura:
Columna	Descripción
A	Proceso
B	Observación
C	Tipo de riesgo
D	Severidad
F	Fecha de compromiso
G	Responsable
I	Correo electrónico del responsable
J	Estado del proceso (Regularizado, Atrasado, Pendiente)


Consideraciones:
El código asume que la estructura del archivo de Excel y los nombres de las columnas son como se describen en la tabla anterior.
El código utiliza el módulo confi para almacenar las credenciales de correo electrónico de forma segura. Se recomienda llenar este modulo para el correcto funcionamiento del script.
El código está diseñado para funcionar con el formulario de auditoría específico de la URL proporcionada. Es posible que sea necesario realizar ajustes si el formulario cambia.

Uso:
Reemplace la ruta del archivo de Excel (excel_file) en el código con la ruta real del archivo que desea utilizar.
Asegúrese de tener las bibliotecas necesarias instaladas en su entorno de Python.
En el módulo confi y guarde las credenciales de correo electrónico en las variables notification_email y passw (Tener en cuenta que se usa el host de gmail en este script, por lo que el correo debe ser gmail).

# Bonus
Se adjunta un archivo .Json exportado desde la aplicacion Rocketbot Studio.
Dicho archivo contiene el robot "pruebaTecnicaConRocketbot" y realiza lo anteriormente mencionado en estedocumento.