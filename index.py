import csv
import smtplib
import requests
import time
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# -------------------------- CONFIGURACIÓN --------------------------

# Parámetros de conexión al servidor SMTP de Amazon SES
SMTP_HOST = 'xxxxxxxx'
SMTP_PORT = 587
SMTP_USERNAME = 'xxxxxxxx'
SMTP_PASSWORD = 'xxxxxxxx'

# Correo remitente (debe estar verificado en SES)
FROM_EMAIL = 'xxxxxxxx'

# Archivo CSV con la lista de destinatarios.
# Se asume que el CSV usa el delimitador ";" y tiene las columnas:
# "email", "archivo" (link al archivo PDF) y "nombre completo"
CSV_FILE = 'destinatarios.csv'

# Asunto del correo
ASUNTO = 'Formulario 107 año fiscal 2024 - UCACUE'

# Plantilla del cuerpo del correo en formato HTML.
CUERPO_HTML_TEMPLATE = """
<html>
  <body>
    <h4>Estimada/o  {nombre_completo},</h4>
    <p>Presente. - </p>
    <p>Reciba un cordial saludo; al tiempo que, adjunto al presente se remite el formulario 107 del Servicio de Rentas Internas - SRI, correspondiente al Comprobante de retenciones en la fuente del impuesto a la renta por ingresos del trabajo en relación de dependencia, año fiscal 2024.</p>
    <p>Atentamente,<br>
        JEFATURA FINANCIERA<br>
        Universidad Católica de Cuenca</p>
        <br>       
        <p style="font-size: 0.9em; color: #666;">
          Este es un mensaje generado automáticamente; por favor, no responda a este correo, ya que la dirección de envío no está habilitada para recibir respuestas.
        </p>     
  </body>
</html>
"""

# Tiempo de espera entre envíos (en segundos)
TIEMPO_ESPERA = 5

# Configuración del logging para generar un archivo de log (envio.log)
logging.basicConfig(
    level=logging.INFO,
    filename='envio.log',
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# -------------------------- FUNCIÓN PARA ADJUNTAR ARCHIVO --------------------------

def adjuntar_archivo(mensaje, url):
    """
    Descarga el archivo PDF desde la URL proporcionada y lo adjunta al mensaje.
    Se asume que todos los archivos son PDF.
    """
    try:
        response = requests.get(url)
        response.raise_for_status()  # Verifica que la descarga fue exitosa
        file_data = response.content
        # Extraer el nombre del archivo a partir de la URL
        filename = url.split('/')[-1]
        # Se crea el adjunto usando MIMEApplication y forzamos el _subtype_ a "pdf"
        adjunto = MIMEApplication(file_data, _subtype="pdf")
        adjunto.add_header('Content-Disposition', 'attachment', filename=filename)
        mensaje.attach(adjunto)
        logging.info(f"Adjunto descargado y agregado: {filename}")
    except Exception as e:
        logging.error(f"Error al descargar o adjuntar el archivo desde {url}: {e}")

# -------------------------- FUNCIÓN PARA ENVIAR CORREOS --------------------------

def enviar_correo(server, destinatario, archivo_url, nombre_completo):
    """
    Prepara y envía un correo electrónico al destinatario.
    Personaliza el saludo con el nombre completo y adjunta el PDF descargado.
    """
    # Crear el mensaje multipart
    mensaje = MIMEMultipart()
    mensaje['Subject'] = ASUNTO
    mensaje['From'] = FROM_EMAIL
    mensaje['To'] = destinatario

    # Personalizar el cuerpo del mensaje
    cuerpo_html = CUERPO_HTML_TEMPLATE.format(nombre_completo=nombre_completo)
    parte_html = MIMEText(cuerpo_html, 'html')
    mensaje.attach(parte_html)

    # Descargar y adjuntar el archivo PDF correspondiente
    adjuntar_archivo(mensaje, archivo_url)

    try:
        server.sendmail(FROM_EMAIL, destinatario, mensaje.as_string())
        logging.info(f"Correo enviado a {destinatario} (Usuario: {nombre_completo})")
        return True
    except Exception as e:
        logging.error(f"Error al enviar correo a {destinatario} (Usuario: {nombre_completo}): {e}")
        return False

# -------------------------- FUNCIÓN PRINCIPAL --------------------------

def main():
    total_correos = 0
    exitosos = 0
    fallidos = 0

    # Conectar al servidor SMTP de Amazon SES
    try:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        logging.info("Conexión SMTP establecida correctamente.")
    except Exception as e:
        logging.error(f"Error al conectar con el servidor SMTP: {e}")
        return

    # Leer el archivo CSV (con delimitador ';')
    try:
        with open(CSV_FILE, newline='', encoding='utf-8') as csvfile:
            lector = csv.DictReader(csvfile, delimiter=';')
            for fila in lector:
                total_correos += 1
                email_destino = fila.get("email")
                archivo_url = fila.get("archivo")
                nombre_completo = fila.get("nombre completo")
                if email_destino and archivo_url and nombre_completo:
                    if enviar_correo(server, email_destino, archivo_url, nombre_completo):
                        exitosos += 1
                    else:
                        fallidos += 1
                    time.sleep(TIEMPO_ESPERA)  # Espera entre envíos
                else:
                    logging.warning("Falta el correo, el link del archivo o el nombre completo en la fila.")
                    fallidos += 1
    except FileNotFoundError:
        logging.error(f"El archivo {CSV_FILE} no se encontró.")
    except Exception as e:
        logging.error(f"Error al leer el archivo CSV: {e}")

    # Cerrar la conexión SMTP
    server.quit()
    logging.info("Conexión SMTP cerrada.")

    # Registrar resumen final
    logging.info(f"Resumen de envíos: Total: {total_correos} | Exitosos: {exitosos} | Fallidos: {fallidos}")
    print(f"Resumen de envíos: Total: {total_correos} | Exitosos: {exitosos} | Fallidos: {fallidos}")

# -------------------------- EJECUCIÓN DEL SCRIPT --------------------------

if __name__ == '__main__':
    main()
