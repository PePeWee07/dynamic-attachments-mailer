import csv
import smtplib
import requests
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
FROM_EMAIL = 'xxxxxxxx@ucatolica.cue.ec'

# Archivo CSV con la lista de destinatarios.
# Se asume que el archivo tiene delimitador ";" y columnas: "email" y "archivo"
CSV_FILE = 'destinatarios.csv'

# Asunto del correo
ASUNTO = 'Aquí tienes el archivo solicitado'

# Cuerpo del correo en formato HTML (incluye la leyenda de no responder)
CUERPO_HTML = """
<html>
  <body>
    <p>Hola,</p>
    <p>Adjunto encontrarás el archivo solicitado.</p>
    <p>Esto es un correo automático, no responder.</p>
  </body>
</html>
"""

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
    except Exception as e:
        print(f"Error al descargar o adjuntar el archivo desde {url}: {e}")

# -------------------------- FUNCIÓN PARA ENVIAR CORREOS --------------------------

def enviar_correo(server, destinatario, archivo_url):
    """
    Prepara y envía un correo electrónico al destinatario.
    Se descarga y adjunta el archivo PDF desde la URL especificada en la fila del CSV.
    """
    # Crear el mensaje multipart
    mensaje = MIMEMultipart()
    mensaje['Subject'] = ASUNTO
    mensaje['From'] = FROM_EMAIL
    mensaje['To'] = destinatario

    # Adjuntar el cuerpo del mensaje (HTML)
    parte_html = MIMEText(CUERPO_HTML, 'html')
    mensaje.attach(parte_html)

    # Descargar y adjuntar el archivo PDF correspondiente
    adjuntar_archivo(mensaje, archivo_url)

    try:
        server.sendmail(FROM_EMAIL, destinatario, mensaje.as_string())
        print(f"Correo enviado a {destinatario}")
    except Exception as e:
        print(f"Error al enviar correo a {destinatario}: {e}")

# -------------------------- FUNCIÓN PRINCIPAL --------------------------

def main():
    # Conectar al servidor SMTP de Amazon SES
    try:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        print("Conexión SMTP establecida correctamente.")
    except Exception as e:
        print(f"Error al conectar con el servidor SMTP: {e}")
        return

    # Leer el archivo CSV (con delimitador ';')
    try:
        with open(CSV_FILE, newline='', encoding='utf-8') as csvfile:
            lector = csv.DictReader(csvfile, delimiter=';')
            for fila in lector:
                email_destino = fila.get("email")
                archivo_url = fila.get("archivo")  # Se asume que la columna con el link se llama "archivo"
                if email_destino and archivo_url:
                    enviar_correo(server, email_destino, archivo_url)
                else:
                    print("Falta el correo o el link del archivo en la fila.")
    except FileNotFoundError:
        print(f"El archivo {CSV_FILE} no se encontró.")
    except Exception as e:
        print(f"Error al leer el archivo CSV: {e}")

    # Cerrar la conexión SMTP
    server.quit()
    print("Conexión SMTP cerrada.")

# -------------------------- EJECUCIÓN DEL SCRIPT --------------------------

if __name__ == '__main__':
    main()
