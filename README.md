# Envío Masivo de Correos con Amazon SES

Este proyecto es una solución en Python para el envío masivo de correos electrónicos utilizando el servicio SMTP de Amazon SES. El script permite:

- Leer una lista de destinatarios desde un archivo CSV.
- Personalizar el mensaje con el nombre completo de cada destinatario.
- Descargar y adjuntar un archivo PDF específico para cada usuario (la URL del archivo se obtiene desde el CSV).
- Reiniciar la conexión SMTP después de enviar un número determinado de correos (por ejemplo, 200 por sesión) para evitar superar los límites de mensajes por conexión.
- Esperar un intervalo configurable (por ejemplo, 5 segundos) entre cada envío para evitar saturar el servidor.
- Generar logs detallados de los envíos, incluyendo información sobre correos enviados, errores y reintentos.

## Características

- **Personalización del mensaje:**  
  Se utiliza una plantilla HTML que saluda al destinatario por su nombre completo e incluye el contenido del mensaje.
  
- **Adjuntos dinámicos:**  
  El script descarga el archivo PDF desde la URL especificada en el CSV y lo adjunta al correo de forma individual.

- **Gestión de conexiones SMTP:**  
  Para evitar errores de "Maximum message count per session reached", el script cierra y reabre la conexión SMTP cada cierto número de envíos.

- **Control de envío:**  
  Se implementa una pausa configurable entre cada envío para respetar los límites del servidor y evitar problemas de conexión.

- **Logging:**  
  Se registra información detallada en un archivo `envio.log` para auditar quién recibió el correo, cuántos envíos fueron exitosos, cuántos fallaron, y otros detalles importantes.

## Prerrequisitos

- **Python 3.x**
- Librerías de Python necesarias (la mayoría son parte de la biblioteca estándar):
  - `csv`
  - `smtplib`
  - `time`
  - `logging`
  - `email`
  - [**requests**](https://pypi.org/project/requests/)

## Instalación

1. **Clona o descarga** este repositorio en tu máquina.
2. **Instala las dependencias** necesarias. Por ejemplo, para instalar `requests`:
   ```bash
   pip install requests
