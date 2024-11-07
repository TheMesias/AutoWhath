import pyautogui
import time
import pandas as pd
import pyperclip
from PIL import Image
import os
import win32clipboard
import io
import webbrowser
from datetime import datetime
import subprocess
import random

# Cargar datos de Excel
data_noviembre = pd.read_excel(r"C:\\automatizacion\\noviembre.xlsx")  # Nombres y teléfonos
data_variado = pd.read_excel(r"C:\\automatizacion\\variados.xlsx")  # Variables adicionales

IMAGES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "imagenes")
ERROR_IMAGE_PATH = os.path.join(os.getcwd(), "error_image.png")
INDEX_FILE_PATH = os.path.join(os.getcwd(), "indice_progreso.txt")


def esperar_random(min_seg=2, max_seg=5):
    time.sleep(random.uniform(min_seg, max_seg))


def simular_escritura(texto):
    for char in texto:
        pyperclip.copy(char)
        pyautogui.hotkey('ctrl', 'v')
        esperar_random(0.05, 0.2)


def guardar_indice_actual(indice):
    with open(INDEX_FILE_PATH, "w") as f:
        f.write(str(indice))


def cargar_indice_actual():
    try:
        with open(INDEX_FILE_PATH, "r") as f:
            return int(f.read().strip())
    except FileNotFoundError:
        return 0


def copiar_imagen_al_portapapeles(imagen):
    output = io.BytesIO()
    imagen.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()


def reiniciar_whatsapp():
    subprocess.call(["taskkill", "/F", "/IM", "WhatsApp.exe"])
    esperar_random(3, 5)
    pyautogui.click(x=1374, y=1059)
    esperar_random(10, 12)

def detectar_error():
    try:
        error_location = pyautogui.locateOnScreen(ERROR_IMAGE_PATH, confidence=0.6)
        return error_location is not None
    except pyautogui.ImageNotFoundException:
        return False

# Cargar las imágenes en memoria al inicio
imagenes_cargadas = {}
for i in range(1, 5):
    for ext in ['png', 'PNG', 'jpg', 'JPG']:
        ruta_imagen = os.path.join(IMAGES_DIR, f"{i}.{ext}")
        if os.path.exists(ruta_imagen):
            imagenes_cargadas[f"{i}.{ext}"] = Image.open(ruta_imagen)
            break
    else:
        print(f"Advertencia: No se encontró la imagen {i}.png o con otras extensiones")


ultimo_indice = cargar_indice_actual()
contador_mensajes = 0

for index, row_noviembre in data_noviembre.iloc[ultimo_indice:].iterrows():
    telefono = str(row_noviembre['TELEFONOS'])
    nombre = row_noviembre['NOMBRES']

    row_variado = data_variado.sample(n=1).iloc[0]
    
    saludo = row_variado['SALUDO']
    emoji = row_variado['EMOJI']
    temporal = row_variado['TEMPORAL']
    ofertas = row_variado['OFERTAS']
    ubicacion = row_variado['UBICACION']
    promociones = row_variado['PROMOCIONES']
    accion = row_variado['ACCION']
    urgencia = row_variado['URGENCIA']

    mensaje = f"""{saludo}, {nombre}
{temporal} {ofertas} 
Envios a {ubicacion}.
{promociones}
    
{accion} 
    
{emoji}Califica tu crédito al instante envianos tus:
{emoji} Nombres y Apellidos
{emoji} Número de cédula
{emoji} Ciudad de contacto    
 
{emoji}{urgencia}
    """

    url = f"whatsapp://send?phone=593{telefono}"
    webbrowser.open(url)
    esperar_random(2, 3)

    if detectar_error():
        print(f"El número {telefono} no tiene cuenta de WhatsApp.")
        pyautogui.click(x=1564, y=556)
        time.sleep(1)  
        while detectar_error():
            print("Esperando que se cierre el cuadro de error...")
            pyautogui.click(x=1564, y=556)
            time.sleep(1)
        time.sleep(1.5)  
        continue  


    pyautogui.click(x=1495, y=987)
    simular_escritura(mensaje)
    pyautogui.press('enter')
    esperar_random(1, 2)

    # Enviar las imágenes en orden aleatorio
    imagenes_a_enviar = list(imagenes_cargadas.items())  # Convertir el diccionario a una lista de tuplas
    random.shuffle(imagenes_a_enviar)  # Reordenar aleatoriamente la lista

    for imagen_nombre, imagen in imagenes_a_enviar:
        copiar_imagen_al_portapapeles(imagen)
        pyautogui.click(x=1495, y=987)  # Clic en el cuadro de texto de WhatsApp
        pyautogui.hotkey('ctrl', 'v')  # Pegar la imagen desde el portapapeles
        esperar_random(2, 3)
        pyautogui.press('enter')  # Enviar la imagen
        esperar_random(2, 3)


    guardar_indice_actual(index + 1)
    contador_mensajes += 1
