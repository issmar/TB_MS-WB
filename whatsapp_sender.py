from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import pandas as pd
import os # Importar el módulo os para manejar rutas de archivos

# --- Configuración del WebDriver ---

# ¡IMPORTANTE! Ajusta esta ruta a la ubicación REAL de tu chromedriver.exe
# Si chromedriver.exe está en la misma carpeta que tu script, puedes usar:
chromedriver_path = "chromedriver.exe" 

# Opciones para el navegador Chrome
chrome_options = Options()
# Mantener la sesión de WhatsApp Web iniciada usando un perfil de usuario.
chrome_options.add_argument(f"user-data-dir=C:/Users/ALPHA_PC/AppData/Local/Google/Chrome/User Data/WhatsappProfile")

# Variables globales para el driver y la URL base
driver = None
whatsapp_url = "https://web.whatsapp.com/"

# ¡AJUSTA ESTA RUTA A LA UBICACIÓN REAL DE TU ARCHIVO EXCEL DE ENTRADA Y SALIDA!
excel_file = "correos.xlsx" 

# DataFrame para almacenar los datos del Excel
df_data = None 

# --- Funciones ---

def load_excel_data():
    """Carga los datos del Excel al inicio."""
    global df_data
    try:
        df_data = pd.read_excel(excel_file)
        
        # Asegúrate de que las columnas de estatus existan; si no, créalas en memoria
        # Esto evitará errores si las columnas no están presentes al inicio.
        # Las creamos con valores predeterminados.
        if 'Estatus Mensaje Whatsapp' not in df_data.columns:
            df_data['Estatus Mensaje Whatsapp'] = 'Pendiente'
        if 'Estatus Respuesta Whatsapp' not in df_data.columns:
            df_data['Estatus Respuesta Whatsapp'] = 'Pendiente'
            
        print("Datos del Excel cargados exitosamente.")
    except FileNotFoundError:
        print(f"ERROR: El archivo Excel '{excel_file}' no se encontró. Verifica la ruta.")
        df_data = None 
    except Exception as e:
        print(f"ERROR al cargar el archivo Excel: {e}")
        df_data = None

def save_excel_data_in_place():
    """Guarda el DataFrame de vuelta al archivo Excel original, sin alterar columnas."""
    global df_data
    if df_data is not None:
        try:
            # Reordenar las columnas del DataFrame antes de guardar
            # para que coincidan con el orden original del Excel si se modificó internamente,
            # aunque pandas.read_excel ya mantiene el orden.
            # Lo más importante es que no añadimos columnas nuevas que no existan en el original.
            df_data.to_excel(excel_file, index=False)
            print("Cambios guardados en el archivo Excel.")
        except Exception as e:
            print(f"ERROR al guardar los cambios en el Excel: {e}")

def initialize_driver():
    """Inicializa el WebDriver y abre WhatsApp Web."""
    global driver
    if driver is None:
        print("Inicializando navegador y abriendo WhatsApp Web...")
        try:
            if chromedriver_path:
                service = Service(executable_path=chromedriver_path)
            else:
                service = Service()
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.get(whatsapp_url)

            print("Por favor, escanea el código QR de WhatsApp Web si aún no has iniciado sesión.")
            print("Una vez iniciado, el programa estará listo para las opciones del menú.")

            chat_panel_xpath = '//*[@id="pane-side"]'
            WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.XPATH, chat_panel_xpath))
            )
            print("¡Iniciaste Whatsapp!")
            return True
        except TimeoutException:
            print("ERROR: No se pudo iniciar WhatsApp Web a tiempo. Asegúrate de escanear el QR rápidamente.")
            if driver:
                driver.quit()
            driver = None
            return False
        except Exception as e:
            print(f"ERROR FATAL al inicializar el navegador: {e}")
            if driver:
                driver.quit()
            driver = None
            return False
    return True # Driver ya inicializado

def send_messages():
    """Envía mensajes a números y actualiza el DataFrame en memoria."""
    global df_data 
    if driver is None:
        print("El navegador no está iniciado. Por favor, inicialízalo primero.")
        return
    if df_data is None:
        print("No se pudieron cargar los datos. Por favor, reinicia el programa.")
        return

    print("Preparando el envío de mensajes...")
    mensaje_a_enviar = "Hola"

    for index, row in df_data.iterrows():
        telefono_original = row['Telefono'] 
        
        # Limpiar el número de teléfono para el link de WhatsApp
        telefono_limpio = str(telefono_original)
        telefono_limpio = ''.join(filter(str.isdigit, telefono_limpio))

        if pd.isna(telefono_limpio) or not telefono_limpio:
            print(f"Saltando fila {index}: Número de teléfono inválido o vacío: {telefono_original}")
            continue
        
        # Solo procesar si el estatus no es ya 'Enviado' o 'Fallido'
        if row['Estatus Mensaje Whatsapp'] == 'Enviado' or row['Estatus Mensaje Whatsapp'] == 'Fallido':
            print(f"Saltando {telefono_original}: Mensaje ya fue '{row['Estatus Mensaje Whatsapp']}'.")
            continue

        # Lógica para prefijo 521 (México)
        if not telefono_limpio.startswith('52'):
            whatsapp_link = f"https://web.whatsapp.com/send?phone=521{telefono_limpio}&text={mensaje_a_enviar}"
        elif len(telefono_limpio) == 10 and telefono_limpio.startswith('52'):
             whatsapp_link = f"https://web.whatsapp.com/send?phone=521{telefono_limpio[2:]}&text={mensaje_a_enviar}"
        else:
            whatsapp_link = f"https://web.whatsapp.com/send?phone={telefono_limpio}&text={mensaje_a_enviar}"

        print(f"Navegando a: {whatsapp_link} para enviar mensaje a {telefono_original}")
        driver.get(whatsapp_link)
        time.sleep(5) 

        try:
            send_button_xpath = '//*[@aria-label="Enviar"]' 
            send_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, send_button_xpath))
            )
            
            send_button.click()
            print(f"Mensaje enviado a: {telefono_original}")
            time.sleep(3) 
            
            # --- Actualizar estatus en DataFrame ---
            df_data.loc[index, 'Estatus Mensaje Whatsapp'] = 'Enviado'
            df_data.loc[index, 'Estatus Respuesta Whatsapp'] = 'En Espera'
            save_excel_data_in_place() # Guarda los cambios después de cada envío exitoso
            
        except Exception as e:
            print(f"ERROR: No se pudo enviar el mensaje a {telefono_original}. Detalles: {e}")
            df_data.loc[index, 'Estatus Mensaje Whatsapp'] = 'Fallido'
            df_data.loc[index, 'Estatus Respuesta Whatsapp'] = 'N/A - Fallido' 
            save_excel_data_in_place() # Guarda los cambios incluso si falla
            continue 

    print("Proceso de envío de mensajes completado.")


def check_responses():
    """Revisa si se han recibido respuestas y actualiza el DataFrame en memoria."""
    global df_data
    if driver is None:
        print("El navegador no está iniciado. Por favor, inicialízalo primero.")
        return
    if df_data is None:
        print("No se pudieron cargar los datos. Por favor, reinicia el programa.")
        return

    print("Revisando respuestas de los números en el Excel...")
        
    respuestas_detectadas = []

    # Iterar sobre las filas del DataFrame directamente para poder actualizarlo
    for index, row in df_data.iterrows():
        # Solo revisar si el estatus es 'En Espera'
        if row['Estatus Respuesta Whatsapp'] != 'En Espera':
            print(f"Saltando revisión de {row['Telefono']}: Estatus de respuesta no es 'En Espera'.")
            continue

        telefono_original = row['Telefono']
        # Limpiar el número de teléfono para el link de WhatsApp
        telefono_limpio = str(telefono_original)
        telefono_limpio = ''.join(filter(str.isdigit, telefono_limpio))
        
        if pd.isna(telefono_limpio) or not telefono_limpio: 
            print(f"Saltando revisión de fila {index}: Número de teléfono inválido o vacío: {telefono_original}")
            continue

        whatsapp_chat_url = f"https://web.whatsapp.com/send?phone={telefono_limpio}"
        print(f"Abriendo chat con: {telefono_original} para revisar respuestas.")
        driver.get(whatsapp_chat_url)
        time.sleep(5) 

        try:
            last_incoming_message_xpath = '(//div[contains(@class, "message-in")]//div[@class="copyable-text"])[last()]'
            
            last_incoming_message = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, last_incoming_message_xpath))
            )
            
            print(f"-> Respuesta detectada de: {telefono_original}")
            respuestas_detectadas.append(telefono_original)
            
            # Actualizar estatus en DataFrame de reporte
            df_data.loc[index, 'Estatus Respuesta Whatsapp'] = 'Respondido'
            save_excel_data_in_place() # Guarda el cambio inmediatamente
            
        except TimeoutException:
            print(f"-> No se detectó una respuesta de: {telefono_original} (o no hay mensajes entrantes recientes).")
            # El estatus se mantiene en "En Espera"
        except NoSuchElementException:
            print(f"-> No se detectó una respuesta de: {telefono_original} (elemento de mensaje no encontrado).")
            # El estatus se mantiene en "En Espera"
        except Exception as e:
            print(f"ERROR al revisar {telefono_original}: {e}")
        
        time.sleep(2) 

    print("\n--- Resumen de Respuestas Detectadas ---")
    if respuestas_detectadas:
        for num in respuestas_detectadas:
            print(f"- {num}")
    else:
        print("No se detectaron nuevas respuestas de los números con estatus 'En Espera'.")


def main_menu():
    """Muestra el menú principal y maneja las opciones."""
    global driver
    
    # Cargar los datos del Excel al inicio del programa
    load_excel_data()
    if df_data is None: 
        print("No se pudieron cargar los datos necesarios del Excel. Saliendo.")
        return

    # Inicializar el driver una vez
    if not initialize_driver():
        print("No se pudo iniciar el navegador. Saliendo.")
        return

    while True:
        print("\n--- Menú Principal ---")
        print("1 - Enviar Mensaje")
        print("2 - Revisar Respuestas")
        print("3 - Salir")
        
        choice = input("Elige una opción: ")

        if choice == '1':
            send_messages()
        elif choice == '2':
            check_responses()
        elif choice == '3':
            print("Saliendo del programa. Cerrando navegador...")
            break
        else:
            print("Opción no válida. Por favor, elige 1, 2 o 3.")

    if driver:
        driver.quit()
        print("Navegador cerrado.")

# --- Ejecutar el menú ---
if __name__ == "__main__":
    main_menu()