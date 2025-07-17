# toolbox.py - Parte 1 de 3

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import time
import imaplib
import email
import os
import socket
from urllib.parse import urlparse
from openpyxl import load_workbook
from openpyxl.styles import Font, Color
import re
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from urllib.parse import urlparse


# === Configuración global ===
smtp_server = "smtp.gmail.com"
smtp_port = 587
imap_server = "imap.gmail.com"
tu_correo = ""
tu_password = ""
ruta_excel = "correos.xlsx"
nombre_archivo_pdf = "ejemplo.pdf"
chromedriver_path = "chromedriver.exe"
whatsapp_profile_path = "C:/Users/ALPHA_PC/AppData/Local/Google/Chrome/User Data/WhatsappProfile"
mensaje_a_enviar = "Hola"

# WhatsApp globals
driver = None
df_data = None  # Se sincroniza con df de correo

# === Excel: Cargar y guardar ===
def cargar_o_crear_df():
    try:
        df = pd.read_excel(ruta_excel, keep_default_na=False)
        print("💾 Archivo 'correos.xlsx' cargado con éxito.")
    except FileNotFoundError:
        print(f"❌ Archivo '{ruta_excel}' no encontrado. Creando uno nuevo...")
        columnas = ["E Mail", "Nombre", "Sitio web", "Telefono",
                    "Estatus Correo", "Estatus Envío", "Estatus Dominios", "Estatus Respuesta",
                    "Estatus Mensaje Whatsapp", "Estatus Respuesta Whatsapp"]
        df = pd.DataFrame(columns=columnas)
        for col in columnas[4:]:
            df[col] = "Pendiente"
    return df

def guardar_df_con_formato(df):
    try:
        with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            sheet = writer.sheets['Sheet1']
            columnas = ["Sitio web", "E Mail", "Estatus Correo", "Estatus Envío",
                        "Estatus Dominios", "Estatus Respuesta", "Estatus Mensaje Whatsapp", "Estatus Respuesta Whatsapp"]
            colores = {"Verde": "FF008000", "Rojo": "FFFF0000", "Naranja": "FFFF8C00"}
            indices = {col: -1 for col in columnas}
            for idx, cell in enumerate(sheet[1]):
                if cell.value in indices:
                    indices[cell.value] = idx + 1
            for index, row in df.iterrows():
                fila = index + 2
                for estatus_col in columnas[2:]:
                    valor = str(row.get(estatus_col, "")).strip().lower()
                    color = None
                    if valor in ["existe", "enviado", "confirmado", "respuesta", "respondido"]:
                        color = colores["Verde"]
                    elif valor in ["pendiente", "en espera"]:
                        color = colores["Naranja"]
                    elif valor in ["n/a", "rechazado", "fallido", "no existe"]:
                        color = colores["Rojo"]
                    if color and indices[estatus_col] != -1:
                        sheet.cell(row=fila, column=indices[estatus_col]).font = Font(color=Color(rgb=color))
            workbook.save(ruta_excel)
            print("✅ Archivo guardado con formato.")
    except Exception as e:
        print(f"❌ Error al guardar Excel: {e}")

# === Validaciones ===
def validar_correo_sintaxis(correo):
    regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(regex, str(correo))

def es_dominio_valido_para_envio(row):
    correo = str(row["E Mail"]).strip().lower()
    sitio_web = str(row.get("Sitio web", "")).strip().lower()
    estatus_dom = str(row.get("Estatus Dominios", "")).strip()
    if estatus_dom == "No existe":
        dominio_correo = correo.split("@")[-1] if "@" in correo else ""
        dominio_sitio = urlparse("http://" + sitio_web).netloc.replace("www.", "")
        return dominio_correo != dominio_sitio
    return True

# toolbox.py - Parte 2 de 3

def enviar_correos(df, limit):
    correos_a_enviar = df[
        (df["Estatus Correo"].str.lower() == "pendiente") & 
        df["E Mail"].apply(validar_correo_sintaxis).astype(bool)
    ].head(limit).copy()

    if correos_a_enviar.empty:
        print("\n✅ No hay correos pendientes.")
        return df, 0

    ruta_pdf = os.path.join(os.getcwd(), nombre_archivo_pdf)
    if not os.path.exists(ruta_pdf):
        print(f"\n❌ No se encontró el archivo '{nombre_archivo_pdf}'.")
        return df, 0

    print(f"\n--- Enviando {len(correos_a_enviar)} correos ---")

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(tu_correo, tu_password)

            for index, row in correos_a_enviar.iterrows():
                correo_destino = str(row["E Mail"]).strip().lower()

                if not es_dominio_valido_para_envio(row):
                    print(f"🚫 Saltado (dominio inválido): {correo_destino}")
                    df.at[index, "Estatus Correo"] = "N/A"
                    df.at[index, "Estatus Envío"] = "N/A"
                    df.at[index, "Estatus Respuesta"] = "N/A"
                else:
                    msg = MIMEMultipart()
                    msg['From'] = tu_correo
                    msg['To'] = correo_destino
                    msg['Subject'] = "Prueba de envío con adjunto"
                    msg.attach(MIMEText("Hola, te envío un archivo de prueba adjunto.", 'plain'))

                    with open(ruta_pdf, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header("Content-Disposition", f"attachment; filename={nombre_archivo_pdf}")
                        msg.attach(part)

                    server.sendmail(tu_correo, correo_destino, msg.as_string())
                    print(f"📤 Enviado: {correo_destino}")
                    df.at[index, "Estatus Correo"] = "Enviado"
                    df.at[index, "Estatus Envío"] = "En Espera"
                    df.at[index, "Estatus Respuesta"] = "En Espera"
                    time.sleep(1)
        print("✅ Lote de correos enviado.")
    except smtplib.SMTPAuthenticationError:
        print("❌ Error de autenticación. Revisa contraseña de aplicación.")
        return df, 0
    except Exception as e:
        print(f"❌ Error al enviar: {e}")
        return df, 0

    return df, len(correos_a_enviar)

def verificar_sitios_web(df, limit):
    """Verifica existencia de sitios web usando DNS y HTTP HEAD."""
    print("\n🌐 Verificando sitios web...")

    filas = df[df["Estatus Dominios"] == "Pendiente"].head(limit)
    if filas.empty:
        print("✅ Nada por verificar.")
        return df, 0

    procesados = 0
    for index, row in filas.iterrows():
        url_raw = str(row.get("Sitio web", "")).strip()
        if not url_raw or url_raw.lower() in ['nan', 'nan.0']:
            df.at[index, "Estatus Dominios"] = "No existe"
            continue

        # Normalizar URL
        if not url_raw.startswith(("http://", "https://")):
            url_http = "http://" + url_raw
        else:
            url_http = url_raw

        dominio = urlparse(url_http).netloc.replace("www.", "")
        existe = False

        # Paso 1: Verificación DNS
        try:
            socket.gethostbyname(dominio)
            existe = True
        except socket.gaierror:
            existe = False

        # Paso 2: Confirmar con HTTP si DNS falló
        if not existe:
            try:
                response = requests.head(url_http, timeout=5, allow_redirects=True)
                if response.status_code < 400:
                    existe = True
            except:
                existe = False

        resultado = "Existe" if existe else "No existe"
        df.at[index, "Estatus Dominios"] = resultado
        print(f"{'✅' if existe else '❌'} {url_raw} → {resultado}")
        procesados += 1

    return df, procesados

def verificar_rebotados(df, limit):
    pendientes = df[df["Estatus Envío"] == "En Espera"].head(limit)
    if pendientes.empty:
        print("✅ No hay correos pendientes de rebote.")
        return df, 0

    rebotados = []
    try:
        with imaplib.IMAP4_SSL(imap_server) as mail:
            mail.login(tu_correo, tu_password)
            carpetas = ["INBOX", "[Gmail]/Spam", "[Gmail]/Trash"]
            for carpeta in carpetas:
                try:
                    mail.select(carpeta)
                    typ, data = mail.search(None, '(FROM "mailer-daemon")')
                    for num in data[0].split():
                        typ, msg_data = mail.fetch(num, '(RFC822)')
                        raw_email = msg_data[0][1]
                        msg = email.message_from_bytes(raw_email)
                        body = ""
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    body += part.get_payload(decode=True).decode(errors="ignore")
                        else:
                            body = msg.get_payload(decode=True).decode(errors="ignore")
                        found = re.findall(r'[\w\.-]+@[\w\.-]+', body)
                        for correo in found:
                            if correo.lower() in pendientes["E Mail"].str.lower().values:
                                rebotados.append(correo.lower())
                except:
                    continue
    except Exception as e:
        print(f"❌ Error IMAP: {e}")
        return df, 0

    rebotados = list(set(rebotados))
    for index, row in pendientes.iterrows():
        correo = row["E Mail"].strip().lower()
        if correo in rebotados:
            df.at[index, "Estatus Envío"] = "Rechazado"
        else:
            df.at[index, "Estatus Envío"] = "Confirmado"
    return df, len(pendientes)

def verificar_respuestas_desde_excel(df, limit):
    print("\n📥 Verificando respuestas de correo...")
    lista = df[df["Estatus Respuesta"] == "En Espera"]["E Mail"].astype(str).str.strip().str.lower().head(limit).unique()
    if lista.size == 0:
        print("✅ No hay correos esperando respuesta.")
        return df, 0
    encontradas = 0
    try:
        with imaplib.IMAP4_SSL(imap_server) as mail:
            mail.login(tu_correo, tu_password)
            mail.select("INBOX")
            for correo in lista:
                search_query = f'(FROM "{correo}")'
                typ, data = mail.search(None, search_query)
                if data[0]:
                    print(f"📬 Respuesta recibida de: {correo}")
                    df.loc[df["E Mail"].str.lower() == correo, "Estatus Respuesta"] = "Respuesta"
                    encontradas += 1
    except Exception as e:
        print(f"❌ Error IMAP: {e}")
        return df, 0
    return df, encontradas
# toolbox.py - Parte 3 de 3

def load_excel_data():
    global df_data
    try:
        df_data = pd.read_excel(ruta_excel)
        if 'Estatus Mensaje Whatsapp' not in df_data.columns:
            df_data['Estatus Mensaje Whatsapp'] = 'Pendiente'
        if 'Estatus Respuesta Whatsapp' not in df_data.columns:
            df_data['Estatus Respuesta Whatsapp'] = 'Pendiente'
    except Exception as e:
        print(f"❌ Error al cargar Excel para WhatsApp: {e}")
        df_data = None

def save_excel_data_in_place():
    global df_data
    if df_data is not None:
        try:
            df_data.to_excel(ruta_excel, index=False)
            print("💾 Cambios guardados.")
        except Exception as e:
            print(f"❌ Error al guardar Excel: {e}")

def initialize_driver():
    global driver
    if driver is None:
        try:
            chrome_options = Options()
            chrome_options.add_argument(f"user-data-dir={whatsapp_profile_path}")
            service = Service(executable_path=chromedriver_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.get("https://web.whatsapp.com/")
            print("🔒 Escanea el código QR si no has iniciado sesión...")
            WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.ID, "pane-side"))
            )
            print("✅ WhatsApp Web iniciado.")
            return True
        except Exception as e:
            print(f"❌ Error al iniciar WebDriver: {e}")
            return False
    return True

def send_messages():
    global df_data
    if df_data is None:
        print("❌ Datos no cargados.")
        return
    for index, row in df_data.iterrows():
        telefono = str(row.get('Telefono', '')).strip()
        telefono = ''.join(filter(str.isdigit, telefono))
        if not telefono:
            continue
        if row.get("Estatus Mensaje Whatsapp") in ["Enviado", "Fallido"]:
            continue
        if not telefono.startswith('52'):
            telefono = f'521{telefono}'
        link = f"https://web.whatsapp.com/send?phone={telefono}&text={mensaje_a_enviar}"
        driver.get(link)
        time.sleep(5)
        try:
            boton = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@aria-label="Enviar"]'))
            )
            boton.click()
            df_data.at[index, 'Estatus Mensaje Whatsapp'] = 'Enviado'
            df_data.at[index, 'Estatus Respuesta Whatsapp'] = 'En Espera'
            print(f"📤 Mensaje enviado a {telefono}")
        except Exception as e:
            print(f"❌ Falló con {telefono}: {e}")
            df_data.at[index, 'Estatus Mensaje Whatsapp'] = 'Fallido'
            df_data.at[index, 'Estatus Respuesta Whatsapp'] = 'N/A - Fallido'
        save_excel_data_in_place()
        time.sleep(2)
    print("✅ Mensajes procesados.")

def check_responses():
    global df_data
    if df_data is None:
        return
    for index, row in df_data.iterrows():
        if row.get("Estatus Respuesta Whatsapp") != "En Espera":
            continue
        telefono = ''.join(filter(str.isdigit, str(row.get("Telefono", ""))))
        if not telefono:
            continue
        if not telefono.startswith('52'):
            telefono = f'521{telefono}'
        driver.get(f"https://web.whatsapp.com/send?phone={telefono}")
        time.sleep(5)
        try:
            mensaje_xpath = '(//div[contains(@class, "message-in")]//div[@class="copyable-text"])[last()]'
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, mensaje_xpath))
            )
            df_data.at[index, 'Estatus Respuesta Whatsapp'] = 'Respondido'
            print(f"📩 Respondió: {telefono}")
        except:
            print(f"⏳ Sin respuesta: {telefono}")
        save_excel_data_in_place()
        time.sleep(2)

def proceso_automatico(df):
    print("\n🔁 Ejecutando proceso automático completo...")
    for i in range(3):
        df, sitios = verificar_sitios_web(df, 5)
        df, enviados = enviar_correos(df, 5)
        guardar_df_con_formato(df)
        if sitios == 0 and enviados == 0:
            print("⚠️ Nada más que procesar.")
            break
        time.sleep(5)
    load_excel_data()
    if initialize_driver():
        send_messages()

def exportar_respuestas_separadas(df):
    try:
        # ------------------------------
        # 1. Respuestas por CORREO
        # ------------------------------
        df_correos = df[df["Estatus Respuesta"].str.lower() == "respuesta"]
        if not df_correos.empty:
            df_correos.to_excel("respuestas_correos.xlsx", index=False)
            print("✅ Archivo 'respuestas_correos.xlsx' generado.")
        else:
            print("📭 No hay respuestas por correo.")

        # ------------------------------
        # 2. Respuestas por WHATSAPP
        # ------------------------------
        df_whatsapp = df[df["Estatus Respuesta Whatsapp"].str.lower() == "respondido"]
        if not df_whatsapp.empty:
            df_whatsapp.to_excel("respuestas_whatsapp.xlsx", index=False)
            print("✅ Archivo 'respuestas_whatsapp.xlsx' generado.")
        else:
            print("📱 No hay respuestas por WhatsApp.")

    except Exception as e:
        print(f"❌ Error al exportar archivos: {e}")

def menu():
    df = cargar_o_crear_df()
    # === Limpieza automática de columnas de estatus ===
    estatus_cols = {
        "Estatus Correo": "Pendiente",
        "Estatus Envío": "Pendiente",
        "Estatus Dominios": "Pendiente",
        "Estatus Respuesta": "En Espera",
        "Estatus Mensaje Whatsapp": "Pendiente",
        "Estatus Respuesta Whatsapp": "En Espera"
    }

    for col, default in estatus_cols.items():
        if col in df.columns:
            df[col] = df[col].fillna("").replace("", default)

    load_excel_data()
    if df is None:
        print("❌ No se pudo cargar el Excel.")
        return
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("===================================")
        print("        TOOLBOX DE COMUNICACIÓN     ")
        print("===================================")
        print("1. Verificar sitios web")
        print("2. Enviar correos electrónicos")
        print("3. Enviar mensajes de WhatsApp")
        print("4. Verificar correos rebotados")
        print("5. Verificar respuestas de correo")
        print("6. Verificar respuestas de WhatsApp")
        print("7. 🔁 Proceso Automático Completo")
        print("8. Salir")
        print("9. Exportar respuestas a nuevo archivo")
        print("-----------------------------------")
        opcion = input("Elige una opción (1-8): ")

        if opcion == "1":
            df, _ = verificar_sitios_web(df, 25)
            guardar_df_con_formato(df)
        elif opcion == "2":
            df, _ = enviar_correos(df, 10)
            guardar_df_con_formato(df)
        elif opcion == "3":
            if initialize_driver():
                send_messages()
        elif opcion == "4":
            df, _ = verificar_rebotados(df, 10)
            guardar_df_con_formato(df)
        elif opcion == "5":
            df, _ = verificar_respuestas_desde_excel(df, 10)
            guardar_df_con_formato(df)
        elif opcion == "6":
            if initialize_driver():
                check_responses()
        elif opcion == "7":
            proceso_automatico(df)
        elif opcion == "8":
            print("👋 Cerrando toolbox...")
            break
        elif opcion == "9":
            exportar_respuestas_separadas(df)
        else:
            print("❌ Opción inválida.")
        input("\nPresiona Enter para continuar...")

    if driver:
        driver.quit()

if __name__ == "__main__":
    menu()
