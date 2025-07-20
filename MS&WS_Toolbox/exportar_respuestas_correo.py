import os
import pandas as pd
from datetime import datetime

def exportar_respuestas_correo():
    GENERADOS_DIR = "./Generados"
    DESTINO_DIR = "./RespuestasMail"
    DESTINO_FILE = os.path.join(DESTINO_DIR, "RespuestasCorreo.xlsx")

    # Asegurar que la carpeta de destino exista
    os.makedirs(DESTINO_DIR, exist_ok=True)

    archivos = [f for f in os.listdir(GENERADOS_DIR) if f.endswith(".xlsx")]
    if not archivos:
        print("❌ No hay archivos en 'Generados'")
        return

    print("📂 Archivos disponibles:")
    for i, nombre in enumerate(archivos, 1):
        print(f"{i}) {nombre}")

    try:
        seleccion = int(input("Seleccione un archivo: "))
        archivo_base = archivos[seleccion - 1]
    except (ValueError, IndexError):
        print("❌ Selección inválida.")
        return

    ruta_base = os.path.join(GENERADOS_DIR, archivo_base)
    print(f"\n🔍 Procesando: {archivo_base}")

    df_base = pd.read_excel(ruta_base)
    col_estado = "Estado Respuesta Correo"

    if col_estado not in df_base.columns:
        print(f"❌ La columna '{col_estado}' no existe.")
        return

    df_respondidos = df_base[df_base[col_estado].astype(str).str.strip().str.lower() == "respondido"].copy()
    if df_respondidos.empty:
        print("ℹ️ No hay registros con respuesta.")
        return

    df_respondidos["Fecha"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    columnas_excluir = [
        "Estado Dominio", "Estado Envío Correo", "Estado Respuesta Correo",
        "Estado Envío WA", "Estado Respuesta WA"
    ]
    df_respondidos = df_respondidos.drop(columns=[col for col in columnas_excluir if col in df_respondidos.columns])

    nuevos_registros = 0

    if os.path.exists(DESTINO_FILE):
        df_hist = pd.read_excel(DESTINO_FILE)

        claves = ["Telefono.", "E Mail"]
        if all(col in df_hist.columns for col in claves):
            claves_hist = df_hist[claves].astype(str)
            claves_nuevas = df_respondidos[claves].astype(str)

            df_nuevos = df_respondidos[
                ~claves_nuevas.apply(tuple, axis=1).isin(claves_hist.apply(tuple, axis=1))
            ]
            nuevos_registros = len(df_nuevos)
            df_final = pd.concat([df_hist, df_nuevos], ignore_index=True)
        else:
            print("⚠️ Archivo de historial no tiene columnas clave. Se sobrescribirá.")
            df_final = df_respondidos.copy()
            nuevos_registros = len(df_final)
    else:
        df_final = df_respondidos.copy()
        nuevos_registros = len(df_final)

    df_final.to_excel(DESTINO_FILE, index=False)
    print(f"\n✅ {nuevos_registros} nuevos registros exportados a: {DESTINO_FILE}")
