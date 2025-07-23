import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

# Columnas de estado por canal
COLUMNAS_ESTADO = {
    "Filtrados": [
        "Estado Dominio",
        "Estado Envío Correo",
        "Estado Respuesta Correo",
        "Estado Envío WA",
        "Estado Respuesta WA"
    ],
    "Correos": [
        "Estado Envío Correo",
        "Estado Respuesta Correo"
    ],
    "Whatsapp": [
        "Estado Envío WA",
        "Estado Respuesta WA"
    ]
}

# Columnas clave para evitar duplicados
COLUMNAS_CLAVE = [
    "Razon comercial", "Telefono.", "E Mail",
    "Municipio", "Calle", "Numero Exterior", "Colonia.", "Código postal."
]

# Estados que se deben descartar (en mayúsculas)
ESTADOS_DESCARTADOS = {
    "ESTADO DE MEXICO", "MICHOACAN", "MORELOS", "GUERRERO",
    "OAXACA", "PUEBLA", "YUCATAN", "CDMX",
    "CIUDAD DE MEXICO", "DF", "DISTRITO FEDERAL"
}

def agregar_columnas_estado(df, columnas_estado):
    for columna in columnas_estado:
        if columna not in df.columns:
            df[columna] = "Pendiente"
    return df

def guardar_datos_ordenados(df_nuevos, ruta_guardado):
    fecha_actual = datetime.today().strftime('%d/%m/%Y')
    df_nuevos["FechaCreada"] = fecha_actual

    # Normalizar Estado y separar registros descartados
    if "Estado" in df_nuevos.columns:
        estado_normalizado = df_nuevos["Estado"].astype(str).str.upper().str.strip()
        df_descartados = df_nuevos[estado_normalizado.isin(ESTADOS_DESCARTADOS)].copy()
        df_validos = df_nuevos[~estado_normalizado.isin(ESTADOS_DESCARTADOS)].copy()
    else:
        print("⚠️ Advertencia: La columna 'Estado' no está presente. No se aplicó filtro de estados.")
        df_descartados = pd.DataFrame(columns=df_nuevos.columns)
        df_validos = df_nuevos.copy()

    # Cargar archivo existente si ya existe
    hojas_existentes = {}
    if os.path.exists(ruta_guardado):
        xls = pd.ExcelFile(ruta_guardado)
        for hoja in xls.sheet_names:
            hojas_existentes[hoja] = xls.parse(hoja)

    # Guardar datos válidos en hojas definidas
    for hoja, columnas_estado in COLUMNAS_ESTADO.items():
        df_copia = df_validos.copy()
        df_copia = agregar_columnas_estado(df_copia, columnas_estado)

        df_existente = hojas_existentes.get(hoja, pd.DataFrame(columns=df_copia.columns))
        df_existente = agregar_columnas_estado(df_existente, columnas_estado)

        df_combinado = pd.concat([df_existente, df_copia])

        columnas_clave_validas = [col for col in COLUMNAS_CLAVE if col in df_combinado.columns]
        if len(columnas_clave_validas) < len(COLUMNAS_CLAVE):
            faltantes = set(COLUMNAS_CLAVE) - set(columnas_clave_validas)
            print(f"⚠️ Columnas faltantes en '{hoja}' y no usadas para evitar duplicados: {faltantes}")

        df_combinado = df_combinado.drop_duplicates(subset=columnas_clave_validas, keep='first')
        df_combinado.sort_index(inplace=True)

        hojas_existentes[hoja] = df_combinado

    # Guardar hoja Descartados si tiene contenido
    if not df_descartados.empty:
        df_descartados.dropna(how='all', inplace=True)
        df_descartados = df_descartados.reset_index(drop=True)
        if "ID" not in df_descartados.columns:
            df_descartados.insert(0, "ID", range(1, len(df_descartados) + 1))
        hojas_existentes["Descartados"] = df_descartados

    # Guardar todas las hojas
    with pd.ExcelWriter(ruta_guardado, engine='openpyxl', mode='w') as writer:
        for hoja, df_hoja in hojas_existentes.items():
            df_hoja.dropna(how='all', inplace=True)
            df_hoja = df_hoja.reset_index(drop=True)

            if "ID" not in df_hoja.columns:
                df_hoja.insert(0, "ID", range(1, len(df_hoja) + 1))

            df_hoja.to_excel(writer, sheet_name=hoja, index=False)

    print(f"\n✅ Archivo actualizado. Hojas procesadas: {', '.join(hojas_existentes.keys())}")

def main():
    ruta_archivo = "lista_contactos.xlsx"
    ruta_guardado = os.path.join("GENERADOS", "contactos_filtrados.xlsx")

    if not os.path.exists("GENERADOS"):
        os.makedirs("GENERADOS")

    print("🚀 Iniciando script...")

    try:
        excel_file = pd.ExcelFile(ruta_archivo)
        hojas = excel_file.sheet_names
        print("\n📄 Hojas disponibles:")
        for i, h in enumerate(hojas, 1):
            print(f"{i}. {h}")
    except Exception as e:
        print(f"❌ Error al abrir archivo Excel: {e}")
        return

    try:
        seleccion = int(input("\n👉 Ingresa el número de hoja con la que deseas trabajar: "))
        if seleccion < 1 or seleccion > len(hojas):
            print("❌ Selección inválida.")
            return
        hoja_seleccionada = hojas[seleccion - 1]
        df = pd.read_excel(ruta_archivo, sheet_name=hoja_seleccionada)
    except Exception as e:
        print(f"❌ Error al leer hoja: {e}")
        return

    total_filas = len(df)
    print(f"\n📊 Total de filas disponibles: {total_filas}")

    try:
        fila_inicio = int(input("🔢 Fila inicial: ")) - 1
        fila_fin = int(input("🔢 Fila final: "))
        if fila_inicio < 0 or fila_fin > total_filas or fila_inicio >= fila_fin:
            print("❌ Rango inválido.")
            return
        df_filtrado = df.iloc[fila_inicio:fila_fin].copy()
    except Exception as e:
        print(f"❌ Error en selección de filas: {e}")
        return

    guardar_datos_ordenados(df_filtrado, ruta_guardado)

if __name__ == "__main__":
    main()
