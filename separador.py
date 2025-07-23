import pandas as pd

# Cargar el archivo Excel original
archivo_entrada = "contactos_filtrados.xlsx"
archivo_salida = "contactos_actualizado.xlsx"

# Leer hoja "Filtrados"
df_filtrados = pd.read_excel(archivo_entrada, sheet_name="Filtrados", engine="openpyxl")

# --- FILTRAR WHATSAPP ---
df_whatsapp = df_filtrados[
    df_filtrados["Telefono. "].notna() &
    (
        (df_filtrados["Estado Envío WA"] == "Pendiente") |
        (df_filtrados["Estado Respuesta WA"] == "Pendiente")
    )
]

columnas_whatsapp = [
    'ID', 'Razon comercial', 'Razon social', 'Estado ', 'Municipio', 'Calle  ',
    'Numero Exterior', 'Colonia. ', 'Codigo postal. ', 'Localidad. ', 'Telefono. ',
    'E Mail', 'Sitio web', 'Giro / Actividad', 'Rango de empleados', 'FechaCreada',
    'Estado Envío WA', 'Estado Respuesta WA'
]
df_whatsapp = df_whatsapp[columnas_whatsapp]

# --- FILTRAR CORREO ---
df_correos = df_filtrados[
    (df_filtrados["Correo"].notna() | df_filtrados["E Mail"].notna()) &
    (df_filtrados["Estado Dominio"] == "Existe")
]

columnas_correos = [
    'ID', 'Razon comercial', 'Razon social', 'Estado ', 'Municipio', 'Calle  ',
    'Numero Exterior', 'Colonia. ', 'Codigo postal. ', 'Localidad. ', 'Telefono. ',
    'E Mail', 'Sitio web', 'Giro / Actividad', 'Rango de empleados', 'FechaCreada',
    'Estado Dominio', 'Estado Envío Correo', 'Correo'
]
df_correos = df_correos[columnas_correos]

# --- GUARDAR ARCHIVO ACTUALIZADO ---
with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
    df_filtrados.to_excel(writer, sheet_name="Filtrados", index=False)
    df_whatsapp.to_excel(writer, sheet_name="Whatsapp", index=False)
    df_correos.to_excel(writer, sheet_name="Correos", index=False)

print("Archivo actualizado guardado como:", archivo_salida)
