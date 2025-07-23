import pandas as pd

# Ruta del archivo original
archivo_entrada = 'contactos_filtrados.xlsx'
hoja_origen = 'Filtrados'
archivo_salida = 'CorteDiario.xlsx'

# Leer la hoja
df = pd.read_excel(archivo_entrada, sheet_name=hoja_origen)

# Columnas que se deben conservar
columnas_deseadas = [
    'ID', 'Razon comercial', 'Razon social', 'Estado ', 'Municipio', 'Calle  ',
    'Numero Exterior', 'Colonia. ', 'Codigo postal. ', 'Localidad. ',
    'Telefono. ', 'E Mail', 'Sitio web', 'Giro / Actividad',
    'Rango de empleados', 'FechaCreada'
]

# Aplicar filtros según tus criterios
correos_existentes = df[df['Estado Envío Correo'].isin(['Enviado', 'No Rebotado'])][columnas_deseadas]
correos_basura = df[df['Estado Envío Correo'].isin(['No Aplica', 'Rebotado'])][columnas_deseadas]
whatsapp_existentes = df[df['Estado Envío WA'] == 'Enviado'][columnas_deseadas]
whatsapp_basura = df[df['Estado Envío WA'].isin(['No Aplica', 'No Existe', 'No existe'])][columnas_deseadas]
respuestas_correo = df[df['Estado Respuesta Correo'] == 'Respondido'][columnas_deseadas]
respuestas_whatsapp = df[df['Estado Respuesta WA'] == 'Respondido'][columnas_deseadas]

# Guardar en Excel
with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
    correos_existentes.to_excel(writer, sheet_name='CorreosExistentes', index=False)
    correos_basura.to_excel(writer, sheet_name='CorreosBasura', index=False)
    whatsapp_existentes.to_excel(writer, sheet_name='WhatsappExistentes', index=False)
    whatsapp_basura.to_excel(writer, sheet_name='WhatsappBasura', index=False)
    respuestas_correo.to_excel(writer, sheet_name='RespuestasCorreo', index=False)
    respuestas_whatsapp.to_excel(writer, sheet_name='RespuestasWhatsapp', index=False)

print(f'Archivo "{archivo_salida}" generado correctamente.')
