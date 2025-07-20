import os
from checador_dns import ejecutar_checador_dns
from separador_filas import ejecutar_separador
from enviador_correo import enviar_correos
from checador_envio_correo import verificar_rebotes_correo
from checador_respuesta_correo import verificar_respuestas_correo
from enviador_wa import enviar_mensajes_wa
from checador_respuesta_wa import verificar_respuestas_wa
from exportar_respuestas_correo import exportar_respuestas_correo
from exportar_respuestas_wa import exportar_respuestas_wa




# ─────────────────────────────────────────────
# Bienvenida y autenticación (solo visual)
# ─────────────────────────────────────────────
print("📧 Inicio de sesión")
correo = input("Correo de Gmail: ").strip()
contrasena = input("Contraseña de aplicación: ").strip()

# ─────────────────────────────────────────────
# Menú principal
# ─────────────────────────────────────────────
def mostrar_menu():
    print("\n📋 Menú principal:")
    print("1) Separar filas (crear por lotes)")
    print("2) Verificar dominios (checador_dns)")
    print("3) Enviar correos (enviador_correo)")
    print("4) Verificar rebotes (checador_envio_correo)")
    print("5) Verificar respuestas (checador_respuesta_correo)")
    print("6) Enviar por WhatsApp (enviador_wa)")
    print("7) Verificar respuestas de WhatsApp")
    print("8) Exportar respuestas de correo")
    print("9) Exportar respuestas de WhatsApp")

    print("4) Salir")

while True:
    mostrar_menu()
    opcion = input("\nSeleccione una opción: ").strip()

    if opcion == "1":
        ejecutar_separador()

    elif opcion == "2":
        ejecutar_checador_dns()

    elif opcion == "3":
        enviar_correos(correo, contrasena)

    elif opcion == "4":
        verificar_rebotes_correo(correo, contrasena)

    elif opcion == "5":
        verificar_respuestas_correo(correo, contrasena)

    elif opcion == "6":
        enviar_mensajes_wa()

    elif opcion == "7":
        verificar_respuestas_wa()

    elif opcion == "8":
        exportar_respuestas_correo()

    elif opcion == "9":
        exportar_respuestas_wa()

    elif opcion == "":
        print("👋 Saliendo...")
        break

    else:
        print("❌ Opción no válida.")
