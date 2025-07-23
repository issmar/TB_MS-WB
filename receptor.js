const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');

// Nombre exacto del grupo donde quieres reenviar los mensajes
const NOMBRE_DEL_GRUPO = "Checker";

// Crear cliente con sesión guardada
const client = new Client({
    authStrategy: new LocalAuth()
});

const mensajesProcesados = new Set();
let grupoDestino = null;

// Mostrar QR
client.on('qr', qr => {
    qrcode.generate(qr, { small: true });
    console.log('🔐 Escanea el QR con tu WhatsApp');
});

// Confirmar inicio
client.on('ready', async () => {
    console.log('✅ Sesión iniciada. Buscando grupo...');

    const chats = await client.getChats();
    grupoDestino = chats.find(chat => chat.isGroup && chat.name === NOMBRE_DEL_GRUPO);

    if (grupoDestino) {
        console.log(`✅ Grupo encontrado: ${grupoDestino.name}`);
    } else {
        console.log(`❌ No se encontró el grupo llamado "${NOMBRE_DEL_GRUPO}"`);
    }
});

// Cuando llega un mensaje
client.on('message', async message => {
    // Evitar mensajes duplicados
    if (mensajesProcesados.has(message.id._serialized)) return;
    mensajesProcesados.add(message.id._serialized);

    const contact = await message.getContact();
    const texto = message.body;

    // Obtener nombre o número
    let nombre = contact.pushname || contact.name;
    if (!nombre || nombre.trim() === '') {
        nombre = contact.number || message.from;
    }

    // Mostrar en consola
    console.log(`\nDe: ${nombre}`);
    console.log(`Mensaje: ${texto}`);

    // Reenviar al grupo si no es del bot y no está vacío
    if (grupoDestino && !message.fromMe && texto.trim() !== "") {
        const mensajeFormateado = `De: ${nombre}\nMensaje: ${texto}`;
        await client.sendMessage(grupoDestino.id._serialized, mensajeFormateado);
        console.log("📤 Mensaje reenviado al grupo.");
    }
});

client.initialize();
