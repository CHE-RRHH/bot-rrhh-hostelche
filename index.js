import makeWASocket, {
  useMultiFileAuthState,
  DisconnectReason,
  fetchLatestBaileysVersion,
  downloadMediaMessage,
} from "@whiskeysockets/baileys";
import { Boom } from "@hapi/boom";
import pino from "pino";
import qrcode from "qrcode-terminal";
import { existsSync, mkdirSync, writeFileSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import fetch from "node-fetch";

const __dirname = dirname(fileURLToPath(import.meta.url));

// ─────────────────────────────────────────────
// ⚙️  CONFIGURACIÓN
// ─────────────────────────────────────────────
const CONFIG = {
  // Zona horaria
  timeZone: "America/Cancun",

  // Número admin que recibe reportes (formato: 521XXXXXXXXXX)
  numeroAdmin: "5219983411564",

  // Google Apps Script URL
  sheetsUrl: "https://script.google.com/macros/s/AKfycbwBwalYtWX1Q1poTu4Re-p_npZbBj8rpIXhnzQ1sA2UHgxPhh-BKrZ0wp5r9AYgIOac/exec",

  // Minutos sin actividad para cancelar sesión de fichaje
  expiracionMin: 5,

  // Sedes con coordenadas y radio en metros
  sedes: [
    { nombre: "Oficina",        lat: 20.623033, lng: -87.079929, radio: 50 },
    { nombre: "Che Playa",      lat: 20.626184, lng: -87.075446, radio: 50 },
    { nombre: "Che Suite Playa",lat: 20.620354, lng: -87.077728, radio: 50 },
  ],
};

// ─────────────────────────────────────────────
// 📍 UTILIDADES DE UBICACIÓN
// ─────────────────────────────────────────────
function distanciaMetros(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const dLat = ((lat2 - lat1) * Math.PI) / 180;
  const dLng = ((lng2 - lng1) * Math.PI) / 180;
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos((lat1 * Math.PI) / 180) *
    Math.cos((lat2 * Math.PI) / 180) *
    Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function verificarUbicacion(lat, lng) {
  for (const sede of CONFIG.sedes) {
    const dist = distanciaMetros(lat, lng, sede.lat, sede.lng);
    if (dist <= sede.radio) {
      return { ok: true, sede: sede.nombre, distancia: Math.round(dist) };
    }
  }
  const cercana = CONFIG.sedes
    .map(s => ({ ...s, dist: Math.round(distanciaMetros(lat, lng, s.lat, s.lng)) }))
    .sort((a, b) => a.dist - b.dist)[0];
  return { ok: false, cercana: cercana.nombre, distancia: cercana.dist };
}

// ─────────────────────────────────────────────
// 🕐 HORA LOCAL
// ─────────────────────────────────────────────
function ahoraLocal() {
  return new Date().toLocaleString("es-MX", { timeZone: CONFIG.timeZone });
}
function horaStr() {
  return new Date().toLocaleTimeString("es-MX", {
    timeZone: CONFIG.timeZone, hour: "2-digit", minute: "2-digit"
  });
}
function fechaStr() {
  return new Date().toLocaleDateString("es-MX", {
    timeZone: CONFIG.timeZone, day: "numeric", month: "numeric", year: "numeric"
  });
}

// ─────────────────────────────────────────────
// 📊 GOOGLE SHEETS
// ─────────────────────────────────────────────
async function sheetsGetEmpleados() {
  try {
    const url = `${CONFIG.sheetsUrl}?action=get_empleados&callback=cb`;
    const res = await fetch(url);
    const text = await res.text();
    const json = text.replace(/^cb\(/, "").replace(/\)$/, "");
    const data = JSON.parse(json);
    return data.ok ? data.empleados : [];
  } catch (e) {
    console.error("Error leyendo empleados:", e.message);
    return [];
  }
}

async function sheetsGetLogHoy() {
  try {
    const now = new Date();
    const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    const endOfDay = startOfDay + 86400000;
    const url = `${CONFIG.sheetsUrl}?action=get_log&ts_start=${startOfDay}&ts_end=${endOfDay}&callback=cb`;
    const res = await fetch(url);
    const text = await res.text();
    const json = text.replace(/^cb\(/, "").replace(/\)$/, "");
    const data = JSON.parse(json);
    return data.ok ? data.log : [];
  } catch (e) {
    return [];
  }
}

async function sheetsFichar(empleado, tipo, hora, fecha, ts, sede, fotoPath) {
  try {
    const params = new URLSearchParams({
      action: "fichar",
      empId: empleado.id,
      nombre: empleado.name,
      dept: empleado.dept,
      tipo,
      hora,
      fecha,
      ts,
      foto: fotoPath ? "Sí (WhatsApp)" : "No",
      ubicacion: sede,
      maps: "",
      canal: "WhatsApp",
    });
    await fetch(`${CONFIG.sheetsUrl}?${params.toString()}`, { method: "GET" });
    return true;
  } catch (e) {
    console.error("Error guardando fichaje:", e.message);
    return false;
  }
}

// ─────────────────────────────────────────────
// 🗃️ ESTADO DE SESIONES (en memoria)
// ─────────────────────────────────────────────
const sesiones = new Map();
// Estado: esperando_foto | esperando_ubicacion | esperando_pin

function getSesion(jid) { return sesiones.get(jid) || null; }
function setSesion(jid, data) {
  sesiones.set(jid, { ...data, ultimaActividad: Date.now() });
}
function delSesion(jid) { sesiones.delete(jid); }

function jidANumero(jid) { return jid.replace("@s.whatsapp.net", ""); }

// Expirar sesiones inactivas
function iniciarExpiraciones() {
  setInterval(() => {
    const ahora = Date.now();
    for (const [jid, sesion] of sesiones.entries()) {
      if (ahora - sesion.ultimaActividad > CONFIG.expiracionMin * 60 * 1000) {
        sesiones.delete(jid);
        console.log(`⏱️ Sesión expirada: ${jidANumero(jid)}`);
      }
    }
  }, 60000);
}

// ─────────────────────────────────────────────
// 🤖 LÓGICA PRINCIPAL
// ─────────────────────────────────────────────
async function procesarMensaje(sock, msg) {
  const jid = msg.key.remoteJid;
  const tipo = Object.keys(msg.message || {})[0];
  const sesion = getSesion(jid);

  // Extraer texto
  const texto = (
    msg.message?.conversation ||
    msg.message?.extendedTextMessage?.text ||
    ""
  ).trim().toLowerCase();

  // ── COMANDO INICIAL ──────────────────────────
  if (!sesion && (texto === "fichar" || texto === "hola" || texto === "fichaje")) {
    setSesion(jid, { paso: "esperando_foto" });
    return `👋 *Hola!* Soy el sistema de fichaje de *Hostel Che*.\n\n📸 Para registrar tu entrada o salida, primero envíame una *selfie* (foto tuya en este momento).`;
  }

  if (!sesion) {
    return `👋 Escribe *fichar* para registrar tu entrada o salida.`;
  }

  // ── PASO 1: ESPERANDO FOTO ───────────────────
  if (sesion.paso === "esperando_foto") {
    const esImagen = tipo === "imageMessage";
    if (!esImagen) {
      setSesion(jid, sesion);
      return `📸 Necesito que me envíes una *selfie* (foto). Por favor tómate una foto y envíala.`;
    }

    // Guardar foto
    let fotoPath = null;
    try {
      const buffer = await downloadMediaMessage(msg, "buffer", {});
      const fotosDir = join(__dirname, "fotos");
      if (!existsSync(fotosDir)) mkdirSync(fotosDir, { recursive: true });
      fotoPath = join(fotosDir, `${jidANumero(jid)}_${Date.now()}.jpg`);
      writeFileSync(fotoPath, buffer);
    } catch (e) {
      console.error("Error guardando foto:", e.message);
    }

    setSesion(jid, { ...sesion, paso: "esperando_ubicacion", fotoPath });
    return `✅ Foto recibida.\n\n📍 Ahora comparte tu *ubicación actual*.\n\n_En WhatsApp: pulsa el clip 📎 → Ubicación → Enviar mi ubicación actual_`;
  }

  // ── PASO 2: ESPERANDO UBICACIÓN ──────────────
  if (sesion.paso === "esperando_ubicacion") {
    const esUbicacion = tipo === "locationMessage";
    if (!esUbicacion) {
      setSesion(jid, sesion);
      return `📍 Necesito tu *ubicación*. Pulsa el clip 📎 → Ubicación → *Enviar mi ubicación actual*.`;
    }

    const lat = msg.message.locationMessage.degreesLatitude;
    const lng = msg.message.locationMessage.degreesLongitude;
    const verificacion = verificarUbicacion(lat, lng);

    if (!verificacion.ok) {
      delSesion(jid);
      return `❌ *Ubicación fuera de rango.*\n\nEstás a *${verificacion.distancia}m* de ${verificacion.cercana} (máximo permitido: 50m).\n\nSi crees que es un error, intenta de nuevo con *fichar*.`;
    }

    setSesion(jid, { ...sesion, paso: "esperando_pin", lat, lng, sede: verificacion.sede });
    return `✅ Ubicación verificada — *${verificacion.sede}* (${verificacion.distancia}m)\n\n🔐 Ahora escribe tu *PIN* de 4 dígitos.`;
  }

  // ── PASO 3: ESPERANDO PIN ────────────────────
  if (sesion.paso === "esperando_pin") {
    if (!/^\d{4}$/.test(texto)) {
      setSesion(jid, sesion);
      return `🔐 Por favor escribe tu *PIN de 4 dígitos*.`;
    }

    const empleados = await sheetsGetEmpleados();
    const empleado = empleados.find(e => String(e.pin) === texto);

    if (!empleado) {
      delSesion(jid);
      return `❌ *PIN incorrecto.* Sesión cancelada.\n\nEscribe *fichar* para intentar de nuevo.`;
    }

    // Determinar si es entrada o salida
    const logHoy = await sheetsGetLogHoy();
    const registrosEmpleado = logHoy
      .filter(r => String(r.empId) === String(empleado.id))
      .sort((a, b) => Number(b.ts) - Number(a.ts));
    const ultimoRegistro = registrosEmpleado[0];
    const tipoFichaje = !ultimoRegistro || ultimoRegistro.type === "salida" ? "entrada" : "salida";

    const hora = horaStr();
    const fecha = fechaStr();
    const ts = Date.now();

    const ok = await sheetsFichar(empleado, tipoFichaje, hora, fecha, ts, sesion.sede, sesion.fotoPath);
    delSesion(jid);

    const emoji = tipoFichaje === "entrada" ? "🟢" : "🔴";
    const accion = tipoFichaje === "entrada" ? "Entrada registrada" : "Salida registrada";

    if (ok) {
      return `${emoji} *${accion}*\n\n👤 ${empleado.name}\n🏨 ${sesion.sede}\n🕐 ${hora} — ${fecha}\n\n_Guardado en el sistema ✓_`;
    } else {
      return `⚠️ Fichaje registrado localmente pero hubo un error al sincronizar. Contacta a sistemas.`;
    }
  }

  return null;
}

// ─────────────────────────────────────────────
// 📋 REPORTE DIARIO (20:00)
// ─────────────────────────────────────────────
function iniciarReporteDiario(sock) {
  setInterval(async () => {
    const ahora = new Date().toLocaleTimeString("es-MX", {
      timeZone: CONFIG.timeZone, hour: "2-digit", minute: "2-digit", hour12: false
    });
    if (ahora !== "20:00") return;

    try {
      const log = await sheetsGetLogHoy();
      if (!log.length) return;

      const empleados = await sheetsGetEmpleados();
      const hoy = fechaStr();
      let reporte = `📋 *Reporte de asistencia — ${hoy}*\n\n`;

      const porEmpleado = {};
      for (const r of log) {
        const emp = empleados.find(e => String(e.id) === String(r.empId));
        const nombre = emp ? emp.name : r.nombre || "Desconocido";
        if (!porEmpleado[nombre]) porEmpleado[nombre] = [];
        porEmpleado[nombre].push(r);
      }

      for (const [nombre, registros] of Object.entries(porEmpleado)) {
        const entradas = registros.filter(r => r.type === "entrada");
        const salidas = registros.filter(r => r.type === "salida");
        const entrada = entradas[0]?.hora || "—";
        const salida = salidas[salidas.length - 1]?.hora || "Sin salida";
        reporte += `👤 *${nombre}*\n   ▶ Entrada: ${entrada} | ◀ Salida: ${salida}\n\n`;
      }

      await sock.sendMessage(`${CONFIG.numeroAdmin}@s.whatsapp.net`, { text: reporte });
    } catch (e) {
      console.error("Error reporte diario:", e.message);
    }
  }, 60000);
}

// ─────────────────────────────────────────────
// 🔌 CONEXIÓN BAILEYS
// ─────────────────────────────────────────────
async function conectar() {
  const { state, saveCreds } = await useMultiFileAuthState("auth_rrhh");
  const { version } = await fetchLatestBaileysVersion();

  const sock = makeWASocket({
    version,
    auth: state,
    logger: pino({ level: "silent" }),
    printQRInTerminal: false,
  });

  sock.ev.on("creds.update", saveCreds);

  sock.ev.on("connection.update", ({ connection, lastDisconnect, qr }) => {
    if (qr) {
      console.log("\n📱 Escanea este QR con WhatsApp:\n");
      qrcode.generate(qr, { small: true });
    }
    if (connection === "close") {
      const shouldReconnect =
        new Boom(lastDisconnect?.error)?.output?.statusCode !== DisconnectReason.loggedOut;
      console.log("🔄 Reconectando:", shouldReconnect);
      if (shouldReconnect) setTimeout(conectar, 3000);
      else console.log("❌ Sesión cerrada. Borra auth_rrhh/ y reinicia.");
    }
    if (connection === "open") {
      console.log("\n✅ Bot RRHH conectado!\n");
      iniciarExpiraciones();
      iniciarReporteDiario(sock);
    }
  });

  sock.ev.on("messages.upsert", async ({ messages, type }) => {
    if (type !== "notify") return;
    for (const msg of messages) {
      try {
        if (msg.key.fromMe) continue;
        if (!msg.message) continue;
        // Solo mensajes privados
        if (msg.key.remoteJid?.endsWith("@g.us")) continue;
        const tipoMsg = Object.keys(msg.message)[0];
        const ignorados = ["protocolMessage", "reactionMessage", "pollUpdateMessage", "senderKeyDistributionMessage"];
        if (ignorados.includes(tipoMsg)) continue;

        const jid = msg.key.remoteJid;
        const nombre = msg.pushName || jidANumero(jid);
        console.log(`📨 ${nombre}: ${tipoMsg}`);

        const respuesta = await procesarMensaje(sock, msg);
        if (respuesta) {
          await sock.sendMessage(jid, { text: respuesta });
        }
      } catch (error) {
        console.error("❌ Error:", error.message);
      }
    }
  });
}

console.log("🏨 Bot RRHH — Hostel Che\n");
conectar();
