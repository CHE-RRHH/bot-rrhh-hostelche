import makeWASocket, {
  useMultiFileAuthState,
  DisconnectReason,
  fetchLatestBaileysVersion,
  downloadMediaMessage,
} from "@whiskeysockets/baileys";
import { Boom } from "@hapi/boom";
import pino from "pino";
import qrcode from "qrcode-terminal";
import fetch from "node-fetch";

// ─────────────────────────────────────────────
// ⚙️  CONFIGURACIÓN
// ─────────────────────────────────────────────
const CONFIG = {
  // Zona horaria
  timeZone: "America/Cancun",

  // Número admin que recibe reportes (formato: 521XXXXXXXXXX)
  numeroAdmin: "5219842326325",

  // Google Apps Script URL (sistema de fichaje — chehostel-fichaje)
  sheetsUrl: "https://script.google.com/macros/s/AKfycbwJ-d2SuAE3o5mdyrZVhFQ82PMPUa935PlDIiqk3BNXCFhep0QkcQENQXlgyFvtjrq1/exec",

  // Minutos sin actividad para cancelar sesión de fichaje
  expiracionMin: 5,

  // Sedes con coordenadas y radio en metros (mismas 12 que el sistema web)
  sedes: [
    { nombre: "Oficinas EUN",     lat: 20.623011, lng: -87.079899, radio: 100 },
    { nombre: "Che Playa",        lat: 20.626166, lng: -87.075494, radio: 100 },
    { nombre: "Che Tulum",        lat: 20.213802, lng: -87.456531, radio: 100 },
    { nombre: "Che Holbox",       lat: 21.521428, lng: -87.374776, radio: 100 },
    { nombre: "Che Bacalar",      lat: 18.682165, lng: -88.386827, radio: 100 },
    { nombre: "Che Merida",       lat: 20.973264, lng: -89.623262, radio: 100 },
    { nombre: "Che Valladolid",   lat: 20.687649, lng: -88.201307, radio: 100 },
    { nombre: "Che Puerto",       lat: 15.835287, lng: -97.041735, radio: 100 },
    { nombre: "Che Zipolite",     lat: 15.663439, lng: -96.519244, radio: 100 },
    { nombre: "Che San Cristobal",lat: 16.738113, lng: -92.635435, radio: 100 },
    { nombre: "Che Suite Playa",  lat: 20.620334, lng: -87.077838, radio: 100 },
    { nombre: "Che Suite Tulum",  lat: 20.214102, lng: -87.456717, radio: 100 },
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
function fechaISO() {
  const d = new Date(new Date().toLocaleString("en-US", { timeZone: CONFIG.timeZone }));
  const pad = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function horaGuion() {
  return horaStr().replace(":", "-");
}

// ─────────────────────────────────────────────
// 📊 GOOGLE SHEETS
// ─────────────────────────────────────────────
async function sheetsGet(params) {
  try {
    const url = CONFIG.sheetsUrl + '?' + new URLSearchParams(params).toString();
    const res = await fetch(url, {
      method: 'GET',
      headers: { 'Content-Type': 'application/json' },
      redirect: 'follow',
    });
    const text = await res.text();
    const trimmed = text.trim();
    // Solo intentar limpiar como JSONP si de verdad no parece JSON plano
    const clean = (trimmed.startsWith('{') || trimmed.startsWith('['))
      ? trimmed
      : trimmed.replace(/^[^(]+\(/, '').replace(/\);?$/, '').trim();
    return JSON.parse(clean);
  } catch (e) {
    console.error("Sheets error:", e.message);
    return { ok: false };
  }
}

async function sheetsGetEmpleados() {
  const data = await sheetsGet({ action: 'get_empleados' });
  if (data.ok) {
    console.log(`📋 Empleados cargados: ${data.empleados.length}`);
    return data.empleados;
  }
  return [];
}

async function sheetsGetLogHoy() {
  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const endOfDay = startOfDay + 86400000;
  const data = await sheetsGet({ action: 'get_log', ts_start: startOfDay, ts_end: endOfDay });
  return data.ok ? data.log : [];
}

async function sheetsFichar(empleado, tipo, hora, fecha, ts, sede, tieneFoto) {
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
      foto: tieneFoto ? "Sí (WhatsApp)" : "No",
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

async function subirFotoADrive(buffer, empleado, tipo) {
  try {
    const b64 = `data:image/jpeg;base64,${buffer.toString("base64")}`;
    const body = JSON.stringify({
      action: "guardar_selfie_fichaje",
      empId: empleado.id,
      nombre: empleado.name,
      apellido: empleado.apellido || "",
      fecha: fechaISO(),
      hora: horaGuion(),
      tipo,
      b64: encodeURIComponent(b64),
    });
    const res = await fetch(CONFIG.sheetsUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body,
    });
    const data = await res.json();
    return !!data.ok;
  } catch (e) {
    console.error("Error subiendo foto a Drive:", e.message);
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
  console.log(`📝 Texto extraido: "${texto}" | sesion previa: ${sesion ? sesion.paso : 'ninguna'}`);

  // ── COMANDO INICIAL ──────────────────────────
  if (!sesion && (texto === "fichar" || texto === "hola" || texto === "fichaje")) {
    // Verificar si el número está registrado antes de continuar
    const numero = jidANumero(jid);
    console.log(`🔍 Buscando empleado para numero: ${numero}`);
    const empleados = await sheetsGetEmpleados();
    const empleado = empleados.find(e => {
      const telSheet = String(e.telefono || "").replace(/\D/g, "");
      const telWA = numero.replace(/\D/g, "");
      return telSheet.slice(-10) === telWA.slice(-10);
    });

    if (!empleado) {
      console.log(`❌ No se encontro empleado con ese numero.`);
      return `❌ Tu número no está registrado en el sistema de fichaje de *Hostel Che*.\n\nContacta a RRHH para que te den de alta.`;
    }

    console.log(`✅ Empleado encontrado: ${empleado.name} (id ${empleado.id})`);
    setSesion(jid, { paso: "esperando_foto", empleadoId: empleado.id });
    return `👋 *Hola ${empleado.name}!*\n\n📸 Para registrar tu fichaje, primero envíame una *selfie* (foto tuya en este momento).`;
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

    // Guardar foto en memoria (no en disco: Railway borra el disco en cada redeploy)
    let fotoBuffer = null;
    try {
      fotoBuffer = await downloadMediaMessage(msg, "buffer", {});
    } catch (e) {
      console.error("Error descargando foto:", e.message);
    }

    setSesion(jid, { ...sesion, paso: "esperando_ubicacion", fotoBuffer });
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

    // Verificar empleado por número de teléfono
    const numero = jidANumero(jid);
    const empleados = await sheetsGetEmpleados();
    const empleado = empleados.find(e => {
      const telSheet = String(e.telefono || "").replace(/\D/g, "");
      const telWA = numero.replace(/\D/g, "");
      // Comparar últimos 10 dígitos para evitar diferencias de prefijo país
      return telSheet.slice(-10) === telWA.slice(-10);
    });

    if (!empleado) {
      delSesion(jid);
      return `❌ Tu número no está registrado en el sistema.\n\nContacta a RRHH para que te den de alta.`;
    }

    // Determinar entrada o salida
    const logHoy = await sheetsGetLogHoy();
    const registrosEmpleado = logHoy
      .filter(r => String(r.empId) === String(empleado.id))
      .sort((a, b) => Number(b.ts) - Number(a.ts));
    const ultimoRegistro = registrosEmpleado[0];
    const tipoFichaje = !ultimoRegistro || ultimoRegistro.type === "salida" ? "entrada" : "salida";

    const hora = horaStr();
    const fecha = fechaStr();
    const ts = Date.now();

    const [ok] = await Promise.all([
      sheetsFichar(empleado, tipoFichaje, hora, fecha, ts, verificacion.sede, !!sesion.fotoBuffer),
      sesion.fotoBuffer ? subirFotoADrive(sesion.fotoBuffer, empleado, tipoFichaje) : Promise.resolve(false),
    ]);
    delSesion(jid);

    const emoji = tipoFichaje === "entrada" ? "🟢" : "🔴";
    const accion = tipoFichaje === "entrada" ? "Entrada registrada" : "Salida registrada";

    if (ok) {
      return `${emoji} *${accion}*\n\n👤 ${empleado.name}\n🏨 ${verificacion.sede}\n🕐 ${hora} — ${fecha}\n\n_Guardado en el sistema ✓_`;
    } else {
      return `⚠️ Fichaje registrado pero hubo un error al sincronizar. Contacta a sistemas.`;
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
        console.log(`💬 Respuesta generada: ${respuesta ? respuesta.slice(0,60) : '(null, no responde)'}`);
        if (respuesta) {
          try {
            await sock.sendMessage(jid, { text: respuesta });
            console.log(`📤 Mensaje enviado a ${jidANumero(jid)}`);
          } catch (sendErr) {
            console.error(`🔴 Error al enviar mensaje:`, sendErr.message);
          }
        }
      } catch (error) {
        console.error("❌ Error:", error.message);
      }
    }
  });
}

console.log("🏨 Bot RRHH — Hostel Che\n");
conectar();
