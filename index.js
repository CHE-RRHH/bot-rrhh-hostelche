const SHEET_ID = '16ucJYG-fAzRs4ZjK_eF7rnGmGaqYnZgSRtCfto1QGPw';
const DRIVE_FOLDER_ID = '1doWm3fsbgzZezx3bucF-FQEt6GgU0H9u';
const FICHAR_BASE_URL = 'https://chehostel-fichaje.pages.dev/fichar.html';

function doGet(e) {
  const data = e.parameter;
  const result = handleRequest(data);
  if (data.callback) {
    return ContentService
      .createTextOutput(data.callback + '(' + result + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    return ContentService.createTextOutput(handleRequest(data)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function crearCarpetaEmpleado(nombre, apellido, id) {
  try {
    const carpetaRaiz = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const nombreCarpeta = `${nombre} ${apellido || ''} [${id}]`.trim();
    const existing = carpetaRaiz.getFoldersByName(nombreCarpeta);
    if (existing.hasNext()) {
      return existing.next().getUrl();
    }
    const nuevaCarpeta = carpetaRaiz.createFolder(nombreCarpeta);
    nuevaCarpeta.createFolder('Identificacion');
    nuevaCarpeta.createFolder('Contrato');
    nuevaCarpeta.createFolder('IMSS');
    nuevaCarpeta.createFolder('Comprobante domicilio');
    nuevaCarpeta.createFolder('Otros');
    return nuevaCarpeta.getUrl();
  } catch(e) {
    return '';
  }
}

function handleRequest(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (data.action === 'get_empleados') {
      let sheet = ss.getSheetByName('Empleados');
      if (!sheet) return JSON.stringify({ok:true, empleados:[]});
      const tz = Session.getScriptTimeZone() || 'America/Cancun';
      const fmtFecha = v => v instanceof Date ? Utilities.formatDate(v, tz, 'dd/MM/yyyy') : String(v || '');
      const rows = sheet.getDataRange().getValues();
      const empleados = rows.slice(1).filter(r => r[0]).map(r => ({
        id: String(r[0]), name: String(r[1]), apellido: String(r[2]||''),
        genero: String(r[3]||''), puesto: String(r[4]||''), dept: String(r[5]||''),
        sede: String(r[6]||''), fechaIngreso: fmtFecha(r[7]),
        telefono: String(r[8]||''), email: String(r[9]||''),
        curp: String(r[10]||''), rfc: String(r[11]||''),
        direccion: String(r[12]||''), contactoEmergencia: String(r[13]||''),
        telefonoEmergencia: String(r[14]||''), pin: String(r[15]||'0000'),
        color: Number(r[16])||0, foto: String(r[17]||''),
        estado: String(r[18]||'activo'), driveUrl: String(r[19]||''),
        fechaNacimiento: fmtFecha(r[20]),
        sinGPS: String(r[21]||'').toLowerCase()==='si',
        sinFoto: String(r[23]||'').toLowerCase()==='si',
        puedeHomeOffice: String(r[24]||'').toLowerCase()==='si',
      }));
      return JSON.stringify({ok:true, empleados});
    }

    if (data.action === 'get_log') {
      let sheet = ss.getSheetByName('Fichajes');
      if (!sheet) return JSON.stringify({ok:true, log:[]});
      const tsStart = Number(data.ts_start);
      const tsEnd = Number(data.ts_end);
      const tz = Session.getScriptTimeZone() || 'America/Cancun';
      const rows = sheet.getDataRange().getValues();
      const log = rows.slice(1).filter(r => {
        const ts = Number(r[6]);
        return ts >= tsStart && ts <= tsEnd;
      }).map(r => {
        const horaVal = r[4];
        const fechaVal = r[5];
        const hora = horaVal instanceof Date ? Utilities.formatDate(horaVal, tz, 'HH:mm') : String(horaVal || '');
        const dateStr = fechaVal instanceof Date ? Utilities.formatDate(fechaVal, tz, 'dd/MM/yyyy') : String(fechaVal || '');
        return {
          empId: String(r[0]), nombre: String(r[1]), dept: String(r[2]),
          type: String(r[3]), hora, dateStr,
          ts: Number(r[6]), foto: String(r[7]||''), ubicacion: String(r[8]||''),
          maps: String(r[9]||''), canal: String(r[10]||'')
        };
      });
      return JSON.stringify({ok:true, log});
    }

    if (data.action === 'fichar') {
      let sheet = ss.getSheetByName('Fichajes');
      if (!sheet) {
        sheet = ss.insertSheet('Fichajes');
        sheet.appendRow(['ID Empleado','Empleado','Departamento','Tipo','Hora','Fecha','Timestamp','Foto','Ubicacion','Maps','Canal']);
      }
      sheet.appendRow([data.empId, data.nombre, data.dept, data.tipo, data.hora, data.fecha, Number(data.ts), data.foto||'No', data.ubicacion||'', data.maps||'', data.canal||'Tablet']);
      return JSON.stringify({ok:true});
    }

    if (data.action === 'add_empleado') {
      let sheet = ss.getSheetByName('Empleados');
      if (!sheet) {
        sheet = ss.insertSheet('Empleados');
        sheet.appendRow(['ID','Nombre','Apellido','Genero','Puesto','Departamento','Sede','FechaIngreso','Telefono','Email','CURP','RFC','Direccion','ContactoEmergencia','TelefonoEmergencia','PIN','Color','Foto','Estado','DriveURL','FechaNacimiento','SinGPS','TelegramID','SinFoto','PuedeHomeOffice']);
      }
      const driveUrl = crearCarpetaEmpleado(data.nombre, data.apellido, data.id);
      sheet.appendRow([
        data.id, data.nombre, data.apellido||'', data.genero||'',
        data.puesto||'', data.dept||'', data.sede||'', data.fechaIngreso||'',
        data.telefono||'', data.email||'', data.curp||'', data.rfc||'',
        data.direccion||'', data.contactoEmergencia||'', data.telefonoEmergencia||'',
        data.pin||'0000', data.color||0, data.foto||'', data.estado||'activo', driveUrl,
        data.fechaNacimiento||'', String(data.sinGPS)==='true'?'Si':'No', '',
        String(data.sinFoto)==='true'?'Si':'No',
        String(data.puedeHomeOffice)==='true'?'Si':'No'
      ]);
      if (data.email) {
        enviarCorreoCredenciales({
          id: data.id, nombre: data.nombre, apellido: data.apellido||'',
          email: data.email, pin: data.pin||'0000',
        });
      }
      return JSON.stringify({ok:true, driveUrl});
    }

    if (data.action === 'update_empleado') {
      let sheet = ss.getSheetByName('Empleados');
      if (!sheet) return JSON.stringify({ok:false, error:'No existe hoja Empleados'});
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          const driveUrl = data.driveUrl || String(rows[i][19]||'');
          const fechaNac = data.fechaNacimiento !== undefined ? data.fechaNacimiento : String(rows[i][20]||'');
          const sinGPS = data.sinGPS !== undefined ? (String(data.sinGPS)==='true'?'Si':'No') : String(rows[i][21]||'No');
          const sinFoto = data.sinFoto !== undefined ? (String(data.sinFoto)==='true'?'Si':'No') : String(rows[i][23]||'No');
          const puedeHomeOffice = data.puedeHomeOffice !== undefined ? (String(data.puedeHomeOffice)==='true'?'Si':'No') : String(rows[i][24]||'No');
          sheet.getRange(i+1, 1, 1, 22).setValues([[
            data.id, data.nombre, data.apellido||'', data.genero||'',
            data.puesto||'', data.dept||'', data.sede||'', data.fechaIngreso||'',
            data.telefono||'', data.email||'', data.curp||'', data.rfc||'',
            data.direccion||'', data.contactoEmergencia||'', data.telefonoEmergencia||'',
            data.pin||'0000', data.color||0, data.foto||'', data.estado||'activo', driveUrl,
            fechaNac, sinGPS
          ]]);
          sheet.getRange(i+1, 24).setValue(sinFoto);
          sheet.getRange(i+1, 25).setValue(puedeHomeOffice);
          return JSON.stringify({ok:true});
        }
      }
      return JSON.stringify({ok:false, error:'Empleado no encontrado'});
    }

    if (data.action === 'crear_carpeta') {
      const driveUrl = crearCarpetaEmpleado(data.nombre, data.apellido, data.id);
      if (driveUrl) {
        let sheet = ss.getSheetByName('Empleados');
        const rows = sheet.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          if (String(rows[i][0]) === String(data.id)) {
            sheet.getRange(i+1, 20).setValue(driveUrl);
            break;
          }
        }
      }
      return JSON.stringify({ok:true, driveUrl});
    }

    if (data.action === 'del_empleado') {
      let sheet = ss.getSheetByName('Empleados');
      if (!sheet) return JSON.stringify({ok:true});
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][0]) === String(data.id)) {
          const driveUrl = String(rows[i][19] || '');
          const match = driveUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
          if (match) {
            try { DriveApp.getFolderById(match[1]).setTrashed(true); } catch (e) {}
          }
          sheet.deleteRow(i + 1);
        }
      }
      return JSON.stringify({ok:true});
    }

    if (data.action === 'subir_foto') {
      try {
        const b64 = decodeURIComponent(data.b64);
        const base64Data = b64.replace(/^data:image\/(png|jpg|jpeg|gif);base64,/, '');
        const mimeType = b64.match(/data:(image\/\w+);/)[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, `foto_perfil.jpg`);

        const carpetaRaiz = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        const nombre = data.nombre || '';
        const apellido = data.apellido || '';
        const id = data.id || '';
        const nombreCarpeta = `${nombre} ${apellido} [${id}]`.trim();

        let carpetaEmp;
        const existing = carpetaRaiz.getFoldersByName(nombreCarpeta);
        if (existing.hasNext()) {
          carpetaEmp = existing.next();
        } else {
          carpetaEmp = carpetaRaiz.createFolder(nombreCarpeta);
        }

        const archivos = carpetaEmp.getFilesByName('foto_perfil.jpg');
        while (archivos.hasNext()) archivos.next().setTrashed(true);

        const file = carpetaEmp.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const url = `https://drive.google.com/uc?id=${file.getId()}`;

        const sheet = ss.getSheetByName('Empleados');
        if (sheet) {
          const rows = sheet.getDataRange().getValues();
          for (let i = 1; i < rows.length; i++) {
            if (String(rows[i][0]) === String(id)) {
              sheet.getRange(i+1, 18).setValue(url);
              break;
            }
          }
        }
        return JSON.stringify({ok: true, url});
      } catch(e) {
        return JSON.stringify({ok: false, error: e.message});
      }
    }

    if (data.action === 'guardar_foto_perfil_drive') return guardarFotoPerfilDrive(data);
    if (data.action === 'guardar_selfie_fichaje') return guardarSelfieFichaje(data);

    if (data.action === 'guardar_telegram_id') {
      let sheet = ss.getSheetByName('Empleados');
      if (!sheet) return JSON.stringify({ ok: false, error: 'No existe hoja Empleados' });
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          sheet.getRange(i + 1, 23).setValue(String(data.telegramId));
          return JSON.stringify({ ok: true });
        }
      }
      return JSON.stringify({ ok: false, error: 'Empleado no encontrado' });
    }

    if (data.action === 'get_empleado_por_telegram_id') {
      let sheet = ss.getSheetByName('Empleados');
      if (!sheet) return JSON.stringify({ ok: false, error: 'No existe hoja Empleados' });
      const tz = Session.getScriptTimeZone() || 'America/Cancun';
      const fmtFecha = v => v instanceof Date ? Utilities.formatDate(v, tz, 'dd/MM/yyyy') : String(v || '');
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][22] || '') === String(data.telegramId)) {
          const r = rows[i];
          return JSON.stringify({ ok: true, empleado: {
            id: String(r[0]), name: String(r[1]), apellido: String(r[2] || ''),
            genero: String(r[3] || ''), puesto: String(r[4] || ''), dept: String(r[5] || ''),
            sede: String(r[6] || ''), fechaIngreso: fmtFecha(r[7]),
            telefono: String(r[8] || ''), email: String(r[9] || ''),
            foto: String(r[17] || ''), estado: String(r[18] || 'activo'),
            fechaNacimiento: fmtFecha(r[20]), sinGPS: String(r[21] || '').toLowerCase() === 'si',
          }});
        }
      }
      return JSON.stringify({ ok: false, error: 'No encontrado' });
    }

    if (data.action === 'get_homeoffice') {
      let sheet = ss.getSheetByName('HomeOffice');
      if (!sheet) return JSON.stringify({ ok: true, dias: [] });
      const rows = sheet.getDataRange().getValues();
      const dias = rows.slice(1).filter(r => r[0] && r[1]).map(r => ({
        empId: String(r[0]), fecha: String(r[1]),
      }));
      return JSON.stringify({ ok: true, dias });
    }

    if (data.action === 'toggle_homeoffice') {
      let sheet = ss.getSheetByName('HomeOffice');
      if (!sheet) {
        sheet = ss.insertSheet('HomeOffice');
        sheet.appendRow(['ID Empleado', 'Fecha']);
      }
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.empId) && String(rows[i][1]) === String(data.fecha)) {
          sheet.deleteRow(i + 1);
          return JSON.stringify({ ok: true, activo: false });
        }
      }
      sheet.appendRow([data.empId, data.fecha]);
      return JSON.stringify({ ok: true, activo: true });
    }

    if (data.action === 'marcar_homeoffice_hoy') {
      let sheet = ss.getSheetByName('HomeOffice');
      if (!sheet) {
        sheet = ss.insertSheet('HomeOffice');
        sheet.appendRow(['ID Empleado', 'Fecha']);
      }
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.empId) && String(rows[i][1]) === String(data.fecha)) {
          return JSON.stringify({ ok: true, yaEstaba: true });
        }
      }
      sheet.appendRow([data.empId, data.fecha]);
      return JSON.stringify({ ok: true, yaEstaba: false });
    }

    if (data.action === 'guardar_homeoffice_lote') {
      let sheet = ss.getSheetByName('HomeOffice');
      if (!sheet) {
        sheet = ss.insertSheet('HomeOffice');
        sheet.appendRow(['ID Empleado', 'Fecha']);
      }
      let cambios = [];
      try { cambios = JSON.parse(data.cambios); } catch (e) { return JSON.stringify({ ok: false, error: 'Datos invalidos' }); }

      const rows = sheet.getDataRange().getValues();
      const filasABorrar = [];
      const fechasExistentes = new Set();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.empId)) fechasExistentes.add(String(rows[i][1]));
      }

      const aAgregar = [];
      for (const c of cambios) {
        if (c.accion === 'add' && !fechasExistentes.has(c.fecha)) {
          aAgregar.push(c.fecha);
        } else if (c.accion === 'remove') {
          for (let i = 1; i < rows.length; i++) {
            if (String(rows[i][0]) === String(data.empId) && String(rows[i][1]) === c.fecha) {
              filasABorrar.push(i + 1);
            }
          }
        }
      }

      // Borrar de mayor a menor indice para no correr las filas restantes
      filasABorrar.sort((a, b) => b - a).forEach(fila => sheet.deleteRow(fila));
      aAgregar.forEach(fecha => sheet.appendRow([data.empId, fecha]));

      return JSON.stringify({ ok: true, agregados: aAgregar.length, quitados: filasABorrar.length });
    }

    if (data.action === 'enviar_credenciales') {
      let sheet = ss.getSheetByName('Empleados');
      if (!sheet) return JSON.stringify({ok:false, error:'No existe hoja Empleados'});
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          const emp = {
            id: String(rows[i][0]), nombre: String(rows[i][1]), apellido: String(rows[i][2]||''),
            email: String(rows[i][9]||''), pin: String(rows[i][15]||'0000'),
          };
          if (!emp.email) return JSON.stringify({ok:false, error:'Este empleado no tiene correo registrado'});
          enviarCorreoCredenciales(emp);
          return JSON.stringify({ok:true});
        }
      }
      return JSON.stringify({ok:false, error:'Empleado no encontrado'});
    }

    return JSON.stringify({ok:false, error:'Accion desconocida'});
  } catch(err) {
    return JSON.stringify({ok:false, error:err.message});
  }
}

function limpiarFichajes() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Fichajes');
  if (!sheet) { Logger.log('No existe Fichajes'); return; }
  const rows = sheet.getDataRange().getValues();
  const limpios = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]).toLowerCase().includes('empleado') || String(r[1]).toLowerCase().includes('departamento')) continue;
    let id='',nombre='',dept='',tipo='',hora='',fecha='',ts='',foto='',ubicacion='',maps='',canal='';
    if (String(r[0]).length > 8 && isNaN(Number(r[1])) && r[1] !== '') {
      id=String(r[0]);nombre=String(r[1]);dept=String(r[2]);tipo=String(r[3]);
      hora=String(r[4]);fecha=String(r[5]);ts=String(r[6]);foto=String(r[7]||'');
      ubicacion=String(r[8]||'');maps=String(r[9]||'');canal=String(r[10]||'');
    } else if (r[0] !== '' && isNaN(Number(r[0]))) {
      id='';nombre=String(r[0]);dept=String(r[1]);tipo=String(r[2]);
      hora=String(r[3]);fecha=String(r[4]);ts=String(r[5]);foto='';ubicacion='';maps='';canal='';
    } else { continue; }
    if (nombre && tipo && hora) limpios.push([id,nombre,dept,tipo,hora,fecha,ts,foto,ubicacion,maps,canal]);
  }
  sheet.clearContents();
  sheet.appendRow(['ID Empleado','Empleado','Departamento','Tipo','Hora','Fecha','Timestamp','Foto','Ubicacion','Maps','Canal']);
  limpios.forEach(r => sheet.appendRow(r));
  Logger.log('Listo. Filas: ' + limpios.length);
}

// ─────────────────────────────────────────────
// 🎂 CUMPLEAÑOS — enviar saludo por correo
// ─────────────────────────────────────────────

// Mientras testeas, usa sistemas@hostelche.com.mx.
// Cuando quede validado, cambia a grupoche@hostelche.com.mx
const CUMPLE_DESTINATARIOS = ['grupoche@hostelche.com.mx'];

function revisarCumpleanos() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Empleados');
  if (!sheet) return;

  const rows = sheet.getDataRange().getValues();
  const hoy = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Cancun';
  const hoyDia = Number(Utilities.formatDate(hoy, tz, 'd'));
  const hoyMes = Number(Utilities.formatDate(hoy, tz, 'M'));

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const id = r[0];
    if (!id) continue;

    const estado = String(r[18] || '').toLowerCase();
    if (estado && estado !== 'activo') continue;

    const fechaNac = r[20];
    if (!fechaNac) continue;

    let dia, mes;
    if (fechaNac instanceof Date) {
      dia = fechaNac.getDate();
      mes = fechaNac.getMonth() + 1;
    } else {
      const partes = String(fechaNac).split(/[\/\-]/).map(Number);
      if (partes.length < 3) continue;
      if (partes[0] > 31) {
        mes = partes[1]; dia = partes[2];
      } else {
        dia = partes[0]; mes = partes[1];
      }
    }

    if (dia === hoyDia && mes === hoyMes) {
      enviarCorreoCumpleanos({
        nombre: String(r[1] || ''),
        apellido: String(r[2] || ''),
        puesto: String(r[4] || ''),
        sede: String(r[6] || ''),
        foto: String(r[17] || ''),
      });
    }
  }
}

function enviarCorreoCumpleanos(emp) {
  const nombreCompleto = (emp.nombre + ' ' + emp.apellido).trim();
  const fotoHtml = emp.foto && emp.foto.startsWith('http')
    ? `<img src="${emp.foto}" width="120" height="120" style="border-radius:50%;object-fit:cover;display:block;margin:0 auto 20px;border:4px solid #1D9E75;" />`
    : `<div style="width:120px;height:120px;border-radius:50%;background:#f0ede8;margin:0 auto 20px;display:flex;align-items:center;justify-content:center;font-size:40px;color:#1D9E75;font-weight:600;">${(emp.nombre[0]||'')}${(emp.apellido[0]||'')}</div>`;

  const html = `
  <div style="background:#f5f4f0;padding:40px 20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;">
    <div style="max-width:420px;margin:0 auto;background:#ffffff;border-radius:20px;padding:36px 28px;border:1px solid #e8e6e0;text-align:center;">
      <div style="font-size:11px;font-weight:600;color:#aaaaaa;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:24px;">Hostel Che</div>
      <div style="font-size:40px;margin-bottom:8px;">&#127881;&#127874;&#127880;</div>
      ${fotoHtml}
      <div style="font-size:22px;font-weight:600;color:#1a1a1a;margin-bottom:4px;">¡Feliz cumpleaños, ${emp.nombre}!</div>
      <div style="font-size:14px;color:#aaaaaa;margin-bottom:20px;">${emp.puesto || ''}${emp.puesto && emp.sede ? ' · ' : ''}${emp.sede || ''}</div>
      <div style="font-size:15px;color:#444444;line-height:1.6;margin-bottom:8px;">
        Todo el equipo de <strong>Hostel Che</strong> te desea un día increíble.<br/>
        Gracias por ser parte de esta familia. &#127796;
      </div>
      <div style="margin-top:28px;padding-top:20px;border-top:1px solid #f0ede8;font-size:12px;color:#cccccc;">
        ${nombreCompleto}
      </div>
    </div>
  </div>`;

  GmailApp.sendEmail(CUMPLE_DESTINATARIOS.join(','), `Hoy cumple años ${nombreCompleto} - Hostel Che`, '', {
    htmlBody: html,
    from: 'rrhh@hostelche.com.mx',
    name: 'Hostel Che',
  });
}

// ─────────────────────────────────────────────
// 🎉 ANIVERSARIO LABORAL — enviar saludo por correo
// ─────────────────────────────────────────────

// Mientras testeas, usa maximiliano@hostelche.com.mx.
// Cuando quede validado, cambia a grupoche@hostelche.com.mx
const ANIVERSARIO_DESTINATARIOS = ['grupoche@hostelche.com.mx'];

function revisarAniversarios() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Empleados');
  if (!sheet) return;

  const rows = sheet.getDataRange().getValues();
  const hoy = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Cancun';
  const hoyDia = Number(Utilities.formatDate(hoy, tz, 'd'));
  const hoyMes = Number(Utilities.formatDate(hoy, tz, 'M'));
  const hoyAnio = Number(Utilities.formatDate(hoy, tz, 'yyyy'));

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const id = r[0];
    if (!id) continue;

    const estado = String(r[18] || '').toLowerCase();
    if (estado && estado !== 'activo') continue;

    const fechaIngreso = r[7];
    if (!fechaIngreso) continue;

    let dia, mes, anio;
    if (fechaIngreso instanceof Date) {
      dia = fechaIngreso.getDate();
      mes = fechaIngreso.getMonth() + 1;
      anio = fechaIngreso.getFullYear();
    } else {
      const partes = String(fechaIngreso).split(/[\/\-]/).map(Number);
      if (partes.length < 3) continue;
      if (partes[0] > 31) {
        anio = partes[0]; mes = partes[1]; dia = partes[2];
      } else if (partes[2] > 31) {
        dia = partes[0]; mes = partes[1]; anio = partes[2];
      } else {
        continue;
      }
    }

    const anios = hoyAnio - anio;
    if (dia === hoyDia && mes === hoyMes && anios >= 1) {
      enviarCorreoAniversario({
        nombre: String(r[1] || ''),
        apellido: String(r[2] || ''),
        puesto: String(r[4] || ''),
        sede: String(r[6] || ''),
        foto: String(r[17] || ''),
        anios: anios,
      });
    }
  }
}

function enviarCorreoAniversario(emp) {
  const nombreCompleto = (emp.nombre + ' ' + emp.apellido).trim();
  const fotoHtml = emp.foto && emp.foto.startsWith('http')
    ? `<img src="${emp.foto}" width="120" height="120" style="border-radius:50%;object-fit:cover;display:block;margin:0 auto 20px;border:4px solid #1D9E75;" />`
    : `<div style="width:120px;height:120px;border-radius:50%;background:#f0ede8;margin:0 auto 20px;display:flex;align-items:center;justify-content:center;font-size:40px;color:#1D9E75;font-weight:600;">${(emp.nombre[0]||'')}${(emp.apellido[0]||'')}</div>`;

  const textoAnios = emp.anios === 1 ? '1 año' : `${emp.anios} años`;

  const html = `
  <div style="background:#f5f4f0;padding:40px 20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;">
    <div style="max-width:420px;margin:0 auto;background:#ffffff;border-radius:20px;padding:36px 28px;border:1px solid #e8e6e0;text-align:center;">
      <div style="font-size:11px;font-weight:600;color:#aaaaaa;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:24px;">Hostel Che</div>
      <div style="font-size:40px;margin-bottom:8px;">&#127881;&#127942;&#127880;</div>
      ${fotoHtml}
      <div style="display:inline-block;padding:5px 16px;border-radius:20px;background:#1D9E75;color:#ffffff;font-size:13px;font-weight:600;margin-bottom:12px;">${textoAnios} en Hostel Che</div>
      <div style="font-size:22px;font-weight:600;color:#1a1a1a;margin-bottom:4px;">¡Feliz aniversario, ${emp.nombre}!</div>
      <div style="font-size:14px;color:#aaaaaa;margin-bottom:20px;">${emp.puesto || ''}${emp.puesto && emp.sede ? ' · ' : ''}${emp.sede || ''}</div>
      <div style="font-size:15px;color:#444444;line-height:1.6;margin-bottom:8px;">
        Gracias por ${textoAnios} de dedicacion y por ser parte fundamental de <strong>Hostel Che</strong>.<br/>
        Celebramos este camino juntos. &#127796;
      </div>
      <div style="margin-top:28px;padding-top:20px;border-top:1px solid #f0ede8;font-size:12px;color:#cccccc;">
        ${nombreCompleto}
      </div>
    </div>
  </div>`;

  GmailApp.sendEmail(ANIVERSARIO_DESTINATARIOS.join(','), `Hoy cumple ${emp.anios === 1 ? '1 año' : emp.anios + ' años'} en Hostel Che ${nombreCompleto}`, '', {
    htmlBody: html,
    from: 'rrhh@hostelche.com.mx',
    name: 'Hostel Che',
  });
}

// ─────────────────────────────────────────────
// 🔑 CREDENCIALES — enviar link + PIN por correo
// ─────────────────────────────────────────────

function enviarCorreoCredenciales(emp) {
  const link = `${FICHAR_BASE_URL}?id=${encodeURIComponent(emp.id)}`;
  const nombreCompleto = (emp.nombre + ' ' + (emp.apellido||'')).trim();

  const html = `
  <div style="background:#f5f4f0;padding:40px 20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;">
    <div style="max-width:440px;margin:0 auto;background:#ffffff;border-radius:20px;padding:36px 28px;border:1px solid #e8e6e0;text-align:center;">
      <div style="font-size:11px;font-weight:600;color:#aaaaaa;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:20px;">Hostel Che &middot; Fichaje</div>
      <div style="font-size:20px;font-weight:600;color:#1a1a1a;margin-bottom:6px;">Hola, ${emp.nombre}</div>
      <div style="font-size:14px;color:#888888;line-height:1.5;margin-bottom:24px;">
        Este es tu acceso personal para registrar tu entrada y salida en Hostel Che.
      </div>

      <a href="${link}" style="display:inline-block;background:#1D9E75;color:#ffffff;text-decoration:none;font-weight:600;font-size:15px;padding:14px 28px;border-radius:14px;margin-bottom:20px;">Abrir mi fichaje</a>

      <div style="font-size:12px;color:#aaaaaa;word-break:break-all;margin-bottom:24px;">${link}</div>

      <div style="background:#f5f4f0;border-radius:14px;padding:18px;margin-bottom:20px;">
        <div style="font-size:11px;color:#aaaaaa;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px;">Tu PIN de acceso</div>
        <div style="font-size:28px;font-weight:700;color:#1a1a1a;letter-spacing:0.15em;">${emp.pin}</div>
      </div>

      <div style="text-align:left;border-top:1px solid #f0ede8;padding-top:20px;margin-top:4px;">
        <div style="font-size:11px;font-weight:600;color:#aaaaaa;letter-spacing:0.08em;text-transform:uppercase;margin-bottom:12px;">Como funciona</div>
        <div style="font-size:13px;color:#444444;line-height:1.6;">
          <div style="margin-bottom:10px;"><strong>1.</strong> Abre tu link personal e ingresa tu PIN de 4 digitos.</div>
          <div style="margin-bottom:10px;"><strong>2.</strong> La primera vez, te vamos a pedir una foto tuya de frente (con la camara). Esa foto se usa despues para verificar tu identidad cada vez que fiches.</div>
          <div style="margin-bottom:10px;"><strong>3.</strong> Para fichar entrada o salida, toma una selfie y confirma tu ubicacion (debes estar dentro de tu sede).</div>
          <div style="margin-bottom:0;"><strong>4.</strong> En la pestana "Mi perfil" puedes ver y actualizar tus datos cuando quieras.</div>
        </div>
      </div>

      <div style="font-size:12px;color:#cccccc;margin-top:20px;">
        Guarda este correo. Si pierdes tu PIN, pide a IT que te lo reenvie.
      </div>
    </div>
  </div>`;

  GmailApp.sendEmail(emp.email, `Tu acceso a fichaje - Hostel Che`, '', {
    htmlBody: html,
    from: 'rrhh@hostelche.com.mx',
    name: 'Hostel Che',
  });
}

// ─────────────────────────────────────────────
// 📸 RECONOCIMIENTO FACIAL — respaldo de fotos en Drive
// ─────────────────────────────────────────────

function obtenerCarpetaEmpleado(empId, nombre, apellido) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Empleados');
    if (sheet) {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(empId)) {
          const driveUrl = String(rows[i][19] || '');
          const match = driveUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
          if (match) {
            try { return DriveApp.getFolderById(match[1]); } catch (e) {}
          }
          break;
        }
      }
    }
  } catch (e) {}
  const carpetaRaiz = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const nombreCarpeta = `${nombre} ${apellido || ''} [${empId}]`.trim();
  const existing = carpetaRaiz.getFoldersByName(nombreCarpeta);
  if (existing.hasNext()) return existing.next();
  return carpetaRaiz.createFolder(nombreCarpeta);
}

function obtenerSubcarpeta(carpetaPadre, nombreSub) {
  const existing = carpetaPadre.getFoldersByName(nombreSub);
  if (existing.hasNext()) return existing.next();
  return carpetaPadre.createFolder(nombreSub);
}

function decodificarImagenB64(b64raw) {
  const data = decodeURIComponent(b64raw);
  const match = data.match(/^data:([^;]+);base64,(.*)$/);
  if (!match) return null;
  return { mimeType: match[1], bytes: Utilities.base64Decode(match[2]) };
}

function guardarFotoPerfilDrive(data) {
  try {
    const carpeta = obtenerCarpetaEmpleado(data.id, data.nombre, data.apellido);
    const img = decodificarImagenB64(data.b64);
    if (!img) return JSON.stringify({ ok: false, error: 'Formato de imagen invalido' });
    const blob = Utilities.newBlob(img.bytes, img.mimeType, 'foto_perfil.jpg');
    const existentes = carpeta.getFilesByName('foto_perfil.jpg');
    while (existentes.hasNext()) existentes.next().setTrashed(true);
    const file = carpeta.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return JSON.stringify({ ok: true, url: file.getUrl() });
  } catch (err) {
    return JSON.stringify({ ok: false, error: err.message });
  }
}

function guardarSelfieFichaje(data) {
  try {
    const carpetaEmp = obtenerCarpetaEmpleado(data.empId, data.nombre, data.apellido);
    const carpetaSelfies = obtenerSubcarpeta(carpetaEmp, 'Selfies fichaje');
    const img = decodificarImagenB64(data.b64);
    if (!img) return JSON.stringify({ ok: false, error: 'Formato de imagen invalido' });
    const fechaSafe = String(data.fecha || '').replace(/\//g, '-');
    const horaSafe = String(data.hora || '').replace(/:/g, '-');
    const nombreArchivo = `fichaje_${fechaSafe}_${horaSafe}_${data.tipo || ''}.jpg`;
    const blob = Utilities.newBlob(img.bytes, img.mimeType, nombreArchivo);
    carpetaSelfies.createFile(blob);
    limpiarSelfiesDeCarpeta(carpetaSelfies);
    return JSON.stringify({ ok: true });
  } catch (err) {
    return JSON.stringify({ ok: false, error: err.message });
  }
}

function limpiarSelfiesDeCarpeta(carpeta) {
  const limite = new Date();
  limite.setDate(limite.getDate() - 15);
  const files = carpeta.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getDateCreated() < limite) f.setTrashed(true);
  }
}

// Funcion para el trigger diario (revisa TODAS las carpetas de empleados)
function limpiarSelfiesAntiguas() {
  const raiz = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const carpetas = raiz.getFolders();
  while (carpetas.hasNext()) {
    const empCarpeta = carpetas.next();
    const subs = empCarpeta.getFoldersByName('Selfies fichaje');
    if (subs.hasNext()) {
      limpiarSelfiesDeCarpeta(subs.next());
    }
  }
}

// ─────────────────────────────────────────────
// ⏱️ AUTO-CIERRE DE TURNOS OLVIDADOS (12 horas)
// ─────────────────────────────────────────────
// Revisa cada hora si alguien ficho entrada y nunca marco salida.
// Si ya pasaron 12 horas, registra una salida automatica (para que
// el reloj deje de correr) y le manda un correo avisando.

const LIMITE_HORAS_TURNO = 12;

function revisarSalidasOlvidadas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetFichajes = ss.getSheetByName('Fichajes');
  const sheetEmpleados = ss.getSheetByName('Empleados');
  if (!sheetFichajes || !sheetEmpleados) return;

  const ahora = Date.now();
  const limiteMs = LIMITE_HORAS_TURNO * 60 * 60 * 1000;
  const tz = Session.getScriptTimeZone() || 'America/Cancun';

  const fichajes = sheetFichajes.getDataRange().getValues();
  const empleadosRows = sheetEmpleados.getDataRange().getValues();

  // Ultimo fichaje por empleado (buscando hacia atras en las ultimas 48h)
  const ultimoPorEmpleado = {};
  for (let i = fichajes.length - 1; i >= 1; i--) {
    const r = fichajes[i];
    const empId = String(r[0]);
    const ts = Number(r[6]);
    if (!empId || !ts) continue;
    if (ahora - ts > 48 * 60 * 60 * 1000) continue;
    if (!ultimoPorEmpleado[empId] || ts > ultimoPorEmpleado[empId].ts) {
      ultimoPorEmpleado[empId] = { tipo: String(r[3]), ts: ts, nombre: String(r[1]), dept: String(r[2]) };
    }
  }

  for (const empId in ultimoPorEmpleado) {
    const ultimo = ultimoPorEmpleado[empId];
    if (ultimo.tipo !== 'entrada') continue;
    if (ahora - ultimo.ts < limiteMs) continue;

    // Ya pasaron 12+ horas desde la entrada sin salida: auto-cerrar
    const tsCierre = ultimo.ts + limiteMs;
    const fechaCierre = new Date(tsCierre);
    const horaStr = Utilities.formatDate(fechaCierre, tz, 'HH:mm');
    const fechaStr = Utilities.formatDate(fechaCierre, tz, 'dd/MM/yyyy');

    sheetFichajes.appendRow([empId, ultimo.nombre, ultimo.dept, 'salida', horaStr, fechaStr, tsCierre, 'No', 'Auto-cierre (turno olvidado)', '', 'Sistema']);

    // Buscar email del empleado para avisarle
    for (let i = 1; i < empleadosRows.length; i++) {
      if (String(empleadosRows[i][0]) === empId) {
        const email = String(empleadosRows[i][9] || '');
        if (email) {
          const horaEntrada = Utilities.formatDate(new Date(ultimo.ts), tz, 'HH:mm');
          const fechaEntrada = Utilities.formatDate(new Date(ultimo.ts), tz, 'dd/MM/yyyy');
          enviarCorreoTurnoOlvidado(email, ultimo.nombre, fechaEntrada, horaEntrada, horaStr);
        }
        break;
      }
    }
  }
}

function enviarCorreoTurnoOlvidado(email, nombre, fechaEntrada, horaEntrada, horaCierre) {
  const html = `
  <div style="background:#f5f4f0;padding:40px 20px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;">
    <div style="max-width:440px;margin:0 auto;background:#ffffff;border-radius:20px;padding:32px 28px;border:1px solid #e8e6e0;">
      <div style="font-size:11px;font-weight:600;color:#aaaaaa;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:20px;">Hostel Che &middot; Fichaje</div>
      <div style="font-size:18px;font-weight:600;color:#1a1a1a;margin-bottom:12px;">Hola ${nombre}, olvidaste marcar tu salida</div>
      <div style="font-size:14px;color:#444444;line-height:1.6;background:#f5f4f0;padding:14px;border-radius:10px;margin-bottom:16px;">
        Registraste tu entrada el <strong>${fechaEntrada}</strong> a las <strong>${horaEntrada}</strong>, pero nunca marcaste tu salida.
        <br/><br/>
        Tu turno se cerro automaticamente a las <strong>${horaCierre}</strong> (${LIMITE_HORAS_TURNO} horas despues de tu entrada), para que el sistema no siga contando tus horas de mas.
      </div>
      <div style="font-size:12px;color:#cccccc;">Recuerda marcar tu salida la proxima vez desde tu link de fichaje o el bot de Telegram.</div>
    </div>
  </div>`;

  GmailApp.sendEmail(email, `Olvidaste marcar tu salida - Hostel Che`, '', {
    htmlBody: html,
    from: 'rrhh@hostelche.com.mx',
    name: 'Hostel Che',
  });
}
