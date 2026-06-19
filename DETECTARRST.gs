/**
 * =====================================================================
 * DETECTARRST.gs — v1 — detecta RST en Drive y marca tickets resueltos
 * =====================================================================
 *
 * Flujo:
 *   1. Cada 5 min revisa la carpeta de Drive (FOLDER_RST_ID)
 *   2. Lista PDFs nuevos (no procesados)
 *   3. Intenta vincular cada PDF con un ticket:
 *      a) Por TCKxxx en el nombre del archivo
 *      b) Por fecha (DDMMYY en nombre) + técnico (dueño del archivo)
 *      c) Por fecha solamente (si solo hay 1 ticket CONFIRMADO ese día)
 *   4. Marca ticket como RESUELTO y notifica a vendedor + jefatura
 *   5. Mueve el PDF a subcarpeta "Procesados" para no duplicar
 */

// Carpeta única donde los técnicos suben TODOS los RST
const FOLDER_RST_ID = '1Mgd9AmfDnsn-eGlQWgiOmm-znTdZcNDg';

function procesarRSTsNuevos() {
  console.log('=== Revisando carpeta de RST ===');

  let folder;
  try {
    folder = DriveApp.getFolderById(FOLDER_RST_ID);
    console.log(`Carpeta: "${folder.getName()}"`);
  } catch(e) {
    log_('ERROR', `No se pudo abrir carpeta RST: ${e.message}`);
    return 0;
  }

  // Recolectar todos los PDFs de la carpeta
  const pdfs = folder.getFilesByType(MimeType.PDF);
  const archivosPorRevisar = [];
  while (pdfs.hasNext()) {
    archivosPorRevisar.push(pdfs.next());
  }

  console.log(`Total de PDFs encontrados: ${archivosPorRevisar.length}`);

  let procesados = 0;
  let fallidos = 0;
  let saltadosProcesados = 0;

  for (const archivo of archivosPorRevisar) {
    const nombre = archivo.getName();

    // Saltar si ya está marcado como procesado
    if (/\[PROCESADO\]/.test(nombre)) {
      saltadosProcesados++;
      continue;
    }

    console.log(`\nAnalizando: ${nombre}`);

    try {
      const resultado = vincularRSTaTicket_(archivo);
      if (resultado.exitoso) {
        marcarTicketResuelto_(resultado.ticketId, archivo, resultado.resumen);
        // Renombrar con [PROCESADO] para no reprocesar
        try {
          const nuevoNombre = nombre.replace(/\.pdf$/i, '') + ' [PROCESADO].pdf';
          archivo.setName(nuevoNombre);
        } catch(e) {
          log_('WARN', `No se pudo renombrar ${nombre}: ${e.message}`);
        }
        procesados++;
      } else {
        console.log(`  No vinculado: ${resultado.motivo}`);
        notificarRSTNoVinculado_(archivo, resultado.motivo);
        fallidos++;
      }
    } catch(err) {
      log_('ERROR', `RST ${nombre}: ${err.message}`);
      fallidos++;
    }
  }

  console.log(`\n=== RESUMEN ===`);
  console.log(`Total PDFs: ${archivosPorRevisar.length}`);
  console.log(`  - Ya procesados (saltados): ${saltadosProcesados}`);
  console.log(`  - Vinculados exitosamente: ${procesados}`);
  console.log(`  - No vinculados: ${fallidos}`);

  log_('INFO', `RST: ${procesados} procesados | ${fallidos} no vinculados | ${saltadosProcesados} ya procesados antes`);

  return procesados;
}

// ---------- Intentar vincular PDF con ticket ----------
function vincularRSTaTicket_(archivo) {
  const nombre = archivo.getName();

  // Estrategia A: TCKxxx en el nombre
  const matchTCK = nombre.match(/TCK\d+/);
  if (matchTCK) {
    const ticketId = matchTCK[0];
    const ctx = buscarTicketPorId_(ticketId);
    if (ctx) {
      console.log(`  Estrategia A: match por Ticket_ID = ${ticketId}`);
      return { exitoso: true, ticketId, resumen: 'Vinculado por Ticket_ID en nombre' };
    }
    return { exitoso: false, motivo: `Se encontró ${ticketId} en el nombre pero no existe en la Sheet` };
  }

  // Estrategia B: fecha en nombre + dueño del archivo como técnico
  const fechaArchivo = extraerFechaDelNombre_(nombre) || archivo.getDateCreated();
  const dueñoEmail = archivo.getOwner() ? archivo.getOwner().getEmail() : null;

  if (dueñoEmail) {
    const ctx = buscarTicketDeTecnicoEnFecha_(dueñoEmail, fechaArchivo);
    if (ctx === 'multiple') {
      return { exitoso: false, motivo: `Múltiples tickets de ${dueñoEmail} en esa fecha` };
    }
    if (ctx) {
      console.log(`  Estrategia B: match por técnico ${dueñoEmail} + fecha`);
      return { exitoso: true, ticketId: ctx.ticketId, resumen: 'Vinculado por técnico + fecha' };
    }
  }

  // Estrategia C: fecha solamente (último recurso)
  const ctxC = buscarTicketEnFecha_(fechaArchivo);
  if (ctxC === 'multiple') {
    return { exitoso: false, motivo: `Múltiples tickets CONFIRMADOS en esa fecha` };
  }
  if (ctxC) {
    console.log(`  Estrategia C: match por fecha única`);
    return { exitoso: true, ticketId: ctxC.ticketId, resumen: 'Vinculado por fecha única' };
  }

  return { exitoso: false, motivo: 'No se pudo vincular con ningún ticket' };
}

// ---------- Extraer fecha del nombre del archivo ----------
/**
 * Reconoce patrones como:
 *   RST-2604241041.pdf  → DDMMYY = 26/04/24
 *   RST-260424.pdf      → 26/04/24
 *   2026-04-26          → 26/04/26
 */
function extraerFechaDelNombre_(nombre) {
  // Patrón YYYY-MM-DD
  let m = nombre.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  }

  // Patrón DDMMYY (6 dígitos) posiblemente seguido de HHMM
  m = nombre.match(/(\d{2})(\d{2})(\d{2})(\d{4})?/);
  if (m) {
    const dd = parseInt(m[1]);
    const mm = parseInt(m[2]);
    const yy = parseInt(m[3]);
    // Asumir 20YY (año 2000+)
    const year = yy < 70 ? 2000 + yy : 1900 + yy;
    if (dd >= 1 && dd <= 31 && mm >= 1 && mm <= 12) {
      return new Date(year, mm - 1, dd);
    }
  }

  return null;
}

// ---------- Buscar ticket de un técnico en una fecha ----------
function buscarTicketDeTecnicoEnFecha_(emailTecnico, fecha) {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TICKETS);
  const data = hoja.getDataRange().getValues();

  const tecnicos = cargarTecnicos_();
  const tec = tecnicos.find(t => t.email.toLowerCase() === emailTecnico.toLowerCase());
  if (!tec) return null;

  const fechaStr = Utilities.formatDate(fecha, CFG.TIMEZONE, 'yyyy-MM-dd');
  const base = hoja.getLastColumn() - 6;

  const matches = [];
  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    const ticketId = fila[base];
    const tecnicoNombre = fila[base + 1];
    const fechaSlot = fila[base + 3];
    const estado = String(fila[base + 5] || '').toUpperCase();

    if (estado !== 'CONFIRMADO') continue;
    if (tecnicoNombre !== tec.nombre) continue;

    const fechaSlotStr = fechaSlot instanceof Date
      ? Utilities.formatDate(fechaSlot, CFG.TIMEZONE, 'yyyy-MM-dd')
      : String(fechaSlot).substring(0, 10);

    if (fechaSlotStr === fechaStr) {
      matches.push({ ticketId, fila: i + 1 });
    }
  }

  if (matches.length === 0) return null;
  if (matches.length > 1) return 'multiple';
  return matches[0];
}

// ---------- Buscar ticket único en una fecha (cualquier técnico) ----------
function buscarTicketEnFecha_(fecha) {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TICKETS);
  const data = hoja.getDataRange().getValues();

  const fechaStr = Utilities.formatDate(fecha, CFG.TIMEZONE, 'yyyy-MM-dd');
  const base = hoja.getLastColumn() - 6;

  const matches = [];
  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    const ticketId = fila[base];
    const fechaSlot = fila[base + 3];
    const estado = String(fila[base + 5] || '').toUpperCase();

    if (estado !== 'CONFIRMADO') continue;

    const fechaSlotStr = fechaSlot instanceof Date
      ? Utilities.formatDate(fechaSlot, CFG.TIMEZONE, 'yyyy-MM-dd')
      : String(fechaSlot).substring(0, 10);

    if (fechaSlotStr === fechaStr) {
      matches.push({ ticketId, fila: i + 1 });
    }
  }

  if (matches.length === 0) return null;
  if (matches.length > 1) return 'multiple';
  return matches[0];
}

// ---------- Marcar ticket como RESUELTO ----------
function marcarTicketResuelto_(ticketId, archivoPDF, metodoVinculacion) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`Ticket ${ticketId} no encontrado`);

  const linkPDF = archivoPDF.getUrl();
  const nombrePDF = archivoPDF.getName();

  // Actualizar estado en Sheet
  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue('RESUELTO');

  // Cambiar color del evento en Calendar a azul oscuro
  if (ctx.evento) {
    try {
      ctx.evento.setColor(CalendarApp.EventColor.BLUE || '9');
      const desc = ctx.evento.getDescription() || '';
      ctx.evento.setDescription(desc + `\n\n[RESUELTO]\nPDF del RST: ${linkPDF}`);
    } catch(_) {}
  }

  log_('INFO', `Ticket ${ticketId} resuelto (${metodoVinculacion}). PDF: ${nombrePDF}`);

  // Notificar al vendedor
  if (ctx.emailVendedor) {
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[RESUELTO ${ticketId}] ${ctx.cliente}`,
      `Tu ticket ${ticketId} fue completado por el técnico.\n\n` +
      `-------------------------------------------------\n` +
      `Ticket:       ${ticketId}\n` +
      `Cliente:      ${ctx.cliente}\n` +
      `Equipo:       ${ctx.equipo}\n` +
      `Servicio:     ${ctx.tipoServicio}\n` +
      `Técnico:      ${ctx.tecnico ? ctx.tecnico.nombre : '(desconocido)'}\n` +
      `Fecha visita: ${ctx.fechaSlot}\n` +
      `-------------------------------------------------\n\n` +
      `Reporte RST (PDF): ${linkPDF}\n\n` +
      `Este ticket queda cerrado en el sistema.`
    );
  }

  // Notificar a jefatura
  GmailApp.sendEmail(
    CFG.EMAIL_JEFATURA,
    `[RESUELTO ${ticketId}] ${ctx.cliente} - RST cargado`,
    `El técnico ${ctx.tecnico ? ctx.tecnico.nombre : '(desconocido)'} completó el ticket.\n\n` +
    `-------------------------------------------------\n` +
    `Ticket:       ${ticketId}\n` +
    `Cliente:      ${ctx.cliente}\n` +
    `Equipo:       ${ctx.equipo}\n` +
    `Servicio:     ${ctx.tipoServicio}\n` +
    `Fecha visita: ${ctx.fechaSlot}\n` +
    `Vendedor:     ${ctx.vendedor}\n` +
    `-------------------------------------------------\n\n` +
    `Reporte RST (PDF): ${linkPDF}\n` +
    `Vinculación:  ${metodoVinculacion}\n\n` +
    `El ticket queda cerrado en el sistema.`
  );

  actualizarDashboard_();
}

// ---------- Notificar RST no vinculado ----------
function notificarRSTNoVinculado_(archivo, motivo) {
  log_('WARN', `RST no vinculado: ${archivo.getName()} — ${motivo}`);

  const archivoId = archivo.getId();
  enviarCorreoUnico_(
    CFG.EMAIL_JEFATURA,
    `[RST SIN VINCULAR] ${archivo.getName()}`,
    `Se detectó un nuevo RST en Drive pero no se pudo vincular automáticamente a un ticket.\n\n` +
    `Archivo: ${archivo.getName()}\n` +
    `Link: ${archivo.getUrl()}\n` +
    `Motivo: ${motivo}\n\n` +
    `Tienes 2 opciones:\n` +
    `1) Renombra el PDF para incluir el TCKxxx correspondiente (ej: RST-TCK202604261041.pdf) — el sistema lo procesará en el siguiente ciclo.\n` +
    `2) Marca el ticket como RESUELTO manualmente en la Sheet.\n`,
    archivoId,
    'rst_no_vinculado'
  );
}

// ---------- Mover archivo a subcarpeta Procesados ----------
function moverAProcesados_(archivo, folderProcesados) {
  try {
    // Agregar tag al nombre para evitar reprocesamiento si falla el move
    const nuevoNombre = archivo.getName().replace(/\.pdf$/i, '') + ' [PROCESADO].pdf';
    archivo.setName(nuevoNombre);
    folderProcesados.addFile(archivo);
    // Quitar de la carpeta original
    const padres = archivo.getParents();
    while (padres.hasNext()) {
      const p = padres.next();
      if (p.getId() !== folderProcesados.getId()) {
        p.removeFile(archivo);
      }
    }
  } catch(e) {
    log_('WARN', `No se pudo mover RST a procesados: ${e.message}`);
  }
}

function obtenerOCrearSubcarpeta_(folderPadre, nombre) {
  const existentes = folderPadre.getFoldersByName(nombre);
  if (existentes.hasNext()) return existentes.next();
  return folderPadre.createFolder(nombre);
}
