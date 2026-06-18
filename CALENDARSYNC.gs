/**
 * =====================================================================
 * CALENDARSYNC.gs — v2 — sin emojis
 * =====================================================================
 */

const SNAPSHOT_TAB = 'SnapshotCalendar';

function detectarCambiosEnCalendar() {
  const ui = SpreadsheetApp.getUi();
  console.log('=== Detectando cambios manuales en Calendar ===');

  try {
    const resultado = sincronizarConCalendar_();

    const msg =
      `Detección completa.\n\n` +
      `- ${resultado.cambiosHora} cambios de hora/duración\n` +
      `- ${resultado.cambiosTecnico} cambios de técnico\n` +
      `- ${resultado.cambiosTitulo} cambios de título\n` +
      `- ${resultado.borrados} eventos borrados (tickets cancelados)\n` +
      `- ${resultado.nuevos} eventos nuevos en snapshot\n` +
      `- ${resultado.sinCambio} sin cambios`;

    ui.alert('Sincronización Calendar', msg, ui.ButtonSet.OK);
    log_('INFO', `Sync Calendar: ${JSON.stringify(resultado)}`);
  } catch (err) {
    log_('ERROR', `detectarCambiosEnCalendar: ${err.message}\n${err.stack}`);
    ui.alert('Error', err.message, ui.ButtonSet.OK);
  }
}

function sincronizarConCalendar_() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  let snap = ss.getSheetByName(SNAPSHOT_TAB);
  if (!snap) {
    snap = ss.insertSheet(SNAPSHOT_TAB);
    snap.getRange(1, 1, 1, 7).setValues([[
      'Event_ID', 'Ticket_ID', 'Inicio', 'Duracion_Horas',
      'Invitados', 'Titulo', 'Ultima_Verificacion'
    ]]);
    snap.hideSheet();
  }

  const snapshotData = snap.getDataRange().getValues();
  const snapshotMap = {};
  for (let i = 1; i < snapshotData.length; i++) {
    const [eventId, ticketId, inicio, duracion, invitados, titulo] = snapshotData[i];
    if (eventId) {
      snapshotMap[eventId] = { ticketId, inicio, duracion, invitados, titulo, filaSnap: i + 1 };
    }
  }

  const hojaTickets = ss.getSheetByName(CFG.TAB_TICKETS);
  const data = hojaTickets.getDataRange().getValues();
  const headers = data[0].map(h => normalizarHeader_(h));
  const baseCol = hojaTickets.getLastColumn() - 6;

  const cal = getCalendarCentral_();
  const resultado = {
    cambiosHora: 0,
    cambiosTecnico: 0,
    cambiosTitulo: 0,
    borrados: 0,
    nuevos: 0,
    sinCambio: 0
  };

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    const estado = String(fila[baseCol + 5] || '').toUpperCase();
    const eventId = fila[baseCol + 4];
    const ticketId = fila[baseCol + 0];

    if (!eventId || !ticketId) continue;
    if (estado !== 'CONFIRMADO') continue;

    let evento = null;
    try {
      evento = cal.getEventById(eventId);
    } catch(_) { evento = null; }

    if (!evento) {
      const snapPrev = snapshotMap[eventId];
      if (snapPrev) {
        manejarEventoBorrado_(ticketId, i + 1, hojaTickets, fila, headers);
        snap.getRange(snapPrev.filaSnap, 1, 1, 7).clearContent();
        resultado.borrados++;
      }
      continue;
    }

    const actual = extraerEstadoEvento_(evento);
    const snapPrev = snapshotMap[eventId];

    if (!snapPrev) {
      escribirSnapshot_(snap, eventId, ticketId, actual);
      resultado.nuevos++;
      continue;
    }

    const cambios = compararEstados_(snapPrev, actual);
    if (cambios.length === 0) {
      resultado.sinCambio++;
      snap.getRange(snapPrev.filaSnap, 7).setValue(new Date());
      continue;
    }

    if (cambios.includes('hora') || cambios.includes('duracion')) resultado.cambiosHora++;
    if (cambios.includes('invitados')) resultado.cambiosTecnico++;
    if (cambios.includes('titulo')) resultado.cambiosTitulo++;

    manejarEventoCambiado_(evento, actual, snapPrev, cambios, ticketId, fila, headers, hojaTickets, i + 1);
    escribirSnapshot_(snap, eventId, ticketId, actual, snapPrev.filaSnap);
  }

  return resultado;
}

function extraerEstadoEvento_(evento) {
  const inicio = evento.getStartTime();
  const fin = evento.getEndTime();
  const duracion = (fin - inicio) / 3600000;

  const guests = evento.getGuestList()
    .map(g => g.getEmail().toLowerCase())
    .sort()
    .join(',');

  return {
    inicio,
    duracion,
    invitados: guests,
    titulo: evento.getTitle()
  };
}

function compararEstados_(prev, actual) {
  const cambios = [];

  const inicioPrev = new Date(prev.inicio);
  const inicioAct = actual.inicio;

  if (Math.abs(inicioPrev.getTime() - inicioAct.getTime()) > 60000) cambios.push('hora');

  const duracionPrev = parseFloat(prev.duracion);
  if (Math.abs(duracionPrev - actual.duracion) > 0.01) cambios.push('duracion');

  if (String(prev.invitados).trim() !== actual.invitados) cambios.push('invitados');
  if (String(prev.titulo).trim() !== actual.titulo) cambios.push('titulo');

  return cambios;
}

function escribirSnapshot_(snap, eventId, ticketId, estado, filaExistente) {
  const row = [
    eventId,
    ticketId,
    Utilities.formatDate(estado.inicio, CFG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss"),
    estado.duracion.toFixed(2),
    estado.invitados,
    estado.titulo,
    new Date()
  ];

  if (filaExistente) {
    snap.getRange(filaExistente, 1, 1, 7).setValues([row]);
  } else {
    snap.appendRow(row);
  }
}

function manejarEventoCambiado_(evento, actual, prev, cambios, ticketId, fila, headers, hojaTickets, filaNum) {
  log_('INFO', `Ticket ${ticketId} cambió: ${cambios.join(', ')}`);

  const t = {};
  headers.forEach((h, i) => t[h] = fila[i]);

  const cliente = t.cliente;
  const direccion = t.direccion || t.ubicacion;
  const equipo = t.equipo;
  const servicio = t.tipo_de_servicio || t.servicio;
  const emailVendedor = t.correo_vendedor || t.email_vendedor;

  const nuevoInicio = evento.getStartTime();
  const nuevaDuracion = actual.duracion;
  const cambiosTexto = cambios.join(', ');

  // Determinar técnico (puede haber cambiado)
  const emailTecnicoActual = actual.invitados.split(',')[0];
  let nombreTecnicoActual = '(desconocido)';
  if (emailTecnicoActual) {
    const tecnicos = cargarTecnicos_();
    const tec = tecnicos.find(tt => tt.email === emailTecnicoActual);
    if (tec) nombreTecnicoActual = tec.nombre;
  }

  // ----- Actualizar la Sheet con los cambios del Calendar -----
  if (hojaTickets && filaNum) {
    const totalCols = hojaTickets.getLastColumn();
    try {
      // Flag con timestamp: vale por 60 segundos
      const ahora = Date.now();
      PropertiesService.getScriptProperties().setProperty(
        'skipOnEdit_' + ticketId,
        String(ahora)
      );

      // Actualizar fecha si cambió hora (solo si el valor actual es diferente)
      if (cambios.includes('hora') || cambios.includes('duracion')) {
        const colFechaSlot = totalCols - 2;
        const nuevaFechaStr = Utilities.formatDate(nuevoInicio, CFG.TIMEZONE, 'yyyy-MM-dd');
        const valorActual = hojaTickets.getRange(filaNum, colFechaSlot).getValue();
        const valorActualStr = valorActual instanceof Date
          ? Utilities.formatDate(valorActual, CFG.TIMEZONE, 'yyyy-MM-dd')
          : String(valorActual).substring(0, 10);

        if (valorActualStr !== nuevaFechaStr) {
          hojaTickets.getRange(filaNum, colFechaSlot).setValue(nuevaFechaStr);
          log_('INFO', `Sheet actualizada: ${ticketId} fecha -> ${nuevaFechaStr}`);
        } else {
          log_('INFO', `Sheet ya tenía la fecha correcta de ${ticketId}, no se modifica`);
        }
      }

      // Actualizar técnico si cambió (solo si el valor actual es diferente)
      if (cambios.includes('invitados')) {
        const colTecnico = totalCols - 4;
        const valorActual = String(hojaTickets.getRange(filaNum, colTecnico).getValue() || '');
        if (valorActual !== nombreTecnicoActual) {
          hojaTickets.getRange(filaNum, colTecnico).setValue(nombreTecnicoActual);
          log_('INFO', `Sheet actualizada: ${ticketId} técnico -> ${nombreTecnicoActual}`);
        }
      }

      // NO borrar el flag - se invalida solo cuando expire (60s)
    } catch(e) {
      log_('WARN', `No se pudo actualizar Sheet de ${ticketId}: ${e.message}`);
    }
  }

  // ----- Construir cuerpo común -----
  const cuerpoBase = `
Se detectó una modificación manual al evento del ticket ${ticketId}.

-------------------------------------------------
Ticket:           ${ticketId}
Cliente:          ${cliente}
Dirección:        ${direccion}
Equipo:           ${equipo}
Servicio:         ${servicio}
Técnico actual:   ${nombreTecnicoActual}
-------------------------------------------------
Cambios detectados: ${cambiosTexto}
Nueva fecha y hora: ${Utilities.formatDate(nuevoInicio, CFG.TIMEZONE, 'EEEE dd MMM, HH:mm')}
Nueva duración:     ${formatearHorasMinutos_(nuevaDuracion)}
-------------------------------------------------
`;

  // ----- 1) Notificar al técnico -----
  if (emailTecnicoActual) {
    enviarCorreoUnico_(
      emailTecnicoActual,
      `[ACTUALIZACIÓN ${ticketId}] ${cliente} - evento modificado`,
      `Hola ${nombreTecnicoActual},\n` +
      cuerpoBase +
      (cambios.includes('invitados')
        ? '\nEste ticket fue reasignado a ti. Por favor confirma.'
        : '\nRevisa tu calendario para ver los detalles actualizados.'),
      ticketId,
      `cambio_${cambiosTexto}_tecnico_${Date.now()}`  // únique por timestamp - permite múltiples ediciones
    );
  }

  // ----- 2) Notificar al vendedor -----
  if (emailVendedor) {
    enviarCorreoUnico_(
      emailVendedor,
      `[ACTUALIZACIÓN ${ticketId}] ${cliente} - cambio en agenda`,
      `Hola,\n\nLa agenda del ticket ${ticketId} fue modificada.` +
      cuerpoBase +
      `\nSi tienes alguna observación, contacta a jefatura.`,
      ticketId,
      `cambio_${cambiosTexto}_vendedor_${Date.now()}`
    );
  }

  // ----- 3) Notificar a jefatura -----
  enviarCorreoUnico_(
    CFG.EMAIL_JEFATURA,
    `[ACTUALIZACIÓN ${ticketId}] ${cliente} - cambio manual en Calendar`,
    `Hola,\n\nUn evento de ticket fue editado manualmente en Calendar.` +
    cuerpoBase +
    `\nVendedor: ${t.vendedor || '(desconocido)'}` +
    (emailVendedor ? `\nVendedor avisado: sí` : '\nVendedor avisado: no (sin correo)'),
    ticketId,
    `cambio_${cambiosTexto}_jefatura_${Date.now()}`
  );

  // Reoptimizar ruta del día del técnico (si existe)
  if (emailTecnicoActual) {
    const tecnicos = cargarTecnicos_();
    const tec = tecnicos.find(tt => tt.email === emailTecnicoActual);
    if (tec) {
      try {
        reoptimizarRutaDelDia_(tec, evento.getStartTime());
      } catch(e) {
        log_('WARN', `Ruta no reoptimizada: ${e.message}`);
      }
    }
  }
}

function manejarEventoBorrado_(ticketId, filaTicket, hojaTickets, fila, headers) {
  log_('INFO', `Ticket ${ticketId} cancelado (evento borrado)`);

  hojaTickets.getRange(filaTicket, hojaTickets.getLastColumn())
    .setValue('CANCELADO_MANUALMENTE');

  const t = {};
  headers.forEach((h, i) => t[h] = fila[i]);
  const emailVendedor = t.correo_vendedor || t.email_vendedor;
  const cliente = t.cliente;
  const equipo = t.equipo;

  // Determinar técnico que tenía asignado (de las columnas auxiliares)
  const base = hojaTickets.getLastColumn() - 6;
  const tecnicoNombre = fila[base + 1];

  // Buscar email del técnico
  let emailTecnico = null;
  if (tecnicoNombre) {
    const tecnicos = cargarTecnicos_();
    const tec = tecnicos.find(tt => tt.nombre === tecnicoNombre);
    if (tec) emailTecnico = tec.email;
  }

  const cuerpoBase = `
El evento del ticket ${ticketId} fue eliminado manualmente del Calendar.

-------------------------------------------------
Ticket:    ${ticketId}
Cliente:   ${cliente}
Equipo:    ${equipo}
Técnico:   ${tecnicoNombre || '(desconocido)'}
-------------------------------------------------

El ticket queda marcado como CANCELADO_MANUALMENTE en la Sheet.
Si deseas reagendarlo, debes generar un nuevo ticket.
`;

  // 1) Vendedor
  if (emailVendedor) {
    enviarCorreoUnico_(
      emailVendedor,
      `[CANCELADO ${ticketId}] ${cliente}`,
      `Tu ticket ${ticketId} para ${cliente} fue cancelado manualmente.\n` + cuerpoBase,
      ticketId,
      'cancelado_vendedor'
    );
  }

  // 2) Técnico
  if (emailTecnico) {
    enviarCorreoUnico_(
      emailTecnico,
      `[CANCELADO ${ticketId}] ${cliente}`,
      `El evento del ticket ${ticketId} fue eliminado.\n` + cuerpoBase,
      ticketId,
      'cancelado_tecnico'
    );
  }

  // 3) Jefatura
  enviarCorreoUnico_(
    CFG.EMAIL_JEFATURA,
    `[CANCELADO ${ticketId}] ${cliente}`,
    `Un evento de ticket fue eliminado manualmente.\n` + cuerpoBase,
    ticketId,
    'cancelado_jefatura'
  );
}

// La función enviarNotificacionCambio_ fue reemplazada por la lógica
// integrada en manejarEventoCambiado_ que notifica a técnico+vendedor+jefatura.
