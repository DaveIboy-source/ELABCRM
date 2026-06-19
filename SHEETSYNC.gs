/**
 * =====================================================================
 * SHEETSYNC.gs — v1 — sincronización Sheet -> Calendar
 * =====================================================================
 *
 * Detecta cambios manuales en la Sheet "Solicitudes de Servicio Tecnico"
 * y los aplica al evento de Calendar correspondiente.
 *
 * Cambios soportados:
 *   - Estado a CANCELADO_MANUALMENTE -> borra evento + notifica
 *   - Fecha_Slot -> mueve fecha del evento
 *   - Cliente, Dirección, Equipo, Teléfono cliente, Notas -> actualizan
 *     título y/o descripción del evento
 *   - Si cambia Dirección -> re-geocodifica y recalcula traslado
 *
 * Setup:
 *   1. Pegar este archivo en Apps Script
 *   2. Triggers (icono reloj) -> Add Trigger
 *      Función: onEditSheet
 *      Evento: From spreadsheet -> On edit
 */

const COL_ESTADO_OFFSET = 0;     // última columna = Estado
const COL_EVENT_ID_OFFSET = -1;  // penúltima = Event_ID
const COL_FECHA_SLOT_OFFSET = -2;
const COL_TRASLADO_OFFSET = -3;
const COL_TECNICO_OFFSET = -4;
const COL_TICKET_ID_OFFSET = -5;

function onEditSheet(e) {
  // Log inicial — útil para detectar si el trigger se está disparando
  try {
    log_('INFO', `onEditSheet: trigger disparado (event=${e ? 'sí' : 'no'})`);
  } catch(_) {}

  if (!e || !e.range) {
    log_('WARN', 'onEditSheet: sin evento o sin range, salida temprana');
    return;
  }

  const hoja = e.range.getSheet();
  log_('INFO', `onEditSheet: edición en hoja "${hoja.getName()}"`);
  if (hoja.getName() !== CFG.TAB_TICKETS) {
    log_('INFO', `onEditSheet: ignorando hoja "${hoja.getName()}" (esperaba "${CFG.TAB_TICKETS}")`);
    return;
  }

  const fila = e.range.getRow();
  if (fila < 2) return;  // header

  const totalCols = hoja.getLastColumn();
  const colEditada = e.range.getColumn();

  try {
    // Identificar el ticket de esta fila
    const ticketId = hoja.getRange(fila, totalCols + COL_TICKET_ID_OFFSET).getValue();
    if (!ticketId) {
      log_('INFO', `onEditSheet: fila ${fila} no tiene ticketId, saliendo`);
      return;
    }

    // Si este cambio fue iniciado por CalendarSync, ignorar (evita loop infinito)
    // El flag tiene timestamp y vale por 60 segundos
    const skipFlag = PropertiesService.getScriptProperties().getProperty('skipOnEdit_' + ticketId);
    if (skipFlag) {
      const timestampFlag = parseInt(skipFlag);
      if (!isNaN(timestampFlag)) {
        const segundosTranscurridos = (Date.now() - timestampFlag) / 1000;
        if (segundosTranscurridos < 60) {
          log_('INFO', `onEditSheet: ignorando cambio de ${ticketId} (vino de CalendarSync hace ${segundosTranscurridos.toFixed(1)}s)`);
          return;
        }
        // Si el flag expiró, lo limpio
        PropertiesService.getScriptProperties().deleteProperty('skipOnEdit_' + ticketId);
      }
    }

    log_('INFO', `onEditSheet: ${ticketId} fila ${fila} col ${colEditada} (totalCols=${totalCols})`);

    // Mapear columna editada
    const headers = hoja.getRange(1, 1, 1, totalCols).getValues()[0];
    const headerEditado = headers[colEditada - 1];
    const headerNorm = normalizarHeader_(headerEditado);
    const valorNuevo = e.value !== undefined ? e.value : e.range.getValue();
    const valorPrevio = e.oldValue;

    log_('INFO', `onEditSheet: header="${headerEditado}" (norm="${headerNorm}") valor: "${valorPrevio}" -> "${valorNuevo}"`);

    // ---------- Cambio en Estado (última columna) ----------
    if (colEditada === totalCols) {
      log_('INFO', `onEditSheet: detectado cambio en Estado`);
      manejarCambioEstado_(hoja, fila, ticketId, valorPrevio, valorNuevo);
      return;
    }

    // ---------- Cambio en Fecha_Slot (penúltima - 1) ----------
    const colFechaSlot = totalCols + COL_FECHA_SLOT_OFFSET;  // totalCols - 2
    if (colEditada === colFechaSlot) {
      log_('INFO', `onEditSheet: detectado cambio en Fecha_Slot (col ${colFechaSlot})`);
      manejarCambioFecha_(hoja, fila, ticketId, valorPrevio, valorNuevo);
      return;
    }

    // ---------- Cambio en columnas del Form ----------
    const camposEvento = ['cliente', 'direccion', 'equipo', 'telefono_cliente', 'notas',
                          'tipo_de_servicio', 'departamento', 'fecha_sugerida', 'hora_sugerida'];
    if (camposEvento.includes(headerNorm)) {
      log_('INFO', `onEditSheet: detectado cambio en campo "${headerNorm}"`);

      // Casos especiales: fecha_sugerida y hora_sugerida también deben mover el evento
      if (headerNorm === 'fecha_sugerida') {
        log_('INFO', `onEditSheet: tratando "Fecha sugerida" como cambio de Fecha_Slot`);
        manejarCambioFecha_(hoja, fila, ticketId, valorPrevio, valorNuevo);
        // También actualizar Fecha_Slot para que ambas columnas queden iguales
        try {
          let fechaStr;
          if (valorNuevo instanceof Date) {
            fechaStr = Utilities.formatDate(valorNuevo, CFG.TIMEZONE, 'yyyy-MM-dd');
          } else if (typeof valorNuevo === 'number' || /^\d+(\.\d+)?$/.test(String(valorNuevo).trim())) {
            // Convertir serial de Sheets a fecha
            const serial = parseFloat(valorNuevo);
            const sheetsEpoch = new Date(1899, 11, 30);
            const fechaConvertida = new Date(sheetsEpoch.getTime() + serial * 86400000);
            fechaStr = Utilities.formatDate(fechaConvertida, CFG.TIMEZONE, 'yyyy-MM-dd');
          } else {
            fechaStr = String(valorNuevo).substring(0, 10);
          }
          hoja.getRange(fila, colFechaSlot).setValue(fechaStr);
          log_('INFO', `Fecha_Slot actualizada a "${fechaStr}"`);
        } catch(e) {
          log_('WARN', `No se pudo actualizar Fecha_Slot: ${e.message}`);
        }
        return;
      }
      if (headerNorm === 'hora_sugerida') {
        log_('INFO', `onEditSheet: detectado cambio en Hora sugerida, moviendo evento`);
        manejarCambioHora_(hoja, fila, ticketId, valorPrevio, valorNuevo);
        return;
      }

      manejarCambioCampo_(hoja, fila, ticketId, headerNorm, valorPrevio, valorNuevo);
      return;
    }

    log_('INFO', `onEditSheet: cambio en columna "${headerEditado}" no requiere acción`);

  } catch (err) {
    log_('ERROR', `onEditSheet ${err.message}\n${err.stack}`);
  }
}

// ----------------------------------------------------------------
// Cambio en columna Estado
// ----------------------------------------------------------------
function manejarCambioEstado_(hoja, fila, ticketId, viejo, nuevo) {
  const nuevoUp = String(nuevo || '').toUpperCase().trim();
  const viejoUp = String(viejo || '').toUpperCase().trim();

  if (nuevoUp === viejoUp) return;
  log_('INFO', `Estado de ${ticketId}: ${viejoUp} -> ${nuevoUp}`);

  if (nuevoUp === 'CANCELADO_MANUALMENTE') {
    cancelarTicketDesdeSheet_(hoja, fila, ticketId);
    return;
  }

  // Estados que solo notifican (no tocan Calendar):
  //   RECHAZADO, RESUELTO (antes LLAMADA_COMPLETADA / REVISION_COMPLETADA)
  // Detecta también si empieza con RECHAZADO (puede ir con motivo: "RECHAZADO: razón")
  const esRechazado = nuevoUp === 'RECHAZADO' || nuevoUp.startsWith('RECHAZADO');
  const esCompletada = nuevoUp === 'RESUELTO';

  if (esRechazado || esCompletada) {
    notificarCambioEstadoManual_(hoja, fila, ticketId, viejoUp, nuevoUp);
    return;
  }

  log_('INFO', `Estado ${nuevoUp} no requiere acción automática`);
}

// ----------------------------------------------------------------
// Notifica a los 3 (vendedor, técnico, jefatura) sobre cambio de estado manual
// ----------------------------------------------------------------
function notificarCambioEstadoManual_(hoja, fila, ticketId, estadoAnterior, estadoNuevo) {
  const ctx = leerContextoFila_(hoja, fila);

  const totalCols = hoja.getLastColumn();
  const tecnicoNombre = hoja.getRange(fila, totalCols + COL_TECNICO_OFFSET).getValue();

  // Buscar email del técnico
  let emailTecnico = null;
  if (tecnicoNombre) {
    const tecnicos = cargarTecnicos_();
    const tec = tecnicos.find(t => t.nombre === tecnicoNombre);
    if (tec) emailTecnico = tec.email;
  }

  // Texto descriptivo del nuevo estado
  let descripcionEstado = '';
  let asuntoEtiqueta = '';

  if (estadoNuevo.startsWith('RECHAZADO')) {
    const motivo = estadoNuevo.replace(/^RECHAZADO\s*:?/, '').trim() || '(sin motivo)';
    descripcionEstado = `El ticket fue marcado como RECHAZADO.\nMotivo: ${motivo}`;
    asuntoEtiqueta = 'RECHAZADO';
  } else if (estadoNuevo === 'RESUELTO') {
    descripcionEstado = 'El ticket fue marcado como RESUELTO manualmente.';
    asuntoEtiqueta = 'RESUELTO';
  }

  const cuerpo = `
${descripcionEstado}

-------------------------------------------------
Ticket:    ${ticketId}
Cliente:   ${ctx.cliente}
Equipo:    ${ctx.equipo}
Servicio:  ${ctx.tipoServicio}
Técnico:   ${tecnicoNombre || '(desconocido)'}
Vendedor:  ${ctx.vendedor || '(desconocido)'}
-------------------------------------------------
Estado anterior: ${estadoAnterior}
Estado nuevo:    ${estadoNuevo}
-------------------------------------------------

Este cambio fue realizado manualmente desde la hoja.
`;

  const asunto = `[${asuntoEtiqueta} ${ticketId}] ${ctx.cliente}`;

  // 1) Vendedor
  if (ctx.emailVendedor) {
    GmailApp.sendEmail(ctx.emailVendedor, asunto, cuerpo, { name: 'Sistema de Tickets' });
    log_('INFO', `Notificado vendedor de cambio ${estadoNuevo} en ${ticketId}`);
  }

  // 2) Técnico
  if (emailTecnico) {
    GmailApp.sendEmail(emailTecnico, asunto, cuerpo, { name: 'Sistema de Tickets' });
    log_('INFO', `Notificado técnico de cambio ${estadoNuevo} en ${ticketId}`);
  }

  // 3) Jefatura
  GmailApp.sendEmail(CFG.EMAIL_JEFATURA, asunto, cuerpo, { name: 'Sistema de Tickets' });
  log_('INFO', `Notificado jefatura de cambio ${estadoNuevo} en ${ticketId}`);
}

function cancelarTicketDesdeSheet_(hoja, fila, ticketId) {
  const totalCols = hoja.getLastColumn();
  const eventIdRaw = hoja.getRange(fila, totalCols + COL_EVENT_ID_OFFSET).getValue();
  const tecnicoNombre = hoja.getRange(fila, totalCols + COL_TECNICO_OFFSET).getValue();

  // Leer datos del ticket para los correos
  const ctx = leerContextoFila_(hoja, fila);

  // 1) Borrar evento(s) del Calendar — puede haber 1 o 2 (modo HOTEL)
  if (eventIdRaw) {
    const idsParaBorrar = String(eventIdRaw).split(',').map(s => s.trim()).filter(Boolean);
    try {
      const cal = getCalendarCentral_();
      for (const evId of idsParaBorrar) {
        try {
          const evento = cal.getEventById(evId);
          if (evento) {
            evento.deleteEvent();
            log_('INFO', `Evento ${evId} borrado por cancelación manual de ${ticketId}`);
          }
        } catch(e) {
          log_('WARN', `No se pudo borrar evento ${evId}: ${e.message}`);
        }
      }
    } catch (err) {
      log_('WARN', `Error general borrando eventos de ${ticketId}: ${err.message}`);
    }
  }

  // 2) Buscar email del técnico
  let emailTecnico = null;
  if (tecnicoNombre) {
    const tecnicos = cargarTecnicos_();
    const tec = tecnicos.find(t => t.nombre === tecnicoNombre);
    if (tec) emailTecnico = tec.email;
  }

  // 3) Notificar a los 3
  const cuerpoBase = `
El ticket ${ticketId} fue cancelado manualmente desde la hoja.

-------------------------------------------------
Ticket:    ${ticketId}
Cliente:   ${ctx.cliente}
Equipo:    ${ctx.equipo}
Servicio:  ${ctx.tipoServicio}
Técnico:   ${tecnicoNombre || '(desconocido)'}
-------------------------------------------------

El evento del Calendar ya fue eliminado.
Si necesitas reagendar, debes generar un nuevo ticket.
`;

  if (ctx.emailVendedor) {
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[CANCELADO ${ticketId}] ${ctx.cliente}`,
      `Tu ticket ${ticketId} fue cancelado.\n` + cuerpoBase,
      { name: 'Sistema de Tickets' }
    );
  }

  if (emailTecnico) {
    GmailApp.sendEmail(
      emailTecnico,
      `[CANCELADO ${ticketId}] ${ctx.cliente}`,
      `El ticket ${ticketId} asignado a ti fue cancelado.\n` + cuerpoBase,
      { name: 'Sistema de Tickets' }
    );
  }

  GmailApp.sendEmail(
    CFG.EMAIL_JEFATURA,
    `[CANCELADO ${ticketId}] ${ctx.cliente} - cancelación manual desde Sheet`,
    `El ticket fue cancelado manualmente desde la Sheet.\n` + cuerpoBase,
    { name: 'Sistema de Tickets' }
  );

  log_('INFO', `Notificaciones de cancelación enviadas para ${ticketId}`);
}

// ----------------------------------------------------------------
// Cambio en Fecha_Slot
// ----------------------------------------------------------------
function manejarCambioFecha_(hoja, fila, ticketId, viejo, nuevo) {
  if (!nuevo) return;
  log_('INFO', `Fecha de ${ticketId}: ${viejo} -> ${nuevo}`);

  const totalCols = hoja.getLastColumn();
  const eventIdRaw = hoja.getRange(fila, totalCols + COL_EVENT_ID_OFFSET).getValue();
  if (!eventIdRaw) {
    log_('WARN', `${ticketId} no tiene Event_ID para mover`);
    return;
  }

  // Soporte modo HOTEL: pueden venir 2 IDs separados por coma
  const idsLista = String(eventIdRaw).split(',').map(s => s.trim()).filter(Boolean);
  const eventId = idsLista[0];
  const eventIdVuelta = idsLista[1] || null;

  try {
    const cal = getCalendarCentral_();
    const evento = cal.getEventById(eventId);
    if (!evento) {
      log_('WARN', `Evento ${eventId} no encontrado`);
      return;
    }

    // Parsear nueva fecha
    let nuevaFecha;
    if (nuevo instanceof Date) {
      nuevaFecha = nuevo;
    } else if (typeof nuevo === 'number' || /^\d+(\.\d+)?$/.test(String(nuevo).trim())) {
      // Número serial de Google Sheets (días desde 1899-12-30)
      // ej: 46164 = 19 de mayo de 2026
      const serial = parseFloat(nuevo);
      // Google Sheets epoch: 1899-12-30 (no 1900-01-01 por el bug histórico)
      const sheetsEpoch = new Date(1899, 11, 30);
      nuevaFecha = new Date(sheetsEpoch.getTime() + serial * 86400000);
      log_('INFO', `Convertido serial ${serial} a fecha ${nuevaFecha}`);
    } else {
      // Intentar parsear como yyyy-MM-dd
      const m = String(nuevo).match(/(\d{4})-(\d{2})-(\d{2})/);
      if (m) {
        nuevaFecha = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
      } else {
        nuevaFecha = new Date(nuevo);
      }
    }
    if (isNaN(nuevaFecha.getTime())) {
      log_('ERROR', `Fecha inválida en ${ticketId}: ${nuevo} (typeof=${typeof nuevo})`);
      return;
    }

    // Mantener la hora original del evento
    const inicioOriginal = evento.getStartTime();
    const finOriginal = evento.getEndTime();
    const duracionMs = finOriginal.getTime() - inicioOriginal.getTime();

    nuevaFecha.setHours(inicioOriginal.getHours(),
                        inicioOriginal.getMinutes(), 0, 0);

    // Si el evento ya está en esa fecha (mismo día, misma hora), no hacer nada.
    // Evita loops cuando CalendarSync escribió la fecha en Sheet y onEditSheet se dispara.
    if (Math.abs(nuevaFecha.getTime() - inicioOriginal.getTime()) < 60000) {
      log_('INFO', `Evento ${eventId} ya está en esa fecha, no se modifica (evita loop)`);
      return;
    }

    const nuevoFin = new Date(nuevaFecha.getTime() + duracionMs);

    evento.setTime(nuevaFecha, nuevoFin);
    log_('INFO', `Evento ${eventId} movido a ${nuevaFecha}`);

    // Actualizar snapshot para que CalendarSync no detecte esto como "cambio manual"
    actualizarSnapshotEvento_(eventId, evento);

    // Si modo HOTEL, mover también el evento de vuelta al día siguiente
    if (eventIdVuelta) {
      try {
        const evV = cal.getEventById(eventIdVuelta);
        if (evV) {
          const inicioVOriginal = evV.getStartTime();
          const finVOriginal = evV.getEndTime();
          const duracionVMs = finVOriginal.getTime() - inicioVOriginal.getTime();

          const nuevaFechaV = new Date(nuevaFecha);
          nuevaFechaV.setDate(nuevaFechaV.getDate() + 1);
          if (nuevaFechaV.getDay() === 6) nuevaFechaV.setDate(nuevaFechaV.getDate() + 2);
          else if (nuevaFechaV.getDay() === 0) nuevaFechaV.setDate(nuevaFechaV.getDate() + 1);

          nuevaFechaV.setHours(inicioVOriginal.getHours(),
                               inicioVOriginal.getMinutes(), 0, 0);
          const nuevoFinV = new Date(nuevaFechaV.getTime() + duracionVMs);
          evV.setTime(nuevaFechaV, nuevoFinV);
          log_('INFO', `Evento de vuelta ${eventIdVuelta} también movido a ${nuevaFechaV}`);
          actualizarSnapshotEvento_(eventIdVuelta, evV);
        }
      } catch(e) {
        log_('WARN', `No se pudo mover evento de vuelta ${eventIdVuelta}: ${e.message}`);
      }
    }

    // Notificar
    const ctx = leerContextoFila_(hoja, fila);
    const tecnicoNombre = hoja.getRange(fila, totalCols + COL_TECNICO_OFFSET).getValue();
    notificarCambioDesdeSheet_(ctx, tecnicoNombre, ticketId, 'Fecha del evento', nuevaFecha);
  } catch (err) {
    log_('ERROR', `manejarCambioFecha ${ticketId}: ${err.message}`);
  }
}

// ----------------------------------------------------------------
// Cambio en campos del Form (cliente, dirección, equipo, etc)
// ----------------------------------------------------------------
function manejarCambioCampo_(hoja, fila, ticketId, headerNorm, viejo, nuevo) {
  log_('INFO', `Campo ${headerNorm} de ${ticketId}: "${viejo}" -> "${nuevo}"`);

  const totalCols = hoja.getLastColumn();
  const eventId = hoja.getRange(fila, totalCols + COL_EVENT_ID_OFFSET).getValue();
  if (!eventId) {
    log_('INFO', `${ticketId} no tiene Event_ID, no se actualiza Calendar`);
    return;
  }

  try {
    const cal = getCalendarCentral_();
    // Puede haber varios Event_IDs separados por coma (modo HOTEL)
    const idsLista = String(eventId).split(',').map(s => s.trim()).filter(Boolean);
    const evento = cal.getEventById(idsLista[0]);
    if (!evento) {
      log_('WARN', `Evento ${idsLista[0]} no encontrado para ${ticketId}`);
      return;
    }

    const ctx = leerContextoFila_(hoja, fila);

    // Reconstruir título con datos actuales
    const tecnicoNombre = hoja.getRange(fila, totalCols + COL_TECNICO_OFFSET).getValue();
    const estado = String(hoja.getRange(fila, totalCols).getValue() || '').toUpperCase();
    const prefijo = (estado === 'BORRADOR') ? '[BORRADOR] ' : '';

    const nuevoTitulo = `${prefijo}[${ctx.prioridad}] ${ctx.cliente} - ${ctx.equipo} ` +
                        `(${ctx.tipoServicio}) - ${tecnicoNombre || '?'}`;
    evento.setTitle(nuevoTitulo);

    // Reconstruir descripción con datos actuales
    let descripcion = `Ticket: ${ticketId}\n` +
                      `Cliente: ${ctx.cliente}\n` +
                      `Teléfono: ${ctx.telefonoCliente || '(no especificado)'}\n` +
                      `Dirección: ${ctx.direccion}\n` +
                      `Departamento: ${ctx.departamento}\n` +
                      `Equipo: ${ctx.equipo}\n` +
                      `Servicio: ${ctx.tipoServicio}\n` +
                      `Vendedor: ${ctx.vendedor}\n` +
                      `Notas: ${ctx.notas || '(sin notas)'}\n\n` +
                      `[ACTUALIZADO desde Sheet: ${headerNorm}]`;

    evento.setDescription(descripcion);

    // Si cambió la dirección, también actualizar location
    if (headerNorm === 'direccion') {
      evento.setLocation(ctx.direccion);
    }

    // SI CAMBIÓ EL TÉCNICO: actualizar invitados de TODOS los eventos vinculados
    // y re-enviar correo de aprobación al técnico nuevo
    if (headerNorm === 'tecnico') {
      log_('INFO', `Cambio de técnico detectado en ${ticketId}: nuevo técnico "${tecnicoNombre}"`);

      // Buscar email del nuevo técnico
      const todosTecnicos = cargarTecnicos_();
      const nuevoTec = todosTecnicos.find(t => normalizar_(t.nombre) === normalizar_(tecnicoNombre));

      if (!nuevoTec || !nuevoTec.email) {
        log_('ERROR', `Técnico "${tecnicoNombre}" no encontrado o sin email`);
        hoja.getRange(fila, totalCols).setValue(estado);  // Mantener estado
        notificarErrorCambioTecnico_(ctx, ticketId, tecnicoNombre);
        return;
      }

      // Actualizar invitado en TODOS los eventos del ticket (IDA + REGRESO en HOTEL)
      for (const id of idsLista) {
        try {
          const ev = cal.getEventById(id);
          if (!ev) continue;

          // Remover invitados viejos
          const invitadosViejos = ev.getGuestList();
          for (const g of invitadosViejos) {
            try { ev.removeGuest(g.getEmail()); } catch(_) {}
          }

          // Agregar al nuevo técnico
          ev.addGuest(nuevoTec.email);
          log_('INFO', `Evento ${id}: invitado actualizado a ${nuevoTec.email}`);
        } catch(e) {
          log_('WARN', `No se pudo actualizar invitado del evento ${id}: ${e.message}`);
        }
      }

      // Si el estado es APROBADO_ESPERA_TECNICO o CONFIRMADO, re-enviar correo al técnico nuevo
      const estadosQueRequierenReenvio = ['APROBADO_ESPERA_TECNICO', 'CONFIRMADO'];
      if (estadosQueRequierenReenvio.includes(estado)) {
        log_('INFO', `Reenviando correo de aprobación al nuevo técnico ${nuevoTec.nombre}`);

        // Reconstruir contexto completo para el correo
        const contextoCompleto = buscarTicketPorId_(ticketId);
        if (contextoCompleto && contextoCompleto.tecnico && contextoCompleto.tecnico.email) {
          try {
            // Si estaba CONFIRMADO, volver a APROBADO_ESPERA_TECNICO porque el técnico nuevo
            // necesita responder horas nuevamente
            if (estado === 'CONFIRMADO') {
              hoja.getRange(fila, totalCols).setValue('APROBADO_ESPERA_TECNICO');
              log_('INFO', `Estado revertido a APROBADO_ESPERA_TECNICO porque cambió técnico`);
            }
            enviarCorreoATecnico_(contextoCompleto);
            log_('INFO', `Correo enviado al nuevo técnico ${nuevoTec.email}`);
          } catch(e) {
            log_('ERROR', `No se pudo enviar correo al nuevo técnico: ${e.message}`);
          }
        }
      }
    }

    log_('INFO', `Evento ${idsLista[0]} actualizado por cambio de ${headerNorm}`);

    // Actualizar snapshot para que CalendarSync no detecte esto como "cambio manual"
    for (const id of idsLista) {
      try {
        const ev = cal.getEventById(id);
        if (ev) actualizarSnapshotEvento_(id, ev);
      } catch(_) {}
    }

    // Notificar
    notificarCambioDesdeSheet_(ctx, tecnicoNombre, ticketId, headerNorm, nuevo);
  } catch (err) {
    log_('ERROR', `manejarCambioCampo ${ticketId}: ${err.message}\n${err.stack}`);
  }
}

/**
 * Notifica al vendedor + jefatura cuando el cambio de técnico falla por técnico inexistente.
 */
function notificarErrorCambioTecnico_(ctx, ticketId, tecnicoMalo) {
  try {
    const cuerpo = `Error al cambiar técnico del ticket ${ticketId}.\n\n` +
                   `El técnico "${tecnicoMalo}" no está registrado o no tiene email en la pestaña Tecnicos.\n\n` +
                   `Por favor verifica el nombre exacto del técnico y vuelve a intentarlo.`;
    GmailApp.sendEmail(
      CFG.EMAIL_JEFATURA,
      `[ERROR CAMBIO TÉCNICO ${ticketId}] Técnico "${tecnicoMalo}" no encontrado`,
      cuerpo,
      { name: 'Sistema de Tickets' }
    );
  } catch(_) {}
}

// ----------------------------------------------------------------
// Helper: leer contexto del ticket desde la fila
// ----------------------------------------------------------------
function leerContextoFila_(hoja, fila) {
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const valores = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
  const t = {};
  headers.forEach((h, i) => t[normalizarHeader_(h)] = valores[i]);

  return {
    vendedor: t.vendedor,
    emailVendedor: t.correo_vendedor || t.email_vendedor,
    cliente: t.cliente,
    telefonoCliente: t.telefono_cliente || t.telefono || '',
    direccion: t.direccion || t.ubicacion,
    departamento: t.departamento,
    equipo: normalizarListaEquipos_(t.equipo),
    tipoServicio: t.tipo_de_servicio || t.servicio,
    prioridad: t.prioridad || 'Media',
    notas: t.notas || t.observaciones || ''
  };
}

// ----------------------------------------------------------------
// Notificar cambio a vendedor + técnico + jefatura
// ----------------------------------------------------------------
function notificarCambioDesdeSheet_(ctx, tecnicoNombre, ticketId, campo, nuevoValor) {
  let emailTecnico = null;
  if (tecnicoNombre) {
    const tecnicos = cargarTecnicos_();
    const tec = tecnicos.find(t => t.nombre === tecnicoNombre);
    if (tec) emailTecnico = tec.email;
  }

  const valorTexto = (nuevoValor instanceof Date)
    ? Utilities.formatDate(nuevoValor, CFG.TIMEZONE, 'yyyy-MM-dd HH:mm')
    : String(nuevoValor);

  const cuerpo = `
Se actualizó el ticket ${ticketId} desde la Sheet.

-------------------------------------------------
Ticket:    ${ticketId}
Cliente:   ${ctx.cliente}
Equipo:    ${ctx.equipo}
Técnico:   ${tecnicoNombre || '(desconocido)'}
-------------------------------------------------
Campo modificado: ${campo}
Nuevo valor:      ${valorTexto}
-------------------------------------------------

El evento de Calendar ya fue actualizado con esta información.
`;

  const asunto = `[ACTUALIZADO ${ticketId}] ${ctx.cliente} - ${campo}`;

  if (ctx.emailVendedor) {
    GmailApp.sendEmail(ctx.emailVendedor, asunto, cuerpo, { name: 'Sistema de Tickets' });
  }
  if (emailTecnico) {
    GmailApp.sendEmail(emailTecnico, asunto, cuerpo, { name: 'Sistema de Tickets' });
  }
  GmailApp.sendEmail(CFG.EMAIL_JEFATURA, asunto, cuerpo, { name: 'Sistema de Tickets' });
}

/**
 * Actualiza la fila correspondiente en SnapshotCalendar para que la
 * próxima ejecución de detectarCambiosEnCalendar no marque este
 * cambio como "manual del usuario" (porque ya lo iniciamos nosotros desde Sheet).
 */
function actualizarSnapshotEvento_(eventId, evento) {
  try {
    const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    const snap = ss.getSheetByName('SnapshotCalendar');
    if (!snap) return;

    const data = snap.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        const inicio = evento.getStartTime();
        const fin = evento.getEndTime();
        const duracion = (fin - inicio) / 3600000;
        const guests = evento.getGuestList()
          .map(g => g.getEmail().toLowerCase())
          .sort()
          .join(',');

        snap.getRange(i + 1, 3).setValue(
          Utilities.formatDate(inicio, CFG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss")
        );
        snap.getRange(i + 1, 4).setValue(duracion.toFixed(2));
        snap.getRange(i + 1, 5).setValue(guests);
        snap.getRange(i + 1, 6).setValue(evento.getTitle());
        snap.getRange(i + 1, 7).setValue(new Date());
        break;
      }
    }
  } catch (e) {
    log_('WARN', `actualizarSnapshotEvento_ falló: ${e.message}`);
  }
}

/**
 * Función helper para configurar AMBOS triggers: el de Sheet y el de Calendar.
 * Solo se corre 1 vez al inicio.
 */
function instalarTriggerOnEdit() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);

  // Limpiar triggers viejos de onEditSheet
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'onEditSheet') {
      ScriptApp.deleteTrigger(t);
    }
    if (t.getHandlerFunction() === 'sincronizarConCalendar_') {
      ScriptApp.deleteTrigger(t);
    }
  }

  // 1) Trigger onEdit en la Sheet (Sheet -> Calendar)
  ScriptApp.newTrigger('onEditSheet')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  console.log('Trigger onEditSheet (Sheet -> Calendar) instalado');

  // 2) Trigger por tiempo cada 5 min para detectar cambios de Calendar (Calendar -> Sheet)
  ScriptApp.newTrigger('sincronizarConCalendar_')
    .timeBased()
    .everyMinutes(5)
    .create();
  console.log('Trigger sincronizarConCalendar_ (Calendar -> Sheet) instalado cada 5 min');
}

/**
 * Diagnóstico: lista los triggers activos y reporta si onEditSheet está instalado.
 * Útil para verificar que los cambios en Sheet se están detectando.
 */
function diagnosticarTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  console.log(`Total de triggers activos: ${triggers.length}\n`);

  let tieneOnEditSheet = false;
  let tieneSincronizar = false;

  for (const t of triggers) {
    const handler = t.getHandlerFunction();
    const tipo = t.getEventType();
    const tipoTexto = String(tipo);

    console.log(`- ${handler}() | tipo: ${tipoTexto}`);

    if (handler === 'onEditSheet') tieneOnEditSheet = true;
    if (handler === 'sincronizarConCalendar_') tieneSincronizar = true;
  }

  console.log('');
  console.log('===== RESUMEN =====');
  console.log(`onEditSheet (Sheet → Calendar): ${tieneOnEditSheet ? '✓ ACTIVO' : '✗ NO INSTALADO'}`);
  console.log(`sincronizarConCalendar_ (Calendar → Sheet): ${tieneSincronizar ? '✓ ACTIVO' : '✗ NO INSTALADO'}`);

  if (!tieneOnEditSheet || !tieneSincronizar) {
    console.log('\n⚠️ FALTA INSTALAR: corre instalarTriggerOnEdit()');
  } else {
    console.log('\n✓ Todos los triggers de sincronización están activos');
  }
}

/**
 * Maneja el cambio de Hora sugerida desde Sheet.
 * Mueve el evento del Calendar a la nueva hora manteniendo el día.
 */
function manejarCambioHora_(hoja, fila, ticketId, viejo, nuevo) {
  if (nuevo === '' || nuevo === null || nuevo === undefined) {
    log_('WARN', `Hora vacía para ${ticketId}, no se mueve`);
    return;
  }

  log_('INFO', `Hora de ${ticketId}: ${viejo} -> ${nuevo}`);

  const totalCols = hoja.getLastColumn();
  const eventIdRaw = hoja.getRange(fila, totalCols + COL_EVENT_ID_OFFSET).getValue();
  if (!eventIdRaw) {
    log_('WARN', `${ticketId} no tiene Event_ID para mover`);
    return;
  }

  // Parsear hora
  let horaDecimal = null;
  if (typeof nuevo === 'number') {
    horaDecimal = nuevo;
  } else {
    // Acepta formatos: "10:00", "10:30", "10.5", "10"
    const s = String(nuevo).trim();
    const matchHora = s.match(/^(\d{1,2}):(\d{2})/);
    if (matchHora) {
      horaDecimal = parseInt(matchHora[1]) + parseInt(matchHora[2]) / 60;
    } else {
      const num = parseFloat(s);
      if (!isNaN(num)) horaDecimal = num;
    }
  }

  if (horaDecimal === null || horaDecimal < 0 || horaDecimal >= 24) {
    log_('ERROR', `Hora inválida en ${ticketId}: "${nuevo}"`);
    return;
  }

  try {
    const cal = getCalendarCentral_();
    const idsLista = String(eventIdRaw).split(',').map(s => s.trim()).filter(Boolean);
    const evento = cal.getEventById(idsLista[0]);
    if (!evento) {
      log_('WARN', `Evento ${idsLista[0]} no encontrado para ${ticketId}`);
      return;
    }

    const inicioOriginal = evento.getStartTime();
    const finOriginal = evento.getEndTime();
    const duracionMs = finOriginal.getTime() - inicioOriginal.getTime();

    // Mantener fecha, cambiar hora
    const h = Math.floor(horaDecimal);
    const m = Math.round((horaDecimal % 1) * 60);
    const nuevoInicio = new Date(inicioOriginal);
    nuevoInicio.setHours(h, m, 0, 0);
    const nuevoFin = new Date(nuevoInicio.getTime() + duracionMs);

    // Verificar que realmente cambia
    if (Math.abs(nuevoInicio.getTime() - inicioOriginal.getTime()) < 60000) {
      log_('INFO', `Evento ${ticketId} ya está en esa hora, no se mueve`);
      return;
    }

    evento.setTime(nuevoInicio, nuevoFin);
    log_('INFO', `Evento ${ticketId} movido a hora ${h}:${String(m).padStart(2,'0')}`);

    actualizarSnapshotEvento_(idsLista[0], evento);

    // Notificar a vendedor + técnico + jefatura
    notificarCambioDesdeSheet_(hoja, fila, ticketId,
      `Hora del ticket cambiada manualmente.\nNueva hora: ${h}:${String(m).padStart(2,'0')}`);

  } catch (err) {
    log_('ERROR', `Error moviendo hora de ${ticketId}: ${err.message}`);
  }
}
