/**
 * =====================================================================
 * APROBACION.gs — v4 — sin emojis + viáticos en correos
 * =====================================================================
 */

function procesarBandejaEntrada() {
  console.log('=== Procesando bandeja de entrada ===');
  try {
    const n1 = procesarAprobacionesJefatura_();
    const n2 = procesarRespuestasTecnicos_();
    const n3 = procesarUrgentesSinSlot_();
    const n4 = procesarTrackingLlamadas_();
    const n5 = procesarRSTsNuevos();
    console.log(`Aprobaciones: ${n1} | Técnicos: ${n2} | Urgentes: ${n3} | Llamadas: ${n4} | RSTs: ${n5}`);
  } catch (err) {
    log_('ERROR', `procesarBandejaEntrada: ${err.message}\n${err.stack}`);
  }
}

// =====================================================================
// A) APROBACIONES DE JEFATURA
// =====================================================================
function procesarAprobacionesJefatura_() {
  // Buscar correos de aprobación: visitas (APROBAR), llamadas/revisiones (APROBAR LLAMADA/REVISION)
  // y escalamientos solicitados por técnicos (ESCALAMIENTO SOLICITADO)
  const query = `(subject:APROBAR OR subject:"ESCALAMIENTO SOLICITADO") newer_than:7d`;
  const hilos = GmailApp.search(query, 0, 30);
  console.log(`[v7] Hilos aprobación: ${hilos.length}`);
  log_('INFO', `[v7] procesarAprobacionesJefatura: ${hilos.length} hilos encontrados`);

  let procesados = 0;

  for (const hilo of hilos) {
    const mensajes = hilo.getMessages();
    const asunto = hilo.getFirstMessageSubject();

    const matchId = asunto.match(/TCK\d+/);
    if (!matchId) continue;
    const ticketId = matchId[0];

    if (mensajes.length < 2) continue;

    // Buscar el último mensaje del hilo, identificando si es del sistema o del usuario
    // Criterio: el sistema envía mensajes que SON respuestas automáticas con texto específico.
    // NO usamos el From (porque jefaturadt envía como sistema Y como usuario).
    let ultimo = null;
    let ultimoEsRespuestaBot = false;  // si el último mensaje (cronológico) es del bot, ya está procesado
    const ultimoMensajeReal = mensajes[mensajes.length - 1];
    const cuerpoUltimoReal = extraerCuerpoPlano_(ultimoMensajeReal).trim();
    const lineaUltimoReal = cuerpoUltimoReal.split('\n')[0].trim();
    if (/^Ticket TCK\d+ (aprobado|rechazado|reasignado|fecha cambiada)/i.test(lineaUltimoReal) ||
        /^Escalamiento de TCK\d+/i.test(lineaUltimoReal) ||
        /^Error:/i.test(lineaUltimoReal) ||
        /notificado al t[eé]cnico/i.test(lineaUltimoReal)) {
      // El bot ya respondió al último comando del usuario → no hay nada nuevo que procesar
      ultimoEsRespuestaBot = true;
    }

    if (ultimoEsRespuestaBot) {
      // Marcar todo el hilo como leído para que no vuelva a entrar
      try {
        for (const m of mensajes) {
          try { m.markRead(); } catch(_) {}
        }
      } catch(_) {}
      continue;
    }

    for (let i = mensajes.length - 1; i >= 0; i--) {
      const m = mensajes[i];
      const cuerpoPrev = extraerCuerpoPlano_(m).trim();
      const lineaPrev = cuerpoPrev.split('\n')[0].trim();

      // Ignorar SOLO respuestas automatizadas del sistema (texto típico exacto)
      if (/^Ticket TCK\d+ (aprobado|rechazado|reasignado|fecha cambiada)/i.test(lineaPrev)) continue;
      if (/^Escalamiento de TCK\d+/i.test(lineaPrev)) continue;
      if (/^Error:/i.test(lineaPrev)) continue;
      // Confirmaciones genéricas del bot
      if (/^Skip:/i.test(lineaPrev)) continue;
      if (/notificado al t[eé]cnico/i.test(lineaPrev)) continue;

      ultimo = m;
      break;
    }

    if (!ultimo) {
      log_('INFO', `Hilo ${ticketId}: ningún mensaje de usuario, ignorando`);
      continue;
    }

    const ctx = buscarTicketPorId_(ticketId);
    if (!ctx) continue;

    const estadoActual = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();

    // ATAJO: si el ticket está en estado terminal/cerrado, ignorar el hilo completo.
    // Esto evita reprocesar respuestas viejas y ensuciar el log.
    const estadosTerminales = [
      'RESUELTO', 'RECHAZADO', 'CANCELADO_MANUALMENTE',
      // Estados viejos (no migrados):
      'LLAMADA_COMPLETADA', 'REVISION_COMPLETADA',
      'LLAMADA_ESCALADA', 'REVISION_ESCALADA',
      'LLAMADA_VISTA', 'REVISION_VISTA'
    ];
    if (estadosTerminales.some(s => estadoActual.startsWith(s))) {
      // Ticket ya cerrado: archivar el hilo para sacarlo de la bandeja
      // y marcar como leído para evitar futuras revisiones
      try {
        hilo.moveToArchive();
        for (const m of mensajes) {
          try { m.markRead(); } catch(_) {}
        }
      } catch(e) {
        // Si falla el archive (ej: ya archivado), continuar silenciosamente
      }
      continue;
    }

    // Estados donde se puede APROBAR (primera vez)
    const estadosAprobables = ['BORRADOR', 'BORRADOR_FORZADO',
                               'LLAMADA_PENDIENTE_APROBACION',
                               'REVISION_PENDIENTE_APROBACION',
                               'ESCALAMIENTO_PENDIENTE_APROBACION'];

    // Estados donde se puede REASIGNAR técnico (después de aprobar también)
    const estadosReasignables = [
      'BORRADOR', 'BORRADOR_FORZADO',
      'LLAMADA_PENDIENTE_APROBACION', 'REVISION_PENDIENTE_APROBACION',
      'LLAMADA_ENVIADA', 'PENDIENTE',
      'REVISION_ENVIADA',
      'APROBADO_ESPERA_TECNICO', 'CONFIRMADO',
      'ESCALAMIENTO_PENDIENTE_APROBACION'
    ];

    const esRemoto = estadoActual.includes('LLAMADA') || estadoActual.includes('REVISION');

    const cuerpo = extraerCuerpoPlano_(ultimo);
    const primeraLinea = cuerpo.split('\n')[0].trim().toUpperCase();
    log_('INFO', `Aprobación ${ticketId} (${estadoActual}) primeraLinea: "${primeraLinea}"`);
    log_('INFO', `Aprobación ${ticketId} cuerpo plano (primeros 200 chars): "${cuerpo.substring(0, 200)}"`);
    console.log(`Ticket ${ticketId} (${estadoActual}): "${primeraLinea}"`);

    // Si el comando es CAMBIAR, permitir desde casi cualquier estado activo
    const esComandoCambiar = /^CAMBIAR\s*:/.test(primeraLinea);
    const esComandoCambiarFecha = /^CAMBIAR_FECHA\s*:/.test(primeraLinea);
    const esComandoRechazado = /^RECHAZADO\b/.test(primeraLinea);
    // APROBADO debe ser SOLO la palabra APROBADO, no parte de otra (APROBADO_ESPERA, etc)
    // Acepta: "APROBADO", "APROBADO ", "APROBADO.", "APROBADO: nombre" (para escalamiento)
    const esComandoAprobado = /^APROBADO\s*[\.\!\?]*\s*$/.test(primeraLinea) ||
                              /^APROBADO\s+\S/.test(primeraLinea) ||  // "APROBADO con comentario"
                              /^APROBADO\s*:\s*\S/.test(primeraLinea);  // "APROBADO: nombre"

    // Validar que el estado permite la acción solicitada
    if (esComandoAprobado && !estadosAprobables.includes(estadoActual)) {
      console.log(`Skip: ticket ${ticketId} ya no está en estado aprobable (${estadoActual})`);
      continue;
    }
    if ((esComandoCambiar || esComandoCambiarFecha) && !estadosReasignables.includes(estadoActual)) {
      console.log(`Skip: ticket ${ticketId} no se puede reasignar en estado ${estadoActual}`);
      continue;
    }
    if (!esComandoAprobado && !esComandoCambiar && !esComandoCambiarFecha && !esComandoRechazado) {
      continue;  // No es un comando válido, no procesar
    }

    try {
      if (esComandoAprobado) {
        // Caso especial: si el ticket está esperando aprobación de escalamiento del técnico
        if (estadoActual === 'ESCALAMIENTO_PENDIENTE_APROBACION') {
          // Permitir "APROBADO: <nombre>" para forzar técnico al aprobar escalamiento
          const matchTec = primeraLinea.match(/^APROBADO\s*:\s*(.+?)\s*$/);
          const tecForzado = matchTec ? matchTec[1].trim() : null;
          aprobarEscalamientoTecnico_(ticketId, tecForzado);
          if (tecForzado) {
            hilo.reply(`Escalamiento de ${ticketId} aprobado con técnico ${tecForzado}. Procesando como visita.`);
          } else {
            hilo.reply(`Escalamiento de ${ticketId} aprobado. Procesando como visita.`);
          }
          procesados++;
        } else if (esRemoto) {
          aprobarLlamadaORevision_(ticketId);
          // Verificar que el estado realmente cambió
          const estadoDespues = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();
          if (estadoDespues === estadoActual) {
            hilo.reply(`Error: el estado del ticket ${ticketId} no cambió. Revisa el log para más detalles.`);
            log_('ERROR', `Aprobación de ${ticketId} no cambió el estado (sigue en ${estadoActual})`);
          } else {
            hilo.reply(`Ticket ${ticketId} aprobado y notificado al técnico.`);
          }
          procesados++;
        } else {
          confirmarTicket_(ticketId);
          // Verificar que el estado realmente cambió
          const estadoDespues = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();
          if (estadoDespues === estadoActual) {
            hilo.reply(`Error: el estado del ticket ${ticketId} no cambió. Revisa el log para más detalles.`);
            log_('ERROR', `confirmarTicket de ${ticketId} no cambió el estado (sigue en ${estadoActual}). Posible problema con el evento Calendar o el técnico.`);
          } else {
            hilo.reply(`Ticket ${ticketId} aprobado y notificado al técnico.`);
          }
          procesados++;
        }
      } else if (esComandoCambiarFecha) {
        const fechaTexto = primeraLinea.replace(/^CAMBIAR_FECHA\s*:/, '').trim();
        cambiarFechaTicket_(ticketId, fechaTexto);
        hilo.reply(`Ticket ${ticketId} fecha cambiada a ${fechaTexto}.`);
        procesados++;
      } else if (esComandoCambiar) {
        const nuevoTec = primeraLinea.replace(/^CAMBIAR\s*:/, '').trim();
        // Caso especial: en ESCALAMIENTO_PENDIENTE_APROBACION, CAMBIAR significa
        // aprobar el escalamiento Y asignar al técnico especificado
        if (estadoActual === 'ESCALAMIENTO_PENDIENTE_APROBACION') {
          aprobarEscalamientoTecnico_(ticketId, nuevoTec);
          hilo.reply(`Escalamiento de ${ticketId} aprobado con técnico ${nuevoTec}. Procesando como visita.`);
        } else if (esRemoto) {
          reasignarLlamadaORevision_(ticketId, nuevoTec);
          hilo.reply(`Ticket ${ticketId} reasignado a ${nuevoTec}.`);
        } else {
          reasignarTicket_(ticketId, nuevoTec);
          hilo.reply(`Ticket ${ticketId} reasignado a ${nuevoTec}.`);
        }
        // CRÍTICO: marcar todos los mensajes del hilo como leídos para evitar loop
        try {
          for (const m of mensajes) {
            try { m.markRead(); } catch(_) {}
          }
        } catch(_) {}
        procesados++;
      } else if (esComandoRechazado) {
        const motivo = primeraLinea.replace(/^RECHAZADO\s*:?/, '').trim() || 'Sin motivo';
        // Caso especial: rechazar escalamiento solicitado por técnico
        if (estadoActual === 'ESCALAMIENTO_PENDIENTE_APROBACION') {
          rechazarEscalamientoTecnico_(ticketId, motivo);
          hilo.reply(`Escalamiento de ${ticketId} rechazado. Ticket queda como completado.`);
        } else {
          rechazarTicket_(ticketId, motivo);
          hilo.reply(`Ticket ${ticketId} rechazado.`);
        }
        procesados++;
      }
    } catch (err) {
      log_('ERROR', `Aprobación ${ticketId}: ${err.message}`);
      hilo.reply(`Error: ${err.message}`);
    }
    ultimo.markRead();
  }
  return procesados;
}

// =====================================================================
// B) URGENTES SIN SLOT
// =====================================================================
function procesarUrgentesSinSlot_() {
  const query = `subject:"URGENTE SIN SLOT" newer_than:7d`;
  const hilos = GmailApp.search(query, 0, 20);
  console.log(`Hilos urgentes sin slot: ${hilos.length}`);

  let procesados = 0;

  for (const hilo of hilos) {
    const mensajes = hilo.getMessages();
    if (mensajes.length < 2) continue;

    const ultimo = mensajes[mensajes.length - 1];
    const asunto = hilo.getFirstMessageSubject();

    const matchId = asunto.match(/TCK\d+/);
    if (!matchId) continue;
    const ticketId = matchId[0];

    const ctx = buscarTicketPorId_(ticketId);
    if (!ctx) continue;

    const estadoActual = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();
    if (!estadoActual.startsWith('URGENTE_SIN')) continue;

    const cuerpo = extraerCuerpoPlano_(ultimo);
    const primeraLinea = cuerpo.split('\n')[0].trim().toUpperCase();
    console.log(`URGENTE ${ticketId}: "${primeraLinea}"`);

    try {
      if (/^DESPLAZAR\s*:/.test(primeraLinea)) {
        const ticketDesplazadoId = primeraLinea.replace(/^DESPLAZAR\s*:/, '').trim();
        ejecutarDesplazamiento_(ticketId, ticketDesplazadoId);
        hilo.reply(`Ticket ${ticketId} agendado. ${ticketDesplazadoId} reubicado.`);
        procesados++;
      } else if (/^RETRASAR\b/.test(primeraLinea)) {
        ejecutarRetrasar_(ticketId);
        hilo.reply(`Ticket ${ticketId} agendado en el próximo hueco disponible.`);
        procesados++;
      } else if (/^FORZAR\s*:/.test(primeraLinea)) {
        const tecnicoForzado = primeraLinea.replace(/^FORZAR\s*:/, '').trim();
        ejecutarForzar_(ticketId, tecnicoForzado);
        hilo.reply(`Ticket ${ticketId} forzado a ${tecnicoForzado}.`);
        procesados++;
      } else if (/^RECHAZADO/.test(primeraLinea)) {
        const motivo = primeraLinea.replace(/^RECHAZADO\s*:?/, '').trim() || 'Sin motivo';
        ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue(`RECHAZADO: ${motivo}`);
        hilo.reply(`Ticket ${ticketId} rechazado.`);
        procesados++;
      } else {
        hilo.reply(`No entendí. Usa DESPLAZAR, RETRASAR, FORZAR o RECHAZADO.`);
      }
    } catch (err) {
      log_('ERROR', `Urgente ${ticketId}: ${err.message}`);
      hilo.reply(`Error: ${err.message}`);
    }
    ultimo.markRead();
  }
  return procesados;
}

// ---------- Desplazamiento ----------
function ejecutarDesplazamiento_(ticketUrgenteId, ticketDesplazadoId) {
  const ctxDesplazado = buscarTicketPorId_(ticketDesplazadoId);
  if (!ctxDesplazado) throw new Error(`No se encontró ${ticketDesplazadoId}`);

  try { ctxDesplazado.evento.deleteEvent(); } catch(_) {}

  const ticketDesplazadoData = leerTicketDeFilaComoNuevo_(ctxDesplazado);

  const ctxUrgente = buscarTicketPorId_(ticketUrgenteId);
  if (!ctxUrgente) throw new Error(`No se encontró ${ticketUrgenteId}`);

  const ticketUrgenteData = leerTicketDeFilaComoNuevo_(ctxUrgente);
  const destino = geocodificar_(ticketUrgenteData.direccion, ticketUrgenteData.departamento);
  ticketUrgenteData.lat = destino.lat;
  ticketUrgenteData.lng = destino.lng;

  const tecnicos = cargarTecnicos_();
  const resultado = evaluarCandidatos_(ticketUrgenteData, tecnicos);
  if (resultado.candidatosDirectos.length === 0) {
    throw new Error('Aún sin slot tras desplazar. Intenta RETRASAR.');
  }

  const elegido = resultado.candidatosDirectos[0];
  const viaticos = calcularViaticos_(ticketUrgenteData, elegido, 1);
  const evento = crearEventoBorrador_(ticketUrgenteData, elegido, viaticos);

  ctxUrgente.hoja.getRange(ctxUrgente.fila, ctxUrgente.hoja.getLastColumn() - 5, 1, 6).setValues([[
    ticketUrgenteId, elegido.tecnico.nombre, elegido.trasladoHoras.toFixed(2),
    elegido.fechaSlot, evento.getId(), 'BORRADOR'
  ]]);
  enviarCorreoAprobacion_(ticketUrgenteData, elegido, evento, viaticos);

  // Reagendar desplazado
  const destino2 = geocodificar_(ticketDesplazadoData.direccion, ticketDesplazadoData.departamento);
  ticketDesplazadoData.lat = destino2.lat;
  ticketDesplazadoData.lng = destino2.lng;
  const res2 = evaluarCandidatos_(ticketDesplazadoData, tecnicos);
  if (res2.candidatosDirectos.length > 0) {
    const e2 = res2.candidatosDirectos[0];
    const v2 = calcularViaticos_(ticketDesplazadoData, e2, 1);
    const ev2 = crearEventoBorrador_(ticketDesplazadoData, e2, v2);
    ctxDesplazado.hoja.getRange(ctxDesplazado.fila, ctxDesplazado.hoja.getLastColumn() - 5, 1, 6).setValues([[
      ticketDesplazadoId, e2.tecnico.nombre, e2.trasladoHoras.toFixed(2),
      e2.fechaSlot, ev2.getId(), 'BORRADOR'
    ]]);
    enviarCorreoAprobacion_(ticketDesplazadoData, e2, ev2, v2);
    log_('INFO', `${ticketDesplazadoId} reagendado a ${e2.fechaSlot}`);
  } else {
    log_('ERROR', `No se pudo reagendar ${ticketDesplazadoId}`);
  }

  actualizarDashboard_();
}

function ejecutarRetrasar_(ticketId) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  const ticket = leerTicketDeFilaComoNuevo_(ctx);
  const destino = geocodificar_(ticket.direccion, ticket.departamento);
  ticket.lat = destino.lat; ticket.lng = destino.lng;

  const tecnicos = cargarTecnicos_();
  const prioCfg = getConfigPrioridad_('Urgente');

  const candidatos = [];
  for (const tec of tecnicos) {
    if (!tec.activo || !tec.email) continue;

    const fechaBase = determinarFechaBase_(ticket.fechaSugerida);
    const origen = obtenerUbicacionActualTecnico_(tec, fechaBase);
    const traslado = distanceMatrix_(origen, `${ticket.lat},${ticket.lng}`);
    if (!traslado) continue;
    const trasladoHoras = traslado.duracionSegundos / 3600;

    const duracion = trasladoHoras * 2 + 1;
    const slot = buscarSlotConCapacidad_(
      tec, fechaBase, duracion,
      prioCfg.ventanaAmpliada,
      ticket.horaSugerida,
      prioCfg.toleranciaHoraSugerida
    );
    if (!slot) continue;

    const diasDemora = Math.floor((slot.fecha - fechaBase) / 86400000);
    candidatos.push({
      tecnico: tec,
      trasladoHoras,
      idaVueltaHoras: trasladoHoras * 2,
      fechaSlot: Utilities.formatDate(slot.fecha, CFG.TIMEZONE, 'yyyy-MM-dd'),
      horaInicioSlot: slot.horaInicio,
      distanciaKm: traslado.distanciaMetros / 1000,
      score: trasladoHoras * 2 + diasDemora * prioCfg.pesoDemora
    });
  }

  if (candidatos.length === 0) throw new Error('Sin slot aun ampliando ventana');
  candidatos.sort((a, b) => a.score - b.score);

  const elegido = candidatos[0];
  const viaticos = calcularViaticos_(ticket, elegido, 1);
  const evento = crearEventoBorrador_(ticket, elegido, viaticos);

  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn() - 5, 1, 6).setValues([[
    ticketId, elegido.tecnico.nombre, elegido.trasladoHoras.toFixed(2),
    elegido.fechaSlot, evento.getId(), 'BORRADOR'
  ]]);
  enviarCorreoAprobacion_(ticket, elegido, evento, viaticos);
  actualizarDashboard_();
}

function ejecutarForzar_(ticketId, nombreTecnico) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  const ticket = leerTicketDeFilaComoNuevo_(ctx);
  const tecnicos = cargarTecnicos_();
  const tec = tecnicos.find(t => normalizar_(t.nombre).includes(normalizar_(nombreTecnico)));
  if (!tec) throw new Error(`Técnico "${nombreTecnico}" no encontrado`);

  const destino = geocodificar_(ticket.direccion, ticket.departamento);
  ticket.lat = destino.lat; ticket.lng = destino.lng;

  const fecha = siguienteDiaHabil_(new Date());
  const origen = obtenerUbicacionActualTecnico_(tec, fecha);
  const traslado = distanceMatrix_(origen, `${ticket.lat},${ticket.lng}`);
  const trasladoHoras = traslado.duracionSegundos / 3600;
  const elegido = {
    tecnico: tec,
    trasladoHoras,
    idaVueltaHoras: trasladoHoras * 2,
    fechaSlot: Utilities.formatDate(fecha, CFG.TIMEZONE, 'yyyy-MM-dd'),
    horaInicioSlot: CFG.HORA_INICIO,
    distanciaKm: traslado.distanciaMetros / 1000
  };
  const viaticos = calcularViaticos_(ticket, elegido, 1);
  const evento = crearEventoBorrador_(ticket, elegido, viaticos);

  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn() - 5, 1, 6).setValues([[
    ticketId, tec.nombre, trasladoHoras.toFixed(2),
    elegido.fechaSlot, evento.getId(), 'BORRADOR_FORZADO'
  ]]);
  enviarCorreoAprobacion_(ticket, elegido, evento, viaticos);
  actualizarDashboard_();
}

function leerTicketDeFilaComoNuevo_(ctx) {
  return {
    id: ctx.ticketId,
    vendedor: ctx.vendedor,
    emailVendedor: ctx.emailVendedor,
    cliente: ctx.cliente,
    telefonoCliente: ctx.telefonoCliente || '',
    direccion: ctx.direccion,
    departamento: ctx.departamento,
    equipo: ctx.equipo,
    tipoServicio: ctx.tipoServicio,
    prioridad: ctx.prioridad || 'Media',
    fechaSugerida: ctx.fechaSugerida,
    horaSugerida: ctx.horaSugerida,
    notas: ctx.notas
  };
}

function confirmarTicket_(ticketId) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  // Verificar que existe el evento Calendar
  if (!ctx.evento) {
    log_('ERROR', `confirmarTicket_: ticket ${ticketId} NO tiene evento Calendar. No se puede confirmar.`);
    throw new Error(`Ticket ${ticketId} no tiene evento en Calendar. Verifica que la fila tenga Event_ID.`);
  }

  // Verificar que existe el técnico
  if (!ctx.tecnico || !ctx.tecnico.email) {
    log_('ERROR', `confirmarTicket_: ticket ${ticketId} NO tiene técnico asignado o el técnico no tiene email.`);
    throw new Error(`Ticket ${ticketId} no tiene técnico con email asignado.`);
  }

  log_('INFO', `confirmarTicket_: ${ticketId} -> ${ctx.tecnico.nombre} (${ctx.tecnico.email})`);

  const { hoja, fila, evento } = ctx;
  const tituloNuevo = evento.getTitle().replace(/^\[BORRADOR\]\s*/, '');
  evento.setTitle(tituloNuevo);
  try { evento.setColor(CFG.COLOR_CONFIRMADO); } catch(_) {}

  hoja.getRange(fila, hoja.getLastColumn()).setValue('APROBADO_ESPERA_TECNICO');
  log_('INFO', `confirmarTicket_: estado de ${ticketId} cambiado a APROBADO_ESPERA_TECNICO`);

  try {
    enviarCorreoATecnico_(ctx);
    log_('INFO', `confirmarTicket_: correo enviado a ${ctx.tecnico.email}`);
  } catch(e) {
    log_('ERROR', `confirmarTicket_: no se pudo enviar correo a técnico: ${e.message}`);
    throw e;  // Re-lanzar para que el handler superior lo vea
  }

  actualizarDashboard_();
}

function enviarCorreoATecnico_(ctx) {
  const { ticketId, tecnico, trasladoHoras, cliente, equipo, tipoServicio,
          direccion, fechaSlot, horaInicioSlot, telefonoCliente } = ctx;

  // VALIDACIONES PREVIAS
  if (!tecnico) {
    log_('ERROR', `enviarCorreoATecnico_: ${ticketId} sin objeto tecnico`);
    throw new Error(`Ticket ${ticketId}: no hay técnico asignado`);
  }

  if (!tecnico.email || String(tecnico.email).trim() === '') {
    log_('ERROR', `enviarCorreoATecnico_: ${ticketId} técnico "${tecnico.nombre}" sin email en Sheet`);
    throw new Error(`Técnico ${tecnico.nombre} no tiene email en la pestaña Tecnicos`);
  }

  const emailLimpio = String(tecnico.email).trim();
  const emailValido = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailLimpio);
  if (!emailValido) {
    log_('ERROR', `enviarCorreoATecnico_: ${ticketId} email mal formado: "${emailLimpio}"`);
    throw new Error(`Email del técnico "${tecnico.nombre}" mal formado: ${emailLimpio}`);
  }

  // Calcular viáticos para mostrar al técnico
  const tickFake = {
    departamento: ctx.departamento,
    horaSugerida: ctx.horaSugerida
  };
  const elegidoFake = {
    trasladoHoras, fechaSlot,
    horaInicioSlot, tecnico
  };
  let lineaViaticos = '';
  try {
    const viaticos = calcularViaticos_(tickFake, elegidoFake, 1);
    lineaViaticos = viaticos.aplica ? `\n${viaticos.resumen}\n` : '';
  } catch(e) {
    log_('WARN', `enviarCorreoATecnico_: no se pudo calcular viáticos: ${e.message}`);
  }

  const asunto = `[TICKET ${ticketId}] Nuevo trabajo asignado - Responde con horas`;
  const cuerpo = `
Hola ${tecnico.nombre},

Tienes un nuevo ticket asignado y aprobado.

-------------------------------------------------
Ticket:       ${ticketId}
Cliente:      ${cliente}
Teléfono:     ${telefonoCliente || '(no especificado)'}
Dirección:    ${direccion}
Equipo:       ${equipo}
Servicio:     ${tipoServicio}
Fecha:        ${fechaSlot} - ${formatearHora_(horaInicioSlot)}
Traslado ida: ${formatearHorasMinutos_(trasladoHoras)}${lineaViaticos}
-------------------------------------------------

¿Cuánto tiempo de TRABAJO necesitas (sin contar el viaje)?

RESPONDE CON UN NÚMERO. Ejemplos: 2, 2.5, 3h

Mantén el asunto intacto.
`;

  log_('INFO', `enviarCorreoATecnico_: enviando a "${emailLimpio}" para ${ticketId}`);
  try {
    GmailApp.sendEmail(emailLimpio, asunto, cuerpo, { name: 'Sistema de Tickets' });
    log_('INFO', `enviarCorreoATecnico_: enviado exitosamente a ${emailLimpio}`);
  } catch(e) {
    log_('ERROR', `enviarCorreoATecnico_: FALLÓ envío a ${emailLimpio}: ${e.message}`);
    throw new Error(`No se pudo enviar correo a ${tecnico.nombre} (${emailLimpio}): ${e.message}`);
  }
}

function reasignarTicket_(ticketId, nuevoNombreTecnico) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  const tecnicos = cargarTecnicos_();
  const nuevoTec = tecnicos.find(t =>
    normalizar_(t.nombre).includes(normalizar_(nuevoNombreTecnico)));
  if (!nuevoTec) throw new Error(`Técnico "${nuevoNombreTecnico}" no encontrado`);

  // Borrar evento anterior (puede haber 1 o 2 en modo HOTEL)
  const totalCols = ctx.hoja.getLastColumn();
  const eventIdActual = ctx.hoja.getRange(ctx.fila, totalCols - 1).getValue();
  if (eventIdActual) {
    const idsLista = String(eventIdActual).split(',').map(s => s.trim()).filter(Boolean);
    const cal = getCalendarCentral_();
    for (const id of idsLista) {
      try {
        const ev = cal.getEventById(id);
        if (ev) ev.deleteEvent();
      } catch(_) {}
    }
  }

  const ticket = leerTicketDeFilaComoNuevo_(ctx);
  const destino = geocodificar_(ticket.direccion, ticket.departamento);
  ticket.lat = destino.lat; ticket.lng = destino.lng;

  // FORZAR técnico sin chequear capacidad (es decisión explícita de jefatura)
  // Usa la fecha sugerida del ticket o el próximo día hábil
  let fecha;
  if (ticket.fechaSugerida) {
    fecha = new Date(ticket.fechaSugerida + 'T00:00:00');
    if (isNaN(fecha.getTime())) fecha = siguienteDiaHabil_(new Date());
  } else {
    fecha = siguienteDiaHabil_(new Date());
  }

  const origen = obtenerUbicacionActualTecnico_(nuevoTec, fecha);
  const traslado = distanceMatrix_(origen, `${ticket.lat},${ticket.lng}`);
  const trasladoHoras = traslado.duracionSegundos / 3600;

  const horaInicio = ticket.horaSugerida || CFG.HORA_INICIO;
  const elegido = {
    tecnico: nuevoTec,
    trasladoHoras,
    idaVueltaHoras: trasladoHoras * 2,
    fechaSlot: Utilities.formatDate(fecha, CFG.TIMEZONE, 'yyyy-MM-dd'),
    horaInicioSlot: horaInicio,
    distanciaKm: traslado.distanciaMetros / 1000
  };

  const viaticos = calcularViaticos_(ticket, elegido, 1);
  const evento = crearEventoBorrador_(ticket, elegido, viaticos);

  // Si es modo HOTEL, guardar ambos Event_IDs
  const eventIdParaSheet = evento._eventoVueltaId
    ? `${evento.getId()},${evento._eventoVueltaId}`
    : evento.getId();

  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn() - 5, 1, 6).setValues([[
    ticketId, nuevoTec.nombre, trasladoHoras.toFixed(2),
    elegido.fechaSlot, eventIdParaSheet, 'BORRADOR'
  ]]);
  enviarCorreoAprobacion_(ticket, elegido, evento, viaticos);
  actualizarDashboard_();
}

// =====================================================================
// Cambiar la fecha de un ticket en BORRADOR (mantiene técnico y todo)
// =====================================================================
function cambiarFechaTicket_(ticketId, fechaTexto) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  // Parsear fecha en formato YYYY-MM-DD
  const m = fechaTexto.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!m) throw new Error(`Formato de fecha inválido: "${fechaTexto}". Usa YYYY-MM-DD`);

  const nuevaFecha = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  if (isNaN(nuevaFecha.getTime())) throw new Error(`Fecha inválida: ${fechaTexto}`);

  // Validar que sea día laboral
  const dow = nuevaFecha.getDay();
  if (dow === 0 || dow === 6) {
    throw new Error(`${fechaTexto} es fin de semana, no es día laboral`);
  }

  if (!ctx.evento) throw new Error(`Ticket ${ticketId} no tiene evento`);

  // Mantener la hora original del evento
  const inicioOriginal = ctx.evento.getStartTime();
  const finOriginal = ctx.evento.getEndTime();
  const duracionMs = finOriginal.getTime() - inicioOriginal.getTime();

  nuevaFecha.setHours(inicioOriginal.getHours(),
                      inicioOriginal.getMinutes(), 0, 0);
  const nuevoFin = new Date(nuevaFecha.getTime() + duracionMs);

  // Mover el evento principal
  ctx.evento.setTime(nuevaFecha, nuevoFin);

  // Si es modo HOTEL, mover también el evento de vuelta al día siguiente
  if (ctx.eventoVuelta) {
    const inicioVOriginal = ctx.eventoVuelta.getStartTime();
    const finVOriginal = ctx.eventoVuelta.getEndTime();
    const duracionVMs = finVOriginal.getTime() - inicioVOriginal.getTime();

    const nuevaFechaV = new Date(nuevaFecha);
    nuevaFechaV.setDate(nuevaFechaV.getDate() + 1);
    // Si cae en sábado, saltar al lunes
    if (nuevaFechaV.getDay() === 6) nuevaFechaV.setDate(nuevaFechaV.getDate() + 2);
    else if (nuevaFechaV.getDay() === 0) nuevaFechaV.setDate(nuevaFechaV.getDate() + 1);

    nuevaFechaV.setHours(inicioVOriginal.getHours(),
                         inicioVOriginal.getMinutes(), 0, 0);
    const nuevoFinV = new Date(nuevaFechaV.getTime() + duracionVMs);
    ctx.eventoVuelta.setTime(nuevaFechaV, nuevoFinV);
    log_('INFO', `Evento de vuelta de ${ticketId} también movido a ${nuevaFechaV}`);
  }

  // Actualizar Fecha_Slot en Sheet
  const totalCols = ctx.hoja.getLastColumn();
  ctx.hoja.getRange(ctx.fila, totalCols - 2).setValue(fechaTexto);

  log_('INFO', `Ticket ${ticketId} fecha cambiada a ${fechaTexto}`);

  // Notificar al vendedor
  if (ctx.emailVendedor) {
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[TICKET ${ticketId}] Fecha actualizada a ${fechaTexto}`,
      `Tu ticket ${ticketId} fue reagendado por jefatura para optimizar la ruta del técnico.\n\n` +
      `Cliente:      ${ctx.cliente}\n` +
      `Técnico:      ${ctx.tecnico.nombre}\n` +
      `Nueva fecha:  ${fechaTexto}\n\n` +
      `El técnico será notificado cuando jefatura apruebe.`,
      { name: 'Sistema de Tickets' }
    );
  }

  actualizarDashboard_();
}


function rechazarTicket_(ticketId, motivo) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  // Borrar evento principal y de vuelta si existen (los remotos no tienen)
  if (ctx.evento) {
    try { ctx.evento.deleteEvent(); } catch(_) {}
  }
  if (ctx.eventoVuelta) {
    try { ctx.eventoVuelta.deleteEvent(); } catch(_) {}
  }
  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue(`RECHAZADO: ${motivo}`);

  if (ctx.emailVendedor) {
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[RECHAZADO ${ticketId}] ${ctx.cliente}`,
      `Tu ticket ${ticketId} fue rechazado.\n\nMotivo: ${motivo}`,
      { name: 'Sistema de Tickets' }
    );
  }
  actualizarDashboard_();
}

// =====================================================================
// APROBACIÓN DE LLAMADAS / REVISIONES (sin evento Calendar)
// =====================================================================
function aprobarLlamadaORevision_(ticketId) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  const tecnicos = cargarTecnicos_();
  const tec = tecnicos.find(t => t.nombre === ctx.tecnico.nombre);
  if (!tec) throw new Error(`Técnico ${ctx.tecnico.nombre} no encontrado`);

  // Construir objeto ticket para enviarCorreoLlamada_
  const ticket = {
    id: ticketId,
    cliente: ctx.cliente,
    telefonoCliente: ctx.telefonoCliente,
    direccion: ctx.direccion,
    departamento: ctx.departamento,
    equipo: ctx.equipo,
    tipoServicio: ctx.tipoServicio,
    vendedor: ctx.vendedor,
    notas: ctx.notas
  };

  // Enviar correo al técnico
  enviarCorreoLlamada_(ticket, tec);

  // Cambiar estado: PENDIENTE_APROBACION -> ENVIADA
  const estadoActual = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();
  const nuevoEstado = estadoActual.includes('REVISION') ? 'REVISION_ENVIADA' : 'LLAMADA_ENVIADA';
  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue(nuevoEstado);

  log_('INFO', `${ticketId} aprobado y enviado a técnico ${tec.nombre}`);

  // Notificar al vendedor que ya fue aprobado
  if (ctx.emailVendedor) {
    const tipoTexto = ctx.tipoServicio || 'Llamada';
    const tipoEtiqueta = tipoTexto.toUpperCase();
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[${tipoEtiqueta} APROBADA ${ticketId}] ${ctx.cliente}`,
      `Tu solicitud de ${tipoTexto.toLowerCase()} fue aprobada y asignada.\n\n` +
      `Técnico: ${tec.nombre}\n` +
      `Cliente: ${ctx.cliente}\n\n` +
      `El técnico contactará al cliente lo antes posible.`,
      { name: 'Sistema de Tickets' }
    );
  }

  actualizarDashboard_();
}

function reasignarLlamadaORevision_(ticketId, nuevoNombreTecnico) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  const tecnicos = cargarTecnicos_();
  const nuevoTec = tecnicos.find(t =>
    normalizar_(t.nombre) === normalizar_(nuevoNombreTecnico) && t.activo
  );
  if (!nuevoTec) throw new Error(`Técnico "${nuevoNombreTecnico}" no encontrado o inactivo`);

  // Actualizar técnico en Sheet
  const totalCols = ctx.hoja.getLastColumn();
  ctx.hoja.getRange(ctx.fila, totalCols - 4).setValue(nuevoTec.nombre);

  log_('INFO', `${ticketId} reasignado de ${ctx.tecnico.nombre} a ${nuevoTec.nombre}`);

  // Aprobar con el nuevo técnico
  aprobarLlamadaORevision_(ticketId);
}

// =====================================================================
// ESCALAR LLAMADA/REVISIÓN A VISITA
// Mantiene el mismo TCK, cambia tipo de servicio, busca técnico+slot,
// crea evento Calendar y manda correo de aprobación normal.
// Si se especifica tecnicoForzado, se pre-asigna ese técnico (jefatura decide).
// =====================================================================
function escalarLlamadaORevision_(ticketId, nuevoTipoServicio, tecnicoForzado) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  const estadoActual = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();
  log_('INFO', `Escalando ${ticketId} (${estadoActual}) a ${nuevoTipoServicio}`);

  // Lista canónica de tipos de visita (deben coincidir con las opciones del Form)
  // Cada entrada: alias normalizado → nombre oficial
  const tiposVisita = {
    'reparacion': 'Reparación',
    'mantenimiento': 'Mantenimiento',
    'instalacion': 'Instalación/Capacitación',
    'instalacion/capacitacion': 'Instalación/Capacitación',
    'capacitacion': 'Capacitación',
    'desinstalacion': 'Desinstalación/Retiro',
    'desinstalacion/retiro': 'Desinstalación/Retiro',
    'retiro': 'Desinstalación/Retiro',
    'visita tecnica': 'Visita técnica',
    'visita': 'Visita técnica'
  };

  const tipoNormBuscado = normalizar_(nuevoTipoServicio.trim());
  const matcheaTipo = tiposVisita[tipoNormBuscado];

  if (!matcheaTipo) {
    throw new Error(
      `Tipo de servicio "${nuevoTipoServicio}" no es válido para escalar. ` +
      `Usa: Reparación, Mantenimiento, Instalación/Capacitación, Capacitación, Desinstalación/Retiro o Visita técnica.`
    );
  }

  const esRevision = estadoActual.includes('REVISION');
  const estadoEscalado = 'ESCALADO';

  // 1. Notificar al técnico anterior si ya tenía la llamada/revisión asignada
  const tieneAsignacion = ['LLAMADA_ENVIADA', 'PENDIENTE',
                           'REVISION_ENVIADA',
                           'RESUELTO'].includes(estadoActual);

  if (tieneAsignacion && ctx.tecnico && ctx.tecnico.nombre) {
    const tecnicos = cargarTecnicos_();
    const tec = tecnicos.find(t => t.nombre === ctx.tecnico.nombre);
    if (tec && tec.email) {
      const tipoAnterior = esRevision ? 'revisión' : 'llamada';
      GmailApp.sendEmail(
        tec.email,
        `[ESCALADO ${ticketId}] ${ctx.cliente} - de ${tipoAnterior} a ${matcheaTipo}`,
        `El ticket ${ticketId} fue escalado por jefatura.\n\n` +
        `Cliente:     ${ctx.cliente}\n` +
        `Equipo:      ${ctx.equipo}\n` +
        `De:          ${tipoAnterior}\n` +
        `A:           ${matcheaTipo}\n\n` +
        `Este ticket ahora requiere visita presencial. Será reasignado y agendado automáticamente. ` +
        `Si el técnico asignado para la visita eres tú nuevamente, recibirás un correo aparte con los detalles.`,
        { name: 'Sistema de Tickets' }
      );
      log_('INFO', `Notificado a ${tec.nombre} del escalamiento de ${ticketId}`);
    }
  }

  // 2. Actualizar tipo de servicio en la Sheet
  // Buscar columna "tipo_servicio" o "Servicio"
  const headers = ctx.hoja.getRange(1, 1, 1, ctx.hoja.getLastColumn()).getValues()[0];
  let colTipo = -1;
  for (let i = 0; i < headers.length; i++) {
    const h = normalizar_(String(headers[i] || ''));
    if (h.includes('tipo') && h.includes('servicio')) {
      colTipo = i + 1;
      break;
    }
    if (h === 'servicio') {
      colTipo = i + 1;
      break;
    }
  }
  if (colTipo > 0) {
    ctx.hoja.getRange(ctx.fila, colTipo).setValue(matcheaTipo);
    log_('INFO', `Tipo de servicio actualizado a ${matcheaTipo} para ${ticketId}`);
  } else {
    log_('WARN', `No se encontró columna de tipo de servicio en la Sheet`);
  }

  // 3. Marcar estado anterior como ESCALADA
  // Lo dejamos brevemente para histórico, después de eso lo cambia el flujo normal
  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue(estadoEscalado);
  log_('INFO', `Estado anterior marcado como ${estadoEscalado}`);

  // 4. Limpiar Event_ID y Técnico previos para procesar como nuevo
  const totalCols = ctx.hoja.getLastColumn();
  log_('INFO', `escalarLlamadaORevision_: totalCols=${totalCols}, fila=${ctx.fila}, limpiando técnico (col ${totalCols - 4}) y event_id (col ${totalCols - 1})`);
  ctx.hoja.getRange(ctx.fila, totalCols - 4).setValue('');  // técnico
  ctx.hoja.getRange(ctx.fila, totalCols - 1).setValue('');  // event_id

  // 4.5. ACTUALIZAR EL TIPO DE SERVICIO en la Sheet al nuevo tipo (Reparación, etc.)
  // Esto es crítico porque procesarTicket_ va a re-leer la fila
  const headersFila = ctx.hoja.getRange(1, 1, 1, totalCols).getValues()[0];
  let colTipoServicio = -1;
  for (let i = 0; i < headersFila.length; i++) {
    const h = normalizar_(String(headersFila[i] || '')).replace(/\s+/g, ' ').trim();
    if (h === 'tipo de servicio' || h === 'tipo_de_servicio' || h === 'servicio') {
      colTipoServicio = i + 1;  // 1-indexed
      break;
    }
  }
  if (colTipoServicio > 0) {
    ctx.hoja.getRange(ctx.fila, colTipoServicio).setValue(matcheaTipo);
    log_('INFO', `escalarLlamadaORevision_: Tipo de servicio actualizado a "${matcheaTipo}" en Sheet (col ${colTipoServicio})`);
  } else {
    log_('WARN', `escalarLlamadaORevision_: no se encontró columna "Tipo de servicio" para actualizar`);
  }

  // 5. Releer el ticket de la fila actualizada y procesarlo como visita
  const ticket = leerTicketDeFila_(ctx.hoja, ctx.fila);
  ticket.tipoServicio = matcheaTipo;  // forzar también en memoria por si acaso
  ticket.id = ticketId;  // mantener el ID original

  log_('INFO', `escalarLlamadaORevision_: ticket releído. id=${ticket.id}, tipo=${ticket.tipoServicio}, direccion=${ticket.direccion}, depto=${ticket.departamento}`);

  // 6. Procesar como ticket normal de visita
  // Esto buscará técnico + slot, creará evento Calendar y mandará correo de aprobación
  try {
    procesarTicket_(ticket, ctx.hoja, ctx.fila);
    log_('INFO', `escalarLlamadaORevision_: procesarTicket_ completado para ${ticketId}`);
  } catch(e) {
    log_('ERROR', `escalarLlamadaORevision_: procesarTicket_ falló para ${ticketId}: ${e.message}`);
    throw e;
  }

  // 6b. Si jefatura forzó un técnico específico, reasignar después de procesar
  // (procesarTicket_ ya creó el evento con un técnico — ahora lo cambiamos al forzado)
  if (tecnicoForzado) {
    try {
      log_('INFO', `Reasignando ${ticketId} al técnico forzado por jefatura: ${tecnicoForzado}`);
      reasignarTicket_(ticketId, tecnicoForzado);
    } catch(e) {
      log_('ERROR', `No se pudo asignar técnico forzado ${tecnicoForzado}: ${e.message}`);
    }
  }

  // 7. Notificar al vendedor
  if (ctx.emailVendedor) {
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[TICKET ESCALADO ${ticketId}] ${ctx.cliente} - ahora ${matcheaTipo}`,
      `Tu ticket ${ticketId} fue escalado por jefatura.\n\n` +
      `Cliente:     ${ctx.cliente}\n` +
      `Equipo:      ${ctx.equipo}\n` +
      `Tipo nuevo:  ${matcheaTipo}\n\n` +
      `El sistema está buscando técnico y horario disponible. ` +
      `Recibirás otra notificación cuando jefatura apruebe la asignación.`,
      { name: 'Sistema de Tickets' }
    );
  }

  actualizarDashboard_();
}


// =====================================================================
// C) RESPUESTAS DE TÉCNICOS
// =====================================================================
function procesarRespuestasTecnicos_() {
  // Buscar todos los hilos con TCK en el asunto, no solo los que empiezan con TICKET
  // Esto es necesario porque al escalar, el hilo del correo puede ser [LLAMADA TCK...] o [APROBAR TCK...]
  const query = `subject:TCK newer_than:7d`;
  const hilos = GmailApp.search(query, 0, 30);
  console.log(`Hilos técnicos (buscando TCK): ${hilos.length}`);

  let procesados = 0;

  for (const hilo of hilos) {
    const mensajes = hilo.getMessages();
    const ultimo = mensajes[mensajes.length - 1];
    const asunto = hilo.getFirstMessageSubject();

    const matchId = asunto.match(/TCK\d+/);
    if (!matchId) continue;
    const ticketId = matchId[0];

    if (mensajes.length < 2) continue;

    const ctx = buscarTicketPorId_(ticketId);
    if (!ctx) continue;

    const estadoActual = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();

    // Si el ticket ya está cerrado, archivar el hilo y saltar
    const estadosTerminalesTec = ['RESUELTO', 'RECHAZADO', 'CANCELADO_MANUALMENTE',
                                   'LLAMADA_COMPLETADA', 'REVISION_COMPLETADA',
                                   'LLAMADA_ESCALADA', 'REVISION_ESCALADA'];
    if (estadosTerminalesTec.some(s => estadoActual.startsWith(s))) {
      try {
        hilo.moveToArchive();
        for (const m of mensajes) {
          try { m.markRead(); } catch(_) {}
        }
      } catch(_) {}
      continue;
    }

    // Procesamos respuestas con horas si el ticket está esperando horas del técnico
    // o si está en BORRADOR/BORRADOR_FORZADO (puede haberse saltado la aprobación por bug)
    const estadosProcesables = ['APROBADO_ESPERA_TECNICO', 'BORRADOR', 'BORRADOR_FORZADO'];
    if (!estadosProcesables.includes(estadoActual)) continue;

    const cuerpo = extraerCuerpoPlano_(ultimo);
    const horas = parsearHoras_(cuerpo);

    // Validar que es respuesta del técnico, no del bot ni del usuario
    const linea0 = cuerpo.split('\n')[0].trim();
    if (/^Ticket TCK\d+/i.test(linea0)) { ultimo.markRead(); continue; }
    if (/^APROBADO\b|^CAMBIAR|^RECHAZADO|^FORZAR|^RETRASAR/i.test(linea0)) { ultimo.markRead(); continue; }
    if (/^Error:|^Escalamiento de/i.test(linea0)) { ultimo.markRead(); continue; }
    if (/notificado al t[eé]cnico/i.test(linea0)) { ultimo.markRead(); continue; }
    if (/^Registrado:/i.test(linea0)) { ultimo.markRead(); continue; }

    try {
      if (horas === null) {
        // No hay horas válidas, ignorar silenciosamente
        ultimo.markRead();
        continue;
      } else if (horas <= 0 || horas > (CFG.HORA_FIN - CFG.HORA_INICIO)) {
        hilo.reply(`Horas fuera de rango (1 a ${CFG.HORA_FIN - CFG.HORA_INICIO}).`);
        ultimo.markRead();
        continue;
      } else {
        // Si estaba en BORRADOR, primero cambiar estado a APROBADO_ESPERA_TECNICO
        // para que finalizarTicketConHoras_ pueda procesarlo correctamente
        if (estadoActual === 'BORRADOR' || estadoActual === 'BORRADOR_FORZADO') {
          ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue('APROBADO_ESPERA_TECNICO');
          log_('INFO', `Auto-aprobando ${ticketId} porque técnico respondió horas en estado ${estadoActual}`);
        }
        finalizarTicketConHoras_(ticketId, horas);
        hilo.reply(`Registrado: ${horas} h de trabajo. Calendario actualizado.`);
        procesados++;
      }
    } catch (err) {
      log_('ERROR', `Técnico ${ticketId}: ${err.message}`);
      hilo.reply(`Error: ${err.message}`);
    }
    ultimo.markRead();
  }
  return procesados;
}

function parsearHoras_(texto) {
  const lineas = texto.split('\n').slice(0, 3).join(' ');
  const m = lineas.match(/(\d+(?:[.,]\d+)?)\s*(h|hr|hrs|hora|horas)?/i);
  if (!m) return null;
  return parseFloat(m[1].replace(',', '.'));
}

function finalizarTicketConHoras_(ticketId, horasTrabajo) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  const { evento, tecnico, trasladoHoras } = ctx;

  // Validar que las horas de TRABAJO no excedan el límite diario (9h)
  if (horasTrabajo > CFG.CAPACIDAD_DIARIA_HORAS) {
    throw new Error(`Las horas de trabajo (${horasTrabajo}h) exceden el límite diario de ${CFG.CAPACIDAD_DIARIA_HORAS}h.`);
  }

  // Duración del evento en Calendar = ida + trabajo + vuelta (jornada completa visible)
  const duracionEvento = trasladoHoras + horasTrabajo + trasladoHoras;

  // El inicio del evento se mantiene tal cual lo dejó crearEventoBorrador_
  // (que ya consideró la hora de salida real).
  const inicio = evento.getStartTime();
  const fin = new Date(inicio.getTime() + duracionEvento * 3600000);

  // Si el evento terminaría después de la HORA_FIN, hay que reagendar.
  // Pero solo re-evaluamos por las horas de TRABAJO, no por la jornada total.
  // (Aceptamos que regrese tarde por traslados largos.)
  const horaInicioDecimal = inicio.getHours() + inicio.getMinutes() / 60;
  const horaTrabajoInicio = horaInicioDecimal + trasladoHoras;
  const horaTrabajoFin = horaTrabajoInicio + horasTrabajo;

  if (horaTrabajoFin > CFG.HORA_FIN) {
    // El TRABAJO en sí termina después de las 5pm → necesitamos reagendar.
    const prio = ctx.prioridad || 'Media';
    const prioCfg = getConfigPrioridad_(prio);
    const slot = buscarSlotConCapacidad_(
      tecnico, inicio, horasTrabajo,
      prioCfg.ventanaDias, null, 0
    );
    if (!slot) {
      const siguiente = siguienteDiaHabil_(inicio);
      const slot2 = buscarSlotConCapacidad_(
        tecnico, siguiente, horasTrabajo,
        prioCfg.ventanaDias, null, 0
      );
      if (!slot2) throw new Error(`No hay capacidad para ${horasTrabajo}h de trabajo.`);
      aplicarSlotConTraslado_(evento, slot2, horasTrabajo, trasladoHoras);
    } else {
      aplicarSlotConTraslado_(evento, slot, horasTrabajo, trasladoHoras);
    }
  } else {
    // El trabajo cabe dentro del horario, simplemente actualizar fin del evento
    evento.setTime(inicio, fin);
  }

  // Recalcular viáticos con las horas reales
  const elegidoFake = {
    trasladoHoras,
    fechaSlot: Utilities.formatDate(evento.getStartTime(), CFG.TIMEZONE, 'yyyy-MM-dd'),
    horaInicioSlot: evento.getStartTime().getHours() + evento.getStartTime().getMinutes()/60,
    tecnico
  };
  const viaticos = calcularViaticos_(
    { departamento: ctx.departamento, horaSugerida: ctx.horaSugerida },
    elegidoFake,
    horasTrabajo
  );

  const desc = evento.getDescription() || '';
  let nuevaDesc = desc.replace(/\*\*\* PENDIENTE.*?\*\*\*/s, '').trim() +
    `\n\nHoras de trabajo: ${formatearHorasMinutos_(horasTrabajo)}` +
    `\nTraslado ida: ${formatearHorasMinutos_(trasladoHoras)}` +
    `\nTraslado vuelta: ${formatearHorasMinutos_(trasladoHoras)}` +
    `\nDuración total del evento: ${formatearHorasMinutos_(duracionEvento)} (incluye ida + trabajo + vuelta)`;

  if (viaticos.aplica) nuevaDesc += `\n${viaticos.resumen}`;
  nuevaDesc += `\n\nCONFIRMADO`;

  evento.setDescription(nuevaDesc);

  try { reoptimizarRutaDelDia_(tecnico, inicio); } catch(e) {
    log_('WARN', `Ruta no reoptimizada: ${e.message}`);
  }

  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue('CONFIRMADO');

  if (ctx.emailVendedor) {
    const lineaViaticos = viaticos.aplica ? `\n${viaticos.resumen}` : '';
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[CONFIRMADO ${ticketId}] ${ctx.cliente}`,
      `Tu ticket ${ticketId} fue confirmado:\n\n` +
      `Técnico: ${tecnico.nombre}\n` +
      `Inicio: ${Utilities.formatDate(evento.getStartTime(), CFG.TIMEZONE, 'EEEE dd MMM, HH:mm')}\n` +
      `Fin: ${Utilities.formatDate(evento.getEndTime(), CFG.TIMEZONE, 'HH:mm')}\n` +
      `Cliente: ${ctx.cliente}\n` +
      `Lugar: ${ctx.direccion}${lineaViaticos}`
    );
  }

  actualizarDashboard_();
}

function aplicarSlot_(evento, slot, duracionHoras) {
  const ini = new Date(slot.fecha);
  ini.setHours(Math.floor(slot.horaInicio),
               Math.round((slot.horaInicio % 1) * 60), 0, 0);
  const fin = new Date(ini.getTime() + duracionHoras * 3600000);
  evento.setTime(ini, fin);
}

/**
 * Igual que aplicarSlot_ pero considera traslado:
 *   - El slot.horaInicio es la hora del TRABAJO
 *   - El evento empieza en (horaInicio - traslado) (puede ser antes de 8am, hasta 5am mínimo)
 *   - El evento dura traslado + trabajo + traslado
 */
function aplicarSlotConTraslado_(evento, slot, horasTrabajo, trasladoHoras) {
  const horaSalida = Math.max(slot.horaInicio - trasladoHoras, CFG.HORA_MINIMA_SALIDA);
  const ini = new Date(slot.fecha);
  ini.setHours(Math.floor(horaSalida),
               Math.round((horaSalida % 1) * 60), 0, 0);

  const duracionTotal = trasladoHoras + horasTrabajo + trasladoHoras;
  const fin = new Date(ini.getTime() + duracionTotal * 3600000);
  evento.setTime(ini, fin);
}

// =====================================================================
// Helpers
// =====================================================================
function buscarTicketPorId_(ticketId) {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TICKETS);
  const data = hoja.getDataRange().getValues();
  const headers = data[0].map(h => normalizarHeader_(h));

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    if (fila.some(c => String(c).includes(ticketId))) {
      const t = {};
      headers.forEach((h, j) => t[h] = fila[j]);

      const base = hoja.getLastColumn() - 6;
      const tecnicoNombre = fila[base + 1];
      const trasladoHoras = parseFloat(fila[base + 2]) || 0;
      const fechaSlot = fila[base + 3];
      const eventId = fila[base + 4];

      const tecnicos = cargarTecnicos_();
      const tecnico = tecnicos.find(tc => tc.nombre === tecnicoNombre);

      let evento = null;
      let eventoVuelta = null;
      if (eventId) {
        try {
          const cal = getCalendarCentral_();
          // El eventId puede ser "id1" (visita normal) o "id1,id2" (modo HOTEL)
          const partes = String(eventId).split(',').map(s => s.trim()).filter(Boolean);
          if (partes.length > 0) {
            evento = cal.getEventById(partes[0]);
          }
          if (partes.length > 1) {
            eventoVuelta = cal.getEventById(partes[1]);
          }
        } catch(_) {}
      }

      let horaSugeridaNum = null;
      const raw = t.hora_sugerida;
      if (raw) {
        const s = String(raw).trim();
        const m = s.match(/(\d{1,2})[:.]?(\d{0,2})/);
        if (m) horaSugeridaNum = parseInt(m[1], 10) + (m[2] ? parseInt(m[2], 10) / 60 : 0);
      }

      return {
        ticketId, hoja, fila: i + 1, tecnico, evento, eventoVuelta,
        trasladoHoras, fechaSlot,
        horaInicioSlot: evento ? (evento.getStartTime().getHours()
                      + evento.getStartTime().getMinutes() / 60) : null,
        cliente: t.cliente,
        telefonoCliente: t.telefono_cliente || t.telefono || '',
        direccion: t.direccion || t.ubicacion,
        departamento: t.departamento,
        equipo: normalizarListaEquipos_(t.equipo),
        tipoServicio: t.tipo_de_servicio || t.servicio,
        prioridad: t.prioridad || 'Media',
        notas: t.notas || t.observaciones || '',
        vendedor: t.vendedor,
        emailVendedor: t.correo_vendedor || t.email_vendedor,
        fechaSugerida: t.fecha_sugerida ? new Date(t.fecha_sugerida) : null,
        horaSugerida: horaSugeridaNum
      };
    }
  }
  return null;
}

function extraerCuerpoPlano_(msg) {
  const raw = msg.getPlainBody();
  const lineas = raw.split('\n');
  const limpias = [];
  for (const l of lineas) {
    if (/^>/.test(l)) break;
    if (/^El .* escribió:/.test(l)) break;
    if (/^On .* wrote:/.test(l)) break;
    limpias.push(l);
  }
  return limpias.join('\n').trim();
}

// =====================================================================
// D) TRACKING DE LLAMADAS
// Detecta cuando el técnico (a) leyó el correo, (b) respondió HECHO.
// Envía un correo a jefatura en cada evento y actualiza la Sheet.
// =====================================================================
function procesarTrackingLlamadas_() {
  // Buscar hilos de llamadas recientes (últimos 7 días)
  const query = `subject:LLAMADA newer_than:7d`;
  const hilos = GmailApp.search(query, 0, 30);
  console.log(`Hilos de llamadas: ${hilos.length}`);

  let procesados = 0;

  for (const hilo of hilos) {
    const asunto = hilo.getFirstMessageSubject();

    // Solo procesar los correos de LLAMADA al técnico, no los de tracking a jefatura
    // Formato: "[LLAMADA TCKxxx] Atención remota - Cliente"
    if (!/^\[LLAMADA\s+TCK\d+\]/.test(asunto)) continue;

    const matchId = asunto.match(/TCK\d+/);
    if (!matchId) continue;
    const ticketId = matchId[0];

    const ctx = buscarTicketPorId_(ticketId);
    if (!ctx) continue;

    const estadoActual = String(
      ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || ''
    ).toUpperCase();

    // Solo procesar tickets remotos en estado activo (Llamada o Revisión)
    const estadosActivos = ['LLAMADA_ENVIADA', 'PENDIENTE', 'REVISION_ENVIADA'];
    if (!estadosActivos.includes(estadoActual)) continue;

    // Determinar prefijo según el tipo del ticket
    const tipoNorm = normalizar_(ctx.tipoServicio || '');
    const esRevision = tipoNorm.includes('revision');
    const prefijo = esRevision ? 'REVISION' : 'LLAMADA';

    const mensajes = hilo.getMessages();
    const primerMensaje = mensajes[0]; // el que le enviamos al técnico

    try {
      // ---------- Detectar VISTA (cambia estado a PENDIENTE, NO envía correo) ----------
      const estadoEnviada = `${prefijo}_ENVIADA`;
      const estadoVista = 'PENDIENTE';
      const estadoCompletada = 'RESUELTO';

      if (estadoActual === estadoEnviada) {
        if (!primerMensaje.isUnread()) {
          log_('INFO', `${prefijo} ${ticketId} vista por el técnico (sin notificación)`);

          // Solo actualizar estado en Sheet, sin enviar correo
          ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue(estadoVista);

          procesados++;
        }
      }

      // ---------- Detectar HECHO ----------
      // Revisar siempre (aunque acabe de pasar a VISTA) para capturar respuesta rápida
      if (mensajes.length > 1) {
        const ultimoMensaje = mensajes[mensajes.length - 1];
        const cuerpo = extraerCuerpoPlano_(ultimoMensaje);
        const primeraLinea = cuerpo.split('\n')[0].trim().toUpperCase();

        if (/^HECHO\b/.test(primeraLinea)) {
          log_('INFO', `${prefijo} ${ticketId} completada por el técnico`);

          // Detectar si el técnico sugiere escalamiento: "HECHO - REQUIERE VISITA: <tipo>"
          const matchEscalamiento = primeraLinea.match(/REQUIERE\s+VISITA\s*:\s*(.+?)(?:[\.\s]*$)/i);
          let tipoSugerido = null;
          if (matchEscalamiento) {
            tipoSugerido = matchEscalamiento[1].trim();
            log_('INFO', `${prefijo} ${ticketId} técnico SUGIERE escalamiento a: ${tipoSugerido}`);
          }

          // Marcar como COMPLETADA siempre
          ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue(estadoCompletada);

          const resumen = primeraLinea.replace(/^HECHO\s*[-:]?\s*/, '').trim() || '(sin resumen adicional)';
          const tipoReal = (ctx.tipoServicio || 'Llamada').toLowerCase();

          if (tipoSugerido) {
            // Técnico sugiere escalamiento: notificar a jefatura para aprobación
            notificarSolicitudEscalamiento_(ticketId, ctx, tipoSugerido, resumen);
            // Después de marcar COMPLETADA, ponemos estado de solicitud de escalamiento
            ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue('ESCALAMIENTO_PENDIENTE_APROBACION');
            // Guardar el tipo sugerido en una propiedad para usarlo después
            PropertiesService.getScriptProperties().setProperty(
              'escalamiento_' + ticketId,
              tipoSugerido
            );
          } else {
            // HECHO simple: notificación normal a jefatura
            notificarJefaturaLlamada_(
              ticketId,
              'COMPLETADA',
              ctx,
              `El técnico ${ctx.tecnico ? ctx.tecnico.nombre : '(desconocido)'} completó la ${tipoReal}.\n\nRespuesta del técnico: ${resumen}`
            );
          }

          ultimoMensaje.markRead();
          procesados++;
        }
      }
    } catch (err) {
      log_('ERROR', `Tracking llamada ${ticketId}: ${err.message}`);
    }
  }

  return procesados;
}

function notificarJefaturaLlamada_(ticketId, evento, ctx, detalle) {
  const tipoTexto = ctx.tipoServicio || 'Llamada';
  const tipoEtiqueta = tipoTexto.toUpperCase();

  const asunto = evento === 'VISTA'
    ? `[${tipoEtiqueta} VISTA ${ticketId}] Técnico leyó el correo`
    : `[${tipoEtiqueta} COMPLETADA ${ticketId}] Técnico finalizó ${tipoTexto.toLowerCase()}`;

  const cuerpo = `
${detalle}

-------------------------------------------------
Ticket:     ${ticketId}
Tipo:       ${tipoTexto}
Cliente:    ${ctx.cliente}
Teléfono:   ${ctx.telefonoCliente || '(no especificado)'}
Equipo:     ${ctx.equipo}
Técnico:    ${ctx.tecnico ? ctx.tecnico.nombre : '(desconocido)'}
Vendedor:   ${ctx.vendedor}
-------------------------------------------------

Este es un mensaje automático del sistema de tickets.
`;

  // 1) Notificar a jefatura
  GmailApp.sendEmail(CFG.EMAIL_JEFATURA, asunto, cuerpo, {
    name: 'Sistema de Tickets'
  });

  // 2) Notificar también al vendedor
  if (ctx.emailVendedor) {
    log_('INFO', `Notificando ${evento} a vendedor ${ctx.emailVendedor} para ${ticketId}`);
    GmailApp.sendEmail(ctx.emailVendedor, asunto, cuerpo, {
      name: 'Sistema de Tickets'
    });
  } else {
    log_('WARN', `Ticket ${ticketId}: emailVendedor vacío, NO se notifica al vendedor`);
  }
}

// =====================================================================
// SOLICITUD DE ESCALAMIENTO (técnico sugiere → jefatura aprueba)
// =====================================================================

/**
 * Notifica a jefatura que el técnico sugiere escalar la llamada/revisión.
 * Jefatura debe responder APROBADO o RECHAZADO.
 */
function notificarSolicitudEscalamiento_(ticketId, ctx, tipoSugerido, resumenTecnico) {
  const tipoActual = ctx.tipoServicio || 'Llamada';
  const tecnicoNombre = ctx.tecnico ? ctx.tecnico.nombre : '(desconocido)';

  const asunto = `[ESCALAMIENTO SOLICITADO ${ticketId}] ${ctx.cliente} - ${tipoActual} → ${tipoSugerido}`;

  const cuerpo = `
El técnico ${tecnicoNombre} sugiere escalar este ticket a visita.

-------------------------------------------------
Ticket:          ${ticketId}
Cliente:         ${ctx.cliente}
Teléfono:        ${ctx.telefonoCliente || '(no especificado)'}
Dirección:       ${ctx.direccion}
Departamento:    ${ctx.departamento}
Equipo:          ${ctx.equipo}
-------------------------------------------------
Tipo actual:     ${tipoActual} (ya completada)
Tipo sugerido:   ${tipoSugerido}
Técnico:         ${tecnicoNombre}
-------------------------------------------------

Resumen del técnico:
${resumenTecnico}

-------------------------------------------------

CÓMO RESPONDER (primera línea del correo):

  APROBADO                       → procesa como visita y asigna técnico óptimo
  APROBADO: <nombre del técnico> → procesa como visita asignando ese técnico
  CAMBIAR: <nombre del técnico>  → equivalente a APROBADO: <nombre>
  RECHAZADO: motivo              → no se escala, queda completada

Mantén el asunto intacto.
`;

  GmailApp.sendEmail(CFG.EMAIL_JEFATURA, asunto, cuerpo, {
    name: 'Sistema de Tickets'
  });

  // También notificar al vendedor que se solicita escalamiento
  if (ctx.emailVendedor) {
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[ESCALAMIENTO SOLICITADO ${ticketId}] ${ctx.cliente}`,
      `El técnico ${tecnicoNombre} completó la llamada/revisión pero sugiere visita presencial.\n\n` +
      `Cliente:        ${ctx.cliente}\n` +
      `Tipo sugerido:  ${tipoSugerido}\n\n` +
      `Resumen del técnico:\n${resumenTecnico}\n\n` +
      `Jefatura decidirá si se aprueba el escalamiento. Te notificaremos del resultado.`,
      { name: 'Sistema de Tickets' }
    );
  }
}

/**
 * Jefatura aprueba el escalamiento. Procesa como visita normal.
 * Si se especifica tecnicoForzado, se asigna ese técnico en vez del que sugiere el sistema.
 */
function aprobarEscalamientoTecnico_(ticketId, tecnicoForzado) {
  // Recuperar el tipo sugerido que se guardó al detectar la solicitud
  const tipoSugerido = PropertiesService.getScriptProperties()
                       .getProperty('escalamiento_' + ticketId);
  if (!tipoSugerido) {
    throw new Error(`No se encontró tipo de escalamiento para ${ticketId}`);
  }

  log_('INFO', `Aprobando escalamiento de ${ticketId} a ${tipoSugerido}` +
               (tecnicoForzado ? ` con técnico forzado: ${tecnicoForzado}` : ''));

  // Validar técnico forzado si se especificó
  if (tecnicoForzado) {
    const tecnicos = cargarTecnicos_();
    const tecValido = tecnicos.find(t =>
      normalizar_(t.nombre) === normalizar_(tecnicoForzado) && t.activo
    );
    if (!tecValido) {
      throw new Error(`Técnico "${tecnicoForzado}" no encontrado o inactivo`);
    }
  }

  // Reusar la función existente para escalar, pasando el técnico forzado si lo hay
  escalarLlamadaORevision_(ticketId, tipoSugerido, tecnicoForzado);

  // Limpiar la propiedad temporal
  PropertiesService.getScriptProperties().deleteProperty('escalamiento_' + ticketId);
}

/**
 * Jefatura rechaza el escalamiento. El ticket queda como completado.
 */
function rechazarEscalamientoTecnico_(ticketId, motivo) {
  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) throw new Error(`No se encontró ${ticketId}`);

  log_('INFO', `Rechazando escalamiento de ${ticketId}: ${motivo}`);

  // Volver al estado RESUELTO (en vez de ESCALAMIENTO_PENDIENTE_APROBACION)
  const estadoFinal = 'RESUELTO';

  ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).setValue(estadoFinal);

  // Limpiar tipo sugerido guardado
  PropertiesService.getScriptProperties().deleteProperty('escalamiento_' + ticketId);

  // Notificar al técnico
  if (ctx.tecnico && ctx.tecnico.nombre) {
    const tecnicos = cargarTecnicos_();
    const tec = tecnicos.find(t => t.nombre === ctx.tecnico.nombre);
    if (tec && tec.email) {
      GmailApp.sendEmail(
        tec.email,
        `[ESCALAMIENTO RECHAZADO ${ticketId}] ${ctx.cliente}`,
        `Jefatura rechazó la solicitud de escalamiento.\n\n` +
        `Cliente: ${ctx.cliente}\n` +
        `Motivo:  ${motivo}\n\n` +
        `El ticket queda como completado.`,
        { name: 'Sistema de Tickets' }
      );
    }
  }

  // Notificar al vendedor
  if (ctx.emailVendedor) {
    GmailApp.sendEmail(
      ctx.emailVendedor,
      `[ESCALAMIENTO RECHAZADO ${ticketId}] ${ctx.cliente}`,
      `Jefatura rechazó el escalamiento sugerido por el técnico.\n\n` +
      `Motivo: ${motivo}\n\n` +
      `El ticket queda como completado.`,
      { name: 'Sistema de Tickets' }
    );
  }

  actualizarDashboard_();
}
