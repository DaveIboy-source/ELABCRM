/**
 * =====================================================================
 * CODIGO.gs — v5 — llamada + viáticos + sin emojis
 * =====================================================================
 */

function onFormSubmit(e) {
  // Protección: si se corre manualmente desde el editor, e es undefined
  if (!e || !e.range) {
    const msg = 'onFormSubmit NO se ejecuta manualmente. Se dispara automáticamente cuando alguien envía el Google Form.\n\n' +
                'Para probar sin enviar Form, corre la función: testSimularUltimoTicket()';
    console.log(msg);
    try {
      SpreadsheetApp.getUi().alert('Información', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(_) {
      // No hay UI (ejecución desde editor sin Sheet abierta)
    }
    return;
  }

  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TICKETS);
  const fila = e.range.getRow();

  // GUARDA ANTI-DUPLICADO: si la fila ya tiene TicketID asignado, NO procesar de nuevo
  // Esto evita que se mande correo de aprobación múltiples veces si el trigger se dispara dos veces
  const totalCols = hoja.getLastColumn();
  const ticketIdExistente = hoja.getRange(fila, totalCols - 5).getValue();
  if (ticketIdExistente && String(ticketIdExistente).startsWith('TCK')) {
    log_('WARN', `onFormSubmit: fila ${fila} YA tiene TicketID "${ticketIdExistente}". Ignorando para evitar duplicado.`);
    return;
  }

  const ticket = leerTicketDeFila_(hoja, fila);
  log_('INFO', `Ticket recibido ${ticket.id} - ${ticket.cliente} - ${ticket.tipoServicio} - ${ticket.prioridad}`);

  try {
    // Tipos remotos (sin agenda, correo inmediato): Llamada y Revisión Equipo
    // Normalizar y colapsar espacios múltiples
    const tipoNorm = normalizar_(ticket.tipoServicio).replace(/\s+/g, ' ').trim();
    log_('INFO', `Ticket ${ticket.id}: tipoServicio original="${ticket.tipoServicio}", normalizado="${tipoNorm}"`);

    // Lista EXACTA de tipos remotos (deben coincidir con el Form)
    // Solo "Llamada" y "Revisión Equipo" son remotos
    // TODOS los demás son visitas con agenda
    const tiposRemotos = ['llamada', 'revision equipo'];
    const esRemoto = tiposRemotos.includes(tipoNorm);

    log_('INFO', `Ticket ${ticket.id}: clasificado como ${esRemoto ? 'REMOTO (sin agenda)' : 'VISITA (con agenda)'}`);

    if (esRemoto) {
      procesarLlamada_(ticket, hoja, fila);
    } else {
      procesarTicket_(ticket, hoja, fila);
    }
  } catch (err) {
    log_('ERROR', `Ticket ${ticket.id}: ${err.message}\n${err.stack}`);
    notificarErrorJefatura_(ticket, err);
  }
}

// =====================================================================
// FLUJO DE LLAMADA / REVISIÓN
// Ahora requiere aprobación de jefatura ANTES de mandar al técnico
// =====================================================================
function procesarLlamada_(ticket, hoja, fila) {
  if (!ticket.telefonoCliente || !String(ticket.telefonoCliente).trim()) {
    throw new Error('Las llamadas requieren teléfono del cliente.');
  }

  const tecnicos = cargarTecnicos_().filter(t => t.activo && t.email);
  if (tecnicos.length === 0) throw new Error('No hay técnicos activos.');

  // Rotación: excluir al técnico que recibió la última llamada/revisión
  const ultimoTecnicoLlamada = obtenerUltimoTecnicoDeLlamada_(hoja, ticket.id);
  console.log(`Último técnico de llamada/revisión: ${ultimoTecnicoLlamada || '(ninguno)'}`);

  let candidatos = tecnicos;
  if (ultimoTecnicoLlamada && tecnicos.length > 1) {
    const filtrados = tecnicos.filter(t => t.nombre !== ultimoTecnicoLlamada);
    if (filtrados.length > 0) candidatos = filtrados;
  }

  // Sugerir el menos cargado entre los candidatos (carga real = eventos + llamadas activas)
  const hoy = new Date();

  let sugerido = null;
  let menorCarga = Infinity;
  const cargasPorTecnico = {};

  for (const tec of candidatos) {
    const cargaReal = obtenerCargaRealTecnico_(tec, hoy);
    const llamadasActivas = contarLlamadasActivas_(tec.nombre, hoy);
    cargasPorTecnico[tec.nombre] = {
      total: cargaReal,
      llamadas: llamadasActivas
    };
    console.log(`  ${tec.nombre}: ${cargaReal.toFixed(1)}h carga real (${llamadasActivas} llamadas activas)`);
    if (cargaReal < menorCarga) {
      menorCarga = cargaReal;
      sugerido = tec;
    }
  }

  log_('INFO', `Remoto ${ticket.id} sugerido: ${sugerido.nombre} (${menorCarga.toFixed(1)}h)`);

  // Determinar tipo de estado
  const tipoNorm = normalizar_(ticket.tipoServicio);
  const esRevision = tipoNorm.includes('revision');
  const estadoInicial = esRevision ? 'REVISION_PENDIENTE_APROBACION' : 'LLAMADA_PENDIENTE_APROBACION';

  // DEFENSA: si por alguna razón ya existe un evento Calendar vinculado a esta fila,
  // borrarlo porque las llamadas/revisiones NO deben tener evento
  const totalCols = hoja.getLastColumn();
  const eventIdExistente = hoja.getRange(fila, totalCols + COL_EVENT_ID_OFFSET).getValue();
  if (eventIdExistente) {
    log_('WARN', `procesarLlamada_: ${ticket.id} ya tenía Event_ID "${eventIdExistente}", borrando evento Calendar`);
    try {
      const cal = getCalendarCentral_();
      const idsLista = String(eventIdExistente).split(',').map(s => s.trim()).filter(Boolean);
      for (const id of idsLista) {
        try {
          const ev = cal.getEventById(id);
          if (ev) ev.deleteEvent();
        } catch(e) {
          log_('WARN', `No se pudo borrar evento ${id}: ${e.message}`);
        }
      }
    } catch(e) {
      log_('WARN', `Error al limpiar eventos huérfanos: ${e.message}`);
    }
  }

  // Guardar en Sheet con técnico sugerido (no enviar al técnico aún)
  // IMPORTANTE: Event_ID siempre vacío para llamadas/revisiones
  hoja.getRange(fila, hoja.getLastColumn() - 5, 1, 6).setValues([[
    ticket.id,
    sugerido.nombre,
    '0',
    Utilities.formatDate(new Date(), CFG.TIMEZONE, 'yyyy-MM-dd'),
    '',
    estadoInicial
  ]]);

  // Enviar correo de aprobación a JEFATURA
  enviarCorreoAprobacionLlamada_(ticket, sugerido, cargasPorTecnico);

  // Confirmar al vendedor que el ticket fue creado
  if (ticket.emailVendedor) {
    const tipoTexto = ticket.tipoServicio || 'Llamada';
    const tipoEtiqueta = tipoTexto.toUpperCase();
    GmailApp.sendEmail(
      ticket.emailVendedor,
      `[${tipoEtiqueta} CREADA ${ticket.id}] ${ticket.cliente}`,
      `Tu solicitud de ${tipoTexto.toLowerCase()} fue creada y está esperando aprobación de jefatura.\n\n` +
      `Técnico sugerido: ${sugerido.nombre}\n` +
      `Cliente: ${ticket.cliente}\n` +
      `Teléfono: ${ticket.telefonoCliente}\n\n` +
      `Cuando jefatura apruebe, el técnico será notificado para contactar al cliente.`,
      { name: 'Sistema de Tickets' }
    );
  } else {
    log_('WARN', `Ticket ${ticket.id}: emailVendedor vacío`);
  }

  actualizarDashboard_();
}

// =====================================================================
// Correo de aprobación para Llamada/Revisión a jefatura
// =====================================================================
function enviarCorreoAprobacionLlamada_(ticket, sugerido, cargasPorTecnico) {
  const tipoTexto = ticket.tipoServicio || 'Llamada';
  const tipoEtiqueta = tipoTexto.toUpperCase();
  const asunto = `[APROBAR ${tipoEtiqueta} ${ticket.id}] [${ticket.prioridad}] ${ticket.cliente} -> ${sugerido.nombre}`;

  // Lista de cargas por técnico para mostrar contexto
  const lineasCargas = Object.entries(cargasPorTecnico)
    .map(([nombre, info]) => {
      if (typeof info === 'object') {
        return `  ${nombre}: ${info.total.toFixed(1)}h total (${info.llamadas} llamadas/revisiones activas)`;
      }
      return `  ${nombre}: ${info.toFixed(1)}h ocupadas`;
    })
    .join('\n');

  const cuerpo = `
Nueva solicitud de ${tipoTexto} requiere aprobación.

-------------------------------------------------
Ticket:        ${ticket.id}
Tipo:          ${tipoTexto}
Prioridad:     ${ticket.prioridad}
Vendedor:      ${ticket.vendedor}
Cliente:       ${ticket.cliente}
Teléfono:      ${ticket.telefonoCliente}
Equipo:        ${ticket.equipo}
Notas:         ${ticket.notas || '(sin notas)'}
-------------------------------------------------
Técnico sugerido: ${sugerido.nombre}

Carga del día (técnicos disponibles):
${lineasCargas}
-------------------------------------------------

CÓMO RESPONDER (primera línea del correo):

  APROBADO
  CAMBIAR: <nombre del técnico>
  RECHAZADO: <motivo>

Mantén el asunto intacto.
`;

  GmailApp.sendEmail(CFG.EMAIL_JEFATURA, asunto, cuerpo, {
    name: 'Sistema de Tickets'
  });
}

function enviarCorreoLlamada_(ticket, tecnico) {
  const tipoTexto = ticket.tipoServicio || 'Llamada';
  // Asunto usa etiqueta [LLAMADA] como identificador interno para tracking,
  // pero muestra el tipo real en el título visible
  const asunto = `[LLAMADA ${ticket.id}] ${tipoTexto} - ${ticket.cliente}`;
  const cuerpo = `
Hola ${tecnico.nombre},

Tienes una solicitud de ${tipoTexto.toLowerCase()} / atención remota. Contacta al cliente lo antes posible.

-------------------------------------------------
Ticket:     ${ticket.id}
Tipo:       ${tipoTexto}
Cliente:    ${ticket.cliente}
Teléfono:   ${ticket.telefonoCliente}
Equipo:     ${ticket.equipo}
Vendedor:   ${ticket.vendedor}
Notas:      ${ticket.notas || '(sin notas)'}
-------------------------------------------------

Esta es una atención remota. No requiere visita en sitio ni agenda en calendario.

IMPORTANTE: cuando termines, responde a este correo con UNA de estas opciones (primera línea):

  HECHO
       → marca como completada simple

  HECHO: <resumen>
       → completada con tu resumen del caso
       Ejemplo: HECHO: problema resuelto, equipo funcionando correctamente

  HECHO - REQUIERE VISITA: <tipo>
       → si el caso no se pudo resolver y necesita visita presencial
       Tipos: Reparación, Mantenimiento, Instalación/Capacitación, Capacitación, Desinstalación/Retiro, Visita técnica
       Ejemplo: HECHO - REQUIERE VISITA: Reparación

Jefatura debe aprobar la visita antes de que se agende.
Mantén el asunto intacto.
`;

  GmailApp.sendEmail(tecnico.email, asunto, cuerpo, {
    name: 'Sistema de Tickets'
  });
}

// =====================================================================
// FLUJO NORMAL
// =====================================================================
function procesarTicket_(ticket, hoja, fila) {
  // GUARDA ANTI-DUPLICADO: si la fila ya tiene Event_ID, NO crear evento nuevo
  // (alguien ya procesó este ticket)
  const totalColsCheck = hoja.getLastColumn();
  const eventIdYa = hoja.getRange(fila, totalColsCheck - 1).getValue();
  if (eventIdYa && String(eventIdYa).trim()) {
    log_('WARN', `procesarTicket_: ${ticket.id} ya tiene Event_ID "${eventIdYa}". Ignorando para evitar duplicado.`);
    return;
  }

  // DEFENSA: si por alguna razón un ticket remoto llega aquí, redirigir
  const tipoNormDef = normalizar_(ticket.tipoServicio).replace(/\s+/g, ' ').trim();
  if (tipoNormDef === 'llamada' || tipoNormDef === 'revision equipo') {
    log_('WARN', `procesarTicket_: ticket ${ticket.id} es REMOTO (${tipoNormDef}), redirigiendo a procesarLlamada_`);
    procesarLlamada_(ticket, hoja, fila);
    return;
  }

  const destino = geocodificar_(ticket.direccion, ticket.departamento);
  ticket.lat = destino.lat;
  ticket.lng = destino.lng;

  const tecnicos = cargarTecnicos_();
  const resultado = evaluarCandidatos_(ticket, tecnicos);

  if (ticket.prioridad === 'Urgente' && resultado.candidatosDirectos.length === 0) {
    if (resultado.candidatosConDesplazamiento.length > 0) {
      manejarUrgenteConDesplazamiento_(ticket, hoja, fila, resultado.candidatosConDesplazamiento);
      return;
    } else {
      manejarUrgenteSinOpciones_(ticket, hoja, fila);
      return;
    }
  }

  const candidatos = resultado.candidatosDirectos;
  if (candidatos.length === 0) {
    throw new Error('Ningún técnico disponible con capacidad.');
  }

  const elegido = candidatos[0];
  log_('INFO', `Ticket ${ticket.id} -> ${elegido.tecnico.nombre} (traslado ${elegido.trasladoHoras.toFixed(2)}h)`);

  // Calcular viáticos ANTES de crear evento para incluirlos en la descripción
  const viaticos = calcularViaticos_(ticket, elegido, 1);

  const evento = crearEventoBorrador_(ticket, elegido, viaticos);

  // Si es modo HOTEL, guardar AMBOS Event_IDs separados por coma
  const eventIdParaSheet = evento._eventoVueltaId
    ? `${evento.getId()},${evento._eventoVueltaId}`
    : evento.getId();

  hoja.getRange(fila, hoja.getLastColumn() - 5, 1, 6).setValues([[
    ticket.id,
    elegido.tecnico.nombre,
    elegido.trasladoHoras.toFixed(2),
    elegido.fechaSlot,
    eventIdParaSheet,
    'BORRADOR'
  ]]);

  enviarCorreoAprobacion_(ticket, elegido, evento, viaticos);
  actualizarDashboard_();
}

function leerTicketDeFila_(hoja, fila) {
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const valores = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
  const t = {};
  headers.forEach((h, i) => t[normalizarHeader_(h)] = valores[i]);

  let horaSugeridaNum = null;
  const raw = t.hora_sugerida;
  if (raw) {
    const s = String(raw).trim();
    const m = s.match(/(\d{1,2})[:.]?(\d{0,2})/);
    if (m) horaSugeridaNum = parseInt(m[1], 10) + (m[2] ? parseInt(m[2], 10) / 60 : 0);
  }

  return {
    id: 'TCK' + Utilities.formatDate(new Date(t.marca_temporal || new Date()),
                                      CFG.TIMEZONE, 'yyyyMMddHHmmss'),
    timestamp: t.marca_temporal,
    vendedor: t.vendedor,
    emailVendedor: t.correo_vendedor || t.email_vendedor,
    cliente: t.cliente,
    telefonoCliente: t.telefono_cliente || t.telefono || '',
    direccion: t.direccion || t.ubicacion,
    departamento: t.departamento,
    equipo: normalizarListaEquipos_(t.equipo),
    tipoServicio: t.tipo_de_servicio || t.servicio,
    prioridad: t.prioridad || 'Media',
    fechaSugerida: t.fecha_sugerida ? new Date(t.fecha_sugerida) : null,
    horaSugerida: horaSugeridaNum,
    notas: t.notas || t.observaciones || ''
  };
}

function normalizarHeader_(h) {
  return String(h).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_|_$/g, '');
}

/**
 * Normaliza el campo Equipo (que ahora puede ser uno o varios separados por coma).
 * Devuelve un string limpio: "MQ60, SD-1, Centurion"
 */
function normalizarListaEquipos_(valor) {
  if (!valor) return '';
  return String(valor)
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .join(', ');
}

function evaluarCandidatos_(ticket, tecnicos) {
  const candidatosDirectos = [];
  const candidatosConDesplazamiento = [];
  const prioCfg = getConfigPrioridad_(ticket.prioridad);

  for (const tec of tecnicos) {
    if (!tec.activo || !tec.email) continue;

    if (tec.departamentos.length > 0 && ticket.departamento) {
      const match = tec.departamentos.some(d =>
        normalizar_(d) === normalizar_(ticket.departamento));
      if (!match) continue;
    }

    // Determinar ORIGEN dinámico según última parada conocida del técnico
    // en la fecha en que se agendaría el ticket.
    const fechaBase = determinarFechaBase_(ticket.fechaSugerida);
    const origen = obtenerUbicacionActualTecnico_(tec, fechaBase);
    console.log(`  [origen] ${tec.nombre}: ${origen}`);

    const traslado = distanceMatrix_(origen, `${ticket.lat},${ticket.lng}`);
    if (!traslado) continue;

    const trasladoHoras = traslado.duracionSegundos / 3600;
    const idaVueltaHoras = trasladoHoras * 2;

    // Modo HOTEL: si ida supera el umbral, se necesitan 2 días (cliente + dormida)
    const modoHotel = trasladoHoras > CFG.UMBRAL_HOTEL_HORAS;

    // Duración del slot a buscar = SOLO el trabajo (no traslado).
    // Al inicio asumimos 1h placeholder; cuando el técnico responda con horas reales
    // se ajusta el evento. Los traslados solo aparecen en la descripción.
    const duracionEstimada = 1;

    // Para tickets urgentes, ignoramos eventos en BORRADOR (no son trabajo confirmado todavía)
    const opcionesSlot = ticket.prioridad === 'Urgente' ? { ignorarBorradores: true } : {};

    const slot = buscarSlotConCapacidad_(
      tec, fechaBase, duracionEstimada,
      prioCfg.ventanaDias,
      ticket.horaSugerida,
      prioCfg.toleranciaHoraSugerida,
      opcionesSlot
    );

    // Si modo hotel, también verificar que el día siguiente tenga espacio para la vuelta
    let slotVuelta = null;
    if (slot && modoHotel) {
      const diaVuelta = siguienteDiaHabil_(slot.fecha);
      slotVuelta = buscarSlotConCapacidad_(
        tec, diaVuelta, trasladoHoras,  // solo vuelta
        2,  // ventana corta, debe ser al día siguiente o subsiguiente
        null, 0,
        opcionesSlot
      );
      if (!slotVuelta) {
        log_('WARN', `${tec.nombre}: día 1 OK pero no hay espacio para regreso`);
        continue;
      }
    }

    if (slot) {
      const diasDemora = Math.floor((slot.fecha - fechaBase) / (1000 * 60 * 60 * 24));

      // Carga REAL del técnico en el día sugerido (eventos + llamadas activas × 0.5h)
      const horasOcupadasDia = obtenerCargaRealTecnico_(tec, slot.fecha);

      // Score base: traslado + demora + carga del día
      let score = trasladoHoras * 2
                + diasDemora * prioCfg.pesoDemora
                + horasOcupadasDia * 1.5;  // 1.5 = peso moderado a la carga

      // Penalización por hora sugerida lejana
      if (ticket.horaSugerida) {
        score += slot.distanciaHoraSugerida * 3;
      }

      // BONUS FUERTE por agrupamiento: si el técnico estará físicamente
      // en el mismo departamento dentro de los próximos 5 días, GANA
      // sobre cualquier otro técnico (ahorra viaje completo).
      const infoZona = tecnicoEstaraEnZona_(tec, slot.fecha, ticket.departamento);
      if (infoZona) {
        // Penalización negativa muy alta para asegurar que gane
        score -= 100;
        log_('INFO', `${tec.nombre}: TECNICO YA EN ZONA ${ticket.departamento} (-100 score)`);
      }

      console.log(`  Score ${tec.nombre}: ${score.toFixed(2)} ` +
                  `[traslado=${trasladoHoras.toFixed(1)}, demora=${diasDemora}d, ` +
                  `carga=${horasOcupadasDia.toFixed(1)}h${infoZona ? ', EN_ZONA' : ''}]`);

      candidatosDirectos.push({
        tecnico: tec,
        trasladoHoras,
        idaVueltaHoras,
        fechaSlot: Utilities.formatDate(slot.fecha, CFG.TIMEZONE, 'yyyy-MM-dd'),
        horaInicioSlot: slot.horaInicio,
        distanciaKm: traslado.distanciaMetros / 1000,
        origenUsado: origen,
        modoHotel,
        infoZona: infoZona,
        fechaVuelta: slotVuelta ? Utilities.formatDate(slotVuelta.fecha, CFG.TIMEZONE, 'yyyy-MM-dd') : null,
        horaVuelta: slotVuelta ? slotVuelta.horaInicio : null,
        score
      });
    } else if (ticket.prioridad === 'Urgente' && prioCfg.permiteDesplazarMediaBaja) {
      const inicioVentana = new Date(fechaBase);
      const finVentana = new Date(fechaBase);
      finVentana.setDate(finVentana.getDate() + prioCfg.ventanaDias);
      finVentana.setHours(23, 59, 59, 999);

      const movibles = buscarTicketsMovibles_(tec, inicioVentana, finVentana);
      if (movibles.length > 0) {
        candidatosConDesplazamiento.push({
          tecnico: tec,
          trasladoHoras,
          idaVueltaHoras,
          distanciaKm: traslado.distanciaMetros / 1000,
          movibles,
          score: trasladoHoras * 2 + 100
        });
      }
    }
  }

  candidatosDirectos.sort((a, b) => a.score - b.score);
  candidatosConDesplazamiento.sort((a, b) => a.score - b.score);

  return { candidatosDirectos, candidatosConDesplazamiento };
}

function determinarFechaBase_(fechaSugerida) {
  const ahora = new Date();
  const hoy = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());

  // Respetar fecha sugerida tal cual (puede ser dentro de 1 día o 2 semanas)
  let base = fechaSugerida
    ? new Date(fechaSugerida.getFullYear(), fechaSugerida.getMonth(), fechaSugerida.getDate())
    : hoy;

  // Solo bloqueo: no permitir fechas pasadas
  if (base < hoy) base = hoy;

  // Si la fecha base es HOY y ya pasó la hora límite, avanzar al siguiente día hábil
  if (base.getTime() === hoy.getTime()) {
    const horaLimite = CFG.HORA_INICIO + (CFG.MARGEN_MINUTOS_INICIO_HOY / 60);
    const horaActual = ahora.getHours() + ahora.getMinutes() / 60;
    if (horaActual >= horaLimite) {
      base = siguienteDiaHabil_(base);
    }
  }

  // Asegurar que sea día hábil (si cae sábado/domingo, mover al lunes siguiente)
  while (!CFG.DIAS_LABORALES.includes(base.getDay())) {
    base = siguienteDiaHabil_(base);
  }
  return base;
}

function siguienteDiaHabil_(fecha) {
  const d = new Date(fecha);
  do {
    d.setDate(d.getDate() + 1);
  } while (!CFG.DIAS_LABORALES.includes(d.getDay()));
  return d;
}

function crearEventoBorrador_(ticket, elegido, viaticos) {
  // DEFENSA CRÍTICA: nunca crear evento Calendar para llamadas/revisiones de equipo
  const tipoNormCheck = normalizar_(ticket.tipoServicio || '').replace(/\s+/g, ' ').trim();
  if (tipoNormCheck === 'llamada' || tipoNormCheck === 'revision equipo') {
    // Capturar stack trace para identificar quién llamó esto
    let stack = '';
    try { throw new Error('trace'); } catch(e) { stack = e.stack || ''; }
    const msg = `BUG INTERNO: crearEventoBorrador_ llamado con tipo remoto "${ticket.tipoServicio}" para ticket ${ticket.id}`;
    log_('ERROR', msg + '\nStack: ' + stack);
    throw new Error(msg);
  }

  const cal = getCalendarCentral_();
  const prioCfg = getConfigPrioridad_(ticket.prioridad);

  const fecha = new Date(elegido.fechaSlot + 'T00:00:00');
  const inicio = new Date(fecha);
  inicio.setHours(Math.floor(elegido.horaInicioSlot),
                  Math.round((elegido.horaInicioSlot % 1) * 60), 0, 0);

  // Duración del evento en Calendar = ida + trabajo (placeholder 1h) + vuelta
  // El bloque completo refleja la jornada del técnico desde que sale hasta que regresa.
  // Cuando el técnico responda con horas reales de trabajo, se ajusta la duración.
  const trabajoPlaceholder = 1;
  const duracionHoras = elegido.trasladoHoras + trabajoPlaceholder + elegido.trasladoHoras;
  const tituloSuffix = elegido.modoHotel ? ' [HOTEL DIA 1/2]' : '';

  // Calcular hora de salida = horaInicioSlot - traslado de ida (puede ser antes de las 8am)
  const horaSalidaDecimal = elegido.horaInicioSlot - elegido.trasladoHoras;
  const horaSalidaFinal = Math.max(horaSalidaDecimal, CFG.HORA_MINIMA_SALIDA);

  // Reemplazar inicio con la hora de salida real
  inicio.setHours(Math.floor(horaSalidaFinal),
                  Math.round((horaSalidaFinal % 1) * 60), 0, 0);

  const fin = new Date(inicio.getTime() + duracionHoras * 3600 * 1000);

  const titulo = `[BORRADOR] [${ticket.prioridad}] ${ticket.cliente} - ${ticket.equipo} (${ticket.tipoServicio}) - ${elegido.tecnico.nombre}${tituloSuffix}`;

  const lineasDesc = [
    `Ticket: ${ticket.id}`,
    `Prioridad: ${ticket.prioridad}`,
    `Técnico asignado: ${elegido.tecnico.nombre} (${elegido.tecnico.email})`,
    `Vendedor: ${ticket.vendedor}`,
    `Cliente: ${ticket.cliente}`,
    `Teléfono cliente: ${ticket.telefonoCliente || '(no especificado)'}`,
    `Dirección: ${ticket.direccion}`,
    `Departamento: ${ticket.departamento}`,
    `Equipo: ${ticket.equipo}`,
    `Servicio: ${ticket.tipoServicio}`,
    `Traslado estimado (ida): ${formatearHorasMinutos_(elegido.trasladoHoras)}`,
    `Distancia: ${elegido.distanciaKm.toFixed(1)} km`,
    `Hora sugerida vendedor: ${ticket.horaSugerida || 'no especificada'}`,
    `Modo: ${elegido.modoHotel ? 'HOTEL (2 días)' : 'NORMAL (1 día)'}`,
    `Notas: ${ticket.notas}`
  ];

  if (viaticos && viaticos.aplica) {
    lineasDesc.push(viaticos.resumen);
  }

  lineasDesc.push('', '*** PENDIENTE DE APROBACIÓN Y DE HORAS DE TRABAJO ***');

  const ev = cal.createEvent(titulo, inicio, fin, {
    description: lineasDesc.join('\n'),
    location: ticket.direccion,
    guests: elegido.tecnico.email,
    sendInvites: false
  });

  try { ev.setColor(prioCfg.colorCalendar || CFG.COLOR_BORRADOR); } catch(_) {}

  // Si es modo hotel, crear evento de regreso día 2
  if (elegido.modoHotel && elegido.fechaVuelta && elegido.horaVuelta !== null) {
    const fechaV = new Date(elegido.fechaVuelta + 'T00:00:00');
    const inicioV = new Date(fechaV);
    inicioV.setHours(Math.floor(elegido.horaVuelta),
                     Math.round((elegido.horaVuelta % 1) * 60), 0, 0);
    const finV = new Date(inicioV.getTime() + elegido.trasladoHoras * 3600 * 1000);

    const tituloV = `[BORRADOR] [${ticket.prioridad}] REGRESO - ${ticket.cliente} - ${elegido.tecnico.nombre} [HOTEL DIA 2/2]`;
    const descV = [
      `Ticket: ${ticket.id}`,
      `Parte: REGRESO (día 2 de 2)`,
      `Técnico: ${elegido.tecnico.nombre}`,
      `Desde: ${ticket.direccion}`,
      `Hacia: base (${elegido.tecnico.base || CFG.ORIGEN_DEFAULT})`,
      `Traslado: ${formatearHorasMinutos_(elegido.trasladoHoras)}`,
      '',
      'Este evento es el regreso del ticket principal del día anterior.'
    ].join('\n');

    const evV = cal.createEvent(tituloV, inicioV, finV, {
      description: descV,
      location: elegido.tecnico.base || CFG.ORIGEN_DEFAULT,
      guests: elegido.tecnico.email,
      sendInvites: false
    });
    try { evV.setColor(prioCfg.colorCalendar || CFG.COLOR_BORRADOR); } catch(_) {}

    // Vincular ambos eventos: guardar ID del evento de vuelta también en descripción del principal
    ev.setDescription(ev.getDescription() + `\n\nEvento_Vuelta_ID: ${evV.getId()}`);

    // Guardar también el ID del principal en la descripción del de vuelta
    evV.setDescription(evV.getDescription() + `\n\nEvento_Principal_ID: ${ev.getId()}`);

    // Marcar que este ticket tiene 2 eventos vinculados (para que se guarden ambos IDs en Sheet)
    ev._eventoVueltaId = evV.getId();
  }

  return ev;
}

function enviarCorreoAprobacion_(ticket, elegido, evento, viaticos) {
  const asunto = `[APROBAR ${ticket.id}] [${ticket.prioridad}] ${ticket.cliente} -> ${elegido.tecnico.nombre}`;

  const lineaViaticos = (viaticos && viaticos.aplica) ? `\n${viaticos.resumen}\n` : '';

  // Alerta de agrupamiento si el técnico ya estará en la zona
  let alertaZona = '';
  if (elegido.infoZona) {
    const fechaInicio = Utilities.formatDate(elegido.infoZona.fechaInicio, CFG.TIMEZONE, 'EEEE dd MMM');
    const fechaUltimo = Utilities.formatDate(elegido.infoZona.fechaUltimoEvento, CFG.TIMEZONE, 'EEEE dd MMM');
    const numEventos = elegido.infoZona.eventos.length;

    alertaZona = `
*** AGRUPAMIENTO DETECTADO ***
${elegido.tecnico.nombre} ya estará en ${ticket.departamento} (${numEventos} evento${numEventos > 1 ? 's' : ''} programado${numEventos > 1 ? 's' : ''} entre ${fechaInicio} y ${fechaUltimo}).

Considera optimizar la fecha del nuevo ticket para aprovechar el viaje:
  CAMBIAR_FECHA: YYYY-MM-DD

Esto ahorra el viaje de ida desde Guatemala.
*****************************
`;
  }

  const cuerpo = `
Nuevo ticket requiere aprobación.
${alertaZona}
-------------------------------------------------
Ticket:        ${ticket.id}
Prioridad:     ${ticket.prioridad}
Vendedor:      ${ticket.vendedor}
Cliente:       ${ticket.cliente}
Teléfono:      ${ticket.telefonoCliente || '(no especificado)'}
Dirección:     ${ticket.direccion}
Departamento:  ${ticket.departamento}
Equipo:        ${ticket.equipo}
Servicio:      ${ticket.tipoServicio}
Notas:         ${ticket.notas || '(sin notas)'}
-------------------------------------------------
Técnico sugerido:   ${elegido.tecnico.nombre}
Sale desde:         ${elegido.origenUsado || elegido.tecnico.base || 'base'}
Fecha propuesta:    ${elegido.fechaSlot} ${formatearHora_(elegido.horaInicioSlot)}
Hora sugerida ven.: ${ticket.horaSugerida ? formatearHora_(ticket.horaSugerida) : 'no especificada'}
Traslado (ida):     ${formatearHorasMinutos_(elegido.trasladoHoras)}
Distancia:          ${elegido.distanciaKm.toFixed(1)} km${lineaViaticos}
-------------------------------------------------

CÓMO RESPONDER (primera línea del correo):

  APROBADO
  CAMBIAR: <nombre del técnico>
  CAMBIAR_FECHA: YYYY-MM-DD (mantiene técnico, mueve fecha)
  RECHAZADO: <motivo>

Mantén el asunto intacto.
`;

  // 1) Correo a jefatura para aprobar
  GmailApp.sendEmail(CFG.EMAIL_JEFATURA, asunto, cuerpo, {
    name: 'Sistema de Tickets'
  });

  // 2) Correo de confirmación al vendedor (sin pedir aprobación)
  if (ticket.emailVendedor) {
    const asuntoVendedor = `[TICKET CREADO ${ticket.id}] ${ticket.cliente}`;
    const cuerpoVendedor = `
Tu ticket fue creado y está esperando aprobación de jefatura.

-------------------------------------------------
Ticket:           ${ticket.id}
Prioridad:        ${ticket.prioridad}
Cliente:          ${ticket.cliente}
Equipo:           ${ticket.equipo}
Servicio:         ${ticket.tipoServicio}
Departamento:     ${ticket.departamento}
Dirección:        ${ticket.direccion}
-------------------------------------------------
Técnico sugerido: ${elegido.tecnico.nombre}
Fecha propuesta:  ${elegido.fechaSlot} ${formatearHora_(elegido.horaInicioSlot)}${lineaViaticos}
-------------------------------------------------

Estado: BORRADOR (esperando aprobación)

Recibirás otro correo cuando el técnico confirme las horas y el ticket pase a CONFIRMADO.

Si tienes alguna observación o necesitas modificar algo, contacta a jefatura.
`;
    GmailApp.sendEmail(ticket.emailVendedor, asuntoVendedor, cuerpoVendedor, {
      name: 'Sistema de Tickets'
    });
  }
}

function manejarUrgenteConDesplazamiento_(ticket, hoja, fila, candidatos) {
  log_('INFO', `Urgente ${ticket.id} requiere desplazamiento`);

  hoja.getRange(fila, hoja.getLastColumn() - 5, 1, 6).setValues([[
    ticket.id,
    '(pendiente decisión)',
    '', '', '',
    'URGENTE_SIN_SLOT'
  ]]);

  enviarCorreoUrgenteJefatura_(ticket, candidatos);
  actualizarDashboard_();
}

function manejarUrgenteSinOpciones_(ticket, hoja, fila) {
  log_('WARN', `Urgente ${ticket.id} sin opciones`);

  hoja.getRange(fila, hoja.getLastColumn() - 5, 1, 6).setValues([[
    ticket.id,
    '(sin opciones)',
    '', '', '',
    'URGENTE_SIN_OPCIONES'
  ]]);

  const asunto = `[URGENTE SIN SLOT ${ticket.id}] Sin técnicos disponibles`;
  const cuerpo = `
ATENCIÓN: Ticket URGENTE sin opciones de agendado.

-------------------------------------------------
Ticket:        ${ticket.id}
Cliente:       ${ticket.cliente}
Teléfono:      ${ticket.telefonoCliente || '(no especificado)'}
Dirección:     ${ticket.direccion}
Equipo:        ${ticket.equipo}
Servicio:      ${ticket.tipoServicio}
Notas:         ${ticket.notas}
-------------------------------------------------

No hay técnicos con slot libre hoy/mañana, y tampoco hay tickets Media/Baja
desplazables. Acciones posibles:

  RETRASAR
  FORZAR: <nombre del técnico>

Responde con una opción en la primera línea.
`;

  GmailApp.sendEmail(CFG.EMAIL_JEFATURA, asunto, cuerpo, {
    name: 'Sistema de Tickets'
  });
  actualizarDashboard_();
}

function enviarCorreoUrgenteJefatura_(ticket, candidatos) {
  const asunto = `[URGENTE SIN SLOT ${ticket.id}] ${ticket.cliente} - decisión`;

  let opciones = '';
  candidatos.slice(0, 3).forEach((c, idx) => {
    opciones += `\nOPCIÓN ${idx + 1}: ${c.tecnico.nombre} (traslado ${formatearHorasMinutos_(c.trasladoHoras)})\n`;
    opciones += `   Tickets movibles:\n`;
    c.movibles.slice(0, 5).forEach(m => {
      opciones += `     - ${m.ticketId} (${m.prioridad}) el ${Utilities.formatDate(m.fechaActual, CFG.TIMEZONE, 'EEE dd/MM HH:mm')} - ${m.duracionHoras.toFixed(1)}h\n`;
    });
  });

  const cuerpo = `
ATENCIÓN: Ticket URGENTE sin hueco directo.

-------------------------------------------------
Ticket:        ${ticket.id}
Cliente:       ${ticket.cliente}
Teléfono:      ${ticket.telefonoCliente || '(no especificado)'}
Dirección:     ${ticket.direccion}
Equipo:        ${ticket.equipo}
Servicio:      ${ticket.tipoServicio}
Notas:         ${ticket.notas}
-------------------------------------------------

${opciones}

-------------------------------------------------

CÓMO RESPONDER:

  DESPLAZAR: <TicketID>
  RETRASAR
  RECHAZADO: <motivo>

Mantén el asunto intacto.
`;

  GmailApp.sendEmail(CFG.EMAIL_JEFATURA, asunto, cuerpo, {
    name: 'Sistema de Tickets'
  });
}

function notificarErrorJefatura_(ticket, err) {
  // NUNCA notificar errores de "sin capacidad" - estos son ruido operacional
  if (err && err.message && /sin capacidad/i.test(err.message)) {
    log_('INFO', `notificarErrorJefatura_: error "${err.message}" silenciado (operacional)`);
    return;
  }
  // NUNCA notificar errores BUG INTERNO de crearEventoBorrador
  if (err && err.message && /BUG INTERNO/i.test(err.message)) {
    log_('INFO', `notificarErrorJefatura_: error BUG INTERNO silenciado (ya está en log)`);
    return;
  }

  // No enviar correos de error para tickets que ya están cerrados (RESUELTO, RECHAZADO, etc.)
  try {
    if (ticket.id) {
      const ctx = buscarTicketPorId_(ticket.id);
      if (ctx) {
        const estado = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();
        const estadosCerrados = ['RESUELTO', 'RECHAZADO', 'CANCELADO_MANUALMENTE',
                                  'LLAMADA_COMPLETADA', 'REVISION_COMPLETADA',
                                  'LLAMADA_ESCALADA', 'REVISION_ESCALADA',
                                  'LLAMADA_ENVIADA', 'REVISION_ENVIADA', 'PENDIENTE',
                                  'APROBADO_ESPERA_TECNICO', 'CONFIRMADO'];
        if (estadosCerrados.some(s => estado.startsWith(s))) {
          log_('INFO', `notificarErrorJefatura_: ticket ${ticket.id} en estado ${estado}, SE OMITE correo de error`);
          return;
        }
      }
    }
  } catch(_) {}

  enviarCorreoUnico_(
    CFG.EMAIL_JEFATURA,
    `[ERROR ${ticket.id}] No se pudo procesar ticket`,
    `Error: ${err.message}\n\n${err.stack}`,
    ticket.id,
    'error_procesamiento'
  );
}

// =====================================================================
// HELPER PARA PRUEBAS DESDE EL EDITOR
// =====================================================================
/**
 * Simula que la última fila del Sheet acaba de llegar como un Form submit.
 * Útil para reprocesar un ticket que falló o probar cambios sin enviar Form.
 *
 * Correr desde el editor: seleccionar testSimularUltimoTicket → Ejecutar.
 */
function testSimularUltimoTicket() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TICKETS);
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    console.log('No hay tickets en la Sheet para simular.');
    return;
  }

  console.log(`Simulando envío de Form con la fila ${ultimaFila}...`);

  const eventoFalso = {
    range: hoja.getRange(ultimaFila, 1),
    values: hoja.getRange(ultimaFila, 1, 1, hoja.getLastColumn()).getValues()[0]
  };

  onFormSubmit(eventoFalso);
  console.log('Simulación completada. Revisa el log de la Sheet y tu correo.');
}

// =====================================================================
// HELPER: obtener el técnico que recibió la ÚLTIMA llamada/revisión
// (recorre la Sheet de atrás hacia adelante buscando tickets tipo
//  remoto y devuelve el primero que encuentre, excluyendo el actual)
// =====================================================================
function obtenerUltimoTecnicoDeLlamada_(hoja, ticketIdActual) {
  const data = hoja.getDataRange().getValues();
  const headers = data[0].map(h => normalizarHeader_(h));
  const colTipo = headers.indexOf('tipo_de_servicio');
  if (colTipo < 0) return null;

  const base = hoja.getLastColumn() - 6;
  const colTicketId = base;
  const colTecnico = base + 1;

  // Recorrer de la última fila hacia la primera
  for (let i = data.length - 1; i >= 1; i--) {
    const fila = data[i];
    const ticketId = fila[colTicketId];
    const tipo = String(fila[colTipo] || '').toLowerCase();
    const tecnicoNombre = fila[colTecnico];

    // Saltar el ticket actual (el que acabamos de recibir)
    if (ticketId === ticketIdActual) continue;

    // Solo contar tickets tipo remoto (Llamada o Revisión Equipo)
    if (tipo.includes('llamada') || tipo.includes('revisión') || tipo.includes('revision')) {
      if (tecnicoNombre && String(tecnicoNombre).trim() !== '') {
        return String(tecnicoNombre).trim();
      }
    }
  }

  return null;
}
