/**
 * =====================================================================
 * UTILIDADES.gs — v4 — helpers comunes
 * =====================================================================
 */

function cargarTecnicos_() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TECNICOS);
  if (!hoja) throw new Error(`Falta la pestaña "${CFG.TAB_TECNICOS}"`);

  const data = hoja.getDataRange().getValues();
  const headers = data.shift().map(h => normalizarHeader_(h));

  return data.filter(r => r[0]).map(row => {
    const t = {};
    headers.forEach((h, i) => t[h] = row[i]);
    return {
      nombre: t.nombre,
      email: String(t.email || '').toLowerCase().trim(),
      base: t.base || CFG.ORIGEN_DEFAULT,
      departamentos: (t.departamentos || '').split(',').map(s => s.trim()).filter(Boolean),
      activo: String(t.activo).toLowerCase() !== 'false'
    };
  });
}

function geocodificar_(direccion, departamento) {
  const cache = CacheService.getScriptCache();
  // Cache key incluye dirección + departamento para evitar colisiones
  const claveTexto = String(direccion).trim() + '||' + String(departamento || '').trim();
  const clave = 'geo_v2_' + Utilities.base64Encode(claveTexto).slice(0, 200);
  const hit = cache.get(clave);
  if (hit) {
    console.log(`Geocoding cache HIT para "${direccion}"`);
    return JSON.parse(hit);
  }

  // Intento 1: dirección tal cual
  let resultado = intentoGeocoding_(direccion);

  // Intento 2: si falló y nos pasaron departamento, añadirlo
  if (!resultado && departamento) {
    const direccionExtendida = `${direccion}, ${departamento}, Guatemala`;
    console.log(`Geocoding reintento con: "${direccionExtendida}"`);
    resultado = intentoGeocoding_(direccionExtendida);
  }

  // Intento 3: solo con el departamento (último recurso)
  if (!resultado && departamento) {
    const soloDepto = `${departamento}, Guatemala`;
    console.log(`Geocoding último intento solo con departamento: "${soloDepto}"`);
    resultado = intentoGeocoding_(soloDepto);
    if (resultado) {
      log_('WARN', `Dirección "${direccion}" no fue precisa. Usando centro del departamento ${departamento}`);
      resultado.fueFallback = true;  // marcar que vino del fallback
    }
  }

  if (!resultado) {
    throw new Error(
      `No se pudo ubicar la dirección "${direccion}". ` +
      `Pide al vendedor que sea más específico (agregue zona, aldea o referencia).`
    );
  }

  // Guardar en cache: pero si vino del fallback, solo 10 min (no 6 horas)
  // así si hay error temporal de Maps no contamina por mucho tiempo
  const ttl = resultado.fueFallback ? 600 : 21600;
  cache.put(clave, JSON.stringify(resultado), ttl);
  return resultado;
}

/**
 * Limpia todo el caché de geocoding y distance matrix.
 * Usar cuando sospeches que el caché está contaminado con resultados malos.
 */
function limpiarCacheMaps() {
  const cache = CacheService.getScriptCache();

  // Apps Script no permite limpiar todas las claves, pero sí podemos
  // poner TTL=1 segundo en las claves más comunes. Mejor enfoque:
  // borrar el caché completo del script.
  try {
    // No hay forma directa de borrar todas las claves del CacheService.
    // Usamos getAll con prefijos comunes para vaciar lo que sabemos.
    const claves = [];
    // Como no podemos enumerar claves, mejor cambiamos la versión del prefijo
    // Eso invalida automáticamente el caché viejo.
    PropertiesService.getScriptProperties().setProperty(
      'cache_version', String(Date.now())
    );
    console.log('Caché de Maps invalidado. La próxima consulta hará llamadas frescas.');
  } catch(e) {
    console.log('Error: ' + e.message);
  }

  console.log('NOTA: como Apps Script no permite enumerar caché, las entradas viejas');
  console.log('expirarán por sí solas en 6 horas. Si quieres forzarlo, espera ese tiempo');
  console.log('o haz un cambio mínimo en geocodificar_ (ej: cambiar v2 -> v3 en el prefijo).');
}

function intentoGeocoding_(direccion) {
  const url = 'https://maps.googleapis.com/maps/api/geocode/json'
    + '?address=' + encodeURIComponent(direccion)
    + '&region=gt&key=' + getMapsApiKey_();

  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());

  if (data.status !== 'OK' || !data.results.length) {
    console.log(`  Geocoding "${direccion}": ${data.status}`);
    return null;
  }
  const loc = data.results[0].geometry.location;
  return {
    lat: loc.lat,
    lng: loc.lng,
    formateada: data.results[0].formatted_address
  };
}

function distanceMatrix_(origen, destino) {
  const cache = CacheService.getScriptCache();
  const clave = 'dm_' + Utilities.base64Encode(origen + '|' + destino).slice(0, 100);
  const hit = cache.get(clave);
  if (hit) return JSON.parse(hit);

  const url = 'https://maps.googleapis.com/maps/api/distancematrix/json'
    + '?origins=' + encodeURIComponent(origen)
    + '&destinations=' + encodeURIComponent(destino)
    + '&mode=driving&region=gt&departure_time=now'
    + '&key=' + getMapsApiKey_();

  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data = JSON.parse(res.getContentText());

  if (data.status !== 'OK') {
    log_('WARN', `Distance Matrix ${data.status}`);
    return null;
  }
  const el = data.rows[0]?.elements[0];
  if (!el || el.status !== 'OK') return null;

  const out = {
    duracionSegundos: (el.duration_in_traffic || el.duration).value,
    distanciaMetros: el.distance.value
  };
  cache.put(clave, JSON.stringify(out), 3600);
  return out;
}

function obtenerEventosDeTecnico_(emailTecnico, inicio, fin, opciones) {
  opciones = opciones || {};
  const cal = getCalendarCentral_();
  const todos = cal.getEvents(inicio, fin);
  const emailLower = emailTecnico.toLowerCase();

  // Por defecto, ignoramos siempre los BORRADORES porque NO son trabajo confirmado
  // (es decisión correcta: no deben bloquear huecos para nuevos tickets)
  const ignorarBorradores = opciones.ignorarBorradores !== false;

  return todos.filter(ev => {
    const guests = ev.getGuestList();
    const esInvitado = guests.some(g => g.getEmail().toLowerCase() === emailLower);
    if (!esInvitado) return false;

    if (ignorarBorradores) {
      const titulo = ev.getTitle() || '';
      if (/^\[BORRADOR\]/i.test(titulo)) return false;
    }

    return true;
  });
}

// =====================================================================
// Cuenta la carga REAL del técnico para una fecha dada:
//   - Eventos del Calendar (visitas reales en horas)
//   - + 0.5h por cada Llamada/Revisión activa (ENVIADA o VISTA) del mismo día
// Devuelve total en horas.
// =====================================================================
function obtenerCargaRealTecnico_(tecnico, fecha) {
  // 1. Carga de eventos del Calendar
  const inicioDia = new Date(fecha); inicioDia.setHours(CFG.HORA_INICIO, 0, 0, 0);
  const finDia = new Date(fecha); finDia.setHours(CFG.HORA_FIN, 0, 0, 0);

  const eventos = obtenerEventosDeTecnico_(tecnico.email, inicioDia, finDia);
  const horasEventos = eventos.reduce((s, e) =>
    s + (e.getEndTime() - e.getStartTime()) / 3600000, 0);

  // 2. Carga de llamadas/revisiones activas del mismo día
  const horasRemotas = contarLlamadasActivas_(tecnico.nombre, fecha) * 0.5;

  return horasEventos + horasRemotas;
}

// =====================================================================
// Cuenta cuántas Llamadas/Revisiones del técnico están en estado
// ENVIADA o VISTA y fueron creadas el mismo día que la fecha objetivo.
// =====================================================================
function contarLlamadasActivas_(nombreTecnico, fecha) {
  try {
    const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    const hoja = ss.getSheetByName(CFG.TAB_TICKETS);
    const data = hoja.getDataRange().getValues();
    if (data.length < 2) return 0;

    const totalCols = hoja.getLastColumn();
    const colTecnico = totalCols - 5;       // 5ta desde el final (0-indexed)
    const colFechaSlot = totalCols - 3;     // 3ra desde el final
    const colEstado = totalCols - 1;        // última

    const fechaStr = Utilities.formatDate(fecha, CFG.TIMEZONE, 'yyyy-MM-dd');
    const estadosActivos = ['LLAMADA_ENVIADA', 'PENDIENTE',
                            'REVISION_ENVIADA'];

    let contador = 0;
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const tec = fila[colTecnico];
      const estado = String(fila[colEstado] || '').toUpperCase();
      const fechaSlot = fila[colFechaSlot];

      if (tec !== nombreTecnico) continue;
      if (!estadosActivos.includes(estado)) continue;

      // Comparar fecha
      const fechaSlotStr = fechaSlot instanceof Date
        ? Utilities.formatDate(fechaSlot, CFG.TIMEZONE, 'yyyy-MM-dd')
        : String(fechaSlot).substring(0, 10);

      if (fechaSlotStr === fechaStr) contador++;
    }

    return contador;
  } catch (e) {
    log_('WARN', `contarLlamadasActivas_ falló: ${e.message}`);
    return 0;
  }
}

// =====================================================================
// Obtener la ubicación actual del técnico para una fecha dada.
// Lógica:
//   1. Si es el mismo día de trabajo:
//      - Si tiene evento que YA terminó → usa esa ubicación (última parada)
//      - Si tiene evento que aún no empieza → usa esa (estará allá)
//   2. Si es un día futuro:
//      - Si tiene eventos ese día → usa el último
//      - Si no tiene eventos → usa base
//   3. Fallback: base del técnico (o ORIGEN_DEFAULT)
// =====================================================================
function obtenerUbicacionActualTecnico_(tecnico, fechaReferencia) {
  const ahora = new Date();
  const inicioDia = new Date(fechaReferencia);
  inicioDia.setHours(CFG.HORA_INICIO, 0, 0, 0);
  const finDia = new Date(fechaReferencia);
  finDia.setHours(CFG.HORA_FIN, 0, 0, 0);

  const eventos = obtenerEventosDeTecnico_(tecnico.email, inicioDia, finDia)
    .filter(ev => ev.getLocation())
    // Ignorar eventos de REGRESO (el técnico va a estar en su base al final, no en la ubicación del evento)
    .filter(ev => !/^\[?REGRESO|REGRESO/i.test(ev.getTitle() || ''))
    .sort((a, b) => a.getStartTime() - b.getStartTime());

  const baseFallback = tecnico.base || CFG.ORIGEN_DEFAULT;

  if (eventos.length === 0) return baseFallback;

  // Mismo día: buscar el evento más cercano en tiempo
  const esHoy = (fechaReferencia.getFullYear() === ahora.getFullYear()
              && fechaReferencia.getMonth() === ahora.getMonth()
              && fechaReferencia.getDate() === ahora.getDate());

  if (esHoy) {
    // Buscar el último evento que ya terminó antes de ahora
    let ultimoTerminado = null;
    let proximoPorVenir = null;
    for (const ev of eventos) {
      if (ev.getEndTime() <= ahora) {
        ultimoTerminado = ev;
      } else if (!proximoPorVenir) {
        proximoPorVenir = ev;
      }
    }
    if (ultimoTerminado) return ultimoTerminado.getLocation();
    if (proximoPorVenir) return proximoPorVenir.getLocation();
    return baseFallback;
  }

  // Día futuro: la "ubicación al final del día" es la última parada
  const ultimoEvento = eventos[eventos.length - 1];
  return ultimoEvento.getLocation();
}

function buscarSlotConCapacidad_(tecnico, fechaBase, horasNecesarias, ventanaDias, horaSugerida, toleranciaHora, opciones) {
  opciones = opciones || {};
  let cursor = new Date(fechaBase);
  cursor.setHours(0, 0, 0, 0);

  const maxDiasCalendario = ventanaDias * 2;
  let diasHabilesRevisados = 0;

  for (let i = 0; i < maxDiasCalendario && diasHabilesRevisados < ventanaDias; i++) {
    if (!CFG.DIAS_LABORALES.includes(cursor.getDay())) {
      cursor.setDate(cursor.getDate() + 1);
      continue;
    }
    diasHabilesRevisados++;

    const inicioDia = new Date(cursor); inicioDia.setHours(CFG.HORA_INICIO, 0, 0, 0);
    const finDia = new Date(cursor); finDia.setHours(CFG.HORA_FIN, 0, 0, 0);

    const eventos = obtenerEventosDeTecnico_(tecnico.email, inicioDia, finDia, opciones);

    // Sin tope de 9h: solo validamos que haya un hueco que quepa en 8am-5pm
    const horaInicio = calcularMejorHoraLibre_(
      eventos, horasNecesarias, horaSugerida, toleranciaHora);

    if (horaInicio !== null && horaInicio + horasNecesarias <= CFG.HORA_FIN) {
      const distanciaPref = horaSugerida ? Math.abs(horaInicio - horaSugerida) : 0;
      return {
        fecha: new Date(cursor),
        horaInicio,
        distanciaHoraSugerida: distanciaPref
      };
    }

    cursor.setDate(cursor.getDate() + 1);
  }
  return null;
}

function calcularMejorHoraLibre_(eventos, horasRequeridas, horaSugerida, toleranciaHora) {
  const ordenados = eventos
    .map(ev => ({
      ini: ev.getStartTime().getHours() + ev.getStartTime().getMinutes() / 60,
      fin: ev.getEndTime().getHours() + ev.getEndTime().getMinutes() / 60
    }))
    .sort((a, b) => a.ini - b.ini);

  const huecos = [];
  let cursor = CFG.HORA_INICIO;
  for (const ev of ordenados) {
    if (ev.ini - cursor >= horasRequeridas) {
      huecos.push({ ini: cursor, fin: ev.ini });
    }
    cursor = Math.max(cursor, ev.fin);
  }
  if (CFG.HORA_FIN - cursor >= horasRequeridas) {
    huecos.push({ ini: cursor, fin: CFG.HORA_FIN });
  }

  if (huecos.length === 0) return null;
  if (!horaSugerida) return huecos[0].ini;

  let mejor = null;
  let mejorDist = Infinity;

  for (const h of huecos) {
    if (horaSugerida >= h.ini && horaSugerida + horasRequeridas <= h.fin) {
      return horaSugerida;
    }
    let candidatoInicio;
    if (horaSugerida < h.ini) candidatoInicio = h.ini;
    else candidatoInicio = h.fin - horasRequeridas;
    if (candidatoInicio < h.ini) candidatoInicio = h.ini;
    if (candidatoInicio + horasRequeridas > h.fin) continue;

    const dist = Math.abs(candidatoInicio - horaSugerida);
    if (dist <= toleranciaHora && dist < mejorDist) {
      mejorDist = dist;
      mejor = candidatoInicio;
    }
  }

  if (mejor === null) return huecos[0].ini;
  return mejor;
}

function buscarTicketsMovibles_(tecnico, inicio, fin) {
  const eventos = obtenerEventosDeTecnico_(tecnico.email, inicio, fin);
  const movibles = [];

  for (const ev of eventos) {
    const desc = ev.getDescription() || '';
    const matchPrio = desc.match(/Prioridad:\s*(\w+)/i);
    const matchId = desc.match(/Ticket:\s*(TCK\d+)/i);
    if (!matchId) continue;

    const prioridad = matchPrio ? matchPrio[1] : 'Media';
    if (prioridad === 'Urgente') continue;

    const ticketId = matchId[1];
    const ctx = buscarTicketPorId_(ticketId);
    if (!ctx) continue;
    const estado = String(ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue() || '').toUpperCase();
    if (estado === 'CONFIRMADO') continue;

    movibles.push({
      evento: ev,
      ticketId,
      prioridad,
      duracionHoras: (ev.getEndTime() - ev.getStartTime()) / 3600000,
      fechaActual: ev.getStartTime()
    });
  }

  return movibles;
}

/**
 * Detecta si el técnico estará físicamente en cierto departamento alrededor
 * de la fecha objetivo. Útil para sugerir agrupar visitas y evitar viajes
 * largos innecesarios.
 *
 * Retorna:
 *   null si NO está en zona
 *   { fechaInicio, fechaUltimoEvento, eventos[] } si SÍ está
 *
 * Reglas:
 *   - Solo cuenta si el evento existente está en el mismo departamento
 *   - El técnico se considera "en zona" desde el día de inicio del evento
 *     hasta máximo 5 días laborales después
 *   - Solo detecta eventos a partir de la fecha objetivo (no antes)
 */
function tecnicoEstaraEnZona_(tecnico, fecha, departamento) {
  // Buscar eventos en ventana de -5 a +5 días desde la fecha objetivo
  const inicio = new Date(fecha);
  inicio.setDate(inicio.getDate() - 5);
  inicio.setHours(0, 0, 0, 0);

  const fin = new Date(fecha);
  fin.setDate(fin.getDate() + 5);
  fin.setHours(23, 59, 59, 999);

  const eventos = obtenerEventosDeTecnico_(tecnico.email, inicio, fin);
  const deptoNorm = normalizar_(departamento);

  // Filtrar eventos que sean en el mismo departamento, ignorando los REGRESO
  const eventosEnZona = eventos.filter(ev => {
    // Ignorar eventos de REGRESO (no son visitas reales en zona, es solo el viaje de vuelta)
    if (/^\[?REGRESO|REGRESO/i.test(ev.getTitle() || '')) return false;

    const desc = ev.getDescription() || '';
    const matchDepto = desc.match(/Departamento:\s*(.+)/i);
    return matchDepto && normalizar_(matchDepto[1].trim()) === deptoNorm;
  });

  if (eventosEnZona.length === 0) return null;

  // Encontrar el evento que cubre la fecha objetivo o el inmediatamente anterior
  const fechaObjetivo = new Date(fecha);
  fechaObjetivo.setHours(12, 0, 0, 0);

  // Ordenar por fecha
  eventosEnZona.sort((a, b) => a.getStartTime() - b.getStartTime());

  // Buscar el evento "ancla" (el más cercano a la fecha objetivo, hacia atrás o el mismo día)
  let anchorEvent = null;
  for (const ev of eventosEnZona) {
    const evFecha = ev.getStartTime();
    const diffDias = Math.floor((fechaObjetivo - evFecha) / (24 * 3600 * 1000));

    // El evento debe estar entre 0 y 5 días antes de la fecha objetivo
    // (no contamos antes porque dijiste que no debe ganar tickets de días anteriores)
    if (diffDias >= 0 && diffDias <= 5) {
      // Verificar que no haya feriado/sábado/domingo en medio
      if (esRangoLaboralValido_(evFecha, fechaObjetivo)) {
        anchorEvent = ev;
      }
    }
  }

  if (!anchorEvent) return null;

  return {
    fechaInicio: anchorEvent.getStartTime(),
    fechaUltimoEvento: eventosEnZona[eventosEnZona.length - 1].getStartTime(),
    eventos: eventosEnZona,
    departamento: departamento
  };
}

/**
 * Wrapper de compatibilidad: devuelve true si tecnicoEstaraEnZona_ encuentra match.
 */
function hayAgrupamientoPosible_(tecnico, fecha, departamento) {
  return tecnicoEstaraEnZona_(tecnico, fecha, departamento) !== null;
}

/**
 * Verifica que el rango entre 2 fechas no atraviese fin de semana
 * (sábado o domingo). Útil para considerar que el técnico se queda
 * en la zona durante días laborales seguidos.
 */
function esRangoLaboralValido_(desde, hasta) {
  const inicio = new Date(desde);
  inicio.setHours(0, 0, 0, 0);
  const fin = new Date(hasta);
  fin.setHours(0, 0, 0, 0);

  const diasDiff = Math.floor((fin - inicio) / (24 * 3600 * 1000));
  if (diasDiff > 5) return false;  // máximo 5 días seguidos
  if (diasDiff <= 1) return true;  // 0 o 1 día siempre válido

  // Verificar que no haya fin de semana en medio
  const cursor = new Date(inicio);
  while (cursor <= fin) {
    const dow = cursor.getDay();
    if (dow === 0 || dow === 6) return false;
    cursor.setDate(cursor.getDate() + 1);
  }
  return true;
}



function reoptimizarRutaDelDia_(tecnico, fecha) {
  const inicioDia = new Date(fecha); inicioDia.setHours(CFG.HORA_INICIO, 0, 0, 0);
  const finDia = new Date(fecha); finDia.setHours(CFG.HORA_FIN, 0, 0, 0);

  const eventos = obtenerEventosDeTecnico_(tecnico.email, inicioDia, finDia)
    .filter(ev => ev.getLocation() && !/LOCK/i.test(ev.getDescription() || ''));

  if (eventos.length < 2) return;

  const puntos = eventos.map(ev => ({
    evento: ev,
    direccion: ev.getLocation(),
    duracion: (ev.getEndTime() - ev.getStartTime()) / 3600000
  }));

  const orden = [];
  let actual = tecnico.base || CFG.ORIGEN_DEFAULT;
  const pendientes = puntos.slice();

  while (pendientes.length) {
    let mejor = -1, mejorTiempo = Infinity;
    for (let i = 0; i < pendientes.length; i++) {
      const dm = distanceMatrix_(actual, pendientes[i].direccion);
      if (dm && dm.duracionSegundos < mejorTiempo) {
        mejorTiempo = dm.duracionSegundos;
        mejor = i;
      }
    }
    if (mejor < 0) break;
    const p = pendientes.splice(mejor, 1)[0];
    p.trasladoHoras = mejorTiempo / 3600;
    orden.push(p);
    actual = p.direccion;
  }

  let cursor = CFG.HORA_INICIO;
  for (const p of orden) {
    cursor += p.trasladoHoras;
    const ini = new Date(fecha);
    ini.setHours(Math.floor(cursor), Math.round((cursor % 1) * 60), 0, 0);
    const fin = new Date(ini.getTime() + p.duracion * 3600000);
    if (fin.getHours() + fin.getMinutes() / 60 > CFG.HORA_FIN) {
      log_('WARN', `Ruta excede jornada: ${tecnico.nombre}`);
      break;
    }
    p.evento.setTime(ini, fin);
    cursor += p.duracion;
  }
}

function normalizar_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function formatearHora_(horaDecimal) {
  const h = Math.floor(horaDecimal);
  const m = Math.round((horaDecimal % 1) * 60);
  const ampm = h >= 12 ? 'pm' : 'am';
  const h12 = ((h + 11) % 12) + 1;
  return `${h12}:${String(m).padStart(2, '0')} ${ampm}`;
}

/**
 * Convierte horas decimales (ej: 1.2) en formato legible "1 h 12 min".
 * Ejemplos:
 *   0.25 -> "15 min"
 *   1.0  -> "1 h"
 *   1.2  -> "1 h 12 min"
 *   3.5  -> "3 h 30 min"
 */
function formatearHorasMinutos_(horasDecimal) {
  if (!horasDecimal || horasDecimal <= 0) return '0 min';

  const totalMinutos = Math.round(horasDecimal * 60);
  const h = Math.floor(totalMinutos / 60);
  const m = totalMinutos % 60;

  if (h === 0) return `${m} min`;
  if (m === 0) return `${h} h`;
  return `${h} h ${m} min`;
}

function log_(nivel, msg) {
  try {
    const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
    let hoja = ss.getSheetByName(CFG.TAB_LOG);
    if (!hoja) {
      hoja = ss.insertSheet(CFG.TAB_LOG);
      hoja.appendRow(['Timestamp', 'Nivel', 'Mensaje']);
    }
    hoja.appendRow([new Date(), nivel, msg]);
    console.log(`[${nivel}] ${msg}`);
  } catch(e) { console.log(`[${nivel}] ${msg}`); }
}
