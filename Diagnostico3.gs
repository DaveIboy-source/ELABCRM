/**
 * DIAGNOSTICO v5: valida que
 *   1. Se quitó el límite de 5h de traslado
 *   2. Se activa modo HOTEL si traslado > 5h
 *   3. Se crean 2 eventos en Calendar en modo hotel
 *   4. Todos los técnicos son candidatos (no solo Moisés)
 *   5. Los viáticos se calculan correctamente
 *
 * Pegar en archivo nuevo "Diagnostico3.gs" y correr
 *   debugValidarV5()   — análisis del último ticket
 *   debugSimular(depto) — simula con cualquier depto
 */

function debugValidarV5() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TICKETS);
  const fila = hoja.getLastRow();

  if (fila < 2) {
    console.log('No hay tickets para analizar');
    return;
  }

  const ticket = leerTicketDeFila_(hoja, fila);
  console.log('============================================');
  console.log('DIAGNOSTICO v5 - Ultimo ticket enviado');
  console.log('============================================');
  console.log(`Ticket:       ${ticket.id}`);
  console.log(`Cliente:      ${ticket.cliente}`);
  console.log(`Direccion:    ${ticket.direccion}`);
  console.log(`Departamento: ${ticket.departamento}`);
  console.log(`Prioridad:    ${ticket.prioridad}`);
  console.log(`Fecha sug.:   ${ticket.fechaSugerida}`);
  console.log(`Hora sug.:    ${ticket.horaSugerida}`);
  console.log(`Umbral hotel: ${CFG.UMBRAL_HOTEL_HORAS}h\n`);

  analizarTicket_(ticket);
}

function debugSimular(departamento, direccionOpcional) {
  console.log('============================================');
  console.log('SIMULACION - ticket ficticio');
  console.log('============================================');

  const ticket = {
    id: 'TCK_SIM',
    cliente: 'Cliente de prueba',
    direccion: direccionOpcional || departamento + ', Guatemala',
    departamento: departamento,
    prioridad: 'Media',
    fechaSugerida: new Date(Date.now() + 86400000),
    horaSugerida: 10,
    notas: 'simulacion',
    vendedor: 'test',
    emailVendedor: '',
    telefonoCliente: '',
    equipo: 'TEST',
    tipoServicio: 'Reparación'
  };

  console.log(`Direccion a geocodificar: ${ticket.direccion}`);
  analizarTicket_(ticket);
}

function analizarTicket_(ticket) {
  // 1. Geocoding
  console.log('--- GEOCODING ---');
  let destino;
  try {
    destino = geocodificar_(ticket.direccion);
    ticket.lat = destino.lat;
    ticket.lng = destino.lng;
    console.log(`  OK: ${destino.formateada}`);
    console.log(`  lat=${destino.lat}, lng=${destino.lng}\n`);
  } catch(e) {
    console.log(`  FALLO: ${e.message}`);
    return;
  }

  // 2. Técnicos
  const tecnicos = cargarTecnicos_();
  console.log(`--- TECNICOS (${tecnicos.length}) ---`);
  tecnicos.forEach(t => {
    console.log(`  ${t.nombre}: base="${t.base}", activo=${t.activo}, email=${t.email ? 'SI' : 'NO'}`);
    console.log(`    departamentos: [${t.departamentos.join(', ') || 'sin filtro'}]`);
  });
  console.log('');

  // 3. Evaluación uno por uno
  console.log('--- EVALUACION ---');
  const resultado = evaluarCandidatos_(ticket, tecnicos);

  console.log(`\nCandidatos directos: ${resultado.candidatosDirectos.length}`);
  resultado.candidatosDirectos.forEach((c, idx) => {
    console.log(`\n  [${idx + 1}] ${c.tecnico.nombre}`);
    console.log(`      Traslado ida: ${c.trasladoHoras.toFixed(2)}h`);
    console.log(`      Distancia:    ${c.distanciaKm.toFixed(1)}km`);
    console.log(`      Modo:         ${c.modoHotel ? 'HOTEL (2 dias)' : 'NORMAL (1 dia)'}`);
    console.log(`      Fecha slot:   ${c.fechaSlot} @ ${c.horaInicioSlot}:00`);
    if (c.modoHotel) {
      console.log(`      Fecha vuelta: ${c.fechaVuelta} @ ${c.horaVuelta}:00`);
    }
    console.log(`      Score:        ${c.score.toFixed(2)}`);
  });

  if (resultado.candidatosConDesplazamiento.length > 0) {
    console.log(`\nCandidatos con desplazamiento: ${resultado.candidatosConDesplazamiento.length}`);
    resultado.candidatosConDesplazamiento.forEach((c, idx) => {
      console.log(`  [${idx + 1}] ${c.tecnico.nombre}: ${c.movibles.length} tickets movibles`);
    });
  }

  // 4. Diagnóstico de modo hotel
  if (resultado.candidatosDirectos.length > 0) {
    const mejor = resultado.candidatosDirectos[0];
    console.log('\n--- VALIDACION ---');
    console.log(`  Traslado del ganador: ${mejor.trasladoHoras.toFixed(2)}h`);
    console.log(`  Umbral hotel: ${CFG.UMBRAL_HOTEL_HORAS}h`);
    console.log(`  Modo esperado: ${mejor.trasladoHoras > CFG.UMBRAL_HOTEL_HORAS ? 'HOTEL' : 'NORMAL'}`);
    console.log(`  Modo detectado: ${mejor.modoHotel ? 'HOTEL' : 'NORMAL'}`);
    const ok = (mejor.trasladoHoras > CFG.UMBRAL_HOTEL_HORAS) === mejor.modoHotel;
    console.log(`  ${ok ? 'CORRECTO' : 'INCORRECTO - revisar logica'}`);

    // 5. Viáticos
    const viaticos = calcularViaticos_(ticket, mejor, 1);
    console.log('\n--- VIATICOS ---');
    if (viaticos.aplica) {
      console.log(`  Desayuno: Q${viaticos.desayuno}`);
      console.log(`  Almuerzo: Q${viaticos.almuerzo}`);
      console.log(`  Cena:     Q${viaticos.cena}`);
      console.log(`  Hotel:    Q${viaticos.hotel}`);
      console.log(`  TOTAL:    Q${viaticos.total}`);
      console.log(`  Resumen:  "${viaticos.resumen}"`);
    } else {
      console.log('  No aplican viáticos (Guatemala)');
    }
  } else {
    console.log('\n!!! SIN CANDIDATOS - ningún técnico disponible !!!');
  }
}
