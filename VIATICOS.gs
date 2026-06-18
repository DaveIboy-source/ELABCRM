/**
 * =====================================================================
 * VIATICOS.gs — v2 — soporte modo hotel (2 días)
 * =====================================================================
 *
 * MODO NORMAL (1 día):
 *   - Desayuno Q50 si sale antes de las 8am
 *   - Almuerzo Q50 si jornada cubre 12-14h
 *   - Cena Q50 si tiene evento al día siguiente mismo depto
 *   - Hotel Q200 si traslado > UMBRAL (5h) — pero este caso ahora es HOTEL
 *
 * MODO HOTEL (2 días):
 *   Día 1 (ida + trabajo):
 *     - Desayuno Q50 (siempre, sale temprano)
 *     - Almuerzo Q50 (si jornada pasa por 12-14)
 *     - Hotel Q200
 *     - Cena Q50
 *   Día 2 (regreso):
 *     - Desayuno Q50
 *     - Almuerzo Q50 (si el regreso pasa por 12-14)
 */

function calcularViaticos_(ticket, elegido, horasTrabajo) {
  const montos = CFG.VIATICOS || {
    desayuno: 50, almuerzo: 50, cena: 50, hotel: 200
  };

  // Si es en Guatemala capital, no aplica
  if (normalizar_(ticket.departamento) === 'guatemala') {
    return { aplica: false, total: 0, resumen: '' };
  }

  const trasladoHoras = elegido.trasladoHoras || 0;
  const horaInicio = elegido.horaInicioSlot;
  const trabajoHoras = horasTrabajo || 1;

  // ========== MODO HOTEL ==========
  if (elegido.modoHotel) {
    const desayuno1 = montos.desayuno;
    const almuerzo1 = montos.almuerzo;
    const cena1 = montos.cena;
    const hotel = montos.hotel;
    const desayuno2 = montos.desayuno;

    // Día 2: almuerzo solo si el regreso cruza el mediodía
    const horaVuelta = elegido.horaVuelta || CFG.HORA_INICIO;
    const horaRegresoFin = horaVuelta + trasladoHoras;
    const almuerzo2 = (horaVuelta <= 14 && horaRegresoFin >= 12) ? montos.almuerzo : 0;

    const total = desayuno1 + almuerzo1 + cena1 + hotel + desayuno2 + almuerzo2;

    return {
      aplica: true,
      total,
      desayuno: desayuno1 + desayuno2,
      almuerzo: almuerzo1 + almuerzo2,
      cena: cena1,
      hotel,
      resumen: `Viáticos estimados: Q${total.toFixed(2)} (modo hotel, 2 días)`
    };
  }

  // ========== MODO NORMAL ==========
  const horaSalida = horaInicio - trasladoHoras;
  const horaRegreso = horaInicio + trabajoHoras + trasladoHoras;

  let desayuno = 0, almuerzo = 0, cena = 0;

  if (horaSalida < 8) desayuno = montos.desayuno;
  if (horaSalida <= 14 && horaRegreso >= 12) almuerzo = montos.almuerzo;

  // Cena si hay otro evento día siguiente mismo depto
  if (elegido.fechaSlot && elegido.tecnico) {
    try {
      const fechaSlot = new Date(elegido.fechaSlot + 'T00:00:00');
      const diaSiguiente = new Date(fechaSlot);
      diaSiguiente.setDate(diaSiguiente.getDate() + 1);
      if (hayAgrupamientoPosible_(elegido.tecnico, diaSiguiente, ticket.departamento)) {
        cena = montos.cena;
      }
    } catch(_) {}
  }

  const total = desayuno + almuerzo + cena;

  return {
    aplica: total > 0,
    total,
    desayuno,
    almuerzo,
    cena,
    hotel: 0,
    resumen: total > 0 ? `Viáticos estimados: Q${total.toFixed(2)}` : ''
  };
}
