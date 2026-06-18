/**
 * Dashboard.gs - Dashboard Ejecutivo del sistema de tickets EquiposLab
 *
 * Genera automáticamente en la pestaña "Dashboard":
 *   - KPIs principales (totales, abiertos, en proceso, resueltos, % resolución, tiempo prom.)
 *   - Gráficos: estados, técnicos, departamentos, equipos, servicios, prioridad, tickets/día
 *
 * Se ejecuta desde el menú: Tickets > Actualizar Dashboard
 * También se llama automáticamente desde otros módulos cuando hay cambios.
 */

// ===================================================================
// Configuración del Dashboard
// ===================================================================

const DASHBOARD_TAB = 'Dashboard';

// Mapeo de estados reales → categorías agrupadas
const ESTADOS_ABIERTO = [
  'BORRADOR', 'BORRADOR_FORZADO',
  'LLAMADA_PENDIENTE_APROBACION', 'REVISION_PENDIENTE_APROBACION',
  'ESCALAMIENTO_PENDIENTE_APROBACION',
  'URGENTE_SIN_SLOT', 'URGENTE_SIN_OPCIONES'
];

const ESTADOS_EN_PROCESO = [
  'APROBADO_ESPERA_TECNICO', 'CONFIRMADO',
  'LLAMADA_ENVIADA', 'REVISION_ENVIADA',
  'PENDIENTE', 'ESCALADO'
];

const ESTADOS_RESUELTO = ['RESUELTO'];

const ESTADOS_EXCLUIDOS = ['RECHAZADO', 'CANCELADO_MANUALMENTE'];

// Colores principales
const COLOR_AZUL = '#1976d2';
const COLOR_VERDE = '#2e7d32';
const COLOR_ROJO = '#c62828';
const COLOR_AMARILLO = '#f9a825';
const COLOR_MORADO = '#7b1fa2';
const COLOR_CYAN = '#00838f';
const COLOR_HEADER = '#263238';

// ===================================================================
// FUNCIÓN PRINCIPAL
// ===================================================================

/**
 * Reconstruye completamente el dashboard.
 * Borra todo el contenido anterior y lo vuelve a generar con datos frescos.
 */
function actualizarDashboard_() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  let hoja = ss.getSheetByName(DASHBOARD_TAB);

  if (!hoja) {
    hoja = ss.insertSheet(DASHBOARD_TAB);
  } else {
    // Limpiar gráficos previos
    const graficos = hoja.getCharts();
    for (const g of graficos) hoja.removeChart(g);
    hoja.clear();
  }

  // Mostrar todas las columnas escondidas previas
  try {
    hoja.showColumns(1, hoja.getMaxColumns());
  } catch(_) {}

  // Leer datos reales
  const datos = leerDatosTickets_();
  if (datos.length === 0) {
    hoja.getRange('A1').setValue('No hay tickets aún.');
    return;
  }

  // Calcular métricas globales
  const metricas = calcularMetricas_(datos);

  // 1. Encabezado principal
  construirEncabezado_(hoja);

  // 2. Tarjetas KPI
  construirKPIs_(hoja, metricas);

  // 3. Gráficos
  construirGraficoEstados_(hoja, metricas);
  construirGraficoTecnicos_(hoja, datos);
  construirGraficoDepartamentos_(hoja, datos);
  construirGraficoEquipos_(hoja, datos);
  construirGraficoServicios_(hoja, datos);
  construirGraficoPrioridad_(hoja, datos);
  construirGraficoTicketsPorDia_(hoja, datos);

  // 4. Estilo general
  aplicarEstiloGeneral_(hoja);

  console.log('Dashboard actualizado: ' + datos.length + ' tickets analizados');
}

// ===================================================================
// LECTURA DE DATOS
// ===================================================================

/**
 * Lee la pestaña Tickets y devuelve array de objetos con las
 * columnas relevantes para el dashboard.
 */
function leerDatosTickets_() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TICKETS);
  if (!hoja) return [];

  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = hoja.getLastColumn();
  const headers = hoja.getRange(1, 1, 1, lastCol).getValues()[0];
  const todasFilas = hoja.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Mapear columnas por header
  const idxTimestamp = encontrarColumna_(headers, ['marca temporal', 'timestamp']);
  const idxCliente = encontrarColumna_(headers, ['cliente']);
  const idxDepartamento = encontrarColumna_(headers, ['departamento']);
  const idxEquipo = encontrarColumna_(headers, ['equipo']);
  const idxServicio = encontrarColumna_(headers, ['tipo de servicio', 'tipo_de_servicio', 'servicio']);
  const idxPrioridad = encontrarColumna_(headers, ['prioridad']);
  const idxVendedor = encontrarColumna_(headers, ['vendedor']);

  // Las últimas 6 columnas son sistema: TicketID, Tecnico, Traslado, FechaSlot, EventID, Estado
  const idxTicketId = lastCol - 6;
  const idxTecnico = lastCol - 5;
  const idxFechaSlot = lastCol - 3;
  const idxEstado = lastCol - 1;

  const datos = [];
  for (const fila of todasFilas) {
    const ticketId = fila[idxTicketId];
    if (!ticketId) continue;  // Saltar filas sin TCK

    datos.push({
      ticketId: ticketId,
      fechaCreacion: fila[idxTimestamp],
      fechaResolucion: fila[idxFechaSlot],
      cliente: fila[idxCliente] || '',
      departamento: fila[idxDepartamento] || 'Sin departamento',
      equipo: fila[idxEquipo] || 'Sin equipo',
      servicio: fila[idxServicio] || 'Sin servicio',
      prioridad: fila[idxPrioridad] || 'Media',
      vendedor: fila[idxVendedor] || '',
      tecnico: fila[idxTecnico] || 'Sin asignar',
      estado: String(fila[idxEstado] || '').toUpperCase().split(':')[0].trim()
    });
  }

  return datos;
}

/**
 * Busca índice de columna por nombre de encabezado (lista de posibles nombres)
 */
function encontrarColumna_(headers, posibles) {
  const normalizados = headers.map(function(h) {
    return String(h || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  });
  for (var p of posibles) {
    var pNorm = String(p).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    var idx = normalizados.findIndex(function(h) { return h.includes(pNorm); });
    if (idx >= 0) return idx;
  }
  return -1;
}

// ===================================================================
// CÁLCULO DE MÉTRICAS
// ===================================================================

function calcularMetricas_(datos) {
  let abiertos = 0, enProceso = 0, resueltos = 0, excluidos = 0;
  let tiempoTotalMs = 0;
  let ticketsResueltosConTiempo = 0;

  for (const d of datos) {
    const estado = d.estado;

    if (ESTADOS_EXCLUIDOS.includes(estado)) {
      excluidos++;
    } else if (ESTADOS_RESUELTO.includes(estado)) {
      resueltos++;
      if (d.fechaCreacion instanceof Date && d.fechaResolucion) {
        const fechaRes = d.fechaResolucion instanceof Date
          ? d.fechaResolucion
          : new Date(d.fechaResolucion);
        if (!isNaN(fechaRes.getTime())) {
          tiempoTotalMs += (fechaRes.getTime() - d.fechaCreacion.getTime());
          ticketsResueltosConTiempo++;
        }
      }
    } else if (ESTADOS_EN_PROCESO.includes(estado)) {
      enProceso++;
    } else if (ESTADOS_ABIERTO.includes(estado)) {
      abiertos++;
    } else {
      // Estado desconocido, asumir abierto
      abiertos++;
    }
  }

  const totalActivos = abiertos + enProceso + resueltos;
  const porcentajeResolucion = totalActivos > 0 ? (resueltos / totalActivos) : 0;
  const tiempoPromedioHoras = ticketsResueltosConTiempo > 0
    ? (tiempoTotalMs / ticketsResueltosConTiempo) / (1000 * 3600)
    : 0;

  return {
    total: totalActivos,
    abiertos: abiertos,
    enProceso: enProceso,
    resueltos: resueltos,
    excluidos: excluidos,
    porcentajeResolucion: porcentajeResolucion,
    tiempoPromedioHoras: tiempoPromedioHoras
  };
}

// ===================================================================
// CONSTRUCCIÓN DEL DASHBOARD
// ===================================================================

function construirEncabezado_(hoja) {
  const ahora = Utilities.formatDate(new Date(), CFG.TIMEZONE, 'dd/MM/yyyy HH:mm');

  hoja.getRange('A1:H1').merge()
    .setValue('DASHBOARD EJECUTIVO - SISTEMA DE TICKETS')
    .setBackground(COLOR_HEADER)
    .setFontColor('#ffffff')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  hoja.setRowHeight(1, 40);

  hoja.getRange('A2:H2').merge()
    .setValue('Actualizado: ' + ahora)
    .setBackground('#37474f')
    .setFontColor('#ffffff')
    .setFontSize(10)
    .setHorizontalAlignment('center');

  hoja.setRowHeight(2, 22);
  hoja.setRowHeight(3, 15);
}

function construirKPIs_(hoja, m) {
  // Primera fila de KPIs (fila 4-5): Total, Abiertos, En proceso, Resueltos
  const kpis = [
    { col: 1, label: 'TOTAL TICKETS', value: m.total, color: COLOR_AZUL, format: '#,##0' },
    { col: 3, label: 'ABIERTOS', value: m.abiertos, color: COLOR_ROJO, format: '#,##0' },
    { col: 5, label: 'EN PROCESO', value: m.enProceso, color: COLOR_AMARILLO, format: '#,##0' },
    { col: 7, label: 'RESUELTOS', value: m.resueltos, color: COLOR_VERDE, format: '#,##0' }
  ];

  for (const k of kpis) {
    hoja.getRange(4, k.col, 1, 2).merge()
      .setValue(k.label)
      .setBackground(k.color)
      .setFontColor('#ffffff')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    hoja.getRange(5, k.col, 1, 2).merge()
      .setValue(k.value)
      .setBackground('#ffffff')
      .setFontColor(k.color)
      .setFontSize(32)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat(k.format)
      .setBorder(true, true, true, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  }

  hoja.setRowHeight(4, 30);
  hoja.setRowHeight(5, 65);
  hoja.setRowHeight(6, 10);

  // Segunda fila de KPIs (fila 7-8): % Resolución, Tiempo Promedio
  const kpis2 = [
    { col: 1, label: '% RESOLUCIÓN', value: m.porcentajeResolucion, color: COLOR_MORADO, format: '0.0%', span: 4 },
    { col: 5, label: 'TIEMPO PROM. RESOLUCIÓN (horas)', value: m.tiempoPromedioHoras, color: COLOR_CYAN, format: '0.0', span: 4 }
  ];

  for (const k of kpis2) {
    hoja.getRange(7, k.col, 1, k.span).merge()
      .setValue(k.label)
      .setBackground(k.color)
      .setFontColor('#ffffff')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    hoja.getRange(8, k.col, 1, k.span).merge()
      .setValue(k.value)
      .setBackground('#ffffff')
      .setFontColor(k.color)
      .setFontSize(24)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setNumberFormat(k.format)
      .setBorder(true, true, true, true, false, false, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  }

  hoja.setRowHeight(7, 25);
  hoja.setRowHeight(8, 50);
  hoja.setRowHeight(9, 15);
}

function construirGraficoEstados_(hoja, m) {
  const startCol = 11;
  hoja.getRange(4, startCol, 1, 2).setValues([['Estado', 'Cantidad']]);
  hoja.getRange(5, startCol, 3, 2).setValues([
    ['Abiertos', m.abiertos],
    ['En proceso', m.enProceso],
    ['Resueltos', m.resueltos]
  ]);

  const chart = hoja.newChart()
    .asPieChart()
    .addRange(hoja.getRange(4, startCol, 4, 2))
    .setPosition(11, 1, 0, 0)
    .setOption('title', 'Estados de Tickets')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('pieHole', 0.4)
    .setOption('colors', [COLOR_ROJO, COLOR_AMARILLO, COLOR_VERDE])
    .setOption('width', 480)
    .setOption('height', 320)
    .setOption('legend', { position: 'right' })
    .build();
  hoja.insertChart(chart);
}

function construirGraficoTecnicos_(hoja, datos) {
  const conteo = {};
  for (const d of datos) {
    if (ESTADOS_EXCLUIDOS.includes(d.estado)) continue;
    const tec = d.tecnico || 'Sin asignar';
    conteo[tec] = (conteo[tec] || 0) + 1;
  }

  const ordenado = Object.entries(conteo).sort(function(a, b) { return b[1] - a[1]; });
  if (ordenado.length === 0) return;

  const startCol = 14;
  hoja.getRange(4, startCol, 1, 2).setValues([['Técnico', 'Tickets']]);
  for (let i = 0; i < ordenado.length; i++) {
    hoja.getRange(5 + i, startCol, 1, 2).setValues([ordenado[i]]);
  }

  const chart = hoja.newChart()
    .asBarChart()
    .addRange(hoja.getRange(4, startCol, ordenado.length + 1, 2))
    .setPosition(11, 5, 0, 0)
    .setOption('title', 'Tickets por Técnico')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('colors', [COLOR_AZUL])
    .setOption('width', 480)
    .setOption('height', 320)
    .setOption('legend', { position: 'none' })
    .build();
  hoja.insertChart(chart);
}

function construirGraficoDepartamentos_(hoja, datos) {
  const conteo = {};
  for (const d of datos) {
    if (ESTADOS_EXCLUIDOS.includes(d.estado)) continue;
    const dep = d.departamento || 'Sin departamento';
    conteo[dep] = (conteo[dep] || 0) + 1;
  }

  const ordenado = Object.entries(conteo).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 15);
  if (ordenado.length === 0) return;

  const startCol = 17;
  hoja.getRange(4, startCol, 1, 2).setValues([['Departamento', 'Tickets']]);
  for (let i = 0; i < ordenado.length; i++) {
    hoja.getRange(5 + i, startCol, 1, 2).setValues([ordenado[i]]);
  }

  const chart = hoja.newChart()
    .asBarChart()
    .addRange(hoja.getRange(4, startCol, ordenado.length + 1, 2))
    .setPosition(28, 1, 0, 0)
    .setOption('title', 'Departamentos Más Visitados')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('colors', [COLOR_VERDE])
    .setOption('width', 480)
    .setOption('height', 360)
    .setOption('legend', { position: 'none' })
    .build();
  hoja.insertChart(chart);
}

function construirGraficoEquipos_(hoja, datos) {
  const conteo = {};
  for (const d of datos) {
    if (ESTADOS_EXCLUIDOS.includes(d.estado)) continue;
    const eq = d.equipo || 'Sin equipo';
    conteo[eq] = (conteo[eq] || 0) + 1;
  }

  const ordenado = Object.entries(conteo).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 10);
  if (ordenado.length === 0) return;

  const startCol = 20;
  hoja.getRange(4, startCol, 1, 2).setValues([['Equipo', 'Tickets']]);
  for (let i = 0; i < ordenado.length; i++) {
    hoja.getRange(5 + i, startCol, 1, 2).setValues([ordenado[i]]);
  }

  const chart = hoja.newChart()
    .asColumnChart()
    .addRange(hoja.getRange(4, startCol, ordenado.length + 1, 2))
    .setPosition(28, 5, 0, 0)
    .setOption('title', 'Equipos con Más Tickets (Top 10)')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('colors', ['#f57c00'])
    .setOption('width', 480)
    .setOption('height', 360)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 30 })
    .build();
  hoja.insertChart(chart);
}

function construirGraficoServicios_(hoja, datos) {
  const conteo = {};
  for (const d of datos) {
    if (ESTADOS_EXCLUIDOS.includes(d.estado)) continue;
    const sv = d.servicio || 'Sin servicio';
    conteo[sv] = (conteo[sv] || 0) + 1;
  }

  const ordenado = Object.entries(conteo).sort(function(a, b) { return b[1] - a[1]; });
  if (ordenado.length === 0) return;

  const startCol = 23;
  hoja.getRange(4, startCol, 1, 2).setValues([['Servicio', 'Tickets']]);
  for (let i = 0; i < ordenado.length; i++) {
    hoja.getRange(5 + i, startCol, 1, 2).setValues([ordenado[i]]);
  }

  const chart = hoja.newChart()
    .asPieChart()
    .addRange(hoja.getRange(4, startCol, ordenado.length + 1, 2))
    .setPosition(46, 1, 0, 0)
    .setOption('title', 'Tipos de Servicio')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('pieHole', 0.4)
    .setOption('width', 480)
    .setOption('height', 360)
    .setOption('legend', { position: 'right' })
    .build();
  hoja.insertChart(chart);
}

function construirGraficoPrioridad_(hoja, datos) {
  const conteo = {};
  for (const d of datos) {
    if (ESTADOS_EXCLUIDOS.includes(d.estado)) continue;
    const prio = d.prioridad || 'Media';
    if (!conteo[prio]) conteo[prio] = { abierto: 0, proceso: 0, resuelto: 0 };

    if (ESTADOS_RESUELTO.includes(d.estado)) conteo[prio].resuelto++;
    else if (ESTADOS_EN_PROCESO.includes(d.estado)) conteo[prio].proceso++;
    else conteo[prio].abierto++;
  }

  const startCol = 26;
  hoja.getRange(4, startCol, 1, 4).setValues([['Prioridad', 'Abiertos', 'En proceso', 'Resueltos']]);

  const orden = ['Urgente', 'Alta', 'Media', 'Baja'];
  const filas = [];
  for (const prio of orden) {
    if (!conteo[prio]) continue;
    filas.push([prio, conteo[prio].abierto, conteo[prio].proceso, conteo[prio].resuelto]);
  }
  // Agregar prioridades no estándar al final
  for (const prio of Object.keys(conteo)) {
    if (orden.includes(prio)) continue;
    filas.push([prio, conteo[prio].abierto, conteo[prio].proceso, conteo[prio].resuelto]);
  }
  if (filas.length === 0) return;

  hoja.getRange(5, startCol, filas.length, 4).setValues(filas);

  const chart = hoja.newChart()
    .asColumnChart()
    .addRange(hoja.getRange(4, startCol, filas.length + 1, 4))
    .setPosition(46, 5, 0, 0)
    .setOption('title', 'Tickets por Prioridad')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('isStacked', true)
    .setOption('colors', [COLOR_ROJO, COLOR_AMARILLO, COLOR_VERDE])
    .setOption('width', 480)
    .setOption('height', 360)
    .setOption('legend', { position: 'top' })
    .build();
  hoja.insertChart(chart);
}

function construirGraficoTicketsPorDia_(hoja, datos) {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const hace30 = new Date(hoy);
  hace30.setDate(hace30.getDate() - 30);

  const conteoPorDia = {};
  for (let i = 0; i <= 30; i++) {
    const d = new Date(hace30);
    d.setDate(d.getDate() + i);
    const key = Utilities.formatDate(d, CFG.TIMEZONE, 'yyyy-MM-dd');
    conteoPorDia[key] = 0;
  }

  for (const d of datos) {
    if (!(d.fechaCreacion instanceof Date)) continue;
    if (d.fechaCreacion < hace30) continue;
    const key = Utilities.formatDate(d.fechaCreacion, CFG.TIMEZONE, 'yyyy-MM-dd');
    if (conteoPorDia[key] !== undefined) conteoPorDia[key]++;
  }

  const dias = Object.keys(conteoPorDia).sort();

  const startCol = 31;
  hoja.getRange(4, startCol, 1, 2).setValues([['Fecha', 'Tickets']]);
  const filas = dias.map(function(d) { return [d, conteoPorDia[d]]; });
  hoja.getRange(5, startCol, filas.length, 2).setValues(filas);

  const chart = hoja.newChart()
    .asLineChart()
    .addRange(hoja.getRange(4, startCol, filas.length + 1, 2))
    .setPosition(63, 1, 0, 0)
    .setOption('title', 'Tickets Creados por Día (últimos 30 días)')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('colors', [COLOR_AZUL])
    .setOption('width', 980)
    .setOption('height', 320)
    .setOption('legend', { position: 'none' })
    .setOption('pointSize', 5)
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .build();
  hoja.insertChart(chart);
}

// ===================================================================
// ESTILO GENERAL
// ===================================================================

function aplicarEstiloGeneral_(hoja) {
  // Ancho de columnas principales (A-H)
  for (let i = 1; i <= 8; i++) {
    hoja.setColumnWidth(i, 130);
  }

  // Ocultar zona de datos auxiliares (columna K en adelante hasta J de cada tabla)
  const lastCol = hoja.getMaxColumns();
  if (lastCol >= 11) {
    try {
      hoja.hideColumns(11, lastCol - 10);
    } catch(_) {}
  }

  // Quitar gridlines
  hoja.setHiddenGridlines(true);
}

// ===================================================================
// API PÚBLICA
// ===================================================================

/**
 * Refresca el dashboard manualmente desde el menú.
 * Llamado por: Tickets > Actualizar Dashboard
 */
function refrescarDashboardManual() {
  actualizarDashboard_();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Dashboard actualizado correctamente',
    'Sistema de Tickets',
    5
  );
}
