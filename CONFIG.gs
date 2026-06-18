/**
 * =====================================================================
 * CONFIG.gs — v4 — con prioridad y viáticos
 * =====================================================================
 */

const CFG = {
  // ---------- Sheet ----------
  SHEET_ID: '1eRgoQLyg8cAUVcp9ZEQa7O44mDeBOrAUdYgBmLhgdqM',
  TAB_TICKETS: 'Solicitudes de Servicio Tecnico',
  TAB_TECNICOS: 'Tecnicos',
  TAB_DASHBOARD: 'Dashboard',
  TAB_LOG: 'Log',

  // ---------- Calendario CENTRAL ----------
  CALENDAR_CENTRAL_ID: 'c_59704623b8ed4d6615555f2a9769269eb0ece41a62bdd45850fbed3db1bad927@group.calendar.google.com',

  // ---------- Correos ----------
  EMAIL_JEFATURA: 'jefaturadt@equiposlab.com',
  EMAIL_REMITENTE_ALIAS: '',
  ETIQUETA_GMAIL_APROBACIONES: 'tickets/aprobaciones',
  ETIQUETA_GMAIL_TECNICOS: 'tickets/tecnicos',

  // ---------- Horario ----------
  HORA_INICIO: 8,
  HORA_FIN: 17,
  HORA_MINIMA_SALIDA: 5,        // los técnicos pueden salir desde las 5am para viajes largos
  DIAS_LABORALES: [1, 2, 3, 4, 5],
  CAPACIDAD_DIARIA_HORAS: 9,    // se aplica solo a horas de trabajo, no traslado

  // ---------- Maps ----------
  ORIGEN_DEFAULT: 'JC4X+PMJ, 17 Calle, Cdad. de Guatemala',

  // ---------- Zona horaria ----------
  TIMEZONE: 'America/Guatemala',

  // ---------- Misc ----------
  MARGEN_MINUTOS_INICIO_HOY: 60,
  VENTANA_BUSQUEDA_DIAS: 10,
  COLOR_BORRADOR: '8',
  COLOR_CONFIRMADO: '10',

  // ---------- Reglas de prioridad ----------
  PRIORIDAD: {
    Urgente: {
      ventanaDias: 2,
      ventanaAmpliada: 3,
      pesoDemora: 50,
      toleranciaHoraSugerida: 4,
      permiteDesplazarMediaBaja: true,
      colorCalendar: '11'
    },
    Media: {
      ventanaDias: 3,
      ventanaAmpliada: 5,
      pesoDemora: 10,
      toleranciaHoraSugerida: 3,
      permiteDesplazarMediaBaja: false,
      colorCalendar: '6'
    },
    Baja: {
      ventanaDias: 10,
      ventanaAmpliada: 20,
      pesoDemora: 2,
      toleranciaHoraSugerida: 8,
      permiteDesplazarMediaBaja: false,
      bonusAgrupamiento: 30,
      colorCalendar: '2'
    }
  },

  // ---------- Viáticos (en Quetzales) ----------
  VIATICOS: {
    desayuno: 50,
    almuerzo: 50,
    cena: 50,
    hotel: 200
  },

  // ---------- Modo Hotel ----------
  UMBRAL_HOTEL_HORAS: 5   // si traslado ida > 5h, requiere pernoctar
};

function getMapsApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
  if (!key) throw new Error('Falta MAPS_API_KEY en Script Properties');
  return key;
}

function getCalendarCentral_() {
  const cal = CalendarApp.getCalendarById(CFG.CALENDAR_CENTRAL_ID);
  if (!cal) throw new Error(`No se puede abrir el calendario central "${CFG.CALENDAR_CENTRAL_ID}".`);
  return cal;
}

function getConfigPrioridad_(prioridad) {
  const p = CFG.PRIORIDAD[prioridad];
  if (!p) {
    console.log(`Prioridad "${prioridad}" no reconocida, usando Media`);
    return CFG.PRIORIDAD.Media;
  }
  return p;
}
