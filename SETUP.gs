/**
 * =====================================================================
 * SETUP.gs — v4 — sin emojis
 * =====================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tickets')
    .addItem('Detectar cambios en Calendar', 'detectarCambiosEnCalendar')
    .addItem('Procesar RSTs de Drive', 'procesarRSTsNuevos')
    .addSeparator()
    .addItem('Actualizar listas Cliente/Equipo del Form', 'actualizarListasDelForm')
    .addSeparator()
    .addItem('Actualizar Dashboard', 'actualizarDashboard')
    .addItem('Procesar bandeja de entrada', 'procesarBandejaEntrada')
    .addSeparator()
    .addItem('Inicializar sistema (solo 1 vez)', 'inicializar')
    .addToUi();
}

function inicializar() {
  crearEstructuraSheets_();
  crearEtiquetasGmail_();
  instalarTriggers_();
  SpreadsheetApp.getActive().toast('Sistema inicializado. Revisa la pestaña Tecnicos.', 'Setup', 10);
}

function crearEstructuraSheets_() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);

  let t = ss.getSheetByName(CFG.TAB_TECNICOS);
  if (!t) {
    t = ss.insertSheet(CFG.TAB_TECNICOS);
    t.getRange(1, 1, 1, 5).setValues([[
      'Nombre', 'Email', 'Base', 'Departamentos', 'Activo'
    ]]).setFontWeight('bold').setBackground('#0b5394').setFontColor('#fff');
    t.getRange(2, 1, 1, 5).setValues([[
      'Ejemplo: Juan Pérez',
      'juan@tudominio.com',
      'Guatemala, Guatemala',
      'Guatemala,Sacatepéquez',
      true
    ]]);
    t.setColumnWidths(1, 5, 180);
  }

  let l = ss.getSheetByName(CFG.TAB_LOG);
  if (!l) {
    l = ss.insertSheet(CFG.TAB_LOG);
    l.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Nivel', 'Mensaje']])
     .setFontWeight('bold').setBackground('#444').setFontColor('#fff');
  }

  const tk = ss.getSheetByName(CFG.TAB_TICKETS);
  if (tk) {
    const lastCol = tk.getLastColumn();
    const headers = tk.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headers.includes('Ticket_ID')) {
      tk.getRange(1, lastCol + 1, 1, 6).setValues([[
        'Ticket_ID', 'Técnico', 'Traslado_h', 'Fecha_Slot', 'Event_ID', 'Estado'
      ]]).setFontWeight('bold').setBackground('#0b5394').setFontColor('#fff');
    }
  }
}

function crearEtiquetasGmail_() {
  if (!GmailApp.getUserLabelByName(CFG.ETIQUETA_GMAIL_APROBACIONES))
    GmailApp.createLabel(CFG.ETIQUETA_GMAIL_APROBACIONES);
  if (!GmailApp.getUserLabelByName(CFG.ETIQUETA_GMAIL_TECNICOS))
    GmailApp.createLabel(CFG.ETIQUETA_GMAIL_TECNICOS);
}

function instalarTriggers_() {
  ScriptApp.getProjectTriggers().forEach(tr => {
    const fn = tr.getHandlerFunction();
    if (['onFormSubmit', 'procesarBandejaEntrada', 'actualizarDashboard'].includes(fn)) {
      ScriptApp.deleteTrigger(tr);
    }
  });

  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('procesarBandejaEntrada').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('actualizarDashboard').timeBased().everyHours(1).create();
}
