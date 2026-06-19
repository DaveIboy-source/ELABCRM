/**
 * =====================================================================
 * NOTIFICACIONES.gs — v1 — anti-duplicados de correos
 * =====================================================================
 * Registra cada notificación automática enviada para no repetirla.
 *
 * Uso:
 *   if (yaSeNotifico_('TCK123', 'rebote_email')) return;
 *   GmailApp.sendEmail(...);
 *   marcarNotificado_('TCK123', 'rebote_email');
 *
 * El registro se guarda en una pestaña oculta "NotificacionesEnviadas".
 */

const TAB_NOTIFICACIONES = 'NotificacionesEnviadas';
const VENTANA_DUPLICADOS_DIAS = 7;  // ignora repetidas dentro de los últimos 7 días

function yaSeNotifico_(ticketId, tipoEvento) {
  const hoja = getHojaNotificaciones_();
  const data = hoja.getDataRange().getValues();
  const limite = new Date(Date.now() - VENTANA_DUPLICADOS_DIAS * 24 * 3600 * 1000);

  for (let i = 1; i < data.length; i++) {
    const [timestamp, tid, tipo] = data[i];
    if (!timestamp) continue;
    if (new Date(timestamp) < limite) continue;
    if (tid === ticketId && tipo === tipoEvento) return true;
  }
  return false;
}

function marcarNotificado_(ticketId, tipoEvento) {
  const hoja = getHojaNotificaciones_();
  hoja.appendRow([new Date(), ticketId, tipoEvento]);
}

function getHojaNotificaciones_() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  let hoja = ss.getSheetByName(TAB_NOTIFICACIONES);
  if (!hoja) {
    hoja = ss.insertSheet(TAB_NOTIFICACIONES);
    hoja.appendRow(['Timestamp', 'Ticket_ID', 'Tipo_Evento']);
    hoja.hideSheet();
  }
  return hoja;
}

/**
 * Wrapper unificado para enviar correos automáticos del sistema.
 * Aplica chequeo anti-duplicados.
 *
 * @returns {boolean} true si se envió, false si fue saltado por duplicado
 */
function enviarCorreoUnico_(destinatario, asunto, cuerpo, ticketId, tipoEvento) {
  if (yaSeNotifico_(ticketId, tipoEvento)) {
    console.log(`Correo "${tipoEvento}" para ${ticketId} ya fue enviado, saltando`);
    return false;
  }

  try {
    GmailApp.sendEmail(destinatario, asunto, cuerpo, {
      name: 'Sistema de Tickets'
    });
    marcarNotificado_(ticketId, tipoEvento);
    return true;
  } catch (err) {
    log_('ERROR', `enviarCorreoUnico_ ${ticketId}/${tipoEvento}: ${err.message}`);
    // Marcamos como notificado igual para no reintentar correos rebotados
    marcarNotificado_(ticketId, tipoEvento);
    return false;
  }
}
