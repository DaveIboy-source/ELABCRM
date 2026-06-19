/**
 * DiagnosticoEmail.gs - verifica el estado del envío de correos
 * Pega este archivo y corre debugEmail()
 */

function debugEmail() {
  console.log('=== DIAGNÓSTICO DE EMAIL ===\n');

  // 1. Verificar quota
  const quotaRestante = MailApp.getRemainingDailyQuota();
  console.log(`Quota restante de correos hoy: ${quotaRestante}`);

  if (quotaRestante === 0) {
    console.log('\n*** PROBLEMA CRÍTICO: Te quedaste sin quota de correos hoy ***');
    console.log('El sistema NO puede enviar correos hasta mañana.');
    console.log('Límite Workspace: 1500/día. Si es cuenta gratuita: 100/día.\n');
  } else if (quotaRestante < 20) {
    console.log(`\nAdvertencia: quedan pocos correos disponibles (${quotaRestante})`);
  }

  // 2. Verificar técnicos cargados
  console.log('\n=== TÉCNICOS ACTIVOS EN LA SHEET ===');
  try {
    const tecnicos = cargarTecnicos_();
    for (const tec of tecnicos) {
      console.log(`  ${tec.nombre} | ${tec.email} | ${tec.activo ? 'activo' : 'INACTIVO'} | base: ${tec.base}`);
    }
  } catch(e) {
    console.log(`ERROR cargando técnicos: ${e.message}`);
  }

  // 3. Test de envío
  console.log('\n=== TEST DE ENVÍO ===');
  console.log('Probando enviar correo de prueba a jefatura...');
  try {
    GmailApp.sendEmail(
      CFG.EMAIL_JEFATURA,
      '[TEST DE EMAIL] Sistema de Tickets - ' + new Date().toLocaleString(),
      'Este es un correo de prueba para verificar que el sistema puede enviar correos.\n\nSi recibes este correo, el envío funciona correctamente.\n\nQuota restante: ' + quotaRestante,
      { name: 'Sistema de Tickets' }
    );
    console.log('✓ Correo de test enviado a jefatura');
    console.log('Revisa tu bandeja en 1 minuto. Si no llega, hay problema de envío.');
  } catch(e) {
    console.log('✗ ERROR al enviar test: ' + e.message);
  }
}

/**
 * Verifica un ticket específico: dónde está el técnico asignado y si su email es válido.
 * CAMBIA el ticketId por el del ticket afectado antes de correr.
 */
function debugTicketEmail() {
  const ticketId = 'TCK20260520'; // CAMBIA ESTE ID por el del ticket afectado

  console.log('=== DIAGNÓSTICO DE TICKET ' + ticketId + ' ===\n');

  const ctx = buscarTicketPorId_(ticketId);
  if (!ctx) {
    console.log('Ticket ' + ticketId + ' NO encontrado en Sheet');
    return;
  }

  console.log('Fila en Sheet: ' + ctx.fila);
  console.log('Cliente: ' + ctx.cliente);
  console.log('Estado: ' + ctx.hoja.getRange(ctx.fila, ctx.hoja.getLastColumn()).getValue());
  console.log('Técnico (objeto): ' + JSON.stringify(ctx.tecnico));

  if (ctx.tecnico) {
    console.log('  Nombre: "' + ctx.tecnico.nombre + '"');
    console.log('  Email: "' + ctx.tecnico.email + '"');
    console.log('  ¿Email vacío?: ' + (!ctx.tecnico.email || ctx.tecnico.email.trim() === ''));
    console.log('  ¿Email válido?: ' + /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(ctx.tecnico.email));
  } else {
    console.log('*** PROBLEMA: ctx.tecnico es null ***');
  }

  if (!ctx.evento) {
    console.log('*** PROBLEMA: ctx.evento es null (no hay evento Calendar) ***');
  } else {
    console.log('Evento Calendar OK: ' + ctx.evento.getTitle());
  }
}
