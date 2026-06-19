/**
 * Diagnóstico: busca correos de aprobación en Gmail
 * Corre test8_buscarCorreosAprobacion
 */

function test8_buscarCorreosAprobacion() {
  console.log('=== Test 8: búsqueda de correos de aprobación ===');
  console.log('Email jefatura configurado: ' + CFG.EMAIL_JEFATURA);
  console.log('Usuario actual: ' + Session.getActiveUser().getEmail());
  console.log('');

  // Query 1: como lo hace el script original
  const q1 = `subject:"APROBAR TCK" is:unread -from:me`;
  const h1 = GmailApp.search(q1, 0, 10);
  console.log(`Query 1: ${q1}`);
  console.log(`  → ${h1.length} hilos`);
  h1.forEach((hilo, i) => {
    const msgs = hilo.getMessages();
    const ult = msgs[msgs.length - 1];
    console.log(`     ${i + 1}. "${hilo.getFirstMessageSubject()}" | de: ${ult.getFrom()} | leído: ${!ult.isUnread()}`);
  });
  console.log('');

  // Query 2: sin filtro de leído
  const q2 = `subject:"APROBAR TCK"`;
  const h2 = GmailApp.search(q2, 0, 10);
  console.log(`Query 2: ${q2}`);
  console.log(`  → ${h2.length} hilos`);
  h2.forEach((hilo, i) => {
    const msgs = hilo.getMessages();
    const ult = msgs[msgs.length - 1];
    console.log(`     ${i + 1}. "${hilo.getFirstMessageSubject()}" | de: ${ult.getFrom()} | leído: ${!ult.isUnread()} | ${msgs.length} mensajes`);
  });
  console.log('');

  // Query 3: los más recientes del usuario (respuestas)
  const q3 = `from:me subject:"APROBAR TCK"`;
  const h3 = GmailApp.search(q3, 0, 10);
  console.log(`Query 3: ${q3}`);
  console.log(`  → ${h3.length} hilos (correos enviados por ti al hilo)`);

  // Query 4: hilos con el texto APROBADO recientes
  const q4 = `APROBADO newer_than:1d`;
  const h4 = GmailApp.search(q4, 0, 10);
  console.log(`Query 4: ${q4}`);
  console.log(`  → ${h4.length} hilos`);
  h4.forEach((hilo, i) => {
    console.log(`     ${i + 1}. "${hilo.getFirstMessageSubject()}"`);
  });
}
