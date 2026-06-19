/**
 * DiagnosticoEmail.gs - función de diagnóstico para verificar
 * que el correo del vendedor se está leyendo correctamente
 */

function debugCorreoVendedor() {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(CFG.TAB_TICKETS);

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    console.log('No hay tickets en la Sheet');
    return;
  }

  console.log(`=== Debug del correo del vendedor en fila ${ultimaFila} ===\n`);

  // 1. Ver headers crudos
  const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  console.log('HEADERS DE LA SHEET:');
  headers.forEach((h, i) => {
    const norm = normalizarHeader_(h);
    console.log(`  Col ${i + 1}: "${h}" -> normalizado: "${norm}"`);
  });

  // 2. Ver valores crudos
  const valores = hoja.getRange(ultimaFila, 1, 1, hoja.getLastColumn()).getValues()[0];
  console.log('\nVALORES DE LA FILA:');
  headers.forEach((h, i) => {
    const norm = normalizarHeader_(h);
    console.log(`  "${h}" (${norm}): "${valores[i]}"`);
  });

  // 3. Construir objeto t como lo hace leerTicketDeFila_
  const t = {};
  headers.forEach((h, i) => t[normalizarHeader_(h)] = valores[i]);

  console.log('\nOBJETO t (lo que ve el código):');
  Object.keys(t).forEach(k => {
    console.log(`  t.${k} = "${t[k]}"`);
  });

  // 4. Mostrar el resultado del fallback
  const emailVendedor = t.correo_vendedor || t.email_vendedor;
  console.log(`\n*** RESULTADO ***`);
  console.log(`emailVendedor = "${emailVendedor}"`);

  if (!emailVendedor) {
    console.log('❌ PROBLEMA: emailVendedor está vacío');
    console.log('Verifica que la columna "Correo vendedor" tenga un email en esa fila');
  } else {
    console.log('✅ Correo leído correctamente');

    // 5. Probar envío real
    console.log('\nIntentando enviar correo de prueba...');
    try {
      GmailApp.sendEmail(
        emailVendedor,
        '[PRUEBA] Sistema de tickets - test de correo',
        'Este es un correo de prueba enviado por debugCorreoVendedor.\n\n' +
        'Si lo recibes, el sistema puede enviarte correos correctamente.',
        { name: 'Sistema de Tickets' }
      );
      console.log(`✅ Correo enviado a ${emailVendedor}`);
    } catch (err) {
      console.log(`❌ Error al enviar: ${err.message}`);
    }
  }
}
