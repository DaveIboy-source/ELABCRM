/**
 * =====================================================================
 * DIAGNOSTICO.gs  —  Prueba cada pieza del sistema por separado
 * =====================================================================
 * Corre cada función desde el editor y mira el registro de ejecución.
 * La que falle te dice exactamente dónde está el problema.
 */

// 1️⃣ Prueba que la API Key de Maps esté guardada
function test1_verificarApiKey() {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
    if (!key) {
      console.log('❌ FALLA: No hay MAPS_API_KEY en Script Properties');
      console.log('   Ve a: Configuración del proyecto → Propiedades del script');
      return;
    }
    console.log('✅ API Key encontrada. Empieza con: ' + key.substring(0, 8) + '...');
    console.log('   Longitud: ' + key.length + ' caracteres');
  } catch (e) {
    console.log('❌ FALLA: ' + e.message);
  }
}

// 2️⃣ Prueba Geocoding API
function test2_geocoding() {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
    const direccion = '6a avenida 10-00 zona 9, Guatemala';
    const url = 'https://maps.googleapis.com/maps/api/geocode/json'
      + '?address=' + encodeURIComponent(direccion)
      + '&region=gt&key=' + key;

    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    console.log('Status HTTP: ' + res.getResponseCode());
    console.log('Status API: ' + data.status);

    if (data.status === 'OK') {
      console.log('✅ Geocoding funciona. Coordenadas: ' + JSON.stringify(data.results[0].geometry.location));
    } else {
      console.log('❌ FALLA en Geocoding: ' + data.status);
      console.log('   Error: ' + (data.error_message || 'sin mensaje'));
      if (data.status === 'REQUEST_DENIED') {
        console.log('   → La API Key no tiene habilitada Geocoding API,');
        console.log('     o tiene restricciones que bloquean Apps Script.');
      }
    }
  } catch (e) {
    console.log('❌ FALLA: ' + e.message);
  }
}

// 3️⃣ Prueba Distance Matrix API
function test3_distanceMatrix() {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
    const url = 'https://maps.googleapis.com/maps/api/distancematrix/json'
      + '?origins=' + encodeURIComponent('Zona 10, Guatemala')
      + '&destinations=' + encodeURIComponent('Zona 9, Guatemala')
      + '&mode=driving&key=' + key;

    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    console.log('Status API: ' + data.status);

    if (data.status === 'OK') {
      const el = data.rows[0].elements[0];
      console.log('✅ Distance Matrix funciona');
      console.log('   Duración: ' + el.duration.text);
      console.log('   Distancia: ' + el.distance.text);
    } else {
      console.log('❌ FALLA en Distance Matrix: ' + data.status);
      console.log('   Error: ' + (data.error_message || 'sin mensaje'));
      if (data.status === 'REQUEST_DENIED') {
        console.log('   → La API Key no tiene habilitada Distance Matrix API,');
        console.log('     o tiene restricciones que bloquean Apps Script.');
      }
    }
  } catch (e) {
    console.log('❌ FALLA: ' + e.message);
  }
}

// 4️⃣ Prueba carga de técnicos
function test4_cargarTecnicos() {
  try {
    const tecnicos = cargarTecnicos_();
    console.log('✅ Se cargaron ' + tecnicos.length + ' técnicos');
    tecnicos.forEach((t, i) => {
      console.log(`  ${i + 1}. ${t.nombre} | ${t.email} | Base: ${t.base} | Activo: ${t.activo} | Deptos: ${t.departamentos.length || '(comodín)'}`);
    });
    const activos = tecnicos.filter(t => t.activo);
    console.log(`   → Activos: ${activos.length}`);
  } catch (e) {
    console.log('❌ FALLA: ' + e.message);
    console.log('   Revisa la pestaña Tecnicos: headers Nombre, Email, Base, Departamentos, Calendar_ID, Activo');
  }
}

// 5️⃣ Prueba el flujo completo simulado
function test5_flujoCompleto() {
  try {
    const ticketFake = {
      id: 'TEST_DIAGNOSTICO',
      direccion: '6a avenida 10-00 zona 9, Guatemala',
      departamento: 'Guatemala',
      fechaSugerida: null
    };

    console.log('→ Geocodificando dirección...');
    const destino = geocodificar_(ticketFake.direccion);
    ticketFake.lat = destino.lat;
    ticketFake.lng = destino.lng;
    console.log('  ✅ Coords: ' + destino.lat + ', ' + destino.lng);

    console.log('→ Cargando técnicos...');
    const tecnicos = cargarTecnicos_();
    console.log('  ✅ ' + tecnicos.length + ' técnicos cargados');

    console.log('→ Evaluando candidatos...');
    const candidatos = evaluarCandidatos_(ticketFake, tecnicos);
    console.log('  ✅ ' + candidatos.length + ' candidatos con capacidad');

    if (candidatos.length === 0) {
      console.log('❌ Ningún técnico con capacidad. Posibles causas:');
      console.log('   - Ningún técnico activo');
      console.log('   - Departamento no coincide (revisa tildes/mayúsculas)');
      console.log('   - Distance Matrix falló para todos');
      console.log('   - Ningún hueco libre en 8am-5pm');
    } else {
      const mejor = candidatos[0];
      console.log('✅ Mejor candidato: ' + mejor.tecnico.nombre);
      console.log('   Traslado: ' + mejor.trasladoHoras.toFixed(2) + ' h');
      console.log('   Fecha slot: ' + mejor.fechaSlot);
    }
  } catch (e) {
    console.log('❌ FALLA: ' + e.message);
    console.log('Stack: ' + e.stack);
  }
}
