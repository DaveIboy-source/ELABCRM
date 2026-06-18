/**
 * =====================================================================
 * CONFIGURAR_FORM.gs — v5 — listas Cliente/Equipo cargadas desde Sheet
 * =====================================================================
 */

const FORM_ID = '1eNWi7etBayr7phq9ljumIyywzQtso8FPJmuhMffI4js';

function configurarFormulario() {
  console.log('Iniciando configuración del Form...');
  const form = FormApp.openById(FORM_ID);

  const items = form.getItems();
  console.log(`Borrando ${items.length} campos existentes...`);
  items.forEach(item => form.deleteItem(item));

  form.setTitle('Solicitud de Ticket de Servicio Técnico');
  form.setDescription(
    'Llena este formulario para generar un nuevo ticket. ' +
    'El sistema asignará automáticamente un técnico disponible ' +
    'según prioridad, zona, capacidad y traslado.'
  );

  form.setRequireLogin(true);
  form.setCollectEmail(true);
  form.setAllowResponseEdits(false);
  form.setShowLinkToRespondAgain(true);

  // Cargar listas dinámicas ANTES de construir el Form
  const listaClientes = cargarListaDeSheet_('Clientes');
  const listaEquipos = cargarListaDeSheet_('Equipos');

  console.log(`Clientes cargados: ${listaClientes.length}`);
  console.log(`Equipos cargados: ${listaEquipos.length}`);

  // --- Datos del vendedor ---
  console.log('Agregando campo: Vendedor');
  form.addTextItem()
      .setTitle('Vendedor')
      .setHelpText('Tu nombre completo')
      .setRequired(true);

  console.log('Agregando campo: Correo vendedor');
  form.addTextItem()
      .setTitle('Correo vendedor')
      .setHelpText('Tu correo — aquí recibirás la confirmación')
      .setRequired(true)
      .setValidation(
        FormApp.createTextValidation()
          .requireTextIsEmail()
          .setHelpText('Debe ser un correo válido')
          .build()
      );

  // --- Datos del cliente ---
  console.log('Agregando campo: Cliente (desplegable)');
  form.addListItem()
      .setTitle('Cliente')
      .setHelpText('Selecciona el cliente de la lista (si falta, avisa a jefatura para agregarlo)')
      .setRequired(true)
      .setChoiceValues(listaClientes);

  console.log('Agregando campo: Teléfono cliente');
  form.addTextItem()
      .setTitle('Teléfono cliente')
      .setHelpText('Obligatorio si el tipo de servicio es "Llamada". Para visitas es recomendable pero no estricto.')
      .setRequired(true);

  console.log('Agregando campo: Dirección');
  form.addParagraphTextItem()
      .setTitle('Dirección')
      .setHelpText('Dirección del cliente. Si el servicio es "Llamada", puedes escribir "N/A".')
      .setRequired(true);

  console.log('Agregando campo: Departamento');
  form.addListItem()
      .setTitle('Departamento')
      .setHelpText('Departamento donde se encuentra el cliente (para llamadas puedes elegir cualquiera).')
      .setRequired(true)
      .setChoiceValues([
        'Guatemala','Sacatepéquez','Chimaltenango','Escuintla','Santa Rosa',
        'Sololá','Totonicapán','Quetzaltenango','Suchitepéquez','Retalhuleu',
        'San Marcos','Huehuetenango','Quiché','Baja Verapaz','Alta Verapaz',
        'Petén','Izabal','Zacapa','Chiquimula','Jalapa','Jutiapa','El Progreso'
      ]);

  // --- Datos del servicio ---
  console.log('Agregando campo: Equipo (selección múltiple)');
  form.addCheckboxItem()
      .setTitle('Equipo')
      .setHelpText('Selecciona uno o varios equipos. Si vas a revisar varios equipos del mismo cliente en una sola visita, márcalos todos.')
      .setRequired(true)
      .setChoiceValues(listaEquipos);

  console.log('Agregando campo: Tipo de servicio');
  form.addMultipleChoiceItem()
      .setTitle('Tipo de servicio')
      .setHelpText(
        'Llamada = atención remota inmediata (sin traslado, sin agenda). ' +
        'Revisión Equipo = revisión remota inmediata (sin traslado, sin agenda). ' +
        'Instalación/Mantenimiento/Reparación/Capacitación = visita en sitio.'
      )
      .setRequired(true)
      .setChoiceValues([
        'Llamada',
        'Revisión Equipo',
        'Instalación',
        'Mantenimiento',
        'Reparación',
        'Capacitación'
      ]);

  console.log('Agregando campo: Prioridad');
  form.addListItem()
      .setTitle('Prioridad')
      .setHelpText(
        'Urgente = detiene la operación del cliente (hoy o mañana). ' +
        'Media = mantenimiento correctivo (hasta 3 días hábiles). ' +
        'Baja = mantenimiento preventivo (se agrupa con rutas). ' +
        'Para "Llamada" la prioridad se ignora (siempre es inmediata).'
      )
      .setRequired(true)
      .setChoiceValues(['Urgente', 'Media', 'Baja']);

  console.log('Agregando campo: Fecha sugerida');
  form.addDateItem()
      .setTitle('Fecha sugerida')
      .setHelpText('Fecha preferida para la visita. No aplica para llamadas.')
      .setRequired(true);

  console.log('Agregando campo: Hora sugerida');
  form.addListItem()
      .setTitle('Hora sugerida')
      .setHelpText(
        'Hora preferida de inicio. Si el técnico no está libre a esa hora, ' +
        'el sistema busca el hueco más cercano. No aplica para llamadas.'
      )
      .setRequired(true)
      .setChoiceValues([
        '08:00', '09:00', '10:00', '11:00', '12:00',
        '13:00', '14:00', '15:00', '16:00'
      ]);

  console.log('Agregando campo: Notas');
  form.addParagraphTextItem()
      .setTitle('Notas')
      .setHelpText('Información adicional relevante (opcional).')
      .setRequired(false);

  console.log('Formulario configurado correctamente');
  console.log('URL pública: ' + form.getPublishedUrl());
}

/**
 * Carga una lista desde una pestaña de la Sheet.
 * Toma todos los valores de columna A (desde fila 2), limpia, ordena
 * alfabéticamente y elimina duplicados.
 */
function cargarListaDeSheet_(nombrePestaña) {
  const ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  const hoja = ss.getSheetByName(nombrePestaña);
  if (!hoja) {
    throw new Error(`Falta la pestaña "${nombrePestaña}" en la Sheet. ` +
                    `Créala con header en A1 y los valores debajo.`);
  }

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    throw new Error(`La pestaña "${nombrePestaña}" está vacía. ` +
                    `Agrega valores desde la fila 2 en la columna A.`);
  }

  const valores = hoja.getRange(2, 1, ultimaFila - 1, 1)
    .getValues()
    .map(r => String(r[0]).trim())
    .filter(Boolean);

  // Eliminar duplicados y ordenar alfabéticamente
  const unicos = Array.from(new Set(valores)).sort((a, b) =>
    a.localeCompare(b, 'es', { sensitivity: 'base' })
  );

  return unicos;
}

/**
 * Función RÁPIDA para actualizar SOLO los desplegables de Cliente y Equipo
 * sin reconstruir todo el Form. Úsala cuando agregues clientes/equipos a las
 * pestañas.
 */
function actualizarListasDelForm() {
  console.log('Actualizando listas Cliente/Equipo del Form...');
  const form = FormApp.openById(FORM_ID);

  const listaClientes = cargarListaDeSheet_('Clientes');
  const listaEquipos = cargarListaDeSheet_('Equipos');

  console.log(`Clientes: ${listaClientes.length}`);
  console.log(`Equipos: ${listaEquipos.length}`);

  let actualizados = 0;
  const items = form.getItems();

  for (const item of items) {
    const titulo = item.getTitle();
    if (titulo === 'Cliente' && item.getType() === FormApp.ItemType.LIST) {
      item.asListItem().setChoiceValues(listaClientes);
      actualizados++;
      console.log('  Cliente actualizado');
    }
    if (titulo === 'Equipo' && item.getType() === FormApp.ItemType.CHECKBOX) {
      item.asCheckboxItem().setChoiceValues(listaEquipos);
      actualizados++;
      console.log('  Equipo actualizado (checkbox)');
    }
    // Compatibilidad: si Equipo aún es lista vieja, también lo actualizamos
    if (titulo === 'Equipo' && item.getType() === FormApp.ItemType.LIST) {
      item.asListItem().setChoiceValues(listaEquipos);
      actualizados++;
      console.log('  Equipo actualizado (lista)');
    }
  }

  console.log(`${actualizados} campo(s) actualizado(s)`);
  if (actualizados < 2) {
    console.log('ADVERTENCIA: no se encontraron ambos campos. ' +
                'Si faltan, corre configurarFormulario() para reconstruir el Form completo.');
  }
}
