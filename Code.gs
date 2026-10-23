/**
 * Copello - pasar correos etiquetados a una hoja de cálculo.
 * No envía correos ni modifica los mensajes.
 *
 * Configuración:
 * 1. Crear una etiqueta de Gmail llamada Copello_Para_Excel.
 * 2. En script.google.com, crear un proyecto y pegar este archivo.
 * 3. Ejecutar prepararCopello() y autorizar permisos.
 * 4. Ejecutar importarCorreosCopello() para probar.
 *
 * Esta versión guarda los campos generales del correo y el cuerpo completo.
 * La extracción de campos específicos se ajusta cuando se define el formato
 * real del correo y las columnas exactas de la planilla.
 */

const CONFIG = {
  etiquetaGmail: 'Copello_Para_Excel',
  nombrePlanilla: 'Copello_Registro_Correos',
  nombreHoja: 'Correos'
};

function prepararCopello() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('COPPELLO_SPREADSHEET_ID');
  let spreadsheet;

  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      spreadsheet = null;
    }
  }

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(CONFIG.nombrePlanilla);
    props.setProperty('COPPELLO_SPREADSHEET_ID', spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName(CONFIG.nombreHoja);
  if (!sheet) sheet = spreadsheet.getSheets()[0].setName(CONFIG.nombreHoja);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Fecha del correo',
      'Remitente',
      'Asunto',
      'Cuerpo del correo',
      'Enlace al correo',
      'ID del mensaje',
      'Fecha de importación'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }

  if (!GmailApp.getUserLabelByName(CONFIG.etiquetaGmail)) {
    GmailApp.createLabel(CONFIG.etiquetaGmail);
  }

  Logger.log('Planilla: ' + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

function importarCorreosCopello() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('COPPELLO_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Primero ejecutá prepararCopello().');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(CONFIG.nombreHoja);
  const label = GmailApp.getUserLabelByName(CONFIG.etiquetaGmail);
  if (!label) throw new Error('No existe la etiqueta ' + CONFIG.etiquetaGmail);

  const existingIds = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 6, sheet.getLastRow() - 1, 1)
      .getDisplayValues()
      .flat()
      .filter(String)
      .forEach(id => existingIds.add(id));
  }

  const threads = label.getThreads(0, 100);
  const rows = [];
  const importedAt = new Date();

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const id = message.getId();
      if (existingIds.has(id)) return;

      rows.push([
        message.getDate(),
        message.getFrom(),
        message.getSubject(),
        message.getPlainBody(),
        'https://mail.google.com/mail/u/0/#all/' + id,
        id,
        importedAt
      ]);
      existingIds.add(id);
    });
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
  }

  Logger.log('Correos nuevos importados: ' + rows.length);
  Logger.log('Planilla: ' + spreadsheet.getUrl());
  return rows.length;
}

/**
 * Opcional: crea un disparador que revisa la etiqueta cada 15 minutos.
 * Ejecutarlo solo después de probar la importación manual.
 */
function instalarRevisionAutomaticaCopello() {
  const exists = ScriptApp.getProjectTriggers()
    .some(trigger => trigger.getHandlerFunction() === 'importarCorreosCopello');
  if (!exists) {
    ScriptApp.newTrigger('importarCorreosCopello')
      .timeBased()
      .everyMinutes(15)
      .create();
  }
}

/** Quita únicamente el disparador de esta importación. */
function quitarRevisionAutomaticaCopello() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'importarCorreosCopello') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
