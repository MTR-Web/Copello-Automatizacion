/**
 * COPELLO — Registro de comprobantes desde Gmail.
 * Solo procesa correos etiquetados "Pagos" o la etiqueta antigua configurada.
 * Guarda adjuntos en Drive, evita duplicados y crea borradores; nunca envía correos.
 *
 * Para OCR de PDF/imagenes: en Apps Script > Servicios + > Drive API (v2).
 */
const CONFIG = {
  etiquetaPagos: 'Pagos',
  etiquetaLegacy: 'Copello_Para_Excel',
  etiquetaQuejas: 'Quejas',
  nombrePlanilla: 'Copello_Registro_Correos',
  nombreHoja: 'Comprobantes',
  nombreHojaCorreos: 'Correos',
  nombreCarpetaAdjuntos: 'Copello_Comprobantes',
  limiteConversaciones: 100,
  crearBorradorRespuesta: true,
  textoBorrador: 'Hola, recibimos el comprobante enviado. Lo revisaremos y lo registraremos si corresponde. Gracias.'
};

const HEADERS = [
  'Fecha del correo','Remitente','Asunto','Importe detectado','Fecha del comprobante',
  'N.º de operación','Pagador detectado','Archivo adjunto','Enlace al archivo',
  'Texto extraído','Estado de revisión','Enlace al correo','ID del mensaje','Fecha de importación',
  'Borrador creado'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Copello | Registro de comprobantes')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function autorizarDriveCopello() {
  const props = PropertiesService.getUserProperties();
  let folderId = props.getProperty('COPPELLO_FOLDER_ID');
  let folder;
  if (folderId) {
    folder = DriveApp.getFolderById(folderId);
    folder.getName();
  } else {
    folder = DriveApp.getFoldersByName(CONFIG.nombreCarpetaAdjuntos);
    folder = folder.hasNext() ? folder.next() : DriveApp.createFolder(CONFIG.nombreCarpetaAdjuntos);
    props.setProperty('COPPELLO_FOLDER_ID', folder.getId());
  }
  return {ok:true,folderUrl:folder.getUrl(),message:'Acceso a Drive autorizado.'};
}

function prepararCopello() {
  const props = PropertiesService.getUserProperties();
  let id = props.getProperty('COPPELLO_SPREADSHEET_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.nombrePlanilla);
    props.setProperty('COPPELLO_SPREADSHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(CONFIG.nombreHoja);
  if (!sh) sh = ss.insertSheet(CONFIG.nombreHoja);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#18352e').setFontColor('#ffffff');
    sh.getRange(1,1,sh.getMaxRows(),1).setNumberFormat('dd/mm/yyyy hh:mm');
    sh.autoResizeColumns(1,HEADERS.length);
  }
  let mailSh = ss.getSheetByName(CONFIG.nombreHojaCorreos);
  if (!mailSh) mailSh = ss.insertSheet(CONFIG.nombreHojaCorreos);
  if (mailSh.getLastRow() === 0) {
    mailSh.appendRow(['Fecha del correo','Remitente','Asunto','Cuerpo del correo','Enlace al correo','ID del mensaje','Fecha de importación','Etiqueta']);
    mailSh.setFrozenRows(1);
    mailSh.getRange(1,1,1,8).setFontWeight('bold').setBackground('#18352e').setFontColor('#ffffff');
  }
  [CONFIG.etiquetaPagos, CONFIG.etiquetaLegacy, CONFIG.etiquetaQuejas].forEach(name => {
    if (!GmailApp.getUserLabelByName(name)) GmailApp.createLabel(name);
  });
  let folderId = props.getProperty('COPPELLO_FOLDER_ID');
  let folder = null;
  if (folderId) { try { folder = DriveApp.getFolderById(folderId); } catch (e) {} }
  if (!folder) {
    const it = DriveApp.getFoldersByName(CONFIG.nombreCarpetaAdjuntos);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.nombreCarpetaAdjuntos);
    props.setProperty('COPPELLO_FOLDER_ID', folder.getId());
  }
  actualizarPlanillaImpresionCopello_();
  const printSheet = ss.getSheetByName('Planilla para imprimir');
  if (printSheet) { printSheet.showSheet(); ss.setActiveSheet(printSheet); }
  [sh, mailSh].forEach(tab => { if (tab && tab.getName() !== 'Planilla para imprimir') tab.hideSheet(); });
  ss.getSheets().filter(tab => tab.getName() === 'Sheet1' && ss.getSheets().length > 1).forEach(tab => ss.deleteSheet(tab));
  return {url:ss.getUrl(),label:CONFIG.etiquetaPagos,folderUrl:folder.getUrl()};
}

function importarCorreosCopello() {
  const props = PropertiesService.getUserProperties();
  const ssId = props.getProperty('COPPELLO_SPREADSHEET_ID');
  if (!ssId) throw new Error('Primero tocá "Preparar sistema" en el panel.');
  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName(CONFIG.nombreHoja);
  const mailSh = ss.getSheetByName(CONFIG.nombreHojaCorreos);
  if (!sh || !mailSh) throw new Error('Falta una hoja. Tocá "Preparar sistema" de nuevo.');
  const labels = [CONFIG.etiquetaPagos, CONFIG.etiquetaLegacy]
    .map(n => GmailApp.getUserLabelByName(n)).filter(Boolean);
  if (!labels.length) throw new Error('No se encontraron etiquetas de pagos.');
  const existing = new Set();
  if (sh.getLastRow() > 1) sh.getRange(2,13,sh.getLastRow()-1,1).getDisplayValues().flat().filter(Boolean).forEach(x=>existing.add(x));
  const mailExisting = new Set();
  if (mailSh.getLastRow() > 1) mailSh.getRange(2,6,mailSh.getLastRow()-1,1).getDisplayValues().flat().filter(Boolean).forEach(x=>mailExisting.add(x));
  const propsFolder = props.getProperty('COPPELLO_FOLDER_ID');
  const folder = DriveApp.getFolderById(propsFolder);
  let imported = 0, drafts = 0, attachmentCount = 0, ocrUnavailable = false;
  const now = new Date(), rows = [], mailRows = [], threadsById = new Map();
  labels.forEach(label => label.getThreads(0,CONFIG.limiteConversaciones).forEach(t=>threadsById.set(t.getId(),t)));
  Array.from(threadsById.values()).forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();
      const accountEmail = String(Session.getActiveUser().getEmail() || '').toLowerCase();
      const fromEmail = String(msg.getFrom() || '').toLowerCase();
      // No importar respuestas propias: las conversaciones etiquetadas pueden incluirlas.
      if (accountEmail && fromEmail.includes(accountEmail)) return;
      if (existing.has(msgId)) return;
      const attachments = msg.getAttachments({includeInlineImages:false,includeAttachments:true});
      const mailLink = 'https://mail.google.com/mail/u/0/#all/' + msgId;
      if (attachments.length === 0) {
        rows.push([msg.getDate(),msg.getFrom(),msg.getSubject(),'','','','','','','','Sin adjunto: revisar',mailLink,msgId,now,'No corresponde']);
      } else {
        attachments.forEach(att => {
          attachmentCount++;
          let savedFile, extracted = '', ocrStatus = 'Pendiente de revisión';
          try {
            const safeName = (msgId + '_' + att.getName()).replace(/[\\/:*?"<>|]/g,'_').slice(0,180);
            savedFile = folder.createFile(att.copyBlob()).setName(safeName);
            try {
              extracted = extraerTextoAdjunto_(savedFile, att.getContentType());
              if (extracted) ocrStatus = 'Texto extraído; verificar datos';
              else ocrStatus = 'No se pudo extraer texto; revisar manualmente';
            } catch (ocrErr) {
              ocrUnavailable = true;
              ocrStatus = 'OCR no disponible; revisar adjunto';
            }
          } catch (saveErr) {
            ocrStatus = 'Error al guardar adjunto: ' + saveErr.message;
          }
          const parsed = extraerCampos_(extracted);
          rows.push([
            msg.getDate(),msg.getFrom(),msg.getSubject(),parsed.importe,parsed.fecha,
            parsed.operacion,parsed.pagador,att.getName(),savedFile ? savedFile.getUrl() : '',
            extracted.slice(0,45000),ocrStatus,mailLink,msgId,now,'Pendiente'
          ]);
        });
      }
      if (!mailExisting.has(msgId)) {
        mailRows.push([msg.getDate(),msg.getFrom(),msg.getSubject(),msg.getPlainBody(),mailLink,msgId,now,thread.getLabels().map(l=>l.getName()).join(', ')]);
        mailExisting.add(msgId);
      }
      if (CONFIG.crearBorradorRespuesta) {
        try {
          thread.createDraftReply(CONFIG.textoBorrador);
          drafts++;
        } catch (e) {
          Logger.log('Error al crear borrador para ' + msgId + ': ' + e.message);
        }
      }
      existing.add(msgId);
      imported++;
    });
  });
  if (rows.length) {
    const start = sh.getLastRow()+1;
    sh.getRange(start,1,rows.length,HEADERS.length).setValues(rows);
    sh.getRange(start,1,rows.length,1).setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange(start,5,rows.length,1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(start,14,rows.length,1).setNumberFormat('dd/mm/yyyy hh:mm');
  }
  if (mailRows.length) mailSh.getRange(mailSh.getLastRow()+1,1,mailRows.length,8).setValues(mailRows);
  actualizarPlanillaImpresionCopello_();
  props.setProperty('COPPELLO_LAST_RUN',new Date().toISOString());
  props.setProperty('COPPELLO_LAST_IMPORT',JSON.stringify({imported,attachmentCount,drafts,ocrUnavailable}));
  return {imported,attachmentCount,drafts,ocrUnavailable,total:Math.max(0,sh.getLastRow()-1),spreadsheetUrl:ss.getUrl()};
}

/** OCR gratuito mediante el servicio avanzado Drive API v2. Activar en Servicios > Drive API. */
function extraerTextoAdjunto_(file, mimeType) {
  const supported = /^application\/pdf$|^image\/(png|jpeg|jpg|webp|gif|bmp|tiff?)$/i.test(mimeType || '');
  if (!supported) return '';
  if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Files.insert) {
    throw new Error('Activá el servicio avanzado Drive API v2.');
  }
  let tempId = null;
  try {
    const blob = file.getBlob();
    const resource = {title:'OCR_TEMP_' + file.getName(),mimeType:'application/vnd.google-apps.document'};
    const docFile = Drive.Files.insert(resource, blob, {convert:true,ocr:true,ocrLanguage:'es'});
    tempId = docFile.id;
    return DocumentApp.openById(tempId).getBody().getText() || '';
  } finally {
    if (tempId) { try { Drive.Files.remove(tempId); } catch (e) {} }
  }
}

function extraerCampos_(text) {
  const t = String(text || '').replace(/\u00a0/g,' ');
  const importeMatch = t.match(/(?:\$\s*|ARS\s*)\d{1,3}(?:[. ]\d{3})*(?:,\d{2})?|(?:importe|total|monto|transferiste|transferencia)[^\d$]{0,25}\$?\s*\d+(?:[.,]\d{2})?/i);
  const fechaMatch = t.match(/\b(\d{1,2}[/-]\d{1,2}[/-](?:20)?\d{2}|\d{4}-\d{2}-\d{2})\b/);
  const opMatch = t.match(/(?:operaci[oó]n|comprobante|transacci[oó]n|referencia|n[uú]mero de operaci[oó]n|id)\s*(?:n[ro.º°]*\s*)?[:#-]?\s*([A-Z0-9-]{5,})/i);
  const payerMatch = t.match(/(?:titular|ordenante|pagador|de\s+parte\s+de|enviado\s+por)\s*[:\-]?\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑa-záéíóúñ .'-]{2,60})/i);
  return {importe:importeMatch ? importeMatch[0].trim() : '',fecha:fechaMatch ? fechaMatch[1] : '',operacion:opMatch ? opMatch[1] : '',pagador:payerMatch ? payerMatch[1].trim() : ''};
}


/** Crea una hoja de cobranza apaisada, lista para imprimir y completar a mano. */
function actualizarPlanillaImpresionCopello_() {
  const props = PropertiesService.getUserProperties();
  const id = props.getProperty('COPPELLO_SPREADSHEET_ID');
  if (!id) return {rows:0};
  const ss = SpreadsheetApp.openById(id);
  const source = ss.getSheetByName(CONFIG.nombreHoja);
  const name = 'Planilla para imprimir';
  let out = ss.getSheetByName(name);
  if (!out) out = ss.insertSheet(name);
  out.clear();
  out.setHiddenGridlines(true);
  out.getRange('A1:J1').merge().setValue('COPELLO ADMINISTRACIÓN: LUNES A VIERNES DE 9 A 18 HS. · CABA');
  out.getRange('A2:J2').merge().setValue('EMAIL: PAGOS Y EXPENSAS');
  out.getRange('A3:J3').merge().setValue('PLANILLA DESCRIPTIVA DE LA COBRANZA  *  LIQUIDACIÓN EXPENSAS  *');
  const headers = ['AF','A4','Apellido y nombre del propietario','Saldo deudor','Fecha de pago','Pago por correo','N.º de cheque','Nombre del banco','Importe del cheque','Observaciones / comprobante'];
  out.getRange(4,1,1,headers.length).setValues([headers]);
  const values = source && source.getLastRow() > 1 ? source.getRange(2,1,source.getLastRow()-1,HEADERS.length).getValues() : [];
  const rows = values.map(r => {
    const sender = String(r[1] || '');
    const senderName = sender.replace(/<[^>]*>/g,'').replace(/["']/g,'').trim();
    const extracted = String(r[9] || '').trim();
    const parsed = extraerCampos_(extracted);
    const payer = String(r[6] || parsed.pagador || '').trim();
    const date = r[4] || parsed.fecha || '';
    const subject = String(r[2] || '').trim();
    const operation = String(r[5] || parsed.operacion || '').trim();
    const attachment = String(r[7] || '').trim();
    const status = String(r[10] || 'Pendiente de revisión').trim();
    const details = [
      subject ? 'Asunto: ' + subject : '',
      attachment ? 'Adjunto: ' + attachment : '',
      operation ? 'Operación: ' + operation : '',
      extracted ? extracted.slice(0,900) : 'OCR sin texto: revisar imagen del comprobante'
    ].filter(Boolean).join('\n');
    return ['', '', payer || '', r[3] || parsed.importe || '', date, sender ? sender : '', operation, '', '', details + (status ? '\nEstado: ' + status : '')];
  });
  if (rows.length) out.getRange(5,1,rows.length,10).setValues(rows);
  const totalRows = Math.max(24, rows.length + 5);
  if (totalRows > out.getMaxRows()) out.insertRowsAfter(out.getMaxRows(), totalRows-out.getMaxRows());
  out.getRange(1,1,3,10).setFontFamily('Arial').setVerticalAlignment('middle');
  out.getRange('A1').setFontSize(10).setFontWeight('bold');
  out.getRange('A2').setFontSize(9);
  out.getRange('A3').setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  out.setRowHeight(1,24); out.setRowHeight(2,20); out.setRowHeight(3,27);
  out.getRange(4,1,1,10).setFontWeight('bold').setFontSize(8).setBackground('#e7e7e7').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  out.setRowHeight(4,48);
  if (rows.length) {
    out.getRange(5,1,rows.length,10).setFontSize(8).setVerticalAlignment('middle').setWrap(true);
    out.setRowHeights(5,rows.length,52);
  }
  [42,48,150,80,75,145,75,90,85,210].forEach((w,i)=>out.setColumnWidth(i+1,w));
  out.getRange(4,1,totalRows-3,10).setBorder(true,true,true,true,true,true,'#444444',SpreadsheetApp.BorderStyle.SOLID);
  out.getRange('A1:J3').setBorder(false,false,false,false,false,false);
  out.getRange('A4:J4').setBorder(true,true,true,true,true,true,'#222222',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  out.getRange(totalRows,1,1,10).merge().setValue('OBSERVACIONES: __________________________________________________________________________________________________________');
  out.setRowHeight(totalRows,25);
  out.getRange(totalRows,1).setFontSize(8).setVerticalAlignment('middle');
  out.setFrozenRows(4);
  out.setTabColor('#2e6b53');
  if (out.getMaxColumns()>10) out.deleteColumns(11,out.getMaxColumns()-10);
  if (rows.length) {
    out.getRange(5,4,rows.length,1).setNumberFormat('$ #,##0.00');
    out.getRange(5,5,rows.length,1).setNumberFormat('dd/mm/yyyy');
    out.getRange(5,9,rows.length,1).setNumberFormat('$ #,##0.00');
  }
  return {rows:rows.length,sheet:name,sheetId:out.getSheetId()};
}

function obtenerEstadoCopello() {
  const props = PropertiesService.getUserProperties();
  const id = props.getProperty('COPPELLO_SPREADSHEET_ID');
  const labels = [CONFIG.etiquetaPagos,CONFIG.etiquetaLegacy];
  const triggers = ScriptApp.getProjectTriggers();
  const state = {ready:false,labelExists:labels.some(n=>Boolean(GmailApp.getUserLabelByName(n))),label:labels.join(' / '),spreadsheetUrl:'',excelUrl:'',total:0,recent:[],triggerEnabled:triggers.some(t=>t.getHandlerFunction()==='importarCorreosCopello'),lastRun:props.getProperty('COPPELLO_LAST_RUN')||'',lastImport:{}};
  try { state.lastImport=JSON.parse(props.getProperty('COPPELLO_LAST_IMPORT')||'{}'); } catch(e) {}
  if (id) {
    try {
      const ss=SpreadsheetApp.openById(id), sh=ss.getSheetByName(CONFIG.nombreHoja);
      state.ready=Boolean(sh); state.spreadsheetUrl=ss.getUrl(); const printSh=ss.getSheetByName('Planilla para imprimir'); state.printUrl=printSh ? ss.getUrl()+'#gid='+printSh.getSheetId() : ss.getUrl(); state.excelUrl='https://docs.google.com/spreadsheets/d/'+id+'/export?format=xlsx';
      if (sh && sh.getLastRow()>1) {
        state.total=sh.getLastRow()-1;
        const from=Math.max(2,sh.getLastRow()-4);
        state.recent=sh.getRange(from,1,sh.getLastRow()-from+1,HEADERS.length).getDisplayValues().reverse().map(r=>({date:r[0],sender:r[1],subject:r[2],link:r[11],status:r[10],amount:r[3],attachment:r[7]}));
      }
    } catch(e) {}
  }
  return state;
}

/** Reinicia el historial de correos importados, sin borrar los correos de Outlook/Gmail ni los adjuntos de Drive. */
function limpiarHistorialCopello() {
  const props = PropertiesService.getUserProperties();
  const id = props.getProperty('COPPELLO_SPREADSHEET_ID');
  if (!id) throw new Error('Primero prepará el sistema.');
  const ss = SpreadsheetApp.openById(id);
  [CONFIG.nombreHoja, CONFIG.nombreHojaCorreos].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  });
  const printSh = ss.getSheetByName('Planilla para imprimir');
  if (printSh) {
    // La próxima importación reconstruye las filas de la planilla imprimible.
    printSh.clear();
  }
  props.deleteProperty('COPPELLO_LAST_RUN');
  props.deleteProperty('COPPELLO_LAST_IMPORT');
  return obtenerEstadoCopello();
}

function prepararDesdePanelCopello() { prepararCopello(); return obtenerEstadoCopello(); }
function importarDesdePanelCopello() { const result=importarCorreosCopello(); return {result,state:obtenerEstadoCopello()}; }
function instalarRevisionAutomaticaCopello() {
  if (!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='importarCorreosCopello')) ScriptApp.newTrigger('importarCorreosCopello').timeBased().everyMinutes(15).create();
  return obtenerEstadoCopello();
}
function quitarRevisionAutomaticaCopello() {
  ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()==='importarCorreosCopello') ScriptApp.deleteTrigger(t);});
  return obtenerEstadoCopello();
}
