const CONFIG = { etiquetaGmail: 'Copello_Para_Excel', nombrePlanilla: 'Copello_Registro_Correos', nombreHoja: 'Correos', limiteConversaciones: 100 };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Copello | Centro de correos').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function prepararCopello() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('COPPELLO_SPREADSHEET_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.nombrePlanilla);
    props.setProperty('COPPELLO_SPREADSHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(CONFIG.nombreHoja);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(CONFIG.nombreHoja); }
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Fecha del correo','Remitente','Asunto','Cuerpo del correo','Enlace al correo','ID del mensaje','Fecha de importación']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#18352e').setFontColor('#ffffff');
    [155,240,280,420,240,220,155].forEach((w,i) => sh.setColumnWidth(i+1,w));
  }
  if (!GmailApp.getUserLabelByName(CONFIG.etiquetaGmail)) GmailApp.createLabel(CONFIG.etiquetaGmail);
  return {url:ss.getUrl(), label:CONFIG.etiquetaGmail};
}

function importarCorreosCopello() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('COPPELLO_SPREADSHEET_ID');
  if (!id) throw new Error('Primero prepará Copello desde el panel.');
  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName(CONFIG.nombreHoja);
  const label = GmailApp.getUserLabelByName(CONFIG.etiquetaGmail);
  if (!label) throw new Error('No existe la etiqueta Gmail ' + CONFIG.etiquetaGmail);
  const ids = new Set();
  if (sh.getLastRow() > 1) sh.getRange(2,6,sh.getLastRow()-1,1).getDisplayValues().flat().filter(String).forEach(v => ids.add(v));
  const threads = label.getThreads(0,CONFIG.limiteConversaciones);
  const now = new Date();
  const rows = [];
  threads.forEach(thread => thread.getMessages().forEach(msg => {
    const msgId = msg.getId();
    if (ids.has(msgId)) return;
    rows.push([msg.getDate(),msg.getFrom(),msg.getSubject(),msg.getPlainBody(),'https://mail.google.com/mail/u/0/#all/'+msgId,msgId,now]);
    ids.add(msgId);
  }));
  if (rows.length) {
    const start = sh.getLastRow()+1;
    sh.getRange(start,1,rows.length,7).setValues(rows);
    sh.getRange(start,1,rows.length,1).setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange(start,7,rows.length,1).setNumberFormat('dd/mm/yyyy hh:mm');
  }
  props.setProperty('COPPELLO_LAST_RUN',new Date().toISOString());
  return {imported:rows.length, total:Math.max(0,sh.getLastRow()-1), spreadsheetUrl:ss.getUrl()};
}

function obtenerEstadoCopello() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('COPPELLO_SPREADSHEET_ID');
  const triggers = ScriptApp.getProjectTriggers();
  const state = {ready:false,labelExists:Boolean(GmailApp.getUserLabelByName(CONFIG.etiquetaGmail)),label:CONFIG.etiquetaGmail,spreadsheetUrl:'',total:0,recent:[],triggerEnabled:triggers.some(t=>t.getHandlerFunction()==='importarCorreosCopello'),lastRun:props.getProperty('COPPELLO_LAST_RUN')||''};
  if (id) {
    try {
      const ss = SpreadsheetApp.openById(id);
      const sh = ss.getSheetByName(CONFIG.nombreHoja);
      state.ready = true; state.spreadsheetUrl = ss.getUrl();
      if (sh && sh.getLastRow()>1) {
        const last=sh.getLastRow(); state.total=last-1;
        const from=Math.max(2,last-4);
        state.recent=sh.getRange(from,1,last-from+1,7).getDisplayValues().reverse().map(r=>({date:r[0],sender:r[1],subject:r[2],link:r[4]}));
      }
    } catch(e) {}
  }
  return state;
}

function prepararDesdePanelCopello() { prepararCopello(); return obtenerEstadoCopello(); }
function importarDesdePanelCopello() { const result=importarCorreosCopello(); return {result:result,state:obtenerEstadoCopello()}; }
function instalarRevisionAutomaticaCopello() {
  if (!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='importarCorreosCopello')) {
    ScriptApp.newTrigger('importarCorreosCopello').timeBased().everyMinutes(15).create();
  }
  return obtenerEstadoCopello();
}
function quitarRevisionAutomaticaCopello() {
  ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()==='importarCorreosCopello') ScriptApp.deleteTrigger(t);});
  return obtenerEstadoCopello();
}