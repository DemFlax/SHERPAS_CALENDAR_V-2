/** 14_integration_updates.js — Actualizaciones de Triggers e Integración
 *  Mejoras en los triggers existentes para integrar nuevos módulos
 */
/* global Sherpas */

/** Actualizaciones a 10_triggers.js */

/**
 * REEMPLAZAR función onEditMaster en 10_triggers.js
 * Integra ValidationEngine y SyncController
 */
function onEditMaster_UPDATED(e){
  try{
    // 1. Validación con rollback automático
    const validationResult = Sherpas.ValidationEngine.validateAndRollback(e);
    if(validationResult.state === Sherpas.ValidationEngine.VALIDATION_STATES.INVALID_REVERTED) {
      console.log('Cambio inválido revertido:', validationResult.message);
      return; // Salir si el cambio fue inválido
    }

    var sh = e.range.getSheet(); if(!Sherpas.CFG.MASTER_MONTH_NAME.test(sh.getName())) return;
    var row = e.range.getRow(), col = e.range.getColumn(); if(row<3) return;

    var blocks = (function _listAllGuideBlocks(sheet){
      var res=[], lastCol=sheet.getLastColumn();
      for(var c=1;c<=lastCol;c++){
        var m = String(sheet.getRange(2,c).getDisplayValue()||'').toUpperCase();
        var t = String(sheet.getRange(2,c+1).getDisplayValue()||'').toUpperCase();
        if((m==='MAÑANA'||m==='M') && (t==='TARDE'||t==='T')){ res.push({colM:c,colT:c+1}); c++; }
      }
      return res;
    })(sh);
    var blk = blocks.find(function(b){ return col===b.colM || col===b.colT; }); if(!blk) return;

    var fechaISO = Sherpas.Util.toISO(Sherpas.Util.fromDisplay(sh.getRange(row,1).getDisplayValue()));
    var isM = (col===blk.colM);
    var accion = String(e.value||'').toUpperCase();

    var validList = isM? Sherpas.CFG.MASTER_DV_M : Sherpas.CFG.MASTER_DV_T;
    if(validList.indexOf(accion)===-1){ e.range.setValue(e.oldValue||''); sh.toast('Valor no válido.'); return; }

    var header = sh.getRange(1, isM?blk.colM:blk.colT).getDisplayValue();
    var code = (header.split('—')[0]||header).trim().toUpperCase();
    var reg = Sherpas.RegistryRepo.list().find(function(r){ return r.codigo===code; });
    if(!reg){ sh.toast('Guía no en REGISTRO'); return; }

    var celda = sh.getRange(row, isM?blk.colM:blk.colT);
    var current = String(celda.getDisplayValue()||'').toUpperCase();
    if(current==='NO DISPONIBLE' && accion!=='LIBERAR'){ e.range.setValue(e.oldValue||''); sh.toast('Bloqueado por GUÍA.'); return; }

    if(accion==='LIBERAR'){
      var slot = isM? 'M' : (current.indexOf('ASIGNADO')===0? current.split(/\s+/).pop() : 'T1');
      celda.setValue('');
      Sherpas.GuideBook.writeCell(reg.fileId, sh.getName(), fechaISO, isM?'MAÑANA':'TARDE', isM?'MAÑANA':'TARDE', false);
      Sherpas.CalendarSvc.remove(fechaISO, slot, reg.email);
      // Usar nuevo template profesional
      Sherpas.EmailTemplates.sendRelease(reg.nombre, reg.codigo, reg.email, fechaISO, isM?'MAÑANA':'TARDE', reg.url);
      return;
    }

    if(accion.indexOf('ASIGNAR')===0){
      var asignado = accion.replace('ASIGNAR','ASIGNADO').trim();
      celda.setValue(asignado);
      Sherpas.GuideBook.writeCell(reg.fileId, sh.getName(), fechaISO, isM?'MAÑANA':'TARDE', asignado, true);
      var slot2 = isM? 'M' : asignado.split(' ').pop();
      Sherpas.CalendarSvc.invite(fechaISO, slot2, reg.email);
      // Usar nuevo template profesional
      Sherpas.EmailTemplates.sendAssignment(reg.nombre, reg.codigo, reg.email, fechaISO, slot2, reg.url);
      return;
    }
  }catch(err){ console.error('onEdit MASTER:', err); }
}

/**
 * REEMPLAZAR función onEditGuide en 10_triggers.js
 * Integra ValidationEngine
 */
function onEditGuide_UPDATED(e){
  try{
    // 1. Validación con rollback automático
    const validationResult = Sherpas.ValidationEngine.validateAndRollback(e);
    if(validationResult.state === Sherpas.ValidationEngine.VALIDATION_STATES.INVALID_REVERTED) {
      console.log('Cambio inválido revertido:', validationResult.message);
      return; // Salir si el cambio fue inválido
    }

    var range = e && e.range; if(!range) return;
    var sh = range.getSheet(); var ss = sh.getParent();
    if(!Sherpas.CFG.GUIDE_MONTH_NAME.test(sh.getName())) return;

    // Si hay filas extra por debajo, recorta y salir
    var p = Sherpas.Util.parseTab_MMYYYY(sh.getName()); if(!p) return;
    var rowsNeed = (1 + Sherpas.Util.monthMeta(p.yyyy,p.mm).weeks*3) + 1; // +1 tampón
    if(sh.getMaxRows() > rowsNeed){
      Sherpas.GuideBook.normalize(sh);
      ss.toast('Se normalizó el mes (filas extra eliminadas).');
      return;
    }

    var r = range.getRow(), c = range.getColumn();
    var off = (r - 2) % 3; if(off!==1 && off!==2) return; // 1=MAÑANA 2=TARDE
    var turno = (off===1)? 'MAÑANA':'TARDE';

    var newVal = (e.value==null? '' : String(e.value).trim().toUpperCase());
    var oldVal = (e.oldValue==null? sh.getRange(r,c).getDisplayValue() : String(e.oldValue).trim().toUpperCase());

    var num = parseInt(sh.getRange(r-off, c).getDisplayValue(),10); if(!num) return;
    var yyyy=p.yyyy, mm=p.mm;
    var fecha = new Date(yyyy, mm-1, num);

    var master = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID));
    var ms = master.getSheetByName(sh.getName()); if(!ms){ ss.toast('MASTER sin pestaña '+sh.getName()); return; }
    var header = ss.getName();
    var mTitle = String(header||'').match(/-(G\d{2})$/i);
    var codigo = mTitle? mTitle[1].toUpperCase() : 'G00';

    var cols = (function _findGuideBlockCols(sheet, codigo){
      var lastCol = sheet.getLastColumn();
      var top = sheet.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(s){ return String(s||'').toUpperCase(); });
      var row2= sheet.getRange(2,1,1,lastCol).getDisplayValues()[0].map(function(s){ return String(s||'').toUpperCase(); });
      var codeUp = codigo.toUpperCase();
      for(var c2=1;c2<=lastCol;c2++){
        var t = top[c2-1];
        if(t.startsWith(codeUp+' ') || t.startsWith(codeUp+'—') || t===codeUp){
          var isM=(row2[c2-1]==='MAÑANA'||row2[c2-1]==='M'), isT=(row2[c2]==='TARDE'||row2[c2]==='T');
          if(isM && isT) return {colM:c2, colT:c2+1};
        }
      }
      return {colM:0,colT:0};
    })(ms, codigo);

    var rowMaster = Sherpas.MasterBook.findRowByISO(ms, Sherpas.Util.toISO(fecha));
    if(!cols.colM || !cols.colT || rowMaster<3){ ss.toast('MASTER preparándose...'); return; }

    var targetCol = (turno==='MAÑANA')? cols.colM : cols.colT;
    var cell = ms.getRange(rowMaster, targetCol);
    var currentMaster = String(cell.getDisplayValue()||'').toUpperCase();

    if(currentMaster.indexOf('ASIGNADO')===0){ range.setValue(turno); ss.toast('Turno asignado por MASTER.'); return; }
    if(newVal==='NO DISPONIBLE'){ cell.setValue('NO DISPONIBLE'); return; }
    if(newVal==='REVERTIR' || newVal===''){ range.setValue(turno); cell.setValue(''); return; }
  }catch(err){ console.error('onEdit GUIDE:', err); }
}

/**
 * ACTUALIZAR casos de uso para usar templates profesionales
 * REEMPLAZAR función CreateGuideUC en 08_casos_uso.js
 */
function CreateGuideUC_UPDATED(nombre, email, codigo){
  if(!nombre||!email||!codigo) throw new Error('Datos incompletos');
  if(!Sherpas.RegistryRepo.checkUniq(codigo, email)) throw new Error('Código o email ya existen');
  var folderId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.DEST_FOLDER_ID) || Sherpas.CFG.DEST_FOLDER_ID_DEFAULT;
  var file = SpreadsheetApp.create('CALENDARIO_'+nombre+'-'+codigo);
  DriveApp.getFileById(file.getId()).moveTo(DriveApp.getFolderById(folderId));
  var now = new Date(), y = now.getFullYear();
  var startM = Math.max(10, now.getMonth()+1);
  for(var m=startM; m<=12; m++){ Sherpas.GuideBook.buildMonth(file.getId(), new Date(y, m-1, 1)); }
  Sherpas.GuideBook.adoptGuide(file.getId());
  Sherpas.RegistryRepo.upsert({codigo:codigo, nombre:nombre, email:String(email).toLowerCase(), fileId:file.getId(), url:file.getUrl()});
  Sherpas.MasterBook.ensureMonthsFromOct();
  Sherpas.MasterBook.ensureGuideColumns(codigo, nombre);
  if(email) file.addEditor(email);
  
  // Enviar email de bienvenida con template profesional
  Sherpas.EmailTemplates.sendWelcome(nombre, codigo, email, file.getUrl());
  
  return file;
}

/**
 * ACTUALIZAR función InitUC para configurar sincronización automática
 * REEMPLAZAR en 08_casos_uso.js
 */
function InitUC_UPDATED(){
  var ss = SpreadsheetApp.getActive();
  var SP = PropertiesService.getScriptProperties();
  
  if(!SP.getProperty(Sherpas.KEYS.MASTER_ID)) SP.setProperty(Sherpas.KEYS.MASTER_ID, ss.getId());
  if(!SP.getProperty(Sherpas.KEYS.DEST_FOLDER_ID)) SP.setProperty(Sherpas.KEYS.DEST_FOLDER_ID, Sherpas.CFG.DEST_FOLDER_ID_DEFAULT);
  
  // Configurar triggers existentes
  Sherpas.TriggerSvc.ensureTimeEvery5m('Sherpas.Triggers.cronReconcile');
  
  // Configurar nuevo trigger de sincronización automática
  Sherpas.SyncController.setupAutoSync();
  
  Sherpas.RegistryRepo.ensure();
  Sherpas.MasterBook.ensureMonthsFromOct();
  Sherpas.MasterBook.applyDVandCF_All();
  Sherpas.RegistryRepo.list().forEach(function(r){ Sherpas.MasterBook.ensureGuideColumns(r.codigo, r.nombre); });
  
  ss.toast('Sistema inicializado con sincronización automática activada');
}

/**
 * AÑADIR nuevas funciones de menú 
 * ACTUALIZAR 09_menu.js añadiendo estas funciones
 */
function TestSync(){
  var result = Sherpas.SyncController.executeBidirectionalSync();
  var message = result.success ? 
    `Sincronización exitosa: ${result.changes.length} cambios, ${result.conflicts.length} conflictos` :
    `Error en sincronización: ${result.errors.length} errores`;
  SpreadsheetApp.getActive().toast(message);
}

function CleanAuditLogs(){
  Sherpas.ValidationEngine.cleanOldAuditLogs();
  SpreadsheetApp.getActive().toast('Logs de auditoría limpiados');
}

function TestEmailTemplate(){
  var ui = SpreadsheetApp.getUi();
  var email = ui.prompt('Email de prueba').getResponseText().trim();
  if(!email) return;
  
  try{
    Sherpas.EmailTemplates.sendWelcome('Juan Pérez', 'G01', email, 'https://example.com/calendario');
    ui.alert('Email de prueba enviado correctamente');
  }catch(e){ 
    ui.alert('Error enviando email: ' + String(e)); 
  }
}

function ViewAuditLog(){
  try{
    var masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
    var masterSS = SpreadsheetApp.openById(masterId);
    var logSheet = masterSS.getSheetByName('AUDIT_LOG');
    
    if(!logSheet){
      SpreadsheetApp.getActive().toast('No existe hoja de auditoría aún');
      return;
    }
    
    masterSS.setActiveSheet(logSheet);
    SpreadsheetApp.getActive().toast('Abriendo log de auditoría...');
  }catch(e){
    SpreadsheetApp.getUi().alert('Error accediendo al log: ' + String(e));
  }
}

/**
 * ACTUALIZAR menú en 09_menu.js
 * REEMPLAZAR función onOpen
 */
function onOpen_UPDATED(){
  SpreadsheetApp.getUi().createMenu('Sherpas')
    .addItem('Inicializar (MASTER)','Sherpas.UI.Inicializar')
    .addSeparator()
    .addItem('Crear Guía (nuevo)','Sherpas.UI.CrearGuia')
    .addItem('Adoptar Guía (URL/ID/Código/Email)','Sherpas.UI.AdoptarGuia')
    .addSeparator()
    .addItem('Sincronizar ahora','Sherpas.UI.SyncAhora')
    .addItem('🔄 Prueba Sincronización Automática','Sherpas.UI.TestSync')
    .addItem('Reparar guías ahora','Sherpas.UI.RepararGuiasAhora')
    .addSeparator()
    .addItem('📧 Test Email Template','Sherpas.UI.TestEmailTemplate')
    .addItem('📋 Ver Log de Auditoría','Sherpas.UI.ViewAuditLog')
    .addItem('🧹 Limpiar Logs Antiguos','Sherpas.UI.CleanAuditLogs')
    .addSeparator()
    .addItem('Reconstruir REGISTRO desde carpeta','Sherpas.UI.RebuildRegistro')
    .addItem('Eliminar Guía (TOTAL)','Sherpas.UI.EliminarGuiaTotal')
    .addToUi();
}

/**
 * ACTUALIZAR configuración para incluir emails admin
 * AÑADIR a 01_config.js en Sherpas.CFG
 */
const CFG_ADMIN_EMAILS_ADDITION = {
  // Añadir estos emails al objeto Sherpas.CFG existente
  ADMIN_EMAILS: ['manager@spainfoodsherpas.com'], // Configurar emails reales
  
  // Configuración de validación
  VALIDATION: {
    MAX_FAILED_ATTEMPTS: 3,
    BLOCK_DURATION_MINUTES: 15,
    LOG_RETENTION_DAYS: 30,
    NOTIFICATION_THRESHOLD: 2
  }
};

/**
 * FUNCIONES HELPER PARA INTEGRACIÓN
 * Añadir estas funciones globales al final de cualquier archivo .js
 */

/**
 * Función global para ser llamada por trigger de sincronización
 */
function triggerBidirectionalSync(){
  try {
    Sherpas.SyncController.triggerAutoSync();
  } catch(error) {
    console.error('Error en trigger bidireccional:', error);
  }
}

/**
 * Función global para limpieza automática de logs
 */
function triggerLogCleanup(){
  try {
    Sherpas.ValidationEngine.cleanOldAuditLogs();
  } catch(error) {
    console.error('Error en limpieza de logs:', error);
  }
}

/**
 * CONFIGURACIÓN DE TRIGGERS ADICIONALES
 * Añadir estas funciones a TriggerSvc en 04_integraciones.js
 */
function ensureDailyLogCleanup(){
  var existing = ScriptApp.getProjectTriggers().some(function(t){
    return t.getHandlerFunction()==='triggerLogCleanup' && t.getTriggerSource()===ScriptApp.TriggerSource.CLOCK;
  });
  
  if(!existing) {
    ScriptApp.newTrigger('triggerLogCleanup')
      .timeBased()
      .everyDays(1)
      .atHour(2) // 2:00 AM
      .create();
    console.log('Trigger de limpieza diaria configurado');
  }
}

/**
 * MEJORAS EN CalendarSvc
 * ACTUALIZAR funciones en 04_integraciones.js para mejor manejo de errores
 */
function invite_IMPROVED(iso, slot, email){
  try{
    if(!email) {
      console.warn('Email vacío en CalendarSvc.invite');
      return false;
    }
    
    var cal = CalendarApp.getCalendarById(Sherpas.CFG.CALENDAR_ID);
    if(!cal) {
      console.error('Calendar no encontrado:', Sherpas.CFG.CALENDAR_ID);
      return false;
    }
    
    var start = _shiftStart(iso, slot);
    var ev = _getEvent(cal, start, slot);
    if(!ev) {
      console.warn('Evento no encontrado para:', iso, slot);
      return false;
    }
    
    ev.addGuest(email);
    console.log('Invitación enviada:', email, iso, slot);
    return true;
  }catch(e){ 
    console.error('CalendarSvc.invite error:', e);
    return false;
  }
}

function remove_IMPROVED(iso, slot, email){
  try{
    if(!email) {
      console.warn('Email vacío en CalendarSvc.remove');
      return false;
    }
    
    var cal = CalendarApp.getCalendarById(Sherpas.CFG.CALENDAR_ID);
    if(!cal) {
      console.error('Calendar no encontrado:', Sherpas.CFG.CALENDAR_ID);
      return false;
    }
    
    var start = _shiftStart(iso, slot);
    var ev = _getEvent(cal, start, slot);
    if(!ev) {
      console.warn('Evento no encontrado para remover:', iso, slot);
      return false;
    }
    
    try{ 
      ev.removeGuest(email); 
      console.log('Invitación removida:', email, iso, slot);
      return true;
    }catch(removeError){
      console.warn('Error removiendo guest (puede no estar invitado):', removeError);
      return false;
    }
  }catch(e){ 
    console.error('CalendarSvc.remove error:', e);
    return false;
  }
}

/**
 * FUNCIÓN DE MONITOREO DE SISTEMA
 * Añadir esta función para diagnósticos
 */
function systemHealthCheck(){
  const health = {
    timestamp: new Date(),
    masterSheet: false,
    guidesCount: 0,
    triggersActive: [],
    lastSync: null,
    auditLogRows: 0,
    errors: []
  };
  
  try {
    // Verificar Master Sheet
    const masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
    if(masterId) {
      SpreadsheetApp.openById(masterId);
      health.masterSheet = true;
    }
    
    // Contar guías
    health.guidesCount = Sherpas.RegistryRepo.list().length;
    
    // Verificar triggers
    const triggers = ScriptApp.getProjectTriggers();
    health.triggersActive = triggers.map(t => t.getHandlerFunction());
    
    // Última sincronización
    const lastSyncProp = PropertiesService.getScriptProperties().getProperty('LAST_SYNC_TIMESTAMP');
    if(lastSyncProp) {
      health.lastSync = new Date(lastSyncProp);
    }
    
    // Audit log
    const masterSS = SpreadsheetApp.openById(masterId);
    const logSheet = masterSS.getSheetByName('AUDIT_LOG');
    if(logSheet) {
      health.auditLogRows = logSheet.getLastRow() - 1; // -1 para header
    }
    
  } catch(error) {
    health.errors.push(error.message);
  }
  
  console.log('=== SYSTEM HEALTH CHECK ===');
  console.log(JSON.stringify(health, null, 2));
  
  return health;
}

/**
 * NOTAS DE IMPLEMENTACIÓN:
 * 
 * 1. Reemplazar las funciones onEditMaster y onEditGuide en 10_triggers.js
 * 2. Actualizar CreateGuideUC e InitUC en 08_casos_uso.js  
 * 3. Añadir nuevas funciones de menú en 09_menu.js
 * 4. Configurar emails de admin en 01_config.js
 * 5. Mejorar CalendarSvc en 04_integraciones.js
 * 6. Añadir funciones globales para triggers
 * 
 * TESTING:
 * - Probar sincronización manual desde menú
 * - Verificar templates de email con función de test
 * - Monitorear audit log después de cambios inválidos
 * - Confirmar que triggers automáticos se configuran correctamente
 */