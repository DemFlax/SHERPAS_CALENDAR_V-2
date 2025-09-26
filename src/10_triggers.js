/** 10_triggers.gs — Triggers */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.Triggers = (function(){
  'use strict';

  function onEditMaster(e){
    try{
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
        Sherpas.GuideBook.writeCell(reg.fileId, sh.getName(), fechaISO, isM?'MAÑANA':'TARDE', isM?'MAÑANA':'TARDE', true);
        Sherpas.CalendarSvc.remove(fechaISO, slot, reg.email);
        Sherpas.MailSvc.send(reg.email, 'Liberación: '+Sherpas.Util.dateES(fechaISO)+' '+code+' '+(isM?'M':'T'),
          '<p>Se liberó tu turno del '+Sherpas.Util.dateES(fechaISO)+' ('+(isM?'Mañana':'Tarde')+').</p><p><a href="'+reg.url+'">Abrir calendario</a></p>');
        return;
      }

      if(accion.indexOf('ASIGNAR')===0){
        var asignado = accion.replace('ASIGNAR','ASIGNADO').trim();
        celda.setValue(asignado);
        Sherpas.GuideBook.writeCell(reg.fileId, sh.getName(), fechaISO, isM?'MAÑANA':'TARDE', asignado, true);
        var slot2 = isM? 'M' : asignado.split(' ').pop();
        Sherpas.CalendarSvc.invite(fechaISO, slot2, reg.email);
        Sherpas.MailSvc.send(reg.email, 'Asignación: '+Sherpas.Util.dateES(fechaISO)+' '+code+' '+slot2,
          '<p>Tienes una asignación el '+Sherpas.Util.dateES(fechaISO)+' ('+slot2+').</p><p><a href="'+reg.url+'">Abrir calendario</a></p>');
        return;
      }
    }catch(err){ console.error('onEdit MASTER:', err); }
  }

  function onEditGuide(e){
    try{
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

      if(newVal==='NO DISPONIBLE' || newVal==='REVERTIR' || newVal===''){
        var slotRef = (turno==='MAÑANA')? 'M' : 'T1';
        var start = Sherpas.CalendarSvc._shiftStart(Sherpas.Util.toISO(fecha), slotRef);
        var diffH = (start.getTime() - (new Date()).getTime())/3600000;
        if(diffH < Sherpas.CFG.LIMIT_HOURS){ range.setValue(oldVal); ss.toast('Fuera de ventana (14h).'); return; }
      }

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

  function cronReconcile(){ Sherpas.UseCases.CronReconcileUC(); }

  return {
    onEditMaster: onEditMaster,
    onEditGuide: onEditGuide,
    cronReconcile: cronReconcile
  };
})();

/**
 * HANDLER GLOBAL DE CAMBIOS - Protección completa contra alteraciones
 */
function globalOnChangeHandler(e) {
  try {
    if (!e || !e.changeType) return;
    
    console.log('Cambio detectado:', e.changeType, 'en', e.source.getName());
    
    var sourceSheet = e.source.getActiveSheet();
    var sourceName = sourceSheet.getName();
    
    // Verificar si es una pestaña de mes (MM_YYYY) 
    if (!Sherpas.CFG.MASTER_MONTH_NAME.test(sourceName)) return;
    
    var masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
    var isInMaster = (e.source.getId() === masterId);
    
    if (isInMaster) {
      console.log('Cambio en MASTER permitido:', sourceName);
      return;
    }
    
    // Cambio en calendario GUÍA - Aplicar todas las protecciones
    handleGuideProtections(e, sourceSheet);
    
  } catch(error) {
    console.error('Error en globalOnChangeHandler:', error);
  }
}

/**
 * Protecciones completas para calendarios guía
 */
function handleGuideProtections(e, sheet) {
  try {
    var changeType = e.changeType;
    
    // 1. PROTECCIÓN ESTRUCTURAL
    if (changeType === 'INSERT_ROW') {
      blockUnauthorizedRows(sheet);
    }
    
    if (changeType === 'INSERT_COLUMN') {
      blockUnauthorizedColumns(sheet);
    }
    
    // 2. PROTECCIÓN DE CONTENIDO (Más crítica)
    if (changeType === 'EDIT' || changeType === 'OTHER') {
      validateAndCorrectContent(sheet);
    }
    
    // 3. PROTECCIÓN DE VALIDACIONES DE DATOS
    if (changeType === 'FORMAT' || changeType === 'OTHER') {
      restoreDataValidations(sheet);
    }
    
  } catch(error) {
    console.error('Error en handleGuideProtections:', error);
  }
}

/**
 * Bloquea filas no autorizadas
 */
function blockUnauthorizedRows(sheet) {
  var parsed = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
  if (!parsed) return;
  
  var meta = Sherpas.Util.monthMeta(parsed.yyyy, parsed.mm);
  var correctRows = 1 + meta.weeks * 3;
  var maxAllowed = correctRows + 3;
  
  var currentRows = sheet.getMaxRows();
  
  if (currentRows > maxAllowed) {
    var excessRows = currentRows - maxAllowed;
    sheet.deleteRows(maxAllowed + 1, excessRows);
    
    showWarning('Filas Bloqueadas', 'No se pueden añadir filas al calendario.');
    console.log('Eliminadas ' + excessRows + ' filas de ' + sheet.getName());
  }
}

/**
 * Bloquea columnas no autorizadas
 */
function blockUnauthorizedColumns(sheet) {
  var currentCols = sheet.getMaxColumns();
  var correctCols = 7; // Lun-Dom
  
  if (currentCols > correctCols + 2) { // Pequeño margen
    var excessCols = currentCols - correctCols;
    sheet.deleteColumns(correctCols + 1, excessCols);
    
    showWarning('Columnas Bloqueadas', 'No se pueden añadir columnas al calendario.');
    console.log('Eliminadas ' + excessCols + ' columnas de ' + sheet.getName());
  }
}

/**
 * CRÍTICO: Valida y corrige contenido de celdas
 */
function validateAndCorrectContent(sheet) {
  try {
    var parsed = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
    if (!parsed) return;
    
    var meta = Sherpas.Util.monthMeta(parsed.yyyy, parsed.mm);
    var mtA1List = Sherpas.Util.monthMT_A1_FromMeta(meta, 2);
    
    var correctionsMade = 0;
    var invalidValues = [];
    
    // Verificar cada celda editable M/T
    mtA1List.forEach(function(a1) {
      var range = sheet.getRange(a1);
      var value = String(range.getDisplayValue() || '').toUpperCase().trim();
      
      // Verificar si el valor es válido
      var isValid = Sherpas.CFG.GUIDE_DV.includes(value) || 
                   value.startsWith('ASIGNADO');
      
      if (!isValid && value !== '') {
        // Valor inválido detectado
        invalidValues.push({a1: a1, value: value});
        
        // Determinar valor correcto basado en la fila
        var pos = Sherpas.Util.a1ToRowCol(a1);
        var rowType = (pos.row - 2) % 3; // 1=MAÑANA, 2=TARDE
        var correctValue = (rowType === 1) ? 'MAÑANA' : 'TARDE';
        
        // Corregir inmediatamente
        range.setValue(correctValue);
        correctionsMade++;
      }
    });
    
    // Mostrar advertencia si se hicieron correcciones
    if (correctionsMade > 0) {
      var warning = 'Se corrigieron ' + correctionsMade + ' valores no válidos:\n';
      invalidValues.slice(0, 3).forEach(function(inv) {
        warning += '• ' + inv.a1 + ': "' + inv.value + '" → valor correcto\n';
      });
      if (invalidValues.length > 3) {
        warning += '• ... y ' + (invalidValues.length - 3) + ' más';
      }
      
      showWarning('Contenido Corregido', warning);
      console.log('Valores corregidos en ' + sheet.getName() + ':', invalidValues);
    }
    
  } catch(error) {
    console.error('Error validando contenido:', error);
  }
}

/**
 * Restaura validaciones de datos si fueron alteradas
 */
function restoreDataValidations(sheet) {
  try {
    var parsed = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
    if (!parsed) return;
    
    var meta = Sherpas.Util.monthMeta(parsed.yyyy, parsed.mm);
    var mtA1List = Sherpas.Util.monthMT_A1_FromMeta(meta, 2);
    
    // Crear regla de validación correcta
    var rule = Sherpas.DVManager.buildListRule(Sherpas.CFG.GUIDE_DV);
    
    var restoredCount = 0;
    
    // Verificar cada celda M/T
    mtA1List.forEach(function(a1) {
      var range = sheet.getRange(a1);
      var currentRule = range.getDataValidation();
      
      // Si no hay regla o la regla es incorrecta, restaurar
      if (!currentRule) {
        range.setDataValidation(rule);
        restoredCount++;
      }
    });
    
    if (restoredCount > 0) {
      showWarning('Desplegables Restaurados', 
        'Se restauraron ' + restoredCount + ' desplegables que habían sido alterados.');
      console.log('Restauradas ' + restoredCount + ' validaciones en ' + sheet.getName());
    }
    
  } catch(error) {
    console.error('Error restaurando validaciones:', error);
  }
}

/**
 * Función auxiliar para mostrar advertencias
 */
function showWarning(title, message) {
  try {
    // Verificar si hay UI disponible
    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getUi) {
      SpreadsheetApp.getUi().alert(
        title, 
        message + '\n\nEsta acción fue automáticamente revertida.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch(e) {
    // Si no hay UI (ejecución automática), solo log
    console.log('Advertencia:', title, '-', message);
  }
}

/**
 * FUNCIÓN ADICIONAL: Reparación completa de un calendario específico
 */
function repairGuideCalendarComplete(guideCode) {
  try {
    var reg = Sherpas.RegistryRepo.resolve(guideCode);
    if (!reg) {
      console.error('Guía no encontrado:', guideCode);
      return false;
    }
    
    var ss = SpreadsheetApp.openById(reg.fileId);
    var repaired = 0;
    
    ss.getSheets().forEach(function(sheet) {
      if (Sherpas.CFG.GUIDE_MONTH_NAME.test(sheet.getName())) {
        // Reparar estructura
        Sherpas.GuideBook.normalize(sheet);
        
        // Reparar contenido
        validateAndCorrectContent(sheet);
        
        // Reparar validaciones
        restoreDataValidations(sheet);
        
        // Reparar formato condicional
        Sherpas.GuideBook.applyCF(sheet);
        
        // Reparar protecciones
        Sherpas.GuideBook.protectEditableMT(sheet);
        
        repaired++;
      }
    });
    
    console.log('Calendario ' + guideCode + ' reparado completamente: ' + repaired + ' pestañas');
    return true;
    
  } catch(error) {
    console.error('Error reparando calendario:', guideCode, error);
    return false;
  }
}

/**
 * Función de debug para verificar triggers
 */
function debugTriggers() {
  try {
    if (typeof Sherpas !== 'undefined' && Sherpas.TriggerSvc && Sherpas.TriggerSvc.countActiveTriggers) {
      return Sherpas.TriggerSvc.countActiveTriggers();
    } else {
      var triggers = ScriptApp.getProjectTriggers();
      console.log('Total triggers activos:', triggers.length);
      return { total: triggers.length };
    }
  } catch(error) {
    console.error('Error en debugTriggers:', error);
    return { error: error.message };
  }
}

/* Entradas globales */
function onOpen(){ Sherpas.UI.onOpen(); }                 // menú
function onEdit(e){ Sherpas.Triggers.onEditMaster(e); }   // simple trigger en MASTER