/** 06_master.gs — Libro MASTER (dominio)
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.MasterBook = (function(){
  'use strict';

  /* ======== acceso ======== */
  function _SP(){ return PropertiesService.getScriptProperties(); }
  function _open(){
    var id = _SP().getProperty(Sherpas.KEYS.MASTER_ID);
    return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
  }
  function _listMonthTabs(ss){
    return ss.getSheets().filter(function(sh){ return Sherpas.CFG.MASTER_MONTH_NAME.test(sh.getName()); });
  }

  /* ======== construcción meses ======== */
  function ensureMonthsFromOct(){
    var ss = _open();
    var now = new Date(); var year = now.getFullYear();
    var startM = Math.max(10, now.getMonth()+1); // desde octubre
    for(var m=startM; m<=12; m++){
      var mm = Sherpas.Util.pad2(m), tab = mm+'_'+year;
      if(!ss.getSheetByName(tab)) _buildMonth(ss, tab);
    }
  }

  function _buildMonth(ss, ym){
    var m = ym.match(Sherpas.CFG.MASTER_MONTH_NAME); if(!m) return null;
    var mm = parseInt(m[1],10), yyyy = parseInt(m[2],10);
    var sh = ss.getSheetByName(ym); if(sh) return sh;
    
    // CORRECCIÓN: Insertar pestañas en posición correcta (orden cronológico)
    var existingSheets = ss.getSheets();
    var insertPosition = _calculateSheetPosition(existingSheets, mm, yyyy);
    
    if(insertPosition >= 0) {
      sh = ss.insertSheet(ym, insertPosition);
    } else {
      sh = ss.insertSheet(ym);
    }
    
    sh.getRange(1,1).setValue('FECHA');
    sh.setFrozenRows(2);

    var first = new Date(yyyy, mm-1, 1);
    var lastDay = new Date(yyyy, mm, 0).getDate();
    var values = Array.from({length:lastDay}, function(_,i){ return [new Date(yyyy, mm-1, i+1)]; });
    
    // CORRECCIÓN: Usar SheetWriter con gestión inteligente de filas
    var rowsNeeded = lastDay + 2; // 2 filas de cabecera + días del mes
    Sherpas.SheetWriter.ensureExactRows(sh, rowsNeeded);
    
    sh.getRange(3,1,lastDay,1).setValues(values);
    sh.getRange(3,1,lastDay,1).setNumberFormat('dd/MM/yyyy');
    return sh;
  }

  /**
   * NUEVA: Calcula posición correcta para insertar pestaña mensual
   */
  function _calculateSheetPosition(existingSheets, targetMonth, targetYear) {
    var monthSheets = [];
    
    // Extraer hojas mensuales existentes
    existingSheets.forEach(function(sheet, index) {
      var match = sheet.getName().match(Sherpas.CFG.MASTER_MONTH_NAME);
      if(match) {
        var mm = parseInt(match[1], 10);
        var yyyy = parseInt(match[2], 10);
        monthSheets.push({
          index: index,
          month: mm,
          year: yyyy,
          name: sheet.getName()
        });
      }
    });
    
    // Ordenar por año y mes
    monthSheets.sort(function(a, b) {
      if(a.year !== b.year) return a.year - b.year;
      return a.month - b.month;
    });
    
    // Encontrar posición de inserción
    for(var i = 0; i < monthSheets.length; i++) {
      var sheet = monthSheets[i];
      if(targetYear < sheet.year || (targetYear === sheet.year && targetMonth < sheet.month)) {
        return sheet.index;
      }
    }
    
    // Si no se encuentra, insertar al final de las hojas mensuales
    if(monthSheets.length > 0) {
      return monthSheets[monthSheets.length - 1].index + 1;
    }
    
    return -1; // Insertar al final
  }

  /* ======== columnas por guía ======== */
  function ensureGuideColumns(codigo, nombre){
    var ss = _open();
    _listMonthTabs(ss).forEach(function(ms){
      var info = _findGuideBlockCols(ms, codigo);
      if(info.colM && info.colT) return;
      
      // CORRECCIÓN: Encontrar verdadera última columna con contenido de guía
      var lastGuideCol = _findLastGuideColumn(ms);
      var insertAfter = Math.max(1, lastGuideCol); // Mínimo columna 1 (FECHA)
      
      ms.insertColumnsAfter(insertAfter, 2);
      var start = insertAfter + 1;
      ms.getRange(2,start,1,2).setValues([['MAÑANA','TARDE']]);
      var top = ms.getRange(1,start,1,2); top.mergeAcross(); top.setValue(codigo+' — '+nombre);
      _applyDVandCF_ForCols(ms, start, start+1);
    });
  }

  /**
   * NUEVA: Encuentra la verdadera última columna con contenido de guía
   */
  function _findLastGuideColumn(sheet) {
    var maxCol = sheet.getMaxColumns();
    var header1 = sheet.getRange(1, 1, 1, maxCol).getDisplayValues()[0];
    var header2 = sheet.getRange(2, 1, 1, maxCol).getDisplayValues()[0];
    
    var lastGuideCol = 1; // Empezar después de FECHA
    
    // Buscar hacia atrás desde la última columna
    for(var col = maxCol; col >= 1; col--) {
      var h1 = String(header1[col-1] || '').trim();
      var h2M = String(header2[col-1] || '').toUpperCase().trim();
      var h2T = String(header2[col] || '').toUpperCase().trim();
      
      // Si encontramos un bloque de guía válido (código en h1, MAÑANA/TARDE en h2)
      if(h1.match(/^G\d{2}\s*—/) && h2M === 'MAÑANA' && h2T === 'TARDE') {
        lastGuideCol = col + 1; // +1 porque TARDE está en la siguiente columna
        break;
      }
    }
    
    return lastGuideCol;
  }

  function _applyDVandCF_ForCols(sheet, colM, colT){
    var lastRow = Math.max(3, sheet.getLastRow());
    var ruleM = Sherpas.DVManager.buildListRule(Sherpas.CFG.MASTER_DV_M);
    var ruleT = Sherpas.DVManager.buildListRule(Sherpas.CFG.MASTER_DV_T);
    sheet.getRange(3,colM,lastRow-2,1).setDataValidation(ruleM);
    sheet.getRange(3,colT,lastRow-2,1).setDataValidation(ruleT);
    Sherpas.CFManager.setMasterRulesForCols(sheet, colM, colT);
  }

  function applyDVandCF_All(){
    var ss = _open();
    _listMonthTabs(ss).forEach(function(ms){
      _listAllGuideBlocks(ms).forEach(function(b){ _applyDVandCF_ForCols(ms, b.colM, b.colT); });
    });
  }

  function _listAllGuideBlocks(sheet){
    var res=[], lastCol=sheet.getLastColumn();
    for(var c=1;c<=lastCol;c++){
      var m = String(sheet.getRange(2,c).getDisplayValue()||'').toUpperCase();
      var t = String(sheet.getRange(2,c+1).getDisplayValue()||'').toUpperCase();
      if((m==='MAÑANA'||m==='M') && (t==='TARDE'||t==='T')){ res.push({colM:c,colT:c+1}); c++; }
    }
    return res;
  }

  function _findGuideBlockCols(sheet, codigo){
    var lastCol = sheet.getLastColumn();
    var top = sheet.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(s){ return String(s||'').toUpperCase(); });
    var row2= sheet.getRange(2,1,1,lastCol).getDisplayValues()[0].map(function(s){ return String(s||'').toUpperCase(); });
    var codeUp = String(codigo||'').toUpperCase();
    for(var c=1;c<=lastCol;c++){
      var t = top[c-1];
      if(t.startsWith(codeUp+' ') || t.startsWith(codeUp+'—') || t===codeUp){
        var isM=(row2[c-1]==='MAÑANA'||row2[c-1]==='M'), isT=(row2[c]==='TARDE'||row2[c]==='T');
        if(isM && isT) return {colM:c, colT:c+1};
      }
    }
    return {colM:0,colT:0};
  }

  /* ======== util fechas en MASTER ======== */
  function findRowByISO(sheet, iso){
    var lr = sheet.getLastRow();
    var vals = sheet.getRange(3,1,Math.max(0,lr-2),1).getDisplayValues();
    for(var i=0;i<vals.length;i++){
      if(Sherpas.Util.toISO(Sherpas.Util.fromDisplay(vals[i][0]))===iso) return i+3;
    }
    return -1;
  }

  /* ======== sincronización guía→master para un mes ======== */
  function syncFromGuidesForMonth(ym){
    var ss = _open();
    var ms = ss.getSheetByName(ym); if(!ms) return;
    var regs = Sherpas.RegistryRepo.list();

    regs.forEach(function(reg){
      ensureGuideColumns(reg.codigo, reg.nombre);
      var cols = _findGuideBlockCols(ms, reg.codigo); if(!cols.colM||!cols.colT) return;
      var gss = SpreadsheetApp.openById(reg.fileId);
      var gsh = gss.getSheetByName(ym); if(!gsh) return;

      var m = ym.match(Sherpas.CFG.MASTER_MONTH_NAME);
      var yyyy=parseInt(m[2],10), mm=parseInt(m[1],10);
      var lastDay = new Date(yyyy, mm, 0).getDate();

      for(var d=1; d<=lastDay; d++){
        var pos = _findDateCellInGuide(gsh, new Date(yyyy,mm-1,d)); if(!pos) continue;
        var vM = String(gsh.getRange(pos.rowM,pos.col).getDisplayValue()||'').toUpperCase();
        var vT = String(gsh.getRange(pos.rowT,pos.col).getDisplayValue()||'').toUpperCase();
        var row = findRowByISO(ms, Sherpas.Util.toISO(new Date(yyyy,mm-1,d)));
        if(row>0){
          ms.getRange(row,cols.colM).setValue(vM==='NO DISPONIBLE' ? 'NO DISPONIBLE' : (vM.indexOf('ASIGNADO')===0? vM : ''));
          ms.getRange(row,cols.colT).setValue(vT==='NO DISPONIBLE' ? 'NO DISPONIBLE' : (vT.indexOf('ASIGNADO')===0? vT : ''));
        }
      }
    });
    applyDVandCF_All();
  }

  /* ======== borrar columnas de un guía en todas las pestañas ======== */
  function removeGuideColumnsEverywhere(codigo){
    var ss = _open();
    _listMonthTabs(ss).forEach(function(ms){
      var c = _findGuideBlockCols(ms, codigo);
      if(c.colM && c.colT) ms.deleteColumns(c.colM, 2);
    });
  }

  /* ======== helper para GUÍA (búsqueda de fecha en cuadrícula guía) ======== */
  function _findDateCellInGuide(sheet, dateObj){
    var mm = Sherpas.Util.pad2(dateObj.getMonth()+1); var yyyy = dateObj.getFullYear();
    if(sheet.getName()!== mm+'_'+yyyy) return null;

    // cuadrícula: por semanas, 3 filas: número / MAÑANA / TARDE
    for(var w=0; w<6; w++){
      var rowNum = 2 + (w*3);
      for(var col=1; col<=7; col++){
        var d = parseInt(sheet.getRange(rowNum, col).getDisplayValue(), 10);
        if(d===dateObj.getDate()) return { rowM:rowNum+1, rowT:rowNum+2, col:col };
      }
    }
    return null;
  }

  /* ======== NUEVA: función para obtener columnas de guía (corrige referencia faltante) ======== */
  function findGuideColumns(sheet, codigo) {
    return _findGuideBlockCols(sheet, codigo);
  }

  return {
    ensureMonthsFromOct: ensureMonthsFromOct,
    ensureGuideColumns: ensureGuideColumns,
    applyDVandCF_All: applyDVandCF_All,
    findRowByISO: findRowByISO,
    syncFromGuidesForMonth: syncFromGuidesForMonth,
    removeGuideColumnsEverywhere: removeGuideColumnsEverywhere,
    findGuideColumns: findGuideColumns
  };
})();

// Función temporal ejecutable para regenerar meses
function ejecutarEnsureMonths() {
  Sherpas.MasterBook.ensureMonthsFromOct();
}