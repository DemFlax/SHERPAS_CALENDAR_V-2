/** 07_guia.gs — Libro GUÍA (dominio)
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.GuideBook = (function(){
  'use strict';

  function _openById(id){ return SpreadsheetApp.openById(id); }

  function _rowsNeededFromSheetName_(sheet){
    var p = Sherpas.Util.parseTab_MMYYYY(sheet.getName()); 
    if(!p) return {rows:1, meta:null};
    var meta = Sherpas.Util.monthMeta(p.yyyy, p.mm);
    return { rows: 1 + meta.weeks*3, meta: meta }; // 1 cabecera + 3 por semana
  }

  /** Normaliza nº de filas del mes SIN tocar DV */
  function normalize_(sheet){
    var need = _rowsNeededFromSheetName_(sheet); 
    if(!need.meta) return;
    
    // CORRECCIÓN: Aplicar ensureExactRows mejorado a través de SheetWriter
    Sherpas.SheetWriter.ensureExactRows(sheet, need.rows);
    
    // Limpieza preventiva de contenido en filas excesivas
    var maxRows = sheet.getMaxRows();
    var contentLimit = need.rows + 2; // Solo 2 filas de buffer para contenido
    
    if(maxRows > contentLimit) {
      var excessStart = contentLimit + 1;
      var excessCount = Math.min(10, maxRows - contentLimit); // Limpiar máximo 10 filas por vez
      try {
        sheet.getRange(excessStart, 1, excessCount, 7).clearContent();
      } catch(e) {
        console.warn('Error limpiando contenido excesivo en ' + sheet.getName() + ':', e);
      }
    }
  }

  function buildMonth(guideIdOrSS, dateObj){
    var ss = (typeof guideIdOrSS==='string')? _openById(guideIdOrSS) : guideIdOrSS;
    var y = dateObj.getFullYear(), m = dateObj.getMonth()+1, mm = Sherpas.Util.pad2(m);
    var tab = mm+'_'+y;
    var sh = ss.getSheetByName(tab) || ss.insertSheet(tab);

    // Cabecera + cuadrícula por lotes
    var model = Sherpas.Util.buildGuideMonthArrays(y, m, 2);
    sh.clear(); sh.setFrozenRows(1);
    
    // CORRECCIÓN: Usar ensureExactRows mejorado en lugar de la versión básica
    Sherpas.SheetWriter.ensureExactRows(sh, model.rowsNeeded);
    Sherpas.SheetWriter.setValues2D(sh, 1, 1, [model.header]);
    if(model.grid.length) Sherpas.SheetWriter.setValues2D(sh, 2, 1, model.grid);

    // DV solo en M/T válidos
    var rule = Sherpas.DVManager.buildListRule(Sherpas.CFG.GUIDE_DV);
    Sherpas.DVManager.applyRuleToA1List(sh, model.mtA1, rule);

    // CF + Protección
    Sherpas.CFManager.setGuideRules(sh);
    Sherpas.ProtectionManager.protectSheetExcept(sh, model.mtA1, 'Sherpas — Guía (solo M/T editables)');
    return sh;
  }

  function applyDV(sheet){
    var p = Sherpas.Util.parseTab_MMYYYY(sheet.getName()); if(!p) return;
    normalize_(sheet);
    var meta = Sherpas.Util.monthMeta(p.yyyy, p.mm);
    var mtA1 = Sherpas.Util.monthMT_A1_FromMeta(meta, 2);
    var rule = Sherpas.DVManager.buildListRule(Sherpas.CFG.GUIDE_DV);
    Sherpas.DVManager.applyRuleToA1List(sheet, mtA1, rule);
  }

  function applyCF(sheet){ Sherpas.CFManager.setGuideRules(sheet); }

  function protectEditableMT(sheet){
    var p = Sherpas.Util.parseTab_MMYYYY(sheet.getName()); if(!p) return;
    normalize_(sheet);
    var meta = Sherpas.Util.monthMeta(p.yyyy, p.mm);
    var mtA1 = Sherpas.Util.monthMT_A1_FromMeta(meta, 2);
    Sherpas.ProtectionManager.protectSheetExcept(sheet, mtA1, 'Sherpas — Guía (solo M/T editables)');
  }

  function _findDateCellInGuide(sheet, dateObj){
    var mm = Sherpas.Util.pad2(dateObj.getMonth()+1), yyyy = dateObj.getFullYear();
    if(sheet.getName()!== mm+'_'+yyyy) return null;
    for(var w=0; w<6; w++){
      var rowNum = 2 + (w*3);
      for(var col=1; col<=7; col++){
        var d = parseInt(sheet.getRange(rowNum, col).getDisplayValue(), 10);
        if(d===dateObj.getDate()) return { rowM:rowNum+1, rowT:rowNum+2, col:col };
      }
    }
    return null;
  }

  function writeCell(guideId, tabName, fechaISO, turnoLabel, texto, lockIt){
    var ss = _openById(guideId);
    var sh = ss.getSheetByName(tabName); if(!sh) return;
    normalize_(sh);
    var pos = _findDateCellInGuide(sh, new Date(fechaISO+'T00:00:00')); if(!pos) return;
    var r = (turnoLabel==='MAÑANA')? pos.rowM : pos.rowT;
    var cell = sh.getRange(r, pos.col);
    cell.setValue(texto);
    if(lockIt){
      try {
        cell.clearDataValidations();
        var prot = cell.protect().setDescription('Sherpas — asignado (solo MASTER)');
        prot.setWarningOnly(false);
      } catch(e) {
        console.warn('Error protegiendo celda asignada:', e);
      }
    }else{
      applyDV(sh); protectEditableMT(sh);
    }
  }

  /** Adopta: normaliza + DV + CF + Protección + onEdit instalable */
  function adoptGuide(guideId){
    var ss = _openById(guideId);
    ss.getSheets().forEach(function(sh){
      if(Sherpas.CFG.GUIDE_MONTH_NAME.test(sh.getName())){
        normalize_(sh);
        applyDV(sh); applyCF(sh); protectEditableMT(sh);
      }
    });
    
    // CORRECCIÓN: Verificar que TriggerSvc existe antes de usar
    if(typeof Sherpas.TriggerSvc !== 'undefined' && Sherpas.TriggerSvc.ensureOnEditForSpreadsheet) {
      try {
        Sherpas.TriggerSvc.ensureOnEditForSpreadsheet(guideId, 'Sherpas.Triggers.onEditGuide');
      } catch(e) {
        console.warn('Error instalando trigger onEdit para guía adoptado:', e);
      }
    }
  }

  function normalize(sheet){ normalize_(sheet); }

  return {
    buildMonth: buildMonth,
    applyDV: applyDV,
    applyCF: applyCF,
    protectEditableMT: protectEditableMT,
    writeCell: writeCell,
    adoptGuide: adoptGuide,
    normalize: normalize
  };
})();