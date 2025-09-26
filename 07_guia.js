/** 07_guia.gs — Libro GUÍA (dominio)
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.GuideBook = (function(){
  'use strict';

  function _openById(id){ return SpreadsheetApp.openById(id); }

  function buildMonth(guideIdOrSS, dateObj){
    var ss = (typeof guideIdOrSS==='string')? _openById(guideIdOrSS) : guideIdOrSS;
    var y = dateObj.getFullYear(), m = dateObj.getMonth()+1, mm = Sherpas.Util.pad2(m);
    var tab = mm+'_'+y;
    var sh = ss.getSheetByName(tab) || ss.insertSheet(tab);

    // Cabecera + cuadrícula por lotes
    var model = Sherpas.Util.buildGuideMonthArrays(y, m, 2);
    sh.clear(); sh.setFrozenRows(1);
    Sherpas.SheetWriter.ensureExactRows(sh, model.rowsNeeded);
    Sherpas.SheetWriter.setValues2D(sh, 1, 1, [model.header]);
    if(model.grid.length) Sherpas.SheetWriter.setValues2D(sh, 2, 1, model.grid);

    // DV M/T
    var rule = Sherpas.DVManager.buildListRule(Sherpas.CFG.GUIDE_DV);
    Sherpas.DVManager.clearAllInRect(sh, 2, 1, Math.max(0, sh.getMaxRows()-1), 7);
    Sherpas.DVManager.applyRuleToA1List(sh, model.mtA1, rule);

    // CF + Protección
    Sherpas.CFManager.setGuideRules(sh);
    Sherpas.ProtectionManager.protectSheetExcept(sh, model.mtA1, 'Sherpas — Guía (solo M/T editables)');
    return sh;
  }

  function applyDV(sheet){
    var p = Sherpas.Util.parseTab_MMYYYY(sheet.getName()); if(!p) return;
    var meta = Sherpas.Util.monthMeta(p.yyyy, p.mm);
    var mtA1 = Sherpas.Util.monthMT_A1_FromMeta(meta, 2);
    var rule = Sherpas.DVManager.buildListRule(Sherpas.CFG.GUIDE_DV);
    Sherpas.DVManager.clearAllInRect(sheet, 2,1, Math.max(0, sheet.getMaxRows()-1), 7);
    Sherpas.DVManager.applyRuleToA1List(sheet, mtA1, rule);
  }

  function applyCF(sheet){ Sherpas.CFManager.setGuideRules(sheet); }

  function protectEditableMT(sheet){
    var p = Sherpas.Util.parseTab_MMYYYY(sheet.getName()); if(!p) return;
    var meta = Sherpas.Util.monthMeta(p.yyyy, p.mm);
    var mtA1 = Sherpas.Util.monthMT_A1_FromMeta(meta, 2);
    Sherpas.ProtectionManager.protectSheetExcept(sheet, mtA1, 'Sherpas — Guía (solo M/T editables)');
  }

  function _findDateCellInGuide(sheet, dateObj){
    var mm = Sherpas.Util.pad2(dateObj.getMonth()+1), yyyy = dateObj.getFullYear();
    if(sheet.getName()!== mm+'_'+yyyy) return null;
    // semanas: número/Mañana/Tarde
    for(var w=0; w<6; w++){
      var rowNum = 2 + (w*3);
      for(var col=1; col<=7; col++){
        var d = parseInt(sheet.getRange(rowNum, col).getDisplayValue(), 10);
        if(d===dateObj.getDate()) return { rowM:rowNum+1, rowT:rowNum+2, col:col };
      }
    }
    return null;
  }

  /**
   * Escribe en la hoja de guía y opcionalmente bloquea esa celda (asignación MASTER).
   * @param {string} guideId
   * @param {string} tabName  MM_YYYY
   * @param {string} fechaISO yyyy-mm-dd
   * @param {'MAÑANA'|'TARDE'} turnoLabel
   * @param {string} texto    'ASIGNADO …' | 'MAÑANA'|'TARDE'
   * @param {boolean} lockIt  si true, protege la celda individual
   */
  function writeCell(guideId, tabName, fechaISO, turnoLabel, texto, lockIt){
    var ss = _openById(guideId);
    var sh = ss.getSheetByName(tabName); if(!sh) return;
    var pos = _findDateCellInGuide(sh, new Date(fechaISO+'T00:00:00')); if(!pos) return;
    var r = (turnoLabel==='MAÑANA')? pos.rowM : pos.rowT;
    var cell = sh.getRange(r, pos.col);
    cell.setValue(texto);
    if(lockIt){
      cell.clearDataValidations();
      var prot = cell.protect().setDescription('Sherpas — asignado (solo MASTER)');
      prot.setWarningOnly(false);
    }else{
      applyDV(sh); protectEditableMT(sh);
    }
  }

  /** Adopta todas las pestañas MM_YYYY de un guía: DV+CF+Protección y trigger onEdit */
  function adoptGuide(guideId){
    var ss = _openById(guideId);
    ss.getSheets().forEach(function(sh){
      if(Sherpas.CFG.GUIDE_MONTH_NAME.test(sh.getName())){
        applyDV(sh); applyCF(sh); protectEditableMT(sh);
      }
    });
    Sherpas.TriggerSvc.ensureOnEditForSpreadsheet(guideId, 'Sherpas.Triggers.onEditGuide'); // instalable
  }

  return {
    buildMonth: buildMonth,
    applyDV: applyDV,
    applyCF: applyCF,
    protectEditableMT: protectEditableMT,
    writeCell: writeCell,
    adoptGuide: adoptGuide
  };
})();
