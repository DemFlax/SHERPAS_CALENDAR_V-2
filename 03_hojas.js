/** 03_hojas.gs — Servicio de Hojas (DV, CF, Protección, escritura por lotes)
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

/** Servicio de escritura por lotes */
Sherpas.SheetWriter = (function(){
  'use strict';
  function setValues2D(sheet, startRow, startCol, grid){
    if(!grid || !grid.length) return;
    sheet.getRange(startRow, startCol, grid.length, grid[0].length).setValues(grid); // setValues en bloque
  }
  function ensureExactRows(sheet, rowsNeeded){
    var max = sheet.getMaxRows();
    if(max < rowsNeeded) sheet.insertRowsAfter(max, rowsNeeded-max);
    if(max > rowsNeeded) sheet.deleteRows(rowsNeeded+1, max-rowsNeeded);
  }
  function clearRange(sheet, r, c, nr, nc){
    sheet.getRange(r,c,nr,nc).clearContent();
  }
  function clearDataValidationsRect(sheet, r, c, nr, nc){
    sheet.getRange(r,c,nr,nc).clearDataValidations();
  }
  return { setValues2D:setValues2D, ensureExactRows:ensureExactRows, clearRange:clearRange, clearDVRect:clearDataValidationsRect };
})();

/** Validación de datos */
Sherpas.DVManager = (function(){
  'use strict';
  function buildListRule(values){
    return SpreadsheetApp.newDataValidation()
      .setAllowInvalid(false)
      .requireValueInList(values, true) // lista con desplegable
      .build(); // Range.setDataValidation(rule)
  }
  function applyRuleToA1List(sheet, a1List, rule){
    if(!a1List || !a1List.length) return;
    sheet.getRangeList(a1List).getRanges().forEach(function(r){ r.setDataValidation(rule); });
  }
  function clearAllInRect(sheet, r, c, nr, nc){
    sheet.getRange(r,c,nr,nc).clearDataValidations();
  }
  return { buildListRule:buildListRule, applyRuleToA1List:applyRuleToA1List, clearAllInRect:clearAllInRect };
})();

/** Formato condicional */
Sherpas.CFManager = (function(){
  'use strict';
  function setGuideRules(sheet){
    var last = sheet.getMaxRows();
    var rng = sheet.getRange(2,1, Math.max(0,last-1), 7);
    var base = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=TRUE').setBackground('#ffffff').setBold(false).setRanges([rng]).build();
    var nodisp = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('NO DISPONIBLE').setBackground('#ffcccc').setBold(true).setRanges([rng]).build();
    var asign = SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('ASIGNADO').setBackground('#c6efce').setBold(false).setRanges([rng]).build();
    sheet.setConditionalFormatRules([base, nodisp, asign]);
  }
  function setMasterRulesForCols(sheet, colM, colT){
    var lastRow = Math.max(3, sheet.getLastRow());
    var r1 = sheet.getRange(3,colM,lastRow-2,1);
    var r2 = sheet.getRange(3,colT,lastRow-2,1);
    var base = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=TRUE').setBackground('#ffffff').setBold(false).setRanges([r1,r2]).build();
    var nodisp = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('NO DISPONIBLE').setBackground('#ffcccc').setBold(true).setRanges([r1,r2]).build();
    var asign = SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('ASIGNADO').setBackground('#c6efce').setBold(false).setRanges([r1,r2]).build();
    sheet.setConditionalFormatRules([base, asign, nodisp]);
  }
  return { setGuideRules:setGuideRules, setMasterRulesForCols:setMasterRulesForCols };
})();

/** Protección de hoja con excepciones editables */
Sherpas.ProtectionManager = (function(){
  'use strict';
  function protectSheetExcept(sheet, a1EditableList, description){
    // limpia protecciones Sherpas previas
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){
      if((p.getDescription()||'').indexOf('Sherpas')===0) p.remove();
    });
    var prot = sheet.protect().setDescription(description || 'Sherpas — protección').setWarningOnly(false);
    var ranges = (a1EditableList||[]).map(function(a1){ return sheet.getRange(a1); });
    if(ranges.length) prot.setUnprotectedRanges(ranges); // define zonas editables
  }
  return { protectSheetExcept:protectSheetExcept };
})();
