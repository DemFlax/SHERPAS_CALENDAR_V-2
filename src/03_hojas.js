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
    
    // ESTRATEGIA DEFINITIVA: Contenido hasta el final para que Google no muestre botón
    var targetRows = Math.max(rowsNeeded + 200, 1000); // Mínimo 1000 filas siempre
    
    // Expandir si es necesario
    if(max < targetRows) {
      sheet.insertRowsAfter(max, targetRows - max);
      console.log(`Sheet ${sheet.getName()}: Expandido de ${max} a ${targetRows} filas`);
    } else if(max > targetRows + 50) {
      // Solo reducir si hay demasiadas filas
      var toDelete = max - targetRows;
      sheet.deleteRows(targetRows + 1, toDelete);
      console.log(`Sheet ${sheet.getName()}: Reducido de ${max} a ${targetRows} filas`);
    }
    
    // CLAVE: Llenar zona de seguridad con contenido "invisible" que ocupe espacio
    fillSafetyZone(sheet, rowsNeeded, targetRows);
  }
  
  /**
   * Llena la zona de seguridad con contenido "fantasma" que ocupe espacio
   * pero sea invisible para el usuario
   */
  function fillSafetyZone(sheet, contentRows, totalRows) {
    if(totalRows <= contentRows + 5) return; // No hay zona de seguridad significativa
    
    var safetyStart = contentRows + 3; // Dejar 2 filas de buffer visual
    var safetyEnd = totalRows;
    var safetyRowCount = safetyEnd - safetyStart + 1;
    
    if(safetyRowCount <= 0) return;
    
    try {
      // Crear matriz de contenido "fantasma" - espacios invisibles que ocupan celdas
      var ghostContent = [];
      for(var i = 0; i < safetyRowCount; i++) {
        var row = [];
        for(var j = 0; j < 7; j++) {
          row.push(' '); // Espacio invisible que ocupa la celda
        }
        ghostContent.push(row);
      }
      
      // Aplicar contenido fantasma
      sheet.getRange(safetyStart, 1, safetyRowCount, 7).setValues(ghostContent);
      
      // Hacer invisible: fondo muy claro, texto muy claro
      sheet.getRange(safetyStart, 1, safetyRowCount, 7)
        .setBackground('#fefefe')
        .setFontColor('#fefefe')
        .setFontSize(6);
      
      // CRÍTICO: Proteger zona fantasma para que no sea editable
      protectGhostZone(sheet, safetyStart, safetyRowCount);
      
      console.log(`Zona fantasma creada en ${sheet.getName()}: filas ${safetyStart}-${safetyEnd}`);
      
    } catch(e) {
      console.error('Error creando zona fantasma:', e.message);
    }
  }
  
  /**
   * Protege la zona fantasma para que sea no editable
   */
  function protectGhostZone(sheet, startRow, rowCount) {
    try {
      var ghostRange = sheet.getRange(startRow, 1, rowCount, 7);
      var protection = ghostRange.protect();
      protection.setDescription('Sherpas — Zona Sistema (No Editar)');
      protection.setWarningOnly(false);
      
      // Quitar todos los editores = solo el propietario del script puede editar
      var editors = protection.getEditors();
      protection.removeEditors(editors);
      
    } catch(e) {
      console.warn('No se pudo proteger zona fantasma:', e.message);
    }
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
    // Solo proteger calendarios de guía, no MASTER ni REGISTRO
    var sheetName = sheet.getName();
    var spreadsheet = sheet.getParent();
    
    // Verificar si es calendario de guía (formato MM_YYYY y no es MASTER)
    var isGuideCalendar = Sherpas.CFG.GUIDE_MONTH_NAME.test(sheetName);
    var masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
    var isMasterSheet = masterId && spreadsheet.getId() === masterId;
    
    // Solo proteger si es calendario de guía (no MASTER)
    if(!isGuideCalendar || isMasterSheet) {
      console.log('Saltando protección para:', sheetName, 'Es calendario guía:', isGuideCalendar, 'Es MASTER:', isMasterSheet);
      return; // No proteger MASTER ni REGISTRO
    }
    
    // limpia protecciones Sherpas previas
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){
      if((p.getDescription()||'').indexOf('Sherpas')===0) p.remove();
    });
    
    // Proteger solo calendarios de guía con celdas editables específicas
    var prot = sheet.protect().setDescription(description || 'Sherpas — protección guía').setWarningOnly(false);
    var ranges = (a1EditableList||[]).map(function(a1){ return sheet.getRange(a1); });
    if(ranges.length) prot.setUnprotectedRanges(ranges); // define zonas editables
  }
  
  return { protectSheetExcept:protectSheetExcept };
})();