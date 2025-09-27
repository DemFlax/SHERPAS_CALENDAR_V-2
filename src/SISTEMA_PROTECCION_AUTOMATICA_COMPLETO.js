/** SISTEMA DE PROTECCIÓN AUTOMÁTICA COMPLETO
 *  Protege automáticamente contra cualquier alteración inválida
 */

/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

/** 
 * MÓDULO DE PROTECCIÓN AUTOMÁTICA
 * Revierte automáticamente cualquier cambio inválido
 */
Sherpas.AutoProtection = (function(){
  'use strict';

  // Variables de control
  var PROTECTION_ACTIVE = true;
  var LAST_VALIDATION = {};

  /**
   * FUNCIÓN PRINCIPAL: Validación automática con rollback inmediato
   * Se ejecuta en CADA edición de celda
   */
  function validateAndRollbackImmediate(e) {
    if(!PROTECTION_ACTIVE || !e || !e.range) return;

    try {
      var sheet = e.range.getSheet();
      var sheetName = sheet.getName();
      
      // Solo aplicar a calendarios de guía (MM_YYYY)
      if(!Sherpas.CFG.GUIDE_MONTH_NAME.test(sheetName)) return;
      
      // Verificar que no es MASTER
      var masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
      if(sheet.getParent().getId() === masterId) return;

      var range = e.range;
      var newValue = String(e.value || '').toUpperCase().trim();
      var oldValue = String(e.oldValue || '').trim();
      var a1 = range.getA1Notation();
      var row = range.getRow();
      var col = range.getColumn();

      console.log('🛡️ AutoProtection validando:', sheetName, a1, 'Valor:', newValue);

      // 1. VALIDACIÓN DE ZONA EDITABLE
      if(!_isEditableCell(sheet, row, col)) {
        console.warn('❌ Celda no editable:', a1);
        range.setValue(oldValue);
        _showProtectionMessage('❌ Esta celda no se puede editar', sheet);
        return;
      }

      // 2. VALIDACIÓN DE VALOR PERMITIDO
      if(!_isValidValue(newValue)) {
        console.warn('❌ Valor inválido:', newValue, 'en', a1);
        var correctValue = _getCorrectDefaultValue(sheet, row, col);
        range.setValue(correctValue);
        _showProtectionMessage('❌ Valor "' + newValue + '" no permitido.\n\nSe restauró el valor correcto: "' + correctValue + '"', sheet);
        return;
      }

      // 3. VALIDACIÓN TEMPORAL (14h antes del turno)
      if(!_isWithinTimeLimit(sheet, row, col, newValue)) {
        console.warn('❌ Fuera de ventana temporal:', a1);
        range.setValue(oldValue || _getCorrectDefaultValue(sheet, row, col));
        _showProtectionMessage('❌ No se puede cambiar disponibilidad.\n\nFaltan menos de 14h para el turno.', sheet);
        return;
      }

      // 4. VALIDACIÓN DE JERARQUÍA (turnos asignados)
      if(!_respectsHierarchy(sheet, row, col, newValue, oldValue)) {
        console.warn('❌ Violación de jerarquía:', a1);
        range.setValue(oldValue);
        _showProtectionMessage('❌ No puedes modificar turnos asignados por el manager.', sheet);
        return;
      }

      // 5. Si llegó aquí, el cambio es válido
      console.log('✅ Cambio válido:', a1, newValue);
      
      // Aplicar formato inmediatamente
      _applyImmediateFormatting(sheet, range, newValue);
      
      // Sincronizar con MASTER
      _syncWithMaster(sheet, row, col, newValue);

    } catch(error) {
      console.error('Error en AutoProtection:', error);
      // En caso de error, revertir por seguridad
      if(e.range && e.oldValue !== undefined) {
        e.range.setValue(e.oldValue || '');
      }
    }
  }

  /**
   * Verifica si la celda es editable (solo M/T válidos)
   */
  function _isEditableCell(sheet, row, col) {
    try {
      var parsed = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
      if(!parsed) return false;

      // Verificar que es fila M/T (offset 1 o 2)
      var offset = (row - 2) % 3;
      if(offset !== 1 && offset !== 2) return false;

      // Verificar que hay número de día en la fila superior
      var dayNumber = parseInt(sheet.getRange(row - offset, col).getDisplayValue(), 10);
      return !isNaN(dayNumber) && dayNumber > 0;

    } catch(e) {
      return false;
    }
  }

  /**
   * Verifica si el valor está en la lista permitida
   */
  function _isValidValue(value) {
    if(!value) return true; // Vacío es válido
    return Sherpas.CFG.GUIDE_DV.includes(value) || value.startsWith('ASIGNADO');
  }

  /**
   * Obtiene el valor correcto por defecto según la posición
   */
  function _getCorrectDefaultValue(sheet, row, col) {
    var offset = (row - 2) % 3;
    return (offset === 1) ? 'MAÑANA' : 'TARDE';
  }

  /**
   * Verifica restricción temporal de 14h
   */
  function _isWithinTimeLimit(sheet, row, col, newValue) {
    try {
      // Solo aplicar a cambios de disponibilidad
      if(newValue !== 'NO DISPONIBLE' && newValue !== 'REVERTIR' && newValue !== '') {
        return true;
      }

      var parsed = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
      if(!parsed) return true;

      var offset = (row - 2) % 3;
      var dayNumber = parseInt(sheet.getRange(row - offset, col).getDisplayValue(), 10);
      if(isNaN(dayNumber)) return true;

      var fecha = new Date(parsed.yyyy, parsed.mm - 1, dayNumber);
      var turno = (offset === 1) ? 'MAÑANA' : 'TARDE';
      var slotRef = (turno === 'MAÑANA') ? 'M' : 'T1';
      
      var start = Sherpas.CalendarSvc._shiftStart(Sherpas.Util.toISO(fecha), slotRef);
      var diffH = (start.getTime() - (new Date()).getTime()) / 3600000;
      
      return diffH >= Sherpas.CFG.LIMIT_HOURS;

    } catch(e) {
      return true; // En caso de error, permitir
    }
  }

  /**
   * Verifica jerarquía de permisos (ASIGNADO no se puede cambiar)
   */
  function _respectsHierarchy(sheet, row, col, newValue, oldValue) {
    // Si el valor anterior era ASIGNADO, no se puede cambiar
    return !String(oldValue).toUpperCase().startsWith('ASIGNADO');
  }

  /**
   * Aplica formato inmediato según el valor
   */
  function _applyImmediateFormatting(sheet, range, value) {
    try {
      if(value === 'NO DISPONIBLE') {
        range.setBackground('#ffcccc').setFontColor('#cc0000').setFontWeight('bold');
      } else if(value.startsWith('ASIGNADO')) {
        range.setBackground('#c6efce').setFontColor('#006100').setFontWeight('bold');
      } else {
        range.setBackground('#ffffff').setFontColor('#000000').setFontWeight('normal');
      }
    } catch(e) {
      console.warn('Error aplicando formato:', e);
    }
  }

  /**
   * Sincroniza automáticamente con MASTER
   */
  function _syncWithMaster(sheet, row, col, newValue) {
    try {
      var ss = sheet.getParent();
      var masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
      if(!masterId) return;

      // Obtener código de guía
      var header = ss.getName();
      var match = String(header || '').match(/-(G\d{2})$/i);
      var codigo = match ? match[1].toUpperCase() : null;
      if(!codigo) return;

      // Calcular fecha
      var parsed = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
      if(!parsed) return;

      var offset = (row - 2) % 3;
      var dayNumber = parseInt(sheet.getRange(row - offset, col).getDisplayValue(), 10);
      if(isNaN(dayNumber)) return;

      var fecha = new Date(parsed.yyyy, parsed.mm - 1, dayNumber);
      var turno = (offset === 1) ? 'MAÑANA' : 'TARDE';

      // Abrir MASTER y sincronizar
      var master = SpreadsheetApp.openById(masterId);
      var masterSheet = master.getSheetByName(sheet.getName());
      if(!masterSheet) return;

      var cols = _findGuideColumnsInMaster(masterSheet, codigo);
      var rowMaster = Sherpas.MasterBook.findRowByISO(masterSheet, Sherpas.Util.toISO(fecha));
      
      if(cols && cols.colM && cols.colT && rowMaster >= 3) {
        var targetCol = (turno === 'MAÑANA') ? cols.colM : cols.colT;
        var masterCell = masterSheet.getRange(rowMaster, targetCol);
        
        // Solo sincronizar si no hay asignación del MASTER
        var currentMaster = String(masterCell.getDisplayValue() || '').toUpperCase();
        if(!currentMaster.startsWith('ASIGNADO')) {
          masterCell.setValue(newValue === 'NO DISPONIBLE' ? 'NO DISPONIBLE' : '');
          console.log('✅ Sincronizado con MASTER:', newValue);
        }
      }

    } catch(e) {
      console.warn('Error sincronizando con MASTER:', e);
    }
  }

  /**
   * Encuentra columnas de guía en MASTER
   */
  function _findGuideColumnsInMaster(sheet, codigo) {
    try {
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 2, lastCol).getDisplayValues();
      var topRow = headers[0];
      var secondRow = headers[1];
      
      for(var c = 0; c < topRow.length; c++) {
        var header = String(topRow[c] || '').toUpperCase();
        if(header.includes(codigo)) {
          var isM = String(secondRow[c] || '').toUpperCase() === 'MAÑANA';
          var isT = String(secondRow[c + 1] || '').toUpperCase() === 'TARDE';
          if(isM && isT) {
            return { colM: c + 1, colT: c + 2 };
          }
        }
      }
      return null;
    } catch(e) {
      return null;
    }
  }

  /**
   * Muestra mensaje de protección al usuario
   */
  function _showProtectionMessage(message, sheet) {
    try {
      sheet.getParent().toast(message, '🛡️ Protección Automática', 8);
    } catch(e) {
      console.warn('No se pudo mostrar mensaje:', e);
    }
  }

  /**
   * Configuración y control del sistema
   */
  function enableProtection() {
    PROTECTION_ACTIVE = true;
    console.log('🛡️ Protección automática ACTIVADA');
  }

  function disableProtection() {
    PROTECTION_ACTIVE = false;
    console.log('🛡️ Protección automática DESACTIVADA');
  }

  function isProtectionActive() {
    return PROTECTION_ACTIVE;
  }

  /**
   * Función de instalación masiva
   */
  function installAutoProtectionForAllGuides() {
    var guides = Sherpas.RegistryRepo.list();
    var installed = 0;

    guides.forEach(function(guide) {
      try {
        // Instalar trigger onEdit para protección automática
        Sherpas.TriggerSvc.ensureOnEditForSpreadsheet(guide.fileId, 'autoProtectionTrigger');
        installed++;
      } catch(e) {
        console.error('Error instalando protección para', guide.codigo, ':', e);
      }
    });

    console.log('🛡️ Protección automática instalada en', installed, 'calendarios');
    return installed;
  }

  // API pública
  return {
    validateAndRollbackImmediate: validateAndRollbackImmediate,
    enableProtection: enableProtection,
    disableProtection: disableProtection,
    isProtectionActive: isProtectionActive,
    installAutoProtectionForAllGuides: installAutoProtectionForAllGuides
  };
})();

/**
 * TRIGGER GLOBAL PARA PROTECCIÓN AUTOMÁTICA
 * Esta función se ejecuta automáticamente en cada edición
 */
function autoProtectionTrigger(e) {
  Sherpas.AutoProtection.validateAndRollbackImmediate(e);
}

/**
 * TRIGGER MEJORADO PARA CALENDARIOS DE GUÍA
 * Reemplaza la función onEditGuide existente
 */
function onEditGuideWithAutoProtection(e) {
  // Primero ejecutar protección automática
  Sherpas.AutoProtection.validateAndRollbackImmediate(e);
  
  // Si llegó aquí, el cambio es válido - ejecutar lógica adicional
  // (La sincronización ya se hace en AutoProtection)
}

/**
 * FUNCIÓN DE INICIALIZACIÓN
 * Ejecutar una vez para instalar el sistema completo
 */
function initializeAutoProtectionSystem() {
  try {
    console.log('🛡️ Inicializando sistema de protección automática...');
    
    // Activar protección
    Sherpas.AutoProtection.enableProtection();
    
    // Instalar triggers en todos los calendarios
    var installed = Sherpas.AutoProtection.installAutoProtectionForAllGuides();
    
    var message = '🛡️ Sistema de protección automática ACTIVADO\n\n';
    message += '✅ Instalado en ' + installed + ' calendarios\n';
    message += '🔒 Cualquier alteración inválida se revertirá automáticamente\n';
    message += '⚡ Protección en tiempo real activa';
    
    SpreadsheetApp.getActive().toast(message, 'Protección Automática', 10);
    
    return {
      success: true,
      installed: installed,
      active: true
    };
    
  } catch(error) {
    console.error('Error inicializando protección automática:', error);
    SpreadsheetApp.getActive().toast('❌ Error activando protección: ' + error.message, 'Error', 10);
    return {
      success: false,
      error: error.message
    };
  }
}