/** 11_sync_controller.js — Controlador de Sincronización Bidireccional
 *  Maneja sincronización tiempo real entre MASTER y calendarios GUÍA
 */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.SyncController = (function() {
  'use strict';

  const SYNC_LOCK_TIMEOUT = 30000; // 30 segundos
  const BATCH_SIZE = 50; // Procesar máximo 50 cambios por lote

  /**
   * Configuración de sincronización automática
   */
  function setupAutoSync() {
    try {
      // Limpiar triggers existentes
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(trigger => {
        if(trigger.getHandlerFunction() === 'autoSync') {
          ScriptApp.deleteTrigger(trigger);
        }
      });

      // Crear nuevo trigger cada 5 minutos
      ScriptApp.newTrigger('autoSync')
        .timeBased()
        .everyMinutes(5)  // ACTUALIZADO: Era 2, ahora 5 minutos
        .create();

      console.log('Trigger de sincronización automática configurado (cada 5 minutos)');
      return true;

    } catch(error) {
      console.error('Error configurando sincronización automática:', error);
      return false;
    }
  }

  /**
   * Función principal de sincronización (llamada por trigger)
   */
  function autoSync() {
    const lock = LockService.getScriptLock();
    
    try {
      // Intentar obtener lock por 30 segundos
      if(!lock.tryLock(SYNC_LOCK_TIMEOUT)) {
        console.warn('No se pudo obtener lock para sincronización - otra ejecución en progreso');
        return;
      }

      console.log('Iniciando sincronización automática...');
      
      // Ejecutar sincronización
      const result = syncAllGuides();
      
      console.log(`Sincronización completada: ${result.synced} guías sincronizadas, ${result.errors} errores`);

    } catch(error) {
      console.error('Error en sincronización automática:', error);
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * Sincroniza todos los calendarios de guías registrados
   */
  function syncAllGuides() {
    const result = { synced: 0, errors: 0, details: [] };
    
    try {
      const guides = Sherpas.RegistryRepo.list();
      const activeMonths = getCurrentActiveMonths();
      
      guides.forEach(guide => {
        try {
          activeMonths.forEach(month => {
            const syncResult = syncGuideForMonth(guide, month);
            if(syncResult.success) {
              result.synced++;
            } else {
              result.errors++;
              result.details.push(`${guide.codigo}-${month}: ${syncResult.error}`);
            }
          });
        } catch(error) {
          result.errors++;
          result.details.push(`${guide.codigo}: ${error.message}`);
        }
      });

    } catch(error) {
      console.error('Error en syncAllGuides:', error);
      result.errors++;
    }

    return result;
  }

  /**
   * Obtiene los meses activos actuales
   */
  function getCurrentActiveMonths() {
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth() + 1; // JavaScript months are 0-based
    
    const months = [];
    
    // Desde octubre del año actual hasta diciembre
    for(let month = Math.max(10, currentMonth); month <= 12; month++) {
      months.push(`${String(month).padStart(2, '0')}_${currentYear}`);
    }
    
    return months;
  }

  /**
   * Sincroniza un guía específico para un mes específico
   */
  function syncGuideForMonth(guide, monthTab) {
    try {
      // Abrir sheets
      const masterSS = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID)
      );
      const guideSS = SpreadsheetApp.openById(guide.fileId);
      
      const masterSheet = masterSS.getSheetByName(monthTab);
      const guideSheet = guideSS.getSheetByName(monthTab);
      
      if(!masterSheet || !guideSheet) {
        return { success: false, error: 'Sheets no encontrados' };
      }

      // Sincronizar cambios bidireccionales
      const changes = detectChanges(masterSheet, guideSheet, guide.codigo);
      applyChanges(masterSheet, guideSheet, changes, guide);

      return { success: true, changes: changes.length };

    } catch(error) {
      console.error(`Error sincronizando ${guide.codigo} para ${monthTab}:`, error);
      return { success: false, error: error.message };
    }
  }

  /**
   * Detecta cambios entre MASTER y calendario GUÍA
   */
  function detectChanges(masterSheet, guideSheet, guideCode) {
    const changes = [];
    
    try {
      // Obtener columnas del guía en MASTER
      const guideColumns = Sherpas.MasterBook.findGuideColumns(masterSheet, guideCode);
      if(!guideColumns.morning || !guideColumns.afternoon) return changes;

      // Comparar cada día del mes
      const monthData = getMonthDatesFromSheet(masterSheet);
      
      monthData.forEach(dateInfo => {
        const masterMorning = masterSheet.getRange(dateInfo.row, guideColumns.morning).getDisplayValue();
        const masterAfternoon = masterSheet.getRange(dateInfo.row, guideColumns.afternoon).getDisplayValue();
        
        const guideData = getGuideDataForDate(guideSheet, dateInfo.date);
        if(!guideData) return;
        
        // Detectar conflictos y cambios
        const change = analyzeChanges({
          date: dateInfo.date,
          row: dateInfo.row,
          master: { morning: masterMorning, afternoon: masterAfternoon },
          guide: guideData,
          columns: guideColumns
        });
        
        if(change) changes.push(change);
      });

    } catch(error) {
      console.error('Error detectando cambios:', error);
    }
    
    return changes;
  }

  /**
   * Obtiene fechas del mes desde el MASTER sheet
   */
  function getMonthDatesFromSheet(masterSheet) {
    const dates = [];
    const lastRow = masterSheet.getLastRow();
    
    for(let row = 3; row <= lastRow; row++) {
      const dateValue = masterSheet.getRange(row, 1).getValue();
      if(dateValue instanceof Date) {
        dates.push({ row: row, date: dateValue });
      }
    }
    
    return dates;
  }

  /**
   * Obtiene datos del guía para una fecha específica
   */
  function getGuideDataForDate(guideSheet, date) {
    try {
      const datePosition = findDateInGuideSheet(guideSheet, date);
      if(!datePosition) return null;
      
      return {
        morning: guideSheet.getRange(datePosition.morningRow, datePosition.col).getDisplayValue(),
        afternoon: guideSheet.getRange(datePosition.afternoonRow, datePosition.col).getDisplayValue(),
        position: datePosition
      };
    } catch(error) {
      return null;
    }
  }

  /**
   * Encuentra la posición de una fecha en el calendario guía
   */
  function findDateInGuideSheet(sheet, targetDate) {
    const day = targetDate.getDate();
    const month = targetDate.getMonth() + 1;
    const year = targetDate.getFullYear();
    
    // Verificar que la sheet corresponde al mes correcto
    const expectedTab = `${String(month).padStart(2, '0')}_${year}`;
    if(sheet.getName() !== expectedTab) return null;
    
    // Buscar en la cuadrícula calendario
    for(let week = 0; week < 6; week++) {
      const numberRow = 2 + (week * 3);
      
      for(let col = 1; col <= 7; col++) {
        const cellDay = parseInt(sheet.getRange(numberRow, col).getDisplayValue());
        if(cellDay === day) {
          return {
            col: col,
            morningRow: numberRow + 1,
            afternoonRow: numberRow + 2
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Analiza cambios y determina acciones necesarias
   */
  function analyzeChanges(data) {
    const { date, row, master, guide, columns } = data;
    
    // Reglas de jerarquía:
    // 1. NO DISPONIBLE (Guía) > TODO (Master no puede override)
    // 2. ASIGNADO (Master) > Disponibilidad (Guía no puede cambiar turno asignado)
    
    const changes = [];
    
    // Verificar mañana
    if(shouldSyncToMaster(master.morning, guide.morning)) {
      changes.push({
        type: 'GUIDE_TO_MASTER',
        target: 'MORNING',
        row: row,
        col: columns.morning,
        value: guide.morning,
        date: date
      });
    } else if(shouldSyncToGuide(master.morning, guide.morning)) {
      changes.push({
        type: 'MASTER_TO_GUIDE',
        target: 'MORNING',
        position: guide.position,
        value: master.morning,
        date: date
      });
    }
    
    // Verificar tarde
    if(shouldSyncToMaster(master.afternoon, guide.afternoon)) {
      changes.push({
        type: 'GUIDE_TO_MASTER',
        target: 'AFTERNOON',
        row: row,
        col: columns.afternoon,
        value: guide.afternoon,
        date: date
      });
    } else if(shouldSyncToGuide(master.afternoon, guide.afternoon)) {
      changes.push({
        type: 'MASTER_TO_GUIDE',
        target: 'AFTERNOON',
        position: guide.position,
        value: master.afternoon,
        date: date
      });
    }
    
    return changes.length > 0 ? changes : null;
  }

  /**
   * Determina si debe sincronizar de GUÍA a MASTER
   */
  function shouldSyncToMaster(masterValue, guideValue) {
    // Guía marca NO DISPONIBLE y MASTER no lo tiene
    if(guideValue === 'NO DISPONIBLE' && masterValue !== 'NO DISPONIBLE') {
      return true;
    }
    
    // Guía libera NO DISPONIBLE
    if(masterValue === 'NO DISPONIBLE' && (guideValue === 'MAÑANA' || guideValue === 'TARDE')) {
      return true;
    }
    
    return false;
  }

  /**
   * Determina si debe sincronizar de MASTER a GUÍA
   */
  function shouldSyncToGuide(masterValue, guideValue) {
    // MASTER asigna turno
    if(masterValue.includes('ASIGNADO') && !guideValue.includes('ASIGNADO')) {
      return true;
    }
    
    // MASTER libera turno asignado
    if(!masterValue.includes('ASIGNADO') && guideValue.includes('ASIGNADO')) {
      return true;
    }
    
    return false;
  }

  /**
   * Aplica los cambios detectados
   */
  function applyChanges(masterSheet, guideSheet, changes, guide) {
    changes.forEach(changeGroup => {
      if(!Array.isArray(changeGroup)) changeGroup = [changeGroup];
      
      changeGroup.forEach(change => {
        try {
          if(change.type === 'GUIDE_TO_MASTER') {
            masterSheet.getRange(change.row, change.col).setValue(change.value);
          } else if(change.type === 'MASTER_TO_GUIDE') {
            const targetRow = change.target === 'MORNING' ? 
              change.position.morningRow : change.position.afternoonRow;
            
            guideSheet.getRange(targetRow, change.position.col).setValue(change.value);
            
            // Proteger celda si es asignación
            if(change.value.includes('ASIGNADO')) {
              protectAssignedCell(guideSheet, targetRow, change.position.col);
            }
          }
        } catch(error) {
          console.error('Error aplicando cambio:', error);
        }
      });
    });
  }

  /**
   * Protege una celda asignada en el calendario guía
   */
  function protectAssignedCell(sheet, row, col) {
    try {
      const range = sheet.getRange(row, col);
      const protection = range.protect();
      protection.setDescription('Sherpas — Turno asignado (solo MASTER puede modificar)');
      protection.setWarningOnly(false);
      protection.removeEditors(protection.getEditors());
    } catch(error) {
      console.warn('No se pudo proteger celda asignada:', error);
    }
  }

  /**
   * NUEVA FUNCIÓN: Limpia filas excesivas en todas las hojas
   */
  function cleanExcessRows() {
    try {
      let sheetsProcessed = 0;
      let errorsFound = 0;

      // Limpiar MASTER
      const masterResult = cleanMasterSheets();
      sheetsProcessed += masterResult.processed;
      errorsFound += masterResult.errors;

      // Limpiar calendarios de guías
      const guidesResult = cleanGuideCalendars();
      sheetsProcessed += guidesResult.processed;
      errorsFound += guidesResult.errors;

      console.log(`Limpieza completada: ${sheetsProcessed} hojas procesadas, ${errorsFound} errores`);
      return errorsFound === 0;

    } catch(error) {
      console.error('Error en cleanExcessRows:', error);
      return false;
    }
  }

  /**
   * Limpia hojas MASTER
   */
  function cleanMasterSheets() {
    const result = { processed: 0, errors: 0 };
    
    try {
      const masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
      if(!masterId) return result;

      const masterSS = SpreadsheetApp.openById(masterId);
      const sheets = masterSS.getSheets();

      sheets.forEach(sheet => {
        if(Sherpas.CFG.MASTER_MONTH_NAME.test(sheet.getName())) {
          try {
            cleanMasterMonthSheet(sheet);
            result.processed++;
          } catch(error) {
            console.error(`Error limpiando sheet MASTER ${sheet.getName()}:`, error);
            result.errors++;
          }
        }
      });

    } catch(error) {
      console.error('Error en cleanMasterSheets:', error);
      result.errors++;
    }

    return result;
  }

  /**
   * Limpia una hoja mensual del MASTER
   */
  function cleanMasterMonthSheet(sheet) {
    const match = sheet.getName().match(Sherpas.CFG.MASTER_MONTH_NAME);
    if(!match) return;

    const month = parseInt(match[1], 10);
    const year = parseInt(match[2], 10);
    const lastDay = new Date(year, month, 0).getDate();
    const optimalRows = lastDay + 2; // Cabecera + días

    Sherpas.SheetWriter.ensureExactRows(sheet, optimalRows);
  }

  /**
   * Limpia calendarios de guías
   */
  function cleanGuideCalendars() {
    const result = { processed: 0, errors: 0 };
    
    try {
      const guides = Sherpas.RegistryRepo.list();
      
      guides.forEach(guide => {
        try {
          const guideSS = SpreadsheetApp.openById(guide.fileId);
          const sheets = guideSS.getSheets();
          
          sheets.forEach(sheet => {
            if(Sherpas.CFG.GUIDE_MONTH_NAME.test(sheet.getName())) {
              try {
                cleanGuideMonthSheet(sheet);
                result.processed++;
              } catch(error) {
                console.error(`Error limpiando sheet GUÍA ${guide.codigo}-${sheet.getName()}:`, error);
                result.errors++;
              }
            }
          });
          
        } catch(error) {
          console.error(`Error accediendo a calendario de guía ${guide.codigo}:`, error);
          result.errors++;
        }
      });

    } catch(error) {
      console.error('Error en cleanGuideCalendars:', error);
      result.errors++;
    }

    return result;
  }

  /**
   * Limpia una hoja mensual de guía
   */
  function cleanGuideMonthSheet(sheet) {
    const parsed = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
    if(!parsed) return;

    const meta = Sherpas.Util.monthMeta(parsed.yyyy, parsed.mm);
    const optimalRows = 1 + meta.weeks * 3; // Cabecera + cuadrícula

    Sherpas.SheetWriter.ensureExactRows(sheet, optimalRows);
  }

  /**
   * Función de mantenimiento programado
   */
  function scheduledMaintenance() {
    try {
      // Ejecutar solo una vez por día
      const lastRun = PropertiesService.getScriptProperties().getProperty('LAST_MAINTENANCE_SYNC');
      const today = Sherpas.Util.toISO(new Date());
      
      if(lastRun === today) {
        return; // Ya se ejecutó hoy
      }
      
      // Ejecutar limpieza
      const success = cleanExcessRows();
      
      // Registrar ejecución
      PropertiesService.getScriptProperties().setProperty('LAST_MAINTENANCE_SYNC', today);
      
      console.log(`Mantenimiento programado completado: ${today}, éxito: ${success}`);
      
    } catch(error) {
      console.error('Error en mantenimiento programado:', error);
    }
  }

  // API Pública
  return {
    setupAutoSync: setupAutoSync,
    autoSync: autoSync,
    syncAllGuides: syncAllGuides,
    syncGuideForMonth: syncGuideForMonth,
    cleanExcessRows: cleanExcessRows,
    scheduledMaintenance: scheduledMaintenance
  };

})();

// Función global para trigger automático
function autoSync() {
  Sherpas.SyncController.autoSync();
}