/** 12_validation_engine.js — Sistema de Rollback Automático
 *  Implementa validación estricta con reversión automática de cambios inválidos
 *  y logging de auditoría
 */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

/** Motor de Validación con Rollback Automático */
Sherpas.ValidationEngine = (function(){
  'use strict';

  /** Configuración de validación */
  const VALIDATION_CONFIG = {
    MAX_FAILED_ATTEMPTS: 3,        // Máximo intentos fallidos antes de bloqueo temporal
    BLOCK_DURATION_MINUTES: 15,    // Duración del bloqueo temporal
    LOG_RETENTION_DAYS: 30,        // Días que se mantienen logs de auditoría
    NOTIFICATION_THRESHOLD: 2      // Intentos antes de notificar al manager
  };

  /** Tipos de validación */
  const VALIDATION_TYPES = {
    INVALID_VALUE: 'INVALID_VALUE',
    UNAUTHORIZED_EDIT: 'UNAUTHORIZED_EDIT', 
    PROTECTED_CELL: 'PROTECTED_CELL',
    TIME_RESTRICTION: 'TIME_RESTRICTION',
    HIERARCHY_VIOLATION: 'HIERARCHY_VIOLATION'
  };

  /** Estados de validación */
  const VALIDATION_STATES = {
    VALID: 'VALID',
    INVALID_REVERTED: 'INVALID_REVERTED',
    BLOCKED: 'BLOCKED',
    WARNING: 'WARNING'
  };

  /**
   * Valida un cambio de edición y aplica rollback si es necesario
   * @param {Object} editEvent - Evento de edición de Google Apps Script
   * @returns {Object} Resultado de validación
   */
  function validateAndRollback(editEvent) {
    const validationResult = {
      state: VALIDATION_STATES.VALID,
      type: null,
      message: '',
      reverted: false,
      logged: false,
      notified: false,
      timestamp: new Date()
    };

    try {
      const context = _buildValidationContext(editEvent);
      
      // 1. Validaciones básicas
      const basicValidation = _performBasicValidations(context);
      if(!basicValidation.isValid) {
        return _handleInvalidEdit(context, basicValidation, validationResult);
      }

      // 2. Validaciones de jerarquía de permisos
      const hierarchyValidation = _performHierarchyValidations(context);
      if(!hierarchyValidation.isValid) {
        return _handleInvalidEdit(context, hierarchyValidation, validationResult);
      }

      // 3. Validaciones temporales (14h antes del turno)
      const timeValidation = _performTimeValidations(context);
      if(!timeValidation.isValid) {
        return _handleInvalidEdit(context, timeValidation, validationResult);
      }

      // 4. Todas las validaciones pasaron
      validationResult.message = 'Cambio válido aplicado correctamente';
      return validationResult;

    } catch(error) {
      console.error('Error en ValidationEngine:', error);
      validationResult.state = VALIDATION_STATES.INVALID_REVERTED;
      validationResult.type = VALIDATION_TYPES.INVALID_VALUE;
      validationResult.message = 'Error interno de validación';
      return validationResult;
    }
  }

  /**
   * Construye contexto de validación desde el evento de edición
   */
  function _buildValidationContext(editEvent) {
    const range = editEvent.range;
    const sheet = range.getSheet();
    const spreadsheet = sheet.getParent();
    
    return {
      // Datos del evento
      range: range,
      newValue: editEvent.value,
      oldValue: editEvent.oldValue,
      user: Session.getActiveUser().getEmail(),
      
      // Datos del contexto
      sheet: sheet,
      sheetName: sheet.getName(),
      spreadsheet: spreadsheet,
      spreadsheetId: spreadsheet.getId(),
      
      // Posición
      row: range.getRow(),
      column: range.getColumn(),
      a1: range.getA1Notation(),
      
      // Tipo de hoja
      isGuideSheet: _isGuideSheet(spreadsheet, sheet),
      isMasterSheet: _isMasterSheet(spreadsheet, sheet),
      
      // Timestamp
      timestamp: new Date()
    };
  }

  /**
   * Determina si es una hoja de guía
   */
  function _isGuideSheet(spreadsheet, sheet) {
    try {
      const guias = Sherpas.RegistryRepo.list();
      const isGuideSpreadsheet = guias.some(g => g.fileId === spreadsheet.getId());
      const isMonthTab = Sherpas.CFG.GUIDE_MONTH_NAME.test(sheet.getName());
      return isGuideSpreadsheet && isMonthTab;
    } catch(e) {
      return false;
    }
  }

  /**
   * Determina si es una hoja MASTER
   */
  function _isMasterSheet(spreadsheet, sheet) {
    try {
      const masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
      const isMasterSpreadsheet = spreadsheet.getId() === masterId;
      const isMonthTab = Sherpas.CFG.MASTER_MONTH_NAME.test(sheet.getName());
      return isMasterSpreadsheet && isMonthTab;
    } catch(e) {
      return false;
    }
  }

  /**
   * Validaciones básicas de formato y valores permitidos
   */
  function _performBasicValidations(context) {
    // Validar valores permitidos según tipo de hoja
    if(context.isGuideSheet) {
      return _validateGuideValues(context);
    } else if(context.isMasterSheet) {
      return _validateMasterValues(context);
    }

    return { isValid: true };
  }

  /**
   * Valida valores en hojas de guía
   */
  function _validateGuideValues(context) {
    const newValue = String(context.newValue || '').toUpperCase();
    const allowedValues = Sherpas.CFG.GUIDE_DV;

    if(!allowedValues.includes(newValue)) {
      return {
        isValid: false,
        type: VALIDATION_TYPES.INVALID_VALUE,
        message: `Valor "${context.newValue}" no permitido. Use solo: ${allowedValues.join(', ')}`
      };
    }

    return { isValid: true };
  }

  /**
   * Valida valores en hojas MASTER
   */
  function _validateMasterValues(context) {
    const newValue = String(context.newValue || '').toUpperCase();
    
    // Determinar si es columna MAÑANA o TARDE
    const isMorning = _isMorningColumn(context);
    const allowedValues = isMorning ? Sherpas.CFG.MASTER_DV_M : Sherpas.CFG.MASTER_DV_T;

    if(!allowedValues.includes(newValue)) {
      return {
        isValid: false,
        type: VALIDATION_TYPES.INVALID_VALUE,
        message: `Valor "${context.newValue}" no permitido para ${isMorning ? 'MAÑANA' : 'TARDE'}. Use: ${allowedValues.join(', ')}`
      };
    }

    return { isValid: true };
  }

  /**
   * Determina si la columna es MAÑANA en MASTER
   */
  function _isMorningColumn(context) {
    try {
      const headerValue = context.sheet.getRange(2, context.column).getDisplayValue();
      return String(headerValue || '').toUpperCase() === 'MAÑANA' || String(headerValue || '').toUpperCase() === 'M';
    } catch(e) {
      return true; // Default a MAÑANA si hay error
    }
  }

  /**
   * Validaciones de jerarquía de permisos
   */
  function _performHierarchyValidations(context) {
    if(context.isMasterSheet) {
      return _validateMasterPermissions(context);
    } else if(context.isGuideSheet) {
      return _validateGuidePermissions(context);
    }

    return { isValid: true };
  }

  /**
   * Valida permisos en MASTER
   */
  function _validateMasterPermissions(context) {
    const currentValue = String(context.oldValue || '').toUpperCase();
    const newValue = String(context.newValue || '').toUpperCase();

    // MASTER no puede asignar sobre "NO DISPONIBLE" sin liberar primero
    if(currentValue === 'NO DISPONIBLE' && newValue.startsWith('ASIGNAR')) {
      return {
        isValid: false,
        type: VALIDATION_TYPES.HIERARCHY_VIOLATION,
        message: 'No se puede asignar sobre NO DISPONIBLE. El guía debe liberar primero.'
      };
    }

    return { isValid: true };
  }

  /**
   * Valida permisos en hoja de guía
   */
  function _validateGuidePermissions(context) {
    const currentValue = String(context.oldValue || '').toUpperCase();
    const newValue = String(context.newValue || '').toUpperCase();

    // Guía no puede modificar turnos ASIGNADOS
    if(currentValue.startsWith('ASIGNADO')) {
      return {
        isValid: false,
        type: VALIDATION_TYPES.PROTECTED_CELL,
        message: 'No puedes modificar turnos ya asignados por el manager.'
      };
    }

    return { isValid: true };
  }

  /**
   * Validaciones temporales (14h antes del turno)
   */
  function _performTimeValidations(context) {
    const newValue = String(context.newValue || '').toUpperCase();
    
    // Solo aplicar restricción temporal para bloqueos/liberaciones de guía
    if(context.isGuideSheet && (newValue === 'NO DISPONIBLE' || newValue === 'REVERTIR' || newValue === '')) {
      return _validateTimeRestriction(context);
    }

    return { isValid: true };
  }

  /**
   * Valida restricción temporal de 14h
   */
  function _validateTimeRestriction(context) {
    try {
      // Calcular fecha del turno
      const fechaTurno = _calculateTurnDate(context);
      if(!fechaTurno) return { isValid: true }; // Si no se puede calcular, permitir

      // Determinar turno (MAÑANA o TARDE) 
      const turno = _determineTurnType(context);
      const slotRef = (turno === 'MAÑANA') ? 'M' : 'T1';

      // Calcular hora de inicio del turno
      const fechaISO = Sherpas.Util.toISO(fechaTurno);
      const startTime = Sherpas.CalendarSvc._shiftStart(fechaISO, slotRef);

      // Verificar si estamos dentro de la ventana de 14h
      const now = new Date();
      const diffHours = (startTime.getTime() - now.getTime()) / (1000 * 60 * 60);

      if(diffHours < Sherpas.CFG.LIMIT_HOURS) {
        return {
          isValid: false,
          type: VALIDATION_TYPES.TIME_RESTRICTION,
          message: `No se puede modificar disponibilidad. Faltan menos de ${Sherpas.CFG.LIMIT_HOURS}h para el turno.`
        };
      }

      return { isValid: true };

    } catch(error) {
      console.error('Error en validación temporal:', error);
      return { isValid: true }; // En caso de error, permitir el cambio
    }
  }

  /**
   * Calcula fecha del turno desde contexto de edición
   */
  function _calculateTurnDate(context) {
    try {
      const p = Sherpas.Util.parseTab_MMYYYY(context.sheetName);
      if(!p) return null;

      // Para hojas de guía: calcular día desde posición en cuadrícula
      const meta = Sherpas.Util.monthMeta(p.yyyy, p.mm);
      const row = context.row;
      const col = context.column;

      // Determinar semana y día dentro de la cuadrícula
      const semana = Math.floor((row - 2) / 3);
      const dow = col - 1; // 0-based day of week
      const dia = (semana * 7) + dow + 1 - meta.firstWeekday;

      if(dia >= 1 && dia <= meta.lastDay) {
        return new Date(p.yyyy, p.mm - 1, dia);
      }

      return null;
    } catch(e) {
      return null;
    }
  }

  /**
   * Determina tipo de turno (MAÑANA/TARDE) desde contexto
   */
  function _determineTurnType(context) {
    try {
      const row = context.row;
      const offset = (row - 2) % 3;
      return (offset === 1) ? 'MAÑANA' : 'TARDE';
    } catch(e) {
      return 'MAÑANA';
    }
  }

  /**
   * Maneja ediciones inválidas aplicando rollback y logging
   */
  function _handleInvalidEdit(context, validation, validationResult) {
    // 1. Revertir el cambio inmediatamente
    const reverted = _revertChange(context);
    
    // 2. Actualizar resultado
    validationResult.state = VALIDATION_STATES.INVALID_REVERTED;
    validationResult.type = validation.type;
    validationResult.message = validation.message;
    validationResult.reverted = reverted;

    // 3. Mostrar mensaje al usuario
    _showUserMessage(context, validation.message);

    // 4. Registrar en log de auditoría
    const logged = _logFailedAttempt(context, validation);
    validationResult.logged = logged;

    // 5. Verificar si necesita notificación al manager
    const shouldNotify = _checkNotificationThreshold(context);
    if(shouldNotify) {
      const notified = _notifyManager(context, validation);
      validationResult.notified = notified;
    }

    return validationResult;
  }

  /**
   * Revierte un cambio inválido
   */
  function _revertChange(context) {
    try {
      const oldValue = context.oldValue || '';
      context.range.setValue(oldValue);
      return true;
    } catch(error) {
      console.error('Error revirtiendo cambio:', error);
      return false;
    }
  }

  /**
   * Muestra mensaje al usuario
   */
  function _showUserMessage(context, message) {
    try {
      context.spreadsheet.toast(message, 'Cambio no permitido', 5);
    } catch(error) {
      console.error('Error mostrando mensaje:', error);
    }
  }

  /**
   * Registra intento fallido en log de auditoría
   */
  function _logFailedAttempt(context, validation) {
    try {
      const logEntry = {
        timestamp: context.timestamp.toISOString(),
        user: context.user,
        spreadsheetId: context.spreadsheetId,
        sheetName: context.sheetName,
        cell: context.a1,
        oldValue: context.oldValue,
        newValue: context.newValue,
        validationType: validation.type,
        message: validation.message,
        reverted: true
      };

      // Guardar en hoja de logs si existe
      _appendToAuditLog(logEntry);
      return true;

    } catch(error) {
      console.error('Error guardando log de auditoría:', error);
      return false;
    }
  }

  /**
   * Añade entrada al log de auditoría
   */
  function _appendToAuditLog(logEntry) {
    try {
      const masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
      const masterSS = SpreadsheetApp.openById(masterId);
      
      let logSheet = masterSS.getSheetByName('AUDIT_LOG');
      if(!logSheet) {
        logSheet = masterSS.insertSheet('AUDIT_LOG');
        // Crear cabeceras
        const headers = ['TIMESTAMP', 'USER', 'SPREADSHEET_ID', 'SHEET', 'CELL', 'OLD_VALUE', 'NEW_VALUE', 'TYPE', 'MESSAGE'];
        logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }

      // Añadir entrada
      const row = [
        logEntry.timestamp,
        logEntry.user,
        logEntry.spreadsheetId,
        logEntry.sheetName,
        logEntry.cell,
        logEntry.oldValue,
        logEntry.newValue,
        logEntry.validationType,
        logEntry.message
      ];

      logSheet.appendRow(row);

    } catch(error) {
      console.error('Error escribiendo audit log:', error);
    }
  }

  /**
   * Verifica si se debe notificar al manager
   */
  function _checkNotificationThreshold(context) {
    try {
      // Contar intentos recientes del usuario
      const recentAttempts = _getRecentFailedAttempts(context.user, 60); // últimos 60 minutos
      return recentAttempts >= VALIDATION_CONFIG.NOTIFICATION_THRESHOLD;
    } catch(error) {
      console.error('Error verificando threshold:', error);
      return false;
    }
  }

  /**
   * Obtiene intentos fallidos recientes de un usuario
   */
  function _getRecentFailedAttempts(userEmail, minutesBack) {
    try {
      const masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
      const masterSS = SpreadsheetApp.openById(masterId);
      const logSheet = masterSS.getSheetByName('AUDIT_LOG');
      
      if(!logSheet) return 0;

      const cutoffTime = new Date(Date.now() - (minutesBack * 60 * 1000));
      const data = logSheet.getDataRange().getValues();
      
      let count = 0;
      for(let i = 1; i < data.length; i++) { // Skip header
        const timestamp = new Date(data[i][0]);
        const user = data[i][1];
        
        if(user === userEmail && timestamp > cutoffTime) {
          count++;
        }
      }

      return count;

    } catch(error) {
      console.error('Error contando intentos recientes:', error);
      return 0;
    }
  }

  /**
   * Notifica al manager sobre intentos repetidos
   */
  function _notifyManager(context, validation) {
    try {
      const subject = `🚨 Intentos repetidos de edición inválida - ${context.user}`;
      const html = `
        <div style="max-width:600px; margin:0 auto; font-family:Arial;">
          <div style="background:#dc3545; color:white; padding:20px; text-align:center;">
            <h1>⚠️ Alerta de Seguridad</h1>
          </div>
          <div style="background:#f8f9fa; padding:20px;">
            <p><strong>Usuario:</strong> ${context.user}</p>
            <p><strong>Hoja:</strong> ${context.sheetName}</p>
            <p><strong>Celda:</strong> ${context.a1}</p>
            <p><strong>Tipo de violación:</strong> ${validation.type}</p>
            <p><strong>Mensaje:</strong> ${validation.message}</p>
            <p><strong>Hora:</strong> ${context.timestamp.toLocaleString()}</p>
            <div style="background:#fff3cd; padding:15px; margin:15px 0;">
              <strong>Este usuario ha tenido múltiples intentos fallidos en la última hora.</strong>
            </div>
          </div>
        </div>
      `;

      // Enviar a emails de admin configurados
      const adminEmails = Sherpas.CFG.ADMIN_EMAILS || ['admin@example.com'];
      adminEmails.forEach(email => {
        Sherpas.MailSvc.send(email, subject, html);
      });

      return true;

    } catch(error) {
      console.error('Error notificando manager:', error);
      return false;
    }
  }

  /**
   * Limpia logs antiguos (mantenimiento)
   */
  function cleanOldAuditLogs() {
    try {
      const masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
      const masterSS = SpreadsheetApp.openById(masterId);
      const logSheet = masterSS.getSheetByName('AUDIT_LOG');
      
      if(!logSheet) return;

      const cutoffDate = new Date(Date.now() - (VALIDATION_CONFIG.LOG_RETENTION_DAYS * 24 * 60 * 60 * 1000));
      const data = logSheet.getDataRange().getValues();
      
      let rowsToDelete = [];
      for(let i = data.length - 1; i >= 1; i--) { // Reverse order, skip header
        const timestamp = new Date(data[i][0]);
        if(timestamp < cutoffDate) {
          rowsToDelete.push(i + 1); // +1 because getRange is 1-based
        }
      }

      // Eliminar filas antiguas
      rowsToDelete.forEach(rowNum => {
        logSheet.deleteRow(rowNum);
      });

      console.log(`Limpiados ${rowsToDelete.length} logs antiguos`);

    } catch(error) {
      console.error('Error limpiando logs antiguos:', error);
    }
  }

  // API Pública
  return {
    validateAndRollback: validateAndRollback,
    cleanOldAuditLogs: cleanOldAuditLogs,
    VALIDATION_CONFIG: VALIDATION_CONFIG,
    VALIDATION_TYPES: VALIDATION_TYPES,
    VALIDATION_STATES: VALIDATION_STATES
  };
})();