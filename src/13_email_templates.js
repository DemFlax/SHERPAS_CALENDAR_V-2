/** 13_email_templates.js — Templates Profesionales HTML
 *  Implementa templates especificados en PDF con diseño moderno
 */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

/** Servicio de Templates de Email */
Sherpas.EmailTemplates = (function(){
  'use strict';

  /** Configuración base de emails */
  const EMAIL_CONFIG = {
    FROM_NAME: 'Spain Food Sherpas',
    BRAND_COLOR: '#1a73e8',
    SUCCESS_COLOR: '#c6efce',
    WARNING_COLOR: '#fff3cd', 
    ERROR_COLOR: '#ffcccc',
    TEXT_COLOR: '#333333',
    FOOTER_COLOR: '#666666'
  };

  /** Template base HTML */
  const BASE_TEMPLATE = `
    <div style="max-width:600px; margin:0 auto; font-family:Arial, sans-serif; color:${EMAIL_CONFIG.TEXT_COLOR};">
      {{HEADER}}
      <div style="background:#f8f9fa; padding:20px; margin:20px 0;">
        {{CONTENT}}
      </div>
      {{FOOTER}}
    </div>
  `;

  /** Header principal */
  const HEADER_TEMPLATE = `
    <div style="background:${EMAIL_CONFIG.BRAND_COLOR}; color:white; padding:20px; text-align:center; border-radius:8px 8px 0 0;">
      <h1 style="margin:0; font-size:24px;">🍽️ Spain Food Sherpas</h1>
      <p style="margin:8px 0 0 0; opacity:0.9;">{{SUBTITLE}}</p>
    </div>
  `;

  /** Footer estándar */
  const FOOTER_TEMPLATE = `
    <div style="background:#f1f3f4; padding:15px; text-align:center; font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; border-radius:0 0 8px 8px;">
      <p style="margin:0;">Este email fue generado automáticamente por el Sistema de Gestión de Tours</p>
      <p style="margin:5px 0 0 0;">© 2025 Spain Food Sherpas - Madrid</p>
    </div>
  `;

  /**
   * Template de bienvenida para nuevo guía
   */
  function buildWelcomeTemplate(nombreGuia, codigoGuia, enlaceCalendario) {
    const content = `
      <div style="text-align:center; margin-bottom:20px;">
        <h2 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 10px 0;">¡Bienvenido al equipo!</h2>
        <p style="font-size:16px; margin:0;">Tu calendario personal está listo para usar</p>
      </div>
      
      <div style="background:white; padding:20px; border-radius:6px; margin:20px 0;">
        <div style="display:flex; align-items:center; margin-bottom:15px;">
          <div style="background:${EMAIL_CONFIG.BRAND_COLOR}; color:white; width:40px; height:40px; border-radius:50%; display:flex; align-items:center; justify-content:center; margin-right:15px; font-weight:bold;">
            👤
          </div>
          <div>
            <h3 style="margin:0; color:${EMAIL_CONFIG.TEXT_COLOR};">${nombreGuia}</h3>
            <p style="margin:0; color:${EMAIL_CONFIG.FOOTER_COLOR}; font-size:14px;">Código de guía: <strong>${codigoGuia}</strong></p>
          </div>
        </div>
        
        <div style="background:${EMAIL_CONFIG.WARNING_COLOR}; padding:15px; border-radius:4px; margin:15px 0;">
          <p style="margin:0; font-weight:bold;">📅 Accede a tu calendario:</p>
          <a href="${enlaceCalendario}" 
             style="display:inline-block; background:${EMAIL_CONFIG.BRAND_COLOR}; color:white; padding:10px 20px; text-decoration:none; border-radius:4px; margin-top:10px; font-weight:bold;">
            🚀 Abrir Mi Calendario
          </a>
        </div>
      </div>

      <div style="background:white; padding:20px; border-radius:6px;">
        <h3 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 15px 0;">📋 Instrucciones Importantes:</h3>
        <div style="background:#f8f9fa; padding:15px; border-left:4px solid ${EMAIL_CONFIG.BRAND_COLOR}; margin:10px 0;">
          <p style="margin:0 0 10px 0;"><strong>✅ Para marcar NO DISPONIBLE:</strong></p>
          <p style="margin:0; font-size:14px; color:${EMAIL_CONFIG.FOOTER_COLOR};">Selecciona la celda del día y turno, escribe "NO DISPONIBLE"</p>
        </div>
        <div style="background:#f8f9fa; padding:15px; border-left:4px solid ${EMAIL_CONFIG.SUCCESS_COLOR}; margin:10px 0;">
          <p style="margin:0 0 10px 0;"><strong>🟢 Los turnos asignados:</strong></p>
          <p style="margin:0; font-size:14px; color:${EMAIL_CONFIG.FOOTER_COLOR};">Aparecerán en verde y no podrás modificarlos</p>
        </div>
        <div style="background:#f8f9fa; padding:15px; border-left:4px solid ${EMAIL_CONFIG.ERROR_COLOR}; margin:10px 0;">
          <p style="margin:0 0 10px 0;"><strong>⚠️ Importante:</strong></p>
          <p style="margin:0; font-size:14px; color:${EMAIL_CONFIG.FOOTER_COLOR};">No modifiques ninguna otra parte del calendario</p>
        </div>
      </div>

      <div style="text-align:center; margin-top:20px; padding:15px; background:${EMAIL_CONFIG.WARNING_COLOR}; border-radius:6px;">
        <p style="margin:0; font-size:14px;">
          <strong>💡 ¿Necesitas ayuda?</strong><br>
          Contacta con el equipo de gestión para cualquier duda
        </p>
      </div>
    `;

    return _buildFullTemplate('¡Tu calendario de tours está listo!', content);
  }

  /**
   * Template de asignación de turno
   */
  function buildAssignmentTemplate(nombreGuia, codigoGuia, fecha, turno, enlaceCalendario) {
    const fechaFormateada = _formatearFecha(fecha);
    const turnoDescripcion = _describeTurno(turno);
    const horaInicio = _getHoraTurno(turno);

    const content = `
      <div style="text-align:center; margin-bottom:20px;">
        <div style="background:${EMAIL_CONFIG.SUCCESS_COLOR}; padding:15px; border-radius:50px; display:inline-block; margin-bottom:15px;">
          <span style="font-size:32px;">✅</span>
        </div>
        <h2 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 10px 0;">¡Nuevo Turno Asignado!</h2>
        <p style="font-size:16px; margin:0;">Se te ha asignado un tour</p>
      </div>
      
      <div style="background:white; padding:20px; border-radius:6px; margin:20px 0;">
        <div style="text-align:center; margin-bottom:20px;">
          <h3 style="color:${EMAIL_CONFIG.TEXT_COLOR}; margin:0 0 5px 0;">${nombreGuia} (${codigoGuia})</h3>
          <p style="color:${EMAIL_CONFIG.FOOTER_COLOR}; margin:0; font-size:14px;">Guía asignado</p>
        </div>
        
        <div style="background:${EMAIL_CONFIG.SUCCESS_COLOR}; padding:20px; border-radius:8px; text-align:center;">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap;">
            <div style="flex:1; min-width:120px; margin:5px;">
              <p style="margin:0; font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Fecha</p>
              <p style="margin:5px 0 0 0; font-size:18px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">📅 ${fechaFormateada}</p>
            </div>
            <div style="flex:1; min-width:120px; margin:5px;">
              <p style="margin:0; font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Turno</p>
              <p style="margin:5px 0 0 0; font-size:18px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">⏰ ${turnoDescripcion}</p>
            </div>
            <div style="flex:1; min-width:120px; margin:5px;">
              <p style="margin:0; font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Hora</p>
              <p style="margin:5px 0 0 0; font-size:18px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">🕐 ${horaInicio}</p>
            </div>
          </div>
        </div>
      </div>

      <div style="background:white; padding:20px; border-radius:6px;">
        <h3 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 15px 0;">📝 Recordatorios:</h3>
        <ul style="padding-left:20px; margin:0;">
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">Este turno aparece ahora en tu calendario personal</li>
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">Recibirás una invitación de Google Calendar</li>
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">No podrás modificar este turno desde tu calendario</li>
          <li style="margin-bottom:0; color:${EMAIL_CONFIG.TEXT_COLOR};">Contacta con gestión para cualquier cambio</li>
        </ul>
      </div>

      <div style="text-align:center; margin-top:20px;">
        <a href="${enlaceCalendario}" 
           style="display:inline-block; background:${EMAIL_CONFIG.BRAND_COLOR}; color:white; padding:12px 25px; text-decoration:none; border-radius:6px; font-weight:bold;">
          📅 Ver Mi Calendario Completo
        </a>
      </div>
    `;

    return _buildFullTemplate('¡Tienes un nuevo turno asignado!', content);
  }

  /**
   * Template de liberación de turno
   */
  function buildReleaseTemplate(nombreGuia, codigoGuia, fecha, turno, enlaceCalendario) {
    const fechaFormateada = _formatearFecha(fecha);
    const turnoDescripcion = _describeTurno(turno);

    const content = `
      <div style="text-align:center; margin-bottom:20px;">
        <div style="background:${EMAIL_CONFIG.WARNING_COLOR}; padding:15px; border-radius:50px; display:inline-block; margin-bottom:15px;">
          <span style="font-size:32px;">🔓</span>
        </div>
        <h2 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 10px 0;">Turno Liberado</h2>
        <p style="font-size:16px; margin:0;">Tu turno ha sido liberado</p>
      </div>
      
      <div style="background:white; padding:20px; border-radius:6px; margin:20px 0;">
        <div style="text-align:center; margin-bottom:20px;">
          <h3 style="color:${EMAIL_CONFIG.TEXT_COLOR}; margin:0 0 5px 0;">${nombreGuia} (${codigoGuia})</h3>
        </div>
        
        <div style="background:${EMAIL_CONFIG.WARNING_COLOR}; padding:20px; border-radius:8px; text-align:center;">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap;">
            <div style="flex:1; min-width:120px; margin:5px;">
              <p style="margin:0; font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Fecha</p>
              <p style="margin:5px 0 0 0; font-size:18px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">📅 ${fechaFormateada}</p>
            </div>
            <div style="flex:1; min-width:120px; margin:5px;">
              <p style="margin:0; font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Turno</p>
              <p style="margin:5px 0 0 0; font-size:18px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">⏰ ${turnoDescripcion}</p>
            </div>
            <div style="flex:1; min-width:120px; margin:5px;">
              <p style="margin:0; font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Estado</p>
              <p style="margin:5px 0 0 0; font-size:18px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">🔓 LIBERADO</p>
            </div>
          </div>
        </div>
      </div>

      <div style="background:white; padding:20px; border-radius:6px;">
        <h3 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 15px 0;">ℹ️ Información:</h3>
        <ul style="padding-left:20px; margin:0;">
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">El turno ha sido liberado de tu calendario</li>
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">La invitación de Google Calendar ha sido cancelada</li>
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">Ya puedes marcar disponibilidad para esta fecha si lo deseas</li>
          <li style="margin-bottom:0; color:${EMAIL_CONFIG.TEXT_COLOR};">El turno está disponible para reasignación</li>
        </ul>
      </div>

      <div style="text-align:center; margin-top:20px;">
        <a href="${enlaceCalendario}" 
           style="display:inline-block; background:${EMAIL_CONFIG.BRAND_COLOR}; color:white; padding:12px 25px; text-decoration:none; border-radius:6px; font-weight:bold;">
          📅 Actualizar Mi Disponibilidad
        </a>
      </div>
    `;

    return _buildFullTemplate('Turno liberado de tu calendario', content);
  }

  /**
   * Template de alerta para manager
   */
  function buildManagerAlertTemplate(userEmail, sheetName, cell, violationType, message, attemptCount) {
    const content = `
      <div style="text-align:center; margin-bottom:20px;">
        <div style="background:${EMAIL_CONFIG.ERROR_COLOR}; padding:15px; border-radius:50px; display:inline-block; margin-bottom:15px;">
          <span style="font-size:32px;">🚨</span>
        </div>
        <h2 style="color:#dc3545; margin:0 0 10px 0;">Alerta de Seguridad</h2>
        <p style="font-size:16px; margin:0;">Detectados múltiples intentos inválidos</p>
      </div>
      
      <div style="background:white; padding:20px; border-radius:6px; margin:20px 0;">
        <h3 style="color:#dc3545; margin:0 0 15px 0;">📊 Detalles del Incidente:</h3>
        
        <div style="background:#f8f9fa; padding:15px; border-radius:4px; margin:10px 0;">
          <div style="display:flex; justify-content:space-between; margin-bottom:10px;">
            <strong>👤 Usuario:</strong>
            <span style="color:${EMAIL_CONFIG.FOOTER_COLOR};">${userEmail}</span>
          </div>
          <div style="display:flex; justify-content:space-between; margin-bottom:10px;">
            <strong>📄 Hoja:</strong>
            <span style="color:${EMAIL_CONFIG.FOOTER_COLOR};">${sheetName}</span>
          </div>
          <div style="display:flex; justify-content:space-between; margin-bottom:10px;">
            <strong>📍 Celda:</strong>
            <span style="color:${EMAIL_CONFIG.FOOTER_COLOR};">${cell}</span>
          </div>
          <div style="display:flex; justify-content:space-between; margin-bottom:10px;">
            <strong>⚠️ Tipo:</strong>
            <span style="color:#dc3545; font-weight:bold;">${violationType}</span>
          </div>
          <div style="display:flex; justify-content:space-between;">
            <strong>🔢 Intentos:</strong>
            <span style="color:#dc3545; font-weight:bold;">${attemptCount}</span>
          </div>
        </div>

        <div style="background:${EMAIL_CONFIG.ERROR_COLOR}; padding:15px; border-radius:4px; border-left:4px solid #dc3545;">
          <p style="margin:0; font-weight:bold;">🛑 Mensaje de Error:</p>
          <p style="margin:5px 0 0 0; font-style:italic;">"${message}"</p>
        </div>
      </div>

      <div style="background:white; padding:20px; border-radius:6px;">
        <h3 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 15px 0;">📋 Acciones Recomendadas:</h3>
        <ul style="padding-left:20px; margin:0;">
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">Verificar que el usuario entiende las reglas del sistema</li>
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">Considerar contacto directo si los intentos persisten</li>
          <li style="margin-bottom:8px; color:${EMAIL_CONFIG.TEXT_COLOR};">Revisar permisos del usuario si es necesario</li>
          <li style="margin-bottom:0; color:${EMAIL_CONFIG.TEXT_COLOR};">Monitorear actividad futura de este usuario</li>
        </ul>
      </div>

      <div style="text-align:center; margin-top:20px; padding:15px; background:${EMAIL_CONFIG.WARNING_COLOR}; border-radius:6px;">
        <p style="margin:0; font-size:14px;">
          <strong>📧 Esta alerta se genera automáticamente</strong><br>
          El sistema continúa bloqueando cambios inválidos automáticamente
        </p>
      </div>
    `;

    return _buildFullTemplate('🚨 Alerta de Seguridad del Sistema', content);
  }

  /**
   * Template de resumen diario para manager
   */
  function buildDailySummaryTemplate(fecha, assignedTurns, availableGuides, conflicts) {
    const fechaFormateada = _formatearFecha(fecha);

    const content = `
      <div style="text-align:center; margin-bottom:20px;">
        <h2 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 10px 0;">📊 Resumen Diario</h2>
        <p style="font-size:16px; margin:0; color:${EMAIL_CONFIG.FOOTER_COLOR};">${fechaFormateada}</p>
      </div>
      
      <div style="background:white; padding:20px; border-radius:6px; margin:20px 0;">
        <h3 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 15px 0;">📈 Estadísticas del Día</h3>
        
        <div style="display:flex; justify-content:space-between; flex-wrap:wrap; margin-bottom:20px;">
          <div style="flex:1; min-width:150px; text-align:center; margin:10px; padding:15px; background:${EMAIL_CONFIG.SUCCESS_COLOR}; border-radius:6px;">
            <div style="font-size:24px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">${assignedTurns}</div>
            <div style="font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Turnos Asignados</div>
          </div>
          <div style="flex:1; min-width:150px; text-align:center; margin:10px; padding:15px; background:${EMAIL_CONFIG.WARNING_COLOR}; border-radius:6px;">
            <div style="font-size:24px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">${availableGuides}</div>
            <div style="font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Guías Disponibles</div>
          </div>
          <div style="flex:1; min-width:150px; text-align:center; margin:10px; padding:15px; background:${conflicts > 0 ? EMAIL_CONFIG.ERROR_COLOR : EMAIL_CONFIG.SUCCESS_COLOR}; border-radius:6px;">
            <div style="font-size:24px; font-weight:bold; color:${EMAIL_CONFIG.TEXT_COLOR};">${conflicts}</div>
            <div style="font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; text-transform:uppercase;">Conflictos</div>
          </div>
        </div>
      </div>

      <div style="text-align:center; margin-top:20px;">
        <p style="margin:0; font-size:14px; color:${EMAIL_CONFIG.FOOTER_COLOR};">
          Sistema funcionando correctamente • Próximo resumen: mañana
        </p>
      </div>
    `;

    return _buildFullTemplate('Resumen diario del sistema', content);
  }

  /**
   * Construye template completo con header y footer
   */
  function _buildFullTemplate(subtitle, content) {
    const header = HEADER_TEMPLATE.replace('{{SUBTITLE}}', subtitle);
    const footer = FOOTER_TEMPLATE;
    
    return BASE_TEMPLATE
      .replace('{{HEADER}}', header)
      .replace('{{CONTENT}}', content)
      .replace('{{FOOTER}}', footer);
  }

  /**
   * Formatea fecha en español
   */
  function _formatearFecha(fecha) {
    const fechaObj = (typeof fecha === 'string') ? new Date(fecha + 'T00:00:00') : fecha;
    const opciones = { 
      weekday: 'long', 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric',
      timeZone: 'Europe/Madrid'
    };
    return fechaObj.toLocaleDateString('es-ES', opciones);
  }

  /**
   * Describe turno de forma legible
   */
  function _describeTurno(turno) {
    const descripciones = {
      'M': 'Mañana',
      'MAÑANA': 'Mañana',
      'T1': 'Tarde 1',
      'T2': 'Tarde 2', 
      'T3': 'Tarde 3',
      'TARDE': 'Tarde'
    };
    return descripciones[turno] || turno;
  }

  /**
   * Obtiene hora de inicio del turno
   */
  function _getHoraTurno(turno) {
    const slot = turno === 'MAÑANA' ? 'M' : turno;
    const hora = Sherpas.CFG.SHIFT_TIMES[slot];
    return hora ? `${hora}h` : 'Por confirmar';
  }

  /**
   * Envía email usando template específico
   */
  function sendWelcome(nombreGuia, codigoGuia, email, enlaceCalendario) {
    const subject = `🎉 Tu calendario de tours está listo - ${nombreGuia}`;
    const html = buildWelcomeTemplate(nombreGuia, codigoGuia, enlaceCalendario);
    return Sherpas.MailSvc.send(email, subject, html);
  }

  /**
   * Envía notificación de asignación
   */
  function sendAssignment(nombreGuia, codigoGuia, email, fecha, turno, enlaceCalendario) {
    const fechaStr = _formatearFecha(fecha);
    const turnoStr = _describeTurno(turno);
    const subject = `✅ Asignación: ${fechaStr} - ${turnoStr}`;
    const html = buildAssignmentTemplate(nombreGuia, codigoGuia, fecha, turno, enlaceCalendario);
    return Sherpas.MailSvc.send(email, subject, html);
  }

  /**
   * Envía notificación de liberación
   */
  function sendRelease(nombreGuia, codigoGuia, email, fecha, turno, enlaceCalendario) {
    const fechaStr = _formatearFecha(fecha);
    const turnoStr = _describeTurno(turno);
    const subject = `🔓 Liberación: ${fechaStr} - ${turnoStr}`;
    const html = buildReleaseTemplate(nombreGuia, codigoGuia, fecha, turno, enlaceCalendario);
    return Sherpas.MailSvc.send(email, subject, html);
  }

  /**
   * Envía alerta de seguridad al manager
   */
  function sendManagerAlert(adminEmails, userEmail, sheetName, cell, violationType, message, attemptCount) {
    const subject = `🚨 Alerta de Seguridad: Intentos inválidos - ${userEmail}`;
    const html = buildManagerAlertTemplate(userEmail, sheetName, cell, violationType, message, attemptCount);
    
    const results = [];
    adminEmails.forEach(email => {
      results.push(Sherpas.MailSvc.send(email, subject, html));
    });
    return results;
  }

  /**
   * Envía resumen diario al manager
   */
  function sendDailySummary(adminEmails, fecha, assignedTurns, availableGuides, conflicts) {
    const fechaStr = _formatearFecha(fecha);
    const subject = `📊 Resumen Diario - ${fechaStr}`;
    const html = buildDailySummaryTemplate(fecha, assignedTurns, availableGuides, conflicts);
    
    const results = [];
    adminEmails.forEach(email => {
      results.push(Sherpas.MailSvc.send(email, subject, html));
    });
    return results;
  }

  // API Pública
  return {
    // Templates de construcción
    buildWelcomeTemplate: buildWelcomeTemplate,
    buildAssignmentTemplate: buildAssignmentTemplate,
    buildReleaseTemplate: buildReleaseTemplate,
    buildManagerAlertTemplate: buildManagerAlertTemplate,
    buildDailySummaryTemplate: buildDailySummaryTemplate,
    
    // Métodos de envío directo
    sendWelcome: sendWelcome,
    sendAssignment: sendAssignment,
    sendRelease: sendRelease,
    sendManagerAlert: sendManagerAlert,
    sendDailySummary: sendDailySummary,
    
    // Configuración
    EMAIL_CONFIG: EMAIL_CONFIG
  };
})();