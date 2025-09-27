/** 04_integraciones.gs — Servicios externos: Calendar, Mail, Triggers */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

/** Servicio de Calendar */
Sherpas.CalendarSvc = (function(){
  'use strict';

  function _shiftStart(iso, slot){
    var parts = String(iso).split('-');
    var y = parseInt(parts[0],10), m = parseInt(parts[1],10), d = parseInt(parts[2],10);
    var hm  = Sherpas.CFG.SHIFT_TIMES[slot]; // 'HH:MM'
    var hh = parseInt(hm.split(':')[0],10), mm = parseInt(hm.split(':')[1],10);
    return new Date(y, m-1, d, hh, mm, 0);
  }
  function _getEvent(cal, start, slot){
    var durMins = Sherpas.CFG.SHIFT_DUR_MINS[slot];
    var end = new Date(start.getTime() + durMins*60000);
    var evs = cal.getEvents(start, end);
    return evs && evs.length ? evs[0] : null;
  }

  /** Invita un email a un evento del slot indicado (M/T1/T2/T3) en la fecha ISO */
  function invite(iso, slot, email){
    try{
      if(!email) return;
      var cal = CalendarApp.getCalendarById(Sherpas.CFG.CALENDAR_ID);
      var start = _shiftStart(iso, slot);
      var ev = _getEvent(cal, start, slot);
      if(!ev) return;
      ev.addGuest(email);
    }catch(e){ console.error('CalendarSvc.invite', e); }
  }

  /** Quita un invitado del evento del slot indicado (M/T1/T2/T3) en la fecha ISO */
  function remove(iso, slot, email){
    try{
      if(!email) return;
      var cal = CalendarApp.getCalendarById(Sherpas.CFG.CALENDAR_ID);
      var start = _shiftStart(iso, slot);
      var ev = _getEvent(cal, start, slot);
      if(!ev) return;
      try{ ev.removeGuest(email); }catch(_){}
    }catch(e){ console.error('CalendarSvc.remove', e); }
  }

  return { invite:invite, remove:remove, _shiftStart:_shiftStart };
})();

/** Servicio de correo */
Sherpas.MailSvc = (function(){
  'use strict';
  function send(to, subject, html){
    try{
      if(!to) {
        console.warn('MailSvc: No se especificó destinatario');
        return false;
      }
      if(MailApp.getRemainingDailyQuota() <= 0) {
        console.warn('MailSvc: Quota de email agotada');
        return false;
      }
      MailApp.sendEmail({
        to: to, 
        subject: subject, 
        htmlBody: html, 
        name: 'Spain Food Sherpas'
      });
      console.log('Email enviado a:', to);
      return true;
    }catch(e){ 
      console.error('MailSvc.send error:', e); 
      return false;
    }
  }
  return { send:send };
})();

/** NUEVO: Servicio de Templates de Email */
Sherpas.EmailTemplates = (function(){
  'use strict';

  /** Configuración base de emails */
  var EMAIL_CONFIG = {
    FROM_NAME: 'Spain Food Sherpas',
    BRAND_COLOR: '#1a73e8',
    SUCCESS_COLOR: '#c6efce',
    WARNING_COLOR: '#fff3cd', 
    ERROR_COLOR: '#ffcccc',
    TEXT_COLOR: '#333333',
    FOOTER_COLOR: '#666666'
  };

  /** Template base HTML */
  function _buildBaseTemplate(headerText, content) {
    return `
      <div style="max-width:600px; margin:0 auto; font-family:Arial, sans-serif; color:${EMAIL_CONFIG.TEXT_COLOR};">
        <div style="background:${EMAIL_CONFIG.BRAND_COLOR}; color:white; padding:20px; text-align:center; border-radius:8px 8px 0 0;">
          <h1 style="margin:0; font-size:24px;">🍽️ Spain Food Sherpas</h1>
          <p style="margin:8px 0 0 0; opacity:0.9;">${headerText}</p>
        </div>
        <div style="background:#f8f9fa; padding:20px; margin:0;">
          ${content}
        </div>
        <div style="background:#f1f3f4; padding:15px; text-align:center; font-size:12px; color:${EMAIL_CONFIG.FOOTER_COLOR}; border-radius:0 0 8px 8px;">
          <p style="margin:0;">Este email fue generado automáticamente por el Sistema de Gestión de Tours</p>
          <p style="margin:5px 0 0 0;">© 2025 Spain Food Sherpas - Madrid</p>
        </div>
      </div>
    `;
  }

  /**
   * Template de bienvenida para nuevo guía
   */
  function buildWelcomeTemplate(nombreGuia, codigoGuia, enlaceCalendario) {
    var content = `
      <div style="text-align:center; margin-bottom:20px;">
        <h2 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 10px 0;">¡Bienvenido al equipo!</h2>
        <p style="font-size:16px; margin:0;">Tu calendario personal está listo para usar</p>
      </div>
      
      <div style="background:white; padding:20px; border-radius:6px; margin:20px 0;">
        <div style="margin-bottom:15px;">
          <h3 style="margin:0 0 5px 0; color:${EMAIL_CONFIG.TEXT_COLOR};">👤 ${nombreGuia}</h3>
          <p style="margin:0; color:${EMAIL_CONFIG.FOOTER_COLOR}; font-size:14px;">Código de guía: <strong>${codigoGuia}</strong></p>
        </div>
        
        <div style="background:${EMAIL_CONFIG.WARNING_COLOR}; padding:15px; border-radius:4px; margin:15px 0;">
          <p style="margin:0 0 10px 0; font-weight:bold;">📅 Accede a tu calendario:</p>
          <a href="${enlaceCalendario}" 
             style="display:inline-block; background:${EMAIL_CONFIG.BRAND_COLOR}; color:white; padding:10px 20px; text-decoration:none; border-radius:4px; font-weight:bold;">
            🚀 Abrir Mi Calendario
          </a>
        </div>
      </div>

      <div style="background:white; padding:20px; border-radius:6px;">
        <h3 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 15px 0;">📋 Instrucciones importantes:</h3>
        <ul style="margin:0; padding-left:20px; line-height:1.6;">
          <li>Marca <strong>"NO DISPONIBLE"</strong> en turnos que no puedes trabajar</li>
          <li>Los turnos asignados aparecerán en <span style="background:${EMAIL_CONFIG.SUCCESS_COLOR}; padding:2px 4px; border-radius:3px;">verde</span></li>
          <li><strong>No modifiques</strong> nada más en el calendario</li>
          <li>Las asignaciones las gestiona el manager desde el sistema central</li>
        </ul>
      </div>

      <div style="background:white; padding:15px; border-radius:6px; margin-top:20px; text-align:center;">
        <p style="margin:0; color:${EMAIL_CONFIG.FOOTER_COLOR}; font-size:14px;">
          ¿Problemas con el calendario? Contacta con el manager
        </p>
      </div>
    `;

    return _buildBaseTemplate('Sistema de Gestión de Tours', content);
  }

  /**
   * Envía email de bienvenida a nuevo guía
   */
  function sendWelcome(nombreGuia, codigoGuia, email, enlaceCalendario) {
    try {
      var subject = `🎉 Tu calendario está listo - ${nombreGuia} (${codigoGuia})`;
      var html = buildWelcomeTemplate(nombreGuia, codigoGuia, enlaceCalendario);
      
      var success = Sherpas.MailSvc.send(email, subject, html);
      
      if(success) {
        console.log(`Email de bienvenida enviado a ${nombreGuia} (${email})`);
        return true;
      } else {
        console.warn(`Error enviando email de bienvenida a ${nombreGuia}`);
        return false;
      }
    } catch(e) {
      console.error('Error en sendWelcome:', e);
      return false;
    }
  }

  /**
   * Template para notificación de asignación
   */
  function sendAssignment(nombreGuia, codigoGuia, email, fecha, turno, enlaceCalendario) {
    var content = `
      <div style="text-align:center; margin-bottom:20px;">
        <h2 style="color:${EMAIL_CONFIG.BRAND_COLOR}; margin:0 0 10px 0;">📅 Nueva Asignación</h2>
        <p style="font-size:16px; margin:0;">Tienes un turno asignado</p>
      </div>
      
      <div style="background:white; padding:20px; border-radius:6px; margin:20px 0;">
        <h3 style="color:${EMAIL_CONFIG.TEXT_COLOR}; margin:0 0 15px 0;">Detalles del turno:</h3>
        <p style="margin:0 0 10px 0;"><strong>Guía:</strong> ${nombreGuia} (${codigoGuia})</p>
        <p style="margin:0 0 10px 0;"><strong>Fecha:</strong> ${fecha}</p>
        <p style="margin:0 0 20px 0;"><strong>Turno:</strong> ${turno}</p>
        
        <a href="${enlaceCalendario}" 
           style="display:inline-block; background:${EMAIL_CONFIG.BRAND_COLOR}; color:white; padding:10px 20px; text-decoration:none; border-radius:4px; font-weight:bold;">
          📋 Ver Calendario
        </a>
      </div>
    `;

    var subject = `📅 Turno asignado: ${fecha} - ${turno}`;
    var html = _buildBaseTemplate('Asignación de Turno', content);
    
    return Sherpas.MailSvc.send(email, subject, html);
  }

  return {
    buildWelcomeTemplate: buildWelcomeTemplate,
    sendWelcome: sendWelcome,
    sendAssignment: sendAssignment
  };
})();

/** Servicio de triggers instalables */
Sherpas.TriggerSvc = (function(){
  'use strict';

  /** Crea un trigger CLOCK cada 5 minutos si no existe para la función dada */
  function ensureTimeEvery5m(fnName){
    var ok = ScriptApp.getProjectTriggers().some(function(t){
      return t.getHandlerFunction()===fnName && t.getTriggerSource()===ScriptApp.TriggerSource.CLOCK;
    });
    if(!ok) ScriptApp.newTrigger(fnName).timeBased().everyMinutes(5).create();
  }

  /** Crea un onEdit instalable para un Spreadsheet concreto si no existe */
  function ensureOnEditForSpreadsheet(spreadsheetId, handlerName){
    var exists = ScriptApp.getProjectTriggers().some(function(t){
      return t.getHandlerFunction()===handlerName &&
             t.getTriggerSource()===ScriptApp.TriggerSource.SPREADSHEETS &&
             t.getTriggerSourceId && t.getTriggerSourceId()===spreadsheetId;
    });
    if(!exists){
      ScriptApp.newTrigger(handlerName).forSpreadsheet(spreadsheetId).onEdit().create();
      return true;
    }
    return false;
  }

  /** Elimina el onEdit instalable para un Spreadsheet concreto (si existe) */
  function deleteOnEditForSpreadsheet(spreadsheetId, handlerName){
    ScriptApp.getProjectTriggers().forEach(function(t){
      if(t.getHandlerFunction()===handlerName &&
         t.getTriggerSource()===ScriptApp.TriggerSource.SPREADSHEETS){
        var id = t.getTriggerSourceId && t.getTriggerSourceId();
        if(id===spreadsheetId) ScriptApp.deleteTrigger(t);
      }
    });
  }

  /** Limpia onEdit huérfanos de un handler dado si su sourceId no está en la lista válida */
  function cleanOnEditOrphans(handlerName, validIds){
    var valid = new Set(validIds||[]);
    ScriptApp.getProjectTriggers().forEach(function(t){
      if(t.getHandlerFunction()===handlerName &&
         t.getTriggerSource()===ScriptApp.TriggerSource.SPREADSHEETS){
        var id = t.getTriggerSourceId && t.getTriggerSourceId();
        if(!id || !valid.has(id)) ScriptApp.deleteTrigger(t);
      }
    });
  }

  /** NUEVO: Instalación masiva de triggers onChange para todas las hojas de guías */
  function ensureOnChangeForAllGuides() {
    var count = 0;
    try {
      var guides = Sherpas.RegistryRepo.list();
      guides.forEach(function(guide) {
        var exists = ScriptApp.getProjectTriggers().some(function(t){
          return t.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
                 t.getTriggerSourceId && t.getTriggerSourceId() === guide.fileId &&
                 t.getEventType() === ScriptApp.EventType.ON_CHANGE;
        });
        
        if(!exists) {
          ScriptApp.newTrigger('globalOnChangeHandler')
            .forSpreadsheet(guide.fileId)
            .onChange()
            .create();
          count++;
        }
      });
    } catch(e) {
      console.error('Error instalando triggers onChange:', e);
    }
    return count;
  }

  /** NUEVO: Instalación de trigger onChange maestro desde MASTER */
  function ensureMasterOnChangeForAllGuides() {
    try {
      var masterId = PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID);
      if(!masterId) return false;
      
      var exists = ScriptApp.getProjectTriggers().some(function(t){
        return t.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
               t.getTriggerSourceId && t.getTriggerSourceId() === masterId &&
               t.getEventType() === ScriptApp.EventType.ON_CHANGE;
      });
      
      if(!exists) {
        ScriptApp.newTrigger('masterOnChangeHandler')
          .forSpreadsheet(masterId)
          .onChange()
          .create();
        return true;
      }
      return false;
    } catch(e) {
      console.error('Error instalando trigger onChange maestro:', e);
      return false;
    }
  }

  /** NUEVO: Contar triggers activos por tipo */
  function countActiveTriggers() {
    var triggers = ScriptApp.getProjectTriggers();
    var counts = {
      onEdit: 0,
      onChange: 0,
      timeBased: 0,
      total: triggers.length
    };
    
    triggers.forEach(function(t) {
      if(t.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS) {
        if(t.getEventType() === ScriptApp.EventType.ON_EDIT) counts.onEdit++;
        if(t.getEventType() === ScriptApp.EventType.ON_CHANGE) counts.onChange++;
      } else if(t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
        counts.timeBased++;
      }
    });
    
    console.log('Triggers activos:', counts);
    return counts;
  }

  return {
    ensureTimeEvery5m: ensureTimeEvery5m,
    ensureOnEditForSpreadsheet: ensureOnEditForSpreadsheet,
    deleteOnEditForSpreadsheet: deleteOnEditForSpreadsheet,
    cleanOnEditOrphans: cleanOnEditOrphans,
    ensureOnChangeForAllGuides: ensureOnChangeForAllGuides,
    ensureMasterOnChangeForAllGuides: ensureMasterOnChangeForAllGuides,
    countActiveTriggers: countActiveTriggers
  };
})();