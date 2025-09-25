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
      if(!to) return;
      if(MailApp.getRemainingDailyQuota() <= 0) return;
      MailApp.sendEmail({to:to, subject:subject, htmlBody:html, name:'Spain Food Sherpas'});
    }catch(e){ console.error('MailSvc.send', e); }
  }
  return { send:send };
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

  return {
    ensureTimeEvery5m: ensureTimeEvery5m,
    ensureOnEditForSpreadsheet: ensureOnEditForSpreadsheet,
    deleteOnEditForSpreadsheet: deleteOnEditForSpreadsheet,
    cleanOnEditOrphans: cleanOnEditOrphans
  };
})();
