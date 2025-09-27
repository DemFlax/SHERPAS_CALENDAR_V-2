/** 08_casos_uso.gs — Casos de uso */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.UseCases = (function(){
  'use strict';
  function _SP(){ return PropertiesService.getScriptProperties(); }

  function InitUC(){
    var ss = SpreadsheetApp.getActive();
    if(!_SP().getProperty(Sherpas.KEYS.MASTER_ID)) _SP().setProperty(Sherpas.KEYS.MASTER_ID, ss.getId());
    if(!_SP().getProperty(Sherpas.KEYS.DEST_FOLDER_ID)) _SP().setProperty(Sherpas.KEYS.DEST_FOLDER_ID, Sherpas.CFG.DEST_FOLDER_ID_DEFAULT);
    
    // Instalar triggers básicos
    Sherpas.TriggerSvc.ensureTimeEvery5m('Sherpas.Triggers.cronReconcile');
    
    // CORRECCIÓN: Instalar trigger onChange para protecciones
    var onChangeInstalled = false;
    try {
      onChangeInstalled = Sherpas.TriggerSvc.ensureMasterOnChangeForAllGuides();
    } catch(e) {
      console.warn('Error instalando trigger onChange desde MASTER, intentando estrategia alternativa:', e);
      try {
        var alternativeCount = Sherpas.TriggerSvc.ensureOnChangeForAllGuides();
        onChangeInstalled = alternativeCount > 0;
      } catch(e2) {
        console.error('Error en estrategia alternativa:', e2);
      }
    }
    
    // Configurar sincronización automática si existe el módulo
    if(typeof Sherpas.SyncController !== 'undefined' && Sherpas.SyncController.setupAutoSync) {
      try {
        Sherpas.SyncController.setupAutoSync();
      } catch(e) {
        console.warn('Error configurando SyncController:', e);
      }
    }
    
    // Inicializar sistema base
    Sherpas.RegistryRepo.ensure();
    Sherpas.MasterBook.ensureMonthsFromOct();
    Sherpas.MasterBook.applyDVandCF_All();
    Sherpas.RegistryRepo.list().forEach(function(r){ 
      Sherpas.MasterBook.ensureGuideColumns(r.codigo, r.nombre); 
    });
    
    // Mensaje de confirmación
    var message = 'Sistema inicializado';
    if(onChangeInstalled) {
      message += ' con protecciones anti-alteración activadas';
    }
    ss.toast(message);
    
    // CORRECCIÓN: Ejecutar debug triggers directamente (sin setTimeout)
    if(typeof Sherpas.TriggerSvc !== 'undefined' && typeof Sherpas.TriggerSvc.countActiveTriggers === 'function') {
      try {
        Sherpas.TriggerSvc.countActiveTriggers();
      } catch(e) {
        console.warn('Error ejecutando debug triggers:', e);
      }
    }
  }

  function CreateGuideUC(nombre, email, codigo){
    if(!nombre||!email||!codigo) throw new Error('Datos incompletos');
    if(!Sherpas.RegistryRepo.checkUniq(codigo, email)) throw new Error('Código o email ya existen');
    var folderId = _SP().getProperty(Sherpas.KEYS.DEST_FOLDER_ID) || Sherpas.CFG.DEST_FOLDER_ID_DEFAULT;
    var file = SpreadsheetApp.create('CALENDARIO_'+nombre+'-'+codigo);
    DriveApp.getFileById(file.getId()).moveTo(DriveApp.getFolderById(folderId));
    var now = new Date(), y = now.getFullYear();
    var startM = Math.max(10, now.getMonth()+1);
    for(var m=startM; m<=12; m++){ 
      Sherpas.GuideBook.buildMonth(file.getId(), new Date(y, m-1, 1)); 
    }
    Sherpas.GuideBook.adoptGuide(file.getId());
    Sherpas.RegistryRepo.upsert({codigo:codigo, nombre:nombre, email:String(email).toLowerCase(), fileId:file.getId(), url:file.getUrl()});
    Sherpas.MasterBook.ensureMonthsFromOct();
    Sherpas.MasterBook.ensureGuideColumns(codigo, nombre);
    if(email) file.addEditor(email);
    
    // CORRECCIÓN: Actualizar triggers onChange para incluir nueva guía
    try {
      Sherpas.TriggerSvc.ensureOnEditForSpreadsheet(file.getId(), 'Sherpas.Triggers.onEditGuide');
      // Si usamos estrategia individual de onChange, añadir para esta guía
      var triggerCount = Sherpas.TriggerSvc.countActiveTriggers();
      if(triggerCount.onChange === 0) {
        // No hay triggers onChange, instalar uno para esta guía
        ScriptApp.newTrigger('globalOnChangeHandler')
          .forSpreadsheet(file.getId())
          .onChange()
          .create();
      }
    } catch(e) {
      console.warn('Error configurando triggers para nueva guía:', e);
    }
    
    // *** NUEVO: Enviar email de bienvenida con template profesional ***
    if(email && typeof Sherpas.EmailTemplates !== 'undefined') {
      try {
        var emailSent = Sherpas.EmailTemplates.sendWelcome(nombre, codigo, email, file.getUrl());
        if(emailSent) {
          console.log('Email de bienvenida enviado correctamente a:', email);
        } else {
          console.warn('Error enviando email de bienvenida a:', email);
          // Intentar notificar al usuario del problema
          SpreadsheetApp.getActive().toast('ADVERTENCIA: El calendario se creó correctamente, pero no se pudo enviar el email de bienvenida. Verifica la quota de Gmail.', 'Email no enviado', 10);
        }
      } catch(e) {
        console.error('Error crítico enviando email de bienvenida:', e);
        SpreadsheetApp.getActive().toast('ADVERTENCIA: Calendario creado, pero falló el envío de email. Error: ' + e.message, 'Error de Email', 10);
      }
    } else {
      console.warn('EmailTemplates no disponible o email vacío');
    }
    
    return file;
  }

  function AdoptGuideUC(key){
    var id = (function resolve(k){
      var m = k.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/); if(m) return m[1];
      var reg = Sherpas.RegistryRepo.resolve(k); if(reg) return reg.fileId;
      if(/^[a-zA-Z0-9\-_]{30,}$/.test(k)) return k;
      return null;
    })(key);
    if(!id) throw new Error('No se pudo resolver ese dato a ID');
    var ss = SpreadsheetApp.openById(id);
    var meta = (function parseTitle(title){
      var m = String(title||'').match(/^CALENDARIO_(.+?)-(G\d{2})$/i);
      return m? { nombre:m[1].trim(), codigo:m[2].toUpperCase() } : { nombre:title, codigo:'G00' };
    })(ss.getName());

    if(!Sherpas.RegistryRepo.checkUniq(meta.codigo, null, id)) throw new Error('Conflicto en REGISTRO');
    Sherpas.GuideBook.adoptGuide(id);
    Sherpas.RegistryRepo.upsert({codigo:meta.codigo, nombre:meta.nombre, email:'', fileId:id, url:ss.getUrl()});
    Sherpas.MasterBook.ensureMonthsFromOct();
    Sherpas.MasterBook.ensureGuideColumns(meta.codigo, meta.nombre);
    
    // CORRECCIÓN: Configurar triggers para guía adoptado
    try {
      Sherpas.TriggerSvc.ensureOnEditForSpreadsheet(id, 'Sherpas.Triggers.onEditGuide');
    } catch(e) {
      console.warn('Error configurando trigger onEdit para guía adoptado:', e);
    }
    
    return true;
  }

  function DeleteGuideTotalUC(codeOrEmail){
    var reg = Sherpas.RegistryRepo.resolve(codeOrEmail); 
    if(!reg) throw new Error('Guía no encontrado');
    
    Sherpas.MasterBook.removeGuideColumnsEverywhere(reg.codigo);
    Sherpas.TriggerSvc.deleteOnEditForSpreadsheet(reg.fileId, 'Sherpas.Triggers.onEditGuide');
    
    // CORRECCIÓN: También eliminar triggers onChange si existen
    try {
      var triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(function(t) {
        if(t.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
           t.getTriggerSourceId && t.getTriggerSourceId() === reg.fileId) {
          ScriptApp.deleteTrigger(t);
        }
      });
    } catch(e) {
      console.warn('Error eliminando triggers del guía:', e);
    }
    
    try{ DriveApp.getFileById(reg.fileId).setTrashed(true); }catch(_){}
    Sherpas.RegistryRepo.removeByFileId(reg.fileId);
    Sherpas.MasterBook.applyDVandCF_All();
    Sherpas.TriggerSvc.cleanOnEditOrphans('Sherpas.Triggers.onEditGuide', Sherpas.RegistryRepo.list().map(function(r){return r.fileId;}));
    return true;
  }

  function SyncNowUC(activeMonthTabName){
    Sherpas.MasterBook.syncFromGuidesForMonth(activeMonthTabName);
    
    // Usar sincronización bidireccional si está disponible
    if(typeof Sherpas.SyncController !== 'undefined' && Sherpas.SyncController.syncAllGuides) {
      try {
        var result = Sherpas.SyncController.syncAllGuides();
        console.log('Sincronización bidireccional completada:', result);
      } catch(e) {
        console.warn('Error en sincronización bidireccional, usando método tradicional:', e);
      }
    }
    
    return true;
  }

  /** NUEVA: Reparación completa de TODAS las guías con problemas de protección */
  function ForceRepairAllGuidesUC(){
    var guias = Sherpas.RegistryRepo.list();
    var totalRepaired = 0;
    var totalErrors = 0;
    var repairLog = [];

    console.log('Iniciando reparación forzada de', guias.length, 'guías...');

    guias.forEach(function(guia) {
      try {
        console.log('Reparando guía:', guia.codigo, '-', guia.nombre);
        
        var ss = SpreadsheetApp.openById(guia.fileId);
        var sheetsRepaired = 0;
        var sheetsWithErrors = 0;

        ss.getSheets().forEach(function(sheet) {
          if(Sherpas.CFG.GUIDE_MONTH_NAME.test(sheet.getName())) {
            try {
              // Reparación completa paso a paso
              console.log('  Reparando hoja:', sheet.getName());

              // 1. Normalizar estructura
              Sherpas.GuideBook.normalize(sheet);

              // 2. Limpiar contenido inválido
              _cleanInvalidContent(sheet);

              // 3. Aplicar protecciones completas 
              var p = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
              if(p) {
                var meta = Sherpas.Util.monthMeta(p.yyyy, p.mm);
                var mtA1 = Sherpas.Util.monthMT_A1_FromMeta(meta, 2);
                
                // Usar función de protección forzada
                if(typeof Sherpas.GuideBook.forceRepairGuide === 'function') {
                  Sherpas.GuideBook.applyDV(sheet);
                  Sherpas.GuideBook.applyCF(sheet);
                  Sherpas.GuideBook.protectEditableMT(sheet);
                } else {
                  // Fallback a métodos individuales
                  Sherpas.GuideBook.applyDV(sheet);
                  Sherpas.GuideBook.applyCF(sheet);
                  Sherpas.GuideBook.protectEditableMT(sheet);
                }
              }

              sheetsRepaired++;
              console.log('    ✅ Reparado:', sheet.getName());

            } catch(sheetError) {
              sheetsWithErrors++;
              console.error('    ❌ Error en hoja', sheet.getName(), ':', sheetError);
              repairLog.push({
                guia: guia.codigo,
                sheet: sheet.getName(),
                error: sheetError.message
              });
            }
          }
        });

        // Reinstalar triggers
        try {
          Sherpas.TriggerSvc.ensureOnEditForSpreadsheet(guia.fileId, 'Sherpas.Triggers.onEditGuide');
          console.log('  ✅ Triggers reinstalados para:', guia.codigo);
        } catch(triggerError) {
          console.warn('  ⚠️ Error reinstalando triggers:', triggerError);
        }

        totalRepaired += sheetsRepaired;
        totalErrors += sheetsWithErrors;

        repairLog.push({
          guia: guia.codigo,
          sheetsRepaired: sheetsRepaired,
          sheetsWithErrors: sheetsWithErrors,
          status: sheetsWithErrors === 0 ? 'COMPLETADO' : 'CON_ERRORES'
        });

      } catch(guideError) {
        totalErrors++;
        console.error('❌ Error completo en guía', guia.codigo, ':', guideError);
        repairLog.push({
          guia: guia.codigo,
          error: guideError.message,
          status: 'FALLIDO'
        });
      }
    });

    // Mostrar resumen
    console.log('=== RESUMEN REPARACIÓN COMPLETA ===');
    console.log('Guías procesadas:', guias.length);
    console.log('Hojas reparadas:', totalRepaired);
    console.log('Errores encontrados:', totalErrors);
    console.log('Log completo:', repairLog);

    // Mensaje al usuario
    var message = '🔧 Reparación Completa Finalizada\n\n';
    message += '📊 Resultados:\n';
    message += '• Guías procesadas: ' + guias.length + '\n';
    message += '• Hojas reparadas: ' + totalRepaired + '\n';
    message += '• Errores: ' + totalErrors + '\n\n';
    
    if(totalErrors === 0) {
      message += '✅ Todos los calendarios reparados exitosamente';
    } else {
      message += '⚠️ Revisa la consola para detalles de errores';
    }

    SpreadsheetApp.getActive().toast(message, 'Reparación Completa', 15);

    return {
      guidesProcessed: guias.length,
      sheetsRepaired: totalRepaired,
      errors: totalErrors,
      log: repairLog
    };
  }

  /**
   * NUEVA: Limpia contenido inválido de una hoja de guía
   */
  function _cleanInvalidContent(sheet) {
    try {
      var p = Sherpas.Util.parseTab_MMYYYY(sheet.getName());
      if(!p) return;

      var meta = Sherpas.Util.monthMeta(p.yyyy, p.mm);
      var mtA1List = Sherpas.Util.monthMT_A1_FromMeta(meta, 2);
      var cleaned = 0;

      mtA1List.forEach(function(a1) {
        var range = sheet.getRange(a1);
        var value = String(range.getDisplayValue() || '').toUpperCase().trim();
        
        // Si el valor no es válido, corregirlo
        if(value && !Sherpas.CFG.GUIDE_DV.includes(value) && !value.startsWith('ASIGNADO')) {
          var pos = Sherpas.Util.a1ToRowCol(a1);
          var rowType = (pos.row - 2) % 3;
          var correctValue = (rowType === 1) ? 'MAÑANA' : 'TARDE';
          
          range.setValue(correctValue);
          cleaned++;
          console.log('    Limpiado valor inválido en', a1, ':', value, '→', correctValue);
        }
      });

      if(cleaned > 0) {
        console.log('    ✅ Limpiados', cleaned, 'valores inválidos en', sheet.getName());
      }

    } catch(e) {
      console.error('Error limpiando contenido inválido:', e);
    }
  }

  /** Repara TODAS las guías: recorta filas, re-aplica DV/CF y re-protege */
  function RepairGuidesUC(){
    return ForceRepairAllGuidesUC(); // Usar la nueva función mejorada
  }

  function CronReconcileUC(){
    try{
      Sherpas.MasterBook.ensureMonthsFromOct();
      Sherpas.MasterBook.applyDVandCF_All();
      RepairGuidesUC(); // incluye normalize + DV + CF + protección
      Sherpas.TriggerSvc.cleanOnEditOrphans('Sherpas.Triggers.onEditGuide', Sherpas.RegistryRepo.list().map(function(r){return r.fileId;}));
      
      // Ejecutar mantenimiento de limpieza si está disponible
      if(typeof Sherpas.SyncController !== 'undefined' && Sherpas.SyncController.scheduledMaintenance) {
        Sherpas.SyncController.scheduledMaintenance();
      }
      
    }catch(e){ console.error('CronReconcileUC', e); }
  }

  /** NUEVA: Función para probar el sistema de emails */
  function TestEmailUC(){
    try {
      var testEmail = Session.getActiveUser().getEmail();
      if(!testEmail) {
        throw new Error('No se pudo obtener el email del usuario activo');
      }
      
      var success = Sherpas.EmailTemplates.sendWelcome(
        'Test Usuario', 
        'G99', 
        testEmail, 
        'https://docs.google.com/spreadsheets/d/test'
      );
      
      if(success) {
        SpreadsheetApp.getActive().toast('Email de prueba enviado correctamente a: ' + testEmail, 'Test Email', 5);
        return true;
      } else {
        SpreadsheetApp.getActive().toast('Error enviando email de prueba. Verifica la quota de Gmail.', 'Error Email', 10);
        return false;
      }
    } catch(e) {
      console.error('Error en TestEmailUC:', e);
      SpreadsheetApp.getActive().toast('Error en test de email: ' + e.message, 'Error', 10);
      return false;
    }
  }

  return {
    InitUC: InitUC,
    CreateGuideUC: CreateGuideUC,
    AdoptGuideUC: AdoptGuideUC,
    DeleteGuideTotalUC: DeleteGuideTotalUC,
    SyncNowUC: SyncNowUC,
    RepairGuidesUC: RepairGuidesUC,
    ForceRepairAllGuidesUC: ForceRepairAllGuidesUC,
    CronReconcileUC: CronReconcileUC,
    TestEmailUC: TestEmailUC
  };
})();