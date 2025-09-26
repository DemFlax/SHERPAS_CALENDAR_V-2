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
    
    // Debug info
    if(typeof Sherpas.TriggerSvc.countActiveTriggers === 'function') {
      setTimeout(function() {
        Sherpas.TriggerSvc.countActiveTriggers();
      }, 1000);
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

  /** Repara TODAS las guías: recorta filas, re-aplica DV/CF y re-protege */
  function RepairGuidesUC(){
    Sherpas.RegistryRepo.list().forEach(function(r){
      try {
        var ss = SpreadsheetApp.openById(r.fileId);
        ss.getSheets().forEach(function(sh){
          if(Sherpas.CFG.GUIDE_MONTH_NAME.test(sh.getName())){
            Sherpas.GuideBook.normalize(sh);
            Sherpas.GuideBook.applyDV(sh);
            Sherpas.GuideBook.applyCF(sh);
            Sherpas.GuideBook.protectEditableMT(sh);
          }
        });
      } catch(e) {
        console.error('Error reparando guía ' + r.codigo + ':', e);
      }
    });
    return true;
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

  return {
    InitUC: InitUC,
    CreateGuideUC: CreateGuideUC,
    AdoptGuideUC: AdoptGuideUC,
    DeleteGuideTotalUC: DeleteGuideTotalUC,
    SyncNowUC: SyncNowUC,
    RepairGuidesUC: RepairGuidesUC,
    CronReconcileUC: CronReconcileUC
  };
})();