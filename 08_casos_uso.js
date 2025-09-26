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
    Sherpas.TriggerSvc.ensureTimeEvery5m('Sherpas.Triggers.cronReconcile');
    Sherpas.RegistryRepo.ensure();
    Sherpas.MasterBook.ensureMonthsFromOct();
    Sherpas.MasterBook.applyDVandCF_All();
    Sherpas.RegistryRepo.list().forEach(function(r){ Sherpas.MasterBook.ensureGuideColumns(r.codigo, r.nombre); });
    ss.toast('Inicializado');
  }

  function CreateGuideUC(nombre, email, codigo){
    if(!nombre||!email||!codigo) throw new Error('Datos incompletos');
    if(!Sherpas.RegistryRepo.checkUniq(codigo, email)) throw new Error('Código o email ya existen');
    var folderId = _SP().getProperty(Sherpas.KEYS.DEST_FOLDER_ID) || Sherpas.CFG.DEST_FOLDER_ID_DEFAULT;
    var file = SpreadsheetApp.create('CALENDARIO_'+nombre+'-'+codigo);
    DriveApp.getFileById(file.getId()).moveTo(DriveApp.getFolderById(folderId));

    // Meses: desde octubre hasta diciembre del año actual
    var now = new Date(), y = now.getFullYear();
    var startM = Math.max(10, now.getMonth()+1);
    for(var m=startM; m<=12; m++){ Sherpas.GuideBook.buildMonth(file.getId(), new Date(y, m-1, 1)); }

    // Adopción: DV/CF/Protección + trigger onEdit
    Sherpas.GuideBook.adoptGuide(file.getId());

    // Registro + MASTER + permisos + correo
    Sherpas.RegistryRepo.upsert({codigo:codigo, nombre:nombre, email:String(email).toLowerCase(), fileId:file.getId(), url:file.getUrl()});
    Sherpas.MasterBook.ensureMonthsFromOct();
    Sherpas.MasterBook.ensureGuideColumns(codigo, nombre);
    if(email) file.addEditor(email);
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
    return true;
  }

  function DeleteGuideTotalUC(codeOrEmail){
    var reg = Sherpas.RegistryRepo.resolve(codeOrEmail); if(!reg) throw new Error('Guía no encontrado');
    // MASTER: quitar columnas
    Sherpas.MasterBook.removeGuideColumnsEverywhere(reg.codigo);
    // Trigger instalable
    Sherpas.TriggerSvc.deleteOnEditForSpreadsheet(reg.fileId, 'Sherpas.Triggers.onEditGuide');
    // Borrar archivo y REGISTRO
    try{ DriveApp.getFileById(reg.fileId).setTrashed(true); }catch(_){}
    Sherpas.RegistryRepo.removeByFileId(reg.fileId);
    // Limpieza final
    Sherpas.MasterBook.applyDVandCF_All();
    Sherpas.TriggerSvc.cleanOnEditOrphans('Sherpas.Triggers.onEditGuide', Sherpas.RegistryRepo.list().map(function(r){return r.fileId;}));
    return true;
  }

  function SyncNowUC(activeMonthTabName){
    Sherpas.MasterBook.syncFromGuidesForMonth(activeMonthTabName);
    return true;
  }

  function CronReconcileUC(){
    try{
      Sherpas.MasterBook.ensureMonthsFromOct();
      Sherpas.MasterBook.applyDVandCF_All();
      // Reforzar guías (primera pestaña válida de cada guía)
      Sherpas.RegistryRepo.list().forEach(function(r){
        var ss = SpreadsheetApp.openById(r.fileId);
        var sh = ss.getSheets().find(function(s){ return Sherpas.CFG.GUIDE_MONTH_NAME.test(s.getName()); });
        if(sh){ Sherpas.GuideBook.applyDV(sh); Sherpas.GuideBook.applyCF(sh); Sherpas.GuideBook.protectEditableMT(sh); }
      });
      // Trigger huérfanos
      Sherpas.TriggerSvc.cleanOnEditOrphans('Sherpas.Triggers.onEditGuide', Sherpas.RegistryRepo.list().map(function(r){return r.fileId;}));
    }catch(e){ console.error('CronReconcileUC', e); }
  }

  return {
    InitUC: InitUC,
    CreateGuideUC: CreateGuideUC,
    AdoptGuideUC: AdoptGuideUC,
    DeleteGuideTotalUC: DeleteGuideTotalUC,
    SyncNowUC: SyncNowUC,
    CronReconcileUC: CronReconcileUC
  };
})();
