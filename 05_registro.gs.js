/** 05_registro.gs — Repositorio de REGISTRO */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.RegistryRepo = (function(){
  'use strict';

  function _SP(){ return PropertiesService.getScriptProperties(); }
  function _master(){
    var id = _SP().getProperty(Sherpas.KEYS.MASTER_ID);
    return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
  }
  function _sheet(){
    var ss = _master();
    var sh = ss.getSheetByName(Sherpas.CFG.REGISTRY_SHEET);
    if(!sh) sh = ss.insertSheet(Sherpas.CFG.REGISTRY_SHEET);
    if(sh.getLastRow()===0) sh.appendRow(Sherpas.CFG.REGISTRY_HEADERS);
    // formato fecha y protección total
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    // protege completamente la hoja (sin rangos desbloqueados)
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){
      if((p.getDescription()||'').indexOf('Sherpas — REGISTRO')===0) p.remove();
    });
    var p = sh.protect().setDescription('Sherpas — REGISTRO (solo script)');
    p.setWarningOnly(false);
    p.setUnprotectedRanges([]); // todo bloqueado
    return sh;
  }

  function ensure(){
    _sheet(); return true;
  }

  function listRaw(){
    var sh = _sheet();
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if(lr<2) return [];
    return sh.getRange(2,1,lr-1,lc).getValues();
  }

  function list(){
    return listRaw().map(function(r){
      return { timestamp:r[0], codigo:r[1], nombre:r[2], email:String(r[3]||'').toLowerCase(), fileId:r[4], url:r[5] };
    });
  }

  function upsert(meta){
    var sh = _sheet();
    var rows = listRaw();
    var byId = rows.findIndex(function(r){ return r[4]===meta.fileId; });
    var row = [new Date(), meta.codigo, meta.nombre, meta.email, meta.fileId, meta.url];
    if(byId>=0){
      sh.getRange(byId+2,1,1,row.length).setValues([row]);
    }else{
      sh.appendRow(row);
    }
  }

  function removeByFileId(fileId){
    var sh = _sheet();
    var rows = listRaw();
    for(var i=0;i<rows.length;i++){
      if(rows[i][4]===fileId){ sh.deleteRow(i+2); return true; }
    }
    return false;
  }

  function checkUniq(codigo, email, fileIdIfAdopt){
    var items = list();
    if(codigo && items.some(function(r){ return r.codigo===codigo && r.fileId!==fileIdIfAdopt; })) return false;
    if(email  && items.some(function(r){ return r.email===String(email).toLowerCase() && r.fileId!==fileIdIfAdopt; })) return false;
    return true;
  }

  function resolve(codeOrEmail){
    var up = String(codeOrEmail||'').toUpperCase();
    var lo = String(codeOrEmail||'').toLowerCase();
    var items = list();
    for(var i=0;i<items.length;i++){
      var r = items[i];
      if(r.codigo.toUpperCase()===up || r.email.toLowerCase()===lo) return r;
    }
    return null;
  }

  function rebuildFromFolder(){
    var folderId = _SP().getProperty(Sherpas.KEYS.DEST_FOLDER_ID) || Sherpas.CFG.DEST_FOLDER_ID_DEFAULT;
    var folder = DriveApp.getFolderById(folderId);
    var it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while(it.hasNext()){
      var f = it.next();
      var meta = _parseGuideFileName(f.getName());
      upsert({ codigo:meta.codigo, nombre:meta.nombre, email:'', fileId:f.getId(), url:f.getUrl() });
    }
  }

  // "CALENDARIO_{NOMBRE}-{GNN}"
  function _parseGuideFileName(title){
    var m = String(title||'').match(/^CALENDARIO_(.+?)-(G\d{2})$/i);
    if(m) return { nombre:m[1].trim(), codigo:m[2].toUpperCase() };
    return { nombre:title, codigo:'G00' };
  }

  return {
    ensure: ensure,
    list: list,
    upsert: upsert,
    removeByFileId: removeByFileId,
    checkUniq: checkUniq,
    resolve: resolve,
    rebuildFromFolder: rebuildFromFolder
  };
})();
