/** 10_triggers.gs — Triggers */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.Triggers = (function(){
  'use strict';

  function onEditMaster(e){
    try{
      var ss = SpreadsheetApp.getActive();
      var sh = e.range.getSheet(); if(!Sherpas.CFG.MASTER_MONTH_NAME.test(sh.getName())) return;
      var row = e.range.getRow(), col = e.range.getColumn(); if(row<3) return;

      // identificar bloque guía
      var blocks = (function _listAllGuideBlocks(sheet){
        var res=[], lastCol=sheet.getLastColumn();
        for(var c=1;c<=lastCol;c++){
          var m = String(sheet.getRange(2,c).getDisplayValue()||'').toUpperCase();
          var t = String(sheet.getRange(2,c+1).getDisplayValue()||'').toUpperCase();
          if((m==='MAÑANA'||m==='M') && (t==='TARDE'||t==='T')){ res.push({colM:c,colT:c+1}); c++; }
        }
        return res;
      })(sh);
      var blk = blocks.find(function(b){ return col===b.colM || col===b.colT; }); if(!blk) return;

      var fechaISO = Sherpas.Util.toISO(Sherpas.Util.fromDisplay(sh.getRange(row,1).getDisplayValue()));
      var isM = (col===blk.colM);
      var accion = String(e.value||'').toUpperCase();

      var validList = isM? Sherpas.CFG.MASTER_DV_M : Sherpas.CFG.MASTER_DV_T;
      if(validList.indexOf(accion)===-1){ e.range.setValue(e.oldValue||''); sh.toast('Valor no válido.'); return; }

      var header = sh.getRange(1, isM?blk.colM:blk.colT).getDisplayValue();
      var code = (header.split('—')[0]||header).trim().toUpperCase();
      var reg = Sherpas.RegistryRepo.list().find(function(r){ return r.codigo===code; });
      if(!reg){ sh.toast('Guía no en REGISTRO'); return; }

      var celda = sh.getRange(row, isM?blk.colM:blk.colT);
      var current = String(celda.getDisplayValue()||'').toUpperCase();
      if(current==='NO DISPONIBLE' && accion!=='LIBERAR'){ e.range.setValue(e.oldValue||''); sh.toast('Bloqueado por GUÍA.'); return; }

      if(accion==='LIBERAR'){
        // Determinar slot previo para quitar invitación si estaba asignado
        var slot = isM? 'M' : (current.indexOf('ASIGNADO')===0? current.split(/\s+/).pop() : 'T1');
        celda.setValue('');
        Sherpas.GuideBook.writeCell(reg.fileId, sh.getName(), fechaISO, isM?'MAÑANA':'TARDE', isM?'MAÑANA':'TARDE', true);
        Sherpas.CalendarSvc.remove(fechaISO, slot, reg.email);
        Sherpas.MailSvc.send(reg.email, 'Liberación: '+Sherpas.Util.dateES(fechaISO)+' '+code+' '+(isM?'M':'T'),
          '<p>Se liberó tu turno del '+Sherpas.Util.dateES(fechaISO)+' ('+(isM?'Mañana':'Tarde')+').</p><p><a href="'+reg.url+'">Abrir calendario</a></p>');
        return;
      }

      if(accion.indexOf('ASIGNAR')===0){
        var asignado = accion.replace('ASIGNAR','ASIGNADO').trim();
        celda.setValue(asignado);
        Sherpas.GuideBook.writeCell(reg.fileId, sh.getName(), fechaISO, isM?'MAÑANA':'TARDE', asignado, true);
        var slot2 = isM? 'M' : asignado.split(' ').pop();
        Sherpas.CalendarSvc.invite(fechaISO, slot2, reg.email);
        Sherpas.MailSvc.send(reg.email, 'Asignación: '+Sherpas.Util.dateES(fechaISO)+' '+code+' '+slot2,
          '<p>Tienes una asignación el '+Sherpas.Util.dateES(fechaISO)+' ('+slot2+').</p><p><a href="'+reg.url+'">Abrir calendario</a></p>');
        return;
      }
    }catch(err){ console.error('onEdit MASTER:', err); }
  }

  function onEditGuide(e){
    try{
      var range = e && e.range; if(!range) return;
      var sh = range.getSheet(); var ss = sh.getParent();
      if(!Sherpas.CFG.GUIDE_MONTH_NAME.test(sh.getName())) return;

      var r = range.getRow(), c = range.getColumn();
      var off = (r - 2) % 3; if(off!==1 && off!==2) return; // 1=MAÑANA 2=TARDE
      var turno = (off===1)? 'MAÑANA':'TARDE';

      var newVal = (e.value==null? '' : String(e.value).trim().toUpperCase());
      var oldVal = (e.oldValue==null? sh.getRange(r,c).getDisplayValue() : String(e.oldValue).trim().toUpperCase());

      // Fecha
      var num = parseInt(sh.getRange(r-off, c).getDisplayValue(),10); if(!num) return;
      var p = Sherpas.Util.parseTab_MMYYYY(sh.getName()); var yyyy=p.yyyy, mm=p.mm;
      var fecha = new Date(yyyy, mm-1, num);

      // Ventana 14h para NO DISPONIBLE/REVERTIR/"" 
      if(newVal==='NO DISPONIBLE' || newVal==='REVERTIR' || newVal===''){
        var slotRef = (turno==='MAÑANA')? 'M' : 'T1';
        var start = Sherpas.CalendarSvc._shiftStart(Sherpas.Util.toISO(fecha), slotRef);
        var diffH = (start.getTime() - (new Date()).getTime())/3600000;
        if(diffH < Sherpas.CFG.LIMIT_HOURS){ range.setValue(oldVal); ss.toast('Fuera de ventana (14h).'); return; }
      }

      // MASTER y jerarquía
      var master = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(Sherpas.KEYS.MASTER_ID));
      var ms = master.getSheetByName(sh.getName()); if(!ms){ ss.toast('MASTER sin pestaña '+sh.getName()); return; }
      var header = ss.getName(); // CALENDARIO_{NOMBRE}-GXX
      var mTitle = String(header||'').match(/-(G\d{2})$/i);
      var codigo = mTitle? mTitle[1].toUpperCase() : 'G00';

      // localizar columnas del guía en MASTER
      var cols = (function _findGuideBlockCols(sheet, codigo){
        var lastCol = sheet.getLastColumn();
        var top = sheet.getRange(1,1,1,lastCol).getDisplayValues()[0].map(function(s){ return String(s||'').toUpperCase(); });
        var row2= sheet.getRange(2,1,1,lastCol).getDisplayValues()[0].map(function(s){ return String(s||'').toUpperCase(); });
        var codeUp = codigo.toUpperCase();
        for(var c2=1;c2<=lastCol;c2++){
          var t = top[c2-1];
          if(t.startsWith(codeUp+' ') || t.startsWith(codeUp+'—') || t===codeUp){
            var isM=(row2[c2-1]==='MAÑANA'||row2[c2-1]==='M'), isT=(row2[c2]==='TARDE'||row2[c2]==='T');
            if(isM && isT) return {colM:c2, colT:c2+1};
          }
        }
        return {colM:0,colT:0};
      })(ms, codigo);

      // fila de la fecha en MASTER
      var rowMaster = Sherpas.MasterBook.findRowByISO(ms, Sherpas.Util.toISO(fecha));
      if(!cols.colM || !cols.colT || rowMaster<3){ ss.toast('MASTER preparándose...'); return; }

      var targetCol = (turno==='MAÑANA')? cols.colM : cols.colT;
      var cell = ms.getRange(rowMaster, targetCol);
      var currentMaster = String(cell.getDisplayValue()||'').toUpperCase();

      if(currentMaster.indexOf('ASIGNADO')===0){ range.setValue(turno); ss.toast('Turno asignado por MASTER.'); return; }
      if(newVal==='NO DISPONIBLE'){ cell.setValue('NO DISPONIBLE'); return; }
      if(newVal==='REVERTIR' || newVal===''){ range.setValue(turno); cell.setValue(''); return; }
    }catch(err){ console.error('onEdit GUIDE:', err); }
  }

  function cronReconcile(){ Sherpas.UseCases.CronReconcileUC(); }

  return {
    onEditMaster: onEditMaster,
    onEditGuide: onEditGuide,
    cronReconcile: cronReconcile
  };
})();

/* Puntos de entrada simples */
function onOpen(){ Sherpas.UI.onOpen(); }                 // menú
function onEdit(e){ Sherpas.Triggers.onEditMaster(e); }   // simple trigger en MASTER
