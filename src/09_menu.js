/** 09_menu.gs — Menú */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

Sherpas.UI = (function(){
  'use strict';

  function onOpen(){
    SpreadsheetApp.getUi().createMenu('Sherpas')
      .addItem('Inicializar (MASTER)','Sherpas.UI.Inicializar')
      .addSeparator()
      .addItem('Crear Guía (nuevo)','Sherpas.UI.CrearGuia')
      .addItem('Adoptar Guía (URL/ID/Código/Email)','Sherpas.UI.AdoptarGuia')
      .addSeparator()
      .addItem('Sincronizar ahora','Sherpas.UI.SyncAhora')
      .addItem('Reparar guías ahora','Sherpas.UI.RepararGuiasAhora')
      .addSeparator()
      .addItem('Limpiar sistema (eliminar botones filas)','Sherpas.UI.LimpiarSistema')
      .addItem('Reconstruir REGISTRO desde carpeta','Sherpas.UI.RebuildRegistro')
      .addItem('Eliminar Guía (TOTAL)','Sherpas.UI.EliminarGuiaTotal')
      .addToUi();
  }

  function Inicializar(){ Sherpas.UseCases.InitUC(); }
  function CrearGuia(){
    var ui = SpreadsheetApp.getUi();
    var nombre = ui.prompt('Nombre del guía').getResponseText().trim();
    var email  = ui.prompt('Email del guía').getResponseText().trim().toLowerCase();
    var codigo = ui.prompt('Código del guía (G01)').getResponseText().trim().toUpperCase();
    try{ var file = Sherpas.UseCases.CreateGuideUC(nombre,email,codigo); ui.alert('Guía creado: '+file.getName()); }
    catch(e){ ui.alert(String(e)); }
  }
  function AdoptarGuia(){
    var key = SpreadsheetApp.getUi().prompt('Pegue URL/ID/Código/Email').getResponseText().trim();
    if(!key) return; try{ Sherpas.UseCases.AdoptGuideUC(key); SpreadsheetApp.getUi().alert('Guía adoptado.'); }
    catch(e){ SpreadsheetApp.getUi().alert(String(e)); }
  }
  function SyncAhora(){
    var master = SpreadsheetApp.getActive();
    var sh = master.getActiveSheet();
    if(!Sherpas.CFG.MASTER_MONTH_NAME.test(sh.getName())){ master.toast('Activa una pestaña MM_YYYY'); return; }
    Sherpas.UseCases.SyncNowUC(sh.getName());
    master.toast('Sincronización completa');
  }
  function RepararGuiasAhora(){ Sherpas.UseCases.RepairGuidesUC(); SpreadsheetApp.getActive().toast('Guías reparadas'); }
  
  function LimpiarSistema(){
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert('Limpieza del Sistema', 
      '¿Ejecutar limpieza completa para eliminar botones "Añadir filas" innecesarios?', 
      ui.ButtonSet.YES_NO);
    
    if(result === ui.Button.YES) {
      try {
        // Usar función de limpieza del sync controller existente
        if(typeof Sherpas.SyncController !== 'undefined' && Sherpas.SyncController.cleanExcessRows) {
          var success = Sherpas.SyncController.cleanExcessRows();
          ui.alert('Limpieza Completada', 
            success ? 'Sistema limpiado exitosamente.' : 'Revisa los logs para más detalles.', 
            ui.ButtonSet.OK);
        } else {
          // Fallback: ejecutar reparación de guías que incluye normalización
          Sherpas.UseCases.RepairGuidesUC();
          ui.alert('Limpieza Completada', 'Se ejecutó reparación completa del sistema.', ui.ButtonSet.OK);
        }
      } catch(e) {
        ui.alert('Error en Limpieza', 'Error: ' + e.message, ui.ButtonSet.OK);
      }
    }
  }
  
  function RebuildRegistro(){ Sherpas.RegistryRepo.rebuildFromFolder(); }
  function EliminarGuiaTotal(){
    var ui = SpreadsheetApp.getUi();
    var key = ui.prompt('Código (GXX) o email del guía a ELIMINAR TOTALMENTE').getResponseText().trim();
    if(!key) return; try{ Sherpas.UseCases.DeleteGuideTotalUC(key); ui.alert('Eliminado correctamente.'); }
    catch(e){ ui.alert(String(e)); }
  }

  return {
    onOpen:onOpen,
    Inicializar:Inicializar,
    CrearGuia:CrearGuia,
    AdoptarGuia:AdoptarGuia,
    SyncAhora:SyncAhora,
    RepararGuiasAhora:RepararGuiasAhora,
    LimpiarSistema:LimpiarSistema,
    RebuildRegistro:RebuildRegistro,
    EliminarGuiaTotal:EliminarGuiaTotal
  };
})();