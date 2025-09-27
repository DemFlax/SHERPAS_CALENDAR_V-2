/** 09_menu.gs — Menú con Sistema de Protección Automática */
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
      .addSeparator()
      .addItem('🛡️ ACTIVAR Protección Automática','Sherpas.UI.ActivarProteccionAuto') // NUEVO
      .addItem('🧪 Test Email Sistema','Sherpas.UI.TestEmail')
      .addItem('🔧 Reparar Guías (normal)','Sherpas.UI.RepararGuias')
      .addItem('🚨 REPARACIÓN FORZADA','Sherpas.UI.ReparacionForzada')
      .addSeparator()
      .addItem('Reconstruir REGISTRO desde carpeta','Sherpas.UI.RebuildRegistro')
      .addItem('Eliminar Guía (TOTAL)','Sherpas.UI.EliminarGuiaTotal')
      .addSeparator()
      .addItem('🧹 Limpiar Sistema','Sherpas.UI.LimpiarSistema')
      .addToUi();
  }

  function Inicializar(){ 
    try {
      Sherpas.UseCases.InitUC(); 
    } catch(e) {
      SpreadsheetApp.getUi().alert('Error en inicialización: ' + e.message);
    }
  }

  function CrearGuia(){
    var ui = SpreadsheetApp.getUi();
    var nombre = ui.prompt('Nombre del guía').getResponseText().trim();
    if(!nombre) return;
    
    var email = ui.prompt('Email del guía').getResponseText().trim().toLowerCase();
    if(!email) return;
    
    var codigo = ui.prompt('Código del guía (G01)').getResponseText().trim().toUpperCase();
    if(!codigo) return;
    
    try{
      var file = Sherpas.UseCases.CreateGuideUC(nombre, email, codigo);
      ui.alert('✅ Guía creado correctamente: ' + file.getName() + '\n\n📧 Se ha enviado un email de bienvenida a: ' + email);
      
      // Auto-instalar protección en nuevo guía
      if(typeof Sherpas.AutoProtection !== 'undefined') {
        try {
          Sherpas.TriggerSvc.ensureOnEditForSpreadsheet(file.getId(), 'autoProtectionTrigger');
          ui.alert('🛡️ Protección automática instalada en el nuevo calendario');
        } catch(e) {
          console.warn('Error instalando protección automática:', e);
        }
      }
      
    }catch(e){ 
      ui.alert('❌ Error creando guía: ' + String(e)); 
    }
  }

  function AdoptarGuia(){
    var key = SpreadsheetApp.getUi().prompt('Pegue URL/ID/Código/Email').getResponseText().trim();
    if(!key) return;
    try{ 
      Sherpas.UseCases.AdoptGuideUC(key); 
      SpreadsheetApp.getUi().alert('✅ Guía adoptado correctamente.'); 
    }
    catch(e){ 
      SpreadsheetApp.getUi().alert('❌ Error adoptando guía: ' + String(e)); 
    }
  }

  function SyncAhora(){
    var master = SpreadsheetApp.getActive();
    var sh = master.getActiveSheet();
    if(!Sherpas.CFG.MASTER_MONTH_NAME.test(sh.getName())){ 
      master.toast('❌ Activa una pestaña MM_YYYY (ej: 10_2025)', 'Error', 5); 
      return; 
    }
    try {
      Sherpas.UseCases.SyncNowUC(sh.getName());
      master.toast('✅ Sincronización completa para ' + sh.getName(), 'Éxito', 3);
    } catch(e) {
      master.toast('❌ Error en sincronización: ' + e.message, 'Error', 10);
    }
  }

  /** NUEVA: Activar Sistema de Protección Automática */
  function ActivarProteccionAuto(){
    var ui = SpreadsheetApp.getUi();
    
    // Verificar si ya está activo
    var isActive = false;
    try {
      isActive = typeof Sherpas.AutoProtection !== 'undefined' && Sherpas.AutoProtection.isProtectionActive();
    } catch(e) {
      isActive = false;
    }
    
    if(isActive) {
      var response = ui.alert(
        '🛡️ Protección Automática',
        '✅ El sistema de protección automática YA ESTÁ ACTIVO\n\n¿Quieres reinstalarlo en todos los calendarios?',
        ui.ButtonSet.YES_NO
      );
      
      if(response === ui.Button.NO) return;
    } else {
      var response = ui.alert(
        '🛡️ Activar Protección Automática',
        '¿Activar protección automática en tiempo real?\n\nEsto hará que:\n\n🔒 Cualquier valor inválido se revierta automáticamente\n⚡ NO más "PUTA" o alteraciones\n🎯 Sincronización automática con MASTER\n🎨 Formato automático (rojo/verde)\n📱 Mensajes informativos al usuario\n\n⭐ RECOMENDADO PARA PROTECCIÓN COMPLETA ⭐',
        ui.ButtonSet.YES_NO
      );
      
      if(response === ui.Button.NO) return;
    }
    
    try {
      // Ejecutar inicialización del sistema
      var result = initializeAutoProtectionSystem();
      
      if(result.success) {
        ui.alert(
          '✅ PROTECCIÓN AUTOMÁTICA ACTIVADA',
          '🛡️ Sistema de protección en tiempo real ACTIVO\n\n📊 Resultados:\n• Calendarios protegidos: ' + result.installed + '\n• Estado: ACTIVO\n\n🔒 Beneficios:\n• Rollback automático de valores inválidos\n• Sincronización instantánea\n• Formato automático\n• Protección 24/7\n\n⚡ Los calendarios están ahora completamente protegidos'
        );
      } else {
        ui.alert('❌ Error activando protección automática:\n\n' + result.error);
      }
      
    } catch(e) {
      ui.alert('❌ Error crítico activando protección:\n\n' + e.message);
    }
  }

  /** Función para probar el sistema de emails */
  function TestEmail(){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '🧪 Test Email Sistema',
      '¿Enviar email de prueba a tu dirección? Esto verificará si el sistema de emails funciona correctamente.',
      ui.ButtonSet.YES_NO
    );
    
    if(response === ui.Button.YES) {
      try {
        var success = Sherpas.UseCases.TestEmailUC();
        if(success) {
          ui.alert('✅ Email de prueba enviado correctamente. Revisa tu bandeja de entrada.');
        } else {
          ui.alert('❌ Error enviando email de prueba. Verifica:\n\n1. Quota de Gmail disponible\n2. Permisos de script\n3. Configuración de email');
        }
      } catch(e) {
        ui.alert('❌ Error ejecutando test: ' + e.message);
      }
    }
  }

  /** Función para reparar todas las guías (método estándar) */
  function RepararGuias(){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '🔧 Reparar Guías (Normal)',
      '¿Reparar todas las hojas de guías con método estándar? Esto:\n\n• Normaliza filas excesivas\n• Reaplica validaciones de datos\n• Restaura protecciones\n• Corrige formato condicional',
      ui.ButtonSet.YES_NO
    );
    
    if(response === ui.Button.YES) {
      try {
        var result = Sherpas.UseCases.RepairGuidesUC();
        if(typeof result === 'object' && result.guidesProcessed) {
          ui.alert('✅ Reparación completa:\n\n• Guías: ' + result.guidesProcessed + '\n• Hojas reparadas: ' + result.sheetsRepaired + '\n• Errores: ' + result.errors);
        } else {
          ui.alert('✅ Reparación completa. Todas las guías han sido restauradas.');
        }
      } catch(e) {
        ui.alert('❌ Error reparando guías: ' + e.message);
      }
    }
  }

  /** Reparación forzada para problemas graves de protección */
  function ReparacionForzada(){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '🚨 REPARACIÓN FORZADA',
      '⚠️ Esta función resuelve problemas GRAVES de protección:\n\n🔹 Elimina TODAS las validaciones existentes\n🔹 Corrige valores inválidos (ej: "PUTA" → "MAÑANA")\n🔹 Reinstala protecciones desde cero\n🔹 Fuerza formato condicional correcto\n🔹 Reinstala triggers\n\n❗ Usar solo si hay problemas críticos ❗\n\n¿Continuar con reparación forzada?',
      ui.ButtonSet.YES_NO
    );
    
    if(response === ui.Button.YES) {
      var confirmacion = ui.alert(
        '🚨 CONFIRMACIÓN FINAL',
        '⚠️ ÚLTIMA ADVERTENCIA ⚠️\n\nEsta reparación forzada modificará TODOS los calendarios de guías existentes. Los cambios son irreversibles.\n\n¿CONFIRMAS que quieres proceder?',
        ui.ButtonSet.YES_NO
      );
      
      if(confirmacion === ui.Button.YES) {
        try {
          // Mostrar progreso
          SpreadsheetApp.getActive().toast('🚨 Iniciando reparación forzada... Por favor espera', 'Reparando', 5);
          
          var result = Sherpas.UseCases.ForceRepairAllGuidesUC();
          
          var message = '🚨 REPARACIÓN FORZADA COMPLETADA\n\n';
          message += '📊 Resultados:\n';
          message += '• Guías procesadas: ' + result.guidesProcessed + '\n';
          message += '• Hojas reparadas: ' + result.sheetsRepaired + '\n';
          message += '• Errores: ' + result.errors + '\n\n';
          
          if(result.errors === 0) {
            message += '✅ TODOS los problemas de protección resueltos\n\n';
            message += '🔹 Validaciones de datos: RESTABLECIDAS\n';
            message += '🔹 Contenido inválido: CORREGIDO\n';
            message += '🔹 Protecciones: REINSTALADAS\n';
            message += '🔹 Formato condicional: APLICADO\n';
            message += '🔹 Triggers: ACTUALIZADOS\n\n';
            message += '💡 RECOMENDACIÓN: Activa "Protección Automática" para prevenir futuros problemas';
          } else {
            message += '⚠️ Revisa la consola del script para detalles de errores';
          }
          
          ui.alert(message);
          
        } catch(e) {
          ui.alert('❌ Error crítico en reparación forzada:\n\n' + e.message + '\n\nRevisa la consola del script para más detalles.');
        }
      }
    }
  }

  function RebuildRegistro(){ 
    try {
      Sherpas.RegistryRepo.rebuildFromFolder(); 
      SpreadsheetApp.getUi().alert('✅ Registro reconstruido desde carpeta.');
    } catch(e) {
      SpreadsheetApp.getUi().alert('❌ Error reconstruyendo registro: ' + e.message);
    }
  }

  function EliminarGuiaTotal(){
    var ui = SpreadsheetApp.getUi();
    var key = ui.prompt('Código (GXX) o email del guía a ELIMINAR TOTALMENTE').getResponseText().trim();
    if(!key) return;
    
    var confirmation = ui.alert(
      '⚠️ ELIMINACIÓN TOTAL',
      '¿CONFIRMAS eliminar completamente el guía: ' + key + '?\n\n⚠️ ESTA ACCIÓN NO SE PUEDE DESHACER ⚠️\n\n• Se eliminará el archivo de calendario\n• Se quitarán las columnas del MASTER\n• Se eliminará del registro\n• Se limpiarán todos los triggers',
      ui.ButtonSet.YES_NO
    );
    
    if(confirmation === ui.Button.YES) {
      try{ 
        Sherpas.UseCases.DeleteGuideTotalUC(key); 
        ui.alert('✅ Guía eliminado completamente.'); 
      }
      catch(e){ 
        ui.alert('❌ Error eliminando guía: ' + String(e)); 
      }
    }
  }

  /** Función para limpiar sistema completo */
  function LimpiarSistema(){
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '🧹 Limpiar Sistema',
      '¿Ejecutar limpieza completa del sistema? Esto:\n\n• Limpia triggers huérfanos\n• Normaliza todas las hojas\n• Reaplica validaciones y protecciones\n• Regenera formato condicional\n• Verifica integridad de datos',
      ui.ButtonSet.YES_NO
    );
    
    if(response === ui.Button.YES) {
      try {
        // Ejecutar limpieza completa
        Sherpas.UseCases.CronReconcileUC();
        
        // Contar triggers después de limpieza
        if(typeof Sherpas.TriggerSvc.countActiveTriggers === 'function') {
          var counts = Sherpas.TriggerSvc.countActiveTriggers();
          ui.alert('✅ Limpieza completa terminada.\n\n📊 Triggers activos:\n• onEdit: ' + counts.onEdit + '\n• onChange: ' + counts.onChange + '\n• Temporales: ' + counts.timeBased + '\n• Total: ' + counts.total + '/20');
        } else {
          ui.alert('✅ Limpieza completa terminada.');
        }
      } catch(e) {
        ui.alert('❌ Error en limpieza: ' + e.message);
      }
    }
  }

  return {
    onOpen: onOpen,
    Inicializar: Inicializar,
    CrearGuia: CrearGuia,
    AdoptarGuia: AdoptarGuia,
    SyncAhora: SyncAhora,
    ActivarProteccionAuto: ActivarProteccionAuto, // NUEVA FUNCIÓN
    TestEmail: TestEmail,
    RepararGuias: RepararGuias,
    ReparacionForzada: ReparacionForzada,
    RebuildRegistro: RebuildRegistro,
    EliminarGuiaTotal: EliminarGuiaTotal,
    LimpiarSistema: LimpiarSistema
  };
})();