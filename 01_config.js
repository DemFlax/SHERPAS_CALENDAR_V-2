/** 01_config.gs — Configuración base y espacio de nombres
 *  Paradigma: OOP con un único namespace global `Sherpas`.
 *  Todos los demás archivos añaden clases u objetos bajo `Sherpas.*`.
 */

/** @namespace */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};

/** Versión del modelo de datos / scripts */
Sherpas.VERSION = 'v3.0.0-oop';

/** Constantes de configuración del proyecto */
Sherpas.CFG = Object.freeze({
  TIMEZONE: 'Europe/Madrid',

  // Carpeta destino por defecto para hojas de GUÍA
  DEST_FOLDER_ID_DEFAULT: '1AAHblDnbhyg-m58d0CaI6TXZvmNEUSw_',

  // Hoja REGISTRO
  REGISTRY_SHEET: 'REGISTRO',
  REGISTRY_HEADERS: ['TIMESTAMP','CODIGO','NOMBRE','EMAIL','FILE_ID','URL'],

  // Nombres de pestañas
  MASTER_MONTH_NAME: /^(\d{2})_(\d{4})$/, // MM_YYYY
  GUIDE_MONTH_NAME:  /^(\d{2})_(\d{4})$/, // MM_YYYY

  // Listas para validación de datos
  GUIDE_DV: ['', 'NO DISPONIBLE', 'REVERTIR'],
  MASTER_DV_M: ['', 'LIBERAR', 'ASIGNAR M'],
  MASTER_DV_T: ['', 'LIBERAR', 'ASIGNAR T1', 'ASIGNAR T2', 'ASIGNAR T3'],

  // Calendar Bookeo
  CALENDAR_ID: 'c_61981c641dc3c970e63f1713ccc2daa49d8fe8962b6ed9f2669c4554496c7bdd@group.calendar.google.com',

  // Turnos (Europe/Madrid)
  SHIFT_TIMES: { M:'12:00', T1:'17:15', T2:'18:15', T3:'19:15' },
  SHIFT_DUR_MINS: { M:180, T1:180, T2:180, T3:180 },

  // Reglas temporales
  LIMIT_HOURS: 14,          // GUÍA puede bloquear/liberar hasta 14h antes
  MAX_WAIT_LOCK_MS: 6000    // Espera de LockService para operaciones críticas
});

/** Claves en PropertiesService */
Sherpas.KEYS = Object.freeze({
  MASTER_ID: 'MASTER_ID',
  DEST_FOLDER_ID: 'DEST_FOLDER_ID'
});

/** Tipos de datos (JSDoc) */
/**
 * @typedef {Object} GuideMeta
 * @property {string} codigo    - Código guía "G01".
 * @property {string} nombre    - Nombre del guía.
 * @property {string} email     - Email del guía.
 * @property {string} fileId    - ID del Spreadsheet del guía.
 * @property {string} url       - URL del Spreadsheet del guía.
 */

/**
 * Helper: asegura TZ del proceso. Úsalo al inicio de casos de uso si necesitas consistencia.
 */
Sherpas.useProjectTZ = function () {
  Session.getScriptTimeZone && Session.getScriptTimeZone(); // no cambia TZ, pero documenta la dependencia
};

/** Utilidades mínimas comunes */
Sherpas.now = function () { return new Date(); };
Sherpas.pad2 = function (n) { return String(n).padStart(2, '0'); };
Sherpas.dateISO = function (d) {
  var dt = (d instanceof Date) ? d : new Date(d);
  return dt.getFullYear() + '-' + Sherpas.pad2(dt.getMonth()+1) + '-' + Sherpas.pad2(dt.getDate());
};
Sherpas.dateES = function (iso) {
  var a = String(iso).split('-'); return a[2] + '/' + a[1] + '/' + a[0];
};
