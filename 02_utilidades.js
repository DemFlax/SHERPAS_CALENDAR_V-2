/** 02_utilidades.gs — Utilidades puras (sin llamadas a Google)
 *  Todo se cuelga de `Sherpas.Util`.
 */
/* global Sherpas */
var Sherpas = (typeof Sherpas !== 'undefined') ? Sherpas : {};
Sherpas.Util = (function () {
  'use strict';

  /* -------- Básicas -------- */
  function pad2(n){ return String(n).padStart(2,'0'); }
  function toISO(d){
    var dt = (d instanceof Date)? d : new Date(d);
    return dt.getFullYear()+'-'+pad2(dt.getMonth()+1)+'-'+pad2(dt.getDate());
  }
  function fromDisplay(text){
    var t=String(text||'');
    var a=t.match(/^(\d{2})\/(\d{2})\/(\d{4})$/); if(a) return new Date(+a[3],+a[2]-1,+a[1]);
    var b=t.match(/^(\d{4})-(\d{2})-(\d{2})$/);    if(b) return new Date(+b[1],+b[2]-1,+b[3]);
    return new Date(t);
  }

  /* -------- A1 helpers -------- */
  function colToA1(col){ var s=''; while(col){ var n=(col-1)%26; s=String.fromCharCode(65+n)+s; col=(col-n-1)/26|0; } return s; }
  function a1ToRowCol(a1){
    var m=String(a1).toUpperCase().match(/^([A-Z]+)(\d+)$/); if(!m) return {row:0,col:0};
    var col=0, letters=m[1]; for(var i=0;i<letters.length;i++){ col = col*26 + (letters.charCodeAt(i)-64); }
    return {row:parseInt(m[2],10), col:col};
  }

  /* -------- Mes y cuadrícula (Lun=0..Dom=6) -------- */
  function firstDayOfMonth(y,m){ return new Date(y, m-1, 1); }
  function lastDayNumber(y,m){ return new Date(y, m, 0).getDate(); }
  function mondayIndex(d){ return (d.getDay()+6)%7; } // Lun=0
  function monthMeta(y,m){
    var first = firstDayOfMonth(y,m);
    var lastN = lastDayNumber(y,m);
    var firstW = mondayIndex(first);
    var weeks = Math.ceil((firstW + lastN)/7);
    return { yyyy:y, mm:m, firstWeekday:firstW, lastDay:lastN, weeks:weeks };
  }

  /**
   * Construye modelo de GUÍA: cabecera + 3 filas por semana (número/Mañana/Tarde).
   * Devuelve cuadrícula para setValues y las posiciones A1 editables M/T.
   * @param {number} y - año
   * @param {number} m - mes (1..12)
   * @param {number} startRow - fila inicial de la cuadrícula (por defecto 2)
   * @returns {{header:string[], grid:(string|number)[][], mtA1:string[], rowsNeeded:number}}
   */
  function buildGuideMonthArrays(y,m,startRow){
    startRow = startRow||2;
    var meta = monthMeta(y,m);
    var header = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
    var grid = [];
    var mtA1 = [];
    var day=1;

    for(var w=0; w<meta.weeks; w++){
      var numRow = new Array(7).fill('');
      var manRow = new Array(7).fill('');
      var tarRow = new Array(7).fill('');
      for(var dow=0; dow<7; dow++){
        var inRange = (w>0 || dow>=meta.firstWeekday) && day<=meta.lastDay;
        if(inRange){
          numRow[dow]=day;
          manRow[dow]='MAÑANA';
          tarRow[dow]='TARDE';
          var colA1 = colToA1(dow+1);
          var rowBase = startRow+(w*3);
          mtA1.push(colA1+(rowBase+1), colA1+(rowBase+2)); // M/T
          day++;
        }
      }
      grid.push(numRow,manRow,tarRow);
    }
    var rowsNeeded = 1 + meta.weeks*3; // 1 cabecera + 3 por semana
    return { header:header, grid:grid, mtA1:mtA1, rowsNeeded:rowsNeeded, meta:meta };
  }

  /**
   * Lista A1 de todas las celdas M/T válidas de un mes dado, asumiendo estructura guía.
   * Útil para DV/Protección. No consulta Sheets.
   * @param {{yyyy:number, mm:number, firstWeekday:number, lastDay:number, weeks:number}} meta
   * @param {number} startRow - por defecto 2
   * @returns {string[]}
   */
  function monthMT_A1_FromMeta(meta, startRow){
    startRow = startRow||2;
    var res = [], day=1;
    for(var w=0; w<meta.weeks; w++){
      var rowBase = startRow + (w*3);
      for(var dow=0; dow<7; dow++){
        var inRange = (w>0 || dow>=meta.firstWeekday) && day<=meta.lastDay;
        if(inRange){
          var a1 = colToA1(dow+1);
          res.push(a1+(rowBase+1), a1+(rowBase+2));
          day++;
        }
      }
    }
    return res;
  }

  /**
   * Parsea nombre de pestaña "MM_YYYY" a números.
   * @returns {{mm:number, yyyy:number}|null}
   */
  function parseTab_MMYYYY(name){
    var m = String(name||'').match(/^(\d{2})_(\d{4})$/);
    if(!m) return null;
    return { mm:parseInt(m[1],10), yyyy:parseInt(m[2],10) };
  }

  return {
    pad2: pad2,
    toISO: toISO,
    fromDisplay: fromDisplay,
    colToA1: colToA1,
    a1ToRowCol: a1ToRowCol,
    firstDayOfMonth: firstDayOfMonth,
    lastDayNumber: lastDayNumber,
    mondayIndex: mondayIndex,
    monthMeta: monthMeta,
    buildGuideMonthArrays: buildGuideMonthArrays,
    monthMT_A1_FromMeta: monthMT_A1_FromMeta,
    parseTab_MMYYYY: parseTab_MMYYYY
  };
})();
