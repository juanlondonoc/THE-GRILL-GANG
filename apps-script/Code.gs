/**
 * The Grill Gang — recibe las cotizaciones de "Quiero que me llamen"
 * y las guarda en la hoja "Cotizaciones" de este Google Sheets.
 *
 * Instalación: en el Google Sheets → Extensiones → Apps Script,
 * pegar este código, guardar e Implementar → Nueva implementación →
 * Tipo "Aplicación web", Ejecutar como "Yo", Acceso "Cualquier usuario".
 * Copiar la URL /exec y ponerla en SHEETS_URL dentro de index.html.
 */
var HOJA = 'Cotizaciones';
var COLUMNAS = ['Fecha de registro', 'Nombre', 'Correo', 'Teléfono', 'Plan', 'Precio p/p',
  'Personas', 'Fecha del evento', 'Ubicación', 'Entradas y adicionales', 'Total estimado'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(COLUMNAS);
      hoja.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold');
      hoja.setFrozenRows(1);
    }
    var p = e.parameter;
    // El apóstrofo evita que Sheets interprete teléfonos o textos como fórmulas
    var txt = function (v) { return "'" + String(v || '').slice(0, 1000); };
    hoja.appendRow([
      new Date(), txt(p.nombre), txt(p.correo), txt(p.telefono), txt(p.plan),
      Number(p.precio_pp) || '', Number(p.personas) || '', txt(p.fecha),
      txt(p.ubicacion), txt(p.extras), Number(p.total) || ''
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
