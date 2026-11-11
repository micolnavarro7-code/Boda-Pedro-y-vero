/*
  Guarda RSVP, canciones, apuestas, consejos y resultados del quiz en una Google Sheet.

  1. Crea una hoja de cálculo nueva en Google Drive.
  2. Extensiones → Apps Script. Borra lo que haya y pega este archivo.
  3. Implementar → Nueva implementación → Tipo: Aplicación web.
     Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario.
  4. Copia la URL que termina en /exec y pégala en CONFIG.endpoint (index.html).

  Cada tipo de envío crea su propia pestaña (rsvp, playlist, apuestas, consejos, quiz).
*/

function doPost(e) {
  const { tipo, datos } = JSON.parse(e.postData.contents);
  const permitidos = ['rsvp', 'playlist', 'apuestas', 'consejos', 'quiz'];
  if (!permitidos.includes(tipo)) return ContentService.createTextOutput('tipo no válido');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(tipo) || ss.insertSheet(tipo);
    let cabecera = hoja.getLastRow() ? hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0] : [];
    const nuevas = Object.keys(datos).filter(k => !cabecera.includes(k));
    if (nuevas.length) {
      cabecera = cabecera.concat(nuevas);
      hoja.getRange(1, 1, 1, cabecera.length).setValues([cabecera]).setFontWeight('bold');
    }
    hoja.appendRow(cabecera.map(k => (k in datos ? String(datos[k]).slice(0, 1000) : '')));
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// Devuelve las últimas canciones para mostrarlas en la web (solo nombre, canción, artista y momento).
function doGet(e) {
  if (e.parameter.tipo !== 'playlist') return json([]);
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('playlist');
  if (!hoja || hoja.getLastRow() < 2) return json([]);
  const [cab, ...filas] = hoja.getDataRange().getValues();
  const col = k => cab.indexOf(k);
  const lista = filas.reverse().slice(0, 8).map(f => ({
    nombre: f[col('nombre')], cancion: f[col('cancion')], artista: f[col('artista')], momento: f[col('momento')],
  }));
  return json(lista);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
