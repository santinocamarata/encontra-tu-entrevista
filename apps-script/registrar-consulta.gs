// Apps Script para el Sheet de ENTREVISTAS. Cada vez que un alumno
// consulta su aula desde "Encontra tu aula", agrega una fila al log de
// trazabilidad en las columnas E (Legajo), F (Nombre) y G (Hora exacta),
// empezando desde la primera fila en blanco (fila 2 en adelante). No
// depende de en qué fila esté ese legajo en las columnas A-D: es un
// registro secuencial de consultas, no una anotación por alumno.
//
// Deploy:
// 1. Abrir el Google Sheet -> Extensiones -> Apps Script.
// 2. Pegar este código (reemplazar el contenido de Code.gs).
// 3. Implementar -> Administrar implementaciones -> Editar (lápiz) ->
//    Versión: Nueva versión -> Implementar.
//    (si es la primera vez: Implementar -> Nueva implementación -> tipo
//    "Aplicación web", Ejecutar como "Yo", Acceso "Cualquier usuario")

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ENTREVISTAS');
  var body = JSON.parse(e.postData.contents);
  var legajo = String(body.legajo || '').trim();
  var nombre = String(body.nombre || '').trim();

  if (!legajo) {
    return respond({ ok: false, error: 'legajo requerido' });
  }

  var fila = proximaFilaLibre(sheet, 5); // columna E
  var hora = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm:ss');

  sheet.getRange(fila, 5).setValue(legajo); // columna E
  sheet.getRange(fila, 6).setValue(nombre); // columna F
  sheet.getRange(fila, 7).setValue(hora); // columna G

  return respond({ ok: true, fila: fila });
}

// Busca la primera fila (desde la 2) donde la columna dada está vacía.
function proximaFilaLibre(sheet, columna) {
  var ultimaFila = Math.max(sheet.getMaxRows(), 2);
  var valores = sheet.getRange(2, columna, ultimaFila - 1, 1).getValues();
  for (var i = 0; i < valores.length; i++) {
    if (valores[i][0] === '' || valores[i][0] === null) {
      return i + 2;
    }
  }
  return valores.length + 2;
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
