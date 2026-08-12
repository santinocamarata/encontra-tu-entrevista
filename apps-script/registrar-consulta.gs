// Apps Script para el Sheet de ENTREVISTAS. Registra en las columnas
// E (Legajo), F (Nombre) y G (Hora exacta) el momento en que un alumno
// consultó su aula desde "Encontra tu aula" — trazabilidad de quién
// entró y cuándo, sin tocar las columnas A-D que ya usa el sistema.
//
// Deploy:
// 1. Abrir el Google Sheet -> Extensiones -> Apps Script.
// 2. Pegar este código (reemplazar el contenido de Code.gs).
// 3. Implementar -> Nueva implementación -> tipo "Aplicación web".
//    - Ejecutar como: Yo (tu cuenta)
//    - Quién tiene acceso: Cualquier usuario
// 4. Copiar la URL que termina en /exec y pasársela a Claude para
//    pegarla en gorrion.html como APPS_SCRIPT_URL.

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ENTREVISTAS');
  var body = JSON.parse(e.postData.contents);
  var legajo = String(body.legajo || '').trim();
  var nombre = String(body.nombre || '').trim();

  if (!legajo) {
    return respond({ ok: false, error: 'legajo requerido' });
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === legajo) {
      var fila = i + 1;
      var hora = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm:ss');
      sheet.getRange(fila, 5).setValue(legajo); // columna E
      sheet.getRange(fila, 6).setValue(nombre || data[i][3] || ''); // columna F
      sheet.getRange(fila, 7).setValue(hora); // columna G
      return respond({ ok: true });
    }
  }

  return respond({ ok: false, error: 'legajo no encontrado' });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
