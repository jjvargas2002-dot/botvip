/**
 * Webhook para el spreadsheet "STATUS DE CAJAS PENDIENTES".
 * Recibe POST desde BotVip (app.py -> notificar_status_cajas) cada vez que un
 * operador resuelve/agrupa (action="add") o desagrupa/reabre (action="remove")
 * una avería (principal o cliente agrupado) desde la página web / app.
 * - action="add": escribe la cuenta en la siguiente fila vacía de la columna
 *   C y "Reparado" en la columna G de esa misma fila.
 * - action="remove": busca la cuenta en la columna C y limpia esa fila
 *   (columnas C y G) en todas las coincidencias.
 * Todo esto en la pestaña identificada por SHEET_GID.
 *
 * NO se llama nunca desde la sincronización automática con el Drive/Excel
 * origen (cuando una cuenta desaparece o el propio Excel ya la marca como
 * resuelta) - eso se maneja aparte en la app como "Fuente sheet". Tampoco se
 * llama para averías creadas manualmente en la app (origen="MANUAL") ni para
 * sedes fuera de Lima (solo LI1/LI2/LI3/LI4/LI7) - eso lo filtra app.py
 * antes de llamar a este webhook.
 *
 * Instalación:
 * 1. Abrir el spreadsheet -> Extensiones -> Apps Script.
 * 2. Pegar este archivo (reemplazando el contenido de Code.gs, o como
 *    archivo nuevo).
 * 3. Implementar -> Gestionar implementaciones (o Nueva implementación) ->
 *    Tipo: Aplicación web.
 *    - Ejecutar como: Yo (tu cuenta).
 *    - Quién tiene acceso: Cualquier usuario.
 * 4. Copiar la URL de la app web y colocarla en BotBitel/.env como
 *    STATUS_CAJAS_WEBHOOK_URL. Debe coincidir con WEBHOOK_SECRET de abajo.
 *
 * Nota: el secreto se guarda como constante en el código (no en
 * PropertiesService) porque leer/escribir Propiedades del Script desde un
 * despliegue de Aplicación Web público falla con
 * "PERMISSION_DENIED al leer desde el almacenamiento" en algunos proyectos
 * tras cambiar el acceso del despliegue. El código fuente no es público,
 * solo la URL del endpoint, así que sigue siendo seguro.
 */

var SHEET_GID = 773109892; // gid de la pestaña "STATUS DE CAJAS" en la URL del sheet
var HEADER_ROW = 2;        // fila 2 tiene los encabezados; los datos empiezan en la fila 3
var COL_CUENTA = 3;        // columna C
var COL_STATUS = 7;        // columna G
var WEBHOOK_SECRET = 'naXl72P0xGs_RhQeeGDkNDtQZsuQGK55VngKJ0ATkys'; // debe coincidir con STATUS_CAJAS_WEBHOOK_SECRET (.env)

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse(e.postData.contents);

    if (!WEBHOOK_SECRET || body.secret !== WEBHOOK_SECRET) {
      return jsonResponse({ ok: false, error: 'unauthorized' });
    }

    var cuenta = (body.cuenta || '').toString().trim();
    if (!cuenta) {
      return jsonResponse({ ok: false, error: 'cuenta requerida' });
    }

    var action = (body.action || 'add').toString();

    var sheet = getSheetByGid(SHEET_GID);
    if (!sheet) {
      return jsonResponse({ ok: false, error: 'sheet con ese gid no encontrado' });
    }

    // Varias cuentas de un mismo grupo se notifican casi al mismo tiempo desde
    // la app. Sin este lock, dos llamadas concurrentes pueden calcular la misma
    // "siguiente fila vacía" y una sobrescribe a la otra.
    lock.waitLock(30000);
    try {
      if (action === 'remove') {
        var removedRows = removeCuenta(sheet, cuenta);
        SpreadsheetApp.flush();
        return jsonResponse({ ok: true, removed: removedRows, cuenta: cuenta });
      }

      var targetRow = findNextEmptyRow(sheet);
      sheet.getRange(targetRow, COL_CUENTA).setValue(cuenta);
      sheet.getRange(targetRow, COL_STATUS).setValue('Reparado');
      SpreadsheetApp.flush();
      return jsonResponse({ ok: true, row: targetRow, cuenta: cuenta });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function removeCuenta(sheet, cuenta) {
  var lastRow = sheet.getLastRow();
  var removedRows = [];
  if (lastRow >= HEADER_ROW + 1) {
    var values = sheet.getRange(HEADER_ROW + 1, COL_CUENTA, lastRow - HEADER_ROW, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if ((values[i][0] || '').toString().trim() === cuenta) {
        var row = HEADER_ROW + 1 + i;
        sheet.getRange(row, COL_CUENTA).setValue('');
        sheet.getRange(row, COL_STATUS).setValue('');
        removedRows.push(row);
      }
    }
  }
  return removedRows;
}

function getSheetByGid(gid) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) {
      return sheets[i];
    }
  }
  return null;
}

function findNextEmptyRow(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow >= HEADER_ROW + 1) {
    var values = sheet.getRange(HEADER_ROW + 1, COL_CUENTA, lastRow - HEADER_ROW, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (!values[i][0] || values[i][0].toString().trim() === '') {
        return HEADER_ROW + 1 + i;
      }
    }
  }
  return lastRow + 1;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
