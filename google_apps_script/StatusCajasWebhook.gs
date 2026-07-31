/**
 * Webhook para el spreadsheet "STATUS DE CAJAS PENDIENTES".
 * Recibe POST desde BotVip (app.py -> notificar_status_cajas) cada vez que un
 * operador resuelve una avería (principal o cliente agrupado) desde la
 * página web / app. Escribe la cuenta en la siguiente fila vacía de la
 * columna C y "Reparado" en la columna G de esa misma fila, en la pestaña
 * identificada por SHEET_GID.
 *
 * NO se llama nunca desde la sincronización automática con el Drive/Excel
 * origen (cuando una cuenta desaparece o el propio Excel ya la marca como
 * resuelta) - eso se maneja aparte en la app como "Fuente sheet".
 *
 * Instalación:
 * 1. Abrir el spreadsheet -> Extensiones -> Apps Script.
 * 2. Pegar este archivo (reemplazando el contenido de Code.gs, o como
 *    archivo nuevo).
 * 3. En el editor, seleccionar la función "setup" en el desplegable de
 *    funciones y ejecutarla una vez (Run). Autoriza los permisos que pida.
 *    Esto guarda el secreto compartido en las Propiedades del Script.
 * 4. Implementar -> Nueva implementación -> Tipo: Aplicación web.
 *    - Ejecutar como: Yo (tu cuenta).
 *    - Quién tiene acceso: Cualquier usuario.
 * 5. Copiar la URL de la app web y colocarla en BotBitel/.env como
 *    STATUS_CAJAS_WEBHOOK_URL. El secreto ya está en .env como
 *    STATUS_CAJAS_WEBHOOK_SECRET (debe coincidir con el que pongas abajo
 *    en setup()).
 */

var SHEET_GID = 773109892; // gid de la pestaña "STATUS DE CAJAS" en la URL del sheet
var HEADER_ROW = 2;        // fila 2 tiene los encabezados; los datos empiezan en la fila 3
var COL_CUENTA = 3;        // columna C
var COL_STATUS = 7;        // columna G

function setup() {
  // Reemplaza el valor por el mismo secreto configurado en STATUS_CAJAS_WEBHOOK_SECRET (.env)
  PropertiesService.getScriptProperties().setProperty('WEBHOOK_SECRET', 'naXl72P0xGs_RhQeeGDkNDtQZsuQGK55VngKJ0ATkys');
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var expectedSecret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');

    if (!expectedSecret || body.secret !== expectedSecret) {
      return jsonResponse({ ok: false, error: 'unauthorized' });
    }

    var cuenta = (body.cuenta || '').toString().trim();
    if (!cuenta) {
      return jsonResponse({ ok: false, error: 'cuenta requerida' });
    }

    var sheet = getSheetByGid(SHEET_GID);
    if (!sheet) {
      return jsonResponse({ ok: false, error: 'sheet con ese gid no encontrado' });
    }

    var targetRow = findNextEmptyRow(sheet);
    sheet.getRange(targetRow, COL_CUENTA).setValue(cuenta);
    sheet.getRange(targetRow, COL_STATUS).setValue('Reparado');

    return jsonResponse({ ok: true, row: targetRow, cuenta: cuenta });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
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
