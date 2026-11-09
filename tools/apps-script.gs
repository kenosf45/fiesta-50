/** @OnlyCurrentDoc */
/**
 * Los 50 de Marco — recibe las confirmaciones de la invitación y las guarda en esta hoja.
 *
 * Cómo se instala (una sola vez):
 *  1. Extensiones → Apps Script, pegar este archivo.
 *  2. Ejecutar la función `setup` (Google pide permiso: Permitir).
 *  3. Implementar → Nueva implementación → Aplicación web
 *     Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario.
 *  4. Copiar la URL /exec en `API` dentro de index.html.
 */

const SHEET_NAME = 'Confirmaciones';
const HEADER_ROW = 6;
const FIRST_DATA_ROW = HEADER_ROW + 1;
const HEADERS = ['Fecha', 'Nombre', 'Teléfono', '¿Viene?', 'Personas', 'Mensaje para Marco', 'Canción para el DJ', 'ID'];
const COL = { fecha: 1, nombre: 2, telefono: 3, viene: 4, personas: 5, mensaje: 6, cancion: 7, id: 8 };

/** Prepara la hoja: título, totales y encabezados. Se puede volver a correr; no borra respuestas. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
  }
  ss.setSpreadsheetTimeZone('America/New_York');

  sh.getRange('A1').setValue('Los 50 de Marco · Confirmaciones')
    .setFontSize(16).setFontWeight('bold').setFontColor('#9a7531');

  const d = 'D' + FIRST_DATA_ROW + ':D';
  const e = 'E' + FIRST_DATA_ROW + ':E';
  sh.getRange('A2:B4').setValues([
    ['Total de personas que vienen', ''],
    ['Respuestas: sí vienen', ''],
    ['Respuestas: no pueden', ''],
  ]);
  sh.getRange('B2').setFormula('=SUMIF(' + d + ',"Sí",' + e + ')');
  sh.getRange('B3').setFormula('=COUNTIF(' + d + ',"Sí")');
  sh.getRange('B4').setFormula('=COUNTIF(' + d + ',"No")');
  sh.getRange('A2:A4').setFontWeight('bold');
  sh.getRange('B2:B4').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('left');
  sh.getRange('B2').setFontColor('#9a7531').setFontSize(16);

  sh.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#0a0907').setFontColor('#d6b262');
  sh.setFrozenRows(HEADER_ROW);

  sh.setColumnWidth(COL.fecha, 220);
  sh.setColumnWidth(COL.nombre, 200);
  sh.setColumnWidth(COL.telefono, 140);
  sh.setColumnWidth(COL.viene, 80);
  sh.setColumnWidth(COL.personas, 80);
  sh.setColumnWidth(COL.mensaje, 360);
  sh.setColumnWidth(COL.cancion, 220);
  sh.getRange(FIRST_DATA_ROW, COL.fecha, sh.getMaxRows() - HEADER_ROW, 1).setNumberFormat('mmm d, h:mm am/pm');
  sh.getRange(FIRST_DATA_ROW, COL.mensaje, sh.getMaxRows() - HEADER_ROW, 1).setWrap(true);
  sh.hideColumns(COL.id);
}

/** Recibe { action: 'rsvp' | 'extras', ... } desde la invitación. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

    if (data.action === 'rsvp') {
      const attending = data.attending !== false;
      const id = clean(data.id, 40) || Utilities.getUuid();
      const values = [
        new Date(),
        text(data.name, 80),
        text(data.phone, 25),
        attending ? 'Sí' : 'No',
        attending ? Math.max(1, Math.min(99, parseInt(data.guests, 10) || 1)) : 0,
      ];
      const row = findRow(sh, id);
      if (row) {
        sh.getRange(row, 1, 1, values.length).setValues([values]);
      } else {
        sh.appendRow(values.concat(['', '', id]));
      }
      return json({ ok: true, id: id });
    }

    if (data.action === 'extras') {
      const row = findRow(sh, clean(data.id, 40));
      if (!row) return json({ ok: false, error: 'not_found' });
      sh.getRange(row, COL.mensaje, 1, 2).setValues([[text(data.message, 1000), text(data.song, 120)]]);
      return json({ ok: true, id: data.id });
    }

    return json({ ok: false, error: 'bad_action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json({ ok: true });
}

function findRow(sh, id) {
  if (!id) return null;
  const last = sh.getLastRow();
  if (last < FIRST_DATA_ROW) return null;
  const ids = sh.getRange(FIRST_DATA_ROW, COL.id, last - HEADER_ROW, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return FIRST_DATA_ROW + i;
  }
  return null;
}

function clean(v, max) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
}

/** Texto seguro para la hoja: nada que empiece con = + - @ se toma como fórmula. */
function text(v, max) {
  const s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
