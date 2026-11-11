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
 *
 * Para cambios posteriores: pegar, guardar, ejecutar `setup`, y luego
 * Implementar → Gestionar implementaciones → editar → Versión nueva (la URL no cambia).
 */

const SHEET_NAME = 'Confirmaciones';
const HEADER_ROW = 6;
const FIRST_DATA_ROW = HEADER_ROW + 1;
const HEADERS = ['Fecha', 'Nombre', 'Teléfono', '¿Viene?', 'Personas', 'Adultos', 'Niños', 'Mensaje para Marco', 'Canción para el DJ', 'ID'];
const COL = { fecha: 1, nombre: 2, telefono: 3, viene: 4, personas: 5, adultos: 6, ninos: 7, mensaje: 8, cancion: 9, id: 10 };

/** Prepara la hoja: título, totales y encabezados. Se puede volver a correr; no borra respuestas. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
  }
  ss.setSpreadsheetTimeZone('America/New_York');

  // La primera versión no tenía Adultos/Niños: insertar esas dos columnas después de Personas
  // (las respuestas que ya existan se corren a la derecha sin perder nada).
  const afterPersonas = sh.getRange(HEADER_ROW, COL.adultos).getValue();
  if (afterPersonas && afterPersonas !== 'Adultos') {
    sh.insertColumnsAfter(COL.personas, 2);
    // Respuestas anteriores con 0 o 1 persona: ya se sabe cuántos adultos y niños son.
    const last = sh.getLastRow();
    if (last >= FIRST_DATA_ROW) {
      const n = last - HEADER_ROW;
      const personas = sh.getRange(FIRST_DATA_ROW, COL.personas, n, 1).getValues();
      sh.getRange(FIRST_DATA_ROW, COL.adultos, n, 2).setValues(
        personas.map((r) => (r[0] === 0 ? [0, 0] : r[0] === 1 ? [1, 0] : ['', ''])));
    }
  }

  sh.getRange('A1').setValue('Los 50 de Marco · Confirmaciones')
    .setFontSize(16).setFontWeight('bold').setFontColor('#9a7531');

  const viene = 'D' + FIRST_DATA_ROW + ':D';
  const sumOf = (letter) => '=SUMIF(' + viene + ',"Sí",' + letter + FIRST_DATA_ROW + ':' + letter + ')';
  sh.getRange('A2:D4').setValues([
    ['Total de personas que vienen', '', 'Adultos', ''],
    ['Respuestas: sí vienen', '', 'Niños', ''],
    ['Respuestas: no pueden', '', '', ''],
  ]);
  sh.getRange('B2').setFormula(sumOf('E'));
  sh.getRange('B3').setFormula('=COUNTIF(' + viene + ',"Sí")');
  sh.getRange('B4').setFormula('=COUNTIF(' + viene + ',"No")');
  sh.getRange('D2').setFormula(sumOf('F'));
  sh.getRange('D3').setFormula(sumOf('G'));
  sh.getRange('A2:A4').setFontWeight('bold');
  sh.getRange('C2:C3').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('B2:B4').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('left');
  sh.getRange('D2:D3').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('left').setFontColor('#9a7531');
  sh.getRange('B2').setFontColor('#9a7531').setFontSize(16);

  sh.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#0a0907').setFontColor('#d6b262');
  sh.setFrozenRows(HEADER_ROW);

  sh.setColumnWidth(COL.fecha, 220);
  sh.setColumnWidth(COL.nombre, 200);
  sh.setColumnWidth(COL.telefono, 140);
  sh.setColumnWidth(COL.viene, 80);
  sh.setColumnWidth(COL.personas, 80);
  sh.setColumnWidth(COL.adultos, 80);
  sh.setColumnWidth(COL.ninos, 80);
  sh.setColumnWidth(COL.mensaje, 360);
  sh.setColumnWidth(COL.cancion, 220);
  const rows = sh.getMaxRows() - HEADER_ROW;
  sh.getRange(FIRST_DATA_ROW, COL.fecha, rows, 1).setNumberFormat('mmm d, h:mm am/pm');
  sh.getRange(FIRST_DATA_ROW, COL.mensaje, rows, 1).setWrap(true);
  // Teléfonos como texto: así no se pierde el 0 inicial (ej. celulares de Ecuador 09…)
  sh.getRange(FIRST_DATA_ROW, COL.telefono, rows, 1).setNumberFormat('@').setHorizontalAlignment('left');
  sh.showColumns(1, HEADERS.length);
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
      // Las invitaciones viejas mandan solo "guests"; se cuentan como adultos.
      const adults = attending ? int(data.adults != null ? data.adults : data.guests, 1, 99, 1) : 0;
      const kids = attending ? int(data.kids, 0, 99, 0) : 0;
      const values = [
        new Date(),
        text(data.name, 80),
        text(data.phone, 25),
        attending ? 'Sí' : 'No',
        adults + kids,
        adults,
        kids,
      ];
      const row = findRow(sh, id);
      if (row) {
        sh.getRange(row, 1, 1, values.length).setValues([values]);
      } else {
        // setValues en la siguiente fila (no appendRow) respeta los formatos de la hoja:
        // fecha corta y teléfono como texto (con su 0 inicial).
        const next = sh.getLastRow() + 1;
        if (next > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
        sh.getRange(next, 1, 1, HEADERS.length).setValues([values.concat(['', '', id])]);
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

function int(v, min, max, fallback) {
  const n = parseInt(v, 10);
  return isNaN(n) ? fallback : Math.max(min, Math.min(max, n));
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
