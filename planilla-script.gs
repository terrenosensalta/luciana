/**
 * Conecta la app de Luciana con esta planilla.
 * Pegar en: Extensiones > Apps Script, y publicar como "Aplicación web".
 * - Pestaña "Resultados": una fila por cada respuesta (bien o a intentar).
 * - Pestaña "Resumen": totales por actividad y por día (se calculan solos).
 * - Pestaña "Progreso": copia de estrellas, figuritas y avances, para recuperarlos en otro celular.
 */
const ZONA = 'America/Argentina/Salta';

function preparar_(ss) {
  let res = ss.getSheetByName('Resultados');
  if (!res) {
    res = ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName('Resultados') : ss.insertSheet('Resultados');
    res.appendRow(['Fecha', 'Hora', 'Actividad', 'Resultado', 'Detalle', 'Estrellas']);
    res.getRange('A1:F1').setFontWeight('bold').setBackground('#FFF4D6');
    res.setFrozenRows(1);
    res.setColumnWidth(3, 200); res.setColumnWidth(5, 380);
  }
  if (!ss.getSheetByName('Resumen')) {
    const r = ss.insertSheet('Resumen');
    r.getRange('A1').setValue('Por actividad').setFontWeight('bold');
    r.getRange('A2').setFormula('=IFERROR(QUERY(Resultados!A:F,"select C, count(C) where C is not null group by C pivot D label C \'Actividad\'",1),"Todavía no hay resultados")');
    r.getRange('F1').setValue('Por día').setFontWeight('bold');
    r.getRange('F2').setFormula('=IFERROR(QUERY(Resultados!A:F,"select A, count(A) where A is not null group by A pivot D label A \'Fecha\'",1),"")');
    r.setColumnWidth(1, 200);
  }
  if (!ss.getSheetByName('Progreso')) {
    const p = ss.insertSheet('Progreso');
    p.getRange('A1:A3').setValues([['Último guardado'], ['Estrellas'], ['Datos (no editar)']]).setFontWeight('bold');
  }
}

function responder_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  preparar_(ss);
  const p = ss.getSheetByName('Progreso');
  const txt = p.getRange('B3').getValue();
  return responder_({ ok: true, url: ss.getUrl(), guardado: String(p.getRange('B1').getDisplayValue() || ''), datos: txt ? JSON.parse(txt) : null });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    preparar_(ss);
    const body = JSON.parse(e.postData.contents);
    const filas = (body.resultados || []).map(r => {
      const d = new Date(r.t);
      return [Utilities.formatDate(d, ZONA, 'dd/MM/yyyy'), Utilities.formatDate(d, ZONA, 'HH:mm'), r.a, r.ok ? '✅ Bien' : '❌ A intentar', r.d || '', r.s];
    });
    if (filas.length) {
      const res = ss.getSheetByName('Resultados');
      res.getRange(res.getLastRow() + 1, 1, filas.length, 6).setNumberFormat('@').setValues(filas);
    }
    if (body.progreso) {
      ss.getSheetByName('Progreso').getRange('B1:B3').setValues([
        [Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm')], [body.progreso.stars || 0], [JSON.stringify(body.progreso)]
      ]);
    }
    return responder_({ ok: true, guardadas: filas.length });
  } finally {
    lock.releaseLock();
  }
}
