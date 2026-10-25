/**
 * Coclia · recibe las respuestas de la encuesta y se las entrega a la página de resultados.
 * 1) Cambia la CLAVE por una palabra que solo tú sepas.
 * 2) Implementar > Nueva implementación > Aplicación web > Ejecutar como: yo > Acceso: cualquier persona.
 */
const CLAVE = 'CAMBIA-ESTA-CLAVE';

const COLUMNAS = ['id', 'fecha', 'ronda', 'codigo', 'nps', 'categoria', 'por_que', 'tarea',
  'dificultades', 'sin_dificultades', 'le_falta', 'no_necesita_mejoras', 'pagaria', 'cop_mes',
  'edad', 'ocupacion', 'tiene_dificultades', 'procedencia'];

const PERMITIDOS = {
  tarea: ['sin_ayuda', 'con_ayuda', 'no'],
  pagaria: ['si', 'no', 'tal_vez', ''],
  edad: ['', 'menos_18', '18_24', '25_34', '35_44', '45_54', '55_64', '65_mas'],
  ocupacion: ['', 'estudiante', 'empleado', 'independiente', 'salud', 'otra'],
  tiene_dificultades: ['', 'si', 'no'],
  procedencia: ['', 'familia', 'companeros', 'externo', 'otros']
};

function hoja_() {
  const libro = SpreadsheetApp.getActive();
  let hoja = libro.getSheetByName('Respuestas');
  if (!hoja) {
    hoja = libro.insertSheet('Respuestas');
    hoja.appendRow(COLUMNAS);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function ronda_() {
  return PropertiesService.getScriptProperties().getProperty('ronda') || 'Ronda 1';
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Evita que un texto se convierta en fórmula dentro de la hoja
function texto_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B-\u001F\u007F]/g, ' ').trim().slice(0, max || 600);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.accion === 'config') return salida_({ ok: true, ronda: ronda_() });
  if (p.clave !== CLAVE) return salida_({ ok: false, error: 'clave' });
  const filas = hoja_().getDataRange().getValues();
  return salida_({ ok: true, ronda: ronda_(), filas: filas });
}

function doPost(e) {
  const candado = LockService.getScriptLock();
  candado.waitLock(15000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (d.accion === 'ronda') {
      if (d.clave !== CLAVE) return salida_({ ok: false, error: 'clave' });
      PropertiesService.getScriptProperties().setProperty('ronda', texto_(d.ronda, 40) || 'Ronda 1');
      return salida_({ ok: true, ronda: ronda_() });
    }

    const nps = Number(d.nps);
    if (!Number.isInteger(nps) || nps < 0 || nps > 10) return salida_({ ok: false, error: 'nps' });
    if (!texto_(d.por_que)) return salida_({ ok: false, error: 'por_que' });
    if (PERMITIDOS.tarea.indexOf(d.tarea) < 0) return salida_({ ok: false, error: 'tarea' });
    const id = texto_(d.id, 60), codigo = texto_(d.codigo, 12), ronda = texto_(d.ronda, 40) || ronda_();
    if (!id) return salida_({ ok: false, error: 'id' });

    const hoja = hoja_();
    const ultima = hoja.getLastRow();
    if (ultima > 1) {
      const datos = hoja.getRange(2, 1, ultima - 1, 4).getValues();
      for (let i = 0; i < datos.length; i++) {
        if (datos[i][0] === id) return salida_({ ok: true, repetida: true });
        if (codigo && datos[i][3] === codigo && datos[i][2] === ronda) return salida_({ ok: true, repetida: true });
      }
    }

    const elegir = (campo) => (PERMITIDOS[campo].indexOf(d[campo] || '') >= 0 ? (d[campo] || '') : '');
    const cop = d.cop_mes === '' || d.cop_mes == null ? '' : Math.max(0, Math.min(10000000, Math.round(Number(d.cop_mes)) || 0));
    const categoria = nps <= 6 ? 'detractor' : nps <= 8 ? 'pasivo' : 'promotor';

    hoja.appendRow([
      id, new Date(), ronda, codigo, nps, categoria,
      texto_(d.por_que), d.tarea,
      texto_(d.dificultades), d.sin_dificultades === 'si' ? 'si' : '',
      texto_(d.le_falta), d.no_necesita_mejoras === 'si' ? 'si' : '',
      elegir('pagaria'), cop,
      elegir('edad'), elegir('ocupacion'), elegir('tiene_dificultades'), elegir('procedencia')
    ]);
    return salida_({ ok: true });
  } catch (err) {
    return salida_({ ok: false, error: 'formato' });
  } finally {
    candado.releaseLock();
  }
}
