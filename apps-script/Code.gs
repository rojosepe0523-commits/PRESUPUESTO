/**
 * API de la app de presupuesto (versión 4).
 *
 * Reemplaza por completo el Code.gs anterior. Es compatible con la app (acepta las
 * mismas acciones "datos", "agregar" y "eliminar") y agrega:
 *   - Catálogo dinámico: crear, editar, renombrar y borrar categorías y métodos de pago desde la app.
 *   - Cierre de ciclo: cada vez que registras un salario se cierra el mes anterior en la pestaña "Cierres".
 *   - Recomendación mensual con IA (Claude) al cerrar cada ciclo. Sin clave de API usa reglas fijas.
 *
 * Instalación (una sola vez):
 *   1. Pega este archivo en lugar de tu Code.gs.
 *   2. (Opcional, para IA) Configuración del proyecto > Propiedades del script > ANTHROPIC_API_KEY = tu clave.
 *   3. Ejecuta "configurar" y acepta los permisos.
 *   4. Implementar > Gestionar implementaciones > editar (lápiz) > Versión: Nueva versión > Implementar.
 *      Así la URL /exec sigue siendo la misma.
 */

const API_VERSION = 4;
const H_GASTOS = 'Gastos', H_INGRESOS = 'Ingresos', H_CATEGORIAS = 'Categorias', H_RESUMEN = 'Resumen', H_CIERRES = 'Cierres';
const ENC_GASTOS = ['Fecha', 'Descripción', 'Categoría', 'Monto', 'Método de pago', 'Notas', 'Registrado', 'ID'];
const ENC_INGRESOS = ['Fecha', 'Descripción', 'Tipo', 'Monto', 'Notas', 'Registrado', 'ID'];
const ENC_CATEGORIAS = ['Categoría', 'Presupuesto mensual', 'Gastado este periodo', 'Disponible', '% usado', 'Tipo'];
const ENC_CIERRES = ['Inicio', 'Fin', 'Días', 'Ingresos', 'Gastos', 'Saldo', 'Ahorro', 'Tasa de ahorro', 'Presupuesto',
  'Detalle por categoría', 'Recomendación', 'Fuente', 'Generada'];
const TIPOS_CATEGORIA = ['fija', 'esencial', 'opcional', 'ahorro'];
const TIPO_POR_DEFECTO = {
  'Casa': 'fija', 'Servicios': 'fija', 'Educación': 'fija', 'Deudas': 'fija', 'Ahorro de emergencia': 'ahorro',
  'Supermercado': 'esencial', 'Transporte': 'esencial', 'Salud': 'esencial', 'Mascotas': 'esencial',
  'Comida': 'opcional', 'Entretenimiento': 'opcional', 'Ropa': 'opcional', 'Otros': 'opcional'
};
const METODOS_POR_DEFECTO = ['Efectivo', 'Débito', 'Crédito', 'Transferencia'];
const MAX_CATEGORIAS = 30;            // Resumen!A8:C37 tiene espacio para 30
const DIAS_DE_HISTORIA = 400;         // datos que se envían a la app
const FORMATO_COP = '"$"#,##0';
const FORMATO_FECHA = 'dd/mm/yyyy';
const IA_MODELO_POR_DEFECTO = 'claude-opus-5-5';
const PROPS = PropertiesService.getScriptProperties();

/* ============================ Entrada web ============================ */

function doGet() {
  return json_({ ok: true, mensaje: 'API de presupuesto v' + API_VERSION + '. Usa la app para conectarte.' });
}

function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Petición inválida' }); }
  if (!p || p.clave !== clave_()) return json_({ ok: false, error: 'Clave incorrecta. Ejecuta "mostrarClave" en Apps Script y pégala en Ajustes.' });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return json_({ ok: false, error: 'La hoja está ocupada, intenta de nuevo.' }); }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const r = { ok: true };
    switch (p.accion) {
      case 'datos': break;
      case 'agregar': Object.assign(r, agregar_(ss, p)); break;
      case 'eliminar': eliminar_(ss, String(p.id || '')); break;
      case 'categoria_guardar': categoriaGuardar_(ss, p); break;
      case 'categoria_eliminar': categoriaEliminar_(ss, String(p.nombre || '')); break;
      case 'metodos_guardar': metodosGuardar_(p.metodos); break;
      case 'recomendacion_generar':
        cerrarCiclos_(ss);
        generarRecomendacion_(ss, true);
        break;
      default: return json_({ ok: false, error: 'Acción desconocida: ' + p.accion });
    }
    r.datos = datos_(ss);
    return json_(r);
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ Configuración ============================ */

/** Ejecútala una vez después de pegar el archivo: crea pestañas, columnas y el activador diario. */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hoja_(ss, H_GASTOS, ENC_GASTOS);
  hoja_(ss, H_INGRESOS, ENC_INGRESOS);
  escribirCategorias_(ss, leerCategorias_(ss));   // agrega la columna "Tipo" sin tocar tus presupuestos
  hoja_(ss, H_CIERRES, ENC_CIERRES);
  clave_();

  ScriptApp.getProjectTriggers()
    .filter(t => ['tareaDiaria', 'generarRecomendacionPendiente'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tareaDiaria').timeBased().everyDays(1).atHour(6).create();

  cerrarCiclos_(ss);
  const conIA = !!PROPS.getProperty('ANTHROPIC_API_KEY');
  ss.toast('Listo. ' + (conIA ? 'Recomendaciones con IA activas.' : 'Sin ANTHROPIC_API_KEY: las recomendaciones usarán reglas fijas.') +
    ' Ahora publica una nueva versión de la implementación.', 'Presupuesto', 10);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Presupuesto')
    .addItem('Configurar (una vez)', 'configurar')
    .addItem('Generar recomendación del último mes', 'generarRecomendacionAhora')
    .addItem('Mostrar clave de la app', 'mostrarClave')
    .addToUi();
}

function mostrarClave() {
  const c = clave_();
  console.log('Tu clave secreta es: ' + c);
  try { SpreadsheetApp.getUi().alert('Clave secreta de la app', c, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* sin interfaz */ }
  return c;
}

/** Usa la clave que ya tenías (si existe) para que la app siga conectada. */
function clave_() {
  const nombres = ['CLAVE', 'clave', 'CLAVE_SECRETA', 'SECRET', 'API_KEY', 'APP_KEY', 'TOKEN'];
  for (let i = 0; i < nombres.length; i++) {
    const v = PROPS.getProperty(nombres[i]);
    if (v) { if (nombres[i] !== 'CLAVE') PROPS.setProperty('CLAVE', v); return v; }
  }
  const nueva = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  PROPS.setProperty('CLAVE', nueva);
  return nueva;
}

/** Activador diario: cierra ciclos pendientes y genera la recomendación si falta. */
function tareaDiaria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  cerrarCiclos_(ss);
  generarRecomendacion_(ss, false);
}

/** Activador de un solo uso que se programa al registrar un salario. */
function generarRecomendacionPendiente() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'generarRecomendacionPendiente').forEach(t => ScriptApp.deleteTrigger(t));
  tareaDiaria();
}

function generarRecomendacionAhora() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  cerrarCiclos_(ss);
  const r = generarRecomendacion_(ss, true);
  ss.toast(r ? 'Recomendación generada (' + r.fuente + '). Mírala en la app o en la pestaña Cierres.' : 'Aún no hay un mes cerrado: registra al menos dos salarios.', 'Presupuesto', 8);
}

/* ============================ Utilidades ============================ */

function hoja_(ss, nombre, encabezados) {
  let h = ss.getSheetByName(nombre);
  if (!h) h = ss.insertSheet(nombre);
  if (encabezados) {
    const actual = h.getRange(1, 1, 1, encabezados.length).getValues()[0];
    if (actual.join('|') !== encabezados.join('|')) {
      h.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
      h.setFrozenRows(1);
    }
  }
  return h;
}

function zona_(ss) { return ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }

function aIso_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

/** "2026-11-05" → fecha a la medianoche en la zona horaria de la hoja (evita el corrimiento a las 22:00). */
function deIso_(s, tz) { return Utilities.parseDate(s, tz, 'yyyy-MM-dd'); }

function sumarDiasIso_(s, n) {
  const p = s.split('-');
  const d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return d.toISOString().slice(0, 10);
}

function diasEntreIso_(a, b) {
  const pa = a.split('-'), pb = b.split('-');
  return Math.round((Date.UTC(+pb[0], +pb[1] - 1, +pb[2]) - Date.UTC(+pa[0], +pa[1] - 1, +pa[2])) / 86400000);
}

function numero_(v) {
  if (typeof v === 'number') return v;
  const s = String(v == null ? '' : v).replace(/[\s$]/g, '').replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

/* ============================ Lectura ============================ */

function leerMovimientos_(ss) {
  const tz = zona_(ss);
  const leer = (nombre, enc, idCol) => {
    const h = hoja_(ss, nombre, enc);
    const n = h.getLastRow() - 1;
    if (n < 1) return [];
    const vals = h.getRange(2, 1, n, enc.length).getValues();
    // Filas escritas a mano sin ID: se les asigna uno para poder borrarlas desde la app.
    let cambio = false;
    vals.forEach(r => { if ((r[0] || r[3]) && !r[idCol]) { r[idCol] = Utilities.getUuid(); cambio = true; } });
    if (cambio) h.getRange(2, idCol + 1, n, 1).setValues(vals.map(r => [r[idCol]]));
    return vals;
  };
  const gastos = leer(H_GASTOS, ENC_GASTOS, 7)
    .filter(r => r[0] && r[3] !== '')
    .map(r => ({ id: String(r[7]), fecha: aIso_(r[0], tz), descripcion: String(r[1] || ''), categoria: String(r[2] || 'Otros'),
      monto: numero_(r[3]) || 0, metodo: String(r[4] || ''), notas: String(r[5] || '') }));
  const ingresos = leer(H_INGRESOS, ENC_INGRESOS, 6)
    .filter(r => r[0] && r[3] !== '')
    .map(r => ({ id: String(r[6]), fecha: aIso_(r[0], tz), descripcion: String(r[1] || ''), tipo: String(r[2] || 'Otro ingreso'),
      monto: numero_(r[3]) || 0, notas: String(r[4] || '') }));
  return { gastos: gastos.filter(g => g.fecha), ingresos: ingresos.filter(i => i.fecha) };
}

function leerCategorias_(ss) {
  const h = hoja_(ss, H_CATEGORIAS);
  const n = h.getLastRow() - 1;
  if (n < 1) return Object.keys(TIPO_POR_DEFECTO).map(k => ({ nombre: k, presupuesto: 0, tipo: TIPO_POR_DEFECTO[k] }));
  return h.getRange(2, 1, n, 6).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => {
      const nombre = String(r[0]).trim();
      const tipo = String(r[5] || '').trim().toLowerCase();
      return { nombre: nombre, presupuesto: numero_(r[1]) || 0,
        tipo: TIPOS_CATEGORIA.indexOf(tipo) >= 0 ? tipo : (TIPO_POR_DEFECTO[nombre] || 'opcional') };
    });
}

function leerMetodos_() {
  try { const m = JSON.parse(PROPS.getProperty('METODOS') || 'null'); if (m && m.length) return m; } catch (e) { /* valor dañado */ }
  return METODOS_POR_DEFECTO.slice();
}

function leerCierres_(ss) {
  const h = hoja_(ss, H_CIERRES, ENC_CIERRES);
  const n = h.getLastRow() - 1;
  if (n < 1) return [];
  const tz = zona_(ss);
  return h.getRange(2, 1, n, ENC_CIERRES.length).getValues().filter(r => r[0]).map((r, i) => {
    let detalle = {}, rec = null;
    try { detalle = JSON.parse(r[9] || '{}'); } catch (e) { /* texto editado a mano */ }
    try { rec = r[10] ? JSON.parse(r[10]) : null; } catch (e) { rec = { titulo: 'Recomendación', diagnostico: String(r[10]) }; }
    return { fila: i + 2, inicio: aIso_(r[0], tz), fin: aIso_(r[1], tz), dias: +r[2] || 0, ingresos: +r[3] || 0, gastos: +r[4] || 0,
      saldo: +r[5] || 0, ahorro: +r[6] || 0, tasaAhorro: +r[7] || 0, presupuesto: +r[8] || 0, porCategoria: detalle,
      recomendacion: rec, fuente: String(r[11] || ''), generada: r[12] instanceof Date ? r[12].toISOString() : String(r[12] || '') };
  });
}

function datos_(ss) {
  const tz = zona_(ss);
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const desde = sumarDiasIso_(hoy, -DIAS_DE_HISTORIA);
  const m = leerMovimientos_(ss);
  const cierres = leerCierres_(ss).slice(-12).map(c => { const o = Object.assign({}, c); delete o.fila; return o; });
  return {
    version: API_VERSION,
    desde: desde,
    categorias: leerCategorias_(ss),
    metodos: leerMetodos_(),
    gastos: m.gastos.filter(g => g.fecha >= desde).map(g => [g.id, g.fecha, g.categoria, g.monto, g.metodo, g.descripcion]),
    ingresos: m.ingresos.filter(i => i.fecha >= desde).map(i => [i.id, i.fecha, i.tipo, i.monto, i.descripcion]),
    cierres: cierres,
    ia: !!PROPS.getProperty('ANTHROPIC_API_KEY')
  };
}

/* ============================ Escritura de movimientos ============================ */

function agregar_(ss, p) {
  const tz = zona_(ss);
  const m = leerMovimientos_(ss);
  const existentes = new Set(m.gastos.map(g => g.id).concat(m.ingresos.map(i => i.id)));
  const ids = [], rechazados = [], filasG = [], filasI = [];
  let haySalario = false;

  const validar = x => {
    const monto = numero_(x.monto);
    if (!(monto > 0)) return 'monto inválido';
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(x.fecha || ''))) return 'fecha inválida';
    return '';
  };
  const registrado = x => { const d = new Date(x.creado); return isNaN(d) ? new Date() : d; };

  (p.gastos || []).forEach(x => {
    if (!x || !x.id) return;
    if (existentes.has(x.id)) { ids.push(x.id); return; }
    const motivo = validar(x);
    if (motivo) { rechazados.push({ id: x.id, motivo: motivo }); return; }
    filasG.push([deIso_(x.fecha, tz), String(x.descripcion || ''), String(x.categoria || 'Otros'), Math.round(numero_(x.monto)),
      String(x.metodo || ''), String(x.notas || ''), registrado(x), x.id]);
    existentes.add(x.id); ids.push(x.id);
  });
  (p.ingresos || []).forEach(x => {
    if (!x || !x.id) return;
    if (existentes.has(x.id)) { ids.push(x.id); return; }
    const motivo = validar(x);
    if (motivo) { rechazados.push({ id: x.id, motivo: motivo }); return; }
    const tipo = x.tipo === 'Salario' ? 'Salario' : 'Otro ingreso';
    if (tipo === 'Salario') haySalario = true;
    filasI.push([deIso_(x.fecha, tz), String(x.descripcion || ''), tipo, Math.round(numero_(x.monto)), String(x.notas || ''), registrado(x), x.id]);
    existentes.add(x.id); ids.push(x.id);
  });

  const anexar = (nombre, enc, filas) => {
    if (!filas.length) return;
    const h = hoja_(ss, nombre, enc);
    const ini = h.getLastRow() + 1;
    h.getRange(ini, 1, filas.length, enc.length).setValues(filas);
    h.getRange(ini, 1, filas.length, 1).setNumberFormat(FORMATO_FECHA);
    h.getRange(ini, 4, filas.length, 1).setNumberFormat(FORMATO_COP);
  };
  anexar(H_GASTOS, ENC_GASTOS, filasG);
  anexar(H_INGRESOS, ENC_INGRESOS, filasI);

  let cicloCerrado = false;
  if (haySalario) {
    SpreadsheetApp.flush();
    cicloCerrado = cerrarCiclos_(ss) > 0;
    if (cicloCerrado) programarRecomendacion_();
  }
  return { ids: ids, rechazados: rechazados, cicloCerrado: cicloCerrado };
}

function eliminar_(ss, id) {
  if (!id) throw new Error('Falta el ID');
  for (const [nombre, col] of [[H_GASTOS, 8], [H_INGRESOS, 7]]) {
    const h = ss.getSheetByName(nombre);
    if (!h || h.getLastRow() < 2) continue;
    const ids = h.getRange(2, col, h.getLastRow() - 1, 1).getValues();
    const i = ids.findIndex(r => String(r[0]) === id);
    if (i >= 0) { h.deleteRow(i + 2); return; }
  }
  // Si no existe, ya estaba borrado: no es un error.
}

/* ============================ Catálogo ============================ */

/** Reescribe el bloque de Categorias (valores + fórmulas). Las referencias de Resumen se mantienen. */
function escribirCategorias_(ss, lista) {
  if (lista.length > MAX_CATEGORIAS) throw new Error('Máximo ' + MAX_CATEGORIAS + ' categorías.');
  const h = hoja_(ss, H_CATEGORIAS);
  const n = lista.length;
  const ultima = Math.max(h.getLastRow(), n + 1, 2);
  h.getRange(1, 1, 1, ENC_CATEGORIAS.length).setValues([ENC_CATEGORIAS]).setFontWeight('bold');
  h.getRange(2, 1, ultima - 1, ENC_CATEGORIAS.length).clearContent();
  if (!n) return;
  h.getRange(2, 1, n, 2).setValues(lista.map(c => [c.nombre, Math.round(c.presupuesto) || 0]));
  h.getRange(2, 6, n, 1).setValues(lista.map(c => [c.tipo]));
  const f = [];
  for (let r = 2; r <= n + 1; r++) {
    f.push(['=SUMIFS(Gastos!D:D,Gastos!C:C,A' + r + ',Gastos!A:A,">="&Resumen!$B$2)', '=B' + r + '-C' + r, '=IF(B' + r + '=0,0,C' + r + '/B' + r + ')']);
  }
  h.getRange(2, 3, n, 3).setFormulas(f);
  h.getRange(2, 2, n, 3).setNumberFormat(FORMATO_COP);
  h.getRange(2, 5, n, 1).setNumberFormat('0%');
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(TIPOS_CATEGORIA, true).setAllowInvalid(false).build();
  h.getRange(2, 6, n, 1).setDataValidation(regla);
}

function categoriaGuardar_(ss, p) {
  const nombre = String(p.nombre || '').trim().replace(/\s+/g, ' ');
  const anterior = String(p.nombreAnterior || '').trim();
  const presupuesto = Math.round(numero_(p.presupuesto));
  const tipo = String(p.tipo || '').toLowerCase();
  if (!nombre) throw new Error('Escribe el nombre de la categoría');
  if (nombre.length > 40) throw new Error('El nombre es muy largo (máximo 40 caracteres)');
  if (!(presupuesto >= 0)) throw new Error('Presupuesto inválido');
  if (TIPOS_CATEGORIA.indexOf(tipo) < 0) throw new Error('Tipo inválido');

  const lista = leerCategorias_(ss);
  const igual = (a, b) => a.toLowerCase() === b.toLowerCase();
  const iAnt = anterior ? lista.findIndex(c => igual(c.nombre, anterior)) : -1;
  const iDup = lista.findIndex(c => igual(c.nombre, nombre));
  if (iDup >= 0 && iDup !== iAnt) {
    if (anterior) throw new Error('Ya existe una categoría llamada "' + lista[iDup].nombre + '"');
    lista[iDup] = { nombre: lista[iDup].nombre, presupuesto: presupuesto, tipo: tipo };   // crear una que ya existe = actualizarla
  } else if (iAnt >= 0) {
    lista[iAnt] = { nombre: nombre, presupuesto: presupuesto, tipo: tipo };
    if (anterior !== nombre) renombrarEnGastos_(ss, anterior, nombre);
  } else {
    lista.push({ nombre: nombre, presupuesto: presupuesto, tipo: tipo });
  }
  escribirCategorias_(ss, lista);
}

function renombrarEnGastos_(ss, de, a) {
  const h = ss.getSheetByName(H_GASTOS);
  if (!h || h.getLastRow() < 2) return;
  const r = h.getRange(2, 3, h.getLastRow() - 1, 1);
  const v = r.getValues();
  let cambio = false;
  v.forEach(f => { if (String(f[0]) === de) { f[0] = a; cambio = true; } });
  if (cambio) r.setValues(v);
}

function categoriaEliminar_(ss, nombre) {
  const lista = leerCategorias_(ss);
  const nueva = lista.filter(c => c.nombre !== nombre);
  if (nueva.length === lista.length) return;
  if (!nueva.length) throw new Error('Debe quedar al menos una categoría');
  // Los gastos antiguos conservan su categoría; la app los sigue mostrando en el historial.
  escribirCategorias_(ss, nueva);
}

function metodosGuardar_(metodos) {
  const limpios = [...new Set((metodos || []).map(m => String(m).trim()).filter(Boolean))].slice(0, 12);
  if (!limpios.length) throw new Error('Debe quedar al menos un método de pago');
  PROPS.setProperty('METODOS', JSON.stringify(limpios));
}

/* ============================ Cierre de ciclo ============================ */

/**
 * Un ciclo va desde un salario hasta el día anterior al siguiente salario.
 * Agrega a "Cierres" los ciclos terminados que aún no estén. Devuelve cuántos agregó.
 */
function cerrarCiclos_(ss) {
  const tz = zona_(ss);
  const m = leerMovimientos_(ss);
  const salarios = [...new Set(m.ingresos.filter(i => i.tipo === 'Salario').map(i => i.fecha))].sort();
  if (salarios.length < 2) return 0;
  const h = hoja_(ss, H_CIERRES, ENC_CIERRES);
  const hechos = new Set(leerCierres_(ss).map(c => c.inicio));
  const presupuesto = leerCategorias_(ss).reduce((s, c) => s + c.presupuesto, 0);
  const ahorroCats = new Set(leerCategorias_(ss).filter(c => c.tipo === 'ahorro').map(c => c.nombre));
  const filas = [];
  for (let i = 0; i < salarios.length - 1; i++) {
    const ini = salarios[i], fin = sumarDiasIso_(salarios[i + 1], -1);
    if (hechos.has(ini)) continue;
    const g = m.gastos.filter(x => x.fecha >= ini && x.fecha <= fin);
    const ing = m.ingresos.filter(x => x.fecha >= ini && x.fecha <= fin).reduce((s, x) => s + x.monto, 0);
    const gas = g.reduce((s, x) => s + x.monto, 0);
    const porCat = {};
    g.forEach(x => { porCat[x.categoria] = (porCat[x.categoria] || 0) + x.monto; });
    const ahorro = Object.keys(porCat).filter(k => ahorroCats.has(k)).reduce((s, k) => s + porCat[k], 0);
    const tasa = ing > 0 ? (ing - gas + ahorro) / ing : 0;
    filas.push([deIso_(ini, tz), deIso_(fin, tz), diasEntreIso_(ini, fin) + 1, ing, gas, ing - gas, ahorro, tasa, presupuesto,
      JSON.stringify(porCat), '', '', '']);
  }
  if (!filas.length) return 0;
  const inicio = h.getLastRow() + 1;
  h.getRange(inicio, 1, filas.length, ENC_CIERRES.length).setValues(filas);
  const n = h.getLastRow() - 1;
  h.getRange(2, 1, n, 2).setNumberFormat(FORMATO_FECHA);
  h.getRange(2, 4, n, 4).setNumberFormat(FORMATO_COP);
  h.getRange(2, 8, n, 1).setNumberFormat('0%');
  h.getRange(2, 9, n, 1).setNumberFormat(FORMATO_COP);
  graficoCierres_(h);
  return filas.length;
}

function graficoCierres_(h) {
  h.getCharts().forEach(c => h.removeChart(c));
  const n = h.getLastRow();
  if (n < 2) return;
  const g = h.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(h.getRange(1, 1, n, 1))
    .addRange(h.getRange(1, 4, n, 3))
    .setNumHeaders(1)
    .setPosition(2, 15, 0, 0)
    .setOption('title', 'Evolución por mes (desde cada salario)')
    .setOption('series', { 2: { type: 'line' } })
    .setOption('width', 620).setOption('height', 360)
    .build();
  h.insertChart(g);
}

function programarRecomendacion_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'generarRecomendacionPendiente').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('generarRecomendacionPendiente').timeBased().after(60 * 1000).create();
}

/* ============================ Recomendación (IA o reglas) ============================ */

const ESQUEMA_RECOMENDACION = {
  type: 'object',
  additionalProperties: false,
  required: ['titulo', 'diagnostico', 'logros', 'alertas', 'acciones', 'ajustes_presupuesto', 'meta_ahorro_cop', 'mensaje'],
  properties: {
    titulo: { type: 'string', description: 'Frase corta que resume el mes' },
    diagnostico: { type: 'string', description: '2 a 4 frases sobre cómo le fue en el mes' },
    logros: { type: 'array', items: { type: 'string' } },
    alertas: { type: 'array', items: { type: 'string' } },
    acciones: {
      type: 'array',
      items: {
        type: 'object', additionalProperties: false, required: ['accion', 'impacto_cop'],
        properties: { accion: { type: 'string' }, impacto_cop: { type: 'integer' } }
      }
    },
    ajustes_presupuesto: {
      type: 'array',
      items: {
        type: 'object', additionalProperties: false, required: ['categoria', 'actual', 'sugerido', 'motivo'],
        properties: { categoria: { type: 'string' }, actual: { type: 'integer' }, sugerido: { type: 'integer' }, motivo: { type: 'string' } }
      }
    },
    meta_ahorro_cop: { type: 'integer', description: 'Cuánto ahorrar el próximo mes' },
    mensaje: { type: 'string', description: 'Un siguiente paso para evolucionar financieramente' }
  }
};

/**
 * Genera la recomendación del último ciclo cerrado. Si forzar = false, solo si aún no tiene.
 * Devuelve { fuente } o null si no hay ciclos cerrados.
 */
function generarRecomendacion_(ss, forzar) {
  const cierres = leerCierres_(ss);
  if (!cierres.length) return null;
  const ult = cierres[cierres.length - 1];
  if (ult.recomendacion && !forzar) return { fuente: ult.fuente };

  const contexto = contextoRecomendacion_(ss, ult, cierres);
  let rec = null, fuente = 'reglas';
  try {
    rec = llamarClaude_(contexto);
    if (rec) fuente = 'IA';
  } catch (e) {
    console.error('IA no disponible: ' + e);
  }
  if (!rec) rec = recomendacionPorReglas_(contexto);

  hoja_(ss, H_CIERRES, ENC_CIERRES).getRange(ult.fila, 11, 1, 3).setValues([[JSON.stringify(rec), fuente, new Date()]]);
  return { fuente: fuente };
}

function contextoRecomendacion_(ss, ciclo, cierres) {
  const cats = leerCategorias_(ss);
  const m = leerMovimientos_(ss);
  const gastos = m.gastos.filter(g => g.fecha >= ciclo.inicio && g.fecha <= ciclo.fin);
  const enviarDescripciones = PROPS.getProperty('IA_ENVIAR_DESCRIPCIONES') !== 'no';
  const porMetodo = {};
  gastos.forEach(g => { porMetodo[g.metodo || 'Sin método'] = (porMetodo[g.metodo || 'Sin método'] || 0) + g.monto; });
  return {
    moneda: 'COP',
    periodo: { inicio: ciclo.inicio, fin: ciclo.fin, dias: ciclo.dias },
    ingresos: ciclo.ingresos, gastos: ciclo.gastos, saldo: ciclo.saldo, ahorro_registrado: ciclo.ahorro,
    tasa_ahorro: Math.round(ciclo.tasaAhorro * 1000) / 10,
    categorias: cats.map(c => ({ nombre: c.nombre, tipo: c.tipo, presupuesto: c.presupuesto, gastado: Math.round(ciclo.porCategoria[c.nombre] || 0) }))
      .concat(Object.keys(ciclo.porCategoria).filter(k => !cats.some(c => c.nombre === k))
        .map(k => ({ nombre: k, tipo: 'sin catálogo', presupuesto: 0, gastado: Math.round(ciclo.porCategoria[k]) }))),
    por_metodo_pago: porMetodo,
    mayores_gastos: gastos.slice().sort((a, b) => b.monto - a.monto).slice(0, 15)
      .map(g => ({ fecha: g.fecha, categoria: g.categoria, monto: g.monto, descripcion: enviarDescripciones ? g.descripcion : '' })),
    compras_pequenas: gastos.filter(g => g.monto <= Math.max(ciclo.ingresos * 0.015, 1)).length,
    meses_anteriores: cierres.slice(-7, -1).map(c => ({ inicio: c.inicio, ingresos: c.ingresos, gastos: c.gastos, saldo: c.saldo,
      tasa_ahorro: Math.round(c.tasaAhorro * 1000) / 10 }))
  };
}

/** Llama a la API de Claude. Devuelve el objeto de recomendación o null si no hay clave. */
function llamarClaude_(contexto) {
  const key = PROPS.getProperty('ANTHROPIC_API_KEY');
  if (!key) return null;
  const sistema = 'Eres un asesor de finanzas personales para una persona que vive en Colombia y maneja pesos colombianos (COP). ' +
    'Analizas el cierre de su mes de presupuesto (que va de un salario al siguiente) y le das una recomendación práctica para ' +
    'evolucionar financieramente: primero estabilidad (no gastar más de lo que entra y fondo de emergencia), luego reducir deudas, ' +
    'luego ahorro e inversión. Escribe en español de Colombia, en un tono cercano y directo, sin juzgar. ' +
    'Usa solo los datos recibidos; si algo no se puede saber, no lo inventes. Los montos van en pesos enteros, sin centavos. ' +
    'Da como máximo 3 logros, 3 alertas, 4 acciones concretas con su impacto estimado en COP para el próximo mes y ' +
    'solo los ajustes de presupuesto que valgan la pena. Las categorías de tipo "fija" son pagos mensuales que no se recortan fácilmente.';
  const cuerpo = {
    model: PROPS.getProperty('IA_MODELO') || IA_MODELO_POR_DEFECTO,
    max_tokens: 8000,
    system: sistema,
    messages: [{ role: 'user', content: 'Datos del mes cerrado (JSON):\n' + JSON.stringify(contexto) }],
    output_config: { effort: PROPS.getProperty('IA_ESFUERZO') || 'medium', format: { type: 'json_schema', schema: ESQUEMA_RECOMENDACION } },
    fallbacks: 'default'
  };
  const pedir = (conRespaldo) => {
    const headers = { 'x-api-key': key, 'anthropic-version': '2026-10-21' };
    const body = Object.assign({}, cuerpo);
    if (conRespaldo) headers['anthropic-beta'] = 'server-side-fallback-2026-07-01'; else delete body.fallbacks;
    return UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true
    });
  };
  let r = pedir(true);
  if (r.getResponseCode() === 400) r = pedir(false);   // por si la cuenta no admite el modelo de respaldo
  const codigo = r.getResponseCode();
  if (codigo !== 200) throw new Error('API de Claude respondió ' + codigo + ': ' + r.getContentText().slice(0, 300));
  const j = JSON.parse(r.getContentText());
  if (j.stop_reason === 'refusal' || j.stop_reason === 'max_tokens') throw new Error('Respuesta incompleta: ' + j.stop_reason);
  const texto = (j.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  const rec = JSON.parse(texto);
  rec.modelo = j.model;
  return rec;
}

/** Recomendación sin IA, para cuando no hay clave o la API falla. */
function recomendacionPorReglas_(c) {
  const dinero = n => '$' + Math.round(n).toLocaleString('es-CO');
  const logros = [], alertas = [], acciones = [], ajustes = [];
  const opcionales = c.categorias.filter(x => x.tipo === 'opcional' && x.gastado > 0).sort((a, b) => b.gastado - a.gastado);
  const excedidas = c.categorias.filter(x => x.presupuesto > 0 && x.gastado > x.presupuesto && x.tipo !== 'ahorro');
  const ahorro = c.categorias.filter(x => x.tipo === 'ahorro');
  const metaAhorroCat = ahorro.reduce((s, x) => s + x.presupuesto, 0);

  if (c.saldo >= 0) logros.push('Cerraste el mes con saldo positivo de ' + dinero(c.saldo) + '.');
  else alertas.push('Gastaste ' + dinero(-c.saldo) + ' más de lo que ingresó.');
  if (c.tasa_ahorro >= 10) logros.push('Tu tasa de ahorro fue ' + c.tasa_ahorro + ' %.');
  excedidas.slice(0, 3).forEach(x => alertas.push(x.nombre + ' superó su presupuesto en ' + dinero(x.gastado - x.presupuesto) + '.'));
  if (!excedidas.length) logros.push('Ninguna categoría superó su presupuesto.');

  if (opcionales[0]) acciones.push({ accion: 'Baja ' + opcionales[0].nombre + ' un 15 % el próximo mes.', impacto_cop: Math.round(opcionales[0].gastado * 0.15) });
  const metaAhorro = Math.max(metaAhorroCat, Math.round(c.ingresos * 0.1 / 10000) * 10000);
  acciones.push({ accion: 'Aparta ' + dinero(metaAhorro) + ' para ahorro el mismo día que recibas el salario.', impacto_cop: metaAhorro });
  if (c.compras_pequenas >= 8) acciones.push({ accion: 'Revisa tus ' + c.compras_pequenas + ' compras pequeñas: elimina la mitad.', impacto_cop: 0 });
  excedidas.filter(x => x.tipo !== 'opcional').forEach(x =>
    ajustes.push({ categoria: x.nombre, actual: x.presupuesto, sugerido: Math.ceil(x.gastado / 10000) * 10000, motivo: 'Es esencial y el presupuesto se quedó corto.' }));

  return {
    titulo: c.saldo >= 0 ? 'Mes en positivo' : 'Mes en rojo: hay que ajustar',
    diagnostico: 'Del ' + c.periodo.inicio + ' al ' + c.periodo.fin + ' ingresaron ' + dinero(c.ingresos) + ' y gastaste ' + dinero(c.gastos) +
      '. Tasa de ahorro: ' + c.tasa_ahorro + ' %.',
    logros: logros, alertas: alertas, acciones: acciones, ajustes_presupuesto: ajustes,
    meta_ahorro_cop: metaAhorro,
    mensaje: c.saldo < 0 ? 'Primer objetivo: que el próximo mes los gastos no superen los ingresos.'
      : 'Siguiente paso: completar un fondo de emergencia de 3 meses de gastos (' + dinero(c.gastos * 3) + ').'
  };
}
