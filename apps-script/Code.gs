/**
 * Canje Rinti · Backend (Google Apps Script)
 *
 * Recibe los registros de la app de campo, guarda los datos en este Google Sheet
 * y las fotos en una carpeta de Drive. También entrega los datos al dashboard.
 *
 * Instalación: ver README.md (pasos 1 a 4). Correr setup() una sola vez.
 */

const TZ = 'America/Lima';
const SACOS_POR_VALE = 8;
const KG = { ricocan: 22, ricocat: 15 };
const P = PropertiesService.getScriptProperties();

const MERCADOS = [
  'Fiori', 'Unicachi Norte Mayorista', 'Unicachi Norte Minorista', 'Unicachi Sur',
  'Huamantanga', 'Ciudad de Dios', 'Parada', 'Caquetá', 'Mercado Productores',
];

const ESQUEMA = {
  Asistencia: ['id', 'registrado', 'fecha', 'hora', 'impulsadora', 'mercado', 'marca', 'lat', 'lng', 'precision_m', 'foto'],
  Pedidos: ['id', 'registrado', 'fecha', 'hora', 'impulsadora', 'mercado', 'puesto', 'cliente', 'celular',
    'tipo_comprobante', 'n_comprobante', 'sacos_ricocan', 'sacos_ricocat', 'sacos_total', 'kg_total',
    'es_canje', 'vales', 'codigos_vale', 'foto_comprobante', 'foto_entrega', 'firma', 'compra_lima',
    'consentimiento', 'cumple_mecanica', 'alertas', 'observaciones', 'validacion', 'lat', 'lng'],
  Incidencias: ['id', 'registrado', 'fecha', 'hora', 'impulsadora', 'mercado', 'tipo_incidencia', 'detalle', 'foto', 'lat', 'lng'],
  Cierres: ['id', 'registrado', 'fecha', 'hora', 'supervisor', 'mercado', 'pavos_inicio', 'pavos_entregados',
    'pavos_saldo', 'vales_reportados', 'observaciones'],
  Vales: ['codigo', 'pedido_id', 'fecha', 'impulsadora', 'mercado'],
  Equipo: ['nombre', 'pin', 'rol', 'mercado', 'activo'],
  Config: ['clave', 'valor', 'nota'],
};

// Columnas que se guardan como texto (para que Sheets no las convierta en fecha/número).
const COLUMNAS_TEXTO = ['id', 'fecha', 'hora', 'celular', 'n_comprobante', 'codigos_vale', 'codigo', 'pedido_id', 'pin', 'valor'];

/* ------------------------------------------------------------------ */
/* Instalación                                                         */
/* ------------------------------------------------------------------ */

function setup() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.create('Canje Rinti · Base de datos');
  P.setProperty('SS_ID', ss.getId());

  Object.keys(ESQUEMA).forEach(function (nombre) {
    const cols = ESQUEMA[nombre];
    const sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cols.length).setValues([cols])
      .setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    cols.forEach(function (c, i) {
      if (COLUMNAS_TEXTO.indexOf(c) >= 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });

  // Borra la hoja vacía por defecto
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1 && sh.getLastRow() === 0) ss.deleteSheet(sh);
  });

  // Validación del supervisor: lista desplegable
  const ped = ss.getSheetByName('Pedidos');
  const colVal = ESQUEMA.Pedidos.indexOf('validacion') + 1;
  ped.getRange(2, colVal, ped.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Pendiente', 'Válido', 'Observado'], true).build());

  const cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 5, 3).setValues([
      ['meta_diaria_por_impulsadora', '6', 'Canjes por impulsadora por día'],
      ['impulsadoras_planificadas', '9', 'Para calcular la meta total y la asistencia'],
      ['hora_ingreso', '09:00', 'Hora de ingreso (tolerancia de 10 min para tardanza)'],
      ['fecha_inicio', '', 'AAAA-MM-DD · inicio de campaña'],
      ['fecha_fin', '', 'AAAA-MM-DD · fin de campaña'],
    ]);
  }

  const eq = ss.getSheetByName('Equipo');
  if (eq.getLastRow() < 2) {
    const filas = MERCADOS.map(function (m, i) {
      return ['Impulsadora ' + (i + 1), pinAleatorio_(), 'impulsadora', m, 'SI'];
    });
    filas.push(['Supervisor 1', pinAleatorio_(), 'supervisor', '', 'SI']);
    eq.getRange(2, 1, filas.length, 5).setValues(filas);
  }

  if (!P.getProperty('FOLDER_ID')) {
    const folder = DriveApp.createFolder('Canje Rinti · Evidencias');
    try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
    P.setProperty('FOLDER_ID', folder.getId());
  }
  if (!P.getProperty('DASH_TOKEN')) {
    P.setProperty('DASH_TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  }
  asegurarAccesos_();
  verAccesos();
}

/** Muestra en el registro de ejecución el token del dashboard y los enlaces. */
function verAccesos() {
  Logger.log('Google Sheet: ' + ss_().getUrl());
  Logger.log('Carpeta de fotos: https://drive.google.com/drive/folders/' + P.getProperty('FOLDER_ID'));
  Logger.log('Usuarios y contraseñas: pestaña Config del Google Sheet (app_usuario, app_clave, cliente_usuario, cliente_clave)');
  Logger.log('Token del dashboard (k): ' + P.getProperty('DASH_TOKEN'));
  Logger.log('Link del cliente: https://TU-USUARIO.github.io/rinti-canje/dashboard.html?k=' + P.getProperty('DASH_TOKEN'));
}

/** Cambia el token si el link del dashboard se filtró. */
function renovarTokenDashboard() {
  P.setProperty('DASH_TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  verAccesos();
}

/* ------------------------------------------------------------------ */
/* Endpoints                                                           */
/* ------------------------------------------------------------------ */

function doGet(e) {
  const accion = String((e && e.parameter && e.parameter.action) || '').toLowerCase();
  try {
    if (accion === 'config') {
      asegurarAccesos_();
      return out_({ ok: true, equipo: equipoPublico_(), config: configPublica_() });
    }
    if (accion === 'data') {
      if (!e.parameter.k || e.parameter.k !== P.getProperty('DASH_TOKEN')) {
        return out_({ ok: false, code: 'auth', error: 'Link no autorizado' });
      }
      const pedidos = leer_('Pedidos').map(function (p) {
        delete p.celular; delete p.consentimiento; // el cliente (Rinti) no ve datos de contacto
        return p;
      });
      return out_({
        ok: true,
        generado: ahora_(),
        config: configPublica_(),
        equipo: equipoPublico_(),
        asistencia: leer_('Asistencia'),
        pedidos: pedidos,
        incidencias: leer_('Incidencias'),
        cierres: leer_('Cierres'),
      });
    }
    return out_({ ok: true, servicio: 'Canje Rinti' });
  } catch (err) {
    return out_({ ok: false, code: 'server', error: String(err && err.message || err) });
  }
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) {
    return out_({ ok: false, code: 'val', error: 'Datos inválidos' });
  }
  try {
    asegurarAccesos_();
    if (b.tipo === 'login_cliente') {
      const c = config_();
      const ok = limpio_(b.usuario).toLowerCase() === String(c.cliente_usuario || '').toLowerCase()
        && String(b.clave || '') === String(c.cliente_clave || '');
      return out_(ok ? { ok: true, k: P.getProperty('DASH_TOKEN') } : { ok: false, code: 'auth', error: 'Usuario o contraseña incorrectos' });
    }
    const u = auth_(b);
    if (!u) return out_({ ok: false, code: 'auth', error: 'Usuario, contraseña o nombre incorrectos' });
    if (b.tipo === 'login') return out_({ ok: true, usuario: u });
    const fn = { asistencia: asistencia_, pedido: pedido_, incidencia: incidencia_, cierre: cierre_ }[b.tipo];
    if (!fn) return out_({ ok: false, code: 'val', error: 'Tipo de registro desconocido' });
    if (!b.id) return out_({ ok: false, code: 'val', error: 'Falta el id del registro' });
    return out_(fn(b, u));
  } catch (err) {
    return out_({ ok: false, code: err && err.val ? 'val' : 'server', error: String(err && err.message || err) });
  }
}

/* ------------------------------------------------------------------ */
/* Registros                                                           */
/* ------------------------------------------------------------------ */

function asistencia_(b, u) {
  if (existeId_('Asistencia', b.id)) return { ok: true, duplicado: true };
  if (!limpio_(b.mercado)) throw val_('Falta el mercado');
  if (['Entrada', 'Salida'].indexOf(b.marca) < 0) throw val_('Marca inválida');
  if (!b.foto) throw val_('Falta la selfie');

  const t = cuando_(b.capturado);
  const subidas = [];
  const foto = foto_(b.foto, t.fecha + '_' + u.nombre + '_' + b.marca, t.fecha, subidas);

  return conLock_(function () {
    if (existeId_('Asistencia', b.id)) { borrar_(subidas); return { ok: true, duplicado: true }; }
    agregar_('Asistencia', {
      id: b.id, registrado: ahora_(), fecha: t.fecha, hora: t.hora, impulsadora: u.nombre,
      mercado: limpio_(b.mercado), marca: b.marca, lat: num_(b.lat), lng: num_(b.lng),
      precision_m: num_(b.acc), foto: foto,
    });
    return { ok: true };
  });
}

function pedido_(b, u) {
  if (existeId_('Pedidos', b.id)) return { ok: true, duplicado: true };

  const rc = entero_(b.ricocan), rt = entero_(b.ricocat), total = rc + rt;
  const mercado = limpio_(b.mercado);
  if (!mercado) throw val_('Falta el mercado');
  if (total <= 0) throw val_('Indica la cantidad de sacos');
  if (!limpio_(b.cliente)) throw val_('Falta el nombre del cliente o negocio');
  if (/productores/i.test(mercado) && !limpio_(b.puesto)) throw val_('Elige el puesto de Mercado Productores');

  const vales = Math.floor(total / SACOS_POR_VALE);
  const esCanje = vales > 0;
  const codigos = (b.codigos || []).map(normCodigo_).filter(String);
  const nComp = limpio_(b.n_comprobante);
  const alertas = [];

  if (esCanje) {
    if (codigos.length !== vales) throw val_('Ingresa ' + vales + ' código(s) de vale');
    if (new Set(codigos).size !== codigos.length) throw val_('Hay códigos de vale repetidos');
    if (!nComp) throw val_('Falta el número de boleta o factura');
    if (!b.foto_comprobante) throw val_('Falta la foto de la boleta o factura');
    if (!b.firma) throw val_('Falta la firma del cliente');
    if (!b.foto_entrega) throw val_('Falta la foto del cliente recibiendo el vale');
    if (!b.compra_lima) throw val_('El canje solo aplica a compras en Lima');
    const usados = codigos.filter(valeUsado_);
    if (usados.length) throw val_('Vale ya registrado antes: ' + usados.join(', '));
  }
  if (limpio_(b.celular) && !b.consentimiento) throw val_('Falta el consentimiento del cliente para guardar su celular');
  if (nComp && comprobanteUsado_(nComp)) alertas.push('Comprobante repetido');
  if (!b.lat) alertas.push('Sin GPS');

  const t = cuando_(b.capturado);
  const base = t.fecha + '_' + t.hora.replace(':', '') + '_' + u.nombre;
  const subidas = [];
  const fotoComp = foto_(b.foto_comprobante, base + '_comprobante', t.fecha, subidas);
  const fotoEnt = foto_(b.foto_entrega, base + '_entrega', t.fecha, subidas);
  const firma = foto_(b.firma, base + '_firma', t.fecha, subidas);

  return conLock_(function () {
    if (existeId_('Pedidos', b.id)) { borrar_(subidas); return { ok: true, duplicado: true }; }
    const usados = codigos.filter(valeUsado_);
    if (usados.length) { borrar_(subidas); throw val_('Vale ya registrado antes: ' + usados.join(', ')); }

    agregar_('Pedidos', {
      id: b.id, registrado: ahora_(), fecha: t.fecha, hora: t.hora, impulsadora: u.nombre, mercado: mercado,
      puesto: limpio_(b.puesto), cliente: limpio_(b.cliente), celular: limpio_(b.celular),
      tipo_comprobante: limpio_(b.tipo_comprobante), n_comprobante: nComp,
      sacos_ricocan: rc, sacos_ricocat: rt, sacos_total: total, kg_total: rc * KG.ricocan + rt * KG.ricocat,
      es_canje: esCanje ? 'SÍ' : 'NO', vales: vales, codigos_vale: codigos.join(', '),
      foto_comprobante: fotoComp, foto_entrega: fotoEnt, firma: firma,
      compra_lima: b.compra_lima ? 'SÍ' : 'NO', consentimiento: b.consentimiento ? 'SÍ' : 'NO',
      cumple_mecanica: !esCanje ? 'N/A' : (alertas.indexOf('Comprobante repetido') >= 0 ? 'REVISAR' : 'SÍ'),
      alertas: alertas.join(', '), observaciones: limpio_(b.observaciones),
      validacion: esCanje ? 'Pendiente' : '', lat: num_(b.lat), lng: num_(b.lng),
    });
    codigos.forEach(function (c) {
      agregar_('Vales', { codigo: c, pedido_id: b.id, fecha: t.fecha, impulsadora: u.nombre, mercado: mercado });
    });
    return { ok: true, vales: vales, alertas: alertas };
  });
}

function incidencia_(b, u) {
  if (existeId_('Incidencias', b.id)) return { ok: true, duplicado: true };
  if (!limpio_(b.mercado)) throw val_('Falta el mercado');
  if (!limpio_(b.tipo_incidencia)) throw val_('Elige el tipo de incidencia');
  if (!limpio_(b.detalle)) throw val_('Describe la incidencia');

  const t = cuando_(b.capturado);
  const subidas = [];
  const foto = foto_(b.foto, t.fecha + '_' + u.nombre + '_incidencia', t.fecha, subidas);
  return conLock_(function () {
    if (existeId_('Incidencias', b.id)) { borrar_(subidas); return { ok: true, duplicado: true }; }
    agregar_('Incidencias', {
      id: b.id, registrado: ahora_(), fecha: t.fecha, hora: t.hora, impulsadora: u.nombre,
      mercado: limpio_(b.mercado), tipo_incidencia: limpio_(b.tipo_incidencia), detalle: limpio_(b.detalle),
      foto: foto, lat: num_(b.lat), lng: num_(b.lng),
    });
    return { ok: true };
  });
}

function cierre_(b, u) {
  if (u.rol !== 'supervisor') throw val_('Solo el supervisor puede enviar el cierre');
  if (existeId_('Cierres', b.id)) return { ok: true, duplicado: true };
  if (!limpio_(b.mercado)) throw val_('Falta el mercado');
  const ini = entero_(b.pavos_inicio), ent = entero_(b.pavos_entregados);
  if (ent > ini) throw val_('Los pavos entregados no pueden ser más que los del inicio');
  const t = cuando_(b.capturado);
  return conLock_(function () {
    if (existeId_('Cierres', b.id)) return { ok: true, duplicado: true };
    agregar_('Cierres', {
      id: b.id, registrado: ahora_(), fecha: t.fecha, hora: t.hora, supervisor: u.nombre,
      mercado: limpio_(b.mercado), pavos_inicio: ini, pavos_entregados: ent, pavos_saldo: ini - ent,
      vales_reportados: entero_(b.vales_reportados), observaciones: limpio_(b.observaciones),
    });
    return { ok: true };
  });
}

/* ------------------------------------------------------------------ */
/* Utilidades                                                          */
/* ------------------------------------------------------------------ */

function ss_() { return SpreadsheetApp.openById(P.getProperty('SS_ID')); }
function hoja_(n) { return ss_().getSheetByName(n); }

function leer_(nombre) {
  const sh = hoja_(nombre);
  if (sh.getLastRow() < 2) return [];
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  const h = v.shift();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const o = {};
    h.forEach(function (k, i) { o[k] = r[i]; });
    return o;
  });
}

function agregar_(nombre, obj) {
  hoja_(nombre).appendRow(ESQUEMA[nombre].map(function (k) {
    return obj[k] === undefined || obj[k] === null ? '' : obj[k];
  }));
}

function buscarEnColumna_(nombre, columna, valor) {
  const sh = hoja_(nombre);
  if (sh.getLastRow() < 2 || !valor) return false;
  const col = ESQUEMA[nombre].indexOf(columna) + 1;
  return !!sh.getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(String(valor)).matchEntireCell(true).findNext();
}
function existeId_(nombre, id) { return buscarEnColumna_(nombre, 'id', id); }
function valeUsado_(codigo) { return buscarEnColumna_('Vales', 'codigo', codigo); }
function comprobanteUsado_(n) { return buscarEnColumna_('Pedidos', 'n_comprobante', n); }

/**
 * Acceso a la app de campo: usuario y contraseña generales (hoja Config) + nombre de la promotora.
 * También acepta el PIN individual de la hoja Equipo, para cuando se pase a accesos individuales.
 */
function auth_(b) {
  if (!b.nombre) return null;
  const r = leer_('Equipo').find(function (x) { return x.nombre === b.nombre && !/^no$/i.test(x.activo); });
  if (!r) return null;
  const c = config_();
  const okGeneral = !!(b.usuario && b.clave && c.app_usuario && c.app_clave
    && limpio_(b.usuario).toLowerCase() === String(c.app_usuario).toLowerCase() && String(b.clave) === String(c.app_clave));
  const okPin = !!(b.pin && r.pin && String(r.pin).trim() === String(b.pin).trim());
  if (!okGeneral && !okPin) return null;
  return { nombre: r.nombre, rol: String(r.rol || 'impulsadora').toLowerCase(), mercado: r.mercado };
}

/** Crea en la hoja Config los usuarios y contraseñas generales si aún no existen. */
function asegurarAccesos_() {
  const faltan = function () {
    const c = config_();
    const f = [];
    if (!c.app_usuario) f.push(['app_usuario', 'promotoras', 'Usuario general de la app de campo']);
    if (!c.app_clave) f.push(['app_clave', claveAleatoria_(), 'Contraseña general de la app de campo (puedes cambiarla aquí)']);
    if (!c.cliente_usuario) f.push(['cliente_usuario', 'rinti', 'Usuario del dashboard del cliente']);
    if (!c.cliente_clave) f.push(['cliente_clave', claveAleatoria_(), 'Contraseña del dashboard del cliente (puedes cambiarla aquí)']);
    return f;
  };
  if (!faltan().length) return;
  conLock_(function () {
    const f = faltan();
    if (!f.length) return;
    const sh = hoja_('Config');
    sh.getRange(sh.getLastRow() + 1, 1, f.length, 3).setValues(f);
  });
}

function claveAleatoria_() {
  const abc = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
}

/** Config sin usuarios ni contraseñas, para enviar a las páginas. */
function configPublica_() {
  const c = config_();
  Object.keys(c).forEach(function (k) { if (/usuario|clave|pin|token/i.test(k)) delete c[k]; });
  return c;
}

function equipoPublico_() {
  return leer_('Equipo').filter(function (x) { return !/^no$/i.test(x.activo); })
    .map(function (x) { return { nombre: x.nombre, rol: String(x.rol || 'impulsadora').toLowerCase(), mercado: x.mercado }; });
}

function config_() {
  const o = {};
  leer_('Config').forEach(function (r) { if (r.clave) o[r.clave] = r.valor; });
  return o;
}

function foto_(dataUrl, nombre, fecha, subidas) {
  if (!dataUrl) return '';
  const m = String(dataUrl).match(/^data:(image\/[a-z+]+);base64,(.+)$/);
  if (!m) throw val_('Formato de foto inválido');
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1],
    String(nombre).replace(/[^\w\-áéíóúñÁÉÍÓÚÑ]+/g, '_') + '.' + ext);
  const f = carpetaDia_(fecha).createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
  subidas.push(f);
  return 'https://drive.google.com/file/d/' + f.getId() + '/view';
}

function carpetaDia_(fecha) {
  const raiz = DriveApp.getFolderById(P.getProperty('FOLDER_ID'));
  const it = raiz.getFoldersByName(fecha);
  return it.hasNext() ? it.next() : raiz.createFolder(fecha);
}

function borrar_(archivos) { archivos.forEach(function (f) { try { f.setTrashed(true); } catch (e) { } }); }

function conLock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(30000);
  try { return fn(); } finally { l.releaseLock(); }
}

function cuando_(iso) {
  let d = iso ? new Date(iso) : new Date();
  if (isNaN(d.getTime()) || d.getTime() - Date.now() > 3600 * 1000) d = new Date();
  return { fecha: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), hora: Utilities.formatDate(d, TZ, 'HH:mm') };
}
function ahora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function limpio_(s) { return String(s === undefined || s === null ? '' : s).trim().slice(0, 500); }
function normCodigo_(s) { return limpio_(s).toUpperCase().replace(/\s+/g, ''); }
function entero_(n) { const x = parseInt(n, 10); return isNaN(x) || x < 0 ? 0 : x; }
function num_(n) { const x = Number(n); return n === '' || n === null || n === undefined || isNaN(x) ? '' : x; }
function pinAleatorio_() { return String(Math.floor(1000 + Math.random() * 9000)); }
function val_(msg) { const e = new Error(msg); e.val = true; return e; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
