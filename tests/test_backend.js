// Pruebas del backend (apps-script/Code.gs) con servicios de Google simulados.
// Correr en macOS:  osascript -l JavaScript tests/test_backend.js
ObjC.import('Foundation');
function readFile(p) { return ObjC.unwrap($.NSString.stringWithContentsOfFileEncodingError(p, $.NSUTF8StringEncoding, null)); }
const here = ObjC.unwrap($.NSFileManager.defaultManager.currentDirectoryPath);
const src = readFile(here + '/apps-script/Code.gs');

// ---- simulación de servicios ----
const props = {}; let idn = 0; const files = {}; const logs = [];
const PropertiesService = { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) };
function mkSheet(name) {
  const rows = [];
  const sh = {
    name, rows,
    getLastRow: () => rows.length, getLastColumn: () => Math.max(0, ...rows.map(r => r.length)), getMaxRows: () => 1000,
    setFrozenRows() {},
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      const rg = {
        setValues(v) { v.forEach((row, i) => { rows[r - 1 + i] = rows[r - 1 + i] || []; row.forEach((x, j) => rows[r - 1 + i][c - 1 + j] = x); }); return rg; },
        setValue(x) { rows[r - 1] = rows[r - 1] || []; rows[r - 1][c - 1] = x; return rg; },
        getDisplayValues() {
          const out = [];
          for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const x = (rows[r - 1 + i] || [])[c - 1 + j]; row.push(x === undefined || x === null ? '' : String(x)); } out.push(row); }
          return out;
        },
        setFontWeight() { return rg; }, setBackground() { return rg; }, setFontColor() { return rg; }, setNumberFormat() { return rg; }, setDataValidation() { return rg; },
        createTextFinder(t) { return { matchEntireCell() { return { findNext() { for (let i = 0; i < nr; i++) { const x = (rows[r - 1 + i] || [])[c - 1]; if (String(x) === t) return {}; } return null; } }; } }; },
      };
      return rg;
    },
    appendRow(v) { rows.push(v.slice()); },
  };
  return sh;
}
const sheets = {};
const ss = { getId: () => 'SS1', getUrl: () => 'url', getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = mkSheet(n)), getSheets: () => Object.values(sheets), deleteSheet(s) { delete sheets[s.name]; } };
const dv = { requireValueInList() { return dv; }, build() { return {}; } };
const SpreadsheetApp = { getActiveSpreadsheet: () => ss, openById: () => ss, create: () => ss, newDataValidation: () => dv };
function mkFolder(id) {
  const subs = {};
  return {
    getId: () => id, setSharing() {},
    getFoldersByName(n) { let done = !subs[n]; return { hasNext: () => !done, next: () => { done = true; return subs[n]; } }; },
    createFolder(n) { return subs[n] = mkFolder('F' + (++idn)); },
    createFile(b) { const fid = 'file' + (++idn); files[fid] = { name: b.name, trashed: false }; return { getId: () => fid, setSharing() {}, setTrashed(t) { files[fid].trashed = t; } }; },
  };
}
const root = mkFolder('ROOT');
const DriveApp = { createFolder: () => root, getFolderById: () => root, Access: {}, Permission: {} };
const Utilities = {
  formatDate(d, tz, f) { const s = d.toLocaleString('sv-SE', { timeZone: tz }); return f === 'yyyy-MM-dd' ? s.slice(0, 10) : f === 'HH:mm' ? s.slice(11, 16) : s; },
  base64Decode: s => s, newBlob: (d, t, n) => ({ name: n, type: t }), getUuid: () => 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
};
const LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
const ContentService = { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType() { return { s }; } }) };
const Logger = { log: s => logs.push(s) };
eval(src);

// ---- pruebas ----
var sh;
const res = []; const ok = (name, cond, extra) => res.push((cond ? 'PASS ' : 'FAIL ') + name + (extra ? ' -> ' + extra : ''));
const post = b => JSON.parse(doPost({ postData: { contents: JSON.stringify(b) } }).s);
const get = p => JSON.parse(doGet({ parameter: p }).s);
setup();
const H = sheets.Equipo.rows[0];
const eq = sheets.Equipo.rows.slice(1).map(r => Object.fromEntries(H.map((h, i) => [h, r[i]])));
const cfg = Object.fromEntries(sheets.Config.rows.slice(1).map(r => [r[0], r[1]]));
const fiori = eq.find(e => e.mercado === 'Fiori'), cdd = eq.find(e => e.mercado === 'Ciudad de Dios'), prod = eq.find(e => e.mercado === 'Mercado Productores'), sup = eq.find(e => e.rol === 'supervisor');
ok('9 usuarios por mercado + 2 supervisores + coordinador', eq.length === 12 && eq.filter(e => e.usuario && e.contrasena && e.contrasena.length === 8).length === 12, eq.map(e => e.usuario).join(','));
const coord = eq.find(e => e.usuario === 'coordinador'), sup2 = eq.find(e => e.usuario === 'supervisor2');
ok('coordinador y supervisor2 con su rol', coord && coord.rol === 'coordinador' && coord.nombre === 'Coordinador de Trade' && sup2 && sup2.rol === 'supervisor');
ok('usuarios con nombre de mercado', fiori.usuario === 'fiori' && prod.usuario === 'productores' && sup.usuario === 'supervisor');
ok('config publica sin claves y con reglas', (() => { const c = get({ action: 'config' }); return !JSON.stringify(c).match(/contrasena|cliente_clave|"pin"/) && c.sacos_por_mercado['Fiori'] === 17; })());
ok('login cliente con espacios', post({ tipo: 'login_cliente', usuario: ' Rinti ', clave: cfg.cliente_clave + ' \n' }).k === props.DASH_TOKEN);
ok('login cliente malo', post({ tipo: 'login_cliente', usuario: 'rinti', clave: 'x' }).code === 'auth');
const F = { usuario: 'FIORI', clave: fiori.contrasena + ' ' };
let r = post({ ...F, tipo: 'login' });
ok('login por mercado', r.ok && r.usuario.mercado === 'Fiori' && r.usuario.nombre === fiori.nombre, JSON.stringify(r));
ok('login con clave de otro mercado', post({ usuario: 'fiori', clave: cdd.contrasena, tipo: 'login' }).code === 'auth');
const img = 'data:image/jpeg;base64,AAAA';
ok('asistencia queda en su mercado aunque mande otro', post({ ...F, tipo: 'asistencia', id: 'a1', mercado: 'Parada', marca: 'Entrada', foto: img }).ok && sheets.Asistencia.rows[1][5] === 'Fiori');
const canje = (extra) => ({ ...F, tipo: 'pedido', cliente: 'Bodega', tipo_comprobante: 'Boleta', foto_comprobante: img, foto_entrega: img, firma: img, compra_lima: true, consentimiento: true, lat: -12, ...extra });
r = post(canje({ id: 'p1', ricocan: 10, ricocat: 6, n_comprobante: 'B1', codigos: ['1000001'], dni: '12345678' }));
ok('Fiori 16 sacos no es canje (pide 17)', r.ok && r.vales === 0, JSON.stringify(r));
r = post(canje({ id: 'p2', ricocan: 10, ricocat: 7, n_comprobante: 'B2', codigos: ['1000002'], dni: '12345678' }));
ok('Fiori 17 sacos mixtos = 1 vale', r.ok && r.vales === 1, JSON.stringify(r));
r = post(canje({ id: 'p3', ricocan: 40, n_comprobante: 'B3', codigos: ['1000003'], dni: '12345678' }));
ok('mismo DNI el mismo día: rechazado', r.code === 'val' && /ya recibió un vale hoy/.test(r.error), JSON.stringify(r));
r = post(canje({ id: 'p4', ricocan: 34, n_comprobante: 'B4', codigos: ['1000004', '1000005'], dni: '87654321' }));
ok('34 sacos sigue siendo 1 vale (un solo código)', r.code === 'val' && /código del vale/.test(r.error), JSON.stringify(r));
r = post(canje({ id: 'p5', ricocan: 34, n_comprobante: 'B5', codigos: ['1000004'], dni: '87654321' }));
ok('34 sacos con 1 código: ok', r.ok && r.vales === 1);
ok('canje sin DNI', post(canje({ id: 'p6', ricocan: 20, n_comprobante: 'B6', codigos: ['1000006'] })).error === 'Falta el DNI del cliente');
ok('DNI inválido', /8 dígitos/.test(post(canje({ id: 'p7', ricocan: 20, n_comprobante: 'B7', codigos: ['1000007'], dni: '123' })).error));
ok('DNI sin consentimiento', /consentimiento/.test(post(canje({ id: 'p8', ricocan: 20, n_comprobante: 'B8', codigos: ['1000008'], dni: '11112222', consentimiento: false })).error));
const C = { usuario: 'ciudaddedios', clave: cdd.contrasena };
r = post({ ...canje({ id: 'p9', ricocan: 3, ricocat: 2, n_comprobante: 'B9', codigos: ['1000009'], dni: '22223333' }), ...C });
ok('Ciudad de Dios 5 sacos = vale', r.ok && r.vales === 1, JSON.stringify(r));
ok('verificar DNI ya usado', (() => { const v = post({ ...C, tipo: 'verificar_dni', dni: '12345678' }); return v.ok && v.disponible === false && /Fiori/.test(v.mensaje); })());
ok('verificar DNI libre', post({ ...C, tipo: 'verificar_dni', dni: '99998888' }).disponible === true);
const S = { usuario: 'supervisor', clave: sup.contrasena };
r = post({ ...canje({ id: 'p10', ricocan: 24, n_comprobante: 'B10', codigos: ['1000010'], dni: '33334444', puesto: 'Sumaq Sara' }), ...S, mercado: 'Mercado Productores' });
ok('Productores 24 sacos no es canje', r.ok && r.vales === 0, JSON.stringify(r));
r = post({ ...canje({ id: 'p11', ricocan: 25, n_comprobante: 'B11', codigos: ['1000011'], dni: '33334444', puesto: 'Sumaq Sara' }), ...S, mercado: 'Mercado Productores' });
ok('Productores 25 sacos = canje (supervisor elige mercado)', r.ok && r.vales === 1);
ok('supervisor con mercado inválido', /Mercado no válido/.test(post({ ...canje({ id: 'p12', ricocan: 25, n_comprobante: 'B12', codigos: ['1000012'], dni: '44445555' }), ...S, mercado: 'Otro' }).error));
ok('cierre del supervisor', post({ ...S, tipo: 'cierre', id: 'c1', mercado: 'Fiori', pavos_inicio: 10, pavos_entregados: 1, vales_reportados: 1 }).ok);
ok('cierre de impulsadora: no', post({ ...F, tipo: 'cierre', id: 'c2', mercado: 'Fiori', pavos_inicio: 10, pavos_entregados: 1 }).code === 'val');
const d = get({ action: 'data', k: props.DASH_TOKEN });
const p2 = d.pedidos.find(p => p.id === 'p2');
ok('dashboard: DNI enmascarado y sin celular', p2.dni === '•••••678' && !d.pedidos.some(p => 'celular' in p), p2.dni);
ok('dashboard: sacos requeridos guardados', p2.sacos_requeridos === '17' && p2.es_canje === 'SÍ');
ok('dashboard: reglas en la respuesta', d.sacos_por_mercado['Mercado Productores'] === 25);
ok('dashboard: sin token no hay datos', get({ action: 'data', k: 'x' }).code === 'auth');
// ---- coordinador, panel y regla del vale ----
sh = n => sheets[n].rows.slice(1).map(r => Object.fromEntries(sheets[n].rows[0].map((h, i) => [h, r[i]])));
const eq2 = sh('Equipo'); const K = u => ({ usuario: u, clave: eq2.find(e => e.usuario === u).contrasena });
r = post({ ...K('coordinador'), tipo: 'login' });
ok('login coordinador', r.ok && r.usuario.rol === 'coordinador', JSON.stringify(r));
ok('migración no duplica coordinador', eq2.filter(e => e.rol === 'coordinador').length === 1 && eq2.filter(e => /supervisor\s*2/i.test(e.nombre)).length === 1, eq2.map(e => e.nombre).join(','));
ok('coordinador hace cierre', post({ ...K('coordinador'), tipo: 'cierre', id: 'c3', mercado: 'Parada', pavos_inicio: 5, pavos_entregados: 0, vales_reportados: 0 }).ok);
r = post({ ...K('coordinador'), tipo: 'panel' });
const fi = r.filas && r.filas.find(f => f.mercado === 'Fiori');
ok('panel: 9 mercados, solo impulsadoras', r.ok && r.filas.length === 9 && !r.filas.some(f => /supervisor|coordinador/i.test(f.nombre)), JSON.stringify(r).slice(0, 200));
ok('panel: Fiori con entrada y canjes', fi && fi.entrada && fi.canjes === 2 && fi.pedidos === 3 && fi.sacos === 16 + 17 + 34 && fi.ultima, JSON.stringify(fi));
ok('panel: impulsadora no entra', post({ ...K('fiori'), tipo: 'panel' }).code === 'val');
ok('vale: letras rechazadas', /7 números \(sin letras\)/.test(post(canje({ ...K('fiori'), id: 'v1', ricocan: 17, n_comprobante: 'BV1', codigos: ['AB12345'], dni: '55556666' })).error));
ok('vale: 6 dígitos rechazado', /escribiste 6/.test(post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '123456' }).mensaje));
ok('vale: 8 dígitos rechazado', post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '12345678' }).valido === false);
ok('vale con guion o espacio se normaliza', post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '200-0099' }).valido === true);
ok('vale repetido avisa dónde', /ya se registró \(Fiori/.test(post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '1000002' }).mensaje));
sheets.Talonario.rows.push(['3000001', 'Fiori', ''], ['3000002', '', ''], ['3000003', 'Parada', '']);
ok('talonario: código fuera de lista', /no es un vale de la campaña/.test(post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '3000099' }).mensaje));
ok('talonario: vale de otro mercado', /es del mercado Parada/.test(post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '3000003' }).mensaje));
r = post(canje({ ...K('fiori'), id: 'v2', ricocan: 17, n_comprobante: 'BV2', codigos: ['300 0001'], dni: '55556666' }));
ok('talonario: vale correcto se canjea', r.ok && r.vales === 1 && sheets.Vales.rows.some(v => v[0] === '3000001'), JSON.stringify(r));
ok('talonario: vale sin mercado sirve en cualquiera', post({ ...K('parada'), tipo: 'verificar_vale', codigo: '3000002' }).valido === true);
// ---- dashboards cliente / interno y validación ----
post({ ...F, tipo: 'incidencia', id: 'i1', tipo_incidencia: 'Otro', detalle: 'algo pasó' });
const dc = get({ action: 'data', k: props.DASH_TOKEN });
ok('cliente: sin incidencias ni observaciones', dc.modo === 'cliente' && dc.incidencias.length === 0 && !dc.pedidos.some(p => 'observaciones' in p || 'alertas' in p || 'celular' in p));
const di = post({ ...S, tipo: 'datos_interno' });
ok('interno: con incidencias, observaciones y DNI completo', di.ok && di.modo === 'interno' && di.incidencias.length === 1 && di.pedidos.some(p => p.dni === '12345678') && 'observaciones' in di.pedidos[0], JSON.stringify(di).slice(0, 120));
ok('interno: impulsadora no entra', post({ ...F, tipo: 'datos_interno' }).code === 'auth');
ok('validar: impulsadora no puede', post({ ...F, tipo: 'validar', id: 'p2', validacion: 'Válido' }).code === 'val');
ok('validar: observado pide motivo', /motivo/.test(post({ ...S, tipo: 'validar', id: 'p2', validacion: 'Observado' }).error));
r = post({ ...S, tipo: 'validar', id: 'p2', validacion: 'Válido' });
const p2v = sh('Pedidos').find(p => p.id === 'p2');
ok('validar: queda Válido con quién y cuándo', r.ok && p2v.validacion === 'Válido' && p2v.validado_por === sup.nombre && p2v.validado_en, JSON.stringify(p2v).slice(0, 80));
post({ ...S, tipo: 'validar', id: 'p5', validacion: 'Observado', nota: 'foto borrosa' });
ok('validar: observado guarda la nota', /foto borrosa/.test(sh('Pedidos').find(p => p.id === 'p5').observaciones));
ok('validar: id inexistente', /No se encontró/.test(post({ ...S, tipo: 'validar', id: 'zz', validacion: 'Válido' }).error));
// Migración: hoja antigua con 5 columnas en Equipo
const viejo = sheets.Equipo.rows; viejo.forEach(x => x.length = 5); viejo[0] = ['nombre', 'pin', 'rol', 'mercado', 'activo'];
props.ESQUEMA = '1'; get({ action: 'config' });
ok('migración agrega usuario y contraseña', sheets.Equipo.rows[0].join() === 'nombre,pin,rol,mercado,activo,usuario,contrasena,celular' && sheets.Equipo.rows.slice(1).every(x => x[5] && x[6]), sheets.Equipo.rows[1].join());
ok('login tolerante: "Unicachi Sur" y clave en mayúsculas', post({ tipo: 'login', usuario: ' Unicachi Sur ', clave: sheets.Equipo.rows.find(e => e[5] === 'unicachi.sur')[6].toUpperCase() }).ok);
ok('login cliente con clave en mayúsculas', !!post({ tipo: 'login_cliente', usuario: 'RINTI', clave: cfg.cliente_clave.toUpperCase() }).k);
ok('migración tampoco duplica jefes', sh('Equipo').filter(e => e.rol === 'coordinador').length === 1 && sh('Equipo').length === 12);
ok('hora sin cero se normaliza', hhmm_('9:05') === '09:05' && hhmm_('13:40') === '13:40' && hhmm_('') === '');
res.join('\n');
