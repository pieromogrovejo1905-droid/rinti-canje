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
ok('config publica sin claves y con reglas', (() => { const c = get({ action: 'config' }); return !JSON.stringify(c).match(/contrasena|cliente_clave|"pin"/) && c.sacos_por_mercado['Fiori'] === 15; })());
ok('login cliente con espacios', post({ tipo: 'login_cliente', usuario: ' Rinti ', clave: cfg.cliente_clave + ' \n' }).k === props.DASH_TOKEN);
ok('login cliente malo', post({ tipo: 'login_cliente', usuario: 'rinti', clave: 'x' }).code === 'auth');
const F = { usuario: 'FIORI', clave: fiori.contrasena + ' ' };
let r = post({ ...F, tipo: 'login' });
ok('login por mercado', r.ok && r.usuario.mercado === 'Fiori' && r.usuario.nombre === fiori.nombre, JSON.stringify(r));
ok('login con clave de otro mercado', post({ usuario: 'fiori', clave: cdd.contrasena, tipo: 'login' }).code === 'auth');
const img = 'data:image/jpeg;base64,AAAA';
ok('asistencia queda en su mercado aunque mande otro', post({ ...F, tipo: 'asistencia', id: 'a1', mercado: 'Parada', marca: 'Entrada', foto: img }).ok && sheets.Asistencia.rows[1][5] === 'Fiori');
const canje = (extra) => ({ ...F, tipo: 'pedido', cliente: 'Bodega', tipo_comprobante: 'Boleta', foto_comprobante: img, foto_entrega: img, firma: img, compra_lima: true, consentimiento: true, lat: -12, ...extra });
r = post(canje({ id: 'p1', ricocan: 7, ricocat: 7, n_comprobante: 'B1', codigos: ['1000001'], dni: '12345678' }));
ok('Fiori 14 sacos no es canje (pide 15)', r.ok && r.vales === 0, JSON.stringify(r));
r = post(canje({ id: 'p2', ricocan: 7, ricocat: 8, n_comprobante: 'B2', codigos: ['1000002'], dni: '12345678' }));
ok('Fiori 15 sacos mixtos = 1 vale', r.ok && r.vales === 1, JSON.stringify(r));
r = post(canje({ id: 'p3', ricocan: 40, n_comprobante: 'B3', codigos: ['1000003'], dni: '12345678' }));
ok('mismo DNI el mismo día: rechazado', r.code === 'val' && /ya recibió un vale hoy/.test(r.error), JSON.stringify(r));
r = post(canje({ id: 'p4', ricocan: 34, n_comprobante: 'B4', codigos: ['1000004', '1000005'], dni: '87654321' }));
ok('34 sacos sigue siendo 1 vale (un solo código)', r.code === 'val' && /código del vale/.test(r.error), JSON.stringify(r));
r = post(canje({ id: 'p5', ricocan: 34, n_comprobante: 'B5', codigos: ['1000004'], dni: '87654321' }));
ok('34 sacos con 1 código: ok', r.ok && r.vales === 1);
ok('canje sin DNI', post(canje({ id: 'p6', ricocan: 30, n_comprobante: 'B6', codigos: ['1000006'] })).error === 'Falta el DNI del cliente');
ok('DNI inválido', /8 dígitos/.test(post(canje({ id: 'p7', ricocan: 30, n_comprobante: 'B7', codigos: ['1000007'], dni: '123' })).error));
ok('DNI sin consentimiento', /consentimiento/.test(post(canje({ id: 'p8', ricocan: 30, n_comprobante: 'B8', codigos: ['1000008'], dni: '11112222', consentimiento: false })).error));
const C = { usuario: 'ciudaddedios', clave: cdd.contrasena };
r = post({ ...canje({ id: 'p9', ricocan: 6, ricocat: 4, n_comprobante: 'B9', codigos: ['1000009'], dni: '22223333' }), ...C });
ok('Ciudad de Dios 10 sacos = vale', r.ok && r.vales === 1, JSON.stringify(r));
ok('verificar DNI ya usado', (() => { const v = post({ ...C, tipo: 'verificar_dni', dni: '12345678' }); return v.ok && v.disponible === false && /Fiori/.test(v.mensaje); })());
ok('verificar DNI libre', post({ ...C, tipo: 'verificar_dni', dni: '99998888' }).disponible === true);
const S = { usuario: 'supervisor', clave: sup.contrasena };
r = post({ ...canje({ id: 'p10', ricocan: 14, n_comprobante: 'B10', codigos: ['1000010'], dni: '33334444', puesto: 'Sumaq Sara' }), ...S, mercado: 'Mercado Productores' });
ok('Productores 14 sacos no es canje', r.ok && r.vales === 0, JSON.stringify(r));
r = post({ ...canje({ id: 'p11', ricocan: 15, n_comprobante: 'B11', codigos: ['1000011'], dni: '33334444', puesto: 'Sumaq Sara' }), ...S, mercado: 'Mercado Productores' });
ok('Productores 15 sacos = canje (supervisor elige mercado)', r.ok && r.vales === 1);
ok('supervisor con mercado inválido', /Mercado no válido/.test(post({ ...canje({ id: 'p12', ricocan: 30, n_comprobante: 'B12', codigos: ['1000012'], dni: '44445555' }), ...S, mercado: 'Otro' }).error));
ok('cierre del supervisor', post({ ...S, tipo: 'cierre', id: 'c1', mercado: 'Fiori', pavos_inicio: 10, pavos_entregados: 1, vales_reportados: 1 }).ok);
ok('cierre de impulsadora: no', post({ ...F, tipo: 'cierre', id: 'c2', mercado: 'Fiori', pavos_inicio: 10, pavos_entregados: 1 }).code === 'val');
const d = get({ action: 'data', k: props.DASH_TOKEN });
const p2 = d.pedidos.find(p => p.id === 'p2');
ok('dashboard: DNI enmascarado y sin celular', p2.dni === '•••••678' && !d.pedidos.some(p => 'celular' in p), p2.dni);
ok('dashboard: sacos requeridos guardados', p2.sacos_requeridos === '15' && p2.es_canje === 'SÍ');
ok('dashboard: reglas en la respuesta', d.sacos_por_mercado['Mercado Productores'] === 15 && d.sacos_por_mercado['Unicachi Sur'] === 8 && d.sacos_por_mercado['Ciudad de Dios'] === 8);
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
ok('panel: Fiori con entrada y canjes', fi && fi.entrada && fi.canjes === 2 && fi.pedidos === 3 && fi.sacos === 14 + 15 + 34 && fi.ultima, JSON.stringify(fi));
ok('panel: impulsadora no entra', post({ ...K('fiori'), tipo: 'panel' }).code === 'val');
ok('vale: letras rechazadas', /7 números \(sin letras\)/.test(post(canje({ ...K('fiori'), id: 'v1', ricocan: 22, n_comprobante: 'BV1', codigos: ['AB12345'], dni: '55556666' })).error));
ok('vale: 6 dígitos rechazado', /escribiste 6/.test(post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '123456' }).mensaje));
ok('vale: 8 dígitos rechazado', post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '12345678' }).valido === false);
ok('vale con guion o espacio se normaliza', post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '200-0099' }).valido === true);
ok('vale repetido avisa dónde', /ya se registró \(Fiori/.test(post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '1000002' }).mensaje));
sheets.Talonario.rows.push(['3000001', 'Fiori', ''], ['3000002', '', ''], ['3000003', 'Parada', '']);
ok('talonario: código fuera de lista', /no es un vale de la campaña/.test(post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '3000099' }).mensaje));
ok('talonario: vale de otro mercado', /es del mercado Parada/.test(post({ ...K('fiori'), tipo: 'verificar_vale', codigo: '3000003' }).mensaje));
r = post(canje({ ...K('fiori'), id: 'v2', ricocan: 22, n_comprobante: 'BV2', codigos: ['300 0001'], dni: '55556666' }));
ok('talonario: vale correcto se canjea', r.ok && r.vales === 1 && sheets.Vales.rows.some(v => v[0] === '3000001'), JSON.stringify(r));
ok('talonario: vale sin mercado sirve en cualquiera', post({ ...K('parada'), tipo: 'verificar_vale', codigo: '3000002' }).valido === true);
// ---- presentaciones Ricocan 15 kg y Ricocat 9 kg ----
r = post(canje({ ...K('fiori'), id: 'v3', ricocan15: 8, ricocat9: 7, n_comprobante: 'BV3', codigos: ['3000002'], dni: '66667777' }));
const v3 = sh('Pedidos').find(p => p.id === 'v3');
ok('Ricocan 15 kg + Ricocat 9 kg cuentan para el vale', r.ok && r.vales === 1 && v3.sacos_total == 15 && v3.sacos_ricocan == 8 && v3.sacos_ricocat == 7, JSON.stringify(r));
ok('detalle y kg de las presentaciones nuevas', v3.sacos_ricocan_15kg == 8 && v3.sacos_ricocat_9kg == 7 && v3.kg_total == 8 * 15 + 7 * 9, JSON.stringify(v3).slice(0, 200));
// ---- relevo de stock ----
const FOTO = 'data:image/jpeg;base64,AAAA';
const skOk = { rc22: { estado: 'Con stock', exhibicion: 4, almacen: 6, transito: 'NO' }, rc15: { estado: 'Con stock', exhibicion: 3, almacen: 0, transito: 'NO' },
  rt15: { estado: 'No trabaja este SKU' }, rt9: { estado: 'Con stock', exhibicion: 2, almacen: 5, transito: 'SÍ', transito_sacos: 10 } };
r = post({ ...K('fiori'), tipo: 'stock', id: 's1', turno: 'Apertura', skus: skOk, foto: FOTO, mercado: 'Parada' });
const s1 = sh('Stock').find(x => x.id === 's1');
ok('stock: guarda relevo con totales automáticos', r.ok && !r.alerta && s1 && s1.mercado === 'Fiori' && s1.rc22_total == 10 && s1.rc15_total == 3 && s1.rt15_total === '' && s1.rt9_transito_sacos == 10 && s1.alerta === 'NO', JSON.stringify(r) + JSON.stringify(s1).slice(0, 300));
ok('stock: pide turno', /turno/.test(post({ ...K('fiori'), tipo: 'stock', id: 's2', skus: skOk, foto: FOTO }).error));
ok('stock: pide foto', /foto/.test(post({ ...K('fiori'), tipo: 'stock', id: 's2', turno: 'Cierre', skus: skOk }).error));
ok('stock: pide estado de cada presentación', /Ricocat 15 kg/.test(post({ ...K('fiori'), tipo: 'stock', id: 's2', turno: 'Cierre', skus: { ...skOk, rt15: {} }, foto: FOTO }).error));
ok('stock: tránsito sin cantidad', /tránsito/.test(post({ ...K('fiori'), tipo: 'stock', id: 's2', turno: 'Cierre', skus: { ...skOk, rt9: { estado: 'Con stock', exhibicion: 5, transito: 'SÍ' } }, foto: FOTO }).error));
const skQ = { ...skOk, rc15: { estado: 'Quiebre total', exhibicion: 9, transito: 'NO' } };
ok('stock: quiebre pide motivo', /motivo/.test(post({ ...K('fiori'), tipo: 'stock', id: 's2', turno: 'Cierre', skus: skQ, foto: FOTO }).error));
ok('stock: quiebre pide aviso', /alertó/.test(post({ ...K('fiori'), tipo: 'stock', id: 's2', turno: 'Cierre', skus: skQ, foto: FOTO, motivos: ['Demora en entrega'] }).error));
r = post({ ...K('fiori'), tipo: 'stock', id: 's2', turno: 'Cierre', skus: skQ, foto: FOTO, motivos: ['Demora en entrega', 'Falta de liquidez'], aviso_distribuidor: 'SÍ' });
const s2 = sh('Stock').find(x => x.id === 's2');
ok('stock: quiebre se guarda con alerta y motivos', r.ok && r.alerta && s2.rc15_total == 0 && s2.rc15_exhibicion == 0 && s2.alerta === 'SÍ' && /Demora.*Falta/.test(s2.motivos) && s2.aviso_distribuidor === 'SÍ', JSON.stringify(s2).slice(0, 300));
ok('stock: menos de 3 sacos es bajo stock', /motivo/.test(post({ ...K('fiori'), tipo: 'stock', id: 's3', turno: 'Cierre', skus: { ...skOk, rc22: { estado: 'Con stock', exhibicion: 2, transito: 'NO' } }, foto: FOTO }).error));
ok('stock: supervisor elige mercado', post({ ...S, tipo: 'stock', id: 's4', turno: 'Apertura', mercado: 'Parada', skus: skOk, foto: FOTO }).ok && sh('Stock').find(x => x.id === 's4').mercado === 'Parada');
ok('stock: id repetido no duplica', post({ ...K('fiori'), tipo: 'stock', id: 's1', turno: 'Apertura', skus: skOk, foto: FOTO }).duplicado && sh('Stock').filter(x => x.id === 's1').length === 1);
// ---- completar canje (pedido guardado sin vale con la regla anterior) ----
const filaVieja = (id, dni) => sheets.Pedidos.rows[0].map(k => ({ id, registrado: '2026-11-09 16:21:34', fecha: '2026-11-09', hora: '16:21', impulsadora: 'Impulsadora 2', mercado: 'Unicachi Norte Mayorista', cliente: 'Alfredo', n_comprobante: 'X' + id, sacos_ricocan: 4, sacos_ricocat: 6, sacos_total: 10, es_canje: 'NO', vales: 0, foto_comprobante: 'https://drive/x', dni, sacos_requeridos: 15, cumple_mecanica: 'N/A' })[k] ?? '');
sheets.Pedidos.rows.push(filaVieja('old1', '6844459'), filaVieja('old2', '72915844'));
sheets.Talonario.rows.push(['4000001', '', ''], ['4000002', '', ''], ['4000003', '', '']);
ok('completar: impulsadora no puede', post({ ...F, tipo: 'completar_canje', id: 'old1', codigo: '4000001', foto_entrega: img, dni: '06844459' }).code === 'val');
ok('completar: DNI de 7 dígitos se rechaza', /8 dígitos/.test(post({ ...S, tipo: 'completar_canje', id: 'old1', codigo: '4000001', foto_entrega: img }).error));
ok('completar: pide foto', /foto/.test(post({ ...S, tipo: 'completar_canje', id: 'old1', codigo: '4000001', dni: '06844459' }).error));
ok('completar: vale ya usado', /ya se registró/.test(post({ ...S, tipo: 'completar_canje', id: 'old1', codigo: '3000002', foto_entrega: img, dni: '06844459' }).error));
r = post({ ...S, tipo: 'completar_canje', id: 'old1', codigo: '400 0001', foto_entrega: img, dni: '06844459' });
const o1 = sh('Pedidos').find(p => p.id === 'old1');
ok('completar: queda canje pendiente con vale, DNI y regla vigente', r.ok && o1.es_canje === 'SÍ' && o1.vales == 1 && o1.codigos_vale === '4000001' && o1.dni === '06844459' && o1.sacos_requeridos == 10 && o1.validacion === 'Pendiente' && o1.foto_entrega && /Sin firma/.test(o1.alertas) && /completado por/.test(o1.observaciones) && sheets.Vales.rows.some(v => v[0] === '4000001' && v[2] === '2026-11-09'), JSON.stringify(o1).slice(0, 400));
ok('completar: no se repite', /ya es un canje/.test(post({ ...S, tipo: 'completar_canje', id: 'old1', codigo: '4000002', foto_entrega: img }).error));
ok('completar: con firma no marca alerta', post({ ...S, tipo: 'completar_canje', id: 'old2', codigo: '4000002', foto_entrega: img, firma: img }).ok && !/Sin firma/.test(sh('Pedidos').find(p => p.id === 'old2').alertas));
ok('completar: pedido con menos sacos que la regla', /se piden/.test(post({ ...S, tipo: 'completar_canje', id: 'p1', codigo: '4000003', foto_entrega: img }).error));
// ---- dashboards cliente / interno y validación ----
post({ ...F, tipo: 'incidencia', id: 'i1', tipo_incidencia: 'Otro', detalle: 'algo pasó' });
const dc = get({ action: 'data', k: props.DASH_TOKEN });
ok('cliente: sin incidencias ni observaciones', dc.modo === 'cliente' && dc.incidencias.length === 0 && !dc.pedidos.some(p => 'observaciones' in p || 'alertas' in p || 'celular' in p));
const di = post({ ...S, tipo: 'datos_interno' });
ok('interno: con incidencias, observaciones y DNI completo', di.ok && di.modo === 'interno' && di.incidencias.length === 1 && di.pedidos.some(p => p.dni === '12345678') && 'observaciones' in di.pedidos[0], JSON.stringify(di).slice(0, 120));
ok('dashboards reciben el stock', dc.stock.length === 3 && di.stock.length === 3);
ok('interno: impulsadora no entra', post({ ...F, tipo: 'datos_interno' }).code === 'auth');
ok('validar: impulsadora no puede', post({ ...F, tipo: 'validar', id: 'p2', validacion: 'Válido' }).code === 'val');
ok('validar: observado pide motivo', /motivo/.test(post({ ...S, tipo: 'validar', id: 'p2', validacion: 'Observado' }).error));
r = post({ ...S, tipo: 'validar', id: 'p2', validacion: 'Válido' });
const p2v = sh('Pedidos').find(p => p.id === 'p2');
ok('validar: queda Válido con quién y cuándo', r.ok && p2v.validacion === 'Válido' && p2v.validado_por === sup.nombre && p2v.validado_en, JSON.stringify(p2v).slice(0, 80));
post({ ...S, tipo: 'validar', id: 'p5', validacion: 'Observado', nota: 'foto borrosa' });
ok('validar: observado guarda la nota', /foto borrosa/.test(sh('Pedidos').find(p => p.id === 'p5').observaciones));
ok('validar: id inexistente', /No se encontró/.test(post({ ...S, tipo: 'validar', id: 'zz', validacion: 'Válido' }).error));
// Kardex de vales de pavo
let k = post({ ...S, tipo: 'kardex' });
ok('kardex: 9 mercados, lote 1350, sin dotación', k.ok && k.mercados.length === 9 && k.lote === 1350 && k.mercados.every(m => m.dotacion === 0), JSON.stringify(k).slice(0, 120));
ok('kardex: canjes por mercado desde Pedidos', k.canjes.Fiori && k.canjes.Fiori.total > 0, JSON.stringify(k.canjes));
ok('selfies: impulsadora no entra', /solo para supervisores/i.test(post({ ...F, tipo: 'selfies', ids: ['abcdefghijk'] }).error));
ok('selfies: ignora archivos que no son de Asistencia', (r => r.ok && Object.keys(r.fotos).length === 0)(post({ ...S, tipo: 'selfies', ids: ['1NoEstaEnAsistencia00'] })));
ok('kardex: impulsadora no entra', /solo para supervisores/.test(post({ ...F, tipo: 'kardex' }).error));
ok('kardex: dotación no puede exceder el lote', /excede el lote/.test(post({ ...S, tipo: 'kardex_guardar', filas: k.mercados.map(m => ({ mercado: m.mercado, dotacion: 200 })) }).error));
ok('kardex: dotación entera', /entero/.test(post({ ...S, tipo: 'kardex_guardar', filas: [{ mercado: 'Fiori', dotacion: 1.5 }] }).error));
k = post({ ...S, tipo: 'kardex_guardar', filas: k.mercados.map(m => ({ mercado: m.mercado, supervisor: 'Supervisor 1', dotacion: 150 })) });
ok('kardex: guarda dotación y supervisor', k.ok && k.mercados.every(m => m.dotacion === 150 && m.supervisor === 'Supervisor 1') && sheets.KardexMercado.rows.length === 10, JSON.stringify(k.mercados[0]));
const KE = { ...S, tipo: 'kardex_entrega', mercado: 'Fiori', recibe: 'Ana', fotos: [FOTO, FOTO] };
ok('entrega: pide foto', /al menos una foto/.test(post({ ...KE, id: 'k0', cantidad: 10, fotos: [] }).error));
ok('entrega: pide quién recibe', /quién recibe/.test(post({ ...KE, id: 'k0', cantidad: 10, recibe: '' }).error));
ok('entrega: cantidad válida', /cantidad/.test(post({ ...KE, id: 'k0', cantidad: 0 }).error));
k = post({ ...KE, id: 'k1', cantidad: 100 });
ok('entrega: se registra con 2 fotos', k.ok && k.entregas.length === 1 && k.entregas[0].fotos.split(' ').length === 2 && k.entregas[0].supervisor === sup.nombre, JSON.stringify(k.entregas));
ok('entrega: id repetido no duplica', post({ ...KE, id: 'k1', cantidad: 100 }).duplicado && sh('KardexEntregas').length === 1);
const nf = Object.keys(files).length;
ok('entrega: valida saldo del mercado', /Saldo por entregar de Fiori: 50/.test(post({ ...KE, id: 'k2', cantidad: 51 }).error));
ok('entrega: rechazo no deja fotos', Object.keys(files).length === nf);
ok('kardex: dotación no baja de lo entregado', /menor que lo ya entregado \(100\)/.test(post({ ...S, tipo: 'kardex_guardar', filas: [{ mercado: 'Fiori', dotacion: 90 }] }).error));
ok('entrega: impulsadora no puede', /Solo el supervisor/.test(post({ ...KE, ...F, id: 'k3', cantidad: 1 }).error));
// Migración: hoja antigua con 5 columnas en Equipo
const viejo = sheets.Equipo.rows; viejo.forEach(x => x.length = 5); viejo[0] = ['nombre', 'pin', 'rol', 'mercado', 'activo'];
props.ESQUEMA = '1'; get({ action: 'config' });
ok('migración agrega usuario y contraseña', sheets.Equipo.rows[0].join() === 'nombre,pin,rol,mercado,activo,usuario,contrasena,celular' && sheets.Equipo.rows.slice(1).every(x => x[5] && x[6]), sheets.Equipo.rows[1].join());
ok('login tolerante: "Unicachi Sur" y clave en mayúsculas', post({ tipo: 'login', usuario: ' Unicachi Sur ', clave: sheets.Equipo.rows.find(e => e[5] === 'unicachi.sur')[6].toUpperCase() }).ok);
ok('login cliente con clave en mayúsculas', !!post({ tipo: 'login_cliente', usuario: 'RINTI', clave: cfg.cliente_clave.toUpperCase() }).k);
ok('migración tampoco duplica jefes', sh('Equipo').filter(e => e.rol === 'coordinador').length === 1 && sh('Equipo').length === 12);
ok('meta 6 hasta el 20/11 y 8 desde el 21/11', metaDe_('2026-11-20', 6) === 6 && metaDe_('2026-11-21', 6) === 8 && metaDe_('2026-11-30', 6, 'Fiori') === 8);
ok('meta 9 en Ciudad de Dios y Unicachi Sur desde el 23/11', metaDe_('2026-11-22', 6, 'Unicachi Sur') === 8 && metaDe_('2026-11-23', 6, 'Unicachi Sur') === 9 && metaDe_('2026-11-23', 6, 'Ciudad de Dios') === 9 && metaDe_('2026-11-23', 6, 'Parada') === 8 && metaDe_('2026-11-23', 6) === 8);
ok('hora sin cero se normaliza', hhmm_('9:05') === '09:05' && hhmm_('13:40') === '13:40' && hhmm_('') === '');
res.join('\n');
