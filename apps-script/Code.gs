/**
 * Canje Rinti · Backend (Google Apps Script)
 *
 * Recibe los registros de la app de campo, guarda los datos en este Google Sheet
 * y las fotos en una carpeta de Drive. También entrega los datos al dashboard.
 *
 * Instalación: ver README.md (pasos 1 a 4). Correr setup() una sola vez.
 */

const TZ = 'America/Lima';
const KG = { ricocan: 22, ricocat: 15 };
const P = PropertiesService.getScriptProperties();
const VERSION_ESQUEMA = '4';

const MERCADOS = [
  'Fiori', 'Unicachi Norte Mayorista', 'Unicachi Norte Minorista', 'Unicachi Sur',
  'Huamantanga', 'Ciudad de Dios', 'Parada', 'Caquetá', 'Mercado Productores',
];

// Sacos (Ricocan y/o Ricocat, se pueden sumar) que el cliente debe comprar para recibir 1 vale.
const SACOS_POR_MERCADO = {
  'Mercado Productores': 30,
  'Fiori': 15,
  'Unicachi Sur': 10,
  'Parada': 10,
  'Unicachi Norte Mayorista': 10,
  'Unicachi Norte Minorista': 10,
  'Huamantanga': 10,
  'Ciudad de Dios': 10,
  'Caquetá': 10,
};

// Usuario de acceso de cada mercado (uno por impulsadora).
const USUARIO_MERCADO = {
  'Fiori': 'fiori',
  'Unicachi Norte Mayorista': 'unicachi.mayorista',
  'Unicachi Norte Minorista': 'unicachi.minorista',
  'Unicachi Sur': 'unicachi.sur',
  'Huamantanga': 'huamantanga',
  'Ciudad de Dios': 'ciudaddedios',
  'Parada': 'parada',
  'Caquetá': 'caqueta',
  'Mercado Productores': 'productores',
};

const ESQUEMA = {
  Asistencia: ['id', 'registrado', 'fecha', 'hora', 'impulsadora', 'mercado', 'marca', 'lat', 'lng', 'precision_m', 'foto'],
  Pedidos: ['id', 'registrado', 'fecha', 'hora', 'impulsadora', 'mercado', 'puesto', 'cliente', 'celular',
    'tipo_comprobante', 'n_comprobante', 'sacos_ricocan', 'sacos_ricocat', 'sacos_total', 'kg_total',
    'es_canje', 'vales', 'codigos_vale', 'foto_comprobante', 'foto_entrega', 'firma', 'compra_lima',
    'consentimiento', 'cumple_mecanica', 'alertas', 'observaciones', 'validacion', 'lat', 'lng',
    'dni', 'sacos_requeridos', 'validado_por', 'validado_en'],
  Incidencias: ['id', 'registrado', 'fecha', 'hora', 'impulsadora', 'mercado', 'tipo_incidencia', 'detalle', 'foto', 'lat', 'lng'],
  Cierres: ['id', 'registrado', 'fecha', 'hora', 'supervisor', 'mercado', 'pavos_inicio', 'pavos_entregados',
    'pavos_saldo', 'vales_reportados', 'observaciones'],
  Vales: ['codigo', 'pedido_id', 'fecha', 'impulsadora', 'mercado'],
  Equipo: ['nombre', 'pin', 'rol', 'mercado', 'activo', 'usuario', 'contrasena', 'celular'],
  // Lista de vales impresos. Si tiene códigos, solo se aceptan esos (y del mercado indicado, si se llena).
  Talonario: ['codigo', 'mercado', 'nota'],
  Config: ['clave', 'valor', 'nota'],
};

// Columnas que se guardan como texto (para que Sheets no las convierta en fecha/número).
const COLUMNAS_TEXTO = ['id', 'fecha', 'hora', 'celular', 'n_comprobante', 'codigos_vale', 'codigo', 'pedido_id', 'pin', 'valor',
  'dni', 'usuario', 'contrasena'];

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
      return ['Impulsadora ' + (i + 1), pinAleatorio_(), 'impulsadora', m, 'SI', USUARIO_MERCADO[m], claveAleatoria_()];
    });
    filas.push(['Supervisor 1', pinAleatorio_(), 'supervisor', '', 'SI', 'supervisor', claveAleatoria_()]);
    eq.getRange(2, 1, filas.length, 7).setValues(filas);
    // supervisor2 y coordinador los agrega asegurarAccesos_()
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
  Logger.log('Accesos de impulsadoras (uno por mercado): pestaña Equipo, columnas usuario y contrasena');
  Logger.log('Acceso del cliente al dashboard: pestaña Config, filas cliente_usuario y cliente_clave');
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
    asegurarAccesos_();
    if (accion === 'config') {
      return out_({ ok: true, equipo: equipoPublico_(), config: configPublica_(), sacos_por_mercado: SACOS_POR_MERCADO });
    }
    if (accion === 'data') {
      if (!e.parameter.k || e.parameter.k !== P.getProperty('DASH_TOKEN')) {
        return out_({ ok: false, code: 'auth', error: 'Link no autorizado' });
      }
      return out_(datos_(false));
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
      const ok = normUsuario_(b.usuario) === normUsuario_(c.cliente_usuario)
        && normClave_(b.clave) !== '' && normClave_(b.clave) === normClave_(c.cliente_clave);
      return out_(ok ? { ok: true, k: P.getProperty('DASH_TOKEN') } : { ok: false, code: 'auth', error: 'Usuario o contraseña incorrectos' });
    }
    const u = auth_(b);
    if (!u) return out_({ ok: false, code: 'auth', error: 'Usuario o contraseña incorrectos' });
    if (b.tipo === 'login') return out_({ ok: true, usuario: u });
    if (b.tipo === 'verificar_dni') return out_(verificarDni_(b, u));
    if (b.tipo === 'verificar_vale') return out_(verificarVale_(b, u));
    if (b.tipo === 'panel') return out_(panel_(u));
    if (b.tipo === 'datos_interno') {
      if (!esJefe_(u)) return out_({ ok: false, code: 'auth', error: 'El dashboard interno es solo para supervisores y coordinador' });
      return out_(Object.assign(datos_(true), { usuario: u }));
    }
    if (b.tipo === 'validar') return out_(validar_(b, u));
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

/**
 * Datos para los dashboards.
 * Cliente (Rinti): sin incidencias, observaciones, alertas, contacto ni DNI completo.
 * Interno (Free): todo.
 */
function datos_(interno) {
  const pedidos = leer_('Pedidos').map(function (p) {
    if (interno) return p;
    ['celular', 'consentimiento', 'observaciones', 'alertas', 'validado_por', 'validado_en', 'lat', 'lng']
      .forEach(function (k) { delete p[k]; });
    p.dni = p.dni ? '•••••' + String(p.dni).slice(-3) : ''; // DNI enmascarado
    return p;
  });
  const cierres = leer_('Cierres').map(function (c) { if (!interno) delete c.observaciones; return c; });
  return {
    ok: true, modo: interno ? 'interno' : 'cliente', generado: ahora_(),
    config: configPublica_(), sacos_por_mercado: SACOS_POR_MERCADO,
    equipo: equipoPublico_(), asistencia: leer_('Asistencia'), pedidos: pedidos,
    incidencias: interno ? leer_('Incidencias') : [], cierres: cierres,
  };
}

/** Validación de un canje desde el dashboard interno (Válido, Observado o volver a Pendiente). */
function validar_(b, u) {
  if (!esJefe_(u)) throw val_('Solo supervisores y coordinador pueden validar');
  const valor = limpio_(b.validacion);
  if (['Válido', 'Observado', 'Pendiente'].indexOf(valor) < 0) throw val_('Validación inválida');
  const nota = limpio_(b.nota);
  if (valor === 'Observado' && !nota) throw val_('Escribe el motivo de la observación');
  return conLock_(function () {
    const sh = hoja_('Pedidos'), cols = ESQUEMA.Pedidos;
    if (sh.getLastRow() < 2) throw val_('No se encontró el canje');
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    let fila = -1;
    for (let i = 0; i < ids.length; i++) if (ids[i][0] === String(b.id)) { fila = i + 2; break; }
    if (fila < 0) throw val_('No se encontró el canje');
    const col = function (k) { return cols.indexOf(k) + 1; };
    sh.getRange(fila, col('validacion')).setValue(valor);
    sh.getRange(fila, col('validado_por')).setValue(valor === 'Pendiente' ? '' : u.nombre);
    sh.getRange(fila, col('validado_en')).setValue(valor === 'Pendiente' ? '' : ahora_());
    if (nota) {
      const obs = sh.getRange(fila, col('observaciones')).getDisplayValues()[0][0];
      sh.getRange(fila, col('observaciones')).setValue((obs ? obs + ' | ' : '') + 'Validación (' + u.nombre + '): ' + nota);
    }
    return { ok: true, validacion: valor, validado_por: valor === 'Pendiente' ? '' : u.nombre };
  });
}

/** Supervisores y coordinador: eligen mercado, hacen el cierre y ven el panel del equipo. */
function esJefe_(u) { return /^(supervisor|coordinador)/.test(String(u && u.rol || '').toLowerCase()); }

/** La impulsadora siempre registra en su mercado; supervisores y coordinador eligen uno de la lista. */
function mercadoDe_(b, u) {
  const m = esJefe_(u) ? limpio_(b.mercado) : limpio_(u.mercado);
  if (!SACOS_POR_MERCADO[m]) throw val_('Mercado no válido: ' + (m || '(vacío)'));
  return m;
}

function asistencia_(b, u) {
  if (existeId_('Asistencia', b.id)) return { ok: true, duplicado: true };
  b.mercado = mercadoDe_(b, u);
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
  const mercado = mercadoDe_(b, u);
  if (total <= 0) throw val_('Indica la cantidad de sacos');
  if (!limpio_(b.cliente)) throw val_('Falta el nombre del cliente o negocio');
  if (/productores/i.test(mercado) && !limpio_(b.puesto)) throw val_('Elige el puesto de Mercado Productores');

  // Regla: en cada mercado se piden N sacos (Ricocan y/o Ricocat) y se da como máximo 1 vale por DNI al día.
  const requeridos = SACOS_POR_MERCADO[mercado];
  const esCanje = total >= requeridos;
  const vales = esCanje ? 1 : 0;
  const codigos = (b.codigos || []).map(normCodigo_).filter(String);
  const nComp = limpio_(b.n_comprobante);
  const dni = normDni_(b.dni);
  const alertas = [];
  const t = cuando_(b.capturado);

  if (dni && !dniValido_(dni)) throw val_('El DNI debe tener 8 dígitos (o carné de extranjería de 9 a 12 caracteres)');
  if (esCanje) {
    if (!dni) throw val_('Falta el DNI del cliente');
    const previo = dniUsadoHoy_(dni, t.fecha);
    if (previo) throw val_('Este DNI ya recibió un vale hoy (' + previo.mercado + ', ' + previo.hora + '). Máximo 1 vale por DNI al día.');
    if (codigos.length !== 1) throw val_('Ingresa el código del vale');
    const errVale = errorVale_(codigos[0], mercado);
    if (errVale) throw val_(errVale);
    if (!nComp) throw val_('Falta el número de boleta o factura');
    if (!b.foto_comprobante) throw val_('Falta la foto de la boleta o factura');
    if (!b.firma) throw val_('Falta la firma del cliente');
    if (!b.foto_entrega) throw val_('Falta la foto del cliente recibiendo el vale');
    if (!b.compra_lima) throw val_('El canje solo aplica a compras en Lima');
  }
  if ((limpio_(b.celular) || dni) && !b.consentimiento) throw val_('Falta el consentimiento del cliente para guardar sus datos');
  if (nComp && comprobanteUsado_(nComp)) alertas.push('Comprobante repetido');
  if (!b.lat) alertas.push('Sin GPS');

  const base = t.fecha + '_' + t.hora.replace(':', '') + '_' + u.nombre;
  const subidas = [];
  const fotoComp = foto_(b.foto_comprobante, base + '_comprobante', t.fecha, subidas);
  const fotoEnt = foto_(b.foto_entrega, base + '_entrega', t.fecha, subidas);
  const firma = foto_(b.firma, base + '_firma', t.fecha, subidas);

  return conLock_(function () {
    if (existeId_('Pedidos', b.id)) { borrar_(subidas); return { ok: true, duplicado: true }; }
    const usados = codigos.filter(valeUsado_);
    if (usados.length) { borrar_(subidas); throw val_('Vale ya registrado antes: ' + usados.join(', ')); }
    const previo = esCanje && dniUsadoHoy_(dni, t.fecha);
    if (previo) { borrar_(subidas); throw val_('Este DNI ya recibió un vale hoy (' + previo.mercado + ', ' + previo.hora + '). Máximo 1 vale por DNI al día.'); }

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
      dni: dni, sacos_requeridos: requeridos,
    });
    codigos.forEach(function (c) {
      agregar_('Vales', { codigo: c, pedido_id: b.id, fecha: t.fecha, impulsadora: u.nombre, mercado: mercado });
    });
    return { ok: true, vales: vales, alertas: alertas };
  });
}

/** Consulta si un DNI ya recibió vale hoy (antes de entregar el vale físico). */
function verificarDni_(b, u) {
  const dni = normDni_(b.dni);
  if (!dniValido_(dni)) return { ok: false, code: 'val', error: 'DNI inválido' };
  const previo = dniUsadoHoy_(dni, cuando_().fecha);
  return previo
    ? { ok: true, disponible: false, mensaje: 'Ya recibió un vale hoy en ' + previo.mercado + ' a las ' + previo.hora }
    : { ok: true, disponible: true, mensaje: 'Puede recibir su vale hoy' };
}

/**
 * Regla del vale: solo se acepta un código de vale real.
 * - Solo números, exactamente 7 dígitos.
 * - Si la hoja Talonario tiene códigos: tiene que estar en la lista y, si tiene mercado, ser de ese mercado.
 * - Cada código se usa una sola vez.
 */
function errorVale_(codigo, mercado) {
  if (!/^\d{7}$/.test(codigo)) return 'El código del vale tiene 7 números' + (/[A-Z]/.test(codigo) ? ' (sin letras)' : codigo ? ' (escribiste ' + codigo.length + ')' : '');
  const lista = leer_('Talonario').filter(function (r) { return normCodigo_(r.codigo); });
  if (lista.length) {
    const r = lista.find(function (x) { return normCodigo_(x.codigo) === codigo; });
    if (!r) return 'El código ' + codigo + ' no es un vale de la campaña. Revisa que esté bien escrito';
    if (limpio_(r.mercado) && normUsuario_(r.mercado) !== normUsuario_(mercado)) return 'El vale ' + codigo + ' es del mercado ' + limpio_(r.mercado);
  }
  if (valeUsado_(codigo)) {
    const v = leer_('Vales').find(function (x) { return x.codigo === codigo; }) || {};
    return 'El vale ' + codigo + ' ya se registró' + (v.mercado ? ' (' + v.mercado + ', ' + v.fecha + ')' : '');
  }
  return '';
}

/** Consulta el código antes de guardar el canje. */
function verificarVale_(b, u) {
  const codigo = normCodigo_(b.codigo);
  const err = errorVale_(codigo, mercadoDe_(b, u));
  return err ? { ok: true, valido: false, mensaje: err } : { ok: true, valido: true, mensaje: 'Código válido' };
}

/** Panel para supervisores y coordinador: quién llegó, quién va tarde y cómo va cada una con su cuota. */
function panel_(u) {
  if (!esJefe_(u)) throw val_('El panel es solo para supervisores y coordinador');
  const t = cuando_(), cfg = config_();
  const meta = Number(cfg.meta_diaria_por_impulsadora) || 6;
  const ingreso = /^\d{1,2}:\d{2}$/.test(cfg.hora_ingreso || '') ? ('0' + cfg.hora_ingreso).slice(-5) : '09:00';
  const limite = minAHora_(horaAMin_(ingreso) + 10);
  const equipo = leer_('Equipo').filter(function (x) { return !/^no$/i.test(x.activo); });
  const jefes = equipo.filter(function (x) { return esJefe_(x); }).map(function (x) { return x.nombre; });
  // Sheets puede guardar "09:54" como "9:54": se normaliza para comparar bien las horas.
  const hoy = function (r) { if (r.fecha !== t.fecha) return false; r.hora = hhmm_(r.hora); return true; };
  const asis = leer_('Asistencia').filter(hoy), ped = leer_('Pedidos').filter(hoy), inc = leer_('Incidencias').filter(hoy);
  const porHora = function (a, b) { return a.hora < b.hora ? -1 : a.hora > b.hora ? 1 : 0; };
  const filas = equipo.filter(function (x) { return !esJefe_(x); }).map(function (x) {
    const m = x.mercado;
    const a = asis.filter(function (r) { return r.mercado === m && jefes.indexOf(r.impulsadora) < 0; }).sort(porHora);
    const ent = a.filter(function (r) { return r.marca === 'Entrada'; })[0];
    const sal = a.filter(function (r) { return r.marca === 'Salida'; }).pop();
    const p = ped.filter(function (r) { return r.mercado === m; });
    const i = inc.filter(function (r) { return r.mercado === m; }).sort(porHora);
    const horas = a.concat(p, i).map(function (r) { return r.hora; }).sort();
    return {
      nombre: x.nombre, mercado: m, celular: limpio_(x.celular),
      entrada: ent ? ent.hora : '', foto_entrada: ent ? ent.foto : '', tarde: !!ent && ent.hora > limite,
      salida: sal ? sal.hora : '',
      canjes: p.filter(function (r) { return r.es_canje === 'SÍ'; }).length, pedidos: p.length,
      sacos: p.reduce(function (s, r) { return s + (Number(r.sacos_total) || 0); }, 0),
      ultima: horas.length ? horas[horas.length - 1] : '',
      incidencias: i.map(function (r) { return { hora: r.hora, tipo: r.tipo_incidencia, detalle: r.detalle }; }),
      requeridos: SACOS_POR_MERCADO[m] || '',
    };
  });
  return { ok: true, fecha: t.fecha, hora: t.hora, meta: meta, hora_ingreso: ingreso, limite: limite, filas: filas };
}
function hhmm_(h) { const m = String(h || '').match(/(\d{1,2}):(\d{2})/); return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : ''; }
function horaAMin_(h) { const p = String(h).split(':'); return Number(p[0]) * 60 + Number(p[1] || 0); }
function minAHora_(m) { return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }

function normDni_(s) { return limpio_(s).toUpperCase().replace(/[\s.-]+/g, ''); }
function dniValido_(d) { return /^\d{8}$/.test(d) || /^[A-Z0-9]{9,12}$/.test(d); }

function dniUsadoHoy_(dni, fecha) {
  if (!dni) return null;
  return leer_('Pedidos').find(function (p) { return p.dni === dni && p.fecha === fecha && p.es_canje === 'SÍ'; }) || null;
}

function incidencia_(b, u) {
  if (existeId_('Incidencias', b.id)) return { ok: true, duplicado: true };
  b.mercado = mercadoDe_(b, u);
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
  if (!esJefe_(u)) throw val_('Solo el supervisor o el coordinador puede enviar el cierre');
  if (existeId_('Cierres', b.id)) return { ok: true, duplicado: true };
  b.mercado = mercadoDe_(b, u);
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

/** Acceso a la app de campo: un usuario y contraseña por mercado (hoja Equipo, columnas usuario y contrasena). */
function auth_(b) {
  const usuario = normUsuario_(b.usuario), clave = normClave_(b.clave);
  if (!usuario || !clave) return null;
  const r = leer_('Equipo').find(function (x) {
    return normUsuario_(x.usuario) === usuario && normClave_(x.contrasena) === clave && !/^no$/i.test(x.activo);
  });
  return r ? { nombre: r.nombre, rol: String(r.rol || 'impulsadora').toLowerCase(), mercado: r.mercado, usuario: r.usuario } : null;
}

/**
 * Deja la hoja lista para la versión actual: agrega columnas nuevas al final (sin mover datos),
 * crea los accesos por mercado en Equipo y el acceso del cliente en Config.
 */
function asegurarAccesos_() {
  if (P.getProperty('ESQUEMA') === VERSION_ESQUEMA && config_().cliente_clave) return;
  conLock_(function () {
    if (P.getProperty('ESQUEMA') === VERSION_ESQUEMA && config_().cliente_clave) return; // otro pedido ya lo hizo
    ['Pedidos', 'Equipo', 'Talonario'].forEach(function (nombre) {
      const cols = ESQUEMA[nombre];
      const sh = hoja_(nombre) || ss_().insertSheet(nombre);
      if (sh.getLastRow() === 0) sh.setFrozenRows(1);
      const actual = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0];
      const esPrefijo = actual.every(function (h, i) { return h === '' || h === cols[i]; });
      if (!esPrefijo) throw new Error('Las columnas de la hoja ' + nombre + ' no coinciden con el sistema');
      if (actual.length < cols.length || actual[cols.length - 1] !== cols[cols.length - 1]) {
        sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
      }
      cols.forEach(function (c, i) {
        if (COLUMNAS_TEXTO.indexOf(c) >= 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
      });
    });

    // Supervisor 2 y Coordinador de Trade (se agregan una sola vez)
    const eq = hoja_('Equipo');
    const existentes = leer_('Equipo').map(function (x) { return normUsuario_(x.usuario) + '|' + normUsuario_(x.nombre); }).join(' ');
    [['Supervisor 2', 'supervisor', 'supervisor2'], ['Coordinador de Trade', 'coordinador', 'coordinador']].forEach(function (j) {
      if (existentes.indexOf(j[2] + '|') >= 0 || existentes.indexOf('|' + normUsuario_(j[0])) >= 0) return;
      eq.getRange(eq.getLastRow() + 1, 1, 1, 8).setValues([[j[0], pinAleatorio_(), j[1], '', 'SI', j[2], claveAleatoria_(), '']]);
    });

    // Usuario y contraseña para cada fila del equipo que no tenga
    if (eq.getLastRow() > 1) {
      const cols = ESQUEMA.Equipo, iU = cols.indexOf('usuario'), iC = cols.indexOf('contrasena');
      const filas = eq.getRange(2, 1, eq.getLastRow() - 1, cols.length).getDisplayValues();
      let sup = 0;
      filas.forEach(function (f, n) {
        if (!f[0]) return;
        const esSup = /supervisor/i.test(f[2]);
        if (esSup) sup++;
        if (!limpio_(f[iU])) eq.getRange(n + 2, iU + 1).setValue(esSup ? 'supervisor' + (sup > 1 ? sup : '') : (USUARIO_MERCADO[f[3]] || claveAleatoria_()));
        if (!limpio_(f[iC])) eq.getRange(n + 2, iC + 1).setValue(claveAleatoria_());
      });
    }

    const c = config_(), falt = [];
    if (!c.cliente_usuario) falt.push(['cliente_usuario', 'rinti', 'Usuario del dashboard del cliente']);
    if (!c.cliente_clave) falt.push(['cliente_clave', claveAleatoria_(), 'Contraseña del dashboard del cliente (puedes cambiarla aquí)']);
    if (falt.length) { const sh = hoja_('Config'); sh.getRange(sh.getLastRow() + 1, 1, falt.length, 3).setValues(falt); }
    P.setProperty('ESQUEMA', VERSION_ESQUEMA);
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
// Tolera lo que pasa al escribir en el celular: mayúsculas, tildes, espacios o puntos de más.
// "Unicachi Sur", "unicachi.sur" y "UnicachiSur" son el mismo usuario.
function normUsuario_(s) { return limpio_(s).toLowerCase().normalize('NFD').replace(/[^a-z0-9]/g, ''); }
function normClave_(s) { return limpio_(s).toLowerCase().replace(/\s+/g, ''); }
function normCodigo_(s) { return limpio_(s).toUpperCase().normalize('NFD').replace(/[^A-Z0-9]/g, ''); }
function entero_(n) { const x = parseInt(n, 10); return isNaN(x) || x < 0 ? 0 : x; }
function num_(n) { const x = Number(n); return n === '' || n === null || n === undefined || isNaN(x) ? '' : x; }
function pinAleatorio_() { return String(Math.floor(1000 + Math.random() * 9000)); }
function val_(msg) { const e = new Error(msg); e.val = true; return e; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
