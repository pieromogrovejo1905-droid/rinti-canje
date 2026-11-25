// Descarga en Excel (.xlsx) de los pedidos y canjes, para supervisores y coordinador.
// La usan el dashboard interno y la app (perfil supervisor/coordinador).
// Hojas: Resumen por mercado · Por punto de venta · Por cliente · Detalle de canjes · Todos los pedidos.
(function () {
  'use strict';
  const num = v => { const x = Number(String(v == null ? '' : v).replace(',', '.')); return isFinite(x) ? x : 0; };
  const txt = v => String(v == null ? '' : v).trim();
  let lib = null;
  function cargar() {
    if (window.XLSX) return Promise.resolve();
    return lib = lib || new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      s.onload = res; s.onerror = () => { lib = null; rej(new Error('No se pudo cargar el generador de Excel. Revisa tu conexión.')); };
      document.head.appendChild(s);
    });
  }
  const esCanje = p => p.es_canje === 'SÍ' || num(p.vales) > 0;
  const punto = p => txt(p.puesto) || '(sin punto de venta)';
  function sacos(p) {
    const rc = num(p.sacos_ricocan), rt = num(p.sacos_ricocat), rc15 = num(p.sacos_ricocan_15kg), rt9 = num(p.sacos_ricocat_9kg);
    const kg = num(p.kg_total) || (rc - rc15) * 22 + rc15 * 15 + (rt - rt9) * 15 + rt9 * 9;
    return { rc, rt, rc22: rc - rc15, rc15, rt15: rt - rt9, rt9, tot: rc + rt, kg };
  }
  const claveCliente = p => p.mercado + '|' + (txt(p.dni).replace(/\W/g, '') || txt(p.cliente).toLowerCase().normalize('NFD').replace(/[^a-z0-9]/g, ''));

  // Agrupa pedidos por una clave y suma lo comercial.
  function agrupar(pedidos, clave, base) {
    const g = new Map();
    pedidos.forEach(p => {
      const k = clave(p);
      if (!g.has(k)) g.set(k, Object.assign(base(p), { ped: 0, cj: 0, v: 0, rc: 0, rt: 0, tot: 0, kg: 0, cli: new Set(), pv: new Set(), pvn: new Set(), ini: p.fecha, fin: p.fecha }));
      const r = g.get(k), s = sacos(p);
      r.ped++; if (esCanje(p)) { r.cj++; r.v += num(p.vales) || 1; }
      r.rc += s.rc; r.rt += s.rt; r.tot += s.tot; r.kg += s.kg;
      r.cli.add(claveCliente(p)); r.pv.add(p.mercado + '|' + punto(p)); r.pvn.add(punto(p));
      if (!r.cel && p.celular) r.cel = txt(p.celular);
      if (p.fecha < r.ini) r.ini = p.fecha; if (p.fecha > r.fin) r.fin = p.fecha;
    });
    return [...g.values()];
  }
  function hoja(XLSX, filas, anchos) {
    const sh = XLSX.utils.aoa_to_sheet(filas);
    sh['!cols'] = anchos.map(w => ({ wch: w }));
    return sh;
  }

  async function descargar(pedidos, o) {
    o = o || {};
    await cargar();
    const XLSX = window.XLSX;
    const P = (pedidos || []).slice().sort((a, b) => (a.fecha + ' ' + a.hora).localeCompare(b.fecha + ' ' + b.hora));
    const C = P.filter(esCanje);
    const wb = XLSX.utils.book_new();
    const cab = [['Canje Rinti · Ricocan & Ricocat'], ['Periodo', o.periodo || 'Toda la campaña'], ['Mercado', o.mercado || 'Todos los mercados'], ['Generado', new Date().toLocaleString('es-PE') + (o.usuario ? ' · ' + o.usuario : '')], []];

    // 1. Resumen por mercado
    const porM = agrupar(P, p => p.mercado, p => ({ m: p.mercado })).sort((a, b) => b.cj - a.cj);
    const T = porM.reduce((t, r) => { ['ped', 'cj', 'v', 'rc', 'rt', 'tot', 'kg'].forEach(k => t[k] += r[k]); return t; }, { ped: 0, cj: 0, v: 0, rc: 0, rt: 0, tot: 0, kg: 0 });
    const cliT = new Set(P.map(claveCliente)).size, pvT = new Set(P.map(p => p.mercado + '|' + punto(p))).size;
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, [...cab,
      ['Mercado', 'Pedidos', 'Canjes', 'Vales', 'Sacos Ricocan', 'Sacos Ricocat', 'Sacos total', 'Kg', 'Clientes', 'Puntos de venta'],
      ...porM.map(r => [r.m, r.ped, r.cj, r.v, r.rc, r.rt, r.tot, r.kg, r.cli.size, r.pv.size]),
      ['TOTAL', T.ped, T.cj, T.v, T.rc, T.rt, T.tot, T.kg, cliT, pvT],
    ], [28, 10, 10, 10, 14, 14, 12, 10, 10, 15]), 'Resumen por mercado');

    // 2. Por punto de venta
    const porPV = agrupar(P, p => p.mercado + '|' + punto(p), p => ({ m: p.mercado, nombre: punto(p) })).sort((a, b) => a.m.localeCompare(b.m) || b.cj - a.cj || b.tot - a.tot);
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, [...cab,
      ['Mercado', 'Punto de venta', 'Pedidos', 'Canjes', 'Vales', 'Sacos Ricocan', 'Sacos Ricocat', 'Sacos total', 'Kg', 'Clientes', 'Primera compra', 'Última compra'],
      ...porPV.map(r => [r.m, r.nombre, r.ped, r.cj, r.v, r.rc, r.rt, r.tot, r.kg, r.cli.size, r.ini, r.fin]),
    ], [26, 40, 9, 9, 8, 13, 13, 11, 9, 9, 14, 14]), 'Por punto de venta');

    // 3. Por cliente
    const porCli = agrupar(P, claveCliente, p => ({ m: p.mercado, cliente: txt(p.cliente), dni: txt(p.dni), cel: txt(p.celular) }));
    porCli.sort((a, b) => a.m.localeCompare(b.m) || b.cj - a.cj || b.tot - a.tot);
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, [...cab,
      ['Mercado', 'Cliente', 'DNI', 'Celular', 'Punto de venta', 'Compras', 'Canjes', 'Vales', 'Sacos Ricocan', 'Sacos Ricocat', 'Sacos total', 'Kg', 'Primera compra', 'Última compra'],
      ...porCli.map(r => [r.m, r.cliente, r.dni, r.cel, [...r.pvn].join(' / '), r.ped, r.cj, r.v, r.rc, r.rt, r.tot, r.kg, r.ini, r.fin]),
    ], [26, 30, 12, 12, 36, 9, 8, 7, 13, 13, 11, 9, 14, 14]), 'Por cliente');

    // 4 y 5. Detalle
    const head = ['Fecha', 'Hora', 'Mercado', 'Punto de venta', 'Impulsadora', 'Cliente', 'DNI', 'Celular', 'Tipo comprobante', 'N° comprobante',
      'Ricocan 22 kg', 'Ricocan 15 kg', 'Ricocat 15 kg', 'Ricocat 9 kg', 'Sacos total', 'Kg', 'Sacos requeridos', 'Es canje', 'Vales', 'Códigos de vale',
      'Compra en Lima', 'Cumple mecánica', 'Validación', 'Validado por', 'Alertas', 'Observaciones', 'Foto comprobante', 'Foto cliente', 'Firma'];
    const fila = p => { const s = sacos(p); return [p.fecha, p.hora, p.mercado, txt(p.puesto), p.impulsadora, p.cliente, txt(p.dni), txt(p.celular), p.tipo_comprobante, txt(p.n_comprobante),
      s.rc22, s.rc15, s.rt15, s.rt9, s.tot, s.kg, num(p.sacos_requeridos) || '', esCanje(p) ? 'SÍ' : 'NO', num(p.vales), txt(p.codigos_vale),
      p.compra_lima, p.cumple_mecanica, p.validacion, p.validado_por, p.alertas, p.observaciones, p.foto_comprobante, p.foto_entrega, p.firma]; };
    const anchos = [11, 6, 24, 30, 14, 26, 11, 11, 13, 14, 9, 9, 9, 9, 9, 8, 10, 8, 6, 14, 10, 12, 11, 16, 20, 30, 40, 40, 40];
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, [head, ...C.map(fila)], anchos), 'Detalle de canjes');
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, [head, ...P.map(fila)], anchos), 'Todos los pedidos');

    const hoy = new Date().toLocaleDateString('en-CA', { timeZone: 'America/Lima' });
    const nom = o.archivo || `canjes-rinti_${hoy}${o.mercado ? '_' + o.mercado.replace(/\s+/g, '-') : ''}.xlsx`;
    XLSX.writeFile(wb, nom);
    return { canjes: C.length, pedidos: P.length, archivo: nom };
  }
  // ---------- Equipo en campo: marcaciones con selfies ----------
  let libX = null;
  function cargarExcelJS() {
    if (window.ExcelJS) return Promise.resolve();
    return libX = libX || new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
      s.onload = res; s.onerror = () => { libX = null; rej(new Error('No se pudo cargar el generador de Excel. Revisa tu conexión.')); };
      document.head.appendChild(s);
    });
  }
  const idFoto = u => (String(u || '').match(/\/d\/([\w-]+)/) || [])[1] || '';
  const aMin = t => { const [h, m] = String(t || '').split(':').map(Number); return h * 60 + (m || 0); };
  const hh = t => { const [h, m] = String(t || '').split(':'); return h ? String(+h).padStart(2, '0') + ':' + (m || '00') : ''; };
  const hhmm = n => String(Math.floor(n / 60)).padStart(2, '0') + ':' + String(Math.round(n % 60)).padStart(2, '0');
  const fecha = f => { const [a, m, d] = String(f).split('-').map(Number); return new Date(Date.UTC(a, m - 1, d)); };
  const ALTO = 96; // alto de la miniatura en px
  // Reduce la foto a miniatura JPEG (base64 sin cabecera) para que el archivo pese poco.
  function miniatura(dataUrl) {
    return new Promise(res => {
      const img = new Image();
      img.onload = () => {
        const s = ALTO / img.naturalHeight, c = document.createElement('canvas');
        c.width = Math.max(1, Math.round(img.naturalWidth * s)); c.height = ALTO;
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        res({ b64: c.toDataURL('image/jpeg', 0.8).split(',')[1], w: c.width, h: ALTO });
      };
      img.onerror = () => res(null);
      img.src = dataUrl;
    });
  }

  // o: { asistencia, dias (fechas de operación), hoy, periodo, mercado, usuario, archivo, traerFotos(ids) => {id: dataUrl}, progreso(txt) }
  async function descargarEquipo(o) {
    await cargarExcelJS();
    const prog = o.progreso || (() => {});
    const A = (o.asistencia || []).map(r => Object.assign({}, r, { hora: hh(r.hora), _id: idFoto(r.foto) }))
      .sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));
    if (!A.length) throw new Error('No hay marcaciones en el periodo elegido.');

    // Fotos: se piden al servidor por tandas
    const ids = [...new Set(A.map(r => r._id).filter(Boolean))], minis = {};
    const TANDA = 40;
    for (let i = 0; i < ids.length; i += TANDA * 3) {
      prog(`Bajando selfies ${Math.min(ids.length, i + TANDA * 3)} de ${ids.length}…`);
      const lotes = [];
      for (let j = i; j < Math.min(ids.length, i + TANDA * 3); j += TANDA) lotes.push(o.traerFotos(ids.slice(j, j + TANDA)).catch(() => ({})));
      for (const r of await Promise.all(lotes)) for (const id in r) minis[id] = await miniatura(r[id]);
    }
    prog('Armando el Excel…');

    // Días por impulsadora: primera entrada y última salida
    const porDia = new Map();
    A.forEach(r => {
      const k = r.fecha + '|' + r.impulsadora;
      if (!porDia.has(k)) porDia.set(k, { fecha: r.fecha, imp: r.impulsadora, mer: r.mercado, E: [], S: [] });
      porDia.get(k)[r.marca === 'Salida' ? 'S' : 'E'].push(r);
    });
    const det = [...porDia.values()].map(d => {
      const e = d.E[0] || null, s = d.S[d.S.length - 1] || null;
      let horas = 0, estado;
      if (e && s && aMin(s.hora) > aMin(e.hora)) { horas = (aMin(s.hora) - aMin(e.hora)) / 60; estado = 'Completa'; }
      else if (e && !s) estado = d.fecha === o.hoy ? 'En curso' : 'Sin salida (no suma horas)';
      else if (s && !e) estado = 'Sin entrada (no suma horas)';
      else estado = 'Revisar horas';
      return Object.assign(d, { e, s, horas, estado, extra: d.E.length + d.S.length - (e ? 1 : 0) - (s ? 1 : 0) });
    });

    const wb = new ExcelJS.Workbook(); wb.creator = 'Free · Canje Rinti';
    const AZUL = 'FF1E22AA', VERDE = 'FFD9F2E3', AMBAR = 'FFFFF1CC', ROJO = 'FFFADBD8', CELESTE = 'FFDCE3F7';
    const borde = { top: { style: 'thin', color: { argb: 'FFD9D9D9' } }, left: { style: 'thin', color: { argb: 'FFD9D9D9' } }, bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } }, right: { style: 'thin', color: { argb: 'FFD9D9D9' } } };
    const fill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: c } });
    const sub = `${o.periodo || 'Toda la campaña'} · ${o.mercado || 'Todos los mercados'} · generado el ${new Date().toLocaleString('es-PE')}${o.usuario ? ' por ' + o.usuario : ''}`;
    const imgIds = {};
    const ponerFoto = (ws, id, col, row) => {
      const m = minis[id]; if (!m) { ws.getCell(row, col).value = id ? 'Sin foto' : ''; return; }
      if (!imgIds[id]) imgIds[id] = wb.addImage({ base64: m.b64, extension: 'jpeg' });
      ws.addImage(imgIds[id], { tl: { col: col - 1 + 0.05, row: row - 1 + 0.05 }, ext: { width: m.w, height: m.h } });
    };
    const link = (cell, url, txt) => { if (url) { cell.value = { text: txt, hyperlink: url }; cell.font = { color: { argb: 'FF1155CC' }, underline: true }; } };
    const mapa = r => r && r.lat ? `https://maps.google.com/?q=${r.lat},${r.lng}` : '';
    function hoja(nombre, titulo, cols) {
      const ws = wb.addWorksheet(nombre, { views: [{ state: 'frozen', xSplit: 2, ySplit: 4, showGridLines: false }] });
      ws.getCell('A1').value = titulo; ws.getCell('A1').font = { bold: true, size: 14, color: { argb: AZUL } };
      ws.getCell('A2').value = sub; ws.getCell('A2').font = { italic: true, color: { argb: 'FF666666' } };
      const h = ws.getRow(4); h.height = 32;
      cols.forEach(([t, w], i) => {
        const c = h.getCell(i + 1); c.value = t; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = fill(AZUL);
        c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; c.border = borde; ws.getColumn(i + 1).width = w;
      });
      return ws;
    }
    const fila = (ws, n, vals, centro) => {
      const r = ws.getRow(n);
      vals.forEach((v, i) => { const c = r.getCell(i + 1); if (v !== undefined) c.value = v; c.border = borde; c.alignment = { vertical: 'middle', wrapText: true, horizontal: centro && centro.includes(i + 1) ? 'center' : undefined }; });
      return r;
    };

    // 1. Resumen
    const dias = o.dias || [], esp = dias.length * 8;
    const ws1 = hoja('Resumen', 'Equipo en campo · Resumen de asistencia y horas', [['Impulsadora', 18], ['Mercado', 26], ['Días con marcación', 12], ['Días completos', 12], ['Días sin salida', 12], ['Horas efectivas', 12], ['Horas esperadas (48 h/sem)', 14], ['% cumplimiento', 13], ['Entrada promedio', 12], ['Salida promedio', 12], ['Primera marcación', 14], ['Última marcación', 14]]);
    const porImp = new Map(); det.forEach(d => { if (!porImp.has(d.imp)) porImp.set(d.imp, []); porImp.get(d.imp).push(d); });
    let n = 5; const T = { d: 0, c: 0, s: 0, h: 0, e: 0 };
    [...porImp.keys()].sort((a, b) => a.localeCompare(b, 'es', { numeric: true })).forEach(imp => {
      const L = porImp.get(imp), cnt = {}; L.forEach(x => cnt[x.mer] = (cnt[x.mer] || 0) + 1);
      const mer = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
      const comp = L.filter(x => x.estado === 'Completa').length, sins = L.filter(x => x.estado.startsWith('Sin salida')).length;
      const hrs = L.reduce((s, x) => s + x.horas, 0), ent = L.filter(x => x.e).map(x => aMin(x.e.hora)), sal = L.filter(x => x.s).map(x => aMin(x.s.hora));
      const p = esp ? hrs / esp : 0;
      const r = fila(ws1, n, [imp, mer, L.length, comp, sins, +hrs.toFixed(2), esp, p, ent.length ? hhmm(ent.reduce((a, b) => a + b) / ent.length) : '', sal.length ? hhmm(sal.reduce((a, b) => a + b) / sal.length) : '',
        fecha(L[0].fecha), fecha(L[L.length - 1].fecha)], [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
      r.getCell(8).numFmt = '0%'; r.getCell(8).fill = fill(p >= 0.95 ? VERDE : p >= 0.85 ? AMBAR : ROJO);
      r.getCell(11).numFmt = r.getCell(12).numFmt = 'dd/mm/yyyy';
      T.d += L.length; T.c += comp; T.s += sins; T.h += hrs; T.e += esp; n++;
    });
    const rt = fila(ws1, n, ['TOTAL', '', T.d, T.c, T.s, +T.h.toFixed(2), T.e, T.e ? T.h / T.e : 0], [3, 4, 5, 6, 7, 8]);
    rt.font = { bold: true }; rt.getCell(8).numFmt = '0%';
    ws1.getCell(n + 2, 1).value = `Horas efectivas = primera entrada a última salida de cada día; un día sin salida no suma horas. Horas esperadas = ${dias.length} días de operación × 8 h (48 h por semana). Semáforo: verde ≥ 95 %, ámbar ≥ 85 %, rojo < 85 %.`;
    ws1.getCell(n + 2, 1).font = { italic: true, color: { argb: 'FF666666' } };

    // 2. Detalle por día
    const ws2 = hoja('Detalle por día', 'Equipo en campo · Detalle diario con selfies', [['Fecha', 11], ['Impulsadora', 16], ['Mercado', 22], ['Hora entrada', 9], ['Selfie entrada', 19], ['Ubicación entrada', 12], ['Hora salida', 9], ['Selfie salida', 19], ['Ubicación salida', 12], ['Horas efectivas', 10], ['Estado', 20], ['Foto entrada', 12], ['Foto salida', 12]]);
    n = 5;
    det.slice().sort((a, b) => b.fecha.localeCompare(a.fecha) || a.imp.localeCompare(b.imp, 'es', { numeric: true })).forEach(d => {
      const r = fila(ws2, n, [fecha(d.fecha), d.imp, d.mer, d.e ? d.e.hora : '—', '', d.e ? (d.e.lat ? undefined : 'Sin GPS') : '', d.s ? d.s.hora : '—', '', d.s ? (d.s.lat ? undefined : 'Sin GPS') : '',
        +d.horas.toFixed(2), d.estado + (d.extra ? ` · ${d.extra} marcación(es) extra` : ''), '', ''], [1, 4, 6, 7, 9, 10, 12, 13]);
      r.height = ALTO * 0.75 + 6; r.getCell(1).numFmt = 'dd/mm/yyyy';
      r.getCell(11).fill = fill(d.estado === 'Completa' ? VERDE : d.estado === 'En curso' ? AMBAR : ROJO);
      if (d.e) { ponerFoto(ws2, d.e._id, 5, n); link(r.getCell(6), mapa(d.e), 'Ver mapa'); link(r.getCell(12), d.e.foto, 'Abrir foto'); }
      if (d.s) { ponerFoto(ws2, d.s._id, 8, n); link(r.getCell(9), mapa(d.s), 'Ver mapa'); link(r.getCell(13), d.s.foto, 'Abrir foto'); }
      n++;
    });
    ws2.autoFilter = { from: { row: 4, column: 1 }, to: { row: Math.max(4, n - 1), column: 13 } };

    // 3. Todas las marcaciones
    const ws3 = hoja('Todas las marcaciones', 'Equipo en campo · Todas las marcaciones', [['Fecha', 11], ['Hora', 8], ['Impulsadora', 16], ['Mercado', 22], ['Marca', 9], ['Selfie', 19], ['Latitud', 11], ['Longitud', 11], ['Precisión GPS (m)', 11], ['Mapa', 10], ['Foto', 12]]);
    n = 5;
    A.slice().reverse().forEach(x => {
      const r = fila(ws3, n, [fecha(x.fecha), x.hora, x.impulsadora, x.mercado, x.marca, '', x.lat !== '' && x.lat != null ? num(x.lat) : 'Sin GPS', x.lng !== '' && x.lng != null ? num(x.lng) : '', x.precision_m !== '' && x.precision_m != null ? num(x.precision_m) : '', '', ''], [1, 2, 5, 7, 8, 9, 10, 11]);
      r.height = ALTO * 0.75 + 6; r.getCell(1).numFmt = 'dd/mm/yyyy';
      r.getCell(5).fill = fill(x.marca === 'Entrada' ? VERDE : CELESTE);
      if (num(x.precision_m) > 100) r.getCell(9).fill = fill(AMBAR);
      ponerFoto(ws3, x._id, 6, n); link(r.getCell(10), mapa(x), 'Ver mapa'); link(r.getCell(11), x.foto, 'Abrir foto');
      n++;
    });
    ws3.autoFilter = { from: { row: 4, column: 1 }, to: { row: Math.max(4, n - 1), column: 11 } };

    const buf = await wb.xlsx.writeBuffer();
    const nom = o.archivo || 'equipo-en-campo.xlsx', a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    a.download = nom; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    return { marcaciones: A.length, dias: det.length, fotos: Object.values(minis).filter(Boolean).length, archivo: nom };
  }

  window.RintiExcel = { descargar, descargarEquipo };
})();
