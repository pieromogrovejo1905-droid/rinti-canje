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
  window.RintiExcel = { descargar };
})();
