// Configuración compartida por la app de campo (index.html) y el dashboard.
window.RINTI = {
  // Pega aquí la URL de la app web de Apps Script (termina en /exec).
  // Vacío = modo demo: la app no envía nada y el dashboard muestra datos de ejemplo.
  API_URL: 'https://script.google.com/macros/s/AKfycbyzZ3p1jaGIdczLAoohmkvxJ7APZyYR1fCotiblNzb0HV7-k80hrpcfCAUct8ullmUYuQ/exec',

  CAMPANA: 'Canje Rinti · Ricocan & Ricocat',
  // Sacos (Ricocan y/o Ricocat) para 1 vale, por mercado. Máximo 1 vale por DNI al día.
  // El backend manda la versión vigente; esta copia se usa sin conexión.
  SACOS_POR_MERCADO: {
    'Mercado Productores': 30,
    'Fiori': 22,
    'Unicachi Sur': 10,
    'Parada': 10,
    'Unicachi Norte Mayorista': 10,
    'Unicachi Norte Minorista': 10,
    'Huamantanga': 10,
    'Ciudad de Dios': 10,
    'Caquetá': 10,
  },
  KG: { ricocan: 22, ricocat: 15 },
  META_DIARIA: 6,              // se reemplaza por la hoja Config si existe
  IMPULSADORAS_PLANIFICADAS: 9,
  // Día 1 de la tabla "Canjes efectivos" del dashboard interno (antes fueron pruebas)
  INICIO_EFECTIVO: '2026-11-09',

  MERCADOS: [
    'Fiori',
    'Unicachi Norte Mayorista',
    'Unicachi Norte Minorista',
    'Unicachi Sur',
    'Huamantanga',
    'Ciudad de Dios',
    'Parada',
    'Caquetá',
    'Mercado Productores',
  ],

  // En Mercado Productores no se recorre: solo estos 8 puestos.
  PUESTOS_PRODUCTORES: [
    'Comercial Vega (bróker) · Pje. San Francisco de Asís, Pto. B38 · Directo',
    'Comercial Vega 2.ª tienda · Pje. San Francisco de Asís, Pto. 8 · Directo',
    'Sumaq Sara · Pje. San Francisco, Pto. 38 · Directo',
    'Corporación Andahuaylas · Pje. Productores, Pto. 17 · Directo',
    'Corporación Vithalex · Pje. Huancayo, Pto. 5 · Distribuidor AC',
    'Gregorio Riveros · Pje. D, Pto. 6 · Distribuidor AC',
    'Rivadeneyra Paucar · Pje. Sr. de los Milagros, Pto. 51 · Distribuidor AC',
    'Riveros Mescua · Pje. San Francisco, Pto. 27 · Distribuidor AC',
  ],

  TIPOS_INCIDENCIA: [
    'Quiebre de stock Ricocan / Ricocat',
    'Sin vales o pavos disponibles',
    'Cliente no cumple requisitos',
    'Puesto cerrado / poco movimiento',
    'Problema con distribuidor',
    'Reclamo de cliente',
    'Seguridad / clima',
    'Otro',
  ],
};
