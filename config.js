// Configuración compartida por la app de campo (index.html) y el dashboard.
window.RINTI = {
  // Pega aquí la URL de la app web de Apps Script (termina en /exec).
  // Vacío = modo demo: la app no envía nada y el dashboard muestra datos de ejemplo.
  API_URL: '',

  CAMPANA: 'Canje Rinti · Ricocan & Ricocat',
  SACOS_POR_VALE: 8,
  KG: { ricocan: 22, ricocat: 15 },
  META_DIARIA: 6,              // se reemplaza por la hoja Config si existe
  IMPULSADORAS_PLANIFICADAS: 9,

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
