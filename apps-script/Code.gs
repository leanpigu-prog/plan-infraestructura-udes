/**
 * Backend de "Plan de Infraestructura UDES" (Física + Tecnológica).
 * Pegar este código en el Google Apps Script vinculado al Google Sheet del proyecto,
 * ejecutar setup() una vez (es idempotente: también migra el Sheet ya en producción
 * sin perder datos), y desplegar como Web App (ver instrucciones adjuntas).
 *
 * Modelo de acceso:
 *  - admin: ve y edita ambos módulos completos.
 *  - infraestructura: ve y edita todo el módulo Física (todos los campus).
 *  - tecnologia: ve y edita todo el módulo Tecnológica.
 *  - responsable_proyecto: cuenta individual de una persona — ve y actualiza (estado,
 *    avance, hitos, evidencias) SOLO los proyectos (de cualquier módulo) donde su
 *    "Nombre asociado" (columna en Usuarios) coincide con el campo "Responsable" del
 *    proyecto. No puede crear ni eliminar proyectos. Estas cuentas se agregan MANUALMENTE en la
 *    hoja "Usuarios" (Usuario, Clave, Campus vacío, Rol="responsable_proyecto",
 *    Nombre asociado="Nombre exacto tal como se escribió en el campo Responsable").
 */

// ID del Google Sheet (se toma del final de su URL: .../d/ESTE_ID/edit).
// Necesario porque SpreadsheetApp.getActiveSpreadsheet() devuelve null cuando
// el script corre como Web App (sin una hoja abierta por un usuario).
const SPREADSHEET_ID = '1S6rO2pOxsSrgCDjGxdAe2CoLWZjQ_zcfufCU1TL1Y3w';

const SHEET_PROYECTOS = 'Proyectos';
const SHEET_PROYECTOS_TI = 'ProyectosTI';
const SHEET_USUARIOS = 'Usuarios';
const SHEET_HISTORIAL = 'Historial_Seguimiento';
const SHEET_HITOS = 'Hitos';
const SHEET_EVIDENCIAS = 'Evidencias';

function getSs_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

const ESTADOS_VALIDOS = ['Sin iniciar', 'En desarrollo', 'Ejecutado', 'Atrasado'];
const PRIORIDADES_VALIDAS = ['Alta', 'Media', 'Baja'];
const HITO_ESTADOS_VALIDOS = ['Pendiente', 'Cumplido', 'Atrasado'];
const EVIDENCIA_TIPOS_VALIDOS = ['Acta', 'Informe', 'Fotografía', 'Contrato/Convenio', 'Soporte de pago', 'Documento técnico', 'Otro'];

// Columna de "Proyectos"/"ProyectosTI" que identifica a la persona responsable de un
// proyecto — se usa para filtrar qué ve una cuenta individual (rol responsable_proyecto).
const CAMPOS_RESPONSABLES = ['Responsable'];

const PROGRAMAS_VALIDOS = [
  'Movilidad y accesibilidad',
  'Expansión académica',
  'Equipamiento cultural y deportivo',
  'Bienestar y vida universitaria',
  'Identidad e imagen institucional',
  'Ciencia, tecnología e innovación',
  'Reserva de suelo',
];

const LINEAS_TI_VALIDAS = [
  'Infraestructura de procesamiento y almacenamiento',
  'Plataformas institucionales',
  'Servicios tecnológicos',
  'Seguridad de la infraestructura tecnológica',
  'Analítica, datos e IA',
  'Infraestructura tecnológica para docencia e investigación',
];

const CAMPUS_VALIDOS = ['Bucaramanga', 'Cúcuta', 'Valledupar'];

/**
 * Config central por módulo — evita duplicar listarProyectos_/crearProyecto_/etc.
 * para Física y Tecnológica: ambas comparten la misma lógica, solo cambia
 * la hoja y el nombre de la columna "categoría".
 */
const MODULOS = {
  fisica: {
    sheet: SHEET_PROYECTOS,
    categoriaCol: 'Programa',
    categoriasValidas: PROGRAMAS_VALIDOS,
    moduloLabel: 'Física',
  },
  tecnologica: {
    sheet: SHEET_PROYECTOS_TI,
    categoriaCol: 'Línea',
    categoriasValidas: LINEAS_TI_VALIDAS,
    moduloLabel: 'Tecnológica',
  },
};

// Qué módulos puede ver/operar cada rol (el filtrado fino de responsable_proyecto ocurre
// aparte, en proyectoVisiblePara_/puedeEditarProyecto_, no aquí).
const ROLE_MODULOS = {
  admin: ['fisica', 'tecnologica'],
  infraestructura: ['fisica'],
  tecnologia: ['tecnologica'],
  responsable_proyecto: ['fisica', 'tecnologica'],
};

const PROYECTOS_BASE = [
  // N°, Proyecto, Campus, Descripción, Programa, Estado inicial, % avance inicial
  [1, 'Edificio Chitareros', 'Bucaramanga', 'Infraestructura académica con aulas activas y laboratorios flexibles, preparada para ampliaciones futuras.', 'Expansión académica', 'Sin iniciar', 0],
  [2, 'Vía al Polvorín', 'Bucaramanga', 'Conexión vial estratégica entre parqueaderos, zonas deportivas y equipamientos institucionales.', 'Movilidad y accesibilidad', 'Sin iniciar', 0],
  [3, 'Deprimido', 'Bucaramanga', 'Paso inferior para mejorar la seguridad y fluidez vehicular y peatonal.', 'Movilidad y accesibilidad', 'Sin iniciar', 0],
  [4, 'Puente', 'Bucaramanga', 'Conexión directa entre el campus académico y las zonas deportivas y culturales.', 'Movilidad y accesibilidad', 'Sin iniciar', 0],
  [5, 'Acceso principal', 'Bucaramanga', 'Reingeniería del ingreso institucional con mayor seguridad y eficiencia operativa.', 'Movilidad y accesibilidad', 'Sin iniciar', 0],
  [6, 'Auditorio', 'Bucaramanga', 'Escenario multipropósito para eventos académicos y culturales de gran formato.', 'Equipamiento cultural y deportivo', 'Sin iniciar', 0],
  [7, 'Coliseo', 'Bucaramanga', 'Equipamiento deportivo y cultural de aforo masivo y uso multifuncional.', 'Equipamiento cultural y deportivo', 'Sin iniciar', 0],
  [8, 'Canchas', 'Bucaramanga', 'Espacios especializados para prácticas deportivas con especificaciones técnicas para uso continuo.', 'Equipamiento cultural y deportivo', 'Sin iniciar', 0],
  [9, 'Residencias estudiantiles', 'Bucaramanga', 'Alojamiento con ambientes de estudio, conectividad y servicios complementarios.', 'Bienestar y vida universitaria', 'Sin iniciar', 0],
  [10, 'CDT INNOVATEC', 'Bucaramanga', 'Centro de Desarrollo Tecnológico consolidado en 2025 (Acuerdo 009), orientado a seguridad alimentaria y energías renovables.', 'Ciencia, tecnología e innovación', 'Sin iniciar', 0],
  [11, 'Cafetería', 'Cúcuta', 'Ampliación del espacio de encuentro universitario en la plaza central para mejorar bienestar y permanencia.', 'Bienestar y vida universitaria', 'En desarrollo', 40],
  [12, 'Ascensor edificio administrativo', 'Cúcuta', 'Mejora de accesibilidad y conectividad vertical en el bloque administrativo.', 'Movilidad y accesibilidad', 'En desarrollo', 40],
  [13, 'Compra Edificio Cruz Roja', 'Cúcuta', 'Ampliación de aulas, laboratorios y espacios de extensión mediante adquisición contigua estratégica.', 'Expansión académica', 'Ejecutado', 100],
  [14, 'Fachada Institucional', 'Cúcuta', 'Renovación del borde institucional para reforzar identidad, iluminación y accesos.', 'Identidad e imagen institucional', 'Sin iniciar', 0],
  [15, 'Modernización de parqueaderos', 'Cúcuta', 'Optimización de flujos, señalización e infraestructura vial interna.', 'Movilidad y accesibilidad', 'Sin iniciar', 0],
  [16, 'Pompeyanos (paso peatonal)', 'Cúcuta', 'Conexión peatonal segura entre el campus y su entorno inmediato.', 'Movilidad y accesibilidad', 'Sin iniciar', 0],
  [17, 'Edificio Calle 14 (3 pisos)', 'Valledupar', 'Construcción de nuevo bloque académico-administrativo.', 'Expansión académica', 'Sin iniciar', 0],
  [18, 'Adquisición lote Cra 5', 'Valledupar', 'Compra de predio para expansión futura del campus.', 'Reserva de suelo', 'Sin iniciar', 0],
  [19, 'Fachada Institucional', 'Valledupar', 'Adecuación arquitectónica y mejoramiento de accesos.', 'Identidad e imagen institucional', 'Sin iniciar', 0],
];

// ─── PROYECTOS BASE: Infraestructura Tecnológica ──────────────────────────────
// Campos por fila: [N°, Proyecto, Línea, Campus, Descripción, Costo estimado,
//                   Fuente de financiación, Estado inicial, Prioridad, % avance]
const PROYECTOS_BASE_TI = [

  // LÍNEA 1: Infraestructura de procesamiento y almacenamiento
  [1, 'Infraestructura de servidores (Moodle + DR)',
    'Infraestructura de procesamiento y almacenamiento', 'Todos',
    'Adquisición e implementación de infraestructura de servidores para soportar Moodle, plataformas institucionales y el plan de recuperación de desastres (DR) de la UDES.',
    4000000000, 'Gestión TIC / Plan Estratégico', 'Sin iniciar', 'Alta', 0],
  [2, 'Dotación tecnológica laboratorios clínicos y simulación (VAL)',
    'Infraestructura de procesamiento y almacenamiento', 'Valledupar',
    'Adquisición de equipos tecnológicos para prácticas de laboratorio en ciencias de la salud, incluyendo simuladores clínicos y dispositivos especializados.',
    296529640, 'Facultad de Ciencias Médicas y de la Salud', 'Sin iniciar', 'Alta', 0],

  // LÍNEA 2: Plataformas institucionales
  [3, 'Optimización plataformas virtuales y Moodle (INIC-A)',
    'Plataformas institucionales', 'Todos',
    'Optimización y modernización de las plataformas de aprendizaje virtual Moodle y sistemas de evaluación en línea, unificando versiones y mejorando la experiencia de usuario.',
    80000000, 'PDI - Transformación Curricular / Política ITD', 'Sin iniciar', 'Alta', 0],
  [4, 'Ruta institucional de virtualización y plataformas (INIC-B)',
    'Plataformas institucionales', 'Todos',
    'Fortalecimiento de la ruta institucional de virtualización mediante integración de plataformas tecnológicas para el aprendizaje a distancia y semipresencial.',
    130000000, 'PDI - Transformación Curricular / Política de Enseñanza', 'Sin iniciar', 'Alta', 0],
  [5, 'Ajuste modelo de evaluación profesoral MEIPU (INIC-E)',
    'Plataformas institucionales', 'Todos',
    'Módulo de software para digitalizar el Modelo de Evaluación Integral de la Práctica de Enseñanza (MEIPU), integrando retroalimentación de calidad docente.',
    76000000, 'PDI - Transformación Curricular / Política de Enseñanza', 'Sin iniciar', 'Media', 0],
  [6, 'Plan prospectivo de infraestructura tecnológica del programa (CUC)',
    'Plataformas institucionales', 'Cúcuta',
    'Fortalecimiento de infraestructura tecnológica al servicio de los programas del campus Cúcuta, incluyendo plataformas digitales y sistemas de información.',
    713003004, 'Facultad de Ciencias Médicas y de la Salud', 'Sin iniciar', 'Alta', 0],

  // LÍNEA 3: Servicios tecnológicos
  [7, 'Digitalización y automatización de procesos institucionales (INIC-F)',
    'Servicios tecnológicos', 'Todos',
    'Digitalización y automatización de procesos académico-administrativos para mejorar la eficiencia institucional mediante flujos de trabajo electrónicos y firma digital.',
    240000000, 'PDI - Transformación Curricular / Política ITD', 'Sin iniciar', 'Alta', 0],
  [8, 'Vigilancia tecnológica e innovación disruptiva institucional',
    'Servicios tecnológicos', 'Todos',
    'Implementación de proceso permanente de vigilancia tecnológica para identificar tendencias disruptivas y orientar la toma de decisiones estratégicas en TI.',
    216000000, 'Política de Innovación, Transformación y Disrupción', 'Sin iniciar', 'Alta', 0],
  [9, 'Sistema de proctoring para evaluaciones en línea (VAL)',
    'Servicios tecnológicos', 'Valledupar',
    'Incorporación de herramientas de proctoring para garantizar la integridad y vigilancia en evaluaciones a distancia del campus Valledupar.',
    50000000, 'Facultad de Ingenierías', 'Sin iniciar', 'Alta', 0],

  // LÍNEA 4: Seguridad de la infraestructura tecnológica
  [10, 'Protección de datos y privacidad en IA institucional',
    'Seguridad de la infraestructura tecnológica', 'Todos',
    'Garantizar el cumplimiento normativo en protección de datos en el uso de IA institucional, incluyendo auditorías de sesgos algorítmicos y marcos éticos.',
    560000000, 'Política de Inteligencia Artificial', 'Sin iniciar', 'Alta', 0],
  [11, 'Gobierno y seguridad de datos institucionales',
    'Seguridad de la infraestructura tecnológica', 'Todos',
    'Marcos de seguridad de la información, gestión de riesgos y protocolos de privacidad para el cumplimiento normativo en el tratamiento de datos institucionales.',
    70000000, 'Política de Gobierno de Datos', 'Sin iniciar', 'Alta', 0],

  // LÍNEA 5: Analítica, datos e IA
  [12, 'Herramientas emergentes IA y simuladores en aprendizaje (INIC-G)',
    'Analítica, datos e IA', 'Todos',
    'Incorporación de tecnologías emergentes (IA, simuladores, realidad virtual/aumentada) en los procesos de enseñanza-aprendizaje con pilotos en programas seleccionados.',
    480000000, 'PDI - Transformación Curricular / Política ITD', 'Sin iniciar', 'Alta', 0],
  [13, 'Modelo de alertas curriculares y permanencia estudiantil (INIC-C)',
    'Analítica, datos e IA', 'Todos',
    'Módulo de seguimiento y alertas curriculares para detección temprana de riesgo de deserción y apoyo personalizado a estudiantes con bajo rendimiento académico.',
    130000000, 'PDI - Transformación Curricular / Política ITD', 'Sin iniciar', 'Media', 0],
  [14, 'Seguimiento resultados de aprendizaje CMIRA (INIC-D)',
    'Analítica, datos e IA', 'Todos',
    'Módulo de seguimiento de resultados dentro del Marco Institucional de Resultados de Aprendizaje (CMIRA), integrado al sistema de evaluación académica.',
    0, 'PDI - Transformación Curricular / Política de Enseñanza', 'Sin iniciar', 'Media', 0],
  [15, 'Arquitectura e integración de datos institucionales (ETL)',
    'Analítica, datos e IA', 'Todos',
    'Unificación de arquitectura de datos, integración de sistemas heterogéneos y reducción de redundancia mediante pipelines ETL y repositorio centralizado.',
    250000000, 'Política de Gobierno de Datos', 'Sin iniciar', 'Alta', 0],
  [16, 'Sistema de analítica predictiva para retención estudiantil',
    'Analítica, datos e IA', 'Todos',
    'Minería de datos y analítica predictiva sobre comportamiento académico para identificar estudiantes en riesgo y activar estrategias personalizadas de acompañamiento.',
    120000000, 'Vicerrectoría de Enseñanza', 'Sin iniciar', 'Alta', 0],

  // LÍNEA 6: Infraestructura tecnológica para docencia e investigación
  [17, 'Renovación progresiva de equipos de laboratorio institucional (INIC-H)',
    'Infraestructura tecnológica para docencia e investigación', 'Todos',
    'Renovación y modernización progresiva de los equipos de laboratorio utilizados en formación práctica, garantizando condiciones técnicas actualizadas en los tres campus.',
    3600000000, 'PDI - Transformación Curricular', 'Sin iniciar', 'Alta', 0],
  [18, 'Actualización recursos digitales y cursos virtuales portables (BUC)',
    'Infraestructura tecnológica para docencia e investigación', 'Bucaramanga',
    'Actualización de cursos virtuales portables con accesibilidad y subtitulación, y nuevos recursos digitales para la Facultad de Ciencias Sociales de Bucaramanga.',
    1625249304, 'Facultad de Ciencias Sociales', 'Sin iniciar', 'Alta', 0],
  [19, 'Fortalecimiento bibliográfico y bases de datos académicas (CUC)',
    'Infraestructura tecnológica para docencia e investigación', 'Cúcuta',
    'Adquisición y actualización de recursos bibliográficos y bases de datos digitales para la comunidad académica del campus Cúcuta.',
    1257926530, 'Facultad de Ciencias Médicas y de la Salud', 'Sin iniciar', 'Alta', 0],
  [20, 'Actualización HW/SW e informática especializada (BUC)',
    'Infraestructura tecnológica para docencia e investigación', 'Bucaramanga',
    'Actualización de recursos informáticos (hardware y software especializado) para laboratorios de computación de programas académicos de FEAC.',
    278128893, 'Facultad de Ciencias Económicas, Administrativas y Contables', 'Sin iniciar', 'Alta', 0],
];

// Cuentas de nivel institucional (setup() las agrega si faltan, sin duplicar).
// Las cuentas individuales de responsable_proyecto se agregan manualmente en el Sheet.
const USUARIOS_BASE = [
  ['admin', 'CAMBIAR-CLAVE-ADMIN', 'Todos', 'admin', ''],
  ['infraestructura', 'CAMBIAR-CLAVE-INFRA', 'Todos', 'infraestructura', ''],
  ['tecnologia', 'CAMBIAR-CLAVE-TI', 'Todos', 'tecnologia', ''],
];

const PROYECTOS_HEADERS = [
  'N°', 'Proyecto', 'Campus', 'Descripción', 'Programa',
  'Costo estimado', 'Fuente de financiación',
  'Fecha inicio', 'Fecha fin', 'Fase actual',
  'Responsable', 'Estado', 'Prioridad', '% avance', 'Última actualización',
];

const PROYECTOS_TI_HEADERS = [
  'N°', 'Proyecto', 'Línea', 'Campus', 'Descripción',
  'Costo estimado', 'Fuente de financiación',
  'Fecha inicio', 'Fecha fin', 'Fase actual', 'Responsable',
  'Estado', 'Prioridad', '% avance', 'Última actualización',
];

const USUARIOS_HEADERS = ['Usuario', 'Clave', 'Campus', 'Rol', 'Nombre asociado'];

const HISTORIAL_HEADERS = ['Fecha', 'N° proyecto', 'Proyecto', 'Usuario', 'Estado anterior', 'Estado nuevo', '% avance', 'Comentario', 'Módulo'];

const HITOS_HEADERS = ['ID', 'Módulo', 'N° proyecto', 'Proyecto', 'Título', 'Periodo', 'Fecha objetivo', 'Estado', 'Responsable', 'Comentario', 'Fecha creación', 'Creado por'];

const EVIDENCIAS_HEADERS = ['ID', 'Módulo', 'N° proyecto', 'Proyecto', 'Tipo', 'Descripción', 'Link', 'Fecha', 'Subido por'];

/**
 * Ejecutar desde el editor de Apps Script (menú Ejecutar > setup) para crear las hojas,
 * encabezados, validaciones y datos base — o para MIGRAR un Sheet ya en producción:
 * agrega columnas/hojas/usuarios nuevos sin duplicar ni perder nada de lo existente.
 * Es seguro volver a ejecutarla las veces que sea necesario.
 *
 * Esta función NO elimina cuentas existentes (por ejemplo, cuentas de campus antiguas):
 * si ya no las quieres, bórralas tú mismo desde la hoja "Usuarios".
 */
function setup() {
  const ss = getSs_();

  const proyectosSheet = getOrCreateSheet_(ss, SHEET_PROYECTOS);
  if (proyectosSheet.getLastRow() === 0) {
    proyectosSheet.appendRow(PROYECTOS_HEADERS);
    PROYECTOS_BASE.forEach(function (p) {
      const now = new Date();
      proyectosSheet.appendRow([
        p[0], p[1], p[2], p[3], p[4],
        '', '',
        '', '', '',
        '', p[5], '', p[6], now
      ]);
    });
    aplicarValidaciones_(proyectosSheet);
    proyectosSheet.setFrozenRows(1);
    proyectosSheet.autoResizeColumns(1, PROYECTOS_HEADERS.length);
  } else {
    asegurarColumnas_(proyectosSheet, PROYECTOS_HEADERS);
  }

  const proyectosTiSheet = getOrCreateSheet_(ss, SHEET_PROYECTOS_TI);
  if (proyectosTiSheet.getLastRow() === 0) {
    proyectosTiSheet.appendRow(PROYECTOS_TI_HEADERS);
    const nowTi = new Date();
    PROYECTOS_BASE_TI.forEach(function (p) {
      proyectosTiSheet.appendRow([
        p[0],          // N°
        p[1],          // Proyecto
        p[2],          // Línea
        p[3],          // Campus
        p[4],          // Descripción
        p[5],          // Costo estimado
        p[6],          // Fuente de financiación
        '',            // Fecha inicio
        '',            // Fecha fin
        'Formulación', // Fase actual
        '',            // Responsable
        p[7],          // Estado
        p[8],          // Prioridad
        p[9],          // % avance
        nowTi,         // Última actualización
      ]);
    });
    aplicarValidaciones_(proyectosTiSheet);
    proyectosTiSheet.setFrozenRows(1);
    proyectosTiSheet.autoResizeColumns(1, PROYECTOS_TI_HEADERS.length);
  } else {
    asegurarColumnas_(proyectosTiSheet, PROYECTOS_TI_HEADERS);
  }

  const usuariosSheet = getOrCreateSheet_(ss, SHEET_USUARIOS);
  if (usuariosSheet.getLastRow() === 0) {
    usuariosSheet.appendRow(USUARIOS_HEADERS);
    USUARIOS_BASE.forEach(function (u) { usuariosSheet.appendRow(u); });
    usuariosSheet.setFrozenRows(1);
  } else {
    asegurarColumnas_(usuariosSheet, USUARIOS_HEADERS);
    const usuariosExistentes = usuariosSheet.getDataRange().getValues().slice(1).map(function (row) { return String(row[0]).toLowerCase(); });
    USUARIOS_BASE.forEach(function (u) {
      if (usuariosExistentes.indexOf(u[0].toLowerCase()) === -1) {
        usuariosSheet.appendRow(u);
      }
    });
  }

  const historialSheet = getOrCreateSheet_(ss, SHEET_HISTORIAL);
  if (historialSheet.getLastRow() === 0) {
    historialSheet.appendRow(HISTORIAL_HEADERS);
    historialSheet.setFrozenRows(1);
  } else {
    asegurarColumnas_(historialSheet, HISTORIAL_HEADERS);
    backfillModuloHistorial_(historialSheet);
  }

  const hitosSheet = getOrCreateSheet_(ss, SHEET_HITOS);
  if (hitosSheet.getLastRow() === 0) {
    hitosSheet.appendRow(HITOS_HEADERS);
    aplicarValidacionLista_(hitosSheet, 'Estado', HITO_ESTADOS_VALIDOS);
    hitosSheet.setFrozenRows(1);
  }

  const evidenciasSheet = getOrCreateSheet_(ss, SHEET_EVIDENCIAS);
  if (evidenciasSheet.getLastRow() === 0) {
    evidenciasSheet.appendRow(EVIDENCIAS_HEADERS);
    aplicarValidacionLista_(evidenciasSheet, 'Tipo', EVIDENCIA_TIPOS_VALIDOS);
    evidenciasSheet.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  Logger.log('Setup/migración completo. Recuerda cambiar las claves en la hoja "Usuarios" (incluidas infraestructura y tecnologia).');
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

/**
 * Migración no destructiva: agrega al final cualquier columna de headersEsperados
 * que no exista todavía en la hoja, sin tocar las columnas/filas ya presentes.
 */
function asegurarColumnas_(sheet, headersEsperados) {
  const lastCol = sheet.getLastColumn();
  const headersActuales = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const presentes = {};
  headersActuales.forEach(function (h) { presentes[h] = true; });
  let siguienteCol = lastCol;
  headersEsperados.forEach(function (h) {
    if (!presentes[h]) {
      siguienteCol += 1;
      sheet.getRange(1, siguienteCol).setValue(h);
      presentes[h] = true;
    }
  });
}

/** Rellena con 'Física' las filas de Historial_Seguimiento que quedaron sin Módulo tras la migración. */
function backfillModuloHistorial_(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colModulo = headers.indexOf('Módulo');
  if (colModulo === -1) return;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][colModulo]) {
      sheet.getRange(i + 1, colModulo + 1).setValue('Física');
    }
  }
}

/** Aplica validación de lista desplegable a las columnas "Estado" y "Prioridad" de una hoja de proyectos, buscándolas por nombre de header (no por posición fija). */
function aplicarValidaciones_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 200);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colEstado = headers.indexOf('Estado') + 1;
  const colPrioridad = headers.indexOf('Prioridad') + 1;
  if (colEstado > 0) {
    const estadoRule = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_VALIDOS, true).setAllowInvalid(false).build();
    sheet.getRange(2, colEstado, lastRow - 1, 1).setDataValidation(estadoRule);
  }
  if (colPrioridad > 0) {
    const prioridadRule = SpreadsheetApp.newDataValidation().requireValueInList(PRIORIDADES_VALIDAS, true).setAllowInvalid(false).build();
    sheet.getRange(2, colPrioridad, lastRow - 1, 1).setDataValidation(prioridadRule);
  }
}

/** Aplica validación de lista desplegable a una columna cualquiera, buscándola por nombre de header. */
function aplicarValidacionLista_(sheet, nombreColumna, valoresValidos) {
  const lastRow = Math.max(sheet.getLastRow(), 200);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = headers.indexOf(nombreColumna) + 1;
  if (col > 0) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(valoresValidos, true).setAllowInvalid(false).build();
    sheet.getRange(2, col, lastRow - 1, 1).setDataValidation(rule);
  }
}

/** Punto de entrada GET — útil para pruebas rápidas desde el navegador (?action=ping). */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'ping') {
    return jsonResponse_({ ok: true, mensaje: 'Backend activo' });
  }
  return jsonResponse_({ ok: false, error: 'Usa POST para las acciones de la aplicación.' });
}

/**
 * Punto de entrada POST. El body debe ser JSON: { action: '...', ...payload }
 * Acciones de proyectos (aceptan "modulo": 'fisica' [por defecto] | 'tecnologica'):
 *   login, listarProyectos, actualizarProyecto, crearProyecto, eliminarProyecto
 * Acciones de seguimiento (requieren "modulo"):
 *   listarSeguimiento, crearHito, actualizarHito, eliminarHito, agregarEvidencia, eliminarEvidencia
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'login') return jsonResponse_(login_(body.usuario, body.clave));
    if (action === 'listarProyectos') return jsonResponse_(listarProyectos_(body.usuario, body.clave, body.modulo));
    if (action === 'actualizarProyecto') return jsonResponse_(actualizarProyecto_(body));
    if (action === 'crearProyecto') return jsonResponse_(crearProyecto_(body));
    if (action === 'eliminarProyecto') return jsonResponse_(eliminarProyecto_(body));

    if (action === 'listarSeguimiento') return jsonResponse_(listarSeguimiento_(body.usuario, body.clave, body.modulo, body.numero));
    if (action === 'crearHito') return jsonResponse_(crearHito_(body));
    if (action === 'actualizarHito') return jsonResponse_(actualizarHito_(body));
    if (action === 'eliminarHito') return jsonResponse_(eliminarHito_(body));
    if (action === 'agregarEvidencia') return jsonResponse_(agregarEvidencia_(body));
    if (action === 'eliminarEvidencia') return jsonResponse_(eliminarEvidencia_(body));

    return jsonResponse_({ ok: false, error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Busca al usuario en la hoja "Usuarios" y valida la clave. Incluye "Nombre asociado" (relevante para responsable_proyecto). */
function autenticar_(usuario, clave) {
  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_USUARIOS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colNombre = headers.indexOf('Nombre asociado');
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).toLowerCase() === String(usuario).toLowerCase() && String(row[1]) === String(clave)) {
      return {
        usuario: row[0],
        campus: row[2],
        rol: row[3],
        nombreAsociado: colNombre !== -1 ? row[colNombre] : '',
      };
    }
  }
  return null;
}

function login_(usuario, clave) {
  const sesion = autenticar_(usuario, clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };
  return { ok: true, sesion: sesion };
}

/** ¿La sesión tiene permitido ver/operar sobre este módulo (a nivel general)? */
function puedeAccederModulo_(sesion, modulo) {
  const permitidos = ROLE_MODULOS[sesion.rol] || [];
  return permitidos.indexOf(modulo) !== -1;
}

/** ¿La sesión puede crear o eliminar proyectos completos en este módulo? (responsable_proyecto NO puede). */
function puedeGestionarProyectos_(sesion, modulo) {
  if (sesion.rol === 'admin') return true;
  if (sesion.rol === 'infraestructura') return modulo === 'fisica';
  if (sesion.rol === 'tecnologia') return modulo === 'tecnologica';
  return false;
}

/** ¿Este proyecto puntual es visible para la sesión? (admin/infraestructura/tecnologia ven todo su alcance; responsable_proyecto solo lo suyo). */
function proyectoVisiblePara_(sesion, modulo, proyecto) {
  if (sesion.rol === 'admin') return true;
  if (sesion.rol === 'infraestructura') return modulo === 'fisica';
  if (sesion.rol === 'tecnologia') return modulo === 'tecnologica';
  if (sesion.rol === 'responsable_proyecto') {
    if (!sesion.nombreAsociado) return false;
    return CAMPOS_RESPONSABLES.some(function (campo) { return proyecto[campo] && proyecto[campo] === sesion.nombreAsociado; });
  }
  return false;
}

/** ¿La sesión puede editar/actualizar (seguimiento incluido) este proyecto puntual? Misma regla que proyectoVisiblePara_. */
function puedeEditarProyecto_(sesion, modulo, proyecto) {
  return proyectoVisiblePara_(sesion, modulo, proyecto);
}

/** Convierte una fila cruda + colIndex en un objeto {header: valor}. */
function filaAObjeto_(fila, colIndex) {
  const obj = {};
  Object.keys(colIndex).forEach(function (h) { obj[h] = fila[colIndex[h]]; });
  return obj;
}

/** Busca la fila de un proyecto (por N°) en la hoja del módulo dado. Devuelve {fila, colIndex} o null. */
function buscarProyecto_(cfg, numero) {
  const ss = getSs_();
  const sheet = ss.getSheetByName(cfg.sheet);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = {};
  headers.forEach(function (h, idx) { colIndex[h] = idx; });
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colIndex['N°']]) === Number(numero)) {
      return { fila: data[i], colIndex: colIndex };
    }
  }
  return null;
}

/** Devuelve los proyectos visibles para el usuario en el módulo indicado ('fisica' por defecto). */
function listarProyectos_(usuario, clave, moduloParam) {
  const sesion = autenticar_(usuario, clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = moduloParam || 'fisica';
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };
  if (!puedeAccederModulo_(sesion, modulo)) return { ok: false, error: 'No tienes acceso a este módulo.' };

  const ss = getSs_();
  const sheet = ss.getSheetByName(cfg.sheet);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const proyectos = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const proyecto = {};
    headers.forEach(function (h, idx) { proyecto[h] = row[idx]; });
    if (proyectoVisiblePara_(sesion, modulo, proyecto)) {
      proyectos.push(proyecto);
    }
  }

  return { ok: true, sesion: sesion, proyectos: proyectos };
}

/**
 * Actualiza los campos editables de un proyecto (identificado por N° dentro de su módulo) y
 * registra el cambio en Historial_Seguimiento. Respeta puedeEditarProyecto_ (admin/infraestructura/
 * tecnologia según el módulo, o responsable_proyecto solo si aparece como responsable del proyecto).
 */
function actualizarProyecto_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = body.modulo || 'fisica';
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };
  if (!puedeAccederModulo_(sesion, modulo)) return { ok: false, error: 'No tienes acceso a este módulo.' };

  const ss = getSs_();
  const sheet = ss.getSheetByName(cfg.sheet);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = {};
  headers.forEach(function (h, idx) { colIndex[h] = idx; });

  const numeroProyecto = Number(body.numero);
  let filaEncontrada = -1;
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colIndex['N°']]) === numeroProyecto) { filaEncontrada = i; break; }
  }
  if (filaEncontrada === -1) return { ok: false, error: 'Proyecto no encontrado: ' + body.numero };

  const filaActual = data[filaEncontrada];
  const proyectoActual = filaAObjeto_(filaActual, colIndex);
  if (!puedeEditarProyecto_(sesion, modulo, proyectoActual)) {
    return { ok: false, error: 'No tienes permiso para editar este proyecto.' };
  }

  if (body.prioridad && PRIORIDADES_VALIDAS.indexOf(body.prioridad) === -1) {
    return { ok: false, error: 'Prioridad inválida: ' + body.prioridad };
  }
  if (body.estado && ESTADOS_VALIDOS.indexOf(body.estado) === -1) {
    return { ok: false, error: 'Estado inválido: ' + body.estado };
  }

  const estadoAnterior = filaActual[colIndex['Estado']];
  const rowNumber = filaEncontrada + 1; // 1-indexado en el Sheet, +1 por encabezado ya incluido en data

  const camposEditables = {
    'Costo estimado': 'costoEstimado',
    'Fuente de financiación': 'fuenteFinanciacion',
    'Fecha inicio': 'fechaInicio',
    'Fecha fin': 'fechaFin',
    'Fase actual': 'faseActual',
    'Responsable': 'responsable',
    'Estado': 'estado',
    'Prioridad': 'prioridad',
    '% avance': 'pctAvance',
  };

  Object.keys(camposEditables).forEach(function (columna) {
    const campoBody = camposEditables[columna];
    if (colIndex[columna] !== undefined && body[campoBody] !== undefined && body[campoBody] !== null) {
      sheet.getRange(rowNumber, colIndex[columna] + 1).setValue(body[campoBody]);
    }
  });
  sheet.getRange(rowNumber, colIndex['Última actualización'] + 1).setValue(new Date());

  const historial = ss.getSheetByName(SHEET_HISTORIAL);
  historial.appendRow([
    new Date(),
    numeroProyecto,
    filaActual[colIndex['Proyecto']],
    sesion.usuario,
    estadoAnterior,
    body.estado || estadoAnterior,
    body.pctAvance !== undefined ? body.pctAvance : '',
    body.comentario || '',
    cfg.moduloLabel
  ]);

  return { ok: true };
}

/**
 * Crea un proyecto nuevo en el módulo indicado. Solo admin/infraestructura/tecnologia pueden
 * crear proyectos (puedeGestionarProyectos_); el campus es siempre informativo y de libre
 * elección entre CAMPUS_VALIDOS o "Todos".
 */
function crearProyecto_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = body.modulo || 'fisica';
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };
  if (!puedeGestionarProyectos_(sesion, modulo)) return { ok: false, error: 'No tienes permiso para crear proyectos en este módulo.' };

  const nombreProyecto = String(body.proyecto || '').trim();
  if (!nombreProyecto) return { ok: false, error: 'El nombre del proyecto es obligatorio.' };

  const categoria = body.categoria !== undefined ? body.categoria : body.programa;
  if (cfg.categoriasValidas.indexOf(categoria) === -1) {
    return { ok: false, error: cfg.categoriaCol + ' inválido: ' + categoria };
  }

  const campus = CAMPUS_VALIDOS.concat(['Todos']).indexOf(body.campus) !== -1 ? body.campus : 'Todos';

  if (body.prioridad && PRIORIDADES_VALIDAS.indexOf(body.prioridad) === -1) {
    return { ok: false, error: 'Prioridad inválida: ' + body.prioridad };
  }

  const ss = getSs_();
  const sheet = ss.getSheetByName(cfg.sheet);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = {};
  headers.forEach(function (h, idx) { colIndex[h] = idx; });

  let maxNumero = 0;
  for (let i = 1; i < data.length; i++) {
    const n = Number(data[i][colIndex['N°']]);
    if (n > maxNumero) maxNumero = n;
  }
  const nuevoNumero = maxNumero + 1;
  const now = new Date();

  const fila = new Array(headers.length).fill('');
  fila[colIndex['N°']] = nuevoNumero;
  fila[colIndex['Proyecto']] = nombreProyecto;
  fila[colIndex[cfg.categoriaCol]] = categoria;
  if (colIndex['Campus'] !== undefined) fila[colIndex['Campus']] = campus;
  fila[colIndex['Descripción']] = String(body.descripcion || '');
  fila[colIndex['Estado']] = 'Sin iniciar';
  fila[colIndex['Prioridad']] = body.prioridad || '';
  fila[colIndex['% avance']] = 0;
  fila[colIndex['Última actualización']] = now;

  const camposCreacion = {
    'Costo estimado': 'costoEstimado',
    'Fuente de financiación': 'fuenteFinanciacion',
    'Fecha inicio': 'fechaInicio',
    'Fecha fin': 'fechaFin',
    'Responsable': 'responsable',
  };
  Object.keys(camposCreacion).forEach(function (columna) {
    if (colIndex[columna] !== undefined) {
      fila[colIndex[columna]] = body[camposCreacion[columna]] || '';
    }
  });

  sheet.appendRow(fila);

  const historial = ss.getSheetByName(SHEET_HISTORIAL);
  historial.appendRow([
    now, nuevoNumero, nombreProyecto, sesion.usuario,
    '—', 'Sin iniciar', 0,
    modulo === 'fisica' ? 'Proyecto creado (fuera del Plan Maestro original).' : 'Proyecto creado.',
    cfg.moduloLabel
  ]);

  return { ok: true, numero: nuevoNumero };
}

/**
 * Elimina un proyecto (identificado por N° dentro de su módulo) y su fila del Sheet.
 * Solo admin/infraestructura/tecnologia pueden eliminar (puedeGestionarProyectos_) —
 * un responsable_proyecto NO puede eliminar aunque el proyecto sea suyo.
 * Deja constancia en Historial_Seguimiento.
 */
function eliminarProyecto_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = body.modulo || 'fisica';
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };
  if (!puedeGestionarProyectos_(sesion, modulo)) return { ok: false, error: 'No tienes permiso para eliminar proyectos en este módulo.' };

  const ss = getSs_();
  const sheet = ss.getSheetByName(cfg.sheet);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = {};
  headers.forEach(function (h, idx) { colIndex[h] = idx; });

  const numeroProyecto = Number(body.numero);
  let filaEncontrada = -1;
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][colIndex['N°']]) === numeroProyecto) { filaEncontrada = i; break; }
  }
  if (filaEncontrada === -1) return { ok: false, error: 'Proyecto no encontrado: ' + body.numero };

  const filaActual = data[filaEncontrada];
  const rowNumber = filaEncontrada + 1;
  sheet.deleteRow(rowNumber);

  const historial = ss.getSheetByName(SHEET_HISTORIAL);
  historial.appendRow([
    new Date(), numeroProyecto, filaActual[colIndex['Proyecto']], sesion.usuario,
    filaActual[colIndex['Estado']], 'Eliminado', '', 'Proyecto eliminado.', cfg.moduloLabel
  ]);

  return { ok: true };
}

/** Devuelve hitos, evidencias e historial de un proyecto puntual, filtrados por módulo + N°. */
function listarSeguimiento_(usuario, clave, modulo, numero) {
  const sesion = autenticar_(usuario, clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };
  const proyectoInfo = buscarProyecto_(cfg, numero);
  if (!proyectoInfo) return { ok: false, error: 'Proyecto no encontrado: ' + numero };
  const proyectoActual = filaAObjeto_(proyectoInfo.fila, proyectoInfo.colIndex);
  if (!proyectoVisiblePara_(sesion, modulo, proyectoActual)) {
    return { ok: false, error: 'No tienes acceso a este proyecto.' };
  }

  const numeroNum = Number(numero);
  const ss = getSs_();

  function leerFiltrado_(nombreSheet) {
    const sheet = ss.getSheetByName(nombreSheet);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colModulo = headers.indexOf('Módulo');
    const colNumero = headers.indexOf('N° proyecto');
    const filas = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][colModulo] === cfg.moduloLabel && Number(data[i][colNumero]) === numeroNum) {
        const obj = {};
        headers.forEach(function (h, idx) { obj[h] = data[i][idx]; });
        filas.push(obj);
      }
    }
    return filas;
  }

  return {
    ok: true,
    hitos: leerFiltrado_(SHEET_HITOS),
    evidencias: leerFiltrado_(SHEET_EVIDENCIAS),
    historial: leerFiltrado_(SHEET_HISTORIAL),
  };
}

/** Crea un hito manual (título, fecha objetivo, periodo libre tipo "2027-Q1") para un proyecto. */
function crearHito_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = body.modulo;
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };

  const proyectoInfo = buscarProyecto_(cfg, body.numero);
  if (!proyectoInfo) return { ok: false, error: 'Proyecto no encontrado: ' + body.numero };
  const proyectoActual = filaAObjeto_(proyectoInfo.fila, proyectoInfo.colIndex);
  if (!puedeEditarProyecto_(sesion, modulo, proyectoActual)) {
    return { ok: false, error: 'No tienes permiso para editar el seguimiento de este proyecto.' };
  }

  const titulo = String(body.titulo || '').trim();
  if (!titulo) return { ok: false, error: 'El título del hito es obligatorio.' };
  if (!body.fechaObjetivo) return { ok: false, error: 'La fecha objetivo es obligatoria.' };

  const id = Utilities.getUuid();
  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_HITOS);
  const nombreProyecto = proyectoInfo.fila[proyectoInfo.colIndex['Proyecto']];
  const now = new Date();
  sheet.appendRow([
    id, cfg.moduloLabel, Number(body.numero), nombreProyecto,
    titulo, body.periodo || '', body.fechaObjetivo, 'Pendiente',
    body.responsable || '', body.comentario || '', now, sesion.usuario
  ]);
  return { ok: true, id: id };
}

/** Actualiza un hito existente (identificado por ID) — típicamente su Estado. */
function actualizarHito_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = body.modulo;
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };

  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_HITOS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = {};
  headers.forEach(function (h, idx) { colIndex[h] = idx; });

  let filaEncontrada = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex['ID']] === body.id) { filaEncontrada = i; break; }
  }
  if (filaEncontrada === -1) return { ok: false, error: 'Hito no encontrado.' };
  const filaActual = data[filaEncontrada];
  const numeroProyecto = filaActual[colIndex['N° proyecto']];

  const proyectoInfo = buscarProyecto_(cfg, numeroProyecto);
  const proyectoActual = proyectoInfo ? filaAObjeto_(proyectoInfo.fila, proyectoInfo.colIndex) : {};
  if (!puedeEditarProyecto_(sesion, modulo, proyectoActual)) {
    return { ok: false, error: 'No tienes permiso para editar este hito.' };
  }

  if (body.estado && HITO_ESTADOS_VALIDOS.indexOf(body.estado) === -1) {
    return { ok: false, error: 'Estado de hito inválido: ' + body.estado };
  }

  const rowNumber = filaEncontrada + 1;
  const camposEditables = {
    'Título': 'titulo', 'Periodo': 'periodo', 'Fecha objetivo': 'fechaObjetivo',
    'Estado': 'estado', 'Responsable': 'responsable', 'Comentario': 'comentario',
  };
  Object.keys(camposEditables).forEach(function (columna) {
    const campoBody = camposEditables[columna];
    if (body[campoBody] !== undefined && body[campoBody] !== null) {
      sheet.getRange(rowNumber, colIndex[columna] + 1).setValue(body[campoBody]);
    }
  });
  return { ok: true };
}

/** Elimina un hito (identificado por ID). */
function eliminarHito_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = body.modulo;
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };

  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_HITOS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = {};
  headers.forEach(function (h, idx) { colIndex[h] = idx; });

  let filaEncontrada = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex['ID']] === body.id) { filaEncontrada = i; break; }
  }
  if (filaEncontrada === -1) return { ok: false, error: 'Hito no encontrado.' };

  const numeroProyecto = data[filaEncontrada][colIndex['N° proyecto']];
  const proyectoInfo = buscarProyecto_(cfg, numeroProyecto);
  const proyectoActual = proyectoInfo ? filaAObjeto_(proyectoInfo.fila, proyectoInfo.colIndex) : {};
  if (!puedeEditarProyecto_(sesion, modulo, proyectoActual)) {
    return { ok: false, error: 'No tienes permiso para eliminar este hito.' };
  }

  sheet.deleteRow(filaEncontrada + 1);
  return { ok: true };
}

/** Agrega una evidencia (tipo del catálogo + link) a un proyecto. */
function agregarEvidencia_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = body.modulo;
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };

  const proyectoInfo = buscarProyecto_(cfg, body.numero);
  if (!proyectoInfo) return { ok: false, error: 'Proyecto no encontrado: ' + body.numero };
  const proyectoActual = filaAObjeto_(proyectoInfo.fila, proyectoInfo.colIndex);
  if (!puedeEditarProyecto_(sesion, modulo, proyectoActual)) {
    return { ok: false, error: 'No tienes permiso para agregar evidencias a este proyecto.' };
  }

  if (EVIDENCIA_TIPOS_VALIDOS.indexOf(body.tipo) === -1) {
    return { ok: false, error: 'Tipo de evidencia inválido: ' + body.tipo };
  }
  const link = String(body.link || '').trim();
  if (!/^https?:\/\//i.test(link)) {
    return { ok: false, error: 'El link de la evidencia debe ser una URL válida (http/https).' };
  }

  const id = Utilities.getUuid();
  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_EVIDENCIAS);
  const nombreProyecto = proyectoInfo.fila[proyectoInfo.colIndex['Proyecto']];
  sheet.appendRow([
    id, cfg.moduloLabel, Number(body.numero), nombreProyecto,
    body.tipo, body.descripcion || '', link, new Date(), sesion.usuario
  ]);
  return { ok: true, id: id };
}

/** Elimina una evidencia (identificada por ID). */
function eliminarEvidencia_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const modulo = body.modulo;
  const cfg = MODULOS[modulo];
  if (!cfg) return { ok: false, error: 'Módulo inválido: ' + modulo };

  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_EVIDENCIAS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = {};
  headers.forEach(function (h, idx) { colIndex[h] = idx; });

  let filaEncontrada = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex['ID']] === body.id) { filaEncontrada = i; break; }
  }
  if (filaEncontrada === -1) return { ok: false, error: 'Evidencia no encontrada.' };

  const numeroProyecto = data[filaEncontrada][colIndex['N° proyecto']];
  const proyectoInfo = buscarProyecto_(cfg, numeroProyecto);
  const proyectoActual = proyectoInfo ? filaAObjeto_(proyectoInfo.fila, proyectoInfo.colIndex) : {};
  if (!puedeEditarProyecto_(sesion, modulo, proyectoActual)) {
    return { ok: false, error: 'No tienes permiso para eliminar esta evidencia.' };
  }

  sheet.deleteRow(filaEncontrada + 1);
  return { ok: true };
}

/**
 * Restablece la hoja ProyectosTI a los 20 proyectos base de infraestructura tecnológica.
 * ADVERTENCIA: elimina TODAS las filas de datos existentes en ProyectosTI.
 * Usar solo en desarrollo / testing — no ejecutar en producción con datos reales.
 */
function resetearTI() {
  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_PROYECTOS_TI);
  if (!sheet) {
    Logger.log('[resetearTI] Hoja "' + SHEET_PROYECTOS_TI + '" no encontrada.');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  const now = new Date();
  PROYECTOS_BASE_TI.forEach(function (p) {
    sheet.appendRow([
      p[0], p[1], p[2], p[3], p[4],
      p[5], p[6],
      '', '', 'Formulación',
      '', p[7], p[8], p[9], now,
    ]);
  });
  aplicarValidaciones_(sheet);
  Logger.log('[resetearTI] Hoja reiniciada con ' + PROYECTOS_BASE_TI.length + ' proyectos base TI.');
}

/**
 * Restablece la hoja Proyectos (física) a los 19 proyectos base del Plan Maestro.
 * ADVERTENCIA: elimina TODAS las filas de datos existentes en Proyectos.
 * Usar solo en desarrollo / testing — no ejecutar en producción con datos reales.
 */
function resetearFisica() {
  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_PROYECTOS);
  if (!sheet) {
    Logger.log('[resetearFisica] Hoja "' + SHEET_PROYECTOS + '" no encontrada.');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  const now = new Date();
  PROYECTOS_BASE.forEach(function (p) {
    sheet.appendRow([
      p[0], p[1], p[2], p[3], p[4],
      '', '',
      '', '', '',
      '', p[5], '', p[6], now,
    ]);
  });
  aplicarValidaciones_(sheet);
  Logger.log('[resetearFisica] Hoja reiniciada con ' + PROYECTOS_BASE.length + ' proyectos base físicos.');
}
