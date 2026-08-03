/**
 * Backend de "Plan de Infraestructura UDES".
 * Pegar este código en el Google Apps Script vinculado al Google Sheet del proyecto,
 * ejecutar setup() una vez, y desplegar como Web App (ver instrucciones adjuntas).
 */

// ID del Google Sheet (se toma del final de su URL: .../d/ESTE_ID/edit).
// Necesario porque SpreadsheetApp.getActiveSpreadsheet() devuelve null cuando
// el script corre como Web App (sin una hoja abierta por un usuario).
const SPREADSHEET_ID = '1S6rO2pOxsSrgCDjGxdAe2CoLWZjQ_zcfufCU1TL1Y3w';

const SHEET_PROYECTOS = 'Proyectos';
const SHEET_USUARIOS = 'Usuarios';
const SHEET_HISTORIAL = 'Historial_Seguimiento';

function getSs_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

const ESTADOS_VALIDOS = ['Sin iniciar', 'En desarrollo', 'Ejecutado', 'Atrasado'];
const PRIORIDADES_VALIDAS = ['Alta', 'Media', 'Baja'];

const PROGRAMAS_VALIDOS = [
  'Movilidad y accesibilidad',
  'Expansión académica',
  'Equipamiento cultural y deportivo',
  'Bienestar y vida universitaria',
  'Identidad e imagen institucional',
  'Ciencia, tecnología e innovación',
  'Reserva de suelo',
];

const CAMPUS_VALIDOS = ['Bucaramanga', 'Cúcuta', 'Valledupar'];

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

const USUARIOS_BASE = [
  ['bucaramanga', 'CAMBIAR-CLAVE-1', 'Bucaramanga', 'responsable'],
  ['cucuta', 'CAMBIAR-CLAVE-2', 'Cúcuta', 'responsable'],
  ['valledupar', 'CAMBIAR-CLAVE-3', 'Valledupar', 'responsable'],
  ['admin', 'CAMBIAR-CLAVE-ADMIN', 'Todos', 'admin'],
];

const PROYECTOS_HEADERS = [
  'N°', 'Proyecto', 'Campus', 'Descripción', 'Programa',
  'Costo estimado', 'Fuente de financiación',
  'Fecha inicio', 'Fecha fin', 'Fase actual',
  'Responsable', 'Estado', 'Prioridad', '% avance', 'Última actualización'
];

const USUARIOS_HEADERS = ['Usuario', 'Clave', 'Campus', 'Rol'];

const HISTORIAL_HEADERS = ['Fecha', 'N° proyecto', 'Proyecto', 'Usuario', 'Estado anterior', 'Estado nuevo', '% avance', 'Comentario'];

/**
 * Ejecutar UNA sola vez desde el editor de Apps Script (menú Ejecutar > setup)
 * para crear las hojas, encabezados, validaciones y datos base.
 * Es seguro volver a ejecutarla: no duplica filas si "Proyectos" y "Usuarios" ya tienen datos.
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
  }

  const usuariosSheet = getOrCreateSheet_(ss, SHEET_USUARIOS);
  if (usuariosSheet.getLastRow() === 0) {
    usuariosSheet.appendRow(USUARIOS_HEADERS);
    USUARIOS_BASE.forEach(function (u) { usuariosSheet.appendRow(u); });
    usuariosSheet.setFrozenRows(1);
  }

  const historialSheet = getOrCreateSheet_(ss, SHEET_HISTORIAL);
  if (historialSheet.getLastRow() === 0) {
    historialSheet.appendRow(HISTORIAL_HEADERS);
    historialSheet.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  Logger.log('Setup completo. Recuerda cambiar las claves en la hoja "Usuarios".');
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function aplicarValidaciones_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 200);
  const estadoRange = sheet.getRange(2, 12, lastRow - 1, 1); // columna "Estado"
  const estadoRule = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_VALIDOS, true).setAllowInvalid(false).build();
  estadoRange.setDataValidation(estadoRule);

  const prioridadRange = sheet.getRange(2, 13, lastRow - 1, 1); // columna "Prioridad"
  const prioridadRule = SpreadsheetApp.newDataValidation().requireValueInList(PRIORIDADES_VALIDAS, true).setAllowInvalid(false).build();
  prioridadRange.setDataValidation(prioridadRule);
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
 * Acciones soportadas: login, listarProyectos, actualizarProyecto, crearProyecto
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'login') return jsonResponse_(login_(body.usuario, body.clave));
    if (action === 'listarProyectos') return jsonResponse_(listarProyectos_(body.usuario, body.clave));
    if (action === 'actualizarProyecto') return jsonResponse_(actualizarProyecto_(body));
    if (action === 'crearProyecto') return jsonResponse_(crearProyecto_(body));

    return jsonResponse_({ ok: false, error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Busca al usuario en la hoja "Usuarios" y valida la clave. */
function autenticar_(usuario, clave) {
  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_USUARIOS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).toLowerCase() === String(usuario).toLowerCase() && String(row[1]) === String(clave)) {
      return { usuario: row[0], campus: row[2], rol: row[3] };
    }
  }
  return null;
}

function login_(usuario, clave) {
  const sesion = autenticar_(usuario, clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };
  return { ok: true, sesion: sesion };
}

/** Devuelve los proyectos visibles para el usuario: todos si es admin, solo los de su campus si es responsable. */
function listarProyectos_(usuario, clave) {
  const sesion = autenticar_(usuario, clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_PROYECTOS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const proyectos = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const campus = row[2];
    if (sesion.rol === 'admin' || campus === sesion.campus) {
      const proyecto = {};
      headers.forEach(function (h, idx) { proyecto[h] = row[idx]; });
      proyectos.push(proyecto);
    }
  }

  return { ok: true, sesion: sesion, proyectos: proyectos };
}

/**
 * Actualiza los campos editables de un proyecto (identificado por N°) y registra el cambio
 * en Historial_Seguimiento. Un responsable solo puede actualizar proyectos de su propio campus.
 */
function actualizarProyecto_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_PROYECTOS);
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
  const campusProyecto = filaActual[colIndex['Campus']];
  if (sesion.rol !== 'admin' && campusProyecto !== sesion.campus) {
    return { ok: false, error: 'No tienes permiso para editar proyectos de otro campus.' };
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
    if (body[campoBody] !== undefined && body[campoBody] !== null) {
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
    body.comentario || ''
  ]);

  return { ok: true };
}

/**
 * Crea un proyecto nuevo (fuera de los 19 del Plan Maestro original), dentro de una
 * de las categorías funcionales ya definidas. Un responsable solo puede crear proyectos
 * en su propio campus; el admin puede elegir cualquiera de los 3 campus.
 */
function crearProyecto_(body) {
  const sesion = autenticar_(body.usuario, body.clave);
  if (!sesion) return { ok: false, error: 'Usuario o clave incorrectos.' };

  const nombreProyecto = String(body.proyecto || '').trim();
  if (!nombreProyecto) return { ok: false, error: 'El nombre del proyecto es obligatorio.' };

  if (PROGRAMAS_VALIDOS.indexOf(body.programa) === -1) {
    return { ok: false, error: 'Programa inválido: ' + body.programa };
  }

  let campus;
  if (sesion.rol === 'admin') {
    if (CAMPUS_VALIDOS.indexOf(body.campus) === -1) {
      return { ok: false, error: 'Campus inválido: ' + body.campus };
    }
    campus = body.campus;
  } else {
    campus = sesion.campus;
  }

  const ss = getSs_();
  const sheet = ss.getSheetByName(SHEET_PROYECTOS);
  const data = sheet.getDataRange().getValues();
  const colIndex = {};
  data[0].forEach(function (h, idx) { colIndex[h] = idx; });

  let maxNumero = 0;
  for (let i = 1; i < data.length; i++) {
    const n = Number(data[i][colIndex['N°']]);
    if (n > maxNumero) maxNumero = n;
  }
  const nuevoNumero = maxNumero + 1;
  const now = new Date();

  sheet.appendRow([
    nuevoNumero, nombreProyecto, campus, String(body.descripcion || ''), body.programa,
    '', '',
    '', '', '',
    '', 'Sin iniciar', '', 0, now
  ]);

  const historial = ss.getSheetByName(SHEET_HISTORIAL);
  historial.appendRow([
    now, nuevoNumero, nombreProyecto, sesion.usuario,
    '—', 'Sin iniciar', 0, 'Proyecto creado (fuera del Plan Maestro original).'
  ]);

  return { ok: true, numero: nuevoNumero };
}
