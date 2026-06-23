// ============================================================
// WebApp.gs — Endpoint público del formulario de figuritas
//
// Dos funciones de entrada:
//   doGet()  → no se usa si el HTML está en GitHub Pages / Netlify.
//              Se mantiene por si se quiere hostear desde Apps Script.
//   doPost() → recibe el JSON del formulario, guarda la foto en Drive
//              y escribe la fila en el Sheet.
// ============================================================

/**
 * doPost — receptor principal del formulario.
 *
 * Recibe un body JSON con:
 *   nombre, mail, area, superpoder, actitud,
 *   fotoBase64 (JPEG en base64, sin prefijo data:),
 *   consentimientoMural (boolean)
 *
 * Responde siempre con JSON:
 *   { status: 'ok' }           en caso de éxito
 *   { status: 'error', message: '...' }  en caso de error
 */
function doPost(e) {
  // Cabeceras CORS permisivas — necesario para fetch desde dominio externo
  const headers = {
    'Access-Control-Allow-Origin':  '*',
    'Access-Control-Allow-Methods': 'POST',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type':                 'application/json'
  };

  try {
    // Parsear payload
    const payload = JSON.parse(e.postData.contents);

    // Validaciones mínimas en el servidor (defensa ante bypass del cliente)
    const errores = validarPayload_(payload);
    if (errores.length > 0) {
      return jsonResponse_({ status: 'error', message: errores.join(' | ') }, headers);
    }

    // Deduplicación: si ya existe un registro con este mail en estado
    // EMAIL_ENVIADO o FIGURITA_CREADA, rechazar silenciosamente.
    if (esDuplicado_(payload.mail)) {
      // Responder ok para no confundir al usuario — ya fue procesado
      return jsonResponse_({ status: 'ok' }, headers);
    }

    // Guardar foto en Drive y obtener ID y URL directa
    const { fileId, directUrl } = guardarFotoEnDrive_(payload.fotoBase64, payload.nombre);

    // Escribir fila en el Sheet
    escribirFila_({
      nombre:              payload.nombre,
      mail:                payload.mail,
      area:                payload.area,
      superpoder:          payload.superpoder,
      actitud:             payload.actitud,
      consentimientoMural: payload.consentimientoMural ? 'Sí' : '',
      idArchivoDrive:      fileId,
      urlArchivoDrive:     directUrl,
      estado:              'PENDIENTE',
      timestampEnvio:      new Date().toISOString(),
    });

    return jsonResponse_({ status: 'ok' }, headers);

  } catch (err) {
    Logger.log('[doPost] Error: ' + err.toString());
    return jsonResponse_({ status: 'error', message: 'Error interno del servidor. Intentá de nuevo.' }, headers);
  }
}

/**
 * doGet — responde al preflight OPTIONS y sirve el HTML
 * si se quiere hostear el form desde Apps Script en lugar de GitHub Pages.
 * Si el frontend vive en GitHub Pages, este doGet no se usa.
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index') // requiere index.html en el proyecto
    .setTitle('Mi Figurita Personal Pay')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// VALIDACIÓN
// ============================================================

function validarPayload_(payload) {
  const errores = [];

  if (!payload.nombre || String(payload.nombre).trim().length < 2)
    errores.push('Nombre inválido.');

  if (!payload.mail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(payload.mail).trim()))
    errores.push('Mail inválido.');

  if (!payload.area || String(payload.area).trim() === '')
    errores.push('Área inválida.');

  if (!payload.superpoder || String(payload.superpoder).trim().length < 3)
    errores.push('Superpoder inválido.');

  if (!payload.actitud || String(payload.actitud).trim().length < 3)
    errores.push('Actitud inválida.');

  if (!payload.fotoBase64 || payload.fotoBase64.length < 1000)
    errores.push('Foto inválida o ausente.');

  return errores;
}


// ============================================================
// DEDUPLICACIÓN
// ============================================================

/**
 * Verifica si ya existe una fila con este mail que haya llegado
 * al menos a FIGURITA_CREADA o EMAIL_ENVIADO.
 * Evita procesar el mismo participante dos veces si envía el form más de una vez.
 */
function esDuplicado_(mail) {
  const sheet  = getSheet_();
  const colMap = getColumnMap_(sheet);
  const data   = sheet.getDataRange().getValues();

  const mailNorm  = String(mail).trim().toLowerCase();
  const colMail   = colMap[CONFIG.COLUMNS.mail];
  const colEstado = colMap[CONFIG.COLUMNS.estado];

  if (!colMail || !colEstado) return false;

  return data.slice(1).some(row => {
    const rowMail   = String(row[colMail - 1]   || '').trim().toLowerCase();
    const rowEstado = String(row[colEstado - 1]  || '').trim();
    return rowMail === mailNorm &&
      ['FIGURITA_CREADA', 'EMAIL_ENVIADO', 'PROCESANDO', 'PENDIENTE'].includes(rowEstado);
  });
}


// ============================================================
// GUARDAR FOTO EN DRIVE
// ============================================================

/**
 * Convierte el base64 JPEG recibido en un Blob y lo guarda en la
 * carpeta de fotos originales en Drive.
 *
 * Devuelve { fileId, directUrl } donde directUrl es del formato
 * https://drive.google.com/uc?export=view&id=FILE_ID
 * que es compatible con =IMAGE() en Sheets y con el script de generación.
 */
function guardarFotoEnDrive_(base64, nombre) {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_FOTOS_ID);

  const nombreSanitizado = String(nombre)
    .replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚüÜñÑ\s]/g, '')
    .trim()
    .substring(0, 30);

  const timestamp  = new Date().getTime();
  const fileName   = `Foto_${nombreSanitizado}_${timestamp}.jpg`;

  const bytes = Utilities.base64Decode(base64);
  const blob  = Utilities.newBlob(bytes, 'image/jpeg', fileName);

  const file = folder.createFile(blob);

  // El archivo de foto no necesita ser público — solo el script lo lee.
  // Sin embargo, debe ser accesible por el script, lo cual está garantizado
  // porque el script corre como el propietario de la carpeta.

  return {
    fileId:    file.getId(),
    directUrl: `https://drive.google.com/uc?export=view&id=${file.getId()}`
  };
}


// ============================================================
// ESCRIBIR FILA EN SHEET
// ============================================================

/**
 * Agrega una nueva fila al Sheet con todos los datos del participante
 * y los campos operativos iniciales.
 */
function escribirFila_(datos) {
  const sheet  = getSheet_();
  const colMap = getColumnMap_(sheet);

  // Calcular el índice de la próxima fila vacía
  const lastRow  = sheet.getLastRow();
  const newRow   = lastRow + 1;

  // Construir el array de valores para toda la fila
  // Primero llenamos con vacíos, luego asignamos por columna
  const totalCols = sheet.getLastColumn();
  const rowValues = new Array(totalCols).fill('');

  const set = (colName, value) => {
    const idx = colMap[colName];
    if (idx) rowValues[idx - 1] = value;
  };

  set(CONFIG.COLUMNS.nombre,              datos.nombre);
  set(CONFIG.COLUMNS.mail,                datos.mail);
  set(CONFIG.COLUMNS.area,                datos.area);
  set(CONFIG.COLUMNS.superpoder,          datos.superpoder);
  set(CONFIG.COLUMNS.actitud,             datos.actitud);
  set(CONFIG.COLUMNS.consentimientoMural, datos.consentimientoMural);
  set(CONFIG.COLUMNS.idArchivoDrive,      datos.idArchivoDrive);
  set(CONFIG.COLUMNS.urlArchivoDrive,     datos.urlArchivoDrive);
  set(CONFIG.COLUMNS.estado,              datos.estado);
  set(CONFIG.COLUMNS.timestampProcesando, '');
  set(CONFIG.COLUMNS.detalleError,        '');

  // Escribir toda la fila de una sola vez (1 llamada API en lugar de N)
  sheet.getRange(newRow, 1, 1, totalCols).setValues([rowValues]);

  Logger.log(`[escribirFila_] Nueva fila en ${newRow}: ${datos.nombre} <${datos.mail}>`);
}


// ============================================================
// HELPER: respuesta JSON
// ============================================================

function jsonResponse_(data, headers) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

  // Apps Script no permite setear headers arbitrarios en ContentService,
  // pero el CORS se maneja automáticamente cuando el Web App está publicado
  // como "Anyone can access" (acceso anónimo).
  // Los headers de la firma se mantienen por documentación.
  return output;
}
