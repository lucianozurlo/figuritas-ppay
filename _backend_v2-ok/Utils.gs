// ============================================================
// Utils.gs — Funciones utilitarias compartidas
// ============================================================

/**
 * Obtiene la hoja de cálculo por nombre definido en CONFIG.
 * Lanza un error descriptivo si no existe.
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error(
      `No existe la hoja "${CONFIG.SHEET_NAME}". ` +
        `Verificar CONFIG.SHEET_NAME o que el script esté vinculado al Sheet correcto.`,
    );
  }
  return sheet;
}

/**
 * Construye un mapa { nombreColumna: índiceColumna (1-based) }
 * a partir de la fila de encabezados del Sheet.
 * Los nombres se normalizan con trim() para evitar errores por espacios.
 */
function getColumnMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    const key = String(header).trim();
    if (key) map[key] = index + 1;
  });
  return map;
}

/**
 * Obtiene el valor de una celda de una fila por nombre de columna.
 * Lanza error descriptivo si la columna no existe en el mapa.
 */
function getCell_(row, colMap, columnName) {
  const colIndex = colMap[columnName];
  if (!colIndex) {
    throw new Error(
      `Columna no encontrada en el Sheet: "${columnName}". ` +
        `Verificar CONFIG.COLUMNS y los encabezados del Sheet.`,
    );
  }
  return row[colIndex - 1];
}

/**
 * Escribe un valor en una celda específica por nombre de columna.
 * Lanza error descriptivo si la columna no existe.
 */
function setCell_(sheet, rowIndex, colMap, columnName, value) {
  const colIndex = colMap[columnName];
  if (!colIndex) {
    throw new Error(
      `Columna no encontrada en el Sheet: "${columnName}". ` +
        `Verificar CONFIG.COLUMNS y los encabezados del Sheet.`,
    );
  }
  sheet.getRange(rowIndex, colIndex).setValue(value);
}

/**
 * Extrae el ID de archivo de Drive desde el valor almacenado en el Sheet.
 * El WebApp.gs ahora guarda directamente el ID puro, pero se mantiene
 * el parser de URLs por compatibilidad o edición manual.
 */
function extractDriveFileId_(value) {
  const text = String(value || "").trim();
  if (!text) throw new Error("La celda de foto está vacía.");

  // ID puro (25+ caracteres alfanuméricos con guiones y underscores)
  if (/^[a-zA-Z0-9_-]{25,}$/.test(text)) return text;

  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]{25,})/, // .../file/d/ID/view
    /[?&]id=([a-zA-Z0-9_-]{25,})/, // ...?id=ID o ...&id=ID
    /\/uc\?export=view&id=([a-zA-Z0-9_-]{25,})/, // URL directa
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1];
  }

  throw new Error(`No se pudo extraer el ID de Drive desde: "${text}"`);
}

/**
 * Reintenta una función con backoff exponencial.
 * @param {Function} fn        - Función a ejecutar. Debe ser síncrona.
 * @param {number}   maxRetries - Reintentos máximos (default: 3).
 */
function retryWithExponentialBackoff_(fn, maxRetries = 3) {
  let lastError;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return fn();
    } catch (err) {
      lastError = err;
      if (attempt < maxRetries) {
        Utilities.sleep(Math.pow(2, attempt) * 1000); // 1s, 2s, 4s
      }
    }
  }
  throw lastError;
}
