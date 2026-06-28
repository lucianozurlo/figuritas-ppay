// ============================================================
// Processing.gs — Generación por lotes, envío y operación
// ============================================================


// ──────────────────────────────────────────────────────────────
// MENÚ DE ADMINISTRACIÓN
// ──────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎴 Figuritas Admin')
    .addItem('1. Procesar lote de figuritas',   'procesarLoteFiguritas')
    .addItem('2. Enviar lote de correos',        'enviarLoteCorreos')
    .addSeparator()
    .addItem('Reprocesar errores',              'reprocesarErrores')
    .addItem('Ver resumen de estados',          'mostrarResumenEstados')
    .addSeparator()
    .addItem('Validar configuración',           'validarConfiguracion')
    .addItem('Crear triggers de tiempo',        'createTimeTriggers')
    .addItem('Eliminar todos los triggers',     'deleteAllTriggers')
    .addSeparator()
    .addItem('🖼️ Generar mural',              'generarMural')
    .addItem('📦 Exportar mural',             'exportarMural')
    .addToUi();
}


// ──────────────────────────────────────────────────────────────
// TRIGGERS
// ──────────────────────────────────────────────────────────────

/**
 * Crea los triggers automáticos de generación y envío.
 * Verifica que no existan antes de crearlos para evitar duplicados.
 */
function createTimeTriggers() {
  const ui       = SpreadsheetApp.getUi();
  const existing = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const creados  = [];

  if (!existing.includes('procesarLoteFiguritas')) {
    ScriptApp.newTrigger('procesarLoteFiguritas').timeBased().everyMinutes(10).create();
    creados.push('procesarLoteFiguritas (cada 10 min)');
  }

  if (!existing.includes('enviarLoteCorreos')) {
    ScriptApp.newTrigger('enviarLoteCorreos').timeBased().everyMinutes(10).create();
    creados.push('enviarLoteCorreos (cada 10 min)');
  }

  if (creados.length > 0) {
    ui.alert('✅ Triggers creados', creados.join('\n'), ui.ButtonSet.OK);
  } else {
    ui.alert('ℹ️ Ya existían', 'Los triggers ya estaban creados. No se duplicaron.', ui.ButtonSet.OK);
  }
}

function deleteAllTriggers() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '⚠️ Confirmar',
    '¿Eliminar todos los triggers? Vas a tener que volver a crearlos.',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ui.alert('✅ Listo', 'Todos los triggers fueron eliminados.', ui.ButtonSet.OK);
}


// ──────────────────────────────────────────────────────────────
// VALIDACIÓN DE CONFIGURACIÓN
// ──────────────────────────────────────────────────────────────

function validarConfiguracion() {
  const ui       = SpreadsheetApp.getUi();
  const errores  = [];
  const avisos   = [];

  // Verificar hoja
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    errores.push(`❌ No existe la hoja "${CONFIG.SHEET_NAME}".`);
  } else {
    // Verificar columnas
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim());

    Object.entries(CONFIG.COLUMNS).forEach(([key, colName]) => {
      if (!headers.includes(colName)) {
        errores.push(`❌ Columna faltante: "${colName}" (CONFIG.COLUMNS.${key})`);
      }
    });
  }

  // Verificar carpeta de figuritas
  try {
    DriveApp.getFolderById(CONFIG.FOLDER_FIGURITAS_ID);
  } catch (_) {
    errores.push(`❌ No se puede acceder a FOLDER_FIGURITAS_ID: "${CONFIG.FOLDER_FIGURITAS_ID}"`);
  }

  // Verificar carpeta de fotos
  try {
    DriveApp.getFolderById(CONFIG.FOLDER_FOTOS_ID);
  } catch (_) {
    errores.push(`❌ No se puede acceder a FOLDER_FOTOS_ID: "${CONFIG.FOLDER_FOTOS_ID}"`);
  }

  // Verificar plantilla de Slides
  try {
    DriveApp.getFileById(CONFIG.SLIDE_TEMPLATE_ID);
  } catch (_) {
    errores.push(`❌ No se puede acceder a SLIDE_TEMPLATE_ID: "${CONFIG.SLIDE_TEMPLATE_ID}"`);
  }

  // Verificar cuota de mail
  const quota = MailApp.getRemainingDailyQuota();
  if (quota < 100) {
    avisos.push(`⚠️ Cuota de mail baja: ${quota} envíos restantes hoy.`);
  }

  // Verificar texto alternativo en la plantilla
  if (CONFIG.ALT_TEXT_FOTO !== 'FOTO_PERFIL_REEMPLAZAR') {
    avisos.push(`ℹ️ ALT_TEXT_FOTO personalizado: "${CONFIG.ALT_TEXT_FOTO}". Verificar que coincide con la plantilla.`);
  }

  // Resultado
  if (errores.length === 0 && avisos.length === 0) {
    ui.alert('✅ Todo en orden', 'Configuración válida. El sistema está listo para operar.', ui.ButtonSet.OK);
  } else {
    const partes = [];
    if (errores.length) partes.push('ERRORES CRÍTICOS (resolver antes de operar):\n' + errores.join('\n'));
    if (avisos.length)  partes.push('AVISOS:\n' + avisos.join('\n'));
    ui.alert('⚠️ Problemas de configuración', partes.join('\n\n'), ui.ButtonSet.OK);
  }
}


// ──────────────────────────────────────────────────────────────
// RECUPERACIÓN DE ZOMBIES
// ──────────────────────────────────────────────────────────────

/**
 * Detecta filas que llevan más de PROCESANDO_TIMEOUT_MINUTOS en estado
 * PROCESANDO (script cortado por timeout) y las devuelve a PENDIENTE.
 * Se llama automáticamente al inicio de procesarLoteFiguritas.
 */
function resetZombies_(sheet, colMap, data) {
  const ahora    = new Date();
  const limitMs  = CONFIG.PROCESANDO_TIMEOUT_MINUTOS * 60 * 1000;
  let   reseteados = 0;

  data.slice(1).forEach((row, i) => {
    const rowIndex = i + 2;
    const estado   = String(getCell_(row, colMap, CONFIG.COLUMNS.estado) || '').trim();
    if (estado !== 'PROCESANDO') return;

    const tsVal = getCell_(row, colMap, CONFIG.COLUMNS.timestampProcesando);
    const ts    = tsVal ? new Date(tsVal) : null;

    if (!ts || isNaN(ts.getTime()) || (ahora - ts) > limitMs) {
      setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado, 'PENDIENTE');
      setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.detalleError,
        `Recuperado de PROCESANDO zombie (${new Date().toISOString()})`);
      reseteados++;
    }
  });

  if (reseteados > 0) Logger.log(`[resetZombies_] ${reseteados} zombie(s) reseteados.`);
}


// ──────────────────────────────────────────────────────────────
// GENERACIÓN DE FIGURITAS POR LOTES
// ──────────────────────────────────────────────────────────────

function procesarLoteFiguritas() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('[procesarLoteFiguritas] No se obtuvo el lock. Otra ejecución en curso.');
    return;
  }

  try {
    const sheet  = getSheet_();
    const colMap = getColumnMap_(sheet);

    // Resetear zombies antes de procesar
    const dataParaZombies = sheet.getDataRange().getValues();
    resetZombies_(sheet, colMap, dataParaZombies);

    // Releer datos frescos después del reset
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    let procesadas = 0;

    for (let i = 0; i < rows.length; i++) {
      if (procesadas >= CONFIG.BATCH_SIZE_GENERACION) break;

      const row      = rows[i];
      const rowIndex = i + 2;
      const estado   = String(getCell_(row, colMap, CONFIG.COLUMNS.estado) || '').trim();

      if (estado && estado !== 'PENDIENTE') continue;

      let presentationCopy = null;

      try {
        // Marcar como PROCESANDO con timestamp
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado,              'PROCESANDO');
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.timestampProcesando, new Date().toISOString());
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.detalleError,        '');

        // Leer datos del participante
        const nombre     = String(getCell_(row, colMap, CONFIG.COLUMNS.nombre)     || '').trim();
        const area       = String(getCell_(row, colMap, CONFIG.COLUMNS.area)       || '').trim();
        const superpoder = String(getCell_(row, colMap, CONFIG.COLUMNS.superpoder) || '').trim();
        const actitud    = String(getCell_(row, colMap, CONFIG.COLUMNS.actitud)    || '').trim();

        // Obtener la foto original desde Drive
        // El WebApp.gs guarda el ID directo en id_archivo_drive
        const fotoId   = String(getCell_(row, colMap, CONFIG.COLUMNS.idArchivoDrive) || '').trim();
        if (!fotoId) throw new Error('No hay ID de foto para este registro. El formulario puede no haber subido la imagen.');

        const fotoBlob = retryWithExponentialBackoff_(() => DriveApp.getFileById(fotoId).getBlob());

        // Validar que el blob de foto tiene contenido real
        if (!fotoBlob || fotoBlob.getBytes().length < 500) {
          throw new Error(`Blob de foto inválido o vacío. ID: ${fotoId}`);
        }

        // Copiar la plantilla de Slides
        const nombreSlug = nombre.replace(/[^a-zA-Z0-9áéíóúüñ\s]/gi, '').trim().substring(0, 30);
        presentationCopy = DriveApp
          .getFileById(CONFIG.SLIDE_TEMPLATE_ID)
          .makeCopy(`Temp_${nombreSlug}_${rowIndex}`);

        const presentation = SlidesApp.openById(presentationCopy.getId());
        const slide        = presentation.getSlides()[0];

        // Reemplazar marcadores de texto
        slide.replaceAllText('{{nombre}}',     nombre);
        slide.replaceAllText('{{area}}',       area);
        slide.replaceAllText('{{superpoder}}', superpoder);
        slide.replaceAllText('{{actitud}}',    actitud);

        // Buscar la forma por texto alternativo y reemplazar con la foto
        const imageShape = slide.getShapes()
          .find(s => s.getAltTextTitle() === CONFIG.ALT_TEXT_FOTO);

        if (!imageShape) {
          throw new Error(
            `Forma con alt text "${CONFIG.ALT_TEXT_FOTO}" no encontrada en la plantilla. ` +
            `Verificar el diseño de la diapositiva.`
          );
        }

        imageShape.replaceWithImage(fotoBlob);
        presentation.saveAndClose();

        // Exportar diapositiva como PNG vía URL autenticada
        const slideId        = slide.getObjectId();
        const presentationId = presentationCopy.getId();
        const exportUrl      = `https://docs.google.com/presentation/d/${presentationId}/export/png?pageid=${slideId}&scale=2`;

        const pngBlob = retryWithExponentialBackoff_(() => {
          const response = UrlFetchApp.fetch(exportUrl, {
            headers:           { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
            muteHttpExceptions: true
          });

          const code = response.getResponseCode();
          if (code !== 200) {
            throw new Error(`Exportación PNG falló con HTTP ${code}.`);
          }

          const blob = response.getBlob();
          if (!blob || blob.getBytes().length < 1000) {
            throw new Error(`PNG exportado está vacío (${blob ? blob.getBytes().length : 0} bytes).`);
          }

          return blob.setName(`Figurita_${nombreSlug}_${rowIndex}.png`);
        });

        // Guardar el PNG en la carpeta de figuritas generadas
        const folder  = DriveApp.getFolderById(CONFIG.FOLDER_FIGURITAS_ID);
        const pngFile = folder.createFile(pngBlob);

        // URL directa compatible con =IMAGE() en Sheets
        const urlFigura = `https://drive.google.com/uc?export=view&id=${pngFile.getId()}`;

        // Guardar resultados en el Sheet
        const idFigurita = `FIG-${rowIndex}-${new Date().getTime()}`;
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.idFigurita,         idFigurita);
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.idFiguraGenerada,   pngFile.getId());
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.urlFiguraGenerada,  urlFigura);
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado,             'FIGURITA_CREADA');

        procesadas++;

      } catch (err) {
        Logger.log(`[procesarLoteFiguritas] Error fila ${rowIndex}: ${err.toString()}`);
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado,       'ERROR');
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.detalleError, err.toString());
        procesadas++;

      } finally {
        // Siempre limpiar la copia temporal
        if (presentationCopy) {
          try {
            DriveApp.getFileById(presentationCopy.getId()).setTrashed(true);
          } catch (_) { /* ignorar error de limpieza */ }
        }
      }
    }

    Logger.log(`[procesarLoteFiguritas] Lote completado. Procesadas: ${procesadas}.`);

  } finally {
    lock.releaseLock();
  }
}


// ──────────────────────────────────────────────────────────────
// ENVÍO DE CORREOS POR LOTES
// ──────────────────────────────────────────────────────────────

function enviarLoteCorreos() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('[enviarLoteCorreos] No se obtuvo el lock.');
    return;
  }

  try {
    const sheet  = getSheet_();
    const colMap = getColumnMap_(sheet);
    const data   = sheet.getDataRange().getValues();
    const rows   = data.slice(1);
    let enviados = 0;

    for (let i = 0; i < rows.length; i++) {
      if (enviados >= CONFIG.BATCH_SIZE_EMAIL) break;

      // Verificar cuota con buffer de seguridad
      if (MailApp.getRemainingDailyQuota() <= CONFIG.QUOTA_BUFFER_EMAIL) {
        Logger.log('[enviarLoteCorreos] Cuota de mail insuficiente. Deteniendo lote.');
        break;
      }

      const row      = rows[i];
      const rowIndex = i + 2;
      const estado   = String(getCell_(row, colMap, CONFIG.COLUMNS.estado) || '').trim();

      if (estado !== 'FIGURITA_CREADA') continue;

      try {
        const nombre  = String(getCell_(row, colMap, CONFIG.COLUMNS.nombre)           || '').trim();
        const mail    = String(getCell_(row, colMap, CONFIG.COLUMNS.mail)             || '').trim();
        const fileId  = String(getCell_(row, colMap, CONFIG.COLUMNS.idFiguraGenerada) || '').trim();
        const fileUrl = String(getCell_(row, colMap, CONFIG.COLUMNS.urlFiguraGenerada)|| '').trim();

        if (!mail)   throw new Error('Campo mail vacío.');
        if (!fileId) throw new Error('No existe id de figura generada para enviar.');
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(mail)) {
          throw new Error(`Formato de email inválido: "${mail}"`);
        }

        const pngFile = DriveApp.getFileById(fileId);

        retryWithExponentialBackoff_(() => {
          MailApp.sendEmail({
            to:          mail,
            subject:     '🎴 Tu figurita de Personal Pay',
            htmlBody:    buildEmailHtml_(nombre, fileUrl),
            attachments: [pngFile.getBlob()]
          });
        });

        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado,       'EMAIL_ENVIADO');
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.detalleError, '');
        enviados++;

      } catch (err) {
        Logger.log(`[enviarLoteCorreos] Error fila ${rowIndex}: ${err.toString()}`);
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado,       'ERROR_EMAIL');
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.detalleError, err.toString());
        enviados++;
      }
    }

    Logger.log(`[enviarLoteCorreos] Lote completado. Enviados: ${enviados}. Cuota restante: ${MailApp.getRemainingDailyQuota()}.`);

  } finally {
    lock.releaseLock();
  }
}

function buildEmailHtml_(nombre, fileUrl) {
  return `
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
                max-width:560px;margin:0 auto;color:#1A1A2E;">
      <h2 style="color:#0062FF;">¡Hola, ${nombre}! 🎉</h2>
      <p style="font-size:15px;line-height:1.6;">
        Gracias por participar en el All Hands de Personal Pay.<br>
        Tu figurita personalizada está adjunta a este correo.
      </p>
      <p style="margin-top:20px;">
        <a href="${fileUrl}"
           style="display:inline-block;background:#0062FF;color:#fff;
                  padding:12px 24px;border-radius:8px;text-decoration:none;
                  font-weight:700;font-size:14px;">
          Ver figurita en Drive →
        </a>
      </p>
      <p style="margin-top:28px;font-size:12px;color:#6B7280;">
        Si no participaste en el evento o recibiste este correo por error,
        podés ignorarlo.
      </p>
    </div>
  `;
}


// ──────────────────────────────────────────────────────────────
// REPROCESAMIENTO DE ERRORES
// ──────────────────────────────────────────────────────────────

function reprocesarErrores() {
  const ui   = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    ui.alert('⚠️ En uso', 'El sistema está procesando datos. Intentá en un momento.', ui.ButtonSet.OK);
    return;
  }

  try {
    const sheet  = getSheet_();
    const colMap = getColumnMap_(sheet);
    const data   = sheet.getDataRange().getValues();
    const rows   = data.slice(1);

    let resetGen   = 0;
    let resetEmail = 0;

    rows.forEach((row, i) => {
      const rowIndex = i + 2;
      const estado   = String(getCell_(row, colMap, CONFIG.COLUMNS.estado)           || '').trim();
      const figId    = String(getCell_(row, colMap, CONFIG.COLUMNS.idFiguraGenerada)  || '').trim();

      if (estado === 'ERROR') {
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado,       'PENDIENTE');
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.detalleError, '');
        resetGen++;
      }

      // Solo volver a FIGURITA_CREADA si el PNG ya existe en Drive
      if (estado === 'ERROR_EMAIL' && figId) {
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado,       'FIGURITA_CREADA');
        setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.detalleError, '');
        resetEmail++;
      }
    });

    ui.alert(
      '✅ Reprocesamiento completado',
      `Errores de generación reiniciados: ${resetGen}\n` +
      `Errores de envío reiniciados: ${resetEmail}`,
      ui.ButtonSet.OK
    );

  } finally {
    lock.releaseLock();
  }
}


// ──────────────────────────────────────────────────────────────
// RESUMEN DE ESTADOS
// ──────────────────────────────────────────────────────────────

function mostrarResumenEstados() {
  const sheet  = getSheet_();
  const colMap = getColumnMap_(sheet);
  const data   = sheet.getDataRange().getValues();

  const conteo = {
    'PENDIENTE':       0,
    'PROCESANDO':      0,
    'FIGURITA_CREADA': 0,
    'EMAIL_ENVIADO':   0,
    'ERROR':           0,
    'ERROR_EMAIL':     0,
    'OTRO':            0,
  };

  data.slice(1).forEach(row => {
    const estado = String(getCell_(row, colMap, CONFIG.COLUMNS.estado) || '').trim();
    if (!estado)                        conteo['PENDIENTE']++;
    else if (conteo[estado] !== undefined) conteo[estado]++;
    else                                conteo['OTRO']++;
  });

  const total = data.length - 1;
  const quota = MailApp.getRemainingDailyQuota();

  const lines = [
    `Total de registros: ${total}`,
    ``,
    `⬜ PENDIENTE:        ${conteo['PENDIENTE']}`,
    `🟡 PROCESANDO:       ${conteo['PROCESANDO']}`,
    `🔵 FIGURITA_CREADA:  ${conteo['FIGURITA_CREADA']}`,
    `🟢 EMAIL_ENVIADO:    ${conteo['EMAIL_ENVIADO']}`,
    `🔴 ERROR:            ${conteo['ERROR']}`,
    `🟠 ERROR_EMAIL:      ${conteo['ERROR_EMAIL']}`,
    conteo['OTRO'] > 0 ? `⚠️ ESTADO DESCONOCIDO: ${conteo['OTRO']}` : null,
    ``,
    `📧 Cuota de mail restante hoy: ${quota}`,
  ].filter(l => l !== null);

  SpreadsheetApp.getUi().alert(
    '📊 Resumen de estados',
    lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}