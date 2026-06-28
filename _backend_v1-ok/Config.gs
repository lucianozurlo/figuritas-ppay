// ============================================================
// Config.gs — Configuración global compartida por todos los scripts
//
// INSTRUCCIONES:
// 1. Reemplazar todos los valores que dicen REEMPLAZAR_CON_...
//    con los IDs reales antes de publicar el Web App.
// 2. Verificar que los COLUMNS coincidan exactamente con los
//    encabezados del Sheet (ejecutar validarConfiguracion()).
// ============================================================

const CONFIG = {

  // Nombre exacto de la hoja dentro del Google Sheet
  SHEET_NAME: 'Registro de Figuritas',

  // ID de la carpeta 03_Figuritas_Generadas en Drive
  // (donde se guardan los PNG finales)
  FOLDER_FIGURITAS_ID: '1ITsMucNJBkiV96MB2kCaq_KeqblJgW8N',

  // ID de la carpeta donde se guardan las fotos originales subidas por el form
  // Puede ser la misma que 03 o una subcarpeta 00_Fotos_Originales
  FOLDER_FOTOS_ID: '1nAfi2Db1Y8-6z9CkRW6xZPvwqDm8qzON',

  // ID de la presentación Plantilla Figurita en Drive
  SLIDE_TEMPLATE_ID: '18ztUYRiU3h6GnWcO-No3J7Hyk6nIWTA93jiCJvzCNds',

  // ID de la presentación del MURAL (slide prediseñada, 1920×1080)
  MURAL_PRESENTATION_ID: '1hO-W_Schmy3Hwt7d1lopXTRAtPsTpjkmfexIE5R8kSI',

  // Altura del header del mural en cm (área reservada para títulos/logos)
  MURAL_HEADER_CM: 2,

  // Padding del área de figuritas en pt (30px ≈ 22.5pt a 96dpi)
  MURAL_PADDING_PT: 22.5,

  // Texto alternativo exacto de la forma de foto en la plantilla de Slides
  ALT_TEXT_FOTO: 'FOTO_PERFIL_REEMPLAZAR',

  // Tamaño máximo de lote por ejecución de generación
  // Reducido a 15 para evitar timeouts (límite Apps Script: 6 min)
  BATCH_SIZE_GENERACION: 15,

  // Tamaño máximo de lote por ejecución de envío de correos
  BATCH_SIZE_EMAIL: 50,

  // Mínimo de cuota de mail restante para continuar enviando
  // (reserva para reprocesos urgentes)
  QUOTA_BUFFER_EMAIL: 10,

  // Minutos máximos en estado PROCESANDO antes de considerarlo zombie
  PROCESANDO_TIMEOUT_MINUTOS: 15,

  // ──────────────────────────────────────────────────────────
  // COLUMNAS DEL SHEET
  //
  // El Sheet ya no es generado por Google Forms — ahora el
  // script WebApp.gs escribe las filas directamente.
  // Los encabezados deben crearse manualmente (ver runbook).
  //
  // Modificar estos valores solo si se cambian los encabezados
  // del Sheet. Los nombres aquí deben coincidir exactamente.
  // ──────────────────────────────────────────────────────────
  COLUMNS: {
    // Datos del participante (escritos por doPost)
    nombre:              'Nombre',
    mail:                'Mail',
    area:                'Area',
    superpoder:          'Superpoder',
    actitud:             'Actitud',
    consentimientoMural: 'Consentimiento mural',

    // Foto (escrita por doPost, apunta al archivo original en Drive)
    idArchivoDrive:      'id_archivo_drive',      // ID del archivo de foto original
    urlArchivoDrive:     'url_archivo_drive',      // URL directa de la foto (para =IMAGE() y generación)

    // Columnas operativas (escritas por el script de procesamiento)
    idFigurita:          'id_figurita',
    idFiguraGenerada:    'id_figurita_generada',   // ID del PNG final (distinto al de la foto original)
    urlFiguraGenerada:   'url_figurita_generada',  // URL del PNG final
    estado:              'estado',
    detalleError:        'detalle_error',
    timestampProcesando: 'timestamp_procesando',
    timestampSubida:     'fecha_subida',
    timestampGeneracion: 'fecha_generacion',
  }
};