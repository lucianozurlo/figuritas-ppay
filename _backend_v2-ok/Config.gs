// ============================================================
// Config.gs — Configuración global compartida por todos los scripts
// ============================================================

const CONFIG = {

  // Nombre exacto de la hoja dentro del Google Sheet
  SHEET_NAME: 'Registro de Figuritas',

  // ID de la carpeta 03_Figuritas_Generadas en Drive (PNGs finales)
  FOLDER_FIGURITAS_ID: '1ITsMucNJBkiV96MB2kCaq_KeqblJgW8N',

  // ID de la carpeta 00_Fotos_Originales en Drive (fotos del formulario)
  FOLDER_FOTOS_ID: '1nAfi2Db1Y8-6z9CkRW6xZPvwqDm8qzON',

  // ID de la presentación Plantilla Figurita en Drive
  SLIDE_TEMPLATE_ID: '18ztUYRiU3h6GnWcO-No3J7Hyk6nIWTA93jiCJvzCNds',

  // ID de la presentación del MURAL (1920×1080)
  MURAL_PRESENTATION_ID: '1hO-W_Schmy3Hwt7d1lopXTRAtPsTpjkmfexIE5R8kSI',

  // ID de la carpeta raíz del evento (donde se crea mural-final/)
  FOLDER_RAIZ_ID: '16DvBIKsMIyiGaCn8jYZCCwUqF9uA_G14',

  // Nombre de la carpeta de exports del mural
  MURAL_EXPORT_FOLDER_NAME: 'mural-final',

  // Altura del header del mural en cm
  MURAL_HEADER_CM: 2.6,

  // Padding base del área de figuritas en pt (≈30px a 96dpi)
  MURAL_PADDING_PT: 22.5,

  // Texto alternativo de la forma de foto en la plantilla de Slides
  ALT_TEXT_FOTO: 'FOTO_PERFIL_REEMPLAZAR',

  // Tamaño máximo de lote por ejecución de generación
  BATCH_SIZE_GENERACION: 15,

  // Tamaño máximo de lote por ejecución de envío de correos
  BATCH_SIZE_EMAIL: 50,

  // Reserva mínima de cuota de mail
  QUOTA_BUFFER_EMAIL: 10,

  // Minutos máximos en estado PROCESANDO antes de considerarlo zombie
  PROCESANDO_TIMEOUT_MINUTOS: 15,

  COLUMNS: {
    nombre:              'Nombre',
    mail:                'Mail',
    area:                'Area',
    superpoder:          'Superpoder',
    actitud:             'Actitud',
    consentimientoMural: 'Consentimiento mural',
    idArchivoDrive:      'id_archivo_drive',
    urlArchivoDrive:     'url_archivo_drive',
    idFigurita:          'id_figurita',
    idFiguraGenerada:    'id_figurita_generada',
    urlFiguraGenerada:   'url_figurita_generada',
    estado:              'estado',
    detalleError:        'detalle_error',
    timestampProcesando: 'timestamp_procesando',
    timestampSubida:     'fecha_subida',
    timestampGeneracion: 'fecha_generacion',
  }
};