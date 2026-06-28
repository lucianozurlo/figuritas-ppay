// ============================================================
// Mural.gs — Mural de figuritas con auto-encadenamiento
// Procesa UNA SLIDE por ejecución, encadena via triggers.
// ============================================================

const CM_TO_PT_ = 28.3465;
const MURAL_IMG_ALT_ = "MURAL_FIGURITA";
const MURAL_BG_ALT_ = "MURAL_FONDO";
const PROP_KEY_ = "MURAL_ESTADO";
const TRIGGER_FN_ = "continuarMuralAuto_";
const COLS_FIJO_ = 14;
const ROWS_FIJO_ = 6;
const MAX_POR_SLIDE_ = 30; // máx figuritas por ejecución para evitar timeout

// Fondos rotativos (se repiten en ciclo)
const FONDOS_IDS_ = [
  "1qNYSUV3JrfhIAy3kbAHFhj9Uep25NcBt",
  "15U6zRJZ1TK8PtHULLMDscpeNnDGIOd6B",
  "1FLCtmxt9kUu9rtc88SHWvWm5uJOhvdAv",
];

// ── ENTRADA PRINCIPAL ────────────────────────────────────────

function generarMural() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty(PROP_KEY_)) {
    const resp = ui.alert(
      "⚠️ Mural en progreso",
      "¿Descartarlo y empezar de cero?",
      ui.ButtonSet.YES_NO,
    );
    if (resp !== ui.Button.YES) return;
    cancelarTriggersMural_();
    props.deleteProperty(PROP_KEY_);
  }

  try {
    const figuritas = obtenerFiguritasParaMural_();
    if (figuritas.length === 0) {
      ui.alert(
        "Sin figuritas",
        "No hay figuritas con consentimiento de mural listas todavía.",
        ui.ButtonSet.OK,
      );
      return;
    }

    const barajadas = barajar_(figuritas);
    const n = barajadas.length;
    const porSlide = COLS_FIJO_ * ROWS_FIJO_; // 84 figuritas por slide
    const totalSlides = Math.ceil(n / porSlide);

    // Calcular cellSize con las dimensiones reales
    const { areaW, areaH } = calcularArea_();
    const cellSize = Math.min(areaW / COLS_FIJO_, areaH / ROWS_FIJO_ / 1.25);

    Logger.log(
      `[generarMural] ${n} figuritas → ${porSlide}/slide → ${totalSlides} slide(s)`,
    );

    // Preparar presentación: solo dejar la slide 1
    const pres = SlidesApp.openById(CONFIG.MURAL_PRESENTATION_ID);
    const existentes = pres.getSlides();
    for (let i = existentes.length - 1; i >= 1; i--) {
      existentes[i].remove();
    }
    // Limpiar imágenes de la slide 1 (restos de ejecuciones anteriores)
    eliminarTodasLasImagenes_(pres.getSlides()[0]);
    pres.saveAndClose();

    // Crear slides adicionales una por una con save intermedio
    // para que la API registre cada slide correctamente
    for (let s = 1; s < totalSlides; s++) {
      const presLoop = SlidesApp.openById(CONFIG.MURAL_PRESENTATION_ID);
      const slides = presLoop.getSlides();
      // Duplicar la slide 0 (el diseño original)
      slides[0].duplicate();
      // La copia queda en posición 1 — moverla al final
      const slidesNow = presLoop.getSlides();
      if (slidesNow.length > 2) {
        slidesNow[1].move(slidesNow.length - 1);
      }
      presLoop.saveAndClose();
      Logger.log(`[generarMural] Slide ${s + 1}/${totalSlides} creada`);
    }

    // Guardar estado
    const estado = {
      figuritas: barajadas,
      porSlide: porSlide,
      totalSlides: totalSlides,
      slideActual: 0,
      imgActual: 0, // índice dentro de la slide actual (para lotes parciales)
      cols: COLS_FIJO_,
      cellSize: cellSize,
      n: n,
    };
    props.setProperty(PROP_KEY_, JSON.stringify(estado));

    // Procesar primer lote
    const res = procesarLoteActual_();

    if (res.completado) {
      ui.alert(
        "✅ Mural listo",
        `${n} figuritas en ${totalSlides} slide(s).`,
        ui.ButtonSet.OK,
      );
      return;
    }

    crearTriggerContinuacion_();

    const est = JSON.parse(props.getProperty(PROP_KEY_));
    ui.alert(
      "🖼️ Mural iniciado",
      `Procesando slide 1/${totalSlides}...\n` +
        `Las restantes se generan automáticamente.\n` +
        `Recibirás un mail al terminar.`,
      ui.ButtonSet.OK,
    );
  } catch (err) {
    Logger.log("[generarMural] ERROR: " + err.toString());
    cancelarTriggersMural_();
    props.deleteProperty(PROP_KEY_);
    ui.alert("❌ Error", err.message, ui.ButtonSet.OK);
  }
}

// ── CONTINUACIÓN AUTOMÁTICA ──────────────────────────────────

function continuarMuralAuto_() {
  cancelarTriggersMural_();
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_KEY_);
  if (!raw) return;

  try {
    const estado = JSON.parse(raw);
    if (estado.slideActual >= estado.totalSlides) {
      props.deleteProperty(PROP_KEY_);
      return;
    }

    const res = procesarLoteActual_();
    Logger.log(
      `[continuarMuralAuto_] Lote: ${res.insertadas} ok, ${res.errores} err. Completado: ${res.completado}`,
    );

    if (!res.completado) {
      crearTriggerContinuacion_();
    } else {
      // Guardar timestamp del mural recién generado (ISO, UTC)
      PropertiesService.getScriptProperties().setProperty(
        "MURAL_ULTIMO_TS",
        new Date().toISOString(),
      );
      // Exportar automáticamente
      exportarMuralSilencioso_();
    }
  } catch (err) {
    Logger.log("[continuarMuralAuto_] ERROR: " + err.toString());
    cancelarTriggersMural_();
    PropertiesService.getScriptProperties().deleteProperty(PROP_KEY_);
  }
}

// ── PROCESAR UN LOTE (máx MAX_POR_SLIDE_ figuritas) ─────────

/**
 * Procesa hasta MAX_POR_SLIDE_ figuritas de la slide actual.
 * Si la slide tiene más figuritas (84 total), se procesan en
 * múltiples lotes encadenados por trigger.
 */
function procesarLoteActual_() {
  const props = PropertiesService.getScriptProperties();
  const estado = JSON.parse(props.getProperty(PROP_KEY_));

  const {
    figuritas,
    porSlide,
    totalSlides,
    slideActual,
    imgActual,
    cols: c,
    cellSize: cs,
    n,
  } = estado;

  const slideInicio = slideActual * porSlide;
  const slideFin = Math.min(slideInicio + porSlide, n);

  // Lote: desde imgActual hasta imgActual + MAX_POR_SLIDE_
  const loteInicio = slideInicio + imgActual;
  const loteFin = Math.min(loteInicio + MAX_POR_SLIDE_, slideFin);
  const enEsteLote = figuritas.slice(loteInicio, loteFin);

  Logger.log(
    `[procesarLote] Slide ${slideActual + 1}/${totalSlides}, imgs ${imgActual}–${imgActual + enEsteLote.length - 1}`,
  );

  const { headerPt, padTop, padLeft } = calcularArea_();
  const dim = getDimensionesMural_();
  const RATIO = 1.25;
  const gap = Math.max(0.5, cs * 0.025);
  const imgW = cs - gap * 2;
  const imgH = imgW * RATIO;

  const pres = SlidesApp.openById(CONFIG.MURAL_PRESENTATION_ID);
  const slides = pres.getSlides();

  if (slideActual >= slides.length) {
    throw new Error(
      `Slide ${slideActual} no existe (hay ${slides.length} slides).`,
    );
  }
  const slide = slides[slideActual];

  // Si es el primer lote de esta slide: aplicar fondo
  if (imgActual === 0) {
    aplicarFondo_(slide, slideActual, dim.w, dim.h);
  }

  let insertadas = 0;
  let errores = 0;

  for (let i = 0; i < enEsteLote.length; i++) {
    const idxEnSlide = imgActual + i;
    const col = idxEnSlide % c;
    const row = Math.floor(idxEnSlide / c);
    const x = padLeft + col * cs + gap;
    const y = headerPt + padTop + row * (imgH + gap * 2) + gap;

    const fileId = enEsteLote[i].fileId;
    try {
      const blob = DriveApp.getFileById(fileId).getBlob();
      const img = slide.insertImage(blob);
      img.setLeft(x).setTop(y).setWidth(imgW).setHeight(imgH);
      img.setTitle(MURAL_IMG_ALT_);
      img.setDescription(enEsteLote[i].nombre || "");
      insertadas++;
    } catch (eImg) {
      Logger.log(
        `[procesarLote] Error img idx ${loteInicio + i} (${fileId}): ${eImg.message}`,
      );
      errores++;
    }
  }

  pres.saveAndClose();

  // Avanzar estado
  const nuevoImgActual = imgActual + enEsteLote.length;
  const slideTerminada = nuevoImgActual >= slideFin - slideInicio;

  if (slideTerminada) {
    estado.slideActual = slideActual + 1;
    estado.imgActual = 0;
  } else {
    estado.imgActual = nuevoImgActual;
  }

  const completado = estado.slideActual >= totalSlides;

  if (completado) {
    props.deleteProperty(PROP_KEY_);
  } else {
    props.setProperty(PROP_KEY_, JSON.stringify(estado));
  }

  return { insertadas, errores, completado };
}

// ── APLICAR FONDO ────────────────────────────────────────────

function aplicarFondo_(slide, slideIndex, slideW, slideH) {
  const fondoId = FONDOS_IDS_[slideIndex % FONDOS_IDS_.length];
  try {
    // Eliminar imágenes anteriores (fondos viejos)
    eliminarTodasLasImagenes_(slide);

    // Insertar el fondo nuevo
    const blob = DriveApp.getFileById(fondoId).getBlob();
    const bg = slide.insertImage(blob);
    bg.setLeft(0).setTop(0).setWidth(slideW).setHeight(slideH);
    bg.setTitle(MURAL_BG_ALT_);

    // En lugar de mandar el fondo al fondo (sendToBack),
    // traer al frente TODOS los demás elementos del diseño
    // (textos, logos, formas) para que queden por encima del fondo.
    // El fondo recién insertado queda al final por defecto.
    const bgId = bg.getObjectId();
    slide.getPageElements().forEach((pe) => {
      try {
        if (pe.getObjectId() !== bgId) pe.bringToFront();
      } catch (_) {}
    });

    Logger.log(
      `[aplicarFondo_] Fondo ${(slideIndex % 3) + 1} → slide ${slideIndex + 1}`,
    );
  } catch (e) {
    Logger.log(`[aplicarFondo_] Error fondo ${fondoId}: ${e.message}`);
  }
}

// ── HELPERS ──────────────────────────────────────────────────

function getDimensionesMural_() {
  const pres = SlidesApp.openById(CONFIG.MURAL_PRESENTATION_ID);
  const w = pres.getPageWidth();
  const h = pres.getPageHeight();
  pres.saveAndClose();
  return { w, h };
}

function calcularArea_() {
  const dim = getDimensionesMural_();
  const headerPt = CONFIG.MURAL_HEADER_CM * CM_TO_PT_;
  const padBase = dim.w * 0.015625;
  const padTop = 1.0 * CM_TO_PT_;
  const padLeft = padBase * 3;
  return {
    headerPt,
    padTop,
    padLeft,
    padBase,
    slideW: dim.w,
    slideH: dim.h,
    areaW: dim.w - padLeft * 2,
    areaH: dim.h - headerPt - padTop - padBase,
  };
}

function eliminarTodasLasImagenes_(slide) {
  if (!slide) return;
  try {
    slide.getImages().forEach((img) => {
      try {
        img.remove();
      } catch (_) {}
    });
  } catch (_) {}
}

function limpiarFiguritasAnteriores_(slide) {
  if (!slide) return;
  try {
    slide.getImages().forEach((img) => {
      try {
        if (img.getTitle() === MURAL_IMG_ALT_) img.remove();
      } catch (_) {}
    });
  } catch (_) {}
}

function obtenerFiguritasParaMural_() {
  const sheet = getSheet_();
  const colMap = getColumnMap_(sheet);
  const data = sheet.getDataRange().getValues();

  const colEstado = colMap[CONFIG.COLUMNS.estado];
  const colMural = colMap[CONFIG.COLUMNS.consentimientoMural];
  const colFigId = colMap[CONFIG.COLUMNS.idFiguraGenerada];
  const colNombre = colMap[CONFIG.COLUMNS.nombre];

  if (!colEstado || !colMural || !colFigId) {
    throw new Error("Columnas faltantes en el Sheet.");
  }

  return data.slice(1).reduce((acc, row) => {
    const estado = String(row[colEstado - 1] || "").trim();
    const mural = String(row[colMural - 1] || "").trim();
    const fileId = String(row[colFigId - 1] || "").trim();
    const nombre = String(row[colNombre - 1] || "").trim();
    const estadoOK = estado === "EMAIL_ENVIADO" || estado === "FIGURITA_CREADA";
    const muralOK = mural === "Sí" || mural === "Si" || mural === "TRUE";
    if (estadoOK && muralOK && fileId) acc.push({ nombre, fileId });
    return acc;
  }, []);
}

function crearTriggerContinuacion_() {
  ScriptApp.newTrigger(TRIGGER_FN_)
    .timeBased()
    .after(90 * 1000)
    .create();
}

function cancelarTriggersMural_() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === TRIGGER_FN_)
    .forEach((t) => {
      try {
        ScriptApp.deleteTrigger(t);
      } catch (_) {}
    });
}

function barajar_(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

// ============================================================
// VERSIÓN SILENCIOSA (sin UI) — para endpoint y triggers
// ============================================================

/**
 * Igual que generarMural() pero sin alerts de UI.
 * Usada cuando se llama desde el endpoint HTTP.
 */
function generarMuralSilencioso_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_KEY_)) {
    Logger.log("[generarMuralSilencioso_] Ya hay un mural en progreso.");
    return;
  }

  const figuritas = obtenerFiguritasParaMural_();
  if (figuritas.length === 0) {
    Logger.log("[generarMuralSilencioso_] Sin figuritas elegibles.");
    return;
  }

  const barajadas = barajar_(figuritas);
  const n = barajadas.length;
  const porSlide = COLS_FIJO_ * ROWS_FIJO_;
  const totalSlides = Math.ceil(n / porSlide);
  const { areaW, areaH } = calcularArea_();
  const cellSize = Math.min(areaW / COLS_FIJO_, areaH / ROWS_FIJO_ / 1.25);

  Logger.log(
    `[generarMuralSilencioso_] ${n} figuritas → ${totalSlides} slide(s)`,
  );

  const pres = SlidesApp.openById(CONFIG.MURAL_PRESENTATION_ID);
  const existentes = pres.getSlides();
  for (let i = existentes.length - 1; i >= 1; i--) existentes[i].remove();
  eliminarTodasLasImagenes_(pres.getSlides()[0]);
  pres.saveAndClose();

  for (let s = 1; s < totalSlides; s++) {
    const p = SlidesApp.openById(CONFIG.MURAL_PRESENTATION_ID);
    const sl = p.getSlides();
    sl[0].duplicate();
    const slNow = p.getSlides();
    if (slNow.length > 2) slNow[1].move(slNow.length - 1);
    p.saveAndClose();
  }

  const estado = {
    figuritas: barajadas,
    porSlide,
    totalSlides,
    slideActual: 0,
    imgActual: 0,
    cols: COLS_FIJO_,
    cellSize,
    n,
  };
  props.setProperty(PROP_KEY_, JSON.stringify(estado));

  const res = procesarLoteActual_();
  if (!res.completado) crearTriggerContinuacion_();
}

// ============================================================
// EXPORT: PNG + PPTX en carpeta versionada de Drive
// ============================================================

/**
 * Exporta la presentación del mural a:
 *   - Un archivo PPTX
 *   - Un PNG por cada slide (1920×1080)
 * Todo en una subcarpeta versionada dentro de mural-final/.
 *
 * Llamada automáticamente al completar el mural, o manualmente
 * desde el menú Sheet → 🎴 Figuritas Admin → 📦 Exportar mural.
 */
function exportarMural() {
  const ui = SpreadsheetApp.getUi();
  try {
    const carpetaVersion = crearCarpetaVersion_();
    const { pptxFile, pdfFile, pngFiles } = ejecutarExport_(carpetaVersion);

    ui.alert(
      "✅ Mural exportado",
      `Carpeta: ${carpetaVersion.getName()}\n` +
        `PPTX: ${pptxFile.getName()}\n` +
        `PDF: ${pdfFile.getName()}\n` +
        `PNGs: ${pngFiles.length} slides\n\n` +
        `Link: ${carpetaVersion.getUrl()}`,
      ui.ButtonSet.OK,
    );
  } catch (err) {
    Logger.log("[exportarMural] ERROR: " + err.toString());
    ui.alert("❌ Error exportando", err.message, ui.ButtonSet.OK);
  }
}

/**
 * Versión silenciosa del export — para llamar al terminar el mural.
 */
function exportarMuralSilencioso_() {
  try {
    const carpetaVersion = crearCarpetaVersion_();
    const { pptxFile, pdfFile, pngFiles } = ejecutarExport_(carpetaVersion);

    // Usar el mail guardado desde el endpoint, o el del propietario como fallback
    const props = PropertiesService.getScriptProperties();
    const mailDest =
      props.getProperty("MURAL_NOTIF_MAIL") ||
      Session.getActiveUser().getEmail();
    props.deleteProperty("MURAL_NOTIF_MAIL");

    Logger.log(
      "[exportarMuralSilencioso_] Export completado: " +
        carpetaVersion.getName(),
    );

    try {
      const pngLinks = pngFiles
        .map(
          (f, i) => `  Slide ${String(i + 1).padStart(2, "0")}: ${f.getUrl()}`,
        )
        .join("\n");

      MailApp.sendEmail({
        to: mailDest,
        subject: "✅ Mural listo — Mi Figurita Personal Pay",
        body:
          `¡El mural fue generado y exportado correctamente!\n\n` +
          `Carpeta: ${carpetaVersion.getName()}\n` +
          `Link: ${carpetaVersion.getUrl()}\n\n` +
          `— PPTX —\n  ${pptxFile.getUrl()}\n\n` +
          `— PDF —\n  ${pdfFile.getUrl()}\n\n` +
          `— PNGs por slide —\n${pngLinks}`,
      });
    } catch (eMail) {
      Logger.log(
        "[exportarMuralSilencioso_] Error enviando mail: " + eMail.toString(),
      );
    }
  } catch (err) {
    Logger.log("[exportarMuralSilencioso_] ERROR: " + err.toString());
  }
}

/**
 * Crea la carpeta versionada: mural-final/vN_YYYY-MM-DD_HH-mm/
 */
function crearCarpetaVersion_() {
  const raiz = DriveApp.getFolderById(CONFIG.FOLDER_RAIZ_ID);

  // Obtener o crear carpeta mural-final/
  let carpetaMural;
  const it = raiz.getFoldersByName(CONFIG.MURAL_EXPORT_FOLDER_NAME);
  carpetaMural = it.hasNext()
    ? it.next()
    : raiz.createFolder(CONFIG.MURAL_EXPORT_FOLDER_NAME);

  // Calcular número de versión
  const versiones = [];
  const itV = carpetaMural.getFolders();
  while (itV.hasNext()) {
    const nombre = itV.next().getName();
    const match = nombre.match(/^v(\d+)_/);
    if (match) versiones.push(parseInt(match[1]));
  }
  const numVersion = versiones.length > 0 ? Math.max(...versiones) + 1 : 1;

  // Nombre de la carpeta: vN_YYYY-MM-DD_HH-mm
  const ahora = new Date();
  const fechaStr = Utilities.formatDate(
    ahora,
    "America/Argentina/Buenos_Aires",
    "yyyy-MM-dd_HH-mm",
  );
  const nombreCarpeta = `v${numVersion}_${fechaStr}`;

  return carpetaMural.createFolder(nombreCarpeta);
}

/**
 * Hace el export real: PPTX + PNGs por slide.
 */
function ejecutarExport_(carpeta) {
  const presentationId = CONFIG.MURAL_PRESENTATION_ID;
  const token = ScriptApp.getOAuthToken();
  const pres = SlidesApp.openById(presentationId);
  const slides = pres.getSlides();
  const nombreBase = carpeta.getName();

  // ── EXPORT PPTX ───────────────────────────────────────────
  const pptxUrl = `https://docs.google.com/presentation/d/${presentationId}/export/pptx`;
  const pptxResp = UrlFetchApp.fetch(pptxUrl, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  if (pptxResp.getResponseCode() !== 200) {
    throw new Error(
      `Error exportando PPTX: HTTP ${pptxResp.getResponseCode()}`,
    );
  }
  const pptxBlob = pptxResp.getBlob().setName(`${nombreBase}.pptx`);
  const pptxFile = carpeta.createFile(pptxBlob);
  Logger.log(`[ejecutarExport_] PPTX: ${pptxFile.getName()}`);

  // ── EXPORT PDF ────────────────────────────────────────────
  const pdfUrl = `https://docs.google.com/presentation/d/${presentationId}/export/pdf`;
  const pdfResp = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  if (pdfResp.getResponseCode() !== 200) {
    throw new Error(`Error exportando PDF: HTTP ${pdfResp.getResponseCode()}`);
  }
  const pdfBlob = pdfResp.getBlob().setName(`${nombreBase}.pdf`);
  const pdfFile = carpeta.createFile(pdfBlob);
  Logger.log(`[ejecutarExport_] PDF: ${pdfFile.getName()}`);

  // ── EXPORT PNG POR SLIDE ───────────────────────────────────
  const pngFiles = [];
  slides.forEach((slide, idx) => {
    const slideId = slide.getObjectId();
    // scale=2 sobre 960×540pt base = 1920×1080px
    const pngUrl = `https://docs.google.com/presentation/d/${presentationId}/export/png?pageid=${slideId}&scale=2`;
    const pngResp = UrlFetchApp.fetch(pngUrl, {
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    });
    if (pngResp.getResponseCode() !== 200) {
      Logger.log(
        `[ejecutarExport_] Error PNG slide ${idx + 1}: HTTP ${pngResp.getResponseCode()}`,
      );
      return;
    }
    const pngBlob = pngResp
      .getBlob()
      .setName(`${nombreBase}_slide${String(idx + 1).padStart(2, "0")}.png`);
    const pngFile = carpeta.createFile(pngBlob);
    pngFiles.push(pngFile);
    Logger.log(`[ejecutarExport_] PNG slide ${idx + 1}: ${pngFile.getName()}`);
  });

  pres.saveAndClose();
  return { pptxFile, pdfFile, pngFiles };
}
