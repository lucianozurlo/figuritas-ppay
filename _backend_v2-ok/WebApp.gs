// ============================================================
// WebApp.gs — Endpoint público del formulario de figuritas
// ============================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({
        status: "error",
        message: "DEBUG: sin postData (ejecutado desde editor?)",
      });
    }

    const payload = JSON.parse(e.postData.contents);
    Logger.log("[doPost] mail=" + payload.mail + " area=" + payload.area);

    const errores = validarPayload_(payload);
    if (errores.length > 0) {
      return jsonResponse_({
        status: "error",
        message: "Validación: " + errores.join(" | "),
      });
    }

    // Buscar si el mail ya tiene una fila (para sobrescribir en vez de duplicar)
    const filaExistente = buscarFilaPorMail_(payload.mail);

    // 1. Guardar foto original en Drive
    let fileId, directUrl;
    try {
      const r = guardarFotoEnDrive_(payload.fotoBase64, payload.nombre);
      fileId = r.fileId;
      directUrl = r.directUrl;
    } catch (eFoto) {
      return jsonResponse_({
        status: "error",
        message: "DEBUG guardarFoto: " + eFoto.toString(),
      });
    }

    // 2. Escribir fila (nueva o sobrescribir la existente)
    let rowIndex;
    try {
      if (filaExistente) {
        // Sobrescribir la fila existente con los datos nuevos
        rowIndex = filaExistente;
        actualizarFila_(rowIndex, {
          nombre: payload.nombre,
          mail: payload.mail,
          area: payload.area,
          superpoder: payload.superpoder,
          actitud: payload.actitud,
          consentimientoMural: payload.consentimientoMural ? "Sí" : "",
          idArchivoDrive: fileId,
          urlArchivoDrive: directUrl,
          estado: "PROCESANDO",
        });
      } else {
        rowIndex = escribirFila_({
          nombre: payload.nombre,
          mail: payload.mail,
          area: payload.area,
          superpoder: payload.superpoder,
          actitud: payload.actitud,
          consentimientoMural: payload.consentimientoMural ? "Sí" : "",
          idArchivoDrive: fileId,
          urlArchivoDrive: directUrl,
          estado: "PROCESANDO",
        });
      }
    } catch (eFila) {
      return jsonResponse_({
        status: "error",
        message: "DEBUG escribirFila: " + eFila.toString(),
      });
    }

    // 3. Generar la figurita EN EL MOMENTO (sincrónico)
    try {
      const resultado = generarFigurita_(rowIndex, {
        nombre: payload.nombre,
        area: payload.area,
        superpoder: payload.superpoder,
        actitud: payload.actitud,
        fotoId: fileId,
      });
      return jsonResponse_({
        status: "ready",
        url: resultado.url,
        base64: resultado.base64,
      });
    } catch (genErr) {
      Logger.log("[doPost] Error generando: " + genErr.toString());
      marcarPendiente_(rowIndex, genErr.toString());
      return jsonResponse_({
        status: "error",
        message: "DEBUG generarFigurita: " + genErr.toString(),
      });
    }
  } catch (err) {
    Logger.log("[doPost] Error: " + err.toString());
    return jsonResponse_({
      status: "error",
      message: "DEBUG fatal: " + err.toString(),
    });
  }
}

/**
 * doGet — polling del frontend + health check.
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === "check") {
    const mail = e.parameter.mail || "";
    return jsonResponse_(checkFiguritaLista_(mail));
  }

  if (action === "exists") {
    const mail = e.parameter.mail || "";
    const fila = buscarFilaPorMail_(mail);
    return jsonResponse_({ exists: fila ? true : false });
  }

  // ── ESTADO DEL MURAL ────────────────────────────────────────
  if (action === "mural_status") {
    const props = PropertiesService.getScriptProperties();
    const enProceso = !!props.getProperty("MURAL_ESTADO");

    // tieneMural: hay al menos una versión exportada en Drive
    let tieneMural = false;
    try {
      const raiz = DriveApp.getFolderById(CONFIG.FOLDER_RAIZ_ID);
      const it = raiz.getFoldersByName(CONFIG.MURAL_EXPORT_FOLDER_NAME);
      if (it.hasNext()) {
        tieneMural = it.next().getFolders().hasNext();
      }
    } catch (_) {}

    // figuritasNuevas: figuritas con consentimiento de mural cuya fecha_generacion
    // es POSTERIOR a la última generación del mural.
    // Si nunca se generó un mural, figuritasNuevas = total elegibles.
    let figuritasNuevas = 0;
    let totalElegibles = 0;
    try {
      const sheet = getSheet_();
      const colMap = getColumnMap_(sheet);
      const data = sheet.getDataRange().getValues();
      const colEstado = colMap[CONFIG.COLUMNS.estado];
      const colMural = colMap[CONFIG.COLUMNS.consentimientoMural];
      const colFigId = colMap[CONFIG.COLUMNS.idFiguraGenerada];
      const colFecha = colMap[CONFIG.COLUMNS.timestampGeneracion];

      // Fecha del último mural (guardada como ISO string al generar)
      const ultimoMuralStr = props.getProperty("MURAL_ULTIMO_TS");
      const ultimoMuralTs = ultimoMuralStr
        ? new Date(ultimoMuralStr).getTime()
        : 0;

      if (colEstado && colMural && colFigId) {
        data.slice(1).forEach((row) => {
          const estado = String(row[colEstado - 1] || "").trim();
          const mural = String(row[colMural - 1] || "").trim();
          const fileId = String(row[colFigId - 1] || "").trim();
          const estadoOK =
            estado === "EMAIL_ENVIADO" || estado === "FIGURITA_CREADA";
          const muralOK = mural === "Sí" || mural === "Si";
          if (!estadoOK || !muralOK || !fileId) return;

          totalElegibles++;

          // Si no hay mural previo, todas son "nuevas"
          if (!ultimoMuralStr) {
            figuritasNuevas++;
            return;
          }

          // Comparar fecha_generacion de la figurita con el timestamp del último mural
          if (colFecha) {
            const fechaStr = String(row[colFecha - 1] || "").trim();
            if (fechaStr) {
              // Formato guardado: dd/MM/yyyy HH:mm:ss (Buenos Aires)
              // Convertir a Date para comparar
              const partes = fechaStr.match(
                /(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})/,
              );
              if (partes) {
                const fechaFig = new Date(
                  `${partes[3]}-${partes[2]}-${partes[1]}T${partes[4]}:${partes[5]}:${partes[6]}-03:00`,
                );
                if (fechaFig.getTime() > ultimoMuralTs) figuritasNuevas++;
              }
            }
          }
        });
      }
    } catch (_) {}

    return jsonResponse_({
      status: "ok",
      enProceso,
      tieneMural,
      totalElegibles,
      figuritasNuevas,
    });
  }

  // ── GENERAR MURAL ────────────────────────────────────────────
  // URL: .../exec?action=mural&mail=xxx@mail.com&overwrite=true/false
  if (action === "mural") {
    try {
      const props = PropertiesService.getScriptProperties();
      const enCurso = props.getProperty("MURAL_ESTADO");
      const overwrite = (e.parameter.overwrite || "").toLowerCase() === "true";
      const mail = (e.parameter.mail || "").trim();

      if (enCurso && !overwrite) {
        return jsonResponse_({
          status: "error",
          enProceso: true,
          message: "Hay un mural en proceso. Confirmá la sobreescritura.",
        });
      }

      // Guardar el mail de notificación en Script Properties
      if (mail) props.setProperty("MURAL_NOTIF_MAIL", mail);

      // Si hay proceso en curso y se confirma overwrite, cancelar triggers
      if (enCurso && overwrite) {
        ScriptApp.getProjectTriggers()
          .filter((t) => t.getHandlerFunction() === "continuarMuralAuto_")
          .forEach((t) => {
            try {
              ScriptApp.deleteTrigger(t);
            } catch (_) {}
          });
        props.deleteProperty("MURAL_ESTADO");
      }

      // Disparar generación en background via trigger de 1 segundo
      ScriptApp.newTrigger("generarMuralDesdeEndpoint_")
        .timeBased()
        .after(1000)
        .create();

      return jsonResponse_({
        status: "ok",
        message: "Generación del mural iniciada.",
      });
    } catch (err) {
      return jsonResponse_({ status: "error", message: err.toString() });
    }
  }

  return jsonResponse_({
    status: "ok",
    service: "Figuritas API",
    area_config: CONFIG.COLUMNS.area,
  });
}

function checkFiguritaLista_(mail) {
  if (!mail) return { status: "pending" };

  try {
    const sheet = getSheet_();
    const colMap = getColumnMap_(sheet);
    const data = sheet.getDataRange().getValues();

    const mailNorm = String(mail).trim().toLowerCase();
    const colMail = colMap[CONFIG.COLUMNS.mail];
    const colEstado = colMap[CONFIG.COLUMNS.estado];
    const colUrl = colMap[CONFIG.COLUMNS.urlFiguraGenerada];

    if (!colMail || !colEstado || !colUrl) return { status: "pending" };

    const row = data.slice(1).find((r) => {
      const rowMail = String(r[colMail - 1] || "")
        .trim()
        .toLowerCase();
      const rowEstado = String(r[colEstado - 1] || "").trim();
      return (
        rowMail === mailNorm &&
        (rowEstado === "FIGURITA_CREADA" || rowEstado === "EMAIL_ENVIADO")
      );
    });

    if (row) {
      const url = String(row[colUrl - 1] || "").trim();
      if (url) return { status: "ready", url };
    }
    return { status: "pending" };
  } catch (err) {
    Logger.log("[checkFiguritaLista_] " + err.toString());
    return { status: "pending" };
  }
}

// ============================================================
// VALIDACIÓN
// ============================================================

function validarPayload_(payload) {
  const errores = [];

  if (!payload.nombre || String(payload.nombre).trim().length < 2)
    errores.push("Nombre invalido.");

  if (
    !payload.mail ||
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(payload.mail).trim())
  )
    errores.push("Mail invalido.");

  if (!payload.area || String(payload.area).trim() === "")
    errores.push("Area invalida.");

  if (!payload.superpoder || String(payload.superpoder).trim().length < 3)
    errores.push("Superpoder invalido.");

  if (!payload.actitud || String(payload.actitud).trim().length < 3)
    errores.push("Actitud invalida.");

  if (!payload.fotoBase64 || payload.fotoBase64.length < 1000)
    errores.push("Foto invalida o ausente.");

  return errores;
}

// ============================================================
// GUARDAR FOTO EN DRIVE
// ============================================================

function guardarFotoEnDrive_(base64, nombre) {
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_FOTOS_ID);

  const nombreSanitizado = String(nombre)
    .replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚüÜñÑ\s]/g, "")
    .trim()
    .substring(0, 30);

  const fileName = `Foto_${nombreSanitizado}_${new Date().getTime()}.jpg`;
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, "image/jpeg", fileName);
  const file = folder.createFile(blob);

  return {
    fileId: file.getId(),
    directUrl: `https://drive.google.com/uc?export=view&id=${file.getId()}`,
  };
}

// ============================================================
// ESCRIBIR FILA EN SHEET
// ============================================================

function escribirFila_(datos) {
  const sheet = getSheet_();
  const colMap = getColumnMap_(sheet);
  const newRow = sheet.getLastRow() + 1;

  // Fecha y hora de subida en horario de Argentina (UTC-3)
  const fechaSubida = Utilities.formatDate(
    new Date(),
    "America/Argentina/Buenos_Aires",
    "dd/MM/yyyy HH:mm:ss",
  );

  const celdas = [
    { col: CONFIG.COLUMNS.nombre, val: datos.nombre },
    { col: CONFIG.COLUMNS.mail, val: datos.mail },
    { col: CONFIG.COLUMNS.area, val: datos.area },
    { col: CONFIG.COLUMNS.superpoder, val: datos.superpoder },
    { col: CONFIG.COLUMNS.actitud, val: datos.actitud },
    {
      col: CONFIG.COLUMNS.consentimientoMural,
      val: datos.consentimientoMural || "",
    },
    { col: CONFIG.COLUMNS.idArchivoDrive, val: datos.idArchivoDrive },
    { col: CONFIG.COLUMNS.urlArchivoDrive, val: datos.urlArchivoDrive },
    { col: CONFIG.COLUMNS.estado, val: datos.estado },
    { col: CONFIG.COLUMNS.timestampSubida, val: fechaSubida },
  ];

  celdas.forEach(({ col, val }) => {
    if (val === undefined || val === null) return;
    const colIdx = colMap[col];
    if (!colIdx) return;
    try {
      sheet.getRange(newRow, colIdx).setValue(val);
    } catch (err) {
      Logger.log(`[escribirFila_] Error en columna "${col}": ${err.message}`);
    }
  });

  Logger.log(`[escribirFila_] Fila ${newRow}: ${datos.nombre} <${datos.mail}>`);
  return newRow;
}

// ============================================================
// GENERACIÓN SINCRÓNICA DE FIGURITA
// ============================================================

function generarFigurita_(rowIndex, datos) {
  const sheet = getSheet_();
  const colMap = getColumnMap_(sheet);
  let presentationCopy = null;

  try {
    const fotoBlob = DriveApp.getFileById(datos.fotoId).getBlob();
    if (!fotoBlob || fotoBlob.getBytes().length < 500) {
      throw new Error("Blob de foto inválido o vacío.");
    }

    const nombreSlug = String(datos.nombre)
      .replace(/[^a-zA-Z0-9áéíóúüñ\s]/gi, "")
      .trim()
      .substring(0, 30);
    presentationCopy = DriveApp.getFileById(CONFIG.SLIDE_TEMPLATE_ID).makeCopy(
      `Temp_${nombreSlug}_${rowIndex}`,
    );

    const presentation = SlidesApp.openById(presentationCopy.getId());
    const slide = presentation.getSlides()[0];

    slide.replaceAllText("{{nombre}}", (datos.nombre || "").toUpperCase());
    slide.replaceAllText("{{area}}", (datos.area || "").toUpperCase());
    slide.replaceAllText("{{superpoder}}", datos.superpoder || "");
    slide.replaceAllText("{{actitud}}", datos.actitud || "");

    let imageShape = slide.getPageElements().find((el) => {
      try {
        return el.getTitle() === CONFIG.ALT_TEXT_FOTO;
      } catch (_) {
        return false;
      }
    });

    // Fallback: si no se encontró por título, buscar el SHAPE vacío
    // (las otras 4 formas tienen los marcadores {{...}}, la de foto está vacía)
    if (!imageShape) {
      imageShape = slide.getPageElements().find((el) => {
        try {
          if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE)
            return false;
          const txt = el.asShape().getText().asString().trim();
          return txt === "";
        } catch (_) {
          return false;
        }
      });
    }

    if (!imageShape) {
      throw new Error(
        "No se encontró la forma de foto en la plantilla (ni por título ni por forma vacía).",
      );
    }

    const left = imageShape.getLeft();
    const top = imageShape.getTop();
    const width = imageShape.getWidth();
    const height = imageShape.getHeight();

    // La foto llega cuadrada (1:1). Se centra usando el lado menor.
    const lado = Math.min(width, height);
    const offsetX = left + (width - lado) / 2;
    const offsetY = top + (height - lado) / 2;

    const insertedImage = slide.insertImage(fotoBlob);
    insertedImage
      .setLeft(offsetX)
      .setTop(offsetY)
      .setWidth(lado)
      .setHeight(lado);

    // Eliminar la forma placeholder
    imageShape.remove();

    // insertImage() siempre coloca la imagen al frente del orden Z.
    // La máscara del diseño queda detrás. Para que la máscara quede encima,
    // mandamos la foto atrás del todo — así todos los elementos del diseño
    // (incluyendo la máscara) quedan por encima de ella.
    const fotoPageEl = slide
      .getPageElements()
      .find((el) => el.getObjectId() === insertedImage.getObjectId());
    if (fotoPageEl) fotoPageEl.sendToBack();

    presentation.saveAndClose();

    // Exportar como PNG
    const exportUrl = `https://docs.google.com/presentation/d/${presentationCopy.getId()}/export/png?pageid=${slide.getObjectId()}&scale=2`;
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) {
      throw new Error(`Export PNG HTTP ${response.getResponseCode()}.`);
    }
    const pngBlob = response.getBlob();
    if (!pngBlob || pngBlob.getBytes().length < 1000) {
      throw new Error("PNG exportado vacío.");
    }
    pngBlob.setName(`Figurita_${nombreSlug}_${rowIndex}.png`);

    const pngFile = DriveApp.getFolderById(
      CONFIG.FOLDER_FIGURITAS_ID,
    ).createFile(pngBlob);
    try {
      pngFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
    } catch (_) {
      /* si la política del dominio no lo permite, sigue igual */
    }

    const urlFigura = `https://drive.google.com/thumbnail?id=${pngFile.getId()}&sz=w1080`;

    setCell_(
      sheet,
      rowIndex,
      colMap,
      CONFIG.COLUMNS.idFigurita,
      `FIG-${rowIndex}-${new Date().getTime()}`,
    );
    setCell_(
      sheet,
      rowIndex,
      colMap,
      CONFIG.COLUMNS.idFiguraGenerada,
      pngFile.getId(),
    );
    setCell_(
      sheet,
      rowIndex,
      colMap,
      CONFIG.COLUMNS.urlFiguraGenerada,
      urlFigura,
    );
    setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado, "FIGURITA_CREADA");

    const ahora = Utilities.formatDate(
      new Date(),
      "America/Argentina/Buenos_Aires",
      "dd/MM/yyyy HH:mm:ss",
    );
    setCell_(
      sheet,
      rowIndex,
      colMap,
      CONFIG.COLUMNS.timestampGeneracion,
      ahora,
    );

    // Devolver URL (para mostrar) y base64 (para descarga/compartir sin CORS)
    return {
      url: urlFigura,
      base64: Utilities.base64Encode(pngBlob.getBytes()),
    };
  } finally {
    if (presentationCopy) {
      try {
        DriveApp.getFileById(presentationCopy.getId()).setTrashed(true);
      } catch (_) {
        /* ignorar */
      }
    }
  }
}

// ============================================================
// HELPERS
// ============================================================

function marcarPendiente_(rowIndex, motivo) {
  try {
    const sheet = getSheet_();
    const colMap = getColumnMap_(sheet);
    setCell_(sheet, rowIndex, colMap, CONFIG.COLUMNS.estado, "PENDIENTE");
    setCell_(
      sheet,
      rowIndex,
      colMap,
      CONFIG.COLUMNS.detalleError,
      "Generación diferida: " + motivo,
    );
  } catch (e) {
    Logger.log("[marcarPendiente_] " + e.toString());
  }
}

function buscarFiguritaExistente_(mail) {
  try {
    const sheet = getSheet_();
    const colMap = getColumnMap_(sheet);
    const data = sheet.getDataRange().getValues();
    const mailNorm = String(mail).trim().toLowerCase();
    const colMail = colMap[CONFIG.COLUMNS.mail];
    const colUrl = colMap[CONFIG.COLUMNS.urlFiguraGenerada];
    if (!colMail || !colUrl) return null;

    const row = data.slice(1).find(
      (r) =>
        String(r[colMail - 1] || "")
          .trim()
          .toLowerCase() === mailNorm &&
        String(r[colUrl - 1] || "").trim() !== "",
    );
    return row ? String(row[colUrl - 1]).trim() : null;
  } catch (_) {
    return null;
  }
}

/**
 * Busca el número de fila (1-based) de un mail en el Sheet.
 * Devuelve el rowIndex o null si no existe.
 */
function buscarFilaPorMail_(mail) {
  try {
    const sheet = getSheet_();
    const colMap = getColumnMap_(sheet);
    const data = sheet.getDataRange().getValues();
    const mailNorm = String(mail).trim().toLowerCase();
    const colMail = colMap[CONFIG.COLUMNS.mail];
    if (!colMail) return null;

    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][colMail - 1] || "")
          .trim()
          .toLowerCase() === mailNorm
      ) {
        return i + 1; // rowIndex 1-based
      }
    }
    return null;
  } catch (_) {
    return null;
  }
}

/**
 * Sobrescribe los datos de una fila existente y limpia los campos
 * de figurita anterior para que se regenere desde cero.
 */
function actualizarFila_(rowIndex, datos) {
  const sheet = getSheet_();
  const colMap = getColumnMap_(sheet);

  const fechaSubida = Utilities.formatDate(
    new Date(),
    "America/Argentina/Buenos_Aires",
    "dd/MM/yyyy HH:mm:ss",
  );

  const celdas = [
    { col: CONFIG.COLUMNS.nombre, val: datos.nombre },
    { col: CONFIG.COLUMNS.mail, val: datos.mail },
    { col: CONFIG.COLUMNS.area, val: datos.area },
    { col: CONFIG.COLUMNS.superpoder, val: datos.superpoder },
    { col: CONFIG.COLUMNS.actitud, val: datos.actitud },
    {
      col: CONFIG.COLUMNS.consentimientoMural,
      val: datos.consentimientoMural || "",
    },
    { col: CONFIG.COLUMNS.idArchivoDrive, val: datos.idArchivoDrive },
    { col: CONFIG.COLUMNS.urlArchivoDrive, val: datos.urlArchivoDrive },
    { col: CONFIG.COLUMNS.estado, val: datos.estado },
    { col: CONFIG.COLUMNS.timestampSubida, val: fechaSubida },
    // Limpiar datos de la figurita anterior — se regenera
    { col: CONFIG.COLUMNS.idFigurita, val: "" },
    { col: CONFIG.COLUMNS.idFiguraGenerada, val: "" },
    { col: CONFIG.COLUMNS.urlFiguraGenerada, val: "" },
    { col: CONFIG.COLUMNS.detalleError, val: "" },
    { col: CONFIG.COLUMNS.timestampGeneracion, val: "" },
  ];

  celdas.forEach(({ col, val }) => {
    const colIdx = colMap[col];
    if (!colIdx) return;
    try {
      sheet.getRange(rowIndex, colIdx).setValue(val);
    } catch (err) {
      Logger.log(`[actualizarFila_] Error en columna "${col}": ${err.message}`);
    }
  });

  Logger.log(
    `[actualizarFila_] Fila ${rowIndex} sobrescrita: ${datos.nombre} <${datos.mail}>`,
  );
}

/**
 * Llamada por trigger desde el endpoint ?action=mural.
 * Llama a generarMural() sin UI (sin alerts).
 */
function generarMuralDesdeEndpoint_() {
  // Eliminar el trigger que disparó esta función
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "generarMuralDesdeEndpoint_")
    .forEach((t) => {
      try {
        ScriptApp.deleteTrigger(t);
      } catch (_) {}
    });

  try {
    generarMuralSilencioso_();
  } catch (err) {
    Logger.log("[generarMuralDesdeEndpoint_] ERROR: " + err.toString());
  }
}

function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
