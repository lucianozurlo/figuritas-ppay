function testExport() {
  Logger.log("=== TEST EXPORT ===");
  try {
    const carpeta = crearCarpetaVersion_();
    Logger.log("Carpeta creada: " + carpeta.getName());
    Logger.log("URL: " + carpeta.getUrl());

    const { pptxFile, pdfFile, pngFiles } = ejecutarExport_(carpeta);
    Logger.log("PPTX: " + pptxFile.getName());
    Logger.log("PDF: " + pdfFile.getName());
    Logger.log("PNGs: " + pngFiles.length);
  } catch (e) {
    Logger.log("ERROR: " + e.toString());
    Logger.log("STACK: " + (e.stack || "sin stack"));
  }
}
