function autorizarDrive() {
  const carpeta = DriveApp.getFolderById(CONFIG.FOLDER_FOTOS_ID);
  Logger.log('Acceso a Drive OK: ' + carpeta.getName());
  const archivo = DriveApp.getFileById(CONFIG.SLIDE_TEMPLATE_ID);
  Logger.log('Acceso a plantilla OK: ' + archivo.getName());
}