function autorizarDrive() {
  Logger.log('FOLDER_FOTOS_ID = "' + CONFIG.FOLDER_FOTOS_ID + '"');
  Logger.log('FOLDER_FIGURITAS_ID = "' + CONFIG.FOLDER_FIGURITAS_ID + '"');
  Logger.log('SLIDE_TEMPLATE_ID = "' + CONFIG.SLIDE_TEMPLATE_ID + '"');
  
  // Probar acceso genérico a Drive primero
  const root = DriveApp.getRootFolder();
  Logger.log('Root folder OK: ' + root.getName());
  
  // Ahora la carpeta específica
  const carpeta = DriveApp.getFolderById(CONFIG.FOLDER_FOTOS_ID);
  Logger.log('Carpeta fotos OK: ' + carpeta.getName());
}