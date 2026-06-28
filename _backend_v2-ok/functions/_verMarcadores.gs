function verMarcadores() {
  const pres  = SlidesApp.openById(CONFIG.SLIDE_TEMPLATE_ID);
  const slide = pres.getSlides()[0];
  slide.getPageElements().forEach((el, i) => {
    try {
      const tipo = el.getPageElementType();
      let texto = '';
      if (tipo === SlidesApp.PageElementType.SHAPE) {
        texto = el.asShape().getText().asString().trim();
      }
      Logger.log(`[${i}] ${tipo} ${el.getTitle() ? '("'+el.getTitle()+'")' : ''} texto: "${texto}"`);
    } catch(e) {
      Logger.log(`[${i}] error: ${e.message}`);
    }
  });
}