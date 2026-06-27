function verPlantilla() {
  const pres = SlidesApp.openById(CONFIG.SLIDE_TEMPLATE_ID);
  const slide = pres.getSlides()[0];
  Logger.log("Total elementos: " + slide.getPageElements().length);
  slide.getPageElements().forEach((el, i) => {
    let titulo = "";
    try {
      titulo = el.getTitle() || "(sin titulo)";
    } catch (e) {
      titulo = "(error)";
    }
    let texto = "";
    try {
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        texto = el.asShape().getText().asString().trim().substring(0, 30);
      }
    } catch (e) {}
    Logger.log(
      `[${i}] tipo=${el.getPageElementType()} titulo="${titulo}" texto="${texto}"`,
    );
  });
}
