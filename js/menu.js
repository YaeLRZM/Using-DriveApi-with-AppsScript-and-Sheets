function abrirInterfazImportacion() {
  const html = HtmlService.createHtmlOutputFromFile("interfaz")
    .setWidth(400)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, "Importar Bitácoras");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Bitácoras")
    .addItem("Importar datos", "abrirInterfazImportacion")
    .addToUi();
}
