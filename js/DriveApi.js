function extraerDatos(
  NOMBRE_EXCEL,
  HOJA_EXCEL,
  HOJA_DESTINO,
  FILA_INICIO_ORIGEN,
  FILA_INICIO_DESTINO,
) {
  FILA_INICIO_ORIGEN = Number(FILA_INICIO_ORIGEN);
  FILA_INICIO_DESTINO = Number(FILA_INICIO_DESTINO);

  const archivos = DriveApp.getFilesByName(NOMBRE_EXCEL);
  if (!archivos.hasNext()) {
    throw new Error("No se encontró el archivo Excel");
  }

  const excelFile = archivos.next();
  const recurso = {
    title: "TEMP_EXCEL_CONVERTIDO",
    mimeType: MimeType.GOOGLE_SHEETS,
  };

  const archivoConvertido = Drive.Files.insert(recurso, excelFile.getBlob());
  const archivoTempId = archivoConvertido.id;

  try {
    const ssTemp = SpreadsheetApp.openById(archivoTempId);
    const hoja = ssTemp.getSheetByName(HOJA_EXCEL);
    if (!hoja) throw new Error("No existe la hoja: " + HOJA_EXCEL);

    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < FILA_INICIO_ORIGEN) {
      throw new Error("No hay datos suficientes");
    }

    const numFilas = ultimaFila - FILA_INICIO_ORIGEN + 1;
    const datos = hoja.getRange(FILA_INICIO_ORIGEN, 1, numFilas, 9).getValues();

    //Modificar a las columnas que se necesiten traer A,B,C. Nota: Sheets usa indice 0 para A y asi sucesivamente
    const columnas = [0, 2, 6, 7, 8];

    let filtrado = datos
      .map((fila) => columnas.map((col) => fila[col]))
      .filter((fila) => fila[0] !== "" && fila[1] !== "");

    //ordenar por fecha y unidad
    filtrado.sort((a, b) => {
      const fechaA = new Date(a[0]);
      fechaA.setHours(0, 0, 0, 0);
      const fechaB = new Date(b[0]);
      fechaB.setHours(0, 0, 0, 0);

      if (fechaA.getTime() !== fechaB.getTime()) {
        return fechaA - fechaB;
      }

      const unidadA = parseInt(a[1], 10) || 0;
      const unidadB = parseInt(b[1], 10) || 0;
      return unidadA - unidadB;
    });

    const ssDestino = SpreadsheetApp.getActive();
    let destino = ssDestino.getSheetByName(HOJA_DESTINO);
    if (!destino) destino = ssDestino.insertSheet(HOJA_DESTINO);

    destino
      .getRange(FILA_INICIO_DESTINO, 1, destino.getMaxRows(), columnas.length)
      .clearContent();

    if (filtrado.length > 0) {
      destino
        .getRange(FILA_INICIO_DESTINO, 1, filtrado.length, filtrado[0].length)
        .setValues(filtrado);
    }

    Logger.log("Importación completada correctamente");
  } finally {
    DriveApp.getFileById(archivoTempId).setTrashed(true);
  }
}
