/**
 * Exportación a PDF de un libro de boletines ya generado y revisado (#21)
 *
 * Operación separada y posterior a la generación: no forma parte del pipeline que crea los boletines.
 * Lee el libro del curso sin modificarlo y guarda un PDF por hoja de estudiante en una carpeta nueva
 * junto al libro. Una falla de exportación no afecta al libro: se corrige la causa y se vuelve a exportar.
 *
 * Usa la URL de exportación de Google Sheets, un mecanismo no documentado oficialmente: queda aislado en
 * urlExportacionPdf_ y exportarHojaPdf_ para reemplazarlo si deja de funcionar.
 * Usa de Generador_EPS.gs: CONFIG_EPS, extraerSpreadsheetId_ y registrarEstado_.
 */
const CONFIG_PDF = {
  // Configuración de impresión validada (README), completa en cada exportación para no depender
  // de lo guardado en el diálogo de impresión. Márgenes en pulgadas: 1 cm = 0,3937 in.
  PARAMETROS: {
    format: 'pdf',
    size: 'A4',
    portrait: 'true',
    fitw: 'true',           // ajustar al ancho
    scale: '2',             // ajustar al ancho, en la forma más nueva del parámetro
    top_margin: '0',
    left_margin: '0',
    right_margin: '0',
    bottom_margin: '0.3937',
    fzr: 'true',            // repetir filas inmovilizadas: la fila 1 del boletín (ver completarInforme_)
    fzc: 'false',
    gridlines: 'false',
    printnotes: 'false',
    printtitle: 'false',
    sheetnames: 'false',
    pagenum: 'UNDEFINED',
    horizontal_alignment: 'CENTER',
    vertical_alignment: 'TOP'
  },
  // Margen ante el límite de 6 minutos por ejecución de Apps Script: pasado este tiempo no se empieza
  // otra exportación, y las hojas que faltan se informan como sin exportar.
  TIEMPO_MAXIMO_MS: 5 * 60 * 1000
};

function exportarPdfLibro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const control = ss.getSheetByName(CONFIG_EPS.HOJA_CONTROL);

  const respuesta = ui.prompt('Exportar boletines a PDF',
    'Link o ID del libro de boletines ya generado y revisado.\n\n' +
    'Se crea un PDF por hoja de estudiante en una carpeta nueva junto al libro. El libro no se modifica.',
    ui.ButtonSet.OK_CANCEL);
  // Comparación como en la documentación de Apps Script (los botones son enums del servicio).
  if (respuesta.getSelectedButton() != ui.Button.OK) return;

  try {
    const inicio = Date.now();
    const libro = SpreadsheetApp.openById(extraerSpreadsheetId_(respuesta.getResponseText()));
    const hojas = hojasEstudiantesLibro_(libro);
    const carpeta = crearCarpetaPdf_(libro);
    const resultados = exportarHojasPdf_(ss, libro, hojas, carpeta, inicio);

    const contar = estado => resultados.filter(r => r.estado === estado).length;
    const segundos = Math.round((Date.now() - inicio) / 1000);
    const exportados = contar('Exportado');
    const resultado = 'PDF exportados: ' + exportados + ' de ' + hojas.length + ' · Con error: ' + contar('Error') +
      ' · Sin exportar: ' + contar('Sin exportar') + ' · Duración: ' + segundos + ' s' +
      (exportados ? ' (' + Math.round(segundos / exportados * 10) / 10 + ' s por PDF)' : '') + '.';
    const detalle = resultados.filter(r => r.estado !== 'Exportado').map(r => r.estado + ' · ' + r.hoja + ': ' + r.detalle);

    registrarEstado_(control, resultado + ' Libro: ' + libro.getName() + ' · Carpeta: ' + carpeta.getUrl(), detalle);
    ui.alert(resultado + '\n\nLibro: ' + libro.getName() + '\nCarpeta: ' + carpeta.getName() + '\n' + carpeta.getUrl() +
      (detalle.length ? '\n\n- ' + detalle.join('\n- ') + '\n\nEl libro no se modificó: se puede volver a exportar.' : ''));
  } catch (err) {
    if (control) registrarEstado_(control, 'Error al exportar PDF: ' + err.message, []);
    ui.alert('No se pudo exportar el libro a PDF.\n\n' + err.message);
    throw err;
  }
}

// Hojas de estudiante del libro de un curso: todas menos Resumen, en el orden del libro.
// Sin hoja Resumen no es un libro de la generación por curso (por ejemplo, el propio generador).
function hojasEstudiantesLibro_(libro) {
  if (!libro.getSheetByName(CONFIG_EPS.HOJA_RESUMEN)) {
    throw new Error('El libro "' + libro.getName() + '" no tiene la hoja ' + CONFIG_EPS.HOJA_RESUMEN +
      ': no parece un libro de boletines generado por curso.');
  }
  const hojas = libro.getSheets().filter(hoja => hoja.getName() !== CONFIG_EPS.HOJA_RESUMEN);
  if (!hojas.length) throw new Error('El libro "' + libro.getName() + '" no tiene hojas de estudiantes.');
  return hojas;
}

// Carpeta nueva en la misma carpeta que el libro (o en Mi unidad si no se puede ver su carpeta).
// Cada exportación crea la suya: volver a exportar no reemplaza ni mezcla PDF anteriores.
function crearCarpetaPdf_(libro) {
  const padres = DriveApp.getFileById(libro.getId()).getParents();
  const padre = padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
  const fecha = Utilities.formatDate(new Date(), CONFIG_EPS.ZONA_HORARIA, 'yyyy-MM-dd HH:mm');
  return padre.createFolder(libro.getName() + ' - PDF ' + fecha);
}

// Una hoja por vez. El error de una hoja sólo afecta su PDF; las demás se exportan igual.
function exportarHojasPdf_(ss, libro, hojas, carpeta, inicio) {
  const token = ScriptApp.getOAuthToken();
  return hojas.map((hoja, i) => {
    const nombre = hoja.getName();
    if (Date.now() - inicio > CONFIG_PDF.TIEMPO_MAXIMO_MS) {
      return { hoja: nombre, estado: 'Sin exportar', detalle: 'se alcanzó el tiempo máximo de la ejecución.' };
    }
    ss.toast('PDF ' + (i + 1) + ' de ' + hojas.length + ': ' + nombre + '.', 'Informes EPS');
    try {
      carpeta.createFile(exportarHojaPdf_(libro.getId(), hoja, token));
      return { hoja: nombre, estado: 'Exportado', detalle: '' };
    } catch (err) {
      return { hoja: nombre, estado: 'Error', detalle: err.message };
    }
  });
}

function exportarHojaPdf_(libroId, hoja, token) {
  const respuesta = UrlFetchApp.fetch(urlExportacionPdf_(libroId, hoja.getSheetId()), {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  const codigo = respuesta.getResponseCode();
  if (codigo !== 200) {
    throw new Error('Google respondió HTTP ' + codigo + (codigo === 429 ? ' (demasiadas exportaciones seguidas)' : '') + '.');
  }
  const pdf = respuesta.getBlob();
  const tipo = String(pdf.getContentType());
  if (tipo.indexOf('application/pdf') !== 0) throw new Error('La respuesta de Google no es un PDF (' + tipo + ').');
  return pdf.setName(hoja.getName() + '.pdf');
}

function urlExportacionPdf_(libroId, gid) {
  const parametros = Object.assign({}, CONFIG_PDF.PARAMETROS, { gid: String(gid) });
  return 'https://docs.google.com/spreadsheets/d/' + libroId + '/export?' +
    Object.keys(parametros).map(clave => clave + '=' + encodeURIComponent(parametros[clave])).join('&');
}
