/**
 * Exportación a PDF de un libro de boletines ya generado y revisado (#21)
 *
 * Operación separada y posterior a la generación: no forma parte del pipeline que crea los boletines.
 * Lee el libro del curso sin modificarlo y guarda un PDF por hoja de estudiante en una carpeta nueva
 * junto al libro. Una falla de exportación no afecta al libro: se vuelve a exportar y, si la última exportación
 * quedó incompleta, se puede completar su carpeta con los PDF que faltan (exportacionIncompleta_).
 * Los pedidos se espacian y se reintentan ante el límite de Google (HTTP 429): ver CONFIG_PDF.PAUSA_MS.
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
  // Margen ante el límite de 6 minutos por ejecución de Apps Script, contado desde que empieza la ejecución
  // (los diálogos también cuentan): pasado este tiempo no se empieza otra exportación, y las hojas que faltan
  // se informan como sin exportar.
  TIEMPO_MAXIMO_MS: 5 * 60 * 1000,
  // La URL de exportación limita los pedidos seguidos (HTTP 429): en Google real, 15 pedidos sin pausa dieron
  // 6 PDF y 9 rechazos. Entre un pedido y el siguiente se espera una pausa que empieza en PAUSA_MS; ante un 429
  // se duplica, hasta PAUSA_MAXIMA_MS, y se reintenta la misma hoja, hasta INTENTOS_POR_HOJA pedidos en total;
  // cada PDF exportado la vuelve a bajar PAUSA_MS. Google no publica el límite: valores iniciales, a ajustar
  // con los tiempos y reintentos reales que registra Control.
  PAUSA_MS: 3 * 1000,
  PAUSA_MAXIMA_MS: 30 * 1000,
  INTENTOS_POR_HOJA: 4
};

function exportarPdfLibro() {
  const inicioEjecucion = Date.now();
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
    const libro = SpreadsheetApp.openById(extraerSpreadsheetId_(respuesta.getResponseText()));
    const hojas = hojasEstudiantesLibro_(libro);
    const padre = carpetaDelLibro_(libro);

    // Si la última exportación de este libro quedó incompleta, se ofrece completarla sin repetir lo ya exportado.
    let carpeta = null;
    let pendientes = hojas;
    const anterior = exportacionIncompleta_(padre, libro, hojas);
    if (anterior) {
      const opcion = ui.alert('Exportación incompleta',
        'La última exportación de este libro quedó incompleta: ' + anterior.carpeta.getName() + ' tiene ' +
        (hojas.length - anterior.faltan.length) + ' de ' + hojas.length + ' PDF.\n\n' +
        'Sí: completar esa carpeta, exportando sólo las ' + anterior.faltan.length + ' hojas que faltan.\n' +
        'No: exportar todas las hojas en una carpeta nueva (por ejemplo, si el libro se corrigió después).\n' +
        'Cancelar: no exportar nada.',
        ui.ButtonSet.YES_NO_CANCEL);
      if (opcion != ui.Button.YES && opcion != ui.Button.NO) return;
      if (opcion == ui.Button.YES) {
        carpeta = anterior.carpeta;
        pendientes = anterior.faltan;
      }
    }
    // Sin margen para el primer pedido (los diálogos estuvieron abiertos casi 5 minutos) no se crea una carpeta vacía.
    if (Date.now() - inicioEjecucion > CONFIG_PDF.TIEMPO_MAXIMO_MS) {
      throw new Error('No queda tiempo en esta ejecución para empezar la exportación: los diálogos estuvieron abiertos ' +
        'casi 5 minutos. No se creó ninguna carpeta; volver a exportar el libro.');
    }
    if (!carpeta) carpeta = crearCarpetaPdf_(padre, libro);

    const inicio = Date.now();
    const exportacion = exportarHojasPdf_(ss, libro, pendientes, carpeta, inicioEjecucion);
    const resultados = exportacion.resultados;

    const contar = estado => resultados.filter(r => r.estado === estado).length;
    const segundos = Math.round((Date.now() - inicio) / 1000);
    const exportados = contar('Exportado');
    const enCarpeta = hojas.length - pendientes.length + exportados;
    const resultado = 'PDF exportados: ' + exportados + ' de ' + pendientes.length +
      (pendientes.length < hojas.length ? ' que faltaban' : '') + ' · Con error: ' + contar('Error') +
      ' · Sin exportar: ' + contar('Sin exportar') + ' · Reintentos por HTTP 429: ' + exportacion.reintentos +
      ' · Duración: ' + segundos + ' s' +
      (exportados ? ' (' + Math.round(segundos / exportados * 10) / 10 + ' s por PDF)' : '') + '.' +
      ' La carpeta tiene ' + enCarpeta + ' de ' + hojas.length + ' PDF.';
    const detalle = resultados.filter(r => r.estado !== 'Exportado').map(r => r.estado + ' · ' + r.hoja + ': ' + r.detalle);

    registrarEstado_(control, resultado + ' Libro: ' + libro.getName() + ' · Carpeta: ' + carpeta.getUrl(), detalle);
    ui.alert(resultado + '\n\nLibro: ' + libro.getName() + '\nCarpeta: ' + carpeta.getName() + '\n' + carpeta.getUrl() +
      (detalle.length ? '\n\n- ' + detalle.join('\n- ') : '') +
      (enCarpeta < hojas.length ? '\n\nEl libro no se modificó. Para completar la carpeta, volver a exportar el mismo libro ' +
        'y elegir Sí: se exportan sólo los PDF que faltan.' : ''));
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

// La carpeta del libro, donde van sus carpetas de PDF (o Mi unidad si no se puede ver su carpeta).
function carpetaDelLibro_(libro) {
  const padres = DriveApp.getFileById(libro.getId()).getParents();
  return padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
}

// Carpeta nueva para cada exportación completa: volver a exportar todo no reemplaza ni mezcla PDF anteriores.
// Su descripción lleva el ID del libro: dos libros pueden llamarse igual (dos regeneraciones del mismo minuto).
function crearCarpetaPdf_(padre, libro) {
  const fecha = Utilities.formatDate(new Date(), CONFIG_EPS.ZONA_HORARIA, 'yyyy-MM-dd HH:mm');
  return padre.createFolder(libro.getName() + ' - PDF ' + fecha).setDescription(descripcionCarpetaPdf_(libro));
}

function descripcionCarpetaPdf_(libro) {
  return 'PDF del libro de boletines ' + libro.getId();
}

// La última carpeta de PDF de este libro (la creada más recientemente, fuera de la papelera, con el ID del libro en
// la descripción; no se mira el nombre) y las hojas cuyo PDF no tiene. null si no hay ninguna o si la última está
// completa: entonces se exporta todo en una carpeta nueva.
function exportacionIncompleta_(padre, libro, hojas) {
  const descripcion = descripcionCarpetaPdf_(libro);
  let ultima = null;
  const carpetas = padre.getFolders();
  while (carpetas.hasNext()) {
    const c = carpetas.next();
    if (!c.isTrashed() && c.getDescription() === descripcion && (!ultima || c.getDateCreated() > ultima.getDateCreated())) ultima = c;
  }
  if (!ultima) return null;

  const existentes = new Set();
  const archivos = ultima.getFiles();
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (!archivo.isTrashed()) existentes.add(archivo.getName());
  }
  const faltan = hojas.filter(hoja => !existentes.has(nombrePdf_(hoja)));
  return faltan.length ? { carpeta: ultima, faltan } : null;
}

function nombrePdf_(hoja) {
  return hoja.getName() + '.pdf';
}

/**
 * Una hoja por vez, con una pausa entre pedidos que se adapta al límite de Google (ver CONFIG_PDF.PAUSA_MS): ante un
 * HTTP 429 se duplica y se reintenta la misma hoja; cada PDF exportado la baja un paso. Otro error, o el 429 en el
 * último intento, sólo afecta el PDF de esa hoja; las demás se exportan igual.
 */
function exportarHojasPdf_(ss, libro, hojas, carpeta, inicioEjecucion) {
  const token = ScriptApp.getOAuthToken();
  let pausa = CONFIG_PDF.PAUSA_MS;
  let reintentos = 0;
  let primero = true;

  const resultados = hojas.map((hoja, i) => {
    const nombre = hoja.getName();
    ss.toast('PDF ' + (i + 1) + ' de ' + hojas.length + ': ' + nombre + '.', 'Informes EPS');
    for (let intento = 1; ; intento++) {
      const espera = primero ? 0 : pausa;
      if (Date.now() + espera - inicioEjecucion > CONFIG_PDF.TIEMPO_MAXIMO_MS) {
        return { hoja: nombre, estado: 'Sin exportar', detalle: 'se alcanzó el tiempo máximo de la ejecución.' };
      }
      if (espera) Utilities.sleep(espera);
      primero = false;
      try {
        carpeta.createFile(exportarHojaPdf_(libro.getId(), hoja, token));
        pausa = Math.max(CONFIG_PDF.PAUSA_MS, pausa - CONFIG_PDF.PAUSA_MS);
        return { hoja: nombre, estado: 'Exportado', detalle: '' };
      } catch (err) {
        if (!err.limiteGoogle || intento === CONFIG_PDF.INTENTOS_POR_HOJA) {
          return { hoja: nombre, estado: 'Error', detalle: err.message + (intento > 1 ? ' (' + intento + ' intentos)' : '') };
        }
        reintentos++;
        pausa = Math.min(pausa * 2, CONFIG_PDF.PAUSA_MAXIMA_MS);
        ss.toast('Google limitó las exportaciones: se espera ' + pausa / 1000 + ' s y se reintenta ' + nombre + '.', 'Informes EPS');
      }
    }
  });
  return { resultados, reintentos };
}

function exportarHojaPdf_(libroId, hoja, token) {
  const respuesta = UrlFetchApp.fetch(urlExportacionPdf_(libroId, hoja.getSheetId()), {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  const codigo = respuesta.getResponseCode();
  if (codigo !== 200) {
    const error = new Error('Google respondió HTTP ' + codigo + (codigo === 429 ? ' (demasiadas exportaciones seguidas)' : '') + '.');
    error.limiteGoogle = codigo === 429;
    throw error;
  }
  const pdf = respuesta.getBlob();
  const tipo = String(pdf.getContentType());
  if (tipo.indexOf('application/pdf') !== 0) throw new Error('La respuesta de Google no es un PDF (' + tipo + ').');
  return pdf.setName(nombrePdf_(hoja));
}

function urlExportacionPdf_(libroId, gid) {
  const parametros = Object.assign({}, CONFIG_PDF.PARAMETROS, { gid: String(gid) });
  return 'https://docs.google.com/spreadsheets/d/' + libroId + '/export?' +
    Object.keys(parametros).map(clave => clave + '=' + encodeURIComponent(parametros[clave])).join('&');
}
