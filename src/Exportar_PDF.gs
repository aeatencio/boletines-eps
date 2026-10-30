/**
 * Exportación a PDF de un libro de boletines ya generado y revisado (#21)
 *
 * Operación separada y posterior a la generación: no forma parte del pipeline que crea los boletines.
 * Lee el libro del curso sin modificarlo y guarda un PDF por hoja de estudiante en una carpeta nueva
 * junto al libro. Una falla de exportación no afecta al libro: se vuelve a exportar y, si la última exportación
 * quedó incompleta, se puede completar su carpeta con los PDF que faltan (ultimaExportacion_).
 * Los pedidos se espacian y se reintentan ante el límite de Google (HTTP 429): ver CONFIG_PDF.PAUSA_MS.
 * Varios libros (los cursos de Control, o una carpeta elegida a mano) se exportan con la misma lógica, el libro más
 * reciente de cada curso, en tantas ejecuciones como haga falta: cada una continúa desde lo que ya está en Drive
 * (confirmacionPlanesPdf_, ejecutarPlanesPdf_ y resultadoPlanesPdf_). Un libro elegido a mano usa el mismo plan por
 * libro (planLibro_). La exportación desde un libro o carpeta usa un diálogo HTML (Exportar_PDF_UI.html) para que la
 * espera de la persona no ocurra dentro de la ejecución que exporta (exportarPdfDesdeDrive).
 *
 * Usa la URL de exportación de Google Sheets, un mecanismo no documentado oficialmente: queda aislado en
 * urlExportacionPdf_ y exportarHojaPdf_ para reemplazarlo si deja de funcionar.
 * Usa de Generador_EPS.gs: CONFIG_EPS, texto_, registrarEstado_ y, para varios cursos, leerCursosMarcados_,
 * normalizarPeriodo_, leerCicloLectivo_, prefijoLibroCurso_, nombresCarpetaLibros_, rutaCarpetaLibros_ y buscarSubcarpeta_.
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

/**
 * Exportar PDF desde un libro o carpeta: operación manual, independiente de Control, en un diálogo HTML
 * (Exportar_PDF_UI.html). La interacción y la exportación son ejecuciones distintas de Apps Script, para que el tiempo que
 * se tarda en buscar el link o en revisar lo que se va a hacer no se descuente del tiempo máximo de la exportación:
 * el menú sólo abre el diálogo; "Analizar" llama a previsualizarExportacionPdf, que no exporta; "Exportar" llama a
 * exportarPdfConfirmado, una ejecución nueva que vuelve a mirar Drive antes de actuar. No hay estado guardado entre
 * ellas: el diálogo devuelve el link y la opción elegida.
 * Un libro: se exporta ese libro. Una carpeta: el libro más reciente de cada curso que contenga.
 * Los cursos marcados o todos, desde Control, están en exportarPdfVariosCursos_.
 */
function exportarPdfDesdeDrive() {
  const html = HtmlService.createHtmlOutputFromFile('Exportar_PDF_UI').setWidth(680).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Exportar PDF desde un libro o carpeta');
}

/**
 * Vista previa, desde el diálogo: qué se encontró y qué se haría. Sólo lee: no exporta, no crea carpetas ni archivos,
 * no modifica libros y no escribe en Control, en ningún camino. Sin nada pendiente, lo dice y no ofrece opciones; si el
 * análisis falla, el error llega al diálogo sin registrarse en Control (todavía no hubo un intento de exportación).
 * Devuelve { entrada, texto, opciones: [{ modo, etiqueta }], firma }. firma resume el plan, para que la exportación
 * avise si encontró otro estado en Drive.
 */
function previsualizarExportacionPdf(entrada) {
  const elegido = elementoDrive_(entrada);
  if (elegido.carpeta) {
    const carpeta = planCarpetaPdf_(elegido.carpeta);
    const pendientes = carpeta.planes.reduce((n, p) => n + p.faltan.length, 0);
    return {
      entrada, texto: confirmacionPlanesPdf_(carpeta.planes, carpeta.contexto), firma: firmaPlanesPdf_(carpeta.planes),
      opciones: pendientes ? [{ modo: 'carpeta', etiqueta: 'Exportar ' + pendientes + ' PDF' }] : []
    };
  }
  const plan = planLibro_({}, elegido.libro, carpetaDelLibro_(elegido.libro));
  const opciones = plan.carpeta && plan.faltan.length
    ? [{ modo: 'completar', etiqueta: 'Completar (' + plan.faltan.length + ' PDF)' },
      { modo: 'rehacer', etiqueta: 'Exportar todo en una carpeta nueva (' + plan.hojas.length + ' PDF)' }]
    : [{ modo: 'rehacer', etiqueta: 'Exportar ' + plan.hojas.length + ' PDF' }];
  return { entrada, texto: vistaLibroPdf_(plan), firma: vistaLibroPdf_(plan), opciones };
}

/**
 * Exportación, desde el diálogo, en una ejecución propia: el tiempo máximo se cuenta desde acá. No confía en la vista
 * previa: vuelve a resolver la entrada y a calcular el plan con el estado actual de Drive, y si no es el que se vio,
 * lo avisa. modo: 'completar' o 'rehacer' para un libro (exportarLibroPdf_); una carpeta no tiene opciones. Una
 * exportación por vez y por usuario (LockService), para no duplicar PDF con dos diálogos abiertos.
 * Devuelve { texto } con el resultado, que también queda en Control.
 */
function exportarPdfConfirmado(entrada, modo, firma) {
  const inicioEjecucion = Date.now();
  return conRegistroPdf_(() => {
    const cerrojo = LockService.getUserLock();
    if (!cerrojo.tryLock(0)) throw new Error('Ya hay una exportación de PDF en curso: esperar a que termine y volver a analizar.');
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const elegido = elementoDrive_(entrada);
      let actual, final;
      if (elegido.carpeta) {
        const carpeta = planCarpetaPdf_(elegido.carpeta);
        actual = firmaPlanesPdf_(carpeta.planes);
        final = resultadoPlanesPdf_(carpeta.planes, carpeta.contexto, ejecutarPlanesPdf_(ss, carpeta.planes, inicioEjecucion));
      } else {
        const plan = planLibro_({}, elegido.libro, carpetaDelLibro_(elegido.libro));
        actual = vistaLibroPdf_(plan);
        final = exportarLibroPdf_(ss, plan, modo, inicioEjecucion);
      }
      registrarEstado_(controlPdf_(), final.resultado, final.detalle);
      return { texto: (actual !== firma ? 'Lo que hay en Drive cambió desde la vista previa: se usó el estado actual.\n\n' : '') + final.texto };
    } finally {
      cerrojo.releaseLock();
    }
  });
}

function controlPdf_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_EPS.HOJA_CONTROL);
}

// Un error de la exportación confirmada queda también en Control, y llega al diálogo para mostrarlo.
function conRegistroPdf_(operacion) {
  try {
    return operacion();
  } catch (err) {
    const control = controlPdf_();
    if (control) registrarEstado_(control, 'Error al exportar PDF: ' + err.message, []);
    throw err;
  }
}

/**
 * Libro (planilla de Google) o carpeta de Drive a partir de un link o ID, sólo con operaciones documentadas:
 * SpreadsheetApp.openById abre planillas y DriveApp.getFolderById, carpetas; cada una falla si el ID no es de su tipo.
 * Un link de carpeta (/folders/) se busca sólo como carpeta y uno de archivo (/d/), sólo como planilla; un ID solo
 * (o ?id=), primero como planilla y después como carpeta. DriveApp.getFileById sólo se usa, sobre archivos, para
 * nombrar en el error un archivo que no es una planilla.
 */
function elementoDrive_(entrada) {
  const ref = referenciaDrive_(entrada);
  if (ref.tipo !== 'carpeta') {
    try {
      return { libro: SpreadsheetApp.openById(ref.id) };
    } catch (err) { /* no es una planilla accesible: se prueba como carpeta, o se informa abajo */ }
  }
  if (ref.tipo !== 'libro') {
    try {
      return { carpeta: DriveApp.getFolderById(ref.id) };
    } catch (err) { /* no es una carpeta accesible */ }
  }
  if (ref.tipo !== 'carpeta') {
    let nombre = '';
    try {
      nombre = DriveApp.getFileById(ref.id).getName();
    } catch (err) { /* tampoco es un archivo accesible */ }
    if (nombre) throw new Error('"' + nombre + '" no es un libro de Google Sheets ni una carpeta de Drive.');
  }
  throw new Error('No se encontró en Drive ' + (ref.tipo === 'carpeta' ? 'una carpeta' : ref.tipo === 'libro' ? 'un libro' :
    'un libro ni una carpeta') + ' con ese link o ID, o no hay acceso.');
}

// ID y, según la forma del link, qué se espera: carpeta (/folders/<id>), libro (/d/<id>) o cualquiera (?id=<id>, ID solo).
function referenciaDrive_(entrada) {
  const texto = texto_(entrada);
  let match = texto.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (match) return { id: match[1], tipo: 'carpeta' };
  match = texto.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return { id: match[1], tipo: 'libro' };
  match = texto.match(/[?&]id=([a-zA-Z0-9_-]+)/) || texto.match(/^([a-zA-Z0-9_-]{20,})$/);
  if (match) return { id: match[1], tipo: null };
  throw new Error('No pude reconocer el ID. Pegá el link completo o el ID de un libro de boletines o de una carpeta de Drive.');
}

// Qué se haría con un libro elegido a mano (vista previa). Elegirlo a mano es una decisión deliberada de exportarlo:
// si su última exportación está completa, o el libro cambió después, se exporta todo en una carpeta nueva; si quedó
// incompleta y vigente, se elige entre completarla o exportar todo.
function vistaLibroPdf_(plan) {
  const total = plan.hojas.length;
  const libro = 'Libro: ' + plan.libro.getName() + ' (' + total + ' hojas de estudiantes).\n\n';
  if (plan.carpeta && plan.faltan.length) {
    return libro + 'La última exportación de este libro quedó incompleta: ' + plan.carpeta.getName() + ' tiene ' +
      plan.yaHabia + ' de ' + total + ' PDF.\n\n' +
      'Completar: completar esa carpeta, exportando sólo las ' + plan.faltan.length + ' hojas que faltan.\n' +
      'Exportar todo: exportar todas las hojas en una carpeta nueva.';
  }
  if (plan.carpeta) return libro + 'Su última exportación está completa (' + plan.carpeta.getName() + '): se exportan de nuevo sus ' + total + ' PDF en una carpeta nueva.';
  if (plan.desactualizada) return libro + 'El libro cambió después de su última exportación (' + plan.desactualizada.getName() + '): se exportan sus ' + total + ' PDF en una carpeta nueva.';
  return libro + 'Se exportan sus ' + total + ' PDF en una carpeta nueva junto al libro.';
}

/**
 * Exporta un libro según lo elegido y el estado actual de Drive (plan): 'completar' completa su carpeta si sigue
 * vigente (si ya está completa, no exporta nada); en cualquier otro caso, o si la carpeta dejó de ser vigente, se
 * exporta todo en una carpeta nueva. Devuelve resultado (Control!F3), detalle (F4) y texto.
 */
function exportarLibroPdf_(ss, plan, modo, inicioEjecucion) {
  const libro = plan.libro;
  const hojas = plan.hojas;
  if (!(modo === 'completar' && plan.carpeta)) Object.assign(plan, { carpeta: null, faltan: hojas, yaHabia: 0 });
  // Sin margen para el primer pedido no se crea una carpeta vacía.
  if (plan.faltan.length && !hayTiempoPdf_(inicioEjecucion, nuevoRitmoPdf_())) {
    throw new Error('No queda tiempo en esta ejecución para empezar la exportación. No se creó ninguna carpeta; volver a intentar.');
  }

  const inicio = Date.now();
  const ritmo = nuevoRitmoPdf_();
  exportarPlanPdf_(ss, plan, inicioEjecucion, ritmo);
  const resultados = plan.resultados;
  const carpeta = plan.carpeta;

  const contar = estado => resultados.filter(r => r.estado === estado).length;
  const segundos = Math.round((Date.now() - inicio) / 1000);
  const exportados = contar('Exportado');
  const enCarpeta = pdfEnCarpeta_(plan);
  const resultado = 'PDF exportados: ' + exportados + ' de ' + plan.faltan.length +
    (plan.faltan.length < hojas.length ? ' que faltaban' : '') + ' · Con error: ' + contar('Error') +
    ' · Sin exportar: ' + contar('Sin exportar') + ' · Reintentos por HTTP 429: ' + ritmo.reintentos +
    ' · Tiempo de exportación: ' + segundos + ' s' +
    (exportados ? ' (' + Math.round(segundos / exportados * 10) / 10 + ' s por PDF)' : '') + '.' +
    ' La carpeta tiene ' + enCarpeta + ' de ' + hojas.length + ' PDF.';
  const detalle = resultados.filter(r => r.estado !== 'Exportado').map(r => r.estado + ' · ' + r.hoja + ': ' + r.detalle);
  return {
    resultado: resultado + ' Libro: ' + libro.getName() + ' · Carpeta: ' + carpeta.getUrl(),
    detalle,
    texto: resultado + '\n\nLibro: ' + libro.getName() + '\nCarpeta: ' + carpeta.getName() + '\n' + carpeta.getUrl() +
      (detalle.length ? '\n\n- ' + detalle.join('\n- ') : '') +
      (enCarpeta < hojas.length ? '\n\nEl libro no se modificó. Para completar la carpeta, volver a analizar el mismo libro ' +
        'y elegir Completar: se exportan sólo los PDF que faltan.' : '')
  };
}

/**
 * Una carpeta elegida a mano: el libro más reciente de cada curso que contenga (librosPorCursoEnCarpeta_), con la misma
 * lógica que los cursos de Control, pero sin depender de Control: el curso, el período y el ciclo salen de cada libro.
 * Devuelve los planes y el contexto de exportarPlanesPdf_.
 */
function planCarpetaPdf_(carpeta) {
  const encontrados = librosPorCursoEnCarpeta_(carpeta);
  const planes = encontrados.grupos.map(grupo => planGrupoPdf_(grupo.etiqueta, grupo, carpeta))
    .concat(encontrados.sinAbrir.map(s => ({ curso: 'Planilla "' + s.archivo.getName() + '"', libros: 1, fecha: fechaPdf_(s.archivo.getDateCreated()),
      libro: null, hojas: null, carpeta: null, faltan: [], yaHabia: 0, error: 'no se pudo abrir: ' + s.error })));
  const anteriores = encontrados.grupos.reduce((n, g) => n + g.libros.length - 1, 0);
  const nombre = carpeta.getName();
  return {
    planes,
    contexto: {
      titulo: 'PDF de la carpeta "' + nombre + '"',
      intro: 'Carpeta: ' + nombre + '.\nLibros de boletines: ' + encontrados.libros + ', de ' + encontrados.grupos.length +
        ' cursos. Se usa el más reciente de cada curso; se omiten ' + anteriores + ' regeneraciones anteriores.' +
        (encontrados.ignoradas ? ' Otras planillas, que no son libros de boletines: ' + encontrados.ignoradas + ', se ignoran.' : ''),
      donde: 'la carpeta "' + nombre + '"',
      repetir: 'volver a analizar y exportar la misma carpeta',
      sinLibros: 'La carpeta no tiene libros de boletines.'
    }
  };
}

function firmaPlanesPdf_(planes) {
  return planes.map(p => lineaCursoPdf_(p, false)).join('\n');
}

/**
 * Exportación de varios cursos desde Control (#23): los cursos marcados, o los 8, del período y ciclo de Control, con su
 * libro más reciente en Mi unidad > Boletines EPS > <ciclo> - <período> (los libros se reconocen igual que en una
 * carpeta elegida a mano: librosPorCursoEnCarpeta_). Un curso sin libro queda informado.
 */
function exportarPdfCursosMarcados() {
  exportarPdfVariosCursos_(false);
}

function exportarPdfTodosLosCursos() {
  exportarPdfVariosCursos_(true);
}

function exportarPdfVariosCursos_(todos) {
  const inicioEjecucion = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const control = ss.getSheetByName(CONFIG_EPS.HOJA_CONTROL);

  try {
    if (!control) throw new Error('No se encontró la hoja Control.');
    const config = leerPeriodoExportacion_(control);
    const cursos = todos ? CONFIG_EPS.CURSOS
      : leerCursosMarcados_(control, 'Para exportar los 8, usar Informes EPS > Exportar PDF de todos los cursos.');
    const ruta = rutaCarpetaLibros_(config);
    const carpetaPeriodo = nombresCarpetaLibros_(config)
      .reduce((padre, nombre) => padre && buscarSubcarpeta_(padre, nombre), DriveApp.getRootFolder());
    const grupos = carpetaPeriodo ? librosPorCursoEnCarpeta_(carpetaPeriodo).grupos : [];
    const planes = cursos.map(curso => {
      const clave = cursoPdf_(curso.hojaOrigen, curso.nivel, config.periodo, config.cicloLectivo).clave;
      return planGrupoPdf_(curso.hojaOrigen + ' ' + curso.nivel, grupos.find(g => g.clave === clave), carpetaPeriodo);
    });
    exportarPlanesPdf_(ss, ui, control, planes, inicioEjecucion, {
      titulo: 'PDF de varios cursos (' + config.periodo + ' ' + config.cicloLectivo + ')',
      intro: 'Período: ' + config.periodo + ' ' + config.cicloLectivo + ' · Libros en ' + ruta + '.',
      donde: ruta,
      repetir: 'volver a ejecutar la misma opción',
      sinLibros: 'Ningún curso elegido tiene libro en ' + ruta + '.'
    });
  } catch (err) {
    if (control) registrarEstado_(control, 'Error al exportar PDF de varios cursos: ' + err.message, []);
    ui.alert('No se pudieron exportar los PDF de los cursos.\n\n' + err.message);
    throw err;
  }
}

/**
 * Varios libros (de una carpeta o de los cursos de Control), uno por plan, con un único ritmo de pausas: el límite de
 * Google es por usuario. Se confirma lo que se va a hacer y cada ejecución avanza hasta el tiempo máximo; la siguiente
 * continúa desde lo que ya está en Drive (las carpetas de PDF de cada libro, ultimaExportacion_), sin estado propio,
 * triggers ni colas. Por plan: carpeta completa y vigente, se omite; incompleta y vigente, se completa; sin carpeta
 * vigente (no había, o el libro cambió después), se exporta todo en una nueva; libro con problema, se informa.
 * contexto: titulo (para Control), intro (para la confirmación), donde (dónde están los libros), repetir (cómo
 * continuar) y sinLibros (si ningún plan tiene libro).
 */
function exportarPlanesPdf_(ss, ui, control, planes, inicioEjecucion, contexto) {
  // Sin nada para exportar no se pregunta ni se pide nada, pero el resultado se informa igual: un curso con un
  // libro que no se puede exportar sigue siendo un problema pendiente.
  let ejecucion = { ritmo: nuevoRitmoPdf_(), inicio: Date.now() };
  if (planes.some(p => p.faltan.length)) {
    const confirmacion = ui.alert('Exportar PDF de varios libros', confirmacionPlanesPdf_(planes, contexto), ui.ButtonSet.OK_CANCEL);
    if (confirmacion != ui.Button.OK) return;
    ejecucion = ejecutarPlanesPdf_(ss, planes, inicioEjecucion);
  }
  const final = resultadoPlanesPdf_(planes, contexto, ejecucion);
  registrarEstado_(control, final.resultado, final.detalle);
  ui.alert(final.texto);
}

// Lo que se va a hacer con cada libro, para confirmarlo antes de exportar. Sin PDF pendientes (sólo en la vista previa
// de una carpeta: Control no confirma en ese caso), lo dice en lugar de pedir confirmación.
function confirmacionPlanesPdf_(planes, contexto) {
  const cuantos = condicion => planes.filter(condicion).length;
  const pendientes = planes.reduce((n, p) => n + p.faltan.length, 0);
  return contexto.intro + '\n' +
    'Completos, se omiten: ' + cuantos(p => p.hojas && !p.faltan.length) +
    ' · Se completan: ' + cuantos(p => p.yaHabia && p.faltan.length) +
    ' · Se rehacen porque el libro cambió: ' + cuantos(p => p.desactualizada) +
    ' · Se exportan enteros: ' + cuantos(p => p.faltan.length && !p.yaHabia && !p.desactualizada) +
    ' · Con problema: ' + cuantos(p => p.error) + '.' +
    (planes.length ? '\n\n- ' + planes.map(p => lineaCursoPdf_(p, false)).join('\n- ') : '') + '\n\n' +
    (pendientes
      ? 'PDF a exportar: ' + pendientes + '. Cada ejecución avanza hasta ' + CONFIG_PDF.TIEMPO_MAXIMO_MS / 60000 +
        ' minutos; si no termina, ' + contexto.repetir + ': continúa donde quedó, sin repetir los PDF ya exportados. ' +
        'Los libros no se modifican.'
      : 'No hay PDF para exportar' + (planes.some(p => p.libros) ? '.' : ': ' + contexto.sinLibros));
}

// Exporta los libros con PDF pendientes, en orden y con un único ritmo, hasta el tiempo máximo de esta ejecución.
function ejecutarPlanesPdf_(ss, planes, inicioEjecucion) {
  const pendientes = planes.filter(p => p.faltan.length);
  const ejecucion = { ritmo: nuevoRitmoPdf_(), inicio: Date.now() };
  pendientes.forEach((plan, i) => {
    // Sin margen para el primer pedido de este libro, no se empieza ni se crea su carpeta: queda para la próxima.
    if (!hayTiempoPdf_(inicioEjecucion, ejecucion.ritmo)) return;
    ss.toast('Libro ' + (i + 1) + ' de ' + pendientes.length + ': ' + plan.curso + '.', 'Informes EPS');
    exportarPlanPdf_(ss, plan, inicioEjecucion, ejecucion.ritmo);
  });
  return ejecucion;
}

// Resultado de varios libros: resultado (Control!F3), detalle (F4) y texto (para mostrar).
function resultadoPlanesPdf_(planes, contexto, ejecucion) {
  const todosResultados = [].concat(...planes.map(p => p.resultados || []));
  const contar = estado => todosResultados.filter(r => r.estado === estado).length;
  // Con libro: todo curso para el que se encontró un libro, se pueda exportar o no.
  const conLibro = planes.filter(p => p.libros);
  const completos = conLibro.filter(p => p.hojas && pdfEnCarpeta_(p) === p.hojas.length).length;
  const conProblema = conLibro.filter(p => p.error).length;
  const faltanPdf = conLibro.length - completos - conProblema;
  const resultado = contexto.titulo + ': cursos completos: ' +
    completos + ' de ' + conLibro.length + ' con libro' + (conProblema ? ' · Con libro que no se puede exportar: ' + conProblema : '') +
    ' · PDF exportados ahora: ' + contar('Exportado') + ' · Con error: ' + contar('Error') +
    ' · Reintentos por HTTP 429: ' + ejecucion.ritmo.reintentos + ' · Tiempo de exportación: ' + Math.round((Date.now() - ejecucion.inicio) / 1000) + ' s.' +
    (faltanPdf ? ' Faltan PDF: ' + contexto.repetir + ' para continuar.' : '') +
    (conProblema ? ' Hay cursos cuyo libro no se puede exportar: ver el detalle, corregirlo y volver a ejecutar.' : '') +
    (conLibro.length ? '' : ' ' + contexto.sinLibros);
  const detalle = [];
  planes.forEach(p => {
    detalle.push(lineaCursoPdf_(p, true) + (p.carpeta ? ' ' + p.carpeta.getUrl() : ''));
    (p.resultados || []).filter(r => r.estado === 'Error').forEach(r => detalle.push('  ' + p.curso + ' · Error · ' + r.hoja + ': ' + r.detalle));
  });
  const texto = resultado + (planes.length ? '\n\n- ' + planes.map(p => lineaCursoPdf_(p, true)).join('\n- ') : '') +
    '\n\nLas carpetas de PDF están junto a cada libro, en ' + contexto.donde + '. El detalle, con los enlaces, queda en ' +
    CONFIG_EPS.HOJA_CONTROL + '.';
  return { resultado, detalle, texto };
}

// Período y ciclo de Control: lo único que necesita la exportación de varios cursos (no lee las grillas).
function leerPeriodoExportacion_(control) {
  const periodo = normalizarPeriodo_(control.getRange(CONFIG_EPS.CELDA_PERIODO).getDisplayValue());
  const periodos = Object.keys(CONFIG_EPS.PERIODOS).map(clave => CONFIG_EPS.PERIODOS[clave]);
  if (periodos.indexOf(periodo) < 0) {
    throw new Error('Falta indicar el período en ' + CONFIG_EPS.HOJA_CONTROL + '!' + CONFIG_EPS.CELDA_PERIODO +
      ', o no está implementado: "' + periodo + '".');
  }
  return { periodo, cicloLectivo: leerCicloLectivo_(control) };
}

/**
 * Libros de boletines de una carpeta, agrupados por curso: hoja origen, nivel, período y ciclo, que la generación
 * escribe en Resumen (A2:B5) al terminar. Un libro cuya generación no terminó no tiene esos datos y se agrupa por el
 * nombre que le dio la generación (prefijoLibroCurso_); si tampoco se reconoce por el nombre pero tiene hoja Resumen,
 * queda solo, como problema. Las demás planillas no son libros de boletines y se ignoran (se cuentan).
 * Una planilla que no puede abrirse también se agrupa por ese nombre, si lo tiene; si no, se devuelve aparte
 * (sinAbrir), para informarla en la operación manual: no hay forma fundada de atribuirla a un curso.
 * De cada curso se elige el candidato creado más recientemente, sea un libro terminado, uno incompleto o uno que no
 * pudo abrirse: en los dos últimos casos el curso queda con un problema (planGrupoPdf_), sin volver en silencio a
 * uno anterior.
 */
function librosPorCursoEnCarpeta_(carpeta) {
  const grupos = {};
  const encontrados = { grupos: [], libros: 0, ignoradas: 0, sinAbrir: [] };
  const archivos = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (archivo.isTrashed()) continue;
    // Candidato: { archivo, libro } abierto, o { archivo, libro: null, error } si no pudo abrirse.
    const candidato = { archivo, libro: null, error: '' };
    let curso;
    try {
      candidato.libro = SpreadsheetApp.openById(archivo.getId());
      curso = cursoDelLibro_(candidato.libro, archivo.getName());
    } catch (err) {
      candidato.error = err.message;
      curso = cursoPorNombre_(archivo.getName());
      if (!curso) {
        encontrados.sinAbrir.push(candidato);
        continue;
      }
    }
    if (!curso) {
      encontrados.ignoradas++;
      continue;
    }
    encontrados.libros++;
    if (!grupos[curso.clave]) grupos[curso.clave] = Object.assign({ libros: [] }, curso);
    grupos[curso.clave].libros.push(candidato);
  }

  const periodos = Object.keys(CONFIG_EPS.PERIODOS).map(clave => CONFIG_EPS.PERIODOS[clave]);
  const posicion = g => {
    const i = CONFIG_EPS.CURSOS.findIndex(c => c.hojaOrigen === g.hojaOrigen && c.nivel === g.nivel);
    return i < 0 ? CONFIG_EPS.CURSOS.length : i;
  };
  encontrados.grupos = Object.keys(grupos).map(clave => grupos[clave])
    .sort((a, b) => String(a.cicloLectivo).localeCompare(String(b.cicloLectivo)) ||
      periodos.indexOf(a.periodo) - periodos.indexOf(b.periodo) || posicion(a) - posicion(b) || a.etiqueta.localeCompare(b.etiqueta));
  encontrados.grupos.forEach(g => {
    g.elegido = g.libros.reduce((a, b) => (b.archivo.getDateCreated() > a.archivo.getDateCreated() ? b : a));
  });
  return encontrados;
}

// Curso de un libro de boletines, o null si la planilla no es uno. Del Resumen terminado; si no, del nombre que le da
// la generación (cursoPorNombre_); si tampoco, el libro solo.
function cursoDelLibro_(libro, nombre) {
  const resumen = libro.getSheetByName(CONFIG_EPS.HOJA_RESUMEN);
  if (resumen) {
    const v = resumen.getRange('A1:B5').getDisplayValues();
    if (v[0][0] === CONFIG_EPS.TITULO_RESUMEN && v[1][1] && v[2][1] && v[3][1]) return cursoPdf_(v[1][1], v[2][1], v[3][1], v[4][1]);
  }
  return cursoPorNombre_(nombre) || (resumen ? { clave: 'libro ' + libro.getId(), etiqueta: 'Libro "' + nombre + '"' } : null);
}

// Curso según el nombre que le da la generación ("Boletines EPS - <hoja> - <nivel> - <período> <ciclo> - <fecha>"), o null.
function cursoPorNombre_(nombre) {
  const partes = nombre.split(' - ');
  const espacio = partes.length === 5 && partes[0] === 'Boletines EPS' ? partes[3].lastIndexOf(' ') : -1;
  if (espacio <= 0) return null;
  const curso = cursoPdf_(partes[1], partes[2], partes[3].slice(0, espacio), partes[3].slice(espacio + 1));
  return prefijoLibroCurso_(curso) === partes.slice(0, 4).join(' - ') + ' - ' ? curso : null;
}

function cursoPdf_(hojaOrigen, nivel, periodo, cicloLectivo) {
  return {
    clave: [hojaOrigen, nivel, periodo, cicloLectivo].join('|'),
    etiqueta: hojaOrigen + ' ' + nivel + ' · ' + periodo + ' ' + cicloLectivo,
    hojaOrigen, nivel, periodo, cicloLectivo
  };
}

/**
 * Lo que falta de un libro: sus hojas (falla si no es un libro de boletines terminado), su última carpeta de PDF
 * vigente y las hojas cuyo PDF falta. Sin carpeta vigente, todas las hojas, en una carpeta nueva junto al libro (padre).
 */
function planLibro_(plan, libro, padre) {
  plan.libro = libro;
  plan.padre = padre;
  plan.hojas = hojasEstudiantesLibro_(libro);
  const ultima = ultimaExportacion_(padre, libro, plan.hojas);
  const vigente = ultima && ultima.vigente;
  plan.desactualizada = ultima && !ultima.vigente ? ultima.carpeta : null;
  plan.carpeta = vigente ? ultima.carpeta : null;
  plan.faltan = vigente ? ultima.faltan : plan.hojas;
  plan.yaHabia = plan.hojas.length - plan.faltan.length;
  return plan;
}

// Plan del libro elegido de un curso (grupo de librosPorCursoEnCarpeta_), o del curso sin libro. Un libro que no pudo
// abrirse o no puede exportarse queda con su error y nada para exportar.
function planGrupoPdf_(etiqueta, grupo, padre) {
  const plan = { curso: etiqueta, libros: grupo ? grupo.libros.length : 0, libro: null, hojas: null, carpeta: null, faltan: [], yaHabia: 0, error: '' };
  if (!grupo) return plan;
  plan.fecha = fechaPdf_(grupo.elegido.archivo.getDateCreated());
  try {
    if (!grupo.elegido.libro) throw new Error('no se pudo abrir: ' + grupo.elegido.error);
    planLibro_(plan, grupo.elegido.libro, padre);
  } catch (err) {
    Object.assign(plan, { hojas: null, carpeta: null, faltan: [], yaHabia: 0, error: err.message });
  }
  return plan;
}

function exportarPlanPdf_(ss, plan, inicioEjecucion, ritmo) {
  if (!plan.carpeta) plan.carpeta = crearCarpetaPdf_(plan.padre, plan.libro);
  plan.resultados = exportarHojasPdf_(ss, plan.libro, plan.faltan, plan.carpeta, inicioEjecucion, ritmo).resultados;
}

function fechaPdf_(fecha) {
  return Utilities.formatDate(fecha, CONFIG_EPS.ZONA_HORARIA, 'yyyy-MM-dd HH:mm');
}

function pdfEnCarpeta_(plan) {
  return plan.yaHabia + (plan.resultados || []).filter(r => r.estado === 'Exportado').length;
}

// Una línea por curso: antes de exportar (qué se va a hacer) o después (cómo quedó).
function lineaCursoPdf_(plan, despues) {
  if (!plan.libros) return plan.curso + ': sin libro del período.';
  const libro = plan.curso + ' (libro ' + plan.fecha + (plan.libros > 1 ? ', el más reciente de ' + plan.libros : '') + ')';
  if (plan.error) return libro + ': no se puede exportar: ' + plan.error;
  const total = plan.hojas.length;
  if (!despues) {
    if (!plan.faltan.length) return libro + ': completo (' + total + ' de ' + total + ' PDF), se omite.';
    if (plan.yaHabia) return libro + ': faltan ' + plan.faltan.length + ' de ' + total + ' PDF, se completa su carpeta.';
    if (plan.desactualizada) return libro + ': cambió después de su última exportación (' + plan.desactualizada.getName() +
      '), se exportan sus ' + total + ' PDF en una carpeta nueva.';
    return libro + ': ' + total + ' PDF a exportar.';
  }
  if (!plan.faltan.length && !plan.resultados) return libro + ': ya estaba completo (' + total + ' de ' + total + ' PDF).';
  if (!plan.resultados) return libro + ': no se empezó por el tiempo máximo; faltan ' + plan.faltan.length + ' de ' + total + ' PDF.';
  const enCarpeta = pdfEnCarpeta_(plan);
  const contar = estado => plan.resultados.filter(r => r.estado === estado).length;
  return libro + ': ' + enCarpeta + ' de ' + total + ' PDF (' + contar('Exportado') + ' exportados ahora' +
    (contar('Error') ? ', ' + contar('Error') + ' con error' : '') +
    (contar('Sin exportar') ? ', ' + contar('Sin exportar') + ' sin exportar por el tiempo máximo' : '') + ')' +
    (enCarpeta < total ? '; faltan ' + (total - enCarpeta) + '.' : '.');
}

// Hojas de estudiante del libro de un curso, o del de la generación por DNI: todas menos Resumen, en el orden del libro.
// Sin hoja Resumen no es un libro de la generación por curso (por ejemplo, el propio generador). Con Resumen vacío, la
// generación se cortó antes de terminar (Resumen se escribe al final, después de copiar todas las hojas): sus hojas
// no son el curso completo.
function hojasEstudiantesLibro_(libro) {
  const resumen = libro.getSheetByName(CONFIG_EPS.HOJA_RESUMEN);
  if (!resumen) {
    throw new Error('El libro "' + libro.getName() + '" no tiene la hoja ' + CONFIG_EPS.HOJA_RESUMEN +
      ': no parece un libro de boletines generado por curso.');
  }
  if (![CONFIG_EPS.TITULO_RESUMEN, CONFIG_EPS.TITULO_RESUMEN_DNI].includes(resumen.getRange('A1').getDisplayValue())) {
    throw new Error('El libro "' + libro.getName() + '" no terminó de generarse (su hoja ' + CONFIG_EPS.HOJA_RESUMEN +
      ' está incompleta): volver a generar el curso y mandar ese libro a la papelera.');
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
// la descripción; no se mira el nombre) y las hojas cuyo PDF no tiene (ninguna si está completa). null si no hay
// ninguna carpeta. Una exportación completa la usa sólo para informar; una incompleta, para completarla.
// vigente: el libro no se modificó en Drive después de crearse la carpeta. Si cambió, sus PDF son de otra versión
// del libro: no se completa ni se da por buena, y se exporta todo en una carpeta nueva.
function ultimaExportacion_(padre, libro, hojas) {
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
  return {
    carpeta: ultima,
    faltan: hojas.filter(hoja => !existentes.has(nombrePdf_(hoja))),
    vigente: DriveApp.getFileById(libro.getId()).getLastUpdated() <= ultima.getDateCreated()
  };
}

function nombrePdf_(hoja) {
  return hoja.getName() + '.pdf';
}

/**
 * Una hoja por vez, con una pausa entre pedidos que se adapta al límite de Google (ver CONFIG_PDF.PAUSA_MS): ante un
 * HTTP 429 se duplica y se reintenta la misma hoja; cada PDF exportado la baja un paso. Otro error, o el 429 en el
 * último intento, sólo afecta el PDF de esa hoja; las demás se exportan igual.
 * ritmo (pausa, reintentos, primer pedido) sigue de un libro al siguiente en la exportación de varios cursos:
 * el límite de Google es por usuario, no por libro.
 */
function exportarHojasPdf_(ss, libro, hojas, carpeta, inicioEjecucion, ritmo) {
  const token = ScriptApp.getOAuthToken();
  ritmo = ritmo || nuevoRitmoPdf_();

  const resultados = hojas.map((hoja, i) => {
    const nombre = hoja.getName();
    ss.toast('PDF ' + (i + 1) + ' de ' + hojas.length + ': ' + nombre + '.', 'Informes EPS');
    for (let intento = 1; ; intento++) {
      const espera = ritmo.primero ? 0 : ritmo.pausa;
      if (!hayTiempoPdf_(inicioEjecucion, ritmo)) {
        return { hoja: nombre, estado: 'Sin exportar', detalle: 'se alcanzó el tiempo máximo de la ejecución.' };
      }
      if (espera) Utilities.sleep(espera);
      ritmo.primero = false;
      try {
        carpeta.createFile(exportarHojaPdf_(libro.getId(), hoja, token));
        ritmo.pausa = Math.max(CONFIG_PDF.PAUSA_MS, ritmo.pausa - CONFIG_PDF.PAUSA_MS);
        return { hoja: nombre, estado: 'Exportado', detalle: '' };
      } catch (err) {
        if (!err.limiteGoogle || intento === CONFIG_PDF.INTENTOS_POR_HOJA) {
          return { hoja: nombre, estado: 'Error', detalle: err.message + (intento > 1 ? ' (' + intento + ' intentos)' : '') };
        }
        ritmo.reintentos++;
        ritmo.pausa = Math.min(ritmo.pausa * 2, CONFIG_PDF.PAUSA_MAXIMA_MS);
        ss.toast('Google limitó las exportaciones: se espera ' + ritmo.pausa / 1000 + ' s y se reintenta ' + nombre + '.', 'Informes EPS');
      }
    }
  });
  return { resultados, reintentos: ritmo.reintentos };
}

function nuevoRitmoPdf_() {
  return { pausa: CONFIG_PDF.PAUSA_MS, reintentos: 0, primero: true };
}

// Si todavía entra el próximo pedido, con su pausa, antes del tiempo máximo de la ejecución.
function hayTiempoPdf_(inicioEjecucion, ritmo) {
  return Date.now() + (ritmo.primero ? 0 : ritmo.pausa) - inicioEjecucion <= CONFIG_PDF.TIEMPO_MAXIMO_MS;
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
