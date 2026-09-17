/**
 * Generador de informes individuales EPS - v0.1
 *
 * Uso:
 * 1. Abrir el archivo generador como Google Sheets.
 * 2. Ir a Extensiones > Apps Script.
 * 3. Pegar este código.
 * 4. Guardar.
 * 5. Volver a la planilla y ejecutar: Informes EPS > Generar informe actual.
 *
 * Esta versión:
 * - genera un informe individual por vez;
 * - completa únicamente 1º bimestre;
 * - busca por DNI;
 * - no modifica la planilla fuente;
 * - no modifica la hoja Plantilla;
 * - deja el resultado en Informe_actual para revisión manual.
 */

const CONFIG_EPS = {
  HOJA_CONTROL: 'Control',
  HOJA_PLANTILLA: 'Plantilla',
  HOJA_SALIDA: 'Informe_actual',

  CELDA_FUENTE: 'B4',
  CELDA_HOJA_ORIGEN: 'B5',
  CELDA_DNI: 'B6',
  CELDA_PERIODO: 'B7',
  CELDA_CICLO: 'B8',

  FILA_INICIO_MATERIAS: 9,
  FILAS_MATERIAS_PLANTILLA: 8,

  VALORACIONES_VALIDAS: ['EN PROCESO', 'SUFICIENTE', 'AVANZADO', 'ACREDITO'],
  VALORES_EXCLUIDOS: ['', 'NO CURSA', 'CURSA']
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Informes EPS')
    .addItem('Generar informe actual', 'generarInformeActual')
    .addToUi();
}

function generarInformeActual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const control = ss.getSheetByName(CONFIG_EPS.HOJA_CONTROL);

  if (!control) {
    ui.alert('No se encontró la hoja Control.');
    return;
  }

  try {
    const fuenteIngresada = texto_(control.getRange(CONFIG_EPS.CELDA_FUENTE).getDisplayValue());
    const hojaOrigen = texto_(control.getRange(CONFIG_EPS.CELDA_HOJA_ORIGEN).getDisplayValue());
    const dniBuscado = normalizarDni_(control.getRange(CONFIG_EPS.CELDA_DNI).getDisplayValue());
    const cicloLectivo = texto_(control.getRange(CONFIG_EPS.CELDA_CICLO).getDisplayValue()) || '2026';

    if (!fuenteIngresada) {
      throw new Error('Falta completar la URL o el ID de la planilla de valoraciones en Control!B4.');
    }
    if (!hojaOrigen) {
      throw new Error('Falta indicar la hoja origen en Control!B5.');
    }
    if (!dniBuscado) {
      throw new Error('Falta completar el DNI del estudiante en Control!B6.');
    }

    const idFuente = extraerSpreadsheetId_(fuenteIngresada);
    const libroFuente = SpreadsheetApp.openById(idFuente);
    const hojaFuente = libroFuente.getSheetByName(hojaOrigen);

    if (!hojaFuente) {
      throw new Error('No se encontró la hoja origen "' + hojaOrigen + '" en la planilla de valoraciones.');
    }

    const valores = hojaFuente.getDataRange().getDisplayValues();
    const mapa = detectarEstructuraFuente_(valores);
    const coincidencias = buscarFilasPorDni_(valores, dniBuscado);

    if (coincidencias.length === 0) {
      throw new Error('No se encontró el DNI indicado en la hoja "' + hojaOrigen + '".');
    }

    if (coincidencias.length > 1) {
      const filas = coincidencias.map(function(f) { return f + 1; }).join(', ');
      throw new Error('El DNI indicado aparece más de una vez en la hoja "' + hojaOrigen + '". Filas detectadas: ' + filas + '. No se genera el informe para evitar ambigüedad.');
    }

    const estudiante = construirRegistroEstudiante_(valores, coincidencias[0], mapa, hojaOrigen);
    const advertencias = estudiante.advertencias.slice();

    if (estudiante.materias.length === 0) {
      advertencias.push('No se detectaron materias con valoración válida para incluir en el informe.');
    }

    crearInforme_(ss, estudiante, cicloLectivo);

    const mensajeFinal = [
      'Informe generado en la hoja Informe_actual.',
      '',
      'Estudiante: ' + estudiante.nombreCompleto,
      'DNI: ' + estudiante.dni,
      'Materias incluidas: ' + estudiante.materias.length,
      advertencias.length ? '\nAdvertencias:\n- ' + advertencias.join('\n- ') : ''
    ].join('\n');

    registrarEstado_(control, 'Informe generado', advertencias);
    ui.alert(mensajeFinal);

  } catch (error) {
    registrarEstado_(control, 'No se generó el informe', [error.message]);
    ui.alert('No se generó el informe.\n\n' + error.message);
  }
}

function detectarEstructuraFuente_(valores) {
  const encabezado1 = valores[0] || [];
  const encabezado2 = valores[1] || [];
  const cantidadColumnas = Math.max(encabezado1.length, encabezado2.length);

  const especiales = {
    orientacion: -1,
    observacionesGenerales: -1,
    observacionesAsistencia: -1,
    observacionesConvivencia: -1,
    ptf: -1
  };

  for (let c = 0; c < cantidadColumnas; c++) {
    const h1 = normalizarTexto_(encabezado1[c]);

    if (h1 === 'ORIENTACION') {
      especiales.orientacion = c;
    } else if (h1 === 'OBSERVACIONES GENERALES SOBRE DESEMPENO ACADEMICO') {
      especiales.observacionesGenerales = c;
    } else if (h1 === 'OBSERVACIONES SOBRE ASISTENCIA') {
      especiales.observacionesAsistencia = c;
    } else if (h1 === 'OBSERVACIONES SOBRE CONVIVENCIA ESCOLAR') {
      especiales.observacionesConvivencia = c;
    } else if (h1.indexOf('PLAN DE TRABAJO FORMATIVO') === 0) {
      especiales.ptf = c;
    }
  }

  const materias = [];

  for (let c = 3; c < cantidadColumnas - 1; c++) {
    const h1Original = texto_(encabezado1[c]);
    const h2Original = texto_(encabezado2[c]);
    const nombreMateria = h2Original || h1Original;

    if (!nombreMateria) {
      continue;
    }

    const nombreNormalizado = normalizarTexto_(nombreMateria);
    const h1Normalizado = normalizarTexto_(h1Original);

    if (/^EPS\s*\d+$/i.test(h1Original) || /^EPS\s*\d+$/i.test(nombreMateria)) {
      continue;
    }

    if (nombreNormalizado === 'ORIENTACION') {
      continue;
    }

    if (nombreNormalizado.indexOf('OBSERVACIONES') === 0 || nombreNormalizado.indexOf('PLAN DE TRABAJO FORMATIVO') === 0) {
      continue;
    }

    if (h1Normalizado.indexOf('OBSERVACIONES') === 0 || h1Normalizado.indexOf('PLAN DE TRABAJO FORMATIVO') === 0) {
      continue;
    }

    materias.push({
      nombre: nombreMateria,
      columnaEtiqueta: c,
      columnaValor: c + 1
    });
  }

  return {
    materias: materias,
    especiales: especiales
  };
}

function buscarFilasPorDni_(valores, dniBuscado) {
  const coincidencias = [];

  for (let r = 0; r < valores.length; r++) {
    const dniFila = normalizarDni_(valores[r][2]);

    if (dniFila && dniFila === dniBuscado) {
      coincidencias.push(r);
    }
  }

  return coincidencias;
}

function construirRegistroEstudiante_(valores, fila, mapa, hojaOrigen) {
  const filaValores = valores[fila] || [];

  const apellido = texto_(filaValores[0]);
  const nombres = texto_(filaValores[1]);
  const dni = texto_(filaValores[2]);
  const nombreCompleto = [apellido, nombres].filter(Boolean).join(', ') || '(sin nombre en la fuente)';

  const nivel = detectarNivelPorFila_(valores, fila);
  const orientacion = obtenerValor_(valores, fila, mapa.especiales.orientacion);
  const nivelVisual = construirNivelVisual_(hojaOrigen, nivel, orientacion);

  const materias = [];
  const advertencias = [];

  mapa.materias.forEach(function(materia) {
    const valorOriginal = obtenerValor_(valores, fila, materia.columnaValor);
    const valorNormalizado = normalizarTexto_(valorOriginal);

    if (CONFIG_EPS.VALORACIONES_VALIDAS.indexOf(valorNormalizado) !== -1) {
      materias.push({
        nombre: materia.nombre,
        valoracion: valorCanonico_(valorNormalizado)
      });
      return;
    }

    if (CONFIG_EPS.VALORES_EXCLUIDOS.indexOf(valorNormalizado) !== -1) {
      return;
    }

    advertencias.push('Valor inesperado en "' + materia.nombre + '": "' + valorOriginal + '".');
  });

  return {
    apellido: apellido,
    nombres: nombres,
    nombreCompleto: nombreCompleto,
    dni: dni,
    nivelVisual: nivelVisual,
    orientacion: orientacion,
    materias: materias,
    observacionesGenerales: obtenerValor_(valores, fila, mapa.especiales.observacionesGenerales),
    observacionesAsistencia: obtenerValor_(valores, fila, mapa.especiales.observacionesAsistencia),
    observacionesConvivencia: obtenerValor_(valores, fila, mapa.especiales.observacionesConvivencia),
    ptf: obtenerValor_(valores, fila, mapa.especiales.ptf),
    advertencias: advertencias
  };
}

function crearInforme_(ss, estudiante, cicloLectivo) {
  const plantilla = ss.getSheetByName(CONFIG_EPS.HOJA_PLANTILLA);

  if (!plantilla) {
    throw new Error('No se encontró la hoja Plantilla.');
  }

  const salidaExistente = ss.getSheetByName(CONFIG_EPS.HOJA_SALIDA);
  if (salidaExistente) {
    ss.deleteSheet(salidaExistente);
  }

  const salida = plantilla.copyTo(ss);
  salida.setName(CONFIG_EPS.HOJA_SALIDA);
  ss.setActiveSheet(salida);

  // Encabezado.
  salida.getRange('J3').setValue('CICLO LECTIVO: ' + cicloLectivo);
  salida.getRange('E4').setValue(estudiante.nombreCompleto);
  salida.getRange('E5').setValue(estudiante.dni);
  salida.getRange('E6').setValue(estudiante.nivelVisual);

  // Limpieza de restos usados por la plantilla original.
  salida.getRange('A1:A3').clearContent();

  const materias = estudiante.materias.length
    ? estudiante.materias
    : [{ nombre: 'Sin materias con valoración válida para incluir', valoracion: '-' }];

  const filasNecesarias = materias.length;
  const filasBase = CONFIG_EPS.FILAS_MATERIAS_PLANTILLA;
  const filaInicio = CONFIG_EPS.FILA_INICIO_MATERIAS;
  const delta = filasNecesarias - filasBase;

  if (filasNecesarias > filasBase) {
    salida.insertRowsBefore(filaInicio + filasBase, filasNecesarias - filasBase);
    for (let r = filaInicio + filasBase; r < filaInicio + filasNecesarias; r++) {
      salida.getRange(filaInicio + filasBase - 1, 1, 1, 12).copyTo(
        salida.getRange(r, 1, 1, 12),
        { formatOnly: true }
      );
      salida.getRange(r, 3, 1, 2).merge();
      salida.setRowHeight(r, salida.getRowHeight(filaInicio + filasBase - 1));
    }
  }

  if (filasNecesarias < filasBase) {
    salida.deleteRows(filaInicio + filasNecesarias, filasBase - filasNecesarias);
  }

  for (let i = 0; i < materias.length; i++) {
    const filaDestino = filaInicio + i;
    salida.getRange(filaDestino, 3).setValue(materias[i].nombre);
    salida.getRange(filaDestino, 5).setValue(materias[i].valoracion);
    salida.getRange(filaDestino, 6, 1, 7).setValues([['-', '-', '-', '-', '-', '-', '-']]);
  }

  // Observaciones. Las filas se mueven si se insertan o eliminan materias.
  const filaObsGenerales = 20 + delta;
  const filaObsAsistencia = 23 + delta;
  const filaObsConvivencia = 26 + delta;
  const filaPtf = 29 + delta;

  escribirCajaTexto_(salida, filaObsGenerales, estudiante.observacionesGenerales);
  escribirCajaTexto_(salida, filaObsAsistencia, estudiante.observacionesAsistencia);
  escribirCajaTexto_(salida, filaObsConvivencia, estudiante.observacionesConvivencia);
  escribirCajaTexto_(salida, filaPtf, estudiante.ptf);

  salida.getDataRange().setWrap(true);
}

function escribirCajaTexto_(hoja, fila, valor) {
  const texto = texto_(valor) || '-';
  hoja.getRange(fila, 4).setValue(texto);
  hoja.getRange(fila, 4, 1, 7).setWrap(true);
  hoja.setRowHeight(fila, Math.max(48, hoja.getRowHeight(fila)));
}

function detectarNivelPorFila_(valores, fila) {
  for (let r = fila; r >= 0; r--) {
    const posibleNivel = texto_(valores[r][0]);
    if (/^EPS\s*\d+$/i.test(posibleNivel)) {
      return posibleNivel.replace(/\s+/g, '');
    }
  }
  return '';
}

function construirNivelVisual_(hojaOrigen, nivel, orientacion) {
  const numeroNivel = nivel ? nivel.replace(/\D/g, '') : '';
  const base = numeroNivel ? hojaOrigen + numeroNivel : hojaOrigen;
  return [base, orientacion].filter(Boolean).join(' - ');
}

function extraerSpreadsheetId_(entrada) {
  const textoEntrada = texto_(entrada);

  const matchUrl = textoEntrada.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (matchUrl && matchUrl[1]) {
    return matchUrl[1];
  }

  const matchId = textoEntrada.match(/^[a-zA-Z0-9-_]{20,}$/);
  if (matchId) {
    return textoEntrada;
  }

  throw new Error('No pude reconocer el ID de la planilla fuente. Pegá el link completo de Google Sheets o el ID del archivo.');
}

function obtenerValor_(valores, fila, columna) {
  if (columna === undefined || columna === null || columna < 0) {
    return '';
  }
  return texto_((valores[fila] || [])[columna]);
}

function valorCanonico_(valorNormalizado) {
  if (valorNormalizado === 'ACREDITO') {
    return 'ACREDITÓ';
  }
  if (valorNormalizado === 'EN PROCESO') {
    return 'En proceso';
  }
  if (valorNormalizado === 'SUFICIENTE') {
    return 'Suficiente';
  }
  if (valorNormalizado === 'AVANZADO') {
    return 'Avanzado';
  }
  return valorNormalizado;
}

function normalizarDni_(valor) {
  return texto_(valor).replace(/\D/g, '');
}

function normalizarTexto_(valor) {
  return texto_(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function texto_(valor) {
  if (valor === null || valor === undefined) {
    return '';
  }
  return String(valor).trim();
}

function registrarEstado_(control, resultado, advertencias) {
  try {
    control.getRange('F4').setValue(new Date());
    control.getRange('F5').setValue(resultado);
    control.getRange('F6').setValue(advertencias && advertencias.length ? advertencias.join('\n') : 'Sin advertencias.');
  } catch (e) {
    // No se interrumpe la generación si falla el registro del estado.
  }
}
