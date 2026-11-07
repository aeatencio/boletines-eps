/**
 * Carga docente EPS (#26) · núcleo de consulta y guardado por campo.
 *
 * Proyecto de Apps Script propio, separado del generador. Este archivo no usa servicios de Google: recibe por
 * parámetro las hojas (Sheet / Range), el lock, el reloj y los datos de configuración, así que corre igual en
 * Apps Script y en las pruebas locales (tests/carga.js). Carga_Google.gs abre los recursos reales.
 * Todas las funciones terminan en "_": google.script.run no puede invocarlas.
 *
 * Contratos compartidos con el generador (src/Generador_EPS.gs), copiados y verificados por las pruebas de paridad:
 * nombres de los períodos, celdas de Control con la grilla de cada período, hojas EPSI / EPSG, estructura de la
 * grilla (materias antes de ORIENTACIÓN, pares etiqueta / valor, bloques de cuatro filas por estudiante),
 * reconocimiento del ID de una planilla y valores que el generador acepta en cada período.
 *
 * Garantías del guardado (ver guardarCampo_):
 *  - cada pedido nombra el período y la grilla que vio el docente; se rechaza si la fuente vigente es otra;
 *  - configuración, ventana, fuente y asignación se validan al empezar y otra vez dentro del lock, después de
 *    esperarlo; la ventana se comprueba además inmediatamente antes de escribir;
 *  - el Registro tiene a lo sumo una fila por operación, y la fila precede a la escritura en la grilla;
 *  - todo lo escrito bajo el lock (Registro y grilla) se confirma con flush antes de liberarlo, en todos los
 *    caminos; si esa confirmación falla, nunca se responde «guardado»;
 *  - una operación intenta escribir la grilla a lo sumo una vez: después de un intento (con o sin error), un
 *    reintento con la misma operación sólo observa e informa. Sólo se repite lo que nunca llegó a intentarse;
 *  - el Registro distingue lo confirmado de lo inferido: 'aplicado' sólo si Sheets aceptó la escritura sin error
 *    y la celda mostró el valor nuevo enseguida, en su lugar; si el valor nuevo se ve pero no se puede atribuir a
 *    esta operación, 'inferido';
 *  - «guardado» sólo se responde con la fila 'aplicado' y el valor nuevo en la celda. «coincide» dice que la celda
 *    muestra el valor pedido sin poder confirmar que lo escribió este pedido. «reintentar» y «rechazado»
 *    garantizan que ese pedido no intentó escribir la grilla.
 *
 * Límites (Sheets no ofrece escrituras condicionales, transacciones entre planillas ni historial por celda):
 *  - un cambio de forma de la grilla (insertar, borrar u ordenar filas o columnas) entre la última lectura y la
 *    escritura, sin llamadas al servicio de por medio, puede desviar esa escritura: se detecta después y queda
 *    'revisar' con la celda a verificar. Condición operativa: no cambiar la forma de las grillas con la carga abierta;
 *  - un cambio de configuración o asignación en el instante entre la revalidación dentro del lock y la escritura
 *    no la detiene;
 *  - después de un error ambiguo, una corrección del Equipo EPS que devuelve el valor anterior no se distingue de
 *    una escritura no aplicada: por eso queda 'incierto' y no se reintenta solo; el docente decide volver a cargar.
 */
const CARGA = {
  // Generador: CONFIG_EPS.PERIODOS (en el orden del año) y CONFIG_EPS.CELDA_FUENTE_*.
  PERIODOS: ['1º bimestre', '1º cuatrimestre', '3º bimestre', '2º cuatrimestre', 'Anual'],
  CELDAS_FUENTE: { '1º bimestre': 'B10', '1º cuatrimestre': 'B11', '3º bimestre': 'B12', '2º cuatrimestre': 'B13', 'Anual': 'B14' },
  HOJA_CONTROL: 'Control',
  HOJAS_ORIGEN: ['EPSI', 'EPSG'],
  // Generador: CONFIG_EPS.ENCABEZADOS_GENERALES y ENCABEZADOS_OPERATIVOS (nunca son materias).
  ENCABEZADOS_COLA: ['OBSERVACIONES GENERALES SOBRE DESEMPEÑO ACADÉMICO', 'OBSERVACIONES SOBRE ASISTENCIA',
    'OBSERVACIONES SOBRE CONVIVENCIA ESCOLAR', 'PLAN DE TRABAJO FORMATIVO (PTF)', 'LISTO PARA EXPORTAR', 'EXPORTADO'],

  CAMPOS: [
    { id: 'calificacion', nombre: 'Calificación', desplazamiento: 0, etiqueta: 'VALORACION' },
    { id: 'calidad', nombre: 'Calidad de las producciones', desplazamiento: 1, etiqueta: 'CALIDAD DE LAS PRODUCCIONES' },
    { id: 'participacion', nombre: 'Participación', desplazamiento: 2, etiqueta: 'PARTICIPACION' },
    { id: 'observaciones', nombre: 'Observaciones', desplazamiento: 3, etiqueta: 'OBSERVACIONES' }
  ],
  PENDIENTE: 'CURSA',
  NO_CURSA: 'NO CURSA',
  // Lo que el docente puede elegir en Calificación. Todo es válido para clasificarValor_ del generador (paridad).
  ESCALA_DOCENTE: {
    '1º bimestre': ['En proceso', 'Suficiente', 'Avanzado'],
    '3º bimestre': ['En proceso', 'Suficiente', 'Avanzado'],
    '1º cuatrimestre': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'],
    '2º cuatrimestre': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'],
    'Anual': ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']
  },
  // Los establece el Equipo EPS: Calificación queda en solo lectura y los otros tres campos siguen editables.
  ADMINISTRATIVOS: ['ACREDITÓ', 'EQUIVALENCIA'],
  // Opciones cerradas actuales de las grillas (hoja Opciones). Sin texto libre; vacío = sin dato.
  OPCIONES: {
    calidad: ['Regular', 'Buena', 'Muy buena', 'Excelente', 'No entrega'],
    participacion: ['Ninguna', 'Poca', 'Activa']
  },
  MAX_OBSERVACIONES: 1000,

  // Ventanas: abren a las 00:00 de «Desde» y cierran a las 8:00 del primer día hábil posterior a «Hasta».
  // America/Argentina/Buenos_Aires es UTC−3 todo el año (sin horario de verano desde 2009).
  ZONA: 'America/Argentina/Buenos_Aires',
  DESFASE: '-03:00',
  HORA_APERTURA: '00:00',
  HORA_CIERRE: '08:00',
  MAX_DIAS_HASTA_CIERRE: 31,

  ESPERA_LOCK_MS: 10000,
  OPERACION: /^[A-Za-z0-9_-]{8,80}$/,

  // Planilla de administración: el Equipo EPS edita Ventanas, Feriados y Asignaciones; el Registro lo escribe la carga.
  HOJAS: { ventanas: 'Ventanas', feriados: 'Feriados', asignaciones: 'Asignaciones', registro: 'Registro' },
  ENCABEZADOS: {
    ventanas: ['Período', 'Desde', 'Hasta'],
    feriados: ['Fecha', 'Descripción'],
    asignaciones: ['Orientación', 'Nivel', 'Materia', 'Docente', 'Correo']
  },
  // Columnas obligatorias de cada hoja (Descripción de Feriados es sólo para el Equipo EPS).
  OBLIGATORIOS: { ventanas: ['Período', 'Desde', 'Hasta'], feriados: ['Fecha'], asignaciones: ['Orientación', 'Nivel', 'Materia', 'Docente', 'Correo'] },
  // Registro: estado y detalle contiguos, para marcarlos con una sola escritura.
  REGISTRO: ['Fecha', 'Estado', 'Detalle', 'Intentos', 'Operación', 'Correo', 'Docente', 'Período', 'Fuente',
    'Orientación', 'Nivel', 'Materia', 'DNI', 'Estudiante', 'Campo', 'Anterior', 'Nuevo']
};
const COL_REGISTRO = Object.fromEntries(CARGA.REGISTRO.map((c, i) => [c, i]));

// ================================================================ errores

/** Rechazo: una regla, la configuración o la autorización impiden la operación. Nunca se escribió la grilla. */
function rechazo_(codigo, motivo) {
  const e = new Error(motivo);
  e.rechazo = { codigo, motivo };
  return e;
}

// ================================================================ fechas (fechas civiles 'aaaa-mm-dd' en hora argentina)

/** Fecha civil de una celda: 'aaaa-mm-dd' o 'dd/mm/aaaa'. Las fechas de Sheets llegan ya convertidas por el adaptador. */
function fechaCivil_(valor) {
  const t = txt_(valor);
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  let a, mes, d;
  if (m) { a = +m[1]; mes = +m[2]; d = +m[3]; } else {
    m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return '';
    d = +m[1]; mes = +m[2]; a = +m[3];
  }
  const f = new Date(Date.UTC(a, mes - 1, d));
  if (f.getUTCFullYear() !== a || f.getUTCMonth() !== mes - 1 || f.getUTCDate() !== d) return '';
  return f.toISOString().slice(0, 10);
}

function instanteLocal_(fechaIso, hora) { return new Date(fechaIso + 'T' + hora + ':00' + CARGA.DESFASE); }

function fechaLocalDe_(instante) {
  const desfaseMs = Number(CARGA.DESFASE.slice(0, 3)) * 3600000;
  return new Date(instante.getTime() + desfaseMs).toISOString().slice(0, 10);
}

function diaSiguiente_(fechaIso) {
  const f = new Date(fechaIso + 'T00:00:00Z');
  f.setUTCDate(f.getUTCDate() + 1);
  return f.toISOString().slice(0, 10);
}

/**
 * 8:00 del primer día hábil posterior a «Hasta»: ni sábado, ni domingo, ni feriado nacional o de CABA.
 * Si un día que hay que mirar es de un año sin feriados cargados, no se adivina: es un problema de configuración.
 */
function cierreVentana_(hastaIso, feriados) {
  const anios = {};
  Object.keys(feriados).forEach(f => { anios[f.slice(0, 4)] = true; });
  let dia = diaSiguiente_(hastaIso);
  for (let i = 0; i < CARGA.MAX_DIAS_HASTA_CIERRE; i++, dia = diaSiguiente_(dia)) {
    if (!anios[dia.slice(0, 4)]) throw new Error('Faltan los feriados de ' + dia.slice(0, 4) + ' en la hoja ' + CARGA.HOJAS.feriados + ': no se puede calcular el cierre.');
    const semana = new Date(dia + 'T12:00:00Z').getUTCDay();
    if (semana !== 0 && semana !== 6 && !feriados[dia]) return { dia, instante: instanteLocal_(dia, CARGA.HORA_CIERRE) };
  }
  throw new Error('No hay un día hábil en el mes posterior al ' + hastaIso + '.');
}

// ================================================================ configuración (planilla de administración + Control)

/** Filas de una hoja de administración como objetos, con las columnas ubicadas por encabezado. */
function tablaAdministracion_(filas, hoja) {
  const nombre = CARGA.HOJAS[hoja];
  if (!filas || !filas.length) throw new Error('Falta la hoja ' + nombre + ' en la planilla de administración, o está vacía.');
  const enc = filas[0].map(norm_);
  const indices = {};
  CARGA.ENCABEZADOS[hoja].forEach(campo => {
    const pos = enc.map((e, i) => (e === norm_(campo) ? i : -1)).filter(i => i >= 0);
    if (pos.length > 1 || (!pos.length && CARGA.OBLIGATORIOS[hoja].includes(campo))) {
      throw new Error('La hoja ' + nombre + ' debe tener una única columna "' + campo + '".');
    }
    if (pos.length) indices[campo] = pos[0];
  });
  const filasDatos = [];
  for (let r = 1; r < filas.length; r++) {
    if (!filas[r].some(v => txt_(v))) continue;
    const obj = { fila: r + 1 };
    Object.keys(indices).forEach(campo => { obj[campo] = txt_(filas[r][indices[campo]]); });
    filasDatos.push(obj);
  }
  return filasDatos;
}

/**
 * Configuración vigente, leída en cada operación. datos: { ventanas, feriados, asignaciones } (matrices con
 * encabezado; fechas como 'aaaa-mm-dd' o 'dd/mm/aaaa') y fuentes: { período: valor de su celda de Control }.
 * Nunca lanza por un problema de datos. Lo anota donde bloquea lo mínimo, siempre cerrando y nunca adivinando:
 *  - problemas: hojas ilegibles o feriados mal cargados; bloquean toda la carga;
 *  - problemas de un período (fuente, fechas, cierre): bloquean ese período;
 *  - materias ambiguas (dos docentes, niveles distintos o una fila inválida): bloquean esa materia;
 *  - advertencias: filas que no se pueden vincular a nada; no habilitan a nadie.
 */
function configuracion_(datos, ahora) {
  const config = { problemas: [], advertencias: [], periodos: {}, asignaciones: [], porMateria: {} };
  CARGA.PERIODOS.forEach(p => { config.periodos[p] = { periodo: p, fuente: '', ventana: null, estado: 'sin ventana', problemas: [], problemasFuente: [] }; });

  // Feriados nacionales y de CABA.
  const feriados = {};
  try {
    tablaAdministracion_(datos.feriados, 'feriados').forEach(f => {
      const fecha = fechaCivil_(f['Fecha']);
      if (!fecha) config.problemas.push('Feriados, fila ' + f.fila + ': fecha no reconocida ("' + f['Fecha'] + '").');
      else feriados[fecha] = true;
    });
  } catch (e) { config.problemas.push(e.message); }

  // Fuentes: la grilla de cada período en Control del generador. Dos períodos nunca comparten grilla.
  const ids = {};
  CARGA.PERIODOS.forEach(p => {
    const valor = txt_((datos.fuentes || {})[p]);
    const id = idPlanilla_(valor);
    config.periodos[p].fuente = id;
    if (valor && !id) config.periodos[p].problemasFuente.push('Control!' + CARGA.CELDAS_FUENTE[p] + ' del generador no tiene una grilla reconocible para ' + p + '.');
    if (id) (ids[id] = ids[id] || []).push(p);
  });
  Object.keys(ids).filter(id => ids[id].length > 1).forEach(id => {
    ids[id].forEach(p => config.periodos[p].problemasFuente.push('La grilla de ' + p + ' es la misma que la de ' + ids[id].filter(o => o !== p).join(' y ') +
      ' en Control del generador: cada período debe tener su propia grilla.'));
  });

  // Ventanas: una fila por período habilitado. Los períodos con las mismas fechas comparten ventana.
  try {
    const vistas = {};
    tablaAdministracion_(datos.ventanas, 'ventanas').forEach(v => {
      const periodo = normalizarPeriodo_(v['Período']);
      if (!config.periodos[periodo]) return config.advertencias.push('Ventanas, fila ' + v.fila + ': período no reconocido ("' + v['Período'] + '"). Esa fila no abre nada.');
      const p = config.periodos[periodo];
      if (vistas[periodo]) return p.problemas.push('El período ' + periodo + ' figura más de una vez en Ventanas (filas ' + vistas[periodo] + ' y ' + v.fila + ').');
      vistas[periodo] = v.fila;
      const desde = fechaCivil_(v['Desde']);
      const hasta = fechaCivil_(v['Hasta']);
      if (!desde || !hasta) return p.problemas.push('Ventanas, fila ' + v.fila + ': fechas no reconocidas para ' + periodo + '.');
      if (desde > hasta) return p.problemas.push('Ventanas, fila ' + v.fila + ': «Desde» es posterior a «Hasta» para ' + periodo + '.');
      p.ventana = { desde, hasta, apertura: instanteLocal_(desde, CARGA.HORA_APERTURA), cierre: null, diaCierre: '' };
    });
  } catch (e) { config.problemas.push(e.message); }

  CARGA.PERIODOS.forEach(nombre => {
    const p = config.periodos[nombre];
    if (!p.ventana) return; // sin ventana nadie lo carga: su fuente no importa
    p.problemas = p.problemas.concat(p.problemasFuente);
    if (!p.fuente) p.problemas.push('Falta la grilla de ' + nombre + ' en Control!' + CARGA.CELDAS_FUENTE[nombre] + ' del generador.');
    try {
      const cierre = cierreVentana_(p.ventana.hasta, feriados);
      p.ventana.cierre = cierre.instante;
      p.ventana.diaCierre = cierre.dia;
    } catch (e) { p.problemas.push(e.message); }
    if (ahora < p.ventana.apertura) p.estado = 'futura';
    else if (!p.ventana.cierre || ahora < p.ventana.cierre) p.estado = 'abierta';
    else p.estado = 'cerrada';
  });

  // Asignaciones: un único docente por orientación y materia, con el nivel de la materia en la grilla.
  // Una fila inválida que nombra una materia reconocible la bloquea: no se decide por descarte quién la carga.
  try {
    tablaAdministracion_(datos.asignaciones, 'asignaciones').forEach(a => {
      const fila = {
        fila: a.fila, orientacion: normalizarHojaOrigen_(a['Orientación']), nivel: normalizarNivel_(a['Nivel']),
        materia: a['Materia'], clave: clave_(a['Materia']), docente: a['Docente'], correo: correo_(a['Correo'])
      };
      const falta = !fila.orientacion ? 'la orientación (EPSI o EPSG)' : !fila.clave ? 'la materia' : !fila.nivel ? 'el nivel (EPS1 a EPS4)' :
        !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(fila.correo) ? 'el correo' : '';
      if (falta) {
        config.advertencias.push('Asignaciones, fila ' + a.fila + ': falta o no se reconoce ' + falta + '. Esa fila no habilita a nadie' +
          (fila.orientacion && fila.clave ? ' y nadie puede cargar ' + fila.orientacion + ' · ' + fila.materia + ' hasta corregirla.' : '.'));
        if (!fila.orientacion || !fila.clave) return;
        fila.invalida = true;
      } else {
        config.asignaciones.push(fila);
      }
      const k = fila.orientacion + '|' + fila.clave;
      (config.porMateria[k] = config.porMateria[k] || []).push(fila);
    });
  } catch (e) { config.problemas.push(e.message); }
  Object.keys(config.porMateria).forEach(k => {
    const filas = config.porMateria[k];
    const correos = filas.map(f => f.correo).filter((c, i, l) => l.indexOf(c) === i);
    const niveles = filas.map(f => f.nivel).filter((c, i, l) => l.indexOf(c) === i);
    if (filas.some(f => f.invalida) || correos.length > 1 || niveles.length > 1) {
      if (!filas.some(f => f.invalida)) {
        config.advertencias.push('Asignaciones, filas ' + filas.map(f => f.fila).join(', ') + ': ' + filas[0].orientacion + ' · ' + filas[0].materia +
          (correos.length > 1 ? ' tiene más de un docente.' : ' tiene niveles distintos.') + ' Nadie puede cargarla hasta corregirlo.');
      }
      filas.forEach(f => { f.ambigua = true; });
    }
  });
  return config;
}

/** Período vigente para una operación: existe, con la ventana abierta, sin problemas y con la fuente que vio el docente. */
function periodoHabilitado_(config, periodo, fuente) {
  const p = config.periodos[periodo];
  if (!p) throw rechazo_('periodo', 'El período "' + txt_(periodo) + '" no es uno de los de la carga.');
  if (p.estado === 'sin ventana' || p.estado === 'futura') throw rechazo_('ventana', 'La carga de ' + periodo + ' no está abierta.');
  if (p.estado === 'cerrada') throw rechazo_('ventana', 'La carga de ' + periodo + ' cerró.');
  if (config.problemas.length || p.problemas.length) {
    throw rechazo_('configuracion', 'La carga de ' + periodo + ' no está disponible por un problema de configuración. Avisá al Equipo EPS.');
  }
  if (fuente !== undefined && txt_(fuente) !== p.fuente) {
    throw rechazo_('fuente', 'La grilla de ' + periodo + ' cambió desde que abriste la carga. Volvé a abrirla.');
  }
  return p;
}

function asignacionDe_(config, orientacion, clave) {
  const filas = config.porMateria[orientacion + '|' + clave] || [];
  return filas.length && !filas[0].ambigua ? filas[0] : null;
}

/** La materia (tal como está en la grilla) está asignada a este correo, con el mismo nivel. Se llama en cada operación. */
function autorizarMateria_(config, correo, orientacion, materia) {
  const a = asignacionDe_(config, orientacion, materia.clave);
  if (!a || a.correo !== correo) throw rechazo_('asignacion', 'Esta materia no está asignada a tu cuenta. Si es un error, avisá al Equipo EPS.');
  if (a.nivel !== materia.nivel) throw rechazo_('asignacion', 'La asignación de ' + materia.nombre + ' tiene un nivel distinto del de la grilla. Avisá al Equipo EPS.');
  return a;
}

// ================================================================ estructura de la grilla (paridad con el generador)

/**
 * Materias de una hoja de la grilla, con las mismas reglas que detectarEstructuraFuente_ del generador, más el
 * nivel (sección EPSn) de cada una. Un problema de estructura es un rechazo: nunca se escribe en una grilla dudosa.
 */
function estructuraGrilla_(valores, periodo) {
  const e1 = valores[0] || [];
  const e2 = valores[1] || [];
  const maxCols = Math.max(e1.length, e2.length);
  const enGrilla = ' en la grilla de ' + periodo + '. Avisá al Equipo EPS.';
  const orientacion = [];
  for (let c = 3; c < maxCols; c++) if (norm_(e1[c]) === 'ORIENTACION' || norm_(e2[c]) === 'ORIENTACION') orientacion.push(c);
  if (orientacion.length !== 1) throw rechazo_('grilla', 'No hay una única columna ORIENTACIÓN' + enGrilla);
  const cola = CARGA.ENCABEZADOS_COLA.map(norm_);
  const materias = [];
  const vistas = {};
  let nivel = '';
  for (let c = 3; c < orientacion[0]; c++) {
    if (/^EPS\d+$/.test(norm_(e1[c]))) nivel = norm_(e1[c]);
    const nombre = txt_(e2[c]) || txt_(e1[c]);
    if (!nombre || /^EPS\d+$/.test(norm_(nombre))) continue;
    if (cola.includes(norm_(nombre))) throw rechazo_('grilla', 'El encabezado "' + nombre + '" está antes de ORIENTACIÓN' + enGrilla);
    if (c + 1 >= orientacion[0] || txt_(e1[c + 1]) || txt_(e2[c + 1])) throw rechazo_('grilla', 'La materia "' + nombre + '" no tiene su columna de valoración' + enGrilla);
    const clave = clave_(nombre);
    if (vistas[clave]) throw rechazo_('grilla', 'La materia "' + nombre + '" está repetida' + enGrilla);
    vistas[clave] = true;
    materias.push({ nombre, clave, nivel, columnaValor: c + 1 });
  }
  return { materias, inicioCola: orientacion[0] };
}

function materiaPorClave_(estructura, clave) {
  const materia = estructura.materias.find(m => m.clave === clave);
  if (!materia) throw rechazo_('grilla', 'La materia ya no figura en la grilla. Avisá al Equipo EPS.');
  return materia;
}

function filasConDni_(valores, dni) {
  const filas = [];
  for (let r = 0; r < valores.length; r++) if (dni_(valores[r][2]) === dni) filas.push(r);
  return filas;
}

/**
 * Fila de Valoración del estudiante, por DNI. El bloque debe estar entero: las tres filas siguientes sin DNI ni
 * marcador de nivel, y con las etiquetas esperadas donde las hay. Nunca se escribe en una fila deducida.
 */
function filaEstudiante_(valores, estructura, dni) {
  const filas = filasConDni_(valores, dni);
  if (filas.length > 1) throw rechazo_('grilla', 'El DNI figura más de una vez en la hoja de la grilla. Avisá al Equipo EPS.');
  if (!filas.length) throw rechazo_('grilla', 'El estudiante ya no figura en la grilla. Volvé a abrir la carga.');
  validarBloque_(valores, estructura, filas[0]);
  return filas[0];
}

function validarBloque_(valores, estructura, r) {
  CARGA.CAMPOS.forEach(campo => {
    const fila = valores[r + campo.desplazamiento];
    if (!fila) throw rechazo_('grilla', 'El bloque del estudiante está incompleto en la grilla. Avisá al Equipo EPS.');
    if (campo.desplazamiento > 0 && (dni_(fila[2]) || /^EPS\d+$/.test(norm_(fila[0])))) {
      throw rechazo_('grilla', 'El bloque del estudiante no tiene sus cuatro filas en la grilla. Avisá al Equipo EPS.');
    }
    const etiquetas = estructura.materias.map(m => etiqueta_(fila[m.columnaValor - 1])).filter(Boolean);
    if (!etiquetas.length || etiquetas.some(e => e !== campo.etiqueta)) {
      throw rechazo_('grilla', 'La fila de ' + campo.nombre + ' del estudiante no tiene la etiqueta esperada en la grilla. Avisá al Equipo EPS.');
    }
  });
}

function nivelDeFila_(valores, fila) {
  for (let r = fila; r >= 0; r--) if (/^EPS\d+$/.test(norm_(valores[r][0]))) return norm_(valores[r][0]);
  return '';
}

// ================================================================ reglas de cada celda

/**
 * Situación de la Calificación: 'no cursa' (vacía o NO CURSA: no aparece), 'pendiente' (CURSA), 'valorada'
 * (escala docente del período), 'administrativo' (ACREDITÓ, EQUIVALENCIA) o 'a revisar' (otro valor).
 */
function situacionCalificacion_(periodo, valor) {
  const v = norm_(valor);
  if (!v || v === CARGA.NO_CURSA) return 'no cursa';
  if (v === CARGA.PENDIENTE) return 'pendiente';
  if (canonico_(CARGA.ESCALA_DOCENTE[periodo], valor)) return 'valorada';
  if (canonico_(CARGA.ADMINISTRATIVOS, valor)) return 'administrativo';
  return 'a revisar';
}

function calificacionEditable_(situacion) { return situacion === 'pendiente' || situacion === 'valorada'; }

/** Qué puede editar el docente según la Calificación: los complementarios siempre que curse; Calificación, si es editable. */
function comprobarReglas_(periodo, campo, calificacion) {
  const situacion = situacionCalificacion_(periodo, calificacion);
  if (situacion === 'no cursa') throw rechazo_('cursada', 'El estudiante ya no cursa esta materia según la grilla.');
  if (campo.id === 'calificacion' && !calificacionEditable_(situacion)) {
    throw rechazo_('calificacion', situacion === 'administrativo' ? 'Esta calificación la establece el Equipo EPS.' : 'La calificación tiene un valor a revisar con el Equipo EPS.');
  }
}

/** Valor a escribir y el que mostrará la celda. Rechaza lo que el docente no puede establecer. */
function normalizarNuevo_(campo, periodo, valor) {
  const crudo = txt_(valor);
  if (campo === 'calificacion') {
    if (norm_(crudo) === CARGA.PENDIENTE) return { escrito: CARGA.PENDIENTE, visto: CARGA.PENDIENTE }; // «Dejar pendiente»
    const canon = canonico_(CARGA.ESCALA_DOCENTE[periodo], crudo);
    if (!canon) throw rechazo_('valor', '"' + crudo + '" no es una calificación que el docente pueda cargar en ' + periodo + '.');
    return { escrito: canon, visto: canon };
  }
  if (campo === 'calidad' || campo === 'participacion') {
    if (!crudo) return { escrito: '', visto: '' };
    if (!CARGA.OPCIONES[campo].includes(crudo)) throw rechazo_('valor', '"' + crudo + '" no es una de las opciones de ' + campoPorId_(campo).nombre + '.');
    return { escrito: crudo, visto: crudo };
  }
  if (campo === 'observaciones') {
    const texto = crudo.replace(/\s+/g, ' ');
    if (texto.length > CARGA.MAX_OBSERVACIONES) throw rechazo_('valor', 'La observación supera ' + CARGA.MAX_OBSERVACIONES + ' caracteres.');
    // El apóstrofo inicial hace que Sheets guarde texto: ni fórmulas, ni fechas, ni números interpretados.
    return { escrito: texto ? "'" + texto : '', visto: texto };
  }
  throw rechazo_('pedido', 'Campo desconocido.');
}

function campoPorId_(id) {
  const campo = CARGA.CAMPOS.find(c => c.id === id);
  if (!campo) throw rechazo_('pedido', 'Campo desconocido.');
  return campo;
}

// ================================================================ consulta

/** Estudiantes que cursan la materia (Calificación no vacía ni NO CURSA), de cualquier grupo, con lo que pueden editar. */
function leerMateria_(valores, estructura, materia, periodo) {
  const estudiantes = [];
  const repetidos = {};
  for (let r = 0; r < valores.length; r++) { const d = dni_(valores[r][2]); if (d) repetidos[d] = (repetidos[d] || 0) + 1; }
  for (let r = 2; r < valores.length; r++) {
    const dni = dni_(valores[r][2]);
    if (!dni) continue;
    const situacion = situacionCalificacion_(periodo, valores[r][materia.columnaValor]);
    if (situacion === 'no cursa') continue;
    const campos = {};
    CARGA.CAMPOS.forEach(c => { campos[c.id] = txt_((valores[r + c.desplazamiento] || [])[materia.columnaValor]); });
    let problema = repetidos[dni] > 1 ? 'El DNI figura más de una vez en la grilla. Avisá al Equipo EPS.' : '';
    if (!problema) { try { validarBloque_(valores, estructura, r); } catch (e) { problema = e.message; } }
    const editable = {};
    CARGA.CAMPOS.forEach(c => { editable[c.id] = !problema && (c.id !== 'calificacion' || calificacionEditable_(situacion)); });
    estudiantes.push({
      dni, nombre: [txt_(valores[r][0]), txt_(valores[r][1])].filter(Boolean).join(', '),
      grupo: nivelDeFila_(valores, r), situacion, campos, editable, problema
    });
  }
  return estudiantes;
}

function contar_(estudiantes) {
  return ['valorada', 'pendiente', 'administrativo', 'a revisar'].reduce((c, s) => { c[s] = estudiantes.filter(x => x.situacion === s).length; return c; }, {});
}

/**
 * Lo que el docente puede ver y cargar ahora: los períodos con ventana abierta y, en cada uno, sus materias
 * asignadas con los estudiantes que las cursan. servicios: { reloj, datos, valoresHoja(fuente, orientación) }.
 * Cada período trae su fuente: el guardado la exige de vuelta y se rechaza si cambió.
 */
function consultarDocente_(servicios, correoDocente) {
  const correo = correo_(correoDocente);
  const ahora = servicios.reloj();
  const config = configuracion_(servicios.datos(), ahora);
  const propias = config.asignaciones.filter(a => a.correo === correo);
  if (!propias.length) return { estado: 'rechazado', codigo: 'asignacion', motivo: 'Tu cuenta no tiene materias asignadas. Si es un error, avisá al Equipo EPS.' };
  if (config.problemas.length) return { estado: 'rechazado', codigo: 'configuracion', motivo: 'La carga no está disponible por un problema de configuración. Avisá al Equipo EPS.' };
  const periodos = CARGA.PERIODOS.map(n => config.periodos[n]).filter(p => p.estado === 'abierta');
  return {
    estado: 'ok', correo, docente: propias[0].docente, al: ahora.toISOString(),
    periodos: periodos.map(p => {
      const base = { periodo: p.periodo, fuente: p.fuente, desde: p.ventana.desde, hasta: p.ventana.hasta };
      if (p.problemas.length) return Object.assign(base, { problema: 'La carga de ' + p.periodo + ' no está disponible por un problema de configuración. Avisá al Equipo EPS.', materias: [] });
      const hojas = {};
      const hoja = orientacion => {
        if (!hojas[orientacion]) {
          try {
            const valores = servicios.valoresHoja(p.fuente, orientacion);
            hojas[orientacion] = { valores, estructura: estructuraGrilla_(valores, p.periodo) };
          } catch (e) { hojas[orientacion] = { problema: e.rechazo ? e.message : 'No se pudo leer la hoja ' + orientacion + ' de la grilla de ' + p.periodo + '.' }; }
        }
        return hojas[orientacion];
      };
      return Object.assign(base, {
        cierre: p.ventana.cierre.toISOString(),
        escala: CARGA.ESCALA_DOCENTE[p.periodo].slice(),
        opciones: { calidad: CARGA.OPCIONES.calidad.slice(), participacion: CARGA.OPCIONES.participacion.slice() },
        materias: propias.map(a => {
          const item = { orientacion: a.orientacion, nivel: a.nivel, materia: a.materia, clave: a.clave, estudiantes: [] };
          if (a.ambigua) return Object.assign(item, { problema: 'La materia tiene una asignación ambigua. Avisá al Equipo EPS.' });
          const h = hoja(a.orientacion);
          if (h.problema) return Object.assign(item, { problema: h.problema });
          const materia = h.estructura.materias.find(m => m.clave === a.clave);
          if (!materia) return Object.assign(item, { problema: 'La materia no figura en la grilla de ' + p.periodo + '. Avisá al Equipo EPS.' });
          if (materia.nivel !== a.nivel) return Object.assign(item, { problema: 'La asignación tiene un nivel distinto del de la grilla. Avisá al Equipo EPS.' });
          const estudiantes = leerMateria_(h.valores, h.estructura, materia, p.periodo);
          return Object.assign(item, { materia: materia.nombre, estudiantes, conteo: contar_(estudiantes) });
        })
      });
    })
  };
}

// ================================================================ Registro

function registroValidar_(registro) {
  const enc = registro.getRange(1, 1, 1, CARGA.REGISTRO.length).getDisplayValues()[0];
  if (CARGA.REGISTRO.some((c, i) => norm_(enc[i]) !== norm_(c))) {
    throw rechazo_('registro', 'El Registro no tiene las columnas esperadas. Avisá al Equipo EPS.');
  }
}

/** Filas (1-based) del Registro con esta operación. Se lee la columna entera: el Equipo EPS puede filtrar u ordenar. */
function registroFilasDe_(registro, operacion) {
  const ultima = registro.getLastRow();
  if (ultima < 2) return [];
  const ops = registro.getRange(2, COL_REGISTRO['Operación'] + 1, ultima - 1, 1).getDisplayValues();
  const filas = [];
  ops.forEach((f, i) => { if (txt_(f[0]) === operacion) filas.push(i + 2); });
  return filas;
}

function registroLeer_(registro, fila) {
  const v = registro.getRange(fila, 1, 1, CARGA.REGISTRO.length).getDisplayValues()[0];
  const obj = { fila };
  CARGA.REGISTRO.forEach((c, i) => { obj[c] = txt_(v[i]); });
  obj['Intentos'] = Number(obj['Intentos']) || 1;
  return obj;
}

// Cada texto con apóstrofo inicial: Sheets lo guarda tal cual (sin convertir "7" en número ni "1/2" en fecha).
function registroArreglo_(datos) {
  return CARGA.REGISTRO.map(c => {
    const v = datos[c];
    if (v === undefined || v === null || v === '') return '';
    return typeof v === 'string' ? "'" + v : v;
  });
}

function registroAgregar_(registro, datos) {
  registro.appendRow(registroArreglo_(datos));
  const ultima = registro.getLastRow();
  if (txt_(registro.getRange(ultima, COL_REGISTRO['Operación'] + 1).getDisplayValue()) === datos['Operación']) return ultima;
  const filas = registroFilasDe_(registro, datos['Operación']);
  if (filas.length !== 1) throw new Error('No se pudo ubicar la fila nueva del Registro.');
  return filas[0];
}

/** Ubica de nuevo la fila de la operación si se movió (el Equipo EPS pudo ordenar o insertar filas). */
function registroUbicar_(registro, fila, operacion) {
  if (txt_(registro.getRange(fila, COL_REGISTRO['Operación'] + 1).getDisplayValue()) === operacion) return fila;
  const filas = registroFilasDe_(registro, operacion);
  if (filas.length !== 1) throw new Error('No se pudo ubicar la fila de la operación en el Registro.');
  return filas[0];
}

function registroReescribir_(registro, fila, datos) {
  const f = registroUbicar_(registro, fila, datos['Operación']);
  registro.getRange(f, 1, 1, CARGA.REGISTRO.length).setValues([registroArreglo_(datos)]);
  return f;
}

function registroMarcar_(registro, fila, operacion, estado, detalle) {
  const f = registroUbicar_(registro, fila, operacion);
  registro.getRange(f, COL_REGISTRO['Estado'] + 1, 1, 2).setValues([["'" + estado, detalle ? "'" + detalle : '']]);
}

// ================================================================ guardado de un campo

/**
 * Lectura mínima del destino dentro del lock, en dos llamadas: encabezado de la materia y bloque del estudiante,
 * de la columna C (DNI) a la de valor. Comprueba DNI y todas las etiquetas del bloque en ese tramo contra la lectura
 * completa. Devuelve los cuatro valores de la materia, o null si la grilla cambió de forma desde esa lectura.
 */
function leerDestino_(grilla, destino) {
  const enc = grilla.getRange(1, destino.columnaValor, 2, 1).getDisplayValues();
  const ancho = destino.columnaValor - 1; // columnas C … valor (1-based: 3 … columnaValor + 1)
  const bloque = grilla.getRange(destino.fila + 1, 3, 4, ancho).getDisplayValues();
  const mismo = clave_(txt_(enc[1][0]) || txt_(enc[0][0])) === destino.clave && dni_(bloque[0][0]) === destino.dni &&
    bloque.every((f, i) => (i === 0 || !dni_(f[0])) && etiquetasTramo_(f, destino.indices).join('\u0001') === destino.etiquetas[i]);
  return mismo ? bloque.map(f => txt_(f[ancho - 1])) : null;
}

// Etiquetas de las materias del tramo C … valor de una fila (índices relativos a la columna C).
function etiquetasTramo_(fila, indices) { return indices.map(i => txt_(fila[i])); }

function destinoDe_(valores, estructura, materia, fila, campo, dni) {
  const indices = estructura.materias.filter(m => m.columnaValor <= materia.columnaValor).map(m => m.columnaValor - 3);
  return {
    fila, filaCampo: fila + campo.desplazamiento, columnaValor: materia.columnaValor, clave: materia.clave, dni, indices,
    etiquetas: CARGA.CAMPOS.map(c => etiquetasTramo_(valores[fila + c.desplazamiento].slice(2), indices).join('\u0001'))
  };
}

/**
 * Guarda un campo. pedido: { operacion, periodo, fuente, orientacion, materia (clave), dni, campo, esperado, nuevo }.
 * ctx: { config, grilla, registro, lock, flush, reloj, periodo, fuente, orientacion, correo, cierre }.
 *
 * Respuestas (todas traen 'actual' cuando se pudo leer la celda):
 *  - 'guardado': esta operación escribió y se confirmó (fila 'aplicado'); la celda muestra el valor nuevo;
 *  - 'coincide': la celda muestra el valor pedido, pero no se puede confirmar que lo escribió este pedido (ya
 *    estaba, otra persona lo puso, o la escritura tuvo un error); nada que reintentar;
 *  - 'conflicto': la celda no tiene lo que vio el docente, o cambió después de escribir;
 *  - 'reintentar': no se intentó escribir la grilla; se puede repetir con la misma operación;
 *  - 'incierto': no se pudo confirmar qué pasó. Repetir con la misma operación nunca vuelve a escribir: sólo
 *    puede aclarar el resultado observando la celda.
 * Los rechazos se lanzan (rechazo_) y siempre ocurren antes de intentar escribir la grilla. Una excepción antes
 * del intento es 'reintentar'; después, 'incierto' (ver guardarDocente_).
 *
 * Estados del Registro: 'pendiente' (anotada, resultado sin observar) · 'aplicado' (confirmada) · 'inferido' (la
 * celda mostró el valor nuevo, sin poder atribuirlo a esta operación) · 'no aplicado' (nunca se intentó escribir;
 * un reintento reutiliza la fila) · 'incierto' (se intentó escribir con resultado ambiguo y la celda no muestra el
 * valor nuevo) · 'revisar' (después de escribir, la celda muestra otra cosa o la grilla cambió de forma: el Equipo
 * EPS debe mirar la celda que indica el detalle).
 */
function guardarCampo_(ctx, pedido) {
  const campo = campoPorId_(pedido.campo);
  const nuevo = normalizarNuevo_(campo.id, ctx.periodo, pedido.nuevo);

  // 1. Ubicación con una lectura completa, fuera del lock.
  const valores = ctx.grilla.getDataRange().getDisplayValues();
  const estructura = estructuraGrilla_(valores, ctx.periodo);
  const materia = materiaPorClave_(estructura, pedido.materia);
  const asignacion = autorizarMateria_(ctx.config, ctx.correo, ctx.orientacion, materia);
  const fila = filaEstudiante_(valores, estructura, pedido.dni);
  comprobarReglas_(ctx.periodo, campo, valores[fila][materia.columnaValor]);
  const destino = destinoDe_(valores, estructura, materia, fila, campo, pedido.dni);
  const datos = {
    'Operación': pedido.operacion, 'Correo': ctx.correo, 'Docente': asignacion.docente, 'Período': ctx.periodo, 'Fuente': ctx.fuente,
    'Orientación': ctx.orientacion, 'Nivel': materia.nivel, 'Materia': materia.nombre, 'DNI': pedido.dni, 'Campo': campo.nombre,
    'Estudiante': [txt_(valores[fila][0]), txt_(valores[fila][1])].filter(Boolean).join(', '), 'Nuevo': nuevo.visto
  };

  // 2. Registro, comprobación, escritura y observación, dentro de un lock corto.
  if (!ctx.lock.tryLock(CARGA.ESPERA_LOCK_MS)) return { estado: 'reintentar', motivo: 'Hay muchos guardados al mismo tiempo. Se reintenta.' };
  const inicioLock = Date.now();
  let r = null, error = null;
  try {
    r = guardarConLock_(ctx, pedido, campo, nuevo, destino, datos, materia);
  } catch (e) { error = e; }
  // Todo lo escrito bajo el lock (Registro y grilla) se confirma antes de liberarlo, en todos los caminos.
  const confirmado = confirmarYLiberar_(ctx.flush, ctx.lock);
  if (error) throw error;
  // Sin confirmación, la marca 'aplicado' pudo perderse: no se responde «guardado». El reintento resuelve sin escribir.
  if (!confirmado && r.estado === 'guardado') r = { estado: 'incierto', actual: r.actual, motivo: 'No se pudo confirmar el Registro. Se reintenta.' };
  r.lockMs = Date.now() - inicioLock;
  return r;
}

/**
 * Confirma con flush los cambios pendientes de Sheets y recién entonces libera el lock, como pide LockService.
 * El lock se libera aunque el flush falle. Devuelve si la confirmación tuvo éxito.
 */
function confirmarYLiberar_(flush, lock) {
  let confirmado = true;
  try { flush(); } catch (e) { confirmado = false; }
  lock.releaseLock();
  return confirmado;
}

function guardarConLock_(ctx, pedido, campo, nuevo, destino, datos, materia) {
  const op = pedido.operacion;
  const esperado = txt_(pedido.esperado);
  registroValidar_(ctx.registro);
  const filas = registroFilasDe_(ctx.registro, op);
  if (filas.length > 1) throw rechazo_('registro', 'La operación figura más de una vez en el Registro. Avisá al Equipo EPS.');
  const previa = filas.length ? registroLeer_(ctx.registro, filas[0]) : null;
  // Identidad de la operación: el reintento debe ser el mismo pedido, campo por campo (normalizados), incluido lo
  // que vio el docente. La fila sólo se crea o reutiliza cuando la celda coincide con «esperado», así que su
  // «Anterior» es siempre el «esperado» del pedido original.
  if (previa && (clave_(previa['Materia']) !== clave_(datos['Materia']) || previa['Anterior'] !== esperado ||
      ['Correo', 'Período', 'Fuente', 'Orientación', 'DNI', 'Campo', 'Nuevo'].some(c => previa[c] !== datos[c]))) {
    throw rechazo_('operacion', 'El identificador de la operación ya se usó para otro cambio. Volvé a abrir la carga.');
  }
  // Mientras se esperaba el lock pudieron cambiar ventana, fuente, asignación o configuración: se validan de nuevo.
  if (ctx.revalidar) ctx.revalidar(materia);

  const vista = leerDestino_(ctx.grilla, destino);
  if (!vista) return { estado: 'reintentar', motivo: 'La grilla cambió mientras se guardaba. Se reintenta.' };
  const actual = vista[campo.desplazamiento];

  if (previa) {
    const estado = previa['Estado'];
    // Ya hubo un intento de escritura: nunca se repite (no se pisa una corrección posterior). Sólo se informa.
    if (estado === 'aplicado') return actual === nuevo.visto ? { estado: 'guardado', actual } : { estado: 'conflicto', actual };
    if (estado === 'inferido' || estado === 'revisar') return actual === nuevo.visto ? { estado: 'coincide', actual } : { estado: 'conflicto', actual };
    if (estado === 'pendiente' || estado === 'incierto') {
      // El valor nuevo en la celda no prueba quién lo puso: queda 'inferido' y no se responde «guardado».
      if (actual === nuevo.visto) {
        registroMarcar_(ctx.registro, previa.fila, op, 'inferido', 'al reintentar, la celda mostraba el valor nuevo; no se puede confirmar que lo escribió esta operación');
        return { estado: 'coincide', actual };
      }
      if (estado === 'pendiente') registroMarcar_(ctx.registro, previa.fila, op, 'incierto', 'al reintentar, la celda mostraba "' + actual + '": la escritura pudo no aplicarse, o aplicarse y cambiar después');
      return { estado: 'incierto', actual, motivo: 'No se pudo confirmar el guardado. La grilla muestra el valor actual.' };
    }
    // 'no aplicado': nunca se intentó escribir. Si la celda ya tiene el valor, lo puso otra persona.
  }
  if (actual === nuevo.visto) return { estado: 'coincide', actual }; // ya estaba: no hay modificación que registrar
  if (actual !== esperado) return { estado: 'conflicto', actual };
  comprobarReglas_(ctx.periodo, campo, vista[0]); // la Calificación pudo cambiar desde la lectura completa
  if (ctx.reloj() >= ctx.cierre) throw rechazo_('ventana', 'La carga de ' + ctx.periodo + ' cerró.');

  // El Registro precede a la grilla: ninguna escritura queda sin su fila.
  const fila = Object.assign({}, datos, { 'Fecha': ctx.reloj(), 'Estado': 'pendiente', 'Detalle': '', 'Intentos': previa ? previa['Intentos'] + 1 : 1, 'Anterior': actual });
  let filaRegistro;
  let movida = false;
  try {
    filaRegistro = previa ? registroReescribir_(ctx.registro, previa.fila, fila) : registroAgregar_(ctx.registro, fila);
    ctx.flush();
    // Última comprobación, inmediatamente antes de escribir: entre esta lectura y la escritura no hay otra llamada
    // al servicio. Cualquier cambio en el bloque desde la primera comprobación cancela el intento.
    const ultima = leerDestino_(ctx.grilla, destino);
    movida = !ultima || ultima.join('\u0001') !== vista.join('\u0001');
    if (!movida && ctx.reloj() >= ctx.cierre) throw rechazo_('ventana', 'La carga de ' + ctx.periodo + ' cerró.');
  } catch (e) {
    anularIntento_(ctx.registro, op, e.rechazo ? e.message : 'falló la anotación en el Registro');
    throw e;
  }
  if (movida) {
    anularIntento_(ctx.registro, op, 'la grilla cambió entre la comprobación y la escritura');
    return { estado: 'reintentar', motivo: 'La grilla cambió mientras se guardaba. Se reintenta.' };
  }

  ctx.escribiendo = true; // desde acá hubo un intento: un error deja el resultado «incierto»
  let error = null;
  try {
    ctx.grilla.getRange(destino.filaCampo + 1, destino.columnaValor + 1).setValue(nuevo.escrito);
    ctx.flush();
  } catch (e) { error = e; } // resultado ambiguo: decide la observación, sin concluir más de lo que muestra
  const resultado = clasificarEscritura_(ctx, destino, campo, actual, nuevo.visto, leerDestino_(ctx.grilla, destino), error);
  // Sólo esta ejecución sabe si la escritura se confirmó: ante una falla pasajera, se intenta marcar dos veces.
  for (let intento = 1; ; intento++) {
    try {
      registroMarcar_(ctx.registro, filaRegistro, op, resultado.estado, resultado.detalle);
      return resultado.respuesta;
    } catch (e) {
      // La fila queda 'pendiente': el reintento con la misma operación sólo observa y la resuelve, sin escribir.
      if (intento === 2) return { estado: 'incierto', motivo: 'No se pudo completar el Registro. Se reintenta.' };
    }
  }
}

/**
 * La fila de este intento queda 'no aplicado': la grilla no se tocó y un reintento puede volver a intentarlo.
 * Si ni esto se puede escribir, la fila sigue 'pendiente' y el reintento, sin poder saberlo, no escribirá.
 */
function anularIntento_(registro, op, motivo) {
  try {
    const f = registroFilasDe_(registro, op);
    if (f.length === 1) registroMarcar_(registro, f[0], op, 'no aplicado', 'no se llegó a escribir: ' + motivo);
  } catch (x) { /* queda 'pendiente' */ }
}

/**
 * Estado del Registro y respuesta, mirando la celda después de intentar escribir. Sólo se confirma ('aplicado')
 * lo que Sheets aceptó sin error y la celda mostró enseguida en su lugar. Lo demás no se adivina: el valor nuevo
 * visto sin esa confirmación queda 'inferido', y un error con el valor anterior queda 'incierto' (la escritura pudo
 * no aplicarse, o aplicarse y ser revertida antes de esta lectura).
 */
function clasificarEscritura_(ctx, destino, campo, anterior, nuevo, despues, error) {
  const msg = error ? ': ' + error.message : '';
  if (despues) {
    const visto = despues[campo.desplazamiento];
    if (visto === nuevo && !error) return { estado: 'aplicado', detalle: '', respuesta: { estado: 'guardado', actual: nuevo } };
    if (visto === nuevo) return { estado: 'inferido', detalle: 'el servicio informó un error y la celda muestra el valor nuevo' + msg, respuesta: { estado: 'coincide', actual: nuevo } };
    if (visto === anterior && error) {
      return { estado: 'incierto', detalle: 'el servicio informó un error y la celda muestra el valor anterior: la escritura pudo no aplicarse, o aplicarse y ser revertida' + msg,
        respuesta: { estado: 'incierto', actual: visto, motivo: 'No se pudo confirmar el guardado. La grilla muestra el valor actual.' } };
    }
    if (visto === anterior) {
      return { estado: 'revisar', detalle: 'después de escribir, la celda volvió a mostrar el valor anterior: otra persona la cambió al mismo tiempo, o la grilla no admitió el valor', respuesta: { estado: 'conflicto', actual: visto } };
    }
    return { estado: 'revisar', detalle: 'después de escribir, la celda muestra "' + visto + '": otra persona la cambió al mismo tiempo, o la grilla altera el valor' + msg, respuesta: { estado: 'conflicto', actual: visto } };
  }
  // La grilla cambió de forma después de la última comprobación: la escritura pudo quedar en otra celda.
  let actual = null;
  try {
    const valores = ctx.grilla.getDataRange().getDisplayValues();
    const estructura = estructuraGrilla_(valores, ctx.periodo);
    const columna = materiaPorClave_(estructura, destino.clave).columnaValor;
    actual = txt_(valores[filaEstudiante_(valores, estructura, destino.dni) + campo.desplazamiento][columna]);
  } catch (e) { /* el destino ya no se puede ubicar */ }
  const celda = 'la celda ' + a1_(destino.filaCampo, destino.columnaValor) + ' de la hoja ' + ctx.orientacion;
  if (actual === nuevo) {
    return { estado: 'inferido', detalle: 'la grilla cambió de forma durante el guardado y la celda del estudiante muestra el valor nuevo; si no lo escribió esta operación, verificar ' + celda + msg,
      respuesta: { estado: 'coincide', actual } };
  }
  const detalle = 'la grilla cambió de forma durante la escritura: verificar ' + celda + ' (historial de ediciones), que pudo quedar con "' + nuevo + '"' + msg;
  return { estado: 'revisar', detalle, respuesta: actual === null ? { estado: 'incierto', motivo: 'No se pudo confirmar el guardado. Volvé a abrir la carga.' } : { estado: 'conflicto', actual } };
}

function validarPedido_(pedido) {
  if (!pedido || typeof pedido !== 'object') throw rechazo_('pedido', 'Pedido inválido.');
  if (!CARGA.OPERACION.test(txt_(pedido.operacion))) throw rechazo_('pedido', 'Falta el identificador de la operación o no es válido.');
  if (!CARGA.HOJAS_ORIGEN.includes(txt_(pedido.orientacion))) throw rechazo_('pedido', 'Orientación desconocida.');
  ['periodo', 'fuente', 'materia', 'dni', 'campo'].forEach(k => { if (!txt_(pedido[k])) throw rechazo_('pedido', 'Pedido incompleto: falta ' + k + '.'); });
  if (pedido.esperado === undefined || pedido.nuevo === undefined) throw rechazo_('pedido', 'Pedido incompleto: falta el valor esperado o el nuevo.');
}

/**
 * Punto de entrada de cada guardado, para un docente ya identificado (la identificación es de #28).
 * servicios: { reloj, datos, hoja(fuente, orientación), registro, lock, flush }, leídos en cada operación.
 * Nunca lanza: devuelve 'guardado', 'coincide', 'conflicto', 'reintentar', 'incierto' o 'rechazado'. Un rechazo
 * garantiza que este pedido no intentó escribir; si la operación ya figuraba en el Registro, «previo» informa su estado.
 */
function guardarDocente_(servicios, correoDocente, pedido) {
  const correo = correo_(correoDocente);
  let ctx = null;
  try {
    validarPedido_(pedido);
    const config = configuracion_(servicios.datos(), servicios.reloj());
    const p = periodoHabilitado_(config, pedido.periodo, pedido.fuente);
    ctx = {
      config, grilla: servicios.hoja(p.fuente, pedido.orientacion), registro: servicios.registro(), lock: servicios.lock(),
      flush: servicios.flush, reloj: servicios.reloj, periodo: p.periodo, fuente: p.fuente, orientacion: pedido.orientacion,
      correo, cierre: p.ventana.cierre, escribiendo: false,
      // Dentro del lock: la configuración vigente en ese momento, con la misma fuente que vio el docente.
      revalidar: materia => {
        const vigente = configuracion_(servicios.datos(), servicios.reloj());
        const q = periodoHabilitado_(vigente, p.periodo, p.fuente);
        autorizarMateria_(vigente, correo, pedido.orientacion, materia);
        ctx.config = vigente;
        ctx.cierre = q.ventana.cierre;
      }
    };
    return guardarCampo_(ctx, { operacion: txt_(pedido.operacion), materia: txt_(pedido.materia), dni: dni_(pedido.dni), campo: pedido.campo, esperado: pedido.esperado, nuevo: pedido.nuevo });
  } catch (e) {
    if (!e.rechazo) {
      // Antes de anotar en el Registro no se escribió nada en la grilla: se puede repetir sin dudas.
      return ctx && ctx.escribiendo
        ? { estado: 'incierto', motivo: 'No se pudo confirmar el guardado. Se reintenta; si persiste, avisá al Equipo EPS.', error: e.message }
        : { estado: 'reintentar', motivo: 'No se pudo guardar. Se reintenta; si persiste, avisá al Equipo EPS.', error: e.message };
    }
    const respuesta = { estado: 'rechazado', codigo: e.rechazo.codigo, motivo: e.rechazo.motivo };
    try {
      if (CARGA.OPERACION.test(txt_(pedido && pedido.operacion))) {
        const registro = servicios.registro();
        const filas = registroFilasDe_(registro, txt_(pedido.operacion));
        if (filas.length === 1) {
          const previa = registroLeer_(registro, filas[0]);
          if (previa['Correo'] === correo) respuesta.previo = { estado: previa['Estado'], nuevo: previa['Nuevo'] };
        }
      }
    } catch (x) { respuesta.previo = { estado: 'desconocido' }; }
    return respuesta;
  }
}

// ================================================================ control del Equipo EPS

/**
 * Resuelve, mirando la celda, las filas 'pendiente' más viejas que edadMs: quedan así si la ejecución se cortó
 * después de anotar y el docente no reintentó. Con el valor nuevo, 'inferido'; con cualquier otro, 'incierto':
 * no se concluye ni que se aplicó ni que no. Nunca escribe la grilla. Las marcas se confirman (flush) antes de
 * liberar el lock; si eso falla, se informa en errores. servicios: { reloj, hoja, registro, lock, flush }.
 */
function resolverPendientes_(servicios, edadMs) {
  const registro = servicios.registro();
  const lock = servicios.lock();
  if (!lock.tryLock(CARGA.ESPERA_LOCK_MS)) throw new Error('No se pudo tomar el lock. Volvé a intentarlo.');
  const resultado = { revisadas: 0, inferido: 0, incierto: 0, errores: [] };
  try {
    registroValidar_(registro);
    const ultima = registro.getLastRow();
    if (ultima < 2) return resultado;
    const filas = registro.getRange(2, 1, ultima - 1, CARGA.REGISTRO.length).getValues();
    const ahora = servicios.reloj();
    filas.forEach((f, i) => {
      if (txt_(f[COL_REGISTRO['Estado']]) !== 'pendiente') return;
      const fecha = f[COL_REGISTRO['Fecha']];
      if (Object.prototype.toString.call(fecha) === '[object Date]' && ahora - fecha < edadMs) return;
      const r = registroLeer_(registro, i + 2);
      resultado.revisadas++;
      try {
        const campo = CARGA.CAMPOS.find(c => c.nombre === r['Campo']);
        const valores = servicios.hoja(r['Fuente'], r['Orientación']).getDataRange().getDisplayValues();
        const estructura = estructuraGrilla_(valores, r['Período']);
        const columna = materiaPorClave_(estructura, clave_(r['Materia'])).columnaValor;
        const actual = txt_(valores[filaEstudiante_(valores, estructura, r['DNI']) + campo.desplazamiento][columna]);
        const estado = actual === r['Nuevo'] ? 'inferido' : 'incierto';
        registroMarcar_(registro, r.fila, r['Operación'], estado, 'resuelto por control: la celda ' + (estado === 'inferido' ?
          'mostraba el valor nuevo; no se puede confirmar que lo escribió esta operación' : 'mostraba "' + actual + '": la escritura pudo no aplicarse, o aplicarse y cambiar después'));
        resultado[estado]++;
      } catch (e) { resultado.errores.push('Registro, fila ' + r.fila + ': ' + e.message); }
    });
  } finally {
    if (!confirmarYLiberar_(servicios.flush, lock)) {
      resultado.errores.push('No se pudo confirmar en el Registro las marcas de esta ejecución: volvé a ejecutarla.');
    }
  }
  return resultado;
}

/**
 * Problemas que el Equipo EPS debe resolver antes o durante una ventana: configuración, asignaciones que no
 * coinciden con la grilla, materias con cursantes y sin docente, DNI repetidos y filas del Registro sin resolver.
 * servicios: { reloj, datos, valoresHoja(fuente, orientación), registro }.
 */
function verificarConfiguracion_(servicios) {
  const ahora = servicios.reloj();
  const config = configuracion_(servicios.datos(), ahora);
  const problemas = config.problemas.concat(config.advertencias);
  const ventanas = [];
  CARGA.PERIODOS.forEach(nombre => {
    const p = config.periodos[nombre];
    p.problemas.forEach(x => problemas.push(nombre + ': ' + x));
    if (!p.ventana || p.estado === 'cerrada') return;
    ventanas.push(nombre + ': ' + p.ventana.desde + ' a ' + p.ventana.hasta + (p.ventana.cierre ? ', cierra ' + p.ventana.diaCierre + ' ' + CARGA.HORA_CIERRE : '') + ' (' + p.estado + ')');
    if (!p.fuente) return;
    CARGA.HOJAS_ORIGEN.forEach(orientacion => {
      let valores, estructura;
      try { valores = servicios.valoresHoja(p.fuente, orientacion); estructura = estructuraGrilla_(valores, nombre); } catch (e) { return problemas.push(nombre + ' · ' + orientacion + ': ' + e.message); }
      const dnis = {};
      for (let r = 2; r < valores.length; r++) { const d = dni_(valores[r][2]); if (d) dnis[d] = (dnis[d] || 0) + 1; }
      Object.keys(dnis).filter(d => dnis[d] > 1).forEach(d => problemas.push(nombre + ' · ' + orientacion + ': el DNI ' + d + ' figura ' + dnis[d] + ' veces.'));
      config.asignaciones.filter(a => a.orientacion === orientacion).forEach(a => {
        const m = estructura.materias.find(x => x.clave === a.clave);
        if (!m) problemas.push(nombre + ' · ' + orientacion + ': la materia asignada "' + a.materia + '" no figura en la grilla (Asignaciones, fila ' + a.fila + ').');
        else if (m.nivel !== a.nivel) problemas.push(nombre + ' · ' + orientacion + ': ' + a.materia + ' es de ' + m.nivel + ' en la grilla y de ' + a.nivel + ' en Asignaciones (fila ' + a.fila + ').');
      });
      estructura.materias.forEach(m => {
        if (asignacionDe_(config, orientacion, m.clave)) return;
        const cursan = leerMateria_(valores, estructura, m, nombre).length;
        if (cursan) problemas.push(nombre + ' · ' + orientacion + ' ' + m.nivel + ' · ' + m.nombre + ': ' + cursan + ' cursantes y ningún docente asignado.');
      });
    });
  });
  try {
    const registro = servicios.registro();
    registroValidar_(registro);
    const ultima = registro.getLastRow();
    if (ultima >= 2) {
      registro.getRange(2, 1, ultima - 1, CARGA.REGISTRO.length).getDisplayValues().forEach((f, i) => {
        const estado = txt_(f[COL_REGISTRO['Estado']]);
        if (estado === 'pendiente' || estado === 'revisar') problemas.push('Registro, fila ' + (i + 2) + ': ' + estado + (f[COL_REGISTRO['Detalle']] ? ' · ' + f[COL_REGISTRO['Detalle']] : ''));
      });
    }
  } catch (e) { problemas.push(e.message); }
  return { ventanas, problemas };
}

// ================================================================ utilidades (copias mínimas de las del generador)

/** Generador: normalizarPeriodo_ (paridad). Variantes escritas a mano → nombre canónico del período. */
function normalizarPeriodo_(valor) {
  const crudo = txt_(valor);
  const t = crudo.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[º°]/g, '').replace(/\s+/g, ' ').trim();
  if (['1 bimestre', '1er bimestre', 'primer bimestre'].includes(t)) return '1º bimestre';
  if (['1 cuatrimestre', '1er cuatrimestre', 'primer cuatrimestre'].includes(t)) return '1º cuatrimestre';
  if (['3 bimestre', '3er bimestre', 'tercer bimestre'].includes(t)) return '3º bimestre';
  if (['2 cuatrimestre', '2do cuatrimestre', 'segundo cuatrimestre'].includes(t)) return '2º cuatrimestre';
  if (['anual', 'calificacion anual'].includes(t)) return 'Anual';
  return crudo;
}

/** Generador: extraerSpreadsheetId_ (paridad), devolviendo '' en lugar de lanzar. */
function idPlanilla_(entrada) {
  const texto = txt_(entrada);
  const url = texto.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (url) return url[1];
  return /^[a-zA-Z0-9-_]{20,}$/.test(texto) ? texto : '';
}

function normalizarHojaOrigen_(valor) { const h = norm_(valor).replace(/\s+/g, ''); return CARGA.HOJAS_ORIGEN.includes(h) ? h : ''; }
function normalizarNivel_(valor) { const m = norm_(valor).replace(/\s+/g, '').match(/^(?:EPS)?(\d+)$/); return m ? 'EPS' + Number(m[1]) : ''; }
function correo_(valor) { return txt_(valor).toLowerCase(); }
function canonico_(lista, valor) { return (lista || []).find(v => norm_(v) === norm_(valor)) || null; }
function etiqueta_(valor) { return norm_(valor).replace(/:$/, ''); }
function clave_(nombre) { return norm_(nombre).replace(/[^A-Z0-9]/g, ''); }
function dni_(valor) { return txt_(valor).replace(/\D/g, ''); }
function norm_(valor) { return txt_(valor).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim(); }
function txt_(valor) { return valor === null || valor === undefined ? '' : String(valor).trim(); }
function a1_(fila0, col0) {
  let letras = '';
  for (let n = col0 + 1; n > 0; n = Math.floor((n - 1) / 26)) letras = String.fromCharCode(65 + (n - 1) % 26) + letras;
  return letras + (fila0 + 1);
}
