#!/usr/bin/env node
/**
 * Pruebas locales de la carga docente (#26): núcleo (src/carga/Carga_Nucleo.gs) y adaptador (Carga_Google.gs).
 * Datos ficticios. Sólo Node.js (no requiere Windows ni tar.exe).
 *
 *   node tests/carga.js [corridas]     corridas de la simulación aleatoria (30 por defecto)
 *
 * La carga y src/Generador_EPS.gs (sin cambios) se cargan en contextos separados, como dos proyectos de Apps Script.
 * El generador se usa para la paridad (estructura, valores, celdas de Control) y para comprobar que lo que guarda la
 * carga sale en el boletín. Los dobles de Sheets inyectan fallas de resultado ambiguo y ediciones del Equipo EPS en
 * puntos exactos del guardado. Nada de esto sustituye el piloto en Google real.
 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const RAIZ = path.resolve(__dirname, "..");
const leer = p => fs.readFileSync(path.join(RAIZ, p), "utf8");
const CORRIDAS = Number(process.argv[2]) || 30;

function contexto(archivos, globales) {
  const ctx = Object.assign({ console }, globales || {});
  vm.createContext(ctx);
  archivos.forEach(a => vm.runInContext(leer(a), ctx, { filename: a }));
  return ctx;
}
const NUCLEO = "src/carga/Carga_Nucleo.gs";
const C = contexto([NUCLEO]);
const G = contexto(["src/Generador_EPS.gs"]);
const enG = (expr, datos) => { G.__d = JSON.parse(JSON.stringify(datos === undefined ? null : datos)); return vm.runInContext(expr, G); };
const CARGA = vm.runInContext("CARGA", C);
const COL = vm.runInContext("COL_REGISTRO", C);

// ================================================================ dobles de Sheets con ganchos y fallas

let SEQ = 0;
// Orden de las operaciones con efecto: 'escritura' (grilla o Registro), 'flush', 'flush-fallido', 'toma' y 'libera' (lock).
let EVENTOS = [];
let OP_ACTUAL = "";
const esFecha = x => Object.prototype.toString.call(x) === "[object Date]";

class Hoja {
  constructor(filas, datos) {
    const ancho = Math.max(1, ...filas.map(f => f.length));
    this.v = filas.map(f => Array.from({ length: ancho }, (_, i) => (f[i] === undefined || f[i] === null ? "" : f[i])));
    Object.assign(this, datos || {});
    this.ganchos = [];    // [{ evento, fn }], de un solo uso: lectura:completa, lectura:parcial, antes:setValue, despues:setValue, appendRow, setValues
    this.fallas = {};     // evento → 'fallar' | 'aplicar-y-fallar' (de un solo uso) o una lista de ellas
    this.formatos = {};   // "fila,col" → función que muestra un número
    this.escrituras = []; // setValue con efecto en la celda: { seq, fila, col, antes, valor, crudo, op }
    this.intentos = [];   // todo setValue, aunque falle: { seq, op }
    this.altas = [];      // Registro: { seq, op } por appendRow o reescritura de fila completa
  }
  disparar(evento) { const i = this.ganchos.findIndex(g => g.evento === evento); if (i >= 0) this.ganchos.splice(i, 1)[0].fn(this); }
  falla(evento) { // una falla, o una lista de fallas consecutivas
    const f = this.fallas[evento];
    if (!Array.isArray(f)) { delete this.fallas[evento]; return f; }
    const x = f.shift(); if (!f.length) delete this.fallas[evento]; return x;
  }
  ancho() { return this.v[0] ? this.v[0].length : 1; }
  asegurar(f, c) {
    while (this.v.length < f) this.v.push(Array(this.ancho()).fill(""));
    if (c > this.ancho()) this.v.forEach(r => { while (r.length < c) r.push(""); });
  }
  ver(x) { return esFecha(x) ? x.toISOString() : String(x); }
  // Como Sheets: el apóstrofo inicial fuerza texto y no se muestra; un formato numérico altera lo que se ve.
  poner(f, c, x) {
    let valor = typeof x === "string" && x.startsWith("'") ? x.slice(1) : x;
    const fmt = this.formatos[(f - 1) + "," + (c - 1)];
    if (fmt && typeof x === "string" && /^\d+$/.test(x)) valor = fmt(Number(x));
    this.v[f - 1][c - 1] = valor;
  }
  getDataRange() {
    return {
      getDisplayValues: () => { this.disparar("lectura:completa"); return this.v.map(f => f.map(x => this.ver(x))); },
      getValues: () => this.v.map(f => f.slice())
    };
  }
  getLastRow() { for (let r = this.v.length; r > 0; r--) if (this.v[r - 1].some(x => x !== "")) return r; return 0; }
  getLastColumn() { return this.ancho(); }
  getRange(f, c, nf, nc) {
    nf = nf || 1; nc = nc || 1;
    const h = this;
    const bloque = crudo => {
      h.disparar("lectura:parcial");
      return Array.from({ length: nf }, (_, i) => Array.from({ length: nc }, (_, j) => {
        const x = (h.v[f - 1 + i] || [])[c - 1 + j];
        const y = x === undefined ? "" : x;
        return crudo ? y : h.ver(y);
      }));
    };
    const rango = {
      getDisplayValues: () => bloque(false),
      getValues: () => bloque(true),
      getDisplayValue: () => bloque(false)[0][0],
      setValue: x => {
        h.intentos.push({ seq: ++SEQ, op: OP_ACTUAL });
        h.disparar("antes:setValue");
        const falla = h.falla("setValue");
        if (falla === "fallar") throw new Error("Servicio de hojas de cálculo no disponible (simulado)");
        h.asegurar(f, c);
        const antes = h.ver(h.v[f - 1][c - 1]);
        h.poner(f, c, x);
        h.escrituras.push({ seq: ++SEQ, fila: f - 1, col: c - 1, antes, valor: h.ver(h.v[f - 1][c - 1]), crudo: x, op: OP_ACTUAL });
        EVENTOS.push("escritura");
        h.disparar("despues:setValue");
        if (falla === "aplicar-y-fallar") throw new Error("Se excedió el tiempo de espera (simulado; la escritura sí se aplicó)");
      },
      setValues: m => {
        h.disparar("setValues");
        if (h.falla("setValues") === "fallar") throw new Error("Falla simulada al escribir el Registro");
        if (m.length !== nf || m.some(r => r.length !== nc)) throw new Error("setValues con dimensiones distintas del rango");
        h.asegurar(f + nf - 1, c + nc - 1);
        m.forEach((r, i) => r.forEach((x, j) => h.poner(f + i, c + j, x)));
        EVENTOS.push("escritura");
        if (nc === CARGA.REGISTRO.length) h.altas.push({ seq: ++SEQ, op: String(m[0][COL["Operación"]]).replace(/^'/, "") });
        return rango;
      },
      setFontWeight: () => rango
    };
    return rango;
  }
  appendRow(fila) {
    this.disparar("appendRow");
    const falla = this.falla("appendRow");
    if (falla === "fallar") throw new Error("Falla simulada al agregar al Registro");
    const r = this.getLastRow() + 1;
    this.asegurar(r, fila.length);
    fila.forEach((x, j) => this.poner(r, j + 1, x));
    EVENTOS.push("escritura");
    this.altas.push({ seq: ++SEQ, op: String(fila[COL["Operación"]]).replace(/^'/, "") });
    if (falla === "aplicar-y-fallar") throw new Error("Se excedió el tiempo de espera (simulado; la fila sí se agregó)");
  }
  insertarFilas(antesDe0, n) { const a = this.ancho(); this.v.splice(antesDe0, 0, ...Array.from({ length: n }, () => Array(a).fill(""))); }
  insertarColumnas(antesDe0, n) { this.v.forEach(f => f.splice(antesDe0, 0, ...Array(n).fill(""))); }
}
// alEsperar: lo que pasa mientras la ejecución espera el lock (otros guardados, cambios del Equipo EPS).
function nuevoLock() {
  return {
    tomado: false, alEsperar: null,
    tryLock() { const f = this.alEsperar; this.alEsperar = null; if (f) f(); if (this.tomado) return false; this.tomado = true; EVENTOS.push("toma"); return true; },
    releaseLock() { this.tomado = false; EVENTOS.push("libera"); }
  };
}

// ================================================================ grillas ficticias con la estructura real

const leerTsv = p => leer(p).replace(/\n$/, "").split("\n").map(l => l.split("\t"));
function prng(s) { return () => { s |= 0; s = s + 0x6D2B79F5 | 0; let t = Math.imul(s ^ s >>> 15, 1 | s); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
const EPSI_3B = "reference/fixtures/3er_bimestre_estructura_real_EPSI.tsv";
const EPSG_3B = "reference/fixtures/3er_bimestre_estructura_real_EPSG.tsv";
const EPSI_1C = "reference/fixtures/1er_cuatrimestre_estructura_real_EPSI.tsv";
const ETIQUETAS = ["Valoración", "Calidad de las producciones:", "Participación:", "Observaciones:"];

/**
 * Encabezado de un fixture real y estudiantes ficticios en bloques de cuatro filas. Cada uno cursa las materias de
 * su grupo y algunos, además, una de otro nivel. Hay CURSA, NO CURSA, ACREDITÓ, Equivalencia (como en la validación
 * real), valores no reconocidos y etiquetas vacías, como en las grillas reales.
 */
function grillaSintetica(fixture, periodo, semilla, dniBase, porNivel) {
  const azar = prng(semilla);
  const base = leerTsv(fixture);
  const enc = [base[0], base[1]];
  const est = C.estructuraGrilla_(enc, periodo);
  const ancho = enc[0].length;
  const escala = CARGA.ESCALA_DOCENTE[periodo];
  const filas = [enc[0].slice(), enc[1].slice()];
  const nombres = ["Ana", "Bruno", "Carla", "Dante", "Elena", "Facundo", "Gala", "Hugo", "Inés", "Juan", "Kiara", "Lucas"];
  const elegir = l => l[Math.floor(azar() * l.length)];
  let n = 0;
  ["EPS1", "EPS2", "EPS3", "EPS4"].forEach(grupo => {
    const marcador = Array(ancho).fill(""); marcador[0] = grupo; filas.push(marcador);
    for (let s = 0; s < (porNivel || 6); s++, n++) {
      const b = [0, 1, 2, 3].map(() => Array(ancho).fill(""));
      b[0][0] = "Ficticio " + grupo.slice(3) + String.fromCharCode(65 + s); b[0][1] = nombres[(n * 5) % nombres.length]; b[0][2] = String(dniBase + n);
      const otroNivel = azar() < 0.25 ? elegir(est.materias.filter(m => m.nivel !== grupo)) : null;
      est.materias.forEach(m => {
        ETIQUETAS.forEach((e, i) => { if (azar() > 0.03) b[i][m.columnaValor - 1] = e; });
        if (m.nivel !== grupo && m !== otroNivel) return;
        const r = azar();
        if (r < 0.05) { b[0][m.columnaValor] = "NO CURSA"; return; }
        if (r < 0.40) { b[0][m.columnaValor] = "CURSA"; return; }
        b[0][m.columnaValor] = r < 0.46 ? "ACREDITÓ" : r < 0.50 ? "Equivalencia" : r < 0.52 ? "Muy bien" : elegir(escala);
        b[1][m.columnaValor] = elegir(CARGA.OPCIONES.calidad);
        b[2][m.columnaValor] = elegir(CARGA.OPCIONES.participacion);
        if (azar() < 0.3) b[3][m.columnaValor] = "Observación ficticia";
      });
      const col = re => enc[0].findIndex(h => re.test(h));
      b[0][col(/ORIENTACI/i)] = fixture.includes("EPSG") ? "Gastronomía" : "Informática";
      if (col(/LISTO PARA EXPORTAR/i) >= 0) b[0][col(/LISTO PARA EXPORTAR/i)] = "SÍ";
      if (col(/^EXPORTADO$/i) >= 0) b[0][col(/^EXPORTADO$/i)] = "NO";
      filas.push(...b);
    }
  });
  return filas;
}

// ================================================================ escenario: administración, Control y grillas

const P1B = "1º bimestre", P1C = "1º cuatrimestre", P3B = "3º bimestre", P2C = "2º cuatrimestre", PAN = "Anual";
const ID = { [P1B]: "GRILLA_1B_ficticia_00000000001", [P1C]: "GRILLA_1C_ficticia_00000000001", [P3B]: "GRILLA_3B_ficticia_00000000001",
  [P2C]: "GRILLA_2C_ficticia_00000000001", [PAN]: "GRILLA_AN_ficticia_00000000001" };
const OTRA_2C = "GRILLA_2C_ficticia_REEMPLAZO_01";
const URL = id => "https://docs.google.com/spreadsheets/d/" + id + "/edit#gid=0";
const MATE = "docente.matematica@bue.edu.ar";
const LENGUA = "docente.lengua@bue.edu.ar";
const REEMPLAZO = "docente.reemplazo@bue.edu.ar";
const AHORA = new Date("2026-11-20T15:00:00-03:00"); // viernes, dentro de la ventana de 2C y anual

// Ventana compartida de 2C y anual: del lunes 16/11 al viernes 4/12. Sábado, domingo y dos feriados (lunes 7,
// ficticio para la prueba, y martes 8, Inmaculada Concepción): cierra el miércoles 9/12 a las 8:00.
// El 3º bimestre ya cerró: hasta el viernes 9/10, con el lunes 12/10 feriado → martes 13/10 a las 8:00.
const VENTANAS = [["Período", "Desde", "Hasta"], ["2º cuatrimestre", "16/11/2026", "04/12/2026"], ["Anual", "2026-11-16", "2026-12-04"], ["3er bimestre", "28/09/2026", "09/10/2026"]];
const FERIADOS = [["Fecha", "Descripción"], ["12/10/2026", "Día del Respeto a la Diversidad Cultural"], ["07/12/2026", "Feriado ficticio de la prueba"],
  ["08/12/2026", "Inmaculada Concepción de María"], ["25/12/2026", "Navidad"]];
const ASIGNACIONES = [["Orientación", "Nivel", "Materia", "Docente", "Correo"],
  ["EPSI", "EPS2", "Matemática II", "Docente Ficticio de Matemática", MATE],
  ["EPSI", "EPS2", "Inglés I", "Docente Ficticio de Matemática", "Docente.Matematica@bue.edu.ar"],
  ["EPSG", "EPS1", "Matemática I", "Docente Ficticio de Matemática", MATE],
  ["EPSI", "EPS1", "Lengua y Literatura I", "Docente Ficticia de Lengua", LENGUA],
  ["EPSI", "EPS3", "Matemática III", "Docente Ficticia de Lengua", LENGUA]];

function escenario() {
  const grillas = {};
  const crear = (id, periodo, fixtureI, semilla) => {
    grillas[id] = {
      EPSI: new Hoja(grillaSintetica(fixtureI, periodo, semilla, 90000000), { periodo, fuente: id, orientacion: "EPSI" }),
      EPSG: new Hoja(grillaSintetica(EPSG_3B, periodo, semilla + 1, 92000000), { periodo, fuente: id, orientacion: "EPSG" })
    };
  };
  crear(ID[P2C], P2C, EPSI_1C, 7);
  crear(ID[PAN], PAN, EPSI_1C, 7);   // mismos estudiantes y valores iniciales que 2C: la independencia se ve al escribir
  crear(ID[P3B], P3B, EPSI_3B, 11);
  crear(OTRA_2C, P2C, EPSI_1C, 7);
  const s = {
    ahora: new Date(AHORA.getTime()), grillas,
    control: { [P1B]: URL(ID[P1B]), [P1C]: ID[P1C], [P3B]: URL(ID[P3B]), [P2C]: URL(ID[P2C]), [PAN]: URL(ID[PAN]) },
    ventanas: VENTANAS.map(f => f.slice()), feriados: FERIADOS.map(f => f.slice()), asignaciones: ASIGNACIONES.map(f => f.slice()),
    registro: new Hoja([CARGA.REGISTRO.slice()], { esRegistro: true }), lock: nuevoLock(), lecturasDatos: 0,
    flushFallas: [] // una entrada por flush, en orden: 'fallar' o vacío
  };
  s.servicios = {
    reloj: () => new Date(s.ahora.getTime()),
    datos: () => { s.lecturasDatos++; return JSON.parse(JSON.stringify({ ventanas: s.ventanas, feriados: s.feriados, asignaciones: s.asignaciones, fuentes: s.control })); },
    hoja: (fuente, orientacion) => { const g = s.grillas[fuente]; if (!g || !g[orientacion]) throw C.rechazo_("grilla", "No se encontró la hoja."); return g[orientacion]; },
    valoresHoja: (fuente, orientacion) => s.servicios.hoja(fuente, orientacion).getDataRange().getDisplayValues(),
    registro: () => s.registro,
    lock: () => s.lock,
    flush: () => {
      if (s.flushFallas.shift() === "fallar") { EVENTOS.push("flush-fallido"); throw new Error("flush simulado: no se pudieron confirmar los cambios"); }
      EVENTOS.push("flush");
    }
  };
  return s;
}
const hoja = (s, periodo, orientacion) => s.grillas[ID[periodo]][orientacion || "EPSI"];
const todasLasEscrituras = s => Object.values(s.grillas).flatMap(g => Object.values(g)).flatMap(h => h.escrituras);
const todosLosIntentos = s => Object.values(s.grillas).flatMap(g => Object.values(g)).flatMap(h => h.intentos);
const filasRegistro = s => s.registro.v.slice(1).filter(f => f.some(x => x !== "")).map(f => f.map(x => s.registro.ver(x)));
const estados = s => filasRegistro(s).map(f => f[COL.Estado]);

let N = 0;
const op = () => "op-prueba-" + String(++N).padStart(6, "0");
function pedido(periodo, orientacion, materia, dni, campo, nuevo, esperado) {
  return { operacion: op(), periodo, fuente: ID[periodo], orientacion, materia: C.clave_(materia), dni, campo, esperado, nuevo };
}
function guardar(s, correo, p) {
  const previo = OP_ACTUAL;
  OP_ACTUAL = p.operacion;
  try { return C.guardarDocente_(s.servicios, correo, p); } finally { OP_ACTUAL = previo; }
}
function cursantes(s, periodo, orientacion, materia) {
  const v = hoja(s, periodo, orientacion).getDataRange().getDisplayValues();
  const e = C.estructuraGrilla_(v, periodo);
  return C.leerMateria_(v, e, C.materiaPorClave_(e, C.clave_(materia)), periodo);
}
const conSituacion = (s, periodo, orientacion, materia, situacion) => cursantes(s, periodo, orientacion, materia).find(x => x.situacion === situacion && !x.problema);
function ubicar(h, periodo, materia, dni, campo) {
  const v = h.getDataRange().getDisplayValues(); const e = C.estructuraGrilla_(v, periodo); const m = C.materiaPorClave_(e, C.clave_(materia));
  return { fila: C.filaEstudiante_(v, e, dni) + C.campoPorId_(campo || "calificacion").desplazamiento, col: m.columnaValor };
}
const celda = (h, periodo, materia, dni, campo) => { const u = ubicar(h, periodo, materia, dni, campo); return h.ver(h.v[u.fila][u.col]); };
const copia = h => h.v.map(f => f.map(x => h.ver(x)));
function diferencias(antes, h) {
  const d = [];
  h.v.forEach((f, r) => f.forEach((x, c) => { if (h.ver(x) !== ((antes[r] || [])[c] === undefined ? "" : antes[r][c])) d.push([r, c, h.ver(x)]); }));
  return d;
}
const pedCal = (periodo, orientacion, materia, e, nuevo) => pedido(periodo, orientacion, materia, e.dni, "calificacion", nuevo, e.campos.calificacion);

/** Boletín del generador con la grilla principal y las anteriores (sólo hoja EPSI), como construirDatosEstudiante_. */
function boletin(principal, valoresPorPeriodo, dni) {
  return enG(`(function (d) {
    const adv = [];
    const p = { periodo: d.principal, hojaOrigen: 'EPSI', valores: d.v[d.principal], otrasHojas: [] };
    p.estructura = detectarEstructuraFuente_(p.valores, p.periodo);
    p.cola = detectarCamposCola_(p.valores, p.estructura, p.periodo, adv);
    const anteriores = Object.keys(d.v).filter(x => x !== d.principal).map(per => {
      const h = { periodo: per, hojaOrigen: 'EPSI', valores: d.v[per] };
      h.estructura = detectarEstructuraFuente_(h.valores, per);
      return { periodo: per, hojaActual: 'EPSI', hojas: [h] };
    });
    const datos = construirDatosEstudiante_({ principal: p, anteriores }, 'EPSI', d.dni, adv);
    return { materias: Object.fromEntries(datos.materias.map(m => [m.nombre, m.valores])), advertencias: adv };
  })(__d)`, { principal, v: valoresPorPeriodo, dni });
}

// ================================================================ utilidades de prueba

let fallas = 0, total = 0;
function prueba(nombre, fn) {
  total++;
  try { fn(); console.log("OK    " + nombre); } catch (e) { fallas++; console.log("FALLA " + nombre + "\n      " + String(e && e.stack || e).split("\n").slice(0, 5).join("\n      ")); }
}
function igual(a, b, msg) { const x = JSON.stringify(a), y = JSON.stringify(b); if (x !== y) throw new Error((msg || "distinto") + ": " + x + " ≠ " + y); }
function cierto(x, msg) { if (!x) throw new Error(msg || "se esperaba verdadero"); }
function lanza(fn, re) { let e = null; try { fn(); } catch (x) { e = x; } cierto(e && (!re || re.test(e.message)), "se esperaba error " + re + ", hubo: " + (e && e.message)); }
function seccion(t) { console.log("\n# " + t); }
const horaBA = d => new Intl.DateTimeFormat("es-AR", { timeZone: "America/Argentina/Buenos_Aires", weekday: "long", day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).format(d);
const config = (s, ahora) => C.configuracion_(s.servicios.datos(), ahora || s.ahora);
/** Cada liberación del lock va inmediatamente precedida por un flush (intentado), sin escrituras de por medio. */
function confirmadoAntesDeLiberar(eventos, contexto) {
  let tomado = false;
  eventos.forEach((ev, i) => {
    if (ev === "toma") tomado = true;
    if (ev !== "libera") return;
    cierto(tomado, "liberación sin toma");
    cierto(i > 0 && /^flush/.test(eventos[i - 1]), (contexto || "") + ": se liberó el lock sin confirmar: " + eventos.slice(Math.max(0, i - 6), i + 1).join(" → "));
    tomado = false;
  });
  cierto(!tomado, (contexto || "") + ": el lock quedó tomado");
}

// ================================================================ 1. paridad con el generador

seccion("1. La carga y el generador interpretan igual la grilla, sus valores y sus fuentes");
[[EPSI_3B, P3B], [EPSG_3B, P3B], [EPSI_1C, P1C], ["reference/fixtures/3er_bimestre.tsv", P3B], [EPSI_1C, P2C], [EPSI_1C, PAN]].forEach(([f, periodo]) => {
  prueba("mismas materias y columnas: " + path.basename(f) + " como " + periodo, () => {
    const v = new Hoja(leerTsv(f)).getDataRange().getDisplayValues();
    igual(C.estructuraGrilla_(v, periodo).materias.map(m => [m.clave, m.columnaValor]),
      enG("detectarEstructuraFuente_(__d.v, __d.p).materias.map(m => [m.clave, m.columnaValor])", { v, p: periodo }));
  });
});
prueba("las grillas que el generador rechaza, la carga también las rechaza (sin ORIENTACIÓN, materia repetida, sin columna de valor, cola antes de ORIENTACIÓN)", () => {
  const base = leerTsv(EPSI_3B).slice(0, 2);
  const mal = [
    base.map(f => f.map(x => (x === "ORIENTACIÓN" ? "ORIENTACION X" : x))),
    base.map((f, i) => f.map(x => (i === 0 && x === "Inglés I" ? "Matemática II" : x))),
    base.map((f, i) => f.map((x, c) => (i === 0 && c === 16 ? "intruso" : x))),
    base.map((f, i) => f.map(x => (i === 0 && x === "Inglés I" ? "EXPORTADO" : x)))
  ];
  mal.forEach((v, i) => {
    lanza(() => enG("detectarEstructuraFuente_(__d, '3º bimestre')", v), /./);
    let e = null; try { C.estructuraGrilla_(v, P3B); } catch (x) { e = x; }
    cierto(e && e.rechazo && e.rechazo.codigo === "grilla", "caso " + i + ": " + (e && e.message));
  });
});
prueba("lo que el docente escribe y los estados administrativos son válidos para el generador en cada período; CURSA y NO CURSA quedan excluidos", () => {
  CARGA.PERIODOS.forEach(p => {
    CARGA.ESCALA_DOCENTE[p].concat(CARGA.ADMINISTRATIVOS).forEach(v => igual(enG("(r => [r.estado, r.valor])(clasificarValor_(__d.p, __d.v))", { p, v }), ["valido", v], p + " " + v));
    ["CURSA", "NO CURSA", ""].forEach(v => igual(enG("clasificarValor_(__d.p, __d.v).estado", { p, v }), "excluido"));
  });
  igual(enG("clasificarValor_('Anual', 'Suficiente').estado"), "invalido");
  igual(enG("clasificarValor_('3º bimestre', '7').estado"), "invalido");
});
prueba("períodos, celdas de Control, hojas, encabezados de la cola y normalizaciones: iguales a los del generador", () => {
  igual(CARGA.PERIODOS, enG("Object.values(CONFIG_EPS.PERIODOS)"));
  igual(CARGA.CELDAS_FUENTE, enG("celdasFuente_()"));
  igual([CARGA.HOJA_CONTROL, CARGA.HOJAS_ORIGEN], enG("[CONFIG_EPS.HOJA_CONTROL, CONFIG_EPS.HOJAS_ORIGEN]"));
  igual(CARGA.ENCABEZADOS_COLA, enG("Object.values(CONFIG_EPS.ENCABEZADOS_GENERALES).concat(Object.values(CONFIG_EPS.ENCABEZADOS_OPERATIVOS))"));
  ["1º bimestre", "1er bimestre", "3 bimestre", "2do cuatrimestre", "Segundo Cuatrimestre", "2° cuatrimestre", "anual", "Calificación anual", "recuperación"].forEach(x =>
    igual(C.normalizarPeriodo_(x), enG("normalizarPeriodo_(__d)", x), x));
  ["EPS2", "eps 3", "4", "EPS", "x"].forEach(x => igual(C.normalizarNivel_(x), enG("normalizarNivel_(__d)", x), x));
  ["epsi", "EPSG", "EPS I", "otra"].forEach(x => igual(C.normalizarHojaOrigen_(x), enG("normalizarHojaOrigen_(__d)", x), x));
  [URL(ID[P2C]), ID[P2C], "https://docs.google.com/spreadsheets/d/" + ID[PAN] + "/edit?usp=sharing"].forEach(x => igual(C.idPlanilla_(x), enG("extraerSpreadsheetId_(__d)", x)));
  igual(C.idPlanilla_("no es un id"), "");
});

// ================================================================ 2. ventanas, feriados y cierre

seccion("2. Ventanas configurables: cierre a las 8:00 del primer día hábil posterior, con feriados nacionales y de CABA");
prueba("2C y anual comparten ventana: cierran el miércoles 9/12 a las 8:00 (salta sábado, domingo y dos feriados); el 3B ya cerró el martes 13/10 a las 8:00", () => {
  const c = config(escenario());
  igual([c.problemas, c.periodos[P2C].problemas, c.periodos[PAN].problemas], [[], [], []]);
  igual([c.periodos[P2C].estado, c.periodos[PAN].estado, c.periodos[P3B].estado, c.periodos[P1B].estado], ["abierta", "abierta", "cerrada", "sin ventana"]);
  igual(c.periodos[P2C].ventana.cierre.toISOString(), "2026-12-09T11:00:00.000Z");
  igual(c.periodos[PAN].ventana.cierre.toISOString(), c.periodos[P2C].ventana.cierre.toISOString());
  igual(horaBA(c.periodos[P2C].ventana.cierre), "miércoles, 09/12/2026, 08:00:00");
  igual(horaBA(c.periodos[P3B].ventana.cierre), "martes, 13/10/2026, 08:00:00");
  igual(horaBA(c.periodos[P2C].ventana.apertura), "lunes, 16/11/2026, 00:00:00");
});
prueba("límites: abre a las 00:00 de «Desde»; sigue abierta a las 7:59:59.999 del día de cierre y cierra a las 8:00:00.000", () => {
  const s = escenario();
  const estado = iso => config(s, new Date(iso)).periodos[P2C].estado;
  igual([estado("2026-11-15T23:59:59.999-03:00"), estado("2026-11-16T00:00:00-03:00"), estado("2026-12-04T23:59:59-03:00"), estado("2026-12-07T12:00:00-03:00"),
    estado("2026-12-09T07:59:59.999-03:00"), estado("2026-12-09T08:00:00.000-03:00")], ["futura", "abierta", "abierta", "abierta", "abierta", "cerrada"]);
});
prueba("sin feriados ni fin de semana: cierra al día siguiente a las 8:00; un feriado de CABA cargado en la lista también se salta", () => {
  const s = escenario();
  s.ventanas = [VENTANAS[0], ["Anual", "17/11/2026", "18/11/2026"]];
  igual(horaBA(config(s).periodos[PAN].ventana.cierre), "jueves, 19/11/2026, 08:00:00");
  s.feriados.push(["19/11/2026", "Feriado de CABA ficticio para la prueba"]);
  igual(horaBA(config(s).periodos[PAN].ventana.cierre), "viernes, 20/11/2026, 08:00:00");
});
prueba("si faltan los feriados del año en que cae el cierre, el período queda bloqueado en lugar de adivinar; al cargarlos, cierra el lunes 4/1/2027", () => {
  const s = escenario();
  s.ventanas = [VENTANAS[0], ["Anual", "21/12/2026", "31/12/2026"]];
  const c = config(s, new Date("2026-12-28T10:00:00-03:00"));
  cierto(/Faltan los feriados de 2027/.test(c.periodos[PAN].problemas.join()), c.periodos[PAN].problemas.join());
  lanza(() => C.periodoHabilitado_(c, PAN, ID[PAN]), /problema de configuración/);
  s.feriados.push(["01/01/2027", "Año Nuevo"]);
  igual(horaBA(config(s).periodos[PAN].ventana.cierre), "lunes, 04/01/2027, 08:00:00");
});
prueba("errores de configuración: feriado ilegible bloquea toda la carga; período repetido, fechas inválidas o invertidas bloquean ese período; período desconocido no abre nada", () => {
  const s = escenario();
  s.feriados.push(["31/02/2026", "inexistente"]);
  cierto(config(s).problemas.some(p => /Feriados, fila 6/.test(p)));
  const s2 = escenario();
  s2.ventanas.push(["Anual", "20/11/2026", "30/11/2026"], ["1º bimestre", "32/03/2026", "10/04/2026"], ["1º cuatrimestre", "10/07/2026", "01/07/2026"], ["2do cuatri", "16/11/2026", "04/12/2026"]);
  const c = config(s2);
  igual(c.problemas, []);
  cierto(/más de una vez/.test(c.periodos[PAN].problemas.join()), "anual repetido");
  igual(c.periodos[P2C].problemas, []);
  cierto(/fechas no reconocidas/.test(c.periodos[P1B].problemas.join()) && /posterior a/.test(c.periodos[P1C].problemas.join()));
  cierto(c.advertencias.some(a => /período no reconocido \("2do cuatri"\)/.test(a)));
});
prueba("la configuración se lee en cada operación: el Equipo EPS adelanta el cierre con la página abierta y el siguiente guardado se rechaza", () => {
  const s = escenario();
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  s.ventanas[1][2] = "18/11/2026"; // 2C cerró el jueves 19/11 a las 8:00; anual sigue
  const r = guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calidad", "Buena", e.campos.calidad));
  igual([r.estado, r.codigo], ["rechazado", "ventana"]);
  igual(guardar(s, MATE, pedCal(PAN, "EPSI", "Matemática II", e, "8")).estado, "guardado");
  cierto(s.lecturasDatos >= 3, "lecturas de configuración: " + s.lecturasDatos);
});

// ================================================================ 3. período y fuente en cada operación

seccion("3. Cada operación identifica el período y su fuente, y se revalida contra Control del generador");
prueba("la consulta devuelve cada período abierto con su fuente; el 3B cerrado no aparece", () => {
  const r = C.consultarDocente_(escenario().servicios, MATE);
  igual(r.estado, "ok");
  igual(r.periodos.map(p => [p.periodo, p.fuente, p.cierre]), [[P2C, ID[P2C], "2026-12-09T11:00:00.000Z"], [PAN, ID[PAN], "2026-12-09T11:00:00.000Z"]]);
  igual(r.periodos[0].escala, ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]);
});
prueba("el Equipo EPS cambia la grilla de 2C en Control con la página abierta: el guardado con la fuente anterior se rechaza sin escribir en ninguna grilla", () => {
  const s = escenario();
  const vista = C.consultarDocente_(s.servicios, MATE);
  const e = vista.periodos[0].materias.find(m => m.materia === "Matemática II").estudiantes.find(x => x.situacion === "pendiente" && !x.problema);
  s.control[P2C] = URL(OTRA_2C);
  const p = { operacion: op(), periodo: P2C, fuente: vista.periodos[0].fuente, orientacion: "EPSI", materia: C.clave_("Matemática II"), dni: e.dni, campo: "calificacion", esperado: "CURSA", nuevo: "7" };
  const r = guardar(s, MATE, p);
  igual([r.estado, r.codigo], ["rechazado", "fuente"]);
  igual([todasLasEscrituras(s).length, filasRegistro(s).length], [0, 0]);
  const otra = C.consultarDocente_(s.servicios, MATE).periodos[0].fuente;
  igual(otra, OTRA_2C);
  igual(guardar(s, MATE, Object.assign({}, p, { operacion: op(), fuente: otra })).estado, "guardado");
  igual(s.grillas[OTRA_2C].EPSI.escrituras.length, 1);
  igual(filasRegistro(s)[0][COL.Fuente], OTRA_2C);
});
prueba("2C y anual con la misma grilla en Control: ambos bloqueados (nunca se mezclan valoraciones); también si 2C repite la grilla de un período cerrado", () => {
  const s = escenario();
  s.control[PAN] = ID[P2C];
  const c = config(s);
  cierto(/misma que la de Anual/.test(c.periodos[P2C].problemas.join()) && /misma que la de 2º cuatrimestre/.test(c.periodos[PAN].problemas.join()));
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).codigo, "configuracion");
  const vista = C.consultarDocente_(s.servicios, MATE);
  cierto(vista.periodos.every(p => p.problema && !p.materias.length));
  const s2 = escenario();
  s2.control[P2C] = ID[P3B];
  cierto(/misma que la de 3º bimestre/.test(config(s2).periodos[P2C].problemas.join()));
});
prueba("período con ventana sin grilla o con una celda ilegible en Control: bloqueado; un período sin ventana no exige grilla", () => {
  const s = escenario();
  s.control[PAN] = ""; s.control[P1B] = "";
  const c = config(s);
  cierto(/Falta la grilla de Anual en Control!B14/.test(c.periodos[PAN].problemas.join()));
  igual(c.periodos[P1B].problemas, []);
  s.control[PAN] = "planilla del anual";
  cierto(/no tiene una grilla reconocible/.test(config(s).periodos[PAN].problemas.join()));
});
prueba("el pedido debe nombrar el período exacto y una orientación conocida", () => {
  const s = escenario();
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  igual(guardar(s, MATE, Object.assign(pedCal(P2C, "EPSI", "Matemática II", e, "7"), { periodo: "2do cuatrimestre" })).codigo, "periodo");
  igual(guardar(s, MATE, Object.assign(pedCal(P2C, "EPSI", "Matemática II", e, "7"), { orientacion: "EPSX" })).codigo, "pedido");
  igual(guardar(s, MATE, Object.assign(pedCal(P2C, "EPSI", "Matemática II", e, "7"), { operacion: "x" })).codigo, "pedido");
  igual(todasLasEscrituras(s).length, 0);
});

// ================================================================ 4. asignaciones y autorizaciones

seccion("4. Asignaciones vigentes en cada operación");
prueba("cada docente guarda sólo en sus materias; el Registro identifica fecha, período, fuente, orientación, nivel, materia, estudiante, campo, valores y docente", () => {
  const s = escenario();
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  const f = filasRegistro(s)[0];
  igual([f[COL.Estado], f[COL.Intentos], f[COL.Correo], f[COL.Docente], f[COL["Período"]], f[COL.Fuente], f[COL["Orientación"]], f[COL.Nivel], f[COL.Materia], f[COL.DNI], f[COL.Estudiante], f[COL.Campo], f[COL.Anterior], f[COL.Nuevo]],
    ["aplicado", "1", MATE, "Docente Ficticio de Matemática", P2C, ID[P2C], "EPSI", "EPS2", "Matemática II", e.dni, e.nombre, "Calificación", "CURSA", "7"]);
  cierto(/^2026-11-20T18:00:00/.test(f[COL.Fecha]), f[COL.Fecha]);
  const l = conSituacion(s, P2C, "EPSI", "Lengua y Literatura I", "pendiente");
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Lengua y Literatura I", l, "7")).codigo, "asignacion");
  igual(guardar(s, "DOCENTE.LENGUA@bue.edu.ar", pedCal(P2C, "EPSI", "Lengua y Literatura I", l, "8")).estado, "guardado", "el correo no distingue mayúsculas");
});
prueba("reemplazo con la página abierta: el docente anterior pierde el acceso en la operación siguiente y el nuevo corrige lo existente", () => {
  const s = escenario();
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  s.asignaciones[1] = ["EPSI", "EPS2", "Matemática II", "Docente Ficticio de Reemplazo", REEMPLAZO];
  igual(guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calificacion", "8", "7")).codigo, "asignacion");
  igual(guardar(s, REEMPLAZO, pedido(P2C, "EPSI", "Matemática II", e.dni, "calificacion", "8", "7")).estado, "guardado");
  igual(C.consultarDocente_(s.servicios, MATE).periodos[0].materias.map(m => m.materia).includes("Matemática II"), false);
});
prueba("dos docentes para la misma materia, niveles distintos o una fila inválida para esa materia: nadie la carga; las demás siguen", () => {
  const s = escenario();
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  const otra = conSituacion(s, P2C, "EPSG", "Matemática I", "pendiente");
  s.asignaciones.push(["EPSI", "EPS2", "Matemática II", "Otro docente", "otro@bue.edu.ar"]);
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).codigo, "asignacion");
  igual(guardar(s, "otro@bue.edu.ar", pedCal(P2C, "EPSI", "Matemática II", e, "7")).codigo, "asignacion");
  igual(guardar(s, MATE, pedCal(P2C, "EPSG", "Matemática I", otra, "7")).estado, "guardado");
  cierto(config(s).advertencias.some(a => /más de un docente/.test(a)));
  const s2 = escenario();
  s2.asignaciones.push(["EPSI", "EPS2", "Matemática II", "Docente", "correo-mal-escrito"]);
  igual(guardar(s2, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).codigo, "asignacion");
  cierto(config(s2).advertencias.some(a => /nadie puede cargar EPSI · Matemática II/.test(a)));
});
prueba("nivel mal cargado en Asignaciones: rechazado; la consulta lo muestra como problema de esa materia", () => {
  const s = escenario();
  s.asignaciones[1][1] = "EPS3";
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
  igual([r.codigo, /nivel distinto/.test(r.motivo)], ["asignacion", true]);
  cierto(/nivel distinto/.test(C.consultarDocente_(s.servicios, MATE).periodos[0].materias.find(m => m.materia === "Matemática II").problema));
});
prueba("consulta: sólo las materias propias; todos los que cursan, de cualquier grupo; NO CURSA y vacía no aparecen; sin asignaciones, rechazada", () => {
  const s = escenario();
  const r = C.consultarDocente_(s.servicios, LENGUA);
  igual(r.periodos[0].materias.map(m => [m.orientacion, m.nivel, m.materia]), [["EPSI", "EPS1", "Lengua y Literatura I"], ["EPSI", "EPS3", "Matemática III"]]);
  const todas = r.periodos.flatMap(p => p.materias.flatMap(m => m.estudiantes.map(x => [m.nivel, x.grupo, x.campos.calificacion])));
  cierto(todas.length && todas.every(([, , c]) => c && c !== "NO CURSA"));
  const v = hoja(s, P2C).getDataRange().getDisplayValues(); const est = C.estructuraGrilla_(v, P2C);
  cierto(est.materias.some(m => C.leerMateria_(v, est, m, P2C).some(x => x.grupo !== m.nivel)), "hay estudiantes de otro grupo en alguna materia");
  igual(C.consultarDocente_(s.servicios, "nadie@bue.edu.ar").codigo, "asignacion");
});

// ================================================================ 5. reglas de los campos y guardado progresivo

seccion("5. Reglas de cada campo y guardado progresivo e independiente");
prueba("cada campo se guarda solo: Calificación, Calidad, Participación y Observaciones cambian exactamente su celda", () => {
  const s = escenario(); const h = hoja(s, P2C);
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  [["calificacion", "7"], ["calidad", "Excelente"], ["participacion", "Poca"], ["observaciones", "Trabaja bien en grupo"]].forEach(([campo, nuevo]) => {
    const antes = copia(h);
    const actual = celda(h, P2C, "Matemática II", e.dni, campo);
    igual(guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, campo, nuevo, actual)).estado, "guardado", campo);
    const d = diferencias(antes, h); const u = ubicar(h, P2C, "Matemática II", e.dni, campo);
    igual(d, [[u.fila, u.col, nuevo]], campo);
  });
  igual(estados(s), ["aplicado", "aplicado", "aplicado", "aplicado"]);
});
const libres = est => est.materias.filter(m => !ASIGNACIONES.some(a => a[0] === "EPSI" && C.clave_(a[2]) === m.clave));
prueba("ACREDITÓ y Equivalencia: Calificación en solo lectura (ni cambiar ni «Dejar pendiente»); los otros tres campos sí", () => {
  const s = escenario();
  const v = hoja(s, P2C).getDataRange().getDisplayValues(); const est = C.estructuraGrilla_(v, P2C);
  ["ACREDITÓ", "Equivalencia"].forEach(valor => {
    const [m, x] = libres(est).map(m => [m, C.leerMateria_(v, est, m, P2C).find(y => y.campos.calificacion === valor && !y.problema)]).find(par => par[1]);
    igual([x.situacion, x.editable], ["administrativo", { calificacion: false, calidad: true, participacion: true, observaciones: true }]);
    const reg = new Hoja([CARGA.REGISTRO.slice()]); s.registro = reg;
    s.asignaciones.push(["EPSI", m.nivel, m.nombre, "Docente", "admin.prueba@bue.edu.ar"]);
    ["8", "CURSA"].forEach(n => igual(guardar(s, "admin.prueba@bue.edu.ar", pedido(P2C, "EPSI", m.nombre, x.dni, "calificacion", n, valor)).codigo, "calificacion"));
    const otra = CARGA.OPCIONES.calidad.find(o => o !== x.campos.calidad);
    igual(guardar(s, "admin.prueba@bue.edu.ar", pedido(P2C, "EPSI", m.nombre, x.dni, "calidad", otra, x.campos.calidad)).estado, "guardado");
    s.asignaciones.pop();
  });
});
prueba("el docente no puede establecer ACREDITÓ, EQUIVALENCIA, NO CURSA, la escala de otro período ni vacío; un valor no reconocido queda en solo lectura", () => {
  const s = escenario();
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  ["ACREDITÓ", "EQUIVALENCIA", "NO CURSA", "Suficiente", "11", "0", "7,5", ""].forEach(n => igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, n)).codigo, "valor", n));
  const v = hoja(s, P2C).getDataRange().getDisplayValues(); const est = C.estructuraGrilla_(v, P2C);
  const [m, x] = libres(est).map(m => [m, C.leerMateria_(v, est, m, P2C).find(y => !y.problema)]).find(par => par[1]);
  const u = ubicar(hoja(s, P2C), P2C, m.nombre, x.dni); hoja(s, P2C).v[u.fila][u.col] = "Muy bien";
  x.campos.calificacion = "Muy bien";
  igual(cursantes(s, P2C, "EPSI", m.nombre).find(y => y.dni === x.dni).situacion, "a revisar");
  s.asignaciones.push(["EPSI", m.nivel, m.nombre, "Docente", "admin.prueba@bue.edu.ar"]);
  igual(guardar(s, "admin.prueba@bue.edu.ar", pedido(P2C, "EPSI", m.nombre, x.dni, "calificacion", "7", x.campos.calificacion)).codigo, "calificacion");
  igual(todasLasEscrituras(s).length, 0);
});
prueba("3º bimestre: escala conceptual en su ventana; los números se rechazan", () => {
  const s = escenario();
  s.ventanas[3] = ["3º bimestre", "16/11/2026", "04/12/2026"];
  s.asignaciones.push(["EPSI", "EPS2", "Matemática II", "Docente Ficticio de Matemática", MATE]); // misma fila: no es ambigua
  const e = conSituacion(s, P3B, "EPSI", "Matemática II", "pendiente");
  igual(guardar(s, MATE, pedCal(P3B, "EPSI", "Matemática II", e, "7")).codigo, "valor");
  igual(guardar(s, MATE, pedCal(P3B, "EPSI", "Matemática II", e, "suficiente")).estado, "guardado");
  igual(celda(hoja(s, P3B), P3B, "Matemática II", e.dni), "Suficiente");
});
prueba("Calidad y Participación: exactamente sus opciones (otra grafía o texto libre, rechazados); vacío permitido", () => {
  const s = escenario();
  const e = cursantes(s, P2C, "EPSI", "Matemática II").find(x => x.campos.calidad && !x.problema);
  ["buena", "Muy Buena", "Bueno", "texto libre"].forEach(n => igual(guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calidad", n, e.campos.calidad)).codigo, "valor", n));
  igual(guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "participacion", "activa", e.campos.participacion)).codigo, "valor");
  igual(guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calidad", "", e.campos.calidad)).estado, "guardado");
  igual(celda(hoja(s, P2C), P2C, "Matemática II", e.dni, "calidad"), "");
});
prueba("Observaciones: texto libre, siempre como texto (ni fórmula, ni fecha, ni número)", () => {
  const s = escenario(); const h = hoja(s, P2C);
  const e = cursantes(s, P2C, "EPSI", "Matemática II").find(x => !x.problema);
  let esperado = e.campos.observaciones;
  ['=IMPORTXML("http://x";"//a")', "1/2", "  dos   espacios \n y salto  "].forEach(n => {
    const r = guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "observaciones", n, esperado));
    igual(r.estado, "guardado", n);
    esperado = r.actual;
  });
  igual(h.escrituras.map(w => w.crudo), ["'=IMPORTXML(\"http://x\";\"//a\")", "'1/2", "'dos espacios y salto"]);
  igual(celda(h, P2C, "Matemática II", e.dni, "observaciones"), "dos espacios y salto");
});
prueba("«Dejar pendiente» vuelve a CURSA sin tocar los otros campos; el boletín deja de mostrar esa materia", () => {
  const s = escenario(); const h = hoja(s, P2C);
  const e = cursantes(s, P2C, "EPSI", "Matemática II").find(x => x.situacion === "valorada" && !x.problema);
  const antes = copia(h);
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "CURSA")).estado, "guardado");
  const u = ubicar(h, P2C, "Matemática II", e.dni);
  igual(diferencias(antes, h), [[u.fila, u.col, "CURSA"]]);
  igual(boletin(P2C, { [P2C]: copia(h) }, e.dni).materias["Matemática II"], undefined);
});
prueba("DNI repetido en la hoja: la consulta lo muestra sin nada editable y el guardado se rechaza; bloque roto: rechazado sin escribir", () => {
  const s = escenario(); const h = hoja(s, P2C);
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  const otro = cursantes(s, P2C, "EPSI", "Matemática II").find(x => x.dni !== e.dni);
  const u = ubicar(h, P2C, "Matemática II", otro.dni); h.v[u.fila][2] = e.dni;
  const x = cursantes(s, P2C, "EPSI", "Matemática II").find(y => y.dni === e.dni);
  igual([/más de una vez/.test(x.problema), Object.values(x.editable).some(Boolean)], [true, false]);
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).codigo, "grilla");
  const s2 = escenario(); const h2 = hoja(s2, P2C);
  const e2 = conSituacion(s2, P2C, "EPSI", "Matemática II", "pendiente");
  h2.insertarFilas(ubicar(h2, P2C, "Matemática II", e2.dni).fila + 1, 1);
  igual(guardar(s2, MATE, pedido(P2C, "EPSI", "Matemática II", e2.dni, "observaciones", "x", e2.campos.observaciones)).codigo, "grilla");
  igual(todasLasEscrituras(s).length + todasLasEscrituras(s2).length, 0);
});

// ================================================================ 6. 2º cuatrimestre y anual

seccion("6. 2º cuatrimestre y anual: misma ventana, grillas y valoraciones independientes");
prueba("el mismo estudiante y materia en 2C y anual: cada guardado cambia sólo su grilla; el Registro los distingue; el boletín anual muestra ambos", () => {
  const s = escenario(); const h2 = hoja(s, P2C); const ha = hoja(s, PAN);
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  igual(conSituacion(s, PAN, "EPSI", "Matemática II", "pendiente").dni, e.dni, "mismos estudiantes en ambas grillas");
  const a2 = copia(h2), aa = copia(ha);
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  igual([diferencias(a2, h2).length, diferencias(aa, ha).length], [1, 0]);
  igual(guardar(s, MATE, pedido(PAN, "EPSI", "Matemática II", e.dni, "calificacion", "9", "CURSA")).estado, "guardado");
  igual(guardar(s, MATE, pedido(PAN, "EPSI", "Matemática II", e.dni, "observaciones", "Sólo del anual", e.campos.observaciones)).estado, "guardado");
  igual([diferencias(a2, h2).length, diferencias(aa, ha).length], [1, 2]);
  igual([celda(h2, P2C, "Matemática II", e.dni, "observaciones"), celda(ha, PAN, "Matemática II", e.dni, "observaciones")], [e.campos.observaciones, "Sólo del anual"]);
  igual(filasRegistro(s).map(f => [f[COL["Período"]], f[COL.Fuente], f[COL.Nuevo]]), [[P2C, ID[P2C], "7"], [PAN, ID[PAN], "9"], [PAN, ID[PAN], "Sólo del anual"]]);
  const b = boletin(PAN, { [PAN]: copia(ha), [P2C]: copia(h2) }, e.dni);
  igual(b.materias["Matemática II"], { [PAN]: "9", [P2C]: "7" });
  igual(boletin(P2C, { [P2C]: copia(h2) }, e.dni).materias["Matemática II"], { [P2C]: "7" });
});
prueba("un pedido de 2C con la fuente del anual (o al revés) se rechaza: el período y la fuente van juntos", () => {
  const s = escenario();
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  igual(guardar(s, MATE, Object.assign(pedCal(P2C, "EPSI", "Matemática II", e, "7"), { fuente: ID[PAN] })).codigo, "fuente");
  igual(guardar(s, MATE, Object.assign(pedCal(PAN, "EPSI", "Matemática II", e, "7"), { fuente: ID[P2C] })).codigo, "fuente");
  igual(todasLasEscrituras(s).length, 0);
});
prueba("ventana de un solo período: si sólo el anual sigue abierto, la consulta y los guardados de 2C se cierran", () => {
  const s = escenario();
  s.ventanas[1] = ["2º cuatrimestre", "02/11/2026", "13/11/2026"]; // cerró el lunes 16/11 a las 8:00
  igual(C.consultarDocente_(s.servicios, MATE).periodos.map(p => p.periodo), [PAN]);
  const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).codigo, "ventana");
  igual(guardar(s, MATE, pedCal(PAN, "EPSI", "Matemática II", e, "7")).estado, "guardado");
});

// ================================================================ 7. reintentos, idempotencia y resultados inciertos

seccion("7. Reintentos seguros: una fila por operación, un solo intento de escritura, y lo inferido separado de lo confirmado");
const base7 = () => { const s = escenario(); return { s, h: hoja(s, P2C), e: conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente") }; };
prueba("respuesta perdida: el reintento responde «guardado» sin escribir de nuevo ni agregar filas", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  igual(guardar(s, MATE, p).estado, "guardado"); igual(guardar(s, MATE, p).estado, "guardado"); igual(guardar(s, MATE, p).estado, "guardado");
  igual([h.escrituras.length, filasRegistro(s).length], [1, 1]);
});
prueba("la misma operación ejecutada dos veces a la vez: una escribe, la otra lo encuentra en el Registro; una sola fila", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  let segunda;
  h.ganchos.push({ evento: "lectura:completa", fn: () => { segunda = guardar(s, MATE, p); } });
  igual(guardar(s, MATE, p).estado, "guardado"); igual(segunda.estado, "guardado");
  igual([h.escrituras.length, filasRegistro(s).length, estados(s)], [1, 1, ["aplicado"]]);
});
prueba("identificador repetido con otro contenido, o de otro docente: rechazado sin escribir; no se confunde con el cambio original", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  igual(guardar(s, MATE, p).estado, "guardado");
  igual(guardar(s, MATE, Object.assign({}, p, { nuevo: "8", esperado: "7" })).codigo, "operacion");
  igual(guardar(s, MATE, Object.assign({}, p, { campo: "calidad", nuevo: "Buena", esperado: e.campos.calidad })).codigo, "operacion");
  const l = conSituacion(s, P2C, "EPSI", "Lengua y Literatura I", "pendiente");
  igual(guardar(s, LENGUA, Object.assign(pedCal(P2C, "EPSI", "Lengua y Literatura I", l, "7"), { operacion: p.operacion })).codigo, "operacion");
  igual([h.escrituras.length, filasRegistro(s).length], [1, 1]);
});
prueba("REGRESIÓN 4: alta fallida («no aplicado»), el Equipo EPS corrige a 8 y se reenvía el mismo identificador con esperado 8: rechazado; la corrección queda", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  s.registro.fallas.appendRow = "aplicar-y-fallar";
  igual([guardar(s, MATE, p).estado, estados(s)], ["reintentar", ["no aplicado"]]);
  const u = ubicar(h, P2C, "Matemática II", e.dni); h.v[u.fila][u.col] = "8";
  const r = guardar(s, MATE, Object.assign({}, p, { esperado: "8" }));
  igual([r.estado, r.codigo, r.previo], ["rechazado", "operacion", { estado: "no aplicado", nuevo: "7" }]);
  igual([h.intentos.length, celda(h, P2C, "Matemática II", e.dni)], [0, "8"]);
  const original = guardar(s, MATE, p); // el pedido original intacto: conflicto, sin escribir
  igual([original.estado, original.actual, h.intentos.length, estados(s)], ["conflicto", "8", 0, ["no aplicado"]]);
});
prueba("REGRESIÓN 4: el mismo identificador con otro «esperado» se rechaza en todos los estados de la fila; el pedido original intacto sigue funcionando", () => {
  const casos = [
    ["aplicado", s => {}, "guardado"],
    ["incierto", s => { hoja(s, P2C).fallas.setValue = "fallar"; }, "incierto"],
    ["pendiente", s => { s.registro.fallas.setValues = ["fallar", "fallar"]; }, "coincide"],
    ["no aplicado", s => { s.registro.fallas.appendRow = "aplicar-y-fallar"; }, "guardado"]
  ];
  casos.forEach(([estado, preparar, respuestaOriginal]) => {
    const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
    preparar(s); guardar(s, MATE, p);
    igual(estados(s), [estado], estado);
    ["8", "", "7", "Suficiente"].forEach(otro => {
      const r = guardar(s, MATE, Object.assign({}, p, { esperado: otro }));
      igual([r.estado, r.codigo], ["rechazado", "operacion"], estado + " con esperado " + JSON.stringify(otro));
    });
    const intentos = h.intentos.length;
    igual(guardar(s, MATE, Object.assign({}, p, { esperado: " " + p.esperado + " " })).estado, respuestaOriginal, estado + ": el original (con espacios alrededor) es el mismo pedido");
    cierto(h.intentos.length - intentos <= (estado === "no aplicado" ? 1 : 0), estado + ": intentos de más");
  });
});
prueba("setValue falla sin aplicar: «incierto» y fila «incierto»; el reintento no vuelve a intentar; una operación nueva (decisión del docente) guarda", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  h.fallas.setValue = "fallar";
  const r = guardar(s, MATE, p);
  igual([r.estado, r.actual, estados(s)], ["incierto", "CURSA", ["incierto"]]);
  const otra = guardar(s, MATE, p);
  igual([otra.estado, otra.actual, h.intentos.length], ["incierto", "CURSA", 1]);
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  igual([estados(s), h.escrituras.length], [["incierto", "aplicado"], 1]);
});
prueba("setValue informa un error pero la celda muestra el valor nuevo: «inferido» y «coincide», nunca «aplicado» ni «guardado»", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  h.fallas.setValue = "aplicar-y-fallar";
  igual(guardar(s, MATE, p).estado, "coincide");
  const f = filasRegistro(s)[0];
  igual([f[COL.Estado], /error/.test(f[COL.Detalle])], ["inferido", true]);
  igual([guardar(s, MATE, p).estado, h.intentos.length], ["coincide", 1]);
});
prueba("una falla pasajera al marcar el Registro se supera con el segundo intento: «guardado», «aplicado»", () => {
  const { s, h, e } = base7();
  s.registro.fallas.setValues = "fallar";
  igual([guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, estados(s)], ["guardado", ["aplicado"]]);
});
prueba("se escribió la celda pero fallaron las dos marcas del Registro: «incierto»; el reintento la ve con el valor nuevo y la deja «inferido» («coincide»), sin escribir", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  s.registro.fallas.setValues = ["fallar", "fallar"];
  igual(guardar(s, MATE, p).estado, "incierto");
  igual(estados(s), ["pendiente"]);
  igual(guardar(s, MATE, p).estado, "coincide");
  igual(estados(s), ["inferido"]);
  igual([h.escrituras.length, h.intentos.length], [1, 1]);
});
prueba("appendRow ambiguo (la fila se agregó, la grilla no): «reintentar» y la fila queda «no aplicado»; el reintento la reutiliza y escribe una sola vez", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  s.registro.fallas.appendRow = "aplicar-y-fallar";
  igual(guardar(s, MATE, p).estado, "reintentar"); igual(h.escrituras.length, 0);
  igual(filasRegistro(s).map(f => [f[COL.Estado], f[COL.Detalle]]), [["no aplicado", "no se llegó a escribir: falló la anotación en el Registro"]]);
  igual(guardar(s, MATE, p).estado, "guardado");
  igual(filasRegistro(s).map(f => [f[COL.Estado], f[COL.Intentos]]), [["aplicado", "2"]]);
});
prueba("appendRow ambiguo y además falla esa corrección: la fila queda «pendiente»; el reintento, sin saber si se escribió, no escribe («incierto»)", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  s.registro.fallas.appendRow = "aplicar-y-fallar"; s.registro.fallas.setValues = "fallar";
  igual(guardar(s, MATE, p).estado, "reintentar");
  igual(estados(s), ["pendiente"]);
  const r = guardar(s, MATE, p);
  igual([r.estado, r.actual, estados(s), h.intentos.length], ["incierto", "CURSA", ["incierto"], 0]);
});
prueba("se escribió, falló la marca y el Equipo EPS revirtió la celda antes del reintento: el reintento no vuelve a escribir («incierto»)", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  s.registro.fallas.setValues = ["fallar", "fallar"];
  igual(guardar(s, MATE, p).estado, "incierto");
  const u = ubicar(h, P2C, "Matemática II", e.dni); h.v[u.fila][u.col] = "CURSA";
  const r = guardar(s, MATE, p);
  igual([r.estado, r.actual, estados(s), h.intentos.length, celda(h, P2C, "Matemática II", e.dni)], ["incierto", "CURSA", ["incierto"], 1, "CURSA"]);
});
prueba("REGRESIÓN 1: setValue falla, Google aplicó la escritura y el Equipo EPS la revirtió antes de la lectura: «incierto»; ningún reintento pisa la corrección", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  const u = ubicar(h, P2C, "Matemática II", e.dni);
  h.fallas.setValue = "aplicar-y-fallar";
  h.ganchos.push({ evento: "despues:setValue", fn: x => { x.v[u.fila][u.col] = "CURSA"; } }); // corrección del Equipo EPS
  const r = guardar(s, MATE, p);
  igual([r.estado, r.actual, estados(s)], ["incierto", "CURSA", ["incierto"]]);
  cierto(/pudo no aplicarse, o aplicarse y ser revertida/.test(filasRegistro(s)[0][COL.Detalle]));
  for (let i = 0; i < 3; i++) igual(guardar(s, MATE, p).estado, "incierto");
  igual([h.intentos.length, celda(h, P2C, "Matemática II", e.dni)], [1, "CURSA"]);
});
prueba("REGRESIÓN 2: el intento falló sin aplicar y el Equipo EPS puso justo el valor pedido: el reintento responde «coincide» y la fila queda «inferido», no «aplicado»", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  h.fallas.setValue = "fallar";
  igual(guardar(s, MATE, p).estado, "incierto");
  const u = ubicar(h, P2C, "Matemática II", e.dni); h.v[u.fila][u.col] = "7"; // lo pone el Equipo EPS
  const r = guardar(s, MATE, p);
  igual([r.estado, estados(s), h.escrituras.length], ["coincide", ["inferido"], 0]);
  cierto(/no se puede confirmar que lo escribió esta operación/.test(filasRegistro(s)[0][COL.Detalle]));
});
prueba("REGRESIÓN 2: el valor pedido ya estaba en la celda: «coincide», sin fila (nadie se atribuye la modificación); una fila «no aplicado» no se convierte en «inferido»", () => {
  const { s, h, e } = base7(); const u = ubicar(h, P2C, "Matemática II", e.dni);
  h.v[u.fila][u.col] = "7";
  igual([guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, filasRegistro(s).length], ["coincide", 0]);
  h.v[u.fila][u.col] = "CURSA";
  const p = pedCal(P2C, "EPSI", "Matemática II", e, "8");
  s.registro.fallas.appendRow = "aplicar-y-fallar";
  igual(guardar(s, MATE, p).estado, "reintentar"); igual(estados(s), ["no aplicado"]);
  h.v[u.fila][u.col] = "8";
  igual([guardar(s, MATE, p).estado, estados(s), h.intentos.length], ["coincide", ["no aplicado"], 0]);
});
prueba("falla antes de anotar en el Registro (al leer la grilla): «reintentar», porque nada se escribió", () => {
  const { s, e } = base7();
  const hojaReal = s.servicios.hoja;
  s.servicios.hoja = () => { throw new Error("Servicio no disponible (simulado)"); };
  const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
  s.servicios.hoja = hojaReal;
  igual([r.estado, filasRegistro(s).length], ["reintentar", 0]);
});
prueba("reintento de una operación ya aplicada después de una corrección del Equipo EPS: conflicto; no pisa la corrección", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  igual(guardar(s, MATE, p).estado, "guardado");
  const u = ubicar(h, P2C, "Matemática II", e.dni); h.v[u.fila][u.col] = "CURSA";
  const r = guardar(s, MATE, p);
  igual([r.estado, r.actual], ["conflicto", "CURSA"]);
  igual(h.escrituras.length, 1);
});
prueba("la celda muestra otra cosa después de escribir (formato 7 → \"7,00\"): «revisar» y conflicto; el reintento no vuelve a escribir e informa el estado previo", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  const u = ubicar(h, P2C, "Matemática II", e.dni); h.formatos[u.fila + "," + u.col] = n => n.toFixed(2).replace(".", ",");
  const r = guardar(s, MATE, p);
  igual([r.estado, r.actual, estados(s)], ["conflicto", "7,00", ["revisar"]]);
  const otra = guardar(s, MATE, p); // "7,00" no es un valor reconocido: Calificación queda en solo lectura
  igual([otra.estado, otra.codigo, otra.previo], ["rechazado", "calificacion", { estado: "revisar", nuevo: "7" }]);
  igual([h.escrituras.length, filasRegistro(s).length], [1, 1]);
});
prueba("reintento después del cierre: rechazado sin escribir, informando el estado previo de la operación", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  igual(guardar(s, MATE, p).estado, "guardado");
  s.ahora = new Date("2026-12-09T08:00:00-03:00");
  const r = guardar(s, MATE, p);
  igual([r.estado, r.codigo, r.previo], ["rechazado", "ventana", { estado: "aplicado", nuevo: "7" }]);
  const nuevo = guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calidad", "Buena", e.campos.calidad));
  igual([nuevo.codigo, nuevo.previo], ["ventana", undefined]);
  igual(h.escrituras.length, 1);
});
prueba("lock ocupado: «reintentar» sin anotar ni escribir", () => {
  const { s, h, e } = base7();
  s.lock.tomado = true;
  const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
  s.lock.tomado = false;
  igual([r.estado, h.escrituras.length, filasRegistro(s).length], ["reintentar", 0, 0]);
});
prueba("control de pendientes: resuelve, mirando la celda y sin escribir, las filas «pendiente» viejas (valor nuevo: «inferido»; otro: «incierto») y deja las recientes", () => {
  const s = escenario(); const h = hoja(s, P2C);
  const ps = ["Matemática II", "Inglés I"].flatMap(m => cursantes(s, P2C, "EPSI", m).filter(x => x.situacion === "pendiente" && !x.problema).map(x => [m, x])).slice(0, 4);
  cierto(ps.length === 4, "faltan pendientes en la grilla ficticia");
  ps.forEach(([m, e]) => { s.registro.fallas.setValues = ["fallar", "fallar"]; igual(guardar(s, MATE, pedCal(P2C, "EPSI", m, e, "7")).estado, "incierto"); });
  igual(estados(s), ["pendiente", "pendiente", "pendiente", "pendiente"]);
  const u1 = ubicar(h, P2C, ps[1][0], ps[1][1].dni); h.v[u1.fila][u1.col] = "CURSA";
  const u2 = ubicar(h, P2C, ps[2][0], ps[2][1].dni); h.v[u2.fila][u2.col] = "ACREDITÓ";
  s.ahora = new Date(s.ahora.getTime() + 30 * 60000);
  s.registro.v[4][COL.Fecha] = new Date(s.ahora.getTime() - 60000); // la cuarta es reciente
  const intentos = todosLosIntentos(s).length;
  const r = C.resolverPendientes_(s.servicios, 10 * 60000);
  igual([r.revisadas, r.inferido, r.incierto, r.errores], [3, 1, 2, []]);
  igual(estados(s), ["inferido", "incierto", "incierto", "pendiente"]);
  igual(todosLosIntentos(s).length, intentos);
});

// ================================================================ 8. concurrencia y cambios del Equipo EPS

seccion("7b. Confirmación (flush) antes de liberar el lock, en todos los caminos que escriben bajo lock");
prueba("cada camino que escribe el Registro o la grilla bajo lock confirma con flush inmediatamente antes de liberarlo", () => {
  const caminos = {
    "guardado normal": s => { const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente"); igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado"); },
    "reintento: pendiente → inferido": s => { const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente"); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
      s.registro.fallas.setValues = ["fallar", "fallar"]; guardar(s, MATE, p); EVENTOS = []; igual(guardar(s, MATE, p).estado, "coincide"); },
    "reintento: pendiente → incierto": s => { const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente"); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
      s.registro.fallas.setValues = ["fallar", "fallar"]; guardar(s, MATE, p); const h = hoja(s, P2C); const u = ubicar(h, P2C, "Matemática II", e.dni); h.v[u.fila][u.col] = "CURSA";
      EVENTOS = []; igual(guardar(s, MATE, p).estado, "incierto"); },
    "anulación: la grilla cambió antes de escribir": s => { const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente");
      s.registro.ganchos.push({ evento: "appendRow", fn: () => hoja(s, P2C).insertarFilas(3, 4) }); igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "reintentar"); },
    "anulación: la ventana cerró antes de escribir (rechazo)": s => { const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente"); s.ahora = new Date("2026-12-09T07:59:59.900-03:00");
      s.registro.ganchos.push({ evento: "appendRow", fn: () => { s.ahora = new Date("2026-12-09T08:00:00-03:00"); } }); igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).codigo, "ventana"); },
    "segunda marca después de una falla": s => { const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente"); s.registro.fallas.setValues = "fallar"; igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado"); },
    "error ambiguo de setValue («incierto»)": s => { const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente"); hoja(s, P2C).fallas.setValue = "fallar"; igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "incierto"); },
    "resolución de pendientes": s => { const e = conSituacion(s, P2C, "EPSI", "Matemática II", "pendiente"); s.registro.fallas.setValues = ["fallar", "fallar"]; guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
      EVENTOS = []; s.ahora = new Date(s.ahora.getTime() + 3600000); igual(C.resolverPendientes_(s.servicios, 0).inferido, 1); }
  };
  Object.keys(caminos).forEach(nombre => {
    const s = escenario(); EVENTOS = [];
    caminos[nombre](s);
    cierto(EVENTOS.includes("escritura") && EVENTOS.includes("libera"), nombre + ": el camino no escribió bajo lock");
    confirmadoAntesDeLiberar(EVENTOS, nombre);
  });
});
prueba("si falla el flush final después de una escritura confirmada: «incierto», nunca «guardado»; el lock se libera; el reintento no vuelve a escribir", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  s.flushFallas = ["", "", "fallar"]; // tras el alta, tras escribir la grilla, y el final antes de liberar
  EVENTOS = [];
  const r = guardar(s, MATE, p);
  igual([r.estado, r.actual, s.lock.tomado, h.intentos.length], ["incierto", "7", false, 1]);
  confirmadoAntesDeLiberar(EVENTOS, "flush final fallido");
  // Si la marca se confirmó igual, el reintento la encuentra «aplicado»; si se perdió, sólo puede inferir.
  igual(guardar(s, MATE, p).estado, "guardado");
  s.registro.v[1][COL.Estado] = "pendiente";
  igual([guardar(s, MATE, p).estado, estados(s), h.intentos.length], ["coincide", ["inferido"], 1]);
});
prueba("si falla el flush final en caminos que no confirman nada (conflicto, coincide, rechazo), la respuesta no cambia y el lock se libera", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  const u = ubicar(h, P2C, "Matemática II", e.dni); h.v[u.fila][u.col] = "6";
  s.flushFallas = ["fallar"];
  igual([guardar(s, MATE, p).estado, s.lock.tomado], ["conflicto", false]);
  h.v[u.fila][u.col] = "7"; s.flushFallas = ["fallar"];
  igual([guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, s.lock.tomado], ["coincide", false]);
  s.flushFallas = ["fallar"]; s.lock.alEsperar = () => { s.asignaciones[1][4] = REEMPLAZO; };
  igual([guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "8")).codigo, s.lock.tomado, h.intentos.length], ["asignacion", false, 0]);
});
prueba("si falla el flush posterior al alta en el Registro: no se intenta escribir la grilla («reintentar»); si falla el posterior a la escritura: «inferido», no «aplicado»", () => {
  const a = base7(); a.s.flushFallas = ["fallar"];
  igual([guardar(a.s, MATE, pedCal(P2C, "EPSI", "Matemática II", a.e, "7")).estado, a.h.intentos.length, estados(a.s)], ["reintentar", 0, ["no aplicado"]]);
  const b = base7(); b.s.flushFallas = ["", "fallar"];
  igual([guardar(b.s, MATE, pedCal(P2C, "EPSI", "Matemática II", b.e, "7")).estado, estados(b.s)], ["coincide", ["inferido"]]);
});
prueba("si falla el flush al resolver pendientes: se informa en errores y el lock se libera", () => {
  const { s, e } = base7();
  s.registro.fallas.setValues = ["fallar", "fallar"]; guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
  s.ahora = new Date(s.ahora.getTime() + 3600000); s.flushFallas = ["fallar"];
  const r = C.resolverPendientes_(s.servicios, 0);
  igual([r.inferido, s.lock.tomado], [1, false]);
  cierto(r.errores.some(x => /No se pudo confirmar en el Registro/.test(x)), JSON.stringify(r.errores));
});

seccion("8. Escrituras concurrentes y cambios administrativos directos en la grilla, el Registro y la configuración");
prueba("dos dispositivos del mismo docente: el segundo cambio, sobre un valor viejo, es conflicto", () => {
  const { s, e } = base7();
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "8"));
  igual([r.estado, r.actual], ["conflicto", "7"]);
});
prueba("el Equipo EPS cambia la celda entre la lectura y el lock: conflicto con su valor, sin fila ni escritura", () => {
  const { s, h, e } = base7(); const u = ubicar(h, P2C, "Matemática II", e.dni);
  h.ganchos.push({ evento: "lectura:parcial", fn: x => { x.v[u.fila][u.col] = "6"; } });
  const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
  igual([r.estado, r.actual, h.escrituras.length, filasRegistro(s).length], ["conflicto", "6", 0, 0]);
});
prueba("el Equipo EPS marca NO CURSA entre la lectura y el lock mientras se guarda Calidad: rechazado sin escribir", () => {
  const { s, h, e } = base7(); const u = ubicar(h, P2C, "Matemática II", e.dni);
  h.ganchos.push({ evento: "lectura:parcial", fn: x => { x.v[u.fila][u.col] = "NO CURSA"; } });
  const r = guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calidad", "Buena", e.campos.calidad));
  igual([r.codigo, h.escrituras.length, filasRegistro(s).length], ["cursada", 0, 0]);
});
prueba("el Equipo EPS escribe la misma celda justo después del script: «revisar»; el docente ve el valor del Equipo", () => {
  const { s, h, e } = base7(); const u = ubicar(h, P2C, "Matemática II", e.dni);
  h.ganchos.push({ evento: "despues:setValue", fn: x => { x.v[u.fila][u.col] = "ACREDITÓ"; } });
  const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
  igual([r.estado, r.actual, estados(s)], ["conflicto", "ACREDITÓ", ["revisar"]]);
});
prueba("estudiantes (4 filas) y materias (2 columnas) agregados antes del guardado: se escribe donde corresponde", () => {
  const { s, h, e } = base7();
  h.insertarFilas(3, 4); h.insertarColumnas(4, 2); h.v[0][4] = "Materia nueva";
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  igual(celda(h, P2C, "Matemática II", e.dni), "7");
});
prueba("filas o columnas insertadas entre la lectura y la comprobación previa: «reintentar» sin fila; el reintento guarda bien", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  h.ganchos.push({ evento: "lectura:parcial", fn: x => x.insertarFilas(3, 4) });
  igual(guardar(s, MATE, p).estado, "reintentar");
  h.ganchos.push({ evento: "lectura:parcial", fn: x => x.insertarColumnas(4, 2) });
  igual(guardar(s, MATE, p).estado, "reintentar");
  igual([h.escrituras.length, filasRegistro(s).length], [0, 0]);
  igual(guardar(s, MATE, p).estado, "guardado"); igual(celda(h, P2C, "Matemática II", e.dni), "7");
});
prueba("fila vacía insertada dentro del bloque del estudiante antes de la comprobación previa: se detecta («reintentar»), nunca escribe en la fila nueva", () => {
  const { s, h, e } = base7(); const u = ubicar(h, P2C, "Matemática II", e.dni);
  h.ganchos.push({ evento: "lectura:parcial", fn: x => x.insertarFilas(u.fila + 1, 1) });
  igual(guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calidad", "Excelente", e.campos.calidad)).estado, "reintentar");
  igual(h.escrituras.length, 0);
});
prueba("filas insertadas después de escribir: la celda del estudiante muestra el valor, pero ya no se puede confirmar en su lugar → «inferido», «coincide»", () => {
  const { s, h, e } = base7();
  h.ganchos.push({ evento: "despues:setValue", fn: x => x.insertarFilas(3, 4) });
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "coincide");
  igual([estados(s), celda(h, P2C, "Matemática II", e.dni)], [["inferido"], "7"]);
  cierto(/verificar la celda [A-Z]+\d+ de la hoja EPSI/.test(filasRegistro(s)[0][COL.Detalle]));
});
prueba("REGRESIÓN 3: filas insertadas entre la primera comprobación y la escritura (durante el alta en el Registro): la última comprobación lo detecta; no se escribe y el reintento reutiliza la fila", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  s.registro.ganchos.push({ evento: "appendRow", fn: () => h.insertarFilas(3, 4) });
  igual(guardar(s, MATE, p).estado, "reintentar");
  igual([h.intentos.length, estados(s)], [0, ["no aplicado"]]);
  cierto(/entre la comprobación y la escritura/.test(filasRegistro(s)[0][COL.Detalle]));
  igual(guardar(s, MATE, p).estado, "guardado");
  igual([filasRegistro(s).map(f => [f[COL.Estado], f[COL.Intentos]]), celda(h, P2C, "Matemática II", e.dni)], [[["aplicado", "2"]], "7"]);
});
prueba("REGRESIÓN 3: el Equipo EPS cambia otra fila del bloque (NO CURSA) entre la primera comprobación y la escritura: no se escribe", () => {
  const { s, h, e } = base7(); const u = ubicar(h, P2C, "Matemática II", e.dni);
  s.registro.ganchos.push({ evento: "appendRow", fn: () => { h.v[u.fila][u.col] = "NO CURSA"; } });
  const r = guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calidad", "Buena", e.campos.calidad));
  igual([r.estado, h.intentos.length, estados(s)], ["reintentar", 0, ["no aplicado"]]);
  igual(guardar(s, MATE, pedido(P2C, "EPSI", "Matemática II", e.dni, "calidad", "Buena", e.campos.calidad)).codigo, "cursada");
});
prueba("filas insertadas JUSTO antes de escribir: «revisar» con la celda a verificar; la operación no reescribe y una nueva guarda bien", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  h.ganchos.push({ evento: "antes:setValue", fn: x => x.insertarFilas(3, 4) });
  const r = guardar(s, MATE, p);
  igual([r.estado, r.actual], ["conflicto", "CURSA"]);
  const f = filasRegistro(s)[0];
  cierto(f[COL.Estado] === "revisar" && /verificar la celda [A-Z]+\d+ de la hoja EPSI/.test(f[COL.Detalle]), f[COL.Detalle]);
  igual(guardar(s, MATE, p).estado, "conflicto"); igual(h.escrituras.length, 1);
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  igual([celda(h, P2C, "Matemática II", e.dni), estados(s)], ["7", ["revisar", "aplicado"]]);
});
prueba("el Equipo EPS inserta filas en el Registro entre el alta y la marca: se marca la fila de la operación, no otra", () => {
  const { s, h, e } = base7();
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).estado, "guardado");
  const l = conSituacion(s, P2C, "EPSI", "Matemática III", "pendiente");
  h.ganchos.push({ evento: "despues:setValue", fn: () => s.registro.insertarFilas(1, 2) });
  igual(guardar(s, LENGUA, pedCal(P2C, "EPSI", "Matemática III", l, "8")).estado, "guardado");
  igual(filasRegistro(s).map(f => [f[COL.Estado], f[COL.Nuevo]]), [["aplicado", "7"], ["aplicado", "8"]]);
});
prueba("Registro con columnas cambiadas: rechazado sin escribir (nunca se escribe en columnas deducidas)", () => {
  const { s, h, e } = base7();
  s.registro.v[0][1] = "Detalle"; s.registro.v[0][2] = "Estado";
  igual(guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7")).codigo, "registro");
  igual(h.escrituras.length, 0);
});
prueba("REGRESIÓN 3: mientras el guardado espera el lock, el Equipo EPS reasigna la materia, cambia la grilla en Control, adelanta el cierre o rompe la configuración: rechazado sin intentar escribir", () => {
  const casos = [
    ["asignacion", s => { s.asignaciones[1] = ["EPSI", "EPS2", "Matemática II", "Reemplazo", REEMPLAZO]; }],
    ["fuente", s => { s.control[P2C] = URL(OTRA_2C); }],
    ["ventana", s => { s.ventanas[1][2] = "18/11/2026"; }],
    ["configuracion", s => { s.feriados.push(["no es una fecha", ""]); }],
    ["asignacion", s => { s.asignaciones[1][1] = "EPS3"; }]
  ];
  casos.forEach(([codigo, cambio]) => {
    const { s, h, e } = base7();
    s.lock.alEsperar = () => cambio(s);
    const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
    igual([r.estado, r.codigo, todosLosIntentos(s).length, filasRegistro(s).length], ["rechazado", codigo, 0, 0], codigo);
  });
});
prueba("REGRESIÓN 3: un reintento cuya asignación se revocó mientras esperaba el lock se rechaza e informa el estado previo, sin escribir", () => {
  const { s, h, e } = base7(); const p = pedCal(P2C, "EPSI", "Matemática II", e, "7");
  s.registro.fallas.setValues = ["fallar", "fallar"];
  igual(guardar(s, MATE, p).estado, "incierto");
  s.lock.alEsperar = () => { s.asignaciones[1] = ["EPSI", "EPS2", "Matemática II", "Reemplazo", REEMPLAZO]; };
  const r = guardar(s, MATE, p);
  igual([r.codigo, r.previo, h.intentos.length], ["asignacion", { estado: "pendiente", nuevo: "7" }, 1]);
});
prueba("la ventana cierra (8:00) durante la operación, entre la autorización y la escritura: rechazado sin anotar ni escribir", () => {
  const { s, h, e } = base7();
  s.ahora = new Date("2026-12-09T07:59:59.900-03:00");
  h.ganchos.push({ evento: "lectura:parcial", fn: () => { s.ahora = new Date("2026-12-09T08:00:00.000-03:00"); } });
  const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
  igual([r.codigo, h.escrituras.length, filasRegistro(s).length], ["ventana", 0, 0]);
});
prueba("la ventana cierra durante el alta en el Registro, justo antes de escribir: rechazado, la fila queda «no aplicado» y no hay intento de escritura", () => {
  const { s, h, e } = base7();
  s.ahora = new Date("2026-12-09T07:59:59.900-03:00");
  s.registro.ganchos.push({ evento: "appendRow", fn: () => { s.ahora = new Date("2026-12-09T08:00:00.000-03:00"); } });
  const r = guardar(s, MATE, pedCal(P2C, "EPSI", "Matemática II", e, "7"));
  igual([r.codigo, h.intentos.length, estados(s)], ["ventana", 0, ["no aplicado"]]);
});

// ================================================================ 9. simulación aleatoria

seccion("9. Simulación aleatoria: fallas ambiguas, respuestas perdidas, reintentos, duplicados y ediciones del Equipo EPS (" + CORRIDAS + " corridas)");
function simular(semilla, pasos) {
  EVENTOS = [];
  const azar = prng(semilla);
  const elegir = l => l[Math.floor(azar() * l.length)];
  const s = escenario();
  const objetivos = [[MATE, "EPSI", "Matemática II"], [MATE, "EPSI", "Inglés I"], [MATE, "EPSG", "Matemática I"], [LENGUA, "EPSI", "Lengua y Literatura I"], [LENGUA, "EPSI", "Matemática III"]];
  const textos = ["Muy bien", "1/2", "=SUMA(1;2)", "", "Entrega tarde", "+5"];
  const viejos = {};
  const llamadas = [];
  const cola = [];
  const est = { pedidos: 0, llamadas: 0, reintentos: 0, perdidas: 0, fallas: 0, equipo: 0, estructurales: 0, duplicadas: 0, configuracion: 0, alterados: 0, flush: 0 };
  for (let paso = 0; paso < pasos; paso++) {
    let item;
    if (cola.length && azar() < 0.45) { item = cola.shift(); est.reintentos++; } else {
      const [correo, orientacion, materia] = elegir(objetivos);
      const periodo = azar() < 0.5 ? P2C : PAN;
      const lista = cursantes(s, periodo, orientacion, materia).filter(x => !x.problema);
      if (!lista.length) continue;
      const x = elegir(lista);
      const campos = CARGA.CAMPOS.map(c => c.id).filter(c => x.editable[c]);
      const campo = elegir(campos);
      const opciones = campo === "calificacion" ? CARGA.ESCALA_DOCENTE[periodo].concat("CURSA") : campo === "observaciones" ? textos : CARGA.OPCIONES[campo].concat("");
      const clave = [periodo, orientacion, materia, x.dni, campo].join("|");
      const esperado = azar() < 0.15 && viejos[clave] !== undefined ? viejos[clave] : x.campos[campo];
      viejos[clave] = x.campos[campo];
      item = { correo, p: pedido(periodo, orientacion, materia, x.dni, campo, elegir(opciones), esperado), intentos: 0 };
      est.pedidos++;
    }
    // A veces el cliente reenvía el mismo identificador con el pedido alterado (otro «esperado» u otro valor).
    let p = item.p, mutado = false;
    if (item.intentos > 0 && azar() < 0.12) {
      p = Object.assign({}, item.p, azar() < 0.7 ? { esperado: elegir(["8", "", "CURSA", "Buena", "otra cosa"]) } : { nuevo: elegir(["9", "Activa", "otra observación"]) });
      mutado = JSON.stringify(p) !== JSON.stringify(item.p);
      if (mutado) est.alterados++;
    }
    const teniaFila = filasRegistro(s).some(f => f[COL["Operación"]] === p.operacion);
    const h = hoja(s, p.periodo, p.orientacion);
    const materiaNombre = objetivos.find(o => C.clave_(o[2]) === p.materia)[2];
    const r = azar();
    let restaurar = null;
    item.inyeccion = r < 0.19 ? "falla" : r < 0.24 ? "equipo antes" : r < 0.28 ? "equipo después" : r < 0.31 ? "estructural" : r < 0.35 ? "duplicada" :
      r < 0.38 ? "error y reversión" : r < 0.41 ? "configuración durante el lock" : r < 0.44 ? "estructural entre comprobaciones" : r < 0.47 ? "flush fallido" : "";
    const editarCelda = () => { try { const u = ubicar(h, p.periodo, materiaNombre, p.dni, p.campo); h.v[u.fila][u.col] = p.campo === "calificacion" ? elegir(["CURSA", "ACREDITÓ", "NO CURSA", "8", "5"]) : elegir(["Buena", "Activa", "Otra cosa", ""]); est.equipo++; } catch (e) { /* ya no está */ } };
    if (r < 0.05) { h.fallas.setValue = "fallar"; est.fallas++; }
    else if (r < 0.10) { h.fallas.setValue = "aplicar-y-fallar"; est.fallas++; }
    else if (r < 0.13) { s.registro.fallas.appendRow = "aplicar-y-fallar"; est.fallas++; }
    else if (r < 0.15) { s.registro.fallas.appendRow = "fallar"; est.fallas++; }
    else if (r < 0.19) { s.registro.fallas.setValues = azar() < 0.5 ? "fallar" : ["fallar", "fallar"]; est.fallas++; }
    else if (r < 0.24) h.ganchos.push({ evento: "lectura:parcial", fn: editarCelda });
    else if (r < 0.28) h.ganchos.push({ evento: "despues:setValue", fn: editarCelda });
    else if (r < 0.31) {
      const dni = p.dni; est.estructurales++;
      h.ganchos.push({ evento: elegir(["lectura:parcial", "antes:setValue", "despues:setValue"]), fn: x => { const f = C.filasConDni_(x.v.map(z => z.map(String)), dni)[0]; if (f !== undefined) x.insertarFilas(f, 4); } });
    }
    else if (r < 0.35) { est.duplicadas++; h.ganchos.push({ evento: "lectura:completa", fn: () => { const rr = guardar(s, item.correo, p); llamadas.push({ p, r: rr, efectivas: [], anidada: true }); } }); }
    else if (r < 0.38) { est.fallas++; est.equipo++; h.fallas.setValue = "aplicar-y-fallar"; h.ganchos.push({ evento: "despues:setValue", fn: () => { try { const u = ubicar(h, p.periodo, materiaNombre, p.dni, p.campo); h.v[u.fila][u.col] = p.esperado; } catch (e) { /* */ } } }); }
    else if (r < 0.41) { est.configuracion++; const guardadas = [JSON.stringify(s.asignaciones), JSON.stringify(s.control)]; restaurar = () => { s.asignaciones = JSON.parse(guardadas[0]); s.control = JSON.parse(guardadas[1]); };
      s.lock.alEsperar = () => { if (azar() < 0.5) s.asignaciones = s.asignaciones.map(a => (a[1] === "EPS2" ? [a[0], a[1], a[2], "Reemplazo", REEMPLAZO] : a)); else s.control[p.periodo] = URL(OTRA_2C); }; }
    else if (r >= 0.44 && r < 0.47) { est.flush++; s.flushFallas = Array(Math.floor(azar() * 3)).fill("").concat("fallar"); }
    else if (r < 0.44) { est.estructurales++; s.registro.ganchos.push({ evento: "appendRow", fn: () => { const f = C.filasConDni_(h.v.map(z => z.map(String)), p.dni)[0]; if (f !== undefined) h.insertarFilas(f, 4); } }); }
    const antes = SEQ;
    const eventosAntes = EVENTOS.length;
    const resp = guardar(s, item.correo, p);
    est.llamadas++;
    h.ganchos = []; h.fallas = {}; s.registro.fallas = {}; s.registro.ganchos = []; s.lock.alEsperar = null;
    if (restaurar) restaurar();
    const efectivas = todasLasEscrituras(s).filter(w => w.seq > antes && w.op === p.operacion && w.antes !== w.valor);
    let despues = null;
    try { despues = celda(h, p.periodo, materiaNombre, p.dni, p.campo); } catch (e) { /* bloque ya no ubicable */ }
    const intentosLlamada = todosLosIntentos(s).filter(w => w.seq > antes && w.op === p.operacion).length;
    const flushFallido = EVENTOS.slice(eventosAntes).includes("flush-fallido");
    s.flushFallas = [];
    llamadas.push({ p, r: resp, efectivas, despues, mutado, teniaFila, intentosLlamada, flushFallido, inyeccion: item.inyeccion + (r < 0.19 ? ":" + r.toFixed(3) : "") });
    if (mutado) { cola.push(item); continue; } // el pedido original sigue su curso
    const perdida = azar() < 0.1;
    if (perdida) est.perdidas++;
    if ((perdida || resp.estado === "reintentar" || resp.estado === "incierto") && ++item.intentos < 5) cola.push(item);
  }
  const resolucion = C.resolverPendientes_(s.servicios, 0);
  confirmadoAntesDeLiberar(EVENTOS, "simulación " + semilla);
  return { s, llamadas, est, resolucion };
}
function invariantes({ s, llamadas, resolucion }) {
  const filas = filasRegistro(s);
  const porOp = {};
  filas.forEach((f, i) => { const o = f[COL["Operación"]]; cierto(!porOp[o], "operación con dos filas en el Registro: " + o); porOp[o] = { f, i }; });
  igual(resolucion.errores, [], "errores al resolver pendientes");
  cierto(!filas.some(f => f[COL.Estado] === "pendiente"), "quedaron filas pendientes");
  const hojas = Object.values(s.grillas).flatMap(g => Object.values(g));
  const intentosPorOp = {};
  hojas.forEach(h => h.intentos.forEach(w => { (intentosPorOp[w.op] = intentosPorOp[w.op] || []).push(w); }));
  const efectivasPorOp = {};
  hojas.forEach(h => h.escrituras.filter(w => w.antes !== w.valor && w.op).forEach(w => { (efectivasPorOp[w.op] = efectivasPorOp[w.op] || []).push(Object.assign({ hoja: h }, w)); }));
  const detalle = o => JSON.stringify({ registro: porOp[o] && porOp[o].f, intentos: (intentosPorOp[o] || []).length,
    llamadas: llamadas.filter(l => l.p.operacion === o).map(l => [l.r.estado, l.r.actual, l.efectivas.length, l.anidada || false, l.inyeccion]) });
  // Un solo intento de escritura por operación, con o sin error: un reintento nunca vuelve a escribir.
  Object.keys(intentosPorOp).forEach(o => {
    cierto(intentosPorOp[o].length === 1, "la operación " + o + " intentó escribir " + intentosPorOp[o].length + " veces: " + detalle(o));
    const reg = porOp[o];
    cierto(reg && ["aplicado", "inferido", "incierto", "revisar"].includes(reg.f[COL.Estado]), "intento de escritura sin su fila resuelta: " + detalle(o));
    cierto(s.registro.altas.some(a => a.op === o && a.seq < intentosPorOp[o][0].seq), "el intento de " + o + " no fue precedido por su fila del Registro");
  });
  Object.keys(efectivasPorOp).forEach(o => {
    const w = efectivasPorOp[o][0], reg = porOp[o];
    cierto(w.hoja.periodo === reg.f[COL["Período"]] && w.hoja.fuente === reg.f[COL.Fuente], "escritura fuera de la grilla de su período: " + o);
  });
  // 'aplicado' sólo con una escritura propia, sin error, del valor nuevo sobre el anterior registrado.
  let inferidas = 0;
  filas.forEach(f => {
    const o = f[COL["Operación"]], estado = f[COL.Estado];
    if (estado === "aplicado") {
      const ws = efectivasPorOp[o] || [];
      cierto(ws.length === 1, "fila «aplicado» sin escritura propia: " + detalle(o));
      igual([ws[0].antes, ws[0].valor], [f[COL.Anterior], f[COL.Nuevo]], "anterior y nuevo de " + o);
    }
    if (estado === "inferido" && !efectivasPorOp[o]) inferidas++;
    if (estado === "no aplicado") cierto(!intentosPorOp[o], "fila «no aplicado» con intento de escritura: " + detalle(o));
  });
  llamadas.forEach(l => {
    cierto(["guardado", "coincide", "conflicto", "reintentar", "incierto", "rechazado"].includes(l.r.estado), "respuesta desconocida " + l.r.estado);
    if (["reintentar", "rechazado"].includes(l.r.estado)) igual(l.efectivas.length, 0, "«" + l.r.estado + "» con escritura");
    if (l.mutado && l.teniaFila) {
      igual([l.r.estado, l.intentosLlamada], ["rechazado", 0], "reenvío alterado de una operación registrada (" + l.p.operacion + ")");
    }
    if (l.flushFallido) cierto(l.r.estado !== "guardado", "«guardado» con un flush fallido (" + l.p.operacion + ")");
    if (l.anidada || l.mutado) return;
    const visto = C.normalizarNuevo_(l.p.campo, l.p.periodo, l.p.nuevo).visto;
    if (l.r.estado === "guardado") {
      igual(l.despues, visto, "«guardado» sin el valor en la celda (" + l.p.operacion + ")");
      cierto(porOp[l.p.operacion] && porOp[l.p.operacion].f[COL.Estado] === "aplicado", "«guardado» sin fila «aplicado»: " + detalle(l.p.operacion));
    }
    if (l.r.estado === "coincide") igual(l.despues, visto, "«coincide» sin el valor en la celda (" + l.p.operacion + ")");
  });
  return { inferidas, filas: filas.length };
}
prueba("invariantes: lock liberado siempre tras un flush; ningún «guardado» con un flush fallido; reenvío alterado de una operación registrada siempre rechazado sin escribir; una fila por operación; a lo sumo un intento de escritura por operación, siempre precedido por su fila; «aplicado» sólo con escritura propia; «guardado» sólo con «aplicado» y el valor en la celda; «reintentar» y «rechazado» nunca escriben", () => {
  const tot = { pedidos: 0, llamadas: 0, reintentos: 0, perdidas: 0, fallas: 0, equipo: 0, estructurales: 0, duplicadas: 0, configuracion: 0, alterados: 0, flush: 0, inferidas: 0, filas: 0, respuestas: {}, estados: {} };
  for (let c = 0; c < CORRIDAS; c++) {
    const res = simular(1000 + c, 120);
    const inv = invariantes(res);
    Object.keys(res.est).forEach(k => { tot[k] += res.est[k]; });
    tot.inferidas += inv.inferidas; tot.filas += inv.filas;
    res.llamadas.forEach(l => { tot.respuestas[l.r.estado] = (tot.respuestas[l.r.estado] || 0) + 1; });
    estados(res.s).forEach(x => { tot.estados[x] = (tot.estados[x] || 0) + 1; });
  }
  console.log("      " + JSON.stringify(tot));
});

// ================================================================ 10. adaptador con dobles de los servicios de Google

seccion("10. Adaptador (Carga_Google.gs) con dobles de SpreadsheetApp, LockService, PropertiesService, Utilities y Session");
function google(s, extra) {
  const control = Array.from({ length: 14 }, () => ["", ""]);
  CARGA.PERIODOS.forEach(p => { control[Number(CARGA.CELDAS_FUENTE[p].slice(1)) - 1][1] = s.control[p]; });
  const fecha = t => { const f = fechaCivilPrueba(t); return f ? new Date(f + "T00:00:00-03:00") : t; };
  const conFechas = filas => filas.map((f, i) => (i === 0 ? f : f.map((x, j) => (j > 0 || filas[0][0] === "Fecha" ? fecha(x) : x))));
  const estado = {
    props: { ADMIN_ID: "ADMIN_ficticia_000000000001", GENERADOR_ID: "GENERADOR_ficticio_00000001" },
    aperturas: [], lock: s.lock, propietario: "eps@cfp7.edu.ar", usuario: "eps@cfp7.edu.ar",
    libros: {
      ADMIN_ficticia_000000000001: { zona: "America/Argentina/Buenos_Aires", hojas: { Ventanas: new Hoja(conFechas(s.ventanas)), Feriados: new Hoja(conFechas(s.feriados)), Asignaciones: new Hoja(s.asignaciones), Registro: s.registro } },
      GENERADOR_ficticio_00000001: { zona: "America/Argentina/Buenos_Aires", hojas: { Control: new Hoja(control) } }
    }
  };
  Object.keys(s.grillas).forEach(id => { estado.libros[id] = { zona: "America/Argentina/Buenos_Aires", hojas: s.grillas[id] }; });
  Object.assign(estado, extra || {});
  const libro = id => {
    const l = estado.libros[id];
    if (!l) throw new Error("No se encontró la planilla " + id);
    estado.aperturas.push(id);
    return { getSheetByName: n => l.hojas[n] || null, getSpreadsheetTimeZone: () => l.zona, insertSheet: n => (l.hojas[n] = new Hoja([[""]])) };
  };
  const globales = {
    SpreadsheetApp: { openById: libro, flush: () => {} },
    LockService: { getScriptLock: () => estado.lock },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => estado.props[k] || null }) },
    Utilities: { formatDate: (d, zona, patron) => { if (patron !== "yyyy-MM-dd") throw new Error("patrón no simulado"); return new Intl.DateTimeFormat("en-CA", { timeZone: zona, year: "numeric", month: "2-digit", day: "2-digit" }).format(d); } },
    Session: { getEffectiveUser: () => ({ getEmail: () => estado.propietario }), getActiveUser: () => ({ getEmail: () => estado.usuario }) }
  };
  globales.console = { log: () => {}, error: () => {} };
  const ctx = contexto([NUCLEO, "src/carga/Carga_Google.gs"], globales);
  return { ctx, estado, control: estado.libros.GENERADOR_ficticio_00000001.hojas.Control };
}
function fechaCivilPrueba(t) { const m = String(t).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if (m) return m[3] + "-" + m[2].padStart(2, "0") + "-" + m[1].padStart(2, "0"); return /^\d{4}-\d{2}-\d{2}$/.test(t) ? t : ""; }
// El adaptador usa el reloj real; las ventanas de la prueba se corren para que hoy caiga dentro de 2C y anual.
function ventanasAlrededorDeHoy(s) {
  const hoy = new Date();
  const iso = d => new Intl.DateTimeFormat("en-CA", { timeZone: "America/Argentina/Buenos_Aires", year: "numeric", month: "2-digit", day: "2-digit" }).format(d);
  const desde = iso(new Date(hoy.getTime() - 3 * 86400000)), hasta = iso(new Date(hoy.getTime() + 10 * 86400000));
  s.ventanas = [VENTANAS[0], ["2º cuatrimestre", desde, hasta], ["Anual", desde, hasta]];
  const anio = Number(iso(hoy).slice(0, 4));
  s.feriados = [FERIADOS[0], [anio + "-12-25", "Navidad"], [(anio + 1) + "-01-01", "Año Nuevo"]];
}
prueba("consulta y guardado de punta a punta: fechas de Sheets, fuente leída de Control!B13 por su URL, y la grilla abierta por ese ID", () => {
  const s = escenario(); ventanasAlrededorDeHoy(s);
  const { ctx, estado } = google(s);
  const vista = ctx.consultarCarga_(MATE);
  igual([vista.estado, vista.periodos.map(p => [p.periodo, p.fuente])], ["ok", [[P2C, ID[P2C]], [PAN, ID[PAN]]]]);
  const e = vista.periodos[0].materias.find(m => m.materia === "Matemática II").estudiantes.find(x => x.situacion === "pendiente" && !x.problema);
  const r = ctx.guardarCampoCarga_(MATE, { operacion: op(), periodo: P2C, fuente: ID[P2C], orientacion: "EPSI", materia: C.clave_("Matemática II"), dni: e.dni, campo: "calificacion", esperado: "CURSA", nuevo: "7" });
  igual(r.estado, "guardado"); cierto(typeof r.ms === "number");
  igual([hoja(s, P2C).escrituras.length, hoja(s, PAN).escrituras.length, estados(s)], [1, 0, ["aplicado"]]);
  cierto(estado.aperturas.every(id => [ID[P2C], ID[PAN], "ADMIN_ficticia_000000000001", "GENERADOR_ficticio_00000001"].includes(id)), estado.aperturas.join());
});
prueba("cambio de Control!B13 con la página abierta: el adaptador relee Control y el guardado con la fuente anterior se rechaza", () => {
  const s = escenario(); ventanasAlrededorDeHoy(s);
  const { ctx, control } = google(s);
  const e = ctx.consultarCarga_(MATE).periodos[0].materias.find(m => m.materia === "Matemática II").estudiantes.find(x => x.situacion === "pendiente" && !x.problema);
  control.v[12][1] = URL(OTRA_2C);
  const r = ctx.guardarCampoCarga_(MATE, { operacion: op(), periodo: P2C, fuente: ID[P2C], orientacion: "EPSI", materia: C.clave_("Matemática II"), dni: e.dni, campo: "calificacion", esperado: "CURSA", nuevo: "7" });
  igual([r.estado, r.codigo, todasLasEscrituras(s).length], ["rechazado", "fuente", 0]);
});
prueba("sin propiedades del script o sin Control: error explícito en la consulta y «reintentar» en el guardado, sin escribir", () => {
  const s = escenario(); ventanasAlrededorDeHoy(s);
  const { ctx, estado } = google(s);
  estado.props = {};
  igual(ctx.consultarCarga_(MATE).estado, "error");
  igual(ctx.guardarCampoCarga_(MATE, { operacion: op() }).estado, "reintentar");
  const g2 = google(s); delete g2.estado.libros.GENERADOR_ficticio_00000001.hojas.Control;
  igual(g2.ctx.consultarCarga_(MATE).estado, "error");
  igual(todasLasEscrituras(s).length, 0);
});
prueba("funciones del propietario: preparar crea sólo las hojas que faltan; verificar informa materias sin docente y filas a revisar; otro usuario no puede", () => {
  const s = escenario(); ventanasAlrededorDeHoy(s);
  const { ctx, estado } = google(s);
  delete estado.libros.ADMIN_ficticia_000000000001.hojas.Feriados;
  const asign = estado.libros.ADMIN_ficticia_000000000001.hojas.Asignaciones; const antes = copia(asign);
  ctx.prepararAdministracionCarga();
  igual(estado.libros.ADMIN_ficticia_000000000001.hojas.Feriados.v[0], ["Fecha", "Descripción"]);
  igual(diferencias(antes, asign), []);
  const r = ctx.verificarCargaDocente();
  cierto(r.problemas.some(p => /Faltan los feriados/.test(p)), "sin feriados, el cierre no se puede calcular");
  cierto(r.problemas.some(p => /ningún docente asignado/.test(p)), r.problemas.slice(0, 3).join(" | "));
  estado.usuario = "visitante@bue.edu.ar";
  lanza(() => ctx.verificarCargaDocente(), /propietario/);
  lanza(() => ctx.resolverPendientesCarga(), /propietario/);
});

console.log("\n" + (fallas ? fallas + " de " + total + " pruebas fallaron" : "Las " + total + " pruebas pasaron"));
process.exit(fallas ? 1 : 0);
