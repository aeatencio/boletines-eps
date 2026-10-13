// Arnés de pruebas local para v0.3
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { execFileSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const WORK = path.join(ROOT, "tmp", "unzipped");

function initWorkDir() {
  if (fs.existsSync(WORK)) {
    fs.rmSync(WORK, { recursive: true, force: true });
  }
  fs.mkdirSync(WORK, { recursive: true });
}

function unzip(rel) {
  const srcPath = path.join(ROOT, rel);
  if (!fs.existsSync(srcPath)) {
    throw new Error(`Archivo requerido no encontrado: ${rel}`);
  }
  const dest = path.join(WORK, path.basename(rel, ".xlsx"));
  if (fs.existsSync(dest)) {
    fs.rmSync(dest, { recursive: true, force: true });
  }
  fs.mkdirSync(dest, { recursive: true });
  execFileSync("tar.exe", ["-xf", srcPath, "-C", dest]);
  return dest;
}

function colNum(letters) {
  return letters.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

function readXlsx(rel) {
  const dir = unzip(rel);
  const rd = p => fs.readFileSync(path.join(dir, p), "utf8");
  const ssPath = path.join(dir, "xl/sharedStrings.xml");
  const shared = fs.existsSync(ssPath)
    ? [...rd("xl/sharedStrings.xml").matchAll(/<(?:\w+:)?si>([\s\S]*?)<\/(?:\w+:)?si>/g)]
        .map(m => [...m[1].matchAll(/<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/g)].map(t => t[1]).join(""))
    : [];
  const wb = rd("xl/workbook.xml");
  const rels = rd("xl/_rels/workbook.xml.rels");
  const out = {};
  for (const m of wb.matchAll(/<(?:\w+:)?sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
    const t = rels.match(new RegExp("Target=\"([^\"]+)\"[^>]*Id=\"" + m[2] + "\"")) || rels.match(new RegExp("Id=\"" + m[2] + "\"[^>]*Target=\"([^\"]+)\""));
    let p = t[1].replace(/^\//, "");
    if (!p.startsWith("xl/")) p = "xl/" + p;
    const x = rd(p);
    const grid = [];
    for (const r of x.matchAll(/<(?:\w+:)?row [^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/(?:\w+:)?row>/g)) {
      for (const c of r[2].matchAll(/<(?:\w+:)?c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g)) {
        const body = c[4] || "";
        let v = (body.match(/<(?:\w+:)?v>([^<]*)</) || [])[1];
        const is = (body.match(/<(?:\w+:)?t[^>]*>([^<]*)</) || [])[1];
        if (/t="s"/.test(c[3]) && v !== undefined) v = shared[+v];
        if (v === undefined) v = is;
        if (v === undefined || v === "") continue;
        const rr = +c[2] - 1, cc = colNum(c[1]) - 1;
        while (grid.length <= rr) grid.push([]);
        grid[rr][cc] = v.replace(/&amp;/g, "&").replace(/&gt;/g, ">").replace(/&lt;/g, "<");
      }
    }
    out[m[1]] = grid;
  }
  return out;
}

function rect(grid) {
  const w = Math.max(...grid.map(r => r.length));
  return grid.map(r => Array.from({ length: w }, (_, i) => r[i] === undefined ? "" : String(r[i])));
}

// Mocks
function a1ToRC(a1) {
  const [a, b] = a1.split(":");
  const p = s => { const m = s.match(/^([A-Z]+)(\d+)$/); return [+m[2], colNum(m[1])]; };
  const [r1, c1] = p(a);
  const [r2, c2] = b ? p(b) : [r1, c1];
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}

class MockSheet {
  constructor(name, grid) {
    this.name = name;
    this.grid = grid.map(r => r.slice());
  }
  getName() { return this.name; }
  setName(n) { this.name = n; }
  get(r, c) { return ((this.grid[r - 1] || [])[c - 1]) ?? ""; }
  set(r, c, v) {
    while (this.grid.length < r) this.grid.push([]);
    this.grid[r - 1][c - 1] = v;
  }
  getRange(a, b, nr, nc) {
    const [r, c, h, w] = typeof a === "string" ? a1ToRC(a) : [a, b, nr || 1, nc || 1];
    const sh = this;
    return {
      getDisplayValue: () => String(sh.get(r, c)),
      setValue: v => {
        for (let i = 0; i < h; i++) {
          for (let j = 0; j < w; j++) {
            sh.set(r + i, c + j, v instanceof Date ? "<fecha>" : v);
          }
        }
      },
      clearContent: () => {
        for (let i = 0; i < h; i++) {
          for (let j = 0; j < w; j++) {
            sh.set(r + i, c + j, "");
          }
        }
      },
      setWrap: () => {},
      copyTo: () => {},
      getDataRange: undefined
    };
  }
  getDataRange() {
    const g = rect(this.grid);
    return { getDisplayValues: () => g.map(r => r.slice()) };
  }
  insertRowsAfter(row, n) {
    this.grid.splice(row, 0, ...Array.from({ length: n }, () => []));
  }
  deleteRows(row, n) {
    this.grid.splice(row - 1, n);
  }
  setRowHeight() {}
  copyTo(ss) {
    const s = new MockSheet("Copia de " + this.name, this.grid);
    ss.sheets.push(s);
    return s;
  }
}

class MockSS {
  constructor(sheets) { this.sheets = sheets; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
  setActiveSheet() {}
}

function hojaControl(control) {
  const ctrlGrid = [];
  for (const [a1, v] of Object.entries(control)) {
    const [r, c] = a1ToRC(a1);
    while (ctrlGrid.length < r) ctrlGrid.push([]);
    ctrlGrid[r - 1][c - 1] = v;
  }
  return new MockSheet("Control", ctrlGrid);
}

// Las grillas fuente son de sólo lectura para el generador: el mock no expone escritura.
function abrirFuenteMock(sources) {
  return id => {
    if (!sources[id]) throw new Error("mock: id inexistente " + id);
    return {
      getSheetByName: n => {
        if (!sources[id][n]) return null;
        const hoja = new MockSheet(id + "/" + n, sources[id][n]);
        return { getName: () => hoja.getName(), getDataRange: () => hoja.getDataRange() };
      }
    };
  };
}

function runTest(testName, { control, sources, templateGrid }) {
  const code = fs.readFileSync(path.join(ROOT, "src/Codigo_Apps_Script_Generador_EPS_v0_3.gs"), "utf8");
  const plantilla = new MockSheet("Plantilla", templateGrid);
  const ctrl = hojaControl(control);
  const ss = new MockSS([plantilla, ctrl]);
  const rawAlerts = [];
  const ctx = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ({ alert: m => rawAlerts.push(m), createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }),
      openById: abrirFuenteMock(sources)
    }
  };
  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  let error = null;
  try {
    vm.runInContext("generarInformeActual()", ctx);
  } catch (e) {
    error = e.message;
  }
  const out = ss.getSheetByName("Informe_actual");

  // Parsear advertencias funcionales reportadas en el diálogo de UI
  const advertencias = [];
  for (const msg of rawAlerts) {
    const match = msg.match(/Advertencias:\n- ([\s\S]+)$/);
    if (match) {
      advertencias.push(...match[1].split("\n- "));
    }
  }
  const estadoAdvertencias = ctrl.get(4, 6); // Celda F4 de Control
  const estadoResultado = ctrl.get(3, 6); // Celda F3 de Control

  return {
    error,
    rawAlerts,
    advertencias,
    estadoAdvertencias,
    estadoResultado,
    out: out ? rect(out.grid) : null
  };
}

// Carga el generador sin ejecutarlo, para probar funciones puras del parser.
// Con fuentes, también puede abrirlas (sin hoja activa ni UI).
function cargarGenerador(sources) {
  const ctx = sources ? { SpreadsheetApp: { openById: abrirFuenteMock(sources) } } : {};
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(ROOT, "src/Codigo_Apps_Script_Generador_EPS_v0_3.gs"), "utf8"), ctx);
  return ctx;
}

// Etapa de datos de la generación, sin crear Informe_actual: expone el resultado interno.
function generarDatos(control, sources) {
  const gen = cargarGenerador(sources);
  const advertencias = [];
  try {
    const informe = gen.generarDatosInforme_(gen.leerConfiguracion_(hojaControl(control)), advertencias);
    return { error: null, informe: JSON.parse(JSON.stringify(informe)), advertencias: advertencias.slice() };
  } catch (e) {
    return { error: e.message, informe: null, advertencias: advertencias.slice() };
  }
}

// Filas de materias impresas en Informe_actual: desde la fila 9 hasta el bloque de observaciones.
function filasMaterias(out) {
  const filas = [];
  for (let i = 8; i < (out || []).length; i++) {
    if (/^OBSERVACIONES GENERALES/.test(out[i][2])) break;
    if (out[i][2]) filas.push(out[i]);
  }
  return filas;
}

function textoBajoEtiqueta(out, prefijo) {
  const i = (out || []).findIndex(f => f[2].startsWith(prefijo));
  return i >= 0 && out[i + 1] ? out[i + 1][3] : null;
}

function readTsv(p) {
  const srcPath = path.join(ROOT, p);
  if (!fs.existsSync(srcPath)) {
    throw new Error(`Archivo TSV requerido no existe: ${p}`);
  }
  return fs.readFileSync(srcPath, "utf8").split("\n").map(l => l.split("\t"));
}

// Inicializar directorio de trabajo limpio para garantizar que cada corrida lea los fixtures vigentes
initWorkDir();

const tpl03 = readXlsx("src/Generador_informes_individuales_EPS_v0_3.xlsx").Plantilla;
const src1B = readXlsx("reference/fixtures/Valoraciones_ficticias_EPS_1er_bimestre_v0_3.xlsx");
const src1C = readXlsx("reference/fixtures/Valoraciones_ficticias_EPS_1er_cuatrimestre_v0_3.xlsx");
const src3B = { EPSI: readTsv("reference/fixtures/3er_bimestre.tsv") };
const src3B_inv = { EPSI: readTsv("reference/fixtures/3er_bimestre_invalido.tsv") };
const src3B_real = {
  EPSI: readTsv("reference/fixtures/3er_bimestre_estructura_real_EPSI.tsv"),
  EPSG: readTsv("reference/fixtures/3er_bimestre_estructura_real_EPSG.tsv")
};
const src1C_real = { EPSI: readTsv("reference/fixtures/1er_cuatrimestre_estructura_real_EPSI.tsv") };

const sources = {
  "ID_1B_xxxxxxxxxxxxxxxxxxxxxxxx": src1B,
  "ID_1C_xxxxxxxxxxxxxxxxxxxxxxxx": src1C,
  "ID_3B_xxxxxxxxxxxxxxxxxxxxxxxx": src3B,
  "ID_3B_INV_xxxxxxxxxxxxxxxxxxxx": src3B_inv,
  "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx": src3B_real,
  "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx": src1C_real
};

const ctl = (p, d, extra) => Object.assign({
  B5: p,
  B6: "EPSI",
  B7: d,
  B8: "2026",
  B10: "ID_1B_xxxxxxxxxxxxxxxxxxxxxxxx",
  B12: "ID_1C_xxxxxxxxxxxxxxxxxxxxxxxx",
  B13: "ID_3B_xxxxxxxxxxxxxxxxxxxxxxxx"
}, extra);

let fail = 0;
function assert(desc, cond, msg) {
  if (!cond) {
    fail++;
    console.error("FAIL:", desc, msg !== undefined ? "-" : "", msg ?? "");
  } else {
    console.log("OK:", desc);
  }
}

// =============================================================================
// 1. Happy path 3B (Camila Uno, DNI 99000001 · 1B + 1C + 3B)
// =============================================================================
let r = runTest("Happy path 3B", { control: ctl("3º bimestre", "99000001"), sources, templateGrid: tpl03 });
assert("Happy path 3B: sin errores", !r.error, r.error);
assert("Happy path 3B: sin advertencias", r.advertencias.length === 0, r.advertencias.join("; "));
assert("Happy path 3B: Control registra 'Sin advertencias'", r.estadoAdvertencias === "Sin advertencias", r.estadoAdvertencias);
assert("Happy path 3B: nivel visual es 'Informática 1'", r.out && r.out[5] && r.out[5][4] === "Informática 1", r.out ? r.out[5][4] : null);

let filaMat = (r.out || []).find(f => f[2] === "Matemática I") || [];
assert("Happy path 3B: Matemática I acumula los 3 períodos (Suficiente|7|Avanzado)", filaMat.slice(4, 7).join("|") === "Suficiente|7|Avanzado", filaMat.slice(4, 7));

let filaLen = (r.out || []).find(f => f[2] === "Lengua y Literatura I") || [];
assert("Happy path 3B: Lengua y Literatura I acumula 1B y 3B (En proceso||Suficiente)", filaLen.slice(4, 7).join("|") === "En proceso||Suficiente", filaLen.slice(4, 7));

let filaFGI = (r.out || []).find(f => f[2] === "FGI I") || [];
assert("Happy path 3B: FGI I acumula los 3 períodos (ACREDITÓ|ACREDITÓ|ACREDITÓ)", filaFGI.slice(4, 7).join("|") === "ACREDITÓ|ACREDITÓ|ACREDITÓ", filaFGI.slice(4, 7));

// =============================================================================
// 2. Materia que aparece por primera vez en 3B (Carla Cuatro, DNI 99000003)
// =============================================================================
r = runTest("Materia que aparece por primera vez en 3B", { control: ctl("3º bimestre", "99000003"), sources, templateGrid: tpl03 });
assert("Materia incorporada en 3B: sin errores", !r.error, r.error);
assert("Materia incorporada en 3B: nivel visual es 'Informática 1'", r.out && r.out[5] && r.out[5][4] === "Informática 1", r.out ? r.out[5][4] : null);
let filaIncorporada = (r.out || []).find(f => f[2] === "Taller C") || [];
assert("Materia incorporada en 3B: Taller C solo tiene valor en 3B (||Suficiente)", filaIncorporada.slice(4, 7).join("|") === "||Suficiente", filaIncorporada.slice(4, 7));

// =============================================================================
// 3. Caso separado: Valoración inválida en 3B (descarte + advertencia + continuidad)
// =============================================================================
let ctlInv = ctl("3º bimestre", "99000001", { B13: "ID_3B_INV_xxxxxxxxxxxxxxxxxxxx" });
let rInv = runTest("Valoración inválida 3B", { control: ctlInv, sources, templateGrid: tpl03 });
assert("Valoración inválida 3B: no interrumpe ejecución", !rInv.error && rInv.out, rInv.error);

let filaLenInv = (rInv.out || []).find(f => f[2] === "Lengua y Literatura I") || [];
assert("Valoración inválida 3B: valor inválido es descartado y conserva 1B (En proceso||)", filaLenInv.slice(4, 7).join("|") === "En proceso||", filaLenInv.slice(4, 7));

const advEsperada = 'Valor inesperado en 3º bimestre para "Lengua y Literatura I": "Bueno"';
assert("Valoración inválida 3B: emite advertencia específica", rInv.advertencias.some(a => a.includes(advEsperada)), rInv.advertencias.join("; "));
assert("Valoración inválida 3B: Control registra la advertencia", rInv.estadoAdvertencias.includes(advEsperada), rInv.estadoAdvertencias);

let filaMatInv = (rInv.out || []).find(f => f[2] === "Matemática I") || [];
assert("Valoración inválida 3B: continuidad (Matemática I acumula correctamente)", filaMatInv.slice(4, 7).join("|") === "Suficiente|7|Avanzado", filaMatInv.slice(4, 7));

// =============================================================================
// 4. Regresión de períodos anteriores (matriz histórica proporcional a v0.2)
// =============================================================================

// 4.1. Regresión 1B (CP-01 de v0.2: Camila Uno, DNI 99000001, EPSI)
let r1B = runTest("Regresión 1B (CP-01)", { control: ctl("1º bimestre", "99000001", { B12: "", B13: "" }), sources, templateGrid: tpl03 });
assert("Regresión 1B: genera sin errores", !r1B.error && r1B.out, r1B.error);
assert("Regresión 1B: sin advertencias", r1B.advertencias.length === 0, r1B.advertencias.join("; "));
assert("Regresión 1B: Control registra 'Sin advertencias'", r1B.estadoAdvertencias === "Sin advertencias", r1B.estadoAdvertencias);
assert("Regresión 1B: nivel visual es 'Informática 1'", r1B.out && r1B.out[5] && r1B.out[5][4] === "Informática 1", r1B.out ? r1B.out[5][4] : null);
let f1bMat = (r1B.out || []).find(f => f[2] === "Matemática I") || [];
assert("Regresión 1B: Matemática I solo valor en columna 1B (Suficiente||)", f1bMat.slice(4, 7).join("|") === "Suficiente||", f1bMat.slice(4, 7));
let f1bLen = (r1B.out || []).find(f => f[2] === "Lengua y Literatura I") || [];
assert("Regresión 1B: Lengua y Literatura I solo valor en columna 1B (En proceso||)", f1bLen.slice(4, 7).join("|") === "En proceso||", f1bLen.slice(4, 7));
let f1bFGI = (r1B.out || []).find(f => f[2] === "FGI I") || [];
assert("Regresión 1B: FGI I solo valor en columna 1B (ACREDITÓ||)", f1bFGI.slice(4, 7).join("|") === "ACREDITÓ||", f1bFGI.slice(4, 7));

// 4.2. Regresión 1C acumulativo (CP-02 de v0.2: Camila Uno, DNI 99000001, EPSI · 1B + 1C)
let r1C = runTest("Regresión 1C acumulativo (CP-02)", { control: ctl("1º cuatrimestre", "99000001", { B13: "" }), sources, templateGrid: tpl03 });
assert("Regresión 1C: genera sin errores", !r1C.error && r1C.out, r1C.error);
assert("Regresión 1C: sin advertencias", r1C.advertencias.length === 0, r1C.advertencias.join("; "));
assert("Regresión 1C: Control registra 'Sin advertencias'", r1C.estadoAdvertencias === "Sin advertencias", r1C.estadoAdvertencias);
assert("Regresión 1C: nivel visual es 'Informática 1'", r1C.out && r1C.out[5] && r1C.out[5][4] === "Informática 1", r1C.out ? r1C.out[5][4] : null);
let f1cMat = (r1C.out || []).find(f => f[2] === "Matemática I") || [];
assert("Regresión 1C: Matemática I acumula 1B y 1C (Suficiente|7|)", f1cMat.slice(4, 7).join("|") === "Suficiente|7|", f1cMat.slice(4, 7));
let f1cLen = (r1C.out || []).find(f => f[2] === "Lengua y Literatura I") || [];
assert("Regresión 1C: Lengua y Literatura I solo en 1B (En proceso||)", f1cLen.slice(4, 7).join("|") === "En proceso||", f1cLen.slice(4, 7));
let f1cSoc = (r1C.out || []).find(f => f[2] === "Ciencias Sociales I") || [];
assert("Regresión 1C: Ciencias Sociales I solo en 1C (|8|)", f1cSoc.slice(4, 7).join("|") === "|8|", f1cSoc.slice(4, 7));
let f1cFGI = (r1C.out || []).find(f => f[2] === "FGI I") || [];
assert("Regresión 1C: FGI I acumula acreditación en 1B y 1C (ACREDITÓ|ACREDITÓ|)", f1cFGI.slice(4, 7).join("|") === "ACREDITÓ|ACREDITÓ|", f1cFGI.slice(4, 7));
let f1cArq = (r1C.out || []).find(f => f[2] === "Arquitectura de Sistemas Informáticos") || [];
assert("Regresión 1C: Arquitectura solo en 1C (|8|)", f1cArq.slice(4, 7).join("|") === "|8|", f1cArq.slice(4, 7));

// 4.3. Regresión 1C valor no válido (CP-03 de v0.2: Bruno Dos, DNI 99000002 · valor '7,5' en 1C)
let r1C_inv = runTest("Regresión 1C valor no válido (CP-03)", { control: ctl("1º cuatrimestre", "99000002", { B13: "" }), sources, templateGrid: tpl03 });
assert("Regresión 1C valor no válido: genera sin abortar", !r1C_inv.error && r1C_inv.out, r1C_inv.error);
let f1cBrunoLen = (r1C_inv.out || []).find(f => f[2] === "Lengua y Literatura I") || [];
assert("Regresión 1C valor no válido: descarta '7,5' y conserva 1B (Suficiente||)", f1cBrunoLen.slice(4, 7).join("|") === "Suficiente||", f1cBrunoLen.slice(4, 7));
assert("Regresión 1C valor no válido: emite advertencia por 7,5", r1C_inv.advertencias.some(a => a.includes("7,5")), r1C_inv.advertencias.join("; "));
let f1cBrunoMat = (r1C_inv.out || []).find(f => f[2] === "Matemática I") || [];
assert("Regresión 1C valor no válido: continuidad (Matemática I tiene |8|)", f1cBrunoMat.slice(4, 7).join("|") === "|8|", f1cBrunoMat.slice(4, 7));


// 4.4. Regresión 1C Orientación Gastronomía (CP-06 de v0.2: Nora Uno, DNI 98000001, EPSG)
let rGastro = runTest("Regresión 1C Gastronomía (CP-06)", {
  control: ctl("1º cuatrimestre", "98000001", { B6: "EPSG", B13: "" }),
  sources,
  templateGrid: tpl03
});
assert("Regresión 1C Gastronomía: genera sin errores", !rGastro.error && rGastro.out, rGastro.error);
assert("Regresión 1C Gastronomía: nivel visual es 'Gastronomía 1'", rGastro.out && rGastro.out[5] && rGastro.out[5][4] === "Gastronomía 1", rGastro.out ? rGastro.out[5][4] : null);
let fGastroAlim = (rGastro.out || []).find(f => f[2] === "Manipulación de alimentos") || [];
assert("Regresión 1C Gastronomía: Manipulación de alimentos acumula 1B y 1C (Avanzado|10|)", fGastroAlim.slice(4, 7).join("|") === "Avanzado|10|", fGastroAlim.slice(4, 7));

// =============================================================================
// 5. Estructura real de las grillas (#10): materias, cola y orden
// =============================================================================
const gen = cargarGenerador();
const P3B = "3º bimestre";
const nombresMaterias = est => est.materias.map(m => m.nombre).join("|");
const valoresFila = f => f.slice(4, 7).join("|");

// 5.1. Detección explícita de materias y cola general/operativa
const estReal = gen.detectarEstructuraFuente_(src3B_real.EPSI, P3B);
assert("Estructura real: 16 materias de los cuatro niveles en orden de grilla, con Formación Profesional por subencabezado",
  nombresMaterias(estReal) === [
    "Matemática I", "Lengua y Literatura I", "Taller Profesional I A", "Taller Profesional I B", "FGI I",
    "Matemática II", "Inglés I", "Taller Profesional II A", "Taller Profesional II B", "FGI II",
    "Matemática III", "Taller Profesional III A", "FGI III",
    "Matemática IV", "Taller Profesional IV A", "FGI IV"
  ].join("|"), nombresMaterias(estReal));
assert("Estructura real: cada materia lee la valoración en la columna siguiente a su etiqueta",
  estReal.materias.every(m => m.columnaValor === m.columnaEtiqueta + 1), JSON.stringify(estReal.materias[0]));
const colaReal = gen.detectarCamposCola_(src3B_real.EPSI, estReal, P3B, []);
assert("Estructura real: la cola empieza en ORIENTACIÓN (columna AN)",
  estReal.inicioCola === 39 && colaReal.orientacion === 39, estReal.inicioCola);
assert("Estructura real: campos generales detectados dentro de la cola",
  [colaReal.observacionesGenerales, colaReal.observacionesAsistencia,
   colaReal.observacionesConvivencia, colaReal.ptf].join(",") === "40,41,42,43",
  JSON.stringify(colaReal));
assert("Estructura real: ninguna columna desde ORIENTACIÓN en adelante es materia (LISTO PARA EXPORTAR, EXPORTADO, columnas posteriores)",
  estReal.materias.every(m => m.columnaEtiqueta < estReal.inicioCola && m.columnaValor < estReal.inicioCola));
assert("Estructura real: encabezados de grupo y marcadores de nivel no son materias",
  !estReal.materias.some(m => /^Formación Profesional|^EPS\d/.test(m.nombre)), nombresMaterias(estReal));

const estV03 = gen.detectarEstructuraFuente_(src3B.EPSI, P3B);
assert("Estructura v0.3: el fixture de 3B anterior conserva sus 6 materias",
  nombresMaterias(estV03) === "Matemática I|Lengua y Literatura I|Ciencias Sociales I|FGI I|Arquitectura de Sistemas Informáticos|Taller C",
  nombresMaterias(estV03));

// 5.2. EQUIVALENCIA, CURSA, NO CURSA y vacío
const clasif = (p, v) => { const c = gen.clasificarValor_(p, v); return c.estado + ":" + c.valor; };
assert("EQUIVALENCIA es válida en 1B, 1C y 3B y se emite canónicamente",
  ["1º bimestre", "1º cuatrimestre", P3B].every(p => ["EQUIVALENCIA", "Equivalencia", " equivalencia "].every(v => clasif(p, v) === "valido:EQUIVALENCIA")),
  clasif("1º cuatrimestre", "Equivalencia"));
assert("CURSA, NO CURSA y vacío quedan excluidos del boletín",
  ["CURSA", "NO CURSA", "No cursa", ""].every(v => clasif(P3B, v) === "excluido:"));

// 5.3. Generación con la grilla de forma real (Ana Real, DNI 97000001, EPS2 · 1C + 3B)
const ctlReal = (d, extra) => ctl(P3B, d, Object.assign({
  B10: "",
  B12: "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx",
  B13: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx"
}, extra));
let rReal = runTest("Estructura real 3B", { control: ctlReal("97000001"), sources, templateGrid: tpl03 });
let matsReal = filasMaterias(rReal.out);
const filaReal = nombre => matsReal.find(f => f[2] === nombre) || [];
assert("Estructura real 3B: sin errores", !rReal.error && rReal.out, rReal.error);
assert("Estructura real 3B: nivel visual es 'Informática 2'", rReal.out && rReal.out[5][4] === "Informática 2", rReal.out ? rReal.out[5][4] : null);
assert("Estructura real 3B: 13 materias en orden de grilla principal y, al final, la que sólo existe en 1C",
  matsReal.map(f => f[2]).join("|") === [
    "Matemática I", "Lengua y Literatura I", "Taller Profesional I A", "Taller Profesional I B", "FGI I",
    "Matemática II", "Taller Profesional II A", "Taller Profesional II B", "FGI II", "Matemática III",
    "Matemática IV", "Taller Profesional IV A", "Taller Profesional I C"
  ].join("|"), matsReal.map(f => f[2]).join("|"));
assert("Estructura real 3B: EQUIVALENCIA se conserva en 1C y 3B (|EQUIVALENCIA|EQUIVALENCIA)",
  valoresFila(filaReal("Taller Profesional I B")) === "|EQUIVALENCIA|EQUIVALENCIA", valoresFila(filaReal("Taller Profesional I B")));
assert("Estructura real 3B: materia sólo en período anterior queda en su lugar de la grilla (Lengua y Literatura I |8|)",
  valoresFila(filaReal("Lengua y Literatura I")) === "|8|", valoresFila(filaReal("Lengua y Literatura I")));
assert("Estructura real 3B: materia de otro nivel sólo en período anterior (Matemática III |6|)",
  valoresFila(filaReal("Matemática III")) === "|6|", valoresFila(filaReal("Matemática III")));
assert("Estructura real 3B: materia sólo en período actual (Taller Profesional II B ||Avanzado)",
  valoresFila(filaReal("Taller Profesional II B")) === "||Avanzado", valoresFila(filaReal("Taller Profesional II B")));
assert("Estructura real 3B: materias de otro nivel acumulan (Matemática I |7|Avanzado, Matemática IV ||En proceso)",
  valoresFila(filaReal("Matemática I")) === "|7|Avanzado" && valoresFila(filaReal("Matemática IV")) === "||En proceso",
  valoresFila(filaReal("Matemática I")) + " / " + valoresFila(filaReal("Matemática IV")));
assert("Estructura real 3B: Formación Profesional se imprime por subencabezado (Taller Profesional I A |6|En proceso)",
  valoresFila(filaReal("Taller Profesional I A")) === "|6|En proceso", valoresFila(filaReal("Taller Profesional I A")));
assert("Estructura real 3B: materias sólo con CURSA, NO CURSA o vacío no aparecen",
  ["Inglés I", "Taller Profesional III A", "FGI III", "FGI IV"].every(n => !matsReal.some(f => f[2] === n)));
assert("Estructura real 3B: ninguna columna de la cola aparece como materia",
  !matsReal.some(f => /ORIENTACIÓN|OBSERVACIONES|PLAN DE TRABAJO|LISTO PARA EXPORTAR|EXPORTADO|SEGUIMIENTO INTERNO|VERIFICADO/.test(f[2])));
assert("Estructura real 3B: con más de ocho materias, las observaciones se desplazan sin pisar materias",
  textoBajoEtiqueta(rReal.out, "OBSERVACIONES GENERALES") === "Obs 3B Ana", textoBajoEtiqueta(rReal.out, "OBSERVACIONES GENERALES"));
assert("Estructura real 3B: sin advertencias salvo la falta de fuente de 1B",
  rReal.advertencias.join("\n") === "No se indicó archivo de 1º bimestre. El informe se generó sin valores anteriores de 1º bimestre.",
  rReal.advertencias.join("; "));

// 5.4. Estudiante nuevo, presente sólo en el período actual (Beto Nuevo, DNI 97000002, EPS1)
let rNuevo = runTest("Estudiante nuevo 3B", { control: ctlReal("97000002"), sources, templateGrid: tpl03 });
let matsNuevo = filasMaterias(rNuevo.out);
assert("Estudiante nuevo: genera sin errores", !rNuevo.error && rNuevo.out, rNuevo.error);
assert("Estudiante nuevo: nivel visual es 'Informática 1'", rNuevo.out && rNuevo.out[5][4] === "Informática 1", rNuevo.out ? rNuevo.out[5][4] : null);
assert("Estudiante nuevo: sólo valores de 3B (Matemática I ||Suficiente, Lengua y Literatura I ||En proceso, FGI I ||ACREDITÓ)",
  matsNuevo.map(f => f[2] + ":" + valoresFila(f)).join(", ") === "Matemática I:||Suficiente, Lengua y Literatura I:||En proceso, FGI I:||ACREDITÓ",
  matsNuevo.map(f => f[2] + ":" + valoresFila(f)).join(", "));
assert("Estudiante nuevo: la ausencia en 1C se informa como advertencia no bloqueante",
  rNuevo.advertencias.some(a => a.startsWith("No se encontró el DNI en la fuente de 1º cuatrimestre.")), rNuevo.advertencias.join("; "));

// 5.5. Hoja EPSG con la forma real (Gael Uno, DNI 96000001)
let rEpsg = runTest("Estructura real EPSG", {
  control: ctlReal("96000001", { B6: "EPSG", B12: "" }), sources, templateGrid: tpl03
});
assert("Estructura real EPSG: genera sin errores", !rEpsg.error && rEpsg.out, rEpsg.error);
assert("Estructura real EPSG: nivel visual es 'Gastronomía 1'", rEpsg.out && rEpsg.out[5][4] === "Gastronomía 1", rEpsg.out ? rEpsg.out[5][4] : null);
assert("Estructura real EPSG: materias y subencabezados de Formación Profesional en orden",
  filasMaterias(rEpsg.out).map(f => f[2]).join("|") === "Matemática I|Lengua y Literatura I|Taller Gastronómico I A|Taller Gastronómico I B|FGI I",
  filasMaterias(rEpsg.out).map(f => f[2]).join("|"));

// 5.6. Estructuras que no pueden identificarse de manera segura
const variante = (grid, cambiar) => { const g = grid.map(f => f.slice()); cambiar(g); return g; };
const columnaDe = (g, encabezado) => g[0].indexOf(encabezado);
const filaDni = (g, dni) => g.findIndex(f => f[2] === dni);
const conVariante = (id, fuente) => Object.assign({}, sources, { [id]: fuente });
const ID_VAR = "ID_VARIANTE_xxxxxxxxxxxxxxxxx";

let rSinOrient = runTest("Sin ORIENTACIÓN", {
  control: ctlReal("97000001", { B13: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "ORIENTACIÓN")] = ""; }) }),
  templateGrid: tpl03
});
assert("Sin ORIENTACIÓN en la principal: se informa error y no se genera",
  rSinOrient.error && rSinOrient.error.includes("No se encontró la columna ORIENTACIÓN") && !rSinOrient.out, rSinOrient.error);

let rOrientDoble = runTest("ORIENTACIÓN repetida", {
  control: ctlReal("97000001", { B13: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "SEGUIMIENTO INTERNO")] = "Orientación"; }) }),
  templateGrid: tpl03
});
assert("ORIENTACIÓN repetida en la principal: se informa error",
  rOrientDoble.error && rOrientDoble.error.includes("La columna ORIENTACIÓN aparece más de una vez (columnas AN, AU)"), rOrientDoble.error);

let rRepetida = runTest("Materia repetida en principal", {
  control: ctlReal("97000001", { B13: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "Matemática IV")] = "Matemática I"; }) }),
  templateGrid: tpl03
});
assert("Materia repetida en la principal: se informa error con ambas columnas, sin elegir una",
  rRepetida.error && rRepetida.error.includes('Materia repetida en la fuente de 3º bimestre: "Matemática I" (columnas E, AH)') && !rRepetida.out,
  rRepetida.error);

let rRepetidaHist = runTest("Materia repetida en fuente anterior", {
  control: ctlReal("97000001", { B12: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src1C_real.EPSI, g => { g[0][columnaDe(g, "Matemática IV")] = "Matemática I"; }) }),
  templateGrid: tpl03
});
assert("Materia repetida en una fuente anterior: advertencia, se descarta esa fuente y se genera con la principal",
  !rRepetidaHist.error && rRepetidaHist.advertencias.some(a => a.startsWith("Materia repetida en la fuente de 1º cuatrimestre") && a.endsWith("Se continuará con la fuente principal."))
    && valoresFila(filasMaterias(rRepetidaHist.out).find(f => f[2] === "Matemática I") || []) === "||Avanzado",
  rRepetidaHist.error || rRepetidaHist.advertencias.join("; "));

let rParRoto = runTest("Materia sin columna de valoración", {
  control: ctlReal("97000001", { B13: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "Matemática I") + 1] = "Columna intrusa"; }) }),
  templateGrid: tpl03
});
assert("Materia sin su columna de valoración: se informa error",
  rParRoto.error && rParRoto.error.includes('La materia "Matemática I" (columna E) no tiene a continuación su columna de valoración'), rParRoto.error);

let rFilaDni = runTest("DNI fuera de la fila Valoración", {
  control: ctlReal("97000001", { B13: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { const r = filaDni(g, "97000001"); g[r][2] = ""; g[r + 1][2] = "97000001"; }) }),
  templateGrid: tpl03
});
assert("DNI fuera de la fila Valoración del bloque: se informa error en vez de leer otra fila",
  rFilaDni.error && rFilaDni.error.includes('La fila del DNI no es la fila "Valoración"') && !rFilaDni.out, rFilaDni.error);

// 5.7. Nombre mostrado: el de la grilla principal si la materia existe en ella
let rNombre = runTest("Nombre de la grilla principal", {
  control: ctlReal("97000001", { B12: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src1C_real.EPSI, g => {
    g[0][columnaDe(g, "Lengua y Literatura I")] = "LENGUA Y LITERATURA I";
    g[1][g[1].indexOf("Taller Profesional I C")] = "TALLER PROFESIONAL I C";
  }) }),
  templateGrid: tpl03
});
let matsNombre = filasMaterias(rNombre.out).map(f => f[2] + ":" + valoresFila(f));
assert("Nombre mostrado: materia de la grilla principal valorada sólo en 1C usa el nombre de la principal",
  matsNombre.includes("Lengua y Literatura I:|8|") && !matsNombre.some(m => m.startsWith("LENGUA")), matsNombre.join(", "));
assert("Nombre mostrado: materia que no existe en la grilla principal conserva el nombre de la fuente anterior, al final",
  matsNombre[matsNombre.length - 1] === "TALLER PROFESIONAL I C:|8|", matsNombre.join(", "));

// =============================================================================
// 6. Campos generales y operativos: sólo desde la fuente principal (#10)
// =============================================================================
const P1C = "1º cuatrimestre";
const ID_VAR_1C = "ID_VARIANTE_1C_xxxxxxxxxxxxxx";
const ADV_SIN_1B = "No se indicó archivo de 1º bimestre. El informe se generó sin valores anteriores de 1º bimestre.";
const colDe = (g, encabezado) => {
  const c = columnaDe(g, encabezado);
  if (c < 0) throw new Error("El fixture no tiene el encabezado " + encabezado);
  return c;
};
const conCeldas = (dni, valoresPorEncabezado) => g => {
  const r = filaDni(g, dni);
  Object.entries(valoresPorEncabezado).forEach(([encabezado, v]) => { g[r][colDe(g, encabezado)] = v; });
};
const conEncabezados = reemplazos => g => {
  Object.entries(reemplazos).forEach(([encabezado, nuevo]) => { g[0][colDe(g, encabezado)] = nuevo; });
};
const variante3B = cambiar => conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, cambiar) });
const ctlAna3B = ctlReal("97000001", { B13: ID_VAR });
const cajas = out => ["OBSERVACIONES GENERALES", "OBSERVACIONES SOBRE ASISTENCIA", "OBSERVACIONES SOBRE CONVIVENCIA", "PLAN DE TRABAJO FORMATIVO"]
  .map(prefijo => textoBajoEtiqueta(out, prefijo)).join("|");
const operativos = d => d.informe ? JSON.stringify(d.informe.camposOperativos) : d.error;
const mismoInforme = (a, b) => !!a.out && JSON.stringify(a.out) === JSON.stringify(b.out);
const OP_SI_NO = '{"listoParaExportar":"SÍ","exportado":"NO"}';

// 6.1. Lectura desde la principal (Ana Real en 3B: LISTO PARA EXPORTAR = SÍ, EXPORTADO = NO)
let dAna = generarDatos(ctlReal("97000001"), sources);
assert("Campos operativos: LISTO PARA EXPORTAR = SÍ y EXPORTADO = NO quedan disponibles en camposOperativos",
  !dAna.error && operativos(dAna) === OP_SI_NO, operativos(dAna));
assert("Campos operativos: no forman parte de estudiante",
  dAna.informe && !("listoParaExportar" in dAna.informe.estudiante) && !("exportado" in dAna.informe.estudiante),
  dAna.informe && Object.keys(dAna.informe.estudiante).join(","));
assert("Campos operativos: valores canónicos no generan advertencias",
  dAna.advertencias.join("\n") === ADV_SIN_1B, dAna.advertencias.join("; "));
assert("Campos operativos: ausentes en la principal (grilla anterior a #10) quedan nulos, sin advertencia",
  (d => operativos(d) === '{"listoParaExportar":null,"exportado":null}' && d.advertencias.length === 0)(generarDatos(ctl(P3B, "99000001"), sources)));

// 6.2. Validación SÍ / NO
const estadoOp = v => { const adv = []; return { estado: gen.leerEstadoOperativo_([[v]], 0, 0, "LISTO PARA EXPORTAR", P3B, adv), adv }; };
assert("Campos operativos: SÍ y NO canónicos (también SÍ con tilde combinada) se aceptan sin advertencia",
  [["SÍ", "SÍ"], ["NO", "NO"], ["SÍ", "SÍ"]].every(([v, e]) => { const x = estadoOp(v); return x.estado === e && x.adv.length === 0; }));
assert("Campos operativos: Si, si, sí, Sí y SI se interpretan como SÍ con advertencia de forma no canónica",
  ["Si", "si", "sí", "Sí", "SI"].every(v => { const x = estadoOp(v); return x.estado === "SÍ" && x.adv.length === 1 && x.adv[0].includes('"' + v + '"') && x.adv[0].includes('la forma canónica es "SÍ"'); }));
assert("Campos operativos: No y no se interpretan como NO con advertencia de forma no canónica",
  ["No", "no"].every(v => { const x = estadoOp(v); return x.estado === "NO" && x.adv.length === 1 && x.adv[0].includes('la forma canónica es "NO"'); }));

let srcNoCanon = variante3B(conCeldas("97000001", { "LISTO PARA EXPORTAR": "Si", "EXPORTADO": "no" }));
let dNoCanon = generarDatos(ctlAna3B, srcNoCanon);
assert("Campos operativos: en la generación, 'Si' y 'no' quedan como SÍ y NO y cada uno advierte",
  operativos(dNoCanon) === OP_SI_NO
    && dNoCanon.advertencias.includes('LISTO PARA EXPORTAR tiene "Si" en la fuente de 3º bimestre. Se interpretó como SÍ, pero la forma canónica es "SÍ".')
    && dNoCanon.advertencias.includes('EXPORTADO tiene "no" en la fuente de 3º bimestre. Se interpretó como NO, pero la forma canónica es "NO".'),
  operativos(dNoCanon) + " / " + dNoCanon.advertencias.join("; "));

let srcOpVacio = variante3B(conCeldas("97000001", { "LISTO PARA EXPORTAR": "" }));
let rOpVacio = runTest("LISTO PARA EXPORTAR vacío", { control: ctlAna3B, sources: srcOpVacio, templateGrid: tpl03 });
let dOpVacio = generarDatos(ctlAna3B, srcOpVacio);
assert("Celda operativa vacía: advertencia no bloqueante y mismo informe",
  !rOpVacio.error && mismoInforme(rOpVacio, rReal)
    && rOpVacio.advertencias.includes("LISTO PARA EXPORTAR está vacío en la fuente de 3º bimestre. Se esperaba SÍ o NO."),
  rOpVacio.error || rOpVacio.advertencias.join("; "));
assert("Celda operativa vacía: el estado queda nulo",
  operativos(dOpVacio) === '{"listoParaExportar":null,"exportado":"NO"}', operativos(dOpVacio));

let srcOpInvalido = variante3B(conCeldas("97000001", { "EXPORTADO": "Tal vez" }));
let rOpInvalido = runTest("EXPORTADO inválido", { control: ctlAna3B, sources: srcOpInvalido, templateGrid: tpl03 });
let dOpInvalido = generarDatos(ctlAna3B, srcOpInvalido);
assert("Valor operativo inválido: advertencia no bloqueante y mismo informe",
  !rOpInvalido.error && mismoInforme(rOpInvalido, rReal)
    && rOpInvalido.advertencias.includes('EXPORTADO tiene un valor no admitido en la fuente de 3º bimestre. Valor: "Tal vez". Se esperaba SÍ o NO; el estado queda sin determinar.'),
  rOpInvalido.error || rOpInvalido.advertencias.join("; "));
assert("Valor operativo inválido: el estado queda nulo",
  operativos(dOpInvalido) === '{"listoParaExportar":"SÍ","exportado":null}', operativos(dOpInvalido));

let rOpDoble = runTest("EXPORTADO repetido", {
  control: ctlAna3B, sources: variante3B(conEncabezados({ "SEGUIMIENTO INTERNO": "Exportado" })), templateGrid: tpl03
});
assert("Encabezado operativo repetido: advertencia, no se elige una columna y se genera igual",
  !rOpDoble.error && mismoInforme(rOpDoble, rReal)
    && rOpDoble.advertencias.includes('El encabezado "EXPORTADO" aparece más de una vez (columnas AT, AU) en la fuente de 3º bimestre. No se leerá ese campo.'),
  rOpDoble.error || rOpDoble.advertencias.join("; "));

// 6.3. LISTO PARA EXPORTAR y EXPORTADO todavía no seleccionan ni alteran la generación individual
let rNoListo = runTest("LISTO PARA EXPORTAR = NO", {
  control: ctlAna3B, sources: variante3B(conCeldas("97000001", { "LISTO PARA EXPORTAR": "NO", "EXPORTADO": "SÍ" })), templateGrid: tpl03
});
assert("LISTO PARA EXPORTAR = NO: el boletín individual se genera igual que con SÍ",
  !rNoListo.error && mismoInforme(rNoListo, rReal) && rNoListo.advertencias.join("\n") === ADV_SIN_1B,
  rNoListo.error || rNoListo.advertencias.join("; "));
let dBeto = generarDatos(ctlReal("97000002"), sources);
assert("Estado inicial NO / NO (Beto Nuevo): se genera y el estado queda disponible sin advertencias operativas",
  !rNuevo.error && rNuevo.out && operativos(dBeto) === '{"listoParaExportar":"NO","exportado":"NO"}'
    && !rNuevo.advertencias.some(a => /LISTO PARA EXPORTAR|EXPORTADO/.test(a)),
  operativos(dBeto) + " / " + rNuevo.advertencias.join("; "));

// 6.4. Nunca son materias
assert("LISTO PARA EXPORTAR y EXPORTADO nunca son materias (estructura ni resultado)",
  !estReal.materias.some(m => ["LISTOPARAEXPORTAR", "EXPORTADO"].includes(m.clave))
    && dAna.informe && dAna.informe.materias.every(m => !/LISTO PARA EXPORTAR|EXPORTADO/i.test(m.nombre)));
let rOpEnMaterias = runTest("LISTO PARA EXPORTAR en la zona de materias", {
  control: ctlAna3B, sources: variante3B(conEncabezados({ "Matemática IV": "LISTO PARA EXPORTAR" })), templateGrid: tpl03
});
assert("LISTO PARA EXPORTAR antes de ORIENTACIÓN: error en vez de leerlo como materia",
  rOpEnMaterias.error && rOpEnMaterias.error.includes('El encabezado "LISTO PARA EXPORTAR" (columna AH) es un campo general u operativo y está antes de ORIENTACIÓN')
    && !rOpEnMaterias.out, rOpEnMaterias.error);

// 6.5. Las fuentes históricas no leen, validan ni aportan campos generales u operativos
const genFuentes = cargarGenerador(sources);
const histAna = genFuentes.leerFuenteHistorica_("ID_1C_REAL_xxxxxxxxxxxxxxxxxxx", "EPSI", "97000001", P1C, []);
assert("Fuente histórica: su resultado sólo contiene materias (sin estudiante, campos generales ni camposOperativos)",
  Object.keys(histAna).sort().join(",") === "materiasPorClave,nombresGrilla,ordenGrilla,ordenMaterias,periodo", Object.keys(histAna).join(","));

const srcHistIntrusa = Object.assign({}, sources, { [ID_VAR_1C]: { EPSI: variante(src1C_real.EPSI, g => {
  conCeldas("97000001", {
    "ORIENTACIÓN": "Gastronomía",
    "OBSERVACIONES GENERALES SOBRE DESEMPEÑO ACADÉMICO": "INTRUSO obs",
    "OBSERVACIONES SOBRE ASISTENCIA": "INTRUSO asistencia",
    "OBSERVACIONES SOBRE CONVIVENCIA ESCOLAR": "",
    "LISTO PARA EXPORTAR": "Tal vez",
    "EXPORTADO": ""
  })(g);
  conEncabezados({ "PLAN DE TRABAJO FORMATIVO (PTF)": "", "EXPORTADO": "LISTO PARA EXPORTAR" })(g);
}) } });
const ctlHistIntrusa = ctlReal("97000001", { B12: ID_VAR_1C });
let rHistIntrusa = runTest("Fuente histórica con cola distinta e inválida", { control: ctlHistIntrusa, sources: srcHistIntrusa, templateGrid: tpl03 });
let dHistIntrusa = generarDatos(ctlHistIntrusa, srcHistIntrusa);
assert("Fuente histórica con orientación, observaciones, PTF y operativos distintos o inválidos: el informe no cambia",
  !rHistIntrusa.error && mismoInforme(rHistIntrusa, rReal) && !JSON.stringify(rHistIntrusa.out).includes("INTRUSO"),
  rHistIntrusa.error);
assert("Fuente histórica con cola distinta e inválida: no produce advertencias",
  rHistIntrusa.advertencias.join("\n") === ADV_SIN_1B, rHistIntrusa.advertencias.join("; "));
assert("Fuente histórica con cola distinta e inválida: estudiante y camposOperativos idénticos a los de la principal",
  dHistIntrusa.informe && dAna.informe
    && JSON.stringify(dHistIntrusa.informe.estudiante) === JSON.stringify(dAna.informe.estudiante)
    && operativos(dHistIntrusa) === OP_SI_NO && dHistIntrusa.informe.estudiante.orientacion === "Informática",
  dHistIntrusa.informe && JSON.stringify(dHistIntrusa.informe.estudiante));

// 6.6. Generación de 3B: campos generales de 3B, jamás de 1C
assert("Generación 3B: observaciones y PTF salen de 3B",
  cajas(rReal.out) === "Obs 3B Ana|Asist 3B Ana|Conv 3B Ana|PTF 3B Ana" && !JSON.stringify(rReal.out).includes("1C Ana"), cajas(rReal.out));
const srcHuecos3B = variante3B(conCeldas("97000001", { "ORIENTACIÓN": "", "OBSERVACIONES SOBRE ASISTENCIA": "", "PLAN DE TRABAJO FORMATIVO (PTF)": "" }));
let rHuecos3B = runTest("Campos generales vacíos en 3B", { control: ctlAna3B, sources: srcHuecos3B, templateGrid: tpl03 });
let dHuecos3B = generarDatos(ctlAna3B, srcHuecos3B);
assert("Generación 3B: un campo general vacío en 3B queda vacío y nunca se completa con 1C",
  !rHuecos3B.error && cajas(rHuecos3B.out) === "Obs 3B Ana||Conv 3B Ana|" && !JSON.stringify(rHuecos3B.out).includes("1C Ana")
    && dHuecos3B.informe.estudiante.orientacion === "",
  rHuecos3B.error || cajas(rHuecos3B.out));

// 6.7. Generación de 1C: ahora 1C es la principal y sus campos generales y operativos sí se usan
const ctl1CReal = ctl(P1C, "97000001", { B10: "", B12: "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx", B13: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx" });
let r1CReal = runTest("Generación 1C con 1C como principal", { control: ctl1CReal, sources, templateGrid: tpl03 });
let d1CReal = generarDatos(ctl1CReal, sources);
assert("Generación 1C: observaciones y PTF salen de 1C",
  !r1CReal.error && cajas(r1CReal.out) === "Obs 1C Ana|Asist 1C Ana|Conv 1C Ana|PTF 1C Ana" && !JSON.stringify(r1CReal.out).includes("3B Ana"),
  r1CReal.error || cajas(r1CReal.out));
assert("Generación 1C: orientación y camposOperativos salen de 1C (SÍ / SÍ)",
  operativos(d1CReal) === '{"listoParaExportar":"SÍ","exportado":"SÍ"}' && d1CReal.informe.estudiante.orientacion === "Informática",
  operativos(d1CReal));
assert("Generación 1C: sólo advierte la falta de 1B",
  r1CReal.advertencias.join("\n") === "No se indicó archivo de 1º bimestre. El informe se generó sin valores anteriores.", r1CReal.advertencias.join("; "));

// 6.8. Encabezado general esperado que no puede identificarse en la principal
let rSinAsist = runTest("Encabezado de asistencia no identificable", {
  control: ctlAna3B, sources: variante3B(conEncabezados({ "OBSERVACIONES SOBRE ASISTENCIA": "ASISTENCIA" })), templateGrid: tpl03
});
assert("Encabezado general no identificable: advertencia clara",
  rSinAsist.advertencias.includes('No se encontró el encabezado "OBSERVACIONES SOBRE ASISTENCIA" después de ORIENTACIÓN en la fuente de 3º bimestre. Ese campo quedará vacío en el informe.'),
  rSinAsist.advertencias.join("; "));
assert("Encabezado general no identificable: se genera, el campo queda vacío y su columna no se lee por posición ni desde 1C",
  !rSinAsist.error && cajas(rSinAsist.out) === "Obs 3B Ana||Conv 3B Ana|PTF 3B Ana"
    && !/Asist (3B|1C) Ana/.test(JSON.stringify(rSinAsist.out))
    && JSON.stringify(filasMaterias(rSinAsist.out)) === JSON.stringify(filasMaterias(rReal.out)),
  rSinAsist.error || cajas(rSinAsist.out));

const encabezadosGenerales = [
  "OBSERVACIONES GENERALES SOBRE DESEMPEÑO ACADÉMICO", "OBSERVACIONES SOBRE ASISTENCIA",
  "OBSERVACIONES SOBRE CONVIVENCIA ESCOLAR", "PLAN DE TRABAJO FORMATIVO (PTF)"
];
const srcColaSinEnc = variante3B(conEncabezados(Object.fromEntries(encabezadosGenerales.map(e => [e, ""]))));
let rColaSinEnc = runTest("Cola con encabezados generales vacíos", { control: ctlAna3B, sources: srcColaSinEnc, templateGrid: tpl03 });
let dColaSinEnc = generarDatos(ctlAna3B, srcColaSinEnc);
assert("Encabezados generales vacíos (forma observada en EPSG): una advertencia por campo, sin adivinar columnas",
  encabezadosGenerales.every(e => rColaSinEnc.advertencias.includes('No se encontró el encabezado "' + e + '" después de ORIENTACIÓN en la fuente de 3º bimestre. Ese campo quedará vacío en el informe.'))
    && rColaSinEnc.advertencias.length === 5,
  rColaSinEnc.advertencias.join("; "));
assert("Encabezados generales vacíos: se genera con las cajas vacías, mismas materias y mismos camposOperativos",
  !rColaSinEnc.error && cajas(rColaSinEnc.out) === "|||"
    && JSON.stringify(filasMaterias(rColaSinEnc.out)) === JSON.stringify(filasMaterias(rReal.out))
    && operativos(dColaSinEnc) === OP_SI_NO,
  rColaSinEnc.error || cajas(rColaSinEnc.out));

// =============================================================================
// 7. Smoke test real: estudiante sólo en 3B con las tres fuentes cargadas (#10)
// =============================================================================
// Reproduce el flujo completo de generarInformeActual() con 1B, 1C y 3B indicadas en Control
// (Beto Nuevo, DNI 97000002, sólo existe en 3B).
const ctlSmoke = periodo => ctl(periodo, "97000002", {
  B10: "ID_1B_xxxxxxxxxxxxxxxxxxxxxxxx",
  B12: "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx",
  B13: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx"
});
const advDniHistorico = periodo => "No se encontró el DNI en la fuente de " + periodo + ". Se continuará con la fuente principal.";
let rSmoke = runTest("Smoke: estudiante sólo en 3B", { control: ctlSmoke(P3B), sources, templateGrid: tpl03 });
let matsSmoke = filasMaterias(rSmoke.out);
assert("Smoke 3B, estudiante sólo en 3B: generarInformeActual termina con éxito y lo registra en Control",
  !rSmoke.error && rSmoke.out && rSmoke.rawAlerts.length === 1 && rSmoke.rawAlerts[0].startsWith("Informe generado correctamente.")
    && rSmoke.estadoResultado === "Generado: Ficticio, Beto Nuevo",
  rSmoke.error || rSmoke.rawAlerts.join(" / "));
assert("Smoke 3B: la ausencia del DNI en 1C y en 1B se advierte sin bloquear (diálogo y Control)",
  rSmoke.advertencias.join("\n") === [advDniHistorico(P1C), advDniHistorico("1º bimestre")].join("\n")
    && rSmoke.estadoAdvertencias === rSmoke.advertencias.join("\n"),
  rSmoke.advertencias.join("; "));
assert("Smoke 3B: las columnas de 1B y 1C quedan vacías y 3B conserva sus valoraciones",
  matsSmoke.map(f => f[2] + ":" + valoresFila(f)).join(", ") === "Matemática I:||Suficiente, Lengua y Literatura I:||En proceso, FGI I:||ACREDITÓ",
  matsSmoke.map(f => f[2] + ":" + valoresFila(f)).join(", "));
assert("Smoke 3B: estudiante y campos generales salen de 3B",
  rSmoke.out && rSmoke.out[3][4] === "Ficticio, Beto Nuevo" && cajas(rSmoke.out) === "Obs 3B Beto|Asist 3B Beto|Conv 3B Beto|PTF 3B Beto",
  rSmoke.out && cajas(rSmoke.out));

// Mismo Control con el período en 1º cuatrimestre: 1C pasa a ser la principal y el DNI ausente bloquea.
let rSmoke1C = runTest("Smoke: mismo estudiante generando 1C", { control: ctlSmoke(P1C), sources, templateGrid: tpl03 });
const errDniPrincipal1C = "No se encontró el DNI en la fuente de 1º cuatrimestre. Es la fuente principal del período a generar (1º cuatrimestre), por eso no se puede continuar.";
assert("Smoke con período 1º cuatrimestre: DNI ausente en la principal bloquea e identifica la fuente principal y el período",
  rSmoke1C.error === errDniPrincipal1C && !rSmoke1C.out
    && rSmoke1C.rawAlerts[0] === "No se pudo generar el informe.\n\n" + errDniPrincipal1C
    && rSmoke1C.estadoResultado === "Error: " + errDniPrincipal1C,
  rSmoke1C.error);

console.log(`\nResultado final: ${fail === 0 ? "TODOS LOS TESTS PASARON" : `${fail} TESTS FALLARON`}`);
process.exit(fail ? 1 : 0);
