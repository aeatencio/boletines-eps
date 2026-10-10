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

function runTest(testName, { control, sources, templateGrid }) {
  const code = fs.readFileSync(path.join(ROOT, "src/Codigo_Apps_Script_Generador_EPS_v0_3.gs"), "utf8");
  const plantilla = new MockSheet("Plantilla", templateGrid);
  const ctrlGrid = [];
  for (const [a1, v] of Object.entries(control)) {
    const [r, c] = a1ToRC(a1);
    while (ctrlGrid.length < r) ctrlGrid.push([]);
    ctrlGrid[r - 1][c - 1] = v;
  }
  const ctrl = new MockSheet("Control", ctrlGrid);
  const ss = new MockSS([plantilla, ctrl]);
  const rawAlerts = [];
  const ctx = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ({ alert: m => rawAlerts.push(m), createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }),
      openById: id => {
        if (!sources[id]) throw new Error("mock: id inexistente " + id);
        return { getSheetByName: n => sources[id][n] ? new MockSheet(id + "/" + n, sources[id][n]) : null };
      }
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

  return {
    error,
    rawAlerts,
    advertencias,
    estadoAdvertencias,
    out: out ? rect(out.grid) : null
  };
}

// Carga el generador sin ejecutarlo, para probar funciones puras del parser.
function cargarGenerador() {
  const ctx = {};
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(ROOT, "src/Codigo_Apps_Script_Generador_EPS_v0_3.gs"), "utf8"), ctx);
  return ctx;
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
assert("Estructura real: la cola empieza en ORIENTACIÓN (columna AN)",
  estReal.inicioCola === 39 && estReal.especiales.orientacion === 39, estReal.inicioCola);
assert("Estructura real: campos generales detectados dentro de la cola",
  [estReal.especiales.observacionesGenerales, estReal.especiales.observacionesAsistencia,
   estReal.especiales.observacionesConvivencia, estReal.especiales.ptf].join(",") === "40,41,42,43",
  JSON.stringify(estReal.especiales));
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

console.log(`\nResultado final: ${fail === 0 ? "TODOS LOS TESTS PASARON" : `${fail} TESTS FALLARON`}`);
process.exit(fail ? 1 : 0);
