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
  // Como Google Sheets: el nombre no puede repetirse en el libro, sin distinguir mayúsculas.
  setName(n) {
    if (this.ss && this.ss.sheets.some(s => s !== this && s.name.toLowerCase() === String(n).toLowerCase())) {
      throw new Error("mock: ya existe una hoja llamada " + n);
    }
    this.name = n;
  }
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
      setValues: filas => {
        if (filas.length !== h || filas.some(f => f.length !== w)) throw new Error("mock: setValues con dimensiones distintas del rango");
        filas.forEach((f, i) => f.forEach((v, j) => sh.set(r + i, c + j, v instanceof Date ? "<fecha>" : v)));
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
  // Como Google Sheets: no se puede ocultar la última hoja visible.
  hideSheet() {
    if (this.ss && this.ss.sheets.every(s => s === this || s.hidden)) throw new Error("mock: no se pueden ocultar todas las hojas");
    this.hidden = true;
  }
  isSheetHidden() { return !!this.hidden; }
  copyTo(ss) {
    const s = new MockSheet("Copia de " + this.name, this.grid);
    s.ss = ss;
    ss.sheets.push(s);
    return s;
  }
}

class MockSS {
  constructor(sheets, name) {
    this.sheets = sheets;
    this.name = name || "Generador";
    sheets.forEach(s => { s.ss = this; });
  }
  getName() { return this.name; }
  getUrl() { return "https://docs.google.com/spreadsheets/d/MOCK/" + encodeURIComponent(this.name); }
  getSpreadsheetTimeZone() { return "America/Argentina/Buenos_Aires"; }
  getSheets() { return this.sheets.slice(); }
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
// El registro anota cada apertura y lectura, y cualquier acceso a algo distinto de la lectura permitida.
function abrirFuenteMock(sources, registro = { aperturas: [], lecturas: [], indebidos: [] }) {
  const soloLectura = (obj, nombre) => new Proxy(obj, {
    get(t, p) {
      if (!(p in t)) registro.indebidos.push(nombre + "." + String(p));
      return t[p];
    }
  });
  return id => {
    if (!sources[id]) throw new Error("mock: id inexistente " + id);
    registro.aperturas.push(id);
    return soloLectura({
      getSheetByName: n => {
        if (!sources[id][n]) return null;
        const hoja = new MockSheet(id + "/" + n, sources[id][n]);
        return soloLectura({
          getName: () => hoja.getName(),
          getDataRange: () => {
            registro.lecturas.push(id + "/" + n);
            return soloLectura(hoja.getDataRange(), id + "/" + n + " (rango)");
          }
        }, id + "/" + n);
      }
    }, id);
  };
}

function runTest(testName, { control, sources, templateGrid, ejecutar = "generarInformeActual()" }) {
  const code = fs.readFileSync(path.join(ROOT, "src/Codigo_Apps_Script_Generador_EPS_v0_3.gs"), "utf8");
  const plantilla = new MockSheet("Plantilla", templateGrid);
  const ctrl = hojaControl(control);
  const ss = new MockSS([plantilla, ctrl]);
  const rawAlerts = [];
  const creados = [];
  const registro = { aperturas: [], lecturas: [], indebidos: [] };
  const ctx = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ({ alert: m => rawAlerts.push(m), createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }),
      openById: abrirFuenteMock(sources, registro),
      create: nombre => {
        const libro = new MockSS([new MockSheet("Hoja 1", [])], nombre);
        creados.push(libro);
        return libro;
      }
    },
    Utilities: { formatDate: () => "<fecha>" }
  };
  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  let error = null;
  try {
    vm.runInContext(ejecutar, ctx);
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
    out: out ? rect(out.grid) : null,
    ss,
    creados,
    registro
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

// Coordenadas de Control tomadas del generador: los tests no repiten dónde está cada fuente.
const CELDAS_CONTROL = vm.runInContext("CONFIG_EPS", cargarGenerador());
const F1B = CELDAS_CONTROL.CELDA_FUENTE_1B;
const F1C = CELDAS_CONTROL.CELDA_FUENTE_1C;
const F3B = CELDAS_CONTROL.CELDA_FUENTE_3B;

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
  [F1B]: "ID_1B_xxxxxxxxxxxxxxxxxxxxxxxx",
  [F1C]: "ID_1C_xxxxxxxxxxxxxxxxxxxxxxxx",
  [F3B]: "ID_3B_xxxxxxxxxxxxxxxxxxxxxxxx"
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
let ctlInv = ctl("3º bimestre", "99000001", { [F3B]: "ID_3B_INV_xxxxxxxxxxxxxxxxxxxx" });
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
let r1B = runTest("Regresión 1B (CP-01)", { control: ctl("1º bimestre", "99000001", { [F1C]: "", [F3B]: "" }), sources, templateGrid: tpl03 });
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
let r1C = runTest("Regresión 1C acumulativo (CP-02)", { control: ctl("1º cuatrimestre", "99000001", { [F3B]: "" }), sources, templateGrid: tpl03 });
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
let r1C_inv = runTest("Regresión 1C valor no válido (CP-03)", { control: ctl("1º cuatrimestre", "99000002", { [F3B]: "" }), sources, templateGrid: tpl03 });
assert("Regresión 1C valor no válido: genera sin abortar", !r1C_inv.error && r1C_inv.out, r1C_inv.error);
let f1cBrunoLen = (r1C_inv.out || []).find(f => f[2] === "Lengua y Literatura I") || [];
assert("Regresión 1C valor no válido: descarta '7,5' y conserva 1B (Suficiente||)", f1cBrunoLen.slice(4, 7).join("|") === "Suficiente||", f1cBrunoLen.slice(4, 7));
assert("Regresión 1C valor no válido: emite advertencia por 7,5", r1C_inv.advertencias.some(a => a.includes("7,5")), r1C_inv.advertencias.join("; "));
let f1cBrunoMat = (r1C_inv.out || []).find(f => f[2] === "Matemática I") || [];
assert("Regresión 1C valor no válido: continuidad (Matemática I tiene |8|)", f1cBrunoMat.slice(4, 7).join("|") === "|8|", f1cBrunoMat.slice(4, 7));


// 4.4. Regresión 1C Orientación Gastronomía (CP-06 de v0.2: Nora Uno, DNI 98000001, EPSG)
let rGastro = runTest("Regresión 1C Gastronomía (CP-06)", {
  control: ctl("1º cuatrimestre", "98000001", { B6: "EPSG", [F3B]: "" }),
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
  [F1B]: "",
  [F1C]: "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx",
  [F3B]: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx"
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
  control: ctlReal("96000001", { B6: "EPSG", [F1C]: "" }), sources, templateGrid: tpl03
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
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "ORIENTACIÓN")] = ""; }) }),
  templateGrid: tpl03
});
assert("Sin ORIENTACIÓN en la principal: se informa error y no se genera",
  rSinOrient.error && rSinOrient.error.includes("No se encontró la columna ORIENTACIÓN") && !rSinOrient.out, rSinOrient.error);

let rOrientDoble = runTest("ORIENTACIÓN repetida", {
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "SEGUIMIENTO INTERNO")] = "Orientación"; }) }),
  templateGrid: tpl03
});
assert("ORIENTACIÓN repetida en la principal: se informa error",
  rOrientDoble.error && rOrientDoble.error.includes("La columna ORIENTACIÓN aparece más de una vez (columnas AN, AU)"), rOrientDoble.error);

let rRepetida = runTest("Materia repetida en principal", {
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "Matemática IV")] = "Matemática I"; }) }),
  templateGrid: tpl03
});
assert("Materia repetida en la principal: se informa error con ambas columnas, sin elegir una",
  rRepetida.error && rRepetida.error.includes('Materia repetida en la fuente de 3º bimestre: "Matemática I" (columnas E, AH)') && !rRepetida.out,
  rRepetida.error);

let rRepetidaHist = runTest("Materia repetida en fuente anterior", {
  control: ctlReal("97000001", { [F1C]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src1C_real.EPSI, g => { g[0][columnaDe(g, "Matemática IV")] = "Matemática I"; }) }),
  templateGrid: tpl03
});
assert("Materia repetida en una fuente anterior: advertencia, se descarta esa fuente y se genera con la principal",
  !rRepetidaHist.error && rRepetidaHist.advertencias.some(a => a.startsWith("Materia repetida en la fuente de 1º cuatrimestre") && a.endsWith("Se continuará con la fuente principal."))
    && valoresFila(filasMaterias(rRepetidaHist.out).find(f => f[2] === "Matemática I") || []) === "||Avanzado",
  rRepetidaHist.error || rRepetidaHist.advertencias.join("; "));

let rParRoto = runTest("Materia sin columna de valoración", {
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "Matemática I") + 1] = "Columna intrusa"; }) }),
  templateGrid: tpl03
});
assert("Materia sin su columna de valoración: se informa error",
  rParRoto.error && rParRoto.error.includes('La materia "Matemática I" (columna E) no tiene a continuación su columna de valoración'), rParRoto.error);

let rFilaDni = runTest("DNI fuera de la fila Valoración", {
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { const r = filaDni(g, "97000001"); g[r][2] = ""; g[r + 1][2] = "97000001"; }) }),
  templateGrid: tpl03
});
assert("DNI fuera de la fila Valoración del bloque: se informa error en vez de leer otra fila",
  rFilaDni.error && rFilaDni.error.includes('La fila del DNI no es la fila "Valoración"') && !rFilaDni.out, rFilaDni.error);

// 5.7. Nombre mostrado: el de la grilla principal si la materia existe en ella
let rNombre = runTest("Nombre de la grilla principal", {
  control: ctlReal("97000001", { [F1C]: ID_VAR }),
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
const ctlAna3B = ctlReal("97000001", { [F3B]: ID_VAR });
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
const histAna = genFuentes.leerFuenteHistorica_(genFuentes.abrirFuenteAnterior_("ID_1C_REAL_xxxxxxxxxxxxxxxxxxx", "EPSI", P1C, []), "97000001", []);
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
const ctlHistIntrusa = ctlReal("97000001", { [F1C]: ID_VAR_1C });
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
const ctl1CReal = ctl(P1C, "97000001", { [F1B]: "", [F1C]: "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx", [F3B]: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx" });
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
  [F1B]: "ID_1B_xxxxxxxxxxxxxxxxxxxxxxxx",
  [F1C]: "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx",
  [F3B]: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx"
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

// =============================================================================
// 8. Generación por curso (#12)
// =============================================================================
// Copia el bloque de cuatro filas de un estudiante con otra identidad y otros estados operativos,
// a continuación del bloque de otro estudiante: queda en su mismo nivel. Cada copia sigue a la anterior.
const conBloques = (modelo, despuesDe, nuevos) => g => {
  let anterior = despuesDe;
  nuevos.forEach(n => {
    const bloque = g.slice(filaDni(g, modelo), filaDni(g, modelo) + 4).map(f => f.slice());
    [bloque[0][0], bloque[0][1], bloque[0][2]] = [n.apellido, n.nombres, n.dni];
    if ("listo" in n) bloque[0][colDe(g, "LISTO PARA EXPORTAR")] = n.listo;
    if ("exportado" in n) bloque[0][colDe(g, "EXPORTADO")] = n.exportado;
    g.splice(filaDni(g, anterior) + 4, 0, ...bloque);
    anterior = n.dni;
  });
};
const aplicar = (...cambios) => g => cambios.forEach(c => c(g));
const ID_CURSO_3B = "ID_CURSO_3B_xxxxxxxxxxxxxxxxx";
const ID_CURSO_1C = "ID_CURSO_1C_xxxxxxxxxxxxxxxxx";
const conCurso = (cambiar3B, cambiar1C) => Object.assign({}, sources, {
  [ID_CURSO_3B]: { EPSI: variante(src3B_real.EPSI, cambiar3B || (() => {})) },
  [ID_CURSO_1C]: { EPSI: variante(src1C_real.EPSI, cambiar1C || (() => {})) }
});
const fuentesCurso = { [F1B]: "", [F1C]: ID_CURSO_1C, [F3B]: ID_CURSO_3B };
const ctlCurso = (nivel, extra) => ctl(P3B, "", Object.assign({ B9: nivel }, fuentesCurso, extra));
const ctlIndividual = dni => ctl(P3B, dni, fuentesCurso);
const runCurso = (nombre, control, srcs) => runTest(nombre, { control, sources: srcs, templateGrid: tpl03, ejecutar: "generarBoletinesPorCurso()" });
const libroDe = r => r.creados[0] || null;
const hojasDe = r => libroDe(r) ? libroDe(r).getSheets().map(s => s.getName()) : [];
const gridHoja = (r, nombre) => { const h = libroDe(r) && libroDe(r).getSheetByName(nombre); return h ? rect(h.grid) : null; };
const resumenDe = r => {
  const g = gridHoja(r, "Resumen") || [];
  const i = g.findIndex(f => f[0] === "Estado");
  return {
    generales: (g.find(f => f[0] === "Advertencias generales") || [])[1],
    filas: i < 0 ? [] : g.slice(i + 1).filter(f => f[0]).map(f => ({
      estado: f[0], hoja: f[1], nombre: f[2], dni: f[3], fila: f[4], exportado: f[5], detalle: f[6]
    }))
  };
};
const enResumen = (r, dni) => resumenDe(r).filas.find(f => f.dni === dni) || {};
const advertenciasTanda = (r, dni) => [resumenDe(r).generales].concat(enResumen(r, dni).detalle.split("\n"))
  .filter(a => a && a !== "Sin advertencias");

// 8.1. Curso EPS2 con estados variados en 3B (principal); en 1C (anterior) los estados operativos son otros.
const srcCurso = conCurso(
  aplicar(
    conCeldas("97000002", { "LISTO PARA EXPORTAR": "SÍ" }), // Beto, EPS1: listo, pero de otro nivel
    conBloques("97000001", "97000001", [
      { apellido: "Ficticio", nombres: "Caro No", dni: "97000012", listo: "NO", exportado: "SÍ" },
      { apellido: "Ficticia", nombres: "Ana Real", dni: "97000011", listo: "SÍ", exportado: "NO" },
      { apellido: "Ficticio", nombres: "Dani Vacio", dni: "97000013", listo: "" },
      { apellido: "Ficticio", nombres: "Eli Invalido", dni: "97000014", listo: "Tal vez" },
      { apellido: "Ficticio", nombres: "Fede Exportado", dni: "97000015", listo: "SÍ", exportado: "SÍ" },
      { apellido: "Ficticio", nombres: "Gabi Si", dni: "97000016", listo: "Si" }
    ])
  ),
  aplicar(
    conCeldas("97000001", { "LISTO PARA EXPORTAR": "Tal vez", "EXPORTADO": "" }),
    conBloques("97000001", "97000001", [
      { apellido: "Ficticio", nombres: "Caro No", dni: "97000012", listo: "SÍ", exportado: "NO" },
      { apellido: "Ficticio", nombres: "Solo Anterior", dni: "97000018", listo: "SÍ", exportado: "NO" }
    ])
  )
);
const fotoFuentes = JSON.stringify(srcCurso);
let rCurso = runCurso("Curso EPS2", ctlCurso("EPS2"), srcCurso);
const resCurso = resumenDe(rCurso);
const estadosCurso = resCurso.filas.map(f => f.estado + ":" + f.dni).join(", ");
const HOJAS_EPS2 = ["Resumen", "Ficticia, Ana Real (97000001)", "Ficticia, Ana Real (97000011)", "Ficticio, Fede Exportado", "Ficticio, Gabi Si"];
const NOMBRE_LIBRO_EPS2 = "Boletines EPS - EPSI - EPS2 - 3º bimestre 2026 - <fecha>";

assert("Por curso: termina sin error y crea un único libro nuevo con el nombre del curso",
  !rCurso.error && rCurso.creados.length === 1 && libroDe(rCurso).getName() === NOMBRE_LIBRO_EPS2,
  rCurso.error || rCurso.creados.map(l => l.getName()).join(", "));
assert("Por curso: una hoja por estudiante seleccionado, después de Resumen y en el orden de la grilla",
  JSON.stringify(hojasDe(rCurso)) === JSON.stringify(HOJAS_EPS2), hojasDe(rCurso).join(" | "));
assert("Por curso: el Resumen lista sólo estudiantes de EPS2 de la principal, en orden de grilla, con su estado",
  estadosCurso === "Generado:97000001, Omitido:97000012, Generado:97000011, Omitido:97000013, Omitido:97000014, Generado:97000015, Generado:97000016",
  estadosCurso);
assert("Por curso: no se seleccionan estudiantes de otro nivel aunque estén listos (Beto, EPS1)",
  !resCurso.filas.some(f => f.dni === "97000002") && !hojasDe(rCurso).some(h => /Beto/.test(h)));
assert("Por curso: un estudiante que sólo existe en una fuente anterior, aunque esté listo allí, no aparece",
  !resCurso.filas.some(f => f.dni === "97000018") && !hojasDe(rCurso).some(h => /Solo Anterior/.test(h)));
assert("Por curso: LISTO PARA EXPORTAR = NO se omite aunque la fuente anterior diga SÍ y EXPORTADO = SÍ",
  enResumen(rCurso, "97000012").estado === "Omitido" && enResumen(rCurso, "97000012").detalle === "LISTO PARA EXPORTAR = NO."
    && enResumen(rCurso, "97000012").exportado === "SÍ",
  JSON.stringify(enResumen(rCurso, "97000012")));
assert("Por curso: LISTO PARA EXPORTAR vacío o inválido se omite con el motivo",
  enResumen(rCurso, "97000013").estado === "Omitido" && enResumen(rCurso, "97000013").detalle === "LISTO PARA EXPORTAR está vacío en la fuente de 3º bimestre. Se esperaba SÍ o NO."
    && enResumen(rCurso, "97000014").estado === "Omitido" && enResumen(rCurso, "97000014").detalle.includes('Valor: "Tal vez"'),
  JSON.stringify([enResumen(rCurso, "97000013"), enResumen(rCurso, "97000014")]));
assert("Por curso: EXPORTADO no filtra (SÍ / SÍ se genera) y se informa en el Resumen",
  enResumen(rCurso, "97000015").estado === "Generado" && enResumen(rCurso, "97000015").exportado === "SÍ"
    && enResumen(rCurso, "97000001").exportado === "NO",
  JSON.stringify(enResumen(rCurso, "97000015")));
assert("Por curso: 'Si' se interpreta como SÍ (contrato de #10) y la forma no canónica queda advertida",
  enResumen(rCurso, "97000016").estado === "Generado"
    && enResumen(rCurso, "97000016").detalle.includes('LISTO PARA EXPORTAR tiene "Si" en la fuente de 3º bimestre. Se interpretó como SÍ'),
  JSON.stringify(enResumen(rCurso, "97000016")));
assert("Por curso: estados operativos de la fuente anterior no se consultan ni se advierten",
  ![resCurso.generales].concat(resCurso.filas.map(f => f.detalle)).join("\n").split("\n")
    .some(a => /LISTO PARA EXPORTAR|EXPORTADO/.test(a) && a.includes("1º cuatrimestre")),
  JSON.stringify(resCurso));
assert("Por curso: nombres de hoja en conflicto se distinguen con el DNI, de forma determinista",
  enResumen(rCurso, "97000001").hoja === "Ficticia, Ana Real (97000001)" && enResumen(rCurso, "97000011").hoja === "Ficticia, Ana Real (97000011)",
  JSON.stringify([enResumen(rCurso, "97000001"), enResumen(rCurso, "97000011")]));
assert("Por curso: el Resumen informa la fila de la fuente y las advertencias generales una sola vez",
  enResumen(rCurso, "97000001").fila === "9" && resCurso.generales === ADV_SIN_1B
    && !resCurso.filas.some(f => f.detalle.includes(ADV_SIN_1B)),
  JSON.stringify(enResumen(rCurso, "97000001")) + " / " + resCurso.generales);

// Equivalencia con la generación individual (mismas fuentes y mismo Control, con el DNI del estudiante)
["97000001", "97000015", "97000016"].forEach(dni => {
  const ind = runTest("Individual " + dni, { control: ctlIndividual(dni), sources: srcCurso, templateGrid: tpl03 });
  const hoja = gridHoja(rCurso, enResumen(rCurso, dni).hoja);
  assert("Por curso: la hoja de " + dni + " es idéntica al Informe_actual de la generación individual",
    !ind.error && hoja && JSON.stringify(hoja) === JSON.stringify(ind.out), ind.error || "hojas distintas");
  assert("Por curso: las advertencias de " + dni + " (generales + propias) son las de la generación individual",
    JSON.stringify(advertenciasTanda(rCurso, dni).sort()) === JSON.stringify(ind.advertencias.slice().sort()),
    advertenciasTanda(rCurso, dni).join("; ") + " ≠ " + ind.advertencias.join("; "));
});
assert("Por curso: la hoja conserva estudiante, nivel visual, materias acumuladas y observaciones",
  (g => g && g[3][4] === "Ficticia, Ana Real" && g[5][4] === "Informática 2" && cajas(g) === "Obs 3B Ana|Asist 3B Ana|Conv 3B Ana|PTF 3B Ana"
    && valoresFila(filasMaterias(g).find(f => f[2] === "Matemática I") || []) === "|7|Avanzado")(gridHoja(rCurso, "Ficticia, Ana Real (97000001)")));

// Lectura única y fuentes de sólo lectura
assert("Por curso: cada fuente se abre y se lee una sola vez para toda la tanda",
  JSON.stringify(rCurso.registro.aperturas.slice().sort()) === JSON.stringify([ID_CURSO_1C, ID_CURSO_3B])
    && rCurso.registro.lecturas.length === 2,
  JSON.stringify(rCurso.registro));
assert("Por curso: ningún acceso a las fuentes fuera de la lectura de valores, y las fuentes quedan intactas",
  rCurso.registro.indebidos.length === 0 && JSON.stringify(srcCurso) === fotoFuentes, rCurso.registro.indebidos.join(", "));
assert("Por curso: el generador no crea Informe_actual ni modifica Plantilla",
  rCurso.ss.getSheets().map(s => s.getName()).join("|") === "Plantilla|Control"
    && JSON.stringify(rect(rCurso.ss.getSheetByName("Plantilla").grid)) === JSON.stringify(rect(tpl03)));
assert("Por curso: Control registra el libro, los conteos y las advertencias generales",
  rCurso.estadoResultado.startsWith("Libro creado: " + NOMBRE_LIBRO_EPS2 + " · Generados: 4 · Omitidos: 3 · Con error: 0.")
    && rCurso.estadoAdvertencias === ADV_SIN_1B && rCurso.rawAlerts.length === 1 && rCurso.rawAlerts[0].startsWith("Libro creado: "),
  rCurso.estadoResultado + " / " + rCurso.estadoAdvertencias);

// 8.2. Otro nivel del mismo Control (nivel indicado como "1")
let rCurso1 = runCurso("Curso EPS1", ctlCurso("1"), srcCurso);
let indBeto = runTest("Individual Beto", { control: ctlIndividual("97000002"), sources: srcCurso, templateGrid: tpl03 });
assert("Por curso EPS1 (nivel '1'): sólo Beto, idéntico a su generación individual",
  !rCurso1.error && hojasDe(rCurso1).join("|") === "Resumen|Ficticio, Beto Nuevo"
    && libroDe(rCurso1).getName() === "Boletines EPS - EPSI - EPS1 - 3º bimestre 2026 - <fecha>"
    && JSON.stringify(gridHoja(rCurso1, "Ficticio, Beto Nuevo")) === JSON.stringify(indBeto.out),
  rCurso1.error || hojasDe(rCurso1).join("|"));

// 8.3. Falla de un estudiante: DNI repetido en la principal (también en EPS1); el resto se genera
const srcDuplicado = conCurso(aplicar(
  conBloques("97000001", "97000001", [{ apellido: "Ficticio", nombres: "Ivan Duplicado", dni: "97000017", listo: "SÍ" }]),
  conBloques("97000002", "97000002", [{ apellido: "Ficticio", nombres: "Ivan Duplicado", dni: "97000017", listo: "NO" }])
));
let rDuplicado = runCurso("Curso con un estudiante fallido", ctlCurso("EPS2"), srcDuplicado);
assert("Falla individual: el estudiante con DNI repetido queda con error y su motivo, sin hoja",
  !rDuplicado.error && enResumen(rDuplicado, "97000017").estado === "Error"
    && enResumen(rDuplicado, "97000017").detalle.startsWith("El DNI aparece más de una vez en la fuente de 3º bimestre.")
    && !hojasDe(rDuplicado).some(h => /Ivan/.test(h)),
  rDuplicado.error || JSON.stringify(enResumen(rDuplicado, "97000017")));
assert("Falla individual: el resto de la tanda se genera y Control cuenta el error",
  hojasDe(rDuplicado).join("|") === "Resumen|Ficticia, Ana Real"
    && rDuplicado.estadoResultado.includes("Generados: 1 · Omitidos: 0 · Con error: 1."),
  hojasDe(rDuplicado).join("|") + " / " + rDuplicado.estadoResultado);

// 8.4. Errores globales: sin LISTO PARA EXPORTAR identificable en la principal, o nivel inexistente
const errSinListo = 'No se encontró una única columna "LISTO PARA EXPORTAR" después de ORIENTACIÓN en la hoja EPSI de la fuente de 3º bimestre.';
let rSinListo = runCurso("Sin LISTO PARA EXPORTAR", ctlCurso("EPS2"), conCurso(conEncabezados({ "LISTO PARA EXPORTAR": "LISTO PARA IMPRIMIR" })));
assert("Sin columna LISTO PARA EXPORTAR (ni se acepta LISTO PARA IMPRIMIR): error claro, sin libro",
  rSinListo.error && rSinListo.error.startsWith(errSinListo) && rSinListo.creados.length === 0
    && rSinListo.estadoResultado === "Error: " + rSinListo.error && rSinListo.rawAlerts[0].startsWith("No se pudo generar el libro del curso."),
  rSinListo.error);
let rListoDoble = runCurso("LISTO PARA EXPORTAR repetido", ctlCurso("EPS2"), conCurso(conEncabezados({ "SEGUIMIENTO INTERNO": "LISTO PARA EXPORTAR" })));
assert("LISTO PARA EXPORTAR repetido: error claro, sin elegir columna ni crear libro",
  rListoDoble.error && rListoDoble.error.startsWith(errSinListo) && rListoDoble.creados.length === 0, rListoDoble.error);
let rSinListoV03 = runCurso("Grilla v0.3 sin columnas operativas", ctl(P3B, "", { B9: "EPS1", [F1C]: "", [F3B]: "ID_3B_xxxxxxxxxxxxxxxxxxxxxxxx" }), sources);
assert("Grilla sin columnas operativas (v0.3): la generación por curso no adivina y falla",
  rSinListoV03.error && rSinListoV03.error.startsWith(errSinListo) && rSinListoV03.creados.length === 0, rSinListoV03.error);
let rNivelInexistente = runCurso("Nivel inexistente", ctlCurso("EPS5"), srcCurso);
assert("Nivel sin marcador en la grilla: error claro, sin libro",
  rNivelInexistente.error === "No se encontró el nivel EPS5 (marcador en la columna A) en la hoja EPSI de la fuente de 3º bimestre."
    && rNivelInexistente.creados.length === 0, rNivelInexistente.error);
let rSinNivel = runCurso("Sin nivel", ctlCurso(""), srcCurso);
assert("Sin nivel en Control: error claro antes de leer fuentes",
  rSinNivel.error && rSinNivel.error.startsWith("Falta indicar el nivel / curso en la celda B9") && rSinNivel.registro.aperturas.length === 0,
  rSinNivel.error);

// 8.5. Tanda sin estudiantes listos: en 3B Ana está en NO (en 1C, fuente anterior, está en SÍ)
let rNadieListo = runCurso("Sin estudiantes listos", ctlCurso("EPS2"),
  conCurso(conCeldas("97000001", { "LISTO PARA EXPORTAR": "NO" }), conCeldas("97000001", { "LISTO PARA EXPORTAR": "SÍ" })));
assert("Sin estudiantes listos: no se crea ningún libro y el resultado lo explica",
  !rNadieListo.error && rNadieListo.creados.length === 0
    && rNadieListo.rawAlerts[0].startsWith("No se creó ningún libro.\n\nNo hay estudiantes con LISTO PARA EXPORTAR = SÍ en EPSI EPS2 (3º bimestre).")
    && rNadieListo.estadoResultado === "Sin libro: No hay estudiantes con LISTO PARA EXPORTAR = SÍ en EPSI EPS2 (3º bimestre). Generados: 0 · Omitidos: 1 · Con error: 0.",
  rNadieListo.error || rNadieListo.rawAlerts.join(" / "));
assert("Sin estudiantes listos: Control conserva el detalle de los omitidos y las advertencias generales",
  rNadieListo.estadoAdvertencias === ADV_SIN_1B + "\nOmitido · fila 9 · Ficticia, Ana Real: LISTO PARA EXPORTAR = NO.",
  rNadieListo.estadoAdvertencias);

// 8.6. Menú: la generación por curso es una acción separada y la individual se conserva
const menu = [];
const genMenu = cargarGenerador();
genMenu.SpreadsheetApp = { getUi: () => ({ createMenu: () => ({ addItem(t, f) { menu.push(t + "→" + f); return this; }, addToUi() {} }) }) };
genMenu.onOpen();
assert("Menú: conserva 'Generar informe actual' y agrega 'Generar boletines por curso'",
  menu.join(", ") === "Generar informe actual→generarInformeActual, Generar boletines por curso→generarBoletinesPorCurso", menu.join(", "));

// 8.7. Resumen oculto: existe y conserva el detalle, pero no se imprime ni se exporta con el libro
const visibilidad = r => libroDe(r) ? libroDe(r).getSheets().map(s => s.getName() + (s.isSheetHidden() ? ":oculta" : ":visible")).join(" | ") : "";
assert("Resumen oculto: queda oculta y todos los boletines quedan visibles",
  visibilidad(rCurso) === HOJAS_EPS2.map((h, i) => h + (i === 0 ? ":oculta" : ":visible")).join(" | ") && resCurso.filas.length === 7,
  visibilidad(rCurso));
assert("Resumen oculto: también cuando hay errores, sin perder el detalle del estudiante fallido",
  visibilidad(rDuplicado) === "Resumen:oculta | Ficticia, Ana Real:visible" && enResumen(rDuplicado, "97000017").estado === "Error",
  visibilidad(rDuplicado));
let rSinHojas = runTest("Ninguna hoja pudo crearse", {
  control: ctlCurso("EPS2"), sources: srcCurso, templateGrid: tpl03,
  ejecutar: "SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla').copyTo = () => { throw new Error('copia rechazada'); }; generarBoletinesPorCurso()"
});
assert("Resumen oculto: si no se generó ningún boletín, Resumen queda visible (única hoja) con los errores",
  !rSinHojas.error && visibilidad(rSinHojas) === "Resumen:visible"
    && resumenDe(rSinHojas).filas.filter(f => f.estado === "Error").every(f => f.detalle.startsWith("No se pudo crear la hoja: copia rechazada"))
    && resumenDe(rSinHojas).filas.filter(f => f.estado === "Error").length === 4,
  rSinHojas.error || visibilidad(rSinHojas));

// =============================================================================
// 9. Control del workbook versionado: selector de nivel para la generación por curso (#12)
// =============================================================================
const XLSX_GENERADOR = "src/Generador_informes_individuales_EPS_v0_3.xlsx";
const controlXlsx = readXlsx(XLSX_GENERADOR).Control;
// XML de una hoja del último xlsx extraído, ubicada por nombre como en readXlsx.
const xmlHoja = (rel, nombre) => {
  const dir = path.join(WORK, path.basename(rel, ".xlsx"));
  const rd = p => fs.readFileSync(path.join(dir, p), "utf8");
  const rid = rd("xl/workbook.xml").match(new RegExp("<(?:\\w+:)?sheet [^>]*name=\"" + nombre + "\"[^>]*r:id=\"([^\"]+)\""))[1];
  const rels = rd("xl/_rels/workbook.xml.rels");
  const t = rels.match(new RegExp("Target=\"([^\"]+)\"[^>]*Id=\"" + rid + "\"")) || rels.match(new RegExp("Id=\"" + rid + "\"[^>]*Target=\"([^\"]+)\""));
  let p = t[1].replace(/^\//, "");
  if (!p.startsWith("xl/")) p = "xl/" + p;
  return rd(p);
};
const validacionB9 = (xmlHoja(XLSX_GENERADOR, "Control").match(/<(?:\w+:)?dataValidation [^>]*sqref="B9"[^>]*>([\s\S]*?)<\/(?:\w+:)?dataValidation>/) || [])[0] || "";
assert("Control (xlsx): A9 rotula el nivel / curso y B9 queda vacío para elegir",
  (controlXlsx[8] || [])[0] === "Nivel / curso" && !(controlXlsx[8] || [])[1], JSON.stringify(controlXlsx[8]));
assert("Control (xlsx): B9 tiene un selector de lista con EPS1, EPS2, EPS3 y EPS4",
  /type="list"/.test(validacionB9) && /<(?:\w+:)?formula1>"EPS1,EPS2,EPS3,EPS4"<\/(?:\w+:)?formula1>/.test(validacionB9), validacionB9);

// Control real del workbook (con sus rótulos y valores por defecto), completado como lo haría el usuario
const letras = c => { let s = ""; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
const ctlDesdeXlsx = extra => {
  const celdas = {};
  controlXlsx.forEach((fila, r) => (fila || []).forEach((v, c) => { if (v !== undefined) celdas[letras(c) + (r + 1)] = v; }));
  return Object.assign(celdas, extra);
};
const datosCursoXlsx = Object.assign({ B5: P3B, B6: "EPSI" }, fuentesCurso);
let rXlsxEps1 = runCurso("Curso con Control del xlsx (EPS1)", ctlDesdeXlsx(Object.assign({ B9: "EPS1" }, datosCursoXlsx)), srcCurso);
let rXlsxEps2 = runCurso("Curso con Control del xlsx (EPS2)", ctlDesdeXlsx(Object.assign({ B9: "EPS2" }, datosCursoXlsx)), srcCurso);
assert("Control (xlsx): la generación por curso usa el nivel elegido en B9",
  !rXlsxEps1.error && !rXlsxEps2.error
    && hojasDe(rXlsxEps1).join("|") === "Resumen|Ficticio, Beto Nuevo" && libroDe(rXlsxEps1).getName().includes(" - EPS1 - ")
    && JSON.stringify(hojasDe(rXlsxEps2)) === JSON.stringify(HOJAS_EPS2) && libroDe(rXlsxEps2).getName().includes(" - EPS2 - "),
  (rXlsxEps1.error || hojasDe(rXlsxEps1).join("|")) + " / " + (rXlsxEps2.error || hojasDe(rXlsxEps2).join("|")));
let rXlsxSinNivel = runCurso("Curso con Control del xlsx sin nivel", ctlDesdeXlsx(datosCursoXlsx), srcCurso);
assert("Control (xlsx): con B9 sin elegir, la generación por curso lo pide y no crea libro",
  rXlsxSinNivel.error && rXlsxSinNivel.error.startsWith("Falta indicar el nivel / curso en la celda B9") && rXlsxSinNivel.creados.length === 0,
  rXlsxSinNivel.error);
let rXlsxIndividual = runTest("Individual con Control del xlsx y B9 de otro nivel", {
  control: ctlDesdeXlsx({ B5: P3B, B6: "EPSI", B7: "97000001", B9: "EPS1", [F1B]: "", [F1C]: "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx", [F3B]: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx" }),
  sources, templateGrid: tpl03
});
assert("Control (xlsx): la generación individual ignora B9 (Ana, EPS2, con B9 = EPS1: mismo informe y sin libro nuevo)",
  !rXlsxIndividual.error && mismoInforme(rXlsxIndividual, rReal) && rXlsxIndividual.creados.length === 0
    && rXlsxIndividual.advertencias.join("\n") === ADV_SIN_1B,
  rXlsxIndividual.error || rXlsxIndividual.advertencias.join("; "));

// =============================================================================
// 10. Columnas operativas: sólo la fuente principal las exige (smoke test de #12)
// =============================================================================
// Mismo Control para todo el bloque; cambia el período (y con él cuál fuente es la principal).
const ID_HIST_1C = "ID_HIST_1C_xxxxxxxxxxxxxxxxxx";
const conHist1C = cambiar => Object.assign({}, sources, { [ID_HIST_1C]: { EPSI: variante(src1C_real.EPSI, cambiar) } });
const ctlContrato = (periodo, dni, extra) => ctl(periodo, dni || "", Object.assign({
  B9: "EPS2", [F1B]: "ID_1B_xxxxxxxxxxxxxxxxxxxxxxxx", [F1C]: ID_HIST_1C, [F3B]: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx"
}, extra));
const sinOperativas = conEncabezados({ "LISTO PARA EXPORTAR": "", "EXPORTADO": "" });
const avisosOperativos = r => (resumenDe(r).generales || "").split("\n").concat(resumenDe(r).filas.map(f => f.detalle))
  .filter(a => /LISTO PARA EXPORTAR|EXPORTADO/.test(a));
const errPrincipalSinListo = periodo => 'No se encontró una única columna "LISTO PARA EXPORTAR" después de ORIENTACIÓN en la hoja EPSI de la fuente de ' + periodo +
  '. Es la fuente principal del período a generar (' + periodo + ', elegido en Control!B5)';

// 10.1. 3B: principal 3B; 1C y 1B (v0.3) sin columnas operativas
const srcSinOp = conHist1C(sinOperativas);
let r3BSinOp = runCurso("3B con históricas sin columnas operativas", ctlContrato(P3B), srcSinOp);
let ind3BSinOp = runTest("Individual 3B con históricas sin columnas operativas", { control: ctlContrato(P3B, "97000001"), sources: srcSinOp, templateGrid: tpl03 });
assert("Contrato operativo, 3B: 1C y 1B sin columnas operativas son válidas; se genera con los valores de 1C",
  !r3BSinOp.error && hojasDe(r3BSinOp).join("|") === "Resumen|Ficticia, Ana Real"
    && valoresFila(filasMaterias(gridHoja(r3BSinOp, "Ficticia, Ana Real")).find(f => f[2] === "Matemática I") || []) === "|7|Avanzado"
    && avisosOperativos(r3BSinOp).length === 0,
  r3BSinOp.error || hojasDe(r3BSinOp).join("|") + " / " + avisosOperativos(r3BSinOp).join("; "));
assert("Contrato operativo, 3B: la selección sale sólo de la principal (EPS2 de 3B: Ana)",
  resumenDe(r3BSinOp).filas.map(f => f.estado + ":" + f.dni).join(",") === "Generado:97000001", JSON.stringify(resumenDe(r3BSinOp).filas));
assert("Contrato operativo, 3B: la generación individual funciona igual y su informe es idéntico a la hoja del libro",
  !ind3BSinOp.error && JSON.stringify(ind3BSinOp.out) === JSON.stringify(gridHoja(r3BSinOp, "Ficticia, Ana Real"))
    && !ind3BSinOp.advertencias.some(a => /LISTO PARA EXPORTAR|EXPORTADO/.test(a)),
  ind3BSinOp.error || ind3BSinOp.advertencias.join("; "));

// 10.2. 1C: principal 1C (con LISTO PARA EXPORTAR); 1B (v0.3) sin columnas operativas
let r1CPrincipal = runCurso("1C con 1B sin columnas operativas", ctlContrato(P1C), conHist1C(() => {}));
let ind1CPrincipal = runTest("Individual 1C con 1B sin columnas operativas", { control: ctlContrato(P1C, "97000001"), sources: conHist1C(() => {}), templateGrid: tpl03 });
assert("Contrato operativo, 1C: 1B sin columnas operativas es válida; la selección usa LISTO PARA EXPORTAR de 1C",
  !r1CPrincipal.error && hojasDe(r1CPrincipal).join("|") === "Resumen|Ficticia, Ana Real"
    && libroDe(r1CPrincipal).getName().includes(" - 1º cuatrimestre 2026 - ")
    && JSON.stringify(gridHoja(r1CPrincipal, "Ficticia, Ana Real")) === JSON.stringify(ind1CPrincipal.out),
  r1CPrincipal.error || hojasDe(r1CPrincipal).join("|"));

// 10.3. 1B: principal 1B; sin LISTO PARA EXPORTAR falla, con la columna se genera
let r1BSinListo = runCurso("1B sin LISTO PARA EXPORTAR", ctlContrato("1º bimestre"), sources);
assert("Contrato operativo, 1B: la principal 1B sin LISTO PARA EXPORTAR falla claramente, sin libro",
  r1BSinListo.error && r1BSinListo.error.startsWith(errPrincipalSinListo("1º bimestre")) && r1BSinListo.creados.length === 0, r1BSinListo.error);
let r1BConListo = runCurso("1B con LISTO PARA EXPORTAR", ctlContrato("1º bimestre", "", { [F1B]: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx" }), sources);
assert("Contrato operativo, 1B: con LISTO PARA EXPORTAR en 1B se genera desde 1B",
  !r1BConListo.error && hojasDe(r1BConListo).join("|") === "Resumen|Ficticia, Ana Real"
    && valoresFila(filasMaterias(gridHoja(r1BConListo, "Ficticia, Ana Real")).find(f => f[2] === "Matemática I") || []) === "Avanzado||",
  r1BConListo.error || hojasDe(r1BConListo).join("|"));

// 10.4. Históricas con columnas operativas duplicadas, inválidas o malformadas: se ignoran
const r3BBase = runCurso("3B con 1C limpia", ctlContrato(P3B), conHist1C(() => {}));
const mismaTanda = (a, b) => !a.error && !b.error && JSON.stringify(resumenDe(a)) === JSON.stringify(resumenDe(b))
  && hojasDe(a).every(h => JSON.stringify(gridHoja(a, h)) === JSON.stringify(gridHoja(b, h))) && hojasDe(a).join("|") === hojasDe(b).join("|");
let rHistListoDoble = runCurso("Histórica con LISTO duplicado", ctlContrato(P3B), conHist1C(conEncabezados({ "SEGUIMIENTO INTERNO": "LISTO PARA EXPORTAR" })));
assert("Contrato operativo: una histórica con LISTO PARA EXPORTAR duplicada se ignora (misma tanda)",
  mismaTanda(rHistListoDoble, r3BBase) && avisosOperativos(rHistListoDoble).length === 0, rHistListoDoble.error);
let rHistExpInvalido = runCurso("Histórica con EXPORTADO inválido", ctlContrato(P3B),
  conHist1C(conCeldas("97000001", { "EXPORTADO": "Tal vez", "LISTO PARA EXPORTAR": "" })));
assert("Contrato operativo: una histórica con EXPORTADO inválido (y LISTO vacío) se ignora (misma tanda)",
  mismaTanda(rHistExpInvalido, r3BBase) && avisosOperativos(rHistExpInvalido).length === 0, rHistExpInvalido.error);
let rHistMalformada = runCurso("Histórica con encabezados operativos malformados", ctlContrato(P3B),
  conHist1C(conEncabezados({ "LISTO PARA EXPORTAR": "LISTO PARA IMPRIMIR", "EXPORTADO": "EXPORTADO?" })));
assert("Contrato operativo: una histórica con encabezados operativos parecidos o malformados no bloquea (misma tanda)",
  mismaTanda(rHistMalformada, r3BBase), rHistMalformada.error);

// 10.5. La misma grilla sin LISTO PARA EXPORTAR: válida como histórica, error sólo como principal
const src1CSinListo = conHist1C(conEncabezados({ "LISTO PARA EXPORTAR": "" }));
let rSinListoComoHist = runCurso("1C sin LISTO como histórica de 3B", ctlContrato(P3B), src1CSinListo);
let rSinListoComoPrincipal = runCurso("1C sin LISTO como principal", ctlContrato(P1C), src1CSinListo);
assert("Contrato operativo: 1C sin LISTO PARA EXPORTAR es válida como histórica de 3B",
  mismaTanda(rSinListoComoHist, r3BBase), rSinListoComoHist.error);
assert("Contrato operativo: esa misma 1C como principal (período 1º cuatrimestre en B5) falla y lo dice",
  rSinListoComoPrincipal.error && rSinListoComoPrincipal.error.startsWith(errPrincipalSinListo(P1C))
    && rSinListoComoPrincipal.error.endsWith("Las fuentes anteriores no la necesitan.") && rSinListoComoPrincipal.creados.length === 0,
  rSinListoComoPrincipal.error);
let r1CListoDoble = runCurso("1C principal con LISTO duplicado", ctlContrato(P1C), conHist1C(conEncabezados({ "SEGUIMIENTO INTERNO": "LISTO PARA EXPORTAR" })));
assert("Contrato operativo: la principal 1C con LISTO PARA EXPORTAR duplicada falla claramente, sin libro",
  r1CListoDoble.error && r1CListoDoble.error.startsWith(errPrincipalSinListo(P1C)) && r1CListoDoble.creados.length === 0, r1CListoDoble.error);

// 10.6. La generación individual ignora B9 aunque tenga un valor no válido
let rIndB9Invalido = runTest("Individual con B9 no válido", { control: ctlReal("97000001", { B9: "Cualquier cosa" }), sources, templateGrid: tpl03 });
assert("Individual: ignora B9 aunque tenga un valor no válido (mismo informe)",
  !rIndB9Invalido.error && mismoInforme(rIndB9Invalido, rReal), rIndB9Invalido.error);

// =============================================================================
// 11. Control del xlsx: cada coordenada que usa el generador tiene su rótulo, sin filas vacías en el bloque
// =============================================================================
const filaColumna = a1 => a1ToRC(a1).slice(0, 2).map(n => n - 1);
const textoXlsx = a1 => { const [r, c] = filaColumna(a1); return ((controlXlsx[r] || [])[c] || "").trim(); };
const rotuloDe = a1 => { const [r, c] = filaColumna(a1); return ((controlXlsx[r] || [])[c - 1] || "").trim(); };
const rotulosEsperados = {
  CELDA_PERIODO: "Período a generar",
  CELDA_HOJA_ORIGEN: "Hoja origen / orientación",
  CELDA_DNI: "DNI del estudiante",
  CELDA_CICLO: "Ciclo lectivo",
  CELDA_NIVEL: "Nivel / curso",
  CELDA_FUENTE_1B: "URL o ID del archivo de valoraciones del 1º bimestre",
  CELDA_FUENTE_1C: "URL o ID del archivo de valoraciones del 1º cuatrimestre",
  CELDA_FUENTE_3B: "URL o ID del archivo de valoraciones del 3º bimestre",
  ESTADO_FECHA: "Última ejecución",
  ESTADO_RESULTADO: "Resultado",
  ESTADO_ADVERTENCIAS: "Advertencias"
};
const celdasGenerador = Object.keys(CELDAS_CONTROL).filter(k => /^(CELDA|ESTADO)_/.test(k));
assert("Control (xlsx): cada celda que usa el generador está junto a su rótulo esperado",
  celdasGenerador.length === Object.keys(rotulosEsperados).length
    && celdasGenerador.every(k => rotuloDe(CELDAS_CONTROL[k]) === rotulosEsperados[k]),
  celdasGenerador.map(k => k + "=" + CELDAS_CONTROL[k] + ":" + rotuloDe(CELDAS_CONTROL[k])).join(", "));
assert("Control (xlsx): las fuentes quedan en B10, B11 y B12, sin fila vacía entre 1º bimestre y 1º cuatrimestre",
  [F1B, F1C, F3B].join(",") === "B10,B11,B12"
    && Array.from({ length: 8 }, (_, i) => (controlXlsx[4 + i] || [])[0]).every(Boolean),
  [F1B, F1C, F3B].join(",") + " / " + JSON.stringify(controlXlsx.slice(4, 12).map(f => (f || [])[0])));
assert("Control (xlsx): valores por defecto conservados (B5 1º cuatrimestre, B6 EPSI, B8 2026) y B9 vacía",
  textoXlsx("B5") === P1C && textoXlsx("B6") === "EPSI" && textoXlsx("B8") === "2026" && textoXlsx(CELDAS_CONTROL.CELDA_NIVEL) === "",
  ["B5", "B6", "B8", "B9"].map(textoXlsx).join("|"));
const validacionesXlsx = [...xmlHoja(XLSX_GENERADOR, "Control").matchAll(/<(?:\w+:)?dataValidation [^>]*sqref="([^"]+)"[^>]*>\s*<(?:\w+:)?formula1>([^<]*)</g)]
  .map(m => m[1] + "=" + m[2]).join(" | ");
assert("Control (xlsx): las validaciones siguen en período, hoja origen y nivel",
  validacionesXlsx === [
    CELDAS_CONTROL.CELDA_PERIODO + '="1º bimestre,1º cuatrimestre,3º bimestre"',
    CELDAS_CONTROL.CELDA_HOJA_ORIGEN + '="EPSI,EPSG"',
    CELDAS_CONTROL.CELDA_NIVEL + '="EPS1,EPS2,EPS3,EPS4"'
  ].join(" | "), validacionesXlsx);

console.log(`\nResultado final: ${fail === 0 ? "TODOS LOS TESTS PASARON" : `${fail} TESTS FALLARON`}`);
process.exit(fail ? 1 : 0);
