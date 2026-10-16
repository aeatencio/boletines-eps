// Arnés de pruebas local
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

// Las fechas se crean dentro del contexto vm, donde instanceof Date no las reconoce.
const esFecha = v => Object.prototype.toString.call(v) === "[object Date]";

class MockSheet {
  constructor(name, grid, reglas) {
    this.name = name;
    this.grid = grid.map(r => r.slice());
    this.reglas = (reglas || []).slice(); // formato condicional
    this.formatos = {};                   // "fila,columna" → { fondo, texto }, sólo lo que asigna el generador
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
            sh.set(r + i, c + j, esFecha(v) ? "<fecha>" : v);
          }
        }
      },
      setValues: filas => {
        if (filas.length !== h || filas.some(f => f.length !== w)) throw new Error("mock: setValues con dimensiones distintas del rango");
        filas.forEach((f, i) => f.forEach((v, j) => sh.set(r + i, c + j, esFecha(v) ? "<fecha>" : v)));
      },
      clearContent: () => {
        for (let i = 0; i < h; i++) {
          for (let j = 0; j < w; j++) {
            sh.set(r + i, c + j, "");
          }
        }
      },
      setWrap: () => {},
      setFontSize: n => {
        for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) sh.formato(r + i, c + j).tamano = n;
      },
      setBackgrounds: m => {
        if (m.length !== h || m.some(f => f.length !== w)) throw new Error("mock: setBackgrounds con dimensiones distintas del rango");
        m.forEach((f, i) => f.forEach((v, j) => { sh.formato(r + i, c + j).fondo = v; }));
      },
      setFontColors: m => {
        if (m.length !== h || m.some(f => f.length !== w)) throw new Error("mock: setFontColors con dimensiones distintas del rango");
        m.forEach((f, i) => f.forEach((v, j) => { sh.formato(r + i, c + j).texto = v; }));
      },
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
  formato(r, c) { return this.formatos[r + "," + c] || (this.formatos[r + "," + c] = {}); }
  getConditionalFormatRules() { return this.reglas.slice(); }
  setConditionalFormatRules(reglas) { this.reglas = reglas.slice(); }
  setFrozenRows(n) { this.filasInmovilizadas = n; }
  getFrozenRows() { return this.filasInmovilizadas || 0; }
  // Como Google Sheets: no se puede ocultar la última hoja visible.
  hideSheet() {
    if (this.ss && this.ss.sheets.every(s => s === this || s.hidden)) throw new Error("mock: no se pueden ocultar todas las hojas");
    this.hidden = true;
  }
  isSheetHidden() { return !!this.hidden; }
  copyTo(ss) {
    const s = new MockSheet("Copia de " + this.name, this.grid, this.reglas);
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
  // Distinta de la de Buenos Aires, como el generador real observado: el nombre del libro no debe depender de ella.
  getSpreadsheetTimeZone() { return "America/Los_Angeles"; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
  setActiveSheet() {}
  toast(mensaje) { (this.toasts = this.toasts || []).push(mensaje); }
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

// Formato real de Utilities.formatDate para el único patrón que usa el generador.
function formatearFecha(fecha, zona, formato) {
  if (formato !== "yyyy-MM-dd HH:mm") throw new Error("mock: formato no soportado " + formato);
  return new Intl.DateTimeFormat("sv-SE", {
    timeZone: zona, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23"
  }).format(fecha);
}

function runTest(testName, { control, sources, templateGrid, ejecutar = "generarInformeActual()", ahora = null, reglasPlantilla = [] }) {
  const code = fs.readFileSync(path.join(ROOT, "src/Generador_EPS.gs"), "utf8");
  const plantilla = new MockSheet("Plantilla", templateGrid, reglasPlantilla);
  const ctrl = hojaControl(control);
  const ss = new MockSS([plantilla, ctrl]);
  const rawAlerts = [];
  const creados = [];
  const registro = { aperturas: [], lecturas: [], indebidos: [] };
  const formatos = [];
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
    // Sin `ahora`, la fecha queda como "<fecha>"; con `ahora`, Date se fija a ese instante y se formatea de verdad.
    Utilities: {
      formatDate: (fecha, zona, formato) => {
        formatos.push(zona + " " + formato);
        return ahora ? formatearFecha(fecha, zona, formato) : "<fecha>";
      }
    }
  };
  if (ahora) {
    const fija = Date.parse(ahora);
    ctx.Date = class extends Date {
      constructor(...args) { if (args.length) super(...args); else super(fija); }
      static now() { return fija; }
    };
  }
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
    registro,
    formatos
  };
}

// Carga el generador sin ejecutarlo, para probar funciones puras del parser.
// Con fuentes, también puede abrirlas (sin hoja activa ni UI).
function cargarGenerador(sources) {
  const ctx = sources ? { SpreadsheetApp: { openById: abrirFuenteMock(sources) } } : {};
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(ROOT, "src/Generador_EPS.gs"), "utf8"), ctx);
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

const tpl = readXlsx("src/Generador_EPS.xlsx").Plantilla;
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
let r = runTest("Happy path 3B", { control: ctl("3º bimestre", "99000001"), sources, templateGrid: tpl });
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
r = runTest("Materia que aparece por primera vez en 3B", { control: ctl("3º bimestre", "99000003"), sources, templateGrid: tpl });
assert("Materia incorporada en 3B: sin errores", !r.error, r.error);
assert("Materia incorporada en 3B: nivel visual es 'Informática 1'", r.out && r.out[5] && r.out[5][4] === "Informática 1", r.out ? r.out[5][4] : null);
let filaIncorporada = (r.out || []).find(f => f[2] === "Taller C") || [];
assert("Materia incorporada en 3B: Taller C solo tiene valor en 3B (||Suficiente)", filaIncorporada.slice(4, 7).join("|") === "||Suficiente", filaIncorporada.slice(4, 7));

// =============================================================================
// 3. Caso separado: Valoración inválida en 3B (descarte + advertencia + continuidad)
// =============================================================================
let ctlInv = ctl("3º bimestre", "99000001", { [F3B]: "ID_3B_INV_xxxxxxxxxxxxxxxxxxxx" });
let rInv = runTest("Valoración inválida 3B", { control: ctlInv, sources, templateGrid: tpl });
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
let r1B = runTest("Regresión 1B (CP-01)", { control: ctl("1º bimestre", "99000001", { [F1C]: "", [F3B]: "" }), sources, templateGrid: tpl });
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
let r1C = runTest("Regresión 1C acumulativo (CP-02)", { control: ctl("1º cuatrimestre", "99000001", { [F3B]: "" }), sources, templateGrid: tpl });
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
let r1C_inv = runTest("Regresión 1C valor no válido (CP-03)", { control: ctl("1º cuatrimestre", "99000002", { [F3B]: "" }), sources, templateGrid: tpl });
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
  templateGrid: tpl
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
let rReal = runTest("Estructura real 3B", { control: ctlReal("97000001"), sources, templateGrid: tpl });
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
let rNuevo = runTest("Estudiante nuevo 3B", { control: ctlReal("97000002"), sources, templateGrid: tpl });
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
  control: ctlReal("96000001", { B6: "EPSG", [F1C]: "" }), sources, templateGrid: tpl
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
  templateGrid: tpl
});
assert("Sin ORIENTACIÓN en la principal: se informa error y no se genera",
  rSinOrient.error && rSinOrient.error.includes("No se encontró la columna ORIENTACIÓN") && !rSinOrient.out, rSinOrient.error);

let rOrientDoble = runTest("ORIENTACIÓN repetida", {
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "SEGUIMIENTO INTERNO")] = "Orientación"; }) }),
  templateGrid: tpl
});
assert("ORIENTACIÓN repetida en la principal: se informa error",
  rOrientDoble.error && rOrientDoble.error.includes("La columna ORIENTACIÓN aparece más de una vez (columnas AN, AU)"), rOrientDoble.error);

let rRepetida = runTest("Materia repetida en principal", {
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "Matemática IV")] = "Matemática I"; }) }),
  templateGrid: tpl
});
assert("Materia repetida en la principal: se informa error con ambas columnas, sin elegir una",
  rRepetida.error && rRepetida.error.includes('Materia repetida en la fuente de 3º bimestre: "Matemática I" (columnas E, AH)') && !rRepetida.out,
  rRepetida.error);

let rRepetidaHist = runTest("Materia repetida en fuente anterior", {
  control: ctlReal("97000001", { [F1C]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src1C_real.EPSI, g => { g[0][columnaDe(g, "Matemática IV")] = "Matemática I"; }) }),
  templateGrid: tpl
});
assert("Materia repetida en una fuente anterior: advertencia, se descarta esa fuente y se genera con la principal",
  !rRepetidaHist.error && rRepetidaHist.advertencias.some(a => a.startsWith("Materia repetida en la fuente de 1º cuatrimestre") && a.endsWith("Se continuará con la fuente principal."))
    && valoresFila(filasMaterias(rRepetidaHist.out).find(f => f[2] === "Matemática I") || []) === "||Avanzado",
  rRepetidaHist.error || rRepetidaHist.advertencias.join("; "));

let rParRoto = runTest("Materia sin columna de valoración", {
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { g[0][columnaDe(g, "Matemática I") + 1] = "Columna intrusa"; }) }),
  templateGrid: tpl
});
assert("Materia sin su columna de valoración: se informa error",
  rParRoto.error && rParRoto.error.includes('La materia "Matemática I" (columna E) no tiene a continuación su columna de valoración'), rParRoto.error);

let rFilaDni = runTest("DNI fuera de la fila Valoración", {
  control: ctlReal("97000001", { [F3B]: ID_VAR }),
  sources: conVariante(ID_VAR, { EPSI: variante(src3B_real.EPSI, g => { const r = filaDni(g, "97000001"); g[r][2] = ""; g[r + 1][2] = "97000001"; }) }),
  templateGrid: tpl
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
  templateGrid: tpl
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
let rOpVacio = runTest("LISTO PARA EXPORTAR vacío", { control: ctlAna3B, sources: srcOpVacio, templateGrid: tpl });
let dOpVacio = generarDatos(ctlAna3B, srcOpVacio);
assert("Celda operativa vacía: advertencia no bloqueante y mismo informe",
  !rOpVacio.error && mismoInforme(rOpVacio, rReal)
    && rOpVacio.advertencias.includes("LISTO PARA EXPORTAR está vacío en la fuente de 3º bimestre. Se esperaba SÍ o NO."),
  rOpVacio.error || rOpVacio.advertencias.join("; "));
assert("Celda operativa vacía: el estado queda nulo",
  operativos(dOpVacio) === '{"listoParaExportar":null,"exportado":"NO"}', operativos(dOpVacio));

let srcOpInvalido = variante3B(conCeldas("97000001", { "EXPORTADO": "Tal vez" }));
let rOpInvalido = runTest("EXPORTADO inválido", { control: ctlAna3B, sources: srcOpInvalido, templateGrid: tpl });
let dOpInvalido = generarDatos(ctlAna3B, srcOpInvalido);
assert("Valor operativo inválido: advertencia no bloqueante y mismo informe",
  !rOpInvalido.error && mismoInforme(rOpInvalido, rReal)
    && rOpInvalido.advertencias.includes('EXPORTADO tiene un valor no admitido en la fuente de 3º bimestre. Valor: "Tal vez". Se esperaba SÍ o NO; el estado queda sin determinar.'),
  rOpInvalido.error || rOpInvalido.advertencias.join("; "));
assert("Valor operativo inválido: el estado queda nulo",
  operativos(dOpInvalido) === '{"listoParaExportar":"SÍ","exportado":null}', operativos(dOpInvalido));

let rOpDoble = runTest("EXPORTADO repetido", {
  control: ctlAna3B, sources: variante3B(conEncabezados({ "SEGUIMIENTO INTERNO": "Exportado" })), templateGrid: tpl
});
assert("Encabezado operativo repetido: advertencia, no se elige una columna y se genera igual",
  !rOpDoble.error && mismoInforme(rOpDoble, rReal)
    && rOpDoble.advertencias.includes('El encabezado "EXPORTADO" aparece más de una vez (columnas AT, AU) en la fuente de 3º bimestre. No se leerá ese campo.'),
  rOpDoble.error || rOpDoble.advertencias.join("; "));

// 6.3. LISTO PARA EXPORTAR y EXPORTADO todavía no seleccionan ni alteran la generación individual
let rNoListo = runTest("LISTO PARA EXPORTAR = NO", {
  control: ctlAna3B, sources: variante3B(conCeldas("97000001", { "LISTO PARA EXPORTAR": "NO", "EXPORTADO": "SÍ" })), templateGrid: tpl
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
  control: ctlAna3B, sources: variante3B(conEncabezados({ "Matemática IV": "LISTO PARA EXPORTAR" })), templateGrid: tpl
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
let rHistIntrusa = runTest("Fuente histórica con cola distinta e inválida", { control: ctlHistIntrusa, sources: srcHistIntrusa, templateGrid: tpl });
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
let rHuecos3B = runTest("Campos generales vacíos en 3B", { control: ctlAna3B, sources: srcHuecos3B, templateGrid: tpl });
let dHuecos3B = generarDatos(ctlAna3B, srcHuecos3B);
assert("Generación 3B: un campo general vacío en 3B queda vacío y nunca se completa con 1C",
  !rHuecos3B.error && cajas(rHuecos3B.out) === "Obs 3B Ana||Conv 3B Ana|" && !JSON.stringify(rHuecos3B.out).includes("1C Ana")
    && dHuecos3B.informe.estudiante.orientacion === "",
  rHuecos3B.error || cajas(rHuecos3B.out));

// 6.7. Generación de 1C: ahora 1C es la principal y sus campos generales y operativos sí se usan
const ctl1CReal = ctl(P1C, "97000001", { [F1B]: "", [F1C]: "ID_1C_REAL_xxxxxxxxxxxxxxxxxxx", [F3B]: "ID_3B_REAL_xxxxxxxxxxxxxxxxxxx" });
let r1CReal = runTest("Generación 1C con 1C como principal", { control: ctl1CReal, sources, templateGrid: tpl });
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
  control: ctlAna3B, sources: variante3B(conEncabezados({ "OBSERVACIONES SOBRE ASISTENCIA": "ASISTENCIA" })), templateGrid: tpl
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
let rColaSinEnc = runTest("Cola con encabezados generales vacíos", { control: ctlAna3B, sources: srcColaSinEnc, templateGrid: tpl });
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
let rSmoke = runTest("Smoke: estudiante sólo en 3B", { control: ctlSmoke(P3B), sources, templateGrid: tpl });
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
let rSmoke1C = runTest("Smoke: mismo estudiante generando 1C", { control: ctlSmoke(P1C), sources, templateGrid: tpl });
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
const runCurso = (nombre, control, srcs) => runTest(nombre, { control, sources: srcs, templateGrid: tpl, ejecutar: "generarBoletinesPorCurso()" });
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
  const ind = runTest("Individual " + dni, { control: ctlIndividual(dni), sources: srcCurso, templateGrid: tpl });
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
    && JSON.stringify(rect(rCurso.ss.getSheetByName("Plantilla").grid)) === JSON.stringify(rect(tpl)));
assert("Por curso: Control registra el libro, los conteos y las advertencias generales",
  rCurso.estadoResultado.startsWith("Libro creado: " + NOMBRE_LIBRO_EPS2 + " · Generados: 4 · Omitidos: 3 · Con error: 0.")
    && rCurso.estadoAdvertencias === ADV_SIN_1B && rCurso.rawAlerts.length === 1 && rCurso.rawAlerts[0].startsWith("Libro creado: "),
  rCurso.estadoResultado + " / " + rCurso.estadoAdvertencias);

// 8.2. Otro nivel del mismo Control (nivel indicado como "1")
let rCurso1 = runCurso("Curso EPS1", ctlCurso("1"), srcCurso);
let indBeto = runTest("Individual Beto", { control: ctlIndividual("97000002"), sources: srcCurso, templateGrid: tpl });
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

// 8.6. Menú: la generación por curso y la de varios cursos son acciones separadas y la individual se conserva
const menu = [];
const genMenu = cargarGenerador();
genMenu.SpreadsheetApp = { getUi: () => ({ createMenu: () => ({ addItem(t, f) { menu.push(t + "→" + f); return this; }, addToUi() {} }) }) };
genMenu.onOpen();
assert("Menú: conserva 'Generar informe actual' y 'Generar boletines por curso', y agrega los cursos marcados y todos los cursos",
  menu.join(", ") === "Generar informe actual→generarInformeActual, Generar boletines por curso→generarBoletinesPorCurso, "
    + "Generar boletines de los cursos marcados→generarBoletinesCursosMarcados, Generar boletines de todos los cursos→generarBoletinesTodosLosCursos",
  menu.join(", "));

// 8.7. Resumen oculto: existe y conserva el detalle, pero no se imprime ni se exporta con el libro
const visibilidad = r => libroDe(r) ? libroDe(r).getSheets().map(s => s.getName() + (s.isSheetHidden() ? ":oculta" : ":visible")).join(" | ") : "";
assert("Resumen oculto: queda oculta y todos los boletines quedan visibles",
  visibilidad(rCurso) === HOJAS_EPS2.map((h, i) => h + (i === 0 ? ":oculta" : ":visible")).join(" | ") && resCurso.filas.length === 7,
  visibilidad(rCurso));
assert("Resumen oculto: también cuando hay errores, sin perder el detalle del estudiante fallido",
  visibilidad(rDuplicado) === "Resumen:oculta | Ficticia, Ana Real:visible" && enResumen(rDuplicado, "97000017").estado === "Error",
  visibilidad(rDuplicado));
let rSinHojas = runTest("Ninguna hoja pudo crearse", {
  control: ctlCurso("EPS2"), sources: srcCurso, templateGrid: tpl,
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
const XLSX_GENERADOR = "src/Generador_EPS.xlsx";
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
  sources, templateGrid: tpl
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
let ind3BSinOp = runTest("Individual 3B con históricas sin columnas operativas", { control: ctlContrato(P3B, "97000001"), sources: srcSinOp, templateGrid: tpl });
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
let ind1CPrincipal = runTest("Individual 1C con 1B sin columnas operativas", { control: ctlContrato(P1C, "97000001"), sources: conHist1C(() => {}), templateGrid: tpl });
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
let rIndB9Invalido = runTest("Individual con B9 no válido", { control: ctlReal("97000001", { B9: "Cualquier cosa" }), sources, templateGrid: tpl });
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
assert("Control (xlsx): las validaciones siguen en período, hoja origen y nivel, y las marcas de los cursos son SÍ / NO",
  validacionesXlsx === [
    CELDAS_CONTROL.CELDA_PERIODO + '="1º bimestre,1º cuatrimestre,3º bimestre"',
    CELDAS_CONTROL.CELDA_HOJA_ORIGEN + '="EPSI,EPSG"',
    CELDAS_CONTROL.CELDA_NIVEL + '="EPS1,EPS2,EPS3,EPS4"',
    CELDAS_CONTROL.CURSOS[0].celda + ":" + CELDAS_CONTROL.CURSOS[7].celda + '="SÍ,NO"'
  ].join(" | "), validacionesXlsx);

// =============================================================================
// 12. Orientación por período: historial cuando el estudiante cambia de orientación (#15)
// =============================================================================
// Cada estudiante declara, por período, en qué hoja(s) figura y con qué valoraciones. Los tres archivos se
// arman con las grillas reales sintéticas: para 1B y 3B sirven las de 3B (mismos valores válidos) y para
// 1C EPSG, la estructura EPSG de 3B. Cada estudiante es una copia del bloque de un modelo, sin sus valoraciones.
const valorar = (dni, valores) => g => {
  const r = filaDni(g, dni);
  g[r].forEach((v, c) => { if (v === "Valoración") g[r][c + 1] = ""; });
  Object.entries(valores).forEach(([materia, v]) => {
    const c = [0, 1].map(i => g[i].indexOf(materia)).find(col => col >= 0);
    if (c === undefined) throw new Error("El fixture no tiene la materia " + materia);
    g[r][c + 1] = v;
  });
};
const TRAYECTOS = {
  "95000001": { nombres: "Ana Cambio", "1B": [["EPSI"], { "Matemática I": "Suficiente", "Taller Profesional I A": "En proceso" }],
    "1C": [["EPSI"], { "Matemática I": "8", "Taller Profesional I A": "7" }], "3B": [["EPSG"], { "Matemática I": "Avanzado", "Taller Gastronómico I A": "Suficiente" }] },
  "95000002": { nombres: "Bruno Cambio", "1B": [["EPSI"], { "Matemática I": "En proceso" }],
    "1C": [["EPSG"], { "Matemática I": "6", "Taller Gastronómico I A": "9" }], "3B": [["EPSG"], { "Matemática I": "Suficiente", "Taller Gastronómico I A": "Avanzado" }] },
  "95000003": { nombres: "Ciro Nuevo", "3B": [["EPSG"], { "Matemática I": "Avanzado" }] },
  "95000004": { nombres: "Fabi Multiple", "1B": [["EPSI"], { "Matemática I": "Suficiente" }],
    "1C": [["EPSG"], { "Matemática I": "9", "Taller Gastronómico I A": "8" }], "3B": [["EPSI"], { "Matemática I": "Avanzado", "Taller Profesional I A": "Suficiente" }] },
  "95000005": { nombres: "Gina Vuelta", "1B": [["EPSG"], { "Matemática I": "Avanzado", "Taller Gastronómico I A": "En proceso" }],
    "1C": [["EPSG"], { "Matemática I": "7" }], "3B": [["EPSI"], { "Matemática I": "Suficiente" }] },
  "95000006": { nombres: "Dario Doble", "1B": [["EPSG"], { "Matemática I": "Suficiente" }],
    "1C": [["EPSI", "EPSG"], { "Matemática I": "5" }], "3B": [["EPSG"], { "Matemática I": "Avanzado" }] },
  "95000007": { nombres: "Elsa Doble", "3B": [["EPSI", "EPSG"], { "Matemática I": "Avanzado" }] }
};
const BASES_TRAYECTO = {
  "1B": { EPSI: [src3B_real.EPSI, "97000002"], EPSG: [src3B_real.EPSG, "96000001"] },
  "1C": { EPSI: [src1C_real.EPSI, "97000001"], EPSG: [src3B_real.EPSG, "96000001"] },
  "3B": { EPSI: [src3B_real.EPSI, "97000002"], EPSG: [src3B_real.EPSG, "96000001"] }
};
const archivoTrayectos = (periodo, cambiar = {}) => Object.fromEntries(["EPSI", "EPSG"].map(hoja => {
  const [base, modelo] = BASES_TRAYECTO[periodo][hoja];
  return [hoja, variante(base, g => {
    let anterior = modelo;
    Object.entries(TRAYECTOS).filter(([, t]) => t[periodo] && t[periodo][0].includes(hoja)).forEach(([dni, t]) => {
      const operativos = periodo === "3B" ? { listo: "SÍ", exportado: "NO" } : {};
      conBloques(modelo, anterior, [Object.assign({ apellido: "Trayecto", nombres: t.nombres, dni }, operativos)])(g);
      valorar(dni, t[periodo][1])(g);
      anterior = dni;
    });
    if (cambiar[hoja]) cambiar[hoja](g);
  })];
}));
const ID_T1B = "ID_TRAYECTO_1B_xxxxxxxxxxxxxx";
const ID_T1C = "ID_TRAYECTO_1C_xxxxxxxxxxxxxx";
const ID_T3B = "ID_TRAYECTO_3B_xxxxxxxxxxxxxx";
const srcTrayectos = (cambiar1C) => Object.assign({}, sources, {
  [ID_T1B]: archivoTrayectos("1B"), [ID_T1C]: archivoTrayectos("1C", cambiar1C), [ID_T3B]: archivoTrayectos("3B")
});
const ctlTrayecto = (hoja, dni) => ctl(P3B, dni || "", { B6: hoja, B9: "EPS1", [F1B]: ID_T1B, [F1C]: ID_T1C, [F3B]: ID_T3B });
const filaMateria = (r, hoja, materia) => valoresFila(filasMaterias(gridHoja(r, hoja)).find(f => f[2] === materia) || []);
const tomadoDe = (periodo, hoja, actual) => "El " + periodo + " se tomó de la hoja " + hoja + " de su fuente, distinta de la orientación actual (" + actual + ").";

const SRC_TRAYECTOS = srcTrayectos();
const fotoTrayectos = JSON.stringify(SRC_TRAYECTOS);
let rTrayEPSG = runCurso("Tanda EPSG con cambios de orientación", ctlTrayecto("EPSG"), SRC_TRAYECTOS);
let rTrayEPSI = runCurso("Tanda EPSI con cambios de orientación", ctlTrayecto("EPSI"), SRC_TRAYECTOS);
const estadosDe = r => resumenDe(r).filas.map(f => f.estado + ":" + f.dni).join(", ");

assert("Orientación por período: la tanda EPSG continúa ante errores y respeta el orden de la grilla",
  !rTrayEPSG.error && estadosDe(rTrayEPSG) === "Omitido:96000001, Generado:95000001, Generado:95000002, Generado:95000003, Error:95000006, Error:95000007"
    && hojasDe(rTrayEPSG).join("|") === "Resumen|Trayecto, Ana Cambio|Trayecto, Bruno Cambio|Trayecto, Ciro Nuevo",
  rTrayEPSG.error || estadosDe(rTrayEPSG));
assert("Caso A (1B EPSI, 1C EPSI, 3B EPSG): conserva 1B y 1C desde EPSI, también las materias que sólo existen en EPSI",
  filaMateria(rTrayEPSG, "Trayecto, Ana Cambio", "Matemática I") === "Suficiente|8|Avanzado"
    && filaMateria(rTrayEPSG, "Trayecto, Ana Cambio", "Taller Gastronómico I A") === "||Suficiente"
    && filaMateria(rTrayEPSG, "Trayecto, Ana Cambio", "Taller Profesional I A") === "En proceso|7|",
  filasMaterias(gridHoja(rTrayEPSG, "Trayecto, Ana Cambio")).map(f => f[2] + ":" + valoresFila(f)).join(", "));
assert("Caso A: advierte de qué hoja se tomó cada período anterior y muestra la orientación actual",
  enResumen(rTrayEPSG, "95000001").detalle.split("\n").join("|") === [tomadoDe(P1C, "EPSI", "EPSG"), tomadoDe("1º bimestre", "EPSI", "EPSG")].join("|")
    && gridHoja(rTrayEPSG, "Trayecto, Ana Cambio")[5][4] === "Gastronomía 1",
  enResumen(rTrayEPSG, "95000001").detalle);
assert("Caso B (1B EPSI, 1C EPSG, 3B EPSG): 1C de la hoja actual y 1B de EPSI, con una sola advertencia",
  filaMateria(rTrayEPSG, "Trayecto, Bruno Cambio", "Matemática I") === "En proceso|6|Suficiente"
    && filaMateria(rTrayEPSG, "Trayecto, Bruno Cambio", "Taller Gastronómico I A") === "|9|Avanzado"
    && enResumen(rTrayEPSG, "95000002").detalle === tomadoDe("1º bimestre", "EPSI", "EPSG"),
  enResumen(rTrayEPSG, "95000002").detalle);
assert("Caso C (estudiante nuevo, sólo en 3B): se genera; la ausencia en cada período anterior se advierte como antes",
  filaMateria(rTrayEPSG, "Trayecto, Ciro Nuevo", "Matemática I") === "||Avanzado"
    && enResumen(rTrayEPSG, "95000003").detalle.split("\n").join("|") === [advDniHistorico(P1C), advDniHistorico("1º bimestre")].join("|"),
  enResumen(rTrayEPSG, "95000003").detalle);
assert("Caso D (DNI en EPSI y EPSG de 1C): error del estudiante, sin hoja ni combinación",
  enResumen(rTrayEPSG, "95000006").detalle.startsWith("El DNI aparece en las hojas EPSG y EPSI de la fuente de 1º cuatrimestre. No se puede decidir cuál corresponde a ese período ni combinarlas")
    && !hojasDe(rTrayEPSG).some(h => /Dario/.test(h)),
  enResumen(rTrayEPSG, "95000006").detalle);
assert("Caso E (DNI en EPSG y EPSI de la principal): error del estudiante en la tanda EPSG",
  enResumen(rTrayEPSG, "95000007").detalle === "El DNI aparece en las hojas EPSG y EPSI de la fuente de 3º bimestre. No se puede decidir en qué orientación cursa este período. Es la fuente principal del período a generar (3º bimestre), por eso no se puede continuar.",
  enResumen(rTrayEPSG, "95000007").detalle);
assert("Caso E: tampoco se genera en la tanda EPSI, así que no hay dos boletines del mismo estudiante",
  !rTrayEPSI.error && estadosDe(rTrayEPSI) === "Omitido:97000002, Generado:95000004, Generado:95000005, Error:95000007"
    && enResumen(rTrayEPSI, "95000007").detalle.startsWith("El DNI aparece en las hojas EPSI y EPSG de la fuente de 3º bimestre.")
    && ![rTrayEPSG, rTrayEPSI].some(r => hojasDe(r).some(h => /Elsa/.test(h))),
  rTrayEPSI.error || estadosDe(rTrayEPSI));
assert("Caso F (EPSI → EPSG → EPSI): cada período sale de su hoja y sólo 1C advierte",
  filaMateria(rTrayEPSI, "Trayecto, Fabi Multiple", "Matemática I") === "Suficiente|9|Avanzado"
    && filaMateria(rTrayEPSI, "Trayecto, Fabi Multiple", "Taller Profesional I A") === "||Suficiente"
    && filaMateria(rTrayEPSI, "Trayecto, Fabi Multiple", "Taller Gastronómico I A") === "|8|"
    && enResumen(rTrayEPSI, "95000004").detalle === tomadoDe(P1C, "EPSG", "EPSI"),
  enResumen(rTrayEPSI, "95000004").detalle);
assert("Cambio EPSG → EPSI (1B y 1C EPSG, 3B EPSI): conserva ambos desde EPSG y muestra la orientación actual",
  filaMateria(rTrayEPSI, "Trayecto, Gina Vuelta", "Matemática I") === "Avanzado|7|Suficiente"
    && filaMateria(rTrayEPSI, "Trayecto, Gina Vuelta", "Taller Gastronómico I A") === "En proceso||"
    && enResumen(rTrayEPSI, "95000005").detalle.split("\n").join("|") === [tomadoDe(P1C, "EPSG", "EPSI"), tomadoDe("1º bimestre", "EPSG", "EPSI")].join("|")
    && gridHoja(rTrayEPSI, "Trayecto, Gina Vuelta")[5][4] === "Informática 1",
  enResumen(rTrayEPSI, "95000005").detalle);

// Mismo comportamiento en la generación individual
const indTrayecto = (hoja, dni) => runTest("Individual " + dni, { control: ctlTrayecto(hoja, dni), sources: SRC_TRAYECTOS, templateGrid: tpl });
[["EPSG", "95000001", "Trayecto, Ana Cambio"], ["EPSI", "95000004", "Trayecto, Fabi Multiple"]].forEach(([hoja, dni, nombreHoja]) => {
  const ind = indTrayecto(hoja, dni);
  const tanda = hoja === "EPSG" ? rTrayEPSG : rTrayEPSI;
  assert("Orientación por período, individual " + dni + ": mismo informe y mismas advertencias que en la tanda",
    !ind.error && JSON.stringify(ind.out) === JSON.stringify(gridHoja(tanda, nombreHoja))
      && JSON.stringify(advertenciasTanda(tanda, dni).sort()) === JSON.stringify(ind.advertencias.slice().sort()),
    ind.error || ind.advertencias.join("; "));
});
const indDario = indTrayecto("EPSG", "95000006");
const indElsa = indTrayecto("EPSG", "95000007");
assert("Orientación por período, individual: los DNI ambiguos (histórica y principal) impiden generar, con el mismo motivo",
  indDario.error === enResumen(rTrayEPSG, "95000006").detalle && !indDario.out
    && indElsa.error === enResumen(rTrayEPSG, "95000007").detalle && !indElsa.out
    && indElsa.rawAlerts[0] === "No se pudo generar el informe.\n\n" + indElsa.error,
  (indDario.error || "") + " / " + (indElsa.error || ""));

// Lectura acotada y sólo lectura
const hojasLeidas = [ID_T3B + "/EPSG", ID_T3B + "/EPSI", ID_T1C + "/EPSG", ID_T1C + "/EPSI", ID_T1B + "/EPSG", ID_T1B + "/EPSI"].sort();
assert("Orientación por período: en la tanda cada archivo se abre una vez y cada hoja se lee una vez, no por estudiante",
  JSON.stringify(rTrayEPSG.registro.aperturas.slice().sort()) === JSON.stringify([ID_T1B, ID_T1C, ID_T3B].sort())
    && JSON.stringify(rTrayEPSG.registro.lecturas.slice().sort()) === JSON.stringify(hojasLeidas),
  JSON.stringify(rTrayEPSG.registro));
assert("Orientación por período: ningún acceso de escritura y fuentes intactas",
  [rTrayEPSG, rTrayEPSI].every(r => r.registro.indebidos.length === 0) && JSON.stringify(SRC_TRAYECTOS) === fotoTrayectos,
  rTrayEPSG.registro.indebidos.concat(rTrayEPSI.registro.indebidos).join(", "));

// Una hoja ilegible en una histórica (aquí, la de la orientación actual): se advierte una vez, se sigue con la otra y no se elige en silencio
const rTrayRota = runCurso("1C EPSG sin ORIENTACIÓN", ctlTrayecto("EPSG"),
  srcTrayectos({ EPSG: g => { g[0][colDe(g, "ORIENTACIÓN")] = ""; } }));
assert("Hoja ilegible en una histórica: advertencia general única; la otra hoja se usa y el DNI en ambas sigue siendo error",
  !rTrayRota.error
    && resumenDe(rTrayRota).generales === "No se encontró la columna ORIENTACIÓN que separa las materias de los campos generales en la fuente de 1º cuatrimestre. No se usará la hoja EPSG de esa fuente."
    && filaMateria(rTrayRota, "Trayecto, Ana Cambio", "Matemática I") === "Suficiente|8|Avanzado"
    && filaMateria(rTrayRota, "Trayecto, Bruno Cambio", "Matemática I") === "En proceso||Suficiente"
    && enResumen(rTrayRota, "95000006").estado === "Error",
  rTrayRota.error || resumenDe(rTrayRota).generales);

const textosOrientacion = [textoXlsx("C6"), (controlXlsx.find(f => /^La hoja origen/.test((f || [])[0] || "")) || [])[0] || ""];
assert("Control (xlsx): los textos ya no afirman una orientación única; B6 es la del período a generar y las anteriores se buscan en ambas hojas",
  textosOrientacion.every(t => !/[Úú]nica/.test(t) && /período a generar/.test(t) && /anteriores/.test(t)), textosOrientacion.join(" / "));

// =============================================================================
// 13. Libro por curso: hora de Buenos Aires en el nombre (#12)
// =============================================================================
const runCursoHora = (nombre, ahora) => runTest(nombre, {
  control: ctlCurso("EPS2"), sources: srcCurso, templateGrid: tpl, ejecutar: "generarBoletinesPorCurso()", ahora
});

// La zona del libro generador en el mock es America/Los_Angeles: el nombre no debe depender de ella.
let rHora = runCursoHora("Hora de Buenos Aires", "2026-10-14T21:07:00Z");
assert("Hora: el nombre del libro lleva la hora de Buenos Aires (21:07 UTC → 18:07), no la del libro generador (14:07)",
  !rHora.error && libroDe(rHora).getName() === "Boletines EPS - EPSI - EPS2 - 3º bimestre 2026 - 2026-10-14 18:07"
    && rHora.ss.getSpreadsheetTimeZone() === "America/Los_Angeles",
  rHora.error || libroDe(rHora).getName());
assert("Hora: se formatea con la zona explícita America/Argentina/Buenos_Aires y el formato yyyy-MM-dd HH:mm",
  JSON.stringify(rHora.formatos) === JSON.stringify(["America/Argentina/Buenos_Aires yyyy-MM-dd HH:mm"]), rHora.formatos.join(", "));
let rHoraNoche = runCursoHora("Hora de Buenos Aires, otro día en UTC", "2026-10-15T02:30:00Z");
assert("Hora: también la fecha es la de Buenos Aires (02:30 UTC del 15 → 2026-10-14 23:30)",
  !rHoraNoche.error && libroDe(rHoraNoche).getName().endsWith(" - 2026-10-14 23:30"), rHoraNoche.error || libroDe(rHoraNoche).getName());
assert("Libro por curso: un único libro nuevo, completo, y el diálogo muestra su nombre y su enlace",
  rHora.creados.length === 1 && JSON.stringify(hojasDe(rHora)) === JSON.stringify(HOJAS_EPS2)
    && rHora.rawAlerts[0].startsWith("Libro creado: " + libroDe(rHora).getName() + "\n" + libroDe(rHora).getUrl() + "\n\n"),
  rHora.rawAlerts[0]);
assert("Sin libro (nadie listo) y generación individual: no se crea ningún libro ni se arma su nombre",
  [rNadieListo, rReal, rSmoke].every(r => r.creados.length === 0 && r.formatos.length === 0));

// =============================================================================
// 14. Impresión: aire superior en las páginas siguientes a la primera
// =============================================================================
// El paginado real de Google Sheets no se modela: se protege el mecanismo (fila 1 inmovilizada, para
// «Repetir filas inmovilizadas») y que el contenido y la cantidad de filas no cambien.
const inmovilizadas = hoja => hoja ? hoja.getFrozenRows() : null;
// Fila (1-based) de la etiqueta del PTF, último bloque del boletín: 28 en la Plantilla, corrida sólo por las materias.
const filaPtf = out => (out || []).findIndex(f => f[2].startsWith("PLAN DE TRABAJO FORMATIVO")) + 1;

// 14.1. Boletín corto (Beto, 3 materias): sólo se agrega la fila inmovilizada, sin filas nuevas
assert("Impresión, boletín corto: Informe_actual tiene la fila 1 inmovilizada y ninguna fila más que antes",
  !rNuevo.error && inmovilizadas(rNuevo.ss.getSheetByName("Informe_actual")) === 1
    && filaPtf(rNuevo.out) === 28 - (8 - 3) && filasMaterias(rNuevo.out).length === 3,
  rNuevo.error || "PTF en la fila " + filaPtf(rNuevo.out));
assert("Impresión: la Plantilla no se modifica (sin filas inmovilizadas)",
  inmovilizadas(rNuevo.ss.getSheetByName("Plantilla")) === 0 && inmovilizadas(rCurso.ss.getSheetByName("Plantilla")) === 0);

// 14.2. Boletín largo (Ana con las 16 materias de 3B valoradas + la que sólo existe en 1C: 17 materias)
const srcLargo = variante3B(valorar("97000001", Object.fromEntries(estReal.materias.map(m => [m.nombre, "Suficiente"]))));
let rLargo = runTest("Boletín largo individual", { control: ctlAna3B, sources: srcLargo, templateGrid: tpl });
let rLargoCurso = runTest("Boletín largo por curso", {
  control: ctlReal("97000001", { [F3B]: ID_VAR, B9: "EPS2" }), sources: srcLargo, templateGrid: tpl, ejecutar: "generarBoletinesPorCurso()"
});
assert("Impresión, boletín largo: conserva las 17 materias, los cuatro bloques inferiores y la fila 1 inmovilizada",
  !rLargo.error && filasMaterias(rLargo.out).length === 17 && filaPtf(rLargo.out) === 28 + (17 - 8)
    && cajas(rLargo.out) === "Obs 3B Ana|Asist 3B Ana|Conv 3B Ana|PTF 3B Ana"
    && inmovilizadas(rLargo.ss.getSheetByName("Informe_actual")) === 1,
  rLargo.error || filasMaterias(rLargo.out).length + " materias / PTF en la fila " + filaPtf(rLargo.out) + " / " + cajas(rLargo.out));
assert("Impresión, boletín largo: la hoja del libro por curso es idéntica al informe individual y también inmoviliza la fila 1",
  !rLargoCurso.error && JSON.stringify(gridHoja(rLargoCurso, "Ficticia, Ana Real")) === JSON.stringify(rLargo.out)
    && inmovilizadas(libroDe(rLargoCurso).getSheetByName("Ficticia, Ana Real")) === 1,
  rLargoCurso.error || hojasDe(rLargoCurso).join("|"));
assert("Impresión, boletín largo: sin escrituras en las fuentes",
  [rLargo, rLargoCurso].every(r => r.registro.indebidos.length === 0), rLargo.registro.indebidos.concat(rLargoCurso.registro.indebidos).join(", "));

// 14.3. Tanda: todos los boletines con la fila 1 inmovilizada; el Resumen (oculto) no se toca
assert("Impresión, tanda: cada boletín tiene la fila 1 inmovilizada y el Resumen no",
  libroDe(rCurso).getSheets().every(h => inmovilizadas(h) === (h.getName() === "Resumen" ? 0 : 1)),
  libroDe(rCurso).getSheets().map(h => h.getName() + ":" + inmovilizadas(h)).join(", "));

// =============================================================================
// 15. Color comunicacional: identidad de cada nivel de la escala conceptual; gris para lo que todavía no corresponde
// =============================================================================
// Contrato (modelos institucionales de 2024): En proceso, Suficiente y Avanzado tienen cada uno su color; las notas
// numéricas, ACREDITÓ y EQUIVALENCIA son datos neutros. Además hay dos estados estructurales, no académicos: celda
// vacía de un período habilitado y período futuro. Las señales se identifican por la tabla del generador
// (CONFIG_EPS): los tests protegen qué significa cada una y que sean distinguibles y legibles.
const SENALES = CELDAS_CONTROL.SENALES;
const FUTURO = CELDAS_CONTROL.COLOR_PERIODO_FUTURO;
const VACIA = CELDAS_CONTROL.COLOR_CELDA_VACIA;
const LETRA = { enProceso: "P", suficiente: "S", avanzado: "A", neutra: "." };
// Trayectoria de Ana con 16 materias (más de las 8 de la Plantilla): 1B y 3B sobre la grilla real de 3B, 1C sobre la de 1C.
const ID_COLOR_1B = "ID_COLOR_1B_xxxxxxxxxxxxxxxxx";
const ID_COLOR_1C = "ID_COLOR_1C_xxxxxxxxxxxxxxxxx";
const ID_COLOR_3B = "ID_COLOR_3B_xxxxxxxxxxxxxxxxx";
const srcColor = Object.assign({}, sources, {
  [ID_COLOR_1B]: { EPSI: variante(src3B_real.EPSI, valorar("97000001", {
    "Matemática I": "En proceso", "Lengua y Literatura I": "Avanzado", "Taller Profesional I A": "Avanzado",
    "Taller Profesional I B": "Equivalencia", "FGI I": "ACREDITÓ", "Matemática II": "En proceso", "Matemática III": "Suficiente" })) },
  [ID_COLOR_1C]: { EPSI: variante(src1C_real.EPSI, valorar("97000001", {
    "Matemática I": "4", "Lengua y Literatura I": "9", "Taller Profesional I A": "7", "Taller Profesional I B": "EQUIVALENCIA",
    "FGI I": "ACREDITÓ", "Matemática II": "5", "Matemática III": "6", "Matemática IV": "3" })) },
  [ID_COLOR_3B]: { EPSI: variante(src3B_real.EPSI, valorar("97000001", {
    "Matemática I": "Suficiente", "Lengua y Literatura I": "Avanzado", "Taller Profesional I A": "En proceso",
    "Taller Profesional I B": "EQUIVALENCIA", "FGI I": "ACREDITÓ", "Matemática II": "En proceso", "Inglés I": "Suficiente",
    "Taller Profesional II A": "Avanzado", "Taller Profesional II B": "Suficiente", "FGI II": "ACREDITÓ", "Matemática III": "Avanzado",
    "Taller Profesional III A": "Suficiente", "FGI III": "ACREDITÓ", "Matemática IV": "Avanzado", "Taller Profesional IV A": "Suficiente",
    "FGI IV": "En proceso" })) }
});
const ctlColor = (periodo, extra) => ctl(periodo, "97000001", Object.assign({ B9: "EPS2", [F1B]: ID_COLOR_1B, [F1C]: ID_COLOR_1C, [F3B]: ID_COLOR_3B }, extra));
const SEMAFORO_V04 = ["E9:E16 En Proceso", "E9:E16 Suficiente", "E9:E16 Avanzado", "G9:G16 En Proceso", "G9:G16 Suficiente", "G9:G16 Avanzado"];
const runColor = (periodo, ejecutar) => runTest("Color " + periodo, {
  control: ctlColor(periodo), sources: srcColor, templateGrid: tpl, reglasPlantilla: SEMAFORO_V04, ejecutar
});
// Por materia, de E (1B) a L (final): P En proceso, S Suficiente, A Avanzado, . dato neutro, V vacía, F período futuro, ? otra cosa.
const senales = hoja => {
  const g = rect(hoja.grid);
  const out = {};
  for (let r = 9; r <= g.length; r++) {
    const materia = g[r - 1][2];
    if (!materia || /^OBSERVACIONES/.test(materia)) break;
    out[materia] = [5, 6, 7, 8, 9, 10, 11, 12].map(c => {
      const f = hoja.formatos[r + "," + c] || {};
      if (f.fondo === FUTURO && f.texto === SENALES.neutra.texto) return "F";
      if (f.fondo === VACIA && f.texto === SENALES.neutra.texto) return "V";
      const cat = Object.keys(SENALES).find(k => SENALES[k].fondo === f.fondo && SENALES[k].texto === f.texto);
      return cat ? LETRA[cat] : "?";
    }).join("");
  }
  return out;
};
const informeDe = r => r.ss.getSheetByName("Informe_actual");

// 15.1. Gramática visual: identidad propia para cada nivel conceptual, sobria y legible
const rgb = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16));
const luminancia = h => { const [r, g, b] = rgb(h).map(c => c / 255).map(c => c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4); return 0.2126 * r + 0.7152 * g + 0.0722 * b; };
const contraste = (a, b) => (Math.max(luminancia(a), luminancia(b)) + 0.05) / (Math.min(luminancia(a), luminancia(b)) + 0.05);
const tono = h => { const [r, g, b] = rgb(h); if (Math.max(r, g, b) - Math.min(r, g, b) <= 8) return "neutro"; return r > g && r > b ? "rojo" : g > r && g > b ? "verde" : b > r && b > g ? "azul" : "otro"; };
const niveles = ["enProceso", "suficiente", "avanzado"];
assert("Gramática visual: sólo los tres niveles conceptuales tienen color; lo neutro no tiene relleno ni texto de color",
  Object.keys(SENALES).sort().join(",") === "avanzado,enProceso,neutra,suficiente"
    && SENALES.neutra.fondo === null && SENALES.neutra.texto === "#000000");
assert("Gramática visual: cada nivel tiene su propio tono, como en los modelos de 2024 (En proceso rojo, Suficiente verde, Avanzado azul), y el período futuro es gris neutro",
  niveles.map(k => tono(SENALES[k].fondo) + "/" + tono(SENALES[k].texto)).join(",") === "rojo/rojo,verde/verde,azul/azul" && tono(FUTURO) === "neutro"
    && new Set(niveles.map(k => SENALES[k].fondo).concat(FUTURO)).size === 4,
  niveles.map(k => SENALES[k].fondo + "/" + SENALES[k].texto).join(", "));
assert("Gramática visual: rellenos claros (sin saturación que domine la página) y texto oscuro del mismo tono con contraste de al menos 4.5:1",
  niveles.every(k => luminancia(SENALES[k].fondo) >= 0.6 && luminancia(SENALES[k].texto) < luminancia(SENALES[k].fondo)
    && contraste(SENALES[k].fondo, SENALES[k].texto) >= 4.5),
  niveles.map(k => k + ":" + contraste(SENALES[k].fondo, SENALES[k].texto).toFixed(1)).join(", "));

const gris = h => { const [r, g, b] = rgb(h); return 0.299 * r + 0.587 * g + 0.114 * b; };
assert("Estados estructurales: vacía y futuro son grises sin tono, distintos entre sí y de cualquier valoración",
  tono(VACIA) === "neutro" && tono(FUTURO) === "neutro" && VACIA !== FUTURO
    && !niveles.some(k => [VACIA, FUTURO].includes(SENALES[k].fondo)) && SENALES.neutra.fondo === null,
  VACIA + " / " + FUTURO);
assert("Estados estructurales: de menos a más marcado, dato neutro (sin relleno) < vacía < futuro, con diferencias perceptibles en grises",
  255 - gris(VACIA) >= 6 && gris(VACIA) - gris(FUTURO) >= 8, gris(VACIA) + " / " + gris(FUTURO));

// 15.2. 3B (16 materias): trayectorias legibles de izquierda a derecha, también en filas agregadas; H a L en gris
let rColor3B = runColor(P3B);
const s3B = informeDe(rColor3B) ? senales(informeDe(rColor3B)) : {};
const fila3B = m => (s3B[m] || "").slice(0, 3);
assert("Color 3B: genera las 16 materias, más de las 8 filas de la Plantilla",
  !rColor3B.error && Object.keys(s3B).length === 16, rColor3B.error || Object.keys(s3B).join("|"));
assert("Trayectoria que mejora: En proceso → 4 → Suficiente (la nota numérica queda neutra)",
  fila3B("Matemática I") === "P.S", s3B["Matemática I"]);
assert("Trayectoria que sigue en proceso: En proceso → 5 → En proceso",
  fila3B("Matemática II") === "P.P", s3B["Matemática II"]);
assert("Desempeño destacado: Avanzado → 9 → Avanzado (Avanzado tiene identidad propia, distinta de Suficiente)",
  fila3B("Lengua y Literatura I") === "A.A", s3B["Lengua y Literatura I"]);
assert("Trayectoria positiva: Suficiente → 6 → Avanzado",
  fila3B("Matemática III") === "S.A", s3B["Matemática III"]);
assert("Trayectoria que retrocede: Avanzado → 7 → En proceso",
  fila3B("Taller Profesional I A") === "A.P", s3B["Taller Profesional I A"]);
assert("ACREDITÓ y EQUIVALENCIA son datos neutros: la escuela no les definió un color, y no se confunden con vacío",
  fila3B("FGI I") === "..." && fila3B("Taller Profesional I B") === "...", s3B["FGI I"] + " / " + s3B["Taller Profesional I B"]);
assert("Notas numéricas: dato neutro en 1C, sea menor a 6 (3, 4, 5) o mayor o igual (6, 7, 9); sin nota, vacía",
  ["Matemática I", "Matemática II", "Lengua y Literatura I", "Matemática III", "Taller Profesional I A", "Matemática IV"].every(m => s3B[m][1] === ".")
    && ["Inglés I", "FGI II", "FGI IV"].every(m => s3B[m][1] === "V"),
  Object.keys(s3B).map(m => m + ":" + s3B[m][1]).join(", "));
assert("Vacía habilitada: materia que sólo tiene 3B (Inglés I) muestra 1B y 1C vacías, no futuras ni neutras",
  fila3B("Inglés I") === "VVS" && fila3B("FGI II") === "VV.", s3B["Inglés I"] + " / " + s3B["FGI II"]);
assert("Filas agregadas dinámicamente: misma gramática (Matemática IV vacía → 3 → Avanzado; FGI IV vacía, vacía, En proceso)",
  fila3B("Matemática IV") === "V.A" && fila3B("FGI IV") === "VVP" && fila3B("Taller Profesional II A") === "VVA",
  [s3B["Matemática IV"], s3B["FGI IV"], s3B["Taller Profesional II A"]].join(" / "));
assert("Color 3B: E a G con su estado (nivel, dato neutro o vacía, sin gris fijo), H a L futuras en todas las filas, sin celdas sin asignar",
  Object.values(s3B).every(v => /^[PSA.V][.V][PSA.V]F{5}$/.test(v)), JSON.stringify(s3B));
assert("Color: sólo se formatea el área de valoraciones (filas de materias, columnas E a L)",
  Object.keys(informeDe(rColor3B).formatos).every(k => { const [r, c] = k.split(",").map(Number); return r >= 9 && r < 9 + 16 && c >= 5 && c <= 12; })
    && Object.keys(informeDe(rColor3B).formatos).length === 16 * 8,
  Object.keys(informeDe(rColor3B).formatos).length);

// 15.3. El gris depende del período generado
let rColor1C = runColor(P1C);
const s1C = informeDe(rColor1C) ? senales(informeDe(rColor1C)) : {};
assert("Color 1C: G (3B) y siguientes futuras; E con su nivel o vacía y F dato neutro (8 materias, las de la Plantilla)",
  !rColor1C.error && Object.keys(s1C).length === 8 && s1C["Matemática I"] === "P.FFFFFF" && s1C["Lengua y Literatura I"] === "A.FFFFFF"
    && s1C["Matemática III"] === "S.FFFFFF" && s1C["Matemática IV"] === "V.FFFFFF" && s1C["FGI I"] === "..FFFFFF",
  rColor1C.error || JSON.stringify(s1C));
let rColor1B = runColor("1º bimestre");
const s1B = informeDe(rColor1B) ? senales(informeDe(rColor1B)) : {};
assert("Color 1B: F (1C) en adelante futuras; sólo E habilitada (7 materias, menos que la Plantilla)",
  !rColor1B.error && Object.keys(s1B).length === 7 && s1B["Matemática I"] === "PFFFFFFF" && s1B["Lengua y Literatura I"] === "AFFFFFFF"
    && s1B["Matemática III"] === "SFFFFFFF" && s1B["FGI I"] === ".FFFFFFF" && Object.values(s1B).every(v => /^[PSA.]F{7}$/.test(v)),
  rColor1B.error || JSON.stringify(s1B));
assert("Futuro tiene precedencia sobre vacía: en 1B y en 3B ninguna celda futura (todas sin contenido) queda como vacía",
  Object.values(s1B).every(v => v.slice(1) === "FFFFFFF") && Object.values(s3B).every(v => v.slice(3) === "FFFFF"));

// 15.4. Individual y por curso: mismo criterio visual; el semáforo heredado se elimina
let rColorCurso = runColor(P3B, "generarBoletinesPorCurso()");
const hojaCurso = libroDe(rColorCurso) && libroDe(rColorCurso).getSheetByName("Ficticia, Ana Real");
assert("Color: la hoja del libro por curso tiene exactamente los mismos fondos y colores de texto que el informe individual",
  !rColorCurso.error && hojaCurso && JSON.stringify(hojaCurso.formatos) === JSON.stringify(informeDe(rColor3B).formatos),
  rColorCurso.error || hojasDe(rColorCurso).join("|"));
assert("Color: las reglas del semáforo de v0.4 que traiga la Plantilla no llegan al boletín (individual ni curso) y la Plantilla no se toca",
  informeDe(rColor3B).getConditionalFormatRules().length === 0 && hojaCurso.getConditionalFormatRules().length === 0
    && rColor3B.ss.getSheetByName("Plantilla").getConditionalFormatRules().length === 6,
  JSON.stringify(informeDe(rColor3B).getConditionalFormatRules()));
assert("Color: sin escrituras en las fuentes",
  [rColor3B, rColor1C, rColor1B, rColorCurso].every(r => r.registro.indebidos.length === 0));

// 15.5. Categoría de cada valoración
const genColor = cargarGenerador();
const categorias = valores => valores.map(v => LETRA[genColor.categoriaVisual_(v)]).join("");
assert("Categoría: En proceso, Suficiente y Avanzado con identidad propia; ACREDITÓ, EQUIVALENCIA, notas (1 a 10) y vacío neutros",
  categorias(["En proceso", "Suficiente", "Avanzado", "ACREDITÓ", "EQUIVALENCIA", "1", "5", "6", "10", "", undefined]) === "PSA........");

// 15.6. Plantilla versionada: sin semáforo; la leyenda institucional (imagen) no se altera
const xmlPlantilla = xmlHoja(XLSX_GENERADOR, "Plantilla");
assert("Plantilla (xlsx): sin reglas de formato condicional",
  !/conditionalFormatting|cfRule/.test(xmlPlantilla));
// Rellenos de la Plantilla según su estilo de celda (styles.xml del xlsx extraído).
const rellenoPlantilla = (() => {
  const dir = path.join(WORK, path.basename(XLSX_GENERADOR, ".xlsx"));
  const st = fs.readFileSync(path.join(dir, "xl/styles.xml"), "utf8");
  const bloque = t => (st.match(new RegExp("<(?:\\w+:)?" + t + "[^>]*>([\\s\\S]*?)</(?:\\w+:)?" + t + ">")) || [])[1] || "";
  const fills = bloque("fills").match(/<(?:\w+:)?fill>[\s\S]*?<\/(?:\w+:)?fill>|<(?:\w+:)?fill\s*\/>/g) || [];
  const xfs = bloque("cellXfs").match(/<(?:\w+:)?xf [^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?xf>)/g) || [];
  const estilos = Object.fromEntries([...xmlPlantilla.matchAll(/<(?:\w+:)?c r="([A-Z]+\d+)"[^>]*? s="(\d+)"/g)].map(m => [m[1], +m[2]]));
  return ref => {
    const xf = xfs[estilos[ref]] || "";
    const fill = fills[+((xf.match(/fillId="(\d+)"/) || [])[1] || 0)] || "";
    return /patternType="(?!none)/.test(fill) ? ((fill.match(/rgb="([0-9A-F]+)"/i) || [])[1] || "con relleno") : null;
  };
})();
const areaPlantilla = "EFGHIJKL".split("").flatMap(c => [9, 10, 11, 12, 13, 14, 15, 16].map(r => c + r));
assert("Plantilla (xlsx): base uniforme en el área de valoraciones, sin rellenos heredados por columna",
  areaPlantilla.every(ref => rellenoPlantilla(ref) === null), areaPlantilla.filter(ref => rellenoPlantilla(ref) !== null).join(","));
assert("Plantilla (xlsx): encabezado de la tabla y nombres de materia conservan su relleno estructural",
  "CEFGHIJKL".split("").every(c => rellenoPlantilla(c + "8") && rellenoPlantilla(c + "8") === rellenoPlantilla("C8"))
    && [9, 16].every(r => rellenoPlantilla("C" + r) && rellenoPlantilla("C" + r) !== rellenoPlantilla("C8")),
  rellenoPlantilla("C8") + " / " + rellenoPlantilla("C9"));
assert("Plantilla (xlsx): sin línea de leyenda agregada en la fila 17 (D17 vacía) y la fila 17 sigue en 192 pt",
  (rect(tpl)[16] || [])[3] === "" && /<(?:\w+:)?row r="17" ht="192" customHeight="1"/.test(xmlPlantilla),
  (rect(tpl)[16] || [])[3]);

// =============================================================================
// 16. Varios cursos en una sola operación (#20)
// =============================================================================
// Los 8 cursos fijos de CONFIG_EPS.CURSOS sobre las grillas reales sintéticas. 3B (principal):
// - EPSI: EPS1 (Beto listo) y EPS2 como en la sección 8; EPS3 sin estudiantes; EPS4 con Dora lista.
// - EPSG: EPS1 (Gael listo); EPS2 con Hugo en NO; EPS3 con Iris lista; EPS4 sin estudiantes.
// 1C (anterior) como en la sección 8; sin fuente de 1B.
// Agrega al final de la grilla el marcador de un nivel y, a continuación, copias del bloque de un modelo.
const conNivel = (nivel, modelo, nuevos = []) => g => {
  const base = modelo ? g.slice(filaDni(g, modelo), filaDni(g, modelo) + 4).map(f => f.slice()) : [];
  g.push(Object.assign(Array(g[0].length).fill(""), { 0: nivel }));
  nuevos.forEach(n => {
    const bloque = base.map(f => f.slice());
    [bloque[0][0], bloque[0][1], bloque[0][2]] = [n.apellido, n.nombres, n.dni];
    bloque[0][colDe(g, "LISTO PARA EXPORTAR")] = n.listo;
    g.push(...bloque);
  });
};
const EPSI_VARIOS = [conNivel("EPS3"), conNivel("EPS4", "97000001", [{ apellido: "Ficticia", nombres: "Dora Cuatro", dni: "97000041", listo: "SÍ" }])];
const EPSG_VARIOS = [
  conCeldas("96000001", { "LISTO PARA EXPORTAR": "SÍ" }),
  conNivel("EPS2", "96000001", [{ apellido: "Gastro", nombres: "Hugo No", dni: "96000021", listo: "NO" }]),
  conNivel("EPS3", "96000001", [{ apellido: "Gastro", nombres: "Iris Tres", dni: "96000031", listo: "SÍ" }])
];
const conVarios = (epsi = EPSI_VARIOS, epsg = EPSG_VARIOS.concat(conNivel("EPS4"))) => Object.assign({}, srcCurso, {
  [ID_CURSO_3B]: { EPSI: variante(srcCurso[ID_CURSO_3B].EPSI, aplicar(...epsi)), EPSG: variante(src3B_real.EPSG, aplicar(...epsg)) }
});
const srcVarios = conVarios();
const fotoVarios = JSON.stringify(srcVarios);
// Control de varios cursos: hoja origen, DNI y nivel con valores que la operación debe ignorar.
const ctlVarios = marcas => ctl(P3B, "99000001", Object.assign({ B6: "EPSG", B9: "EPS3" }, fuentesCurso, marcas));
const MARCA = Object.fromEntries(CELDAS_CONTROL.CURSOS.map(c => [c.hojaOrigen + " " + c.nivel, c.celda]));
const marcas = (valores) => Object.fromEntries(Object.entries(valores).map(([curso, v]) => [MARCA[curso], v]));
const runVarios = (nombre, control, srcs = srcVarios, ejecutar = "generarBoletinesCursosMarcados()") =>
  runTest(nombre, { control, sources: srcs, templateGrid: tpl, ejecutar });
// Libro completo: nombre y, por hoja, nombre, visibilidad, filas inmovilizadas, valores y colores.
const fotoLibro = libro => libro && JSON.stringify([libro.getName(), libro.getSheets().map(h =>
  [h.getName(), h.isSheetHidden(), h.getFrozenRows(), rect(h.grid), h.formatos])]);
const NOMBRE_LIBRO = curso => "Boletines EPS - " + curso.replace(" ", " - ") + " - 3º bimestre 2026 - <fecha>";
const CON_LIBRO = ["EPSI EPS1", "EPSI EPS2", "EPSI EPS4", "EPSG EPS1", "EPSG EPS3"];

// 16.1. Todos: los 8 cursos sin marcar ninguno; un libro por curso con estudiantes listos
let rTodos = runVarios("Todos los cursos", ctlVarios({}), srcVarios, "generarBoletinesTodosLosCursos()");
const nombresTodos = rTodos.creados.map(l => l.getName());
assert("Todos: sin marcas, termina sin error y crea un libro por cada curso con estudiantes listos, en orden fijo",
  !rTodos.error && JSON.stringify(nombresTodos) === JSON.stringify(CON_LIBRO.map(NOMBRE_LIBRO)),
  rTodos.error || nombresTodos.join(" | "));
assert("Todos: como máximo un libro por curso y ninguno para los cursos sin estudiantes listos (EPSI EPS3, EPSG EPS2, EPSG EPS4)",
  new Set(nombresTodos).size === nombresTodos.length
    && ["EPSI EPS3", "EPSG EPS2", "EPSG EPS4"].every(c => !nombresTodos.includes(NOMBRE_LIBRO(c))));
const porCursoDe = curso => {
  const [hoja, nivel] = curso.split(" ");
  return runCurso("Por curso " + curso, ctlCurso(nivel, { B6: hoja }), srcVarios);
};
const porCurso = Object.fromEntries(CELDAS_CONTROL.CURSOS.map(c => c.hojaOrigen + " " + c.nivel).map(curso => [curso, porCursoDe(curso)]));
const libroTodos = curso => rTodos.creados.find(l => l.getName() === NOMBRE_LIBRO(curso)) || null;
assert("Todos: cada libro es idéntico (hojas, Resumen, valores, colores, visibilidad) al de la generación por curso del mismo curso",
  Object.entries(porCurso).every(([curso, r]) => !r.error && fotoLibro(libroDe(r)) === fotoLibro(libroTodos(curso))),
  Object.entries(porCurso).filter(([curso, r]) => r.error || fotoLibro(libroDe(r)) !== fotoLibro(libroTodos(curso))).map(([c, r]) => c + ":" + (r.error || "distinto")).join(", "));
assert("Todos: el contenido esperado (EPSG EPS3 sólo con Iris; EPSI EPS4 sólo con Dora; EPSG EPS2 omite a Hugo sin libro)",
  libroTodos("EPSG EPS3").getSheets().map(h => h.getName()).join("|") === "Resumen|Gastro, Iris Tres"
    && libroTodos("EPSI EPS4").getSheets().map(h => h.getName()).join("|") === "Resumen|Ficticia, Dora Cuatro"
    && rect(libroTodos("EPSI EPS4").getSheetByName("Ficticia, Dora Cuatro").grid)[5][4] === "Informática 4"
    && porCurso["EPSG EPS2"].rawAlerts[0].startsWith("No se creó ningún libro.\n\nNo hay estudiantes con LISTO PARA EXPORTAR = SÍ en EPSG EPS2 (3º bimestre)."));
assert("Todos: Control resume cursos, libros, cursos sin libro y duración, con una línea por curso",
  /^Cursos: 8 · Libros creados: 5 · Sin libro: 3 · Duración: \d+ s\.$/.test(rTodos.estadoResultado)
    && rTodos.estadoAdvertencias.split("\n").filter(l => !l.startsWith("  ")).join("\n") === CELDAS_CONTROL.CURSOS.map(c => c.hojaOrigen + " " + c.nivel).map(curso =>
      CON_LIBRO.includes(curso)
        ? curso + ": Libro creado: " + NOMBRE_LIBRO(curso) + " · " + porCurso[curso].estadoResultado.split(" · ").slice(1).join(" · ")
        : curso + ": Sin libro: " + porCurso[curso].estadoResultado.replace(/^Sin libro: /, "")).join("\n"),
  rTodos.estadoResultado + "\n" + rTodos.estadoAdvertencias);
assert("Todos: Control detalla las advertencias generales por curso y los estudiantes de los cursos sin libro",
  rTodos.estadoAdvertencias.includes("  EPSG EPS1 · " + ADV_SIN_1B)
    && rTodos.estadoAdvertencias.includes("  EPSG EPS2 · Omitido · fila 10 · Gastro, Hugo No: LISTO PARA EXPORTAR = NO."),
  rTodos.estadoAdvertencias);
assert("Todos: un único diálogo final con una línea por curso y un aviso de progreso por curso",
  rTodos.rawAlerts.length === 1 && rTodos.rawAlerts[0].startsWith(rTodos.estadoResultado + "\n\n- EPSI EPS1: Libro creado: ")
    && JSON.stringify(rTodos.ss.toasts) === JSON.stringify(CELDAS_CONTROL.CURSOS.map((c, i) => "Curso " + (i + 1) + " de 8: " + c.hojaOrigen + " " + c.nivel + ".")),
  rTodos.rawAlerts.join(" / ") + " / " + JSON.stringify(rTodos.ss.toasts));
assert("Todos: cada fuente se abre y cada hoja se lee una sola vez por orientación (no por curso)",
  JSON.stringify(rTodos.registro.aperturas) === JSON.stringify([ID_CURSO_3B, ID_CURSO_1C, ID_CURSO_3B, ID_CURSO_1C])
    && rTodos.registro.lecturas.length === 6,
  JSON.stringify(rTodos.registro));
assert("Todos: sin accesos de escritura a las fuentes; fuentes, Plantilla y hojas del generador intactas",
  rTodos.registro.indebidos.length === 0 && JSON.stringify(srcVarios) === fotoVarios
    && rTodos.ss.getSheets().map(s => s.getName()).join("|") === "Plantilla|Control"
    && JSON.stringify(rect(rTodos.ss.getSheetByName("Plantilla").grid)) === JSON.stringify(rect(tpl)),
  rTodos.registro.indebidos.join(", "));
let rTodosConMarcas = runVarios("Todos con marcas", ctlVarios(marcas({ "EPSI EPS2": "SÍ", "EPSG EPS1": "NO", "EPSG EPS3": "cualquier cosa" })),
  srcVarios, "generarBoletinesTodosLosCursos()");
assert("Todos: ignora las marcas, aun con valores no válidos (mismos 5 libros)",
  !rTodosConMarcas.error && JSON.stringify(rTodosConMarcas.creados.map(fotoLibro)) === JSON.stringify(rTodos.creados.map(fotoLibro)),
  rTodosConMarcas.error);

// 16.2. Cursos marcados: uno, varios, y un curso marcado sin estudiantes listos
let rVarios = runVarios("Cursos marcados", ctlVarios(marcas({ "EPSI EPS1": "NO", "EPSI EPS2": "SÍ", "EPSG EPS3": "Si", "EPSI EPS3": "SÍ" })));
assert("Marcados: sólo los cursos con SÍ ('Si' también), en orden fijo; el marcado sin estudiantes listos no crea libro",
  !rVarios.error && rVarios.creados.map(l => l.getName()).join(" | ") === ["EPSI EPS2", "EPSG EPS3"].map(NOMBRE_LIBRO).join(" | ")
    && /^Cursos: 3 · Libros creados: 2 · Sin libro: 1 · /.test(rVarios.estadoResultado)
    && rVarios.estadoAdvertencias.includes("EPSI EPS3: Sin libro: No hay estudiantes con LISTO PARA EXPORTAR = SÍ en EPSI EPS3 (3º bimestre)."),
  rVarios.error || rVarios.creados.map(l => l.getName()).join(" | ") + " / " + rVarios.estadoResultado);
assert("Marcados: cada libro es idéntico al de la generación por curso (y al de Todos)",
  ["EPSI EPS2", "EPSG EPS3"].every(curso => fotoLibro(rVarios.creados.find(l => l.getName() === NOMBRE_LIBRO(curso))) === fotoLibro(libroDe(porCurso[curso]))));
assert("Marcados: con cursos de una sola orientación (EPSI), no se abre la otra hoja como principal ni se lee dos veces",
  (r => !r.error && r.creados.length === 1 && r.registro.aperturas.length === 2 && r.registro.lecturas.length === 3)(
    runVarios("Un curso marcado", ctlVarios(marcas({ "EPSI EPS2": "SÍ" })))));
let rUno = runVarios("Un curso marcado (EPSG EPS1)", ctlVarios(marcas({ "EPSG EPS1": "SÍ" })));
assert("Marcados: un solo curso produce exactamente el libro de la generación por curso",
  !rUno.error && rUno.creados.length === 1 && fotoLibro(rUno.creados[0]) === fotoLibro(libroDe(porCurso["EPSG EPS1"])),
  rUno.error);

// 16.3. Errores de configuración: sin cursos marcados o marca no admitida; nada se lee ni se crea
let rSinMarcas = runVarios("Sin cursos marcados", ctlVarios(marcas({ "EPSI EPS1": "NO" })));
assert("Sin cursos marcados: error claro que sugiere Todos, sin leer fuentes ni crear libros",
  rSinMarcas.error === "No hay cursos marcados con SÍ en Control!B15:B22. Para generar los 8, usar Informes EPS > Generar boletines de todos los cursos."
    && rSinMarcas.registro.aperturas.length === 0 && rSinMarcas.creados.length === 0
    && rSinMarcas.estadoResultado === "Error: " + rSinMarcas.error && rSinMarcas.rawAlerts[0].startsWith("No se pudieron generar los libros de los cursos."),
  rSinMarcas.error);
let rMarcaInvalida = runVarios("Marca no admitida", ctlVarios(marcas({ "EPSI EPS1": "SÍ", "EPSG EPS2": "x" })));
assert("Marca no admitida: error que nombra el curso y la celda, sin leer fuentes ni crear libros",
  rMarcaInvalida.error === 'La marca de EPSG EPS2 (Control!B20) tiene un valor no admitido: "x". Se esperaba SÍ, NO o vacío.'
    && rMarcaInvalida.registro.aperturas.length === 0 && rMarcaInvalida.creados.length === 0,
  rMarcaInvalida.error);
let rVariosSinFuente = runVarios("Varios sin fuente principal", ctlVarios(Object.assign(marcas({ "EPSI EPS1": "SÍ" }), { [F3B]: "" })));
assert("Varios cursos: la configuración se valida como en la generación por curso (falta la fuente principal)",
  rVariosSinFuente.error === "Falta la URL o ID del archivo de valoraciones del 3º bimestre." && rVariosSinFuente.creados.length === 0,
  rVariosSinFuente.error);

// 16.4. Errores globales: abortan antes de crear el primer libro, aunque falle un curso posterior
let rSinNivelEpsg = runVarios("Todos sin EPS4 en EPSG", ctlVarios({}), conVarios(EPSI_VARIOS, EPSG_VARIOS), "generarBoletinesTodosLosCursos()");
assert("Error global en el último curso (EPSG sin marcador EPS4): ningún libro, tampoco de los cursos anteriores",
  rSinNivelEpsg.error === "No se encontró el nivel EPS4 (marcador en la columna A) en la hoja EPSG de la fuente de 3º bimestre."
    && rSinNivelEpsg.creados.length === 0 && rSinNivelEpsg.estadoResultado === "Error: " + rSinNivelEpsg.error,
  rSinNivelEpsg.error);
let rEpsgSinListo = runVarios("EPSG sin LISTO PARA EXPORTAR", ctlVarios(marcas({ "EPSI EPS2": "SÍ", "EPSG EPS1": "SÍ" })),
  conVarios(EPSI_VARIOS, EPSG_VARIOS.concat(conNivel("EPS4"), conEncabezados({ "LISTO PARA EXPORTAR": "LISTO PARA IMPRIMIR" }))));
assert("Error global de estructura en una orientación (EPSG sin LISTO PARA EXPORTAR): ningún libro",
  rEpsgSinListo.error && rEpsgSinListo.error.startsWith('No se encontró una única columna "LISTO PARA EXPORTAR" después de ORIENTACIÓN en la hoja EPSG')
    && rEpsgSinListo.creados.length === 0,
  rEpsgSinListo.error);

// 16.5. El error de un estudiante no afecta a los demás ni a otros cursos (Dora, EPSI EPS4, también en EPSG EPS3)
const conDoraEnEpsg = conNivel("EPS4");
let rDora = runVarios("Estudiante en ambas orientaciones", ctlVarios({}),
  conVarios(EPSI_VARIOS, EPSG_VARIOS.concat(g => {
    conBloques("96000031", "96000031", [{ apellido: "Ficticia", nombres: "Dora Cuatro", dni: "97000041", listo: "SÍ" }])(g);
    conDoraEnEpsg(g);
  })), "generarBoletinesTodosLosCursos()");
assert("Falla de un estudiante: error en ambos cursos; EPSI EPS4 queda sin libro y EPSG EPS3 se genera con Iris",
  !rDora.error && rDora.creados.length === 4 && !rDora.creados.some(l => l.getName() === NOMBRE_LIBRO("EPSI EPS4"))
    && rDora.creados.find(l => l.getName() === NOMBRE_LIBRO("EPSG EPS3")).getSheets().map(h => h.getName()).join("|") === "Resumen|Gastro, Iris Tres"
    && rDora.estadoAdvertencias.includes("EPSI EPS4: Sin libro: Ningún estudiante seleccionado de EPSI EPS4 (3º bimestre) pudo prepararse.")
    && /\n  EPSI EPS4 · Error · fila \d+ · Ficticia, Dora Cuatro: El DNI aparece en las hojas EPSI y EPSG /.test(rDora.estadoAdvertencias),
  rDora.error || rDora.creados.map(l => l.getName()).join(" | ") + "\n" + rDora.estadoAdvertencias);
assert("Falla de un estudiante: los demás cursos producen los mismos libros que sin la falla",
  ["EPSI EPS1", "EPSI EPS2", "EPSG EPS1"].every(curso => fotoLibro(rDora.creados.find(l => l.getName() === NOMBRE_LIBRO(curso))) === fotoLibro(libroTodos(curso))));

// 16.6. Falla inesperada después de crear algún libro: se informa lo ya creado (no se borra ni se reintenta)
let rCorte = runVarios("Falla al crear el segundo libro", ctlVarios({}), srcVarios,
  "const crear = SpreadsheetApp.create; let n = 0; SpreadsheetApp.create = nombre => { if (++n === 2) throw new Error('cuota de Drive'); return crear(nombre); }; generarBoletinesTodosLosCursos()");
assert("Falla inesperada: error propagado, Control y diálogo listan el libro ya creado",
  rCorte.error === "cuota de Drive" && rCorte.creados.length === 1 && rCorte.estadoResultado === "Error: cuota de Drive"
    && rCorte.estadoAdvertencias.startsWith("Libros ya creados antes del error:\nEPSI EPS1: Libro creado: " + NOMBRE_LIBRO("EPSI EPS1"))
    && rCorte.rawAlerts[0].includes("Libros ya creados antes del error:\n- EPSI EPS1: Libro creado: "),
  rCorte.error + " / " + rCorte.estadoAdvertencias);

// 16.7. Las generaciones existentes no cambian con marcas en Control
let rIndConMarcas = runTest("Individual con marcas", { control: ctlReal("97000001", marcas({ "EPSI EPS1": "SÍ", "EPSG EPS2": "x" })), sources, templateGrid: tpl });
assert("Individual: ignora las marcas de varios cursos (mismo informe, sin libros)",
  !rIndConMarcas.error && mismoInforme(rIndConMarcas, rReal) && rIndConMarcas.creados.length === 0, rIndConMarcas.error);
let rCursoConMarcas = runCurso("Por curso con marcas", ctlCurso("EPS2", marcas({ "EPSI EPS1": "SÍ", "EPSG EPS2": "x" })), srcCurso);
assert("Por curso: ignora las marcas de varios cursos (mismo libro que sin marcas)",
  !rCursoConMarcas.error && rCursoConMarcas.creados.length === 1 && fotoLibro(libroDe(rCursoConMarcas)) === fotoLibro(libroDe(rCurso)),
  rCursoConMarcas.error);

// 16.8. Control del xlsx: bloque de varios cursos con los 8 cursos fijos
const genCursos = cargarGenerador();
assert("Control (xlsx): A14 encabeza el bloque de varios cursos",
  textoXlsx("A14") === "Varios cursos" && textoXlsx("B14") === "Generar", textoXlsx("A14"));
assert("Control (xlsx): cada marca está junto al nombre de su curso y empieza vacía",
  CELDAS_CONTROL.CURSOS.length === 8 && CELDAS_CONTROL.CURSOS.every(c =>
    rotuloDe(c.celda) === genCursos.construirNivelVisual_(c.hojaOrigen, c.nivel, "") + " (" + c.hojaOrigen + " " + c.nivel + ")" && textoXlsx(c.celda) === ""),
  CELDAS_CONTROL.CURSOS.map(c => c.celda + ":" + rotuloDe(c.celda)).join(", "));
let rXlsxVarios = runVarios("Varios cursos con Control del xlsx",
  ctlDesdeXlsx(Object.assign({ B5: P3B }, fuentesCurso, marcas({ "EPSI EPS1": "SÍ", "EPSG EPS3": "SÍ" }))));
assert("Control (xlsx): con las marcas del bloque se generan sólo esos cursos, sin depender de B6 ni B9",
  !rXlsxVarios.error && rXlsxVarios.creados.map(l => l.getName()).join(" | ") === ["EPSI EPS1", "EPSG EPS3"].map(NOMBRE_LIBRO).join(" | "),
  rXlsxVarios.error || rXlsxVarios.creados.map(l => l.getName()).join(" | "));
assert("Control (xlsx): el manual y las reglas se conservan debajo del bloque",
  textoXlsx("A24") === "Manual breve" && textoXlsx("A31") === "Reglas centrales", textoXlsx("A24") + " / " + textoXlsx("A31"));

// =============================================================================
// 17. Textos que no deben cortarse: nombre del estudiante y EQUIVALENCIA
// =============================================================================
// El ancho real del texto no se modela: se protege el mecanismo (celda combinada del nombre y tamaño de EQUIVALENCIA).
const combinadas = [...xmlPlantilla.matchAll(/<(?:\w+:)?mergeCell ref="([^"]+)"/g)].map(m => m[1]);
assert("Plantilla (xlsx): el nombre ocupa E4:L4, sobre celdas vacías; DNI y curso siguen en E5:F5 y E6:F6",
  combinadas.includes("E4:L4") && !combinadas.includes("E4:F4") && combinadas.includes("E5:F5") && combinadas.includes("E6:F6")
    && "GHIJKL".split("").every(c => !rect(tpl)[3][colNum(c) - 1]),
  combinadas.join(", "));
// Celdas con texto de una hoja generada y el tamaño de letra que les asignó el generador.
const tamanos = hoja => rect(hoja.grid).flatMap((f, r) => f.map((v, c) => [v, (hoja.formatos[(r + 1) + "," + (c + 1)] || {}).tamano]))
  .filter(([v, t]) => t !== undefined || v === "EQUIVALENCIA");
const soloEquivalenciaEn7 = hoja => { const t = tamanos(hoja); return t.length > 0 && t.every(([v, n]) => v === "EQUIVALENCIA" && n === 7); };
assert("EQUIVALENCIA: en el informe individual va en 7 y ninguna otra celda cambia de tamaño (Ana: 1C y 3B)",
  soloEquivalenciaEn7(rReal.ss.getSheetByName("Informe_actual")) && tamanos(rReal.ss.getSheetByName("Informe_actual")).length === 2,
  JSON.stringify(tamanos(rReal.ss.getSheetByName("Informe_actual"))));
assert("EQUIVALENCIA: igual en la hoja del libro por curso y en la de varios cursos",
  soloEquivalenciaEn7(libroDe(rCurso).getSheetByName("Ficticia, Ana Real (97000001)"))
    && soloEquivalenciaEn7(libroTodos("EPSI EPS2").getSheetByName("Ficticia, Ana Real (97000001)")));
assert("EQUIVALENCIA: un boletín sin EQUIVALENCIA no cambia ningún tamaño (Beto)",
  tamanos(rNuevo.ss.getSheetByName("Informe_actual")).length === 0, JSON.stringify(tamanos(rNuevo.ss.getSheetByName("Informe_actual"))));

console.log(`\nResultado final: ${fail === 0 ? "TODOS LOS TESTS PASARON" : `${fail} TESTS FALLARON`}`);
process.exit(fail ? 1 : 0);
