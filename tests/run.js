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
    this.sheetId = MockSheet.ultimoId = (MockSheet.ultimoId || 0) + 1; // gid, único como en Google Sheets
    this.name = name;
    this.grid = grid.map(r => r.slice());
    this.reglas = (reglas || []).slice(); // formato condicional
    this.formatos = {};                   // "fila,columna" → { fondo, texto }, sólo lo que asigna el generador
  }
  getName() { return this.name; }
  getSheetId() { return this.sheetId; }
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
      getDisplayValues: () => Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => String(sh.get(r + i, c + j)))),
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
  getId() { return this.id; }
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

// Carpeta de Drive simulada: hijas por nombre (con repetidos, como Drive), papelera y ruta para los tests.
// `creadas` registra la ruta de cada carpeta que crea el generador.
function carpetaDrive(nombre, padre = null, creadas = []) {
  const c = {
    nombre, padre, hijas: [], enPapelera: false,
    getName: () => nombre,
    isTrashed: () => c.enPapelera,
    getFoldersByName: n => { const l = c.hijas.filter(h => h.nombre === n); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; },
    createFolder: n => { const h = carpetaDrive(n, c, creadas); c.hijas.push(h); creadas.push(h.ruta()); return h; },
    ruta: () => (padre ? padre.ruta() + " > " : "") + nombre
  };
  return c;
}

function runTest(testName, { control, sources, templateGrid, ejecutar = "generarInformeActual()", ahora = null, reglasPlantilla = [], prepararDrive = () => {} }) {
  const code = fs.readFileSync(path.join(ROOT, "src/Generador_EPS.gs"), "utf8");
  const plantilla = new MockSheet("Plantilla", templateGrid, reglasPlantilla);
  const ctrl = hojaControl(control);
  const ss = new MockSS([plantilla, ctrl]);
  const rawAlerts = [];
  const creados = [];
  const registro = { aperturas: [], lecturas: [], indebidos: [] };
  const formatos = [];
  // Mi unidad: prepararDrive arma las carpetas que ya existían; después se registra lo que hace el generador.
  const carpetasCreadas = [];
  const miUnidad = carpetaDrive("Mi unidad", null, carpetasCreadas);
  prepararDrive(miUnidad);
  carpetasCreadas.length = 0;
  const accesosDrive = [];
  const ctx = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ({ alert: m => rawAlerts.push(m), createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }),
      openById: abrirFuenteMock(sources, registro),
      // Como en Google: el libro nuevo queda en Mi unidad.
      create: nombre => {
        const libro = new MockSS([new MockSheet("Hoja 1", [])], nombre);
        libro.id = "ID_LIBRO_CREADO_" + (creados.length + 1);
        libro.carpeta = miUnidad;
        creados.push(libro);
        return libro;
      }
    },
    // Al mover, se anota también cuántas hojas tenía el libro: el generador lo mueve recién creado.
    DriveApp: {
      getRootFolder: () => { accesosDrive.push("getRootFolder"); return miUnidad; },
      getFileById: id => {
        accesosDrive.push("getFileById");
        const libro = creados.find(l => l.getId() === id);
        if (!libro) throw new Error("mock: archivo inexistente " + id);
        return {
          moveTo: carpeta => {
            if (!carpeta || !carpeta.ruta) throw new Error("mock: moveTo sin carpeta de Drive");
            libro.carpeta = carpeta;
            libro.hojasAlMover = libro.getSheets().length;
          }
        };
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
    formatos,
    miUnidad,
    carpetasCreadas,
    accesosDrive,
    contexto: ctx
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
genMenu.SpreadsheetApp = { getUi: () => ({ createMenu: () => ({ addItem(t, f) { menu.push(t + "→" + f); return this; }, addSeparator() { menu.push("---"); return this; }, addToUi() {} }) }) };
genMenu.onOpen();
assert("Menú: conserva las generaciones y agrega, separadas, las exportaciones a PDF (un libro, cursos marcados, todos)",
  menu.join(", ") === "Generar informe actual→generarInformeActual, Generar boletines por curso→generarBoletinesPorCurso, "
    + "Generar boletines de los cursos marcados→generarBoletinesCursosMarcados, Generar boletines de todos los cursos→generarBoletinesTodosLosCursos, "
    + "Generar boletines por DNI…→generarBoletinesPorDni, ---, Exportar PDF desde un libro o carpeta…→exportarPdfDesdeDrive, Exportar PDF de los cursos marcados→exportarPdfCursosMarcados, "
    + "Exportar PDF de todos los cursos→exportarPdfTodosLosCursos",
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
assert("Libro por curso: un único libro nuevo, completo, y el diálogo muestra su nombre, su enlace y su carpeta",
  rHora.creados.length === 1 && JSON.stringify(hojasDe(rHora)) === JSON.stringify(HOJAS_EPS2)
    && rHora.rawAlerts[0].startsWith("Libro creado: " + libroDe(rHora).getName() + "\n" + libroDe(rHora).getUrl()
      + "\nCarpeta: Mi unidad > Boletines EPS > 2026 - 3º bimestre\n\n"),
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

// =============================================================================
// 18. Exportación a PDF de un libro ya generado (#21)
// =============================================================================
// Se exporta el libro EPS2 de la sección 8. La URL de exportación y Drive se simulan: el PDF real se valida en Google.
const ID_LIBRO = "ID_LIBRO_EPS2_xxxxxxxxxxxxxxx";
const LINK_LIBRO = "https://docs.google.com/spreadsheets/d/" + ID_LIBRO + "/edit#gid=0";
const blobMock = tipo => ({ nombre: "", getContentType: () => tipo, setName(n) { this.nombre = n; return this; }, getName() { return this.nombre; } });
const respuestaMock = (codigo, tipo) => ({ getResponseCode: () => codigo, getBlob: () => blobMock(tipo) });
const PDF_OK = () => respuestaMock(200, "application/pdf");

// Registra cada método del libro y de sus hojas que usa la exportación.
const conAccesos = (obj, nombre, accesos) => new Proxy(obj, {
  get(t, p) {
    const v = t[p];
    if (typeof v !== "function") return v;
    return (...args) => {
      accesos.add(nombre + "." + String(p));
      const r = v.apply(t, args);
      if (p === "getSheets") return r.map(h => conAccesos(h, "hoja", accesos));
      if (p === "getSheetByName" && r) return conAccesos(r, "hoja", accesos);
      return r;
    };
  }
});

// Ejecuta exportarPdfDesdeDrive() (u otra operación) con los dos archivos de src cargados juntos, como en Apps Script.
// responder(n, url, momento) da la respuesta del pedido n, hecho en ese momento; cada pedido avanza el reloj msPorPdf, y cada Utilities.sleep, lo que espera.
// previas: carpetas que ya están junto al libro ({ nombre, descripcion, archivos, enPapelera, archivosEnPapelera, creada: ms antes }).
// opcion: botón que se elige si se ofrece completar una exportación incompleta. msPrompt: lo que se tarda en pegar el link.
// antes: código que se ejecuta antes (por ejemplo, para cambiar CONFIG_PDF).
// Varios cursos: libros ({ libro, creada: ms antes, enPapelera }) en la carpeta del libro, que es la del período
// (Mi unidad > Boletines EPS > nombrePeriodo); control, valores de Control; ejecutar, la operación; confirmar, el botón.
// Cada libro tiene su última modificación en Drive (modificado: ms antes; por defecto, 2 días antes, anterior a las
// carpetas previas de los tests).
function runExportar({ libro, respuesta = LINK_LIBRO, boton = "OK", responder = PDF_OK, msPorPdf = 0, msPrompt = 0, padres, previas = [],
  opcion = "YES", antes = "", libros = [], control = {}, ejecutar = "exportarPdfDesdeDrive()", confirmar = "OK", msConfirmar = 0, nombrePeriodo = "Carpeta del curso",
  modificado, idCarpeta = "ID_CARPETA_DEL_LIBRO_xxxxxxxxx", cambiarDrive = null, exportacionEnCurso = false }) {
  const ctrl = hojaControl(control);
  const generador = new MockSS([new MockSheet("Plantilla", tpl), ctrl]);
  const alerts = [], preguntas = [], pedidos = [], carpetas = [], abiertos = [], esperas = [], consultasDrive = [], accesos = new Set();
  const dialogos = [], vistas = [], ejecuciones = [];
  let cerrojoTomado = exportacionEnCurso;
  let archivosCreados = 0;
  // Efectos al terminar la vista previa (también si falló): Control, carpetas, archivos, pedidos y accesos a los libros.
  const fotoControl = () => JSON.stringify([ctrl.grid, ctrl.formatos]);
  const controlInicial = fotoControl();
  let trasVista = null;
  let reloj = Date.parse("2026-10-16T12:00:00Z");
  const iterador = lista => { let i = 0; return { hasNext: () => i < lista.length, next: () => lista[i++] }; };
  let carpetasConId = 0;
  const todasCarpetas = [];
  const carpetaMock = (nombre, padre, datos = {}) => {
    const c = {
      id: datos.id || "ID_CARPETA_" + String(++carpetasConId).padStart(4, "0") + "_xxxxxxxxxxxxx", nombre, padre, hijas: [], libros: [], archivos: (datos.archivos || []).slice(), archivosEnPapelera: datos.archivosEnPapelera || [],
      creada: reloj - (datos.creada || 0), enPapelera: !!datos.enPapelera, descripcion: datos.descripcion || "",
      getName: () => nombre,
      getId: () => c.id,
      getUrl: () => "https://drive.google.com/drive/folders/" + encodeURIComponent(nombre),
      isTrashed: () => c.enPapelera,
      getDescription: () => c.descripcion,
      setDescription: d => { c.descripcion = d; return c; },
      getDateCreated: () => new Date(c.creada),
      getFolders: () => iterador(c.hijas),
      getFoldersByName: n => iterador(c.hijas.filter(h => h.nombre === n)),
      getFilesByType: tipo => iterador(c.libros.filter(l => l.getMimeType() === tipo)),
      getFiles: () => iterador(c.archivos.map(n => ({ getName: () => n, isTrashed: () => false }))
        .concat(c.archivosEnPapelera.map(n => ({ getName: () => n, isTrashed: () => true })))),
      createFolder: n => { const nueva = carpetaMock(n, nombre); c.hijas.push(nueva); carpetas.push(nueva); return nueva; },
      createFile: blob => { c.archivos.push(blob.getName()); archivosCreados++; return {}; }
    };
    todasCarpetas.push(c);
    return c;
  };
  const miUnidad = carpetaMock("Mi unidad");
  const boletines = carpetaMock("Boletines EPS", "Mi unidad");
  const delCurso = carpetaMock(nombrePeriodo, "Boletines EPS", { id: idCarpeta });
  miUnidad.hijas.push(boletines);
  boletines.hijas.push(delCurso);
  previas.forEach(p => delCurso.hijas.push(carpetaMock(p.nombre, nombrePeriodo, p)));
  const todosLibros = (libro ? [{ libro, modificado }] : []).concat(libros);
  // Archivos de los libros en la carpeta del período, como los ve DriveApp.
  libros.forEach(l => delCurso.libros.push({
    getId: () => l.libro.getId(), getName: () => l.libro.getName(), isTrashed: () => !!l.enPapelera,
    getDateCreated: () => new Date(reloj - (l.creada || 0)), getMimeType: () => l.tipo || "application/vnd.google-apps.spreadsheet"
  }));
  const listaPadres = padres || [delCurso];
  const ctx = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => generador,
      getUi: () => ({
        ButtonSet: { OK_CANCEL: "OK_CANCEL", YES_NO_CANCEL: "YES_NO_CANCEL" },
        Button: { OK: "OK", CANCEL: "CANCEL", YES: "YES", NO: "NO", CLOSE: "CLOSE" },
        // La exportación desde Drive usa un diálogo HTML: un prompt bloquearía la ejecución mientras se busca el link.
        prompt: () => { throw new Error("mock: no se espera ui.prompt"); },
        showModalDialog: (html, titulo) => dialogos.push({ archivo: html.archivo, titulo }),
        alert: (titulo, mensaje, botones) => {
          if (!botones) return alerts.push(titulo);
          preguntas.push(mensaje);
          if (botones !== "OK_CANCEL") return opcion;
          reloj += msConfirmar;
          return confirmar;
        }
      }),
      // Como Apps Script: sólo abre planillas de Google (no carpetas ni otros archivos). abiertos: sólo las que abrió.
      openById: id => {
        consultasDrive.push("openById " + id);
        const l = todosLibros.find(x => x.libro.getId() === id);
        if (!l || (l.tipo && l.tipo !== "application/vnd.google-apps.spreadsheet")) throw new Error("mock: no es una planilla " + id);
        if (l.inaccesible) throw new Error("No tenés permiso para acceder a este documento.");
        abiertos.push(id);
        return conAccesos(l.libro, "libro", accesos);
      }
    },
    DriveApp: {
      // Como lo documenta Apps Script: getFileById devuelve archivos y getFolderById, carpetas; cada uno rechaza el otro
      // tipo. No se simula que getFileById encuentre carpetas ni dé su tipo (conducta no documentada).
      getFileById: id => {
        consultasDrive.push("getFileById " + id);
        if (todasCarpetas.some(c => c.id === id)) throw new Error("mock: getFileById no acepta carpetas " + id);
        const l = todosLibros.find(x => x.libro.getId() === id);
        if (!l) throw new Error("mock: archivo inexistente " + id);
        if (l.inaccesible) throw new Error("No tenés permiso para acceder a este archivo.");
        return {
          getName: () => l.libro.getName(),
          getParents: () => iterador(listaPadres), getLastUpdated: () => new Date(reloj - (l.modificado === undefined ? 2 * 86400000 : l.modificado))
        };
      },
      getFolderById: id => {
        consultasDrive.push("getFolderById " + id);
        const carpeta = todasCarpetas.find(c => c.id === id);
        if (!carpeta) throw new Error("mock: no es una carpeta " + id);
        return carpeta;
      },
      getRootFolder: () => miUnidad
    },
    MimeType: { GOOGLE_SHEETS: "application/vnd.google-apps.spreadsheet" },
    ScriptApp: { getOAuthToken: () => "TOKEN_FICTICIO" },
    UrlFetchApp: {
      fetch: (url, opciones) => {
        pedidos.push({ url, opciones, momento: reloj });
        reloj += msPorPdf;
        return responder(pedidos.length, url, reloj - msPorPdf);
      }
    },
    Utilities: {
      formatDate: () => "<fecha>",
      sleep: ms => { esperas.push(ms); reloj += ms; }
    },
    HtmlService: {
      createHtmlOutputFromFile: nombre => {
        if (!fs.existsSync(path.join(ROOT, "src", nombre + ".html"))) throw new Error("mock: no existe el HTML " + nombre);
        return { archivo: nombre, setWidth() { return this; }, setHeight() { return this; } };
      }
    },
    // Cerrojo por usuario: una sola exportación confirmada a la vez.
    LockService: {
      getUserLock: () => ({
        tryLock: () => { if (cerrojoTomado) return false; cerrojoTomado = true; return true; },
        releaseLock: () => { cerrojoTomado = false; }
      })
    },
    Date: class extends Date {
      constructor(...args) { if (args.length) super(...args); else super(reloj); }
      static now() { return reloj; }
    }
  };
  vm.createContext(ctx);
  ["src/Generador_EPS.gs", "src/Exportar_PDF.gs"].forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), "utf8"), ctx));
  // Cada llamada al servidor es una ejecución de Apps Script: se anota cuándo empieza.
  const llamar = (nombre, args) => {
    ejecuciones.push({ nombre, inicio: reloj });
    return vm.runInContext(nombre + "(" + args.map(a => JSON.stringify(a)).join(", ") + ")", ctx);
  };
  let error = null;
  try {
    if (ejecutar !== "exportarPdfDesdeDrive()") {
      vm.runInContext(antes + ejecutar, ctx);
    } else {
      // El diálogo, como lo usa una persona: el menú lo abre y termina; mientras se busca y pega el link no corre
      // ninguna ejecución (boton distinto de OK: se cierra sin analizar); "Analizar" y "Exportar" son ejecuciones nuevas.
      vm.runInContext(antes, ctx);
      llamar("exportarPdfDesdeDrive", []);
      if (boton === "OK") {
        reloj += msPrompt;
        let vista;
        try {
          vista = llamar("previsualizarExportacionPdf", [respuesta]);
        } finally {
          trasVista = { controlIntacto: fotoControl() === controlInicial, carpetas: carpetas.length, archivos: archivosCreados,
            pedidos: pedidos.length, accesos: [...accesos].sort() };
        }
        vistas.push(vista);
        const modos = vista.opciones.map(o => o.modo);
        let modo = null;
        if (!modos.length) alerts.push(vista.texto);
        else if (modos.includes("completar")) {
          preguntas.push(vista.texto);
          modo = opcion === "YES" ? "completar" : opcion === "NO" ? "rehacer" : null;
        } else {
          if (modos.includes("carpeta")) preguntas.push(vista.texto);
          modo = confirmar === "OK" ? modos[0] : null;
        }
        if (modo) {
          reloj += msConfirmar;
          if (cambiarDrive) cambiarDrive({ delCurso, libros: todosLibros, carpetaMock });
          alerts.push(llamar("exportarPdfConfirmado", [vista.entrada, modo, vista.firma]).texto);
        }
      }
    }
  } catch (e) {
    error = e.message;
  }
  return {
    error, alerts, preguntas, pedidos, carpetas, abiertos, esperas, consultasDrive, contexto: ctx, accesos: [...accesos].sort(), generador, delCurso,
    dialogos, vistas, ejecuciones, cerrojoTomado, trasVista,
    archivos: carpetas.length ? carpetas[0].archivos : [],
    resultado: ctrl.get(3, 6), advertencias: ctrl.get(4, 6), toasts: generador.toasts || []
  };
}

const libroPdf = libroDe(rCurso);
libroPdf.id = ID_LIBRO;
const fotoLibroPdf = fotoLibro(libroPdf);
const HOJAS_PDF = HOJAS_EPS2.slice(1); // sin Resumen
const CARPETA_PDF = NOMBRE_LIBRO_EPS2 + " - PDF <fecha>";
const parametrosPdf = url => Object.fromEntries(new URL(url).searchParams);
const CONFIG_PDF_TEST = (() => { const c = {}; vm.createContext(c); ["src/Generador_EPS.gs", "src/Exportar_PDF.gs"]
  .forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), "utf8"), c)); return vm.runInContext("CONFIG_PDF", c); })();
const LECTURAS_LIBRO = ["hoja.getName", "hoja.getRange", "hoja.getSheetId", "libro.getId", "libro.getName", "libro.getSheetByName", "libro.getSheets"];
const soloLeyo = r => r.accesos.every(a => LECTURAS_LIBRO.includes(a));

// 18.1. Libro completo: un PDF por hoja de estudiante, con la configuración de impresión validada
let rPdf = runExportar({ libro: libroPdf });
assert("PDF: termina sin error y crea una carpeta nueva junto al libro, con el nombre del libro",
  !rPdf.error && rPdf.carpetas.length === 1 && rPdf.carpetas[0].nombre === CARPETA_PDF && rPdf.carpetas[0].padre === "Carpeta del curso",
  rPdf.error || JSON.stringify(rPdf.carpetas.map(c => [c.nombre, c.padre])));
assert("PDF: un archivo por hoja de estudiante, con el nombre de la hoja, en el orden del libro y sin Resumen",
  JSON.stringify(rPdf.archivos) === JSON.stringify(HOJAS_PDF.map(h => h + ".pdf")), rPdf.archivos.join(" | "));
assert("PDF: cada pedido exporta una sola hoja del libro (su gid) por la URL de exportación",
  rPdf.pedidos.length === HOJAS_PDF.length
    && rPdf.pedidos.every((p, i) => p.url.startsWith("https://docs.google.com/spreadsheets/d/" + ID_LIBRO + "/export?")
      && parametrosPdf(p.url).gid === String(libroPdf.getSheetByName(HOJAS_PDF[i]).getSheetId())),
  rPdf.pedidos.map(p => p.url).join("\n"));
const pPdf = rPdf.pedidos.length ? parametrosPdf(rPdf.pedidos[0].url) : {};
assert("PDF: A4 vertical, ajustado al ancho y con filas inmovilizadas repetidas",
  pPdf.format === "pdf" && pPdf.size === "A4" && pPdf.portrait === "true" && pPdf.fitw === "true" && pPdf.scale === "2" && pPdf.fzr === "true",
  JSON.stringify(pPdf));
assert("PDF: márgenes superior, izquierdo y derecho de 0 cm e inferior de 1 cm (en pulgadas)",
  pPdf.top_margin === "0" && pPdf.left_margin === "0" && pPdf.right_margin === "0" && Math.abs(Number(pPdf.bottom_margin) * 2.54 - 1) < 0.001,
  JSON.stringify(pPdf));
assert("PDF: todos los pedidos llevan la misma configuración completa, sin depender de lo guardado en el diálogo de impresión",
  rPdf.pedidos.every(p => { const q = parametrosPdf(p.url); delete q.gid; const base = Object.assign({}, pPdf); delete base.gid; return JSON.stringify(q) === JSON.stringify(base); })
    && ["gridlines", "printnotes", "printtitle", "sheetnames", "fzc"].every(k => pPdf[k] === "false") && pPdf.pagenum === "UNDEFINED",
  JSON.stringify(pPdf));
assert("PDF: cada pedido usa el token del usuario y no lanza excepciones por códigos HTTP",
  rPdf.pedidos.every(p => p.opciones.headers.Authorization === "Bearer TOKEN_FICTICIO" && p.opciones.muteHttpExceptions === true));
assert("PDF: el libro no se modifica (mismas hojas, valores, formatos, visibilidad y filas inmovilizadas) y sólo se lee",
  fotoLibro(libroPdf) === fotoLibroPdf && soloLeyo(rPdf), rPdf.accesos.join(", "));
assert("PDF: Control registra conteos, duración, libro y carpeta; el diálogo da el enlace a la carpeta",
  rPdf.resultado === "PDF exportados: 4 de 4 · Con error: 0 · Sin exportar: 0 · Reintentos por HTTP 429: 0 · Tiempo de exportación: 9 s (2.3 s por PDF)."
    + " La carpeta tiene 4 de 4 PDF. Libro: " + NOMBRE_LIBRO_EPS2
    + " · Carpeta: https://drive.google.com/drive/folders/" + encodeURIComponent(CARPETA_PDF)
    && rPdf.advertencias === "Sin advertencias"
    && rPdf.alerts.length === 1 && rPdf.alerts[0].includes(CARPETA_PDF + "\nhttps://drive.google.com/drive/folders/"),
  rPdf.resultado + " / " + rPdf.alerts.join(" / "));
assert("PDF: la descripción de la carpeta nueva identifica al libro por su ID",
  rPdf.carpetas.length === 1 && rPdf.carpetas[0].descripcion.includes(ID_LIBRO), rPdf.carpetas.length && rPdf.carpetas[0].descripcion);
assert("PDF: el primer pedido sale enseguida y entre uno y otro se espera la pausa inicial (3 s), sin preguntar nada",
  JSON.stringify(rPdf.esperas) === JSON.stringify([3000, 3000, 3000]) && rPdf.pedidos[0].momento === Date.parse("2026-10-16T12:00:00Z")
    && rPdf.preguntas.length === 0,
  JSON.stringify(rPdf.esperas));
assert("PDF: un aviso por hoja exportada; el generador sólo cambia en el estado de Control",
  rPdf.toasts.length === 4 && rPdf.toasts[0] === "PDF 1 de 4: " + HOJAS_PDF[0] + "."
    && rPdf.generador.getSheets().map(h => h.getName()).join("|") === "Plantilla|Control",
  rPdf.toasts.join(" / "));

// 18.2. Fallas por hoja: se informan, las demás se exportan y el libro queda igual
// porHoja: respuestas de cada hoja (por su posición en HOJAS_PDF), una por intento; sin respuestas, PDF.
const porHoja = respuestas => {
  const intentos = {};
  return (n, url) => {
    const i = HOJAS_PDF.findIndex(h => String(libroPdf.getSheetByName(h).getSheetId()) === parametrosPdf(url).gid);
    const k = intentos[i] = (intentos[i] || 0) + 1;
    const r = (respuestas[i] || [])[k - 1];
    if (r instanceof Error) throw r;
    return r ? r() : PDF_OK();
  };
};
const HTTP_429 = () => respuestaMock(429, "text/html");
let rPdfFallas = runExportar({
  libro: libroPdf,
  responder: porHoja({
    1: [HTTP_429, HTTP_429, HTTP_429, HTTP_429],
    2: [() => respuestaMock(200, "text/html; charset=utf-8")],
    3: [new Error("Tiempo de espera agotado")]
  })
});
assert("PDF con fallas: termina sin error y exporta las hojas que pudo; la que sigue en 429 se pide 4 veces",
  !rPdfFallas.error && JSON.stringify(rPdfFallas.archivos) === JSON.stringify([HOJAS_PDF[0] + ".pdf"]) && rPdfFallas.pedidos.length === 7,
  rPdfFallas.error || rPdfFallas.archivos.join(" | "));
assert("PDF con fallas: Control lista cada hoja fallida con su motivo (429 tras los reintentos, respuesta que no es PDF, excepción)",
  rPdfFallas.resultado.startsWith("PDF exportados: 1 de 4 · Con error: 3 · Sin exportar: 0 · Reintentos por HTTP 429: 3 · ")
    && rPdfFallas.resultado.includes(" La carpeta tiene 1 de 4 PDF.")
    && rPdfFallas.advertencias === [
      "Error · " + HOJAS_PDF[1] + ": Google respondió HTTP 429 (demasiadas exportaciones seguidas). (4 intentos)",
      "Error · " + HOJAS_PDF[2] + ": La respuesta de Google no es un PDF (text/html; charset=utf-8).",
      "Error · " + HOJAS_PDF[3] + ": Tiempo de espera agotado"
    ].join("\n")
    && rPdfFallas.alerts[0].includes("El libro no se modificó. Para completar la carpeta, volver a analizar el mismo libro y elegir Completar"),
  rPdfFallas.resultado + "\n" + rPdfFallas.advertencias);
assert("PDF con fallas: ante cada 429 la pausa se duplica (6, 12, 24 s) y la mayor queda para las hojas siguientes; sólo otros errores no la cambian",
  JSON.stringify(rPdfFallas.esperas) === JSON.stringify([3000, 6000, 12000, 24000, 24000, 24000])
    && rPdfFallas.toasts.includes("Google limitó las exportaciones: se espera 6 s y se reintenta " + HOJAS_PDF[1] + "."),
  JSON.stringify(rPdfFallas.esperas));
assert("PDF con fallas: el libro no se modifica",
  fotoLibro(libroPdf) === fotoLibroPdf && soloLeyo(rPdfFallas), rPdfFallas.accesos.join(", "));
let rPdf429UnaVez = runExportar({ libro: libroPdf, responder: porHoja({ 1: [HTTP_429], 2: [HTTP_429, HTTP_429] }) });
assert("PDF: un 429 se reintenta y la hoja se exporta; cada PDF exportado baja la pausa 3 s, sin bajar de 3",
  !rPdf429UnaVez.error && JSON.stringify(rPdf429UnaVez.archivos) === JSON.stringify(HOJAS_PDF.map(h => h + ".pdf"))
    && rPdf429UnaVez.resultado.startsWith("PDF exportados: 4 de 4 · Con error: 0 · Sin exportar: 0 · Reintentos por HTTP 429: 3 · ")
    && JSON.stringify(rPdf429UnaVez.esperas) === JSON.stringify([3000, 6000, 3000, 6000, 12000, 9000])
    && rPdf429UnaVez.advertencias === "Sin advertencias",
  rPdf429UnaVez.error || rPdf429UnaVez.resultado + " " + JSON.stringify(rPdf429UnaVez.esperas));
let rPdfDeNuevo = runExportar({ libro: libroPdf, previas: [{ nombre: CARPETA_PDF, archivos: HOJAS_PDF.map(h => h + ".pdf") }] });
assert("PDF: si la última exportación está completa, volver a exportar crea una carpeta nueva con los 4 PDF, sin preguntar",
  !rPdfDeNuevo.error && JSON.stringify(rPdfDeNuevo.archivos) === JSON.stringify(rPdf.archivos) && rPdfDeNuevo.carpetas.length === 1
    && rPdfDeNuevo.preguntas.length === 0,
  rPdfDeNuevo.error);

// 18.3. Tiempo máximo, contado desde el inicio de la ejecución: no se empieza otra exportación y las hojas que faltan se informan
let rPdfTiempo = runExportar({ libro: libroPdf, msPorPdf: 3 * 60 * 1000 });
assert("PDF con tiempo máximo: exporta 2 de 4 y las otras quedan sin exportar, informadas en Control",
  !rPdfTiempo.error && rPdfTiempo.archivos.length === 2 && rPdfTiempo.pedidos.length === 2
    && rPdfTiempo.resultado.startsWith("PDF exportados: 2 de 4 · Con error: 0 · Sin exportar: 2 · Reintentos por HTTP 429: 0 · Tiempo de exportación: 363 s (181.5 s por PDF).")
    && rPdfTiempo.advertencias === HOJAS_PDF.slice(2).map(h => "Sin exportar · " + h + ": se alcanzó el tiempo máximo de la ejecución.").join("\n"),
  rPdfTiempo.error || rPdfTiempo.resultado + "\n" + rPdfTiempo.advertencias);
// Desde el diálogo, el tiempo de buscar el link y de revisar la vista previa no se descuenta: la exportación es una
// ejecución nueva. 4 min hasta analizar y 10 min hasta exportar (más que una ejecución entera) + 30 s por PDF: los 4.
let rPdfDialogoLento = runExportar({ libro: libroPdf, msPrompt: 4 * 60 * 1000, msConfirmar: 10 * 60 * 1000, msPorPdf: 30 * 1000 });
const ejecucionDe = (r, nombre) => r.ejecuciones.find(e => e.nombre === nombre) || {};
assert("Diálogo: el tiempo de pegar el link (4 min) y de revisar la vista previa (10 min) no reduce la exportación: 4 de 4",
  !rPdfDialogoLento.error && rPdfDialogoLento.pedidos.length === 4
    && rPdfDialogoLento.resultado.startsWith("PDF exportados: 4 de 4 · Con error: 0 · Sin exportar: 0 · "),
  rPdfDialogoLento.error || rPdfDialogoLento.resultado);
assert("Diálogo: menú, vista previa y exportación son ejecuciones separadas; la exportación empieza después de la espera y sus pedidos entran en sus 5 minutos",
  JSON.stringify(rPdfDialogoLento.ejecuciones.map(e => e.nombre)) === JSON.stringify(["exportarPdfDesdeDrive", "previsualizarExportacionPdf", "exportarPdfConfirmado"])
    && ejecucionDe(rPdfDialogoLento, "previsualizarExportacionPdf").inicio - ejecucionDe(rPdfDialogoLento, "exportarPdfDesdeDrive").inicio === 4 * 60 * 1000
    && ejecucionDe(rPdfDialogoLento, "exportarPdfConfirmado").inicio - ejecucionDe(rPdfDialogoLento, "previsualizarExportacionPdf").inicio === 10 * 60 * 1000
    && rPdfDialogoLento.pedidos.every(p => p.momento - ejecucionDe(rPdfDialogoLento, "exportarPdfConfirmado").inicio < CONFIG_PDF_TEST.TIEMPO_MAXIMO_MS),
  JSON.stringify(rPdfDialogoLento.ejecuciones));

// 18.4. Libros que no son de la generación por curso, cancelación y link inválido: no se crea nada
const sinCambios = r => r.carpetas.length === 0 && r.pedidos.length === 0;
// Un Resumen terminado, como lo deja la generación: título y, debajo, el curso (ver hojasEstudiantesLibro_ y cursoDelLibro_).
const resumenTerminado = (hojaOrigen = "EPSI", nivel = "EPS9", periodo = P3B, ciclo = "2026") => new MockSheet("Resumen",
  [[CELDAS_CONTROL.TITULO_RESUMEN], ["Hoja origen", hojaOrigen], ["Nivel", nivel], ["Período", periodo], ["Ciclo lectivo", ciclo]]);
const otroLibro = (id, hojas) => { const l = new MockSS(hojas.map(n => n === "Resumen" ? resumenTerminado() : new MockSheet(n, [["x"]])), "Otro libro"); l.id = id; return l; };
let rPdfGenerador = runExportar({ libro: otroLibro("ID_GENERADOR_xxxxxxxxxxxxxxxx", ["Control", "Plantilla", "Informe_actual"]),
  respuesta: "ID_GENERADOR_xxxxxxxxxxxxxxxx" });
assert("PDF: un libro sin Resumen (por ejemplo, el generador) se rechaza sin crear carpeta ni exportar",
  rPdfGenerador.error === 'El libro "Otro libro" no tiene la hoja Resumen: no parece un libro de boletines generado por curso.'
    && sinCambios(rPdfGenerador) && rPdfGenerador.resultado === "" && rPdfGenerador.trasVista.controlIntacto,
  rPdfGenerador.error);
let rPdfSoloResumen = runExportar({ libro: otroLibro("ID_SOLO_RESUMEN_xxxxxxxxxxxxx", ["Resumen"]), respuesta: "ID_SOLO_RESUMEN_xxxxxxxxxxxxx" });
assert("PDF: un libro sin hojas de estudiantes se rechaza sin crear carpeta",
  rPdfSoloResumen.error === 'El libro "Otro libro" no tiene hojas de estudiantes.' && sinCambios(rPdfSoloResumen), rPdfSoloResumen.error);
let rPdfCancelado = runExportar({ libro: libroPdf, boton: "CANCEL" });
assert("PDF: cancelar el pedido del link no abre el libro ni registra nada",
  !rPdfCancelado.error && rPdfCancelado.abiertos.length === 0 && sinCambios(rPdfCancelado) && rPdfCancelado.resultado === "" && rPdfCancelado.alerts.length === 0);
let rPdfLinkMalo = runExportar({ libro: libroPdf, respuesta: "boletines del curso" });
assert("PDF: un link no reconocible se informa sin abrir ningún libro",
  /^No pude reconocer el ID/.test(rPdfLinkMalo.error || "") && rPdfLinkMalo.abiertos.length === 0 && sinCambios(rPdfLinkMalo),
  rPdfLinkMalo.error);
let rPdfSinPadre = runExportar({ libro: libroPdf, padres: [] });
assert("PDF: si no se puede ver la carpeta del libro, la carpeta de PDF se crea en Mi unidad",
  !rPdfSinPadre.error && rPdfSinPadre.carpetas.length === 1 && rPdfSinPadre.carpetas[0].padre === "Mi unidad" && rPdfSinPadre.archivos.length === 4,
  rPdfSinPadre.error);

// 18.5. Separación: la generación no usa la exportación ni sus servicios (DriveApp, sólo para guardar sus libros; sección 19)
const codigoGenerador = fs.readFileSync(path.join(ROOT, "src/Generador_EPS.gs"), "utf8");
assert("PDF: Generador_EPS.gs no usa UrlFetchApp ni ScriptApp; sólo nombra la exportación en el menú",
  !/UrlFetchApp|ScriptApp/.test(codigoGenerador) && (codigoGenerador.match(/exportarPdf\w*/g) || []).join(",") === "exportarPdfDesdeDrive,exportarPdfCursosMarcados,exportarPdfTodosLosCursos");
assert("PDF: Generador_EPS.gs usa de DriveApp sólo la raíz (carpeta de los libros) y getFileById (mover el libro nuevo)",
  JSON.stringify((codigoGenerador.match(/DriveApp\.\w+/g) || []).sort()) === JSON.stringify(["DriveApp.getFileById", "DriveApp.getRootFolder"]),
  (codigoGenerador.match(/DriveApp\.\w+/g) || []).join(", "));

// 18.6. Límite de Google en un curso como el de la prueba real (15 hojas)
// Google no publica el límite. Modelo supuesto: acepta hasta 6 pedidos en cualquier ventana de 60 s y rechaza con 429 el
// resto (los rechazados no cuentan). Reproduce lo observado en producción con pedidos seguidos: 6 PDF y 9 rechazos en ~24 s.
const ID_LIBRO_15 = "ID_LIBRO_15_HOJAS_xxxxxxxxxxx";
const HOJAS_15 = Array.from({ length: 15 }, (_, i) => "Ficticio, Estudiante " + String(i + 1).padStart(2, "0"));
const libro15 = new MockSS([resumenTerminado()].concat(HOJAS_15.map(n => new MockSheet(n, [["x"]]))), "Libro sintético de 15");
libro15.id = ID_LIBRO_15;
const limiteVentana = (maximo, ventanaMs) => {
  const aceptados = [];
  return (n, url, momento) => {
    if (aceptados.filter(t => momento - t < ventanaMs).length >= maximo) return HTTP_429();
    aceptados.push(momento);
    return PDF_OK();
  };
};
const runLimite = (opciones) => runExportar(Object.assign({ libro: libro15, respuesta: ID_LIBRO_15, msPorPdf: 1600 }, opciones));
let rSinPausa = runLimite({ responder: limiteVentana(6, 60000), antes: "CONFIG_PDF.PAUSA_MS = 0; CONFIG_PDF.INTENTOS_POR_HOJA = 1; " });
assert("Límite (modelo): sin pausa ni reintentos, como antes, reproduce la prueba real: 6 de 15 y 9 rechazos 429 en 24 s",
  !rSinPausa.error && rSinPausa.archivos.length === 6
    && rSinPausa.resultado.startsWith("PDF exportados: 6 de 15 · Con error: 9 · Sin exportar: 0 · Reintentos por HTTP 429: 0 · Tiempo de exportación: 24 s"),
  rSinPausa.error || rSinPausa.resultado);
let rConPausa = runLimite({ responder: limiteVentana(6, 60000) });
const segundosDe = r => Number((r.resultado.match(/Tiempo de exportación: (\d+) s/) || [])[1]);
assert("Límite (modelo): con pausa y reintentos se exportan los 15, en orden, dentro del tiempo máximo",
  !rConPausa.error && JSON.stringify(rConPausa.archivos) === JSON.stringify(HOJAS_15.map(h => h + ".pdf"))
    && rConPausa.resultado.startsWith("PDF exportados: 15 de 15 · Con error: 0 · Sin exportar: 0 · ")
    && segundosDe(rConPausa) < 300 && rConPausa.advertencias === "Sin advertencias",
  rConPausa.error || rConPausa.resultado);

// 18.7. Recuperación: completar la última exportación incompleta sin repetir lo ya exportado
const CARPETA_PREVIA = NOMBRE_LIBRO_EPS2 + " - PDF 2026-10-16 08:00";
// Las carpetas anteriores del libro llevan la descripción que les puso la exportación (18.1), no una escrita en el test.
const DESCRIPCION_LIBRO_PDF = rPdf.carpetas[0].descripcion;
const previaIncompleta = (extra) => Object.assign({ nombre: CARPETA_PREVIA, descripcion: DESCRIPCION_LIBRO_PDF,
  archivos: [HOJAS_PDF[0] + ".pdf", HOJAS_PDF[2] + ".pdf"], creada: 3600000 }, extra);
const gidsPedidos = r => r.pedidos.map(p => parametrosPdf(p.url).gid);
const gidsDe = hojas => hojas.map(h => String(libroPdf.getSheetByName(h).getSheetId()));
let rCompletar = runExportar({ libro: libroPdf, previas: [previaIncompleta()] });
const previaDe = r => r.delCurso.hijas.find(c => c.nombre === CARPETA_PREVIA);
assert("Completar: se pregunta, con la carpeta y cuántos PDF tiene",
  rCompletar.preguntas.length === 1
    && rCompletar.preguntas[0].startsWith("Libro: " + NOMBRE_LIBRO_EPS2 + " (4 hojas de estudiantes).\n\nLa última exportación de este libro quedó incompleta: " + CARPETA_PREVIA + " tiene 2 de 4 PDF.")
    && rCompletar.preguntas[0].includes("Completar: completar esa carpeta, exportando sólo las 2 hojas que faltan.")
    && JSON.stringify(rCompletar.vistas[0].opciones.map(o => o.modo)) === JSON.stringify(["completar", "rehacer"]),
  rCompletar.preguntas.join(" / "));
assert("Completar (Sí): se piden sólo las hojas que faltan y sus PDF van a esa carpeta, sin crear otra",
  !rCompletar.error && JSON.stringify(gidsPedidos(rCompletar)) === JSON.stringify(gidsDe([HOJAS_PDF[1], HOJAS_PDF[3]]))
    && rCompletar.carpetas.length === 0
    && JSON.stringify(previaDe(rCompletar).archivos) === JSON.stringify([0, 2, 1, 3].map(i => HOJAS_PDF[i] + ".pdf")),
  rCompletar.error || JSON.stringify(previaDe(rCompletar).archivos));
assert("Completar (Sí): Control dice cuántos faltaban y que la carpeta quedó completa, con su enlace; el libro sólo se lee",
  rCompletar.resultado.startsWith("PDF exportados: 2 de 2 que faltaban · Con error: 0 · Sin exportar: 0 · Reintentos por HTTP 429: 0 · ")
    && rCompletar.resultado.includes(" La carpeta tiene 4 de 4 PDF. Libro: " + NOMBRE_LIBRO_EPS2 + " · Carpeta: https://drive.google.com/drive/folders/" + encodeURIComponent(CARPETA_PREVIA))
    && !rCompletar.alerts[0].includes("Para completar") && fotoLibro(libroPdf) === fotoLibroPdf && soloLeyo(rCompletar),
  rCompletar.resultado);
let rNoCompletar = runExportar({ libro: libroPdf, previas: [previaIncompleta()], opcion: "NO" });
assert("Completar (No): todas las hojas en una carpeta nueva; la anterior no se toca",
  !rNoCompletar.error && rNoCompletar.carpetas.length === 1 && rNoCompletar.carpetas[0].nombre === CARPETA_PDF
    && rNoCompletar.archivos.length === 4 && rNoCompletar.pedidos.length === 4 && previaDe(rNoCompletar).archivos.length === 2,
  rNoCompletar.error);
const sinNada = r => !r.error && r.pedidos.length === 0 && r.carpetas.length === 0 && r.resultado === "" && r.alerts.length === 0;
assert("Completar (Cancelar o cerrar el diálogo): no se exporta, no se crea carpeta ni se registra nada",
  ["CANCEL", "CLOSE"].every(o => sinNada(runExportar({ libro: libroPdf, previas: [previaIncompleta()], opcion: o }))));
const preguntaCon = previas => runExportar({ libro: libroPdf, previas, opcion: "NO" }).preguntas.length === 1;
assert("Completar: sólo cuenta la última carpeta de PDF del libro, fuera de la papelera",
  !preguntaCon([previaIncompleta({ creada: 7200000 }), previaIncompleta({ nombre: NOMBRE_LIBRO_EPS2 + " - PDF 2026-10-16 10:00", archivos: HOJAS_PDF.map(h => h + ".pdf"), creada: 60000 })])
    && preguntaCon([previaIncompleta({ creada: 60000 }), previaIncompleta({ nombre: NOMBRE_LIBRO_EPS2 + " - PDF 2026-10-16 10:00", archivos: HOJAS_PDF.map(h => h + ".pdf"), creada: 7200000 })])
    && !preguntaCon([previaIncompleta({ enPapelera: true })]));
assert("Completar: un PDF en la papelera cuenta como faltante",
  preguntaCon([previaIncompleta({ archivos: HOJAS_PDF.slice(1).map(h => h + ".pdf"), archivosEnPapelera: [HOJAS_PDF[0] + ".pdf"] })]));
assert("Completar: el libro se reconoce por la descripción, no por el nombre: una carpeta renombrada se ofrece; una sin descripción, no",
  preguntaCon([previaIncompleta({ nombre: "PDF para revisar" })]) && !preguntaCon([previaIncompleta({ descripcion: "" })]));

// Regresión: dos libros distintos (IDs distintos) con exactamente el mismo nombre, en la misma carpeta del período,
// como dos regeneraciones del mismo curso en el mismo minuto. La exportación incompleta de uno no es del otro.
const ID_GEMELO = "ID_LIBRO_EPS2_GEMELO_xxxxxxxxx";
const gemelo = new MockSS(libroPdf.getSheets().map(h => new MockSheet(h.getName(), h.grid)), libroPdf.getName());
gemelo.id = ID_GEMELO;
const previaDelOtro = previaIncompleta();
let rGemelo = runExportar({ libro: gemelo, respuesta: ID_GEMELO, previas: [previaDelOtro] });
assert("Mismo nombre (regresión): los dos libros tienen IDs distintos y exactamente el mismo nombre",
  gemelo.getId() !== libroPdf.getId() && gemelo.getName() === libroPdf.getName() && CARPETA_PREVIA.startsWith(gemelo.getName() + " - PDF "));
assert("Mismo nombre (regresión): al exportar el otro libro, la carpeta incompleta del primero no se ofrece ni se completa",
  !rGemelo.error && rGemelo.preguntas.length === 0 && rGemelo.carpetas.length === 1 && rGemelo.archivos.length === 4
    && rGemelo.pedidos.every(p => p.url.startsWith("https://docs.google.com/spreadsheets/d/" + ID_GEMELO + "/export?"))
    && JSON.stringify(previaDe(rGemelo).archivos) === JSON.stringify(previaDelOtro.archivos)
    && rGemelo.carpetas[0].descripcion.includes(ID_GEMELO) && !rGemelo.carpetas[0].descripcion.includes(ID_LIBRO),
  rGemelo.error || rGemelo.preguntas.join(" / ") + " / " + JSON.stringify(previaDe(rGemelo).archivos));
assert("Mismo nombre (regresión): la misma carpeta sí se ofrece al exportar el libro al que pertenece",
  runExportar({ libro: libroPdf, previas: [previaIncompleta()], opcion: "CANCEL" }).preguntas.length === 1);

// Recuperación de punta a punta con un límite más estricto que el del modelo: una corrida no alcanza y la siguiente completa
let rCorrida1 = runLimite({ responder: limiteVentana(3, 60000) });
const carpetaCorrida1 = rCorrida1.carpetas[0];
let rCorrida2 = runLimite({ responder: limiteVentana(3, 60000),
  previas: [{ nombre: carpetaCorrida1.nombre, descripcion: carpetaCorrida1.descripcion, archivos: carpetaCorrida1.archivos, creada: 60000 }] });
const faltaban = 15 - carpetaCorrida1.archivos.length;
assert("Recuperación: con un límite de 3 por minuto la primera corrida queda incompleta, lo informa y dice cómo completar",
  !rCorrida1.error && faltaban > 0 && rCorrida1.resultado.includes(" La carpeta tiene " + (15 - faltaban) + " de 15 PDF.")
    && rCorrida1.alerts[0].includes("Para completar la carpeta, volver a analizar el mismo libro y elegir Completar"),
  rCorrida1.error || rCorrida1.resultado);
assert("Recuperación: la segunda corrida pide sólo las hojas que faltaban y la carpeta queda con los 15 PDF",
  !rCorrida2.error && rCorrida2.pedidos.length >= faltaban && rCorrida2.carpetas.length === 0
    && new Set(rCorrida2.pedidos.map(p => parametrosPdf(p.url).gid)).size === faltaban
    && rCorrida2.resultado.startsWith("PDF exportados: " + faltaban + " de " + faltaban + " que faltaban · Con error: 0 · Sin exportar: 0 · ")
    && rCorrida2.resultado.includes(" La carpeta tiene 15 de 15 PDF.")
    && JSON.stringify(rCorrida2.delCurso.hijas[0].archivos.slice().sort()) === JSON.stringify(HOJAS_15.map(h => h + ".pdf")),
  rCorrida2.error || rCorrida2.resultado);

// =============================================================================
// 19. Organización de los libros en Drive (#23)
// =============================================================================
// Drive se simula (carpetas, papelera y moveTo); los permisos y el comportamiento real se validan en Google.
const CARPETA_3B = "Mi unidad > Boletines EPS > 2026 - 3º bimestre";
const rutaLibro = libro => libro.carpeta.ruta();

// 19.1. Por curso: carpeta creada la primera vez, libro movido recién creado y completo en ella
assert("Drive, por curso: el libro queda en Mi unidad > Boletines EPS > <ciclo> - <período>, creadas al no existir",
  !rCurso.error && rutaLibro(libroDe(rCurso)) === CARPETA_3B
    && JSON.stringify(rCurso.carpetasCreadas) === JSON.stringify(["Mi unidad > Boletines EPS", CARPETA_3B]),
  rCurso.error || rutaLibro(libroDe(rCurso)) + " / " + rCurso.carpetasCreadas.join(", "));
assert("Drive, por curso: se mueve recién creado (con su única hoja inicial), antes de copiar los boletines",
  libroDe(rCurso).hojasAlMover === 1 && JSON.stringify(hojasDe(rCurso)) === JSON.stringify(HOJAS_EPS2),
  libroDe(rCurso).hojasAlMover);
assert("Drive, por curso: Control conserva el resultado y el enlace del libro, sin cambios",
  rCurso.estadoResultado === "Libro creado: " + NOMBRE_LIBRO_EPS2 + " · Generados: 4 · Omitidos: 3 · Con error: 0. " + libroDe(rCurso).getUrl(),
  rCurso.estadoResultado);

// 19.2. Carpetas existentes: se reutilizan; una en la papelera se ignora; otro ciclo o período, otra subcarpeta
const conCarpetas = (...rutas) => miUnidad => rutas.forEach(ruta => ruta.split(" > ").reduce((padre, nombre) =>
  padre.getFoldersByName(nombre).hasNext() ? padre.getFoldersByName(nombre).next() : padre.createFolder(nombre), miUnidad));
const runCursoDrive = (nombre, prepararDrive, extra) =>
  runTest(nombre, { control: ctlCurso("EPS2", extra), sources: srcCurso, templateGrid: tpl, ejecutar: "generarBoletinesPorCurso()", prepararDrive });
let rCarpetaExistente = runCursoDrive("Carpeta existente", conCarpetas("Boletines EPS > 2026 - 3º bimestre", "Boletines EPS > 2026 - 1º cuatrimestre"));
assert("Drive: con las carpetas ya creadas, el libro va a la existente y no se crea ninguna",
  !rCarpetaExistente.error && rCarpetaExistente.carpetasCreadas.length === 0 && rutaLibro(libroDe(rCarpetaExistente)) === CARPETA_3B
    && rCarpetaExistente.miUnidad.hijas.length === 1 && rCarpetaExistente.miUnidad.hijas[0].hijas.length === 2,
  rCarpetaExistente.error || rCarpetaExistente.carpetasCreadas.join(", "));
let rEnPapelera = runCursoDrive("Carpeta en la papelera", miUnidad => {
  conCarpetas("Boletines EPS > 2026 - 3º bimestre")(miUnidad);
  miUnidad.hijas[0].hijas[0].enPapelera = true;
});
const subcarpetas3B = r => r.miUnidad.hijas[0].hijas.filter(c => c.nombre === "2026 - 3º bimestre");
assert("Drive: una carpeta con el mismo nombre en la papelera se ignora y se crea otra",
  !rEnPapelera.error && JSON.stringify(rEnPapelera.carpetasCreadas) === JSON.stringify([CARPETA_3B])
    && subcarpetas3B(rEnPapelera).length === 2 && libroDe(rEnPapelera).carpeta === subcarpetas3B(rEnPapelera)[1],
  rEnPapelera.error || rEnPapelera.carpetasCreadas.join(", "));
let rOtroCiclo = runCursoDrive("Otro ciclo", conCarpetas("Boletines EPS > 2026 - 3º bimestre"), { B8: "2027" });
assert("Drive: otro ciclo lectivo usa su propia subcarpeta dentro de Boletines EPS",
  !rOtroCiclo.error && rutaLibro(libroDe(rOtroCiclo)) === "Mi unidad > Boletines EPS > 2027 - 3º bimestre"
    && JSON.stringify(rOtroCiclo.carpetasCreadas) === JSON.stringify(["Mi unidad > Boletines EPS > 2027 - 3º bimestre"]),
  rOtroCiclo.error || rutaLibro(libroDe(rOtroCiclo)));
let rOtroPeriodo = runTest("Otro período", { control: ctl("1º cuatrimestre", "", Object.assign({ B9: "EPS2" }, fuentesCurso)),
  sources: srcCurso, templateGrid: tpl, ejecutar: "generarBoletinesPorCurso()" });
assert("Drive: otro período usa su propia subcarpeta (1º cuatrimestre)",
  !rOtroPeriodo.error && rutaLibro(libroDe(rOtroPeriodo)) === "Mi unidad > Boletines EPS > 2026 - 1º cuatrimestre",
  rOtroPeriodo.error || rutaLibro(libroDe(rOtroPeriodo)));

// 19.3. Varios cursos: una sola búsqueda de la carpeta y todos los libros en ella
assert("Drive, varios cursos: los 5 libros quedan en la misma carpeta, buscada y creada una sola vez",
  !rTodos.error && rTodos.creados.every(l => rutaLibro(l) === CARPETA_3B && l.hojasAlMover === 1)
    && JSON.stringify(rTodos.carpetasCreadas) === JSON.stringify(["Mi unidad > Boletines EPS", CARPETA_3B])
    && rTodos.accesosDrive.filter(a => a === "getRootFolder").length === 1,
  rTodos.error || rTodos.creados.map(rutaLibro).join(", ") + " / " + rTodos.accesosDrive.join(", "));
assert("Drive, varios cursos: el diálogo indica la carpeta de los libros, después de la línea de cada curso",
  rTodos.rawAlerts[0].includes(" s.\n\n- EPSI EPS1: Libro creado: ")
    && rTodos.rawAlerts[0].includes("\n\nCarpeta de los libros: " + CARPETA_3B + ".\n\nEl detalle por estudiante"),
  rTodos.rawAlerts[0]);
let rCorteDrive = runVarios("Falla al crear el segundo libro (Drive)", ctlVarios({}), srcVarios,
  "const crear = SpreadsheetApp.create; let n = 0; SpreadsheetApp.create = nombre => { if (++n === 2) throw new Error('cuota de Drive'); return crear(nombre); }; generarBoletinesTodosLosCursos()");
assert("Drive, varios cursos: si la ejecución falla a mitad, el libro ya creado igual quedó en la carpeta",
  rCorteDrive.error === "cuota de Drive" && rCorteDrive.creados.length === 1 && rutaLibro(rCorteDrive.creados[0]) === CARPETA_3B,
  rCorteDrive.error);

// 19.4. Sin libro, errores previos y generación individual: Drive no se toca
assert("Drive: sin estudiantes listos, con un error global o en la generación individual, no se busca ni crea ninguna carpeta",
  [rNadieListo, rSinListo, rNivelInexistente, rSinNivelEpsg, rSinMarcas, rReal, rSmoke].every(r => r.accesosDrive.length === 0 && r.carpetasCreadas.length === 0),
  [rNadieListo, rSinListo, rNivelInexistente, rSinNivelEpsg, rSinMarcas, rReal, rSmoke].map(r => r.accesosDrive.length).join(","));
let rVariosSinListos = runVarios("Marcados sin estudiantes listos", ctlVarios(marcas({ "EPSI EPS3": "SÍ", "EPSG EPS2": "SÍ" })));
assert("Drive, varios cursos: si ningún curso marcado tiene libro, no se crea la carpeta y el diálogo no la menciona",
  !rVariosSinListos.error && rVariosSinListos.creados.length === 0 && rVariosSinListos.accesosDrive.length === 0
    && !rVariosSinListos.rawAlerts[0].includes("Carpeta"),
  rVariosSinListos.error || rVariosSinListos.rawAlerts[0]);

// 19.5. Fallas de Drive
let rSinCarpeta = runCursoDrive("No se puede crear la carpeta", miUnidad => {
  miUnidad.createFolder = () => { throw new Error("Acceso denegado: DriveApp."); };
});
assert("Drive: si no se puede crear la carpeta, no se crea ningún libro y el error queda en Control",
  rSinCarpeta.error === "Acceso denegado: DriveApp." && rSinCarpeta.creados.length === 0
    && rSinCarpeta.estadoResultado === "Error: Acceso denegado: DriveApp.",
  rSinCarpeta.error);
let rSinMover = runTest("No se puede mover el libro", { control: ctlCurso("EPS2"), sources: srcCurso, templateGrid: tpl,
  ejecutar: "DriveApp.getFileById = () => ({ moveTo: () => { throw new Error('Sin permiso.'); } }); generarBoletinesPorCurso()" });
assert("Drive: si no se puede mover el libro nuevo, la generación se detiene y el error nombra la carpeta y el libro vacío a borrar",
  rSinMover.error === "No se pudo mover el libro nuevo a " + CARPETA_3B + ": Sin permiso. El libro \"" + NOMBRE_LIBRO_EPS2
      + "\" quedó vacío en Mi unidad y se puede borrar: " + libroDe(rSinMover).getUrl()
    && rSinMover.creados.length === 1 && hojasDe(rSinMover).join("|") === "Hoja 1" && rSinMover.estadoResultado === "Error: " + rSinMover.error,
  rSinMover.error);

// 19.6. La organización no cambia el contenido del libro
assert("Drive: el libro organizado es idéntico al de la generación por curso sin carpetas previas",
  fotoLibro(libroDe(rCarpetaExistente)) === fotoLibro(libroDe(rCurso)) && fotoLibro(libroDe(rEnPapelera)) === fotoLibro(libroDe(rCurso)));

// =============================================================================
// 20. Exportación a PDF de varios cursos (#23)
// =============================================================================
// Los libros están en la carpeta del período (Mi unidad > Boletines EPS > 2026 - 3º bimestre). Drive y la URL de
// exportación se simulan; los tiempos y el límite reales se validan en Google.
const PERIODO_PDF = "2026 - 3º bimestre";
const ctlPdf = marcasCtl => Object.assign({ B5: P3B, B8: "2026" }, marcasCtl || {});
// Los 5 libros que generó "todos los cursos" (sección 16), con sus nombres reales y en el orden en que se crearon.
const LIBROS_TODOS = rTodos.creados.map((libro, i) => ({ libro, creada: 600000 - i * 1000 }));
const fotosTodos = rTodos.creados.map(fotoLibro);
const hojasPdfDe = libro => libro.getSheets().map(h => h.getName()).filter(n => n !== "Resumen").map(n => n + ".pdf");
const runPdfCursos = opciones => runExportar(Object.assign({
  libros: LIBROS_TODOS, nombrePeriodo: PERIODO_PDF, control: ctlPdf(), ejecutar: "exportarPdfTodosLosCursos()"
}, opciones));
// Lo que quedó en la carpeta del período al terminar, como carpetas previas de la ejecución siguiente.
const INICIO_RELOJ = Date.parse("2026-10-16T12:00:00Z");
const quedaEnDrive = r => r.delCurso.hijas.map(c => ({
  nombre: c.nombre, descripcion: c.descripcion, archivos: c.archivos.slice(), creada: INICIO_RELOJ - c.creada + 3600000
}));
const carpetaDeLibro = (r, libro) => r.delCurso.hijas.filter(c => c.descripcion.includes(libro.getId()));
const pedidosDeLibro = (r, libro) => r.pedidos.filter(p => p.url.includes("/d/" + libro.getId() + "/export?")).length;

// 20.1. Todos los cursos, sin PDF previos: un libro por curso con libro, todo en una ejecución
let rPdfTodos = runPdfCursos({});
assert("PDF de todos: se confirma una vez, con una línea por curso (los cursos sin libro, informados)",
  rPdfTodos.preguntas.length === 1
    && rPdfTodos.preguntas[0].includes("\n- EPSI EPS1 (libro <fecha>): 1 PDF a exportar.\n- EPSI EPS2 (libro <fecha>): 4 PDF a exportar.\n- EPSI EPS3: sin libro del período.")
    && rPdfTodos.preguntas[0].includes("PDF a exportar: 8. Cada ejecución avanza hasta 5 minutos; si no termina, volver a ejecutar la misma opción"),
  rPdfTodos.preguntas.join(" / "));
assert("PDF de todos: cada libro tiene su carpeta nueva junto a él, con su ID en la descripción y un PDF por hoja",
  !rPdfTodos.error && rPdfTodos.carpetas.length === 5
    && rTodos.creados.every(l => carpetaDeLibro(rPdfTodos, l).length === 1
      && JSON.stringify(carpetaDeLibro(rPdfTodos, l)[0].archivos) === JSON.stringify(hojasPdfDe(l))
      && carpetaDeLibro(rPdfTodos, l)[0].nombre === l.getName() + " - PDF <fecha>"),
  rPdfTodos.error || rPdfTodos.carpetas.map(c => c.nombre + ":" + c.archivos.length).join(", "));
assert("PDF de todos: 8 pedidos (uno por hoja de estudiante), con la configuración validada; la pausa sigue entre un libro y otro",
  rPdfTodos.pedidos.length === 8 && rTodos.creados.every(l => pedidosDeLibro(rPdfTodos, l) === hojasPdfDe(l).length)
    && rPdfTodos.pedidos.every(p => parametrosPdf(p.url).size === "A4" && parametrosPdf(p.url).fzr === "true")
    && JSON.stringify(rPdfTodos.esperas) === JSON.stringify(Array(7).fill(3000)),
  JSON.stringify(rPdfTodos.esperas));
assert("PDF de todos: Control resume cursos completos y PDF, con una línea por curso y el enlace de cada carpeta",
  rPdfTodos.resultado.startsWith("PDF de varios cursos (3º bimestre 2026): cursos completos: 5 de 5 con libro · PDF exportados ahora: 8 · Con error: 0 · Reintentos por HTTP 429: 0 · Tiempo de exportación: ")
    && !rPdfTodos.resultado.includes("Faltan PDF")
    && rPdfTodos.advertencias.split("\n")[1].startsWith("EPSI EPS2 (libro <fecha>): 4 de 4 PDF (4 exportados ahora). https://drive.google.com/drive/folders/")
    && rPdfTodos.advertencias.split("\n")[2] === "EPSI EPS3: sin libro del período."
    && rPdfTodos.alerts.length === 1 && rPdfTodos.alerts[0].includes("Las carpetas de PDF están junto a cada libro, en Mi unidad > Boletines EPS > 2026 - 3º bimestre."),
  rPdfTodos.resultado + "\n" + rPdfTodos.advertencias);
assert("PDF de todos: los libros no se modifican y sólo se leen",
  JSON.stringify(rTodos.creados.map(fotoLibro)) === JSON.stringify(fotosTodos) && soloLeyo(rPdfTodos), rPdfTodos.accesos.join(", "));

// 20.2. Volver a ejecutar con todo completo: no se repite nada ni se pregunta
let rPdfTodosOtraVez = runPdfCursos({ previas: quedaEnDrive(rPdfTodos) });
assert("PDF de todos, otra vez: con todo completo no se pide ningún PDF, no se crean carpetas ni se pide confirmación",
  !rPdfTodosOtraVez.error && rPdfTodosOtraVez.pedidos.length === 0 && rPdfTodosOtraVez.carpetas.length === 0 && rPdfTodosOtraVez.preguntas.length === 0
    && rPdfTodosOtraVez.resultado.startsWith("PDF de varios cursos (3º bimestre 2026): cursos completos: 5 de 5 con libro · PDF exportados ahora: 0 · Con error: 0 ")
    && !/Faltan PDF|no se puede exportar/.test(rPdfTodosOtraVez.resultado)
    && rPdfTodosOtraVez.advertencias.split("\n")[0].startsWith("EPSI EPS1 (libro <fecha>): ya estaba completo (1 de 1 PDF). https://drive.google.com/"),
  rPdfTodosOtraVez.error || rPdfTodosOtraVez.resultado + "\n" + rPdfTodosOtraVez.advertencias);

// 20.3. Cursos marcados
let rPdfMarcados = runPdfCursos({ control: ctlPdf(marcas({ "EPSI EPS2": "SÍ", "EPSG EPS3": "Si", "EPSI EPS1": "NO" })), ejecutar: "exportarPdfCursosMarcados()" });
assert("PDF de marcados: sólo los cursos con SÍ, en orden fijo",
  !rPdfMarcados.error && rPdfMarcados.carpetas.length === 2 && rPdfMarcados.pedidos.length === 5
    && rPdfMarcados.preguntas[0].includes("\n- EPSI EPS2 (libro <fecha>): 4 PDF a exportar.\n- EPSG EPS3 (libro <fecha>): 1 PDF a exportar.\n\n")
    && rPdfMarcados.resultado.includes("cursos completos: 2 de 2 con libro"),
  rPdfMarcados.error || rPdfMarcados.resultado);
let rPdfSinMarcas = runPdfCursos({ control: ctlPdf(), ejecutar: "exportarPdfCursosMarcados()" });
assert("PDF de marcados: sin cursos marcados, error que sugiere exportar todos, sin crear nada",
  rPdfSinMarcas.error === "No hay cursos marcados con SÍ en Control!B15:B22. Para exportar los 8, usar Informes EPS > Exportar PDF de todos los cursos."
    && rPdfSinMarcas.carpetas.length === 0 && rPdfSinMarcas.abiertos.length === 0
    && rPdfSinMarcas.resultado === "Error al exportar PDF de varios cursos: " + rPdfSinMarcas.error,
  rPdfSinMarcas.error);
let rPdfMarcaMala = runPdfCursos({ control: ctlPdf(marcas({ "EPSI EPS2": "x" })), ejecutar: "exportarPdfCursosMarcados()" });
assert("PDF de marcados: una marca no admitida es un error, como en la generación",
  /^La marca de EPSI EPS2 \(Control!B16\) tiene un valor no admitido/.test(rPdfMarcaMala.error || "") && rPdfMarcaMala.carpetas.length === 0,
  rPdfMarcaMala.error);
let rPdfSinPeriodo = runPdfCursos({ control: { B5: "" } });
assert("PDF de varios cursos: sin período en Control, error claro sin abrir libros",
  /^Falta indicar el período en Control!B5/.test(rPdfSinPeriodo.error || "") && rPdfSinPeriodo.abiertos.length === 0, rPdfSinPeriodo.error);

// 20.4. Varios libros del mismo curso: se usa el más reciente
const ID_EPS2 = rTodos.creados[1].getId();
const libroCopia = (original, id, nombre) => { const l = new MockSS(original.getSheets().map(h => new MockSheet(h.getName(), h.grid)), nombre || original.getName()); l.id = id; return l; };
const viejoEps2 = libroCopia(rTodos.creados[1], "ID_LIBRO_EPS2_VIEJO_xxxxxxxxxx", NOMBRE_LIBRO("EPSI EPS2").replace("<fecha>", "2026-10-15 09:00"));
const gemeloEps2 = libroCopia(rTodos.creados[1], "ID_LIBRO_EPS2_GEMELO_NUEVO_xxx");
const enPapeleraEps2 = libroCopia(rTodos.creados[1], "ID_LIBRO_EPS2_PAPELERA_xxxxxx");
const pdfConNombreEps2 = libroCopia(rTodos.creados[1], "ID_PDF_CON_NOMBRE_xxxxxxxxxxx");
let rPdfVarios = runPdfCursos({
  control: ctlPdf(marcas({ "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()",
  libros: [
    { libro: viejoEps2, creada: 86400000 }, { libro: rTodos.creados[1], creada: 600000 },
    { libro: gemeloEps2, creada: 300000 }, // mismo nombre exacto, creado después: es el más reciente
    { libro: enPapeleraEps2, creada: 1000, enPapelera: true }, { libro: pdfConNombreEps2, creada: 1000, tipo: "application/pdf" }
  ]
});
assert("Varios libros del curso: se exporta sólo el creado más recientemente (aun con el mismo nombre), fuera de la papelera y que sea una planilla",
  !rPdfVarios.error && rPdfVarios.carpetas.length === 1 && rPdfVarios.carpetas[0].descripcion.includes("ID_LIBRO_EPS2_GEMELO_NUEVO_xxx")
    && rPdfVarios.pedidos.every(p => p.url.includes("/d/ID_LIBRO_EPS2_GEMELO_NUEVO_xxx/"))
    // Se abren las planillas de la carpeta para reconocerlas por su Resumen; ni la de la papelera ni el PDF.
    && JSON.stringify(rPdfVarios.abiertos) === JSON.stringify(["ID_LIBRO_EPS2_VIEJO_xxxxxxxxxx", ID_EPS2, "ID_LIBRO_EPS2_GEMELO_NUEVO_xxx"]),
  rPdfVarios.error || JSON.stringify(rPdfVarios.abiertos));
assert("Varios libros del curso: la confirmación y Control dicen cuál se usa y cuántos hay",
  rPdfVarios.preguntas[0].includes("- EPSI EPS2 (libro <fecha>, el más reciente de 3): 4 PDF a exportar.")
    && rPdfVarios.advertencias.startsWith("EPSI EPS2 (libro <fecha>, el más reciente de 3): 4 de 4 PDF"),
  rPdfVarios.preguntas.join(" / "));

// 20.5. Una exportación incompleta de un curso se completa sin preguntar aparte; las carpetas de otro libro no cuentan
const libroEps2 = rTodos.creados[1];
const previaEps2 = { nombre: libroEps2.getName() + " - PDF 2026-10-16 08:00", descripcion: "PDF del libro de boletines " + ID_EPS2,
  archivos: hojasPdfDe(libroEps2).slice(0, 2), creada: 3600000 };
const previaDeOtroLibro = { nombre: libroEps2.getName() + " - PDF 2026-10-16 08:30", descripcion: "PDF del libro de boletines OTRO_ID_xxxxxxxxxxxxxx",
  archivos: hojasPdfDe(libroEps2).slice(0, 3), creada: 1800000 };
let rPdfIncompleta = runPdfCursos({ previas: [previaEps2, previaDeOtroLibro] });
const previaEps2Final = rPdfIncompleta.delCurso.hijas.find(c => c.nombre === previaEps2.nombre);
assert("Incompleta: la confirmación anuncia que se completa, y no hay otra pregunta",
  rPdfIncompleta.preguntas.length === 1 && rPdfIncompleta.preguntas[0].includes("- EPSI EPS2 (libro <fecha>): faltan 2 de 4 PDF, se completa su carpeta.")
    && rPdfIncompleta.preguntas[0].includes("PDF a exportar: 6."),
  rPdfIncompleta.preguntas.join(" / "));
assert("Incompleta: se piden sólo las 2 hojas que faltaban y van a esa carpeta; la del otro libro no se toca",
  !rPdfIncompleta.error && pedidosDeLibro(rPdfIncompleta, libroEps2) === 2 && rPdfIncompleta.carpetas.length === 4
    && JSON.stringify(previaEps2Final.archivos.slice().sort()) === JSON.stringify(hojasPdfDe(libroEps2).slice().sort())
    && rPdfIncompleta.delCurso.hijas.find(c => c.nombre === previaDeOtroLibro.nombre).archivos.length === 3
    && rPdfIncompleta.advertencias.split("\n")[1].startsWith("EPSI EPS2 (libro <fecha>): 4 de 4 PDF (2 exportados ahora).")
    && rPdfIncompleta.resultado.includes("cursos completos: 5 de 5 con libro · PDF exportados ahora: 6 "),
  rPdfIncompleta.error || rPdfIncompleta.advertencias);

// 20.6. Corte por tiempo y continuación: tres cursos de 15 hojas no entran en una ejecución
const libroSintetico = (curso, id, n) => {
  const l = new MockSS([resumenTerminado(...curso.split(" "))].concat(Array.from({ length: n }, (_, i) =>
    new MockSheet("Ficticio, " + curso.replace(" ", "") + " Estudiante " + String(i + 1).padStart(2, "0"), [["x"]]))),
    NOMBRE_LIBRO(curso).replace("<fecha>", "2026-10-19 10:00"));
  l.id = id;
  return l;
};
const LIBROS_GRANDES = [["EPSI EPS1", "ID_GRANDE_EPSI_EPS1_xxxxxxxxx"], ["EPSI EPS2", "ID_GRANDE_EPSI_EPS2_xxxxxxxxx"], ["EPSG EPS1", "ID_GRANDE_EPSG_EPS1_xxxxxxxxx"]]
  .map(([curso, id], i) => ({ libro: libroSintetico(curso, id, 15), creada: 600000 - i }));
// 10 s por pedido + 3 s de pausa: unos 23 PDF por ejecución; los 45 necesitan 2 o 3 ejecuciones.
const corridas = [];
let previasGrandes = [];
for (let n = 0; n < 5; n++) {
  const r = runPdfCursos({ libros: LIBROS_GRANDES, msPorPdf: 10000, previas: previasGrandes });
  corridas.push(r);
  previasGrandes = quedaEnDrive(r);
  if (!r.resultado.includes("Faltan PDF")) break;
}
const [corrida1, corrida2] = corridas;
const ultimaCorrida = corridas[corridas.length - 1];
const pedidosTotales = [].concat(...corridas.map(r => r.pedidos.map(p => p.url)));
assert("Continuación: la primera ejecución termina por tiempo, con el tercer libro sin empezar y sin carpeta vacía",
  !corrida1.error && corridas.length > 1 && corrida1.resultado.includes("Faltan PDF: volver a ejecutar la misma opción para continuar.")
    && corrida1.carpetas.length === 2 && carpetaDeLibro(corrida1, LIBROS_GRANDES[2].libro).length === 0
    && corrida1.advertencias.includes("EPSG EPS1 (libro <fecha>): no se empezó por el tiempo máximo; faltan 15 de 15 PDF.")
    && /EPSI EPS2 \(libro <fecha>\): \d+ de 15 PDF \(\d+ exportados ahora, \d+ sin exportar por el tiempo máximo\); faltan \d+\./.test(corrida1.advertencias)
    && corrida1.pedidos.every(p => p.momento - INICIO_RELOJ < 5 * 60 * 1000),
  corrida1.error || corrida1.resultado + "\n" + corrida1.advertencias);
assert("Continuación: la segunda ejecución completa la carpeta del libro a medias y empieza el que faltaba",
  !corrida2.error && /- EPSI EPS2 \(libro <fecha>\): faltan \d+ de 15 PDF, se completa su carpeta\./.test(corrida2.preguntas[0])
    && corrida2.preguntas[0].includes("- EPSI EPS1 (libro <fecha>): completo (15 de 15 PDF), se omite.")
    && corrida2.preguntas[0].includes("- EPSG EPS1 (libro <fecha>): 15 PDF a exportar."),
  corrida2.preguntas.join(" / "));
assert("Continuación: al final, una sola carpeta por libro con sus 15 PDF y ninguna hoja pedida dos veces (45 pedidos en " + corridas.length + " ejecuciones)",
  !ultimaCorrida.error && ultimaCorrida.resultado.includes("cursos completos: 3 de 3 con libro") && !ultimaCorrida.resultado.includes("Faltan PDF")
    && LIBROS_GRANDES.every(l => carpetaDeLibro(ultimaCorrida, l.libro).length === 1
      && JSON.stringify(carpetaDeLibro(ultimaCorrida, l.libro)[0].archivos.slice().sort()) === JSON.stringify(hojasPdfDe(l.libro).slice().sort()))
    && pedidosTotales.length === 45 && new Set(pedidosTotales).size === 45,
  ultimaCorrida.resultado + " / pedidos: " + pedidosTotales.length);

// 20.7. Errores: una hoja que falla se reintenta en la ejecución siguiente; un libro ilegible no frena a los demás
const ID_EPS1 = rTodos.creados[0].getId();
let rPdfFallaHoja = runPdfCursos({ responder: (n, url) => url.includes("/d/" + ID_EPS2 + "/") && parametrosPdf(url).gid === String(libroEps2.getSheets()[2].getSheetId())
  ? respuestaMock(500, "text/html") : PDF_OK() });
const hojaFallida = libroEps2.getSheets()[2].getName();
assert("Error en una hoja: el curso queda incompleto, Control detalla la hoja y pide volver a ejecutar",
  !rPdfFallaHoja.error && rPdfFallaHoja.resultado.includes("cursos completos: 4 de 5 con libro · PDF exportados ahora: 7 · Con error: 1 ")
    && rPdfFallaHoja.resultado.endsWith("Faltan PDF: volver a ejecutar la misma opción para continuar.")
    && rPdfFallaHoja.advertencias.includes("EPSI EPS2 (libro <fecha>): 3 de 4 PDF (3 exportados ahora, 1 con error); faltan 1.")
    && rPdfFallaHoja.advertencias.includes("\n  EPSI EPS2 · Error · " + hojaFallida + ": Google respondió HTTP 500."),
  rPdfFallaHoja.resultado + "\n" + rPdfFallaHoja.advertencias);
let rPdfReintentoHoja = runPdfCursos({ previas: quedaEnDrive(rPdfFallaHoja) });
assert("Error en una hoja: la ejecución siguiente pide sólo esa hoja y completa el curso",
  !rPdfReintentoHoja.error && rPdfReintentoHoja.pedidos.length === 1 && parametrosPdf(rPdfReintentoHoja.pedidos[0].url).gid === String(libroEps2.getSheets()[2].getSheetId())
    && rPdfReintentoHoja.resultado.includes("cursos completos: 5 de 5 con libro"),
  rPdfReintentoHoja.error || rPdfReintentoHoja.resultado);
const sinResumen = new MockSS([new MockSheet("Hoja suelta", [["x"]])], rTodos.creados[0].getName());
sinResumen.id = "ID_SIN_RESUMEN_xxxxxxxxxxxxxx";
let rPdfIlegible = runPdfCursos({ libros: [{ libro: sinResumen, creada: 1000 }].concat(LIBROS_TODOS) });
assert("Libro que no es de la generación (sin Resumen), el más reciente de su curso: cuenta como problema y los demás cursos se exportan",
  !rPdfIlegible.error && rPdfIlegible.carpetas.length === 4 && carpetaDeLibro(rPdfIlegible, sinResumen).length === 0
    && rPdfIlegible.advertencias.startsWith("EPSI EPS1 (libro <fecha>, el más reciente de 2): no se puede exportar: El libro \"" + sinResumen.getName() + "\" no tiene la hoja Resumen")
    && rPdfIlegible.resultado.includes("cursos completos: 4 de 5 con libro · Con libro que no se puede exportar: 1 · PDF exportados ahora: 7 ")
    && rPdfIlegible.resultado.endsWith(" Hay cursos cuyo libro no se puede exportar: ver el detalle, corregirlo y volver a ejecutar.")
    && !rPdfIlegible.resultado.includes("Faltan PDF"),
  rPdfIlegible.error || rPdfIlegible.advertencias);
let rPdfSinCarpetaPeriodo = runPdfCursos({ nombrePeriodo: "2026 - 1º cuatrimestre" });
assert("Sin carpeta del período: ningún curso tiene libro, se informa sin crear ni pedir nada",
  !rPdfSinCarpetaPeriodo.error && rPdfSinCarpetaPeriodo.pedidos.length === 0 && rPdfSinCarpetaPeriodo.carpetas.length === 0
    && rPdfSinCarpetaPeriodo.resultado.startsWith("PDF de varios cursos (3º bimestre 2026): cursos completos: 0 de 0 con libro · ")
    && rPdfSinCarpetaPeriodo.resultado.endsWith(" Ningún curso elegido tiene libro en Mi unidad > Boletines EPS > 2026 - 3º bimestre.")
    && rPdfSinCarpetaPeriodo.preguntas.length === 0
    && rPdfSinCarpetaPeriodo.advertencias.split("\n").every(l => l.endsWith(": sin libro del período.")),
  rPdfSinCarpetaPeriodo.error || rPdfSinCarpetaPeriodo.advertencias);
let rPdfCancelaCursos = runPdfCursos({ confirmar: "CANCEL" });
assert("Cancelar la confirmación: no se pide ni crea nada, y Control no cambia",
  !rPdfCancelaCursos.error && rPdfCancelaCursos.pedidos.length === 0 && rPdfCancelaCursos.carpetas.length === 0
    && rPdfCancelaCursos.resultado === "" && rPdfCancelaCursos.alerts.length === 0);
let rPdfConfirmaTarde = runPdfCursos({ msConfirmar: 5 * 60 * 1000 + 1000 });
assert("Confirmación abierta más de 5 minutos: no se empieza ningún libro ni se crea ninguna carpeta; se pide volver a ejecutar",
  !rPdfConfirmaTarde.error && rPdfConfirmaTarde.pedidos.length === 0 && rPdfConfirmaTarde.carpetas.length === 0
    && rPdfConfirmaTarde.resultado.includes("cursos completos: 0 de 5 con libro") && rPdfConfirmaTarde.resultado.endsWith("para continuar."),
  rPdfConfirmaTarde.error || rPdfConfirmaTarde.resultado);

// 20.8. Ritmo compartido: tras dos 429 en un libro, el siguiente empieza con la pausa ajustada, no sin pausa ni con 3 s
let rPdfRitmo = runPdfCursos({ control: ctlPdf(marcas({ "EPSI EPS1": "SÍ", "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()",
  responder: (() => { let rechazos = 0; return (n, url) => url.includes("/d/" + ID_EPS1 + "/") && rechazos++ < 2 ? HTTP_429() : PDF_OK(); })() });
assert("Ritmo compartido: dos 429 en el primer libro (6 y 12 s); el PDF baja la pausa a 9 s y el segundo libro empieza con 9 s",
  !rPdfRitmo.error && JSON.stringify(rPdfRitmo.esperas) === JSON.stringify([6000, 12000, 9000, 6000, 3000, 3000])
    && rPdfRitmo.resultado.includes("Reintentos por HTTP 429: 2"),
  rPdfRitmo.error || JSON.stringify(rPdfRitmo.esperas));

// 20.9. Libro inválido sin nada más para exportar: no se informa como "sin pendientes"
let rPdfSoloProblema = runPdfCursos({ previas: quedaEnDrive(rPdfTodos), libros: [{ libro: sinResumen, creada: 1000 }].concat(LIBROS_TODOS) });
assert("Libro inválido y los demás completos: sin confirmación ni pedidos, pero el resultado lo marca como problema pendiente",
  !rPdfSoloProblema.error && rPdfSoloProblema.pedidos.length === 0 && rPdfSoloProblema.preguntas.length === 0
    && rPdfSoloProblema.resultado.startsWith("PDF de varios cursos (3º bimestre 2026): cursos completos: 4 de 5 con libro · Con libro que no se puede exportar: 1 · PDF exportados ahora: 0 ")
    && rPdfSoloProblema.resultado.endsWith("Hay cursos cuyo libro no se puede exportar: ver el detalle, corregirlo y volver a ejecutar.")
    && rPdfSoloProblema.alerts[0].startsWith(rPdfSoloProblema.resultado),
  rPdfSoloProblema.error || rPdfSoloProblema.resultado);

// 20.10. Generación cortada (regresión): el libro más reciente del curso tiene Resumen pero la generación no terminó
// Con el generador real: se corta antes de escribir Resumen (como un corte por tiempo después de copiar las hojas).
let rCortada = runTest("Generación cortada", { control: ctlCurso("EPS2"), sources: srcCurso, templateGrid: tpl,
  ejecutar: "escribirResumen_ = () => { throw new Error('corte simulado'); }; generarBoletinesPorCurso()" });
const libroCortado = rCortada.creados[0];
libroCortado.id = "ID_LIBRO_CORTADO_xxxxxxxxxxxxx";
assert("Generación cortada: el generador real deja un libro con Resumen vacío y sus hojas de estudiantes",
  rCortada.error === "corte simulado" && libroCortado.getSheetByName("Resumen") && rect(libroCortado.getSheetByName("Resumen").grid).length === 0
    && libroCortado.getSheets().length === HOJAS_EPS2.length,
  rCortada.error);
let rPdfCortadoIndividual = runExportar({ libro: libroCortado, respuesta: libroCortado.getId() });
assert("Generación cortada: la exportación de un libro lo rechaza sin crear carpeta ni pedir PDF",
  rPdfCortadoIndividual.error === "El libro \"" + libroCortado.getName() + "\" no terminó de generarse (su hoja Resumen está incompleta): "
      + "volver a generar el curso y mandar ese libro a la papelera."
    && rPdfCortadoIndividual.carpetas.length === 0 && rPdfCortadoIndividual.pedidos.length === 0,
  rPdfCortadoIndividual.error);
// Estado descrito en la revisión: Resumen (vacío) y sólo algunas hojas de estudiantes, más reciente que un libro completo.
const cortadoParcial = new MockSS([new MockSheet("Resumen", [])].concat(libroEps2.getSheets().slice(1, 3).map(h => new MockSheet(h.getName(), h.grid))), libroEps2.getName());
cortadoParcial.id = "ID_LIBRO_CORTADO_PARCIAL_xxxxx";
let rPdfCortado = runPdfCursos({ control: ctlPdf(marcas({ "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()",
  libros: [{ libro: libroEps2, creada: 600000 }, { libro: cortadoParcial, creada: 1000 }] });
assert("Generación cortada (varios cursos): el libro más reciente queda como problema; no se exportan sus pocas hojas ni se usa en silencio el anterior",
  !rPdfCortado.error && rPdfCortado.pedidos.length === 0 && rPdfCortado.carpetas.length === 0
    && JSON.stringify(rPdfCortado.abiertos) === JSON.stringify([libroEps2.getId(), cortadoParcial.getId()])
    && rPdfCortado.advertencias.startsWith("EPSI EPS2 (libro <fecha>, el más reciente de 2): no se puede exportar: El libro \"" + libroEps2.getName() + "\" no terminó de generarse")
    && rPdfCortado.resultado.includes("cursos completos: 0 de 1 con libro · Con libro que no se puede exportar: 1 "),
  rPdfCortado.error || rPdfCortado.resultado + "\n" + rPdfCortado.advertencias);

// 20.11. Libro modificado después de exportar: la carpeta anterior no está vigente
// La carpeta previa se creó hace 1 h; el libro se modificó hace 30 min (getLastUpdated).
const conModificado = (ms) => LIBROS_TODOS.map(l => l.libro === libroEps2 ? Object.assign({}, l, { modificado: ms }) : l);
const previaCompletaEps2 = Object.assign({}, previaEps2, { archivos: hojasPdfDe(libroEps2) });
let rPdfModParcial = runPdfCursos({ control: ctlPdf(marcas({ "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()",
  libros: conModificado(30 * 60000), previas: [previaEps2] });
const nuevaEps2 = r => r.carpetas.find(c => c.descripcion.includes(ID_EPS2));
assert("Modificado tras una exportación parcial: no se completa; la confirmación lo dice y se exporta todo en una carpeta nueva",
  !rPdfModParcial.error && rPdfModParcial.preguntas[0].includes("- EPSI EPS2 (libro <fecha>): cambió después de su última exportación ("
      + previaEps2.nombre + "), se exportan sus 4 PDF en una carpeta nueva.")
    && rPdfModParcial.carpetas.length === 1 && JSON.stringify(nuevaEps2(rPdfModParcial).archivos) === JSON.stringify(hojasPdfDe(libroEps2))
    && rPdfModParcial.pedidos.length === 4
    && rPdfModParcial.delCurso.hijas.find(c => c.nombre === previaEps2.nombre).archivos.length === 2,
  rPdfModParcial.error || rPdfModParcial.preguntas.join(" / "));
let rPdfModCompleta = runPdfCursos({ control: ctlPdf(marcas({ "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()",
  libros: conModificado(30 * 60000), previas: [previaCompletaEps2] });
assert("Modificado tras una exportación completa: no se omite como vigente; se exporta todo en una carpeta nueva",
  !rPdfModCompleta.error && rPdfModCompleta.preguntas.length === 1
    && rPdfModCompleta.preguntas[0].includes("cambió después de su última exportación")
    && rPdfModCompleta.carpetas.length === 1 && nuevaEps2(rPdfModCompleta).archivos.length === 4 && rPdfModCompleta.pedidos.length === 4
    && rPdfModCompleta.resultado.includes("cursos completos: 1 de 1 con libro"),
  rPdfModCompleta.error || rPdfModCompleta.resultado);
let rPdfNoModCompleta = runPdfCursos({ control: ctlPdf(marcas({ "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()",
  libros: conModificado(2 * 3600000), previas: [previaCompletaEps2] });
assert("Modificado antes de la exportación completa: sigue vigente y se omite",
  !rPdfNoModCompleta.error && rPdfNoModCompleta.pedidos.length === 0 && rPdfNoModCompleta.carpetas.length === 0
    && rPdfNoModCompleta.advertencias.startsWith("EPSI EPS2 (libro <fecha>): ya estaba completo (4 de 4 PDF)."),
  rPdfNoModCompleta.error || rPdfNoModCompleta.advertencias);
// quedaEnDrive pasa la ejecución anterior 1 h al pasado: la modificación del libro (30 min antes de ella) también.
let rPdfTrasModificar = runPdfCursos({ control: ctlPdf(marcas({ "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()",
  libros: conModificado(90 * 60000), previas: quedaEnDrive(rPdfModParcial) });
assert("Modificado: la carpeta nueva, creada después del cambio, queda vigente y la ejecución siguiente la da por completa",
  !rPdfTrasModificar.error && rPdfTrasModificar.pedidos.length === 0 && rPdfTrasModificar.carpetas.length === 0
    && rPdfTrasModificar.resultado.includes("cursos completos: 1 de 1 con libro"),
  rPdfTrasModificar.error || rPdfTrasModificar.resultado);
let rPdfModIndividual = runExportar({ libro: libroPdf, previas: [previaIncompleta()], modificado: 30 * 60000 });
assert("Modificado (exportación de un libro): no se ofrece completar la carpeta anterior; se exporta todo en una nueva",
  !rPdfModIndividual.error && rPdfModIndividual.preguntas.length === 0 && rPdfModIndividual.carpetas.length === 1
    && rPdfModIndividual.archivos.length === 4 && previaDe(rPdfModIndividual).archivos.length === 2,
  rPdfModIndividual.error || rPdfModIndividual.preguntas.join(" / "));
assert("Sin modificar (exportación de un libro): se sigue ofreciendo completar, y No sigue rehaciendo todo en una carpeta nueva",
  runExportar({ libro: libroPdf, previas: [previaIncompleta()], modificado: 2 * 3600000, opcion: "CANCEL" }).preguntas.length === 1
    && (r => !r.error && r.carpetas.length === 1 && r.archivos.length === 4)(runExportar({ libro: libroPdf, previas: [previaIncompleta()], opcion: "NO" })));

// =============================================================================
// 21. Exportar PDF desde un libro o carpeta de Drive
// =============================================================================
// Una carpeta cualquiera, elegida a mano, sin Control: el curso, el período y el ciclo salen del Resumen de cada libro.
const ID_CARPETA_MANUAL = "ID_CARPETA_ELEGIDA_A_MANO_xxxx";
const URL_CARPETA = "https://drive.google.com/drive/folders/" + ID_CARPETA_MANUAL + "?usp=sharing";
const runCarpeta = opciones => runExportar(Object.assign({
  libros: LIBROS_TODOS, nombrePeriodo: "Boletines para revisar", idCarpeta: ID_CARPETA_MANUAL, respuesta: URL_CARPETA, control: {}
}, opciones));
const descripcionDe = libro => "PDF del libro de boletines " + libro.getId();
const previaCompletaDe = (libro, extra) => Object.assign({ nombre: libro.getName() + " - PDF 2026-10-16 08:00", descripcion: descripcionDe(libro),
  archivos: hojasPdfDe(libro), creada: 3600000 }, extra);
const intro = r => (r.preguntas[0] || "").split("\n\n- ")[0];

// 21.1. Entrada libro: por cualquier forma de link o ID se exporta exactamente ese libro (el resto, en la sección 18)
let rLibroOpenId = runExportar({ libro: libroPdf, respuesta: "https://drive.google.com/open?id=" + ID_LIBRO });
assert("Desde Drive, libro: un link open?id= se reconoce como libro y exporta sus 4 PDF junto a él, sin confirmación de carpeta",
  !rLibroOpenId.error && rLibroOpenId.archivos.length === 4 && rLibroOpenId.preguntas.length === 0
    && rLibroOpenId.pedidos.every(p => p.url.includes("/d/" + ID_LIBRO + "/")),
  rLibroOpenId.error);
let rLibroCompletoOtraVez = runExportar({ libro: libroPdf, previas: [previaCompletaDe(libroPdf)] });
assert("Desde Drive, libro: con su exportación completa y vigente, elegirlo a mano rehace todo en una carpeta nueva (deliberado)",
  !rLibroCompletoOtraVez.error && rLibroCompletoOtraVez.carpetas.length === 1 && rLibroCompletoOtraVez.archivos.length === 4
    && rLibroCompletoOtraVez.preguntas.length === 0,
  rLibroCompletoOtraVez.error);

// 21.2. Entrada carpeta con varios cursos, sin depender de Control
let rCarpeta = runCarpeta({});
assert("Desde Drive, carpeta: el link de carpeta se reconoce y la confirmación resume libros, cursos y qué se hará",
  !rCarpeta.error && rCarpeta.preguntas.length === 1
    && intro(rCarpeta) === "Carpeta: Boletines para revisar.\nLibros de boletines: 5, de 5 cursos. Se usa el más reciente de cada curso; se omiten 0 regeneraciones anteriores.\n"
      + "Completos, se omiten: 0 · Se completan: 0 · Se rehacen porque el libro cambió: 0 · Se exportan enteros: 5 · Con problema: 0."
    && rCarpeta.preguntas[0].includes("\n- EPSI EPS1 · 3º bimestre 2026 (libro <fecha>): 1 PDF a exportar.\n- EPSI EPS2 · 3º bimestre 2026 (libro <fecha>): 4 PDF a exportar.")
    && rCarpeta.preguntas[0].includes("PDF a exportar: 8. Cada ejecución avanza hasta 5 minutos; si no termina, volver a analizar y exportar la misma carpeta"),
  rCarpeta.error || rCarpeta.preguntas.join(" / "));
assert("Desde Drive, carpeta: sin período ni cursos en Control, cada libro tiene su carpeta de PDF en la carpeta elegida, con su ID",
  rTodos.creados.every(l => carpetaDeLibro(rCarpeta, l).length === 1 && JSON.stringify(carpetaDeLibro(rCarpeta, l)[0].archivos) === JSON.stringify(hojasPdfDe(l)))
    && rCarpeta.carpetas.every(c => c.padre === "Boletines para revisar") && rCarpeta.pedidos.length === 8
    && rCarpeta.resultado.startsWith("PDF de la carpeta \"Boletines para revisar\": cursos completos: 5 de 5 con libro · PDF exportados ahora: 8 · Con error: 0 ")
    && rCarpeta.alerts[0].includes("Las carpetas de PDF están junto a cada libro, en la carpeta \"Boletines para revisar\".")
    && JSON.stringify(rTodos.creados.map(fotoLibro)) === JSON.stringify(fotosTodos) && soloLeyo(rCarpeta),
  rCarpeta.resultado);
assert("Desde Drive, carpeta: también con el link /drive/u/0/folders/ o con el ID solo",
  ["https://drive.google.com/drive/u/0/folders/" + ID_CARPETA_MANUAL, ID_CARPETA_MANUAL].every(respuesta =>
    (r => !r.error && r.preguntas.length === 1 && r.preguntas[0].startsWith("Carpeta: Boletines para revisar."))(runCarpeta({ respuesta, confirmar: "CANCEL" }))));

// 21.3. Regeneraciones, otro período, planillas que no son libros y archivos que no son planillas
const libro1C = libroCopia(rTodos.creados[0], "ID_LIBRO_EPS1_1C_xxxxxxxxxxxxx"); // mismo nombre que el de 3B: decide el Resumen
libro1C.getSheetByName("Resumen").set(4, 2, "1º cuatrimestre");
const planillaAjena = new MockSS([new MockSheet("Asistencia", [["x"]])], "Planilla de asistencia sintética");
planillaAjena.id = "ID_PLANILLA_AJENA_xxxxxxxxxxxx";
const pdfSuelto = libroCopia(rTodos.creados[0], "ID_PDF_SUELTO_xxxxxxxxxxxxxxxx");
let rMezcla = runCarpeta({ libros: LIBROS_TODOS.concat([
  { libro: viejoEps2, creada: 86400000 }, { libro: libro1C, creada: 1000 },
  { libro: planillaAjena, creada: 1000 }, { libro: pdfSuelto, creada: 1000, tipo: "application/pdf" }
]) });
assert("Carpeta con mezcla: agrupa por el curso del Resumen (no por el nombre), usa el más reciente e ignora lo que no es un libro",
  !rMezcla.error && intro(rMezcla).startsWith("Carpeta: Boletines para revisar.\nLibros de boletines: 7, de 6 cursos. Se usa el más reciente de cada curso; "
      + "se omiten 1 regeneraciones anteriores. Otras planillas, que no son libros de boletines: 1, se ignoran.")
    && rMezcla.preguntas[0].includes("\n- EPSI EPS1 · 1º cuatrimestre 2026 (libro <fecha>): 1 PDF a exportar.\n- EPSI EPS1 · 3º bimestre 2026 (libro <fecha>): 1 PDF a exportar.")
    && rMezcla.preguntas[0].includes("- EPSI EPS2 · 3º bimestre 2026 (libro <fecha>, el más reciente de 2): 4 PDF a exportar.")
    && carpetaDeLibro(rMezcla, viejoEps2).length === 0 && carpetaDeLibro(rMezcla, libro1C).length === 1
    && carpetaDeLibro(rMezcla, planillaAjena).length === 0 && !rMezcla.abiertos.includes(pdfSuelto.getId())
    && rMezcla.resultado.includes("cursos completos: 6 de 6 con libro"),
  rMezcla.error || rMezcla.preguntas.join(" / "));

// 21.4. Libro más reciente incompleto o que no se puede abrir: problema informado, sin usar el anterior
const sinAcceso = libroCopia(rTodos.creados[4], "ID_SIN_ACCESO_xxxxxxxxxxxxxxx", "Planilla compartida sintética");
let rCarpetaCortada = runCarpeta({ libros: [{ libro: libroEps2, creada: 600000 }, { libro: cortadoParcial, creada: 1000 },
  { libro: rTodos.creados[0], creada: 500000 }, { libro: sinAcceso, creada: 1000, inaccesible: true }] });
assert("Carpeta con el libro más reciente cortado: el curso queda con problema y no se exporta el anterior; los demás sí",
  !rCarpetaCortada.error && carpetaDeLibro(rCarpetaCortada, libroEps2).length === 0 && carpetaDeLibro(rCarpetaCortada, rTodos.creados[0]).length === 1
    && rCarpetaCortada.preguntas[0].includes("- EPSI EPS2 · 3º bimestre 2026 (libro <fecha>, el más reciente de 2): no se puede exportar: El libro \"" + libroEps2.getName() + "\" no terminó de generarse")
    && rCarpetaCortada.resultado.includes("cursos completos: 1 de 3 con libro · Con libro que no se puede exportar: 2 "),
  rCarpetaCortada.error || rCarpetaCortada.resultado);
assert("Carpeta con una planilla que no se puede abrir: se informa como problema, no se ignora en silencio",
  rCarpetaCortada.preguntas[0].includes("- Planilla \"Planilla compartida sintética\" (libro <fecha>): no se puede exportar: no se pudo abrir: No tenés permiso")
    && rCarpetaCortada.advertencias.includes("Planilla \"Planilla compartida sintética\""),
  rCarpetaCortada.preguntas.join(" / "));

// 21.5. Carpeta sin libros de boletines
let rCarpetaVacia = runCarpeta({ libros: [{ libro: planillaAjena, creada: 1000 }, { libro: pdfSuelto, creada: 1000, tipo: "application/pdf" }] });
assert("Carpeta sin libros de boletines: la vista previa lo dice, sin opciones de exportar, pedidos, carpetas ni escritura en Control",
  !rCarpetaVacia.error && rCarpetaVacia.preguntas.length === 0 && rCarpetaVacia.pedidos.length === 0 && rCarpetaVacia.carpetas.length === 0
    && rCarpetaVacia.vistas[0].opciones.length === 0
    && rCarpetaVacia.vistas[0].texto.startsWith("Carpeta: Boletines para revisar.\nLibros de boletines: 0, de 0 cursos.")
    && rCarpetaVacia.vistas[0].texto.endsWith("\n\nNo hay PDF para exportar: La carpeta no tiene libros de boletines.")
    && rCarpetaVacia.resultado === "" && rCarpetaVacia.trasVista.controlIntacto,
  rCarpetaVacia.error || rCarpetaVacia.vistas.map(v => v.texto).join(" / "));

// 21.6. Exportaciones previas completas, incompletas y de un libro modificado después
const libroEps4 = rTodos.creados[2];
let rCarpetaPrevias = runCarpeta({
  libros: LIBROS_TODOS.map(l => l.libro === libroEps4 ? Object.assign({}, l, { modificado: 30 * 60000 }) : l),
  previas: [previaCompletaDe(rTodos.creados[0]), previaEps2, previaCompletaDe(libroEps4)]
});
assert("Carpeta con exportaciones previas: completa vigente se omite, incompleta se completa, la del libro modificado se rehace",
  !rCarpetaPrevias.error && intro(rCarpetaPrevias).endsWith("Completos, se omiten: 1 · Se completan: 1 · Se rehacen porque el libro cambió: 1 · Se exportan enteros: 2 · Con problema: 0.")
    && pedidosDeLibro(rCarpetaPrevias, rTodos.creados[0]) === 0 && pedidosDeLibro(rCarpetaPrevias, libroEps2) === 2
    && pedidosDeLibro(rCarpetaPrevias, libroEps4) === 1 && carpetaDeLibro(rCarpetaPrevias, libroEps4).length === 2
    && rCarpetaPrevias.resultado.includes("cursos completos: 5 de 5 con libro · PDF exportados ahora: 5 "),
  rCarpetaPrevias.error || intro(rCarpetaPrevias) + "\n" + rCarpetaPrevias.resultado);

// 21.7. Corte por tiempo y continuación, sin repetir pedidos
const corridasCarpeta = [];
let previasCarpeta = [];
for (let n = 0; n < 5; n++) {
  const r = runCarpeta({ libros: LIBROS_GRANDES, msPorPdf: 10000, previas: previasCarpeta });
  corridasCarpeta.push(r);
  previasCarpeta = quedaEnDrive(r);
  if (!r.resultado.includes("Faltan PDF")) break;
}
const pedidosCarpeta = [].concat(...corridasCarpeta.map(r => r.pedidos.map(p => p.url)));
assert("Carpeta, continuación: la primera ejecución pide volver a exportar la misma carpeta; al final, 45 PDF sin pedidos repetidos y una carpeta por libro",
  corridasCarpeta.length > 1 && corridasCarpeta[0].resultado.endsWith("Faltan PDF: volver a analizar y exportar la misma carpeta para continuar.")
    && corridasCarpeta[1].preguntas[0].includes("se completa su carpeta.")
    && pedidosCarpeta.length === 45 && new Set(pedidosCarpeta).size === 45
    && LIBROS_GRANDES.every(l => carpetaDeLibro(corridasCarpeta[corridasCarpeta.length - 1], l.libro).length === 1)
    && corridasCarpeta[corridasCarpeta.length - 1].resultado.includes("cursos completos: 3 de 3 con libro"),
  corridasCarpeta.map(r => r.resultado).join("\n"));

// 21.8. Coexistencia con las opciones desde Control: las dos ven el mismo progreso en Drive
const runPeriodoComoCarpeta = opciones => runExportar(Object.assign({ libros: LIBROS_TODOS, nombrePeriodo: PERIODO_PDF,
  idCarpeta: ID_CARPETA_MANUAL, respuesta: URL_CARPETA, control: ctlPdf() }, opciones));
let rCarpetaTrasControl = runPeriodoComoCarpeta({ previas: quedaEnDrive(rPdfTodos) });
assert("Coexistencia: después de exportar todos los cursos desde Control, la misma carpeta desde Drive no tiene nada pendiente",
  !rCarpetaTrasControl.error && rCarpetaTrasControl.pedidos.length === 0 && rCarpetaTrasControl.carpetas.length === 0
    && rCarpetaTrasControl.vistas[0].opciones.length === 0
    && rCarpetaTrasControl.vistas[0].texto.includes("Completos, se omiten: 5 · Se completan: 0 · ")
    && rCarpetaTrasControl.vistas[0].texto.endsWith("\n\nNo hay PDF para exportar.") && rCarpetaTrasControl.trasVista.controlIntacto,
  rCarpetaTrasControl.error || rCarpetaTrasControl.vistas.map(v => v.texto).join(" / "));
let rCarpetaDelPeriodo = runPeriodoComoCarpeta({});
let rControlTrasCarpeta = runPdfCursos({ previas: quedaEnDrive(rCarpetaDelPeriodo) });
assert("Coexistencia: después de exportar la carpeta del período desde Drive, todos los cursos desde Control no tiene nada pendiente",
  !rCarpetaDelPeriodo.error && rCarpetaDelPeriodo.pedidos.length === 8
    && !rControlTrasCarpeta.error && rControlTrasCarpeta.pedidos.length === 0 && rControlTrasCarpeta.carpetas.length === 0
    && rControlTrasCarpeta.resultado.includes("cursos completos: 5 de 5 con libro"),
  rControlTrasCarpeta.error || rControlTrasCarpeta.resultado);

// 21.9. Entradas que no son un libro ni una carpeta
let rEsPdf = runCarpeta({ libros: [{ libro: pdfSuelto, tipo: "application/pdf" }], respuesta: pdfSuelto.getId() });
assert("Desde Drive: el ID de un archivo que no es planilla ni carpeta da un error claro, sin abrir ni crear nada",
  rEsPdf.error === "\"" + pdfSuelto.getName() + "\" no es un libro de Google Sheets ni una carpeta de Drive."
    && rEsPdf.abiertos.length === 0 && rEsPdf.carpetas.length === 0 && rEsPdf.resultado === "" && rEsPdf.trasVista.controlIntacto,
  rEsPdf.error);
let rPdfPorLink = runCarpeta({ libros: [{ libro: pdfSuelto, tipo: "application/pdf" }], respuesta: "https://drive.google.com/file/d/" + pdfSuelto.getId() + "/view" });
assert("Desde Drive: el link /file/d/ de un archivo que no es planilla da el mismo error, sin buscarlo como carpeta",
  rPdfPorLink.error === rEsPdf.error && !rPdfPorLink.consultasDrive.some(c => c.startsWith("getFolderById")),
  rPdfPorLink.error);
let rNoExiste = runCarpeta({ respuesta: "ID_QUE_NO_EXISTE_xxxxxxxxxxxx" });
assert("Desde Drive: un ID que no está en Drive da un error claro",
  rNoExiste.error === "No se encontró en Drive un libro ni una carpeta con ese link o ID, o no hay acceso." && rNoExiste.carpetas.length === 0,
  rNoExiste.error);
let rCarpetaNoExiste = runCarpeta({ respuesta: "https://drive.google.com/drive/folders/" + ID_LIBRO });
assert("Desde Drive: un link de carpeta con el ID de un libro no lo abre como libro: error de carpeta no encontrada",
  rCarpetaNoExiste.error === "No se encontró en Drive una carpeta con ese link o ID, o no hay acceso." && rCarpetaNoExiste.abiertos.length === 0,
  rCarpetaNoExiste.error);

// 21.10. Regresión: getFileById no acepta IDs de carpetas (como en Apps Script) y la entrada por carpeta funciona igual
const getFileByIdDeCarpeta = r => vm.runInContext("(() => { try { DriveApp.getFileById('" + ID_CARPETA_MANUAL + "'); return 'aceptado'; } "
  + "catch (e) { return 'rechazado'; } })()", r.contexto);
// Las consultas de cada ejecución se copian antes de sondear el mock (la sonda también queda anotada).
const entradasCarpeta = [URL_CARPETA, ID_CARPETA_MANUAL, "https://drive.google.com/open?id=" + ID_CARPETA_MANUAL]
  .map(respuesta => runCarpeta({ respuesta })).map(r => Object.assign(r, { consultas: r.consultasDrive.slice() }));
assert("Regresión: el mock de getFileById rechaza el ID de una carpeta, y aun así la carpeta se exporta por link, por ID solo y por ?id=",
  entradasCarpeta.every(r => getFileByIdDeCarpeta(r) === "rechazado" && !r.error && r.pedidos.length === 8 && r.carpetas.length === 5
    && r.consultas.includes("getFolderById " + ID_CARPETA_MANUAL)
    && !r.consultas.includes("getFileById " + ID_CARPETA_MANUAL)),
  entradasCarpeta.map(r => r.error || r.consultas.filter(c => c.includes(ID_CARPETA_MANUAL)).join(",")).join(" / "));
// La entrada se resuelve dos veces: en la vista previa y, de nuevo, al exportar.
assert("Regresión: un link de carpeta va directo a getFolderById; un ID solo se prueba primero como planilla y después como carpeta",
  JSON.stringify(entradasCarpeta[0].consultas.filter(c => c.includes(ID_CARPETA_MANUAL))) === JSON.stringify(Array(2).fill("getFolderById " + ID_CARPETA_MANUAL))
    && JSON.stringify(entradasCarpeta[1].consultas.filter(c => c.includes(ID_CARPETA_MANUAL)))
      === JSON.stringify(["openById " + ID_CARPETA_MANUAL, "getFolderById " + ID_CARPETA_MANUAL, "openById " + ID_CARPETA_MANUAL, "getFolderById " + ID_CARPETA_MANUAL]),
  entradasCarpeta.map(r => r.consultas.filter(c => c.includes(ID_CARPETA_MANUAL)).join(",")).join(" / "));
assert("Regresión: un libro por link /d/ o por ID solo se abre como planilla, sin pasar por getFolderById",
  [LINK_LIBRO, ID_LIBRO].every(respuesta => (r => !r.error && r.archivos.length === 4 && !r.consultasDrive.some(c => c.startsWith("getFolderById")))(
    runExportar({ libro: libroPdf, respuesta }))));

// 21.11. Regresión: una planilla que no puede abrirse, con el nombre que le da la generación, participa de la elección
// del más reciente de su curso. Mismo nombre que el libro de EPSI EPS2; sin acceso (openById falla).
const eps2SinAcceso = libroCopia(libroEps2, "ID_EPS2_SIN_ACCESO_xxxxxxxxxxx");
const problemaSinAcceso = "no se puede exportar: no se pudo abrir: No tenés permiso para acceder a este documento.";
const sinAccesoMasReciente = [{ libro: libroEps2, creada: 600000 }, { libro: eps2SinAcceso, creada: 1000, inaccesible: true }];
let rSinAccesoCarpeta = runCarpeta({ libros: sinAccesoMasReciente });
let rSinAccesoControl = runPdfCursos({ libros: sinAccesoMasReciente, control: ctlPdf(marcas({ "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()" });
assert("Regresión, sin acceso más reciente (carpeta): el curso queda con problema y no se exporta el libro válido anterior",
  !rSinAccesoCarpeta.error && rSinAccesoCarpeta.pedidos.length === 0 && rSinAccesoCarpeta.carpetas.length === 0
    && rSinAccesoCarpeta.vistas[0].opciones.length === 0
    && rSinAccesoCarpeta.vistas[0].texto.includes("· Con problema: 1.\n\n- EPSI EPS2 · 3º bimestre 2026 (libro <fecha>, el más reciente de 2): " + problemaSinAcceso)
    && rSinAccesoCarpeta.trasVista.controlIntacto,
  rSinAccesoCarpeta.error || rSinAccesoCarpeta.vistas.map(v => v.texto).join(" / "));
assert("Regresión, sin acceso más reciente (Control): el curso marcado queda con problema y no se exporta el libro válido anterior",
  !rSinAccesoControl.error && rSinAccesoControl.pedidos.length === 0 && rSinAccesoControl.carpetas.length === 0
    && rSinAccesoControl.advertencias.startsWith("EPSI EPS2 (libro <fecha>, el más reciente de 2): " + problemaSinAcceso)
    && rSinAccesoControl.resultado.includes("cursos completos: 0 de 1 con libro · Con libro que no se puede exportar: 1 "),
  rSinAccesoControl.error || rSinAccesoControl.resultado + "\n" + rSinAccesoControl.advertencias);
const sinAccesoAnterior = [{ libro: eps2SinAcceso, creada: 600000, inaccesible: true }, { libro: libroEps2, creada: 1000 }];
let rSinAccesoViejoCarpeta = runCarpeta({ libros: sinAccesoAnterior });
let rSinAccesoViejoControl = runPdfCursos({ libros: sinAccesoAnterior, control: ctlPdf(marcas({ "EPSI EPS2": "SÍ" })), ejecutar: "exportarPdfCursosMarcados()" });
assert("Regresión, sin acceso anterior: el libro válido más reciente se usa y se exporta, desde la carpeta y desde Control",
  [rSinAccesoViejoCarpeta, rSinAccesoViejoControl].every(r => !r.error && carpetaDeLibro(r, libroEps2).length === 1
    && pedidosDeLibro(r, libroEps2) === 4 && r.resultado.includes("cursos completos: 1 de 1 con libro · PDF exportados ahora: 4 ")
    && !r.resultado.includes("no se puede exportar") && r.advertencias.includes("(libro <fecha>, el más reciente de 2): 4 de 4 PDF")),
  [rSinAccesoViejoCarpeta, rSinAccesoViejoControl].map(r => r.error || r.resultado).join(" / "));
// Sin acceso y con un nombre que no identifica un curso: problema propio en la carpeta; desde Control no se atribuye a ningún curso.
const ajenaSinAcceso = [{ libro: sinAcceso, creada: 1000, inaccesible: true }];
let rAjenaCarpeta = runCarpeta({ libros: LIBROS_TODOS.concat(ajenaSinAcceso) });
let rAjenaControl = runPdfCursos({ libros: LIBROS_TODOS.concat(ajenaSinAcceso) });
assert("Regresión, sin acceso y nombre irreconocible: en la carpeta se informa como problema propio, sin afectar a ningún curso",
  !rAjenaCarpeta.error && rAjenaCarpeta.pedidos.length === 8
    && rAjenaCarpeta.advertencias.includes("Planilla \"Planilla compartida sintética\" (libro <fecha>): no se puede exportar: no se pudo abrir:")
    && rAjenaCarpeta.resultado.includes("cursos completos: 5 de 6 con libro · Con libro que no se puede exportar: 1 "),
  rAjenaCarpeta.error || rAjenaCarpeta.resultado);
assert("Regresión, sin acceso y nombre irreconocible: desde Control no se atribuye a ningún curso y los 5 se exportan completos",
  !rAjenaControl.error && rAjenaControl.pedidos.length === 8 && !rAjenaControl.advertencias.includes("Planilla compartida sintética")
    && rAjenaControl.resultado.includes("cursos completos: 5 de 5 con libro") && !rAjenaControl.resultado.includes("no se puede exportar"),
  rAjenaControl.error || rAjenaControl.resultado);

// =============================================================================
// 22. Exportar desde un libro o carpeta: el diálogo separa la interacción de la exportación
// =============================================================================
// 22.1. El menú sólo abre el diálogo; analizar no exporta, no crea carpetas ni escribe en Control
let rSoloVistaLibro = runExportar({ libro: libroPdf, confirmar: "CANCEL" });
let rSoloVistaCarpeta = runCarpeta({ confirmar: "CANCEL" });
assert("Diálogo: el menú abre Exportar_PDF_UI.html y termina, sin pedir nada con ui.prompt",
  JSON.stringify(rSoloVistaLibro.dialogos) === JSON.stringify([{ archivo: "Exportar_PDF_UI", titulo: "Exportar PDF desde un libro o carpeta" }])
    && !rSoloVistaLibro.error,
  rSoloVistaLibro.error || JSON.stringify(rSoloVistaLibro.dialogos));
assert("Diálogo: analizar un libro o una carpeta muestra qué se haría, sin pedir PDF, crear carpetas ni escribir en Control",
  [rSoloVistaLibro, rSoloVistaCarpeta].every(r => !r.error && r.vistas.length === 1 && r.pedidos.length === 0 && r.carpetas.length === 0
    && r.resultado === "" && !r.ejecuciones.some(e => e.nombre === "exportarPdfConfirmado"))
    && rSoloVistaLibro.vistas[0].texto === "Libro: " + NOMBRE_LIBRO_EPS2 + " (4 hojas de estudiantes).\n\nSe exportan sus 4 PDF en una carpeta nueva junto al libro."
    && JSON.stringify(rSoloVistaLibro.vistas[0].opciones) === JSON.stringify([{ modo: "rehacer", etiqueta: "Exportar 4 PDF" }])
    && JSON.stringify(rSoloVistaCarpeta.vistas[0].opciones) === JSON.stringify([{ modo: "carpeta", etiqueta: "Exportar 8 PDF" }]),
  JSON.stringify(rSoloVistaLibro.vistas.concat(rSoloVistaCarpeta.vistas).map(v => v.opciones)));
let rVistaCompleta = runExportar({ libro: libroPdf, previas: [previaCompletaDe(libroPdf)], confirmar: "CANCEL" });
let rVistaCambiado = runExportar({ libro: libroPdf, previas: [previaIncompleta()], modificado: 30 * 60000, confirmar: "CANCEL" });
assert("Diálogo: la vista de un libro dice si su exportación está completa (se rehace) o si el libro cambió después (se rehace)",
  rVistaCompleta.vistas[0].texto.endsWith("Su última exportación está completa (" + previaCompletaDe(libroPdf).nombre + "): se exportan de nuevo sus 4 PDF en una carpeta nueva.")
    && rVistaCambiado.vistas[0].texto.endsWith("El libro cambió después de su última exportación (" + CARPETA_PREVIA + "): se exportan sus 4 PDF en una carpeta nueva.")
    && rVistaCambiado.vistas[0].opciones.length === 1,
  rVistaCompleta.vistas[0].texto + " / " + rVistaCambiado.vistas[0].texto);

// 22.1 bis. Garantía general: analizar no deja ningún efecto persistente, en ningún camino. Se mira el estado al terminar
// la vista previa (antes de cualquier exportación): Control igual que al empezar, sin carpetas, archivos ni pedidos de PDF,
// y los libros sólo leídos. Incluye libro y carpeta con pendientes, carpeta sin pendientes (vacía, completa o con un
// problema) y errores de análisis (link no reconocible, archivo que no es planilla, ID inexistente, libro sin Resumen,
// generación cortada), y también ejecuciones que después sí exportaron.
let rVistaLinkMalo = runExportar({ libro: libroPdf, respuesta: "boletines del curso" });
const caminosDeAnalisis = {
  "libro con pendientes": rSoloVistaLibro, "carpeta con pendientes": rSoloVistaCarpeta, "libro incompleto": rCompletar,
  "carpeta vacía": rCarpetaVacia, "carpeta completa": rCarpetaTrasControl, "carpeta con un problema": rSinAccesoCarpeta,
  "link no reconocible": rVistaLinkMalo, "archivo que no es planilla": rEsPdf, "ID inexistente": rNoExiste,
  "libro sin Resumen": rPdfGenerador, "generación cortada": rPdfCortadoIndividual, "carpeta que después se exportó": rCarpeta,
  "libro que después se exportó": rPdf
};
const sinEfectos = r => r.trasVista && r.trasVista.controlIntacto && r.trasVista.carpetas === 0 && r.trasVista.archivos === 0
  && r.trasVista.pedidos === 0 && r.trasVista.accesos.every(a => LECTURAS_LIBRO.includes(a));
assert("Analizar es sólo lectura en todos los caminos: Control intacto, sin carpetas, archivos ni PDF, y los libros sólo leídos",
  Object.values(caminosDeAnalisis).every(sinEfectos),
  Object.entries(caminosDeAnalisis).filter(([, r]) => !sinEfectos(r)).map(([c, r]) => c + ": " + JSON.stringify(r.trasVista)).join(" / "));
assert("Analizar con error: el error llega al diálogo y Control queda sin registrar nada (no hubo intento de exportación)",
  [rVistaLinkMalo, rEsPdf, rNoExiste, rPdfGenerador, rPdfCortadoIndividual].every(r => r.error && r.resultado === "" && r.advertencias === ""
    && JSON.stringify(r.ejecuciones.map(e => e.nombre)) === JSON.stringify(["exportarPdfDesdeDrive", "previsualizarExportacionPdf"])),
  [rVistaLinkMalo, rEsPdf, rNoExiste, rPdfGenerador, rPdfCortadoIndividual].map(r => r.error + " → " + r.resultado).join(" / "));
// 22.2. Al exportar se vuelve a mirar Drive y se usa el estado nuevo
const AVISO_CAMBIO = "Lo que hay en Drive cambió desde la vista previa: se usó el estado actual.\n\n";
let rSinCambios = runCarpeta({});
assert("Diálogo: sin cambios entre la vista previa y la exportación, no hay aviso",
  !rSinCambios.error && rSinCambios.alerts.length === 1 && !rSinCambios.alerts[0].startsWith(AVISO_CAMBIO), rSinCambios.alerts[0]);
// Entre la vista previa y la exportación, otra exportación completó el libro de EPSI EPS2.
let rOtroCompleto = runCarpeta({ cambiarDrive: ({ delCurso, carpetaMock }) =>
  delCurso.hijas.push(carpetaMock(libroEps2.getName() + " - PDF otra", delCurso.nombre, previaCompletaDe(libroEps2, { creada: 60000 }))) });
assert("Diálogo: una carpeta que se completó entre la vista previa y la exportación no se vuelve a exportar, y se avisa",
  !rOtroCompleto.error && rOtroCompleto.preguntas[0].includes("- EPSI EPS2 · 3º bimestre 2026 (libro <fecha>): 4 PDF a exportar.")
    && pedidosDeLibro(rOtroCompleto, libroEps2) === 0 && rOtroCompleto.pedidos.length === 4 && carpetaDeLibro(rOtroCompleto, libroEps2).length === 1
    && rOtroCompleto.alerts[0].startsWith(AVISO_CAMBIO + "PDF de la carpeta \"Boletines para revisar\": cursos completos: 5 de 5 con libro · PDF exportados ahora: 4 "),
  rOtroCompleto.error || rOtroCompleto.alerts[0]);
// Se eligió Completar, pero el libro se modificó entre la vista previa y la exportación.
let rCambioLibro = runExportar({ libro: libroPdf, previas: [previaIncompleta()], opcion: "YES",
  cambiarDrive: ({ libros }) => { libros[0].modificado = 60000; } });
assert("Diálogo: si el libro cambió después de la vista previa, no se completa la carpeta vieja: se exporta todo en una nueva, y se avisa",
  !rCambioLibro.error && rCambioLibro.carpetas.length === 1 && rCambioLibro.archivos.length === 4 && rCambioLibro.pedidos.length === 4
    && previaDe(rCambioLibro).archivos.length === 2 && rCambioLibro.alerts[0].startsWith(AVISO_CAMBIO + "PDF exportados: 4 de 4 "),
  rCambioLibro.error || rCambioLibro.alerts[0]);
// Se eligió Completar, pero otra exportación terminó esa carpeta mientras tanto.
let rYaCompletada = runExportar({ libro: libroPdf, previas: [previaIncompleta()], opcion: "YES",
  cambiarDrive: ({ delCurso }) => { delCurso.hijas.find(c => c.nombre === CARPETA_PREVIA).archivos = HOJAS_PDF.map(h => h + ".pdf"); } });
assert("Diálogo: si la carpeta a completar ya se completó, no se pide ningún PDF ni se crea otra carpeta",
  !rYaCompletada.error && rYaCompletada.pedidos.length === 0 && rYaCompletada.carpetas.length === 0
    && rYaCompletada.alerts[0].startsWith(AVISO_CAMBIO + "PDF exportados: 0 de 0 que faltaban · ")
    && rYaCompletada.resultado.includes(" La carpeta tiene 4 de 4 PDF."),
  rYaCompletada.error || rYaCompletada.alerts[0]);
let rYaNoEsta = runExportar({ libro: libroPdf, cambiarDrive: ({ libros }) => { libros[0].inaccesible = true; } });
assert("Diálogo: si el libro dejó de estar accesible después de la vista previa, la exportación da un error claro y no crea nada",
  rYaNoEsta.error === "No se encontró en Drive un libro con ese link o ID, o no hay acceso." && rYaNoEsta.carpetas.length === 0
    && rYaNoEsta.resultado === "Error al exportar PDF: " + rYaNoEsta.error && !rYaNoEsta.cerrojoTomado,
  rYaNoEsta.error);

// 22.3. Una exportación por vez
let rEnCurso = runExportar({ libro: libroPdf, exportacionEnCurso: true });
assert("Diálogo: con otra exportación en curso (cerrojo del usuario tomado) no se exporta y se avisa",
  rEnCurso.error === "Ya hay una exportación de PDF en curso: esperar a que termine y volver a analizar."
    && rEnCurso.pedidos.length === 0 && rEnCurso.carpetas.length === 0 && rEnCurso.resultado === "Error al exportar PDF: " + rEnCurso.error,
  rEnCurso.error);
assert("Diálogo: al terminar una exportación, con o sin error, el cerrojo queda libre",
  !rPdf.cerrojoTomado && !rPdfFallas.cerrojoTomado && !rYaNoEsta.cerrojoTomado && rEnCurso.cerrojoTomado);

// El registro en Control queda para las exportaciones efectivas.
assert("Exportar sí registra en Control, con resultado y con error: libro, carpeta, error al exportar y cerrojo tomado",
  rPdf.resultado.startsWith("PDF exportados: 4 de 4 ") && rCarpeta.resultado.startsWith("PDF de la carpeta \"Boletines para revisar\": cursos completos: 5 de 5")
    && rYaNoEsta.resultado === "Error al exportar PDF: " + rYaNoEsta.error && rEnCurso.resultado === "Error al exportar PDF: " + rEnCurso.error,
  [rPdf, rCarpeta, rYaNoEsta, rEnCurso].map(r => r.resultado).join(" / "));
assert("Marcados y todos, desde Control, siguen registrando su resultado y sus errores en Control",
  rPdfTodos.resultado.startsWith("PDF de varios cursos (3º bimestre 2026): cursos completos: 5 de 5 con libro")
    && rPdfTodosOtraVez.resultado.startsWith("PDF de varios cursos (3º bimestre 2026): cursos completos: 5 de 5 con libro · PDF exportados ahora: 0 ")
    && rPdfSinMarcas.resultado === "Error al exportar PDF de varios cursos: " + rPdfSinMarcas.error,
  [rPdfTodos, rPdfTodosOtraVez, rPdfSinMarcas].map(r => r.resultado).join(" / "));

// 22.4. Marcados y todos, desde Control, no cambian: confirman con ui.alert dentro de su ejecución
assert("Control: marcados y todos no abren el diálogo ni pasan por la vista previa",
  [rPdfTodos, rPdfMarcados, rPdfCancelaCursos].every(r => r.dialogos.length === 0 && r.vistas.length === 0 && r.ejecuciones.length === 0)
    && rPdfTodos.preguntas.length === 1 && rPdfConfirmaTarde.carpetas.length === 0);

// 22.5. El cliente del diálogo (Exportar_PDF_UI.html), con un DOM mínimo: una operación por vez y nunca algo distinto de lo visto
const HTML_DIALOGO = fs.readFileSync(path.join(ROOT, "src/Exportar_PDF_UI.html"), "utf8");
const cargarDialogo = () => {
  const crear = id => ({
    id, value: "", textContent: "", className: "", disabled: false, onclick: null, oninput: null, hijos: [],
    appendChild(hijo) { this.hijos.push(hijo); }, replaceChildren() { this.hijos = []; },
    querySelectorAll(selector) { return selector === "button" ? this.hijos.slice() : []; }
  });
  const ids = [...HTML_DIALOGO.matchAll(/ id="([^"]+)"/g)].map(m => m[1]);
  const elementos = Object.fromEntries(ids.map(id => [id, crear(id)]));
  const llamadas = [];
  const run = () => {
    const manejadores = {};
    const cadena = new Proxy({}, { get: (t, p) => p === "withSuccessHandler" ? f => { manejadores.exito = f; return cadena; }
      : p === "withFailureHandler" ? f => { manejadores.falla = f; return cadena; }
      : (...args) => llamadas.push(Object.assign({ funcion: p, args }, manejadores)) });
    return cadena;
  };
  const ctx = { document: { getElementById: id => elementos[id], createElement: () => crear("") }, google: { script: { get run() { return run(); } } } };
  vm.createContext(ctx);
  vm.runInContext(HTML_DIALOGO.match(/<script>([\s\S]*)<\/script>/)[1], ctx);
  return { e: elementos, llamadas, botones: () => elementos.opciones.hijos };
};
const servidorExporta = fs.readFileSync(path.join(ROOT, "src/Exportar_PDF.gs"), "utf8");
const llamadasDelHtml = [...new Set([...HTML_DIALOGO.matchAll(/\.withFailureHandler\(\w+\)\s*\.(\w+)\(/g)].map(m => m[1]))];
assert("Cliente: llama sólo a funciones públicas que existen en Exportar_PDF.gs (vista previa y exportación)",
  JSON.stringify(llamadasDelHtml) === JSON.stringify(["previsualizarExportacionPdf", "exportarPdfConfirmado"])
    && llamadasDelHtml.every(f => new RegExp("^function " + f + "\\(", "m").test(servidorExporta)),
  llamadasDelHtml.join(", "));
const d = cargarDialogo();
d.e.entrada.value = URL_CARPETA;
d.e.analizar.onclick();
d.e.analizar.onclick();
assert("Cliente: Analizar pide la vista previa una sola vez, con el link, y queda desactivado mientras espera",
  d.llamadas.length === 1 && d.llamadas[0].funcion === "previsualizarExportacionPdf" && d.llamadas[0].args[0] === URL_CARPETA && d.e.analizar.disabled);
d.llamadas[0].exito({ entrada: URL_CARPETA, texto: "Plan sintético", firma: "F1", opciones: [{ modo: "carpeta", etiqueta: "Exportar 8 PDF" }] });
assert("Cliente: la vista previa se muestra con un botón por opción, sin exportar",
  d.e.vista.textContent === "Plan sintético" && d.botones().length === 1 && d.botones()[0].textContent === "Exportar 8 PDF"
    && !d.e.analizar.disabled && d.llamadas.length === 1);
d.e.entrada.value = "otro link";
d.e.entrada.oninput();
assert("Cliente: cambiar el link borra la vista previa y sus botones: no se exporta algo distinto de lo que se ve",
  d.botones().length === 0 && d.e.vista.textContent === "");
d.e.entrada.value = URL_CARPETA;
d.e.analizar.onclick();
d.llamadas[1].exito({ entrada: URL_CARPETA, texto: "Plan sintético", firma: "F2", opciones: [{ modo: "carpeta", etiqueta: "Exportar 8 PDF" }] });
const exportarBoton = d.botones()[0];
exportarBoton.onclick();
exportarBoton.onclick();
d.e.analizar.onclick();
assert("Cliente: Exportar llama una sola vez a la exportación con el link y la firma de la vista; mientras tanto, nada se puede volver a pedir",
  d.llamadas.length === 3 && d.llamadas[2].funcion === "exportarPdfConfirmado"
    && JSON.stringify(d.llamadas[2].args) === JSON.stringify([URL_CARPETA, "carpeta", "F2"]) && exportarBoton.disabled && d.e.analizar.disabled);
d.llamadas[2].exito({ texto: "Resultado sintético" });
exportarBoton.onclick();
assert("Cliente: el resultado se muestra; la vista previa ya usada no permite exportar otra vez sin volver a analizar",
  d.e.resultado.textContent === "Resultado sintético" && exportarBoton.disabled && !d.e.analizar.disabled && d.llamadas.length === 3);
const dLibro = cargarDialogo();
dLibro.e.entrada.value = LINK_LIBRO;
dLibro.e.analizar.onclick();
dLibro.llamadas[0].exito({ entrada: LINK_LIBRO, texto: "Libro sintético", firma: "L", opciones: [{ modo: "completar", etiqueta: "Completar" }, { modo: "rehacer", etiqueta: "Exportar todo" }] });
dLibro.botones()[0].onclick();
assert("Cliente: un libro con exportación incompleta ofrece Completar o Exportar todo, y envía la opción elegida",
  dLibro.botones().length === 2 && dLibro.botones().every(b => b.disabled) && JSON.stringify(dLibro.llamadas[1].args) === JSON.stringify([LINK_LIBRO, "completar", "L"]));
const dError = cargarDialogo();
dError.e.analizar.onclick();
dError.llamadas[0].falla(new Error("No pude reconocer el ID."));
assert("Cliente: un error del servidor se muestra y deja volver a analizar",
  dError.e.estado.textContent === "Error: No pude reconocer el ID." && dError.e.estado.className === "error" && !dError.e.analizar.disabled);

// =============================================================================
// 23. Generación por DNI: un único libro con los estudiantes pedidos, de cualquier curso
// =============================================================================
// Mismas grillas sintéticas que varios cursos (sección 16). Cada DNI se genera como en la generación individual,
// sin mirar LISTO PARA EXPORTAR (se informa). En 3B (principal):
// - EPSI EPS1: Beto 97000002 (SÍ). EPSI EPS2: Ana 97000001 (SÍ), Caro 97000012 (NO), Dani 97000013 (vacío),
//   Fede 97000015 (SÍ), Gabi 97000016 ("Si"), entre otros. EPSI EPS4: Dora 97000041 (SÍ).
// - EPSG EPS1: Gael 96000001 (SÍ). EPSG EPS2: Hugo 96000021 (NO). EPSG EPS3: Iris 96000031 (SÍ).
// - 97000018 sólo existe en 1C (fuente anterior).
// El Control de varios cursos tiene hoja origen, DNI y nivel que la operación debe ignorar.
const runDni = (nombre, entrada, opciones = {}) => {
  const r = runTest(nombre, Object.assign({ control: ctlVarios({}), sources: srcVarios, templateGrid: tpl,
    ejecutar: (opciones.antes || "") + "var respuestaDni = generarLibroPorDni(" + JSON.stringify(entrada) + ")" }, opciones));
  r.respuesta = r.contexto.respuestaDni;
  r.libro = r.creados[0] || null;
  return r;
};
const NOMBRE_LIBRO_DNI = "Boletines EPS - Por DNI - 3º bimestre 2026 - <fecha>";
const RUTA_DNI = "Mi unidad > Boletines EPS > 2026 - 3º bimestre > Por DNI";
const hojasLibro = libro => libro ? libro.getSheets().map(h => h.getName()) : [];
const resumenDni = libro => {
  const g = rect(libro.getSheetByName("Resumen").grid);
  const i = g.findIndex(f => f[0] === "Estado");
  return {
    cabecera: g.slice(0, 6).map(f => f[0]),
    pedidos: (g.find(f => f[0] === "DNI pedidos") || [])[1],
    encabezado: g[i],
    filas: g.slice(i + 1).filter(f => f[0]).map(f => ({ estado: f[0], hoja: f[1], nombre: f[2], dni: f[3], curso: f[4], fila: f[5], listo: f[6], exportado: f[7], detalle: f[8] }))
  };
};
const estadosDni = r => r.libro ? resumenDni(r.libro).filas.map(f => f.estado + ":" + f.dni + ":" + f.curso).join(", ") : "";
// La hoja del estudiante en el libro de su curso (generación por curso de la sección 16, mismas fuentes).
const hojaDelCurso = (curso, hoja) => gridHoja(porCurso[curso], hoja);
const mismaHoja = (r, hojaDni, curso, hojaCurso) => {
  const h = r.libro && r.libro.getSheetByName(hojaDni);
  const c = libroDe(porCurso[curso]).getSheetByName(hojaCurso || hojaDni);
  return !!h && !!c && JSON.stringify([rect(h.grid), h.formatos, h.getFrozenRows()]) === JSON.stringify([rect(c.grid), c.formatos, c.getFrozenRows()]);
};

// 23.1. Un DNI válido
let rDni1 = runDni("DNI: uno", "97000015");
assert("Por DNI, uno: un libro nuevo con Resumen y la hoja del estudiante, con el nombre de la generación por DNI",
  !rDni1.error && rDni1.creados.length === 1 && rDni1.libro.getName() === NOMBRE_LIBRO_DNI
    && hojasLibro(rDni1.libro).join("|") === "Resumen|Ficticio, Fede Exportado",
  rDni1.error || hojasLibro(rDni1.libro).join("|"));
assert("Por DNI, uno: la hoja es idéntica (valores, colores, fila inmovilizada) a la del libro de su curso",
  mismaHoja(rDni1, "Ficticio, Fede Exportado", "EPSI EPS2"));
assert("Por DNI, uno: el libro queda en la subcarpeta Por DNI de la del período, creada la primera vez, y se mueve recién creado",
  rDni1.libro.carpeta.ruta() === RUTA_DNI && JSON.stringify(rDni1.carpetasCreadas) === JSON.stringify(["Mi unidad > Boletines EPS", "Mi unidad > Boletines EPS > 2026 - 3º bimestre", RUTA_DNI])
    && rDni1.libro.hojasAlMover === 1,
  rDni1.libro.carpeta.ruta() + " / " + rDni1.carpetasCreadas.join(", "));
assert("Por DNI, uno: Resumen oculto, con su título propio (no el de un curso), período, ciclo y DNI pedidos",
  rDni1.libro.getSheetByName("Resumen").isSheetHidden()
    && JSON.stringify(resumenDni(rDni1.libro).cabecera) === JSON.stringify(["Boletines EPS por DNI", "Período", "Ciclo lectivo", "Fecha de generación", "DNI pedidos", "Advertencias generales"])
    && rect(rDni1.libro.getSheetByName("Resumen").grid)[1][1] === P3B && rect(rDni1.libro.getSheetByName("Resumen").grid)[2][1] === "2026"
    && resumenDni(rDni1.libro).pedidos === "97000015"
    && resumenDni(rDni1.libro).encabezado.join("|") === "Estado|Hoja|Estudiante|DNI|Curso|Fila en la fuente|LISTO PARA EXPORTAR en la grilla|EXPORTADO en la grilla|Motivo o advertencias",
  JSON.stringify(resumenDni(rDni1.libro)));
assert("Por DNI, uno: el Resumen informa estado, hoja, curso, fila, LISTO PARA EXPORTAR, EXPORTADO y las advertencias generales una vez",
  (f => f.estado === "Generado" && f.hoja === "Ficticio, Fede Exportado" && f.curso === "EPSI EPS2" && f.listo === "SÍ" && f.exportado === "SÍ"
    && f.fila === enResumen(porCurso["EPSI EPS2"], "97000015").fila)(resumenDni(rDni1.libro).filas[0] || {})
    && rect(rDni1.libro.getSheetByName("Resumen").grid)[5][1] === ADV_SIN_1B,
  JSON.stringify(resumenDni(rDni1.libro).filas));
assert("Por DNI, uno: el diálogo recibe el texto y el link; Control registra el libro, los conteos y el link",
  rDni1.respuesta && rDni1.respuesta.url === rDni1.libro.getUrl()
    && rDni1.respuesta.texto.startsWith("Libro creado: " + NOMBRE_LIBRO_DNI + "\nCarpeta: " + RUTA_DNI + "\n\nDNI pedidos: 1 · Generados: 1 · Con error: 0 · No encontrados: 0.\n\n- 97000015: Generado · EPSI EPS2 · Ficticio, Fede Exportado (hoja \"Ficticio, Fede Exportado\")")
    && rDni1.respuesta.texto.includes("Exportar PDF desde un libro o carpeta…, pegar el link del libro")
    && rDni1.estadoResultado === "Libro por DNI creado: " + NOMBRE_LIBRO_DNI + " · DNI pedidos: 1 · Generados: 1 · Con error: 0 · No encontrados: 0. " + rDni1.libro.getUrl()
    && rDni1.estadoAdvertencias.split("\n")[0] === "97000015: Generado · EPSI EPS2 · Ficticio, Fede Exportado (hoja \"Ficticio, Fede Exportado\")"
    && rDni1.rawAlerts.length === 0,
  JSON.stringify(rDni1.respuesta) + " / " + rDni1.estadoResultado);
assert("Por DNI: fuentes de sólo lectura e intactas; el generador no cambia salvo el estado de Control",
  rDni1.registro.indebidos.length === 0 && JSON.stringify(srcVarios) === fotoVarios
    && rDni1.ss.getSheets().map(s => s.getName()).join("|") === "Plantilla|Control"
    && JSON.stringify(rect(rDni1.ss.getSheetByName("Plantilla").grid)) === JSON.stringify(rect(tpl)),
  rDni1.registro.indebidos.join(", "));

// 23.2. Varios del mismo curso, en el orden pedido
let rDniMismo = runDni("DNI: mismo curso", "97000016\n97000001");
assert("Por DNI, mismo curso: una hoja por DNI, en el orden pedido y no en el de la grilla",
  !rDniMismo.error && hojasLibro(rDniMismo.libro).join("|") === "Resumen|Ficticio, Gabi Si|Ficticia, Ana Real"
    && estadosDni(rDniMismo) === "Generado:97000016:EPSI EPS2, Generado:97000001:EPSI EPS2",
  rDniMismo.error || hojasLibro(rDniMismo.libro).join("|"));
assert("Por DNI, mismo curso: sólo los pedidos (no Fede ni Ana 97000011, también listos) y hojas idénticas a las del curso",
  !resumenDni(rDniMismo.libro).filas.some(f => ["97000015", "97000011"].includes(f.dni))
    && mismaHoja(rDniMismo, "Ficticio, Gabi Si", "EPSI EPS2")
    && mismaHoja(rDniMismo, "Ficticia, Ana Real", "EPSI EPS2", "Ficticia, Ana Real (97000001)"));
let rDniHomonimos = runDni("DNI: homónimos", "97000011 97000001");
assert("Por DNI: dos estudiantes con el mismo nombre se distinguen con el DNI, como en el libro del curso",
  hojasLibro(rDniHomonimos.libro).join("|") === "Resumen|Ficticia, Ana Real (97000011)|Ficticia, Ana Real (97000001)",
  hojasLibro(rDniHomonimos.libro).join("|"));

// 23.3. Cursos y orientaciones diferentes, mezclados
let rDniMezcla = runDni("DNI: cursos y orientaciones distintos", "96000031, 97000002; 97000041\t96000001");
assert("Por DNI, mezclados: un único libro con los cuatro, en el orden pedido, de tres niveles y dos orientaciones",
  !rDniMezcla.error && rDniMezcla.creados.length === 1
    && hojasLibro(rDniMezcla.libro).join("|") === "Resumen|Gastro, Iris Tres|Ficticio, Beto Nuevo|Ficticia, Dora Cuatro|Gastro, Gael Uno"
    && estadosDni(rDniMezcla) === "Generado:96000031:EPSG EPS3, Generado:97000002:EPSI EPS1, Generado:97000041:EPSI EPS4, Generado:96000001:EPSG EPS1",
  rDniMezcla.error || hojasLibro(rDniMezcla.libro).join("|") + " / " + estadosDni(rDniMezcla));
assert("Por DNI, mezclados: cada hoja es idéntica a la del libro de su curso (orientación y nivel visual incluidos)",
  mismaHoja(rDniMezcla, "Gastro, Iris Tres", "EPSG EPS3") && mismaHoja(rDniMezcla, "Ficticio, Beto Nuevo", "EPSI EPS1")
    && mismaHoja(rDniMezcla, "Ficticia, Dora Cuatro", "EPSI EPS4") && mismaHoja(rDniMezcla, "Gastro, Gael Uno", "EPSG EPS1")
    && rect(rDniMezcla.libro.getSheetByName("Gastro, Iris Tres").grid)[5][4] === hojaDelCurso("EPSG EPS3", "Gastro, Iris Tres")[5][4]);
assert("Por DNI, mezclados: la principal se abre una vez para buscar y cada orientación con DNI pedidos se lee completa una sola vez",
  JSON.stringify(rDniMezcla.registro.aperturas) === JSON.stringify([ID_CURSO_3B, ID_CURSO_3B, ID_CURSO_1C, ID_CURSO_3B, ID_CURSO_1C])
    && rDniMezcla.registro.lecturas.length === 8,
  JSON.stringify(rDniMezcla.registro));
assert("Por DNI, una sola orientación pedida: la otra sólo se mira para buscar los DNI",
  JSON.stringify(rDni1.registro.aperturas) === JSON.stringify([ID_CURSO_3B, ID_CURSO_3B, ID_CURSO_1C]),
  JSON.stringify(rDni1.registro.aperturas));

// 23.4. DNI inexistente: no se crea libro ni se toca Drive
let rDniNada = runDni("DNI: inexistente", "99999999");
assert("Por DNI, inexistente: sin libro ni carpetas, informado como no encontrado; sólo se buscó en la principal",
  !rDniNada.error && rDniNada.creados.length === 0 && rDniNada.carpetasCreadas.length === 0 && rDniNada.accesosDrive.length === 0
    && rDniNada.respuesta.url === ""
    && rDniNada.respuesta.texto === "No se creó ningún libro: ningún estudiante pedido pudo generarse.\n\n"
      + "DNI pedidos: 1 · Generados: 0 · Con error: 0 · No encontrados: 1.\n\n"
      + "- 99999999: No encontrado. No figura en las hojas EPSI ni EPSG de la fuente de 3º bimestre."
    && JSON.stringify(rDniNada.registro.aperturas) === JSON.stringify([ID_CURSO_3B]) && rDniNada.registro.lecturas.length === 2
    && rDniNada.estadoResultado === "Sin libro por DNI: ningún estudiante pedido pudo generarse. DNI pedidos: 1 · Generados: 0 · Con error: 0 · No encontrados: 1.",
  rDniNada.error || JSON.stringify(rDniNada.respuesta));
let rDniSoloAnterior = runDni("DNI: sólo en una fuente anterior", "97000018");
assert("Por DNI: un estudiante que sólo figura en una fuente anterior no se encuentra (la principal es la del período)",
  !rDniSoloAnterior.error && rDniSoloAnterior.creados.length === 0 && rDniSoloAnterior.respuesta.texto.includes("- 97000018: No encontrado."),
  rDniSoloAnterior.error);

// 23.5. DNI repetido en la entrada: un solo boletín
let rDniRepetido = runDni("DNI: repetido", "97000015\n97000015 97.000.015\n");
assert("Por DNI, repetido: se genera una sola vez (también con puntos) y se informa como repetido",
  !rDniRepetido.error && hojasLibro(rDniRepetido.libro).join("|") === "Resumen|Ficticio, Fede Exportado"
    && resumenDni(rDniRepetido.libro).filas.length === 1 && resumenDni(rDniRepetido.libro).pedidos === "97000015"
    && rDniRepetido.respuesta.texto.includes("DNI pedidos: 1 · Generados: 1 · Con error: 0 · No encontrados: 0 · Repetidos en la entrada: 1.")
    && rDniRepetido.respuesta.texto.includes("- Repetidos en la entrada (se pidieron una sola vez): 97000015"),
  rDniRepetido.error || JSON.stringify(rDniRepetido.respuesta));

// 23.6. LISTO PARA EXPORTAR distinto de SÍ: pedidos explícitamente, se generan como en la generación individual
let rDniNoListos = runDni("DNI: LISTO PARA EXPORTAR distinto de SÍ", "97000012 96000021 97000013 97000014");
assert("Por DNI, no listos: NO, vacío o no válido no impiden generar; la grilla dice lo suyo en el Resumen y en el diálogo",
  !rDniNoListos.error && hojasLibro(rDniNoListos.libro).join("|") === "Resumen|Ficticio, Caro No|Gastro, Hugo No|Ficticio, Dani Vacio|Ficticio, Eli Invalido"
    && resumenDni(rDniNoListos.libro).filas.map(f => f.estado + ":" + f.listo).join(",") === "Generado:NO,Generado:NO,Generado:,Generado:"
    && rDniNoListos.respuesta.texto.includes("- 97000012: Generado · EPSI EPS2 · Ficticio, Caro No (hoja \"Ficticio, Caro No\") · LISTO PARA EXPORTAR en la grilla: NO.")
    && rDniNoListos.respuesta.texto.includes("- 97000013: Generado · EPSI EPS2 · Ficticio, Dani Vacio (hoja \"Ficticio, Dani Vacio\") · LISTO PARA EXPORTAR en la grilla: vacío o no válido.")
    && !rDniNoListos.respuesta.texto.includes("Fede Exportado") && rDniNoListos.respuesta.texto.includes("Generados: 4 · Con error: 0"),
  rDniNoListos.error || hojasLibro(rDniNoListos.libro).join("|") + " / " + JSON.stringify(rDniNoListos.respuesta));
[["97000012", "EPSI", "Ficticio, Caro No"], ["96000021", "EPSG", "Gastro, Hugo No"], ["97000014", "EPSI", "Ficticio, Eli Invalido"]].forEach(([dni, hoja, nombre]) => {
  const ind = runTest("Individual " + dni, { control: ctlVarios({ B6: hoja, B7: dni }), sources: srcVarios, templateGrid: tpl });
  const f = resumenDni(rDniNoListos.libro).filas.find(x => x.dni === dni) || {};
  assert("Por DNI, no listos: la hoja de " + dni + " es idéntica al Informe_actual de la generación individual, con sus mismas advertencias",
    !ind.error && JSON.stringify(rect(rDniNoListos.libro.getSheetByName(nombre).grid)) === JSON.stringify(ind.out)
      && JSON.stringify([rect(rDniNoListos.libro.getSheetByName("Resumen").grid)[5][1]].concat(f.detalle ? f.detalle.split("\n") : []).filter(a => a !== "Sin advertencias").sort())
        === JSON.stringify(ind.advertencias.slice().sort()),
    ind.error || (f.detalle + " ≠ " + ind.advertencias.join("; ")));
});
let rDniCurso = runDni("DNI: el curso no cambia", "97000012");
assert("Por DNI: la generación por curso sigue omitiendo a quien tiene LISTO PARA EXPORTAR = NO",
  enResumen(porCurso["EPSI EPS2"], "97000012").estado === "Omitido" && hojasLibro(rDniCurso.libro).join("|") === "Resumen|Ficticio, Caro No");

// 23.7. Mezcla de válidos, no elegibles, inexistentes y entradas sin dígitos
let rDniMixto = runDni("DNI: mezcla", " 97000015\r\n99999999\nabc\n97000012;;97.000.041 ");
assert("Por DNI, mezcla: los encontrados tienen hoja (aunque no estén listos); cada pedido queda en el Resumen con su estado, en el orden pedido",
  !rDniMixto.error && hojasLibro(rDniMixto.libro).join("|") === "Resumen|Ficticio, Fede Exportado|Ficticio, Caro No|Ficticia, Dora Cuatro"
    && estadosDni(rDniMixto) === "Generado:97000015:EPSI EPS2, Generado:97000012:EPSI EPS2, Generado:97000041:EPSI EPS4, No encontrado:99999999:"
    && resumenDni(rDniMixto.libro).pedidos === "97000015, 99999999, 97000012, 97000041",
  rDniMixto.error || hojasLibro(rDniMixto.libro).join("|") + " / " + estadosDni(rDniMixto));
assert("Por DNI, mezcla: el texto ordena por pedido e informa no encontrados e ignorados",
  rDniMixto.respuesta.texto.includes("DNI pedidos: 4 · Generados: 3 · Con error: 0 · No encontrados: 1 · Ignorados en la entrada: 1.\n\n"
    + "- 97000015: Generado · EPSI EPS2 · Ficticio, Fede Exportado (hoja \"Ficticio, Fede Exportado\")\n"
    + "- 97000012: Generado · EPSI EPS2 · Ficticio, Caro No (hoja \"Ficticio, Caro No\") · LISTO PARA EXPORTAR en la grilla: NO.\n"
    + "- 97000041: Generado · EPSI EPS4 · Ficticia, Dora Cuatro (hoja \"Ficticia, Dora Cuatro\")\n"
    + "- 99999999: No encontrado.")
    && rDniMixto.respuesta.texto.includes("- Ignorados en la entrada (sin dígitos): \"abc\""),
  rDniMixto.respuesta && rDniMixto.respuesta.texto);

// 23.8. Entrada vacía: error claro, sin leer fuentes ni crear nada
let rDniVacio = runDni("DNI: entrada vacía", "  \n , ; ");
assert("Por DNI, vacía: error que lo dice, en Control; no se abre ninguna fuente ni se crea nada",
  rDniVacio.error === "No se ingresó ningún DNI." && rDniVacio.estadoResultado === "Error: No se ingresó ningún DNI."
    && rDniVacio.registro.aperturas.length === 0 && rDniVacio.creados.length === 0 && rDniVacio.carpetasCreadas.length === 0,
  rDniVacio.error);
let rDniLetras = runDni("DNI: sólo texto", "DNI");
assert("Por DNI: una entrada sin dígitos tampoco pide nada", rDniLetras.error === "No se ingresó ningún DNI." && rDniLetras.creados.length === 0, rDniLetras.error);

// 23.9. Errores: un DNI que figura en las dos orientaciones no genera ningún boletín; la configuración se valida
const srcDoraDoble = conVarios(EPSI_VARIOS, EPSG_VARIOS.concat(g => {
  conBloques("96000031", "96000031", [{ apellido: "Ficticia", nombres: "Dora Cuatro", dni: "97000041", listo: "SÍ" }])(g);
  conNivel("EPS4")(g);
}));
let rDniDoble = runDni("DNI: en las dos orientaciones", "97000041\n97000015", { sources: srcDoraDoble });
assert("Por DNI, en dos orientaciones: un error, como en la generación individual, sin hoja (nunca dos boletines), y el resto se genera",
  !rDniDoble.error && hojasLibro(rDniDoble.libro).join("|") === "Resumen|Ficticio, Fede Exportado"
    && estadosDni(rDniDoble) === "Error:97000041:EPSI EPS4, Generado:97000015:EPSI EPS2"
    && resumenDni(rDniDoble.libro).filas[0].detalle.startsWith("El DNI aparece en las hojas EPSI y EPSG de la fuente de 3º bimestre.")
    && rDniDoble.respuesta.texto.includes("Con error: 1"),
  rDniDoble.error || estadosDni(rDniDoble));
let rDniSinFuente = runDni("DNI: sin fuente principal", "97000015", { control: ctlVarios({ [F3B]: "" }) });
assert("Por DNI: la configuración se valida como en varios cursos (falta la fuente principal), sin crear nada",
  rDniSinFuente.error === "Falta la URL o ID del archivo de valoraciones del 3º bimestre." && rDniSinFuente.creados.length === 0,
  rDniSinFuente.error);
// Orientaciones: sólo se leen completas las que tienen algún DNI pedido.
const srcEpsgSinListo = conVarios(EPSI_VARIOS, EPSG_VARIOS.concat(conNivel("EPS4"), conEncabezados({ "LISTO PARA EXPORTAR": "LISTO PARA IMPRIMIR" })));
let rDniEpsgSinListo = runDni("DNI: EPSG sin LISTO PARA EXPORTAR", "97000015 96000001", { sources: srcEpsgSinListo });
assert("Por DNI: como en la generación individual, la columna LISTO PARA EXPORTAR no se exige",
  !rDniEpsgSinListo.error && hojasLibro(rDniEpsgSinListo.libro).join("|") === "Resumen|Ficticio, Fede Exportado|Gastro, Gael Uno"
    && (resumenDni(rDniEpsgSinListo.libro).filas[1] || {}).listo === "",
  rDniEpsgSinListo.error);
const srcEpsgRota = conVarios(EPSI_VARIOS, EPSG_VARIOS.concat(conNivel("EPS4"), conEncabezados({ "ORIENTACIÓN": "OTRA COSA" })));
let rDniEpsgRota = runDni("DNI: EPSG con estructura rota, pedidos de EPSI", "97000015", { sources: srcEpsgRota });
assert("Por DNI: una orientación con estructura ilegible no impide generar los DNI de la otra",
  !rDniEpsgRota.error && hojasLibro(rDniEpsgRota.libro).join("|") === "Resumen|Ficticio, Fede Exportado"
    && JSON.stringify(rDniEpsgRota.registro.aperturas) === JSON.stringify([ID_CURSO_3B, ID_CURSO_3B, ID_CURSO_1C]),
  rDniEpsgRota.error);
let rDniEpsgRotaPedida = runDni("DNI: EPSG con estructura rota, con un DNI de EPSG", "97000015 96000001", { sources: srcEpsgRota });
assert("Por DNI: si la orientación rota tiene un DNI pedido, error de la fuente principal antes de crear el libro",
  rDniEpsgRotaPedida.error && rDniEpsgRotaPedida.error.startsWith("No se encontró la columna ORIENTACIÓN que separa las materias de los campos generales en la fuente de 3º bimestre.")
    && rDniEpsgRotaPedida.creados.length === 0 && rDniEpsgRotaPedida.estadoResultado === "Error: " + rDniEpsgRotaPedida.error,
  rDniEpsgRotaPedida.error);
const srcSinEpsg = Object.assign({}, srcVarios, { [ID_CURSO_3B]: { EPSI: srcVarios[ID_CURSO_3B].EPSI } });
let rDniSinEpsg = runDni("DNI: fuente sin hoja EPSG", "97000015 96000001", { sources: srcSinEpsg });
assert("Por DNI: sin la hoja EPSG se generan los de EPSI; los de EPSG quedan no encontrados, con la advertencia de que no se buscaron",
  !rDniSinEpsg.error && hojasLibro(rDniSinEpsg.libro).join("|") === "Resumen|Ficticio, Fede Exportado"
    && rDniSinEpsg.respuesta.texto.includes("- 96000001: No encontrado.")
    && rDniSinEpsg.respuesta.texto.includes("- No se encontró la hoja EPSG en la fuente de 3º bimestre: no se buscaron DNI en ella."),
  rDniSinEpsg.error || JSON.stringify(rDniSinEpsg.respuesta));
let rDniFuenteIlegible = runDni("DNI: fuente principal inexistente", "97000015", { control: ctlVarios({ [F3B]: "ID_INEXISTENTE_xxxxxxxxxxxxxx" }) });
assert("Por DNI: una fuente principal que no puede abrirse es un error de la fuente principal, sin crear nada",
  rDniFuenteIlegible.error && rDniFuenteIlegible.error.endsWith("Es la fuente principal del período a generar (3º bimestre), por eso no se puede continuar.")
    && rDniFuenteIlegible.creados.length === 0,
  rDniFuenteIlegible.error);
let rDniSinHojas = runDni("DNI: ninguna hoja pudo crearse", "97000015",
  { antes: "SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plantilla').copyTo = () => { throw new Error('copia rechazada'); }; " });
assert("Por DNI: si la hoja no puede crearse, el libro queda con el Resumen visible y el error (como el del curso)",
  !rDniSinHojas.error && hojasLibro(rDniSinHojas.libro).join("|") === "Resumen" && !rDniSinHojas.libro.getSheetByName("Resumen").isSheetHidden()
    && estadosDni(rDniSinHojas) === "Error:97000015:EPSI EPS2" && resumenDni(rDniSinHojas.libro).filas[0].detalle.startsWith("No se pudo crear la hoja: copia rechazada"),
  rDniSinHojas.error || estadosDni(rDniSinHojas));

// 23.10. Carpeta existente: se reutiliza
let rDniCarpeta = runDni("DNI: carpeta existente", "97000015", { prepararDrive: raiz => raiz.createFolder("Boletines EPS").createFolder("2026 - 3º bimestre").createFolder("Por DNI") });
assert("Por DNI: la subcarpeta Por DNI existente se reutiliza", !rDniCarpeta.error && rDniCarpeta.carpetasCreadas.length === 0
  && rDniCarpeta.libro.carpeta.ruta() === RUTA_DNI, rDniCarpeta.error || rDniCarpeta.carpetasCreadas.join(", "));

// 23.11. PDF: el libro por DNI se exporta desde su link como uno de curso; las exportaciones de cursos no lo toman
const libroDniPdf = rDniMezcla.libro;
libroDniPdf.id = "ID_LIBRO_POR_DNI_xxxxxxxxxxxxx";
const fotoLibroDni = fotoLibro(libroDniPdf);
let rPdfDni = runExportar({ libro: libroDniPdf, respuesta: "https://docs.google.com/spreadsheets/d/" + libroDniPdf.id + "/edit" });
assert("PDF del libro por DNI: un PDF por hoja de estudiante, con su nombre y sin Resumen, en una carpeta nueva junto al libro",
  !rPdfDni.error && rPdfDni.carpetas.length === 1 && rPdfDni.carpetas[0].nombre === NOMBRE_LIBRO_DNI + " - PDF <fecha>"
    && JSON.stringify(rPdfDni.archivos) === JSON.stringify(["Gastro, Iris Tres.pdf", "Ficticio, Beto Nuevo.pdf", "Ficticia, Dora Cuatro.pdf", "Gastro, Gael Uno.pdf"])
    && rPdfDni.vistas[0].texto === "Libro: " + NOMBRE_LIBRO_DNI + " (4 hojas de estudiantes).\n\nSe exportan sus 4 PDF en una carpeta nueva junto al libro."
    && fotoLibro(libroDniPdf) === fotoLibroDni && soloLeyo(rPdfDni),
  rPdfDni.error || rPdfDni.archivos.join(" | "));
let rPdfDniCompletar = runExportar({ libro: libroDniPdf, respuesta: libroDniPdf.id, opcion: "YES",
  previas: [{ nombre: NOMBRE_LIBRO_DNI + " - PDF 2026-10-16 08:00", descripcion: "PDF del libro de boletines " + libroDniPdf.id, archivos: ["Gastro, Iris Tres.pdf"], creada: 3600000 }] });
assert("PDF del libro por DNI: una exportación incompleta se completa sin repetir los PDF ya exportados",
  !rPdfDniCompletar.error && rPdfDniCompletar.carpetas.length === 0 && rPdfDniCompletar.pedidos.length === 3
    && rPdfDniCompletar.vistas[0].opciones.map(o => o.modo).join(",") === "completar,rehacer",
  rPdfDniCompletar.error);
let rPdfCursosConDni = runPdfCursos({ libros: LIBROS_TODOS.concat([{ libro: libroDniPdf, creada: 1000 }]) });
assert("PDF de todos los cursos: un libro por DNI en la carpeta del período, aunque sea el más reciente, no reemplaza al de ningún curso",
  !rPdfCursosConDni.error && rPdfCursosConDni.pedidos.length === 8 && pedidosDeLibro(rPdfCursosConDni, libroDniPdf) === 0
    && rPdfCursosConDni.preguntas[0] === rPdfTodos.preguntas[0],
  rPdfCursosConDni.error || rPdfCursosConDni.preguntas.join(" / "));
let rCarpetaConDni = runCarpeta({ libros: LIBROS_TODOS.concat([{ libro: libroDniPdf, creada: 1000 }]) });
assert("PDF de una carpeta elegida a mano: el libro por DNI va aparte, como un libro más, sin cambiar los de los cursos",
  !rCarpetaConDni.error && rCarpetaConDni.pedidos.length === 12 && pedidosDeLibro(rCarpetaConDni, libroDniPdf) === 4
    && rTodos.creados.every(l => pedidosDeLibro(rCarpetaConDni, l) === hojasPdfDe(l).length)
    && rCarpetaConDni.preguntas[0].includes("\n- Libro \"" + NOMBRE_LIBRO_DNI + "\" (libro <fecha>): 4 PDF a exportar."),
  rCarpetaConDni.error || rCarpetaConDni.preguntas.join(" / "));

// 23.12. Las otras generaciones no cambian
let rCursoDespues = runCurso("Por curso, regresión", ctlCurso("EPS2"), srcCurso);
assert("Por DNI: la generación por curso produce el mismo libro de siempre (cada hoja, Resumen y su título de curso)",
  !rCursoDespues.error && fotoLibro(libroDe(rCursoDespues)) === fotoLibro(libroDe(rCurso))
    && gridHoja(rCursoDespues, "Resumen")[0][0] === "Boletines EPS por curso" && libroDe(rCursoDespues).carpeta.ruta() === "Mi unidad > Boletines EPS > 2026 - 3º bimestre",
  rCursoDespues.error);

// 23.13. El menú abre el diálogo; el cliente (Generar_DNI_UI.html) hace una generación por vez
const genDialogoDni = cargarGenerador();
const dialogosDni = [];
genDialogoDni.HtmlService = { createHtmlOutputFromFile: nombre => {
  if (!fs.existsSync(path.join(ROOT, "src", nombre + ".html"))) throw new Error("mock: no existe el HTML " + nombre);
  return { archivo: nombre, setWidth() { return this; }, setHeight() { return this; } };
} };
genDialogoDni.SpreadsheetApp = { getUi: () => ({ showModalDialog: (html, titulo) => dialogosDni.push(html.archivo + ":" + titulo) }) };
genDialogoDni.generarBoletinesPorDni();
assert("Por DNI: el menú sólo abre el diálogo (archivo existente en src)", dialogosDni.join() === "Generar_DNI_UI:Generar boletines por DNI", dialogosDni.join());
const HTML_DNI = fs.readFileSync(path.join(ROOT, "src/Generar_DNI_UI.html"), "utf8");
const llamadasDni = [...new Set([...HTML_DNI.matchAll(/\.withFailureHandler\(\w+\)\s*\.(\w+)\(/g)].map(m => m[1]))];
assert("Cliente DNI: llama sólo a generarLibroPorDni, pública en Generador_EPS.gs",
  llamadasDni.join() === "generarLibroPorDni" && /^function generarLibroPorDni\(/m.test(codigoGenerador), llamadasDni.join());
const cargarDialogoDni = () => {
  const crear = id => ({ id, value: "", textContent: "", className: "", disabled: false, hidden: true, href: "", onclick: null, select() {} });
  const elementos = Object.fromEntries([...HTML_DNI.matchAll(/ id="([^"]+)"/g)].map(m => [m[1], crear(m[1])]));
  const llamadas = [];
  const run = () => {
    const manejadores = {};
    const cadena = new Proxy({}, { get: (t, p) => p === "withSuccessHandler" ? f => { manejadores.exito = f; return cadena; }
      : p === "withFailureHandler" ? f => { manejadores.falla = f; return cadena; }
      : (...args) => llamadas.push(Object.assign({ funcion: p, args }, manejadores)) });
    return cadena;
  };
  const ctx = { document: { getElementById: id => elementos[id] }, google: { script: { get run() { return run(); } } } };
  vm.createContext(ctx);
  vm.runInContext(HTML_DNI.match(/<script>([\s\S]*)<\/script>/)[1], ctx);
  return { e: elementos, llamadas };
};
const dd = cargarDialogoDni();
dd.e.dnis.value = "97000015\n96000031";
dd.e.generar.onclick();
dd.e.generar.onclick();
assert("Cliente DNI: Generar envía el texto pegado tal cual, una sola vez, y queda desactivado mientras espera",
  dd.llamadas.length === 1 && dd.llamadas[0].funcion === "generarLibroPorDni" && dd.llamadas[0].args[0] === "97000015\n96000031" && dd.e.generar.disabled);
dd.llamadas[0].exito({ texto: "Resultado sintético", url: "https://docs.google.com/spreadsheets/d/LIBRO_SINTETICO/edit" });
assert("Cliente DNI: muestra el resultado, el enlace para revisar y el link para exportar; se puede volver a generar",
  dd.e.resultado.textContent === "Resultado sintético" && !dd.e.libro.hidden && dd.e.abrir.href.includes("LIBRO_SINTETICO")
    && dd.e.link.value.includes("LIBRO_SINTETICO") && !dd.e.generar.disabled);
dd.e.generar.onclick();
dd.llamadas[1].exito({ texto: "Sin libro sintético", url: "" });
assert("Cliente DNI: sin libro, no muestra enlace (tampoco el de la generación anterior)",
  dd.e.resultado.textContent === "Sin libro sintético" && dd.e.libro.hidden);
dd.e.generar.onclick();
dd.llamadas[2].falla(new Error("No se ingresó ningún DNI."));
assert("Cliente DNI: un error del servidor se muestra y deja volver a generar",
  dd.e.estado.textContent === "Error: No se ingresó ningún DNI." && dd.e.estado.className === "error" && !dd.e.generar.disabled);

console.log(`\nResultado final: ${fail === 0 ? "TODOS LOS TESTS PASARON" : `${fail} TESTS FALLARON`}`);
process.exit(fail ? 1 : 0);
