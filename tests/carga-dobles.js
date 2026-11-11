/**
 * Carga docente EPS (#27) · entorno local compartido por tests/carga-interfaz.js y tests/carga-dev.js.
 *
 * Carga en un contexto propio, como un proyecto de Apps Script, los tres archivos reales del servidor
 * (Carga_Nucleo.gs, Carga_Google.gs y Carga_Web.gs) con dobles de SpreadsheetApp, LockService, PropertiesService,
 * Utilities, Session y HtmlService, sobre una planilla de administración, un Control y grillas con datos ficticios.
 *
 * Única sustitución del código del servidor: identificarDocente_ (el límite de #28), por un mapa de credenciales
 * aleatorias que crea este entorno (ver sustituirIdentidad). Todo lo demás corre tal cual.
 *
 * Los dobles de Sheets son más simples que los de tests/carga.js (que prueban el núcleo a fondo): alcanzan para la
 * interfaz e inyectan las fallas que ésta debe mostrar bien (escritura ambigua, lock ocupado, cambios del Equipo EPS).
 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const crypto = require("crypto");

const RAIZ = path.resolve(__dirname, "..");
const leer = p => fs.readFileSync(path.join(RAIZ, p), "utf8");
const ARCHIVOS_SERVIDOR = ["src/carga/Carga_Nucleo.gs", "src/carga/Carga_Google.gs", "src/carga/Carga_Web.gs"];
const ZONA = "America/Argentina/Buenos_Aires";

// ================================================================ hojas

const esFecha = x => Object.prototype.toString.call(x) === "[object Date]";

class Hoja {
  constructor(filas) {
    const ancho = Math.max(1, ...filas.map(f => f.length));
    this.v = filas.map(f => Array.from({ length: ancho }, (_, i) => (f[i] === undefined || f[i] === null ? "" : f[i])));
    this.fallas = [];      // próximas fallas de setValue: 'fallar' (sin aplicar) o 'aplicar-y-fallar'
    this.alEscribir = [];  // funciones de un solo uso, después de aplicar un setValue (el Equipo EPS edita justo después)
    this.escrituras = [];  // setValue aplicados: { fila, col, valor }
    this.intentos = 0;     // todo setValue, aunque falle
  }
  ancho() { return this.v[0] ? this.v[0].length : 1; }
  asegurar(f, c) {
    while (this.v.length < f) this.v.push(Array(this.ancho()).fill(""));
    if (c > this.ancho()) this.v.forEach(r => { while (r.length < c) r.push(""); });
  }
  ver(x) { return esFecha(x) ? x.toISOString() : String(x); }
  // Como Sheets: el apóstrofo inicial fuerza texto y no se muestra.
  poner(f, c, x) { this.v[f - 1][c - 1] = typeof x === "string" && x.startsWith("'") ? x.slice(1) : x; }
  getDataRange() {
    return { getDisplayValues: () => this.v.map(f => f.map(x => this.ver(x))), getValues: () => this.v.map(f => f.slice()) };
  }
  getLastRow() { for (let r = this.v.length; r > 0; r--) if (this.v[r - 1].some(x => x !== "")) return r; return 0; }
  getLastColumn() { return this.ancho(); }
  getRange(f, c, nf, nc) {
    nf = nf || 1; nc = nc || 1;
    const h = this;
    const bloque = crudo => Array.from({ length: nf }, (_, i) => Array.from({ length: nc }, (_, j) => {
      const x = (h.v[f - 1 + i] || [])[c - 1 + j];
      const y = x === undefined ? "" : x;
      return crudo ? y : h.ver(y);
    }));
    const rango = {
      getDisplayValues: () => bloque(false),
      getValues: () => bloque(true),
      getDisplayValue: () => bloque(false)[0][0],
      setValue: x => {
        h.intentos++;
        const falla = h.fallas.shift();
        if (falla === "fallar") throw new Error("Servicio de hojas de cálculo no disponible (simulado)");
        h.asegurar(f, c);
        h.poner(f, c, x);
        h.escrituras.push({ fila: f - 1, col: c - 1, valor: h.ver(h.v[f - 1][c - 1]) });
        const despues = h.alEscribir.shift();
        if (despues) despues(h, f - 1, c - 1);
        if (falla === "aplicar-y-fallar") throw new Error("Se excedió el tiempo de espera (simulado; la escritura sí se aplicó)");
        return rango;
      },
      setValues: m => {
        h.asegurar(f + nf - 1, c + nc - 1);
        m.forEach((r, i) => r.forEach((x, j) => h.poner(f + i, c + j, x)));
        return rango;
      },
      setFontWeight: () => rango
    };
    return rango;
  }
  appendRow(fila) {
    const r = this.getLastRow() + 1;
    this.asegurar(r, fila.length);
    fila.forEach((x, j) => this.poner(r, j + 1, x));
  }
}

function nuevoLock() {
  return {
    tomado: false, ocupadas: 0, // ocupadas: cuántos tryLock siguientes fallan (otro guardado lo tiene)
    tryLock() { if (this.ocupadas > 0) { this.ocupadas--; return false; } if (this.tomado) return false; this.tomado = true; return true; },
    releaseLock() { this.tomado = false; }
  };
}

// ================================================================ grillas ficticias con la estructura real

// const CARGA del núcleo: no es una propiedad del contexto (sólo las declaraciones function lo son).
const cargaDe = C => vm.runInContext("CARGA", C);
const leerTsv = p => leer(p).replace(/\n$/, "").split("\n").map(l => l.split("\t"));
function prng(s) { return () => { s |= 0; s = s + 0x6D2B79F5 | 0; let t = Math.imul(s ^ s >>> 15, 1 | s); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
const EPSI_1C = "reference/fixtures/1er_cuatrimestre_estructura_real_EPSI.tsv";
const EPSG_3B = "reference/fixtures/3er_bimestre_estructura_real_EPSG.tsv";
const ETIQUETAS = ["Valoración", "Calidad de las producciones:", "Participación:", "Observaciones:"];
// Nombres inventados (no corresponden a estudiantes reales).
const APELLIDOS = ["Ficticia Acosta", "Ficticio Benítez", "Ficticia Cabrera", "Ficticio Domínguez", "Ficticia Escobar", "Ficticio Funes",
  "Ficticia Giménez", "Ficticio Herrera", "Ficticia Ibarra", "Ficticio Juárez", "Ficticia Ledesma", "Ficticio Medina", "Ficticia Núñez",
  "Ficticio Ortiz", "Ficticia Paz", "Ficticio Quiroga", "Ficticia Ríos", "Ficticio Sosa", "Ficticia Toledo", "Ficticio Vera"];
const NOMBRES = ["Ana", "Bruno", "Carla", "Dante", "Elena", "Facundo", "Gala", "Hugo", "Inés", "Juan", "Kiara", "Lucas", "Mora", "Nahuel"];

// En estas materias los casos que la interfaz debe mostrar están siempre: dos pendientes, ACREDITÓ, Equivalencia,
// un valor no reconocido («Muy bien») y una calificación; además, el primer estudiante de cada otro grupo la cursa.
const MATERIAS_CON_CASOS = ["Matemática II", "Matemática I", "Lengua y Literatura I"];
function casoFijo(m, grupo, s, escala) {
  if (!MATERIAS_CON_CASOS.includes(m.nombre)) return null;
  if (m.nivel !== grupo) return s === 0 ? "CURSA" : null;
  return [ "CURSA", "CURSA", "ACREDITÓ", "Equivalencia", "Muy bien", escala[escala.length - 2]][s] || null;
}

/**
 * Encabezado de un fixture real y estudiantes ficticios en bloques de cuatro filas. Cada uno cursa las materias de
 * su grupo y algunos, además, una de otro nivel. Hay CURSA, NO CURSA, ACREDITÓ, Equivalencia y valores no reconocidos.
 */
function grillaSintetica(C, fixture, periodo, semilla, dniBase, porNivel) {
  const azar = prng(semilla);
  const base = leerTsv(fixture);
  const enc = [base[0], base[1]];
  const est = C.estructuraGrilla_(enc, periodo);
  const CARGA = cargaDe(C);
  const ancho = enc[0].length;
  const escala = CARGA.ESCALA_DOCENTE[periodo];
  const filas = [enc[0].slice(), enc[1].slice()];
  const elegir = l => l[Math.floor(azar() * l.length)];
  let n = 0;
  ["EPS1", "EPS2", "EPS3", "EPS4"].forEach(grupo => {
    const marcador = Array(ancho).fill(""); marcador[0] = grupo; filas.push(marcador);
    for (let s = 0; s < (porNivel || 6); s++, n++) {
      const b = [0, 1, 2, 3].map(() => Array(ancho).fill(""));
      b[0][0] = APELLIDOS[(n * 7 + semilla) % APELLIDOS.length]; b[0][1] = NOMBRES[(n * 5) % NOMBRES.length]; b[0][2] = String(dniBase + n);
      const otroNivel = azar() < 0.3 ? elegir(est.materias.filter(m => m.nivel !== grupo)) : null;
      est.materias.forEach(m => {
        ETIQUETAS.forEach((e, i) => { b[i][m.columnaValor - 1] = e; });
        const fijo = casoFijo(m, grupo, s, escala);
        if (fijo === null && m.nivel !== grupo && m !== otroNivel) return;
        const r = azar();
        const valor = fijo !== null ? fijo : r < 0.05 ? "NO CURSA" : r < 0.45 ? "CURSA" :
          r < 0.50 ? "ACREDITÓ" : r < 0.53 ? "Equivalencia" : r < 0.55 ? "Muy bien" : elegir(escala);
        b[0][m.columnaValor] = valor;
        if (valor === "NO CURSA" || valor === "CURSA") return;
        b[1][m.columnaValor] = elegir(CARGA.OPCIONES.calidad);
        b[2][m.columnaValor] = elegir(CARGA.OPCIONES.participacion);
        if (azar() < 0.3) b[3][m.columnaValor] = "Observación ficticia de prueba";
      });
      const col = re => enc[0].findIndex(h => re.test(h));
      b[0][col(/ORIENTACI/i)] = fixture.includes("EPSG") ? "Gastronomía" : "Informática";
      if (col(/LISTO PARA EXPORTAR/i) >= 0) b[0][col(/LISTO PARA EXPORTAR/i)] = "SÍ";
      filas.push(...b);
    }
  });
  return filas;
}

// ================================================================ escenario

const P2C = "2º cuatrimestre", PAN = "Anual", P3B = "3º bimestre";
const ID = { [P2C]: "GRILLA_2C_ficticia_00000000001", [PAN]: "GRILLA_AN_ficticia_00000000001", [P3B]: "GRILLA_3B_ficticia_00000000001" };
const OTRA_2C = "GRILLA_2C_ficticia_REEMPLAZO_01";
const ADMIN_ID = "ADMIN_ficticia_000000000001";
const GENERADOR_ID = "GENERADOR_ficticio_00000001";
const URL = id => "https://docs.google.com/spreadsheets/d/" + id + "/edit#gid=0";
const DOCENTES = {
  matematica: { correo: "docente.matematica@ficticio.example", nombre: "Docente Ficticio de Matemática" },
  lengua: { correo: "docente.lengua@ficticio.example", nombre: "Docente Ficticia de Lengua" },
  reemplazo: { correo: "docente.reemplazo@ficticio.example", nombre: "Docente Ficticio de Reemplazo" },
  sinMaterias: { correo: "docente.sin.materias@ficticio.example", nombre: "Docente sin materias" }
};
const ASIGNACIONES = [["Orientación", "Nivel", "Materia", "Docente", "Correo"],
  ["EPSI", "EPS2", "Matemática II", DOCENTES.matematica.nombre, DOCENTES.matematica.correo],
  ["EPSI", "EPS2", "Inglés I", DOCENTES.matematica.nombre, DOCENTES.matematica.correo],
  ["EPSG", "EPS1", "Matemática I", DOCENTES.matematica.nombre, DOCENTES.matematica.correo],
  ["EPSI", "EPS1", "Lengua y Literatura I", DOCENTES.lengua.nombre, DOCENTES.lengua.correo],
  ["EPSI", "EPS3", "Matemática III", DOCENTES.lengua.nombre, DOCENTES.lengua.correo]];

const isoBA = ms => new Intl.DateTimeFormat("en-CA", { timeZone: ZONA, year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date(ms));

/**
 * Estado completo de un entorno: planilla de administración (Ventanas con 2C y anual compartiendo ventana alrededor
 * de «hoy», 3B ya cerrado; Feriados; Asignaciones; Registro), Control del generador y grillas de 2C, anual y 3B.
 * opciones: { hoy (ms; por defecto, ahora), porNivel }.
 */
function escenario(C, opciones) {
  const o = opciones || {};
  const hoy = o.hoy || Date.now();
  const desde = isoBA(hoy - 3 * 86400000), hasta = isoBA(hoy + 10 * 86400000);
  const anio = Number(isoBA(hoy).slice(0, 4));
  const grillas = {};
  const crear = (id, periodo, semilla) => {
    grillas[id] = {
      EPSI: new Hoja(grillaSintetica(C, EPSI_1C, periodo, semilla, 90000000, o.porNivel)),
      EPSG: new Hoja(grillaSintetica(C, EPSG_3B, periodo, semilla + 1, 92000000, o.porNivel))
    };
  };
  crear(ID[P2C], P2C, 7);
  crear(ID[PAN], PAN, 7); // mismos estudiantes y valores iniciales que 2C: la independencia se ve al guardar
  crear(OTRA_2C, P2C, 7);
  grillas[ID[P3B]] = { EPSI: new Hoja(grillaSintetica(C, "reference/fixtures/3er_bimestre_estructura_real_EPSI.tsv", P3B, 11, 90000000, o.porNivel)),
    EPSG: new Hoja(grillaSintetica(C, EPSG_3B, P3B, 12, 92000000, o.porNivel)) };
  const control = Array.from({ length: 14 }, () => ["", ""]);
  control[11][1] = URL(ID[P3B]); control[12][1] = URL(ID[P2C]); control[13][1] = ID[PAN]; // B12, B13 (URL) y B14 (ID)
  return {
    hoy, grillas,
    admin: {
      Ventanas: new Hoja([["Período", "Desde", "Hasta"], [P2C, desde, hasta], [PAN, desde, hasta], ["3er bimestre", isoBA(hoy - 60 * 86400000), isoBA(hoy - 45 * 86400000)]]),
      Feriados: new Hoja([["Fecha", "Descripción"], [(anio - 1) + "-12-25", "Navidad"], [anio + "-12-25", "Navidad"], [(anio + 1) + "-01-01", "Año Nuevo"]]),
      Asignaciones: new Hoja(ASIGNACIONES.map(f => f.slice())),
      Registro: new Hoja([cargaDe(C).REGISTRO.slice()])
    },
    control: new Hoja(control),
    lock: nuevoLock()
  };
}

// ================================================================ contexto del servidor (los tres .gs reales)

/**
 * Contexto de Apps Script con los archivos del servidor y dobles de los servicios. reloj.desfase (ms) corre la hora
 * del servidor (para simular el cierre de la ventana con la página abierta); sólo afecta a este contexto.
 */
function servidor(opciones) {
  const o = opciones || {};
  const reloj = { desfase: 0 };
  const RealDate = Date;
  class RelojDate extends RealDate {
    constructor(...a) { if (a.length) super(...a); else super(RealDate.now() + reloj.desfase); }
    static now() { return RealDate.now() + reloj.desfase; }
  }
  const registroConsola = [];
  const consola = { log: m => registroConsola.push(String(m)), error: m => registroConsola.push("ERROR " + String(m)), warn: () => {} };
  const estado = { s: null, aperturas: [], propietario: "eps@ficticio.example", usuario: "eps@ficticio.example" };
  const libro = id => {
    const s = estado.s;
    const hojas = id === ADMIN_ID ? s.admin : id === GENERADOR_ID ? { Control: s.control } : s.grillas[id];
    if (!hojas) throw new Error("No se encontró la planilla " + id + " (simulado)");
    estado.aperturas.push(id);
    return { getSheetByName: n => hojas[n] || null, getSpreadsheetTimeZone: () => ZONA, insertSheet: n => (hojas[n] = new Hoja([[""]])) };
  };
  const globales = {
    console: consola,
    Date: RelojDate,
    SpreadsheetApp: { openById: libro, flush: () => {} },
    LockService: { getScriptLock: () => estado.s.lock },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => ({ ADMIN_ID, GENERADOR_ID })[k] || null }) },
    Utilities: { formatDate: (d, zona) => new Intl.DateTimeFormat("en-CA", { timeZone: zona, year: "numeric", month: "2-digit", day: "2-digit" }).format(d) },
    Session: { getEffectiveUser: () => ({ getEmail: () => estado.propietario }), getActiveUser: () => ({ getEmail: () => estado.usuario }) },
    HtmlService: {
      createTemplateFromFile: nombre => {
        const t = { evaluate: () => ({ contenido: renderPagina(t.credencial, {}, nombre), setTitle() { return this; }, addMetaTag() { return this; } }) };
        return t;
      },
      createHtmlOutputFromFile: nombre => ({ getContent: () => leer("src/carga/" + nombre + ".html") })
    }
  };
  const ctx = Object.assign({}, globales);
  vm.createContext(ctx);
  ARCHIVOS_SERVIDOR.forEach(a => vm.runInContext(leer(a), ctx, { filename: a }));
  estado.s = escenario(ctx, o);
  return { ctx, estado, reloj, consola: registroConsola, get s() { return estado.s; }, reiniciar: op => { estado.s = escenario(ctx, Object.assign({}, o, op)); } };
}

/**
 * LÍMITE #28 en el entorno local: reemplaza identificarDocente_ por un mapa de credenciales aleatorias creadas acá.
 * La página sólo conoce su credencial; el servidor decide el correo. Devuelve { credencial → correo }.
 */
function sustituirIdentidad(ctx, correos) {
  const mapa = Object.create(null);
  correos.forEach(c => { mapa[crypto.randomBytes(18).toString("base64url")] = c; });
  ctx.__credencialesLocales = mapa;
  vm.runInContext("identificarDocente_ = function (credencial) { var c = __credencialesLocales[credencial]; return c ? { correo: c } : null; };", ctx);
  return mapa;
}

// ================================================================ página

const escHtml = t => String(t).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

/**
 * La página como la arma HtmlService: incluye Carga_Cliente y escapa la credencial (<?= ?>). Sólo admite esos dos
 * scriptlets; cualquier otro es un error, para que el servidor local no diverja de Apps Script sin aviso.
 * opciones.cabecera: HTML que se agrega al final de <head> (el servidor local inyecta ahí su google.script.run).
 */
function renderPagina(credencial, opciones, nombre) {
  let html = leer("src/carga/" + (nombre || "Carga_Pagina") + ".html");
  html = html.replace(/<\?!=\s*incluirCarga_\('([A-Za-z_]+)'\);?\s*\?>/g, (_, n) => leer("src/carga/" + n + ".html"));
  html = html.replace(/<\?=\s*credencial\s*\?>/g, escHtml(credencial === undefined || credencial === null ? "" : credencial));
  if (/<\?/.test(html)) throw new Error("La página usa un scriptlet que el entorno local no reproduce.");
  if (opciones && opciones.cabecera) html = html.replace("</head>", opciones.cabecera + "\n</head>");
  return html;
}

/** Lógica del cliente (Carga_Cliente.html) en un contexto propio, como en el navegador. */
function cliente() {
  const fuente = leer("src/carga/Carga_Cliente.html").match(/<script>([\s\S]*)<\/script>/)[1];
  const ctx = { crypto: globalThis.crypto, console };
  vm.createContext(ctx);
  vm.runInContext(fuente + "\n;this.CargaCliente = CargaCliente;", ctx, { filename: "Carga_Cliente.html" });
  return ctx.CargaCliente;
}

module.exports = { cargaDe, Hoja, nuevoLock, grillaSintetica, escenario, servidor, sustituirIdentidad, renderPagina, cliente, escHtml, leer,
  P2C, PAN, P3B, ID, OTRA_2C, URL, DOCENTES, ADMIN_ID, GENERADOR_ID, ZONA, isoBA };
