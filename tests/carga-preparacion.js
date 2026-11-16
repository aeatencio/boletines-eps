/**
 * Carga docente EPS (#32) · pruebas locales de la preparación de 2º cuatrimestre y anual (Carga_Preparacion.gs y
 * Carga_Preparacion_Google.gs), con todos los .gs reales de src/carga y los dobles de tests/carga-dobles.js.
 *
 * El escenario reproduce la situación institucional con datos ficticios y los encabezados reales de los fixtures:
 * grilla de 1º cuatrimestre; 3º bimestre con la matrícula vigente (altas, una baja, un cambio de orientación, un
 * nombre corregido); y dos borradores copiados de 1º cuatrimestre y editados a mano de forma distinta. El de
 * 2º cuatrimestre se «limpió» vaciando las calificaciones copiadas (como los reales); el de anual conserva casi todas.
 * En los dos, las materias de EPS4 sólo tienen celdas vaciadas: sin preparar bien esas celdas, no tendrían cursantes
 * ni fila en el borrador de asignaciones (el defecto de la primera vista previa real).
 *
 * Usa dobles de Google: no reemplaza la primera aplicación revisada sobre los borradores reales.
 *
 * Uso: node tests/carga-preparacion.js [corridas]   (corridas de la prueba aleatoria de invariantes; 40 por defecto)
 */
"use strict";
const vm = require("vm");
const D = require("./carga-dobles.js");
const { Hoja } = D;

let fallas = 0, total = 0;
function prueba(nombre, fn) {
  total++;
  try { fn(); console.log("OK    " + nombre); } catch (e) { fallas++; console.log("FALLA " + nombre + "\n      " + String(e && e.stack || e).split("\n").slice(0, 5).join("\n      ")); }
}
const corto = t => (t.length > 400 ? t.slice(0, 400) + "…" : t);
function igual(a, b, msg) { const x = JSON.stringify(a), y = JSON.stringify(b); if (x !== y) throw new Error((msg || "distinto") + ":\n      " + corto(x) + "\n    ≠ " + corto(y)); }
function cierto(x, msg) { if (!x) throw new Error(msg || "se esperaba verdadero"); }
function lanza(fn, re) { let e = null; try { fn(); } catch (x) { e = x; } if (!e || !re.test(e.message)) throw new Error("se esperaba un error " + re + (e ? ", fue: " + e.message : "")); return e; }

// ================================================================ escenario

const P1C = "1º cuatrimestre", P3B = "3º bimestre", P2C = "2º cuatrimestre", PAN = "Anual";
const ID = { [P1C]: "GRILLA_1C_ficticia_00000000001", [P3B]: "GRILLA_3B_ficticia_00000000001", [P2C]: "GRILLA_2C_ficticia_00000000001", [PAN]: "GRILLA_AN_ficticia_00000000001" };
const PROF = n => "prof." + n + "@ficticio.example";
const BORRADOR = "Asignaciones (borrador)";
const leerTsv = p => D.leer(p).replace(/\n$/, "").split("\n").map(l => l.split("\t"));
const ENC = {
  EPSI_1C: leerTsv("reference/fixtures/1er_cuatrimestre_estructura_real_EPSI.tsv").slice(0, 2),
  EPSI_3B: leerTsv("reference/fixtures/3er_bimestre_estructura_real_EPSI.tsv").slice(0, 2),
  EPSG: leerTsv("reference/fixtures/3er_bimestre_estructura_real_EPSG.tsv").slice(0, 2)
};
const ETIQUETAS = ["Valoración", "Calidad de las producciones:", "Participación:", "Observaciones:"];
const nrm = t => String(t).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();

/** Grilla con el encabezado real y los estudiantes dados: { grupo, dni, ap, no, v: { materia: valor | [cuatro campos] }, cola }. */
function grilla(C, enc, periodo, estudiantes) {
  const ancho = Math.max(enc[0].length, enc[1].length);
  const pad = f => Array.from({ length: ancho }, (_, i) => f[i] || "");
  const filas = [pad(enc[0]), pad(enc[1])];
  const est = C.estructuraGrilla_(filas, periodo);
  const colCola = h => filas[0].findIndex(x => nrm(x) === nrm(h));
  ["EPS1", "EPS2", "EPS3", "EPS4"].forEach(grupo => {
    const marcador = Array(ancho).fill(""); marcador[0] = grupo; filas.push(marcador);
    estudiantes.filter(e => e.grupo === grupo).forEach(e => {
      const b = [0, 1, 2, 3].map(() => Array(ancho).fill(""));
      b[0][0] = e.ap; b[0][1] = e.no; b[0][2] = e.dni || "";
      est.materias.forEach(m => ETIQUETAS.forEach((t, i) => { b[i][m.columnaValor - 1] = t; }));
      Object.keys(e.v || {}).forEach(nombre => {
        const m = est.materias.find(x => x.nombre === nombre);
        if (!m) throw new Error("La materia " + nombre + " no está en el encabezado de " + periodo);
        [].concat(e.v[nombre]).forEach((x, i) => { b[i][m.columnaValor] = x; });
      });
      Object.keys(e.cola || {}).forEach(h => { b[0][colCola(h)] = e.cola[h]; });
      filas.push(...b);
    });
  });
  return new Hoja(filas);
}

// La fórmula del correo del borrador, evaluada como Sheets: VLOOKUP exacto (sin distinguir mayúsculas) en la hoja Docentes.
let adminActual = null;
const verOriginal = Hoja.prototype.ver;
Hoja.prototype.ver = function (x) {
  const m = typeof x === "string" && /^=IFERROR\(VLOOKUP\(([A-Z]+)(\d+),Docentes!\$A:\$B,2,FALSE\),""\)$/.exec(x);
  if (!m) return verOriginal.call(this, x);
  const col = m[1].split("").reduce((n, l) => n * 26 + l.charCodeAt(0) - 64, 0) - 1;
  const buscado = String((this.v[Number(m[2]) - 1] || [])[col] || "").trim().toLowerCase();
  const nomina = adminActual && adminActual.Docentes;
  const fila = buscado && nomina ? nomina.v.find(f => String(f[0]).trim().toLowerCase() === buscado) : null;
  return fila ? String(fila[1]) : "";
};

const INF = { "ORIENTACIÓN": "Informática" }, GAS = { "ORIENTACIÓN": "Gastronomía" };
// 1º cuatrimestre (EPSI). Los borradores son copias de estos estudiantes, editadas.
const A = { grupo: "EPS1", dni: "30000001", ap: "Ficticia", no: "Ana", cola: Object.assign({ "LISTO PARA EXPORTAR": "SÍ", "EXPORTADO": "SÍ", "OBSERVACIONES GENERALES SOBRE DESEMPEÑO ACADÉMICO": "Obs general 1C" }, INF),
  v: { "Matemática I": ["7", "Buena", "Activa", "Obs 1C Ana"], "Lengua y Literatura I": ["8", "Muy buena"], "Taller Profesional I A": "6", "Taller Profesional I C": ["9", "Excelente"], "FGI I": "ACREDITÓ" } };
const B = { grupo: "EPS1", dni: "30000002", ap: "Ficticio", no: "Bruno", cola: Object.assign({ "LISTO PARA EXPORTAR": "SÍ" }, INF),
  v: { "Matemática I": ["4", "Regular", "Poca"], "Lengua y Literatura I": "EQUIVALENCIA", "Taller Profesional I A": "5", "Taller Profesional I B": "NO CURSA" } };
const Cc = { grupo: "EPS2", dni: "30000003", ap: "Ficticia", no: "Carla", cola: INF, v: { "Inglés I": "9", "Matemática II": "8", "Taller Profesional II A": "7", "Matemática I": "6" } };
const Dd = { grupo: "EPS2", dni: "30000004", ap: "Ficticio", no: "Dante", cola: INF, v: { "Matemática II": ["5", "Regular"], "Inglés I": "CURSA", "Taller Profesional II A": "NO CURSA" } };
const E = { grupo: "EPS3", dni: "30000005", ap: "Ficticia", no: "Elena", cola: INF, v: { "Matemática III": "7", "Taller Profesional III A": "8" } };
const F = { grupo: "EPS4", dni: "30000006", ap: "Ficticio", no: "Facundo", cola: INF, v: { "Matemática IV": "10", "Taller Profesional IV A": "9" } };
const L = { grupo: "EPS1", dni: "30000011", ap: "Ficticia", no: "Luz", cola: INF, v: { "Matemática I": "7" } };
const J = { grupo: "EPS1", dni: "31000001", ap: "Ficticio", no: "Joaquín", cola: GAS, v: { "Matemática I": ["6", "Buena"], "Taller Gastronómico I A": "7" } };
const K = { grupo: "EPS2", dni: "31000002", ap: "Ficticia", no: "Kiara", cola: GAS, v: { "Matemática II": "8", "Taller Gastronómico II A": "EQUIVALENCIA" } };
const G = { grupo: "EPS1", dni: "30000007", ap: "Ficticio", no: "Gael", cola: INF };
const con = (e, v, extra) => Object.assign({}, e, { v: Object.assign({}, e.v, v) }, extra || {});
// Limpieza de la copia: la calificación vaciada; Calidad, Participación y Observaciones a veces quedaron.
const vaciar = (e, materias, extra) => con(e, Object.fromEntries(materias.map(m => [m, [""].concat([].concat(e.v[m]).slice(1))])), extra);

function escenarioPrep(C) {
  const g = {};
  g[ID[P1C]] = { EPSI: grilla(C, ENC.EPSI_1C, P1C, [A, B, Cc, Dd, E, F, L]), EPSG: grilla(C, ENC.EPSG, P1C, [J, K]) };
  g[ID[P3B]] = {
    EPSI: grilla(C, ENC.EPSI_3B, P3B, [
      Object.assign({}, A, { v: { "Matemática I": ["Suficiente", "Buena"], "Lengua y Literatura I": "Avanzado", "Taller Profesional I A": "En proceso", "FGI I": "ACREDITÓ" } }),
      Object.assign({}, B, { v: { "Matemática I": "En proceso", "Lengua y Literatura I": "EQUIVALENCIA", "Taller Profesional I A": "NO CURSA", "Taller Profesional I B": "Suficiente" } }),
      Object.assign({}, Cc, { no: "Karla", v: { "Inglés I": "Suficiente", "Matemática II": "Avanzado", "Taller Profesional II A": "Suficiente", "Matemática I": "ACREDITÓ" } }),
      Object.assign({}, Dd, { v: { "Matemática II": "Suficiente", "Inglés I": "CURSA" } }),
      Object.assign({}, F, { v: { "Matemática IV": "Suficiente", "Taller Profesional IV A": "Avanzado" } }),
      Object.assign({}, G, { v: { "Matemática I": "Suficiente", "Lengua y Literatura I": "En proceso" } }),
      { grupo: "EPS3", dni: "30000008", ap: "Ficticia", no: "Helena", cola: INF, v: { "Matemática III": "Suficiente" } },
      { grupo: "EPS2", dni: "30000009", ap: "Ficticio", no: "Iván", cola: INF, v: { "Matemática II": "Suficiente" } }]),
    EPSG: grilla(C, ENC.EPSG, P3B, [Object.assign({}, J, { v: { "Matemática I": "Suficiente", "Taller Gastronómico I A": "Avanzado" } }),
      Object.assign({}, K, { v: { "Matemática II": "Suficiente", "Taller Gastronómico II A": "EQUIVALENCIA" } }),
      Object.assign({}, L, { cola: GAS, v: { "Matemática I": "Avanzado", "Taller Gastronómico I A": "Suficiente" } })])
  };
  const Lg = Object.assign({}, L, { cola: GAS, v: {} });
  const Fvaciado = vaciar(F, ["Matemática IV", "Taller Profesional IV A"]);
  g[ID[P2C]] = {
    EPSI: grilla(C, ENC.EPSI_1C, P2C, [
      vaciar(A, ["Matemática I", "Lengua y Literatura I", "Taller Profesional I A"]),
      con(vaciar(B, ["Taller Profesional I A"]), { "Matemática I": ["CURSA", "Regular", "Poca"] }),
      con(vaciar(Cc, ["Inglés I", "Taller Profesional II A", "Matemática I"]), { "Matemática II": "9" }),
      con(vaciar(Dd, ["Matemática II"]), { "Taller Profesional II A": "CURSA", "Matemática II": ["", "Regular", "", "Nota EPS"] }),
      Object.assign({}, Fvaciado, { grupo: "EPS3" }),
      Object.assign({}, G, { v: { "Matemática I": "CURSA" } }), { grupo: "EPS2", dni: "30000090", ap: "Ficticio", no: "Iván", cola: INF },
      { grupo: "EPS2", dni: "", ap: "Ficticia", no: "Sin Documento", v: { "Inglés I": "CURSA" } }]),
    EPSG: grilla(C, ENC.EPSG, P2C, [vaciar(J, ["Matemática I", "Taller Gastronómico I A"]), vaciar(K, ["Matemática II"]), Object.assign({}, Lg, { v: { "Matemática I": "NO CURSA" } })])
  };
  g[ID[PAN]] = {
    EPSI: grilla(C, ENC.EPSI_1C, PAN, [A, B, Cc, Dd, E, Fvaciado, Object.assign({}, G, { v: { "Matemática I": "CURSA" } })]),
    EPSG: grilla(C, ENC.EPSG, PAN, [J, K, Lg])
  };
  const control = Array.from({ length: 14 }, () => ["", ""]);
  control[10][1] = D.URL(ID[P1C]); control[11][1] = ID[P3B]; control[12][1] = D.URL(ID[P2C]); control[13][1] = D.URL(ID[PAN]);
  return {
    hoy: Date.now(), grillas: g, control: new Hoja(control), lock: D.nuevoLock(),
    admin: {
      Ventanas: new Hoja([["Período", "Desde", "Hasta"], ["3er bimestre", "2026-09-14", "2026-09-25"]]),
      Asignaciones: new Hoja([["Orientación", "Nivel", "Materia", "Docente", "Correo"], ["EPSI", "EPS1", "Lengua y Literatura I", "Cuenta de prueba", PROF("prueba")]]),
      Enlaces: new Hoja([["Correo", "Versión"]]),
      Registro: new Hoja([vm.runInContext("CARGA", C).REGISTRO.slice()]),
      Correos: new Hoja([vm.runInContext("CORREO", C).COLUMNAS.slice()]),
      Docentes: new Hoja([["Docente", "Correo"], ["Prof. Mate Ficticio", PROF("mate")], ["Prof. Lengua Ficticia", PROF("lengua")], ["Prof. Taller Ficticio", PROF("taller")],
        ["Prof. Cocina Ficticia", PROF("cocina")]]),
      [BORRADOR]: new Hoja([["Orientación", "Nivel", "Materia", "Docente", "Correo"],
        ["EPSI", "EPS2", "Matemática II", "Prof. Mate Ficticio", PROF("mate")],
        ["EPSI", "EPS3", "Inglés I", "Prof. Inglés Ficticia", PROF("ingles")],
        ["EPSG", "EPS1", "Panadería", "Docente Ficticio", "otro@ficticio.example"]])
    }
  };
}

function entorno(ajustar) {
  const S = D.servidor();
  S.estado.s = escenarioPrep(S.ctx);
  if (ajustar) ajustar(S.estado.s, S.ctx);
  adminActual = S.estado.s.admin;
  return { S, ctx: S.ctx, s: S.estado.s };
}
const copia = h => JSON.stringify(h.v);
const fotos = s => JSON.stringify(Object.keys(s.grillas).map(id => [id, s.grillas[id].EPSI.v, s.grillas[id].EPSG.v]));
const tabla = h => h.getDataRange().getDisplayValues();
function celda(ctx, s, periodo, hoja, materia, dni, desp) {
  const v = tabla(s.grillas[ID[periodo]][hoja]);
  const m = ctx.estructuraGrilla_(v, periodo).materias.find(x => x.nombre === materia);
  const fila = v.findIndex(f => f[2] === dni);
  if (!m || fila < 0) throw new Error("No se ubicó " + materia + " / " + dni);
  return v[fila + (desp || 0)][m.columnaValor];
}
function poner(ctx, s, periodo, hoja, materia, dni, valor) {
  const h = s.grillas[ID[periodo]][hoja];
  const v = tabla(h);
  const m = ctx.estructuraGrilla_(v, periodo).materias.find(x => x.nombre === materia);
  h.v[v.findIndex(f => f[2] === dni)][m.columnaValor] = valor;
}
const informe = s => tabla(s.admin["Preparación"]);
const avisos = s => {
  const f = informe(s);
  const i = f.findIndex(x => x[0] === "Prioridad" && x[9] === "Detalle");
  return f.slice(i + 1).map(x => ({ prioridad: x[0], tipo: x[1], periodo: x[2], hoja: x[3], nivel: x[4], materia: x[5], fila: x[6], dni: x[7], estudiante: x[8], detalle: x[9] }));
};
const porTipo = s => {
  const f = informe(s);
  const i = f.findIndex(x => x[0] === "Prioridad" && x[2] === "Casos");
  const r = [];
  for (let k = i + 1; k < f.length && f[k][0]; k++) r.push({ prioridad: f[k][0], tipo: f[k][1], casos: Number(f[k][2]), accion: f[k][3], ejemplos: f[k][4] });
  return r;
};
const hay = (lista, cond, msg) => { const r = lista.filter(cond); cierto(r.length, msg || "no se encontró el aviso esperado"); return r; };
const borrador = s => tabla(s.admin[BORRADOR]);
const filaB = (s, o, m) => borrador(s).find(f => f[0] === o && f[2] === m);

// ================================================================ pruebas

console.log("\n# Decisión por celda (reglas de #32)");
prueba("vacía y NO CURSA con calificación ordinaria en 3B → CURSA, aunque en 1C tuviera nota (la limpieza de la copia no es dejar de cursar)", () => {
  const { ctx } = entorno();
  const d = (D0, C0, E0, quitada) => ctx.decidirCalificacionPrep_(P2C, D0, C0, E0, "", quitada);
  ["", "NO CURSA"].forEach(D0 => {
    [null, "", "NO CURSA", "7", "10", "CURSA", "ACREDITÓ", "EQUIVALENCIA"].forEach(C0 => {
      ["Suficiente", "En proceso", "Avanzado"].forEach(E0 => igual(!!d(D0, C0, E0).cursa, true, "«" + D0 + "» / 1C «" + C0 + "» / 3B «" + E0 + "»"));
      // sin calificación ordinaria en 3B: nada; ACREDITÓ/EQUIVALENCIA en 3B: se informa, nunca se copia ni da CURSA
      ["", "NO CURSA", "CURSA"].forEach(E0 => igual(d(D0, C0, E0), {}));
      ["ACREDITÓ", "EQUIVALENCIA"].forEach(E0 => igual([!!d(D0, C0, E0).cursa, d(D0, C0, E0).tipo], [false, "administrativo3b"]));
      // una cursada que la preparación ya puso y alguien quitó: se respeta
      igual([!!d(D0, C0, "Suficiente", true).cursa, d(D0, C0, "Suficiente", true).tipo], [false, "quitadaDespues"]);
    });
  });
  igual(/se había vaciado/.test(d("", "7", "Suficiente").motivo), true, d("", "7", "Suficiente").motivo);
});
prueba("notas, CURSA y estados administrativos del borrador: sólo una nota copiada con evidencia pasa a CURSA; nada se convierte en valoración nueva", () => {
  const { ctx } = entorno();
  const d = (D0, C0, E0, falla) => ctx.decidirCalificacionPrep_(P2C, D0, C0, E0, falla || "");
  igual(!!d("7", "7", "Suficiente").cursa, true);
  ["ACREDITÓ", "EQUIVALENCIA", "CURSA", "", "NO CURSA", "Muy bien"].forEach(e => igual([!!d("7", "7", e).cursa, d("7", "7", e).tipo], [false, "notaSinEvidencia"], "7 copiado y 3B «" + e + "»"));
  igual(d("8", "7", "Suficiente").tipo, "notaPropia");
  igual(d("8", null, "Suficiente").tipo, "notaPropia");
  igual(d("7,00", "7", "Suficiente").tipo, "valorNoReconocido");
  igual(d("CURSA", "7", "Suficiente"), {});
  igual(d("CURSA", "", "ACREDITÓ").tipo, "cursaSinEvidencia");
  igual(d("CURSA", "", "").tipo, "cursaSinEvidencia");
  igual(d("ACREDITÓ", "ACREDITÓ", "ACREDITÓ"), {});
  igual(d("EQUIVALENCIA", "", "Suficiente").tipo, "administrativoBorrador");
  // sin evidencia comparable (estudiante o materia no identificados): nada cambia
  igual(d("7", "7", null, "sin estudiante"), { sinPreparar: true });
  igual(d("", "7", null, "otro nivel"), { sinPreparar: false });
  // Anual: la misma regla, con 3º bimestre (nunca con 2º cuatrimestre)
  igual(!!ctx.decidirCalificacionPrep_(PAN, "", "7", "Suficiente", "").cursa, true);
});

console.log("\n# Vista previa y aplicación sobre el escenario institucional ficticio");
prueba("vista previa: escribe sólo el informe; no toca grillas, borrador, nómina, Asignaciones, Ventanas, Registro ni Correos", () => {
  const { ctx, s } = entorno();
  const nombres = [BORRADOR, "Docentes", "Asignaciones", "Ventanas", "Enlaces", "Registro", "Correos"];
  const antes = fotos(s), adm = nombres.map(n => copia(s.admin[n]));
  const r = ctx.previsualizarPeriodosCarga();
  igual([r.aplicar, r.puedeAplicar, r.aplicacion], [false, true, null]);
  cierto(r.cambios > 0);
  igual(fotos(s), antes, "grillas");
  igual(nombres.map(n => copia(s.admin[n])), adm);
  igual(s.admin["Preparación · cambios"], undefined);
  cierto(/VISTA PREVIA/.test(informe(s)[0][0]), informe(s)[0][0]);
});

prueba("aplicación: las celdas vaciadas al limpiar la copia pasan a CURSA con evidencia de 3B; se conservan estados administrativos, notas propias y cambios humanos", () => {
  const { ctx, s } = entorno();
  const r = ctx.prepararPeriodosCarga();
  igual([r.aplicacion.error, r.aplicacion.revisar, r.aplicacion.descartados], ["", 0, 0]);
  igual(r.aplicacion.aplicados, r.cambios);
  const c = (p, h, m, dni, desp) => celda(ctx, s, p, h, m, dni, desp);
  // 2C (limpiado): vaciadas con calificación en 3B → CURSA; complementarios copiados, borrados; ACREDITÓ se conserva.
  igual([c(P2C, "EPSI", "Matemática I", A.dni), c(P2C, "EPSI", "Matemática I", A.dni, 1), c(P2C, "EPSI", "Matemática I", A.dni, 3), c(P2C, "EPSI", "Lengua y Literatura I", A.dni),
    c(P2C, "EPSI", "Lengua y Literatura I", A.dni, 1), c(P2C, "EPSI", "Taller Profesional I A", A.dni), c(P2C, "EPSI", "FGI I", A.dni)], ["CURSA", "", "", "CURSA", "", "CURSA", "ACREDITÓ"]);
  // Taller Profesional I C no está en 3B: su 9 queda (Decidir); su Calidad copiada se borra.
  igual([c(P2C, "EPSI", "Taller Profesional I C", A.dni), c(P2C, "EPSI", "Taller Profesional I C", A.dni, 1)], ["9", ""]);
  // B: vaciado y NO CURSA en 3B → sigue sin cursar; EQUIVALENCIA se conserva; NO CURSA con 3B → CURSA.
  igual([c(P2C, "EPSI", "Matemática I", B.dni), c(P2C, "EPSI", "Matemática I", B.dni, 1), c(P2C, "EPSI", "Lengua y Literatura I", B.dni), c(P2C, "EPSI", "Taller Profesional I A", B.dni),
    c(P2C, "EPSI", "Taller Profesional I B", B.dni)], ["CURSA", "", "EQUIVALENCIA", "", "CURSA"]);
  // C: nota propia 9 se conserva; vaciada con ACREDITÓ en 3B queda vacía (no se copian estados); las demás, CURSA.
  igual([c(P2C, "EPSI", "Matemática II", Cc.dni), c(P2C, "EPSI", "Matemática I", Cc.dni), c(P2C, "EPSI", "Inglés I", Cc.dni), c(P2C, "EPSI", "Taller Profesional II A", Cc.dni)], ["9", "", "CURSA", "CURSA"]);
  // D: vaciada → CURSA; Observaciones del Equipo EPS y CURSA puesto a mano se conservan.
  igual([c(P2C, "EPSI", "Matemática II", Dd.dni), c(P2C, "EPSI", "Matemática II", Dd.dni, 1), c(P2C, "EPSI", "Matemática II", Dd.dni, 3), c(P2C, "EPSI", "Taller Profesional II A", Dd.dni)],
    ["CURSA", "", "Nota EPS", "CURSA"]);
  // F: sus dos materias de EPS4 sólo tenían celdas vaciadas, en los dos borradores → CURSA en los dos.
  igual([c(P2C, "EPSI", "Matemática IV", F.dni), c(P2C, "EPSI", "Taller Profesional IV A", F.dni), c(PAN, "EPSI", "Matemática IV", F.dni), c(PAN, "EPSI", "Taller Profesional IV A", F.dni)],
    ["CURSA", "CURSA", "CURSA", "CURSA"]);
  // G (alta) y L (cambió a EPSG; NO CURSA en 2C, vacía en anual).
  igual([c(P2C, "EPSI", "Lengua y Literatura I", G.dni), c(P2C, "EPSG", "Matemática I", L.dni), c(PAN, "EPSG", "Matemática I", L.dni), c(PAN, "EPSG", "Taller Gastronómico I A", L.dni)],
    ["CURSA", "CURSA", "CURSA", "CURSA"]);
  // EPSG de 2C limpiado: vaciadas → CURSA; EQUIVALENCIA se conserva.
  igual([c(P2C, "EPSG", "Matemática I", J.dni), c(P2C, "EPSG", "Matemática I", J.dni, 1), c(P2C, "EPSG", "Matemática II", K.dni), c(P2C, "EPSG", "Taller Gastronómico II A", K.dni)],
    ["CURSA", "", "CURSA", "EQUIVALENCIA"]);
  // Anual (no limpiado): notas copiadas con evidencia → CURSA; independencia de 2C (D sigue sin Taller II A).
  igual([c(PAN, "EPSI", "Matemática I", B.dni), c(PAN, "EPSI", "Matemática II", Cc.dni), c(PAN, "EPSI", "Taller Profesional II A", Dd.dni), c(PAN, "EPSI", "Matemática III", E.dni), c(PAN, "EPSI", "Matemática I", Cc.dni)],
    ["CURSA", "CURSA", "NO CURSA", "7", "6"]);
  // Nada fuera de las columnas de materias; 1C y 3B sólo se leen.
  igual(s.grillas[ID[P2C]].EPSI.escrituras.filter(w => w.col <= 2), []);
  igual(tabla(s.grillas[ID[P2C]].EPSI).find(f => f[2] === A.dni).slice(41).filter(Boolean), ["Informática", "Obs general 1C", "SÍ", "SÍ"], "campos generales y operativos intactos");
  igual([s.grillas[ID[P1C]].EPSI.intentos, s.grillas[ID[P3B]].EPSI.intentos, s.grillas[ID[P3B]].EPSG.intentos], [0, 0, 0]);
});

prueba("aplicación: cada celda queda en «Preparación · cambios» con anterior, nuevo, motivo y estado verificado", () => {
  const { ctx, s } = entorno();
  const r = ctx.prepararPeriodosCarga();
  const log = tabla(s.admin["Preparación · cambios"]);
  igual(log[0], ["Fecha", "Estado", "Detalle", "Período", "Grilla", "Hoja", "Celda", "DNI", "Estudiante", "Nivel", "Materia", "Campo", "Anterior", "Nuevo", "Motivo"]);
  igual(log.length - 1, r.cambios);
  cierto(log.slice(1).every(f => f[1] === "aplicado"), "todas aplicadas");
  const a = log.find(f => f[3] === P2C && f[7] === A.dni && f[10] === "Matemática I" && f[11] === "Calificación");
  igual([a[4], a[5], a[12], a[13]], [ID[P2C], "EPSI", "", "CURSA"]);
  cierto(/en 3º bimestre tiene «Suficiente»; en 1º cuatrimestre tenía «7», que se había vaciado/i.test(a[14]), a[14]);
  const b = log.find(f => f[3] === PAN && f[7] === A.dni && f[10] === "Matemática I" && f[11] === "Calificación");
  igual([b[12], b[13]], ["7", "CURSA"]);
});

prueba("informe: un renglón por tipo de caso con cantidad y qué hacer; ningún caso por las celdas vaciadas con evidencia; los excepcionales siguen visibles", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  const cab = informe(s).map(f => f[0]);
  cierto(/APLICADA/.test(cab[0]), cab[0]);
  cierto(cab.some(l => /^2º cuatrimestre: 10 estudiantes \(EPSI 7, EPSG 3\)\. Se pusieron \d+ CURSA con evidencia de 3º bimestre \(0 en lugar de notas copiadas de 1º cuatrimestre, \d+ en celdas que se habían vaciado/.test(l)),
    cab.slice(0, 6).join(" | "));
  const x = avisos(s), t = porTipo(s);
  // El resumen por tipo cuenta exactamente los casos del detalle y está ordenado por prioridad.
  igual(t.map(y => y.casos).reduce((a, b) => a + b, 0), x.length);
  t.forEach(y => igual(x.filter(z => z.tipo === y.tipo).length, y.casos, y.tipo));
  const orden = t.map(y => ["Bloqueo", "Decidir", "Revisar"].indexOf(y.prioridad));
  igual(orden.slice().sort((p, q) => p - q), orden);
  cierto(t.every(y => y.accion), "cada tipo dice qué hacer");
  // Lo que antes eran cientos de «Revisar» (vacía con evidencia) ya no existe como caso.
  igual(x.filter(z => /Vacía en el borrador .*pero en 3º bimestre tiene/.test(z.detalle)), []);
  // Los casos excepcionales siguen, con su tipo.
  hay(x, a => a.tipo === "Calificación en el borrador que no viene de 1º cuatrimestre" && a.dni === Cc.dni && a.periodo === P2C);
  hay(x, a => a.tipo === "Calificación de 1º cuatrimestre sin evidencia de cursada en 3º bimestre" && a.dni === Cc.dni && a.periodo === PAN && /figura «ACREDITÓ»/.test(a.detalle));
  hay(x, a => a.tipo === "Calificación de 1º cuatrimestre sin evidencia de cursada en 3º bimestre" && a.dni === B.dni && a.periodo === PAN && a.materia === "Taller Profesional I A");
  hay(x, a => a.prioridad === "Revisar" && a.tipo === "Vacía en el borrador; ACREDITÓ o EQUIVALENCIA en 3º bimestre" && a.dni === Cc.dni && a.periodo === P2C);
  hay(x, a => a.tipo === "Materia sin evidencia en 3º bimestre (falta o tiene otro nivel)" && a.materia === "Taller Profesional I C");
  hay(x, a => a.tipo === "Fila con nombre y sin DNI");
  hay(x, a => a.tipo === "Estudiante no identificado en 3º bimestre, con calificaciones de 1º cuatrimestre" && a.dni === E.dni);
  hay(x, a => a.dni === "30000090" && /DNI 30000009 \(¿DNI mal escrito\?\)/.test(a.detalle));
  hay(x, a => a.dni === "30000008" && a.tipo === "Diferencia de matrícula con 3º bimestre");
  hay(x, a => a.dni === Cc.dni && a.tipo === "Nombre, grupo u orientación distintos de 3º bimestre o 1º cuatrimestre" && /Karla/.test(a.detalle));
  hay(x, a => a.dni === F.dni && /Grupo EPS3 en el borrador y EPS4/.test(a.detalle));
  hay(x, a => a.dni === Dd.dni && a.tipo === "CURSA en el borrador sin calificación ordinaria en 3º bimestre");
  hay(x, a => a.dni === Dd.dni && a.tipo === "Calidad, Participación u Observaciones que no vienen de 1º cuatrimestre");
  hay(x, a => a.tipo === "Diferencia entre 2º cuatrimestre y Anual" && a.dni === Dd.dni && /Cursa en 2º cuatrimestre \(«CURSA»\) y no en Anual/.test(a.detalle));
  hay(x, a => a.tipo === "Campos generales u operativos iguales a 1º cuatrimestre");
  // Asignaciones: un solo caso agrupado para las materias sin docente; las filas inconsistentes, una por una.
  igual(x.filter(a => a.tipo === "Materias sin docente elegido").length, 1);
  hay(x, a => a.tipo === "Fila de asignación incompleta o inconsistente" && a.materia === "Inglés I" && /no está en la hoja Docentes/.test(a.detalle) && /EPS2/.test(a.detalle));
  hay(x, a => a.tipo === "Fila de asignación incompleta o inconsistente" && a.materia === "Panadería");
  igual(x.filter(a => a.prioridad === "Bloqueo"), []);
  igual(x.filter(a => a.dni === L.dni), [], "L cambió de orientación de forma consistente");
});

console.log("\n# Asignaciones");
prueba("todas las materias con cursantes tienen fila (incluidas las que sólo tenían celdas vaciadas); ninguna sin cursantes", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  // Materias con algún cursante en las grillas preparadas, calculadas aparte.
  const esperadas = new Set();
  [P2C, PAN].forEach(p => ["EPSI", "EPSG"].forEach(h => {
    const v = tabla(s.grillas[ID[p]][h]);
    const est = ctx.estructuraGrilla_(v, p);
    est.materias.forEach(m => { if (v.some((f, i) => i > 1 && f[2] && !["", "NO CURSA"].includes(f[m.columnaValor]))) esperadas.add(h + "|" + m.nombre); });
  }));
  const b = borrador(s).slice(1);
  const enBorrador = new Set(b.map(f => f[0] + "|" + f[2]));
  esperadas.forEach(k => cierto(enBorrador.has(k), "falta la fila de " + k));
  b.filter(f => !esperadas.has(f[0] + "|" + f[2])).forEach(f => cierto(f[2] === "Panadería", "fila sin cursantes: " + f.join(" | ")));
  // Las de EPS4 sólo existen porque se prepararon las celdas vaciadas.
  cierto(enBorrador.has("EPSI|Matemática IV") && enBorrador.has("EPSI|Taller Profesional IV A"));
  igual(b.filter(f => f[0] === "EPSI" && f[2] === "Inglés I").length, 1, "no se duplica una materia con fila");
  // Orientación, nivel y materia salen de la grilla.
  igual(filaB(s, "EPSI", "Matemática IV").slice(0, 3), ["EPSI", "EPS4", "Matemática IV"]);
  igual(filaB(s, "EPSG", "Taller Gastronómico I A").slice(0, 3), ["EPSG", "EPS1", "Taller Gastronómico I A"]);
});

prueba("el Equipo EPS elige el docente de la nómina y el correo sale solo, consistente; las filas existentes no se tocan", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  const h = s.admin[BORRADOR];
  igual(borrador(s)[0], ["Orientación", "Nivel", "Materia", "Docente", "Correo", "Cursantes", "Revisión"]);
  // Filas existentes: sus cinco columnas, intactas; la consistente no tiene nada para revisar.
  igual(borrador(s).slice(1, 4).map(f => f.slice(0, 5)), [["EPSI", "EPS2", "Matemática II", "Prof. Mate Ficticio", PROF("mate")], ["EPSI", "EPS3", "Inglés I", "Prof. Inglés Ficticia", PROF("ingles")],
    ["EPSG", "EPS1", "Panadería", "Docente Ficticio", "otro@ficticio.example"]]);
  igual(filaB(s, "EPSI", "Matemática II").slice(5), ["2º cuatrimestre: 2 · Anual: 2", ""]);
  // Filas nuevas: docente vacío, correo por fórmula (vacío hasta elegir) y la lista de la hoja Docentes en toda la columna.
  const nueva = filaB(s, "EPSI", "Matemática I");
  igual(nueva.slice(3, 5), ["", ""]);
  cierto(/^Falta elegir el docente\.$/.test(nueva[6]), nueva[6]);
  const fila = borrador(s).findIndex(f => f[0] === "EPSI" && f[2] === "Matemática I") + 1;
  igual(h.v[fila - 1][4], '=IFERROR(VLOOKUP(D' + fila + ',Docentes!$A:$B,2,FALSE),"")');
  const val = h.validaciones.slice(-1)[0];
  igual([val.fila, val.col, val.filas, val.regla.rango.a1, val.regla.rango.hoja === s.admin.Docentes, val.regla.permiteInvalidos], [2, 4, borrador(s).length - 1, "A2:A", true, false]);
  cierto(/FGI I/.test(filaB(s, "EPSI", "FGI I")[2]) && /Sólo ACREDITÓ o EQUIVALENCIA/.test(filaB(s, "EPSI", "FGI I")[6]));
  // Elegir docentes: el correo aparece por la fórmula; el mismo docente puede tener varias materias.
  h.v[fila - 1][3] = "Prof. Mate Ficticio";
  const f4 = borrador(s).findIndex(f => f[2] === "Matemática IV");
  h.v[f4][3] = "Prof. Mate Ficticio";
  igual([filaB(s, "EPSI", "Matemática I")[4], filaB(s, "EPSI", "Matemática IV")[4]], [PROF("mate"), PROF("mate")]);
  // Un docente que no está en la nómina, o un correo pisado a mano, quedan señalados en la siguiente ejecución.
  const fl = borrador(s).findIndex(f => f[2] === "Lengua y Literatura I");
  h.v[fl][3] = "Prof. Inexistente";
  const fc = borrador(s).findIndex(f => f[2] === "Taller Profesional I A");
  h.v[fc][3] = "Prof. Taller Ficticio"; h.v[fc][4] = PROF("otro");
  const antes = borrador(s).length;
  ctx.prepararPeriodosCarga();
  igual(borrador(s).length, antes, "no se duplican filas");
  igual([filaB(s, "EPSI", "Matemática I")[3], filaB(s, "EPSI", "Matemática I")[4], filaB(s, "EPSI", "Matemática I")[6]], ["Prof. Mate Ficticio", PROF("mate"), ""]);
  cierto(/no está en la hoja Docentes/.test(filaB(s, "EPSI", "Lengua y Literatura I")[6]));
  cierto(/El correo no es el de Docentes \(prof\.taller@ficticio\.example\)/.test(filaB(s, "EPSI", "Taller Profesional I A")[6]), filaB(s, "EPSI", "Taller Profesional I A")[6]);
  igual(h.v[fc][4], PROF("otro"), "el correo escrito a mano no se pisa");
  hay(avisos(s), a => a.tipo === "Fila de asignación incompleta o inconsistente" && a.materia === "Lengua y Literatura I");
  const sin = avisos(s).find(a => a.tipo === "Materias sin docente elegido");
  const total = borrador(s).slice(1).filter(f => !f[3] && !/Sin cursantes|no figura/.test(f[6])).length;
  cierto(sin && new RegExp("^" + total + " materias con cursantes no tienen docente").test(sin.detalle), sin && sin.detalle);
  // La hoja Asignaciones (la que lee la carga) no se tocó: no se habilita nada.
  igual(tabla(s.admin.Asignaciones).length, 2);
});

prueba("sin hoja Docentes ni borrador: los crea (la nómina vacía, para completar una vez); nómina con encabezados distintos: el borrador no se toca, las grillas sí", () => {
  const sin = entorno(s => { delete s.admin.Docentes; delete s.admin[BORRADOR]; });
  sin.ctx.prepararPeriodosCarga();
  igual(tabla(sin.s.admin.Docentes), [["Docente", "Correo"]]);
  igual(borrador(sin.s)[0], ["Orientación", "Nivel", "Materia", "Docente", "Correo", "Cursantes", "Revisión"]);
  cierto(/La hoja Docentes está vacía/.test(avisos(sin.s).find(a => a.tipo === "Materias sin docente elegido").detalle));
  const mal = entorno(s => { s.admin.Docentes.v[0] = ["Correo", "Docente"]; });
  const b = copia(mal.s.admin[BORRADOR]);
  const r = mal.ctx.prepararPeriodosCarga();
  igual(copia(mal.s.admin[BORRADOR]), b);
  cierto(/no se modificó: La hoja Docentes debe tener/.test(r.aplicacion.borrador), r.aplicacion.borrador);
  cierto(r.aplicacion.aplicados > 0);
  hay(avisos(mal.s), a => a.prioridad === "Bloqueo" && a.tipo === "Borrador de asignaciones o nómina con columnas no reconocidas");
  const rep = entorno(s => { s.admin.Docentes.v.push(["prof. mate ficticio", PROF("mate2")], ["", PROF("x")], ["Sin Correo", ""]); });
  rep.ctx.previsualizarPeriodosCarga();
  igual(avisos(rep.s).filter(a => a.tipo === "Nómina de docentes con datos repetidos o incompletos").length, 3);
});

console.log("\n# Repetición y trabajo humano");
prueba("repetir: no cambia nada ni duplica filas; respeta decisiones posteriores (cursada quitada, nota decidida, docente elegido)", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  const grillas = fotos(s);
  const r2 = ctx.prepararPeriodosCarga();
  igual([r2.cambios, r2.aplicacion.aplicados, fotos(s)], [0, 0, grillas]);
  // Después de preparar, el Equipo EPS: quita una cursada (vacía) y otra (NO CURSA), decide el 9 de Taller I C, elige un docente.
  poner(ctx, s, P2C, "EPSI", "Matemática I", A.dni, "");
  poner(ctx, s, PAN, "EPSI", "Matemática IV", F.dni, "NO CURSA");
  poner(ctx, s, P2C, "EPSI", "Taller Profesional I C", A.dni, "NO CURSA");
  const h = s.admin[BORRADOR];
  const fl = borrador(s).findIndex(f => f[2] === "Lengua y Literatura I");
  h.v[fl][3] = "Prof. Lengua Ficticia";
  const filas = borrador(s).length, despues = fotos(s);
  const r3 = ctx.prepararPeriodosCarga();
  igual([r3.cambios, fotos(s), borrador(s).length], [0, despues, filas]);
  igual([celda(ctx, s, P2C, "EPSI", "Matemática I", A.dni), celda(ctx, s, PAN, "EPSI", "Matemática IV", F.dni)], ["", "NO CURSA"]);
  igual(borrador(s)[fl].slice(3, 5), ["Prof. Lengua Ficticia", PROF("lengua")]);
  const x = avisos(s);
  hay(x, a => a.tipo === "Cursada que la preparación puso y después se quitó" && a.dni === A.dni && a.periodo === P2C && a.materia === "Matemática I");
  hay(x, a => a.tipo === "Cursada que la preparación puso y después se quitó" && a.dni === F.dni && a.periodo === PAN);
  cierto(!x.some(a => a.materia === "Taller Profesional I C" && a.periodo === P2C && a.prioridad === "Decidir"), "el caso decidido ya no figura");
  hay(x, a => a.dni === Cc.dni && a.materia === "Matemática II" && a.periodo === P2C, "lo no decidido sigue informado");
});

console.log("\n# Salvaguardas");
prueba("con una ventana cargada para 2º cuatrimestre o anual no se modifica nada; la vista previa informa el bloqueo", () => {
  const { ctx, s } = entorno(s => { s.admin.Ventanas.v.push(["2do cuatrimestre", "2026-12-01", "2026-12-04"]); });
  const antes = fotos(s), b = copia(s.admin[BORRADOR]);
  lanza(() => ctx.prepararPeriodosCarga(), /no se aplicó.*Ventanas, fila 3: hay una ventana de carga de 2º cuatrimestre/);
  igual([fotos(s), copia(s.admin[BORRADOR]), s.admin["Preparación · cambios"]], [antes, b, undefined]);
  cierto(/NO SE APLICÓ/.test(informe(s)[0][0]));
  igual(ctx.previsualizarPeriodosCarga().puedeAplicar, false);
  igual(avisos(s)[0].prioridad, "Bloqueo");
});

prueba("sin 3º bimestre en Control (o con la misma grilla que un borrador): bloqueo general, nada modificado", () => {
  const { ctx, s } = entorno(s => { s.control.v[11][1] = ""; });
  const antes = fotos(s);
  lanza(() => ctx.prepararPeriodosCarga(), /Falta la grilla de 3º bimestre en Control!B12.*evidencia de cursada/);
  igual(fotos(s), antes);
  const otro = entorno(s => { s.control.v[13][1] = D.URL(ID[P2C]); });
  lanza(() => otro.ctx.prepararPeriodosCarga(), /misma que la de/);
  igual(otro.s.grillas[ID[P2C]].EPSI.escrituras.length, 0);
});

prueba("estructura no reconocida en una hoja de un borrador: esa hoja no se toca; la otra sí se prepara", () => {
  const { ctx, s } = entorno(s => { const h = s.grillas[ID[P2C]].EPSI; h.v[0][h.v[0].length - 1] = "ORIENTACIÓN"; });
  igual(ctx.prepararPeriodosCarga().aplicacion.error, "");
  igual(s.grillas[ID[P2C]].EPSI.escrituras.length, 0);
  cierto(s.grillas[ID[P2C]].EPSG.escrituras.length > 0);
  hay(avisos(s), a => a.prioridad === "Bloqueo" && a.periodo === P2C && a.hoja === "EPSI" && /ORIENTACIÓN.*No se prepara esta hoja/.test(a.detalle));
});

prueba("DNI repetido o bloque incompleto en el borrador: ese estudiante no se prepara y queda como bloqueo", () => {
  const { ctx, s } = entorno(s => {
    const h = s.grillas[ID[PAN]].EPSI;
    h.v[h.v.findIndex(f => f[2] === B.dni)][2] = A.dni;
    h.v[h.v.findIndex(f => f[2] === F.dni) + 2][4] = "Otra etiqueta";
  });
  ctx.prepararPeriodosCarga();
  const log = tabla(s.admin["Preparación · cambios"]).slice(1).filter(f => f[3] === PAN && f[5] === "EPSI");
  igual(log.filter(f => f[7] === A.dni || f[7] === F.dni), []);
  hay(avisos(s), a => a.prioridad === "Bloqueo" && a.periodo === PAN && a.dni === A.dni && /figura 2 veces/.test(a.detalle));
  hay(avisos(s), a => a.prioridad === "Bloqueo" && a.periodo === PAN && a.dni === F.dni && /etiqueta esperada/.test(a.detalle));
  cierto(log.length > 0, "los demás se preparan");
});

prueba("celdas en las que la carga docente ya escribió (Registro) no se tocan, aunque estén vacías y haya evidencia", () => {
  const { ctx, s } = entorno((s, C) => {
    const R = vm.runInContext("CARGA", C).REGISTRO;
    s.admin.Registro.v.push(R.map(c => ({ "Fecha": "2026-11-17", "Estado": "aplicado", "Período": P2C, "Fuente": ID[P2C], "Orientación": "EPSI", "DNI": A.dni, "Materia": "Matemática I", "Campo": "Calificación" }[c] || "")));
  });
  ctx.prepararPeriodosCarga();
  igual([celda(ctx, s, P2C, "EPSI", "Matemática I", A.dni), celda(ctx, s, P2C, "EPSI", "Matemática I", A.dni, 1)], ["", "Buena"]);
  igual(celda(ctx, s, PAN, "EPSI", "Matemática I", A.dni), "CURSA", "anual es otra grilla");
  hay(avisos(s), a => a.dni === A.dni && a.materia === "Matemática I" && a.periodo === P2C && a.tipo === "Celda en la que la carga docente ya escribió");
});

prueba("una celda que el Equipo EPS cambia entre la lectura y la escritura no se pisa: queda «no aplicado»", () => {
  const { ctx, s } = entorno();
  const h = s.grillas[ID[P2C]].EPSI;
  const original = h.getDataRange.bind(h);
  let lecturas = 0;
  h.getDataRange = () => {
    if (++lecturas === 2) poner(ctx, s, P2C, "EPSI", "Matemática I", A.dni, "NO CURSA");
    return original();
  };
  const r = ctx.prepararPeriodosCarga();
  igual(r.aplicacion.descartados, 1);
  igual(celda(ctx, s, P2C, "EPSI", "Matemática I", A.dni), "NO CURSA");
  const log = tabla(s.admin["Preparación · cambios"]).find(f => f[1] === "no aplicado");
  cierto(log && log[7] === A.dni && /ahora dice «NO CURSA»/.test(log[2]), JSON.stringify(log));
});

prueba("si Sheets falla al escribir, las celdas quedan «revisar» en el registro de cambios y la ejecución lo informa sin cortarse", () => {
  const { ctx, s } = entorno();
  s.grillas[ID[PAN]].EPSG.getRangeList = () => ({ setValue: () => { throw new Error("Servicio no disponible (simulado)"); }, clearContent: () => { throw new Error("Servicio no disponible (simulado)"); } });
  const r = ctx.prepararPeriodosCarga();
  cierto(r.aplicacion.revisar > 0 && r.aplicacion.aplicados > 0, JSON.stringify(r.aplicacion));
  const rev = tabla(s.admin["Preparación · cambios"]).filter(f => f[1] === "revisar");
  cierto(rev.length === r.aplicacion.revisar && rev.every(f => f[3] === PAN && f[5] === "EPSG" && /Servicio no disponible/.test(f[2])));
});

prueba("sólo la cuenta propietaria puede ejecutarlas", () => {
  const { S, s } = entorno();
  S.estado.usuario = "visitante@ficticio.example";
  const antes = fotos(s);
  lanza(() => S.ctx.prepararPeriodosCarga(), /propietario/);
  lanza(() => S.ctx.previsualizarPeriodosCarga(), /propietario/);
  igual(fotos(s), antes);
});

console.log("\n# Compatibilidad con la carga y el generador");
prueba("después de elegir docentes y copiar el borrador a Asignaciones (E1), la carga ve los cursantes como pendientes y verificar no señala materias sin docente", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  const h = s.admin[BORRADOR];
  const nombres = ["Prof. Mate Ficticio", "Prof. Lengua Ficticia", "Prof. Taller Ficticio", "Prof. Cocina Ficticia"];
  // Filas nuevas: elegir de la lista. Filas existentes: Matemática II ya era consistente; Inglés I se corrige a mano (nivel, docente y correo).
  borrador(s).forEach((f, i) => { if (i > 3) h.v[i][3] = f[0] === "EPSG" ? nombres[3] : nombres[i % 3]; });
  const fi = borrador(s).findIndex(f => f[2] === "Inglés I");
  h.v[fi][1] = "EPS2"; h.v[fi][3] = "Prof. Lengua Ficticia"; h.v[fi][4] = PROF("lengua");
  h.v[borrador(s).findIndex(f => f[2] === "Matemática IV")][3] = "Prof. Mate Ficticio";
  ctx.prepararPeriodosCarga(); // vuelve a revisar el borrador
  const listas = borrador(s).slice(1).filter(f => f[2] !== "Panadería");
  cierto(listas.every(f => f[3] && f[4] && !f[6].replace(/Sólo ACREDITÓ o EQUIVALENCIA[^.]*\./, "").trim()), listas.filter(f => f[6]).map(f => f[2] + ": " + f[6]).join(" | "));
  // E1: pegar sólo valores de las cinco columnas en Asignaciones; ventana futura (E3).
  s.admin.Asignaciones.v = [["Orientación", "Nivel", "Materia", "Docente", "Correo"]].concat(listas.map(f => f.slice(0, 5)));
  const hoy = D.isoBA(Date.now()), luego = D.isoBA(Date.now() + 5 * 86400000);
  s.admin.Ventanas.v = [["Período", "Desde", "Hasta"], [P2C, hoy, luego], [PAN, hoy, luego]];
  const r = ctx.verificarCargaDocente();
  igual(r.problemas.filter(p => /ningún docente asignado|no figura en la grilla|es de EPS/.test(p)), []);
  cierto(/0 cargadas, 1 del Equipo EPS, 0 a revisar; 0 con Calidad/.test(r.grillas.find(x => x.startsWith(PAN + " · EPSG"))), r.grillas.join(" | "));
  const v = ctx.consultarCarga_(PROF("mate"));
  const iv = v.periodos.find(p => p.periodo === P2C).materias.find(m => m.materia === "Matemática IV");
  cierto(iv && iv.estudiantes.length === 1 && iv.estudiantes[0].situacion === "pendiente", JSON.stringify(v.periodos.map(p => p.materias.map(m => m.materia))));
});

prueba("el generador no ve calificaciones de 1º cuatrimestre en las grillas preparadas (CURSA queda excluido del boletín)", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  const Gx = {}; vm.createContext(Gx); vm.runInContext(D.leer("src/Generador_EPS.gs"), Gx);
  const v = tabla(s.grillas[ID[PAN]].EPSG);
  const est = Gx.detectarEstructuraFuente_(v, PAN);
  igual(Gx.leerMateriasEstudiante_(v, v.findIndex(f => f[2] === J.dni), est, PAN, []).ordenMaterias, []);
  const k = Gx.leerMateriasEstudiante_(v, v.findIndex(f => f[2] === K.dni), est, PAN, []);
  igual(Object.values(k.materiasPorClave).map(m => m.valores[PAN]), ["EQUIVALENCIA"]);
});

console.log("\n# Invariantes con grillas aleatorias");
prueba("CURSA sólo con calificación ordinaria de 3B sobre vacía, NO CURSA o nota copiada; nada más cambia; repetir no cambia nada, ni después de quitar cursadas", () => {
  const corridas = Number(process.argv[2]) || 40;
  const ESC3 = ["En proceso", "Suficiente", "Avanzado"], OPC = ["", "", "CURSA", "NO CURSA", "ACREDITÓ", "EQUIVALENCIA"];
  for (let semilla = 1; semilla <= corridas; semilla++) {
    let x = semilla * 9301 + 49297;
    const azar = () => (x = (x * 9301 + 49297) % 233280) / 233280;
    const elegir = l => l[Math.floor(azar() * l.length)];
    const { ctx, s } = entorno((s, C) => {
      const nombres3 = C.estructuraGrilla_(ENC.EPSI_3B.map(f => f.slice()), P3B).materias.map(m => m.nombre);
      const nombres1 = C.estructuraGrilla_(ENC.EPSI_1C.map(f => f.slice()), P1C).materias.map(m => m.nombre);
      const est = [], est3 = [], est2 = [], estA = [];
      for (let n = 0; n < 12; n++) {
        const grupo = elegir(["EPS1", "EPS2", "EPS3", "EPS4"]);
        const base = { grupo, dni: String(40000000 + n), ap: "Ficticio" + n, no: "Azar", cola: INF };
        const v1 = {}, v3 = {}, v2 = {}, va = {};
        nombres1.forEach(m => { if (azar() < 0.5) v1[m] = azar() < 0.6 ? [String(1 + Math.floor(azar() * 10)), elegir(["", "Buena"]), elegir(["", "Activa"]), elegir(["", "obs"])] : elegir(OPC); });
        nombres3.forEach(m => { if (azar() < 0.5) v3[m] = azar() < 0.6 ? elegir(ESC3) : elegir(OPC); });
        nombres1.forEach(m => {
          [v2, va].forEach(v => { const r = azar(); v[m] = r < 0.5 ? v1[m] || "" : r < 0.7 ? "" : r < 0.8 ? "CURSA" : r < 0.9 ? "NO CURSA" : r < 0.95 ? String(1 + Math.floor(azar() * 10)) : "EQUIVALENCIA"; });
        });
        if (azar() < 0.9) est.push(Object.assign({}, base, { v: v1 }));
        if (azar() < 0.9) est3.push(Object.assign({}, base, { v: v3, grupo: azar() < 0.9 ? grupo : "EPS2" }));
        if (azar() < 0.95) est2.push(Object.assign({}, base, { v: v2 }));
        if (azar() < 0.95) estA.push(Object.assign({}, base, { v: va }));
      }
      s.grillas[ID[P1C]].EPSI = grilla(C, ENC.EPSI_1C, P1C, est);
      s.grillas[ID[P3B]].EPSI = grilla(C, ENC.EPSI_3B, P3B, est3);
      s.grillas[ID[P2C]].EPSI = grilla(C, ENC.EPSI_1C, P2C, est2);
      s.grillas[ID[PAN]].EPSI = grilla(C, ENC.EPSI_1C, PAN, estA);
    });
    const antes = { [P2C]: tabla(s.grillas[ID[P2C]].EPSI), [PAN]: tabla(s.grillas[ID[PAN]].EPSI) };
    const v1 = tabla(s.grillas[ID[P1C]].EPSI), v3 = tabla(s.grillas[ID[P3B]].EPSI);
    const e1 = ctx.estructuraGrilla_(v1, P1C), e3 = ctx.estructuraGrilla_(v3, P3B);
    ctx.prepararPeriodosCarga();
    const puestas = [];
    [P2C, PAN].forEach(P => {
      const ahora = tabla(s.grillas[ID[P]].EPSI), et = ctx.estructuraGrilla_(ahora, P);
      for (let r = 0; r < ahora.length; r++) for (let c = 0; c < ahora[r].length; c++) {
        const ant = antes[P][r][c];
        const m = et.materias.find(y => y.columnaValor === c);
        let base = r; while (base > 1 && !antes[P][base][2]) base--;
        const dni = antes[P][base][2], desp = r - base;
        const r1 = v1.findIndex(f => f[2] === dni), m1 = m && e1.materias.find(y => y.clave === m.clave);
        const previo = r1 >= 0 && m1 ? v1[r1 + desp][m1.columnaValor] : null;
        const r3 = v3.findIndex(f => f[2] === dni), m3 = m && e3.materias.find(y => y.clave === m.clave);
        const ev = r3 >= 0 && m3 && m3.nivel === m.nivel ? v3[r3][m3.columnaValor] : null;
        // Toda calificación vacía o NO CURSA de un estudiante identificado en 3B con calificación ordinaria quedó en CURSA.
        if (m && desp === 0 && dni && ["", "NO CURSA"].includes(ant) && ESC3.includes(ev) && v3.filter(f => f[2] === dni).length === 1) {
          igual(ahora[r][c], "CURSA", "semilla " + semilla + ": vacía con evidencia sin preparar (" + P + ", " + m.nombre + ")");
        }
        if (ahora[r][c] === ant) continue;
        cierto(m, "semilla " + semilla + ": cambió una celda fuera de las materias (" + r + "," + c + ")");
        if (desp > 0) { igual([ahora[r][c], ant], ["", previo], "semilla " + semilla + ": sólo se borra un complementario copiado"); continue; }
        igual(ahora[r][c], "CURSA", "semilla " + semilla);
        cierto(ESC3.includes(ev), "semilla " + semilla + ": CURSA sin calificación ordinaria en 3B");
        cierto(["", "NO CURSA"].includes(ant) || ant === previo, "semilla " + semilla + ": se cambió «" + ant + "», que no era vacía, NO CURSA ni resto de la copia (1C «" + previo + "»)");
        puestas.push([P, m.nombre, dni]);
      }
    });
    igual(ctx.prepararPeriodosCarga().cambios, 0, "semilla " + semilla + ": la segunda aplicación cambia cosas");
    // El Equipo EPS quita la mitad de las cursadas puestas: la tercera aplicación las respeta.
    const quitadas = puestas.filter((p, i) => i % 2 === 0);
    quitadas.forEach(([P, materia, dni], i) => poner(ctx, s, P, "EPSI", materia, dni, i % 4 ? "" : "NO CURSA"));
    igual(ctx.prepararPeriodosCarga().cambios, 0, "semilla " + semilla + ": volvió a poner cursadas quitadas");
    quitadas.forEach(([P, materia, dni]) => cierto(celda(ctx, s, P, "EPSI", materia, dni) !== "CURSA"));
  }
});

console.log("\n" + (fallas ? fallas + " de " + total + " pruebas fallaron" : "Las " + total + " pruebas pasaron"));
process.exitCode = fallas ? 1 : 0; // sin process.exit: con la salida a una tubería, cortaría el texto pendiente
