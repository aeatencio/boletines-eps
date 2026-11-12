/**
 * Carga docente EPS (#32) · pruebas locales de la preparación de 2º cuatrimestre y anual (Carga_Preparacion.gs y
 * Carga_Preparacion_Google.gs), con todos los .gs reales de src/carga y los dobles de tests/carga-dobles.js.
 *
 * El escenario reproduce la situación institucional con datos ficticios y los encabezados reales de los fixtures:
 * grilla de 1º cuatrimestre; 3º bimestre con la matrícula vigente (altas, una baja, un cambio de orientación, un
 * nombre corregido); y dos borradores copiados de 1º cuatrimestre y editados a mano de forma distinta (CURSA ya
 * puestos, una nota cargada, una fila agregada con un DNI mal escrito, una fila sin DNI, un cambio de grupo).
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
function igual(a, b, msg) { const x = JSON.stringify(a), y = JSON.stringify(b); if (x !== y) throw new Error((msg || "distinto") + ":\n      " + x + "\n    ≠ " + y); }
function cierto(x, msg) { if (!x) throw new Error(msg || "se esperaba verdadero"); }
function lanza(fn, re) { let e = null; try { fn(); } catch (x) { e = x; } if (!e || !re.test(e.message)) throw new Error("se esperaba un error " + re + (e ? ", fue: " + e.message : "")); return e; }

// ================================================================ escenario

const P1C = "1º cuatrimestre", P3B = "3º bimestre", P2C = "2º cuatrimestre", PAN = "Anual";
const ID = { [P1C]: "GRILLA_1C_ficticia_00000000001", [P3B]: "GRILLA_3B_ficticia_00000000001", [P2C]: "GRILLA_2C_ficticia_00000000001", [PAN]: "GRILLA_AN_ficticia_00000000001" };
const PROPIETARIA = "eps@ficticio.example";
const REFERENTE = "referente@ficticio.example";
const PROF = n => "prof." + n + "@ficticio.example";
const leerTsv = p => D.leer(p).replace(/\n$/, "").split("\n").map(l => l.split("\t"));
const ENC = {
  EPSI_1C: leerTsv("reference/fixtures/1er_cuatrimestre_estructura_real_EPSI.tsv").slice(0, 2),
  EPSI_3B: leerTsv("reference/fixtures/3er_bimestre_estructura_real_EPSI.tsv").slice(0, 2),
  EPSG: leerTsv("reference/fixtures/3er_bimestre_estructura_real_EPSG.tsv").slice(0, 2)
};
const ETIQUETAS = ["Valoración", "Calidad de las producciones:", "Participación:", "Observaciones:"];
const nrm = t => String(t).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().trim();

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

const INF = { "ORIENTACIÓN": "Informática" }, GAS = { "ORIENTACIÓN": "Gastronomía" };
const conCola = (e, extra) => Object.assign({}, e, { cola: Object.assign({}, e.cola, extra) });
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
const con = (e, v, extra) => Object.assign({}, e, { v: Object.assign({}, e.v, v) }, extra || {});
const G = { grupo: "EPS1", dni: "30000007", ap: "Ficticio", no: "Gael", cola: INF };

function escenarioPrep(C) {
  const g = {};
  g[ID[P1C]] = { EPSI: grilla(C, ENC.EPSI_1C, P1C, [A, B, Cc, Dd, E, F, L]), EPSG: grilla(C, ENC.EPSG, P1C, [J, K]) };
  g[ID[P3B]] = {
    EPSI: grilla(C, ENC.EPSI_3B, P3B, [
      con(A, { "Matemática I": ["Suficiente", "Buena"], "Lengua y Literatura I": "Avanzado", "Taller Profesional I A": "En proceso", "FGI I": "ACREDITÓ" }, { v: { "Matemática I": ["Suficiente", "Buena"], "Lengua y Literatura I": "Avanzado", "Taller Profesional I A": "En proceso", "FGI I": "ACREDITÓ" } }),
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
  const LgNoCursa = Object.assign({}, Lg, { v: { "Matemática I": "NO CURSA" } }); // fila nueva: NO CURSA equivale a vacía
  g[ID[P2C]] = {
    EPSI: grilla(C, ENC.EPSI_1C, P2C, [A, con(B, { "Matemática I": ["CURSA", "Regular", "Poca"] }), con(Cc, { "Matemática II": "9" }),
      con(Dd, { "Taller Profesional II A": "CURSA", "Matemática II": ["5", "Regular", "", "Nota EPS"] }), Object.assign({}, F, { grupo: "EPS3" }),
      Object.assign({}, G, { v: { "Matemática I": "CURSA" } }), { grupo: "EPS2", dni: "30000090", ap: "Ficticio", no: "Iván", cola: INF },
      { grupo: "EPS2", dni: "", ap: "Ficticia", no: "Sin Documento", v: { "Inglés I": "CURSA" } }]),
    EPSG: grilla(C, ENC.EPSG, P2C, [J, K, LgNoCursa])
  };
  g[ID[PAN]] = {
    EPSI: grilla(C, ENC.EPSI_1C, PAN, [A, B, Cc, Dd, E, F, Object.assign({}, G, { v: { "Matemática I": "CURSA" } })]),
    EPSG: grilla(C, ENC.EPSG, PAN, [J, K, Lg])
  };
  // Protecciones por columna de la carga directa anterior. La propietaria y la referente figuran en todas.
  const prot = (hoja, periodo, materia, editores, extra) => {
    const m = C.estructuraGrilla_(hoja.getDataRange().getDisplayValues(), periodo).materias.find(x => x.nombre === materia);
    return Object.assign({ filas: [2, 400], columnas: [m.columnaValor - 1, m.columnaValor], editores }, extra || {});
  };
  const eq = [PROPIETARIA, REFERENTE];
  const i3 = g[ID[P3B]].EPSI;
  i3.protecciones = [
    { filas: [0, 1], columnas: [0, 47], editores: [PROPIETARIA] }, // encabezado: no es de una materia
    prot(i3, P3B, "Matemática I", eq.concat(PROF("mate"))), prot(i3, P3B, "Lengua y Literatura I", eq.concat(PROF("lengua"))),
    prot(i3, P3B, "Taller Profesional I A", eq.concat(PROF("taller.a"), PROF("taller.b"))), prot(i3, P3B, "Matemática II", eq.concat(PROF("mate"))),
    prot(i3, P3B, "Inglés I", eq.concat(PROF("ingles"))), prot(i3, P3B, "FGI I", eq.slice()),
    prot(i3, P3B, "Matemática III", [PROPIETARIA, PROF("falsa")], { advertencia: true })];
  const i1 = g[ID[P1C]].EPSI;
  i1.protecciones = [prot(i1, P1C, "Matemática I", eq.concat(PROF("mate"))), prot(i1, P1C, "Inglés I", eq.concat(PROF("ingles.anterior"))), prot(i1, P1C, "Matemática IV", eq.concat(PROF("mate4")))];
  g[ID[P2C]].EPSI.protecciones = i1.protecciones.map(p => Object.assign({}, p)); // la copia conservó las de 1º cuatrimestre
  const g3 = g[ID[P3B]].EPSG;
  g3.protecciones = [prot(g3, P3B, "Matemática I", eq.concat(PROF("mate.g"))), prot(g3, P3B, "Taller Gastronómico I A", eq.concat(PROF("cocina")))];

  const control = Array.from({ length: 14 }, () => ["", ""]);
  control[10][1] = D.URL(ID[P1C]); control[11][1] = ID[P3B]; control[12][1] = D.URL(ID[P2C]); control[13][1] = D.URL(ID[PAN]);
  return {
    hoy: Date.now(), grillas: g, control: new Hoja(control), lock: D.nuevoLock(),
    admin: {
      Ventanas: new Hoja([["Período", "Desde", "Hasta"], ["3er bimestre", "2026-09-14", "2026-09-25"]]),
      Asignaciones: new Hoja([["Orientación", "Nivel", "Materia", "Docente", "Correo"], ["EPSI", "EPS1", "Lengua y Literatura I", "Prof. Lengua Ficticia", PROF("lengua")]]),
      Enlaces: new Hoja([["Correo", "Versión"]]),
      Registro: new Hoja([vm.runInContext("CARGA", C).REGISTRO.slice()]),
      Correos: new Hoja([vm.runInContext("CORREO", C).COLUMNAS.slice()]),
      "Asignaciones (borrador)": new Hoja([["Orientación", "Nivel", "Materia", "Docente", "Correo"],
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
  return { S, ctx: S.ctx, s: S.estado.s };
}
const copia = h => JSON.stringify(h.v);
const fotos = s => JSON.stringify(Object.keys(s.grillas).map(id => [id, s.grillas[id].EPSI.v, s.grillas[id].EPSG.v]));
const tabla = h => h.getDataRange().getDisplayValues();
function celda(ctx, s, periodo, hoja, materia, dni, desp) {
  const h = s.grillas[ID[periodo]][hoja];
  const v = tabla(h);
  const m = ctx.estructuraGrilla_(v, periodo).materias.find(x => x.nombre === materia);
  const fila = v.findIndex(f => f[2] === dni);
  if (!m || fila < 0) throw new Error("No se ubicó " + materia + " / " + dni);
  return v[fila + (desp || 0)][m.columnaValor];
}
const informe = s => tabla(s.admin["Preparación"]);
const avisos = s => { const f = informe(s); const i = f.findIndex(x => x[0] === "Prioridad"); return f.slice(i + 1).map(x => ({ prioridad: x[0], periodo: x[1], hoja: x[2], nivel: x[3], materia: x[4], fila: x[5], dni: x[6], estudiante: x[7], detalle: x[8] })); };
const hay = (lista, cond, msg) => { const r = lista.filter(cond); cierto(r.length, msg || "no se encontró el aviso esperado"); return r; };

// ================================================================ pruebas

console.log("\n# Decisión por celda (reglas de #32)");
prueba("decidirCalificacionPrep_: sólo una calificación ordinaria de 3º bimestre lleva a CURSA; ACREDITÓ, EQUIVALENCIA, CURSA sin calificar y vacío no", () => {
  const { ctx } = entorno();
  const d = (D0, C0, E0, falla) => ctx.decidirCalificacionPrep_(P2C, D0, C0, E0, falla);
  // resto de la copia (igual a 1º cuatrimestre)
  igual(!!d("7", "7", "Suficiente").cursa, true);
  ["ACREDITÓ", "EQUIVALENCIA", "CURSA", "", "NO CURSA", "Muy bien"].forEach(e => {
    const r = d("7", "7", e); igual([!!r.cursa, r.prioridad], [false, "Decidir"], "7 copiado y 3B «" + e + "»");
  });
  // vacío o NO CURSA copiados: CURSA si 3B tiene calificación; ACREDITÓ en 3B sólo se informa
  igual(!!d("", "", "Avanzado").cursa, true);
  igual(!!d("NO CURSA", "", "En proceso").cursa, true, "NO CURSA y vacío son equivalentes como resto");
  igual([!!d("", "", "ACREDITÓ").cursa, d("", "", "ACREDITÓ").prioridad], [false, "Revisar"]);
  igual(d("", "", "").cursa, undefined);
  // fila o columna que no venía de la copia: vacía y NO CURSA son equivalentes (#13) → CURSA con evidencia
  igual(!!d("", null, "Suficiente").cursa, true);
  igual(!!d("NO CURSA", null, "Suficiente").cursa, true);
  igual(!!d("no cursa", null, "Avanzado").cursa, true);
  // sin calificación ordinaria en 3B, ni vacía ni NO CURSA cambian; con ACREDITÓ en 3B, las dos se informan igual
  ["", "NO CURSA"].forEach(D0 => {
    [null, "", "NO CURSA"].forEach(C0 => {
      ["", "NO CURSA", "CURSA"].forEach(E0 => igual(d(D0, C0, E0), {}, "«" + D0 + "» / 1C «" + C0 + "» / 3B «" + E0 + "»"));
      igual(d(D0, C0, "EQUIVALENCIA").prioridad, "Revisar");
      igual(!!d(D0, C0, "Suficiente").cursa, true);
    });
    // quitada a mano respecto de 1C (cursaba): se conserva, sea vacía o NO CURSA
    igual([!!d(D0, "7", "Suficiente").cursa, d(D0, "7", "Suficiente").prioridad], [false, "Revisar"]);
    igual([!!d(D0, "CURSA", "Suficiente").cursa, d(D0, "CURSA", "Suficiente").prioridad], [false, "Revisar"]);
  });
  // cambios humanos: se conservan
  igual([!!d("NO CURSA", "7", "Suficiente").cursa, d("NO CURSA", "7", "Suficiente").prioridad], [false, "Revisar"]);
  igual(d("8", "7", "Suficiente").prioridad, "Decidir", "nota distinta de 1º cuatrimestre: no se toca");
  igual(d("8", null, "Suficiente").prioridad, "Decidir");
  igual(d("CURSA", "7", "Suficiente"), {});
  igual(d("CURSA", "7", "ACREDITÓ").prioridad, "Revisar");
  igual(d("ACREDITÓ", "ACREDITÓ", "ACREDITÓ"), {});
  igual(d("EQUIVALENCIA", "EQUIVALENCIA", "Suficiente").prioridad, "Revisar");
  igual(d("7,00", "7", "Suficiente").prioridad, "Decidir", "valor no reconocido");
  // sin evidencia comparable: nada cambia; un resto con nota queda contado para la línea del estudiante o la materia
  igual(d("7", "7", null, "sin estudiante"), { sinPreparar: true });
  igual(d("", "", null, "sin materia"), { sinPreparar: false });
  // Anual usa la misma regla, con 3º bimestre (no con 2º cuatrimestre)
  igual(!!ctx.decidirCalificacionPrep_(PAN, "7", "7", "Suficiente").cursa, true);
});

console.log("\n# Vista previa y aplicación sobre el escenario institucional ficticio");
prueba("vista previa: escribe sólo el informe; no toca grillas, borrador, Asignaciones, Ventanas, Registro ni Correos", () => {
  const { ctx, s } = entorno();
  const antes = fotos(s), adm = ["Asignaciones (borrador)", "Asignaciones", "Ventanas", "Enlaces", "Registro", "Correos"].map(n => copia(s.admin[n]));
  const r = ctx.previsualizarPeriodosCarga();
  igual([r.aplicar, r.puedeAplicar, r.aplicacion], [false, true, null]);
  cierto(r.cambios > 0);
  igual(fotos(s), antes, "grillas");
  igual(["Asignaciones (borrador)", "Asignaciones", "Ventanas", "Enlaces", "Registro", "Correos"].map(n => copia(s.admin[n])), adm);
  igual(s.admin["Preparación · cambios"], undefined);
  cierto(/VISTA PREVIA/.test(informe(s)[0][0]), informe(s)[0][0]);
  cierto(s.grillas[ID[P2C]].EPSI.escrituras.length === 0 && s.grillas[ID[P1C]].EPSI.intentos === 0);
});

prueba("aplicación: CURSA sólo con calificación de 3º bimestre, restos de la copia borrados, ACREDITÓ/EQUIVALENCIA y cambios humanos conservados", () => {
  const { ctx, s } = entorno();
  const r = ctx.prepararPeriodosCarga();
  igual([r.aplicacion.error, r.aplicacion.revisar, r.aplicacion.descartados], ["", 0, 0]);
  igual(r.aplicacion.aplicados, r.cambios);
  const c = (p, h, m, dni, desp) => celda(ctx, s, p, h, m, dni, desp);
  // A: 7, 8 y 6 copiados con calificación en 3B → CURSA; sus complementarios copiados, vacíos; ACREDITÓ se conserva.
  igual([c(P2C, "EPSI", "Matemática I", A.dni), c(P2C, "EPSI", "Matemática I", A.dni, 1), c(P2C, "EPSI", "Matemática I", A.dni, 2), c(P2C, "EPSI", "Matemática I", A.dni, 3)], ["CURSA", "", "", ""]);
  igual([c(P2C, "EPSI", "Lengua y Literatura I", A.dni), c(P2C, "EPSI", "Lengua y Literatura I", A.dni, 1), c(P2C, "EPSI", "Taller Profesional I A", A.dni), c(P2C, "EPSI", "FGI I", A.dni)], ["CURSA", "", "CURSA", "ACREDITÓ"]);
  // Taller Profesional I C no está en 3B: el 9 queda (sin evidencia), su Calidad copiada se borra.
  igual([c(P2C, "EPSI", "Taller Profesional I C", A.dni), c(P2C, "EPSI", "Taller Profesional I C", A.dni, 1)], ["9", ""]);
  // B: CURSA ya puesto se conserva; EQUIVALENCIA se conserva; 5 con NO CURSA en 3B queda; NO CURSA → CURSA por 3B.
  igual([c(P2C, "EPSI", "Matemática I", B.dni), c(P2C, "EPSI", "Matemática I", B.dni, 1), c(P2C, "EPSI", "Lengua y Literatura I", B.dni), c(P2C, "EPSI", "Taller Profesional I A", B.dni), c(P2C, "EPSI", "Taller Profesional I B", B.dni)],
    ["CURSA", "", "EQUIVALENCIA", "5", "CURSA"]);
  // C: la nota 9 escrita en el borrador no se toca; 6 con ACREDITÓ en 3B queda; las demás pasan a CURSA.
  igual([c(P2C, "EPSI", "Matemática II", Cc.dni), c(P2C, "EPSI", "Matemática I", Cc.dni), c(P2C, "EPSI", "Inglés I", Cc.dni), c(P2C, "EPSI", "Taller Profesional II A", Cc.dni)], ["9", "6", "CURSA", "CURSA"]);
  // D: Observaciones del Equipo EPS se conserva; CURSA agregado a mano se conserva; CURSA sin calificar en 3B sigue.
  igual([c(P2C, "EPSI", "Matemática II", Dd.dni), c(P2C, "EPSI", "Matemática II", Dd.dni, 1), c(P2C, "EPSI", "Matemática II", Dd.dni, 3), c(P2C, "EPSI", "Taller Profesional II A", Dd.dni), c(P2C, "EPSI", "Inglés I", Dd.dni)],
    ["CURSA", "", "Nota EPS", "CURSA", "CURSA"]);
  // G (alta): vacía con calificación en 3B → CURSA. L (cambió a EPSG): CURSA en EPSG por 3B, desde NO CURSA (2C) o vacía (anual).
  igual([c(P2C, "EPSI", "Lengua y Literatura I", G.dni), c(P2C, "EPSG", "Matemática I", L.dni), c(P2C, "EPSG", "Taller Gastronómico I A", L.dni)], ["CURSA", "CURSA", "CURSA"]);
  // Anual: la misma evidencia, independientemente de lo que tenga 2C. D sigue sin cursar Taller II A (como en 3B).
  igual([c(PAN, "EPSI", "Matemática I", B.dni), c(PAN, "EPSI", "Matemática II", Cc.dni), c(PAN, "EPSI", "Taller Profesional II A", Dd.dni), c(PAN, "EPSI", "Matemática III", E.dni)],
    ["CURSA", "CURSA", "NO CURSA", "7"]);
  igual([c(PAN, "EPSG", "Matemática I", J.dni), c(PAN, "EPSG", "Matemática I", J.dni, 1), c(PAN, "EPSG", "Taller Gastronómico II A", K.dni)], ["CURSA", "", "EQUIVALENCIA"]);
  igual([c(PAN, "EPSG", "Matemática I", L.dni), c(PAN, "EPSG", "Taller Gastronómico I A", L.dni)], ["CURSA", "CURSA"]);
  // Ninguna calificación anterior quedó como valoración nueva salvo las informadas; nada cambió fuera de las columnas de materias.
  const cola = s.grillas[ID[P2C]].EPSI.escrituras.filter(w => w.col <= 2);
  igual(cola, []);
  igual(tabla(s.grillas[ID[P2C]].EPSI).find(f => f[2] === A.dni).slice(41).filter(Boolean), ["Informática", "Obs general 1C", "SÍ", "SÍ"], "campos generales y operativos intactos");
  // Grillas de origen y evidencia: sólo lectura.
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
  igual([a[4], a[5], a[12], a[13]], [ID[P2C], "EPSI", "7", "CURSA"]);
  cierto(/copiada de 1º cuatrimestre; en 3º bimestre tiene «Suficiente»/.test(a[14]), a[14]);
  const obs = log.find(f => f[7] === A.dni && f[10] === "Matemática I" && f[11] === "Observaciones" && f[3] === P2C);
  igual([obs[12], obs[13]], ["Obs 1C Ana", ""]);
});

prueba("informe: resumen claro y excepciones por prioridad (bloqueo, decidir, revisar) con período, hoja, materia, fila y estudiante", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  const cab = informe(s).map(f => f[0]);
  cierto(/APLICADA/.test(cab[0]), cab[0]);
  cierto(cab.some(l => /^2º cuatrimestre: \d+ estudiantes \(EPSI 7, EPSG 3\)\. Se pasaron a CURSA \d+ calificaciones/.test(l)), cab.slice(0, 6).join(" | "));
  cierto(cab.some(l => /^Escritura: \d+ celdas aplicadas y verificadas, 0 a revisar, 0 no aplicadas/.test(l)));
  const x = avisos(s);
  const orden = x.map(a => ["Bloqueo", "Decidir", "Revisar"].indexOf(a.prioridad));
  igual(orden.slice().sort((p, q) => p - q), orden, "ordenado por prioridad");
  hay(x, a => a.prioridad === "Decidir" && a.materia === "Taller Profesional I A" && a.dni === B.dni && /Quedó la calificación «5» de 1º cuatrimestre y en 3º bimestre dice NO CURSA/.test(a.detalle));
  hay(x, a => a.prioridad === "Decidir" && a.materia === "Matemática II" && a.dni === Cc.dni && a.periodo === P2C && /«9» en el borrador, que no viene de 1º cuatrimestre \(allí: «8»\)/.test(a.detalle));
  hay(x, a => a.prioridad === "Decidir" && a.materia === "Matemática I" && a.dni === Cc.dni && /figura «ACREDITÓ»/.test(a.detalle));
  hay(x, a => a.prioridad === "Decidir" && a.materia === "Taller Profesional I C" && !a.dni && /No figura en la hoja EPSI de 3º bimestre.*1 calificación/.test(a.detalle));
  hay(x, a => a.prioridad === "Decidir" && a.estudiante === "Ficticia, Sin Documento" && /sin DNI/.test(a.detalle));
  hay(x, a => a.prioridad === "Decidir" && a.periodo === PAN && a.dni === E.dni && /No figura en 3º bimestre.*2 calificaciones/.test(a.detalle));
  hay(x, a => a.prioridad === "Revisar" && a.dni === "30000090" && /DNI 30000009 \(¿DNI mal escrito\?\)/.test(a.detalle));
  hay(x, a => a.prioridad === "Revisar" && a.dni === "30000009" && /donde figura «Ficticio, Iván» con DNI 30000090/.test(a.detalle));
  hay(x, a => a.dni === "30000008" && /Figura en 3º bimestre .* y no en el borrador/.test(a.detalle));
  hay(x, a => a.dni === Cc.dni && /nombres distintos: el borrador «Ficticia, Carla»; 3º bimestre «Ficticia, Karla»/.test(a.detalle));
  hay(x, a => a.dni === F.dni && a.periodo === P2C && /Grupo EPS3 en el borrador y EPS4 en 3º bimestre/.test(a.detalle));
  hay(x, a => a.dni === Dd.dni && a.materia === "Matemática II" && /Observaciones «Nota EPS» en el borrador/.test(a.detalle));
  hay(x, a => a.dni === Dd.dni && a.materia === "Taller Profesional II A" && a.periodo === P2C && /CURSA en el borrador, pero en 3º bimestre no la cursaba/.test(a.detalle));
  hay(x, a => a.periodo === P2C + " y " + PAN && a.dni === Dd.dni && /Cursa en 2º cuatrimestre \(«CURSA»\) y no en Anual \(«NO CURSA»\)/.test(a.detalle));
  hay(x, a => a.periodo === P2C + " y " + PAN && a.dni === E.dni && /Figura en Anual .* y no en 2º cuatrimestre/.test(a.detalle));
  hay(x, a => a.hoja === "EPSI" && a.periodo === P2C && /2 estudiantes con LISTO PARA EXPORTAR = SÍ igual que en 1º cuatrimestre/.test(a.detalle));
  hay(x, a => a.periodo === P2C && /1 estudiante con observaciones generales/.test(a.detalle));
  igual(x.filter(a => a.prioridad === "Bloqueo"), [], "sin bloqueos en el escenario");
  // L cambió de orientación entre 1C y 3B y los borradores lo reflejan: no es una excepción.
  igual(x.filter(a => a.dni === L.dni), []);
});

prueba("Asignaciones (borrador): una fila por materia con cursantes, orientación/nivel/materia de la grilla; correo sólo si todas las protecciones coinciden", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  const b = tabla(s.admin["Asignaciones (borrador)"]);
  igual(b[0], ["Orientación", "Nivel", "Materia", "Docente", "Correo", "Cursantes", "Origen del correo", "Revisión"]);
  const fila = (o, m) => b.find(f => f[0] === o && f[2] === m);
  // Filas existentes: sus cinco columnas, intactas.
  igual(b.slice(1, 4).map(f => f.slice(0, 5)), [["EPSI", "EPS2", "Matemática II", "Prof. Mate Ficticio", PROF("mate")], ["EPSI", "EPS3", "Inglés I", "Prof. Inglés Ficticia", PROF("ingles")],
    ["EPSG", "EPS1", "Panadería", "Docente Ficticio", "otro@ficticio.example"]]);
  igual([fila("EPSI", "Matemática II")[5], fila("EPSI", "Matemática II")[7]], ["2º cuatrimestre: 2 · Anual: 2", ""]);
  cierto(/El nivel de la materia en las grillas es EPS2/.test(fila("EPSI", "Inglés I")[7]), fila("EPSI", "Inglés I")[7]);
  cierto(/no figura en las grillas/.test(fila("EPSG", "Panadería")[7]));
  // Nuevas: unánime (3B, 1C y la copia 2C) → correo y nombre ya escrito para ese correo.
  igual(fila("EPSI", "Matemática I").slice(0, 7), ["EPSI", "EPS1", "Matemática I", "Prof. Mate Ficticio", PROF("mate"), "2º cuatrimestre: 4 · Anual: 4", "Protecciones de 3º bimestre, 2º cuatrimestre y 1º cuatrimestre"]);
  igual(fila("EPSI", "Lengua y Literatura I").slice(3, 5), ["Prof. Lengua Ficticia", PROF("lengua")], "nombre tomado de Asignaciones");
  igual(fila("EPSI", "Matemática IV").slice(3, 5), ["", PROF("mate4")]);
  cierto(/Falta el nombre del docente/.test(fila("EPSI", "Matemática IV")[7]));
  // Ambigua o sin protecciones: sin correo, con lo que dicen las protecciones.
  igual(fila("EPSI", "Taller Profesional I A")[4], "");
  cierto(/varias personas: 3º bimestre: prof\.taller\.a@ficticio\.example, prof\.taller\.b@ficticio\.example/.test(fila("EPSI", "Taller Profesional I A")[7]));
  igual([fila("EPSI", "FGI I")[4], fila("EPSI", "Matemática III")[4]], ["", ""], "sólo editores generales / protección de advertencia");
  igual([fila("EPSG", "Matemática I")[4], fila("EPSG", "Taller Gastronómico I A")[4]], [PROF("mate.g"), PROF("cocina")]);
  cierto(!b.some(f => f[4] === PROPIETARIA || f[4] === REFERENTE), "la propietaria y el Equipo EPS nunca se proponen");
  // Sin cursantes, sin fila (Taller Profesional I B de EPSG no existe; Taller II B no tiene cursantes).
  igual(fila("EPSI", "Taller Profesional II B"), undefined);
  igual(b.filter(f => f[0] === "EPSI" && f[2] === "Inglés I").length, 1, "no se duplica una materia que ya tenía fila");
  // La hoja Asignaciones (la que lee la carga) no cambió.
  igual(tabla(s.admin.Asignaciones).length, 2);
});

prueba("repetir la aplicación: no cambia nada, no duplica filas y conserva correcciones humanas posteriores", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  // El Equipo EPS decide algunos casos y completa el borrador después de la primera aplicación.
  const h = s.grillas[ID[P2C]].EPSI;
  const v = tabla(h);
  const m = ctx.estructuraGrilla_(v, P2C).materias.find(x => x.nombre === "Taller Profesional I C");
  const fa = v.findIndex(f => f[2] === A.dni);
  h.v[fa][m.columnaValor] = "NO CURSA";
  const bor = s.admin["Asignaciones (borrador)"];
  const fl = bor.v.findIndex(f => f[2] === "Taller Profesional I A");
  bor.v[fl][3] = "Prof. Taller Ficticio"; bor.v[fl][4] = PROF("taller.b");
  const grillas = fotos(s), filasBorrador = bor.v.length;
  const r = ctx.prepararPeriodosCarga();
  igual([r.cambios, r.aplicacion.aplicados], [0, 0]);
  igual(fotos(s), grillas);
  igual(bor.v.length, filasBorrador);
  igual(bor.v[fl].slice(3, 5), ["Prof. Taller Ficticio", PROF("taller.b")]);
  igual(bor.v[fl][7].includes("Falta"), false, bor.v[fl][7]);
  igual(celda(ctx, s, P2C, "EPSI", "Taller Profesional I C", A.dni), "NO CURSA");
  cierto(!avisos(s).some(a => a.materia === "Taller Profesional I C" && a.prioridad === "Decidir" && a.periodo === P2C), "el caso decidido ya no figura");
  hay(avisos(s), a => a.dni === Cc.dni && a.materia === "Matemática II" && a.periodo === P2C, "lo no decidido sigue informado");
});

console.log("\n# Salvaguardas");
prueba("con una ventana cargada para 2º cuatrimestre o anual no se modifica nada (también con fechas futuras); la vista previa informa el bloqueo", () => {
  const { ctx, s } = entorno(s => { s.admin.Ventanas.v.push(["2do cuatrimestre", "2026-12-01", "2026-12-04"]); });
  const antes = fotos(s), b = copia(s.admin["Asignaciones (borrador)"]);
  lanza(() => ctx.prepararPeriodosCarga(), /no se aplicó.*Ventanas, fila 3: hay una ventana de carga de 2º cuatrimestre/);
  igual([fotos(s), copia(s.admin["Asignaciones (borrador)"]), s.admin["Preparación · cambios"]], [antes, b, undefined]);
  cierto(/NO SE APLICÓ/.test(informe(s)[0][0]));
  const r = ctx.previsualizarPeriodosCarga();
  igual(r.puedeAplicar, false);
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
  const r = ctx.prepararPeriodosCarga();
  igual(r.aplicacion.error, "");
  igual(s.grillas[ID[P2C]].EPSI.escrituras.length, 0);
  cierto(s.grillas[ID[P2C]].EPSG.escrituras.length > 0);
  hay(avisos(s), a => a.prioridad === "Bloqueo" && a.periodo === P2C && a.hoja === "EPSI" && /ORIENTACIÓN.*No se prepara esta hoja/.test(a.detalle));
});

prueba("DNI repetido o bloque incompleto en el borrador: ese estudiante no se prepara y queda como bloqueo", () => {
  const { ctx, s } = entorno(s => {
    const h = s.grillas[ID[PAN]].EPSI;
    const fb = h.v.findIndex(f => f[2] === B.dni); h.v[fb][2] = A.dni;           // B con el DNI de A
    const ff = h.v.findIndex(f => f[2] === F.dni); h.v[ff + 2][4] = "Otra etiqueta"; // bloque de F alterado
  });
  ctx.prepararPeriodosCarga();
  const log = tabla(s.admin["Preparación · cambios"]).slice(1).filter(f => f[3] === PAN && f[5] === "EPSI");
  igual(log.filter(f => f[7] === A.dni || f[7] === F.dni), []);
  hay(avisos(s), a => a.prioridad === "Bloqueo" && a.periodo === PAN && a.dni === A.dni && /figura 2 veces/.test(a.detalle));
  hay(avisos(s), a => a.prioridad === "Bloqueo" && a.periodo === PAN && a.dni === F.dni && /etiqueta esperada/.test(a.detalle));
  cierto(log.length > 0, "los demás se preparan");
});

prueba("celdas en las que la carga docente ya escribió (Registro) no se tocan", () => {
  const { ctx, s } = entorno((s, C) => {
    const R = vm.runInContext("CARGA", C).REGISTRO;
    const fila = R.map(c => ({ "Fecha": "2026-11-17", "Estado": "aplicado", "Período": P2C, "Fuente": ID[P2C], "Orientación": "EPSI", "DNI": A.dni, "Materia": "Matemática I", "Campo": "Calificación" }[c] || ""));
    s.admin.Registro.v.push(fila);
  });
  ctx.prepararPeriodosCarga();
  igual([celda(ctx, s, P2C, "EPSI", "Matemática I", A.dni), celda(ctx, s, P2C, "EPSI", "Matemática I", A.dni, 1)], ["7", "Buena"]);
  igual(celda(ctx, s, PAN, "EPSI", "Matemática I", A.dni), "CURSA", "anual es otra grilla");
  hay(avisos(s), a => a.dni === A.dni && a.materia === "Matemática I" && a.periodo === P2C && /ya guardó cambios/.test(a.detalle));
});

prueba("una celda que el Equipo EPS cambia entre la lectura y la escritura no se pisa: queda «no aplicado»", () => {
  const { ctx, s } = entorno();
  const h = s.grillas[ID[P2C]].EPSI;
  const original = h.getDataRange.bind(h);
  let lecturas = 0;
  h.getDataRange = () => {
    if (++lecturas === 2) { // la relectura dentro del lock: justo antes, alguien edita la celda
      const v = original().getDisplayValues();
      const m = ctx.estructuraGrilla_(v, P2C).materias.find(x => x.nombre === "Matemática I");
      h.v[v.findIndex(f => f[2] === A.dni)][m.columnaValor] = "NO CURSA";
    }
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
  const h = s.grillas[ID[PAN]].EPSG;
  h.getRangeList = () => ({ setValue: () => { throw new Error("Servicio no disponible (simulado)"); }, clearContent: () => { throw new Error("Servicio no disponible (simulado)"); } });
  const r = ctx.prepararPeriodosCarga();
  cierto(r.aplicacion.revisar > 0 && r.aplicacion.aplicados > 0, JSON.stringify(r.aplicacion));
  const rev = tabla(s.admin["Preparación · cambios"]).filter(f => f[1] === "revisar");
  cierto(rev.length === r.aplicacion.revisar && rev.every(f => f[3] === PAN && f[5] === "EPSG" && /Servicio no disponible/.test(f[2])));
  cierto(informe(s).some(f => /\d+ a revisar/.test(f[0])));
});

prueba("borrador con columnas propias faltantes se completa a la derecha; sin columna estándar, el borrador no se toca pero las grillas sí", () => {
  const { ctx, s } = entorno(s => { s.admin["Asignaciones (borrador)"].v[0][1] = "Curso"; });
  const b = copia(s.admin["Asignaciones (borrador)"]);
  const r = ctx.prepararPeriodosCarga();
  igual(copia(s.admin["Asignaciones (borrador)"]), b);
  cierto(/no se modificó: .*no tiene la columna "Nivel"/.test(r.aplicacion.borrador), r.aplicacion.borrador);
  cierto(r.aplicacion.aplicados > 0);
  const sin = entorno(s => { delete s.admin["Asignaciones (borrador)"]; });
  sin.ctx.prepararPeriodosCarga();
  igual(tabla(sin.s.admin["Asignaciones (borrador)"])[0], ["Orientación", "Nivel", "Materia", "Docente", "Correo", "Cursantes", "Origen del correo", "Revisión"]);
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
prueba("después de preparar, la carga ve a los cursantes como CURSA pendientes y verificarCargaDocente no encuentra restos de la copia salvo lo informado", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  // Paso E1 del manual: las filas completas del borrador pasan a Asignaciones; ventana futura (paso E3).
  const b = tabla(s.admin["Asignaciones (borrador)"]).slice(1).filter(f => f[4] && f[3] && !/nivel|no figura/i.test(f[7]));
  s.admin.Asignaciones.v = [["Orientación", "Nivel", "Materia", "Docente", "Correo"]].concat(b.map(f => f.slice(0, 5)));
  const hoy = D.isoBA(Date.now()), luego = D.isoBA(Date.now() + 5 * 86400000);
  s.admin.Ventanas.v = [["Período", "Desde", "Hasta"], [P2C, hoy, luego], [PAN, hoy, luego]];
  const r = ctx.verificarCargaDocente();
  const linea = r.grillas.find(g => g.startsWith(PAN + " · EPSG"));
  cierto(/0 cargadas, 1 del Equipo EPS, 0 a revisar; 0 con Calidad/.test(linea), linea);
  const l2 = r.grillas.find(g => g.startsWith(P2C + " · EPSI"));
  // Quedan sólo las cuatro notas informadas como «Decidir» (C: 9 y 6; A: 9 de Taller I C; B: 5) y la observación del Equipo EPS.
  cierto(/, 4 cargadas, .*; 1 con Calidad/.test(l2), l2);
  const v = ctx.consultarCarga_(PROF("mate"));
  const mat = v.periodos.find(p => p.periodo === PAN).materias.find(m => m.materia === "Matemática I");
  igual(mat.estudiantes.map(e => [e.dni, e.situacion]).sort(), [[A.dni, "pendiente"], [B.dni, "pendiente"], [Cc.dni, "valorada"], [G.dni, "pendiente"]].sort());
});

prueba("el generador no ve calificaciones de 1º cuatrimestre en las grillas preparadas (CURSA queda excluido del boletín)", () => {
  const { ctx, s } = entorno();
  ctx.prepararPeriodosCarga();
  const Gx = {}; vm.createContext(Gx); vm.runInContext(D.leer("src/Generador_EPS.gs"), Gx);
  const v = tabla(s.grillas[ID[P2C]].EPSG);
  const est = Gx.detectarEstructuraFuente_(v, P2C);
  const leidas = Gx.leerMateriasEstudiante_(v, v.findIndex(f => f[2] === J.dni), est, P2C, []);
  igual(leidas.ordenMaterias, []);
  const k = Gx.leerMateriasEstudiante_(v, v.findIndex(f => f[2] === K.dni), est, P2C, []);
  igual(Object.values(k.materiasPorClave).map(m => m.valores[P2C]), ["EQUIVALENCIA"]);
});

console.log("\n# Invariantes con grillas aleatorias");
prueba("en grillas aleatorias: sólo cambian restos de la copia o vacías nuevas; CURSA sólo con calificación de 3B; la segunda aplicación no cambia nada", () => {
  const corridas = Number(process.argv[2]) || 40;
  const ESC3 = ["En proceso", "Suficiente", "Avanzado"], OPC = ["", "", "CURSA", "NO CURSA", "ACREDITÓ", "EQUIVALENCIA"];
  for (let semilla = 1; semilla <= corridas; semilla++) {
    let x = semilla * 9301 + 49297;
    const azar = () => (x = (x * 9301 + 49297) % 233280) / 233280;
    const elegir = l => l[Math.floor(azar() * l.length)];
    const { ctx, s } = entorno((s, C) => {
      const nombresI = C.estructuraGrilla_(ENC.EPSI_3B.map(f => f.slice()), P3B).materias.map(m => m.nombre);
      const nombresI1 = C.estructuraGrilla_(ENC.EPSI_1C.map(f => f.slice()), P1C).materias.map(m => m.nombre);
      const est = [], est3 = [], est2 = [], estA = [];
      for (let n = 0; n < 12; n++) {
        const grupo = elegir(["EPS1", "EPS2", "EPS3", "EPS4"]);
        const base = { grupo, dni: String(40000000 + n), ap: "Ficticio" + n, no: "Azar", cola: INF };
        const v1 = {}, v3 = {}, v2 = {}, va = {};
        nombresI1.forEach(m => { if (azar() < 0.5) v1[m] = azar() < 0.6 ? [String(1 + Math.floor(azar() * 10)), elegir(["", "Buena"]), elegir(["", "Activa"]), elegir(["", "obs"])] : elegir(OPC); });
        nombresI.forEach(m => { if (azar() < 0.5) v3[m] = azar() < 0.6 ? elegir(ESC3) : elegir(OPC); });
        nombresI1.forEach(m => {
          [v2, va].forEach(v => { const r = azar(); v[m] = r < 0.7 ? v1[m] || "" : r < 0.8 ? "CURSA" : r < 0.9 ? "NO CURSA" : r < 0.95 ? String(1 + Math.floor(azar() * 10)) : ""; });
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
      s.__fuentes = { est, est3 };
    });
    const antes = { [P2C]: tabla(s.grillas[ID[P2C]].EPSI), [PAN]: tabla(s.grillas[ID[PAN]].EPSI) };
    const v1 = tabla(s.grillas[ID[P1C]].EPSI), v3 = tabla(s.grillas[ID[P3B]].EPSI);
    const e1 = ctx.estructuraGrilla_(v1, P1C), e3 = ctx.estructuraGrilla_(v3, P3B);
    ctx.prepararPeriodosCarga();
    [P2C, PAN].forEach(P => {
      const ahora = tabla(s.grillas[ID[P]].EPSI), et = ctx.estructuraGrilla_(ahora, P);
      for (let r = 0; r < ahora.length; r++) for (let c = 0; c < ahora[r].length; c++) {
        if (ahora[r][c] === antes[P][r][c]) continue;
        const m = et.materias.find(x => x.columnaValor === c);
        cierto(m, "semilla " + semilla + ": cambió una celda fuera de las materias (" + r + "," + c + ")");
        let base = r; while (!antes[P][base][2]) base--;
        const desp = r - base, dni = antes[P][base][2];
        const r1 = v1.findIndex(f => f[2] === dni), m1 = e1.materias.find(x => x.clave === m.clave);
        const previo = r1 >= 0 && m1 ? v1[r1 + desp][m1.columnaValor] : null;
        if (desp > 0) { igual([ahora[r][c], antes[P][r][c]], ["", previo], "semilla " + semilla + ": sólo se borra un complementario copiado"); continue; }
        igual(ahora[r][c], "CURSA", "semilla " + semilla);
        const r3 = v3.findIndex(f => f[2] === dni), m3 = e3.materias.find(x => x.clave === m.clave);
        cierto(r3 >= 0 && m3 && ESC3.includes(v3[r3][m3.columnaValor]), "semilla " + semilla + ": CURSA sin calificación ordinaria en 3B");
        const ant = antes[P][r][c];
        cierto(previo === null ? ["", "NO CURSA"].includes(ant) : (["", "NO CURSA"].includes(ant) && ["", "NO CURSA"].includes(previo)) || ant === previo, "semilla " + semilla + ": se cambió algo que no era resto de la copia: «" + ant + "» (1C «" + previo + "»)");
      }
    });
    const segunda = ctx.prepararPeriodosCarga();
    igual(segunda.cambios, 0, "semilla " + semilla + ": la segunda aplicación cambia cosas");
  }
});

console.log("\n" + (fallas ? fallas + " de " + total + " pruebas fallaron" : "Las " + total + " pruebas pasaron"));
process.exit(fallas ? 1 : 0);
