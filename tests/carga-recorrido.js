/**
 * Carga docente EPS (#29) · ensayo local del recorrido de docs/puesta-en-marcha-carga-docente.md, etapas B a D.
 *
 * Recorre, con los dobles de tests/carga-dobles.js y un calendario fijo, los pasos del manual que tienen un
 * resultado esperado observable sin Google (registro de las funciones del editor, hoja Correos, Registro, respuestas
 * de la app) y comprueba que coincide con lo que el manual cita. Si una prueba falla después de cambiar el código,
 * hay que actualizar el manual o el código: no deben divergir.
 *
 * No reemplaza la prueba integrada en Google: los puntos [G] del manual sólo se observan allí.
 *
 * Uso: node tests/carga-recorrido.js
 */
"use strict";
const D = require("./carga-dobles.js");

const A = "docente.prueba@ficticio.example";
const REGISTRO = ["Fecha", "Estado", "Detalle", "Intentos", "Operación", "Correo", "Docente", "Período", "Fuente", "Orientación", "Nivel", "Materia",
  "DNI", "Estudiante", "Campo", "Anterior", "Nuevo"];
const CORREOS = ["Fecha", "Día", "Estado", "Detalle", "Correo", "Docente", "Tipo", "Acceso", "Eventos", "Pendientes", "Asunto"];
const SOLO_ACTIVAR = ["Los correos no están activados (activarCorreosCarga)."];

let fallas = 0, total = 0;
function paso(nombre, fn) {
  total++;
  try { fn(); console.log("OK    " + nombre); } catch (e) { fallas++; console.log("FALLA " + nombre + "\n      " + String(e && e.stack || e).split("\n").slice(0, 4).join("\n      ")); }
}
function igual(a, b, msg) { const x = JSON.stringify(a), y = JSON.stringify(b); if (x !== y) throw new Error((msg || "distinto") + ": " + x + " ≠ " + y); }
function cierto(x, msg) { if (!x) throw new Error(msg || "se esperaba verdadero"); }

// Calendario: día 0 lunes 16/11/2026; días 1 a 3, martes a jueves (hábiles). Cierre de la ventana de la prueba: día 3, 8:00.
const DIA = { 0: "2026-11-16", 1: "2026-11-17", 2: "2026-11-18", 3: "2026-11-19" };
const S = D.servidor();
const { ctx, estado } = S;
S.aniosFicticios.forEach(a => D.sinAnioDeCalendario(ctx, a)); // sólo el calendario real del código
const ad = () => S.s.admin;
const en = local => { S.reloj.desfase = Date.parse(local + ":00-03:00") - Date.now(); };
const registroEjecucion = () => S.consola.splice(0);
const filas = (hoja, columnas) => hoja.getDataRange().getDisplayValues().slice(1).map(f => Object.fromEntries(columnas.map((c, i) => [c, f[i]])));
const registro = () => filas(ad().Registro, REGISTRO);
const correos = () => filas(ad().Correos, CORREOS);
const mails = () => estado.correo.enviados;
const credencial = m => (m.htmlBody.match(/\?k=([A-Za-z0-9_-]{43})/) || [])[1] || "";
let nOp = 0;
const t = {}; // lo que los pasos van dejando: enlaces, períodos y estudiante de la prueba

// A8: administración vacía; sin activadores. Los feriados son los del calendario del código (2026).
ad().Ventanas.v = [["Período", "Desde", "Hasta"]];
ad().Asignaciones.v = [["Orientación", "Nivel", "Materia", "Docente", "Correo"]];
en(DIA[0] + "T10:00");

console.log("\n# B. Verificación sin habilitar el servicio");
paso("B1: verificarCargaDocente sin ventanas: el calendario de feriados y sólo «Los correos no están activados», con las secciones del registro", () => {
  const r = ctx.verificarCargaDocente();
  igual([r.ventanas, r.grillas, r.problemas], [[], [], SOLO_ACTIVAR]);
  igual(registroEjecucion()[0].split("\n"), ["Calendario de feriados: 2026", "", "Ventanas vigentes o próximas:", "(ninguna)", "", "Grillas de esas ventanas:", "(ninguna)", "",
    "Problemas:", SOLO_ACTIVAR[0]]);
});
paso("B2: previsualizarCorreosCarga: \"simulados\":0 y el motivo citado", () => {
  ctx.previsualizarCorreosCarga();
  const l = registroEjecucion()[0];
  cierto(/"simulados":0/.test(l) && /"motivo":"No hay ventanas en curso ni cierres para informar\."/.test(l), l);
});
paso("B3: la URL sola o con una credencial cualquiera: «Acceso no verificado» con el motivo citado", () => {
  [ctx.cargaConsultar(""), ctx.cargaConsultar("A".repeat(43))].forEach(r => {
    igual([r.estado, r.codigo], ["rechazado", "acceso"]);
    cierto(r.motivo.startsWith("El enlace no es válido, fue reemplazado o es de una carga anterior"), r.motivo);
  });
});
paso("B4 / C1: ventana de la prueba a futuro: ventanas «(futura)» con cierre el día 3, una línea de grilla por período y hoja, nadie tiene enlace", () => {
  ad().Asignaciones.v.push(["EPSI", "EPS2", "Matemática II", "Docente de prueba", A]);
  ad().Ventanas.v.push(["2º cuatrimestre", DIA[1], DIA[2]], ["Anual", DIA[1], DIA[2]]);
  const r = ctx.verificarCargaDocente();
  registroEjecucion();
  igual(r.ventanas, ["2º cuatrimestre: " + DIA[1] + " a " + DIA[2] + ", cierra " + DIA[3] + " 08:00 (futura)", "Anual: " + DIA[1] + " a " + DIA[2] + ", cierra " + DIA[3] + " 08:00 (futura)"]);
  igual(r.grillas.map(g => g.split(":")[0]), ["2º cuatrimestre · EPSI", "2º cuatrimestre · EPSG", "Anual · EPSI", "Anual · EPSG"]);
  cierto(r.problemas.some(p => /cursantes y ningún docente asignado\.$/.test(p)) && r.problemas.includes(SOLO_ACTIVAR[0]));
  ctx.previsualizarCorreosCarga();
  cierto(/"simulados":0/.test(registroEjecucion()[0]));
  igual(Object.keys(D.enlacesReales(ctx, [A])).length, 0);
});

console.log("\n# C2. Día 1");
en(DIA[1] + "T09:00");
paso("C2.1: previsualizar: sólo la cuenta de prueba, apertura de los dos períodos", () => {
  ctx.previsualizarCorreosCarga();
  const sim = registroEjecucion().filter(x => /"correo":"simulado"/.test(x));
  igual(sim.length, 1);
  cierto(sim[0].includes('"a":"' + A + '"') && sim[0].includes('"tipos":["apertura"]') && sim[0].includes('"periodos":["2º cuatrimestre (apertura)","Anual (apertura)"]'), sim[0]);
});
paso("C2.2 / C2.3: enviar: una fila «enviado», «apertura», acceso «<día 1>|v1»; remitente y asunto citados", () => {
  igual(ctx.enviarCorreosCarga().enviados, 1);
  registroEjecucion();
  const f = correos();
  igual(f.map(x => [x["Estado"], x["Tipo"], x["Acceso"], x["Correo"]]), [["enviado", "apertura", DIA[1] + "|v1", A]]);
  const m = mails()[0];
  igual(m.name, "Equipo EPS");
  cierto(/^Carga de valoraciones EPS: abrió 2º cuatrimestre y Anual · \d+ calificaciones pendientes, cierra el 19\/11$/.test(m.subject), m.subject);
  t.k1 = credencial(m);
  cierto(t.k1, "el correo trae el enlace");
});
paso("C2.4: segunda ejecución del día: no envía (yaRecibieron) y Correos no cambia", () => {
  const r = ctx.enviarCorreosCarga();
  registroEjecucion();
  igual([r.enviados, r.yaRecibieron, correos().length, mails().length], [0, 1, 1, 1]);
});
paso("C2.5: el enlace abre los dos períodos con la materia de prueba y sus cursantes", () => {
  const q = ctx.cargaConsultar(t.k1);
  igual([q.estado, q.periodos.map(p => p.periodo), q.periodos.map(p => p.materias.map(m => m.materia))], ["ok", ["2º cuatrimestre", "Anual"], [["Matemática II"], ["Matemática II"]]]);
  t.p2c = q.periodos[0]; t.pan = q.periodos[1];
  t.e1 = t.p2c.materias[0].estudiantes.find(e => e.situacion === "pendiente" && !e.problema && !e.campos.calidad && !e.campos.participacion && !e.campos.observaciones &&
    t.pan.materias[0].estudiantes.some(x => x.dni === e.dni && x.situacion === "pendiente" && !x.problema));
  cierto(t.e1, "hay un estudiante como E1");
});
paso("C2.7: enlace con un carácter cambiado o recortado: «Acceso no verificado»", () => {
  const otro = t.k1.slice(0, -1) + (t.k1.slice(-1) === "A" ? "B" : "A");
  [otro, t.k1.slice(0, 20)].forEach(k => igual(ctx.cargaConsultar(k).codigo, "acceso"));
});

const pedido = (p, campo, esperado, nuevo, op) => ({ operacion: op || "op-recorrido-" + String(++nOp).padStart(4, "0"), periodo: p.periodo, fuente: p.fuente,
  orientacion: "EPSI", materia: p.materias[0].clave, dni: t.e1.dni, campo, esperado, nuevo });
const guardar = (p, campo, esperado, nuevo, op) => ctx.cargaGuardar(t.k, pedido(p, campo, esperado, nuevo, op));

paso("C2.8: una materia agregada aparece al actualizar; retirada con la página abierta, el guardado se rechaza sin escribir", () => {
  ad().Asignaciones.v.push(["EPSI", "EPS2", "Inglés I", "Docente de prueba", A]);
  const q = ctx.cargaConsultar(t.k1);
  const ingles = q.periodos[0].materias.find(m => m.materia === "Inglés I");
  cierto(ingles, "aparece la materia agregada");
  const e = ingles.estudiantes.find(x => x.situacion === "pendiente" && !x.problema);
  ad().Asignaciones.v.pop();
  const antes = registro().length;
  const r = ctx.cargaGuardar(t.k1, { operacion: "op-recorrido-ingles", periodo: q.periodos[0].periodo, fuente: q.periodos[0].fuente, orientacion: "EPSI",
    materia: ingles.clave, dni: e.dni, campo: "calificacion", esperado: "CURSA", nuevo: "7" });
  igual([r.estado, r.codigo, registro().length], ["rechazado", "asignacion", antes]);
  cierto(r.motivo.startsWith("Esta materia no está asignada a tu cuenta"), r.motivo);
});
paso("C2.9 a C2.12: guardado, corrección, complementarios y anual independiente; cada uno, una fila «aplicado»", () => {
  t.k = t.k1;
  igual(guardar(t.p2c, "calificacion", "CURSA", "7").estado, "guardado");
  const f = registro()[0];
  igual([f["Estado"], f["Anterior"], f["Nuevo"], f["Correo"], f["Período"], f["Fuente"], f["DNI"], f["Materia"]],
    ["aplicado", "CURSA", "7", A, "2º cuatrimestre", t.p2c.fuente, t.e1.dni, "Matemática II"]);
  igual(guardar(t.p2c, "calificacion", "7", "8").estado, "guardado");
  igual([guardar(t.p2c, "calidad", "", "Buena").estado, guardar(t.p2c, "participacion", "", "Activa").estado, guardar(t.p2c, "observaciones", "", "Prueba integrada").estado],
    ["guardado", "guardado", "guardado"]);
  igual(guardar(t.pan, "calificacion", "CURSA", "9").estado, "guardado");
  const q = ctx.cargaConsultar(t.k);
  const de = i => q.periodos[i].materias[0].estudiantes.find(x => x.dni === t.e1.dni).campos;
  igual([de(0).calificacion, de(0).calidad, de(0).participacion, de(0).observaciones, de(1).calificacion], ["8", "Buena", "Activa", "Prueba integrada", "9"]);
  igual(registro().map(x => x["Estado"]), Array(6).fill("aplicado"));
  igual(registro()[5]["Período"], "Anual");
});
paso("C2.13: la grilla cambió (edición del Equipo EPS): «conflicto» sin fila; guardar el valor propio: «aplicado» con el Anterior de la grilla", () => {
  const h = S.s.grillas[t.p2c.fuente].EPSI;
  const v = h.getDataRange().getDisplayValues();
  const e = ctx.estructuraGrilla_(v, "2º cuatrimestre");
  const col = ctx.materiaPorClave_(e, t.p2c.materias[0].clave).columnaValor;
  const fila = ctx.filaEstudiante_(v, e, t.e1.dni);
  h.v[fila + 3][col] = "Cambio EPS";
  const antes = registro().length;
  const r = guardar(t.p2c, "observaciones", "Prueba integrada", "Prueba 2");
  igual([r.estado, r.actual, registro().length], ["conflicto", "Cambio EPS", antes]);
  igual(guardar(t.p2c, "observaciones", "Cambio EPS", "Prueba 2").estado, "guardado");
  igual([registro()[antes]["Estado"], registro()[antes]["Anterior"]], ["aplicado", "Cambio EPS"]);
});
paso("C2.14: el mismo pedido reenviado (respuesta perdida) no duplica: una sola fila para esa operación", () => {
  const antes = registro().length;
  const op = "op-recorrido-reintento";
  igual(guardar(t.p2c, "participacion", "Activa", "Poca", op).estado, "guardado");
  igual(guardar(t.p2c, "participacion", "Activa", "Poca", op).estado, "guardado");
  igual([registro().length - antes, registro().filter(x => x["Operación"] === op).length], [1, 1]);
});
paso("C2.18: restauración desde la app: cada celda vuelve a su valor original", () => {
  ["guardado", "guardado", "guardado", "guardado", "guardado"].forEach((esperado, i) => igual([
    () => guardar(t.p2c, "calificacion", "8", "CURSA"), () => guardar(t.p2c, "calidad", "Buena", ""), () => guardar(t.p2c, "participacion", "Poca", ""),
    () => guardar(t.p2c, "observaciones", "Prueba 2", ""), () => guardar(t.pan, "calificacion", "9", "CURSA")][i]().estado, esperado));
  const q = ctx.cargaConsultar(t.k);
  const de = i => q.periodos[i].materias[0].estudiantes.find(x => x.dni === t.e1.dni).campos;
  igual([de(0), de(1).calificacion], [{ calificacion: "CURSA", calidad: "", participacion: "", observaciones: "" }, "CURSA"]);
});
paso("C2.19: revocación (Enlaces, versión 2): el guardado con la página abierta se rechaza sin fila y el enlace deja de abrir", () => {
  ad().Enlaces.v.push([A, 2]);
  const antes = registro().length;
  const r = guardar(t.p2c, "calidad", "", "Buena");
  igual([r.estado, r.codigo, registro().length, ctx.cargaConsultar(t.k1).codigo], ["rechazado", "acceso", antes, "acceso"]);
});
paso("C2.20: activarCorreosCarga: el texto citado y un único activador", () => {
  registroEjecucion();
  ctx.activarCorreosCarga();
  const l = registroEjecucion()[0];
  cierto(/^Correos activados \(activador [^)]+\)\. Corre todos los días entre las 8:00 y las 9:00; sólo envía en días hábiles\.$/.test(l), l);
  igual(estado.activadores.length, 1);
});
paso("C2.21: verificar sin filas del Registro ni avisos de correos; resolverPendientesCarga: revisadas 0", () => {
  const r = ctx.verificarCargaDocente();
  cierto(!r.problemas.some(p => /^Registro|correos|activador/i.test(p)), r.problemas.filter(p => /^Registro|correos|activador/i.test(p)).join(" | "));
  igual(ctx.resolverPendientesCarga().revisadas, 0);
  registroEjecucion();
});

console.log("\n# C3. Día 2");
paso("C3.1 a C3.4: el activador envía «seguimiento, enlace» con acceso «<día 1>|v2»; el enlace nuevo abre y el anterior no", () => {
  en(DIA[2] + "T08:30");
  igual(ctx.enviarCorreosCarga({ triggerUid: estado.activadores[0].uid }).enviados, 1);
  registroEjecucion();
  const f = correos()[1];
  igual([f["Día"], f["Estado"], f["Tipo"], f["Acceso"]], [DIA[2], "enviado", "seguimiento, enlace", DIA[1] + "|v2"]);
  const m = mails()[1];
  cierto(m.body.includes("Tu enlace de acceso cambió: usá el de este correo"));
  t.k2 = credencial(m);
  igual([ctx.cargaConsultar(t.k2).estado, ctx.cargaConsultar(t.k1).codigo], ["ok", "acceso"]);
});
paso("C3.5: «Hasta» al día 1 con la página abierta: «La carga de 2º cuatrimestre cerró.», sin fila; la página queda sin cargas abiertas", () => {
  en(DIA[2] + "T09:30");
  t.k = t.k2;
  const q = ctx.cargaConsultar(t.k2);
  t.p2c = q.periodos[0];
  ad().Ventanas.v[1][2] = DIA[1]; ad().Ventanas.v[2][2] = DIA[1];
  const antes = registro().length;
  const r = guardar(t.p2c, "calificacion", "CURSA", "7");
  igual([r.estado, r.codigo, r.motivo, registro().length], ["rechazado", "ventana", "La carga de 2º cuatrimestre cerró.", antes]);
  igual(ctx.cargaConsultar(t.k2).periodos.length, 0);
  igual(ctx.verificarCargaDocente().ventanas, []);
  registroEjecucion();
});
paso("C3.6: todas las filas de la prueba están «aplicado»", () => {
  igual(registro().filter(x => x["Estado"] !== "aplicado").length, 0);
});

console.log("\n# C4. Día 3");
paso("C4: el activador envía el cierre, sin enlace", () => {
  en(DIA[3] + "T08:30");
  igual(ctx.enviarCorreosCarga({ triggerUid: estado.activadores[0].uid }).enviados, 1);
  registroEjecucion();
  igual(correos()[2]["Tipo"], "cierre");
  const m = mails()[2];
  cierto(!credencial(m) && m.body.includes("Cerró la carga de 2º cuatrimestre y Anual"), m.subject);
});

console.log("\n# D. Limpieza");
paso("D1: desactivarCorreosCarga: el texto citado y sin activadores", () => {
  ctx.desactivarCorreosCarga();
  igual([registroEjecucion()[0], estado.activadores.length], ["Activadores de correos borrados: 1.", 0]);
});
paso("D2, D3, D6 y D7: sin filas de prueba en la configuración y con «Nota EPS» a la derecha de Registro y Correos, verificar como en B1; los enlaces de la prueba no abren", () => {
  ["Ventanas", "Asignaciones", "Enlaces"].forEach(h => { ad()[h].v = [ad()[h].v[0]]; });
  ["Registro", "Correos"].forEach(h => {
    const v = ad()[h].v;
    v[0].push("Nota EPS");
    v.slice(1).forEach(f => f.push("Prueba integrada #29"));
  });
  igual(ctx.verificarCargaDocente().problemas, SOLO_ACTIVAR);
  ctx.previsualizarCorreosCarga();
  cierto(registroEjecucion().some(l => /"simulados":0/.test(l)));
  igual([ctx.cargaConsultar(t.k1).codigo, ctx.cargaConsultar(t.k2).codigo], ["acceso", "acceso"]);
});

console.log("\n" + (fallas ? fallas + " de " + total + " pasos fallaron" : "Los " + total + " pasos coinciden con el manual"));
process.exit(fallas ? 1 : 0);
