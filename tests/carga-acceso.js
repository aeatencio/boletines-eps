#!/usr/bin/env node
/**
 * Pruebas locales del acceso individual por enlace (#28): Carga_Acceso.gs, identificarDocente_ (Carga_Web.gs) y el
 * gancho «acceso» del núcleo, con todos los .gs reales sobre los dobles de tests/carga-dobles.js. Datos y secreto
 * ficticios; sólo Node.js. No reemplazan el piloto en Google.
 *
 *   node tests/carga-acceso.js
 */
"use strict";
const crypto = require("crypto");
const D = require("./carga-dobles.js");

const pruebas = [];
function seccion(t) { pruebas.push({ seccion: t }); }
function prueba(nombre, fn) { pruebas.push({ nombre, fn }); }
function igual(a, b, msg) { const x = JSON.stringify(a), y = JSON.stringify(b); if (x !== y) throw new Error((msg || "distinto") + ": " + x + " ≠ " + y); }
function cierto(x, msg) { if (!x) throw new Error(msg || "se esperaba verdadero"); }
const json = x => JSON.parse(JSON.stringify(x));

const MATE = D.DOCENTES.matematica.correo, LENGUA = D.DOCENTES.lengua.correo, REEMPLAZO = D.DOCENTES.reemplazo.correo;
const DIA = 86400000;
const enlace = (S, correo) => { const m = D.enlacesReales(S.ctx, [correo]); return Object.keys(m)[0] || ""; };
const grillas = S => Object.values(S.s.grillas).flatMap(g => Object.values(g));
const intentos = S => grillas(S).reduce((n, h) => n + h.intentos, 0);
const filasRegistro = S => S.s.admin.Registro.getLastRow() - 1;
/** Un pedido válido de Calificación para un pendiente de Matemática II en 2C, a partir de la consulta. */
function pedidoPendiente(S, credencial, op) {
  const v = S.ctx.cargaConsultar(credencial);
  igual(v.estado, "ok", "consulta");
  const p = v.periodos.find(x => x.periodo === D.P2C), m = p.materias.find(x => x.materia === "Matemática II");
  const e = m.estudiantes.find(x => x.situacion === "pendiente" && !x.problema);
  return { operacion: op || "op-acceso-" + crypto.randomBytes(6).toString("hex"), periodo: p.periodo, fuente: p.fuente, orientacion: m.orientacion,
    materia: m.clave, dni: e.dni, campo: "calificacion", esperado: "CURSA", nuevo: "7" };
}
/** Ejecuta fn justo antes de que el guardado obtenga el lock (simula un cambio mientras espera). */
function alTomarLock(S, fn) {
  const lock = S.s.lock, original = lock.tryLock;
  lock.tryLock = function () { lock.tryLock = original; fn(); return original.apply(this, arguments); };
}
const enlaces = (S, filas) => { S.s.admin.Enlaces.v = [["Correo", "Versión"]].concat(filas); };

// ================================================================ 1. la credencial

seccion("1. La credencial: HMAC-SHA256 del secreto sobre correo, versión y tramo; nunca guardada");

prueba("es base64url sin relleno (43 caracteres) del HMAC-SHA256 con CLAVE_ENLACES sobre «contexto · correo · versión · tramo», calculado aparte con Node", () => {
  const S = D.servidor();
  const k = enlace(S, MATE);
  cierto(/^[A-Za-z0-9_-]{43}$/.test(k), k);
  const s = S.ctx.serviciosCarga_();
  const config = S.ctx.configuracion_(s.datos(), new Date());
  const tramo = S.ctx.tramoAcceso_(config, new Date());
  const esperado = crypto.createHmac("sha256", D.CLAVE_FICTICIA).update(["boletines-eps/carga/enlace/v1", MATE, "1", tramo.desde].join("\n")).digest("base64url");
  igual(k, esperado);
  igual(tramo.desde, config.periodos[D.P2C].ventana.desde, "2C y anual comparten ventana: el tramo empieza en su «Desde»");
});

prueba("cada docente tiene su credencial; el correo no está en ella; ni la planilla de administración ni el Registro la guardan", async () => {
  const S = D.servidor();
  const a = enlace(S, MATE), b = enlace(S, LENGUA);
  cierto(a && b && a !== b);
  cierto(!a.includes("matematica") && !Buffer.from(a, "base64url").toString("latin1").includes("matematica"));
  const r = S.ctx.cargaGuardar(a, pedidoPendiente(S, a));
  igual(r.estado, "guardado");
  const todo = JSON.stringify(Object.values(S.s.admin).map(h => h.v)) + S.consola.join("\n");
  cierto(!todo.includes(a) && !todo.includes(b), "la credencial aparece en una hoja o en el registro de ejecución");
});

prueba("credenciales falsas: 2000 al azar con la forma correcta, la de otro docente con un carácter cambiado y la de otro tramo se rechazan", () => {
  const S = D.servidor();
  const a = enlace(S, MATE);
  for (let i = 0; i < 2000; i++) igual(S.ctx.cargaConsultar(crypto.randomBytes(32).toString("base64url")).codigo, "acceso");
  const cambiada = a.slice(0, 20) + (a[20] === "A" ? "B" : "A") + a.slice(21);
  igual(S.ctx.cargaConsultar(cambiada).codigo, "acceso");
  const firmar = S.ctx.serviciosCarga_().firmar;
  igual(S.ctx.cargaConsultar(S.ctx.credencialDocente_(firmar, MATE, 1, { desde: "2020-03-01" })).codigo, "acceso");
  igual(S.ctx.cargaConsultar(a).estado, "ok");
  igual(intentos(S), 0);
});

prueba("igualesSeguro_: compara sin cortar en la primera diferencia y rechaza tipos y largos distintos", () => {
  const S = D.servidor();
  igual([S.ctx.igualesSeguro_("abc", "abc"), S.ctx.igualesSeguro_("abc", "abd"), S.ctx.igualesSeguro_("abc", "ab"), S.ctx.igualesSeguro_(null, null), S.ctx.igualesSeguro_("", "")],
    [true, false, false, false, true]);
});

// ================================================================ 2. estable durante la ventana, vence con el tramo siguiente

seccion("2. Estable durante la ventana; vence cuando empieza el tramo siguiente");

prueba("el mismo enlace vale todos los días de la ventana y en ambos períodos (2C y anual); después del cierre abre la página sin períodos", () => {
  const S = D.servidor();
  const k = enlace(S, MATE);
  const pedido = pedidoPendiente(S, k);
  const v0 = S.ctx.cargaConsultar(k);
  igual(v0.periodos.map(p => p.periodo), [D.P2C, D.PAN]);
  const cierre = new Date(v0.periodos[0].cierre).getTime();
  for (let t = Date.now(); t < cierre; t += DIA) {
    S.reloj.desfase = t - Date.now();
    igual(enlace(S, MATE), k, "cambió el enlace");
    igual(S.ctx.cargaConsultar(k).estado, "ok");
  }
  S.reloj.desfase = cierre + 60000 - Date.now();
  const despues = S.ctx.cargaConsultar(k);
  igual([despues.estado, despues.periodos.length], ["ok", 0]);
  igual([S.ctx.cargaGuardar(k, pedido).codigo, intentos(S)], ["ventana", 0], "el enlace vale, pero la ventana cerró");
});

prueba("cuando empieza la ventana siguiente (otro tramo), el enlace anterior deja de valer y vale el nuevo", () => {
  const S = D.servidor();
  const k = enlace(S, MATE);
  const cierre = new Date(S.ctx.cargaConsultar(k).periodos[0].cierre).getTime();
  const desde = D.isoBA(cierre + 20 * DIA), hasta = D.isoBA(cierre + 30 * DIA);
  S.s.admin.Ventanas.v.push(["1º bimestre", desde, hasta]);
  S.s.control.v[9][1] = D.URL(D.ID[D.P3B]); // grilla ficticia para 1B (B10)
  S.reloj.desfase = cierre + 10 * DIA - Date.now();
  igual(S.ctx.cargaConsultar(k).estado, "ok", "entre ventanas el enlace sigue abriendo la página");
  S.reloj.desfase = new Date(desde + "T00:00:00-03:00").getTime() + 3600000 - Date.now();
  igual(S.ctx.cargaConsultar(k).codigo, "acceso");
  const nuevo = enlace(S, MATE);
  cierto(nuevo && nuevo !== k);
  igual(S.ctx.cargaConsultar(nuevo).periodos.map(p => p.periodo), ["1º bimestre"]);
});

prueba("ventanas que se superponen forman un solo tramo: el enlace no cambia cuando cierra la primera y sigue la otra", () => {
  const S = D.servidor();
  const hoy = Date.now();
  // Anual empieza después de emitido el enlace de 2C y antes de su cierre.
  S.s.admin.Ventanas.v = [["Período", "Desde", "Hasta"], [D.P2C, D.isoBA(hoy - 3 * DIA), D.isoBA(hoy + 2 * DIA)], [D.PAN, D.isoBA(hoy + 1 * DIA), D.isoBA(hoy + 12 * DIA)]];
  const k = enlace(S, MATE);
  igual(S.ctx.cargaConsultar(k).periodos.map(p => p.periodo), [D.P2C]);
  S.reloj.desfase = 2 * DIA;
  igual([enlace(S, MATE), S.ctx.cargaConsultar(k).periodos.map(p => p.periodo)], [k, [D.P2C, D.PAN]], "abre anual: el enlace no cambia");
  S.reloj.desfase = 8 * DIA;
  const v = S.ctx.cargaConsultar(k);
  igual([v.estado, v.periodos.map(p => p.periodo)], ["ok", [D.PAN]]);
  // Si no se superpusieran (anual empieza después del cierre de 2C), serían tramos distintos.
  S.reloj.desfase = 0;
  S.s.admin.Ventanas.v[2][1] = D.isoBA(hoy + 7 * DIA);
  const k2 = enlace(S, MATE);
  S.reloj.desfase = 8 * DIA;
  igual(S.ctx.cargaConsultar(k2).codigo, "acceso");
});

prueba("tramoAcceso_: sin ventanas o antes de la primera apertura no hay tramo; una ventana futura no cambia el vigente", () => {
  const S = D.servidor();
  const s = S.ctx.serviciosCarga_();
  const datos = s.datos();
  const ahora = new Date();
  const t = S.ctx.tramoAcceso_(S.ctx.configuracion_(datos, ahora), ahora);
  datos.ventanas = datos.ventanas.concat([["1º bimestre", D.isoBA(Date.now() + 60 * DIA), D.isoBA(Date.now() + 70 * DIA)]]);
  igual(json(S.ctx.tramoAcceso_(S.ctx.configuracion_(datos, ahora), ahora)), json(t));
  datos.ventanas = [datos.ventanas[0]];
  igual(S.ctx.tramoAcceso_(S.ctx.configuracion_(datos, ahora), ahora), null);
});

// ================================================================ 3. revocación

seccion("3. Revocación con la hoja Enlaces");

prueba("subir la versión revoca el enlace en la llamada siguiente (consulta y guardado), sin leer grillas ni escribir; el nuevo enlace funciona", () => {
  const S = D.servidor();
  const viejo = enlace(S, MATE);
  const pedido = pedidoPendiente(S, viejo);
  enlaces(S, [[MATE, 2]]);
  const antes = S.estado.aperturas.length;
  igual(S.ctx.cargaConsultar(viejo).codigo, "acceso");
  const g = S.ctx.cargaGuardar(viejo, pedido);
  igual([g.estado, g.codigo, intentos(S), filasRegistro(S)], ["rechazado", "acceso", 0, 0]);
  igual(S.estado.aperturas.slice(antes).filter(id => id !== D.ADMIN_ID && id !== D.GENERADOR_ID), [], "no abrió grillas");
  const nuevo = enlace(S, MATE);
  cierto(nuevo !== viejo);
  igual(S.ctx.cargaGuardar(nuevo, pedido).estado, "guardado");
  // Revocar a un docente no afecta a los demás.
  igual(S.ctx.cargaConsultar(enlace(S, LENGUA)).estado, "ok");
});

prueba("REVOCACIÓN MIENTRAS EL GUARDADO ESPERA EL LOCK: se rechaza («acceso») sin escribir la grilla ni el Registro; revocar a otro docente no lo afecta", () => {
  const S = D.servidor();
  const k = enlace(S, MATE);
  const pedido = pedidoPendiente(S, k);
  alTomarLock(S, () => enlaces(S, [[MATE, 2]]));
  const r = S.ctx.cargaGuardar(k, pedido);
  igual([r.estado, r.codigo, intentos(S), filasRegistro(S)], ["rechazado", "acceso", 0, 0]);
  igual(S.s.lock.tomado, false, "el lock se liberó");
  // Otro docente revocado durante la espera: el guardado sigue.
  const S2 = D.servidor();
  const k2 = enlace(S2, MATE);
  const p2 = pedidoPendiente(S2, k2);
  alTomarLock(S2, () => enlaces(S2, [[LENGUA, 5]]));
  igual(S2.ctx.cargaGuardar(k2, p2).estado, "guardado");
});

prueba("mientras espera el lock: si sacan al docente de todas sus asignaciones, «acceso»; si le reasignan sólo esa materia, «asignacion»; nunca se escribe", () => {
  const S = D.servidor();
  const k = enlace(S, MATE);
  const pedido = pedidoPendiente(S, k);
  alTomarLock(S, () => { S.s.admin.Asignaciones.v = S.s.admin.Asignaciones.v.filter(f => f[4] !== MATE); });
  igual([S.ctx.cargaGuardar(k, pedido).codigo, intentos(S)], ["acceso", 0]);
  const S2 = D.servidor();
  const k2 = enlace(S2, MATE);
  const p2 = pedidoPendiente(S2, k2);
  alTomarLock(S2, () => { S2.s.admin.Asignaciones.v[1][3] = "Docente Ficticio de Reemplazo"; S2.s.admin.Asignaciones.v[1][4] = REEMPLAZO; });
  igual([S2.ctx.cargaGuardar(k2, p2).codigo, intentos(S2), filasRegistro(S2)], ["asignacion", 0, 0]);
});

prueba("un reintento de una operación ya aplicada, con el enlace revocado, se rechaza e informa el estado previo del Registro (sin escribir)", () => {
  const S = D.servidor();
  const k = enlace(S, MATE);
  const pedido = pedidoPendiente(S, k);
  igual(S.ctx.cargaGuardar(k, pedido).estado, "guardado");
  enlaces(S, [[MATE, 2]]);
  const r = S.ctx.cargaGuardar(k, pedido);
  igual([r.estado, r.codigo, intentos(S)], ["rechazado", "acceso", 1]);
});

prueba("Enlaces: versión 1 explícita = sin fila; filas repetidas valen por la mayor; correo en mayúsculas o con espacios se reconoce", () => {
  const S = D.servidor();
  const base = enlace(S, MATE);
  enlaces(S, [[MATE, 1]]);
  igual(enlace(S, MATE), base);
  enlaces(S, [[" " + MATE.toUpperCase() + " ", 3], [MATE, 2]]);
  const v3 = enlace(S, MATE);
  enlaces(S, [[MATE, 3]]);
  igual(enlace(S, MATE), v3);
  igual(S.ctx.cargaConsultar(base).codigo, "acceso");
});

prueba("fallas cerradas: versión ilegible bloquea a ese docente; sin la hoja Enlaces (o sin sus columnas) ningún enlace vale; verificar lo informa", () => {
  const S = D.servidor();
  const k = enlace(S, MATE), l = enlace(S, LENGUA);
  enlaces(S, [[MATE, "dos"]]);
  igual([S.ctx.cargaConsultar(k).codigo, enlace(S, MATE)], ["acceso", ""]);
  igual(S.ctx.cargaConsultar(l).estado, "ok", "los demás siguen");
  enlaces(S, [[MATE, 0]]);
  igual(S.ctx.cargaConsultar(k).codigo, "acceso");
  enlaces(S, [["docente.inexistente@ficticio.example", 2], ["no-es-un-correo", 2]]);
  const problemas = S.ctx.verificarCargaDocente().problemas.join("\n");
  cierto(/no figura en Asignaciones/.test(problemas) && /correo no reconocido/.test(problemas), problemas);
  delete S.s.admin.Enlaces;
  igual([S.ctx.cargaConsultar(k).codigo, S.ctx.cargaConsultar(l).codigo], ["acceso", "acceso"]);
  cierto(/Falta la hoja Enlaces/.test(S.ctx.verificarCargaDocente().problemas.join("\n")));
  S.s.admin.Enlaces = new D.Hoja([["Correo", "Otra"]]);
  igual(S.ctx.cargaConsultar(l).codigo, "acceso");
});

prueba("cambiar el secreto invalida todos los enlaces a la vez; sin secreto (o muy corto), la consulta da error y el guardado «reintentar», sin escribir", () => {
  const S = D.servidor();
  const k = enlace(S, MATE);
  const pedido = pedidoPendiente(S, k);
  S.estado.props.CLAVE_ENLACES = "otra-clave-ficticia-0123456789abcdef0123456789";
  igual(S.ctx.cargaConsultar(k).codigo, "acceso");
  delete S.estado.props.CLAVE_ENLACES;
  igual(S.ctx.cargaConsultar(k).estado, "error");
  igual(S.ctx.cargaGuardar(k, pedido).estado, "reintentar");
  S.estado.props.CLAVE_ENLACES = "corta";
  igual(S.ctx.cargaConsultar(k).estado, "error");
  igual([intentos(S), filasRegistro(S)], [0, 0]);
});

// ================================================================ 4. costo

seccion("4. Costo: la identificación no agrega lecturas de la administración");

prueba("consulta: una lectura de Asignaciones (la de la identificación, reutilizada); guardado: dos (antes del lock y dentro); una credencial mal formada, ninguna", () => {
  const S = D.servidor();
  const k = enlace(S, MATE);
  const pedido = pedidoPendiente(S, k);
  const h = S.s.admin.Asignaciones, original = h.getDataRange.bind(h);
  let lecturas = 0;
  h.getDataRange = () => { lecturas++; return original(); };
  S.ctx.cargaConsultar(k);
  igual(lecturas, 1, "consulta");
  lecturas = 0;
  igual(S.ctx.cargaGuardar(k, pedido).estado, "guardado");
  igual(lecturas, 2, "guardado");
  lecturas = 0;
  S.ctx.cargaConsultar("no-tiene-forma");
  igual(lecturas, 0);
  // Firmas por llamada: una por docente asignado al identificar, más una por cada revalidación.
  const antes = S.estado.hmac;
  S.ctx.cargaConsultar(k);
  igual(S.estado.hmac - antes, 2, "dos docentes asignados en el escenario");
});

// ================================================================ ejecución

(async () => {
  let fallas = 0, total = 0;
  for (const x of pruebas) {
    if (x.seccion) { console.log("\n# " + x.seccion); continue; }
    total++;
    try { await x.fn(); console.log("OK    " + x.nombre); }
    catch (e) { fallas++; console.log("FALLA " + x.nombre + "\n      " + String(e && e.stack || e).split("\n").slice(0, 6).join("\n      ")); }
  }
  console.log("\n" + (fallas ? fallas + " de " + total + " pruebas fallaron" : "Las " + total + " pruebas pasaron"));
  process.exit(fallas ? 1 : 0);
})();
