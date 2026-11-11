#!/usr/bin/env node
/**
 * Pruebas locales de la interfaz docente (#27): Carga_Web.gs, Carga_Cliente.html y su integración con el núcleo.
 * Datos ficticios; sólo Node.js. No reemplazan la prueba en Google ni en un navegador real.
 *
 *   node tests/carga-interfaz.js
 *
 * 1. Límite de identidad (#28): sin el reemplazo de identificarDocente_, nada se lee ni se escribe; la página nunca
 *    elige la identidad.
 * 2. Reglas del cliente iguales a las del núcleo (paridad).
 * 3. Cola de guardado, con un transporte simulado: identidad de cada operación en los reintentos, estados y avisos.
 * 4. Cola + servidor real (los tres .gs) con dobles de Google y fallas inyectadas.
 */
"use strict";
const vm = require("vm");
const D = require("./carga-dobles.js");

const K = D.cliente();
const pruebas = [];
let seccionActual = "";
function seccion(t) { pruebas.push({ seccion: t }); }
function prueba(nombre, fn) { pruebas.push({ nombre, fn }); }
function igual(a, b, msg) { const x = JSON.stringify(a), y = JSON.stringify(b); if (x !== y) throw new Error((msg || "distinto") + ": " + x + " ≠ " + y); }
function cierto(x, msg) { if (!x) throw new Error(msg || "se esperaba verdadero"); }
const drenar = async (n) => { for (let i = 0; i < (n || 30); i++) await new Promise(r => setImmediate(r)); };
const json = x => JSON.parse(JSON.stringify(x));

// ================================================================ utilidades de la cola

/** Transporte simulado: cada respuesta es un objeto, una función(pedido) o 'falla' (sin respuesta del servidor). */
function transporte(respuestas, porDefecto) {
  const envios = [];
  const enVuelo = { ahora: 0, max: 0 };
  const f = pedido => {
    envios.push(json(pedido));
    enVuelo.ahora++; enVuelo.max = Math.max(enVuelo.max, enVuelo.ahora);
    const r = respuestas.length ? respuestas.shift() : (porDefecto || (p => ({ estado: "guardado", actual: K.visto(p.campo, p.nuevo) })));
    const x = typeof r === "function" ? r(pedido) : r;
    return new Promise((ok, mal) => setImmediate(() => { enVuelo.ahora--; x === "falla" ? mal(new Error("red caída (simulada)")) : ok(json(x)); }));
  };
  f.envios = envios; f.enVuelo = enVuelo;
  return f;
}
/** Transporte que el test resuelve a mano: sirve para cambiar el valor mientras un pedido está en vuelo. */
function transporteManual() {
  const pendientes = [];
  const f = pedido => new Promise((ok, mal) => pendientes.push({ pedido: json(pedido), ok: r => ok(json(r)), mal }));
  f.pendientes = pendientes;
  return f;
}
let ids = 0;
function nuevaCola(enviar, extra) {
  const rechazos = [];
  const cola = K.crearCola(Object.assign({
    enviar, esperar: () => Promise.resolve(), nuevoId: () => "op-test-" + String(++ids).padStart(6, "0"), ahora: () => Date.now(),
    alCambiar: () => {}, alRechazo: (clave, r, p) => rechazos.push({ clave, r, p })
  }, extra || {}));
  cola.rechazos = rechazos;
  return cola;
}
const BASE = { periodo: D.P2C, fuente: D.ID[D.P2C], orientacion: "EPSI", materia: "MATEMATICAII", dni: "90000001", campo: "calificacion" };
const baseDe = (cambios) => Object.assign({}, BASE, cambios || {});
const clave = b => K.claveCampo(b);
const campos = ["calificacion", "calidad", "participacion", "observaciones"];

// ================================================================ 1. límite de identidad (#28)

seccion("1. Límite de identidad (#28): la página nunca elige la identidad; sin #28, nada se lee ni se escribe");

prueba("sin el reemplazo de #28, toda credencial se rechaza («acceso») sin abrir ninguna planilla ni escribir: vacía, un correo asignado, un token, un objeto, muy larga", () => {
  const S = D.servidor();
  const credenciales = [undefined, null, "", "   ", D.DOCENTES.matematica.correo, "Zx81y_token-valido-de-aspecto", { correo: D.DOCENTES.matematica.correo }, "a".repeat(5000)];
  credenciales.forEach(c => {
    const r = S.ctx.cargaConsultar(c);
    igual([r.estado, r.codigo], ["rechazado", "acceso"], "consulta con " + JSON.stringify(c));
    const g = S.ctx.cargaGuardar(c, { operacion: "op-identidad-0001", periodo: D.P2C, fuente: D.ID[D.P2C], orientacion: "EPSI", materia: "MATEMATICAII",
      dni: "90000001", campo: "calificacion", esperado: "CURSA", nuevo: "7", correo: D.DOCENTES.matematica.correo });
    igual([g.estado, g.codigo], ["rechazado", "acceso"], "guardado con " + JSON.stringify(c));
  });
  igual(S.estado.aperturas, [], "no se abrió ninguna planilla");
  igual(Object.values(S.s.grillas).flatMap(g => Object.values(g)).reduce((n, h) => n + h.intentos, 0), 0);
  igual(S.s.admin.Registro.getLastRow(), 1, "Registro sin filas");
});

prueba("con credenciales locales: la identidad sale de la credencial; un «correo» en el pedido no llega al núcleo y el Registro atribuye el cambio al dueño de la credencial", () => {
  const S = D.servidor();
  const mapa = D.sustituirIdentidad(S.ctx, [D.DOCENTES.lengua.correo, D.DOCENTES.matematica.correo]);
  const deLengua = Object.keys(mapa).find(k => mapa[k] === D.DOCENTES.lengua.correo);
  const deMate = Object.keys(mapa).find(k => mapa[k] === D.DOCENTES.matematica.correo);
  igual(S.ctx.cargaConsultar("credencial-inventada").codigo, "acceso");
  const vista = S.ctx.cargaConsultar(deMate);
  const p = vista.periodos[0], m = p.materias.find(x => x.materia === "Matemática II");
  const e = m.estudiantes.find(x => x.situacion === "pendiente");
  const pedido = { operacion: "op-identidad-0002", periodo: p.periodo, fuente: p.fuente, orientacion: m.orientacion, materia: m.clave, dni: e.dni,
    campo: "calificacion", esperado: "CURSA", nuevo: "8", correo: D.DOCENTES.lengua.correo, docente: "otro" };
  // Lengua no puede guardar en Matemática II aunque nombre el correo de Matemática en el pedido.
  const ajeno = S.ctx.cargaGuardar(deLengua, Object.assign({}, pedido, { correo: D.DOCENTES.matematica.correo }));
  igual([ajeno.estado, ajeno.codigo], ["rechazado", "asignacion"]);
  const propio = S.ctx.cargaGuardar(deMate, pedido);
  igual(propio.estado, "guardado");
  const fila = S.s.admin.Registro.getDataRange().getDisplayValues()[1];
  const COL = D.cargaDe(S.ctx).REGISTRO;
  igual([fila[COL.indexOf("Correo")], fila[COL.indexOf("Estado")]], [D.DOCENTES.matematica.correo, "aplicado"]);
});

prueba("pedidoDePagina_ deja sólo los nueve campos conocidos y sólo si son texto; un pedido que no es objeto llega como nulo y el núcleo lo rechaza", () => {
  const S = D.servidor();
  const limpio = json(S.ctx.pedidoDePagina_({ operacion: "op-x-000001", periodo: "Anual", fuente: "F", orientacion: "EPSI", materia: "M", dni: "1", campo: "calidad",
    esperado: "", nuevo: "Buena", correo: "x@y.z", __proto__: { nuevo: "otra" }, extra: 1, dniNumero: 5 }));
  igual(Object.keys(limpio).sort(), ["campo", "dni", "esperado", "fuente", "materia", "nuevo", "operacion", "orientacion", "periodo"]);
  igual(json(S.ctx.pedidoDePagina_({ esperado: 7, nuevo: null })), {});
  igual(S.ctx.pedidoDePagina_("texto"), null);
  const mapa = D.sustituirIdentidad(S.ctx, [D.DOCENTES.matematica.correo]);
  const r = S.ctx.cargaGuardar(Object.keys(mapa)[0], "no es un pedido");
  igual([r.estado, r.codigo], ["rechazado", "pedido"]);
});

prueba("si la verificación de identidad falla con una excepción: la consulta informa un error y el guardado responde «reintentar», sin escribir", () => {
  const S = D.servidor();
  vm.runInContext("identificarDocente_ = function () { throw new Error('servicio de enlaces caído (simulado)'); };", S.ctx);
  igual(S.ctx.cargaConsultar("x").estado, "error");
  igual(S.ctx.cargaGuardar("x", { operacion: "op-x-000002" }).estado, "reintentar");
  igual(S.estado.aperturas, []);
});

prueba("funciones invocables desde la página: sólo doGet, cargaConsultar y cargaGuardar, más las tres del propietario (que exigen ser el propietario)", () => {
  const publicas = ["src/carga/Carga_Nucleo.gs", "src/carga/Carga_Google.gs", "src/carga/Carga_Web.gs"]
    .flatMap(a => [...D.leer(a).matchAll(/^function\s+([A-Za-z0-9_$]+)\s*\(/gm)].map(m => m[1]))
    .filter(n => !n.endsWith("_")).sort();
  igual(publicas, ["cargaConsultar", "cargaGuardar", "doGet", "prepararAdministracionCarga", "resolverPendientesCarga", "verificarCargaDocente"]);
  const S = D.servidor();
  S.estado.usuario = "visitante@ficticio.example";
  ["prepararAdministracionCarga", "verificarCargaDocente", "resolverPendientesCarga"].forEach(f => {
    let error = null; try { S.ctx[f](); } catch (e) { error = e; }
    cierto(error && /propietario/.test(error.message), f + " sin control de propietario");
  });
});

prueba("doGet: la credencial del enlace llega escapada a la página (sin inyección), se descarta si supera 512 caracteres, y la página incluye su lógica", () => {
  const S = D.servidor();
  const ataque = '"><script>alert(1)</script>';
  const html = S.ctx.doGet({ parameter: { k: ataque } }).contenido;
  cierto(html.includes('data-credencial="&quot;&gt;&lt;script&gt;alert(1)&lt;/script&gt;"'), "credencial sin escapar");
  cierto(!html.includes("<script>alert(1)"), "inyección");
  cierto(html.includes("var CargaCliente = (function"), "falta Carga_Cliente");
  cierto(S.ctx.doGet({ parameter: { k: "x".repeat(600) } }).contenido.includes('data-credencial=""'));
  cierto(S.ctx.doGet({}).contenido.includes('data-credencial=""'));
});

prueba("la página no carga recursos externos (la credencial no viaja en ningún Referer) y sus pedidos sólo llevan credencial y pedido", () => {
  const html = D.renderPagina("cred");
  igual([...html.matchAll(/\b(?:src|href)\s*=\s*["']?(?:https?:)?\/\//gi)].length, 0, "recursos externos");
  igual([...html.matchAll(/llamar\('([a-zA-Z]+)'([^)]*)\)/g)].map(m => m[1] + m[2]), ["cargaGuardar, CREDENCIAL, pedido", "cargaConsultar, CREDENCIAL"]);
  cierto(!/correo/i.test(D.leer("src/carga/Carga_Cliente.html")), "el cliente no maneja correos");
});

// ================================================================ 2. paridad con el núcleo

seccion("2. Reglas del cliente iguales a las del núcleo");

prueba("situación de la Calificación: igual que situacionCalificacion_ en los cinco períodos, con grafías, acentos y espacios", () => {
  const S = D.servidor();
  const CARGA = D.cargaDe(S.ctx);
  const valores = ["", "  ", "CURSA", "cursa", " Cursa ", "NO CURSA", "no  cursa", "ACREDITÓ", "acredito", "Acreditó", "EQUIVALENCIA", "Equivalencia",
    "7", "07", "7,00", "10", "1", "0", "11", "En proceso", "en proceso", "EN PROCESO", "Suficiente", "Avanzado", "avanzado", "Muy bien", "Aprobado", "x"];
  CARGA.PERIODOS.forEach(p => valores.forEach(v => igual(K.situacion(CARGA.ESCALA_DOCENTE[p], v), S.ctx.situacionCalificacion_(p, v), p + " «" + v + "»")));
});

prueba("Observaciones: el valor que el cliente espera ver es el que el núcleo escribe (espacios, tabulaciones y renglones)", () => {
  const S = D.servidor();
  ["", "  hola  ", "dos  espacios", "renglón\nnuevo", "tab\tulado", " =1+1 ", "'apóstrofo", "1/2"].forEach(t =>
    igual(K.visto("observaciones", t), S.ctx.normalizarNuevo_("observaciones", D.P2C, t).visto, JSON.stringify(t)));
  ["Buena", ""].forEach(t => igual(K.visto("calidad", t), S.ctx.normalizarNuevo_("calidad", D.P2C, t).visto));
});

prueba("identificadores de operación: cumplen CARGA.OPERACION y no se repiten en 5000", () => {
  const S = D.servidor();
  const re = D.cargaDe(S.ctx).OPERACION;
  const vistos = new Set();
  for (let i = 0; i < 5000; i++) { const id = K.nuevoIdOperacion(); cierto(re.test(id), id); vistos.add(id); }
  igual(vistos.size, 5000);
});

prueba("modelo: 2C y anual son materias distintas con claves distintas (la clave incluye período y fuente); se marcan los cursantes de otros grupos", () => {
  const S = D.servidor();
  const mapa = D.sustituirIdentidad(S.ctx, [D.DOCENTES.matematica.correo]);
  const m = K.modelo(S.ctx.cargaConsultar(Object.keys(mapa)[0]));
  igual(m.periodos.map(p => p.periodo), [D.P2C, D.PAN]);
  const [a, b] = m.periodos.map(p => p.materias.find(x => x.materia === "Matemática II"));
  cierto(a.id !== b.id);
  const ea = a.estudiantes[0], eb = b.estudiantes.find(x => x.dni === ea.dni);
  campos.forEach(c => cierto(ea.claves[c] !== eb.claves[c] && ea.claves[c].includes(D.ID[D.P2C]) && eb.claves[c].includes(D.ID[D.PAN])));
  const otros = a.estudiantes.filter(e => e.otroGrupo);
  cierto(otros.length >= 2 && otros.every(e => e.grupo !== "EPS2"), "cursantes de otros grupos: " + otros.map(e => e.grupo));
  igual(m.maxObservaciones, 1000);
  const c = K.contar(m.periodos[0], a, e => e.valores.calificacion);
  igual([c.pendiente, c.administrativo, c["a revisar"]], [5, 2, 1]);
  // Pendientes primero
  const orden = K.ordenar(m.periodos[0], a.estudiantes, e => e.valores.calificacion);
  cierto(orden.slice(0, 5).every(e => e.valores.calificacion === "CURSA") && orden[5].valores.calificacion !== "CURSA");
});

// ================================================================ 3. cola de guardado

seccion("3. Cola de guardado (transporte simulado)");

prueba("guardado simple: «esperado» es el valor de la grilla, «nuevo» el elegido; sólo «guardado» confirma", async () => {
  const t = transporte([]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  igual(cola.campo(clave(BASE)).estado, "en-espera");
  await drenar();
  const c = cola.campo(clave(BASE));
  igual([c.estado, c.confirmado, t.envios.length], ["guardado", "7", 1]);
  igual(Object.keys(t.envios[0]).sort(), ["campo", "dni", "esperado", "fuente", "materia", "nuevo", "operacion", "orientacion", "periodo"]);
  igual([t.envios[0].esperado, t.envios[0].nuevo], ["CURSA", "7"]);
  cierto(Object.isFrozen(c.ultimo), "el pedido se congela");
});

prueba("elegir el valor que ya tiene la grilla no envía nada (en Calificación, sin importar la grafía)", async () => {
  const t = transporte([]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "avanzado");
  cola.cambiar(clave(BASE), "Avanzado");
  const b2 = baseDe({ campo: "observaciones" });
  cola.registrar(clave(b2), b2, "texto con espacios");
  cola.cambiar(clave(b2), "  texto   con\nespacios ");
  await drenar();
  igual(t.envios.length, 0);
});

prueba("cambios rápidos mientras un guardado está en vuelo: sólo se envía el último, después, con «esperado» = lo que confirmó el servidor", async () => {
  const t = transporteManual();
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "5");
  await drenar();
  cola.cambiar(clave(BASE), "6"); cola.cambiar(clave(BASE), "8"); cola.cambiar(clave(BASE), "9");
  igual(t.pendientes.length, 1);
  t.pendientes[0].ok({ estado: "guardado", actual: "5" });
  await drenar();
  igual(t.pendientes.length, 2);
  igual([t.pendientes[1].pedido.esperado, t.pendientes[1].pedido.nuevo], ["5", "9"]);
  cierto(t.pendientes[1].pedido.operacion !== t.pendientes[0].pedido.operacion, "operación nueva");
  t.pendientes[1].ok({ estado: "guardado", actual: "9" });
  await drenar();
  igual([cola.campo(clave(BASE)).estado, cola.campo(clave(BASE)).confirmado], ["guardado", "9"]);
});

prueba("«reintentar» se reintenta solo con la MISMA operación, byte a byte; al agotarse queda «sin-enviar» y el reintento manual también repite el mismo pedido", async () => {
  const R = { estado: "reintentar", motivo: "Hay muchos guardados al mismo tiempo. Se reintenta." };
  const t = transporte([R, R, R, R, R]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  igual([cola.campo(clave(BASE)).estado, t.envios.length], ["sin-enviar", 4]);
  cola.reintentar(clave(BASE));
  await drenar();
  igual([cola.campo(clave(BASE)).estado, t.envios.length], ["guardado", 6]);
  cierto(t.envios.every(e => JSON.stringify(e) === JSON.stringify(t.envios[0])), "todos los envíos son el mismo pedido");
});

prueba("sin respuesta del servidor: reintentos con la misma operación; si sigue sin respuesta, «sin-confirmar» (nunca «guardado» ni un pedido distinto)", async () => {
  const t = transporte(["falla", "falla", "falla", "falla"]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  const c = cola.campo(clave(BASE));
  igual([c.estado, c.propuesto, c.confirmado, cola.hayPendientes()], ["sin-confirmar", "7", "CURSA", true]);
  cola.reintentar(clave(BASE));
  await drenar();
  igual(cola.campo(clave(BASE)).estado, "guardado");
  igual(new Set(t.envios.map(e => JSON.stringify(e))).size, 1);
});

prueba("sin-confirmar y el docente elige otro valor: primero se reenvía la operación pendiente tal cual y después la nueva, con «esperado» actualizado", async () => {
  const t = transporte(["falla", "falla", "falla", "falla"]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  igual(cola.campo(clave(BASE)).estado, "sin-confirmar");
  cola.cambiar(clave(BASE), "9");
  await drenar();
  igual(t.envios.length, 6);
  igual(t.envios[4], t.envios[0], "el reenvío es idéntico");
  igual([t.envios[5].esperado, t.envios[5].nuevo], ["7", "9"]);
  cierto(t.envios[5].operacion !== t.envios[0].operacion);
  igual(cola.campo(clave(BASE)).confirmado, "9");
});

prueba("sin-enviar (todas las respuestas fueron «reintentar», nunca intentó escribir) y el docente elige otro valor: la operación vieja se abandona y no se reenvía", async () => {
  const R = { estado: "reintentar" };
  const t = transporte([R, R, R, R]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  igual(cola.campo(clave(BASE)).estado, "sin-enviar");
  cola.cambiar(clave(BASE), "9");
  await drenar();
  igual(t.envios.length, 5);
  igual([t.envios[4].esperado, t.envios[4].nuevo], ["CURSA", "9"]);
});

prueba("«incierto»: se vuelve a consultar con la misma operación; si sigue incierto, se muestra el valor de la grilla y el del docente, sin reenviar solo", async () => {
  const I = { estado: "incierto", actual: "CURSA", motivo: "No se pudo confirmar el guardado. La grilla muestra el valor actual." };
  const t = transporte([I, I, I]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  let c = cola.campo(clave(BASE));
  igual([c.estado, c.confirmado, c.propuesto, t.envios.length, cola.hayPendientes()], ["incierto", "CURSA", "7", 3, false]);
  igual(new Set(t.envios.map(e => JSON.stringify(e))).size, 1);
  await drenar();
  igual(t.envios.length, 3, "no se reenvía solo");
  // «Volver a cargar» es una decisión del docente: operación nueva, con «esperado» = lo que muestra la grilla.
  cola.usarPropuesto(clave(BASE));
  await drenar();
  c = cola.campo(clave(BASE));
  igual([c.estado, t.envios[3].esperado, t.envios[3].nuevo], ["guardado", "CURSA", "7"]);
  cierto(t.envios[3].operacion !== t.envios[0].operacion);
});

prueba("«incierto» que se aclara al volver a consultar: «coincide» se informa como tal, nunca como «guardado»", async () => {
  const t = transporte([{ estado: "incierto", actual: "CURSA" }, { estado: "coincide", actual: "7" }]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  igual([cola.campo(clave(BASE)).estado, cola.campo(clave(BASE)).confirmado], ["coincide", "7"]);
  igual(t.envios[0], t.envios[1]);
});

prueba("«incierto» sin valor actual: un cambio del docente primero vuelve a consultar la operación (que sólo observa) y recién después decide", async () => {
  const t = transporte([{ estado: "incierto" }, { estado: "incierto" }, { estado: "incierto" }, { estado: "incierto", actual: "CURSA" }]);
  const cola = nuevaCola(t, { esperasIncierto: [1, 1] });
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  igual([cola.campo(clave(BASE)).estado, cola.campo(clave(BASE)).confirmado], ["incierto", null]);
  cola.cambiar(clave(BASE), "8");
  await drenar();
  igual(t.envios[3], t.envios[0], "la verificación repite la operación");
  igual([t.envios[4].esperado, t.envios[4].nuevo], ["CURSA", "8"]);
});

prueba("conflicto: el campo muestra el valor de la grilla; ni el pedido ni un cambio en curso se reenvían solos; «Guardar mi valor» es una operación nueva", async () => {
  const t = transporteManual();
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  cola.cambiar(clave(BASE), "8"); // en vuelo
  t.pendientes[0].ok({ estado: "conflicto", actual: "6" });
  await drenar();
  let c = cola.campo(clave(BASE));
  igual([c.estado, c.confirmado, c.propuesto, t.pendientes.length], ["conflicto", "6", "8", 1]);
  cola.usarPropuesto(clave(BASE));
  await drenar();
  igual([t.pendientes[1].pedido.esperado, t.pendientes[1].pedido.nuevo], ["6", "8"]);
  t.pendientes[1].ok({ estado: "guardado", actual: "8" });
  await drenar();
  igual(cola.campo(clave(BASE)).estado, "guardado");
  // «Dejar» el valor de la grilla
  const b2 = baseDe({ campo: "calidad" });
  cola.registrar(clave(b2), b2, "");
  cola.cambiar(clave(b2), "Buena");
  await drenar();
  t.pendientes[2].ok({ estado: "conflicto", actual: "Regular" });
  await drenar();
  cola.descartar(clave(b2));
  c = cola.campo(clave(b2));
  igual([c.estado, c.confirmado, c.propuesto], ["listo", "Regular", null]);
});

prueba("rechazo por ventana cerrada: se informa al alcance «periodo»; lo que espera turno no sale y el cambio en curso queda sin enviar", async () => {
  const t = transporteManual();
  const rechazos = [];
  // Como Carga_Pagina: el rechazo bloquea su alcance en el momento.
  const cola = nuevaCola(t, { maxEnVuelo: 1, alRechazo: (clave, r, p) => { rechazos.push(r); cola.bloquear(b => b.periodo === p.periodo, r.motivo); } });
  cola.rechazos = rechazos;
  const otra = baseDe({ dni: "90000002" });
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.registrar(clave(otra), otra, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  cola.cambiar(clave(otra), "8");  // espera turno
  cola.cambiar(clave(BASE), "9");  // siguiente
  t.pendientes[0].ok({ estado: "rechazado", codigo: "ventana", motivo: "La carga de 2º cuatrimestre cerró." });
  await drenar();
  igual(rechazos.map(r => [r.codigo, K.alcanceRechazo(r.codigo)]), [["ventana", "periodo"]]);
  igual(t.pendientes.length, 1, "nada más salió");
  igual([cola.campo(clave(BASE)).estado, cola.campo(clave(BASE)).propuesto], ["rechazado", "9"]);
  igual(cola.campo(clave(otra)).estado, "no-enviado");
  igual(cola.cambiar(clave(BASE), "10"), false, "bloqueado");
  igual(cola.hayPendientes(), false);
});

prueba("rechazo local (valor inválido): el cambio siguiente del docente se procesa normalmente", async () => {
  const t = transporte([{ estado: "rechazado", codigo: "valor", motivo: "inválido" }]);
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "77");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  igual([cola.campo(clave(BASE)).estado, t.envios.length, t.envios[1].esperado], ["guardado", 2, "CURSA"]);
});

prueba("a lo sumo 3 guardados en vuelo a la vez (para no competir por el lock); los 12 terminan", async () => {
  const t = transporte([]);
  const cola = nuevaCola(t);
  const bases = Array.from({ length: 12 }, (_, i) => baseDe({ dni: String(90000100 + i) }));
  bases.forEach(b => cola.registrar(clave(b), b, "CURSA"));
  bases.forEach(b => cola.cambiar(clave(b), "8"));
  await drenar(60);
  igual([t.enVuelo.max, t.envios.length], [3, 12]);
  cierto(bases.every(b => cola.campo(clave(b)).estado === "guardado"));
});

prueba("actualizar desde una consulta no pisa una respuesta más nueva ni un campo con una operación abierta; sí actualiza un campo sin novedades", async () => {
  const t = transporteManual();
  const cola = nuevaCola(t);
  const b2 = baseDe({ dni: "90000002" }), b3 = baseDe({ dni: "90000003" });
  [BASE, b2, b3].forEach(b => cola.registrar(clave(b), b, "CURSA"));
  const desde = Date.now() - 1000; // la consulta salió antes de la respuesta
  cola.cambiar(clave(BASE), "7");
  cola.cambiar(clave(b2), "8");
  await drenar();
  t.pendientes[0].ok({ estado: "guardado", actual: "7" });
  await drenar();
  cola.registrar(clave(BASE), BASE, "CURSA", desde); // consulta vieja
  cola.registrar(clave(b2), b2, "CURSA", desde);     // operación abierta
  cola.registrar(clave(b3), b3, "4", desde);         // cambió en la grilla
  igual([cola.campo(clave(BASE)).confirmado, cola.campo(clave(b2)).pedido.nuevo, cola.campo(clave(b3)).confirmado], ["7", "8", "4"]);
});

prueba("el pedido congelado no se puede alterar entre reintentos (aunque el transporte lo intente)", async () => {
  let primero = null;
  const t = pedido => { if (!primero) { primero = pedido; try { pedido.nuevo = "1"; pedido.esperado = "8"; } catch (e) { /* congelado */ } return Promise.reject(new Error("red")); } return Promise.resolve({ estado: "guardado", actual: pedido.nuevo }); };
  const cola = nuevaCola(t);
  cola.registrar(clave(BASE), BASE, "CURSA");
  cola.cambiar(clave(BASE), "7");
  await drenar();
  igual([primero.nuevo, primero.esperado, cola.campo(clave(BASE)).confirmado], ["7", "CURSA", "7"]);
});

// ================================================================ 4. integración con el servidor real

seccion("3b. Indicador global de la barra");

const resumenDe = r => Object.assign({ enCurso: 0, sinConfirmar: 0, conflicto: 0, noGuardado: 0 }, r);

prueba("REGRESIÓN r1: con Observaciones escrita y sin enviar después de un guardado, la barra no dice «Todo guardado» sino «sin guardar»", () => {
  const ind = json(K.indicadorGlobal(resumenDe({}), 1, true));
  cierto(ind.texto !== "✓ Todo guardado", "dijo «Todo guardado» con un cambio sin enviar");
  igual(ind, { texto: "✎ 1 cambio sin guardar", clase: "aviso", avisos: 0 });
  igual(K.indicadorGlobal(resumenDe({}), 2, false).texto, "✎ 2 cambios sin guardar");
});

prueba("indicador global: «Todo guardado» sólo sin nada en curso, sin avisos y sin cambios por enviar (todas las combinaciones de 0 a 2)", () => {
  let casos = 0;
  [0, 1, 2].forEach(enCurso => [0, 1, 2].forEach(sinConfirmar => [0, 1, 2].forEach(conflicto => [0, 1, 2].forEach(noGuardado =>
    [0, 1, 2].forEach(sinGuardar => [false, true].forEach(hubo => {
      const ind = K.indicadorGlobal({ enCurso, sinConfirmar, conflicto, noGuardado }, sinGuardar, hubo);
      const nada = !enCurso && !sinConfirmar && !conflicto && !noGuardado && !sinGuardar;
      igual(ind.texto === "✓ Todo guardado", nada && hubo, JSON.stringify({ enCurso, sinConfirmar, conflicto, noGuardado, sinGuardar, hubo }));
      igual(ind.avisos, sinConfirmar + conflicto + noGuardado);
      if (nada && !hubo) igual(ind.texto, "");
      casos++;
    }))))));
  igual(casos, 486);
  igual(K.indicadorGlobal(resumenDe({ enCurso: 1 }), 1, true).texto, "Guardando 1 cambio…");
  igual(K.indicadorGlobal(resumenDe({ conflicto: 1 }), 1, true).texto, "⚠ 1 cambio necesita atención");
});

prueba("indicador global con la cola real: en curso, conflicto, rechazo, incierto y sin respuesta se informan; «Todo guardado» sólo cuando todo terminó bien", async () => {
  const t = transporteManual();
  const cola = nuevaCola(t, { esperas: [], esperasIncierto: [] });
  const bases = ["90000201", "90000202", "90000203", "90000204", "90000205"].map(dni => baseDe({ dni }));
  bases.forEach(b => cola.registrar(clave(b), b, "CURSA"));
  const ind = sinGuardar => K.indicadorGlobal(cola.resumen(), sinGuardar || 0, true);
  bases.forEach(b => cola.cambiar(clave(b), "7"));
  await drenar();
  igual([ind().texto, ind().clase], ["Guardando 5 cambios…", "curso"]);
  igual(t.pendientes.length, 3, "a lo sumo 3 en vuelo; el resto espera turno y también cuenta como en curso");
  t.pendientes[0].ok({ estado: "guardado", actual: "7" });
  t.pendientes[1].ok({ estado: "conflicto", actual: "5" });
  t.pendientes[2].ok({ estado: "rechazado", codigo: "valor", motivo: "inválido" });
  await drenar();
  igual(ind().texto, "Guardando 2 cambios…", "mientras algo está en curso, la barra lo dice");
  t.pendientes[3].ok({ estado: "incierto", actual: "CURSA" });
  t.pendientes[4].mal(new Error("sin conexión"));
  await drenar();
  igual(json(ind()), { texto: "⚠ 4 cambios necesitan atención", clase: "aviso", avisos: 4 });
  // El docente resuelve cada aviso aceptando el valor de la grilla o volviendo a cargar.
  [1, 2, 3].forEach(i => cola.descartar(clave(bases[i])));
  cola.reintentar(clave(bases[4]));
  await drenar();
  igual(ind().texto, "Guardando 1 cambio…");
  t.pendientes[5].ok({ estado: "guardado", actual: "7" });
  await drenar();
  igual(ind().texto, "✓ Todo guardado");
  igual(ind(1).texto, "✎ 1 cambio sin guardar", "lo escrito sin enviar sigue contando");
});

prueba("REGRESIÓN r1 (página): la barra usa indicadorGlobal con los cambios sin enviar (S.sucias) y se repinta al escribir en Observaciones", () => {
  const pagina = D.leer("src/carga/Carga_Pagina.html");
  const barra = pagina.match(/function pintarBarra\(\) \{([\s\S]*?)\n    \}/)[1];
  cierto(/K\.indicadorGlobal\(cola\.resumen\(\), Object\.keys\(S\.sucias\)\.length, S\.huboGuardados\)/.test(barra), "pintarBarra no considera S.sucias");
  cierto(!/'✓ Todo guardado'/.test(barra), "pintarBarra decide el texto por su cuenta");
  const entrada = pagina.match(/control\.addEventListener\('input', function \(\) \{([\s\S]*?)\n        \}\);/)[1];
  cierto(/S\.sucias\[clave\] = control\.value/.test(entrada) && /programarPintura\(\)/.test(entrada), "escribir en Observaciones no repinta la barra");
  const guardarObs = pagina.match(/function guardarObservacion\(clave, control\) \{([\s\S]*?)\n    \}/)[1];
  cierto(/delete S\.sucias\[clave\][\s\S]*programarPintura\(\)/.test(guardarObs), "al enviar Observaciones no se repinta la barra");
});

seccion("4. Cola + servidor real (Carga_Web.gs → Carga_Google.gs → Carga_Nucleo.gs) con dobles de Google");

/**
 * Página sin DOM: consulta, modelo, cola y bloqueos como en Carga_Pagina, con transporte JSON hacia cargaGuardar.
 * red: lista de fallas de transporte para los próximos envíos: 'caida' (no llega al servidor) o 'perdida' (el
 * servidor la ejecuta y la respuesta se pierde).
 */
function pagina(S, correo, opciones) {
  const mapa = D.sustituirIdentidad(S.ctx, [correo]);
  const credencial = Object.keys(mapa)[0];
  const red = [];
  const envios = [];
  const bloqueos = [];
  const enviar = pedido => new Promise((ok, mal) => setImmediate(() => {
    envios.push(json(pedido));
    const falla = red.shift();
    if (falla === "caida") return mal(new Error("sin conexión (simulado)"));
    const r = json(S.ctx.cargaGuardar(credencial, json(pedido)));
    if (falla === "perdida") return mal(new Error("respuesta perdida (simulado)"));
    ok(r);
  }));
  const P = { credencial, red, envios, bloqueos, modelo: null };
  P.cola = K.crearCola(Object.assign({
    enviar, esperar: () => Promise.resolve(), nuevoId: () => K.nuevoIdOperacion(), ahora: () => Date.now(),
    alRechazo: (clave, r, p) => {
      const alcance = K.alcanceRechazo(r.codigo);
      bloqueos.push([r.codigo, alcance]);
      const filtro = { todo: () => true, periodo: b => b.periodo === p.periodo, materia: b => b.orientacion === p.orientacion && b.materia === p.materia,
        estudiante: b => b.periodo === p.periodo && b.materia === p.materia && b.dni === p.dni, campo: b => K.claveCampo(b) === clave }[alcance];
      if (filtro) P.cola.bloquear(filtro, r.motivo);
    }
  }, opciones || {}));
  P.consultar = () => {
    const desde = Date.now();
    const r = json(S.ctx.cargaConsultar(credencial));
    if (r.estado !== "ok") return r;
    P.modelo = K.modelo(r);
    P.modelo.periodos.forEach(p => p.materias.forEach(m => m.estudiantes.forEach(e => campos.forEach(c => P.cola.registrar(e.claves[c], K.base(p, m, e, c), e.valores[c], desde)))));
    return r;
  };
  P.materia = (periodo, nombre) => { const p = P.modelo.periodos.find(x => x.periodo === periodo); return { p, m: p.materias.find(x => x.materia === nombre) }; };
  P.estudiante = (periodo, nombre, filtro) => { const { p, m } = P.materia(periodo, nombre); return { p, m, e: m.estudiantes.find(filtro) }; };
  P.campo = (e, c) => P.cola.campo(e.claves[c]);
  return P;
}
const grilla = (S, periodo, orientacion) => S.s.grillas[D.ID[periodo]][orientacion || "EPSI"];
const filasRegistro = S => S.s.admin.Registro.getDataRange().getDisplayValues().slice(1).filter(f => f.some(x => x !== ""));
const colRegistro = (S, nombre) => D.cargaDe(S.ctx).REGISTRO.indexOf(nombre);
const celda = (S, periodo, materia, dni, campo, orientacion) => {
  const v = grilla(S, periodo, orientacion).getDataRange().getDisplayValues();
  const e = S.ctx.estructuraGrilla_(v, periodo);
  const m = S.ctx.materiaPorClave_(e, S.ctx.clave_(materia));
  return v[S.ctx.filaEstudiante_(v, e, dni) + S.ctx.campoPorId_(campo).desplazamiento][m.columnaValor];
};
const pendiente = x => x.valores.calificacion === "CURSA" && !x.problema;

prueba("flujo normal: guardar Calificación en 2C escribe sólo la grilla de 2C y el Registro lo confirma; la misma celda de anual queda intacta y se guarda aparte", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  igual(P.campo(e, "calificacion").estado, "guardado");
  igual([celda(S, D.P2C, "Matemática II", e.dni, "calificacion"), celda(S, D.PAN, "Matemática II", e.dni, "calificacion")], ["7", "CURSA"]);
  igual([grilla(S, D.P2C).escrituras.length, grilla(S, D.PAN).escrituras.length], [1, 0]);
  const eAnual = P.estudiante(D.PAN, "Matemática II", x => x.dni === e.dni).e;
  igual(P.campo(eAnual, "calificacion").confirmado, "CURSA", "el modelo de anual no cambió");
  P.cola.cambiar(eAnual.claves.calificacion, "9");
  await drenar();
  igual([celda(S, D.P2C, "Matemática II", e.dni, "calificacion"), celda(S, D.PAN, "Matemática II", e.dni, "calificacion")], ["7", "9"]);
  const filas = filasRegistro(S);
  igual(filas.map(f => [f[colRegistro(S, "Estado")], f[colRegistro(S, "Período")], f[colRegistro(S, "Fuente")], f[colRegistro(S, "Anterior")], f[colRegistro(S, "Nuevo")]]),
    [["aplicado", D.P2C, D.ID[D.P2C], "CURSA", "7"], ["aplicado", D.PAN, D.ID[D.PAN], "CURSA", "9"]]);
});

prueba("los cuatro campos se guardan por separado y sin depender de la Calificación; «Dejar pendiente» vuelve a CURSA sin tocar los demás", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  P.cola.cambiar(e.claves.calidad, "Muy buena");
  P.cola.cambiar(e.claves.participacion, "Activa");
  P.cola.cambiar(e.claves.observaciones, "  Muy  buen\ntrabajo ");
  await drenar();
  igual(campos.map(c => celda(S, D.P2C, "Matemática II", e.dni, c)), ["CURSA", "Muy buena", "Activa", "Muy buen trabajo"]);
  igual(grilla(S, D.P2C).escrituras.length, 3);
  igual(P.campo(e, "observaciones").confirmado, "Muy buen trabajo");
  P.cola.cambiar(e.claves.calificacion, "8");
  await drenar();
  P.cola.cambiar(e.claves.calificacion, "CURSA");
  await drenar();
  igual(campos.map(c => celda(S, D.P2C, "Matemática II", e.dni, c)), ["CURSA", "Muy buena", "Activa", "Muy buen trabajo"]);
  igual(filasRegistro(S).map(f => f[colRegistro(S, "Nuevo")]), ["Muy buena", "Activa", "Muy buen trabajo", "8", "CURSA"]);
});

prueba("ACREDITÓ y Equivalencia: la Calificación no es editable, los otros tres sí; si una pantalla vieja intenta cambiarla, el servidor la rechaza sin escribir", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { p, m } = P.materia(D.P2C, "Matemática II");
  const admin = m.estudiantes.filter(x => K.situacion(p.escala, x.valores.calificacion) === "administrativo");
  igual(admin.map(x => x.valores.calificacion).sort(), ["ACREDITÓ", "Equivalencia"]);
  admin.forEach(x => igual([x.editable.calificacion, x.editable.calidad, x.editable.participacion, x.editable.observaciones], [false, true, true, true]));
  P.cola.cambiar(admin[0].claves.participacion, "Poca");
  P.cola.cambiar(admin[0].claves.calificacion, "CURSA");
  await drenar();
  igual(P.campo(admin[0], "participacion").estado, "guardado");
  igual([P.campo(admin[0], "calificacion").estado, P.campo(admin[0], "calificacion").codigo], ["rechazado", "calificacion"]);
  igual(celda(S, D.P2C, "Matemática II", admin[0].dni, "calificacion"), admin[0].valores.calificacion);
});

prueba("cursantes de otros grupos: aparecen en la materia y se les guarda igual", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", x => x.otroGrupo && pendiente(x));
  cierto(e && e.grupo !== "EPS2");
  P.cola.cambiar(e.claves.calificacion, "6");
  await drenar();
  igual([P.campo(e, "calificacion").estado, celda(S, D.P2C, "Matemática II", e.dni, "calificacion")], ["guardado", "6"]);
});

prueba("escritura que falla sin aplicarse: la página termina en «incierto» (nunca «guardado»), un único intento de escritura, una fila; volver a cargar es otra operación", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  grilla(S, D.P2C).fallas.push("fallar");
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  const c = P.campo(e, "calificacion");
  igual([c.estado, c.confirmado, c.propuesto], ["incierto", "CURSA", "7"]);
  igual(grilla(S, D.P2C).intentos, 1);
  igual(new Set(P.envios.map(x => JSON.stringify(x))).size, 1, "los reenvíos son la misma operación");
  igual(filasRegistro(S).map(f => f[colRegistro(S, "Estado")]), ["incierto"]);
  P.cola.usarPropuesto(e.claves.calificacion);
  await drenar();
  igual([P.campo(e, "calificacion").estado, celda(S, D.P2C, "Matemática II", e.dni, "calificacion")], ["guardado", "7"]);
  igual(filasRegistro(S).map(f => f[colRegistro(S, "Estado")]), ["incierto", "aplicado"]);
});

prueba("escritura aplicada que informa error: «coincide» (la grilla muestra el valor), nunca «guardado»; el Registro dice «inferido»", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  grilla(S, D.P2C).fallas.push("aplicar-y-fallar");
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  igual([P.campo(e, "calificacion").estado, filasRegistro(S)[0][colRegistro(S, "Estado")]], ["coincide", "inferido"]);
});

prueba("respuesta perdida: el servidor guardó, la página no se enteró; el reintento con la misma operación responde «guardado» sin escribir otra vez ni duplicar la fila", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  P.red.push("perdida");
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  igual([P.campo(e, "calificacion").estado, grilla(S, D.P2C).intentos, filasRegistro(S).length], ["guardado", 1, 1]);
  igual(P.envios.length, 2);
  igual(P.envios[0], P.envios[1]);
});

prueba("sin conexión hasta agotar los reintentos: «sin-confirmar»; el reintento manual, con la misma operación, guarda una sola vez", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  P.red.push("perdida", "caida", "caida", "caida");
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  igual(P.campo(e, "calificacion").estado, "sin-confirmar");
  P.cola.reintentar(e.claves.calificacion);
  await drenar();
  igual([P.campo(e, "calificacion").estado, grilla(S, D.P2C).intentos, filasRegistro(S).length], ["guardado", 1, 1]);
  igual(new Set(P.envios.map(x => JSON.stringify(x))).size, 1);
});

prueba("lock ocupado dos veces: la página reintenta sola con la misma operación y guarda; una fila", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  S.s.lock.ocupadas = 2;
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  igual([P.campo(e, "calificacion").estado, P.envios.length, filasRegistro(S).length], ["guardado", 3, 1]);
});

prueba("conflicto con el Equipo EPS: cambió la celda con la página abierta; la página muestra su valor, no escribe, y el docente decide volver a cargar", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  const h = grilla(S, D.P2C);
  const v = h.getDataRange().getDisplayValues(); const est = S.ctx.estructuraGrilla_(v, D.P2C);
  h.v[S.ctx.filaEstudiante_(v, est, e.dni)][S.ctx.materiaPorClave_(est, "MATEMATICAII").columnaValor] = "5";
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  let c = P.campo(e, "calificacion");
  igual([c.estado, c.confirmado, c.propuesto, h.intentos], ["conflicto", "5", "7", 0]);
  P.cola.usarPropuesto(e.claves.calificacion);
  await drenar();
  c = P.campo(e, "calificacion");
  igual([c.estado, celda(S, D.P2C, "Matemática II", e.dni, "calificacion"), filasRegistro(S)[0][colRegistro(S, "Anterior")]], ["guardado", "7", "5"]);
});

prueba("el Equipo EPS marca NO CURSA con la página abierta: el guardado se rechaza (alcance estudiante) y la situación pasa a «no cursa» al actualizar", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  const h = grilla(S, D.P2C);
  const v = h.getDataRange().getDisplayValues(); const est = S.ctx.estructuraGrilla_(v, D.P2C);
  h.v[S.ctx.filaEstudiante_(v, est, e.dni)][S.ctx.materiaPorClave_(est, "MATEMATICAII").columnaValor] = "NO CURSA";
  P.cola.cambiar(e.claves.calidad, "Buena");
  await drenar();
  igual([P.campo(e, "calidad").estado, P.bloqueos], ["rechazado", [["cursada", "estudiante"]]]);
  igual(P.cola.cambiar(e.claves.observaciones, "x"), false, "el estudiante quedó bloqueado");
  P.consultar();
  cierto(!P.materia(D.P2C, "Matemática II").m.estudiantes.some(x => x.dni === e.dni), "ya no figura");
  igual(P.cola.cambiar(e.claves.observaciones, "x"), false, "sigue bloqueado: ya no está en la carga");
});

prueba("la ventana cierra con la página abierta: el cierre que conoce la página es el del servidor; después, el guardado se rechaza sin escribir y el período queda bloqueado", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { p, e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  const cierre = S.ctx.configuracion_(S.ctx.datosCarga_({ getSheetByName: n => S.s.admin[n], getSpreadsheetTimeZone: () => D.ZONA },
    { getSheetByName: () => S.s.control }), new Date()).periodos[D.P2C].ventana.cierre.getTime();
  igual(p.cierre, cierre);
  S.reloj.desfase = p.cierre - Date.now() + 1000; // un segundo después del cierre (sólo en el servidor)
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  igual([P.campo(e, "calificacion").estado, P.campo(e, "calificacion").codigo, grilla(S, D.P2C).intentos], ["rechazado", "ventana", 0]);
  igual(P.bloqueos, [["ventana", "periodo"]]);
  const otro = P.estudiante(D.P2C, "Inglés I", pendiente).e;
  igual(P.cola.cambiar(otro.claves.calificacion, "7"), false, "todo el período quedó bloqueado");
  const anual = P.estudiante(D.PAN, "Matemática II", pendiente).e;
  cierto(P.cola.cambiar(anual.claves.calificacion, "7"), "anual tiene su propio bloqueo; aquí cerró a la misma hora");
  await drenar();
  igual(P.campo(anual, "calificacion").codigo, "ventana");
});

prueba("reintento después del cierre de una operación cuya respuesta se perdió: rechazado, e informa que el Registro la tenía «aplicado»", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo, { esperas: [1] });
  P.consultar();
  const { p, e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  P.red.push("perdida");
  // Antes del reintento, cierra la ventana.
  const enviarOriginal = S.ctx.cargaGuardar;
  let llamadas = 0;
  S.ctx.cargaGuardar = (c, x) => { const r = enviarOriginal(c, x); if (++llamadas === 1) S.reloj.desfase = p.cierre - Date.now() + 1000; return r; };
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  const c = P.campo(e, "calificacion");
  igual([c.estado, c.codigo, c.previo && c.previo.estado], ["rechazado", "ventana", "aplicado"]);
  igual(grilla(S, D.P2C).intentos, 1);
});

prueba("cambio de asignación con la página abierta: el docente anterior queda rechazado y bloqueado en esa materia (en ambos períodos); el nuevo la ve y continúa", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  S.s.admin.Asignaciones.v[1][3] = D.DOCENTES.reemplazo.nombre; S.s.admin.Asignaciones.v[1][4] = D.DOCENTES.reemplazo.correo;
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  igual([P.campo(e, "calificacion").codigo, P.bloqueos], ["asignacion", [["asignacion", "materia"]]]);
  const anual = P.estudiante(D.PAN, "Matemática II", x => x.dni === e.dni).e;
  igual(P.cola.cambiar(anual.claves.calificacion, "7"), false);
  cierto(P.cola.cambiar(P.estudiante(D.P2C, "Inglés I", pendiente).e.claves.calificacion, "7"), "las otras materias siguen");
  P.consultar();
  cierto(!P.modelo.periodos[0].materias.some(m => m.materia === "Matemática II"), "ya no figura para el anterior");
  const N = pagina(S, D.DOCENTES.reemplazo.correo);
  N.consultar();
  const nuevo = N.estudiante(D.P2C, "Matemática II", x => x.dni === e.dni).e;
  N.cola.cambiar(nuevo.claves.calificacion, "7");
  await drenar();
  igual(N.campo(nuevo, "calificacion").estado, "guardado");
});

prueba("el Equipo EPS cambia la grilla de 2C en Control: el guardado con la fuente anterior se rechaza (período bloqueado); al actualizar, las claves cambian y se guarda en la grilla nueva", async () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.matematica.correo);
  P.consultar();
  const { e } = P.estudiante(D.P2C, "Matemática II", pendiente);
  S.s.control.v[12][1] = D.URL(D.OTRA_2C);
  P.cola.cambiar(e.claves.calificacion, "7");
  await drenar();
  igual([P.campo(e, "calificacion").codigo, P.bloqueos], ["fuente", [["fuente", "periodo"]]]);
  P.consultar();
  const nuevo = P.estudiante(D.P2C, "Matemática II", x => x.dni === e.dni).e;
  cierto(nuevo.claves.calificacion !== e.claves.calificacion && nuevo.claves.calificacion.includes(D.OTRA_2C));
  igual(P.cola.cambiar(e.claves.calificacion, "7"), false, "la clave vieja sigue bloqueada");
  P.cola.cambiar(nuevo.claves.calificacion, "7");
  await drenar();
  igual([P.campo(nuevo, "calificacion").estado, S.s.grillas[D.OTRA_2C].EPSI.escrituras.length, grilla(S, D.P2C).escrituras.length], ["guardado", 1, 0]);
});

prueba("docente sin materias asignadas y período sin ventanas: la consulta lo dice sin datos", () => {
  const S = D.servidor();
  const P = pagina(S, D.DOCENTES.sinMaterias.correo);
  igual([P.consultar().estado, P.consultar().codigo], ["rechazado", "asignacion"]);
  S.s.admin.Ventanas.v = [S.s.admin.Ventanas.v[0]];
  const Q = pagina(S, D.DOCENTES.matematica.correo);
  igual([Q.consultar().estado, Q.modelo.periodos.length], ["ok", 0]);
});

// ================================================================ ejecución

(async () => {
  let fallas = 0, total = 0;
  for (const x of pruebas) {
    if (x.seccion) { seccionActual = x.seccion; console.log("\n# " + x.seccion); continue; }
    total++;
    try { await x.fn(); console.log("OK    " + x.nombre); }
    catch (e) { fallas++; console.log("FALLA " + x.nombre + "\n      " + String(e && e.stack || e).split("\n").slice(0, 6).join("\n      ")); }
  }
  console.log("\n" + (fallas ? fallas + " de " + total + " pruebas fallaron" : "Las " + total + " pruebas pasaron"));
  process.exit(fallas ? 1 : 0);
})();
