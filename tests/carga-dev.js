#!/usr/bin/env node
/**
 * Carga docente EPS (#27, #28) · servidor local de desarrollo de la interfaz, con datos ficticios.
 *
 *   node tests/carga-dev.js [puerto]            (8787 por defecto; sólo escucha en 127.0.0.1)
 *
 * Sirve la página real (src/carga/Carga_Pagina.html con Carga_Cliente.html) y atiende sus llamadas con los tres .gs
 * reales del servidor sobre dobles de Google (tests/carga-dobles.js). Nada sale de la máquina: no usa recursos
 * Google ni datos reales.
 *
 * Identidad: los enlaces de los docentes ficticios son los reales de #28, calculados por el código del servidor con
 * un secreto ficticio (D.CLAVE_FICTICIA); nada del servidor se sustituye. Una credencial inventada o revocada se
 * rechaza como en producción. En /dev/correos se ven los correos que se enviarían ahora (simulados: no sale nada).
 *
 * Panel de simulación en /dev: fallas en el próximo guardado (escritura ambigua, lock ocupado, respuesta perdida, sin
 * conexión, lentitud), cambios del Equipo EPS con la página abierta (celdas, asignaciones, grilla en Control,
 * cierre de la ventana) y el Registro.
 */
"use strict";
const http = require("http");
const D = require("./carga-dobles.js");

const PUERTO = Number(process.argv[2]) || 8787;
const S = D.servidor({ porNivel: 7 });
const DOCENTES = [D.DOCENTES.matematica, D.DOCENTES.lengua, D.DOCENTES.reemplazo, D.DOCENTES.sinMaterias];
// Se calculan en cada pedido: cambian si el Equipo EPS revoca un enlace o reasigna una materia.
const credencialDe = correo => Object.keys(D.enlacesReales(S.ctx, [correo]))[0] || "";
const red = { proximas: [], lentoMs: 0 }; // fallas de transporte para los próximos guardados: 'caida' | 'perdida'
const historial = [];

function registrarAccion(t) { historial.unshift(new Date().toLocaleTimeString("es-AR") + " · " + t); historial.length = Math.min(historial.length, 30); }

// ---------------------------------------------------------------- página con google.script.run simulado

const SIMULADOR = `
<script>
/* Sólo en el servidor local: google.script.run sobre fetch('/rpc'). En Apps Script lo provee HtmlService. */
(function () {
  function corredor(ok, mal) {
    return new Proxy({}, { get: function (_, nombre) {
      if (nombre === 'withSuccessHandler') return function (f) { return corredor(f, mal); };
      if (nombre === 'withFailureHandler') return function (f) { return corredor(ok, f); };
      return function () {
        var args = Array.prototype.slice.call(arguments);
        fetch('/rpc', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fn: nombre, args: args }) })
          .then(function (r) { return r.json(); })
          .then(function (x) { if (x.error) { if (mal) mal(new Error(x.error)); } else if (ok) ok(x.resultado); })
          .catch(function (e) { if (mal) mal(e); });
      };
    } });
  }
  window.google = { script: { run: corredor(null, null) } };
  document.addEventListener('DOMContentLoaded', function () {
    var d = document.createElement('a');
    d.href = '/dev'; d.target = '_blank'; d.rel = 'noopener';
    d.textContent = 'Servidor local · datos ficticios · panel';
    d.setAttribute('style', 'position:fixed;left:8px;bottom:8px;z-index:99;background:#3b1d5e;color:#fff;font:12px system-ui;padding:5px 9px;border-radius:6px;opacity:.9;text-decoration:none');
    document.body.appendChild(d);
  });
})();
</script>`;

// ---------------------------------------------------------------- acciones del panel

function grillas() { return Object.values(S.s.grillas).flatMap(g => Object.values(g)); }
function proximaFallaEscritura(tipo) {
  const compartida = [tipo]; // la primera escritura en cualquier grilla la consume
  grillas().forEach(h => { h.fallas = compartida; });
}
function cierreVigente() {
  const cfg = S.ctx.configuracion_(S.ctx.datosCarga_({ getSheetByName: n => S.s.admin[n], getSpreadsheetTimeZone: () => D.ZONA }, { getSheetByName: () => S.s.control }), new Date(Date.now() + S.reloj.desfase));
  return cfg.periodos[D.P2C].ventana && cfg.periodos[D.P2C].ventana.cierre ? cfg.periodos[D.P2C].ventana.cierre.getTime() : null;
}
function editarCelda(periodo, orientacion, materia, dni, campo, valor) {
  const id = S.s.control.v[{ [D.P2C]: 12, [D.PAN]: 13 }[periodo]][1];
  const fuente = S.ctx.idPlanilla_(id);
  const h = S.s.grillas[fuente] && S.s.grillas[fuente][orientacion];
  if (!h) throw new Error("No hay grilla para " + periodo + " " + orientacion);
  const v = h.getDataRange().getDisplayValues();
  const est = S.ctx.estructuraGrilla_(v, periodo);
  const m = S.ctx.materiaPorClave_(est, S.ctx.clave_(materia));
  const fila = S.ctx.filaEstudiante_(v, est, dni) + S.ctx.campoPorId_(campo).desplazamiento;
  const antes = v[fila][m.columnaValor];
  h.v[fila][m.columnaValor] = valor;
  return antes;
}

const ACCIONES = {
  "escritura-falla": () => { proximaFallaEscritura("fallar"); return "La próxima escritura en una grilla falla sin aplicarse (resultado incierto)."; },
  "escritura-aplica-y-falla": () => { proximaFallaEscritura("aplicar-y-fallar"); return "La próxima escritura se aplica pero informa error (coincide / inferido)."; },
  "lock-2": () => { S.s.lock.ocupadas = 2; return "El lock está ocupado para los próximos 2 intentos (la página reintenta sola)."; },
  "lock-6": () => { S.s.lock.ocupadas = 6; return "El lock está ocupado para los próximos 6 intentos (se agotan los reintentos automáticos)."; },
  "respuesta-perdida": () => { red.proximas.push("perdida"); return "El próximo guardado se ejecuta, pero la respuesta se pierde."; },
  "sin-conexion": () => { red.proximas.push("caida", "caida", "caida", "caida"); return "Los próximos 4 envíos no llegan al servidor."; },
  "lento": () => { red.lentoMs = 5000; return "Las próximas respuestas tardan 5 s (hasta restablecer)."; },
  "red-normal": () => { red.proximas = []; red.lentoMs = 0; grillas().forEach(h => { h.fallas = []; }); S.s.lock.ocupadas = 0; return "Sin fallas pendientes."; },
  "celda": q => {
    const antes = editarCelda(q.periodo, q.orientacion, q.materia, q.dni, q.campo, q.valor);
    return "Equipo EPS: " + q.periodo + " · " + q.orientacion + " · " + q.materia + " · DNI " + q.dni + " · " + q.campo + ": «" + antes + "» → «" + q.valor + "».";
  },
  "reasignar": () => {
    const a = S.s.admin.Asignaciones.v[1];
    const vuelve = a[4] === D.DOCENTES.reemplazo.correo;
    const d = vuelve ? D.DOCENTES.matematica : D.DOCENTES.reemplazo;
    a[3] = d.nombre; a[4] = d.correo;
    return "Asignaciones: EPSI · Matemática II ahora es de " + d.nombre + ".";
  },
  "fuente": () => {
    const actual = S.ctx.idPlanilla_(S.s.control.v[12][1]);
    const nueva = actual === D.OTRA_2C ? D.ID[D.P2C] : D.OTRA_2C;
    S.s.control.v[12][1] = D.URL(nueva);
    return "Control!B13 (2º cuatrimestre) apunta ahora a " + nueva + ".";
  },
  "cierre-90s": q => {
    const segundos = Math.max(1, Number(q && q.segundos) || 90);
    S.reloj.desfase = 0;
    const cierre = cierreVigente();
    if (!cierre) throw new Error("La ventana de 2C no tiene cierre.");
    S.reloj.desfase = cierre - segundos * 1000 - Date.now();
    return "Reloj del servidor adelantado: faltan " + segundos + " s para el cierre de 2º cuatrimestre y anual (8:00 del día de cierre). Actualizá la página para que lo tome.";
  },
  "cerrar-ya": () => {
    const v = S.s.admin.Ventanas.v;
    const pasado = D.isoBA(Date.now() - 20 * 86400000), antes = D.isoBA(Date.now() - 25 * 86400000);
    v.forEach((f, i) => { if (i && (f[0] === D.P2C || f[0] === D.PAN)) { f[1] = antes; f[2] = pasado; } });
    return "Ventanas: 2º cuatrimestre y anual cerraron (fechas en el pasado). La página abierta se entera en el próximo guardado.";
  },
  "revocar": q => {
    const correo = q && q.correo === D.DOCENTES.lengua.correo ? D.DOCENTES.lengua.correo : D.DOCENTES.matematica.correo;
    const v = S.s.admin.Enlaces.v;
    const fila = v.find((f, i) => i && f[0] === correo);
    if (fila) fila[1] = Number(fila[1]) + 1; else v.push([correo, 2]);
    return "Enlaces: se revocó el enlace de " + correo + " (versión " + (fila ? fila[1] : 2) + "). El anterior deja de valer en la próxima llamada.";
  },
  "reiniciar": () => { S.reiniciar(); S.reloj.desfase = 0; red.proximas = []; red.lentoMs = 0; return "Escenario reiniciado: grillas, administración y Registro como al empezar."; }
};

// ---------------------------------------------------------------- panel

const esc = D.escHtml;
function panel(base) {
  const reg = S.s.admin.Registro.getDataRange().getDisplayValues();
  const enc = reg[0];
  const filas = reg.slice(1).filter(f => f.some(x => x !== "")).reverse();
  const columnas = ["Fecha", "Estado", "Detalle", "Intentos", "Operación", "Correo", "Período", "Orientación", "Materia", "DNI", "Campo", "Anterior", "Nuevo"].map(c => enc.indexOf(c));
  const materias = ["Matemática II", "Inglés I", "Matemática I", "Lengua y Literatura I", "Matemática III"];
  const boton = (a, t) => `<button name="accion" value="${a}">${esc(t)}</button>`;
  const desfase = S.reloj.desfase ? `<p class="nota">Reloj del servidor corrido ${Math.round(S.reloj.desfase / 60000)} min.</p>` : "";
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Panel local · carga docente</title>
<style>body{font:14px/1.4 system-ui,sans-serif;margin:0;padding:16px;background:#f5f3f8;color:#1c1c28}h1{font-size:18px}h2{font-size:15px;margin:18px 0 6px}
section{background:#fff;border:1px solid #ddd;border-radius:8px;padding:10px 12px;margin-bottom:12px}button{margin:3px;padding:5px 10px}
table{border-collapse:collapse;font-size:12px;width:100%}td,th{border-bottom:1px solid #eee;padding:3px 5px;text-align:left;vertical-align:top}
.nota{color:#666}.aviso{background:#fff3dc;border:1px solid #f0c77a;padding:6px 10px;border-radius:6px}form.celda label{margin-right:6px}.esc{overflow:auto}</style></head><body>
<h1>Panel local de la carga docente (#27) · datos ficticios</h1>
<p class="nota">Enlaces reales de #28, calculados por el servidor con un secreto ficticio. <a href="/dev/correos" target="_blank">Correos que se enviarían ahora</a> (simulados).</p>
${historial.length ? `<p class="aviso">${esc(historial[0])}</p>` : ""}${desfase}
<section><h2>Enlaces de docentes ficticios</h2><ul>${DOCENTES.map(d => `<li><a href="${base}/?k=${encodeURIComponent(credencialDe(d.correo))}" target="_blank">${esc(d.nombre)}</a> <span class="nota">${esc(d.correo)}</span></li>`).join("")}
<li><a href="${base}/?k=credencial-inventada" target="_blank">Credencial inventada</a> <span class="nota">(debe rechazarse)</span></li></ul>
<form method="post" action="/dev/accion"><input type="hidden" name="accion" value="revocar"><select name="correo">${DOCENTES.slice(0, 2).map(d => `<option value="${esc(d.correo)}">${esc(d.nombre)}</option>`).join("")}</select> <button>Revocar su enlace (Enlaces)</button></form></section>
<form method="post" action="/dev/accion">
<section><h2>Próximo guardado</h2>${boton("escritura-falla", "Escritura falla sin aplicarse")}${boton("escritura-aplica-y-falla", "Escritura aplicada con error")}${boton("lock-2", "Lock ocupado ×2")}${boton("lock-6", "Lock ocupado ×6")}${boton("respuesta-perdida", "Respuesta perdida")}${boton("sin-conexion", "Sin conexión ×4")}${boton("lento", "Respuestas lentas (5 s)")}${boton("red-normal", "Quitar fallas")}</section>
<section><h2>Equipo EPS con la página abierta</h2>${boton("reasignar", "Reasignar / devolver Matemática II (EPSI)")}${boton("fuente", "Cambiar / restaurar la grilla de 2C en Control")}${boton("cierre-90s", "Cierre de 2C y anual en 90 s")}${boton("cerrar-ya", "Cerrar 2C y anual ya (Ventanas)")}${boton("reiniciar", "Reiniciar escenario")}</section>
</form>
<section><h2>El Equipo EPS edita una celda</h2><form class="celda" method="post" action="/dev/accion"><input type="hidden" name="accion" value="celda">
<label>Período <select name="periodo"><option>${D.P2C}</option><option>${D.PAN}</option></select></label>
<label>Orientación <select name="orientacion"><option>EPSI</option><option>EPSG</option></select></label>
<label>Materia <select name="materia">${materias.map(m => `<option>${esc(m)}</option>`).join("")}</select></label>
<label>DNI <input name="dni" size="9" placeholder="90000007"></label>
<label>Campo <select name="campo"><option value="calificacion">Calificación</option><option value="calidad">Calidad</option><option value="participacion">Participación</option><option value="observaciones">Observaciones</option></select></label>
<label>Valor <input name="valor" size="12" placeholder="5, NO CURSA, ACREDITÓ…"></label><button>Escribir en la grilla</button></form>
<p class="nota">Los DNI ficticios empiezan en 90000000 (EPSI) y 92000000 (EPSG).</p></section>
<section><h2>Registro (más reciente primero)</h2><div class="esc"><table><tr>${columnas.map(i => `<th>${esc(enc[i])}</th>`).join("")}</tr>
${filas.map(f => `<tr>${columnas.map(i => `<td>${esc(f[i])}</td>`).join("")}</tr>`).join("") || `<tr><td colspan="13" class="nota">Sin filas</td></tr>`}</table></div></section>
<section><h2>Acciones</h2><ul>${historial.map(h => `<li>${esc(h)}</li>`).join("")}</ul></section>
</body></html>`;
}

// ---------------------------------------------------------------- servidor HTTP

function cuerpo(req) {
  return new Promise((ok, mal) => {
    let datos = "";
    req.on("data", c => { datos += c; if (datos.length > 1e6) { mal(new Error("pedido demasiado grande")); req.destroy(); } });
    req.on("end", () => ok(datos));
  });
}
const enviar = (res, codigo, tipo, contenido) => { res.writeHead(codigo, { "Content-Type": tipo, "Cache-Control": "no-store", "Referrer-Policy": "no-referrer" }); res.end(contenido); };

http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://127.0.0.1");
  try {
    if (req.method === "GET" && url.pathname === "/") {
      // doGet real con un HtmlService simulado; luego, el google.script.run local.
      const html = S.ctx.doGet({ parameter: Object.fromEntries(url.searchParams) }).contenido.replace("</head>", SIMULADOR + "\n</head>");
      return enviar(res, 200, "text/html; charset=utf-8", html);
    }
    if (req.method === "POST" && url.pathname === "/rpc") {
      const { fn, args } = JSON.parse(await cuerpo(req));
      if (fn !== "cargaConsultar" && fn !== "cargaGuardar") return enviar(res, 200, "application/json", JSON.stringify({ error: "Función no disponible: " + fn }));
      const falla = fn === "cargaGuardar" ? red.proximas.shift() : undefined;
      if (red.lentoMs) await new Promise(r => setTimeout(r, red.lentoMs));
      if (falla === "caida") return enviar(res, 200, "application/json", JSON.stringify({ error: "Sin conexión (simulado)" }));
      const resultado = S.ctx[fn].apply(null, Array.isArray(args) ? args : []);
      if (falla === "perdida") return enviar(res, 200, "application/json", JSON.stringify({ error: "Respuesta perdida (simulado; el servidor sí lo procesó)" }));
      return enviar(res, 200, "application/json", JSON.stringify({ resultado }));
    }
    if (url.pathname === "/favicon.ico") { res.writeHead(204); return res.end(); }
    if (req.method === "GET" && url.pathname === "/dev") return enviar(res, 200, "text/html; charset=utf-8", panel(""));
    if (req.method === "POST" && url.pathname === "/dev/accion") {
      const tipo = req.headers["content-type"] || "";
      const datos = await cuerpo(req);
      const q = tipo.includes("json") ? JSON.parse(datos) : Object.fromEntries(new URLSearchParams(datos));
      let texto;
      try { texto = ACCIONES[q.accion] ? ACCIONES[q.accion](q) : "Acción desconocida."; } catch (e) { texto = "Error: " + e.message; }
      registrarAccion(texto);
      if (tipo.includes("json")) return enviar(res, 200, "application/json", JSON.stringify({ ok: !/^Error/.test(texto), texto }));
      res.writeHead(303, { Location: "/dev" }); return res.end();
    }
    if (req.method === "GET" && url.pathname === "/dev/correos") {
      // Simulación (sin anotar ni enviar): los enlaces apuntan a este servidor local.
      const r = S.ctx.enviarCorreosDia_(S.ctx.serviciosCorreo_(), { simular: true });
      const html = `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>Correos simulados</title><style>body{font:14px system-ui;margin:16px;background:#f5f3f8}
section{background:#fff;border:1px solid #ddd;border-radius:8px;padding:10px;margin-bottom:14px}iframe{width:100%;height:520px;border:1px solid #eee}</style></head><body>
<h1 style="font-size:18px">Correos que se enviarían ahora (simulados, día ${esc(r.dia)})</h1>${r.motivo ? `<p>${esc(r.motivo)}</p>` : ""}${r.problemas.map(p => `<p style="color:#b3261e">${esc(p)}</p>`).join("")}
${r.mensajes.map(m => `<section><b>Para:</b> ${esc(m.correo)} · <b>Asunto:</b> ${esc(m.asunto)} · ${m.bytes} bytes<iframe sandbox="allow-popups allow-popups-to-escape-sandbox" srcdoc="${esc(m.html.split(D.URL_APP).join("http://localhost:" + PUERTO + "/"))}"></iframe></section>`).join("")}</body></html>`;
      return enviar(res, 200, "text/html; charset=utf-8", html);
    }
    if (req.method === "GET" && url.pathname === "/dev/credenciales") {
      return enviar(res, 200, "application/json", JSON.stringify(Object.fromEntries(DOCENTES.map(d => [d.correo, credencialDe(d.correo)]))));
    }
    if (req.method === "GET" && url.pathname === "/dev/registro") {
      return enviar(res, 200, "application/json", JSON.stringify(S.s.admin.Registro.getDataRange().getDisplayValues()));
    }
    enviar(res, 404, "text/plain; charset=utf-8", "No encontrado");
  } catch (e) {
    enviar(res, 500, "text/plain; charset=utf-8", "Error: " + e.message);
  }
}).listen(PUERTO, "127.0.0.1", () => {
  const base = "http://localhost:" + PUERTO;
  console.log("Carga docente (#27, #28) · servidor local con datos ficticios");
  console.log("Panel de simulación: " + base + "/dev\n");
  DOCENTES.forEach(d => console.log((d.nombre + ":").padEnd(34) + base + "/?k=" + encodeURIComponent(credencialDe(d.correo))));
  console.log("\nCorreos simulados: " + base + "/dev/correos\nCtrl+C para terminar.");
});
