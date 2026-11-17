#!/usr/bin/env node
/**
 * Pruebas locales de los correos de #28 (Carga_Correo.gs y Carga_Correo_Google.gs) con todos los .gs reales sobre
 * los dobles de tests/carga-dobles.js: MailApp guarda los mensajes en memoria (no se envía nada), ScriptApp guarda
 * los activadores en memoria y el reloj del servidor se mueve a voluntad. Datos ficticios; sólo Node.js.
 *
 *   node tests/carga-correo.js [corridas]     corridas de la simulación de una ventana completa (10 por defecto)
 *
 * Calendario: el real de 2026 (Carga_Calendario.gs). 2º cuatrimestre y anual comparten la ventana del sábado 5/12/2026
 * al viernes 11/12/2026; el lunes 7/12 es día no laborable (hábil para la carga) y el martes 8/12, feriado; el cierre
 * es el lunes 14/12/2026 a las 8:00.
 */
"use strict";
const vm = require("vm");
const D = require("./carga-dobles.js");

const CORRIDAS = Number(process.argv[2]) || 10;
const pruebas = [];
function seccion(t) { pruebas.push({ seccion: t }); }
function prueba(nombre, fn) { pruebas.push({ nombre, fn }); }
function igual(a, b, msg) { const x = JSON.stringify(a), y = JSON.stringify(b); if (x !== y) throw new Error((msg || "distinto") + ": " + x + " ≠ " + y); }
function cierto(x, msg) { if (!x) throw new Error(msg || "se esperaba verdadero"); }
function lanza(fn, re) { let e = null; try { fn(); } catch (x) { e = x; } cierto(e && re.test(e.message), "se esperaba un error " + re + (e ? ": " + e.message : "")); }

const MATE = D.DOCENTES.matematica.correo, LENGUA = D.DOCENTES.lengua.correo, REEMPLAZO = D.DOCENTES.reemplazo.correo;
const P2C = D.P2C, PAN = D.PAN;
const VENTANAS = [["Período", "Desde", "Hasta"], [P2C, "2026-12-05", "2026-12-11"], [PAN, "2026-12-05", "2026-12-11"]];

/** Entorno con el calendario ficticio. */
function entorno(opciones) {
  const S = D.servidor(opciones);
  S.s.admin.Ventanas.v = VENTANAS.map(f => f.slice());
  S.en = local => { S.reloj.desfase = Date.parse(local + ":00-03:00") - Date.now(); };
  S.correr = () => S.ctx.enviarCorreosDia_(S.ctx.serviciosCorreo_(), {});
  S.mails = () => S.estado.correo.enviados;
  S.filas = () => S.ctx.correosLeer_(S.s.admin.Correos);
  return S;
}
const para = (S, correo) => S.mails().filter(m => m.to === correo);
const credencialDe = m => { const x = m.htmlBody.match(/\?k=([A-Za-z0-9_-]{43})/); return x ? x[1] : ""; };

/** Celdas de Calificación (y su bloque) de una materia en la grilla de un período. */
function celdasMateria(S, periodo, orientacion, materia) {
  const h = S.s.grillas[D.ID[periodo]][orientacion];
  const v = h.getDataRange().getDisplayValues();
  const e = S.ctx.estructuraGrilla_(v, periodo);
  const m = S.ctx.materiaPorClave_(e, S.ctx.clave_(materia));
  const filas = [];
  for (let r = 2; r < v.length; r++) if (S.ctx.dni_(v[r][2])) filas.push(r);
  return filas.map(r => ({
    fila: r, dni: S.ctx.dni_(v[r][2]), get valor() { return h.v[r][m.columnaValor]; },
    poner: (x, desplazamiento) => { h.v[r + (desplazamiento || 0)][m.columnaValor] = x; }
  }));
}
const MATERIAS = { [LENGUA]: [["EPSI", "Lengua y Literatura I"], ["EPSI", "Matemática III"]], [MATE]: [["EPSI", "Matemática II"], ["EPSI", "Inglés I"], ["EPSG", "Matemática I"]] };
/** Deja al día a un docente: sin CURSA y sin desaprobaciones (todo lo cursado, 8), en ambos períodos. */
function alDia(S, correo) {
  [P2C, PAN].forEach(p => MATERIAS[correo].forEach(([o, m]) => celdasMateria(S, p, o, m).forEach(c => {
    const s = S.ctx.situacionCalificacion_(p, c.valor);
    if (s === "pendiente" || (s === "valorada" && Number(c.valor) <= 5)) c.poner("8");
  })));
}
function ponerCursa(S, periodo, orientacion, materia) {
  const c = celdasMateria(S, periodo, orientacion, materia).find(x => S.ctx.situacionCalificacion_(periodo, x.valor) === "valorada");
  if (c) c.poner("CURSA");
  return c || null;
}

// ================================================================ 1. calendario

seccion("1. Días hábiles de CABA y horario");

prueba("diaHabil_ con el calendario del código: sábados, domingos y feriados no son hábiles (el 7/12, no laborable, sí); un año sin calendario no se adivina; primer día hábil desde un sábado es el lunes", () => {
  const S = entorno();
  const f = S.ctx.feriadosCalendario_();
  igual(["2026-12-04", "2026-12-05", "2026-12-06", "2026-12-07", "2026-12-08", "2026-12-09"].map(d => S.ctx.diaHabil_(d, f)), [true, false, false, true, false, true]);
  lanza(() => S.ctx.diaHabil_("2099-03-01", f), /Falta el calendario de feriados de 2099 en Carga_Calendario\.gs/);
  igual([S.ctx.primerDiaHabilDesde_("2026-12-05", f), S.ctx.primerDiaHabilDesde_("2026-12-08", f), S.ctx.primerDiaHabilDesde_("2026-12-24", f)], ["2026-12-07", "2026-12-09", "2026-12-24"]);
});

prueba("antes de las 8:00 no se envía nada, ni se lee la configuración; sábado, domingo y feriado tampoco, aunque la ventana esté abierta", () => {
  const S = entorno();
  S.en("2026-12-07T07:59");
  const lecturas = S.estado.aperturas.length;
  igual(S.correr().motivo, "Antes de las 8:00 no se envían correos.");
  igual(S.estado.aperturas.length, lecturas, "no abrió planillas");
  ["2026-12-05T08:30", "2026-12-06T08:30", "2026-12-08T08:30"].forEach(t => { S.en(t); cierto(/no es día hábil/.test(S.correr().motivo), t); });
  igual([S.mails().length, S.filas().length], [0, 0]);
});

// ================================================================ 2. apertura, seguimiento y cierre

seccion("2. Apertura, seguimiento y cierre (un correo por docente y día, aunque haya dos períodos)");

prueba("apertura el primer día hábil (lunes 7/12, la ventana abrió el sábado): un único correo por docente con 2C y anual, su enlace y el estado inicial; el enlace funciona y es del destinatario", () => {
  const S = entorno();
  S.en("2026-12-07T08:12");
  const r = S.correr();
  igual([r.enviados.length, r.errores.length, r.problemas], [2, 0, []]);
  [MATE, LENGUA].forEach(c => {
    const m = para(S, c);
    igual(m.length, 1, c);
    cierto(/abrió 2º cuatrimestre y Anual/.test(m[0].subject), m[0].subject);
    cierto(m[0].htmlBody.includes("2º cuatrimestre · abrió la carga") && m[0].htmlBody.includes("Anual · abrió la carga"));
    cierto(m[0].body.includes("Estado de las grillas al lunes 7/12/2026, 8:12") && m[0].body.includes("El plazo para completar la carga es el viernes 11/12. Como margen, el sistema seguirá abierto hasta el lunes 14/12 a las 8:00; después ya no se podrán hacer cambios."), m[0].body.slice(0, 400));
    igual(m[0].name, "Equipo EPS");
    cierto(m[0].htmlBody.startsWith('<div') && m[0].htmlBody.includes(D.URL_APP + "?k="));
    const v = S.ctx.cargaConsultar(credencialDe(m[0]));
    igual([v.estado, v.correo], ["ok", c], "el enlace del correo abre la carga de su destinatario");
  });
  igual(S.filas().map(f => [f.dia, f.correo, f.estado]), [["2026-12-07", MATE, "enviado"], ["2026-12-07", LENGUA, "enviado"]]);
  igual(S.filas()[0].eventos, ["abierta|2º cuatrimestre|2026-12-05", "abierta|Anual|2026-12-05"]);
});

prueba("otra ejecución el mismo día (manual o un activador repetido) no manda un segundo correo, aunque cambien los pendientes", () => {
  const S = entorno();
  S.en("2026-12-07T08:12"); S.correr();
  ponerCursa(S, P2C, "EPSI", "Lengua y Literatura I");
  S.en("2026-12-07T15:40");
  const r = S.correr();
  igual([r.enviados.length, r.yaRecibieron.sort()], [0, [LENGUA, MATE].sort()]);
  igual(S.mails().length, 2);
});

prueba("seguimiento: sólo a quien tiene pendientes; se suspende al quedar al día y se reanuda si aparecen nuevos", () => {
  const S = entorno();
  S.en("2026-12-07T08:05"); S.correr();
  alDia(S, LENGUA);
  S.en("2026-12-09T08:05");
  let r = S.correr();
  igual(r.enviados.map(x => [x.correo, x.tipos]), [[MATE, ["seguimiento"]]]);
  cierto(/^Carga de valoraciones EPS: plazo viernes 11\/12 · \d+ calificaciones pendientes$/.test(para(S, MATE)[1].subject), para(S, MATE)[1].subject);
  ponerCursa(S, PAN, "EPSI", "Matemática III");
  S.en("2026-12-10T08:05");
  r = S.correr();
  igual(r.enviados.map(x => x.correo).sort(), [LENGUA, MATE].sort());
  const l = para(S, LENGUA)[1];
  cierto(/1 calificación pendiente/.test(l.subject), l.subject);
  cierto(l.htmlBody.includes("<b>Pendiente</b>"));
  // Al día otra vez: sin correo.
  alDia(S, LENGUA);
  S.en("2026-12-11T08:05");
  igual(S.correr().enviados.map(x => x.correo), [MATE]);
});

prueba("cierre: el lunes 14/12 después de las 8:00, a todos los asignados (también a quien estaba al día), con el resumen final por período, sin enlace; antes de las 8:00, nada", () => {
  const S = entorno();
  S.en("2026-12-07T08:05"); S.correr();
  alDia(S, LENGUA);
  S.en("2026-12-14T07:59");
  igual(S.correr().enviados.length, 0);
  S.en("2026-12-14T08:20");
  const r = S.correr();
  igual(r.enviados.map(x => [x.correo, x.tipos]).sort(), [[LENGUA, ["cierre"]], [MATE, ["cierre"]]].sort());
  const m = para(S, LENGUA)[1];
  cierto(/cerró 2º cuatrimestre y Anual, resumen final/.test(m.subject), m.subject);
  cierto(m.htmlBody.includes("Cerró el lunes 14/12 a las 8:00.") && /no certifica que la carga esté completa/.test(m.body));
  igual(credencialDe(m), "", "sin enlace después del cierre");
  // Al día siguiente, nada más.
  S.en("2026-12-15T08:20");
  igual(S.correr().enviados.length, 0);
});

prueba("cierre atrasado: si el trabajo no corrió el lunes, el cierre sale el siguiente día hábil en que corre, hasta 7 días después; más tarde, no", () => {
  const S = entorno();
  S.en("2026-12-07T08:05"); S.correr();
  S.en("2026-12-15T08:05");
  igual(S.correr().enviados.map(x => x.tipos[0]), ["cierre", "cierre"]);
  const T = entorno();
  T.en("2026-12-22T08:05");
  igual(T.correr().enviados.length, 0);
  T.en("2026-12-21T08:05");
  igual(T.correr().enviados.length, 2, "7 días después todavía sale");
});

prueba("apertura atrasada: si el trabajo no corrió el primer día hábil, el siguiente envía la apertura (no un seguimiento) a todos", () => {
  const S = entorno();
  alDia(S, LENGUA);
  S.en("2026-12-09T08:05");
  igual(S.correr().enviados.map(x => [x.correo, x.tipos]).sort(), [[LENGUA, ["apertura"]], [MATE, ["apertura"]]].sort());
});

prueba("ventanas distintas el mismo día (cierra 2C, sigue anual): un solo correo con el cierre de 2C, el estado de anual y su enlace", () => {
  const S = entorno();
  S.s.admin.Ventanas.v = [VENTANAS[0], [P2C, "2026-12-05", "2026-12-09"], [PAN, "2026-12-05", "2026-12-18"]];
  S.en("2026-12-07T08:05"); S.correr();
  S.en("2026-12-10T08:30");
  const r = S.correr();
  igual(r.enviados.map(x => x.correo).sort(), [LENGUA, MATE].sort());
  const m = para(S, MATE)[1];
  cierto(m.htmlBody.includes("2º cuatrimestre · resumen final") && m.htmlBody.includes("Anual · carga abierta"), m.subject);
  cierto(credencialDe(m), "el enlace sigue, por anual");
  igual(S.mails().filter(x => x.to === MATE).length, 2);
});

prueba("#33 plazo y margen, con el calendario de noviembre de 2026 (Hasta viernes 20/11; sábado, domingo y el feriado del lunes 23/11 de margen; cierre martes 24/11 a las 8:00): el asunto lleva el plazo; el cuerpo lo explica una vez, con el margen y su hora; sin correos en el margen", () => {
  const S = entorno();
  S.s.admin.Ventanas.v = [VENTANAS[0], [P2C, "2026-11-16", "2026-11-20"], [PAN, "2026-11-16", "2026-11-20"]];
  const frase = "El plazo para completar la carga es el viernes 20/11. Como margen, el sistema seguirá abierto hasta el martes 24/11 a las 8:00; después ya no se podrán hacer cambios.";
  S.en("2026-11-16T08:35");
  igual(S.correr().enviados.length, 2);
  const m = para(S, MATE)[0];
  cierto(/^Carga de valoraciones EPS: abrió 2º cuatrimestre y Anual · plazo viernes 20\/11 · \d+ calificaciones pendientes$/.test(m.subject), m.subject);
  cierto(!/24\/11/.test(m.subject), "el asunto no lleva el cierre del sistema");
  igual([m.body.split(frase).length - 1, m.htmlBody.split(frase).length - 1], [1, 1], "una sola vez en texto y en HTML, aunque haya dos períodos");
  cierto(m.body.indexOf(frase) < m.body.indexOf("Abrir mi carga"), "arriba, antes del enlace");
  cierto(m.htmlBody.includes('<p style="margin:0 0 8px">' + frase + '</p>'), "un párrafo propio en el HTML");
  [m.subject, m.body, m.htmlBody].forEach(x => cierto(!/[Cc]ierra el/.test(x), "ninguna parte presenta el día del cierre como un día de carga: " + x.slice(0, 200)));
  // Seguimiento y último día del plazo.
  S.en("2026-11-17T08:35");
  S.correr();
  igual(/^Carga de valoraciones EPS: plazo viernes 20\/11 · \d+ calificaciones pendientes$/.test(para(S, MATE).slice(-1)[0].subject), true, para(S, MATE).slice(-1)[0].subject);
  S.en("2026-11-20T08:10");
  S.correr();
  const u = para(S, MATE).slice(-1)[0];
  cierto(/^Carga de valoraciones EPS: plazo hoy, viernes 20\/11 · \d+ calificaciones pendientes$/.test(u.subject), u.subject);
  cierto(u.body.includes("El plazo para completar la carga vence hoy, viernes 20/11. Como margen, el sistema seguirá abierto hasta el martes 24/11 a las 8:00;"), u.body.slice(0, 500));
  // El margen no tiene días hábiles: no hay correos entre el plazo y el cierre.
  const antes = S.mails().length;
  ["2026-11-21T08:30", "2026-11-22T08:30", "2026-11-23T08:30"].forEach(t => { S.en(t); cierto(/no es día hábil/.test(S.correr().motivo), t); });
  S.en("2026-11-24T07:59");
  S.correr();
  igual(S.mails().length, antes);
  // Cierre: dice cuándo cerró, sin volver sobre el plazo.
  S.en("2026-11-24T08:20");
  S.correr();
  const c = para(S, MATE).slice(-1)[0];
  igual(c.subject, "Carga de valoraciones EPS: cerró 2º cuatrimestre y Anual, resumen final");
  igual(c.body.split("Cerró el martes 24/11 a las 8:00.").length - 1, 2, "una vez por período cerrado");
  cierto(!/plazo/i.test(c.body), "sin plazo en el cierre");
  // Un plazo ya vencido no debería llegar a un correo (no hay días hábiles en el margen), pero se redacta bien.
  const g = { hasta: "2026-11-20", cierre: new Date("2026-11-24T08:00:00-03:00"), periodos: [P2C] };
  igual([S.ctx.plazoAsunto_(g, "2026-11-21", false), S.ctx.plazoAsunto_(g, "2026-11-19", true)], ["venció el plazo (viernes 20/11)", "plazo viernes 20/11 (2º cuatrimestre)"]);
  igual(S.ctx.plazoFrase_(g, "2026-11-21", true), "El plazo para completar la carga de 2º cuatrimestre venció el viernes 20/11. Como margen, el sistema seguirá abierto hasta el martes 24/11 a las 8:00; después ya no se podrán hacer cambios.");
});

prueba("#33 períodos con plazos distintos (2C hasta el viernes 11/12, anual hasta el martes 15/12, que cierra el miércoles 16/12 a las 8:00, sin fin de semana de por medio): un plazo por período, nombrado; el asunto, el más próximo; con un período cerrado en el mismo correo, el abierto se nombra", () => {
  const S = entorno();
  S.s.admin.Ventanas.v = [VENTANAS[0], [P2C, "2026-12-05", "2026-12-11"], [PAN, "2026-12-05", "2026-12-15"]];
  S.en("2026-12-07T08:05");
  S.correr();
  const m = para(S, MATE)[0];
  cierto(/^Carga de valoraciones EPS: abrió 2º cuatrimestre y Anual · plazo viernes 11\/12 \(2º cuatrimestre\) · \d+ calificaciones pendientes$/.test(m.subject), m.subject);
  cierto(m.body.includes("El plazo para completar la carga de 2º cuatrimestre es el viernes 11/12. Como margen, el sistema seguirá abierto hasta el lunes 14/12 a las 8:00;"), m.body.slice(0, 700));
  cierto(m.body.includes("El plazo para completar la carga de Anual es el martes 15/12. Como margen, el sistema seguirá abierto hasta el miércoles 16/12 a las 8:00;"), m.body.slice(0, 700));
  cierto(/== 2º cuatrimestre · abrió la carga ==\n\d+ calificaciones pendientes\.\n/.test(m.body) && /== Anual · abrió la carga ==\n\d+ calificaciones pendientes\.\n/.test(m.body), "con dos períodos abiertos, los pendientes de cada uno");
  // 2C cerró el lunes 14/12 (ese día no corrió: se recupera el martes 15/12, último día del plazo de anual).
  S.en("2026-12-15T08:05");
  S.correr();
  const u = para(S, MATE).slice(-1)[0];
  cierto(/^Carga de valoraciones EPS: plazo hoy, martes 15\/12 \(Anual\) · \d+ calificaciones pendientes · cerró 2º cuatrimestre, resumen final$/.test(u.subject), u.subject);
  cierto(u.body.includes("El plazo para completar la carga de Anual vence hoy, martes 15/12. Como margen, el sistema seguirá abierto hasta el miércoles 16/12 a las 8:00;"), u.body.slice(0, 600));
  cierto(u.body.includes("== 2º cuatrimestre · resumen final ==\nCerró el lunes 14/12 a las 8:00.\n"), "la sección cerrada dice cuándo cerró");
  cierto(/== Anual · carga abierta ==\n\n/.test(u.body), "con un solo período abierto, sus pendientes van sólo arriba");
});

// ================================================================ 3. asignaciones y enlaces

seccion("3. Asignaciones vigentes y enlaces");

prueba("docente asignado a mitad de la ventana: recibe la apertura el siguiente día hábil, sólo con su materia; el anterior deja de verla", () => {
  const S = entorno();
  S.en("2026-12-07T08:05"); S.correr();
  S.s.admin.Asignaciones.v[1][3] = D.DOCENTES.reemplazo.nombre; S.s.admin.Asignaciones.v[1][4] = REEMPLAZO; // Matemática II
  S.en("2026-12-09T08:05");
  const r = S.correr();
  igual(r.enviados.find(x => x.correo === REEMPLAZO).tipos, ["apertura"]);
  const nuevo = para(S, REEMPLAZO)[0], anterior = para(S, MATE)[1];
  cierto(nuevo.htmlBody.includes("EPSI EPS2 · Matemática II") && !nuevo.htmlBody.includes("Inglés I"));
  cierto(!anterior.htmlBody.includes("EPSI EPS2 · Matemática II") && anterior.htmlBody.includes("Inglés I"));
  igual(S.ctx.cargaConsultar(credencialDe(nuevo)).correo, REEMPLAZO);
});

prueba("revocación: el docente recibe su nuevo enlace el siguiente día hábil aunque esté al día; el viejo deja de valer; sin revocar, no hay correo", () => {
  const S = entorno();
  S.en("2026-12-07T08:05"); S.correr();
  const viejo = credencialDe(para(S, LENGUA)[0]);
  alDia(S, LENGUA);
  S.s.admin.Enlaces.v.push([LENGUA, 2]);
  S.en("2026-12-09T08:05");
  const r = S.correr();
  igual(r.enviados.find(x => x.correo === LENGUA).tipos, ["enlace"]);
  const m = para(S, LENGUA)[1];
  cierto(/tu nuevo enlace/.test(m.subject) && /El anterior ya no funciona/.test(m.body));
  const nuevo = credencialDe(m);
  cierto(nuevo && nuevo !== viejo);
  igual([S.ctx.cargaConsultar(viejo).codigo, S.ctx.cargaConsultar(nuevo).estado], ["acceso", "ok"]);
  S.en("2026-12-10T08:05");
  igual(S.correr().enviados.map(x => x.correo), [MATE]);
});

prueba("una materia con asignación ambigua o que no figura en la grilla aparece con su problema, sin estudiantes; versión ilegible en Enlaces: no se le envía e informa", () => {
  const S = entorno();
  S.s.admin.Asignaciones.v.push(["EPSI", "EPS3", "Matemática III", D.DOCENTES.reemplazo.nombre, REEMPLAZO]);
  S.s.admin.Asignaciones.v.push(["EPSG", "EPS1", "Materia inexistente", D.DOCENTES.matematica.nombre, MATE]);
  S.s.admin.Enlaces.v.push([REEMPLAZO, "dos"]);
  S.en("2026-12-07T08:05");
  const r = S.correr();
  const l = para(S, LENGUA)[0].htmlBody;
  cierto(l.includes("asignación ambigua") && para(S, MATE)[0].htmlBody.includes("no figura en la grilla"));
  igual(para(S, REEMPLAZO).length, 0);
  cierto(r.problemas.some(p => p.startsWith(REEMPLAZO) && /versión/.test(p)), r.problemas.join(" | "));
});

// ================================================================ 4. contenido y privacidad

seccion("4. Contenido: datos guardados, sin DNI ni contenido de los campos complementarios, sólo las asignaciones propias");

prueba("cada correo lista a todos los cursantes de sus materias con su Calificación, como la grilla; sin DNI, sin Calidad/Participación/Observaciones, sin materias ajenas", () => {
  const S = entorno();
  // Marcas únicas en Observaciones de todas las grillas; Calidad y Participación ya tienen sus opciones.
  let n = 0;
  Object.values(S.s.grillas).flatMap(g => Object.values(g)).forEach(h => h.v.forEach((f, r) => f.forEach((x, c) => {
    if (r > 0 && /^Observaci/.test(String(h.v[r][c - 1] || "")) && !h.v[r][c]) h.v[r][c] = "OBS-SECRETA-" + (++n);
  })));
  S.en("2026-12-07T08:05"); S.correr();
  const dnis = Object.values(S.s.grillas).flatMap(g => Object.values(g)).flatMap(h => h.v.map(f => S.ctx.dni_(f[2])).filter(Boolean));
  const opciones = D.cargaDe(S.ctx).OPCIONES;
  S.mails().forEach(m => {
    const todo = m.subject + m.body + m.htmlBody;
    dnis.forEach(d => cierto(!todo.includes(d), "DNI " + d + " en el correo a " + m.to));
    cierto(!/OBS-SECRETA|Observación ficticia/.test(todo), "contenido de Observaciones");
    opciones.calidad.concat(opciones.participacion).forEach(o => cierto(!new RegExp("\\b" + o + "\\b").test(todo), "opción «" + o + "» en el correo"));
  });
  const mate = para(S, MATE)[0].htmlBody;
  cierto(!mate.includes("Lengua y Literatura I") && !mate.includes("EPSI EPS3 · Matemática III"), "materias de otro docente");
  // Lo listado coincide con la consulta (la misma lectura de las grillas).
  const v = S.ctx.consultarCarga_(MATE);
  v.periodos.forEach(p => p.materias.forEach(m => {
    m.estudiantes.forEach(e => cierto(mate.includes(S.ctx.esc_(e.nombre)), e.nombre));
    const pendientes = m.conteo.pendiente ? m.conteo.pendiente + (m.conteo.pendiente === 1 ? " calificación pendiente" : " calificaciones pendientes") : "sin calificaciones pendientes";
    cierto(mate.includes("<li style=\"margin:0 0 4px\">" + S.ctx.esc_(m.orientacion + " " + m.nivel + " · " + m.materia) + ": <span style=\"" + (m.conteo.pendiente ? "font-weight:bold" : "color:#1e6b34") + "\">" + pendientes + "</span>"), m.materia);
  }));
});

prueba("los nombres se escapan en el HTML; el texto plano trae lo mismo que el HTML", () => {
  const S = entorno();
  const c = celdasMateria(S, P2C, "EPSI", "Matemática II")[0];
  S.s.grillas[D.ID[P2C]].EPSI.v[c.fila][0] = '<img src=x onerror="alert(1)">';
  S.en("2026-12-07T08:05"); S.correr();
  const m = para(S, MATE)[0];
  cierto(!m.htmlBody.includes("<img") && m.htmlBody.includes("&lt;img src=x onerror=&quot;alert(1)&quot;&gt;"));
  cierto(m.body.includes('<img src=x onerror="alert(1)">'), "texto plano sin escapar (no es HTML)");
});

prueba("mejora no bloqueante: una desaprobación (≤5, o «En proceso» en los bimestres) con Calidad, Participación u Observaciones vacías pide completarlas, sin reproducirlas, y mantiene el seguimiento", () => {
  const S = entorno();
  S.en("2026-12-07T08:05"); S.correr();
  alDia(S, LENGUA);
  const c = celdasMateria(S, P2C, "EPSI", "Lengua y Literatura I").find(x => S.ctx.situacionCalificacion_(P2C, x.valor) === "valorada");
  c.poner("4"); c.poner("Buena", 1); c.poner("Activa", 2); c.poner("", 3);
  S.en("2026-12-09T08:05");
  const r = S.correr();
  igual(r.enviados.find(x => x.correo === LENGUA).tipos, ["seguimiento"]);
  const m = para(S, LENGUA)[1];
  cierto(/1 desaprobación por completar/.test(m.subject) && m.htmlBody.includes("4 · completar Calidad, Participación u Observaciones"), m.subject);
  cierto(/- EPSI EPS1 · Lengua y Literatura I: [^\n]*1 desaprobación por completar/.test(m.body), "el resumen de la materia lo dice");
  cierto(!/\bBuena\b|\bActiva\b/.test(m.htmlBody + m.body));
  c.poner("Observación completa", 3);
  S.en("2026-12-10T08:05");
  igual(S.correr().enviados.map(x => x.correo), [MATE]);
  const d = S.ctx.desaprobada_;
  igual([d("3º bimestre", { situacion: "valorada", campos: { calificacion: "En proceso" } }), d("3º bimestre", { situacion: "valorada", campos: { calificacion: "Suficiente" } }),
    d("Anual", { situacion: "valorada", campos: { calificacion: "6" } }), d("Anual", { situacion: "valorada", campos: { calificacion: "5" } })], [true, false, false, true]);
});

prueba("#35 resumen legible en un teléfono: un renglón por materia con lo que pide algo al docente (pendientes, valores no reconocidos), sin tabla de conteos ni categorías internas; ACREDITÓ y EQUIVALENCIA dicen quién las pone; un valor no reconocido no se declara inválido y se revisa con el Equipo EPS; en un período cerrado, «quedaron pendientes»", () => {
  const S = entorno();
  S.en("2026-12-07T08:05"); S.correr();
  const m = para(S, MATE)[0];
  [m.htmlBody, m.body].forEach(x => cierto(!/valoradas|a revisar|del Equipo EPS,|válid/i.test(x), "categorías internas o invalidez: " + (x.match(/valoradas|a revisar|del Equipo EPS,|válid/i) || [])[0]));
  cierto(!/<th\b/.test(m.htmlBody), "sin encabezados de tabla");
  (m.htmlBody.match(/<tr>.*?<\/tr>/g) || []).forEach(f => cierto((f.match(/<td\b/g) || []).length <= 2, "a lo sumo dos columnas: " + f.slice(0, 120)));
  cierto(/: ACREDITÓ \(la establece el Equipo EPS\)\n/.test(m.body) && /: Equivalencia \(la establece el Equipo EPS\)\n/.test(m.body), "quién pone ACREDITÓ y EQUIVALENCIA");
  cierto(m.body.includes(": Muy bien (la carga no reconoce este valor: revisalo con el Equipo EPS)\n"), "qué hacer con un valor no reconocido");
  const v = S.ctx.consultarCarga_(MATE);
  let noValidas = 0;
  v.periodos.forEach(p => p.materias.forEach(x => {
    noValidas += x.conteo["a revisar"];
    const partes = [x.conteo.pendiente + (x.conteo.pendiente === 1 ? " calificación pendiente" : " calificaciones pendientes")];
    if (x.conteo["a revisar"]) partes.push(x.conteo["a revisar"] === 1 ? "1 calificación con un valor no reconocido (revisalo con el Equipo EPS)" : x.conteo["a revisar"] + " calificaciones con valores no reconocidos (revisalos con el Equipo EPS)");
    cierto(m.body.includes("\n- " + x.orientacion + " " + x.nivel + " · " + x.materia + ": " + partes.join(" · ") + "\n"), x.materia + ": " + partes.join(" · "));
  }));
  cierto(noValidas > 1 && m.body.includes("\n" + noValidas + " calificaciones tienen en la grilla un valor que la carga no reconoce, así que no podés cambiarlas: revisalas con el Equipo EPS.\n"), "aviso arriba");
  // Dos períodos abiertos y nada pendiente (enlace nuevo a un docente al día): «Sin calificaciones pendientes.», no «0 …».
  alDia(S, LENGUA);
  S.s.admin.Enlaces.v.push([LENGUA, 2]);
  S.en("2026-12-09T08:05"); S.correr();
  const l = para(S, LENGUA)[1];
  igual([l.body.split("\nSin calificaciones pendientes.\n").length - 1, /\b0 calificaciones/.test(l.body)], [2, false], l.body.slice(0, 900));
  cierto(/: sin calificaciones pendientes/.test(l.body));
  // Cierre: lo pendiente ya no se puede cargar.
  S.en("2026-12-14T08:20"); S.correr();
  const c = para(S, MATE).slice(-1)[0];
  cierto(/: \d+ calificaciones quedaron pendientes/.test(c.body) && !/: \d+ calificaciones pendientes/.test(c.body), c.body.slice(0, 900));
});

prueba("#35 período cerrado: lo que quedó se informa como estado, sin pedir lo que ya no se puede hacer (ni «por completar» ni «completar …»); revisar un valor con el Equipo EPS sigue; con un período abierto en el mismo correo, lo que todavía se puede hacer dice de qué período es", () => {
  const S = entorno();
  S.s.admin.Ventanas.v = [VENTANAS[0], [P2C, "2026-12-05", "2026-12-09"], [PAN, "2026-12-05", "2026-12-18"]];
  S.en("2026-12-07T08:05"); S.correr();
  alDia(S, LENGUA);
  const c = celdasMateria(S, P2C, "EPSI", "Lengua y Literatura I").find(x => S.ctx.situacionCalificacion_(P2C, x.valor) === "valorada");
  c.poner("4"); c.poner("", 1); c.poner("", 2); c.poner("", 3);
  // Mixto: 2C cerró el jueves 10/12 a las 8:00; anual sigue.
  S.en("2026-12-10T08:30"); S.correr();
  const m = para(S, LENGUA).slice(-1)[0];
  const [abierta, cerrada] = m.body.split("== 2º cuatrimestre · resumen final ==");
  cierto(cerrada && /: \d+ · completar Calidad/.test(m.body) === false, "la fila del cerrado no pide completar");
  cierto(/Lengua y Literatura I: [^\n]*1 desaprobación quedó con Calidad, Participación u Observaciones sin completar/.test(cerrada) && !/por completar/.test(cerrada), cerrada.slice(0, 500));
  cierto(/: 4 · Calidad, Participación u Observaciones sin completar\n/.test(cerrada), "estado de la fila");
  cierto(/\(la carga no reconoce este valor: revisalo con el Equipo EPS\)/.test(cerrada), "revisar con el Equipo EPS sigue siendo posible");
  cierto(abierta.includes("\nNo tenés calificaciones pendientes en Anual.\n") && abierta.includes("\n1 calificación de Anual tiene en la grilla un valor que la carga no reconoce, así que no podés cambiarla: revisala con el Equipo EPS.\n"), abierta.slice(0, 700));
  cierto(!/desaprobaci/.test(m.subject) && !/desaprobaci/.test(abierta), "la desaprobación incompleta es de 2C, ya cerrado: no se pide arriba ni en el asunto");
  // Sólo cierres: lo mismo en HTML, en gris y sin «por completar».
  S.en("2026-12-21T08:20"); S.correr();
  const f = para(S, LENGUA).slice(-1)[0];
  cierto(/cerró Anual, resumen final$/.test(f.subject), f.subject);
  igual(S.ctx.estadoMateria_({ conteo: { pendiente: 0, "a revisar": 0 }, completar: 2 }, true), [{ texto: "sin calificaciones pendientes", tipo: "listo" },
    { texto: "2 desaprobaciones quedaron con Calidad, Participación u Observaciones sin completar", tipo: "registro" }]);
  igual(S.ctx.calificacionTexto_({ situacion: "valorada", calificacion: "3", completar: true }, true), "3 · Calidad, Participación u Observaciones sin completar");
  igual(S.ctx.calificacionTexto_({ situacion: "valorada", calificacion: "3", completar: true }, false), "3 · completar Calidad, Participación u Observaciones");
  // Resumen final recortado por tamaño: sin enlace, el pie no remite a la carga.
  const T = entorno();
  T.en("2026-12-07T08:05"); T.correr();
  vm.runInContext("CORREO.MAX_HTML = 3000", T.ctx);
  T.en("2026-12-14T08:20"); T.correr();
  const g = para(T, MATE).slice(-1)[0];
  cierto(credencialDe(g) === "" && g.body.includes("El detalle completo no entra en un correo: este sólo lista lo que quedó pendiente o sin completar y los valores a revisar.") &&
    !/está en la carga/.test(g.body), g.body.slice(-500));
});

prueba("tamaño: un correo de un docente con muchas materias queda lejos del recorte de Gmail; si superara el límite, el detalle sólo lista lo que requiere atención", () => {
  const S = entorno({ porNivel: 30 });
  S.en("2026-12-07T08:05");
  const r = S.correr();
  const max = Math.max(...r.enviados.map(x => x.bytes));
  cierto(max < 90000, "HTML de " + max + " bytes");
  const T = entorno();
  vm.runInContext("CORREO.MAX_HTML = 3000", T.ctx);
  T.en("2026-12-07T08:05"); T.correr();
  const m = para(T, MATE)[0];
  cierto(/El detalle completo no entra en un correo/.test(m.body));
  cierto(m.htmlBody.includes("<b>Pendiente</b>"));
  const valorada = T.ctx.consultarCarga_(MATE).periodos[0].materias[0].estudiantes.find(e => e.situacion === "valorada" && Number(e.campos.calificacion) > 5);
  cierto(!m.htmlBody.includes(T.ctx.esc_(valorada.nombre)) || /\(EPS/.test(m.htmlBody), "en el resumido no se listan las valoradas");
  vm.runInContext("CORREO.MAX_HTML = 90000", T.ctx);
  console.log("      tamaño máximo con 30 estudiantes por nivel: " + max + " bytes de HTML");
});

// ================================================================ 5. fallas

seccion("5. Fallas: envío con error, ejecución cortada, cuota, tiempo y configuración");

prueba("un envío con error queda «error» y otra ejecución del mismo día lo reintenta en la misma fila; nunca hay dos «enviado» por docente y día", () => {
  const S = entorno();
  S.estado.correo.fallas = ["fallar"];
  S.en("2026-12-07T08:05");
  const r = S.correr();
  igual([r.enviados.length, r.errores.length], [1, 1]);
  igual(S.filas().map(f => f.estado).sort(), ["enviado", "error"]);
  S.en("2026-12-07T10:00");
  igual(S.correr().enviados.length, 1);
  igual([S.filas().length, S.filas().map(f => f.estado)], [2, ["enviado", "enviado"]]);
  igual(S.mails().map(m => m.to).sort(), [LENGUA, MATE].sort());
});

prueba("un error de envío no reintentado ese día: al siguiente día hábil, la apertura sale como apertura", () => {
  const S = entorno();
  S.estado.correo.fallas = ["fallar"];
  S.en("2026-12-07T08:05"); S.correr();
  const fallado = S.filas().find(f => f.estado === "error").correo;
  S.en("2026-12-09T08:05");
  igual(S.correr().enviados.find(x => x.correo === fallado).tipos, ["apertura"]);
});

prueba("ejecución cortada después de enviar (no pudo marcar «enviado»): la fila queda «pendiente», cuenta como enviada y no se repite ese día ni como apertura", () => {
  const S = entorno();
  const h = S.s.admin.Correos, getRange = h.getRange.bind(h);
  let cortar = 1;
  h.getRange = (f, c, nf, nc) => {
    const r = getRange(f, c, nf, nc);
    if (c === 3 && nc === 2 && cortar-- > 0) r.setValues = () => { throw new Error("se cortó la ejecución (simulado)"); };
    return r;
  };
  S.en("2026-12-07T08:05");
  const r = S.correr();
  igual([S.mails().length, r.problemas.length], [2, 1]);
  igual(S.filas().map(f => f.estado).sort(), ["enviado", "pendiente"]);
  S.en("2026-12-07T11:00");
  igual(S.correr().enviados.length, 0);
  S.en("2026-12-09T08:05");
  cierto(S.correr().enviados.every(x => !x.tipos.includes("apertura")));
});

prueba("cuota de MailApp agotada o tiempo agotado: lo que falta queda «sin procesar» (sin fila) y otra ejecución del día lo completa", () => {
  const S = entorno();
  S.estado.correo.cuota = 1;
  S.en("2026-12-07T08:05");
  let r = S.correr();
  igual([r.enviados.length, r.sinProcesar.length, S.filas().length], [1, 1, 1]);
  S.estado.correo.cuota = 100;
  S.en("2026-12-07T09:30");
  r = S.correr();
  igual([r.enviados.length, S.filas().length, S.mails().length], [1, 2, 2]);
  const T = entorno();
  T.en("2026-12-07T08:05");
  const s = T.ctx.serviciosCorreo_();
  s.ms = () => 5 * 60 * 1000;
  r = T.ctx.enviarCorreosDia_(s, {});
  igual([r.enviados.length, r.sinProcesar.length, T.mails().length], [0, 2, 0]);
});

prueba("configuración: año sin calendario o problema global, ningún correo; un período con problema (sin grilla en Control) no se informa y se avisa", () => {
  const S = entorno();
  const restituir = D.sinAnioDeCalendario(S.ctx, "2026");
  S.en("2026-12-07T08:05");
  let r = S.correr();
  restituir();
  igual([r.enviados.length], [0]);
  cierto(/Falta el calendario de feriados de 2026 en Carga_Calendario\.gs/.test(r.problemas.join()), r.problemas.join());
  const asignaciones = S.s.admin.Asignaciones.v;
  S.s.admin.Asignaciones.v = [asignaciones[0].concat(["Correo"])].concat(asignaciones.slice(1));
  r = S.correr();
  igual([r.enviados.length], [0]);
  cierto(/^Configuración: La hoja Asignaciones debe tener una única columna/.test(r.problemas.join()), r.problemas.join());
  S.s.admin.Asignaciones.v = asignaciones;
  S.s.control.v[13][1] = ""; // sin grilla de anual
  r = S.correr();
  igual(r.enviados.length, 2);
  cierto(/^Anual: no se envían sus correos/.test(r.problemas[0]), r.problemas.join());
  cierto(!para(S, MATE)[0].htmlBody.includes("Anual"));
});

prueba("#29: sin ventanas en curso ni cierres por informar, un año sin calendario no bloquea: la ejecución termina sin problemas; con algo que informar, sí", () => {
  const S = entorno();
  const restituir = D.sinAnioDeCalendario(S.ctx, "2026");
  try {
    S.en("2026-11-20T08:30"); // la ventana del 5/12 todavía no empezó
    let r = S.correr();
    igual([r.motivo, r.problemas, r.enviados.length], ["No hay ventanas en curso ni cierres para informar.", [], 0]);
    S.en("2026-12-07T08:30");
    r = S.correr();
    igual(r.enviados.length, 0);
    cierto(/Falta el calendario de feriados de 2026/.test(r.problemas.join()), r.problemas.join());
  } finally { restituir(); }
});

prueba("sin URL_CARGA válida no se envían correos con enlace (se informa); un cierre, que no lleva enlace, sale igual", () => {
  const S = entorno();
  S.estado.props.URL_CARGA = "https://ejemplo.invalid/app";
  S.en("2026-12-07T08:05");
  let r = S.correr();
  igual([r.enviados.length, r.problemas.length], [0, 2]);
  S.en("2026-12-14T08:05");
  r = S.correr();
  igual(r.enviados.map(x => x.tipos), [["cierre"], ["cierre"]]);
});

// ================================================================ 6. activador y funciones de la propietaria

seccion("6. Activador diario y funciones de la propietaria (MailApp y ScriptApp simulados)");

prueba("activarCorreosCarga: un único activador diario a las 8 en hora de Buenos Aires (idempotente); exige secreto, URL y hoja Correos; desactivar lo borra", () => {
  const S = entorno();
  S.estado.props.URL_CARGA = "";
  lanza(() => S.ctx.activarCorreosCarga(), /URL_CARGA/);
  igual(S.estado.activadores.length, 0);
  S.estado.props.URL_CARGA = D.URL_APP;
  S.ctx.activarCorreosCarga(); S.ctx.activarCorreosCarga();
  igual(S.estado.activadores.map(t => [t.funcion, t.config]), [["enviarCorreosCarga", { hora: 8, cadaDias: 1, zona: "America/Argentina/Buenos_Aires" }]]);
  cierto(S.consola.some(l => /entre las 8:00 y las 9:00/.test(l)));
  S.ctx.desactivarCorreosCarga();
  igual(S.estado.activadores.length, 0);
  delete S.s.admin.Correos;
  lanza(() => S.ctx.activarCorreosCarga(), /Falta la hoja Correos/);
});

prueba("enviarCorreosCarga: corre con el evento del activador real (aunque quien ejecuta no sea la propietaria); con otra ejecución en curso no hace nada; con problemas termina con error para que Google avise", () => {
  const S = entorno();
  S.ctx.activarCorreosCarga();
  const uid = S.estado.activadores[0].uid;
  S.estado.usuario = "";
  S.en("2026-12-07T08:40");
  const r = S.ctx.enviarCorreosCarga({ triggerUid: uid, hour: 8 });
  igual([r.enviados, r.problemas], [2, 0]);
  S.estado.userLock.tomado = true;
  S.en("2026-12-09T08:40");
  igual(S.ctx.enviarCorreosCarga({ triggerUid: uid }), null);
  igual(S.mails().length, 2);
  S.estado.userLock.tomado = false;
  S.estado.correo.fallas = ["fallar"];
  lanza(() => S.ctx.enviarCorreosCarga({ triggerUid: uid }), /1 con error/);
  igual(S.estado.userLock.tomado, false, "liberó el lock");
  cierto(S.consola.every(l => !/\?k=/.test(l)), "el registro de ejecución no lleva enlaces");
});

prueba("previsualizarCorreosCarga: arma los mismos mensajes sin anotar ni enviar nada, y sin escribir enlaces en el registro", () => {
  const S = entorno();
  S.en("2026-12-07T08:05");
  const r = S.ctx.previsualizarCorreosCarga();
  igual([r.simulados, r.enviados, S.mails().length, S.filas().length], [2, 0, 0, 0]);
  cierto(S.consola.some(l => /"correo":"simulado"/.test(l)) && S.consola.every(l => !/\?k=/.test(l)));
  S.estado.usuario = "visitante@ficticio.example";
  lanza(() => S.ctx.previsualizarCorreosCarga(), /propietario/);
});

// ================================================================ 7. simulación de una ventana completa

seccion("7. Simulación: una ventana completa, día por día, con cambios del Equipo EPS, fallas de envío y ejecuciones repetidas (" + CORRIDAS + " corridas)");

function prng(s) { return () => { s |= 0; s = s + 0x6D2B79F5 | 0; let t = Math.imul(s ^ s >>> 15, 1 | s); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }

prueba("invariantes: ≤1 correo por docente y día; sólo días hábiles desde las 8:00; apertura y cierre a todos; seguimiento si y sólo si hay pendientes o desaprobaciones por completar; lo informado coincide con la grilla al enviar; sin DNI", () => {
  const stats = { correos: 0, apertura: 0, seguimiento: 0, cierre: 0, enlace: 0, errores: 0, repetidas: 0 };
  for (let corrida = 0; corrida < CORRIDAS; corrida++) {
    const azar = prng(1000 + corrida);
    const S = entorno();
    const dnis = Object.values(S.s.grillas).flatMap(g => Object.values(g)).flatMap(h => h.v.map(f => S.ctx.dni_(f[2])).filter(Boolean));
    const dias = ["2026-12-04", "2026-12-05", "2026-12-06", "2026-12-07", "2026-12-08", "2026-12-09", "2026-12-10", "2026-12-11", "2026-12-12", "2026-12-13", "2026-12-14", "2026-12-15", "2026-12-16"];
    const recibidoCierre = {}, recibidoApertura = {};
    dias.forEach(dia => {
      // Cambios del día: docentes que cargan, el Equipo EPS que agrega cursadas o revoca un enlace.
      if (azar() < 0.4) alDia(S, azar() < 0.5 ? LENGUA : MATE);
      if (azar() < 0.3) ponerCursa(S, azar() < 0.5 ? P2C : PAN, "EPSI", azar() < 0.5 ? "Matemática III" : "Matemática II");
      if (azar() < 0.1) S.s.admin.Enlaces.v.push([azar() < 0.5 ? LENGUA : MATE, 2 + Math.floor(azar() * 5)]);
      if (azar() < 0.2) S.estado.correo.fallas.push("fallar");
      const ejecuciones = 1 + (azar() < 0.3 ? 1 : 0);
      for (let i = 0; i < ejecuciones; i++) {
        const minuto = Math.floor(azar() * 60);
        S.en(dia + "T" + (i ? "1" + Math.floor(azar() * 9) : "08") + ":" + String(minuto).padStart(2, "0"));
        const antes = S.mails().length;
        const r = S.correr();
        stats.repetidas += i;
        stats.errores += r.errores.length;
        const ahora = new Date();
        const habil = S.ctx.diaHabil_(dia, { "2026-12-08": true, "2026-12-25": true });
        if (!habil) igual(S.mails().length, antes, "envió en día no hábil " + dia);
        // Lo que recién se envió refleja la grilla de este momento.
        r.enviados.forEach(e => {
          stats.correos++; e.tipos.forEach(t => { stats[t]++; });
          const m = S.mails().filter(x => x.to === e.correo).slice(-1)[0];
          dnis.forEach(d => cierto(!m.htmlBody.includes(d) && !m.body.includes(d), "DNI"));
          const v = S.ctx.consultarCarga_(e.correo);
          if (v.periodos.length) {
            const pend = v.periodos.reduce((n, p) => n + p.materias.reduce((k, x) => k + (x.conteo ? x.conteo.pendiente : 0), 0), 0);
            igual(e.pendientes, pend, "pendientes informados vs grilla, " + e.correo + " " + dia);
          }
          if (e.tipos.includes("apertura")) recibidoApertura[e.correo] = true;
          if (e.tipos.includes("cierre")) recibidoCierre[e.correo] = true;
        });
        // Quien no recibió nada hoy (y ya tuvo apertura) no tiene pendientes ni nada por completar.
        if (habil && ahora.getTime() < Date.parse("2026-12-14T08:00:00-03:00")) {
          [MATE, LENGUA].forEach(c => {
            if (S.filas().some(f => f.correo === c && f.dia === dia)) return;
            const v = S.ctx.consultarCarga_(c);
            const atencion = v.periodos.some(p => p.materias.some(x => x.conteo && (x.conteo.pendiente ||
              x.estudiantes.some(est => S.ctx.complementariosFaltan_(p.periodo, est)))));
            cierto(!atencion || !recibidoApertura[c], "sin correo con pendientes: " + c + " " + dia);
          });
        }
      }
      // ≤1 por docente y día.
      const porDia = {};
      S.filas().filter(f => f.estado !== "error").forEach(f => { const k = f.dia + f.correo; porDia[k] = (porDia[k] || 0) + 1; });
      cierto(Object.values(porDia).every(n => n === 1), "más de un correo por docente y día");
      igual(S.mails().filter(m => m.to === MATE && m.al.slice(0, 10) === new Date(Date.parse(dia + "T12:00:00-03:00")).toISOString().slice(0, 10)).length <= 1, true);
    });
    // Sin errores sin reintentar, todos tuvieron apertura y cierre.
    S.filas().forEach(f => cierto(f.dia >= "2026-12-07" && f.dia !== "2026-12-08" && !["2026-12-12", "2026-12-13"].includes(f.dia), "fila de " + f.dia));
    [MATE, LENGUA].forEach(c => { cierto(recibidoApertura[c], "apertura de " + c); cierto(recibidoCierre[c], "cierre de " + c); });
  }
  console.log("      " + JSON.stringify(stats));
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
