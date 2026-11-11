/**
 * Carga docente EPS (#29) · calendario de feriados para los días hábiles de CABA. Sin servicios de Google.
 *
 * Los días hábiles deciden el cierre de cada ventana (8:00 del primer día hábil posterior a «Hasta») y los días en
 * que salen los correos. Día hábil: lunes a viernes que no figura en «feriados» del año.
 *
 * Criterio institucional:
 *  - feriados: los feriados nacionales (inamovibles y trasladables, en su fecha efectiva) y los que rigen en CABA,
 *    también los que establece un decreto para un año. Uno que cae en fin de semana se incluye igual (no cambia nada);
 *  - noLaborables: los días no laborables (con fines turísticos, Jueves Santo) se anotan para que la decisión quede a
 *    la vista, pero NO cambian los días hábiles. Si el Equipo EPS decide que uno debe contar, se pasa a «feriados».
 *    Los no laborables sólo para quienes profesan una religión (Ley 27.399, arts. 2 y 3) no se anotan.
 *
 * Año sin entrada: no se adivina. Lo que necesita un día de ese año (el cierre de una ventana, el envío de correos)
 * queda bloqueado e informado; lo que no lo necesita sigue funcionando. No borrar años pasados mientras haya
 * ventanas de ese año en la hoja Ventanas.
 *
 * Para preparar un año nuevo: agregar su entrada con las fechas del calendario oficial y sus fuentes, guardar y
 * publicar una versión nueva de la app web (docs/puesta-en-marcha-carga-docente.md, F6 y F8).
 */
const CALENDARIO_FERIADOS = {
  // Fuentes: Ley 27.399 (arts. 1, 6 y 7), Decreto DNU 1103/2026 (BO 28/9/2026), Resolución JGM 164/2025 (BO 26/12/2025)
  // y el calendario de días inhábiles 2026 de argentina.gob.ar/tribunalfiscal/calendario. Revisado el 11/11/2026.
  2026: {
    feriados: {
      '2026-01-01': 'Año Nuevo',
      '2026-02-16': 'Carnaval',
      '2026-02-17': 'Carnaval',
      '2026-03-24': 'Día Nacional de la Memoria por la Verdad y la Justicia',
      '2026-04-02': 'Día del Veterano y de los Caídos en la Guerra de Malvinas (también Jueves Santo)',
      '2026-04-03': 'Viernes Santo',
      '2026-05-01': 'Día del Trabajador',
      '2026-05-25': 'Día de la Revolución de Mayo',
      '2026-06-15': 'Paso a la Inmortalidad del Gral. Martín Miguel de Güemes (17/6, trasladado)',
      '2026-06-20': 'Paso a la Inmortalidad del Gral. Manuel Belgrano',
      '2026-07-09': 'Día de la Independencia',
      '2026-08-17': 'Paso a la Inmortalidad del Gral. José de San Martín',
      '2026-10-12': 'Día del Respeto a la Diversidad Cultural',
      '2026-11-09': 'Visita del Papa León XIV (DNU 1103/2026, feriado nacional)',
      '2026-11-10': 'Visita del Papa León XIV (DNU 1103/2026, feriado en CABA y Córdoba)',
      '2026-11-23': 'Día de la Soberanía Nacional (20/11, trasladado)',
      '2026-12-08': 'Inmaculada Concepción de María',
      '2026-12-25': 'Navidad'
    },
    noLaborables: {
      '2026-03-23': 'Día no laborable con fines turísticos (Res. JGM 164/2025)',
      '2026-07-10': 'Día no laborable con fines turísticos (Res. JGM 164/2025)',
      '2026-12-07': 'Día no laborable con fines turísticos (Res. JGM 164/2025)'
    }
  }
};

/** Feriados de todos los años del calendario: { 'aaaa-mm-dd': true }. Un año sin ninguna fecha no tiene calendario. */
function feriadosCalendario_() {
  const feriados = {};
  Object.keys(CALENDARIO_FERIADOS).forEach(anio => Object.keys(CALENDARIO_FERIADOS[anio].feriados).forEach(f => { feriados[f] = true; }));
  return feriados;
}

/** Años con calendario, en orden. */
function aniosCalendario_() {
  return Object.keys(CALENDARIO_FERIADOS).sort();
}

/** Motivo del bloqueo cuando un día cae en un año sin calendario. */
function sinCalendario_(anio, consecuencia) {
  return 'Falta el calendario de feriados de ' + anio + ' en Carga_Calendario.gs: ' + consecuencia + '.';
}
