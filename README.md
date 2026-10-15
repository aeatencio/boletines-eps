# boletines-eps

Generador de boletines para la Escuela Profesional Secundaria (EPS) del CFP N.º 7, donde se desarrolla y se usa.

El equipo docente carga las valoraciones en grillas de Google Sheets. `boletines-eps` arma a partir de ellas el boletín de cada estudiante con el formato institucional, listo para revisar, exportar e imprimir.

```
Grillas de valoración → Generador (Google Sheets + Apps Script) → Boletín → Revisión, exportación e impresión
```

El proyecto está en desarrollo activo. La versión actual es v0.4, y el trabajo pendiente se organiza en los issues del repositorio.

## Puesta en funcionamiento

1. Subir `src/Generador_EPS_v0_4.xlsx` a Google Drive y abrirlo como Google Sheets.
2. En *Extensiones > Apps Script*, pegar `src/Generador_EPS_v0_4.gs`, guardar y recargar.
3. Completar la hoja `Control` con el período, el estudiante y las grillas de origen.
4. Ejecutar *Informes EPS > Generar informe actual* y revisar el resultado antes de exportar o imprimir.
5. Para un curso completo, elegir el nivel / curso (`EPS1`…`EPS4`) en `Control` y ejecutar *Informes EPS > Generar boletines por curso*. Se crea un libro nuevo con una hoja por estudiante con `LISTO PARA EXPORTAR = SÍ` en la grilla del período. El libro se crea en Mi unidad, y su organización posterior en carpetas queda a cargo del usuario. La hoja `Resumen`, con generados, omitidos, advertencias y errores, queda oculta para que no se imprima ni se exporte con el libro.
6. Para imprimir o exportar a PDF: A4 vertical, ajustar al ancho, márgenes superior y laterales de 0 cm e inferior de 1 cm, y la opción *Repetir filas inmovilizadas* del diálogo de impresión activada. Así, si un boletín ocupa dos páginas, la segunda empieza con aire superior.

## Organización

- `src/`: versión actual del generador (planilla y Apps Script).
- `tests/`: pruebas locales.
- `reference/`: casos de prueba y datos sintéticos.
- `legacy/`: versiones anteriores, como referencia.

## Versiones

`src/` define el estado versionado del producto. Cada versión es un tag anotado `vX.Y` sobre un commit de `main`: cuando el contenido de `src/` en `main` difiere del último tag, el workflow [`versionar-producto`](.github/workflows/versionar-producto.yml) crea automáticamente el siguiente `Y`. Los cambios fuera de `src/` no generan versión.

No se usa SemVer: el número es correlativo, y los tags no se mueven ni se reutilizan. El salto de serie (por ejemplo, `v1.0`) se crea a mano. Las versiones anteriores al repositorio se conservan como snapshots en `legacy/`.

## Pruebas

En Windows, desde PowerShell o `cmd`, con Node.js:

```
node tests/run.js
```

Los casos cubiertos están descritos en [`reference/casos_de_prueba.md`](reference/casos_de_prueba.md).

## Datos

Los datos de prueba versionados son sintéticos. Las grillas reales y cualquier dato de estudiantes deben quedar fuera de Git.
