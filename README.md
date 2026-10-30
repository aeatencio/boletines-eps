# boletines-eps

Generador de boletines para la Escuela Profesional Secundaria (EPS) del CFP N.º 7, donde se desarrolla y se usa.

El equipo docente carga las valoraciones en grillas de Google Sheets. `boletines-eps` arma a partir de ellas el boletín de cada estudiante con el formato institucional, listo para revisar, exportar e imprimir.

```
Grillas de valoración → Generador (Google Sheets + Apps Script) → Boletín → Revisión, exportación e impresión
```

El proyecto está en desarrollo activo, y el trabajo pendiente se organiza en los issues del repositorio.

## Puesta en funcionamiento

1. Subir `src/Generador_EPS.xlsx` a Google Drive y abrirlo como Google Sheets.
2. En *Extensiones > Apps Script*, pegar `src/Generador_EPS.gs` y, en un segundo archivo de script, `src/Exportar_PDF.gs`. Crear además un archivo HTML llamado `Exportar_PDF_UI` (*Archivo > Nuevo > HTML*; Apps Script no admite dos archivos con el mismo nombre base, como `Exportar_PDF.gs` y `Exportar_PDF.html`) y pegar `src/Exportar_PDF_UI.html`; es el diálogo de la exportación desde un libro o carpeta. Del mismo modo, crear el archivo HTML `Generar_DNI_UI` y pegar `src/Generar_DNI_UI.html`, el diálogo de la generación por DNI. Guardar y recargar.
3. Completar la hoja `Control` con el período, el estudiante y las grillas de origen.
4. Ejecutar *Informes EPS > Generar informe actual* y revisar el resultado antes de exportar o imprimir.
5. Para un curso completo, elegir el nivel / curso (`EPS1`…`EPS4`) en `Control` y ejecutar *Informes EPS > Generar boletines por curso*. Se crea un libro nuevo con una hoja por estudiante con `LISTO PARA EXPORTAR = SÍ` en la grilla del período. El libro se guarda en *Mi unidad > Boletines EPS > <ciclo> - <período>* (por ejemplo, `2026 - 3º bimestre`); las carpetas se crean la primera vez y después se reutilizan. La hoja `Resumen`, con generados, omitidos, advertencias y errores, queda oculta para que no se imprima ni se exporte con el libro.
6. Para varios cursos en una sola operación, marcar con `SÍ` los cursos del bloque *Varios cursos* de `Control` y ejecutar *Informes EPS > Generar boletines de los cursos marcados*; para los 8 cursos (EPSI y EPSG, EPS1…EPS4), ejecutar *Informes EPS > Generar boletines de todos los cursos*, que no mira las marcas. Se usan el período, el ciclo lectivo y las fuentes de `Control`; la hoja origen, el DNI y el nivel se ignoran. Cada curso produce el mismo libro que la generación por curso, en la misma carpeta, y un curso sin estudiantes listos no genera libro. Antes de crear el primer libro se validan todas las fuentes y cursos: un error de configuración o de estructura no deja libros a medias. `Control` registra una línea por curso y la duración total.
7. Para imprimir o exportar a PDF a mano: A4 vertical, ajustar al ancho, márgenes superior y laterales de 0 cm e inferior de 1 cm, y la opción *Repetir filas inmovilizadas* del diálogo de impresión activada. Así, si un boletín ocupa dos páginas, la segunda empieza con aire superior.
8. Para obtener un PDF por estudiante de un libro ya generado y revisado, ejecutar *Informes EPS > Exportar PDF desde un libro o carpeta…*, pegar el link del libro en el diálogo y elegir *Analizar*: muestra qué se va a hacer, sin exportar nada. La exportación empieza recién al elegir *Exportar*, en una ejecución nueva que vuelve a mirar Drive. Así, el tiempo de buscar el link y de revisar no se descuenta de los 5 minutos de la exportación. Si en lugar del libro se pega el link de una carpeta de Drive, se exporta el libro más reciente de cada curso que contenga, como en el punto siguiente pero sin depender de `Control`: el curso, el período y el ciclo salen del `Resumen` de cada libro. Es una operación separada de la generación: no modifica el libro y usa siempre la configuración de impresión del punto anterior, sin depender de lo guardado en el diálogo de impresión. Los PDF, uno por hoja de estudiante y con el nombre de la hoja, quedan en una carpeta nueva junto al libro. Google limita las exportaciones seguidas, así que entre una y otra se espera unos segundos, más si Google lo pide, y un curso completo tarda algunos minutos. `Control` registra cuántos se exportaron, los reintentos, la duración y cada hoja que falló. Si la exportación queda incompleta (por errores o por el tiempo máximo de 5 minutos), se vuelve a analizar el mismo libro y se elige *Completar*, que agrega a esa carpeta los PDF que faltan, o *Exportar todo en una carpeta nueva*. Si el libro se modificó después de esa exportación, no se ofrece completarla: se exporta todo en una carpeta nueva. Para una carpeta de Drive, se vuelve a analizar y exportar la misma carpeta, que continúa donde quedó. La primera vez, Google pide autorizar el acceso a Drive y a servicios externos (la exportación usa la URL de exportación de Google Sheets).
9. Para los PDF de varios cursos, ejecutar *Informes EPS > Exportar PDF de los cursos marcados* (con las mismas marcas `SÍ` del bloque *Varios cursos*) o *Exportar PDF de todos los cursos*. Se usan el período y el ciclo lectivo de `Control`, y de cada curso, su libro más reciente en *Mi unidad > Boletines EPS > <ciclo> - <período>*. Antes de empezar, un diálogo muestra qué se va a hacer con cada curso: los cursos con PDF completos se omiten, las carpetas incompletas se completan y los demás se exportan en una carpeta nueva junto al libro. Si un libro se modificó después de su última exportación, se exporta todo de nuevo en una carpeta nueva. Un libro que no terminó de generarse, o que no es de la generación, queda informado como problema y no se exporta. Cada ejecución avanza hasta 5 minutos. Si `Control` dice *Faltan PDF*, se vuelve a ejecutar la misma opción, que continúa donde quedó sin repetir los PDF ya exportados.
10. Para rehacer los boletines de unos pocos estudiantes, de cualquier curso, ejecutar *Informes EPS > Generar boletines por DNI…*, pegar sus DNI (uno por línea, o separados por espacios, comas o punto y coma) y elegir *Generar*. Se usan el período, el ciclo lectivo y las fuentes de `Control`. Se crea un único libro, con una hoja por estudiante pedido, igual a la de la generación individual, en *Mi unidad > Boletines EPS > <ciclo> - <período> > Por DNI*. Como en la generación individual, `LISTO PARA EXPORTAR` no se mira: el estudiante se pidió explícitamente, y el valor de la grilla se informa. El diálogo informa cada DNI (generado, con error o no encontrado) y los repetidos, y da el link del libro, para revisarlo y para exportar sus PDF con *Exportar PDF desde un libro o carpeta…* (punto 8). Ese libro no reemplaza al de ningún curso en la exportación de los cursos marcados o de todos.

## Organización

- `src/`: estado actual del generador (planilla y Apps Script).
- `tests/`: pruebas locales.
- `reference/`: casos de prueba y datos sintéticos.
- `legacy/`: versiones anteriores, como referencia.

## Versiones

`src/` define el estado versionado del producto. Sus archivos usan nombres estables, sin versión embebida, y Git registra su evolución. Cada versión es un tag anotado `vX.Y` sobre un commit de `main`: cuando el contenido de `src/` en `main` difiere del último tag, el workflow [`versionar-producto`](.github/workflows/versionar-producto.yml) crea automáticamente el siguiente `Y`. Los cambios fuera de `src/` no generan versión.

No se usa SemVer: el número es correlativo, y los tags no se mueven ni se reutilizan. El salto de serie (por ejemplo, `v1.0`) se crea a mano. Las versiones anteriores al repositorio se conservan como snapshots en `legacy/`.

## Pruebas

En Windows, desde PowerShell o `cmd`, con Node.js:

```
node tests/run.js
```

Los casos cubiertos están descritos en [`reference/casos_de_prueba.md`](reference/casos_de_prueba.md).

## Datos

Los datos de prueba versionados son sintéticos. Las grillas reales y cualquier dato de estudiantes deben quedar fuera de Git.
