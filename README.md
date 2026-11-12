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
3. Completar la hoja `Control` con el período, el estudiante y las grillas de origen. Los períodos son 1º bimestre, 1º cuatrimestre, 3º bimestre, 2º cuatrimestre y anual, cada uno con su propio libro de grillas (`B10` a `B14`). La grilla del período elegido es la principal, y de ella salen los datos generales, `LISTO PARA EXPORTAR` y `EXPORTADO`. Las grillas de los períodos anteriores sólo aportan sus valoraciones al historial del boletín.
4. Ejecutar *Informes EPS > Generar informe actual* y revisar el resultado antes de exportar o imprimir.
5. Para un curso completo, elegir el nivel / curso (`EPS1`…`EPS4`) en `Control` y ejecutar *Informes EPS > Generar boletines por curso*. Se crea un libro nuevo con una hoja por estudiante con `LISTO PARA EXPORTAR = SÍ` en la grilla del período. El libro se guarda en *Mi unidad > Boletines EPS > <ciclo> - <período>* (por ejemplo, `2026 - 3º bimestre`); las carpetas se crean la primera vez y después se reutilizan. La hoja `Resumen`, con generados, omitidos, advertencias y errores, queda oculta para que no se imprima ni se exporte con el libro.
6. Para varios cursos en una sola operación, marcar con `SÍ` los cursos del bloque *Varios cursos* de `Control` y ejecutar *Informes EPS > Generar boletines de los cursos marcados*; para los 8 cursos (EPSI y EPSG, EPS1…EPS4), ejecutar *Informes EPS > Generar boletines de todos los cursos*, que no mira las marcas. Se usan el período, el ciclo lectivo y las fuentes de `Control`; la hoja origen, el DNI y el nivel se ignoran. Cada curso produce el mismo libro que la generación por curso, en la misma carpeta, y un curso sin estudiantes listos no genera libro. Antes de crear el primer libro se validan todas las fuentes y cursos: un error de configuración o de estructura no deja libros a medias. `Control` registra una línea por curso y la duración total.
7. Para imprimir o exportar a PDF a mano: A4 vertical, ajustar al ancho, márgenes superior y laterales de 0 cm e inferior de 1 cm, y la opción *Repetir filas inmovilizadas* del diálogo de impresión activada. Así, si un boletín ocupa dos páginas, la segunda empieza con aire superior.
8. Para obtener un PDF por estudiante de un libro ya generado y revisado, ejecutar *Informes EPS > Exportar PDF desde un libro o carpeta…*, pegar el link del libro en el diálogo y elegir *Analizar*: muestra qué se va a hacer, sin exportar nada. La exportación empieza recién al elegir *Exportar*, en una ejecución nueva que vuelve a mirar Drive. Así, el tiempo de buscar el link y de revisar no se descuenta de los 5 minutos de la exportación. Si en lugar del libro se pega el link de una carpeta de Drive, se exporta el libro más reciente de cada curso que contenga, como en el punto siguiente pero sin depender de `Control`: el curso, el período y el ciclo salen del `Resumen` de cada libro. Es una operación separada de la generación: no modifica el libro y usa siempre la configuración de impresión del punto anterior, sin depender de lo guardado en el diálogo de impresión. Los PDF, uno por hoja de estudiante y con el nombre de la hoja, quedan en una carpeta nueva junto al libro. Google limita las exportaciones seguidas, así que entre una y otra se espera unos segundos, más si Google lo pide, y un curso completo tarda algunos minutos. `Control` registra cuántos se exportaron, los reintentos, la duración y cada hoja que falló. Si la exportación queda incompleta (por errores o por el tiempo máximo de 5 minutos), se vuelve a analizar el mismo libro y se elige *Completar*, que agrega a esa carpeta los PDF que faltan, o *Exportar todo en una carpeta nueva*. Si el libro se modificó después de esa exportación, no se ofrece completarla: se exporta todo en una carpeta nueva. Para una carpeta de Drive, se vuelve a analizar y exportar la misma carpeta, que continúa donde quedó. La primera vez, Google pide autorizar el acceso a Drive y a servicios externos (la exportación usa la URL de exportación de Google Sheets).
9. Para los PDF de varios cursos, ejecutar *Informes EPS > Exportar PDF de los cursos marcados* (con las mismas marcas `SÍ` del bloque *Varios cursos*) o *Exportar PDF de todos los cursos*. Se usan el período y el ciclo lectivo de `Control`, y de cada curso, su libro más reciente en *Mi unidad > Boletines EPS > <ciclo> - <período>*. Antes de empezar, un diálogo muestra qué se va a hacer con cada curso: los cursos con PDF completos se omiten, las carpetas incompletas se completan y los demás se exportan en una carpeta nueva junto al libro. Si un libro se modificó después de su última exportación, se exporta todo de nuevo en una carpeta nueva. Un libro que no terminó de generarse, o que no es de la generación, queda informado como problema y no se exporta. Cada ejecución avanza hasta 5 minutos. Si `Control` dice *Faltan PDF*, se vuelve a ejecutar la misma opción, que continúa donde quedó sin repetir los PDF ya exportados.
10. Para rehacer los boletines de unos pocos estudiantes, de cualquier curso, ejecutar *Informes EPS > Generar boletines por DNI…*, pegar sus DNI (uno por línea, o separados por espacios, comas o punto y coma) y elegir *Generar*. Se usan el período, el ciclo lectivo y las fuentes de `Control`. Se crea un único libro, con una hoja por estudiante pedido, igual a la de la generación individual, en *Mi unidad > Boletines EPS > <ciclo> - <período> > Por DNI*. Como en la generación individual, `LISTO PARA EXPORTAR` no se mira: el estudiante se pidió explícitamente, y el valor de la grilla se informa. El diálogo informa cada DNI (generado, con error o no encontrado) y los repetidos, y da el link del libro, para revisarlo y para exportar sus PDF con *Exportar PDF desde un libro o carpeta…* (punto 8). Ese libro no reemplaza al de ningún curso en la exportación de los cursos marcados o de todos.

## Actualizar el generador institucional

Para que un generador ya en uso admita 2º cuatrimestre y anual, se actualizan juntos su hoja `Control` y el script `Generador_EPS.gs`. Las grillas no se tocan: el generador sólo las lee y este procedimiento no las abre. El script nuevo busca las marcas de *Varios cursos* en `B17:B24`, y sólo las usa si cada una está junto al rótulo de su curso, que debe conservar el código entre paréntesis (por ejemplo, `Informática 1 (EPSI EPS1)`). Con la hoja `Control` anterior, la generación y la exportación a PDF de los cursos marcados se detienen con un error, sin generar ni exportar nada. Las demás operaciones no leen las marcas.

1. Hacer una copia del libro del generador (*Archivo > Hacer una copia*) como respaldo.
2. En `Control`, seleccionar la fila 13 (vacía) e insertar dos filas arriba. Copiar el formato de la fila 12 a las filas nuevas y completar `A13`, `C13`, `A14` y `C14` con los textos de `src/Generador_EPS.xlsx`. Comprobar que *Varios cursos* quedó en `A16` y que cada marca, en `B17:B24`, sigue junto a su curso.
3. En `B5`, en *Datos > Validación de datos*, agregar `2º cuatrimestre` y `Anual` a la lista.
4. Copiar de `src/Generador_EPS.xlsx` los textos actualizados de `A2`, `C5`, `C10` a `C12`, el *Manual breve* (`A27`), los *Valores válidos* (`E27:F30`) y las *Reglas centrales* (`A35` y `A37`).
5. En *Extensiones > Apps Script*, reemplazar el contenido de `Generador_EPS.gs` por `src/Generador_EPS.gs` y guardar.
6. Completar `B13` y `B14` con las grillas de 2º cuatrimestre y anual cuando existan. Antes de usarlo, generar con *Generar informe actual* un boletín de un período ya usado y comprobar que coincide con el que se había generado antes.

Otra opción es armar un generador nuevo como en *Puesta en funcionamiento* y copiar a su `Control` el período, el ciclo, las fuentes y las marcas del anterior. Así se obtiene un libro con otro enlace, y hay que volver a autorizar los scripts.

## Carga docente (en desarrollo)

La carga docente reemplazará la edición directa de las grillas por parte de los docentes ([#30](https://github.com/aeatencio/boletines-eps/issues/30)). Es un proyecto de Apps Script propio, separado del generador, en `src/carga/`. Las grillas siguen siendo la fuente de las valoraciones: la carga escribe en ellas y el generador las lee como hasta ahora.

**Implementado**, sólo con pruebas locales:
- [#26](https://github.com/aeatencio/boletines-eps/issues/26): el núcleo de consulta y guardado por campo (`Carga_Nucleo.gs`) y su adaptador a los servicios de Google (`Carga_Google.gs`); el calendario de feriados (`Carga_Calendario.gs`, #29);
- [#27](https://github.com/aeatencio/boletines-eps/issues/27), primer incremento: la página docente (`Carga_Web.gs`, con `doGet`, y los archivos HTML `Carga_Pagina` y `Carga_Cliente`). El docente recorre sus materias por período, ve a todos los que las cursan y guarda cada campo por separado, con estados de guardado, conflicto y resultado sin confirmar;
- [#28](https://github.com/aeatencio/boletines-eps/issues/28): el enlace individual de cada docente (`Carga_Acceso.gs`) y los correos de apertura, seguimiento y cierre (`Carga_Correo.gs`, y su adaptador a MailApp y al activador diario, `Carga_Correo_Google.gs`);
- [#32](https://github.com/aeatencio/boletines-eps/issues/32): la preparación de las grillas de 2º cuatrimestre y anual, copiadas de 1º cuatrimestre, y del borrador de asignaciones (`Carga_Preparacion.gs` y su adaptador, `Carga_Preparacion_Google.gs`). Ver más abajo.

**Pendiente** ([#29](https://github.com/aeatencio/boletines-eps/issues/29)): la prueba integrada controlada sobre la infraestructura definitiva (acceso a la app web, escrituras reales, entrega del correo y activador) y la habilitación institucional. Las dos se hacen con el manual operativo [`docs/puesta-en-marcha-carga-docente.md`](docs/puesta-en-marcha-carga-docente.md), que reemplaza al piloto separado con datos ficticios.

**No habilitar la carga a los docentes antes de completar esa prueba y de que el Equipo EPS autorice la apertura.** El acceso y los correos sólo se probaron con dobles locales de Google. Publicar la app web no la habilita: sin filas en `Ventanas` ningún enlace vale, y los correos sólo salen hacia quienes figuran en `Asignaciones`.

El Equipo EPS configura la carga en una planilla de administración, sin tocar el código:

- `Ventanas`: `Período`, `Desde`, `Hasta`, una fila por período habilitado. La carga abre a las 0:00 de `Desde` y cierra a las 8:00 (hora de Buenos Aires) del primer día hábil posterior a `Hasta`. Los períodos con las mismas fechas, como 2º cuatrimestre y anual, comparten ventana, pero cada uno escribe en su propia grilla.
- `Asignaciones`: `Orientación` (EPSI o EPSG), `Nivel` (EPS1 a EPS4), `Materia` (como en el encabezado de la grilla), `Docente` y `Correo`. Un único docente por orientación y materia; si hay dos, o una fila incompleta, nadie carga esa materia hasta corregirla.
- `Enlaces`: `Correo` y `Versión`. Normalmente vacía. Para revocar el enlace de un docente (por ejemplo, si lo reenvió), agregar su correo con `Versión` 2, o aumentar la que tenga: el enlace anterior deja de valer en la operación siguiente, también si un guardado está en curso, y el nuevo le llega en el próximo correo, el siguiente día hábil. No borrar filas ni bajar versiones: eso rehabilita enlaces revocados. Sin esta hoja, ningún enlace vale.
- `Registro`: lo escribe la carga, una fila por cambio intentado, con fecha, período, grilla, orientación, nivel, materia, estudiante, campo, valor anterior y nuevo, y docente. Su `Estado` distingue lo confirmado de lo inferido:
  - `aplicado`: la escritura se confirmó;
  - `inferido`: la celda mostró el valor nuevo, pero no se puede asegurar que lo haya escrito la carga;
  - `incierto`: la escritura falló de forma ambigua y la celda no muestra el valor nuevo. Pudo no aplicarse, o aplicarse y ser corregida; la carga no la repite;
  - `revisar`: hay que mirar la celda que indica el detalle;
  - `no aplicado`: no se llegó a escribir;
  - `pendiente`: sin resolver.

  Conviene protegerlo para que sólo lo edite la cuenta propietaria, y consultarlo con vistas de filtro.
- `Correos`: lo escribe la carga, una fila por correo enviado o intentado: día, `Estado` (`enviado`, `error` o `pendiente`, si la ejecución se cortó después de enviarlo), docente, tipo (`apertura`, `seguimiento`, `cierre`, `enlace`), períodos informados, calificaciones `CURSA` pendientes y asunto. Nunca guarda el enlace. Es también la memoria que evita un segundo correo el mismo día. Protegerla como el Registro.

Mientras una ventana está abierta, el Equipo EPS no debería insertar, borrar ni ordenar filas o columnas de las grillas. La carga ubica cada celda por DNI y encabezado, y vuelve a comprobarla antes de escribir. Aun así, un cambio de forma en el instante exacto de una escritura puede desviarla; si ocurre, queda `revisar` con la celda a verificar. Editar valores en la grilla sí se puede: la carga no pisa lo que cambió desde que el docente lo vio.

La grilla de cada período es la que figura en `Control` del generador (`B10` a `B14`), leída en cada operación. Si el Equipo EPS la cambia con la carga abierta, los guardados con la grilla anterior se rechazan. Dos períodos nunca comparten grilla: si en `Control` figuran iguales, ninguno de los dos abre. Ventanas y asignaciones también se leen en cada operación.

Los feriados no se administran en una planilla: están en el código, en `Carga_Calendario.gs`, organizados por año, con sus fuentes oficiales. Hoy incluye 2026. Día hábil es de lunes a viernes, salvo los feriados nacionales y los que rigen en CABA. Los días no laborables (por ejemplo, los turísticos) están anotados aparte y no cambian los días hábiles. Lo que necesita un día de un año sin calendario no adivina: el período cuyo cierre cae en ese año no abre, y los correos de ese día no salen y se informa el problema. Lo demás sigue funcionando. Al preparar un ciclo lectivo, se agrega el año y se publica una versión nueva de la app web ([manual](docs/puesta-en-marcha-carga-docente.md#f8-actualizar-el-calendario-de-feriados)). Una hoja `Feriados` que haya quedado en la planilla de administración no se lee: no hace falta borrarla.

Propiedades del script: `ADMIN_ID` (planilla de administración), `GENERADOR_ID` (planilla del generador), `CLAVE_ENLACES` (secreto de los enlaces; lo crea `prepararAdministracionCarga`; borrarlo y volver a crearlo invalida todos los enlaces a la vez; con una ventana abierta, hay que aumentar además la versión de cada docente en `Enlaces` para que todos reciban el nuevo, porque los correos no detectan el cambio de secreto) y `URL_CARGA` (la URL `/exec` de la app web publicada, para los enlaces de los correos). Desde el editor, la cuenta propietaria puede ejecutar `prepararAdministracionCarga` (crea las hojas que falten, sólo con sus encabezados, y `CLAVE_ENLACES` si falta), `verificarCargaDocente` (problemas de configuración, de las grillas, del Registro, de los enlaces y de los correos, y un resumen de lo que contiene cada grilla de las ventanas vigentes o próximas, para revisar una grilla recién preparada) `resolverPendientesCarga` (resuelve las filas `pendiente` viejas mirando la celda) y, antes de cargar las ventanas de 2º cuatrimestre y anual, `previsualizarPeriodosCarga` y `prepararPeriodosCarga` (#32, más abajo).

### Preparación de 2º cuatrimestre y anual (#32)

Los borradores de esas dos grillas se copiaron de la de 1º cuatrimestre y se editaron a mano. `prepararPeriodosCarga`, desde el editor y con la cuenta propietaria, los deja listos para la carga en una sola ejecución; `previsualizarPeriodosCarga` muestra lo mismo sin modificar nada. Las dos escriben un informe en la hoja `Preparación` de la planilla de administración, con lo hecho y las excepciones por prioridad (**Bloqueo**, **Decidir**, **Revisar**).

- **Cursadas**: `CURSA` donde 3º bimestre (último período calificado, matrícula vigente) tiene una calificación ordinaria. `ACREDITÓ`, `EQUIVALENCIA`, `CURSA` sin calificar o vacío no habilitan `CURSA`. Anual usa la misma evidencia: 2º cuatrimestre no está calificado.
- **Restos de la copia**: sólo se modifica lo que sigue igual a 1º cuatrimestre (o una calificación vacía en una fila o columna agregada): una nota copiada pasa a `CURSA` sólo con evidencia, y Calidad, Participación y Observaciones copiadas se borran. Ninguna calificación anterior queda como valoración nueva sin informarse.
- **Asignaciones (borrador)**: una fila por orientación y materia con cursantes, sin transcribir orientación, nivel ni materia. El correo, sólo si todas las protecciones por columna de las grillas que cubren la materia nombran a la misma persona. No toca la hoja `Asignaciones`.
- **Revisión general**: matrícula, nombres y DNI, grupos, orientación, materias y niveles de cada borrador contra 3º bimestre y 1º cuatrimestre, y de los dos borradores entre sí.

Nunca inserta, borra ni mueve filas o columnas, ni toca `ORIENTACIÓN` y los campos posteriores, 1º cuatrimestre o 3º bimestre. Conserva los cambios del Equipo EPS y lo que la carga docente ya escribió (`Registro`). Con una fila de 2º cuatrimestre o anual en `Ventanas`, no modifica nada. Escribe con el lock del script, vuelve a leer cada hoja antes de escribirla y anota cada celda en `Preparación · cambios` (valor anterior, nuevo, motivo y estado verificado). Se puede repetir: lo ya preparado no cambia y las filas del borrador no se duplican. El procedimiento está en el [manual, paso A3.1](docs/puesta-en-marcha-carga-docente.md#a31-preparación-automática-de-cursadas-y-asignaciones-32).

### Acceso y correos (#28)

Cada docente entra con **su enlace**, que le llega por correo. El enlace no se guarda en ningún lado: el servidor lo verifica en cada operación a partir del correo de `Asignaciones`, su versión en `Enlaces` y la ventana vigente. Es el mismo durante toda la ventana (2º cuatrimestre y anual comparten uno), sigue abriendo la página después del cierre, sin nada para cargar, y vence cuando empieza la ventana siguiente. Un docente sin asignaciones no tiene enlace. Quien tenga el enlace actúa como ese docente, así que no hay que reenviarlo. Las asignaciones vigentes siguen decidiendo qué ve y qué carga.

Los correos salen de la cuenta propietaria, como «Equipo EPS», como máximo uno por docente y día, y sólo en días hábiles de CABA (lunes a viernes que no sean feriado, según `Carga_Calendario.gs`). Si hay algo que informar y el año no tiene calendario, no se envía ninguno. Cada correo muestra el estado de las grillas en el momento de enviarlo, por período, materia y nivel. Lista a cada estudiante con su calificación, sin DNI y sin el contenido de Calidad, Participación ni Observaciones:

- **Apertura**: el primer día hábil de la ventana, a todos los docentes asignados, con el enlace y el estado inicial. Quien no la recibió (porque lo asignaron después o porque el envío falló) la recibe el siguiente día hábil.
- **Seguimiento**: los días hábiles siguientes, sólo a quien tiene calificaciones `CURSA` pendientes o desaprobaciones (5 o menos, o `En proceso`) con Calidad, Participación u Observaciones vacías.
- **Cierre**: el primer día hábil posterior a la ventana, después del cierre de las 8:00, a todos los asignados, con el resumen final. No certifica que la carga esté completa. Si ese día no salió, sale en los 7 días siguientes.
- **Enlace nuevo**: después de una revocación, aunque el docente no tenga pendientes.

Cuando coinciden varios períodos, todo va en un mismo correo.

Para activarlos: publicar la app web (ejecutar como la cuenta propietaria), guardar su URL `/exec` en `URL_CARGA` y ejecutar `activarCorreosCarga`. La instalación completa, la prueba y la habilitación, con sus comprobaciones, están en el [manual de puesta en marcha](docs/puesta-en-marcha-carga-docente.md). El activador corre todos los días en algún minuto entre las 8:00 y las 9:00 (Apps Script no garantiza el minuto) y nunca envía antes de las 8:00. `previsualizarCorreosCarga` muestra en el registro de ejecución qué se enviaría en ese momento, sin enviar nada. `desactivarCorreosCarga` borra el activador. Si una ejecución tiene problemas, errores de envío o correos sin procesar (por cuota o tiempo), termina con error, y Google avisa por correo a la cuenta propietaria. Ejecutar `enviarCorreosCarga` a mano el mismo día completa lo que faltó, sin repetir a quien ya lo recibió.

## Organización

- `src/`: estado actual del generador (planilla y Apps Script); en `src/carga/`, la carga docente.
- `docs/`: manual operativo de la carga docente para el Equipo EPS ([`puesta-en-marcha-carga-docente.md`](docs/puesta-en-marcha-carga-docente.md)).
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

Las pruebas de la carga docente sólo necesitan Node.js, en cualquier sistema. Cada prueba describe su caso; el número opcional es la cantidad de corridas de la simulación aleatoria (30 por defecto):

```
node tests/carga.js [corridas]
```

Usan dobles de los servicios de Google: no reemplazan el piloto en Google real.

La interfaz docente tiene sus propias pruebas y un servidor local para usarla en el navegador, con datos ficticios:

```
node tests/carga-interfaz.js
node tests/carga-dev.js [puerto]
```

El servidor local escucha sólo en `127.0.0.1` (puerto 8787 por defecto) y muestra un enlace por docente ficticio. Usa los archivos reales de `src/carga/`, sin sustituir nada, con un secreto de enlaces ficticio. En `/dev` hay un panel para simular fallas de guardado y cambios del Equipo EPS con la página abierta (incluida la revocación de un enlace) y para ver el Registro. En `/dev/correos` se ven los correos que se enviarían en ese momento; no se envía nada.

El acceso y los correos tienen sus propias pruebas. El número opcional es la cantidad de corridas de la simulación de una ventana completa (10 por defecto):

```
node tests/carga-acceso.js
node tests/carga-correo.js [corridas]
```

El recorrido del [manual de puesta en marcha](docs/puesta-en-marcha-carga-docente.md), de la verificación a la limpieza, tiene su ensayo local: falla si el código deja de coincidir con los resultados que cita el manual.

```
node tests/carga-recorrido.js
```

La preparación de 2º cuatrimestre y anual (#32) tiene sus propias pruebas, con un escenario institucional ficticio y una prueba aleatoria de invariantes (el número opcional es la cantidad de corridas; 40 por defecto):

```
node tests/carga-preparacion.js [corridas]
```

## Datos

Los datos de prueba versionados son sintéticos. Las grillas reales y cualquier dato de estudiantes deben quedar fuera de Git.
