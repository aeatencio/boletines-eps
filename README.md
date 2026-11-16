# boletines-eps

Herramientas para las valoraciones y los boletines de la Escuela Profesional Secundaria (EPS) del CFP N.º 7, donde se desarrollan y se usan.

Las valoraciones de cada período están en grillas de Google Sheets, una planilla por período. Los docentes cargan las suyas desde la **carga docente**, una página web a la que entran con un enlace personal. El Equipo EPS administra las grillas directamente en Sheets. El **generador** arma, a partir de ellas, el boletín de cada estudiante con el formato institucional, listo para revisar, exportar e imprimir.

```
Carga docente (página web) ─┐
                            ├→ Grillas de valoración → Generador (Google Sheets + Apps Script) → Boletín → Revisión, exportación e impresión
Equipo EPS (Google Sheets) ─┘
```

Son dos proyectos de Apps Script independientes. Sólo comparten las grillas: la carga escribe en ellas y el generador las lee.

El proyecto está en desarrollo activo, y el trabajo pendiente se organiza en los issues del repositorio.

## Generador de boletines

### Instalación

1. Subir `src/Generador_EPS.xlsx` a Google Drive y abrirlo como Google Sheets.
2. En *Extensiones > Apps Script*, pegar `src/Generador_EPS.gs` y, en un segundo archivo de script, `src/Exportar_PDF.gs`. Crear además un archivo HTML llamado `Exportar_PDF_UI` (*Archivo > Nuevo > HTML*; Apps Script no admite dos archivos con el mismo nombre base, como `Exportar_PDF.gs` y `Exportar_PDF.html`) y pegar `src/Exportar_PDF_UI.html`; es el diálogo de la exportación desde un libro o carpeta. Del mismo modo, crear el archivo HTML `Generar_DNI_UI` y pegar `src/Generar_DNI_UI.html`, el diálogo de la generación por DNI. Guardar y recargar.

### Uso

1. Completar la hoja `Control` con el período, el estudiante y las grillas de origen. Los períodos son 1º bimestre, 1º cuatrimestre, 3º bimestre, 2º cuatrimestre y anual, cada uno con su propio libro de grillas (`B10` a `B14`). La grilla del período elegido es la principal, y de ella salen los datos generales, `LISTO PARA EXPORTAR` y `EXPORTADO`. Las grillas de los períodos anteriores sólo aportan sus valoraciones al historial del boletín.
2. Ejecutar *Informes EPS > Generar informe actual* y revisar el resultado antes de exportar o imprimir.
3. Para un curso completo, elegir el nivel / curso (`EPS1`…`EPS4`) en `Control` y ejecutar *Informes EPS > Generar boletines por curso*. Se crea un libro nuevo con una hoja por estudiante con `LISTO PARA EXPORTAR = SÍ` en la grilla del período. El libro se guarda en *Mi unidad > Boletines EPS > <ciclo> - <período>* (por ejemplo, `2026 - 3º bimestre`); las carpetas se crean la primera vez y después se reutilizan. La hoja `Resumen`, con generados, omitidos, advertencias y errores, queda oculta para que no se imprima ni se exporte con el libro.
4. Para varios cursos en una sola operación, marcar con `SÍ` los cursos del bloque *Varios cursos* de `Control` y ejecutar *Informes EPS > Generar boletines de los cursos marcados*; para los 8 cursos (EPSI y EPSG, EPS1…EPS4), ejecutar *Informes EPS > Generar boletines de todos los cursos*, que no mira las marcas. Se usan el período, el ciclo lectivo y las fuentes de `Control`; la hoja origen, el DNI y el nivel se ignoran. Cada curso produce el mismo libro que la generación por curso, en la misma carpeta, y un curso sin estudiantes listos no genera libro. Antes de crear el primer libro se validan todas las fuentes y cursos: un error de configuración o de estructura no deja libros a medias. `Control` registra una línea por curso y la duración total.
5. Para imprimir o exportar a PDF a mano: A4 vertical, ajustar al ancho, márgenes superior y laterales de 0 cm e inferior de 1 cm, y la opción *Repetir filas inmovilizadas* del diálogo de impresión activada. Así, si un boletín ocupa dos páginas, la segunda empieza con aire superior.
6. Para obtener un PDF por estudiante de un libro ya generado y revisado, ejecutar *Informes EPS > Exportar PDF desde un libro o carpeta…*, pegar el link del libro en el diálogo y elegir *Analizar*: muestra qué se va a hacer, sin exportar nada. La exportación empieza recién al elegir *Exportar*, en una ejecución nueva que vuelve a mirar Drive. Así, el tiempo de buscar el link y de revisar no se descuenta de los 5 minutos de la exportación. Si en lugar del libro se pega el link de una carpeta de Drive, se exporta el libro más reciente de cada curso que contenga, como en el punto siguiente pero sin depender de `Control`: el curso, el período y el ciclo salen del `Resumen` de cada libro. Es una operación separada de la generación: no modifica el libro y usa siempre la configuración de impresión del punto anterior, sin depender de lo guardado en el diálogo de impresión. Los PDF, uno por hoja de estudiante y con el nombre de la hoja, quedan en una carpeta nueva junto al libro. Google limita las exportaciones seguidas, así que entre una y otra se espera unos segundos, más si Google lo pide, y un curso completo tarda algunos minutos. `Control` registra cuántos se exportaron, los reintentos, la duración y cada hoja que falló. Si la exportación queda incompleta (por errores o por el tiempo máximo de 5 minutos), se vuelve a analizar el mismo libro y se elige *Completar*, que agrega a esa carpeta los PDF que faltan, o *Exportar todo en una carpeta nueva*. Si el libro se modificó después de esa exportación, no se ofrece completarla: se exporta todo en una carpeta nueva. Para una carpeta de Drive, se vuelve a analizar y exportar la misma carpeta, que continúa donde quedó. La primera vez, Google pide autorizar el acceso a Drive y a servicios externos (la exportación usa la URL de exportación de Google Sheets).
7. Para los PDF de varios cursos, ejecutar *Informes EPS > Exportar PDF de los cursos marcados* (con las mismas marcas `SÍ` del bloque *Varios cursos*) o *Exportar PDF de todos los cursos*. Se usan el período y el ciclo lectivo de `Control`, y de cada curso, su libro más reciente en *Mi unidad > Boletines EPS > <ciclo> - <período>*. Antes de empezar, un diálogo muestra qué se va a hacer con cada curso: los cursos con PDF completos se omiten, las carpetas incompletas se completan y los demás se exportan en una carpeta nueva junto al libro. Si un libro se modificó después de su última exportación, se exporta todo de nuevo en una carpeta nueva. Un libro que no terminó de generarse, o que no es de la generación, queda informado como problema y no se exporta. Cada ejecución avanza hasta 5 minutos. Si `Control` dice *Faltan PDF*, se vuelve a ejecutar la misma opción, que continúa donde quedó sin repetir los PDF ya exportados.
8. Para rehacer los boletines de unos pocos estudiantes, de cualquier curso, ejecutar *Informes EPS > Generar boletines por DNI…*, pegar sus DNI (uno por línea, o separados por espacios, comas o punto y coma) y elegir *Generar*. Se usan el período, el ciclo lectivo y las fuentes de `Control`. Se crea un único libro, con una hoja por estudiante pedido, igual a la de la generación individual, en *Mi unidad > Boletines EPS > <ciclo> - <período> > Por DNI*. Como en la generación individual, `LISTO PARA EXPORTAR` no se mira: el estudiante se pidió explícitamente, y el valor de la grilla se informa. El diálogo informa cada DNI (generado, con error o no encontrado) y los repetidos, y da el link del libro, para revisarlo y para exportar sus PDF con *Exportar PDF desde un libro o carpeta…* (punto 6). Ese libro no reemplaza al de ningún curso en la exportación de los cursos marcados o de todos.

### Actualizar el generador institucional

Las grillas no se tocan: el generador sólo las lee y este procedimiento no las abre.

1. Hacer una copia del libro del generador (*Archivo > Hacer una copia*) como respaldo.
2. En *Extensiones > Apps Script*, reemplazar el contenido de cada archivo que cambió por el de `src/` y guardar.
3. Si la versión nueva cambia la hoja `Control`, aplicar esos cambios a mano en el libro en uso (ver el caso siguiente).
4. Generar con *Generar informe actual* un boletín de un período ya usado y comprobar que coincide con el que se había generado antes.

**Desde un generador anterior a 2º cuatrimestre y anual** (hasta `v0.9`), `Control` necesita dos filas nuevas, y se actualiza junto con `Generador_EPS.gs`. El script nuevo busca las marcas de *Varios cursos* en `B17:B24`, y sólo las usa si cada una está junto al rótulo de su curso, que debe conservar el código entre paréntesis (por ejemplo, `Informática 1 (EPSI EPS1)`). Con la hoja `Control` anterior, la generación y la exportación a PDF de los cursos marcados se detienen con un error, sin generar ni exportar nada; las demás operaciones no leen las marcas. En el paso 3:

1. En `Control`, seleccionar la fila 13 (vacía) e insertar dos filas arriba. Copiar el formato de la fila 12 a las filas nuevas y completar `A13`, `C13`, `A14` y `C14` con los textos de `src/Generador_EPS.xlsx`. Comprobar que *Varios cursos* quedó en `A16` y que cada marca, en `B17:B24`, sigue junto a su curso.
2. En `B5`, en *Datos > Validación de datos*, agregar `2º cuatrimestre` y `Anual` a la lista.
3. Copiar de `src/Generador_EPS.xlsx` los textos actualizados de `A2`, `C5`, `C10` a `C12`, el *Manual breve* (`A27`), los *Valores válidos* (`E27:F30`) y las *Reglas centrales* (`A35` y `A37`).
4. Completar `B13` y `B14` con las grillas de 2º cuatrimestre y anual cuando existan.

Otra opción es armar un generador nuevo como en *Instalación* y copiar a su `Control` el período, el ciclo, las fuentes y las marcas del anterior. Así se obtiene un libro con otro enlace, y hay que volver a autorizar los scripts.

## Carga docente

La carga docente es la página web con la que cada docente carga sus valoraciones, sin editar las grillas generales. Es un proyecto de Apps Script propio, separado del generador, en `src/carga/`. Las grillas siguen siendo la fuente de las valoraciones: la carga escribe en ellas y el generador las lee como siempre. La grilla de cada período es la que figura en `Control` del generador (`B10` a `B14`).

La carga es sólo para docentes. El Equipo EPS sigue administrando en Sheets, con sus propias cuentas, las cursadas, `ACREDITÓ` y `EQUIVALENCIA`, los datos generales y todo lo que no es valoración docente; la carga no tiene roles ni pantallas administrativas.

Se usa desde la carga de 2º cuatrimestre y anual de 2026, cuya ventana abrió el 16/11/2026. La instalación, la prueba previa, la apertura de cada ventana y la operación diaria están en el [manual de la carga docente](docs/puesta-en-marcha-carga-docente.md). Las reglas de esta sección son las que el manual aplica.

### Qué hace el docente

- Ve las materias que tiene asignadas, por período, y en cada una a **todos los estudiantes que la cursan**, sin filtrar por grupo: un estudiante puede cursar materias de otros niveles. `CURSA` en la grilla es una calificación pendiente; con la celda vacía o `NO CURSA`, el estudiante no aparece en esa materia.
- Carga por separado cuatro campos: **Calificación**, con los valores del período (de 1 a 10 en los cuatrimestres y en anual; `En proceso`, `Suficiente` o `Avanzado` en los bimestres); **Calidad de las producciones** y **Participación**, con las opciones cerradas de las grillas; y **Observaciones**, texto libre. Los tres últimos son opcionales y no condicionan la calificación.
- Cada campo se guarda por separado, en el momento, y la página muestra si quedó guardado, si falló o si alguien lo cambió en la grilla mientras tanto. Nunca da por guardado algo que no pudo confirmar.
- `ACREDITÓ` y `EQUIVALENCIA` los pone el Equipo EPS: el docente los ve en sólo lectura, pero puede completar los otros tres campos.
- Elegir «Pendiente» en una calificación la devuelve a `CURSA`, sin tocar los otros campos.
- Cuando una ventana incluye dos períodos, como 2º cuatrimestre y anual, cada uno se carga por separado, en su propia grilla: nada se copia de uno al otro.

### Planilla de administración

El Equipo EPS configura la carga en una planilla de administración, sin tocar el código:

- `Ventanas`: `Período`, `Desde`, `Hasta`, una fila por período habilitado. La carga abre a las 0:00 de `Desde` y cierra a las 8:00 (hora de Buenos Aires) del primer día hábil posterior a `Hasta`. Los períodos con las mismas fechas, como 2º cuatrimestre y anual, comparten ventana, pero cada uno escribe en su propia grilla.
- `Asignaciones`: `Orientación` (EPSI o EPSG), `Nivel` (EPS1 a EPS4), `Materia` (como en el encabezado de la grilla), `Docente` y `Correo`. Un único docente por orientación y materia; si hay dos, o una fila incompleta, nadie carga esa materia hasta corregirla. Un docente puede tener varias materias, de las dos orientaciones. Si se cambia el docente de una materia, el nuevo puede seguir y corregir lo cargado, y el anterior deja de verla.
- `Enlaces`: `Correo` y `Versión`. Normalmente vacía. Para revocar el enlace de un docente (por ejemplo, si lo reenvió), agregar su correo con `Versión` 2, o aumentar la que tenga: el enlace anterior deja de valer en la operación siguiente, también si un guardado está en curso, y el nuevo le llega en el próximo correo, el siguiente día hábil. No borrar filas ni bajar versiones: eso rehabilita enlaces revocados. Sin esta hoja, ningún enlace vale.
- `Registro`: lo escribe la carga, una fila por cambio intentado, con fecha, período, grilla, orientación, nivel, materia, estudiante, campo, valor anterior y nuevo, y docente. No incluye las ediciones que el Equipo EPS hace directamente en Sheets. Su `Estado` distingue lo confirmado de lo inferido:
  - `aplicado`: la escritura se confirmó;
  - `inferido`: la celda mostró el valor nuevo, pero no se puede asegurar que lo haya escrito la carga;
  - `incierto`: la escritura falló de forma ambigua y la celda no muestra el valor nuevo. Pudo no aplicarse, o aplicarse y ser corregida; la carga no la repite;
  - `revisar`: hay que mirar la celda que indica el detalle;
  - `no aplicado`: no se llegó a escribir;
  - `pendiente`: sin resolver.

  Conviene protegerlo para que sólo lo edite la cuenta propietaria, y consultarlo con vistas de filtro.
- `Correos`: lo escribe la carga, una fila por correo enviado o intentado: día, `Estado` (`enviado`, `error` o `pendiente`, si la ejecución se cortó después de enviarlo), docente, tipo (`apertura`, `seguimiento`, `cierre`, `enlace`), períodos informados, calificaciones `CURSA` pendientes y asunto. Nunca guarda el enlace. Es también la memoria que evita un segundo correo el mismo día. Protegerla como el Registro.

Ventanas, asignaciones y la grilla de cada período se leen en cada operación, así que los cambios rigen enseguida, aunque el docente tenga la página abierta. Si el Equipo EPS cambia la grilla de un período en `Control` con la carga abierta, los guardados con la grilla anterior se rechazan. Dos períodos nunca comparten grilla: si en `Control` figuran iguales, ninguno de los dos abre.

Mientras una ventana está abierta, el Equipo EPS no debería insertar, borrar ni ordenar filas o columnas de las grillas. La carga ubica cada celda por DNI y encabezado, y vuelve a comprobarla antes de escribir. Aun así, un cambio de forma en el instante exacto de una escritura puede desviarla; si ocurre, queda `revisar` con la celda a verificar. Editar valores en la grilla sí se puede: la carga no pisa lo que cambió desde que el docente lo vio.

### Días hábiles y feriados

Día hábil es de lunes a viernes, salvo los feriados nacionales y los que rigen en CABA. Deciden el cierre de cada ventana y los días en que salen los correos. Los feriados no se administran en una planilla: están en el código, en `Carga_Calendario.gs`, organizados por año, con sus fuentes oficiales. Los días no laborables (por ejemplo, los turísticos) están anotados aparte y no cambian los días hábiles.

Para un año que no está en el calendario, la carga no adivina: el período cuyo cierre cae en ese año no abre, y los correos de ese día no salen y se informa el problema. Lo demás sigue funcionando. Por eso, **antes de la primera ventana de cada año** hay que agregar el año y publicar una versión nueva de la app web ([manual, F8](docs/puesta-en-marcha-carga-docente.md#f8-actualizar-el-calendario-de-feriados)). La primera línea de `verificarCargaDocente` dice qué años tiene el calendario instalado. Una hoja `Feriados` que haya quedado en la planilla de administración no se lee: no hace falta borrarla.

### Acceso y correos

Cada docente entra con **su enlace**, que le llega por correo. El enlace no se guarda en ningún lado: el servidor lo verifica en cada operación a partir del correo de `Asignaciones`, su versión en `Enlaces` y la ventana vigente. Es el mismo durante toda la ventana (2º cuatrimestre y anual comparten uno), sigue abriendo la página después del cierre, sin nada para cargar, y vence cuando empieza la ventana siguiente. Un docente sin asignaciones no tiene enlace. Quien tenga el enlace actúa como ese docente, así que no hay que reenviarlo. Las asignaciones vigentes siguen decidiendo qué ve y qué carga.

Los correos salen de la cuenta propietaria, como «Equipo EPS», como máximo uno por docente y día, y sólo en días hábiles. Si hay algo que informar y el año no tiene calendario, no se envía ninguno. Cada correo muestra el estado de las grillas en el momento de enviarlo, por período, materia y nivel. Lista a cada estudiante con su calificación, sin DNI y sin el contenido de Calidad, Participación ni Observaciones:

- **Apertura**: el primer día hábil de la ventana, a todos los docentes asignados, con el enlace y el estado inicial. Quien no la recibió (porque lo asignaron después o porque el envío falló) la recibe el siguiente día hábil.
- **Seguimiento**: los días hábiles siguientes, sólo a quien tiene calificaciones `CURSA` pendientes o desaprobaciones (5 o menos, o `En proceso`) con Calidad, Participación u Observaciones vacías. Es sólo un recordatorio: no impide guardar ni generar boletines.
- **Cierre**: el primer día hábil posterior a la ventana, después del cierre de las 8:00, a todos los asignados, con el resumen final. No certifica que la carga esté completa. Si ese día no salió, sale en los 7 días siguientes. El cierre de la carga no depende de este correo: a las 8:00 deja de aceptar cambios aunque el correo falle.
- **Enlace nuevo**: después de una revocación, aunque el docente no tenga pendientes.

Cuando coinciden varios períodos, todo va en un mismo correo.

**Plazo y cierre.** Los correos y la página distinguen dos fechas: el **plazo** para completar la carga, el `Hasta` de `Ventanas`, y el **cierre del sistema**, a las 8:00 del primer día hábil posterior. Lo que media entre ambos es sólo un margen, y el día del cierre no es un día de carga. Por ejemplo, con `Hasta` el viernes 20/11 y feriado el lunes 23/11, el asunto lleva el plazo («plazo viernes 20/11», o «plazo hoy, …» el último día) y el cuerpo lo explica una vez: «El plazo para completar la carga es el viernes 20/11. Como margen, el sistema seguirá abierto hasta el martes 24/11 a las 8:00; después ya no se podrán hacer cambios.» La página muestra «Plazo: viernes 20/11 · el sistema cierra el martes 24/11 a las 8:00», con cuenta regresiva en el margen y en las últimas 24 horas.

Los correos los envía un activador diario, que corre en algún minuto entre las 8:00 y las 9:00 (Apps Script no garantiza el minuto) y nunca envía antes de las 8:00. Si una ejecución tiene problemas, errores de envío o correos sin procesar (por cuota o tiempo), termina con error, y Google avisa por correo a la cuenta propietaria. Ejecutar `enviarCorreosCarga` a mano el mismo día completa lo que faltó, sin repetir a quien ya lo recibió.

### Funciones del editor y propiedades

Sólo la cuenta propietaria las ejecuta, desde el editor del proyecto:

| Función | Qué hace |
|---|---|
| `prepararAdministracionCarga` | Crea las hojas de administración que falten, sólo con sus encabezados, y `CLAVE_ENLACES` si falta |
| `verificarCargaDocente` | Informa problemas de configuración, de las grillas, del Registro, de los enlaces y de los correos, y resume lo que contiene cada grilla de las ventanas vigentes o próximas (sirve para revisar una grilla recién preparada) |
| `resolverPendientesCarga` | Resuelve las filas `pendiente` viejas del Registro mirando la celda |
| `activarCorreosCarga` / `desactivarCorreosCarga` | Crea o borra el activador diario de los correos |
| `previsualizarCorreosCarga` | Muestra en el registro de ejecución qué se enviaría en ese momento, sin enviar nada |
| `enviarCorreosCarga` | El envío diario; a mano, completa el del día |
| `previsualizarPeriodosCarga` / `prepararPeriodosCarga` | Preparación de 2º cuatrimestre y anual (abajo) |

Propiedades del script:

- `ADMIN_ID`: la planilla de administración.
- `GENERADOR_ID`: la planilla del generador.
- `URL_CARGA`: la URL `/exec` de la app web publicada, para los enlaces de los correos.
- `CLAVE_ENLACES`: el secreto de los enlaces. Lo crea `prepararAdministracionCarga`. Borrarlo y volver a crearlo invalida todos los enlaces a la vez; con una ventana abierta, hay que aumentar además la versión de cada docente en `Enlaces` para que todos reciban el nuevo, porque los correos no detectan el cambio de secreto.

### Preparación de 2º cuatrimestre y anual

Las grillas de 2º cuatrimestre y anual se arman copiando la de 1º cuatrimestre, y hay que dejar `CURSA` donde corresponde, limpiar lo que quedó de la copia y asignar un docente a cada materia. `prepararPeriodosCarga` lo hace en una sola ejecución, a partir de lo que el sistema ya sabe; `previsualizarPeriodosCarga` muestra lo mismo sin modificar nada. Las dos escriben un informe en la hoja `Preparación` de la planilla de administración, con lo hecho y los casos que necesitan una decisión del Equipo EPS, por prioridad (**Bloqueo**, **Decidir**, **Revisar**).

- **Cursadas**: la evidencia es 3º bimestre, el último período calificado, con la matrícula vigente. Una calificación ordinaria allí deja `CURSA` en el borrador vacío, con `NO CURSA` o con la nota copiada de 1º cuatrimestre: vaciar las notas copiadas fue limpiar la copia, no decidir que el estudiante dejó de cursar. `ACREDITÓ`, `EQUIVALENCIA` o `CURSA` sin calificar no son evidencia. Anual usa la misma evidencia, porque 2º cuatrimestre todavía no está calificado.
- **Restos de la copia**: ninguna calificación de 1º cuatrimestre queda como valoración nueva sin informarse, y se borran Calidad, Participación y Observaciones iguales a las de 1º cuatrimestre.
- **Asignaciones**: arma la hoja `Asignaciones (borrador)`, con una fila por orientación y materia con cursantes. No hay una fuente fiable del docente de cada materia: el Equipo EPS lo elige de una lista armada con su nómina (hoja `Docentes`). No toca la hoja `Asignaciones`.
- **Revisión general**: matrícula, nombres y DNI, grupos, orientación, materias y niveles de cada borrador contra 3º bimestre y 1º cuatrimestre, y de los dos borradores entre sí.

Nunca inserta, borra ni mueve filas o columnas, ni toca `ORIENTACIÓN` y los campos posteriores, 1º cuatrimestre o 3º bimestre. Conserva las decisiones del Equipo EPS (estados especiales, notas que no vienen de la copia, cursadas quitadas después de prepararlas) y lo que la carga docente ya escribió. Con una fila de 2º cuatrimestre o anual en `Ventanas`, no modifica nada. Anota cada celda que cambia en `Preparación · cambios`, y se puede repetir sin deshacer ni duplicar lo hecho. Las reglas completas y el procedimiento están en el [manual, paso A3.1](docs/puesta-en-marcha-carga-docente.md#a31-preparación-automática-de-cursadas-y-asignaciones).

## Organización

- `src/`: estado actual del generador (planilla y Apps Script); en `src/carga/`, la carga docente.
- `docs/`: manual de la carga docente para el Equipo EPS ([`puesta-en-marcha-carga-docente.md`](docs/puesta-en-marcha-carga-docente.md)).
- `tests/`: pruebas locales.
- `reference/`: casos de prueba del generador y datos sintéticos.
- `legacy/`: versiones anteriores, como referencia.

El problema que dio origen a la carga docente y sus reglas institucionales están en [#13](https://github.com/aeatencio/boletines-eps/issues/13); el plan de trabajo, en [#30](https://github.com/aeatencio/boletines-eps/issues/30).

## Versiones

`src/` define el estado versionado del producto. Sus archivos usan nombres estables, sin versión embebida, y Git registra su evolución. Cada versión es un tag anotado `vX.Y` sobre un commit de `main`: cuando el contenido de `src/` en `main` difiere del último tag, el workflow [`versionar-producto`](.github/workflows/versionar-producto.yml) crea automáticamente el siguiente `Y`. Los cambios fuera de `src/` no generan versión.

No se usa SemVer: el número es correlativo, y los tags no se mueven ni se reutilizan. El salto de serie (por ejemplo, `v1.0`) se crea a mano. Las versiones anteriores al repositorio se conservan como snapshots en `legacy/`.

## Pruebas

Todas las pruebas son locales, con Node.js, y usan dobles de los servicios de Google: no reemplazan una comprobación en Google real.

Las del generador necesitan además el `tar.exe` de Windows para leer los `.xlsx` (en Windows, desde PowerShell o `cmd`; también desde WSL):

```
node tests/run.js
```

Los casos cubiertos están descritos en [`reference/casos_de_prueba.md`](reference/casos_de_prueba.md).

Las de la carga docente funcionan en cualquier sistema. Cada prueba describe su caso al ejecutarse. El número opcional es la cantidad de corridas de la simulación aleatoria (30 por defecto en `carga.js`, 10 en `carga-correo.js` y 40 en `carga-preparacion.js`):

```
node tests/carga.js [corridas]              # núcleo de consulta y guardado
node tests/carga-interfaz.js                # página docente
node tests/carga-acceso.js                  # enlaces
node tests/carga-correo.js [corridas]       # correos (simula una ventana completa)
node tests/carga-preparacion.js [corridas]  # preparación de 2º cuatrimestre y anual
node tests/carga-recorrido.js               # recorrido del manual
```

`carga-recorrido.js` recorre los pasos del [manual](docs/puesta-en-marcha-carga-docente.md), de la verificación a la limpieza, y falla si el código deja de coincidir con los resultados que el manual cita.

Para usar la página docente en el navegador, con datos ficticios:

```
node tests/carga-dev.js [puerto]
```

El servidor local escucha sólo en `127.0.0.1` (puerto 8787 por defecto) y muestra un enlace por docente ficticio. Usa los archivos reales de `src/carga/`, sin sustituir nada, con un secreto de enlaces ficticio. En `/dev` hay un panel para simular fallas de guardado y cambios del Equipo EPS con la página abierta (incluida la revocación de un enlace) y para ver el Registro. En `/dev/correos` se ven los correos que se enviarían en ese momento; no se envía nada.

## Datos

Los datos de prueba versionados son sintéticos. Las grillas reales y cualquier dato de estudiantes deben quedar fuera de Git.
