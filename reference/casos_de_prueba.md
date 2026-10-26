# Casos de prueba del generador EPS

Este documento describe los escenarios mínimos de prueba reproducibles para el Generador EPS. Las secciones 1 a 4 vienen de v0.3 (incorporación del 3º bimestre, acumulación con 1º bimestre y 1º cuatrimestre y regresiones históricas); las secciones 5 a 8 cubren v0.4 (grillas reales, generación por curso y orientación por período), la sección 11, la generación de varios cursos (#20); la sección 13, la exportación a PDF (#21), la sección 14, la organización de los libros en Drive (#23); la sección 15, la exportación a PDF de varios cursos (#23); la sección 16, la exportación desde un libro o una carpeta de Drive, y la sección 17, su diálogo.

## Arnés de pruebas local

Para facilitar la reproducibilidad sin requerir interacción manual con Google Sheets en cada iteración de desarrollo, los escenarios están automatizados en `tests/run.js`.

### Características y funcionamiento del arnés

- **Ejecución de código real**: `tests/run.js` carga y ejecuta el archivo de código fuente real `src/Generador_EPS.gs` dentro de un contexto aislado (`node:vm`).
- **Mocks locales mínimos**: Emula exclusivamente la superficie mínima de Google Apps Script requerida por la lógica de negocio (`SpreadsheetApp.getActiveSpreadsheet()`, `SpreadsheetApp.openById()`, `SpreadsheetApp.create()`, `Utilities.formatDate()`, `Sheet`, `Range`, `Ui.alert()`, `Ui.createMenu()`; para guardar los libros, `DriveApp` con carpetas, papelera y `moveTo()`; para la exportación de varios libros, también `Folder.getFilesByType()`, `DriveApp.getFolderById()`, el tipo de cada archivo y `MimeType`; para la exportación a PDF, también `Ui.prompt()`, `UrlFetchApp.fetch()`, `DriveApp` y `ScriptApp.getOAuthToken()`).
- **Consumo directo de fixtures vigentes**: En cada ejecución se vacía el directorio de trabajo temporal `tmp/unzipped/` y se extraen directamente los archivos XLSX y TSV que existen en `reference/fixtures/`. Si falta algún fixture o se modifica, el runner valida estrictamente el estado actual del disco y no reutiliza extracciones de corridas previas.
- **Requisito de plataforma y shell**: El runner debe ejecutarse en un entorno **Windows** desde **PowerShell** o **`cmd`**, dependiendo exclusivamente del comando `tar.exe` provisto por el sistema operativo en `%SystemRoot%\System32` (bsdtar nativo de Windows). **No debe asumirse que funciona desde Git Bash**, ya que en dicho entorno `tar.exe` puede resolverse al binario provisto por MSYS y fallar al abrir los archivos XLSX empaquetados.

### Aspectos NO cubiertos por el arnés local

El arnés local valida la lógica de lectura tabular, normalización, acumulación, descarte de valores y asignación en la cuadrícula de salida. **NO cubre ni emula**:
1. El motor completo de Google Apps Script ni la infraestructura de Google Cloud / Google Drive (permisos, autenticación, cuotas, triggers instalables).
2. El motor de cálculo de Google Sheets ni validaciones nativas de datos en celdas.
3. Aspectos visuales y de formato del workbook: tipografías, alineación, bordes, anchos de columna y altos de fila efectivos en Google Sheets.
4. Evaluación del formato condicional en tiempo real en la interfaz de Google Sheets.
5. Importación real de planillas en formato nativo de Google Sheets.
6. Presentación final para impresión ni el PDF que produce Google: la exportación se prueba hasta el pedido (URL, parámetros y manejo de la respuesta).

Por estas razones, la tarea **#4 (verificación manual en entorno real de Google Sheets)** sigue siendo estrictamente necesaria para validar visualización, menús de usuario, permisos y la experiencia operativa completa antes de cualquier liberación.

### Cómo ejecutar las pruebas (PowerShell o cmd)
```bash
node tests/run.js
```
El script finaliza con código `0` si todas las aserciones pasan, o código `1` indicando las fallas detectadas.

## Fuentes de datos (Fixtures)

Todos los fixtures utilizan datos ficticios, sin información de estudiantes reales:
- `reference/fixtures/Valoraciones_ficticias_EPS_1er_bimestre_v0_3.xlsx`: planilla fuente ficticia de 1º bimestre (idéntica a v0.2).
- `reference/fixtures/Valoraciones_ficticias_EPS_1er_cuatrimestre_v0_3.xlsx`: planilla fuente ficticia de 1º cuatrimestre (idéntica a v0.2).
- `reference/fixtures/3er_bimestre.tsv`: planilla fuente ficticia de 3º bimestre en texto plano con tabulaciones. Reproduce la estructura de niveles colocando el marcador `EPS1` en columna A para permitir la detección del nivel visual. Contiene calificaciones válidas para el escenario *happy path*.
- `reference/fixtures/3er_bimestre_invalido.tsv`: fixture idéntico al anterior pero con una valoración no admitida fuera del glosario ("Bueno" en Lengua y Literatura I para Camila Uno), destinado exclusivamente a aislar y comprobar el manejo de datos inválidos.
- `reference/fixtures/3er_bimestre_estructura_real_EPSI.tsv`, `3er_bimestre_estructura_real_EPSG.tsv` y `1er_cuatrimestre_estructura_real_EPSI.tsv` (#10): grillas sintéticas con la forma de las grillas reales (bloques de cuatro filas, cuatro niveles, Formación Profesional en fila 2 y la cola desde `ORIENTACIÓN` con `LISTO PARA EXPORTAR`, `EXPORTADO` y columnas internas posteriores).

## Escenarios cubiertos y resultados esperados

### 1. Happy path 3B (1B + 1C + 3B acumulado)
- **Caso**: Camila Uno (DNI `99000001`, orientación Informática · EPSI), presente con valoraciones válidas en los 3 períodos.
- **Resultado esperado**:
  - Generación sin errores.
  - Finalización limpia sin advertencias de datos inválidos (`Sin advertencias` en la celda `F4` de Control y sin bloque de advertencias en el diálogo de UI).
  - Nivel visual en celda `E6` de `Informe_actual`: `Informática 1`.
  - Acumulación correcta de columnas en materias compartidas:
    - *Matemática I*: `Suficiente` (1B) | `7` (1C) | `Avanzado` (3B).
    - *FGI I*: `ACREDITÓ` (1B) | `ACREDITÓ` (1C) | `ACREDITÓ` (3B).
    - *Lengua y Literatura I*: `En proceso` (1B) | vacío (1C) | `Suficiente` (3B).

### 2. Materia que aparece por primera vez en 3B
- **Caso**: Carla Cuatro (DNI `99000003`, EPSI), cursa "Taller C", materia que aparece por primera vez para ella en el 3º bimestre (sin valoraciones en 1B ni 1C).
- **Resultado esperado**:
  - Generación sin errores.
  - Nivel visual en `E6`: `Informática 1`.
  - En la fila de *Taller C*, columnas 1B y 1C vacías, y columna 3B con `Suficiente` (`||Suficiente`).

### 3. Valoración inválida en 3B (descarte + advertencia + continuidad)
- **Caso separado**: Camila Uno (DNI `99000001`) procesada con la fuente `3er_bimestre_invalido.tsv`, donde Lengua y Literatura I contiene el valor `"Bueno"` (no admitido para bimestres).
- **Resultado esperado**:
  - **Descarte y conservación**: La columna de 3B para Lengua y Literatura I queda vacía; se conserva la valoración previa válida de 1B (`En proceso||`).
  - **Advertencia**: Se emite la advertencia específica `Valor inesperado en 3º bimestre para "Lengua y Literatura I": "Bueno"`, registrada tanto en el diálogo de UI como en la celda `F4` de Control.
  - **Continuidad**: La ejecución no se interrumpe y el informe se genera completo con las materias restantes válidas (*Matemática I* acumula `Suficiente|7|Avanzado`).

### 4. Regresión histórica de períodos anteriores (matriz v0.2)

La suite v0.3 reutiliza de forma automatizada una selección de 4 casos representativos de la matriz histórica de `legacy/v0.2/Casos_de_prueba_Generador_EPS_v0_2.xlsx`:
- **CP-01** (generación básica de 1º bimestre)
- **CP-02** (acumulación de 1º bimestre y 1º cuatrimestre)
- **CP-03** (valor no válido en cuatrimestre: descarte, advertencia y continuidad)
- **CP-06** (orientación Gastronomía / EPSG)

Como recorte proporcional y deliberado, los casos **CP-04** (estudiante ausente en fuente anterior) y **CP-05** (DNI duplicado en fuente principal) no forman parte de esta suite automatizada de #3.

- **4.1. Regresión 1º bimestre puro (CP-01)**:
  - Camila Uno (DNI `99000001`) con fuentes posteriores vacías.
  - Resultado esperado: sin advertencias, nivel visual `Informática 1`, y solo valores en columna 1B (*Matemática I*: `Suficiente||`, *Lengua*: `En proceso||`, *FGI I*: `ACREDITÓ||`).
- **4.2. Regresión 1º cuatrimestre acumulativo (CP-02)**:
  - Camila Uno (DNI `99000001`) con fuentes 1B + 1C (sin 3B).
  - Resultado esperado: sin advertencias, nivel visual `Informática 1`, acumulación de materias presentes en ambos (*Matemática I*: `Suficiente|7|`, *FGI I*: `ACREDITÓ|ACREDITÓ|`), solo en 1B (*Lengua*: `En proceso||`) y solo en 1C (*Ciencias Sociales I*: `|8|`, *Arquitectura*: `|8|`).
- **4.3. Regresión 1º cuatrimestre con valor inválido (CP-03)**:
  - Bruno Dos (DNI `99000002`) con calificación decimal `"7,5"` en 1C en Lengua y Literatura I.
  - Resultado esperado: el valor `7,5` es descartado, se conserva la calificación `Suficiente` de 1B en Lengua y Literatura I (`Suficiente||`), se emite advertencia por `7,5`, y se comprueba la continuidad operativa verificando que las materias válidas restantes se completen normalmente (*Matemática I*: `|8|`).
- **4.4. Regresión Orientación Gastronomía (CP-06)**:
  - Nora Uno (DNI `98000001`, EPSG) con fuentes 1B + 1C.
  - Resultado esperado: nivel visual en `E6` es `Gastronomía 1` y la materia *Manipulación de alimentos* acumula correctamente `Avanzado|10|`.

### 5. Estructura real de las grillas (#10)

Con los fixtures `*_estructura_real_*.tsv`:
- **Materias y cola**: las materias son los pares etiqueta/valoración entre la columna D y `ORIENTACIÓN`; desde `ORIENTACIÓN` en adelante ninguna columna es materia. Formación Profesional se lee por subencabezado de fila 2; los marcadores `EPS1`…`EPS4` y los encabezados de grupo no son materias.
- **Valores**: `EQUIVALENCIA` se imprime literalmente; `CURSA`, `NO CURSA` y celdas vacías no se imprimen.
- **Acumulación**: toda materia con una valoración válida en algún período aparece, aunque sea de otro nivel o no tenga valor en el período principal. Orden: el de la grilla principal y, al final, las materias que sólo existen en una fuente anterior.
- **Más de ocho materias**: las cajas de observaciones se desplazan sin pisar materias.
- **Estudiante nuevo** (Beto Nuevo, DNI `97000002`): sólo presente en 3B; la ausencia en 1C es una advertencia no bloqueante. Con 1B, 1C y 3B indicadas en Control y el flujo completo de `generarInformeActual()`, el informe se genera con una advertencia por cada fuente anterior y las columnas de 1B y 1C vacías. Si el período elegido es 1º cuatrimestre, 1C es la principal y el DNI ausente bloquea con un error que identifica la fuente principal y el período.
- **EPSG** (Gael Uno, DNI `96000001`): nivel visual `Gastronomía 1`.
- **Estructuras ambiguas**: falta o repetición de `ORIENTACIÓN`, materia repetida, materia sin columna de valoración y DNI fuera de la fila "Valoración" son errores en la fuente principal; en una fuente anterior, esa hoja se descarta con advertencia (sección 8).

### 6. Campos generales y operativos: sólo desde la fuente principal (#10)

**Regla**: `ORIENTACIÓN`, las observaciones generales, de asistencia y de convivencia, el PTF, `LISTO PARA EXPORTAR` y `EXPORTADO` se leen únicamente de la fuente principal del período solicitado. Las fuentes anteriores sólo aportan materias y valoraciones: no leen, validan ni advierten sobre esos campos.

Cada campo de la cola se identifica por su encabezado exacto (normalizado sin tildes, mayúsculas ni espacios extra), nunca por posición:

| Campo | Encabezado |
|---|---|
| Observaciones generales | `OBSERVACIONES GENERALES SOBRE DESEMPEÑO ACADÉMICO` |
| Asistencia | `OBSERVACIONES SOBRE ASISTENCIA` |
| Convivencia | `OBSERVACIONES SOBRE CONVIVENCIA ESCOLAR` |
| PTF | `PLAN DE TRABAJO FORMATIVO (PTF)` |
| Listo para exportar | `LISTO PARA EXPORTAR` |
| Exportado | `EXPORTADO` |

**Columnas operativas** (`LISTO PARA EXPORTAR`, `EXPORTADO`; valores canónicos `SÍ` / `NO`, estado inicial `NO`):

| Situación en la fuente principal | Estado | Advertencia |
|---|---|---|
| Columna ausente (grilla anterior a #10) | nulo | no |
| `SÍ` o `NO` | `SÍ` / `NO` | no |
| Celda vacía | nulo | sí, no bloqueante |
| `Si`, `si`, `sí`, `Sí`, `SI` | `SÍ` | sí, forma no canónica |
| `No`, `no` | `NO` | sí, forma no canónica |
| Cualquier otro valor | nulo | sí |
| Encabezado repetido | nulo | sí; no se elige columna |

Estos estados se leen y quedan en el resultado interno (`camposOperativos`, separado de `estudiante`). No alteran el boletín individual y nunca se escriben en la fuente. Sólo `LISTO PARA EXPORTAR` de la fuente principal selecciona estudiantes, y únicamente en la generación por curso (sección 7).

Casos automatizados (Ana Real, DNI `97000001`, salvo indicación):
- **Lectura**: en 3B, `LISTO PARA EXPORTAR = SÍ` y `EXPORTADO = NO` quedan en `camposOperativos`, sin advertencias. Con la grilla v0.3 de 3B (sin columnas operativas) ambos quedan nulos, sin advertencia.
- **Validación**: variantes no canónicas, celda vacía, valor inválido (`Tal vez`) y encabezado repetido producen la advertencia correspondiente; el informe generado es idéntico al del caso válido.
- **Sin efecto sobre la generación**: con `LISTO PARA EXPORTAR = NO` el boletín se genera igual; Beto Nuevo (`NO` / `NO`) se genera normalmente.
- **Nunca materias**: `LISTO PARA EXPORTAR` y `EXPORTADO` no aparecen entre las materias; si uno de esos encabezados aparece antes de `ORIENTACIÓN`, se informa un error en vez de leerlo como materia.
- **Fuente histórica**: una 1C con orientación, observaciones y operativos distintos o inválidos, y encabezados de cola faltantes o repetidos, no cambia el informe ni produce advertencias. El resultado de una fuente histórica sólo contiene materias.
- **Generación de 3B**: orientación, observaciones y PTF salen de 3B; si un campo está vacío en 3B, queda vacío y no se completa con 1C.
- **Generación de 1C**: con 1C como principal, esos mismos campos y `camposOperativos` (`SÍ` / `SÍ`) salen de 1C, y no se lee 3B.
- **Encabezado general no identificable en la principal**: un encabezado distinto del esperado (por ejemplo `ASISTENCIA`) o vacío —forma observada en la cola de una copia anonimizada de EPSG— produce una advertencia por campo; el informe se genera con esa caja vacía, sin leer la columna por posición ni desde otra fuente.

El mock de fuentes del arnés sólo expone lectura (`getDataRange().getDisplayValues()`): cualquier intento de escribir en una grilla fuente hace fallar la suite. Además registra cada apertura, cada lectura y cualquier acceso a algo distinto de esa lectura, aunque el generador capture el error.

### 7. Generación por curso (#12)

**Uso**: en `Control`, además del período, la hoja origen, el ciclo lectivo y las fuentes, se elige el nivel en `B9` (*Nivel / curso*, selector `EPS1`…`EPS4` incluido en el workbook versionado; también se acepta sólo el número). *Informes EPS > Generar boletines por curso* crea un libro nuevo `Boletines EPS - <hoja origen> - <nivel> - <período> <ciclo> - <aaaa-mm-dd hh:mm>` con la hoja `Resumen` y una hoja por estudiante generado. La hora es siempre la de `America/Argentina/Buenos_Aires`, sin depender de la zona horaria del libro generador, del proyecto ni de la cuenta. El libro se guarda en `Mi unidad > Boletines EPS > <ciclo> - <período>` (sección 14). El DNI de `B7` no se usa. La generación individual ignora `B9` y no cambia.

**Selección**: sólo desde la fuente principal ya leída. El nivel de cada fila es el último marcador `EPSn` de la columna A. Se selecciona una fila del nivel con DNI cuyo `LISTO PARA EXPORTAR` resulta `SÍ` según la tabla de la sección 6 (una forma no canónica como `Si` cuenta como `SÍ` y queda advertida). `NO`, vacío o inválido se omite con su motivo. `EXPORTADO` se informa en el `Resumen` pero no selecciona. Una fila del nivel con `LISTO PARA EXPORTAR = SÍ` y sin DNI se informa como error.

**Lectura única**: cada fuente se abre, se lee y se valida una vez por tanda (`leerFuentes_`). Después, para cada estudiante, se construyen los datos con la misma función que usa la generación individual (`construirDatosEstudiante_`) y cada hoja se completa con la misma lógica que `Informe_actual` (`completarInforme_` sobre una copia de `Plantilla`).

**Errores**:
- Globales, sin libro: configuración incompleta, error estructural de la fuente principal, falta de una única columna `LISTO PARA EXPORTAR` después de `ORIENTACIÓN` en la fuente principal (no se usa `LISTO PARA IMPRIMIR` ni ninguna otra columna en su lugar) y nivel sin marcador en la grilla. El mensaje nombra el período elegido en `Control!B5`, porque ese período decide cuál fuente es la principal. Las fuentes anteriores no necesitan columnas operativas y nunca se inspeccionan.
- Por estudiante: su error (por ejemplo, DNI repetido) queda en el `Resumen` con el motivo; el resto se genera.
- Si ningún estudiante puede generarse, no se crea el libro y el detalle queda en el diálogo y en `Control`.

**Resumen**: advertencias generales una sola vez (falta de una fuente anterior, encabezados de la cola) y, por estudiante del nivel en orden de grilla: estado (`Generado`, `Omitido`, `Error`), hoja, nombre, DNI, fila en la fuente, `EXPORTADO` y motivo o advertencias propias. Si varios estudiantes producirían el mismo nombre de hoja (sin distinguir mayúsculas), todos llevan su DNI entre paréntesis. Al terminar, `Resumen` se oculta para que no entre en la impresión ni en el PDF del libro completo. Si no se generó ningún boletín, queda visible porque Google Sheets exige una hoja visible.

Casos automatizados (grillas reales sintéticas con bloques agregados en el test):
- Curso EPS2 con `SÍ`, `NO` (con `EXPORTADO = SÍ`), vacío, `Tal vez`, `SÍ` con `EXPORTADO = SÍ`, `Si` y dos estudiantes con el mismo nombre. En 1C (fuente anterior), otros estados operativos y un estudiante listo que no existe en 3B, sin efecto.
- Un estudiante de EPS1 listo no entra en EPS2; con nivel `1` se genera sólo él.
- Equivalencia: cada hoja es idéntica al `Informe_actual` individual del mismo estudiante, y sus advertencias (generales + propias) coinciden con las de la generación individual.
- Cada fuente se abre y se lee una sola vez; no hay accesos fuera de la lectura de valores; las fuentes, `Plantilla` y las hojas del generador (salvo el estado en `Control`) quedan intactas.
- DNI repetido en la principal: error de ese estudiante, el resto se genera.
- Sin columna `LISTO PARA EXPORTAR` (renombrada `LISTO PARA IMPRIMIR`, repetida o ausente en la grilla v0.3), nivel inexistente o sin nivel: error claro y ningún libro.
- Sin estudiantes listos: ningún libro; el diálogo y `Control` lo explican.
- Menú: conserva *Generar informe actual* y agrega *Generar boletines por curso*.
- Hora del nombre (sección 13; en el mock, la zona del libro generador es `America/Los_Angeles`): 21:07 UTC queda como `18:07` (no `14:07`), y 02:30 UTC del día 15 queda como `2026-10-14 23:30`. La zona se pasa explícitamente a `Utilities.formatDate`. El diálogo muestra el nombre y el enlace del libro; sin libro o en la generación individual no se crea ninguno.
- `Resumen` oculta y boletines visibles, también con un estudiante fallido. Si ninguna hoja pudo crearse, `Resumen` queda visible con los errores (el mock rechaza ocultar la última hoja visible, como Google Sheets).
- `Control` del xlsx versionado: `A9` = *Nivel / curso*, `B9` vacío con selector `EPS1,EPS2,EPS3,EPS4`. Con esa hoja, la generación por curso usa el nivel elegido, pide `B9` si no se eligió, y la generación individual lo ignora (aunque tenga un valor no válido).
- Coordenadas de `Control`: las fuentes de 1º bimestre, 1º cuatrimestre y 3º bimestre están en `B10`, `B11` y `B12`, sin fila vacía intermedia. Los tests toman todas las coordenadas de `CONFIG_EPS` y comprueban en el xlsx que cada una esté junto a su rótulo, junto con los valores por defecto y las validaciones de `B5`, `B6` y `B9`.
- Columnas operativas sólo en la principal: en 3B, 1C y 1B sin columnas operativas son válidas. En 1C, 1B sin ellas es válida. En 1B, la principal 1B las exige. Una histórica con `LISTO PARA EXPORTAR` duplicada, `EXPORTADO` inválido o encabezados operativos malformados produce la misma tanda que una histórica limpia. La misma grilla 1C sin `LISTO PARA EXPORTAR` es válida como histórica de 3B y es un error sólo como principal (período 1º cuatrimestre). La principal con la columna duplicada también es un error.

**Cambio en la generación individual**: el contenido del informe no cambia. Las advertencias de las fuentes (falta de una fuente anterior, fuente anterior ilegible, encabezados de la cola) ahora aparecen antes que las del estudiante.

**No cubierto por el arnés** (queda para la validación manual de #11): copia real de `Plantilla` a otro libro (imágenes, celdas combinadas, formato condicional, anchos y altos), permisos de creación en Drive (y de carpetas, sección 14), tiempos y límites de Apps Script con un curso completo, y PDF e impresión del libro.

### 8. Orientación por período (#15)

**Regla**: la orientación es del período, no del estudiante. `Control!B6` elige la hoja de la fuente principal; de ella salen la selección por curso, los datos generales, la orientación mostrada y `LISTO PARA EXPORTAR` / `EXPORTADO`. En cada fuente anterior, el DNI se busca en `EPSI` y en `EPSG`, empezando por la hoja de `B6`. Individual y por curso comparten este comportamiento.

| DNI en las hojas de un período | Fuente principal | Fuente anterior |
|---|---|---|
| Ninguna | Error (como antes) | Advertencia y período vacío (como antes) |
| Una | Se usa la hoja de `B6` | Se usa esa hoja; si no es la de `B6`, advertencia informativa *"El 1º cuatrimestre se tomó de la hoja EPSI de su fuente, distinta de la orientación actual (EPSG)."* |
| Ambas | Error del estudiante: *"El DNI aparece en las hojas EPSG y EPSI de la fuente de 3º bimestre. No se puede decidir en qué orientación cursa este período."* | Error del estudiante: no se elige ni se combina |

En la generación por curso, el error sólo afecta a ese estudiante; los demás se generan. Un DNI en ambas hojas de la principal no se genera en ninguna de las dos tandas.

**Lectura acotada**: cada archivo se abre una vez y cada hoja de orientación se lee una vez por tanda. La otra hoja de la principal sólo se usa para detectar DNI repetidos entre orientaciones. Una hoja que falta en una fuente anterior no se advierte; si faltan todas, sí. Una hoja ilegible (estructura no identificable) se advierte una vez y no se usa, pero se sigue consultando su columna de DNI para no elegir otra hoja en silencio.

Casos automatizados (trayectorias sintéticas armadas en el test sobre las grillas reales de EPSI y EPSG, con tandas EPSG y EPSI de EPS1):
- **A** (1B EPSI, 1C EPSI, 3B EPSG): conserva 1B y 1C, incluidas las materias que sólo existen en EPSI (al final), con una advertencia por período y la orientación actual en el boletín.
- **B** (1B EPSI, 1C EPSG, 3B EPSG): 1C desde la hoja actual y 1B desde EPSI.
- **C** (sólo en 3B): se genera con la advertencia de ausencia de siempre.
- **D** (1C en EPSI y EPSG) y **E** (3B en EPSI y EPSG): error del estudiante, en la tanda y en la generación individual; en E, en ninguna de las dos tandas.
- **F** (EPSI → EPSG → EPSI) y EPSG → EPSI: cada período sale de su hoja.
- Mismo informe y mismas advertencias en individual y por curso; una apertura por archivo y una lectura por hoja; sin accesos de escritura.
- Hoja ilegible en una fuente anterior: advertencia única, se usa la otra hoja, y un DNI en ambas sigue siendo error.
- `Control` del xlsx: `C6` y la regla de hoja origen describen la orientación del período a generar y la búsqueda en ambas hojas.

### 9. Impresión de boletines de más de una página

**Contexto**: la `Plantilla` está diseñada para una página con margen superior 0: los logos están anclados al borde superior de la fila 1. Un boletín con muchas materias (con la configuración de impresión de abajo, unas 14 o más) pasa a una segunda página. Google Sheets aplica el mismo margen superior a todas las páginas de la hoja, no parte filas y no permite fijar saltos de página por script, así que la segunda página empezaba pegada al borde.

**Regla**: cada boletín (`Informe_actual` y cada hoja del libro por curso, en `completarInforme_`) tiene la fila 1 inmovilizada. Esa fila sólo contiene los logos, sin texto, bordes ni relleno, y mide 39,75 pt (≈1 cm impresa con "ajustar al ancho"). Al imprimir con *Repetir filas inmovilizadas*, la primera página no cambia y cada página siguiente empieza con esa fila vacía como aire superior. Un boletín de una página se imprime igual que antes. No se calcula dónde cae el salto, así que funciona con cualquier cantidad de materias. `Plantilla` y `Resumen` no se modifican.

**Configuración de impresión**: A4 vertical, ajustar al ancho, márgenes superior y laterales de 0 cm e inferior de 1 cm, y *Repetir filas inmovilizadas* activada.

Casos automatizados (sección 14 de `tests/run.js`; el paginado real de Sheets no se modela):
- Boletín corto (3 materias): fila 1 inmovilizada y ninguna fila más; el PTF queda en la fila esperada.
- Boletín largo (17 materias): conserva todas las materias y los cuatro bloques inferiores, con la fila 1 inmovilizada; la hoja del libro por curso es idéntica al informe individual; sin escrituras en las fuentes.
- Tanda: todos los boletines con la fila 1 inmovilizada; `Resumen` y `Plantilla` sin cambios.

**Validación manual** (#11): queda pendiente comprobar en Google Sheets que, con la opción activada, la segunda página tenga el aire esperado sin repetir los logos y sin agregar una tercera página.

### 10. Color de las valoraciones: identidad de la escala conceptual

**Origen**: los modelos institucionales "Informe Valoración + PTF" de abril de 2024 (horizontal y vertical) dan a cada nivel de la escala conceptual de los bimestres un color propio: En Proceso, Suficiente y Avanzado. Las notas numéricas aparecen en un gris neutro, sin color según aprobación. No hay texto que defina los colores; la leyenda sólo define los niveles y la escala de 1 a 10 con mínimo 6.

**Regla**: en el área de valoraciones (filas de materias, columnas de períodos de `E` a `L`), cada nivel conceptual tiene identidad propia, con relleno claro y texto oscuro del mismo tono. Lo aplica `completarInforme_` en cada generación, así que vale igual para el informe individual y para el libro por curso:

| Valoración | Tratamiento |
|---|---|
| `En proceso` | tono rojo claro, texto rojo oscuro (versión accesible del rojo saturado de 2024) |
| `Suficiente` | tono verde claro, texto verde oscuro (como en 2024) |
| `Avanzado` | tono celeste, texto azul oscuro (como en 2024) |
| Notas numéricas (cualquier valor), `ACREDITÓ`, `EQUIVALENCIA` | dato neutro: sin relleno, texto negro |
| Celda vacía de un período habilitado | gris muy tenue: "no hay contenido en esta celda", sin juzgar por qué |
| Período posterior al generado (aunque la celda esté vacía) | gris: "todavía no corresponde en este informe" |

- Vacía y futuro son estados estructurales, no académicos. Dependen sólo de si el generador escribió un valor y de la posición del período, nunca del valor. Son grises sin tono, de menos a más marcados (dato neutro < vacía < futuro), y futuro tiene precedencia sobre vacía. Una celda queda vacía cuando el generador no escribe nada en ella, por ejemplo con `CURSA`, `NO CURSA` o un valor descartado.
- Los tonos están en `CONFIG_EPS.SENALES`, `COLOR_CELDA_VACIA` y `COLOR_PERIODO_FUTURO`. Los rellenos son claros (luminancia de 0,6 o más) y el texto tiene contraste de 4,5:1 o más. No se usa negrita: el color acompaña a la palabra, que siempre está.
- No se colorean las notas según menos de 6 o 6 y más: los modelos no lo hacen y la escuela no lo definió.
- `ACREDITÓ` y `EQUIVALENCIA` no aparecen en los modelos. Quedan neutras hasta que se decida otra cosa.
- Trayectorias, de izquierda a derecha:
  - `En proceso → 4 → Suficiente`: el rojo pasa a verde, y la nota queda neutra;
  - `Avanzado → 9 → Avanzado`: celeste, neutro, celeste.
- Columnas en gris según el período generado: con 1B, de `F` a `L`; con 1C, de `G` a `L`; con 3B, de `H` a `L`. Las columnas `H` a `L` no se completan en esta versión.
- El color se calcula sobre todas las filas reales de materias, también las agregadas cuando hay más de 8, sin formato condicional. Sustituye al semáforo de formato condicional de v0.4 (los mismos tres tonos, saturados, sólo en las 8 filas de la `Plantilla`). Esas 6 reglas se quitaron de la `Plantilla`, y el boletín generado queda sin formato condicional aunque un generador anterior todavía las tenga.
- `Plantilla` sin generar: el área de valoraciones (`E9:L16`) tiene una base uniforme, sin relleno. Se quitó el gris fijo de las columnas `F` y `H` a `L`, heredado del gris neutro de las notas del modelo de 2024, que ya no representaba ninguna regla. Encabezado y nombres de materia conservan su relleno estructural. Vacía y futuro aparecen sólo al generar.
- Leyenda: la imagen institucional ya define En Proceso, Suficiente y Avanzado, y el color va pegado a esas palabras. No se agrega un texto sobre colores; `D17` queda vacía, como en v0.4.

Casos automatizados (sección 15 de `tests/run.js`):
- Gramática: sólo los tres niveles tienen color; tono propio por nivel (rojo, verde y azul), con texto del mismo tono; gris neutro para lo futuro; rellenos claros y contraste suficiente.
- Estados estructurales: vacía y futuro son grises sin tono, distintos entre sí y de toda valoración, con orden perceptible en grises.
- Trayectoria de 16 materias en 3B: mejora, en proceso sostenido, Avanzado sostenido, `Suficiente → 6 → Avanzado`, retroceso, `ACREDITÓ` y `EQUIVALENCIA` como datos neutros, notas de 1C neutras, 1B y 1C vacías en materias que sólo tienen 3B, la misma gramática en filas agregadas y sólo el área de valoraciones formateada.
- Futuro según el período, con precedencia sobre vacía: 3B (16 materias), 1C (8, con una celda vacía en 1B) y 1B (7).
- Individual y por curso con los mismos fondos y colores de texto; sin el semáforo heredado; sin escrituras en las fuentes.
- Categoría de cada valor.
- `Plantilla` del xlsx sin formato condicional, con base uniforme en `E9:L16`, con encabezado y nombres de materia intactos y con `D17` vacía, y la fila 17 en 192 pt.

### 11. Varios cursos en una sola operación (#20)

**Uso**: `Control` tiene el bloque *Varios cursos* (filas 14 a 22) con los 8 cursos fijos de CFP7 (`EPSI` y `EPSG`, `EPS1` a `EPS4`, en `CONFIG_EPS.CURSOS`) y una marca `SÍ` / `NO` por curso en `B15:B22`. *Informes EPS > Generar boletines de los cursos marcados* toma los cursos con `SÍ` (también `Si`); `NO` o vacío no lo incluye. *Informes EPS > Generar boletines de todos los cursos* toma los 8 y no mira las marcas. Ambas usan el período, el ciclo lectivo y las fuentes de `Control`, e ignoran `B6`, `B7` y `B9`. No hay descubrimiento de cursos.

**Reutilización**: cada curso es una tanda de la generación por curso. `generarTandaCurso_` quedó dividida en `prepararTandaCurso_` (selección y datos de cada estudiante, sin crear nada) y `crearLibroTanda_` (libro, hojas y `Resumen`). La generación por curso llama a las dos seguidas; la de varios cursos primero prepara todos los cursos elegidos y después crea los libros en orden fijo.

**Lectura**: las fuentes se leen una vez por orientación (`leerFuentes_`) y las comparten sus cursos. Cada tanda recibe su propia copia de las advertencias generales.

**Errores**:
- Configuración (ningún curso marcado, una marca distinta de `SÍ`, `NO` o vacío, falta de período o de fuente principal) y estructura (fuente principal ilegible, falta de `LISTO PARA EXPORTAR`, nivel sin marcador en la hoja de su orientación): abortan antes de crear el primer libro, aunque el curso con problemas sea el último.
- Por estudiante: como en la generación por curso; no afecta a los demás estudiantes ni a otros cursos.
- Curso sin estudiantes listos: no genera libro y los demás se procesan.
- Falla inesperada al crear libros (por ejemplo, de Drive): no se borra lo creado; `Control` y el diálogo listan los libros ya creados, que ya están en su carpeta (sección 14).

**Resultado**: `Control!F3` muestra `Cursos: n · Libros creados: n · Sin libro: n · Duración: n s.`. `F4` muestra una línea por curso (libro con conteos y enlace, o motivo sin libro), las advertencias generales de cada curso y, para los cursos sin libro, el detalle por estudiante. Un aviso (`toast`) indica el curso en curso. No hay jobs, colas, triggers ni cortes por tiempo: el rendimiento de los 8 cursos se mide en una prueba real.

Casos automatizados (sección 16 de `tests/run.js`, sobre las grillas reales sintéticas con los 8 niveles en EPSI y EPSG; 3 cursos sin estudiantes listos):
- Todos: sin marcas, 5 libros en orden fijo y ninguno para los 3 cursos sin estudiantes listos. Cada libro es idéntico al de la generación por curso del mismo curso (hojas, `Resumen`, valores, colores, visibilidad y fila inmovilizada). `Control`, diálogo y avisos con una línea por curso. Una apertura por archivo y orientación. Sin escrituras en las fuentes. Ignora las marcas, aun con valores no válidos.
- Cursos marcados: sólo los marcados, en orden fijo, y el marcado sin estudiantes listos no crea libro. Con un solo curso, el libro es el de la generación por curso. Con una sola orientación, no se lee la otra como principal.
- Sin cursos marcados, una marca no admitida o falta de la fuente principal: error claro, sin leer fuentes (en las marcas) ni crear libros.
- Error global en el último curso (EPSG sin marcador `EPS4`) o en una orientación (EPSG sin `LISTO PARA EXPORTAR`): ningún libro.
- Un estudiante en las hojas principales de EPSI y EPSG: error en los dos cursos; su curso queda sin libro, el otro se genera con el resto y los demás libros no cambian.
- Falla al crear el segundo libro: el error se propaga y se informa el libro ya creado.
- La generación individual y la generación por curso ignoran las marcas.
- `Control` del xlsx: encabezado `A14`, rótulo de cada curso junto a su marca vacía, validación `SÍ,NO` en `B15:B22`, manual y reglas corridos debajo del bloque. Con ese `Control`, se generan sólo los cursos marcados.

**No cubierto por el arnés**: tiempo real y límites de Apps Script con los 8 cursos (límite de 6 minutos por ejecución), cuotas de creación en Drive, aspecto del bloque y de los selectores en Google Sheets, y el aviso de progreso.

### 12. Textos que no deben cortarse: nombre del estudiante y EQUIVALENCIA

**Nombre**: en la `Plantilla`, el nombre del estudiante (`E4`) era la celda combinada `E4:F4` (≈184 px). Google Sheets no deja que el texto de una celda combinada desborde a las vecinas, así que un nombre largo en Montserrat 11 se cortaba (por ejemplo, *"Robledo Mencacci, Alejo Lucas"* ocupa ≈226 px). Ahora la combinación es `E4:L4`, sobre celdas vacías: el nombre llega hasta el borde derecho del boletín y no cambian el DNI (`E5:F5`), el curso (`E6:F6`), los anchos de columna, los altos de fila ni la impresión.

**EQUIVALENCIA**: el área de valoraciones usa Montserrat 9 con ajuste de texto. `EQUIVALENCIA` es una sola palabra de ≈94 px y una columna de período tiene ≈80 px útiles (la más angosta, `G`), así que Sheets la partía por letras (`EQUIVALENCI` / `A`). `completarInforme_` le asigna tamaño 7 (`CONFIG_EPS.TAMANO_EQUIVALENCIA`, ≈73 px) sólo a las celdas con ese valor. El resto del área queda en 9, y las otras valoraciones válidas ya entran enteras en 9. Mediciones: anchos de avance de Montserrat 400; anchos de columna estimados desde el xlsx (≈7 px por unidad).

Casos automatizados (sección 17 de `tests/run.js`; el ancho real del texto no se modela):
- `Plantilla` (xlsx): `E4:L4` combinada sobre celdas vacías, sin `E4:F4`, y `E5:F5` / `E6:F6` sin cambios.
- `EQUIVALENCIA` en 7 en el informe individual, en el libro por curso y en el de varios cursos; ninguna otra celda cambia de tamaño. Un boletín sin `EQUIVALENCIA` no cambia ningún tamaño.

**Validación manual**: comprobar en Google Sheets, y en el PDF, que un nombre largo se vea completo y que `EQUIVALENCIA` quede en una línea en las columnas de 1º bimestre, 1º cuatrimestre y 3º bimestre.

### 13. Exportación a PDF de un libro ya generado (#21)

**Uso**: *Informes EPS > Exportar PDF desde un libro o carpeta…* abre un diálogo (sección 17) donde se pega el link o ID de un libro de la generación por curso (o de varios cursos), ya revisado, o de una carpeta (sección 16). Crea una carpeta nueva junto al libro, `<nombre del libro> - PDF <fecha y hora>`, con un PDF por hoja de estudiante (todas menos `Resumen`), con el nombre de la hoja. Si no se puede ver la carpeta del libro, la crea en Mi unidad. Si la última carpeta de PDF del libro quedó incompleta, la vista previa ofrece completarla (sección *Recuperación*).

**Separación**: es una operación posterior y separada de la generación, en su propio archivo (`src/Exportar_PDF.gs`). `Generador_EPS.gs` sólo la nombra en el menú y no usa sus servicios (`UrlFetchApp`, `ScriptApp`); usa `DriveApp` únicamente para guardar sus libros (sección 14). Como los libros quedan en la carpeta de su ciclo y período, las carpetas de PDF quedan ahí, junto al libro. El libro sólo se lee: una falla de exportación nunca obliga a regenerarlo. Cada exportación completa crea su carpeta, así que volver a exportar todo no reemplaza ni mezcla PDF anteriores; sólo completar una exportación incompleta agrega PDF a una carpeta existente.

**Mecanismo**: la URL de exportación de Google Sheets (`/export?format=pdf&gid=…`), pedida con `UrlFetchApp` y el token del usuario. No está documentada oficialmente, por eso queda aislada en `urlExportacionPdf_` y `exportarHojaPdf_`. Cada pedido lleva la configuración de impresión completa (`CONFIG_PDF.PARAMETROS`), sin depender de lo guardado en el diálogo de impresión: `size=A4`, `portrait=true`, `fitw=true` y `scale=2` (ajustar al ancho), `top_margin`, `left_margin` y `right_margin` en 0, `bottom_margin=0.3937` (1 cm en pulgadas) y `fzr=true` (repetir filas inmovilizadas, la fila 1 de cada boletín; ver sección 9), sin cuadrícula, notas, título, nombre de hoja ni números de página.

**Errores**:
- Link no reconocible, libro sin `Resumen` (por ejemplo, el propio generador) o sin hojas de estudiantes: error claro, sin crear la carpeta.
- Generación que no terminó: `Resumen` se escribe al final de la generación, en una sola operación y después de copiar todas las hojas, así que su primera celda (`CONFIG_EPS.TITULO_RESUMEN`) indica que el libro está completo. Un libro con `Resumen` vacío, de una generación cortada (por ejemplo, por tiempo), se rechaza: sus hojas no son el curso completo. Vale también para los libros generados antes de este cambio.
- Límite de Google (HTTP 429): en la prueba real con un curso de 15 hojas, los pedidos seguidos dieron 6 PDF y 9 rechazos 429 en 24 s. Ahora entre un pedido y el siguiente se espera una pausa que empieza en 3 s (`CONFIG_PDF.PAUSA_MS`). Ante un 429 se duplica, hasta 30 s (`PAUSA_MAXIMA_MS`), y se reintenta la misma hoja, hasta 4 pedidos en total (`INTENTOS_POR_HOJA`). Cada PDF exportado la vuelve a bajar 3 s. Google no publica el límite: los valores son iniciales y se ajustan con los reintentos y tiempos que registra `Control`.
- Por hoja: el 429 en el último intento, otra respuesta HTTP distinta de 200, una respuesta que no es PDF o una excepción del pedido se registran en esa hoja, y las demás se exportan igual. Sólo el 429 se reintenta.
- Tiempo: pasados 5 minutos (`CONFIG_PDF.TIEMPO_MAXIMO_MS`) desde el inicio de la ejecución que exporta, no se empieza otra exportación ni otra pausa, y las hojas que faltan se informan como sin exportar. Desde el diálogo, esa ejecución empieza al elegir *Exportar*: el tiempo de buscar el link y de revisar la vista previa no cuenta, porque no corre ninguna ejecución mientras tanto (sección 17). En las exportaciones desde `Control`, la confirmación con `ui.alert` sí ocurre dentro de la ejecución y cuenta. Sin margen para el primer pedido no se crea una carpeta vacía.

**Recuperación**: cada carpeta de PDF lleva en su descripción de Drive el ID del libro exportado (`PDF del libro de boletines <ID>`). Al exportar un libro, se busca junto a él su última carpeta de PDF: la creada más recientemente, fuera de la papelera, con el ID de ese libro en la descripción. No se usa el nombre, porque dos libros distintos pueden llamarse igual (dos regeneraciones del mismo curso en el mismo minuto). Una carpeta sin esa descripción, como las de exportaciones anteriores a este cambio, no se ofrece para completar. Si le faltan PDF de alguna hoja (los de la papelera cuentan como faltantes), la vista previa ofrece dos botones: *Completar* pide sólo las hojas que faltan y las agrega a esa carpeta; *Exportar todo en una carpeta nueva* rehace todo, por ejemplo para un libro corregido. Cerrar el diálogo no hace nada. Si la última carpeta está completa, o no hay ninguna, el único botón exporta todo en una carpeta nueva. Tampoco se ofrece completar si el libro se modificó en Drive (`getLastUpdated`) después de crearse esa carpeta: completarla mezclaría PDF de dos versiones del libro, así que se exporta todo en una carpeta nueva. La exportación vuelve a comprobar todo esto al empezar (sección 17).

**Resultado**: `Control!F3` muestra `PDF exportados: n de n · Con error: n · Sin exportar: n · Reintentos por HTTP 429: n · Tiempo de exportación: n s (n s por PDF). La carpeta tiene n de n PDF.`, el libro y el enlace a la carpeta. Al completar una carpeta, el primer total son las hojas que faltaban (`n de n que faltaban`). Si la carpeta queda incompleta, el diálogo explica cómo completarla. `F4` lista cada hoja con error o sin exportar, con el motivo. Un aviso (`toast`) indica la hoja en curso.

Casos automatizados (sección 18 de `tests/run.js`, sobre el libro EPS2 de la generación por curso):
- Libro completo: carpeta junto al libro, 4 PDF en el orden del libro y sin `Resumen`, un pedido por hoja con su `gid`, configuración de impresión validada e idéntica en todos los pedidos, token del usuario y `muteHttpExceptions`.
- El libro no cambia (hojas, valores, formatos, visibilidad, filas inmovilizadas) y sólo se usan métodos de lectura.
- Pausa de 3 s entre pedidos, el primero sin espera; sin carpeta anterior incompleta, no se pregunta nada.
- Fallas por hoja (429 en los 4 intentos, respuesta HTML y excepción): las demás se exportan, `Control` detalla cada una con los intentos y el libro no cambia. La pausa se duplica con cada 429 (6, 12, 24 s) y queda para las hojas siguientes. Un 429 seguido de éxito exporta la hoja, y cada PDF baja la pausa 3 s, sin bajar de 3.
- Tiempo máximo: con 3 minutos por PDF se exportan 2 de 4 y las otras 2 quedan informadas. Con 4 minutos hasta analizar, 10 hasta exportar y 30 s por PDF, se exportan los 4: la espera en el diálogo no cuenta (sección 17).
- Límite de Google, con un modelo supuesto (hasta 6 pedidos aceptados en cualquier ventana de 60 s) y un libro sintético de 15 hojas: sin pausa ni reintentos, como antes, reproduce la prueba real (6 de 15, 9 rechazos 429, 24 s); con la pausa adaptativa se exportan los 15 dentro del tiempo máximo.
- Recuperación: con una carpeta anterior incompleta, la vista previa muestra cuántos PDF tiene y ofrece *Completar* o *Exportar todo en una carpeta nueva*. *Completar* pide sólo las hojas que faltan y las agrega a esa carpeta; *Exportar todo* exporta todo en una nueva sin tocar la anterior; cerrar no hace nada. Sólo cuenta la última carpeta del libro fuera de la papelera, y un PDF en la papelera cuenta como faltante. El libro se reconoce por el ID en la descripción: una carpeta renombrada se ofrece y una sin descripción, no. Regresión: con dos libros distintos, de IDs distintos y exactamente el mismo nombre en la misma carpeta, la exportación incompleta de uno no se ofrece ni se completa al exportar el otro, que crea su propia carpeta; al exportar el libro al que pertenece, sí se ofrece. Si la última está completa, se exporta todo en una carpeta nueva sin preguntar.
- Recuperación de punta a punta con un límite más estricto (3 por minuto): la primera corrida queda incompleta por tiempo y lo informa; la segunda pide sólo las hojas que faltaban y completa los 15 PDF.
- Libro sin `Resumen`, libro sin hojas de estudiantes, cancelación y link inválido: no se crea carpeta ni se exporta.
- `Generador_EPS.gs` no usa `UrlFetchApp` ni `ScriptApp`, y de `DriveApp` sólo usa `getRootFolder` y `getFileById` (sección 14).

**No cubierto por el arnés**: el PDF que produce Google (tamaño, márgenes, ajuste al ancho, repetición de la fila 1, cortes de página), la vigencia de los parámetros de la URL, los permisos y cuotas reales, los tiempos y el límite real de la URL de exportación (el modelo de 6 por minuto sólo reproduce lo observado; la pausa adaptativa debe probarse con un curso real), y el comportamiento real de `getFolders`, `getFiles`, `getDateCreated`, `setDescription` y `getDescription` en la carpeta del libro. Queda pendiente la validación en Google real que pide #21: un boletín de una página, uno de dos páginas, un curso completo, independencia de la configuración guardada en el diálogo de impresión y tiempos reales.

### 14. Organización de los libros en Drive (#23)

**Uso**: la generación por curso y la de varios cursos guardan cada libro en `Mi unidad > Boletines EPS > <ciclo> - <período>`, con el ciclo lectivo y el período de `Control` (por ejemplo, `2026 - 3º bimestre`). Es una carpeta por entrega: ahí quedan los libros de todos los cursos de ese período, las regeneraciones (cada libro conserva la fecha y hora en el nombre) y, como la exportación a PDF crea su carpeta junto al libro, también los PDF. Las carpetas se crean la primera vez; después se reutilizan las existentes con ese nombre, salvo las que están en la papelera. El diálogo final indica la carpeta; `Control` conserva el enlace de cada libro. La generación individual no crea libros ni usa Drive.

**Mecanismo**: `SpreadsheetApp.create()` crea el libro en Mi unidad y `File.moveTo()` lo mueve a la carpeta. La carpeta se busca o se crea una sola vez por operación, después de validar todo y antes del primer libro, y sólo si se va a crear alguno (`carpetaLibros_`). Cada libro se mueve recién creado, antes de copiar los boletines (`moverLibro_`): si la ejecución se corta después (por ejemplo, por el límite de 6 minutos), el libro incompleto igual queda en la carpeta y no suelto en Mi unidad.

**Errores**:
- Si no se puede buscar o crear la carpeta: error en `Control` y en el diálogo, sin crear ningún libro.
- Si no se puede mover un libro: la operación se detiene como ante cualquier otra falla de Drive. El error nombra la carpeta y el libro recién creado, que quedó vacío en Mi unidad, con su enlace para borrarlo; en varios cursos, además se listan los libros ya creados, que ya están en la carpeta.

Casos automatizados (sección 19 de `tests/run.js`; Drive simulado con carpetas, papelera y `moveTo`):
- Por curso: el libro queda en `Mi unidad > Boletines EPS > 2026 - 3º bimestre`, y las dos carpetas se crean al no existir. Se mueve con su única hoja inicial, antes de copiar los boletines. `Control` no cambia y el diálogo muestra la carpeta.
- Carpetas existentes: se reutilizan sin crear otras; una con el mismo nombre en la papelera se ignora. Otro ciclo lectivo (`2027`) u otro período (`1º cuatrimestre`) usan su propia subcarpeta dentro de `Boletines EPS`.
- Varios cursos: los 5 libros en la misma carpeta, con una sola búsqueda; el diálogo indica la carpeta después de la línea de cada curso. Si falla la creación del segundo libro, el primero ya está en la carpeta.
- Sin estudiantes listos, errores globales o de configuración, cursos marcados sin libro y generación individual: no se usa Drive ni se crea ninguna carpeta.
- Fallas: sin poder crear la carpeta, ningún libro; sin poder mover el libro, error con la carpeta y el enlace del libro vacío.
- El libro organizado es idéntico al de la generación por curso.

**No cubierto por el arnés**: el comportamiento real de Drive. En particular: si `getFoldersByName` devuelve carpetas de la papelera, el orden en que devuelve carpetas repetidas, el tiempo que agrega mover cada libro, y que la exportación a PDF cree su carpeta en la carpeta del período. Esto se valida en Google real.

### 15. Exportación a PDF de varios cursos (#23)

**Problema**: con los tiempos reales (un curso de 15 hojas: 104 s, 6,9 s por PDF, con pausas y reintentos por el límite de Google), los 8 cursos llevan unos 14 minutos, más que una ejecución de Apps Script. La exportación de un libro por vez obligaba a buscar, copiar y pegar el link de cada uno.

**Uso**: *Informes EPS > Exportar PDF de los cursos marcados* toma los cursos con `SÍ` en `B15:B22`, las mismas marcas de la generación de varios cursos. *Exportar PDF de todos los cursos* toma los 8. Ambas usan el período (`B5`) y el ciclo lectivo (`B8`) de `Control`, pero no las fuentes, porque no leen las grillas.

**Selección del libro**: para cada curso, se toma el libro más reciente (`getDateCreated`) de la carpeta del período (`Mi unidad > Boletines EPS > <ciclo> - <período>`). Los libros se reconocen y agrupan como en una carpeta elegida a mano (sección 16, `librosPorCursoEnCarpeta_`): por el curso que la generación escribe en `Resumen` y, si la generación no terminó, por el nombre que le dio. Si hay varios libros del curso (regeneraciones, aun con el mismo nombre), se usa sólo el más reciente, y la confirmación y `Control` dicen de cuántos. Un curso sin libro queda informado. Si el libro más reciente no se puede exportar (no es de la generación, o su generación no terminó: `Resumen` vacío), el curso queda como problema: no se usa en su lugar un libro anterior, no se exportan sus hojas y no frena a los demás cursos.

**Qué se hace con cada curso** (`ultimaExportacion_`, la misma recuperación de la exportación de un libro, asociada al libro por su ID):
- Última carpeta de PDF completa y vigente (el libro no se modificó en Drive después de crearla): se omite.
- Incompleta y vigente: se completa con las hojas que faltan, sin preguntar aparte; la confirmación lo anuncia.
- No vigente (el libro cambió después de crearse la carpeta), completa o no: no se da por buena ni se completa, porque mezclaría versiones. Se exportan todas las hojas en una carpeta nueva, y la confirmación lo anuncia. Esa carpeta nueva es vigente, así que la ejecución siguiente la continúa o la da por completa.
- Para rehacer a propósito un libro sin cambios, se usa la exportación desde un libro y se elige *Exportar todo en una carpeta nueva*.
- Sin carpeta: se exporta todo en una carpeta nueva junto al libro.

**Confirmación**: antes de empezar, un diálogo resume cuántos cursos se omiten por completos, se completan, se rehacen porque el libro cambió, se exportan enteros o tienen problema, lista lo que se hará con cada curso y cuántos PDF se van a exportar; *Cancelar* no hace nada. Si no hay nada para exportar, no se pregunta ni se pide nada, pero el resultado se informa igual (ver *Resultado*): un curso con un libro que no se puede exportar sigue figurando como problema.

**Varias ejecuciones sin estado propio**: cada ejecución avanza hasta el tiempo máximo (5 minutos desde el inicio, con el diálogo incluido). No se empieza un libro, ni se crea su carpeta, si no entra su primer pedido con la pausa. La ejecución siguiente vuelve a calcular desde Drive qué falta y continúa sin repetir PDF ya exportados. No hay triggers, colas ni registro de progreso fuera de las carpetas de PDF. Las hojas con error quedan faltantes y se vuelven a pedir en la ejecución siguiente. La pausa adaptativa sigue de un libro al siguiente dentro de la ejecución, porque el límite de Google es por usuario.

**Resultado**: `Control!F3` muestra `PDF de varios cursos (<período> <ciclo>): cursos completos: n de n con libro · PDF exportados ahora: n · Con error: n · Reintentos por HTTP 429: n · Tiempo de exportación: n s.` Si faltan PDF, se agrega `Faltan PDF: volver a ejecutar la misma opción para continuar.` El total *con libro* incluye los cursos cuyo libro no se puede exportar; si hay alguno, el resultado agrega `Con libro que no se puede exportar: n` y pide corregirlo. Se informa aunque no haya nada para exportar (entonces sin confirmación ni pedidos), así que un curso con un libro inválido nunca queda como si no hubiera pendientes. Si ningún curso elegido tiene libro, se dice explícitamente. `F4` muestra una línea por curso (cuántos PDF tiene su carpeta y cuántos se exportaron ahora, o por qué no) con el enlace de la carpeta, y las hojas con error.

Casos automatizados (sección 20 de `tests/run.js`; los libros de la sección 16, con sus nombres de la generación, y libros sintéticos de 15 hojas):
- Todos, sin PDF previos: una confirmación con una línea por curso, una carpeta nueva por libro con su ID y un PDF por hoja, 8 pedidos con la configuración validada, y la pausa también entre libros. `Control` con una línea y un enlace por curso, y los cursos sin libro informados. Los libros sólo se leen.
- Otra vez, con todo completo: ningún pedido ni carpeta, sin confirmación, y el resultado con todos los cursos completos.
- Marcados: sólo los cursos con `SÍ`. Sin marcas, error que sugiere *todos*; marca no admitida y período ausente, error claro.
- Varios libros del curso: se usa el creado más recientemente, aun con exactamente el mismo nombre que otro. No cuentan los libros en la papelera ni un archivo que no es planilla.
- Exportación incompleta de un curso: se completa (sólo las hojas que faltan, en esa carpeta), y la carpeta de otro libro con nombre parecido no se toca.
- Corte y continuación: tres libros de 15 hojas con 10 s por pedido. La primera ejecución exporta 24 PDF y deja el tercer libro sin empezar y sin carpeta; la segunda completa el libro a medias y exporta el tercero. En total hay 45 pedidos, ninguno repetido, y una carpeta por libro.
- Error en una hoja: el curso queda incompleto con la hoja detallada; la ejecución siguiente pide sólo esa hoja. Un libro sin `Resumen` cuenta como problema en el total (4 de 5) y los demás se exportan. Con los demás ya completos, el libro inválido sigue marcado como problema, sin confirmación ni pedidos. Sin carpeta del período, se informa que ningún curso tiene libro.
- Generación cortada (regresión): el generador real, cortado antes de escribir `Resumen`, deja un libro que la exportación de un libro rechaza. En varios cursos, un libro así (con `Resumen` vacío y sólo 2 de las 4 hojas) que es el más reciente del curso queda como problema: no se exporta y no se usa en silencio el libro completo anterior.
- Libro modificado después de exportar (regresión): tras una exportación parcial, no se completa y se exporta todo en una carpeta nueva, sin tocar la anterior; tras una completa, no se omite y se exporta todo en una nueva. Modificado antes de la exportación, sigue vigente y se omite. La ejecución siguiente da por completa la carpeta nueva. En la exportación de un libro, con el libro modificado no se ofrece completar; sin modificar, se sigue ofreciendo, y *No* sigue rehaciendo todo.
- Cancelar la confirmación: nada. Confirmación abierta más de 5 minutos: ningún libro ni carpeta.
- Ritmo compartido: después de dos 429 en un libro, el siguiente empieza con la pausa ajustada.

**No cubierto por el arnés**: los tiempos reales con varios cursos y cuántas ejecuciones hacen falta para los 8; si el límite de Google cambia con exportaciones de varios libros seguidas; `getFilesByType` y `getDateCreated` reales en la carpeta del período; que `getLastUpdated` del libro cambie sólo cuando se modifica su contenido (no al abrirlo, exportarlo ni moverlo) y su precisión frente a `getDateCreated` de la carpeta; y el aspecto del diálogo de confirmación con 8 líneas.

### 16. Exportar PDF desde un libro o carpeta de Drive

**Rol**: la operación manual. La exportación de un libro (sección 13) se amplió para aceptar también una carpeta, sin depender de `Control`: el período, el ciclo y los cursos salen de los libros. Desde `Control` quedan *Exportar PDF de los cursos marcados* y *de todos los cursos* (sección 15).

**Entrada**: el diálogo acepta un link de planilla (`/d/<id>`), de carpeta (`/folders/<id>`, también `/drive/u/0/folders/<id>`), con `?id=<id>`, o el ID solo. Sólo se usan operaciones documentadas, cada una para su tipo:
- **Link de carpeta:** se busca sólo como carpeta (`DriveApp.getFolderById`).
- **Link de archivo:** se abre sólo como planilla (`SpreadsheetApp.openById`).
- **ID solo, o `?id=`:** primero como planilla y, si no lo es, como carpeta. Cada operación falla con un ID del otro tipo, así que el resultado no es ambiguo.

Si no es ninguna de las dos, se da un error claro. `DriveApp.getFileById`, que está documentado para archivos y no para carpetas, sólo se usa para nombrar en ese error un archivo que no es planilla. Un ID que no está en Drive, o sin acceso, también da error.

**Modelo común**: cada libro se convierte en un plan (`planLibro_`): sus hojas (si es un libro de boletines terminado), su última carpeta de PDF vigente y lo que falta. Los tres caminos usan ese plan; lo que cambia es cómo se arma la lista:
- **Un libro:** ese libro; si la última exportación vigente quedó incompleta, la vista previa ofrece *Completar* o *Exportar todo en una carpeta nueva*. Si está completa, o el libro cambió después, se exporta todo en una carpeta nueva: elegir un libro a mano es la forma deliberada de rehacerlo.
- **Una carpeta:** el libro más reciente de cada curso que contenga (`librosPorCursoEnCarpeta_`).
- **`Control`:** la misma búsqueda en la carpeta del período, filtrada por los cursos marcados o todos, con el período y el ciclo de `Control`.

Varios libros (carpeta o `Control`) se exportan con una sola función (`exportarPlanesPdf_`), que hace la confirmación, el ritmo compartido, el corte por tiempo, la continuación desde Drive y el resultado.

**Reconocimiento de libros en una carpeta**: se recorren las planillas de Google fuera de la papelera (los demás archivos no se miran):
- **`Resumen` terminado:** el curso es el que la generación escribió en `A2:B5` (hoja origen, nivel, período y ciclo), no el nombre del archivo.
- **`Resumen` sin terminar, o sin hoja `Resumen`:** el curso sale del nombre que le da la generación (`Boletines EPS - <hoja> - <nivel> - <período> <ciclo> - <fecha>`), así que la generación cortada se agrupa con su curso. Con hoja `Resumen` pero sin nombre reconocible, queda sola como problema.
- **Las demás planillas:** no son libros de boletines; se ignoran y se cuentan.
- **Las que no se pueden abrir:** si su nombre es el que da la generación, se agrupan con su curso por ese nombre y compiten por ser la más reciente; si no, se informan como problema propio en la operación manual. Desde `Control` no se atribuyen a ningún curso, porque no hay forma fundada de hacerlo.

De cada curso (hoja origen, nivel, período y ciclo) se elige el libro creado más recientemente. Las regeneraciones anteriores se omiten y se cuentan. El elegido puede ser un libro terminado, uno incompleto o uno que no pudo abrirse. En los dos últimos casos el curso queda como problema, sin volver a uno anterior. Las carpetas de PDF se crean en la carpeta elegida, junto a los libros.

**Confirmación**: carpeta, cantidad de libros de boletines y de cursos, regeneraciones anteriores que se omiten y planillas ignoradas; cuántos cursos se omiten por completos, se completan, se rehacen, se exportan enteros o tienen problema; una línea por curso (con su período) y los PDF a exportar. El resultado (`PDF de la carpeta "<nombre>": cursos completos: …`) tiene la misma forma que el de los cursos de `Control`, y si falta algo pide volver a exportar la misma carpeta.

Casos automatizados (sección 21 de `tests/run.js`):
- **Libro:** con un link `open?id=` se exporta exactamente ese libro. Con su exportación completa y vigente, elegirlo a mano lo rehace en una carpeta nueva. El resto, en la sección 18, que ya usa esta opción.
- **Carpeta con varios cursos**, sin período ni cursos en `Control`: confirmación con el resumen, una carpeta por libro en la carpeta elegida y con su ID, 8 PDF, y libros sólo leídos. También con `/drive/u/0/folders/` y con el ID solo.
- **Mezcla:**
  - una regeneración anterior (omitida);
  - un libro de otro período con el mismo nombre que uno de 3º bimestre (agrupado por su `Resumen`, en su propio curso);
  - una planilla ajena (ignorada);
  - un archivo que no es planilla (ni se abre).
- **Problemas:** el libro más reciente cortado deja su curso como problema y no se exporta el anterior. Una planilla que no se puede abrir se informa como problema.
- **Carpeta sin libros de boletines:** la vista previa lo dice, sin opciones de exportar, pedidos ni escritura en `Control`.
- **Exportaciones previas:** completa vigente (se omite), incompleta (se completa) y de un libro modificado después (se rehace en una carpeta nueva).
- **Corte y continuación:** 45 PDF en varias ejecuciones, sin pedidos repetidos y con una carpeta por libro.
- **Coexistencia:** después de exportar todos los cursos desde `Control`, la carpeta del período desde Drive no tiene nada pendiente, y al revés.
- **Entradas inválidas:** un archivo que no es planilla ni carpeta (por ID y por link `/file/d/`, sin buscarlo como carpeta), un ID inexistente, y un link de carpeta con el ID de un libro (no se abre como libro).
- **Regresión:** el mock de `getFileById` rechaza IDs de carpeta, como lo documenta Apps Script; aun así la carpeta se exporta por link, por ID solo y por `?id=`, sin pedir nunca la carpeta a `getFileById`. Un link de carpeta va directo a `getFolderById`; un ID solo se prueba primero como planilla. Un libro, por link o por ID, no pasa por `getFolderById`.
- **Mocks:** `openById` sólo abre planillas y `getFolderById` sólo carpetas, como en Apps Script.
- **Regresión, planillas que no se pueden abrir:** con un libro válido anterior y una regeneración más reciente del mismo curso sin acceso (mismo nombre), el curso queda como problema y no se exporta el anterior, desde la carpeta y desde `Control`. Con la planilla sin acceso como la anterior, se usa y exporta el libro válido más reciente. Una sin acceso y con nombre irreconocible se informa como problema propio en la carpeta y no afecta a ningún curso desde `Control`.

**No cubierto por el arnés**:
- que `SpreadsheetApp.openById` falle con el ID de una carpeta y `DriveApp.getFolderById` con el de un archivo, como dice la documentación, y cuánto tarda el intento fallido con un ID solo de carpeta;
- que `getDisplayValues` de `Resumen!A1:B5` devuelva el ciclo como texto igual al de `Control` (por ejemplo, `2026` y no `2.026`);
- el tiempo de abrir cada planilla de una carpeta con muchos libros antes de la confirmación (se descuenta de los 5 minutos);
- el aspecto de la confirmación con muchas líneas.

### 17. Diálogo de la exportación desde un libro o carpeta

**Problema**: la exportación desde Drive pedía el link con `ui.prompt` y confirmaba con `ui.alert` dentro de la misma ejecución que exportaba. Esos diálogos dejan la ejecución esperando, y el límite de 6 minutos de Apps Script sigue corriendo. En una prueba real, con los diálogos abiertos varios minutos, quedaron 3 de 20 PDF sin exportar: se exportaron 17 en 95 s, sin errores ni 429, y se alcanzó el tiempo máximo.

**Diseño**: un diálogo HTML (`src/Exportar_PDF_UI.html`, `HtmlService`; en Apps Script, el archivo HTML `Exportar_PDF_UI`, porque no puede llamarse igual que `Exportar_PDF.gs`) que separa la interacción de la exportación en ejecuciones distintas, sin estado guardado entre ellas:
- El menú (`exportarPdfDesdeDrive`) sólo abre el diálogo y termina.
- *Analizar* llama a `previsualizarExportacionPdf(entrada)`: resuelve el link, arma el plan y devuelve el texto de la vista previa, las opciones y una firma del plan. Sólo lee, en todos los caminos: no exporta, no crea carpetas ni archivos, no modifica libros y no escribe en `Control`. Si en una carpeta no hay nada para exportar, la vista previa lo dice y no ofrece exportar. Si el análisis falla (link no reconocible, archivo que no es planilla, libro inválido o sin acceso), el error se muestra en el diálogo y no se registra en `Control`, porque todavía no hubo un intento de exportación. `Control` registra sólo las exportaciones confirmadas, con su resultado o su error.
- Mientras se revisa la vista previa no corre ninguna ejecución.
- *Exportar* (o *Completar* / *Exportar todo en una carpeta nueva*) llama a `exportarPdfConfirmado(entrada, modo, firma)`: una ejecución nueva, cuyo tiempo máximo se cuenta desde ese momento. No confía en la vista previa: vuelve a resolver la entrada y a calcular el plan con el estado actual de Drive, actúa según ese estado y avisa si cambió respecto de lo visto.
- Una sola exportación por vez y por usuario (`LockService.getUserLock()`), por ejemplo con dos diálogos abiertos. El cerrojo se libera al terminar, también con error.
- El diálogo desactiva sus botones mientras espera al servidor. Una vista previa sirve para una sola exportación, y cambiar el link la borra.

Las exportaciones desde `Control` (cursos marcados y todos) no cambian: siguen confirmando con `ui.alert` dentro de su ejecución. Comparten con el diálogo la confirmación, la ejecución y el resultado (`confirmacionPlanesPdf_`, `ejecutarPlanesPdf_`, `resultadoPlanesPdf_`).

**Cuando Drive cambia entre la vista previa y la exportación**, se usa el estado nuevo:
- Una carpeta que otra exportación completó mientras tanto no se vuelve a exportar.
- Si se eligió *Completar* pero el libro se modificó, no se completa la carpeta vieja: se exporta todo en una nueva.
- Si la carpeta a completar ya se completó, no se pide nada.
- Si el libro dejó de estar accesible, se da un error claro.

Casos automatizados (sección 22 de `tests/run.js`; en la sección 18 y siguientes, el arnés simula el diálogo: menú, espera, vista previa, espera y exportación, cada una como ejecución propia):
- El menú abre `Exportar_PDF_UI.html` sin `ui.prompt`. Garantía general, medida al terminar la vista previa: *Analizar* deja `Control` igual, sin carpetas, archivos ni pedidos de PDF, y los libros sólo leídos. Se comprueba con:
  - libro y carpeta con pendientes, y libro con exportación incompleta;
  - carpeta sin pendientes (vacía, ya completa o con un problema);
  - errores de análisis: link no reconocible, archivo que no es planilla, ID inexistente, libro sin `Resumen` y generación cortada;
  - ejecuciones que después sí exportaron.

  Los errores de análisis llegan al diálogo sin registrarse. Las exportaciones confirmadas sí registran su resultado y sus errores (también con el cerrojo tomado), y marcados y todos siguen registrando como antes. La vista de un libro dice si se exporta por primera vez, si su exportación está completa o si el libro cambió.
- Tiempo: con 4 y 10 minutos de espera se exportan los 4 PDF. Las tres ejecuciones son distintas, y los pedidos entran en los 5 minutos de la que exporta.
- Estado nuevo: los cuatro casos de arriba, y sin cambios no hay aviso.
- Cerrojo: con otra exportación en curso, error y nada exportado; al terminar, con o sin error, el cerrojo queda libre.
- `Control`: marcados y todos no abren el diálogo ni pasan por la vista previa.
- Cliente, con un DOM mínimo:
  - sólo llama a las dos funciones públicas del servidor;
  - *Analizar* pide una sola vez y se desactiva;
  - cambiar el link borra la vista previa;
  - *Exportar* llama una sola vez, con el link y la firma de la vista, y bloquea todo mientras espera;
  - una vista previa ya usada no permite volver a exportar;
  - un libro incompleto ofrece las dos opciones;
  - un error se muestra y deja volver a analizar.

**No cubierto por el arnés**:
- el diálogo real en Google Sheets: aspecto, tamaño y comportamiento de `google.script.run`;
- que una llamada de `google.script.run` pueda durar hasta el límite de 6 minutos y que su resultado llegue al diálogo;
- que `ss.toast` se vea durante esa llamada;
- que `LockService.getUserLock()` funcione como se espera;
- qué pasa si se cierra el diálogo durante la exportación (la ejecución debería seguir, sin mostrar el resultado; queda en `Control`).
