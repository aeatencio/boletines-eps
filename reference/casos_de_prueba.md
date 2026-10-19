# Casos de prueba del generador EPS

Este documento describe los escenarios mínimos de prueba reproducibles para el Generador EPS. Las secciones 1 a 4 vienen de v0.3 (incorporación del 3º bimestre, acumulación con 1º bimestre y 1º cuatrimestre y regresiones históricas); las secciones 5 a 8 cubren v0.4 (grillas reales, generación por curso y orientación por período), la sección 11, la generación de varios cursos (#20); la sección 13, la exportación a PDF (#21), y la sección 14, la organización de los libros en Drive (#23).

## Arnés de pruebas local

Para facilitar la reproducibilidad sin requerir interacción manual con Google Sheets en cada iteración de desarrollo, los escenarios están automatizados en `tests/run.js`.

### Características y funcionamiento del arnés

- **Ejecución de código real**: `tests/run.js` carga y ejecuta el archivo de código fuente real `src/Generador_EPS.gs` dentro de un contexto aislado (`node:vm`).
- **Mocks locales mínimos**: Emula exclusivamente la superficie mínima de Google Apps Script requerida por la lógica de negocio (`SpreadsheetApp.getActiveSpreadsheet()`, `SpreadsheetApp.openById()`, `SpreadsheetApp.create()`, `Utilities.formatDate()`, `Sheet`, `Range`, `Ui.alert()`, `Ui.createMenu()`; para guardar los libros, `DriveApp` con carpetas, papelera y `moveTo()`; para la exportación a PDF, también `Ui.prompt()`, `UrlFetchApp.fetch()`, `DriveApp` y `ScriptApp.getOAuthToken()`).
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

**Uso**: *Informes EPS > Exportar PDF de un libro de boletines* pide el link o ID de un libro de la generación por curso (o de varios cursos), ya revisado. Crea una carpeta nueva junto al libro, `<nombre del libro> - PDF <fecha y hora>`, con un PDF por hoja de estudiante (todas menos `Resumen`), con el nombre de la hoja. Si no se puede ver la carpeta del libro, la crea en Mi unidad.

**Separación**: es una operación posterior y separada de la generación, en su propio archivo (`src/Exportar_PDF.gs`). `Generador_EPS.gs` sólo la nombra en el menú y no usa sus servicios (`UrlFetchApp`, `ScriptApp`); usa `DriveApp` únicamente para guardar sus libros (sección 14). Como los libros quedan en la carpeta de su ciclo y período, las carpetas de PDF quedan ahí, junto al libro. El libro sólo se lee: una falla de exportación nunca obliga a regenerarlo. Cada exportación crea su carpeta, así que volver a exportar no reemplaza ni mezcla PDF anteriores.

**Mecanismo**: la URL de exportación de Google Sheets (`/export?format=pdf&gid=…`), pedida con `UrlFetchApp` y el token del usuario. No está documentada oficialmente, por eso queda aislada en `urlExportacionPdf_` y `exportarHojaPdf_`. Cada pedido lleva la configuración de impresión completa (`CONFIG_PDF.PARAMETROS`), sin depender de lo guardado en el diálogo de impresión: `size=A4`, `portrait=true`, `fitw=true` y `scale=2` (ajustar al ancho), `top_margin`, `left_margin` y `right_margin` en 0, `bottom_margin=0.3937` (1 cm en pulgadas) y `fzr=true` (repetir filas inmovilizadas, la fila 1 de cada boletín; ver sección 9), sin cuadrícula, notas, título, nombre de hoja ni números de página.

**Errores**:
- Link no reconocible, libro sin `Resumen` (por ejemplo, el propio generador) o sin hojas de estudiantes: error claro, sin crear la carpeta.
- Por hoja: una respuesta HTTP distinta de 200 (429 se señala como demasiadas exportaciones seguidas), una respuesta que no es PDF o una excepción del pedido se registran en esa hoja, y las demás se exportan igual. No hay reintentos.
- Tiempo: pasados 5 minutos (`CONFIG_PDF.TIEMPO_MAXIMO_MS`), antes del límite de 6 de Apps Script, no se empieza otra exportación y las hojas que faltan se informan como sin exportar.

**Resultado**: `Control!F3` muestra `PDF exportados: n de n · Con error: n · Sin exportar: n · Duración: n s (n s por PDF).`, el libro y el enlace a la carpeta. `F4` lista cada hoja con error o sin exportar, con el motivo. Un aviso (`toast`) indica la hoja en curso.

Casos automatizados (sección 18 de `tests/run.js`, sobre el libro EPS2 de la generación por curso):
- Libro completo: carpeta junto al libro, 4 PDF en el orden del libro y sin `Resumen`, un pedido por hoja con su `gid`, configuración de impresión validada e idéntica en todos los pedidos, token del usuario y `muteHttpExceptions`.
- El libro no cambia (hojas, valores, formatos, visibilidad, filas inmovilizadas) y sólo se usan métodos de lectura.
- Fallas por hoja (HTTP 429, respuesta HTML y excepción): las demás se exportan, `Control` detalla cada una y el libro no cambia; volver a exportar da los 4 PDF.
- Tiempo máximo: con 3 minutos por PDF se exportan 2 de 4 y las otras 2 quedan informadas.
- Libro sin `Resumen`, libro sin hojas de estudiantes, cancelación y link inválido: no se crea carpeta ni se exporta.
- `Generador_EPS.gs` no usa `UrlFetchApp` ni `ScriptApp`, y de `DriveApp` sólo usa `getRootFolder` y `getFileById` (sección 14).

**No cubierto por el arnés**: el PDF que produce Google (tamaño, márgenes, ajuste al ancho, repetición de la fila 1, cortes de página), la vigencia de los parámetros de la URL, los permisos y cuotas reales, los tiempos y la aparición de errores 429. Queda pendiente la validación en Google real que pide #21: un boletín de una página, uno de dos páginas, un curso completo, independencia de la configuración guardada en el diálogo de impresión y tiempos reales.

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
