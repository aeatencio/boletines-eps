# Casos de prueba v0.3 (3º bimestre)

Este documento describe los escenarios mínimos de prueba reproducibles para el Generador EPS v0.3, cubriendo la incorporación del 3º bimestre, su acumulación con períodos anteriores (1º bimestre y 1º cuatrimestre) y la protección contra regresiones históricas.

## Arnés de pruebas local

Para facilitar la reproducibilidad sin requerir interacción manual con Google Sheets en cada iteración de desarrollo, los escenarios están automatizados en `tests/run.js`.

### Características y funcionamiento del arnés

- **Ejecución de código real**: `tests/run.js` carga y ejecuta el archivo de código fuente real `src/Codigo_Apps_Script_Generador_EPS_v0_3.gs` dentro de un contexto aislado (`node:vm`).
- **Mocks locales mínimos**: Emula exclusivamente la superficie mínima de Google Apps Script requerida por la lógica de negocio (`SpreadsheetApp.getActiveSpreadsheet()`, `SpreadsheetApp.openById()`, `SpreadsheetApp.create()`, `Utilities.formatDate()`, `Sheet`, `Range`, `Ui.alert()`, `Ui.createMenu()`).
- **Consumo directo de fixtures vigentes**: En cada ejecución se vacía el directorio de trabajo temporal `tmp/unzipped/` y se extraen directamente los archivos XLSX y TSV que existen en `reference/fixtures/`. Si falta algún fixture o se modifica, el runner valida estrictamente el estado actual del disco y no reutiliza extracciones de corridas previas.
- **Requisito de plataforma y shell**: El runner debe ejecutarse en un entorno **Windows** desde **PowerShell** o **`cmd`**, dependiendo exclusivamente del comando `tar.exe` provisto por el sistema operativo en `%SystemRoot%\System32` (bsdtar nativo de Windows). **No debe asumirse que funciona desde Git Bash**, ya que en dicho entorno `tar.exe` puede resolverse al binario provisto por MSYS y fallar al abrir los archivos XLSX empaquetados.

### Aspectos NO cubiertos por el arnés local

El arnés local valida la lógica de lectura tabular, normalización, acumulación, descarte de valores y asignación en la cuadrícula de salida. **NO cubre ni emula**:
1. El motor completo de Google Apps Script ni la infraestructura de Google Cloud / Google Drive (permisos, autenticación, cuotas, triggers instalables).
2. El motor de cálculo de Google Sheets ni validaciones nativas de datos en celdas.
3. Aspectos visuales y de formato del workbook: tipografías, alineación, bordes, anchos de columna y altos de fila efectivos en Google Sheets.
4. Evaluación del formato condicional en tiempo real en la interfaz de Google Sheets.
5. Importación real de planillas en formato nativo de Google Sheets.
6. Presentación final para impresión ni generación de PDF.

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
- **Estructuras ambiguas**: falta o repetición de `ORIENTACIÓN`, materia repetida, materia sin columna de valoración y DNI fuera de la fila "Valoración" son errores en la fuente principal; en una fuente anterior, esa fuente se descarta con advertencia.

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

**Uso**: en `Control`, además del período, la hoja origen, el ciclo lectivo y las fuentes, se elige el nivel en `B9` (*Nivel / curso*, selector `EPS1`…`EPS4` incluido en el workbook versionado; también se acepta sólo el número). *Informes EPS > Generar boletines por curso* crea un libro nuevo `Boletines EPS - <hoja origen> - <nivel> - <período> <ciclo> - <fecha y hora>` con la hoja `Resumen` y una hoja por estudiante generado. El DNI de `B7` no se usa. La generación individual ignora `B9` y no cambia.

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
- `Resumen` oculta y boletines visibles, también con un estudiante fallido. Si ninguna hoja pudo crearse, `Resumen` queda visible con los errores (el mock rechaza ocultar la última hoja visible, como Google Sheets).
- `Control` del xlsx versionado: `A9` = *Nivel / curso*, `B9` vacío con selector `EPS1,EPS2,EPS3,EPS4`. Con esa hoja, la generación por curso usa el nivel elegido, pide `B9` si no se eligió, y la generación individual lo ignora (aunque tenga un valor no válido).
- Coordenadas de `Control`: las fuentes de 1º bimestre, 1º cuatrimestre y 3º bimestre están en `B10`, `B11` y `B12`, sin fila vacía intermedia. Los tests toman todas las coordenadas de `CONFIG_EPS` y comprueban en el xlsx que cada una esté junto a su rótulo, junto con los valores por defecto y las validaciones de `B5`, `B6` y `B9`.
- Columnas operativas sólo en la principal: en 3B, 1C y 1B sin columnas operativas son válidas. En 1C, 1B sin ellas es válida. En 1B, la principal 1B las exige. Una histórica con `LISTO PARA EXPORTAR` duplicada, `EXPORTADO` inválido o encabezados operativos malformados produce la misma tanda que una histórica limpia. La misma grilla 1C sin `LISTO PARA EXPORTAR` es válida como histórica de 3B y es un error sólo como principal (período 1º cuatrimestre). La principal con la columna duplicada también es un error.

**Cambio en la generación individual**: el contenido del informe no cambia. Las advertencias de las fuentes (falta de una fuente anterior, fuente anterior ilegible, encabezados de la cola) ahora aparecen antes que las del estudiante.

**No cubierto por el arnés** (queda para la validación manual de #11): copia real de `Plantilla` a otro libro (imágenes, celdas combinadas, formato condicional, anchos y altos), permisos de creación en Drive, tiempos y límites de Apps Script con un curso completo, y PDF e impresión del libro.
