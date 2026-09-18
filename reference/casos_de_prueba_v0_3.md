# Casos de prueba v0.3 (3º bimestre)

Este documento describe los escenarios mínimos de prueba reproducibles para el Generador EPS v0.3, cubriendo la incorporación del 3º bimestre, su acumulación con períodos anteriores (1º bimestre y 1º cuatrimestre) y la protección contra regresiones históricas.

## Arnés de pruebas local

Para facilitar la reproducibilidad sin requerir interacción manual con Google Sheets en cada iteración de desarrollo, los escenarios están automatizados en `tests/run.js`.

### Características y funcionamiento del arnés

- **Ejecución de código real**: `tests/run.js` carga y ejecuta el archivo de código fuente real `src/Codigo_Apps_Script_Generador_EPS_v0_3.gs` dentro de un contexto aislado (`node:vm`).
- **Mocks locales mínimos**: Emula exclusivamente la superficie mínima de Google Apps Script requerida por la lógica de negocio (`SpreadsheetApp.getActiveSpreadsheet()`, `SpreadsheetApp.openById()`, `Sheet`, `Range`, `Ui.alert()`, `Ui.createMenu()`).
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

### 2. Materia nueva en 3B
- **Caso**: Carla Cuatro (DNI `99000003`, EPSI), cursa "Taller Nuevo 3B" incorporado únicamente en el 3º bimestre.
- **Resultado esperado**:
  - Generación sin errores.
  - Nivel visual en `E6`: `Informática 1`.
  - En la fila de *Taller Nuevo 3B*, columnas 1B y 1C vacías, y columna 3B con `Suficiente` (`||Suficiente`).

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
