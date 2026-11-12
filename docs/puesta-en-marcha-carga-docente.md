# Puesta en marcha de la carga docente

Manual operativo del Equipo EPS para instalar, probar, habilitar y mantener la carga docente (`src/carga/`). Las reglas funcionales (ventanas, enlaces, correos, Registro) están en el [README](../README.md#carga-docente-en-desarrollo). Este manual sólo dice qué hacer, en qué orden y cómo comprobarlo.

**Estado (#29).** Nada de la carga se ejecutó todavía en Google real. Los resultados esperados salen del código y de las pruebas locales; los de las etapas B a D, además, están ensayados con dobles de Google en `tests/carga-recorrido.js`, que falla si el código deja de coincidir con este manual. La marca **[G]** señala comportamientos de Google que el código no controla: la prueba integrada (etapa C) los tiene que observar y anotar.

**Regla general.** Si un resultado no coincide con el esperado, parar, anotarlo en el acta (anexo 2) y resolverlo antes de seguir. No corregir datos a mano para que «dé bien».

| Etapa | Qué se logra | Quién la habilita |
|---|---|---|
| [A. Preparación](#a-preparación-de-los-recursos-productivos) | Recursos definitivos creados y configurados; app publicada pero sin acceso | Responsable técnico, con la cuenta propietaria |
| [B. Verificación sin servicio](#b-verificación-sin-habilitar-el-servicio) | Configuración comprobada, nadie puede cargar | Ídem |
| [C. Prueba integrada](#c-prueba-integrada-controlada) | Recorrido real con cuentas de prueba, sin dejar cambios en las valoraciones | Ídem, con aviso al Equipo EPS |
| [D. Limpieza](#d-limpieza-y-estado-posterior) | Estado previo restaurado y comprobado; acta cerrada | Ídem |
| [E. Apertura](#e-apertura-de-la-ventana-institucional) | Ventana real cargada; correos activados | **Autorización expresa del Equipo EPS** |
| [F. Operación](#f-operación-cierre-y-mantenimiento) | Seguimiento diario, incidencias, cierre y períodos siguientes | Equipo EPS |

---

## 1. Piezas y cuentas

| Pieza | Qué es | Quién edita |
|---|---|---|
| Generador | La planilla del generador con su Apps Script. La carga sólo lee `Control!B10:B14` (la grilla de cada período) | Equipo EPS |
| Grillas de 2º cuatrimestre y anual | Una planilla por período, con hojas `EPSI` y `EPSG`. Son la fuente de las valoraciones: la carga escribe en ellas y el generador las lee | Equipo EPS en Sheets; los docentes, sólo a través de la app |
| Planilla de administración | `Ventanas`, `Asignaciones`, `Enlaces` (las edita el Equipo EPS), `Registro` y `Correos` (las escribe la carga). La preparación de A3 agrega `Asignaciones (borrador)`, `Preparación` y `Preparación · cambios` | Equipo EPS; `Registro`, `Correos` y `Preparación · cambios`, sólo la cuenta propietaria |
| Proyecto de Apps Script «Carga docente EPS» | Proyecto independiente con los archivos de `src/carga/`: la app web y el activador diario de los correos | Sólo la cuenta propietaria |
| Calendario de feriados | `Carga_Calendario.gs`, dentro del proyecto: los feriados por año que deciden los días hábiles (cierre de las ventanas y días de envío de correos). No hay hoja de feriados | Responsable técnico, al preparar cada año (F8) |
| Cuenta propietaria | Cuenta institucional del Equipo EPS (hoy, una cuenta de `cfp7.edu.ar`). Es dueña del proyecto, ejecuta la app para todos (los docentes no necesitan permisos sobre las planillas), crea el activador y envía los correos como «Equipo EPS» | — |

Todo lo del proyecto (publicar, ejecutar funciones, crear el activador) se hace **con la cuenta propietaria**. Las funciones del editor exigen sólo que no las ejecute un visitante de la app: si otra cuenta con permiso de edición sobre el proyecto creara el activador, los correos saldrían de esa cuenta. Por eso el proyecto no se comparte con otras cuentas.

## 2. Ficha de la instalación

Los datos propios de esta instalación no van al repositorio. Llevarlos en un documento del Equipo EPS, al que sólo acceda el equipo:

| Dato | Dónde se obtiene |
|---|---|
| Cuenta propietaria | — |
| ID del generador, de la planilla de administración y de las grillas de 2º cuatrimestre y anual | La parte de la URL entre `/d/` y `/edit` |
| ID del proyecto de Apps Script | *Configuración del proyecto* |
| URL `/exec` e ID de la implementación, con su número de versión | *Implementar > Gestionar implementaciones* |
| Commit (o tag) del repositorio instalado | `git log -1` en la copia de la que se pegó el código |
| Modalidad de acceso elegida (paso A9) | — |
| Cuentas de prueba, fechas de la prueba y acta | Etapa C |
| Autorización de apertura: quién, cuándo y fechas de la ventana | Etapa E |

**Nunca** anotar ni copiar la propiedad `CLAVE_ENLACES` ni un enlace de docente. Tampoco reenviar correos de la carga.

---

## A. Preparación de los recursos productivos

### A1. Antes de empezar

- La cuenta propietaria tiene permiso de edición sobre el generador y sobre las grillas de 2º cuatrimestre y anual, y puede crear proyectos en <https://script.google.com>.
- Se conoce el commit del repositorio que se va a instalar (el revisado) y se trabaja desde esa copia.
- El Equipo EPS sabe que se va a probar sobre las grillas reales (etapa C) y quién será la cuenta docente de prueba.

### A2. Generador actualizado

El generador institucional tiene que admitir 2º cuatrimestre y anual: hoja `Control` con las filas 13 y 14 y el script de `src/Generador_EPS.gs`. Si todavía no se hizo, seguir [Actualizar el generador institucional](../README.md#actualizar-el-generador-institucional).

**Comprobar**: `Control!A13` y `A14` dicen 2º cuatrimestre y anual; *Varios cursos* está en `A16`; `B5` ofrece `2º cuatrimestre` y `Anual`; un boletín de un período ya usado coincide con el generado antes.

### A3. Grillas de 2º cuatrimestre y anual

Se preparan a partir de la del 1º cuatrimestre (*Archivo > Hacer una copia*, **sin** marcar «Compartir con las mismas personas»). Una planilla por período: dos períodos nunca comparten grilla.

| Qué | Regla (la del código) |
|---|---|
| Hojas | `EPSI` y `EPSG`, con esos nombres exactos. La hoja `Opciones` queda igual |
| Encabezados (filas 1 y 2) | Como en 1º cuatrimestre: una única columna `ORIENTACIÓN`; cada materia, antes de `ORIENTACIÓN`, con su columna de valoración a la derecha; marcadores `EPS1`…`EPS4` en la fila 1 sobre las materias de cada nivel |
| Estudiantes | Agrupados bajo filas `EPS1`…`EPS4` en la columna A. Cada uno, en un bloque de cuatro filas (Valoración, Calidad de las producciones, Participación, Observaciones) con sus etiquetas; DNI en la columna C de la primera fila, una sola vez por hoja |
| Calificación (columna de valoración, primera fila del bloque) | `CURSA` si lo cursa en ese período; `NO CURSA` o vacío si no; `ACREDITÓ` o `EQUIVALENCIA`, si corresponde, los pone el Equipo EPS. **Ningún otro valor**: un número que quedó de 1º cuatrimestre aparecería como ya cargado |
| Calidad, Participación, Observaciones | Vacías. Si quedan las de 1º cuatrimestre, el docente las vería como propias del período |
| Campos generales y operativos (después de `ORIENTACIÓN`) | Los decide el Equipo EPS para el generador (`LISTO PARA EXPORTAR`, `EXPORTADO`, observaciones generales). La carga no los toca |
| Formato de la columna de valoración | El de 1º cuatrimestre. Un formato con decimales (`7,00`) hace que la carga marque el guardado como `revisar` |
| Protecciones | Si hay rangos protegidos, la cuenta propietaria tiene que poder editar las cuatro filas de cada estudiante en las columnas de valoración |
| Compartir | Equipo EPS y cuenta propietaria como editores. **Los docentes no necesitan acceso**: si lo tuvieran, podrían editar fuera de la ventana y sin Registro |

En el generador, completar `Control!B13` (2º cuatrimestre) y `B14` (anual) con la URL o el ID de cada grilla. `B11` (1º cuatrimestre) y `B12` (3º bimestre) tienen que seguir apuntando a sus grillas: la preparación las usa.

#### A3.1 Preparación automática de cursadas y asignaciones (#32)

Las reglas de la tabla anterior no se aplican a mano: lo hace `prepararPeriodosCarga`, después de A5 a A7 (necesita el proyecto y la planilla de administración) y **antes de cargar `Ventanas`** de 2º cuatrimestre o anual: con una fila de esos períodos en `Ventanas`, no modifica nada. Respeta las ediciones que el Equipo EPS ya hizo en los borradores.

| Qué hace | Regla |
|---|---|
| Cursadas | `CURSA` donde **3º bimestre** (el último período calificado, con la matrícula vigente) tiene una calificación ordinaria (`En proceso`, `Suficiente`, `Avanzado`). `ACREDITÓ`, `EQUIVALENCIA`, `CURSA` sin calificar o vacío no habilitan `CURSA`. Anual se prepara con la misma evidencia: 2º cuatrimestre todavía no está calificado y no se usa |
| Restos de la copia | Una calificación que sigue igual a la de 1º cuatrimestre pasa a `CURSA` sólo con esa evidencia; si no la hay, queda y se informa. Calidad, Participación y Observaciones iguales a las de 1º cuatrimestre se borran |
| Lo que no toca | Lo que ya no es igual a 1º cuatrimestre (cambios del Equipo EPS), `ACREDITÓ` y `EQUIVALENCIA`, filas y columnas (nunca inserta, borra ni mueve), `ORIENTACIÓN` y los campos posteriores (`LISTO PARA EXPORTAR`, observaciones generales…), lo que la carga docente ya escribió (`Registro`), 1º cuatrimestre y 3º bimestre (sólo las lee) |
| `Asignaciones (borrador)` | Una fila por orientación y materia con cursantes, con orientación, nivel y materia de la grilla. Correo: sólo si **todas** las protecciones por columna de las grillas (las de la carga directa anterior) que cubren esa materia nombran a la misma persona; si no, queda vacío y `Revisión` dice qué encontró. Nombre: sólo si ese correo ya tiene uno escrito en el borrador o en `Asignaciones`. En las filas existentes nunca cambia `Orientación`, `Nivel`, `Materia`, `Docente` ni `Correo`; reescribe sus columnas propias, `Cursantes` y `Revisión`, y agrega `Origen del correo` en las filas que crea. **No toca la hoja `Asignaciones`** |
| Revisión | Matrícula, nombres, DNI, grupos, orientación, materias y niveles de cada borrador contra 3º bimestre y 1º cuatrimestre, y de los dos borradores entre sí |

Pasos, con la cuenta propietaria, desde el editor del proyecto:

1. Asignar nombre a la versión actual de las dos grillas («Antes de la preparación <fecha>»), como en C1.3.
2. Ejecutar `previsualizarPeriodosCarga`. No modifica grillas ni el borrador: escribe el informe en la hoja `Preparación` de la planilla de administración y un resumen en el registro de ejecución.
3. Revisar `Preparación`: el resumen (qué se pasaría a `CURSA`, qué se borraría, cuántas calificaciones eran iguales a 1º cuatrimestre) y la tabla, con una vista de filtro por `Prioridad`:
   - **Bloqueo**: impide preparar todo, una hoja o un estudiante (estructura no reconocida, DNI repetido, bloque incompleto, ventana cargada, falta una grilla en `Control`). Se corrige antes de aplicar.
   - **Decidir**: quedó sin preparar y lo decide una persona: una calificación de 1º cuatrimestre sin evidencia en 3º bimestre, una nota que no viene de la copia, un valor no reconocido, una fila sin DNI, una materia que no está en 3º bimestre, una fila del borrador sin correo o nombre.
   - **Revisar**: se conservó lo del borrador, pero difiere de la evidencia o de lo esperado (estudiantes que faltan o sobran respecto de 3º bimestre, nombres distintos para un DNI, posible DNI mal escrito, cambios de grupo u orientación, diferencias entre 2º cuatrimestre y anual, `LISTO PARA EXPORTAR` u observaciones generales iguales a las de 1º cuatrimestre).
4. Ejecutar `prepararPeriodosCarga`. Vuelve a armar el plan y, si no hay bloqueos generales, lo aplica: cada celda modificada queda en `Preparación · cambios` con su valor anterior y el estado verificado (`aplicado`; `no aplicado` si alguien la cambió desde la lectura; `revisar` si no quedó como se esperaba). Si no se aplicó, la ejecución termina con error y `Preparación` dice por qué.
5. Resolver en las grillas y en `Asignaciones (borrador)` lo marcado **Decidir** y lo que corresponda de **Revisar**, y volver a ejecutar `prepararPeriodosCarga` hasta que sólo queden casos entendidos. Repetirla no cambia lo ya preparado ni duplica filas, y no deshace decisiones del Equipo EPS.

**Comprobar**: `Preparación` sin **Bloqueo**; en `Preparación · cambios`, ninguna fila `pendiente` ni `revisar`; `Asignaciones (borrador)` con correo y nombre en cada fila y `Revisión` vacía. Después lo revisa `verificarCargaDocente` en el paso B4 (estructura, DNI repetidos, valores no reconocidos, bloques incompletos y un resumen de cada hoja). `Registro` y `Preparación · cambios` se protegen igual (A8).

### A4. Planilla de administración

1. Con la cuenta propietaria, crear una planilla nueva, por ejemplo «Carga docente EPS · Administración».
2. *Archivo > Configuración*: configuración regional **Argentina** y zona horaria **(GMT-03:00) Buenos Aires**. Con otra configuración regional, `05/12/2026` se lee como 12 de mayo.
3. Compartirla sólo con el Equipo EPS, como editores. Las hojas las crea el paso A7.

### A5. Proyecto de Apps Script

1. Con la cuenta propietaria, en <https://script.google.com>, *Nuevo proyecto*. Nombre: «Carga docente EPS». No vincularlo a ninguna planilla.
2. *Configuración del proyecto*: activar «Mostrar el archivo de manifiesto `appsscript.json` en el editor».
3. Crear los archivos con estos nombres exactos y pegar el contenido de `src/carga/`. Apps Script agrega la extensión; el orden no importa.

   | Tipo (*+ > …*) | Nombre | Contenido |
   |---|---|---|
   | Secuencia de comandos | `Carga_Calendario` | `Carga_Calendario.gs` |
   | Secuencia de comandos | `Carga_Nucleo` | `Carga_Nucleo.gs` |
   | Secuencia de comandos | `Carga_Google` | `Carga_Google.gs` |
   | Secuencia de comandos | `Carga_Acceso` | `Carga_Acceso.gs` |
   | Secuencia de comandos | `Carga_Web` | `Carga_Web.gs` |
   | Secuencia de comandos | `Carga_Correo` | `Carga_Correo.gs` |
   | Secuencia de comandos | `Carga_Correo_Google` | `Carga_Correo_Google.gs` |
   | Secuencia de comandos | `Carga_Preparacion` | `Carga_Preparacion.gs` |
   | Secuencia de comandos | `Carga_Preparacion_Google` | `Carga_Preparacion_Google.gs` |
   | HTML | `Carga_Pagina` | `Carga_Pagina.html` |
   | HTML | `Carga_Cliente` | `Carga_Cliente.html` |

4. Borrar el `Código.gs` inicial. Reemplazar el contenido de `appsscript.json` por el de `src/carga/appsscript.json`.
5. Guardar.

**Comprobar**: el editor guarda sin errores y lista 9 archivos `.gs`, 2 `.html` y `appsscript.json`. En una instalación hecha antes de #32, agregar los dos archivos `Carga_Preparacion*` (no cambian la app web: son funciones del editor). El manifiesto tiene zona `America/Argentina/Buenos_Aires`.

### A6. Propiedades del script

En *Configuración del proyecto > Propiedades del script*, agregar:

| Propiedad | Valor |
|---|---|
| `ADMIN_ID` | ID de la planilla de administración |
| `GENERADOR_ID` | ID del generador |

`CLAVE_ENLACES` la crea el paso A7 y `URL_CARGA` se completa en el A9. **No** crear `CLAVE_ENLACES` a mano.

### A7. Hojas de administración y secreto

1. En el editor, elegir `prepararAdministracionCarga` y *Ejecutar*. La primera vez Google pide autorizar cuatro permisos: planillas, enviar correo como la cuenta, administrar activadores y ver la dirección de correo. **[G]** Anotar en el acta cualquier advertencia o bloqueo de la pantalla de autorización.
2. **Resultado esperado** en el registro de ejecución: «Se creó la propiedad CLAVE_ENLACES…» y «Listo. Completar Ventanas y Asignaciones…», con los años del calendario de feriados.

**Comprobar**: la planilla de administración tiene `Ventanas`, `Asignaciones`, `Enlaces`, `Registro` y `Correos`, cada una con sus encabezados en negrita, y la propiedad `CLAVE_ENLACES` existe. Al ejecutarla de nuevo, no crea nada ni cambia el secreto.

### A8. Contenido inicial de la administración

| Hoja | Ahora |
|---|---|
| `Registro`, `Correos` | Vacías. Protegerlas: *Datos > Proteger hojas y rangos > Hoja*, editar sólo la cuenta propietaria. **Nunca** borrar ni editar sus filas, ni sus columnas `Estado` y `Detalle`. Para anotar algo, agregar una columna a la derecha (por ejemplo, «Nota EPS»): la carga sólo usa sus propias columnas |
| `Ventanas` | **Vacía** (sólo el encabezado). Mientras esté vacía, ningún enlace vale |
| `Asignaciones` | **Vacía**. La lista real se prepara en la hoja `Asignaciones (borrador)`, con las mismas columnas, que crea y completa la preparación (A3.1). La carga sólo lee la hoja llamada exactamente `Asignaciones` |
| `Enlaces` | Vacía |

No hay que cargar feriados: están en el código (`Carga_Calendario.gs`). Si la planilla tiene una hoja `Feriados` de una instalación anterior, la carga no la lee; no hace falta borrarla ni modificarla.

### A9. Publicar la app web

1. *Implementar > Nueva implementación > Seleccionar tipo > Aplicación web*. Descripción: «Carga docente EPS».
2. *Ejecutar como*: **Yo** (la cuenta propietaria).
3. *Quién tiene acceso* (**decisión de acceso**). Las dos opciones posibles:

   | Opción del diálogo | Manifiesto | Qué pasa al abrir el enlace |
   |---|---|---|
   | «Cualquier persona» (*Anyone*) | `ANYONE_ANONYMOUS` | Abre sin iniciar sesión en Google |
   | «Cualquier persona con una cuenta de Google» (*Anyone with Google account*) | `ANYONE` | Google pide iniciar sesión con alguna cuenta antes de abrir |

   En las dos, la app se ejecuta como la cuenta propietaria, la credencial de la app es el enlace individual firmado (sin un enlace vigente no se ve ni se guarda nada) y las funciones del editor siguen reservadas a la cuenta propietaria. La diferencia es el requisito previo de Google: con `ANYONE`, Google exige además iniciar sesión con alguna cuenta antes de abrir la página; con `ANYONE_ANONYMOUS`, no. Esa sesión no identifica al docente: la app no la usa, y para cuentas de otra organización, como las `@bue.edu.ar`, ni siquiera puede leer su correo (comprobado en una prueba previa con esas cuentas). A cambio, el inicio de sesión agrega fricción: dentro de la app de Gmail en el celular, o con varias cuentas iniciadas en el navegador.
   - **Elegida**: «Cualquier persona» (`ANYONE_ANONYMOUS`), la del manifiesto. **[G]** Falta comprobar en Google real que la organización la ofrezca y que abra bien (B3 y C2.6).
   - **Si el diálogo no la ofrece** (política de la organización): «Cualquier persona con una cuenta de Google»: la app sigue usando el mismo enlace firmado como credencial, y Google exige además iniciar sesión. En ese caso, cambiar `"access"` a `"ANYONE"` en el `appsscript.json` del proyecto y avisar para actualizarlo en el repositorio, de modo que el manifiesto que se pegue en el futuro coincida con lo elegido. **[G]** Después de cada actualización (F6), comprobar en *Gestionar implementaciones* que la modalidad de acceso siga siendo la elegida.
4. *Implementar*. Copiar la **URL de la aplicación web** (termina en `/exec`) y guardarla en la propiedad `URL_CARGA`. Anotar en la ficha la URL, el ID de la implementación y la modalidad elegida.

**Resultado esperado**: la app queda publicada **sin habilitar a nadie**, porque `Ventanas` está vacía.

**[G]** Si la URL tiene la forma `https://script.google.com/a/macros/<dominio>/s/…/exec`, el paso B3 comprueba que se abra sin pedir una cuenta de ese dominio. Si la pide, usar la forma `https://script.google.com/macros/s/<mismo ID>/exec` en `URL_CARGA`. La carga acepta las dos.

---

## B. Verificación sin habilitar el servicio

### B1. `verificarCargaDocente`

Ejecutarla desde el editor. **Resultado esperado** en el registro:

```
Calendario de feriados: 2026

Ventanas vigentes o próximas:
(ninguna)

Grillas de esas ventanas:
(ninguna)

Problemas:
Los correos no están activados (activarCorreosCarga).
```

La primera línea lista los años del calendario de feriados del código. Si la función falla con un error como «aniosCalendario_ is not defined», falta el archivo `Carga_Calendario` (A5). Cualquier otro problema se resuelve antes de seguir. Por ejemplo, un `URL_CARGA` mal copiado o una hoja sin sus columnas.

### B2. `previsualizarCorreosCarga`

**Resultado esperado**: una línea JSON con `"simulados":0` y el `motivo` «No hay ventanas en curso ni cierres para informar.» (antes de las 8:00, «Antes de las 8:00 no se envían correos.»).

### B3. La página no habilita a nadie

En una ventana privada del navegador, sin sesión de Google:

1. Abrir la URL `/exec` sola. **Esperado**: «Acceso no verificado», con «El enlace no es válido, fue reemplazado o es de una carga anterior…».
2. Abrir la URL con `?k=` y 43 letras cualesquiera. **Esperado**: lo mismo.
3. **[G]** Con la opción elegida en el paso A9 («Cualquier persona»), no se pide iniciar sesión. Anotar el aviso de Apps Script que se vea («Esta aplicación fue creada por otro usuario…» o similar).

### B4. Grillas y configuración de la prueba, sin abrir

Se hace el **día hábil anterior** a la prueba (día 0), cuando ya están decididos los datos del paso C1. La ventana de la prueba queda cargada **a futuro**: se verifica todo, pero nadie tiene enlace hasta las 0:00 del día 1.

1. Cargar `Asignaciones` y `Ventanas` de la prueba como dice C1.
2. Ejecutar `verificarCargaDocente`. **Resultado esperado**:
   - `Ventanas vigentes o próximas`: `2º cuatrimestre: <día 1> a <día 2>, cierra <día 3> 08:00 (futura)` y lo mismo para Anual. `<día 3>` tiene que ser el día hábil siguiente a `<día 2>` según el calendario del código; si falta el año, aparece «Falta el calendario de feriados de …» (ver F8).
   - `Grillas de esas ventanas`: una línea por período y hoja, con `0 cargadas`, `0 a revisar` y `0 con Calidad, Participación u Observaciones cargadas`, salvo valores que el Equipo EPS haya puesto a propósito. `CURSA` es el total de calificaciones pendientes de esa hoja.
   - `Problemas`: «… cursantes y ningún docente asignado» para las materias reales (es esperable: durante la prueba sólo está asignada la cuenta de prueba) y «Los correos no están activados». **Ningún otro**: un DNI repetido, una calificación que la carga no reconoce o un bloque incompleto se corrigen en la grilla, en la fila indicada, antes de la prueba.
3. Ejecutar `previsualizarCorreosCarga`: `"simulados":0` (la ventana todavía no empezó).

---

## C. Prueba integrada controlada

### C0. Alcance y salvaguardas

La prueba usa los recursos definitivos y escribe en celdas reales, pero:

- **Sólo la cuenta de prueba tiene enlace y recibe correos**, porque `Asignaciones` contiene únicamente sus filas.
- **Toda escritura pasa por la app** y queda en `Registro` antes de tocar la grilla. La restauración también se hace desde la app. Así, las filas del Registro son la lista completa de celdas afectadas.
- Se escribe sólo en **una materia** y en **uno o dos estudiantes** que hoy tienen `CURSA` y los complementarios vacíos, en las dos grillas. Al terminar, cada celda vuelve a su valor original.
- No se cambia la forma de ninguna grilla (filas ni columnas). La única edición directa en Sheets es la del paso C2.13, en una celda que después se restaura desde la app.
- No se borra ninguna fila de `Registro` ni de `Correos`. Las de la prueba se identifican por la cuenta de prueba y las fechas.

**Calendario**: el día 1 y el día 2 son días hábiles consecutivos, y el día 3 es el día hábil siguiente (opcional, para el correo de cierre). Cada día se empieza después de las 8:00. El día 1 lleva una o dos horas; los días 2 y 3, unos 20 minutos cada uno. En las hojas, `<día n>` es la fecha de ese día.

### C1. Preparación (día 0)

1. **Cuenta de prueba A**: un buzón del dominio de los docentes (`@bue.edu.ar`), de alguien del Equipo EPS o de un docente que aceptó participar y puede revisar su correo. Va a recibir nombres reales de estudiantes de una materia. Opcional: **B**, la propia cuenta propietaria, como segundo docente (sirve de respaldo si el correo a A no llega).
2. **Materia de prueba M**: una materia de `EPSI` o `EPSG` que esté en las dos grillas, con pocos cursantes. **Estudiante E1** (y opcionalmente **E2**): cursa M con `CURSA` en 2º cuatrimestre y en anual, y tiene Calidad, Participación y Observaciones vacías. Anotar en el acta orientación, nivel, materia, DNI y fila de cada uno, y los valores originales de sus cuatro celdas en cada grilla.
3. **Respaldo**: en la grilla de 2º cuatrimestre, en la de anual y en el generador, *Archivo > Historial de versiones > Asignar nombre a la versión actual*: «Antes de la prueba integrada <fecha>».
4. **`Asignaciones`**: una fila `Orientación | Nivel | Materia | Docente | Correo` para A con M, con el nombre de la materia como en el encabezado de la grilla y el nivel de su sección (`EPS1`…`EPS4`). Opcional: una fila para B con otra materia M2.
5. **`Ventanas`**: `2º cuatrimestre | <día 1> | <día 2>` y `Anual | <día 1> | <día 2>`.
6. **Calendario**: los días 1 a 3 son hábiles y su año figura en `Calendario de feriados` (primera línea de `verificarCargaDocente`).
7. Hacer **B4**. Avisar al Equipo EPS que entre el día 1 y el día 2 nadie edite, en las grillas, las celdas de E1 y E2 en la materia M.

### C2. Día 1: correo, acceso, guardado, generador y revocación

Cada paso tiene un identificador para el acta.

**Correo de apertura**

| Id | Hacer | Esperado |
|---|---|---|
| C2.1 | `previsualizarCorreosCarga` | Una línea `"correo":"simulado"` por cuenta de prueba, con `"tipos":["apertura"]` y `"periodos":["2º cuatrimestre (apertura)","Anual (apertura)"]`. Ningún otro destinatario |
| C2.2 | `enviarCorreosCarga`, desde el editor | La ejecución termina sin error; `"enviados"` es el número de cuentas de prueba. En `Correos`, una fila por cuenta: `Estado` `enviado`, `Tipo` `apertura`, `Acceso` `<día 1>\|v1` |
| C2.3 | **[G]** En el buzón de A (buscar también en Spam) | Llega a la bandeja de entrada, remitente «Equipo EPS» con la dirección de la cuenta propietaria y asunto «Carga de valoraciones EPS: abrió 2º cuatrimestre y Anual · N calificaciones pendientes, cierra el d/m». En *⋮ > Mostrar original*, anotar SPF, DKIM y DMARC, la carpeta y la demora |
| C2.4 | `enviarCorreosCarga` otra vez | No se envía nada (`"yaRecibieron"` = número de cuentas de prueba) y `Correos` no cambia |

**Acceso**

| Id | Hacer | Esperado |
|---|---|---|
| C2.5 | Abrir «Abrir mi carga» desde el correo, en la computadora | La página saluda a A, muestra «2º cuatrimestre» y «Anual», y en cada uno la materia M con tantos estudiantes como `CURSA` (más valores del Equipo EPS) tiene esa columna en la grilla del período |
| C2.6 | **[G]** Abrir el enlace desde la app de Gmail en el celular y en un navegador con varias cuentas de Google iniciadas | Abre y muestra lo mismo, sin pedir iniciar sesión (con «Cualquier persona»). Anotar cualquier aviso |
| C2.7 | Cambiar un carácter del parámetro `k` del enlace; borrarlo; recortarlo | «Acceso no verificado» en los tres casos |
| C2.8 | Opcional, asignaciones vigentes: agregar en `Asignaciones` una fila de A con otra materia, y en la página tocar *Actualizar* | Aparece la materia nueva. Abrirla, borrar la fila en `Asignaciones` e intentar cambiar una calificación: «✕ No se guardó» con «Esta materia no está asignada a tu cuenta…». La grilla y el Registro no cambian |

**Guardado y corrección (2º cuatrimestre, E1, materia M)**

| Id | Hacer | Esperado |
|---|---|---|
| C2.9 | Calificación: de «Pendiente» a `7` | «✓ Guardado hh:mm». La celda de la grilla muestra `7`. En `Registro`, una fila: `Estado` `aplicado`, `Anterior` `CURSA`, `Nuevo` `7`, `Correo` de A, `Período` `2º cuatrimestre`, `Fuente` el ID de la grilla de 2º cuatrimestre, con su DNI y su materia |
| C2.10 | Corregir a `8` | «✓ Guardado»; otra fila `aplicado` con `Anterior` `7` y `Nuevo` `8` |
| C2.11 | Calidad `Buena`, Participación `Activa`, Observaciones «Prueba integrada» (escribir y salir del campo) | «✓ Guardado» en cada uno; tres filas `aplicado` y las tres celdas de la grilla con esos valores |

**Independencia de los períodos (anual, E1, materia M)**

| Id | Hacer | Esperado |
|---|---|---|
| C2.12 | En «Anual», la calificación de E1 en M sigue «Pendiente». Cambiarla a `9` | «✓ Guardado». En la grilla de anual, `9`; en la de 2º cuatrimestre sigue `8`. Fila del Registro con `Período` `Anual` y la `Fuente` de anual |

**Conflicto y falta de conexión (2º cuatrimestre, E1)**

| Id | Hacer | Esperado |
|---|---|---|
| C2.13 | Con la página abierta, en la grilla de 2º cuatrimestre (como Equipo EPS) escribir «Cambio EPS» en las Observaciones de E1 en M. En la página, sin actualizar, cambiar Observaciones a «Prueba 2» y salir del campo | «⚠ Cambió en la grilla», con el aviso: la grilla ahora dice «Cambio EPS» (…). Tu valor «Prueba 2» no se guardó. Ninguna fila nueva en el Registro. Elegir el botón *Guardar «Prueba 2»*: «✓ Guardado» y una fila `aplicado` con `Anterior` «Cambio EPS» |
| C2.14 | Cortar la conexión (en Chrome, *DevTools > Network > Offline*, o modo avión en el celular), cambiar Participación a `Poca` y esperar unos segundos. Reconectar | Mientras no hay conexión: «Reintentando…» o «⚠ Sin respuesta del servidor». Al reconectar: «✓ Guardado» o «= Ya figura así en la grilla». En el Registro, **una sola** fila para ese cambio (`aplicado` o `inferido`) y la celda con `Poca` |

Un estado `incierto` del servidor no se puede provocar a voluntad en Google real; lo cubren las pruebas locales (`tests/carga.js`) y el panel `/dev` del servidor local. Si aparece en la prueba, anotar la fila del Registro y seguir con C3.6.

**Boletín y exportación desde las mismas grillas**

| Id | Hacer | Esperado |
|---|---|---|
| C2.15 | En el generador, anotar `Control!B5`, ponerlo en `Anual` y ejecutar *Informes EPS > Generar boletines por DNI…* con el DNI de E1 | El diálogo informa E1 como generado y da el link del libro, en *Boletines EPS > <ciclo> - Anual > Por DNI*. En el boletín, la fila de M muestra `8` en 2º cuatrimestre (columna H) y `9` en anual (columna I) |
| C2.16 | *Informes EPS > Exportar PDF desde un libro o carpeta…* con el link de ese libro: *Analizar* y después *Exportar* | Un PDF de E1 en una carpeta nueva junto al libro, con los mismos valores |
| C2.17 | Volver `Control!B5` a su valor. Mover a la papelera **sólo** el libro de C2.15 y la carpeta de PDF de C2.16 (los anotados en el acta). No vaciar la papelera ni eliminar nada de forma permanente | Ningún archivo de prueba en *Boletines EPS*; Drive vuelve a eliminarlos solo de la papelera, según su plazo |

**Restauración desde la app**

| Id | Hacer | Esperado |
|---|---|---|
| C2.18 | En 2º cuatrimestre, E1, M: Calificación «Pendiente», Calidad y Participación «—», Observaciones vacía. En anual: Calificación «Pendiente». Lo mismo con E2, si se usó | «✓ Guardado» en cada uno, con una fila `aplicado` por cambio. Las celdas vuelven a los valores originales anotados en C1 |

**Revocación y activador**

| Id | Hacer | Esperado |
|---|---|---|
| C2.19 | En `Enlaces`, agregar `<correo de A> \| 2`. En la página todavía abierta, cambiar Calidad de E1 a `Buena` | «✕ No se guardó» con el motivo de acceso; la página queda bloqueada. La grilla y el Registro no cambian. Al recargar el enlace: «Acceso no verificado» |
| C2.20 | `activarCorreosCarga` | «Correos activados (activador …). Corre todos los días entre las 8:00 y las 9:00…». En *Activadores*, uno solo: `enviarCorreosCarga`, basado en tiempo, diario |
| C2.21 | `verificarCargaDocente` y `resolverPendientesCarga` | Entre los problemas, sólo las materias sin docente (esperables, como en B4): ninguna fila `pendiente` ni `revisar` del Registro y ningún aviso de correos. `resolverPendientesCarga` devuelve `"revisadas": 0` |

### C3. Día 2: seguimiento automático, enlace nuevo y cierre efectivo

Después de las 9:15:

| Id | Hacer | Esperado |
|---|---|---|
| C3.1 | **[G]** *Ejecuciones* del proyecto | Una ejecución de `enviarCorreosCarga` de tipo activador, *Completada*, entre las 8:00 y las 9:00. Anotar la hora |
| C3.2 | `Correos` | Fila del día 2 para A: `enviado`, `Tipo` `seguimiento, enlace`, `Acceso` `<día 1>\|v2` |
| C3.3 | Buzón de A | Correo con «Tu enlace de acceso cambió: usá el de este correo…» y las calificaciones pendientes |
| C3.4 | Abrir el enlace nuevo; abrir el del día 1 | El nuevo abre la página con M; el del día 1, «Acceso no verificado» |
| C3.5 | **Cierre efectivo.** Con la página abierta en 2º cuatrimestre (enlace nuevo), cambiar en `Ventanas` el `Hasta` de las dos filas a `<día 1>`. Así el cierre pasa a las 8:00 del día 2, que ya ocurrió. En la página, cambiar la calificación de E1 a `7` | «✕ No se guardó» con «La carga de 2º cuatrimestre cerró.»; la grilla y el Registro no cambian. Al recargar: «No hay cargas abiertas». `verificarCargaDocente` ya no lista las ventanas |
| C3.6 | Revisar el Registro con una vista de filtro por el correo de A | Todas las filas de la prueba, cada una `aplicado` (o `inferido`, por C2.14), sin `pendiente`, `incierto` ni `revisar`. Anotar en el acta el rango de filas |

### C4. Día 3 (opcional): correo de cierre

Después de las 9:15: en `Correos`, fila del día 3 para A con `Tipo` `cierre`. En el buzón: «Cerró la carga de 2º cuatrimestre y Anual…», con el resumen final y **sin enlace**. Si se omite el día 3, el correo de cierre queda por verificar en el primer cierre real (anotarlo).

---

## D. Limpieza y estado posterior

| Id | Hacer | Esperado |
|---|---|---|
| D1 | `desactivarCorreosCarga` | «Activadores de correos borrados: 1.»; *Activadores* vacío |
| D2 | Borrar las filas de la prueba en `Ventanas` y `Asignaciones` (que queden sólo los encabezados) y la de A en `Enlaces` | Hojas como en A8. Borrar la fila de `Enlaces` es seguro: los enlaces de la prueba son del tramo que empezó el día 1 y no vuelven a valer cuando empieza una ventana con otro «Desde» |
| D3 | En `Registro` y `Correos`, **no borrar ni editar** las filas de la prueba. En una columna agregada a la derecha («Nota EPS»), marcarlas «Prueba integrada #29» | El registro de auditoría queda completo e identificado |
| D4 | Para cada celda que figura en las filas de A del Registro, comparar su valor actual con el `Anterior` de la primera fila de esa celda. En la grilla, *Archivo > Historial de versiones*, comparar con la versión «Antes de la prueba integrada…» | Todas con su valor original. En el historial, las celdas de E1 y E2 en M sólo tienen, después de esa versión, los cambios de la prueba (incluido el de C2.13). Si alguna difiere, restaurarla a mano y anotarlo |
| D5 | Generador y Drive | `Control!B5` como estaba; el libro y la carpeta de PDF de la prueba, en la papelera; ningún otro archivo tocado |
| D6 | `verificarCargaDocente` | Como en B1: sin ventanas ni grillas, y sólo «Los correos no están activados» |
| D7 | `previsualizarCorreosCarga`; abrir los enlaces de los días 1 y 2 | `"simulados":0`; los dos enlaces dan «Acceso no verificado» |

**Queda en producción**: la app publicada (sin acceso), las propiedades, las filas identificadas de la prueba en `Registro` y `Correos`, y las versiones con nombre de las grillas. **No queda**: ventanas, asignaciones, activador ni archivos de boletín de la prueba.

**Para pasar a la etapa E** hacen falta en el acta:

- C2.3 en la bandeja de entrada (no en Spam), con SPF, DKIM y DMARC anotados;
- C2.5, C2.6, C2.9 a C2.14, C2.18, C2.19, C3.1 a C3.5 y D4 correctos;
- ninguna fila `revisar`, `incierto` o `pendiente` sin explicar.

Si el correo cae en Spam o es rechazado, la causa probable es la autenticación del dominio remitente (SPF y DKIM de la cuenta propietaria): la corrige su administrador, no el producto. Repetir C2.2 a C2.4 con otra cuenta de prueba (otro día, por la regla de un correo diario) después de la corrección.

---

## E. Apertura de la ventana institucional

**Requisito**: acta de la prueba aprobada y **autorización expresa del Equipo EPS** con las fechas de la ventana. Anotar en la ficha quién autorizó, cuándo y las fechas.

| Id | Hacer | Esperado / comprobar |
|---|---|---|
| E1 | Ejecutar `prepararPeriodosCarga` una última vez (A3.1, sin ventanas cargadas) y copiar la lista de `Asignaciones (borrador)` a `Asignaciones` (las columnas `Cursantes`, `Origen del correo` y `Revisión` se pueden copiar o no: la carga las ignora) | Una fila por orientación y materia, con un único docente; materia como en el encabezado de la grilla (mayúsculas y acentos no importan; letras y números sí); nivel de su sección; correo institucional del docente. En el borrador, `Revisión` vacía en todas las filas |
| E2 | Revisar la primera línea de `verificarCargaDocente` | `Calendario de feriados` incluye el año de la ventana, de su día de cierre y de los 7 días siguientes. Si no, actualizar el calendario (F8) antes de seguir. Revisar también que no haya feriados nuevos decretados para esas fechas |
| E3 | Cargar en `Ventanas` las filas de 2º cuatrimestre y anual con las fechas autorizadas, **al menos un día hábil antes del «Desde»** | La carga abre sola a las 0:00 de «Desde». Hasta entonces nadie tiene enlace |
| E4 | `verificarCargaDocente` | Las dos ventanas `(futura)`, con el día de cierre esperado; grillas con `0 cargadas` y `0 a revisar` (salvo valores intencionales); **ningún problema**, salvo «Los correos no están activados». En especial, ninguna materia con «cursantes y ningún docente asignado», salvo que sea intencional (nadie podrá cargarla) |
| E5 | Revisar el acceso a las grillas de 2º cuatrimestre y anual | Sólo el Equipo EPS y la cuenta propietaria, como en A3 |
| E6 | `activarCorreosCarga` | Un activador, como en C2.20 |
| E7 | Avisar a los docentes, por el canal habitual, que el primer día hábil de la ventana, entre las 8:00 y las 9:00, recibirán desde «Equipo EPS» (dirección de la cuenta propietaria) un correo con su enlace personal, que no deben reenviar | — |
| E8 | El primer día hábil de la ventana, después de las 9:15: *Ejecuciones* y `Correos` | Ejecución *Completada*; una fila `enviado` de tipo `apertura` por docente asignado. Si la ejecución falló o hay filas `error`, ejecutar `enviarCorreosCarga` a mano ese mismo día: completa lo que faltó sin repetir |

Si E4 muestra problemas que no se pueden resolver antes del «Desde», correr las fechas en `Ventanas` (con la ventana todavía futura, no cambia ningún enlace ya enviado).

---

## F. Operación, cierre y mantenimiento

### F1. Durante la ventana

- **Cada día hábil, después de las 9:15**: en `Correos`, que no haya filas `error` del día; si una ejecución falla, Google avisa por correo a la cuenta propietaria. Ante errores o envíos sin procesar, `enviarCorreosCarga` a mano el mismo día.
- **Una o dos veces por semana**: `verificarCargaDocente`. Las filas `pendiente` del Registro con más de 10 minutos se resuelven con `resolverPendientesCarga`. Las `revisar` y las `incierto` se atienden como en F3.
- **No** insertar, borrar ni ordenar filas o columnas de las grillas, ni cambiar `Control!B13`/`B14`. Editar valores en Sheets sí se puede: la carga no pisa lo que cambió desde que el docente lo vio.

### F2. Cambios habituales

| Situación | Qué hacer | Efecto |
|---|---|---|
| Cambiar el docente de una materia | Editar `Docente` y `Correo` de su fila en `Asignaciones` | El anterior la pierde en su operación siguiente. El nuevo recibe el correo de apertura en el próximo envío diario (ese mismo día, si el cambio se hace antes de las 8:00) |
| Agregar un docente o una materia | Nueva fila en `Asignaciones` | Apertura en el próximo envío diario |
| Un docente reenvió su enlace o lo perdió | En `Enlaces`, agregar su correo con `Versión` 2, o aumentar la que tenga | El enlace anterior deja de valer enseguida. El nuevo llega en el próximo envío diario; si ese día ya recibió su correo, al día hábil siguiente (a lo sumo un correo por día). **No** borrar filas ni bajar versiones durante la ventana |
| Extender la ventana | Cambiar sólo `Hasta`, en las dos filas si comparten fechas | El cierre se recalcula y los enlaces no cambian. **No** cambiar el `Desde` de una ventana abierta: cambia todos los enlaces y los nuevos llegan recién en el próximo envío diario |
| El Equipo EPS corrige una valoración | Editarla en la grilla | Si un docente la tenía abierta, ve «Cambió en la grilla» y decide |
| «No me llegó el correo» | Buscar al docente en `Correos`. Sin fila: revisar su correo en `Asignaciones` y que sea día hábil. Con `error`: `enviarCorreosCarga`. Con `enviado`: pedirle que busque en Spam | — |
| «Acceso no verificado» | Que use el enlace del **último** correo; que su correo figure en `Asignaciones`; que su versión en `Enlaces` sea un número | — |

### F3. Incidencias del Registro

- **`revisar`**: el `Detalle` indica la celda. Abrirla, ver *clic derecho > Mostrar historial de ediciones*, decidir el valor con el docente y corregirlo en la grilla. Anotar la resolución en la columna «Nota EPS», **sin** editar `Estado` ni `Detalle`.
- **`incierto`**: la escritura pudo no aplicarse, o aplicarse y ser corregida. La celda muestra el valor vigente; el docente decide si vuelve a cargar.
- **`pendiente`** viejas: `resolverPendientesCarga`.

### F4. Recuperación

| Problema | Qué hacer |
|---|---|
| Un valor erróneo en una celda | El Registro da el `Anterior`; la celda, su historial de ediciones. Corregir en la grilla |
| Restaurar una grilla entera | Sólo con una versión con nombre y si nadie editó después; si no, celda por celda |
| Se filtró `CLAVE_ENLACES` (por ejemplo, se copió a un documento) | Borrar la propiedad y ejecutar `prepararAdministracionCarga`: todos los enlaces dejan de valer. Con una ventana abierta, aumentar además la versión de **cada** docente en `Enlaces`: los correos no detectan el cambio de secreto, y sin eso quien no tiene pendientes no recibiría el enlace nuevo |
| Falta el activador o hay más de uno (lo informa `verificarCargaDocente`) | `activarCorreosCarga` |
| Cuota diaria de correo agotada (envíos «sin procesar») | Se completan al día siguiente, o con `enviarCorreosCarga` cuando Google renueve la cuota |
| La app no responde | Ver *Ejecuciones*. La carga se puede interrumpir sin perder nada: lo guardado está en la grilla y en el Registro |

### F5. Cierre

- Automático: a las 8:00 del primer día hábil posterior a `Hasta`, nadie puede escribir, aunque tenga la página abierta. Ese día, entre las 8:00 y las 9:00, sale el correo de cierre a todos los asignados; si no sale, se recupera en los 7 días siguientes.
- **Comprobar**: en el Registro, ninguna fila con fecha posterior al cierre; en `Correos`, una fila `cierre` por docente.
- Después: `desactivarCorreosCarga` (opcional; sin ventanas abiertas no envía nada). Las correcciones fuera de término las hace el Equipo EPS en la grilla. Los boletines se generan y exportan como siempre (README, *Puesta en funcionamiento*), desde las mismas grillas.

### F6. Actualizar el código

Preferentemente sin ventanas abiertas:

1. Anotar el commit nuevo en la ficha.
2. Pegar los archivos cambiados de `src/carga/` y guardar. **El activador usa enseguida el código guardado; la app web, no** hasta el paso 3.
3. *Implementar > Gestionar implementaciones*, editar la implementación existente, *Versión: Nueva versión*, *Implementar*. **No** crear una implementación nueva: cambiaría la URL y los enlaces ya enviados apuntarían a la anterior.
4. Repetir B1 y B3 (y, si cambió algo del guardado, una prueba corta como C2.9, C2.10 y C2.18, antes de abrir).

### F7. Períodos y ciclos siguientes

- **Cada período**: preparar su grilla como en A3 (copia de la del período anterior, con `CURSA` donde corresponda; la preparación automática de A3.1 es sólo para 2º cuatrimestre y anual copiados de 1º cuatrimestre, con evidencia de 3º bimestre), ponerla en `Control!B10:B14` del generador sin ventanas abiertas, actualizar `Asignaciones` y repetir la etapa E. Cada período tiene una sola fila en `Ventanas`; se reemplaza la anterior.
- **Bimestres** (1º y 3º): escala `En proceso`, `Suficiente`, `Avanzado`. La carga los admite, pero nunca se probaron en Google: la primera vez, hacer antes una prueba corta (C2.5, C2.9, C2.10, C2.18 y D4) con la cuenta de prueba.
- **Ciclo nuevo**: antes de la primera ventana del año, actualizar el calendario de feriados (F8). Al preparar el ciclo, se pueden depurar las filas de `Enlaces`: los enlaces de la ventana anterior vencen cuando empieza la siguiente. `Registro` y `Correos` se conservan.

### F8. Actualizar el calendario de feriados

Los feriados forman parte del código: `src/carga/Carga_Calendario.gs`, una entrada por año. Día hábil: de lunes a viernes, salvo los `feriados` del año. Los `noLaborables` se anotan aparte y **no** cambian los días hábiles (criterio institucional). Por ejemplo, el lunes 7/12/2026, no laborable con fines turísticos, es hábil para la carga: una ventana con `Hasta` el viernes 4/12 cierra ese lunes a las 8:00 y ese día sale el correo de cierre. Si el Equipo EPS decide que un no laborable debe contar, se pasa a `feriados`.

**Cuándo**: al preparar cada ciclo lectivo (antes de su primera ventana), y durante el año si un decreto agrega un feriado nacional o de CABA.

**Cómo** (responsable técnico):

1. En el repositorio, agregar o corregir la entrada del año en `Carga_Calendario.gs`, desde fuentes oficiales: la Ley 27.399 (feriados inamovibles y trasladables, y la regla de traslado), los decretos o resoluciones del año (días con fines turísticos, feriados extraordinarios) y lo que rija en CABA. Los trasladables van en su **fecha efectiva**. Anotar las fuentes en el comentario del año. No agregar años sin datos oficiales.
2. No borrar años pasados mientras `Ventanas` tenga ventanas de ese año.
3. Correr las pruebas (`node tests/carga.js`, que valida las fechas del calendario, y las demás del README).
4. En el proyecto de Apps Script, reemplazar el contenido de `Carga_Calendario` y publicar una versión nueva de la implementación existente (F6). El activador de los correos usa el archivo apenas se guarda; la app web, recién con la versión nueva.
5. **Comprobar**: la primera línea de `verificarCargaDocente` lista el año, y las ventanas muestran el día de cierre esperado.

**Si falta un año**, la carga no adivina. Si el cierre de un período cae en ese año, el período no abre, y `verificarCargaDocente` muestra «Falta el calendario de feriados de …». Si ese día hay correos que informar, no sale ninguno y la ejecución termina con error, que Google avisa por correo a la cuenta propietaria. Lo que no depende del calendario sigue funcionando: sin ventanas en curso ni cierres por informar, los correos no lo necesitan.

---

## Anexo 1. Pendiente de verificación con Google real

Ninguno de estos puntos se comprobó. La etapa C cubre los marcados con su paso:

| Punto | Paso |
|---|---|
| Autorización de los cuatro permisos en la cuenta propietaria y restricciones del administrador | A7 |
| Modalidad de acceso disponible («Cualquier persona») y forma de la URL `/exec` | A9, B3 |
| Apertura sin iniciar sesión, desde el celular (app de Gmail) y con varias cuentas | B3, C2.6 |
| Firma de los enlaces con `Utilities.computeHmacSha256Signature`: el enlace que emite el correo es el que acepta la app | C2.5 |
| Entrega real a `@bue.edu.ar`: carpeta, remitente, SPF, DKIM, DMARC | C2.3 |
| Escritura, `flush` y lock reales; un guardado por fila del Registro; reintento sin duplicar | C2.9 a C2.14 |
| Ejecución desatendida del activador, autorización por `triggerUid` y minuto real | C3.1 |
| Correo de cierre | C4 (opcional) o primer cierre real |
| Cuota real de MailApp con el número real de docentes, y duración de la ejecución diaria | E8 |
| Latencia y contención con muchos docentes a la vez | Primeros días de la ventana real (F1) |
| Bimestres en Google | F7 |
| Calendario instalado: la primera línea de `verificarCargaDocente` muestra 2026 y el cierre de la prueba cae donde se espera | B1, B4 |
| Estado `incierto` del servidor | No se puede provocar; sólo pruebas locales |
| Preparación (#32): lectura real de las protecciones (`getEditors`) con la cuenta propietaria, escritura por `getRangeList`, duración con las grillas reales, y que el informe coincida con lo que el Equipo EPS sabe de los borradores | A3.1 (vista previa revisada antes de aplicar) |

## Anexo 2. Acta de la prueba

Copiar en un documento del Equipo EPS (no en el repositorio):

```
Prueba integrada de la carga docente · fechas de los días 0 a 3:
Commit instalado:                 Implementación / versión:
Modalidad de acceso (A9):         Cuenta propietaria:
Cuentas de prueba (A, B):         Materia M / E1 / E2 (orientación, nivel, filas):
Valores originales (2C y anual, cuatro celdas por estudiante):

Id     Resultado (OK / falla)   Observación (texto exacto, hora, captura)
A7
B1 … B4
C2.1 … C2.21
C3.1 … C3.6
C4
D1 … D7
Filas de la prueba en Registro:            en Correos:
Decisión: pasar a la etapa E (sí / no) · quién · fecha
```
