# boletines-eps

Generador de boletines para la Escuela Profesional Secundaria (EPS) del CFP N.º 7, donde se desarrolla y se usa.

El equipo docente carga las valoraciones en grillas de Google Sheets. `boletines-eps` arma a partir de ellas el boletín de cada estudiante con el formato institucional, listo para revisar, exportar e imprimir.

```
Grillas de valoración → Generador (Google Sheets + Apps Script) → Boletín → Revisión, exportación e impresión
```

El proyecto está en desarrollo activo. La versión actual es v0.3, y el trabajo pendiente se organiza en los issues del repositorio.

## Puesta en funcionamiento

1. Subir `src/Generador_informes_individuales_EPS_v0_3.xlsx` a Google Drive y abrirlo como Google Sheets.
2. En *Extensiones > Apps Script*, pegar `src/Codigo_Apps_Script_Generador_EPS_v0_3.gs`, guardar y recargar.
3. Completar la hoja `Control` con el período, el estudiante y las grillas de origen.
4. Ejecutar *Informes EPS > Generar informe actual* y revisar el resultado antes de exportar o imprimir.

## Organización

- `src/`: versión actual del generador (planilla y Apps Script).
- `tests/`: pruebas locales.
- `reference/`: casos de prueba y datos sintéticos.
- `legacy/`: versiones anteriores, como referencia.

## Pruebas

En Windows, desde PowerShell o `cmd`, con Node.js:

```
node tests/run.js
```

Los casos cubiertos están descritos en [`reference/casos_de_prueba_v0_3.md`](reference/casos_de_prueba_v0_3.md).

## Datos

Los datos de prueba versionados son sintéticos. Las grillas reales y cualquier dato de estudiantes deben quedar fuera de Git.
