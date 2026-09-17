Generador de informes individuales EPS v0.2
================================================

Contenido del paquete
---------------------
- Generador_informes_individuales_EPS_v0_1_1.xlsx
- Codigo_Apps_Script_Generador_EPS_v0_1_1.gs
- Generador_informes_individuales_EPS_v0_2.xlsx
- Codigo_Apps_Script_Generador_EPS_v0_2.gs
- Valoraciones_ficticias_EPS_1er_bimestre_v0_2.xlsx
- Valoraciones_ficticias_EPS_1er_cuatrimestre_v0_2.xlsx
- Casos_de_prueba_Generador_EPS_v0_2.xlsx
- Especificacion_Generador_EPS_v0_2.docx
- Manual_Generador_EPS_v0_2.docx
- CHANGELOG_Generador_EPS.txt

Instalación rápida
------------------
1. Subir Generador_informes_individuales_EPS_v0_2.xlsx a Google Drive.
2. Abrirlo como Google Sheets.
3. Abrir Extensiones > Apps Script.
4. Pegar el contenido de Codigo_Apps_Script_Generador_EPS_v0_2.gs.
5. Guardar el proyecto y recargar la planilla.
6. Completar Control.
7. Ejecutar Informes EPS > Generar informe actual.
8. Revisar Informe_actual antes de imprimir o exportar.

Reglas centrales v0.2
---------------------
- La hoja origen es única: EPSI o EPSG.
- La generación sigue siendo individual, por DNI.
- La versión implementa solo 1º bimestre y 1º cuatrimestre.
- El boletín de 1º cuatrimestre es acumulativo.
- Las celdas sin calificación válida quedan vacías.
- En 1º cuatrimestre se aceptan solo enteros del 1 al 10 y ACREDITÓ.
- No se aceptan decimales, texto junto con número ni notas escritas como texto.
