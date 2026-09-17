Generador de informes individuales EPS - v0.1

Archivos del paquete:

1. Generador_informes_individuales_EPS_v0_1.xlsx
   Archivo base del generador. Al subirlo a Google Drive y abrirlo como Google Sheets, contiene:
   - Control
   - Plantilla
   - Informe_actual

2. Codigo_Apps_Script_Generador_EPS_v0_1.gs
   Código para pegar en Extensiones > Apps Script dentro del archivo generador.

3. Valoraciones_ficticias_EPS_v0_1.xlsx
   Planilla fuente de prueba, con estudiantes ficticios y estructura equivalente a la planilla de valoraciones.
   DNI de prueba:
   - EPSI: 99000001, 99000002, 99000003, 99000004
   - EPSG: 98000001, 98000002, 98000003

Uso básico:

1. Subir Generador_informes_individuales_EPS_v0_1.xlsx a Google Drive.
2. Abrirlo como Google Sheets.
3. Ir a Extensiones > Apps Script.
4. Pegar el código del archivo Codigo_Apps_Script_Generador_EPS_v0_1.gs.
5. Guardar.
6. Volver a la planilla.
7. En Control, completar:
   - URL o ID de la planilla de valoraciones.
   - Hoja origen: EPSI o EPSG.
   - DNI del estudiante.
8. Ejecutar el menú Informes EPS > Generar informe actual.
9. Revisar Informe_actual antes de imprimir o exportar.

Alcance de esta versión:

- Genera un informe individual por vez.
- Completa únicamente 1º bimestre.
- Usa DNI como identificador.
- No modifica la planilla fuente.
- No modifica la hoja Plantilla.
- Incluye materias con En proceso, Suficiente, Avanzado o ACREDITÓ.
- Excluye celdas vacías, NO CURSA y CURSA.
- Muestra advertencias ante valores inesperados.
- Deja el resultado en Informe_actual para revisión manual.

Notas:

- El archivo .xlsx no puede llevar Apps Script embebido. Por eso el código va separado.
- Para probar con datos ficticios, subir también Valoraciones_ficticias_EPS_v0_1.xlsx a Drive, abrirlo como Google Sheets y usar su URL como fuente.
