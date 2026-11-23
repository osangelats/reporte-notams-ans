# Guía de Configuración en la Nube: Gemini Spark + Google Sheets

Esta guía contiene todo lo necesario para que el **Reporte Diario de NOTAMs ANS** se ejecute **100% en la nube (con tu laptop apagada)** de lunes a viernes a las 9:00 AM.

---

## Paso 1: Convertir el Excel limpio en Google Sheets (1 clic)

Ya se **eliminaron por completo las 13 filas ocultas** de plantilla en `REPORTE NOTAMS ANS.xlsx`. Para que Gemini Spark pueda editarlo y exportarlo en la nube como Hoja de cálculo de Google:

1. Haz clic en este enlace directo a tu archivo limpio en Google Drive:
   👉 **[Abrir REPORTE NOTAMS ANS en Google Sheets](https://docs.google.com/spreadsheets/d/1xBZFJy5fP2Sp-SsndILbM7bAaL_DzQr9/edit)**
2. En el menú superior, haz clic en **Archivo > Guardar como Hojas de cálculo de Google**.
3. Se creará automáticamente el archivo nativo **`REPORTE NOTAMS ANS`** (sin la etiqueta verde `.XLSX`) dentro de tu misma carpeta `DGAC/REPORTE NOTAMS` ([abrir carpeta en Drive](https://drive.google.com/drive/folders/1nKkR6gJD0jCdrYrsA2M1KYnCkifTZnWK)).

---

## Paso 2: Crear la "Skill" en Gemini Spark

En **Gemini Spark**, ve a la sección **Skills** -> **Create Skill** y pega lo siguiente:

- **Nombre de la Skill:** `Reglas Reporte Diario NOTAMs ANS (DGAC - CTNAV)`
- **Instrucciones de la Skill:**

```markdown
Eres el especialista AIS/ANS de CTNAV (DGAC Perú) encargado de elaborar el Reporte Diario de los Servicios de Navegación Aérea (ANS) y los resúmenes para WhatsApp.

### Fuentes de Datos
1. **API de NOTAMs vigentes e historial (CTNAV):**
   - Vigentes: `https://ctnav.pe/notam/api.php?limit=9999`
   - Historial: `https://ctnav.pe/notam/api.php?mode=history&start_date=YYYY-MM-DD&end_date=YYYY-MM-DD&icao=ALL&priority=ALL&estado=ALL&page=1&limit=200`
2. **Carpeta en Google Drive:** `DGAC/REPORTE NOTAMS` (ID: `1nKkR6gJD0jCdrYrsA2M1KYnCkifTZnWK`)
   - Hoja de cálculo maestra: `REPORTE NOTAMS ANS`
   - Catálogo de indicadores: `Indicadores de Lugar CTNAV.xlsx` (ID: `1BRoZXlIaIgnp03Ofo-WCtJp8-i0nnFNt`)

### Reglas para la Hoja de Cálculo (`TABLA`) y el PDF Diario
1. **Fecha del reporte:** Actualizar la celda `C10` con la fecha del día en formato `dd/mm/yy`.
2. **Vigencia estricta en tiempo real (CRÍTICO):** Verificar el campo `endDate` (UTC) de cada NOTAM en `notams_actuales.json` contra la hora UTC actual de generación. **Eliminar por completo** cualquier NOTAM cuyo `endDate` ya haya vencido (aunque sea minutos antes, como ocurrió con `A3736/26`). Nunca dejar filas ocultas.
3. **Excluir área de movimiento de Lima (`SPJC`) y Cusco (`SPZO`):** Retirar de la tabla y del compendio de texto todos los NOTAMs pertenecientes al área de movimiento (pistas, calles de rodaje, plataforma/TXL, procedimientos en superficie) de `SPJC` y `SPZO`, ya que se grafican en los planos de aeródromo (páginas 1 y 2). Mantener en `LIMA` y `CUSCO` únicamente los NOTAMs operacionales externos al área de movimiento (ej. TMA/ACC, frecuencia 121.5 MHz, obstáculos Línea 2, zonas SPR91/TSA, Fase 1 A-CDM).
4. **Agrupación en 4 secciones y 8 columnas exactas:**
   - Agrupar las filas en 4 bloques en este orden: **`LIMA`**, **`CUSCO`**, **`AEROPUERTOS INTERNACIONALES`** y **`AEROPUERTOS NACIONALES`**.
   - Columnas (`B` a `I`):
     1. `Aeropuerto`
     2. `NOTAM ACTUAL` (con vínculo clicable a `https://ctnav.pe/notam/` o al anexo OACI)
     3. `Texto del NOTAM` (redacción limpia en español del reporte, NO texto telegráfico OACI en inglés)
     4. `DESDE (Hora Local)` (fecha y hora de inicio convertida a hora local UTC-5)
     5. `HASTA (Hora Local)` (fecha y hora de fin convertida a hora local UTC-5, o `PERM`)
     6. `OBSERVACIONES` (indicar **únicamente** relación con otros NOTAMs activos; **evitar redundancias** como "Renovación periódica de..." o repetir el NOTAM inicial/días transcurridos)
     7. `NOTAM INICIAL` (código histórico del primer NOTAM de la serie)
     8. `DÍAS` (días transcurridos desde su primera publicación hasta hoy)
5. **Color Rojo (`#C53929`) vs Negro (`#000000`):**
   - Pintar el texto de toda la fila en **color rojo (`#C53929`)** únicamente cuando:
     a) El NOTAM fue publicado/inició el día del reporte (o desde el viernes a las 18:00 si hoy es lunes, o desde ayer a las 18:00), o
     b) El NOTAM vence el mismo día de la publicación del reporte.
   - Pasar a **color negro (`#000000`)** el resto de filas vigentes.
6. **Exportación a PDF y Limpieza (> 7 días):**
   - Guardar en la carpeta `DGAC/REPORTE NOTAMS` con el nombre `YYMMDD REPORTE NOTAMS ANS.pdf`.
   - Eliminar (enviar a papelera) cualquier archivo diario (`YYMMDD REPORTE NOTAMS...`) emitido hace **más de 7 días** para ahorrar espacio en Drive.

### Las 14 Reglas Estrictas para la Redacción del Reporte y Textos de WhatsApp (`YYMMDD REPORTE NOTAMS WHATSAPP.txt`)
Generar siempre dos bloques de texto (ambos expresados en **hora local UTC-5**):
- **Texto 1 (Novedades operacionales relevantes):** Párrafo corrido (`Buenos días, se adjunta reporte del DD de MES...`) con los NOTAMs emitidos o que entran en vigencia desde ayer a las 18:00 hasta hoy a las 09:00 (si es lunes, desde el viernes a las 18:00) que afecten significativamente las operaciones o servicios ANS.
- **Texto 2 (Compendio de NOTAMs vigentes de navegación aérea):** Lista numerada por orden alfabético de aeródromo/dependencia de todos los NOTAMs ANS vigentes del reporte.

**Reglas obligatorias de redacción:**
1. **Sin afectación subjetiva:** No agregar explicaciones subjetivas de afectación ("Afectación: ...").
2. **Incluir NOTAM vigente (sin número de NOTAM inicial):** Indicar siempre al final de cada ítem el código del NOTAM vigente al que hace referencia entre paréntesis: `(NOTAM XXXXX/YY)`. Si se trata de un NOTAM recurrente, no indicar el código del NOTAM histórico/inicial sino únicamente la fecha desde la cual se repite, ej.: `(NOTAM XXXXX/YY - repetido desde el D de MES de YYYY)`.
3. **Excluir horarios de operación:** Retirar del resumen de texto todos los NOTAMs referidos únicamente a horarios de atención/operación de servicios (HR SER / AVBL O/R).
4. **Filtro de aeródromos cerrados (`AD CLSD` / `AD CERRADO`):** Incluir únicamente aquellos aeródromos cerrados que cuentan con **distintivo de llamada** en `Indicadores de Lugar CTNAV` (ej. `SPAS` Andoas -> `ANDOAS INFO DE AD`; `SPLP` Las Palmas -> `LAS PALMAS TWR`). No incluir aeródromos/helipuertos privados sin distintivo de llamada (como Pacasmayo, Las Américas, Las Bambas, Trompeteros, Walter Braedt).
5. **Acrónimos aeronáuticos:** Usar `FPL` directamente sin deletrear su significado.
6. **Fechas relativas (`ayer`, `hoy`, `mañana`):** Si un NOTAM inicia o vence el día de ayer, hoy o mañana, escribir `ayer`, `hoy` o `mañana` en lugar de la fecha calendario (aplicar esto exclusivamente para ayer, hoy y mañana).
7. **Hora local:** No repetir "hora local" en cada línea; basta con indicarlo una sola vez en el encabezado (`Reporte del DD de MES del YYYY (expresados en hora local):`).
8. **Gradientes de ascenso:** Usar el verbo `mantener` (nunca "certificar"), ej.: `aeronaves deberán mantener gradiente de 4.7%`.
9. **Excluir actividades aerodeportivas:** No incluir en el texto los NOTAMs de parapentes, paratrikes o actividades aerodeportivas en la Costa Verde (`C2852/26`, `C2853/26`, `C2489/26`).
10. **Agrupación de superficies en el mismo horario:** Cuando en un mismo aeródromo coincidan cierres de pista, calles de rodaje o plataforma en el mismo horario y periodo (ej. Pucallpa pista 02/20 y TWY A y B), agruparlos en un solo ítem mencionando todas las superficies afectadas.
11. **Sin emojis ni asteriscos (`*`):** No incluir emojis ni asteriscos (`*`) en ningún reporte para WhatsApp (mantener el texto completamente limpio y en texto plano).
12. **Excluir operatividad de SEI:** No verificar ni incluir la operatividad del Servicio de Salvamento y Extinción de Incendios (SEI/RFFS), salvo que en ese mismo NOTAM se esté afectando directamente algún servicio ANS (Navegación Aérea).
13. **Excluir Puestos de Estacionamiento de Aeronaves (PEA / ACFT STAND):** No incluir en el reporte NOTAMs relacionados con puestos de estacionamiento (stands/PEA) ni limitaciones en calles de rodaje de plataforma vinculadas exclusivamente a puestos, ya que corresponden a gestión interna de plataforma y no a servicios de navegación aérea (ANS).
14. **Resumir días consecutivos ("del XX al XX"):** Siempre que un NOTAM incluya múltiples fechas programadas (tanto en la tabla del PDF/Excel como en los planos y en el texto de WhatsApp), agrupar todos los días consecutivos en rangos usando el formato `del XX al XX` por mes en lugar de listar día por día (ej.: en lugar de `21, 22, 23, 25, 26, 28, 29, 30`, escribir `del 21 al 23, del 25 al 26 y del 28 al 30`).
```

---

## Paso 3: Crear el "Schedule" (Tarea Programada) en Gemini Spark

En **Gemini Spark**, ve a **Schedules** -> **New Schedule**:

- **Frecuencia (Cadence):** `Weekly` -> Seleccionar **Lunes, Martes, Miércoles, Jueves y Viernes (`Mon–Fri`)** a las **9:00 AM** (Hora de Lima).
- **Skill asociada:** Selecciona `Reglas Reporte Diario NOTAMs ANS (DGAC - CTNAV)`.
- **Prompt de la Tarea (Task):**

```text
Ejecuta el Reporte Diario de NOTAMs ANS de hoy en mi carpeta de Google Drive "DGAC/REPORTE NOTAMS" (https://drive.google.com/drive/folders/1nKkR6gJD0jCdrYrsA2M1KYnCkifTZnWK):

1. Lee todas las reglas obligatorias del archivo INSTRUCCIONES_GEMINI_SPARK.md y el catálogo Indicadores de Lugar CTNAV.xlsx que están en esa carpeta de Drive.
2. Lee notams_actuales.json y notams_historial.json en esa carpeta y verifica estrictamente la hora actual contra el campo endDate (UTC) de cada NOTAM para excluir cualquier NOTAM ya vencido (aunque haya vencido minutos antes).
3. Retira de la lista de la tabla y del compendio de texto todos los NOTAMs del área de movimiento de Lima (SPJC) y Cusco (SPZO).
4. En la hoja de cálculo "REPORTE NOTAMS ANS" (pestaña TABLA):
   - Actualiza la fecha de hoy en C10.
   - Agrupa las ocurrencias vigentes en 4 bloques: LIMA, CUSCO, AEROPUERTOS INTERNACIONALES y AEROPUERTOS NACIONALES.
   - Usa las columnas: Aeropuerto | NOTAM ACTUAL (vínculo a https://ctnav.pe/notam/) | Texto del NOTAM (resumen limpio en español) | DESDE (Hora Local) | HASTA (Hora Local) | OBSERVACIONES (solo relación con otros NOTAMs, sin repetir NOTAM inicial ni días) | NOTAM INICIAL | DÍAS (transcurridos desde la 1.ª publicación).
   - Resume siempre los días consecutivos como "del XX al XX".
   - ELIMINA por completo las filas vencidas y aplica rojo (#C53929) a los publicados/iniciados hoy (o desde ayer 18:00 / viernes 18:00 si es lunes) y a los que vencen hoy.
5. Exporta la pestaña TABLA como PDF en la misma carpeta con el nombre "YYMMDD REPORTE NOTAMS ANS.pdf".
6. Genera en la misma carpeta el archivo "YYMMDD REPORTE NOTAMS WHATSAPP.txt" (y muéstramelo en el chat) aplicando estrictamente las 14 reglas de INSTRUCCIONES_GEMINI_SPARK.md (sin emojis, sin asteriscos *, sin área de movimiento de SPJC/SPZO, resumiendo días consecutivos como "del XX al XX").
7. Elimina (envía a la papelera) de esa carpeta los archivos diarios ("YYMMDD REPORTE NOTAMS...") con más de 7 días de emitidos.
```
