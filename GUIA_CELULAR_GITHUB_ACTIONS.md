# Ejecutar el Reporte Unificado (6 páginas) desde tu Celular sin PC prendida

Esta guía configura **GitHub Actions + Google Apps Script** (100% gratis) para que el reporte unificado completo (planos vectoriales SPJC y SPZO + tabla de 4 bloques + anexo OACI clicable + Excel + WhatsApp + limpieza > 7 días) se genere automáticamente **Lunes a Viernes a las 9:00 AM**, o cuando tú lo lances **con 1 toque desde tu celular** con la laptop apagada.

---

## Paso 1: Publicar `AppsScript_Puente_NOTAMs.gs` como Web App (2 minutos)
1. Abre tu proyecto de Google Apps Script en el navegador y reemplaza el código con el contenido actualizado de `AppsScript_Puente_NOTAMs.gs`.
2. Haz clic en **Implementar** (arriba a la derecha) -> **Nueva implementación**.
3. En **Tipo**, elige **Aplicación web**:
   - **Ejecutar como:** `Yo (tu correo)`
   - **Quién tiene acceso:** `Cualquier usuario`
4. Haz clic en **Implementar**, autoriza los permisos y **copia la URL de la aplicación web** (termina en `/exec`).

---

## Paso 2: Subir la carpeta a un repositorio privado de GitHub
1. Crea un repositorio **privado** en [github.com/new](https://github.com/new) (ej. `reporte-notams-ans`).
2. Sube los archivos de esta carpeta (incluyendo `.github/workflows/reporte_notams.yml`, `generar_reporte_unificado.py`, `assets/` y `REPORTE NOTAMS/REPORTE NOTAMS ANS.xlsx`).
3. En tu repositorio de GitHub ve a **Settings -> Secrets and variables -> Actions -> New repository secret** y crea:
   - **`DRIVE_WEBHOOK_URL`**: pega la URL `/exec` que copiaste en el Paso 1.
   - *(Opcional)* **`DRIVE_WEBHOOK_SECRET`**: `DGAC_CTNAV_NOTAM_2026`

---

## Paso 3: Cómo ejecutarlo desde tu celular (con la PC apagada)
1. Instala la app oficial **GitHub** en tu celular (iOS / Android) e inicia sesión.
2. Entra a tu repositorio `reporte-notams-ans` -> **Actions** -> **Reporte Diario NOTAMs ANS**.
3. Toca **Run workflow** (1 toque).
4. En ~25 segundos:
   - El **PDF unificado de 6 páginas**, el **Excel** y el **TXT** aparecerán automáticamente en tu carpeta de Google Drive **`DGAC/REPORTE NOTAMS`** (y también podrás descargarlos desde la misma app de GitHub en *Artifacts*).
   - Los archivos diarios con más de 7 días se eliminarán automáticamente de Drive.
