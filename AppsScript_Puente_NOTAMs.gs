/**
 * Receptor minimalista en Google Drive para el PDF Diario de NOTAMs ANS
 * - Recibe únicamente el PDF unificado desde GitHub Actions.
 * - Reemplaza si ya existe el PDF del mismo día.
 * - Elimina automáticamente (envía a papelera) reportes con más de 7 días.
 */

const FOLDER_ID = '1nKkR6gJD0jCdrYrsA2M1KYnCkifTZnWK'; // Carpeta DGAC/REPORTE NOTAMS
const WEBHOOK_SECRET = 'DGAC_CTNAV_NOTAM_2026';

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.secret !== WEBHOOK_SECRET) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const guardados = [];
    if (Array.isArray(payload.files)) {
      for (const item of payload.files) {
        const bytes = Utilities.base64Decode(item.content_b64);
        const blob = Utilities.newBlob(bytes, item.mimeType || 'application/pdf', item.name);
        const existing = folder.getFilesByName(item.name);
        while (existing.hasNext()) {
          existing.next().setTrashed(true);
        }
        folder.createFile(blob);
        guardados.push(item.name);
      }
    }
    limpiarArchivosAntiguos();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, guardados: guardados }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function limpiarArchivosAntiguos() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const limiteMs = new Date().getTime() - 7 * 24 * 60 * 60 * 1000;
  const regex = /^(\d{2})(\d{2})(\d{2})\s+REPORTE NOTAMS/i;
  while (files.hasNext()) {
    const file = files.next();
    const m = file.getName().match(regex);
    if (!m) continue;
    const fileDate = new Date(2000 + Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (fileDate.getTime() < limiteMs) {
      file.setTrashed(true);
    }
  }
}

/**
 * Dispara el workflow de GitHub Actions (workflow_dispatch) de lunes a viernes.
 * Configurar un Activador (Trigger) diario en Apps Script entre 9:00 y 10:00 AM (hora Lima)
 * y guardar el token de GitHub en Propiedades del Script bajo la clave: GITHUB_TOKEN
 */
function dispararReporteGitHub() {
  const dia = new Date().getDay(); // 0 = Dom, 6 = Sab
  if (dia === 0 || dia === 6) {
    Logger.log('Fin de semana: no se genera reporte.');
    return;
  }
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    throw new Error('Falta configurar la propiedad GITHUB_TOKEN en Configuración del proyecto > Propiedades del script.');
  }
  const url = 'https://api.github.com/repos/osangelats/reporte-notams-ans/actions/workflows/reporte_notams.yml/dispatches';
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
    },
    payload: JSON.stringify({ ref: 'main' }),
    muteHttpExceptions: true,
  });
  Logger.log('GitHub dispatch status: ' + resp.getResponseCode() + ' ' + resp.getContentText());
}
