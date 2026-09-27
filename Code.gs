/**
 * Google Apps Script backend untuk website absensi.
 *
 * Cara penggunaan:
 * 1. Buat atau buka Google Spreadsheet.
 * 2. Extensions > Apps Script, lalu tempelkan file ini.
 * 3. Deploy > New deployment > Web app.
 * 4. Execute as: Me. Who has access: Anyone.
 * 5. Salin URL /exec ke Environment Variable APPS_SCRIPT_URL di Vercel.
 */

const SPREADSHEET_ID = ''; // Kosongkan jika script dibuat dari spreadsheet tersebut.
const SHEET_NAME = 'Absensi';
const FOLDER_NAME = 'Foto Absensi';
const MAX_NAME_LENGTH = 120;
const MAX_INSTITUTION_LENGTH = 160;
const MAX_STATUS_LENGTH = 60;

function doGet(event) {
  try {
    const action = event && event.parameter && event.parameter.action;
    if (action !== 'list') return json_({ ok: true, service: 'attendance', message: 'API aktif.' });

    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return json_({ ok: true, data: [] });

    const rows = values.slice(1).filter(row => row[0]).map(row => ({
      id: String(row[0]),
      name: String(row[1] || ''),
      institution: String(row[2] || ''),
      status: String(row[3] || ''),
      time: formatTime_(row[4]),
      photoUrl: String(row[6] || '')
    })).reverse();

    return json_({ ok: true, data: rows });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  }
}

function doPost(event) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const payload = JSON.parse(event.postData.contents || '{}');
    if (payload.action !== 'create') throw new Error('Aksi tidak dikenal.');

    const name = cleanText_(payload.name, MAX_NAME_LENGTH);
    const institution = cleanText_(payload.institution, MAX_INSTITUTION_LENGTH);
    const status = cleanText_(payload.status, MAX_STATUS_LENGTH);
    const photoData = String(payload.photoData || '');

    if (!name || !institution || !status) throw new Error('Nama, instansi, dan status wajib diisi.');
    if (!/^data:image\/(jpeg|jpg|png|webp);base64,/i.test(photoData)) {
      throw new Error('Format foto tidak didukung.');
    }

    const photo = savePhoto_(photoData, name);
    const id = Utilities.getUuid();
    const createdAt = new Date();
    getSheet_().appendRow([id, name, institution, status, createdAt, photo.fileId, photo.url]);

    return json_({ ok: true, data: { id, photoUrl: photo.url } });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function getSheet_() {
  const spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Spreadsheet tidak ditemukan. Isi SPREADSHEET_ID atau buat script dari Spreadsheet.');

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Nama', 'Instansi', 'Status', 'Waktu', 'File ID', 'Link Foto']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function savePhoto_(dataUrl, name) {
  const match = dataUrl.match(/^data:(image\/[a-z0-9.+-]+);base64,(.*)$/i);
  const mimeType = match[1].toLowerCase() === 'image/jpg' ? 'image/jpeg' : match[1].toLowerCase();
  const bytes = Utilities.base64Decode(match[2]);
  const extension = mimeType.split('/')[1].replace('jpeg', 'jpg');
  const safeName = name.replace(/[^a-z0-9_-]+/gi, '-').replace(/^-|-$/g, '').slice(0, 40) || 'pengunjung';
  const fileName = `${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')}-${safeName}.${extension}`;
  const folder = getFolder_();
  const file = folder.createFile(Utilities.newBlob(bytes, mimeType, fileName));

  // Foto galeri harus dapat dibaca oleh browser pengunjung.
  // Jika akun Workspace melarang public link, atur izin folder secara manual.
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (ignored) {}
  return {
    fileId: file.getId(),
    url: `https://drive.google.com/thumbnail?id=${encodeURIComponent(file.getId())}&sz=w1200`
  };
}

function getFolder_() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
}

function cleanText_(value, maxLength) {
  const text = String(value || '').trim().slice(0, maxLength);
  // Mencegah formula injection ketika nilai dibuka di Google Sheets.
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function formatTime_(value) {
  if (!(value instanceof Date)) return String(value || '');
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ' WIB';
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
