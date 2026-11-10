/**
 * ABelSekolah - Backend (Google Apps Script)
 * Pasang script ini dari Google Sheet: Extensions > Apps Script
 * Deploy sebagai Web app (Execute as: Me, Who has access: Anyone). Frontend ada di GitHub/Cloudflare Pages.
 */
const SHEETS = {
  Jadwal:     ['ID', 'Nama', 'Jam', 'Hari', 'Nada', 'Aktif'],
  Pengaturan: ['Kunci', 'Nilai'],
  Riwayat:    ['Waktu', 'Nama Bel', 'Jam'],
  Audio:      ['ID', 'Nama', 'FileId', 'Ukuran (KB)']
};
const FOLDER_NAME = 'ABelSekolah Audio';
const MAX_BYTES = 5 * 1024 * 1024; // batas 5 MB per file

/* ---------- API (dipanggil dari Cloudflare Pages lewat fetch) ---------- */
function doGet() {
  return json_({ ok: true, app: 'ABelSekolah', status: 'API aktif' });
}

const WRITES = ['saveJadwal', 'toggleJadwal', 'deleteJadwal', 'saveSettings', 'uploadAudio', 'deleteAudio', 'uploadLogo'];

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = {
      getData: getData, getLogo: getLogo, getAudioData: getAudioData, logBell: logBell,
      saveJadwal: saveJadwal, toggleJadwal: toggleJadwal, deleteJadwal: deleteJadwal, saveSettings: saveSettings,
      uploadAudio: uploadAudio, deleteAudio: deleteAudio, uploadLogo: uploadLogo
    };
    if (!fns[req.fn]) throw new Error('Fungsi tidak dikenal.');
    if (req.fn === 'getData') setup_();
    if (WRITES.indexOf(req.fn) >= 0) {
      const pin = getPin_(); // PIN diisi di sheet Pengaturan (baris "pin"); kosong = tanpa PIN
      if (pin && String(req.pin || '') !== pin) return json_({ ok: false, code: 'PIN', error: 'PIN admin salah atau belum diisi.' });
    }
    return json_({ ok: true, data: fns[req.fn].apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function getPin_() {
  const r = rows_('Pengaturan').filter(function (x) { return x[0] === 'pin'; })[0];
  return r ? String(r[1]).trim() : '';
}

/* ---------- Setup otomatis ---------- */
function setup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]])
        .setFontWeight('bold').setBackground('#12263F').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      sh.getRange('A:F').setNumberFormat('@'); // simpan sebagai teks agar jam tidak berubah format
      if (name === 'Pengaturan') {
        sh.getRange(2, 1, 2, 2).setValues([['namaSekolah', 'Nama Sekolah Anda'], ['volume', '80']]);
      }
      if (name === 'Jadwal') {
        sh.getRange(2, 1, 3, 6).setValues([
          [Utilities.getUuid(), 'Bel Masuk',  '07:00', '1,2,3,4,5', 'klasik',  'TRUE'],
          [Utilities.getUuid(), 'Istirahat',  '09:30', '1,2,3,4,5', 'digital', 'TRUE'],
          [Utilities.getUuid(), 'Bel Pulang', '14:00', '1,2,3,4,5', 'sekolah', 'TRUE']
        ]);
      }
    }
  });
  // kolom logo: isi dengan link/ID file Google Drive atau URL gambar; kosong = logo bawaan
  const ps = ss.getSheetByName('Pengaturan');
  if (findRow_(ps, 'logo') < 0) {
    ps.appendRow(['logo', '']);
    ps.getRange(ps.getLastRow(), 2).setNote('Isi dengan link Google Drive, ID file, atau URL gambar (https://...). Kosongkan untuk memakai logo bawaan.');
  }
  if (findRow_(ps, 'pin') < 0) {
    ps.appendRow(['pin', '']);
    ps.getRange(ps.getLastRow(), 2).setNote('PIN admin (opsional). Jika diisi, PIN diminta setiap kali mengubah data. Kosongkan untuk menonaktifkan.');
  }
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function rows_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const v = sh.getDataRange().getDisplayValues();
  v.shift();
  return v.filter(function (r) { return r[0] !== ''; });
}

/* ---------- API untuk client ---------- */
function getData() {
  const settings = {};
  rows_('Pengaturan').forEach(function (r) { if (r[0] !== 'pin') settings[r[0]] = r[1]; }); // PIN tidak pernah dikirim ke client
  const jadwal = rows_('Jadwal').map(function (r) {
    return { id: r[0], nama: r[1], jam: r[2], hari: r[3].split(',').filter(String).map(Number), nada: r[4], aktif: r[5] === 'TRUE' };
  }).sort(function (a, b) { return a.jam.localeCompare(b.jam); });
  const riwayat = rows_('Riwayat').slice(-50).reverse().map(function (r) { return { waktu: r[0], nama: r[1], jam: r[2] }; });
  const audio = rows_('Audio').map(function (r) { return { id: r[0], nama: r[1], ukuran: r[3] }; });
  return { settings: settings, jadwal: jadwal, riwayat: riwayat, audio: audio };
}

function saveJadwal(d) {
  if (!d.nama || !/^\d{2}:\d{2}$/.test(d.jam)) throw new Error('Nama dan jam wajib diisi dengan benar.');
  if (!d.hari || !d.hari.length) throw new Error('Pilih minimal satu hari.');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jadwal');
  const row = [d.id || Utilities.getUuid(), d.nama, d.jam, d.hari.join(','), d.nada, d.aktif === false ? 'FALSE' : 'TRUE'];
  const idx = d.id ? findRow_(sh, d.id) : -1;
  if (idx > 0) sh.getRange(idx, 1, 1, 6).setValues([row]); else sh.appendRow(row);
  return true;
}

function toggleJadwal(id, aktif) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jadwal');
  const idx = findRow_(sh, id);
  if (idx > 0) sh.getRange(idx, 6).setValue(aktif ? 'TRUE' : 'FALSE');
  return true;
}

function deleteJadwal(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jadwal');
  const idx = findRow_(sh, id);
  if (idx > 0) sh.deleteRow(idx);
  return true;
}

function saveSettings(s) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengaturan');
  Object.keys(s).forEach(function (k) {
    if (k === 'pin') return; // PIN hanya boleh diubah langsung di Google Sheet
    const idx = findRow_(sh, k);
    if (idx > 0) sh.getRange(idx, 2).setValue(String(s[k])); else sh.appendRow([k, String(s[k])]);
  });
  return true;
}

function logBell(nama, jam) {
  const tz = Session.getScriptTimeZone();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Riwayat')
    .appendRow([Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'), nama, jam]);
  return true;
}

/* ---------- Audio: file disimpan di Google Drive, data & ID-nya dicatat di Google Sheet ---------- */
function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function uploadAudio(nama, mime, base64) {
  if (!/^audio\//.test(mime)) throw new Error('File harus berupa audio (MP3, WAV, OGG, M4A).');
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > MAX_BYTES) throw new Error('Ukuran file maksimal 5 MB.');
  const file = folder_().createFile(Utilities.newBlob(bytes, mime, nama));
  const id = Utilities.getUuid();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audio')
    .appendRow([id, nama.replace(/\.[^.]+$/, ''), file.getId(), String(Math.round(bytes.length / 1024))]);
  return id;
}

function getAudioData(id) {
  const r = rows_('Audio').filter(function (x) { return x[0] === id; })[0];
  if (!r) throw new Error('File audio tidak ditemukan.');
  const blob = DriveApp.getFileById(r[2]).getBlob();
  return { mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}

function deleteAudio(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Audio');
  const idx = findRow_(sh, id);
  if (idx > 0) {
    try { DriveApp.getFileById(sh.getRange(idx, 3).getDisplayValue()).setTrashed(true); } catch (e) {}
    sh.deleteRow(idx);
  }
  // jadwal yang memakai audio ini dikembalikan ke nada Klasik
  const js = ss.getSheetByName('Jadwal'), last = js.getLastRow();
  if (last > 1) {
    const rg = js.getRange(2, 5, last - 1, 1), v = rg.getValues();
    rg.setValues(v.map(function (r) { return [r[0] === 'audio:' + id ? 'klasik' : r[0]]; }));
  }
  return true;
}

/* ---------- Logo sekolah ---------- */
function getLogo() {
  const row = rows_('Pengaturan').filter(function (r) { return r[0] === 'logo'; })[0];
  const v = row ? String(row[1]).trim() : '';
  if (!v) return '';
  const isUrl = /^https?:\/\//.test(v);
  if (isUrl && !/(drive|docs)\.google\.com/.test(v)) return v; // URL gambar biasa
  const m = v.match(/[-\w]{25,}/);
  if (!m) throw new Error('Isi sel "logo" di sheet Pengaturan tidak dikenali.');
  const blob = DriveApp.getFileById(m[0]).getBlob();
  if (!/^image\//.test(blob.getContentType())) throw new Error('File logo harus berupa gambar.');
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function uploadLogo(mime, base64) {
  if (!/^image\//.test(mime)) throw new Error('File logo harus berupa gambar (PNG, JPG, SVG).');
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > 2 * 1024 * 1024) throw new Error('Ukuran logo maksimal 2 MB.');
  const file = folder_().createFile(Utilities.newBlob(bytes, mime, 'logo-sekolah'));
  saveSettings({ logo: file.getId() });
  return true;
}

function findRow_(sh, key) {
  const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < col.length; i++) if (String(col[i][0]) === String(key)) return i + 1;
  return -1;
}
