/**
 * PAKPOL - POLSEK JATIROGO (Backend Google Apps Script)
 * Semua data disimpan di folder Google Drive "PAKPOL Polsek Jatirogo":
 *   - Database  : Google Sheets "PAKPOL (Database)" (sheet "Arsip" dan "Personel")
 *   - PDF       : subfolder "PDF"
 *
 * LANGKAH PASANG:
 * 1. script.google.com > Proyek baru > tempel kode ini.
 * 2. GANTI nilai TOKEN di bawah dengan kode rahasia buatan Anda sendiri.
 * 3. Jalankan fungsi "setup" sekali (izinkan akses Drive & Sheets).
 * 4. Terapkan > Penerapan baru > Aplikasi web.
 *    Jalankan sebagai: Saya | Yang memiliki akses: Siapa saja > Terapkan.
 * 5. Salin URL Web App (berakhiran /exec), lalu isi di PAKPOL (ikon roda gigi)
 *    bersama kode rahasia yang sama.
 * Setiap mengubah kode: Terapkan > Kelola penerapan > pensil > Versi baru > Terapkan.
 */
var TOKEN = 'polsek2026';
var FOLDER_NAME = 'PAKPOL Polsek Jatirogo';
var DB_NAME = 'PAKPOL (Database)';
var HEAD = ['id', 'Jenis', 'Judul', 'Tanggal', 'Operator', 'Link PDF', 'Disimpan', 'Data (JSON)'];

function root_() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function ss_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DB_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (x) {}
  }
  var ss = SpreadsheetApp.create(DB_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(root_());
  props.setProperty('DB_ID', ss.getId());
  return ss;
}

function arsip_() {
  var ss = ss_();
  var sh = ss.getSheetByName('Arsip');
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName('Arsip');
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function pers_() {
  var ss = ss_();
  var sh = ss.getSheetByName('Personel') || ss.insertSheet('Personel');
  if (sh.getLastRow() === 0) sh.appendRow(['{"t":"","ts":0}']);
  return sh;
}

function pdfFolder_() {
  var r = root_();
  var it = r.getFoldersByName('PDF');
  return it.hasNext() ? it.next() : r.createFolder('PDF');
}

// Jalankan SEKALI untuk memberi izin dan membuat database + folder
function setup() {
  arsip_();
  pers_();
  pdfFolder_();
  Logger.log('PAKPOL siap. Folder dan database sudah dibuat di Google Drive.');
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function rowOf_(sh, id) {
  if (!id || sh.getLastRow() < 1) return 0;
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; });
  return ids.indexOf(id) + 1;
}

// Membaca data (arsip + data personel)
function doGet(e) {
  if (!e || !e.parameter || e.parameter.token !== TOKEN) return out_({ error: 'Kode rahasia salah' });
  var rows = arsip_().getDataRange().getValues();
  var recs = [];
  for (var i = 1; i < rows.length; i++) {
    try { recs.push(JSON.parse(rows[i][7])); } catch (x) {}
  }
  var p = { t: '', ts: 0 };
  try { p = JSON.parse(pers_().getRange(1, 1).getValue()); } catch (x) {}
  return out_({ records: recs, pers: p });
}

// Menyimpan / menghapus
function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ error: 'Data tidak valid' }); }
  if (b.token !== TOKEN) return out_({ error: 'Kode rahasia salah' });

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var act = b.action;

    if (act === 'persSave') {
      pers_().getRange(1, 1).setValue(JSON.stringify({ t: b.t, ts: b.ts }));
      return out_({ ok: true });
    }

    var sh = arsip_();

    if (act === 'save') {
      var r = b.record;
      var vals = [r.id, r.j, r.s, r.tgl, r.oleh || '', r.u || '', new Date(r.saved || Date.now()), JSON.stringify(r)];
      var row = rowOf_(sh, r.id);
      if (row > 1) sh.getRange(row, 1, 1, vals.length).setValues([vals]);
      else sh.appendRow(vals);
      return out_({ ok: true });
    }

    if (act === 'delete') {
      var drow = rowOf_(sh, b.id);
      if (drow > 1) {
        try {
          var rec = JSON.parse(sh.getRange(drow, 8).getValue());
          if (rec.u) {
            var m = String(rec.u).match(/[-\w]{25,}/);
            if (m) {
              var f = DriveApp.getFileById(m[0]);
              var ps = f.getParents();
              while (ps.hasNext()) {
                if (ps.next().getId() === pdfFolder_().getId()) { f.setTrashed(true); break; }
              }
            }
          }
        } catch (x) {}
        sh.deleteRow(drow);
      }
      return out_({ ok: true });
    }

    if (act === 'pdf') {
      var blob = Utilities.newBlob(Utilities.base64Decode(b.data), 'application/pdf', b.filename);
      var file = pdfFolder_().createFile(blob);
      return out_({ ok: true, url: file.getUrl(), fileId: file.getId() });
    }

    return out_({ error: 'Aksi tidak dikenal' });
  } finally {
    lock.releaseLock();
  }
}
