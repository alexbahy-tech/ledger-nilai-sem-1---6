// ══════════════════════════════════════════════════════════════════
//  SMK Ledger Nilai — Google Apps Script Web App
//  Version: 4.0 — FIXED: mapping kolom disesuaikan spreadsheet aktual
//  Spreadsheet: https://docs.google.com/spreadsheets/d/1IwPOMqIPk3wtx4WA5pxhTxqUjBJvsOPouq2OpJYRITY
//  Sheet: LEDGER, GID: 263852942
//
//  ⚠️  WAJIB: Buat NEW DEPLOYMENT setelah paste kode ini!
//  Deploy → New Deployment → Web App → Execute as: Me → Access: Anyone
//  Copy URL baru → update di Pengaturan dashboard
// ══════════════════════════════════════════════════════════════════

var SPREADSHEET_ID = '1IwPOMqIPk3wtx4WA5pxhTxqUjBJvsOPouq2OpJYRITY';
var SHEET_GID      = '263852942';
var DATA_START_ROW = 3;  // Baris 1=header semester, Baris 2=nama mapel, Baris 3=data siswa

// ──────────────────────────────────────────────────────────────────
//  MAPPING KOLOM (1-indexed, A=1) — diverifikasi dari file Excel
//
//  A=1  No | B=2  NIS | C=3  NISN | D=4  Nama Siswa | E=5  Program Keahlian | F=6  Kelas
//
//  SEM 1 → G(7)  s/d S(19)  — 13 mapel
//  SEM 2 → T(20) s/d AF(32) — 13 mapel (identik Sem 1)
//  SEM 3 → AG(33)s/d AR(44) — 12 mapel (ada PJOK+Sejarah, tanpa Seni Budaya/Informatika/Projek IPAS/DPK)
//  SEM 4 → AS(45)s/d BD(56) — 12 mapel (identik Sem 3)
//  SEM 5 → BE(57)s/d BN(66) — 10 mapel (tanpa PJOK & Sejarah)
//  SEM 6 → BO(67)s/d BX(76) — 10 mapel (identik Sem 5)
// ──────────────────────────────────────────────────────────────────

var SEM_COL_START = {
  '1':  7,   // kolom G
  '2':  20,  // kolom T
  '3':  33,  // kolom AG
  '4':  45,  // kolom AS
  '5':  57,  // kolom BE  ← FIXED (lama: 55=BC, SALAH)
  '6':  67   // kolom BO  ← FIXED (lama: 65=BM, SALAH)
};

// ⚠️  Urutan HARUS sama persis dengan urutan kolom di spreadsheet!
var SEM_MAPEL = {
  '1': [
    'Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah',
    'Seni Budaya','Mulok','Matematika','Bhs. Inggris',
    'Informatika','Projek IPAS','DPK','Rata-Rata'
  ],
  '2': [
    'Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah',
    'Seni Budaya','Mulok','Matematika','Bhs. Inggris',
    'Informatika','Projek IPAS','DPK','Rata-Rata'
  ],
  '3': [
    'Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah',
    'Mulok','Matematika','Bhs. Inggris',
    'KK','PKK','Mapel Pilihan','Rata-Rata'
  ],
  '4': [
    // FIXED: ditambah PJOK & Sejarah (kode lama hanya 10 mapel, aktual 12)
    'Pend. Agama','Pend. Pancasila','Bhs. Indonesia','PJOK','Sejarah',
    'Mulok','Matematika','Bhs. Inggris',
    'KK','PKK','Mapel Pilihan','Rata-Rata'
  ],
  '5': [
    // Sem 5 tidak ada PJOK & Sejarah — 10 mapel
    'Pend. Agama','Pend. Pancasila','Bhs. Indonesia',
    'Mulok','Matematika','Bhs. Inggris',
    'KK','PKK','Mapel Pilihan','Rata-Rata'
  ],
  '6': [
    // FIXED: dihapus PJOK & Seni Budaya (kode lama ada 12 mapel, aktual 10)
    'Pend. Agama','Pend. Pancasila','Bhs. Indonesia',
    'Mulok','Matematika','Bhs. Inggris',
    'KK','PKK','Mapel Pilihan','Rata-Rata'
  ]
};

// ─── JSON response helper ──────────────────────────────────────────
function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Ambil sheet berdasarkan GID ───────────────────────────────────
function getSheet() {
  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getSheetId()) === String(SHEET_GID)) return sheets[i];
  }
  // Fallback: cari sheet bernama LEDGER
  var ledger = ss.getSheetByName('LEDGER');
  if (ledger) return ledger;
  return ss.getSheets()[0];
}

// ─── Cari baris berdasarkan NIS (kolom B=2) ────────────────────────
function findRowByNIS(sheet, nis) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return -1;
  var nisStr = String(nis).trim();
  var vals   = sheet.getRange(DATA_START_ROW, 2, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === nisStr) return DATA_START_ROW + i;
  }
  return -1;
}

// ─── Angka kolom → huruf Excel ─────────────────────────────────────
function colToLetter(col) {
  var letter = '';
  while (col > 0) {
    var temp = (col - 1) % 26;
    letter   = String.fromCharCode(temp + 65) + letter;
    col      = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ══════════════════════════════════════════════════════════════════
//  doGet — entry point semua request dari browser
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var params  = e && e.parameter ? e.parameter : {};
    var action  = params.action || '';
    var payload = {};
    if (params.payload) {
      try { payload = JSON.parse(decodeURIComponent(params.payload)); }
      catch (pe) { return jsonOut({ status: 'error', message: 'Payload tidak valid: ' + pe.toString() }); }
    }

    switch (action) {
      case 'ping':
        return jsonOut({
          status  : 'ok',
          message : 'GAS v4.0 aktif! Sheet: LEDGER (GID:' + SHEET_GID + ')',
          ts      : new Date().toISOString()
        });
      case 'updateNilai':
        return jsonOut(updateNilai(payload));
      case 'addSiswa':
        return jsonOut(addSiswa(payload));
      case 'updateSiswa':
        return jsonOut(updateSiswa(payload));
      case 'getSiswa':
        return jsonOut(getSiswaData());
      case 'debugCols':
        return jsonOut(debugColumns(payload.nis, parseInt(payload.sem) || 1));
      default:
        return jsonOut({
          status    : 'ok',
          message   : 'GAS SMK Ledger v4.0',
          actions   : ['ping','updateNilai','addSiswa','updateSiswa','getSiswa','debugCols'],
          semConfig : SEM_COL_START
        });
    }
  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'updateNilai':  return jsonOut(updateNilai(body.data || body));
      case 'addSiswa':     return jsonOut(addSiswa(body.data || body));
      case 'updateSiswa':  return jsonOut(updateSiswa(body.data || body));
      default: return jsonOut({ status: 'error', message: 'Action tidak dikenal: ' + body.action });
    }
  } catch (err) {
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

// ══════════════════════════════════════════════════════════════════
//  updateNilai — simpan nilai semester ke spreadsheet
//  Payload: { nis, semester (1-6), nilai: { 'Nama Mapel': angka } }
// ══════════════════════════════════════════════════════════════════
function updateNilai(data) {
  if (!data)          return { status: 'error', message: 'Payload kosong' };
  if (!data.nis)      return { status: 'error', message: 'Field "nis" wajib' };
  if (!data.semester) return { status: 'error', message: 'Field "semester" wajib' };
  if (!data.nilai)    return { status: 'error', message: 'Field "nilai" wajib (object mapel: angka)' };

  var sheet    = getSheet();
  var rowNum   = findRowByNIS(sheet, data.nis);

  if (rowNum === -1) {
    return {
      status : 'error',
      message: 'NIS "' + data.nis + '" tidak ditemukan di kolom B (mulai baris ' + DATA_START_ROW + '). ' +
               'Pastikan NIS persis sama, tanpa spasi atau nol ekstra.'
    };
  }

  var semKey   = String(parseInt(data.semester));
  var mapel    = SEM_MAPEL[semKey];
  var colStart = SEM_COL_START[semKey];

  if (!mapel || !colStart) {
    return { status: 'error', message: 'Semester tidak valid: "' + semKey + '". Harus 1-6.' };
  }

  var saved   = [];
  var skipped = [];
  var errors  = [];

  for (var idx = 0; idx < mapel.length; idx++) {
    var m   = mapel[idx];
    var val = data.nilai[m];

    // Skip Rata-Rata — biarkan formula spreadsheet yang menghitung
    if (m === 'Rata-Rata') {
      skipped.push('Rata-Rata (formula spreadsheet)');
      continue;
    }

    if (val === undefined || val === null || String(val).trim() === '') {
      skipped.push(m);
      continue;
    }

    var num = parseFloat(val);
    if (isNaN(num))            { errors.push(m + ': bukan angka (' + val + ')'); continue; }
    if (num < 0 || num > 100)  { errors.push(m + ': di luar range 0-100 (' + num + ')'); continue; }

    var targetCol = colStart + idx;
    try {
      sheet.getRange(rowNum, targetCol).setValue(num);
      saved.push(m + '=' + num + ' [' + colToLetter(targetCol) + rowNum + ']');
    } catch (cellErr) {
      errors.push(m + ' [' + colToLetter(targetCol) + ']: ' + cellErr.toString());
    }
  }

  SpreadsheetApp.flush();

  return {
    status : errors.length > 0 ? 'warning' : 'ok',
    message: 'Berhasil simpan ' + saved.length + ' nilai Semester ' + semKey +
             ' untuk NIS: ' + data.nis + ' (baris ' + rowNum + ')',
    detail : {
      nis          : data.nis,
      row          : rowNum,
      semester     : semKey,
      colStart     : colStart,
      colStartLtr  : colToLetter(colStart),
      totalMapel   : mapel.length,
      savedCount   : saved.length,
      saved        : saved,
      skippedCount : skipped.length,
      skipped      : skipped,
      errors       : errors
    }
  };
}

// ══════════════════════════════════════════════════════════════════
//  addSiswa — tambah siswa baru di baris berikutnya
// ══════════════════════════════════════════════════════════════════
function addSiswa(data) {
  if (!data || !data.nis || !data.nama)
    return { status: 'error', message: 'Field "nis" dan "nama" wajib' };

  var sheet    = getSheet();
  var existing = findRowByNIS(sheet, data.nis);
  if (existing !== -1)
    return { status: 'error', message: 'NIS "' + data.nis + '" sudah ada di baris ' + existing };

  var newRow = Math.max(sheet.getLastRow() + 1, DATA_START_ROW);
  sheet.getRange(newRow, 1, 1, 6).setValues([[
    data.no || (newRow - DATA_START_ROW + 1),
    data.nis,
    data.nisn  || '',
    data.nama,
    data.prodi || '',
    data.kelas || ''
  ]]);
  SpreadsheetApp.flush();

  return { status: 'ok', message: '"' + data.nama + '" ditambahkan di baris ' + newRow, row: newRow };
}

// ══════════════════════════════════════════════════════════════════
//  updateSiswa — edit identitas siswa
// ══════════════════════════════════════════════════════════════════
function updateSiswa(data) {
  if (!data || !data.nis) return { status: 'error', message: 'Field "nis" wajib' };
  var sheet  = getSheet();
  var rowNum = findRowByNIS(sheet, data.nis);
  if (rowNum === -1) return { status: 'error', message: 'NIS "' + data.nis + '" tidak ditemukan' };

  if (data.nisn  !== undefined) sheet.getRange(rowNum, 3).setValue(data.nisn);
  if (data.nama  !== undefined) sheet.getRange(rowNum, 4).setValue(data.nama);
  if (data.prodi !== undefined) sheet.getRange(rowNum, 5).setValue(data.prodi);
  if (data.kelas !== undefined) sheet.getRange(rowNum, 6).setValue(data.kelas);

  SpreadsheetApp.flush();
  return { status: 'ok', message: 'Siswa NIS "' + data.nis + '" diperbarui di baris ' + rowNum };
}

// ══════════════════════════════════════════════════════════════════
//  getSiswaData — ambil data identitas semua siswa
// ══════════════════════════════════════════════════════════════════
function getSiswaData() {
  var sheet   = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { status: 'ok', data: [], count: 0 };

  var vals = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();
  var out  = vals
    .filter(function(r) { return String(r[0]).trim() !== '' || String(r[3]).trim() !== ''; })
    .map(function(r) {
      return {
        no    : r[0],
        nis   : String(r[1]||'').trim(),
        nisn  : String(r[2]||'').trim(),
        nama  : String(r[3]||'').trim(),
        prodi : String(r[4]||'').trim(),
        kelas : String(r[5]||'').trim()
      };
    });
  return { status: 'ok', data: out, count: out.length };
}

// ══════════════════════════════════════════════════════════════════
//  debugColumns — verifikasi mapping kolom (untuk troubleshoot)
// ══════════════════════════════════════════════════════════════════
function debugColumns(nis, sem) {
  var semKey   = String(sem || 1);
  var mapel    = SEM_MAPEL[semKey] || [];
  var colStart = SEM_COL_START[semKey] || 0;
  var sheet    = getSheet();
  var rowNum   = nis ? findRowByNIS(sheet, String(nis)) : -1;

  var mapping = mapel.map(function(m, i) {
    var col = colStart + i;
    return {
      index      : i,
      mapel      : m,
      col_1based : col,
      colLetter  : colToLetter(col),
      currentVal : (rowNum > 0) ? sheet.getRange(rowNum, col).getValue() : '(NIS tidak ditemukan/tidak diisi)'
    };
  });

  return {
    status         : 'ok',
    nis            : nis || '(tidak diisi)',
    rowFound       : rowNum,
    semester       : semKey,
    colStart       : colStart,
    colStartLetter : colToLetter(colStart),
    totalMapel     : mapel.length,
    allSemConfig   : SEM_COL_START,
    mapping        : mapping
  };
}