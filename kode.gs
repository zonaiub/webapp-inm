/**
 * =================================================================
 * SISTEM INPUT INDIKATOR NASIONAL MUTU (INM) - BERBASIS DATA MASTER
 * DENGAN FITUR AUTO-ARCHIVING KE GOOGLE DRIVE
 * =================================================================
 */

// =================================================================
// KONFIGURASI FOLDER ARSIP (WAJIB DIISI)
// =================================================================
// 1. Buat folder "00_ARSIP_DATA_MUTU" di Google Drive.
// 2. Buka folder tersebut, lihat URL-nya di browser.
// 3. Copy kode ID yang ada di ujung URL (setelah "folders/") dan paste di bawah ini.
const ARCHIVE_FOLDER_ID = '10hynL1WTC5--T74rQGM5ff74B1xU6arq';
// =================================================================

const CONFIG_SHEET_NAME = 'Config';
const DATA_SHEET_NAME = 'Data Master';
const USERS_SHEET_NAME = 'Users';

// Header kolom di sheet "Data Master". JANGAN diubah urutannya
const DATA_HEADERS = ['Timestamp', 'Tanggal', 'Ruangan', 'Indikator', 'Numerator', 'Denominator', 'Diisi Oleh', 'Keterangan'];

// Indikator yang arah targetnya kebalik: target itu batas MAKSIMAL.
const REVERSE_INDICATOR_NAMES = ['Penundaan Operasi Elektif'];

/**
 * Membersihkan teks sebelum dibandingkan
 */
function normText(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Menyeragamkan tampilan nama (misal "novi" jadi "Novi")
 */
function titleCase(s) {
  return String(s || '').trim().toLowerCase().replace(/(^|\s)\S/g, function (c) { return c.toUpperCase(); });
}

// Keterangan Numerator & Denominator untuk tiap indikator.
const INDICATOR_DETAILS = [
  { nama: 'Kepatuhan Penggunaan APD', num: 'Jumlah petugas yang patuh menggunakan APD sesuai indikasi dalam periode observasi', den: 'Jumlah seluruh petugas yang terindikasi menggunakan APD dalam periode observasi' },
  { nama: 'Kepatuhan Identifikasi Pasien', num: 'Jumlah pemberi pelayanan yang melakukan identifikasi pasien secara benar dalam periode observasi', den: 'Jumlah pemberi pelayanan yang diobservasi dalam periode observasi' },
  { nama: 'Waktu Tanggap Operasi Seksio Sesarea Emergensi', num: 'Jumlah pasien yang diputuskan tindakan seksio sesarea emergensi kategori I (satu) yang mendapatkan tindakan seksio sesarea emergensi ≤ 30 menit', den: 'Jumlah pasien yang diputuskan tindakan seksio sesarea emergensi kategori I (satu)' },
  { nama: 'Waktu Tunggu Rawat Jalan', num: 'Jumlah pasien rawat jalan dengan waktu tunggu ≤ 60 menit', den: 'Jumlah pasien rawat jalan yang diobservasi' },
  { nama: 'Penundaan Operasi Elektif', num: 'Jumlah pasien yang waktu jadwal operasinya tertunda lebih dari 1 jam', den: 'Jumlah pasien operasi elektif' },
  { nama: 'Ketepatan Waktu Visite Dokter', num: 'Jumlah di-visite Dokter pada pukul 06.00 - 14:00', den: 'Jumlah pasien yang diobservasi' },
  { nama: 'Pelaporan Hasil Kritis Laboratorium', num: 'Jumlah hasil kritis laboratorium yang dilaporkan ≤30 menit', den: 'Jumlah hasil kritis laboratorium yang diobservasi' },
  { nama: 'Kepatuhan Penggunaan Formularium Nasional', num: 'Jumlah R/ recipe dalam lembar resep yang sesuai dengan formularium nasional', den: 'Jumlah R/ recipe dalam lembar resep yang diobservasi' },
  { nama: 'Kepatuhan Terhadap Alur Klinis (Clinical Pathway)', num: 'Jumlah pelayanan oleh PPA yang sesuai dengan clinical pathway', den: 'Jumlah seluruh pelayanan oleh PPA pada clinical pathway yang diobservasi' },
  { nama: 'Kepatuhan Upaya Pencegahan Risiko Pasien Jatuh', num: 'Jumlah pasien rawat inap berisiko tinggi jatuh yang mendapatkan ketiga upaya pencegahan risiko jatuh', den: 'Jumlah pasien rawat inap berisiko tinggi jatuh yang diobservasi' },
  { nama: 'Kecepatan Waktu Tanggap Komplain', num: 'Jumlah komplain yang ditanggapi dan ditindaklanjuti sesuai waktu yang ditetapkan berdasarkan grading', den: 'Jumlah Komplain yang disurvei' },
  { nama: 'Kepuasan Pasien', num: 'Total nilai persepsi seluruh responden', den: 'Total unsur yang terisi dari seluruh responden' }
];

/**
 * =================================================================
 * FITUR DATA ARCHIVING
 * =================================================================
 */

/**
 * Fungsi Pemicu Uji Coba: Jalankan dari editor untuk mencoba arsip tahun lalu
 */
function testRunArchiving() {
  const currentYear = new Date().getFullYear();
  const testYear = currentYear - 1; // Otomatis uji coba untuk data tahun lalu
  Logger.log('Memulai simulasi pengarsipan untuk tahun: ' + testYear);
  archiveData(testYear);
}

// TARUH DI LUAR SINI BRO, JANGAN DI DALAM FUNGSI LAIN
function autoArchiveTahunan() {
  const prevYear = new Date().getFullYear() - 1;
  Logger.log('Menjalankan Auto-Archive untuk tahun: ' + prevYear);
  archiveData(prevYear);
}

/**
 * Logika utama untuk memindahkan data lama dari Sheet Utama ke file Arsip di Google Drive
 */
function archiveData(yearToArchive) {
  if (!ARCHIVE_FOLDER_ID || ARCHIVE_FOLDER_ID === 'GANTI_DENGAN_ID_FOLDER_00_ARSIP_DATA_MUTU_DI_SINI') {
    throw new Error("GAGAL: Silakan isi ARCHIVE_FOLDER_ID di baris paling atas kode terlebih dahulu.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("Sheet Data Master kosong.");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
  const rowsToArchive = [];
  const rowsToKeep = [];

  data.forEach(function(row) {
    if (!row[1]) return;
    const tgl = new Date(row[1]);
    if (tgl.getFullYear() === yearToArchive) {
      rowsToArchive.push(row);
    } else {
      rowsToKeep.push(row);
    }
  });

  if (rowsToArchive.length === 0) {
    Logger.log("Tidak ada data untuk tahun " + yearToArchive + " yang perlu diarsipkan dari Sheet Utama.");
    return;
  }

  const folder = DriveApp.getFolderById(ARCHIVE_FOLDER_ID);
  const fileName = "Arsip_INM_" + yearToArchive;

  let archiveSs;
  const existingFiles = folder.searchFiles("title = '" + fileName + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
  
  if (existingFiles.hasNext()) {
    archiveSs = SpreadsheetApp.openById(existingFiles.next().getId());
    Logger.log("File arsip ditemukan, menambahkan data...");
  } else {
    const newFile = SpreadsheetApp.create(fileName);
    const file = DriveApp.getFileById(newFile.getId());
    file.moveTo(folder);
    archiveSs = SpreadsheetApp.openById(newFile.getId());
    archiveSs.getSheets()[0].appendRow(DATA_HEADERS);
    archiveSs.getSheets()[0].setFrozenRows(1);
    Logger.log("Membuat file arsip baru: " + fileName);
  }

  const archiveSheet = archiveSs.getSheets()[0];
  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, DATA_HEADERS.length).setValues(rowsToArchive);

  // Bersihkan Sheet Utama dan tulis ulang data yang tersisa
  sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).clearContent();
  if (rowsToKeep.length > 0) {
    sheet.getRange(2, 1, rowsToKeep.length, DATA_HEADERS.length).setValues(rowsToKeep);
  }

  Logger.log("Selesai! Berhasil memindahkan " + rowsToArchive.length + " baris data tahun " + yearToArchive + " ke arsip.");
}

/**
 * Smart Data Reader: Mengambil data dari Sheet Utama + File Arsip sesuai tahun.
 * Dilengkapi deduplikasi agar data edit (di Sheet Utama) menimpa data lama (di Arsip).
 */
function getRawDataByYears(startYear, endYear) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let allData = [];

  // 1. Selalu ambil dari Sheet Utama (Data Master) sebagai prioritas tertinggi
  const sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    allData = allData.concat(sheet.getRange(2, 1, sheet.getLastRow() - 1, DATA_HEADERS.length).getValues());
  }

  // 2. Ambil dari folder Arsip jika ID valid
  if (ARCHIVE_FOLDER_ID && ARCHIVE_FOLDER_ID !== 'GANTI_DENGAN_ID_FOLDER_00_ARSIP_DATA_MUTU_DI_SINI') {
    for (let y = startYear; y <= endYear; y++) {
      try {
        const folder = DriveApp.getFolderById(ARCHIVE_FOLDER_ID);
        const files = folder.searchFiles("title = 'Arsip_INM_" + y + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
        if (files.hasNext()) {
          const archiveSs = SpreadsheetApp.openById(files.next().getId());
          const archSheet = archiveSs.getSheets()[0];
          if (archSheet && archSheet.getLastRow() > 1) {
            allData = allData.concat(archSheet.getRange(2, 1, archSheet.getLastRow() - 1, DATA_HEADERS.length).getValues());
          }
        }
      } catch (e) {
        // Abaikan jika arsip tahun tersebut belum ada
      }
    }
  }

  // 3. Filter berdasarkan rentang tahun
  const filtered = allData.filter(function(row) {
    if(!row[1]) return false;
    const y = new Date(row[1]).getFullYear();
    return y >= startYear && y <= endYear;
  });

  // 4. Deduplikasi: Karena allData diisi Data Master duluan, data baru akan dipertahankan
  const uniqueData = {};
  filtered.forEach(function(row) {
    const rowTanggal = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const key = rowTanggal + '_' + normText(row[2]) + '_' + normText(row[3]);
    if (!uniqueData[key]) {
      uniqueData[key] = row;
    }
  });

  return Object.values(uniqueData);
}

/**
 * =================================================================
 * WEB APP & FUNGSI INTI
 * =================================================================
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Form')
    .setTitle('Input Data INM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function isiSemuaTarget() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  sheet.getRange('E1').setValue('Target (%)');

  const targets = {
    'Kepatuhan Kebersihan Tangan': 85, 'Kepatuhan Penggunaan APD': 100, 'Kepatuhan Identifikasi Pasien': 100,
    'Waktu Tanggap Operasi Seksio Sesarea Emergensi': 80, 'Waktu Tunggu Rawat Jalan': 80, 'Penundaan Operasi Elektif': 5,
    'Ketepatan Waktu Visite Dokter': 80, 'Pelaporan Hasil Kritis Laboratorium': 100, 'Kepatuhan Penggunaan Formularium Nasional': 80,
    'Kepatuhan Terhadap Alur Klinis (Clinical Pathway)': 80, 'Kepatuhan Upaya Pencegahan Risiko Pasien Jatuh': 100,
    'Kecepatan Waktu Tanggap Komplain': 80, 'Kepuasan Pasien': 76.61
  };

  const lastRow = sheet.getLastRow();
  const names = sheet.getRange('B2:B' + lastRow).getValues();
  names.forEach(function (row, i) {
    if (targets[row[0]] !== undefined) sheet.getRange(i + 2, 5).setValue(targets[row[0]]);
  });
}

function tambahKolomTarget() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  sheet.getRange('E1').setValue('Target (%)');
}

function setupUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) usersSheet = ss.insertSheet(USERS_SHEET_NAME);
  if (usersSheet.getLastRow() > 0) return;

  usersSheet.appendRow(['Username', 'Password', 'Role', 'Ruangan']);
  usersSheet.setFrozenRows(1);

  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const rooms = configSheet.getRange('A2:A' + configSheet.getLastRow()).getValues().flat().filter(String);

  const rows = rooms.map(function (room) {
    const username = room.toLowerCase().replace(/[^a-z0-9]/g, '');
    return [username, 'inm123', 'pic', room];
  });
  rows.push(['admin', 'admin123', 'admin', '']);

  usersSheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

function login(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Sistem login belum di-setup.' };

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const inputUser = String(username || '').trim().toLowerCase();
  const inputPass = String(password || '').trim();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === inputUser && String(data[i][1] || '').trim() === inputPass) {
      return { success: true, role: data[i][2], ruangan: data[i][3] };
    }
  }
  return { success: false, message: 'Username atau password salah.' };
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) dataSheet = ss.insertSheet(DATA_SHEET_NAME);
  if (dataSheet.getLastRow() === 0) {
    dataSheet.appendRow(DATA_HEADERS);
    dataSheet.setFrozenRows(1);
  }

  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(['Daftar Ruangan', 'Daftar Indikator', 'Keterangan Numerator', 'Keterangan Denominator']);
    const rooms = [
      'ARIMBI', 'BROTOJOYO', 'CITROANGGODO', 'DEWARUCI', 'GATOTKACA', 'HUDOWO',
      'ECTS', 'SRIKANDI A', 'SRIKANDI B', 'YUDISTIRA', 'UPIP', 'JANOKO', 'RIPD',
      'IGD', 'RAJAL', 'RAMASINTA', 'LAB', 'ICU', 'IBS', 'BISMA', 'ENDROTENOYO',
      'HUMAS', 'FARMASI', 'FISIOTERAPI', 'GIZI', 'NAPZA', 'REHABSOS',
      'POLI GIGI', 'RADIOLOGI'
    ];
    const maxLen = Math.max(rooms.length, INDICATOR_DETAILS.length);
    for (let i = 0; i < maxLen; i++) {
      configSheet.getRange(i + 2, 1).setValue(rooms[i] || '');
      if (INDICATOR_DETAILS[i]) {
        configSheet.getRange(i + 2, 2).setValue(INDICATOR_DETAILS[i].nama);
        configSheet.getRange(i + 2, 3).setValue(INDICATOR_DETAILS[i].num);
        configSheet.getRange(i + 2, 4).setValue(INDICATOR_DETAILS[i].den);
      }
    }
  }
}

function getRooms() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  return sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat().filter(String);
}

function getIndicatorDetails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const data = sheet.getRange('B2:E' + sheet.getLastRow()).getValues();
  return data.filter(function (row) { return row[0]; }).map(function (row) {
    const target = (row[3] === '' || row[3] === undefined) ? null : Number(row[3]);
    const reverse = REVERSE_INDICATOR_NAMES.map(normText).indexOf(normText(row[0])) > -1;
    return { nama: row[0], num: row[1] || '', den: row[2] || '', target: target, tipe: 'persen', reverse: reverse };
  });
}

function getExistingEntries(ruangan, tanggal) {
  const targetYear = new Date(tanggal).getFullYear();
  const data = getRawDataByYears(targetYear, targetYear);
  const result = {};
  const targetRuangan = normText(ruangan);
  
  data.forEach(function (row) {
    const rowTanggal = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (rowTanggal === tanggal && normText(row[2]) === targetRuangan) {
      result[normText(row[3])] = { numerator: row[4], denominator: row[5], keterangan: row[7] || '' };
    }
  });
  return result;
}

function getMonthlyRecap(ruangan, bulan, tahun) {
  const data = getRawDataByYears(Number(tahun), Number(tahun));
  const totals = {};
  data.forEach(function (row) {
    const tgl = new Date(row[1]);
    if (row[2] === ruangan && (tgl.getMonth() + 1) === Number(bulan) && tgl.getFullYear() === Number(tahun)) {
      const ind = row[3];
      if (!totals[ind]) totals[ind] = { numerator: 0, denominator: 0 };
      totals[ind].numerator += Number(row[4]) || 0;
      totals[ind].denominator += Number(row[5]) || 0;
    }
  });
  return Object.keys(totals).map(function (ind) {
    const t = totals[ind];
    return { indikator: ind, numerator: t.numerator, denominator: t.denominator, hasil: t.denominator > 0 ? (t.numerator / t.denominator * 100) : null };
  });
}

/**
 * Menyimpan data selalu terpusat di Data Master.
 * Archive akan menanganinya nanti.
 */
function submitData(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues() : [];

  const tanggal = payload.tanggal;
  const ruangan = payload.ruangan;
  const targetRuangan = normText(ruangan);
  const now = new Date();

  payload.entries.forEach(function (entry) {
    if (entry.numerator === '' && entry.denominator === '') return;

    const targetIndikator = normText(entry.indikator);
    let foundRowIndex = -1;
    for (let i = 0; i < existing.length; i++) {
      const rowTanggal = Utilities.formatDate(new Date(existing[i][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rowTanggal === tanggal && normText(existing[i][2]) === targetRuangan && normText(existing[i][3]) === targetIndikator) {
        foundRowIndex = i + 2;
        break;
      }
    }

    const rowValues = [now, new Date(tanggal), ruangan, entry.indikator, Number(entry.numerator) || 0, Number(entry.denominator) || 0, payload.diisiOleh || '', entry.keterangan || ''];
    if (foundRowIndex > -1) sheet.getRange(foundRowIndex, 1, 1, DATA_HEADERS.length).setValues([rowValues]);
    else sheet.appendRow(rowValues);
  });

  return { status: 'ok', message: 'Data tersimpan.' };
}

function getRecapMatrix(mode, params) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const rooms = configSheet.getRange('A2:A' + configSheet.getLastRow()).getValues().flat().filter(String);
  const configValues = configSheet.getRange('B2:E' + configSheet.getLastRow()).getValues().filter(function (r) { return r[0]; });
  const indicatorNames = configValues.map(function (r) { return r[0]; });

  let targetYear;
  if (mode === 'harian') targetYear = new Date(params.tanggal).getFullYear();
  else targetYear = Number(params.tahun);
  
  const data = getRawDataByYears(targetYear, targetYear);
  const totals = {};
  const roomPics = {};

  data.forEach(function (row) {
    const tgl = new Date(row[1]);
    let periodMatch = false;

    if (mode === 'harian') {
      periodMatch = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd') === params.tanggal;
    } else if (mode === 'mingguan') {
      const hari = tgl.getDate();
      const minggu = hari <= 7 ? 1 : hari <= 14 ? 2 : hari <= 21 ? 3 : 4;
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) && tgl.getFullYear() === Number(params.tahun) && minggu === Number(params.minggu);
    } else {
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) && tgl.getFullYear() === Number(params.tahun);
    }
    if (!periodMatch) return;

    const indKey = normText(row[3]);
    const roomKey = normText(row[2]);
    if (!totals[indKey]) totals[indKey] = {};
    if (!totals[indKey][roomKey]) totals[indKey][roomKey] = { numerator: 0, denominator: 0 };
    totals[indKey][roomKey].numerator += Number(row[4]) || 0;
    totals[indKey][roomKey].denominator += Number(row[5]) || 0;

    const picName = String(row[6] || '').trim();
    if (picName) {
      if (!roomPics[roomKey]) roomPics[roomKey] = {};
      roomPics[roomKey][picName.toUpperCase()] = picName;
    }
  });

  const rows = indicatorNames.map(function (ind) {
    const indKey = normText(ind);
    const values = rooms.map(function (room) {
      const roomKey = normText(room);
      const t = totals[indKey] && totals[indKey][roomKey];
      return t && t.denominator > 0 ? Math.round((t.numerator / t.denominator) * 1000) / 10 : null;
    });
    return { indikator: ind, values: values };
  });

  const pics = rooms.map(function (room) {
    const p = roomPics[normText(room)];
    return p ? Object.keys(p).sort().map(function (k) { return titleCase(p[k]); }).join(', ') : '';
  });

  return { rooms: rooms, rows: rows, pics: pics };
}

function getRecap(ruangan, mode, params) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const configValues = configSheet.getRange('B2:E' + configSheet.getLastRow()).getValues().filter(function (r) { return r[0]; });
  const targetMap = {};
  configValues.forEach(function (r) { targetMap[normText(r[0])] = (r[3] === '' || r[3] === undefined) ? null : Number(r[3]); });
  const REVERSE_INDICATORS = REVERSE_INDICATOR_NAMES.map(normText);

  let targetYear;
  if (mode === 'harian') targetYear = new Date(params.tanggal).getFullYear();
  else targetYear = Number(params.tahun);

  const data = getRawDataByYears(targetYear, targetYear);
  const totals = {};
  const targetRuangan = normText(ruangan);

  data.forEach(function (row) {
    if (ruangan !== 'SEMUA' && normText(row[2]) !== targetRuangan) return;
    const tgl = new Date(row[1]);
    let periodMatch = false;

    if (mode === 'harian') {
      periodMatch = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd') === params.tanggal;
    } else if (mode === 'mingguan') {
      const hari = tgl.getDate();
      const minggu = hari <= 7 ? 1 : hari <= 14 ? 2 : hari <= 21 ? 3 : 4;
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) && tgl.getFullYear() === Number(params.tahun) && minggu === Number(params.minggu);
    } else {
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) && tgl.getFullYear() === Number(params.tahun);
    }
    if (!periodMatch) return;

    const ind = normText(row[3]);
    if (!totals[ind]) totals[ind] = { numerator: 0, denominator: 0, pics: {}, alasanList: [] };
    totals[ind].numerator += Number(row[4]) || 0;
    totals[ind].denominator += Number(row[5]) || 0;
    
    const picName = String(row[6] || '').trim();
    if (picName) totals[ind].pics[picName.toUpperCase()] = picName;
    
    const ket = String(row[7] || '').trim();
    if (ket) totals[ind].alasanList.push(Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'dd/MM') + ': ' + ket);
  });

  return configValues.map(function (r) { return r[0]; }).filter(function (ind) { return totals[normText(ind)]; }).map(function (ind) {
    const key = normText(ind);
    const t = totals[key];
    const hasil = t.denominator > 0 ? (t.numerator / t.denominator * 100) : null;
    const target = targetMap[key] !== undefined ? targetMap[key] : null;
    let memenuhi = null;
    if (hasil !== null && target !== null) memenuhi = REVERSE_INDICATORS.indexOf(key) > -1 ? (hasil <= target) : (hasil >= target);
    
    return { 
      indikator: ind, numerator: t.numerator, denominator: t.denominator, 
      hasil: hasil, target: target, memenuhi: memenuhi, 
      pic: Object.keys(t.pics).sort().map(function (k) { return titleCase(t.pics[k]); }).join(', '), 
      alasan: t.alasanList.join(' | ') 
    };
  });
}

function computeReportData(startDateStr, endDateStr) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const configValues = configSheet.getRange('B2:E' + configSheet.getLastRow()).getValues().filter(function (r) { return r[0]; });
  const indicatorList = configValues.map(function (r) { return { nama: r[0], target: r[3] === '' || r[3] === undefined ? null : Number(r[3]) }; });

  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  endDate.setHours(23, 59, 59, 999);
  
  const rawData = getRawDataByYears(startDate.getFullYear(), endDate.getFullYear());
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];

  const isQ1Only = startDate.getMonth() === 0 && startDate.getDate() === 1 && endDate.getMonth() === 2 && endDate.getDate() >= 28 && startDate.getFullYear() === endDate.getFullYear();
  const buckets = [];

  if (isQ1Only) {
    let cursor = new Date(startDate);
    let weekNum = 1;
    while (cursor <= endDate) {
      const bStart = new Date(cursor);
      const bEnd = new Date(cursor);
      bEnd.setDate(bEnd.getDate() + 6);
      bEnd.setHours(23, 59, 59, 999);
      if (bEnd > endDate) bEnd.setTime(endDate.getTime());
      buckets.push({ label: 'Minggu ' + weekNum, start: bStart, end: bEnd });
      cursor.setDate(cursor.getDate() + 7);
      weekNum++;
    }
  } else {
    const rangeDays = (endDate - startDate) / (1000 * 60 * 60 * 24);
    if (rangeDays <= 35) {
      let cursor = new Date(startDate);
      let weekNum = 1;
      while (cursor <= endDate) {
        const bStart = new Date(cursor);
        const bEnd = new Date(cursor);
        bEnd.setDate(bEnd.getDate() + 6);
        bEnd.setHours(23, 59, 59, 999);
        if (bEnd > endDate) bEnd.setTime(endDate.getTime());
        buckets.push({ label: 'Minggu ' + weekNum, start: bStart, end: bEnd });
        cursor.setDate(cursor.getDate() + 7);
        weekNum++;
      }
    } else {
      let cursor = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
      while (cursor <= endDate) {
        const bStart = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
        const bEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0, 23, 59, 59, 999);
        buckets.push({ label: monthNames[cursor.getMonth()] + ' ' + cursor.getFullYear(), start: bStart, end: bEnd });
        cursor.setMonth(cursor.getMonth() + 1);
      }
    }
  }

  return indicatorList.map(function (ind) {
    const bucketResults = buckets.map(function (b) {
      let num = 0, den = 0;
      rawData.forEach(function (row) {
        if (normText(row[3]) !== normText(ind.nama)) return;
        const tgl = new Date(row[1]);
        if (tgl >= b.start && tgl <= b.end) {
          num += Number(row[4]) || 0;
          den += Number(row[5]) || 0;
        }
      });
      return { label: b.label, hasil: den > 0 ? Math.round((num / den) * 10000) / 100 : null };
    });

    const alasanList = [];
    rawData.forEach(function (row) {
      if (normText(row[3]) !== normText(ind.nama)) return;
      const tgl = new Date(row[1]);
      if (tgl < startDate || tgl > endDate) return;
      const ket = String(row[7] || '').trim();
      if (ket) alasanList.push(Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'dd/MM') + ' (' + row[2] + '): ' + ket);
    });

    return { nama: ind.nama, target: ind.target, buckets: bucketResults, alasan: alasanList.join('\n') };
  });
}

function buildPptPresentation(startDateStr, endDateStr) {
  const reportData = computeReportData(startDateStr, endDateStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempChart_' + new Date().getTime());

  const presTitle = 'Laporan Indikator Mutu ' + startDateStr + ' sd ' + endDateStr;
  const pres = SlidesApp.create(presTitle);
  const titleSlide = pres.getSlides()[0];
  titleSlide.getShapes().forEach(function (sh) { try { sh.remove(); } catch (e) { } });
  
  titleSlide.insertTextBox('LAPORAN INDIKATOR MUTU', 40, 180, 860, 80).getText().getTextStyle().setFontSize(32).setBold(true);
  titleSlide.insertTextBox(startDateStr + ' s.d ' + endDateStr, 40, 270, 860, 40).getText().getTextStyle().setFontSize(18);

  let chartRow = 1;
  reportData.forEach(function (ind) {
    const headerRow = chartRow;
    tempSheet.getRange(headerRow, 1, 1, 3).setValues([['Periode', 'Hasil (%)', 'Target (%)']]);
    const rows = ind.buckets.map(function (b) { return [b.label, b.hasil, ind.target]; });
    tempSheet.getRange(headerRow + 1, 1, rows.length, 3).setValues(rows);

    const dataRange = tempSheet.getRange(headerRow, 1, rows.length + 1, 3);
    const chart = tempSheet.newChart().asComboChart().addRange(dataRange).setNumHeaders(1)
      .setOption('title', '').setOption('legend', { position: 'bottom' })
      .setOption('series', { 
        0: { type: 'bars', color: '#f4b400', dataLabel: 'value' }, 
        // Bagian angka dihilangkan, hanya menampilkan garis lurus biasa
        1: { type: 'line', color: '#ea4335', lineWidth: 2, pointSize: 0 } 
      })
      .setPosition(headerRow, 6, 0, 0).build();
    tempSheet.insertChart(chart);
    SpreadsheetApp.flush();

    const chartsOnSheet = tempSheet.getCharts();
    const embeddedChart = chartsOnSheet[chartsOnSheet.length - 1];
    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    const margin = 20, contentWidth = pres.getPageWidth() - margin * 2;

    // Menentukan teks dan simbol target
    const isReverse = REVERSE_INDICATOR_NAMES.map(normText).indexOf(normText(ind.nama)) > -1;
    const targetValText = (ind.target !== null && ind.target !== undefined && ind.target !== '') 
      ? (isReverse ? '≤ ' : '≥ ') + ind.target + '%' 
      : '-';

    const leftBoxWidth = 80;
    const targetBoxWidth = 140;
    const nameBoxWidth = contentWidth - leftBoxWidth - targetBoxWidth;

    // Header Kiri (INM)
    const bannerLeft = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin, 15, leftBoxWidth, 36);
    bannerLeft.getFill().setSolidFill('#4a86c8');
    bannerLeft.getBorder().setTransparent();
    bannerLeft.getText().setText('INM').getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(14);
    bannerLeft.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerLeft.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    // Header Tengah (Nama Indikator)
    const bannerRight = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin + leftBoxWidth, 15, nameBoxWidth, 36);
    bannerRight.getFill().setSolidFill('#5a9bd8');
    bannerRight.getBorder().setTransparent();
    bannerRight.getText().setText(ind.nama).getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(14);
    bannerRight.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerRight.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    // Header Kanan (Badge Angka Target berwarna merah)
    const bannerTarget = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin + leftBoxWidth + nameBoxWidth, 15, targetBoxWidth, 36);
    bannerTarget.getFill().setSolidFill('#c5221f');
    bannerTarget.getBorder().setTransparent();
    bannerTarget.getText().setText('Target: ' + targetValText).getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(13);
    bannerTarget.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerTarget.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    slide.insertSheetsChart(embeddedChart, margin, 65, contentWidth, pres.getPageHeight() - 20 - 60 - 15 - 65);

    const noteBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin, pres.getPageHeight() - 20 - 60, contentWidth, 60);
    noteBox.getFill().setSolidFill('#d9ead3');
    noteBox.getBorder().setTransparent();
    noteBox.getText().setText((ind.alasan ? 'ANALISA (alasan dari ruangan):\n' + ind.alasan : 'ANALISA : [isi analisa di sini]') + '\nRTL : [isi rencana tindak lanjut di sini]');
    noteBox.getText().getTextStyle().setFontSize(11).setBold(true);
    noteBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    chartRow += rows.length + 3;
  });

  SpreadsheetApp.flush();
  return { presId: pres.getId(), filename: presTitle.replace(/[^a-zA-Z0-9]/g, '_') + '.pptx', tempSheetName: tempSheet.getName() };
}function buildPptPresentation(startDateStr, endDateStr) {
  const reportData = computeReportData(startDateStr, endDateStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempChart_' + new Date().getTime());

  const presTitle = 'Laporan Indikator Mutu ' + startDateStr + ' sd ' + endDateStr;
  const pres = SlidesApp.create(presTitle);
  const titleSlide = pres.getSlides()[0];
  titleSlide.getShapes().forEach(function (sh) { try { sh.remove(); } catch (e) { } });
  
  titleSlide.insertTextBox('LAPORAN INDIKATOR MUTU', 40, 180, 860, 80).getText().getTextStyle().setFontSize(32).setBold(true);
  titleSlide.insertTextBox(startDateStr + ' s.d ' + endDateStr, 40, 270, 860, 40).getText().getTextStyle().setFontSize(18);

  let chartRow = 1;
  reportData.forEach(function (ind) {
    const headerRow = chartRow;
    tempSheet.getRange(headerRow, 1, 1, 3).setValues([['Periode', 'Hasil (%)', 'Target (%)']]);
    const rows = ind.buckets.map(function (b) { return [b.label, b.hasil, ind.target]; });
    tempSheet.getRange(headerRow + 1, 1, rows.length, 3).setValues(rows);

    const dataRange = tempSheet.getRange(headerRow, 1, rows.length + 1, 3);
    const chart = tempSheet.newChart().asComboChart().addRange(dataRange).setNumHeaders(1)
      .setOption('title', '').setOption('legend', { position: 'bottom' })
      .setOption('series', { 
        0: { type: 'bars', color: '#f4b400', dataLabel: 'value' }, 
        // Bagian angka dihilangkan, hanya menampilkan garis lurus biasa
        1: { type: 'line', color: '#ea4335', lineWidth: 2, pointSize: 0 } 
      })
      .setPosition(headerRow, 6, 0, 0).build();
    tempSheet.insertChart(chart);
    SpreadsheetApp.flush();

    const chartsOnSheet = tempSheet.getCharts();
    const embeddedChart = chartsOnSheet[chartsOnSheet.length - 1];
    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    const margin = 20, contentWidth = pres.getPageWidth() - margin * 2;

    // Menentukan teks dan simbol target
    const isReverse = REVERSE_INDICATOR_NAMES.map(normText).indexOf(normText(ind.nama)) > -1;
    const targetValText = (ind.target !== null && ind.target !== undefined && ind.target !== '') 
      ? (isReverse ? '≤ ' : '≥ ') + ind.target + '%' 
      : '-';

    const leftBoxWidth = 80;
    const targetBoxWidth = 140;
    const nameBoxWidth = contentWidth - leftBoxWidth - targetBoxWidth;

    // Header Kiri (INM)
    const bannerLeft = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin, 15, leftBoxWidth, 36);
    bannerLeft.getFill().setSolidFill('#4a86c8');
    bannerLeft.getBorder().setTransparent();
    bannerLeft.getText().setText('INM').getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(14);
    bannerLeft.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerLeft.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    // Header Tengah (Nama Indikator)
    const bannerRight = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin + leftBoxWidth, 15, nameBoxWidth, 36);
    bannerRight.getFill().setSolidFill('#5a9bd8');
    bannerRight.getBorder().setTransparent();
    bannerRight.getText().setText(ind.nama).getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(14);
    bannerRight.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerRight.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    // Header Kanan (Badge Angka Target berwarna merah)
    const bannerTarget = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin + leftBoxWidth + nameBoxWidth, 15, targetBoxWidth, 36);
    bannerTarget.getFill().setSolidFill('#c5221f');
    bannerTarget.getBorder().setTransparent();
    bannerTarget.getText().setText('Target: ' + targetValText).getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(13);
    bannerTarget.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerTarget.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    slide.insertSheetsChart(embeddedChart, margin, 65, contentWidth, pres.getPageHeight() - 20 - 60 - 15 - 65);

    const noteBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin, pres.getPageHeight() - 20 - 60, contentWidth, 60);
    noteBox.getFill().setSolidFill('#d9ead3');
    noteBox.getBorder().setTransparent();
    noteBox.getText().setText((ind.alasan ? 'ANALISA (alasan dari ruangan):\n' + ind.alasan : 'ANALISA : [isi analisa di sini]') + '\nRTL : [isi rencana tindak lanjut di sini]');
    noteBox.getText().getTextStyle().setFontSize(11).setBold(true);
    noteBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    chartRow += rows.length + 3;
  });

  SpreadsheetApp.flush();
  return { presId: pres.getId(), filename: presTitle.replace(/[^a-zA-Z0-9]/g, '_') + '.pptx', tempSheetName: tempSheet.getName() };
}

function exportPresentationToPptx(presId, filename, tempSheetName) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + presId + '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.presentationml.presentation');
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) ss.deleteSheet(tempSheet);
  try { DriveApp.getFileById(presId).setTrashed(true); } catch (e) { }

  if (response.getResponseCode() !== 200) throw new Error('Gagal export PPT. Coba lagi.');
  return { base64: Utilities.base64Encode(response.getBlob().getBytes()), filename: filename };
}

function getMissingRooms(bulan, tahun) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const allRooms = configSheet.getRange('A2:A' + configSheet.getLastRow()).getValues().flat().filter(String);
  const rawData = getRawDataByYears(Number(tahun), Number(tahun));
  const roomsWithData = {};

  rawData.forEach(function (row) {
    const tgl = new Date(row[1]);
    if ((tgl.getMonth() + 1) === Number(bulan) && tgl.getFullYear() === Number(tahun)) roomsWithData[normText(row[2])] = true;
  });

  return allRooms.filter(function (room) { return !roomsWithData[normText(room)]; });
}

function testGeneratePptReport() {
  Logger.log('Mulai tes...');
  try {
    const built = buildPptPresentation('2026-01-01', '2026-12-31');
    Utilities.sleep(8000);
    const url = 'https://www.googleapis.com/drive/v3/files/' + built.presId + '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.presentationml.presentation');
    UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(built.tempSheetName));
    DriveApp.getFileById(built.presId).setTrashed(true);
  } catch (err) { Logger.log('GAGAL: ' + err.message); }
}

function testGeneratePptReportKeepFile() {
  const reportData = computeReportData('2026-01-01', '2026-12-31');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempChartDiag_' + new Date().getTime());
  const pres = SlidesApp.create('DIAGNOSA_Laporan_' + new Date().getTime());
  pres.getSlides()[0].insertTextBox('LAPORAN INDIKATOR MUTU (DIAGNOSA)', 40, 180, 860, 80);

  let chartRow = 1;
  reportData.forEach(function (ind, idx) {
    try {
      tempSheet.getRange(chartRow, 1, 1, 3).setValues([['Periode', 'Hasil (%)', 'Target (%)']]);
      const rows = ind.buckets.map(function (b) { return [b.label, b.hasil, ind.target]; });
      tempSheet.getRange(chartRow + 1, 1, rows.length, 3).setValues(rows);
      const chart = tempSheet.newChart().asLineChart().addRange(tempSheet.getRange(chartRow, 1, rows.length + 1, 3)).setNumHeaders(1).setOption('title', ind.nama).setPosition(chartRow, 6, 0, 0).build();
      tempSheet.insertChart(chart);
      SpreadsheetApp.flush();
      
      const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slide.insertTextBox(ind.nama, 20, 15, 900, 40);
      slide.insertSheetsChart(tempSheet.getCharts()[tempSheet.getCharts().length - 1], 20, 65, 520, 300);
      chartRow += rows.length + 3;
    } catch (err) { }
  });
  Logger.log('BUKA LINK INI BUAT CEK MANUAL: https://docs.google.com/presentation/d/' + pres.getId() + '/edit');
}

function buildFormattedSheet(title, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempExport_' + new Date().getTime());
  const numCols = headers.length;

  tempSheet.getRange(1, 1, 1, numCols).merge();
  tempSheet.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  tempSheet.getRange(2, 1, 1, numCols).setValues([headers]).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f0f2f5');
  if (rows.length > 0) tempSheet.getRange(3, 1, rows.length, numCols).setValues(rows).setHorizontalAlignment('center');
  
  tempSheet.getRange(2, 1, rows.length + 1, numCols).setBorder(true, true, true, true, true, true);
  tempSheet.autoResizeColumns(1, numCols);
  SpreadsheetApp.flush();
  return tempSheet;
}

function exportSheetAs(tempSheet, format) {
  const url = 'https://docs.google.com/spreadsheets/d/' + SpreadsheetApp.getActiveSpreadsheet().getId() + 
              (format === 'pdf' ? '/export?format=pdf&gid=' + tempSheet.getSheetId() + '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&attachment=true&horizontal_alignment=CENTER' 
                                : '/export?format=xlsx&gid=' + tempSheet.getSheetId());
  return UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
}

function exportRecapPdf(title, headers, rows) {
  const tempSheet = buildFormattedSheet(title, headers, rows);
  const blob = exportSheetAs(tempSheet, 'pdf');
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempSheet);
  return { base64: Utilities.base64Encode(blob.getBytes()), filename: title.replace(/[^a-zA-Z0-9]/g, '_') + '.pdf' };
}

function exportRecapExcel(title, headers, rows) {
  const tempSheet = buildFormattedSheet(title, headers, rows);
  const blob = exportSheetAs(tempSheet, 'xlsx');
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempSheet);
  return { base64: Utilities.base64Encode(blob.getBytes()), filename: title.replace(/[^a-zA-Z0-9]/g, '_') + '.xlsx' };
}

/**
 * =================================================================
 * FITUR GANTI PASSWORD & RESET PASSWORD
 * =================================================================
 */

function changePassword(username, oldPassword, newPassword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const inputUser = String(username || '').trim().toLowerCase();
  const inputOldPass = String(oldPassword || '').trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === inputUser) {
      if (String(data[i][1] || '').trim() === inputOldPass) {
        sheet.getRange(i + 1, 2).setValue(String(newPassword).trim());
        return { success: true, message: 'Mantap! Password berhasil diubah.' };
      } else {
        return { success: false, message: 'Gagal: Password lama salah!' };
      }
    }
  }
  return { success: false, message: 'Gagal: Username tidak ditemukan.' };
}

function resetPassword(username, ruangan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const inputUser = String(username || '').trim().toLowerCase();
  const inputRuangan = normText(ruangan);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === inputUser) {
      if (normText(data[i][3]) === inputRuangan) {
        sheet.getRange(i + 1, 2).setValue('inm123'); // Kembali ke password awal
        return { success: true, message: 'Berhasil! Password direset menjadi: inm123' };
      } else {
        return { success: false, message: 'Gagal: Username dan Ruangan tidak cocok!' };
      }
    }
  }
  return { success: false, message: 'Gagal: Username tidak ditemukan.' };
}
