/**
 * =================================================================
 * SISTEM INPUT INDIKATOR NASIONAL MUTU (INM) - BERBASIS DATA MASTER
 * =================================================================
 * Cara pakai singkat:
 * 1. Buka Google Sheet rekap kamu (atau bikin Sheet baru khusus sistem ini).
 * 2. Buat 2 sheet baru dengan nama PERSIS seperti ini:
 *    - "Config"       -> berisi daftar ruangan & indikator
 *    - "Data Master"  -> tempat semua data harian tersimpan (jangan diisi manual)
 * 3. Buka menu Extensions > Apps Script, hapus isi default, lalu paste file ini.
 * 4. Buat file HTML baru bernama "Form" (File > New > HTML), paste isi Form.html.
 * 5. Klik Deploy > New deployment > pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy link web app yang dihasilkan, bagikan ke semua PIC ruangan
 *    (tambahkan ?ruangan=NAMARUANGAN di belakang link untuk mengunci ruangan).
 * =================================================================
 */

/**
 * =================================================================
 * DAFTAR ISI FUNGSI — buat cari cepat kalau ada masalah
 * =================================================================
 * Ctrl+F nama fungsinya di kotak pencarian editor buat langsung loncat.
 *
 * -- SETUP AWAL (jalankan manual 1x kalau bikin sheet baru) --
 * setupSheets()            -> bikin sheet Config & Data Master otomatis
 * setupUsers()              -> bikin sheet Users + akun default tiap ruangan
 *
 * -- WEB APP & LOGIN --
 * doGet()                   -> dipanggil otomatis saat link web app dibuka
 * login(username, password) -> cek username/password waktu orang login
 *
 * -- BACA DATA BUAT FORM --
 * getRooms()                 -> ambil daftar ruangan (buat dropdown)
 * getIndicatorDetails()      -> ambil daftar indikator + keterangan + target
 * getExistingEntries()       -> cek data yg SUDAH ada (buat Mode Edit di form)
 *
 * -- SIMPAN DATA --
 * submitData(payload)        -> dipanggil pas PIC klik "Simpan Data"
 *                               (kalau data sudah ada, ditimpa bukan dobel)
 *
 * -- REKAP & LAPORAN --
 * getRecap()                 -> rekap 1 ruangan (harian/mingguan/bulanan)
 * getRecapMatrix()            -> rekap SEMUA ruangan sekaligus (fitur admin)
 * getMissingRooms()           -> cek ruangan yg belum isi data bulan tertentu
 * computeReportData()         -> hitung data buat bikin grafik laporan PPT
 * buildPptPresentation()      -> LANGKAH 1 bikin laporan PPT (bikin slide+grafik)
 * exportPresentationToPptx()  -> LANGKAH 2 bikin laporan PPT (export ke file)
 *
 * -- EXPORT FILE --
 * buildFormattedSheet()       -> siapin sheet sementara yg rapi buat export
 * exportSheetAs()              -> proses export sheet sementara jadi PDF/Excel
 * exportRecapPdf()              -> tombol "Unduh PDF" di tab Lihat Rekap
 * exportRecapExcel()             -> tombol "Unduh Excel" di tab Lihat Rekap
 *
 * -- FUNGSI BANTUAN (dipakai fungsi lain, jarang perlu disentuh) --
 * normText(s)                 -> "membersihkan" teks sebelum dibandingkan,
 *                                 dipakai di HAMPIR SEMUA fungsi di atas biar
 *                                 nama ruangan/indikator yg beda spasi/huruf
 *                                 besar-kecil tetap dianggap sama
 *
 * -- KHUSUS DIAGNOSA (boleh diabaikan kalau semua normal) --
 * testGeneratePptReport()          -> tes bikin PPT langsung dari editor
 * testGeneratePptReportKeepFile()  -> sama, tapi file Slides-nya disimpan
 *                                     (buat cek manual kalau PPT kosong lagi)
 *
 * -- PANDUAN "ADA MASALAH DI MANA, CEK FUNGSI APA" --
 * "Rekap kosong padahal data ada"     -> cek getRecap() atau getRecapMatrix()
 * "Login gagal"                        -> cek login() dan sheet Users
 * "Data ganda / numpuk"                 -> cek submitData() bagian upsert
 * "Mode Edit tapi kolom kosong"          -> cek getExistingEntries() DAN
 *                                          bagian loadExisting() di Form.html
 *                                          (dua-duanya harus samakan cara
 *                                          "membersihkan" teksnya)
 * "PPT kosong/gagal"                    -> cek buildPptPresentation() dan
 *                                          exportPresentationToPptx()
 * "Target/Status salah di rekap"         -> cek sheet Config kolom Target
 * =================================================================
 */

const CONFIG_SHEET_NAME = 'Config';
const DATA_SHEET_NAME = 'Data Master';
const USERS_SHEET_NAME = 'Users';

// Header kolom di sheet "Data Master". JANGAN diubah urutannya
// tanpa menyesuaikan kode di bawah.
const DATA_HEADERS = ['Timestamp', 'Tanggal', 'Ruangan', 'Indikator', 'Numerator', 'Denominator', 'Diisi Oleh', 'Keterangan'];

// Indikator yang arah targetnya kebalik: target itu batas MAKSIMAL,
// bukan minimal (misal Penundaan Operasi Elektif, target <5% berarti
// makin KECIL makin bagus, bukan makin besar).
const REVERSE_INDICATOR_NAMES = ['Penundaan Operasi Elektif'];

/**
 * Membersihkan teks sebelum dibandingkan — membuang SEMUA karakter
 * selain huruf dan angka (termasuk spasi biasa, spasi ganda, dan
 * karakter tak terlihat yang kadang kebawa dari copy-paste), lalu
 * menyamakan ke huruf besar. Ini paling aman untuk mencocokkan nama
 * ruangan/indikator walau ketikannya sedikit beda.
 */
function normText(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Menyeragamkan tampilan nama (misal "novi" atau "NOVI" jadi "Novi"),
 * dipakai buat nampilin daftar nama PIC di rekap supaya nama yang sama
 * (beda huruf besar/kecil) nggak keitung dobel/beda.
 */
function titleCase(s) {
  return String(s || '').trim().toLowerCase().replace(/(^|\s)\S/g, function (c) { return c.toUpperCase(); });
}

// Keterangan Numerator & Denominator untuk tiap indikator.
const INDICATOR_DETAILS = [
  {
    nama: 'Kepatuhan Penggunaan APD',
    num: 'Jumlah petugas yang patuh menggunakan APD sesuai indikasi dalam periode observasi',
    den: 'Jumlah seluruh petugas yang terindikasi menggunakan APD dalam periode observasi'
  },
  {
    nama: 'Kepatuhan Identifikasi Pasien',
    num: 'Jumlah pemberi pelayanan yang melakukan identifikasi pasien secara benar dalam periode observasi',
    den: 'Jumlah pemberi pelayanan yang diobservasi dalam periode observasi'
  },
  {
    nama: 'Waktu Tanggap Operasi Seksio Sesarea Emergensi',
    num: 'Jumlah pasien yang diputuskan tindakan seksio sesarea emergensi kategori I (satu) yang mendapatkan tindakan seksio sesarea emergensi ≤ 30 menit',
    den: 'Jumlah pasien yang diputuskan tindakan seksio sesarea emergensi kategori I (satu)'
  },
  {
    nama: 'Waktu Tunggu Rawat Jalan',
    num: 'Jumlah pasien rawat jalan dengan waktu tunggu ≤ 60 menit',
    den: 'Jumlah pasien rawat jalan yang diobservasi'
  },
  {
    nama: 'Penundaan Operasi Elektif',
    num: 'Jumlah pasien yang waktu jadwal operasinya tertunda lebih dari 1 jam',
    den: 'Jumlah pasien operasi elektif'
  },
  {
    nama: 'Ketepatan Waktu Visite Dokter',
    num: 'Jumlah di-visite Dokter pada pukul 06.00 - 14:00',
    den: 'Jumlah pasien yang diobservasi'
  },
  {
    nama: 'Pelaporan Hasil Kritis Laboratorium',
    num: 'Jumlah hasil kritis laboratorium yang dilaporkan ≤30 menit',
    den: 'Jumlah hasil kritis laboratorium yang diobservasi'
  },
  {
    nama: 'Kepatuhan Penggunaan Formularium Nasional',
    num: 'Jumlah R/ recipe dalam lembar resep yang sesuai dengan formularium nasional',
    den: 'Jumlah R/ recipe dalam lembar resep yang diobservasi'
  },
  {
    nama: 'Kepatuhan Terhadap Alur Klinis (Clinical Pathway)',
    num: 'Jumlah pelayanan oleh PPA yang sesuai dengan clinical pathway',
    den: 'Jumlah seluruh pelayanan oleh PPA pada clinical pathway yang diobservasi'
  },
  {
    nama: 'Kepatuhan Upaya Pencegahan Risiko Pasien Jatuh',
    num: 'Jumlah pasien rawat inap berisiko tinggi jatuh yang mendapatkan ketiga upaya pencegahan risiko jatuh',
    den: 'Jumlah pasien rawat inap berisiko tinggi jatuh yang diobservasi'
  },
  {
    nama: 'Kecepatan Waktu Tanggap Komplain',
    num: 'Jumlah komplain yang ditanggapi dan ditindaklanjuti sesuai waktu yang ditetapkan berdasarkan grading',
    den: 'Jumlah Komplain yang disurvei'
  },
  {
    nama: 'Kepuasan Pasien',
    num: 'Total nilai persepsi seluruh responden',
    den: 'Total unsur yang terisi dari seluruh responden'
  }
];

/**
 * Dipanggil otomatis saat web app dibuka lewat link.
 * Sekarang semua orang (PIC maupun admin) pakai 1 link yang sama —
 * ruangan/peran ditentukan lewat login, bukan lewat parameter URL lagi.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Form')
    .setTitle('Input Data INM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Jalankan fungsi ini SEKALI untuk mengisi semua 13 target indikator
 * INM sekaligus di sheet Config kolom E, sesuai acuan resmi RS.
 * Catatan: "Penundaan Operasi Elektif" arah targetnya kebalik
 * (target di BAWAH 5%, bukan di atas), tapi tetap disimpan sebagai
 * angka 5 di kolom Target.
 */
function isiSemuaTarget() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  sheet.getRange('E1').setValue('Target (%)');

  const targets = {
    'Kepatuhan Kebersihan Tangan': 85,
    'Kepatuhan Penggunaan APD': 100,
    'Kepatuhan Identifikasi Pasien': 100,
    'Waktu Tanggap Operasi Seksio Sesarea Emergensi': 80,
    'Waktu Tunggu Rawat Jalan': 80,
    'Penundaan Operasi Elektif': 5,
    'Ketepatan Waktu Visite Dokter': 80,
    'Pelaporan Hasil Kritis Laboratorium': 100,
    'Kepatuhan Penggunaan Formularium Nasional': 80,
    'Kepatuhan Terhadap Alur Klinis (Clinical Pathway)': 80,
    'Kepatuhan Upaya Pencegahan Risiko Pasien Jatuh': 100,
    'Kecepatan Waktu Tanggap Komplain': 80,
    'Kepuasan Pasien': 76.61
  };

  const lastRow = sheet.getLastRow();
  const names = sheet.getRange('B2:B' + lastRow).getValues();
  let filled = 0;
  let notFound = [];

  names.forEach(function (row, i) {
    const nama = row[0];
    if (targets[nama] !== undefined) {
      sheet.getRange(i + 2, 5).setValue(targets[nama]);
      filled++;
    } else if (nama) {
      notFound.push(nama);
    }
  });

  Logger.log('Berhasil isi ' + filled + ' target.');
  if (notFound.length > 0) {
    Logger.log('Nama indikator yang TIDAK cocok (cek ejaan di Config kolom B): ' + notFound.join(', '));
  }
}

/**
 * Jalankan fungsi ini SEKALI untuk menambahkan kolom "Target (%)"
 * di sheet Config (kolom E), kalau sheet Config kamu sudah ada
 * sebelum fitur laporan ini dibuat. Isi manual target tiap indikator
 * setelah kolom ini muncul.
 */
function tambahKolomTarget() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  sheet.getRange('E1').setValue('Target (%)');
  const knownTargets = {
    'Kepatuhan Kebersihan Tangan': 85,
    'Kepatuhan Penggunaan APD': 100,
    'Kepatuhan Identifikasi Pasien': 100
  };
  const lastRow = sheet.getLastRow();
  const names = sheet.getRange('B2:B' + lastRow).getValues();
  names.forEach(function (row, i) {
    const nama = row[0];
    if (knownTargets[nama] !== undefined) {
      sheet.getRange(i + 2, 5).setValue(knownTargets[nama]);
    }
  });
}

/**
 * Jalankan fungsi ini SEKALI untuk membuat sheet "Users" dan mengisi
 * akun default: 1 akun per ruangan + 1 akun admin.
 * Username dibuat otomatis dari nama ruangan (huruf kecil, tanpa spasi).
 * Password default disamakan dulu, WAJIB diganti manual nanti di sheet
 * Users kolom B setelah dibagikan ke tiap ruangan.
 */
function setupUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(USERS_SHEET_NAME);
  }
  if (usersSheet.getLastRow() > 0) {
    Logger.log('Sheet Users sudah ada isinya, tidak ditimpa. Hapus manual dulu kalau mau reset ulang.');
    return;
  }

  usersSheet.appendRow(['Username', 'Password', 'Role', 'Ruangan']);
  usersSheet.setFrozenRows(1);

  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const lastRow = configSheet.getLastRow();
  const rooms = configSheet.getRange('A2:A' + lastRow).getValues().flat().filter(String);

  const rows = rooms.map(function (room) {
    const username = room.toLowerCase().replace(/[^a-z0-9]/g, '');
    return [username, 'inm123', 'pic', room];
  });
  rows.push(['admin', 'admin123', 'admin', '']);

  usersSheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

/**
 * Dipanggil dari Form.html saat orang login. Mencocokkan username +
 * password ke sheet Users (tidak peka besar/kecil huruf untuk username).
 */
function login(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Sistem login belum di-setup.' };

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const inputUser = String(username || '').trim().toLowerCase();
  const inputPass = String(password || '').trim();

  for (let i = 0; i < data.length; i++) {
    const rowUser = String(data[i][0] || '').trim().toLowerCase();
    const rowPass = String(data[i][1] || '').trim();
    if (rowUser === inputUser && rowPass === inputPass) {
      return { success: true, role: data[i][2], ruangan: data[i][3] };
    }
  }
  return { success: false, message: 'Username atau password salah.' };
}

/**
 * Jalankan fungsi ini SEKALI secara manual (kalau sheet Config/Data Master
 * belum pernah dibuat sama sekali) untuk membuat semuanya otomatis.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    dataSheet = ss.insertSheet(DATA_SHEET_NAME);
  }
  if (dataSheet.getLastRow() === 0) {
    dataSheet.appendRow(DATA_HEADERS);
    dataSheet.setFrozenRows(1);
  }

  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
  }
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

/** Dipanggil dari Form.html untuk mengisi dropdown ruangan. */
function getRooms() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const values = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat().filter(String);
  return values;
}

/** Dipanggil dari Form.html untuk mengisi nama + keterangan tiap indikator. */
function getIndicatorDetails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange('B2:E' + lastRow).getValues();
  return data
    .filter(function (row) { return row[0]; })
    .map(function (row) {
      const target = (row[3] === '' || row[3] === undefined) ? null : Number(row[3]);
      const reverse = REVERSE_INDICATOR_NAMES.map(normText).indexOf(normText(row[0])) > -1;
      return { nama: row[0], num: row[1] || '', den: row[2] || '', target: target, tipe: 'persen', reverse: reverse };
    });
}

/**
 * Dipanggil dari Form.html untuk mengambil data yang sudah pernah
 * diisi (kalau ada) untuk ruangan + tanggal tertentu, supaya bisa diedit
 * atau ditampilkan di rekap harian.
 */
function getExistingEntries(ruangan, tanggal) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
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

/**
 * Dipanggil dari Form.html tab "Lihat Rekap" (mode Bulanan) untuk
 * menampilkan total Numerator/Denominator/Hasil per indikator.
 */
function getMonthlyRecap(ruangan, bulan, tahun) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
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
    const hasil = t.denominator > 0 ? (t.numerator / t.denominator * 100) : null;
    return { indikator: ind, numerator: t.numerator, denominator: t.denominator, hasil: hasil };
  });
}

/**
 * Menyimpan data yang dikirim dari form. Kalau data untuk
 * ruangan+tanggal+indikator yang sama sudah pernah diisi, baris lama
 * akan DIPERBARUI (bukan menumpuk duplikat).
 */
function submitData(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const lastRow = sheet.getLastRow();

  const existing = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues()
    : [];

  const tanggal = payload.tanggal;
  const ruangan = payload.ruangan;
  const targetRuangan = normText(ruangan);
  const now = new Date();

  payload.entries.forEach(function (entry) {
    if (entry.numerator === '' && entry.denominator === '') {
      return;
    }

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

    if (foundRowIndex > -1) {
      sheet.getRange(foundRowIndex, 1, 1, DATA_HEADERS.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  });

  return { status: 'ok', message: 'Data tersimpan.' };
}

/**
 * Fungsi rekap terpadu, dipakai tab "Lihat Rekap".
 * ruangan: nama ruangan tertentu, ATAU 'SEMUA' untuk gabungan semua ruangan (Total RS).
 * mode: 'harian' | 'mingguan' | 'bulanan'
 * params: { tanggal } untuk harian, { bulan, tahun, minggu } untuk mingguan,
 *         { bulan, tahun } untuk bulanan.
 * minggu 1 = tanggal 1-7, minggu 2 = tanggal 8-14, minggu 3 = tanggal 15-21, minggu 4 = tanggal 22-31.
 */
/**
 * Buat admin: rekap SEMUA ruangan sekaligus dalam 1 tabel, tiap ruangan
 * jadi kolom terpisah (bukan digabung jadi 1 angka total).
 */
function getRecapMatrix(mode, params) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastConfigRow = configSheet.getLastRow();
  const rooms = configSheet.getRange('A2:A' + lastConfigRow).getValues().flat().filter(String);
  const configValues = configSheet.getRange('B2:E' + lastConfigRow).getValues().filter(function (r) { return r[0]; });
  const indicatorNames = configValues.map(function (r) { return r[0]; });

  const lastRow = dataSheet.getLastRow();
  const data = lastRow >= 2 ? dataSheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues() : [];

  const totals = {};
  const roomPics = {};

  data.forEach(function (row) {
    const tgl = new Date(row[1]);
    let periodMatch = false;

    if (mode === 'harian') {
      const rowTanggal = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      periodMatch = rowTanggal === params.tanggal;
    } else if (mode === 'mingguan') {
      const hari = tgl.getDate();
      const minggu = hari <= 7 ? 1 : hari <= 14 ? 2 : hari <= 21 ? 3 : 4;
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) &&
        tgl.getFullYear() === Number(params.tahun) &&
        minggu === Number(params.minggu);
    } else {
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) &&
        tgl.getFullYear() === Number(params.tahun);
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
      if (!t) return null;
      return t.denominator > 0 ? Math.round((t.numerator / t.denominator) * 1000) / 10 : null;
    });
    return { indikator: ind, values: values };
  });

  const pics = rooms.map(function (room) {
    const roomKey = normText(room);
    const p = roomPics[roomKey];
    if (!p) return '';
    return Object.keys(p).sort().map(function (k) { return titleCase(p[k]); }).join(', ');
  });

  return { rooms: rooms, rows: rows, pics: pics };
}

function getRecap(ruangan, mode, params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastConfigRow = configSheet.getLastRow();
  const configValues = configSheet.getRange('B2:E' + lastConfigRow).getValues().filter(function (r) { return r[0]; });
  const targetMap = {};
  configValues.forEach(function (r) {
    targetMap[normText(r[0])] = (r[3] === '' || r[3] === undefined) ? null : Number(r[3]);
  });
  // Indikator yang arah targetnya kebalik: target itu batas MAKSIMAL, bukan minimal.
  const REVERSE_INDICATORS = REVERSE_INDICATOR_NAMES.map(normText);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
  const totals = {};
  const targetRuangan = normText(ruangan);

  data.forEach(function (row) {
    const roomMatch = (ruangan === 'SEMUA') || (normText(row[2]) === targetRuangan);
    if (!roomMatch) return;

    const tgl = new Date(row[1]);
    let periodMatch = false;

    if (mode === 'harian') {
      const rowTanggal = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      periodMatch = rowTanggal === params.tanggal;
    } else if (mode === 'mingguan') {
      const hari = tgl.getDate();
      const minggu = hari <= 7 ? 1 : hari <= 14 ? 2 : hari <= 21 ? 3 : 4;
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) &&
        tgl.getFullYear() === Number(params.tahun) &&
        minggu === Number(params.minggu);
    } else {
      periodMatch = (tgl.getMonth() + 1) === Number(params.bulan) &&
        tgl.getFullYear() === Number(params.tahun);
    }
    if (!periodMatch) return;

    const ind = normText(row[3]);
    if (!totals[ind]) totals[ind] = { numerator: 0, denominator: 0, pics: {}, alasanList: [] };
    totals[ind].numerator += Number(row[4]) || 0;
    totals[ind].denominator += Number(row[5]) || 0;
    const picName = String(row[6] || '').trim();
    if (picName) totals[ind].pics[picName.toUpperCase()] = picName;
    const ket = String(row[7] || '').trim();
    if (ket) {
      const tglStr = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'dd/MM');
      totals[ind].alasanList.push(tglStr + ': ' + ket);
    }
  });

  return configValues
    .map(function (r) { return r[0]; })
    .filter(function (ind) { return totals[normText(ind)]; })
    .map(function (ind) {
      const key = normText(ind);
      const t = totals[key];
      const hasil = t.denominator > 0 ? (t.numerator / t.denominator * 100) : null;
      const target = targetMap[key] !== undefined ? targetMap[key] : null;
      const picList = Object.keys(t.pics).sort().map(function (k) { return titleCase(t.pics[k]); }).join(', ');
      const alasan = t.alasanList.join(' | ');
      let memenuhi = null;
      if (hasil !== null && target !== null) {
        memenuhi = REVERSE_INDICATORS.indexOf(key) > -1 ? (hasil <= target) : (hasil >= target);
      }
      return { indikator: ind, numerator: t.numerator, denominator: t.denominator, hasil: hasil, target: target, memenuhi: memenuhi, pic: picList, alasan: alasan };
    });
}

/**
 * Menghitung data laporan untuk rentang tanggal bebas, digabung dari
 * SEMUA ruangan, dipecah otomatis per minggu (kalau rentang pendek)
 * atau per bulan (kalau rentang panjang).
 */
function computeReportData(startDateStr, endDateStr) {
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastConfigRow = configSheet.getLastRow();
  const configValues = configSheet.getRange('B2:E' + lastConfigRow).getValues().filter(function (r) { return r[0]; });
  const indicatorList = configValues.map(function (r) {
    return { nama: r[0], target: r[3] === '' || r[3] === undefined ? null : Number(r[3]) };
  });

  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  endDate.setHours(23, 59, 59, 999);

  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];

  // Mode Triwulan 1: kalau rentang persis 1 Jan - 31 Mar tahun yang sama,
  // dipecah per minggu. Laporan tahunan penuh (1 Jan - 31 Des) TETAP per bulan.
  const isQ1Only = startDate.getMonth() === 0 && startDate.getDate() === 1 &&
    endDate.getMonth() === 2 && endDate.getDate() >= 28 &&
    startDate.getFullYear() === endDate.getFullYear();

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

  const lastDataRow = dataSheet.getLastRow();
  const rawData = lastDataRow >= 2
    ? dataSheet.getRange(2, 1, lastDataRow - 1, DATA_HEADERS.length).getValues()
    : [];

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
      const hasil = den > 0 ? Math.round((num / den) * 10000) / 100 : null;
      return { label: b.label, hasil: hasil };
    });

    const alasanList = [];
    rawData.forEach(function (row) {
      if (normText(row[3]) !== normText(ind.nama)) return;
      const tgl = new Date(row[1]);
      if (tgl < startDate || tgl > endDate) return;
      const ket = String(row[7] || '').trim();
      if (ket) {
        const tglStr = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'dd/MM');
        alasanList.push(tglStr + ' (' + row[2] + '): ' + ket);
      }
    });

    return { nama: ind.nama, target: ind.target, buckets: bucketResults, alasan: alasanList.join('\n') };
  });
}

/**
 * Langkah 1: bikin presentasi Slides lengkap dengan semua slide & grafik.
 * TIDAK export ke PPT di sini, dan file BELUM dihapus.
 */
function buildPptPresentation(startDateStr, endDateStr) {
  const reportData = computeReportData(startDateStr, endDateStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempChart_' + new Date().getTime());

  const presTitle = 'Laporan Indikator Mutu ' + startDateStr + ' sd ' + endDateStr;
  const pres = SlidesApp.create(presTitle);

  const titleSlide = pres.getSlides()[0];
  titleSlide.getShapes().forEach(function (sh) {
    try { sh.remove(); } catch (e) { /* ignore */ }
  });
  const titleBox = titleSlide.insertTextBox('LAPORAN INDIKATOR MUTU', 40, 180, 860, 80);
  titleBox.getText().getTextStyle().setFontSize(32).setBold(true);
  const subBox = titleSlide.insertTextBox(startDateStr + ' s.d ' + endDateStr, 40, 270, 860, 40);
  subBox.getText().getTextStyle().setFontSize(18);

  let chartRow = 1;
  reportData.forEach(function (ind) {
    const headerRow = chartRow;
    tempSheet.getRange(headerRow, 1, 1, 3).setValues([['Periode', 'Hasil (%)', 'Target (%)']]);
    const rows = ind.buckets.map(function (b) {
      return [b.label, b.hasil, ind.target];
    });
    tempSheet.getRange(headerRow + 1, 1, rows.length, 3).setValues(rows);

    const dataRange = tempSheet.getRange(headerRow, 1, rows.length + 1, 3);
    const chartBuilder = tempSheet.newChart()
      .asComboChart()
      .addRange(dataRange)
      .setNumHeaders(1)
      .setOption('title', '')
      .setOption('legend', { position: 'bottom' })
      .setOption('series', {
        0: { type: 'bars', color: '#f4b400', dataLabel: 'value' },
        1: { type: 'line', color: '#ea4335', lineWidth: 2, pointSize: 0 }
      })
      .setPosition(headerRow, 6, 0, 0);
    const chart = chartBuilder.build();
    tempSheet.insertChart(chart);
    SpreadsheetApp.flush();

    const chartsOnSheet = tempSheet.getCharts();
    const embeddedChart = chartsOnSheet[chartsOnSheet.length - 1];

    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    const pageWidth = pres.getPageWidth();
    const pageHeight = pres.getPageHeight();
    const margin = 20;
    const contentWidth = pageWidth - margin * 2;
    const leftBannerWidth = 100;
    const rightBannerWidth = contentWidth - leftBannerWidth;

    const bottomMargin = 20;
    const noteBoxHeight = 60;
    const gapBeforeNote = 15;
    const chartTop = 65;
    const noteBoxTop = pageHeight - bottomMargin - noteBoxHeight;
    const chartHeight = noteBoxTop - gapBeforeNote - chartTop;

    const bannerLeft = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin, 15, leftBannerWidth, 36);
    bannerLeft.getFill().setSolidFill('#4a86c8');
    bannerLeft.getBorder().setTransparent();
    bannerLeft.getText().setText('INM');
    bannerLeft.getText().getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(14);
    bannerLeft.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerLeft.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    const bannerRight = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin + leftBannerWidth, 15, rightBannerWidth, 36);
    bannerRight.getFill().setSolidFill('#5a9bd8');
    bannerRight.getBorder().setTransparent();
    bannerRight.getText().setText(ind.nama);
    bannerRight.getText().getTextStyle().setForegroundColor('#ffffff').setBold(true).setFontSize(14);
    bannerRight.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    bannerRight.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    slide.insertSheetsChart(embeddedChart, margin, chartTop, contentWidth, chartHeight);

    const noteBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, margin, noteBoxTop, contentWidth, noteBoxHeight);
    noteBox.getFill().setSolidFill('#d9ead3');
    noteBox.getBorder().setTransparent();
    const analisaText = ind.alasan
      ? 'ANALISA (alasan dari ruangan):\n' + ind.alasan
      : 'ANALISA : [isi analisa di sini]';
    noteBox.getText().setText(analisaText + '\nRTL : [isi rencana tindak lanjut di sini]');
    noteBox.getText().getTextStyle().setFontSize(11).setBold(true);
    noteBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    chartRow += rows.length + 3;
  });

  SpreadsheetApp.flush();

  return {
    presId: pres.getId(),
    filename: presTitle.replace(/[^a-zA-Z0-9]/g, '_') + '.pptx',
    tempSheetName: tempSheet.getName()
  };
}

/**
 * Langkah 2: export presentasi yang sudah jadi (dari buildPptPresentation)
 * ke format PPTX, lalu bersihkan file & sheet sementaranya. Dipanggil
 * terpisah beberapa detik setelah langkah 1 supaya Google sempat
 * "menyimpan" presentasinya secara utuh dulu sebelum di-export.
 */
function exportPresentationToPptx(presId, filename, tempSheetName) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + presId +
    '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.presentationml.presentation');
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  const blob = response.getBlob();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) ss.deleteSheet(tempSheet);
  try { DriveApp.getFileById(presId).setTrashed(true); } catch (e) { /* ignore kalau gagal hapus */ }

  if (responseCode !== 200) {
    throw new Error('Gagal export PPT (kode ' + responseCode + '). Coba lagi.');
  }

  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: filename
  };
}


/**
 * Cek ruangan mana saja yang SAMA SEKALI belum ada data di bulan/tahun
 * tertentu (belum submit sama sekali, bukan cuma sebagian indikator).
 */
function getMissingRooms(bulan, tahun) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  const lastConfigRow = configSheet.getLastRow();
  const allRooms = configSheet.getRange('A2:A' + lastConfigRow).getValues().flat().filter(String);

  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET_NAME);
  const lastDataRow = dataSheet.getLastRow();
  const roomsWithData = {};

  if (lastDataRow >= 2) {
    const rawData = dataSheet.getRange(2, 1, lastDataRow - 1, DATA_HEADERS.length).getValues();
    rawData.forEach(function (row) {
      const tgl = new Date(row[1]);
      if ((tgl.getMonth() + 1) === Number(bulan) && tgl.getFullYear() === Number(tahun)) {
        roomsWithData[normText(row[2])] = true;
      }
    });
  }

  return allRooms.filter(function (room) { return !roomsWithData[normText(room)]; });
}

/**
 * DIAGNOSA: jalankan fungsi ini langsung dari editor Apps Script
 * (pilih di dropdown fungsi, klik Jalankan) untuk tes bikin PPT tanpa
 * lewat web app. Setelah selesai, buka "Log eksekusi" untuk lihat
 * detail prosesnya dan cek apakah ada error.
 */
function testGeneratePptReport() {
  Logger.log('Mulai tes...');
  try {
    const built = buildPptPresentation('2026-01-01', '2026-12-31');
    Logger.log('Slide selesai dibuat. presId: ' + built.presId);
    Utilities.sleep(8000);

    const url = 'https://www.googleapis.com/drive/v3/files/' + built.presId +
      '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.presentationml.presentation');
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    Logger.log('Response code: ' + response.getResponseCode());
    Logger.log('Response headers: ' + JSON.stringify(response.getHeaders()));
    const contentType = response.getHeaders()['Content-Type'] || response.getHeaders()['content-type'] || '';
    Logger.log('Content-Type: ' + contentType);

    if (contentType.indexOf('text') > -1 || contentType.indexOf('html') > -1 || contentType.indexOf('json') > -1) {
      Logger.log('ISI RESPONSE (kemungkinan error, bukan file PPT): ' + response.getContentText().substring(0, 1000));
    } else {
      Logger.log('Content-Type terlihat seperti file biner (kemungkinan pptx beneran). Ukuran: ' + response.getBlob().getBytes().length + ' bytes.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tempSheet = ss.getSheetByName(built.tempSheetName);
    if (tempSheet) ss.deleteSheet(tempSheet);
    DriveApp.getFileById(built.presId).setTrashed(true);
  } catch (err) {
    Logger.log('GAGAL dengan error: ' + err.message);
    Logger.log('Detail: ' + err.stack);
  }
}

/**
 * DIAGNOSA LANJUTAN: sama seperti generatePptReport, tapi file Google
 * Slides sementaranya TIDAK dihapus, supaya bisa dibuka manual dan
 * dicek langsung apakah slide & grafiknya beneran ada atau kosong.
 * Setelah selesai, cek Log eksekusi untuk link-nya, buka link itu.
 * INGAT: hapus manual file ini nanti dari Drive setelah selesai dicek.
 */
function testGeneratePptReportKeepFile() {
  const reportData = computeReportData('2026-01-01', '2026-12-31');
  Logger.log('Jumlah indikator: ' + reportData.length);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempChartDiag_' + new Date().getTime());
  const pres = SlidesApp.create('DIAGNOSA_Laporan_' + new Date().getTime());

  const titleSlide = pres.getSlides()[0];
  titleSlide.getShapes().forEach(function (sh) {
    try { sh.remove(); } catch (e) { Logger.log('Gagal hapus shape judul: ' + e.message); }
  });
  titleSlide.insertTextBox('LAPORAN INDIKATOR MUTU (DIAGNOSA)', 40, 180, 860, 80);

  let chartRow = 1;
  reportData.forEach(function (ind, idx) {
    try {
      const headerRow = chartRow;
      tempSheet.getRange(headerRow, 1, 1, 3).setValues([['Periode', 'Hasil (%)', 'Target (%)']]);
      const rows = ind.buckets.map(function (b) { return [b.label, b.hasil, ind.target]; });
      tempSheet.getRange(headerRow + 1, 1, rows.length, 3).setValues(rows);

      const dataRange = tempSheet.getRange(headerRow, 1, rows.length + 1, 3);
      const chart = tempSheet.newChart().asLineChart().addRange(dataRange).setNumHeaders(1)
        .setOption('title', ind.nama).setPosition(headerRow, 6, 0, 0).build();
      tempSheet.insertChart(chart);
      SpreadsheetApp.flush();

      const chartsOnSheet = tempSheet.getCharts();
      const embeddedChart = chartsOnSheet[chartsOnSheet.length - 1];

      const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slide.insertTextBox(ind.nama, 20, 15, 900, 40);
      slide.insertSheetsChart(embeddedChart, 20, 65, 520, 300);

      chartRow += rows.length + 3;
      Logger.log('Slide ke-' + (idx + 2) + ' (' + ind.nama + ') berhasil dibuat.');
    } catch (err) {
      Logger.log('GAGAL bikin slide untuk "' + ind.nama + '": ' + err.message);
    }
  });

  Logger.log('Jumlah slide akhir: ' + pres.getSlides().length);
  Logger.log('BUKA LINK INI BUAT CEK MANUAL: https://docs.google.com/presentation/d/' + pres.getId() + '/edit');
  Logger.log('Setelah dicek, hapus manual file ini dan sheet "' + tempSheet.getName() + '" ya.');
}

/**
 * Bikin sheet sementara yang sudah diformat rapi (judul, garis tabel,
 * rata tengah) untuk dipakai export PDF maupun Excel.
 */
function buildFormattedSheet(title, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = ss.insertSheet('TempExport_' + new Date().getTime());
  const numCols = headers.length;

  tempSheet.getRange(1, 1, 1, numCols).merge();
  tempSheet.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  tempSheet.getRange(2, 1, 1, numCols).setValues([headers])
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#f0f2f5');

  if (rows.length > 0) {
    tempSheet.getRange(3, 1, rows.length, numCols).setValues(rows)
      .setHorizontalAlignment('center');
  }

  const totalRows = rows.length + 2;
  tempSheet.getRange(2, 1, totalRows - 1, numCols)
    .setBorder(true, true, true, true, true, true);

  tempSheet.autoResizeColumns(1, numCols);
  SpreadsheetApp.flush();
  return tempSheet;
}

function exportSheetAs(tempSheet, format) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = tempSheet.getSheetId();
  let url;
  if (format === 'pdf') {
    url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
      '/export?format=pdf&gid=' + sheetId +
      '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&attachment=true&horizontal_alignment=CENTER';
  } else {
    url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
      '/export?format=xlsx&gid=' + sheetId;
  }
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  return response.getBlob();
}

function exportRecapPdf(title, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = buildFormattedSheet(title, headers, rows);
  const blob = exportSheetAs(tempSheet, 'pdf');
  ss.deleteSheet(tempSheet);
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: title.replace(/[^a-zA-Z0-9]/g, '_') + '.pdf'
  };
}

function exportRecapExcel(title, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempSheet = buildFormattedSheet(title, headers, rows);
  const blob = exportSheetAs(tempSheet, 'xlsx');
  ss.deleteSheet(tempSheet);
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: title.replace(/[^a-zA-Z0-9]/g, '_') + '.xlsx'
  };
}

