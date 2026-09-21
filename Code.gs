/**
 * ============================================================
 *  HELPDESK IT RSUD  -  Backend (Google Apps Script)   versi 2
 * ============================================================
 *
 *  PEMASANGAN PERTAMA:
 *  1. Buat Google Spreadsheet baru di akun Gmail khusus helpdesk.
 *  2. Extensions > Apps Script. Hapus isi Code.gs, tempel file ini.
 *  3. Jalankan  setup()  sekali, izinkan akses (Spreadsheet + Drive).
 *  4. Buat PIN petugas dengan setPinPetugas('nama', 'pin').
 *  5. Terapkan > Deployment baru > Aplikasi web
 *       Jalankan sebagai : Saya
 *       Akses            : Siapa saja
 *
 *  MEMPERBARUI DARI VERSI 1:
 *  1. Ganti seluruh isi Code.gs dengan file ini, simpan.
 *  2. Jalankan  setup()  lagi. Kolom baru ditambahkan otomatis,
 *     data lama tidak dihapus. Izinkan akses Drive saat diminta.
 *  3. Terapkan > Kelola deployment > pensil > Versi baru > Terapkan.
 * ============================================================
 */

// ------------------------------------------------------------
//  KONFIGURASI UMUM
// ------------------------------------------------------------

const ZONA = 'Asia/Jakarta';
const NAMA_INSTANSI = 'Instalasi Teknologi Informasi RSUD SSMA';

/** Kirim email saat ada tiket baru. Kuota Gmail biasa: 100 email/hari. */
const KIRIM_EMAIL = false;
const EMAIL_TIM_IT = 'ganti@gmail.com';

/** Batas ukuran satu lampiran setelah dikompres di HP. */
const MAKS_FILE_MB = 10;

const SESI_DETIK = 21600; // 6 jam

const SHEET = {
  tiket:    'Tiket',
  pegawai:  'Pegawai',
  petugas:  'Petugas',
  log:       'Log',
  terhapus:  'Terhapus',
  perangkat: 'Perangkat'
};

/** Urutan kolom untuk pemasangan baru. Pada spreadsheet lama,
 *  kolom yang belum ada ditambahkan di ujung kanan oleh setup(). */
const KOLOM_TIKET = ['ID', 'Dibuat', 'NIP', 'Nama', 'Unit', 'No HP',
                     'Grup', 'Kategori', 'Prioritas', 'Judul', 'Deskripsi',
                     'Detail', 'Lampiran', 'Status', 'Teknisi', 'Catatan',
                     'Diupdate', 'Selesai', 'Durasi Jam', 'Data'];

const KOLOM_PEGAWAI = ['NIP', 'Nama', 'Unit', 'No HP', 'Terdaftar'];
const KOLOM_PETUGAS = ['Username', 'Nama', 'Hash PIN', 'Peran', 'Aktif'];
const KOLOM_LOG     = ['Waktu', 'ID Tiket', 'Oleh', 'Aksi', 'Detail'];

/** Jejak tiket yang dihapus petugas. Sengaja TIDAK menyimpan isi detail
 *  dan lampiran, karena tiket sering dihapus justru karena isinya memuat
 *  data yang tidak semestinya (KTP orang lain, data pasien). */
const KOLOM_TERHAPUS = ['ID', 'Dihapus', 'Oleh', 'Alasan', 'Dibuat', 'NIP', 'Nama',
                        'Unit', 'Grup', 'Kategori', 'Judul', 'Status Terakhir',
                        'Lampiran Dibuang'];

/** HP/browser yang sudah mengizinkan notifikasi. Satu baris per perangkat. */
const KOLOM_PERANGKAT = ['Token', 'Peran', 'Pemilik', 'Nama', 'Terdaftar', 'Terakhir'];

/** Kolom berisi angka panjang yang harus disimpan sebagai teks.
 *  Tanpa ini Sheets mengubah NIP 18 digit jadi 1.98706E+17 dan
 *  angka 0 di depan nomor HP hilang. */
const KOLOM_TEKS = ['NIP', 'No HP'];

const PRIORITAS = ['Kritis', 'Tinggi', 'Normal', 'Rendah'];

const STATUS = ['Baru', 'Ditugaskan', 'Diproses', 'Menunggu Sparepart',
                'Diteruskan ke Bagian Umum', 'Selesai', 'Ditolak'];

const STATUS_SELESAI = ['Selesai', 'Ditolak'];

/** Unit / instalasi, sama dengan form lama. "Lainnya" membuka isian bebas. */
const UNIT = [
  'Rawat Jalan', 'Rawat Inap', 'Farmasi', 'Rekam Medis', 'Radiologi',
  'Laboratorium', 'Manajemen', 'Elektromedis', 'IPAL', 'K3RS',
  'Gudang Farmasi', 'Gizi', 'Laundry', 'Pemulasaran Jenazah', 'Hemodialisa',
  'Lainnya'
];

/** Unit pelayanan pasien langsung. Laporan gangguan dari unit ini
 *  otomatis minimal berprioritas Tinggi. */
const UNIT_KRITIS = ['Rawat Jalan', 'Rawat Inap', 'Farmasi', 'Radiologi',
                     'Laboratorium', 'Hemodialisa'];

const GOLONGAN_DARAH = ['A', 'B', 'AB', 'O'];

/** Isi dropdown "Pilih Permohonan", sama dengan form lama. */
const PILIHAN_FITUR_SIPINTER = [
  'Penambahan menu/fitur',
  'Perubahan menu/fitur'
];

const TAUTAN_CONTOH_WO =
  'https://docs.google.com/document/d/1zwUKWbD2_FTuStabOTnsEbU2hlPb0mowSXjqDqT35YA/edit?usp=sharing';

const PERNYATAAN_KTP =
  'Dengan mengunggah KTP melalui formulir ini, saya menyatakan bahwa saya telah ' +
  'memberikan data tersebut secara sadar dan sukarela. Saya menyetujui penggunaan ' +
  'data KTP saya untuk keperluan administrasi dan operasional sistem, termasuk ' +
  'proses verifikasi dan validasi identitas, serta keperluan lain yang berkaitan ' +
  'langsung dengan penyelenggaraan sistem.\n\n' +
  'Saya memahami bahwa data KTP saya akan digunakan hanya untuk kepentingan ' +
  'tersebut dan dikelola dengan memperhatikan keamanan serta kerahasiaan data ' +
  'pribadi sesuai dengan ketentuan yang berlaku.';

// ------------------------------------------------------------
//  JENIS TIKET
// ------------------------------------------------------------
//  tipe isian: teks, angka, paragraf, pilihan, file, setuju
//  terima (untuk file): 'gambar' atau 'gambar_pdf'
//  isiAwal: diisi otomatis dari profil ('nama', 'hp', 'nipAsn')
// ------------------------------------------------------------

const GRUP_GANGGUAN = 'Lapor Gangguan';
const GRUP_PERMOHONAN = 'Permohonan';

const POLA_HP = '^[0-9+][0-9 -]{8,16}$';

function gangguan_(kode, nama, ringkas, labelBarang, contoh, keterangan) {
  return {
    kode: kode, grup: GRUP_GANGGUAN, nama: nama, ringkas: ringkas,
    keterangan: keterangan || '',
    pakaiPrioritas: true,
    fields: [
      { id: 'barang', label: labelBarang, tipe: 'teks', wajib: true, maks: 120,
        bantuan: 'Contoh: ' + contoh },
      { id: 'penjelasan', label: 'Penjelasan masalah', tipe: 'paragraf', wajib: true,
        maks: 3000, bantuan: 'Sejak kapan? Pesan error apa yang muncul? Sudah dicoba apa?' },
      { id: 'foto', label: 'Foto error atau kerusakan', tipe: 'file', wajib: false,
        terima: 'gambar', bantuan: 'Opsional. Jangan memotret layar yang menampilkan data pasien.' }
    ]
  };
}

const JENIS = [
  gangguan_('hw', 'Komputer / Perangkat Keras',
    'PC, laptop, monitor, keyboard, mouse, UPS rusak atau error',
    'Nama barang yang rusak', 'PC pendaftaran loket 2',
    'Hanya untuk perangkat yang masih bisa diperbaiki. Jika mati total, ajukan lewat Permohonan > Perangkat Keras Baru.'),
  gangguan_('net', 'Jaringan / Internet',
    'Internet mati, WiFi tidak tersambung, kabel LAN',
    'Perangkat atau lokasi yang terdampak', 'Semua PC di nurse station lantai 2'),
  gangguan_('app', 'SIPINTER / Aplikasi Error',
    'SIPINTER atau aplikasi lain error, lambat, tidak bisa dibuka',
    'Nama aplikasi dan menu', 'SIPINTER, menu pendaftaran rawat jalan'),
  gangguan_('prn', 'Printer / Scanner',
    'Tidak bisa mencetak, kertas macet, hasil cetak rusak',
    'Nama printer atau scanner', 'Printer label farmasi'),
  gangguan_('tel', 'Telepon / CCTV',
    'Telepon extension mati, CCTV tidak tampil',
    'Nomor extension atau lokasi kamera', 'Ext. 112 ruang Hemodialisa'),
  gangguan_('akun', 'Akun Terkunci / Lupa Password',
    'Tidak bisa masuk SIPINTER, email, atau komputer',
    'Aplikasi atau akun yang bermasalah', 'Login SIPINTER'),

  {
    kode: 'fitur', grup: GRUP_PERMOHONAN, nama: 'Menu / Fitur SIPINTER',
    ringkas: 'Penambahan atau perubahan menu dan fitur SIPINTER',
    keterangan: 'Wajib melampirkan Work Order / surat permohonan yang sudah disetujui atasan langsung.',
    fields: [
      { id: 'permohonan', label: 'Pilih permohonan', tipe: 'pilihan', wajib: true,
        pilihan: PILIHAN_FITUR_SIPINTER },
      { id: 'wo', label: 'File Work Order / Surat Permohonan', tipe: 'file', wajib: true,
        terima: 'gambar_pdf',
        bantuan: 'Gunakan contoh WO atau versi dari unit kerja masing-masing. Wajib ada persetujuan atasan langsung. Boleh foto atau PDF.',
        tautan: { teks: 'Unduh contoh WO', url: TAUTAN_CONTOH_WO } },
      { id: 'uraian', label: 'Keterangan tambahan', tipe: 'paragraf', wajib: false,
        maks: 2000, bantuan: 'Opsional.' }
    ]
  },
  {
    kode: 'akunsip', grup: GRUP_PERMOHONAN, nama: 'Akun Pengguna SIPINTER',
    ringkas: 'Pembuatan akun baru untuk SIPINTER',
    fields: [
      { id: 'jabatan', label: 'Jabatan', tipe: 'teks', wajib: true, maks: 120 },
      { id: 'hp', label: 'Nomor handphone', tipe: 'teks', wajib: true, maks: 20,
        isiAwal: 'hp', inputmode: 'tel', pola: POLA_HP,
        pesanPola: 'Nomor handphone tidak valid.' },
      { id: 'ktp', label: 'Foto KTP', tipe: 'file', wajib: true, terima: 'gambar_pdf',
        bantuan: 'Foto KTP asli, terbaca jelas. Hanya dapat dilihat oleh tim IT.' },
      { id: 'setujuKtp', label: 'Pernyataan persetujuan penggunaan data KTP',
        tipe: 'setuju', wajib: true, pernyataan: PERNYATAAN_KTP,
        teksSetuju: 'Saya telah membaca, memahami, dan menyetujui penggunaan data KTP saya untuk kepentingan sistem sebagaimana dijelaskan di atas.' }
    ]
  },
  {
    kode: 'parkir', grup: GRUP_PERMOHONAN, nama: 'Kartu Parkir',
    ringkas: 'Pembuatan kartu parkir pegawai',
    keterangan: 'Silakan isi sesuai dengan status kepegawaian.',
    fields: [
      { id: 'namaGelar', label: 'Nama pegawai + gelar', tipe: 'teks', wajib: true,
        maks: 120, isiAwal: 'nama' },
      { id: 'nip', label: 'NIP', tipe: 'teks', wajib: true, maks: 18,
        isiAwal: 'nipAsn', inputmode: 'numeric', pola: '^(0|[0-9]{18})$',
        pesanPola: 'NIP harus 18 digit. Jika Non-ASN, isi dengan angka 0.',
        bantuan: 'Jika Non-ASN, isi dengan angka 0.' },
      { id: 'jabatan', label: 'Jabatan sesuai SK', tipe: 'teks', wajib: true, maks: 120 },
      { id: 'goldar', label: 'Golongan darah', tipe: 'pilihan', wajib: true,
        pilihan: GOLONGAN_DARAH },
      { id: 'foto', label: 'Foto pegawai', tipe: 'file', wajib: true, terima: 'gambar',
        maksSisi: 2000, bantuan: 'Jika ASN, gunakan foto di MyASN.' }
    ]
  },
  {
    kode: 'hwbaru', grup: GRUP_PERMOHONAN, nama: 'Perangkat Keras Baru',
    ringkas: 'Penambahan atau penggantian perangkat keras',
    keterangan: 'Permohonan ini menjadi data kajian kebutuhan perangkat keras rumah sakit dan akan diteruskan ke Bagian Umum.',
    fields: [
      { id: 'perangkat', label: 'Nama perangkat keras yang dibutuhkan', tipe: 'teks',
        wajib: true, maks: 120 },
      { id: 'jumlah', label: 'Jumlah', tipe: 'angka', wajib: true },
      { id: 'merk', label: 'Merk atau jenis', tipe: 'teks', wajib: false, maks: 120,
        bantuan: 'Opsional.' }
    ]
  },
  {
    kode: 'lain', grup: GRUP_PERMOHONAN, nama: 'Lainnya',
    ringkas: 'Kebutuhan selain pilihan di atas',
    keterangan: 'Jika permintaan bukan ranah IT, kami akan menjawab lewat WhatsApp.',
    fields: [
      { id: 'kebutuhan', label: 'Kebutuhan Bapak/Ibu', tipe: 'paragraf', wajib: true,
        maks: 3000, bantuan: 'Sampaikan permohonan dengan jelas dan rinci.' }
    ]
  }
];

/** Judul tiket dibuat otomatis dari isian. */
const JUDUL = {
  fitur:   function (d) { return 'Fitur SIPINTER: ' + d.permohonan; },
  akunsip: function (d, s) { return 'Akun SIPINTER untuk ' + s.nama; },
  parkir:  function (d) { return 'Kartu parkir ' + d.namaGelar; },
  hwbaru:  function (d) { return d.jumlah + ' x ' + d.perangkat; },
  lain:    function (d) { return d.kebutuhan.split('\n')[0].slice(0, 80); }
};

// ------------------------------------------------------------
//  SETUP
// ------------------------------------------------------------

/** Jalankan setiap kali memasang atau memperbarui. Aman diulang. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shTiket = siapkanSheet_(ss, SHEET.tiket, KOLOM_TIKET);
  const shPegawai = siapkanSheet_(ss, SHEET.pegawai, KOLOM_PEGAWAI);
  siapkanSheet_(ss, SHEET.petugas, KOLOM_PETUGAS);
  siapkanSheet_(ss, SHEET.log, KOLOM_LOG);
  siapkanTerhapus_(ss);
  siapkanPerangkat_(ss);

  const kategori = JENIS.map(function (j) { return j.nama; });
  validasi_(shTiket, 'Kategori', kategori);
  validasi_(shTiket, 'Prioritas', PRIORITAS);
  validasi_(shTiket, 'Status', STATUS);
  validasi_(shTiket, 'Grup', [GRUP_GANGGUAN, GRUP_PERMOHONAN]);

  [shTiket, shPegawai].forEach(function (sh) {
    KOLOM_TEKS.forEach(function (nama) {
      const k = nomorKolom_(sh, nama);
      if (k) sh.getRange(2, k, 5000, 1).setNumberFormat('@');
    });
  });

  const folder = folderLampiran_();

  const pesan = 'Setup selesai.\n\nFolder lampiran di Google Drive: "' +
    folder.getName() + '"\n\nLangkah berikutnya: Terapkan > Kelola deployment > Versi baru.';
  try { SpreadsheetApp.getUi().alert(pesan); } catch (e) { Logger.log(pesan); }
}

function siapkanSheet_(ss, nama, kolom) {
  let sh = ss.getSheetByName(nama);
  if (!sh) sh = ss.insertSheet(nama);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, kolom.length).setValues([kolom]);
  } else {
    const ada = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const kurang = kolom.filter(function (k) { return ada.indexOf(k) === -1; });
    if (kurang.length) {
      sh.getRange(1, ada.length + 1, 1, kurang.length).setValues([kurang]);
    }
  }
  const lebar = sh.getLastColumn();
  sh.getRange(1, 1, 1, lebar).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  return sh;
}

function siapkanTerhapus_(ss) {
  const sh = siapkanSheet_(ss, SHEET.terhapus, KOLOM_TERHAPUS);
  const k = nomorKolom_(sh, 'NIP');
  if (k) sh.getRange(2, k, 5000, 1).setNumberFormat('@');
  return sh;
}

function siapkanPerangkat_(ss) {
  const sh = siapkanSheet_(ss, SHEET.perangkat, KOLOM_PERANGKAT);
  const k = nomorKolom_(sh, 'Pemilik');
  if (k) sh.getRange(2, k, 5000, 1).setNumberFormat('@');
  return sh;
}

function nomorKolom_(sh, nama) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const i = head.indexOf(nama);
  return i === -1 ? 0 : i + 1;
}

function validasi_(sh, namaKolom, daftar) {
  const k = nomorKolom_(sh, namaKolom);
  if (!k) return;
  const aturan = SpreadsheetApp.newDataValidation()
    .requireValueInList(daftar, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, k, 5000, 1).setDataValidation(aturan);
}

function folderLampiran_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_LAMPIRAN');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* dibuat ulang */ }
  }
  const folder = DriveApp.createFolder('Lampiran Helpdesk IT');
  props.setProperty('FOLDER_LAMPIRAN', folder.getId());
  return folder;
}

/**
 * Membuat atau mengganti PIN petugas IT.
 * Contoh: setPinPetugas('budi', '482913')
 */
function setPinPetugas(username, pin) {
  if (!username || !pin) {
    throw new Error('Fungsi ini tidak dijalankan dari tombol Jalankan. ' +
      'Panggil dari fungsi lain, contoh: setPinPetugas("budi", "482913")');
  }
  const sh = sheet_(SHEET.petugas);
  const data = sh.getDataRange().getValues();
  const hash = hashPin_(pin);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(username).toLowerCase()) {
      sh.getRange(i + 1, 3).setValue(hash);
      Logger.log('PIN untuk "' + username + '" diperbarui.');
      return;
    }
  }
  sh.appendRow([username, username, hash, 'teknisi', 'YA']);
  Logger.log('Petugas "' + username + '" dibuat dengan PIN baru.');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Helpdesk IT')
    .addItem('Setup / perbarui', 'setup')
    .addItem('Rekap hari ini', 'rekapHariIni')
    .addSeparator()
    .addItem('Hapus semua data uji…', 'bersihkanDataUji')
    .addToUi();
}

/**
 * Mengosongkan sheet Tiket, Log, dan Pegawai, lalu memindahkan semua
 * lampiran ke Sampah. Sheet Petugas (PIN) tidak disentuh.
 *
 * Sengaja hanya bisa dijalankan dari menu Helpdesk IT di spreadsheet
 * dan harus dikonfirmasi dengan mengetik HAPUS. Salinan cadangan
 * spreadsheet selalu dibuat lebih dulu.
 */
function bersihkanDataUji() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    throw new Error('Fungsi ini hanya bisa dijalankan dari spreadsheet: ' +
      'menu Helpdesk IT > Hapus semua data uji. Tujuannya supaya data ' +
      'tidak terhapus tanpa sengaja dari editor.');
  }

  const jawab = ui.prompt(
    'Hapus semua data uji',
    'Ini akan MENGOSONGKAN sheet Tiket, Log, Pegawai, dan Terhapus, serta memindahkan ' +
    'semua lampiran (foto, KTP, WO) ke Sampah Google Drive.\n\n' +
    'Sheet Petugas dan PIN tim tidak disentuh. Salinan cadangan spreadsheet ' +
    'dibuat otomatis sebelum menghapus.\n\n' +
    'Ketik HAPUS untuk melanjutkan:',
    ui.ButtonSet.OK_CANCEL);

  if (jawab.getSelectedButton() !== ui.Button.OK ||
      String(jawab.getResponseText()).trim().toUpperCase() !== 'HAPUS') {
    ui.alert('Dibatalkan. Tidak ada data yang dihapus.');
    return;
  }

  const h = hapusDataUji_();
  ui.alert(
    'Selesai.\n\n' +
    'Tiket dihapus   : ' + h.tiket + '\n' +
    'Pegawai dihapus : ' + h.pegawai + '\n' +
    'Lampiran ke Sampah: ' + h.file + '\n\n' +
    'Cadangan tersimpan di Google Drive dengan nama:\n"' + h.cadangan + '"\n\n' +
    'Kalau lampiran berisi foto KTP asli, kosongkan juga Sampah di Google Drive.');
}

function hapusDataUji_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stempel = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH.mm');
  const cadangan = DriveApp.getFileById(ss.getId())
    .makeCopy('Cadangan Helpdesk IT ' + stempel);

  const jumlah = {};
  [SHEET.tiket, SHEET.log, SHEET.pegawai, SHEET.terhapus, SHEET.perangkat].forEach(function (nama) {
    const sh = ss.getSheetByName(nama);
    jumlah[nama] = 0;
    if (!sh) return;
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      // clearContent, bukan deleteRows, supaya format teks untuk NIP/No HP
      // dan dropdown validasi tetap ada untuk data berikutnya.
      sh.getRange(2, 1, n, sh.getLastColumn()).clearContent();
      jumlah[nama] = n;
    }
  });

  let file = 0;
  const it = folderLampiran_().getFiles();
  while (it.hasNext()) {
    it.next().setTrashed(true);
    file++;
  }

  return {
    tiket: jumlah[SHEET.tiket],
    pegawai: jumlah[SHEET.pegawai],
    file: file,
    cadangan: cadangan.getName()
  };
}

function rekapHariIni() {
  const s = hitungStatistik_();
  SpreadsheetApp.getUi().alert(
    'Tiket hari ini : ' + s.hariIni + '\n' +
    'Belum selesai  : ' + s.terbuka + '\n' +
    'Kritis terbuka : ' + s.kritis + '\n' +
    'Total tiket    : ' + s.total
  );
}

// ------------------------------------------------------------
//  ENTRY POINT WEB APP
// ------------------------------------------------------------

function doGet() {
  return json_({ ok: true, pesan: 'API Helpdesk IT aktif.', versi: 2 });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, pesan: 'Permintaan kosong.' });
    }
    const req = JSON.parse(e.postData.contents);
    const fn = ROUTER[req.aksi];
    if (!fn) return json_({ ok: false, pesan: 'Aksi tidak dikenal: ' + req.aksi });
    return json_(fn(req));
  } catch (err) {
    if (err && err.tolak) return json_({ ok: false, pesan: err.message });
    return json_({ ok: false, pesan: 'Kesalahan server: ' + err.message });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Error yang pesannya aman ditampilkan apa adanya ke pengguna. */
function tolak_(pesan) {
  const e = new Error(pesan);
  e.tolak = true;
  return e;
}

const ROUTER = {
  konfigurasi:   aksiKonfigurasi,
  loginPegawai:  aksiLoginPegawai,
  daftarPegawai: aksiDaftarPegawai,
  loginPetugas:  aksiLoginPetugas,
  buatTiket:     aksiBuatTiket,
  tiketSaya:     aksiTiketSaya,
  semuaTiket:    aksiSemuaTiket,
  detailTiket:   aksiDetailTiket,
  lampiran:      aksiLampiran,
  updateTiket:   aksiUpdateTiket,
  hapusTiket:    aksiHapusTiket,
  statistik:     aksiStatistik,
  daftarNotif:   aksiDaftarNotif,
  hapusNotif:    aksiHapusNotif
};

// ------------------------------------------------------------
//  AKSI: KONFIGURASI & LOGIN
// ------------------------------------------------------------

function aksiKonfigurasi() {
  return {
    ok: true,
    grup: [GRUP_GANGGUAN, GRUP_PERMOHONAN],
    jenis: JENIS,
    prioritas: PRIORITAS,
    status: STATUS,
    unit: UNIT,
    teknisi: daftarNamaPetugas_(),
    maksFileMB: MAKS_FILE_MB
  };
}

/** NIP 18 digit (ASN) atau NIK 16 digit (Non-ASN). */
function idValid_(id) {
  return /^[0-9]{16}$/.test(id) || /^[0-9]{18}$/.test(id);
}

function rapikanId_(v) {
  return String(v === null || v === undefined ? '' : v).replace(/[\s.'-]/g, '');
}

function aksiLoginPegawai(req) {
  const id = rapikanId_(req.nip);
  if (!id) return { ok: false, pesan: 'NIP atau NIK wajib diisi.' };

  const pegawai = cariPegawai_(id);
  if (!pegawai) {
    if (!idValid_(id)) {
      return { ok: false, pesan: 'NIP harus 18 digit, NIK harus 16 digit. Periksa kembali angkanya.' };
    }
    return { ok: false, kode: 'BELUM_TERDAFTAR', pesan: 'Belum terdaftar.' };
  }

  const sesi = {
    peran: 'pegawai',
    nip: rapikanId_(pegawai.NIP),
    nama: pegawai.Nama,
    unit: pegawai.Unit,
    hp: String(pegawai['No HP'] || '')
  };
  return { ok: true, token: buatSesi_(sesi), profil: sesi };
}

function aksiDaftarPegawai(req) {
  const id   = rapikanId_(req.nip);
  const nama = String(req.nama || '').trim().slice(0, 120);
  const unit = String(req.unit || '').trim().slice(0, 80);
  const hp   = String(req.hp   || '').trim().slice(0, 20);

  if (!idValid_(id)) return { ok: false, pesan: 'NIP harus 18 digit, NIK harus 16 digit.' };
  if (!nama) return { ok: false, pesan: 'Nama lengkap wajib diisi.' };
  if (!unit || unit === 'Lainnya') return { ok: false, pesan: 'Unit / instalasi wajib diisi.' };
  if (hp && !new RegExp(POLA_HP).test(hp)) return { ok: false, pesan: 'Nomor WhatsApp tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (cariPegawai_(id)) {
      return { ok: false, pesan: 'Nomor ini sudah terdaftar. Silakan langsung masuk.' };
    }
    tulisBaris_(sheet_(SHEET.pegawai), {
      'NIP': id, 'Nama': aman_(nama), 'Unit': aman_(unit),
      'No HP': hp, 'Terdaftar': new Date()
    });
  } finally {
    lock.releaseLock();
  }

  const sesi = { peran: 'pegawai', nip: id, nama: nama, unit: unit, hp: hp };
  return { ok: true, token: buatSesi_(sesi), profil: sesi };
}

function aksiLoginPetugas(req) {
  const username = String(req.username || '').trim().toLowerCase();
  const pin      = String(req.pin || '').trim();
  if (!username || !pin) return { ok: false, pesan: 'Username dan PIN wajib diisi.' };

  const p = bacaSheet_(SHEET.petugas).filter(function (r) {
    return String(r.Username).toLowerCase() === username &&
           String(r.Aktif).toUpperCase() === 'YA';
  })[0];

  if (!p || !p['Hash PIN'] || String(p['Hash PIN']) !== hashPin_(pin)) {
    return { ok: false, pesan: 'Username atau PIN salah.' };
  }

  const sesi = { peran: 'petugas', username: p.Username, nama: p.Nama, level: p.Peran };
  return { ok: true, token: buatSesi_(sesi), profil: sesi };
}

// ------------------------------------------------------------
//  AKSI: BUAT TIKET
// ------------------------------------------------------------

const MIME_GAMBAR = ['image/jpeg', 'image/png', 'image/webp'];
const EKSTENSI = { 'image/jpeg': '.jpg', 'image/png': '.png',
                   'image/webp': '.webp', 'application/pdf': '.pdf' };

function aksiBuatTiket(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();
  if (sesi.peran !== 'pegawai') return { ok: false, pesan: 'Tiket dibuat dari akun pegawai.' };

  const jenis = cariJenis_(req.jenis);
  if (!jenis) return { ok: false, pesan: 'Jenis tiket tidak dikenal.' };

  // ---- 1. Validasi semua isian sebelum menyimpan apa pun ----
  const masuk = req.data || {};
  const berkasMasuk = req.file || {};
  const isi = {};
  const baris = [];
  const unggah = [];
  const stempel = Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm');

  jenis.fields.forEach(function (f) {
    if (f.tipe === 'file') {
      const b = berkasMasuk[f.id];
      if (!b || !b.data) {
        if (f.wajib) throw tolak_(f.label + ' wajib dilampirkan.');
        return;
      }
      const boleh = f.terima === 'gambar_pdf'
        ? MIME_GAMBAR.concat(['application/pdf']) : MIME_GAMBAR;
      if (boleh.indexOf(b.mime) === -1) {
        throw tolak_(f.label + ': format file tidak didukung. Gunakan foto' +
          (f.terima === 'gambar_pdf' ? ' atau PDF.' : '.'));
      }
      const bytes = Utilities.base64Decode(b.data);
      if (bytes.length > MAKS_FILE_MB * 1024 * 1024) {
        throw tolak_(f.label + ' terlalu besar. Maksimal ' + MAKS_FILE_MB + ' MB.');
      }
      unggah.push({ field: f, mime: b.mime, bytes: bytes });
      return;
    }

    if (f.tipe === 'setuju') {
      if (masuk[f.id] !== true) {
        if (f.wajib) throw tolak_('Centang ' + f.label.toLowerCase() + ' untuk melanjutkan.');
        return;
      }
      isi[f.id] = 'Disetujui ' + stempel;
      baris.push(f.label + ': Disetujui ' + stempel);
      return;
    }

    let v = String(masuk[f.id] === null || masuk[f.id] === undefined ? '' : masuk[f.id]).trim();
    if (f.maks) v = v.slice(0, f.maks);
    if (!v) {
      if (f.wajib) throw tolak_(f.label + ' wajib diisi.');
      return;
    }
    if (f.tipe === 'pilihan' && f.pilihan.indexOf(v) === -1) {
      throw tolak_(f.label + ': pilihan tidak valid.');
    }
    if (f.tipe === 'angka') {
      const n = Number(v);
      if (!/^[0-9]+$/.test(v) || n < 1 || n > 9999) {
        throw tolak_(f.label + ' harus berupa angka 1 sampai 9999.');
      }
      v = String(n);
    }
    if (f.pola && !new RegExp(f.pola).test(v)) {
      throw tolak_(f.pesanPola || (f.label + ' tidak valid.'));
    }
    isi[f.id] = v;
    baris.push(f.label + ': ' + v);
  });

  // ---- 2. Prioritas ----
  let prioritas = 'Normal';
  if (jenis.pakaiPrioritas) {
    prioritas = PRIORITAS.indexOf(req.prioritas) !== -1 ? req.prioritas : 'Normal';
    if (UNIT_KRITIS.indexOf(sesi.unit) !== -1 &&
        ['Normal', 'Rendah'].indexOf(prioritas) !== -1) {
      prioritas = 'Tinggi';
    }
  }

  const judul = String(JUDUL[jenis.kode] ? JUDUL[jenis.kode](isi, sesi) : isi.barang).slice(0, 150);
  const deskripsi = jenis.grup === GRUP_GANGGUAN ? isi.penjelasan : baris.join('\n');

  // ---- 3. Simpan lampiran ke Drive (privat, hanya akun helpdesk) ----
  const berkasTersimpan = [];
  const idLampiran = {};
  if (unggah.length) {
    const folder = folderLampiran_();
    const sementara = Utilities.formatDate(new Date(), ZONA, 'yyyyMMdd-HHmmss');
    unggah.forEach(function (u) {
      const nama = 'baru-' + sementara + '-' + u.field.id + EKSTENSI[u.mime];
      const file = folder.createFile(Utilities.newBlob(u.bytes, u.mime, nama));
      berkasTersimpan.push({ file: file, field: u.field, mime: u.mime });
      idLampiran[u.field.id] = file.getId();
    });
  }

  const teksLampiran = berkasTersimpan.map(function (b) {
    return b.field.label + ': https://drive.google.com/file/d/' + b.file.getId() + '/view';
  }).join('\n');

  const dataJson = JSON.stringify({ jenis: jenis.kode, isi: isi, lampiran: idLampiran });

  // ---- 4. Tulis tiket ----
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let id;
  try {
    const sh = sheet_(SHEET.tiket);
    id = idBaru_(sh);
    tulisBaris_(sh, {
      'ID': id, 'Dibuat': new Date(), 'NIP': sesi.nip, 'Nama': aman_(sesi.nama),
      'Unit': aman_(sesi.unit), 'No HP': sesi.hp || '',
      'Grup': jenis.grup, 'Kategori': jenis.nama, 'Prioritas': prioritas,
      'Judul': aman_(judul), 'Deskripsi': aman_(deskripsi),
      'Detail': aman_(baris.join('\n')), 'Lampiran': teksLampiran,
      'Status': 'Baru', 'Diupdate': new Date(), 'Data': dataJson
    });
  } finally {
    lock.releaseLock();
  }

  berkasTersimpan.forEach(function (b) {
    try { b.file.setName(id + '-' + b.field.id + EKSTENSI[b.mime]); } catch (e) { /* nama saja */ }
  });

  catat_(id, sesi.nama, 'Tiket dibuat', jenis.nama + ' / ' + prioritas);
  if (KIRIM_EMAIL) kirimEmailTiketBaru_(id, sesi, jenis, prioritas, judul, baris);

  kirimNotif_({ peran: 'petugas' }, {
    judul: (jenis.grup === GRUP_GANGGUAN ? 'Gangguan baru' : 'Permohonan baru') +
           (jenis.pakaiPrioritas && prioritas !== 'Normal' ? ' · ' + prioritas : ''),
    isi: jenis.nama + ' · ' + sesi.unit + '\n' + judul,
    id: id
  });

  return { ok: true, id: id, pesan: 'Tiket ' + id + ' berhasil dikirim.' };
}

// ------------------------------------------------------------
//  AKSI: MEMBACA TIKET
// ------------------------------------------------------------

function aksiTiketSaya(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();

  const milik = bacaSheet_(SHEET.tiket)
    .filter(function (t) { return rapikanId_(t.NIP) === sesi.nip; })
    .map(function (t) { return bersihkanTiket_(t, false); })
    .reverse();

  return { ok: true, tiket: milik };
}

function aksiSemuaTiket(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();
  if (sesi.peran !== 'petugas') return { ok: false, pesan: 'Akses ditolak.' };

  let data = bacaSheet_(SHEET.tiket);

  if (req.status && req.status !== 'Semua') {
    if (req.status === 'Terbuka') {
      data = data.filter(function (t) { return STATUS_SELESAI.indexOf(t.Status) === -1; });
    } else {
      data = data.filter(function (t) { return t.Status === req.status; });
    }
  }
  if (req.grup && req.grup !== 'Semua') {
    data = data.filter(function (t) { return (t.Grup || GRUP_GANGGUAN) === req.grup; });
  }
  if (req.cari) {
    const q = String(req.cari).toLowerCase();
    data = data.filter(function (t) {
      return [t.ID, t.Judul, t.Nama, t.Unit, t.Kategori, t.Deskripsi, t.Detail]
        .join(' ').toLowerCase().indexOf(q) !== -1;
    });
  }

  const bobot = { 'Kritis': 0, 'Tinggi': 1, 'Normal': 2, 'Rendah': 3 };
  data.sort(function (a, b) {
    const pa = bobot[a.Prioritas] === undefined ? 9 : bobot[a.Prioritas];
    const pb = bobot[b.Prioritas] === undefined ? 9 : bobot[b.Prioritas];
    if (pa !== pb) return pa - pb;
    return new Date(a.Dibuat) - new Date(b.Dibuat);
  });

  return { ok: true, tiket: data.slice(0, 300).map(function (t) { return bersihkanTiket_(t, false); }) };
}

function aksiDetailTiket(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();

  const t = cariTiket_(req.id);
  if (!t) return { ok: false, pesan: 'Tiket tidak ditemukan.' };
  if (!bolehLihat_(sesi, t)) return { ok: false, pesan: 'Akses ditolak.' };

  const riwayat = bacaSheet_(SHEET.log)
    .filter(function (l) { return String(l['ID Tiket']) === String(req.id); })
    .map(function (l) {
      return { waktu: iso_(l.Waktu), oleh: l.Oleh, aksi: l.Aksi, detail: l.Detail };
    });

  return { ok: true, tiket: bersihkanTiket_(t, true), riwayat: riwayat };
}

/** Mengambil isi lampiran. Hanya file milik tiket itu yang boleh dibuka. */
function aksiLampiran(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();

  const t = cariTiket_(req.id);
  if (!t) return { ok: false, pesan: 'Tiket tidak ditemukan.' };
  if (!bolehLihat_(sesi, t)) return { ok: false, pesan: 'Akses ditolak.' };

  const data = bacaData_(t);
  const fileId = data.lampiran[req.field];
  if (!fileId) return { ok: false, pesan: 'Lampiran tidak ditemukan.' };

  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    return {
      ok: true,
      mime: blob.getContentType(),
      nama: blob.getName(),
      data: Utilities.base64Encode(blob.getBytes())
    };
  } catch (e) {
    return { ok: false, pesan: 'Lampiran sudah dihapus atau tidak dapat dibuka.' };
  }
}

// ------------------------------------------------------------
//  AKSI: PETUGAS
// ------------------------------------------------------------

function aksiUpdateTiket(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();
  if (sesi.peran !== 'petugas') return { ok: false, pesan: 'Akses ditolak.' };

  const sh = sheet_(SHEET.tiket);
  const data = sh.getDataRange().getValues();
  const kol = {};
  data[0].forEach(function (n, i) { kol[n] = i; });

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][kol['ID']]) !== String(req.id)) continue;

    const statusLama = data[i][kol['Status']];
    const perubahan = [];
    const set = function (nama, nilai) { sh.getRange(i + 1, kol[nama] + 1).setValue(nilai); };
    let statusBerubah = false;

    if (req.status && STATUS.indexOf(req.status) !== -1 && req.status !== statusLama) {
      statusBerubah = true;
      set('Status', req.status);
      perubahan.push(statusLama + ' -> ' + req.status);

      if (STATUS_SELESAI.indexOf(req.status) !== -1) {
        const selesai = new Date();
        const dibuat = new Date(data[i][kol['Dibuat']]);
        set('Selesai', selesai);
        set('Durasi Jam', Math.round((selesai - dibuat) / 36000) / 100);
      } else {
        set('Selesai', '');
        set('Durasi Jam', '');
      }
    }

    if (req.teknisi !== undefined && req.teknisi !== data[i][kol['Teknisi']]) {
      set('Teknisi', aman_(req.teknisi));
      perubahan.push('teknisi: ' + (req.teknisi || '-'));
    }

    if (req.prioritas && PRIORITAS.indexOf(req.prioritas) !== -1 &&
        req.prioritas !== data[i][kol['Prioritas']]) {
      set('Prioritas', req.prioritas);
      perubahan.push('prioritas: ' + req.prioritas);
    }

    const catatan = String(req.catatan || '').trim().slice(0, 2000);
    if (catatan) {
      const lama = String(data[i][kol['Catatan']] || '');
      const cap = Utilities.formatDate(new Date(), ZONA, 'dd/MM HH:mm');
      set('Catatan', aman_((lama ? lama + '\n' : '') + '[' + cap + ' ' + sesi.nama + '] ' + catatan));
      perubahan.push('catatan ditambahkan');
    }

    set('Diupdate', new Date());
    if (perubahan.length) catat_(req.id, sesi.nama, 'Update', perubahan.join('; '));

    // Kabari pegawai hanya untuk hal yang relevan baginya
    if (statusBerubah || catatan) {
      kirimNotif_({ peran: 'pegawai', pemilik: rapikanId_(data[i][kol['NIP']]) }, {
        judul: statusBerubah ? 'Tiket Anda: ' + req.status : 'Catatan baru dari tim IT',
        isi: String(data[i][kol['Judul']]) + (catatan ? '\n' + catatan : ''),
        id: req.id
      });
    }

    return { ok: true, pesan: 'Tiket ' + req.id + ' diperbarui.' };
  }

  return { ok: false, pesan: 'Tiket tidak ditemukan.' };
}

/**
 * Menghapus satu tiket: baris dihapus dari sheet Tiket, jejak singkatnya
 * dicatat di sheet Terhapus, dan lampirannya dibuang dari Drive.
 */
function aksiHapusTiket(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();
  if (sesi.peran !== 'petugas') return { ok: false, pesan: 'Akses ditolak.' };

  const alasan = String(req.alasan || '').trim().slice(0, 500);
  if (alasan.length < 5) return { ok: false, pesan: 'Tulis alasan penghapusan (minimal 5 huruf).' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let t;
  try {
    const sh = sheet_(SHEET.tiket);
    const data = sh.getDataRange().getValues();
    const head = data[0];
    const kId = head.indexOf('ID');
    let baris = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][kId]) === String(req.id)) { baris = i; break; }
    }
    if (baris === -1) return { ok: false, pesan: 'Tiket tidak ditemukan. Mungkin sudah dihapus petugas lain.' };

    t = {};
    head.forEach(function (h, j) { t[h] = data[baris][j]; });

    const shHapus = ss.getSheetByName(SHEET.terhapus) || siapkanTerhapus_(ss);
    const lampiran = bacaData_(t).lampiran;
    tulisBaris_(shHapus, {
      'ID': t.ID, 'Dihapus': new Date(), 'Oleh': aman_(sesi.nama), 'Alasan': aman_(alasan),
      'Dibuat': t.Dibuat, 'NIP': rapikanId_(t.NIP), 'Nama': aman_(t.Nama), 'Unit': aman_(t.Unit),
      'Grup': t.Grup || GRUP_GANGGUAN, 'Kategori': t.Kategori, 'Judul': aman_(t.Judul),
      'Status Terakhir': t.Status, 'Lampiran Dibuang': Object.keys(lampiran).length
    });

    sh.deleteRow(baris + 1);
  } finally {
    lock.releaseLock();
  }

  const lampiran = bacaData_(t).lampiran;
  Object.keys(lampiran).forEach(function (k) {
    try { DriveApp.getFileById(lampiran[k]).setTrashed(true); } catch (e) { /* sudah tidak ada */ }
  });

  catat_(t.ID, sesi.nama, 'Tiket dihapus', alasan);
  kirimNotif_({ peran: 'pegawai', pemilik: rapikanId_(t.NIP) }, {
    judul: 'Tiket Anda dihapus tim IT',
    isi: String(t.Judul) + '\nAlasan: ' + alasan,
    id: ''
  });
  return { ok: true, pesan: 'Tiket ' + t.ID + ' dihapus.' };
}

function aksiStatistik(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();
  if (sesi.peran !== 'petugas') return { ok: false, pesan: 'Akses ditolak.' };

  const s = hitungStatistik_();
  s.ok = true;
  return s;
}

// ------------------------------------------------------------
//  HELPER
// ------------------------------------------------------------

function sheet_(nama) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
  if (!sh) throw new Error('Sheet "' + nama + '" belum ada. Jalankan setup() dulu.');
  return sh;
}

function bacaSheet_(nama) {
  const sh = sheet_(nama);
  if (sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const head = data[0];
  return data.slice(1)
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      const o = {};
      head.forEach(function (h, i) { o[h] = r[i]; });
      return o;
    });
}

/** Menulis satu baris berdasarkan nama kolom, bukan urutan kolom. */
function tulisBaris_(sh, obj) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
}

/** Mencegah isian pengguna terbaca sebagai rumus spreadsheet. */
function aman_(v) {
  const s = String(v === null || v === undefined ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function cariJenis_(kode) {
  return JENIS.filter(function (j) { return j.kode === kode; })[0] || null;
}

function cariTiket_(id) {
  return bacaSheet_(SHEET.tiket).filter(function (r) {
    return String(r.ID) === String(id);
  })[0] || null;
}

function bolehLihat_(sesi, t) {
  if (sesi.peran === 'petugas') return true;
  return rapikanId_(t.NIP) === sesi.nip;
}

function bacaData_(t) {
  try {
    const d = JSON.parse(t.Data || '{}');
    return { jenis: d.jenis || '', isi: d.isi || {}, lampiran: d.lampiran || {} };
  } catch (e) {
    return { jenis: '', isi: {}, lampiran: {} };
  }
}

function bersihkanTiket_(t, lengkap) {
  const o = {
    id:        t.ID,
    dibuat:    iso_(t.Dibuat),
    nama:      t.Nama,
    unit:      t.Unit,
    grup:      t.Grup || GRUP_GANGGUAN,
    kategori:  t.Kategori,
    prioritas: t.Prioritas,
    judul:     t.Judul,
    status:    t.Status,
    teknisi:   t.Teknisi || '',
    diupdate:  iso_(t.Diupdate),
    selesai:   iso_(t.Selesai),
    durasi:    t['Durasi Jam'] || ''
  };
  if (lengkap) {
    const d = bacaData_(t);
    o.nip = rapikanId_(t.NIP);
    o.hp = String(t['No HP'] || '');
    o.deskripsi = t.Deskripsi;
    o.catatan = t.Catatan || '';
    o.jenis = d.jenis;
    o.isi = d.isi;
    o.lampiran = Object.keys(d.lampiran);
  }
  return o;
}

function iso_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, "yyyy-MM-dd'T'HH:mm:ss");
  return String(v);
}

/** Nomor tiket berikutnya untuk hari ini. Tiket yang sudah dihapus ikut
 *  dihitung, supaya nomornya tidak pernah dipakai ulang. */
function idBaru_(sh) {
  const tgl = Utilities.formatDate(new Date(), ZONA, 'yyyyMMdd');
  const awalan = 'TKT-' + tgl + '-';
  let n = 0;
  const periksa = function (s) {
    if (!s || s.getLastRow() < 2) return;
    s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().forEach(function (r) {
      const v = String(r[0]);
      if (v.indexOf(awalan) === 0) {
        const urut = parseInt(v.substring(awalan.length), 10);
        if (urut > n) n = urut;
      }
    });
  };
  periksa(sh);
  periksa(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.terhapus));
  const urut = String(n + 1);
  return awalan + (urut.length >= 3 ? urut : ('000' + urut).slice(-3));
}

function cariPegawai_(id) {
  return bacaSheet_(SHEET.pegawai).filter(function (p) {
    return rapikanId_(p.NIP) === id;
  })[0] || null;
}

function daftarNamaPetugas_() {
  return bacaSheet_(SHEET.petugas)
    .filter(function (p) { return String(p.Aktif).toUpperCase() === 'YA'; })
    .map(function (p) { return p.Nama; });
}

function catat_(idTiket, oleh, aksi, detail) {
  try {
    sheet_(SHEET.log).appendRow([new Date(), idTiket, aman_(oleh), aksi, aman_(detail || '')]);
  } catch (e) {
    // kegagalan log tidak boleh menggagalkan transaksi utama
  }
}

function hitungStatistik_() {
  const semua = bacaSheet_(SHEET.tiket);
  const tglKini = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd');

  let hariIni = 0, terbuka = 0, kritis = 0, selesai = 0, totalJam = 0, nSelesai = 0;

  semua.forEach(function (t) {
    if (t.Dibuat instanceof Date &&
        Utilities.formatDate(t.Dibuat, ZONA, 'yyyy-MM-dd') === tglKini) hariIni++;

    if (STATUS_SELESAI.indexOf(t.Status) === -1) {
      terbuka++;
      if (t.Prioritas === 'Kritis') kritis++;
    } else {
      selesai++;
      const d = parseFloat(t['Durasi Jam']);
      if (!isNaN(d)) { totalJam += d; nSelesai++; }
    }
  });

  return {
    total: semua.length, hariIni: hariIni, terbuka: terbuka, kritis: kritis,
    selesai: selesai,
    rataJam: nSelesai ? Math.round(totalJam / nSelesai * 10) / 10 : 0
  };
}

// ------------------------------------------------------------
//  SESI
// ------------------------------------------------------------

function buatSesi_(data) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('sesi_' + token, JSON.stringify(data), SESI_DETIK);
  return token;
}

function bacaSesi_(token) {
  if (!token) return null;
  const c = CacheService.getScriptCache();
  const v = c.get('sesi_' + token);
  if (!v) return null;
  c.put('sesi_' + token, v, SESI_DETIK);
  return JSON.parse(v);
}

function sesiHabis_() {
  return { ok: false, kode: 'SESI_HABIS', pesan: 'Sesi berakhir. Silakan masuk lagi.' };
}

function hashPin_(pin) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(pin), Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    const h = (b < 0 ? b + 256 : b).toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
}

// ------------------------------------------------------------
//  NOTIFIKASI (Firebase Cloud Messaging)
// ------------------------------------------------------------
//  Aktif hanya kalau Script Property FCM_SERVICE_ACCOUNT sudah diisi
//  dengan isi file JSON akun layanan dari Firebase. Kalau belum diisi,
//  semua fungsi di bawah diam saja dan aplikasi tetap berjalan normal.
//
//  JANGAN menaruh isi akun layanan di Code.gs, index.html, atau GitHub.
// ------------------------------------------------------------

function aksiDaftarNotif(req) {
  const sesi = bacaSesi_(req.token);
  if (!sesi) return sesiHabis_();
  const fcm = String(req.fcm || '').trim();
  if (fcm.length < 20 || fcm.length > 4096 || /\s/.test(fcm)) {
    return { ok: false, pesan: 'Token notifikasi tidak valid.' };
  }
  const pemilik = sesi.peran === 'petugas' ? String(sesi.username) : String(sesi.nip);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET.perangkat) || siapkanPerangkat_(ss);
    const data = sh.getDataRange().getValues();
    const head = data[0];
    const kol = {};
    head.forEach(function (h, i) { kol[h] = i; });

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][kol.Token]) === fcm) {
        // HP yang sama dipakai orang lain: pindahkan ke pemilik baru
        sh.getRange(i + 1, kol.Peran + 1).setValue(sesi.peran);
        sh.getRange(i + 1, kol.Pemilik + 1).setValue(pemilik);
        sh.getRange(i + 1, kol.Nama + 1).setValue(aman_(sesi.nama));
        sh.getRange(i + 1, kol.Terakhir + 1).setValue(new Date());
        return { ok: true };
      }
    }
    tulisBaris_(sh, {
      'Token': fcm, 'Peran': sesi.peran, 'Pemilik': pemilik, 'Nama': aman_(sesi.nama),
      'Terdaftar': new Date(), 'Terakhir': new Date()
    });
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/** Dipanggil saat keluar, supaya HP bersama tidak menerima notifikasi
 *  milik orang yang sudah keluar. Cukup dengan token, tanpa sesi. */
function aksiHapusNotif(req) {
  const fcm = String(req.fcm || '').trim();
  if (fcm) hapusToken_([fcm]);
  return { ok: true };
}

function hapusToken_(daftar) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.perangkat);
  if (!sh || sh.getLastRow() < 2) return;
  const data = sh.getDataRange().getValues();
  const k = data[0].indexOf('Token');
  for (let i = data.length - 1; i >= 1; i--) {
    if (daftar.indexOf(String(data[i][k])) !== -1) sh.deleteRow(i + 1);
  }
}

function akunLayanan_() {
  const teks = PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT');
  if (!teks) return null;
  let sa;
  try { sa = JSON.parse(teks); } catch (e) {
    throw new Error('FCM_SERVICE_ACCOUNT bukan JSON yang valid. Tempel ulang seluruh isi file akun layanan.');
  }
  if (!sa.client_email || !sa.private_key || !sa.project_id) {
    throw new Error('FCM_SERVICE_ACCOUNT tidak lengkap (butuh client_email, private_key, project_id).');
  }
  return sa;
}

function b64url_(s) {
  return Utilities.base64EncodeWebSafe(s).replace(/=+$/, '');
}

/** Token akses Google untuk mengirim lewat FCM, disimpan 50 menit. */
function tokenAkses_(sa) {
  const cache = CacheService.getScriptCache();
  const ada = cache.get('fcm_akses');
  if (ada) return ada;

  const kini = Math.floor(Date.now() / 1000);
  const kepala = b64url_(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const klaim = b64url_(JSON.stringify({
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    iat: kini, exp: kini + 3600
  }));
  const tanda = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(kepala + '.' + klaim, sa.private_key)).replace(/=+$/, '');

  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: kepala + '.' + klaim + '.' + tanda
    },
    muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) {
    throw new Error('Firebase menolak akun layanan: ' + r.getContentText().slice(0, 300));
  }
  const akses = JSON.parse(r.getContentText()).access_token;
  cache.put('fcm_akses', akses, 3000);
  return akses;
}

/**
 * Mengirim notifikasi ke sekelompok perangkat.
 *   sasaran: { peran: 'petugas' }  atau  { peran: 'pegawai', pemilik: NIP }
 *   pesan:   { judul, isi, id }
 * Tidak pernah menggagalkan transaksi utama; kegagalan dicatat di Log.
 */
function kirimNotif_(sasaran, pesan) {
  const hasil = { terkirim: 0, gagal: 0, perangkat: 0, pesan: '' };
  try {
    const sa = akunLayanan_();
    if (!sa) { hasil.pesan = 'Notifikasi belum diaktifkan (FCM_SERVICE_ACCOUNT kosong).'; return hasil; }

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.perangkat);
    if (!sh || sh.getLastRow() < 2) { hasil.pesan = 'Belum ada perangkat terdaftar.'; return hasil; }

    let baris = bacaSheet_(SHEET.perangkat).filter(function (r) { return r.Peran === sasaran.peran; });
    if (sasaran.peran === 'petugas') {
      const aktif = bacaSheet_(SHEET.petugas)
        .filter(function (p) { return String(p.Aktif).toUpperCase() === 'YA'; })
        .map(function (p) { return String(p.Username).toLowerCase(); });
      baris = baris.filter(function (r) { return aktif.indexOf(String(r.Pemilik).toLowerCase()) !== -1; });
    } else {
      baris = baris.filter(function (r) { return rapikanId_(r.Pemilik) === String(sasaran.pemilik); });
    }
    hasil.perangkat = baris.length;
    if (!baris.length) { hasil.pesan = 'Tidak ada perangkat tujuan.'; return hasil; }

    const akses = tokenAkses_(sa);
    const url = 'https://fcm.googleapis.com/v1/projects/' + sa.project_id + '/messages:send';
    const permintaan = baris.map(function (r) {
      return {
        url: url, method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + akses },
        payload: JSON.stringify({
          message: {
            token: String(r.Token),
            data: {
              judul: String(pesan.judul).slice(0, 120),
              isi: String(pesan.isi || '').slice(0, 300),
              id: String(pesan.id || '')
            },
            webpush: { headers: { Urgency: 'high', TTL: '86400' } }
          }
        })
      };
    });

    const mati = [];
    UrlFetchApp.fetchAll(permintaan).forEach(function (res, i) {
      const kode = res.getResponseCode();
      if (kode === 200) { hasil.terkirim++; return; }
      hasil.gagal++;
      const teks = res.getContentText();
      if (kode === 404 || /UNREGISTERED|registration token/i.test(teks)) {
        mati.push(String(baris[i].Token));   // aplikasi dihapus / izin dicabut
      } else if (!hasil.pesan) {
        hasil.pesan = 'FCM ' + kode + ': ' + teks.slice(0, 200);
      }
    });
    if (mati.length) hapusToken_(mati);
    if (hasil.pesan) catat_('-', 'sistem', 'Notifikasi gagal', hasil.pesan);
  } catch (e) {
    hasil.pesan = e.message;
    catat_('-', 'sistem', 'Notifikasi gagal', e.message);
  }
  return hasil;
}

/**
 * Jalankan dari editor untuk memastikan notifikasi sudah berfungsi.
 * Mengirim pesan uji ke semua HP petugas yang sudah mengaktifkan notifikasi.
 */
function ujiNotifikasi() {
  const h = kirimNotif_({ peran: 'petugas' }, {
    judul: 'Uji notifikasi Helpdesk IT',
    isi: 'Kalau pesan ini muncul, notifikasi sudah berfungsi.',
    id: ''
  });
  const laporan = 'Perangkat petugas: ' + h.perangkat + ', terkirim: ' + h.terkirim +
                  ', gagal: ' + h.gagal + (h.pesan ? '\n' + h.pesan : '');
  Logger.log(laporan);
  if (!h.terkirim) throw new Error(laporan);
}

// ------------------------------------------------------------
//  EMAIL (opsional)
// ------------------------------------------------------------

function kirimEmailTiketBaru_(id, sesi, jenis, prioritas, judul, baris) {
  try {
    MailApp.sendEmail({
      to: EMAIL_TIM_IT,
      subject: '[' + prioritas + '] ' + jenis.nama + ' - ' + id,
      body: 'Tiket baru masuk di ' + NAMA_INSTANSI + '\n\n' +
            'ID        : ' + id + '\n' +
            'Pelapor   : ' + sesi.nama + '\n' +
            'Unit      : ' + sesi.unit + '\n' +
            'No HP     : ' + (sesi.hp || '-') + '\n' +
            'Jenis     : ' + jenis.grup + ' / ' + jenis.nama + '\n' +
            'Prioritas : ' + prioritas + '\n\n' +
            judul + '\n\n' + baris.join('\n') + '\n'
    });
  } catch (e) {
    catat_(id, 'sistem', 'Email gagal', e.message);
  }
}
