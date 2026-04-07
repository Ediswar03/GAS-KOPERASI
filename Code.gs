// =====================================================
// KOPERASI KONSUMEN THEOLI BINA MANDIRI
// Backend Google Apps Script - Integrasi Google Sheets
// =====================================================
// CARA SETUP:
// 1. Buka Google Sheets baru, lalu buka Tools > Apps Script
// 2. Salin semua kode ini ke editor Apps Script
// 3. Ganti SPREADSHEET_ID dengan ID Google Sheets Anda (dari URL)
// 4. Jalankan fungsi setupDatabase() SATU KALI untuk inisialisasi
// 5. Deploy sebagai Web App: Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Salin URL Web App dan update di Index.html (jika diperlukan)
// =====================================================

const SPREADSHEET_ID = '1tx_seOxA3AUcmFZVMwL2dyLO1DgZvF7cRbYvddfRPDs'; // Ganti dengan ID Spreadsheet Anda

// ============================================
// 1. INISIALISASI WEB APP (Entry Point)
// ============================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Kop-Tech: Koperasi System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================
// 2. HELPER: Ambil Sheet berdasarkan nama
// ============================================
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Auto-buat sheet jika belum ada
    sheet = ss.insertSheet(sheetName);
    Logger.log('Sheet baru dibuat: ' + sheetName);
  }
  return sheet;
}

// ============================================
// 3. API: LOGIN & AUTENTIKASI ROLE
// ============================================
function prosesLogin(username, password) {
  try {
    const sheet = getSheet('Users');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username && data[i][2] === password) {
        if (data[i][4] !== 'Aktif') {
          return { success: false, message: 'Akun belum diverifikasi Admin.' };
        }
        return {
          success: true,
          user: { id: data[i][0], username: data[i][1], role: data[i][3] }
        };
      }
    }
    return { success: false, message: 'Username atau Password salah.' };
  } catch (e) {
    Logger.log('Error prosesLogin: ' + e.message);
    return { success: false, message: 'Terjadi kesalahan sistem: ' + e.message };
  }
}

// ============================================
// 4. API: ANGGOTA - Ambil Saldo Simpanan
// ============================================
function getSaldoSimpanan(userId) {
  try {
    const sheet = getSheet('Simpanan');
    const data = sheet.getDataRange().getValues();
    let total = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == userId) {
        total += parseFloat(data[i][4]) || 0;
      }
    }
    return total;
  } catch (e) {
    Logger.log('Error getSaldoSimpanan: ' + e.message);
    return 0;
  }
}

// ============================================
// 5. API: ANGGOTA - Ajukan Pinjaman
// ============================================
function ajukanPinjaman(userId, jumlah, tenor) {
  try {
    const sheet = getSheet('Pinjaman');
    const idPinjaman = 'PJM-' + new Date().getTime();
    const tgl = new Date().toLocaleDateString('id-ID');
    
    sheet.appendRow([idPinjaman, userId, tgl, jumlah, tenor, 'Pending', 'Pending']);
    return { success: true, message: 'Pengajuan pinjaman berhasil dikirim untuk disetujui Admin.' };
  } catch (e) {
    Logger.log('Error ajukanPinjaman: ' + e.message);
    return { success: false, message: 'Gagal mengajukan pinjaman: ' + e.message };
  }
}

// ============================================
// 6. API: ADMIN/KETUA - Ambil Data Pinjaman Pending
// ============================================
function getPendingPinjaman() {
  try {
    const sheet = getSheet('Pinjaman');
    const data = sheet.getDataRange().getValues();
    let pending = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] === 'Pending' || data[i][6] === 'Pending') {
        pending.push({
          id: data[i][0],
          userId: data[i][1],
          tanggal: data[i][2],
          jumlah: data[i][3],
          tenor: data[i][4],
          statusAdmin: data[i][5],
          statusKetua: data[i][6]
        });
      }
    }
    return pending;
  } catch (e) {
    Logger.log('Error getPendingPinjaman: ' + e.message);
    return [];
  }
}

// ============================================
// 7. API: Ambil Semua Daftar Pinjaman dari Google Sheets
//    Sheet: DaftarPinjaman
//    Header: ID | Nama | Alamat | Besar | Tipe | Tanggal | SistemTenor | Tenor | Saldo
// ============================================
function getDaftarPinjaman() {
  try {
    const sheet = getSheet('DaftarPinjaman');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return []; // Hanya header, belum ada data
    
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Skip baris kosong
      result.push({
        id: data[i][0],
        nama: data[i][1] || '',
        alamat: data[i][2] || '',
        besar: parseFloat(data[i][3]) || 0,
        tipe: data[i][4] || 'Harian',
        tanggal: data[i][5] ? data[i][5].toString() : '',
        sistemTenor: data[i][6] || 'Harian',
        tenor: parseInt(data[i][7]) || 1,
        saldo: parseFloat(data[i][8]) || 0
      });
    }
    return result;
  } catch (e) {
    Logger.log('Error getDaftarPinjaman: ' + e.message);
    return [];
  }
}

// ============================================
// 8. API: Simpan Semua Data Pinjaman Sekaligus (Bulk Sync)
//    Dipanggil setiap kali frontend melakukan perubahan besar
// ============================================
function simpanDaftarPinjamanAll(dataArray) {
  try {
    const sheet = getSheet('DaftarPinjaman');
    
    // Pastikan header ada
    if (sheet.getLastRow() === 0) {
      const headers = ['ID', 'Nama', 'Alamat', 'Besar', 'Tipe', 'Tanggal', 'SistemTenor', 'Tenor', 'Saldo'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    }
    
    // Hapus data lama (kecuali header)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 9).clearContent();
    }
    
    // Tulis data baru
    if (dataArray && dataArray.length > 0) {
      const rows = dataArray.map(p => [
        p.id,
        p.nama || '',
        p.alamat || '',
        p.besar || 0,
        p.tipe || 'Harian',
        p.tanggal || '',
        p.sistemTenor || 'Harian',
        p.tenor || 1,
        p.saldo || 0
      ]);
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    return { success: true, message: 'Semua data pinjaman berhasil disinkronkan ke Google Sheets. Total: ' + (dataArray ? dataArray.length : 0) + ' data.' };
  } catch (e) {
    Logger.log('Error simpanDaftarPinjamanAll: ' + e.message);
    return { success: false, message: 'Gagal sinkronisasi: ' + e.message };
  }
}

// ============================================
// 9. API: Simpan Satu Data Pinjaman Baru
// ============================================
function simpanPinjaman(pinjamanObj) {
  try {
    const sheet = getSheet('DaftarPinjaman');
    
    // Pastikan header ada
    if (sheet.getLastRow() === 0) {
      const headers = ['ID', 'Nama', 'Alamat', 'Besar', 'Tipe', 'Tanggal', 'SistemTenor', 'Tenor', 'Saldo'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    }
    
    sheet.appendRow([
      pinjamanObj.id,
      pinjamanObj.nama || '',
      pinjamanObj.alamat || '',
      pinjamanObj.besar || 0,
      pinjamanObj.tipe || 'Harian',
      pinjamanObj.tanggal || '',
      pinjamanObj.sistemTenor || 'Harian',
      pinjamanObj.tenor || 1,
      pinjamanObj.saldo !== undefined ? pinjamanObj.saldo : ((pinjamanObj.besar * 1.2) || 0)
    ]);
    
    return { success: true, message: 'Data pinjaman berhasil disimpan ke Google Sheets.' };
  } catch (e) {
    Logger.log('Error simpanPinjaman: ' + e.message);
    return { success: false, message: 'Gagal menyimpan data: ' + e.message };
  }
}

// ============================================
// 10. API: Update Data Pinjaman yang Sudah Ada
// ============================================
function updatePinjamanBackend(pinjamanObj) {
  try {
    const sheet = getSheet('DaftarPinjaman');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == pinjamanObj.id) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          pinjamanObj.id,
          pinjamanObj.nama || '',
          pinjamanObj.alamat || '',
          pinjamanObj.besar || 0,
          pinjamanObj.tipe || 'Harian',
          pinjamanObj.tanggal || '',
          pinjamanObj.sistemTenor || 'Harian',
          pinjamanObj.tenor || 1,
          pinjamanObj.saldo !== undefined ? pinjamanObj.saldo : ((pinjamanObj.besar * 1.2) || 0)
        ]]);
        return { success: true, message: 'Data pinjaman berhasil diperbarui di Google Sheets.' };
      }
    }
    return { success: false, message: 'Data pinjaman dengan ID tersebut tidak ditemukan.' };
  } catch (e) {
    Logger.log('Error updatePinjamanBackend: ' + e.message);
    return { success: false, message: 'Gagal update data: ' + e.message };
  }
}

// ============================================
// 11. API: Hapus Data Pinjaman Berdasarkan ID
// ============================================
function hapusPinjamanBackend(id) {
  try {
    const sheet = getSheet('DaftarPinjaman');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1); // +1 karena getValues 0-indexed, deleteRow 1-indexed
        return { success: true, message: 'Data pinjaman berhasil dihapus dari Google Sheets.' };
      }
    }
    return { success: false, message: 'Data pinjaman dengan ID tersebut tidak ditemukan.' };
  } catch (e) {
    Logger.log('Error hapusPinjamanBackend: ' + e.message);
    return { success: false, message: 'Gagal menghapus data: ' + e.message };
  }
}

// ============================================
// 12. API: Bayar Angsuran - Kurangi Saldo & Catat Log
//     Sheet: Angsuran (log pembayaran)
//     Kolom Saldo di DaftarPinjaman akan berkurang otomatis
// ============================================
function bayarAngsuranBackend(pinjamanId, nominal) {
  try {
    // 1. Cari & update saldo di sheet DaftarPinjaman
    const sheetPinjaman = getSheet('DaftarPinjaman');
    const data = sheetPinjaman.getDataRange().getValues();
    let found = false;
    let sisaSaldo = 0;
    let namaNasabah = '';
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == pinjamanId) {
        const currentSaldo = parseFloat(data[i][8]) || 0;
        
        if (nominal > currentSaldo) {
          return { 
            success: false, 
            message: 'Nominal pembayaran (Rp ' + nominal.toLocaleString('id-ID') + ') melebihi sisa saldo pinjaman (Rp ' + currentSaldo.toLocaleString('id-ID') + ')!' 
          };
        }
        
        sisaSaldo = Math.max(0, currentSaldo - nominal);
        sheetPinjaman.getRange(i + 1, 9).setValue(sisaSaldo); // Kolom ke-9 = Saldo
        namaNasabah = data[i][1];
        found = true;
        break;
      }
    }
    
    if (!found) {
      return { success: false, message: 'Data pinjaman dengan ID ' + pinjamanId + ' tidak ditemukan.' };
    }
    
    // 2. Catat log pembayaran ke sheet Angsuran
    const sheetAngsuran = getSheet('Angsuran');
    
    // Pastikan header ada
    if (sheetAngsuran.getLastRow() === 0) {
      const headers = ['ID_Angsuran', 'ID_Pinjaman', 'Nama_Nasabah', 'Tanggal_Bayar', 'Jumlah_Bayar', 'Sisa_Saldo'];
      sheetAngsuran.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheetAngsuran.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    }
    
    const idAngsuran = 'ANG-' + new Date().getTime();
    const tglBayar = new Date().toLocaleDateString('id-ID');
    sheetAngsuran.appendRow([idAngsuran, pinjamanId, namaNasabah, tglBayar, nominal, sisaSaldo]);
    
    return { 
      success: true, 
      message: 'Pembayaran angsuran sebesar Rp ' + nominal.toLocaleString('id-ID') + ' berhasil dicatat.',
      sisaSaldo: sisaSaldo,
      lunas: sisaSaldo <= 0
    };
  } catch (e) {
    Logger.log('Error bayarAngsuranBackend: ' + e.message);
    return { success: false, message: 'Gagal mencatat angsuran: ' + e.message };
  }
}

// ============================================
// 13. API: Ambil Riwayat Angsuran Berdasarkan ID Pinjaman
// ============================================
function getRiwayatAngsuran(pinjamanId) {
  try {
    const sheet = getSheet('Angsuran');
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == pinjamanId) {
        result.push({
          idAngsuran: data[i][0],
          idPinjaman: data[i][1],
          namaNasabah: data[i][2] || '',
          tanggalBayar: data[i][3] ? data[i][3].toString() : '',
          jumlahBayar: parseFloat(data[i][4]) || 0,
          sisaSaldo: parseFloat(data[i][5]) || 0
        });
      }
    }
    return result;
  } catch (e) {
    Logger.log('Error getRiwayatAngsuran: ' + e.message);
    return [];
  }
}

// ============================================
// 14. API: Ambil Statistik Dashboard untuk Chart
//     Mengembalikan data untuk ditampilkan di grafik dashboard
// ============================================
function getStatistikDashboard() {
  try {
    const sheetPinjaman = getSheet('DaftarPinjaman');
    const data = sheetPinjaman.getDataRange().getValues();
    
    let totalPinjaman = 0;
    let totalSaldo = 0;
    let countHarian = 0;
    let countMingguan = 0;
    let countTempo = 0;
    let countLunas = 0;
    let countAktif = 0;
    const bulanData = {};
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      const besar = parseFloat(data[i][3]) || 0;
      const saldo = parseFloat(data[i][8]) || 0;
      const tipe = data[i][4] || '';
      const tanggal = data[i][5] ? data[i][5].toString() : '';
      
      totalPinjaman += besar;
      totalSaldo += saldo;
      
      // Hitung per tipe
      if (tipe === 'Harian') countHarian++;
      else if (tipe === 'Mingguan') countMingguan++;
      else if (tipe === 'Tempo') countTempo++;
      
      // Status
      if (saldo <= 0) countLunas++;
      else countAktif++;
      
      // Data per bulan (dari tanggal pencairan)
      if (tanggal) {
        const parts = tanggal.split('/');
        if (parts.length >= 2) {
          const bulanKey = parts[1] + '/' + (parts[2] ? parts[2].substring(2) : '');
          bulanData[bulanKey] = (bulanData[bulanKey] || 0) + besar;
        }
      }
    }
    
    return {
      success: true,
      totalPinjaman: totalPinjaman,
      totalSaldo: totalSaldo,
      distribusiTipe: { Harian: countHarian, Mingguan: countMingguan, Tempo: countTempo },
      distribusiStatus: { Aktif: countAktif, Lunas: countLunas },
      dataBulan: bulanData,
      totalNasabah: data.length - 1
    };
  } catch (e) {
    Logger.log('Error getStatistikDashboard: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ============================================
// 15. PLACEHOLDER: Tambah Record Simpanan (untuk fitur masa depan)
// ============================================
function tambahSimpanan(userId, jenis, jumlah) {
  try {
    const sheet = getSheet('Simpanan');
    
    // Pastikan header ada
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 5).setValues([['ID_Transaksi', 'ID_User', 'Tanggal', 'Jenis', 'Jumlah']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f3f3');
    }
    
    const idTx = 'SMP-' + new Date().getTime();
    const tgl = new Date().toLocaleDateString('id-ID');
    sheet.appendRow([idTx, userId, tgl, jenis, jumlah]);
    
    return { success: true, message: 'Simpanan berhasil dicatat.' };
  } catch (e) {
    Logger.log('Error tambahSimpanan: ' + e.message);
    return { success: false, message: 'Gagal mencatat simpanan: ' + e.message };
  }
}

// ============================================
// 16. PLACEHOLDER: Persetujuan Pinjaman oleh Admin
// ============================================
function setujuiPinjaman(idPinjaman, role) {
  try {
    const sheet = getSheet('Pinjaman');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == idPinjaman) {
        if (role === 'Admin') {
          sheet.getRange(i + 1, 6).setValue('Disetujui');
        } else if (role === 'Ketua') {
          sheet.getRange(i + 1, 7).setValue('Disetujui');
        }
        return { success: true, message: 'Pengajuan pinjaman berhasil disetujui oleh ' + role + '.' };
      }
    }
    return { success: false, message: 'Data pengajuan tidak ditemukan.' };
  } catch (e) {
    Logger.log('Error setujuiPinjaman: ' + e.message);
    return { success: false, message: 'Gagal menyetujui pinjaman: ' + e.message };
  }
}

// ============================================
// 17. PLACEHOLDER: Tolak Pinjaman oleh Admin/Ketua
// ============================================
function tolakPinjaman(idPinjaman, role, alasan) {
  try {
    const sheet = getSheet('Pinjaman');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == idPinjaman) {
        if (role === 'Admin') {
          sheet.getRange(i + 1, 6).setValue('Ditolak: ' + (alasan || '-'));
        } else if (role === 'Ketua') {
          sheet.getRange(i + 1, 7).setValue('Ditolak: ' + (alasan || '-'));
        }
        return { success: true, message: 'Pengajuan pinjaman berhasil ditolak oleh ' + role + '.' };
      }
    }
    return { success: false, message: 'Data pengajuan tidak ditemukan.' };
  } catch (e) {
    Logger.log('Error tolakPinjaman: ' + e.message);
    return { success: false, message: 'Gagal menolak pinjaman: ' + e.message };
  }
}

// ============================================
// DATABASE SETUP & SEEDER
// Jalankan fungsi ini SATU KALI dari Apps Script Editor
// ============================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Skema Database
  const schema = {
    'Users': ['ID_User', 'Username', 'Password', 'Role', 'Status'],
    'Simpanan': ['ID_Transaksi', 'ID_User', 'Tanggal', 'Jenis', 'Jumlah'],
    'Pinjaman': ['ID_Pinjaman', 'ID_User', 'Tanggal', 'Jumlah', 'Tenor', 'Status_Admin', 'Status_Ketua'],
    'DaftarPinjaman': ['ID', 'Nama', 'Alamat', 'Besar', 'Tipe', 'Tanggal', 'SistemTenor', 'Tenor', 'Saldo'],
    'Angsuran': ['ID_Angsuran', 'ID_Pinjaman', 'Nama_Nasabah', 'Tanggal_Bayar', 'Jumlah_Bayar', 'Sisa_Saldo']
  };

  // Buat sheet dan set header
  for (const sheetName in schema) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log('Sheet dibuat: ' + sheetName);
    }
    const headers = schema[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#047857')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  // Seed data dummy Users (hanya jika belum ada data)
  const userSheet = ss.getSheetByName('Users');
  if (userSheet.getLastRow() <= 1) {
    const dummyUsers = [
      ['USR-001', 'budi', 'budi123', 'Anggota', 'Aktif'],
      ['USR-002', 'siti', 'siti123', 'Anggota', 'Aktif'],
      ['USR-003', 'admin', 'admin123', 'Admin', 'Aktif'],
      ['USR-004', 'ketua', 'ketua123', 'Ketua', 'Aktif']
    ];
    userSheet.getRange(2, 1, dummyUsers.length, dummyUsers[0].length).setValues(dummyUsers);
    Logger.log('Data dummy Users berhasil ditambahkan.');
  }

  // Seed data dummy DaftarPinjaman
  const pinjamanSheet = ss.getSheetByName('DaftarPinjaman');
  if (pinjamanSheet.getLastRow() <= 1) {
    const dummyPinjaman = [
      ['1745000001', 'Budi Santoso', 'Jl. Merdeka No. 10, Jakarta', 5000000, 'Harian', '25/01/25', 'Harian', 30, 3500000],
      ['1745000002', 'Siti Rahayu', 'Jl. Sudirman No. 5, Bandung', 10000000, 'Mingguan', '15/02/25', 'Mingguan', 12, 8000000],
      ['1745000003', 'Ahmad Ridwan', 'Jl. Gatot Subroto No. 20, Surabaya', 7500000, 'Tempo', '01/03/25', 'Tempo', 3, 7500000],
    ];
    pinjamanSheet.getRange(2, 1, dummyPinjaman.length, dummyPinjaman[0].length).setValues(dummyPinjaman);
    Logger.log('Data dummy DaftarPinjaman berhasil ditambahkan.');
  }

  // Hapus Sheet1 default jika ada
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
  
  Logger.log('✅ Setup Database Selesai! Jalankan Web App untuk memulai.');
  SpreadsheetApp.getUi().alert('✅ Setup Database Berhasil!\n\nSheet yang dibuat:\n- Users (4 user dummy)\n- DaftarPinjaman (3 data dummy)\n- Angsuran\n- Simpanan\n- Pinjaman\n\nSilakan deploy sebagai Web App.');
}