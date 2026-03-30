// ======================================================================
// BACKEND: ARSIP SIABA (CRUD & DASHBOARD)
// ======================================================================

var ARSIP_SIABA_CONFIG = {
  ID_DB_UNIT: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA", // Untuk daftar unit kerja
  ID_DB_ARSIP: "1sMLUihDFeHufn5kWFG9Sj0G8xSHHOUi8usoeL4EgjqU", // Untuk menyimpan data arsip
  FOLDER_ROOT_ID: "1D0rwRT_tIj9QZTPPG3cRk4NRcbhMzDHm" // Folder Drive Induk
};

// 1. Dapatkan Daftar Unit Kerja
function arsipsiaba_getDaftarUnit() {
  try {
    var ss = SpreadsheetApp.openById(ARSIP_SIABA_CONFIG.ID_DB_UNIT);
    var sheet = ss.getSheetByName("Unit_Siaba");
    if (!sheet) return { error: "Sheet Unit_Siaba tidak ditemukan." };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { data: [] };
    
    // PERBAIKAN: Ambil Range dari Kolom A sampai C (3 Kolom)
    // Array Index -> 0: Kolom A (NPSN), 1: Kolom B (Jenjang), 2: Kolom C (Unit Kerja)
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    var units = [];
    
    for (var i = 0; i < data.length; i++) {
      // Prioritaskan mengambil nama dari Kolom C (Index 2)
      var namaUnit = data[i][2]; 
      
      // Fallback (Cadangan): Jika Kolom C ternyata kosong, baru ambil dari Kolom B
      if (!namaUnit || namaUnit === "") {
         namaUnit = data[i][1];
      }
      
      if (namaUnit && namaUnit !== "") {
        units.push(namaUnit.trim());
      }
    }
    
    // Hapus duplikat dan urutkan abjad
    units = [...new Set(units)].sort();
    return { data: units };
  } catch (e) {
    return { error: e.message };
  }
}

// 2. Dapatkan Data Arsip yang sudah diunggah
function arsipsiaba_getDataArsip(tahunFilter, bulanFilter) {
  try {
    var ss = SpreadsheetApp.openById(ARSIP_SIABA_CONFIG.ID_DB_ARSIP);
    var sheet = ss.getSheetByName("arsip_siaba");
    if (!sheet) return { error: "Sheet arsip_siaba tidak ditemukan." };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { data: [] };
    
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      
      // Filter (jika dikirim)
      if (tahunFilter && tahunFilter !== "" && row[2] !== tahunFilter) continue;
      if (bulanFilter && bulanFilter !== "" && row[1] !== bulanFilter) continue;
      
      result.push({
        rowId: i + 2,
        unitKerja: row[0],
        bulan: row[1],
        tahun: row[2],
        fileUrl: row[3],
        tglUnggah: row[4],
        userUnggah: row[5],
        tglEdit: row[6],
        userEdit: row[7]
      });
    }
    
    return { data: result };
  } catch (e) {
    return { error: e.message };
  }
}

// 3. Helper untuk membuat Folder Bersarang (Nested Folder)
function arsipsiaba_getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

// 4. Proses Upload File & Simpan Database
function arsipsiaba_simpanArsip(payload) {
  try {
    var isEdit = payload.mode === "edit";
    var ss = SpreadsheetApp.openById(ARSIP_SIABA_CONFIG.ID_DB_ARSIP);
    var sheet = ss.getSheetByName("arsip_siaba");
    if (!sheet) return { error: "Sheet arsip_siaba tidak ditemukan." };
    
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var fileUrl = payload.existingUrl || "";
    
    // VAKSIN 1: Normalisasi Payload (Hapus spasi tersembunyi & jadikan kapital)
    var pUnit  = String(payload.unitKerja).trim().toUpperCase();
    var pBulan = String(payload.bulan).trim().toUpperCase();
    var pTahun = String(payload.tahun).trim().toUpperCase();
    
    // Cek Duplikat Jika Insert
    if (!isEdit) {
      var existData = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < existData.length; i++) {
        // VAKSIN 2: Normalisasi Data di Spreadsheet
        var dbUnit  = String(existData[i][0]).trim().toUpperCase();
        var dbBulan = String(existData[i][1]).trim().toUpperCase();
        var dbTahun = String(existData[i][2]).trim().toUpperCase();
        
        // VAKSIN 3: Pengecekan Mutlak 3 Kombinasi Secara Presisi
        if (dbUnit === pUnit && dbBulan === pBulan && dbTahun === pTahun) {
          return { error: "Arsip untuk " + payload.unitKerja + " bulan " + payload.bulan + " " + payload.tahun + " sudah ada. Silakan gunakan tombol Edit pada tabel." };
        }
      }
    }
    
    // Proses File Baru Jika Ada
    if (payload.fileData && payload.fileData.data) {
      var rootFolder = DriveApp.getFolderById(ARSIP_SIABA_CONFIG.FOLDER_ROOT_ID);
      var folderTahun = arsipsiaba_getOrCreateFolder(rootFolder, payload.tahun);
      var folderBulan = arsipsiaba_getOrCreateFolder(folderTahun, payload.bulan);
      
      var fileName = "Siaba " + payload.bulan + " " + payload.tahun + " " + payload.unitKerja + ".pdf";
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.fileData.data), payload.fileData.mimeType, fileName);
      
      var file = folderBulan.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }
    
    if (isEdit) {
      var rowIdx = parseInt(payload.rowId);
      sheet.getRange(rowIdx, 4).setValue(fileUrl); // Kolom D
      sheet.getRange(rowIdx, 7).setValue(timestamp); // Kolom G
      sheet.getRange(rowIdx, 8).setValue(payload.userLogin); // Kolom H
    } else {
      var newRow = [
        payload.unitKerja, // A
        payload.bulan,     // B
        payload.tahun,     // C
        fileUrl,           // D
        timestamp,         // E
        payload.userLogin, // F
        "",                // G
        ""                 // H
      ];
      sheet.appendRow(newRow);
    }
    
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

// 5. Data Dashboard (Belum Lapor)
function arsipsiaba_getDashboardBelumLapor(tahun, bulan) {
  try {
    // Ambil Semua Unit
    var refUnits = arsipsiaba_getDaftarUnit();
    if (refUnits.error) return { error: refUnits.error };
    var allUnits = refUnits.data;
    
    // Ambil Data Arsip yg Sudah Lapor
    var dataArsip = arsipsiaba_getDataArsip(tahun, bulan);
    if (dataArsip.error) return { error: dataArsip.error };
    
    var unitSudahLapor = new Set();
    dataArsip.data.forEach(function(row) {
      unitSudahLapor.add(row.unitKerja);
    });
    
    var belumLapor = [];
    allUnits.forEach(function(u) {
      if (!unitSudahLapor.has(u)) {
        belumLapor.push(u);
      }
    });
    
    return { 
      data: {
        sudah: unitSudahLapor.size,
        belum: belumLapor.length,
        total: allUnits.length,
        listBelum: belumLapor
      }
    };
  } catch (e) {
    return { error: e.message };
  }
}

// 6. Proses Hapus Arsip (Soft Delete)
function arsipsiaba_hapusArsip(rowId, userLogin) {
  try {
    var ss = SpreadsheetApp.openById(ARSIP_SIABA_CONFIG.ID_DB_ARSIP);
    var sheet = ss.getSheetByName("arsip_siaba");
    var sheetTrash = ss.getSheetByName("arsip_siaba_trash");
    
    // Auto-create sheet trash jika belum ada
    if (!sheetTrash) {
      sheetTrash = ss.insertSheet("arsip_siaba_trash");
      sheetTrash.appendRow(["Unit Kerja", "Bulan", "Tahun", "File URL", "Tgl Unggah", "User Unggah", "Tgl Edit", "User Edit", "Tgl Hapus", "User Hapus"]);
    }
    
    var rowIdx = parseInt(rowId);
    if (isNaN(rowIdx) || rowIdx < 2) return { error: "ID Baris tidak valid." };
    
    var rowData = sheet.getRange(rowIdx, 1, 1, 8).getDisplayValues()[0];
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    
    // Tambahkan info waktu hapus dan eksekutor
    rowData.push(timestamp, userLogin);
    
    // Pindah ke Trash lalu hapus dari Utama
    sheetTrash.appendRow(rowData);
    sheet.deleteRow(rowIdx);
    
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}