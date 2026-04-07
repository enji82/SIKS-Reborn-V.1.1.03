/* ======================================================================
   MODUL: KELOLA PTK SD
   Spreadsheet ID: 1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE
   Sheet: Master Data GTK
   ====================================================================== */

var ID_DB_PTK = "1t0-Lmy0YD_GxHzimFWJGh5R5x6RhGL13uqKeVwWoCYE";
var SHEET_PTK = "Master Data GTK";

// 1. AMBIL OPSI FILTER (UNIT & STATUS)
function getFilterOptionsPTK() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = "ptk_filter_options";
    const cached = cache.get(cacheKey);
    if (cached) return cached;
    
    var ss = SpreadsheetApp.openById(ID_DB_PTK);
    var sheet = ss.getSheetByName(SHEET_PTK);
    if (!sheet) return JSON.stringify({ units: [], statuses: [] });
    
    var lastRow = sheet.getLastRow();
    if(lastRow < 2) return JSON.stringify({ units: [], statuses: [] });

    var data = sheet.getRange(2, 1, lastRow - 1, 20).getValues(); // Diperlebar
    var unitSet = new Set();
    var statusSet = new Set();
    
    for(var i=0; i<data.length; i++){
        if(data[i][2]) unitSet.add(String(data[i][2]).trim());
        if(data[i][19]) statusSet.add(String(data[i][19]).trim()); // Status Pegawai bergeser ke T (19)
    }
    
    const result = JSON.stringify({ units: Array.from(unitSet).sort(), statuses: Array.from(statusSet).sort() });
    cache.put(cacheKey, result, 3600);
    return result;
  } catch(e) { return JSON.stringify({ error: "Terjadi kesalahan saat mengambil filter." }); }
}

// 2. AMBIL DATA UTAMA (OPTIMASI DISPLAY VALUES)
function getDataPTKSD(filterUnit, filterStatus) {
  var ss = SpreadsheetApp.openById(ID_DB_PTK);
  var sheet = ss.getSheetByName(SHEET_PTK);
  var data = sheet.getDataRange().getValues();
  data.shift(); 
  
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var tglLahirISO = parseIndoDate(row[9]);
    var tmtJabISO   = parseIndoDate(row[21]); // Bergeser +1
    var tmtGolISO   = parseIndoDate(row[23]); // Bergeser +1

    result.push({
      id: row[0],              // A
      npsn: row[1],            // B
      unit: row[2],            // C
      gelar_depan: row[3],     // D
      nama_no_gelar: row[4],   // E
      gelar_belakang: row[5],  // F
      nama_lengkap: row[6],    // G
      nip: row[7],             // H
      tmp_lahir: row[8],       // I
      tgl_lahir: tglLahirISO,  // J
      nik: row[10],            // K
      lp: row[11],             // L
      agama: row[12],          // M
      pendidikan: row[13],     // N
      jurusan: row[14],        // O
      thn_lulus: row[15],      // P
      alamat_ktp: row[16],     // Q (Alamat KTP)
      alamat_domisili: row[17],// R (Alamat Domisili) - KOLOM BARU
      hp: row[18],             // S
      status_peg: row[19],     // T
      jabatan: row[20],        // U
      tmt_jabatan: tmtJabISO,  // V
      pangkat: row[22],        // W
      tmt_gol: tmtGolISO,      // X
      mkg: row[24],            // Y
      kelas_jab: row[25],      // Z
      tugas: row[26],          // AA
      nuptk: row[27],          // AB
      serdik: row[28],         // AC
      dapodik: row[29],        // AD
      tugtam: row[30],         // AE
      email: row[31],          // AF
      diinput: row[32] ? Utilities.formatDate(new Date(row[32]), Session.getScriptTimeZone(), "dd/MM/yy HH:mm") : "", // AG
      user_input: row[33],     // AH
      diedit: row[34] ? Utilities.formatDate(new Date(row[34]), Session.getScriptTimeZone(), "dd/MM/yy HH:mm") : "",  // AI
      user_edit: row[35]       // AJ
    });
  }
  return JSON.stringify(result);
}

// 3. UPDATE DATA PTK
function updateDataPTK(form) {
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PTK);
    var sheet = ss.getSheetByName(SHEET_PTK);
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for(var i=1; i<data.length; i++){ if(String(data[i][0]) === String(form.id)){ rowIndex = i + 1; break; } }
    if(rowIndex === -1) return "Error: ID PTK tidak ditemukan.";

    var inputNip = String(form.nip || "").trim().replace(/[^0-9]/g, '');
    if (inputNip !== "" && inputNip !== "-") {
        for (var i = 1; i < data.length; i++) {
            var rowNip = String(data[i][7]).replace(/[^0-9]/g, ''); 
            var rowId = String(data[i][0]);
            if (rowNip === inputNip && rowId !== String(form.id)) return "Gagal: NIP " + inputNip + " sudah dipakai oleh " + data[i][6];
        }
    }

    var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + form.nama_lengkap + (form.gelar_belakang ? ", " + form.gelar_belakang : "");
    var mkg = ""; if (form.mkg_thn || form.mkg_bln) { mkg = (form.mkg_thn || "0") + " Tahun " + (form.mkg_bln || "0") + " Bulan"; }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    var user = form.user_login || "Admin";

    if (form.npsn_baru && form.unit_kerja) {
        sheet.getRange(rowIndex, 2).setValue("'" + form.npsn_baru); 
        sheet.getRange(rowIndex, 3).setValue(form.unit_kerja);      
    }

    sheet.getRange(rowIndex, 4).setValue(form.gelar_depan || "");         // D
    sheet.getRange(rowIndex, 5).setValue(form.nama_lengkap || "");        // E
    sheet.getRange(rowIndex, 6).setValue(form.gelar_belakang || "");      // F
    sheet.getRange(rowIndex, 7).setValue(namaFull);                       // G
    sheet.getRange(rowIndex, 8).setValue("'"+(form.nip || ""));           // H
    sheet.getRange(rowIndex, 9).setValue(form.tmp_lahir || "");           // I
    sheet.getRange(rowIndex, 10).setValue(form.tgl_lahir || "");    // J
    sheet.getRange(rowIndex, 11).setValue("'"+(form.nik || ""));          // K
    sheet.getRange(rowIndex, 12).setValue(form.lp || "");                 // L
    sheet.getRange(rowIndex, 13).setValue(form.agama || "");              // M
    sheet.getRange(rowIndex, 14).setValue(form.pendidikan || "");         // N
    sheet.getRange(rowIndex, 15).setValue(form.jurusan || "");            // O
    sheet.getRange(rowIndex, 16).setValue(form.thn_lulus || "");          // P
    sheet.getRange(rowIndex, 17).setValue(form.alamat_ktp || "");         // Q (Alamat KTP)
    sheet.getRange(rowIndex, 18).setValue(form.alamat_domisili || "");    // R (Alamat Domisili)
    sheet.getRange(rowIndex, 19).setValue("'"+(form.hp || ""));           // S (HP)
    sheet.getRange(rowIndex, 20).setValue(form.status_peg || "");         // T
    sheet.getRange(rowIndex, 21).setValue(form.jabatan || "");            // U
    sheet.getRange(rowIndex, 22).setValue(form.tmt_jabatan || "");  // V
    sheet.getRange(rowIndex, 23).setValue(form.pangkat || "");            // W
    sheet.getRange(rowIndex, 24).setValue(form.tmt_gol || "");      // X
    sheet.getRange(rowIndex, 25).setValue(mkg);                           // Y
    sheet.getRange(rowIndex, 27).setValue(form.tugas || "");              // AA
    sheet.getRange(rowIndex, 28).setValue("'"+(form.nuptk || ""));        // AB
    sheet.getRange(rowIndex, 29).setValue(form.serdik || "");             // AC
    sheet.getRange(rowIndex, 30).setValue(form.dapodik || "");            // AD
    sheet.getRange(rowIndex, 31).setValue(form.tugtam || "");             // AE
    sheet.getRange(rowIndex, 32).setValue(form.email || "");              // AF 
    
    sheet.getRange(rowIndex, 35).setValue(now);                           // AI (Diedit)
    sheet.getRange(rowIndex, 36).setValue(user);                          // AJ (User Edit)

    return "Sukses";
  } catch(e) { return "Error: " + e.message; }
}

// 4. INSERT DATA PTK (AUTO FILL LOGIC)
function insertDataPTK(form) {
  var ss = SpreadsheetApp.openById(ID_DB_PTK); 
  var sheet = ss.getSheetByName(SHEET_PTK);
  if (!sheet) return "Error: Sheet 'Master Data GTK' tidak ditemukan.";

  var inputNip = String(form.nip || "").trim().replace(/[^0-9]/g, ''); 
  if (inputNip !== "" && inputNip !== "-") {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowNip = String(data[i][7]).replace(/[^0-9]/g, ''); 
        if (rowNip === inputNip) return "Gagal: NIP " + inputNip + " sudah terdaftar atas nama " + data[i][6];
      }
  }

  var newId = "GTK-" + new Date().getTime();
  var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + form.nama_lengkap + (form.gelar_belakang ? ", " + form.gelar_belakang : "");
  var mkg = ""; if (form.mkg_thn || form.mkg_bln) { mkg = (form.mkg_thn || "0") + " Tahun " + (form.mkg_bln || "0") + " Bulan"; }
  var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");

  var rowData = [
      newId,                  // A (0)
      form.npsn_baru || form.npsn_login || "",  // B (1)
      form.unit_kerja || form.unit_login || "", // C (2)
      form.gelar_depan || "", // D (3)
      form.nama_lengkap || "",// E (4)
      form.gelar_belakang || "",// F (5)
      namaFull || "",         // G (6)
      "'" + (form.nip || ""), // H (7)
      form.tmp_lahir || "",   // I (8)
      form.tgl_lahir || "",   // J (9)
      "'" + (form.nik || ""), // K (10)
      form.lp || "",          // L (11)
      form.agama || "",       // M (12)
      form.pendidikan || "",  // N (13)
      form.jurusan || "",     // O (14)
      form.thn_lulus || "",   // P (15)
      form.alamat_ktp || "",  // Q (16)
      form.alamat_domisili || "", // R (17) NEW
      "'" + (form.hp || ""),  // S (18)
      form.status_peg || "",  // T (19)
      form.jabatan || "",     // U (20)
      form.tmt_jabatan || "", // V (21)
      form.pangkat || "",     // W (22)
      form.tmt_gol || "",     // X (23)
      mkg,                    // Y (24) 
      "",                     // Z (25)
      form.tugas || "",       // AA (26)
      "'" + (form.nuptk || ""),// AB (27)
      form.serdik || "",      // AC (28)
      form.dapodik || "",     // AD (29)
      form.tugtam || "",      // AE (30)
      form.email || "",       // AF (31) 
      timestamp,              // AG (32)
      form.user_login || "",  // AH (33)
      "",                     // AI (34)
      ""                      // AJ (35)
  ];

  sheet.appendRow(rowData);
  return "Sukses";
}

// ======================================================================
// MODUL: REFERENSI (TIDAK BERUBAH)
// ======================================================================
function getReferensiPTK() {
  var ss = SpreadsheetApp.openById(ID_DB_PTK);
  function getColData(sheetName, colIndex) {
    var s = ss.getSheetByName(sheetName); if (!s) return []; var last = s.getLastRow(); if (last < 2) return [];
    var data = s.getRange(2, colIndex, last - 1, 1).getValues(); var res = [];
    for (var i = 0; i < data.length; i++) { var val = String(data[i][0]).trim(); if (val !== "") res.push(val); } return res;
  }
  function getPangkat() {
     var s = ss.getSheetByName("data_pangkat"); if(!s) return []; var last = s.getLastRow(); if (last < 2) return [];
     var data = s.getRange(2, 1, last-1, 1).getValues(); var res = [];
     for(var i=0; i<data.length; i++) { var val = String(data[i][0]).trim(); if(val !== "") res.push(val); } return res;
  }
  return JSON.stringify({ jabatan_non_asn: getColData("isian_jabatan", 1), jabatan_asn: getColData("isian_jabatan", 2), tugas_non_asn: getColData("isian_tugas_di_sekolah", 1), tugas_asn: getColData("isian_tugas_di_sekolah", 2), pangkat: getPangkat() });
}

function getUnitKerjaByNpsnPTK(npsn) {
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Database Sekolah");
    if (!sheet) return JSON.stringify({ error: "Sheet 'Database Sekolah' tidak ditemukan." });
    var lastRow = sheet.getLastRow(); if (lastRow < 2) return JSON.stringify({ error: "Database Sekolah kosong." });
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues(); var searchNpsn = String(npsn).trim().toUpperCase();
    for (var i = 0; i < data.length; i++) { if (String(data[i][0]).trim().toUpperCase() === searchNpsn) return JSON.stringify({ unitKerja: String(data[i][2]).trim() }); }
    return JSON.stringify({ error: "NPSN tidak terdaftar." });
  } catch (e) { return JSON.stringify({ error: "Gagal memuat Database Sekolah PTK: " + e.message }); }
}

function parseIndoDate(dateStr) {
  if (!dateStr || dateStr === "-" || dateStr === "") return "";
  var str = String(dateStr).trim();
  if (str.match(/^\d{4}-\d{2}-\d{2}$/)) return str;
  var slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) { var day = slashMatch[1].length === 1 ? "0" + slashMatch[1] : slashMatch[1]; var month = slashMatch[2].length === 1 ? "0" + slashMatch[2] : slashMatch[2]; return slashMatch[3] + "-" + month + "-" + day; }
  var months = { 'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04', 'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08', 'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12', 'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'Jun': '06', 'Jul': '07', 'Agu': '08', 'Sep': '09', 'Okt': '10', 'Nov': '11', 'Des': '12' };
  var parts = str.split(' '); 
  if (parts.length >= 3) { var dayRaw = parts[0].replace(/[^0-9]/g, ''); var day = dayRaw.length === 1 ? "0" + dayRaw : dayRaw; var month = months[parts[1]]; if (month && parts[2].match(/^\d{4}$/)) return parts[2] + "-" + month + "-" + day; }
  try { var d = new Date(dateStr); if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"); } catch(e) {}
  return "";
}

function moveDataPTKToNonAktif(id, reason, userLogin) {
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheetSource = ss.getSheetByName(SHEET_PTK); var sheetTarget = ss.getSheetByName("gtk_non_aktif");
    if (!sheetTarget) { sheetTarget = ss.insertSheet("gtk_non_aktif"); var headers = sheetSource.getRange(1, 1, 1, sheetSource.getLastColumn()).getValues(); headers[0].push("Alasan Hapus", "Tanggal Hapus", "User Hapus"); sheetTarget.getRange(1, 1, 1, headers[0].length).setValues(headers); }
    var data = sheetSource.getDataRange().getValues(); var rowIndex = -1;
    for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { rowIndex = i; break; } }
    if (rowIndex === -1) return "Data tidak ditemukan.";
    var rowData = data[rowIndex]; var deleteTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    rowData.push(reason, deleteTime, userLogin); sheetTarget.appendRow(rowData); sheetSource.deleteRow(rowIndex + 1);
    return "Sukses";
  } catch (e) { return "Error: " + e.message; }
}

function getDataKeadaanGTK() { var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Keadaan GTK"); if (!sheet) return []; var lastRow = sheet.getLastRow(); if (lastRow < 3) return []; return sheet.getRange(3, 1, lastRow - 2, 56).getDisplayValues(); }
function getDataKebutuhanGuru() { var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Kebutuhan Guru"); if (!sheet) return []; var lastRow = sheet.getLastRow(); if (lastRow < 3) return []; return sheet.getRange(3, 1, lastRow - 2, 42).getDisplayValues(); }

// =============================================================
// BACKEND: KELOLA DATA PTK SD SWASTA (SDS)
// =============================================================
function getDataPTKSDS() {
  try {
    var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Master Data GTK SDS"); if (!sheet) return JSON.stringify([]);
    var lastRow = sheet.getLastRow(); if (lastRow < 2) return JSON.stringify([]); 
    var data = sheet.getRange(2, 1, lastRow - 1, 33).getDisplayValues(); // Diperlebar
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i]; if(row[0] === "") continue; 
      result.push({
        id: row[0], npsn: row[1], unit: row[2], gelar_depan: row[3], nama_no_gelar: row[4], gelar_belakang: row[5], 
        nama_lengkap: row[6], niy: row[7], tmp_lahir: row[8], tgl_lahir: row[9], nik: row[10], lp: row[11], agama: row[12],         
        pendidikan: row[13], jurusan: row[14], thn_lulus: row[15], 
        alamat_ktp: row[16], alamat_domisili: row[17], hp: row[18], status_peg: row[19], jabatan: row[20], tmt_jabatan: row[21],   
        inpassing: row[22], tmt_inpassing: row[23], nuptk: row[24], serdik: row[25], dapodik: row[26], tugtam: row[27], 
        diinput: row[28], user_input: row[29], diedit: row[30], user_edit: row[31], email: row[32] || "" 
      });
    }
    return JSON.stringify(result);
  } catch(e) { return JSON.stringify([]); }
}

function insertDataPTKSDS(form) {
  var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Master Data GTK SDS");
  var data = sheet.getDataRange().getValues(); var inputNik = String(form.nik).trim(); 
  for (var i = 1; i < data.length; i++) { var rowNik = String(data[i][10]).replace(/'/g, "").trim(); if (rowNik === inputNik) return "NIK " + inputNik + " sudah terdaftar atas nama " + data[i][6] + ", hubungi admin."; }
  var newId = "SDS-" + new Date().getTime(); var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + form.nama_lengkap + (form.gelar_belakang ? ", " + form.gelar_belakang : ""); var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  var rowData = [
    newId, form.npsn_baru || form.npsn_login || "", form.unit_kerja || form.unit_login || "", form.gelar_depan || "", form.nama_lengkap || "", form.gelar_belakang || "",    
    namaFull || "", form.niy || "", form.tmp_lahir || "", form.tgl_lahir || "", "'" + (form.nik || ""), form.lp || "", form.agama || "",             
    form.pendidikan || "", form.jurusan || "", form.thn_lulus || "", 
    form.alamat_ktp || "", form.alamat_domisili || "", "'" + (form.hp || ""), form.status_peg || "", form.jabatan || "", form.tmt_jabatan || "",       
    form.inpassing || "", form.tmt_inpassing || "", "'" + (form.nuptk || ""), form.serdik || "", form.dapodik || "", form.tugtam || "",            
    timestamp, form.user_login || "", "", "", form.email || ""  
  ];
  sheet.appendRow(rowData); return "Sukses";
}

function updateDataPTKSDS(form) {
  var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Master Data GTK SDS"); var data = sheet.getDataRange().getValues();
  var rowIdx = -1; for (var i = 0; i < data.length; i++) { if (data[i][0] == form.id) { rowIdx = i + 1; break; } }
  if (rowIdx == -1) return "Error: ID tidak ditemukan.";
  var inputNik = String(form.nik || "").trim();
  if (inputNik !== "") { for (var i = 1; i < data.length; i++) { var rowNik = String(data[i][10]).replace(/'/g, '').trim(); if (rowNik === inputNik && String(data[i][0]) !== String(form.id)) return "Gagal: NIK " + inputNik + " sudah dipakai oleh " + data[i][6]; } }
  var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + form.nama_lengkap + (form.gelar_belakang ? ", " + form.gelar_belakang : "");
  var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  if (form.npsn_baru && form.unit_kerja) { sheet.getRange(rowIdx, 2).setValue("'" + form.npsn_baru); sheet.getRange(rowIdx, 3).setValue(form.unit_kerja); }
  var updateValues = [[ form.gelar_depan || "", form.nama_lengkap || "", form.gelar_belakang || "", namaFull || "", form.niy || "", form.tmp_lahir || "", form.tgl_lahir || "", "'" + (form.nik || ""), form.lp || "", form.agama || "", form.pendidikan || "", form.jurusan || "", form.thn_lulus || "", form.alamat_ktp || "", form.alamat_domisili || "", "'" + (form.hp || ""), form.status_peg || "", form.jabatan || "", form.tmt_jabatan || "", form.inpassing || "", form.tmt_inpassing || "", "'" + (form.nuptk || ""), form.serdik || "", form.dapodik || "", form.tugtam || "" ]];
  sheet.getRange(rowIdx, 4, 1, 25).setValues(updateValues); // Digeser 25 Kolom
  sheet.getRange(rowIdx, 31).setValue(timestamp); sheet.getRange(rowIdx, 32).setValue(form.user_login); sheet.getRange(rowIdx, 33).setValue(form.email || ""); return "Sukses";
}

function deleteDataPTKSDS(id, alasan, userLogin) {
  var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheetSource = ss.getSheetByName("Master Data GTK SDS"); var sheetTarget = ss.getSheetByName("gtk_non_aktif_sds"); 
  if (!sheetTarget) { sheetTarget = ss.insertSheet("gtk_non_aktif_sds"); var headers = sheetSource.getRange(1, 1, 1, sheetSource.getLastColumn()).getValues(); headers[0].push("Alasan Hapus", "Tanggal Hapus", "User Hapus"); sheetTarget.getRange(1, 1, 1, headers[0].length).setValues(headers); }
  var data = sheetSource.getDataRange().getValues(); var rowIdx = -1; var rowData = [];
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === String(id)) { rowIdx = i + 1; rowData = data[i]; break; } }
  if (rowIdx == -1) return "Error: Data tidak ditemukan.";
  rowData.push(alasan, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss"), userLogin);
  sheetTarget.appendRow(rowData); sheetSource.deleteRow(rowIdx); return "Sukses";
}

function getDataKeadaanGTKSDS() { var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Keadaan GTK SDS"); if (!sheet) return []; var lastRow = sheet.getLastRow(); if (lastRow < 3) return []; return sheet.getRange(3, 1, lastRow - 2, 27).getDisplayValues(); }
function getDataKebutuhanGuruSDS() { var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Kebutuhan Guru SDS"); if (!sheet) return []; var lastRow = sheet.getLastRow(); if (lastRow < 3) return []; return sheet.getRange(3, 1, lastRow - 2, 27).getDisplayValues(); }

// =============================================================
// BACKEND: KELOLA DATA PTK PAUD
// ID Spreadsheet: 1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU
// =============================================================
var ID_SPREADSHEET_PAUD = "1XetGkBymmN2NZQlXpzZ2MQyG0nhhZ0sXEPcNsLffhEU";

function getDataPTKPAUD() {
  try {
    var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD); var sheet = ss.getSheetByName("Master Data GTK PAUD"); if (!sheet) return JSON.stringify([]);
    var lastRow = sheet.getLastRow(); if (lastRow < 2) return JSON.stringify([]); 
    var data = sheet.getRange(2, 1, lastRow - 1, 34).getDisplayValues(); 
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i]; if(row[0] === "") continue; 
      result.push({
        id: row[0], npsn: row[1], unit: row[2], jenjang: row[3], gelar_depan: row[4], nama_no_gelar: row[5], gelar_belakang: row[6], nama_lengkap: row[7], niy: row[8], tmp_lahir: row[9], tgl_lahir: row[10], nik: row[11], lp: row[12], agama: row[13], pendidikan: row[14], jurusan: row[15], thn_lulus: row[16], 
        alamat_ktp: row[17], alamat_domisili: row[18], hp: row[19], status_peg: row[20], jabatan: row[21], tmt_jabatan: row[22], inpassing: row[23], tmt_inpassing: row[24], nuptk: row[25], serdik: row[26], dapodik: row[27], tugtam: row[28], 
        diinput: row[29], user_input: row[30], diedit: row[31], user_edit: row[32], email: row[33] || "" 
      });
    }
    return JSON.stringify(result);
  } catch(e) { return JSON.stringify([]); }
}

function insertDataPTKPAUD(form) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD); var sheet = ss.getSheetByName("Master Data GTK PAUD");
  var data = sheet.getDataRange().getValues(); var inputNik = String(form.nik).trim(); 
  for (var i = 1; i < data.length; i++) { var rowNik = String(data[i][11]).replace(/'/g, "").trim(); if (rowNik === inputNik) return "NIK " + inputNik + " sudah terdaftar atas nama " + data[i][7] + "."; }
  var newId = "PAUD-" + new Date().getTime(); var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + form.nama_lengkap + (form.gelar_belakang ? ", " + form.gelar_belakang : ""); var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  var rowData = [
    newId, form.npsn_baru || form.npsn_login || "", form.unit_kerja || form.unit_login || "", form.jenjang || "", form.gelar_depan || "", form.nama_lengkap || "", form.gelar_belakang || "", namaFull || "", form.niy || "", form.tmp_lahir || "", form.tgl_lahir || "", "'" + (form.nik || ""), form.lp || "", form.agama || "", form.pendidikan || "", form.jurusan || "", form.thn_lulus || "", 
    form.alamat_ktp || "", form.alamat_domisili || "", "'" + (form.hp || ""), form.status_peg || "", form.jabatan || "", form.tmt_jabatan || "", form.inpassing || "", form.tmt_inpassing || "", "'" + (form.nuptk || ""), form.serdik || "", form.dapodik || "", form.tugtam || "", timestamp, form.user_login || "", "", "", form.email || ""      
  ];
  sheet.appendRow(rowData); return "Sukses";
}

function updateDataPTKPAUD(form) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD); var sheet = ss.getSheetByName("Master Data GTK PAUD"); var data = sheet.getDataRange().getValues();
  var rowIdx = -1; for (var i = 0; i < data.length; i++) { if (data[i][0] == form.id) { rowIdx = i + 1; break; } }
  if (rowIdx == -1) return "Error: ID tidak ditemukan.";
  var inputNik = String(form.nik || "").trim();
  if (inputNik !== "") { for (var i = 1; i < data.length; i++) { var rowNik = String(data[i][11]).replace(/'/g, '').trim(); if (rowNik === inputNik && String(data[i][0]) != form.id) return "Gagal: NIK " + inputNik + " sudah dipakai oleh " + data[i][7]; } }
  var namaFull = (form.gelar_depan ? form.gelar_depan + " " : "") + form.nama_lengkap + (form.gelar_belakang ? ", " + form.gelar_belakang : ""); var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
  if (form.npsn_baru && form.unit_kerja) { sheet.getRange(rowIdx, 2).setValue("'" + form.npsn_baru); sheet.getRange(rowIdx, 3).setValue(form.unit_kerja); }
  var updateValues = [[ form.jenjang || "", form.gelar_depan || "", form.nama_lengkap || "", form.gelar_belakang || "", namaFull || "", form.niy || "", form.tmp_lahir || "", form.tgl_lahir || "", "'" + (form.nik || ""), form.lp || "", form.agama || "", form.pendidikan || "", form.jurusan || "", form.thn_lulus || "", form.alamat_ktp || "", form.alamat_domisili || "", "'" + (form.hp || ""), form.status_peg || "", form.jabatan || "", form.tmt_jabatan || "", form.inpassing || "", form.tmt_inpassing || "", "'" + (form.nuptk || ""), form.serdik || "", form.dapodik || "", form.tugtam || "" ]];
  sheet.getRange(rowIdx, 4, 1, 26).setValues(updateValues); // Digeser ke 26 Kolom
  sheet.getRange(rowIdx, 32).setValue(timestamp); sheet.getRange(rowIdx, 33).setValue(form.user_login); sheet.getRange(rowIdx, 34).setValue(form.email || ""); return "Sukses";
}

function deleteDataPTKPAUD(id, alasan, userLogin) {
  var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD); var sheetSource = ss.getSheetByName("Master Data GTK PAUD"); var sheetTarget = ss.getSheetByName("gtk_non_aktif_paud"); 
  if (!sheetTarget) { sheetTarget = ss.insertSheet("gtk_non_aktif_paud"); var headers = sheetSource.getRange(1, 1, 1, sheetSource.getLastColumn()).getValues(); headers[0].push("Alasan Hapus", "Tanggal Hapus", "User Hapus"); sheetTarget.getRange(1, 1, 1, headers[0].length).setValues(headers); }
  var data = sheetSource.getDataRange().getValues(); var rowIdx = -1; var rowData = [];
  for (var i = 1; i < data.length; i++) { if (data[i][0] == id) { rowIdx = i + 1; rowData = data[i]; break; } }
  if (rowIdx == -1) return "Error: Data tidak ditemukan.";
  rowData.push(alasan, Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss"), userLogin); sheetTarget.appendRow(rowData); sheetSource.deleteRow(rowIdx); return "Sukses";
}

function getJenjangByNPSN(npsn) {
  try { var ss = SpreadsheetApp.openById(ID_DB_PTK); var sheet = ss.getSheetByName("Database Sekolah"); if (!sheet) return "Sheet Tidak Ditemukan"; var lastRow = sheet.getLastRow(); var data = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues(); var searchNpsn = String(npsn).trim(); for (var i = 0; i < data.length; i++) { if (String(data[i][0]).trim() === searchNpsn) return String(data[i][1]).trim(); } return ""; } catch (e) { return ""; }
}

function getDataKeadaanGTKPAUD() { var ss = SpreadsheetApp.openById(ID_SPREADSHEET_PAUD); var sheet = ss.getSheetByName("Keadaan GTK PAUD"); if (!sheet) return []; var lastRow = sheet.getLastRow(); if (lastRow < 3) return []; return sheet.getRange(3, 1, lastRow - 2, 28).getDisplayValues(); }