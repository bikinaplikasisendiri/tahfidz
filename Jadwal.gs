// ========================================
// JADWAL.gs - Jadwal Pelajaran
// ========================================

// --- JADWAL HAFALAN PER KELAS ---
function getJadwalByKelas(kelas) {
  var data = getSheetData('Data_Jadwal');
  var result = [];
  if (data && data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(kelas)) {
        var jm = data[i][2];
        var js = data[i][3];
        if (jm instanceof Date) jm = jm.getHours().toString().padStart(2,'0') + ':' + jm.getMinutes().toString().padStart(2,'0');
        else jm = String(jm || '');
        if (js instanceof Date) js = js.getHours().toString().padStart(2,'0') + ':' + js.getMinutes().toString().padStart(2,'0');
        else js = String(js || '');
        
        result.push({
          no: i,
          hari: String(data[i][0] || ''),
          kelas: String(data[i][1] || ''),
          jamMulai: jm,
          jamSelesai: js,
          mapel: String(data[i][4] || ''),
          ruangan: String(data[i][5] || ''),
          guru: String(data[i][6] || '')
        });
      }
    }
  }
  return result;
}

// --- TAMBAH JADWAL ---
function addJadwal(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Jadwal');
  if (!sheet) return { status: 'error', message: 'Sheet Data_Jadwal tidak ditemukan!' };
  
  // Set format kolom waktu sebagai teks
  try {
    var maxRow = sheet.getMaxRows();
    sheet.getRange(1, 3, maxRow, 2).setNumberFormat('@');
  } catch(e) {}
  
  var rowData = [
    String(d.hari || '').trim(),
    String(d.kelas || '').trim(),
    padTime_(d.jamMulai),
    padTime_(d.jamSelesai),
    String(d.mapel || '').trim(),
    String(d.ruangan || '').trim(),
    String(d.guru || '').trim(),
    String(d.guruLogin || '').trim()
  ];
  
  sheet.appendRow(rowData);
  CacheService.getScriptCache().remove('sheet_Data_Jadwal');
  
  return { status: 'success', message: 'Rows after save: ' + sheet.getLastRow() };
}

// --- UPDATE JADWAL ---
function updateJadwal(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Jadwal');
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan!' };
  var r = d.no + 1;
  sheet.getRange(r, 1).setValue(d.hari);
  sheet.getRange(r, 2).setValue(d.kelas);
  sheet.getRange(r, 3).setValue(d.jamMulai || '');
  sheet.getRange(r, 4).setValue(d.jamSelesai || '');
  sheet.getRange(r, 5).setValue(d.mapel || '');
  sheet.getRange(r, 6).setValue(d.ruangan || '');
  sheet.getRange(r, 7).setValue(d.guru || '');
  CacheService.getScriptCache().remove('sheet_Data_Jadwal');
  return { status: 'success' };
}

// --- HAPUS JADWAL ---
function deleteJadwal(no) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Jadwal');
  if (!sheet) return { status: 'error' };
  sheet.deleteRow(no + 1);
  CacheService.getScriptCache().remove('sheet_Data_Jadwal');
  return { status: 'success' };
}

// --- BUAT SHEET JADWAL ---
function createJadwalSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Jadwal');
  if (!sheet) {
    sheet = ss.insertSheet('Data_Jadwal');
    sheet.appendRow(['Hari', 'Kelas', 'Jam_Mulai', 'Jam_Selesai', 'Mata_Pelajaran', 'Ruangan', 'Guru', 'Guru_Login']);
    sheet.setFrozenRows(1);
    return 'Sheet Data_Jadwal berhasil dibuat!';
  }
  return 'Sheet Data_Jadwal sudah ada dengan ' + sheet.getLastRow() + ' baris.';
}

// --- FIX JADWAL TIME ---
function fixJadwalTime() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Jadwal');
  if (!sheet) return 'Sheet tidak ditemukan';
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'Tidak ada data';
  
  // Set format teks untuk kolom waktu
  sheet.getRange(2, 3, lastRow - 1, 2).setNumberFormat('@');
  
  var fixed = 0;
  
  for (var r = 2; r <= lastRow; r++) {
    var jmCell = sheet.getRange(r, 3);
    var jsCell = sheet.getRange(r, 4);
    
    var jmVal = jmCell.getValue();
    var jsVal = jsCell.getValue();
    
    // Fix Jam Mulai
    if (jmVal instanceof Date) {
      jmCell.setValue(padTime_(jmVal.getHours() + ':' + jmVal.getMinutes()));
      fixed++;
    } else if (jmVal && String(jmVal).indexOf(':') === -1) {
      // Bukan format waktu, skip
    } else if (jmVal) {
      var fixedJm = padTime_(jmVal);
      if (fixedJm !== String(jmVal)) {
        jmCell.setValue(fixedJm);
        fixed++;
      }
    }
    
    // Fix Jam Selesai
    if (jsVal instanceof Date) {
      jsCell.setValue(padTime_(jsVal.getHours() + ':' + jsVal.getMinutes()));
      fixed++;
    } else if (jsVal) {
      var fixedJs = padTime_(jsVal);
      if (fixedJs !== String(jsVal)) {
        jsCell.setValue(fixedJs);
        fixed++;
      }
    }
  }
  
  CacheService.getScriptCache().remove('sheet_Data_Jadwal');
  return 'Selesai! ' + fixed + ' sel waktu difix.';
}

// --- GET ALL JADWAL ---
function getAllJadwal(filterRole, filterId) {
  var data = getSheetData('Data_Jadwal', false); // false = bypass cache
  var result = [];
  if (data && data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      var hari = String(data[i][0] || '');
      var kelas = String(data[i][1] || '');
      var guru = String(data[i][6] || '');
      var guruLogin = String(data[i][7] || '').trim();
      var filterIdStr = String(filterId || '').trim();
      
      if (filterRole === 'Ust' || filterRole === 'Ustz') {
        if (guruLogin != filterIdStr) continue;
      } else if (filterRole === 'Santri') {
        if (kelas !== filterIdStr) continue;
      }
      
      // Konversi waktu ke format HH:MM
      var jm = data[i][2];
      var js = data[i][3];
      jm = padTime_(jm instanceof Date ? jm.getHours() + ':' + jm.getMinutes() : jm);
      js = padTime_(js instanceof Date ? js.getHours() + ':' + js.getMinutes() : js);
      
      result.push({ no: i, hari: hari, kelas: kelas, jamMulai: jm, jamSelesai: js, mapel: String(data[i][4] || ''), ruangan: String(data[i][5] || ''), guru: guru, guruLogin: guruLogin });
    }
  }
  return result;
}
