// ========================================
// SANTRI.gs - CRUD Data Santri
// ========================================

// --- SANTRI CRUD ---
function getListSantri() { 
  var data = getSheetData('Data_Santri'); var list = []; 
  if(data) for(var i=1; i<data.length; i++) list.push({ nis: data[i][0], nama: data[i][1], kelas: data[i][8] || '' }); 
  return list; 
}

function getSantriDetail(nis) {
  var data = getSheetData('Data_Santri');
  if(!data) return null;
  for (var i = 1; i < data.length; i++) if (String(data[i][0]) === String(nis)) return { nis: data[i][0], nama: data[i][1], gender: data[i][2], tgl: data[i][3] ? Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "", wali: data[i][4], hp: data[i][5], foto: data[i][7], kelas: data[i][8] || '' };
  return null;
}

function addSantri(d) {
  if (d.userId && !validateAdminRole(d.userId)) {
    return { status: 'error', message: 'Hanya Admin yang bisa menambah data santri' };
  }
  var ss = SpreadsheetApp.openById(SHEET_ID); var sheet = ss.getSheetByName('Data_Santri');
  if(!sheet) return {status:'error', message:'Database belum disetup!'};
  var data = sheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) { if(String(data[i][0]) === String(d.nis)) return {status:'error', message:'NIS sudah terdaftar!'}; }
  var hashedPass = hashPassword('123456');
  sheet.appendRow([sanitizeInput(d.nis), sanitizeInput(d.nama), sanitizeInput(d.gender), sanitizeInput(d.tgl), sanitizeInput(d.wali), sanitizeInput(d.hp), hashedPass, sanitizeInput(d.foto), sanitizeInput(d.kelas || '')]); 
  CacheService.getScriptCache().remove('sheet_Data_Santri');
  return {status:'success'};
}

// --- BULK ADD SANTRI ---
function bulkAddSantri(rows, userId) {
  if (userId && !validateAdminRole(userId)) {
    return { status: 'error', message: 'Hanya Admin yang bisa import data santri' };
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Santri');
  if (!sheet) return { status: 'error', message: 'Sheet Data_Santri tidak ditemukan!' };

  // Ambil NIS yang sudah ada di sheet
  var existingData = sheet.getDataRange().getValues();
  var existingNIS = {};
  for (var i = 1; i < existingData.length; i++) {
    existingNIS[String(existingData[i][0])] = true;
  }

  var successCount = 0;
  var failedCount = 0;
  var failedList = [];
  var seenNIS = {}; // Cek duplikat dalam file yang sama
  var hashedPass = hashPassword('123456');

  for (var j = 0; j < rows.length; j++) {
    var row = rows[j];
    var nis = String(row.nis || '').trim();

    // Cek NIS kosong
    if (!nis) {
      failedCount++;
      failedList.push({ baris: j + 1, nis: '-', nama: row.nama || '-', alasan: 'NIS kosong' });
      continue;
    }

    // Cek duplikat di sheet yang sudah ada
    if (existingNIS[nis]) {
      failedCount++;
      failedList.push({ baris: j + 1, nis: nis, nama: row.nama || '-', alasan: 'NIS sudah ada di database' });
      continue;
    }

    // Cek duplikat dalam file yang sama
    if (seenNIS[nis]) {
      failedCount++;
      failedList.push({ baris: j + 1, nis: nis, nama: row.nama || '-', alasan: 'NIS duplikat dalam file' });
      continue;
    }

    seenNIS[nis] = true;

    sheet.appendRow([
      sanitizeInput(nis),
      sanitizeInput(row.nama || ''),
      sanitizeInput(row.gender || 'Laki-laki'),
      sanitizeInput(row.tgl || ''),
      sanitizeInput(row.wali || ''),
      sanitizeInput(row.hp || ''),
      hashedPass,
      '',
      sanitizeInput(row.kelas || '')
    ]);

    existingNIS[nis] = true;
    successCount++;
  }

  CacheService.getScriptCache().remove('sheet_Data_Santri');
  return { status: 'success', success: successCount, failed: failedCount, failedList: failedList };
}

function updateSantriData(d) {
  if (d.userId && !validateAdminRole(d.userId)) {
    return { status: 'error', message: 'Hanya Admin yang bisa mengedit data santri' };
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Santri');
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(d.oldNis)) {
      var r = i + 1;
      var oldNama = String(data[i][1]);
      var newNama = sanitizeInput(d.nama);
      var newNis = sanitizeInput(d.newNis);

      sheet.getRange(r, 1).setValue(newNis); sheet.getRange(r, 2).setValue(newNama);
      sheet.getRange(r, 3).setValue(sanitizeInput(d.gender)); sheet.getRange(r, 4).setValue(sanitizeInput(d.tgl));
      sheet.getRange(r, 5).setValue(sanitizeInput(d.wali)); sheet.getRange(r, 6).setValue(sanitizeInput(d.hp));
      if(d.password) sheet.getRange(r, 7).setValue(hashPassword(d.password));
      if(d.foto !== "") sheet.getRange(r, 8).setValue(d.foto);
      sheet.getRange(r, 9).setValue(sanitizeInput(d.kelas || ''));

      // Sync nama & NIS ke Data_Ziyadah dan Data_Murajaah
      var sheetsToSync = ['Data_Ziyadah', 'Data_Murajaah'];
      sheetsToSync.forEach(function(sheetName) {
        var targetSheet = ss.getSheetByName(sheetName);
        if (!targetSheet) return;
        var targetData = targetSheet.getDataRange().getValues();
        for (var j = 1; j < targetData.length; j++) {
          var rowNis = String(targetData[j][2]);
          var rowNama = String(targetData[j][1]);
          var nisChanged = (rowNis === String(d.oldNis));
          var namaChanged = (rowNama === oldNama && nisChanged);
          if (nisChanged) {
            targetSheet.getRange(j + 1, 3).setValue(newNis);
          }
          if (namaChanged) {
            targetSheet.getRange(j + 1, 2).setValue(newNama);
          }
        }
      });

      CacheService.getScriptCache().remove('sheet_Data_Santri');
      CacheService.getScriptCache().remove('sheet_Data_Ziyadah');
      CacheService.getScriptCache().remove('sheet_Data_Murajaah');
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'NIS tidak ditemukan' };
}

// --- SYNC DATA: Sejajarkan Nama & NIS dari Data_Santri ke Data_Ziyadah & Data_Murajaah ---
// Jalankan dari Apps Script editor jika ada perubahan manual di sheet Data_Santri
function syncSantriData() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var santriSheet = ss.getSheetByName('Data_Santri');
  if (!santriSheet) return 'Sheet Data_Santri tidak ditemukan';

  var santriData = santriSheet.getDataRange().getValues();

  // Bangun map: NIS lama → { nama, nisBaru }
  var santriMap = {};
  for (var i = 1; i < santriData.length; i++) {
    var nis = String(santriData[i][0]);
    var nama = String(santriData[i][1]);
    if (nis) santriMap[nis] = nama;
  }

  var totalUpdated = 0;
  var sheetsToSync = ['Data_Ziyadah', 'Data_Murajaah'];

  sheetsToSync.forEach(function(sheetName) {
    var targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) return;

    var targetData = targetSheet.getDataRange().getValues();
    for (var j = 1; j < targetData.length; j++) {
      var rowNis = String(targetData[j][2]);
      var rowNama = String(targetData[j][1]);

      // Cari NIS yang cocok di Data_Santri
      if (santriMap[rowNis] !== undefined) {
        var correctNama = santriMap[rowNis];

        // Update nama jika berbeda
        if (rowNama !== correctNama) {
          targetSheet.getRange(j + 1, 2).setValue(correctNama);
          totalUpdated++;
        }
      }
    }
  });

  // Clear cache
  CacheService.getScriptCache().remove('sheet_Data_Ziyadah');
  CacheService.getScriptCache().remove('sheet_Data_Murajaah');

  return 'Sync selesai! ' + totalUpdated + ' baris nama diperbarui di Data_Ziyadah & Data_Murajaah.';
}

// --- BULK DELETE SANTRI ---
function bulkDeleteSantri(nisList, userId) {
  if (userId && !validateAdminRole(userId)) {
    return { status: 'error', message: 'Hanya Admin yang bisa menghapus data santri' };
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Santri');
  if (!sheet) return { status: 'error', message: 'Sheet Data_Santri tidak ditemukan!' };

  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  // Cari baris yang NIS-nya ada di list
  for (var i = 1; i < data.length; i++) {
    var rowNis = String(data[i][0]);
    for (var j = 0; j < nisList.length; j++) {
      if (rowNis === String(nisList[j])) {
        rowsToDelete.push(i + 1); // +1 karena baris pertama adalah header
        break;
      }
    }
  }

  // Hapus dari belakang agar indeks tidak bergeser
  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var k = 0; k < rowsToDelete.length; k++) {
    sheet.deleteRow(rowsToDelete[k]);
  }

  // Hapus juga riwayat setoran dari Data_Ziyadah dan Data_Murajaah
  var sheetsToClean = ['Data_Ziyadah', 'Data_Murajaah'];
  sheetsToClean.forEach(function(sheetName) {
    var targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) return;
    var targetData = targetSheet.getDataRange().getValues();
    var targetRows = [];
    for (var m = 1; m < targetData.length; m++) {
      var rowNis = String(targetData[m][2]);
      for (var n = 0; n < nisList.length; n++) {
        if (rowNis === String(nisList[n])) {
          targetRows.push(m + 1);
          break;
        }
      }
    }
    targetRows.sort(function(a, b) { return b - a; });
    for (var p = 0; p < targetRows.length; p++) {
      targetSheet.deleteRow(targetRows[p]);
    }
  });

  CacheService.getScriptCache().remove('sheet_Data_Santri');
  CacheService.getScriptCache().remove('sheet_Data_Ziyadah');
  CacheService.getScriptCache().remove('sheet_Data_Murajaah');

  return { status: 'success', success: rowsToDelete.length };
}

function deleteSantriData(nis, userId) {
  if (userId && !validateAdminRole(userId)) {
    return { status: 'error', message: 'Hanya Admin yang bisa menghapus data santri' };
  }
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Santri');
  
  if (!sheet) return { status: 'error', message: 'Database tidak ditemukan' };
  
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(nis)) {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove('sheet_Data_Santri');
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Data santri tidak ditemukan atau sudah terhapus' };
}

function getSantriCompleteStats(nis) {
  var dataZ = getSheetData('Data_Ziyadah');
  var dataM = getSheetData('Data_Murajaah');
  
  var uniqueZiyadah = new Set();
  var totalMurajaah = 0;
  var activeJuzSet = new Set(); // Juz yang pernah disetor
  
  // Hitung Ziyadah & Juz Aktif
  if(dataZ) {
    for(var i=1; i<dataZ.length; i++) {
      if(String(dataZ[i][2]) === String(nis)) {
         var surat = dataZ[i][3];
         var start = parseInt(dataZ[i][5]);
         var end = parseInt(dataZ[i][6]);
         var juz = parseInt(dataZ[i][4]);
         
         if(!isNaN(juz)) activeJuzSet.add(juz);
         
         if(!isNaN(start) && !isNaN(end) && surat) {
            for(var k=start; k<=end; k++) uniqueZiyadah.add(surat.trim() + "_" + k);
         }
      }
    }
  }
  
  // Hitung Murajaah
  if(dataM) {
    for(var j=1; j<dataM.length; j++) {
      if(String(dataM[j][2]) === String(nis)) {
         var s = parseInt(dataM[j][5]);
         var e = parseInt(dataM[j][6]);
         if(!isNaN(s) && !isNaN(e)) totalMurajaah += (e - s + 1);
      }
    }
  }
  
  return {
    ziyadahCount: uniqueZiyadah.size,
    murajaahCount: totalMurajaah,
    activeJuz: Array.from(activeJuzSet) // Convert Set ke Array untuk dikirim ke JS
  };
}

function getSantriStats(nis) {
  var dataZ = getSheetData('Data_Ziyadah'); var dataM = getSheetData('Data_Murajaah');
  var tz = 0, tm = 0;
  if(dataZ) for(var i=1; i<dataZ.length; i++) if(String(dataZ[i][2]) === String(nis)) tz++;
  if(dataM) for(var j=1; j<dataM.length; j++) if(String(dataM[j][2]) === String(nis)) tm++;
  return { ziyadah: tz, murajaah: tm };
}

function getTotalCount(nis) { 
  var dataZ = getSheetData('Data_Ziyadah'); var dataM = getSheetData('Data_Murajaah'); var count = 0;
  if(dataZ) for(var i=1; i<dataZ.length; i++) if(String(dataZ[i][2]) == String(nis)) count++; 
  if(dataM) for(var j=1; j<dataM.length; j++) if(String(dataM[j][2]) == String(nis)) count++; 
  return count; 
}
