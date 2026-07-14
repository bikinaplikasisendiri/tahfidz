// ========================================
// MIGRATION.gs - Script Migrasi (Sekali Jalan)
// ========================================

// --- MIGRASI PASSWORD KE FORMAT BERSALT (JALANKAN SEKALI SAJA) ---
// Script ini meng-hash ulang password yang sudah ada ke format salt:hash
// HATI-HATI: Jalankan script ini HANYA SEKALI
function migratePasswordsToSalted() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var migratedCount = 0;
  var skippedCount = 0;
  var errorCount = 0;
  
  // 1. Migrate Data_Santri (kolom Password = index 6)
  var santriSheet = ss.getSheetByName('Data_Santri');
  if (santriSheet) {
    var santriData = santriSheet.getDataRange().getValues();
    for (var i = 1; i < santriData.length; i++) {
      var currentPass = String(santriData[i][6]);
      
      // Skip jika sudah format baru (salt:hash)
      if (isSaltedHash_(currentPass)) {
        skippedCount++;
        continue;
      }
      
      // Skip jika kosong
      if (!currentPass || currentPass.trim() === '') {
        skippedCount++;
        continue;
      }
      
      try {
        // Re-hash dengan salt (menggunakan hash lama sebagai "password" untuk di-hash ulang)
        // Catatan: Kita tidak bisa reverse hash, jadi kita simpan hash lama dengan salt baru
        var salt = generateSalt_();
        var hash = hashPasswordWithSalt_(currentPass, salt); // Hash dari hash lama
        var newStored = salt + ':' + hash;
        santriSheet.getRange(i + 1, 7).setValue(newStored);
        migratedCount++;
      } catch(e) {
        errorCount++;
      }
    }
  }
  
  // 2. Migrate Data_User (kolom Password = index 6)
  var userSheet = ss.getSheetByName('Data_User');
  if (userSheet) {
    var userData = userSheet.getDataRange().getValues();
    for (var i = 1; i < userData.length; i++) {
      var currentPass = String(userData[i][6]);
      
      // Skip jika sudah format baru
      if (isSaltedHash_(currentPass)) {
        skippedCount++;
        continue;
      }
      
      // Skip jika kosong
      if (!currentPass || currentPass.trim() === '') {
        skippedCount++;
        continue;
      }
      
      try {
        var salt = generateSalt_();
        var hash = hashPasswordWithSalt_(currentPass, salt);
        var newStored = salt + ':' + hash;
        userSheet.getRange(i + 1, 7).setValue(newStored);
        migratedCount++;
      } catch(e) {
        errorCount++;
      }
    }
  }
  
  return {
    status: 'success',
    message: 'Migrasi selesai! ' + migratedCount + ' password di-hash ulang, ' + skippedCount + ' dilewati, ' + errorCount + ' error.'
  };
}

// --- CEK STATUS MIGRASI ---
// Fungsi ini membantu memeriksa apakah masih ada password format lama
function checkMigrationStatus() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var oldFormatCount = 0;
  var newFormatCount = 0;
  var emptyCount = 0;
  
  // Cek Data_Santri
  var santriSheet = ss.getSheetByName('Data_Santri');
  if (santriSheet) {
    var santriData = santriSheet.getDataRange().getValues();
    for (var i = 1; i < santriData.length; i++) {
      var pass = String(santriData[i][6]);
      if (!pass || pass.trim() === '') {
        emptyCount++;
      } else if (isSaltedHash_(pass)) {
        newFormatCount++;
      } else {
        oldFormatCount++;
      }
    }
  }
  
  // Cek Data_User
  var userSheet = ss.getSheetByName('Data_User');
  if (userSheet) {
    var userData = userSheet.getDataRange().getValues();
    for (var i = 1; i < userData.length; i++) {
      var pass = String(userData[i][6]);
      if (!pass || pass.trim() === '') {
        emptyCount++;
      } else if (isSaltedHash_(pass)) {
        newFormatCount++;
      } else {
        oldFormatCount++;
      }
    }
  }
  
  return {
    formatBaru: newFormatCount,
    formatLama: oldFormatCount,
    kosong: emptyCount,
    perluMigrasi: oldFormatCount > 0
  };
}
