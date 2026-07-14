// ========================================
// NOTIFIKASI.gs - Sistem Notifikasi
// ========================================

// --- NOTIFIKASI SYSTEM ---
function getNotifications(userId) {
  var data = getSheetData('Data_Notifikasi');
  if(!data || data.length <= 1) return [];
  
  // Cek Role User
  var userRole = 'Santri'; 
  var uData = getSheetData('Data_User');
  if(uData) {
    for(var k=1; k<uData.length; k++) {
      if(String(uData[k][2]) == String(userId)) { userRole = uData[k][3]; break; }
    }
  }
  var isAdmin = (userRole === 'Admin' || userRole === 'Ust');
  
  var result = [];
  // Ambil dari bawah (terbaru)
  for(var i=data.length-1; i>=1; i--) {
    var row = data[i];
    var targetId = String(row[1]);
    
    // Tampilkan jika milik user ATAU user adalah Admin
    if(targetId === String(userId) || targetId === 'ALL' || isAdmin) {
      var timeDiff = new Date() - new Date(row[0]);
      var timeStr = "Baru saja";
      var min = Math.floor(timeDiff/60000);
      var hour = Math.floor(min/60);
      
      if(hour >= 24) timeStr = Math.floor(hour/24) + " hari lalu";
      else if(hour > 0) timeStr = hour + " jam lalu";
      else if(min > 0) timeStr = min + " menit lalu";
      
      result.push({
        id: i, time: timeStr, title: row[2], msg: row[3], type: row[4], read: row[5] === 'read'
      });
      if(result.length >= 10) break; // Batasi 10 notif
    }
  }
  return result;
}

// --- TANDAI DIBACA ---
function markAllRead(userId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Notifikasi');
  if(!sheet) return;
  
  var data = sheet.getDataRange().getValues();
  var ranges = [];
  
  for(var i=1; i<data.length; i++) {
    var targetId = String(data[i][1]);
    var status = data[i][5];
    
    if(status !== 'read' && targetId === String(userId)) {
        ranges.push('F' + (i+1));
    }
  }
  
  if (ranges.length > 0) {
    sheet.getRangeList(ranges).setValue('read');
  }
}

// --- HAPUS SEMUA NOTIFIKASI ---
function deleteAllNotifications(userId) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_Notifikasi');
  if(!sheet) return { status: 'success' };
  
  var data = sheet.getDataRange().getValues();
  
  var userRole = 'Santri';
  var uData = getSheetData('Data_User');
  if(uData) {
    for(var k=1; k<uData.length; k++) {
      if(String(uData[k][2]) == String(userId)) { userRole = uData[k][3]; break; }
    }
  }
  var isAdmin = (userRole === 'Admin' || userRole === 'Ust' || userRole === 'Ustadz');
  
  var deletedCount = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var targetId = String(row[1]);
    
    if (isAdmin) {
      // Admin bisa hapus semua notifikasi
      sheet.deleteRow(i + 1);
      deletedCount++;
    } else {
      // Santri hanya hapus milik sendiri
      if (targetId === String(userId)) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
  }
  
  if (isAdmin) {
    return { status: 'success', message: deletedCount + ' notifikasi berhasil dihapus.' };
  } else {
    return { status: 'success', message: 'Notifikasi berhasil dihapus.' };
  }
}
