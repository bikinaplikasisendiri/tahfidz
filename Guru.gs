// ========================================
// GURU.gs - CRUD Data Guru
// ========================================

// --- DATA GURU CRUD ---
function getListGuru() {
  var data = getSheetData('Data_User');
  var list = [];
  if (data && data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      var role = String(data[i][3]);
      if (role === 'Ust' || role === 'Ustz') {
        list.push({
          login: data[i][2],
          nama: data[i][0],
          hp: data[i][1] || '',
          role: role,
          mapel: data[i][4] || '',
          status: data[i][5] || 'Aktif'
        });
      }
    }
  }
  return list;
}

function getGuruDetail(login) {
  var data = getSheetData('Data_User');
  if (!data) return null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(login) && String(data[i][3]) !== 'Santri') {
      return {
        login: data[i][2],
        nama: data[i][0],
        hp: data[i][1] || '',
        role: data[i][3] || 'Guru',
        mapel: data[i][4] || '',
        status: data[i][5] || 'Aktif',
        foto: data[i][7] || ''
      };
    }
  }
  return null;
}

function addGuru(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_User');
  if (!sheet) return { status: 'error', message: 'Sheet Data_User tidak ditemukan!' };
  
  var data = sheet.getDataRange().getValues();
  var login = d.login || d.id;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(login)) {
      return { status: 'error', message: 'Username sudah digunakan!' };
    }
  }
  
  var hashedPass = d.password ? hashPassword(d.password) : '';
  sheet.appendRow([
    sanitizeInput(d.nama),
    sanitizeInput(d.hp || ''),
    sanitizeInput(login),
    d.role || 'Ust',
    sanitizeInput(d.mapel || ''),
    d.status || 'Aktif',
    hashedPass,
    d.foto || ''
  ]);
  
  CacheService.getScriptCache().remove('sheet_Data_User');
  return { status: 'success' };
}

function updateGuru(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_User');
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan!' };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(d.login)) {
      var r = i + 1;
      sheet.getRange(r, 1).setValue(sanitizeInput(d.nama));
      sheet.getRange(r, 2).setValue(sanitizeInput(d.hp));
      sheet.getRange(r, 4).setValue(d.role || 'Ust');
      sheet.getRange(r, 5).setValue(sanitizeInput(d.mapel));
      sheet.getRange(r, 6).setValue(d.status || 'Aktif');
      if (d.password) {
        sheet.getRange(r, 7).setValue(hashPassword(d.password));
      }
      if (d.foto !== undefined && d.foto !== '') {
        sheet.getRange(r, 8).setValue(d.foto);
      }
      CacheService.getScriptCache().remove('sheet_Data_User');
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Guru tidak ditemukan' };
}

function deleteGuru(login) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_User');
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(login) && String(data[i][3]) !== 'Santri') {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove('sheet_Data_User');
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Guru tidak ditemukan' };
}
