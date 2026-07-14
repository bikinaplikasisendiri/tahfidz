// ========================================
// HAFALAN.gs - Setoran, Riwayat, Statistik
// ========================================

// --- RIWAYAT PER JUZ (Untuk Santri) ---
function getRiwayatByJuz(nis, juz, filter, page) {
  var pageSize = 10;
  var startIdx = (page || 0) * pageSize;
  var result = [];
  
  // Ambil data Ziyadah
  if (filter === 'all' || filter === 'Ziyadah') {
    var dataZ = getSheetData('Data_Ziyadah');
    if (dataZ && dataZ.length > 1) {
      for (var i = 1; i < dataZ.length; i++) {
        if (String(dataZ[i][2]) === String(nis) && String(dataZ[i][4]) === String(juz)) {
          result.push({
            tipe: 'Ziyadah',
            tanggal: dataZ[i][0],
            nama: dataZ[i][1],
            nis: dataZ[i][2],
            surat: dataZ[i][3],
            juz: dataZ[i][4],
            ayatMulai: dataZ[i][5],
            ayatSelesai: dataZ[i][6],
            nilai: dataZ[i][7] || '-',
            catatan: dataZ[i][8] || ''
          });
        }
      }
    }
  }
  
  // Ambil data Murajaah
  if (filter === 'all' || filter === 'Murajaah') {
    var dataM = getSheetData('Data_Murajaah');
    if (dataM && dataM.length > 1) {
      for (var i = 1; i < dataM.length; i++) {
        if (String(dataM[i][2]) === String(nis) && String(dataM[i][4]) === String(juz)) {
          result.push({
            tipe: 'Murajaah',
            tanggal: dataM[i][0],
            nama: dataM[i][1],
            nis: dataM[i][2],
            surat: dataM[i][3],
            juz: dataM[i][4],
            ayatMulai: dataM[i][5],
            ayatSelesai: dataM[i][6],
            lanjut: dataM[i][7] || '-',
            catatan: dataM[i][8] || ''
          });
        }
      }
    }
  }
  
  // Sort by tanggal (terbaru dulu)
  result.sort(function(a, b) {
    var dateA = parseDate_(a.tanggal);
    var dateB = parseDate_(b.tanggal);
    return dateB - dateA;
  });
  
  var total = result.length;
  var paged = result.slice(startIdx, startIdx + pageSize);
  
  return {
    data: paged,
    total: total,
    hasMore: (startIdx + pageSize) < total
  };
}

// --- STATISTIK PER JUZ ---
function getJuzStats(nis, juz) {
  var dataZ = getSheetData('Data_Ziyadah');
  var dataM = getSheetData('Data_Murajaah');
  var ziyadahCount = 0;
  var murajaahCount = 0;
  var ziyadahAyatSet = new Set();
  
  // Hitung Ziyadah unik per Juz
  if (dataZ && dataZ.length > 1) {
    for (var i = 1; i < dataZ.length; i++) {
      if (String(dataZ[i][2]) === String(nis) && String(dataZ[i][4]) === String(juz)) {
        var surat = dataZ[i][3];
        var start = parseInt(dataZ[i][5]);
        var end = parseInt(dataZ[i][6]);
        if (!isNaN(start) && !isNaN(end) && surat) {
          for (var k = start; k <= end; k++) {
            ziyadahAyatSet.add(surat.trim() + "_" + k);
          }
        }
      }
    }
  }
  ziyadahCount = ziyadahAyatSet.size;
  
  // Hitung Murajaah
  if (dataM && dataM.length > 1) {
    for (var i = 1; i < dataM.length; i++) {
      if (String(dataM[i][2]) === String(nis) && String(dataM[i][4]) === String(juz)) {
        var start = parseInt(dataM[i][5]);
        var end = parseInt(dataM[i][6]);
        if (!isNaN(start) && !isNaN(end)) {
          murajaahCount += (end - start + 1);
        }
      }
    }
  }
  
  return {
    juz: juz,
    ziyadah: ziyadahCount,
    murajaah: murajaahCount
  };
}

// --- SIMPAN SETORAN ---
function simpanSetoran(form) { 
  var ss = SpreadsheetApp.openById(SHEET_ID); 
  var targetSheetName = (form.kategori == 'Ziyadah') ? 'Data_Ziyadah' : 'Data_Murajaah';
  var sheet = ss.getSheetByName(targetSheetName);
  var notifSheet = ss.getSheetByName('Data_Notifikasi');
  
  if (!sheet) return "Error: Sheet " + targetSheetName + " tidak ditemukan!";
  
  var ayatStart = parseInt(form.ayatMulai) || 0;
  var ayatEnd = parseInt(form.ayatSelesai) || 0;
  if (ayatStart <= 0 || ayatEnd <= 0 || ayatStart > ayatEnd) {
    return "Error: Nomor ayat tidak valid!";
  }
  
  var now = new Date();
  var tglFormatted = "'" + Utilities.formatDate(now, "Asia/Jakarta", "dd-MM-yyyy HH:mm:ss");
  
  sheet.appendRow([
    tglFormatted, 
    sanitizeInput(form.namaSantri), 
    sanitizeInput(form.nisSantri), 
    sanitizeInput(form.surat), 
    sanitizeInput(form.juz), 
    form.ayatMulai, 
    form.ayatSelesai, 
    (form.kategori == 'Ziyadah' ? sanitizeInput(form.nilai) : sanitizeInput(form.lanjut)), 
    sanitizeInput(form.catatan), 
    sanitizeInput(form.namaUst)
  ]);
  
  if (notifSheet) {
    var title = "Setoran " + form.kategori;
    var msg = sanitizeInput(form.surat) + " (Ayat " + form.ayatMulai + "-" + form.ayatSelesai + ")";
    notifSheet.appendRow([new Date(), sanitizeInput(form.nisSantri), title, msg, sanitizeInput(form.kategori), 'unread']);
  }
  
  // Clear cache
  var cache = CacheService.getScriptCache();
  cache.remove('sheet_' + targetSheetName);
  cache.remove('sheet_Data_Notifikasi');
  
  return "Data Berhasil Disimpan!"; 
}

// --- UPDATE SETORAN ---
function updateSetoranData(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = (d.kategori === 'Ziyadah') ? ss.getSheetByName('Data_Ziyadah') : ss.getSheetByName('Data_Murajaah');
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
  var data = sheet.getDataRange().getValues();
  var clean = function(s) { return String(s).replace(/'/g, "").trim(); };
  for (var i = 1; i < data.length; i++) {
    if (clean(data[i][0]) === clean(d.oldDate) && String(data[i][2]) === String(d.nis)) {
      var r = i+1;
      sheet.getRange(r, 4).setValue(d.surat); sheet.getRange(r, 5).setValue(d.juz);
      sheet.getRange(r, 6).setValue(d.ayatMulai); sheet.getRange(r, 7).setValue(d.ayatSelesai);
      sheet.getRange(r, 8).setValue(d.nilai); sheet.getRange(r, 9).setValue(d.catatan);
      CacheService.getScriptCache().remove('sheet_' + sheet.getName());
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Data tidak ditemukan.' };
}

// --- HAPUS SETORAN ---
function deleteSetoranData(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = (d.kategori === 'Ziyadah') ? ss.getSheetByName('Data_Ziyadah') : ss.getSheetByName('Data_Murajaah');
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
  var data = sheet.getDataRange().getValues();
  var clean = function(s) { return String(s).replace(/'/g, "").trim(); };
  for (var i = 1; i < data.length; i++) {
    if (clean(data[i][0]) === clean(d.date) && String(data[i][2]) === String(d.nis)) {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove('sheet_' + sheet.getName());
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Gagal menghapus data' };
}

// --- RIWAYAT HAFALAN ---
function getRiwayatHafalan(userId, role) {
  var result = [];
  var userRole = String(role).toLowerCase();
  var isAdminOrUst = (userRole === 'admin' || userRole === 'ust' || userRole === 'ustadz');

  function parseIndoDate(dateStr) {
    if (!dateStr) return new Date(0);
    if (dateStr instanceof Date) return dateStr;
    try {
      var cleanStr = String(dateStr).replace(/'/g, "").trim();
      var parts = cleanStr.split(' ');
      var d = parts[0].split('-');
      if (parts.length < 2) return new Date(d[2], d[1]-1, d[0]);
      var t = parts[1].split(':');
      return new Date(d[2], d[1]-1, d[0], t[0], t[1], t[2]);
    } catch (e) { return new Date(0); }
  }

  function processData(sheetName, typeLabel) {
    var rawData = getSheetData(sheetName);
    if (!rawData || rawData.length <= 1) return; 
    for (var i = 1; i < rawData.length; i++) {
      try {
        var row = rawData[i];
        if (isAdminOrUst || String(row[2]) === String(userId)) {
          var dateObj = parseIndoDate(row[0]);
          result.push({
            tipe: typeLabel, timestamp: dateObj.getTime(), rawDate: row[0], 
            tanggal: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd MMM HH:mm"),
            nama: row[1], nis: row[2], surat: row[3], juz: row[4], ayatStart: row[5], ayatEnd: row[6], ayat: row[5] + "-" + row[6], nilai_info: row[7] || '-', catatan: row[8] || ''
          });
        }
      } catch (e) { }
    }
  }
  processData('Data_Ziyadah', 'Ziyadah');
  processData('Data_Murajaah', 'Murajaah');
  result.sort(function(a,b){ return b.timestamp - a.timestamp; });
  return result.slice(0, 20);
}

// --- LEADERBOARD ---
function getLeaderboard() {
  var data = getSheetData('Data_Ziyadah');
  var stats = {};

  if (data && data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var nis = String(row[2]); // NIS sebagai key unik
      var nama = row[1];
      var surat = row[3];
      var start = parseInt(row[5]);
      var end = parseInt(row[6]);

      if (!stats[nis]) {
        stats[nis] = { nama: nama, set: new Set() };
      }

      if (!isNaN(start) && !isNaN(end) && surat) {
        for (var k = start; k <= end; k++) {
          var uniqueKey = surat.trim() + "_" + k;
          stats[nis].set.add(uniqueKey);
        }
      }
    }
  }

  var ranking = [];
  for (var nis in stats) {
    ranking.push({ nama: stats[nis].nama, total: stats[nis].set.size });
  }

  ranking.sort(function(a, b){ return b.total - a.total; });
  return ranking.slice(0, 50);
}
