// ========================================
// UTILS.gs - Utilitas Umum, Cache, Helper
// ========================================

// --- FUNGSI AMBIL DATA AMAN DENGAN CACHE ---
function getSheetData(name, useCache) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'sheet_' + name;
  
  // Cek cache dulu
  if (useCache !== false) {
    var cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch(e) {}
    }
  }
  
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var s = ss.getSheetByName(name);
    if (!s) {
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().trim() === name.trim()) {
          s = sheets[i];
          break;
        }
      }
      if (!s) return [];
    }
    var data = s.getDataRange().getValues();
    
    // Simpan ke cache (5 menit)
    if (data && data.length > 0) {
      cache.put(cacheKey, JSON.stringify(data), 300);
    }
    
    return data || [];
  } catch (e) {
    return [];
  }
}

function clearAllCache() {
  var cache = CacheService.getScriptCache();
  var sheets = ['Data_Santri', 'Data_User', 'Data_Ziyadah', 'Data_Murajaah', 'Data_Quran', 'Data_Notifikasi', 'Data_Jadwal'];
  sheets.forEach(function(name) {
    cache.remove('sheet_' + name);
  });
  return 'Cache dibersihkan!';
}

// Helper parse tanggal
function parseDate_(dateStr) {
  if (!dateStr) return new Date(0);
  if (dateStr instanceof Date) return dateStr;
  try {
    var cleanStr = String(dateStr).replace(/'/g, "").trim();
    var parts = cleanStr.split(' ');
    if (parts.length < 2) return new Date(0);
    var d = parts[0].split('-');
    var t = parts[1].split(':');
    return new Date(d[2], d[1]-1, d[0], t[0], t[1], t[2]);
  } catch (e) { return new Date(0); }
}

function padTime_(t) {
  if (!t) return '';
  var str = String(t).trim();
  if (str.indexOf(':') === -1) return str;
  var parts = str.split(':');
  var h = (parts[0] || '0').padStart(2, '0');
  var m = (parts[1] || '0').padStart(2, '0');
  return h + ':' + m;
}

function getKelasList() {
  var data = getSheetData('Data_Santri');
  var kelasSet = {};
  if(data) for(var i=1; i<data.length; i++) {
    var k = String(data[i][8] || '').trim();
    if(k) kelasSet[k] = true;
  }
  return Object.keys(kelasSet).sort();
}

function getListSurat() { 
  var data = getSheetData('Data_Quran'); var list = []; 
  if(data) for(var i=1; i<data.length; i++) list.push(data[i][1]); 
  return list; 
}

// --- SETUP DATABASE (JALANKAN SEKALI SAJA) ---
function doSetup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var db = [
    {n:"Data_Santri", h:["NIS","Nama","Jenis_Kelamin","Tanggal_Lahir","Wali","HP","Password","Foto","Kelas"]},
    {n:"Data_User", h:["Nama_Lengkap","HP","Login","Role","Mata_Pelajaran","Status","Password","Foto"]},
    {n:"Data_Ziyadah", h:["Tanggal","Nama","NIS","Surat","Juz","Ayat_Mulai","Ayat_Selesai","Nilai","Catatan","Muhafidz"]},
    {n:"Data_Murajaah", h:["Tanggal","Nama","NIS","Surat","Juz","Ayat_Mulai","Ayat_Selesai","Lanjut","Catatan","Muhafidz"]},
    {n:"Data_Quran", h:["No","Surat","Jumlah_Ayat"]},
    {n:"Data_Notifikasi", h:["Timestamp","Target_ID","Judul","Pesan","Tipe","Status"]},
    {n:"Data_Jadwal", h:["Hari","Kelas","Jam_Mulai","Jam_Selesai","Mata_Pelajaran","Ruangan","Guru","Guru_Login"]}
  ];

  db.forEach(function(d) {
    var s = ss.getSheetByName(d.n);
    if(!s) {
      s=ss.insertSheet(d.n);
      s.appendRow(d.h);
      s.setFrozenRows(1);
      // Set format teks untuk kolom waktu di Data_Jadwal
      if (d.n === 'Data_Jadwal') {
        s.getRange(1, 3, 100, 2).setNumberFormat('@');
      }
    } 
  });
}
