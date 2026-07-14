// ========================================
// REPORTS.gs - Laporan Bulanan
// ========================================

// --- LAPORAN PER SANTRI ---
function getSantriMonthlyReport(nis, bulanAwal, bulanAkhir, tahun) {
  var dataZ = getSheetData('Data_Ziyadah');
  var dataM = getSheetData('Data_Murajaah');
  var records = [];
  var totalZiyadah = 0;
  var totalMurajaah = 0;
  var totalSetor = 0;

  function isInRange(tgl) {
    var m = tgl.getMonth() + 1;
    var y = tgl.getFullYear();
    if (y != tahun) return false;
    return m >= bulanAwal && m <= bulanAkhir;
  }

  if (dataZ && dataZ.length > 1) {
    for (var i = 1; i < dataZ.length; i++) {
      if (String(dataZ[i][2]) !== String(nis)) continue;
      var tgl = parseDate_(dataZ[i][0]);
      if (!isInRange(tgl)) continue;
      var start = parseInt(dataZ[i][5]);
      var end = parseInt(dataZ[i][6]);
      var ayatCount = (!isNaN(start) && !isNaN(end)) ? (end - start + 1) : 0;
      records.push({
        tanggal: tgl, tipe: 'Ziyadah', surat: dataZ[i][3] || '',
        ayat: (start && end) ? start + '-' + end : '-',
        nilai: dataZ[i][7] || '-', catatan: dataZ[i][8] || ''
      });
      totalZiyadah += ayatCount;
      totalSetor++;
    }
  }

  if (dataM && dataM.length > 1) {
    for (var i = 1; i < dataM.length; i++) {
      if (String(dataM[i][2]) !== String(nis)) continue;
      var tgl = parseDate_(dataM[i][0]);
      if (!isInRange(tgl)) continue;
      var start = parseInt(dataM[i][5]);
      var end = parseInt(dataM[i][6]);
      var ayatCount = (!isNaN(start) && !isNaN(end)) ? (end - start + 1) : 0;
      records.push({
        tanggal: tgl, tipe: 'Murajaah', surat: dataM[i][3] || '',
        ayat: (start && end) ? start + '-' + end : '-',
        nilai: dataM[i][7] || '-', catatan: dataM[i][8] || ''
      });
      totalMurajaah += ayatCount;
      totalSetor++;
    }
  }

  records.sort(function(a, b) { return b.tanggal - a.tanggal; });

  var formatted = records.map(function(r, idx) {
    return {
      no: idx + 1,
      tanggal: Utilities.formatDate(r.tanggal, "Asia/Jakarta", "dd MMM yyyy HH:mm"),
      tipe: r.tipe, surat: r.surat, ayat: r.ayat, nilai: r.nilai, catatan: r.catatan
    };
  });

  return {
    records: formatted,
    summary: { ziyadah: totalZiyadah, murajaah: totalMurajaah, totalSetor: totalSetor }
  };
}

// --- LAPORAN BULANAN ---
function getMonthlyReport(bulanAwal, bulanAkhir, tahun, kelas) {
  var dataZ = getSheetData('Data_Ziyadah');
  var dataM = getSheetData('Data_Murajaah');
  var dataS = getSheetData('Data_Santri');
  
  // Build santri map
  var santriMap = {};
  if (dataS && dataS.length > 1) {
    for (var i = 1; i < dataS.length; i++) {
      santriMap[String(dataS[i][0])] = {
        nama: dataS[i][1],
        kelas: dataS[i][8] || ''
      };
    }
  }
  
  // Filter by kelas
  var filteredNIS = {};
  if (kelas && kelas !== '') {
    for (var nis in santriMap) {
      if (santriMap[nis].kelas === kelas) filteredNIS[nis] = true;
    }
  } else {
    for (var nis in santriMap) filteredNIS[nis] = true;
  }
  
  var stats = {};
  for (var nis in filteredNIS) {
    stats[nis] = {
      nama: santriMap[nis].nama,
      kelas: santriMap[nis].kelas,
      ziyadahAyat: 0, murajaahAyat: 0, setorZiyadah: 0, setorMurajaah: 0
    };
  }
  
  function isInRange(tgl) {
    var m = tgl.getMonth() + 1;
    var y = tgl.getFullYear();
    if (y != tahun) return false;
    return m >= bulanAwal && m <= bulanAkhir;
  }
  
  if (dataZ && dataZ.length > 1) {
    for (var i = 1; i < dataZ.length; i++) {
      var rowNIS = String(dataZ[i][2]);
      if (!stats[rowNIS]) continue;
      var tgl = parseDate_(dataZ[i][0]);
      if (!isInRange(tgl)) continue;
      var start = parseInt(dataZ[i][5]);
      var end = parseInt(dataZ[i][6]);
      var ayatCount = (!isNaN(start) && !isNaN(end)) ? (end - start + 1) : 0;
      stats[rowNIS].ziyadahAyat += ayatCount;
      stats[rowNIS].setorZiyadah++;
    }
  }
  
  if (dataM && dataM.length > 1) {
    for (var i = 1; i < dataM.length; i++) {
      var rowNIS = String(dataM[i][2]);
      if (!stats[rowNIS]) continue;
      var tgl = parseDate_(dataM[i][0]);
      if (!isInRange(tgl)) continue;
      var start = parseInt(dataM[i][5]);
      var end = parseInt(dataM[i][6]);
      var ayatCount = (!isNaN(start) && !isNaN(end)) ? (end - start + 1) : 0;
      stats[rowNIS].murajaahAyat += ayatCount;
      stats[rowNIS].setorMurajaah++;
    }
  }
  
  var result = [];
  var totalZ = 0, totalM = 0, totalSetorZ = 0, totalSetorM = 0, totalSantri = 0;
  
  for (var nis in stats) {
    var s = stats[nis];
    result.push({
      nis: nis, nama: s.nama, kelas: s.kelas,
      ziyadahAyat: s.ziyadahAyat, murajaahAyat: s.murajaahAyat,
      setorZiyadah: s.setorZiyadah, setorMurajaah: s.setorMurajaah
    });
    totalZ += s.ziyadahAyat; totalM += s.murajaahAyat;
    totalSetorZ += s.setorZiyadah; totalSetorM += s.setorMurajaah;
    if (s.ziyadahAyat > 0 || s.murajaahAyat > 0) totalSantri++;
  }
  
  result.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  
  return {
    data: result,
    summary: {
      totalSantri: totalSantri, totalZiyadah: totalZ, totalMurajaah: totalM,
      totalSetor: totalSetorZ + totalSetorM
    }
  };
}
