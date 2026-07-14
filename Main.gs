// ========================================
// Index.gs - Entry Point & Dashboard Stats
// ========================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Hifdzul Qur\'an App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// --- STATISTIK DASHBOARD ---
function getDashboardStats(role, id) {
  var dataZ = getSheetData('Data_Ziyadah');
  var dataM = getSheetData('Data_Murajaah');
  var dataS = getSheetData('Data_Santri');
  
  var totalAyatZiyadah = 0;
  var totalAyatMurajaah = 0; // Murajaah biarkan overlap (karena mengulang)
  var countSantriSetorToday = 0;
  
  var uniqueJuzMap = {};
  var santriActiveSet = {};
  
  // Set untuk menghitung Ziyadah Unik
  var ziyadahUniqueSet = new Set(); 

  var totalSantriRegistered = dataS ? Math.max(0, dataS.length - 1) : 0;
  
  var now = new Date();
  var days = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  var months = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agust','Sep','Okt','Nov','Des'];
  var dateString = days[now.getDay()] + ", " + now.getDate() + " " + months[now.getMonth()] + " " + now.getFullYear();

  function countAyat(row) {
    var start = parseInt(row[5]), end = parseInt(row[6]);   
    return (!isNaN(start) && !isNaN(end)) ? Math.max(0, (end - start) + 1) : 0;
  }
  
  function isToday(dateString) {
    if(!dateString) return false;
    var check;
    if(dateString instanceof Date) check = dateString;
    else { try { var c=String(dateString).replace(/'/g,"").trim().split(' ')[0].split('-'); check=new Date(c[2],c[1]-1,c[0]); } catch(e){return false;} }
    return check.getDate()===now.getDate() && check.getMonth()===now.getMonth() && check.getFullYear()===now.getFullYear();
  }

  // --- LOGIKA SANTRI ---
  if (role === 'Santri') {
    // 1. Hitung Ziyadah Unik
    if(dataZ) {
        for(var i=1; i<dataZ.length; i++) {
            // Cek apakah data milik user ini
            if(String(dataZ[i][2]) === String(id)) { 
                var s = dataZ[i][3]; // Surat
                var j = dataZ[i][4]; // Juz
                var st = parseInt(dataZ[i][5]);
                var en = parseInt(dataZ[i][6]);
                
                if (j) uniqueJuzMap[j] = true;

                // Masukkan ke Set Unik
                if(!isNaN(st) && !isNaN(en) && s) {
                    for(var k=st; k<=en; k++) {
                        ziyadahUniqueSet.add(s.trim() + "_" + k);
                    }
                }
            }
        }
    }
    totalAyatZiyadah = ziyadahUniqueSet.size; // Ambil jumlah unik

    // 2. Hitung Murajaah (Tetap Akumulasi Biasa)
    if(dataM) for(var j=1; j<dataM.length; j++) if(String(dataM[j][2]) === String(id)) totalAyatMurajaah += countAyat(dataM[j]);
    
    return { 
      role: 'Santri', dateStr: dateString, mainNum: totalAyatZiyadah, mainUnit: "Ayat", secNum: Object.keys(uniqueJuzMap).length, secUnit: "Juz", 
      percent: Math.min((totalAyatZiyadah/6236)*100, 100), ziyadahCount: totalAyatZiyadah, murajaahCount: totalAyatMurajaah, juzCount: Object.keys(uniqueJuzMap).length 
    };

  } else {
    // --- LOGIKA ADMIN (Tidak perlu unik per santri untuk total global harian, cukup akumulasi) ---
    // Tapi jika ingin Total Hafalan Seluruh Santri Unik, logikanya kompleks.
    // Untuk Dashboard Admin "Real Time Ayat" biasanya adalah "Ayat yang disetor HARI INI".
    // Jadi kita gunakan logika "Setor Hari Ini" saja (tanpa cek unik history).

    var ayatToday = 0;

    if(dataZ) for(var k=1; k<dataZ.length; k++) {
        if(isToday(dataZ[k][0])) {
            ayatToday += countAyat(dataZ[k]);
            if(dataZ[k][4]) uniqueJuzMap[dataZ[k][4]] = true;
            santriActiveSet[dataZ[k][2]] = true;
        }
    }
    if(dataM) for(var l=1; l<dataM.length; l++) {
        if(isToday(dataM[l][0])) {
            ayatToday += countAyat(dataM[l]);
            santriActiveSet[dataM[l][2]] = true;
        }
    }
    
    // Tapi biasanya dashboard admin memisahkan Ziyadah Total vs Harian. 
    // Di kode lama, mainNum = totalAyatZiyadah + totalAyatMurajaah (Global).
    // Mari kita kembalikan ke total akumulasi biasa untuk Admin agar performa cepat, 
    // karena Admin melihat "Aktivitas", bukan "Hafalan Unik Global".
    
    // Kita hitung total Ziyadah (Akumulasi Biasa) untuk Admin agar konsisten dengan "Aktivitas"
    // (Kecuali Anda mau Total Unik juga untuk Admin, tapi itu berat di loading)
    
    var totalZiyadahAdmin = 0;
    var totalMurajaahAdmin = 0;
    if(dataZ) for(var k=1; k<dataZ.length; k++) totalZiyadahAdmin += countAyat(dataZ[k]);
    if(dataM) for(var l=1; l<dataM.length; l++) totalMurajaahAdmin += countAyat(dataM[l]);
    
    // Hitung ulang yang aktif hari ini
    countSantriSetorToday = Object.keys(santriActiveSet).length;
    var percentage = totalSantriRegistered > 0 ? Math.round((countSantriSetorToday / totalSantriRegistered) * 100) : 0;

    return { 
      role: 'Admin', dateStr: dateString, mainNum: ayatToday, mainUnit: "Ayat (Hari Ini)", secNum: totalSantriRegistered, secUnit: "Total Santri", 
      setorCount: countSantriSetorToday, percent: percentage, ziyadahCount: totalZiyadahAdmin, murajaahCount: totalMurajaahAdmin, juzCount: 30 
    };
  }
}
