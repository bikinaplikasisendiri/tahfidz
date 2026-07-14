// ========================================
// AUTH.gs - Autentikasi, Keamanan, Hash
// ========================================

// Ambil SHEET_ID dari PropertiesService (atau set pertama kali)
function getSheetId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (!id) {
    id = SpreadsheetApp.getActiveSpreadsheet().getId();
    props.setProperty('SHEET_ID', id);
  }
  return id;
}
var SHEET_ID = getSheetId_();

// --- RATE LIMITING LOGIN ---
// Batas: 5 percobaan gagal per 15 menit per username
// Setelah itu, akun terkunci selama 15 menit

var RATE_LIMIT_MAX_ATTEMPTS = 5;
var RATE_LIMIT_WINDOW_MS = 15 * 60 * 1000; // 15 menit dalam milidetik

// Ambil data percobaan login dari PropertiesService
function getLoginAttempts_(username) {
  var props = PropertiesService.getScriptProperties();
  var key = 'login_attempts_' + username;
  var data = props.getProperty(key);
  
  if (!data) {
    return { count: 0, firstAttempt: 0, lockedUntil: 0 };
  }
  
  try {
    return JSON.parse(data);
  } catch(e) {
    return { count: 0, firstAttempt: 0, lockedUntil: 0 };
  }
}

// Simpan data percobaan login
function saveLoginAttempts_(username, attempts) {
  var props = PropertiesService.getScriptProperties();
  var key = 'login_attempts_' + username;
  props.setProperty(key, JSON.stringify(attempts));
}

// Cek apakah username terkunci
function isAccountLocked_(username) {
  var attempts = getLoginAttempts_(username);
  var now = Date.now();
  
  // Cek apakah masih dalam masa lock
  if (attempts.lockedUntil > 0 && attempts.lockedUntil > now) {
    return {
      locked: true,
      remainingSeconds: Math.ceil((attempts.lockedUntil - now) / 1000)
    };
  }
  
  return { locked: false, remainingSeconds: 0 };
}

// Catat percobaan login gagal
function recordFailedLogin_(username) {
  var attempts = getLoginAttempts_(username);
  var now = Date.now();
  
  // Jika sudah lewat window waktu, reset counter
  if (now - attempts.firstAttempt > RATE_LIMIT_WINDOW_MS) {
    attempts = { count: 1, firstAttempt: now, lockedUntil: 0 };
  } else {
    attempts.count++;
  }
  
  // Jika melebihi batas, kunci akun
  if (attempts.count >= RATE_LIMIT_MAX_ATTEMPTS) {
    attempts.lockedUntil = now + RATE_LIMIT_WINDOW_MS;
  }
  
  saveLoginAttempts_(username, attempts);
  return attempts;
}

// Reset percobaan login (saat login berhasil)
function resetLoginAttempts_(username) {
  var props = PropertiesService.getScriptProperties();
  var key = 'login_attempts_' + username;
  props.deleteProperty(key);
}

// Fungsi untuk membersihkan data rate limiting yang sudah expired (opsional, bisa dijadwalkan)
function cleanupExpiredRateLimits() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  var now = Date.now();
  var cleaned = 0;
  
  keys.forEach(function(key) {
    if (key.indexOf('login_attempts_') === 0) {
      try {
        var data = JSON.parse(props.getProperty(key));
        // Hapus jika sudah lewat window + lock time
        if (now - data.firstAttempt > RATE_LIMIT_WINDOW_MS * 2) {
          props.deleteProperty(key);
          cleaned++;
        }
      } catch(e) {
        props.deleteProperty(key);
        cleaned++;
      }
    }
  });
  
  return cleaned;
}

// --- PASSWORD HASHING (SHA-256 + Salt) ---
// Format penyimpanan: salt:hash (contoh: abc123:def456...)

// Generate random salt (32 karakter hex)
function generateSalt_() {
  var salt = '';
  var chars = '0123456789abcdef';
  for (var i = 0; i < 32; i++) {
    salt += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return salt;
}

// Hash password dengan salt
function hashPasswordWithSalt_(password, salt) {
  // Iterasi 10000x untuk memperkuat hash (key stretching)
  var hash = salt + password;
  for (var i = 0; i < 10000; i++) {
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, hash);
    hash = '';
    for (var j = 0; j < digest.length; j++) {
      hash += ('0' + (digest[j] & 0xFF).toString(16)).slice(-2);
    }
  }
  return hash;
}

// Hash password baru dengan salt baru (untuk registrasi/simpan)
function hashPassword(password) {
  var salt = generateSalt_();
  var hash = hashPasswordWithSalt_(password, salt);
  return salt + ':' + hash;
}

// Verifikasi password (mendukung format plain text, SHA-256, dan salted hash)
function verifyPassword(password, storedValue) {
  if (!storedValue) return false;
  storedValue = String(storedValue);
  
  // 1. Cek format baru: salt:hash (panjang total > 64)
  if (storedValue.indexOf(':') !== -1) {
    var parts = storedValue.split(':');
    if (parts.length === 2 && parts[0].length === 32 && parts[1].length === 64) {
      var salt = parts[0];
      var expectedHash = parts[1];
      var computedHash = hashPasswordWithSalt_(password, salt);
      return computedHash === expectedHash;
    }
  }
  
  // 2. Cek format lama: SHA-256 hash (64 karakter hex)
  if (storedValue.length === 64 && /^[a-f0-9]+$/.test(storedValue)) {
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
    var hash = '';
    for (var i = 0; i < digest.length; i++) {
      hash += ('0' + (digest[i] & 0xFF).toString(16)).slice(-2);
    }
    return hash === storedValue;
  }
  
  // 3. Plain text: bandingkan langsung (untuk backward compatibility)
  // Catatan: Setelah login berhasil, password akan otomatis di-hash
  return password === storedValue;
}

// Cek apakah password sudah menggunakan format baru (dengan salt)
function isSaltedHash_(storedValue) {
  if (!storedValue) return false;
  storedValue = String(storedValue);
  if (storedValue.indexOf(':') !== -1) {
    var parts = storedValue.split(':');
    return parts.length === 2 && parts[0].length === 32 && parts[1].length === 64;
  }
  return false;
}

// Cek apakah password adalah plain text (bukan hash)
function isPlainTextPassword_(storedValue) {
  if (!storedValue) return false;
  storedValue = String(storedValue);
  // Hash selalu 64 karakter hex, plain text biasanya lebih pendek dan ada karakter non-hex
  if (storedValue.length === 64 && /^[a-f0-9]+$/.test(storedValue)) {
    return false; // Ini SHA-256 hash
  }
  if (isSaltedHash_(storedValue)) {
    return false; // Ini salted hash
  }
  return true; // Plain text
}

// --- HASH PASSWORD PLAIN TEXT YANG SUDAH ADA DI SPREADSHEET ---
// Jalankan fungsi ini SETELAH membuat akun admin di spreadsheet
// untuk meng-hash password plain text menjadi format yang benar
function hashAllPlainPasswords() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hashedCount = 0;
  
  // Hash Data_User
  var userSheet = ss.getSheetByName('Data_User');
  if (userSheet) {
    var data = userSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var pass = String(data[i][6]);
      if (isPlainTextPassword_(pass) && pass !== '') {
        var newHash = hashPassword(pass);
        userSheet.getRange(i + 1, 7).setValue(newHash);
        hashedCount++;
      }
    }
  }
  
  // Hash Data_Santri
  var santriSheet = ss.getSheetByName('Data_Santri');
  if (santriSheet) {
    var dataS = santriSheet.getDataRange().getValues();
    for (var j = 1; j < dataS.length; j++) {
      var passS = String(dataS[j][6]);
      if (isPlainTextPassword_(passS) && passS !== '') {
        var newHashS = hashPassword(passS);
        santriSheet.getRange(j + 1, 7).setValue(newHashS);
        hashedCount++;
      }
    }
  }
  
  return hashedCount + ' password berhasil di-hash!';
}

// --- INPUT SANITIZATION ---
function sanitizeInput(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// --- ROLE VALIDATION HELPER ---
function validateAdminRole(userId) {
  var userRole = 'Santri';
  var uData = getSheetData('Data_User');
  if (uData) {
    for (var k = 1; k < uData.length; k++) {
      if (String(uData[k][2]) == String(userId)) { 
        userRole = uData[k][3]; 
        break; 
      }
    }
  }
  // Cek juga di Data_Santri
  if (userRole === 'Santri') {
    var sData = getSheetData('Data_Santri');
    if (sData) {
      for (var j = 1; j < sData.length; j++) {
        if (String(sData[j][0]) == String(userId)) {
          return false; // Santri tidak boleh CRUD
        }
      }
    }
  }
  return (userRole === 'Admin' || userRole === 'Ust' || userRole === 'Ustadz');
}

// --- LOGIN & DASHBOARD ---
function loginUser(username, password) {
  // RATE LIMITING: Cek apakah akun terkunci
  var lockStatus = isAccountLocked_(username);
  if (lockStatus.locked) {
    var minutes = Math.ceil(lockStatus.remainingSeconds / 60);
    return {
      status: 'failed',
      message: 'Akun terkunci karena terlalu banyak percobaan gagal. Coba lagi dalam ' + minutes + ' menit.'
    };
  }
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var loginFound = false;
  
  // Cek Data_User
  var userSheet = ss.getSheetByName('Data_User');
  if (userSheet) {
    var data = userSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]) == username) {
        loginFound = true;
        var storedPass = String(data[i][6]);
        
        if (verifyPassword(password, storedPass)) {
          // Login berhasil: reset rate limit
          resetLoginAttempts_(username);
          // Auto-upgrade: jika password masih format lama, hash ulang dengan salt
          if (!isSaltedHash_(storedPass)) {
            var newHash = hashPassword(password);
            userSheet.getRange(i + 1, 7).setValue(newHash);
          }
          return { status: 'success', id: data[i][2], nama: data[i][0], role: data[i][3], loginTime: Date.now() };
        }
      }
    }
  }
  
  // Cek Data_Santri
  var santriSheet = ss.getSheetByName('Data_Santri');
  if (santriSheet) {
    var dataS = santriSheet.getDataRange().getValues();
    for (var j = 1; j < dataS.length; j++) {
      if (String(dataS[j][0]) == username) {
        loginFound = true;
        var storedPassS = String(dataS[j][6]);
        
        if (verifyPassword(password, storedPassS)) {
          // Login berhasil: reset rate limit
          resetLoginAttempts_(username);
          // Auto-upgrade: jika password masih format lama, hash ulang dengan salt
          if (!isSaltedHash_(storedPassS)) {
            var newHash = hashPassword(password);
            santriSheet.getRange(j + 1, 7).setValue(newHash);
          }
          return { status: 'success', id: dataS[j][0], nama: dataS[j][1], role: 'Santri', kelas: dataS[j][8] || '', loginTime: Date.now() };
        }
      }
    }
  }
  
  // Login gagal: catat percobaan
  var attempts = recordFailedLogin_(username);
  var remaining = RATE_LIMIT_MAX_ATTEMPTS - attempts.count;
  
  if (attempts.count >= RATE_LIMIT_MAX_ATTEMPTS) {
    return {
      status: 'failed',
      message: 'Akun terkunci karena 5 percobaan gagal. Coba lagi dalam 15 menit.'
    };
  } else if (loginFound) {
    return {
      status: 'failed',
      message: 'Password salah. Sisa percobaan: ' + remaining
    };
  } else {
    // Username tidak ditemukan tetap tetap catat untuk mencegah enumeration
    recordFailedLogin_(username);
    return {
      status: 'failed',
      message: 'Username atau Password salah.'
    };
  }
}

function addUser(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Data_User');
  if (!sheet) return { status: 'error', message: 'Sheet Data_User tidak ditemukan!' };
  
  // Cek duplikat login
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(d.login)) {
      return { status: 'error', message: 'Username sudah digunakan!' };
    }
  }
  
  var hashedPass = hashPassword(d.password);
  sheet.appendRow([
    sanitizeInput(d.nama),
    sanitizeInput(d.hp || ''),
    sanitizeInput(d.login),
    d.role || 'Ust',
    sanitizeInput(d.mapel || ''),
    d.status || 'Aktif',
    hashedPass,
    ''
  ]);
  
  return { status: 'success' };
}

function updateUserProfile(d) {
  var ss = SpreadsheetApp.openById(SHEET_ID); var sheetName = (d.role === 'Santri') ? 'Data_Santri' : 'Data_User';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
  var data = sheet.getDataRange().getValues(); 
  for (var i = 1; i < data.length; i++) { 
    var searchKey = (d.role === 'Santri') ? data[i][0] : data[i][2]; 
    if (String(searchKey) == String(d.id)) { 
      if (d.role === 'Santri') { 
        sheet.getRange(i + 1, 2).setValue(d.nama); 
        if(d.pass) sheet.getRange(i + 1, 7).setValue(hashPassword(d.pass)); 
        if(d.hp) sheet.getRange(i+1, 6).setValue(d.hp); 
      } else { 
        // Data_User: Nama=0, HP=1, Login=2, Role=3, Mapel=4, Status=5, Password=6, Foto=7
        sheet.getRange(i + 1, 1).setValue(d.nama); 
        if(d.pass) sheet.getRange(i + 1, 7).setValue(hashPassword(d.pass)); 
        if(d.hp) sheet.getRange(i+1, 2).setValue(d.hp); 
      } 
      return { status: 'success' }; 
    } 
  } 
  return { status: 'error', message: 'User ID tidak ditemukan' };
}
