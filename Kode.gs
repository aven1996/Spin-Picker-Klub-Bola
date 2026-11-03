// Mengambil sisa klub, peserta, dan hasil terkini
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = getSyncData(ss);
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Menyimpan hasil putaran dengan sistem Kunci (Lock)
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Tunggu maksimal 10 detik jika ada pengguna lain yang sedang menyimpan bersamaan
    lock.waitLock(10000);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload = JSON.parse(e.postData.contents);
    var targetUser = payload.user;
    var targetClub = payload.club;
    
    var current = getSyncData(ss);
    
    // Validasi anti-bentrok: pastikan belum diambil orang lain
    if (!current.availableUsers.includes(targetUser)) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "Nama peserta sudah digunakan oleh orang lain!" 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (!current.availableClubs.includes(targetClub)) {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "Klub ini baru saja diambil orang lain! Roda akan disegarkan." 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Tulis ke tab 'Hasil'
    var hasilSheet = getSheetOrCreate(ss, "Hasil", ["Waktu", "Nama Peserta", "Klub Terpilih"]);
    var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    hasilSheet.appendRow([timestamp, targetUser, targetClub]);
    
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      user: targetUser, 
      club: targetClub 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Fungsi pembantu menghitung sisa klub & peserta
function getSyncData(ss) {
  var pesertaSheet = getSheetOrCreate(ss, "Peserta", ["Nama Peserta"]);
  var klubSheet = getSheetOrCreate(ss, "Klub", ["Nama Klub"]);
  var hasilSheet = getSheetOrCreate(ss, "Hasil", ["Waktu", "Nama Peserta", "Klub Terpilih"]);
  
  // Baca master peserta
  var allPeserta = [];
  if (pesertaSheet.getLastRow() > 1) {
    allPeserta = pesertaSheet.getRange(2, 1, pesertaSheet.getLastRow() - 1, 1).getValues()
      .map(r => r[0].toString().trim()).filter(v => v !== "");
  }
  
  // Baca master klub
  var allKlub = [];
  if (klubSheet.getLastRow() > 1) {
    allKlub = klubSheet.getRange(2, 1, klubSheet.getLastRow() - 1, 1).getValues()
      .map(r => r[0].toString().trim()).filter(v => v !== "");
  }
  
  // Baca data yang sudah terambil
  var takenUsers = [];
  var takenClubs = [];
  var results = [];
  
  if (hasilSheet.getLastRow() > 1) {
    var rawHasil = hasilSheet.getRange(2, 1, hasilSheet.getLastRow() - 1, 3).getValues();
    rawHasil.forEach(r => {
      if (r[1] && r[2]) {
        takenUsers.push(r[1].toString().trim());
        takenClubs.push(r[2].toString().trim());
        results.push({ time: r[0], user: r[1], club: r[2] });
      }
    });
  }
  
  // Sisa klub dan peserta = Master dikurangi yang sudah ada di tab 'Hasil'
  var availableUsers = allPeserta.filter(u => !takenUsers.includes(u));
  var availableClubs = allKlub.filter(c => !takenClubs.includes(c));
  
  return {
    availableUsers: availableUsers,
    availableClubs: availableClubs,
    results: results
  };
}

function getSheetOrCreate(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) sheet.appendRow(headers);
  }
  return sheet;
}
