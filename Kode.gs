// Mengambil data peserta, klub, hasil spin, dan jadwal pertandingan
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = getSyncData(ss);
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Menangani: Simpan Spin Picker, Update Skor Pertandingan, & Reset Skor
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload = JSON.parse(e.postData.contents);

    // KASUS 1: Update Skor Pertandingan
    if (payload.action === "updateScore") {
      var jadwalSheet = getSheetOrCreate(ss, "Jadwal", ["ID", "Leg", "Tim Home", "Skor Home", "Skor Away", "Tim Away"]);
      var lastRow = jadwalSheet.getLastRow();
      
      if (lastRow > 1) {
        var idList = jadwalSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < idList.length; i++) {
          if (Number(idList[i][0]) === Number(payload.matchId)) {
            var rowIndex = i + 2;
            jadwalSheet.getRange(rowIndex, 4).setValue(payload.homeScore !== null ? payload.homeScore : "");
            jadwalSheet.getRange(rowIndex, 5).setValue(payload.awayScore !== null ? payload.awayScore : "");
            
            SpreadsheetApp.flush(); // Memaksa Google Sheet langsung menyimpan perubahan saat itu juga
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Skor berhasil diperbarui" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // KASUS 2: Reset Semua Skor Liga
    if (payload.action === "resetScores") {
      var jadwalSheet = ss.getSheetByName("Jadwal");
      if (jadwalSheet && jadwalSheet.getLastRow() > 1) {
        jadwalSheet.getRange(2, 4, jadwalSheet.getLastRow() - 1, 2).clearContent();
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Skor di-reset" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // KASUS 3: Buat Ulang Jadwal Baru dari Tab Hasil
    if (payload.action === "regenerateSchedule") {
      var jadwalSheet = ss.getSheetByName("Jadwal");
      if (jadwalSheet) {
        ss.deleteSheet(jadwalSheet);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Jadwal di-reset ulang" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // KASUS 4: Simpan Hasil Spin Picker (Bawaan program sebelumnya)
    if (payload.user && payload.club) {
      var current = getSyncData(ss);
      if (!current.availableUsers.includes(payload.user)) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Peserta sudah digunakan!" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (!current.availableClubs.includes(payload.club)) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Klub sudah diambil orang lain!" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var hasilSheet = getSheetOrCreate(ss, "Hasil", ["Waktu", "Nama Peserta", "Klub Terpilih"]);
      var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
      hasilSheet.appendRow([timestamp, payload.user, payload.club]);

      return ContentService.createTextOutput(JSON.stringify({ status: "success", user: payload.user, club: payload.club }))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Fungsi pembantu membaca semua data dan jadwal liga
function getSyncData(ss) {
  var pesertaSheet = getSheetOrCreate(ss, "Peserta", ["Nama Peserta"]);
  var klubSheet = getSheetOrCreate(ss, "Klub", ["Nama Klub"]);
  var hasilSheet = getSheetOrCreate(ss, "Hasil", ["Waktu", "Nama Peserta", "Klub Terpilih"]);
  var jadwalSheet = getSheetOrCreate(ss, "Jadwal", ["ID", "Leg", "Tim Home", "Skor Home", "Skor Away", "Tim Away"]);

  // 1. Ambil Master Peserta & Klub
  var allPeserta = pesertaSheet.getLastRow() > 1 ? pesertaSheet.getRange(2, 1, pesertaSheet.getLastRow() - 1, 1).getValues().map(r => r[0].toString().trim()).filter(v => v !== "") : [];
  var allKlub = klubSheet.getLastRow() > 1 ? klubSheet.getRange(2, 1, klubSheet.getLastRow() - 1, 1).getValues().map(r => r[0].toString().trim()).filter(v => v !== "") : [];

  // 2. Ambil Hasil Undian
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

  var availableUsers = allPeserta.filter(u => !takenUsers.includes(u));
  var availableClubs = allKlub.filter(c => !takenClubs.includes(c));

  // 3. Ambil / Buat Jadwal Pertandingan 2 Leg
  var matches = [];
  var uniqueClubs = [...new Set(takenClubs)];

  if (jadwalSheet.getLastRow() > 1) {
    // Jika tab Jadwal sudah ada datanya, baca pertandingan & skornya
    var rawMatches = jadwalSheet.getRange(2, 1, jadwalSheet.getLastRow() - 1, 6).getValues();
    rawMatches.forEach(r => {
      matches.push({
        id: Number(r[0]),
        leg: Number(r[1]),
        home: r[2].toString(),
        homeScore: (r[3] !== "" && r[3] !== null) ? Number(r[3]) : null,
        awayScore: (r[4] !== "" && r[4] !== null) ? Number(r[4]) : null,
        away: r[5].toString()
      });
    });
  } else if (uniqueClubs.length >= 2) {
    // Jika tab Jadwal masih kosong, buat jadwal otomatis 2 Leg di Google Sheet
    var matchId = 1;
    var leg1 = [];
    for (var i = 0; i < uniqueClubs.length; i++) {
      for (var j = i + 1; j < uniqueClubs.length; j++) {
        leg1.push([matchId++, 1, uniqueClubs[i], "", "", uniqueClubs[j]]);
      }
    }
    leg1.sort(() => Math.random() - 0.5); // Acak Leg 1

    var leg2 = [];
    leg1.forEach(m => {
      leg2.push([matchId++, 2, m[5], "", "", m[2]]); // Balik Home-Away untuk Leg 2
    });
    leg2.sort(() => Math.random() - 0.5); // Acak Leg 2

    var allRows = leg1.concat(leg2);
    for (var k = 0; k < allRows.length; k++) {
      allRows[k][0] = k + 1;
      matches.push({
        id: k + 1,
        leg: allRows[k][1],
        home: allRows[k][2],
        homeScore: null,
        awayScore: null,
        away: allRows[k][5]
      });
    } 

    // Tulis ke Google Sheet tab Jadwal
    jadwalSheet.getRange(2, 1, allRows.length, 6).setValues(allRows);
  }

  var isLeagueStarted = matches.length > 0;

  return {
    availableUsers: availableUsers,
    availableClubs: availableClubs,
    results: results,
    clubsInLeague: uniqueClubs,
    matches: matches,
    isLeagueStarted: isLeagueStarted
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