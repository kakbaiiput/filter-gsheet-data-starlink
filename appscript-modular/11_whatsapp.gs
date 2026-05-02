// ========================================
// 📱 8. WHATSAPP REMINDER FUNCTIONS
// ========================================

/**
 * 🔄 UPDATED: Main WhatsApp function dengan kolom mapping baru
 */
function kirimWhatsApp(skipEmail) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Hapus filter jika ada
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  // Delay 3 detik
  Utilities.sleep(3000);

  var data = sheet.getDataRange().getValues();
  var headerRow = data[0];
  var waColumnIndex = -1;
  
  // ✅ UPDATED: Column indices sesuai mapping baru
  var tandaColumnIndex = COLUMNS.LABEL_JATUH_TEMPO; // U: "Jatuh Tempo"
  var statusColumnIndex = COLUMNS.STATUS; // O: Status "Ready"
  var linkColumnIndex = COLUMNS.TOMBOL_WHATSAPP; // Y: Link WA
  var prosesColumnIndex = COLUMNS.LABEL_PROSES; // V: "PROSES"
  var segaraColumnIndex = COLUMNS.LABEL_SEGERA; // W: "SEGERA"
  var observasiColumnIndex = COLUMNS.LABEL_OBSERVASI; // X: "OBSERVASI"
  
  // ✅ FIX: Gunakan object untuk logs agar bisa dimodifikasi by reference
  var logs = {
    clientProsesLog: "CLIENT YANG PROSES HARI INI:\n\n",
    clientReminderH3Log: "CLIENT YANG HARUS REMINDER H-3 HARI INI:\n\n",
    clientReminderH2Log: "CLIENT YANG HARUS REMINDER H-2 HARI INI:\n\n",
    clientReminderH1Log: "CLIENT YANG HARUS REMINDER H-1 HARI INI:\n\n",
    clientJatuhTempoLog: "CLIENT YANG JATUH TEMPO DAN BELUM:\n\n",
    clientWarningLog: "CLIENT WARNING - HARI H (AKAN MATI HARI INI):\n\n",
    errorLog: ""
  };

  // Cari kolom untuk Tombol WhatsApp
  for (var c = 0; c < headerRow.length; c++) {
    if (headerRow[c].toString().trim().toLowerCase() === "tombol whatsapp") {
      waColumnIndex = c + 1;
      break;
    }
  }

  // Jika belum ada, buat kolom "Tombol WhatsApp"
  if (waColumnIndex === -1) {
    waColumnIndex = sheet.getLastColumn() + 1;
    sheet.getRange(1, waColumnIndex).setValue("Tombol WhatsApp").setFontWeight("bold").setHorizontalAlignment("center");
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  Logger.log("Memulai perulangan baris data...");

  // Generate link WA untuk H-3 dan HARI H
  for (var i = 1; i < data.length; i++) {
    var nama = data[i][COLUMNS.NAMA];  // A: Nama
    var email = data[i][COLUMNS.LOGIN_STARLINK]; // C: Login Starlink
    var nomorWA = data[i][COLUMNS.NOMOR_CS]; // G: Nomor WA
    var kitNumber = data[i][COLUMNS.KIT_NUMBER]; // I: KIT Number ✅ NEW
    var serialNumber = data[i][COLUMNS.SERIAL_NUMBER]; // J: Serial Number ✅ NEW
    var jatuhTempoRaw = data[i][COLUMNS.JATUH_TEMPO]; // K: Jatuh Tempo ✅ MOVED
    var paket = data[i][COLUMNS.PAKET]; // P: Paket ✅ MOVED
    var status = data[i][statusColumnIndex]; // O: Status ✅ MOVED
    var tipePelanggan = data[i][COLUMNS.TIPE_PELANGGAN]; // S: Tipe Pelanggan ✅ NEW
    
    // ✅ UPDATED: Cell references dengan kolom baru
    var tandaCell = sheet.getRange(i + 1, tandaColumnIndex + 1); // U: Tanda "Jatuh Tempo"
    var cell = sheet.getRange(i + 1, waColumnIndex);
    var prosesCell = sheet.getRange(i + 1, prosesColumnIndex + 1); // V: PROSES
    var segaraCell = sheet.getRange(i + 1, segaraColumnIndex + 1); // W: SEGERA
    var observasiCell = sheet.getRange(i + 1, observasiColumnIndex + 1); // X: OBSERVASI

    // Validasi hanya untuk field wajib (nama + tanggal)
    if (!nama || !jatuhTempoRaw) {
      var logWajib = `❌ Baris ke-${i + 1} dilewati karena data wajib tidak lengkap:`;
      if (!nama) logWajib += "\n   - Nama: [Kosong]";
      if (!jatuhTempoRaw) logWajib += "\n   - Jatuh Tempo: [Kosong]";
      
      Logger.log(logWajib);
      logs.errorLog += logWajib + "\n";
      
      // Clear semua status karena data wajib tidak lengkap
      tandaCell.setValue("");
      cell.setValue("").setFormula("");
      prosesCell.setValue("");
      segaraCell.setValue("");
      observasiCell.setValue("");
      continue;
    }

    // Validasi tanggal
    var jatuhTempoDate = new Date(jatuhTempoRaw);
    if (isNaN(jatuhTempoDate.getTime())) {
      var logTanggal = `❌ Baris ke-${i + 1} dilewati karena format tanggal tidak valid.`;
      Logger.log(logTanggal);
      logs.errorLog += logTanggal + "\n";
      
      // Clear semua status karena tanggal invalid
      tandaCell.setValue("");
      cell.setValue("").setFormula("");
      prosesCell.setValue("");
      segaraCell.setValue("");
      observasiCell.setValue("");
      continue;
    }

    // Validasi untuk field optional (tetap catat tapi tidak skip)
    var logOptional = `⚠️ Baris ke-${i + 1} data tidak lengkap (tetap diproses):`;
    var adaOptionalKosong = false;
    if (!email) { logOptional += "\n   - Email: [Kosong]"; adaOptionalKosong = true; }
    if (!nomorWA) { logOptional += "\n   - Nomor WA: [Kosong]"; adaOptionalKosong = true; }
    if (!kitNumber) { logOptional += "\n   - KIT Number: [Kosong]"; adaOptionalKosong = true; }
    if (!paket) { logOptional += "\n   - Paket: [Kosong]"; adaOptionalKosong = true; }

    if (adaOptionalKosong) {
      Logger.log(logOptional);
      logs.errorLog += logOptional + "\n";
    }

    jatuhTempoDate.setHours(0, 0, 0, 0);
    var selisihHari = Math.floor((jatuhTempoDate - today) / (1000 * 60 * 60 * 24));

    // Cek apakah status adalah "Pending"
    if (status.toString().trim().toLowerCase() === "pending") {
      observasiCell.setValue("OBSERVASI")
                   .setFontWeight("bold")
                   .setFontSize(13)
                   .setHorizontalAlignment("center")
                   .setVerticalAlignment("middle");
    } else {
      observasiCell.setValue("");
    }

    // Label "JATUH TEMPO"
    if (selisihHari === 0) {
      tandaCell.setValue("JATUH TEMPO").setFontWeight("bold").setFontColor("#1F497D").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
    } else {
      tandaCell.setValue("");
    }

    // Status "Belum"
    if (status.toString().trim().toLowerCase() === "belum") {
      if (tandaCell.getValue() === "JATUH TEMPO") {
        logs.clientJatuhTempoLog += `${i + 1}. - ${nama} | ${email || 'N/A'} | ${nomorWA || 'N/A'} | ${kitNumber || 'N/A'}\n\n`;
      }
    }

    // Label "PROSES" jika tanggal jatuh tempo + 1 hari
    var jatuhTempoPlusOne = new Date(jatuhTempoDate);
    jatuhTempoPlusOne.setDate(jatuhTempoPlusOne.getDate() + 1);
    if (jatuhTempoPlusOne.getTime() === today.getTime()) {
      prosesCell.setValue("PROSES").setFontWeight("bold").setFontColor("#1F497D").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
      logs.clientProsesLog += `${i + 1}. - ${nama} | ${email || 'N/A'} | ${nomorWA || 'N/A'} | ${kitNumber || 'N/A'}\n\n`;
    } else {
      prosesCell.setValue("");
    }

    // Label "SEGERA" jika status Ready atau "Planned" dan sudah lewat jatuh tempo
    if ((status.toString().trim().toLowerCase() === "ready" || status.toString().trim().toLowerCase() === "planned") && jatuhTempoDate < today) {
      segaraCell.setValue("SEGERA")
                .setFontWeight("bold")
                .setFontColor("#003366")
                .setFontSize(13)
                .setHorizontalAlignment("center")
                .setVerticalAlignment("middle");
    } else {
      segaraCell.setValue("");
    }

    // Helper function untuk cek status yang memenuhi syarat
    function isEligibleStatus(status) {
      var statusLower = status.toString().trim().toLowerCase();
      return statusLower === "belum" || 
             statusLower === "active" || 
             statusLower === "pending reminder" || 
             statusLower === ""; // kosong
    }

    // ✅ FIX: Process reminder logic dengan logs object yang bisa dimodifikasi
    processReminderLogic(selisihHari, isEligibleStatus(status), {
      nama: nama,
      nomorWA: nomorWA,
      kitNumber: kitNumber,
      serialNumber: serialNumber,
      paket: paket,
      jatuhTempoDate: jatuhTempoDate,
      status: status,
      tipePelanggan: tipePelanggan,
      cell: cell,
      i: i
    }, logs); // ✅ Pass logs object by reference
  }

  // Send consolidated email
  // Gunakan === true agar trigger event object tidak ikut disable email
  if (skipEmail !== true) sendConsolidatedEmail(logs);

  // Reactivate filter
  reactivateFilter(sheet);
}

/**
 * ✅ FIXED: Process reminder logic (now properly modifies logs object)
 */
function processReminderLogic(selisihHari, isEligible, clientData, logs) {
  var { nama, nomorWA, kitNumber, serialNumber, paket, jatuhTempoDate, status, tipePelanggan, cell, i } = clientData;
  
  if (selisihHari === 3) {
    // H-3: Reminder H-3
    if (isEligible) {
      if (nomorWA) {
        var waLinkH3 = generateWhatsAppLink(nama, nomorWA, kitNumber, serialNumber, paket, jatuhTempoDate, 'H-3', tipePelanggan);
        setReminderCell(cell, "📅 Reminder H-3", waLinkH3, "#2563eb");
        logs.clientReminderH3Log += `${i + 1}. - ${nama} (H-3, Status: ${status}) | ${waLinkH3}\n\n`;
      } else {
        setReminderCell(cell, "📅 Reminder H-3", null, "#2563eb");
        logs.clientReminderH3Log += `${i + 1}. - ${nama} (H-3, Status: ${status}) | [NO WA]\n\n`;
      }
    } else {
      cell.setValue("").setFormula("");
    }
  } else if (selisihHari === 2) {
    // H-2: Reminder H-2
    if (isEligible) {
      if (nomorWA) {
        var waLinkH2 = generateWhatsAppLink(nama, nomorWA, kitNumber, serialNumber, paket, jatuhTempoDate, 'H-2', tipePelanggan);
        setReminderCell(cell, "⚠️ Reminder H-2", waLinkH2, "#f59e0b");
        logs.clientReminderH2Log += `${i + 1}. - ${nama} (H-2, Status: ${status}) | ${waLinkH2}\n\n`;
      } else {
        setReminderCell(cell, "⚠️ Reminder H-2", null, "#f59e0b");
        logs.clientReminderH2Log += `${i + 1}. - ${nama} (H-2, Status: ${status}) | [NO WA]\n\n`;
      }
    } else {
      cell.setValue("").setFormula("");
    }
  } else if (selisihHari === 1) {
    // H-1: Reminder H-1
    if (isEligible) {
      if (nomorWA) {
        var waLinkH1 = generateWhatsAppLink(nama, nomorWA, kitNumber, serialNumber, paket, jatuhTempoDate, 'H-1', tipePelanggan);
        setReminderCell(cell, "🚨 Reminder H-1", waLinkH1, "#dc2626");
        logs.clientReminderH1Log += `${i + 1}. - ${nama} (H-1, Status: ${status}) | ${waLinkH1}\n\n`;
      } else {
        setReminderCell(cell, "🚨 Reminder H-1", null, "#dc2626");
        logs.clientReminderH1Log += `${i + 1}. - ${nama} (H-1, Status: ${status}) | [NO WA]\n\n`;
      }
    } else {
      cell.setValue("").setFormula("");
    }
  } else if (selisihHari === 0) {
    // HARI H: Warning Reminder
    if (isEligible) {
      if (nomorWA) {
        var waLinkHariH = generateWhatsAppLink(nama, nomorWA, kitNumber, serialNumber, paket, jatuhTempoDate, 'WARNING', tipePelanggan);
        setReminderCell(cell, "💀 Warning Reminder", waLinkHariH, "#7c2d12");
        logs.clientWarningLog += `${i + 1}. - ${nama} (Warning, Status: ${status}) | ${waLinkHariH}\n\n`;
      } else {
        setReminderCell(cell, "💀 Warning Reminder", null, "#7c2d12");
        logs.clientWarningLog += `${i + 1}. - ${nama} (Warning, Status: ${status}) | [NO WA]\n\n`;
      }
    } else {
      cell.setValue("").setFormula("");
    }
  } else {
    // Clear cell untuk hari lain
    cell.setValue("").setFormula("");
  }
}

/**
 * ✅ FIXED: Send consolidated email (now uses logs object)
 */
function sendConsolidatedEmail(logs) {
  var emailBody = `${logs.clientProsesLog}\n${logs.clientReminderH3Log}\n${logs.clientReminderH2Log}\n${logs.clientReminderH1Log}\n${logs.clientWarningLog}\n${logs.clientJatuhTempoLog}\n\nBerikut adalah daftar data yang tidak lengkap:\n\n${logs.errorLog}`;
  
  if (emailBody.trim() !== "") {
    Logger.log("Mengirim email dengan body:\n" + emailBody);
    MailApp.sendEmail({
      to: "baiiput@gmail.com, octobiztech@gmail.com",
      subject: "Client Reminder H-3, H-2, H-1, Warning (Hari H), dan Jatuh Tempo",
      body: emailBody
    });
    Logger.log("Email telah dikirim dengan informasi lengkap H-3, H-2, H-1, dan Warning Hari H.");
  } else {
    Logger.log("emailBody kosong. Tidak ada email yang dikirim.");
  }
}

/**
 * ✅ NEW: Generate WhatsApp link dengan KIT Number, Serial Number, dan Tipe Pelanggan
 * ✅ UPDATE: Auto-rotasi 7 variasi pesan berdasarkan nomor WA agar aman dari filter spam WhatsApp
 */
function generateWhatsAppLink(nama, nomorWA, kitNumber, serialNumber, paket, jatuhTempoDate, type, tipePelanggan) {
  var kitInfo = '';

  // Handle KIT Number dan Serial Number
  if (kitNumber || serialNumber) {
    var kitLines = [];
    var kitArray = kitNumber ? (kitNumber.includes('\n') ? kitNumber.split('\n') : [kitNumber]) : ['Data KIT tidak tersedia'];
    var serialArray = serialNumber ? (serialNumber.includes('\n') ? serialNumber.split('\n') : [serialNumber]) : ['Data Serial tidak tersedia'];
    var paketArray = paket ? (paket.includes('\n') ? paket.split('\n') : [paket]) : ['Data paket tidak tersedia'];

    var maxLength = Math.max(kitArray.length, serialArray.length, paketArray.length);

    for (var j = 0; j < maxLength; j++) {
      var kitItem = kitArray[j] || kitArray[0] || 'N/A';
      var serialItem = serialArray[j] || serialArray[0] || 'N/A';
      var paketItem = paketArray[j] || paketArray[0] || 'N/A';
      kitLines.push(`   ✅ KIT: ${kitItem} (${paketItem})`);
    }

    kitInfo = kitLines.join('\n');
  }

  // Pilih variasi secara acak tiap kali link dibuat → tiap client selalu dapat pesan berbeda
  var idx = Math.floor(Math.random() * 7);

  var messages = {
    'H-3': generateH3Message(nama, kitInfo, jatuhTempoDate, tipePelanggan, idx),
    'H-2': generateH2Message(nama, kitInfo, jatuhTempoDate, tipePelanggan, idx),
    'H-1': generateH1Message(nama, kitInfo, jatuhTempoDate, tipePelanggan, idx),
    'WARNING': generateWarningMessage(nama, kitInfo, jatuhTempoDate, tipePelanggan, idx)
  };

  var message = messages[type] || messages['H-3'];
  return `https://web.whatsapp.com/send?phone=${nomorWA}&text=${encodeURIComponent(message)}`;
}

/**
 * ✅ NEW: Set reminder cell with proper formatting
 */
function setReminderCell(cell, text, url, color) {
  if (url) {
    var richText = SpreadsheetApp.newRichTextValue()
      .setText(text)
      .setLinkUrl(url)
      .build();
    cell.setRichTextValue(richText);
  } else {
    cell.setValue(text);
  }
  
  cell.setFontSize(12)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontColor(color);
}

/**
 * ✅ NEW: Message generators — masing-masing punya 7 variasi pesan
 * Variasi dipilih otomatis berdasarkan nomor WA (idx) sehingga tiap client
 * mendapat pesan berbeda tanpa perubahan apapun di Google Sheets.
 */

// =============================================================================
// H-3: 7 VARIASI PESAN (3 HARI SEBELUM JATUH TEMPO)
// =============================================================================
function generateH3Message(nama, kitInfo, jatuhTempoDate, tipePelanggan, variantIndex) {
  var satuHariSebelumJatuhTempo = new Date(jatuhTempoDate);
  satuHariSebelumJatuhTempo.setDate(satuHariSebelumJatuhTempo.getDate() - 1);
  var namaHariSebelumnya = getNamaHari(satuHariSebelumJatuhTempo);
  var tanggal = formatTanggalIndonesia(jatuhTempoDate);
  var tipePelangganInfo = tipePelanggan ? `🔹 *Tipe Pelanggan:* ${tipePelanggan}\n` : '';
  var idx = variantIndex || 0;

  var variants = [

    // Variasi 0 — Formal & informatif
    `📡 *Halo ${nama}, Pengguna Layanan Internet Starlink* 📡\n\n` +
    `📌 *Informasi Layanan Anda:*\n` +
    `🔹 *Nama:* ${nama}\n` +
    `🔹 *Detail Starlink:*\n${kitInfo}\n` +
    `🔹 *Tanggal Jatuh Tempo:* ${tanggal}\n` +
    tipePelangganInfo +
    `\n📅 Kami ingin menginformasikan bahwa layanan Starlink Anda akan jatuh tempo dalam 3 hari pada tanggal tersebut.\n\n` +
    `Untuk memastikan kesinambungan layanan internet Anda, mohon konfirmasi perpanjangan layanan maksimal pada hari *${namaHariSebelumnya}*.\n\n` +
    `Jika memerlukan bantuan atau informasi lebih lanjut, silakan hubungi kami.\n\n` +
    `Terima kasih atas perhatian dan kepercayaan Anda.\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`,

    // Variasi 1 — Ramah & santai
    `🛰️ *Selamat pagi/siang, ${nama}!*\n\n` +
    `Kami dari *Octolink.id* ingin menyampaikan informasi penting mengenai layanan Starlink Anda.\n\n` +
    `📋 *Detail Langganan:*\n` +
    `👤 *Pelanggan:* ${nama}\n` +
    `📡 *Perangkat Starlink:*\n${kitInfo}\n` +
    `📅 *Jatuh Tempo:* ${tanggal}\n` +
    tipePelangganInfo +
    `\nLayanan Anda akan berakhir *3 hari lagi*. Agar koneksi internet tidak terputus, harap lakukan konfirmasi perpanjangan sebelum hari *${namaHariSebelumnya}*.\n\n` +
    `Kami siap membantu proses perpanjangan Anda. Jangan ragu untuk menghubungi kami kapan saja.\n\n` +
    `Salam hangat,\n🌐 *Octolink.id - Starlink UnOfficial Support* 🌐`,

    // Variasi 2 — Singkat & to the point
    `Halo *${nama}* 👋\n\n` +
    `Ini pemberitahuan dari tim *Octolink.id* mengenai masa aktif layanan Starlink Anda.\n\n` +
    `🔍 *Ringkasan Layanan:*\n` +
    `• Nama: *${nama}*\n` +
    `• Starlink:\n${kitInfo}\n` +
    `• Masa Berlaku Hingga: *${tanggal}*\n` +
    tipePelangganInfo +
    `\n⏳ Tinggal *3 hari* lagi layanan Anda akan jatuh tempo.\n\n` +
    `Silakan konfirmasi perpanjangan paling lambat *${namaHariSebelumnya}* agar layanan tetap aktif tanpa gangguan.\n\n` +
    `Hubungi kami jika ada pertanyaan — kami senang membantu!\n\n` +
    `🚀 *Octolink.id* — Mitra Starlink Anda 🚀`,

    // Variasi 3 — Formal gaya surat
    `📢 *Pengingat Layanan Starlink — ${nama}*\n\n` +
    `Yth. Bapak/Ibu *${nama}*,\n\n` +
    `Bersama ini kami sampaikan bahwa masa aktif layanan Starlink Anda akan segera berakhir.\n\n` +
    `📌 *Data Langganan:*\n` +
    `🔸 Nama: ${nama}\n` +
    `🔸 Info Perangkat:\n${kitInfo}\n` +
    `🔸 Tanggal Jatuh Tempo: ${tanggal}\n` +
    tipePelangganInfo +
    `\nSisa waktu: *3 hari*. Mohon lakukan perpanjangan sebelum *${namaHariSebelumnya}* untuk menghindari putusnya koneksi.\n\n` +
    `Tim kami selalu siap membantu. Hubungi kami untuk informasi lebih lanjut.\n\n` +
    `Hormat kami,\n🚀 *Octolink.id - Starlink UnOfficial Support*`,

    // Variasi 4 — Personal & casual
    `🌐 Halo *${nama}*! Apa kabar?\n\n` +
    `Kami ingin mengingatkan bahwa paket internet Starlink Anda akan segera habis masa berlakunya.\n\n` +
    `📊 *Info Akun Anda:*\n` +
    `➤ Nama: *${nama}*\n` +
    `➤ Perangkat:\n${kitInfo}\n` +
    `➤ Expired: *${tanggal}*\n` +
    tipePelangganInfo +
    `\n📅 Jatuh tempo dalam *3 hari* ke depan. Agar tidak ada gangguan, harap konfirmasi perpanjangan sebelum hari *${namaHariSebelumnya}*.\n\n` +
    `Ada yang bisa kami bantu? Jangan sungkan untuk menghubungi kami ya!\n\n` +
    `Terima kasih 😊\n🚀 *Octolink.id - Starlink UnOfficial Support*`,

    // Variasi 5 — Profesional & terstruktur
    `📡 *Notifikasi Starlink untuk ${nama}*\n\n` +
    `Kepada Yth. *${nama}*,\n\n` +
    `Kami mengingatkan bahwa layanan internet Starlink Anda mendekati tanggal jatuh tempo.\n\n` +
    `🗂️ *Informasi Berlangganan:*\n` +
    `✦ Pelanggan: ${nama}\n` +
    `✦ Detail Perangkat:\n${kitInfo}\n` +
    `✦ Jatuh Tempo: ${tanggal}\n` +
    tipePelangganInfo +
    `\nLayanan akan berakhir dalam *3 hari*. Untuk menjaga kestabilan koneksi Anda, mohon konfirmasi perpanjangan sebelum *${namaHariSebelumnya}*.\n\n` +
    `Kami siap melayani Anda. Silakan hubungi kami bila membutuhkan bantuan.\n\n` +
    `Salam,\n🛰️ *Octolink.id - Starlink UnOfficial Support* 🛰️`,

    // Variasi 6 — Friendly reminder
    `👋 Halo *${nama}*!\n\n` +
    `Sebelum terlambat, kami ingin mengingatkan Anda tentang jatuh tempo layanan Starlink Anda.\n\n` +
    `📝 *Detail Langganan Anda:*\n` +
    `▸ Nama Pelanggan: ${nama}\n` +
    `▸ Perangkat Starlink:\n${kitInfo}\n` +
    `▸ Berakhir Pada: ${tanggal}\n` +
    tipePelangganInfo +
    `\n⏰ Layanan Anda akan jatuh tempo *3 hari* lagi. Pastikan perpanjangan sudah dikonfirmasi sebelum *${namaHariSebelumnya}* agar koneksi Anda tidak terganggu.\n\n` +
    `Butuh bantuan? Kami siap membantu kapan pun!\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`

  ];

  return variants[idx % variants.length];
}

// =============================================================================
// H-2: 7 VARIASI PESAN (2 HARI SEBELUM JATUH TEMPO)
// =============================================================================
function generateH2Message(nama, kitInfo, jatuhTempoDate, tipePelanggan, variantIndex) {
  var tanggal = formatTanggalIndonesia(jatuhTempoDate);
  var tipePelangganInfo = tipePelanggan ? `🔹 *Tipe Pelanggan:* ${tipePelanggan}\n` : '';
  var idx = variantIndex || 0;

  var variants = [

    // Variasi 0 — Peringatan tegas
    `📡 *Halo ${nama}, Pengguna Layanan Internet Starlink* 📡\n\n` +
    `📌 *Informasi Layanan Anda:*\n` +
    `🔹 *Nama:* ${nama}\n` +
    `🔹 *Detail Starlink:*\n${kitInfo}\n` +
    `🔹 *Tanggal Jatuh Tempo:* ${tanggal}\n` +
    tipePelangganInfo +
    `\n⚠️ Pemberitahuan penting: layanan Starlink Anda akan jatuh tempo dalam 2 hari pada tanggal tersebut.\n\n` +
    `Untuk menghindari gangguan layanan, kami merekomendasikan agar perpanjangan segera dilakukan.\n\n` +
    `⏰ Waktu perpanjangan semakin terbatas, mohon segera lakukan konfirmasi.\n\n` +
    `Jika memerlukan bantuan atau informasi lebih lanjut, silakan hubungi kami.\n\n` +
    `Terima kasih atas perhatian dan kepercayaan Anda.\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`,

    // Variasi 1 — Mendesak tapi sopan
    `⚠️ *Pemberitahuan Penting — ${nama}*\n\n` +
    `Halo *${nama}*,\n\n` +
    `Kami mengingatkan Anda bahwa masa aktif layanan Starlink Anda tinggal *2 hari* lagi.\n\n` +
    `📋 *Detail Akun:*\n` +
    `👤 Nama: ${nama}\n` +
    `📡 Perangkat:\n${kitInfo}\n` +
    `📅 Jatuh Tempo: ${tanggal}\n` +
    tipePelangganInfo +
    `\n🔔 Segera lakukan konfirmasi perpanjangan agar layanan internet Anda tidak terganggu.\n\n` +
    `Waktu semakin singkat — jangan sampai tertinggal!\n\n` +
    `Kami siap membantu Anda. Hubungi kami sekarang.\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`,

    // Variasi 2 — Ringkas & jelas
    `🔔 Halo *${nama}*!\n\n` +
    `Ini pengingat dari *Octolink.id* — layanan Starlink Anda *2 hari lagi* akan berakhir.\n\n` +
    `📊 *Data Layanan:*\n` +
    `• Pelanggan: *${nama}*\n` +
    `• Starlink:\n${kitInfo}\n` +
    `• Expired: *${tanggal}*\n` +
    tipePelangganInfo +
    `\n⏳ Jangan tunda lagi! Konfirmasi perpanjangan segera agar koneksi Anda tetap lancar.\n\n` +
    `Langsung hubungi kami ya — kami siap melayani!\n\n` +
    `🛰️ *Octolink.id - Starlink UnOfficial Support* 🛰️`,

    // Variasi 3 — Formal mendesak
    `📣 *Peringatan Jatuh Tempo Layanan Starlink*\n\n` +
    `Yth. *${nama}*,\n\n` +
    `Dengan hormat, kami menginformasikan bahwa masa aktif layanan Starlink Anda akan berakhir dalam *2 hari*.\n\n` +
    `🗂️ *Rincian Langganan:*\n` +
    `🔸 Pelanggan: ${nama}\n` +
    `🔸 Perangkat Starlink:\n${kitInfo}\n` +
    `🔸 Berakhir: ${tanggal}\n` +
    tipePelangganInfo +
    `\n⚠️ Mohon segera lakukan konfirmasi perpanjangan untuk memastikan layanan tetap berjalan tanpa gangguan.\n\n` +
    `Tim kami tersedia untuk membantu Anda kapan saja.\n\n` +
    `Hormat kami,\n🚀 *Octolink.id - Starlink UnOfficial Support*`,

    // Variasi 4 — Personal & urgent
    `Halo *${nama}* 👋\n\n` +
    `Wah, sudah hampir habis nih masa aktif Starlink Anda! Tinggal *2 hari* lagi.\n\n` +
    `📌 *Info Layanan:*\n` +
    `➤ Nama: *${nama}*\n` +
    `➤ Perangkat:\n${kitInfo}\n` +
    `➤ Tanggal Berakhir: *${tanggal}*\n` +
    tipePelangganInfo +
    `\n⏰ Ayo segera konfirmasi perpanjangan agar internet Anda tidak terputus di tengah aktivitas!\n\n` +
    `Hubungi kami sekarang — proses cepat dan mudah.\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support*`,

    // Variasi 5 — Tegas & profesional
    `⚡ *Urgent: Layanan Starlink ${nama} Berakhir 2 Hari Lagi*\n\n` +
    `Kepada *${nama}*,\n\n` +
    `Harap perhatikan bahwa layanan internet Starlink Anda akan jatuh tempo *2 hari lagi*.\n\n` +
    `📋 *Informasi Akun:*\n` +
    `✦ Nama: ${nama}\n` +
    `✦ Perangkat:\n${kitInfo}\n` +
    `✦ Jatuh Tempo: ${tanggal}\n` +
    tipePelangganInfo +
    `\nUntuk menghindari pemutusan layanan, segera konfirmasi perpanjangan kepada kami.\n\n` +
    `⏰ Waktu terus berjalan — jangan sampai layanan Anda terputus!\n\n` +
    `Silakan hubungi kami untuk informasi lebih lanjut.\n\n` +
    `🛰️ *Octolink.id - Starlink UnOfficial Support* 🛰️`,

    // Variasi 6 — Friendly warning
    `🌐 *Hai ${nama}!*\n\n` +
    `Sekedar mengingatkan — paket Starlink Anda akan berakhir dalam *2 hari*. Jangan sampai lupa ya!\n\n` +
    `📝 *Detail Paket Anda:*\n` +
    `▸ Nama: ${nama}\n` +
    `▸ Starlink:\n${kitInfo}\n` +
    `▸ Berlaku Hingga: ${tanggal}\n` +
    tipePelangganInfo +
    `\n⚠️ Segera lakukan perpanjangan agar koneksi internet Anda tidak terganggu.\n\n` +
    `Kami selalu siap bantu! Hubungi tim kami kapan saja.\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`

  ];

  return variants[idx % variants.length];
}

// =============================================================================
// H-1: 7 VARIASI PESAN (1 HARI SEBELUM JATUH TEMPO)
// =============================================================================
function generateH1Message(nama, kitInfo, jatuhTempoDate, tipePelanggan, variantIndex) {
  var tanggal = formatTanggalIndonesia(jatuhTempoDate);
  var tipePelangganInfo = tipePelanggan ? `🔹 *Tipe Pelanggan:* ${tipePelanggan}\n` : '';
  var idx = variantIndex || 0;

  var variants = [

    // Variasi 0 — Urgent kritis
    `📡 *Halo ${nama}, Pengguna Layanan Internet Starlink* 📡\n\n` +
    `📌 *Informasi Layanan Anda:*\n` +
    `🔹 *Nama:* ${nama}\n` +
    `🔹 *Detail Starlink:*\n${kitInfo}\n` +
    `🔹 *Tanggal Jatuh Tempo:* ${tanggal}\n` +
    tipePelangganInfo +
    `\n🚨 Peringatan urgent: layanan Starlink Anda akan jatuh tempo BESOK pada tanggal tersebut.\n\n` +
    `Mohon konfirmasi perpanjangan layanan HARI INI untuk menghindari terputusnya koneksi internet.\n\n` +
    `⏰ Ini adalah pengingat terakhir sebelum masa aktif berakhir.\n\n` +
    `Jika memerlukan bantuan atau informasi lebih lanjut, silakan hubungi kami segera.\n\n` +
    `Terima kasih atas perhatian dan kepercayaan Anda.\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`,

    // Variasi 1 — Mendesak & personal
    `🚨 *PENTING: ${nama}, Layanan Starlink Anda Berakhir BESOK!*\n\n` +
    `Halo *${nama}*,\n\n` +
    `Ini adalah pengingat *TERAKHIR* — layanan Starlink Anda berakhir *BESOK*, ${tanggal}.\n\n` +
    `📋 *Detail Akun:*\n` +
    `👤 Nama: ${nama}\n` +
    `📡 Perangkat:\n${kitInfo}\n` +
    `📅 Berakhir: ${tanggal}\n` +
    tipePelangganInfo +
    `\n⏰ *Mohon konfirmasi perpanjangan HARI INI* sebelum layanan diputus.\n\n` +
    `Hubungi kami segera — kami siap memproses perpanjangan Anda sekarang!\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`,

    // Variasi 2 — Tegas & singkat
    `⏰ *Terakhir! Starlink ${nama} Berakhir Besok*\n\n` +
    `Halo *${nama}* 👋\n\n` +
    `Masa aktif layanan Starlink Anda tinggal *1 hari lagi*!\n\n` +
    `🔍 *Info Layanan:*\n` +
    `• Nama: *${nama}*\n` +
    `• Starlink:\n${kitInfo}\n` +
    `• Expired: *${tanggal}*\n` +
    tipePelangganInfo +
    `\n🚨 Segera konfirmasi perpanjangan *HARI INI* untuk menghindari gangguan koneksi.\n\n` +
    `Jangan tunda lagi — hubungi kami sekarang!\n\n` +
    `🛰️ *Octolink.id - Starlink UnOfficial Support* 🛰️`,

    // Variasi 3 — Formal & kritis
    `🚨 *PENGINGAT KRITIS — Jatuh Tempo Besok*\n\n` +
    `Yth. *${nama}*,\n\n` +
    `Dengan hormat, kami menginformasikan bahwa layanan Starlink Anda akan berakhir *BESOK*, ${tanggal}.\n\n` +
    `📌 *Data Langganan:*\n` +
    `🔸 Pelanggan: ${nama}\n` +
    `🔸 Perangkat:\n${kitInfo}\n` +
    `🔸 Jatuh Tempo: ${tanggal}\n` +
    tipePelangganInfo +
    `\n⚠️ Mohon lakukan konfirmasi perpanjangan *HARI INI* untuk memastikan kelangsungan layanan internet Anda.\n\n` +
    `Ini adalah pemberitahuan terakhir sebelum masa aktif berakhir.\n\n` +
    `Hormat kami,\n🚀 *Octolink.id - Starlink UnOfficial Support*`,

    // Variasi 4 — Dramatik & urgent
    `🔴 *ALERT! Layanan Starlink Berakhir 1 Hari Lagi!*\n\n` +
    `*${nama}*, tolong perhatikan ini!\n\n` +
    `Koneksi internet Starlink Anda akan *MATI BESOK* jika tidak diperpanjang hari ini.\n\n` +
    `📊 *Status Langganan:*\n` +
    `➤ Nama: *${nama}*\n` +
    `➤ Perangkat:\n${kitInfo}\n` +
    `➤ Berakhir: *${tanggal}* (BESOK!)\n` +
    tipePelangganInfo +
    `\n🚨 Ambil tindakan sekarang! Konfirmasi perpanjangan *HARI INI* sebelum terlambat.\n\n` +
    `Kami siap proses cepat — hubungi kami segera!\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support*`,

    // Variasi 5 — Langsung & mendesak
    `⚡ Halo *${nama}*!\n\n` +
    `Hanya ada *1 hari* tersisa sebelum layanan Starlink Anda berakhir pada ${tanggal}.\n\n` +
    `📋 *Ringkasan:*\n` +
    `✦ Nama: ${nama}\n` +
    `✦ Starlink:\n${kitInfo}\n` +
    `✦ Expired: ${tanggal}\n` +
    tipePelangganInfo +
    `\n⏰ Ini pengingat terakhir dari kami. Mohon konfirmasi perpanjangan *SEKARANG* agar layanan tidak terputus.\n\n` +
    `Hubungi kami untuk proses yang cepat dan mudah!\n\n` +
    `🛰️ *Octolink.id - Starlink UnOfficial Support* 🛰️`,

    // Variasi 6 — Empatik & mendesak
    `😟 *Waduh, Starlink ${nama} Mau Berakhir Besok!*\n\n` +
    `Halo *${nama}*,\n\n` +
    `Kami khawatir dengan kelancaran internet Anda karena layanan Starlink *berakhir BESOK*, ${tanggal}.\n\n` +
    `📝 *Detail Paket:*\n` +
    `▸ Nama: ${nama}\n` +
    `▸ Perangkat Starlink:\n${kitInfo}\n` +
    `▸ Berakhir: ${tanggal}\n` +
    tipePelangganInfo +
    `\n🙏 Mohon segera konfirmasi perpanjangan *HARI INI* ya — jangan sampai internet Anda mati!\n\n` +
    `Kami di sini untuk membantu. Hubungi kami sekarang!\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`

  ];

  return variants[idx % variants.length];
}

// =============================================================================
// WARNING: 7 VARIASI PESAN (HARI JATUH TEMPO / HARI INI)
// =============================================================================
function generateWarningMessage(nama, kitInfo, jatuhTempoDate, tipePelanggan, variantIndex) {
  var tanggal = formatTanggalIndonesia(jatuhTempoDate);
  var tipePelangganInfo = tipePelanggan ? `🔹 *Tipe Pelanggan:* ${tipePelanggan}\n` : '';
  var idx = variantIndex || 0;

  var variants = [

    // Variasi 0 — Darurat penuh
    `⚠️ *PENTING: Layanan Internet Starlink Berakhir Hari Ini* ⚠️\n\n` +
    `👤 *${nama}*\n` +
    `📅 *Masa Aktif Berakhir: HARI INI* (${tanggal})\n` +
    tipePelangganInfo +
    `\n🛰️ *Starlink yang terpengaruh:*\n${kitInfo}\n\n` +
    `⏰ *Mohon konfirmasi perpanjangan secepatnya untuk menghindari pemutusan layanan.*\n\n` +
    `Untuk informasi perpanjangan atau bantuan teknis, silakan hubungi kami.\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`,

    // Variasi 1 — Kritis & mendesak
    `🚨🚨 *DARURAT: Starlink ${nama} Berakhir HARI INI!* 🚨🚨\n\n` +
    `*${nama}*, layanan internet Starlink Anda *BERAKHIR HARI INI*, ${tanggal}!\n\n` +
    `📡 *Perangkat yang terdampak:*\n${kitInfo}\n` +
    tipePelangganInfo +
    `\n🔴 Tanpa perpanjangan *SEKARANG*, koneksi internet Anda akan segera diputus.\n\n` +
    `⏰ *Hubungi kami SEGERA untuk memproses perpanjangan!*\n\n` +
    `Jangan tunda lagi — setiap menit berharga!\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`,

    // Variasi 2 — Singkat & to the point
    `🔴 *LAYANAN BERAKHIR HARI INI!*\n\n` +
    `Halo *${nama}* 👋\n\n` +
    `Masa aktif Starlink Anda *habis HARI INI* (${tanggal}).\n\n` +
    `🛰️ *Perangkat:*\n${kitInfo}\n` +
    tipePelangganInfo +
    `\n⚠️ Untuk menghindari pemutusan internet, konfirmasi perpanjangan *SEKARANG*.\n\n` +
    `Hubungi kami langsung — proses cepat!\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support*`,

    // Variasi 3 — Formal darurat
    `⛔ *PEMBERITAHUAN DARURAT — Jatuh Tempo Hari Ini*\n\n` +
    `Yth. *${nama}*,\n\n` +
    `Kami menginformasikan bahwa layanan Starlink Anda *BERAKHIR HARI INI*, ${tanggal}.\n\n` +
    `📌 *Detail Layanan:*\n` +
    `🔸 Pelanggan: ${nama}\n` +
    `🔸 Perangkat:\n${kitInfo}\n` +
    `🔸 Status: *BERAKHIR HARI INI*\n` +
    tipePelangganInfo +
    `\n🚨 Segera lakukan konfirmasi perpanjangan untuk menghindari pemutusan layanan internet.\n\n` +
    `Tim kami siap memproses perpanjangan Anda dengan cepat.\n\n` +
    `Hormat kami,\n🚀 *Octolink.id - Starlink UnOfficial Support*`,

    // Variasi 4 — Dramatis & personal
    `😱 *${nama}! Internet Starlink Anda Berakhir HARI INI!*\n\n` +
    `Ini *peringatan terakhir* — masa aktif Starlink Anda habis *SEKARANG*, ${tanggal}!\n\n` +
    `📊 *Status Akun:*\n` +
    `➤ Nama: *${nama}*\n` +
    `➤ Starlink:\n${kitInfo}\n` +
    `➤ Status: *BERAKHIR HARI INI* ❗\n` +
    tipePelangganInfo +
    `\n🔴 Jika tidak diperpanjang, koneksi internet Anda akan *segera diputus*.\n\n` +
    `Hubungi kami *SEKARANG JUGA* untuk perpanjangan cepat!\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`,

    // Variasi 5 — Tegas & langsung
    `⚡ *STATUS: BERAKHIR HARI INI — ${nama}*\n\n` +
    `*${nama}*, tanpa tindakan segera, layanan internet Starlink Anda akan *diputus hari ini*.\n\n` +
    `📋 *Info Akun:*\n` +
    `✦ Nama: ${nama}\n` +
    `✦ Perangkat:\n${kitInfo}\n` +
    `✦ Berakhir: *${tanggal} (HARI INI)*\n` +
    tipePelangganInfo +
    `\n⏰ Konfirmasi perpanjangan *SEKARANG* untuk menjaga koneksi Anda tetap aktif.\n\n` +
    `Kami siap membantu — hubungi segera!\n\n` +
    `🛰️ *Octolink.id - Starlink UnOfficial Support* 🛰️`,

    // Variasi 6 — Empatik & mendesak
    `🙏 Halo *${nama}*, mohon perhatikan pesan ini!\n\n` +
    `Dengan berat hati kami menginformasikan bahwa layanan Starlink Anda *berakhir HARI INI*, ${tanggal}.\n\n` +
    `📝 *Detail Layanan:*\n` +
    `▸ Nama: ${nama}\n` +
    `▸ Perangkat Starlink:\n${kitInfo}\n` +
    `▸ Status: *Berakhir ${tanggal}*\n` +
    tipePelangganInfo +
    `\n⚠️ Agar aktivitas internet Anda tidak terganggu, mohon segera konfirmasi perpanjangan kepada kami.\n\n` +
    `Kami siap membantu Anda. Jangan ragu untuk menghubungi kami sekarang!\n\n` +
    `🚀 *Octolink.id - Starlink UnOfficial Support* 🚀`

  ];

  return variants[idx % variants.length];
}

/**
 * ✅ NEW: Reactivate filter
 */
function reactivateFilter(sheet) {
  Utilities.sleep(3000);
  var rangeToFilter = sheet.getRange(1, 1, sheet.getLastRow(), COLUMNS.TOMBOL_WHATSAPP + 1); // Y column
  sheet.getFilter() && sheet.getFilter().remove();
  rangeToFilter.createFilter();
  Logger.log("✅ Filter dari kolom A sampai Y diaktifkan kembali.");
  _sortByJatuhTempo(sheet);
  Logger.log("✅ Data diurutkan A–Z berdasarkan Tanggal Jatuh Tempo (kolom K).");
  _formatColumnsStyle(sheet);
}

