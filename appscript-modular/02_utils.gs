// ========================================
// 🛠️ 2. UTILITY FUNCTIONS
// ========================================

function formatRupiah(amount) {
  if (!amount || isNaN(amount)) return 'Rp 0';
  return 'Rp ' + Number(amount).toLocaleString('id-ID');
}

function formatTanggalIndonesia(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return '-';
  var bulanIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return date.getDate() + " " + bulanIndo[date.getMonth()] + " " + date.getFullYear();
}

function getNamaHari(date) {
  var hariIndo = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  return hariIndo[date.getDay()];
}

/**
 * ✅ NEW: Parse FormData format dari POST request (untuk add-v2.html)
 * Converts "key1=value1&key2=value2" to {key1: value1, key2: value2}
 */
function parseFormData(contents) {
  var params = {};
  var pairs = contents.split('&');

  for (var i = 0; i < pairs.length; i++) {
    var pair = pairs[i].split('=');
    if (pair.length >= 2) {
      var key = decodeURIComponent(pair[0]);
      var value = decodeURIComponent(pair[1].replace(/\+/g, ' '));
      params[key] = value;
    }
  }

  return params;
}

