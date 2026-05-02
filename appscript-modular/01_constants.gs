// ========================================
// UPDATED: Cleaned unused functions + Enhanced with Auto Filter for Today's Data
// New Column Structure:
// I = KIT Number | J = Serial Number | K = Tanggal Jatuh Tempo
// L-R = Data lain (geser +1) | S-T = Kosong (cadangan)
// U = JATUH TEMPO | V = PROSES | W = SEGERA | X = OBSERVASI | Y = Tombol WhatsApp
// ========================================

// ========================================
// 📋 1. CONFIGURATION & CONSTANTS
// ========================================

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const REQUEST_ID_CACHE_KEY_PREFIX = 'SUBMITTED_REQUEST_';
const REQUEST_ID_CACHE_DURATION = 300000; // 5 minutes in milliseconds

const REPORT_SPREADSHEET_ID = "1wc21VKoWbvGXHIC0Izcy-_H50bzNy4vrLLV6jBMH-a0";
const REPORT_SPREADSHEET_NAME = "Laporan Starlink 2025";

// Column Mapping (NEW STRUCTURE)
const COLUMNS = {
  NAMA: 0,              // A
  LOGIN_GMAIL: 1,       // B  
  LOGIN_STARLINK: 2,    // C
  LOGIN_ALTERNATIF: 3,  // D
  ACC_NO: 4,            // E
  EMAIL_CLIENT: 5,      // F
  NOMOR_CS: 6,          // G
  ALAMAT: 7,            // H
  KIT_NUMBER: 8,        // I ✅ NEW: KIT Number
  SERIAL_NUMBER: 9,     // J ✅ NEW: Serial Number
  JATUH_TEMPO: 10,      // K ✅ MOVED: Tanggal Jatuh Tempo
  KODE: 11,             // L
  PAYMENT: 12,          // M
  TRANSACTION_ID: 13,   // N
  STATUS: 14,           // O
  PAKET: 15,            // P
  LAST_4_DIGIT: 16,     // Q
  NO_REGISTER: 17,      // R
  TIPE_PELANGGAN: 18,   // S ✅ NEW: Tipe Pelanggan
  CADANGAN_2: 19,       // T ✅ NEW: Kosong
  LABEL_JATUH_TEMPO: 20, // U ✅ MOVED: "JATUH TEMPO"
  LABEL_PROSES: 21,     // V ✅ MOVED: "PROSES"
  LABEL_SEGERA: 22,     // W ✅ MOVED: "SEGARA"
  LABEL_OBSERVASI: 23,  // X ✅ MOVED: "OBSERVASI"
  TOMBOL_WHATSAPP: 24   // Y ✅ MOVED: Tombol WhatsApp
};

