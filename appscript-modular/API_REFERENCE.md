# Appscript API Reference
> Panduan lengkap untuk developer HTML. Semua fungsi yang bisa dipanggil dari HTML ada di sini.

---

## Cara Memanggil dari HTML

Semua request dikirim ke **satu URL** (Web App URL dari deployment Apps Script):

```js
const SCRIPT_URL = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
```

**Format GET request:**
```js
fetch(`${SCRIPT_URL}?action=namaAction&param1=nilai1&param2=nilai2`)
  .then(r => r.json())
  .then(data => console.log(data));
```

**Format POST request:**
```js
fetch(SCRIPT_URL, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ action: 'namaAction', ...data })
});
```

**Response selalu berformat:**
```json
{ "status": "success" | "error", "data": ..., "message": "..." }
```

---

## Daftar Actions

### Data Client

| Action | Method | Parameter | Deskripsi | File |
|--------|--------|-----------|-----------|------|
| `getAllClientDataComplete` | GET | — | Semua client dari sheet Client Aktif | `08_client_management.gs` |
| `getClientLepasData` | GET | — | Semua client dari sheet Client Lepas | `08_client_management.gs` |
| `getClientNonAktifData` | GET | — | Semua client dari sheet Client Non Aktif | `08_client_management.gs` |
| `getClientTertagihData` | GET | — | Semua client dari sheet Client Tertagih | `08_client_management.gs` |
| `getAkunKosongData` | GET | — | Data dari sheet Akun Kosong | `08_client_management.gs` |
| `getClientsBySheet` | GET | `sheetName` *(opsional, default: 'Client Aktif')* | Data client per sheet tertentu | `08_client_management.gs` |
| `getClientNames` | GET | — | Daftar nama semua client | `13_client_status.gs` |
| `getAllStatus` | GET | — | Semua client dikelompokkan per status | `13_client_status.gs` |

**Contoh:**
```js
// Ambil semua client aktif
fetch(`${SCRIPT_URL}?action=getAllClientDataComplete`)

// Ambil client dari sheet tertentu
fetch(`${SCRIPT_URL}?action=getClientsBySheet&sheetName=Client Lepas`)
```

---

### Search

Semua search menggunakan `action=search`. Bedakan admin vs customer dengan parameter `version`:

| Parameter | Nilai | Efek |
|-----------|-------|------|
| `version` | `admin` | Data lengkap tanpa masking |
| `version` | *(kosong)* | Data customer, field sensitif di-mask |
| `multiple` | `true` | Return semua hasil (array), bukan hanya 1 |

**Parameter pencarian (pilih salah satu):**

| Parameter | Cari berdasarkan | File |
|-----------|-----------------|------|
| `kit` | KIT Number | `09_search_admin.gs` / `10_search_customer.gs` |
| `serial` | Serial Number | `09_search_admin.gs` / `10_search_customer.gs` |
| `login` | Login Starlink | `09_search_admin.gs` / `10_search_customer.gs` |
| `email` | Email Client | `09_search_admin.gs` / `10_search_customer.gs` |
| `phone` | Nomor CS | `09_search_admin.gs` / `10_search_customer.gs` |
| `account` | Account Number | `09_search_admin.gs` / `10_search_customer.gs` |
| `name` | Nama client | `09_search_admin.gs` / `10_search_customer.gs` |

**Contoh:**
```js
// Search by KIT (admin, semua hasil)
fetch(`${SCRIPT_URL}?action=search&kit=KIT-001&version=admin&multiple=true`)

// Search by nama (customer view, 1 hasil)
fetch(`${SCRIPT_URL}?action=search&name=John Doe`)

// Search by serial (admin, 1 hasil)
fetch(`${SCRIPT_URL}?action=search&serial=SN123456&version=admin`)
```

---

### Submit & Validasi Pembayaran

| Action | Method | Parameter | Deskripsi | File |
|--------|--------|-----------|-----------|------|
| `validateKitMulti` | GET | `kit`, `month`, `year` | Validasi KIT sebelum submit pembayaran | `14_form_validation.gs` |
| `validateClientByName` | GET | `clientName`, `month`, `year` | Validasi client by nama | `14_form_validation.gs` |
| `checkDuplicates` | GET | `month`, `year` | Cek duplikat transaksi bulan tsb | `12_report_management.gs` |
| `checkDuplicates` | GET | `kitNumbers` *(comma-separated)*, `month`, `year` | Cek duplikat by KIT numbers | `12_report_management.gs` |
| `submitPaymentMultiKit` | GET/POST | `data` *(JSON string)* | Submit pembayaran multi-KIT | `12_report_management.gs` |

**Contoh:**
```js
// Validasi KIT
fetch(`${SCRIPT_URL}?action=validateKitMulti&kit=KIT-001&month=5&year=2026`)

// Submit pembayaran
const formData = { /* ... */ };
fetch(`${SCRIPT_URL}?action=submitPaymentMultiKit&data=${encodeURIComponent(JSON.stringify(formData))}`)

// Cek duplikat
fetch(`${SCRIPT_URL}?action=checkDuplicates&month=5&year=2026`)
```

---

### Submit Client Baru

| Action | Method | Parameter | Deskripsi | File |
|--------|--------|-----------|-----------|------|
| `submitClientData` | GET/POST | `data` *(JSON string)* | Tambah client baru ke spreadsheet | `08_client_management.gs` |

**Contoh:**
```js
const clientData = {
  nama: 'John Doe',
  loginGmail: '...',
  // ... field lainnya
};
fetch(`${SCRIPT_URL}?action=submitClientData&data=${encodeURIComponent(JSON.stringify(clientData))}`)
```

---

### Laporan

| Action | Method | Parameter | Deskripsi | File |
|--------|--------|-----------|-----------|------|
| `getReportData` | GET | `month`, `year` *(opsional, default bulan ini)* | Data laporan bulanan | `12_report_management.gs` |

**Contoh:**
```js
fetch(`${SCRIPT_URL}?action=getReportData&month=5&year=2026`)
```

---

### Sistem Approval (Admin)

| Action | Method | Parameter | Deskripsi | File |
|--------|--------|-----------|-----------|------|
| `isAdmin` | GET | `email` | Cek apakah email adalah admin | `07_api_handlers.gs` |
| `getPendingApprovals` | GET | `filterStatus` *(opsional, default: 'Pending')* | Ambil daftar approval pending | `07_api_handlers.gs` |
| `addPendingApproval` | GET | `rowId`, `changedBy`, `changedByName`, `sourceSheet`, `changedColumns` *(JSON)*, `originalData` *(JSON)*, `newData` *(JSON)* | Tambah request approval | `07_api_handlers.gs` |
| `approveChange` | GET | `approvalId`, `adminEmail` | Setujui satu perubahan | `07_api_handlers.gs` |
| `rejectChange` | GET | `approvalId`, `adminEmail`, `reason` | Tolak satu perubahan | `07_api_handlers.gs` |
| `approveAll` | GET | `adminEmail` | Setujui semua perubahan pending | `07_api_handlers.gs` |

---

## Fungsi Spreadsheet (google.script.run)

Untuk HTML yang di-*embed* langsung di Spreadsheet (bukan web app eksternal), gunakan `google.script.run`:

| Fungsi | Deskripsi | File |
|--------|-----------|------|
| `resetAndActivateAllFilters()` | Reset & aktifkan filter semua sheet | `03_menu_filter.gs` |
| `resetAndActivateFilterClientAktif()` | Reset & aktifkan filter Client Aktif | `03_menu_filter.gs` |
| `filterSegera()` | Filter baris dengan label SEGERA | `03_menu_filter.gs` |
| `filterProses()` | Filter baris dengan label PROSES | `03_menu_filter.gs` |
| `filterJatuhTempo()` | Filter baris JATUH TEMPO | `03_menu_filter.gs` |
| `filterReady()` | Filter baris READY | `03_menu_filter.gs` |
| `filterConfirmed()` | Filter baris CONFIRMED | `03_menu_filter.gs` |
| `filterPaid()` | Filter baris PAID | `03_menu_filter.gs` |
| `filterPlanned()` | Filter baris PLANNED | `03_menu_filter.gs` |
| `filterIdTransaksi()` | Filter by ID Transaksi | `03_menu_filter.gs` |
| `showFilterCustomDialog()` | Buka dialog filter custom multiple | `03_menu_filter.gs` |
| `refreshWALink()` | Refresh semua tombol WhatsApp | `03_menu_filter.gs` |
| `kirimWhatsApp(skipEmail)` | Kirim reminder WhatsApp ke semua client | `11_whatsapp.gs` |

**Contoh:**
```js
google.script.run
  .withSuccessHandler(function(result) { console.log(result); })
  .withFailureHandler(function(err) { console.error(err); })
  .filterSegera();
```

---

## Konstanta Kolom (COLUMNS)

Saat perlu referensi nomor kolom di spreadsheet, pakai konstanta dari `01_constants.gs`:

```js
// Contoh: COLUMNS.KIT_NUMBER = 8 (kolom I, 0-indexed)
COLUMNS.NAMA              // A = 0
COLUMNS.KIT_NUMBER        // I = 8
COLUMNS.SERIAL_NUMBER     // J = 9
COLUMNS.JATUH_TEMPO       // K = 10
COLUMNS.TRANSACTION_ID    // N = 13
COLUMNS.STATUS            // O = 14
COLUMNS.PAKET             // P = 15
COLUMNS.TIPE_PELANGGAN    // S = 18
COLUMNS.LABEL_JATUH_TEMPO // U = 20
COLUMNS.LABEL_PROSES      // V = 21
COLUMNS.LABEL_SEGERA      // W = 22
COLUMNS.LABEL_OBSERVASI   // X = 23
COLUMNS.TOMBOL_WHATSAPP   // Y = 24
```

---

## Quick Cheatsheet — HTML Mau Buat Apa?

| Mau buat HTML untuk... | Action yang dipakai |
|------------------------|---------------------|
| Lihat daftar client | `getAllClientDataComplete`, `getClientsBySheet` |
| Search client | `action=search` + parameter kit/serial/nama/dll |
| Form tambah client baru | `submitClientData` |
| Form bayar / submit transaksi | `validateKitMulti` → `checkDuplicates` → `submitPaymentMultiKit` |
| Laporan bulanan | `getReportData` |
| Dashboard status client | `getAllStatus`, `getClientNames` |
| Admin approval system | `isAdmin` → `getPendingApprovals` → `approveChange` / `rejectChange` |
| Filter spreadsheet (embedded) | `google.script.run.filterSegera()` dll |
