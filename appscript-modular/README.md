# Appscript Modular — Starlink Management System

Folder ini berisi versi **modular** dari appscript, dipecah per fitur/HTML agar tidak saling overwrite.

## Kenapa Dipisah?

Masalah sebelumnya: satu file `appscript-270426` berisi 6621 baris untuk semua fitur.
Akibatnya, edit appscript untuk HTML A bisa menimpa bagian yang dipakai HTML B.

**Solusi:** Setiap fitur punya file `.gs` sendiri. Di Google Apps Script, semua file `.gs`
dalam satu project tetap bisa saling akses fungsinya — jadi tidak ada yang rusak.

---

## Peta File → HTML / Fitur

| File | Isi | Dipakai oleh |
|------|-----|--------------|
| `01_constants.gs` | COLUMNS config, konstanta global | **Semua** (shared) |
| `02_utils.gs` | formatRupiah, formatTanggal, parseFormData | **Semua** (shared) |
| `03_menu_filter.gs` | onOpen menu, resetFilter, filterSegera/Proses/dll | **HTML: Filter / Spreadsheet toolbar** |
| `04_transaction_id.gs` | generateTransactionID, generateBatch, updateClientDataBatch | **HTML: Pembayaran / Submit Transaksi** |
| `05_report_formatting.gs` | calculateRunningTotal, sortAndOrganize, applyFormatting, dailySummary | **HTML: Laporan** |
| `06_maintenance.gs` | fixAllExistingSheets, checkSystemStatus, initializeStarlinkSystem | **Admin / Maintenance** |
| `07_api_handlers.gs` | doGet, doPost, doOptions, createCorsResponse | **Semua HTML** (API layer) |
| `08_client_management.gs` | submitClientData, getClientsBySheet, getAllClientData, getClientLepas/NonAktif/dll | **HTML: Manajemen Client** |
| `09_search_admin.gs` | searchByKit, searchBySerial, searchByEmail, searchByNama (versi admin) | **HTML: Search Admin** |
| `10_search_customer.gs` | searchByKit/Serial/Email/Nama (versi customer, masked) + universalSearch | **HTML: Search Customer** |
| `11_whatsapp.gs` | kirimWhatsApp, generateWALink, generateH1/H2/H3/WarningMessage | **HTML: WhatsApp Reminder** |
| `12_report_management.gs` | getReportSheet, submitPaymentDataMultiKit, checkDuplicate, validateKit | **HTML: Submit Pembayaran** |
| `13_client_status.gs` | getAllClientNames, getAllClientsByStatus | **HTML: Dashboard / Status Client** |
| `14_form_validation.gs` | handleFormValidation, validateClientByName, checkDuplicatesInMonth | **HTML: Form Validasi** |

---

## Cara Kerja di Google Apps Script

1. Buka Apps Script project kamu
2. **Hapus** konten file `.gs` yang ada (atau rename jadi `_archive.gs`)
3. **Buat file baru** untuk setiap file di folder ini, copy-paste isinya
4. Semua fungsi tetap bisa saling panggil antar file — GAS compile semua `.gs` jadi satu scope

---

## Aturan Edit

> **Sebelum edit: cek dulu tabel di atas, file mana yang relevan dengan HTML yang sedang dikerjakan.**

- Kerja di HTML Filter → edit `03_menu_filter.gs`
- Kerja di HTML Pembayaran → edit `04_transaction_id.gs` + `12_report_management.gs`
- Kerja di HTML Search → edit `09_search_admin.gs` atau `10_search_customer.gs`
- Ubah konstanta kolom → edit `01_constants.gs`
- Ubah format API response → edit `07_api_handlers.gs`

File shared (`01`, `02`, `07`) hati-hati diedit karena dipakai semua HTML.
