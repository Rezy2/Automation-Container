# Inventory LCI - WWT/RWT

Otomasi pencatatan stok material untuk container **WWT** dan **RWT** berbasis Google Sheets dan Google Apps Script, dengan input dari pesan WhatsApp melalui gateway **Fonnte MD**.

Pesan teks dari WhatsApp diparse menjadi draft transaksi. Draft tidak langsung mengubah stok. User harus membalas `OK` untuk menyimpan, atau `REVISI <perbaikan>` untuk membuat draft baru.

## Fitur

- Parsing pesan WhatsApp berbahasa Indonesia untuk transaksi `IN`/`OUT`.
- Deteksi container `WWT`/`RWT`, qty, satuan, user, dan catatan.
- Material yang belum ada di master otomatis dibuat dengan kode baru.
- Alias material disimpan agar penulisan berikutnya lebih mudah dikenali.
- Draft preview sebelum commit, sehingga salah input tidak langsung masuk stok.
- Commit idempotent berdasarkan WA message ID.
- Validasi stok untuk transaksi `OUT`.
- Audit log untuk setiap commit.
- Inbox/outbox Fonnte dengan retry otomatis.
- Trigger outbox setiap 1 menit.
- Data dummy untuk demo dan pengujian.
- Test parser, material, draft, dan Fonnte tanpa dependency eksternal.

## Alur Sistem

```text
WhatsApp
   |
   v
Fonnte MD webhook
   |
   v
Apps Script doPost()
   |
   +--> validasi token webhook
   +--> abaikan group, media, dan nomor yang tidak diizinkan
   +--> simpan ke tab Inbox (deduplikasi)
   |
   v
handleIncomingMessage_()
   |
   +--> parse pesan
   +--> cari material / buat material baru
   +--> buat draft
   |
   v
Balasan preview draft melalui Outbox
   |
   +--> user balas OK      -> commit transaksi
   +--> user balas REVISI  -> draft baru
```

Saat commit, sistem menulis transaksi ke tab `WWT In`, `WWT Out`, `RWT In`, atau `RWT Out`, mengubah status draft menjadi `POSTED`, dan menulis audit log.

## Struktur Repository

```text
.
├── appsscript.json          # Manifest Apps Script
├── package.json             # Script build dan test
├── scripts/build.js         # Menggabungkan src/*.gs menjadi setup.gs
├── setup.gs                 # Bundle siap tempel ke Apps Script (generated)
├── src/
│   ├── config.gs            # Konstanta, schema, config, helper spreadsheet
│   ├── parser.gs            # Parser pesan WhatsApp
│   ├── material.gs          # Pencarian dan auto-create material
│   ├── draft.gs             # State machine draft dan commit
│   ├── fonnte.gs            # Webhook, inbox, outbox, dan pengiriman Fonnte
│   ├── setup.gs             # Setup schema, menu, formula saldo
│   └── dummy.gs             # Seed dan clear data dummy
└── test/
    ├── parser.test.js
    ├── material.test.js
    ├── draft.test.js
    ├── fonnte.test.js
    └── parser-corpus.csv    # Hasil test parser (generated)
```

`setup.gs` di root adalah hasil build. Jangan edit file itu langsung. Edit file di `src/`, lalu jalankan:

```bash
npm run build
```

## Prasyarat

- Google account dengan akses ke Google Sheets dan Google Apps Script.
- Spreadsheet tujuan.
- Akun Fonnte MD dan perangkat WhatsApp yang sudah terhubung.
- Node.js untuk build dan test lokal. Tidak ada dependency npm eksternal.

## Setup Spreadsheet dan Apps Script

1. Buat atau siapkan spreadsheet.
2. Buka `src/config.gs`, sesuaikan `SPREADSHEET_ID` jika memakai spreadsheet lain.
3. Jalankan build:

   ```bash
   npm run build
   ```

4. Buka spreadsheet, lalu buka **Extensions > Apps Script**.
5. Salin seluruh isi `setup.gs` ke file Apps Script, atau buat file baru dan tempel isinya.
6. Simpan project.
7. Jalankan fungsi `setupSchema()` dari editor Apps Script dan setujui permission.
8. Kembali ke spreadsheet. Menu **Inventory LCI** akan muncul.

`setupSchema()` akan:

- Membuat tab `Data Material`, `WWT In`, `WWT Out`, `RWT In`, `RWT Out`.
- Membuat tab internal `Config`, `Drafts`, `Inbox`, `Outbox`, `Audit Log`, `Code Sequences`, dan `Material Alias`.
- Menyembunyikan tab internal.
- Menulis header, format, validasi, dan formula saldo.
- Mengisi default Config.
- Memasang trigger menu on-open.

Jika spreadsheet sudah berisi data dengan header lama, setup akan menolak dan meminta spreadsheet baru atau data dikosongkan. Ini mencegah kerusakan data.

## Konfigurasi

### Tab Config

Buka menu **Inventory LCI > Buka Config** jika tab Config tersembunyi.

| Key | Default | Keterangan |
| --- | --- | --- |
| `TIMEZONE` | `Asia/Jakarta` | Timezone timestamp dan laporan |
| `FUZZY_THRESHOLD` | `0.75` | Ambang minimal fuzzy matching |
| `FUZZY_MARGIN` | `0.10` | Selisih minimal kandidat teratas dan kedua |
| `AI_ENABLED` | `FALSE` | Reserved untuk AI vision fase 2 |
| `DUMMY_SEEDED` | `FALSE` | Penanda data dummy aktif |
| `FONNTE_ENABLED` | `FALSE` | Aktifkan webhook Fonnte |
| `FONNTE_ALLOWED_NUMBERS` | kosong | Nomor WhatsApp yang diizinkan, pisahkan dengan koma |
| `FONNTE_QUOTE_REPLY` | `TRUE` | Balas dengan quote `inboxid` Fonnte |

Contoh `FONNTE_ALLOWED_NUMBERS`:

```text
6281234567890,6289876543210
```

Nomor akan dinormalisasi ke format `62...`. Nomor `08...`, `8...`, dan `+62...` dianggap sama.

### Script Properties

Buka **Project Settings > Script Properties** di Apps Script.

| Property | Keterangan |
| --- | --- |
| `FONNTE_TOKEN` | Token perangkat Fonnte untuk mengirim pesan |
| `FONNTE_WEBHOOK_TOKEN` | Token rahasia untuk URL webhook |

Jangan menyimpan token di tab Config, source code, atau repository.

## Setup Fonnte

1. Login ke [https://md.fonnte.com](https://md.fonnte.com).
2. Tambahkan device dan scan QR WhatsApp.
3. Buka pengaturan device:
   - Aktifkan **Autoread**.
   - Aktifkan **Inbox**.
   - Matikan pemrosesan group.
4. Salin token device.
5. Isi webhook device dengan URL Apps Script:

   ```text
   https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?token=<FONNTE_WEBHOOK_TOKEN>
   ```

6. Jika memakai Fonnte Flow:
   - `Start` menggunakan **Webhook**.
   - `Response Source` diarahkan ke **Flow**.
7. Simpan pengaturan device.

Catatan:

- Webhook Fonnte mengirim JSON ke Apps Script.
- Apps Script tidak bisa membaca header request, sehingga autentikasi memakai query parameter `token`.
- `doPost()` selalu mengembalikan HTTP 200 dengan body `OK` agar Fonnte tidak melakukan retry berulang saat terjadi error pemrosesan. Error tetap dicatat di Apps Script logs dan tab Inbox.
- Pesan group dan pesan media diabaikan. Integrasi ini hanya memproses teks direct message.

## Deploy Web App

1. Di Apps Script, pilih **Deploy > New deployment**.
2. Pilih tipe **Web app**.
3. Isi:
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
4. Deploy dan setujui permission.
5. Salin URL yang berakhiran `/exec`.
6. Tambahkan `?token=<FONNTE_WEBHOOK_TOKEN>`.
7. Pasang URL tersebut di webhook Fonnte.

Gunakan URL `/exec`, bukan `/dev`. Setelah mengubah kode, buat versi deployment baru agar perubahan dipakai oleh webhook.

## Aktivasi dan Pengecekan

Setelah konfigurasi selesai, jalankan dari editor Apps Script:

```javascript
installFonnteOutboxTrigger();
fonnteStatus();
testFonnteSend();
```

- `installFonnteOutboxTrigger()` membuat trigger `processOutbox` setiap 1 menit.
- `fonnteStatus()` menampilkan status config, token, dan jumlah outbox pending.
- `testFonnteSend()` mengirim pesan tes ke nomor pertama di `FONNTE_ALLOWED_NUMBERS`.

Menu **Inventory LCI** juga menyediakan:

- `Setup Schema`
- `Seed Data Dummy`
- `Hapus Data Dummy`
- `Perbaiki Formula Saldo`
- `Buka Config`
- `Fonnte Status`
- `Test Kirim Fonnte`
- `Proses Outbox`
- `Pasang Trigger Outbox`

## Format Pesan

Field wajib:

- Arah transaksi: `in`/`masuk` atau `out`/`keluar`.
- Nama material.
- Qty dan satuan.
- Container: `WWT` atau `RWT`.
- User yang melakukan transaksi.

Catatan opsional ditulis setelah kata `untuk` atau `buat`.

Contoh:

```text
in elbow 2 pcs ke WWT oleh Fariz
out seal tape 1 roll dari WWT oleh Fariz untuk perbaikan pompa
masuk bearing 2 pcs ke RWT oleh Dimas untuk stok awal
```

Satuan yang didukung antara lain:

```text
pcs, roll, pasang, kg, unit, box, pack, set, batang, lembar, meter, liter, tube, kaleng
```

Satuan mengikuti master material. Jika satuan pesan berbeda dengan master, transaksi ditolak dan user diminta memakai satuan master. Tidak ada konversi satuan otomatis.

## Draft dan Konfirmasi

Setelah pesan valid, sistem membalas preview:

```text
Material: Seal Tape
Material Code: C-001
Qty: 1 roll
Container: WWT
User: Fariz
Catatan: perbaikan pompa

Draft: DRF-20261030-123456-ABCD
Balas OK untuk simpan, atau REVISI <perbaikan>.
```

Perintah konfirmasi:

| Balasan | Aksi |
| --- | --- |
| `OK`, `OKE`, `OKAY`, `SIMPAN`, `POST`, `LANJUT`, `BENAR`, `YA` | Commit draft |
| `REVISI <teks baru>` | Batalkan draft lama dan buat draft baru |
| `REVISI` tanpa teks | Batalkan draft |
| Balasan lain | Tampilkan ulang preview draft |

Status draft:

| Status | Arti |
| --- | --- |
| `PENDING` | Menunggu konfirmasi |
| `NEEDS_CLARIFICATION` | Data belum lengkap atau ambigu |
| `REVISI` | Dibatalkan dan diganti draft baru |
| `POSTED` | Sudah menjadi transaksi |
| `CANCELLED` | Dibatalkan |

## Auto-Create Material

Jika nama material tidak ditemukan di master:

1. Material diklasifikasikan ke kategori:
   - `F` untuk fitting/perpipaan
   - `M` untuk material lain
   - `C` untuk consumable
   - `T` untuk tools
2. Kode baru dibuat dengan format `PREFIX-NNN`, misalnya `F-001`.
3. Nama tampilan diambil dari teks pesan.
4. Satuan diambil dari pesan.
5. `Min Stock` diisi `0`.
6. Spesifikasi dan bahan dikosongkan.
7. Status `ACTIVE`, `Is Dummy` false.
8. Alias ditambahkan ke tab `Material Alias`.

Auto-create hanya dilakukan jika material belum ada. Parser tidak menebak material yang mirip; jika ambigu, user diminta klarifikasi.

## Tab Spreadsheet

| Tab | Fungsi |
| --- | --- |
| `Data Material` | Master material, saldo WWT/RWT, reorder |
| `WWT In` | Transaksi masuk WWT |
| `WWT Out` | Transaksi keluar WWT |
| `RWT In` | Transaksi masuk RWT |
| `RWT Out` | Transaksi keluar RWT |
| `Config` | Konfigurasi aplikasi |
| `Drafts` | Draft transaksi dan statusnya |
| `Inbox` | Pesan webhook yang diterima dan status proses |
| `Outbox` | Antrian balasan WhatsApp |
| `Audit Log` | Jejak commit dan perubahan |
| `Code Sequences` | Sequence kode material |
| `Material Alias` | Alias nama material |

Saldo di `Data Material` dihitung dengan formula `SUMIF` dari tab transaksi. Jangan mengisi kolom saldo secara manual.

## Inbox, Outbox, dan Retry

### Inbox

Setiap pesan webhook dicatat dengan status:

- `PROCESSING`
- `PROCESSED`
- `FAILED`

Deduplikasi memakai `inboxid` Fonnte. Jika `inboxid` tidak ada, sistem memakai hash dari device, sender, timestamp, dan teks pesan.

### Outbox

Balasan disimpan dulu di tab Outbox dengan status:

- `PENDING`
- `SENT`
- `FAILED`

Pengiriman memakai endpoint:

```text
POST https://api.fonnte.com/send
```

Header:

```text
Authorization: <FONNTE_TOKEN>
```

Body dikirim sebagai `application/x-www-form-urlencoded`:

```text
target=<nomor>
message=<teks>
preview=false
connectOnly=false
inboxid=<inboxid>   # jika FONNTE_QUOTE_REPLY = TRUE
```

Retry otomatis:

- Maksimal 5 attempt.
- Backoff 30 detik, 60 detik, 120 detik, 240 detik.
- Maksimal jeda 15 menit.
- Retry untuk HTTP `0`, `408`, `409`, `425`, `429`, `5xx`, `device disconnected`, dan `server issue`.

Trigger `processOutbox` berjalan setiap 1 menit. Outbox juga langsung diproses setelah pesan masuk.

## Build dan Test

```bash
npm run build
npm test
```

`npm test` akan:

1. Membangun ulang `setup.gs`.
2. Menjalankan test parser dengan corpus 62 pesan.
3. Menjalankan test edge case parser.
4. Menjalankan test auto-create material.
5. Menjalankan test draft preview dan commit.
6. Menjalankan test adapter Fonnte.

Output yang diharapkan:

```text
Built setup.gs from 7 source files.
parser tests passed: 62 messages
parser edge cases passed
material auto-create tests passed
draft preview and commit tests passed
fonnte tests passed
```

Test tidak membutuhkan koneksi ke Google atau Fonnte. Test memakai mock untuk Spreadsheet, Properties, UrlFetch, dan LockService.

## Data Dummy

Untuk demo:

```javascript
seedDummy();
```

Untuk menghapus:

```javascript
clearDummy();
```

Data dummy ditandai dengan `Is Dummy` pada master dan `Source = DUMMY` pada transaksi, sehingga bisa dihapus tanpa mengganggu data asli.

## Troubleshooting

| Masalah | Pemeriksaan |
| --- | --- |
| Webhook tidak masuk | Pastikan `FONNTE_ENABLED = TRUE`, URL `/exec` benar, deployment `Anyone`, dan token URL sama dengan `FONNTE_WEBHOOK_TOKEN` |
| Tidak ada balasan | Cek `fonnteStatus()`, token Fonnte, status device, dan outbox pending |
| Pesan diabaikan | Pastikan bukan group, bukan media, dan nomor ada di `FONNTE_ALLOWED_NUMBERS` |
| Balasan terlambat | Jalankan `processOutbox()` atau cek trigger `processOutbox` |
| Material tidak dikenali | Kirim nama material lebih spesifik atau cek tab `Material Alias` |
| Satuan ditolak | Gunakan satuan yang sama dengan master material |
| Setup gagal karena header | Spreadsheet sudah berisi header lama; gunakan spreadsheet baru atau kosongkan data |
| Stok tidak berubah | Pastikan draft sudah dibalas `OK` dan status draft `POSTED` |
| Formula saldo rusak | Jalankan `repairFormulas()` |

## Keamanan

- Simpan `FONNTE_TOKEN` dan `FONNTE_WEBHOOK_TOKEN` di Script Properties.
- Gunakan webhook token yang panjang dan acak.
- Selalu isi `FONNTE_ALLOWED_NUMBERS`; jangan biarkan kosong di production.
- Batasi akses spreadsheet sesuai kebutuhan. Spreadsheet yang dapat diedit publik memungkinkan perubahan data di luar aplikasi.
- Jangan commit `.clasp.json`, `.env`, token, atau kredensial lain. File tersebut sudah masuk `.gitignore`.
- Audit log membantu melacak perubahan, tetapi bukan pengganti backup spreadsheet.

## Batasan

- Hanya memproses pesan teks direct message.
- Media, gambar, dokumen, poll, dan pesan group diabaikan.
- Tidak ada konversi satuan.
- AI vision belum aktif; `AI_ENABLED` hanya placeholder.
- Spreadsheet ID dikonfigurasi di `src/config.gs`.
- Kapasitas dan quota mengikuti batas Google Apps Script dan Fonnte.
- Fonnte dapat melakukan retry webhook; adapter mengembalikan 200 agar retry tidak menumpuk, sementara error dicatat di log dan Inbox.

## Pengembangan

1. Edit file di `src/`.
2. Jalankan `npm run build`.
3. Jalankan `npm test`.
4. Salin `setup.gs` yang baru ke Apps Script.
5. Deploy versi baru Web App.
6. Commit perubahan source dan `setup.gs` yang dihasilkan.

Untuk perubahan schema, tambahkan migrasi yang kompatibel. `setupSchema()` menolak tab dengan header lama untuk mencegah data rusak.
