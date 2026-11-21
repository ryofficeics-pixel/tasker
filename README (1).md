# Staff Performance Tracking System

Sistem internal Task Assignment, Project Monitoring, dan Staff Performance Tracking dalam **satu file Google Sheets**.

| File | Fungsi |
|---|---|
| `STAFF_PERFORMANCE_TRACKING.xlsx` | Workbook 10 sheet sesuai spesifikasi + 1 sheet `SYSTEM LOG` di paling kanan (catatan email/form/error) |
| `Index.html` | Tampilan web app (HP) admin & staf — ditambahkan ke proyek Apps Script yang sama |
| `Code.gs` | Google Apps Script: Form HP (tugas & progres), Submit Task, Task ID, validasi, email, reminder, overdue, progress log |
| `README.md` | Panduan ini |

**Revisi terbaru (v3):** (1) **Web app HP** (`Index.html`) untuk admin (input tugas, monitoring, penilaian) dan staf (baca penugasan, lapor progres, **upload foto bukti**); Google Sheet tetap jadi database & rekap. (2) Semua pembaruan kini berupa **patch**: link Google Sheet yang ada **tidak berubah** (jangan upload ulang XLSX; lihat bagian 9). (3) Foto bukti tersimpan di Google Drive, link-nya otomatis masuk EVIDENCE LOG.

---

## 1. Upload & konversi ke Google Sheets

1. Buka <https://drive.google.com> > **New > File upload** > pilih `STAFF_PERFORMANCE_TRACKING.xlsx`.
2. Klik kanan file > **Open with > Google Sheets**.
3. **File > Save as Google Sheets** (agar menjadi file Sheets asli, bukan XLSX yang diedit).
4. **File > Settings**: set **Locale = Indonesia** dan **Time zone = (GMT+07:00) Jakarta**. *Wajib* — `NOW()`/`TODAY()` (deteksi overdue) dan Apps Script memakai zona ini.
5. **File > Settings > Calculation**: set *Recalculation* = **On change and every hour** (agar OVERDUE ikut berubah tanpa diedit).
6. Cek cepat setelah konversi:
   - Sheet pertama = `TASK INPUT`.
   - Klik sel di kolom *Project*, *PIC*, *Priority*, *Created By*: dropdown muncul. Jika dropdown hilang di sheet tertentu, pilih kolomnya > **Data > Data validation** > *Dropdown (from a range)* > arahkan ke range yang sama (lihat tabel di bagian 6).
   - `DASHBOARD` menampilkan KPI = 0 dan 3 grafik kosong (database masih kosong — normal).

> Data awal di STAFF MASTER, PROJECT MASTER, dan SETTINGS (Admin) adalah **contoh** (`Staf 01`, `Proyek Contoh A`, `manager@example.com`). Ganti dengan data asli sebelum dipakai.

---

## 2. Instalasi Apps Script

1. Di Google Sheets: **Extensions > Apps Script**.
2. Hapus isi `Code.gs` bawaan, tempel seluruh isi `Code.gs` dari paket ini, **Save**.
3. **Project Settings (ikon roda gigi) > centang "Show appsscript.json manifest file in editor"**. Buka `appsscript.json`, pastikan:
   ```json
   {
     "timeZone": "Asia/Jakarta",
     "runtimeVersion": "V8",
     "exceptionLogging": "STACKDRIVER"
   }
   ```
4. Kembali ke Sheets, **refresh (F5)**. Menu **Task System** muncul di bar menu.
5. Jalankan **Task System > Self Test**. Google meminta izin (Sheets, Gmail/MailApp, Triggers) — izinkan (akan muncul peringatan "app not verified" untuk script milik sendiri: *Advanced > Go to … (unsafe)*).
6. Self Test: `Form HP …` dan `Trigger …` akan `[GAGAL]` sampai Anda menjalankan Setup Form (di atas) dan Install Triggers (bagian 3). Yang lain harus `[OK]`.

### Input via HP (Google Form) — cara utama
Aplikasi Google Sheets di HP **tidak menampilkan menu/tombol Apps Script**, jadi input dari HP memakai Google Form yang dibuat otomatis oleh script:

1. Di laptop: **Task System > Mobile Forms (input via HP) > Setup / Buat Form**. Izinkan akses Google Forms bila diminta.
2. Alert menampilkan 2 link: **Form TUGAS** (khusus admin) dan **Form PROGRES** (staf & admin). Kirim link ke HP masing-masing; di browser HP pilih *Tambahkan ke layar utama* agar jadi ikon.
3. Pengirim harus **login Google dengan email yang terdaftar**:
   - Form Tugas: email harus ada di **SETTINGS > Admin Email**. *Created By* diisi otomatis dari email login (tidak bisa dipalsukan).
   - Form Progres: email staf di **STAFF MASTER** (hanya boleh mengubah tugas yang menjadi miliknya) atau email admin. *Updated By* otomatis dari email.
4. Dropdown (Project, PIC, Priority, daftar tugas terbuka) mengikuti master data. Setelah menambah staf/proyek, jalankan **Mobile Forms > Refresh Pilihan Form** (juga otomatis tiap hari dan setiap ada tugas/progres baru).
5. Input banyak tugas cepat: setelah kirim, ketuk **Kirim tanggapan lain**.
6. Jika kiriman ditolak (bukan admin, duplikat, deadline lewat, dsb.) **alasannya dikirim ke email pengirim** dan dicatat di `SYSTEM LOG`. Jawaban asli tetap tersimpan di Google Form (tab Responses) sebagai arsip.

`TASK INPUT` tetap ada untuk input massal di laptop: isi banyak baris lalu **Task System > Submit Tasks**.

---

## 3. Konfigurasi email & reminder

Semua pengaturan ada di sheet **SETTINGS** (kolom Config Key / Config Value) — tidak perlu mengubah kode.

1. Isi **Admin Name + Admin Email** (kolom C–D) dan **Email** semua staf di STAFF MASTER dengan alamat asli. (Data contoh memakai `@example.com` — Self Test akan menandainya.)
2. `EMAIL_ENABLED = YES` (sekarang default **YES**). `NO` mematikan semua email.
3. **Task System > Send Test Email** — email tes masuk ke akun Anda. Izin *Gmail/MailApp* harus disetujui saat pertama kali.
4. **Task System > Install Triggers** — reminder harian pada jam `REMINDER_HOUR` (default 08:00).
5. **Task System > Self Test** — semua baris harus `[OK]` (kecuali yang belum Anda pakai).

**Jika email tidak masuk, cek berurutan:**
1. Sheet **SYSTEM LOG**: setiap percobaan email tercatat `TERKIRIM ke …` atau `GAGAL ke …: <alasan>`. Tidak ada baris sama sekali = email tidak dicoba (lihat langkah 2).
2. Pesan hasil **Submit Tasks** sekarang menulis jelas: `EMAIL TIDAK DIKIRIM: EMAIL_ENABLED = NO …` atau `Email GAGAL: <nama>: <alasan>`.
3. Penyebab umum: `EMAIL_ENABLED` masih NO; email PIC kosong/salah di STAFF MASTER (hanya PIC yang menerima email penugasan); izin Gmail belum disetujui (jalankan Send Test Email); kuota harian habis; email masuk **Spam**.
4. Email penugasan dikirim ke **PIC**, dengan CC ke admin pembuat bila `EMAIL_CC_ADMINS = YES`. Email penolakan Form HP dikirim ke pengirim form (tidak bergantung `EMAIL_ENABLED`).

| Email | Kapan | Penerima |
|---|---|---|
| Tugas baru | Saat Submit (sheet atau Form HP) | PIC (CC admin pembuat) |
| Progres penting | Status jadi COMPLETED / PENDING / CANCELLED | Admin pembuat tugas |
| Reminder harian | Tugas overdue / ≤ `DUE_SOON_DAYS` hari lagi | PIC + ringkasan ke semua admin |
| Penolakan Form HP | Kiriman ditolak | Pengirim form |

Batas Google: ±100 email/hari (Gmail biasa), ±1.500/hari (Workspace). Kegagalan email **tidak** membatalkan data.

---

## 4. Cara memakai (harian)

### A. Memberi tugas
**Dari HP (utama):** buka Form TUGAS > isi *Project, PIC, Task Title, Priority, Deadline Date* (opsional: Description, Start Date, Parent Task ID, Notes) > Kirim. Tugas langsung masuk TASK DATABASE dengan Task ID, dan PIC menerima email.

**Dari laptop (massal):** sheet `TASK INPUT`, satu tugas per baris mulai baris 5 (baris 3 hanya contoh). Kolom biru = isi manual; abu-abu = otomatis.
1. Wajib: **Created By, Project, PIC, Task Title, Priority, Deadline Date**. Opsional: Input Date (default hari ini), Start Date (default = Input Date), Parent Task ID, Description, Notes.
2. Kolom **Validation Message** langsung menampilkan `SIAP SUBMIT` atau `PERBAIKI: …` (tanpa Apps Script).
3. **Task System > Submit Tasks**. Baris valid masuk database, mendapat Task ID dan status `SUBMITTED` (terkunci). Baris tidak valid **tidak** dikirim dan alasan per baris ditampilkan.
4. Opsional **Clear Submitted Rows** untuk membersihkan baris yang sudah terkirim (data aman di database).

**Deadline:** tidak ada input jam. Deadline = **akhir hari (23:59)** pada tanggal deadline; tugas baru menjadi OVERDUE setelah tanggal itu lewat.

Catatan: *Notes* ditambahkan ke akhir *Task Description* di database (`Catatan: …`). Task ID dibuat **saat Submit** agar unik walau beberapa admin submit bersamaan. Parent Task ID harus tugas yang **sudah ada** di database.

### B. Update progres
**Dari HP (utama):** Form PROGRES > pilih tugas (daftar tugas terbuka) > status, progres, catatan, alasan keterlambatan > Kirim. Staf hanya bisa mengubah tugasnya sendiri; CANCELLED hanya admin. Kiriman yang melanggar aturan ditolak dengan email penjelasan.

**Dari laptop:** di `PROGRESS LOG`, staf/admin menambah **baris baru**: Task ID, Updated By, New Status, Progress (opsional), Update Notes, Delay Reason, Revised Estimated Completion → **Task System > Process Progress Updates**. Script menambah Log ID, Timestamp, Previous Status, lalu memperbarui TASK DATABASE (status, progres, tanggal selesai, alasan, last updated). Baris log yang sudah punya Log ID tidak boleh diubah (riwayat permanen).

Aturan: COMPLETED otomatis 100% dan mengisi Completion Date; progres 100% hanya untuk COMPLETED; PENDING atau tugas lewat deadline wajib Delay Reason (`REQUIRE_DELAY_REASON`); tugas COMPLETED/CANCELLED tidak bisa dibuka lagi (`ALLOW_REOPEN_CLOSED = NO`); tidak bisa kembali ke NEW.

### C. Evidence (EVIDENCE LOG)
Tempel link Google Drive + Task ID + Uploaded By di baris baru > **Process Evidence Log**. Link terbaru juga disalin ke kolom *Evidence Link* di database.

**Upload foto langsung via Google Form (opsional):**
1. Buat Form **manual** (script Google tidak dapat membuat pertanyaan upload file) dengan pertanyaan berjudul persis: `Task ID` (short answer), `Uploaded By`, `Evidence Description`, `Foto / File` (File upload). Tab *Responses* > **Link to Sheets** > pilih spreadsheet ini.
2. **Task System > Install Form Trigger**. Setiap kiriman Form otomatis masuk EVIDENCE LOG, divalidasi, dan menautkan ke tugas.
3. Judul pertanyaan dapat diubah di `CONFIG.FORMS.EVIDENCE` pada Code.gs.

### D. Penilaian admin (TASK DATABASE)
Pada tugas COMPLETED, admin mengisi **Quality Rating** (dropdown) dan, bila terlambat, **Delay Category** (`Staff`, `External Party`, `Instruction Change`). Kualitas **tidak** dihitung otomatis dari jumlah tugas selesai.

### E. Dashboard & laporan
- **DASHBOARD**: filter (basis periode Deadline/Created At, tanggal dari–sampai, proyek, staf) di sel kuning C4:C8; KPI, 3 grafik, daftar overdue, daftar mendekati deadline.
- **STAFF PERFORMANCE**: rekap per staf + skor; filter periode di C4:C5 (kosong = semua).
- **PROJECT REPORT**: rekap per proyek; pilih proyek di C27 untuk daftar pekerjaan belum selesai (`ALL` = semua).

---

## 5. Logika penting

- **Status operasional** (input): NEW, IN PROGRESS, PENDING, COMPLETED, CANCELLED. **OVERDUE tidak diinput** — kolom *Effective Status* di TASK DATABASE menghitungnya otomatis: belum COMPLETED/CANCELLED **dan** Deadline < sekarang (termasuk jam).
- Tugas **COMPLETED tidak pernah OVERDUE**. Kolom *Timeliness* mencatat `ON TIME`, `LATE`, atau `LATE (EXCUSED)`.
- **Keterlambatan staf vs eksternal**: `Delay Category` + tabel di SETTINGS (kolom *Counts Against Staff*). Kategori `NO` (External Party, Instruction Change) = *excused*, tidak mengurangi On-Time Rate yang masuk skor. Kosong = dianggap kesalahan staf.
- **Skor performa** = rata-rata berbobot: Completion Rate, On-Time (disesuaikan), Kualitas (rata-rata rating admin), Konsistensi update. Bobot diubah di SETTINGS > Performance Weights; komponen tanpa data (mis. belum ada rating) otomatis dikeluarkan.
- Dashboard: status New/In Progress/Pending/Completed/Overdue/Cancelled saling lepas (jumlahnya = Total). Completion Rate = Completed ÷ (Total − Cancelled).
- **Tidak ada penghapusan/penimpaan**: script hanya menambah baris di TASK DATABASE / PROGRESS LOG / EVIDENCE LOG. Hindari menghapus baris di sheet tersebut secara manual; opsional jalankan **Apply Protections** (peringatan bila kolom otomatis/database diedit).
- **Anti duplikat Submit** (berlapis): status SUBMITTED dilewati; Task ID yang sudah ada di database tidak dibuat ulang; tugas identik (proyek + PIC + judul + deadline) ditolak; ID dikunci dengan *script lock* sehingga dua admin tidak bisa mendapat ID sama.

---

## 6. Dropdown & range (dinamis)

Dropdown membaca range kapasitas besar, sehingga bertambah otomatis saat data master ditambah (tanpa mengedit validasi):

| Dropdown | Sumber |
|---|---|
| Project (TASK INPUT) | `PROJECT MASTER!H2:H201` — daftar otomatis proyek *Active* |
| PIC | `STAFF MASTER!H2:H101` — daftar otomatis staf *Active* |
| Created By | `SETTINGS!C5:C14` (Admin Name) |
| Priority / Status / Delay Category / Quality Rating | `SETTINGS!A5:A14`, `B5:B14`, `E5:E14`, `G5:G14` |
| Filter dashboard | `STAFF MASTER!I2:I102`, `PROJECT MASTER!I2:I202` (ALL + semua nama) |

Staf/proyek yang diubah menjadi **Inactive** hilang dari dropdown input tetapi riwayatnya tetap tampil. Nama staf & proyek harus **unik** (dicek validasi + warna merah). Named range (`StaffActiveList`, `ProjectActiveList`, `PriorityList`, dst.) juga didefinisikan untuk pemakaian lanjutan.

Kapasitas (dropdown Form maks. 400 tugas terbuka terbaru): 100 staf, 200 proyek, 300 baris input per batch (gunakan *Clear Submitted Rows*), 10.000 tugas di database (rumus bantu otomatis diperluas script setiap 500 baris). Grafik dashboard menampilkan 10 staf/proyek pertama; tabel Staff Performance & Project Report 20 baris pertama.

---

## 7. Keterbatasan (jujur)

| Hal | Penjelasan | Solusi yang disediakan |
|---|---|---|
| XLSX tidak punya tombol/script | Submit, ID unik, email, reminder tidak bisa jalan di file XLSX saja | `Code.gs` + menu **Task System** |
| Menu/tombol script di aplikasi Sheets HP | Tidak didukung Google | Input HP lewat Google Form (bagian 2) |
| Form HP butuh login Google | Identitas pengirim diambil dari email terverifikasi; staf/admin tanpa akun Google tidak bisa memakai form | Daftarkan email Google mereka di STAFF MASTER / SETTINGS |
| Pilihan dropdown Form tidak otomatis real-time | Diperbarui saat ada tugas/progres baru, harian, atau menu Refresh Pilihan Form | Jalankan Refresh setelah mengubah master data |
| Task ID tidak tampil sebelum Submit | Agar unik dan aman dari balapan 2 admin | ID dibuat saat Submit |
| Kunci baris SUBMITTED | `onEdit` membatalkan edit **satu sel**; paste/hapus banyak sel ke baris terkunci hanya diperingatkan (gunakan Ctrl+Z). Database tetap aman karena terpisah dari TASK INPUT | Proteksi opsional (**Apply Protections**) + anti-duplikat |
| Dropdown "Active only" | Daftar dropdown berupa range rumus; Excel/Sheets dapat menampilkan beberapa baris kosong di akhir daftar | Hanya kosmetik |
| Hak akses | Google Sheets tidak bisa membatasi *tampilan* per sel tanpa proteksi; staf yang boleh mengedit sheet bisa mengedit semuanya | Beri staf akses **hanya** lewat Google Form, atau proteksi sheet (Data > Protect sheets and ranges) |
| Zona waktu | Spreadsheet dan script harus sama; jika beda, Submit menolak dengan pesan jelas | Bagian 1 langkah 4 & bagian 2 langkah 3 |
| `NOW()` di Sheets | Dihitung ulang tiap menit/jam sesuai pengaturan Calculation | Bagian 1 langkah 5 |
| Chart | Konversi XLSX→Sheets dapat sedikit mengubah warna/ukuran chart | Edit di Sheets bila perlu; sumber data tetap di sel K4:Q37 DASHBOARD |

---

## 8. Hasil pengujian sebelum rilis

Dilakukan di lingkungan pengembangan (bukan Google):

- **Formula**: workbook dihitung ulang dengan LibreOffice — 17.217 rumus, 0 error. Diuji dengan data 12 tugas (semua status, tepat waktu/terlambat/excused, lewat jam, cancelled): hasil KPI, Staff Performance (termasuk skor), Project Report (termasuk daftar PIC), daftar overdue/mendekati deadline, dan pesan validasi TASK INPUT dicocokkan dengan hitungan manual — semua sesuai.
- **Code.gs**: 145 pemeriksaan otomatis menggunakan tiruan API Google (Task ID unik & berurutan, penambahan banyak tugas, submit ulang tidak menggandakan, pemulihan setelah kegagalan tulis, validasi/pesan error, deadline akhir hari, progress log, evidence, penguncian edit, email & reminder, deteksi overdue, perluasan rumus, pembuatan Form HP, penolakan non-admin/duplikat/bukan-PIC, pencatatan SYSTEM LOG) — semua lulus.
- **Belum teruji**: eksekusi di server Google asli (izin, kuota email, pembuatan Google Form sungguhan, trigger waktu, tampilan dropdown hasil konversi). Karena itu **wajib** menjalankan *Self Test* dan uji coba kecil (3–5 tugas dengan email Anda sendiri) sebelum dipakai tim — lihat bagian 2 dan 3.

### Uji pasca-instalasi (5 menit)
1. Buka Form TUGAS di HP, kirim 1 tugas > cek TASK DATABASE, email PIC, dan SYSTEM LOG.
2. Kirim lagi tugas yang sama persis > ditolak (duplikat), alasan masuk email Anda.
2b. Di laptop: isi 3 baris TASK INPUT (1 sengaja kurang Project) > Submit > 2 masuk, 1 ditolak; Submit lagi > tidak ada duplikat.
3. Form PROGRES di HP (login sebagai staf PIC) > COMPLETED > status database berubah, Dashboard terupdate.
4. Ubah Deadline salah satu tugas di database ke kemarin > *Effective Status* menjadi OVERDUE, muncul di daftar overdue Dashboard.
5. Setelah selesai uji, hapus baris uji **bukan** dengan menghapus di database: gunakan status CANCELLED (riwayat tidak boleh dihapus).

---

## 9. Web app (admin & staf) + foto bukti — PATCH, link Sheet tetap

Tidak perlu hosting terpisah: web app berjalan di proyek Apps Script yang sama, semua data tetap di Google Sheet Anda.

### 9.1 Pasang patch (sekali, ±10 menit)
1. Buka Google Sheet yang sudah ada > **Extensions > Apps Script**.
2. Buka `Code.gs` > **hapus semua isi, tempel isi `Code.gs` baru** > Save.
3. Klik **+** (Files) > **HTML** > beri nama persis `Index` > tempel isi `Index.html` > Save.
4. Di Sheet: muat ulang halaman (F5). Menu **Task System** > **Apply Patch** > Yes. Patch otomatis:
   - membuat salinan cadangan Sheet di Google Drive Anda (nama: `... BACKUP ...`),
   - menyesuaikan layout lama (mis. menghapus kolom *Deadline Time* bila masih ada, membuat `SYSTEM LOG`, kolom *Access Code*, konfigurasi web app),
   - aman dijalankan berulang (langkah yang sudah beres dilewati). Data tugas/progres/evidence **tidak** diubah atau dihapus.
5. Menu **Task System > Web App > Generate Access Codes**. Kode ada di `STAFF MASTER` kolom F (staf) dan `SETTINGS` kolom L (admin). Bagikan kode ke tiap orang secara pribadi. Ganti kode = kosongkan sel, jalankan menu lagi.
6. Apps Script > **Deploy > New deployment > Web app** > *Execute as*: **Me**; *Who has access*: **Anyone** > Deploy > izinkan akses (Drive, email). Salin link.
7. Menu **Task System > Web App > Link Web App** untuk menampilkan link. Bagikan link + kode. Di HP: buka link > menu browser > **Tambahkan ke layar utama**.
8. Menu **Self Test** harus semua lulus.

### 9.2 Pembaruan berikutnya
Tempel `Code.gs` / `Index.html` baru di editor yang sama > Save > **Deploy > Manage deployments > Edit (pensil) > Version: New version > Deploy** (link web app tetap sama) > di Sheet: **Apply Patch** bila diminta Self Test. Jangan membuat deployment baru dan jangan upload ulang XLSX.

### 9.3 Cara pakai
- **Staf**: login kode > *Tugas Saya* > ketuk tugas > pilih status/progres/catatan > **Ambil / pilih foto** (kamera HP; otomatis dikecilkan) > Kirim. Foto bisa ditambah kapan saja, juga setelah COMPLETED. Tab *Laporan*: skor & statistik pribadi.
- **Admin**: *Monitor* (KPI, overdue, mendekati deadline, kinerja staf, proyek), *Tugas* (semua tugas, filter), *Input* (satu atau banyak tugas sekaligus), *Laporan* (per staf), ketuk tugas untuk menilai (Delay Category & Quality) atau mengubah status (termasuk CANCELLED).
- Semua yang dikirim lewat web masuk ke sheet yang sama (TASK DATABASE, PROGRESS LOG, EVIDENCE LOG, SYSTEM LOG) dan memicu email seperti jalur lain. Task ID, anti-duplikat, validasi, dan aturan overdue identik.

### 9.4 Foto bukti
- Disimpan di Drive pemilik script: folder `... Evidence` > subfolder per Task ID. Maks. `MAX_PHOTOS_PER_TASK` (default 20) per tugas.
- `EVIDENCE_PUBLIC_LINK = YES` (default): siapa pun yang punya link foto bisa melihat (link sulit ditebak). `NO`: hanya akun yang diberi akses Drive.
- `REQUIRE_EVIDENCE_ON_COMPLETE = YES`: tugas tidak bisa COMPLETED tanpa minimal 1 foto (hanya jalur web).

### 9.5 Keamanan & keterbatasan web app
- Login memakai **kode akses** (bukan akun Google), jadi staf tanpa akun Google tetap bisa. Siapa pun yang memegang kode bisa bertindak sebagai orang itu; staf hanya melihat/ubah tugasnya sendiri, admin semuanya. Kode tersimpan di Sheet: batasi akses edit Sheet ke admin saja. Menonaktifkan staf (Inactive) langsung mencabut aksesnya; `WEB_ENABLED = NO` menutup seluruh web app.
- Fungsi menu (Submit, Install Triggers, dsb.) menolak dipanggil dari web app.
- Email dari web app dikirim atas nama pemilik script; kuota email harian Google berlaku.
- Web app butuh koneksi internet; tampilan memakai data real-time dari Sheet (lambat bila database >5.000 tugas).
- **Belum teruji di server Google asli** (deploy, izin Drive, kamera HP, kuota). Lakukan uji kecil: 1 staf + 1 foto + 1 tugas sebelum dipakai tim.

### 9.5b Tampilan (UI)
Seluruh tampilan (staf **dan** admin) memakai UI `karyawan.html` dari aplikasi ICS Absensi (`ryofficeics-pixel/ICS-absensi`): tema kertas terang, header TASK ICS + jam, kartu profil, tab segmen, tombol hitam Syne, toast atas. Admin memakai tema yang sama dengan lebar layar lebih besar di desktop. Font Syne / DM Sans / IBM Plex Mono dari Google Fonts (perlu internet). Logika absensi (GPS/kamera/Firebase) tidak ikut — hanya tampilan.

### 9.6 Hasil pengujian (lingkungan pengembangan)
270 pemeriksaan otomatis lulus: patch dari layout lama (v1) dan v2, idempotensi, backup, layout asing ditolak, kode akses, otorisasi staf/admin, upload foto (tipe/ukuran/ganda/batas), progres via web, penilaian, anti-duplikat, angka analitik web = hitungan LibreOffice pada workbook demo. Tampilan HP (390 px) diuji dengan browser otomatis: login, monitor, daftar, form input, update + foto — tanpa error JS.
