# Staff Performance Tracking System

Sistem internal Task Assignment, Project Monitoring, dan Staff Performance Tracking dalam **satu file Google Sheets**.

| File | Fungsi |
|---|---|
| `STAFF_PERFORMANCE_TRACKING.xlsx` | Workbook 10 sheet (input, database, log, master, dashboard, laporan) |
| `Code.gs` | Google Apps Script: Submit Task, Task ID, validasi, email, reminder, overdue, progress log |
| `README.md` | Panduan ini |

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
6. Self Test harus menampilkan semua `[OK]`. Hanya `Trigger harian terpasang` yang boleh `[GAGAL]` sampai langkah 4 di bagian 3.

### Tombol SUBMIT TASK (opsional)
Menu **Task System > Submit Tasks** sudah berfungsi. Untuk tombol di sheet: pada `TASK INPUT` pilih **Insert > Drawing**, buat kotak "SUBMIT TASK", **Save and close**, klik kotak > titik tiga > **Assign script** > ketik `submitTasks`.

---

## 3. Konfigurasi email & reminder

Semua pengaturan ada di sheet **SETTINGS** (kolom Config Key / Config Value) — tidak perlu mengubah kode.

1. Isi **Admin Name + Admin Email** (kolom C–D) dan **Email** semua staf di STAFF MASTER dengan alamat asli.
2. Set `EMAIL_ENABLED = YES` (default **NO** agar tidak mengirim email ke alamat contoh).
3. **Task System > Send Test Email** — pastikan email tes masuk ke akun Anda.
4. **Task System > Install Triggers** — memasang reminder harian pada jam `REMINDER_HOUR` (default 08:00).
5. Cek: Self Test > `Trigger harian terpasang` = OK. Uji manual: **Run Daily Reminder Now**.

| Email | Kapan | Penerima |
|---|---|---|
| Tugas baru | Saat Submit | PIC (CC admin pembuat bila `EMAIL_CC_ADMINS = YES`) |
| Progres penting | Status jadi COMPLETED / PENDING / CANCELLED | Admin pembuat tugas |
| Reminder harian | Tugas overdue / ≤ `DUE_SOON_DAYS` hari lagi | PIC + ringkasan ke semua admin |

Batas Google: ±100 email/hari (akun Gmail biasa), ±1.500/hari (Workspace). Email dikelompokkan per PIC untuk menghemat kuota. Kegagalan email **tidak** membatalkan data; dilaporkan di pesan hasil Submit.

---

## 4. Cara memakai (harian)

### A. Memberi tugas (TASK INPUT)
1. Isi satu tugas per baris mulai baris 5 (baris 3 hanya contoh). Kolom biru = isi manual; abu-abu = otomatis.
2. Wajib: **Created By, Project, PIC, Task Title, Priority, Deadline Date**. Opsional: Input Date (default hari ini), Start Date (default = Input Date), Deadline Time (default `DEFAULT_DEADLINE_TIME`, 17:00), Parent Task ID, Description, Notes.
3. Kolom **Validation Message** langsung menampilkan `SIAP SUBMIT` atau `PERBAIKI: …` (field yang kurang/salah) — tanpa Apps Script.
4. Klik **Task System > Submit Tasks**. Baris valid masuk ke TASK DATABASE, mendapat Task ID (`TSK-00001`, …) dan status `SUBMITTED` (terkunci). Baris tidak valid **tidak** dikirim dan alasan per baris ditampilkan.
5. Opsional **Clear Submitted Rows** untuk membersihkan baris yang sudah terkirim (data aman di database).

Catatan: kolom *Notes* ditambahkan ke akhir *Task Description* di database (`Catatan: …`). Task ID dibuat **saat Submit** (bukan saat mengetik) agar dijamin unik walau 2 admin submit bersamaan. Parent Task ID harus tugas yang **sudah ada** di database (kirim tugas induk lebih dulu).

### B. Update progres (PROGRESS LOG)
Staf/admin menambah **baris baru**: Task ID, Updated By, New Status, Progress (opsional), Update Notes, Delay Reason, Revised Estimated Completion → **Task System > Process Progress Updates**. Script menambah Log ID, Timestamp, Previous Status, lalu memperbarui TASK DATABASE (status, progres, tanggal selesai, alasan, last updated). Baris log yang sudah punya Log ID tidak boleh diubah (riwayat permanen).

Aturan: COMPLETED otomatis 100% dan mengisi Completion Date; progres 100% hanya untuk COMPLETED; PENDING atau tugas lewat deadline wajib Delay Reason (`REQUIRE_DELAY_REASON`); tugas COMPLETED/CANCELLED tidak bisa dibuka lagi (`ALLOW_REOPEN_CLOSED = NO`); tidak bisa kembali ke NEW.

### C. Evidence (EVIDENCE LOG)
Tempel link Google Drive + Task ID + Uploaded By di baris baru > **Process Evidence Log**. Link terbaru juga disalin ke kolom *Evidence Link* di database.

**Upload foto langsung via Google Form (opsional):**
1. Buat Form dengan pertanyaan berjudul persis: `Task ID` (short answer), `Uploaded By`, `Evidence Description`, `Foto / File` (File upload). Tab *Responses* > **Link to Sheets** > pilih spreadsheet ini.
2. **Task System > Install Form Trigger**. Setiap kiriman Form otomatis masuk EVIDENCE LOG, divalidasi, dan menautkan ke tugas.
3. Form progres (opsional): judul pertanyaan `Task ID`, `Updated By`, `New Status`, `Progress Percentage`, `Update Notes`, `Delay Reason`, `Revised Estimated Completion` — diproses dengan trigger yang sama. Judul pertanyaan bisa diubah di `CONFIG.FORMS` pada Code.gs.

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

Kapasitas: 100 staf, 200 proyek, 300 baris input per batch (gunakan *Clear Submitted Rows*), 10.000 tugas di database (rumus bantu otomatis diperluas script setiap 500 baris). Grafik dashboard menampilkan 10 staf/proyek pertama; tabel Staff Performance & Project Report 20 baris pertama.

---

## 7. Keterbatasan (jujur)

| Hal | Penjelasan | Solusi yang disediakan |
|---|---|---|
| XLSX tidak punya tombol/script | Submit, ID unik, email, reminder tidak bisa jalan di file XLSX saja | `Code.gs` + menu **Task System** |
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
- **Code.gs**: 98 pemeriksaan otomatis menggunakan tiruan API Google (Task ID unik & berurutan, penambahan banyak tugas, submit ulang tidak menggandakan, pemulihan setelah kegagalan tulis, validasi/pesan error, progress log, evidence, penguncian edit, email & reminder, deteksi overdue, perluasan rumus) — semua lulus.
- **Belum teruji**: eksekusi di server Google asli (izin, kuota email, trigger waktu, tampilan dropdown hasil konversi). Karena itu **wajib** menjalankan *Self Test* dan uji coba kecil (3–5 tugas dengan email Anda sendiri) sebelum dipakai tim — lihat bagian 2 dan 3.

### Uji pasca-instalasi (5 menit)
1. Isi 3 baris TASK INPUT (1 sengaja kurang Project) > Submit > 2 masuk, 1 ditolak dengan alasan.
2. Klik Submit lagi > tidak ada tugas baru, jumlah baris database tetap.
3. Tambah baris PROGRESS LOG (COMPLETED) > Process Progress Updates > status di database berubah, Dashboard terupdate.
4. Ubah Deadline salah satu tugas di database ke kemarin > *Effective Status* menjadi OVERDUE, muncul di daftar overdue Dashboard.
5. Setelah selesai uji, hapus baris uji **bukan** dengan menghapus di database: gunakan status CANCELLED (riwayat tidak boleh dihapus).
