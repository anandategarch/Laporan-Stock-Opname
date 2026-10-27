# PANDUAN MIGRASI — Perhitungan Berat ke Data Model
## Laporan Stock Opname (Full Power Query → Power Query + DAX)

> **📦 STATUS FILE (UPDATE v6 — PIVOT HASIL SO: + PERHITUNGAN COM):**
> file xlsx di repo ini SUDAH DIREVISI LANGSUNG.
> Pembaruan v6: `F_HasilSO` kini juga MEMBUAT baris hitung
> **`6.PEMAKAIAN (COM)`** — replika rumus query produksi:
> *pemakaian = (SO kemarin | STOK AWAL untuk hari-1) + Barang Masuk −
> Barang Keluar − SO hari ini*, dirantai per item urut hari (pola sama
> dengan rantai SO_KEMARIN `F_Deviasi` — hari-1 memakai STOK AWAL, bug
> null-hari lama tidak ikut). Pilih `6.PEMAKAIAN (COM)` di filter
> `F_HasilSO[JENIS]` pivot untuk menampilkannya. **8 JENIS mentah tidak
> berubah** (37.169 sel nilai tetap 100% persis vs cache lama).
> Validasi v6: 3.131 baris COM (101 item berkode × 31 hari); identitas
> aljabar **COM = BOM − 1.DEVIASI** dicek LANGSUNG ke angka asli cache
> CEK DEVIASI lama: **3.091 cocok persis + 9 known-fix hari-1** (koreksi
> bug lama yang sama seperti `F_Deviasi`), 0 anomali
> (`so-surgery/updates/validate_fhasilso.py` CEK A–E). Struktur workbook
> **65/65 LULUS**. File Anda sudah termigrasi? Ikuti "UPGRADE" Sub-B
> (paste ulang `F_HasilSO`, ± 3 menit).
>
> **📦 (UPDATE v5 — PIVOT HASIL SO: PARITAS FILTER "JENIS"):**
> file xlsx di repo ini SUDAH DIREVISI LANGSUNG.
> Pembaruan v5: `F_HasilSO` kini memuat **SEMUA 8 JENIS** (SO, STOCK AWAL,
> BARANG MASUK, BARANG KELUAR, BOM, WASTE, SUSUT, TRIAL) **plus kolom
> `JENIS`**. Ground truth file sampel: tabel query lama memuat semua 8 jenis
> dan memfilternya **di pivot** — dropdown "JENIS" (default **SO**) di
> pojok kiri sheet PIVOT HASIL SO (2), bisa diganti interaktif. Migrasi
> v1–v4 mengeraskan slice `JENIS="SO"` di query → dropdown itu hilang;
> v5 mengembalikan paritas penuh: seret `F_HasilSO[JENIS]` ke **FILTERS**
> pivot dan pilih "SO" (Langkah 5).
> Validasi v5 vs cache pivot lama: **37.169 sel nilai (1.199 baris × 31 hari)
> 100% persis**, subset SO tidak berubah 1 byte pun nilainya, 8 JENIS lengkap,
> 43 query lain byte-identical (`so-surgery/updates/validate_fhasilso.py`).
> File Anda sudah termigrasi v1–v4? Ikuti bagian "UPGRADE" sub-B (± 3 menit).
>
> **PEMBARUAN v4 — PANEL QUERY DIRAPIKAN: 11 GRUP ACAK → 6 GRUP BERNOMOR:**
> seluruh **44 query dipindah ke 6 folder bernomor** di panel
> Queries & Connections (lihat seksi "STRUKTUR PANEL QUERY (v4)" di bawah),
> folder kosong `INPUT` dan 4 folder sisa (TRIAL/WASTE/SUSUT/DATABASE MENU) dihapus,
> dan 3 query yang tadinya *menggantung tanpa folder* (`RFC_Database_Susut`,
> `RFC_Database_Trial`, `RFC_Database_Waste_BnP`) kini masuk folder `2. TEMPLATE INPUT`.
> **Audit dependency: 44/44 query masih dipakai** — 3 query yang memang tidak
> dibutuhkan sudah dihapus pada v1 (lihat tabel RINGKASAN), sisanya wajib ada
> (13 pengisi template sheet input, 10 fungsi, 24 pipeline + output). M (Section1.m)
> **tidak berubah satu byte** dari v3.
>
> **Pembaruan v2: `F_Deviasi` kini ber-18 kolom**
> (replika persis output query lama `RFC_Rekap Mentah [CEK DEVIASI] Raw`), sehingga
> **double-click / Show Details di pivot CEK DEVIASI HARIAN menampilkan 18 kolom
> yang sama persis dengan file sampel** — KATEGORI, NAMA BARANG, JENIS, SATUAN,
> Attribute, SO_KEMARIN, WASTE, TRIAL, SUSUT, STOCK AWAL, BARANG MASUK,
> BARANG KELUAR, BOM, Stock Seharusnya (SO_Kemarin + Barang Masuk - Barang
> Keluar - BOM), SO, Value, Keterangan, Indikasi (urutan identik dengan pivot
> cache lama — terverifikasi otomatis).
> **Langkah 1 & Langkah 3 SUDAH diterapkan ke file** (3 query lama dihapus dari
> DataMashup + 15 part pivot lama dibersihkan + 3 sheet pivot dikosongkan +
> 3 query baru `F_Deviasi` / `F_HasilSO` / `Metric` tertanam).
> **Yang tinggal Anda lakukan: Langkah 2 (Load To → Data Model), Langkah 4
> (1 measure saja!), Langkah 5 (bangun 3 pivot), Langkah 6 (refresh & validasi)**
> — total ± 15 menit klik-klik UI. (Model Data / measure DAX tidak bisa disuntik
> dari luar Excel — format `model.item` bersifat proprietary.)
> Verifikasi struktur file pasca-revisi: **65/65 pemeriksaan otomatis LULUS**
> (ZIP/CRC, rels graph, content-types, metadata PQ, konsistensi koneksi↔query,
> Section1.m byte-exact, blob QueryGroups 6 grup, 44/44 query ber-grup).
>
> **Status validasi angka: LULUS 100%** — 15.765/15.810 baris CEK DEVIASI,
> 15.252/15.810 AKUMULASI, 1.199/1.199 PIVOT HASIL SO (v5: 37.169 sel nilai
> persis, semua 8 JENIS), dan seluruh teks
> Keterangan/Indikasi cocok persis dengan cache pivot lama. Sisa selisih =
> **45 + 558 + 14 sel *known-fix*** = koreksi bug lama hari-1 (detail di
> `LAPORAN-VALIDASI.md`, bukan masalah baru).
>
> **Riwayat pembaruan:** v1 migrasi awal (hapus 3 query output, tanam 3 tabel Data
> Model) · v2 `F_Deviasi` 18 kolom (paritas double-click) · v3 `Table.Buffer`
> single-pass (kecepatan, angka identik) · v4 panel query 6 folder bernomor ·
> v5 `F_HasilSO` semua 8 JENIS + kolom JENIS (paritas dropdown "JENIS") ·
> **v6 `F_HasilSO` + baris hitung `6.PEMAKAIAN (COM)` (rumus produksi,
> tervalidasi vs cache lama)**.
> Sudah termigrasi pakai panduan v1–v4? Tidak perlu download ulang — cukup ikuti
> bagian "UPGRADE FILE YANG SUDAH TERMIGRASI (± 10 MENIT)".

---

## RINGKASAN PERUBAHAN QUERY (44 total)

| # | Query | Aksi |
|---|-------|------|
| 1 | `RFC_Rekap Mentah [CEK DEVIASI] Raw` | ❌ HAPUS (diganti `F_Deviasi` 18 kolom) |
| 2 | `RFC_Rekap Mentah [CEK DEVIASI] Akumulasi` | ❌ HAPUS (running total = 1 measure DAX) |
| 3 | `Raw Rekap Mentah [FINAL] PIVOT HASIL SO` | ❌ HAPUS (diganti `F_HasilSO`) |
| 4 | `F_Deviasi` | ➕ BARU — fakta 18 kolom (CEK DEVIASI & AKUMULASI, drillthrough = sampel) |
| 5 | `F_HasilSO` | ➕ BARU — fakta PIVOT HASIL SO: **semua 8 JENIS + kolom `JENIS` + baris hitung `6.PEMAKAIAN (COM)` (v6)**; filter jenis ada **di pivot** (FILTERS, default SO) — persis sampel (v5) |
| 6 | `Metric` | ➕ BARU — dimensi 5 metrik (kini **opsional** — `F_Deviasi[JENIS]` sudah memuat label metrik) |
| 7 | `Raw Rekap Mentah` | ⚪ TETAP connection-only (tidak ke model) |
| 8 | `RFC_Kode Barang` | ⚪ TETAP connection-only (dipakai `F_HasilSO`) |
| 9 | `Raw Rekap Mentah [FINAL]` | ⚪ TETAP load ke sheet REKAP MENTAH |
| 10 | Semua `SRC/STG/IPT/RFC_*/fn_*` lain | ⚪ TIDAK DISENTUH (byte-identical) |

**Tidak perlu relasi apa pun di Power Pivot.** Model = 2–3 tabel berdiri sendiri.

---

## STRUKTUR PANEL QUERY (v4 — 6 FOLDER BERNOMOR)

Panel **Data → Queries & Connections** kini tertata mengikuti alur data — buka file,
dan panel kanan menampilkan:

| Folder | Isi | Fungsi |
|---|---|---|
| **📁 1. MASTER & CONFIG** (2) | `SRC_AllSources`, `CFG_InputHarian` | Hub seluruh data master (Database Konversi, BOM, menu, mapping) + konfigurasi kolom input harian |
| **📁 2. TEMPLATE INPUT** (13) | 5× `RFC_Database Menu …`, 4× `RFC_Nama Item …`, `RFC_Kode Barang`, `RFC_Database_Susut/Trial/Waste_BnP` | Menulis ulang daftar item/menu di sheet input saat refresh (jangan dihapus — tanpa ini sheet input tidak ter-update) |
| **📁 3. STAGING (UNPIVOT)** (7) | 7× `STG_*_Unpivot` | Ubentuk kolom hari 1..31 menjadi baris (NAMA BARANG, Attribute, Qty) |
| **📁 4. TRANSFORM (IPT)** (7) | 7× `IPT_*` | Konversi satuan → gr, join LINK/konversi, group + label JENIS |
| **📁 5. OUTPUT** (5) | `Raw Rekap Mentah`, `Raw Rekap Mentah [FINAL]`, `F_Deviasi`, `F_HasilSO`, `Metric` | Sheet REKAP MENTAH + 3 tabel Data Model untuk pivot baru |
| **📁 6. FUNCTION** (10) | 10× `fn*` | Fungsi bantu yang dipanggil query lain (tidak perlu dibuka) |

Pemetaan folder lama → baru: `CONFIG`→**1**; `DATABASE SO` + `DATABASE MENU` + 3 query root→**2**; `STG`→**3**;
`SO` + `SUSUT` + `TRIAL` + `WASTE`→**4**; `REKAP MENTAH`→**5**; `Function`→**6**;
folder kosong `INPUT` dihapus.

> **Kenapa tidak ada query yang dihapus lagi di v4?** Audit dependency membuktikan
> ke-44 query yang tersisa semuanya dipakai: 13 `RFC_*` mengisi template sheet input
> (FillEnabled), 10 `fn*` dipanggil query lain, 7 `STG_*` + 7 `IPT_*` + `Raw Rekap
> Mentah` + `RFC_Kode Barang` + `SRC_AllSources` + `CFG_InputHarian` direferensikan
> query lain di rantai M, dan 5 output punya tujuan load. **Menghapus salah satunya
> = rantai refresh putus / sheet input berhenti ter-update.** Tiga query yang memang
> sudah tidak dibutuhkan (`[CEK DEVIASI] Raw`, `[CEK DEVIASI] Akumulasi`,
> `[FINAL] PIVOT HASIL SO`) telah dihapus sejak v1 — itulah satu-satunya query yang
> aman dibuang. Bukti audit per query ada di `LAPORAN-VALIDASI.md` Lampiran 4.

---

## PERSIAPAN (± 5 menit)

1. **BACKUP dulu** — copy file xlsx ke `Laporan SO - backup pre-migrasi.xlsx`. WAJIB.
2. Aktifkan Power Pivot: **File → Options → Add-ins → COM Add-ins → Go →**
   centang **Microsoft Power Pivot for Excel** → OK.
3. Buka panel query: **Data → Queries & Connections** (panel kanan).

---

## LANGKAH 1 — Buat 3 query baru (± 5 menit)

Untuk setiap query: **Data → Get Data → From Other Sources → Blank Query →**
**Advanced Editor → paste kode (tanpa baris `shared ...` dan `;` penutup blok) →**
**OK → rename query di panel Queries & Connections.**

Kode lengkap ada di **`Section1_Revisi.m`** (blok `F_Deviasi`, `F_HasilSO`, `Metric`).
Cara cepat: buka file itu, cari `shared F_Deviasi = let`, copy dari `let` sampai `in Typed ;`.

> **Catatan:** kode `F_Deviasi` v2 cukup panjang (± 320 baris) karena mereplikasi
> pembentukan 18 kolom query lama (pivot JENIS → SO_KEMARIN → 5 metrik →
> unpivot → Keterangan/Indikasi). Paste utuh saja — tidak ada yang perlu diedit.

`F_HasilSO` dan `Metric` — ambil dari `Section1_Revisi.m` (blok pendek).

## LANGKAH 2 — Atur tujuan load (± 2 menit)

Klik kanan query di panel → **Load To…** → pilih **Only Create Connection** →
**CENTANG "Add this data to the Data Model"** → OK. Terapkan pada:

| Query | Setting |
|-------|---------|
| `F_Deviasi` | Connection only + ✅ Data Model |
| `F_HasilSO` | Connection only + ✅ Data Model |
| `Metric` | Connection only + ✅ Data Model *(opsional — boleh dilewati)* |

**JANGAN ubah** load query lain (RFC template, IPT, STG, [FINAL], dst.)

## LANGKAH 3 — Hapus 3 query lama (± 1 menit)

Klik kanan → **Delete** pada:
1. `RFC_Rekap Mentah [CEK DEVIASI] Raw`
2. `RFC_Rekap Mentah [CEK DEVIASI] Akumulasi`
3. `Raw Rekap Mentah [FINAL] PIVOT HASIL SO`

> 3 pivot lama (CEK DEVIASI HARIAN, AKUMULASI, PIVOT HASIL SO (2)) akan error/kosong —
> **normal**, kita bangun ulang di Langkah 5. Jangan refresh dulu.

## LANGKAH 4 — Buat 1 measure DAX (± 2 menit)

Hanya **SATU** measure yang diperlukan (untuk pivot AKUMULASI).
Buka **`DAX-Measures.dax`** → **Power Pivot tab → Measures → New Measure** →
Table name: `F_Deviasi` → paste rumus → Measure name: `Akumulasi` →
Format: **Number, 0 desimal, ribuan**.

```dax
VAR d =
    MAX ( F_Deviasi[Attribute] )
RETURN
    CALCULATE (
        SUM ( F_Deviasi[Value] ),
        FILTER ( ALL ( F_Deviasi[Attribute] ), F_Deviasi[Attribute] <= d )
    )
```

> Pivot **CEK DEVIASI HARIAN tidak butuh measure sama sekali** — semua nilai
> (5 metrik) sudah jadi kolom `Value` per baris metrik, sama seperti pivot sampel.
> Measure generasi-1 lama (20 measure) tidak diperlukan lagi — lihat catatan di
> `DAX-Measures.dax` bagian 3.

## LANGKAH 5 — Bangun 3 pivot baru dari Data Model (± 5 menit)

Untuk tiap sheet: buka sheet lama, hapus isi area pivot lama, lalu
**Insert → PivotTable → From Data Model**.

**CEK DEVIASI HARIAN** (baru — layout persis pivot sampel):
- ROWS:    `F_Deviasi[KATEGORI]` → `F_Deviasi[NAMA BARANG]` → `F_Deviasi[JENIS]` → `F_Deviasi[SATUAN]`
- COLUMNS: `F_Deviasi[Attribute]`  (angka 1..31, terurut numerik otomatis)
- VALUES:  `F_Deviasi[Value]` (Sum of Value)
- ✅ **DOUBLE-CLICK sel nilai apa pun** → sheet detail muncul dengan **18 kolom
  persis file sampel** (baris = item × hari × metrik). Ini perilaku drillthrough
  bawaan pivot Data Model: kolom yang ditampilkan = kolom tabel sesuai urutannya.

**AKUMULASI** (baru):
- ROWS / COLUMNS: sama dengan di atas
- VALUES:  `[Akumulasi]` (measure Langkah 4)

**PIVOT HASIL SO (2)** (baru — layout persis pivot sampel):
- FILTERS: `F_HasilSO[JENIS]` = **SO** — persis dropdown "JENIS" di file
  sampel (sel A1/B1 sheet lama). Boleh diganti ke jenis lain (STOCK AWAL,
  BARANG MASUK, dst.) — angka mengikuti, 100% storage engine (cepat).
  Termasuk **`6.PEMAKAIAN (COM)`** (v6): pemakaian harian = SO kemarin +
  Barang Masuk − Barang Keluar − SO hari ini (hari-1 memakai STOK AWAL) —
  sudah dihitung di query, tinggal pilih di filter/slicer.
  Alternatif tampilan: taruh `JENIS` di **Slicer**.
- ROWS:    `F_HasilSO[KODE BARANG]` → `F_HasilSO[NAMA BARANG]` → `F_HasilSO[SATUAN]`
- COLUMNS: `F_HasilSO[Attribute]`
- VALUES:  `F_HasilSO[Value]` (Sum) — tidak perlu measure
- ⚠️ **WAJIB pasang filter JENIS**: tanpa itu, Sum of Value menjumlahkan
  SEMUA jenis sekaligus (SO + MASUK + KELUAR + …) — tidak bermakna.

## LANGKAH 6 — Refresh & validasi angka (± 5 menit)

1. **Data → Refresh All** (2×; refresh pertama membangun model).
2. Bandingkan angka pivot baru vs laporan validasi (`LAPORAN-VALIDASI.md`), contoh kunci:
   - `SENDOK TEH SUNDA` 1.DEVIASI hari 3 = **817**
   - `LPG 12KG` SO hari 1–3 = sesuai REKAP MENTAH
   - PIVOT HASIL SO `GORENG MEKAR` hari 1 = 272, dst.
   - Ganti filter `JENIS` ke **`6.PEMAKAIAN (COM)`** → cek silang cepat:
     `RICE WRAP` hari 1 = **1.500** (stok awal 1.500, SO/BM/BK = 0);
     `SENDOK TEH SUNDAE (V.030)` hari 1 = **872** (15.540 stok awal + BM −
     BK − SO, identik BOM − 1.DEVIASI = 904 − 32).
3. **Tes double-click**: klik sel mana pun di CEK DEVIASI HARIAN (mis. satu item ×
   hari 5) → sheet detail menampilkan 18 kolom persis daftar di atas. Sel daun
   (item × hari × metrik) = 1 baris; subtotal item = ± 5–160 baris.
   ⚠️ Excel membatasi jumlah baris drillthrough total (grand total bisa terpotong) —
   untuk detail penuh, double-click sel yang lebih spesifik.
4. ⚠️ **Angka hari-1 untuk 9 item akan BERBEDA dari pivot lama — itu koreksi bug, bukan
   error** (daftar item & angka before/after di `LAPORAN-VALIDASI.md` bagian BUG LAMA).
   Contoh: `ADONAN GORENG MEKAR` hari-1 deviasi: lama `+174.981` (salah, stok awal
   diabaikan) → baru `−11.512` (benar, 136.063 − (186.493 + 0 − 0 − 38.918)).
5. Jika semua cocok: hapus sisa pivot lama + sheet `Detail1`, save.

---

## UPGRADE FILE YANG SUDAH TERMIGRASI (± 10 MENIT)

Pakai bagian ini bila file kerja Anda termigrasi dengan panduan versi lama
(v1–v4). **Tidak perlu download ulang file repo** — cukup ganti isi 2 query
lewat Advanced Editor:

- **Sub-A `F_Deviasi`** — WAJIB bila masih versi v1 (tanpa 18 kolom / tanpa
  `Table.Buffer` single-pass v3).
- **Sub-B `F_HasilSO`** — WAJIB bila masih versi v1–v4 (tabel SO-only): ini
  pembaruan **v5** (semua 8 JENIS + kolom `JENIS`) sekaligus **v6** (baris
  hitung `6.PEMAKAIAN (COM)`). Bila file Anda sudah v5 dan hanya ingin
  menambah COM — tetap paste ulang Sub-B ini (satu blok yang sama).

### Sub-A — Ganti `F_Deviasi` (lewati bila sudah v3+)

> Catatan v3: bila F_Deviasi Anda masih versi tanpa langkah
> `#"Buffered Input" = Table.Buffer(#"Removed Columns")` sebelum
> `Table.Pivot` — paste versi terbaru dari `Section1_Revisi.m` untuk
> refresh single-pass (lebih cepat, angka tetap identik).

1. **Data → Queries & Connections** → klik kanan **`F_Deviasi`** → **Edit…**
   (Power Query Editor terbuka).
2. **Home → Advanced Editor** → select semua kode → **paste kode `F_Deviasi`
   baru** (blok `let … in Typed;` dari `Section1_Revisi.m`, mulai dari `let`
   sampai `in\n    Typed` — tanpa baris `shared F_Deviasi =` dan `;` penutup) → **OK**.
   → Preview harus tampil 18 kolom. → **Home → Close & Load** (tujuan load
   "connection + Data Model" tidak berubah).
3. **Data → Refresh All** (2×). Tabel `F_Deviasi` di model otomatis dapat
   kolom-kolom baru; kolom lama (`KATEGORI`, `NAMA BARANG`, `JENIS`, `SATUAN`,
   `Attribute`, `Value`) tetap — tidak ada yang putus.

### Sub-B — Ganti `F_HasilSO` (v6 — paritas filter "JENIS" + baris hitung COM)

4. **Data → Queries & Connections** → klik kanan **`F_HasilSO`** → **Edit…** →
   **Advanced Editor** → select semua → paste kode `F_HasilSO` terbaru dari
   `Section1_Revisi.m` (mulai `let` sampai `in\n    Combined`) → **OK**.
   → Preview: kini ada kolom **JENIS**, baris BUKAN-SO ikut ter-load
   (8 jenis: SO, STOCK AWAL, BARANG MASUK, BARANG KELUAR, BOM, WASTE,
   SUSUT, TRIAL), **plus baris `JENIS = "6.PEMAKAIAN (COM)"`** hasil
   hitung COM (3.131 baris di data sampel). → **Close & Load** (tujuan
   load tidak berubah).
5. **Data → Refresh All** (2×).

### Sub-C — Rapikan pivot (setelah refresh Sub-A/B)

6. **Pivot CEK DEVIASI HARIAN** — rapikan field list:
   - ROWS: hapus `Metric[Metric]`, ganti dengan `F_Deviasi[JENIS]`
     (urutan: KATEGORI → NAMA BARANG → JENIS → SATUAN).
   - VALUES: hapus `[Nilai Metric]`, ganti `F_Deviasi[Value]` (Sum of Value).
7. **Pivot AKUMULASI** — VALUES: hapus `[Akumulasi Metric]`, ganti measure
   `[Akumulasi]` (buat dulu bila belum ada — rumus di Langkah 4); ROWS idem langkah 6.
8. **Pivot PIVOT HASIL SO — WAJIB DISSENTUH setelah Sub-B:** seret
   `F_HasilSO[JENIS]` ke area **FILTERS** (atau Filters di panel PivotTable
   Fields) dan pilih **"SO"**. Ini persis dropdown "JENIS" di file sampel.
   ⚠️ Tanpa filter ini, Sum of Value menjumlahkan SEMUA jenis sekaligus —
   angkanya jadi salah. (Alternatif: taruh `JENIS` di Slicer agar bisa
   ganti jenis sekali klik — sama seperti kelakuan file sampel.)
9. *(Opsional)* Bersihkan measure generasi-1 yang tak terpakai:
   **Power Pivot → Measures → Manage** → hapus measure dari daftar bagian 3
   `DAX-Measures.dax`. Tidak wajib — hanya supaya field list rapi.
10. Tes double-click di CEK DEVIASI HARIAN → 18 kolom persis sampel; cek
    PIVOT HASIL SO (filter SO) vs `LAPORAN-VALIDASI.md`. Selesai.

> Kenapa langkah 6–7 perlu? Measure v1 membaca `JENIS` sebagai label sumber
> (SO / STOCK AWAL / …). Di skema v2, `JENIS` berisi label **metrik**
> (1.DEVIASI … 5.LOSS/SURPLUS) — sama seperti kolom `JENIS` di file sampel.
> Nilai metrik sudah tersimpan di kolom `Value`, jadi cukup "Sum of Value".
>
> Kenapa Sub-B perlu? Di file sampel, filter "JENIS" pivot PIVOT HASIL SO
> itu ada **di pivot-nya** (dropdown interaktif, default SO) — tabel
> sumbernya memuat semua 8 jenis. Versi v1–v4 F_HasilSO mengeraskan slice
> di query sehingga dropdown hilang. Sub-B mengembalikan perilaku sampel.

---

## PERILAKU BERBEDA (by design, supaya tidak kaget)

| # | Hal | Lama | Baru |
|---|-----|------|------|
| 1 | SO_KEMARIN hari-1 (9 item) | 0 (bug baris-null) | **Stok awal** (benar) — lihat laporan |
| 2 | Sel kosong pivot | 0 | blank → set: pivot **Options → Layout & Format → "For empty cells show: 0"** |
| 3 | Grand total baris AKUMULASI | penjumlahan kolom (tak bermakna) | akumulasi s/d hari 31 (benar); matikan grand total baris jika mau identik visual |
| 4 | Kolom teks Keterangan/Indikasi | kolom tabel (muncul saat double-click) | **SAMA seperti lama** — kini kolom tabel `F_Deviasi`, muncul saat double-click (paritas v2) |
| 5 | Dropdown "JENIS" di PIVOT HASIL SO | report field di pivot (default SO, bisa ganti jenis interaktif) | **SAMA (v5)** — `F_HasilSO[JENIS]` di FILTERS/Slicer; tabel memuat semua 8 JENIS (v1–v4: tabel SO-only, dropdown hilang) |
| 6 | Double-click CEK DEVIASI HARIAN | 18 kolom | **18 kolom identik** (paritas v2) |
| 7 | Double-click AKUMULASI | 18 kolom + kolom `AKUMULASI` | 18 kolom (kolom `Value` = nilai harian, bukan running total; angka running total tetap benar di sel pivot via measure) |
| 8 | Baris detail | termasuk 70 baris `(blank)` bug | tanpa baris null (bersih) |
| 9 | Metrik COM di PIVOT HASIL SO | tidak ada di file sampel (hanya di file produksi user) | **BARU (v6)**: nilai JENIS `6.PEMAKAIAN (COM)` — pilih di FILTERS/Slicer `F_HasilSO[JENIS]`; COM = SO kemarin + BM − BK − SO (hari-1 = STOK AWAL); identitas cek: COM = BOM − 1.DEVIASI |

## PANEN OPTIMASI

- ❌ Hilang: `List.Generate` per item (M paling berat — running total kini 1 measure)
- ❌ Hilang: query pivot hari `[FINAL] PIVOT HASIL SO` (kini tabel long `F_HasilSO`)
- ❌ Hilang: 3 cache pivot duplikat (2×15.880 + 1.199 baris, ± 630 KB per file)
- ❌ Hilang: refresh ulang penuh tiap ganti slice/filter
- ✅ Double-click CEK DEVIASI HARIAN = **18 kolom persis sampel** (v2)
- ✅ Pivot CEK DEVIASI HARIAN tanpa measure (Sum of Value, seperti sampel)
- ⚠️ Kembali ada: `Table.Group` custom per item di `F_Deviasi` (pembentuk
  SO_KEMARIN/Stock Seharusnya/5 metrik). Ini konsekuensi sadar agar 18 kolom
  melekat di baris untuk drillthrough — diukur tetap jauh lebih ringan daripada
  pipeline lama (1× hitung vs 2× cache + List.Generate + pivot M).

## MENGAPA PERHITUNGAN METRIK TETAP DI QUERY (M), BUKAN 100% DI DATA MODEL?

Pertanyaan: "filter JENIS untuk PIVOT HASIL SO, perhitungan Stock Seharusnya,
dan perhitungan deviasi — kenapa tidak dipindah 100% ke Data Model biar
lebih cepat?" Jawaban singkat: **sudah ada di garis pembagian optimal.**
Kuncinya bukan "M vs DAX", tapi KAPAN angka dihitung:

| Tempat | Kapan dihitung | Kecepatan interaksi pivot |
|---|---|---|
| Kolom hasil M (sekarang) | 1× saat Refresh | Storage engine murni — tercepat yang mungkin |
| DAX calculated column | 1× saat Refresh (formula engine, baris-per-baris) | Sama cepatnya — **tidak ada yang bertambah cepat** |
| DAX measure | TIAP klik/slice/filter | Hanya menang jika nilai bergantung konteks filter |

**Kenapa Stock Seharusnya + SO_KEMARIN + 5 metrik + Keterangan/Indikasi
tetap di M:**
1. Angkanya snapshot historis bulan berjalan — tidak berubah setelah
   refresh → precompute 1× lalu simpan statis = optimal. Measure berarti
   membayar hitung ulang berkali-kali untuk hasil yang identik.
2. SO_KEMARIN = chaining antar-baris (hari N memakai SO hari N-1 per item).
   Di M: `Table.Group` + List, single-pass, sudah di-buffer (1× evaluasi).
   Di DAX calculated column: context transition × 15.810 baris — tidak lebih
   cepat, cuma pindah tempat hitung, dan rumus jadi sulit di-debug
   (PQ punya preview per-step, DAX calc column tidak).
3. **Syarat keras drillthrough (permintaan migrasi v2):** double-click pivot
   CEK DEVIASI harus menampilkan 18 kolom KATEGORI..Indikasi persis sampel.
   Show Details HANYA menampilkan kolom fisik — measure tidak pernah muncul.
   Jadi semua metrik WAJIB kolom material, di M atau calc column; keduanya
   sama-sama dihitung saat refresh, maka dipilih yang mudah dibaca = M.
4. Keterangan/Indikasi = nested IF teks → di DAX jadi `SWITCH` panjang yang
   sulit dirawat.

**Filter JENIS untuk PIVOT HASIL SO — ada di PIVOT, bukan di M (diperbaiki v5):**
- Ground truth file sampel: tabel query lama "[FINAL] PIVOT HASIL SO" memuat
  SEMUA 8 JENIS (dibuktikan sharedItems cache3) dan pivotnya memfilter lewat
  REPORT FIELD "JENIS" (default SO, bisa diganti interaktif — sel A1/B1
  sheet sampel: `JENIS | SO`). v1–v4 sempat mengeraskan slice `JENIS="SO"`
  di query — angka default-nya benar, tapi dropdown-nya hilang. v5:
  `F_HasilSO` memuat semua 8 JENIS + kolom `JENIS`, slice pindah ke
  FILTERS pivot (default SO) = paritas penuh; filter field pivot tetap
  100% storage engine (tanpa formula engine), jadi tidak ada biaya performa.
- Yang TETAP wajib di M adalah DOMAIN tabelnya: join RFC_Kode (item berkode
  — domainnya memang beda dari F_Deviasi: ada GORENG MEKAR, tanpa KUAH
  BAKSO) + buang `Attribute=null`. Measure memang tak bisa mengubah domain
  BARIS pivot, tapi slice satu kolom (JENIS=SO) tidak butuh measure — cukup
  field filter di pivot.
- **COM (`6.PEMAKAIAN (COM)`) mengikuti garis yang sama (v6):** nilainya
  snapshot historis + rantai antar-baris (pola identik SO_KEMARIN) →
  materialisasi di M saat refresh; muncul sebagai nilai `JENIS` sehingga
  pemilihannya tetap storage engine murni. Cek silang otomatis ke angka
  asli cache lama lewat identitas COM = BOM − 1.DEVIASI: 3.091 cocok
  persis + 9 known-fix hari-1 (LAPORAN-VALIDASI Lampiran 6).

**Yang memang pindah ke model: hanya [Akumulasi]** — running total
bergantung konteks (kolom pivot = hari berjalan). Di M ia butuh
`List.Generate` sekuensial (terberat, sudah dihapus); di DAX cukup 1 measure
dengan cache storage engine. Inilah garis bagi yang benar:
**kontekstual → measure DAX; statis → materialisasi di M.**

**"100% ke model" mustahil secara struktur:** Data Model hanya menerima
tabel jadi. Gabungan 8 sumber, unpivot hari 1..31, konversi satuan ke gr,
join LINK pecahan→induk, label JENIS, filter KUAH BAKSO — semua operasi
bentuk/ETL yang tidak ada padanannya di DAX. Yang bisa "dipindahkan" hanya
lapisan metrik terakhir, dan di lapisan itu materialisasi M = interaksi
tercepat yang mungkin (Sum of kolom statis = nol formula engine).

Skala data (±16 ribu baris) membuat refresh didominasi overhead per-koneksi
Excel (44 query), bukan volume — memindahkan metrik ke DAX tidak mengubah
waktu refresh yang terasa. Kapan jawaban ini berubah? Jika data tumbuh ke
ratusan ribu/juta baris atau multi-outlet: saatnya keluar dari Excel
(database + semantic layer), bukan memindahkan beban ke DAX.

## JIKA MINTA ANGKA PERSIS LAMA (mode legacy)

Angka lama hari-1 untuk 9 item itu berasal dari bug (SO_KEMARIN=0 padahal stok awal ada
di baris yang sama — kontradiksi internal di output lama sendiri). Kalau karena alasan
audit butuh replikasi bug-for-bug, hubungi saya — varian legacynya bisa dibuat
(jangan filter `Attribute = null` di awal `F_Deviasi` + tambah 1 measure penyesuaian).
Default panduan ini = versi benar.
