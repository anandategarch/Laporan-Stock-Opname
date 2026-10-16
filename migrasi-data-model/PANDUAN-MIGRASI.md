# PANDUAN MIGRASI — Perhitungan Berat ke Data Model
## Laporan Stock Opname (Full Power Query → Power Query + DAX)

> **📦 STATUS FILE (UPDATE v3 — PARITAS DOUBLE-CLICK + OPTIMASI SINGLE-PASS):
> file xlsx di repo ini SUDAH DIREVISI LANGSUNG.** Pembaruan v2: `F_Deviasi` kini
> ber-**18 kolom**
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
> Verifikasi struktur file pasca-revisi: **53/53 pemeriksaan otomatis LULUS**
> (ZIP/CRC, rels graph, content-types, metadata PQ, konsistensi koneksi↔query,
> Section1.m byte-exact).
>
> **⚡ Pembaruan v3 (fokus kecepatan, tanpa perubahan angka):** langkah
> `Table.Buffer(#"Removed Columns")` ditambahkan SEBELUM `Table.Pivot` di
> `F_Deviasi`. Alasan: `Table.Pivot` meng-enumerate tabel 2× (List.Distinct
> untuk daftar JENIS + scan data) — tanpa buffer, M dapat mengevaluasi ulang
> seluruh rantai upstream `Raw Rekap Mentah` (gabungan 8 sumber + join +
> konversi + group) **dua kali per refresh**. Buffer ± 16 ribu baris menjamin
> evaluasi 1× saja. Validasi ulang: hasil CEK1–4 identik byte-per-byte (15.765
> + 45 known-fix + 0 anomali) — angka tidak berubah sama sekali.
>
> **⚡ SUDAH TERMIGRASI PAKAI PANDUAN v1? Jangan download ulang — lihat bagian
> "UPGRADE FILE YANG SUDAH TERMIGRASI (± 10 MENIT)" di bawah.**
>
> **Status validasi angka: LULUS 100%** — 15.765/15.810 baris CEK DEVIASI,
> 15.252/15.810 AKUMULASI, 1.199/1.199 PIVOT HASIL SO, dan seluruh teks
> Keterangan/Indikasi cocok persis dengan cache pivot lama. Sisa selisih =
> **45 + 558 + 14 sel *known-fix*** = koreksi bug lama hari-1 (detail di
> `LAPORAN-VALIDASI.md`, bukan masalah baru).

---

## RINGKASAN PERUBAHAN QUERY (44 total)

| # | Query | Aksi |
|---|-------|------|
| 1 | `RFC_Rekap Mentah [CEK DEVIASI] Raw` | ❌ HAPUS (diganti `F_Deviasi` 18 kolom) |
| 2 | `RFC_Rekap Mentah [CEK DEVIASI] Akumulasi` | ❌ HAPUS (running total = 1 measure DAX) |
| 3 | `Raw Rekap Mentah [FINAL] PIVOT HASIL SO` | ❌ HAPUS (diganti `F_HasilSO`) |
| 4 | `F_Deviasi` | ➕ BARU — fakta 18 kolom (CEK DEVIASI & AKUMULASI, drillthrough = sampel) |
| 5 | `F_HasilSO` | ➕ BARU — fakta PIVOT HASIL SO |
| 6 | `Metric` | ➕ BARU — dimensi 5 metrik (kini **opsional** — `F_Deviasi[JENIS]` sudah memuat label metrik) |
| 7 | `Raw Rekap Mentah` | ⚪ TETAP connection-only (tidak ke model) |
| 8 | `RFC_Kode Barang` | ⚪ TETAP connection-only (dipakai `F_HasilSO`) |
| 9 | `Raw Rekap Mentah [FINAL]` | ⚪ TETAP load ke sheet REKAP MENTAH |
| 10 | Semua `SRC/STG/IPT/RFC_*/fn_*` lain | ⚪ TIDAK DISENTUH (byte-identical) |

**Tidak perlu relasi apa pun di Power Pivot.** Model = 2–3 tabel berdiri sendiri.

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

**PIVOT HASIL SO (2)** (baru):
- ROWS:    `F_HasilSO[KODE BARANG]` → `F_HasilSO[NAMA BARANG]` → `F_HasilSO[SATUAN]`
- COLUMNS: `F_HasilSO[Attribute]`
- VALUES:  `F_HasilSO[Value]` (Sum) — tidak perlu measure/filter apa pun

## LANGKAH 6 — Refresh & validasi angka (± 5 menit)

1. **Data → Refresh All** (2×; refresh pertama membangun model).
2. Bandingkan angka pivot baru vs laporan validasi (`LAPORAN-VALIDASI.md`), contoh kunci:
   - `SENDOK TEH SUNDA` 1.DEVIASI hari 3 = **817**
   - `LPG 12KG` SO hari 1–3 = sesuai REKAP MENTAH
   - PIVOT HASIL SO `GORENG MEKAR` hari 1 = 272, dst.
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

Pakai bagian ini bila file kerja Anda **sudah** termigrasi dengan panduan v1
(sudah load F_Deviasi ke Data Model, sudah pasang measure, sudah bangun pivot).
**Tidak perlu download ulang file repo** — cukup ganti isi query `F_Deviasi`:

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
4. **Pivot CEK DEVIASI HARIAN** — rapikan field list:
   - ROWS: hapus `Metric[Metric]`, ganti dengan `F_Deviasi[JENIS]`
     (urutan: KATEGORI → NAMA BARANG → JENIS → SATUAN).
   - VALUES: hapus `[Nilai Metric]`, ganti `F_Deviasi[Value]` (Sum of Value).
5. **Pivot AKUMULASI** — VALUES: hapus `[Akumulasi Metric]`, ganti measure
   `[Akumulasi]` (buat dulu bila belum ada — rumus di Langkah 4); ROWS idem langkah 4.
6. **Pivot PIVOT HASIL SO** — tidak perlu disentuh.
7. *(Opsional)* Bersihkan measure generasi-1 yang tak terpakai:
   **Power Pivot → Measures → Manage** → hapus measure dari daftar bagian 3
   `DAX-Measures.dax`. Tidak wajib — hanya supaya field list rapi.
8. Tes double-click di CEK DEVIASI HARIAN → 18 kolom persis sampel. Selesai.

> Kenapa langkah 4–5 perlu? Measure v1 membaca `JENIS` sebagai label sumber
> (SO / STOCK AWAL / …). Di skema v2, `JENIS` berisi label **metrik**
> (1.DEVIASI … 5.LOSS/SURPLUS) — sama seperti kolom `JENIS` di file sampel.
> Nilai metrik sudah tersimpan di kolom `Value`, jadi cukup "Sum of Value".

---

## PERILAKU BERBEDA (by design, supaya tidak kaget)

| # | Hal | Lama | Baru |
|---|-----|------|------|
| 1 | SO_KEMARIN hari-1 (9 item) | 0 (bug baris-null) | **Stok awal** (benar) — lihat laporan |
| 2 | Sel kosong pivot | 0 | blank → set: pivot **Options → Layout & Format → "For empty cells show: 0"** |
| 3 | Grand total baris AKUMULASI | penjumlahan kolom (tak bermakna) | akumulasi s/d hari 31 (benar); matikan grand total baris jika mau identik visual |
| 4 | Kolom teks Keterangan/Indikasi | kolom tabel (muncul saat double-click) | **SAMA seperti lama** — kini kolom tabel `F_Deviasi`, muncul saat double-click (paritas v2) |
| 5 | Double-click CEK DEVIASI HARIAN | 18 kolom | **18 kolom identik** (paritas v2) |
| 6 | Double-click AKUMULASI | 18 kolom + kolom `AKUMULASI` | 18 kolom (kolom `Value` = nilai harian, bukan running total; angka running total tetap benar di sel pivot via measure) |
| 7 | Baris detail | termasuk 70 baris `(blank)` bug | tanpa baris null (bersih) |

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

## JIKA MINTA ANGKA PERSIS LAMA (mode legacy)

Angka lama hari-1 untuk 9 item itu berasal dari bug (SO_KEMARIN=0 padahal stok awal ada
di baris yang sama — kontradiksi internal di output lama sendiri). Kalau karena alasan
audit butuh replikasi bug-for-bug, hubungi saya — varian legacynya bisa dibuat
(jangan filter `Attribute = null` di awal `F_Deviasi` + tambah 1 measure penyesuaian).
Default panduan ini = versi benar.
