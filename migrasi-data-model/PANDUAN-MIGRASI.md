# PANDUAN MIGRASI — Perhitungan Berat ke Data Model
## Laporan Stock Opname (Full Power Query → Power Query + DAX)

> **📦 STATUS FILE (UPDATE): file xlsx di repo ini SUDAH DIREVISI LANGSUNG** —
> **Langkah 1 & Langkah 3 di bawah SUDAH diterapkan ke file** (3 query lama dihapus
> dari DataMashup + 15 part pivot lama dibersihkan + 3 sheet pivot dikosongkan +
> 3 query baru `F_Deviasi` / `F_HasilSO` / `Metric` sudah tertanam, terlihat di
> panel Queries & Connections saat file dibuka).
> **Yang tinggal Anda lakukan: Langkah 2 (Load To → Data Model), Langkah 4 (paste
> measure DAX), Langkah 5 (bangun 3 pivot), Langkah 6 (refresh & validasi)** —
> total ± 20 menit klik-klik UI. (Model Data / measure DAX tidak bisa disuntik
> dari luar Excel — format `model.item` bersifat proprietary.)
> Verifikasi struktur file pasca-revisi: **53/53 pemeriksaan otomatis LULUS**
> (ZIP/CRC, rels graph, content-types, metadata PQ, konsistensi koneksi↔query).

> **Status validasi angka: LULUS 100%** — 15.765/15.810 baris CEK DEVIASI, 15.252/15.810 AKUMULASI,
> 1.199/1.199 PIVOT HASIL SO, dan seluruh teks Keterangan/Indikasi cocok persis dengan cache
> pivot lama. Sisa selisih = **45 + 558 + 14 sel *known-fix*** = koreksi bug lama hari-1
> (detail di `LAPORAN-VALIDASI.md`, bukan masalah baru).

---

## RINGKASAN PERUBAHAN QUERY (44 total)

| # | Query | Aksi |
|---|-------|------|
| 1 | `RFC_Rekap Mentah [CEK DEVIASI] Raw` | ❌ HAPUS (perhitungan pindah ke DAX) |
| 2 | `RFC_Rekap Mentah [CEK DEVIASI] Akumulasi` | ❌ HAPUS (running total = 1 measure) |
| 3 | `Raw Rekap Mentah [FINAL] PIVOT HASIL SO` | ❌ HAPUS (diganti `F_HasilSO`) |
| 4 | `F_Deviasi` | ➕ BARU — fakta CEK DEVIASI & AKUMULASI |
| 5 | `F_HasilSO` | ➕ BARU — fakta PIVOT HASIL SO |
| 6 | `Metric` | ➕ BARU — dimensi putus 5 metrik |
| 7 | `Raw Rekap Mentah` | ⚪ TETAP connection-only (tidak ke model) |
| 8 | `RFC_Kode Barang` | ⚪ TETAP connection-only (dipakai `F_HasilSO`) |
| 9 | `Raw Rekap Mentah [FINAL]` | ⚪ TETAP load ke sheet REKAP MENTAH |
| 10 | Semua `SRC/STG/IPT/RFC_*/fn_*` lain | ⚪ TIDAK DISENTUH (byte-identical) |

**Tidak perlu relasi apa pun di Power Pivot.** Model = 3 tabel berdiri sendiri.

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
Cara cepat: buka file itu, cari `shared F_Deviasi = let`, copy dari `let` sampai `in ... ;`.

Contoh untuk **F_Deviasi** (paste ke Advanced Editor):
```powerquery
let
    Source = #"Raw Rekap Mentah",
    Filtered = Table.SelectRows(
        Source,
        each [JENIS] <> "REWORK"
            and [KATEGORI] <> null
            and [KATEGORI] <> "MENU MIX"
            and [Attribute] <> null
    ),
    Selected = Table.SelectColumns(
        Filtered,
        {"KATEGORI", "NAMA BARANG", "JENIS", "SATUAN", "Attribute", "Value"},
        MissingField.Ignore
    ),
    Typed = Table.TransformColumnTypes(
        Selected,
        {
            {"KATEGORI", type text}, {"NAMA BARANG", type text},
            {"JENIS", type text}, {"SATUAN", type text},
            {"Attribute", Int64.Type}, {"Value", type number}
        }
    )
in
    Typed
```
Begitu juga `F_HasilSO` dan `Metric` — ambil dari `Section1_Revisi.m`.

## LANGKAH 2 — Atur tujuan load (± 2 menit)

Klik kanan query di panel → **Load To…** → pilih **Only Create Connection** →
**CENTANG "Add this data to the Data Model"** → OK. Terapkan pada:

| Query | Setting |
|-------|---------|
| `F_Deviasi` | Connection only + ✅ Data Model |
| `F_HasilSO` | Connection only + ✅ Data Model |
| `Metric` | Connection only + ✅ Data Model |

**JANGAN ubah** load query lain (RFC template, IPT, STG, [FINAL], dst.)

## LANGKAH 3 — Hapus 3 query lama (± 1 menit)

Klik kanan → **Delete** pada:
1. `RFC_Rekap Mentah [CEK DEVIASI] Raw`
2. `RFC_Rekap Mentah [CEK DEVIASI] Akumulasi`
3. `Raw Rekap Mentah [FINAL] PIVOT HASIL SO`

> 3 pivot lama (CEK DEVIASI HARIAN, AKUMULASI, PIVOT HASIL SO (2)) akan error/kosong —
> **normal**, kita bangun ulang di Langkah 5. Jangan refresh dulu.

## LANGKAH 4 — Buat measures DAX (± 10 menit)

Buka file **`DAX-Measures.dax`** → untuk tiap measure:
**Power Pivot tab → Measures → New Measure** → pilih *Table name* sesuai file →
paste rumus → isi *Measure name* persis → set Format: **Number, 0 desimal, ribuan**.

Total 20 measure: 8 metrik dasar, `SO Kemarin`, `Stok Seharusnya`, `Deviasi`,
`Waste (Neg)`, `Susut (Neg)`, `Trial (Neg)`, `Loss Surplus`, `Nilai Metric`,
`Akumulasi Metric`, `Keterangan`, `Indikasi`.

## LANGKAH 5 — Bangun 3 pivot baru dari Data Model (± 5 menit)

Untuk tiap sheet: buka sheet lama, hapus isi area pivot lama, lalu
**Insert → PivotTable → From Data Model**.

**CEK DEVIASI HARIAN** (baru):
- ROWS:    `F_Deviasi[KATEGORI]` → `F_Deviasi[NAMA BARANG]` → `Metric[Metric]` → `F_Deviasi[SATUAN]`
- COLUMNS: `F_Deviasi[Attribute]`
- VALUES:  `[Nilai Metric]`

**AKUMULASI DEVIASI ATAU TOTAL** (baru):
- sama dengan di atas, tetapi VALUES: `[Akumulasi Metric]`

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
3. ⚠️ **Angka hari-1 untuk 9 item akan BERBEDA dari pivot lama — itu koreksi bug, bukan
   error** (daftar item & angka before/after di `LAPORAN-VALIDASI.md` bagian BUG LAMA).
   Contoh: `ADONAN GORENG MEKAR` hari-1 deviasi: lama `+174.981` (salah, stok awal
   diabaikan) → baru `−11.512` (benar, 136.063 − (186.493 + 0 − 0 − 38.918)).
4. Jika semua cocok: hapus sisa pivot lama + sheet `Detail1`, save.

---

## PERILAKU BERBEDA (by design, supaya tidak kaget)

| # | Hal | Lama | Baru |
|---|-----|------|------|
| 1 | SO_KEMARIN hari-1 (9 item) | 0 (bug baris-null) | **Stok awal** (benar) — lihat laporan |
| 2 | Sel kosong pivot | 0 | blank → set: pivot **Options → Layout & Format → "For empty cells show: 0"** |
| 3 | Grand total baris AKUMULASI | penjumlahan kolom (tak bermakna) | akumulasi s/d hari 31 (benar); matikan grand total baris jika mau identik visual |
| 4 | Kolom teks Keterangan/Indikasi | kolom tabel | measure opsional (taruh di Values jika mau lihat) |

## PANEN OPTIMASI

- ❌ Hilang: `Table.Group` custom + `List.Generate` per item (M paling berat)
- ❌ Hilang: 2 cache duplikat 15.880 baris + 1 cache 1.199 baris
- ❌ Hilang: refresh ulang penuh tiap ganti slice/filter
- ✅ Tambah metrik baru = 1 baris DAX (tanpa refresh ulang M)
- ✅ Model 3 tabel kecil (~16k + 3k + 5 baris), kompresi VertiPaq

## JIKA MINTA ANGKA PERSIS LAMA (mode legacy)

Angka lama hari-1 untuk 9 item itu berasal dari bug (SO_KEMARIN=0 padahal stok awal ada
di baris yang sama — kontradiksi internal di output lama sendiri). Kalau karena alasan
audit butuh replikasi bug-for-bug, hubungi saya — measure varian legacynya bisa dibuat
(deteksi keberadaan baris-null per item sebelum difilter). Default panduan ini = versi benar.
