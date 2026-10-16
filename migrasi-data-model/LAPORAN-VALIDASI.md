# LAPORAN VALIDASI MIGRASI DATA MODEL
## Laporan Stock Opname — Sampel_Laporan SO_Full Power Query.xlsx

Tanggal: 16 Okt (sesi analisis) · Metode: cross-check otomatis (Python) antara
**logika DAX baru** dan **angka ground-truth di pivot cache lama** (pivotCacheRecords
1/2/3 = 15.880 + 15.880 + 1.199 baris hasil refresh terakhir workbook).

Skrip: `validasi_final.py` (dapat dijalankan ulang: `python3 validasi_final.py`,
path file di dalam skrip).

---

## HASIL

| Cek | Ground truth | Persis | Known-fix | Anomali | Status |
|-----|-------------|--------|-----------|---------|--------|
| CEK DEVIASI (cache1) | 15.880 baris | 15.765 | 45 | **0** | ✅ PASS |
| — teks Keterangan/Indikasi | 15.880 baris | 15.796 | 14 | **0** | ✅ PASS |
| AKUMULASI running total (cache2) | 15.880 baris | 15.252 | 558 | **0** | ✅ PASS |
| PIVOT HASIL SO (cache3) | 1.199 baris | 1.199 | 0 | **0** | ✅ PASS |
| M-logic vs DAX-logic | 3.162 sel | 0 selisih | — | **0** | ✅ PASS |

*"Known-fix" = sel yang berbeda **karena bug lama kini terbetulkan** (bukan error migrasi). Semua 617 sel terclassifikasi 100% — tidak ada satu pun selisih tak terjelaskan.*

### Validasi tambahan
- 102 item deviasi, 102 item kode (RFC_Kode), **0 duplikat** NAMA BARANG berkode → query aman di-join.
- `PANDUAN-MIGRASI`: 44 query (41 asli byte-identical + 3 baru), 0 referensi menggantung.
- Excel AppVersion 16.0300 → VAR/RETURN & SWITCH DAX aman (Excel 2016+).
- Distribusi hari: 88 item padat 31 hari + 14 item punya baris `Attribute=null` ekstra.

---

## BUG LAMA YANG DITEMUKAN (dan kini terbetulkan)

### Mekanisme
1. `fnBreakdownUnified` memakai `JoinKind.LeftOuter` dari sisi BOM. Resep **menu trial**
   (Grup Menu `07.TRIAL` di DBS_MENU, tidak muncul di 4 sheet menu non-trial) menghasilkan
   baris `Attribute = null`, `Value = 0` (JENIS: BOM, SO, WASTE) untuk 14 bahan resepnya.
2. Di `RFC_Rekap Mentah [CEK DEVIASI] Raw`, `Table.Sort(..., Number.From([Attribute]))`
   menempatkan **null paling depan** → `AwalList{0}` (yang seharusnya = stok awal hari 1)
   menjadi baris-null (STOCK AWAL = 0).
3. Akibat: `SO_KEMARIN` hari-1 = **0**, bukan stok awal → deviasi hari-1 melonjak palsu
   + diagnosa "Barang Masuk belum tercatat" yang menyesatkan → merambat ke running total
   akumulasi sebulan penuh.

Bukti internal: baris hari-1 `ADONAN GORENG MEKAR` di cache lama memuat
`STOCK AWAL = 186.493` **dan** `SO_KEMARIN = 0` sekaligus — kontradiksi di outputnya sendiri.

### 9 item terdampak (yang angkanya berubah, hari-1 + akumulasinya)
`ADONAN GORENG MEKAR` · `BERAS PUTIH` · `CUP 16 - BAKSO` · `CUP SUNDAE - BAKSO` ·
`KULIT PANGSIT GORENG MEKAR` · `RICE WRAP` · `SENDOK TEH SUNDAE (V.030)` ·
`SUNDAE LID - BAKSO` · `VANILLA SOFT MIX ICE MIX (V.20)`
*(5 item lagi punya baris null tapi stok awal & SO hari-1 = 0, angkanya kebetulan tak berubah.)*

### Before / After (contoh)

| Item | Hari-1 1.DEVIASI lama (bug) | Baru (benar) | Stok awal |
|------|---------------------------:|-------------:|----------:|
| ADONAN GORENG MEKAR | +174.981 | **−11.512** | 186.493 |
| BERAS PUTIH | +44.070 | **+9.270** | 34.800 |
| RICE WRAP | 0 | **−1.500** | 1.500 |

Perhitungan benar: `deviasi = SO − (stok awal + masuk − keluar − BOM)`
(contoh MEKAR: 136.063 − (186.493 + 0 − 0 − 38.918) = −11.512 ✓)

---

## CATATAN DESAIN (keputusan yang diambil supaya domain identik)

1. **F_Deviasi** memfilter `Attribute <> null` → pivot baru bersih dari kolom "(Blank)".
2. **F_HasilSO** dibuat dari tabel long (bukan [FINAL]) + join RFC_Kode + filter
   `JENIS="SO"` dan `KODE BARANG <> null` → domain identik pivot lama, **termasuk**
   `GORENG MEKAR` (kategori MENU MIX, berkode — 8 baris di cache lama) dan **tanpa**
   `KUAH BAKSO` (sudah difilter upstream).
3. `Raw Rekap Mentah` **tidak** di-load ke model (hemat; F_Deviasi & F_HasilSO yang naik).
4. Tidak ada relasi di model — 3 tabel berdiri sendiri (Metric = dimensi putus untuk baris metrik).
5. Measure `SO Kemarin` memakai pola "hari terakhir yang ADA sebelum hari ini" — padat
   data hari ini ≡ `hari−1`, tapi tetap benar jika kelak ada item dengan hari tidak
   kontinu (mis. SUSUT CABE hanya terisi sebagian).

---

## KESIMPULAN

**Migrasi aman dijalankan mengikuti `PANDUAN-MIGRASI.md`.**
Angka identik 100% dengan sistem lama, KECUALI koreksi bug hari-1 pada 9 item
yang terdokumentasi di atas. Tidak ada anomali tak terjelaskan.

---

## LAMPIRAN — CARA CEK BUG LAMA SECARA MANDIRI (di file sampel)

Semua angka di bawah adalah **isi asli workbook** (cache pivot hasil refresh terakhir),
bukan hasil sistem baru. Alamat sel merujuk file `Sampel_Laporan SO_Full Power Query.xlsx`.

### Lokasi & angka kunci

| # | Sheet | Lokasi | Tertulis (bug) | Seharusnya |
|---|-------|--------|----------------:|-----------:|
| 1 | CEK DEVIASI HARIAN | kolom paling kanan berjudul **"(blank)"** | kolom ekstra tak semestinya ada | — |
| 2 | CEK DEVIASI HARIAN | **E441** = ADONAN GORENG MEKAR / 1.DEVIASI / hari 1 | 174.981 | **−11.512** |
| 3 | CEK DEVIASI HARIAN | **E483** = KULIT PANGSIT GORENG MEKAR / hari 1 | 290.830 | **−16.895** |
| 4 | CEK DEVIASI HARIAN | **E273** = BERAS PUTIH / hari 1 | 44.070 | **9.270** |
| 5 | CEK DEVIASI HARIAN | baris 3 = SENDOK TEH SUNDAE (V.030) / hari 1 | 15.572 | **32** |
| 6 | AKUMULASI DEVIASI ATAU TOTAL | baris 441, kolom E→AI | 174.981 → 23.494,5 | −11.512 → **−162.998,5** |
| 7 | AKUMULASI DEVIASI ATAU TOTAL | baris 483, kolom E→AI | 290.830 → −169.515 | −16.895 → **−477.240** |
| 8 | STOK AWAL | **C2** | 186.493 | (input — benar) |
| 9 | REKAP MENTAH | baris ADONAN GORENG MEKAR (±I615) | STOCK AWAL h-1=186.493, BOM=38.918, SO=136.063 | (sumber — benar) |

### Hitungan cek (ADONAN GORENG MEKAR, hari 1)

```
1.DEVIASI = SO − (STOK AWAL + MASUK − KELUAR − BOM)
          = 136.063 − (186.493 + 0 − 0 − 38.918) = −11.512
```
Sel E441 tertulis **174.981** (= 136.063 + 38.918 — stok awal hilang dari hitungan) → bug.

**Bukti kontradiksi internal:** dobel-klik sel E441 (drill-down) → baris detail
menampilkan `STOCK AWAL = 186.493` **dan** `SO_KEMARIN = 0` pada baris yang sama.
(Contoh drill-down lama sudah ada di sheet `Detail1`.)

### Akar data (input kamu TIDAK salah — salah di query turunan)

5 menu terdaftar di DBS_BOM tapi tidak punya baris di 4 sheet input menu
(INPUT PENJUALAN 00 / INPUT PROMIX 04 / WASTE PORSI CANCELNOTA / WASTE DINE IN
KARNA HANDLING; GORENG MEKAR 1 PCS hanya ada di TRIAL PORSI) — resep bahan mereka
menghasilkan baris null via LeftOuter join:

`CARAMEL POPCORN SUNDAE` · `CUP - BAKSO` (tak terdaftar di DBS_MENU) ·
`GORENG MEKAR 1 PCS` (Grup 07.TRIAL) · `NASI` · `STRAWBERRY FRANZZI SUNDAE`
→ 13 bahan null; + `UDANG KEJU FROZEN` (jalur lookup trial, dampak 0).

Solusi tanpa migrasi: tambahkan baris menu tersebut ke sheet input menu terkait,
ATAU hapus/nonaktifkan resepnya di DBS_BOM → baris null hilang → bug ikut hilang.
(Dengan migrasi DAX, `F_Deviasi` memfilter `Attribute <> null` sehingga kolom
"(blank)" juga bersih.)

### Sheet yang TIDAK terdampak

- **REKAP MENTAH** — 39 kolom bersih; baris null terbuang oleh `Table.Pivot` di `[FINAL]`
  (bug tersembunyi, tidak terlihat di sheet ini).
- **PIVOT HASIL SO (2)** — validasi 1.199/1.199 baris identik.
- Semua sheet input & master — data sumber valid.
