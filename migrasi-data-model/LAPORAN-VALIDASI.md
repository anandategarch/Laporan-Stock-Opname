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

---

## LAMPIRAN 2 — VALIDASI v2: F_DEVIASI 18 KOLOM (PARITAS DOUBLE-CLICK)

**Permintaan:** double-click (Show Details / drillthrough) pivot **CEK DEVIASI
HARIAN** harus menampilkan kolom yang sama persis dengan file sampel, yaitu 18
kolom (urutan sesuai pivot cache lama): `KATEGORI · NAMA BARANG · JENIS ·
SATUAN · Attribute · SO_KEMARIN · WASTE · TRIAL · SUSUT · STOCK AWAL ·
BARANG MASUK · BARANG KELUAR · BOM · Stock Seharusnya (SO_Kemarin + Barang
Masuk - Barang Keluar - BOM) · SO · Value · Keterangan · Indikasi`.

**Perubahan v2 (commit ini):** `F_Deviasi` ditulis ulang menjadi replika
pipa hitung query lama `RFC_Rekap Mentah [CEK DEVIASI] Raw` (pivot JENIS →
`SO_KEMARIN` via Table.Group per NAMA BARANG → 5 metrik → unpivot metrik →
Keterangan/Indikasi) + dua penyesuaian: (1) filter `Attribute <> null`
dipindah ke **awal** pipa (perbaikan bug hari-1 tetap dipertahankan),
(2) langkah penutup `Table.ReorderColumns` eksplisit 18 kolom + pengetikan
kolom (Attribute = Int64 supaya kolom pivot hari terurut numerik 1..31,
sama seperti cache lama yang bertipe integer).

**Metode validasi** (`so-surgery/updates/validate_fdeviasi.py`, ground truth =
`pivotCacheDefinition1` + `pivotCacheRecords1` file sampel):

| Pemeriksaan | Hasil |
|---|---|
| Urutan & nama 18 kolom kode M == cache1 | **IDENTIK** (18/18) |
| Jumlah baris F_Deviasi baru | **15.810** (= cache 15.880 − 70 null) |
| `JENIS` (label metrik) | hanya `1.DEVIASI … 5.LOSS/SURPLUS` |
| `Attribute` | 1..31 semua integer, **0 null** |
| Nilai 18 kolom + teks vs cache1 | **15.765 persis + 45 known-fix + 0 anomali** |
| Baris ekstra / hilang | **0 / 0** |
| known-fix hari-1 | 9 item sama persis validasi v1 (ADONAN GORENG MEKAR, BERAS PUTIH, CUP 16 - BAKSO, CUP SUNDAE - BAKSO, KULIT PANGSIT GORENG MEKAR, RICE WRAP, SENDOK TEH SUNDAE (V.030), SUNDAE LID - BAKSO, VANILLA SOFT MIX ICE MIX (V.20)) |

Contoh before/after (sama dengan validasi v1):
`ADONAN GORENG MEKAR` hari-1 1.DEVIASI: lama **+174.981** → baru **−11.512**
(stok awal 186.493); `BERAS PUTIH`: +44.070 → **+9.270**; `RICE WRAP`: 0 → **−1.500**.

**Struktur workbook pasca-v2:** `build_new_xlsx.py` dijalankan ulang terhadap
`original.xlsx` — DataMashup berisi Section1.m baru (byte-exact), 44 koneksi ↔
44 query sinkron, metadata PQ konsisten, 15 part pivot tetap bersih, sheet lain
byte-identical. **53/53 pemeriksaan otomatis LULUS** (`so-surgery/validate.py`).

**Dampak desain yang perlu diketahui:**
1. `Table.Group` custom kembali ada di M (pembentuk SO_KEMARIN/Stock Seharusnya/
   5 metrik) — konsekuensi agar 18 kolom melekat di baris untuk drillthrough.
   `List.Generate` (running total AKUMULASI) TETAP hilang → 1 measure DAX
   `[Akumulasi]`; pivot cache duplikat TETAP hilang.
2. Measure generasi-1 (20 measure v1) obsolete — pivot CEK DEVIASI HARIAN kini
   cukup `Sum of Value` (persis pivot sampel). Lihat `DAX-Measures.dax` v2.
3. Drillthrough pivot AKUMULASI menampilkan kolom `Value` (nilai harian), bukan
   kolom `AKUMULASI` seperti sampel — angka running total di sel pivot tetap
   benar (dihitung measure). Selain itu 17 kolom lainnya identik.
4. Excel membatasi jumlah baris drillthrough pada pivot Data Model (grand total
   bisa terpotong); double-click sel spesifik (item × hari) selalu lengkap
   (1–5 baris per sel daun).

---

## LAMPIRAN 3 — Validasi Ulang v3 (Optimasi Single-Pass, fokus kecepatan)

Perubahan v3 SATU-SATUNYA: `#"Buffered Input" = Table.Buffer(#"Removed Columns")`
ditambahkan sebelum `Table.Pivot` di `F_Deviasi`. Latar: `Table.Pivot` harus
meng-enumerate tabel 2× (a) `List.Distinct(…[JENIS])` untuk daftar kolom pivot,
(b) scan data — dan karena M lazily-evaluated, tanpa buffer rantai upstream
`Raw Rekap Mentah` (gabungan 8 sumber + join + konversi + group) berpotensi
dievaluasi 2× per refresh. Buffer ± 15.810 baris menjamin evaluasi 1×.

Hasil validasi ulang (`so-surgery/updates/validate_fdeviasi.py`) — **identik
persis dengan hasil v2, tanpa satu angka pun berubah**:

| Pemeriksaan | Hasil |
|---|---|
| CEK1 urutan 18 kolom vs cache1 | PASS — identik |
| CEK2 domain | 15.810 baris, 0 null, Attribute 1..31 — PASS |
| CEK3 nilai vs cache1 | 15.765 persis + 45 known-fix + 0 anomali + 0 extra — PASS |
| CEK4 contoh before/after | ADONAN −11.512 / BERAS +9.270 / RICE WRAP −1.500 — PASS |
| Struktur workbook (validate.py) | 53/53 LULUS, Section1.m byte-exact kecuali blok F_Deviasi |

Sifat `Table.Buffer`: mematerialkan isi tabel di memori sekali per evaluasi —
urutan baris, urutan kolom, dan nilai TIDAK berubah; hanya pola enumerasi
yang berubah. Itulah sebabnya seluruh angka validasi tetap identik.

---

## LAMPIRAN 4 — v4: AUDIT KEBUTUHAN 44 QUERY + REGRUPING PANEL (6 FOLDER)

Permintaan: *"Query yang gak dibutuhkan langsung aja hapus dari file sampel, dan
query-nya dibuat grup biar mudah dibaca."*

### 4.1 Hasil audit — tidak ada lagi query yang aman dihapus

Graf dependensi seluruh 44 query di `Section1_Revisi.m` di-parse ulang (referensi
antar-query: bentuk `#"...multi kata..."` dan identifier telanjang), lalu
disandingkan dengan target pemuatan di metadata (`FillEnabled`, `FillTarget`,
`FillToDataModelEnabled`, `FillObjectType`):

| Kelas | Jumlah | Bukti dipakai |
|---|---|---|
| Pengisi template sheet input (`RFC_*`) | 13 | `FillEnabled=l1` + `FillTarget` (SUSUT, TRIAL, 5 sheet menu, 4 sheet nama item, WASTE BnP) — meregenerasi daftar item/menu di sheet input saat refresh; tanpa ini sheet input beku |
| Fungsi bantu (`fn*`) | 10 | dipanggil query lain: `fnSafeNumber`←`fnUnpivotDailyStandard`, `fnUnpivotTrialSusut`←`STG_Susut/Trial`, `fnBreakdownUnified`←`STG_AllBreakdown`+`F_Deviasi`, dst. |
| Intermediate pipeline (`STG_*`, `IPT_*`, `Raw Rekap Mentah`, `RFC_Kode Barang`, `SRC_AllSources`, `CFG_InputHarian`) | 24 | semuanya direferensikan minimal 1 query lain dalam M (0 orphan) |
| Output (load ke sheet/model) | 5 | `[FINAL]`→sheet REKAP MENTAH; `F_Deviasi`/`F_HasilSO`/`Metric`→Data Model |
| **Total** | **44** | **44/44 dipakai → 0 query dihapus di v4** |

Tiga query yang memang sudah tidak dibutuhkan — `RFC_Rekap Mentah [CEK DEVIASI]
Raw`, `[CEK DEVIASI] Akumulasi`, `[FINAL] PIVOT HASIL SO` — **sudah dihapus
sejak revisi v1** (bagian migrasi Data Model), jadi file saat ini memang sudah
bersih dari query mati. Menghapus 44 yang tersisa akan memutus rantai refresh
atau membekukan sheet input.

### 4.2 Temuan struktur panel lama (11 folder — sulit dibaca)

Blob biner `QueryGroups` di `customXml/item5.xml` (LocalPackageMetadataFile)
didekode — formatnya: `[int64 jumlah][GUID 16B][len 1B][nama]` per record
dipisah `[00 00][int64 NN]`. Isi file lama: folder `STG`(7), `SO`(4),
`TRIAL`(1), `WASTE`(1), `SUSUT`(1), **`INPUT` — KOSONG, 0 anggota**,
`DATABASE MENU`(5), `DATABASE SO`(5), `REKAP MENTAH`(5), `Function`(10),
`CONFIG`(2), **plus 3 query tanpa folder sama sekali**
(`RFC_Database_Susut/Trial/Waste_BnP`) — inilah yang membuat panel tampak
berantakan dan terkesan ada query "sampah".

### 4.3 Struktur baru — 6 folder bernomor mengikuti alur data

| # | Folder baru | GUID (reuse) | Member | Dari folder lama |
|---|---|---|---|---|
| 1 | `1. MASTER & CONFIG` | f749f097 (CONFIG) | 2 | CONFIG |
| 2 | `2. TEMPLATE INPUT` | cb7847cd (DATABASE SO) | 13 | DATABASE SO + DATABASE MENU + 3 root yatim |
| 3 | `3. STAGING (UNPIVOT)` | 55b371f9 (STG) | 7 | STG |
| 4 | `4. TRANSFORM (IPT)` | 7e851d46 (SO) | 7 | SO + SUSUT + TRIAL + WASTE |
| 5 | `5. OUTPUT` | 53168195 (REKAP MENTAH) | 5 | REKAP MENTAH |
| 6 | `6. FUNCTION` | 23687609 (Function) | 10 | Function |
| — | *dihapus* | — | — | INPUT (kosong), TRIAL, WASTE, SUSUT, DATABASE MENU |

2+13+7+7+5+10 = **44 — semua query ber-grup, 0 yatim, 0 folder kosong.**
GUID folder sengaja **di-reuse** sehingga 36 query yang grupnya tidak berubah
tidak perlu disentuh; hanya 8 query yang dipindah GID + 3 query root yang
disisipkan `<Entry Type="QueryGroupID">` pertamanya.

### 4.4 Validasi v4 (semua otomatis, `so-surgery/validate.py`)

| Pemeriksaan | Hasil |
|---|---|
| Blob QueryGroups: header=6, 6 record utuh 0 byte sisa, GUID+nama = desain, separator NN=0..5 | PASS |
| 44 query-level item metadata; semua punya QueryGroupID; 0 menunjuk grup lama (DATABASE MENU/SUSUT/TRIAL/WASTE/INPUT) | PASS |
| Member per folder 2/13/7/7/5/10; nama query metadata ≡ M ≡ koneksi (44/44/44) | PASS |
| 3 eks-root query masuk folder 2 | PASS |
| **Section1.m (M) byte-exact == v3** — tidak ada logika/angka yang berubah | PASS |
| Seluruh cek struktur v1–v3 (ZIP/CRC, XML, rels, content-types, koneksi, sheet) | PASS — **65/65 LULUS, 0 FAIL** |

Karena M tidak berubah, seluruh angka validasi Lampiran 1–3 berlaku identik
(15.765 + 45 known-fix + 0 anomali; AKUMULASI 15.252 + 558; PIVOT HASIL SO
1.199/1.199). Perubahan v4 murni kosmetik pada metadata panel — bukan pada
perhitungan.

---

## LAMPIRAN 5 — v5: PARITAS FILTER "JENIS" DI PIVOT HASIL SO

Pertanyaan user: *"di F_HasilSO tidak bisa ditambahkan JENIS yang sumbernya
dari Rekap Mentah?"* — jawabannya **bisa, dan ternyata memang seharusnya
begitu.** Penyelidikan ulang file sampel menemukan bahwa pemahaman v1–v4
tentang filter JENIS keliru arah.

### 5.1 Temuan ground truth (bukti dari file sampel asli)

| Bukti | Hasil |
|---|---|
| `pivotCacheDefinition3.xml` — sharedItems kolom `JENIS` | **8 nilai**: BARANG KELUAR, BARANG MASUK, SO, STOCK AWAL, BOM, SUSUT, TRIAL, WASTE → tabel query lama `"Raw Rekap Mentah [FINAL] PIVOT HASIL SO"` (cache3, 1.199 baris) memuat **SEMUA jenis**, TIDAK difilter ke SO |
| `pivotTable3.xml` (pivot PIVOT HASIL SO (2)) | `<pageField fld="2" .../>` → `JENIS` adalah **REPORT FILTER di pivot** (fld 2 = JENIS), bukan filter query |
| Render sheet PIVOT HASIL SO (2), sel A1/B1 | `JENIS \| SO` → filter tampil di pojok kiri pivot, **bisa diganti interaktif** (dropdown) |
| Query lama `"[FINAL] PIVOT HASIL SO"` (Section1.m asli) | hanya join `RFC_Kode Barang` + filter `KODE BARANG <> null` — **tidak ada filter JENIS** |

Artinya: di file sampel, kemampuan ganti-ganti jenis (SO → STOCK AWAL →
BARANG MASUK → …) itu **fitur interaktif pivot**. Migrasi v1–v4 mengeraskan
slice `JENIS="SO"` di query M dan membuang kolomnya — tampilan default
benar, tapi **dropdown itu hilang** (regresi interaktivitas yang tidak
disadari sampai user menanyakannya).

### 5.2 Perbaikan v5

`F_HasilSO` diubah (satu blok query, 43 lainnya tidak disentuh):
1. Filter `[JENIS] = "SO"` **dihapus** — semua 8 jenis ikut ter-load.
2. Kolom **`JENIS` ikut** (SelectColumns + TransformColumnTypes text).
3. Domain lain tetap: join `RFC_Kode` (item berkode) + `Attribute <> null`.
4. Posisi slice pindah ke pivot: `F_HasilSO[JENIS]` di **FILTERS** pivot,
   default **"SO"** (layout di PANDUAN-MIGRASI Langkah 5 / DAX-Measures.dax
   seksi 2). Filter field pivot = murni storage engine → **tanpa biaya
   performa**.

### 5.3 Validasi v5 (`so-surgery/updates/validate_fhasilso.py`, 15 cek LULUS)

| Pemeriksaan | Hasil |
|---|---|
| Subset SO v5 == tabel v4 (set baris + nilai persis) | **PASS — 0 mismatch** → angka SO tidak berubah |
| Distinct JENIS tabel v5 == sharedItems cache3 (8 jenis) | PASS |
| Bijeksi 5-tuple sheet↔cache (0 duplikat) → perbandingan baris-ke-baris 1:1 sah | PASS (1.199/1.199) |
| Meta baris cache3: KODE + KATEGORI + SATUAN | PASS 1.199/1.199 |
| **Nilai sel 1:1: 37.169 sel (1.199 baris × 31 hari) vs cache3** | **PASS — 100% persis, 0 selisih** |
| Pivot-view (sum per NAMA×JENIS×hari, cara pivot lama menjumlah) | PASS — 0 mismatch dari 20.274 sel berdata |
| Struktur M: filter `JENIS="SO"` hilang, kolom+tipe JENIS ada, filter KODE/Attribute tetap | PASS |
| Diff vs git: hanya blok `F_HasilSO` yang berubah, 43 query byte-identical | PASS |
| Struktur workbook (rebuild + `so-surgery/validate.py`) | **65/65 LULUS** |

Catatan cakupan: perbandingan nilai memakai rekonstruksi dari sheet
REKAP MENTAH (wide, 0-filled) — superset dari baris long sesungguhnya;
selisih hanya baris berisi 0 artifak fill yang tidak memengaruhi nilai
maupun tampilan pivot (perilaku blank-vs-0 sudah terdokumentasi di
PANDUAN "PERILAKU BERBEDA" #2, berlaku sama untuk semua jenis).

### 5.4 Konsekuensi & upgrade

- **File baru (download repo):** bangun pivot PIVOT HASIL SO dengan
  `F_HasilSO[JENIS]` di FILTERS = "SO" — perilaku = sampel penuh
  (dropdown bisa ganti jenis).
- **File yang sudah termigrasi v1–v4:** ikuti PANDUAN "UPGRADE" Sub-B
  (paste ulang `F_HasilSO`) + Sub-C langkah 8 (pasang filter JENIS di
  pivot). ⚠️ Tanpa filter, Sum of Value menjumlahkan semua jenis.
- Angka hasil (default SO) **identik** dengan semua validasi sebelumnya
  (1.199/1.199; subset SO 0 mismatch).
