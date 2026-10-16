section Section1;

shared #"IPT_Trial Bahan" = let
    Result =
        fnProcessLinkLookup(
            #"STG_Trial Bahan_Unpivot",
            SRC_AllSources[MappingSusut],
            "TRIAL",
            "TRIAL BAHAN"
        ),
    #"Added Custom" = Table.AddColumn(Result, "KETERANGAN", each "TRIAL")
in
    #"Added Custom";

shared #"IPT_Waste Bahan dan Packaging" = let
    BufferedUnpivot = Table.Buffer(STG_WasteBnP_Unpivot),

    //======================================================
    // 1. BASE
    //======================================================
    // 🔥 FIX: Table.ReplaceValue native (sebelumnya Table.TransformColumns dengan lambda)
    SafeKonversi = Table.ReplaceValue(
        BufferedUnpivot,
        null,
        0,
        Replacer.ReplaceValue,
        {"KONVERSI"}
    ),

    CalcBase = Table.AddColumn(
        SafeKonversi,
        "Value",
        each [Qty] * [KONVERSI],
        type number
    ),

    BaseValues = Table.SelectColumns(
        CalcBase,
        {
            "NAMA BARANG",
            "KETERANGAN",
            "PENYEBAB TERJADINYA WASTE",
            "Attribute",
            "Value"
        },
        MissingField.Ignore
    ),

    BaseFiltered = Table.SelectRows(
        BaseValues,
        each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
    ),

    //======================================================
    // 2. TOTAL QTY PER PENYEBAB (satu-satunya Group)
    //======================================================
    TotalPenyebabAll = Table.Group(
        BufferedUnpivot,
        {"PENYEBAB TERJADINYA WASTE", "Attribute"},
        {
            {"TotalQty", each List.Sum([Qty]), type number}
        }
    ),

    ValidTotalPenyebab = Table.SelectRows(
        TotalPenyebabAll,
        each [PENYEBAB TERJADINYA WASTE] <> null
            and [PENYEBAB TERJADINYA WASTE] <> ""
    ),

    //======================================================
    // 3. MAPPING
    //======================================================
    Mapping = Table.TransformColumnTypes(
        SRC_AllSources[MappingWasteBnP],
        {
            {"NAMA BARANG", type text},
            {"LINK BARANG", type text},
            {"KONVERSI", type number},
            {"KETERANGAN", type text},
            {"PENYEBAB TERJADINYA WASTE", type text}
        }
    ),

    ValidMapping = Table.Buffer(
        Table.SelectRows(
            Mapping,
            each [LINK BARANG] <> null
                and [LINK BARANG] <> ""
                and [LINK BARANG] <> "0"
        )
    ),

    //======================================================
    // 4. LINK CONTRIBUTION
    //======================================================
    JoinedMapping = Table.NestedJoin(
        ValidMapping,
        {"LINK BARANG"},
        ValidTotalPenyebab,
        {"PENYEBAB TERJADINYA WASTE"},
        "MapData",
        JoinKind.Inner
    ),

    ExpandedMapping = Table.ExpandTableColumn(
        JoinedMapping,
        "MapData",
        {"Attribute", "TotalQty"},
        {"Attribute", "TotalQty"}
    ),

    // 🔥 FIX: Table.ReplaceValue native dengan default 1 (untuk link path)
    //    Sebelumnya: Table.AddColumn dengan inline if [KONVERSI] = null then 1 else [KONVERSI]
    SafeMapKonversi = Table.ReplaceValue(
        ExpandedMapping,
        null,
        1,
        Replacer.ReplaceValue,
        {"KONVERSI"}
    ),

    CalcLink = Table.AddColumn(
        SafeMapKonversi,
        "Value",
        each [TotalQty] * [KONVERSI],
        type number
    ),

    LinkValues = Table.SelectColumns(
        CalcLink,
        {
            "NAMA BARANG",
            "KETERANGAN",
            "PENYEBAB TERJADINYA WASTE",
            "Attribute",
            "Value"
        },
        MissingField.Ignore
    ),

    LinkFiltered = Table.SelectRows(
        LinkValues,
        each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
    ),

    //======================================================
    // 5. COMBINE TANPA GROUP FINAL
    //======================================================
    CombinedValues = Table.Combine({BaseFiltered, LinkFiltered}),

    AddJenis = Table.AddColumn(
        CombinedValues,
        "JENIS",
        each "WASTE",
        type text
    ),

    AddSource = Table.AddColumn(
        AddJenis,
        "SOURCE DATA",
        each "WASTE BAHAN DAN PACKAGING",
        type text
    ),

    TypedAttr = Table.TransformColumnTypes(
        AddSource,
        {{"Attribute", Int64.Type}}
    )

in
    TypedAttr;

shared #"IPT_Input SO Harian" = let
    Result =
        fnProcessSOHarian(
            #"STG_Input SO Harian_Unpivot",
            SRC_AllSources[MappingSO],
            "SO",
            "INPUT SO HARIAN"
        ),
    #"Added Custom" = Table.AddColumn(Result, "KETERANGAN", each "SO")
in
    #"Added Custom";

shared #"IPT_Input Barang Masuk" = let
    Result =
        fnProcessSOHarian(
            #"STG_Input Barang Masuk_Unpivot",
            SRC_AllSources[MappingSO],
            "BARANG MASUK",
            "INPUT BARANG MASUK"
        ),
    #"Added Custom" = Table.AddColumn(Result, "KETERANGAN", each "BARANG MASUK")
in
    #"Added Custom";

shared #"IPT_Input Barang Keluar" = let
    Result =
        fnProcessSOHarian(
            #"STG_Input Barang Keluar_Unpivot",
            SRC_AllSources[MappingSO],
            "BARANG KELUAR",
            "INPUT BARANG KELUAR"
        ),
    #"Added Custom" = Table.AddColumn(Result, "KETERANGAN", each "BARANG KELUAR")
in
    #"Added Custom";

shared #"IPT_Stock Awal" = let

    Source =
        Excel.CurrentWorkbook(){[Name="IPT_STOCK_AWAL"]}[Content],

    // Tangkap Excel Error seperti #N/A, #VALUE!, #DIV/0!, dll
    SafeSource =
        Table.ReplaceErrorValues(
            Source,
            List.Transform(
                Table.ColumnNames(Source),
                each {_, null}
            )
        ),

    Header =
        fnHeaderStockAwal(),

    RenameHeader =
        Table.RenameColumns(
            SafeSource,
            List.Zip({
                List.FirstN(
                    Table.ColumnNames(SafeSource),
                    List.Count(Header)
                ),
                Header
            }),
            MissingField.Ignore
        ),

    #"Selected Columns" =
        Table.SelectColumns(
            RenameHeader,
            {"NAMA BARANG", "STOK AWAL"},
            MissingField.UseNull
        ),

    //======================================================
    // NAMA BARANG
    // Error / null → ""
    //======================================================
    #"Safe Nama Barang" =
        Table.TransformColumns(
            #"Selected Columns",
            {
                {
                    "NAMA BARANG",
                    each
                        try
                            if _ = null then
                                ""
                            else
                                Text.From(_)
                        otherwise
                            "",
                    type text
                }
            }
        ),

    //======================================================
    // STOK AWAL
    // HANYA NUMBER ASLI → DIPERTAHANKAN
    // SELAIN ITU → 0
    //======================================================
    #"Safe Stok Awal" =
        Table.TransformColumns(
            #"Safe Nama Barang",
            {
                {
                    "STOK AWAL",
                    each
                        if Value.Is(_, type number) then
                            _
                        else
                            0,
                    type number
                }
            }
        ),

    #"Renamed Value" =
        Table.RenameColumns(
            #"Safe Stok Awal",
            {{"STOK AWAL", "Value"}}
        ),

    #"Added Attribute" =
        Table.AddColumn(
            #"Renamed Value",
            "Attribute",
            each "1",
            type text
        ),

    #"Added Source Data" =
        Table.AddColumn(
            #"Added Attribute",
            "SOURCE DATA",
            each "STOCK AWAL",
            type text
        ),

    #"Added Jenis" =
        Table.AddColumn(
            #"Added Source Data",
            "JENIS",
            each "STOCK AWAL",
            type text
        ),

    #"Final Columns" =
        Table.SelectColumns(
            #"Added Jenis",
            {
                "NAMA BARANG",
                "Attribute",
                "Value",
                "SOURCE DATA",
                "JENIS"
            }
        ),
    #"Added Custom" = Table.AddColumn(#"Final Columns", "KETERANGAN", each "STOK AWAL")

in
    #"Added Custom";

shared #"RFC_Nama Item Stock Awal" = let
 Source = SRC_AllSources[NamaItem_StockAwal]
in
    Source;

shared #"RFC_Nama Item dan Konversi SO Harian" = let
 Source = SRC_AllSources[NamaItem_SOHarian]
in
    Source;

shared #"RFC_Nama Item dan Konversi Barang Masuk" = let
 Source = SRC_AllSources[NamaItem_BarangMasuk]
in
    Source;

shared #"RFC_Nama Item dan Konversi Barang Keluar" = let
 Source = SRC_AllSources[NamaItem_BarangKeluar]
in
 Source;

shared IPT_Susut = let
    Result =
        fnProcessLinkLookup(
            STG_Susut_Unpivot,
            SRC_AllSources[MappingSusut],
            "SUSUT",
            "SUSUT"
        ),

    Appended = Table.Combine({Result, SRC_AllSources[SusutCabe]}),
    #"Added Custom" = Table.AddColumn(Appended, "KETERANGAN", each "SUSUT")

in
    #"Added Custom";

shared #"RFC_Database Menu Promix 00" = let
 Source = SRC_AllSources[Menu_Promix00]
in
    Source;

shared #"RFC_Database Menu Promix 04" = let
    Source = SRC_AllSources[Menu_Promix04]
in
    Source;

shared #"RFC_Database Menu Waste Handling Dine IN" = let
    Source = SRC_AllSources[Menu_WasteHandling]
in
    Source;

shared #"RFC_Database Menu Waste Cancel Nota" = let
    Source = SRC_AllSources[Menu_WasteCN]
in
    Source;

shared #"RFC_Database Menu Trial Porsi" = let
    Source = SRC_AllSources[Menu_TrialPorsi]
in
    Source;

shared #"Raw Rekap Mentah" = let
    Source = Table.Combine({
        // === 1. STG_AllBreakdown (gabungan 5 breakdown: Promix00, Promix04,
        //        CancelNota, TrialPorsi, WasteHandling) ===
        STG_AllBreakdown,

        // === 2. IPT_Waste Bahan dan Packaging ===
        #"IPT_Waste Bahan dan Packaging",

        // === 3. IPT_Trial Bahan ===
        #"IPT_Trial Bahan",

        // === 4. IPT_Susut ===
        IPT_Susut,

        // === 5. IPT_Input SO Harian ===
        #"IPT_Input SO Harian",

        // === 6. IPT_Input Barang Masuk ===
        #"IPT_Input Barang Masuk",

        // === 7. IPT_Input Barang Keluar ===
        #"IPT_Input Barang Keluar",

        // === 8. IPT_Stock Awal ===
        #"IPT_Stock Awal"
    }),

    MainCols = {
        "NAMA BARANG", "Attribute", "Value", "SOURCE DATA", "JENIS",
        "KETERANGAN", "PENYEBAB TERJADINYA WASTE"
    },

    SelectedBaseColumns =
        Table.SelectColumns(Source, MainCols, MissingField.Ignore),

    ChangedType =
        Table.TransformColumnTypes(
            SelectedBaseColumns,
            {
                {"Attribute", Int64.Type},
                {"Value", type number},
                {"KETERANGAN", type text}
            }
        ),
    // 🔥 FIX P0-4: Pindahkan filter KUAH BAKSO SEBELUM join.
    //    Mengurangi baris yang perlu di-join dengan Kategori.
    FilteredKUAHBAKSO =
        Table.SelectRows(ChangedType, each [NAMA BARANG] <> "KUAH BAKSO"),

    // Ambil Kategori dari SRC_AllSources (sudah di-buffer)
    KategoriLookup =
        Table.SelectColumns(
            SRC_AllSources[Kategori],
            {"NAMA BARANG", "KATEGORI", "STATUS", "SATUAN"},
            MissingField.Ignore
        ),

    MergedQueries =
        Table.NestedJoin(
            FilteredKUAHBAKSO, {"NAMA BARANG"},
            KategoriLookup,    {"NAMA BARANG"},
            "Kategori_Lookup",
            JoinKind.LeftOuter
        ),

    ColsToExpand = {"KATEGORI", "STATUS", "SATUAN"},

    ExpandedKategori =
        Table.ExpandTableColumn(
            MergedQueries, "Kategori_Lookup",
            ColsToExpand, ColsToExpand
        ),

    Buffered = Table.Buffer(ExpandedKategori)
in
    Buffered;

shared #"Raw Rekap Mentah [FINAL]" = let
    Source = #"Raw Rekap Mentah",
    Reordered = Table.ReorderColumns(Source, {"KATEGORI", "SOURCE DATA", "JENIS", "STATUS", "NAMA BARANG", "KETERANGAN", "PENYEBAB TERJADINYA WASTE", "SATUAN", "Attribute", "Value"}),
    TypedAttr = Table.TransformColumnTypes(Reordered, {{"Attribute", type text}}, "en-ID"),
    BufferedForPivot = Table.Buffer(TypedAttr),
    DistinctDays = List.Transform({1..31}, Text.From),
    Pivoted = Table.Pivot(BufferedForPivot, DistinctDays, "Attribute", "Value", List.Sum),
    Replaced = Table.ReplaceValue(Pivoted, null, 0, Replacer.ReplaceValue, List.Transform({1..31}, Text.From))
in
    Replaced
;

shared F_Deviasi = let
    // =============================================================
    // TABEL FAKTA UNTUK DATA MODEL (pivot CEK DEVIASI & AKUMULASI).
    // Skema 18 kolom = REPLIKA PERSIS output query lama "RFC_Rekap
    // Mentah [CEK DEVIASI] Raw" (format hybrid: kolom lebar + baris
    // per metrik), tujuannya agar DOUBLE-CLICK / Show Details pada
    // pivot Data Model menampilkan baris & kolom yang sama persis
    // dengan file sampel:
    //   KATEGORI | NAMA BARANG | JENIS | SATUAN | Attribute |
    //   SO_KEMARIN | WASTE | TRIAL | SUSUT | STOCK AWAL |
    //   BARANG MASUK | BARANG KELUAR | BOM | Stock Seharusnya
    //   (SO_Kemarin + Barang Masuk - Barang Keluar - BOM) | SO |
    //   Value | Keterangan | Indikasi
    // Domain = JENIS<>REWORK, KATEGORI<>null, KATEGORI<>"MENU MIX"
    // (sama dengan query lama).
    // PERBAIKAN disengaja vs file lama: baris Attribute=null dibuang
    // SEBELUM dihitung (baris Value=0 hasil LeftOuter join BOM menu
    // trial di fnBreakdownUnified) -> SO_KEMARIN hari-1 = STOCK AWAL,
    // bukan 0 (bug lama pada 9 item terbetulkan - lihat
    // LAPORAN-VALIDASI.md).
    // Catatan desain: 5 metrik + Keterangan + Indikasi melekat
    // sebagai KOLOM baris (syarat agar muncul saat double-click pivot
    // Data Model / drillthrough). Konsekuensi: pivot CEK DEVIASI
    // HARIAN TIDAK lagi butuh measure (persis pivot sampel: JENIS di
    // Rows, Attribute di Columns, Sum of Value di Values), dan
    // running total AKUMULASI cukup 1 measure DAX [Akumulasi].
    // Query lama yang digantikan: "RFC_Rekap Mentah [CEK DEVIASI]
    // Raw" + "RFC_Rekap Mentah [CEK DEVIASI] Akumulasi" + Raw
    // Rekap Mentah [FINAL] PIVOT HASIL SO -> F_Deviasi / F_HasilSO.
    // Load: Only Create Connection + Add to Data Model.
    // =============================================================
    Source = #"Raw Rekap Mentah",
    #"Filtered Rows" = Table.SelectRows(
        Source,
        each [JENIS] <> "REWORK"
            and [KATEGORI] <> null
            and [KATEGORI] <> "MENU MIX"
            and [Attribute] <> null
    ),
    #"Removed Columns" = Table.RemoveColumns(
        #"Filtered Rows",
        {"SOURCE DATA", "KETERANGAN", "PENYEBAB TERJADINYA WASTE", "STATUS"}
    ),
    // 🔥 OPTIMASI SINGLE-PASS (fokus kecepatan): Table.Pivot memerlukan
    // daftar nilai JENIS (List.Distinct -> enumerasi #1) lalu memindai
    // ulang tabelnya (enumerasi #2). Tanpa buffer, M yang laziness-nya
    // mengevaluasi ulang seluruh rantai upstream "Raw Rekap Mentah"
    // (gabungan 8 sumber + join + konversi + group) dua kali per
    // refresh. Table.Buffer (~16 ribu baris, murah) menjamin 1x saja.
    #"Buffered Input" = Table.Buffer(#"Removed Columns"),
    #"Pivoted Column" = Table.Pivot(
        #"Buffered Input",
        List.Distinct(#"Buffered Input"[JENIS]),
        "JENIS",
        "Value",
        List.Sum
    ),

    GetVal = (r as record, fieldName as text) as number =>
        let
            v = Record.FieldOrDefault(r, fieldName, 0)
        in
            if v = null then 0 else v,

    #"Stok Seharusnya" =
        let
            Grouped = Table.Group(
                #"Pivoted Column",
                {"NAMA BARANG"},
                {
                    {
                        "Data",
                        (t) =>
                            let
                                SortedT = Table.Sort(
                                    t,
                                    each try Number.From([Attribute]) otherwise [Attribute]
                                ),
                                BufferedT = Table.Buffer(SortedT),
                                RowCount = Table.RowCount(BufferedT),

                                GetListSafe = (col as text) as list =>
                                    if Table.HasColumns(BufferedT, col)
                                    then List.Transform(Table.Column(BufferedT, col), each if _ = null then 0 else _)
                                    else List.Repeat({0}, RowCount),

                                SOList = GetListSafe("SO"),
                                AwalList = GetListSafe("STOCK AWAL"),
                                MasukList = GetListSafe("BARANG MASUK"),
                                KeluarList = GetListSafe("BARANG KELUAR"),
                                BOMList = GetListSafe("BOM"),

                                StockAwal1 = if List.IsEmpty(AwalList) then 0 else AwalList{0},
                                SOKemarinList = List.FirstN({StockAwal1} & SOList, RowCount),

                                StokList = List.Transform(
                                    List.Numbers(0, RowCount),
                                    each SOKemarinList{_} + MasukList{_} - KeluarList{_} - BOMList{_}
                                ),

                                Output = Table.FromColumns(
                                    Table.ToColumns(BufferedT) & {SOKemarinList, StokList},
                                    Table.ColumnNames(BufferedT) & {"SO_KEMARIN", "STOK_SEHARUSNYA"}
                                )
                            in
                                Output
                    }
                }
            ),

            Result = Table.Combine(Grouped[Data])
        in
            Result,

    #"Added 1.Deviasi" =
        Table.AddColumn(
            #"Stok Seharusnya",
            "1.DEVIASI",
            each GetVal(_, "SO") - [STOK_SEHARUSNYA],
            type number
        ),

    #"Added 2.Waste" =
        Table.AddColumn(
            #"Added 1.Deviasi",
            "2.WASTE",
            each - Number.Abs(GetVal(_, "WASTE")),
            type number
        ),
    #"Added 3.Susut" =
        Table.AddColumn(
            #"Added 2.Waste",
            "3.SUSUT",
            each - Number.Abs(GetVal(_, "SUSUT")),
            type number
        ),
    #"Added 4.Trial" =
        Table.AddColumn(
            #"Added 3.Susut",
            "4.TRIAL",
            each - Number.Abs(GetVal(_, "TRIAL")),
            type number
        ),
    #"Added 5.Loss/Surplus" =
        Table.AddColumn(
            #"Added 4.Trial",
            "5.LOSS/SURPLUS",
            each [1.DEVIASI] - [2.WASTE] - [3.SUSUT] - [4.TRIAL],
            type number
        ),
    #"Renamed Columns" =
        Table.RenameColumns(
            #"Added 5.Loss/Surplus",
            {
                {"STOK_SEHARUSNYA", "Stock Seharusnya (SO_Kemarin + Barang Masuk - Barang Keluar - BOM)"}
            },
            MissingField.Ignore
        ),
    #"Replaced Value1" = Table.ReplaceValue(#"Renamed Columns", null, 0, Replacer.ReplaceValue, {"5.LOSS/SURPLUS"}),

    MetricCols = List.Select(
        Table.ColumnNames(#"Replaced Value1"),
        each List.Contains({"1.DEVIASI", "2.WASTE", "3.SUSUT", "4.TRIAL", "5.LOSS/SURPLUS"}, _)
    ),

    #"Unpivoted Columns" = Table.Unpivot(#"Replaced Value1", MetricCols, "JENIS", "Value"),

    #"Added Keterangan" = Table.AddColumn(
        #"Unpivoted Columns",
        "Keterangan",
        each
            if List.Contains({"1.DEVIASI", "5.LOSS/SURPLUS"}, [JENIS]) then
                if [Value] < 0 then "SO lebih sedikit daripada perhitungan stock seharusnya"
                else if [Value] > 0 then "SO lebih banyak daripada perhitungan stock seharusnya"
                else "Stok Sesuai Perhitungan"
            else "-",
        type text
    ),

    #"Added Indikasi" = Table.AddColumn(
        #"Added Keterangan",
        "Indikasi",
        each
            if not List.Contains({"1.DEVIASI", "5.LOSS/SURPLUS"}, [JENIS]) then "-"
            else
                let
                    val = [Value],
                    soNow = GetVal(_, "SO"),
                    soPrev = GetVal(_, "SO_KEMARIN"),
                    masuk = GetVal(_, "BARANG MASUK"),
                    keluar = GetVal(_, "BARANG KELUAR"),
                    bom = GetVal(_, "BOM")
                in
                    if soNow > soPrev and masuk = 0 then
                        "Barang Masuk belum tercatat (Stok SO naik tapi Barang Masuk = 0)"
                    else if soNow = soPrev and bom > 0 and masuk = 0 then
                        "Stok SO tidak berkurang meski ada penjualan/BOM"
                    else if soNow < soPrev and keluar = 0 and bom = 0 then
                        "FIX: Barang Keluar / BOM belum tercatat (Stok SO turun tapi pemakaian di sistem = 0)"
                    else if val > 0 then
                        "1. Barang Masuk belum tercatat | 2. Under Takaran | 3. Salah catat Barang Keluar"
                    else if val < 0 then
                        "1. Barang Keluar belum tercatat (Kiriman/retur ke resto lain/DC) | 2. Pencatatan Promix 00 kurang tepat | 3. Loss karna Waste/Susut atau Over Takaran"
                    else
                        "-",
        type text
    ),

    // Urutan kolom PERSIS pivot cache lama (ground truth cacheDefinition1):
    // double-click pivot Data Model menampilkan kolom sesuai urutan ini.
    #"Reordered Columns" = Table.ReorderColumns(
        #"Added Indikasi",
        {
            "KATEGORI",
            "NAMA BARANG",
            "JENIS",
            "SATUAN",
            "Attribute",
            "SO_KEMARIN",
            "WASTE",
            "TRIAL",
            "SUSUT",
            "STOCK AWAL",
            "BARANG MASUK",
            "BARANG KELUAR",
            "BOM",
            "Stock Seharusnya (SO_Kemarin + Barang Masuk - Barang Keluar - BOM)",
            "SO",
            "Value",
            "Keterangan",
            "Indikasi"
        },
        MissingField.Ignore
    ),
    // Tipe kolom utk Data Model (Attribute = integer -> kolom pivot
    // hari 1..31 terurut numerik, sama seperti sampel). Kolom lebar
    // dari Table.Pivot hanya dikonversi jika kolomnya ada (tahan
    // terhadap bulan tanpa data salah satu JENIS).
    Typed = Table.TransformColumnTypes(
        #"Reordered Columns",
        List.Select(
            {
                {"KATEGORI", type text},
                {"NAMA BARANG", type text},
                {"JENIS", type text},
                {"SATUAN", type text},
                {"Attribute", Int64.Type},
                {"SO_KEMARIN", type number},
                {"WASTE", type number},
                {"TRIAL", type number},
                {"SUSUT", type number},
                {"STOCK AWAL", type number},
                {"BARANG MASUK", type number},
                {"BARANG KELUAR", type number},
                {"BOM", type number},
                {"Stock Seharusnya (SO_Kemarin + Barang Masuk - Barang Keluar - BOM)", type number},
                {"SO", type number},
                {"Value", type number},
                {"Keterangan", type text},
                {"Indikasi", type text}
            },
            each Table.HasColumns(#"Reordered Columns", {_{0}})
        )
    )
in
    Typed;

shared F_HasilSO = let
    // =============================================================
    // TABEL FAKTA UNTUK DATA MODEL (pivot PIVOT HASIL SO).
    // Pengganti query lama "Raw Rekap Mentah [FINAL] PIVOT HASIL SO"
    // tapi dari tabel LONG (tanpa pivot hari di M):
    //   - join KODE BARANG via RFC_Kode Barang (LeftOuter, filter non-null)
    //   - hanya JENIS = "SO" (sesuai tampilan pivot lama)
    //   - buang Attribute=null
    // Domain identik dengan pivot lama: item berkode (mencakup
    // GORENG MEKAR / kategori MENU MIX, tanpa KUAH BAKSO yang sudah
    // difilter di Raw Rekap Mentah).
    // Load: Only Create Connection + Add to Data Model.
    // =============================================================
    Source = #"Raw Rekap Mentah",
    Joined = Table.NestedJoin(
        Source, {"NAMA BARANG"},
        #"RFC_Kode Barang", {"NAMA BARANG"},
        "Kode", JoinKind.LeftOuter
    ),
    Expanded = Table.ExpandTableColumn(
        Joined, "Kode", {"KODE BARANG"}, {"KODE BARANG"}
    ),
    Filtered = Table.SelectRows(
        Expanded,
        each [KODE BARANG] <> null
            and [JENIS] = "SO"
            and [Attribute] <> null
    ),
    Selected = Table.SelectColumns(
        Filtered,
        {
            "KATEGORI",
            "KODE BARANG",
            "NAMA BARANG",
            "SATUAN",
            "Attribute",
            "Value"
        },
        MissingField.Ignore
    ),
    Typed = Table.TransformColumnTypes(
        Selected,
        {
            {"KATEGORI", type text},
            {"KODE BARANG", type text},
            {"NAMA BARANG", type text},
            {"SATUAN", type text},
            {"Attribute", Int64.Type},
            {"Value", type number}
        }
    )
in
    Typed;

shared Metric = let
    // =============================================================
    // TABEL DIMENSI PUTUS (disconnected) untuk baris metrik di
    // pivot CEK DEVIASI & AKUMULASI. JANGAN dibuat relasi ke tabel
    // lain di Power Pivot. Load: Only Create Connection + Add to
    // Data Model.
    // =============================================================
    Source = #table(
        type table [Metric = text],
        {
            {"1.DEVIASI"},
            {"2.WASTE"},
            {"3.SUSUT"},
            {"4.TRIAL"},
            {"5.LOSS/SURPLUS"}
        }
    )
in
    Source;

shared fnBreakdownUnified = (flagColumn as text, iptTable as table, sourceLabel as text, jenisLabel as text,
  optional dayShift as nullable logical) as table =>
let
    NeedShift = dayShift = true,

    BomSource = SRC_AllSources[BOM],


    Base =
        Table.SelectRows(
            BomSource,
            each Record.FieldOrDefault(_, flagColumn, false) = true
        ),

    FilteredKuahnya =
        Table.SelectRows(
            Base,
            each Record.FieldOrDefault(_, "Nama Barang Grup Barang", null) <> "KUAH BAKSO"
        ),

    Cleaned =
        Table.SelectColumns(
            FilteredKuahnya,
            {
                "Nama Barang Grup Barang",
                "NAMA BARANG",
                "Kuantitas"
            },
            MissingField.Ignore
        ),

    TypedCleaned =
        Table.TransformColumnTypes(
            Cleaned,
            {
                {"Nama Barang Grup Barang", type text},
                {"NAMA BARANG", type text},
                {"Kuantitas", type number}
            }
        ),
    UnpIPT =
        Table.UnpivotOtherColumns(
            iptTable,
            {"Menu"},
            "Attribute",
            "RawValue"
        ),
    ShiftedAttr =
        if NeedShift then
            Table.TransformColumns(
                UnpIPT,
                {
                    {
                        "Attribute",
                        each
                            let
                                n = Number.From(_)
                            in
                                if n = 1 then "31" else Text.From(n - 1),
                        type text
                    }
                }
            )
        else
            UnpIPT,

    Joined =
        Table.NestedJoin(
            TypedCleaned,
            {"Nama Barang Grup Barang"},
            ShiftedAttr,
            {"Menu"},
            "X",
            JoinKind.LeftOuter
        ),

    Expanded =
        Table.ExpandTableColumn(
            Joined,
            "X",
            {"Attribute", "RawValue"},
            {"Attribute", "RawValue"}
        ),

    AddedVal =
        Table.AddColumn(
            Expanded,
            "Value",
            each [Kuantitas] * (if [RawValue] = null then 0 else [RawValue]),
            type number
        ),


    WithSrc =
        Table.AddColumn(
            AddedVal,
            "SOURCE DATA",
            each sourceLabel,
            type text
        ),

    WithJns =
        Table.AddColumn(
            WithSrc,
            "JENIS",
            each jenisLabel,
            type text
        ),

    Final =
        Table.SelectColumns(
            WithJns,
            {
                "NAMA BARANG",
                "Attribute",
                "Value",
                "SOURCE DATA",
                "JENIS"
            },
            MissingField.Ignore
        )
in
    Final;

shared fnHeaderPromix = () as list =>
{
    "Menu",
    "1",
    "2",
    "3",
    "4",
    "5",
    "6",
    "7",
    "8",
    "9",
    "10",
    "11",
    "12",
    "13",
    "14",
    "15",
    "16",
    "17",
    "18",
    "19",
    "20",
    "21",
    "22",
    "23",
    "24",
    "25",
    "26",
    "27",
    "28",
    "29",
    "30",
    "31"
};

shared fnHeaderPromix04 = () as list =>
{
    "Menu",
    "2",
    "3",
    "4",
    "5",
    "6",
    "7",
    "8",
    "9",
    "10",
    "11",
    "12",
    "13",
    "14",
    "15",
    "16",
    "17",
    "18",
    "19",
    "20",
    "21",
    "22",
    "23",
    "24",
    "25",
    "26",
    "27",
    "28",
    "29",
    "30",
    "31",
    "1"
};

shared fnHeaderStockAwal = () as list =>
{
    "NAMA BARANG",
    "SATUAN",
    "STOK AWAL"

};

shared fnLoadMenuInput = (tableName as text, headerFn as function, optional dayOrderMode as text) as table =>
let
    Source = Excel.CurrentWorkbook(){[Name=tableName]}[Content],
    AllColumns = Table.ColumnNames(Source),
    TotalCols = List.Count(AllColumns),

    MetaCount   = 1,
    DayCount    = 31,

    Mode = if dayOrderMode = null then "standard" else dayOrderMode,

    MenuCol = if TotalCols > 0 then AllColumns{0} else null,
    AvailableDayIndices = if TotalCols > MetaCount then
        List.FirstN({MetaCount..TotalCols - 1}, DayCount)
    else
        {},
    DayColNames = List.Transform(AvailableDayIndices, each AllColumns{_}),

    Validation = 
        if MenuCol = null then
            error "Tabel " & tableName & " tidak memiliki kolom Menu (index 0)."
        else if List.IsEmpty(DayColNames) then
            error "Tabel " & tableName & " tidak memiliki kolom hari."
        else
            null,

    ColsToKeep = List.RemoveNulls({MenuCol}) & DayColNames,
    SelectedColumns = Table.SelectColumns(Source, ColsToKeep, MissingField.Ignore),

    TypeSpec = List.RemoveNulls({
        if MenuCol <> null then {MenuCol, type text} else null
    }) & List.Transform(DayColNames, each {_, type number}),

    ChangedType = Table.TransformColumnTypes(SelectedColumns, TypeSpec),


    ReplaceNull = Table.ReplaceValue(
        ChangedType, null, 0,
        Replacer.ReplaceValue, DayColNames
    ),

    RenameMenu = if MenuCol <> null and MenuCol <> "Menu" then
        {{MenuCol, "Menu"}}
    else
        {},

    DayOrder = if Mode = "rotated" then
        {"2","3","4","5","6","7","8","9","10","11","12","13","14","15","16",
         "17","18","19","20","21","22","23","24","25","26","27","28","29","30","31","1"}
    else
        List.Transform({1..31}, Text.From),

    RenameDays = List.Transform({0..List.Count(AvailableDayIndices) - 1}, each
        {AllColumns{AvailableDayIndices{_}}, DayOrder{_}}
    ),

    AllRenames = RenameMenu & RenameDays,
    Renamed = if List.IsEmpty(AllRenames) then ReplaceNull
              else Table.RenameColumns(ReplaceNull, AllRenames),

    Filtered = Table.SelectRows(Renamed, each [Menu] <> null and [Menu] <> "0")
in
    Filtered;

shared fnProcessLinkLookup = 
(unpivotTable as table, mappingTable as table, jenisLabel as text, sourceLabel as text) as table =>
let
    BufferedUnpivot = Table.Buffer(unpivotTable),


    SafeKonversiBase = Table.ReplaceValue(
        BufferedUnpivot,
        null,
        0,
        Replacer.ReplaceValue,
        {"KONVERSI"}
    ),
    CalcBase = Table.AddColumn(
        SafeKonversiBase,
        "Value",
        each [Qty] * [KONVERSI],
        type number
    ),
    BaseAll = Table.SelectColumns(CalcBase, {"NAMA BARANG", "Attribute", "Value"}),
    BaseFiltered = Table.SelectRows(
        BaseAll,
        each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
    ),


    TypedMapping = Table.TransformColumnTypes(
        Table.SelectColumns(
            mappingTable,
            {"NAMA BARANG", "LINK BARANG", "KONVERSI"},
            MissingField.Ignore
        ),
        {{"NAMA BARANG", type text}, {"LINK BARANG", type text}, {"KONVERSI", type number}}
    ),
    ValidMapping = Table.Buffer(
        Table.SelectRows(
            TypedMapping,
            each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
                and [LINK BARANG] <> null and [LINK BARANG] <> "" and [LINK BARANG] <> "0"
        )
    ),

    Joined = Table.NestedJoin(
        ValidMapping,
        {"LINK BARANG"},
        BufferedUnpivot,
        {"KETERANGAN"},
        "KetData",
        JoinKind.Inner
    ),
    Expanded = Table.ExpandTableColumn(
        Joined,
        "KetData",
        {"Attribute", "Qty"},
        {"Attribute", "Qty"}
    ),


    SafeMapKonversi = Table.ReplaceValue(
        Expanded,
        null,
        1,
        Replacer.ReplaceValue,
        {"KONVERSI"}
    ),
    CalcLink = Table.AddColumn(
        SafeMapKonversi,
        "Value",
        each [Qty] * [KONVERSI],
        type number
    ),
    LinkAll = Table.SelectColumns(CalcLink, {"NAMA BARANG", "Attribute", "Value"}),
    LinkFiltered = Table.SelectRows(
        LinkAll,
        each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
    ),


    Combined = Table.Combine({BaseFiltered, LinkFiltered}),
    GroupFinal = Table.Group(
        Combined,
        {"NAMA BARANG", "Attribute"},
        {{"Value", each List.Sum([Value]), type number}}
    ),

    AddJenis = Table.AddColumn(GroupFinal, "JENIS", each jenisLabel, type text),
    AddSource = Table.AddColumn(AddJenis, "SOURCE DATA", each sourceLabel, type text),
    TypedAttr = Table.TransformColumnTypes(AddSource, {{"Attribute", Int64.Type}})
in
    TypedAttr;

shared fnProcessSOHarian = 
(unpivotTable as table, mappingTable as table, jenisLabel as text, sourceLabel as text) as table =>
let
    BufferedUnpivot = Table.Buffer(unpivotTable),


    SafeKonversiBase = Table.ReplaceValue(
        BufferedUnpivot,
        null,
        0,
        Replacer.ReplaceValue,
        {"Konversi"}
    ),
    CalcBase = Table.AddColumn(
        SafeKonversiBase,
        "Value",
        each [Qty] * [Konversi],
        type number
    ),
    BaseAll = Table.SelectColumns(CalcBase, {"NAMA BARANG", "Attribute", "Value"}),
    BaseFiltered = Table.SelectRows(
        BaseAll,
        each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
    ),

   
    TypedMapping = Table.TransformColumnTypes(
        Table.SelectColumns(
            mappingTable,
            {"NAMA BARANG", "Konversi", "LINK"},
            MissingField.Ignore
        ),
        {{"NAMA BARANG", type text}, {"Konversi", type number}, {"LINK", type text}}
    ),
    ValidMapping = Table.Buffer(
        Table.SelectRows(
            TypedMapping,
            each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
                and [LINK] <> null and [LINK] <> "" and [LINK] <> "0"
        )
    ),


    Joined = Table.NestedJoin(
        ValidMapping,
        {"LINK"},
        BufferedUnpivot,
        {"ITEM STOCK OPNAME"},
        "MapData",
        JoinKind.Inner
    ),
    Expanded = Table.ExpandTableColumn(
        Joined,
        "MapData",
        {"Attribute", "Qty"},
        {"Attribute", "Qty"}
    ),

    SafeMapKonversi = Table.ReplaceValue(
        Expanded,
        null,
        1,
        Replacer.ReplaceValue,
        {"Konversi"}
    ),
    CalcLink = Table.AddColumn(
        SafeMapKonversi,
        "Value",
        each [Qty] * [Konversi],
        type number
    ),
    LinkAll = Table.SelectColumns(CalcLink, {"NAMA BARANG", "Attribute", "Value"}),
    LinkFiltered = Table.SelectRows(
        LinkAll,
        each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
    ),


    Combined = Table.Combine({BaseFiltered, LinkFiltered}),
    GroupFinal = Table.Group(
        Combined,
        {"NAMA BARANG", "Attribute"},
        {{"Value", each List.Sum([Value]), type number}}
    ),

    AddJenis = Table.AddColumn(GroupFinal, "JENIS", each jenisLabel, type text),
    AddSource = Table.AddColumn(AddJenis, "SOURCE DATA", each sourceLabel, type text),
    TypedAttr = Table.TransformColumnTypes(AddSource, {{"Attribute", Int64.Type}})
in
    TypedAttr;

shared fnUnpivotDailyStandard = 

(sourceTable as table, config as record) as table =>
let

    DayCount = Record.FieldOrDefault(config, "DayCount", 31),
    MetaCount = Record.FieldOrDefault(config, "MetaCount", 6),
    DayOrderMode = Record.FieldOrDefault(config, "DayOrderMode", "standard"),
    AttributeName = Record.FieldOrDefault(config, "AttributeName", "Attribute"),
    ValueName = Record.FieldOrDefault(config, "ValueName", "Qty"),
    AttributeType = Record.FieldOrDefault(config, "AttributeType", Int64.Type),
    EnsureAllDays = Record.FieldOrDefault(config, "EnsureAllDays", false),
    ReplaceNullsWithZero = Record.FieldOrDefault(config, "ReplaceNullsWithZero", true),
    FilterBlankKey = Record.FieldOrDefault(config, "FilterBlankKey", false),
    KeyColumn = Record.FieldOrDefault(config, "KeyColumn", null),
    MetaMap = Record.FieldOrDefault(config, "MetaMap", {}),


    AllColumns = Table.ColumnNames(sourceTable),
    TotalCols = List.Count(AllColumns),


    DayOrder =
        if DayOrderMode = "rotated" then
            List.Transform({2..DayCount}, Text.From) & {"1"}
        else
            List.Transform({1..DayCount}, Text.From),

    AvailableDayIndices =
        if TotalCols > MetaCount then
            List.FirstN({MetaCount..TotalCols - 1}, DayCount)
        else
            {},

    DayColNames = List.Transform(AvailableDayIndices, each AllColumns{_}),
    ExistingTargetDays = List.FirstN(DayOrder, List.Count(DayColNames)),


    RenameDays =
        if List.IsEmpty(DayColNames) then
            {}
        else
            List.Transform(
                List.Numbers(0, List.Count(DayColNames)),
                each {DayColNames{_}, ExistingTargetDays{_}}
            ),

    ResolveMeta =
        List.Transform(
            MetaMap,
            each
                let
                    m = _,
                    SourceName = Record.FieldOrDefault(m, "SourceName", null),
                    SourceIndex = Record.FieldOrDefault(m, "SourceIndex", null),
                    Target = Record.Field(m, "Target"),
                    Resolved =
                        if SourceName <> null and List.Contains(AllColumns, SourceName) then
                            SourceName
                        else if SourceIndex <> null and TotalCols > SourceIndex then
                            AllColumns{SourceIndex}
                        else
                            null
                in
                    Record.Combine({m, [ResolvedSource = Resolved]})
        ),

    MetaSourceCols =
        List.RemoveNulls(
            List.Transform(ResolveMeta, each Record.Field(_, "ResolvedSource"))
        ),

    ColsToKeep = List.Distinct(MetaSourceCols & DayColNames),

    Selected =
        Table.SelectColumns(
            sourceTable,
            ColsToKeep,
            MissingField.Ignore
        ),

MetaResolved =
    List.Select(
        ResolveMeta,
        each Record.Field(_, "ResolvedSource") <> null
    ),

NumericMetaSourceCols =
    List.Transform(
        List.Select(
            MetaResolved,
            each Record.FieldOrDefault(_, "Type", type any) = type number
        ),
        each Record.Field(_, "ResolvedSource")
    ),

NonNumericMetaTypeSpec =
    List.Transform(
        List.Select(
            MetaResolved,
            each Record.FieldOrDefault(_, "Type", type any) <> type number
        ),
        each {
            Record.Field(_, "ResolvedSource"),
            Record.FieldOrDefault(_, "Type", type any)
        }
    ),

TypedNonNumeric =
    Table.TransformColumnTypes(
        Selected,
        NonNumericMetaTypeSpec
    ),

SafeDays =
    Table.TransformColumns(
        TypedNonNumeric,
        List.Transform(DayColNames, each {_, fnSafeNumber, type number})
    ),

Typed =
    if List.IsEmpty(NumericMetaSourceCols) then
        SafeDays
    else
        Table.TransformColumns(
            SafeDays,
            List.Transform(NumericMetaSourceCols, each {_, fnSafeNumber, type number})
        ),

    MetaRenames =
        List.Transform(
            List.Select(
                ResolveMeta,
                each Record.Field(_, "ResolvedSource") <> null
                    and Record.Field(_, "ResolvedSource") <> Record.Field(_, "Target")
            ),
            each {Record.Field(_, "ResolvedSource"), Record.Field(_, "Target")}
        ),

    Renamed =
        Table.RenameColumns(
            Typed,
            RenameDays & MetaRenames,
            MissingField.Ignore
        ),

    EnsureMeta =
        List.Accumulate(
            ResolveMeta,
            Renamed,
            (state, m) =>
                let
                    Target = Record.Field(m, "Target"),
                    ColType = Record.FieldOrDefault(m, "Type", type any),
                    DefaultValue =
                        if ColType = type number then 0
                        else if ColType = type logical then false
                        else null,
                    Next =
                        if List.Contains(Table.ColumnNames(state), Target) then
                            state
                        else
                            Table.AddColumn(state, Target, each DefaultValue, ColType)
                in
                    Next
        ),

    TargetDaysToUse =
        if EnsureAllDays then DayOrder else ExistingTargetDays,

    EnsureDays =
        List.Accumulate(
            TargetDaysToUse,
            EnsureMeta,
            (state, d) =>
                if List.Contains(Table.ColumnNames(state), d) then
                    state
                else
                    Table.AddColumn(state, d, each 0, type number)
        ),


    NumericMetaTargets =
        List.Transform(
            List.Select(
                ResolveMeta,
                each Record.FieldOrDefault(_, "Type", type any) = type number
            ),
            each Record.Field(_, "Target")
        ),

    NumericColsToReplace =
        List.Select(
            NumericMetaTargets & TargetDaysToUse,
            each List.Contains(Table.ColumnNames(EnsureDays), _)
        ),

    ReplacedNulls =
        if ReplaceNullsWithZero and not List.IsEmpty(NumericColsToReplace) then
            Table.ReplaceValue(
                EnsureDays,
                null,
                0,
                Replacer.ReplaceValue,
                NumericColsToReplace
            )
        else
            EnsureDays,

    Filtered =
        if FilterBlankKey = true
            and KeyColumn <> null
            and List.Contains(Table.ColumnNames(ReplacedNulls), KeyColumn)
        then
            Table.SelectRows(
                ReplacedNulls,
                each Record.Field(_, KeyColumn) <> null
                    and Record.Field(_, KeyColumn) <> ""
            )
        else
            ReplacedNulls,


    MetaTypeTargets =
        List.Transform(
            ResolveMeta,
            each {Record.Field(_, "Target"), Record.FieldOrDefault(_, "Type", type any)}
        ),

    Unpivoted =
        if List.IsEmpty(TargetDaysToUse) then
            let
                EmptyCols =
                    List.Transform(ResolveMeta, each Record.Field(_, "Target"))
                    & {AttributeName, ValueName},

                EmptyTable =
                    Table.FromColumns(
                        List.Repeat({{}}, List.Count(EmptyCols)),
                        EmptyCols
                    ),

                TypedEmpty =
                    Table.TransformColumnTypes(
                        EmptyTable,
                        MetaTypeTargets
                        & {{AttributeName, type text}, {ValueName, type number}}
                    )
            in
                TypedEmpty
        else
            Table.Unpivot(
                Filtered,
                TargetDaysToUse,
                AttributeName,
                ValueName
            ),


    TypedValue =
        Table.TransformColumnTypes(
            Unpivoted,
            {{ValueName, type number}}
        ),

    ReplacedValueNull =
        if ReplaceNullsWithZero then
            Table.ReplaceValue(
                TypedValue,
                null,
                0,
                Replacer.ReplaceValue,
                {ValueName}
            )
        else
            TypedValue,

    TypedAttribute =
        Table.TransformColumnTypes(
            ReplacedValueNull,
            {{AttributeName, AttributeType}}
        )
in
    TypedAttribute;

shared RFC_Database_Susut = let
    Source = SRC_AllSources[RFC_Database_Susut]
in
    Source;

shared RFC_Database_Trial = let
    Source = SRC_AllSources[RFC_Database_Trial]
in
    Source;

shared RFC_Database_Waste_BnP = let
    Source = SRC_AllSources[WasteBnP],

    Filtered = Table.SelectRows(
        Source,
        each [Helper] <> "HIDE"
    ),

    Removed = Table.RemoveColumns(
        Filtered,
        {"Helper", "LINK BARANG"},
        MissingField.Ignore
    ),

    TambahIndex = Table.AddIndexColumn(
        Removed,
        "IndexUrutan",
        1,
        1,
        type number
    ),

    CariAlatMesin = Table.SelectRows(
        TambahIndex,
        each [KETERANGAN] <> null
            and Text.Contains([KETERANGAN], "Alat/ Mesin", Comparer.OrdinalIgnoreCase)
    ),

    NamaKolom = Table.ColumnNames(Removed),

    TemplateKosong = Record.FromList(
        List.Repeat({null}, List.Count(NamaKolom)),
        NamaKolom
    ),

    BuatBarisKosong = Table.TransformRows(
        CariAlatMesin,
        each Record.Combine({TemplateKosong, [IndexUrutan = [IndexUrutan] + 0.5]})
    ),

    TabelBarisKosong = Table.FromRecords(
        BuatBarisKosong,
        Table.ColumnNames(TambahIndex)
    ),

    GabungTabel = Table.Combine({TambahIndex, TabelBarisKosong}),

    UrutkanData = Table.Sort(
        GabungTabel,
        {{"IndexUrutan", Order.Ascending}}
    ),

    HasilAkhir = Table.RemoveColumns(
        UrutkanData,
        {"IndexUrutan"}
    )

in
    HasilAkhir;

shared #"STG_Trial Bahan_Unpivot" = let
    #"STG_Trial Bahan_Unpivot" = fnUnpivotTrialSusut("TRIAL_BAHAN")
in
    #"STG_Trial Bahan_Unpivot";

shared STG_Susut_Unpivot = let
    STG_Susut_Unpivot = fnUnpivotTrialSusut("SUSUT")
in
    STG_Susut_Unpivot;

shared #"STG_Input SO Harian_Unpivot" = let
    Source = Excel.CurrentWorkbook(){[Name="IPT_SO_HARIAN"]}[Content],
    Result = fnUnpivotDailyStandard(Source, CFG_InputHarian)
in
    Result;

shared #"STG_Input Barang Masuk_Unpivot" = let
    Source = Excel.CurrentWorkbook(){[Name="IPT_BARANG_MASUK"]}[Content],
    Result = fnUnpivotDailyStandard(Source, CFG_InputHarian)
in
    Result;

shared #"STG_Input Barang Keluar_Unpivot" = let
    Source = Excel.CurrentWorkbook(){[Name="IPT_BARANG_KELUAR"]}[Content],
    Result = fnUnpivotDailyStandard(Source, CFG_InputHarian)
in
    Result;

shared STG_WasteBnP_Unpivot = let
    Source = Excel.CurrentWorkbook(){[Name="IPT_WASTE_BAHAN_DAN_PACKAGING"]}[Content],

    AllColumns = Table.ColumnNames(Source),
    TotalCols = List.Count(AllColumns),

    MetaCount = 5,
    DayCount = 31,

    NamaBarangCol = if TotalCols > 0 then AllColumns{0} else null,
    KeteranganCol = if TotalCols > 1 then AllColumns{1} else null,
    PenyebabCol   = if TotalCols > 2 then AllColumns{2} else null,
    KonversiCol   = if TotalCols > 4 then AllColumns{4} else null,

    AvailableDayIndices =
        if TotalCols > MetaCount then
            List.FirstN({MetaCount..TotalCols - 1}, DayCount)
        else
            {},

    DayColNames = List.Transform(AvailableDayIndices, each AllColumns{_}),

    ColsToKeep =
        List.RemoveNulls({
            NamaBarangCol,
            KeteranganCol,
            PenyebabCol,
            KonversiCol
        }) & DayColNames,

    Selected =
        Table.SelectColumns(
            Source,
            ColsToKeep,
            MissingField.Ignore
        ),

    TypeSpec =
        List.RemoveNulls({
            if NamaBarangCol <> null then {NamaBarangCol, type text} else null,
            if KeteranganCol <> null then {KeteranganCol, type text} else null,
            if PenyebabCol <> null then {PenyebabCol, type text} else null,
            if KonversiCol <> null then {KonversiCol, type number} else null
        })
        &
        List.Transform(DayColNames, each {_, type number}),

    ChangedType =
        Table.TransformColumnTypes(Selected, TypeSpec),

    ReplaceNullDays =
        Table.ReplaceValue(
            ChangedType,
            null,
            0,
            Replacer.ReplaceValue,
            DayColNames
        ),

    RenameMeta =
        List.RemoveNulls({
            if NamaBarangCol <> null and NamaBarangCol <> "NAMA BARANG"
                then {NamaBarangCol, "NAMA BARANG"} else null,

            if KeteranganCol <> null and KeteranganCol <> "KETERANGAN"
                then {KeteranganCol, "KETERANGAN"} else null,

            if PenyebabCol <> null and PenyebabCol <> "PENYEBAB TERJADINYA WASTE"
                then {PenyebabCol, "PENYEBAB TERJADINYA WASTE"} else null,

            if KonversiCol <> null and KonversiCol <> "KONVERSI"
                then {KonversiCol, "KONVERSI"} else null
        }),

    RenameDays =
        List.Transform(
            AvailableDayIndices,
            each {AllColumns{_}, Text.From(_ - MetaCount + 1)}
        ),

    AllRenames = RenameMeta & RenameDays,

    Renamed =
        if List.IsEmpty(AllRenames)
        then ReplaceNullDays
        else Table.RenameColumns(ReplaceNullDays, AllRenames),

    EnsureNama =
        if List.Contains(Table.ColumnNames(Renamed), "NAMA BARANG")
        then Renamed
        else Table.AddColumn(Renamed, "NAMA BARANG", each null, type text),

    EnsureKet =
        if List.Contains(Table.ColumnNames(EnsureNama), "KETERANGAN")
        then EnsureNama
        else Table.AddColumn(EnsureNama, "KETERANGAN", each null, type text),

    EnsurePenyebab =
        if List.Contains(Table.ColumnNames(EnsureKet), "PENYEBAB TERJADINYA WASTE")
        then EnsureKet
        else Table.AddColumn(EnsureKet, "PENYEBAB TERJADINYA WASTE", each null, type text),

    EnsureKonversi =
        if List.Contains(Table.ColumnNames(EnsurePenyebab), "KONVERSI")
        then EnsurePenyebab
        else Table.AddColumn(EnsurePenyebab, "KONVERSI", each 0, type number),

    DayColStandard =
        List.Transform(
            AvailableDayIndices,
            each Text.From(_ - MetaCount + 1)
        ),

    FilterNama =
        Table.SelectRows(
            EnsureKonversi,
            each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
        ),

    Unpivot =
        if List.IsEmpty(DayColStandard)
        then
            Table.FromColumns(
                { {}, {}, {}, {}, {}, {} },
                {
                    "NAMA BARANG",
                    "KETERANGAN",
                    "PENYEBAB TERJADINYA WASTE",
                    "KONVERSI",
                    "Attribute",
                    "Qty"
                }
            )
        else
            Table.Unpivot(
                FilterNama,
                DayColStandard,
                "Attribute",
                "Qty"
            ),

    TypedAttr =
        Table.TransformColumnTypes(
            Unpivot,
            {{"Attribute", Int64.Type}}
        )
in
    TypedAttr;

shared STG_AllBreakdown = let
    Promix00 = fnBreakdownUnified(
        "BOM [Promix 00 & 04]",
        fnLoadMenuInput("IPT_PROMIX_00", fnHeaderPromix),
        "INPUT PENJUALAN DINE IN/TAKE AWAY 00", "BOM", false),

    Promix04 = fnBreakdownUnified(
        "BOM [Promix 00 & 04]",
        fnLoadMenuInput("IPT_PROMIX_04", fnHeaderPromix04, "rotated"),
        "INPUT PENJUALAN DINE IN/TAKE AWAY 04", "SO", true),

    CancelNota = fnBreakdownUnified(
        "Cancel Nota [Waste CN Dine In & TA]",
        fnLoadMenuInput("IPT_WASTE_CN", fnHeaderPromix),
        "WASTE PORSI CANCELNOTA DINE IN/TAKE AWAY", "WASTE", false),

    TrialPorsi = fnBreakdownUnified(
        "Trial Porsi [Trial Porsi]",
        fnLoadMenuInput("IPT_TRIAL_PORSI", fnHeaderPromix),
        "TRIAL PORSI", "TRIAL", false),

    WasteHandling = fnBreakdownUnified(
        "Waste Handling [Waste Handling]",
        fnLoadMenuInput("IPT_WASTE_HANDLING_DINE_IN", fnHeaderPromix),
        "WASTEPORSI DINEIN KARNAHANDLING", "WASTE", false),

    Combined = Table.Combine({Promix00, Promix04, CancelNota, TrialPorsi, WasteHandling}),
    #"Added Custom" = Table.AddColumn(Combined, "KETERANGAN", each if [SOURCE DATA] = "WASTE PORSI CANCELNOTA DINE IN/TAKE AWAY" then "Cancel Nota" else if [SOURCE DATA] = "WASTEPORSI DINEIN KARNAHANDLING" then "Handling" else if [SOURCE DATA] = "TRIAL PORSI" then "TRIAL" else if [SOURCE DATA] = "INPUT PENJUALAN DINE IN/TAKE AWAY 00" then "BOM" else if [SOURCE DATA] = "INPUT PENJUALAN DINE IN/TAKE AWAY 04" then "SO" else " "),
    #"Added Custom1" = Table.AddColumn(
    #"Added Custom",
    "PENYEBAB TERJADINYA WASTE",
    each if [SOURCE DATA] = "WASTE PORSI CANCELNOTA DINE IN/TAKE AWAY"
        then [NAMA BARANG] & " Karena Cancel Nota"
        else if [SOURCE DATA] = "WASTEPORSI DINEIN KARNAHANDLING"
        then [NAMA BARANG] & " Karena Handling"
        else " "
)
in
    #"Added Custom1";

shared SRC_AllSources = let
    //==============================================================
    // 0. FLAG BOM CONFIG + HELPER
    //==============================================================
    BOMFlagCols =
        {
            "BOM [Promix 00 & 04]",
            "Cancel Nota [Waste CN Dine In & TA]",
            "Trial Porsi [Trial Porsi]",
            "Waste Handling [Waste Handling]"
        },

    FlagToString = (v) as text =>
        if v = null then "" else Text.From(v),

    FlagToLogical = (v) as nullable logical =>
        let
            s = Text.Upper(Text.Trim(FlagToString(v)))
        in
            if s = "TRUE" then true
            else if s = "FALSE" then false
            else null,

    // 🔥 FIX: AnyTrueIgnoreNA — hapus List.Transform & List.IsEmpty redundan.
    AnyTrueIgnoreNA = (values as list) as logical =>
        let
            Converted = List.Transform(values, each FlagToLogical(_)),
            NonNull = List.Select(Converted, each _ <> null),
            Result = List.AnyTrue(NonNull)
        in
            Result,

    //==============================================================
    // 1. KONVERSI
    // 🔥 FIX: Hapus KonversiFiltered. Pakai KonversiRaw langsung.
    //    Filter KODE BARANG <> null HANYA di BOMTrueFalseSource & Kategori.
    //    NamaItem_SOHarian & NamaItem_BarangMasukKeluar TIDAK filter KODE BARANG.
    //==============================================================
    KonversiRaw =
        Excel.CurrentWorkbook(){[Name="Database_Konversi"]}[Content],

    KonversiTyped =
        Table.TransformColumnTypes(
            KonversiRaw,
            {
                {"NAMA BARANG", type text},
                {"KODE BARANG", type text},
                {"ITEM STOCK OPNAME", type text},
                {"SATUAN", type text},
                {"Konversi", type number},
                {"SATUAN 2", type text},
                {"INPUT SO HARIAN", type text},
                {"INPUT BARANG MASUK/KELUAR", type text},
                {"LINK", type text},
                {"Helper", type text},
                {"KATEGORI", type text},
                {"STATUS", type text},
                {"BOM [Promix 00 & 04]", type text},
                {"Cancel Nota [Waste CN Dine In & TA]", type text},
                {"Trial Porsi [Trial Porsi]", type text},
                {"Waste Handling [Waste Handling]", type text}
            }
        ),

    // 🔥 FIX P1-4a: Pre-convert 4 BOM flag text -> logical SEKALI di sini.
    //    N/A -> false (dianggap bukan flag aktif)
    KonversiWithFlags =
        Table.TransformColumns(
            KonversiTyped,
            List.Transform(
                BOMFlagCols,
                (c) => {
                    c,
                    (v) => let lg = FlagToLogical(v) in if lg = null then false else lg,
                    type logical
                }
            )
        ),

    Konversi =
        Table.Buffer(KonversiWithFlags),

    //==============================================================
    // 2. BOM TRUE-FALSE
    // 🔥 FIX: Tambah filter KODE BARANG valid di sini
    //    (sebelumnya dapat dari KonversiFiltered yang sekarang dihapus)
    //==============================================================
    BOMTrueFalseSource =
        Table.SelectRows(
            Konversi,
            each
                let
                    nama = Record.FieldOrDefault(_, "NAMA BARANG", null),
                    namaValid =
                        nama <> null
                        and Text.Trim(Text.From(nama)) <> "",

                    // 🔥 TAMBAHAN: KODE BARANG valid (hanya untuk BOM)
                    kode = Record.FieldOrDefault(_, "KODE BARANG", null),
                    kodeValid =
                        kode <> null
                        and Text.Trim(Text.From(kode)) <> "",

                    adaFlagTrue =
                        List.AnyTrue(
                            List.Transform(
                                BOMFlagCols,
                                (c) => Record.FieldOrDefault(_, c, false) = true
                            )
                        )
                in
                    namaValid and kodeValid and adaFlagTrue
        ),

    BOMTrueFalse =
        Table.Buffer(
            Table.Group(
                BOMTrueFalseSource,
                {"NAMA BARANG"},
                List.Transform(
                    BOMFlagCols,
                    (c) => {
                        c,
                        each List.AnyTrue(Table.Column(_, c)),
                        type logical
                    }
                )
            )
        ),

    //==============================================================
    // 3. BOM TRIAL
    //==============================================================
    BOMTrialRaw =
        Excel.CurrentWorkbook(){[Name="DBS_BARANG_TRIAL"]}[Content],

    BOMTrialBaseCols =
        {
            "Nama Barang Grup Barang",
            "NAMA BARANG",
            "Kuantitas"
        },

    BOMTrialSelected =
        Table.SelectColumns(
            BOMTrialRaw,
            BOMTrialBaseCols & BOMFlagCols,
            MissingField.Ignore
        ),

    BOMTrialTyped =
        Table.TransformColumnTypes(
            BOMTrialSelected,
            List.Select(
                {
                    {"Nama Barang Grup Barang", type text},
                    {"NAMA BARANG", type text},
                    {"Kuantitas", Int64.Type}
                }
                &
                List.Transform(BOMFlagCols, each {_, type text}),
                each List.Contains(Table.ColumnNames(BOMTrialSelected), _{0})
            )
        ),

    BOMTrialEnsureFlags =
        List.Accumulate(
            BOMFlagCols,
            BOMTrialTyped,
            (state, c) =>
                if List.Contains(Table.ColumnNames(state), c) then state
                else Table.AddColumn(state, c, each null, type text)
        ),

    BOMTrialLogical =
        Table.TransformColumns(
            BOMTrialEnsureFlags,
            List.Transform(
                BOMFlagCols,
                (c) =>
                    {
                        c,
                        (v) =>
                            let lg = FlagToLogical(v) in
                            if lg = null then false else lg,
                        type logical
                    }
            )
        ),

    BOMTrial =
        Table.Buffer(BOMTrialLogical),

    //==============================================================
    // 4. BOM MAIN
    //==============================================================
    BOMRaw =
        Table.SelectColumns(
            Excel.CurrentWorkbook(){[Name="DBS_BOM"]}[Content],
            {
                "Nama Barang Grup Barang",
                "NAMA BARANG",
                "Kuantitas",
                "Satuan2"
            },
            MissingField.Ignore
        ),

    BOMTyped =
        Table.TransformColumnTypes(
            BOMRaw,
            {
                {"Nama Barang Grup Barang", type text},
                {"NAMA BARANG", type text},
                {"Kuantitas", type number},
                {"Satuan2", type text}
            }
        ),

    BOMJoined =
        Table.NestedJoin(
            BOMTyped,
            {"NAMA BARANG"},
            BOMTrueFalse,
            {"NAMA BARANG"},
            "J",
            JoinKind.LeftOuter
        ),

    BOMExpanded =
        Table.ExpandTableColumn(
            BOMJoined,
            "J",
            BOMFlagCols,
            BOMFlagCols
        ),

    BOMExpandedFinal =
        Table.SelectColumns(
            BOMExpanded,
            {
                "Nama Barang Grup Barang",
                "NAMA BARANG",
                "Kuantitas"
            }
            & BOMFlagCols,
            MissingField.Ignore
        ),

    BOMTrialFinal =
        Table.SelectColumns(
            BOMTrial,
            {
                "Nama Barang Grup Barang",
                "NAMA BARANG",
                "Kuantitas"
            }
            & BOMFlagCols,
            MissingField.Ignore
        ),

    BOMCombined =
        Table.Combine({BOMExpandedFinal, BOMTrialFinal}),

    BOMFiltered =
        Table.SelectRows(
            BOMCombined,
            each [Nama Barang Grup Barang] <> null
                and [Nama Barang Grup Barang] <> "KUAH BAKSO"
        ),

    BOM =
        Table.Buffer(BOMFiltered),

    //==============================================================
    // 5. KATEGORI
    // ✅ Sudah punya filter KODE BARANG sendiri — tidak berubah
    //==============================================================
    Kategori = Table.Buffer(
        Table.SelectColumns(
            Table.SelectRows(
                Konversi,
                each Record.FieldOrDefault(_, "KODE BARANG", null) <> null
                    and Record.FieldOrDefault(_, "KODE BARANG", null) <> ""
            ),
            {"NAMA BARANG", "KATEGORI", "STATUS", "SATUAN"},
            MissingField.Ignore
        )
    ),

    //==============================================================
    // 6. MENU
    //==============================================================
    MenuRaw = Table.TransformColumnTypes(
        Excel.CurrentWorkbook(){[Name="DBS_MENU"]}[Content],
        {
            {"Menu", type text},
            {"Dine In/Take Away", type text},
            {"Grup Menu", type text},
            {"Aktif", type text}
        }
    ),

    Menu = Table.Buffer(
        Table.SelectRows(MenuRaw, each [Aktif] = "YA")
    ),

    //==============================================================
    // 7. TRIAL SUSUT RAW
    //==============================================================
    TrialSusutRaw = Table.Buffer(
        Table.TransformColumnTypes(
            Excel.CurrentWorkbook(){[Name="DBS_TRIAL_DAN_SUSUT"]}[Content],
            {
                {"Helper", type any},
                {"NAMA BARANG", type text},
                {"KETERANGAN", type any},
                {"SATUAN", type text},
                {"KONVERSI", type number},
                {"LINK BARANG", type any}
            }
        )
    ),

    //==============================================================
    // 8. WASTE BnP
    //==============================================================
    WasteBnP = Table.Buffer(
        Table.TransformColumnTypes(
            Excel.CurrentWorkbook(){[Name="SRC_WASTE_Bnp"]}[Content],
            {
                {"Helper", type text},
                {"NAMA BARANG", type text},
                {"KETERANGAN", type text},
                {"PENYEBAB TERJADINYA WASTE", type text},
                {"SATUAN", type text},
                {"LINK BARANG", type text},
                {"KONVERSI", type number}
            }
        )
    ),

    //==============================================================
    // 9. SUSUT CABE
    //==============================================================
    SusutCabeRaw = Excel.CurrentWorkbook(){[Name="Tabel_SUSUT_CABE"]}[Content],

    SusutCabeNoCol2 = Table.RemoveColumns(SusutCabeRaw, {"Column2"}, MissingField.Ignore),

    SusutCabeHdr = Table.PromoteHeaders(SusutCabeNoCol2, [PromoteAllScalars = true]),

    SusutCabeRen = Table.RenameColumns(
        SusutCabeHdr,
        {{"DATA MONITORING SUSUT CABE", "NAMA BARANG"}},
        MissingField.Ignore
    ),

    SusutCabeDayCols = List.Select(Table.ColumnNames(SusutCabeRen), each _ <> "NAMA BARANG"),

    SusutCabeTyped = Table.TransformColumnTypes(
        SusutCabeRen,
        {{"NAMA BARANG", type text}} & List.Transform(SusutCabeDayCols, each {_, Int64.Type})
    ),

    SusutCabeFilt = Table.SelectRows(
        SusutCabeTyped,
        each [NAMA BARANG] = "Susut harian aktual di resto"
    ),

    SusutCabeRen2 = Table.ReplaceValue(
        SusutCabeFilt,
        "Susut harian aktual di resto",
        "CABE RAWIT - RESTO (V.20)",
        Replacer.ReplaceValue,
        {"NAMA BARANG"}
    ),

    SusutCabeUnp = Table.UnpivotOtherColumns(
        SusutCabeRen2,
        {"NAMA BARANG"},
        "Attribute",
        "Value"
    ),

    SusutCabeSrc = Table.AddColumn(SusutCabeUnp, "SOURCE DATA", each "SUSUT", type text),

    SusutCabe = Table.Buffer(
        Table.AddColumn(SusutCabeSrc, "JENIS", each "SUSUT", type text)
    ),

    //==============================================================
    // 10. MAPPING
    //==============================================================
    MappingSO =
        Table.RemoveColumns(
            Konversi,
            {
                "KODE BARANG",
                "SATUAN",
                "SATUAN 2",
                "INPUT SO HARIAN",
                "INPUT BARANG MASUK/KELUAR",
                "Helper",
                "KATEGORI",
                "STATUS",
                "BOM [Promix 00 & 04]",
                "Cancel Nota [Waste CN Dine In & TA]",
                "Trial Porsi [Trial Porsi]",
                "Waste Handling [Waste Handling]"
            },
            MissingField.Ignore
        ),

    MappingSusut =
        Table.RemoveColumns(
            TrialSusutRaw,
            {"Helper", "KETERANGAN", "SATUAN"},
            MissingField.Ignore
        ),

    MappingWasteBnP =
        Table.RemoveColumns(
            Table.SelectRows(
                WasteBnP,
                each [NAMA BARANG] <> null and [NAMA BARANG] <> ""
            ),
            {"Helper", "SATUAN"},
            MissingField.Ignore
        ),

    //==============================================================
    // 11. RFC DATABASE SUSUT & TRIAL
    //==============================================================
    RFC_Database_Susut =
        Table.RemoveColumns(
            Table.SelectRows(TrialSusutRaw, each [Helper] = null),
            {"Helper", "LINK BARANG"},
            MissingField.Ignore
        ),

    RFC_Database_Trial = RFC_Database_Susut,

    //==============================================================
    // 12. RFC DATABASE MENU
    //==============================================================
    Menu_NonTrial = Table.Buffer(
        Table.SelectColumns(
            Table.SelectRows(Menu, each [Grup Menu] <> "07.TRIAL"),
            {"Menu"}
        )
    ),

    Menu_NonTrial_Full = Table.Buffer(
        Table.RemoveColumns(
            Table.SelectRows(Menu, each [Grup Menu] <> "07.TRIAL"),
            {"Dine In/Take Away", "Grup Menu", "Aktif", "No"},
            MissingField.Ignore
        )
    ),

    Menu_All_Full =
        Table.RemoveColumns(
            Menu,
            {"Dine In/Take Away", "Grup Menu", "Aktif", "No"},
            MissingField.Ignore
        ),

    Menu_Promix00 = Menu_NonTrial,
    Menu_Promix04 = Menu_NonTrial,
    Menu_WasteHandling = Menu_NonTrial_Full,
    Menu_WasteCN = Menu_NonTrial_Full,
    Menu_TrialPorsi = Menu_All_Full,

    //==============================================================
    // 13. RFC NAMA ITEM
    //==============================================================
    // ✅ NamaItem_StockAwal: filter KODE BARANG <> null and <> "0" (sudah benar)
    NamaItem_StockAwal =
        Table.SelectColumns(
            Table.SelectRows(
                Konversi,
                each Record.FieldOrDefault(_, "KODE BARANG", null) <> null
                    and Record.FieldOrDefault(_, "KODE BARANG", null) <> "0"
                    and Record.FieldOrDefault(_, "INPUT SO HARIAN", null) = "YA"
            ),
            {"NAMA BARANG", "SATUAN"},
            MissingField.Ignore
        ),

    // 🔥 NamaItem_SOHarian: TIDAK filter KODE BARANG (sesuai instruksi)
    NamaItem_SOHarian =
        Table.SelectColumns(
            Table.SelectRows(
                Konversi,
                each Record.FieldOrDefault(_, "INPUT SO HARIAN", null) = "YA"
                    and Record.FieldOrDefault(_, "Helper", null) <> "HIDE"
            ),
            {
                "NAMA BARANG",
                "KODE BARANG",
                "ITEM STOCK OPNAME",
                "SATUAN",
                "Konversi"
            },
            MissingField.Ignore
        ),

    // 🔥 NamaItem_BarangMasukKeluar: TIDAK filter KODE BARANG (sesuai instruksi)
    NamaItem_BarangMasukKeluar =
        Table.SelectColumns(
            Table.SelectRows(
                Konversi,
                each Record.FieldOrDefault(_, "INPUT BARANG MASUK/KELUAR", null) = "YA"
                    and Record.FieldOrDefault(_, "Helper", null) <> "HIDE"
            ),
            {
                "NAMA BARANG",
                "KODE BARANG",
                "ITEM STOCK OPNAME",
                "SATUAN",
                "Konversi"
            },
            MissingField.Ignore
        ),

    NamaItem_BarangMasuk = NamaItem_BarangMasukKeluar,
    NamaItem_BarangKeluar = NamaItem_BarangMasukKeluar

in
    [
        BOM = BOM,
        Konversi = Konversi,
        Kategori = Kategori,
        Menu = Menu,
        TrialSusutRaw = TrialSusutRaw,
        WasteBnP = WasteBnP,
        SusutCabe = SusutCabe,

        MappingSO = MappingSO,
        MappingSusut = MappingSusut,
        MappingWasteBnP = MappingWasteBnP,

        RFC_Database_Susut = RFC_Database_Susut,
        RFC_Database_Trial = RFC_Database_Trial,

        Menu_Promix00 = Menu_Promix00,
        Menu_Promix04 = Menu_Promix04,
        Menu_WasteHandling = Menu_WasteHandling,
        Menu_WasteCN = Menu_WasteCN,
        Menu_TrialPorsi = Menu_TrialPorsi,

        NamaItem_StockAwal = NamaItem_StockAwal,
        NamaItem_SOHarian = NamaItem_SOHarian,
        NamaItem_BarangMasuk = NamaItem_BarangMasuk,
        NamaItem_BarangKeluar = NamaItem_BarangKeluar
    ];

shared CFG_InputHarian = 
let
    Config = [
        MetaCount = 5,
        DayCount = 31,
        DayOrderMode = "standard",
        AttributeName = "Attribute",
        ValueName = "Qty",
        AttributeType = Int64.Type,
        EnsureAllDays = true,
        ReplaceNullsWithZero = true,
        FilterBlankKey = false,
        KeyColumn = "NAMA BARANG",
        MetaMap = {
            [SourceIndex = 0, Target = "NAMA BARANG", Type = type text],
            [SourceIndex = 2, Target = "ITEM STOCK OPNAME", Type = type text],
            [SourceIndex = 4, Target = "Konversi", Type = type number]
        }
    ]
in
    Config;

shared fnUnpivotTrialSusut = // fnUnpivotTrialSusut
(tableName as text) as table =>
let
    Source = Excel.CurrentWorkbook(){[Name=tableName]}[Content],
    DayCols = List.Transform({1..31}, Text.From),
    MetaCols = {"NAMA BARANG", "KETERANGAN", "KONVERSI"},
    ColsToKeep = MetaCols & DayCols,
    Selected = Table.SelectColumns(Source, ColsToKeep, MissingField.Ignore),
    ExistingMetaCols = List.Select(MetaCols, each List.Contains(Table.ColumnNames(Selected), _)),
    ExistingDayCols = List.Select(DayCols, each List.Contains(Table.ColumnNames(Selected), _)),
    TypeSpec = List.Transform(ExistingMetaCols, each if _ = "KONVERSI" then {_, type number} else {_, type text})
               & List.Transform(ExistingDayCols, each {_, type number}),
    Typed = Table.TransformColumnTypes(Selected, TypeSpec),
    AddMissingNama = if List.Contains(Table.ColumnNames(Typed), "NAMA BARANG") then Typed
                     else Table.AddColumn(Typed, "NAMA BARANG", each null, type text),
    AddMissingKet = if List.Contains(Table.ColumnNames(AddMissingNama), "KETERANGAN") then AddMissingNama
                    else Table.AddColumn(AddMissingNama, "KETERANGAN", each null, type text),
    AddMissingKonversi = if List.Contains(Table.ColumnNames(AddMissingKet), "KONVERSI") then AddMissingKet
                         else Table.AddColumn(AddMissingKet, "KONVERSI", each 0, type number),
    Replaced = Table.ReplaceValue(AddMissingKonversi, null, 0, Replacer.ReplaceValue, ExistingDayCols & {"KONVERSI"}),
    Unpivot = if List.IsEmpty(ExistingDayCols)
              then Table.FromColumns({{},{},{},{},{}}, {"NAMA BARANG","KETERANGAN","KONVERSI","Attribute","Qty"})
              else Table.Unpivot(Replaced, ExistingDayCols, "Attribute", "Qty"),
    TypedAttr = Table.TransformColumnTypes(Unpivot, {{"Attribute", Int64.Type}})
in
    TypedAttr;

shared fnSafeNumber =  (v as nullable any) as number =>
let
    IsNumber =
        v <> null
        and Value.Is(v, type number),

    TextValue =
        if IsNumber then
            ""
        else
            try Text.From(v) otherwise "",

    Trimmed =
        Text.Trim(TextValue),

    Result =
        if IsNumber then
            v
        else if Trimmed = "" then
            0
        else
            0
in
    Result;

shared #"RFC_Kode Barang" = let
    Source = #"RFC_Nama Item dan Konversi SO Harian",
    #"Filtered Rows" = Table.SelectRows(Source, each ([KODE BARANG] <> null)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"ITEM STOCK OPNAME", "Konversi"})
in
    #"Removed Columns";