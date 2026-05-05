// ============================================================
// DIM TABLES – paste each block as a separate query
// Replace "Sorted" on the Source line with whatever your
// cleaned fact query is actually named in the left panel
// ============================================================


// ── DIM BRAND ────────────────────────────────────────────────
let
    Source   = Sorted,
    KeepCol  = Table.SelectColumns(Source, {"Brand"}),
    Distinct = Table.Distinct(KeepCol),
    Output   = Table.Sort(Distinct, {{"Brand", Order.Ascending}})
in
    Output


// ── DIM CUSTOMER ─────────────────────────────────────────────
let
    Source   = Sorted,
    KeepCol  = Table.SelectColumns(Source, {"Customer"}),
    Distinct = Table.Distinct(KeepCol),
    Output   = Table.Sort(Distinct, {{"Customer", Order.Ascending}})
in
    Output


// ── DIM MONTH ────────────────────────────────────────────────
let
    Source   = Sorted,
    KeepCols = Table.SelectColumns(Source, {"Calendar Month", "Month Name"}),
    Distinct = Table.Distinct(KeepCols),
    Output   = Table.Sort(Distinct, {{"Calendar Month", Order.Ascending}})
in
    Output


// ── DIM TYPE ─────────────────────────────────────────────────
let
    Source   = Sorted,
    KeepCol  = Table.SelectColumns(Source, {"Type"}),
    Distinct = Table.Distinct(KeepCol),
    Output   = Table.Sort(Distinct, {{"Type", Order.Ascending}})
in
    Output


// ── DIM VARIATION ────────────────────────────────────────────
let
    Source   = Sorted,
    KeepCol  = Table.SelectColumns(Source, {"Variation"}),
    Distinct = Table.Distinct(KeepCol),
    Output   = Table.Sort(Distinct, {{"Variation", Order.Ascending}})
in
    Output


// ── DIM YEAR ─────────────────────────────────────────────────
let
    Source   = Sorted,
    KeepCol  = Table.SelectColumns(Source, {"Calendar Year"}),
    Distinct = Table.Distinct(KeepCol),
    Output   = Table.Sort(Distinct, {{"Calendar Year", Order.Ascending}})
in
    Output
