let
    Source = Excel.Workbook(File.Contents("C:\Users\ishaa\Downloads\Drilldown\brand_dataset_raw_v2.xlsx"), null, true),
    #"Raw Data_Sheet" = Source{[Item="Raw Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Raw Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Brand", type text}, {"Type", type text}, {"Variation", type text}, {"Item", type text}, {"Calendar Year", Int64.Type}, {"Calendar Month", Int64.Type}, {"Month Name", type text}, {"Customer", type text}, {"Unit", Int64.Type}, {"CAD", type number}}),

    // ── STEP 1: Uppercase Brand & Item for consistent matching ──
    #"Upper Case" = Table.TransformColumns(#"Changed Type", {
        {"Brand", each Text.Upper(Text.Trim(_)), type text},
        {"Item",  each Text.Upper(Text.Trim(_)), type text}
    }),

    // ── STEP 2: Display flag ──
    // Any item containing a display keyword → Type = Display
    #"Display Flag" = Table.AddColumn(#"Upper Case", "_IsDisplay",
        each List.AnyTrue(List.Transform(
            {"DISP","PDQ","FLOORSTAND","PALLET","CLIPSTRIP","ENCAP","LAUNCH COUNTER","POWER"},
            (kw) => Text.Contains([Item], kw)
        )),
        type logical
    ),

    // ── STEP 3: Clean Brand ──
    // Brand F sub-brands (dirty names) → "Brand F"
    // Casing variants of Brand A–E → proper label
    #"Clean Brand" = Table.AddColumn(#"Display Flag", "Brand_Clean",
        each
            let b = [Brand] in
            if List.Contains({"GLOSS BLONDE","VELVET BRUNETTE","SILK BRUNETTE","SMOOTH FRIZZ","GLOW RED","PURE BLONDE","LIFT VOLUME"}, b)
            then "Brand F"
            else if Text.StartsWith(b, "BRAND A") then "Brand A"
            else if Text.StartsWith(b, "BRAND B") then "Brand B"
            else if Text.StartsWith(b, "BRAND C") then "Brand C"
            else if Text.StartsWith(b, "BRAND D") then "Brand D"
            else if Text.StartsWith(b, "BRAND E") then "Brand E"
            else if Text.StartsWith(b, "BRAND F") then "Brand F"
            else b,
        type text
    ),

    // ── STEP 4: Clean Type ──
    // Brand A (Ban)         → Roll On if ROLL in item, else Stick
    // Brand B (Biore)       → Core if serum/treatment/mixed product, else UV
    // Brand C (Bondi Sands) → Sun Protection if BODY/FACE/LIP, else Self Tan
    // Brand D (Curel)       → Global if CLEANSER, else Essentials
    // Brand E (JNG)         → Base always
    // Brand F (John Frieda) → Shampoo / Conditioner / Styler / Serum / Treatment / Detox & Repair / Hair Care
    // Display overrides all brands
    #"Clean Type" = Table.AddColumn(#"Clean Brand", "Type_Clean",
        each
            let b = [Brand_Clean], i = [Item] in
            if [_IsDisplay] then "Display"
            else if b = "Brand A" then
                if Text.Contains(i, "ROLL") then "Roll On" else "Stick"
            else if b = "Brand B" then
                if Text.Contains(i, "MIXED PRODUCT") or Text.Contains(i, "SERUM") or Text.Contains(i, "TREATMENT")
                then "Core" else "UV"
            else if b = "Brand C" then
                if Text.Contains(i, "BODY") or Text.Contains(i, "FACE") or Text.Contains(i, "LIP")
                then "Sun Protection" else "Self Tan"
            else if b = "Brand D" then
                if Text.Contains(i, "CLEANSER") then "Global" else "Essentials"
            else if b = "Brand E" then "Base"
            else if b = "Brand F" then
                if Text.Contains(i, "DETOX") then "Detox & Repair"
                else if Text.Contains(i, "SHAMPOO") or Text.Contains(i, "SHAMP") or Text.StartsWith(i, "SH ") then "Shampoo"
                else if Text.Contains(i, "CONDITIONER") or Text.Contains(i, "COND") or Text.StartsWith(i, "CD ") then "Conditioner"
                else if Text.Contains(i, "STYLER") then "Styler"
                else if Text.Contains(i, "SERUM") then "Serum"
                else if Text.Contains(i, "TREATMENT") or Text.Contains(i, "TRTMENT") or Text.Contains(i, " TRT") then "Treatment"
                else "Hair Care"
            else "Unknown",
        type text
    ),

    // ── STEP 5: Clean Variation ──
    // Global (all brands, checked first):
    //   GIFT PACK      → Gift Pack
    //   MIXED PRODUCT  → On Pack
    // Brand A: UNSCENTED → Unscented, else Scented
    // Brand B: AQUA RICH → Aqua Rich | PATCH → Patches | SCRUB → Cleanser
    //          MASK → Mask | CLEANSER → Cleanser | STRIP → Pore Strips
    // Brand C: FOAM → Foam | DROP → Drops | ERASER → Eraser | MITT → Mitt | LOTION → Lotion
    // Brand D: FOOT → Foot | MOISTURI → Moisturizers | FACE CREAM → Base Cream | CLEANSER → Cleansers | else Base
    // Brand E: MITT → Mitt | TOWELETTE → Towelette | INSTANT SUN → Instant | HYDRA GEL → Hydra Gel | else Gradual
    // Brand F: DETOX → Detox and Repair | DEEP SEA → Deep Sea | SHAMPOO/SHAMP/SH → Shampoo
    //          COND/CD → Conditioner | STYLER → Styler | SERUM → Serum | TREATMENT/TRTMENT/TRT → Treatment
    #"Clean Variation" = Table.AddColumn(#"Clean Type", "Variation_Clean",
        each
            let b = [Brand_Clean], i = [Item] in
            if Text.Contains(i, "GIFT PACK")      then "Gift Pack"
            else if Text.Contains(i, "MIXED PRODUCT") then "On Pack"
            else if b = "Brand A" then
                if Text.Contains(i, "UNSCENTED") then "Unscented" else "Scented"
            else if b = "Brand B" then
                if Text.Contains(i, "AQUA RICH")   then "Aqua Rich"
                else if Text.Contains(i, "PATCH")   then "Patches"
                else if Text.Contains(i, "SCRUB")   then "Cleanser"
                else if Text.Contains(i, "MASK")    then "Mask"
                else if Text.Contains(i, "CLEANSER") then "Cleanser"
                else if Text.Contains(i, "STRIP")   then "Pore Strips"
                else "General"
            else if b = "Brand C" then
                if Text.Contains(i, "FOAM")   then "Foam"
                else if Text.Contains(i, "DROP")   then "Drops"
                else if Text.Contains(i, "ERASER") then "Eraser"
                else if Text.Contains(i, "MITT")   then "Mitt"
                else if Text.Contains(i, "LOTION") then "Lotion"
                else "General"
            else if b = "Brand D" then
                if Text.Contains(i, "FOOT")          then "Foot"
                else if Text.Contains(i, "MOISTURI")  then "Moisturizers"
                else if Text.Contains(i, "FACE CREAM") then "Base Cream"
                else if Text.Contains(i, "CLEANSER")   then "Cleansers"
                else "Base"
            else if b = "Brand E" then
                if Text.Contains(i, "MITT")          then "Mitt"
                else if Text.Contains(i, "TOWELETTE") then "Towelette"
                else if Text.Contains(i, "INSTANT SUN") then "Instant"
                else if Text.Contains(i, "HYDRA GEL")   then "Hydra Gel"
                else "Gradual"
            else if b = "Brand F" then
                if Text.Contains(i, "DETOX")      then "Detox and Repair"
                else if Text.Contains(i, "DEEP SEA") then "Deep Sea"
                else if Text.Contains(i, "SHAMPOO") or Text.Contains(i, "SHAMP") or Text.StartsWith(i, "SH ") then "Shampoo"
                else if Text.Contains(i, "CONDITIONER") or Text.Contains(i, "COND") or Text.StartsWith(i, "CD ") then "Conditioner"
                else if Text.Contains(i, "STYLER")    then "Styler"
                else if Text.Contains(i, "SERUM")     then "Serum"
                else if Text.Contains(i, "TREATMENT") or Text.Contains(i, "TRTMENT") or Text.Contains(i, " TRT") then "Treatment"
                else "General"
            else "General",
        type text
    ),

    // ── STEP 6: Title-case Item ──
    #"Title Case Item" = Table.TransformColumns(#"Clean Variation", {
        {"Item", each Text.Proper(_), type text}
    }),

    // ── STEP 7: Remove helper columns, replace Brand/Type/Variation ──
    #"Remove Helpers" = Table.RemoveColumns(#"Title Case Item", {"Brand", "Type", "Variation", "_IsDisplay"}),

    #"Rename Clean Cols" = Table.RenameColumns(#"Remove Helpers", {
        {"Brand_Clean",     "Brand"},
        {"Type_Clean",      "Type"},
        {"Variation_Clean", "Variation"}
    }),

    // ── STEP 8: Reorder columns to match original layout ──
    #"Reorder Columns" = Table.ReorderColumns(#"Rename Clean Cols", {
        "Brand", "Type", "Variation", "Item",
        "Calendar Year", "Calendar Month", "Month Name",
        "Customer", "Unit", "CAD"
    }),

    // ── STEP 9: Final sort ──
    #"Sorted" = Table.Sort(#"Reorder Columns", {
        {"Calendar Year",  Order.Ascending},
        {"Calendar Month", Order.Ascending},
        {"Brand",          Order.Ascending}
    })

in
    #"Sorted"
