# Power Query M-kód pre master VIK_2026-27

Táto verzia je určená pre master súbor ktorý obsahuje štyri kritéria hodnotenia.

Predpoklady:

- master súbor má pomenovanú bunku `pFolderHodnotitelia`, ktorá obsahuje cestu k priečinku s hodnotiteľskými súbormi,
- hodnotiteľské súbory sú vo formáte `.xlsx`,
- každý hodnotiteľský súbor obsahuje Excel tabuľku s názvom `tblHodnotenie`,
- tabuľka `tblHodnotenie` má stĺpce `Ateliér`, `Priezvisko`, `Meno`, `kód`, `Projekt`, `Portfólio`, `Motivačný list`, `Pohovor`,
- názov hodnotiteľa sa berie z názvu súboru, napr. `hodnotitel_01.xlsx` → `hodnotitel_01`,
- prázdne hodnotiace bunky sa vo výstupe zachovajú ako prázdne bunky,
- reálne zadané hodnotenie `0` zostane vo výstupe ako `0`,
- každý uchádzač má vo výstupe jeden riadok,
- riadky sú zoradené podľa `Priezvisko`, potom `Meno`,
- výstupné hodnotiace stĺpce sú zoradené podľa kritérií v poradí `Projekt`, `Portfólio`, `Motivačný list`, `Pohovor`; vnútri každého kritéria sú hodnotitelia zoradení abecedne podľa názvu súboru.

Odporúčané anonymizované názvy hodnotiteľských súborov:

```text
hodnotitel_01.xlsx
hodnotitel_02.xlsx
hodnotitel_03.xlsx
```

Kód vložte do Power Query cez `Domov → Rozšírený editor`.

```powerquery
let
    // Cesta k priečinku s hodnotiteľskými súbormi sa neukladá napevno v M-kóde.
    // Načíta sa z pomenovanej bunky pFolderHodnotitelia v master zošite.
    FolderPath = Excel.CurrentWorkbook(){[Name="pFolderHodnotitelia"]}[Content]{0}[Column1],
    Zdroj = Folder.Files(FolderPath),

    // Načítajú sa iba skutočné .xlsx súbory.
    // Dočasné Excel súbory začínajúce na ~$ sa ignorujú.
    #"Len Excel súbory" =
        Table.SelectRows(
            Zdroj,
            each [Extension] = ".xlsx" and not Text.StartsWith([Name], "~$")
        ),

    // Pre ďalšie spracovanie stačí binárny obsah súboru a jeho názov.
    #"Ponechané stĺpce" =
        Table.SelectColumns(
            #"Len Excel súbory",
            {"Content", "Name"}
        ),

    // Z každého súboru sa načítajú excelovské objekty: tabuľky, hárky a pomenované rozsahy.
    #"Pridané Excel tabuľky" =
        Table.AddColumn(
            #"Ponechané stĺpce",
            "ExcelData",
            each Excel.Workbook([Content], true)
        ),

    #"Rozbalené Excel tabuľky" =
        Table.ExpandTableColumn(
            #"Pridané Excel tabuľky",
            "ExcelData",
            {"Name", "Data", "Kind"},
            {"TableName", "Data", "Kind"}
        ),

    // Z každého hodnotiteľského súboru sa berie iba tabuľka tblHodnotenie.
    #"Len tblHodnotenie" =
        Table.SelectRows(
            #"Rozbalené Excel tabuľky",
            each [TableName] = "tblHodnotenie" and [Kind] = "Table"
        ),

    // Rozbalenie dát z tabuľky tblHodnotenie.
    #"Rozbalené hodnotenia" =
        Table.ExpandTableColumn(
            #"Len tblHodnotenie",
            "Data",
            {
                "Ateliér",
                "Priezvisko",
                "Meno",
                "kód",
                "Projekt",
                "Portfólio",
                "Motivačný list",
                "Pohovor"
            },
            {
                "Ateliér",
                "Priezvisko",
                "Meno",
                "kód",
                "Projekt",
                "Portfólio",
                "Motivačný list",
                "Pohovor"
            }
        ),

    // Názov súboru sa použije ako anonymizovaný identifikátor hodnotiteľa.
    #"Premenovaný súbor na hodnotiteľa" =
        Table.RenameColumns(
            #"Rozbalené hodnotenia",
            {{"Name", "Hodnotiteľ"}}
        ),

    // Odstráni sa prípona .xlsx, napr. hodnotitel_01.xlsx sa zmení na hodnotitel_01.
    #"Odstránená prípona xlsx" =
        Table.ReplaceValue(
            #"Premenovaný súbor na hodnotiteľa",
            ".xlsx",
            "",
            Replacer.ReplaceText,
            {"Hodnotiteľ"}
        ),

    // Power Query pri odpivotovaní štandardne vynecháva prázdne hodnoty.
    // Preto sa prázdne hodnotiace bunky dočasne nahradia technickou hodnotou -1.
    // Po pivote sa -1 vráti späť na null, teda prázdnu bunku.
    // Reálne hodnotenie 0 tým zostane odlíšené od prázdnej bunky.
    #"Prázdne body ako mínus jedna" =
        Table.ReplaceValue(
            #"Odstránená prípona xlsx",
            null,
            -1,
            Replacer.ReplaceValue,
            {
                "Projekt",
                "Portfólio",
                "Motivačný list",
                "Pohovor"
            }
        ),

    // Štyri hodnotiace kritériá sa prevedú zo stĺpcov do riadkov.
    #"Odpivotované kritériá" =
        Table.UnpivotOtherColumns(
            #"Prázdne body ako mínus jedna",
            {
                "Content",
                "Hodnotiteľ",
                "TableName",
                "Kind",
                "Ateliér",
                "Priezvisko",
                "Meno",
                "kód"
            },
            "Kritérium",
            "Body"
        ),

    // Pomocné poradie kritérií pre výsledné usporiadanie stĺpcov.
    #"Poradie kritérií" =
        Table.AddColumn(
            #"Odpivotované kritériá",
            "PoradieKritéria",
            each
                if [Kritérium] = "Projekt" then 1
                else if [Kritérium] = "Portfólio" then 2
                else if [Kritérium] = "Motivačný list" then 3
                else if [Kritérium] = "Pohovor" then 4
                else 99,
            Int64.Type
        ),

    // Výsledný názov stĺpca má tvar Kritérium_Hodnotiteľ, napr. Projekt_hodnotitel_01.
    #"Pridaný názov výstupného stĺpca" =
        Table.AddColumn(
            #"Poradie kritérií",
            "StĺpecVýstup",
            each [Kritérium] & "_" & [Hodnotiteľ],
            type text
        ),

    // Zoradenie ešte pred pivotom určí poradie výsledných hodnotiacich stĺpcov.
    // Najprv ide poradie kritérií, vnútri kritéria hodnotitelia abecedne podľa názvu súboru.
    #"Zoradené pred pivotom" =
        Table.Sort(
            #"Pridaný názov výstupného stĺpca",
            {
                {"PoradieKritéria", Order.Ascending},
                {"Hodnotiteľ", Order.Ascending}
            }
        ),

    // Explicitný zoznam výsledných stĺpcov pre pivot v správnom poradí.
    ZoznamStlpcov =
        List.Distinct(#"Zoradené pred pivotom"[StĺpecVýstup]),

    // Technické stĺpce sa pred pivotom odstránia.
    #"Odstránené technické stĺpce" =
        Table.RemoveColumns(
            #"Zoradené pred pivotom",
            {
                "Content",
                "TableName",
                "Kind",
                "Hodnotiteľ",
                "Kritérium",
                "PoradieKritéria"
            }
        ),

    #"Zmenený typ bodov" =
        Table.TransformColumnTypes(
            #"Odstránené technické stĺpce",
            {{"Body", Int64.Type}}
        ),

    // Pivot vytvorí jeden riadok na uchádzača a samostatné stĺpce pre kritérium/hodnotiteľa.
    #"Kontingenčný stĺpec" =
        Table.Pivot(
            #"Zmenený typ bodov",
            ZoznamStlpcov,
            "StĺpecVýstup",
            "Body",
            List.Max
        ),

    // Technická hodnota -1 sa po pivote vráti späť na null, teda prázdnu bunku.
    #"Mínus jedna späť na prázdne" =
        Table.ReplaceValue(
            #"Kontingenčný stĺpec",
            -1,
            null,
            Replacer.ReplaceValue,
            ZoznamStlpcov
        ),

    // Finálne zoradenie uchádzačov podľa abecedy.
    #"Zoradené riadky" =
        Table.Sort(
            #"Mínus jedna späť na prázdne",
            {
                {"Priezvisko", Order.Ascending},
                {"Meno", Order.Ascending}
            }
        )

in
    #"Zoradené riadky"
```
