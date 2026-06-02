# Power Query M-kód

Tento M-kód načíta všetky `.xlsx` súbory z priečinka definovaného pomenovanou bunkou `pFolderHodnotitelia`.

V `master_template.xlsx` je v hárku `Konfiguracia` bunka `B2`, ktorú treba pomenovať ako `pFolderHodnotitelia` cez:

```text
Vzorce → Definovať názov
```

Potom v Power Query vytvorte nový prázdny dotaz a vložte tento kód do Rozšíreného editora.

```powerquery
let
    FolderPath = Excel.CurrentWorkbook(){[Name="pFolderHodnotitelia"]}[Content]{0}[Column1],
    Zdroj = Folder.Files(FolderPath),
    #"Filtrované súbory" = Table.SelectRows(Zdroj, each ([Extension] = ".xlsx") and not Text.StartsWith([Name], "~$")),
    #"Ponechané stĺpce" = Table.SelectColumns(#"Filtrované súbory", {"Content", "Name"}),
    #"Pridané dáta" = Table.AddColumn(#"Ponechané stĺpce", "Data", each Excel.Workbook([Content], true)),
    #"Rozbalené workbooky" = Table.ExpandTableColumn(#"Pridané dáta", "Data", {"Name", "Data", "Kind"}, {"Data.Name", "Data.Data", "Data.Kind"}),
    #"Iba tblHodnotenie" = Table.SelectRows(#"Rozbalené workbooky", each ([Data.Name] = "tblHodnotenie") and ([Data.Kind] = "Table")),
    #"Rozbalené hodnotenie" = Table.ExpandTableColumn(#"Iba tblHodnotenie", "Data.Data", {"Uchádzač", "Projekt", "Portfólio"}, {"Uchádzač", "Projekt", "Portfólio"}),
    #"Odstránená prípona" = Table.ReplaceValue(#"Rozbalené hodnotenie", ".xlsx", "", Replacer.ReplaceText, {"Name"}),
    #"Premenovaný hodnotiteľ" = Table.RenameColumns(#"Odstránená prípona", {{"Name", "Hodnotiteľ"}}),
    #"Zrušená kontingenčnosť stĺpcov" = Table.Unpivot(#"Premenovaný hodnotiteľ", {"Projekt", "Portfólio"}, "Atribút", "Hodnota"),
    #"Pridaný výstupný stĺpec" = Table.AddColumn(#"Zrušená kontingenčnosť stĺpcov", "StĺpecVýstup", each [Atribút] & "_" & [Hodnotiteľ], type text),
    #"Ponechané pre pivot" = Table.SelectColumns(#"Pridaný výstupný stĺpec", {"Uchádzač", "Hodnota", "StĺpecVýstup"}),
    #"Kontingenčný stĺpec" = Table.Pivot(#"Ponechané pre pivot", List.Distinct(#"Ponechané pre pivot"[StĺpecVýstup]), "StĺpecVýstup", "Hodnota", List.Max)
in
    #"Kontingenčný stĺpec"
```
