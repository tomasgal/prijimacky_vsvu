# Hodnotenie uchádzačov – Excel/Power Query nástroje

Tento repozitár obsahuje pracovné podklady pre zber a spracovanie hodnotení uchádzačov v Exceli pomocou Power Query.

Každý hodnotiteľ vypĺňa samostatný Excel súbor. Všetky hodnotiteľské súbory sú uložené v priečinku `hodnotitelia/` a master súbor ich cez Power Query načíta do spoločnej tabuľky.

## Aktuálna verzia: VIK_2026-27

Aktuálna verzia pracuje so štyrmi hodnotiacimi kritériami:

- `Projekt`
- `Portfólio`
- `Motivačný list`
- `Pohovor`

Každé kritérium sa hodnotí bodmi `0–5`. Prázdne hodnotiace bunky sa vo výstupe zachovajú ako prázdne bunky. Reálne zadaná hodnota `0` zostáva vo výstupe ako `0`.

Výstupný master má jedného uchádzača na jednom riadku. Uchádzači sú zoradení podľa `Priezvisko`, potom `Meno`. Hodnotiace stĺpce sú usporiadané po blokoch v poradí `Projekt`, `Portfólio`, `Motivačný list`, `Pohovor`; vnútri každého bloku sú hodnotitelia zoradení abecedne podľa názvu súboru.

Aktuálny M-kód je v:

```text
docs/power_query_m_code_vik_2026_27.md
```

## Prvá verzia: dvojkritériový prototyp

Ako prvá vznikla jednoduchšia dvojkritériová verzia prototypu. Tá pracovala iba s kritériami:

- `Projekt`
- `Portfólio`

Táto verzia zostáva v repozitári ako dokumentovaný historický prototyp a môže byť užitočná v prípadoch, keď stačí jednoduchší zber hodnotení s dvomi hodnotiacimi stĺpcami.

Dokumentácia a Power Query skript pre dvojkritériovú verziu sú v:

```text
docs/power_query_m_code.md
```

## Štruktúra hodnotiteľského súboru

Každý hodnotiteľský súbor v aktuálnej verzii musí obsahovať Excel tabuľku s názvom:

```text
tblHodnotenie
```

Tabuľka musí mať presné názvy stĺpcov:

```text
Ateliér | Priezvisko | Meno | kód | Projekt | Portfólio | Motivačný list | Pohovor
```

Názov súboru sa používa ako identifikátor hodnotiteľa. Odporúčaný anonymizovaný tvar názvov súborov je:

```text
hodnotitel_01.xlsx
hodnotitel_02.xlsx
hodnotitel_03.xlsx
```

Súbor `hodnotitel_01.xlsx` vytvorí vo výstupe stĺpce:

```text
Projekt_hodnotitel_01
Portfólio_hodnotitel_01
Motivačný list_hodnotitel_01
Pohovor_hodnotitel_01
```

## Odporúčaná štruktúra priečinkov

```text
VIK_2026-27/
├── VIK_2026-27.xlsx
└── hodnotitelia/
    ├── hodnotitel_01.xlsx
    ├── hodnotitel_02.xlsx
    └── ...
```

Master súbor `VIK_2026-27.xlsx` má mať konfiguračný hárok s pomenovanou bunkou:

```text
pFolderHodnotitelia
```

Táto bunka obsahuje cestu k priečinku `hodnotitelia`.

## Dokumentácia

```text
docs/power_query_m_code_vik_2026_27.md
docs/power_query_m_code.md
docs/navod_na_pouzitie.md
```

## Dôležité upozornenie

Do verejného repozitára nedávajte reálne mená uchádzačov, hodnotiteľov ani reálne bodové hodnotenia. Repozitár má slúžiť ako technická dokumentácia a šablóna pracovného postupu.
