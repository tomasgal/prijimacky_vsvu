# Hodnotenie uchádzačov – Excel/Power Query nástroje

Tento repozitár obsahuje pracovné podklady pre zber a spracovanie hodnotení uchádzačov v Exceli pomocou Power Query.

Riešenie je postavené na jednoduchom princípe: každý hodnotiteľ vypĺňa samostatný Excel súbor, všetky hodnotiteľské súbory sú uložené v jednom priečinku a master súbor ich následne načíta do spoločnej tabuľky.

## Aktuálna verzia: VIK_2026-27

Aktuálna verzia pracuje so štyrmi hodnotiacimi kritériami:

- `Projekt`
- `Portfólio`
- `Motivačný list`
- `Pohovor`

Každé kritérium sa hodnotí bodmi `0–5`. Hodnotiteľské bunky môžu zostať prázdne; Power Query ich vo výstupe zachová ako prázdne bunky. Reálne zadaná hodnota `0` zostáva vo výstupe ako `0`, takže sa nerozlišovanie medzi prázdnou bunkou a nulovým hodnotením nestráca.

Výstupný master má jedného uchádzača na jednom riadku. Uchádzači sú zoradení podľa `Priezvisko`, potom `Meno`. Hodnotiace stĺpce sú usporiadané po blokoch v poradí `Projekt`, `Portfólio`, `Motivačný list`, `Pohovor`; vnútri každého bloku sú hodnotitelia zoradení abecedne podľa názvu súboru.

Hlavný M-kód pre túto verziu je uložený v:

```text
docs/power_query_m_code_vik_2026_27.md
```

## Štruktúra hodnotiteľského súboru

Každý hodnotiteľský súbor musí obsahovať Excel tabuľku s názvom:

```text
tblHodnotenie
```

Tabuľka musí mať presné názvy stĺpcov:

```text
Ateliér | Priezvisko | Meno | kód | Projekt | Portfólio | Motivačný list | Pohovor
```

Názov súboru sa používa ako identifikátor hodnotiteľa. Napríklad súbor:

```text
Bálik.xlsx
```

vytvorí vo výstupe stĺpce:

```text
Projekt_Bálik
Portfólio_Bálik
Motivačný list_Bálik
Pohovor_Bálik
```

## Odporúčaná štruktúra priečinkov

```text
VIK_2026-27/
├── VIK_2026-27.xlsx
└── hodnotitelia/
    ├── Bálik.xlsx
    ├── Benčík.xlsx
    └── ...
```

Master súbor `VIK_2026-27.xlsx` má mať konfiguračný hárok s pomenovanou bunkou:

```text
pFolderHodnotitelia
```

Táto bunka obsahuje cestu k priečinku `hodnotitelia`. Power Query potom nemusí mať cestu k priečinku zapísanú napevno v M-kóde.

## Pridanie alebo výmena hodnotiteľského súboru

Hodnotiteľský súbor sa uloží do priečinka `hodnotitelia/`. Po otvorení master súboru stačí použiť:

```text
Údaje → Obnoviť všetko
```

Power Query načíta aktuálny obsah priečinka a z každého `.xlsx` súboru použije tabuľku `tblHodnotenie`. Dočasné Excel súbory začínajúce na `~$` sa ignorujú.

## Dokumentácia

V repozitári sú dve hlavné verzie Power Query kódu:

```text
docs/power_query_m_code.md
```

Staršia verzia prototypu s menším počtom hodnotiacich kritérií.

```text
docs/power_query_m_code_vik_2026_27.md
```

Aktuálna verzia pre master `VIK_2026-27.xlsx`, so štyrmi kritériami a stabilným spracovaním prázdnych buniek.

Používateľský návod je v:

```text
docs/navod_na_pouzitie.md
```

## Dôležité upozornenie

Do verejného repozitára nedávajte reálne mená uchádzačov, hodnotiteľov ani reálne bodové hodnotenia. Tento repozitár má slúžiť ako technická dokumentácia a šablóna pracovného postupu, nie ako úložisko produkčných osobných údajov.
