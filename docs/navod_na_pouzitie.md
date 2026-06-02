# Návod na použitie: zber hodnotení uchádzačov cez Excel a Power Query

Tento návod je určený pre používateľa, ktorý bežne ovláda MS Excel, ale nemusí byť informatik. Riešenie slúži na to, aby viacerí hodnotitelia vyplnili body pre uchádzačov vo vlastných Excel súboroch a aby sa výsledky automaticky zozbierali do jedného master súboru.

## 1. Čo riešenie robí

Každý hodnotiteľ dostane vlastný Excel súbor. V tomto súbore je zoznam uchádzačov a dva hodnotiace stĺpce:

```text
Uchádzač | Projekt | Portfólio
```

Hodnotiteľ vyberá body v stĺpcoch `Projekt` a `Portfólio` z rozbaľovacieho zoznamu `0–5`. Nemá meniť mená uchádzačov ani štruktúru tabuľky.

Administrátor potom otvorí súbor `master_template.xlsx` a použije funkciu:

```text
Údaje → Obnoviť všetko
```

Master súbor načíta všetky hodnotiteľské súbory z priečinka `hodnotitelia/` a vytvorí spoločnú výslednú tabuľku. V nej je každý uchádzač v jednom riadku a hodnotenia od jednotlivých hodnotiteľov sú v samostatných stĺpcoch.

## 2. Odporúčaná štruktúra priečinkov

Balík má mať túto štruktúru:

```text
prijimacky_vsvu/
├── master_template.xlsx
├── hodnotitel_template.xlsx
├── hodnotitelia/
│   ├── hodnotitel_01_sample.xlsx
│   └── hodnotitel_02_sample.xlsx
└── docs/
    └── navod_na_pouzitie.md
```

V praxi budú v priečinku `hodnotitelia/` uložené súbory jednotlivých hodnotiteľov, napríklad:

```text
hodnotitel_01.xlsx
hodnotitel_02.xlsx
hodnotitel_03.xlsx
```

Názov súboru sa používa ako identifikátor hodnotiteľa. Ak sa súbor volá `hodnotitel_03.xlsx`, v master tabuľke vzniknú napríklad stĺpce:

```text
Projekt_hodnotitel_03
Portfólio_hodnotitel_03
```

## 3. Ako pripraviť hodnotiteľský súbor

Najjednoduchší postup je použiť existujúci súbor `hodnotitel_template.xlsx` a skopírovať ho pre každého hodnotiteľa.

Postup:

1. Skopírujte `hodnotitel_template.xlsx`.
2. Kópiu premenujte napríklad na `hodnotitel_01.xlsx`.
3. Súbor uložte do priečinka `hodnotitelia/`.
4. Rovnaký postup zopakujte pre ďalších hodnotiteľov.

Každý hodnotiteľský súbor musí obsahovať tabuľku s názvom:

```text
tblHodnotenie
```

a musí mať presne tieto stĺpce:

```text
Uchádzač | Projekt | Portfólio
```

Názvy stĺpcov nemeňte. Ak sa zmení názov stĺpca, master súbor ho nemusí vedieť načítať.

## 4. Ako hodnotiteľ vypĺňa body

Hodnotiteľ otvorí svoj súbor, napríklad:

```text
hodnotitel_01.xlsx
```

V tabuľke vyplní iba stĺpce:

```text
Projekt
Portfólio
```

V bunkách je rozbaľovacie menu s hodnotami:

```text
0, 1, 2, 3, 4, 5
```

Hodnotiteľ vyberie príslušný počet bodov pre každého uchádzača. Po vyplnení súbor uloží.

Hodnotiteľ nemá meniť názvy stĺpcov, názov tabuľky, mená uchádzačov ani štruktúru hárka.

## 5. Ako administrátor obnoví master súbor

Administrátor otvorí:

```text
master_template.xlsx
```

Potom použije:

```text
Údaje → Obnoviť všetko
```

Excel načíta všetky `.xlsx` súbory z priečinka `hodnotitelia/`. Výsledkom je tabuľka, v ktorej je každý uchádzač v jednom riadku a hodnotenia sú v samostatných stĺpcoch.

Príklad výsledku:

```text
Uchádzač | Projekt_hodnotitel_01 | Portfólio_hodnotitel_01 | Projekt_hodnotitel_02 | Portfólio_hodnotitel_02
Uchádzač 001 | 4 | 5 | 3 | 4
Uchádzač 002 | 5 | 4 | 4 | 5
```

Ak niektorá bunka obsahuje `null`, znamená to, že príslušné hodnotenie ešte nebolo vyplnené alebo bolo v hodnotiteľskom súbore prázdne.

## 6. Ako pridať nového hodnotiteľa

Ak pribudne nový hodnotiteľ:

1. Skopírujte `hodnotitel_template.xlsx`.
2. Premenujte kópiu napríklad na:

```text
hodnotitel_03.xlsx
```

3. Uložte ju do priečinka:

```text
hodnotitelia/
```

4. Hodnotiteľ vyplní body.
5. V master súbore použite:

```text
Údaje → Obnoviť všetko
```

Master automaticky pridá nové stĺpce pre hodnotiteľa podľa názvu súboru.

## 7. Ako vymeniť zoznam uchádzačov

Ak sa má použiť nový zoznam uchádzačov, treba ho vložiť do hodnotiteľskej šablóny tak, aby ostala zachovaná tabuľka `tblHodnotenie`.

Odporúčaný postup:

1. Otvorte `hodnotitel_template.xlsx`.
2. V stĺpci `Uchádzač` nahraďte vzorové mená novým zoznamom uchádzačov.
3. Skontrolujte, že nové mená sú stále vo vnútri tabuľky `tblHodnotenie`.
4. Stĺpce `Projekt` a `Portfólio` ponechajte s rozbaľovacím menu `0–5`.
5. Súbor uložte ako novú šablónu alebo ho skopírujte pre jednotlivých hodnotiteľov.

Ak je zoznam uchádzačov dlhší alebo kratší než pôvodný, je potrebné skontrolovať, či rozbaľovacie menu a prípadná ochrana hárka platia pre všetky riadky, ktoré majú hodnotitelia vypĺňať.

## 8. Dôležité pravidlá

Dodržujte tieto pravidlá:

1. Všetky hodnotiteľské súbory musia byť v priečinku `hodnotitelia/`.
2. Všetky hodnotiteľské súbory musia byť vo formáte `.xlsx`.
3. Každý hodnotiteľský súbor musí obsahovať tabuľku `tblHodnotenie`.
4. Tabuľka musí obsahovať stĺpce `Uchádzač`, `Projekt`, `Portfólio`.
5. Názvy stĺpcov nemeňte.
6. Názvy hodnotiteľských súborov musia byť jednoznačné.
7. Do priečinka `hodnotitelia/` nevkladajte iné pracovné Excel súbory, ktoré nemajú rovnakú štruktúru.
8. Súbory začínajúce `~$` sú dočasné Excel súbory a master ich ignoruje.

## 9. Verejný repozitár a ochrana dát

Tento repozitár má obsahovať iba anonymizované alebo vzorové dáta. Do verejného repozitára nepatria:

- reálne mená uchádzačov,
- reálne hodnotenia,
- osobné údaje,
- produkčné hodnotiteľské súbory,
- interné pracovné verzie prijímacieho konania.

Reálne pracovné dáta ukladajte mimo verejný GitHub repozitár, napríklad do interného úložiska organizácie alebo do chráneného SharePoint/OneDrive priestoru.

## 10. Najčastejšie problémy

### Master nenačítal nový hodnotiteľský súbor

Skontrolujte, či je nový súbor uložený v priečinku `hodnotitelia/`, či má príponu `.xlsx` a či obsahuje tabuľku `tblHodnotenie`.

### V masteri sú prázdne hodnoty alebo `null`

To zvyčajne znamená, že hodnotiteľ príslušné bunky nevyplnil alebo súbor neuložil. Otvorte hodnotiteľský súbor, skontrolujte hodnoty a súbor uložte.

### Jeden uchádzač sa zobrazuje viackrát

Môže ísť o rozdiel v zápise mena, napríklad medzera navyše, iná diakritika alebo preklep. Pri produkčnom použití je vhodné pridať aj jednoznačný identifikátor uchádzača, napríklad `ID_uchádzača`.

### Master hlási chybu pri obnovovaní

Najčastejšie príčiny sú zmenený názov tabuľky, zmenené názvy stĺpcov alebo nesprávna cesta k priečinku `hodnotitelia/`.

## 11. Poznámka k jednoznačnej identifikácii uchádzačov

V základnej verzii sa uchádzači identifikujú podľa textu v stĺpci `Uchádzač`. To je jednoduché, ale nie úplne bezpečné, ak sa v zozname vyskytnú dvaja uchádzači s rovnakým menom alebo ak sa meno niekde zapíše mierne odlišne.

Pre ostrejšie alebo produkčné použitie sa odporúča doplniť stĺpec:

```text
ID_uchádzača
```

Potom by tabuľka mala mať tvar:

```text
ID_uchádzača | Uchádzač | Projekt | Portfólio
```

Táto verzia je robustnejšia, ale vyžaduje aj úpravu Power Query dotazu v master súbore.

## 12. Odporúčaný pracovný postup

Pre bežné použitie odporúčame tento postup:

1. Pripraviť jeden správny hodnotiteľský template.
2. Skopírovať ho pre každého hodnotiteľa.
3. Každému hodnotiteľovi prideliť jeho vlastný súbor.
4. Po vyplnení všetky súbory uložiť do priečinka `hodnotitelia/`.
5. Otvoriť `master_template.xlsx`.
6. Použiť `Údaje → Obnoviť všetko`.
7. Skontrolovať výslednú tabuľku.
8. Až z výslednej tabuľky vypočítavať súčty, priemery, poradie alebo rozhodnutie o postupe.
