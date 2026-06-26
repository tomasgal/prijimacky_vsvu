# prijimacky_vsvu

Pomocný repozitár pre Excel/Power Query nástroje používané pri spracovaní prijímacích hodnotení na VŠVU. Cieľom je mať opakovateľný a kontrolovateľný spôsob, ako zozbierať hodnotenia viacerých hodnotiteľov z jednotlivých Excel súborov a automaticky ich zlúčiť do jedného master zošita.

Repozitár neobsahuje reálne hodnotiteľské súbory ani produkčný master zošit s osobnými údajmi. Dokumentácia používa anonymizované názvy súborov a technické príklady.

## Obsah

* [Aktuálna verzia: VIK_2026-27](#aktuálna-verzia-vik_2026-27)
* [Súvisiaca dokumentácia](#súvisiaca-dokumentácia)
* [Základná štruktúra priečinkov](#základná-štruktúra-priečinkov)
* [Hodnotiteľské súbory](#hodnotiteľské-súbory)
* [Master zošit](#master-zošit)
* [Power Query spracovanie](#power-query-spracovanie)
* [Vlastné stĺpce v master tabuľke](#vlastné-stĺpce-v-master-tabuľke)
* [Dynamický 60-percentný prah](#dynamický-60-percentný-prah)
* [Pridanie alebo odstránenie hodnotiteľa](#pridanie-alebo-odstránenie-hodnotiteľa)
* [Starší prototyp](#starší-prototyp)
* [Bezpečnostné a anonymizačné pravidlá](#bezpečnostné-a-anonymizačné-pravidlá)

## Aktuálna verzia: VIK_2026-27

Aktuálna verzia pracuje so štyrmi hodnotiacimi kritériami:

* `Projekt`
* `Portfólio`
* `Motivačný list`
* `Pohovor`

Každé kritérium sa hodnotí bodmi od `0` do `5`. Prázdna bunka zostáva prázdnou bunkou a reálna hodnota `0` zostáva hodnotou `0`. Toto je dôležité, pretože prázdna hodnota znamená nevyplnené hodnotenie, zatiaľ čo `0` je platné bodové hodnotenie.

Výsledný master výstup má jeden riadok na jedného uchádzača. Uchádzači sú zoradení podľa `Priezvisko` a potom `Meno`. Hodnotiace stĺpce sú zoradené najprv podľa kritéria a vnútri každého kritéria abecedne podľa názvu súboru hodnotiteľa.

## Súvisiaca dokumentácia

Hlavné dokumenty:

* [Power Query M-kód pre VIK_2026-27](docs/power_query_m_code_vik_2026_27.md)
* [Nastavenie master zošita VIK_2026-27](docs/master_vik_2026_27.md)
* [Návod na použitie](docs/navod_na_pouzitie.md)
* [Starší dvojkritériový Power Query prototyp](docs/power_query_m_code.md)

Odporúčané poradie čítania:

1. Najprv [Návod na použitie](docs/navod_na_pouzitie.md).
2. Potom [Nastavenie master zošita VIK_2026-27](docs/master_vik_2026_27.md).
3. Pri úprave alebo kontrole dotazu použiť [Power Query M-kód pre VIK_2026-27](docs/power_query_m_code_vik_2026_27.md).

## Základná štruktúra priečinkov

Odporúčaná lokálna štruktúra:

```text
VIK_2026-27/
├── VIK_2026-27.xlsx
└── hodnotitelia/
    ├── hodnotitel_01.xlsx
    ├── hodnotitel_02.xlsx
    ├── hodnotitel_03.xlsx
    └── ...
```

Master zošit `VIK_2026-27.xlsx` je uložený v hlavnom priečinku. Hodnotiteľské súbory sú uložené v podpriečinku `hodnotitelia`.

Reálne súbory s osobnými údajmi alebo s identifikovateľnými hodnotiteľmi sa do repozitára neukladajú.

## Hodnotiteľské súbory

Každý hodnotiteľský súbor musí obsahovať Excel tabuľku s názvom:

```text
tblHodnotenie
```

Očakávané stĺpce v tabuľke:

```text
Ateliér | Priezvisko | Meno | kód | Projekt | Portfólio | Motivačný list | Pohovor
```

Hodnotiteľ vypĺňa iba hodnotiace stĺpce:

```text
Projekt
Portfólio
Motivačný list
Pohovor
```

Odporúčané nastavenie hodnotiacich buniek je rozbaľovací zoznam s hodnotami `0` až `5`. Prázdne bunky sú prípustné.

## Master zošit

Master zošit načítava hodnotiteľské súbory cez Power Query z podpriečinka `hodnotitelia`.

V master zošite je odporúčané mať konfiguračný hárok `Konfiguracia`, v ktorom bunka `B1` obsahuje cestu k priečinku s hodnotiteľskými súbormi. Bunka je pomenovaná ako:

```text
pFolderHodnotitelia
```

Príklad vzorca pre slovenské/lokálne nastavenie Excelu:

```excel
=LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1)&"hodnotitelia\"
```

Power Query dotaz potom neobsahuje natvrdo zapísanú lokálnu cestu. Zošit je tým jednoduchšie prenositeľný medzi počítačmi.

Výsledná načítaná tabuľka v master zošite môže mať názov:

```text
Hodnotitelia
```

## Power Query spracovanie

Power Query robí tieto kroky:

1. Načíta cestu k priečinku hodnotiteľov z pomenovanej bunky `pFolderHodnotitelia`.
2. Načíta iba súbory s príponou `.xlsx`.
3. Ignoruje dočasné Excel súbory začínajúce na `~$`.
4. Z každého súboru berie iba tabuľku `tblHodnotenie`.
5. Rozbalí identifikačné stĺpce uchádzača a štyri hodnotiace kritériá.
6. Ako identifikátor hodnotiteľa použije názov súboru bez prípony `.xlsx`.
7. Prázdne hodnotiace bunky dočasne nahradí technickou hodnotou `-1`, aby sa pri odpivotovaní nestratili.
8. Odpivotuje hodnotiace kritériá.
9. Vytvorí výstupné stĺpce v tvare `Kritérium_Hodnotiteľ`, napríklad `Projekt_hodnotitel_01`.
10. Po pivote vráti technickú hodnotu `-1` späť na prázdnu bunku.
11. Zoradí uchádzačov podľa priezviska a mena.

Kompletný M-kód je v dokumente [Power Query M-kód pre VIK_2026-27](docs/power_query_m_code_vik_2026_27.md).

## Vlastné stĺpce v master tabuľke

Za hodnotiace stĺpce je možné v master tabuľke doplniť vlastné pracovné stĺpce. Aktuálne používané vlastné stĺpce sú:

```text
Súčet bodov
Poradie
Návrh komisie
Ateliér - návrh prijať
```

Tieto stĺpce sa nachádzajú napravo od stĺpcov generovaných Power Query. Ak sú súčasťou tej istej Excel tabuľky, Excel sa ich pri obnove dotazu snaží zachovať a v prípade pridania nových hodnotiacich stĺpcov ich posunúť doprava.

Odporúčanie: vlastné stĺpce nevkladať medzi Power Query stĺpce, ale vždy až napravo od nich.

## Dynamický 60-percentný prah

Stĺpec `Súčet bodov` slúži na súčet bodov zo všetkých hodnotiacich stĺpcov. Pri ôsmich hodnotiteľoch a štyroch kritériách je maximum:

```text
8 × 4 × 5 = 160
```

Šesťdesiatpercentný prah je v tomto prípade:

```text
96
```

Ak sa počet hodnotiteľov zmení, prah sa má zmeniť automaticky. Preto je lepšie nepoužívať pevné číslo `96`, ale dynamický výpočet podľa počtu hodnotiacich stĺpcov.

Ak hodnotiace stĺpce začínajú v stĺpci `E` a stĺpec `Súčet bodov` je v aktuálnom stave v stĺpci `AK`, dynamický prah možno vyjadriť takto:

```excel
0,6*(COLUMN(AK2)-COLUMN($E2))*5
```

Vzorec pre stĺpec `Návrh komisie` môže byť:

```excel
=IF(AK2=0;"nedostavil sa";IF(AK2>=0,6*(COLUMN(AK2)-COLUMN($E2))*5;"prijať";"neprijať"))
```

Logika:

* ak je súčet `0`, návrh je `nedostavil sa`;
* ak je súčet aspoň 60 % z aktuálneho maxima, návrh je `prijať`;
* inak je návrh `neprijať`.

Podmienené formátovanie stĺpca `Súčet bodov` môže používať vzorec:

```excel
=AK2<0,6*(COLUMN(AK2)-COLUMN($E2))*5
```

Tým sa zvýraznia uchádzači pod aktuálnou 60-percentnou hranicou. Ak sa po pridaní hodnotiteľa stĺpec `Súčet bodov` posunie doprava, vzorec sa vie spolu s tabuľkou prepočítať na nový počet hodnotiacich stĺpcov.

Ak Excel používa desatinnú bodku namiesto desatinnej čiarky, treba použiť variant s `0.6`.

## Pridanie alebo odstránenie hodnotiteľa

Pridanie hodnotiteľa:

1. Do priečinka `hodnotitelia` sa pridá nový `.xlsx` súbor.
2. Súbor musí obsahovať tabuľku `tblHodnotenie`.
3. Po obnove dotazu Power Query pridá nové stĺpce pre všetky štyri kritériá daného hodnotiteľa.
4. Vlastné stĺpce napravo od hodnotiacich stĺpcov by sa mali posunúť doprava.

Odstránenie hodnotiteľa:

1. Hodnotiteľský súbor sa odstráni z priečinka alebo sa premenuje tak, aby už nemal príponu `.xlsx`, napríklad na `.bak`.
2. Power Query ho pri ďalšej obnove ignoruje.
3. Po obnove sa príslušné hodnotiace stĺpce z master tabuľky odstránia.

Odporúčanie: pred väčšou zmenou počtu hodnotiteľov uložiť kópiu master zošita.

## Tabuľkový štýl a pruhované riadky

Excel tabuľka môže používať pruhované riadky, napríklad so svetlozeleným striedaním riadkov. Toto je vlastnosť štýlu tabuľky, nie bežná ručná výplň buniek.

Ak sa v niektorom vlastnom stĺpci stratí pruhovanie, najčastejšie je príčinou ručne nastavená výplň buniek. Bezpečnejší postup je odstrániť iba ručnú výplň, nie všetky formáty:

```text
Domov → Farba výplne → Bez výplne
```

Neodporúča sa ako prvý krok použiť:

```text
Domov → Vymazať → Vymazať formáty
```

Tento postup môže odstrániť aj podmienené formátovanie alebo iné užitočné nastavenia.

## Starší prototyp

V repozitári zostáva aj starší dvojkritériový Power Query prototyp:

* [Starší dvojkritériový Power Query prototyp](docs/power_query_m_code.md)

Tento prototyp pracoval s menším rozsahom kritérií a slúži ako historická alebo porovnávacia verzia. Aktuálna verzia pre `VIK_2026-27` používa štyri hodnotiace kritériá a je dokumentovaná samostatne.

## Bezpečnostné a anonymizačné pravidlá

Do repozitára nepatria:

* reálne master zošity s osobnými údajmi uchádzačov;
* reálne hodnotiteľské súbory;
* súbory s menami hodnotiteľov, ak nie sú určené na zverejnenie;
* exporty obsahujúce osobné údaje, interné hodnotenia alebo rozhodnutia komisie.

Do dokumentácie patria iba anonymizované príklady, napríklad:

```text
hodnotitel_01.xlsx
hodnotitel_02.xlsx
hodnotitel_03.xlsx
```

Pred každým commitom treba skontrolovať, či sa do repozitára nedostal produkčný `.xlsx` súbor alebo dočasný Excel súbor začínajúci na `~$`.

## Poznámka k zdieľaniu cez OneDrive

Ak sa pracovné dokumenty zdieľajú prostredníctvom infraštruktúry Microsoft OneDrive, musia byť uložené v priečinku, ktorý je skutočne synchronizovaný alebo spravovaný cez OneDrive. Nestačí, aby sa súbory nachádzali na ľubovoľnom mieste lokálneho disku.

Prakticky to znamená, že master zošit aj priečinok `hodnotitelia/` majú byť uložené v príslušnom podadresári OneDrive. Až potom ich možno spoľahlivo zdieľať, synchronizovať a spravovať cez OneDrive.
