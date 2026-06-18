# Master súbor VIK_2026-27

Tento dokument opisuje aktuálnu štruktúru master súboru `VIK_2026-27.xlsx`. Do repozitára sa nemá ukladať produkčný master so skutočnými menami uchádzačov alebo hodnotiteľov. Dokumentácia preto používa anonymizované názvy hodnotiteľov, napríklad `hodnotitel_01.xlsx`.

## Základná štruktúra

Master súbor obsahuje hárok `hodnotitelia`, na ktorom je výsledná Excel tabuľka `Hodnotitelia`. Dáta z hodnotiteľských súborov sa načítavajú cez Power Query z podpriečinka `hodnotitelia/`. Cesta k tomuto priečinku je uložená v pomenovanej bunke `pFolderHodnotitelia` na hárku `Konfiguracia`.

Power Query časť tabuľky obsahuje identifikačné údaje uchádzača a hodnotiace stĺpce. Pri aktuálnej konfigurácii sa používajú štyri kritériá:

```text
Projekt
Portfólio
Motivačný list
Pohovor
```

Pre každého hodnotiteľa vznikne jeden stĺpec pre každé kritérium. Pri ôsmich hodnotiteľoch teda vznikne 32 hodnotiacich stĺpcov. Stĺpce sú zoradené najprv podľa kritéria a vnútri kritéria podľa názvu hodnotiteľského súboru.

Typický anonymizovaný tvar výsledných stĺpcov je:

```text
Projekt_hodnotitel_01
Projekt_hodnotitel_02
...
Portfólio_hodnotitel_01
Portfólio_hodnotitel_02
...
Motivačný list_hodnotitel_01
...
Pohovor_hodnotitel_01
...
```

## Vlastné stĺpce v master tabuľke

Za importovanými Power Query stĺpcami sú v master tabuľke pridané vlastné pracovné stĺpce:

```text
Súčet bodov
Poradie
Návrh komisie
Ateliér - návrh prijať
```

Tieto stĺpce sú súčasťou tej istej Excel tabuľky `Hodnotitelia`. Ak sa do priečinka `hodnotitelia/` pridá ďalší hodnotiteľský `.xlsx` súbor a Power Query sa obnoví, pribudnú štyri nové hodnotiace stĺpce. Excel by mal vlastné stĺpce posunúť doprava. Po takejto zmene treba skontrolovať najmä vzorec v stĺpci `Súčet bodov`, pretože práve tento stĺpec určuje aj automatický návrh komisie.

## Odporúčaný dynamický súčet bodov

Stĺpec `Súčet bodov` má sčítať všetky hodnotiace stĺpce v aktuálnom riadku od prvého bodového stĺpca po posledný bodový stĺpec pred súčtom. Odporúčaný vzorec pre prvý dátový riadok, ak prvé hodnotiace stĺpce začínajú v stĺpci `E`, je:

```excel
=SUM($E2:INDEX(2:2;1;COLUMN()-1))
```

Tento vzorec neobsahuje pevnú hranicu typu `AJ`. Po pridaní ďalšieho hodnotiteľa a posunutí stĺpca `Súčet bodov` doprava sa koniec sčítavaného rozsahu odvodí od stĺpca tesne pred aktuálnym stĺpcom súčtu. Vzorec treba skopírovať do všetkých dátových riadkov tabuľky.

## Poradie

Stĺpec `Poradie` počíta poradie uchádzača podľa stĺpca `Súčet bodov`. V tabuľke možno použiť štruktúrovaný vzorec:

```excel
=RANK.EQ([@[Súčet bodov]];Hodnotitelia[Súčet bodov];0)
```

Tento vzorec dáva rovnaké poradie uchádzačom s rovnakým počtom bodov a následne preskočí ďalšie číslo poradia. Ide teda o štandardné súťažné poradie.

## Dynamický návrh komisie

Stĺpec `Návrh komisie` obsahuje automatický návrh rozhodnutia. Pravidlo je:

```text
0 bodov → nedostavil sa
aspoň 60 % maxima → prijať
menej ako 60 % maxima → neprijať
```

Pri aktuálnej konfigurácii s ôsmimi hodnotiteľmi a štyrmi kritériami je maximum `8 × 4 × 5 = 160`, takže 60 % je `96`. Tento prah však nemá byť zapísaný napevno. Odporúčaný vzorec pre prvý dátový riadok, ak je `Súčet bodov` aktuálne v stĺpci `AK`, je:

```excel
=IF(AK2=0;"nedostavil sa";IF(AK2>=0,6*(COLUMN(AK2)-COLUMN($E2))*5;"prijať";"neprijať"))
```

Ak Excel používa desatinnú bodku, možno použiť variant:

```excel
=IF(AK2=0;"nedostavil sa";IF(AK2>=0.6*(COLUMN(AK2)-COLUMN($E2))*5;"prijať";"neprijať"))
```

Vzorec počíta počet hodnotiacich stĺpcov podľa vzdialenosti medzi prvým hodnotiacim stĺpcom `E` a aktuálnym stĺpcom `Súčet bodov`. Ak pribudne ďalší hodnotiteľ a `Súčet bodov` sa posunie doprava, hranica 60 % sa prepočíta automaticky. Pravidlo pre 0 bodov zostáva nadradené.

Ak sa automatický návrh v konkrétnom riadku manuálne zmení cez rozbaľovací zoznam, vzorec v tejto bunke sa prepíše. Je to v poriadku, ak ide o vedomé rozhodnutie komisie.

## Rozbaľovací zoznam pre návrh komisie

Pre stĺpec `Návrh komisie` je vhodné použiť overenie údajov typu zoznam s hodnotami:

```text
prijať;neprijať;nedostavil sa
```

Automatický vzorec tým nevytvára nové položky v zozname. Vypočíta iba predvolený návrh podľa bodov. Komisia môže hodnotu manuálne zmeniť výberom jednej z povolených možností.

## Podmienené formátovanie súčtu bodov

Stĺpec `Súčet bodov` možno zvýrazniť podmieneným formátovaním, ak je hodnota pod hranicou 60 %. Namiesto pevného čísla, napríklad `96`, sa má použiť dynamický vzorec.

Ak je `Súčet bodov` aktuálne v stĺpci `AK`, pravidlo pre prvý dátový riadok je:

```excel
=AK2<0,6*(COLUMN(AK2)-COLUMN($E2))*5
```

S desatinnou bodkou:

```excel
=AK2<0.6*(COLUMN(AK2)-COLUMN($E2))*5
```

Ak sa majú nuly nechať bez zvýraznenia, možno použiť variant:

```excel
=AND(AK2>0;AK2<0,6*(COLUMN(AK2)-COLUMN($E2))*5)
```

Pravidlo sa má aplikovať iba na dátové bunky stĺpca `Súčet bodov`, nie na celú tabuľku. Staré pravidlá s pevnou hodnotou, napríklad `96`, treba odstrániť cez `Domov → Podmienené formátovanie → Spravovať pravidlá`.

## Pruhované riadky a ručné formátovanie

Excel tabuľka používa tabuľkový štýl so striedaním farby riadkov. Ak sa v niektorom stĺpci ručne nastaví výplň bunky, môže prekryť pruhovanie tabuľky. Na obnovenie pruhovania v danom stĺpci je bezpečnejšie označiť príslušné bunky a zvoliť:

```text
Domov → Farba výplne → Bez výplne
```

Netreba začínať príkazom `Vymazať formáty`, pretože ten môže odstrániť aj podmienené formátovanie alebo iné užitočné formáty. Zároveň treba skontrolovať, že na karte `Návrh tabuľky` je zapnutá voľba `Pruhované riadky`.

## Pridanie alebo odstránenie hodnotiteľa

Hodnotiteľ sa do master tabuľky pridá tak, že sa do priečinka `hodnotitelia/` vloží nový `.xlsx` súbor so správnou tabuľkou `tblHodnotenie`. Po obnovení údajov sa pre neho vytvoria štyri nové hodnotiace stĺpce.

Hodnotiteľa možno dočasne vyradiť premenovaním súboru tak, aby už nemal príponu `.xlsx`, napríklad:

```text
hodnotitel_03.xlsx → hodnotitel_03.bak
```

Power Query načítava iba skutočné `.xlsx` súbory a dočasné súbory Excelu začínajúce na `~$` ignoruje.

Po pridaní alebo odstránení hodnotiteľa treba skontrolovať:

```text
1. či sa vlastné stĺpce posunuli doprava a neboli prepísané,
2. či Súčet bodov zahŕňa všetky aktuálne hodnotiace stĺpce,
3. či Návrh komisie používa dynamický 60 % prah,
4. či podmienené formátovanie neobsahuje staré pevné pravidlo, napríklad 96.
```
