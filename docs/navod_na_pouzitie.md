# Návod na použitie: zber hodnotení uchádzačov cez Excel a Power Query

Tento návod je určený pre používateľa, ktorý bežne pracuje s MS Excelom, ale nemusí byť informatik. Cieľom riešenia je pripraviť jednoduchý a kontrolovateľný spôsob, ako môžu viacerí hodnotitelia vyplniť body pre uchádzačov a ako sa tieto hodnotenia následne automaticky zozbierajú do jedného výsledného Excel súboru.

Riešenie je postavené na bežných funkciách MS Excelu: excelovskej tabuľke, rozbaľovacích zoznamoch, uzamknutí hárka a Power Query. Hodnotiteľ nemusí rozumieť Power Query. Hodnotiteľ iba otvorí svoj súbor a vyplní body. Power Query používa administrátor v master súbore na načítanie výsledkov.

## 1. Čo riešenie robí

Každý hodnotiteľ dostane vlastný Excel súbor. V tomto súbore je zoznam uchádzačov a dva hodnotiace stĺpce:

```text
Uchádzač | Projekt | Portfólio
```

Hodnotiteľ v stĺpcoch `Projekt` a `Portfólio` vyberá body z rozbaľovacieho zoznamu `0–5`. Nemá meniť mená uchádzačov, názvy stĺpcov ani štruktúru hárka.

Administrátor potom otvorí súbor `master_template.xlsx` a použije:

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

Preto je vhodné pomenúvať hodnotiteľské súbory konzistentne, bez diakritiky a bez medzier. Odporúčaný tvar je napríklad `hodnotitel_01.xlsx`, `hodnotitel_02.xlsx`, prípadne anonymizované označenia `h01.xlsx`, `h02.xlsx`.

## 3. Dôležitá poznámka k ceste k priečinku `hodnotitelia/`

Toto je jedna z najdôležitejších technických častí celého riešenia. Power Query, teda excelovský nástroj, ktorý načítava hodnotiteľské súbory do master súboru, si pri použití príkazu `Z priečinka` spravidla zapíše **absolútnu cestu** k priečinku. To znamená, že si nezapamätá iba informáciu „priečinok hodnotitelia vedľa master súboru“, ale celú konkrétnu cestu v počítači.

Príklad absolútnej cesty:

```text
C:\Users\Admin\Documents\phd prijimacky\hodnotitelia
```

Takáto cesta funguje na počítači, na ktorom bola vytvorená. Ak však celý balík presuniete do iného priečinka, na USB kľúč, na iný počítač alebo ho pošlete niekomu inému, táto cesta už nemusí existovať. Master súbor sa potom pri obnove údajov môže pokúsiť hľadať hodnotiteľské súbory na pôvodnom mieste, napríklad v priečinku používateľa `Admin`, ktorý na inom počítači vôbec nemusí existovať.

Preto treba rozlišovať dve situácie:

```text
Lokálny prototyp: treba dávať pozor na cestu k priečinku hodnotitelia.
Cloudové použitie: ideálne je mať súbory na stabilnom mieste v OneDrive alebo SharePointe.
```

Pri lokálnom prenášaní prototypu je vhodné použiť konfiguračný hárok a pomenovanú bunku, ktorá obsahuje cestu k priečinku `hodnotitelia/`. Vďaka tomu netreba upravovať Power Query kód pri každom presune. Stačí upraviť jednu bunku v Exceli.

## 4. Optimálne použitie: OneDrive alebo SharePoint

Najpraktickejšie a dlhodobo najčistejšie použitie tohto riešenia je v cloude, teda v organizačnom OneDrive alebo SharePointe. V takom prípade sú `master_template.xlsx` aj priečinok `hodnotitelia/` uložené na stabilnom spoločnom mieste a administrátor má nad nimi kontrolu.

Typická cloudová štruktúra môže vyzerať takto:

```text
SharePoint / OneDrive
└── prijimacky_vsvu/
    ├── master.xlsx
    └── hodnotitelia/
        ├── hodnotitel_01.xlsx
        ├── hodnotitel_02.xlsx
        └── hodnotitel_03.xlsx
```

Výhodou cloudového použitia je, že hodnotiteľské súbory možno sprístupniť konkrétnym hodnotiteľom cez webový Excel. Každý hodnotiteľ môže mať prístup iba k svojmu súboru. Administrátor potom v master súbore obnoví údaje a načíta hodnotenia zo všetkých súborov.

Cloudové použitie je vhodnejšie najmä vtedy, ak:

1. na hodnotení pracuje viac ľudí,
2. súbory sa nemajú posielať e-mailom,
3. hodnotitelia majú vypĺňať svoje hodnotenia cez web,
4. treba udržať poriadok v prístupových právach,
5. treba znížiť riziko, že niekto pracuje so starou verziou súboru.

Lokálny ZIP balík alebo lokálny priečinok je vhodný najmä na prototyp, ukážku alebo jednorazové spracovanie. Pre reálne pracovné použitie je cloud spravidla bezpečnejší a prehľadnejší.

## 5. Konfiguračný hárok a pomenovaná bunka `pFolderHodnotitelia`

Aby master súbor nebol pevne naviazaný na jednu absolútnu cestu v jednom konkrétnom počítači, používa sa konfiguračný hárok. Ide o obyčajný hárok v Exceli, napríklad s názvom:

```text
Konfiguracia
```

Na tomto hárku je bunka, v ktorej je uvedená cesta k priečinku s hodnotiteľskými súbormi. Táto bunka je pomenovaná:

```text
pFolderHodnotitelia
```

Pomenovaná bunka znamená, že Excel sa na konkrétnu bunku nevie odkazovať iba súradnicami, napríklad `B2`, ale aj menom. Power Query potom nečíta cestu priamo z kódu, ale z tejto pomenovanej bunky.

Príklad obsahu bunky:

```text
C:\Users\Admin\Documents\phd prijimacky\hodnotitelia\
```

alebo pri inom umiestnení:

```text
D:\pracovne\prijimacky_vsvu\hodnotitelia\
```

Ak sa balík presunie, administrátor nemusí hľadať a upravovať Power Query kód. Stačí zmeniť cestu v konfiguračnej bunke.

## 6. Ako vytvoriť konfiguračný hárok

Ak master súbor ešte nemá konfiguračný hárok, vytvorte ho takto:

1. V master súbore pridajte nový hárok.
2. Pomenujte ho napríklad:

```text
Konfiguracia
```

3. Do bunky `A1` napíšte napríklad:

```text
Priečinok s hodnotiteľskými súbormi
```

4. Do bunky `B1` alebo `B2` vložte cestu k priečinku `hodnotitelia/`.

Príklad:

```text
C:\Users\Admin\Documents\phd prijimacky\hodnotitelia\
```

Ak chcete, aby sa cesta automaticky odvodzovala od miesta, kde je uložený master súbor, možno použiť vzorec. V slovenskom alebo európskom nastavení Excelu môže vyzerať napríklad takto:

```excel
=LEFT(CELL("filename";A1);FIND("[";CELL("filename";A1))-1)&"hodnotitelia\"
```

V niektorých jazykových nastaveniach Excel používa namiesto bodkočiarky čiarku:

```excel
=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"hodnotitelia\"
```

Tento vzorec sa snaží zistiť priečinok, v ktorom je uložený master súbor, a pripojí k nemu podpriečinok `hodnotitelia\`. Prakticky to znamená: „hľadaj priečinok hodnotitelia vedľa master súboru“.

Dôležité: vzorec `CELL("filename")` funguje spoľahlivo až vtedy, keď je master súbor uložený na disku. Ak ide o nový neuložený súbor, Excel ešte nemusí poznať jeho cestu.

## 7. Ako pomenovať bunku s cestou

Keď máte v konfiguračnom hárku bunku s cestou k priečinku `hodnotitelia/`, treba ju pomenovať. Predpokladajme, že cesta je v bunke `B1`.

Postup:

1. Kliknite na bunku `B1`, v ktorej je cesta k priečinku.
2. Otvorte kartu:

```text
Vzorce
```

3. Vyberte:

```text
Definovať názov
```

4. Ako názov zadajte:

```text
pFolderHodnotitelia
```

5. Potvrďte.

Od tejto chvíle má bunka technický názov `pFolderHodnotitelia`. Tento názov používa Power Query v master súbore. Ak sa cesta zmení, stačí upraviť hodnotu v bunke, nie celý dotaz.

## 8. Ako má Power Query používať konfiguračnú bunku

V Power Query kóde by zdroj nemal byť zapísaný natvrdo takto:

```powerquery
Zdroj = Folder.Files("C:\Users\Admin\Documents\phd prijimacky\hodnotitelia"),
```

Takýto zápis znamená, že Power Query bude vždy hľadať presne tento priečinok. Pri prenose súboru na iný počítač sa to môže pokaziť.

Lepší zápis je tento:

```powerquery
FolderPath = Excel.CurrentWorkbook(){[Name="pFolderHodnotitelia"]}[Content]{0}[Column1],
Zdroj = Folder.Files(FolderPath),
```

Tento zápis znamená: „najprv si z Excelu prečítaj hodnotu pomenovanej bunky `pFolderHodnotitelia` a až potom z tejto cesty načítaj súbory“.

Ak už je Power Query dotaz vytvorený a obsahuje absolútnu cestu, možno ju upraviť cez:

```text
Údaje → Dotazy a pripojenia → dvojklik na dotaz → Domov → Rozšírený editor
```

V rozšírenom editore treba nahradiť pôvodný riadok so zdrojom dvoma riadkami uvedenými vyššie.

## 9. Ako rozpoznať problém s absolútnou cestou

Problém s absolútnou cestou sa typicky prejaví po presune súborov. Napríklad riešenie fungovalo na jednom počítači, ale na druhom počítači master pri obnove údajov hlási chybu alebo nenačíta hodnotiteľské súbory.

Typické príznaky:

1. Excel hlási, že priečinok neexistuje.
2. Power Query hľadá cestu začínajúcu na `C:\Users\...`, ktorá patrí inému počítaču alebo inému používateľovi.
3. Master súbor po presune nenačíta žiadne hodnotiteľské súbory.
4. V rozšírenom editore je vidieť pevne zapísanú cestu, napríklad:

```text
C:\Users\Admin\Documents\...
```

V takom prípade treba buď upraviť cestu priamo v konfiguračnej bunke, alebo upraviť Power Query dotaz tak, aby konfiguračnú bunku používal.

## 10. Čo je excelovská tabuľka a prečo je dôležitá

V tomto riešení nestačí, aby bol zoznam uchádzačov iba obyčajný rozsah buniek. Je potrebné, aby bol zoznam vložený ako **excelovská tabuľka**. V Exceli je tabuľka špeciálny objekt, ktorý má vlastný názov, vlastné stĺpce a vie sa rozširovať alebo spracúvať ako jeden celok.

Rozdiel medzi obyčajným rozsahom a tabuľkou je dôležitý:

```text
Obyčajný rozsah: napríklad A1:C72
Excelovská tabuľka: objekt s názvom, napríklad tblHodnotenie
```

Power Query v master súbore nehľadá náhodné bunky na hárku. Hľadá tabuľku s presným názvom:

```text
tblHodnotenie
```

Ak tabuľka nie je pomenovaná správne, master súbor ju nemusí nájsť. Ak je zoznam uchádzačov iba obyčajný rozsah buniek, Power Query s ním nebude pracovať tak spoľahlivo ako s tabuľkou.

## 11. Ako vytvoriť tabuľku v hodnotiteľskom súbore

Ak pripravujete hodnotiteľský súbor od začiatku, postupujte takto:

1. Do prvého riadku vložte hlavičky stĺpcov:

```text
Uchádzač | Projekt | Portfólio
```

2. Pod hlavičku `Uchádzač` vložte zoznam uchádzačov.
3. Označte celý rozsah vrátane hlavičiek, napríklad:

```text
A1:C72
```

4. Stlačte:

```text
Ctrl + T
```

5. Excel sa opýta, či tabuľka obsahuje hlavičky. Zaškrtnite možnosť:

```text
Tabuľka obsahuje hlavičky
```

6. Potvrďte vytvorenie tabuľky.

Od tejto chvíle už nejde iba o obyčajné bunky. Excel vytvoril tabuľku, ktorú vie Power Query spoľahlivo nájsť a spracovať.

## 12. Ako pomenovať tabuľku `tblHodnotenie`

Toto je jeden z najdôležitejších krokov. Každý hodnotiteľský súbor musí obsahovať tabuľku s rovnakým názvom:

```text
tblHodnotenie
```

Postup v slovenskom Exceli:

1. Kliknite hocikam dovnútra vytvorenej tabuľky.
2. Hore na páse s nástrojmi sa zobrazí kontextová karta, spravidla:

```text
Návrh tabuľky
```

alebo:

```text
Nástroje tabuliek → Návrh tabuľky
```

3. Na ľavej strane tejto karty nájdite pole:

```text
Názov tabuľky
```

4. Pôvodný názov, napríklad `Tabuľka1`, prepíšte na:

```text
tblHodnotenie
```

5. Potvrďte klávesom Enter.

Názov tabuľky nesmie obsahovať medzery. Odporúčame nepoužívať diakritiku. Preto je názov `tblHodnotenie` vhodný: je krátky, jednoznačný a technicky stabilný.

Ak tabuľku nepomenujete správne, master súbor ju pri obnove údajov nemusí načítať. V takom prípade sa typicky objaví chyba v Power Query alebo v masteri nebudú hodnotenia z daného súboru.

## 13. Ako pripraviť hodnotiteľský súbor zo šablóny

Najjednoduchší postup je použiť existujúci súbor `hodnotitel_template.xlsx` a skopírovať ho pre každého hodnotiteľa.

Postup:

1. Skopírujte `hodnotitel_template.xlsx`.
2. Kópiu premenujte napríklad na:

```text
hodnotitel_01.xlsx
```

3. Súbor uložte do priečinka:

```text
hodnotitelia/
```

4. Rovnaký postup zopakujte pre ďalších hodnotiteľov.

Každý hodnotiteľský súbor musí obsahovať tabuľku `tblHodnotenie` a presné názvy stĺpcov:

```text
Uchádzač | Projekt | Portfólio
```

Názvy stĺpcov nemeňte. Ak sa zmení napríklad `Portfólio` na `Portfolio`, `Portfolium` alebo `portfólio`, master súbor už nemusí vedieť dáta správne načítať.

## 14. Ako hodnotiteľ vypĺňa body

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

Hodnotiteľ nemá meniť názvy stĺpcov, názov tabuľky, mená uchádzačov ani štruktúru hárka. Ak sú v šablóne uzamknuté časti hárka, je to zámerné: má sa tým zabrániť náhodnej zmene zoznamu uchádzačov alebo technickej štruktúry súboru.

## 15. Rozbaľovacie menu 0–5 bodov

Rozbaľovacie menu sa v Exceli nastavuje cez funkciu:

```text
Údaje → Overenie údajov
```

Používa sa typ overenia:

```text
Zoznam
```

a ako zdroj hodnôt sa uvedie:

```text
0;1;2;3;4;5
```

V niektorých jazykových nastaveniach Excelu sa namiesto bodkočiarky používa čiarka:

```text
0,1,2,3,4,5
```

Ak už používate pripravenú šablónu, rozbaľovacie menu by malo byť nastavené vopred. Pri zmene počtu uchádzačov však treba skontrolovať, či sa rozbaľovacie menu vzťahuje aj na všetky nové riadky.

## 16. Čo znamená rozsah buniek a kedy ho treba upraviť

V Exceli sa mnohé pravidlá nastavujú na konkrétny rozsah buniek. Napríklad ak máte 71 uchádzačov a hodnotiace stĺpce sú `Projekt` a `Portfólio`, rozsah hodnotiacich buniek môže byť:

```text
B2:C72
```

Tento zápis znamená:

```text
od bunky B2 po bunku C72
```

Teda dva stĺpce a 71 riadkov hodnotení.

Ak sa počet uchádzačov zmení, treba skontrolovať, či rozsah stále sedí. Napríklad:

```text
50 uchádzačov:  B2:C51
71 uchádzačov:  B2:C72
90 uchádzačov:  B2:C91
```

Toto je dôležité najmä pri dvoch veciach:

1. rozbaľovacie menu `0–5`,
2. ochrana hárka, teda ktoré bunky smie hodnotiteľ upravovať.

Samotný master súbor číta tabuľku `tblHodnotenie`, nie pevný rozsah `B2:C72`. To znamená, že ak je zoznam uchádzačov správne vo vnútri tabuľky, Power Query ho vie načítať. Ale rozbaľovacie menu a ochrana hárka môžu byť stále nastavené na konkrétny rozsah. Preto ich treba pri zmene počtu riadkov skontrolovať.

## 17. Ako skontrolovať alebo upraviť rozsah tabuľky

Ak pridáte alebo odstránite uchádzačov, treba overiť, že tabuľka `tblHodnotenie` skutočne obsahuje všetky riadky.

Postup:

1. Kliknite do tabuľky.
2. Otvorte kartu:

```text
Návrh tabuľky
```

3. Skontrolujte, či tabuľka vizuálne obopína všetky riadky uchádzačov.
4. Ak tabuľka nepokrýva všetky riadky, použite malý úchyt v pravom dolnom rohu tabuľky a potiahnite tabuľku tak, aby obsahovala celý zoznam.

Alternatívne možno použiť funkciu:

```text
Zmeniť veľkosť tabuľky
```

Tá sa nachádza na karte `Návrh tabuľky`. Tam možno zadať nový rozsah, napríklad:

```text
=$A$1:$C$91
```

Tento príklad znamená, že tabuľka začína v bunke `A1`, končí v bunke `C91` a obsahuje hlavičky aj všetkých uchádzačov.

## 18. Ako upraviť rozsah rozbaľovacieho menu

Ak pribudnú riadky a nové bunky v stĺpcoch `Projekt` a `Portfólio` nemajú rozbaľovacie menu, treba ho doplniť.

Postup:

1. Označte bunky, kde má byť výber bodov, napríklad:

```text
B2:C91
```

2. Otvorte:

```text
Údaje → Overenie údajov
```

3. Vyberte:

```text
Povoliť: Zoznam
```

4. Ako zdroj zadajte:

```text
0;1;2;3;4;5
```

5. Potvrďte.

Odporúčanie: po úprave rozsahu vyskúšajte jednu bunku v prvom riadku, jednu v strede a jednu v poslednom riadku. V každej by sa malo zobraziť rozbaľovacie menu s hodnotami `0–5`.

## 19. Ako upraviť ochranu hárka

Ak je hárok chránený, hodnotiteľ má upravovať iba bunky s bodmi. Typicky ide o rozsah:

```text
B2:C72
```

Ak sa počet uchádzačov zmení, treba upraviť aj odomknuté bunky. Inak sa môže stať, že nové riadky síce budú v tabuľke, ale hodnotiteľ ich nebude môcť vyplniť.

Základný postup:

1. Najprv vypnite ochranu hárka:

```text
Revízia → Zrušiť ochranu hárka
```

2. Označte bunky, ktoré má hodnotiteľ vypĺňať, napríklad:

```text
B2:C91
```

3. Kliknite pravým tlačidlom a vyberte:

```text
Formát buniek → Ochrana
```

4. Odškrtnite možnosť:

```text
Zamknuté
```

5. Potvrďte.
6. Opäť zapnite ochranu hárka:

```text
Revízia → Chrániť hárok
```

7. Povoľte aspoň možnosť:

```text
Vybrať odomknuté bunky
```

Takto hodnotiteľ bude môcť vypĺňať iba určené bunky s bodmi.

## 20. Ako administrátor obnoví master súbor

Administrátor otvorí:

```text
master_template.xlsx
```

Pred obnovou je vhodné skontrolovať konfiguračný hárok a bunku `pFolderHodnotitelia`. Tá musí ukazovať na priečinok, v ktorom sú uložené hodnotiteľské súbory. Ak je cesta nesprávna, master bude obnovovať nesprávny priečinok alebo zahlási chybu.

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

## 21. Ako pridať nového hodnotiteľa

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

## 22. Ako vymeniť zoznam uchádzačov

Ak sa má použiť nový zoznam uchádzačov, treba ho vložiť do hodnotiteľskej šablóny tak, aby ostala zachovaná tabuľka `tblHodnotenie`.

Odporúčaný postup:

1. Otvorte `hodnotitel_template.xlsx`.
2. V stĺpci `Uchádzač` nahraďte vzorové mená novým zoznamom uchádzačov.
3. Skontrolujte, že nové mená sú stále vo vnútri tabuľky `tblHodnotenie`.
4. Ak sa počet riadkov zmenil, upravte veľkosť tabuľky.
5. Skontrolujte, že stĺpce `Projekt` a `Portfólio` majú rozbaľovacie menu `0–5` vo všetkých riadkoch.
6. Ak je hárok chránený, skontrolujte, že hodnotiteľ môže upravovať všetky potrebné bunky.
7. Súbor uložte ako novú šablónu alebo ho skopírujte pre jednotlivých hodnotiteľov.

Ak je zoznam uchádzačov dlhší alebo kratší než pôvodný, nestačí zmeniť iba mená. Treba skontrolovať aj tabuľku, rozbaľovacie menu a ochranu hárka.

## 23. Dôležité pravidlá

Dodržujte tieto pravidlá:

1. Všetky hodnotiteľské súbory musia byť v priečinku `hodnotitelia/`.
2. Všetky hodnotiteľské súbory musia byť vo formáte `.xlsx`.
3. Každý hodnotiteľský súbor musí obsahovať tabuľku `tblHodnotenie`.
4. Tabuľka musí obsahovať stĺpce `Uchádzač`, `Projekt`, `Portfólio`.
5. Názvy stĺpcov nemeňte.
6. Názvy hodnotiteľských súborov musia byť jednoznačné.
7. Do priečinka `hodnotitelia/` nevkladajte iné pracovné Excel súbory, ktoré nemajú rovnakú štruktúru.
8. Súbory začínajúce `~$` sú dočasné Excel súbory a master ich ignoruje.
9. Pri zmene počtu uchádzačov skontrolujte veľkosť tabuľky, rozbaľovacie menu aj ochranu hárka.
10. Pri prenose balíka medzi počítačmi skontrolujte konfiguračnú bunku `pFolderHodnotitelia`, pretože Power Query môže mať problém s pôvodnou absolútnou cestou.

## 24. Verejný repozitár a ochrana dát

Tento repozitár má obsahovať iba anonymizované alebo vzorové dáta. Do verejného repozitára nepatria:

- reálne mená uchádzačov,
- reálne hodnotenia,
- osobné údaje,
- produkčné hodnotiteľské súbory,
- interné pracovné verzie prijímacieho konania.

Reálne pracovné dáta ukladajte mimo verejný GitHub repozitár, napríklad do interného úložiska organizácie alebo do chráneného SharePoint/OneDrive priestoru.

## 25. Najčastejšie problémy

### Master nenačítal nový hodnotiteľský súbor

Skontrolujte, či je nový súbor uložený v priečinku `hodnotitelia/`, či má príponu `.xlsx` a či obsahuje tabuľku `tblHodnotenie`.

### Master po presune hľadá súbory v nesprávnom priečinku

Toto je pravdepodobne problém absolútnej cesty. Otvorte konfiguračný hárok a skontrolujte bunku `pFolderHodnotitelia`. Musí obsahovať aktuálnu cestu k priečinku `hodnotitelia/`. Ak Power Query stále používa starú cestu, treba skontrolovať Rozšírený editor a overiť, že zdroj používa `FolderPath`, nie pevne zapísanú cestu.

### Master hlási, že nenašiel `tblHodnotenie`

Pravdepodobne tabuľka v hodnotiteľskom súbore nemá správny názov. Otvorte hodnotiteľský súbor, kliknite do tabuľky, prejdite na `Návrh tabuľky` a skontrolujte pole `Názov tabuľky`. Musí tam byť presne:

```text
tblHodnotenie
```

### V masteri sú prázdne hodnoty alebo `null`

To zvyčajne znamená, že hodnotiteľ príslušné bunky nevyplnil alebo súbor neuložil. Otvorte hodnotiteľský súbor, skontrolujte hodnoty a súbor uložte.

### Jeden uchádzač sa zobrazuje viackrát

Môže ísť o rozdiel v zápise mena, napríklad medzera navyše, iná diakritika alebo preklep. Pri produkčnom použití je vhodné pridať aj jednoznačný identifikátor uchádzača, napríklad `ID_uchádzača`.

### Hodnotiteľ nevie vyplniť nové riadky

Pravdepodobne sú nové bunky zamknuté alebo na ne nebolo rozšírené rozbaľovacie menu. Skontrolujte rozsah odomknutých buniek a overenie údajov.

### Nové riadky sa v masteri nezobrazujú

Skontrolujte, či nové riadky patria do tabuľky `tblHodnotenie`. Nestačí, aby boli pod tabuľkou vizuálne dopísané. Musia byť súčasťou samotnej excelovskej tabuľky.

### Master hlási chybu pri obnovovaní

Najčastejšie príčiny sú zmenený názov tabuľky, zmenené názvy stĺpcov, nesprávna cesta k priečinku `hodnotitelia/` alebo vložený súbor s inou štruktúrou.

## 26. Poznámka k jednoznačnej identifikácii uchádzačov

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

## 27. Odporúčaný pracovný postup

Pre bežné použitie odporúčame tento postup:

1. Pripraviť jeden správny hodnotiteľský template.
2. Skontrolovať, že ide o tabuľku `tblHodnotenie`.
3. Skontrolovať, že tabuľka obsahuje všetkých uchádzačov.
4. Skontrolovať rozbaľovacie menu `0–5` v stĺpcoch `Projekt` a `Portfólio`.
5. Skontrolovať ochranu hárka a odomknuté bunky na vypĺňanie.
6. Skontrolovať konfiguračný hárok v master súbore a bunku `pFolderHodnotitelia`.
7. Skopírovať template pre každého hodnotiteľa.
8. Každému hodnotiteľovi prideliť jeho vlastný súbor.
9. Po vyplnení všetky súbory uložiť do priečinka `hodnotitelia/`.
10. Otvoriť `master_template.xlsx`.
11. Použiť `Údaje → Obnoviť všetko`.
12. Skontrolovať výslednú tabuľku.
13. Až z výslednej tabuľky vypočítavať súčty, priemery, poradie alebo rozhodnutie o postupe.

## 28. Krátke zhrnutie pre administrátora

Najdôležitejšie je ustrážiť štyri veci. Po prvé, hodnotiteľské súbory musia byť v správnom priečinku. Po druhé, každá hodnotiteľská tabuľka sa musí volať `tblHodnotenie`. Po tretie, stĺpce musia mať presné názvy `Uchádzač`, `Projekt`, `Portfólio`. Po štvrté, master súbor musí vedieť, kde je priečinok `hodnotitelia/`; preto treba venovať pozornosť konfiguračnému hárku a pomenovanej bunke `pFolderHodnotitelia`.

Ak sa mení počet uchádzačov, treba okrem samotných mien skontrolovať aj veľkosť tabuľky, rozbaľovacie menu a ochranu hárka. Ak sa balík presúva medzi počítačmi, treba skontrolovať cestu k priečinku s hodnotiteľskými súbormi. Pre stabilné pracovné použitie je vhodnejšie držať celé riešenie v cloude, najmä v SharePointe alebo OneDrive, a lokálnu verziu používať najmä na prototypovanie alebo demonštráciu.
