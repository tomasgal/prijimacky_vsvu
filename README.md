# Hodnotenie uchádzačov – Excel/Power Query prototyp

Tento repozitár obsahuje anonymizovaný prototyp zberu hodnotení uchádzačov.

## Princíp

Každý hodnotiteľ má vlastný Excel súbor v podadresári `hodnotitelia/`. V hodnotiteľskom súbore vypĺňa iba stĺpce:

- `Projekt`
- `Portfólio`

Hodnoty sa vyberajú z rozbaľovacieho zoznamu `0–5`.

Master súbor následne cez Power Query načíta všetky hodnotiteľské súbory z podadresára `hodnotitelia/` a vytvorí agregovanú tabuľku, kde je každý uchádzač v jednom riadku a hodnotenia od jednotlivých hodnotiteľov sú v samostatných stĺpcoch.

## Súbory

```text
master_template.xlsx
hodnotitel_template.xlsx
hodnotitelia/
  hodnotitel_01_sample.xlsx
  hodnotitel_02_sample.xlsx
docs/
  power_query_m_code.md
```

## Hodnotiteľský súbor

Každý hodnotiteľský súbor musí obsahovať tabuľku s názvom:

```text
tblHodnotenie
```

a presné názvy stĺpcov:

```text
Uchádzač | Projekt | Portfólio
```

Názov súboru sa používa ako identifikátor hodnotiteľa. Napríklad súbor `hodnotitel_03.xlsx` vytvorí v masteri stĺpce:

```text
Projekt_hodnotitel_03
Portfólio_hodnotitel_03
```

## Pridanie nového hodnotiteľa

1. Skopírujte `hodnotitel_template.xlsx`.
2. Premenujte ho napríklad na `hodnotitel_03.xlsx`.
3. Uložte ho do podadresára `hodnotitelia/`.
4. Hodnotiteľ vyplní stĺpce `Projekt` a `Portfólio`.
5. V master súbore použite `Údaje → Obnoviť všetko`.

## Dôležité upozornenie

Do verejného repozitára nedávajte reálne mená uchádzačov, hodnotiteľov ani reálne bodové hodnotenia. Tento balík používa anonymizované vzorové dáta.
