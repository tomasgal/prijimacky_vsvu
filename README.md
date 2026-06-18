# Hodnotenie uchádzačov – Excel/Power Query nástroje

Tento repozitár obsahuje pracovné podklady pre zber a spracovanie hodnotení uchádzačov v Exceli pomocou Power Query.

Každý hodnotiteľ vypĺňa samostatný Excel súbor. Všetky hodnotiteľské súbory sú uložené v priečinku `hodnotitelia/` a master súbor ich cez Power Query načíta do spoločnej tabuľky.

## Aktuálna verzia: VIK_2026-27

Aktuálna verzia pracuje so štyrmi hodnotiacimi kritériami:

- `Projekt`
- `Portfólio`
- `Motivačný list`
- `Pohovor`

Každé kritérium sa hodnotí bodmi `0–5`. Pr