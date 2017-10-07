
# Conflict Finder

Inneholder
- Json fil som inneholder data fra excell arket, og følger json standarden.
- Script som går gjenom .json filen og avdekker konflikter på modulnivå eller sub_serie nivå. 

Scriptet utføres slik:
```sh
python conflict_finder.py -i <input-file> -[s/m/b]
s - sub_series level
m - module level
b - both
```

- Når scriptet blir kjørt, blir det opprettet nye filer med relevant informasjon.