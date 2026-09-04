# GLS Evotec Rapportgenerator

Lille lokalt værktøj til GLS/Evotec, der beriger Evotec / Mapon Excel-eksporter med metadata fra Evotec GPS API.

Programmet ændrer ikke distance- eller timekolonnerne fra Excel-filen. Det tilføjer kun metadata:

- Groups
- Depot
- Client
- Fuel type
- Unit ID

## Sådan bruges programmet

1. Hent eller klon dette repository.
2. Dobbeltklik på `Start GLS Rapport.bat`.
3. Browseren åbner rapportgeneratoren.
4. Vælg Evotec/Mapon `Distance by hours` Excel-rapporten.
5. Indtast Evotec API-nøglen første gang.
6. Klik **Lav rapport**.
7. Klik **Download CSV**.

API-nøglen gemmes kun lokalt i browseren på brugerens computer. Den skal ikke lægges på GitHub.

## Krav

Programmet er lavet til Windows og kræver kun Python. Det kræver ikke `pip`, `openpyxl`, PyInstaller eller andre ekstra Python-pakker.

Hvis Windows ikke kan finde Python, kan det installeres fra Microsoft Store eller python.org.

## Standard API-adresse

Programmet bruger Evotecs GPS-adresse som standard:

```text
https://gps.evotec.dk
```

Teknisk bruges Mapon-kompatible API-endpoints under `/api/v1/...`, men brugerfladen er tilpasset Evotec/GLS.

## Sikkerhed

Læg ikke disse ting i repository:

- API-nøgler
- Kundedata
- Excel-eksporter fra GLS
- Genererede rapporter

## Filer

- `app.py` - selve rapportgeneratoren
- `Start GLS Rapport.bat` - startfil til almindelige Windows-brugere
- `.gitignore` - udelukker logs, cache og kundedata
