# GLS Evotec Rapportgenerator

Lille lokalt værktøj til GLS/Evotec, der beriger Evotec / Mapon Excel-eksporter med metadata fra Evotec GPS API.

Programmet ændrer ikke distance- eller timekolonnerne fra Excel-filen. Det tilføjer kun metadata:

- Groups
- Depot
- Client
- Fuel type
- Unit ID

## Anbefalet brug: Windows EXE

Til almindelige brugere anbefales den færdige Windows-version:

```text
GLS Rapportgenerator.exe
```

Den kræver ikke Python på computeren.

Sådan hentes EXE-filen:

1. Åbn dette repository på GitHub.
2. Klik **Actions**.
3. Klik **Build Windows EXE**.
4. Åbn den nyeste succesfulde kørsel.
5. Download artifactet **GLS-Rapportgenerator-Windows**.
6. Pak ZIP-filen ud.
7. Dobbeltklik på `GLS Rapportgenerator.exe`.

Første gang kan Windows/Defender vise en advarsel, fordi programmet ikke er signeret. Vælg **Flere oplysninger** og derefter **Kør alligevel**, hvis filen er hentet fra Evotecs eget GitHub-repository.

## Brug af programmet

1. Åbn `GLS Rapportgenerator.exe`.
2. Browseren åbner rapportgeneratoren.
3. Vælg Evotec/Mapon `Distance by hours` Excel-rapporten.
4. Indtast Evotec API-nøglen første gang.
5. Klik **Lav rapport**.
6. Klik **Download CSV**.

API-nøglen gemmes kun lokalt i browseren på brugerens computer. Den skal ikke lægges på GitHub.

## Alternativ: kør fra Python

Hvis man ikke bruger EXE-filen, kan programmet også startes med Python:

1. Hent eller klon dette repository.
2. Dobbeltklik på `Start GLS Rapport.bat`.

Denne metode kræver, at Python er installeret på computeren. Programmet kræver ikke `pip`, `openpyxl`, PyInstaller eller andre ekstra Python-pakker.

## Byg EXE igen

EXE-filen bygges automatisk med GitHub Actions.

Manuel rebuild:

1. Gå til **Actions**.
2. Vælg **Build Windows EXE**.
3. Klik **Run workflow**.
4. Vælg branch `main`.
5. Klik **Run workflow**.
6. Download det nye artifact, når workflowet er færdigt.

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
- `Start GLS Rapport.bat` - startfil til Python-varianten
- `.github/workflows/build-windows-exe.yml` - bygger Windows EXE med GitHub Actions
- `.gitignore` - udelukker logs, cache og kundedata
