# GLS Evotec Rapportgenerator

Lille lokalt værktøj til GLS/Evotec, der beriger Evotec / Mapon Excel-eksporter med metadata fra Evotec GPS API.

Programmet ændrer ikke distance- eller timekolonnerne fra Excel-filen. Det tilføjer kun metadata:

- Groups
- Depot
- Client
- Fuel type
- Unit ID

## Download for almindelige brugere

Almindelige brugere skal hente den færdige Windows-version under **Releases** ude til højre på siden:

```text
GLS Rapportgenerator.exe
```

Den kræver ikke Python på computeren.

Download-link:

```text
https://github.com/Evotec/GLS-Rapportgenerator/releases
```

Sådan hentes programmet:

1. Åbn linket til **Releases**.
2. Vælg den nyeste release.
3. Download Windows-filen, typisk `GLS-Rapportgenerator-Windows.zip` eller `GLS Rapportgenerator.exe`.
4. Hvis det er en ZIP-fil, så pak den ud først.
5. Dobbeltklik på `GLS Rapportgenerator.exe`.

Første gang kan Windows/Defender vise en advarsel, fordi programmet ikke er signeret. Vælg **Flere oplysninger** og derefter **Kør alligevel**, hvis filen er hentet fra Evotecs eget GitHub-repository.

## Brug af programmet

1. Åbn `GLS Rapportgenerator.exe`.
2. Browseren åbner rapportgeneratoren.
3. Vælg Evotec/Mapon `Distance by hours` Excel-rapporten.
4. Indtast Evotec API-nøglen første gang.
5. Klik **Lav rapport**.
6. Klik **Download CSV**.

API-nøglen gemmes kun lokalt i browseren på brugerens computer. Den skal ikke lægges på GitHub.

## For teknikere: build via GitHub Actions

EXE-filen bygges med GitHub Actions.

Manuel rebuild:

1. Gå til **Actions**.
2. Vælg **Build Windows EXE**.
3. Klik **Run workflow**.
4. Vælg branch `main`.
5. Klik **Run workflow**.
6. Når workflowet er færdigt, download artifactet **GLS-Rapportgenerator-Windows**.
7. Upload den nye ZIP/EXE til **Releases**, så almindelige brugere kan hente den derfra.

Vigtigt: **Actions artifacts er primært til teknikere. Almindelige brugere skal hente programmet fra Releases.**

## Alternativ: kør fra Python

Hvis man ikke bruger EXE-filen, kan programmet også startes med Python:

1. Hent eller klon dette repository.
2. Dobbeltklik på `Start GLS Rapport.bat`.

Denne metode kræver, at Python er installeret på computeren. Programmet kræver ikke `pip`, `openpyxl`, PyInstaller eller andre ekstra Python-pakker.

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
