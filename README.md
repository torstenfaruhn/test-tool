# DL sporttools

Gecombineerde Flask-webapp voor De Limburger / MHNL met converters voor Cue Print en Cue Web.

## Tools in deze repo

- **Amateurvoetbal: Excel > Cue Print** via `/convert/amateur`
- **Amateurvoetbal: mutaties > Cue Print** via `/convert/amateur-mutaties`
- **Amateurvoetbal: Cue Print > Cue Web** via `/convert/amateur-online`
- **Regiosport: Excel > Cue Print** via `/convert/regiosport`
- **Amateurvoetbal: topscorers > Cue Web** via `/convert/topscorers`
- **Amateurvoetbal: gecumuleerde topscorers > Cue Web** via `/convert/topscorers-cumulated`

## Opbouw van de repo

- `app.py`  
  Flask-app met alle routes, security headers en bestandsdownloads.
- `converter_amateur.py`  
  Converter voor uitslagen Amateurvoetbal van Excel naar Cue Print-tekst.
- `converter_amateur_mutaties.py`  
  Converter voor mutaties Amateurvoetbal van Excel naar Cue Print-tekst.
- `converter_amateur_online.py`  
  Converter van Cue Print-tekst naar Word voor Cue Web.
- `converter_regiosport.py`  
  Converter voor Regiosport van Excel naar Cue Print-tekst.
- `converter_topscorers.py`  
  Converter voor topscorers van tekst naar Word.
- `converter_topscorers_cumulated.py`  
  Converter voor gecumuleerde topscorers van bronbestand + Excel naar Word.
- `templates/index.html`  
  Front-end met cards en formulieren.
- `static/style.css`  
  Styling van de interface.
- `static/app.js`  
  Front-endlogica voor uploaden, converteren en automatische downloads.
- `static/docs/`  
  Voorbeelddocumenten en lege invoerbestanden.
- `notebooks/`  
  Backup van de oorspronkelijke notebooks; deze map wordt niet door de app gebruikt.

## Nieuwe mutatieconverter

De nieuwe mutatieconverter is gebaseerd op het notebook:
`notebooks/mutatieconverter_amateur_colab_v8_sortering_normalisaties.ipynb`

De webapp gebruikt hiervoor:
`converter_amateur_mutaties.py`

De converter:
- leest alleen het eerste werkblad uit een `.xlsx`-bestand;
- gebruikt dezelfde normalisaties als het notebook;
- groepeert op divisie/klasse;
- houdt de sortering uit het notebook aan;
- maakt direct een `.txt`-bestand voor Cue Print;
- slaat geen inhoud online op;
- bewaart het geüploade bestand niet na verwerking.

## Beveiliging en privacy

Deze app:
- gebruikt geen tracking, analytics of third-party scripts;
- zet een Content Security Policy-header;
- verwerkt uploads alleen in het geheugen, behalve de bestaande 2-staps topscorers-tool die tijdelijk in `/tmp` werkt en daarna direct opruimt;
- slaat geen artikelteksten of persoonsdata op in logs;
- gebruikt alleen technische foutmeldingen richting de gebruiker.

## Lokaal draaien

### 1. Virtuele omgeving maken
```bash
python -m venv .venv
```

### 2. Virtuele omgeving activeren
macOS / Linux:
```bash
source .venv/bin/activate
```

Windows PowerShell:
```powershell
.venv\Scripts\Activate.ps1
```

### 3. Dependencies installeren
```bash
pip install -r requirements.txt
```

### 4. App starten
```bash
python app.py
```

De app draait dan op:
`http://localhost:8000`

## Github klaarzetten

### 1. Nieuwe repository maken
Maak in GitHub een nieuwe lege repository aan.

### 2. Bestanden in deze map plaatsen
Zorg dat deze bestanden in de root van de repository staan:
- `app.py`
- `requirements.txt`
- `render.yaml`
- `Procfile`
- converters, `templates/`, `static/`, `notebooks/`

### 3. Git initialiseren
```bash
git init
git add .
git commit -m "Initial commit DL sporttools"
```

### 4. GitHub koppelen
Vervang `JOUW-REPO-NAAM` door de echte naam.
```bash
git branch -M main
git remote add origin git@github.com:JOUW-ACCOUNT/JOUW-REPO-NAAM.git
git push -u origin main
```

## Deployen naar Render

Gebruik voor deze repo een **Web Service**.

### Render-instellingen

- **Service type:** Web Service
- **Environment:** Python
- **Branch:** `main`
- **Root directory:** leeg laten
- **Build command:** `pip install -r requirements.txt`
- **Start command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2`
- **Python version:** `3.12.6`
- **Auto deploy:** aan

### Stappen in Render

1. Log in op Render.
2. Kies **New +**.
3. Kies **Web Service**.
4. Koppel je GitHub-repository.
5. Selecteer deze repository.
6. Controleer of Render de instellingen uit `render.yaml` overneemt.
7. Controleer handmatig:
   - type = Web Service
   - environment = Python
   - build command = `pip install -r requirements.txt`
   - start command = `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2`
8. Start de deploy.
9. Wacht tot de build klaar is.
10. Open daarna de Render-URL en controleer de homepage en de converters.

## Bekende afhankelijkheden

- `Flask`  
  Webframework voor de app.
- `gunicorn`  
  Productieserver voor Render.
- `pandas`  
  Inlezen en verwerken van Exceldata.
- `openpyxl`  
  Lezen en schrijven van `.xlsx`.
- `python-docx`  
  Maken van Word-bestanden.
- `xlrd`  
  Lezen van oudere `.xls`-bestanden voor de gecumuleerde topscorers-tool.

## Controle na deploy

Controleer minimaal deze punten:
- homepage opent zonder mixed content of CSP-fouten;
- de nieuwe card **Amateurvoetbal: mutaties > Cue Print** staat links onderin;
- upload van een geldig `.xlsx`-mutatiebestand geeft direct een `.txt`-download;
- een fout bestandstype geeft een nette foutmelding;
- bestaande cards blijven werken.
