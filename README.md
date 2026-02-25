
# DL sporttools (Regiosport + Amateurvoetbal)

Gecombineerde Flask‑webapp met tools:
- **DL amateurvoetbal tool** → `/convert/amateur`
- **Amateurvoetbal online (Cue Print → Word)** → `/convert/amateur-online`
- **Amateurvoetbal topscorers (tekst → Word)** → `/convert/topscorers`
- **Amateurvoetbal gecumuleerde topscorers (bron + Excel → Word)** → `/convert/topscorers-cumulated`
- **DL regiosport tool** → `/convert/regiosport`

## Lokaal draaien
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py  # http://localhost:8000
```

## Deploy op Render
1. Push deze map naar GitHub (nieuwe repo).
2. Maak op Render.com een *Web Service* op deze repo.
3. Render gebruikt `render.yaml` (incl. `PYTHON_VERSION=3.12.6`).

## Mappering
- `converter_regiosport.py` – definitieve omzetting voor Regiosport (uit je notebook).
- `converter_amateur.py` – omzetting voor Amateurvoetbal (Excel → Cue Print).
- `converter_amateur_online.py` – Cue Print (txt) → Word (.docx) voor Cue Web.
- `converter_topscorers.py` – Topscorers (txt/docx) → Word (.docx) voor Cue Web.
- `templates/index.html` – gecombineerde UI met meerdere blokken.
- `static/style.css` – stijlen.
- `notebooks/` – referentie-notebooks (niet gebruikt door de app).

## Amateurvoetbal online (Cue Print -> Cue Web)

De converter voor **Amateurvoetbal online** zet een Cue Print `.txt` om naar Cue Web **HTML-code** en levert die bewust uit als **.txt-bestand** (kopieerbaar/plakbaar).

### Definitieve keuze: “Optie 1” (volledige classnamen)
We gebruiken in `converter_amateur_online.py` de **volledige class-attributen** (zoals uit het werkende notebook). Dit is bewust zo gedaan omdat de doelomgeving de markup/styling het meest betrouwbaar herkent wanneer de classnamen exact overeenkomen met de notebook-output.

Een eerdere “Optie 2” (minimalistische markup zonder hashed classnamen) is verwijderd om verwarring te voorkomen.
