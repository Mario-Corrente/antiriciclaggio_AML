# Antiriciclaggio AML - Software di Profilatura Rischio

Software per la valutazione del rischio antiriciclaggio conforme alle Linee Guida VEDA 02-2020 e D.Lgs. 231/2007.

## Caratteristiche

- ✅ Valutazione automatica del rischio cliente e operazione
- ✅ Riconoscimento intelligente della natura giuridica
- ✅ Export in formato Word professionale
- ✅ Database configurabile (clienti, avvocati, luoghi a rischio)
- ✅ Calcolo automatico del rischio inerente e specifico
- ✅ Interfaccia grafica intuitiva

## Requisiti

- Python 3.8 o superiore (per sviluppo)
- Windows 10/11 (per eseguibile standalone)

## Installazione (Sviluppo)

```bash
# Clona il repository
git clone <repository-url>
cd antiriciclaggio_AML

# Installa le dipendenze
pip install -r requirements.txt

# Avvia l'applicazione
python antiriciclaggio.py
```

## Download Eseguibile Windows

Scarica l'ultima versione dell'eseguibile dalle [Release](../../releases) o dagli [Artifacts](../../actions) delle GitHub Actions.

## Struttura del Progetto

```
antiriciclaggio_AML/
├── antiriciclaggio.py          # File principale
├── icona.ico                   # Icona applicazione
├── requirements.txt            # Dipendenze Python
└── config/                     # File di configurazione JSON
    ├── avvocati.json
    ├── clienti_studio.json
    ├── configurazione.json
    ├── fattori_rischio.json
    ├── luoghi_rischio.json
    ├── natura_giuridica.json
    └── prestazioni_veda.json
```

## Build Manuale EXE (Windows)

```bash
pip install pyinstaller
pyinstaller --name="Antiriciclaggio_AML" --onefile --windowed --icon=icona.ico --add-data "config;config" --add-data "icona.ico;." antiriciclaggio.py
```

L'eseguibile sarà generato in `dist/Antiriciclaggio_AML.exe`.

## GitHub Actions

Il progetto include una GitHub Action che:
- Compila automaticamente l'EXE ad ogni push su `main`/`master`
- Carica l'EXE come artifact (disponibile per 30 giorni)
- Crea automaticamente una Release quando viene creato un tag `v*`

### Creare una Release

```bash
git tag -a v1.0.0 -m "Release version 1.0.0"
git push origin v1.0.0
```

## Licenza

Uso interno - Tutti i diritti riservati

## Autore

Sviluppato per Studio Legale
