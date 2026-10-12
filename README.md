# Commons Uploader

Uno script per caricare tante foto su [Wikimedia Commons](https://commons.wikimedia.org/) in una volta sola, pensato per la comunità di Wikipedia in italiano.

Il foglio di calcolo con nomi, descrizioni e categorie si prepara con [Pattypan](https://commons.wikimedia.org/wiki/Commons:Pattypan/it); il caricamento vero e proprio lo fa questo script, che lascia più libertà di Pattypan e mostra un'anteprima prima di pubblicare qualsiasi cosa.

## ✨ Caratteristiche

- **📋 Parte dal foglio di Pattypan**: nessun formato nuovo da imparare, si compila un normale foglio Excel/LibreOffice.
- **🔎 Anteprima senza rischi**: con `--prova` vedi la pagina di descrizione che verrebbe creata, senza accedere e senza caricare nulla.
- **✅ Controlli prima di partire**: file mancanti, nomi doppi o non validi, licenza o autore vuoti vengono segnalati riga per riga.
- **📷 Fotocamera automatica**: il modello viene letto dai dati EXIF di ogni foto; nessun dato da adattare a mano.
- **🧑 Diritti della personalità**: l'avviso `{{personality rights}}` si attiva per tutte le foto o foto per foto.
- **🔁 Si può rilanciare**: se il caricamento si interrompe, i file già presenti su Commons vengono saltati.
- **🔐 Nessuna password salvata**: la password viene chiesta a ogni avvio e non finisce in nessun file.

## 📂 Struttura del progetto

```
commons-uploader/
├── upload_to_commons.py   # Lo script di caricamento
├── config.example.ini     # Modello delle impostazioni (copiato in config.ini al primo avvio)
├── requirements.txt       # Librerie Python necessarie
├── CITATION.cff           # Come citare il software
├── codemeta.json          # Metadati del software
└── LICENCE.txt            # Licenza BSD a 3 clausole
```

Il file `config.ini` con le tue impostazioni, le foto e i fogli di calcolo restano solo sul tuo computer: Git li ignora.

## 🚀 Guida passo passo

### Cosa serve

- Un'utenza su Wikimedia Commons (è la stessa di Wikipedia).
- [Python](https://www.python.org/downloads/) 3.9 o successivo. Su Windows, durante l'installazione spunta **Add python.exe to PATH**.
- [Pattypan](https://github.com/yarl/pattypan/releases), che richiede Java.

### 1. Crea il foglio con Pattypan, e fermati lì

1. Metti tutte le foto da caricare in una cartella.
2. Apri Pattypan e scegli **Generate spreadsheet**.
3. Seleziona la cartella delle foto.
4. Scegli il modello **Information** e le colonne proposte (descrizione, data, autore, licenza, categorie). La data può essere letta dalle foto.
5. Crea il file: Pattypan salva nella cartella delle foto un foglio chiamato `pattypan AAAA-MM-GG HH_MM_SS.xls`.

> [!IMPORTANT]
> Fermati qui e chiudi Pattypan. **Non** proseguire con *Validate & upload*: da questo punto in poi lavora lo script.

La [guida di Pattypan in italiano](https://commons.wikimedia.org/wiki/Commons:Pattypan/it) spiega ogni schermata.

### 2. Compila il foglio

Apri il foglio con Excel o LibreOffice e compila una riga per foto. Le colonne sono descritte [più sotto](#-colonne-del-foglio); in breve:

- **name**: il titolo del file su Commons. Sceglilo descrittivo (non `DSC01234`); l'estensione viene aggiunta da sola.
- **description**: la descrizione, per esempio `{{it|1=Il Duomo di Pisa visto da est.}}`.
- **categories**: le categorie separate da punto e virgola, senza «Category:», per esempio `Pisa Cathedral;October 2026 in Pisa`.

Salva mantenendo il formato `.xls` e lascia il foglio nella cartella delle foto.

### 3. Scarica lo script e installa le librerie

Scarica questo progetto (pulsante **Code → Download ZIP** qui su GitHub, poi estrai lo ZIP) oppure clonalo:

```bash
git clone https://github.com/ftosoni/commons-uploader.git
```

Apri un terminale nella cartella del progetto (su Windows: PowerShell) e installa le librerie:

```bash
python -m pip install -r requirements.txt
```

### 4. Compila le impostazioni

Lancia lo script una prima volta:

```bash
python upload_to_commons.py --prova
```

Lo script crea il file `config.ini` e si ferma. Aprilo con il Blocco note e compila almeno queste due righe:

```ini
utente = Mario Rossi@caricamenti
cartella = C:\Users\mario\Pictures\gita a Pisa
```

Ogni voce del file è spiegata da un commento; l'elenco completo è [più sotto](#-impostazioni).

### 5. Crea una bot password

Una *bot password* è una password secondaria, limitata ad alcune operazioni: così non dai allo script la tua password principale, e puoi revocarla quando vuoi.

1. Vai su [Special:BotPasswords](https://commons.wikimedia.org/wiki/Special:BotPasswords).
2. Scegli un nome (per esempio `caricamenti`) e concedi i permessi per **modificare e creare pagine** e per **caricare file** (*Edit existing pages*; *Create, edit, and move pages*; *Upload new files*).
3. Copia la password generata e conservala: viene mostrata una volta sola.

Il nome utente da scrivere in `config.ini` ha la forma `TuoNomeUtente@caricamenti`.

### 6. Prova, poi carica

Guarda l'anteprima: l'elenco dei file e la pagina di descrizione del primo.

```bash
python upload_to_commons.py --prova
```

Quando ti convince, carica. Lo script chiede la bot password (mentre scrivi non compare nulla: è normale) e poi una conferma.

```bash
python upload_to_commons.py
```

Per un primo tentativo prudente puoi caricare un solo file e controllarlo su Commons:

```bash
python upload_to_commons.py --limite 1
```

Rilanciando lo script, i file già caricati vengono saltati.

## 🔧 Impostazioni

Tutte le impostazioni stanno in `config.ini`.

| Voce | A cosa serve | Valori |
| --- | --- | --- |
| `utente` | Nome utente della bot password | `NomeUtente@nomebot` |
| `cartella` | Cartella con le foto e il foglio di Pattypan | un percorso |
| `fotocamera` | Fotocamera indicata nella pagina del file | `auto` (dai dati EXIF), `nessuna`, oppure un nome come `Sony ILCE-7M3` |
| `luogo` | Luogo dello scatto, in inglese, aggiunto alla data | per esempio `Italy`; vuoto per ometterlo |
| `personality_rights` | Avviso sui diritti della personalità su tutte le foto | `sì` / `no` |
| `pausa` | Secondi di attesa tra un file e l'altro | un numero, di norma `2` |
| `ignora_avvisi` | Carica anche se Commons segnala un problema (es. file identico già presente) | `sì` / `no` |

La fotocamera compare solo se su Commons esiste il modello corrispondente (`Template:Taken with …`); altrimenti viene omessa senza errori. I dati EXIF restano comunque visibili nella pagina del file.

## 📑 Colonne del foglio

| Colonna | Obbligatoria | Contenuto |
| --- | :---: | --- |
| `path` | sì | Percorso della foto sul computer; lo compila Pattypan |
| `name` | sì | Titolo del file su Commons |
| `description` | consigliata | Descrizione, meglio se in più lingue: `{{it\|1=…}} {{en\|1=…}}` |
| `date` | consigliata | Data dello scatto, per esempio `2026-10-06 12:38` |
| `author` | sì | Autore, per esempio `[[User:Mario Rossi\|Mario Rossi]]` |
| `license` | sì | Licenza, per esempio `{{self\|cc-by-sa-4.0}}` |
| `categories` | consigliata | Categorie separate da `;`, senza «Category:» |
| `source` | no | Fonte; se manca vale `{{own}}` (opera propria) |
| `permission` | no | Eventuale autorizzazione |
| `personality_rights` | no | Colonna da aggiungere a mano: `sì` nelle righe con persone riconoscibili |

Se sposti la cartella su un altro computer non serve correggere la colonna `path`: le foto vengono cercate anche accanto al foglio, per nome.

## ❓ Problemi comuni

- **«Accesso non riuscito»**: controlla che `utente` abbia la forma `NomeUtente@nomebot` e che la password sia quella generata in *Special:BotPasswords*, non quella con cui entri su Wikipedia.
- **«Non caricato: su Commons c'è già un file identico»**: la foto è già stata caricata, magari con un altro nome. Di norma non va ricaricata.
- **«Il nome contiene caratteri non ammessi»**: nei titoli non si possono usare `# < > [ ] | { } : /`.
- **`python` non viene riconosciuto**: reinstalla Python spuntando *Add python.exe to PATH*, oppure prova con `py` al posto di `python`.

## 🔐 Sicurezza

- La password non viene mai scritta su disco. Se in `config.ini` compare una voce `password`, lo script si rifiuta di partire.
- `config.ini`, le foto e i fogli di calcolo sono esclusi da Git tramite `.gitignore`, così non finiscono per sbaglio in un repository pubblico.
- Chi preferisce può passare le credenziali con le variabili d'ambiente `COMMONS_USERNAME` e `COMMONS_PASSWORD`, valide per la sessione del terminale.

> [!NOTE]
> Di ciò che carichi rispondi tu: assicurati di avere i diritti sulle foto e di rispettare le [regole sulle licenze di Commons](https://commons.wikimedia.org/wiki/Commons:Licensing/it).

---

## 🛠️ Tecnologie e stato del progetto

<p align="left">
  <!-- Stato del progetto e licenza -->
  <a href="./LICENCE.txt"><img src="https://img.shields.io/badge/license-BSD_3--Clause-blue?style=flat-square" alt="Licenza"></a>
  <a href="https://archive.softwareheritage.org/browse/origin/?origin_url=https://github.com/ftosoni/commons-uploader"><img src="https://archive.softwareheritage.org/badge/origin/https://github.com/ftosoni/commons-uploader/" alt="SWH Origin"></a>
  <a href="https://archive.softwareheritage.org/swh:1:dir:777ba36942bfbaf7faf70e89a9725ec7363bd6cd;origin=https://github.com/ftosoni/commons-uploader;visit=swh:1:snp:2f7e22c26136270a21282380b1381e2f093a8ad1;anchor=swh:1:rev:dd6bd0eeaf20f77b0ee5e80c070f7081256efc70"><img src="https://archive.softwareheritage.org/badge/swh:1:dir:777ba36942bfbaf7faf70e89a9725ec7363bd6cd/" alt="SWH Directory"></a>
</p>

<p align="left">
  <!-- Tecnologie -->
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.9+-3776ab?logo=python&logoColor=white" alt="Python 3.9+"></a>
  <a href="https://github.com/mwclient/mwclient"><img src="https://img.shields.io/badge/mwclient-MediaWiki_API-green" alt="mwclient"></a>
  <a href="https://commons.wikimedia.org/"><img src="https://img.shields.io/badge/Wikimedia_Commons-upload-3366cc?logo=wikimedia-commons&logoColor=white" alt="Wikimedia Commons"></a>
  <a href="https://commons.wikimedia.org/wiki/Commons:Pattypan/it"><img src="https://img.shields.io/badge/Pattypan-foglio_di_calcolo-orange" alt="Pattypan"></a>
</p>

Lo sviluppo di questo software segue il [Codice di condotta](./CODE_OF_CONDUCT.md) dei progetti tecnici Wikimedia.

## 📄 Licenza

[Licenza BSD a 3 clausole](./LICENCE.txt). Per citare il software vedi [CITATION.cff](./CITATION.cff).
