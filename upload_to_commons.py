# -*- coding: utf-8 -*-
# BSD 3-Clause License
#
# Copyright (c) 2026, Francesco Tosoni
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# 3. Neither the name of the copyright holder nor the names of its
#    contributors may be used to endorse or promote products derived from
#    this software without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""Carica su Wikimedia Commons le foto elencate in un foglio creato con Pattypan.

Il foglio (.xls) si prepara con Pattypan fermandosi alla creazione del foglio;
le impostazioni personali (nome utente, fotocamera, ...) stanno in config.ini.
La password non viene mai salvata: viene chiesta a ogni avvio.

    python upload_to_commons.py --prova    anteprima, non carica nulla
    python upload_to_commons.py            carica, dopo aver chiesto conferma

Istruzioni complete nel file README.md.
"""
import argparse
import configparser
import getpass
import glob
import os
import shutil
import sys
import time

try:
    import pandas as pd
    import mwclient
except ImportError as e:
    raise SystemExit(
        "Manca una libreria Python (%s).\n"
        "Installa il necessario con:\n"
        "    python -m pip install -r requirements.txt" % e.name)

try:
    from PIL import Image  # only needed for 'fotocamera = auto'
except ImportError:
    Image = None

try:
    sys.stdout.reconfigure(encoding="utf-8")  # so accented characters print on Windows
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

VERSION = '1.0'
SITE_URL = 'commons.wikimedia.org'
PROJECT_URL = 'https://github.com/ftosoni/commons-uploader'
HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(HERE, 'config.ini')
CONFIG_EXAMPLE = os.path.join(HERE, 'config.example.ini')

EXTENSIONS = ('.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff', '.svg', '.webp',
              '.xcf', '.pdf', '.djvu', '.ogg', '.oga', '.ogv', '.opus', '.webm',
              '.wav', '.flac', '.mp3', '.mid', '.stl')
YES = ('s', 'si', 'sì', 'y', 'yes', 'x', '1', 'true', 'vero')
# Characters MediaWiki does not accept in a file title.
FORBIDDEN = '#<>[]|{}:/\\'

# The most common upload warnings, explained to the user.
WARNINGS = {
    'duplicate': "su Commons c'è già un file identico: %s",
    'exists': "su Commons c'è già un file con questo nome",
    'exists-normalized': "su Commons c'è già un file con un nome quasi uguale: %s",
    'page-exists': "su Commons c'è già una pagina con questo nome",
    'was-deleted': "un file con questo nome è stato cancellato in passato",
    'duplicate-archive': "un file identico è stato cancellato in passato: %s",
    'badfilename': "nome non valido; Commons propone: %s",
}


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
def load_config():
    """Read config.ini; on the first run create it from config.example.ini."""
    if not os.path.exists(CONFIG_FILE):
        shutil.copyfile(CONFIG_EXAMPLE, CONFIG_FILE)
        raise SystemExit(
            "Ho creato il file delle impostazioni:\n    %s\n"
            "Aprilo con il Blocco note, compila almeno «utente» e «cartella»,\n"
            "salva e rilancia lo script." % CONFIG_FILE)
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(CONFIG_FILE, encoding='utf-8-sig')
    for section in cfg.sections():
        for key in cfg[section]:
            if 'password' in key or 'token' in key:
                raise SystemExit(
                    "In config.ini c'è una voce «%s»: le password NON vanno scritte\n"
                    "nei file. Cancella quella riga: la password ti verrà chiesta\n"
                    "all'avvio." % key)
    return cfg


def setting(cfg, section, key, default=''):
    return cfg.get(section, key, fallback=default).strip().strip('"').strip()


def find_sheet(cfg, argument):
    """The spreadsheet: command-line argument, else the newest 'pattypan*.xls'
    in the configured folder (backups excluded)."""
    if argument:
        if not os.path.isfile(argument):
            raise SystemExit("Non trovo il foglio: %s" % argument)
        return argument
    folder = setting(cfg, 'foto', 'cartella')
    if not folder:
        raise SystemExit("In config.ini manca «cartella»: la cartella con le foto e\n"
                         "il foglio creato da Pattypan.")
    if not os.path.isdir(folder):
        raise SystemExit("La «cartella» indicata in config.ini non esiste:\n    %s" % folder)
    cands = [f for f in glob.glob(os.path.join(folder, 'pattypan*.xls*')) if '.BACKUP.' not in f]
    if not cands:
        raise SystemExit("Nella cartella non c'è nessun foglio «pattypan ... .xls»:\n    %s\n"
                         "Crealo con Pattypan (vedi README.md)." % folder)
    return max(cands, key=os.path.getmtime)


def read_sheet(path):
    """The 'Data' sheet as a list of dicts; every cell is text, empty if blank."""
    book = pd.read_excel(path, sheet_name=None, dtype=str)
    df = book.get('Data', next(iter(book.values()))).fillna('')
    df.columns = [str(c).strip().lower().replace(' ', '_') for c in df.columns]
    for column in ('path', 'name'):
        if column not in df.columns:
            raise SystemExit("Nel foglio manca la colonna «%s»: è un foglio creato da Pattypan?" % column)
    rows = [dict((k, str(v).strip()) for k, v in r.items()) for r in df.to_dict('records')]
    return [r for r in rows if r['path'] or r['name']]


# ---------------------------------------------------------------------------
# One file
# ---------------------------------------------------------------------------
def resolve_image(path, folder):
    """Locate the local file of a row. Prefer the 'path' as written by Pattypan;
    otherwise a file of the same name inside the sheet's folder, so that the
    whole folder can be moved to another computer."""
    if os.path.isfile(path):
        return path
    candidate = os.path.join(folder, os.path.basename(path.replace('\\', '/')))
    return candidate if os.path.isfile(candidate) else None


def target_name(name, local_path):
    """The file name on Commons. Pattypan's 'name' has no extension: add the one
    of the local file."""
    if name.lower().endswith(EXTENSIONS):
        return name
    return name + os.path.splitext(local_path)[1]


def exif_camera(path):
    """(make, model) from the EXIF data, or None."""
    if Image is None:
        return None
    try:
        with Image.open(path) as im:
            exif = im.getexif()
        make, model = str(exif.get(271) or '').strip(), str(exif.get(272) or '').strip()
    except Exception:
        return None
    return (make, model) if model else None


class Cameras:
    """Resolves the name to put in {{camera|...}}. That template shows
    {{Taken with <name>}}, which exists on Commons only for some cameras: when
    it does not exist the field is left out (Commons shows the EXIF data anyway)."""

    def __init__(self, site, choice):
        self.site = site
        self.choice = choice
        self.known = {}
        self.told = set()

    def exists(self, name):
        if name not in self.known:
            self.known[name] = self.site.pages['Template:Taken with ' + name].exists
        return self.known[name]

    def tell(self, message):
        if message not in self.told:
            self.told.add(message)
            print("  Nota: " + message)

    def name_for(self, path):
        if self.choice.lower() in ('', 'no', 'nessuna'):
            return ''
        if self.choice.lower() != 'auto':
            if self.exists(self.choice):
                return self.choice
            self.tell("su Commons non esiste «Template:Taken with %s»: ometto la fotocamera." % self.choice)
            return ''
        if Image is None:
            self.tell("per «fotocamera = auto» serve Pillow (python -m pip install Pillow): la ometto.")
            return ''
        camera = exif_camera(path)
        if not camera:
            return ''
        make, model = camera
        brand = make.split()[0].capitalize() if make else ''
        for candidate in (model, '%s %s' % (brand, model), '%s %s' % (make, model)):
            candidate = candidate.strip()
            if self.exists(candidate):
                return candidate
        self.tell("su Commons non esiste un modello «Taken with» per %s %s: ometto la fotocamera." % (make, model))
        return ''


def build_wikitext(row, camera, location, personality_rights):
    date = row.get('date', '')
    if date and '{{' not in date:
        if location and '|location=' not in date:
            date += '|location=' + location
        date = '{{taken on|%s}}' % date

    cats = []
    for c in row.get('categories', '').split(';'):
        c = c.strip()
        if c.lower().startswith('category:'):
            c = c[len('category:'):].strip()
        if c:
            cats.append('[[Category:%s]]' % c)
    # Like Pattypan: files without categories are flagged as uncategorised.
    categories = '\n'.join(cats) if cats else '{{subst:unc}}'

    other_fields = '\n|other fields={{camera|%s}}' % camera if camera else ''
    personality = '{{personality rights}}\n\n' if personality_rights else ''

    return """== {{int:filedesc}} ==
{{Information
|description=%s
|date=%s
|source=%s
|author=%s
|permission=%s
|other versions=%s%s
}}

== {{int:license-header}} ==
%s

%s%s
""" % (row.get('description', ''), date, row.get('source', '') or '{{own}}',
       row.get('author', ''), row.get('permission', ''),
       row.get('other_versions', ''), other_fields,
       row.get('license', ''), personality, categories)


# ---------------------------------------------------------------------------
# Checks before uploading
# ---------------------------------------------------------------------------
def prepare(rows, folder):
    """Returns (jobs, errors, notes). A job is (row, local_path, target_name)."""
    jobs, errors, notes, seen = [], [], [], {}
    for n, row in enumerate(rows, start=2):  # row 1 of the sheet is the header
        where = "riga %d" % n
        local = resolve_image(row['path'], folder)
        if not local:
            errors.append("%s: non trovo il file %s" % (where, row['path']))
            continue
        if not row['name']:
            errors.append("%s: la colonna «name» è vuota" % where)
            continue
        name = target_name(row['name'], local)
        bad = sorted(set(ch for ch in name if ch in FORBIDDEN))
        if bad:
            errors.append("%s: il nome contiene caratteri non ammessi (%s): %s" % (where, ' '.join(bad), name))
        if name.lower() in seen:
            errors.append("%s: stesso nome della riga %d: %s" % (where, seen[name.lower()], name))
        seen[name.lower()] = n
        if not row.get('license'):
            errors.append("%s: la colonna «license» è vuota (%s)" % (where, name))
        if not row.get('author'):
            errors.append("%s: la colonna «author» è vuota (%s)" % (where, name))
        if not row.get('description'):
            notes.append("%s: descrizione vuota (%s)" % (where, name))
        if not row.get('categories'):
            notes.append("%s: nessuna categoria, verrà segnalato come non categorizzato (%s)" % (where, name))
        jobs.append((row, local, name))
    return jobs, errors, notes


def wants_personality_rights(row, default):
    """Per row if the sheet has a 'personality_rights' column, else config.ini."""
    if 'personality_rights' in row:
        return row['personality_rights'].lower() in YES
    return default


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Carica su Wikimedia Commons le foto di un foglio creato con Pattypan.")
    parser.add_argument('foglio', nargs='?',
                        help="il foglio .xls (se omesso: il più recente nella «cartella» di config.ini)")
    parser.add_argument('--prova', action='store_true',
                        help="mostra che cosa verrebbe caricato, senza accedere e senza caricare nulla")
    parser.add_argument('--limite', type=int, metavar='N',
                        help="carica solo i primi N file (utile per un primo tentativo)")
    parser.add_argument('--si', action='store_true', help="non chiedere conferma prima di caricare")
    args = parser.parse_args()

    cfg = load_config()
    sheet = find_sheet(cfg, args.foglio)
    folder = os.path.dirname(os.path.abspath(sheet))
    location = setting(cfg, 'foto', 'luogo')
    default_pr = setting(cfg, 'foto', 'personality_rights', 'no').lower() in YES
    pause = float(setting(cfg, 'caricamento', 'pausa', '2') or 2)
    ignore_warnings = setting(cfg, 'caricamento', 'ignora_avvisi', 'no').lower() in YES

    print("Foglio: %s" % sheet)
    jobs, errors, notes = prepare(read_sheet(sheet), folder)
    for note in notes:
        print("  Nota: " + note)
    if errors:
        print("\nCi sono problemi da correggere nel foglio prima di caricare:")
        for error in errors:
            print("  - " + error)
        raise SystemExit(1)
    if not jobs:
        raise SystemExit("Il foglio non contiene righe da caricare.")
    if args.limite:
        jobs = jobs[:args.limite]

    # Environment variables win over config.ini (handy for those who already use them).
    username = os.environ.get('COMMONS_USERNAME', '').strip() or setting(cfg, 'commons', 'utente')
    user_agent = 'commons-uploader/%s (%s; User:%s)' % (
        VERSION, PROJECT_URL, username.split('@')[0] or 'anonymous preview')
    site = mwclient.Site(SITE_URL, clients_useragent=user_agent)
    cameras = Cameras(site, setting(cfg, 'foto', 'fotocamera', 'auto'))

    def wikitext(row, local):
        return build_wikitext(row, cameras.name_for(local), location,
                              wants_personality_rights(row, default_pr))

    if args.prova:
        print("\nPROVA: non carico nulla. File pronti: %d\n" % len(jobs))
        for row, local, name in jobs:
            print("  %s  ->  File:%s" % (os.path.basename(local), name))
        row, local, name = jobs[0]
        page = wikitext(row, local)
        print("\nPagina di descrizione del primo file (File:%s):\n" % name)
        print("-" * 72)
        print(page)
        print("-" * 72)
        print("Se è tutto a posto, rilancia senza --prova.")
        return

    if not username:
        raise SystemExit("In config.ini manca «utente» (es. Mario Rossi@caricamenti).")
    password = os.environ.get('COMMONS_PASSWORD') or getpass.getpass(
        "Bot password di %s (non viene mostrata né salvata): " % username)
    try:
        site.login(username, password)
    except mwclient.errors.LoginError:
        raise SystemExit(
            "Accesso non riuscito. Controlla che «utente» abbia la forma\n"
            "NomeUtente@nomebot e che la password sia quella generata in\n"
            "https://commons.wikimedia.org/wiki/Special:BotPasswords")
    print("Accesso effettuato come %s." % username)

    if not args.si:
        answer = input("Sto per caricare %d file su Wikimedia Commons. Procedo? [s/N] " % len(jobs))
        if answer.strip().lower() not in YES:
            raise SystemExit("Annullato: non ho caricato nulla.")

    done, skipped, failed = 0, 0, []
    for i, (row, local, name) in enumerate(jobs, start=1):
        print("[%d/%d] %s" % (i, len(jobs), name))
        try:
            # Re-running after an interruption only uploads what is missing.
            if site.images[name].exists:
                print("  Già presente su Commons: salto.")
                skipped += 1
                continue
            with open(local, 'rb') as f:
                result = site.upload(f, name, wikitext(row, local), ignore=ignore_warnings,
                                     comment='Caricato con commons-uploader')
            result = result.get('upload', result)  # chunked uploads nest the result
            if result.get('result') == 'Success':
                print("  Fatto.")
                done += 1
            else:
                reasons = []
                for code, detail in (result.get('warnings') or {}).items():
                    text = WARNINGS.get(code, code)
                    if '%s' in text:
                        text = text % (', '.join(detail) if isinstance(detail, list) else detail)
                    reasons.append(text)
                reason = '; '.join(reasons) or str(result)
                print("  NON caricato: " + reason)
                failed.append((name, reason))
            time.sleep(pause)
        except mwclient.errors.FileExists:
            print("  Già presente su Commons: salto.")
            skipped += 1
        except Exception as e:
            print("  Errore: %s" % e)
            failed.append((name, str(e)))
            print("  Attendo 10 secondi prima di proseguire...")
            time.sleep(10)

    print("\nCaricati: %d   Già presenti: %d   Non caricati: %d" % (done, skipped, len(failed)))
    for name, reason in failed:
        print("  - %s: %s" % (name, reason))
    if failed:
        print("Dopo aver sistemato, rilancia lo script: i file già caricati vengono saltati.")


if __name__ == "__main__":
    main()
