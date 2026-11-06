import pandas as pd
import pypdf
import io
import os
import warnings
import datetime
import calendar
import unicodedata
import subprocess
import shutil
import math
import sys
import json
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------------------- storage
# elstat_config.json (next to the program) only remembers WHERE the data folder is.
# Everything else (hotel details, history, archive, preferences) lives in the data folder.
CONFIG_FILE = os.path.join(BASE_DIR, "elstat_config.json")
DEFAULT_DATA_DIR = os.path.join(BASE_DIR, "ELSTAT_Data")
DEFAULT_TEMPLATE = os.path.join(BASE_DIR, "template_elstat_report.pdf")
ARCHIVE_NAME = "ELSTAT_Archive"
DATA_FILES = ("hotel_settings.csv", "elstat_history.csv", "elstat_months.csv", "elstat_preferences.json")

DATA_DIR = SETTINGS_CSV = HISTORY_CSV = MONTHS_CSV = PREFS_JSON = ARCHIVE_DIR = ""


def set_data_dir(path):
    """Points every data path of the program at `path`."""
    global DATA_DIR, SETTINGS_CSV, HISTORY_CSV, MONTHS_CSV, PREFS_JSON, ARCHIVE_DIR
    DATA_DIR = os.path.abspath(path)
    SETTINGS_CSV = os.path.join(DATA_DIR, "hotel_settings.csv")
    HISTORY_CSV = os.path.join(DATA_DIR, "elstat_history.csv")
    MONTHS_CSV = os.path.join(DATA_DIR, "elstat_months.csv")
    PREFS_JSON = os.path.join(DATA_DIR, "elstat_preferences.json")
    ARCHIVE_DIR = os.path.join(DATA_DIR, ARCHIVE_NAME)


set_data_dir(DEFAULT_DATA_DIR)


def ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def _same_path(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _is_inside(child, parent):
    c, p = os.path.normcase(os.path.abspath(child)), os.path.normcase(os.path.abspath(parent))
    return c == p or c.startswith(p + os.sep)


def data_exists(folder):
    """True if `folder` already holds data of this program."""
    return (any(os.path.exists(os.path.join(folder, f)) for f in DATA_FILES) or
            os.path.isdir(os.path.join(folder, ARCHIVE_NAME)))


def copy_data_files(src_dir, dst_dir):
    """Copies the program's data (not deleting anything) from one folder to another."""
    for name in DATA_FILES:
        p = os.path.join(src_dir, name)
        if os.path.isfile(p):
            shutil.copy2(p, os.path.join(dst_dir, name))
    a = os.path.join(src_dir, ARCHIVE_NAME)
    if os.path.isdir(a):
        shutil.copytree(a, os.path.join(dst_dir, ARCHIVE_NAME), dirs_exist_ok=True)


def load_app_config():
    try:
        with open(CONFIG_FILE, encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_app_config(cfg):
    """Returns False if the program folder is read-only."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def init_storage():
    """Run once at start-up: finds/creates the data folder and, the first time after an update,
    copies the old files that used to sit next to the program. Returns messages to show the user."""
    notices = []
    wanted = str(load_app_config().get('data_dir') or '').strip()
    path = wanted or DEFAULT_DATA_DIR
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        if wanted:
            notices.append(f"Ο φάκελος δεδομένων\n{wanted}\nδεν είναι διαθέσιμος αυτή τη στιγμή (π.χ. "
                           "αποσυνδεδεμένος δίσκος). Χρησιμοποιείται προσωρινά ο προεπιλεγμένος φάκελος.")
        path = DEFAULT_DATA_DIR
        try:
            os.makedirs(path, exist_ok=True)
        except OSError:
            path = os.path.join(os.path.expanduser('~'), 'ELSTAT_Data')
            os.makedirs(path, exist_ok=True)
    set_data_dir(path)
    if _same_path(path, DEFAULT_DATA_DIR) and not data_exists(path) and data_exists(BASE_DIR):
        try:
            copy_data_files(BASE_DIR, path)
            notices.append("Τα δεδομένα του προγράμματος (στοιχεία ξενοδοχείου, ιστορικό, αρχείο PDF) "
                           f"αντιγράφηκαν στον φάκελο:\n{path}\n\nΤα παλιά αρχεία έμειναν στη θέση τους ως "
                           "αντίγραφο ασφαλείας - μπορείτε να τα σβήσετε όταν βεβαιωθείτε ότι όλα λειτουργούν.\n"
                           "Τον φάκελο δεδομένων τον αλλάζετε από την καρτέλα «Ρυθμίσεις».")
        except Exception:
            pass
    return notices


# ------------------------------------------------------------- preferences
BOOL_PREFS = ('open_after', 'confirm_overwrite', 'keep_archive', 'show_reminder', 'prefill_employees')
DEFAULT_PREFS = {
    'output_mode': 'program',                  # 'program' | 'excel' | 'custom'
    'output_dir': '',                          # used when output_mode == 'custom'
    'filename_pattern': 'ELSTAT_{year}_{month}',
    'excel_dir': '',                           # '' = remember the last folder
    'template_pdf': '',                        # '' = template_elstat_report.pdf next to the program
    'open_after': False,
    'confirm_overwrite': True,
    'keep_archive': True,
    'show_reminder': True,
    'prefill_employees': False,
}


def pattern_error(pattern):
    """None if the file-name pattern is fine, else a Greek explanation."""
    if not pattern.strip():
        return "Το όνομα δεν μπορεί να είναι κενό."
    try:
        pattern.format(year=2026, month="01")
    except Exception:
        return "Μη έγκυρο όνομα. Χρησιμοποιήστε μόνο τα {year} και {month} μέσα σε άγκιστρα."
    if '{year}' not in pattern or '{month}' not in pattern:
        return ("Πρέπει να περιέχει και το {year} και το {month}, αλλιώς κάθε μήνας θα γράφει "
                "πάνω στον προηγούμενο.")
    return None


def build_pdf_name(pattern, year, month):
    """File name (with .pdf) for a month; characters Windows does not allow become '_'."""
    mm = f"{int(month):02d}"
    try:
        name = pattern.format(year=year, month=mm)
    except (KeyError, IndexError, ValueError, AttributeError):
        name = DEFAULT_PREFS['filename_pattern'].format(year=year, month=mm)
    name = ''.join('_' if (ch in '<>:"/\\|?*' or ord(ch) < 32) else ch for ch in name).strip().rstrip('.')
    if name.lower().endswith('.pdf'):
        name = name[:-4]
    return (name or f"ELSTAT_{year}_{mm}") + ".pdf"


def load_prefs():
    prefs = dict(DEFAULT_PREFS)
    try:
        with open(PREFS_JSON, encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):
            for k, default in DEFAULT_PREFS.items():
                if k in data:
                    prefs[k] = bool(data[k]) if isinstance(default, bool) else str(data[k] or '')
    except Exception:
        pass
    if prefs['output_mode'] not in ('program', 'excel', 'custom'):
        prefs['output_mode'] = 'program'
    if pattern_error(prefs['filename_pattern']):
        prefs['filename_pattern'] = DEFAULT_PREFS['filename_pattern']
    return prefs


def save_prefs(prefs):
    ensure_data_dir()
    with open(PREFS_JSON, 'w', encoding='utf-8') as f:
        json.dump(prefs, f, ensure_ascii=False, indent=2)

_LATIN_TO_GREEK = str.maketrans({
    'A': 'Α', 'B': 'Β', 'E': 'Ε', 'H': 'Η', 'I': 'Ι', 'K': 'Κ', 'M': 'Μ',
    'N': 'Ν', 'O': 'Ο', 'P': 'Ρ', 'T': 'Τ', 'Y': 'Υ', 'X': 'Χ', 'Z': 'Ζ'})

def norm(s):
    """Upper-case, strip all whitespace and turn Latin look-alikes into Greek,
    so Excel names and template text compare equal however they are spaced/typed."""
    s = unicodedata.normalize('NFD', str(s).upper())
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')   # drop accents / diaeresis
    return ''.join(s.split()).translate(_LATIN_TO_GREEK)


SETTING_KEYS = ('type', 'name', 'category', 'email', 'phone', 'afm', 'company')

def load_hotel_settings():
    """Hotel details come ONLY from hotel_settings.csv. If the file is missing, all fields are empty."""
    settings = {k: '' for k in SETTING_KEYS}
    settings['last_dir'] = ''
    settings['last_employees'] = ''
    if os.path.exists(SETTINGS_CSV):
        try:
            df = pd.read_csv(SETTINGS_CSV, encoding='utf-8-sig', dtype=str, keep_default_na=False)
            if not df.empty:
                saved = df.iloc[0].to_dict()
                for k in settings:
                    settings[k] = str(saved.get(k, '') or '').strip()
        except Exception:
            pass
    return settings

def save_hotel_settings(settings_data):
    ensure_data_dir()
    df = pd.DataFrame([settings_data])
    df.to_csv(SETTINGS_CSV, index=False, encoding='utf-8-sig')

def _to_date(v):
    """Excel cell -> datetime.date (real date cells or text such as 10/31/2026), else None."""
    if isinstance(v, (datetime.datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%d.%m.%Y"):
            try:
                return datetime.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None

def read_excel_period(excel_path):
    """Reads the report period (Από / Έως Ημερομηνία) printed in the Protel header.
    Returns (date_from, date_to) or None if it cannot be found."""
    try:
        df = pd.read_excel(excel_path, header=None, nrows=15)
    except Exception:
        return None
    # the label cells are Greek text saved as cp1253 -> repair like in the main reader
    def fix(t):
        if isinstance(t, str):
            try:
                return t.encode('iso-8859-1').decode('cp1253')
            except Exception:
                return t
        return t
    found = {}
    for r in range(len(df)):
        for c in range(df.shape[1]):
            label = norm(fix(df.iat[r, c])) if isinstance(df.iat[r, c], str) else ''
            for key, tag in (('ΑΠΟΗΜΕΡΟΜΗΝΙΑ', 'from'), ('ΕΩΣΗΜΕΡΟΜΗΝΙΑ', 'to')):
                if label.startswith(key) and tag not in found:
                    for c2 in range(c + 1, df.shape[1]):     # first date to the right of the label
                        d = _to_date(df.iat[r, c2])
                        if d:
                            found[tag] = d
                            break
    if 'from' in found and 'to' in found and found['from'] <= found['to']:
        return found['from'], found['to']
    return None

def period_as_month(period):
    """(year, month, last_day) if the period is exactly one calendar month, else None."""
    if not period:
        return None
    d1, d2 = period
    last = calendar.monthrange(d1.year, d1.month)[1]
    if d1.year == d2.year and d1.month == d2.month and d1.day == 1 and d2.day == last:
        return d1.year, d1.month, last
    return None

def open_file(path):
    try:
        if sys.platform.startswith('win'):
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])
    except Exception:
        pass

def print_file(path):
    """Sends the PDF straight to the default printer. Returns (ok, message).
    If silent printing is not possible on this PC the PDF is opened instead (then press Ctrl+P)."""
    try:
        if sys.platform.startswith('win'):
            os.startfile(path, "print")          # uses the PDF program's own "print" action
        else:
            cmd = shutil.which('lp') or shutil.which('lpr')
            if not cmd:
                raise OSError("no print command")
            subprocess.run([cmd, path], check=True, timeout=30,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True, "✓ Το δελτίο στάλθηκε στον προεπιλεγμένο εκτυπωτή."
    except Exception:
        open_file(path)
        return False, ("Δεν ήταν δυνατή η αυτόματη εκτύπωση. Άνοιξα το PDF - πατήστε Ctrl+P "
                       "μέσα στο πρόγραμμα του PDF για να το εκτυπώσετε.")

def validate_afm(afm):
    """AFM must be exactly 9 digits (0-9)."""
    afm = str(afm).strip()
    return len(afm) == 9 and afm.isascii() and afm.isdigit()

def validate_hotel(settings):
    errors = []
    if not validate_afm(settings.get('afm', '')):
        errors.append("Το Α.Φ.Μ. πρέπει να έχει ακριβώς 9 ψηφία (μόνο αριθμοί).")
    for key, label in [('type', 'Είδος Καταλύματος'), ('name', 'Τίτλος Καταλύματος'),
                       ('category', 'Κατηγορία'), ('company', 'Επωνυμία Επιχείρησης')]:
        if not settings.get(key, '').strip():
            errors.append(f"Συμπληρώστε το πεδίο «{label}» (στην καρτέλα «Στοιχεία ξενοδοχείου»).")
    return errors

def validate_month(monthly):
    errors = []
    year, month = monthly.get('year', ''), monthly.get('month', '')
    y_ok = year.isascii() and year.isdigit() and len(year) == 4
    m_ok = month.isascii() and month.isdigit() and 1 <= int(month) <= 12
    if not y_ok:
        errors.append("Το Έτος πρέπει να έχει 4 ψηφία (π.χ. 2026).")
    if not m_ok:
        errors.append("Ο Μήνας πρέπει να είναι αριθμός από 1 έως 12.")

    try:
        d_to = datetime.datetime.strptime(monthly.get('date_to', ''), "%d/%m/%Y")
        if y_ok and m_ok and (d_to.year != int(year) or d_to.month != int(month)):
            errors.append("Η ημερομηνία «Έως» δεν ταιριάζει με το Έτος/Μήνα.")
    except ValueError:
        errors.append("Η ημερομηνία «Έως» πρέπει να είναι της μορφής ΗΗ/ΜΜ/ΕΕΕΕ και υπαρκτή.")

    days = monthly.get('days_of_month', '')
    if not (days.isascii() and days.isdigit() and 1 <= int(days) <= 31):
        errors.append("Οι Ημέρες Λειτουργίας πρέπει να είναι αριθμός από 1 έως 31.")
    elif y_ok and m_ok and int(days) > calendar.monthrange(int(year), int(month))[1]:
        errors.append("Οι Ημέρες Λειτουργίας είναι περισσότερες από τις ημέρες του μήνα.")

    emp = monthly.get('employees', '')
    if not emp:
        errors.append("Πληκτρολογήστε το Σύνολο Εργαζομένων (δεν υπάρχει στο Excel).")
    elif not (emp.isascii() and emp.isdigit() and len(emp) <= 3):
        errors.append("Το Σύνολο Εργαζομένων πρέπει να είναι αριθμός (0-999).")
    return errors

def validate_inputs(settings, monthly, excel, template, out):
    errors = validate_hotel(settings) + validate_month(monthly)
    if not excel or not os.path.isfile(excel):
        errors.append("Επιλέξτε υπαρκτό αρχείο Excel.")
    if not template or not os.path.isfile(template):
        errors.append("Το άδειο PDF (template) δεν βρέθηκε.")
    if not out:
        errors.append("Επιλέξτε πού θα αποθηκευτεί το PDF.")
    return errors

COUNTRY_MAPPING = {
    'ΕΛΛΑΔΑ': 'ΕΛΛΑΔΑ', 'AUSTRIA': 'ΑΥΣΤΡΙΑ', 'BELGIUM': 'ΒΕΛΓΙΟ', 
    'BULGARIA': 'ΒΟΥΛΓΑΡΙΑ', 'FRANCE': 'ΓΑΛΛΙΑ', 'GERMANY': 'ΓΕΡΜΑΝΙΑ', 
    'DENMARK': 'ΔΑΝΙΑ', 'ESTHONIA': 'ΕΣΘΟΝΙΑ', 'GREAT BRITAIN': 'ΗΝΩΜΕΝΟ ΒΑΣΙΛΕΙΟ', 
    'IRELAND': 'ΙΡΛΑΝΔΙΑ', 'SPAIN': 'ΙΣΠΑΝΙΑ', 'ITALY': 'ΙΤΑΛΙΑ', 
    'CROATIA': 'ΚΡΟΑΤΙΑ', 'CYPRUS': 'ΚΥΠΡΟΣ', 'LATVIA': 'ΛΕΤΟΝΙΑ', 
    'LITHUANIA': 'ΛΙΘΟΥΑΝΙΑ', 'LUXENBURG': 'ΛΟΥΞΕΜΒΟΥΡΓΟ', 'MALTA': 'ΜΑΛΤΑ', 
    'NETHERLANDS': 'ΟΛΛΑΝΔΙΑ', 'HUNGARY': 'ΟΥΓΓΑΡΙΑ', 'POLAND': 'ΠΟΛΩΝΙΑ', 
    'PORTUGAL': 'ΠΟΡΤΟΓΑΛΙΑ', 'ROMANIA': 'ΡΟΥΜΑΝΙΑ', 'SLOVAKIA': 'ΣΛΟΒΑΚΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 
    'SLOVENIA': 'ΣΛΟΒΕΝΙΑ', 'SWEDEN': 'ΣΟΥΗΔΙΑ', 'CZECH REPUBLIC': 'ΤΣΕΧΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 
    'FINLAND': 'ΦΙΝΛΑΝΔΙΑ', 'SWITZERLAND': 'ΕΛΒΕΤΙΑ', 'NORWAY': 'ΝΟΡΒΗΓΙΑ', 
    'ALBANIA': 'ΑΛΒΑΝΙΑ', 'GEORGIA': 'ΓΕΩΡΓΙΑ', 'SERBIA': 'ΣΕΡΒΙΑ', 
    'UKRAINE': 'ΟΥΚΡΑΝΙΑ', 'F.Y.R.O.M.': 'ΒΟΡΕΙΑΣ ΜΑΚΕΔΟΝΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 
    'RUSIA': 'ΡΩΣΙΑ', 'TURKEY': 'ΤΟΥΡΚΙΑ', 'U.S.A.': 'Η.Π.Α.', 
    'CANADA': 'ΚΑΝΑΔΑΣ', 'BRAZIL': 'ΒΡΑΖΙΛΙΑ', 'INDIA': 'ΙΝΔΙΑ', 
    'ISRAEL': 'ΙΣΡΑΗΛ', 'CHINA': 'ΚΙΝΑΣ ΛΑΪΚΗ ΔΗΜΟΚΡΑΤΙΑ', 
    'AUSTRALIA': 'ΑΥΣΤΡΑΛΙΑ', 'NEW ZEALAND': 'ΝΕΑ ΖΗΛΑΝΔΙΑ ΚΑΙ ΛΟΙΠΕΣ ΠΕΡΙΟΧΕΣ',
    'PAKISTAN': 'ΠΑΚΙΣΤΑΝ', 'QATAR': 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΣΙΑΣ', 'THAILAND': 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΣΙΑΣ',
    'SAUDI ARABIA': 'ΣΑΟΥΔΙΚΗ ΑΡΑΒΙΑ',
    'UNITED ARAB EMIRATES': 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΣΙΑΣ',
    'SINGAPOURE': 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΣΙΑΣ',
    'No Country': 'Δεν δήλωσαν χώρα'
}

# extra spellings / countries that exist in the form but were missing from the Protel list
COUNTRY_MAPPING.update({
    'UNITED KINGDOM': 'ΗΝΩΜΕΝΟ ΒΑΣΙΛΕΙΟ', 'UK': 'ΗΝΩΜΕΝΟ ΒΑΣΙΛΕΙΟ',
    'LUXEMBOURG': 'ΛΟΥΞΕΜΒΟΥΡΓΟ', 'ESTONIA': 'ΕΣΘΟΝΙΑ', 'RUSSIA': 'ΡΩΣΙΑ',
    'CZECHIA': 'ΤΣΕΧΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 'NORTH MACEDONIA': 'ΒΟΡΕΙΑΣ ΜΑΚΕΔΟΝΙΑΣ ΔΗΜΟΚΡΑΤΙΑ',
    'ICELAND': 'ΙΣΛΑΝΔΙΑ', 'LIECHTENSTEIN': 'ΛΙΧΤΕΝΣΤΑΪΝ', 'MONTENEGRO': 'ΜΑΥΡΟΒΟΥΝΙΟ',
    'MOLDOVA': 'ΜΟΛΔΑΒΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 'ARGENTINA': 'ΑΡΓΕΝΤΙΝΗ', 'MEXICO': 'ΜΕΞΙΚΟ',
    'USA': 'Η.Π.Α.', 'UNITED STATES': 'Η.Π.Α.',
    'JAPAN': 'ΙΑΠΩΝΙΑ', 'IRAQ': 'ΙΡΑΚ', 'IRAN': 'ΙΡΑΝ', 'SOUTH KOREA': 'ΚΟΡΕΑΣ ΔΗΜΟΚΡΑΤΙΑ',
    'KOREA': 'ΚΟΡΕΑΣ ΔΗΜΟΚΡΑΤΙΑ', 'LEBANON': 'ΛΙΒΑΝΟΣ', 'SYRIA': 'ΣΥΡΙΑ', 'PHILIPPINES': 'ΦΙΛΙΠΠΙΝΕΣ',
    'EGYPT': 'ΑΙΓΥΠΤΟΣ', 'ETHIOPIA': 'ΑΙΘΙΟΠΙΑ', 'ALGERIA': 'ΑΛΓΕΡΙΑ', 'NIGERIA': 'ΝΙΓΗΡΙΑ',
    'SOUTH AFRICA': 'ΝΟΤΙΑ ΑΦΡΙΚΗ',
})
COUNTRY_MAPPING_NORM = {norm(k): norm(v) for k, v in COUNTRY_MAPPING.items()}


def fix_greek(text):
    """Protel saves Greek text as cp1253 read as latin-1; this repairs it."""
    if isinstance(text, str):
        try:
            return text.encode('iso-8859-1').decode('cp1253')
        except Exception:
            return text
    return text


def read_nationality_rows(excel_path):
    """Reads the Protel 'Nationalities Statistics' sheet.
    Returns ([(country_as_in_excel, clients, nights), ...], excel_grand_total_or_None).
    Used both for the live preview and for building the PDF, so they always agree."""
    df = pd.read_excel(excel_path, header=None)
    df = df.map(fix_greek) if hasattr(df, 'map') else df.applymap(fix_greek)
    rows, excel_total = [], None
    for idx in range(len(df)):
        row = df.iloc[idx]
        if pd.isna(row[0]):
            continue
        country_excel = str(row[0]).strip()
        key = norm(country_excel)
        if not key:
            continue
        try:
            pax = int(row[7])
            nights = int(row[9])
        except (ValueError, TypeError):
            continue
        if key.startswith('ΣΥΝΟΛ') or key.startswith('ΤΟΤΑΛ'):   # Protel writes "Σύνολα :"
            excel_total = (pax, nights)      # Protel's own grand total, used as a check
            continue
        if pax <= 0 and nights <= 0:
            continue
        rows.append((country_excel, pax, nights))
    return rows, excel_total


def load_history():
    """{year: {month: [(country, clients, nights), ...]}} from elstat_history.csv."""
    if not os.path.exists(HISTORY_CSV):
        return {}
    try:
        df = pd.read_csv(HISTORY_CSV, encoding='utf-8-sig', dtype=str, keep_default_na=False)
    except Exception:
        return {}
    out = {}
    for _, r in df.iterrows():
        try:
            y, m = int(r['Etos']), int(r['Minas'])
            c, n = int(float(r['Pelates'])), int(float(r['Dianykterefseis']))
        except (ValueError, KeyError):
            continue
        out.setdefault(y, {}).setdefault(m, []).append((pretty_country(str(r['Chwra'])), c, n))
    return out


def find_recent_excel(folder):
    """Newest .xls/.xlsx file in `folder` (ignores Excel's ~$ lock files), or None."""
    try:
        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('~$')]
    except OSError:
        return None
    return max(files, key=os.path.getmtime) if files else None


MAPPED_TARGETS = set(COUNTRY_MAPPING_NORM.values())
PRETTY_COUNTRY = {norm(v): v for v in COUNTRY_MAPPING.values()}


def pretty_country(name):
    """History stores names without spaces (ΗΝΩΜΕΝΟΒΑΣΙΛΕΙΟ); this gives back the readable form."""
    return PRETTY_COUNTRY.get(norm(name), name)


def resolve_country(name):
    """Row of the form a country belongs to. Accepts the Protel (English) name or the form's Greek name."""
    k = norm(name)
    return COUNTRY_MAPPING_NORM.get(k) or (k if k in MAPPED_TARGETS else None)


# (ARCHIVE_DIR / MONTHS_CSV are set by set_data_dir() at the top of the file)


def archive_path(year, month):
    return os.path.join(ARCHIVE_DIR, f"ELSTAT_{year}_{month:02d}.pdf")


def archive_copy(year, month, src):
    """Keeps a copy of every generated form in the archive folder (so old months can be reprinted)."""
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    dst = archive_path(year, month)
    if os.path.abspath(src) != os.path.abspath(dst):
        shutil.copy2(src, dst)
    return dst


def load_month_records():
    """{(year, month): {'employees':.., 'days':.., 'date_to':..}} - the typed-in values of each month."""
    if not os.path.exists(MONTHS_CSV):
        return {}
    try:
        df = pd.read_csv(MONTHS_CSV, encoding='utf-8-sig', dtype=str, keep_default_na=False)
        return {(int(r['Etos']), int(r['Minas'])): {'employees': r['Ypalliloi'], 'days': r['Meres'],
                                                    'date_to': r['DateTo']} for _, r in df.iterrows()}
    except Exception:
        return {}


def save_month_record(year, month, employees, days, date_to):
    rec = {'Etos': str(year), 'Minas': str(month), 'Ypalliloi': str(employees), 'Meres': str(days),
           'DateTo': str(date_to), 'Created': datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}
    df = pd.DataFrame([rec])
    if os.path.exists(MONTHS_CSV):
        try:
            old = pd.read_csv(MONTHS_CSV, encoding='utf-8-sig', dtype=str, keep_default_na=False)
            keep = ~((pd.to_numeric(old['Etos'], errors='coerce') == year) &
                     (pd.to_numeric(old['Minas'], errors='coerce') == month))
            df = pd.concat([old[keep], df], ignore_index=True)
        except Exception:
            pass
    ensure_data_dir()
    df.to_csv(MONTHS_CSV, index=False, encoding='utf-8-sig')


def delete_month(year, month):
    """Removes one month from the history, the month records and the archive."""
    for path in (HISTORY_CSV, MONTHS_CSV):
        if os.path.exists(path):
            df = pd.read_csv(path, encoding='utf-8-sig', dtype=str, keep_default_na=False)
            hit = ((pd.to_numeric(df['Etos'], errors='coerce') == year) &
                   (pd.to_numeric(df['Minas'], errors='coerce') == month))
            df[~hit].to_csv(path, index=False, encoding='utf-8-sig')
    try:
        os.remove(archive_path(year, month))
    except OSError:
        pass


def year_stats(hist, year):
    """Totals of one year, months without a form between the first and last one, and the change
    against the previous year (only over months that exist in both years, so it is a fair comparison)."""
    months = hist.get(year, {})
    cl = sum(c for rows in months.values() for _, c, _ in rows)
    ni = sum(n for rows in months.values() for _, _, n in rows)
    present = sorted(months)
    gaps = [m for m in range(present[0], present[-1] + 1) if m not in months] if present else []
    prev, cmp_ = hist.get(year - 1, {}), None
    common = [m for m in months if m in prev]
    if common:
        cmp_ = dict(n=len(common),
                    cur_cl=sum(c for m in common for _, c, _ in months[m]),
                    cur_ni=sum(n for m in common for _, _, n in months[m]),
                    prev_cl=sum(c for m in common for _, c, _ in prev[m]),
                    prev_ni=sum(n for m in common for _, _, n in prev[m]))
    return dict(clients=cl, nights=ni, avg=(ni / cl if cl else 0), months=present, gaps=gaps, cmp=cmp_)


def missing_month_reminder(hist, today=None):
    """(year, month) of last month if its form is missing although the hotel was working until recently."""
    if not hist:
        return None
    today = today or datetime.date.today()
    py, pm = (today.year - 1, 12) if today.month == 1 else (today.year, today.month - 1)
    last = max((y, m) for y, months in hist.items() for m in months)
    gap = (py * 12 + pm) - (last[0] * 12 + last[1])
    return (py, pm) if 1 <= gap <= 2 else None


def greek_fonts():
    """(regular, bold) font names able to print Greek, falling back to Helvetica."""
    regular = bold = "Helvetica"
    pairs = [(r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
             (r"C:\Windows\Fonts\tahoma.ttf", r"C:\Windows\Fonts\tahomabd.ttf"),
             ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")]
    for reg_p, bold_p in pairs:
        if os.path.exists(reg_p):
            try:
                pdfmetrics.registerFont(TTFont('GreekFont', reg_p))
                regular = bold = 'GreekFont'
                if os.path.exists(bold_p):
                    pdfmetrics.registerFont(TTFont('GreekFontBold', bold_p))
                    bold = 'GreekFontBold'
                break
            except Exception:
                continue
    return regular, bold


def fmt_int(n):
    return f"{int(n):,}".replace(",", ".")


def fmt_avg(nights, clients):
    return f"{nights / clients:.1f}".replace(".", ",") if clients else "-"


def make_year_summary_pdf(year, hist, hotel, out_path):
    """One-page A4 summary of a year: month table, Greek vs foreign guests, top countries."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    months = hist.get(year, {})
    if not months:
        raise ValueError("Δεν υπάρχουν δεδομένα για αυτό το έτος.")
    reg, bold = greek_fonts()
    W, H = A4
    L, R = 40, W - 40
    navy, text_c, band = colors.HexColor(NAVY), colors.HexColor(TEXT), colors.HexColor('#e8f0fe')
    c = canvas.Canvas(out_path, pagesize=A4)

    c.setFillColor(navy)
    c.rect(0, H - 92, W, 92, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(bold, 19)
    c.drawString(L, H - 46, f"Σύνοψη κίνησης έτους {year}")
    c.setFont(reg, 10.5)
    who = " · ".join(x for x in (hotel.get('type'), hotel.get('name'), hotel.get('category')) if x)
    c.drawString(L, H - 68, who or "Τουριστικό κατάλυμα")

    def header(y, first, cols):
        c.setFillColor(navy)
        c.rect(L, y - 6, R - L, 22, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont(bold, 9.5)
        c.drawString(L + 8, y, first)
        for x, t in cols:
            c.drawRightString(x - 6, y, t)

    greece = norm('ΕΛΛΑΔΑ')
    cols = [(220, "Πελάτες"), (320, "Διανυκτερεύσεις"), (410, "Μέση διαμονή"), (480, "Ελλάδα"), (R, "Αλλοδαποί")]
    y = H - 130
    header(y, "Μήνας", cols)
    y -= 24
    tcl = tni = tgr = 0
    for i, month in enumerate(sorted(months)):
        rows = months[month]
        cl, ni = sum(r[1] for r in rows), sum(r[2] for r in rows)
        gr = sum(r[1] for r in rows if norm(r[0]) == greece)
        tcl, tni, tgr = tcl + cl, tni + ni, tgr + gr
        if i % 2 == 0:
            c.setFillColor(colors.HexColor('#f4f7fb'))
            c.rect(L, y - 6, R - L, 20, stroke=0, fill=1)
        c.setFillColor(text_c)
        c.setFont(reg, 10)
        c.drawString(L + 8, y, GREEK_MONTHS[month - 1] if 1 <= month <= 12 else str(month))
        for (x, _), v in zip(cols, (fmt_int(cl), fmt_int(ni), fmt_avg(ni, cl), fmt_int(gr), fmt_int(cl - gr))):
            c.drawRightString(x - 6, y, v)
        y -= 20
    c.setFillColor(band)
    c.rect(L, y - 6, R - L, 22, stroke=0, fill=1)
    c.setFillColor(navy)
    c.setFont(bold, 10.5)
    c.drawString(L + 8, y, f"ΣΥΝΟΛΟ {year}")
    for (x, _), v in zip(cols, (fmt_int(tcl), fmt_int(tni), fmt_avg(tni, tcl), fmt_int(tgr), fmt_int(tcl - tgr))):
        c.drawRightString(x - 6, y, v)

    by_country = {}
    for rows in months.values():
        for country, cl, ni in rows:
            t = by_country.setdefault(country, [0, 0])
            t[0] += cl
            t[1] += ni
    top = sorted(by_country.items(), key=lambda kv: -kv[1][0])[:8]
    y -= 48
    c.setFillColor(navy)
    c.setFont(bold, 12)
    c.drawString(L, y, "Κορυφαίες χώρες προέλευσης")
    y -= 26
    cols2 = [(300, "Πελάτες"), (410, "Διανυκτερεύσεις"), (R, "% των πελατών")]
    header(y, "Χώρα", cols2)
    y -= 24
    for i, (country, (cl, ni)) in enumerate(top):
        if i % 2 == 0:
            c.setFillColor(colors.HexColor('#f4f7fb'))
            c.rect(L, y - 6, R - L, 20, stroke=0, fill=1)
        c.setFillColor(text_c)
        c.setFont(reg, 10)
        c.drawString(L + 8, y, country)
        pct = f"{cl / tcl * 100:.1f}%".replace(".", ",") if tcl else "-"
        for (x, _), v in zip(cols2, (fmt_int(cl), fmt_int(ni), pct)):
            c.drawRightString(x - 6, y, v)
        y -= 20

    c.setFillColor(colors.HexColor(MUTED))
    c.setFont(reg, 8.5)
    c.drawString(L, 30, f"Δημιουργήθηκε στις {datetime.date.today().strftime('%d/%m/%Y')} από το πρόγραμμα ΕΛΣΤΑΤ")
    c.save()


def process_elstat_report(excel_path, template_pdf, output_pdf, settings_data, monthly_data,
                         rows_override=None, save_history=True):
    # 1. Φόρτωση Γραμματοσειράς
    font_name = "Helvetica"
    font_paths = [
        r"C:\Windows\Fonts\arial.ttf", 
        r"C:\Windows\Fonts\tahoma.ttf", 
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            pdfmetrics.registerFont(TTFont('GreekFont', fp))
            font_name = 'GreekFont'
            break

    # 3. Αντιστοίχιση Χωρών (με Εμιράτα, Σιγκαπούρη, No Country)
    country_mapping = COUNTRY_MAPPING

    category_map = {
        '2.1': ['ΑΥΣΤΡΙΑ', 'ΒΕΛΓΙΟ', 'ΒΟΥΛΓΑΡΙΑ', 'ΓΑΛΛΙΑ', 'ΓΕΡΜΑΝΙΑ', 'ΔΑΝΙΑ', 'ΕΣΘΟΝΙΑ', 'ΗΝΩΜΕΝΟ ΒΑΣΙΛΕΙΟ', 'ΙΡΛΑΝΔΙΑ', 'ΙΣΠΑΝΙΑ', 'ΙΤΑΛΙΑ', 'ΚΡΟΑΤΙΑ', 'ΚΥΠΡΟΣ', 'ΛΕΤΟΝΙΑ', 'ΛΙΘΟΥΑΝΙΑ', 'ΛΟΥΞΕΜΒΟΥΡΓΟ', 'ΜΑΛΤΑ', 'ΟΛΛΑΝΔΙΑ', 'ΟΥΓΓΑΡΙΑ', 'ΠΟΛΩΝΙΑ', 'ΠΟΡΤΟΓΑΛΙΑ', 'ΡΟΥΜΑΝΙΑ', 'ΣΛΟΒΑΚΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 'ΣΛΟΒΕΝΙΑ', 'ΣΟΥΗΔΙΑ', 'ΤΣΕΧΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 'ΦΙΝΛΑΝΔΙΑ'],
        '2.2': ['ΕΛΒΕΤΙΑ', 'ΛΙΧΤΕΝΣΤΑΪΝ', 'ΙΣΛΑΝΔΙΑ', 'ΝΟΡΒΗΓΙΑ'],
        '2.3': ['ΑΛΒΑΝΙΑ', 'ΓΕΩΡΓΙΑ', 'ΜΑΥΡΟΒΟΥΝΙΟ', 'ΜΟΛΔΑΒΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 'ΟΥΚΡΑΝΙΑ', 'ΒΟΡΕΙΑΣ ΜΑΚΕΔΟΝΙΑΣ ΔΗΜΟΚΡΑΤΙΑ', 'ΡΩΣΙΑ', 'ΣΕΡΒΙΑ', 'ΤΟΥΡΚΙΑ', 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΕΥΡΩΠΗΣ'],
        '3': ['Η.Π.Α.', 'ΚΑΝΑΔΑΣ', 'ΑΡΓΕΝΤΙΝΗ', 'ΒΡΑΖΙΛΙΑ', 'ΜΕΞΙΚΟ', 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΜΕΡΙΚΗΣ'],
        '4': ['ΙΑΠΩΝΙΑ', 'ΙΝΔΙΑ', 'ΙΡΑΚ', 'ΙΡΑΝ', 'ΙΣΡΑΗΛ', 'ΚΙΝΑΣ ΛΑΪΚΗ ΔΗΜΟΚΡΑΤΙΑ', 'ΚΟΡΕΑΣ ΔΗΜΟΚΡΑΤΙΑ', 'ΛΙΒΑΝΟΣ', 'ΠΑΚΙΣΤΑΝ', 'ΣΑΟΥΔΙΚΗ ΑΡΑΒΙΑ', 'ΣΥΡΙΑ', 'ΦΙΛΙΠΠΙΝΕΣ', 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΣΙΑΣ'],
        '5': ['ΑΙΓΥΠΤΟΣ', 'ΑΙΘΙΟΠΙΑ', 'ΑΛΓΕΡΙΑ', 'ΝΙΓΗΡΙΑ', 'ΝΟΤΙΑ ΑΦΡΙΚΗ', 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΦΡΙΚΗΣ'],
        '6': ['ΑΥΣΤΡΑΛΙΑ', 'ΝΕΑ ΖΗΛΑΝΔΙΑ ΚΑΙ ΛΟΙΠΕΣ ΠΕΡΙΟΧΕΣ'],
        '7': ['Δεν δήλωσαν χώρα']
    }

    country_to_cat = {}
    for cat, countries in category_map.items():
        for c in countries:
            country_to_cat[norm(c)] = cat

    pdf_data = {}
    totals = {key: {'clients': 0, 'nights': 0} for key in category_map.keys()}

    mapping_norm = COUNTRY_MAPPING_NORM
    unmapped = []
    excel_total = None

    if rows_override is not None:        # re-printing an old month straight from the history
        rows, excel_total = list(rows_override), None
    else:
        rows, excel_total = read_nationality_rows(excel_path)
    for country_excel, pax, nights in rows:
        pdf_target = resolve_country(country_excel)
        if pdf_target is None:
            unmapped.append((country_excel, pax, nights))
            continue
        pdf_data.setdefault(pdf_target, {'clients': 0, 'nights': 0})
        pdf_data[pdf_target]['clients'] += pax
        pdf_data[pdf_target]['nights'] += nights
        cat = country_to_cat.get(pdf_target)
        if cat:
            totals[cat]['clients'] += pax
            totals[cat]['nights'] += nights

    totals['2'] = {
        'clients': totals['2.1']['clients'] + totals['2.2']['clients'] + totals['2.3']['clients'],
        'nights': totals['2.1']['nights'] + totals['2.2']['nights'] + totals['2.3']['nights']
    }
    
    totals['8'] = {
        'clients': totals['2']['clients'] + totals['3']['clients'] + totals['4']['clients'] + totals['5']['clients'] + totals['6']['clients'] + totals['7']['clients'],
        'nights': totals['2']['nights'] + totals['3']['nights'] + totals['4']['nights'] + totals['5']['nights'] + totals['6']['nights'] + totals['7']['nights']
    }

    greece_cl = pdf_data.get(norm('ΕΛΛΑΔΑ'), {}).get('clients', 0)
    greece_ni = pdf_data.get(norm('ΕΛΛΑΔΑ'), {}).get('nights', 0)
    
    totals['9'] = {
        'clients': greece_cl + totals['8']['clients'],
        'nights': greece_ni + totals['8']['nights']
    }

    reader = pypdf.PdfReader(template_pdf)
    if len(reader.pages) < 2:
        raise ValueError("Το template PDF πρέπει να έχει 2 σελίδες.")
    coord_map = {0: {}, 1: {}}
    for i in range(2):
        page = reader.pages[i]
        def visitor_text(text, cm, tm, fontDict, fontSize, _i=i):
            n = norm(text)
            if not n:
                return
            for target in sorted(pdf_data.keys(), key=len, reverse=True):
                if n.startswith(target):
                    coord_map[_i].setdefault(target, tm[5])
                    break
        page.extract_text(visitor_text=visitor_text)

    if norm('ΤΟΥΡΚΙΑ') in pdf_data and norm('ΤΟΥΡΚΙΑ') not in coord_map[1]:
        coord_map[1][norm('ΤΟΥΡΚΙΑ')] = 555.2222

    fixed_y = {
        '2.1': 46.5, '2.2': 699.5, '2.3': 520.0, '2': 500.1,
        '3': 407.1, '4': 231.7, '5': 141.5, '6': 95.2, 
        '8': 60.5, '9': 43.5
    }

    def create_overlay(page_num, data_dict, coords, totals_dict):
        packet = io.BytesIO()
        can = canvas.Canvas(packet)
        
        # ΣΕΛΙΔΑ 1
        if page_num == 0:
            can.setFont(font_name, 11)
            can.drawString(160, 613, settings_data['type'])
            can.drawString(340, 612, settings_data['category'])
            can.drawString(165, 597, settings_data['name'])
            can.drawString(110, 548, settings_data['phone'])
            can.drawString(420, 548, settings_data['email'])  
            can.drawString(435, 385, settings_data['company'])
            
            afm_str = settings_data['afm']
            afm_start_x = 424  
            afm_step = 14.0    
            
            for i, digit in enumerate(afm_str):
                can.drawString(afm_start_x + (i * afm_step), 342, digit)
            
            can.drawString(380, 675, monthly_data['year'])
            can.drawString(480, 675, monthly_data['month'])
            
            date_from_str = f"01           01           {monthly_data['year']}"
            can.drawString(420, 261, date_from_str)
            
            dt_parts = monthly_data['date_to'].split('/')
            if len(dt_parts) == 3:
                date_to_str = f"{dt_parts[0]}           {dt_parts[1]}           {dt_parts[2]}"
            else:
                date_to_str = monthly_data['date_to']
            can.drawString(420, 233, date_to_str)
            
            days_val = monthly_data['days_of_month']
            if len(days_val) == 1:
                days_val = "0" + days_val  
                
            days_start_x = 457  
            days_step = 16.0    
            
            for i, digit in enumerate(days_val):
                can.drawString(days_start_x + (i * days_step), 169, digit)
                
            emp_val = monthly_data.get('employees', '0')
            emp_val = emp_val.zfill(3) if len(emp_val) < 3 else emp_val
            
            emp_start_x = 448  
            emp_step = 16.0    
            emp_y = 94       
            
            for i, digit in enumerate(emp_val):
                can.drawString(emp_start_x + (i * emp_step), emp_y, digit)
            
            can.setFont(font_name, 10)
            if totals_dict['2.1']['clients'] > 0:
                can.drawString(265, fixed_y['2.1'], str(totals_dict['2.1']['clients']))
                can.drawString(335, fixed_y['2.1'], str(totals_dict['2.1']['nights']))

        # ΣΕΛΙΔΑ 2
        else:
            can.setFont(font_name, 10)
            for key in fixed_y.keys():
                if key != '2.1' and totals_dict[key]['clients'] > 0:
                    can.drawString(265, fixed_y[key], str(totals_dict[key]['clients']))
                    can.drawString(335, fixed_y[key], str(totals_dict[key]['nights']))
                    
            today = datetime.date.today()
            year_suffix = str(today.year)[-2:] if str(today.year).startswith("20") else str(today.year)
            today_str = f"{today.strftime('%d')}      {today.strftime('%m')}          {year_suffix}"
            can.drawString(454, 206, today_str)

        can.setFont(font_name, 9)
        for country, y_val in coords.items():
            if country in data_dict:
                cl, ni = data_dict[country]['clients'], data_dict[country]['nights']
                can.drawString(265, y_val, str(cl))
                can.drawString(335, y_val, str(ni))
                
        can.save()
        packet.seek(0)
        return pypdf.PdfReader(packet)

    writer = pypdf.PdfWriter()
    for i in range(2):
        page = reader.pages[i]
        overlay_pdf = create_overlay(i, pdf_data, coord_map[i], totals)
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    if save_history:
        # Αυτόματη δημιουργία ή ενημέρωση ιστορικού CSV
        csv_filename = HISTORY_CSV
        csv_data = []
        for country, vals in pdf_data.items():
            csv_data.append({
                'Etos': str(int(monthly_data['year'])),
                'Minas': str(int(monthly_data['month'])),
                'Chwra': country,
                'Pelates': vals['clients'],
                'Dianykterefseis': vals['nights']
            })
        df_csv = pd.DataFrame(csv_data)
    
        if os.path.exists(csv_filename):
            # re-running the same month replaces its old rows instead of duplicating them
            old_df = pd.read_csv(csv_filename, encoding='utf-8-sig', dtype=str)
            keep = ~((pd.to_numeric(old_df['Etos'], errors='coerce') == int(monthly_data['year'])) &
                     (pd.to_numeric(old_df['Minas'], errors='coerce') == int(monthly_data['month'])))
            df_csv = pd.concat([old_df[keep], df_csv.astype(str)], ignore_index=True)
        ensure_data_dir()
        df_csv.to_csv(csv_filename, index=False, encoding='utf-8-sig')

    computed = (totals['9']['clients'], totals['9']['nights'])
    gr = pdf_data.get(norm('ΕΛΛΑΔΑ'), {'clients': 0, 'nights': 0})
    summary = [
        ("1.  Ελλάδα", gr['clients'], gr['nights'], False),
        ("2.1  Χώρες Ευρωπαϊκής Ένωσης", totals['2.1']['clients'], totals['2.1']['nights'], False),
        ("2.2  Χώρες Ε.Ζ.Ε.Σ. (EFTA)", totals['2.2']['clients'], totals['2.2']['nights'], False),
        ("2.3  Λοιπές χώρες Ευρώπης", totals['2.3']['clients'], totals['2.3']['nights'], False),
        ("2.  Σύνολο Ευρώπης (εκτός Ελλάδας)", totals['2']['clients'], totals['2']['nights'], True),
        ("3.  Σύνολο Αμερικής", totals['3']['clients'], totals['3']['nights'], True),
        ("4.  Σύνολο Ασίας", totals['4']['clients'], totals['4']['nights'], True),
        ("5.  Σύνολο Αφρικής", totals['5']['clients'], totals['5']['nights'], True),
        ("6.  Σύνολο Ωκεανίας", totals['6']['clients'], totals['6']['nights'], True),
        ("7.  Δεν δήλωσαν χώρα", totals['7']['clients'], totals['7']['nights'], False),
        ("8.  Σύνολο αλλοδαπών", totals['8']['clients'], totals['8']['nights'], True),
        ("9.  ΓΕΝΙΚΟ ΣΥΝΟΛΟ", totals['9']['clients'], totals['9']['nights'], True),
    ]
    return {'unmapped': unmapped, 'excel_total': excel_total, 'computed_total': computed, 'summary': summary}


# ----------------------------------------------------------------- look & feel
BG = "#eef2f7"
CARD = "#ffffff"
NAVY = "#14304f"
ACCENT = "#2563eb"
ACCENT_DARK = "#1d4ed8"
TEAL = "#0f766e"
TEAL_DARK = "#115e59"
BORDER = "#d3dce8"
TEXT = "#1e293b"
MUTED = "#64748b"
GOOD, BAD, WARN = "#15803d", "#b91c1c", "#b45309"
FONT = "Segoe UI" if sys.platform.startswith("win") else "Arial"
GREEK_MONTHS = ["Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος", "Ιούνιος",
                "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος", "Νοέμβριος", "Δεκέμβριος"]


def autowrap(label, container, margin=40, minimum=160):
    """Makes a label re-wrap its text whenever `container` is resized."""
    def _update(_event=None):
        w = container.winfo_width()
        if w and w > 1:
            label.configure(wraplength=max(minimum, w - margin))
    container.bind('<Configure>', _update, add='+')


class Card(tk.Frame):
    """White 'card' with a thin border, an optional numbered badge and a title.
    Put your widgets into `card.body`."""

    def __init__(self, parent, number=None, title=""):
        super().__init__(parent, bg=BORDER)
        inner = tk.Frame(self, bg=CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        head = tk.Frame(inner, bg=CARD)
        head.pack(fill="x", padx=16, pady=(14, 2))
        if number is not None:
            tk.Label(head, text=str(number), bg=ACCENT, fg="white", width=2,
                     font=(FONT, 10, 'bold')).pack(side="left")
        tk.Label(head, text=title, bg=CARD, fg=NAVY,
                 font=(FONT, 12, 'bold')).pack(side="left", padx=(10 if number is not None else 0, 0))
        self.body = ttk.Frame(inner, style='Card.TFrame', padding=(16, 6, 16, 16))
        self.body.pack(fill="x")


GREEK_MONTHS_SHORT = ["Ιαν", "Φεβ", "Μαρ", "Απρ", "Μάι", "Ιουν", "Ιουλ", "Αυγ", "Σεπ", "Οκτ", "Νοε", "Δεκ"]


class StatCard(tk.Frame):
    """Small dashboard tile: caption, big number and a coloured sub-line."""

    def __init__(self, parent, title):
        super().__init__(parent, bg=BORDER)
        inner = tk.Frame(self, bg=CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text=title, bg=CARD, fg=MUTED, font=(FONT, 9, 'bold'),
                 anchor="w").pack(fill="x", padx=14, pady=(10, 0))
        self.value = tk.Label(inner, text="–", bg=CARD, fg=NAVY, font=(FONT, 20, 'bold'), anchor="w")
        self.value.pack(fill="x", padx=14)
        self.sub = tk.Label(inner, text="", bg=CARD, fg=MUTED, font=(FONT, 9), anchor="w", justify="left")
        self.sub.pack(fill="x", padx=14, pady=(0, 10))
        inner.bind('<Configure>', lambda e: self.sub.configure(wraplength=max(100, e.width - 28)))

    def set(self, value, sub="", color=MUTED):
        self.value.configure(text=value)
        self.sub.configure(text=sub, fg=color)


class HistoryChart(tk.Canvas):
    """Bar chart with one bar per month. Clicking a bar calls on_pick(month)."""

    def __init__(self, parent, on_pick=None, height=160):
        super().__init__(parent, height=height, bg=CARD, highlightthickness=1,
                         highlightbackground=BORDER, borderwidth=0)
        self.values, self.title, self.color, self.selected = [0] * 12, "", ACCENT, None
        self.on_pick = on_pick
        self._slots = []
        self.bind('<Configure>', lambda e: self.draw())
        self.bind('<Button-1>', self._click)

    def set_data(self, values, title, color=ACCENT, selected=None):
        self.values, self.title, self.color, self.selected = list(values), title, color, selected
        self.draw()

    @staticmethod
    def _nice(v):
        if v <= 0:
            return 1
        mag = 10 ** math.floor(math.log10(v))
        for f in (1, 2, 2.5, 5, 10):
            if v <= f * mag:
                return f * mag
        return 10 * mag

    @staticmethod
    def _fmt(v):
        return f"{int(round(v)):,}".replace(",", ".")

    def draw(self):
        self.delete('all')
        self._slots = []
        w, h = self.winfo_width(), self.winfo_height()
        if w < 120 or h < 90:
            return
        left, right, top, bottom = 52, 14, 38, 26
        plot_h = h - top - bottom
        top_val = self._nice(max(self.values))
        self.create_text(14, 14, text=self.title, anchor="w", fill=NAVY, font=(FONT, 10, 'bold'))
        for frac in (0, .5, 1):
            y = top + plot_h * (1 - frac)
            self.create_line(left, y, w - right, y, fill='#e5eaf1')
            self.create_text(left - 8, y, text=self._fmt(top_val * frac), anchor="e", fill=MUTED, font=(FONT, 8))
        slot = (w - left - right) / 12
        bw = slot * 0.6
        for i, v in enumerate(self.values):
            cx = left + slot * (i + .5)
            if v > 0:
                y0 = top + plot_h * (1 - v / top_val)
                fill = NAVY if self.selected == i + 1 else self.color
                self.create_rectangle(cx - bw / 2, y0, cx + bw / 2, top + plot_h, fill=fill, outline='')
                self.create_text(cx, y0 - 8, text=self._fmt(v), fill=TEXT, font=(FONT, 8))
                self._slots.append((cx - slot / 2, cx + slot / 2, i + 1))
            else:
                self.create_line(cx - bw / 2, top + plot_h - 1, cx + bw / 2, top + plot_h - 1,
                                 fill='#cbd5e1', width=2)
            self.create_text(cx, h - 12, text=GREEK_MONTHS_SHORT[i], fill=MUTED, font=(FONT, 8))

    def _click(self, event):
        for x0, x1, month in self._slots:
            if x0 <= event.x <= x1 and self.on_pick:
                self.on_pick(month)
                return


class ScrollFrame(ttk.Frame):
    """A frame that stretches to the window width and gets a scrollbar only when it is too tall."""

    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, bg=BG)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.body = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.canvas._sf = self                      # lets the mouse-wheel handler find us
        self.body.bind('<Configure>', self._on_body)
        self.canvas.bind('<Configure>', self._on_canvas)

    def _on_body(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._update_bar()

    def _on_canvas(self, event):
        self.canvas.itemconfigure(self._win, width=event.width)   # content follows the window width
        self._update_bar()

    def _needs_scroll(self):
        return self.body.winfo_reqheight() > self.canvas.winfo_height() > 1

    def _update_bar(self):
        if self._needs_scroll():
            self.vbar.grid()
        else:
            self.vbar.grid_remove()
            self.canvas.yview_moveto(0)

    def scroll(self, units):
        if self._needs_scroll():
            self.canvas.yview_scroll(units, "units")


def setup_styles(root):
    st = ttk.Style(root)
    st.theme_use('clam')
    st.configure('.', font=(FONT, 10), background=BG, foreground=TEXT)
    st.configure('TFrame', background=BG)
    st.configure('Card.TFrame', background=CARD)
    st.configure('TLabel', background=BG, foreground=TEXT)
    st.configure('Card.TLabel', background=CARD)
    st.configure('CardBold.TLabel', background=CARD, font=(FONT, 10, 'bold'))
    st.configure('TSeparator', background=BORDER)
    st.configure('Status.TLabel', background=CARD, foreground=MUTED, font=(FONT, 9), padding=(14, 6))
    st.configure('TRadiobutton', background=BG, focuscolor=BG)
    for name in ('Card.TRadiobutton', 'Card.TCheckbutton'):
        st.configure(name, background=CARD, focuscolor=CARD)
        st.map(name, background=[('active', CARD)])
    st.configure('TCombobox', fieldbackground='white', background='#e8edf4', bordercolor=BORDER,
                 lightcolor=BORDER, darkcolor=BORDER, arrowcolor=TEXT, padding=4)
    st.map('TCombobox', fieldbackground=[('readonly', 'white')], bordercolor=[('focus', ACCENT)])
    root.option_add('*TCombobox*Listbox.font', (FONT, 10))

    # coloured text styles, one set for the grey background and one for white cards
    for p, bg in (('', BG), ('Card', CARD)):
        st.configure(f'{p}Hint.TLabel', background=bg, foreground=MUTED, font=(FONT, 9))
        st.configure(f'{p}Good.TLabel', background=bg, foreground=GOOD)
        st.configure(f'{p}Bad.TLabel', background=bg, foreground=BAD)
        st.configure(f'{p}Warn.TLabel', background=bg, foreground=WARN)
    st.configure('GoodTitle.TLabel', foreground=GOOD, font=(FONT, 15, 'bold'))
    st.configure('BadTitle.TLabel', foreground=WARN, font=(FONT, 15, 'bold'))

    # status "pills" shown above the main buttons
    for name, bg, fg in (('StatusGood', '#dcfce7', '#166534'), ('StatusBad', '#fee2e2', '#991b1b'),
                         ('StatusWarn', '#fef3c7', '#92400e')):
        st.configure(f'{name}.TLabel', background=bg, foreground=fg, padding=(12, 8))

    # tabs
    st.configure('TNotebook', background=BG, borderwidth=0, tabmargins=(2, 4, 2, 0))
    st.configure('TNotebook.Tab', padding=(22, 9), font=(FONT, 10, 'bold'), background='#d9e1ec',
                 foreground=MUTED, borderwidth=0, focuscolor='#d9e1ec')
    st.map('TNotebook.Tab', background=[('selected', ACCENT), ('active', '#c8d4e4')],
           foreground=[('selected', 'white')], focuscolor=[('selected', ACCENT)])

    # buttons (flat, coloured)
    def button(name, bg, fg, active, border=None, font=None, padding=(16, 8)):
        border = border or bg
        opts = dict(background=bg, foreground=fg, bordercolor=border, lightcolor=bg, darkcolor=bg,
                    focuscolor=bg, borderwidth=1, relief='flat', padding=padding)
        if font:
            opts['font'] = font
        st.configure(name, **opts)
        st.map(name,
               background=[('disabled', '#c7d0dc'), ('pressed', active), ('active', active)],
               foreground=[('disabled', '#f8fafc')],
               bordercolor=[('disabled', '#c7d0dc'), ('active', active)],
               lightcolor=[('pressed', active), ('active', active)],
               darkcolor=[('pressed', active), ('active', active)])
    button('TButton', '#e8edf4', TEXT, '#d6deea', border=BORDER)
    button('Primary.TButton', ACCENT, 'white', ACCENT_DARK, font=(FONT, 12, 'bold'), padding=(14, 11))
    button('Print.TButton', TEAL, 'white', TEAL_DARK, font=(FONT, 12, 'bold'), padding=(14, 11))
    button('PrintSm.TButton', TEAL, 'white', TEAL_DARK, font=(FONT, 10, 'bold'), padding=(14, 7))
    button('Danger.TButton', '#fee2e2', BAD, '#fecaca', border='#fca5a5', padding=(14, 7))

    # entries
    st.configure('TEntry', fieldbackground='white', bordercolor=BORDER, lightcolor=BORDER,
                 darkcolor=BORDER, padding=6)
    st.map('TEntry', bordercolor=[('focus', ACCENT)], lightcolor=[('focus', ACCENT)],
           darkcolor=[('focus', ACCENT)])
    st.configure('Bad.TEntry', fieldbackground='#fee2e2', bordercolor='#f87171',
                 lightcolor='#f87171', darkcolor='#f87171', padding=6)

    # scrollbars
    st.configure('Vertical.TScrollbar', background='#c5cfdd', troughcolor=BG, bordercolor=BG,
                 arrowcolor=MUTED, relief='flat')

    # result table
    st.configure('Treeview', rowheight=28, font=(FONT, 10), background=CARD, fieldbackground=CARD,
                 foreground=TEXT, bordercolor=BORDER, borderwidth=1)
    st.configure('Treeview.Heading', font=(FONT, 10, 'bold'), background=NAVY, foreground='white',
                 padding=(8, 7), relief='flat')
    st.map('Treeview.Heading', background=[('active', NAVY)])
    st.map('Treeview', background=[('selected', '#dbeafe')], foreground=[('selected', TEXT)])


class ResultDialog(tk.Toplevel):
    """Shown after the PDF is made: the totals that went into the form, checks and print button."""

    def __init__(self, parent, result, out_path, warnings_txt, print_status=None):
        super().__init__(parent)
        self.title("Αποτέλεσμα")
        self.configure(bg=BG)
        self.transient(parent)
        self.minsize(520, 480)
        self.out_path = out_path
        ok = not warnings_txt

        body = ttk.Frame(self, padding=18)
        body.pack(fill="both", expand=True)

        def note(text, style, pady=(2, 0)):
            lbl = ttk.Label(body, text=text, style=style, justify="left")
            lbl.pack(anchor="w", fill="x", pady=pady)
            autowrap(lbl, self, margin=50, minimum=260)
            return lbl

        if ok:
            ttk.Label(body, text="✓ Το δελτίο δημιουργήθηκε επιτυχώς", style='GoodTitle.TLabel').pack(anchor="w")
            gc, gn = result['computed_total']
            note(f"Τα σύνολα ταιριάζουν με το Excel του Protel ({gc} πελάτες / {gn} διανυκτερεύσεις).",
                 'Good.TLabel', pady=(2, 12))
        else:
            ttk.Label(body, text="⚠ Το δελτίο δημιουργήθηκε, αλλά θέλει έλεγχο",
                      style='BadTitle.TLabel').pack(anchor="w")
            for w in warnings_txt:
                note("• " + w, 'Bad.TLabel', pady=(4, 0))
            ttk.Label(body, text="").pack()

        # buttons first (bottom) so they never get squeezed out when the window is small
        btns = ttk.Frame(body)
        btns.pack(side="bottom", fill="x", pady=(14, 0))
        ttk.Button(btns, text="Εκτύπωση", style='Print.TButton',
                   command=self.print_now).pack(side="left")
        ttk.Button(btns, text="Άνοιγμα PDF", command=lambda: open_file(out_path)).pack(side="left", padx=8)
        ttk.Button(btns, text="Άνοιγμα φακέλου",
                   command=lambda: open_file(os.path.dirname(out_path) or ".")).pack(side="left")
        ttk.Button(btns, text="Κλείσιμο", command=self.destroy).pack(side="right")

        self.print_var = tk.StringVar()
        self.print_lbl = ttk.Label(body, textvariable=self.print_var, style='Hint.TLabel', justify="left")
        self.print_lbl.pack(side="bottom", anchor="w", fill="x", pady=(10, 0))
        autowrap(self.print_lbl, self, margin=50, minimum=260)
        if print_status:
            self._show_print(*print_status)

        steps = ttk.Label(body, justify="left",
                          text="Επόμενα βήματα: εκτυπώστε το δελτίο, γράψτε το όνομα, βάλτε υπογραφή και σφραγίδα "
                               "και υποβάλετε το στην ΕΛΣΤΑΤ.")
        steps.pack(side="bottom", anchor="w", fill="x", pady=(10, 0))
        autowrap(steps, self, margin=50, minimum=260)
        where = ttk.Label(body, text="Αποθηκεύτηκε στο:\n" + out_path, style='Hint.TLabel', justify="left")
        where.pack(side="bottom", anchor="w", fill="x", pady=(10, 0))
        autowrap(where, self, margin=50, minimum=260)

        tree_frame = ttk.Frame(body)
        tree_frame.pack(side="top", fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=('c', 'n'), show='tree headings',
                            height=min(len(result['summary']), 12))
        tree.heading('#0', text='Κατηγορία (όπως στο δελτίο)', anchor='w')
        tree.heading('c', text='Πελάτες')
        tree.heading('n', text='Διανυκτερεύσεις')
        tree.column('#0', width=270, minwidth=180, stretch=True)
        tree.column('c', width=90, minwidth=70, anchor='e', stretch=False)
        tree.column('n', width=130, minwidth=100, anchor='e', stretch=False)
        tree.tag_configure('bold', font=(FONT, 10, 'bold'), background='#e8f0fe')
        tree.tag_configure('plain', background=CARD)
        for label, c, n, bold in result['summary']:
            tree.insert('', 'end', text=label, values=(c, n), tags=('bold' if bold else 'plain',))
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.bind('<Control-p>', lambda e: self.print_now())
        self.bind('<Escape>', lambda e: self.destroy())
        self.update_idletasks()
        try:   # centre over the main window
            x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
            self.grab_set()
        except (tk.TclError, TypeError):
            pass

    def _show_print(self, ok, msg):
        self.print_var.set(msg)
        self.print_lbl.configure(style='Good.TLabel' if ok else 'Warn.TLabel')

    def print_now(self):
        self._show_print(*print_file(self.out_path))


class ElstatApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ΕΛΣΤΑΤ – Μηνιαίο Δελτίο Κίνησης Καταλύματος")
        self.geometry("920x840")
        self.minsize(640, 560)
        self.configure(bg=BG)

        self._storage_notices = init_storage()
        self.prefs = load_prefs()
        self.current_settings = load_hotel_settings()
        self.output_custom = False      # True once the user picks the output file themselves
        self._busy = False
        self._loaded_excel = None

        setup_styles(self)
        self._init_vars()
        self._build_ui()
        self.pv['filename_pattern'].trace_add('write', self._update_name_note)
        self._after_prefs_change()

        watched = (list(self.set_vars.values()) + list(self.mon_vars.values()) +
                   [self.excel_var, self.out_var, self.pdf_var])
        for v in watched:
            v.trace_add('write', self.refresh_validation)
        self.refresh_validation()
        if not os.path.exists(SETTINGS_CSV):
            self.notebook.select(self.tab_hotel)

        self.bind_all('<MouseWheel>', self._on_wheel)
        self.bind_all('<Button-4>', self._on_wheel)
        self.bind_all('<Button-5>', self._on_wheel)
        self.bind('<Control-o>', lambda e: self.browse_excel())
        self.bind('<Control-Return>', self._try_generate)
        self.bind('<Control-p>', lambda e: self._try_generate(print_after=True))
        self._center()
        self.after(150, self.excel_btn.focus_set)
        if self._storage_notices:
            self.after(400, lambda: messagebox.showinfo("Ρυθμίσεις", "\n\n".join(self._storage_notices),
                                                        parent=self))

    # ------------------------------------------------------------------ setup
    def _init_vars(self):
        cs = self.current_settings
        self.set_vars = {k: tk.StringVar(value=cs.get(k, '')) for k in SETTING_KEYS}
        # month details start EMPTY: the Excel fills most of them, the rest are typed by hand
        self.mon_vars = {k: tk.StringVar() for k in ('year', 'month', 'date_to', 'days_of_month', 'employees')}
        self.excel_var = tk.StringVar()
        self.out_var = tk.StringVar(value=self._suggest_output())
        saved_tpl = self.prefs.get('template_pdf', '')
        self.pdf_var = tk.StringVar(value=saved_tpl if saved_tpl and os.path.isfile(saved_tpl) else DEFAULT_TEMPLATE)
        self.period_var = tk.StringVar()
        self.month_name_var = tk.StringVar()
        self.afm_counter_var = tk.StringVar()
        self.check_var = tk.StringVar()
        self.entries = {}
        self.status_var = tk.StringVar(value="Έτοιμο.")
        self.hist_year_var = tk.StringVar()
        self.hist_metric_var = tk.StringVar(value='nights')
        self._hist = {}
        self._reminder_dismissed = False
        # settings-tab variables
        self.pv = {k: tk.StringVar(value=self.prefs[k])
                   for k in ('output_mode', 'output_dir', 'filename_pattern', 'excel_dir')}
        self.pv.update({k: tk.BooleanVar(value=self.prefs[k]) for k in BOOL_PREFS})
        self.data_dir_var = tk.StringVar(value=DATA_DIR)
        self.name_note_var = tk.StringVar()
        self.out_note_var = tk.StringVar()

    def _build_ui(self):
        header = tk.Frame(self, bg=NAVY)
        header.pack(fill="x")
        row = tk.Frame(header, bg=NAVY)
        row.pack(fill="x", padx=22, pady=16)
        tk.Label(row, text="ΕΛΣΤΑΤ", bg=ACCENT, fg="white", font=(FONT, 11, 'bold'),
                 padx=12, pady=6).pack(side="left", anchor="n")
        texts = tk.Frame(row, bg=NAVY)
        texts.pack(side="left", fill="x", expand=True, padx=(16, 0))
        title = tk.Label(texts, text="Μηνιαίο Δελτίο Κίνησης Τουριστικού Καταλύματος", bg=NAVY, fg="white",
                         font=(FONT, 16, 'bold'), justify="left")
        title.pack(anchor="w")
        sub = tk.Label(texts, text="Διαλέξτε το Excel από το Protel και το πρόγραμμα συμπληρώνει και "
                                   "εκτυπώνει το δελτίο της ΕΛΣΤΑΤ.", bg=NAVY, fg="#b6c4d8",
                       font=(FONT, 10), justify="left")
        sub.pack(anchor="w", pady=(2, 0))
        autowrap(title, header, margin=150)
        autowrap(sub, header, margin=150)
        tk.Frame(self, height=3, bg=ACCENT).pack(fill="x")      # accent stripe

        bar = ttk.Frame(self, style='Card.TFrame')
        bar.pack(side="bottom", fill="x")
        ttk.Separator(bar, orient="horizontal").pack(fill="x")
        ttk.Label(bar, textvariable=self.status_var, style='Status.TLabel').pack(fill="x")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=12, pady=(10, 0))
        self.tab_report = ttk.Frame(self.notebook)
        self.tab_hotel = ttk.Frame(self.notebook)
        self.tab_history = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_report, text="Δελτίο μήνα")
        self.notebook.add(self.tab_hotel, text="Στοιχεία ξενοδοχείου")
        self.notebook.add(self.tab_history, text="Ιστορικό")
        self.notebook.add(self.tab_settings, text="Ρυθμίσεις")
        self._build_report_tab()
        self._build_hotel_tab()
        self._build_history_tab()
        self._build_settings_tab()
        self.refresh_history()

    def _build_report_tab(self):
        tab = self.tab_report
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)          # scrolling steps take all spare space
        sf = ScrollFrame(tab)
        sf.grid(row=0, column=0, sticky="nsew")
        body = sf.body
        body.columnconfigure(0, weight=1)

        # ---- reminder banner (only shown when last month's form is missing)
        self.reminder_frame = tk.Frame(body, bg='#fef3c7', highlightthickness=1, highlightbackground='#f5d68a')
        self.reminder_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(12, 0))
        self.reminder_label = tk.Label(self.reminder_frame, text="", bg='#fef3c7', fg='#92400e',
                                       font=(FONT, 10), justify="left", anchor="w")
        self.reminder_label.pack(side="left", fill="x", expand=True, padx=12, pady=9)
        ttk.Button(self.reminder_frame, text="Κλείσιμο", command=self._dismiss_reminder).pack(side="right", padx=8)
        autowrap(self.reminder_label, body, margin=190)
        self.reminder_frame.grid_remove()

        # ---- Step 1
        f1 = Card(body, 1, "Αρχείο από το Protel")
        f1.grid(row=1, column=0, sticky="ew", padx=4, pady=(12, 6))
        c1 = f1.body
        c1.columnconfigure(0, weight=1)
        ent = ttk.Entry(c1, textvariable=self.excel_var)
        ent.grid(row=0, column=0, sticky="ew")
        ent.bind('<FocusOut>', self._excel_typed)
        ent.bind('<Return>', self._excel_typed)
        self.excel_btn = ttk.Button(c1, text="Επιλογή Excel…", command=self.browse_excel)
        self.excel_btn.grid(row=0, column=1, padx=(8, 0))
        ttk.Button(c1, text="Πιο πρόσφατο", command=self.use_recent_excel).grid(row=0, column=2, padx=(6, 0))
        self.period_label = ttk.Label(c1, textvariable=self.period_var, style='CardHint.TLabel', justify="left")
        self.period_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 0))
        autowrap(self.period_label, body, margin=110)

        # live preview of what the Excel contains (appears after a file is chosen)
        self.preview_frame = ttk.Frame(c1, style='Card.TFrame')
        self.preview_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        self.preview_frame.columnconfigure(0, weight=1)
        self.preview_summary = ttk.Label(self.preview_frame, text="", style='CardBold.TLabel', justify="left")
        self.preview_summary.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        autowrap(self.preview_summary, body, margin=110)
        self.preview_tree = ttk.Treeview(self.preview_frame, columns=('p', 'n', 's'), show='tree headings', height=6)
        self.preview_tree.heading('#0', text='Χώρα (όπως στο Excel)', anchor='w')
        self.preview_tree.heading('p', text='Πελάτες')
        self.preview_tree.heading('n', text='Διανυκτ.')
        self.preview_tree.heading('s', text='Στο δελτίο', anchor='w')
        self.preview_tree.column('#0', width=220, minwidth=140, stretch=True)
        self.preview_tree.column('p', width=80, minwidth=60, anchor='e', stretch=False)
        self.preview_tree.column('n', width=80, minwidth=60, anchor='e', stretch=False)
        self.preview_tree.column('s', width=190, minwidth=120, stretch=False)
        self.preview_tree.tag_configure('bad', foreground=BAD, background='#fee2e2')
        self.preview_tree.tag_configure('ok', background=CARD)
        psb = ttk.Scrollbar(self.preview_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=psb.set)
        self.preview_tree.grid(row=1, column=0, sticky="ew")
        psb.grid(row=1, column=1, sticky="ns")
        self.preview_frame.grid_remove()
        self.period_var.set("Πατήστε «Επιλογή Excel…» και διαλέξτε το αρχείο «Nationalities Statistics» "
                            "που βγάζει το Protel (.xls ή .xlsx).")

        # ---- Step 2
        f2 = Card(body, 2, "Στοιχεία του μήνα")
        f2.grid(row=2, column=0, sticky="ew", padx=4, pady=6)
        c2 = f2.body
        c2.columnconfigure(2, weight=1)
        fields = [
            ("Έτος", 'year', 8, "αυτόματα από το Excel"),
            ("Μήνας (αριθμός)", 'month', 8, None),
            ("Τελευταία ημέρα («Έως»)", 'date_to', 14, "αυτόματα από το Excel (ΗΗ/ΜΜ/ΕΕΕΕ)"),
            ("Ημέρες λειτουργίας", 'days_of_month', 8, "αυτόματα από το Excel"),
            ("Σύνολο εργαζομένων", 'employees', 8,
             "✎ ΔΕΝ υπάρχει στο Excel - πληκτρολογήστε το εσείς (στο τέλος του μήνα, "
             "μαζί με ιδιοκτήτες/οικογένεια)" +
             (f"  ·  τελευταία φορά: {self.current_settings['last_employees']}"
              if self.current_settings.get('last_employees') else "")),
        ]
        for i, (lbl, key, width, hint) in enumerate(fields):
            ttk.Label(c2, text=lbl + ":", style='Card.TLabel').grid(row=i, column=0, sticky="e",
                                                                    padx=(0, 10), pady=5)
            e = ttk.Entry(c2, textvariable=self.mon_vars[key], width=width)
            e.grid(row=i, column=1, sticky="w", pady=5)
            self.entries[key] = e
            if key == 'month':
                ttk.Label(c2, textvariable=self.month_name_var, style='CardGood.TLabel',
                          font=(FONT, 10, 'bold')).grid(row=i, column=2, sticky="w", padx=12)
            else:
                h = ttk.Label(c2, text=hint, style='CardHint.TLabel', justify="left")
                h.grid(row=i, column=2, sticky="w", padx=12)
                autowrap(h, body, margin=430, minimum=140)

        # ---- Step 3
        f3 = Card(body, 3, "Πού να αποθηκευτεί το PDF")
        f3.grid(row=3, column=0, sticky="ew", padx=4, pady=(6, 14))
        c3 = f3.body
        c3.columnconfigure(0, weight=1)
        ttk.Entry(c3, textvariable=self.out_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(c3, text="Αλλαγή…", command=self.browse_output).grid(row=0, column=1, padx=(8, 0))
        h3 = ttk.Label(c3, text="Προτείνεται αυτόματα με βάση την καρτέλα «Ρυθμίσεις» (φάκελος και όνομα αρχείου).",
                       style='CardHint.TLabel', justify="left")
        h3.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        autowrap(h3, body, margin=110)

        # ---- bottom bar: always visible, never scrolls away
        bottom = ttk.Frame(tab, style='Card.TFrame')
        bottom.grid(row=1, column=0, sticky="ew")
        ttk.Separator(bottom, orient="horizontal").pack(fill="x")
        self.check_label = ttk.Label(bottom, textvariable=self.check_var, justify="left",
                                     style='StatusWarn.TLabel')
        self.check_label.pack(fill="x", padx=14, pady=(10, 6))
        autowrap(self.check_label, tab, margin=70)

        btnrow = ttk.Frame(bottom, style='Card.TFrame')
        btnrow.pack(fill="x", padx=14, pady=(4, 2))
        btnrow.columnconfigure(0, weight=1, uniform='b')
        btnrow.columnconfigure(1, weight=1, uniform='b')
        self.run_btn = ttk.Button(btnrow, text="Δημιουργία PDF", style='Primary.TButton',
                                  command=self.generate)
        self.run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.print_btn = ttk.Button(btnrow, text="Δημιουργία & Εκτύπωση", style='Print.TButton',
                                    command=lambda: self.generate(print_after=True))
        self.print_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(bottom, text="Συντομεύσεις:  Ctrl+O = επιλογή Excel  ·  Ctrl+Enter = δημιουργία  ·  "
                               "Ctrl+P = δημιουργία & εκτύπωση",
                  style='CardHint.TLabel').pack(anchor="e", padx=16, pady=(4, 10))

    def _build_hotel_tab(self):
        sf = ScrollFrame(self.tab_hotel)
        sf.pack(fill="both", expand=True)
        body = sf.body
        body.columnconfigure(0, weight=1)

        card = Card(body, None, "Στοιχεία ξενοδοχείου (αλλάζουν σπάνια)")
        card.grid(row=0, column=0, sticky="ew", padx=4, pady=(12, 6))
        f = card.body
        f.columnconfigure(1, weight=1)          # entries stretch with the window
        fields = [
            ("Είδος καταλύματος", 'type', "όπως θέλετε να γραφτεί στο δελτίο"),
            ("Τίτλος καταλύματος", 'name', "το όνομα όπως θα εμφανιστεί στο δελτίο"),
            ("Κατηγορία", 'category', "όπως θέλετε να γραφτεί στο δελτίο"),
            ("E-mail", 'email', None),
            ("Τηλέφωνο", 'phone', None),
            ("Α.Φ.Μ.", 'afm', "ακριβώς 9 ψηφία"),
            ("Επωνυμία επιχείρησης", 'company', None),
        ]
        r = 0
        for lbl, key, hint in fields:
            ttk.Label(f, text=lbl + ":", style='CardBold.TLabel').grid(row=r, column=0, sticky="e",
                                                                       padx=(0, 10), pady=(8, 0))
            e = ttk.Entry(f, textvariable=self.set_vars[key])
            e.grid(row=r, column=1, sticky="ew", pady=(8, 0))
            self.entries[key] = e
            if key == 'afm':
                vcmd = (self.register(lambda P: P == "" or (P.isascii() and P.isdigit() and len(P) <= 9)), '%P')
                e.config(validate='key', validatecommand=vcmd)
                ttk.Label(f, textvariable=self.afm_counter_var, style='CardBold.TLabel',
                          font=(FONT, 9, 'bold')).grid(row=r, column=2, sticky="w", padx=(10, 0), pady=(8, 0))
            if hint:
                h = ttk.Label(f, text=hint, style='CardHint.TLabel', justify="left")
                h.grid(row=r + 1, column=1, sticky="w", pady=(2, 0))
                autowrap(h, body, margin=290)
            r += 2

        if os.path.exists(SETTINGS_CSV):
            note = ("Τα στοιχεία διαβάστηκαν από το αρχείο hotel_settings.csv. Αν τα αλλάξετε, αποθηκεύονται "
                    "αυτόματα όταν δημιουργήσετε δελτίο.")
        else:
            note = ("Δεν βρέθηκε το αρχείο hotel_settings.csv. Συμπληρώστε τα στοιχεία εδώ - θα αποθηκευτούν "
                    "αυτόματα την πρώτη φορά που θα δημιουργήσετε δελτίο και δεν θα τα ξαναγράψετε.")
        nl = ttk.Label(body, text=note, style='Hint.TLabel', justify="left")
        nl.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 8))
        autowrap(nl, body, margin=50)

        card2 = Card(body, None, "Άδειο PDF της ΕΛΣΤΑΤ (template)")
        card2.grid(row=2, column=0, sticky="ew", padx=4, pady=(6, 14))
        f2 = card2.body
        f2.columnconfigure(0, weight=1)
        ttk.Entry(f2, textvariable=self.pdf_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(f2, text="Αναζήτηση…", command=self.browse_pdf).grid(row=0, column=1, padx=(8, 0))
        h2 = ttk.Label(f2, text="Το κενό δελτίο πάνω στο οποίο γράφονται τα στοιχεία. "
                                "Συνήθως δεν χρειάζεται να το αλλάξετε.", style='CardHint.TLabel', justify="left")
        h2.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        autowrap(h2, body, margin=110)

    # ------------------------------------------------------------ settings tab
    def _build_settings_tab(self):
        sf = ScrollFrame(self.tab_settings)
        sf.pack(fill="both", expand=True)
        body = sf.body
        body.columnconfigure(0, weight=1)
        pv = self.pv

        def hint(parent, text, row, columnspan=1, pady=(8, 0), margin=110, padx=0):
            lbl = ttk.Label(parent, text=text, style='CardHint.TLabel', justify="left")
            lbl.grid(row=row, column=0, columnspan=columnspan, sticky="w", pady=pady, padx=padx)
            autowrap(lbl, body, margin=margin)
            return lbl

        def commit_on(entry):
            entry.bind('<FocusOut>', self._pref_changed)
            entry.bind('<Return>', self._pref_changed)

        # ---- 1. where the PDFs go
        c1 = Card(body, None, "Πού αποθηκεύονται τα PDF")
        c1.grid(row=0, column=0, sticky="ew", padx=4, pady=(12, 6))
        f = c1.body
        f.columnconfigure(0, weight=1)
        for r, (text, value) in enumerate((("Στον φάκελο του προγράμματος", 'program'),
                                           ("Στον φάκελο του αρχείου Excel (από το Protel)", 'excel'),
                                           ("Σε συγκεκριμένο φάκελο:", 'custom'))):
            ttk.Radiobutton(f, text=text, value=value, variable=pv['output_mode'], style='Card.TRadiobutton',
                            command=self._pref_changed).grid(row=r, column=0, columnspan=2, sticky="w",
                                                             pady=(0 if r == 0 else 6, 0))
        self.out_dir_entry = ttk.Entry(f, textvariable=pv['output_dir'])
        self.out_dir_entry.grid(row=3, column=0, sticky="ew", padx=(24, 0), pady=(4, 0))
        commit_on(self.out_dir_entry)
        self.out_dir_btn = ttk.Button(f, text="Αναζήτηση…", command=self.browse_output_dir)
        self.out_dir_btn.grid(row=3, column=1, padx=(8, 0), pady=(4, 0))
        ttk.Separator(f, orient="horizontal").grid(row=4, column=0, columnspan=2, sticky="ew", pady=14)
        ttk.Label(f, text="Όνομα αρχείου PDF:", style='CardBold.TLabel').grid(row=5, column=0, columnspan=2,
                                                                              sticky="w")
        name_entry = ttk.Entry(f, textvariable=pv['filename_pattern'])
        name_entry.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        commit_on(name_entry)
        self.name_note = hint(f, "", 7, columnspan=2)
        self.name_note.configure(textvariable=self.name_note_var)
        self.out_note = hint(f, "", 8, columnspan=2, pady=(6, 0))
        self.out_note.configure(textvariable=self.out_note_var)

        # ---- 2. data folder
        c2 = Card(body, None, "Φάκελος δεδομένων")
        c2.grid(row=1, column=0, sticky="ew", padx=4, pady=6)
        f2 = c2.body
        f2.columnconfigure(0, weight=1)
        ttk.Entry(f2, textvariable=self.data_dir_var, state='readonly').grid(row=0, column=0, sticky="ew")
        ttk.Button(f2, text="Αλλαγή…", command=self.change_data_dir).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(f2, text="Άνοιγμα", command=lambda: open_file(DATA_DIR)).grid(row=0, column=2, padx=(6, 0))
        hint(f2, "Εδώ αποθηκεύονται τα στοιχεία του ξενοδοχείου, το ιστορικό, τα αρχειοθετημένα PDF και οι "
                 "ρυθμίσεις. Μπορείτε να διαλέξετε π.χ. φάκελο μέσα στο OneDrive ώστε να υπάρχει αυτόματο "
                 "αντίγραφο. Η θέση του φακέλου θυμάται το αρχείο elstat_config.json δίπλα στο πρόγραμμα.",
             1, columnspan=3)
        row2 = ttk.Frame(f2, style='Card.TFrame')
        row2.grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Button(row2, text="Προεπιλεγμένος φάκελος",
                   command=lambda: self._switch_data_dir(DEFAULT_DATA_DIR)).pack(side="left")
        ttk.Button(row2, text="Αντίγραφο ασφαλείας (.zip)…", command=self.backup_data).pack(side="left", padx=8)

        # ---- 3. Excel folder
        c3 = Card(body, None, "Αρχεία Excel από το Protel")
        c3.grid(row=2, column=0, sticky="ew", padx=4, pady=6)
        f3 = c3.body
        f3.columnconfigure(0, weight=1)
        ex_entry = ttk.Entry(f3, textvariable=pv['excel_dir'])
        ex_entry.grid(row=0, column=0, sticky="ew")
        commit_on(ex_entry)
        ttk.Button(f3, text="Αναζήτηση…", command=self.browse_excel_dir).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(f3, text="Καθαρισμός", command=self.clear_excel_dir).grid(row=0, column=2, padx=(6, 0))
        hint(f3, "Από ποιον φάκελο ξεκινούν τα «Επιλογή Excel…» και «Πιο πρόσφατο». "
                 "Κενό = ο φάκελος του τελευταίου Excel που χρησιμοποιήσατε.", 1, columnspan=3)

        # ---- 4. behaviour
        c4 = Card(body, None, "Συμπεριφορά του προγράμματος")
        c4.grid(row=3, column=0, sticky="ew", padx=4, pady=6)
        f4 = c4.body
        f4.columnconfigure(0, weight=1)
        behaviours = [
            ('open_after', "Άνοιγμα του PDF αμέσως μετά τη δημιουργία", None),
            ('confirm_overwrite', "Ερώτηση πριν αντικατασταθεί PDF που υπάρχει ήδη", None),
            ('keep_archive', "Αρχειοθέτηση αντιγράφου κάθε δελτίου στον φάκελο δεδομένων",
             "Επιτρέπει να ξανανοίξετε ή να εκτυπώσετε παλιούς μήνες από το Ιστορικό. Αν το κλείσετε, το "
             "δελτίο ξαναφτιάχνεται από τα δεδομένα του ιστορικού όταν το ζητήσετε."),
            ('show_reminder', "Υπενθύμιση όταν λείπει το δελτίο του προηγούμενου μήνα", None),
            ('prefill_employees', "Προσυμπλήρωση του «Σύνολο εργαζομένων» με την τελευταία τιμή",
             "Ελέγχετε πάντα τον αριθμό πριν δημιουργήσετε το δελτίο."),
        ]
        for r, (key, text, note) in enumerate(behaviours):
            ttk.Checkbutton(f4, text=text, variable=pv[key], style='Card.TCheckbutton',
                            command=self._pref_changed).grid(row=r * 2, column=0, sticky="w",
                                                             pady=(8 if r else 0, 0))
            if note:
                hint(f4, note, r * 2 + 1, padx=(24, 0), pady=(2, 0), margin=130)

        # ---- reset
        bottom = ttk.Frame(body)
        bottom.grid(row=4, column=0, sticky="ew", padx=4, pady=(6, 16))
        ttk.Button(bottom, text="Επαναφορά ρυθμίσεων", style='Danger.TButton',
                   command=self.reset_prefs).pack(side="left")
        ttk.Label(bottom, text="Δεν διαγράφει ιστορικό ή στοιχεία ξενοδοχείου.",
                  style='Hint.TLabel').pack(side="left", padx=10)

    # ---- settings: reading / saving
    def _save_prefs(self):
        try:
            save_prefs(self.prefs)
        except OSError as e:
            self._status(f"⚠ Δεν ήταν δυνατή η αποθήκευση των ρυθμίσεων: {e}")

    def _pref_changed(self, *_):
        p, pv = self.prefs, self.pv
        p['output_mode'] = pv['output_mode'].get()
        p['output_dir'] = pv['output_dir'].get().strip()
        p['excel_dir'] = pv['excel_dir'].get().strip()
        pattern = pv['filename_pattern'].get().strip()
        if pattern_error(pattern) is None:
            p['filename_pattern'] = pattern
        for k in BOOL_PREFS:
            p[k] = bool(pv[k].get())
        self._save_prefs()
        self._after_prefs_change()

    def _sync_pref_vars(self):
        for k, var in self.pv.items():
            var.set(self.prefs[k])

    def _after_prefs_change(self):
        custom = self.prefs['output_mode'] == 'custom'
        for w in (self.out_dir_entry, self.out_dir_btn):
            w.state(['!disabled' if custom else 'disabled'])
        self._update_name_note()
        self._refresh_output_suggestion()
        self.out_note_var.set("Στο Βήμα 3 του δελτίου προτείνεται: " + self._suggest_output() +
                              "  (μπορείτε πάντα να το αλλάξετε εκεί)")
        self.update_reminder()

    def _update_name_note(self, *_):
        pattern = self.pv['filename_pattern'].get().strip()
        err = pattern_error(pattern)
        if err:
            self.name_note_var.set("⚠ " + err + " (δεν αποθηκεύτηκε)")
            self.name_note.configure(style='CardBad.TLabel')
        else:
            today = datetime.date.today()
            self.name_note_var.set(f"Παράδειγμα: {build_pdf_name(pattern, today.year, today.month)}   ·   " +
                                   "{year} = έτος, {month} = μήνας (2 ψηφία)")
            self.name_note.configure(style='CardHint.TLabel')

    # ---- settings: where PDFs / Excel files go
    def _output_folder(self, excel=None):
        mode, folder = self.prefs['output_mode'], self.prefs['output_dir']
        if mode == 'excel':
            excel = excel or self.excel_var.get().strip()
            if excel and os.path.dirname(excel):
                return os.path.dirname(excel)
        elif mode == 'custom' and folder:
            return folder
        return BASE_DIR

    def _suggest_output(self, excel=None):
        year, month = self.mon_vars['year'].get().strip(), self.mon_vars['month'].get().strip()
        if year.isascii() and year.isdigit() and len(year) == 4 and month.isascii() and month.isdigit() \
                and 1 <= int(month) <= 12:
            name = build_pdf_name(self.prefs['filename_pattern'], year, int(month))
        else:
            name = "ELSTAT_Report.pdf"
        return os.path.join(self._output_folder(excel), name)

    def _refresh_output_suggestion(self, excel=None):
        if not self.output_custom:        # never overwrite a location the user picked by hand
            self.out_var.set(self._suggest_output(excel))

    def _excel_start_dir(self):
        d = self.prefs.get('excel_dir', '')
        return d if d and os.path.isdir(d) else self._start_dir()

    def _pdf_start_dir(self):
        d = self.prefs.get('output_dir', '')
        if self.prefs['output_mode'] == 'custom' and d and os.path.isdir(d):
            return d
        return self._start_dir()

    def _pick_folder(self, title, current):
        return filedialog.askdirectory(title=title, initialdir=current if current and os.path.isdir(current)
                                       else BASE_DIR)

    def browse_output_dir(self):
        d = self._pick_folder("Διαλέξτε φάκελο για τα PDF", self.pv['output_dir'].get().strip())
        if d:
            self.pv['output_dir'].set(os.path.normpath(d))
            self.pv['output_mode'].set('custom')
            self._pref_changed()

    def browse_excel_dir(self):
        d = self._pick_folder("Διαλέξτε φάκελο των Excel", self.pv['excel_dir'].get().strip())
        if d:
            self.pv['excel_dir'].set(os.path.normpath(d))
            self._pref_changed()

    def clear_excel_dir(self):
        self.pv['excel_dir'].set('')
        self._pref_changed()

    def _remember_template(self, path):
        self.prefs['template_pdf'] = '' if _same_path(path, DEFAULT_TEMPLATE) else path
        self._save_prefs()

    # ---- settings: data folder
    def change_data_dir(self):
        d = self._pick_folder("Διαλέξτε τον φάκελο δεδομένων", DATA_DIR)
        if d:
            self._switch_data_dir(os.path.normpath(d))

    def _switch_data_dir(self, new):
        old = DATA_DIR
        if _same_path(new, old):
            return
        if _is_inside(new, os.path.join(old, ARCHIVE_NAME)):
            messagebox.showerror("Φάκελος δεδομένων", "Δεν μπορείτε να διαλέξετε φάκελο μέσα στο αρχείο PDF.")
            return
        copy = False
        if data_exists(new):
            if not messagebox.askyesno(
                    "Φάκελος δεδομένων",
                    f"Ο φάκελος\n{new}\nπεριέχει ήδη δεδομένα του προγράμματος.\n\n"
                    "Να χρησιμοποιηθούν αυτά; (Τα τωρινά δεδομένα δεν διαγράφονται.)", default='yes'):
                return
        elif data_exists(old):
            ans = messagebox.askyesnocancel(
                "Φάκελος δεδομένων",
                "Να αντιγραφούν τα υπάρχοντα δεδομένα (στοιχεία ξενοδοχείου, ιστορικό, αρχείο PDF) στον "
                "νέο φάκελο;\n\nΝαι = αντιγραφή (τα παλιά αρχεία μένουν και στον παλιό φάκελο)\n"
                "Όχι = ξεκίνημα από το μηδέν στον νέο φάκελο\nΑκύρωση = καμία αλλαγή", default='yes')
            if ans is None:
                return
            copy = ans
        try:
            os.makedirs(new, exist_ok=True)
            if copy:
                copy_data_files(old, new)
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Δεν ήταν δυνατή η αντιγραφή στον νέο φάκελο:\n{e}")
            return
        set_data_dir(new)
        if not save_app_config({'data_dir': '' if _same_path(new, DEFAULT_DATA_DIR) else new}):
            set_data_dir(old)
            messagebox.showerror("Σφάλμα", "Δεν μπορώ να θυμηθώ τη νέα θέση, γιατί ο φάκελος του προγράμματος "
                                           "δεν επιτρέπει εγγραφή. Η αλλαγή ακυρώθηκε.")
            return
        # preferences live in the data folder: use the ones found there, or carry the current ones over
        if os.path.exists(PREFS_JSON):
            self.prefs = load_prefs()
            self._sync_pref_vars()
        else:
            self._save_prefs()
        if os.path.exists(SETTINGS_CSV):
            self.current_settings = load_hotel_settings()
            for k in SETTING_KEYS:
                self.set_vars[k].set(self.current_settings.get(k, ''))
        self.data_dir_var.set(DATA_DIR)
        self.refresh_history()
        self._after_prefs_change()
        self._status(f"✓ Ο φάκελος δεδομένων είναι τώρα: {DATA_DIR}")

    def backup_data(self):
        name = f"ELSTAT_backup_{datetime.date.today().strftime('%Y-%m-%d')}.zip"
        fp = filedialog.asksaveasfilename(title="Αντίγραφο ασφαλείας", defaultextension=".zip",
                                          initialdir=self._start_dir(), initialfile=name,
                                          filetypes=[("ZIP", "*.zip")])
        if not fp:
            return
        try:
            ensure_data_dir()
            with tempfile.TemporaryDirectory() as tmp:
                z = shutil.make_archive(os.path.join(tmp, 'backup'), 'zip', DATA_DIR)
                shutil.copy2(z, fp)
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Δεν ήταν δυνατή η δημιουργία αντιγράφου:\n{e}")
            return
        self._status(f"✓ Αντίγραφο ασφαλείας: {fp}")

    def reset_prefs(self):
        if not messagebox.askyesno("Επαναφορά ρυθμίσεων",
                                   "Να γυρίσουν όλες οι ρυθμίσεις στις προεπιλογές;\n"
                                   "(Δεν διαγράφονται δεδομένα ή ιστορικό.)", default='no'):
            return
        keep_template = self.prefs.get('template_pdf', '')
        self.prefs = dict(DEFAULT_PREFS, template_pdf=keep_template)
        self._sync_pref_vars()
        self._save_prefs()
        self._after_prefs_change()
        self._status("Οι ρυθμίσεις επανήλθαν στις προεπιλογές.")

    # ---------------------------------------------------- window behaviour
    def _center(self):
        try:
            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            x = max((self.winfo_screenwidth() - w) // 2, 0)
            y = max((self.winfo_screenheight() - h) // 3, 0)
            self.geometry(f"+{x}+{y}")
        except (tk.TclError, TypeError):
            pass

    def _on_wheel(self, event):
        """Mouse wheel scrolls whichever scrolling area is under the pointer."""
        try:
            w = self.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return
        if isinstance(w, ttk.Treeview):
            return
        while w is not None:
            sf = getattr(w, '_sf', None)
            if sf is not None:
                if getattr(event, 'num', None) == 4:
                    units = -3
                elif getattr(event, 'num', None) == 5:
                    units = 3
                else:
                    units = -1 if event.delta > 0 else 1
                    units *= 3
                sf.scroll(units)
                return
            w = getattr(w, 'master', None)

    def _try_generate(self, _event=None, print_after=False):
        if not self._busy and self.run_btn.instate(['!disabled']):
            self.generate(print_after=print_after)

    # ------------------------------------------------------------ live checks
    def _collect(self):
        s = {k: v.get().strip() for k, v in self.set_vars.items()}
        m = {k: v.get().strip() for k, v in self.mon_vars.items()}
        return s, m, self.excel_var.get().strip(), self.pdf_var.get().strip(), self.out_var.get().strip()

    @staticmethod
    def _field_states(s, m):
        st = {'afm': validate_afm(s['afm'])}
        for k in ('type', 'name', 'category', 'company'):
            st[k] = bool(s[k])
        st['year'] = m['year'].isascii() and m['year'].isdigit() and len(m['year']) == 4
        st['month'] = m['month'].isascii() and m['month'].isdigit() and 1 <= int(m['month']) <= 12
        try:
            datetime.datetime.strptime(m['date_to'], "%d/%m/%Y")
            st['date_to'] = True
        except ValueError:
            st['date_to'] = False
        d = m['days_of_month']
        st['days_of_month'] = d.isascii() and d.isdigit() and 1 <= int(d) <= 31
        e = m['employees']
        st['employees'] = e.isascii() and e.isdigit() and len(e) <= 3
        return st

    def refresh_validation(self, *_):
        if not hasattr(self, 'print_btn'):
            return
        s, m, excel, template, out = self._collect()

        # red / normal field colours
        excel_ok = bool(excel) and os.path.isfile(excel)
        for key, ok in self._field_states(s, m).items():
            if key in self.entries:
                quiet = key in m and not m[key] and not excel_ok    # empty month fields before an Excel is chosen
                self.entries[key].configure(style='TEntry' if (ok or quiet) else 'Bad.TEntry')

        # helpers next to the fields
        n = len(s['afm'])
        self.afm_counter_var.set(f"{n}/9 ψηφία" + ("  ✓" if validate_afm(s['afm']) else ""))
        if m['month'].isdigit() and 1 <= int(m['month']) <= 12 and m['year']:
            self.month_name_var.set(f"= {GREEK_MONTHS[int(m['month']) - 1]} {m['year']}")
        else:
            self.month_name_var.set("")

        # overall verdict
        if excel_ok:
            errors = validate_inputs(s, m, excel, template, out)
        else:   # month fields cannot be judged before the Excel is chosen
            errors = ["Επιλέξτε πρώτα το αρχείο Excel από το Protel (Βήμα 1)."] + validate_hotel(s)
        if errors:
            shown = errors[:5]
            more = f"\n… και {len(errors) - 5} ακόμη" if len(errors) > 5 else ""
            self.check_var.set("Για να δημιουργηθεί το δελτίο, διορθώστε:\n" + "\n".join("• " + e for e in shown) + more)
            self.check_label.configure(style='StatusWarn.TLabel' if not excel_ok else 'StatusBad.TLabel')
        else:
            self.check_var.set("✓ Όλα είναι έτοιμα. Πατήστε «Δημιουργία PDF» ή «Δημιουργία & Εκτύπωση».")
            self.check_label.configure(style='StatusGood.TLabel')

        for btn in (self.run_btn, self.print_btn):
            if errors or self._busy:
                btn.state(['disabled'])
            else:
                btn.state(['!disabled'])

    def _set_busy(self, busy):
        self._busy = busy
        self.config(cursor='watch' if busy else '')
        self.run_btn.configure(text="Παρακαλώ περιμένετε…" if busy else "Δημιουργία PDF")
        self.print_btn.configure(text="Παρακαλώ περιμένετε…" if busy else "Δημιουργία & Εκτύπωση")
        self.refresh_validation()
        self.update_idletasks()

    # ----------------------------------------------------------- file choices
    def _start_dir(self):
        d = self.current_settings.get('last_dir', '')
        return d if d and os.path.isdir(d) else BASE_DIR

    def browse_excel(self):
        fp = filedialog.askopenfilename(title="Διαλέξτε το Excel από το Protel", initialdir=self._excel_start_dir(),
                                        filetypes=[("Excel Files", "*.xls *.xlsx")])
        if fp:
            self.excel_var.set(fp)
            self.load_period(fp)

    def _excel_typed(self, _event=None):
        fp = self.excel_var.get().strip()
        if fp and fp != self._loaded_excel:
            self.load_period(fp)

    def load_period(self, fp):
        """Reads the month from the Excel header and fills in step 2 and the output name."""
        self._loaded_excel = fp
        if not os.path.isfile(fp):
            self.period_var.set("⚠ Το αρχείο δεν βρέθηκε.")
            self.period_label.configure(style='CardBad.TLabel')
            return
        self.load_preview(fp)
        info = period_as_month(read_excel_period(fp))
        if info:
            year, month, last = info
            self.mon_vars['year'].set(str(year))
            self.mon_vars['month'].set(str(month))
            self.mon_vars['date_to'].set(f"{last:02d}/{month:02d}/{year}")
            self.mon_vars['days_of_month'].set(str(last))
            self._refresh_output_suggestion(excel=fp)
            last_emp = self.current_settings.get('last_employees', '')
            if self.prefs['prefill_employees'] and last_emp and not self.mon_vars['employees'].get().strip():
                self.mon_vars['employees'].set(last_emp)
            self.period_var.set(f"✓ Το Excel αφορά τον μήνα {GREEK_MONTHS[month - 1]} {year}. "
                                "Τα στοιχεία του μήνα συμπληρώθηκαν αυτόματα - πληκτρολογήστε μόνο τον αριθμό εργαζομένων.")
            self.period_label.configure(style='CardGood.TLabel')
        else:
            self.period_var.set("⚠ Δεν βρέθηκε ολόκληρος μήνας στο Excel. Συμπληρώστε τα στοιχεία του Βήματος 2 "
                                "με το χέρι.")
            self.period_label.configure(style='CardWarn.TLabel')

    def use_recent_excel(self):
        """Picks the newest Excel in the folder used last time (or next to the program)."""
        fp = find_recent_excel(self._excel_start_dir())
        if not fp:
            messagebox.showinfo("Δεν βρέθηκε αρχείο",
                                f"Δεν υπάρχει αρχείο Excel στον φάκελο:\n{self._excel_start_dir()}\n\n"
                                "Πατήστε «Επιλογή Excel…» για να το διαλέξετε.")
            return
        self.excel_var.set(fp)
        self.load_period(fp)

    def load_preview(self, fp):
        """Shows the countries found in the Excel and flags the ones the program does not know."""
        self.preview_frame.grid_remove()
        try:
            rows, _total = read_nationality_rows(fp)
        except Exception:
            return
        if not rows:
            return
        self.preview_tree.delete(*self.preview_tree.get_children())
        bad = 0
        for country, pax, nights in rows:
            known = resolve_country(country) is not None
            bad += not known
            self.preview_tree.insert('', 'end', text=country, values=(pax, nights,
                                     "✓" if known else "⚠ λείπει από το πρόγραμμα"),
                                     tags=('ok' if known else 'bad',))
        self.preview_tree.configure(height=max(3, min(len(rows), 7)))
        text = (f"Βρέθηκαν {len(rows)} χώρες · {sum(r[1] for r in rows)} πελάτες · "
                f"{sum(r[2] for r in rows)} διανυκτερεύσεις")
        if bad:
            text += f"   ⚠ {bad} δεν αναγνωρίστηκαν και ΔΕΝ θα μπουν στο PDF"
            self.preview_summary.configure(style='CardBad.TLabel')
        else:
            self.preview_summary.configure(style='CardGood.TLabel')
        self.preview_summary.configure(text=text)
        self.preview_frame.grid()

    def _build_history_tab(self):
        tab = self.tab_history
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(4, weight=1)

        top = ttk.Frame(tab)
        top.grid(row=0, column=0, sticky="ew", padx=6, pady=(12, 2))
        ttk.Label(top, text="Έτος:").pack(side="left")
        self.hist_year_cb = ttk.Combobox(top, textvariable=self.hist_year_var, state="readonly", width=7)
        self.hist_year_cb.pack(side="left", padx=(6, 16))
        self.hist_year_cb.bind('<<ComboboxSelected>>', lambda e: self.update_dashboard())
        ttk.Radiobutton(top, text="Διανυκτερεύσεις", value='nights', variable=self.hist_metric_var,
                        command=self.update_dashboard).pack(side="left")
        ttk.Radiobutton(top, text="Πελάτες", value='clients', variable=self.hist_metric_var,
                        command=self.update_dashboard).pack(side="left", padx=(10, 0))
        ttk.Button(top, text="Άνοιγμα CSV",
                   command=lambda: os.path.exists(HISTORY_CSV) and open_file(HISTORY_CSV)).pack(side="right")
        ttk.Button(top, text="Ανανέωση", command=self.refresh_history).pack(side="right", padx=(0, 8))

        self.hist_info = tk.StringVar()
        ttk.Label(tab, textvariable=self.hist_info, style='Hint.TLabel').grid(row=1, column=0, sticky="w",
                                                                              padx=8, pady=(0, 8))
        kp = ttk.Frame(tab)
        kp.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 8))
        for i in range(3):
            kp.columnconfigure(i, weight=1, uniform='kpi')
        self.kpi_clients = StatCard(kp, "ΠΕΛΑΤΕΣ")
        self.kpi_clients.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.kpi_nights = StatCard(kp, "ΔΙΑΝΥΚΤΕΡΕΥΣΕΙΣ")
        self.kpi_nights.grid(row=0, column=1, sticky="ew", padx=3)
        self.kpi_avg = StatCard(kp, "ΜΕΣΗ ΔΙΑΜΟΝΗ (ΝΥΧΤΕΣ)")
        self.kpi_avg.grid(row=0, column=2, sticky="ew", padx=(6, 0))

        self.hist_chart = HistoryChart(tab, on_pick=self._chart_pick, height=170)
        self.hist_chart.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 8))

        frame = ttk.Frame(tab)
        frame.grid(row=4, column=0, sticky="nsew", padx=6, pady=(0, 8))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        self.hist_tree = ttk.Treeview(frame, columns=('c', 'n', 'a', 'p'), show='tree headings', height=6)
        self.hist_tree.heading('#0', text='Έτος / Μήνας / Χώρα', anchor='w')
        self.hist_tree.heading('c', text='Πελάτες')
        self.hist_tree.heading('n', text='Διανυκτερεύσεις')
        self.hist_tree.heading('a', text='Μέση διαμονή')
        self.hist_tree.heading('p', text='PDF')
        self.hist_tree.column('#0', width=260, minwidth=170, stretch=True)
        self.hist_tree.column('c', width=90, minwidth=70, anchor='e', stretch=False)
        self.hist_tree.column('n', width=120, minwidth=100, anchor='e', stretch=False)
        self.hist_tree.column('a', width=110, minwidth=90, anchor='e', stretch=False)
        self.hist_tree.column('p', width=60, minwidth=50, anchor='center', stretch=False)
        self.hist_tree.tag_configure('year', font=(FONT, 10, 'bold'), background='#dbe7fb')
        self.hist_tree.tag_configure('month', font=(FONT, 10, 'bold'), background='#eef4fd')
        self.hist_tree.tag_configure('country', background=CARD)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.hist_tree.yview)
        self.hist_tree.configure(yscrollcommand=sb.set)
        self.hist_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.hist_tree.bind('<<TreeviewSelect>>', self._on_hist_select)
        self.hist_tree.bind('<Return>', lambda e: self.hist_open())

        act = ttk.Frame(tab)
        act.grid(row=5, column=0, sticky="ew", padx=6, pady=(0, 12))
        self.month_btns, self.year_btns = [], []
        ttk.Label(act, text="Δελτίο μήνα:", style='Hint.TLabel').pack(side="left", padx=(0, 6))
        for text, cmd, style in (("Άνοιγμα", self.hist_open, 'TButton'),
                                 ("Αποθήκευση ως…", self.hist_save_as, 'TButton'),
                                 ("Εκτύπωση", self.hist_print, 'PrintSm.TButton')):
            b = ttk.Button(act, text=text, command=cmd, style=style)
            b.pack(side="left", padx=(0, 6))
            self.month_btns.append(b)
        ttk.Separator(act, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Label(act, text="Σύνοψη έτους:", style='Hint.TLabel').pack(side="left", padx=(0, 6))
        for text, cmd, style in (("Αποθήκευση PDF…", lambda: self.hist_year_pdf('save'), 'TButton'),
                                 ("Εκτύπωση", lambda: self.hist_year_pdf('print'), 'PrintSm.TButton')):
            b = ttk.Button(act, text=text, command=cmd, style=style)
            b.pack(side="left", padx=(0, 6))
            self.year_btns.append(b)
        self.del_btn = ttk.Button(act, text="Διαγραφή μήνα", command=self.hist_delete, style='Danger.TButton')
        self.del_btn.pack(side="right")

        # right-click menu
        self.hist_menu = tk.Menu(self, tearoff=0)
        self.hist_menu.add_command(label="Άνοιγμα PDF", command=self.hist_open)
        self.hist_menu.add_command(label="Αποθήκευση ως…", command=self.hist_save_as)
        self.hist_menu.add_command(label="Εκτύπωση δελτίου", command=self.hist_print)
        self.hist_menu.add_separator()
        self.hist_menu.add_command(label="Σύνοψη έτους - αποθήκευση PDF…", command=lambda: self.hist_year_pdf('save'))
        self.hist_menu.add_command(label="Σύνοψη έτους - εκτύπωση", command=lambda: self.hist_year_pdf('print'))
        self.hist_menu.add_separator()
        self.hist_menu.add_command(label="Διαγραφή μήνα", command=self.hist_delete)
        self.hist_tree.bind('<Button-3>', self._hist_popup)
        if sys.platform == 'darwin':
            self.hist_tree.bind('<Button-2>', self._hist_popup)

    # ---- history: helpers
    @staticmethod
    def _parse_iid(iid):
        nums = [int(x) for x in iid[1:].split('-')]
        return (nums[0], None) if iid[0] == 'y' else (nums[0], nums[1])

    def _hist_sel(self):
        cur = self.hist_tree.selection()
        return self._parse_iid(cur[0]) if cur else None

    def _status(self, text):
        self.status_var.set(text)

    def _hist_popup(self, event):
        iid = self.hist_tree.identify_row(event.y)
        if iid:
            self.hist_tree.selection_set(iid)
        try:
            self.hist_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.hist_menu.grab_release()

    def _chart_pick(self, month):
        try:
            year = int(self.hist_year_var.get())
        except ValueError:
            return
        iid = f"m{year}-{month}"
        if self.hist_tree.exists(iid):
            self.hist_tree.item(f"y{year}", open=True)
            self.hist_tree.selection_set(iid)
            self.hist_tree.see(iid)

    def _on_hist_select(self, _event=None):
        sel = self._hist_sel()
        if sel and str(sel[0]) != self.hist_year_var.get():
            self.hist_year_var.set(str(sel[0]))
        self.update_dashboard()
        self._update_hist_buttons()

    def _update_hist_buttons(self):
        sel = self._hist_sel()
        month_ok = bool(sel and sel[1] is not None)
        year_ok = bool(self._hist)
        for b in self.month_btns + [self.del_btn]:
            b.state(['!disabled'] if month_ok else ['disabled'])
        for b in self.year_btns:
            b.state(['!disabled'] if year_ok else ['disabled'])

    # ---- history: fill the tab
    def refresh_history(self):
        """Fills the history tab from elstat_history.csv (year > month > country) and redraws the dashboard."""
        tree = self.hist_tree
        keep_sel = tree.selection()[0] if tree.selection() else None

        def walk(parent=''):
            for i in tree.get_children(parent):
                yield i
                yield from walk(i)
        was_open = {i for i in walk() if tree.item(i, 'open')}
        first_time = not tree.get_children()

        hist = self._hist = load_history()
        tree.delete(*tree.get_children())

        def avg(c, n):
            return f"{n / c:.1f}".replace('.', ',') if c else "-"

        months_total = 0
        for year in sorted(hist, reverse=True):
            ycl = sum(c for rows in hist[year].values() for _, c, _ in rows)
            yni = sum(n for rows in hist[year].values() for _, _, n in rows)
            yid = f"y{year}"
            tree.insert('', 'end', iid=yid, text=str(year), values=(ycl, yni, avg(ycl, yni), ""),
                        tags=('year',), open=(year == max(hist)) if first_time else (yid in was_open))
            for month in sorted(hist[year], reverse=True):
                rows = hist[year][month]
                mcl, mni = sum(r[1] for r in rows), sum(r[2] for r in rows)
                months_total += 1
                mid = f"m{year}-{month}"
                name = GREEK_MONTHS[month - 1] if 1 <= month <= 12 else str(month)
                tree.insert(yid, 'end', iid=mid, text=name,
                            values=(mcl, mni, avg(mcl, mni), "✓" if os.path.isfile(archive_path(year, month)) else ""),
                            tags=('month',), open=(mid in was_open))
                for i, (country, c, n) in enumerate(sorted(rows, key=lambda r: -r[1])):
                    tree.insert(mid, 'end', iid=f"c{year}-{month}-{i}", text=country,
                                values=(c, n, avg(c, n), ""), tags=('country',))
        if keep_sel and tree.exists(keep_sel):
            tree.selection_set(keep_sel)

        years = [str(y) for y in sorted(hist, reverse=True)]
        self.hist_year_cb.configure(values=years)
        if self.hist_year_var.get() not in years:
            self.hist_year_var.set(years[0] if years else "")
        self._hist_info_base = (f"{months_total} μήνες καταχωρημένοι. Κάντε κλικ σε μήνα (πίνακας ή γράφημα) "
                                "για να τον ανοίξετε, αποθηκεύσετε ή εκτυπώσετε." if hist else
                                "Δεν υπάρχει ιστορικό ακόμη - δημιουργείται αυτόματα με το πρώτο δελτίο.")
        self.hist_info.set(self._hist_info_base)
        self.update_dashboard()
        self._update_hist_buttons()
        self.update_reminder()

    def update_dashboard(self, *_):
        """Cards + bar chart for the chosen year."""
        try:
            year = int(self.hist_year_var.get())
        except ValueError:
            year = None
        by_nights = self.hist_metric_var.get() == 'nights'
        if year is None or year not in self._hist:
            for card in (self.kpi_clients, self.kpi_nights, self.kpi_avg):
                card.set("–", "")
            self.hist_chart.set_data([0] * 12, "Δεν υπάρχουν δεδομένα ακόμη")
            return
        st = year_stats(self._hist, year)

        def delta(cur, prev):
            if not prev:
                return "", MUTED
            pct = (cur - prev) / prev * 100
            arrow = "▲" if pct > 0.05 else ("▼" if pct < -0.05 else "■")
            text = f"{arrow} {pct:+.1f}% έναντι {year - 1} (ίδιοι μήνες)".replace('.', ',')
            return text, (GOOD if pct > 0.05 else BAD if pct < -0.05 else MUTED)

        cmp_ = st['cmp']
        t, col = delta(cmp_['cur_cl'], cmp_['prev_cl']) if cmp_ else ("", MUTED)
        self.kpi_clients.set(fmt_int(st['clients']), t or f"{len(st['months'])} μήνες με δελτίο", col)
        t, col = delta(cmp_['cur_ni'], cmp_['prev_ni']) if cmp_ else ("", MUTED)
        self.kpi_nights.set(fmt_int(st['nights']), t or "συνολικά στο έτος", col)
        self.kpi_avg.set(fmt_avg(st['nights'], st['clients']), "διανυκτερεύσεις ανά πελάτη")
        gap_txt = ("   ·   Μήνες χωρίς δελτίο μέσα στο " + str(year) + ": " +
                   ", ".join(GREEK_MONTHS_SHORT[m - 1] for m in st['gaps'])) if st['gaps'] else ""
        self.hist_info.set(getattr(self, '_hist_info_base', "") + gap_txt)
        per_month = [sum((r[2] if by_nights else r[1]) for r in self._hist[year].get(m, [])) for m in range(1, 13)]
        sel = self._hist_sel()
        self.hist_chart.set_data(
            per_month, f"{'Διανυκτερεύσεις' if by_nights else 'Πελάτες'} ανά μήνα · {year}",
            color=ACCENT if by_nights else TEAL, selected=sel[1] if sel and sel[0] == year else None)

    # ---- reminder banner
    def update_reminder(self):
        r = None if (self._reminder_dismissed or not self.prefs['show_reminder']) \
            else missing_month_reminder(self._hist)
        if r:
            self.reminder_label.configure(
                text=f"Υπενθύμιση: δεν υπάρχει ακόμη δελτίο για τον {GREEK_MONTHS[r[1] - 1]} {r[0]}. "
                     "Διαλέξτε το Excel του μήνα στο Βήμα 1 για να το δημιουργήσετε.")
            self.reminder_frame.grid()
        else:
            self.reminder_frame.grid_remove()

    def _dismiss_reminder(self):
        self._reminder_dismissed = True
        self.update_reminder()

    # ---- history: actions on a month
    def _ensure_month_pdf(self, year, month):
        """Path of the archived PDF of a month. If it does not exist (month made before the archive
        existed) the form is rebuilt from the history data and archived."""
        path = archive_path(year, month)
        if os.path.isfile(path):
            return path
        rows = self._hist.get(year, {}).get(month)
        if not rows:
            messagebox.showerror("Ιστορικό", "Δεν βρέθηκαν στοιχεία για αυτόν τον μήνα.")
            return None
        s, _m, _excel, template, _out = self._collect()
        errors = validate_hotel(s)
        if not template or not os.path.isfile(template):
            errors.append("Το άδειο PDF (template) δεν βρέθηκε - δείτε την καρτέλα «Στοιχεία ξενοδοχείου».")
        if errors:
            messagebox.showerror("Δεν μπορώ να ξαναφτιάξω το δελτίο", "\n".join("• " + e for e in errors))
            return None
        rec = load_month_records().get((year, month), {})
        employees = rec.get('employees', '')
        name = f"{GREEK_MONTHS[month - 1]} {year}"
        if not employees:
            employees = simpledialog.askstring(
                "Σύνολο εργαζομένων",
                f"Το δελτίο του {name} δημιουργήθηκε πριν αποθηκεύεται ο αριθμός εργαζομένων.\n"
                "Πληκτρολογήστε το Σύνολο Εργαζομένων εκείνου του μήνα:", parent=self)
            if employees is None:
                return None
            employees = employees.strip()
        last = calendar.monthrange(year, month)[1]
        mon = dict(year=str(year), month=str(month), date_to=f"{last:02d}/{month:02d}/{year}",
                   days_of_month=str(rec.get('days') or last), employees=employees)
        problems = validate_month(mon)
        if problems:
            messagebox.showerror("Λάθος στοιχεία", "\n".join("• " + e for e in problems))
            return None
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            process_elstat_report(None, template, path, s, mon, rows_override=rows, save_history=False)
            save_month_record(year, month, employees, mon['days_of_month'], mon['date_to'])
        except PermissionError:
            messagebox.showerror("Σφάλμα", "Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό και δοκιμάστε ξανά.")
            return None
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Κάτι πήγε στραβά:\n{e}")
            return None
        finally:
            self.config(cursor='')
        self._status(f"Το δελτίο {name} ξαναδημιουργήθηκε από το ιστορικό.")
        self.refresh_history()
        return path

    def _month_from_selection(self):
        sel = self._hist_sel()
        return sel if sel and sel[1] is not None else None

    def hist_open(self):
        sel = self._month_from_selection()
        path = sel and self._ensure_month_pdf(*sel)
        if path:
            open_file(path)
            self._status(f"Άνοιξε το δελτίο {GREEK_MONTHS[sel[1] - 1]} {sel[0]}.")

    def hist_save_as(self):
        sel = self._month_from_selection()
        path = sel and self._ensure_month_pdf(*sel)
        if not path:
            return
        fp = filedialog.asksaveasfilename(title="Αποθήκευση δελτίου ως", defaultextension=".pdf",
                                          initialdir=self._pdf_start_dir(),
                                          initialfile=os.path.basename(path), filetypes=[("PDF Files", "*.pdf")])
        if not fp:
            return
        try:
            shutil.copy2(path, fp)
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Δεν ήταν δυνατή η αποθήκευση:\n{e}")
            return
        self._status(f"✓ Αποθηκεύτηκε: {fp}")

    def hist_print(self):
        sel = self._month_from_selection()
        path = sel and self._ensure_month_pdf(*sel)
        if path:
            self._status(print_file(path)[1])

    def hist_delete(self):
        sel = self._month_from_selection()
        if not sel:
            return
        name = f"{GREEK_MONTHS[sel[1] - 1]} {sel[0]}"
        if not messagebox.askyesno("Διαγραφή μήνα",
                                   f"Να διαγραφεί ο μήνας {name} από το ιστορικό;\n"
                                   "(Θα διαγραφεί και το αρχειοθετημένο PDF του. Το αρχείο που αποθηκεύσατε "
                                   "εσείς δεν θα πειραχτεί.)", default='no'):
            return
        try:
            delete_month(*sel)
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Δεν ήταν δυνατή η διαγραφή:\n{e}")
            return
        self._status(f"Ο μήνας {name} διαγράφηκε από το ιστορικό.")
        self.refresh_history()

    def hist_year_pdf(self, mode):
        """One-page summary of the year as PDF: mode 'save' (choose a file) or 'print'."""
        sel = self._hist_sel()
        try:
            year = sel[0] if sel else int(self.hist_year_var.get())
        except ValueError:
            return
        if year not in self._hist:
            return
        s = {k: v.get().strip() for k, v in self.set_vars.items()}
        name = f"ELSTAT_Synopsi_{year}.pdf"
        if mode == 'save':
            fp = filedialog.asksaveasfilename(title="Αποθήκευση σύνοψης έτους", defaultextension=".pdf",
                                              initialdir=self._pdf_start_dir(), initialfile=name,
                                              filetypes=[("PDF Files", "*.pdf")])
            if not fp:
                return
        else:
            os.makedirs(ARCHIVE_DIR, exist_ok=True)
            fp = os.path.join(ARCHIVE_DIR, name)
        try:
            make_year_summary_pdf(year, self._hist, s, fp)
        except PermissionError:
            messagebox.showerror("Σφάλμα", "Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό και δοκιμάστε ξανά.")
            return
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Κάτι πήγε στραβά:\n{e}")
            return
        if mode == 'save':
            self._status(f"✓ Η σύνοψη του {year} αποθηκεύτηκε: {fp}")
            open_file(fp)
        else:
            self._status(print_file(fp)[1])

    def browse_pdf(self):
        fp = filedialog.askopenfilename(title="Διαλέξτε το άδειο PDF της ΕΛΣΤΑΤ", initialdir=BASE_DIR,
                                        filetypes=[("PDF Files", "*.pdf")])
        if fp:
            self.pdf_var.set(fp)
            self._remember_template(fp)

    def browse_output(self):
        cur = self.out_var.get().strip()
        fp = filedialog.asksaveasfilename(title="Αποθήκευση PDF ως", defaultextension=".pdf",
                                          initialdir=os.path.dirname(cur) or BASE_DIR,
                                          initialfile=os.path.basename(cur) or "ELSTAT_Report.pdf",
                                          filetypes=[("PDF Files", "*.pdf")])
        if fp:
            self.out_var.set(fp)
            self.output_custom = True

    # --------------------------------------------------------------- generate
    def generate(self, print_after=False):
        s, m, excel, template, out = self._collect()
        errors = validate_inputs(s, m, excel, template, out)
        if errors:
            messagebox.showerror("Λάθος στοιχεία", "\n".join("• " + e for e in errors))
            return

        info = period_as_month(read_excel_period(excel))
        if info and (str(info[0]), str(int(m['month']))) != (m['year'], str(info[1])):
            if not messagebox.askyesno(
                    "Διαφορετική περίοδος",
                    f"Το Excel αφορά τον μήνα {GREEK_MONTHS[info[1] - 1]} {info[0]}, ενώ στη φόρμα έχετε "
                    f"{GREEK_MONTHS[int(m['month']) - 1]} {m['year']}.\n\nΝα συνεχίσω παρ' όλα αυτά;",
                    default='no'):
                return

        if self.prefs['confirm_overwrite'] and os.path.exists(out):
            if not messagebox.askyesno("Υπάρχει ήδη αρχείο",
                                       f"Το αρχείο υπάρχει ήδη:\n{out}\n\nΝα αντικατασταθεί;", default='no'):
                return

        saved = dict(s, last_dir=os.path.dirname(excel), last_employees=m['employees'])
        try:
            save_hotel_settings(saved)
        except Exception as e:
            messagebox.showwarning("Προσοχή", f"Δεν ήταν δυνατή η αποθήκευση των στοιχείων στον φάκελο "
                                              f"δεδομένων:\n{DATA_DIR}\n\n{e}")
        self.current_settings = saved
        if template != (self.prefs.get('template_pdf') or DEFAULT_TEMPLATE):
            self._remember_template(template)

        self._set_busy(True)
        try:
            os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
            result = process_elstat_report(excel, template, out, s, m)
        except PermissionError:
            messagebox.showerror("Σφάλμα", "Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό σε άλλο "
                                           "πρόγραμμα και δοκιμάστε ξανά.")
            return
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Κάτι πήγε στραβά:\n{e}")
            return
        finally:
            self._set_busy(False)

        warnings_txt = []
        if result['unmapped']:
            lines = ", ".join(f"{c} ({p} πελάτες / {n} διανυκτ.)" for c, p, n in result['unmapped'])
            warnings_txt.append("Αυτές οι χώρες ΔΕΝ μπήκαν στο PDF (λείπουν από το πρόγραμμα): " + lines)
        if result['excel_total'] and tuple(result['excel_total']) != tuple(result['computed_total']):
            warnings_txt.append(f"Το γενικό σύνολο του PDF {result['computed_total']} διαφέρει από το σύνολο "
                                f"του Excel {result['excel_total']}. Ελέγξτε το πριν το υποβάλετε!")

        try:
            gy, gm = int(m['year']), int(m['month'])
            if self.prefs['keep_archive']:
                archive_copy(gy, gm, out)
            save_month_record(gy, gm, m['employees'], m['days_of_month'], m['date_to'])
            self._status(f"✓ Το δελτίο {GREEK_MONTHS[gm - 1]} {gy} δημιουργήθηκε και αρχειοθετήθηκε.")
        except Exception:
            pass
        self.refresh_history()
        print_status = None
        if print_after:
            if warnings_txt:      # never send a form with problems straight to the printer
                print_status = (False, "Δεν εκτυπώθηκε αυτόματα επειδή υπάρχουν προειδοποιήσεις. "
                                       "Ελέγξτε το δελτίο και μετά πατήστε «Εκτύπωση».")
            else:
                print_status = print_file(out)
        if self.prefs['open_after'] and not print_after:
            open_file(out)
        ResultDialog(self, result, out, warnings_txt, print_status)


if __name__ == "__main__":
    if sys.platform.startswith('win'):      # sharp text on high-DPI screens
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = ElstatApp()
    app.mainloop()
