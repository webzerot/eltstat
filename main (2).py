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

# ---------------------------------------------------------------- language
# The interface is Greek by default; English can be chosen in the "Settings" tab.
# Every text the user sees is written as T("greek", "english") and is looked up at the moment
# it is needed, so switching language only requires building the screen again.
LANGUAGES = {'el': "Ελληνικά", 'en': "English"}
LANG = 'el'


def set_language(code):
    global LANG
    LANG = code if code in LANGUAGES else 'el'


def T(el, en):
    """The text in the current language."""
    return en if LANG == 'en' else el


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
            notices.append((
                f"Ο φάκελος δεδομένων\n{wanted}\nδεν είναι διαθέσιμος αυτή τη στιγμή (π.χ. "
                "αποσυνδεδεμένος δίσκος). Χρησιμοποιείται προσωρινά ο προεπιλεγμένος φάκελος.",
                f"The data folder\n{wanted}\nis not available right now (e.g. a disconnected "
                "drive). The default folder is being used temporarily."))
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
            notices.append((
                "Τα δεδομένα του προγράμματος (στοιχεία ξενοδοχείου, ιστορικό, αρχείο PDF) "
                f"αντιγράφηκαν στον φάκελο:\n{path}\n\nΤα παλιά αρχεία έμειναν στη θέση τους ως "
                "αντίγραφο ασφαλείας - μπορείτε να τα σβήσετε όταν βεβαιωθείτε ότι όλα λειτουργούν.\n"
                "Τον φάκελο δεδομένων τον αλλάζετε από την καρτέλα «Ρυθμίσεις».",
                "The program's data (hotel details, history, PDF archive) "
                f"was copied to the folder:\n{path}\n\nThe old files were left where they were as a "
                "backup - you can delete them once you are sure everything works.\n"
                "You can change the data folder in the “Settings” tab."))
        except Exception:
            pass
    return notices


# ------------------------------------------------------------- preferences
BOOL_PREFS = ('open_after', 'confirm_overwrite', 'keep_archive', 'show_reminder', 'prefill_employees')
DEFAULT_PREFS = {
    'language': 'el',                          # 'el' | 'en'  (interface language)
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
        return T("Το όνομα δεν μπορεί να είναι κενό.", "The name cannot be empty.")
    try:
        pattern.format(year=2026, month="01")
    except Exception:
        return T("Μη έγκυρο όνομα. Χρησιμοποιήστε μόνο τα {year} και {month} μέσα σε άγκιστρα.",
                 "Invalid name. Use only {year} and {month} inside curly braces.")
    if '{year}' not in pattern or '{month}' not in pattern:
        return (T("Πρέπει να περιέχει και το {year} και το {month}, αλλιώς κάθε μήνας θα γράφει "
                  "πάνω στον προηγούμενο.",
                  "It must contain both {year} and {month}, otherwise each month will "
                  "overwrite the previous one."))
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
    if prefs['language'] not in LANGUAGES:
        prefs['language'] = DEFAULT_PREFS['language']
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
        return True, T("✓ Το δελτίο στάλθηκε στον προεπιλεγμένο εκτυπωτή.",
                       "✓ The form was sent to the default printer.")
    except Exception:
        open_file(path)
        return False, (T("Δεν ήταν δυνατή η αυτόματη εκτύπωση. Άνοιξα το PDF - πατήστε Ctrl+P "
                         "μέσα στο πρόγραμμα του PDF για να το εκτυπώσετε.",
                         "Automatic printing was not possible. I opened the PDF - press Ctrl+P "
                         "in the PDF program to print it."))

def validate_afm(afm):
    """AFM must be exactly 9 digits (0-9)."""
    afm = str(afm).strip()
    return len(afm) == 9 and afm.isascii() and afm.isdigit()

def validate_hotel(settings):
    errors = []
    if not validate_afm(settings.get('afm', '')):
        errors.append(T("Το Α.Φ.Μ. πρέπει να έχει ακριβώς 9 ψηφία (μόνο αριθμοί).",
                        "The Tax ID (AFM) must have exactly 9 digits (numbers only)."))
    for key, label in [('type', T('Είδος Καταλύματος', 'Property type')), ('name', T('Τίτλος Καταλύματος', 'Property name')),
                       ('category', T('Κατηγορία', 'Category')), ('company', T('Επωνυμία Επιχείρησης', 'Business name'))]:
        if not settings.get(key, '').strip():
            errors.append(T(f"Συμπληρώστε το πεδίο «{label}» (στην καρτέλα «Στοιχεία ξενοδοχείου»).",
                            f"Fill in the “{label}” field (in the “Hotel details” tab)."))
    return errors

def validate_month(monthly):
    errors = []
    year, month = monthly.get('year', ''), monthly.get('month', '')
    y_ok = year.isascii() and year.isdigit() and len(year) == 4
    m_ok = month.isascii() and month.isdigit() and 1 <= int(month) <= 12
    if not y_ok:
        errors.append(T("Το Έτος πρέπει να έχει 4 ψηφία (π.χ. 2026).",
                        "The Year must have 4 digits (e.g. 2026)."))
    if not m_ok:
        errors.append(T("Ο Μήνας πρέπει να είναι αριθμός από 1 έως 12.",
                        "The Month must be a number from 1 to 12."))

    try:
        d_to = datetime.datetime.strptime(monthly.get('date_to', ''), "%d/%m/%Y")
        if y_ok and m_ok and (d_to.year != int(year) or d_to.month != int(month)):
            errors.append(T("Η ημερομηνία «Έως» δεν ταιριάζει με το Έτος/Μήνα.",
                            "The “To” date does not match the Year/Month."))
    except ValueError:
        errors.append(T("Η ημερομηνία «Έως» πρέπει να είναι της μορφής ΗΗ/ΜΜ/ΕΕΕΕ και υπαρκτή.",
                        "The “To” date must be in the format DD/MM/YYYY and must exist."))

    days = monthly.get('days_of_month', '')
    if not (days.isascii() and days.isdigit() and 1 <= int(days) <= 31):
        errors.append(T("Οι Ημέρες Λειτουργίας πρέπει να είναι αριθμός από 1 έως 31.",
                        "Operating days must be a number from 1 to 31."))
    elif y_ok and m_ok and int(days) > calendar.monthrange(int(year), int(month))[1]:
        errors.append(T("Οι Ημέρες Λειτουργίας είναι περισσότερες από τις ημέρες του μήνα.",
                        "Operating days exceed the number of days in the month."))

    emp = monthly.get('employees', '')
    if not emp:
        errors.append(T("Πληκτρολογήστε το Σύνολο Εργαζομένων (δεν υπάρχει στο Excel).",
                        "Enter the Total employees (it is not in the Excel)."))
    elif not (emp.isascii() and emp.isdigit() and len(emp) <= 3):
        errors.append(T("Το Σύνολο Εργαζομένων πρέπει να είναι αριθμός (0-999).",
                        "Total employees must be a number (0-999)."))
    return errors

def validate_inputs(settings, monthly, excel, template, out):
    errors = validate_hotel(settings) + validate_month(monthly)
    if not excel or not os.path.isfile(excel):
        errors.append(T("Επιλέξτε υπαρκτό αρχείο Excel.", "Choose an existing Excel file."))
    if not template or not os.path.isfile(template):
        errors.append(T("Το άδειο PDF (template) δεν βρέθηκε.", "The blank PDF (template) was not found."))
    if not out:
        errors.append(T("Επιλέξτε πού θα αποθηκευτεί το PDF.", "Choose where the PDF will be saved."))
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


def filter_hist(hist, country):
    """History reduced to one country (months stay, with empty row lists, so gaps are still right)."""
    if not country:
        return hist
    k = norm(country)
    return {y: {m: [r for r in rows if norm(r[0]) == k] for m, rows in months.items()}
            for y, months in hist.items()}


def export_history_excel(hist, path):
    """Writes the whole history to .xlsx (sheet 1: one row per month, sheet 2: one row per country)."""
    greece = norm('ΕΛΛΑΔΑ')
    months_rows, country_rows = [], []
    for y in sorted(hist):
        for m in sorted(hist[y]):
            rows = hist[y][m]
            cl, ni = sum(r[1] for r in rows), sum(r[2] for r in rows)
            gr = sum(r[1] for r in rows if norm(r[0]) == greece)
            months_rows.append({T("Έτος", "Year"): y, T("Μήνας", "Month"): m,
                                T("Πελάτες", "Guests"): cl, T("Διανυκτερεύσεις", "Overnight stays"): ni,
                                T("Μέση διαμονή", "Average stay"): round(ni / cl, 2) if cl else None,
                                T("Ελλάδα", "Greece"): gr, T("Αλλοδαποί", "Foreign"): cl - gr})
            for country, c, n in sorted(rows, key=lambda r: -r[1]):
                country_rows.append({T("Έτος", "Year"): y, T("Μήνας", "Month"): m,
                                     T("Χώρα", "Country"): country, T("Πελάτες", "Guests"): c,
                                     T("Διανυκτερεύσεις", "Overnight stays"): n})
    with pd.ExcelWriter(path, engine='openpyxl') as xw:
        pd.DataFrame(months_rows).to_excel(xw, sheet_name=T("Μήνες", "Months"), index=False)
        pd.DataFrame(country_rows).to_excel(xw, sheet_name=T("Χώρες", "Countries"), index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = max(12, min(30, max(len(str(c.value or '')) for c in col) + 3))
            ws.freeze_panes = 'A2'


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
    s = f"{int(n):,}"
    return s if LANG == 'en' else s.replace(",", ".")


def fmt_dec(x, signed=False):
    """One decimal digit with the decimal mark of the current language."""
    s = f"{x:+.1f}" if signed else f"{x:.1f}"
    return s if LANG == 'en' else s.replace(".", ",")


def fmt_avg(nights, clients):
    return fmt_dec(nights / clients) if clients else "-"


def make_year_summary_pdf(year, hist, hotel, out_path):
    """One-page A4 summary of a year: month table, Greek vs foreign guests, top countries."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    months = hist.get(year, {})
    if not months:
        raise ValueError(T("Δεν υπάρχουν δεδομένα για αυτό το έτος.", "There is no data for this year."))
    reg, bold = greek_fonts()
    W, H = A4
    L, R = 40, W - 40
    navy, text_c, band = colors.HexColor(NAVY), colors.HexColor(TEXT), colors.HexColor('#e8f0fe')
    c = canvas.Canvas(out_path, pagesize=A4)

    c.setFillColor(navy)
    c.rect(0, H - 92, W, 92, stroke=0, fill=1)
    c.setFillColor(colors.white)
    c.setFont(bold, 19)
    c.drawString(L, H - 46, T(f"Σύνοψη κίνησης έτους {year}", f"Year {year} movement summary"))
    c.setFont(reg, 10.5)
    who = " · ".join(x for x in (hotel.get('type'), hotel.get('name'), hotel.get('category')) if x)
    c.drawString(L, H - 68, who or T("Τουριστικό κατάλυμα", "Tourist accommodation"))

    def header(y, first, cols):
        c.setFillColor(navy)
        c.rect(L, y - 6, R - L, 22, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont(bold, 9.5)
        c.drawString(L + 8, y, first)
        for x, t in cols:
            c.drawRightString(x - 6, y, t)

    greece = norm('ΕΛΛΑΔΑ')
    cols = [(220, T("Πελάτες", "Guests")), (320, T("Διανυκτερεύσεις", "Overnight stays")), (410, T("Μέση διαμονή", "Average stay")), (480, T("Ελλάδα", "Greece")), (R, T("Αλλοδαποί",
                                                                                                      "Foreign"))]
    y = H - 130
    header(y, T("Μήνας", "Month"), cols)
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
        c.drawString(L + 8, y, month_name(month) if 1 <= month <= 12 else str(month))
        for (x, _), v in zip(cols, (fmt_int(cl), fmt_int(ni), fmt_avg(ni, cl), fmt_int(gr), fmt_int(cl - gr))):
            c.drawRightString(x - 6, y, v)
        y -= 20
    c.setFillColor(band)
    c.rect(L, y - 6, R - L, 22, stroke=0, fill=1)
    c.setFillColor(navy)
    c.setFont(bold, 10.5)
    c.drawString(L + 8, y, T(f"ΣΥΝΟΛΟ {year}", f"TOTAL {year}"))
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
    c.drawString(L, y, T("Κορυφαίες χώρες προέλευσης", "Top countries of origin"))
    y -= 26
    cols2 = [(300, T("Πελάτες", "Guests")), (410, T("Διανυκτερεύσεις", "Overnight stays")), (R, T("% των πελατών", "% of guests"))]
    header(y, T("Χώρα", "Country"), cols2)
    y -= 24
    for i, (country, (cl, ni)) in enumerate(top):
        if i % 2 == 0:
            c.setFillColor(colors.HexColor('#f4f7fb'))
            c.rect(L, y - 6, R - L, 20, stroke=0, fill=1)
        c.setFillColor(text_c)
        c.setFont(reg, 10)
        c.drawString(L + 8, y, country)
        pct = f"{fmt_dec(cl / tcl * 100)}%" if tcl else "-"
        for (x, _), v in zip(cols2, (fmt_int(cl), fmt_int(ni), pct)):
            c.drawRightString(x - 6, y, v)
        y -= 20

    c.setFillColor(colors.HexColor(MUTED))
    c.setFont(reg, 8.5)
    c.drawString(L, 30, T(f"Δημιουργήθηκε στις {datetime.date.today().strftime('%d/%m/%Y')} από το πρόγραμμα ΕΛΣΤΑΤ",
                          f"Created on {datetime.date.today().strftime('%d/%m/%Y')} by the ELSTAT program"))
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
        raise ValueError(T("Το template PDF πρέπει να έχει 2 σελίδες.", "The template PDF must have 2 pages."))
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
        (T("1.  Ελλάδα", "1.  Greece"), gr['clients'], gr['nights'], False),
        (T("2.1  Χώρες Ευρωπαϊκής Ένωσης", "2.1  European Union countries"), totals['2.1']['clients'], totals['2.1']['nights'], False),
        (T("2.2  Χώρες Ε.Ζ.Ε.Σ. (EFTA)", "2.2  EFTA countries"), totals['2.2']['clients'], totals['2.2']['nights'], False),
        (T("2.3  Λοιπές χώρες Ευρώπης", "2.3  Other European countries"), totals['2.3']['clients'], totals['2.3']['nights'], False),
        (T("2.  Σύνολο Ευρώπης (εκτός Ελλάδας)", "2.  Total Europe (excluding Greece)"), totals['2']['clients'], totals['2']['nights'], True),
        (T("3.  Σύνολο Αμερικής", "3.  Total America"), totals['3']['clients'], totals['3']['nights'], True),
        (T("4.  Σύνολο Ασίας", "4.  Total Asia"), totals['4']['clients'], totals['4']['nights'], True),
        (T("5.  Σύνολο Αφρικής", "5.  Total Africa"), totals['5']['clients'], totals['5']['nights'], True),
        (T("6.  Σύνολο Ωκεανίας", "6.  Total Oceania"), totals['6']['clients'], totals['6']['nights'], True),
        (T("7.  Δεν δήλωσαν χώρα", "7.  Country not declared"), totals['7']['clients'], totals['7']['nights'], False),
        (T("8.  Σύνολο αλλοδαπών", "8.  Total foreign"), totals['8']['clients'], totals['8']['nights'], True),
        (T("9.  ΓΕΝΙΚΟ ΣΥΝΟΛΟ", "9.  GRAND TOTAL"), totals['9']['clients'], totals['9']['nights'], True),
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
ENGLISH_MONTHS = ["January", "February", "March", "April", "May", "June",
                  "July", "August", "September", "October", "November", "December"]


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
ENGLISH_MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def month_name(month):
    """Name of month 1-12 in the current language."""
    return (ENGLISH_MONTHS if LANG == 'en' else GREEK_MONTHS)[month - 1]


def month_short(month):
    return (ENGLISH_MONTHS_SHORT if LANG == 'en' else GREEK_MONTHS_SHORT)[month - 1]


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
    """Bar chart with one bar per month. Clicking a bar calls on_pick(month).
    Optionally shows the previous year as small tick marks and a tooltip on hover."""

    def __init__(self, parent, on_pick=None, height=160):
        super().__init__(parent, height=height, bg=CARD, highlightthickness=1,
                         highlightbackground=BORDER, borderwidth=0)
        self.values, self.title, self.color, self.selected = [0] * 12, "", ACCENT, None
        self.prev, self.prev_label, self.hover = None, "", None
        self.on_pick = on_pick
        self._slots = []
        self.bind('<Configure>', lambda e: self.draw())
        self.bind('<Button-1>', self._click)
        self.bind('<Motion>', self._motion)
        self.bind('<Leave>', self._leave)

    def set_data(self, values, title, color=ACCENT, selected=None, prev=None, prev_label=""):
        self.values, self.title, self.color, self.selected = list(values), title, color, selected
        self.prev = list(prev) if prev and any(prev) else None
        self.prev_label = prev_label
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
        return fmt_int(round(v))

    def draw(self):
        self.delete('all')
        self._slots = []
        w, h = self.winfo_width(), self.winfo_height()
        if w < 120 or h < 90:
            return
        left, right, top, bottom = 52, 14, 38, 26
        plot_h = h - top - bottom
        top_val = self._nice(max(self.values + (self.prev or [0])))
        self.create_text(14, 14, text=self.title, anchor="w", fill=NAVY, font=(FONT, 10, 'bold'))
        if self.prev:     # legend
            self.create_line(w - right - 150, 14, w - right - 134, 14, fill='#f59e0b', width=3)
            self.create_text(w - right - 128, 14, text=self.prev_label, anchor="w", fill=MUTED, font=(FONT, 8))
        for frac in (0, .5, 1):
            y = top + plot_h * (1 - frac)
            self.create_line(left, y, w - right, y, fill='#e5eaf1')
            self.create_text(left - 8, y, text=self._fmt(top_val * frac), anchor="e", fill=MUTED, font=(FONT, 8))
        slot = (w - left - right) / 12
        bw = slot * 0.6
        for i, v in enumerate(self.values):
            cx = left + slot * (i + .5)
            if self.hover == i + 1:
                self.create_rectangle(cx - slot / 2 + 2, top - 6, cx + slot / 2 - 2, top + plot_h + 2,
                                      fill='#f1f5f9', outline='')
        for i, v in enumerate(self.values):
            cx = left + slot * (i + .5)
            if v > 0:
                y0 = top + plot_h * (1 - v / top_val)
                fill = NAVY if self.selected == i + 1 else self.color
                self.create_rectangle(cx - bw / 2, y0, cx + bw / 2, top + plot_h, fill=fill, outline='')
                self.create_text(cx, y0 - 8, text=self._fmt(v), fill=TEXT, font=(FONT, 8))
            else:
                self.create_line(cx - bw / 2, top + plot_h - 1, cx + bw / 2, top + plot_h - 1,
                                 fill='#cbd5e1', width=2)
            pv = self.prev[i] if self.prev else 0
            if pv > 0:
                yp = top + plot_h * (1 - pv / top_val)
                self.create_line(cx - bw / 2 - 3, yp, cx + bw / 2 + 3, yp, fill='#f59e0b', width=3)
            if v > 0 or pv > 0:
                self._slots.append((cx - slot / 2, cx + slot / 2, i + 1))
            self.create_text(cx, h - 12, text=month_short(i + 1), fill=MUTED, font=(FONT, 8))
        if self.hover:
            self._tooltip(self.hover, left, slot, top)

    def _tooltip(self, month, left, slot, top):
        v = self.values[month - 1]
        lines = [f"{month_name(month)}: {self._fmt(v)}"]
        if self.prev:
            pv = self.prev[month - 1]
            lines.append(f"{self.prev_label}: {self._fmt(pv)}")
            if pv > 0:
                lines.append(f"{fmt_dec((v - pv) / pv * 100, signed=True)}%")
        cx = left + slot * (month - .5)
        w = self.winfo_width()
        tw = 150
        x = cx + slot / 2 + 4
        if x + tw > w - 4:
            x = cx - slot / 2 - 4 - tw
        y = top - 4
        self.create_rectangle(x, y, x + tw, y + 14 * len(lines) + 10, fill=NAVY, outline='')
        self.create_text(x + 8, y + 6, text="\n".join(lines), anchor="nw", fill="white",
                         font=(FONT, 9), justify="left")

    def _month_at(self, x):
        for x0, x1, month in self._slots:
            if x0 <= x <= x1:
                return month
        return None

    def _motion(self, event):
        m = self._month_at(event.x)
        if m != self.hover:
            self.hover = m
            self.configure(cursor='hand2' if m else '')
            self.draw()

    def _leave(self, _event=None):
        if self.hover is not None:
            self.hover = None
            self.draw()

    def _click(self, event):
        m = self._month_at(event.x)
        if m and self.on_pick:
            self.on_pick(m)


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
        self.title(T("Αποτέλεσμα", "Result"))
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
            ttk.Label(body, text=T("✓ Το δελτίο δημιουργήθηκε επιτυχώς", "✓ The form was created successfully"), style='GoodTitle.TLabel').pack(anchor="w")
            gc, gn = result['computed_total']
            note(T(f"Τα σύνολα ταιριάζουν με το Excel του Protel ({gc} πελάτες / {gn} διανυκτερεύσεις).",
                   f"The totals match the Protel Excel ({gc} guests / {gn} overnight stays)."),
                 'Good.TLabel', pady=(2, 12))
        else:
            ttk.Label(body, text=T("⚠ Το δελτίο δημιουργήθηκε, αλλά θέλει έλεγχο",
                                   "⚠ The form was created, but needs checking"),
                      style='BadTitle.TLabel').pack(anchor="w")
            for w in warnings_txt:
                note("• " + w, 'Bad.TLabel', pady=(4, 0))
            ttk.Label(body, text="").pack()

        # buttons first (bottom) so they never get squeezed out when the window is small
        btns = ttk.Frame(body)
        btns.pack(side="bottom", fill="x", pady=(14, 0))
        ttk.Button(btns, text=T("Εκτύπωση", "Print"), style='Print.TButton',
                   command=self.print_now).pack(side="left")
        ttk.Button(btns, text=T("Άνοιγμα PDF", "Open PDF"), command=lambda: open_file(out_path)).pack(side="left", padx=8)
        ttk.Button(btns, text=T("Άνοιγμα φακέλου", "Open folder"),
                   command=lambda: open_file(os.path.dirname(out_path) or ".")).pack(side="left")
        ttk.Button(btns, text=T("Κλείσιμο", "Close"), command=self.destroy).pack(side="right")

        self.print_var = tk.StringVar()
        self.print_lbl = ttk.Label(body, textvariable=self.print_var, style='Hint.TLabel', justify="left")
        self.print_lbl.pack(side="bottom", anchor="w", fill="x", pady=(10, 0))
        autowrap(self.print_lbl, self, margin=50, minimum=260)
        if print_status:
            self._show_print(*print_status)

        steps = ttk.Label(body, justify="left",
                          text=T("Επόμενα βήματα: εκτυπώστε το δελτίο, γράψτε το όνομα, βάλτε υπογραφή και σφραγίδα "
                                 "και υποβάλετε το στην ΕΛΣΤΑΤ.",
                                 "Next steps: print the form, write your name, sign and stamp it, "
                                 "and submit it to ELSTAT."))
        steps.pack(side="bottom", anchor="w", fill="x", pady=(10, 0))
        autowrap(steps, self, margin=50, minimum=260)
        where = ttk.Label(body, text=T("Αποθηκεύτηκε στο:\n", "Saved to:\n") + out_path, style='Hint.TLabel', justify="left")
        where.pack(side="bottom", anchor="w", fill="x", pady=(10, 0))
        autowrap(where, self, margin=50, minimum=260)

        tree_frame = ttk.Frame(body)
        tree_frame.pack(side="top", fill="both", expand=True)
        tree = ttk.Treeview(tree_frame, columns=('c', 'n'), show='tree headings',
                            height=min(len(result['summary']), 12))
        tree.heading('#0', text=T('Κατηγορία (όπως στο δελτίο)', 'Category (as on the form)'), anchor='w')
        tree.heading('c', text=T('Πελάτες', 'Guests'))
        tree.heading('n', text=T('Διανυκτερεύσεις', 'Overnight stays'))
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
        self.geometry("920x840")
        self.minsize(640, 560)
        self.configure(bg=BG)

        self._storage_notices = init_storage()
        self.prefs = load_prefs()
        set_language(self.prefs['language'])
        self.current_settings = load_hotel_settings()
        self._rebuilding = False
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
            self.after(400, lambda: messagebox.showinfo(T("Ρυθμίσεις", "Settings"),
                                                        "\n\n".join(T(*n) for n in self._storage_notices),
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
        self._period_msg = (lambda: T("Πατήστε «Επιλογή Excel…» και διαλέξτε το αρχείο «Nationalities Statistics» "
                                      "που βγάζει το Protel (.xls ή .xlsx).",
                                      "Click “Choose Excel…” and select the “Nationalities Statistics” file "
                                      "produced by Protel (.xls or .xlsx)."), 'CardHint.TLabel')
        self.month_name_var = tk.StringVar()
        self.afm_counter_var = tk.StringVar()
        self.check_var = tk.StringVar()
        self.entries = {}
        self.status_var = tk.StringVar(value=T("Έτοιμο.", "Ready."))
        self.hist_year_var = tk.StringVar()
        self.hist_metric_var = tk.StringVar(value='nights')
        self.hist_country_var = tk.StringVar()
        self.top_title_var = tk.StringVar()
        self._countries = []
        self._hist = {}
        self._reminder_dismissed = False
        # settings-tab variables
        self.pv = {k: tk.StringVar(value=self.prefs[k])
                   for k in ('language', 'output_mode', 'output_dir', 'filename_pattern', 'excel_dir')}
        self.pv.update({k: tk.BooleanVar(value=self.prefs[k]) for k in BOOL_PREFS})
        self.data_dir_var = tk.StringVar(value=DATA_DIR)
        self.name_note_var = tk.StringVar()
        self.out_note_var = tk.StringVar()

    def _build_ui(self):
        self.title(T("ΕΛΣΤΑΤ – Μηνιαίο Δελτίο Κίνησης Καταλύματος",
                     "ELSTAT – Monthly Accommodation Movement Report"))
        header = tk.Frame(self, bg=NAVY)
        header.pack(fill="x")
        row = tk.Frame(header, bg=NAVY)
        row.pack(fill="x", padx=22, pady=16)
        tk.Label(row, text=T("ΕΛΣΤΑΤ", "ELSTAT"), bg=ACCENT, fg="white", font=(FONT, 11, 'bold'),
                 padx=12, pady=6).pack(side="left", anchor="n")
        texts = tk.Frame(row, bg=NAVY)
        texts.pack(side="left", fill="x", expand=True, padx=(16, 0))
        title = tk.Label(texts, text=T("Μηνιαίο Δελτίο Κίνησης Τουριστικού Καταλύματος",
                                       "Monthly Tourist Accommodation Movement Report"), bg=NAVY, fg="white",
                         font=(FONT, 16, 'bold'), justify="left")
        title.pack(anchor="w")
        sub = tk.Label(texts, text=T("Διαλέξτε το Excel από το Protel και το πρόγραμμα συμπληρώνει και "
                                     "εκτυπώνει το δελτίο της ΕΛΣΤΑΤ.",
                                     "Choose the Excel from Protel and the program fills in and "
                                     "prints the ELSTAT form."), bg=NAVY, fg="#b6c4d8",
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
        self.notebook.add(self.tab_report, text=T("Δελτίο μήνα", "Monthly form"))
        self.notebook.add(self.tab_hotel, text=T("Στοιχεία ξενοδοχείου", "Hotel details"))
        self.notebook.add(self.tab_history, text=T("Ιστορικό", "History"))
        self.notebook.add(self.tab_settings, text=T("Ρυθμίσεις", "Settings"))
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
        ttk.Button(self.reminder_frame, text=T("Κλείσιμο", "Close"), command=self._dismiss_reminder).pack(side="right", padx=8)
        autowrap(self.reminder_label, body, margin=190)
        self.reminder_frame.grid_remove()

        # ---- Step 1
        f1 = Card(body, 1, T("Αρχείο από το Protel", "File from Protel"))
        f1.grid(row=1, column=0, sticky="ew", padx=4, pady=(12, 6))
        c1 = f1.body
        c1.columnconfigure(0, weight=1)
        ent = ttk.Entry(c1, textvariable=self.excel_var)
        ent.grid(row=0, column=0, sticky="ew")
        ent.bind('<FocusOut>', self._excel_typed)
        ent.bind('<Return>', self._excel_typed)
        self.excel_btn = ttk.Button(c1, text=T("Επιλογή Excel…", "Choose Excel…"), command=self.browse_excel)
        self.excel_btn.grid(row=0, column=1, padx=(8, 0))
        ttk.Button(c1, text=T("Πιο πρόσφατο", "Latest"), command=self.use_recent_excel).grid(row=0, column=2, padx=(6, 0))
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
        self.preview_tree.heading('#0', text=T('Χώρα (όπως στο Excel)', 'Country (as in the Excel)'), anchor='w')
        self.preview_tree.heading('p', text=T('Πελάτες', 'Guests'))
        self.preview_tree.heading('n', text=T('Διανυκτ.', 'Nights'))
        self.preview_tree.heading('s', text=T('Στο δελτίο', 'On the form'), anchor='w')
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
        self._render_period()

        # ---- Step 2
        f2 = Card(body, 2, T("Στοιχεία του μήνα", "Month details"))
        f2.grid(row=2, column=0, sticky="ew", padx=4, pady=6)
        c2 = f2.body
        c2.columnconfigure(2, weight=1)
        fields = [
            (T("Έτος", "Year"), 'year', 8, T("αυτόματα από το Excel", "filled in from the Excel")),
            (T("Μήνας (αριθμός)", "Month (number)"), 'month', 8, None),
            (T("Τελευταία ημέρα («Έως»)", "Last day (“To”)"), 'date_to', 14, T("αυτόματα από το Excel (ΗΗ/ΜΜ/ΕΕΕΕ)",
                                                         "filled in from the Excel (DD/MM/YYYY)")),
            (T("Ημέρες λειτουργίας", "Operating days"), 'days_of_month', 8, T("αυτόματα από το Excel", "filled in from the Excel")),
            (T("Σύνολο εργαζομένων", "Total employees"), 'employees', 8,
             T("✎ ΔΕΝ υπάρχει στο Excel - πληκτρολογήστε το εσείς (στο τέλος του μήνα, "
               "μαζί με ιδιοκτήτες/οικογένεια)",
               "✎ NOT in the Excel - type it yourself (at the end of the month, "
               "including owners/family)") +
             (T(f"  ·  τελευταία φορά: {self.current_settings['last_employees']}",
                f"  ·  last time: {self.current_settings['last_employees']}")
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
        f3 = Card(body, 3, T("Πού να αποθηκευτεί το PDF", "Where to save the PDF"))
        f3.grid(row=3, column=0, sticky="ew", padx=4, pady=(6, 14))
        c3 = f3.body
        c3.columnconfigure(0, weight=1)
        ttk.Entry(c3, textvariable=self.out_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(c3, text=T("Αλλαγή…", "Change…"), command=self.browse_output).grid(row=0, column=1, padx=(8, 0))
        h3 = ttk.Label(c3, text=T("Προτείνεται αυτόματα με βάση την καρτέλα «Ρυθμίσεις» (φάκελος και όνομα αρχείου).",
                                  "Suggested automatically based on the “Settings” tab (folder and file name)."),
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
        self.run_btn = ttk.Button(btnrow, text=T("Δημιουργία PDF", "Create PDF"), style='Primary.TButton',
                                  command=self.generate)
        self.run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.print_btn = ttk.Button(btnrow, text=T("Δημιουργία & Εκτύπωση", "Create & Print"), style='Print.TButton',
                                    command=lambda: self.generate(print_after=True))
        self.print_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        ttk.Label(bottom, text=T("Συντομεύσεις:  Ctrl+O = επιλογή Excel  ·  Ctrl+Enter = δημιουργία  ·  "
                                 "Ctrl+P = δημιουργία & εκτύπωση",
                                 "Shortcuts:  Ctrl+O = choose Excel  ·  Ctrl+Enter = create  ·  "
                                 "Ctrl+P = create & print"),
                  style='CardHint.TLabel').pack(anchor="e", padx=16, pady=(4, 10))

    def _build_hotel_tab(self):
        sf = ScrollFrame(self.tab_hotel)
        sf.pack(fill="both", expand=True)
        body = sf.body
        body.columnconfigure(0, weight=1)

        card = Card(body, None, T("Στοιχεία ξενοδοχείου (αλλάζουν σπάνια)", "Hotel details (rarely change)"))
        card.grid(row=0, column=0, sticky="ew", padx=4, pady=(12, 6))
        f = card.body
        f.columnconfigure(1, weight=1)          # entries stretch with the window
        fields = [
            (T("Είδος καταλύματος", "Property type"), 'type', T("όπως θέλετε να γραφτεί στο δελτίο",
                                            "as you want it written on the form")),
            (T("Τίτλος καταλύματος", "Property name"), 'name', T("το όνομα όπως θα εμφανιστεί στο δελτίο",
                                             "the name as it will appear on the form")),
            (T("Κατηγορία", "Category"), 'category', T("όπως θέλετε να γραφτεί στο δελτίο",
                                        "as you want it written on the form")),
            ("E-mail", 'email', None),
            (T("Τηλέφωνο", "Phone"), 'phone', None),
            (T("Α.Φ.Μ.", "Tax ID (AFM)"), 'afm', T("ακριβώς 9 ψηφία", "exactly 9 digits")),
            (T("Επωνυμία επιχείρησης", "Business name"), 'company', None),
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
            note = (T("Τα στοιχεία διαβάστηκαν από το αρχείο hotel_settings.csv. Αν τα αλλάξετε, αποθηκεύονται "
                      "αυτόματα όταν δημιουργήσετε δελτίο.",
                      "The details were read from the file hotel_settings.csv. If you change them, "
                      "they are saved automatically when you create a form."))
        else:
            note = (T("Δεν βρέθηκε το αρχείο hotel_settings.csv. Συμπληρώστε τα στοιχεία εδώ - θα αποθηκευτούν "
                      "αυτόματα την πρώτη φορά που θα δημιουργήσετε δελτίο και δεν θα τα ξαναγράψετε.",
                      "The file hotel_settings.csv was not found. Fill in the details here - they will "
                      "be saved automatically the first time you create a form, and you won't have to "
                      "type them again."))
        nl = ttk.Label(body, text=note, style='Hint.TLabel', justify="left")
        nl.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 8))
        autowrap(nl, body, margin=50)

        card2 = Card(body, None, T("Άδειο PDF της ΕΛΣΤΑΤ (template)", "Blank ELSTAT PDF (template)"))
        card2.grid(row=2, column=0, sticky="ew", padx=4, pady=(6, 14))
        f2 = card2.body
        f2.columnconfigure(0, weight=1)
        ttk.Entry(f2, textvariable=self.pdf_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(f2, text=T("Αναζήτηση…", "Browse…"), command=self.browse_pdf).grid(row=0, column=1, padx=(8, 0))
        h2 = ttk.Label(f2, text=T("Το κενό δελτίο πάνω στο οποίο γράφονται τα στοιχεία. "
                                  "Συνήθως δεν χρειάζεται να το αλλάξετε.",
                                  "The blank form on which the details are written. "
                                  "You usually don't need to change it."), style='CardHint.TLabel', justify="left")
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

        # ---- 0. interface language
        c0 = Card(body, None, T("Γλώσσα / Language", "Language"))
        c0.grid(row=0, column=0, sticky="ew", padx=4, pady=(12, 6))
        f0 = c0.body
        f0.columnconfigure(0, weight=1)
        lang_row = ttk.Frame(f0, style='Card.TFrame')
        lang_row.grid(row=0, column=0, sticky="w")
        for code, label in LANGUAGES.items():       # the names stay in their own language on purpose
            ttk.Radiobutton(lang_row, text=label, value=code, variable=pv['language'],
                            style='Card.TRadiobutton',
                            command=self._language_clicked).pack(side="left", padx=(0, 24))
        hint(f0, T("Αφορά μόνο το περιβάλλον του προγράμματος και αλλάζει αμέσως. Το επίσημο δελτίο της "
                   "ΕΛΣΤΑΤ (PDF) συμπληρώνεται πάντα στα ελληνικά.",
                   "Applies to the program's interface only and changes immediately. The official ELSTAT "
                   "form (PDF) is always filled in Greek."), 1)

        # ---- 1. where the PDFs go
        c1 = Card(body, None, T("Πού αποθηκεύονται τα PDF", "Where PDFs are saved"))
        c1.grid(row=1, column=0, sticky="ew", padx=4, pady=6)
        f = c1.body
        f.columnconfigure(0, weight=1)
        for r, (text, value) in enumerate(((T("Στον φάκελο του προγράμματος", "In the program folder"), 'program'),
                                           (T("Στον φάκελο του αρχείου Excel (από το Protel)",
                                              "In the folder of the Excel file (from Protel)"), 'excel'),
                                           (T("Σε συγκεκριμένο φάκελο:", "In a specific folder:"), 'custom'))):
            ttk.Radiobutton(f, text=text, value=value, variable=pv['output_mode'], style='Card.TRadiobutton',
                            command=self._pref_changed).grid(row=r, column=0, columnspan=2, sticky="w",
                                                             pady=(0 if r == 0 else 6, 0))
        self.out_dir_entry = ttk.Entry(f, textvariable=pv['output_dir'])
        self.out_dir_entry.grid(row=3, column=0, sticky="ew", padx=(24, 0), pady=(4, 0))
        commit_on(self.out_dir_entry)
        self.out_dir_btn = ttk.Button(f, text=T("Αναζήτηση…", "Browse…"), command=self.browse_output_dir)
        self.out_dir_btn.grid(row=3, column=1, padx=(8, 0), pady=(4, 0))
        ttk.Separator(f, orient="horizontal").grid(row=4, column=0, columnspan=2, sticky="ew", pady=14)
        ttk.Label(f, text=T("Όνομα αρχείου PDF:", "PDF file name:"), style='CardBold.TLabel').grid(row=5, column=0, columnspan=2,
                                                                              sticky="w")
        name_entry = ttk.Entry(f, textvariable=pv['filename_pattern'])
        name_entry.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        commit_on(name_entry)
        self.name_note = hint(f, "", 7, columnspan=2)
        self.name_note.configure(textvariable=self.name_note_var)
        self.out_note = hint(f, "", 8, columnspan=2, pady=(6, 0))
        self.out_note.configure(textvariable=self.out_note_var)

        # ---- 2. data folder
        c2 = Card(body, None, T("Φάκελος δεδομένων", "Data folder"))
        c2.grid(row=2, column=0, sticky="ew", padx=4, pady=6)
        f2 = c2.body
        f2.columnconfigure(0, weight=1)
        ttk.Entry(f2, textvariable=self.data_dir_var, state='readonly').grid(row=0, column=0, sticky="ew")
        ttk.Button(f2, text=T("Αλλαγή…", "Change…"), command=self.change_data_dir).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(f2, text=T("Άνοιγμα", "Open"), command=lambda: open_file(DATA_DIR)).grid(row=0, column=2, padx=(6, 0))
        hint(f2, T("Εδώ αποθηκεύονται τα στοιχεία του ξενοδοχείου, το ιστορικό, τα αρχειοθετημένα PDF και οι "
                   "ρυθμίσεις. Μπορείτε να διαλέξετε π.χ. φάκελο μέσα στο OneDrive ώστε να υπάρχει αυτόματο "
                   "αντίγραφο. Η θέση του φακέλου θυμάται το αρχείο elstat_config.json δίπλα στο πρόγραμμα.",
                   "This is where the hotel details, history, archived PDFs and settings are saved. "
                   "You can choose, for example, a folder inside OneDrive so that an automatic copy "
                   "exists. The folder location is remembered in the file elstat_config.json next to "
                   "the program."),
             1, columnspan=3)
        row2 = ttk.Frame(f2, style='Card.TFrame')
        row2.grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Button(row2, text=T("Προεπιλεγμένος φάκελος", "Default folder"),
                   command=lambda: self._switch_data_dir(DEFAULT_DATA_DIR)).pack(side="left")
        ttk.Button(row2, text=T("Αντίγραφο ασφαλείας (.zip)…", "Backup (.zip)…"), command=self.backup_data).pack(side="left", padx=8)

        # ---- 3. Excel folder
        c3 = Card(body, None, T("Αρχεία Excel από το Protel", "Excel files from Protel"))
        c3.grid(row=3, column=0, sticky="ew", padx=4, pady=6)
        f3 = c3.body
        f3.columnconfigure(0, weight=1)
        ex_entry = ttk.Entry(f3, textvariable=pv['excel_dir'])
        ex_entry.grid(row=0, column=0, sticky="ew")
        commit_on(ex_entry)
        ttk.Button(f3, text=T("Αναζήτηση…", "Browse…"), command=self.browse_excel_dir).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(f3, text=T("Καθαρισμός", "Clear"), command=self.clear_excel_dir).grid(row=0, column=2, padx=(6, 0))
        hint(f3, T("Από ποιον φάκελο ξεκινούν τα «Επιλογή Excel…» και «Πιο πρόσφατο». "
                   "Κενό = ο φάκελος του τελευταίου Excel που χρησιμοποιήσατε.",
                   "The folder where “Choose Excel…” and “Latest” start. "
                   "Empty = the folder of the last Excel you used."), 1, columnspan=3)

        # ---- 4. behaviour
        c4 = Card(body, None, T("Συμπεριφορά του προγράμματος", "Program behaviour"))
        c4.grid(row=4, column=0, sticky="ew", padx=4, pady=6)
        f4 = c4.body
        f4.columnconfigure(0, weight=1)
        behaviours = [
            ('open_after', T("Άνοιγμα του PDF αμέσως μετά τη δημιουργία",
                             "Open the PDF right after it is created"), None),
            ('confirm_overwrite', T("Ερώτηση πριν αντικατασταθεί PDF που υπάρχει ήδη",
                                    "Ask before replacing a PDF that already exists"), None),
            ('keep_archive', T("Αρχειοθέτηση αντιγράφου κάθε δελτίου στον φάκελο δεδομένων",
                               "Archive a copy of every form in the data folder"),
             T("Επιτρέπει να ξανανοίξετε ή να εκτυπώσετε παλιούς μήνες από το Ιστορικό. Αν το κλείσετε, το "
               "δελτίο ξαναφτιάχνεται από τα δεδομένα του ιστορικού όταν το ζητήσετε.",
               "Lets you reopen or print old months from History. If you turn it off, the "
               "form is rebuilt from the history data when you ask for it.")),
            ('show_reminder', T("Υπενθύμιση όταν λείπει το δελτίο του προηγούμενου μήνα",
                                "Remind me when last month's form is missing"), None),
            ('prefill_employees', T("Προσυμπλήρωση του «Σύνολο εργαζομένων» με την τελευταία τιμή",
                                    "Pre-fill “Total employees” with the last value"),
             T("Ελέγχετε πάντα τον αριθμό πριν δημιουργήσετε το δελτίο.",
               "Always check the number before creating the form.")),
        ]
        for r, (key, text, note) in enumerate(behaviours):
            ttk.Checkbutton(f4, text=text, variable=pv[key], style='Card.TCheckbutton',
                            command=self._pref_changed).grid(row=r * 2, column=0, sticky="w",
                                                             pady=(8 if r else 0, 0))
            if note:
                hint(f4, note, r * 2 + 1, padx=(24, 0), pady=(2, 0), margin=130)

        # ---- reset
        bottom = ttk.Frame(body)
        bottom.grid(row=5, column=0, sticky="ew", padx=4, pady=(6, 16))
        ttk.Button(bottom, text=T("Επαναφορά ρυθμίσεων", "Reset settings"), style='Danger.TButton',
                   command=self.reset_prefs).pack(side="left")
        ttk.Label(bottom, text=T("Δεν διαγράφει ιστορικό ή στοιχεία ξενοδοχείου.",
                                 "Does not delete history or hotel details."),
                  style='Hint.TLabel').pack(side="left", padx=10)

    # ---- settings: reading / saving
    def _save_prefs(self):
        try:
            save_prefs(self.prefs)
        except OSError as e:
            self._status(T(f"⚠ Δεν ήταν δυνατή η αποθήκευση των ρυθμίσεων: {e}",
                           f"⚠ Could not save the settings: {e}"))

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

    def _language_clicked(self):
        code = self.pv['language'].get()
        if code == LANG:
            return
        self.prefs['language'] = code
        self._save_prefs()
        self.after_idle(self._apply_language)       # not inside the click that is about to be destroyed

    def _apply_language(self):
        """Switches the interface language: the screen is built again, all typed-in values are kept."""
        set_language(self.prefs['language'])
        self.pv['language'].set(self.prefs['language'])
        self._rebuilding = True                     # the field-change traces must not touch half-built widgets
        try:
            try:
                tab = self.notebook.index(self.notebook.select())
            except tk.TclError:
                tab = 0
            for w in self.winfo_children():
                w.destroy()
            self.entries = {}
            self._build_ui()
            self.notebook.select(tab)
        finally:
            self._rebuilding = False
        self._after_prefs_change()
        self.refresh_validation()
        self._render_period()
        if self._loaded_excel and os.path.isfile(self._loaded_excel):
            self.load_preview(self._loaded_excel)
        self._status(T("Η γλώσσα άλλαξε σε Ελληνικά.", "Language changed to English."))

    def _sync_pref_vars(self):
        for k, var in self.pv.items():
            var.set(self.prefs[k])

    def _after_prefs_change(self):
        custom = self.prefs['output_mode'] == 'custom'
        for w in (self.out_dir_entry, self.out_dir_btn):
            w.state(['!disabled' if custom else 'disabled'])
        self._update_name_note()
        self._refresh_output_suggestion()
        self.out_note_var.set(T("Στο Βήμα 3 του δελτίου προτείνεται: ",
                                "In Step 3 of the form the suggestion is: ") + self._suggest_output() +
                              T("  (μπορείτε πάντα να το αλλάξετε εκεί)", "  (you can always change it there)"))
        self.update_reminder()

    def _update_name_note(self, *_):
        if self._rebuilding:
            return
        pattern = self.pv['filename_pattern'].get().strip()
        err = pattern_error(pattern)
        if err:
            self.name_note_var.set("⚠ " + err + T(" (δεν αποθηκεύτηκε)", " (not saved)"))
            self.name_note.configure(style='CardBad.TLabel')
        else:
            today = datetime.date.today()
            self.name_note_var.set(T(f"Παράδειγμα: {build_pdf_name(pattern, today.year, today.month)}   ·   ",
                                     f"Example: {build_pdf_name(pattern, today.year, today.month)}   ·   ") +
                                   T("{year} = έτος, {month} = μήνας (2 ψηφία)",
                                     "{year} = year, {month} = month (2 digits)"))
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
        d = self._pick_folder(T("Διαλέξτε φάκελο για τα PDF", "Choose a folder for the PDFs"), self.pv['output_dir'].get().strip())
        if d:
            self.pv['output_dir'].set(os.path.normpath(d))
            self.pv['output_mode'].set('custom')
            self._pref_changed()

    def browse_excel_dir(self):
        d = self._pick_folder(T("Διαλέξτε φάκελο των Excel", "Choose the folder of the Excel files"), self.pv['excel_dir'].get().strip())
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
        d = self._pick_folder(T("Διαλέξτε τον φάκελο δεδομένων", "Choose the data folder"), DATA_DIR)
        if d:
            self._switch_data_dir(os.path.normpath(d))

    def _switch_data_dir(self, new):
        old = DATA_DIR
        if _same_path(new, old):
            return
        if _is_inside(new, os.path.join(old, ARCHIVE_NAME)):
            messagebox.showerror(T("Φάκελος δεδομένων", "Data folder"), T("Δεν μπορείτε να διαλέξετε φάκελο μέσα στο αρχείο PDF.",
                                                        "You cannot choose a folder inside the PDF archive."))
            return
        copy = False
        if data_exists(new):
            if not messagebox.askyesno(
                    T("Φάκελος δεδομένων", "Data folder"),
                    T(f"Ο φάκελος\n{new}\nπεριέχει ήδη δεδομένα του προγράμματος.\n\n"
                      "Να χρησιμοποιηθούν αυτά; (Τα τωρινά δεδομένα δεν διαγράφονται.)",
                      f"The folder\n{new}\nalready contains data from the program.\n\n"
                      "Use it? (Your current data is not deleted.)"), default='yes'):
                return
        elif data_exists(old):
            ans = messagebox.askyesnocancel(
                T("Φάκελος δεδομένων", "Data folder"),
                T("Να αντιγραφούν τα υπάρχοντα δεδομένα (στοιχεία ξενοδοχείου, ιστορικό, αρχείο PDF) στον "
                  "νέο φάκελο;\n\nΝαι = αντιγραφή (τα παλιά αρχεία μένουν και στον παλιό φάκελο)\n"
                  "Όχι = ξεκίνημα από το μηδέν στον νέο φάκελο\nΑκύρωση = καμία αλλαγή",
                  "Copy the existing data (hotel details, history, PDF archive) to the "
                  "new folder?\n\nYes = copy (the old files also stay in the old folder)\n"
                  "No = start from scratch in the new folder\nCancel = no change"), default='yes')
            if ans is None:
                return
            copy = ans
        try:
            os.makedirs(new, exist_ok=True)
            if copy:
                copy_data_files(old, new)
        except Exception as e:
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Δεν ήταν δυνατή η αντιγραφή στον νέο φάκελο:\n{e}",
                                             f"Could not copy to the new folder:\n{e}"))
            return
        set_data_dir(new)
        if not save_app_config({'data_dir': '' if _same_path(new, DEFAULT_DATA_DIR) else new}):
            set_data_dir(old)
            messagebox.showerror(T("Σφάλμα", "Error"), T("Δεν μπορώ να θυμηθώ τη νέα θέση, γιατί ο φάκελος του προγράμματος "
                                             "δεν επιτρέπει εγγραφή. Η αλλαγή ακυρώθηκε.",
                                             "I can't remember the new location because the program folder "
                                             "does not allow writing. The change was cancelled."))
            return
        # preferences live in the data folder: use the ones found there, or carry the current ones over
        if os.path.exists(PREFS_JSON):
            self.prefs = load_prefs()
            self._sync_pref_vars()
            if self.prefs['language'] != LANG:
                self._apply_language()
        else:
            self._save_prefs()
        if os.path.exists(SETTINGS_CSV):
            self.current_settings = load_hotel_settings()
            for k in SETTING_KEYS:
                self.set_vars[k].set(self.current_settings.get(k, ''))
        self.data_dir_var.set(DATA_DIR)
        self.refresh_history()
        self._after_prefs_change()
        self._status(T(f"✓ Ο φάκελος δεδομένων είναι τώρα: {DATA_DIR}", f"✓ The data folder is now: {DATA_DIR}"))

    def backup_data(self):
        name = f"ELSTAT_backup_{datetime.date.today().strftime('%Y-%m-%d')}.zip"
        fp = filedialog.asksaveasfilename(title=T("Αντίγραφο ασφαλείας", "Backup"), defaultextension=".zip",
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
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Δεν ήταν δυνατή η δημιουργία αντιγράφου:\n{e}",
                                             f"Could not create the backup:\n{e}"))
            return
        self._status(T(f"✓ Αντίγραφο ασφαλείας: {fp}", f"✓ Backup: {fp}"))

    def reset_prefs(self):
        if not messagebox.askyesno(T("Επαναφορά ρυθμίσεων", "Reset settings"),
                                   T("Να γυρίσουν όλες οι ρυθμίσεις στις προεπιλογές;\n"
                                     "(Δεν διαγράφονται δεδομένα ή ιστορικό.)",
                                     "Reset all settings to their defaults?\n"
                                     "(No data or history is deleted.)"), default='no'):
            return
        keep_template = self.prefs.get('template_pdf', '')
        self.prefs = dict(DEFAULT_PREFS, template_pdf=keep_template, language=self.prefs['language'])
        self._sync_pref_vars()
        self._save_prefs()
        self._after_prefs_change()
        self._status(T("Οι ρυθμίσεις επανήλθαν στις προεπιλογές.", "Settings were reset to their defaults."))

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
        if self._rebuilding or not hasattr(self, 'print_btn'):
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
        self.afm_counter_var.set(T(f"{n}/9 ψηφία", f"{n}/9 digits") + ("  ✓" if validate_afm(s['afm']) else ""))
        if m['month'].isdigit() and 1 <= int(m['month']) <= 12 and m['year']:
            self.month_name_var.set(f"= {month_name(int(m['month']))} {m['year']}")
        else:
            self.month_name_var.set("")

        # overall verdict
        if excel_ok:
            errors = validate_inputs(s, m, excel, template, out)
        else:   # month fields cannot be judged before the Excel is chosen
            errors = [T("Επιλέξτε πρώτα το αρχείο Excel από το Protel (Βήμα 1).",
                        "First choose the Excel file from Protel (Step 1).")] + validate_hotel(s)
        if errors:
            shown = errors[:5]
            more = T(f"\n… και {len(errors) - 5} ακόμη", f"\n… and {len(errors) - 5} more") if len(errors) > 5 else ""
            self.check_var.set(T("Για να δημιουργηθεί το δελτίο, διορθώστε:\n",
                                 "To create the form, please fix:\n") + "\n".join("• " + e for e in shown) + more)
            self.check_label.configure(style='StatusWarn.TLabel' if not excel_ok else 'StatusBad.TLabel')
        else:
            self.check_var.set(T("✓ Όλα είναι έτοιμα. Πατήστε «Δημιουργία PDF» ή «Δημιουργία & Εκτύπωση».",
                                 "✓ Everything is ready. Click “Create PDF” or “Create & Print”."))
            self.check_label.configure(style='StatusGood.TLabel')

        for btn in (self.run_btn, self.print_btn):
            if errors or self._busy:
                btn.state(['disabled'])
            else:
                btn.state(['!disabled'])

    def _set_busy(self, busy):
        self._busy = busy
        self.config(cursor='watch' if busy else '')
        self.run_btn.configure(text=T("Παρακαλώ περιμένετε…", "Please wait…") if busy else T("Δημιουργία PDF", "Create PDF"))
        self.print_btn.configure(text=T("Παρακαλώ περιμένετε…", "Please wait…") if busy else T("Δημιουργία & Εκτύπωση",
                                                                            "Create & Print"))
        self.refresh_validation()
        self.update_idletasks()

    # ----------------------------------------------------------- file choices
    def _start_dir(self):
        d = self.current_settings.get('last_dir', '')
        return d if d and os.path.isdir(d) else BASE_DIR

    def browse_excel(self):
        fp = filedialog.askopenfilename(title=T("Διαλέξτε το Excel από το Protel",
                                                "Choose the Excel file from Protel"), initialdir=self._excel_start_dir(),
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
            self._show_period(lambda: T("⚠ Το αρχείο δεν βρέθηκε.", "⚠ The file was not found."),
                              'CardBad.TLabel')
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
            self._show_period(
                lambda: T(f"✓ Το Excel αφορά τον μήνα {month_name(month)} {year}. "
                          "Τα στοιχεία του μήνα συμπληρώθηκαν αυτόματα - πληκτρολογήστε μόνο τον αριθμό εργαζομένων.",
                          f"✓ The Excel is for {month_name(month)} {year}. "
                          "The month details were filled in automatically - just type the number of employees."),
                'CardGood.TLabel')
        else:
            self._show_period(
                lambda: T("⚠ Δεν βρέθηκε ολόκληρος μήνας στο Excel. Συμπληρώστε τα στοιχεία του Βήματος 2 "
                          "με το χέρι.",
                          "⚠ No complete month was found in the Excel. Fill in the Step 2 details "
                          "by hand."),
                'CardWarn.TLabel')

    def _show_period(self, make_text, style):
        """Message under the Excel chooser. Kept as a function so it can be re-worded after a language change."""
        self._period_msg = (make_text, style)
        self._render_period()

    def _render_period(self):
        make_text, style = self._period_msg
        self.period_var.set(make_text())
        self.period_label.configure(style=style)

    def use_recent_excel(self):
        """Picks the newest Excel in the folder used last time (or next to the program)."""
        fp = find_recent_excel(self._excel_start_dir())
        if not fp:
            messagebox.showinfo(T("Δεν βρέθηκε αρχείο", "File not found"),
                                T(f"Δεν υπάρχει αρχείο Excel στον φάκελο:\n{self._excel_start_dir()}\n\n"
                                  "Πατήστε «Επιλογή Excel…» για να το διαλέξετε.",
                                  f"There is no Excel file in the folder:\n{self._excel_start_dir()}\n\n"
                                  "Click “Choose Excel…” to select one."))
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
                                     "✓" if known else T("⚠ λείπει από το πρόγραμμα",
                                                         "⚠ missing from the program")),
                                     tags=('ok' if known else 'bad',))
        self.preview_tree.configure(height=max(3, min(len(rows), 7)))
        text = (T(f"Βρέθηκαν {len(rows)} χώρες · {sum(r[1] for r in rows)} πελάτες · "
                  f"{sum(r[2] for r in rows)} διανυκτερεύσεις",
                  f"Found {len(rows)} countries · {sum(r[1] for r in rows)} guests · "
                  f"{sum(r[2] for r in rows)} overnight stays"))
        if bad:
            text += T(f"   ⚠ {bad} δεν αναγνωρίστηκαν και ΔΕΝ θα μπουν στο PDF",
                      f"   ⚠ {bad} not recognised and will NOT be included in the PDF")
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
        ttk.Label(top, text=T("Έτος:", "Year:")).pack(side="left")
        self.hist_year_cb = ttk.Combobox(top, textvariable=self.hist_year_var, state="readonly", width=7)
        self.hist_year_cb.pack(side="left", padx=(6, 12))
        self.hist_year_cb.bind('<<ComboboxSelected>>', lambda e: self.update_dashboard())
        ttk.Label(top, text=T("Χώρα:", "Country:")).pack(side="left")
        self.hist_country_cb = ttk.Combobox(top, textvariable=self.hist_country_var, state="readonly", width=20)
        self.hist_country_cb.pack(side="left", padx=(6, 12))
        self.hist_country_cb.bind('<<ComboboxSelected>>', lambda e: self._country_changed())
        ttk.Radiobutton(top, text=T("Διανυκτερεύσεις", "Overnight stays"), value='nights', variable=self.hist_metric_var,
                        command=self.update_dashboard).pack(side="left")
        ttk.Radiobutton(top, text=T("Πελάτες", "Guests"), value='clients', variable=self.hist_metric_var,
                        command=self.update_dashboard).pack(side="left", padx=(10, 0))
        ttk.Button(top, text=T("Άνοιγμα CSV", "Open CSV"),
                   command=lambda: os.path.exists(HISTORY_CSV) and open_file(HISTORY_CSV)).pack(side="right")
        ttk.Button(top, text=T("Εξαγωγή Excel…", "Export Excel…"), command=self.hist_export).pack(side="right", padx=(0, 8))
        ttk.Button(top, text=T("Ανανέωση", "Refresh"), command=self.refresh_history).pack(side="right", padx=(0, 8))

        info_row = ttk.Frame(tab)
        info_row.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        info_row.columnconfigure(0, weight=1)
        self.hist_info = tk.StringVar()
        self.hist_info_lbl = ttk.Label(info_row, textvariable=self.hist_info, style='Hint.TLabel', justify="left")
        self.hist_info_lbl.grid(row=0, column=0, sticky="w")
        autowrap(self.hist_info_lbl, tab, margin=260, minimum=200)
        ttk.Button(info_row, text=T("Σύμπτυξη", "Collapse"), command=lambda: self._tree_expand(False)).grid(
            row=0, column=1, padx=(8, 0))
        ttk.Button(info_row, text=T("Ανάπτυξη", "Expand"), command=lambda: self._tree_expand(True)).grid(
            row=0, column=2, padx=(6, 0))

        kp = ttk.Frame(tab)
        kp.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 8))
        for i in range(4):
            kp.columnconfigure(i, weight=1, uniform='kpi')
        self.kpi_clients = StatCard(kp, T("ΠΕΛΑΤΕΣ", "GUESTS"))
        self.kpi_clients.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.kpi_nights = StatCard(kp, T("ΔΙΑΝΥΚΤΕΡΕΥΣΕΙΣ", "OVERNIGHT STAYS"))
        self.kpi_nights.grid(row=0, column=1, sticky="ew", padx=4)
        self.kpi_avg = StatCard(kp, T("ΜΕΣΗ ΔΙΑΜΟΝΗ (ΝΥΧΤΕΣ)", "AVERAGE STAY (NIGHTS)"))
        self.kpi_avg.grid(row=0, column=2, sticky="ew", padx=4)
        self.kpi_share = StatCard(kp, T("ΑΛΛΟΔΑΠΟΙ", "FOREIGN GUESTS"))
        self.kpi_share.grid(row=0, column=3, sticky="ew", padx=(4, 0))

        self.hist_chart = HistoryChart(tab, on_pick=self._chart_pick, height=170)
        self.hist_chart.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 8))

        pane = ttk.PanedWindow(tab, orient="horizontal")
        pane.grid(row=4, column=0, sticky="nsew", padx=6, pady=(0, 8))

        frame = ttk.Frame(pane)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        pane.add(frame, weight=3)
        self.hist_tree = ttk.Treeview(frame, columns=('c', 'n', 'a', 's', 'p'), show='tree headings', height=6)
        self.hist_tree.heading('#0', text=T('Έτος / Μήνας / Χώρα', 'Year / Month / Country'), anchor='w')
        self.hist_tree.heading('c', text=T('Πελάτες', 'Guests'))
        self.hist_tree.heading('n', text=T('Διανυκτ.', 'Nights'))
        self.hist_tree.heading('a', text=T('Μέση δ.', 'Avg'))
        self.hist_tree.heading('s', text=T('% πελατών', '% guests'))
        self.hist_tree.heading('p', text='PDF')
        self.hist_tree.column('#0', width=200, minwidth=140, stretch=True)
        self.hist_tree.column('c', width=70, minwidth=55, anchor='e', stretch=False)
        self.hist_tree.column('n', width=80, minwidth=60, anchor='e', stretch=False)
        self.hist_tree.column('a', width=60, minwidth=50, anchor='e', stretch=False)
        self.hist_tree.column('s', width=75, minwidth=60, anchor='e', stretch=False)
        self.hist_tree.column('p', width=45, minwidth=40, anchor='center', stretch=False)
        self.hist_tree.tag_configure('year', font=(FONT, 10, 'bold'), background='#dbe7fb')
        self.hist_tree.tag_configure('month', font=(FONT, 10, 'bold'), background='#eef4fd')
        self.hist_tree.tag_configure('country', background=CARD)
        self.hist_tree.tag_configure('hl', background='#fef3c7')
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.hist_tree.yview)
        self.hist_tree.configure(yscrollcommand=sb.set)
        self.hist_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.hist_tree.bind('<<TreeviewSelect>>', self._on_hist_select)
        self.hist_tree.bind('<Return>', lambda e: self.hist_open())
        self.hist_tree.bind('<Double-1>', self._hist_double)
        self.hist_tree.bind('<Delete>', lambda e: self.hist_delete())

        side = ttk.Frame(pane, style='Card.TFrame')
        side.columnconfigure(0, weight=1)
        side.rowconfigure(2, weight=1)
        pane.add(side, weight=2)
        ttk.Label(side, textvariable=self.top_title_var, style='CardBold.TLabel').grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 0))
        ttk.Label(side, text=T("Διπλό κλικ σε χώρα = φίλτρο", "Double-click a country = filter"),
                  style='CardHint.TLabel').grid(row=1, column=0, sticky="w", padx=10, pady=(0, 6))
        self.top_tree = ttk.Treeview(side, columns=('c', 's'), show='tree headings', height=6)
        self.top_tree.heading('#0', text=T('Χώρα', 'Country'), anchor='w')
        self.top_tree.heading('c', text=T('Πελάτες', 'Guests'))
        self.top_tree.heading('s', text=T('Μερίδιο', 'Share'), anchor='w')
        self.top_tree.column('#0', width=130, minwidth=90, stretch=True)
        self.top_tree.column('c', width=60, minwidth=50, anchor='e', stretch=False)
        self.top_tree.column('s', width=130, minwidth=90, anchor='w', stretch=False)
        self.top_tree.tag_configure('hl', background='#fef3c7')
        self.top_tree.grid(row=2, column=0, sticky="nsew")
        self.top_tree.bind('<Double-1>', self._top_double)

        act = ttk.Frame(tab)
        act.grid(row=5, column=0, sticky="ew", padx=6, pady=(0, 12))
        self.month_btns, self.year_btns = [], []
        ttk.Label(act, text=T("Δελτίο μήνα:", "Monthly form:"), style='Hint.TLabel').pack(side="left", padx=(0, 6))
        for text, cmd, style in ((T("Άνοιγμα", "Open"), self.hist_open, 'TButton'),
                                 (T("Αποθήκευση ως…", "Save as…"), self.hist_save_as, 'TButton'),
                                 (T("Εκτύπωση", "Print"), self.hist_print, 'PrintSm.TButton')):
            b = ttk.Button(act, text=text, command=cmd, style=style)
            b.pack(side="left", padx=(0, 6))
            self.month_btns.append(b)
        ttk.Separator(act, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Label(act, text=T("Σύνοψη έτους:", "Year summary:"), style='Hint.TLabel').pack(side="left", padx=(0, 6))
        for text, cmd, style in ((T("Αποθήκευση PDF…", "Save PDF…"), lambda: self.hist_year_pdf('save'), 'TButton'),
                                 (T("Εκτύπωση", "Print"), lambda: self.hist_year_pdf('print'), 'PrintSm.TButton')):
            b = ttk.Button(act, text=text, command=cmd, style=style)
            b.pack(side="left", padx=(0, 6))
            self.year_btns.append(b)
        self.del_btn = ttk.Button(act, text=T("Διαγραφή μήνα", "Delete month"), command=self.hist_delete, style='Danger.TButton')
        self.del_btn.pack(side="right")

        # right-click menu
        self.hist_menu = tk.Menu(self, tearoff=0)
        self.hist_menu.add_command(label=T("Άνοιγμα PDF", "Open PDF"), command=self.hist_open)
        self.hist_menu.add_command(label=T("Αποθήκευση ως…", "Save as…"), command=self.hist_save_as)
        self.hist_menu.add_command(label=T("Εκτύπωση δελτίου", "Print form"), command=self.hist_print)
        self.hist_menu.add_separator()
        self.hist_menu.add_command(label=T("Σύνοψη έτους - αποθήκευση PDF…", "Year summary - save PDF…"), command=lambda: self.hist_year_pdf('save'))
        self.hist_menu.add_command(label=T("Σύνοψη έτους - εκτύπωση", "Year summary - print"), command=lambda: self.hist_year_pdf('print'))
        self.hist_menu.add_separator()
        self.hist_menu.add_command(label=T("Διαγραφή μήνα", "Delete month"), command=self.hist_delete)
        self.hist_tree.bind('<Button-3>', self._hist_popup)
        if sys.platform == 'darwin':
            self.hist_tree.bind('<Button-2>', self._hist_popup)

    # ---- history: new helpers
    def _all_label(self):
        return T("Όλες οι χώρες", "All countries")

    def _country_filter(self):
        v = self.hist_country_var.get()
        return v if v in self._countries else None

    def _country_changed(self):
        self.refresh_history()

    def _tree_expand(self, state):
        def walk(parent=''):
            for i in self.hist_tree.get_children(parent):
                yield i
                yield from walk(i)
        for i in walk():
            if self.hist_tree.get_children(i):
                self.hist_tree.item(i, open=state)

    def _hist_double(self, event):
        iid = self.hist_tree.identify_row(event.y)
        if iid and iid[0] == 'm':
            self.hist_open()

    def _top_double(self, _event=None):
        cur = self.top_tree.selection()
        if not cur:
            return
        country = self.top_tree.item(cur[0], 'text')
        if self._country_filter() and norm(self._country_filter()) == norm(country):
            self.hist_country_var.set(self._all_label())      # second double-click clears the filter
        elif country in self._countries:
            self.hist_country_var.set(country)
        self.refresh_history()

    def hist_export(self):
        if not self._hist:
            messagebox.showinfo(T("Ιστορικό", "History"), T("Δεν υπάρχουν δεδομένα για εξαγωγή.", "There is no data to export."))
            return
        fp = filedialog.asksaveasfilename(title=T("Εξαγωγή ιστορικού", "Export history"), defaultextension=".xlsx",
                                          initialdir=self._pdf_start_dir(), initialfile="ELSTAT_History.xlsx",
                                          filetypes=[("Excel", "*.xlsx")])
        if not fp:
            return
        try:
            export_history_excel(self._hist, fp)
        except PermissionError:
            messagebox.showerror(T("Σφάλμα", "Error"), T("Δεν μπορώ να γράψω το αρχείο. Κλείστε το αν είναι ανοιχτό και δοκιμάστε ξανά.",
                                             "I can't write the file. Close it if it is open and try again."))
            return
        except Exception as e:
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Δεν ήταν δυνατή η εξαγωγή:\n{e}", f"Could not export:\n{e}"))
            return
        self._status(T(f"✓ Το ιστορικό εξήχθη: {fp}", f"✓ History exported: {fp}"))
        open_file(fp)

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

        # country filter list (most guests first)
        totals = {}
        for months in hist.values():
            for rows in months.values():
                for country, c, _n in rows:
                    totals[country] = totals.get(country, 0) + c
        self._countries = sorted(totals, key=lambda k: -totals[k])
        self.hist_country_cb.configure(values=[self._all_label()] + self._countries)
        if self.hist_country_var.get() not in self._countries:
            self.hist_country_var.set(self._all_label())
        flt = self._country_filter()
        flt_key = norm(flt) if flt else None

        def avg(c, n):
            return fmt_avg(n, c)

        def share(part, whole):
            return f"{fmt_dec(part / whole * 100)}%" if whole else ""

        months_total = 0
        for year in sorted(hist, reverse=True):
            ycl = sum(c for rows in hist[year].values() for _, c, _ in rows)
            yni = sum(n for rows in hist[year].values() for _, _, n in rows)
            yid = f"y{year}"
            tree.insert('', 'end', iid=yid, text=str(year), values=(ycl, yni, avg(ycl, yni), "", ""),
                        tags=('year',), open=(year == max(hist)) if first_time else (yid in was_open))
            for month in sorted(hist[year], reverse=True):
                rows = hist[year][month]
                mcl, mni = sum(r[1] for r in rows), sum(r[2] for r in rows)
                months_total += 1
                mid = f"m{year}-{month}"
                name = month_name(month) if 1 <= month <= 12 else str(month)
                tree.insert(yid, 'end', iid=mid, text=name,
                            values=(mcl, mni, avg(mcl, mni), share(mcl, ycl),
                                    "✓" if os.path.isfile(archive_path(year, month)) else ""),
                            tags=('month',), open=(mid in was_open))
                for i, (country, c, n) in enumerate(sorted(rows, key=lambda r: -r[1])):
                    hl = flt_key is not None and norm(country) == flt_key
                    tree.insert(mid, 'end', iid=f"c{year}-{month}-{i}", text=country,
                                values=(c, n, avg(c, n), share(c, mcl), ""),
                                tags=('hl' if hl else 'country',))
        if keep_sel and tree.exists(keep_sel):
            tree.selection_set(keep_sel)

        years = [str(y) for y in sorted(hist, reverse=True)]
        self.hist_year_cb.configure(values=years)
        if self.hist_year_var.get() not in years:
            self.hist_year_var.set(years[0] if years else "")
        self._hist_info_base = (T(f"{months_total} μήνες καταχωρημένοι. Κλικ σε μήνα (πίνακας ή γράφημα) = επιλογή · "
                                  "διπλό κλικ = άνοιγμα δελτίου.",
                                  f"{months_total} months recorded. Click a month (table or chart) = select · "
                                  "double-click = open the form.") if hist else
                                T("Δεν υπάρχει ιστορικό ακόμη - δημιουργείται αυτόματα με το πρώτο δελτίο.",
                                  "No history yet - it is created automatically with the first form."))
        self.hist_info.set(self._hist_info_base)
        self.update_dashboard()
        self._update_hist_buttons()
        self.update_reminder()

    def update_dashboard(self, *_):
        """Cards, bar chart and top countries for the chosen year (and country filter)."""
        try:
            year = int(self.hist_year_var.get())
        except ValueError:
            year = None
        by_nights = self.hist_metric_var.get() == 'nights'
        cards = (self.kpi_clients, self.kpi_nights, self.kpi_avg, self.kpi_share)
        if year is None or year not in self._hist:
            for card in cards:
                card.set("–", "")
            self.hist_chart.set_data([0] * 12, T("Δεν υπάρχουν δεδομένα ακόμη", "No data yet"))
            self.top_tree.delete(*self.top_tree.get_children())
            self.top_title_var.set("")
            return
        flt = self._country_filter()
        view = filter_hist(self._hist, flt)
        st = year_stats(view, year)
        full = year_stats(self._hist, year)

        def delta(cur, prev):
            if not prev:
                return "", MUTED
            pct = (cur - prev) / prev * 100
            arrow = "▲" if pct > 0.05 else ("▼" if pct < -0.05 else "■")
            pct_txt = fmt_dec(pct, signed=True)
            text = T(f"{arrow} {pct_txt}% έναντι {year - 1} (ίδιοι μήνες)",
                     f"{arrow} {pct_txt}% vs {year - 1} (same months)")
            return text, (GOOD if pct > 0.05 else BAD if pct < -0.05 else MUTED)

        cmp_ = st['cmp']
        t, col = delta(cmp_['cur_cl'], cmp_['prev_cl']) if cmp_ else ("", MUTED)
        self.kpi_clients.set(fmt_int(st['clients']), t or T(f"{len(st['months'])} μήνες με δελτίο",
                                                            f"{len(st['months'])} months with a form"), col)
        t, col = delta(cmp_['cur_ni'], cmp_['prev_ni']) if cmp_ else ("", MUTED)
        self.kpi_nights.set(fmt_int(st['nights']), t or T("συνολικά στο έτος", "total for the year"), col)
        self.kpi_avg.set(fmt_avg(st['nights'], st['clients']), T("διανυκτερεύσεις ανά πελάτη",
                                                                 "overnight stays per guest"))
        greece = norm('ΕΛΛΑΔΑ')
        if flt:
            pct = st['clients'] / full['clients'] * 100 if full['clients'] else 0
            self.kpi_share.set(f"{fmt_dec(pct)}%", T(f"των πελατών του {year} ({flt})",
                                                    f"of the {year} guests ({flt})"))
        else:
            gr = sum(c for rows in self._hist[year].values() for k, c, _ in rows if norm(k) == greece)
            foreign = full['clients'] - gr
            pct = foreign / full['clients'] * 100 if full['clients'] else 0
            self.kpi_share.set(f"{fmt_dec(pct)}%", T(f"{fmt_int(foreign)} αλλοδαποί · {fmt_int(gr)} Ελλάδα",
                                                    f"{fmt_int(foreign)} foreign · {fmt_int(gr)} Greece"))

        gap_txt = (T("   ·   Μήνες χωρίς δελτίο μέσα στο ", "   ·   Months without a form in ") + str(year) + ": " +
                   ", ".join(month_short(m) for m in full['gaps'])) if full['gaps'] else ""
        flt_txt = (T(f"   ·   Φίλτρο: {flt}", f"   ·   Filter: {flt}")) if flt else ""
        self.hist_info.set(getattr(self, '_hist_info_base', "") + flt_txt + gap_txt)

        def series(y):
            return [sum((r[2] if by_nights else r[1]) for r in view.get(y, {}).get(m, [])) for m in range(1, 13)]
        per_month = series(year)
        prev = series(year - 1) if (year - 1) in view else None
        metric = (T('Διανυκτερεύσεις', 'Overnight stays') if by_nights else T('Πελάτες', 'Guests'))
        sel = self._hist_sel()
        self.hist_chart.set_data(
            per_month, f"{metric} · {year}" + (f" · {flt}" if flt else ""),
            color=ACCENT if by_nights else TEAL, selected=sel[1] if sel and sel[0] == year else None,
            prev=prev, prev_label=str(year - 1))

        # top countries of the year
        by_country = {}
        for rows in self._hist[year].values():
            for country, c, _n in rows:
                by_country[country] = by_country.get(country, 0) + c
        total = sum(by_country.values())
        top = sorted(by_country.items(), key=lambda kv: -kv[1])[:10]
        peak = top[0][1] if top else 1
        self.top_tree.delete(*self.top_tree.get_children())
        for i, (country, c) in enumerate(top):
            bar = "█" * max(1, round(c / peak * 10))
            hl = flt is not None and norm(country) == norm(flt)
            self.top_tree.insert('', 'end', iid=f"t{i}", text=country,
                                 values=(fmt_int(c), f"{bar} {fmt_dec(c / total * 100)}%"),
                                 tags=('hl',) if hl else ())
        self.top_title_var.set(T(f"Κορυφαίες χώρες · {year}", f"Top countries · {year}"))

    # ---- reminder banner
    def update_reminder(self):
        r = None if (self._reminder_dismissed or not self.prefs['show_reminder']) \
            else missing_month_reminder(self._hist)
        if r:
            self.reminder_label.configure(
                text=T(f"Υπενθύμιση: δεν υπάρχει ακόμη δελτίο για τον {month_name(r[1])} {r[0]}. "
                       "Διαλέξτε το Excel του μήνα στο Βήμα 1 για να το δημιουργήσετε.",
                       f"Reminder: there is no form yet for {month_name(r[1])} {r[0]}. "
                       "Choose the month's Excel in Step 1 to create it."))
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
            messagebox.showerror(T("Ιστορικό", "History"), T("Δεν βρέθηκαν στοιχεία για αυτόν τον μήνα.",
                                               "No data was found for this month."))
            return None
        s, _m, _excel, template, _out = self._collect()
        errors = validate_hotel(s)
        if not template or not os.path.isfile(template):
            errors.append(T("Το άδειο PDF (template) δεν βρέθηκε - δείτε την καρτέλα «Στοιχεία ξενοδοχείου».",
                            "The blank PDF (template) was not found - see the “Hotel details” tab."))
        if errors:
            messagebox.showerror(T("Δεν μπορώ να ξαναφτιάξω το δελτίο", "I can't rebuild the form"), "\n".join("• " + e for e in errors))
            return None
        rec = load_month_records().get((year, month), {})
        employees = rec.get('employees', '')
        name = f"{month_name(month)} {year}"
        if not employees:
            employees = simpledialog.askstring(
                T("Σύνολο εργαζομένων", "Total employees"),
                T(f"Το δελτίο του {name} δημιουργήθηκε πριν αποθηκεύεται ο αριθμός εργαζομένων.\n"
                  "Πληκτρολογήστε το Σύνολο Εργαζομένων εκείνου του μήνα:",
                  f"The form for {name} was created before the number of employees was saved.\n"
                  "Type the Total employees for that month:"), parent=self)
            if employees is None:
                return None
            employees = employees.strip()
        last = calendar.monthrange(year, month)[1]
        mon = dict(year=str(year), month=str(month), date_to=f"{last:02d}/{month:02d}/{year}",
                   days_of_month=str(rec.get('days') or last), employees=employees)
        problems = validate_month(mon)
        if problems:
            messagebox.showerror(T("Λάθος στοιχεία", "Invalid details"), "\n".join("• " + e for e in problems))
            return None
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            process_elstat_report(None, template, path, s, mon, rows_override=rows, save_history=False)
            save_month_record(year, month, employees, mon['days_of_month'], mon['date_to'])
        except PermissionError:
            messagebox.showerror(T("Σφάλμα", "Error"), T("Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό και δοκιμάστε ξανά.",
                                             "I can't write the PDF. Close it if it is open and try again."))
            return None
        except Exception as e:
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"))
            return None
        finally:
            self.config(cursor='')
        self._status(T(f"Το δελτίο {name} ξαναδημιουργήθηκε από το ιστορικό.",
                       f"The form for {name} was rebuilt from the history."))
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
            self._status(T(f"Άνοιξε το δελτίο {month_name(sel[1])} {sel[0]}.",
                           f"Opened the form for {month_name(sel[1])} {sel[0]}."))

    def hist_save_as(self):
        sel = self._month_from_selection()
        path = sel and self._ensure_month_pdf(*sel)
        if not path:
            return
        fp = filedialog.asksaveasfilename(title=T("Αποθήκευση δελτίου ως", "Save form as"), defaultextension=".pdf",
                                          initialdir=self._pdf_start_dir(),
                                          initialfile=os.path.basename(path), filetypes=[("PDF Files", "*.pdf")])
        if not fp:
            return
        try:
            shutil.copy2(path, fp)
        except Exception as e:
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Δεν ήταν δυνατή η αποθήκευση:\n{e}", f"Could not save:\n{e}"))
            return
        self._status(T(f"✓ Αποθηκεύτηκε: {fp}", f"✓ Saved: {fp}"))

    def hist_print(self):
        sel = self._month_from_selection()
        path = sel and self._ensure_month_pdf(*sel)
        if path:
            self._status(print_file(path)[1])

    def hist_delete(self):
        sel = self._month_from_selection()
        if not sel:
            return
        name = f"{month_name(sel[1])} {sel[0]}"
        if not messagebox.askyesno(T("Διαγραφή μήνα", "Delete month"),
                                   T(f"Να διαγραφεί ο μήνας {name} από το ιστορικό;\n"
                                     "(Θα διαγραφεί και το αρχειοθετημένο PDF του. Το αρχείο που αποθηκεύσατε "
                                     "εσείς δεν θα πειραχτεί.)",
                                     f"Delete the month {name} from the history?\n"
                                     "(Its archived PDF will also be deleted. The file you saved "
                                     "yourself will not be touched.)"), default='no'):
            return
        try:
            delete_month(*sel)
        except Exception as e:
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Δεν ήταν δυνατή η διαγραφή:\n{e}", f"Could not delete:\n{e}"))
            return
        self._status(T(f"Ο μήνας {name} διαγράφηκε από το ιστορικό.",
                       f"The month {name} was deleted from the history."))
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
            fp = filedialog.asksaveasfilename(title=T("Αποθήκευση σύνοψης έτους", "Save year summary"), defaultextension=".pdf",
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
            messagebox.showerror(T("Σφάλμα", "Error"), T("Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό και δοκιμάστε ξανά.",
                                             "I can't write the PDF. Close it if it is open and try again."))
            return
        except Exception as e:
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"))
            return
        if mode == 'save':
            self._status(T(f"✓ Η σύνοψη του {year} αποθηκεύτηκε: {fp}", f"✓ The {year} summary was saved: {fp}"))
            open_file(fp)
        else:
            self._status(print_file(fp)[1])

    def browse_pdf(self):
        fp = filedialog.askopenfilename(title=T("Διαλέξτε το άδειο PDF της ΕΛΣΤΑΤ",
                                                "Choose the blank ELSTAT PDF"), initialdir=BASE_DIR,
                                        filetypes=[("PDF Files", "*.pdf")])
        if fp:
            self.pdf_var.set(fp)
            self._remember_template(fp)

    def browse_output(self):
        cur = self.out_var.get().strip()
        fp = filedialog.asksaveasfilename(title=T("Αποθήκευση PDF ως", "Save PDF as"), defaultextension=".pdf",
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
            messagebox.showerror(T("Λάθος στοιχεία", "Invalid details"), "\n".join("• " + e for e in errors))
            return

        info = period_as_month(read_excel_period(excel))
        if info and (str(info[0]), str(int(m['month']))) != (m['year'], str(info[1])):
            if not messagebox.askyesno(
                    T("Διαφορετική περίοδος", "Different period"),
                    T(f"Το Excel αφορά τον μήνα {month_name(info[1])} {info[0]}, ενώ στη φόρμα έχετε "
                      f"{month_name(int(m['month']))} {m['year']}.\n\nΝα συνεχίσω παρ' όλα αυτά;",
                      f"The Excel is for {month_name(info[1])} {info[0]}, but the form has "
                      f"{month_name(int(m['month']))} {m['year']}.\n\nContinue anyway?"),
                    default='no'):
                return

        if self.prefs['confirm_overwrite'] and os.path.exists(out):
            if not messagebox.askyesno(T("Υπάρχει ήδη αρχείο", "File already exists"),
                                       T(f"Το αρχείο υπάρχει ήδη:\n{out}\n\nΝα αντικατασταθεί;",
                                         f"The file already exists:\n{out}\n\nReplace it?"), default='no'):
                return

        saved = dict(s, last_dir=os.path.dirname(excel), last_employees=m['employees'])
        try:
            save_hotel_settings(saved)
        except Exception as e:
            messagebox.showwarning(T("Προσοχή", "Warning"), T(f"Δεν ήταν δυνατή η αποθήκευση των στοιχείων στον φάκελο "
                                                f"δεδομένων:\n{DATA_DIR}\n\n{e}",
                                                f"Could not save the details to the data "
                                                f"folder:\n{DATA_DIR}\n\n{e}"))
        self.current_settings = saved
        if template != (self.prefs.get('template_pdf') or DEFAULT_TEMPLATE):
            self._remember_template(template)

        self._set_busy(True)
        try:
            os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
            result = process_elstat_report(excel, template, out, s, m)
        except PermissionError:
            messagebox.showerror(T("Σφάλμα", "Error"), T("Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό σε άλλο "
                                             "πρόγραμμα και δοκιμάστε ξανά.",
                                             "I can't write the PDF. Close it if it is open in another "
                                             "program and try again."))
            return
        except Exception as e:
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"))
            return
        finally:
            self._set_busy(False)

        warnings_txt = []
        if result['unmapped']:
            lines = ", ".join(T(f"{c} ({p} πελάτες / {n} διανυκτ.)", f"{c} ({p} guests / {n} nights)") for c, p, n in result['unmapped'])
            warnings_txt.append(T("Αυτές οι χώρες ΔΕΝ μπήκαν στο PDF (λείπουν από το πρόγραμμα): ",
                                  "These countries were NOT included in the PDF (missing from the program): ") + lines)
        if result['excel_total'] and tuple(result['excel_total']) != tuple(result['computed_total']):
            warnings_txt.append(T(f"Το γενικό σύνολο του PDF {result['computed_total']} διαφέρει από το σύνολο "
                                  f"του Excel {result['excel_total']}. Ελέγξτε το πριν το υποβάλετε!",
                                  f"The PDF grand total {result['computed_total']} differs from the Excel "
                                  f"total {result['excel_total']}. Check it before submitting!"))

        try:
            gy, gm = int(m['year']), int(m['month'])
            if self.prefs['keep_archive']:
                archive_copy(gy, gm, out)
            save_month_record(gy, gm, m['employees'], m['days_of_month'], m['date_to'])
            self._status(T(f"✓ Το δελτίο {month_name(gm)} {gy} δημιουργήθηκε και αρχειοθετήθηκε.",
                           f"✓ The form for {month_name(gm)} {gy} was created and archived."))
        except Exception:
            pass
        self.refresh_history()
        print_status = None
        if print_after:
            if warnings_txt:      # never send a form with problems straight to the printer
                print_status = (False, T("Δεν εκτυπώθηκε αυτόματα επειδή υπάρχουν προειδοποιήσεις. "
                                         "Ελέγξτε το δελτίο και μετά πατήστε «Εκτύπωση».",
                                         "It was not printed automatically because there are warnings. "
                                         "Check the form and then click “Print”."))
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
