import pandas as pd
import pypdf
import io
import os
import warnings
import datetime
import calendar
import unicodedata
import re
import subprocess
import shutil
import math
import sys
import json
import hashlib
import base64
import tempfile
import copy
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Optional native file drag-and-drop support. Install with: pip install tkinterdnd2
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception as _dnd_err:          # ImportError, or the tkdnd binaries failing to load
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False
    DND_ERROR = repr(_dnd_err)
else:
    DND_ERROR = ""

TkBase = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

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


def install_edit_shortcuts(root):
    """Add familiar clipboard shortcuts and a right-click menu to text inputs."""
    def editable_widget(widget):
        return isinstance(widget, (tk.Entry, ttk.Entry, tk.Text))

    def is_readonly(widget):
        try:
            return str(widget.cget('state')) in ('disabled', 'readonly')
        except (tk.TclError, TypeError):
            return False

    def select_all(widget):
        if isinstance(widget, tk.Text):
            widget.tag_add('sel', '1.0', 'end-1c')
            widget.mark_set('insert', 'end-1c')
            widget.see('insert')
        else:
            widget.selection_range(0, 'end')
            widget.icursor('end')
        return 'break'

    def clipboard_action(widget, virtual_event):
        if not editable_widget(widget):
            return None
        if virtual_event in ('<<Paste>>', '<<Cut>>') and is_readonly(widget):
            return 'break'
        try:
            widget.event_generate(virtual_event)
        except tk.TclError:
            pass
        return 'break'

    def on_shortcut(event, action):
        widget = event.widget
        if not editable_widget(widget):
            return None
        if action == 'select_all':
            return select_all(widget)
        return clipboard_action(widget, {'copy': '<<Copy>>', 'cut': '<<Cut>>', 'paste': '<<Paste>>'}[action])

    def popup(event):
        widget = event.widget
        if not editable_widget(widget):
            return None
        menu = tk.Menu(root, tearoff=False)
        try:
            has_selection = bool(widget.tag_ranges('sel')) if isinstance(widget, tk.Text) else bool(widget.selection_present())
        except tk.TclError:
            has_selection = False
        menu.add_command(label=T('Αντιγραφή', 'Copy'), command=lambda: clipboard_action(widget, '<<Copy>>'),
                         state='normal' if has_selection else 'disabled')
        menu.add_command(label=T('Αποκοπή', 'Cut'), command=lambda: clipboard_action(widget, '<<Cut>>'),
                         state='normal' if has_selection and not is_readonly(widget) else 'disabled')
        menu.add_command(label=T('Επικόλληση', 'Paste'), command=lambda: clipboard_action(widget, '<<Paste>>'),
                         state='normal' if not is_readonly(widget) else 'disabled')
        menu.add_separator()
        menu.add_command(label=T('Επιλογή όλων', 'Select all'), command=lambda: select_all(widget))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return 'break'

    for sequence, action in (('<Control-c>', 'copy'), ('<Control-C>', 'copy'),
                             ('<Control-x>', 'cut'), ('<Control-X>', 'cut'),
                             ('<Control-v>', 'paste'), ('<Control-V>', 'paste'),
                             ('<Control-a>', 'select_all'), ('<Control-A>', 'select_all')):
        root.bind_all(sequence, lambda event, a=action: on_shortcut(event, a), add='+')
    root.bind_all('<Button-3>', popup, add='+')
    # Some Linux desktop environments use the middle button for the context menu.
    root.bind_all('<Button-2>', popup, add='+')


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------------------- storage
# elstat_config.json (next to the program) only remembers WHERE the data folder is.
# Everything else (hotel details, history, archive, preferences) lives in the data folder.
CONFIG_FILE = os.path.join(BASE_DIR, "elstat_config.json")
DEFAULT_DATA_DIR = os.path.join(BASE_DIR, "ELSTAT_Data")
DEFAULT_TEMPLATE = os.path.join(BASE_DIR, "template_elstat_report.pdf")
ARCHIVE_NAME = "ELSTAT_Archive"
DATA_FILES = ("hotel_settings.csv", "elstat_history.csv", "elstat_months.csv", "elstat_preferences.json", "elstat_country_mappings.json", "elstat_template_layout.json", "elstat_manual_layout.json")

DATA_DIR = SETTINGS_CSV = HISTORY_CSV = MONTHS_CSV = PREFS_JSON = ARCHIVE_DIR = COUNTRY_MAPPINGS_FILE = ""


def set_data_dir(path):
    """Points every data path of the program at `path`."""
    global DATA_DIR, SETTINGS_CSV, HISTORY_CSV, MONTHS_CSV, PREFS_JSON, ARCHIVE_DIR, COUNTRY_MAPPINGS_FILE
    DATA_DIR = os.path.abspath(path)
    SETTINGS_CSV = os.path.join(DATA_DIR, "hotel_settings.csv")
    HISTORY_CSV = os.path.join(DATA_DIR, "elstat_history.csv")
    MONTHS_CSV = os.path.join(DATA_DIR, "elstat_months.csv")
    PREFS_JSON = os.path.join(DATA_DIR, "elstat_preferences.json")
    ARCHIVE_DIR = os.path.join(DATA_DIR, ARCHIVE_NAME)
    COUNTRY_MAPPINGS_FILE = os.path.join(DATA_DIR, "elstat_country_mappings.json")


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
BOOL_PREFS = ('open_after', 'confirm_overwrite', 'keep_archive', 'show_reminder', 'prefill_employees', 'preview_before_result')
DEFAULT_PREFS = {
    'language': 'el',                          # 'el' | 'en'  (interface language)
    'theme': 'light',                          # 'light' | 'dark' | 'auto'  (follow the system)
    'output_mode': 'program',                  # 'program' | 'excel' | 'custom'
    'output_dir': '',                          # used when output_mode == 'custom'
    'filename_pattern': 'ELSTAT_{year}_{month}',
    'excel_dir': '',                           # '' = remember the last folder
    'template_pdf': '',                        # '' = template_elstat_report.pdf next to the program
    'smart_ack': False,                        # the 'experimental' notice of smart layout was accepted
    'pdf_mode': 'overlay',                     # 'overlay' (fixed positions) | 'smart' (layout read from the template) | 'manual' (layout editor)
    'open_after': False,
    'confirm_overwrite': True,
    'keep_archive': True,
    'show_reminder': True,
    'prefill_employees': False,
    'preview_before_result': False,             # show an in-app PDF preview before the Result dialog
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
    if prefs['pdf_mode'] not in ('overlay', 'smart', 'manual'):
        prefs['pdf_mode'] = DEFAULT_PREFS['pdf_mode']
    if prefs['theme'] not in ('light', 'dark', 'auto'):
        prefs['theme'] = DEFAULT_PREFS['theme']
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

    if 'sign_date' in monthly:
        try:
            datetime.datetime.strptime(monthly['sign_date'], "%d/%m/%Y")
        except ValueError:
            errors.append(T("Η ημερομηνία υπογραφής πρέπει να είναι της μορφής ΗΗ/ΜΜ/ΕΕΕΕ και υπαρκτή.",
                            "The signing date must be in the format DD/MM/YYYY and must exist."))

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

# User-added mappings are stored in the data folder and survive program updates.
CUSTOM_COUNTRY_MAPPING = {}
CUSTOM_COUNTRY_MAPPING_NORM = {}
COUNTRY_EXTRA_DESTINATIONS = (
    'ΛΟΙΠΕΣ ΧΩΡΕΣ ΕΥΡΩΠΗΣ', 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΜΕΡΙΚΗΣ',
    'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΣΙΑΣ', 'ΛΟΙΠΕΣ ΧΩΡΕΣ ΑΦΡΙΚΗΣ',
    'ΝΕΑ ΖΗΛΑΝΔΙΑ ΚΑΙ ΛΟΙΠΕΣ ΠΕΡΙΟΧΕΣ', 'Δεν δήλωσαν χώρα'
)

def country_mapping_destinations():
    """All valid destinations on the ELSTAT form, including regional 'other' rows."""
    values = (set(COUNTRY_MAPPING.values()) | set(COUNTRY_EXTRA_DESTINATIONS)) - {'Δεν δήλωσαν χώρα'}
    return sorted(values, key=lambda value: norm(value))

def load_custom_country_mappings():
    global CUSTOM_COUNTRY_MAPPING, CUSTOM_COUNTRY_MAPPING_NORM
    try:
        with open(COUNTRY_MAPPINGS_FILE, encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, dict):
            data = {}
    except (OSError, ValueError, TypeError):
        data = {}
    allowed = {norm(v) for v in country_mapping_destinations()}
    clean = {}
    for country, destination in data.items():
        key, value = norm(str(country)), norm(str(destination))
        if key and value in allowed:
            clean[str(country).strip()] = str(destination).strip()
    CUSTOM_COUNTRY_MAPPING = clean
    CUSTOM_COUNTRY_MAPPING_NORM = {norm(k): norm(v) for k, v in clean.items()}

def save_custom_country_mapping(country, destination):
    """Save one user-approved country-to-form-row mapping."""
    key, value = str(country).strip(), str(destination).strip()
    allowed = {norm(v) for v in country_mapping_destinations()}
    if not norm(key) or norm(value) not in allowed:
        raise ValueError(T("Μη έγκυρη αντιστοίχιση χώρας.", "Invalid country mapping."))
    updated = dict(CUSTOM_COUNTRY_MAPPING)
    # Replace spelling variants of the same source country rather than duplicating them.
    for old in list(updated):
        if norm(old) == norm(key):
            del updated[old]
    updated[key] = value
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = COUNTRY_MAPPINGS_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(updated, f, ensure_ascii=False, indent=2)
    os.replace(tmp, COUNTRY_MAPPINGS_FILE)
    load_custom_country_mappings()


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
    """Row of the form a country belongs to, including user-added country mappings."""
    k = norm(name)
    return (CUSTOM_COUNTRY_MAPPING_NORM.get(k) or COUNTRY_MAPPING_NORM.get(k)
            or (k if k in MAPPED_TARGETS else None))


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
    navy, text_c, band = colors.HexColor('#14304f'), colors.HexColor('#1e293b'), colors.HexColor('#e8f0fe')
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

    c.setFillColor(colors.HexColor('#64748b'))
    c.setFont(reg, 8.5)
    c.drawString(L, 30, T(f"Δημιουργήθηκε στις {datetime.date.today().strftime('%d/%m/%Y')} από το πρόγραμμα ΕΛΣΤΑΤ",
                          f"Created on {datetime.date.today().strftime('%d/%m/%Y')} by the ELSTAT program"))
    c.save()


# ------------------------------------------------------- template layout ("smart layout" mode)
# Default mode writes every value at fixed positions on top of the blank PDF.  "Smart layout" mode
# first READS the template (every piece of text and where it sits), compares it with a saved
# snapshot of the template the fixed positions were made for (the "reference layout") and moves
# every value by exactly as much as that part of the form has moved.
LAYOUT_NAME = "elstat_template_layout.json"
MIN_MATCH_RATIO = 0.6       # share of the reference's text that must be found again in the new template
MIN_ANCHORS = 8             # ...and at least this many matches on every page
LARGE_SHIFT = 40.0         # points: a bigger move than this deserves a human look
DUP_TOLERANCE = 30.0        # points: how far a repeated word may be from where it is expected


def layout_file():
    return os.path.join(DATA_DIR, LAYOUT_NAME)


def _median(values):
    v = sorted(values)
    n = len(v)
    if not n:
        return 0.0
    return v[n // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2.0


def _fit_line(pairs):
    """Robust v = a*u + b through (u, v) pairs. `a` stays 1.0 unless the data clearly shows a scale change."""
    if not pairs:
        return 1.0, 0.0
    sample = pairs if len(pairs) <= 150 else pairs[::max(1, len(pairs) // 150)]
    a = 1.0
    if len(sample) >= 6:
        slopes = [(v2 - v1) / (u2 - u1) for i, (u1, v1) in enumerate(sample)
                  for (u2, v2) in sample[i + 1:] if abs(u2 - u1) > 20]
        if slopes:
            s = _median(slopes)
            if 0.8 <= s <= 1.25 and abs(s - 1.0) > 0.002:
                a = s
    return a, _median([v - a * u for u, v in pairs])


def read_template_snapshot(template_pdf):
    """What the program needs to know about a template: for its first 2 pages the page size and every
    piece of text with its position (x, y in PDF points, in the page's own coordinates)."""
    reader = pypdf.PdfReader(template_pdf)
    if len(reader.pages) < 2:
        raise ValueError(T("Το template PDF πρέπει να έχει 2 σελίδες.", "The template PDF must have 2 pages."))
    with open(template_pdf, 'rb') as f:
        digest = hashlib.sha1(f.read()).hexdigest()
    pages = []
    for i in range(2):
        page = reader.pages[i]
        items = []

        def visit(text, cm, tm, font_dict, font_size, _items=items):
            n = norm(text)
            if not n:
                return
            x = tm[4] * cm[0] + tm[5] * cm[2] + cm[4]
            y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
            if not any(t == n and abs(x - ox) < 0.5 and abs(y - oy) < 0.5 for t, ox, oy in _items):
                _items.append([n, round(float(x), 2), round(float(y), 2)])

        page.extract_text(visitor_text=visit)
        box = page.mediabox
        pages.append({'size': [float(box.width), float(box.height)], 'items': items})
    return {'version': 1, 'sha1': digest, 'pages': pages}


def load_layout_reference():
    """The saved reference layout, or None."""
    try:
        with open(layout_file(), encoding='utf-8') as f:
            data = json.load(f)
        pages = data['pages']
        if len(pages) == 2 and all(isinstance(p.get('items'), list) for p in pages):
            return data
    except Exception:
        pass
    return None


def save_layout_reference(template_pdf):
    snap = read_template_snapshot(template_pdf)
    snap['saved'] = datetime.datetime.now().isoformat(timespec='seconds')
    snap['template'] = os.path.basename(template_pdf)
    ensure_data_dir()
    with open(layout_file(), 'w', encoding='utf-8') as f:
        json.dump(snap, f, ensure_ascii=False)
    return snap


def match_snapshots(ref, new):
    """Pairs the text of the reference template with the text of a new one, page by page.
    -> [{'anchors': [(ref_x, ref_y, dx, dy), ...], 'matched': n, 'total': n}, ...]"""
    out = []
    for p in range(2):
        r_items, n_items = ref['pages'][p]['items'], new['pages'][p]['items']
        r_by, n_by = {}, {}
        for t, x, y in r_items:
            r_by.setdefault(t, []).append((x, y))
        for t, x, y in n_items:
            n_by.setdefault(t, []).append((x, y))
        # words that appear once in both templates give a safe first estimate of the overall shift/scale
        pairs = [(r_by[t][0], n_by[t][0]) for t in r_by if len(r_by[t]) == 1 and len(n_by.get(t, ())) == 1]
        ax, bx = _fit_line([(r[0], n[0]) for r, n in pairs])
        ay, by = _fit_line([(r[1], n[1]) for r, n in pairs])
        # repeated words are paired with the closest candidate to where they are expected
        for t, rs in r_by.items():
            ns = list(n_by.get(t, ()))
            if not ns or (len(rs) == 1 and len(ns) == 1):
                continue
            for rx, ry in sorted(rs, key=lambda q: (-q[1], q[0])):
                if not ns:
                    break
                px, py = ax * rx + bx, ay * ry + by
                best = min(ns, key=lambda q: (q[0] - px) ** 2 + (q[1] - py) ** 2)
                if math.hypot(best[0] - px, best[1] - py) <= DUP_TOLERANCE:
                    ns.remove(best)
                    pairs.append(((rx, ry), best))
        # one overall shift/scale for the page, from every pair that was found
        mx = _fit_line([(r[0], n[0]) for r, n in pairs])
        my = _fit_line([(r[1], n[1]) for r, n in pairs])
        out.append({'anchors': [(r[0], r[1], n[0] - r[0], n[1] - r[1]) for r, n in pairs],
                    'model': (mx, my), 'matched': len(pairs), 'total': len(r_items)})
    return out


class TemplateLayout:
    """Tells where a value that belongs at (x, y) on the reference template goes on the new template.
    Two things are combined: the overall shift/scale of the page, and the small local differences
    measured at the text around the position (every matched piece of text is an anchor that knows
    how far it moved), so a section that moved on its own is followed and the rest stays put."""
    NEAREST = 6
    AGREE = 4.0         # points: a neighbour that disagrees with the others by more than this is ignored

    def __init__(self, pages, sizes, models):
        self.sizes, self.models = sizes, models
        self.pages = []                 # per page: anchors as (x, y, what the overall model misses in x / y)
        for anchors, ((ax, bx), (ay, by)) in zip(pages, models):
            self.pages.append([(x, y, dx - ((ax - 1.0) * x + bx), dy - ((ay - 1.0) * y + by))
                               for x, y, dx, dy in anchors])

    def size(self, page):
        return tuple(self.sizes[page])

    def offset(self, page, x, y):
        (ax, bx), (ay, by) = self.models[page]
        gx, gy = (ax - 1.0) * x + bx, (ay - 1.0) * y + by
        anchors = self.pages[page]
        if not anchors:
            return gx, gy
        near = sorted(anchors, key=lambda a: (a[0] - x) ** 2 + (a[1] - y) ** 2)[:self.NEAREST]
        mdx, mdy = _median([a[2] for a in near]), _median([a[3] for a in near])
        keep = [a for a in near if abs(a[2] - mdx) <= self.AGREE and abs(a[3] - mdy) <= self.AGREE] or near
        wsum = sx = sy = 0.0
        for a in keep:
            w = 1.0 / ((a[0] - x) ** 2 + (a[1] - y) ** 2 + 25.0)
            wsum += w
            sx += w * a[2]
            sy += w * a[3]
        return gx + sx / wsum, gy + sy / wsum

    def place(self, page, x, y):
        dx, dy = self.offset(page, x, y)
        return x + dx, y + dy


def analyze_template(template_pdf):
    """-> (TemplateLayout or None, report). The report says what was found (see describe_report)."""
    rep = {'ok': False, 'identical': False, 'reason': '', 'pages': []}
    ref = load_layout_reference()
    if ref is None:
        rep['reason'] = 'no_reference'
        return None, rep
    try:
        new = read_template_snapshot(template_pdf)
    except Exception as e:
        rep['reason'], rep['error'] = 'unreadable', str(e)
        return None, rep
    rep['identical'] = bool(ref.get('sha1')) and ref.get('sha1') == new['sha1']
    matches = match_snapshots(ref, new)
    for p, m in enumerate(matches):
        ratio = m['matched'] / m['total'] if m['total'] else 0.0
        n = len(m['anchors'])
        rep['pages'].append({
            'matched': m['matched'], 'total': m['total'], 'ratio': ratio,
            'mean_dx': sum(a[2] for a in m['anchors']) / n if n else 0.0,
            'mean_dy': sum(a[3] for a in m['anchors']) / n if n else 0.0,
            'max_shift': max((math.hypot(a[2], a[3]) for a in m['anchors']), default=0.0)})
        if not rep['reason'] and (m['matched'] < MIN_ANCHORS or ratio < MIN_MATCH_RATIO):
            rep['reason'] = 'low_match'
    if rep['reason']:
        return None, rep
    rep['ok'] = True
    return TemplateLayout([m['anchors'] for m in matches], [pg['size'] for pg in new['pages']],
                          [m['model'] for m in matches]), rep


def describe_report(rep):
    """The report of analyze_template() in words (current language)."""
    reason = rep.get('reason')
    if reason == 'no_reference':
        return T("Δεν έχει αποθηκευτεί ακόμη διάταξη αναφοράς.\n\nΠατήστε «Αποθήκευση ως αναφορά…» όταν το δελτίο "
                 "συμπληρώνεται σωστά με το τρέχον template.",
                 "No reference layout has been saved yet.\n\nPress “Save as reference…” when the form fills in "
                 "correctly with the current template.")
    if reason == 'unreadable':
        return T(f"Δεν μπόρεσα να διαβάσω το template: {rep.get('error', '')}\n\n"
                 "Βεβαιωθείτε ότι το αρχείο είναι έγκυρο PDF με 2 σελίδες και ότι δεν είναι κατεστραμμένο.",
                 f"I could not read the template: {rep.get('error', '')}\n\n"
                 "Make sure the file is a valid 2-page PDF and is not damaged.")
    if reason == 'low_match':
        head = T("✗ Το template δεν αναγνωρίζεται με ασφάλεια: έχει λίγα κοινά στοιχεία με την αναφορά.\n"
                 "Ίσως δεν είναι το σωστό έντυπο ή άλλαξε πολύ. Διαλέξτε το σωστό αρχείο στην καρτέλα «Στοιχεία "
                 "ξενοδοχείου» ή χρησιμοποιήστε την προεπιλεγμένη μέθοδο.",
                 "✗ The template cannot be recognised safely: it has too little in common with the reference.\n"
                 "It may not be the right form, or it changed a lot. Pick the right file in the “Hotel details” tab "
                 "or use the default method.")
    elif rep.get('identical'):
        head = T("✓ Το template είναι ίδιο με την αναφορά - δεν χρειάζεται καμία προσαρμογή.",
                 "✓ The template is identical to the reference - no adjustment is needed.")
    else:
        head = T("✓ Το template αναγνωρίστηκε. Κάθε τιμή θα τοποθετηθεί σύμφωνα με τη νέα του διάταξη.",
                 "✓ The template was recognised. Every value will be placed according to its new layout.")
    lines = [head, ""]
    for i, pg in enumerate(rep.get('pages', []), 1):
        lines.append(T(f"Σελίδα {i}: ταυτίστηκαν {pg['matched']} από {pg['total']} στοιχεία · "
                       f"μέση μετατόπιση {pg['mean_dx']:+.1f} / {pg['mean_dy']:+.1f} pt (μέγιστη {pg['max_shift']:.1f} pt)",
                       f"Page {i}: {pg['matched']} of {pg['total']} items matched · "
                       f"average shift {pg['mean_dx']:+.1f} / {pg['mean_dy']:+.1f} pt (largest {pg['max_shift']:.1f} pt)"))
    if rep.get('ok') and max((pg['max_shift'] for pg in rep.get('pages', [])), default=0.0) > LARGE_SHIFT:
        lines += ["", T("⚠ Το template έχει μετακινηθεί πολύ. Δημιουργήστε ένα «Δοκιμαστικό δελτίο» και ελέγξτε το πριν το υποβάλετε.",
                         "⚠ The template moved a lot. Create a “Test form” and check it before submitting.")]
    return "\n".join(lines)


# ------------------------------------------------------- manual layout ("Set manually" mode)
# In this mode the person decides where every value goes. The positions, fonts, sizes and colours
# are made in the built-in layout editor (class LayoutEditor) and kept in elstat_manual_layout.json.
MANUAL_LAYOUT_NAME = "elstat_manual_layout.json"
MANUAL_TOTAL_KEYS = ('2.1', '2.2', '2.3', '2', '3', '4', '5', '6', '8', '9')
MANUAL_TOTAL_Y = {'2.1': 46.5, '2.2': 699.5, '2.3': 520.0, '2': 500.1, '3': 407.1,
                  '4': 231.7, '5': 141.5, '6': 95.2, '8': 60.5, '9': 43.5}
MANUAL_TOTAL_NAMES = {
    '2.1': ("2.1 Σύνολο χωρών Ε.Ε.", "2.1 Total EU countries"),
    '2.2': ("2.2 Σύνολο χωρών Ε.Ζ.Ε.Σ.", "2.2 Total EFTA countries"),
    '2.3': ("2.3 Σύνολο λοιπών ευρωπαϊκών χωρών", "2.3 Total other European countries"),
    '2': ("2. Σύνολο Ευρώπης", "2. Total Europe"),
    '3': ("3. Σύνολο Αμερικής", "3. Total America"),
    '4': ("4. Σύνολο Ασίας", "4. Total Asia"),
    '5': ("5. Σύνολο Αφρικής", "5. Total Africa"),
    '6': ("6. Σύνολο Ωκεανίας", "6. Total Oceania"),
    '8': ("8. Σύνολο αλλοδαπών", "8. Total foreign"),
    '9': ("9. ΓΕΝΙΚΟ ΣΥΝΟΛΟ", "9. GRAND TOTAL"),
}
MANUAL_GROUPS = [
    ('header', ("Επικεφαλίδα", "Heading")),
    ('hotel', ("Στοιχεία καταλύματος", "Hotel details")),
    ('period', ("Περίοδος και προσωπικό", "Period and staff")),
    ('totals', ("Σύνολα κατηγοριών", "Category totals")),
    ('rows', ("Γραμμές χωρών", "Country rows")),
    ('sign', ("Ημερομηνία υπογραφής", "Signature date")),
    ('extra', ("Πρόσθετα στοιχεία", "Extra items")),
    ('custom', ("Δικά μου κείμενα", "My own texts")),
]
# fields whose text has a fixed length can be split into one item per character
MANUAL_SPLIT_LEN = {'afm': 9, 'days': 2, 'employees': 3, 'year': 4}


def _manual_field_table():
    """[(id, group, (greek, english), page, x, y, size, step)] - the places the program always used."""
    t = []

    def add(fid, group, label, page, x, y, size=11, step=0.0):
        t.append((fid, group, label, page, float(x), float(y), float(size), float(step)))

    add('year', 'header', ("Έτος έρευνας", "Survey year"), 0, 380, 675)
    add('month', 'header', ("Μήνας", "Month"), 0, 480, 675)
    add('type', 'hotel', ("Είδος καταλύματος", "Type of accommodation"), 0, 160, 613)
    add('category', 'hotel', ("Κατηγορία", "Category"), 0, 340, 612)
    add('name', 'hotel', ("Τίτλος καταλύματος", "Hotel name"), 0, 165, 597)
    add('phone', 'hotel', ("Τηλέφωνο", "Phone"), 0, 110, 548)
    add('email', 'hotel', ("E-mail", "E-mail"), 0, 420, 548)
    add('company', 'hotel', ("Επωνυμία επιχείρησης", "Company name"), 0, 435, 385)
    add('afm', 'hotel', ("ΑΦΜ (ένα ψηφίο ανά κουτάκι)", "VAT number (one digit per box)"), 0, 424, 342, step=14)
    add('date_from_d', 'period', ("Από: ημέρα", "From: day"), 0, 420, 261)
    add('date_from_m', 'period', ("Από: μήνας", "From: month"), 0, 465.9, 261)
    add('date_from_y', 'period', ("Από: έτος", "From: year"), 0, 511.7, 261)
    add('date_to_d', 'period', ("Έως: ημέρα", "To: day"), 0, 420, 233)
    add('date_to_m', 'period', ("Έως: μήνας", "To: month"), 0, 465.9, 233)
    add('date_to_y', 'period', ("Έως: έτος", "To: year"), 0, 511.7, 233)
    add('days', 'period', ("Ημέρες λειτουργίας", "Days of operation"), 0, 457, 169, step=16)
    add('employees', 'period', ("Σύνολο εργαζομένων", "Total employees"), 0, 448, 94, step=16)
    for key in MANUAL_TOTAL_KEYS:
        page = 0 if key == '2.1' else 1
        el, en = MANUAL_TOTAL_NAMES[key]
        add(f'tot_{key}_c', 'totals', (el + " · Πελάτες", en + " · Guests"), page, 265, MANUAL_TOTAL_Y[key], 10)
        add(f'tot_{key}_n', 'totals', (el + " · Διανυκτ.", en + " · Nights"), page, 335, MANUAL_TOTAL_Y[key], 10)
    add('rows0_c', 'rows', ("Σελίδα 1 · στήλη Πελάτες (όλες οι χώρες)", "Page 1 · Guests column (all countries)"), 0, 265, 0, 9)
    add('rows0_n', 'rows', ("Σελίδα 1 · στήλη Διανυκτ. (όλες οι χώρες)", "Page 1 · Nights column (all countries)"), 0, 335, 0, 9)
    add('rows1_c', 'rows', ("Σελίδα 2 · στήλη Πελάτες (όλες οι χώρες)", "Page 2 · Guests column (all countries)"), 1, 265, 0, 9)
    add('rows1_n', 'rows', ("Σελίδα 2 · στήλη Διανυκτ. (όλες οι χώρες)", "Page 2 · Nights column (all countries)"), 1, 335, 0, 9)
    add('sign_d', 'sign', ("Ημερομηνία: ημέρα", "Date: day"), 1, 454, 206, 10)
    add('sign_m', 'sign', ("Ημερομηνία: μήνας", "Date: month"), 1, 481.8, 206, 10)
    add('sign_y', 'sign', ("Ημερομηνία: έτος", "Date: year"), 1, 520.7, 206, 10)
    return t


MANUAL_FIELDS = _manual_field_table()
MANUAL_FIELD_INDEX = {f[0]: f for f in MANUAL_FIELDS}


def manual_new_item(page=0, x=100.0, y=400.0, size=11.0, step=0.0, custom=False, text=""):
    it = {'page': int(page), 'x': float(x), 'y': float(y), 'size': float(size), 'font': '',
          'bold': False, 'italic': False, 'color': '#000000', 'align': 'left',
          'step': float(step), 'visible': True}
    if custom:
        it['custom'] = True
        it['text'] = text
    return it


def manual_default_layout():
    items = {}
    for fid, _group, _label, page, x, y, size, step in MANUAL_FIELDS:
        items[fid] = manual_new_item(page, x, y, size, step)
    return {'version': 1, 'items': items}


def manual_layout_file():
    return os.path.join(DATA_DIR, MANUAL_LAYOUT_NAME)


def manual_layout_saved():
    return os.path.isfile(manual_layout_file())


def _merge_manual_item(base, saved):
    """Copies the usable values of a saved item over `base` (anything odd is ignored)."""
    for key in ('x', 'y', 'size', 'step'):
        try:
            value = float(saved[key])
            if math.isfinite(value):
                base[key] = value
        except (KeyError, TypeError, ValueError):
            pass
    base['size'] = min(max(base['size'], 3.0), 96.0)
    base['step'] = min(max(base['step'], 0.0), 200.0)
    if saved.get('page') in (0, 1):
        base['page'] = saved['page']
    if isinstance(saved.get('font'), str):
        base['font'] = saved['font']
    for key in ('bold', 'italic', 'visible'):
        if isinstance(saved.get(key), bool):
            base[key] = saved[key]
    if isinstance(saved.get('color'), str) and re.fullmatch(r'#[0-9a-fA-F]{6}', saved['color']):
        base['color'] = saved['color'].lower()
    if saved.get('align') in ('left', 'center', 'right'):
        base['align'] = saved['align']
    if base.get('custom') and isinstance(saved.get('text'), str):
        base['text'] = saved['text']
    for key in ('parent', 'country', 'source'):
        if isinstance(saved.get(key), str) and saved[key]:
            base[key] = saved[key]
    if isinstance(saved.get('index'), int) and not isinstance(saved.get('index'), bool):
        base['index'] = saved['index']
    if isinstance(saved.get('split'), bool):
        base['split'] = saved['split']


def load_manual_layout():
    """The saved manual layout merged over the program's own positions (so new fields always exist)."""
    layout = manual_default_layout()
    try:
        with open(manual_layout_file(), encoding='utf-8') as f:
            data = json.load(f)
        saved = data.get('items', {}) if isinstance(data, dict) else {}
    except Exception:
        saved = {}
    if not isinstance(saved, dict):
        saved = {}
    for fid, it in saved.items():
        if not isinstance(it, dict):
            continue
        if fid in layout['items']:
            _merge_manual_item(layout['items'][fid], it)
        elif not isinstance(fid, str):
            continue
        elif it.get('custom') and fid.startswith('custom_'):
            layout['items'][fid] = base = manual_new_item(custom=True)
            _merge_manual_item(base, it)
        elif it.get('source') and fid.startswith('extra_'):
            layout['items'][fid] = base = manual_new_item(size=10.0)
            _merge_manual_item(base, it)
        elif it.get('parent') and '@' in fid:
            layout['items'][fid] = base = manual_new_item()
            _merge_manual_item(base, it)
    items = layout['items']
    for fid in [f for f, i in items.items() if i.get('parent') and not items.get(i['parent'], {}).get('split')]:
        del items[fid]                      # a piece whose whole is not split any more
    for fid, i in items.items():
        if i.get('split') and not any(c.get('parent') == fid for c in items.values()):
            i['split'] = False
    return layout


def save_manual_layout(layout):
    ensure_data_dir()
    with open(manual_layout_file(), 'w', encoding='utf-8') as f:
        json.dump({'version': 1, 'items': layout['items']}, f, ensure_ascii=False, indent=1)


def delete_manual_layout():
    try:
        os.remove(manual_layout_file())
    except FileNotFoundError:
        pass


def _pretty_target(key):
    """Readable name of a form row stored as a normalised key."""
    for v in country_mapping_destinations():
        if norm(v) == key:
            return v
    return pretty_country(key)


_EXTRA_SOURCES = (
    ('month_name', ("Όνομα μήνα (π.χ. Μάρτιος)", "Month name (e.g. March)")),
    ('year2', ("Έτος (2 ψηφία)", "Year (2 digits)")),
    ('date_from_full', ("Από: πλήρης ημερομηνία", "From: full date")),
    ('date_to_full', ("Έως: πλήρης ημερομηνία", "To: full date")),
    ('sign_full', ("Ημερομηνία υπογραφής (πλήρης)", "Signing date (full)")),
    ('hotel_line', ("Είδος · Τίτλος · Κατηγορία", "Type · Name · Category")),
)


def manual_source_catalog():
    """[(key, (greek, english), group)] - everything that can be put on the form (searched in the editor)."""
    cat = []
    for fid, group, label, *_rest in MANUAL_FIELDS:
        if group != 'rows':
            cat.append((fid, label, group))
    for key, label in _EXTRA_SOURCES:
        cat.append((key, label, 'extra'))
    for name in country_mapping_destinations():
        key = norm(name)
        cat.append((f'country:{key}:c', (f"{name} · Πελάτες", f"{name} · Guests"), 'country'))
        cat.append((f'country:{key}:n', (f"{name} · Διανυκτ.", f"{name} · Nights"), 'country'))
    return cat


def manual_source_label(source):
    for key, label, _group in manual_source_catalog():
        if key == source:
            return T(*label)
    return source


def manual_group(fid, it):
    if it.get('custom'):
        return 'custom'
    if it.get('source'):
        return 'extra'
    if it.get('parent'):
        fid = it['parent']
    return MANUAL_FIELD_INDEX[fid][1] if fid in MANUAL_FIELD_INDEX else 'custom'


def manual_label(fid, it):
    if it.get('custom'):
        return it.get('text') or T("(κενό κείμενο)", "(empty text)")
    if it.get('source'):
        return manual_source_label(it['source'])
    parent = it.get('parent')
    if parent and it.get('country'):
        col = T("Πελάτες", "Guests") if parent.endswith('_c') else T("Διανυκτ.", "Nights")
        return f"{_pretty_target(it['country'])} · {col} · " + T(f"Σελ. {it.get('page', 0) + 1}", f"Pg {it.get('page', 0) + 1}")
    if parent:
        return f"{manual_label(parent, {})} · #{it.get('index', 0) + 1}"
    entry = MANUAL_FIELD_INDEX.get(fid)
    return T(*entry[2]) if entry else fid


def manual_split_item(layout, fid, sample_text, row_keys):
    """Replaces one grouped item (country column, digit boxes) by individual items. Returns the new ids."""
    items = layout['items']
    it = items[fid]
    if it.get('split') or it.get('parent') or it.get('custom') or it.get('source'):
        return []
    keep = {k: it[k] for k in ('page', 'size', 'font', 'bold', 'italic', 'color', 'visible', 'align')}
    made = []
    if fid.startswith('rows'):
        for key in row_keys:
            cid = f"{fid}@{key}"
            child = manual_new_item(it['page'], it['x'], it['y'], it['size'])
            child.update(keep, parent=fid, country=key, step=0.0)
            items[cid] = child
            made.append(cid)
    else:
        text = sample_text
        if len(text) < 2:
            return []
        step = it['step']
        if step:
            xs = [it['x'] + i * step for i in range(len(text))]
            align = it['align']
        else:
            font = manual_pdf_font(it['font'], it['bold'], it['italic'])
            widths = [pdfmetrics.stringWidth(ch, font, it['size']) for ch in text]
            total = sum(widths)
            start = it['x'] - {'left': 0.0, 'center': total / 2.0, 'right': total}.get(it['align'], 0.0)
            xs = [start + sum(widths[:i]) for i in range(len(text))]
            align = 'left'
        for i, x in enumerate(xs):
            cid = f"{fid}@{i}"
            child = manual_new_item(it['page'], round(x, 2), it['y'], it['size'])
            child.update(keep, parent=fid, index=i, step=0.0, align=align)
            items[cid] = child
            made.append(cid)
    if made:
        it['split'] = True
    return made


def manual_merge_item(layout, fid):
    """Removes the individual pieces of `fid` and makes it a single item again."""
    items = layout['items']
    for cid in [c for c, i in items.items() if i.get('parent') == fid]:
        del items[cid]
    if fid in items:
        items[fid]['split'] = False


# ---- fonts: only families that are really installed are offered
_FONT_DIRS = [r"C:\Windows\Fonts", "/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/truetype/liberation",
              "/usr/share/fonts/truetype/liberation2", "/usr/share/fonts/TTF", "/usr/share/fonts/dejavu",
              "/Library/Fonts", "/System/Library/Fonts/Supplemental", os.path.expanduser("~/Library/Fonts"),
              os.path.expanduser("~/.fonts"), os.path.expanduser("~/.local/share/fonts")]
# family -> candidate file names of (regular, bold, italic, bold italic)
_FONT_FILES = {
    'Arial': (('arial.ttf', 'Arial.ttf'), ('arialbd.ttf', 'Arial Bold.ttf'), ('ariali.ttf', 'Arial Italic.ttf'),
              ('arialbi.ttf', 'Arial Bold Italic.ttf')),
    'Tahoma': (('tahoma.ttf',), ('tahomabd.ttf',), (), ()),
    'DejaVu Sans': (('DejaVuSans.ttf',), ('DejaVuSans-Bold.ttf',), ('DejaVuSans-Oblique.ttf',),
                    ('DejaVuSans-BoldOblique.ttf',)),
    'Liberation Sans': (('LiberationSans-Regular.ttf',), ('LiberationSans-Bold.ttf',),
                        ('LiberationSans-Italic.ttf',), ('LiberationSans-BoldItalic.ttf',)),
    'Calibri': (('calibri.ttf',), ('calibrib.ttf',), ('calibrii.ttf',), ('calibriz.ttf',)),
    'Segoe UI': (('segoeui.ttf',), ('segoeuib.ttf',), ('segoeuii.ttf',), ('segoeuiz.ttf',)),
    'Verdana': (('verdana.ttf', 'Verdana.ttf'), ('verdanab.ttf', 'Verdana Bold.ttf'),
                ('verdanai.ttf', 'Verdana Italic.ttf'), ('verdanaz.ttf', 'Verdana Bold Italic.ttf')),
    'Times New Roman': (('times.ttf', 'Times New Roman.ttf'), ('timesbd.ttf', 'Times New Roman Bold.ttf'),
                        ('timesi.ttf', 'Times New Roman Italic.ttf'), ('timesbi.ttf', 'Times New Roman Bold Italic.ttf')),
    'Courier New': (('cour.ttf', 'Courier New.ttf'), ('courbd.ttf', 'Courier New Bold.ttf'),
                    ('couri.ttf', 'Courier New Italic.ttf'), ('courbi.ttf', 'Courier New Bold Italic.ttf')),
    'DejaVu Serif': (('DejaVuSerif.ttf',), ('DejaVuSerif-Bold.ttf',), ('DejaVuSerif-Italic.ttf',),
                     ('DejaVuSerif-BoldItalic.ttf',)),
    'Liberation Serif': (('LiberationSerif-Regular.ttf',), ('LiberationSerif-Bold.ttf',),
                         ('LiberationSerif-Italic.ttf',), ('LiberationSerif-BoldItalic.ttf',)),
    'DejaVu Sans Mono': (('DejaVuSansMono.ttf',), ('DejaVuSansMono-Bold.ttf',), ('DejaVuSansMono-Oblique.ttf',),
                         ('DejaVuSansMono-BoldOblique.ttf',)),
    'Liberation Mono': (('LiberationMono-Regular.ttf',), ('LiberationMono-Bold.ttf',),
                        ('LiberationMono-Italic.ttf',), ('LiberationMono-BoldItalic.ttf',)),
}
_FONT_FAMILIES_CACHE = []


def _find_font_file(names):
    for folder in _FONT_DIRS:
        for name in names:
            path = os.path.join(folder, name)
            if os.path.isfile(path):
                return path
    return None


def manual_font_families():
    """Font families that are installed on this computer (the first one is the program's usual font)."""
    if not _FONT_FAMILIES_CACHE:
        _FONT_FAMILIES_CACHE.extend(f for f, files in _FONT_FILES.items() if _find_font_file(files[0]))
    return list(_FONT_FAMILIES_CACHE)


def manual_pdf_font(family, bold=False, italic=False):
    """Name of a reportlab font for the family/style (registered on first use)."""
    families = manual_font_families()
    if family not in families:
        family = families[0] if families else None
    if family is None:
        return ('Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique', 'Helvetica-BoldOblique')[(1 if bold else 0) + (2 if italic else 0)]
    files = _FONT_FILES[family]
    wanted = (1 if bold else 0) + (2 if italic else 0)            # 0 regular, 1 bold, 2 italic, 3 bold italic
    order = [wanted] + ([1] if bold and wanted == 3 else []) + ([2] if italic and wanted == 3 else []) + [0]
    for idx in order:
        path = _find_font_file(files[idx]) if files[idx] else None
        if not path:
            continue
        name = f"MF_{family.replace(' ', '')}_{idx}"
        if name not in pdfmetrics.getRegisteredFontNames():
            try:
                pdfmetrics.registerFont(TTFont(name, path))
            except Exception:
                continue
        return name
    return 'Helvetica'


# ---- the text that goes into each field
def manual_field_values(settings_data, monthly_data, totals):
    """{field id: text} for every automatic field."""
    v = {k: settings_data.get(k, '') for k in ('type', 'category', 'name', 'phone', 'email', 'company', 'afm')}
    v['year'] = monthly_data['year']
    v['month'] = monthly_data['month']
    v.update(date_from_d='01', date_from_m='01', date_from_y=monthly_data['year'])
    parts = monthly_data['date_to'].split('/')
    if len(parts) == 3:
        v.update(date_to_d=parts[0], date_to_m=parts[1], date_to_y=parts[2])
    else:
        v.update(date_to_d=monthly_data['date_to'], date_to_m='', date_to_y='')
    v['days'] = monthly_data['days_of_month'].zfill(2)
    v['employees'] = monthly_data.get('employees', '0').zfill(3)
    try:
        today = datetime.datetime.strptime(monthly_data.get('sign_date', ''), "%d/%m/%Y").date()
    except ValueError:
        today = datetime.date.today()
    v.update(sign_d=today.strftime('%d'), sign_m=today.strftime('%m'),
             sign_y=str(today.year)[-2:] if str(today.year).startswith("20") else str(today.year))
    for key in MANUAL_TOTAL_KEYS:
        if totals[key]['clients'] > 0:
            v[f'tot_{key}_c'] = str(totals[key]['clients'])
            v[f'tot_{key}_n'] = str(totals[key]['nights'])
    try:
        mi = int(monthly_data['month'])
    except (ValueError, TypeError):
        mi = 0
    v['month_name'] = GREEK_MONTHS[mi - 1] if 1 <= mi <= 12 else ''
    v['year2'] = str(monthly_data['year'])[-2:]
    v['date_from_full'] = f"01/01/{monthly_data['year']}"
    v['date_to_full'] = monthly_data['date_to']
    v['sign_full'] = today.strftime('%d/%m/%Y')
    v['hotel_line'] = ' · '.join(x for x in (v['type'], v['name'], v['category']) if x)
    return v


def manual_sample_inputs():
    """Made-up hotel + month, used for the editor preview and its test PDF."""
    sample = {'type': 'ΞΕΝΟΔΟΧΕΙΟ', 'name': 'TEST HOTEL', 'category': '5*', 'email': 'test@example.com',
              'phone': '2310000000', 'afm': '123456789', 'company': 'TEST COMPANY'}
    today = datetime.date.today()
    last = calendar.monthrange(today.year, today.month)[1]
    mon = dict(year=str(today.year), month=str(today.month), date_to=f"{last:02d}/{today.month:02d}/{today.year}",
               days_of_month=str(last), employees='25', sign_date=today.strftime('%d/%m/%Y'))
    return sample, mon


def manual_sample_values():
    sample, mon = manual_sample_inputs()
    totals = {key: {'clients': 1234, 'nights': 5678} for key in MANUAL_TOTAL_KEYS}
    return manual_field_values(sample, mon, totals)


# ---- drawing
def _hex_rgb(color):
    try:
        return tuple(int(color[i:i + 2], 16) / 255.0 for i in (1, 3, 5))
    except (ValueError, TypeError, IndexError):
        return (0.0, 0.0, 0.0)


def _manual_put(can, it, x, y, text):
    step = it.get('step') or 0
    draw = {'center': can.drawCentredString, 'right': can.drawRightString}.get(it.get('align'), can.drawString)
    if step:
        for i, ch in enumerate(text):
            draw(x + i * step, y, ch)
    else:
        draw(x, y, text)


def draw_manual_page(can, page_num, layout, values, row_ys, row_vals):
    """Writes every visible item of one page. row_ys: {country: y on this page}; row_vals: {country: {...}}."""
    for fid, it in layout['items'].items():
        if it.get('page') != page_num or not it.get('visible', True) or it.get('split'):
            continue
        can.setFont(manual_pdf_font(it.get('font', ''), it.get('bold', False), it.get('italic', False)), it['size'])
        can.setFillColorRGB(*_hex_rgb(it.get('color', '#000000')))
        parent, country, source = it.get('parent'), it.get('country'), it.get('source')
        if country:                               # one country of a split column
            column = 'clients' if parent.endswith('_c') else 'nights'
            if country in row_ys and country in row_vals:
                _manual_put(can, it, it['x'], row_ys[country] + it['y'], str(row_vals[country][column]))
        elif parent:                              # one character of a split field
            text = str(values.get(parent, ''))
            idx = it.get('index', 0)
            if 0 <= idx < len(text):
                _manual_put(can, it, it['x'], it['y'], text[idx])
        elif fid.startswith('rows'):
            column = 'clients' if fid.endswith('_c') else 'nights'
            for country, y in row_ys.items():
                if country in row_vals:
                    _manual_put(can, it, it['x'], y + it['y'], str(row_vals[country][column]))
        elif source and source.startswith('country:'):
            _tag, target, col = source.split(':')
            if target in row_vals:
                _manual_put(can, it, it['x'], it['y'],
                            str(row_vals[target]['clients' if col == 'c' else 'nights']))
        else:
            text = it.get('text', '') if it.get('custom') else values.get(source or fid, '')
            if text:
                _manual_put(can, it, it['x'], it['y'], str(text))


# ---- reading the template (used by the editor)
def detect_template_rows(template_pdf):
    """[{country: y}, {country: y}] - where each country row sits on page 1 and page 2 of the template."""
    reader = pypdf.PdfReader(template_pdf)
    targets = sorted(MAPPED_TARGETS | {norm('ΕΛΛΑΔΑ')}, key=len, reverse=True)
    found = [{}, {}]
    for i in range(min(2, len(reader.pages))):
        def visitor_text(text, cm, tm, fontDict, fontSize, _i=i):
            n = norm(text)
            if not n:
                return
            for target in targets:
                if n.startswith(target):
                    found[_i].setdefault(target, tm[5])
                    break
        reader.pages[i].extract_text(visitor_text=visitor_text)
    if norm('ΤΟΥΡΚΙΑ') not in found[1]:
        found[1][norm('ΤΟΥΡΚΙΑ')] = 555.2222
    return found


def template_text_items(template_pdf):
    """[[(text, x, y, size)], [...]] for the 2 pages - a plain stand-in when the page cannot be drawn."""
    reader = pypdf.PdfReader(template_pdf)
    pages = [[], []]
    for i in range(min(2, len(reader.pages))):
        def visitor_text(text, cm, tm, fontDict, fontSize, _i=i):
            if text.strip():
                x = tm[4] * cm[0] + tm[5] * cm[2] + cm[4]
                y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
                pages[_i].append((text.strip(), x, y, max(5.0, float(fontSize or 8) * abs(cm[0] or 1))))
        reader.pages[i].extract_text(visitor_text=visitor_text)
    return pages


def _png_bytes(width, height, rgb):
    import struct
    import zlib
    stride = width * 3
    raw = b''.join(b'\x00' + rgb[y * stride:(y + 1) * stride] for y in range(height))

    def chunk(kind, data):
        body = kind + data
        return struct.pack('>I', len(data)) + body + struct.pack('>I', zlib.crc32(body) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw, 3)) + chunk(b'IEND', b''))


def render_pdf_page_png(pdf_path, page_index, scale):
    """PNG bytes of one PDF page drawn at `scale` pixels per point, or None if nothing here can draw PDFs.
    Tried in turn: pypdfium2, PyMuPDF, Poppler's pdftoppm."""
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(pdf_path)
        try:
            bitmap = doc[page_index].render(scale=scale)
            try:
                buf = io.BytesIO()
                bitmap.to_pil().convert('RGB').save(buf, format='PNG')
                return buf.getvalue()
            except Exception:
                arr = bitmap.to_numpy()
                rgb = arr[:, :, :3][:, :, ::-1].copy().tobytes()
                return _png_bytes(arr.shape[1], arr.shape[0], rgb)
        finally:
            try:
                doc.close()
            except Exception:
                pass
    except Exception:
        pass
    try:
        import fitz
        doc = fitz.open(pdf_path)
        try:
            return doc[page_index].get_pixmap(matrix=fitz.Matrix(scale, scale)).tobytes('png')
        finally:
            doc.close()
    except Exception:
        pass
    exe = shutil.which('pdftoppm')
    if exe:
        tmp = tempfile.mkdtemp()
        try:
            prefix = os.path.join(tmp, 'p')
            subprocess.run([exe, '-png', '-r', f"{72 * scale:.1f}", '-f', str(page_index + 1), '-l', str(page_index + 1),
                            '-singlefile', pdf_path, prefix], check=True, capture_output=True, timeout=40,
                           creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            with open(prefix + '.png', 'rb') as f:
                return f.read()
        except Exception:
            pass
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    return None


def process_elstat_report(excel_path, template_pdf, output_pdf, settings_data, monthly_data,
                         rows_override=None, save_history=True, pdf_mode='overlay', manual_layout=None):
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

    # "Smart layout": find out where everything sits on THIS template; any failure -> default method
    layout, layout_warning, layout_info = None, None, None
    if pdf_mode == 'smart':
        layout, layout_report = analyze_template(template_pdf)
        if layout is None:
            layout_warning = (T("Η έξυπνη διάταξη δεν εφαρμόστηκε - χρησιμοποιήθηκε η προεπιλεγμένη μέθοδος.\n",
                                "Smart layout was not applied - the default method was used.\n")
                              + describe_report(layout_report))
        else:
            shifts = layout_report['pages']
            biggest = max(p['max_shift'] for p in shifts)
            layout_info = T(f"Έξυπνη διάταξη: μέγιστη μετατόπιση {biggest:.1f} pt.",
                            f"Smart layout: largest shift {biggest:.1f} pt.")
            if biggest > LARGE_SHIFT:
                layout_warning = T(f"Η έξυπνη διάταξη εφαρμόστηκε, αλλά το template έχει μετακινηθεί πολύ ({biggest:.0f} pt). "
                                   "Ελέγξτε προσεκτικά το δελτίο πριν το υποβάλετε.",
                                   f"Smart layout was applied, but the template moved a lot ({biggest:.0f} pt). "
                                   "Check the form carefully before submitting.")

    def place(page_num, x, y):
        return layout.place(page_num, x, y) if layout else (x, y)

    coord_map = {0: {}, 1: {}}
    for i in range(2):
        page = reader.pages[i]
        def visitor_text(text, cm, tm, fontDict, fontSize, _i=i):
            n = norm(text)
            if not n:
                return
            for target in sorted(pdf_data.keys(), key=len, reverse=True):
                if n.startswith(target):
                    y = tm[5] if layout is None else tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
                    coord_map[_i].setdefault(target, y)
                    break
        page.extract_text(visitor_text=visitor_text)

    if norm('ΤΟΥΡΚΙΑ') in pdf_data and norm('ΤΟΥΡΚΙΑ') not in coord_map[1]:
        coord_map[1][norm('ΤΟΥΡΚΙΑ')] = place(1, 265, 555.2222)[1]

    fixed_y = {
        '2.1': 46.5, '2.2': 699.5, '2.3': 520.0, '2': 500.1,
        '3': 407.1, '4': 231.7, '5': 141.5, '6': 95.2, 
        '8': 60.5, '9': 43.5
    }

    def create_overlay(page_num, data_dict, coords, totals_dict):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=layout.size(page_num)) if layout else canvas.Canvas(packet)

        def put(x, y, s):
            nx, ny = place(page_num, x, y)
            can.drawString(nx, ny, s)

        def put_row(x, y, s):       # country rows: y was read from this template, only x may need to move
            can.drawString(x + (layout.offset(page_num, x, y)[0] if layout else 0.0), y, s)
        
        # ΣΕΛΙΔΑ 1
        if page_num == 0:
            can.setFont(font_name, 11)
            put(160, 613, settings_data['type'])
            put(340, 612, settings_data['category'])
            put(165, 597, settings_data['name'])
            put(110, 548, settings_data['phone'])
            put(420, 548, settings_data['email'])  
            put(435, 385, settings_data['company'])
            
            afm_str = settings_data['afm']
            afm_start_x = 424  
            afm_step = 14.0    
            
            for i, digit in enumerate(afm_str):
                put(afm_start_x + (i * afm_step), 342, digit)
            
            put(380, 675, monthly_data['year'])
            put(480, 675, monthly_data['month'])
            
            date_from_str = f"01           01           {monthly_data['year']}"
            put(420, 261, date_from_str)
            
            dt_parts = monthly_data['date_to'].split('/')
            if len(dt_parts) == 3:
                date_to_str = f"{dt_parts[0]}           {dt_parts[1]}           {dt_parts[2]}"
            else:
                date_to_str = monthly_data['date_to']
            put(420, 233, date_to_str)
            
            days_val = monthly_data['days_of_month']
            if len(days_val) == 1:
                days_val = "0" + days_val  
                
            days_start_x = 457  
            days_step = 16.0    
            
            for i, digit in enumerate(days_val):
                put(days_start_x + (i * days_step), 169, digit)
                
            emp_val = monthly_data.get('employees', '0')
            emp_val = emp_val.zfill(3) if len(emp_val) < 3 else emp_val
            
            emp_start_x = 448  
            emp_step = 16.0    
            emp_y = 94       
            
            for i, digit in enumerate(emp_val):
                put(emp_start_x + (i * emp_step), emp_y, digit)
            
            can.setFont(font_name, 10)
            if totals_dict['2.1']['clients'] > 0:
                put(265, fixed_y['2.1'], str(totals_dict['2.1']['clients']))
                put(335, fixed_y['2.1'], str(totals_dict['2.1']['nights']))

        # ΣΕΛΙΔΑ 2
        else:
            can.setFont(font_name, 10)
            for key in fixed_y.keys():
                if key != '2.1' and totals_dict[key]['clients'] > 0:
                    put(265, fixed_y[key], str(totals_dict[key]['clients']))
                    put(335, fixed_y[key], str(totals_dict[key]['nights']))
                    
            try:
                today = datetime.datetime.strptime(monthly_data.get('sign_date', ''), "%d/%m/%Y").date()
            except ValueError:
                today = datetime.date.today()
            year_suffix = str(today.year)[-2:] if str(today.year).startswith("20") else str(today.year)
            today_str = f"{today.strftime('%d')}      {today.strftime('%m')}          {year_suffix}"
            put(454, 206, today_str)

        can.setFont(font_name, 9)
        for country, y_val in coords.items():
            if country in data_dict:
                cl, ni = data_dict[country]['clients'], data_dict[country]['nights']
                put_row(265, y_val, str(cl))
                put_row(335, y_val, str(ni))
                
        can.save()
        packet.seek(0)
        return pypdf.PdfReader(packet)

    # "Set manually": positions, fonts and colours come from the layout made in the layout editor
    manual = manual_values = None
    if pdf_mode == 'manual':
        manual = manual_layout if manual_layout is not None else load_manual_layout()
        manual_values = manual_field_values(settings_data, monthly_data, totals)

    def create_manual_overlay(page_num):
        box = reader.pages[page_num].mediabox
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(float(box.width), float(box.height)))
        draw_manual_page(can, page_num, manual, manual_values, coord_map[page_num], pdf_data)
        can.save()
        packet.seek(0)
        return pypdf.PdfReader(packet)

    writer = pypdf.PdfWriter()
    for i in range(2):
        page = reader.pages[i]
        overlay_pdf = create_manual_overlay(i) if manual is not None else create_overlay(i, pdf_data, coord_map[i], totals)
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    if save_history and pdf_mode == 'overlay' and load_layout_reference() is None:
        try:        # the first real form made the default way shows the template the fixed positions fit
            save_layout_reference(template_pdf)
        except Exception:
            pass

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
    return {'unmapped': unmapped, 'excel_total': excel_total, 'computed_total': computed, 'summary': summary,
            'layout_warning': layout_warning, 'layout_info': layout_info}


# ----------------------------------------------------------------- look & feel
PALETTES = {
    'light': dict(
        BG="#f3f6fb", CARD="#ffffff", NAVY="#172b4d", TEXT="#243247", MUTED="#718096", BORDER="#e2e8f0",
        HEADER_BG="#172b4d", HEADER_FG="#ffffff", HEADER_MUTED="#c4d0e2",
        ACCENT="#4263eb", ACCENT_DARK="#3451d1", ACCENT_SOFT="#edf1ff", TEAL="#0f8b83", TEAL_DARK="#08766f",
        GOOD="#16845b", BAD="#c2414b", WARN="#a86b12",
        FIELD="#ffffff", BTN="#edf1f7", BTN_HOVER="#e1e8f2", TAB_BG="#e9eef6", TAB_HOVER="#dfe7f2",
        OFF_BG="#cbd5e1", OFF_FG="#f8fafc", SCROLL="#cbd5e1",
        GOOD_BG="#dcfce7", GOOD_FG="#166534", BAD_BG="#fee2e2", BAD_FG="#991b1b", BAD_LINE="#f87171",
        WARN_BG="#fef3c7", WARN_FG="#92400e", WARN_LINE="#f5d68a",
        DANGER_HOVER="#fecaca", DANGER_LINE="#fca5a5",
        ZONE_IDLE_BG="#f8fafc", ZONE_IDLE_LINE="#b6c2d3", ZONE_HOVER_BG="#eff6ff", ZONE_DRAG_BG="#dbeafe",
        ZONE_DONE_BG="#f0fdf4", ZONE_DONE_LINE="#86efac",
        SEL="#e5ebff", HEAD_BG="#edf1f7", HEAD_HOVER="#e3eaf4",
        ROW_BOLD="#edf1ff", ROW_YEAR="#e3eaff", ROW_MONTH="#f2f5fc", ROW_HL="#fff5dc",
        GRID="#e7ecf3", TRACK="#f2f5fa", TICK="#cbd5e1", TIP_BG="#172b4d", TIP_FG="#ffffff",
    ),
    'dark': dict(
        BG="#101827", CARD="#192538", NAVY="#edf3ff", TEXT="#dce6f5", MUTED="#91a2ba", BORDER="#2b3a50",
        HEADER_BG="#0c1422", HEADER_FG="#f5f8ff", HEADER_MUTED="#9aacc5",
        ACCENT="#6d8cff", ACCENT_DARK="#5575ed", ACCENT_SOFT="#24365d", TEAL="#22a69b", TEAL_DARK="#16877e",
        GOOD="#55d6a2", BAD="#ff7b86", WARN="#f2c36b",
        FIELD="#111d2d", BTN="#26364d", BTN_HOVER="#31445e", TAB_BG="#1d2b40", TAB_HOVER="#293b55",
        OFF_BG="#33445b", OFF_FG="#8798b0", SCROLL="#40536d",
        GOOD_BG="#12301f", GOOD_FG="#86efac", BAD_BG="#3b1618", BAD_FG="#fca5a5", BAD_LINE="#ef4444",
        WARN_BG="#3a2c0f", WARN_FG="#fcd34d", WARN_LINE="#5c4514",
        DANGER_HOVER="#4d1d20", DANGER_LINE="#7f1d1d",
        ZONE_IDLE_BG="#131c27", ZONE_IDLE_LINE="#3a4a5f", ZONE_HOVER_BG="#152a45", ZONE_DRAG_BG="#16304f",
        ZONE_DONE_BG="#11291d", ZONE_DONE_LINE="#2f7a4d",
        SEL="#2b4169", HEAD_BG="#22334a", HEAD_HOVER="#2b405b",
        ROW_BOLD="#243650", ROW_YEAR="#263c5b", ROW_MONTH="#1e3048", ROW_HL="#493a20",
        GRID="#2b3b52", TRACK="#1b2a3e", TICK="#526783", TIP_BG="#edf3ff", TIP_FG="#101827",
    ),
}
THEME = 'light'


def system_theme():
    """'dark' or 'light' according to the operating system (light if it cannot be found out)."""
    try:
        if sys.platform.startswith('win'):
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
                return 'light' if winreg.QueryValueEx(k, "AppsUseLightTheme")[0] else 'dark'
        if sys.platform == 'darwin':
            r = subprocess.run(['defaults', 'read', '-g', 'AppleInterfaceStyle'],
                               capture_output=True, text=True, timeout=2)
            return 'dark' if 'dark' in r.stdout.lower() else 'light'
    except Exception:
        pass
    return 'light'


def set_theme(pref):
    """pref: 'light' | 'dark' | 'auto'. Points every colour name (BG, CARD, ...) at the chosen palette.
    The screen has to be built again afterwards."""
    global THEME
    THEME = pref if pref in ('light', 'dark') else system_theme()
    globals().update(PALETTES[THEME])


set_theme('light')
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
    """White 'card' with a thin border, an optional numbered step badge and a title.
    The badge turns into a green tick when set_done(True) is called. Put widgets into `card.body`."""

    def __init__(self, parent, number=None, title=""):
        super().__init__(parent, bg=BORDER)
        inner = tk.Frame(self, bg=CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        head = tk.Frame(inner, bg=CARD)
        head.pack(fill="x", padx=18, pady=(16, 2))
        self.number = number
        self.badge = None
        self._done = None
        if number is not None:
            self.badge = tk.Canvas(head, width=26, height=26, bg=CARD, highlightthickness=0)
            self.badge.pack(side="left")
            self.set_done(False)
        tk.Label(head, text=title, bg=CARD, fg=NAVY,
                 font=(FONT, 13, 'bold')).pack(side="left", padx=(10 if number is not None else 0, 0))
        self.body = ttk.Frame(inner, style='Card.TFrame', padding=(20, 10, 20, 20))
        self.body.pack(fill="x")

    def set_done(self, done):
        if self.badge is None or done == self._done:
            return
        self._done = done
        colour = GOOD if done else ACCENT
        self.badge.delete('all')
        self.badge.create_oval(1, 1, 25, 25, fill=colour, outline=colour)
        self.badge.create_text(13, 13, text="\u2713" if done else str(self.number), fill="white",
                               font=(FONT, 10, 'bold'))


class DropZone(tk.Canvas):
    """Large dashed 'drop your file here' area. Clicking it calls `command` (the file chooser)."""

    def __init__(self, parent, command, height=96):
        super().__init__(parent, height=height, bg=CARD, highlightthickness=0, cursor='hand2')
        self._command = command
        self._title = self._sub = ""
        self._done = self._hover = self._drag = False
        self.bind('<Configure>', lambda e: self._draw())
        self.bind('<Button-1>', lambda e: self._command())
        self.bind('<Enter>', lambda e: self._set('_hover', True))
        self.bind('<Leave>', lambda e: self._set('_hover', False))

    def _set(self, attr, value):
        setattr(self, attr, value)
        self._draw()

    def set_drag(self, on):
        self._set('_drag', bool(on))

    def set_text(self, title, sub, done=False):
        self._title, self._sub, self._done = title, sub, done
        self._draw()

    def _draw(self):
        try:
            self.delete('all')
            w, h = self.winfo_width(), self.winfo_height()
            if w < 40 or h < 40:
                return
            if self._drag:
                fill, line, fg = ZONE_DRAG_BG, ACCENT, ACCENT
            elif self._done:
                fill, line, fg = ZONE_DONE_BG, ZONE_DONE_LINE, GOOD
            elif self._hover:
                fill, line, fg = ZONE_HOVER_BG, ACCENT, ACCENT
            else:
                fill, line, fg = ZONE_IDLE_BG, ZONE_IDLE_LINE, NAVY
            self.create_rectangle(2, 2, w - 3, h - 3, fill=fill, outline=line, width=2,
                                  dash=None if self._done else (6, 4))
            self.create_text(w // 2, h // 2 - 12, text=self._title, fill=fg, width=w - 40,
                             font=(FONT, 12, 'bold'))
            self.create_text(w // 2, h // 2 + 14, text=self._sub, fill=MUTED, width=w - 40,
                             font=(FONT, 9))
        except tk.TclError:
            pass            # the widget was destroyed while the screen was being rebuilt


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

    def set(self, value, sub="", color=None):
        self.value.configure(text=value)
        self.sub.configure(text=sub, fg=color or MUTED)


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

    def set_data(self, values, title, color=None, selected=None, prev=None, prev_label=""):
        self.values, self.title, self.color, self.selected = list(values), title, (color or ACCENT), selected
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
            self.create_line(left, y, w - right, y, fill=GRID)
            self.create_text(left - 8, y, text=self._fmt(top_val * frac), anchor="e", fill=MUTED, font=(FONT, 8))
        slot = (w - left - right) / 12
        bw = slot * 0.6
        for i, v in enumerate(self.values):
            cx = left + slot * (i + .5)
            if self.hover == i + 1:
                self.create_rectangle(cx - slot / 2 + 2, top - 6, cx + slot / 2 - 2, top + plot_h + 2,
                                      fill=TRACK, outline='')
        for i, v in enumerate(self.values):
            cx = left + slot * (i + .5)
            if v > 0:
                y0 = top + plot_h * (1 - v / top_val)
                fill = NAVY if self.selected == i + 1 else self.color
                self.create_rectangle(cx - bw / 2, y0, cx + bw / 2, top + plot_h, fill=fill, outline='')
                self.create_text(cx, y0 - 8, text=self._fmt(v), fill=TEXT, font=(FONT, 8))
            else:
                self.create_line(cx - bw / 2, top + plot_h - 1, cx + bw / 2, top + plot_h - 1,
                                 fill=TICK, width=2)
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
        self.create_rectangle(x, y, x + tw, y + 14 * len(lines) + 10, fill=TIP_BG, outline='')
        self.create_text(x + 8, y + 6, text="\n".join(lines), anchor="nw", fill=TIP_FG,
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
    st.configure('TNotebook', tabmargins=(4, 6, 4, 0))
    st.configure('TFrame', background=BG)
    st.configure('Card.TFrame', background=CARD)
    st.configure('TLabel', background=BG, foreground=TEXT)
    st.configure('Card.TLabel', background=CARD)
    st.configure('CardBold.TLabel', background=CARD, font=(FONT, 10, 'bold'))
    st.configure('TSeparator', background=BORDER)
    st.configure('Status.TLabel', background=CARD, foreground=MUTED, font=(FONT, 9), padding=(14, 6))
    st.configure('TRadiobutton', background=BG, focuscolor=BG)
    st.configure('TRadiobutton', indicatorbackground=FIELD, indicatorforeground=TEXT)
    st.map('TRadiobutton', indicatorbackground=[('pressed', BTN), ('disabled', BG)])
    for name in ('Card.TRadiobutton', 'Card.TCheckbutton'):
        st.configure(name, background=CARD, focuscolor=CARD, indicatorbackground=FIELD, indicatorforeground=TEXT)
        st.map(name, background=[('active', CARD)], indicatorbackground=[('pressed', BTN), ('disabled', CARD)])
    st.configure('TCombobox', fieldbackground=FIELD, background=BTN, bordercolor=BORDER,
                 lightcolor=BORDER, darkcolor=BORDER, arrowcolor=TEXT, foreground=TEXT, padding=6)
    st.map('TCombobox', fieldbackground=[('readonly', FIELD)], bordercolor=[('focus', ACCENT)])
    root.option_add('*TCombobox*Listbox.font', (FONT, 10))
    root.option_add('*TCombobox*Listbox.background', FIELD)
    root.option_add('*TCombobox*Listbox.foreground', TEXT)
    root.option_add('*TCombobox*Listbox.selectBackground', ACCENT)
    root.option_add('*TCombobox*Listbox.selectForeground', 'white')

    # coloured text styles, one set for the grey background and one for white cards
    for p, bg in (('', BG), ('Card', CARD)):
        st.configure(f'{p}Hint.TLabel', background=bg, foreground=MUTED, font=(FONT, 9))
        st.configure(f'{p}Good.TLabel', background=bg, foreground=GOOD)
        st.configure(f'{p}Bad.TLabel', background=bg, foreground=BAD)
        st.configure(f'{p}Warn.TLabel', background=bg, foreground=WARN)
    st.configure('GoodTitle.TLabel', foreground=GOOD, font=(FONT, 15, 'bold'))
    st.configure('BadTitle.TLabel', foreground=WARN, font=(FONT, 15, 'bold'))

    # status "pills" shown above the main buttons
    for name, bg, fg in (('StatusGood', GOOD_BG, GOOD_FG), ('StatusBad', BAD_BG, BAD_FG),
                         ('StatusWarn', WARN_BG, WARN_FG)):
        st.configure(f'{name}.TLabel', background=bg, foreground=fg, padding=(12, 8))

    # tabs
    st.configure('TNotebook', background=BG, borderwidth=0, tabmargins=(2, 4, 2, 0))
    st.configure('TNotebook.Tab', padding=(20, 11), font=(FONT, 10, 'bold'), background=TAB_BG,
                 foreground=MUTED, borderwidth=0, focuscolor=TAB_BG)
    st.map('TNotebook.Tab', background=[('selected', ACCENT), ('active', TAB_HOVER)],
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
               background=[('disabled', OFF_BG), ('pressed', active), ('active', active)],
               foreground=[('disabled', OFF_FG)],
               bordercolor=[('disabled', OFF_BG), ('active', active)],
               lightcolor=[('pressed', active), ('active', active)],
               darkcolor=[('pressed', active), ('active', active)])
    button('TButton', BTN, TEXT, BTN_HOVER, border=BORDER)
    button('Primary.TButton', ACCENT, 'white', ACCENT_DARK, font=(FONT, 11, 'bold'), padding=(16, 13))
    # the second action is an outlined button so the main one stands out
    button('Print.TButton', CARD, ACCENT, ACCENT_SOFT, border=BORDER, font=(FONT, 11, 'bold'), padding=(16, 13))
    button('PrintSm.TButton', TEAL, 'white', TEAL_DARK, font=(FONT, 10, 'bold'), padding=(14, 7))
    button('Danger.TButton', BAD_BG, BAD, DANGER_HOVER, border=DANGER_LINE, padding=(14, 7))

    # entries
    st.configure('TEntry', fieldbackground=FIELD, foreground=TEXT, insertcolor=TEXT, bordercolor=BORDER, lightcolor=BORDER,
                 darkcolor=BORDER, padding=8)
    st.map('TEntry', bordercolor=[('focus', ACCENT)], lightcolor=[('focus', ACCENT)],
           darkcolor=[('focus', ACCENT)])
    st.configure('Bad.TEntry', fieldbackground=BAD_BG, bordercolor=BAD_LINE,
                 lightcolor=BAD_LINE, darkcolor=BAD_LINE, padding=8)

    # scrollbars
    st.configure('Vertical.TScrollbar', background=SCROLL, troughcolor=BG, bordercolor=BG,
                 arrowcolor=MUTED, relief='flat')

    # result table
    st.configure('Treeview', rowheight=30, font=(FONT, 10), background=CARD, fieldbackground=CARD,
                 foreground=TEXT, bordercolor=BORDER, borderwidth=1)
    st.configure('Treeview.Heading', font=(FONT, 10, 'bold'), background=HEAD_BG, foreground=NAVY,
                 bordercolor=BORDER, padding=(8, 8), relief='flat')
    st.map('Treeview.Heading', background=[('active', HEAD_HOVER)])
    st.map('Treeview', background=[('selected', SEL)], foreground=[('selected', TEXT)])


# ------------------------------------------------------------------ user guide
# Every text is a (greek, english) pair. **bold** and `Key` are marked inline.
# Block types: p paragraph, h sub-heading, li bullet, step numbered step, kb key + description,
#              tip / warn callout, go link  ('tab:<name>' opens a tab, 'sec:<id>' another section).
MANUAL = [
    dict(id='start', title=("Γρήγορη εκκίνηση", "Quick start"), blocks=[
        ('p', ("Το πρόγραμμα παίρνει το αρχείο Excel που βγάζει το Protel και συμπληρώνει αυτόματα το μηνιαίο "
               "δελτίο της ΕΛΣΤΑΤ σε PDF. Ο μήνας ολοκληρώνεται σε πέντε βήματα:",
               "The program takes the Excel file that Protel produces and fills in the monthly ELSTAT form as a PDF "
               "automatically. A month takes five steps:")),
        ('live', 'next'),
        ('step', ("Την πρώτη φορά, ανοίξτε την καρτέλα **Στοιχεία ξενοδοχείου** και συμπληρώστε τα στοιχεία σας. "
                  "Γίνεται μία φορά και το πρόγραμμα τα θυμάται.",
                  "The first time, open the **Hotel details** tab and fill in your details. You do this once; "
                  "the program remembers them.")),
        ('step', ("Στην καρτέλα **Δελτίο μήνα**, σύρτε το Excel του μήνα μέσα στο πλαίσιο του Βήματος 1 ή πατήστε "
                  "πάνω του για να το διαλέξετε.",
                  "In the **Monthly form** tab, drop the month's Excel into the box in Step 1, or click it to choose the file.")),
        ('step', ("Ελέγξτε τα **Στοιχεία του μήνα** (Βήμα 2). Τα περισσότερα συμπληρώνονται μόνα τους - εσείς "
                  "πληκτρολογείτε μόνο το **Σύνολο εργαζομένων**.",
                  "Check the **Month details** (Step 2). Most are filled in for you - you only type the **Total employees**.")),
        ('step', ("Δείτε πού θα αποθηκευτεί το PDF (Βήμα 3) και πατήστε **Δημιουργία PDF** ή **Δημιουργία & Εκτύπωση**.",
                  "See where the PDF will be saved (Step 3) and click **Create PDF** or **Create & Print**.")),
        ('step', ("Στο παράθυρο αποτελέσματος ελέγξτε τα σύνολα και τις προειδοποιήσεις και μετά εκτυπώστε ή ανοίξτε το PDF.",
                  "In the result window check the totals and any warnings, then print or open the PDF.")),
        ('tip', ("Κάθε δελτίο αρχειοθετείται αυτόματα. Μπορείτε να το ξανανοίξετε ή να το εκτυπώσετε όποτε θέλετε από το **Ιστορικό**.",
                 "Every form is archived automatically. You can reopen or print it whenever you like from **History**.")),
        ('go', ("Ξεκινήστε από το Δελτίο μήνα", "Go to the Monthly form"), 'tab:report'),
    ]),
    dict(id='status', title=("Κατάσταση συστήματος", "Live status"), blocks=[
        ('p', ("Αυτή η σελίδα ελέγχει το πρόγραμμα αυτή τη στιγμή και ανανεώνεται κάθε φορά που την ανοίγετε. "
               "Πατήστε σε μια γραμμή για να πάτε εκεί που διορθώνεται.",
               "This page checks the program right now and refreshes every time you open it. "
               "Click a line to jump to where it is fixed.")),
        ('live', 'next'),
        ('h', ("Έλεγχοι", "Checks")),
        ('live', 'status'),
        ('h', ("Γρήγορες ενέργειες", "Quick actions")),
        ('act', ("Επιλογή Excel του μήνα", "Choose the month's Excel"), 'choose_excel'),
        ('act', ("Άνοιγμα του επεξεργαστή διάταξης", "Open the layout editor"), 'open_editor'),
        ('act', ("Άνοιγμα του φακέλου δεδομένων", "Open the data folder"), 'open_data'),
        ('act', ("Αντίγραφο ασφαλείας (.zip)…", "Backup (.zip)…"), 'backup'),
        ('act', ("Εναλλαγή εμφάνισης (φωτεινή / σκούρα)", "Switch appearance (light / dark)"), 'theme'),
        ('h', ("Τρέχουσες ρυθμίσεις", "Current settings")),
        ('li', ("Μέθοδος συμπλήρωσης PDF: **[[pdf_mode_name]]**", "PDF filling method: **[[pdf_mode_name]]**")),
        ('li', ("Φάκελος δεδομένων: `[[data_dir]]`", "Data folder: `[[data_dir]]`")),
        ('li', ("Το επόμενο PDF θα αποθηκευτεί ως: `[[next_pdf]]`", "The next PDF will be saved as: `[[next_pdf]]`")),
        ('li', ("Template: `[[template_name]]`", "Template: `[[template_name]]`")),
        ('li', ("Γλώσσα: [[language_name]] · Εμφάνιση: [[theme_name]]", "Language: [[language_name]] · Appearance: [[theme_name]]")),
        ('li', ("Δελτία στο ιστορικό: **[[months_count]]** · Τελευταίο: [[last_month]]",
                "Forms in the history: **[[months_count]]** · Latest: [[last_month]]")),
        ('li', ("Δικές σας αντιστοιχίσεις χωρών: **[[custom_count]]**", "Your own country mappings: **[[custom_count]]**")),
    ]),
    dict(id='monthly', title=("Δελτίο μήνα", "Monthly form"), blocks=[
        ('p', ("Η καρτέλα που θα χρησιμοποιείτε κάθε μήνα. Έχει τρία βήματα και, κάτω, τα κουμπιά δημιουργίας.",
               "The tab you will use every month. It has three steps and, at the bottom, the create buttons.")),
        ('h', ("Βήμα 1 - Αρχείο από το Protel", "Step 1 - File from Protel")),
        ('p', ("Το αρχείο που χρειάζεται λέγεται «Nationalities Statistics» και βγαίνει από το Protel ως .xls ή .xlsx. "
               "Μπορείτε να το διαλέξετε με όποιον τρόπο σας βολεύει:",
               "The file you need is called “Nationalities Statistics” and comes from Protel as .xls or .xlsx. "
               "Choose it whichever way suits you:")),
        ('li', ("Σύρτε το πάνω στο διακεκομμένο πλαίσιο (ή πάνω σε ολόκληρη την κάρτα).",
                "Drop it onto the dashed box (or anywhere on the card).")),
        ('li', ("Πατήστε το πλαίσιο ή το κουμπί **Επιλογή Excel…** (`Ctrl+O`).",
                "Click the box or the **Choose Excel…** button (`Ctrl+O`).")),
        ('li', ("Πατήστε **Πιο πρόσφατο** για να διαλέξει το πρόγραμμα το νεότερο Excel από τον φάκελο που χρησιμοποιήσατε τελευταία.",
                "Click **Latest** and the program picks the newest Excel in the folder you used last.")),
        ('li', ("Ή γράψτε / επικολλήστε τη διαδρομή του αρχείου στο πεδίο.",
                "Or type / paste the file path into the field.")),
        ('p', ("Μόλις διαλέξετε αρχείο, το πρόγραμμα το διαβάζει: δείχνει την περίοδο που βρήκε και έναν πίνακα με τις χώρες. "
               "Η στήλη **Στο δελτίο** λέει σε ποια γραμμή του δελτίου θα μπει κάθε χώρα. Οι κόκκινες γραμμές χρειάζονται προσοχή.",
               "As soon as you choose a file the program reads it: it shows the period it found and a table of the countries. "
               "The **On the form** column says which line of the form each country goes to. Red lines need your attention.")),
        ('h', ("Βήμα 2 - Στοιχεία του μήνα", "Step 2 - Month details")),
        ('p', ("Το Έτος, ο Μήνας, η Τελευταία ημέρα («Έως») και οι Ημέρες λειτουργίας συμπληρώνονται από το Excel. "
               "Η Ημερομηνία υπογραφής παίρνει τη σημερινή - αλλάξτε την (ΗΗ/ΜΜ/ΕΕΕΕ) αν το δελτίο πρέπει να φαίνεται υπογεγραμμένο άλλη μέρα. "
               "Ένα πεδίο που κοκκινίζει έχει λάθος τιμή.",
               "The Year, Month, Last day (“To”) and Operating days are filled in from the Excel. "
               "The Signing date is today's date - change it (DD/MM/YYYY) if the form should look signed on another day. "
               "A field that turns red has a wrong value.")),
        ('warn', ("Το **Σύνολο εργαζομένων** ΔΕΝ υπάρχει στο Excel. Πληκτρολογήστε το εσείς, στο τέλος του μήνα, μαζί με ιδιοκτήτες και οικογένεια. "
                  "Στις Ρυθμίσεις μπορείτε να ζητήσετε να προσυμπληρώνεται με την τελευταία τιμή.",
                  "The **Total employees** is NOT in the Excel. Type it yourself at the end of the month, including owners and family. "
                  "In Settings you can ask for it to be pre-filled with the last value.")),
        ('h', ("Βήμα 3 - Πού αποθηκεύεται το PDF", "Step 3 - Where the PDF is saved")),
        ('p', ("Η θέση προτείνεται αυτόματα με βάση τις Ρυθμίσεις (φάκελος και όνομα αρχείου). Για άλλη θέση πατήστε **Αλλαγή…**.",
               "The location is suggested automatically from the Settings (folder and file name). For another location click **Change…**.")),
        ('li', ("Αυτή τη στιγμή προτείνεται: `[[next_pdf]]`", "Right now it suggests: `[[next_pdf]]`")),
        ('h', ("Δημιουργία του δελτίου", "Creating the form")),
        ('p', ("Πάνω από τα κουμπιά υπάρχει πάντα ένα μήνυμα ελέγχου. Αν κάτι λείπει ή είναι λάθος, το μήνυμα το απαριθμεί και τα κουμπιά "
               "είναι γκρι. Όταν όλα είναι σωστά γράφει «Όλα είναι έτοιμα».",
               "There is always a check message above the buttons. If something is missing or wrong it lists it and the buttons are grey. "
               "When everything is right it says “Everything is ready”.")),
        ('li', ("**Δημιουργία PDF** (`Ctrl+Enter`) φτιάχνει το δελτίο.",
                "**Create PDF** (`Ctrl+Enter`) makes the form.")),
        ('li', ("**Δημιουργία & Εκτύπωση** (`Ctrl+P`) το φτιάχνει και το στέλνει στον προεπιλεγμένο εκτυπωτή.",
                "**Create & Print** (`Ctrl+P`) makes it and sends it to the default printer.")),
        ('h', ("Τι ελέγχει το πρόγραμμα", "What the program checks")),
        ('li', ("Αν η περίοδος του Excel διαφέρει από τον μήνα και το έτος της φόρμας, ρωτά αν θέλετε να συνεχίσετε.",
                "If the Excel's period differs from the form's month and year, it asks whether to continue.")),
        ('li', ("Αν το PDF υπάρχει ήδη, ρωτά πριν το αντικαταστήσει (η ερώτηση κλείνει από τις Ρυθμίσεις).",
                "If the PDF already exists it asks before replacing it (this question can be turned off in Settings).")),
        ('li', ("Μετά τη δημιουργία συγκρίνει το γενικό σύνολο του PDF με το σύνολο του Excel και σας προειδοποιεί αν διαφέρουν.",
                "After creating it compares the PDF's grand total with the Excel's total and warns you if they differ.")),
        ('li', ("Αν μια χώρα δεν είναι γνωστή στο πρόγραμμα, σας λέει ποια δεν μπήκε στο PDF.",
                "If a country is unknown to the program it tells you which one was left out of the PDF.")),
        ('warn', ("Όταν υπάρχουν προειδοποιήσεις, η αυτόματη εκτύπωση δεν γίνεται. Ελέγξτε το δελτίο και μετά πατήστε **Εκτύπωση** στο παράθυρο αποτελέσματος.",
                  "When there are warnings, automatic printing does not happen. Check the form, then click **Print** in the result window.")),
        ('h', ("Υπενθύμιση", "Reminder")),
        ('p', ("Αν λείπει το δελτίο του προηγούμενου μήνα, εμφανίζεται κίτρινη υπενθύμιση στην κορυφή. Κλείνει με το **Κλείσιμο** "
               "ή απενεργοποιείται από τις Ρυθμίσεις.",
               "If last month's form is missing, a yellow reminder appears at the top. Close it with **Close** or turn it off in Settings.")),
        ('go', ("Άνοιγμα της καρτέλας Δελτίο μήνα", "Open the Monthly form tab"), 'tab:report'),
    ]),
    dict(id='hotel', title=("Στοιχεία ξενοδοχείου", "Hotel details"), blocks=[
        ('p', ("Τα στοιχεία που γράφονται σε κάθε δελτίο. Τα συμπληρώνετε μία φορά και το πρόγραμμα τα θυμάται.",
               "The details written on every form. You fill them in once and the program remembers them.")),
        ('li', ("**Είδος καταλύματος, Τίτλος, Κατηγορία**: όπως θέλετε να γραφτούν στο δελτίο.",
                "**Property type, Name, Category**: as you want them written on the form.")),
        ('li', ("**E-mail** και **Τηλέφωνο**.", "**E-mail** and **Phone**.")),
        ('li', ("**Α.Φ.Μ.**: ακριβώς 9 ψηφία. Δέχεται μόνο αριθμούς και ο μετρητής δείχνει ✓ όταν είναι έγκυρο.",
                "**Tax ID (AFM)**: exactly 9 digits. It accepts numbers only and the counter shows ✓ when it is valid.")),
        ('li', ("**Επωνυμία επιχείρησης**.", "**Business name**.")),
        ('tip', ("Αλλάξτε ό,τι θέλετε εδώ: τα στοιχεία αποθηκεύονται αυτόματα όταν δημιουργήσετε δελτίο. Ένα πεδίο που κοκκινίζει πρέπει να διορθωθεί.",
                 "Change anything here: the details are saved automatically when you create a form. A field that turns red must be corrected.")),
        ('h', ("Άδειο PDF της ΕΛΣΤΑΤ (template)", "Blank ELSTAT PDF (template)")),
        ('p', ("Το κενό δελτίο πάνω στο οποίο γράφονται τα στοιχεία (`template_elstat_report.pdf`, δίπλα στο πρόγραμμα). "
               "Συνήθως δεν χρειάζεται να το αλλάξετε. Αν η ΕΛΣΤΑΤ δώσει νέο έντυπο, διαλέξτε το με **Αναζήτηση…** ή σύρτε το πάνω στην κάρτα.",
               "The blank form on which the details are written (`template_elstat_report.pdf`, next to the program). "
               "You usually don't need to change it. If ELSTAT issues a new form, choose it with **Browse…** or drop it onto the card.")),
        ('go', ("Άνοιγμα της καρτέλας Στοιχεία ξενοδοχείου", "Open the Hotel details tab"), 'tab:hotel'),
    ]),
    dict(id='history', title=("Ιστορικό", "History"), blocks=[
        ('p', ("Κάθε δελτίο που δημιουργείτε καταχωρείται στο Ιστορικό. Εκεί βλέπετε στατιστικά, συγκρίσεις με το προηγούμενο έτος "
               "και ξανανοίγετε παλιά δελτία.",
               "Every form you create is recorded in History. There you see statistics, comparisons with the previous year, and reopen old forms.")),
        ('h', ("Φίλτρα", "Filters")),
        ('li', ("**Έτος** και **Χώρα** περιορίζουν τα στοιχεία που βλέπετε.", "**Year** and **Country** narrow down what you see.")),
        ('li', ("**Διανυκτερεύσεις / Πελάτες** επιλέγει τι δείχνει το γράφημα.", "**Overnight stays / Guests** chooses what the chart shows.")),
        ('h', ("Τι βλέπετε", "What you see")),
        ('li', ("Τέσσερα πλακίδια: **Πελάτες, Διανυκτερεύσεις, Μέση διαμονή** (σε νύχτες) και **Αλλοδαποί**.",
                "Four tiles: **Guests, Overnight stays, Average stay** (in nights) and **Foreign guests**.")),
        ('li', ("Γράφημα με μία στήλη ανά μήνα. Αν υπάρχουν στοιχεία του προηγούμενου έτους, εμφανίζονται ως μικρές κίτρινες γραμμές. "
                "Περάστε το ποντίκι για λεπτομέρειες και πατήστε μια στήλη για να επιλέξετε τον μήνα.",
                "A chart with one bar per month. If the previous year has data it appears as small yellow ticks. "
                "Hover for details and click a bar to select the month.")),
        ('li', ("Πίνακας Έτος / Μήνας / Χώρα με πελάτες, διανυκτερεύσεις, μέση διαμονή, ποσοστό και στήλη PDF. "
                "**Ανάπτυξη** και **Σύμπτυξη** ανοίγουν ή κλείνουν τα πάντα.",
                "A Year / Month / Country table with guests, nights, average stay, share and a PDF column. "
                "**Expand** and **Collapse** open or close everything.")),
        ('li', ("Δεξιά βλέπετε τις κορυφαίες χώρες. Διπλό κλικ σε χώρα την κάνει φίλτρο - δεύτερο διπλό κλικ το καθαρίζει.",
                "On the right are the top countries. Double-click a country to filter by it - double-click again to clear.")),
        ('h', ("Ενέργειες", "Actions")),
        ('li', ("Σε επιλεγμένο μήνα: **Άνοιγμα** (ή διπλό κλικ, ή `Enter`), **Αποθήκευση ως…** και **Εκτύπωση**.",
                "On a selected month: **Open** (or double-click, or `Enter`), **Save as…** and **Print**.")),
        ('li', ("**Σύνοψη έτους**: PDF μίας σελίδας για ολόκληρο το έτος - αποθήκευση ή εκτύπωση.",
                "**Year summary**: a one-page PDF for the whole year - save or print.")),
        ('li', ("**Διαγραφή μήνα** (`Delete`) τον αφαιρεί από το ιστορικό και σβήνει το αρχειοθετημένο PDF του. Το αρχείο που αποθηκεύσατε εσείς δεν πειράζεται.",
                "**Delete month** (`Delete`) removes it from the history and deletes its archived PDF. The file you saved yourself is not touched.")),
        ('li', ("Δεξί κλικ σε μήνα ανοίγει μενού με όλες τις ενέργειες.", "Right-click a month for a menu with all the actions.")),
        ('li', ("**Εξαγωγή Excel…** αποθηκεύει όλο το ιστορικό σε αρχείο Excel και **Άνοιγμα CSV** ανοίγει το αρχείο του ιστορικού.",
                "**Export Excel…** saves the whole history as an Excel file and **Open CSV** opens the history file.")),
        ('tip', ("Αν ένας παλιός μήνας δεν έχει αρχειοθετημένο PDF, το πρόγραμμα ξαναφτιάχνει το δελτίο από τα δεδομένα του ιστορικού όταν το ζητήσετε. "
                 "Χρειάζονται τα Στοιχεία ξενοδοχείου και το template.",
                 "If an old month has no archived PDF, the program rebuilds the form from the history data when you ask for it. "
                 "It needs the Hotel details and the template.")),
        ('go', ("Άνοιγμα της καρτέλας Ιστορικό", "Open the History tab"), 'tab:history'),
    ]),
    dict(id='settings', title=("Ρυθμίσεις", "Settings"), blocks=[
        ('p', ("Όλα όσα προσαρμόζουν το πρόγραμμα. Οι αλλαγές αποθηκεύονται αμέσως.",
               "Everything that customises the program. Changes are saved immediately.")),
        ('h', ("Γλώσσα και εμφάνιση", "Language and appearance")),
        ('li', ("**Γλώσσα**: Ελληνικά ή English, αλλάζει αμέσως. Το επίσημο δελτίο (PDF) γράφεται πάντα στα ελληνικά.",
                "**Language**: Ελληνικά or English, changes immediately. The official form (PDF) is always written in Greek.")),
        ('li', ("**Εμφάνιση**: Φωτεινή, Σκούρα ή Όπως το σύστημα. Το κουμπί ☾ / ☀ πάνω δεξιά την αλλάζει γρήγορα.",
                "**Appearance**: Light, Dark or Match system. The ☾ / ☀ button at the top right switches it quickly.")),
        ('h', ("Πού αποθηκεύονται τα PDF", "Where PDFs are saved")),
        ('p', ("Διαλέξτε τον φάκελο του προγράμματος, τον φάκελο του Excel ή έναν δικό σας φάκελο. "
               "Το όνομα του αρχείου φτιάχνεται από ένα μοτίβο, π.χ. `ELSTAT_{year}_{month}` γίνεται ELSTAT_2026_03. "
               "Πρέπει να περιέχει και το {year} και το {month}, αλλιώς κάθε μήνας θα γράφει πάνω στον προηγούμενο.",
               "Choose the program folder, the Excel's folder or a folder of your own. "
               "The file name is built from a pattern, e.g. `ELSTAT_{year}_{month}` becomes ELSTAT_2026_03. "
               "It must contain both {year} and {month}, otherwise each month would overwrite the previous one.")),
        ('h', ("Φάκελος δεδομένων", "Data folder")),
        ('p', ("Εδώ αποθηκεύονται τα στοιχεία του ξενοδοχείου, το ιστορικό, τα αρχειοθετημένα PDF και οι ρυθμίσεις.",
               "This is where the hotel details, history, archived PDFs and settings are stored.")),
        ('li', ("**Αλλαγή…** διαλέγει άλλον φάκελο και προσφέρεται να αντιγράψει τα υπάρχοντα δεδομένα. **Άνοιγμα** τον ανοίγει στον Εξερευνητή.",
                "**Change…** picks another folder and offers to copy the existing data. **Open** shows it in the file explorer.")),
        ('li', ("**Προεπιλεγμένος φάκελος** επιστρέφει στη θέση δίπλα στο πρόγραμμα.",
                "**Default folder** goes back to the location next to the program.")),
        ('li', ("**Αντίγραφο ασφαλείας (.zip)…** αποθηκεύει όλα τα δεδομένα σε ένα αρχείο.",
                "**Backup (.zip)…** saves all the data into one file.")),
        ('tip', ("Διαλέξτε έναν φάκελο μέσα στο OneDrive ή σε άλλη υπηρεσία συγχρονισμού και θα έχετε αυτόματο αντίγραφο ασφαλείας.",
                 "Pick a folder inside OneDrive or another sync service and you get an automatic backup.")),
        ('h', ("Αρχεία Excel από το Protel", "Excel files from Protel")),
        ('p', ("Από ποιον φάκελο ξεκινούν τα **Επιλογή Excel…** και **Πιο πρόσφατο**. Κενό σημαίνει ο φάκελος του τελευταίου Excel που χρησιμοποιήσατε.",
               "The folder where **Choose Excel…** and **Latest** start. Empty means the folder of the last Excel you used.")),
        ('h', ("Τρόπος συμπλήρωσης του PDF (η έξυπνη διάταξη είναι πειραματική)",
               "How the PDF is filled in (smart layout is experimental)")),
        ('p', ("**Πάνω στο template** (προεπιλογή) γράφει τις τιμές σε σταθερές θέσεις πάνω στο άδειο PDF. "
               "**Έξυπνη διάταξη** (πειραματική) διαβάζει το template, βρίσκει πού βρίσκεται τώρα κάθε τμήμα του και "
               "τοποθετεί τις τιμές αναλόγως - χρήσιμο αν η ΕΛΣΤΑΤ αλλάξει λίγο το έντυπο (μετατοπισμένες γραμμές, περιθώρια, αποστάσεις).",
               "**On top of the template** (default) writes the values at fixed positions on the blank PDF. "
               "**Smart layout** (experimental) reads the template, finds where each part of it now sits and places "
               "the values accordingly - useful if ELSTAT changes the form slightly (shifted lines, margins, spacing).")),
        ('warn', ("Η έξυπνη διάταξη είναι πειραματική λειτουργία. Ελέγχετε πάντα το δελτίο πριν το υποβάλετε. "
                  "Αν κάτι δεν πάει καλά, επιλέξτε ξανά «Πάνω στο template».",
                  "Smart layout is an experimental feature. Always check the form before submitting it. "
                  "If anything looks wrong, select “On top of the template” again.")),
        ('step', ("Επιλέξτε «Έξυπνη διάταξη». Τα πρόσθετα κουμπιά εμφανίζονται μόνο όταν είναι επιλεγμένη.",
                  "Select “Smart layout”. The extra buttons appear only while it is selected.")),
        ('step', ("Αποθηκεύστε το σημείο σύγκρισης: το πρόγραμμα το κάνει μόνο του με το πρώτο δελτίο που φτιάχνετε με την "
                  "προεπιλεγμένη μέθοδο, ή πατήστε **Αποθήκευση ως αναφορά…** όταν το δελτίο συμπληρώνεται σωστά.",
                  "Save the point of comparison: the program does it by itself with the first form you make in the default "
                  "mode, or press **Save as reference…** when the form fills in correctly.")),
        ('step', ("Όταν αλλάξει το template, πατήστε **Έλεγχος template** (δείχνει αν αναγνωρίζεται και πόσο μετακινήθηκε) "
                  "και **Δοκιμαστικό δελτίο** (PDF με πλασματικά στοιχεία για οπτικό έλεγχο - μην το υποβάλετε).",
                  "When the template changes, press **Check template** (shows whether it is recognised and how far it moved) "
                  "and **Test form** (a PDF with made-up data for a visual check - do not submit it).")),
        ('tip', ("Αν το template δεν αναγνωριστεί με ασφάλεια ή έχει μετακινηθεί πάρα πολύ, το δελτίο φτιάχνεται με την "
                 "προεπιλεγμένη μέθοδο ή εμφανίζεται προειδοποίηση και δεν γίνεται αυτόματη εκτύπωση.",
                 "If the template is not recognised safely or has moved a lot, the form is made with the default method "
                 "or a warning appears and nothing is printed automatically.")),
        ('p', ("**Χειροκίνητη τοποθέτηση**: επιλέξτε την και ανοίγει ο **επεξεργαστής διάταξης**. Βλέπετε το έντυπο ως σελίδα· "
               "κάθε στοιχείο (ΑΦΜ, ημερομηνίες, σύνολα, στήλες χωρών…) είναι ένα κείμενο που το σέρνετε εκεί που ανήκει.",
               "**Set manually**: select it and the **layout editor** opens. You see the form as a page; "
               "every item (VAT number, dates, totals, country columns…) is a piece of text you drag to where it belongs.")),
        ('step', ("Πατήστε ένα στοιχείο στη σελίδα ή στη λίστα. Δεξιά αλλάζετε **γραμματοσειρά, μέγεθος, χρώμα, στοίχιση** και την ακριβή θέση. "
                  "Τα βελάκια του πληκτρολογίου μετακινούν ανά 0,5 pt (με Shift ανά 5 pt).",
                  "Click an item on the page or in the list. On the right you change the **font, size, colour, alignment** and the exact position. "
                  "The arrow keys move it by 0.5 pt (5 pt with Shift).")),
        ('step', ("Οι στήλες χωρών, ο Α.Φ.Μ., οι ημέρες λειτουργίας και οι εργαζόμενοι μετακινούνται αρχικά ως σύνολο. "
                  "Πατήστε **Χωρισμός σε μεμονωμένα στοιχεία** για να μετακινείτε κάθε χώρα ή κάθε ψηφίο ξεχωριστά (**Συγχώνευση** για να ξαναενωθούν).",
                  "The country columns, the VAT number, operating days and employees first move as one. "
                  "Press **Split into individual items** to move each country or digit on its own (**Merge** puts them back together).")),
        ('step', ("Με το **+ Προσθήκη στοιχείου…** ψάχνετε ό,τι μπορεί να μπει στο δελτίο (στοιχεία ξενοδοχείου, σύνολα, ημερομηνίες, αριθμοί συγκεκριμένης χώρας) "
                  "και το τοποθετείτε όσες φορές θέλετε. Το πεδίο **Αναζήτηση** φιλτράρει τη λίστα στοιχείων. Ctrl+Z / Ctrl+Y = αναίρεση / επανάληψη, Ctrl+D = διπλασιασμός.",
                  "**+ Add item…** lets you search for anything that can go on the form (hotel details, totals, dates, numbers of a specific country) "
                  "and place it as many times as you like. The **Search** box filters the item list. Ctrl+Z / Ctrl+Y = undo / redo, Ctrl+D = duplicate.")),
        ('step', ("Πατήστε **Δοκιμαστικό PDF** για να δείτε το αποτέλεσμα με πλασματικά στοιχεία και **Αποθήκευση** (Ctrl+S) όταν είστε έτοιμοι.",
                  "Press **Test PDF** to see the result with made-up data and **Save** (Ctrl+S) when you are happy.")),
        ('go', ("Δείτε τον πλήρη οδηγό του επεξεργαστή", "See the full layout editor guide"), 'sec:editor'),
        ('h', ("Συμπεριφορά του προγράμματος", "Program behaviour")),
        ('li', ("Άνοιγμα του PDF αμέσως μετά τη δημιουργία.", "Open the PDF right after it is created.")),
        ('li', ("Ερώτηση πριν αντικατασταθεί PDF που υπάρχει ήδη.", "Ask before replacing a PDF that already exists.")),
        ('li', ("Αρχειοθέτηση αντιγράφου κάθε δελτίου - επιτρέπει το ξανάνοιγμα και την εκτύπωση παλιών μηνών από το Ιστορικό.",
                "Archive a copy of every form - lets you reopen and print old months from History.")),
        ('li', ("Υπενθύμιση όταν λείπει το δελτίο του προηγούμενου μήνα.", "Remind me when last month's form is missing.")),
        ('li', ("Προσυμπλήρωση του «Σύνολο εργαζομένων» με την τελευταία τιμή - ελέγχετε πάντα τον αριθμό πριν δημιουργήσετε το δελτίο.",
                "Pre-fill “Total employees” with the last value - always check the number before creating the form.")),
        ('p', ("**Επαναφορά ρυθμίσεων** γυρίζει τις ρυθμίσεις στις προεπιλογές. Δεν διαγράφει ιστορικό ή στοιχεία ξενοδοχείου.",
               "**Reset settings** returns the settings to their defaults. It does not delete the history or the hotel details.")),
        ('go', ("Άνοιγμα της καρτέλας Ρυθμίσεις", "Open the Settings tab"), 'tab:settings'),
    ]),
    dict(id='editor', title=("Επεξεργαστής διάταξης", "Layout editor"), blocks=[
        ('p', ("Ο επεξεργαστής διάταξης σας αφήνει να τοποθετήσετε μόνοι σας κάθε τιμή πάνω στο έντυπο της ΕΛΣΤΑΤ - "
               "θέση, γραμματοσειρά, μέγεθος, χρώμα, στοίχιση. Ανοίγει όταν πατήσετε **Άνοιγμα επεξεργαστή διάταξης** στις Ρυθμίσεις.",
               "The layout editor lets you place every value on the ELSTAT form yourself - position, font, size, colour, alignment. "
               "It opens when you press **Open layout editor** in Settings.")),
        ('act', ("Άνοιγμα του επεξεργαστή διάταξης", "Open the layout editor"), 'open_editor'),
        ('li', ("Τρέχουσα μέθοδος συμπλήρωσης: **[[pdf_mode_name]]**", "Current filling method: **[[pdf_mode_name]]**")),
        ('h', ("Η οθόνη", "The screen")),
        ('li', ("**Αριστερά**: η λίστα όλων των στοιχείων, με πεδίο αναζήτησης από πάνω.", "**Left**: the list of every item, with a search box on top.")),
        ('li', ("**Στο κέντρο**: η σελίδα του εντύπου. Κάθε τιμή φαίνεται εκεί που θα τυπωθεί. Το κουμπί **Σελίδα 1 / Σελίδα 2** αλλάζει σελίδα, τα **− / + / Προσαρμογή** το μέγεθος (Ctrl + τροχός = ζουμ).",
                "**Centre**: the page of the form. Every value appears where it will be printed. **Page 1 / Page 2** switches the page, **− / + / Fit** the size (Ctrl + wheel = zoom).")),
        ('li', ("**Δεξιά**: οι ιδιότητες του επιλεγμένου στοιχείου. Στο κάτω μέρος του παραθύρου μια μπάρα συμβουλών αλλάζει ανάλογα με το τι έχετε επιλέξει.",
                "**Right**: the properties of the selected item. At the bottom of the window a tip bar changes with what you have selected.")),
        ('h', ("Επιλογή και μετακίνηση", "Selecting and moving")),
        ('step', ("Πατήστε ένα στοιχείο πάνω στη σελίδα ή στη λίστα.", "Click an item on the page or in the list.")),
        ('step', ("Σύρτε το με το ποντίκι ή χρησιμοποιήστε τα βελάκια του πληκτρολογίου: 0,5 pt ανά πάτημα, 5 pt με Shift.",
                  "Drag it with the mouse or use the arrow keys: 0.5 pt per press, 5 pt with Shift.")),
        ('step', ("Για ακριβή θέση γράψτε τιμές στα **X** και **Y**, ή χρησιμοποιήστε τα βελάκια ◀ ▲ ▼ ▶ με επιλεγμένο βήμα (0,1 έως 5 pt).",
                  "For an exact position type values in **X** and **Y**, or use the ◀ ▲ ▼ ▶ buttons with a chosen step (0.1 to 5 pt).")),
        ('p', ("Οι συντεταγμένες είναι σε σημεία PDF (pt) από την κάτω αριστερή γωνία της σελίδας - 72 pt = 1 ίντσα. Καθώς περνάτε το ποντίκι πάνω από τη σελίδα, το X και το Y του δείκτη φαίνονται κάτω αριστερά.",
               "Coordinates are in PDF points (pt) from the bottom-left corner of the page - 72 pt = 1 inch. As you move the mouse over the page, the pointer's X and Y show at the bottom left.")),
        ('h', ("Μορφοποίηση", "Formatting")),
        ('li', ("**Γραμματοσειρά**: προσφέρονται μόνο όσες υπάρχουν στον υπολογιστή σας (όλες υποστηρίζουν ελληνικά).", "**Font**: only the ones installed on your computer are offered (all support Greek).")),
        ('li', ("**Μέγεθος, Έντονα, Πλάγια, Χρώμα** (παλέτα, κωδικός #rrggbb ή γρήγορα χρώματα) και **Στοίχιση**.", "**Size, Bold, Italic, Colour** (palette, #rrggbb code or quick colours) and **Align**.")),
        ('li', ("**Απόσταση ψηφίων**: 0 = κανονικό κείμενο. Μεγαλύτερη τιμή βάζει κάθε γράμμα σε σταθερή απόσταση, για τα κουτάκια του Α.Φ.Μ., των ημερών και των εργαζομένων.",
                "**Digit spacing**: 0 = normal text. A larger value puts every character at a fixed distance, for the boxes of the VAT number, days and employees.")),
        ('li', ("**Εμφάνιση στο PDF**: αν το ξετσεκάρετε, το στοιχείο κρύβεται (στη λίστα σημειώνεται με ○).", "**Show in the PDF**: untick it to hide the item (marked ○ in the list).")),
        ('h', ("Μεμονωμένα στοιχεία", "Individual items")),
        ('p', ("Κάποια στοιχεία κινούνται αρχικά ως ομάδα: οι στήλες χωρών (πελάτες και διανυκτερεύσεις, ανά σελίδα), ο Α.Φ.Μ., οι ημέρες λειτουργίας, οι εργαζόμενοι και το έτος. "
               "Για να τα ελέγχετε ένα-ένα:",
               "Some items first move as a group: the country columns (guests and nights, per page), the VAT number, operating days, employees and the year. "
               "To control them one by one:")),
        ('step', ("Επιλέξτε το στοιχείο και πατήστε **Χωρισμός σε μεμονωμένα στοιχεία**.", "Select the item and press **Split into individual items**.")),
        ('step', ("Στη λίστα εμφανίζονται τα μέλη του (κάθε χώρα ή κάθε ψηφίο) κάτω από το στοιχείο - ανοίξτε το με το βελάκι. Κάθε μέλος σέρνεται και μορφοποιείται ξεχωριστά.",
                  "Its members (each country or digit) appear under it in the list - open it with the arrow. Each member is dragged and formatted on its own.")),
        ('step', ("**Ίδια μορφή σε όλα τα αδέλφια** αντιγράφει γραμματοσειρά, μέγεθος, χρώμα και στοίχιση ενός μέλους στα υπόλοιπα. **Συγχώνευση ξανά σε ένα** τα ενώνει πάλι.",
                  "**Same format for all siblings** copies one member's font, size, colour and alignment to the others. **Merge back into one** joins them again.")),
        ('tip', ("Στις μεμονωμένες χώρες το X είναι ελεύθερο, ενώ το Y είναι μετατόπιση από τη γραμμή της χώρας στο template - έτσι ακολουθούν αυτόματα το ύψος των γραμμών.",
                 "For individual countries X is free while Y is an offset from the country's row in the template - so they automatically follow the row heights.")),
        ('h', ("Προσθήκη και αναζήτηση", "Adding and searching")),
        ('li', ("Το πεδίο **Αναζήτηση** (`Ctrl+F`) φιλτράρει τη λίστα στοιχείων ανά όνομα ή ομάδα.", "The **Search** box (`Ctrl+F`) filters the item list by name or group.")),
        ('li', ("**+ Προσθήκη στοιχείου…** ανοίγει έναν κατάλογο με ό,τι μπορεί να μπει στο δελτίο: στοιχεία ξενοδοχείου, σύνολα, ημερομηνίες, όνομα μήνα και πελάτες ή διανυκτερεύσεις συγκεκριμένης χώρας. Διπλό κλικ προσθέτει, και μπορείτε να το προσθέσετε πολλές φορές.",
                "**+ Add item…** opens a catalog of everything that can go on the form: hotel details, totals, dates, the month name and the guests or nights of a specific country. Double-click adds it, and you can add it many times.")),
        ('li', ("**+ Δικό μου κείμενο** βάζει ελεύθερο κείμενο που γράφετε εσείς.", "**+ My own text** places free text that you type.")),
        ('li', ("**Διπλασιασμός** (`Ctrl+D`) φτιάχνει αντίγραφο του επιλεγμένου στοιχείου και **Διαγραφή** (`Delete`) αφαιρεί όσα προσθέσατε εσείς.",
                "**Duplicate** (`Ctrl+D`) makes a copy of the selected item and **Delete** (`Delete`) removes the ones you added.")),
        ('h', ("Έλεγχος και αποθήκευση", "Checking and saving")),
        ('li', ("**Δείγμα δεδομένων**: δείχνει πλασματικές τιμές πάνω στη σελίδα. Αν το ξετσεκάρετε βλέπετε μόνο τα ονόματα των στοιχείων.", "**Sample data**: shows made-up values on the page. Untick it to see only the item names.")),
        ('li', ("**Δοκιμαστικό PDF** φτιάχνει ένα πραγματικό PDF με πλασματικά στοιχεία, ώστε να δείτε το τελικό αποτέλεσμα. Μην το υποβάλετε.", "**Test PDF** makes a real PDF with made-up data so you can see the final result. Do not submit it.")),
        ('li', ("**↶ / ↷** (`Ctrl+Z` / `Ctrl+Y`) αναιρούν και επαναλαμβάνουν μέχρι 60 αλλαγές.", "**↶ / ↷** (`Ctrl+Z` / `Ctrl+Y`) undo and redo up to 60 changes.")),
        ('li', ("**Αποθήκευση** (`Ctrl+S`) κρατά τη διάταξη. Αν κλείσετε με μη αποθηκευμένες αλλαγές, το πρόγραμμα ρωτά.", "**Save** (`Ctrl+S`) keeps the layout. If you close with unsaved changes, the program asks.")),
        ('li', ("**Επαναφορά αυτού του στοιχείου** και **Επαναφορά όλων** γυρίζουν στις προεπιλεγμένες θέσεις.", "**Reset this item** and **Reset all** go back to the default positions.")),
        ('warn', ("Η διάταξη αποθηκεύεται στον φάκελο δεδομένων (`elstat_manual_layout.json`) και ισχύει για όλα τα επόμενα δελτία. Κάντε πάντα ένα δοκιμαστικό PDF μετά από αλλαγές.",
                  "The layout is saved in the data folder (`elstat_manual_layout.json`) and applies to every following form. Always make a test PDF after changes.")),
        ('p', ("Πατήστε `F1` μέσα στον επεξεργαστή για να ανοίξει αυτός ο οδηγός.", "Press `F1` inside the editor to open this guide.")),
    ]),
    dict(id='countries', title=("Χώρες και αντιστοιχίσεις", "Countries and mappings"), blocks=[
        ('p', ("Το Protel γράφει τις χώρες με αγγλικά ονόματα. Το πρόγραμμα τις αντιστοιχίζει στις γραμμές του δελτίου της ΕΛΣΤΑΤ με μια ενσωματωμένη λίστα.",
               "Protel writes countries with English names. The program maps them to the lines of the ELSTAT form with a built-in list.")),
        ('h', ("Όταν μια χώρα είναι άγνωστη", "When a country is unknown")),
        ('step', ("Στο Βήμα 1 η άγνωστη χώρα εμφανίζεται με κόκκινο και η στήλη «Στο δελτίο» γράφει ότι λείπει.", "In Step 1 the unknown country shows in red and the “On the form” column says it is missing.")),
        ('step', ("Επιλέξτε την και πατήστε **Προσθήκη / αντιστοίχιση επιλεγμένης χώρας…** (ή κάντε διπλό κλικ πάνω της).", "Select it and press **Add / map selected country…** (or double-click it).")),
        ('step', ("Διαλέξτε σε ποια γραμμή ή κατηγορία του δελτίου ανήκει και πατήστε **Αποθήκευση**.", "Choose which row or category of the form it belongs to and press **Save mapping**.")),
        ('step', ("Πατήστε ξανά **Δημιουργία PDF**. Η αντιστοίχιση θυμάται και για τους επόμενους μήνες.", "Press **Create PDF** again. The mapping is remembered for future months.")),
        ('li', ("Έχετε ήδη **[[custom_count]]** δικές σας αντιστοιχίσεις. Αποθηκεύονται στο `elstat_country_mappings.json` του φακέλου δεδομένων και μπαίνουν στο αντίγραφο ασφαλείας.",
                "You already have **[[custom_count]]** mappings of your own. They are stored in `elstat_country_mappings.json` in the data folder and included in the backup.")),
        ('h', ("Οι κατηγορίες του δελτίου", "The categories of the form")),
        ('kb', ("1", ("Ελλάδα", "Greece"))),
        ('kb', ("2.1", ("Χώρες Ευρωπαϊκής Ένωσης", "European Union countries"))),
        ('kb', ("2.2", ("Χώρες Ε.Ζ.Ε.Σ. (Ελβετία, Λιχτενστάιν, Ισλανδία, Νορβηγία)", "EFTA countries (Switzerland, Liechtenstein, Iceland, Norway)"))),
        ('kb', ("2.3", ("Λοιπές ευρωπαϊκές χώρες", "Other European countries"))),
        ('kb', ("3 - 6", ("Αμερική, Ασία, Αφρική, Ωκεανία", "America, Asia, Africa, Oceania"))),
        ('kb', ("7", ("Δεν δήλωσαν χώρα", "Country not declared"))),
        ('p', ("Μια χώρα χωρίς δική της γραμμή πηγαίνει στη γραμμή «Λοιπές χώρες» της ηπείρου της.", "A country without a line of its own goes to the “Other countries” line of its continent.")),
        ('warn', ("Μετά τη δημιουργία το πρόγραμμα συγκρίνει το γενικό σύνολο με του Excel. Αν διαφέρουν, κάτι δεν μπήκε - ελέγξτε τις χώρες πριν υποβάλετε το δελτίο.",
                  "After creating the form the program compares the grand total with the Excel's. If they differ, something was left out - check the countries before submitting.")),
        ('go', ("Δείτε τα Συχνά προβλήματα", "See Troubleshooting"), 'sec:trouble'),
    ]),
    dict(id='dnd', title=("Μεταφορά αρχείων (drag & drop)", "Drag and drop"), blocks=[
        ('p', ("Μπορείτε να σέρνετε αρχεία και φακέλους από τον Εξερευνητή μέσα στο πρόγραμμα. Ένα αρχείο κάθε φορά.",
               "You can drag files and folders from the file explorer into the program. One file at a time.")),
        ('li', ("**Excel** (.xls / .xlsx) → το πλαίσιο ή η κάρτα του Βήματος 1.",
                "**Excel** (.xls / .xlsx) → the box or card in Step 1.")),
        ('li', ("**PDF** → η κάρτα του template (Στοιχεία ξενοδοχείου) ή η κάρτα του Βήματος 3 (ορίζει αυτό το υπάρχον αρχείο ως έξοδο).",
                "**PDF** → the template card (Hotel details) or the Step 3 card (sets that existing file as the output).")),
        ('li', ("**Φάκελος** → τα πεδία φακέλων στις Ρυθμίσεις (φάκελος PDF, φάκελος Excel, φάκελος δεδομένων).",
                "**Folder** → the folder fields in Settings (PDF folder, Excel folder, data folder).")),
        ('warn', ("Αν δεν δουλεύει: χρειάζεται το πακέτο `tkinterdnd2` (εγκατάσταση: `pip install tkinterdnd2`) και το πρόγραμμα δεν πρέπει να τρέχει "
                  "ως Διαχειριστής. Η γραμμή κατάστασης κάτω λέει αν η μεταφορά αρχείων είναι ενεργή. Σε κάθε περίπτωση, μπορείτε να πατήσετε στα κουμπιά επιλογής.",
                  "If it doesn't work: it needs the `tkinterdnd2` package (install with `pip install tkinterdnd2`) and the program must not run "
                  "as Administrator. The status bar at the bottom says whether drag and drop is active. You can always use the choose buttons instead.")),
    ]),
    dict(id='shortcuts', title=("Συντομεύσεις", "Shortcuts"), blocks=[
        ('p', ("Γρήγοροι τρόποι για τα πιο συχνά βήματα.", "Quick ways to do the most common steps.")),
        ('h', ("Παντού", "Everywhere")),
        ('kb', ("Ctrl+O", ("Επιλογή αρχείου Excel", "Choose the Excel file"))),
        ('kb', ("Ctrl+Enter", ("Δημιουργία PDF", "Create PDF"))),
        ('kb', ("Ctrl+P", ("Δημιουργία & Εκτύπωση", "Create & Print"))),
        ('kb', ("F1", ("Άνοιγμα του οδηγού στην ενότητα της τρέχουσας καρτέλας", "Open this guide at the section for the current tab"))),
        ('kb', ("Ctrl+F", ("Αναζήτηση μέσα στον οδηγό", "Search inside the guide"))),
        ('h', ("Στο Ιστορικό", "In History")),
        ('kb', ("Enter", ("Άνοιγμα του επιλεγμένου μήνα", "Open the selected month"))),
        ('kb', ("Delete", ("Διαγραφή του επιλεγμένου μήνα", "Delete the selected month"))),
        ('kb', ("Διπλό κλικ", ("Σε μήνα: άνοιγμα. Σε χώρα: φίλτρο", "On a month: open. On a country: filter"))),
        ('h', ("Στο παράθυρο αποτελέσματος", "In the result window")),
        ('kb', ("Ctrl+P", ("Εκτύπωση", "Print"))),
        ('kb', ("Esc", ("Κλείσιμο", "Close"))),
        ('h', ("Στον επεξεργαστή διάταξης", "In the layout editor")),
        ('kb', ("Ctrl+S", ("Αποθήκευση", "Save"))),
        ('kb', ("Ctrl+Z / Ctrl+Y", ("Αναίρεση / επανάληψη", "Undo / redo"))),
        ('kb', ("Ctrl+D", ("Διπλασιασμός στοιχείου", "Duplicate item"))),
        ('kb', ("Ctrl+F", ("Αναζήτηση στη λίστα στοιχείων", "Search the item list"))),
        ('kb', ("Βελάκια", ("Μετακίνηση 0,5 pt (Shift = 5 pt)", "Move 0.5 pt (Shift = 5 pt)"))),
        ('kb', ("Ctrl + τροχός", ("Μεγέθυνση / σμίκρυνση", "Zoom in / out"))),
        ('kb', ("Delete", ("Διαγραφή στοιχείου που προσθέσατε", "Delete an item you added"))),
        ('kb', ("Esc", ("Αποεπιλογή", "Deselect"))),
        ('kb', ("F1", ("Άνοιγμα αυτού του οδηγού", "Open this guide"))),
    ]),
    dict(id='data', title=("Δεδομένα και αντίγραφα ασφαλείας", "Data and backups"), blocks=[
        ('p', ("Όλα τα δεδομένα του προγράμματος βρίσκονται σε έναν φάκελο (δείτε Ρυθμίσεις → Φάκελος δεδομένων):",
               "All of the program's data lives in one folder (see Settings → Data folder):")),
        ('li', ("Τρέχουσα θέση: `[[data_dir]]`", "Current location: `[[data_dir]]`")),
        ('li', ("`hotel_settings.csv` - τα στοιχεία του ξενοδοχείου.", "`hotel_settings.csv` - the hotel details.")),
        ('li', ("`elstat_history.csv` και `elstat_months.csv` - το ιστορικό.", "`elstat_history.csv` and `elstat_months.csv` - the history.")),
        ('li', ("`elstat_preferences.json` - οι ρυθμίσεις.", "`elstat_preferences.json` - the settings.")),
        ('li', ("`elstat_country_mappings.json` - οι δικές σας αντιστοιχίσεις χωρών.", "`elstat_country_mappings.json` - your own country mappings.")),
        ('li', ("`elstat_manual_layout.json` και `elstat_template_layout.json` - οι διατάξεις του επεξεργαστή και της έξυπνης διάταξης.",
                "`elstat_manual_layout.json` and `elstat_template_layout.json` - the layouts of the editor and of smart layout.")),
        ('li', ("`ELSTAT_Archive` - τα αρχειοθετημένα PDF κάθε μήνα.", "`ELSTAT_Archive` - the archived PDF of every month.")),
        ('p', ("Δίπλα στο πρόγραμμα υπάρχει και το `elstat_config.json`, που θυμάται μόνο **πού** είναι ο φάκελος δεδομένων.",
               "Next to the program there is also `elstat_config.json`, which only remembers **where** the data folder is.")),
        ('h', ("Αντίγραφα ασφαλείας", "Backups")),
        ('li', ("Φτιάξτε αντίγραφο από Ρυθμίσεις → **Αντίγραφο ασφαλείας (.zip)…**, π.χ. μία φορά τον μήνα, και κρατήστε το σε άλλο δίσκο ή στο σύννεφο.",
                "Make a copy from Settings → **Backup (.zip)…**, e.g. once a month, and keep it on another drive or in the cloud.")),
        ('li', ("Για επαναφορά, αποσυμπιέστε το .zip σε έναν φάκελο και διαλέξτε τον ως **Φάκελο δεδομένων**. Το πρόγραμμα θα ρωτήσει αν θέλετε να χρησιμοποιήσει τα δεδομένα που υπάρχουν εκεί.",
                "To restore, unzip the .zip into a folder and choose it as the **Data folder**. The program will ask whether to use the data that is already there.")),
        ('tip', ("Αν ο φάκελος δεδομένων δεν είναι διαθέσιμος (π.χ. αποσυνδεδεμένος δίσκος), το πρόγραμμα χρησιμοποιεί προσωρινά τον προεπιλεγμένο και σας το λέει.",
                 "If the data folder is not available (e.g. a disconnected drive), the program temporarily uses the default one and tells you.")),
        ('go', ("Άνοιγμα των Ρυθμίσεων", "Open Settings"), 'tab:settings'),
    ]),
    dict(id='faq', title=("Συχνές ερωτήσεις", "FAQ"), blocks=[
        ('h', ("Πού είναι αποθηκευμένα τα δεδομένα μου;", "Where is my data stored?")),
        ('p', ("Στον φάκελο δεδομένων: `[[data_dir]]`. Τον αλλάζετε από τις Ρυθμίσεις.", "In the data folder: `[[data_dir]]`. You change it in Settings.")),
        ('h', ("Μπορώ να το χρησιμοποιήσω σε δύο υπολογιστές;", "Can I use it on two computers?")),
        ('p', ("Ναι: βάλτε τον φάκελο δεδομένων σε OneDrive ή άλλη υπηρεσία συγχρονισμού και διαλέξτε τον και στους δύο υπολογιστές. Μη δουλεύετε ταυτόχρονα και στους δύο.",
               "Yes: put the data folder in OneDrive or another sync service and choose it on both computers. Do not work on both at the same time.")),
        ('h', ("Έκανα λάθος στον αριθμό εργαζομένων. Τι κάνω;", "I typed the wrong number of employees. What now?")),
        ('p', ("Ξανά-δημιουργήστε το δελτίο του μήνα με το σωστό αριθμό (το παλιό PDF αντικαθίσταται μετά από ερώτηση) ή διαγράψτε τον μήνα από το Ιστορικό και φτιάξτε τον από την αρχή.",
               "Create the month's form again with the right number (the old PDF is replaced after a question) or delete the month from History and make it again.")),
        ('h', ("Η ΕΛΣΤΑΤ άλλαξε το έντυπο. Τι κάνω;", "ELSTAT changed the form. What do I do?")),
        ('p', ("Διαλέξτε το νέο PDF στην καρτέλα Στοιχεία ξενοδοχείου. Αν οι τιμές πέφτουν εκτός θέσης, δοκιμάστε την **Έξυπνη διάταξη** ή τοποθετήστε τα μόνοι σας με τον **επεξεργαστή διάταξης**. Κάντε πάντα δοκιμαστικό δελτίο πριν υποβάλετε.",
               "Choose the new PDF in the Hotel details tab. If values land out of place, try **Smart layout** or place them yourself with the **layout editor**. Always make a test form before submitting.")),
        ('go', ("Άνοιγμα του επεξεργαστή διάταξης", "See the layout editor"), 'sec:editor'),
        ('h', ("Μπορώ να ξανατυπώσω παλιό μήνα;", "Can I reprint an old month?")),
        ('p', ("Ναι, από το Ιστορικό: επιλέξτε τον μήνα και πατήστε Εκτύπωση. Αν λείπει το αρχειοθετημένο PDF, το πρόγραμμα ξαναφτιάχνει το δελτίο από το ιστορικό.",
               "Yes, from History: select the month and press Print. If the archived PDF is missing, the program rebuilds the form from the history.")),
        ('h', ("Το PDF γράφεται στα αγγλικά αν αλλάξω γλώσσα;", "Is the PDF written in English if I change the language?")),
        ('p', ("Όχι. Η γλώσσα αφορά μόνο το πρόγραμμα. Το επίσημο δελτίο συμπληρώνεται πάντα στα ελληνικά.", "No. The language only affects the program. The official form is always filled in Greek.")),
        ('h', ("Πόσους μήνες έχω ήδη καταχωρημένους;", "How many months do I already have recorded?")),
        ('p', ("**[[months_count]]** μήνες. Τελευταίος: [[last_month]].", "**[[months_count]]** months. Latest: [[last_month]].")),
    ]),
    dict(id='glossary', title=("Γλωσσάρι", "Glossary"), blocks=[
        ('p', ("Οι όροι που θα συναντήσετε στο πρόγραμμα.", "The terms you will meet in the program.")),
        ('kb', ("Protel", ("Το πρόγραμμα του ξενοδοχείου που βγάζει το Excel «Nationalities Statistics».", "The hotel program that produces the “Nationalities Statistics” Excel."))),
        ('kb', ("ΕΛΣΤΑΤ", ("Η Ελληνική Στατιστική Αρχή, στην οποία υποβάλλεται το δελτίο.", "The Hellenic Statistical Authority, to which the form is submitted."))),
        ('kb', ("Template", ("Το κενό δελτίο PDF πάνω στο οποίο γράφονται οι τιμές.", "The blank PDF form on which the values are written."))),
        ('kb', ("Αρχείο / Archive", ("Αντίγραφο κάθε δελτίου στον φάκελο δεδομένων, για να ξανανοίγεται από το Ιστορικό.", "A copy of every form in the data folder, so it can be reopened from History."))),
        ('kb', ("Αναφορά", ("Το template όπως ήταν όταν το δελτίο συμπληρωνόταν σωστά - η έξυπνη διάταξη μετρά τις αλλαγές σε σχέση με αυτό.", "The template as it was when the form came out right - smart layout measures changes against it."))),
        ('kb', ("pt", ("Σημείο PDF, 1/72 της ίντσας. Οι θέσεις του επεξεργαστή μετριούνται σε pt.", "A PDF point, 1/72 of an inch. The editor's positions are measured in pt."))),
        ('kb', ("Α.Φ.Μ.", ("Ο 9ψήφιος Αριθμός Φορολογικού Μητρώου.", "The 9-digit tax number (AFM)."))),
        ('kb', ("Ε.Ζ.Ε.Σ.", ("Ευρωπαϊκή Ζώνη Ελεύθερων Συναλλαγών (EFTA).", "European Free Trade Association (EFTA)."))),
        ('kb', ("Μεμονωμένο", ("Στοιχείο που προέκυψε από τον χωρισμό μιας ομάδας (χώρα ή ψηφίο) και κινείται μόνο του.", "An item that came from splitting a group (a country or digit) and moves on its own."))),
    ]),
    dict(id='trouble', title=("Συχνά προβλήματα", "Troubleshooting"), blocks=[
        ('h', ("Δεν μπορώ να γράψω το PDF", "I can't write the PDF")),
        ('p', ("Το αρχείο είναι πιθανότατα ανοιχτό σε άλλο πρόγραμμα. Κλείστε το και δοκιμάστε ξανά.",
               "The file is most likely open in another program. Close it and try again.")),
        ('h', ("Το κουμπί «Δημιουργία PDF» είναι γκρι", "The “Create PDF” button is grey")),
        ('p', ("Διαβάστε το μήνυμα ελέγχου ακριβώς από πάνω του: λέει τι πρέπει να διορθώσετε (π.χ. δεν έχετε διαλέξει Excel, ή λείπει ο Α.Φ.Μ.).",
               "Read the check message right above it: it says what to fix (e.g. no Excel chosen, or the Tax ID is missing).")),
        ('h', ("Μια χώρα δεν μπήκε στο PDF", "A country was left out of the PDF")),
        ('p', ("Η χώρα δεν είναι γνωστή στο πρόγραμμα. Εμφανίζεται με κόκκινο στον πίνακα του Βήματος 1 και στο παράθυρο αποτελέσματος, "
               "μαζί με τους πελάτες και τις διανυκτερεύσεις της. Επιλέξτε την κόκκινη χώρα στον πίνακα και πατήστε «Προσθήκη / αντιστοίχιση χώρας» για να τη συνδέσετε με τη σωστή γραμμή της ΕΛΣΤΑΤ. Η επιλογή αποθηκεύεται για τους επόμενους μήνες. Ελέγξτε τα σύνολα πριν υποβάλετε το δελτίο.",
               "The country is unknown to the program. It shows in red in the Step 1 table and in the result window, with its guests and nights. "
               "Select the red country in the table and click “Add / map country” to assign it to the correct ELSTAT row. The choice is saved for future months. Check the totals before submitting.")),
        ('h', ("Τα σύνολα διαφέρουν", "The totals differ")),
        ('p', ("Το γενικό σύνολο του PDF δεν ταιριάζει με το Excel - συνήθως επειδή κάτι δεν μπήκε (δείτε το προηγούμενο). Μην υποβάλετε το δελτίο πριν το ελέγξετε.",
               "The PDF's grand total does not match the Excel - usually because something was left out (see the previous point). Don't submit the form before checking it.")),
        ('h', ("Λάθος μήνας ή έτος", "Wrong month or year")),
        ('p', ("Αν το πρόγραμμα ρωτήσει «Διαφορετική περίοδος», το Excel αφορά άλλον μήνα από αυτόν της φόρμας. Διαλέξτε το σωστό Excel ή διορθώστε το Έτος και τον Μήνα στο Βήμα 2.",
               "If the program asks about a “Different period”, the Excel is for another month than the form. Choose the right Excel or correct the Year and Month in Step 2.")),
        ('h', ("Η εκτύπωση δεν ξεκίνησε", "Printing did not start")),
        ('p', ("Σε αυτό τον υπολογιστή δεν ήταν δυνατή η αυτόματη εκτύπωση, οπότε το πρόγραμμα άνοιξε το PDF. Πατήστε `Ctrl+P` μέσα στο πρόγραμμα του PDF.",
               "Automatic printing was not possible on this PC, so the program opened the PDF. Press `Ctrl+P` in the PDF program.")),
        ('h', ("Το drag & drop δεν δουλεύει", "Drag and drop doesn't work")),
        ('go', ("Δείτε την ενότητα Μεταφορά αρχείων", "See the Drag and drop section"), 'sec:dnd'),
    ]),
]
SECTION_BY_ID = {sec['id']: sec for sec in MANUAL}
_INLINE = re.compile(r'(\*\*.+?\*\*|`.+?`)')
_PLACEHOLDER = re.compile(r'\[\[(\w+)\]\]')
HELP_CTX = {}          # live values the guide text can show as [[name]] (filled in by ElstatApp.help_context)


def _tx(pair):
    """The Greek or English half of a (greek, english) pair, with [[name]] replaced by the live value."""
    text = pair[1] if LANG == 'en' else pair[0]
    return _PLACEHOLDER.sub(lambda m: str(HELP_CTX.get(m.group(1), m.group(0))), text)


def _plain(text):
    """Text without the **bold** / `key` markers."""
    return text.replace('**', '').replace('`', '')


def _fold(s):
    """Upper-case without accents, for searching (keeps the length of normal text)."""
    return ''.join(ch for ch in unicodedata.normalize('NFD', str(s).upper()) if unicodedata.category(ch) != 'Mn')


class HelpPanel(ttk.Frame):
    """The built-in user guide: section list with search on the left, the text on the right."""

    def __init__(self, parent, app, start='start'):
        super().__init__(parent)
        self.app = app
        self.current = None
        self.nq = ''
        self.items = {}
        self.callouts = []
        self._link_tags = []
        self._resize_job = None
        self._query_job = None
        self._corpus_cache = {}
        self._ph_on = False
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # ---- left: title, search, section list
        side = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER, width=260)
        side.grid(row=0, column=0, sticky="ns", padx=(12, 0), pady=12)
        side.pack_propagate(False)
        tk.Label(side, text=T("Οδηγός χρήσης", "User guide"), bg=CARD, fg=NAVY,
                 font=(FONT, 14, 'bold'), anchor="w").pack(fill="x", padx=16, pady=(16, 2))
        tk.Label(side, text=T("Όλα όσα χρειάζεστε, σε ενότητες.", "Everything you need, in sections."), bg=CARD, fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(fill="x", padx=16, pady=(0, 10))
        srow = tk.Frame(side, bg=CARD)
        srow.pack(fill="x", padx=14, pady=(0, 8))
        self.q = tk.StringVar()
        self.entry = ttk.Entry(srow, textvariable=self.q)
        self.entry.pack(side="left", fill="x", expand=True)
        self.clear_lbl = tk.Label(srow, text="\u2715", bg=CARD, fg=MUTED, font=(FONT, 10), cursor='hand2', padx=6)
        self.clear_lbl.bind('<Button-1>', lambda e: self.clear_search())
        self._ph = T("Αναζήτηση στον οδηγό…", "Search the guide…")
        self.entry.bind('<FocusIn>', self._ph_hide)
        self.entry.bind('<FocusOut>', self._ph_show)
        self.entry.bind('<Escape>', lambda e: self.clear_search())
        self._ph_show()
        self.listbox = tk.Frame(side, bg=CARD)
        self.listbox.pack(fill="both", expand=True, pady=(0, 10))
        for sec in MANUAL:
            self._make_item(sec)
        # Search text is static until the language is rebuilt; prepare it once
        # instead of normalizing every section on every key press.
        self._corpus_cache = {sec['id']: self._corpus(sec) for sec in MANUAL}

        # ---- right: breadcrumb + text
        main = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        main.grid(row=0, column=1, sticky="nsew", padx=12, pady=12)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        top = tk.Frame(main, bg=CARD)
        top.grid(row=0, column=0, columnspan=2, sticky="ew", padx=34, pady=(14, 0))
        self.crumb = tk.Label(top, text="", bg=CARD, fg=MUTED, font=(FONT, 9), anchor="w")
        self.crumb.pack(side="left")
        self.back_lbl = tk.Label(top, text="", bg=CARD, fg=ACCENT, font=(FONT, 9, 'underline'), cursor='hand2')
        self.back_lbl.pack(side="right")
        self.back_lbl.bind('<Button-1>', lambda e: self._go_back())
        self.text = tk.Text(main, wrap="word", bg=CARD, fg=TEXT, relief="flat", borderwidth=0, highlightthickness=0,
                            padx=34, pady=6, font=(FONT, 11), cursor="arrow", spacing2=4, state="disabled",
                            insertwidth=0, selectbackground=SEL, selectforeground=TEXT, takefocus=0)
        vbar = ttk.Scrollbar(main, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=vbar.set)
        self.text.grid(row=1, column=0, sticky="nsew", pady=(0, 14))
        vbar.grid(row=1, column=1, sticky="ns", pady=(0, 14))
        self.text.bind('<Configure>', lambda e: self._schedule_resize())
        self._configure_tags()

        self.q.trace_add('write', self._on_query)
        self.show(start if start in SECTION_BY_ID else 'start')
        self.refresh_back()

    # ------------------------------------------------------------ sidebar
    def _make_item(self, sec):
        frame = tk.Frame(self.listbox, bg=CARD, cursor='hand2')
        bar = tk.Frame(frame, width=4, bg=CARD)
        bar.pack(side="left", fill="y")
        lbl = tk.Label(frame, text=_tx(sec['title']), bg=CARD, fg=TEXT, font=(FONT, 10), anchor="w",
                       justify="left", padx=12, pady=6, wraplength=190)
        lbl.pack(side="left", fill="x", expand=True)
        badge = tk.Label(frame, text="", bg=CARD, fg=MUTED, font=(FONT, 9), padx=10)
        badge.pack(side="right")
        sid = sec['id']
        for w in (frame, bar, lbl, badge):
            w.bind('<Button-1>', lambda e, i=sid: self.show(i))
            w.bind('<Enter>', lambda e, i=sid: self._hover(i, True))
            w.bind('<Leave>', lambda e, i=sid: self._hover(i, False))
        self.items[sid] = (frame, bar, lbl, badge)
        frame.pack(fill="x")

    def _hover(self, sid, on):
        if sid == self.current:
            return
        frame, bar, lbl, badge = self.items[sid]
        colour = BG if on else CARD
        for w in (frame, bar, lbl, badge):
            w.configure(bg=colour)

    def _paint_items(self):
        for sid, (frame, bar, lbl, badge) in self.items.items():
            sel = sid == self.current
            bg = ACCENT_SOFT if sel else CARD
            frame.configure(bg=bg)
            bar.configure(bg=ACCENT if sel else CARD)
            lbl.configure(bg=bg, fg=ACCENT if sel else TEXT, font=(FONT, 10, 'bold' if sel else 'normal'))
            badge.configure(bg=bg)

    # ------------------------------------------------------------ search
    def _ph_show(self, _e=None):
        if not self.q.get():
            self._ph_on = True
            self.entry.configure(foreground=MUTED)
            self.q.set(self._ph)

    def _ph_hide(self, _e=None):
        if self._ph_on:
            self._ph_on = False
            self.q.set('')
            self.entry.configure(foreground=TEXT)

    def focus_search(self):
        self.entry.focus_set()
        self.entry.select_range(0, 'end')

    def clear_search(self):
        self._ph_on = False
        self.q.set('')
        self.entry.configure(foreground=TEXT)
        self.text.focus_set()

    def _corpus(self, sec):
        parts = [_tx(sec['title'])]
        for blk in sec['blocks']:
            if blk[0] == 'kb':
                # Keyboard-help entries are stored as ('kb', (shortcut, (el, en))).
                # Also accept the older three-field layout for compatibility.
                if len(blk) == 2:
                    shortcut, description = blk[1]
                else:
                    shortcut, description = blk[1], blk[2]
                parts += [shortcut, _tx(description)]
            elif blk[0] not in ('go', 'live'):
                parts.append(_tx(blk[1]))
        return _fold(_plain(' '.join(parts)))

    def _on_query(self, *_):
        # Debounce typing: repainting the guide's Text widget on every keystroke
        # caused visible lag, especially on slower PCs.
        if self._query_job is not None:
            try:
                self.after_cancel(self._query_job)
            except tk.TclError:
                pass
        self._query_job = self.after(110, self._apply_query)

    def _apply_query(self):
        self._query_job = None
        text = '' if self._ph_on else self.q.get().strip()
        self.nq = _fold(text) if text else ''
        if self.nq:
            self.clear_lbl.pack(side="right")
        else:
            self.clear_lbl.pack_forget()
        counts = {sid: (corpus.count(self.nq) if self.nq else 0)
                  for sid, corpus in self._corpus_cache.items()}
        visible = [sec['id'] for sec in MANUAL if not self.nq or counts[sec['id']] > 0]
        for sid, (frame, _bar, _lbl, badge) in self.items.items():
            frame.pack_forget()
            badge.configure(text=str(counts[sid]) if self.nq else "")
        for sid in visible:
            self.items[sid][0].pack(fill="x")
        if not visible:
            self._show_empty()
        elif self.current not in visible:
            # Change section only when the current one is filtered out. Rebuilding
            # the complete Text widget on every search update caused unnecessary lag.
            self.show(visible[0])
        # If the current section is still visible, leave its rendered content alone.

    def _show_empty(self):
        self.current = None
        self._paint_items()
        self._reset_text()
        t = self.text
        t.configure(state='normal')
        t.insert('end', T("Δεν βρέθηκε τίποτα", "Nothing found") + "\n", ('title',))
        t.insert('end', T(f"Δεν υπάρχει αποτέλεσμα για «{self.q.get().strip()}». Δοκιμάστε μια άλλη λέξη ή λιγότερα γράμματα.",
                          f"No result for “{self.q.get().strip()}”. Try another word or fewer letters."), ('p',))
        t.configure(state='disabled')
        self.crumb.configure(text="")

    # ------------------------------------------------------------ text
    def _configure_tags(self):
        t = self.text
        t.tag_configure('title', font=(FONT, 22, 'bold'), foreground=NAVY, spacing1=4, spacing3=14)
        t.tag_configure('h', font=(FONT, 13, 'bold'), foreground=NAVY, spacing1=22, spacing3=6)
        t.tag_configure('p', spacing3=9)
        t.tag_configure('li', lmargin1=8, lmargin2=28, spacing3=6, tabs=('28',))
        t.tag_configure('step', lmargin1=2, lmargin2=34, spacing3=9, tabs=('34',))
        t.tag_configure('kb', lmargin1=2, lmargin2=170, spacing3=8, tabs=('170',))
        t.tag_configure('nav', spacing1=26)
        t.tag_configure('bullet', foreground=ACCENT, font=(FONT, 11, 'bold'))
        t.tag_configure('stepnum', foreground=ACCENT, font=(FONT, 12, 'bold'))
        t.tag_configure('b', font=(FONT, 11, 'bold'))
        t.tag_configure('kbd', font=(FONT, 10, 'bold'), background=BTN, foreground=NAVY)
        t.tag_configure('hit', background=ROW_HL, foreground=TEXT)
        t.tag_configure('mark_ok', foreground=GOOD, font=(FONT, 11, 'bold'))
        t.tag_configure('mark_bad', foreground=BAD, font=(FONT, 11, 'bold'))
        t.tag_configure('mark_warn', foreground=WARN, font=(FONT, 11, 'bold'))

    def _reset_text(self):
        t = self.text
        for fr, _box, _lbl in self.callouts:
            try:
                fr.destroy()
            except tk.TclError:
                pass
        self.callouts = []
        for name in self._link_tags:
            t.tag_delete(name)
        self._link_tags = []
        t.configure(state='normal')
        t.delete('1.0', 'end')
        t.configure(state='disabled')

    def _ins(self, s, tags=()):
        """Inserts text, marking search matches."""
        t = self.text
        q = self.nq
        folded = _fold(s) if q else ''
        if not q or len(folded) != len(s):
            t.insert('end', s, tags)
            return
        i = 0
        while True:
            j = folded.find(q, i)
            if j < 0:
                break
            if j > i:
                t.insert('end', s[i:j], tags)
            t.insert('end', s[j:j + len(q)], tuple(tags) + ('hit',))
            i = j + len(q)
        if i < len(s):
            t.insert('end', s[i:], tags)

    def _inline(self, s, base):
        for part in _INLINE.split(s):
            if not part:
                continue
            if part.startswith('**'):
                self._ins(part[2:-2], base + ('b',))
            elif part.startswith('`'):
                self._ins('\u00a0' + part[1:-1] + '\u00a0', base + ('kbd',))
            else:
                self._ins(part, base)

    def _link(self, label, action, base):
        name = f"link{len(self._link_tags) + 1}"
        t = self.text
        t.tag_configure(name, foreground=ACCENT, underline=True)
        t.tag_bind(name, '<Button-1>', lambda e: action())
        t.tag_bind(name, '<Enter>', lambda e: t.configure(cursor='hand2'))
        t.tag_bind(name, '<Leave>', lambda e: t.configure(cursor='arrow'))
        t.insert('end', label, tuple(base) + (name,))
        self._link_tags.append(name)

    def _callout(self, kind, txt):
        accent, bg = (ACCENT, ACCENT_SOFT) if kind == 'tip' else (WARN, WARN_BG)
        title = T("Συμβουλή", "Tip") if kind == 'tip' else T("Προσοχή", "Important")
        fr = tk.Frame(self.text, bg=bg)
        tk.Frame(fr, bg=accent, width=4).pack(side="left", fill="y")
        box = tk.Frame(fr, bg=bg)
        box.pack(side="left", fill="both", expand=True, padx=14, pady=10)
        tk.Label(box, text=title, bg=bg, fg=accent, font=(FONT, 10, 'bold'), anchor="w").pack(anchor="w")
        lbl = tk.Label(box, text=_plain(txt), bg=bg, fg=TEXT, font=(FONT, 10), justify="left", anchor="w")
        lbl.pack(anchor="w", pady=(2, 0))
        fr.pack_propagate(False)
        fr.configure(width=400, height=80)
        self.callouts.append((fr, box, lbl))
        self.text.window_create('end', window=fr, pady=6)
        self.text.insert('end', '\n')

    def _schedule_resize(self):
        if self._resize_job is not None:
            try:
                self.after_cancel(self._resize_job)
            except tk.TclError:
                pass
        self._resize_job = self.after(60, self._resize_callouts)

    def _resize_callouts(self):
        self._resize_job = None
        try:
            avail = self.text.winfo_width() - 2 * 34 - 8
            if avail < 160:
                return
            # One geometry pass for all embedded callouts, rather than one
            # synchronous idle-task flush per callout during window resizing.
            for _fr, _box, lbl in self.callouts:
                lbl.configure(wraplength=max(120, avail - 52))
            self.update_idletasks()
            for fr, box, _lbl in self.callouts:
                fr.configure(width=avail, height=box.winfo_reqheight() + 20)
        except tk.TclError:
            pass

    def _nav(self, target):
        if target.startswith('tab:'):
            self.app.goto_tab(target[4:])
        else:
            self.show(target[4:])

    def _live(self, what):
        """Blocks whose content is worked out from the state of the program every time they are shown."""
        t = self.text
        if what == 'next':
            label, target = self.app.help_next()
            t.insert('end', T("Επόμενο βήμα:", "Next step:") + ' ', ('p', 'b'))
            self._link(label + '  →', lambda tg=target: self._nav(tg), ('p',))
            t.insert('end', '\n', ('p',))
        elif what == 'status':
            for state, label, detail, target in self.app.help_status():
                t.insert('end', {'ok': '✓', 'bad': '✗', 'warn': '⚠'}[state] + '\t', ('li', 'mark_' + state))
                self._link(label, lambda tg=target: self._nav(tg), ('li', 'b'))
                t.insert('end', ' – ' + detail + '\n', ('li',))

    def refresh_live(self):
        """The guide tab was opened: show the current section again with fresh live values."""
        if self.current:
            y = self.text.yview()[0]
            self.show(self.current)
            self.text.yview_moveto(y)

    def show(self, sid):
        sec = SECTION_BY_ID.get(sid)
        if sec is None:
            return
        try:
            HELP_CTX.update(self.app.help_context())
        except Exception:
            pass
        self.current = sid
        self.app._help_last = sid
        self._paint_items()
        self._reset_text()
        t = self.text
        t.configure(state='normal')
        self._ins(_tx(sec['title']), ('title',))
        t.insert('end', '\n', ('title',))
        step_no = 0
        prev = None
        for blk in sec['blocks']:
            kind = blk[0]
            step_no = step_no + 1 if kind == 'step' else 0
            if kind == 'p':
                self._inline(_tx(blk[1]), ('p',))
                t.insert('end', '\n', ('p',))
            elif kind == 'h':
                self._ins(_tx(blk[1]), ('h',))
                t.insert('end', '\n', ('h',))
            elif kind == 'li':
                t.insert('end', '\u2022\t', ('li', 'bullet'))
                self._inline(_tx(blk[1]), ('li',))
                t.insert('end', '\n', ('li',))
            elif kind == 'step':
                t.insert('end', f"{step_no}\t", ('step', 'stepnum'))
                self._inline(_tx(blk[1]), ('step',))
                t.insert('end', '\n', ('step',))
            elif kind == 'kb':
                if len(blk) == 2:
                    shortcut, description = blk[1]
                else:
                    shortcut, description = blk[1], blk[2]
                self._ins('\u00a0' + shortcut + '\u00a0', ('kb', 'kbd'))
                t.insert('end', '\t', ('kb',))
                self._inline(_tx(description), ('kb',))
                t.insert('end', '\n', ('kb',))
            elif kind in ('tip', 'warn'):
                self._callout(kind, _tx(blk[1]))
            elif kind == 'go':
                target = blk[2]
                if target.startswith('tab:'):
                    action = lambda n=target[4:]: self.app.goto_tab(n)
                else:
                    action = lambda i=target[4:]: self.show(i)
                self._link(_tx(blk[1]) + '  →', action, ('p',))
                t.insert('end', '\n', ('p',))
            elif kind == 'live':
                self._live(blk[1])
            elif kind == 'act':
                self._link('▸  ' + _tx(blk[1]), lambda k=blk[2]: self.app.help_action(k), ('p',))
                t.insert('end', '\n', ('p',))
            prev = kind
        # previous / next section
        order = [s['id'] for s in MANUAL]
        i = order.index(sid)
        t.insert('end', '\n', ('nav',))
        if i > 0:
            self._link('\u2039  ' + _tx(MANUAL[i - 1]['title']), lambda s=order[i - 1]: self.show(s), ('nav',))
        if i > 0 and i < len(order) - 1:
            t.insert('end', '          ', ('nav',))
        if i < len(order) - 1:
            self._link(_tx(MANUAL[i + 1]['title']) + '  \u203a', lambda s=order[i + 1]: self.show(s), ('nav',))
        t.tag_raise('hit')
        t.configure(state='disabled')
        t.yview_moveto(0)
        self.crumb.configure(text=T(f"Ενότητα {i + 1} από {len(order)}", f"Section {i + 1} of {len(order)}"))
        self._schedule_resize()

    # ------------------------------------------------------------ back link
    def refresh_back(self):
        prev = getattr(self.app, '_help_prev', None)
        try:
            if prev and str(prev) in self.app.notebook.tabs() and str(prev) != str(self.app.tab_help):
                name = self.app.notebook.tab(prev, 'text')
                self.back_lbl.configure(text=T(f"\u2190  Πίσω στο «{name}»", f"\u2190  Back to “{name}”"))
                return
        except tk.TclError:
            pass
        self.back_lbl.configure(text="")

    def _go_back(self):
        prev = getattr(self.app, '_help_prev', None)
        try:
            if prev and str(prev) in self.app.notebook.tabs():
                self.app.notebook.select(prev)
        except tk.TclError:
            pass


class PDFPreviewDialog(tk.Toplevel):
    """In-app preview of the generated PDF; Continue opens the normal Result dialog."""

    def __init__(self, parent, pdf_path, on_continue):
        super().__init__(parent)
        self.title(T("Προεπισκόπηση PDF", "PDF preview"))
        self.configure(bg=BG)
        self.transient(parent)
        self.minsize(720, 600)
        self._on_continue = on_continue
        self._continued = False
        self._images = []

        header = ttk.Frame(self, padding=(16, 12, 16, 8))
        header.pack(fill="x")
        ttk.Label(
            header,
            text=T("Ελέγξτε το δελτίο πριν συνεχίσετε",
                   "Review the form before continuing"),
            style='CardBold.TLabel'
        ).pack(anchor="w")
        ttk.Label(
            header,
            text=T("Αυτή είναι η προεπισκόπηση του PDF που δημιουργήθηκε. "
                   "Όταν τελειώσετε τον έλεγχο, πατήστε «Συνέχεια στο αποτέλεσμα».",
                   "This is a preview of the generated PDF. When you finish reviewing it, "
                   "press “Continue to result”."),
            style='Hint.TLabel',
            wraplength=900,
            justify="left"
        ).pack(anchor="w", pady=(4, 0))

        viewer = ttk.Frame(self)
        viewer.pack(fill="both", expand=True, padx=14, pady=(4, 10))
        self.canvas = tk.Canvas(viewer, bg="#596579", highlightthickness=0)
        scrollbar = ttk.Scrollbar(viewer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        content = tk.Frame(self.canvas, bg="#596579")
        self._content_window = self.canvas.create_window((0, 0), window=content, anchor="n")
        content.bind("<Configure>", lambda _e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._resize_content)

        rendered = 0
        try:
            reader = pypdf.PdfReader(pdf_path)
            page_count = len(reader.pages)
        except Exception:
            page_count = 0

        for page_index in range(page_count):
            ttk.Label(
                content,
                text=T(f"Σελίδα {page_index + 1}", f"Page {page_index + 1}")
            ).pack(pady=(14, 5))
            png = render_pdf_page_png(pdf_path, page_index, 0.85)
            if png:
                try:
                    photo = tk.PhotoImage(data=base64.b64encode(png).decode("ascii"))
                    self._images.append(photo)
                    image_label = tk.Label(content, image=photo, bg="white",
                                            bd=1, relief="solid")
                    image_label.pack(padx=12, pady=(0, 14))
                    rendered += 1
                except tk.TclError:
                    pass

        if not rendered:
            ttk.Label(
                content,
                text=T("Δεν ήταν δυνατή η εμφάνιση προεπισκόπησης σε αυτόν τον υπολογιστή. "
                       "Μπορείτε να ανοίξετε το PDF εξωτερικά και μετά να συνεχίσετε.",
                       "A preview could not be rendered on this computer. "
                       "You can open the PDF externally and then continue."),
                style='CardHint.TLabel',
                wraplength=650,
                justify="center"
            ).pack(padx=30, pady=50)
            ttk.Button(
                content,
                text=T("Άνοιγμα PDF", "Open PDF"),
                command=lambda: open_file(pdf_path)
            ).pack(pady=(0, 24))

        footer = ttk.Frame(self, padding=(14, 8, 14, 14))
        footer.pack(fill="x")
        ttk.Button(
            footer,
            text=T("Άνοιγμα PDF", "Open PDF"),
            command=lambda: open_file(pdf_path)
        ).pack(side="left")
        ttk.Button(
            footer,
            text=T("Συνέχεια στο αποτέλεσμα  →", "Continue to result  →"),
            style='Primary.TButton',
            command=self._continue
        ).pack(side="right")

        self.bind("<Escape>", lambda _e: self._continue())
        self.protocol("WM_DELETE_WINDOW", self._continue)
        self.update_idletasks()
        try:
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            w, h = min(1100, sw - 50), min(900, sh - 70)
            x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
            self.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
            self.grab_set()
        except tk.TclError:
            self.geometry("1000x800")

    def _resize_content(self, event):
        self.canvas.coords(self._content_window, event.width // 2, 0)

    def _continue(self):
        if self._continued:
            return
        self._continued = True
        callback = self._on_continue
        self.destroy()
        if callback:
            callback()


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
        tree.tag_configure('bold', font=(FONT, 10, 'bold'), background=ROW_BOLD)
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


class Segmented(tk.Frame):
    """Small pill-shaped switch with a few options, e.g. [ Page 1 | Page 2 ]."""

    def __init__(self, parent, options, command=None, value=None):
        super().__init__(parent, bg=BORDER)
        self._command = command
        self._labels = {}
        self.value = value
        for i, (val, text) in enumerate(options):
            lbl = tk.Label(self, text=text, font=(FONT, 9, 'bold'), padx=12, pady=5, cursor='hand2')
            lbl.pack(side="left", padx=(1, 1 if i == len(options) - 1 else 0), pady=1)
            lbl.bind('<Button-1>', lambda e, v=val: self.set(v, True))
            self._labels[val] = lbl
        self.set(value, False)

    def set(self, value, notify=False):
        self.value = value
        for val, lbl in self._labels.items():
            on = val == value
            lbl.configure(bg=ACCENT if on else BTN, fg='white' if on else TEXT)
        if notify and self._command:
            self._command(value)


class LayoutEditor(tk.Toplevel):
    """Built-in layout editor of the "Set manually" mode: the blank form is shown as a page, every value
    is a piece of text on it that can be dragged, and font, size, colour... are set on the right."""
    PAD = 28
    NUDGE_STEPS = ("0.1", "0.5", "1", "2", "5")
    QUICK_COLORS = ('#000000', '#1f2937', '#1d4ed8', '#047857', '#b91c1c', '#7c3aed')

    def __init__(self, app, template_path, on_save=None):
        super().__init__(app)
        self.app = app
        self.template = template_path
        self.on_save = on_save
        self.layout = load_manual_layout()
        self.page = 0
        self.zoom = 1.0
        self.ox = self.PAD
        self.sel = None
        self.dirty = False
        self._loading = False
        self._sync = False
        self._drag = None
        self._photos = {}
        self._fit_pending = True
        self._redraw_job = None
        self._last_width = 0
        self._fallback_texts = None
        self.undo, self.redo = [], []
        self._snap_key, self._snap_time, self._stagger = None, 0.0, 0
        self.q_var = tk.StringVar()
        self.tip_var = tk.StringVar()
        self.sizes = [(595.0, 842.0), (595.0, 842.0)]
        self.row_ys = [{}, {}]
        self._read_template()
        self.sample = manual_sample_values()
        self.families = manual_font_families()
        self.show_sample = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar()

        self.configure(bg=BG)
        self.minsize(1060, 680)
        self._build()
        self._fill_tree()
        self._load_props()
        self._update_title()
        self.protocol('WM_DELETE_WINDOW', self._close)
        for seq in ('<Left>', '<Right>', '<Up>', '<Down>', '<Shift-Left>', '<Shift-Right>', '<Shift-Up>', '<Shift-Down>'):
            self.bind(seq, self._on_arrow)
        self.bind('<Control-s>', lambda e: self.save())
        self.bind('<Control-S>', lambda e: self.save())
        self.bind('<Delete>', self._on_delete_key)
        self.bind('<Control-z>', lambda e: None if self._typing() else self._undo())
        self.bind('<Control-y>', lambda e: None if self._typing() else self._redo())
        self.bind('<Control-d>', lambda e: None if self._typing() else self._duplicate())
        self.bind('<F1>', lambda e: self._open_guide())
        self.bind('<Control-f>', lambda e: (self.search_entry.focus_set(), self.search_entry.select_range(0, 'end'), 'break')[2])
        self.bind('<Escape>', lambda e: self._select(None))
        self.transient(app)
        try:
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            w, h = min(1340, sw - 60), min(880, sh - 90)
            x = app.winfo_rootx() + (app.winfo_width() - w) // 2
            y = app.winfo_rooty() + (app.winfo_height() - h) // 2
            self.geometry(f"{w}x{h}+{max(0, min(x, sw - w))}+{max(0, min(y, sh - h - 40))}")
        except tk.TclError:
            self.geometry("1280x820")
        self.after(150, self._try_grab)

    # ------------------------------------------------------------------ setup
    def _try_grab(self):
        try:
            self.grab_set()
            self.focus_set()
        except tk.TclError:
            pass

    def _read_template(self):
        try:
            reader = pypdf.PdfReader(self.template)
            for i in range(2):
                box = reader.pages[i].mediabox
                self.sizes[i] = (float(box.width), float(box.height))
        except Exception:
            pass
        try:
            self.row_ys = detect_template_rows(self.template)
        except Exception:
            self.row_ys = [{}, {}]

    def _card(self, parent):
        outer = tk.Frame(parent, bg=BORDER)
        inner = ttk.Frame(outer, style='Card.TFrame')
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        return outer, inner

    def _build(self):
        st = ttk.Style(self)
        st.configure('Editor.Treeview', rowheight=26)
        st.configure('TSpinbox', fieldbackground=FIELD, foreground=TEXT, insertcolor=TEXT, bordercolor=BORDER,
                     lightcolor=BORDER, darkcolor=BORDER, arrowcolor=TEXT, background=BTN, padding=5)
        st.map('TSpinbox', bordercolor=[('focus', ACCENT)], lightcolor=[('focus', ACCENT)],
               darkcolor=[('focus', ACCENT)])
        st.configure('Horizontal.TScrollbar', background=SCROLL, troughcolor=BG, bordercolor=BG,
                     arrowcolor=MUTED, relief='flat')
        # ---- header strip, like the main window
        head = tk.Frame(self, bg=HEADER_BG)
        head.pack(fill="x")
        tk.Label(head, text=T("Επεξεργαστής διάταξης", "Layout editor"), bg=HEADER_BG, fg=HEADER_FG,
                 font=(FONT, 14, 'bold')).pack(side="left", padx=(18, 12), pady=12)
        tk.Label(head, text=T("Σύρετε κάθε στοιχείο εκεί που ανήκει στο έντυπο · βελάκια = λεπτομερής μετακίνηση",
                              "Drag each item to where it belongs on the form · arrow keys = fine movement"),
                 bg=HEADER_BG, fg=HEADER_MUTED, font=(FONT, 10)).pack(side="left")

        # ---- toolbar
        bar = ttk.Frame(self, padding=(14, 10, 14, 6))
        bar.pack(fill="x")
        self.page_seg = Segmented(bar, ((0, T("Σελίδα 1", "Page 1")), (1, T("Σελίδα 2", "Page 2"))),
                                  self._on_page_seg, 0)
        self.page_seg.pack(side="left")
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=14)
        ttk.Button(bar, text="−", width=3, command=lambda: self._set_zoom(self.zoom / 1.2)).pack(side="left")
        self.zoom_lbl = ttk.Label(bar, text="100%", width=6, anchor="center")
        self.zoom_lbl.pack(side="left")
        ttk.Button(bar, text="+", width=3, command=lambda: self._set_zoom(self.zoom * 1.2)).pack(side="left")
        ttk.Button(bar, text=T("Προσαρμογή", "Fit"), command=self._fit).pack(side="left", padx=(8, 0))
        self.undo_btn = ttk.Button(bar, text="\u21b6", width=3, command=self._undo)
        self.undo_btn.pack(side="left", padx=(14, 2))
        self.redo_btn = ttk.Button(bar, text="\u21b7", width=3, command=self._redo)
        self.redo_btn.pack(side="left")
        self._update_undo_buttons()
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=14)
        ttk.Checkbutton(bar, text=T("Δείγμα δεδομένων", "Sample data"), variable=self.show_sample,
                        command=self._draw_items).pack(side="left")
        ttk.Button(bar, text=T("Αποθήκευση", "Save"), style='PrintSm.TButton', command=self.save).pack(side="right")
        ttk.Button(bar, text=T("Δοκιμαστικό PDF", "Test PDF"), command=self._test_pdf).pack(side="right", padx=8)
        ttk.Button(bar, text=T("Επαναφορά όλων", "Reset all"), style='Danger.TButton',
                   command=self._reset_all).pack(side="right")

        # ---- status line (bottom, packed before the body so it never disappears)
        status = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        status.pack(side="bottom", fill="x")
        ttk.Label(status, textvariable=self.status_var, style='Status.TLabel').pack(side="left")
        self.render_note = ttk.Label(status, text="", style='Status.TLabel')
        self.render_note.pack(side="right")

        tipbar = tk.Frame(self, bg=ACCENT_SOFT, highlightthickness=1, highlightbackground=BORDER)
        tipbar.pack(side="bottom", fill="x")
        more = tk.Label(tipbar, text=T("Πλήρης οδηγός (F1) \u203a", "Full guide (F1) \u203a"), bg=ACCENT_SOFT, fg=ACCENT,
                        font=(FONT, 9, 'underline'), cursor='hand2', padx=12)
        more.pack(side="right")
        more.bind('<Button-1>', lambda e: self._open_guide())
        self.tip_lbl = tk.Label(tipbar, textvariable=self.tip_var, bg=ACCENT_SOFT, fg=TEXT, font=(FONT, 9),
                                anchor="w", justify="left", padx=12, pady=6)
        self.tip_lbl.pack(side="left", fill="x", expand=True)
        tipbar.bind('<Configure>', lambda e: self.tip_lbl.configure(wraplength=max(300, e.width - 190)))

        # ---- body: field list | page | properties
        body = ttk.PanedWindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=14, pady=(4, 10))

        left_outer, left = self._card(body)
        ttk.Label(left, text=T("Στοιχεία του δελτίου", "Items on the form"), style='CardBold.TLabel').pack(
            anchor="w", padx=14, pady=(12, 4))
        srow = ttk.Frame(left, style='Card.TFrame')
        srow.pack(fill="x", padx=14, pady=(0, 6))
        ttk.Label(srow, text=T("Αναζήτηση:", "Search:"), style='CardHint.TLabel').pack(side="left", padx=(0, 6))
        self.search_entry = ttk.Entry(srow, textvariable=self.q_var)
        self.search_entry.pack(side="left", fill="x", expand=True)
        clear = tk.Label(srow, text="\u2715", bg=CARD, fg=MUTED, font=(FONT, 10), cursor='hand2', padx=6)
        clear.pack(side="left")
        clear.bind('<Button-1>', lambda e: self.q_var.set(''))
        self.q_var.trace_add('write', lambda *a: self._fill_tree())
        tree_box = ttk.Frame(left, style='Card.TFrame')
        tree_box.pack(fill="both", expand=True, padx=8)
        self.tree = ttk.Treeview(tree_box, show='tree', selectmode='browse', height=18, style='Editor.Treeview')
        self.tree.column('#0', width=250, minwidth=180, stretch=True)
        sb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure('group', font=(FONT, 10, 'bold'), foreground=NAVY)
        self.tree.tag_configure('hidden', foreground=MUTED)
        self.tree.bind('<<TreeviewSelect>>', self._on_tree_select)
        lrow = ttk.Frame(left, style='Card.TFrame')
        lrow.pack(fill="x", padx=14, pady=(12, 4))
        ttk.Button(lrow, text=T("+ Προσθήκη στοιχείου…", "+ Add item…"), style='PrintSm.TButton',
                   command=self._add_item_dialog).pack(side="left")
        ttk.Button(lrow, text=T("+ Δικό μου κείμενο", "+ My own text"), command=self._add_custom).pack(side="left", padx=8)
        lrow2 = ttk.Frame(left, style='Card.TFrame')
        lrow2.pack(fill="x", padx=14, pady=(0, 12))
        self.dup_btn = ttk.Button(lrow2, text=T("Διπλασιασμός", "Duplicate"), command=self._duplicate)
        self.dup_btn.pack(side="left")
        self.del_btn = ttk.Button(lrow2, text=T("Διαγραφή", "Delete"), command=self._delete_selected)
        self.del_btn.pack(side="left", padx=8)
        body.add(left_outer, weight=0)

        mid = tk.Frame(body, bg=BORDER)
        mid_in = tk.Frame(mid, bg=TAB_BG)
        mid_in.pack(fill="both", expand=True, padx=1, pady=1)
        mid_in.rowconfigure(0, weight=1)
        mid_in.columnconfigure(0, weight=1)
        self.cv = tk.Canvas(mid_in, bg=TAB_BG, highlightthickness=0, width=640, height=560, takefocus=True)
        vbar = ttk.Scrollbar(mid_in, orient="vertical", command=self.cv.yview)
        hbar = ttk.Scrollbar(mid_in, orient="horizontal", command=self.cv.xview)
        self.cv.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        self.cv.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        self.cv.bind('<Configure>', self._on_canvas_configure)
        self.cv.bind('<ButtonPress-1>', self._on_press)
        self.cv.bind('<B1-Motion>', self._on_drag)
        self.cv.bind('<ButtonRelease-1>', self._on_release)
        self.cv.bind('<Motion>', self._on_hover)
        self.cv.bind('<MouseWheel>', self._on_wheel)
        self.cv.bind('<Button-4>', self._on_wheel)
        self.cv.bind('<Button-5>', self._on_wheel)
        body.add(mid, weight=1)

        right_outer, right = self._card(body)
        self._build_props(right)
        body.add(right_outer, weight=0)

    def _build_props(self, parent):
        self.empty_lbl = ttk.Label(
            parent, style='CardHint.TLabel', justify="left", wraplength=250,
            text=T("Επιλέξτε ένα στοιχείο από τη λίστα ή πατήστε πάνω του στη σελίδα.\n\n"
                   "Μετά σύρετέ το εκεί που θέλετε και διαλέξτε γραμματοσειρά, μέγεθος και χρώμα.",
                   "Choose an item from the list or click it on the page.\n\n"
                   "Then drag it where you want it and pick the font, size and colour."))
        self.empty_lbl.pack(anchor="nw", padx=18, pady=18)
        self.props = ttk.Frame(parent, style='Card.TFrame')
        p = self.props
        p.columnconfigure(1, weight=1)

        self.v_text = tk.StringVar()
        self.v_font = tk.StringVar()
        self.v_size = tk.StringVar()
        self.v_bold = tk.BooleanVar()
        self.v_italic = tk.BooleanVar()
        self.v_color = tk.StringVar()
        self.v_step = tk.StringVar()
        self.v_x = tk.StringVar()
        self.v_y = tk.StringVar()
        self.v_visible = tk.BooleanVar()
        self.v_nudge = tk.StringVar(value="0.5")

        self.p_title = ttk.Label(p, text="", style='CardBold.TLabel', wraplength=250, justify="left")
        self.p_title.grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(14, 0))
        self.p_sub = ttk.Label(p, text="", style='CardHint.TLabel')
        self.p_sub.grid(row=1, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 8))

        def label(row, text):
            ttk.Label(p, text=text, style='Card.TLabel').grid(row=row, column=0, sticky="w", padx=(16, 8), pady=5)

        label(2, T("Κείμενο", "Text"))
        self.e_text = ttk.Entry(p, textvariable=self.v_text, width=18)
        self.e_text.grid(row=2, column=1, sticky="ew", padx=(0, 16))
        self.text_row_widgets = [p.grid_slaves(row=2, column=0)[0], self.e_text]

        label(3, T("Γραμματοσειρά", "Font"))
        fonts = [T("(προεπιλογή)", "(default)")] + self.families
        self.cb_font = ttk.Combobox(p, textvariable=self.v_font, values=fonts, state='readonly', width=18)
        self.cb_font.grid(row=3, column=1, sticky="ew", padx=(0, 16))
        self.cb_font.bind('<<ComboboxSelected>>', lambda e: self._prop_changed())

        label(4, T("Μέγεθος", "Size"))
        self.sp_size = ttk.Spinbox(p, textvariable=self.v_size, from_=3, to=96, increment=0.5, width=8,
                                   command=self._prop_changed)
        self.sp_size.grid(row=4, column=1, sticky="w")

        style_row = ttk.Frame(p, style='Card.TFrame')
        style_row.grid(row=5, column=0, columnspan=2, sticky="w", padx=16, pady=(2, 2))
        ttk.Checkbutton(style_row, text=T("Έντονα", "Bold"), variable=self.v_bold, style='Card.TCheckbutton',
                        command=self._prop_changed).pack(side="left")
        ttk.Checkbutton(style_row, text=T("Πλάγια", "Italic"), variable=self.v_italic, style='Card.TCheckbutton',
                        command=self._prop_changed).pack(side="left", padx=14)

        label(6, T("Χρώμα", "Colour"))
        color_row = ttk.Frame(p, style='Card.TFrame')
        color_row.grid(row=6, column=1, sticky="w")
        self.swatch = tk.Label(color_row, width=4, bg='#000000', cursor='hand2', highlightthickness=1,
                               highlightbackground=BORDER)
        self.swatch.pack(side="left")
        self.swatch.bind('<Button-1>', lambda e: self._pick_color())
        ttk.Entry(color_row, textvariable=self.v_color, width=9).pack(side="left", padx=(8, 0))
        quick = ttk.Frame(p, style='Card.TFrame')
        quick.grid(row=7, column=1, sticky="w", pady=(2, 4))
        for col in self.QUICK_COLORS:
            dot = tk.Label(quick, width=2, bg=col, cursor='hand2', highlightthickness=1, highlightbackground=BORDER)
            dot.pack(side="left", padx=(0, 5))
            dot.bind('<Button-1>', lambda e, c=col: self.v_color.set(c))

        label(8, T("Στοίχιση", "Align"))
        self.align_seg = Segmented(p, (('left', T("Αριστερά", "Left")), ('center', T("Κέντρο", "Centre")),
                                       ('right', T("Δεξιά", "Right"))), lambda v: self._prop_changed(), 'left')
        self.align_seg.grid(row=8, column=1, sticky="w")

        label(9, T("Απόσταση ψηφίων", "Digit spacing"))
        self.sp_step = ttk.Spinbox(p, textvariable=self.v_step, from_=0, to=200, increment=0.5, width=8,
                                   command=self._prop_changed)
        self.sp_step.grid(row=9, column=1, sticky="w")
        ttk.Label(p, text=T("0 = κανονικό κείμενο. Αλλιώς κάθε γράμμα μπαίνει σε σταθερή απόσταση (για κουτάκια).",
                            "0 = normal text. Otherwise each character sits at a fixed distance (for digit boxes)."),
                  style='CardHint.TLabel', wraplength=250, justify="left").grid(
            row=10, column=0, columnspan=2, sticky="w", padx=16, pady=(0, 4))

        ttk.Separator(p, orient="horizontal").grid(row=11, column=0, columnspan=2, sticky="ew", padx=16, pady=8)
        self.x_lbl = ttk.Label(p, text="X (pt)", style='Card.TLabel')
        self.x_lbl.grid(row=12, column=0, sticky="w", padx=(16, 8), pady=5)
        self.sp_x = ttk.Spinbox(p, textvariable=self.v_x, from_=-100, to=1000, increment=0.5, width=8,
                                command=self._prop_changed)
        self.sp_x.grid(row=12, column=1, sticky="w")
        self.y_lbl = ttk.Label(p, text="Y (pt)", style='Card.TLabel')
        self.y_lbl.grid(row=13, column=0, sticky="w", padx=(16, 8), pady=5)
        self.sp_y = ttk.Spinbox(p, textvariable=self.v_y, from_=-1000, to=1500, increment=0.5, width=8,
                                command=self._prop_changed)
        self.sp_y.grid(row=13, column=1, sticky="w")

        pad = ttk.Frame(p, style='Card.TFrame')
        pad.grid(row=14, column=0, columnspan=2, sticky="w", padx=16, pady=(6, 2))
        for col, (txt, dx, dy) in enumerate((("◀", -1, 0), ("▲", 0, 1), ("▼", 0, -1), ("▶", 1, 0))):
            ttk.Button(pad, text=txt, width=3, command=lambda a=dx, b=dy: self._nudge(a, b)).grid(row=0, column=col, padx=(0, 4))
        ttk.Combobox(pad, textvariable=self.v_nudge, values=self.NUDGE_STEPS, width=4).grid(row=0, column=4, padx=(8, 0))
        ttk.Label(pad, text="pt", style='CardHint.TLabel').grid(row=0, column=5, padx=(4, 0))

        ttk.Checkbutton(p, text=T("Εμφάνιση στο PDF", "Show in the PDF"), variable=self.v_visible,
                        style='Card.TCheckbutton', command=self._prop_changed).grid(
            row=15, column=0, columnspan=2, sticky="w", padx=16, pady=(8, 0))
        self.split_sep = ttk.Separator(p, orient="horizontal")
        self.split_sep.grid(row=16, column=0, columnspan=2, sticky="ew", padx=16, pady=(10, 6))
        self.split_btn = ttk.Button(p, text="", command=self._toggle_split)
        self.split_btn.grid(row=17, column=0, columnspan=2, sticky="w", padx=16)
        self.split_hint = ttk.Label(p, text="", style='CardHint.TLabel', wraplength=250, justify="left")
        self.split_hint.grid(row=18, column=0, columnspan=2, sticky="w", padx=16, pady=(4, 0))
        self.sib_btn = ttk.Button(p, text=T("Ίδια μορφή σε όλα τα αδέλφια", "Same format for all siblings"),
                                  command=self._apply_siblings)
        self.sib_btn.grid(row=19, column=0, columnspan=2, sticky="w", padx=16, pady=(6, 0))
        self.reset_btn = ttk.Button(p, text=T("Επαναφορά αυτού του στοιχείου", "Reset this item"),
                                    command=self._reset_item)
        self.reset_btn.grid(row=20, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 16))

        for var in (self.v_text, self.v_size, self.v_color, self.v_step, self.v_x, self.v_y):
            var.trace_add('write', lambda *a: self._prop_changed())

    # ------------------------------------------------------------------ list of items
    def _tree_text(self, fid, it):
        return ("" if it.get('visible', True) else "○  ") + manual_label(fid, it)

    def _fill_tree(self):
        self.tree.delete(*self.tree.get_children())
        q = _fold(self.q_var.get().strip())
        items = self.layout['items']
        groups = dict(MANUAL_GROUPS)
        rows = []
        for fid, it in items.items():
            gid = manual_group(fid, it)
            if q and q not in _fold(manual_label(fid, it) + ' ' + T(*groups.get(gid, (gid, gid)))):
                continue
            rows.append((fid, it, gid))
        used = {g for _f, _i, g in rows}
        for gid, (gel, gen) in MANUAL_GROUPS:
            if not q or gid in used:
                self.tree.insert('', 'end', iid='g:' + gid, text=T(gel, gen), open=(gid != 'totals') or bool(q),
                                 tags=('group',))
        for fid, it, gid in rows:
            parent = 'g:' + gid
            if not q and it.get('parent') in items:       # pieces of a split item sit under it
                parent = it['parent']
            self.tree.insert(parent, 'end', iid=fid, text=self._tree_text(fid, it),
                             tags=() if it.get('visible', True) else ('hidden',))
        self._sync_tree_selection()

    def _sync_tree_selection(self):
        self._sync = True
        try:
            if self.sel and self.tree.exists(self.sel):
                self.tree.selection_set(self.sel)
                self.tree.see(self.sel)
            else:
                self.tree.selection_remove(*self.tree.selection())
        finally:
            self._sync = False

    def _on_tree_select(self, _event=None):
        if self._sync:
            return
        chosen = self.tree.selection()
        if chosen and chosen[0] in self.layout['items'] and chosen[0] != self.sel:
            self._select(chosen[0], from_tree=True)

    # ------------------------------------------------------------------ drawing
    def _px(self, x):
        return self.ox + x * self.zoom

    def _py(self, y):
        return self.PAD + (self.sizes[self.page][1] - y) * self.zoom

    def _page_photo(self):
        key = (self.page, round(self.zoom, 3))
        if key in self._photos:
            return self._photos[key]
        png = render_pdf_page_png(self.template, self.page, self.zoom)
        photo = None
        if png:
            try:
                photo = tk.PhotoImage(data=base64.b64encode(png).decode('ascii'))
            except tk.TclError:
                photo = None
        if len(self._photos) > 8:
            self._photos.clear()
        self._photos[key] = photo
        return photo

    def _draw_all(self):
        cv, z = self.cv, self.zoom
        pw, ph = self.sizes[self.page]
        w, h = pw * z, ph * z
        cw = max(cv.winfo_width(), 100)
        self.ox = max(self.PAD, int((cw - w) / 2))
        cv.delete('all')
        cv.configure(scrollregion=(0, 0, max(self.ox * 2 + w, cw), h + 2 * self.PAD))
        cv.create_rectangle(self.ox + 4, self.PAD + 4, self.ox + w + 4, self.PAD + h + 4, fill=BORDER, outline='')
        self.config(cursor='watch')
        try:
            photo = self._page_photo()
        finally:
            self.config(cursor='')
        if photo:
            cv.create_image(self.ox, self.PAD, image=photo, anchor='nw')
            self.render_note.configure(text="")
        else:
            cv.create_rectangle(self.ox, self.PAD, self.ox + w, self.PAD + h, fill='white', outline='')
            self._draw_fallback_texts()
            self.render_note.configure(text=T("Η σελίδα του εντύπου δεν μπορεί να σχεδιαστεί εδώ - εγκαταστήστε το pypdfium2 "
                                              "(pip install pypdfium2) για πλήρη προβολή.",
                                              "The form page cannot be drawn here - install pypdfium2 "
                                              "(pip install pypdfium2) for the full view."))
        cv.create_rectangle(self.ox, self.PAD, self.ox + w, self.PAD + h, outline=ZONE_IDLE_LINE)
        self.zoom_lbl.configure(text=f"{round(z * 100)}%")
        self._draw_items()

    def _draw_fallback_texts(self):
        if self._fallback_texts is None:
            try:
                self._fallback_texts = template_text_items(self.template)
            except Exception:
                self._fallback_texts = [[], []]
        for text, x, y, size in self._fallback_texts[self.page]:
            self.cv.create_text(self._px(x), self._py(y), text=text, anchor='sw', fill='#9aa5b4',
                                font=(FONT, -max(3, int(size * self.zoom))))

    def _draw_items(self):
        self.cv.delete('item')
        for fid, it in self.layout['items'].items():
            if it['page'] == self.page:
                self._draw_item(fid, it)
        self._draw_selection()

    def _draw_item(self, fid, it):
        if it.get('split'):                     # drawn through its individual pieces
            return
        cv, z = self.cv, self.zoom
        tag = 'it_' + fid
        px_size = max(3, int(round(it['size'] * z)))
        family = it['font'] if it['font'] in self.families else (self.families[0] if self.families else FONT)
        style = ' '.join(s for s, on in (('bold', it['bold']), ('italic', it['italic'])) if on) or 'normal'
        anchor = {'left': 'sw', 'center': 's', 'right': 'se'}.get(it['align'], 'sw')
        color = MUTED if not it.get('visible', True) else it['color']
        drop = px_size * 0.2
        step = it['step']
        sample_on = self.show_sample.get()
        parent, country, source = it.get('parent'), it.get('country'), it.get('source')
        if country:
            y0 = self.row_ys[it['page']].get(country)
            if y0 is None:
                return
            shown = ('12' if parent.endswith('_c') else '34') if sample_on else ''
            entries = [(it['x'], y0 + it['y'], shown)]
        elif parent:
            text = self.sample.get(parent, '')
            idx = it.get('index', 0)
            entries = [(it['x'], it['y'], text[idx] if sample_on and 0 <= idx < len(text) else '')]
        elif fid.startswith('rows'):
            ys = list(self.row_ys[it['page']].values()) or [400.0]
            if sample_on:
                shown = '12' if fid.endswith('_c') else '34'
                entries = [(it['x'], y + it['y'], shown) for y in ys]
            else:
                entries = [(it['x'], ys[0] + it['y'], '')]
        elif it.get('custom'):
            entries = [(it['x'], it['y'], it.get('text', ''))]
        elif source:
            if source.startswith('country:'):
                shown = ('1234' if source.endswith(':c') else '5678') if sample_on else ''
            else:
                shown = self.sample.get(source, '') if sample_on else ''
            entries = [(it['x'], it['y'], shown)]
        else:
            entries = [(it['x'], it['y'], self.sample.get(fid, '') if sample_on else '')]
        for x, y, txt in entries:
            if not txt:           # nothing to show: the item's name stands in, so it can still be found
                cv.create_text(self._px(x), self._py(y) + 2, text=manual_label(fid, it), anchor='sw', fill=ACCENT,
                               font=(FONT, 9), tags=('item', tag))
            elif step:
                for i, ch in enumerate(txt):
                    cv.create_text(self._px(x + i * step), self._py(y) + drop, text=ch, anchor=anchor, fill=color,
                                   font=(family, -px_size, style), tags=('item', tag))
            else:
                cv.create_text(self._px(x), self._py(y) + drop, text=txt, anchor=anchor, fill=color,
                               font=(family, -px_size, style), tags=('item', tag))

    def _draw_selection(self):
        self.cv.delete('selbox')
        if not self.sel or self.sel not in self.layout['items']:
            return
        it = self.layout['items'][self.sel]
        if it['page'] != self.page or it.get('split'):
            return
        box = self.cv.bbox('it_' + self.sel)
        if box:
            self.cv.create_rectangle(box[0] - 3, box[1] - 2, box[2] + 3, box[3] + 2, outline=ACCENT, dash=(4, 3),
                                     tags='selbox')
        x, y = self._px(it['x']), self._py(it['y'] + self._row_base_y(self.sel, it))
        self.cv.create_oval(x - 3, y - 3, x + 3, y + 3, fill=ACCENT, outline='white', tags='selbox')

    @staticmethod
    def _is_rows(fid, it):
        return fid.startswith('rows') and not it.get('parent')

    def _row_base_y(self, fid, it):
        if self._is_rows(fid, it):
            return (list(self.row_ys[it['page']].values()) or [400.0])[0]
        if it.get('country'):
            return self.row_ys[it['page']].get(it['country'], 0.0)
        return 0.0

    def _on_canvas_configure(self, event):
        if self._fit_pending and event.width > 300 and event.height > 200:
            self._fit_pending = False
            self.after(10, self._fit)
            return
        if abs(event.width - self._last_width) > 2:
            self._last_width = event.width
            if self._redraw_job:
                self.after_cancel(self._redraw_job)
            self._redraw_job = self.after(80, self._draw_all)

    def _hit(self, event):
        x, y = self.cv.canvasx(event.x), self.cv.canvasy(event.y)
        for item in reversed(self.cv.find_overlapping(x - 4, y - 4, x + 4, y + 4)):
            for tag in self.cv.gettags(item):
                if tag.startswith('it_'):
                    return tag[3:]
        return None

    def _on_press(self, event):
        self.cv.focus_set()
        fid = self._hit(event)
        self._select(fid)
        self._drag = [self.cv.canvasx(event.x), self.cv.canvasy(event.y), False] if fid else None

    def _on_drag(self, event):
        if not self._drag or not self.sel:
            return
        cx, cy = self.cv.canvasx(event.x), self.cv.canvasy(event.y)
        dx, dy = cx - self._drag[0], cy - self._drag[1]
        if not self._drag[2] and abs(dx) + abs(dy) < 3:
            return
        if not self._drag[2]:
            self._snap(None)
        self._drag[0], self._drag[1], self._drag[2] = cx, cy, True
        it = self.layout['items'][self.sel]
        it['x'] += dx / self.zoom
        it['y'] -= dy / self.zoom
        self.cv.move('it_' + self.sel, dx, dy)
        self.cv.move('selbox', dx, dy)
        self._set_dirty()
        self._load_xy()

    def _on_release(self, _event):
        moved = bool(self._drag and self._drag[2])
        self._drag = None
        if moved and self.sel:
            it = self.layout['items'][self.sel]
            it['x'], it['y'] = round(it['x'], 1), round(it['y'], 1)
            self._load_xy()
            self._draw_items()

    def _on_hover(self, event):
        x = (self.cv.canvasx(event.x) - self.ox) / self.zoom
        y = self.sizes[self.page][1] - (self.cv.canvasy(event.y) - self.PAD) / self.zoom
        self.status_var.set(f"X {x:7.1f}   Y {y:7.1f}  pt")
        self.cv.configure(cursor='fleur' if self._hit(event) else '')

    def _on_wheel(self, event):
        num = getattr(event, 'num', None)
        delta = 120 if num == 4 else -120 if num == 5 else getattr(event, 'delta', 0)
        if not delta:
            return 'break'
        if event.state & 0x4:                              # Ctrl + wheel = zoom
            self._set_zoom(self.zoom * (1.12 if delta > 0 else 1 / 1.12))
        elif event.state & 0x1:                            # Shift + wheel = sideways
            self.cv.xview_scroll(-3 if delta > 0 else 3, 'units')
        else:
            self.cv.yview_scroll(-3 if delta > 0 else 3, 'units')
        return 'break'

    def _typing(self):
        return isinstance(self.focus_get(), (tk.Entry, ttk.Entry, ttk.Spinbox, ttk.Combobox, tk.Text))

    def _on_arrow(self, event):
        if not self.sel or self._typing() or isinstance(self.focus_get(), ttk.Treeview):
            return None
        amount = 5.0 if event.state & 0x1 else 0.5
        dx = {'Left': -1, 'Right': 1}.get(event.keysym, 0)
        dy = {'Up': 1, 'Down': -1}.get(event.keysym, 0)
        self._move_by(dx * amount, dy * amount)
        return 'break'

    def _on_delete_key(self, _event):
        if self.sel and not self._typing():
            self._delete_selected()

    # ------------------------------------------------------------------ view
    def _set_zoom(self, zoom):
        zoom = min(max(zoom, 0.4), 3.0)
        fx, fy = self.cv.xview()[0], self.cv.yview()[0]
        self.zoom = zoom
        self._draw_all()
        self.cv.xview_moveto(fx)
        self.cv.yview_moveto(fy)

    def _fit(self):
        pw, ph = self.sizes[self.page]
        cw, ch = max(self.cv.winfo_width(), 300), max(self.cv.winfo_height(), 200)
        self.zoom = min((cw - 2 * self.PAD) / pw, (ch - 2 * self.PAD) / ph)
        self.zoom = min(max(round(self.zoom, 2), 0.4), 3.0)
        self._draw_all()
        self.cv.xview_moveto(0)
        self.cv.yview_moveto(0)

    def _on_page_seg(self, value):
        self.page = value
        if self.sel and self.layout['items'][self.sel]['page'] != value:
            self.sel = None
            self._sync_tree_selection()
            self._load_props()
        self._draw_all()

    def _reveal(self):
        box = self.cv.bbox('it_' + self.sel) if self.sel else None
        if not box:
            return
        try:
            sr = [float(v) for v in str(self.cv.cget('scrollregion')).split()]
        except ValueError:
            return
        total_w, total_h = max(sr[2] - sr[0], 1), max(sr[3] - sr[1], 1)
        cw, ch = self.cv.winfo_width(), self.cv.winfo_height()
        cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
        x0, y0 = self.cv.canvasx(0), self.cv.canvasy(0)
        if not (x0 + 30 < cx < x0 + cw - 30):
            self.cv.xview_moveto(max(0.0, (cx - cw / 2) / total_w))
        if not (y0 + 30 < cy < y0 + ch - 30):
            self.cv.yview_moveto(max(0.0, (cy - ch / 2) / total_h))

    # ------------------------------------------------------------------ selection / properties
    def _select(self, fid, from_tree=False):
        self.sel = fid if fid in self.layout['items'] else None
        if self.sel and self.layout['items'][self.sel]['page'] != self.page:
            self.page = self.layout['items'][self.sel]['page']
            self.page_seg.set(self.page)
            self._draw_all()
        else:
            self._draw_selection()
        if not from_tree:
            self._sync_tree_selection()
        self._load_props()
        if from_tree:
            self._reveal()

    def _load_props(self):
        if not self.sel:
            self.props.pack_forget()
            self.empty_lbl.pack(anchor="nw", padx=18, pady=18)
            self.del_btn.state(['disabled'])
            self.dup_btn.state(['disabled'])
            self.tip_var.set(self._context_tip())
            return
        self.tip_var.set(self._context_tip())
        items = self.layout['items']
        it = items[self.sel]
        self.empty_lbl.pack_forget()
        self.props.pack(fill="both", expand=True)
        self._loading = True
        try:
            custom = bool(it.get('custom'))
            rows = self._is_rows(self.sel, it)
            rowchild = bool(it.get('country'))
            child = bool(it.get('parent'))
            deletable = custom or bool(it.get('source'))
            self.p_title.configure(text=manual_label(self.sel, it))
            sub = T(f"Σελίδα {it['page'] + 1}", f"Page {it['page'] + 1}")
            if rows and not it.get('split'):
                sub += T(" · μετακινεί όλες τις γραμμές της στήλης", " · moves every row of the column")
            if it.get('split'):
                sub += T(" · χωρισμένο - επεξεργαστείτε τα μέλη του ξεχωριστά",
                         " · split - edit its individual members")
            self.p_sub.configure(text=sub)
            for w in self.text_row_widgets:
                w.grid() if custom else w.grid_remove()
            self.v_text.set(it.get('text', '') if custom else '')
            self.v_font.set(it['font'] if it['font'] in self.families else T("(προεπιλογή)", "(default)"))
            self.v_size.set(f"{it['size']:g}")
            self.v_bold.set(it['bold'])
            self.v_italic.set(it['italic'])
            self.v_color.set(it['color'])
            self.swatch.configure(bg=it['color'])
            self.align_seg.set(it['align'])
            self.v_step.set(f"{it['step']:g}")
            self.v_visible.set(it.get('visible', True))
            self.x_lbl.configure(text=T("X (στήλη, pt)", "X (column, pt)") if rows else "X (pt)")
            self.y_lbl.configure(text=T("Μετατόπιση Y (pt)", "Y offset (pt)") if (rows or rowchild) else "Y (pt)")
            self._load_xy()
            can_reset = not custom and not child and not it.get('source') and self.sel in MANUAL_FIELD_INDEX
            self.reset_btn.state(['!disabled'] if can_reset else ['disabled'])
            self.del_btn.state(['!disabled'] if deletable else ['disabled'])
            self.dup_btn.state(['!disabled'] if (not child and not rows and not it.get('split')) else ['disabled'])
            # individual controls
            tid = it.get('parent') or self.sel
            can_split = (tid.startswith('rows') or tid in MANUAL_SPLIT_LEN) and not custom and not it.get('source')
            if can_split:
                is_split = bool(items[tid].get('split'))
                self.split_btn.configure(text=T("Συγχώνευση ξανά σε ένα", "Merge back into one") if is_split
                                         else T("Χωρισμός σε μεμονωμένα στοιχεία", "Split into individual items"))
                if tid.startswith('rows'):
                    hint = T("Κάθε χώρα γίνεται ξεχωριστό στοιχείο που μετακινείται ελεύθερα (το ύψος γραμμής ακολουθεί το template).",
                             "Each country becomes its own item that moves freely (the row height still follows the template).")
                else:
                    hint = T("Κάθε ψηφίο γίνεται ξεχωριστό στοιχείο.", "Each digit becomes its own item.")
                self.split_hint.configure(text=hint)
                for w in (self.split_sep, self.split_btn, self.split_hint):
                    w.grid()
            else:
                for w in (self.split_sep, self.split_btn, self.split_hint):
                    w.grid_remove()
            self.sib_btn.grid() if child else self.sib_btn.grid_remove()
        finally:
            self._loading = False

    def _load_xy(self):
        if not self.sel:
            return
        it = self.layout['items'][self.sel]
        was, self._loading = self._loading, True
        try:
            self.v_x.set(f"{it['x']:.1f}")
            self.v_y.set(f"{it['y']:.1f}")
        finally:
            self._loading = was

    @staticmethod
    def _num(var, low, high):
        try:
            return min(max(float(var.get().replace(',', '.')), low), high)
        except ValueError:
            return None

    def _prop_changed(self, *_):
        if self._loading or not self.sel:
            return
        self._snap('prop:' + self.sel)
        it = self.layout['items'][self.sel]
        if it.get('custom'):
            it['text'] = self.v_text.get()
        font = self.v_font.get()
        it['font'] = font if font in self.families else ''
        for key, var, low, high in (('size', self.v_size, 3, 96), ('step', self.v_step, 0, 200),
                                    ('x', self.v_x, -100, 1000), ('y', self.v_y, -1000, 1500)):
            value = self._num(var, low, high)
            if value is not None:
                it[key] = value
        it['bold'], it['italic'] = bool(self.v_bold.get()), bool(self.v_italic.get())
        color = self.v_color.get().strip()
        if re.fullmatch(r'#[0-9a-fA-F]{6}', color):
            it['color'] = color.lower()
            self.swatch.configure(bg=it['color'])
        it['align'] = self.align_seg.value
        it['visible'] = bool(self.v_visible.get())
        self.p_title.configure(text=manual_label(self.sel, it))
        if self.tree.exists(self.sel):
            self.tree.item(self.sel, text=self._tree_text(self.sel, it),
                           tags=() if it['visible'] else ('hidden',))
        self._set_dirty()
        self._draw_items()

    def _pick_color(self):
        if not self.sel:
            return
        current = self.layout['items'][self.sel]['color']
        chosen = colorchooser.askcolor(color=current, parent=self, title=T("Χρώμα κειμένου", "Text colour"))
        if chosen and chosen[1]:
            self.v_color.set(chosen[1])

    def _move_by(self, dx, dy):
        if not self.sel:
            return
        self._snap('move:' + self.sel)
        it = self.layout['items'][self.sel]
        it['x'] = round(it['x'] + dx, 2)
        it['y'] = round(it['y'] + dy, 2)
        self._load_xy()
        self._set_dirty()
        self._draw_items()

    def _nudge(self, dx, dy):
        amount = self._num(self.v_nudge, 0.01, 100) or 0.5
        self._move_by(dx * amount, dy * amount)

    def _reset_item(self):
        if not self.sel:
            return
        it = self.layout['items'][self.sel]
        if it.get('custom') or it.get('source') or it.get('parent') or self.sel not in MANUAL_FIELD_INDEX:
            return
        self._snap(None)
        manual_merge_item(self.layout, self.sel)
        self.layout['items'][self.sel] = manual_default_layout()['items'][self.sel]
        self._set_dirty()
        self._fill_tree()
        self._load_props()
        self._draw_items()

    def _add_custom(self):
        self._snap(None)
        n = 1
        while f'custom_{n}' in self.layout['items']:
            n += 1
        fid = f'custom_{n}'
        pw, ph = self.sizes[self.page]
        try:                                           # drop it in the middle of what is on screen
            cx = (self.cv.canvasx(self.cv.winfo_width() / 2) - self.ox) / self.zoom
            cy = ph - (self.cv.canvasy(self.cv.winfo_height() / 2) - self.PAD) / self.zoom
        except tk.TclError:
            cx, cy = pw / 2, ph / 2
        cx, cy = min(max(cx, 20), pw - 120), min(max(cy, 20), ph - 20)
        it = manual_new_item(self.page, round(cx, 1), round(cy, 1), custom=True, text=T("Νέο κείμενο", "New text"))
        self.layout['items'][fid] = it
        self.tree.insert('g:custom', 'end', iid=fid, text=self._tree_text(fid, it))
        self._set_dirty()
        self._select(fid)
        self._draw_items()

    def _delete_selected(self):
        if not self.sel:
            return
        it = self.layout['items'][self.sel]
        if not (it.get('custom') or it.get('source')):
            return
        self._snap(None)
        fid = self.sel
        self.sel = None
        del self.layout['items'][fid]
        self._set_dirty()
        self._fill_tree()
        self._load_props()
        self._draw_items()

    def _reset_all(self):
        if not messagebox.askyesno(T("Επαναφορά όλων", "Reset all"),
                                   T("Όλες οι θέσεις και οι μορφοποιήσεις θα γυρίσουν στις προεπιλεγμένες και τα δικά σας "
                                     "κείμενα θα διαγραφούν.\n\nΣυνέχεια;",
                                     "Every position and format goes back to the default and your own texts are "
                                     "deleted.\n\nContinue?"), default='no', parent=self):
            return
        self._snap(None)
        self.layout = manual_default_layout()
        self.sel = None
        self._fill_tree()
        self._set_dirty()
        self._load_props()
        self._draw_items()

    def _context_tip(self):
        """A short tip that fits what is selected right now."""
        if not self.sel or self.sel not in self.layout['items']:
            return T("Συμβουλή: πατήστε ένα στοιχείο στη σελίδα ή στη λίστα για να το μετακινήσετε. "
                     "Ctrl+F = αναζήτηση · + Προσθήκη στοιχείου = βάλτε ό,τι άλλο θέλετε στο δελτίο.",
                     "Tip: click an item on the page or in the list to move it. "
                     "Ctrl+F = search · + Add item = put anything else you want on the form.")
        it = self.layout['items'][self.sel]
        if it.get('split'):
            return T("Αυτό το στοιχείο είναι χωρισμένο: ανοίξτε το στη λίστα και επιλέξτε ένα μέλος του, ή πατήστε πάνω του στη σελίδα.",
                     "This item is split: open it in the list and pick one of its members, or click it on the page.")
        if it.get('country'):
            return T("Μεμονωμένη χώρα: το X είναι ελεύθερο και το Y είναι μετατόπιση από τη γραμμή της στο template. Απλώς σύρτε την.",
                     "Individual country: X is free and Y is an offset from its row in the template. Just drag it.")
        if it.get('parent'):
            return T("Ένα ψηφίο: μετακινείται μόνο του. «Ίδια μορφή σε όλα τα αδέλφια» περνά τη μορφή του στα υπόλοιπα ψηφία.",
                     "One digit: it moves on its own. “Same format for all siblings” copies its format to the other digits.")
        if self._is_rows(self.sel, it):
            return T("Η στήλη μετακινείται ως σύνολο. Για να μετακινείτε κάθε χώρα ξεχωριστά πατήστε «Χωρισμός σε μεμονωμένα στοιχεία».",
                     "The column moves as one. To move each country separately press “Split into individual items”.")
        if self.sel in MANUAL_SPLIT_LEN:
            return T("Κουτάκια ψηφίων: ρυθμίστε την «Απόσταση ψηφίων» ή χωρίστε το σε μεμονωμένα ψηφία για πλήρη ελευθερία.",
                     "Digit boxes: adjust the “Digit spacing” or split it into single digits for full freedom.")
        if it.get('custom'):
            return T("Δικό σας κείμενο: αλλάξτε το από το πεδίο «Κείμενο». Ctrl+D = αντίγραφο, Delete = διαγραφή.",
                     "Your own text: edit it in the “Text” field. Ctrl+D = copy, Delete = remove.")
        if it.get('source'):
            return T("Στοιχείο δεδομένων: η τιμή του μπαίνει αυτόματα όταν δημιουργείται το δελτίο. Μπορείτε να το βάλετε όσες φορές θέλετε.",
                     "Data item: its value is filled in automatically when the form is created. You can place it as many times as you like.")
        return T("Σύρτε για μετακίνηση · βελάκια = 0,5 pt (Shift = 5 pt) · «Εμφάνιση στο PDF» το κρύβει χωρίς να το διαγράψει.",
                 "Drag to move · arrow keys = 0.5 pt (Shift = 5 pt) · “Show in the PDF” hides it without deleting it.")

    def _open_guide(self):
        """F1: open the guide at the layout-editor section (the editor stops being modal so the main window can be used)."""
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.app.open_help('editor')
        self.app.lift()
        self.app.focus_force()

    # ------------------------------------------------------------------ undo / redo
    def _snap(self, key):
        """Remembers the layout before a change. The same `key` within a second counts as one change."""
        now = time.time()
        if key is not None and key == self._snap_key and now - self._snap_time < 1.0:
            self._snap_time = now
            return
        self._snap_key, self._snap_time = key, now
        self.undo.append(copy.deepcopy(self.layout['items']))
        del self.undo[:-60]
        self.redo.clear()
        self._update_undo_buttons()

    def _update_undo_buttons(self):
        for btn, stack in ((getattr(self, 'undo_btn', None), self.undo), (getattr(self, 'redo_btn', None), self.redo)):
            if btn is not None:
                btn.state(['!disabled'] if stack else ['disabled'])

    def _restore(self):
        self._snap_key = None
        if self.sel not in self.layout['items']:
            self.sel = None
        self._fill_tree()
        self._load_props()
        self._draw_items()
        self._set_dirty()
        self._update_undo_buttons()

    def _undo(self):
        if self.undo:
            self.redo.append(copy.deepcopy(self.layout['items']))
            self.layout['items'] = self.undo.pop()
            self._restore()

    def _redo(self):
        if self.redo:
            self.undo.append(copy.deepcopy(self.layout['items']))
            self.layout['items'] = self.redo.pop()
            self._restore()

    # ------------------------------------------------------------------ individual items / catalog
    def _toggle_split(self):
        if not self.sel:
            return
        items = self.layout['items']
        tid = items[self.sel].get('parent') or self.sel
        target = items[tid]
        self._snap(None)
        if target.get('split'):
            manual_merge_item(self.layout, tid)
            new_sel = tid
        elif tid.startswith('rows'):
            ys = sorted(self.row_ys[target['page']].items(), key=lambda kv: -kv[1])
            made = manual_split_item(self.layout, tid, '', [k for k, _ in ys])
            new_sel = made[0] if made else tid
        else:
            made = manual_split_item(self.layout, tid, '0' * MANUAL_SPLIT_LEN.get(tid, 0), [])
            new_sel = made[0] if made else tid
        self._set_dirty()
        self._fill_tree()
        self._draw_items()
        self._select(new_sel)

    def _apply_siblings(self):
        items = self.layout['items']
        it = items.get(self.sel or '')
        if not it or not it.get('parent'):
            return
        self._snap(None)
        for other in items.values():
            if other.get('parent') == it['parent']:
                for key in ('font', 'size', 'bold', 'italic', 'color', 'align'):
                    other[key] = it[key]
        self._set_dirty()
        self._draw_items()

    def _view_center(self):
        pw, ph = self.sizes[self.page]
        try:
            cx = (self.cv.canvasx(self.cv.winfo_width() / 2) - self.ox) / self.zoom
            cy = ph - (self.cv.canvasy(self.cv.winfo_height() / 2) - self.PAD) / self.zoom
        except tk.TclError:
            cx, cy = pw / 2, ph / 2
        self._stagger = (self._stagger + 1) % 6          # items added one after another do not hide each other
        return min(max(cx, 20), pw - 120) + 12 * self._stagger, min(max(cy, 20), ph - 20) - 12 * self._stagger

    def _free_id(self, prefix):
        n = 1
        while f'{prefix}{n}' in self.layout['items']:
            n += 1
        return f'{prefix}{n}'

    def _add_source_item(self, key):
        self._snap(None)
        fid = self._free_id('extra_')
        cx, cy = self._view_center()
        it = manual_new_item(self.page, round(cx, 1), round(cy, 1), size=10.0)
        it['source'] = key
        self.layout['items'][fid] = it
        self._set_dirty()
        self._fill_tree()
        self._draw_items()
        self._select(fid)

    def _duplicate(self):
        if not self.sel:
            return
        items = self.layout['items']
        it = items[self.sel]
        if it.get('parent') or it.get('split') or self._is_rows(self.sel, it):
            return
        self._snap(None)
        new = copy.deepcopy(it)
        new['x'], new['y'] = round(it['x'] + 10, 1), round(it['y'] - 12, 1)
        if it.get('custom'):
            fid = self._free_id('custom_')
        else:
            fid = self._free_id('extra_')
            new['source'] = it.get('source') or self.sel
        items[fid] = new
        self._set_dirty()
        self._fill_tree()
        self._draw_items()
        self._select(fid)

    def _add_item_dialog(self):
        win = tk.Toplevel(self)
        win.title(T("Προσθήκη στοιχείου", "Add item"))
        win.configure(bg=BG)
        win.transient(self)
        box = ttk.Frame(win, padding=14)
        box.pack(fill="both", expand=True)
        ttk.Label(box, text=T("Αναζητήστε ό,τι θέλετε να βάλετε στο δελτίο και πατήστε «Προσθήκη» (ή διπλό κλικ). "
                              "Μπορείτε να προσθέσετε το ίδιο στοιχείο όσες φορές θέλετε και να το τοποθετήσετε ελεύθερα.",
                              "Search for anything you want to put on the form and press “Add” (or double-click). "
                              "You can add the same item as many times as you like and place it freely."),
                  style='Hint.TLabel', wraplength=430, justify="left").pack(anchor="w")
        qv = tk.StringVar()
        ent = ttk.Entry(box, textvariable=qv)
        ent.pack(fill="x", pady=8)
        frame = ttk.Frame(box)
        frame.pack(fill="both", expand=True)
        tv = ttk.Treeview(frame, columns=('k',), show='tree headings', height=14, selectmode='extended')
        tv.heading('#0', text=T("Στοιχείο", "Item"), anchor='w')
        tv.heading('k', text=T("Ομάδα", "Group"), anchor='w')
        tv.column('#0', width=290, stretch=True)
        tv.column('k', width=130, stretch=False)
        sb = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        kinds = dict(MANUAL_GROUPS)
        kinds['country'] = ("Αριθμοί χωρών", "Country numbers")
        catalog = manual_source_catalog()

        def fill(*_):
            tv.delete(*tv.get_children())
            q = _fold(qv.get().strip())
            for key, label, gid in catalog:
                text, kind = T(*label), T(*kinds.get(gid, (gid, gid)))
                if q and q not in _fold(text + ' ' + kind):
                    continue
                tv.insert('', 'end', iid=key, text=text, values=(kind,))

        def add(_e=None):
            for key in tv.selection():
                self._add_source_item(key)
            return 'break'

        def close(_e=None):
            try:
                win.grab_release()
            except tk.TclError:
                pass
            win.destroy()
            self._try_grab()

        qv.trace_add('write', fill)
        fill()
        btns = ttk.Frame(box)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text=T("Προσθήκη", "Add"), style='PrintSm.TButton', command=add).pack(side="left")
        ttk.Button(btns, text=T("Κλείσιμο", "Close"), command=close).pack(side="right")
        tv.bind('<Double-1>', add)
        ent.bind('<Return>', add)
        win.bind('<Escape>', close)
        win.protocol('WM_DELETE_WINDOW', close)
        win.geometry(f"480x560+{self.winfo_rootx() + 90}+{self.winfo_rooty() + 70}")
        try:
            win.grab_set()
        except tk.TclError:
            pass
        ent.focus_set()

    # ------------------------------------------------------------------ saving / closing / testing
    def _set_dirty(self, value=True):
        if self.dirty != value:
            self.dirty = value
            self._update_title()

    def _update_title(self):
        self.title(T("Επεξεργαστής διάταξης", "Layout editor") + (" *" if self.dirty else ""))

    def save(self):
        try:
            save_manual_layout(self.layout)
        except OSError as e:
            messagebox.showerror(T("Αποθήκευση", "Save"),
                                 T(f"Δεν ήταν δυνατή η αποθήκευση:\n{e}", f"Could not save:\n{e}"), parent=self)
            return False
        self._set_dirty(False)
        self.status_var.set(T("✓ Η διάταξη αποθηκεύτηκε.", "✓ The layout was saved."))
        if self.on_save:
            self.on_save()
        return True

    def _close(self):
        if self.dirty:
            answer = messagebox.askyesnocancel(
                T("Επεξεργαστής διάταξης", "Layout editor"),
                T("Υπάρχουν αλλαγές που δεν έχουν αποθηκευτεί.\n\nΝα αποθηκευτούν;",
                  "There are changes that are not saved.\n\nSave them?"), parent=self)
            if answer is None:
                return
            if answer and not self.save():
                return
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()

    def _test_pdf(self):
        title = T("Δοκιμαστικό PDF", "Test PDF")
        sample, mon = manual_sample_inputs()
        rows = [(name, 12, 34) for name in sorted(MAPPED_TARGETS)]
        path = os.path.join(tempfile.gettempdir(), "ELSTAT_manual_layout_test.pdf")
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            process_elstat_report(None, self.template, path, sample, mon, rows_override=rows, save_history=False,
                                  pdf_mode='manual', manual_layout=self.layout)
        except PermissionError:
            messagebox.showerror(title, T("Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό και δοκιμάστε ξανά.",
                                          "I can't write the PDF. Close it if it is open and try again."), parent=self)
            return
        except Exception as e:
            messagebox.showerror(title, T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"), parent=self)
            return
        finally:
            self.config(cursor='')
        open_file(path)
        self.status_var.set(T("Το δοκιμαστικό PDF έχει πλασματικά στοιχεία - μην το υποβάλετε.",
                              "The test PDF has made-up data - do not submit it."))


class ElstatApp(TkBase):
    def __init__(self):
        super().__init__()
        self.geometry("920x840")
        self.minsize(640, 560)
        self.configure(bg=BG)

        self._storage_notices = init_storage()
        load_custom_country_mappings()
        self.prefs = load_prefs()
        set_language(self.prefs['language'])
        set_theme(self.prefs['theme'])
        self.configure(bg=BG)
        self.current_settings = load_hotel_settings()
        self._rebuilding = False
        self.output_custom = False      # True once the user picks the output file themselves
        self._busy = False
        self._loaded_excel = None
        self._help_last = 'start'       # section of the user guide that was open last
        self._help_prev = None          # tab to return to from the guide
        self._last_tab = None

        setup_styles(self)
        self._init_vars()
        self._build_ui()
        install_edit_shortcuts(self)
        self._register_drag_drop_targets()
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
        self.bind('<F1>', lambda e: self.open_help())
        self.bind('<Control-f>', self._help_find)
        self.bind('<Control-Return>', self._try_generate)
        self.bind('<Control-p>', lambda e: self._try_generate(print_after=True))
        self._center()
        if not DND_AVAILABLE and DND_ERROR:
            self._status("Drag & drop unavailable: " + DND_ERROR)
        elif not DND_AVAILABLE:
            self._status(T("Για μεταφορά αρχείων με drag & drop εγκαταστήστε: pip install tkinterdnd2",
                           f"For drag-and-drop support run: \"{sys.executable}\" -m pip install tkinterdnd2"))
        self.after(100, self._apply_titlebar)
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
        self.mon_vars = {k: tk.StringVar() for k in ('year', 'month', 'date_to', 'days_of_month', 'employees', 'sign_date')}
        self.mon_vars['sign_date'].set(datetime.date.today().strftime('%d/%m/%Y'))   # today unless changed
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
                   for k in ('language', 'theme', 'output_mode', 'output_dir', 'filename_pattern', 'excel_dir', 'pdf_mode')}
        self.pv.update({k: tk.BooleanVar(value=self.prefs[k]) for k in BOOL_PREFS})
        self.data_dir_var = tk.StringVar(value=DATA_DIR)
        self.name_note_var = tk.StringVar()
        self.out_note_var = tk.StringVar()
        self.layout_note_var = tk.StringVar()

    def _build_ui(self):
        self.title(T("ΕΛΣΤΑΤ – Μηνιαίο Δελτίο Κίνησης Καταλύματος",
                     "ELSTAT – Monthly Accommodation Movement Report"))
        header = tk.Frame(self, bg=HEADER_BG)
        header.pack(fill="x")
        row = tk.Frame(header, bg=HEADER_BG)
        row.pack(fill="x", padx=26, pady=18)
        tk.Label(row, text=T("ΕΛΣΤΑΤ", "ELSTAT"), bg=ACCENT, fg="white", font=(FONT, 10, 'bold'),
                 padx=13, pady=8).pack(side="left", anchor="n")
        self.theme_btn = tk.Label(row, text="\u2600" if THEME == 'dark' else "\u263e", bg=HEADER_BG, fg=HEADER_MUTED,
                                  font=(FONT, 17), cursor='hand2', padx=9)
        self.theme_btn.pack(side="right", anchor="n")
        self.theme_btn.bind('<Button-1>', lambda e: self._toggle_theme())
        self.theme_btn.bind('<Enter>', lambda e: self.theme_btn.configure(fg=HEADER_FG))
        self.theme_btn.bind('<Leave>', lambda e: self.theme_btn.configure(fg=HEADER_MUTED))
        self.help_btn = tk.Label(row, text="?", bg=HEADER_BG, fg=HEADER_MUTED, font=(FONT, 15, 'bold'), cursor='hand2', padx=9)
        self.help_btn.pack(side="right", anchor="n")
        self.help_btn.bind('<Button-1>', lambda e: self.open_help())
        self.help_btn.bind('<Enter>', lambda e: self.help_btn.configure(fg=HEADER_FG))
        self.help_btn.bind('<Leave>', lambda e: self.help_btn.configure(fg=HEADER_MUTED))
        texts = tk.Frame(row, bg=HEADER_BG)
        texts.pack(side="left", fill="x", expand=True, padx=(14, 0))
        title = tk.Label(texts, text=T("Μηνιαίο Δελτίο Κίνησης Τουριστικού Καταλύματος",
                                       "Monthly Tourist Accommodation Movement Report"), bg=HEADER_BG, fg=HEADER_FG,
                         font=(FONT, 16, 'bold'), justify="left")
        title.pack(anchor="w")
        sub = tk.Label(texts, text=T("Διαλέξτε το Excel από το Protel και το πρόγραμμα συμπληρώνει και "
                                     "εκτυπώνει το δελτίο της ΕΛΣΤΑΤ.",
                                     "Choose the Excel from Protel and the program fills in and "
                                     "prints the ELSTAT form."), bg=HEADER_BG, fg=HEADER_MUTED,
                       font=(FONT, 10), justify="left")
        sub.pack(anchor="w", pady=(2, 0))
        autowrap(title, header, margin=150)
        autowrap(sub, header, margin=150)
        tk.Frame(self, height=3, bg=ACCENT).pack(fill="x")      # subtle accent under the branded header

        bar = ttk.Frame(self, style='Card.TFrame')
        bar.pack(side="bottom", fill="x")
        ttk.Separator(bar, orient="horizontal").pack(fill="x")
        ttk.Label(bar, textvariable=self.status_var, style='Status.TLabel').pack(fill="x")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=18, pady=(14, 0))
        self.tab_report = ttk.Frame(self.notebook)
        self.tab_hotel = ttk.Frame(self.notebook)
        self.tab_history = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)
        self.tab_help = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_report, text=T("Δελτίο μήνα", "Monthly form"))
        self.notebook.add(self.tab_hotel, text=T("Στοιχεία ξενοδοχείου", "Hotel details"))
        self.notebook.add(self.tab_history, text=T("Ιστορικό", "History"))
        self.notebook.add(self.tab_settings, text=T("Ρυθμίσεις", "Settings"))
        self.notebook.add(self.tab_help, text=T("Οδηγός", "Guide"))
        self.notebook.bind('<<NotebookTabChanged>>', self._tab_changed)
        self._build_report_tab()
        self._build_hotel_tab()
        self._build_history_tab()
        self._build_settings_tab()
        self.guide = HelpPanel(self.tab_help, self, self._help_last)
        self.guide.pack(fill="both", expand=True)
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
        self.reminder_frame = tk.Frame(body, bg=WARN_BG, highlightthickness=1, highlightbackground=WARN_LINE)
        self.reminder_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(12, 0))
        self.reminder_label = tk.Label(self.reminder_frame, text="", bg=WARN_BG, fg=WARN_FG,
                                       font=(FONT, 10), justify="left", anchor="w")
        self.reminder_label.pack(side="left", fill="x", expand=True, padx=12, pady=9)
        ttk.Button(self.reminder_frame, text=T("Κλείσιμο", "Close"), command=self._dismiss_reminder).pack(side="right", padx=8)
        autowrap(self.reminder_label, body, margin=190)
        self.reminder_frame.grid_remove()

        # ---- Step 1
        f1 = self.excel_card = Card(body, 1, T("Αρχείο από το Protel", "File from Protel"))
        f1.grid(row=1, column=0, sticky="ew", padx=4, pady=(12, 6))
        c1 = f1.body
        c1.columnconfigure(0, weight=1)
        self.excel_zone = DropZone(c1, self.browse_excel)
        self.excel_zone.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(2, 12))
        ent = ttk.Entry(c1, textvariable=self.excel_var)
        self.excel_path_entry = ent
        ent.grid(row=1, column=0, sticky="ew")
        ent.bind('<FocusOut>', self._excel_typed)
        ent.bind('<Return>', self._excel_typed)
        self.excel_btn = ttk.Button(c1, text=T("Επιλογή Excel…", "Choose Excel…"), command=self.browse_excel)
        self.excel_btn.grid(row=1, column=1, padx=(8, 0))
        ttk.Button(c1, text=T("Πιο πρόσφατο", "Latest"), command=self.use_recent_excel).grid(row=1, column=2, padx=(6, 0))
        self.period_label = ttk.Label(c1, textvariable=self.period_var, style='CardHint.TLabel', justify="left")
        self.period_label.grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))
        autowrap(self.period_label, body, margin=110)

        # live preview of what the Excel contains (appears after a file is chosen)
        self.preview_frame = ttk.Frame(c1, style='Card.TFrame')
        self.preview_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 0))
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
        self.preview_tree.tag_configure('bad', foreground=BAD, background=BAD_BG)
        self.preview_tree.tag_configure('ok', background=CARD)
        psb = ttk.Scrollbar(self.preview_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=psb.set)
        self.preview_tree.grid(row=1, column=0, sticky="ew")
        psb.grid(row=1, column=1, sticky="ns")
        self.preview_tree.bind('<Double-1>', self._map_selected_country)
        self.map_country_btn = ttk.Button(
            self.preview_frame,
            text=T("Προσθήκη / αντιστοίχιση επιλεγμένης χώρας…", "Add / map selected country…"),
            command=self._map_selected_country)
        self.map_country_btn.grid(row=2, column=0, sticky="w", pady=(8, 2))
        self.preview_hint = ttk.Label(
            self.preview_frame,
            text=T("Επιλέξτε την κόκκινη χώρα και πατήστε εδώ για να τη βάλετε στη σωστή γραμμή του δελτίου.",
                   "Select a red country and click here to assign it to the correct form row."),
            style='CardHint.TLabel', wraplength=620, justify='left')
        self.preview_hint.grid(row=3, column=0, columnspan=2, sticky="w", pady=(2, 0))
        self.preview_frame.grid_remove()
        self._render_period()

        # ---- Step 2
        f2 = self.month_card = Card(body, 2, T("Στοιχεία του μήνα", "Month details"))
        f2.grid(row=2, column=0, sticky="ew", padx=4, pady=6)
        c2 = f2.body
        c2.columnconfigure(2, weight=1)
        c2.columnconfigure(3, minsize=24)
        self.month_field_status = {}
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
            (T("Ημερομηνία υπογραφής", "Signing date"), 'sign_date', 14,
             T("συμπληρώνεται αυτόματα με τη σημερινή - αλλάξτε την (ΗΗ/ΜΜ/ΕΕΕΕ) αν το δελτίο "
               "πρέπει να φαίνεται υπογεγραμμένο άλλη μέρα",
               "filled in automatically with today's date - change it (DD/MM/YYYY) if the form "
               "should look signed on another day")),
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
                h = ttk.Label(c2, text=hint, style='CardWarn.TLabel' if key == 'employees' else 'CardHint.TLabel',
                              justify="left")
                h.grid(row=i, column=2, sticky="w", padx=12)
                autowrap(h, body, margin=430, minimum=140)
            # Small completion marker at the end of each row; it follows the app's
            # existing card styling and updates whenever the field changes.
            marker = ttk.Label(c2, text="○", style='CardHint.TLabel',
                               font=(FONT, 12, 'bold'), anchor='center', width=2)
            marker.grid(row=i, column=3, sticky='e', padx=(8, 0), pady=5)
            self.month_field_status[key] = marker
            self.mon_vars[key].trace_add('write', self._update_month_field_status)

        self.month_status_legend = ttk.Label(
            c2, text=T("● Συμπληρωμένο     ○ Εκκρεμεί", "● Complete     ○ Pending"),
            style='CardHint.TLabel', font=(FONT, 8))
        self.month_status_legend.grid(row=len(fields), column=1, columnspan=3,
                                      sticky='w', padx=(0, 8), pady=(5, 0))
        self._update_month_field_status()

        # ---- Step 3
        f3 = self.output_card = Card(body, 3, T("Πού να αποθηκευτεί το PDF", "Where to save the PDF"))
        f3.grid(row=3, column=0, sticky="ew", padx=4, pady=(6, 14))
        c3 = f3.body
        c3.columnconfigure(0, weight=1)
        self.output_path_entry = ttk.Entry(c3, textvariable=self.out_var)
        self.output_path_entry.grid(row=0, column=0, sticky="ew")
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
                                 "Ctrl+P = δημιουργία & εκτύπωση  ·  F1 = οδηγός",
                                 "Shortcuts:  Ctrl+O = choose Excel  ·  Ctrl+Enter = create  ·  "
                                 "Ctrl+P = create & print  ·  F1 = guide"),
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

        card2 = self.template_card = Card(body, None, T("Άδειο PDF της ΕΛΣΤΑΤ (template)", "Blank ELSTAT PDF (template)"))
        card2.grid(row=2, column=0, sticky="ew", padx=4, pady=(6, 14))
        f2 = card2.body
        f2.columnconfigure(0, weight=1)
        self.template_path_entry = ttk.Entry(f2, textvariable=self.pdf_var)
        self.template_path_entry.grid(row=0, column=0, sticky="ew")
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

        # Give the Settings page a little more breathing room without changing
        # its controls, wording, colours, or behaviour.
        def polish_settings_card(card):
            card.configure(highlightthickness=1, highlightbackground=BORDER)
            children = card.winfo_children()
            if children:
                inner = children[0]
                inner_children = inner.winfo_children()
                if inner_children:
                    head = inner_children[0]
                    head.pack_configure(pady=(18, 5))
                card.body.configure(padding=(22, 11, 22, 21))

        def hint(parent, text, row, columnspan=1, pady=(8, 0), margin=110, padx=0):
            lbl = ttk.Label(parent, text=text, style='CardHint.TLabel', justify="left")
            lbl.grid(row=row, column=0, columnspan=columnspan, sticky="w", pady=pady, padx=padx)
            autowrap(lbl, body, margin=margin)
            return lbl

        def commit_on(entry):
            entry.bind('<FocusOut>', self._pref_changed)
            entry.bind('<Return>', self._pref_changed)

        # ---- 0. interface language
        c0 = Card(body, None, T("Γλώσσα και εμφάνιση / Language and appearance", "Language and appearance"))
        c0.grid(row=0, column=0, sticky="ew", padx=4, pady=(12, 6))
        polish_settings_card(c0)
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
        theme_row = ttk.Frame(f0, style='Card.TFrame')
        theme_row.grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Label(theme_row, text=T("Εμφάνιση:", "Appearance:"), style='CardBold.TLabel').pack(side="left", padx=(0, 16))
        for code, label in (('light', T("Φωτεινή", "Light")), ('dark', T("Σκούρα", "Dark")),
                            ('auto', T("Όπως το σύστημα", "Match system"))):
            ttk.Radiobutton(theme_row, text=label, value=code, variable=pv['theme'],
                            style='Card.TRadiobutton',
                            command=self._theme_clicked).pack(side="left", padx=(0, 24))

        # ---- 1. where the PDFs go
        c1 = Card(body, None, T("Πού αποθηκεύονται τα PDF", "Where PDFs are saved"))
        c1.grid(row=1, column=0, sticky="ew", padx=4, pady=6)
        polish_settings_card(c1)
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
        polish_settings_card(c2)
        f2 = c2.body
        f2.columnconfigure(0, weight=1)
        self.data_dir_entry = ttk.Entry(f2, textvariable=self.data_dir_var, state='readonly')
        self.data_dir_entry.grid(row=0, column=0, sticky="ew")
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
        polish_settings_card(c3)
        f3 = c3.body
        f3.columnconfigure(0, weight=1)
        ex_entry = ttk.Entry(f3, textvariable=pv['excel_dir'])
        self.excel_dir_entry = ex_entry
        ex_entry.grid(row=0, column=0, sticky="ew")
        commit_on(ex_entry)
        ttk.Button(f3, text=T("Αναζήτηση…", "Browse…"), command=self.browse_excel_dir).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(f3, text=T("Καθαρισμός", "Clear"), command=self.clear_excel_dir).grid(row=0, column=2, padx=(6, 0))
        hint(f3, T("Από ποιον φάκελο ξεκινούν τα «Επιλογή Excel…» και «Πιο πρόσφατο». "
                   "Κενό = ο φάκελος του τελευταίου Excel που χρησιμοποιήσατε.",
                   "The folder where “Choose Excel…” and “Latest” start. "
                   "Empty = the folder of the last Excel you used."), 1, columnspan=3)

        # ---- 3b. how the PDF is filled in
        cp = Card(body, None, T("Τρόπος συμπλήρωσης του PDF", "How the PDF is filled in"))
        cp.grid(row=4, column=0, sticky="ew", padx=4, pady=6)
        polish_settings_card(cp)
        fp = cp.body
        fp.columnconfigure(0, weight=1)
        ttk.Radiobutton(fp, text=T("Πάνω στο template (προεπιλογή)", "On top of the template (default)"),
                        value='overlay', variable=pv['pdf_mode'], style='Card.TRadiobutton',
                        command=self._pdf_mode_clicked).grid(row=0, column=0, sticky="w")
        hint(fp, T("Οι τιμές γράφονται σε σταθερές θέσεις πάνω στο άδειο PDF. Έτσι δούλευε πάντα το πρόγραμμα.",
                   "Values are written at fixed positions on top of the blank PDF. This is how the program has always worked."),
             1, padx=(24, 0), pady=(2, 0), margin=130)
        ttk.Radiobutton(fp, text=T("Έξυπνη διάταξη (αναδημιουργία από το template)  ·  ΠΕΙΡΑΜΑΤΙΚΟ",
                                   "Smart layout (rebuilt from the template)  ·  EXPERIMENTAL"),
                        value='smart', variable=pv['pdf_mode'], style='Card.TRadiobutton',
                        command=self._pdf_mode_clicked).grid(row=2, column=0, sticky="w", pady=(10, 0))
        hint(fp, T("Πριν συμπληρωθεί το δελτίο, το πρόγραμμα διαβάζει το template, βρίσκει πού βρίσκεται τώρα κάθε τμήμα του "
                   "και τοποθετεί κάθε τιμή ακριβώς εκεί που ανήκει. Χρήσιμο αν η ΕΛΣΤΑΤ αλλάξει λίγο το έντυπο "
                   "(μετατοπισμένες γραμμές, περιθώρια, αποστάσεις). Αν το template δεν αναγνωριστεί με ασφάλεια, "
                   "χρησιμοποιείται αυτόματα η προεπιλεγμένη μέθοδος και εμφανίζεται προειδοποίηση.",
                   "Before the form is filled in, the program reads the template, finds where each part of it now sits "
                   "and places every value exactly where it belongs. Useful if ELSTAT changes the form slightly "
                   "(shifted lines, margins, spacing). If the template is not recognised safely, the default method "
                   "is used automatically and a warning is shown."),
             3, padx=(24, 0), pady=(2, 0), margin=130)

        # everything below is shown only while "Smart layout" is selected
        self.smart_frame = sm = ttk.Frame(fp, style='Card.TFrame')
        sm.grid(row=4, column=0, sticky="ew")
        sm.columnconfigure(0, weight=1)
        ttk.Separator(sm, orient="horizontal").grid(row=0, column=0, sticky="ew", pady=14)
        warn = ttk.Label(sm, text=T("⚠ Πειραματική λειτουργία: ελέγχετε πάντα το δελτίο πριν το υποβάλετε. Αν κάτι δεν "
                                    "πάει καλά, γυρίστε στο «Πάνω στο template».",
                                    "⚠ Experimental feature: always check the form before submitting it. If anything "
                                    "looks wrong, switch back to “On top of the template”."),
                         style='CardWarn.TLabel', justify="left")
        warn.grid(row=1, column=0, sticky="w")
        autowrap(warn, body, margin=110)
        self.layout_note = ttk.Label(sm, textvariable=self.layout_note_var, style='CardHint.TLabel', justify="left")
        self.layout_note.grid(row=2, column=0, sticky="w", pady=(10, 0))
        autowrap(self.layout_note, body, margin=110)
        lrow = ttk.Frame(sm, style='Card.TFrame')
        lrow.grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Button(lrow, text=T("Έλεγχος template", "Check template"), command=self.layout_check).pack(side="left")
        ttk.Button(lrow, text=T("Αποθήκευση ως αναφορά…", "Save as reference…"),
                   command=self.layout_save_reference).pack(side="left", padx=8)
        ttk.Button(lrow, text=T("Δοκιμαστικό δελτίο", "Test form"), command=self.layout_test_form).pack(side="left")
        hint(sm, T("Το «Δοκιμαστικό δελτίο» φτιάχνει ένα PDF με πλασματικά στοιχεία (μην το υποβάλετε) ώστε να δείτε αν όλα "
                   "πέφτουν στη θέση τους με τη μέθοδο που έχετε διαλέξει.",
                   "“Test form” creates a PDF with made-up data (do not submit it) so you can see whether everything "
                   "lands in place with the method you chose."), 4, pady=(8, 0))

        # ---- set manually (layout editor)
        ttk.Radiobutton(fp, text=T("Χειροκίνητη τοποθέτηση (επεξεργαστής διάταξης)", "Set manually (layout editor)"),
                        value='manual', variable=pv['pdf_mode'], style='Card.TRadiobutton',
                        command=self._pdf_mode_clicked).grid(row=5, column=0, sticky="w", pady=(10, 0))
        hint(fp, T("Ανοίγει ένας επεξεργαστής όπου βλέπετε το έντυπο και σέρνετε κάθε στοιχείο εκεί που ανήκει. "
                   "Για κάθε τιμή ορίζετε επίσης γραμματοσειρά, μέγεθος, χρώμα και στοίχιση. "
                   "Χρήσιμο όταν θέλετε πλήρη έλεγχο ή το έντυπο της ΕΛΣΤΑΤ αλλάξει.",
                   "Opens an editor where you see the form and drag each item to where it belongs. "
                   "For every value you also choose the font, size, colour and alignment. "
                   "Useful when you want full control or the ELSTAT form changes."),
             6, padx=(24, 0), pady=(2, 0), margin=130)
        self.manual_note_var = tk.StringVar()
        self.manual_frame = mf = ttk.Frame(fp, style='Card.TFrame')
        mf.grid(row=7, column=0, sticky="ew")
        mf.columnconfigure(0, weight=1)
        ttk.Separator(mf, orient="horizontal").grid(row=0, column=0, sticky="ew", pady=14)
        self.manual_note = ttk.Label(mf, textvariable=self.manual_note_var, style='CardHint.TLabel', justify="left")
        self.manual_note.grid(row=1, column=0, sticky="w")
        autowrap(self.manual_note, body, margin=110)
        mrow = ttk.Frame(mf, style='Card.TFrame')
        mrow.grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Button(mrow, text=T("Άνοιγμα επεξεργαστή διάταξης", "Open layout editor"), style='PrintSm.TButton',
                   command=self.open_layout_editor).pack(side="left")
        ttk.Button(mrow, text=T("Δοκιμαστικό δελτίο", "Test form"), command=self.layout_test_form).pack(side="left", padx=8)
        ttk.Button(mrow, text=T("Επαναφορά θέσεων", "Reset positions"), command=self.manual_reset).pack(side="left")
        self._update_smart_panel()

        # ---- 4. behaviour
        c4 = Card(body, None, T("Συμπεριφορά του προγράμματος", "Program behaviour"))
        c4.grid(row=5, column=0, sticky="ew", padx=4, pady=6)
        polish_settings_card(c4)
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
            ('preview_before_result', T("Προεπισκόπηση PDF πριν από το παράθυρο «Αποτέλεσμα»",
                                        "Preview the PDF before the “Result” window"),
             T("Εμφανίζει το δελτίο μέσα στο πρόγραμμα. Πατήστε «Συνέχεια στο αποτέλεσμα» για να δείτε "
               "τους ελέγχους και τα σύνολα.",
               "Shows the form inside the application. Press “Continue to result” to see the checks and totals.")),
        ]
        for r, (key, text, note) in enumerate(behaviours):
            ttk.Checkbutton(f4, text=text, variable=pv[key], style='Card.TCheckbutton',
                            command=self._pref_changed).grid(row=r * 2, column=0, sticky="w",
                                                             pady=(8 if r else 0, 0))
            if note:
                hint(f4, note, r * 2 + 1, padx=(24, 0), pady=(2, 0), margin=130)

        # ---- reset
        bottom = ttk.Frame(body)
        bottom.grid(row=6, column=0, sticky="ew", padx=4, pady=(6, 16))
        ttk.Button(bottom, text=T("Επαναφορά ρυθμίσεων", "Reset settings"), style='Danger.TButton',
                   command=self.reset_prefs).pack(side="left")
        ttk.Label(bottom, text=T("Δεν διαγράφει ιστορικό ή στοιχεία ξενοδοχείου.",
                                 "Does not delete history or hotel details."),
                  style='Hint.TLabel').pack(side="left", padx=10)

    # ---- manual layout (settings)
    def _update_manual_panel(self):
        frame = getattr(self, 'manual_frame', None)
        if frame is None:
            return
        try:
            if self.prefs.get('pdf_mode') == 'manual':
                frame.grid()
            else:
                frame.grid_remove()
            if manual_layout_saved():
                note = T("✓ Υπάρχει αποθηκευμένη δική σας διάταξη. Πατήστε «Άνοιγμα επεξεργαστή διάταξης» για αλλαγές.",
                         "✓ Your own layout is saved. Press “Open layout editor” to change it.")
            else:
                note = T("⚠ Δεν έχετε αποθηκεύσει ακόμη δική σας διάταξη - χρησιμοποιούνται οι προεπιλεγμένες θέσεις. "
                         "Ανοίξτε τον επεξεργαστή για να τις προσαρμόσετε.",
                         "⚠ You have not saved your own layout yet - the default positions are used. "
                         "Open the editor to adjust them.")
            self.manual_note_var.set(note)
            self.manual_note.configure(style='CardHint.TLabel' if manual_layout_saved() else 'CardWarn.TLabel')
        except tk.TclError:
            pass

    def open_layout_editor(self):
        title = T("Επεξεργαστής διάταξης", "Layout editor")
        existing = getattr(self, '_layout_editor', None)
        try:
            if existing is not None and existing.winfo_exists():
                existing.lift()
                return
        except tk.TclError:
            pass
        template = self._template_or_error(title)
        if not template:
            return
        try:
            if len(pypdf.PdfReader(template).pages) < 2:
                raise ValueError(T("Το template PDF πρέπει να έχει 2 σελίδες.", "The template PDF must have 2 pages."))
        except Exception as e:
            messagebox.showerror(title, T(f"Δεν μπορώ να διαβάσω το template:\n{e}", f"I could not read the template:\n{e}"))
            return
        self._layout_editor = LayoutEditor(self, template, on_save=self._update_manual_panel)

    def manual_reset(self):
        title = T("Επαναφορά θέσεων", "Reset positions")
        if not manual_layout_saved():
            messagebox.showinfo(title, T("Δεν υπάρχει αποθηκευμένη δική σας διάταξη - χρησιμοποιούνται ήδη οι προεπιλεγμένες θέσεις.",
                                         "There is no saved layout of yours - the default positions are already in use."))
            return
        if messagebox.askyesno(title, T("Η αποθηκευμένη διάταξη θα διαγραφεί και θα ξαναχρησιμοποιηθούν οι προεπιλεγμένες θέσεις.\n\nΣυνέχεια;",
                                        "The saved layout will be deleted and the default positions will be used again.\n\nContinue?"),
                               default='no'):
            try:
                delete_manual_layout()
            except OSError as e:
                messagebox.showerror(title, str(e))
                return
            self._update_manual_panel()
            self._status(T("✓ Οι προεπιλεγμένες θέσεις επανήλθαν.", "✓ The default positions are back."))

    # ---- smart layout (settings)
    def _template_path_now(self):
        return self.pdf_var.get().strip() or DEFAULT_TEMPLATE

    def _update_smart_panel(self):
        """Shows the smart-layout options only while that method is selected."""
        self._update_manual_panel()
        frame = getattr(self, 'smart_frame', None)
        if frame is None:
            return
        if self.prefs.get('pdf_mode') == 'smart':
            frame.grid()
        else:
            frame.grid_remove()
        self._update_layout_note()

    def _update_layout_note(self):
        var = getattr(self, 'layout_note_var', None)
        if var is None:
            return
        ref = load_layout_reference()
        if ref is None:
            var.set(T("⚠ Δεν υπάρχει ακόμη σημείο σύγκρισης (διάταξη αναφοράς). Πατήστε «Αποθήκευση ως αναφορά…» όταν το "
                      "δελτίο συμπληρώνεται σωστά με το τρέχον template.",
                      "⚠ There is no point of comparison (reference layout) yet. Press “Save as reference…” when the "
                      "form fills in correctly with the current template."))
            style = 'CardWarn.TLabel'
        else:
            try:
                when = datetime.datetime.fromisoformat(ref.get('saved', '')).strftime('%d/%m/%Y')
            except ValueError:
                when = "-"
            name = ref.get('template', '')
            var.set(T(f"✓ Διάταξη αναφοράς: {name} (αποθηκεύτηκε {when}).", f"✓ Reference layout: {name} (saved {when})."))
            style = 'CardGood.TLabel'
        try:
            self.layout_note.configure(style=style)
        except (AttributeError, tk.TclError):
            pass

    def _template_or_error(self, title):
        template = self._template_path_now()
        if not os.path.isfile(template):
            messagebox.showerror(title, T("Το άδειο PDF (template) δεν βρέθηκε - δείτε την καρτέλα «Στοιχεία ξενοδοχείου».",
                                          "The blank PDF (template) was not found - see the “Hotel details” tab."))
            return None
        return template

    def _pdf_mode_clicked(self):
        """A radio button of the PDF method was clicked."""
        if self.pv['pdf_mode'].get() == 'manual':
            # Selecting manual mode only reveals its controls; the editor opens
            # only when the user presses the dedicated button.
            self._pref_changed()
            return
        if self.pv['pdf_mode'].get() == 'smart' and not self.prefs.get('smart_ack'):
            if not messagebox.askyesno(
                    T("Πειραματική λειτουργία", "Experimental feature"),
                    T("Η «Έξυπνη διάταξη» είναι πειραματική.\n\n"
                      "Προσπαθεί να βρει μόνη της πού πρέπει να γραφτεί κάθε τιμή αν το έντυπο της ΕΛΣΤΑΤ αλλάξει λίγο. "
                      "Μπορεί να κάνει λάθος, γι' αυτό ελέγχετε πάντα το δελτίο πριν το υποβάλετε. Η προεπιλεγμένη "
                      "μέθοδος παραμένει διαθέσιμη ανά πάσα στιγμή.\n\nΝα ενεργοποιηθεί;",
                      "“Smart layout” is experimental.\n\n"
                      "It tries to work out by itself where each value should be written if the ELSTAT form changes "
                      "slightly. It can get things wrong, so always check the form before submitting it. The default "
                      "method stays available at any time.\n\nTurn it on?"), default='no'):
                self.pv['pdf_mode'].set('overlay')
                self._pref_changed()
                return
            self.prefs['smart_ack'] = True
        self._pref_changed()
        if self.prefs['pdf_mode'] == 'smart' and load_layout_reference() is None:
            self._offer_reference(T("Έξυπνη διάταξη", "Smart layout"))

    def _save_reference(self, template, title):
        try:
            save_layout_reference(template)
        except PermissionError:
            messagebox.showerror(title, T("Δεν μπορώ να γράψω στον φάκελο δεδομένων. Ελέγξτε ότι δεν είναι μόνο για ανάγνωση.",
                                          "I can't write to the data folder. Check that it is not read-only."))
            return False
        except Exception as e:
            messagebox.showerror(title, T(f"Δεν ήταν δυνατή η αποθήκευση:\n{e}", f"Could not save:\n{e}"))
            return False
        self._update_layout_note()
        self._status(T("✓ Το template αποθηκεύτηκε ως διάταξη αναφοράς.", "✓ The template was saved as the reference layout."))
        return True

    def _offer_reference(self, title):
        """There is no reference layout yet: explain and offer to save the current template as one."""
        template = self._template_or_error(title)
        if not template:
            return False
        if messagebox.askyesno(
                title,
                T("Η έξυπνη διάταξη χρειάζεται ένα σημείο σύγκρισης: το template όπως είναι όταν το δελτίο "
                  "συμπληρώνεται σωστά.\n\nΝα αποθηκευτεί το τρέχον template ως σημείο σύγκρισης;\n"
                  "(Πατήστε «Ναι» μόνο αν με την προεπιλεγμένη μέθοδο το δελτίο βγαίνει σωστό.)",
                  "Smart layout needs a point of comparison: the template as it is when the form fills in "
                  "correctly.\n\nSave the current template as the point of comparison?\n"
                  "(Press “Yes” only if the form comes out right with the default method.)"), default='no'):
            return self._save_reference(template, title)
        return False

    def layout_check(self):
        title = T("Έλεγχος template", "Check template")
        template = self._template_or_error(title)
        if not template:
            return
        if load_layout_reference() is None:
            self._offer_reference(title)
            return
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            _layout, report = analyze_template(template)
        except Exception as e:
            messagebox.showerror(title, T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"))
            return
        finally:
            self.config(cursor='')
        (messagebox.showinfo if report['ok'] else messagebox.showwarning)(title, describe_report(report))

    def layout_save_reference(self):
        title = T("Αποθήκευση ως αναφορά", "Save as reference")
        template = self._template_or_error(title)
        if not template:
            return
        if not messagebox.askyesno(
                title,
                T("Να χρησιμοποιηθεί το τρέχον template ως διάταξη αναφοράς;\n\n"
                  "Κάντε το μόνο αν το δελτίο συμπληρώνεται σωστά με την προεπιλεγμένη μέθοδο πάνω σε αυτό το template. "
                  "Η έξυπνη διάταξη θα μετράει τις αλλαγές των επόμενων templates σε σχέση με αυτό.",
                  "Use the current template as the reference layout?\n\n"
                  "Only do this if the form is filled in correctly with the default method on this template. "
                  "Smart layout will measure the changes of later templates against it."), default='no'):
            return
        self._save_reference(template, title)

    def layout_test_form(self):
        title = T("Δοκιμαστικό δελτίο", "Test form")
        template = self._template_or_error(title)
        if not template:
            return
        if self.pv['pdf_mode'].get() == 'smart' and load_layout_reference() is None:
            if not self._offer_reference(title):
                return
        sample = {'type': 'TEST', 'name': 'TEST HOTEL', 'category': '5*', 'email': 'test@example.com',
                  'phone': '2310000000', 'afm': '123456789', 'company': 'TEST COMPANY'}
        today = datetime.date.today()
        last = calendar.monthrange(today.year, today.month)[1]
        mon = dict(year=str(today.year), month=str(today.month), date_to=f"{last:02d}/{today.month:02d}/{today.year}",
                   days_of_month=str(last), employees='25', sign_date=today.strftime('%d/%m/%Y'))
        rows = [(name, 11, 22) for name in sorted(MAPPED_TARGETS)]
        path = os.path.join(tempfile.gettempdir(), "ELSTAT_test_form.pdf")
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            result = process_elstat_report(None, template, path, sample, mon, rows_override=rows,
                                           save_history=False, pdf_mode=self.pv['pdf_mode'].get())
        except PermissionError:
            messagebox.showerror(title, T("Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό και δοκιμάστε ξανά.",
                                          "I can't write the PDF. Close it if it is open and try again."))
            return
        except Exception as e:
            messagebox.showerror(title, T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"))
            return
        finally:
            self.config(cursor='')
        if result.get('layout_warning'):
            messagebox.showwarning(title, result['layout_warning'])
        open_file(path)

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
        p['pdf_mode'] = pv['pdf_mode'].get()
        pattern = pv['filename_pattern'].get().strip()
        if pattern_error(pattern) is None:
            p['filename_pattern'] = pattern
        for k in BOOL_PREFS:
            p[k] = bool(pv[k].get())
        self._save_prefs()
        self._after_prefs_change()

    # ---- user guide
    def goto_tab(self, name):
        """Opens one of the tabs by name: report, hotel, history, settings."""
        try:
            self.notebook.select(getattr(self, 'tab_' + name))
        except (tk.TclError, AttributeError):
            pass

    def open_help(self, section=None):
        """Opens the guide - by default at the section that belongs to the tab you are on."""
        try:
            cur = self.notebook.select()
        except tk.TclError:
            return
        if section is None and cur != str(self.tab_help):
            section = {str(self.tab_report): 'monthly', str(self.tab_hotel): 'hotel',
                       str(self.tab_history): 'history', str(self.tab_settings): 'settings'}.get(cur)
        if cur != str(self.tab_help):
            self._help_prev = cur
        self.notebook.select(self.tab_help)
        if section:
            self.guide.show(section)
        self.guide.refresh_back()

    def help_context(self):
        """Live values for the [[name]] placeholders of the guide."""
        hist = self._hist or {}
        last = max(((y, m) for y, ms in hist.items() for m in ms), default=None)
        modes = {'overlay': T("Πάνω στο template", "On top of the template"),
                 'smart': T("Έξυπνη διάταξη (πειραματική)", "Smart layout (experimental)"),
                 'manual': T("Χειροκίνητη τοποθέτηση", "Set manually")}
        themes = {'light': T("Φωτεινή", "Light"), 'dark': T("Σκούρα", "Dark"), 'auto': T("Όπως το σύστημα", "Match system")}
        return dict(
            data_dir=DATA_DIR,
            next_pdf=self.out_var.get().strip() or self._suggest_output(),
            pdf_mode_name=modes.get(self.prefs.get('pdf_mode'), ''),
            language_name=LANGUAGES.get(LANG, LANG),
            theme_name=themes.get(self.prefs.get('theme'), ''),
            months_count=sum(len(ms) for ms in hist.values()),
            last_month=f"{month_name(last[1])} {last[0]}" if last else T("κανένας ακόμη", "none yet"),
            custom_count=len(CUSTOM_COUNTRY_MAPPING),
            template_name=os.path.basename(self._template_path_now()),
        )

    def _unknown_country_count(self):
        """How many countries of the loaded Excel are not known to the program."""
        excel = self.excel_var.get().strip()
        if not (excel and os.path.isfile(excel)):
            return 0
        try:
            return sum('bad' in self.preview_tree.item(i, 'tags') for i in self.preview_tree.get_children())
        except tk.TclError:
            return 0

    def help_status(self):
        """[(state, label, detail, target)] - what the 'Live status' page of the guide shows."""
        s = {k: v.get().strip() for k, v in self.set_vars.items()}
        hist = self._hist or {}
        out = []
        errs = validate_hotel(s)
        out.append(('bad' if errs else 'ok', T("Στοιχεία ξενοδοχείου", "Hotel details"),
                    T(f"{len(errs)} πεδία θέλουν διόρθωση", f"{len(errs)} field(s) need fixing") if errs
                    else T("συμπληρωμένα", "complete"), 'tab:hotel'))
        tpl = self._template_path_now()
        tpl_ok = os.path.isfile(tpl)
        out.append(('ok' if tpl_ok else 'bad', T("Άδειο PDF (template)", "Blank PDF (template)"),
                    os.path.basename(tpl) if tpl_ok else T("δεν βρέθηκε", "not found"), 'tab:hotel'))
        months = sum(len(ms) for ms in hist.values())
        out.append(('ok' if months else 'warn', T("Ιστορικό", "History"),
                    T(f"{months} μήνες καταχωρημένοι", f"{months} months recorded") if months
                    else T("κανένας μήνας ακόμη", "no month yet"), 'tab:history'))
        rem = missing_month_reminder(hist)
        if rem:
            out.append(('warn', T("Εκκρεμές δελτίο", "Pending form"),
                        T(f"λείπει ο {month_name(rem[1])} {rem[0]}", f"{month_name(rem[1])} {rem[0]} is missing"), 'tab:report'))
        excel = self.excel_var.get().strip()
        excel_ok = bool(excel) and os.path.isfile(excel)
        out.append(('ok' if excel_ok else 'warn', T("Excel του μήνα", "Month's Excel"),
                    os.path.basename(excel) if excel_ok else T("δεν έχει επιλεγεί ακόμη", "not chosen yet"), 'tab:report'))
        bad = self._unknown_country_count()
        if bad:
            out.append(('bad', T("Άγνωστες χώρες", "Unknown countries"),
                        T(f"{bad} δεν θα μπουν στο PDF", f"{bad} will not be in the PDF"), 'sec:countries'))
        mode = self.prefs.get('pdf_mode')
        if mode == 'manual':
            saved = manual_layout_saved()
            out.append(('ok' if saved else 'warn', T("Χειροκίνητη διάταξη", "Manual layout"),
                        T("αποθηκευμένη", "saved") if saved
                        else T("δεν έχει αποθηκευτεί - ισχύουν οι προεπιλεγμένες θέσεις", "not saved - default positions apply"),
                        'sec:editor'))
        elif mode == 'smart':
            ref = load_layout_reference() is not None
            out.append(('ok' if ref else 'warn', T("Έξυπνη διάταξη", "Smart layout"),
                        T("υπάρχει διάταξη αναφοράς", "reference layout exists") if ref
                        else T("λείπει η διάταξη αναφοράς", "no reference layout yet"), 'tab:settings'))
        out.append(('ok' if DND_AVAILABLE else 'warn', 'Drag & drop',
                    T("ενεργό", "active") if DND_AVAILABLE else T("μη διαθέσιμο (pip install tkinterdnd2)", "not available (pip install tkinterdnd2)"),
                    'sec:dnd'))
        return out

    def help_next(self):
        """(text, target) - the one thing the person most likely wants to do next."""
        s = {k: v.get().strip() for k, v in self.set_vars.items()}
        if validate_hotel(s):
            return T("Συμπληρώστε τα στοιχεία του ξενοδοχείου", "Fill in the hotel details"), 'tab:hotel'
        if not os.path.isfile(self._template_path_now()):
            return T("Βρείτε το άδειο PDF (template)", "Find the blank PDF (template)"), 'tab:hotel'
        excel = self.excel_var.get().strip()
        if not (excel and os.path.isfile(excel)):
            rem = missing_month_reminder(self._hist or {})
            if rem:
                return (T(f"Φτιάξτε το δελτίο {month_name(rem[1])} {rem[0]}: διαλέξτε το Excel του",
                          f"Make the form for {month_name(rem[1])} {rem[0]}: choose its Excel"), 'tab:report')
            return T("Διαλέξτε το Excel του μήνα", "Choose the month's Excel"), 'tab:report'
        if self._unknown_country_count():
            return T("Αντιστοιχίστε τις άγνωστες χώρες", "Map the unknown countries"), 'sec:countries'
        if not self.mon_vars['employees'].get().strip():
            return T("Πληκτρολογήστε το σύνολο εργαζομένων", "Type the total employees"), 'tab:report'
        return T("Όλα έτοιμα - δημιουργήστε το PDF", "Everything is ready - create the PDF"), 'tab:report'

    def help_action(self, key):
        if key == 'open_data':
            open_file(DATA_DIR)
        elif key == 'choose_excel':
            self.goto_tab('report')
            self.browse_excel()
        elif key == 'open_editor':
            self.goto_tab('settings')
            self.open_layout_editor()
        elif key == 'backup':
            self.backup_data()
        elif key == 'theme':
            self._toggle_theme()

    def _help_find(self, _event=None):
        """Ctrl+F: jump to the search box of the guide when the guide is open."""
        try:
            if self.notebook.select() == str(self.tab_help):
                self.guide.focus_search()
                return 'break'
        except (tk.TclError, AttributeError):
            pass

    def _tab_changed(self, _event=None):
        if not hasattr(self, 'guide'):
            return
        try:
            cur = self.notebook.select()
        except tk.TclError:
            return
        if cur == str(self.tab_help):
            if self._last_tab and self._last_tab != cur:
                self._help_prev = self._last_tab
            self.guide.refresh_back()
            self.guide.refresh_live()
        else:
            self._last_tab = cur

    def _rebuild_screen(self):
        """Builds the whole screen again (language / theme change). Typed-in values are kept."""
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
            self._register_drag_drop_targets()
            self.notebook.select(tab)
        finally:
            self._rebuilding = False
        self._after_prefs_change()
        self.refresh_validation()
        self._render_period()
        if self._loaded_excel and os.path.isfile(self._loaded_excel):
            self.load_preview(self._loaded_excel)

    def _theme_clicked(self):
        code = self.pv['theme'].get()
        if code == self.prefs['theme']:
            return
        self.prefs['theme'] = code
        self._save_prefs()
        self.after_idle(self._apply_theme)

    def _toggle_theme(self):
        """The sun / moon button in the header: switches between light and dark."""
        self.prefs['theme'] = 'light' if THEME == 'dark' else 'dark'
        self._save_prefs()
        self.after_idle(self._apply_theme)

    def _apply_theme(self):
        set_theme(self.prefs['theme'])
        self.pv['theme'].set(self.prefs['theme'])
        setup_styles(self)
        self.configure(bg=BG)
        self._rebuild_screen()
        self._apply_titlebar()
        self._status(T("Η εμφάνιση άλλαξε σε σκούρα." if THEME == 'dark' else "Η εμφάνιση άλλαξε σε φωτεινή.",
                       "Switched to the dark appearance." if THEME == 'dark' else "Switched to the light appearance."))

    def _apply_titlebar(self):
        """Windows only: dark title bar when the dark theme is on."""
        if not sys.platform.startswith('win'):
            return
        try:
            import ctypes
            self.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            val = ctypes.c_int(1 if THEME == 'dark' else 0)
            for attr in (20, 19):          # 20 = Windows 11 / newer 10, 19 = older Windows 10
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(val), ctypes.sizeof(val)) == 0:
                    break
        except Exception:
            pass

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
        self._rebuild_screen()
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
        self._update_smart_panel()

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
        load_custom_country_mappings()
        if not save_app_config({'data_dir': '' if _same_path(new, DEFAULT_DATA_DIR) else new}):
            set_data_dir(old)
            load_custom_country_mappings()
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
        """Scroll the area under the pointer, including embedded Help callouts."""
        try:
            w = self.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return
        if isinstance(w, ttk.Treeview):
            return

        # Windows reports +/-120 (or multiples); Linux uses Button-4/5.
        num = getattr(event, 'num', None)
        if num == 4:
            units = -3
        elif num == 5:
            units = 3
        else:
            delta = getattr(event, 'delta', 0)
            if not delta:
                return
            units = -max(1, int(abs(delta) / 120)) * (1 if delta > 0 else -1) * 3

        # A Help callout is a child Frame/Label embedded inside a Text widget.
        # Walk its master chain so the wheel still scrolls the Help text itself.
        while w is not None:
            if isinstance(w, tk.Text):
                try:
                    first, last = w.yview()
                    if last - first < 0.999:
                        w.yview_scroll(units, 'units')
                    return 'break'
                except tk.TclError:
                    return 'break'

            sf = getattr(w, '_sf', None)
            if sf is not None:
                sf.scroll(units)
                return 'break'
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
        try:
            datetime.datetime.strptime(m['sign_date'], "%d/%m/%Y")
            st['sign_date'] = True
        except ValueError:
            st['sign_date'] = False
        return st

    def _update_dropzone(self):
        zone = getattr(self, 'excel_zone', None)
        if zone is None:
            return
        fp = self.excel_var.get().strip()
        if fp and os.path.isfile(fp):
            zone.set_text("\u2713  " + os.path.basename(fp),
                          T("Πατήστε ή σύρετε εδώ άλλο αρχείο για να το αλλάξετε",
                            "Click or drop another file here to change it"), done=True)
        elif DND_AVAILABLE:
            zone.set_text(T("Σύρετε εδώ το αρχείο Excel", "Drop the Excel file here"),
                          T("ή πατήστε για να το διαλέξετε από τον υπολογιστή",
                            "or click to choose it from your computer"))
        else:
            zone.set_text(T("Πατήστε για να διαλέξετε το αρχείο Excel", "Click to choose the Excel file"),
                          T("το αρχείο «Nationalities Statistics» από το Protel (.xls ή .xlsx)",
                            "the “Nationalities Statistics” file from Protel (.xls or .xlsx)"))

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

        # drop zone text and the step badges (number -> green tick when that step is done)
        self._update_dropzone()
        states = self._field_states(s, m)
        self.excel_card.set_done(excel_ok)
        self.month_card.set_done(excel_ok and all(ok for k, ok in states.items() if k in m))
        self.output_card.set_done(bool(out))

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

    # ----------------------------------------------------------- drag and drop
    def _register_drag_drop_targets(self):
        """Enable OS file/folder drops on the matching path fields when tkinterdnd2 is installed."""
        if not DND_AVAILABLE:
            return
        targets = (
            (getattr(self, 'excel_zone', None), self._drop_excel),
            (getattr(self, 'excel_card', None), self._drop_excel),
            (getattr(self, 'excel_path_entry', None), self._drop_excel),
            (getattr(self, 'template_card', None), self._drop_template),
            (getattr(self, 'output_card', None), self._drop_output),
            (getattr(self, 'template_path_entry', None), self._drop_template),
            (getattr(self, 'output_path_entry', None), self._drop_output),
            (getattr(self, 'out_dir_entry', None), self._drop_output_folder),
            (getattr(self, 'excel_dir_entry', None), self._drop_excel_folder),
            (getattr(self, 'data_dir_entry', None), self._drop_data_folder),
        )
        failed = []
        for widget, callback in targets:
            if widget is None:
                continue
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', callback)
            except Exception as e:
                # A failed target registration must not affect the rest of the interface.
                failed.append(repr(e))
        zone = getattr(self, 'excel_zone', None)
        if zone is not None:
            try:
                zone.dnd_bind('<<DropEnter>>', lambda e: (zone.set_drag(True), e.action)[1])
                zone.dnd_bind('<<DropLeave>>', lambda e: (zone.set_drag(False), e.action)[1])
            except Exception:
                pass
        if failed:
            self._status("Drag & drop registration failed: " + failed[0])
        else:
            self._status(T("Drag & drop ενεργό: αφήστε αρχεία πάνω στα πεδία.",
                           "Drag & drop is active: drop files onto the fields."))

    def _dropped_paths(self, event):
        """Parse Tk's drop payload correctly, including paths containing spaces."""
        try:
            return [os.path.normpath(p) for p in self.tk.splitlist(event.data) if p]
        except (tk.TclError, AttributeError, TypeError):
            raw = str(getattr(event, 'data', '')).strip()
            return [raw.strip('{}')] if raw else []

    def _drop_one(self, event, kind, extensions, callback):
        paths = self._dropped_paths(event)
        if len(paths) != 1:
            messagebox.showwarning(
                T("Μεταφορά αρχείου", "Drag and drop"),
                T("Αφήστε ένα αρχείο κάθε φορά.", "Please drop one file at a time."),
                parent=self)
            return event.action
        path = paths[0]
        if not os.path.isfile(path) or os.path.splitext(path)[1].lower() not in extensions:
            messagebox.showwarning(
                T("Μη έγκυρο αρχείο", "Invalid file"),
                T(f"Αφήστε ένα υπάρχον αρχείο {kind}.", f"Drop an existing {kind} file."),
                parent=self)
            return event.action
        callback(path)
        return event.action

    def _drop_excel(self, event):
        zone = getattr(self, 'excel_zone', None)
        if zone is not None:
            zone.set_drag(False)
        return self._drop_one(event, "Excel (.xls/.xlsx)", {'.xls', '.xlsx'}, self._accept_dropped_excel)

    def _accept_dropped_excel(self, path):
        self.excel_var.set(path)
        self.load_period(path)
        self._status(T(f"✓ Επιλέχθηκε το Excel: {path}", f"✓ Excel selected: {path}"))

    def _drop_template(self, event):
        return self._drop_one(event, "PDF", {'.pdf'}, self._accept_dropped_template)

    def _accept_dropped_template(self, path):
        self.pdf_var.set(path)
        self._remember_template(path)
        self._status(T(f"✓ Επιλέχθηκε το template: {path}", f"✓ Template selected: {path}"))

    def _drop_output(self, event):
        return self._drop_one(event, "PDF", {'.pdf'}, self._accept_dropped_output)

    def _accept_dropped_output(self, path):
        self.out_var.set(path)
        self.output_custom = True
        self._status(T(f"✓ Ορίστηκε το αρχείο εξόδου: {path}", f"✓ Output file set: {path}"))

    def _drop_folder(self, event, callback):
        paths = self._dropped_paths(event)
        if len(paths) != 1 or not os.path.isdir(paths[0]):
            messagebox.showwarning(
                T("Μη έγκυρος φάκελος", "Invalid folder"),
                T("Αφήστε έναν υπάρχοντα φάκελο.", "Drop one existing folder."),
                parent=self)
            return event.action
        callback(paths[0])
        return event.action

    def _drop_output_folder(self, event):
        return self._drop_folder(event, self._accept_output_folder)

    def _accept_output_folder(self, path):
        self.pv['output_dir'].set(path)
        self.pv['output_mode'].set('custom')
        self._pref_changed()
        self._status(T(f"✓ Φάκελος PDF: {path}", f"✓ PDF folder: {path}"))

    def _drop_excel_folder(self, event):
        return self._drop_folder(event, self._accept_excel_folder)

    def _accept_excel_folder(self, path):
        self.pv['excel_dir'].set(path)
        self._pref_changed()
        self._status(T(f"✓ Φάκελος Excel: {path}", f"✓ Excel folder: {path}"))

    def _drop_data_folder(self, event):
        return self._drop_folder(event, self._switch_data_dir)

    # ----------------------------------------------------------- file choices
    def _start_dir(self):
        d = self.current_settings.get('last_dir', '')
        return d if d and os.path.isdir(d) else BASE_DIR

    def browse_excel(self):
        fp = filedialog.askopenfilename(title=T("Διαλέξτε το Excel από το Protel",
                                                "Choose the Excel file from Protel"), initialdir=self._excel_start_dir(),
                                        filetypes=[("Excel Files", "*.xls *.xlsx")])
        if fp:
            self._accept_dropped_excel(fp)

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

    def _update_month_field_status(self, *_):
        """Refresh the compact completion marks in the Month details card."""
        widgets = getattr(self, 'month_field_status', {})
        for key, marker in widgets.items():
            filled = bool(self.mon_vars.get(key) and self.mon_vars[key].get().strip())
            marker.configure(text='✓' if filled else '○',
                             style='CardGood.TLabel' if filled else 'CardHint.TLabel')

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

    def _map_selected_country(self, _event=None):
        """Let the user map an unknown Excel country to the correct ELSTAT form row."""
        if _event is not None and hasattr(_event, 'y'):
            clicked = self.preview_tree.identify_row(_event.y)
            if clicked:
                self.preview_tree.selection_set(clicked)
        selection = self.preview_tree.selection()
        if not selection:
            messagebox.showinfo(T("Επιλέξτε χώρα", "Select a country"),
                                T("Επιλέξτε πρώτα μια κόκκινη χώρα από τον πίνακα.",
                                  "First select a red country in the table."), parent=self)
            return 'break'
        country = self.preview_tree.item(selection[0], 'text')
        existing_custom = CUSTOM_COUNTRY_MAPPING_NORM.get(norm(country))
        if resolve_country(country) is not None and existing_custom is None:
            messagebox.showinfo(T("Η χώρα είναι ήδη γνωστή", "Country already recognised"),
                                T(f"Η χώρα «{country}» αναγνωρίζεται ήδη από το πρόγραμμα.",
                                  f"“{country}” is already recognised by the program."), parent=self)
            return 'break'

        win = tk.Toplevel(self)
        win.title(T("Αντιστοίχιση χώρας", "Map country"))
        win.configure(bg=BG)
        win.transient(self)
        win.resizable(False, False)
        box = ttk.Frame(win, padding=18)
        box.pack(fill='both', expand=True)
        ttk.Label(box, text=T(f"Η χώρα «{country}» δεν είναι γνωστή.",
                              f"The country “{country}” is not recognised."),
                  style='CardBold.TLabel', wraplength=440).pack(anchor='w')
        ttk.Label(box, text=T("Σε ποια γραμμή / κατηγορία του δελτίου της ΕΛΣΤΑΤ ανήκει;",
                              "Which ELSTAT form row / category should it be assigned to?"),
                  style='CardHint.TLabel', wraplength=440).pack(anchor='w', pady=(8, 10))
        choices = country_mapping_destinations()
        choice_var = tk.StringVar(value='')
        combo = ttk.Combobox(box, textvariable=choice_var, values=choices,
                             state='readonly', width=48)
        combo.pack(fill='x', pady=(0, 12))
        if existing_custom:
            for i, option in enumerate(choices):
                if norm(option) == existing_custom:
                    combo.current(i)
                    break
        ttk.Label(box, text=T("Η επιλογή αποθηκεύεται και θα χρησιμοποιείται και στους επόμενους μήνες. "
                              "Αν δεν είστε βέβαιοι, ελέγξτε πρώτα τη σωστή κατηγορία.",
                              "Your choice is saved and will also be used in future months. "
                              "If unsure, verify the correct category first."),
                  style='CardHint.TLabel', wraplength=440).pack(anchor='w', pady=(0, 12))
        buttons = ttk.Frame(box)
        buttons.pack(fill='x')
        def save_mapping():
            destination = choice_var.get().strip()
            if not destination:
                messagebox.showwarning(T("Επιλέξτε κατηγορία", "Choose a category"),
                                        T("Επιλέξτε πού ανήκει η χώρα.",
                                          "Choose where the country belongs."), parent=win)
                return
            try:
                save_custom_country_mapping(country, destination)
            except Exception as exc:
                messagebox.showerror(T("Δεν αποθηκεύτηκε", "Could not save"), str(exc), parent=win)
                return
            win.destroy()
            fp = self.excel_var.get().strip()
            if fp and os.path.isfile(fp):
                self.load_preview(fp)
            messagebox.showinfo(T("Η αντιστοίχιση αποθηκεύτηκε", "Mapping saved"),
                                T(f"Η χώρα «{country}» αντιστοιχίστηκε στη γραμμή «{destination}». "
                                  "Πατήστε ξανά «Δημιουργία PDF» για να δημιουργηθεί το σωστό δελτίο.",
                                  f"“{country}” was mapped to “{destination}”. "
                                  "Click “Create PDF” again to generate the corrected form."), parent=self)
        ttk.Button(buttons, text=T("Αποθήκευση", "Save mapping"), command=save_mapping).pack(side='left')
        ttk.Button(buttons, text=T("Ακύρωση", "Cancel"), command=win.destroy).pack(side='right')
        win.bind('<Return>', lambda _e: save_mapping())
        win.bind('<Escape>', lambda _e: win.destroy())
        win.update_idletasks()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")
        win.grab_set()
        combo.focus_set()
        return 'break'

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
            destination = resolve_country(country)
            shown_destination = (next((v for v in country_mapping_destinations()
                                        if norm(v) == destination), destination)
                                  if destination else T("⚠ λείπει από το πρόγραμμα",
                                                        "⚠ missing from the program"))
            self.preview_tree.insert('', 'end', text=country, values=(pax, nights,
                                     shown_destination),
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
        self.hist_tree.tag_configure('year', font=(FONT, 10, 'bold'), background=ROW_YEAR)
        self.hist_tree.tag_configure('month', font=(FONT, 10, 'bold'), background=ROW_MONTH)
        self.hist_tree.tag_configure('country', background=CARD)
        self.hist_tree.tag_configure('hl', background=ROW_HL)
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
        self.top_tree.tag_configure('hl', background=ROW_HL)
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
        else:
            self.hist_tree.selection_remove(self.hist_tree.selection())

        # Keep the context menu in sync with the clicked row so disabled actions
        # cannot accidentally operate on a previously selected month.
        sel = self._hist_sel()
        month_ok = bool(sel and sel[1] is not None)
        year_ok = bool(self._hist)
        for index in (0, 1, 2, 7):
            self.hist_menu.entryconfigure(index, state='normal' if month_ok else 'disabled')
        for index in (4, 5):
            self.hist_menu.entryconfigure(index, state='normal' if year_ok else 'disabled')
        try:
            self.hist_menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.hist_menu.grab_release()
            except tk.TclError:
                pass
        return 'break'

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
        sign_date = self.mon_vars['sign_date'].get().strip() or datetime.date.today().strftime('%d/%m/%Y')
        while True:       # ask which signing date to print (today by default)
            sign_date = simpledialog.askstring(
                T("Ημερομηνία υπογραφής", "Signing date"),
                T(f"Ποια ημερομηνία υπογραφής να γραφτεί στο δελτίο του {name};\n(ΗΗ/ΜΜ/ΕΕΕΕ - προτείνεται η σημερινή)",
                  f"Which signing date should be printed on the form for {name}?\n"
                  "(DD/MM/YYYY - today is suggested)"),
                initialvalue=datetime.date.today().strftime('%d/%m/%Y') if not sign_date else sign_date,
                parent=self)
            if sign_date is None:
                return None
            sign_date = sign_date.strip()
            try:
                datetime.datetime.strptime(sign_date, "%d/%m/%Y")
                break
            except ValueError:
                messagebox.showerror(T("Λάθος ημερομηνία", "Invalid date"),
                                     T("Η ημερομηνία πρέπει να είναι της μορφής ΗΗ/ΜΜ/ΕΕΕΕ και υπαρκτή.",
                                       "The date must be in the format DD/MM/YYYY and must exist."))
        mon = dict(year=str(year), month=str(month), date_to=f"{last:02d}/{month:02d}/{year}",
                   days_of_month=str(rec.get('days') or last), employees=employees, sign_date=sign_date)
        problems = validate_month(mon)
        if problems:
            messagebox.showerror(T("Λάθος στοιχεία", "Invalid details"), "\n".join("• " + e for e in problems))
            return None
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            process_elstat_report(None, template, path, s, mon, rows_override=rows, save_history=False,
                                  pdf_mode=self.prefs.get('pdf_mode', 'overlay'))
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
            self._accept_dropped_template(fp)

    def browse_output(self):
        cur = self.out_var.get().strip()
        fp = filedialog.asksaveasfilename(title=T("Αποθήκευση PDF ως", "Save PDF as"), defaultextension=".pdf",
                                          initialdir=os.path.dirname(cur) or BASE_DIR,
                                          initialfile=os.path.basename(cur) or "ELSTAT_Report.pdf",
                                          filetypes=[("PDF Files", "*.pdf")])
        if fp:
            self._accept_dropped_output(fp)

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
            result = process_elstat_report(excel, template, out, s, m, pdf_mode=self.prefs.get('pdf_mode', 'overlay'))
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

        if result.get('layout_warning'):
            warnings_txt.append(result['layout_warning'])

        extra = ("  ·  " + result['layout_info']) if result.get('layout_info') else ""
        try:
            gy, gm = int(m['year']), int(m['month'])
            if self.prefs['keep_archive']:
                archive_copy(gy, gm, out)
            save_month_record(gy, gm, m['employees'], m['days_of_month'], m['date_to'])
            self._status(T(f"✓ Το δελτίο {month_name(gm)} {gy} δημιουργήθηκε και αρχειοθετήθηκε.",
                           f"✓ The form for {month_name(gm)} {gy} was created and archived.") + extra)
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
        show_result = lambda: ResultDialog(self, result, out, warnings_txt, print_status)
        if self.prefs.get('preview_before_result', False):
            PDFPreviewDialog(self, out, show_result)
        else:
            show_result()


if __name__ == "__main__":
    if sys.platform.startswith('win'):      # sharp text on high-DPI screens
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = ElstatApp()
    app.mainloop()
