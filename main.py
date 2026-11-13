APP_VERSION = "1.0.4"      # raise this number every time you push a new main.py to GitHub
FORCE_UPDATE = False       # set to True in the file you push to GitHub (together with a higher APP_VERSION): every
                           # program that starts online then installs this version by itself, without asking anyone.
                           # With False the update stays optional (the user presses "Install update" as before).

# ------------------------------------------------- first-run setup (requirements.txt from GitHub)
# The very first time the program runs it downloads requirements.txt from the GitHub repository into the
# SAME folder as this program and installs what is listed there. This block must stay above the other
# imports, because those are exactly the packages that may still be missing.
def _first_run_setup():
    import os, sys, subprocess, importlib, importlib.util, urllib.request
    if getattr(sys, 'frozen', False):                       # a packaged .exe already contains everything
        return
    here = os.path.dirname(os.path.abspath(__file__))        # the program's own folder (nothing to configure)
    marker = os.path.join(here, ".elstat_setup_done")
    req_path = os.path.join(here, "requirements.txt")
    required = ("pandas", "pypdf", "reportlab", "openpyxl")
    fallback = "pandas\npypdf\nreportlab\nopenpyxl\npillow\ntkinterdnd2\n"   # used if GitHub has no requirements.txt

    def missing():
        importlib.invalidate_caches()
        return [m for m in required if importlib.util.find_spec(m) is None]

    if os.path.exists(marker) and not missing():
        return                                               # already set up: start normally and instantly

    # ---- a small "please wait" window (skipped silently if there is no display)
    win = label = None
    base_text = ''
    try:
        import tkinter as _tk
        win = _tk.Tk()
        win.title("ELSTAT")
        base_text = ("Πρώτη εκκίνηση: εγκατάσταση απαιτούμενων πακέτων…\n"
                     "First run: installing the required packages…\n\nΠαρακαλώ περιμένετε / Please wait")
        label = _tk.Label(win, padx=30, pady=20, text=base_text)
        label.pack()
        win.update()
    except Exception:
        win = None

    result = {'ok': True, 'error': ''}

    def work():                                              # runs in the background so the window stays responsive
        try:
            # 1) download requirements.txt from GitHub into the program folder
            url = "https://raw.githubusercontent.com/webzerot/eltstat/main/requirements.txt"
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'ELSTAT-setup'})
                with urllib.request.urlopen(req, timeout=20) as resp:
                    data = resp.read(1024 * 1024)
                with open(req_path, 'wb') as f:
                    f.write(data)
            except Exception:
                if not os.path.exists(req_path):             # offline / not in the repo yet: use the built-in list
                    with open(req_path, 'w', encoding='utf-8') as f:
                        f.write(fallback)
            # 2) install it
            flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            cmd = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--timeout", "20",
                   "--retries", "2", "-r", req_path]
            run = subprocess.run(cmd, capture_output=True, text=True, creationflags=flags, stdin=subprocess.DEVNULL)
            if run.returncode != 0 and "No module named pip" in (run.stderr or ""):
                subprocess.run([sys.executable, "-m", "ensurepip", "--upgrade"], capture_output=True,
                               creationflags=flags, stdin=subprocess.DEVNULL)
                run = subprocess.run(cmd, capture_output=True, text=True, creationflags=flags,
                                     stdin=subprocess.DEVNULL)
            if run.returncode != 0:
                result['ok'], result['error'] = False, (run.stderr or run.stdout or "")[-800:]
        except Exception as e:
            result['ok'], result['error'] = False, str(e)

    import threading, time
    worker = threading.Thread(target=work, daemon=True)
    worker.start()
    started = time.time()
    while worker.is_alive():
        if win is not None:
            try:
                label.config(text=base_text + "  (%d s)" % int(time.time() - started))
                win.update()
            except Exception:
                win = None
        worker.join(0.1)
    ok, error = result['ok'], result['error']

    still = missing()
    if still:
        ok = False
    if ok:
        try:
            with open(marker, 'w', encoding='utf-8') as f:
                f.write("setup done\n")
        except OSError:
            pass
    if win is not None:
        try:
            win.destroy()
        except Exception:
            pass
    if still:                                                # the program cannot start without these
        msg = ("Δεν ήταν δυνατή η εγκατάσταση: " + ", ".join(still) + "\nCould not install: " + ", ".join(still)
               + "\n\nRun / Εκτελέστε:  pip install -r requirements.txt\n\n" + error)
        try:
            from tkinter import messagebox
            r = _tk.Tk(); r.withdraw()
            messagebox.showerror("ELSTAT", msg)
        except Exception:
            print(msg)
        sys.exit(1)


_first_run_setup()
import pandas as pd
import pypdf
import io
import csv
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
import difflib
import json
import hashlib
import base64
import tempfile
import copy
import time
import threading
import queue
import weakref
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from collections import OrderedDict
try:    # Pillow makes zooming smooth: the page is scaled instantly and the sharp picture follows in the background
    from PIL import Image as PILImage, ImageTk as PILImageTk
except Exception:
    PILImage = PILImageTk = None

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
        record(widget)
        return 'break'

    # Tk already copies / cuts / pastes by itself on Ctrl+C / Ctrl+X / Ctrl+V when the keyboard is on an
    # English layout. This handler (bound to "all", so it runs AFTER Tk's own binding, and cannot cancel it)
    # only has to step in for keys Tk does not understand: another layout such as Greek, or a shortcut the
    # user changed in Settings. Without this check the action ran twice and a paste landed twice.
    native_keys = {'copy': ('c', 'C'), 'cut': ('x', 'X'), 'paste': ('v', 'V')}

    def handled_by_tk(event, action):
        keys = native_keys.get(action)
        if not keys or sys.platform == 'darwin':          # macOS uses Command, not Control
            return False
        state = getattr(event, 'state', 0)
        state = state if isinstance(state, int) else 0
        ctrl = bool(state & 0x4)
        alt = bool(state & (0x20000 if sys.platform.startswith('win') else 0x8))
        return ctrl and not alt and str(getattr(event, 'keysym', '') or '') in keys

    def on_shortcut(event, action):
        widget = event.widget
        if not editable_widget(widget):
            return None
        if action == 'select_all':
            return select_all(widget)
        if handled_by_tk(event, action):
            return None                                   # Tk has already done it once
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
        menu.add_command(label=T('Αναίρεση', 'Undo'), command=lambda: do_history(widget, False))
        menu.add_command(label=T('Επανάληψη', 'Redo'), command=lambda: do_history(widget, True))
        menu.add_separator()
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

    binder = ShortcutBinder(root, use_all=True)      # the keys can be changed in Settings -> Keyboard shortcuts
    root._text_shortcuts = binder
    for sid in ('copy', 'cut', 'paste', 'select_all'):
        binder.add(sid, lambda event, a=sid: on_shortcut(event, a))
    # ---- undo / redo for every text input (Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z)
    history = weakref.WeakKeyDictionary()
    pending = weakref.WeakKeyDictionary()

    def record(widget):
        if not isinstance(widget, (tk.Entry, ttk.Entry)):
            return
        try:
            value = widget.get()
        except tk.TclError:
            return
        h = history.get(widget)
        if h is None:
            history[widget] = [[value], 0]
            return
        stack, pos = h
        if stack[pos] != value:
            del stack[pos + 1:]
            stack.append(value)
            if len(stack) > 200:
                del stack[0]
            h[1] = len(stack) - 1

    def do_history(widget, redo):
        if not isinstance(widget, (tk.Entry, ttk.Entry)) or is_readonly(widget):
            return None                     # tk.Text has its own undo (enabled on focus)
        record(widget)
        h = history.get(widget)
        if not h:
            return 'break'
        stack, pos = h
        new = pos + 1 if redo else pos - 1
        if 0 <= new < len(stack):
            h[1] = new
            try:
                widget.delete(0, 'end')
                widget.insert(0, stack[new])
                widget.icursor('end')
            except tk.TclError:
                pass
        return 'break'

    def on_focus_in(event):
        w = event.widget
        if isinstance(w, tk.Text):
            try:
                if str(w.cget('state')) != 'disabled':
                    w.configure(undo=True, autoseparators=True, maxundo=200)
            except tk.TclError:
                pass
        else:
            record(w)

    def on_key_release(event):
        w = event.widget
        if not isinstance(w, (tk.Entry, ttk.Entry)):
            return
        job = pending.pop(w, None)
        if job:
            try:
                w.after_cancel(job)
            except tk.TclError:
                pass
        try:
            pending[w] = w.after(500, lambda: record(w))
        except tk.TclError:
            pass

    binder.add('undo', lambda event: do_history(event.widget, False))
    binder.add('redo', lambda event: do_history(event.widget, True))
    binder.add('redo_alt', lambda event: do_history(event.widget, True))
    binder.apply()
    root.bind_all('<FocusIn>', on_focus_in, add='+')
    root.bind_all('<KeyRelease>', on_key_release, add='+')
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
DATA_FILES = ("hotel_settings.csv", "elstat_history.csv", "elstat_months.csv", "elstat_preferences.json", "elstat_country_mappings.json", "elstat_template_layout.json", "elstat_manual_layout.json", "elstat_detected_layout.json", "elstat_custom_fields.json")

DATA_DIR = SETTINGS_CSV = HISTORY_CSV = MONTHS_CSV = PREFS_JSON = ARCHIVE_DIR = COUNTRY_MAPPINGS_FILE = ""


def _write_json(path, data, **kw):
    """Saves JSON safely: written to a temporary file first, then swapped in, so a crash or a full disk
    in the middle of saving can never leave a half-written (corrupt) settings or layout file."""
    tmp = path + '.tmp'
    try:
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, **kw)
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:
                pass
        os.replace(tmp, path)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


_last_error_popup = 0.0


def _handle_error(exc, val, tb, parent=None):
    """Any unexpected error: written to elstat_error.log and shown once in plain words - the program keeps running."""
    import traceback
    global _last_error_popup
    text = ''.join(traceback.format_exception(exc, val, tb))
    where = ''
    for folder in (globals().get('DATA_DIR'), tempfile.gettempdir()):
        try:
            if folder:
                where = os.path.join(folder, 'elstat_error.log')
                with open(where, 'a', encoding='utf-8') as f:
                    f.write(f"\n=== {datetime.datetime.now():%Y-%m-%d %H:%M:%S} ===\n{text}")
                break
        except Exception:
            where = ''
    try:
        sys.stderr.write(text)
    except Exception:
        pass
    now = time.time()
    if now - _last_error_popup < 8:          # a burst of the same error must not bury the screen in pop-ups
        return
    _last_error_popup = now
    try:
        messagebox.showerror(
            T("Κάτι δεν πήγε καλά", "Something went wrong"),
            T("Παρουσιάστηκε ένα απρόσμενο σφάλμα, αλλά το πρόγραμμα συνεχίζει να λειτουργεί. "
              "Τα αποθηκευμένα δεδομένα σας δεν επηρεάστηκαν.", "An unexpected error happened, but the program keeps running. "
              "Your saved data was not affected.")
            + (T(f"\n\nΤεχνικές λεπτομέρειες: {where}", f"\n\nTechnical details: {where}") if where else '')
            + f"\n\n{exc.__name__}: {val}", parent=parent)
    except Exception:
        pass


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
        _write_json(CONFIG_FILE, cfg, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


# ------------------------------------------------------- online update (GitHub)
# The program updates itself from a GitHub repository: it reads APP_VERSION from main.py on the chosen
# branch, and if that is newer it can replace this file (the old one is kept as main.py.bak).
UPDATE_FILE = "main.py"
UPDATE_DEFAULT_REPO = "webzerot/eltstat"            # "owner/repository", e.g. "myname/elstat" (can also be typed in Settings)
UPDATE_DEFAULT_BRANCH = "main"
_VERSION_RE = re.compile(r'^APP_VERSION\s*=\s*[\'"]([^\'"]+)[\'"]', re.M)
_FORCE_RE = re.compile(r'^FORCE_UPDATE\s*=\s*(True|1)\b', re.M)      # read from the file on GitHub


def parse_github_repo(text):
    """Accepts 'owner/repo' or a github.com link and returns 'owner/repo' ('' if it is not valid)."""
    text = (text or '').strip()
    m = re.match(r'^(?:https?://)?(?:www\.)?github\.com/([\w.-]+)/([\w.-]+?)(?:\.git)?/?$', text)
    if m:
        return f"{m.group(1)}/{m.group(2)}"
    m = re.match(r'^([\w.-]+)/([\w.-]+?)(?:\.git)?$', text)
    return f"{m.group(1)}/{m.group(2)}" if m else ''


def version_tuple(v):
    return tuple(int(x) for x in re.findall(r'\d+', str(v or '')))


def changelog_section(text, version):
    """From a CHANGELOG.md returns the part under the heading that mentions `version` (e.g. "## 1.0.2").
    If no heading matches (or the section is empty), an empty string is returned and nothing is shown."""
    lines = text.replace('\r\n', '\n').split('\n')
    want = version_tuple(version)
    start = None
    for i, line in enumerate(lines):
        if line.lstrip().startswith('#') and version_tuple(line) == want and want:
            start = i
            break
    if start is None:
        return ''
    end = len(lines)
    for j in range(start + 1, len(lines)):
        if lines[j].lstrip().startswith('#') and version_tuple(lines[j]):      # the next version's heading
            end = j
            break
    body = [l for l in lines[start + 1:end] if l.strip()]
    return '\n'.join(lines[start:end]).strip()[:1500] if body else ''


def fetch_remote_update(repo, branch):
    """Downloads main.py from the repository. Returns {'version', 'newer', 'code' (bytes)}."""
    base = f"https://raw.githubusercontent.com/{repo}/{branch}/"
    stamp = f"?t={int(time.time())}"                  # a new address each time, so GitHub's cache is skipped
    headers = {'User-Agent': 'ELSTAT-updater', 'Cache-Control': 'no-cache'}
    with urllib.request.urlopen(urllib.request.Request(base + UPDATE_FILE + stamp, headers=headers), timeout=20) as resp:
        code = resp.read(30 * 1024 * 1024)
    notes = ''
    try:                                              # optional: a CHANGELOG.md in the repo is shown before installing
        with urllib.request.urlopen(urllib.request.Request(base + "CHANGELOG.md" + stamp, headers=headers),
                                    timeout=8) as resp:
            notes = resp.read(100 * 1024).decode('utf-8', errors='replace')
    except Exception:
        pass
    remote_text = code.decode('utf-8', errors='replace')
    m = _VERSION_RE.search(remote_text)
    if not m:
        raise ValueError("APP_VERSION not found in the repository's " + UPDATE_FILE)
    remote = m.group(1)
    return {'version': remote, 'code': code, 'repo': repo, 'branch': branch,
            'force': bool(_FORCE_RE.search(remote_text)),
            'notes': changelog_section(notes, remote) if notes.strip() else '',
            'newer': version_tuple(remote) > version_tuple(APP_VERSION)}


def apply_update(code):
    """Checks the downloaded program and replaces this file with it. Returns the path of the backup."""
    if getattr(sys, 'frozen', False):
        raise RuntimeError("frozen")
    text = code.decode('utf-8')
    compile(text, UPDATE_FILE, 'exec')                         # must at least be valid Python
    if 'class ElstatApp' not in text or not _VERSION_RE.search(text):
        raise ValueError("The downloaded file does not look like this program.")
    target = os.path.abspath(__file__)
    backup = target + '.bak'
    shutil.copy2(target, backup)
    tmp = target + '.new'
    with open(tmp, 'wb') as f:
        f.write(code)
    os.replace(tmp, target)
    return backup


def restore_backup():
    """Puts back the program that was saved as main.py.bak before the last update."""
    target = os.path.abspath(__file__)
    backup = target + '.bak'
    if not os.path.isfile(backup):
        raise FileNotFoundError(backup)
    with open(backup, 'rb') as f:
        code = f.read()
    compile(code.decode('utf-8'), UPDATE_FILE, 'exec')
    tmp = target + '.new'
    with open(tmp, 'wb') as f:
        f.write(code)
    os.replace(tmp, target)


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
BOOL_PREFS = ('open_after', 'confirm_overwrite', 'keep_archive', 'show_reminder', 'prefill_employees', 'preview_before_result', 'show_extra_fields', 'show_loading', 'auto_populate_layout', 'center_window')
DEFAULT_PREFS = {
    'language': 'el',                          # 'el' | 'en'  (interface language)
    'theme': 'light',                          # 'light' | 'dark' | 'auto'  (follow the system)
    'output_mode': 'program',                  # 'program' | 'excel' | 'custom'
    'output_dir': '',                          # used when output_mode == 'custom'
    'filename_pattern': 'ELSTAT_{year}_{month}',
    'excel_dir': '',                           # '' = remember the last folder
    'template_pdf': '',                        # '' = template_elstat_report.pdf next to the program
    'smart_ack': False,                        # the 'experimental' notice of Smart Shift was accepted
    'detect_ack': False,                       # the 'experimental' notice of Smart Field Detection was accepted
    'pdf_mode': 'overlay',                     # 'overlay' (Legacy: fixed coordinates) | 'smart' (Smart Shift) | 'detect' (Smart Field Detection) | 'manual' (layout editor)
    'manual_base': 'overlay',                  # "Set manually" starts from: 'overlay' | 'smart' | 'detect' | 'blank'
    'auto_populate_layout': True,              # the layout editor starts from what the selected PDF method would do (off = blank canvas)
    'open_after': False,
    'confirm_overwrite': True,
    'keep_archive': True,
    'show_reminder': True,
    'prefill_employees': False,
    'preview_before_result': False,             # show an in-app PDF preview before the Result dialog
    'show_extra_fields': True,                  # the optional "extra details" drop-down of the Hotel details tab
    'center_window': True,                     # open the window in the centre of the screen
    'show_loading': True,                       # animated loading indicator while something is working
    'shortcuts': {},                            # keyboard shortcuts the user changed: {id: "Ctrl+Shift+S"} ('' = none)
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
    prefs['shortcuts'] = {}
    try:
        with open(PREFS_JSON, encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):
            for k, default in DEFAULT_PREFS.items():
                if k in data:
                    if isinstance(default, dict):
                        prefs[k] = ({str(a): str(b or '') for a, b in data[k].items()}
                                    if isinstance(data[k], dict) else {})
                    else:
                        prefs[k] = bool(data[k]) if isinstance(default, bool) else str(data[k] or '')
            if 'manual_base' not in data and data.get('auto_populate_layout') is False:
                prefs['manual_base'] = 'blank'        # older settings: "auto-populate" off meant a blank canvas
    except Exception:
        pass
    if prefs['output_mode'] not in ('program', 'excel', 'custom'):
        prefs['output_mode'] = 'program'
    if prefs['language'] not in LANGUAGES:
        prefs['language'] = DEFAULT_PREFS['language']
    if prefs['pdf_mode'] not in ('overlay', 'smart', 'detect', 'manual'):
        prefs['pdf_mode'] = DEFAULT_PREFS['pdf_mode']
    if prefs['manual_base'] not in MANUAL_BASES:
        prefs['manual_base'] = DEFAULT_PREFS['manual_base']
    set_manual_base(prefs['manual_base'])
    if prefs['theme'] not in ('light', 'dark', 'auto'):
        prefs['theme'] = DEFAULT_PREFS['theme']
    if pattern_error(prefs['filename_pattern']):
        prefs['filename_pattern'] = DEFAULT_PREFS['filename_pattern']
    return prefs


def save_prefs(prefs):
    ensure_data_dir()
    _write_json(PREFS_JSON, prefs, ensure_ascii=False, indent=2)

_LATIN_TO_GREEK = str.maketrans({
    'A': 'Α', 'B': 'Β', 'E': 'Ε', 'H': 'Η', 'I': 'Ι', 'K': 'Κ', 'M': 'Μ',
    'N': 'Ν', 'O': 'Ο', 'P': 'Ρ', 'T': 'Τ', 'Y': 'Υ', 'X': 'Χ', 'Z': 'Ζ'})

# ------------------------------------------------------------- keyboard shortcuts
# Every keyboard shortcut of the program is listed in SHORTCUTS, grouped by the window it belongs to. The defaults
# are written here; the ones the person changes are saved in the preferences ('shortcuts') and edited in
# Settings -> Keyboard shortcuts. Windows bind them through ShortcutBinder, so a change works at once.
SHORTCUT_GROUPS = [
    ('general', ("Κύριο παράθυρο", "Main window")),
    ('text', ("Πεδία κειμένου", "Text fields")),
    ('history', ("Ιστορικό", "History")),
    ('result', ("Παράθυρο αποτελέσματος", "Result window")),
    ('preview', ("Προεπισκόπηση PDF", "PDF preview")),
    ('editor', ("Επεξεργαστής διάταξης", "Layout editor")),
]
# (id, group, (greek, english), default, flags)
#   flags: 'hold'  = one key that is held down (pressed / released), 'mod' = needs Ctrl or Alt (a plain key would
#          block typing), 'shift' = Shift + the key does the same with a bigger step
SHORTCUTS = [
    ('open_excel', 'general', ("Επιλογή αρχείου Excel ή CSV", "Choose the Excel or CSV file"), "Ctrl+O", ''),
    ('generate', 'general', ("Δημιουργία PDF", "Create PDF"), "Ctrl+Enter", ''),
    ('generate_print', 'general', ("Δημιουργία & Εκτύπωση", "Create & Print"), "Ctrl+P", ''),
    ('help', 'general', ("Άνοιγμα του οδηγού", "Open the guide"), "F1", ''),
    ('help_search', 'general', ("Αναζήτηση μέσα στον οδηγό", "Search inside the guide"), "Ctrl+F", ''),

    ('copy', 'text', ("Αντιγραφή", "Copy"), "Ctrl+C", 'mod'),
    ('cut', 'text', ("Αποκοπή", "Cut"), "Ctrl+X", 'mod'),
    ('paste', 'text', ("Επικόλληση", "Paste"), "Ctrl+V", 'mod'),
    ('select_all', 'text', ("Επιλογή όλων", "Select all"), "Ctrl+A", 'mod'),
    ('undo', 'text', ("Αναίρεση", "Undo"), "Ctrl+Z", 'mod'),
    ('redo', 'text', ("Επανάληψη", "Redo"), "Ctrl+Y", 'mod'),
    ('redo_alt', 'text', ("Επανάληψη (δεύτερη συντόμευση)", "Redo (second shortcut)"), "Ctrl+Shift+Z", 'mod'),

    ('hist_open', 'history', ("Άνοιγμα του επιλεγμένου μήνα", "Open the selected month"), "Enter", ''),
    ('hist_delete', 'history', ("Διαγραφή του επιλεγμένου μήνα", "Delete the selected month"), "Delete", ''),

    ('result_print', 'result', ("Εκτύπωση", "Print"), "Ctrl+P", ''),
    ('result_close', 'result', ("Κλείσιμο", "Close"), "Esc", ''),

    ('pv_zoom_in', 'preview', ("Μεγέθυνση", "Zoom in"), "Plus", ''),
    ('pv_zoom_out', 'preview', ("Σμίκρυνση", "Zoom out"), "Minus", ''),
    ('pv_fit', 'preview', ("Προσαρμογή στο παράθυρο", "Fit to the window"), "Ctrl+0", ''),
    ('pv_prev_page', 'preview', ("Προηγούμενη σελίδα", "Previous page"), "PageUp", ''),
    ('pv_next_page', 'preview', ("Επόμενη σελίδα", "Next page"), "PageDown", ''),
    ('pv_up', 'preview', ("Κύλιση πάνω", "Scroll up"), "Up", ''),
    ('pv_down', 'preview', ("Κύλιση κάτω", "Scroll down"), "Down", ''),
    ('pv_left', 'preview', ("Κύλιση αριστερά", "Scroll left"), "Left", ''),
    ('pv_right', 'preview', ("Κύλιση δεξιά", "Scroll right"), "Right", ''),
    ('pv_close', 'preview', ("Κλείσιμο της προεπισκόπησης", "Close the preview"), "Esc", ''),

    ('ed_save', 'editor', ("Αποθήκευση", "Save"), "Ctrl+S", ''),
    ('ed_undo', 'editor', ("Αναίρεση", "Undo"), "Ctrl+Z", ''),
    ('ed_redo', 'editor', ("Επανάληψη", "Redo"), "Ctrl+Y", ''),
    ('ed_duplicate', 'editor', ("Διπλασιασμός στοιχείου", "Duplicate the item"), "Ctrl+D", ''),
    ('ed_delete', 'editor', ("Διαγραφή στοιχείου που προσθέσατε", "Delete an item you added"), "Delete", ''),
    ('ed_search', 'editor', ("Αναζήτηση στη λίστα στοιχείων", "Search the item list"), "Ctrl+F", ''),
    ('ed_deselect', 'editor', ("Αποεπιλογή", "Deselect"), "Esc", ''),
    ('ed_guide', 'editor', ("Άνοιγμα του οδηγού", "Open the guide"), "F1", ''),
    ('ed_zoom_in', 'editor', ("Μεγέθυνση", "Zoom in"), "Plus", ''),
    ('ed_zoom_out', 'editor', ("Σμίκρυνση", "Zoom out"), "Minus", ''),
    ('ed_fit', 'editor', ("Προσαρμογή στο παράθυρο", "Fit to the window"), "Ctrl+0", ''),
    ('ed_pan', 'editor', ("Μετακίνηση της σελίδας (κρατήστε πατημένο και σύρετε)", "Pan the page (hold and drag)"), "Space", 'hold'),
    ('ed_left', 'editor', ("Μετακίνηση στοιχείου αριστερά (Shift = μεγαλύτερο βήμα)", "Move the item left (Shift = bigger step)"), "Left", 'shift'),
    ('ed_right', 'editor', ("Μετακίνηση στοιχείου δεξιά (Shift = μεγαλύτερο βήμα)", "Move the item right (Shift = bigger step)"), "Right", 'shift'),
    ('ed_up', 'editor', ("Μετακίνηση στοιχείου πάνω (Shift = μεγαλύτερο βήμα)", "Move the item up (Shift = bigger step)"), "Up", 'shift'),
    ('ed_down', 'editor', ("Μετακίνηση στοιχείου κάτω (Shift = μεγαλύτερο βήμα)", "Move the item down (Shift = bigger step)"), "Down", 'shift'),
]
SHORTCUT_INFO = {s[0]: dict(id=s[0], group=s[1], label=s[2], default=s[3], flags=s[4]) for s in SHORTCUTS}
SHORTCUT_OVERRIDES = {}          # id -> text of the shortcut the user chose ('' = none); only what differs from the default

_MODIFIER_KEYSYMS = {'Control_L', 'Control_R', 'Shift_L', 'Shift_R', 'Alt_L', 'Alt_R', 'Meta_L', 'Meta_R', 'Caps_Lock',
                     'Num_Lock', 'Scroll_Lock', 'Super_L', 'Super_R', 'Win_L', 'Win_R', 'ISO_Level3_Shift',
                     'Mode_switch', 'Multi_key'}
# name used in the saved text -> Tk key names it stands for
_NAME_KEYSYMS = {
    'Space': ('space',), 'Enter': ('Return', 'KP_Enter'), 'Esc': ('Escape',), 'Delete': ('Delete',),
    'Backspace': ('BackSpace',), 'Tab': ('Tab',), 'Insert': ('Insert',), 'Home': ('Home',), 'End': ('End',),
    'PageUp': ('Prior',), 'PageDown': ('Next',), 'Up': ('Up',), 'Down': ('Down',), 'Left': ('Left',),
    'Right': ('Right',), 'Plus': ('plus', 'KP_Add', 'equal'), 'Minus': ('minus', 'KP_Subtract'),
    'Comma': ('comma',), 'Period': ('period',), 'Slash': ('slash',),
}
_KEYSYM_NAMES = {'space': 'Space', 'Return': 'Enter', 'KP_Enter': 'Enter', 'Escape': 'Esc', 'Delete': 'Delete',
                 'KP_Delete': 'Delete', 'BackSpace': 'Backspace', 'Tab': 'Tab', 'ISO_Left_Tab': 'Tab',
                 'Insert': 'Insert', 'Home': 'Home', 'End': 'End', 'Prior': 'PageUp', 'Next': 'PageDown',
                 'Up': 'Up', 'Down': 'Down', 'Left': 'Left', 'Right': 'Right', 'plus': 'Plus', 'KP_Add': 'Plus',
                 'equal': 'Plus', 'minus': 'Minus', 'KP_Subtract': 'Minus', 'comma': 'Comma', 'period': 'Period',
                 'slash': 'Slash'}
_SYMBOL_KEYS = ('Plus', 'Minus', 'Comma', 'Period', 'Slash')      # Shift is part of the character itself
_NAME_LOOKUP = {k.lower(): k for k in _NAME_KEYSYMS}
_MOD_ORDER = (('Control', 'Ctrl'), ('Alt', 'Alt'), ('Shift', 'Shift'))


def accel_parse(accel):
    """'Ctrl+Shift+S' -> ({'Control', 'Shift'}, 'S')"""
    mods, key = set(), ''
    for part in str(accel or '').split('+'):
        part = part.strip()
        low = part.lower()
        if not part:
            continue
        if low in ('ctrl', 'control'):
            mods.add('Control')
        elif low == 'alt':
            mods.add('Alt')
        elif low == 'shift':
            mods.add('Shift')
        elif low in _NAME_LOOKUP:
            key = _NAME_LOOKUP[low]
        elif len(part) == 1 or re.fullmatch(r'[fF]\d{1,2}', part):
            key = part.upper()
        else:
            key = part
    return mods, key


def accel_normalize(accel):
    mods, key = accel_parse(accel)
    if not key:
        return ''
    return '+'.join([name for m, name in _MOD_ORDER if m in mods] + [key])


def accel_valid(accel):
    mods, key = accel_parse(accel)
    if not key:
        return False
    if len(key) == 1:
        return key.isascii() and key.isalnum()
    return key in _NAME_KEYSYMS or re.fullmatch(r'F([1-9]|1[0-2])', key) is not None


def accel_display(accel):
    return accel if accel else T("Κανένα", "None")


_GREEK_LAYOUT_KEYSYMS = {
    'A': ('Greek_alpha',), 'B': ('Greek_beta',), 'C': ('Greek_psi',),
    'D': ('Greek_delta',), 'E': ('Greek_epsilon',), 'F': ('Greek_phi',),
    'G': ('Greek_gamma',), 'H': ('Greek_eta',), 'I': ('Greek_iota',),
    'J': ('Greek_xi',), 'K': ('Greek_kappa',), 'L': ('Greek_lamda', 'Greek_lambda'),
    'M': ('Greek_mu',), 'N': ('Greek_nu',), 'O': ('Greek_omicron',),
    'P': ('Greek_pi',), 'R': ('Greek_rho',), 'S': ('Greek_sigma',),
    'T': ('Greek_tau',), 'U': ('Greek_theta',), 'V': ('Greek_omega',),
    'W': ('Greek_finalsmallsigma',), 'X': ('Greek_chi',), 'Y': ('Greek_upsilon',),
    'Z': ('Greek_zeta',),
}
_GREEK_CHAR_TO_LATIN = {
    'α':'A', 'β':'B', 'ψ':'C', 'δ':'D', 'ε':'E', 'φ':'F', 'γ':'G',
    'η':'H', 'ι':'I', 'ξ':'J', 'κ':'K', 'λ':'L', 'μ':'M', 'ν':'N',
    'ο':'O', 'π':'P', 'ρ':'R', 'σ':'S', 'τ':'T', 'θ':'U', 'ω':'V',
    'ς':'W', 'χ':'X', 'υ':'Y', 'ζ':'Z',
}
_GREEK_KEYSYM_TO_LATIN = {
    keysym.lower(): letter
    for letter, keysyms in _GREEK_LAYOUT_KEYSYMS.items()
    for keysym in keysyms
}


def accel_keysyms(key):
    if len(key) == 1 and key.isalpha():
        return [key.lower(), key.upper(), *_GREEK_LAYOUT_KEYSYMS.get(key.upper(), ())]
    if len(key) == 1 and key.isdigit():
        return [key, 'KP_' + key]
    return list(_NAME_KEYSYMS.get(key, (key,)))


def accel_sequences(accel, release=False, shift_variant=False):
    """The Tk key sequences of a shortcut text ('Ctrl+S' -> <Control-s>, <Control-S>)."""
    mods, key = accel_parse(accel)
    if not key:
        return []
    variants = [set(mods)]
    if shift_variant:
        variants.append(set(mods) | {'Shift'})
    seqs = []
    for m in variants:
        prefix = ''.join(name + '-' for name in ('Control', 'Alt', 'Shift') if name in m)
        for ks in accel_keysyms(key):
            seqs.append('<%s%s%s>' % (prefix, 'KeyRelease-' if release else '', ks))
    return seqs


def event_to_accel(event):
    """Shortcut text of a key press, normalized to the physical US-layout letter."""
    ks = str(getattr(event, 'keysym', '') or '')
    if ks in _MODIFIER_KEYSYMS:
        return None
    state = getattr(event, 'state', 0)
    ctrl = bool(state & 0x4)
    alt = bool(state & (0x20000 if sys.platform.startswith('win') else 0x8))
    shift = bool(state & 0x1)
    key = _GREEK_KEYSYM_TO_LATIN.get(ks.lower())
    if key is None and len(ks) == 1:
        key = _GREEK_CHAR_TO_LATIN.get(ks.lower())
    if key is None and ks in _KEYSYM_NAMES:
        key = _KEYSYM_NAMES[ks]
        if key in _SYMBOL_KEYS:
            shift = False
    elif key is None and len(ks) == 1 and ks.isascii() and ks.isalnum():
        key = ks.upper()
    elif key is None and re.fullmatch(r'F([1-9]|1[0-2])', ks):
        key = ks
    elif key is None and re.fullmatch(r'KP_[0-9]', ks):
        key = ks[3:]
    if key is None:
        return None
    return '+'.join((['Ctrl'] if ctrl else []) + (['Alt'] if alt else []) + (['Shift'] if shift else []) + [key])


def shortcut_accel(sid):
    return SHORTCUT_OVERRIDES.get(sid, SHORTCUT_INFO[sid]['default'])


def shortcut_display(sid):
    return accel_display(shortcut_accel(sid))


def shortcut_is_default(sid):
    return shortcut_accel(sid) == SHORTCUT_INFO[sid]['default']


def set_shortcut(sid, accel):
    accel = accel_normalize(accel)
    if accel == SHORTCUT_INFO[sid]['default']:
        SHORTCUT_OVERRIDES.pop(sid, None)
    else:
        SHORTCUT_OVERRIDES[sid] = accel


def load_shortcut_overrides(data):
    SHORTCUT_OVERRIDES.clear()
    if not isinstance(data, dict):
        return
    for sid, accel in data.items():
        if sid not in SHORTCUT_INFO:
            continue
        accel = accel_normalize(accel)
        if accel and not accel_valid(accel):
            continue
        if accel != SHORTCUT_INFO[sid]['default']:
            SHORTCUT_OVERRIDES[sid] = accel


def shortcut_conflict(sid, accel):
    """Another shortcut of the same window that already uses `accel` (None when it is free)."""
    accel = accel_normalize(accel)
    if not accel:
        return None
    group = SHORTCUT_INFO[sid]['group']
    for other, info in SHORTCUT_INFO.items():
        if other != sid and info['group'] == group and shortcut_accel(other) == accel:
            return other
    return None


def sc_text(text):
    """Replaces {sc:id} in a text with the shortcut that is set right now."""
    return re.sub(r'\{sc:(\w+)\}', lambda m: shortcut_display(m.group(1)) if m.group(1) in SHORTCUT_INFO else m.group(0), text)


class ShortcutBinder:
    """Binds shortcuts for one window. On Windows, letter shortcuts are matched
    by the physical virtual-key code so the active input language does not matter."""
    _instances = weakref.WeakSet()

    def __init__(self, widget, use_all=False):
        self.widget = widget
        self.use_all = use_all
        self._entries = []
        self._bound = set()
        self._map = {}
        self._win_map = {}
        self._win_release_map = {}
        self._win_press_bound = False
        self._win_release_bound = False
        ShortcutBinder._instances.add(self)

    def add(self, sid, fn, release=None):
        self._entries.append((sid, fn, release))
        return self

    def _bind(self, sequence, handler):
        if self.use_all:
            self.widget.bind_all(sequence, handler, add='+')
        else:
            self.widget.bind(sequence, handler, add='+')

    def apply(self):
        table = {}
        win_map = {}
        win_release_map = {}
        is_windows = sys.platform.startswith('win')
        for sid, fn, release in self._entries:
            accel = shortcut_accel(sid)
            if not accel:
                continue
            mods, key = accel_parse(accel)
            info = {
                'shift': 'Shift' in mods,
                'ctrl': 'Control' in mods,
                'alt': 'Alt' in mods,
                'letter': len(key) == 1 and key.isalpha(),
                'plain': not (mods & {'Control', 'Alt'}) and
                         (len(key) == 1 or key == 'Space' or key in _SYMBOL_KEYS),
            }

            # Tk's <Control-Greek_sigma> matching is inconsistent across Windows/Tk
            # versions. Use the Windows virtual-key code for A-Z instead.
            if is_windows and info['letter']:
                variants = [dict(info)]
                if 'shift' in SHORTCUT_INFO[sid]['flags']:
                    variants.append(dict(info, shift=True))
                vk = ord(key.upper())
                for variant in variants:
                    win_map.setdefault(vk, []).append((fn, variant))
                    if release:
                        win_release_map.setdefault(vk, []).append(
                            (release, dict(variant, plain=False)))
                continue

            for seq in accel_sequences(accel, shift_variant='shift' in SHORTCUT_INFO[sid]['flags']):
                table.setdefault(seq, []).append((fn, info))
            if release:
                for seq in accel_sequences(accel, release=True):
                    table.setdefault(seq, []).append((release, dict(info, plain=False, letter=False)))

        self._map = table
        self._win_map = win_map
        self._win_release_map = win_release_map

        for seq in table:
            if seq in self._bound:
                continue
            try:
                handler = lambda event, s=seq: self._dispatch(s, event)
                self._bind(seq, handler)
                self._bound.add(seq)
            except tk.TclError:
                pass

        if is_windows and win_map and not self._win_press_bound:
            try:
                self._bind('<KeyPress>', self._dispatch_windows_press)
                self._win_press_bound = True
            except tk.TclError:
                pass
        if is_windows and win_release_map and not self._win_release_bound:
            try:
                self._bind('<KeyRelease>', self._dispatch_windows_release)
                self._win_release_bound = True
            except tk.TclError:
                pass

    @staticmethod
    def _windows_letter_keycode(event):
        # Tk on Windows normally reports the Windows VK_A..VK_Z values here.
        code = getattr(event, 'keycode', None)
        if isinstance(code, int) and 0x41 <= code <= 0x5A:
            return code
        # Fallback for Tk builds that expose the translated Greek keysym.
        ks = str(getattr(event, 'keysym', '') or '')
        letter = _GREEK_KEYSYM_TO_LATIN.get(ks.lower())
        if letter is None and len(ks) == 1:
            letter = _GREEK_CHAR_TO_LATIN.get(ks.lower())
        return ord(letter) if letter else None

    @staticmethod
    def _dispatch_windows_entries(entries, event):
        if not entries:
            return None
        state = getattr(event, 'state', 0)
        state = state if isinstance(state, int) else 0
        ctrl = bool(state & 0x4)
        alt = bool(state & (0x20000 if sys.platform.startswith('win') else 0x8))
        shift = bool(state & 0x1)
        caps = bool(state & 0x2)
        effective_shift = shift != caps
        result = None
        for fn, info in entries:
            if ctrl != info['ctrl'] or alt != info['alt'] or effective_shift != info['shift']:
                continue
            if info['plain'] and isinstance(event.widget, (tk.Entry, ttk.Entry, ttk.Combobox, ttk.Spinbox, tk.Text)):
                continue
            if fn(event) == 'break':
                result = 'break'
        return result

    def _dispatch_windows_press(self, event):
        vk = self._windows_letter_keycode(event)
        return self._dispatch_windows_entries(self._win_map.get(vk, ()), event)

    def _dispatch_windows_release(self, event):
        vk = self._windows_letter_keycode(event)
        return self._dispatch_windows_entries(self._win_release_map.get(vk, ()), event)

    def _dispatch(self, seq, event):
        result = None
        state = getattr(event, 'state', 0)
        state = state if isinstance(state, int) else 0
        for fn, info in self._map.get(seq, ()):
            if info['letter'] and (bool(state & 0x1) != bool(state & 0x2)) != info['shift']:
                continue
            if info['plain'] and isinstance(event.widget, (tk.Entry, ttk.Entry, ttk.Combobox, ttk.Spinbox, tk.Text)):
                continue
            if fn(event) == 'break':
                result = 'break'
        return result

    @classmethod
    def refresh_all(cls):
        for inst in list(cls._instances):
            try:
                if inst.widget.winfo_exists():
                    inst.apply()
            except tk.TclError:
                pass


def norm(s):
    """Upper-case, strip all whitespace and turn Latin look-alikes into Greek,
    so Excel names and template text compare equal however they are spaced/typed."""
    s = unicodedata.normalize('NFD', str(s).upper())
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')   # drop accents / diaeresis
    return ''.join(s.split()).translate(_LATIN_TO_GREEK)


REQUIRED_KEYS = ('type', 'name', 'category', 'email', 'phone', 'afm', 'company')
# optional address details - never required, written next to the matching labels of the form
EXTRA_KEYS = ('nomos', 'dimotiki', 'odos', 'arithmos', 'tk')
SETTING_KEYS = REQUIRED_KEYS + EXTRA_KEYS
CF_PREFIX = 'cf:'          # custom field values travel inside the settings dict as 'cf:<id>'


def custom_fields_file():
    return os.path.join(DATA_DIR, "elstat_custom_fields.json")


def load_custom_fields():
    """[{'id': 'cf1', 'name': 'Label', 'value': 'text'}] - the user's own optional fields."""
    out, seen = [], set()
    try:
        with open(custom_fields_file(), encoding='utf-8') as f:
            data = json.load(f)
        items = data.get('fields', []) if isinstance(data, dict) else []
    except Exception:
        items = []
    for it in items if isinstance(items, list) else []:
        if not isinstance(it, dict):
            continue
        fid, name = str(it.get('id', '')).strip(), str(it.get('name', '')).strip()
        if not re.fullmatch(r'cf\d+', fid) or not name or fid in seen:
            continue
        seen.add(fid)
        out.append({'id': fid, 'name': name, 'value': str(it.get('value', '') or '')})
    return out


def save_custom_fields(fields):
    ensure_data_dir()
    _write_json(custom_fields_file(), {'version': 1, 'fields': fields}, ensure_ascii=False, indent=1)


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
    df = pd.DataFrame([{k: v for k, v in settings_data.items() if not str(k).startswith(CF_PREFIX)}])
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

# ---------------------------------------------------------------- Protel file: Excel (.xls/.xlsx) or CSV
SHEET_EXTS = ('.xls', '.xlsx', '.csv')
_DATE_RE = re.compile(r'^\s*(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})\s*$')
# where the readers expect each figure (the same positions the Excel sheet has)
_COL_PAX, _COL_ROOMN, _COL_NIGHTS, _COL_LOGIS, _COL_FB, _COL_EXTRAS, _COL_TOTAL = 7, 8, 9, 10, 11, 12, 13
_SHEET_WIDTH = 19


def is_csv_path(path):
    return str(path).lower().endswith('.csv')


def _decode_csv_bytes(raw):
    """CSV text from bytes: UTF-8 (with/without BOM) or UTF-16 when marked, otherwise Windows-1253 (what Protel writes)."""
    if raw.startswith(b'\xef\xbb\xbf'):
        return raw[3:].decode('utf-8', errors='replace')
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return raw.decode('utf-16', errors='replace')
    try:
        return raw.decode('utf-8')
    except UnicodeDecodeError:
        return raw.decode('cp1253', errors='replace')


def _csv_delimiter(text):
    """Comma, semicolon or tab - whichever splits the first lines into the most columns."""
    sample = text[:200000]
    best, best_n = ',', 0
    for d in (',', ';', '\t'):
        n = 0
        try:
            for i, r in enumerate(csv.reader(io.StringIO(sample, newline=''), delimiter=d)):
                n = max(n, len(r))
                if i >= 40:
                    break
        except csv.Error:
            continue
        if n > best_n:
            best, best_n = d, n
    return best


def _csv_number(v):
    """'20.341,00' / '813,00' / '256' / '1,234.50' / '-12.19' / '100%' -> number (whole numbers stay int); None if not a number."""
    if isinstance(v, (int, float)):
        return v
    t = str(v).strip().replace('\xa0', '').replace(' ', '')
    if t.endswith('%'):
        t = t[:-1]
    if not t or not re.fullmatch(r'[-+]?[\d.,]+', t) or not any(ch.isdigit() for ch in t):
        return None
    neg = t.startswith('-')
    t = t.lstrip('+-')
    if ',' in t and '.' in t:                         # the separator that comes last is the decimal one
        dec = ',' if t.rfind(',') > t.rfind('.') else '.'
        t = t.replace('.' if dec == ',' else ',', '').replace(dec, '.')
    elif ',' in t:
        t = t.replace(',', '.') if t.count(',') == 1 else t.replace(',', '')
    elif '.' in t and re.fullmatch(r'\d{1,3}(\.\d{3})+', t):
        t = t.replace('.', '')                        # 1.113 -> 1113 (Greek thousands dot)
    try:
        x = float(t)
    except ValueError:
        return None
    if neg:
        x = -x
    return int(x) if '.' not in t and x == int(x) else x


def _csv_parse_date(text, month_first=False):
    m = _DATE_RE.match(str(text))
    if not m:
        return None
    a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    d, mo = (b, a) if month_first else (a, b)
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return None


def _csv_resolve_period(s_from, s_to):
    """The two header dates as dates. Day/month (Greek) first; month/day only if that gives a sensible period instead."""
    cands = []
    for month_first in (False, True):
        a, b = _csv_parse_date(s_from, month_first), _csv_parse_date(s_to, month_first)
        if a and b and a <= b:
            cands.append((a, b))
    for c in cands:
        if period_as_month(c):
            return c
    return cands[0] if cands else None


def _csv_find_period(rows, limit=15):
    """Locates the 'Από/Έως Ημερομηνία' labels in the first rows. -> ({'from': (r, c), 'to': (r, c)}, (date_from, date_to) or None)."""
    pos = {}
    for r in range(min(len(rows), limit)):
        for c, cell in enumerate(rows[r]):
            label = norm(fix_greek(cell)) if isinstance(cell, str) else ''
            for key, tag in (('ΑΠΟΗΜΕΡΟΜΗΝΙΑ', 'from'), ('ΕΩΣΗΜΕΡΟΜΗΝΙΑ', 'to')):
                if label.startswith(key) and tag not in pos:
                    for c2 in range(c + 1, len(rows[r])):
                        if _DATE_RE.match(str(rows[r][c2])):
                            pos[tag] = (r, c2)
                            break
    period = None
    if 'from' in pos and 'to' in pos:
        period = _csv_resolve_period(rows[pos['from'][0]][pos['from'][1]], rows[pos['to'][0]][pos['to'][1]])
    return pos, period


def _csv_modif_index(row):
    for j, c in enumerate(row):
        if re.sub(r'\s+', '', str(c)).lower().startswith('%modif'):
            return j
    return None


def _flat_columns(headers):
    """Index (0-15) of each figure among a flattened record's 16 values, read from the headers; usual positions if unclear."""
    default = {'pax': 4, 'roomn': 5, 'nights': 6, 'logis': 7, 'fb': 8, 'extras': 9, 'total': 10}
    pos = {}
    for i, h in enumerate(headers):
        t = re.sub(r'\s+', '', str(h)).lower()
        if t.startswith('pr.'):
            continue
        if 'pax' in t and 'arriv' in t:
            pos.setdefault('pax', i)
        elif 'room' in t and 'night' in t:
            pos.setdefault('roomn', i)
        elif 'adult' in t and 'night' in t:
            pos.setdefault('nights', i)
        elif t.startswith('logis'):
            pos.setdefault('logis', i)
        elif t.startswith('f&b'):
            pos.setdefault('fb', i)
        elif t.startswith('extras'):
            pos.setdefault('extras', i)
        elif t.startswith('total'):
            pos.setdefault('total', i)
    if len(pos) == 7 and len(set(pos.values())) == 7:
        return pos
    return default


def _sheet_row(name, values, cols):
    """One Excel-shaped row: name in column 0, the 16 values from column 3, key figures at their fixed Excel columns."""
    row = [None] * _SHEET_WIDTH
    row[0] = name
    conv = []
    for i, v in enumerate(values[:16]):
        n = _csv_number(v)
        conv.append(n)
        row[3 + i] = n if n is not None else (str(v).strip() or None)
    for key, col in (('pax', _COL_PAX), ('roomn', _COL_ROOMN), ('nights', _COL_NIGHTS), ('logis', _COL_LOGIS),
                     ('fb', _COL_FB), ('extras', _COL_EXTRAS), ('total', _COL_TOTAL)):
        i = cols[key]
        row[col] = conv[i] if i < len(conv) else None
    return row


def _csv_flat_grid(rows):
    """Protel's 'Nationalities Statistics Comparison' CSV repeats the header and totals on every line (one line per country).
    Turns it into the same grid as the Excel sheet. None when the file is not in that shape."""
    recs = []
    for r in rows:
        j = _csv_modif_index(r)
        if j is not None and len(r) >= j + 18:
            recs.append((r, j))
    if not recs:
        return None
    grid, total_row = [], None
    pos, period = _csv_find_period([recs[0][0]], limit=1)
    head = [[None] * _SHEET_WIDTH for _ in range(3)]
    head[0][0], head[1][0] = 'Από Ημερομηνία', 'Έως Ημερομηνία'
    if period:
        head[0][1], head[1][1] = period
    for col, label in ((_COL_PAX, 'Pax'), (_COL_ROOMN, 'Room Nights'), (_COL_NIGHTS, 'Adult Nights'),
                       (_COL_LOGIS, 'Logis'), (_COL_FB, 'F&B'), (_COL_EXTRAS, 'Extras'), (_COL_TOTAL, 'Total')):
        head[2][col] = label
    for r, j in recs:
        cols = _flat_columns(r[max(0, j - 15):j + 1] if j >= 15 else [])
        name = str(r[j + 1]).strip()
        if name:
            grid.append(_sheet_row(name, r[j + 2:j + 18], cols))
        k = j + 18                                    # "Σύνολα :" block that follows the country
        if total_row is None and k < len(r) and norm(fix_greek(r[k])).startswith(('ΣΥΝΟΛ', 'ΤΟΤΑΛ')):
            total_row = _sheet_row(str(r[k]).strip(), r[k + 1:k + 17], cols)
    if total_row:
        grid.append(total_row)
    return head + grid


def _csv_plain_grid(rows):
    """An ordinary CSV copy of the Excel sheet: same columns, cells converted the way Excel would store them."""
    pos, period = _csv_find_period(rows)
    out = []
    for r, row in enumerate(rows):
        cells = []
        for c, cell in enumerate(row):
            t = str(cell).strip()
            if t == '':
                cells.append(None)
            elif period and (r, c) == pos.get('from'):
                cells.append(period[0])
            elif period and (r, c) == pos.get('to'):
                cells.append(period[1])
            elif c == 0:
                cells.append(t)
            else:
                n = _csv_number(t)
                cells.append(n if n is not None else t)
        out.append(cells)
    return out


def _read_csv_grid(path):
    with open(path, 'rb') as f:
        text = _decode_csv_bytes(f.read())
    delim = _csv_delimiter(text)
    rows = [r for r in csv.reader(io.StringIO(text, newline=''), delimiter=delim)]
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        raise ValueError(T("Το αρχείο CSV είναι κενό.", "The CSV file is empty."))
    grid = _csv_flat_grid(rows)
    if grid is None:
        grid = _csv_plain_grid(rows)
    width = max(len(r) for r in grid)
    width = max(width, _SHEET_WIDTH)
    return pd.DataFrame([r + [None] * (width - len(r)) for r in grid], dtype=object)


def _sniff_kind(path):
    """What the file really is, whatever its extension says: 'xls' (old binary Excel), 'xlsx' (zip), 'html' or 'text'."""
    try:
        with open(path, 'rb') as f:
            head = f.read(2048)
    except OSError:
        return 'text'
    if head.startswith(b'\xd0\xcf\x11\xe0'):
        return 'xls'
    if head.startswith(b'PK'):
        return 'xlsx'
    if head.lstrip(b'\xef\xbb\xbf \t\r\n').lower().startswith((b'<html', b'<!doctype', b'<table', b'<?xml')):
        return 'html'
    return 'text'


def _read_excel_any(path, nrows=None):
    """Reads an Excel file (.xls / .xlsx) as a header-less table. Tries every reader that is available and,
    as a last resort, converts the file with LibreOffice, so a missing 'xlrd' no longer makes a valid .xls unreadable."""
    kind = _sniff_kind(path)
    errors = []
    engines = [None, 'calamine'] if kind == 'xls' else [None, 'openpyxl']
    for eng in engines:
        try:
            kw = {'engine': eng} if eng else {}
            return pd.read_excel(path, header=None, nrows=nrows, **kw)
        except Exception as e:
            errors.append(str(e).splitlines()[0] if str(e) else repr(e))
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if soffice:
        tmpdir = tempfile.mkdtemp(prefix='xls2xlsx_')
        try:
            src = os.path.join(tmpdir, 'in' + ('.xls' if kind == 'xls' else '.xlsx'))
            shutil.copyfile(path, src)
            subprocess.run([soffice, '--headless', '--convert-to', 'xlsx', '--outdir', tmpdir, src],
                           capture_output=True, timeout=90)
            out = os.path.join(tmpdir, 'in.xlsx')
            if os.path.exists(out) and out != src:
                return pd.read_excel(out, header=None, nrows=nrows, engine='openpyxl')
        except Exception as e:
            errors.append(repr(e))
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
    hint = T("Εγκαταστήστε το πακέτο xlrd:  pip install xlrd", "Install the xlrd package:  pip install xlrd")
    raise ValueError(' | '.join(errors[:2]) + '\n' + hint)


def read_sheet(path, nrows=None):
    """The Protel sheet as a header-less table - the same shape pd.read_excel(header=None) gives.
    .xls, .xlsx and .csv are all accepted (the real content decides, not the extension), and each one is turned into
    that same table, so the preview, the PDF, the history and the revenue treat them alike."""
    kind = _sniff_kind(path)
    if kind in ('xls', 'xlsx'):
        return _read_excel_any(path, nrows)
    if kind == 'html':
        try:
            with open(path, 'rb') as f:
                tables = pd.read_html(io.StringIO(_decode_csv_bytes(f.read())), header=None)
            df = tables[0].astype(object)
            return df.iloc[:nrows] if nrows else df
        except Exception:
            pass
    if is_csv_path(path) or kind in ('text', 'html'):
        df = _read_csv_grid(path)
        return df.iloc[:nrows] if nrows else df
    return _read_excel_any(path, nrows)


def _csv_is_protel(path):
    """True for a CSV that looks like the Protel nationality report (so unrelated CSVs in the folder are ignored)."""
    try:
        with open(path, 'rb') as f:
            head = f.read(400000)
    except OSError:
        return False
    t = _decode_csv_bytes(head).lower()
    return 'nationalit' in t or ('arrivals' in t and 'revenue' in t)


def read_excel_period(excel_path):
    """Reads the report period (Από / Έως Ημερομηνία) printed in the Protel header.
    Returns (date_from, date_to) or None if it cannot be found."""
    try:
        df = read_sheet(excel_path, nrows=15)
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
        errors.append(T("Επιλέξτε υπαρκτό αρχείο Excel ή CSV.", "Choose an existing Excel or CSV file."))
    else:
        errors.extend(protel_problems(excel))
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
    df = read_sheet(excel_path)
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


def verify_protel_file(path):
    """Checks that a file really is a Protel 'Nationalities Statistics' export and not just any Excel/CSV.
    Returns a list of readable problems (empty list = the file is genuine).
    Checks: Protel header fingerprint (protel / report title / report id / Printdate), the expected column headers,
    the report period, Protel's own 'Totals' row, and that the country rows add up to that Totals row
    (so an edited or hand-made table is caught)."""
    problems = []
    try:
        df = read_sheet(path)
        df = df.map(fix_greek) if hasattr(df, 'map') else df.applymap(fix_greek)
    except Exception as e:
        return [T("Το αρχείο δεν μπορεί να διαβαστεί ως Excel/CSV: ", "The file cannot be read as Excel/CSV: ") + str(e)]
    if df.empty:
        return [T("Το αρχείο είναι κενό.", "The file is empty.")]

    head = []                                     # text cells of the Protel header block (top of the file)
    if is_csv_path(path) and _sniff_kind(path) == 'text':   # a CSV's real header is read raw (read_sheet rebuilds the header rows)
        try:
            with open(path, 'rb') as f:
                text = _decode_csv_bytes(f.read(400000))
            raw = list(csv.reader(io.StringIO(text, newline=''), delimiter=_csv_delimiter(text)))[:30]
        except Exception:
            raw = []
        cells = [v for row in raw for v in row]
    else:
        cells = [df.iat[r, c] for r in range(min(len(df), 30)) for c in range(df.shape[1])]
    for v in cells:
        v = fix_greek(v)
        if isinstance(v, str) and v.strip():
            head.append(''.join(v.lower().split()))
    joined = ' '.join(head)

    if 'protel' not in head:
        problems.append(T("Δεν βρέθηκε η υπογραφή «protel» στην κεφαλίδα.", "The “protel” signature was not found in the header."))
    if not any(h.startswith('nationalitystatistik') for h in head):
        problems.append(T("Δεν βρέθηκε ο κωδικός αναφοράς «nationalitystatistik» του Protel.",
                          "The Protel report id “nationalitystatistik” was not found."))
    if 'nationalitiesstatistics' not in joined:
        problems.append(T("Ο τίτλος αναφοράς δεν είναι «Nationalities Statistics».",
                          "The report title is not “Nationalities Statistics”."))
    if not any(h.startswith('printdate') for h in head):
        problems.append(T("Λείπει η γραμμή «Printdate» που τυπώνει το Protel.", "The “Printdate” line that Protel prints is missing."))
    missing_cols = [w for w in ('arrivals', 'nights', 'logis', 'f&b', 'extras', 'revenue') if w not in joined]
    if missing_cols:
        problems.append(T("Λείπουν οι στήλες του Protel: ", "Protel columns missing: ") + ', '.join(missing_cols))
    if not read_excel_period(path):
        problems.append(T("Δεν βρέθηκε η περίοδος (Από / Έως Ημερομηνία) στην κεφαλίδα.",
                          "The report period (From / To date) was not found in the header."))
    if problems:                                   # no point checking the numbers of a file that is not a Protel report
        return problems

    try:
        rows, total = read_nationality_rows(path)
    except Exception:
        return [T("Τα δεδομένα χωρών δεν μπορούν να διαβαστούν.", "The country data cannot be read.")]
    if not rows:
        problems.append(T("Δεν βρέθηκαν γραμμές χωρών.", "No country rows were found."))
    if total is None:
        problems.append(T("Λείπει η γραμμή «Σύνολα» του Protel.", "Protel's “Totals” row is missing."))
    elif rows and (sum(r[1] for r in rows), sum(r[2] for r in rows)) != tuple(total):
        problems.append(T(f"Τα αθροίσματα δεν ταιριάζουν με τη γραμμή «Σύνολα» (χώρες: {sum(r[1] for r in rows)}/{sum(r[2] for r in rows)}, "
                          f"σύνολα: {total[0]}/{total[1]}). Το αρχείο μοιάζει να έχει τροποποιηθεί.",
                          f"The country rows do not add up to the “Totals” row (countries: {sum(r[1] for r in rows)}/{sum(r[2] for r in rows)}, "
                          f"totals: {total[0]}/{total[1]}). The file looks edited."))
    return problems


PROTEL_OVERRIDES = set()      # files the user chose to use even though the Protel check failed


def _override_key(path):
    try:
        return os.path.normcase(os.path.abspath(path))
    except Exception:
        return str(path)


def protel_problems(path):
    """Like verify_protel_file, but empty for a file the user already chose to use anyway."""
    if _override_key(path) in PROTEL_OVERRIDES:
        return []
    return verify_protel_file(path)


def ask_use_anyway(problems, parent=None):
    """Asks whether to use a file that failed the Protel check. Remembers a 'yes' for that file (done by the caller)."""
    return messagebox.askyesno(
        T("Μη αναγνωρισμένο αρχείο", "File not recognised"),
        T("Το αρχείο δεν φαίνεται να είναι αναφορά «Nationalities Statistics» από το Protel:",
          "This file does not look like a “Nationalities Statistics” report from Protel:")
        + "\n\n" + "\n".join("• " + x for x in problems)
        + "\n\n" + T("Να χρησιμοποιηθεί παρ' όλα αυτά;", "Use it anyway?"),
        icon='warning', default='no', parent=parent)


def protel_error_text(path):
    """One message for dialogs, or '' when the file is genuine."""
    p = verify_protel_file(path)
    if not p:
        return ''
    return (T("Το αρχείο δεν φαίνεται να είναι αναφορά «Nationalities Statistics» από το Protel:",
              "This file does not look like a “Nationalities Statistics” report from Protel:")
            + '\n' + '\n'.join('• ' + x for x in p))


def read_revenue_rows(excel_path, strict=False):
    """Revenue columns of the Protel 'Nationalities Statistics' sheet:
    {country_as_in_excel: (rooms, f&b, extras, total, room_nights)}.  Columns are found by their header (Logis / F&B / Extras /
    Total), falling back to the usual positions.  Empty dict when nothing is found (strict=True raises instead)."""
    try:
        df = read_sheet(excel_path)
        df = df.map(fix_greek) if hasattr(df, 'map') else df.applymap(fix_greek)
        cols = {}
        for r in range(min(len(df), 25)):
            for c in range(df.shape[1]):
                v = df.iat[r, c]
                if not isinstance(v, str):
                    continue
                t = ''.join(v.lower().split())
                for key, names in (('rooms', ('logis',)), ('fb', ('f&b',)), ('extras', ('extras',)),
                                   ('total', ('total',))):
                    if t in names and key not in cols and c >= 8:
                        cols[key] = c
        idx = [cols.get(k) for k in ('rooms', 'fb', 'extras', 'total')]
        if None in idx or len(set(idx)) < 4:
            idx = [10, 11, 12, 13]
        idx.append(8)                      # room nights, for the average room rate
        if df.shape[1] <= max(idx):
            if strict:
                raise ValueError(f"only {df.shape[1]} columns found, revenue columns missing")
            return {}
        def num(v):
            if v is None or (isinstance(v, float) and v != v):
                return 0.0
            if isinstance(v, datetime.time):       # a 0 cell formatted as time reads back as 00:00
                return (v.hour * 3600 + v.minute * 60 + v.second) / 86400.0
            if isinstance(v, datetime.datetime):   # a number formatted as a date: turn back into the Excel serial
                return (v - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
            if isinstance(v, str):
                v = v.strip().replace('\xa0', '').replace(' ', '')
                if ',' in v and '.' in v:
                    v = v.replace('.', '').replace(',', '.') if v.rfind(',') > v.rfind('.') else v.replace(',', '')
                else:
                    v = v.replace(',', '.')
                if v in ('', '-'):
                    return 0.0
            return float(v)

        out, sample = {}, None
        for i in range(len(df)):
            row = df.iloc[i]
            if pd.isna(row.iloc[0]):
                continue
            name = str(row.iloc[0]).strip()
            key = norm(name)
            if not key or key.startswith('ΣΥΝΟΛ') or key.startswith('ΤΟΤΑΛ'):
                continue
            try:                                   # same test as the nationality reader: a real country row
                int(row.iloc[7]); int(row.iloc[9])
            except (ValueError, TypeError):
                continue
            try:
                out[name] = tuple(num(row.iloc[j]) for j in idx)
            except (ValueError, TypeError):
                sample = sample or [repr(row.iloc[j]) for j in idx]
        if not out and strict:
            raise ValueError(f"no country rows with revenue (columns used: {idx}, sheet {df.shape[0]}x{df.shape[1]}, sample {sample})")
        return out
    except Exception:
        if strict:
            raise
        return {}


def attach_revenue(excel_path):
    """Adds the revenue of a Protel 'Nationalities Statistics' file to a month that is already in the history.
    Returns (year, month, total_revenue) or raises ValueError with a readable reason."""
    bad = protel_problems(excel_path)
    if bad:
        raise ValueError(T("Δεν είναι αρχείο Protel: ", "Not a Protel file: ") + '; '.join(bad))
    info = period_as_month(read_excel_period(excel_path))
    if not info:
        raise ValueError(T("Το αρχείο δεν καλύπτει έναν ακριβώς ημερολογιακό μήνα.",
                           "The file does not cover exactly one calendar month."))
    year, month = info[0], info[1]
    rev = read_revenue_rows(excel_path, strict=True)
    if not rev:
        raise ValueError(T("Δεν βρέθηκαν στήλες εσόδων στο αρχείο.", "No revenue columns found in the file."))
    if not os.path.exists(HISTORY_CSV):
        raise ValueError(T("Δεν υπάρχει ιστορικό ακόμη.", "There is no history yet."))
    df = pd.read_csv(HISTORY_CSV, encoding='utf-8-sig', dtype=str, keep_default_na=False)
    for col in ('Logis', 'FnB', 'Extras', 'Esoda', 'Domatia'):
        if col not in df.columns:
            df[col] = ''
    by_target = {}
    for country_excel, vals in rev.items():
        target = resolve_country(country_excel)
        if target is None:
            continue
        old = by_target.get(target, (0.0,) * 5)
        by_target[target] = tuple(a + b for a, b in zip(old, vals))
    hit = ((pd.to_numeric(df['Etos'], errors='coerce') == year) & (pd.to_numeric(df['Minas'], errors='coerce') == month))
    if not hit.any():
        raise ValueError(T(f"Ο μήνας {month:02d}/{year} δεν υπάρχει στο ιστορικό.",
                           f"{month:02d}/{year} is not in the history."))
    total = 0.0
    for i in df.index[hit]:
        v = by_target.get(df.at[i, 'Chwra'])
        if v:
            df.loc[i, ['Logis', 'FnB', 'Extras', 'Esoda', 'Domatia']] = [str(round(x, 2)) for x in v]
            total += v[3]
    ensure_data_dir()
    df.to_csv(HISTORY_CSV, index=False, encoding='utf-8-sig')
    return year, month, total


def load_revenue_by_country(year):
    """{country: revenue} for one year, from the revenue saved with the history."""
    out = {}
    if not os.path.exists(HISTORY_CSV):
        return out
    try:
        df = pd.read_csv(HISTORY_CSV, encoding='utf-8-sig', dtype=str, keep_default_na=False)
    except Exception:
        return out
    if 'Esoda' not in df.columns:
        return out
    for _, r in df.iterrows():
        try:
            if int(r['Etos']) != year or str(r['Esoda']).strip() == '':
                continue
            c = pretty_country(str(r['Chwra']))
            out[c] = out.get(c, 0.0) + float(r['Esoda'])
        except (ValueError, KeyError):
            continue
    return out


def load_revenue_detail(year):
    """{country: {'total','rooms','fb','extras','roomn'}} for one year, from the revenue saved with the history."""
    out = {}
    if not os.path.exists(HISTORY_CSV):
        return out
    try:
        df = pd.read_csv(HISTORY_CSV, encoding='utf-8-sig', dtype=str, keep_default_na=False)
    except Exception:
        return out
    if 'Esoda' not in df.columns:
        return out
    for _, r in df.iterrows():
        try:
            if int(r['Etos']) != year or str(r['Esoda']).strip() == '':
                continue
            d = out.setdefault(pretty_country(str(r['Chwra'])),
                               {'total': 0.0, 'rooms': 0.0, 'fb': 0.0, 'extras': 0.0, 'roomn': 0.0})
            d['total'] += float(r['Esoda'])
            for k, col in (('rooms', 'Logis'), ('fb', 'FnB'), ('extras', 'Extras'), ('roomn', 'Domatia')):
                try:
                    d[k] += float(r.get(col, '') or 0)
                except ValueError:
                    pass
        except (ValueError, KeyError):
            continue
    return out


def load_revenue():
    """{(year, month): {'rooms','fb','extras','total'}} summed from the revenue columns of elstat_history.csv.
    Months saved before revenue was tracked are simply missing."""
    if not os.path.exists(HISTORY_CSV):
        return {}
    try:
        df = pd.read_csv(HISTORY_CSV, encoding='utf-8-sig', dtype=str, keep_default_na=False)
    except Exception:
        return {}
    if 'Esoda' not in df.columns:
        return {}
    out = {}
    for _, r in df.iterrows():
        try:
            y, m = int(r['Etos']), int(r['Minas'])
            if str(r['Esoda']).strip() == '':
                continue
            tot = float(r['Esoda'])
        except (ValueError, KeyError):
            continue
        d = out.setdefault((y, m), {'rooms': 0.0, 'fb': 0.0, 'extras': 0.0, 'total': 0.0, 'roomn': 0.0})
        d['total'] += tot
        for k, col in (('rooms', 'Logis'), ('fb', 'FnB'), ('extras', 'Extras'), ('roomn', 'Domatia')):
            try:
                d[k] += float(r.get(col, '') or 0)
            except ValueError:
                pass
    return out


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
    """Newest .xls/.xlsx (or Protel .csv) file in `folder` (ignores Excel's ~$ lock files), or None."""
    try:
        files = []
        for f in os.listdir(folder):
            if f.startswith('~$'):
                continue
            p, low = os.path.join(folder, f), f.lower()
            if low.endswith(('.xls', '.xlsx')) or (low.endswith('.csv') and os.path.isfile(p) and _csv_is_protel(p)):
                files.append(p)
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


# ------------------------------------------------------- template layout ("Smart Shift" mode)
# Default mode writes every value at fixed positions on top of the blank PDF.  "Smart Shift" mode
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


def delete_layout_reference():
    """Removes the saved reference layout (Smart Shift then asks for a new one). -> True when there was one."""
    try:
        os.remove(layout_file())
        return True
    except FileNotFoundError:
        return False


def save_layout_reference(template_pdf):
    snap = read_template_snapshot(template_pdf)
    snap['saved'] = datetime.datetime.now().isoformat(timespec='seconds')
    snap['template'] = os.path.basename(template_pdf)
    ensure_data_dir()
    _write_json(layout_file(), snap, ensure_ascii=False)
    return snap


def _prune_pairs(pairs, k=6, tol=15.0):
    """Drops pairs whose move disagrees with what the texts around them did (a wrongly paired repeated word,
    or a text that was put somewhere else). A whole section that moved together agrees with itself and stays."""
    if len(pairs) < 3 * k:
        return pairs
    vec = [(n[0] - r[0], n[1] - r[1]) for r, n in pairs]
    keep = []
    for i, (r, _n) in enumerate(pairs):
        near = sorted((j for j in range(len(pairs)) if j != i),
                      key=lambda j: (pairs[j][0][0] - r[0]) ** 2 + (pairs[j][0][1] - r[1]) ** 2)[:k]
        mdx, mdy = _median([vec[j][0] for j in near]), _median([vec[j][1] for j in near])
        if abs(vec[i][0] - mdx) <= tol and abs(vec[i][1] - mdy) <= tol:
            keep.append(pairs[i])
    return keep if len(keep) >= MIN_ANCHORS else pairs


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
        # texts that changed a little (a year, a typo, a split word) are still found: the most similar text close to
        # where the reference text is expected counts as the same text
        used_r = {r for r, _n in pairs}
        used_n = {n for _r, n in pairs}
        for t, x, y in r_items:
            if (x, y) in used_r or len(t) < 4:
                continue
            px, py = ax * x + bx, ay * y + by
            best = None
            for t2, x2, y2 in n_items:
                if (x2, y2) in used_n or abs(len(t2) - len(t)) > 3:
                    continue
                d = math.hypot(x2 - px, y2 - py)
                if d > DUP_TOLERANCE:
                    continue
                sim = difflib.SequenceMatcher(None, t, t2).ratio()
                if sim >= 0.8 and (best is None or (sim, -d) > best[0]):
                    best = ((sim, -d), (x2, y2))
            if best:
                pairs.append(((x, y), best[1]))
                used_r.add((x, y))
                used_n.add(best[1])
        found = len(pairs)
        pairs = _prune_pairs(pairs)             # unreliable pairs do not steer the page model or the local shifts
        # one overall shift/scale for the page, from every reliable pair
        mx = _fit_line([(r[0], n[0]) for r, n in pairs])
        my = _fit_line([(r[1], n[1]) for r, n in pairs])
        out.append({'anchors': [(r[0], r[1], n[0] - r[0], n[1] - r[1]) for r, n in pairs],
                    'model': (mx, my), 'matched': found, 'total': len(r_items)})
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
        # share of the reference texts that moved differently from the page as a whole (0 = the page moved as one piece)
        self.spread = [sum(1 for a in anchors if abs(a[2]) > self.AGREE or abs(a[3]) > self.AGREE) / len(anchors)
                       if anchors else 1.0 for anchors in self.pages]

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

    def confidence(self, page, x, y):
        """-> (score 0..1, why).  How far a position can be trusted: it needs reference texts around it that
        were found again AND that agree with each other about how far this part of the form moved."""
        anchors = self.pages[page]
        if not anchors:
            return 0.4, 'shift_far'
        ranked = sorted([((a[0] - x) ** 2 + (a[1] - y) ** 2, a) for a in anchors], key=lambda t: t[0])
        near = [a for _d, a in ranked[:self.NEAREST]]
        nearest = math.sqrt(ranked[0][0])
        mdx, mdy = _median([a[2] for a in near]), _median([a[3] for a in near])
        bad = sum(1 for a in near if abs(a[2] - mdx) > self.AGREE or abs(a[3] - mdy) > self.AGREE)
        close = sum(1 for d, _a in ranked if d <= 80.0 ** 2)
        # Being far from any text only matters when different parts of the page moved differently: a page that moved
        # as one piece is as trustworthy in an empty cell as next to a label.
        het = min(1.0, 0.15 + 6.0 * self.spread[page])
        penalties = []                      # (factor, reason)
        if nearest > 60.0:
            penalties.append((1.0 - (1.0 - max(0.55, 1.0 - (nearest - 60.0) / 250.0)) * het, 'shift_far'))
        elif close < 3:
            penalties.append((1.0 - 0.15 * het, 'shift_few'))
        if bad:
            penalties.append((1.0 - 0.6 * bad / len(near), 'shift_disagree'))
        dx, dy = self.offset(page, x, y)
        if math.hypot(dx, dy) > LARGE_SHIFT:
            penalties.append((0.9, 'shift_large'))
        score = 1.0
        for factor, _why in penalties:
            score *= factor
        worst = min(penalties) if penalties else (1.0, '')
        why = worst[1] if worst[0] < 0.97 else ''
        return score, why


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
    tl = TemplateLayout([m['anchors'] for m in matches], [pg['size'] for pg in new['pages']],
                        [m['model'] for m in matches])
    try:
        rep['confidence'] = smart_shift_confidence(tl, template_pdf, rep['identical'])
    except Exception:
        rep['confidence'] = None
    return tl, rep


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
    conf_text = describe_confidence(rep.get('confidence')) if rep.get('ok') else ''
    if conf_text:
        lines += ["", conf_text]
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
    ('vars', ("Μεταβλητές", "Variables")),
    ('extra', ("Πρόσθετα στοιχεία", "Extra items")),
    ('custom', ("Δικά μου κείμενα", "My own texts")),
]
# fields whose text has a fixed length can be split into one item per character
MANUAL_SPLIT_LEN = {'afm': 9, 'days': 2, 'employees': 3, 'year': 4}


# where the optional address values start (just after the printed label) and how wide they may be (points)
EXTRA_POS = {'nomos': (79.0, 581.0), 'dimotiki': (352.0, 581.0), 'odos': (72.0, 565.0),
             'arithmos': (388.0, 565.0), 'tk': (477.0, 565.0)}
EXTRA_WIDTH = {'nomos': 178.0, 'dimotiki': 205.0, 'odos': 262.0, 'arithmos': 68.0, 'tk': 80.0}


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
    add('nomos', 'hotel', ("Νομός", "Prefecture (ΝΟΜΟΣ)"), 0, *EXTRA_POS['nomos'], 10)
    add('dimotiki', 'hotel', ("Δημοτική ενότητα", "Municipal unit"), 0, *EXTRA_POS['dimotiki'], 10)
    add('odos', 'hotel', ("Οδός", "Street"), 0, *EXTRA_POS['odos'], 10)
    add('arithmos', 'hotel', ("Αριθμός", "Street number"), 0, *EXTRA_POS['arithmos'], 10)
    add('tk', 'hotel', ("Τ.Κ.", "Postal code"), 0, *EXTRA_POS['tk'], 10)
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
    try:
        width = float(saved['maxw'])
        if math.isfinite(width) and width > 0:
            base['maxw'] = width
    except (KeyError, TypeError, ValueError):
        pass
    if isinstance(saved.get('split'), bool):
        base['split'] = saved['split']
    if isinstance(saved.get('tail'), bool):
        base['tail'] = saved['tail']


def _layout_from_saved(data):
    """A layout made of saved data, or None when the data is not a layout.  Layouts saved by this version carry
    'known' (the program's own fields at the time): an own field that is not in them stays out - a blank canvas stays
    blank - and only fields added to the program later are filled in with their default position.  Older files
    (no 'known') are merged over the program's own positions, as before."""
    if not isinstance(data, dict):
        return None
    saved = data.get('items', {})
    if not isinstance(saved, dict):
        saved = {}
    defaults = manual_default_layout()
    known = data.get('known')
    if isinstance(known, list):
        layout = {'version': 1, 'items': {fid: it for fid, it in defaults['items'].items() if fid not in known}}
    else:
        layout = defaults
    items = layout['items']
    for fid, it in saved.items():
        if not isinstance(it, dict) or not isinstance(fid, str):
            continue
        if fid in defaults['items']:
            _merge_manual_item(items.setdefault(fid, copy.deepcopy(defaults['items'][fid])), it)
        elif it.get('custom') and fid.startswith('custom_'):
            items[fid] = base = manual_new_item(custom=True)
            _merge_manual_item(base, it)
        elif it.get('source') and fid.startswith('extra_'):
            items[fid] = base = manual_new_item(size=10.0)
            _merge_manual_item(base, it)
        elif it.get('parent') and '@' in fid:
            items[fid] = base = manual_new_item()
            _merge_manual_item(base, it)
    for fid in [f for f, i in items.items() if i.get('parent') and not items.get(i['parent'], {}).get('split')]:
        del items[fid]                      # a piece whose whole is not split any more
    for fid, i in items.items():
        if i.get('split') and not any(c.get('parent') == fid for c in items.values()):
            i['split'] = False
    layout['known'] = [f[0] for f in MANUAL_FIELDS]
    return layout


def load_manual_layout():
    """The saved manual layout (or the program's own positions when nothing was saved)."""
    try:
        with open(manual_layout_file(), encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        data = {}
    return _layout_from_saved(data) or manual_default_layout()


def save_manual_layout(layout):
    ensure_data_dir()
    _write_json(manual_layout_file(), {'version': 1, 'items': layout['items'], 'known': [fid for fid, *_r in MANUAL_FIELDS]},
                ensure_ascii=False, indent=1)


def delete_manual_layout():
    try:
        os.remove(manual_layout_file())
    except FileNotFoundError:
        pass


# ------------------------------------------------------- recognition confidence
# "The template was recognised" is too coarse: some parts of a form are found with certainty (a printed label with a
# drawn line next to it) and others only guessed (no line, a width that had to be estimated, a row code without its
# printed word).  Every field therefore gets a score from 0 to 1 and a reason, the scores are averaged per section of
# the form, and the fields that fall below CONF_OK are outlined in the preview so the person knows exactly what to look at.
CONF_OK = 0.90          # at or above: nothing to check
CONF_LOW = 0.70         # between CONF_LOW and CONF_OK: please check (amber); below: low (red)
CONF_SECTIONS = [
    ('hotel', ("Στοιχεία καταλύματος", "Hotel information")),
    ('dates', ("Πεδία ημερομηνιών", "Date fields")),
    ('country', ("Πίνακας χωρών", "Country table")),
    ('totals', ("Σύνολα", "Totals")),
]
CONF_BOX_W = {'name': 180.0, 'type': 120.0, 'category': 90.0, 'phone': 95.0, 'email': 125.0, 'company': 200.0,
              'year': 45.0, 'month': 60.0, 'nomos': 80.0, 'dimotiki': 90.0, 'odos': 110.0, 'arithmos': 40.0, 'tk': 45.0}
CONF_WHY = {
    'missing': ("Δεν βρέθηκε στο template - τοποθετήστε το στον επεξεργαστή", "Not found on the template - place it in the editor"),
    'ambiguous_label': ("Η ετικέτα εμφανίζεται περισσότερες από μία φορές", "The printed label appears more than once"),
    'border': ("Το πλάτος προήλθε από το τυπωμένο όριο του κελιού", "The width comes from the printed cell border"),
    'no_line': ("Δεν βρέθηκε γραμμή γραφής - η θέση εκτιμήθηκε", "No writing line found - the position is an estimate"),
    'open_width': ("Δεν βρέθηκε όριο πλάτους - το πλάτος εκτιμήθηκε", "No width limit found - the width is an estimate"),
    'narrow': ("Πολύ λίγος χώρος - το κείμενο ίσως δεν χωράει", "Very little room - the text may not fit"),
    'slashes': ("Οι θέσεις προήλθαν από τις κάθετες «/» και όχι από γραμμές", "Slots were taken from the “/” marks, not from drawn lines"),
    'box_count': ("Ο αριθμός των κουτιών δεν είναι ο αναμενόμενος", "The number of digit boxes is not what was expected"),
    'ambiguous_boxes': ("Βρέθηκαν πολλά πιθανά κουτιά", "Several candidate digit boxes were found"),
    'code_only': ("Βρέθηκε ο κωδικός γραμμής αλλά όχι η λέξη «Σύνολο» δίπλα του", "The row code was found, but not the word “Total” beside it"),
    'ambiguous_code': ("Ο κωδικός γραμμής εμφανίζεται περισσότερες από μία φορές", "The row code appears more than once"),
    'rows_partial': ("Δεν βρέθηκαν όλες οι γραμμές χωρών", "Not every country row was found"),
    'shift_few': ("Λίγα σημεία αναφοράς κοντά", "Few reference texts nearby"),
    'shift_far': ("Μακριά από κάθε σημείο αναφοράς", "Far from any reference text"),
    'shift_disagree': ("Τα γειτονικά σημεία αναφοράς διαφωνούν μεταξύ τους", "The nearby reference texts disagree with each other"),
    'shift_large': ("Μεγάλη μετατόπιση σε αυτό το σημείο", "Large shift at this spot"),
    'unrated': ("Δεν έχει αξιολογηθεί", "Not rated"),
}


def conf_section(fid):
    """The section of the form a field belongs to (one of CONF_SECTIONS)."""
    if fid.startswith('rows'):
        return 'country'
    if fid.startswith('tot_'):
        return 'totals'
    if fid.startswith(('date_from_', 'date_to_', 'sign_')):
        return 'dates'
    return 'hotel'


def conf_level(score, why=''):
    if why == 'missing':
        return 'missing'
    return 'ok' if score >= CONF_OK else 'check' if score >= CONF_LOW else 'low'


def conf_percent(score):
    """Whole percent, rounded DOWN: 100% is only shown when nothing at all is in doubt."""
    return 100 if score >= 0.9995 else max(0, min(99, int(score * 100)))


def conf_why(code):
    pair = CONF_WHY.get(code) or CONF_WHY['unrated']
    return T(*pair)


def conf_sections(scores):
    """{fid: score} -> [{'id', 'name', 'score', 'pct', 'level', 'count', 'weak': [fids]}] in form order."""
    out = []
    for sid, name in CONF_SECTIONS:
        ids = [f for f in scores if conf_section(f) == sid]
        if not ids:
            continue
        mean = sum(scores[f] for f in ids) / len(ids)
        out.append({'id': sid, 'name': name, 'score': mean, 'pct': conf_percent(mean), 'level': conf_level(mean),
                    'count': len(ids), 'weak': [f for f in ids if scores[f] < CONF_OK]})
    return out


def conf_item_box(fid, it, digits=None):
    """[x0, y0, x1, y1] in PDF points: the area a value will occupy (for the outline in the preview)."""
    size = float(it.get('size') or 10.0)
    step = float(it.get('step') or 0.0)
    if step:                                                 # one digit per box, centred in the first box
        w = step * (digits or MANUAL_SPLIT_LEN.get(fid) or 2)
        x0 = it['x'] - step / 2.0
    else:
        w = float(it.get('maxw') or CONF_BOX_W.get(fid, 70.0))
        x0 = it['x'] - w / 2.0 if it.get('align') == 'center' else it['x']
    return [round(x0, 1), round(it['y'] - 0.3 * size, 1), round(x0 + w, 1), round(it['y'] + size, 1)]


def _detect_confidence(wanted, items, found, conf, colx, row_ys, coverage, boxes):
    """The confidence of Smart Field Detection: {'fields': {fid: {...}}, 'sections': [...]}."""
    fields = {}
    for fid in wanted:
        if fid not in found:
            fields[fid] = {'score': 0.0, 'why': 'missing'}
            continue
        it = items[fid]
        # the table fields carry no rating of their own: they are as sure as the share of rows that was found
        c = {'score': 1.0, 'why': ''} if fid.startswith('rows') else conf.get(fid, {'score': 0.9, 'why': 'unrated'})
        score, why = c['score'], c['why']
        entry = {'page': it['page'], 'x': it['x'], 'y': it['y']}
        if fid.startswith('rows'):
            # the table is as sure as the share of country rows that were found again
            if coverage < score:
                score, why = coverage, 'rows_partial'
            if coverage < 0.9995:
                entry['flag'] = True            # even 95% means some countries get no value: always worth a look
            ys = list(row_ys[it['page']].values()) if it['page'] < len(row_ys) else []
            if ys:
                x0, x1 = colx.get((it['page'], fid[-1]), (it['x'] - 30.0, it['x'] + 30.0))
                entry['band'] = [round(x0, 1), round(min(ys) - 4, 1), round(x1, 1), round(max(ys) + 10, 1)]
        else:
            entry['box'] = conf_item_box(fid, it, boxes.get(fid))
        entry.update({'score': score, 'why': why})
        fields[fid] = entry
    return {'fields': fields, 'sections': conf_sections({f: v['score'] for f, v in fields.items()})}


def smart_shift_confidence(layout, template_pdf, identical=False):
    """The confidence of Smart Shift for every field, from the reference texts around the place it is moved to.
    `identical`: the template is byte-for-byte the reference, so nothing moved and everything is certain."""
    base = manual_default_layout()
    try:
        rows = detect_template_rows(template_pdf)
    except Exception:
        rows = [{}, {}]
    expected = len(MAPPED_TARGETS | {norm('ΕΛΛΑΔΑ')})
    coverage = min(1.0, (len(rows[0]) + len(rows[1])) / expected) if expected else 0.0
    fields = {}
    for fid, it in base['items'].items():
        page = it['page']
        if fid.startswith('rows'):
            ys = sorted(rows[page].values()) if page < len(rows) else []
            ry = ys[len(ys) // 2] if ys else 400.0
            score, why = layout.confidence(page, it['x'], ry)
            x = it['x'] + layout.offset(page, it['x'], ry)[0]
            if coverage < score:
                score, why = coverage, 'rows_partial'
            entry = {'page': page, 'x': round(x, 2), 'y': 0.0, 'score': score, 'why': why}
            if coverage < 0.9995:
                entry['flag'] = True
            if ys:
                entry['band'] = [round(x - 32.0, 1), round(ys[0] - 4, 1), round(x + 32.0, 1), round(ys[-1] + 10, 1)]
        else:
            x, y = layout.place(page, it['x'], it['y'])
            score, why = layout.confidence(page, it['x'], it['y'])
            placed = dict(it, x=round(x, 2), y=round(y, 2))
            entry = {'page': page, 'x': placed['x'], 'y': placed['y'], 'score': score, 'why': why,
                     'box': conf_item_box(fid, placed)}
        fields[fid] = entry
    if identical:
        for entry in fields.values():
            entry.update({'score': 1.0, 'why': ''})
            entry.pop('flag', None)
    return {'fields': fields, 'sections': conf_sections({f: v['score'] for f, v in fields.items()})}


def confidence_view(conf, layout=None):
    """What the screen needs: {'sections': [...], 'flags': [...]}, or None when there is no recognition to report.
    `layout` is the layout the preview was made with.  A position that the person has moved (or hidden, or placed by
    hand when it was not found) in the editor counts as checked, so fine-tuning raises the confidence."""
    if not conf or not conf.get('fields'):
        return None
    items = layout.get('items', {}) if layout else {}
    scores, flags = {}, []
    for fid, f in conf['fields'].items():
        score, why = f['score'], f.get('why', '')
        it = items.get(fid) if layout else None
        adjusted = False
        if it is not None and (score < CONF_OK or f.get('flag')):
            if 'x' not in f:                                   # was not found: placed by hand when it is visible now
                adjusted = bool(it.get('visible', True))
            elif not it.get('visible', True):                  # hidden on purpose: nothing will be written there
                adjusted = True
            else:
                moved_x = abs(it.get('x', 0.0) - f['x']) > 0.6
                moved_y = (not fid.startswith('rows')) and abs(it.get('y', 0.0) - f['y']) > 0.6
                adjusted = it.get('page', f['page']) != f['page'] or moved_x or moved_y
        scores[fid] = 1.0 if adjusted else score
        level = conf_level(score, why)
        if level == 'ok' and f.get('flag'):
            level = 'check'
        if level != 'ok' and not adjusted:
            flags.append({'fid': fid, 'label': manual_label(fid, {}, None), 'score': score, 'pct': conf_percent(score),
                          'level': level, 'why': why, 'section': conf_section(fid), 'page': f.get('page'),
                          'box': f.get('box'), 'band': f.get('band')})
    flags.sort(key=lambda f: (f['score'], f['label']))
    sections = conf_sections(scores)
    for sec in sections:
        sec['weak'] = [f['fid'] for f in flags if f['section'] == sec['id']]
    return {'sections': sections, 'flags': flags}


def describe_confidence(conf, layout=None):
    """The confidence of every section in words (current language), plus the least certain positions."""
    view = confidence_view(conf, layout)
    if not view:
        return ""
    lines = [T("Βεβαιότητα αναγνώρισης ανά ενότητα:", "Recognition confidence per section:")]
    for sec in view['sections']:
        extra = T(f" · {len(sec['weak'])} προς έλεγχο", f" · {len(sec['weak'])} to check") if sec['weak'] else ""
        lines.append(f"{'⚠' if sec['weak'] else '✓'} {T(*sec['name'])}: {sec['pct']}%{extra}")
    if view['flags']:
        lines += ["", T("Ελέγξτε αυτές τις θέσεις:", "Please check these positions:")]
        for f in view['flags'][:8]:
            lines.append(f"  – {f['label']}: {f['pct']}% ({conf_why(f['why'])})")
        if len(view['flags']) > 8:
            lines.append(T(f"  … και άλλες {len(view['flags']) - 8}", f"  … and {len(view['flags']) - 8} more"))
    return "\n".join(lines)


# ------------------------------------------------------- Smart Field Detection ("AI / Auto-Detect")
# The third way of filling the form.  Nothing is hardcoded for the program's own fields: the blank template is
# scanned on the fly (every word with its exact box, every drawn line), the printed labels are recognised, the
# table grid is aligned and a layout with a position for every field is computed from it.  The result is a
# normal layout (the same kind the layout editor works with), so it is drawn by the same code as "Set manually".
# Only what the user adds himself (own texts, variables, extra items) stays manual - it is made in the editor.
DETECT_LAYOUT_NAME = "elstat_detected_layout.json"
DETECT_MIN_FIELDS = 0.55        # share of the fields that must be found, otherwise the template is not recognised
DETECT_MIN_ROWS = 0.6           # ...and the same for the country rows
DETECT_GAP = 4.0                # points between a printed label and the value written after it
DETECT_LIFT = 2.5               # points a value sits above the line it is written on
DETECT_TOTAL_CODES = {'2.1': '05', '2.2': '10', '2.3': '20', '2': '25', '3': '35', '4': '45', '5': '55',
                      '6': '60', '8': '65', '9': '70'}      # the printed code of every total row of the form


class DetectError(Exception):
    """The template could not be scanned or recognised."""


def detect_layout_file():
    return os.path.join(DATA_DIR, DETECT_LAYOUT_NAME)


def detect_cache_saved():
    return os.path.isfile(detect_layout_file())


def _dn(s):
    """Comparison form of a printed text: normalised, without colons and dots."""
    return norm(s).replace(':', '').replace('.', '')


def _page_words(page):
    """[{'s','x0','x1','y0','y1','by'}] - every word of a page. y0/y1 = box of the word, by = baseline of its first letter."""
    tp = page.get_textpage()
    count = tp.count_chars()
    text = tp.get_text_range()
    words, cur = [], None
    for i in range(count):
        ch = text[i] if i < len(text) else ' '
        try:
            l, b, r, t = tp.get_charbox(i)
        except Exception:
            continue
        if ch.isspace() or r <= l:
            if cur:
                words.append(cur)
                cur = None
            continue
        if cur and (l - cur['x1'] > 0.9 * max(t - b, 2.0) or abs(b - cur['by']) > 2.6 or l < cur['x0'] - 1):
            words.append(cur)
            cur = None
        if cur is None:
            cur = {'s': ch, 'x0': l, 'x1': r, 'y0': b, 'y1': t, 'by': b}
        else:
            cur['s'] += ch
            cur['x1'] = max(cur['x1'], r)
            cur['y0'] = min(cur['y0'], b)
            cur['y1'] = max(cur['y1'], t)
    if cur:
        words.append(cur)
    return words


def _page_lines(page, pdfium_raw):
    """([{'y','x0','x1'}], [{'x','y0','y1'}]) - the straight horizontal and vertical lines drawn on a page."""
    hl, vl = [], []
    for obj in page.get_objects(filter=[pdfium_raw.FPDF_PAGEOBJ_PATH]):
        try:
            l, b, r, t = obj.get_bounds()
        except Exception:
            continue
        if t - b < 1.7 and r - l > 3:
            hl.append({'y': (b + t) / 2, 'x0': l + 0.5, 'x1': r - 0.5})
        elif r - l < 1.7 and t - b > 3:
            vl.append({'x': (l + r) / 2, 'y0': b + 0.5, 'y1': t - 0.5})
    return hl, vl


def _page_rects(page, pdfium_raw):
    """[(l, b, r, t)] - the drawn rectangles (cell and frame borders) of a page."""
    out = []
    try:
        for obj in page.get_objects(filter=[pdfium_raw.FPDF_PAGEOBJ_PATH]):
            try:
                l, b, r, t = obj.get_bounds()
            except Exception:
                continue
            if 6 <= r - l < 500 and 6 <= t - b < 400:
                out.append((l, b, r, t))
    except Exception:
        pass
    return out


def _border_right(page_geo, x, y):
    """x of the nearest printed vertical border to the right of x that spans height y (lines and rectangle edges)."""
    cands = [v['x'] for v in page_geo['v'] if v['x'] > x and v['y0'] - 1 <= y <= v['y1'] + 1]
    for (l, b, r, t) in page_geo.get('r', []):
        if b - 1 <= y <= t + 1:
            cands += [e for e in (l, r) if e > x]
    return min(cands) if cands else None


def _read_geometry(template_pdf):
    """Words and lines of the first 2 pages of a template."""
    try:
        import pypdfium2 as pdfium
        import pypdfium2.raw as pdfium_raw
    except Exception:
        raise DetectError(T("Η αυτόματη ανίχνευση χρειάζεται το pypdfium2 (pip install pypdfium2).",
                            "Auto-detection needs pypdfium2 (pip install pypdfium2)."))
    pages = []
    with _PDFIUM_LOCK:
        try:
            doc = _open_pdfium(pdfium, template_pdf)
        except Exception as e:
            raise DetectError(T(f"Δεν μπόρεσα να διαβάσω το template: {e}", f"I could not read the template: {e}"))
        try:
            if len(doc) < 2:
                raise DetectError(T("Το template PDF πρέπει να έχει 2 σελίδες.", "The template PDF must have 2 pages."))
            for pi in range(2):
                page = doc[pi]
                width, height = page.get_size()
                hl, vl = _page_lines(page, pdfium_raw)
                pages.append({'size': (float(width), float(height)), 'words': _page_words(page), 'h': hl, 'v': vl,
                              'r': _page_rects(page, pdfium_raw)})
        except DetectError:
            raise
        except Exception as e:
            raise DetectError(T(f"Δεν μπόρεσα να σαρώσω το template: {e}", f"I could not scan the template: {e}"))
        finally:
            try:
                doc.close()
            except Exception:
                pass
    return pages


def _find_phrase(page, phrase, max_y=None, min_y=None, near_y=None, prefer_line=False):
    """The first run of consecutive words of one text line that spells `phrase` (compared in a normalised form).
    -> {'x0','x1','y','words'} or None.  max_y/min_y limit where the phrase may be, near_y prefers the closest one."""
    target = _dn(phrase)
    words = sorted(page['words'], key=lambda w: (-round(w['by'] / 2.0), w['x0']))
    found = []
    for i, w in enumerate(words):
        acc = ''
        for j in range(i, len(words)):
            wj = words[j]
            if abs(wj['by'] - w['by']) > 2.6 or (j > i and wj['x0'] < words[j - 1]['x0']):
                break
            acc += _dn(wj['s'])
            if acc == target:
                y = w['by']
                if (max_y is None or y < max_y) and (min_y is None or y > min_y):
                    found.append({'x0': w['x0'], 'x1': wj['x1'], 'y': y, 'words': words[i:j + 1]})
                break
            if not target.startswith(acc):
                break
    if not found:
        return None
    if prefer_line and len(found) > 1:
        # deep pass: of several identical printed labels, the one with a drawn writing line after it is the real one
        lined = [m for m in found if any(0 <= h['x0'] - m['x1'] <= 40 for h in _underlines(page, m['y'], m['x1'], tol=4.5))]
        if lined:
            found = lined
    best = min(found, key=lambda m: abs(m['y'] - near_y)) if near_y is not None else found[0]
    best['n'] = len(found)          # more than one candidate = the label is ambiguous (lowers the confidence)
    return best


def _find_word_starting(page, prefix):
    """First word whose normalised text starts with `prefix` -> the word dict, or None."""
    p = _dn(prefix)
    for w in sorted(page['words'], key=lambda w: (-round(w['by'] / 2.0), w['x0'])):
        if _dn(w['s']).startswith(p):
            return w
    return None


def _underlines(page, y, x_from, x_to=None, tol=4.5):
    """Horizontal lines at height `y` (+-tol) that start at or after x_from, left to right."""
    out = [h for h in page['h'] if abs(h['y'] - y) <= tol and h['x1'] > x_from + 2 and (x_to is None or h['x0'] < x_to)]
    return sorted(out, key=lambda h: h['x0'])


def _box_groups(page):
    """Rows of equally spaced vertical lines of the same height (the digit boxes of the form).
    -> [{'x0': first line, 'pitch', 'n': number of boxes, 'y0', 'y1'}]"""
    short = [v for v in page['v'] if 5 < v['y1'] - v['y0'] < 30]
    short.sort(key=lambda v: (round(v['y0'] / 1.5), v['x']))
    clusters = []
    for v in short:
        for c in clusters:
            if abs(c[0]['y0'] - v['y0']) < 1.2 and abs(c[0]['y1'] - v['y1']) < 1.2:
                c.append(v)
                break
        else:
            clusters.append([v])
    groups = []
    for c in clusters:
        c.sort(key=lambda v: v['x'])
        run = [c[0]]
        for v in c[1:]:
            if len(run) < 2 or abs((v['x'] - run[-1]['x']) - (run[1]['x'] - run[0]['x'])) < 1.2:
                if len(run) == 1 and not 6 < v['x'] - run[0]['x'] < 30:
                    run = [v]
                    continue
                run.append(v)
            else:
                if len(run) >= 2:
                    groups.append(run)
                run = [v]
        if len(run) >= 2:
            groups.append(run)
    return [{'x0': r[0]['x'], 'pitch': (r[-1]['x'] - r[0]['x']) / (len(r) - 1), 'n': len(r) - 1,
             'y0': r[0]['y0'], 'y1': r[0]['y1']} for r in groups]


def _table_columns(page):
    """The columns of the country table of one page: {'name': (x0,x1), 'code': ..., 'clients': ..., 'nights': ...}
    taken from the long vertical lines and from where the headers PELATON / DIANYKTEREYSEON are printed."""
    xs = sorted({round(v['x'], 1) for v in page['v'] if v['y1'] - v['y0'] > 150})
    merged = []
    for x in xs:
        if not merged or x - merged[-1] > 2.0:
            merged.append(x)
    cols = list(zip(merged, merged[1:]))
    cp, cn = _find_word_starting(page, 'ΠΕΛΑΤΩΝ'), _find_word_starting(page, 'ΔΙΑΝΥΚΤΕΡΕΥ')
    if not cols or cp is None or cn is None:
        return None

    def holding(word):
        mid = (word['x0'] + word['x1']) / 2.0
        return next((i for i, (a, b) in enumerate(cols) if a <= mid <= b), None)
    ic, inn = holding(cp), holding(cn)
    if ic is None or inn is None or ic < 2 or inn <= ic:
        return None
    return {'name': cols[ic - 2], 'code': cols[ic - 1], 'clients': cols[ic], 'nights': cols[inn]}


def _page_rows(page, columns, targets, relaxed=False):
    """{target: baseline} of every country row found in the name column of a page."""
    x0, x1 = columns['name']
    sl, sr = (6, 30) if relaxed else (2, 6)
    inside = [w for w in page['words'] if w['x0'] >= x0 - sl and w['x1'] <= x1 + sr]
    inside.sort(key=lambda w: (-round(w['by'] / 2.0), w['x0']))
    lines, cur = [], []
    for w in inside:
        if cur and abs(w['by'] - cur[0]['by']) > 2.6:
            lines.append(cur)
            cur = []
        cur.append(w)
    if cur:
        lines.append(cur)
    rows = {}
    for line in lines:
        line.sort(key=lambda w: w['x0'])
        text = ''.join(norm(w['s']) for w in line)
        hit = text if text in targets else next((t for t in sorted(targets, key=len, reverse=True)
                                                 if text.startswith(t)), None)
        if relaxed and not hit:
            # deep pass: a name with a number or a stray mark around it, or only the first line of a wrapped name
            hit = next((t for t in sorted(targets, key=len, reverse=True) if len(t) >= 5 and t in text), None)
            if not hit and len(text) >= 6:
                starts = [t for t in targets if t.startswith(text)]
                hit = starts[0] if len(starts) == 1 else None
        if hit and hit not in rows:
            rows[hit] = round(line[0]['by'], 2)
    return rows


def _item(page, x, y, size, align='left', step=0.0, maxw=None):
    it = manual_new_item(page, round(float(x), 2), round(float(y), 2), size, step)
    it['align'] = align
    if maxw:
        it['maxw'] = round(float(maxw), 1)
    return it


def _after_label(page_geo, page, label, size, fields_out, fid, gap=DETECT_GAP, rate=None, deep=False):
    """A value written after a printed label (and on its line, if one is drawn there).
    `rate(fid, score, why)` is told how sure the placement is: a drawn line is the best evidence,
    a neighbouring text that limits the width is good, a guessed width is weak."""
    start = label['x1'] + gap
    limit = None
    under = [h for h in _underlines(page_geo, label['y'], label['x1'], tol=5.0 if deep else 3.5)]
    if deep and under:                       # only the line closest to the printed baseline counts
        dy = min(abs(h['y'] - label['y']) for h in under)
        under = [h for h in under if abs(h['y'] - label['y']) - dy < 0.8]
    if under:
        first = under[0]
        if first['x0'] > start and first['x0'] - label['x1'] <= (45 if deep else 25):
            start = first['x0'] + 3.0
        if first['x0'] <= start + 6:
            limit = first['x1']
    anchored = limit is not None
    border = None
    if deep and not anchored:                # no writing line: the printed border of the cell is the next best evidence
        bx = _border_right(page_geo, start + 12, label['y'])
        if bx is not None and bx - start < 260:
            border = limit = bx - 3.0
    rivals = [w['x0'] for w in page_geo['words'] if abs(w['by'] - label['y']) < 3 and w['x0'] > start + 8]
    if rivals:
        limit = min(limit, min(rivals) - 6) if limit else min(rivals) - 6
    bounded = limit is not None
    limit = limit or min(page_geo['size'][0] - 40, start + 200)
    fields_out[fid] = _item(page, start, label['y'] + DETECT_LIFT, size, maxw=max(limit - start, 20))
    if rate:
        score, why = ((1.0, '') if anchored else (0.96, 'border') if border else
                      (0.9, 'no_line') if bounded else (0.8, 'open_width'))
        if limit - start < 40:
            score, why = min(score, 0.7), 'narrow'
        if label.get('n', 1) > 1:
            score *= 0.93
            why = why or 'ambiguous_label'
        rate(fid, score, why)
    return True


def _slot_item(page, slot, y, size):
    return _item(page, (slot[0] + slot[1]) / 2.0, y, size, align='center', maxw=max(slot[1] - slot[0] - 1, 8))


def _date_slots(page_geo, label, deep=False):
    """Three writing slots (day / month / year) of a date row: the drawn underlines, else the gaps between the slashes."""
    for tol in ((4.5, 7.0) if deep else (4.5,)):
        under = _underlines(page_geo, label['y'], label['x1'], tol=tol)
        under = [u for u in under if u['x1'] - u['x0'] < 90]
        if deep and under:
            dy = min(abs(u['y'] - label['y']) for u in under)
            under = [u for u in under if abs(u['y'] - label['y']) - dy < 0.8]
        under = under[:3]
        if len(under) == 3:
            return [(u['x0'], u['x1']) for u in under], under[0]['y'], 'lines'
    slashes = sorted([w for w in page_geo['words'] if w['s'] == '/' and abs(w['by'] - label['y']) < 3
                      and w['x0'] > label['x1']], key=lambda w: w['x0'])[:2]
    if len(slashes) == 2:
        return [(label['x1'] + 3, slashes[0]['x0'] - 2), (slashes[0]['x1'] + 2, slashes[1]['x0'] - 2),
                (slashes[1]['x1'] + 2, slashes[1]['x1'] + 34)], label['y'] - 0.6, 'slashes'
    return None, None, None


def _detect_core(template_pdf, geo, deep=False):
    """Scans the blank template and computes a position for every field of the form.
    deep=True adds the extra-evidence checks of the deep pass (see detect_template_fields).
    -> (layout, report).  Raises DetectError when the template cannot be read or is not recognised safely.
    report = {'ok', 'found': [ids], 'missing': [ids], 'rows': {'found', 'expected', 'missing': [targets]},
              'row_ys': [{target: y}, {target: y}], 'sha1', 'sizes'}"""
    with open(template_pdf, 'rb') as f:
        sha1 = hashlib.sha1(f.read()).hexdigest()
    p0, p1 = geo[0], geo[1]
    out, how = {}, {}
    conf = {}                       # fid -> {'score': 0..1, 'why': code}: how sure each placement is

    def find(page, phrase, **kw):
        return _find_phrase(page, phrase, prefer_line=deep, **kw)

    def got(fid):
        how[fid] = True

    def rate(fid, score, why=''):
        conf[fid] = {'score': max(0.0, min(1.0, float(score))), 'why': why}

    def lab(label):                 # a printed label that appears more than once on the page is ambiguous
        return (1.0, '') if label.get('n', 1) <= 1 else (0.93, 'ambiguous_label')

    def rate_slot(fid, slot, suffix, how_, label, consistent=False):
        # slots from the "/" marks are more believable when the day and month slots are equally wide (deep pass)
        score, why = (1.0, '') if how_ == 'lines' else ((0.93, 'slashes') if consistent else (0.85, 'slashes'))
        if slot[1] - slot[0] < (22.0 if suffix == 'y' else 11.0):
            score, why = min(score, 0.7), 'narrow'
        s2, w2 = lab(label)
        rate(fid, score * s2, why or w2)

    # ---- survey year and month: written on the line after their labels
    year = find(p0, 'Έτος έρευνας')
    month = find(p0, 'Μήνας', near_y=year['y'] if year else None)
    for fid, label in (('year', year), ('month', month)):
        if label:
            under = [u for u in _underlines(p0, label['y'], label['x1'] - 1, tol=6.0 if deep else 4.5)
                     if u['x0'] - label['x1'] < (60 if deep else 45)]
            if under:
                out[fid] = _slot_item(0, (under[0]['x0'], under[0]['x1']), under[0]['y'] + DETECT_LIFT - 0.3, 11.0)
                got(fid)
                rate(fid, lab(label)[0], lab(label)[1])
            else:
                out[fid] = _item(0, label['x1'] + 12, label['y'] + DETECT_LIFT, 11.0)
                got(fid)
                rate(fid, 0.7 * lab(label)[0], 'no_line')

    # ---- the hotel lines: a value after each printed label
    type_label = find(p0, 'Είδος καταλύματος')
    ceiling = type_label['y'] + 2 if type_label else None
    for fid, phrase, size in (('type', 'Είδος καταλύματος', 11.0), ('category', 'Κατηγορία', 11.0),
                              ('name', 'Τίτλος καταλύματος', 11.0), ('nomos', 'Νομός', 10.0),
                              ('dimotiki', 'Δημοτική ενότητα', 10.0), ('odos', 'Οδός', 10.0),
                              ('arithmos', 'Αριθμός', 10.0), ('tk', 'ΤΚ', 10.0),
                              ('phone', 'Τηλέφωνα', 11.0)):
        label = find(p0, phrase, max_y=ceiling + 1 if fid != 'type' and ceiling else None,
                             near_y=type_label['y'] if (type_label and fid in ('type', 'category')) else None)
        if fid == 'phone':
            label = find(p0, phrase, max_y=type_label['y'] - 1 if type_label else None)
        if label and _after_label(p0, 0, label, size, out, fid, rate=rate, deep=deep):
            got(fid)
    mail = None
    for w in p0['words']:
        if 'mail' in w['s'].lower() or _dn(w['s']).endswith('ΜΑΙΛ'):
            if type_label is None or w['by'] < type_label['y']:
                mail = {'x0': w['x0'], 'x1': w['x1'], 'y': w['by']}
                break
    if mail and _after_label(p0, 0, mail, 11.0, out, 'email', rate=rate, deep=deep):
        got('email')

    # ---- the company: written under its printed label, inside the panel that is drawn there
    comp = find(p0, 'Επωνυμία επιχείρησης')
    if comp:
        cx = (comp['x0'] + comp['x1']) / 2.0
        below = sorted([h for h in p0['h'] if h['y'] < comp['y'] - 8 and h['x0'] <= cx <= h['x1']
                        and h['x1'] - h['x0'] < 250], key=lambda h: -h['y'])
        if below:
            panel = below[0]
            out['company'] = _item(0, (panel['x0'] + panel['x1']) / 2.0, panel['y'] + 0.3 * (comp['y'] - panel['y']),
                                   11.0, align='center', maxw=panel['x1'] - panel['x0'] - 14)
            got('company')
            roomy = panel['x1'] - panel['x0'] - 14 >= 60
            rate('company', (1.0 if roomy else 0.7) * lab(comp)[0], '' if roomy else 'narrow')

    # ---- digit boxes: VAT number, operating days, employees
    groups = _box_groups(p0)

    def put_boxes(fid, group, size=11.0, cands=1, label=None):
        h = group['y1'] - group['y0']
        y = group['y0'] + (h - 0.72 * size) / 2.0 + 0.6
        out[fid] = _item(0, group['x0'] + group['pitch'] / 2.0, y, size, align='center', step=group['pitch'])
        got(fid)
        how[fid] = group['n']
        expected = MANUAL_SPLIT_LEN.get(fid)        # the form prints a fixed number of boxes for these
        score, why = (1.0, '') if expected in (None, group['n']) else (0.6, 'box_count')
        if cands > 1:
            score *= 0.93
            why = why or 'ambiguous_boxes'
        if label:
            score *= lab(label)[0]
        rate(fid, score, why)

    def exact_count(fid, cand):             # deep pass: of several box rows, prefer the one with the printed number of boxes
        if deep and MANUAL_SPLIT_LEN.get(fid):
            exact = [g for g in cand if g['n'] == MANUAL_SPLIT_LEN[fid]]
            return exact or cand
        return cand

    def nearest_only(cand, dist):           # deep pass: one box row clearly closest to its label is not "ambiguous"
        if deep and len(cand) > 1:
            ranked = sorted(cand, key=dist)
            if dist(ranked[1]) > 2.0 * max(dist(ranked[0]), 5.0):
                return ranked[:1]
        return cand

    afm = find(p0, 'ΑΦΜ')
    if afm:
        cand = [g for g in groups if g['y0'] - 4 <= afm['y'] <= g['y1'] + 4 and g['x0'] >= afm['x1'] - 2]
        cand = nearest_only(exact_count('afm', cand), lambda g: abs(g['x0'] - afm['x1']))
        if cand:
            put_boxes('afm', min(cand, key=lambda g: g['x0']), cands=len(cand), label=afm)
    for fid, phrase in (('days', 'Ημέρες λειτουργίας'), ('employees', 'Σύνολο εργαζομένων')):
        label = find(p0, phrase)
        if label:
            cx = (label['x0'] + label['x1']) / 2.0
            cand = [g for g in groups if g['y1'] < label['y'] and abs(g['x0'] + g['pitch'] * g['n'] / 2.0 - cx) < 80]
            cand = nearest_only(exact_count(fid, cand), lambda g: abs(g['x0'] + g['pitch'] * g['n'] / 2.0 - cx) + (label['y'] - g['y1']))
            if cand:
                put_boxes(fid, max(cand, key=lambda g: g['y1']), cands=len(cand), label=label)

    # ---- the period of operation: day / month / year slots of the "From" and "To" rows
    period = find(p0, 'Περίοδος λειτουργίας')
    for prefix, phrase in (('date_from', 'Από'), ('date_to', 'Έως')):
        label = find(p0, phrase, max_y=period['y'] if period else None, near_y=(period['y'] - 50) if period else None)
        slots, base, how_ = _date_slots(p0, label, deep) if label else (None, None, None)
        if slots:
            cons = deep and abs((slots[0][1] - slots[0][0]) - (slots[1][1] - slots[1][0])) < 2.5
            for suffix, slot in zip(('d', 'm', 'y'), slots):
                out[f'{prefix}_{suffix}'] = _slot_item(0, slot, base + 1.7, 11.0)
                got(f'{prefix}_{suffix}')
                rate_slot(f'{prefix}_{suffix}', slot, suffix, how_, label, cons)

    # ---- signature date (page 2)
    sign = find(p1, 'Ημερομηνία')
    slots, base, how_ = _date_slots(p1, sign, deep) if sign else (None, None, None)
    if slots:
        cons = deep and abs((slots[0][1] - slots[0][0]) - (slots[1][1] - slots[1][0])) < 2.5
        for suffix, slot in zip(('d', 'm', 'y'), slots):
            out[f'sign_{suffix}'] = _slot_item(1, slot, base + 1.4, 10.0)
            got(f'sign_{suffix}')
            rate_slot(f'sign_{suffix}', slot, suffix, how_, sign, cons)

    # ---- the country table: columns, total rows, country rows
    cols = [_table_columns(p0), _table_columns(p1)]
    targets = set(MAPPED_TARGETS) | {norm('ΕΛΛΑΔΑ')}
    row_ys = [{}, {}]
    colx = {}                       # (page, 'c'|'n') -> x range of the column, for highlighting
    for pi, (pg, col) in enumerate(zip(geo, cols)):
        if col is None:
            continue
        for kind, size, tag in (('clients', 9.0, 'c'), ('nights', 9.0, 'n')):
            mid = (col[kind][0] + col[kind][1]) / 2.0
            out[f'rows{pi}_{tag}'] = _item(pi, mid, 0.0, size, align='center')
            got(f'rows{pi}_{tag}')
            colx[(pi, tag)] = col[kind]
        row_ys[pi] = _page_rows(pg, col, targets, relaxed=deep)
        cx0, cx1 = col['code']
        for key, code in DETECT_TOTAL_CODES.items():
            if (key == '2.1') != (pi == 0):
                continue
            cand = [w for w in pg['words'] if w['s'] == code and cx0 - 1 <= w['x0'] and w['x1'] <= cx1 + 1]
            if not cand:
                continue
            if deep and len(cand) > 1:
                tot = [c for c in cand if 'ΣΥΝΟΛΟ' in ''.join(norm(w['s']) for w in pg['words']
                                                           if abs(w['by'] - c['by']) < 11 and w['x1'] <= cx0 + 1)]
                cand = tot or cand
            ys = round(cand[0]['by'], 2)
            # second witness: the word "Σύνολο" printed to the left of the code, on the same line
            left = sorted((w for w in pg['words'] if abs(w['by'] - cand[0]['by']) < 4.5 and w['x1'] <= cx0 + 1),
                          key=lambda w: w['x0'])
            tscore, twhy = (1.0, '') if 'ΣΥΝΟΛΟ' in ''.join(norm(w['s']) for w in left) else (0.88, 'code_only')
            if deep and twhy == 'code_only':
                # deep pass: look further around the code for the word, and for the ruling of a table row
                near = sorted((w for w in pg['words'] if abs(w['by'] - cand[0]['by']) < 11 and w['x1'] <= cx0 + 1),
                              key=lambda w: (-round(w['by'] / 2.0), w['x0']))
                above = any(0 < h['y'] - cand[0]['by'] < 14 and h['x0'] <= cx0 + 2 and h['x1'] >= cx1 - 2 for h in pg['h'])
                below = any(0 < cand[0]['by'] - h['y'] < 14 and h['x0'] <= cx0 + 2 and h['x1'] >= cx1 - 2 for h in pg['h'])
                if 'ΣΥΝΟΛΟ' in ''.join(norm(w['s']) for w in near):
                    tscore, twhy = 0.99, ''
                elif above and below:
                    tscore, twhy = 0.96, ''
            if len(cand) > 1:
                tscore *= 0.93
                twhy = twhy or 'ambiguous_code'
            for kind, tag in (('clients', 'c'), ('nights', 'n')):
                mid = (col[kind][0] + col[kind][1]) / 2.0
                out[f'tot_{key}_{tag}'] = _item(pi, mid, ys, 10.0, align='center')
                got(f'tot_{key}_{tag}')
                rate(f'tot_{key}_{tag}', tscore, twhy)

    # ---- deep pass, cross-check: fields printed on the same row of the form must sit on the same baseline
    if deep:
        for group in (('nomos', 'dimotiki'), ('odos', 'arithmos', 'tk'), ('type', 'category'),
                      ('date_from_d', 'date_from_m', 'date_from_y'), ('date_to_d', 'date_to_m', 'date_to_y'),
                      ('sign_d', 'sign_m', 'sign_y')):
            ids = [f for f in group if f in out]
            if len(ids) < 2:
                continue
            med = sorted(out[f]['y'] for f in ids)[len(ids) // 2]
            agree = [f for f in ids if abs(out[f]['y'] - med) <= 0.8]
            if len(agree) < 2:
                continue
            for f in ids:
                c = conf.get(f)
                if f in agree:
                    if c and c['why'] == 'ambiguous_label':     # an independent witness removes the doubt
                        rate(f, min(1.0, c['score'] / 0.93 * 0.99), '')
                elif abs(out[f]['y'] - med) <= 6.0 and not f.startswith(('date_', 'sign_')):
                    out[f]['y'] = med                           # the odd one out is pulled onto the shared baseline

    # ---- what was found, what was not
    wanted = [f[0] for f in MANUAL_FIELDS]
    found = [f for f in wanted if f in out]
    missing = [f for f in wanted if f not in out]
    expected_rows = len(targets)
    found_rows = sum(len(r) for r in row_ys)
    rows_missing = sorted(t for t in targets if not any(t in r for r in row_ys))
    report = {'ok': True, 'found': found, 'missing': missing, 'sha1': sha1, 'row_ys': row_ys,
              'rows': {'found': found_rows, 'expected': expected_rows, 'missing': rows_missing},
              'sizes': [g['size'] for g in geo], 'boxes': {k: v for k, v in how.items() if isinstance(v, int) and v is not True}}
    if len(found) < DETECT_MIN_FIELDS * len(wanted) or found_rows < DETECT_MIN_ROWS * expected_rows:
        report['ok'] = False
        raise DetectError(T("Το template δεν αναγνωρίζεται με ασφάλεια: βρέθηκαν "
                            f"{len(found)} από {len(wanted)} πεδία και {found_rows} από {expected_rows} γραμμές χωρών. "
                            "Ίσως δεν είναι το σωστό έντυπο.",
                            "The template cannot be recognised safely: "
                            f"{len(found)} of {len(wanted)} fields and {found_rows} of {expected_rows} country rows were found. "
                            "It may not be the right form."))
    layout = manual_default_layout()
    items = {}
    for fid in wanted:
        if fid in out:
            items[fid] = out[fid]
        else:                                           # not found: parked and hidden, to be placed by hand
            base = layout['items'][fid]
            parked = manual_new_item(base['page'], geo[base['page']]['size'][0] / 2.0, geo[base['page']]['size'][1] / 2.0,
                                     base['size'], base['step'])
            parked['visible'] = False
            items[fid] = parked
    layout = {'version': 1, 'items': items, 'known': wanted}
    coverage = found_rows / expected_rows if expected_rows else 0.0
    report['confidence'] = _detect_confidence(wanted, items, out, conf, colx, row_ys, coverage, report['boxes'])
    return layout, report


DETECT_DEEP_NAME = "elstat_detect_deep.json"


def detect_deep_file():
    return os.path.join(DATA_DIR, DETECT_DEEP_NAME)


def _file_sha1(path):
    try:
        with open(path, 'rb') as f:
            return hashlib.sha1(f.read()).hexdigest()
    except OSError:
        return ''


def detect_deep_saved(template_pdf):
    """True when the person pressed Rerun Auto-Detection for this very template: the deep pass is used from then on."""
    try:
        with open(detect_deep_file(), encoding='utf-8') as f:
            data = json.load(f)
        sha1 = data.get('sha1', '')
        return bool(sha1) and sha1 == _file_sha1(template_pdf)
    except Exception:
        return False


def set_detect_deep(template_pdf, on=True):
    if on:
        ensure_data_dir()
        _write_json(detect_deep_file(), {'version': 1, 'sha1': _file_sha1(template_pdf),
                                         'saved': datetime.datetime.now().isoformat(timespec='seconds')})
    else:
        try:
            os.remove(detect_deep_file())
        except FileNotFoundError:
            pass


def _overall_score(conf):
    fields = (conf or {}).get('fields') or {}
    return sum(f['score'] for f in fields.values()) / len(fields) if fields else 0.0


def _merge_detections(bl, br, dl, dr):
    """Field by field, whichever pass is surer wins - so the deep pass can only raise the confidence."""
    wanted = [f[0] for f in MANUAL_FIELDS]
    bf, df = br['confidence']['fields'], dr['confidence']['fields']
    deep_rows = dr['rows']['found'] > br['rows']['found']
    rows_rep = dr if deep_rows else br
    items, fields, boxes, improved = {}, {}, {}, []
    for fid in wanted:
        b, d = bf[fid], df[fid]
        if fid.startswith('rows'):
            take = deep_rows
        else:
            take = d['score'] > b['score'] + 1e-9
        src_l, src_f, src_r = (dl, df, dr) if take else (bl, bf, br)
        items[fid], fields[fid] = src_l['items'][fid], src_f[fid]
        if fid in src_r['boxes']:
            boxes[fid] = src_r['boxes'][fid]
        if fields[fid]['score'] > b['score'] + 1e-9:
            improved.append({'fid': fid, 'before': b['score'], 'after': fields[fid]['score'],
                             'why': b.get('why', ''), 'now': fields[fid].get('why', '')})
    rep = dict(br)
    rep['found'] = [f for f in wanted if fields[f].get('why') != 'missing']
    rep['missing'] = [f for f in wanted if fields[f].get('why') == 'missing']
    rep['row_ys'], rep['rows'], rep['boxes'] = rows_rep['row_ys'], rows_rep['rows'], boxes
    rep['confidence'] = {'fields': fields, 'sections': conf_sections({f: v['score'] for f, v in fields.items()})}
    rep['deep'] = {'before': _overall_score(br['confidence']), 'after': _overall_score(rep['confidence']),
                   'improved': improved}
    return {'version': 1, 'items': items, 'known': wanted}, rep


def detect_template_fields(template_pdf, deep=None):
    """Smart Field Detection of a blank template -> (layout, report).
    The standard pass is always made.  When `deep` is true (None = the person pressed Rerun Auto-Detection for this
    template before) a second pass looks for extra evidence - the writing line that really belongs to a repeated label,
    printed cell borders, a wider search for lines, the printed number of digit boxes, the ruling of total rows,
    looser matching of country names - and the surer result of the two is kept for every field."""
    if deep is None:
        deep = detect_deep_saved(template_pdf)
    geo = _read_geometry(template_pdf)
    try:
        bl, br = _detect_core(template_pdf, geo, False)
    except DetectError:
        if not deep:
            raise
        bl = br = None
        first_error = sys.exc_info()[1]
    if not deep:
        return bl, br
    try:
        dl, dr = _detect_core(template_pdf, geo, True)
    except DetectError:
        if br is None:
            raise first_error
        br['deep'] = {'before': _overall_score(br['confidence']), 'after': _overall_score(br['confidence']), 'improved': []}
        return bl, br
    if br is None:
        dr['deep'] = {'before': 0.0, 'after': _overall_score(dr['confidence']), 'improved': []}
        return dl, dr
    return _merge_detections(bl, br, dl, dr)


def describe_detect_report(rep, tuned=False):
    """The report of detect_template_fields() in words (current language)."""
    n_found, n_missing = len(rep['found']), len(rep['missing'])
    lines = [T(f"✓ Βρέθηκαν {n_found} από {n_found + n_missing} πεδία του δελτίου.",
               f"✓ {n_found} of {n_found + n_missing} fields of the form were found.")]
    rows = rep['rows']
    lines.append(T(f"Γραμμές χωρών: {rows['found']} από {rows['expected']}.",
                   f"Country rows: {rows['found']} of {rows['expected']}."))
    if rep['missing']:
        names = ", ".join(manual_label(f, {}, None) for f in rep['missing'])
        lines.append(T(f"⚠ Δεν βρέθηκαν (κρύφτηκαν - τοποθετήστε τα στον επεξεργαστή): {names}",
                       f"⚠ Not found (hidden - place them in the editor): {names}"))
    if rows['missing']:
        lines.append(T("⚠ Χώρες χωρίς γραμμή στο template: ", "⚠ Countries without a row in the template: ")
                     + ", ".join(_pretty_target(t) for t in rows['missing']))
    deep = rep.get('deep')
    if deep:
        a, b = conf_percent(deep['before']), conf_percent(deep['after'])
        lines.append(T(f"⚡ Βαθιά ανίχνευση: εμπιστοσύνη {a}% → {b}%.", f"⚡ Deep detection: confidence {a}% → {b}%."))
    if tuned:
        lines.append(T("✎ Χρησιμοποιούνται οι προσαρμογές σας από τον επεξεργαστή.",
                       "✎ Your fine-tuning from the editor is in use."))
    return "\n".join(lines)


# ---- the cache of a fine-tuned detected layout (it lives only while Smart Field Detection stays selected)
def load_detect_cache():
    """The fine-tuned detected layout, or None."""
    try:
        with open(detect_layout_file(), encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        return None
    layout = _layout_from_saved(data)
    if layout is not None:
        layout['sha1'] = data.get('sha1', '')
    return layout


def save_detect_cache(layout, template_pdf=None):
    ensure_data_dir()
    sha1 = ''
    if template_pdf:
        try:
            with open(template_pdf, 'rb') as f:
                sha1 = hashlib.sha1(f.read()).hexdigest()
        except OSError:
            pass
    _write_json(detect_layout_file(), {'version': 1, 'items': layout['items'], 'known': [fid for fid, *_r in MANUAL_FIELDS],
                   'sha1': sha1, 'saved': datetime.datetime.now().isoformat(timespec='seconds')},
                ensure_ascii=False, indent=1)


def delete_detect_cache():
    try:
        os.remove(detect_layout_file())
    except FileNotFoundError:
        pass


# ---- the starting layout of the editor: whatever the selected way of filling the form would do
def blank_layout():
    return {'version': 1, 'items': {}, 'known': [fid for fid, *_r in MANUAL_FIELDS]}


def smart_shift_layout(template_pdf):
    """The fixed positions moved the way Smart Shift would move them on this template. -> (layout, note or None)"""
    base = manual_default_layout()
    layout, report = analyze_template(template_pdf)
    if layout is None:
        return base, T("Η Έξυπνη μετατόπιση δεν εφαρμόστηκε - χρησιμοποιήθηκαν οι σταθερές θέσεις.\n",
                       "Smart Shift was not applied - the fixed positions were used.\n") + describe_report(report)
    try:
        rows = detect_template_rows(template_pdf)
    except Exception:
        rows = [{}, {}]
    for fid, it in base['items'].items():
        page = it['page']
        if fid.startswith('rows'):
            ys = sorted(rows[page].values()) if page < len(rows) else []
            ry = ys[len(ys) // 2] if ys else 400.0
            it['x'] = round(it['x'] + layout.offset(page, it['x'], ry)[0], 2)
        else:
            x, y = layout.place(page, it['x'], it['y'])
            it['x'], it['y'] = round(x, 2), round(y, 2)
    return base, None


def build_mode_layout(mode, template_pdf):
    """The layout that a way of filling the form produces on this template, as the editor's starting point.
    -> (layout, note or None).  Smart Field Detection always scans again (it never reads the fine-tuned cache)."""
    if mode == 'smart':
        return smart_shift_layout(template_pdf)
    if mode == 'detect':
        try:
            layout, _rep = detect_template_fields(template_pdf)
            return layout, None
        except DetectError as e:
            return manual_default_layout(), str(e) + "\n" + T("Χρησιμοποιήθηκαν οι σταθερές θέσεις.",
                                                              "The fixed positions were used.")
    if mode == 'manual':
        return load_manual_layout(), None
    return manual_default_layout(), None


MANUAL_BASES = ('overlay', 'smart', 'detect', 'blank')
CURRENT_MANUAL_BASE = 'overlay'


def set_manual_base(base):
    global CURRENT_MANUAL_BASE
    CURRENT_MANUAL_BASE = base if base in MANUAL_BASES else 'overlay'


def manual_base_name(base):
    return {'overlay': T("Παλιά μέθοδος (σταθερές θέσεις)", "Legacy (fixed positions)"),
            'smart': T("Έξυπνη μετατόπιση", "Smart Shift"),
            'detect': T("Έξυπνη ανίχνευση πεδίων", "Smart Field Detection"),
            'blank': T("Κενός καμβάς", "Blank canvas")}.get(base, base)


def manual_base_layout(base, template_pdf):
    """The auto-fill the "Set manually" method starts from. -> (layout, note or None)"""
    if base == 'blank':
        return blank_layout(), None
    try:
        return build_mode_layout(base if base in MANUAL_BASES else 'overlay', template_pdf)
    except Exception as e:
        return manual_default_layout(), T(f"Η επιλεγμένη αυτόματη μέθοδος απέτυχε ({e}) - χρησιμοποιήθηκαν οι σταθερές θέσεις.",
                                          f"The chosen auto-fill failed ({e}) - the fixed positions were used.")


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
    for cf in load_custom_fields():
        cat.append((CF_PREFIX + cf['id'], (cf['name'], cf['name']), 'cfield'))
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
    if str(it.get('source') or '').startswith(CF_PREFIX):
        return 'vars'
    if it.get('source'):
        return 'extra'
    if it.get('parent'):
        fid = it['parent']
    return MANUAL_FIELD_INDEX[fid][1] if fid in MANUAL_FIELD_INDEX else 'custom'


def manual_label(fid, it, items=None):
    if it.get('custom'):
        return it.get('text') or T("(κενό κείμενο)", "(empty text)")
    if it.get('source'):
        return manual_source_label(it['source'])
    parent = it.get('parent')
    if parent and it.get('country'):
        col = T("Πελάτες", "Guests") if parent.endswith('_c') else T("Διανυκτ.", "Nights")
        return f"{_pretty_target(it['country'])} · {col} · " + T(f"Σελ. {it.get('page', 0) + 1}", f"Pg {it.get('page', 0) + 1}")
    if parent:
        return f"{manual_label(parent, (items or {}).get(parent, {}), items)} · #{it.get('index', 0) + 1}"
    entry = MANUAL_FIELD_INDEX.get(fid)
    return T(*entry[2]) if entry else fid


def manual_item_text(items, fid, values):
    """The text an item shows: its own text (custom), the value of its source (variables ...) or of its id."""
    it = items.get(fid, {})
    if it.get('custom'):
        return str(it.get('text', '') or '')
    if it.get('source'):
        return str(values.get(it['source'], '') or '')
    return str(values.get(fid, '') or '')


def manual_piece_text(items, it, values):
    """What one piece of a split text writes. The last piece of a free-length text also writes any surplus."""
    text = manual_item_text(items, it.get('parent'), values)
    idx = it.get('index', 0)
    if it.get('tail'):
        return text[idx:] if 0 <= idx < len(text) else ''
    return text[idx] if 0 <= idx < len(text) else ''


def manual_splittable(it):
    """Every text can be split into letters; only the numbers of a single country (a column of the form) cannot."""
    return not it.get('parent') and not str(it.get('source') or '').startswith('country:')


def manual_split_item(layout, fid, sample_text, row_keys):
    """Replaces one grouped item (country column, digit boxes) by individual items. Returns the new ids."""
    items = layout['items']
    it = items[fid]
    if it.get('split') or it.get('parent') or str(it.get('source') or '').startswith('country:'):
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
        free_length = fid not in MANUAL_SPLIT_LEN          # the length changes: the last piece takes the rest
        for i, x in enumerate(xs):
            cid = f"{fid}@{i}"
            child = manual_new_item(it['page'], round(x, 2), it['y'], it['size'])
            child.update(keep, parent=fid, index=i, step=0.0, align=align)
            if free_length and i == len(xs) - 1:
                child['tail'] = True
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
    for k in EXTRA_KEYS:
        v[k] = settings_data.get(k, '') or ''
    for k, val in settings_data.items():
        if str(k).startswith(CF_PREFIX):
            v[k] = val or ''
    return v


def manual_sample_inputs():
    """Made-up hotel + month, used for the editor preview and its test PDF."""
    sample = {'type': 'ΞΕΝΟΔΟΧΕΙΟ', 'name': 'TEST HOTEL', 'category': '5*', 'email': 'test@example.com',
              'phone': '2310000000', 'afm': '123456789', 'company': 'TEST COMPANY',
              'nomos': 'ΘΕΣΣΑΛΟΝΙΚΗΣ', 'dimotiki': 'ΘΕΣΣΑΛΟΝΙΚΗΣ', 'odos': 'ΤΣΙΜΙΣΚΗ', 'arithmos': '10', 'tk': '54623'}
    for cf in load_custom_fields():
        sample[CF_PREFIX + cf['id']] = cf['value'] or cf['name']
    today = datetime.date.today()
    last = calendar.monthrange(today.year, today.month)[1]
    mon = dict(year=str(today.year), month=str(today.month), date_to=f"{last:02d}/{today.month:02d}/{today.year}",
               days_of_month=str(last), employees='25', sign_date=today.strftime('%d/%m/%Y'))
    return sample, mon


def manual_sample_values():
    sample, mon = manual_sample_inputs()
    totals = {key: {'clients': 1234, 'nights': 5678} for key in MANUAL_TOTAL_KEYS}
    return manual_field_values(sample, mon, totals)


def real_report_inputs():
    """Saved hotel details, own fields and the latest saved month -> (settings, mon, rows, has_data).
    has_data is False when nothing has been saved yet (then there is nothing real to show)."""
    settings = load_hotel_settings()
    has_details = any(str(settings.get(k, '')).strip() for k in SETTING_KEYS)
    for cf in load_custom_fields():
        settings[CF_PREFIX + cf['id']] = cf['value']
        has_details = has_details or bool(str(cf['value']).strip())
    hist = load_history()
    latest = max(((y, m) for y, ms in hist.items() for m in ms), default=None)
    today = datetime.date.today()
    y, m = latest if latest else (today.year, today.month)
    rows = list(hist.get(y, {}).get(m, [])) if latest else []
    rec = load_month_records().get((y, m), {})
    last = calendar.monthrange(y, m)[1]
    mon = dict(year=str(y), month=str(m), date_to=f"{last:02d}/{m:02d}/{y}",
               days_of_month=str(rec.get('days') or last),
               employees=str(rec.get('employees') or settings.get('last_employees') or '0'),
               sign_date=today.strftime('%d/%m/%Y'))
    return settings, mon, rows, bool(has_details or latest)


# ---- drawing
def _hex_rgb(color):
    try:
        return tuple(int(color[i:i + 2], 16) / 255.0 for i in (1, 3, 5))
    except (ValueError, TypeError, IndexError):
        return (0.0, 0.0, 0.0)


def _manual_put(can, it, x, y, text):
    step = it.get('step') or 0
    maxw = it.get('maxw')
    if maxw and not step and text:          # a long value shrinks instead of running into the next label
        font, size = can._fontname, it['size']
        while size > 6.0 and pdfmetrics.stringWidth(text, font, size) > maxw:
            size -= 0.5
        if size != it['size']:
            can.setFont(font, size)
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
            piece = manual_piece_text(layout['items'], it, values)
            if piece:
                _manual_put(can, it, it['x'], it['y'], piece)
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


_PDFIUM_LOCK = threading.RLock()     # PDFium is not thread-safe: one page is drawn at a time (zoom draws in the background)
_PDF_BYTES = {}


def _pdf_file_bytes(path):
    """The PDF read once and kept in memory until it changes: zooming never touches the disk again and the
    file stays unlocked, so the program can still overwrite it."""
    try:
        st = os.stat(path)
        stamp = (st.st_mtime_ns, st.st_size)
        hit = _PDF_BYTES.get(path)
        if hit and hit[0] == stamp:
            return hit[1]
        with open(path, 'rb') as f:
            data = f.read()
        if len(_PDF_BYTES) > 6:
            _PDF_BYTES.clear()
        _PDF_BYTES[path] = (stamp, data)
        return data
    except OSError:
        return None


def _open_pdfium(pdfium, path):
    data = _pdf_file_bytes(path)
    if data is not None:
        try:
            return pdfium.PdfDocument(data)
        except Exception:
            pass
    return pdfium.PdfDocument(path)


def render_pdf_page_image(pdf_path, page_index, scale):
    """Pillow image (RGB) of one PDF page drawn at `scale` pixels per point, or None when it cannot be done."""
    if PILImage is None:
        return None
    with _PDFIUM_LOCK:
        try:
            import pypdfium2 as pdfium
            doc = _open_pdfium(pdfium, pdf_path)
            try:
                return doc[page_index].render(scale=scale).to_pil().convert('RGB')
            finally:
                try:
                    doc.close()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            import fitz
            data = _pdf_file_bytes(pdf_path)
            doc = fitz.open(stream=data, filetype='pdf') if data is not None else fitz.open(pdf_path)
            try:
                pix = doc[page_index].get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                return PILImage.frombytes('RGB', (pix.width, pix.height), pix.samples)
            finally:
                doc.close()
        except Exception:
            pass
    return None


def render_pdf_page_png(pdf_path, page_index, scale):
    """PNG bytes of one PDF page drawn at `scale` pixels per point, or None if nothing here can draw PDFs.
    Tried in turn: pypdfium2, PyMuPDF, Poppler's pdftoppm."""
    with _PDFIUM_LOCK:
        try:
            import pypdfium2 as pdfium
            doc = _open_pdfium(pdfium, pdf_path)
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


def commit_report_history(pdf_data, monthly_data, template_pdf, pdf_mode='overlay'):
    """Remembers a finished form: the first form made the default way becomes the layout reference, and the
    month's rows go into the history CSV (re-running a month replaces its old rows).
    Kept separate so a previewed form can be committed only after the user chose to save it."""
    if pdf_mode == 'overlay' and load_layout_reference() is None:
        try:        # the first real form made the default way shows the template the fixed positions fit
            save_layout_reference(template_pdf)
        except Exception:
            pass
    csv_filename = HISTORY_CSV
    csv_data = []
    for country, vals in pdf_data.items():
        csv_data.append({
            'Etos': str(int(monthly_data['year'])),
            'Minas': str(int(monthly_data['month'])),
            'Chwra': country,
            'Pelates': vals['clients'],
            'Dianykterefseis': vals['nights'],
            **({'Logis': round(vals['rev'][0], 2), 'FnB': round(vals['rev'][1], 2),
                'Extras': round(vals['rev'][2], 2), 'Esoda': round(vals['rev'][3], 2),
                'Domatia': round(vals['rev'][4], 2)} if 'rev' in vals else {})
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


def process_elstat_report(excel_path, template_pdf, output_pdf, settings_data, monthly_data,
                         rows_override=None, save_history=True, pdf_mode='overlay', manual_layout=None,
                         detect_layout=None):
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
    revenue_rows = read_revenue_rows(excel_path) if rows_override is None else {}
    for country_excel, pax, nights in rows:
        pdf_target = resolve_country(country_excel)
        if pdf_target is None:
            unmapped.append((country_excel, pax, nights))
            continue
        pdf_data.setdefault(pdf_target, {'clients': 0, 'nights': 0})
        pdf_data[pdf_target]['clients'] += pax
        pdf_data[pdf_target]['nights'] += nights
        if country_excel in revenue_rows:
            old = pdf_data[pdf_target].get('rev', (0, 0, 0, 0, 0))
            pdf_data[pdf_target]['rev'] = tuple(a + b for a, b in zip(old, revenue_rows[country_excel]))
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

    # "Smart Shift": find out where everything sits on THIS template; any failure -> default method
    layout, layout_warning, layout_info = None, None, None
    confidence = None               # per-section / per-field recognition confidence (Smart Shift and Smart Field Detection)
    if pdf_mode == 'smart':
        layout, layout_report = analyze_template(template_pdf)
        if layout is None:
            layout_warning = (T("Η έξυπνη μετατόπιση δεν εφαρμόστηκε - χρησιμοποιήθηκε η προεπιλεγμένη μέθοδος.\n",
                                "Smart Shift was not applied - the default method was used.\n")
                              + describe_report(layout_report))
        else:
            confidence = layout_report.get('confidence')
            shifts = layout_report['pages']
            biggest = max(p['max_shift'] for p in shifts)
            layout_info = T(f"Έξυπνη μετατόπιση: μέγιστη μετατόπιση {biggest:.1f} pt.",
                            f"Smart Shift: largest shift {biggest:.1f} pt.")
            if biggest > LARGE_SHIFT:
                layout_warning = T(f"Η έξυπνη μετατόπιση εφαρμόστηκε, αλλά το template έχει μετακινηθεί πολύ ({biggest:.0f} pt). "
                                   "Ελέγξτε προσεκτικά το δελτίο πριν το υποβάλετε.",
                                   f"Smart Shift was applied, but the template moved a lot ({biggest:.0f} pt). "
                                   "Check the form carefully before submitting.")

    # "Smart Field Detection": scan the blank template, then fill it with the detected (or fine-tuned) layout
    detect_rows = detect_report = detect_used = None
    if pdf_mode == 'detect':
        try:
            fresh, detect_report = detect_template_fields(template_pdf)
        except DetectError as e:
            pdf_mode = 'overlay'
            layout_warning = (T("Η Έξυπνη ανίχνευση πεδίων δεν εφαρμόστηκε - χρησιμοποιήθηκε η παλιά μέθοδος (σταθερές συντεταγμένες).\n",
                                "Smart Field Detection was not applied - the Legacy method (fixed coordinates) was used.\n") + str(e))
        else:
            detect_rows = detect_report['row_ys']
            confidence = detect_report.get('confidence')
            tuned = detect_layout if detect_layout is not None else load_detect_cache()
            detect_used = tuned if tuned is not None else fresh
            n_found = len(detect_report['found'])
            layout_info = T(f"Έξυπνη ανίχνευση: {n_found} από {n_found + len(detect_report['missing'])} πεδία",
                            f"Smart Field Detection: {n_found} of {n_found + len(detect_report['missing'])} fields")
            notes = []
            if tuned is None and detect_report['missing']:
                notes.append(T("Δεν βρέθηκαν και δεν γράφτηκαν: ", "Not found and not written: ")
                             + ", ".join(manual_label(f, {}, None) for f in detect_report['missing']))
            if tuned is not None and detect_layout is None and tuned.get('sha1') and tuned['sha1'] != detect_report['sha1']:
                notes.append(T("Το template άλλαξε μετά τις προσαρμογές σας στον επεξεργαστή. Ελέγξτε το δελτίο ή ξανατρέξτε "
                               "την ανίχνευση από τις Ρυθμίσεις.",
                               "The template changed after your fine-tuning in the editor. Check the form or rerun the "
                               "detection from Settings."))
            if detect_report['rows']['missing']:
                notes.append(T("Δεν βρέθηκε γραμμή στο template για: ", "No row found in the template for: ")
                             + ", ".join(_pretty_target(t) for t in detect_report['rows']['missing']
                                         if t in {norm(c) for c in pdf_data}))
                if notes[-1].endswith(": "):
                    notes.pop()
            if notes:
                layout_warning = T("Έξυπνη ανίχνευση πεδίων:\n", "Smart Field Detection:\n") + "\n".join(notes)

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

    if detect_rows:                         # rows read from the template by the detection win over the plain text scan
        for i in range(2):
            for target, y in detect_rows[i].items():
                if target in pdf_data:
                    coord_map[i][target] = y

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

            for key in EXTRA_KEYS:          # optional address details: only written when filled in
                text = str(settings_data.get(key, '') or '').strip()
                if text:
                    size = 10.0
                    while size > 6.0 and pdfmetrics.stringWidth(text, font_name, size) > EXTRA_WIDTH[key]:
                        size -= 0.5         # long values shrink instead of running into the next label
                    can.setFont(font_name, size)
                    put(EXTRA_POS[key][0], EXTRA_POS[key][1], text)
            can.setFont(font_name, 11)
            
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
        if manual_layout is not None:
            manual = manual_layout
        elif manual_layout_saved() or CURRENT_MANUAL_BASE in ('overlay', 'blank'):
            manual = load_manual_layout()
        else:        # nothing saved yet: fill the form with the auto-fill chosen for "Set manually"
            manual = manual_base_layout(CURRENT_MANUAL_BASE, template_pdf)[0]
    elif pdf_mode == 'detect':
        manual = copy.deepcopy(detect_used)
    if manual is not None:
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

    if save_history:
        commit_report_history(pdf_data, monthly_data, template_pdf, pdf_mode)

    if confidence:
        view = confidence_view(confidence, manual)
        flagged = view['flags'] if view else []
        if flagged:
            tail = T(f"{len(flagged)} προς έλεγχο", f"{len(flagged)} to check")
            layout_info = (layout_info + " · " if layout_info else "") + tail
        low = [f for f in flagged if f['level'] == 'low']       # 'missing' ones are already reported above
        if low:
            note = T(f"Χαμηλή βεβαιότητα αναγνώρισης ({len(low)}): " + ", ".join(f['label'] for f in low[:6])
                     + ". Δείτε τις επισημασμένες περιοχές στην προεπισκόπηση.",
                     f"Low recognition confidence ({len(low)}): " + ", ".join(f['label'] for f in low[:6])
                     + ". See the highlighted areas in the preview.")
            layout_warning = (layout_warning + "\n" if layout_warning else "") + note

    computed = (totals['9']['clients'], totals['9']['nights'])
    try:        # the real values of every field (used by the temporary layout editor of the preview)
        live_values = manual_field_values(settings_data, monthly_data, totals)
    except Exception:
        live_values = None
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
            'layout_warning': layout_warning, 'layout_info': layout_info, 'pdf_data': pdf_data,
            'live_values': live_values, 'layout_used': copy.deepcopy(manual) if manual is not None else None,
            'detect_report': detect_report, 'pdf_mode_used': pdf_mode, 'confidence': confidence}


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
    state = {'job': None, 'last': None}

    def _apply():
        state['job'] = None
        try:
            w = container.winfo_width()
            if w and w > 1:
                wrap = max(minimum, w - margin)
                if state['last'] is None or abs(wrap - state['last']) >= 2:      # skip needless re-layouts
                    state['last'] = wrap
                    label.configure(wraplength=wrap)
        except tk.TclError:
            pass

    def _update(_event=None):
        if state['job'] is None:                  # many Configure events in a row -> one update
            state['job'] = label.after_idle(_apply)
    container.bind('<Configure>', _update, add='+')


class HoverTip:
    """A small tip that appears when the pointer rests on a widget for a moment (like a file's tooltip in Explorer).
    `note` is a smaller second line, e.g. "Click for more in the guide"."""

    def __init__(self, widget, text, note=None, delay=700, wrap=340):
        self.widget, self.text, self.note, self.delay, self.wrap = widget, text, note, delay, wrap
        self._job = None
        self._tip = None
        widget.bind('<Enter>', self._schedule, add='+')
        widget.bind('<Leave>', self.hide, add='+')
        widget.bind('<ButtonPress>', self.hide, add='+')
        widget.bind('<Destroy>', self.hide, add='+')

    def _schedule(self, _event=None):
        self.hide()
        self._job = self.widget.after(self.delay, self._show)

    def _show(self):
        self._job = None
        try:
            tip = tk.Toplevel(self.widget)
            tip.wm_overrideredirect(True)
            try:
                tip.wm_attributes('-topmost', True)
            except tk.TclError:
                pass
            frame = tk.Frame(tip, bg=BORDER)
            frame.pack()
            inner = tk.Frame(frame, bg=CARD)
            inner.pack(padx=1, pady=1)
            tk.Label(inner, text=self.text, bg=CARD, fg=TEXT, font=(FONT, 9), justify="left",
                     wraplength=self.wrap).pack(anchor="w", padx=10, pady=(7, 7 if not self.note else 2))
            if self.note:
                tk.Label(inner, text=self.note, bg=CARD, fg=ACCENT, font=(FONT, 8, 'bold'), justify="left",
                         wraplength=self.wrap).pack(anchor="w", padx=10, pady=(0, 7))
            tip.update_idletasks()
            w, h = tip.winfo_reqwidth(), tip.winfo_reqheight()
            x = min(self.widget.winfo_rootx() + 14, max(self.widget.winfo_screenwidth() - w - 8, 0))
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
            if y + h > self.widget.winfo_screenheight() - 8:          # no room below: show it above
                y = max(self.widget.winfo_rooty() - h - 8, 0)
            tip.wm_geometry(f"+{x}+{y}")
            self._tip = tip
        except tk.TclError:
            self._tip = None

    def hide(self, _event=None):
        if self._job is not None:
            try:
                self.widget.after_cancel(self._job)
            except tk.TclError:
                pass
            self._job = None
        if self._tip is not None:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None


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
        self._head = head
        self._title = tk.Label(head, text=title, bg=CARD, fg=NAVY, font=(FONT, 13, 'bold'))
        self._title.pack(side="left", padx=(10 if number is not None else 0, 0))
        self.body = ttk.Frame(inner, style='Card.TFrame', padding=(20, 10, 20, 20))
        self.body.pack(fill="x")

    def make_collapsible(self, start_open=False):
        """Click the title to fold the card up to just its heading (less clutter)."""
        state = {'open': start_open}
        arrow = tk.Label(self._head, bg=CARD, fg=ACCENT, font=(FONT, 12, 'bold'), cursor='hand2')
        arrow.pack(side="right")

        def show():
            arrow.configure(text="\u25be" if state['open'] else "\u25b8")
            if state['open']:
                self.body.pack(fill="x")
            else:
                self.body.pack_forget()

        def toggle(_e=None):
            state['open'] = not state['open']
            show()

        for w in (self._head, self._title, arrow):
            w.configure(cursor='hand2')
            w.bind('<Button-1>', toggle)
        show()

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
        self._pos = {}
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

    @staticmethod
    def _axis(v):
        if v >= 1_000_000:
            return f"{v / 1_000_000:g}M"
        if v >= 10_000:
            return f"{v / 1000:g}k"
        return fmt_int(round(v)) if v >= 100 else f"{v:g}"

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
        for frac in ((0, .25, .5, .75, 1) if plot_h >= 150 else (0, .5, 1)):
            y = top + plot_h * (1 - frac)
            self.create_line(left, y, w - right, y, fill=GRID)
            self.create_text(left - 8, y, text=self._axis(top_val * frac), anchor="e", fill=MUTED, font=(FONT, 8))
        shown = [i for i in range(12) if self.values[i] > 0 or (self.prev and self.prev[i] > 0)]
        if not shown:
            self.create_text((left + w - right) / 2, top + plot_h / 2, text=T("Δεν υπάρχουν δεδομένα", "No data"),
                             fill=MUTED, font=(FONT, 10))
            return
        slot = (w - left - right) / len(shown)
        bw = min(slot * 0.6, 56)
        self._pos = {}
        small = slot < 44                                    # too narrow for a label on every bar
        for k, i in enumerate(shown):
            cx = left + slot * (k + .5)
            self._pos[i + 1] = (cx, slot)
            v = self.values[i]
            if self.hover == i + 1:
                hw = bw / 2 + 4
                self.create_rectangle(cx - hw, top - 6, cx + hw, top + plot_h + 2, fill=TRACK, outline='')
            if v > 0:
                y0 = top + plot_h * (1 - v / top_val)
                fill = ACCENT_DARK if self.selected == i + 1 else self.color
                self.create_rectangle(cx - bw / 2, y0, cx + bw / 2, top + plot_h, fill=fill, outline='')
                if not small or self.hover == i + 1:
                    self.create_text(cx, y0 - 8, text=self._fmt(v), fill=TEXT, font=(FONT, 8))
            pv = self.prev[i] if self.prev else 0
            if pv > 0:
                yp = top + plot_h * (1 - pv / top_val)
                self.create_line(cx - bw / 2 - 3, yp, cx + bw / 2 + 3, yp, fill='#f59e0b', width=3)
            hw = min(slot / 2, bw / 2 + 6)
            self._slots.append((cx - hw, cx + hw, i + 1))
            self.create_text(cx, h - 12, text=month_short(i + 1), fill=MUTED, font=(FONT, 8))
        if self.hover and self.hover in self._pos and self.prev:   # only worth a tooltip when it adds the comparison
            self._tooltip(self.hover, left, slot, top)

    def _tooltip(self, month, left, slot, top):
        v = self.values[month - 1]
        lines = [f"{month_name(month)}: {self._fmt(v)}"]
        if self.prev:
            pv = self.prev[month - 1]
            lines.append(f"{self.prev_label}: {self._fmt(pv)}")
            if pv > 0:
                lines.append(f"{fmt_dec((v - pv) / pv * 100, signed=True)}%")
        cx, slot = self._pos[month]
        w = self.winfo_width()
        tw = max(110, 8 * max(len(l) for l in lines) + 20)
        half = min(slot / 2, 40)
        x = cx + half + 4
        if x + tw > w - 4:
            x = cx - half - 4 - tw
        x = max(4, min(x, w - tw - 4))
        y = top - 4
        self.create_rectangle(x, y, x + tw, y + 14 * len(lines) + 10, fill=HEAD_BG, outline=ACCENT)
        self.create_text(x + 8, y + 6, text="\n".join(lines), anchor="nw", fill=TEXT,
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
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, bg=BG, yscrollincrement=1)   # 1 unit = 1 pixel
        self._pending = 0.0                         # pixels still to scroll (the wheel only sets a target)
        self._anim = None
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self._vbar_command)
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

    def _vbar_command(self, *args):
        if len(args) == 3 and args[0] == 'scroll' and args[2] == 'units':      # arrow clicks: 1 unit is only 1 px now
            args = ('scroll', int(args[1]) * 24, 'units')
        self.canvas.yview(*args)

    def _on_body(self, _event=None):
        box = self.canvas.bbox("all")
        if box != getattr(self, '_last_box', None):
            self._last_box = box
            self.canvas.configure(scrollregion=box)
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
        """Mouse wheel: glides to the target in small steps instead of jumping a third of the window at once."""
        if not self._needs_scroll():
            return
        self._pending += units * 14                 # one wheel notch (3 units) = about 42 px
        if self._anim is None:
            self._anim = self.after(10, self._glide)

    def _glide(self):
        self._anim = None
        if abs(self._pending) < 1:
            self._pending = 0.0
            return
        step = self._pending * 0.35
        step = max(1, int(abs(step))) * (1 if step > 0 else -1)
        before = self.canvas.yview()[0]
        self.canvas.yview_scroll(step, "units")
        self._pending -= step
        if self.canvas.yview()[0] == before:        # reached the top or bottom
            self._pending = 0.0
            return
        self._anim = self.after(10, self._glide)


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
        ('p', ("Το πρόγραμμα παίρνει το αρχείο Excel (ή CSV) που βγάζει το Protel και συμπληρώνει αυτόματα το μηνιαίο "
               "δελτίο της ΕΛΣΤΑΤ σε PDF. Ο μήνας ολοκληρώνεται σε πέντε βήματα:",
               "The program takes the Excel (or CSV) file that Protel produces and fills in the monthly ELSTAT form as a PDF "
               "automatically. A month takes five steps:")),
        ('live', 'next'),
        ('step', ("Την πρώτη φορά, ανοίξτε την καρτέλα **Στοιχεία ξενοδοχείου** και συμπληρώστε τα στοιχεία σας. "
                  "Γίνεται μία φορά και το πρόγραμμα τα θυμάται.",
                  "The first time, open the **Hotel details** tab and fill in your details. You do this once; "
                  "the program remembers them.")),
        ('step', ("Στην καρτέλα **Δελτίο μήνα**, σύρτε το Excel ή CSV του μήνα μέσα στο πλαίσιο του Βήματος 1 ή πατήστε "
                  "πάνω του για να το διαλέξετε.",
                  "In the **Monthly form** tab, drop the month's Excel or CSV into the box in Step 1, or click it to choose the file.")),
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
        ('act', ("Επιλογή Excel / CSV του μήνα", "Choose the month's Excel / CSV"), 'choose_excel'),
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
        ('p', ("Το αρχείο που χρειάζεται λέγεται «Nationalities Statistics» και βγαίνει από το Protel ως .xls, .xlsx ή .csv. "
               "Μπορείτε να το διαλέξετε με όποιον τρόπο σας βολεύει:",
               "The file you need is called “Nationalities Statistics” and comes from Protel as .xls, .xlsx or .csv. "
               "Choose it whichever way suits you:")),
        ('li', ("Σύρτε το πάνω στο διακεκομμένο πλαίσιο (ή πάνω σε ολόκληρη την κάρτα).",
                "Drop it onto the dashed box (or anywhere on the card).")),
        ('li', ("Πατήστε το πλαίσιο ή το κουμπί **Επιλογή Excel / CSV…** (`Ctrl+O`).",
                "Click the box or the **Choose Excel / CSV…** button (`Ctrl+O`).")),
        ('li', ("Πατήστε **Πιο πρόσφατο** για να διαλέξει το πρόγραμμα το νεότερο Excel ή CSV από τον φάκελο που χρησιμοποιήσατε τελευταία.",
                "Click **Latest** and the program picks the newest Excel or CSV in the folder you used last.")),
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
        ('h', ("Αρχεία Excel / CSV από το Protel", "Excel / CSV files from Protel")),
        ('p', ("Από ποιον φάκελο ξεκινούν τα **Επιλογή Excel / CSV…** και **Πιο πρόσφατο**. Κενό σημαίνει ο φάκελος του τελευταίου αρχείου που χρησιμοποιήσατε.",
               "The folder where **Choose Excel / CSV…** and **Latest** start. Empty means the folder of the last file you used.")),
        ('h', ("Τρόπος συμπλήρωσης του PDF", "How the PDF is filled in")),
        ('p', ("Υπάρχουν τέσσερις τρόποι. Η **Παλιά μέθοδος** είναι η δοκιμασμένη· η **Έξυπνη μετατόπιση** και η **Έξυπνη "
               "ανίχνευση πεδίων** μπορεί περιστασιακά να αποτύχουν, οπότε το δελτίο φτιάχνεται με την παλιά μέθοδο και "
               "εμφανίζεται προειδοποίηση.",
               "There are four ways. The **Legacy** method is the proven one; **Smart Shift** and **Smart Field Detection** are "
               "can occasionally fail, in which case the form is made with the Legacy method and a warning is shown.")),
        ('li', ("**Παλιά μέθοδος (σταθερές συντεταγμένες) · LEGACY**: οι τιμές γράφονται σε σταθερές θέσεις πάνω στο άδειο PDF. Έτσι δούλευε πάντα το πρόγραμμα.",
                "**Legacy (Fixed Coordinates)**: the values are written at fixed positions on the blank PDF. This is how the program has always worked.")),
        ('li', ("**Έξυπνη μετατόπιση (Smart Shift)**: διαβάζει σημεία-άγκυρες του template, μετρά πόσο μετακινήθηκαν σε "
                "σχέση με τη διάταξη αναφοράς (διανύσματα μετατόπισης και χωρική παρεμβολή IDW) και μετατοπίζει ανάλογα τις σταθερές "
                "συντεταγμένες. Χρειάζεται διάταξη αναφοράς.",
                "**Smart Shift**: reads anchor texts of the template, measures how far they moved against the reference "
                "layout (delta vectors and IDW spatial interpolation) and shifts the fixed coordinates accordingly. It needs a reference layout.")),
        ('li', ("**Έξυπνη ανίχνευση πεδίων (AI / αυτόματη ανίχνευση)**: πλήρως αυτόνομη. Σαρώνει το άδειο PDF εκείνη τη "
                "στιγμή, διαβάζει τις ετικέτες, ευθυγραμμίζει τον πίνακα και υπολογίζει τις συντεταγμένες χωρίς καμία σταθερή τιμή για "
                "τα πεδία του προγράμματος. Δεν χρειάζεται αναφορά. Χρειάζεται το pypdfium2. Δικά σας πεδία και κείμενα μένουν χειροκίνητα (επεξεργαστής διάταξης).",
                "**Smart Field Detection (AI / Auto-Detect)**: fully autonomous. It scans the blank PDF on the fly, reads "
                "the labels, aligns the table grid and computes the coordinates with nothing hardcoded for the program's own fields. "
                "No reference is needed. It requires pypdfium2. Your own fields and texts stay manual (layout editor).")),
        ('li', ("**Χειροκίνητη τοποθέτηση**: τα βάζετε όλα εσείς στον επεξεργαστή διάταξης.",
                "**Set manually**: you place everything yourself in the layout editor.")),
        ('warn', ("Οι έξυπνες μέθοδοι δεν αντικαθιστούν τον έλεγχο: κοιτάξτε πάντα το δελτίο πριν το υποβάλετε.",
                  "The smart methods do not replace checking: always look at the form before submitting it.")),
        ('h', ("Έξυπνη ανίχνευση πεδίων: ρύθμιση και έλεγχος", "Smart Field Detection: configure and inspect")),
        ('step', ("Επιλέξτε «Έξυπνη ανίχνευση πεδίων». Από κάτω εμφανίζεται το κουμπί **👉 Ρύθμιση & έλεγχος ανίχνευσης…**.",
                  "Select “Smart Field Detection”. The button **👉 Configure & Inspect Auto-Detection…** appears right beneath it.")),
        ('step', ("Στο παράθυρο ελέγχου βλέπετε τη λίστα των πεδίων που βρέθηκαν (✓) ή όχι (○, κρυμμένα) και ένα δείγμα συμπληρωμένου "
                  "δελτίου με πλασματικά δεδομένα. Πατώντας ένα πεδίο, σημειώνεται η θέση του στη σελίδα.",
                  "The inspection window lists the fields that were found (✓) or not (○, hidden) and shows a sample form filled with "
                  "made-up data. Click a field and its position is marked on the page.")),
        ('step', ("**🔄 Επανάληψη ανίχνευσης** σαρώνει ξανά το template (και διαγράφει τις προσαρμογές σας, με ερώτηση). "
                  "**Άνοιγμα στον επεξεργαστή διάταξης…** ανοίγει τον επεξεργαστή για λεπτομερή ρύθμιση.",
                  "**🔄 Rerun Auto-Detection** scans the template again (and discards your fine-tuning, after asking). "
                  "**Open in Layout Editor…** opens the editor for fine-tuning.")),
        ('p', ("Οι προσαρμογές που αποθηκεύετε στον επεξεργαστή θυμούνται στα επόμενα δελτία όσο μένετε στην Έξυπνη ανίχνευση πεδίων. "
               "Αν επιλέξετε άλλον τρόπο και ξαναγυρίσετε, οι προσαρμογές διαγράφονται και η ανίχνευση ξεκινά από την αρχή.",
               "The fine-tuning you save in the editor is remembered for the following forms as long as you stay on Smart Field Detection. "
               "If you pick another method and come back, the fine-tuning is discarded and the detection starts from scratch.")),
        ('h', ("Δύο διαφορετικές προεπισκοπήσεις", "Two different previews")),
        ('li', ("**Παράθυρο ελέγχου ανίχνευσης** (από τις Ρυθμίσεις): για να ελέγξετε τα πεδία και να δοκιμάσετε δείγμα δεδομένων πριν "
                "φτιάξετε πραγματικό δελτίο. Έχει **Επανάληψη ανίχνευσης** και **Άνοιγμα στον επεξεργαστή διάταξης…**.",
                "**Detection inspection window** (from Settings): to check the fields and try sample data before any real report. "
                "It has **Rerun Auto-Detection** and **Open in Layout Editor…**.")),
        ('li', ("**Προεπισκόπηση τελικού ελέγχου** (όταν πατάτε «Δημιουργία PDF» και είναι ενεργή η προεπισκόπηση): τελευταία ματιά στο "
                "συμπληρωμένο δελτίο. Έχει **Επεξεργασία διάταξης…** (προσωρινά, μόνο για αυτό το δελτίο) αλλά ΟΧΙ επανάληψη ανίχνευσης, "
                "γιατί η ανίχνευση αυτού του δελτίου είναι ήδη οριστική.",
                "**Final verification preview** (when you press “Create PDF” and the preview is on): a last look at the filled form. "
                "It has **Edit layout…** (temporary, for this form only) but NOT rerun detection, because the detection of this form is already final.")),
        ('h', ("Βεβαιότητα αναγνώρισης", "Recognition confidence")),
        ('p', ("Η Έξυπνη μετατόπιση και η Έξυπνη ανίχνευση πεδίων δείχνουν πόσο σίγουρο είναι το πρόγραμμα για κάθε ενότητα του δελτίου "
               "(Στοιχεία καταλύματος, Πεδία ημερομηνιών, Πίνακας χωρών, Σύνολα). Οι θέσεις που αξίζει να ελέγξετε εμφανίζονται σε λίστα και "
               "σημειώνονται στη σελίδα με πορτοκαλί (έλεγχος) ή κόκκινο (χαμηλή βεβαιότητα) περίγραμμα και το ποσοστό τους. "
               "Αν μετακινήσετε μια θέση στον επεξεργαστή διάταξης, θεωρείται ελεγμένη και η βεβαιότητα ανεβαίνει. Το 100% εμφανίζεται μόνο "
               "όταν τίποτα δεν αμφισβητείται.",
               "Smart Shift and Smart Field Detection show how sure the program is about each section of the form (Hotel information, "
               "Date fields, Country table, Totals). The positions worth checking are listed and outlined on the page - amber (check) or "
               "red (low confidence) - together with their percentage. If you move a position in the layout editor it counts as checked "
               "and the confidence goes up. 100% is only shown when nothing at all is in doubt.")),
        ('h', ("Αυτόματη προσυμπλήρωση του προτύπου χειροκίνητης διάταξης", "Auto-populate manual layout template")),
        ('p', ("Ενεργό (προεπιλογή): ο επεξεργαστής διάταξης ανοίγει με τα στοιχεία τοποθετημένα όπως θα τα έβαζε ο επιλεγμένος τρόπος "
               "(παλιά μέθοδος, Έξυπνη μετατόπιση ή Έξυπνη ανίχνευση). Ανενεργό: ανοίγει εντελώς κενός καμβάς. Μια χειροκίνητη διάταξη "
               "που έχετε ήδη αποθηκεύσει ανοίγει πάντα όπως την αφήσατε.",
               "On (default): the layout editor opens with the items placed the way the selected method (Legacy, Smart Shift or Smart Field "
               "Detection) would place them. Off: it opens a completely blank canvas. A manual layout you already saved always opens as you left it.")),
        ('p', ("Μέσα στον επεξεργαστή: **Εκκαθάριση όλων** αδειάζει τον καμβά και **Αυτόματη τοποθέτηση** τοποθετεί ξανά όλα τα στοιχεία "
               "με τον τρόπο που έχετε επιλέξει.",
               "Inside the editor: **Clear all** empties the canvas and **Auto-place items** places every item again with the method you selected.")),
        ('h', ("Έξυπνη μετατόπιση: βήματα", "Smart Shift: steps")),
        ('step', ("Επιλέξτε «Έξυπνη μετατόπιση». Τα πρόσθετα κουμπιά εμφανίζονται μόνο όταν είναι επιλεγμένη.",
                  "Select “Smart Shift”. The extra buttons appear only while it is selected.")),
        ('step', ("Αποθηκεύστε το σημείο σύγκρισης: το πρόγραμμα το κάνει μόνο του με το πρώτο δελτίο που φτιάχνετε με την "
                  "παλιά μέθοδο, ή πατήστε **Αποθήκευση ως αναφορά…** όταν το δελτίο συμπληρώνεται σωστά.",
                  "Save the point of comparison: the program does it by itself with the first form you make with the Legacy "
                  "method, or press **Save as reference…** when the form fills in correctly.")),
        ('step', ("Όταν αλλάξει το template, πατήστε **Έλεγχος template** και **Δοκιμαστικό δελτίο** (PDF με πλασματικά στοιχεία - μην το υποβάλετε).",
                  "When the template changes, press **Check template** and **Test form** (a PDF with made-up data - do not submit it).")),
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
        ('step', ("Από το **Δεδομένα** πάνω επιλέγετε τι φαίνεται στη σελίδα: **Δείγμα** (πλασματικές τιμές), **Πραγματικά** (τα αποθηκευμένα δικά σας στοιχεία και ο τελευταίος μήνας) ή **Ονόματα** (μόνο τα ονόματα των μεταβλητών). Πατήστε **Αποθήκευση** (Ctrl+S) όταν είστε έτοιμοι.",
                  "Use **Data** at the top to choose what the page shows: **Sample** (made-up values), **Real** (your saved details and the latest month) or **Names** (only the variable names). Press **Save** (Ctrl+S) when you are happy.")),
        ('go', ("Δείτε τον πλήρη οδηγό του επεξεργαστή", "See the full layout editor guide"), 'sec:editor'),
        ('h', ("Ενημέρωση προγράμματος και πρώτη εκκίνηση", "Program update and first-run setup")),
        ('p', ("Το πρόγραμμα ενημερώνεται μόνο του από το GitHub. Διαβάζει τον αριθμό έκδοσης από το αρχείο `main.py` του αποθετηρίου και, αν είναι μεγαλύτερος από τον δικό σας, μπορείτε να το εγκαταστήσετε με ένα κλικ. Το προηγούμενο αρχείο μένει ως `main.py.bak`.",
               "The program updates itself from GitHub. It reads the version number from the repository's `main.py` and, if it is higher than yours, you can install it with one click. The previous file is kept as `main.py.bak`.")),
        ('li', ("**Αποθετήριο GitHub** και **Κλάδος**: από πού γίνεται ο έλεγχος (προεπιλογή `webzerot/eltstat`, κλάδος `main`).",
                "**GitHub repository** and **Branch**: where the check looks (default `webzerot/eltstat`, branch `main`).")),
        ('li', ("**Έλεγχος για ενημερώσεις**: δείχνει αν υπάρχει νέα έκδοση και, αν υπάρχει, τι άλλαξε (από το `CHANGELOG.md`).",
                "**Check for updates**: shows whether a new version exists and, if so, what changed (from `CHANGELOG.md`).")),
        ('li', ("**Εγκατάσταση ενημέρωσης**: κατεβάζει τη νέα έκδοση και αντικαθιστά το `main.py`. Μετά, ανοίξτε ξανά το πρόγραμμα.",
                "**Install update**: downloads the new version and replaces `main.py`. Afterwards, open the program again.")),
        ('li', ("**Έλεγχος για ενημέρωση όταν ανοίγει το πρόγραμμα**: ο έλεγχος γίνεται αυτόματα σε κάθε εκκίνηση.",
                "**Check for updates when the program starts**: the check runs automatically at every start.")),
        ('li', ("**Επαναφορά προηγούμενης έκδοσης**: γυρίζει στο `main.py.bak` αν κάτι δεν δουλεύει μετά την ενημέρωση.",
                "**Restore previous version**: goes back to `main.py.bak` if something stops working after an update.")),
        ('tip', ("Μετά από κάθε νέο ανέβασμα στο GitHub περιμένετε περίπου 5 λεπτά πριν πατήσετε τον έλεγχο, γιατί το GitHub κρατά προσωρινά την παλιά έκδοση.",
                 "After every push to GitHub wait about 5 minutes before checking, because GitHub keeps serving the old version for a short time.")),
        ('h', ("Πρώτη εκκίνηση", "First run")),
        ('p', ("Την πρώτη φορά που ανοίγει το πρόγραμμα κατεβάζει το `requirements.txt` από το GitHub στον φάκελο του προγράμματος και εγκαθιστά τα πακέτα που χρειάζεται. Εμφανίζεται ένα μικρό παράθυρο αναμονής με χρονόμετρο· μπορεί να κρατήσει μερικά λεπτά και χρειάζεται internet.",
               "The first time the program starts it downloads `requirements.txt` from GitHub into the program's folder and installs the packages it needs. A small waiting window with a timer is shown; it can take a few minutes and needs internet.")),
        ('li', ("Όταν τελειώσει, δημιουργείται το κρυφό αρχείο `.elstat_setup_done` και οι επόμενες εκκινήσεις είναι άμεσες.",
                "When it finishes, the hidden file `.elstat_setup_done` is created and later starts are instant.")),
        ('li', ("Αν λείπει κάποιο πακέτο αργότερα, η εγκατάσταση ξαναγίνεται αυτόματα. Δεν χρειάζεται να αλλάξετε κανέναν φάκελο.",
                "If a package goes missing later, the installation runs again by itself. You do not need to change any folder.")),
        ('h', ("Συμπεριφορά του προγράμματος", "Program behaviour")),
        ('p', ("Κάθε επιλογή έχει ένα ⓘ στις Ρυθμίσεις· πατώντας το έρχεστε ακριβώς εδώ.",
               "Every option has an ⓘ in Settings; clicking it brings you straight here.")),
        ('anchor', 'opt_center_window'),
        ('li', ("**Κεντράρισμα του παραθύρου κατά την εκκίνηση**: το κύριο παράθυρο ανοίγει στο κέντρο της οθόνης και εμφανίζεται μία φορά, "
                "χωρίς να «πηδάει». Αν το κλείσετε, τη θέση την επιλέγει το σύστημα. Δεν επηρεάζει τα υπόλοιπα παράθυρα.",
                "**Centre the window when the program starts**: the main window opens in the middle of the screen and appears once, "
                "without jumping. If you turn it off, the system chooses the position. It does not affect the other windows.")),
        ('anchor', 'opt_open_after'),
        ('li', ("**Άνοιγμα του PDF αμέσως μετά τη δημιουργία**: όταν αποθηκευτεί το δελτίο, ανοίγει στο πρόγραμμα PDF του υπολογιστή σας. "
                "Με την προεπισκόπηση ανοίγει αφού πατήσετε «Αποθήκευση & συνέχεια». Δεν ανοίγει όταν στέλνετε το δελτίο κατευθείαν στον εκτυπωτή.",
                "**Open the PDF right after it is created**: once the form is saved it opens in your computer's PDF viewer. "
                "With the preview it opens after you press “Save & continue”. It does not open when you send the form straight to the printer.")),
        ('anchor', 'opt_confirm_overwrite'),
        ('li', ("**Ερώτηση πριν αντικατασταθεί PDF που υπάρχει ήδη**: αν υπάρχει αρχείο με το ίδιο όνομα, ρωτάει «Να αντικατασταθεί;» (η προεπιλογή είναι «Όχι»). "
                "Αν το κλείσετε, το παλιό αρχείο αντικαθίσταται χωρίς ερώτηση. Αν το PDF είναι ανοιχτό αλλού, το πρόγραμμα δεν μπορεί να το γράψει και σας το λέει.",
                "**Ask before replacing a PDF that already exists**: if a file with the same name exists you are asked “Replace it?” (the default answer is No). "
                "If you turn it off, the old file is replaced without asking. If the PDF is open in another program it cannot be written and you are told.")),
        ('anchor', 'opt_keep_archive'),
        ('li', ("**Αρχειοθέτηση αντιγράφου κάθε δελτίου**: κάθε δελτίο αντιγράφεται στον φάκελο δεδομένων. Έτσι μπορείτε να ξανανοίξετε ή να "
                "εκτυπώσετε παλιούς μήνες από το Ιστορικό. Αν το κλείσετε, το δελτίο ξαναφτιάχνεται από τα δεδομένα του ιστορικού όταν το ζητήσετε.",
                "**Archive a copy of every form**: each form is copied into the data folder, so you can reopen or print old months from History. "
                "If you turn it off, the form is rebuilt from the history data when you ask for it.")),
        ('anchor', 'opt_show_reminder'),
        ('li', ("**Υπενθύμιση όταν λείπει το δελτίο του προηγούμενου μήνα**: στην καρτέλα του δελτίου εμφανίζεται μια γραμμή υπενθύμισης με τον μήνα που λείπει. "
                "Μπορείτε να την κλείσετε από εκεί για τη συγκεκριμένη χρήση.",
                "**Remind me when last month's form is missing**: a reminder line with the missing month appears on the form tab. "
                "You can dismiss it there for the current session.")),
        ('anchor', 'opt_prefill_employees'),
        ('li', ("**Προσυμπλήρωση του «Σύνολο εργαζομένων»**: το πεδίο παίρνει την τελευταία τιμή που χρησιμοποιήσατε. "
                "Ελέγχετε πάντα τον αριθμό πριν δημιουργήσετε το δελτίο.",
                "**Pre-fill “Total employees”**: the field gets the last value you used. "
                "Always check the number before creating the form.")),
        ('anchor', 'opt_show_extra_fields'),
        ('li', ("**Πρόσθετα προαιρετικά στοιχεία στα «Στοιχεία ξενοδοχείου»**: εμφανίζει ένα αναδιπλούμενο μενού με Νομό, Δημοτική ενότητα, Οδό, Αριθμό, Τ.Κ. και δικά σας πεδία. "
                "Αν το κλείσετε, τίποτα από αυτά δεν γράφεται στο δελτίο - οι τιμές όμως διατηρούνται.",
                "**Optional extra details in “Hotel details”**: shows a drop-down with Prefecture, Municipal unit, Street, Number, Postal code and your own fields. "
                "If you turn it off, none of them is written on the form - the values are kept.")),
        ('anchor', 'opt_show_loading'),
        ('li', ("**Ενδείξεις φόρτωσης**: μια μικρή κινούμενη ένδειξη πάνω δεξιά όσο διαβάζεται το Excel ή δημιουργείται το PDF. "
                "Αν την κλείσετε, το πρόγραμμα εργάζεται ακριβώς το ίδιο, απλώς χωρίς ένδειξη.",
                "**Loading indicators**: a small animated badge at the top right while the Excel is read or the PDF is created. "
                "If you turn it off, the program works exactly the same, just without the badge.")),
        ('anchor', 'opt_preview_before_result'),
        ('li', ("**Προεπισκόπηση PDF πριν την αποθήκευση**: το δελτίο εμφανίζεται μέσα στο πρόγραμμα με ζουμ και προσωρινό επεξεργαστή διάταξης. "
                "Το αρχείο αποθηκεύεται (και γίνεται η ερώτηση «υπάρχει ήδη») μόνο όταν πατήσετε «Αποθήκευση & συνέχεια». Αν κλείσετε την προεπισκόπηση, δεν αποθηκεύεται τίποτα.",
                "**Preview the PDF before it is saved**: the form is shown inside the program with zoom and a temporary layout editor. "
                "The file is saved (and the “already exists” question asked) only when you press “Save & continue”. If you close the preview, nothing is saved.")),
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
        ('li', ("**Αριστερά**: τα στοιχεία που τυπώνονται στο PDF, ανά ενότητα και με τον αριθμό τους, με πεδίο αναζήτησης από πάνω. Από κάτω τα κουμπιά προσθήκης και επεξεργασίας (δείτε παρακάτω).",
                "**Left**: the items that are printed on the PDF, by section and with their count, with a search box on top. Below it, the add and edit buttons (see below).")),
        ('li', ("**Στο κέντρο**: η σελίδα του εντύπου. Κάθε τιμή φαίνεται εκεί που θα τυπωθεί. Το κουμπί **Σελίδα 1 / Σελίδα 2** αλλάζει σελίδα, τα **− / + / Προσαρμογή** το μέγεθος (Ctrl + τροχός = ζουμ).",
                "**Centre**: the page of the form. Every value appears where it will be printed. **Page 1 / Page 2** switches the page, **− / + / Fit** the size (Ctrl + wheel = zoom).")),
        ('li', ("**Δεξιά, πάνω**: οι ιδιότητες του επιλεγμένου στοιχείου (κυλίστε αν το παράθυρο είναι μικρό). **Δεξιά, κάτω**: το παράθυρο **Κρυμμένα και αχρησιμοποίητα**. Η διαχωριστική γραμμή ανάμεσά τους σέρνεται.",
                "**Right, top**: the properties of the selected item (scroll if the window is small). **Right, bottom**: the **Hidden and not used** window. The divider between them can be dragged.")),
        ('li', ("Στο κάτω μέρος του παραθύρου μια μπάρα συμβουλών αλλάζει ανάλογα με το τι έχετε επιλέξει.",
                "At the bottom of the window a tip bar changes with what you have selected.")),
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
        ('li', ("**Εμφάνιση στο PDF**: αν το ξετσεκάρετε, το στοιχείο κρύβεται (δείτε Κρυμμένα και αχρησιμοποίητα παρακάτω).",
                "**Show in the PDF**: untick it to hide the item (see Hidden and not used below).")),
        ('h', ("Κρυμμένα και αχρησιμοποίητα", "Hidden and not used")),
        ('p', ("Το παράθυρο κάτω δεξιά έχει δύο λίστες, ώστε να ξέρετε πάντα τι **δεν** τυπώνεται:",
               "The window at the bottom right has two lists, so you always know what is **not** printed:")),
        ('li', ("**Κρυμμένα στο PDF**: στοιχεία που έχετε τοποθετήσει αλλά δεν τυπώνονται (τα κρύψατε εσείς ή η αυτόματη ανίχνευση δεν βρήκε θέση). Με ένα κλικ επιλέγονται και στη σελίδα.",
                "**Hidden on the PDF**: items you placed but that are not printed (you hid them, or auto-detection found no place for them). One click also selects them on the page.")),
        ('li', ("**Δεν υπάρχουν στο έντυπο**: ό,τι μπορεί να μπει στο δελτίο αλλά δεν έχει τοποθετηθεί πουθενά - πεδία, στήλες χωρών, πρόσθετα στοιχεία και οι μεταβλητές σας, ανά ενότητα.",
                "**Not on the form**: everything that could go on the form but is not placed anywhere - fields, country columns, extra items and your variables, by section.")),
        ('step', ("Βρείτε το στοιχείο (έχει δικό του πεδίο αναζήτησης).", "Find the item (the window has its own search box).")),
        ('step', ("Διπλό κλικ, ή επιλέξτε και πατήστε **Βάλε στο PDF** (ή `Enter`). Επιλέγετε πολλά μαζί με `Ctrl`/`Shift`.",
                  "Double-click it, or select it and press **Put on the PDF** (or `Enter`). Pick several with `Ctrl`/`Shift`.")),
        ('step', ("Το στοιχείο εμφανίζεται στη σελίδα και μπαίνει στην αριστερή λίστα. **Εμφάνιση όλων των κρυμμένων** τα επαναφέρει όλα μαζί.",
                  "The item appears on the page and joins the left list. **Show all hidden** brings every hidden item back at once.")),
        ('tip', ("Για να βγάλετε κάτι από το PDF χωρίς να το χάσετε, πατήστε **Απόκρυψη** (δίπλα στο Διπλασιασμός και Διαγραφή) - μπορείτε να το φέρετε πίσω οποιαδήποτε στιγμή. Η απόκρυψη και η επαναφορά αναιρούνται με `Ctrl+Z`.",
                 "To take something off the PDF without losing it, press **Hide** (next to Duplicate and Delete) - you can bring it back any time. Hiding and restoring can be undone with `Ctrl+Z`.")),
        ('h', ("Μεγάλη ή μικρή οθόνη", "Big or small window")),
        ('kb', ("F11", ("Μεγιστοποίηση / επαναφορά του επεξεργαστή (όπως και το κουμπί του παραθύρου)", "Maximize / restore the editor (same as the window's own button)"))),
        ('li', ("Η σελίδα **προσαρμόζεται μόνη της** όταν αλλάζετε το μέγεθος του παραθύρου, μέχρι να κάνετε εσείς ζουμ. Το **Προσαρμογή** την ξαναβάζει σε αυτόματη προσαρμογή.",
                "The page **fits itself** when you resize the window, until you zoom by hand. **Fit** turns automatic fitting back on.")),
        ('li', ("Σε στενό παράθυρο η γραμμή εργαλείων σπάει σε δεύτερη γραμμή και τα πλευρικά πλαίσια παραμένουν ολόκληρα ορατά.",
                "In a narrow window the toolbar wraps onto a second row and the side panels stay fully visible.")),
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
        ('li', ("**Δεδομένα**: **Δείγμα** δείχνει πλασματικές τιμές, **Πραγματικά** τα αποθηκευμένα στοιχεία σας (ξενοδοχείο, δικά σας πεδία, τελευταίος μήνας) και **Ονόματα** μόνο τα ονόματα των στοιχείων.", "**Data**: **Sample** shows made-up values, **Real** shows your saved details (hotel, own fields, latest month) and **Names** shows only the item names.")),
        ('li', ("**↶ / ↷** (`Ctrl+Z` / `Ctrl+Y`) αναιρούν και επαναλαμβάνουν μέχρι 60 αλλαγές.", "**↶ / ↷** (`Ctrl+Z` / `Ctrl+Y`) undo and redo up to 60 changes.")),
        ('li', ("**Αποθήκευση** (`Ctrl+S`) κρατά τη διάταξη. Αν κλείσετε με μη αποθηκευμένες αλλαγές, το πρόγραμμα ρωτά.", "**Save** (`Ctrl+S`) keeps the layout. If you close with unsaved changes, the program asks.")),
        ('li', ("**Επαναφορά αυτού του στοιχείου** γυρίζει στη θέση που θα έβαζε ο επιλεγμένος τρόπος. **Εκκαθάριση όλων** αδειάζει τον καμβά· **Αυτόματη τοποθέτηση** τοποθετεί ξανά όλα με τον επιλεγμένο τρόπο.", "**Reset this item** goes back to where the selected method places it. **Clear all** empties the canvas; **Auto-place items** places everything again with the selected method.")),
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
        ('li', ("**Excel / CSV** (.xls / .xlsx / .csv) → το πλαίσιο ή η κάρτα του Βήματος 1.",
                "**Excel / CSV** (.xls / .xlsx / .csv) → the box or card in Step 1.")),
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
        ('kb', ("{sc:open_excel}", ("Επιλογή αρχείου Excel ή CSV", "Choose the Excel or CSV file"))),
        ('kb', ("{sc:generate}", ("Δημιουργία PDF", "Create PDF"))),
        ('kb', ("{sc:generate_print}", ("Δημιουργία & Εκτύπωση", "Create & Print"))),
        ('kb', ("{sc:help}", ("Άνοιγμα του οδηγού στην ενότητα της τρέχουσας καρτέλας", "Open this guide at the section for the current tab"))),
        ('kb', ("{sc:help_search}", ("Αναζήτηση μέσα στον οδηγό", "Search inside the guide"))),
        ('kb', ("{sc:undo} / {sc:redo}", ("Αναίρεση / επανάληψη σε οποιοδήποτε πεδίο κειμένου (και στο δεξί κλικ)", "Undo / redo in any text field (also in the right-click menu)"))),
        ('h', ("Στο Ιστορικό", "In History")),
        ('kb', ("{sc:hist_open}", ("Άνοιγμα του επιλεγμένου μήνα", "Open the selected month"))),
        ('kb', ("{sc:hist_delete}", ("Διαγραφή του επιλεγμένου μήνα", "Delete the selected month"))),
        ('kb', ("Διπλό κλικ", ("Σε μήνα: άνοιγμα. Σε χώρα: φίλτρο", "On a month: open. On a country: filter"))),
        ('h', ("Στο παράθυρο αποτελέσματος", "In the result window")),
        ('kb', ("{sc:result_print}", ("Εκτύπωση", "Print"))),
        ('kb', ("{sc:result_close}", ("Κλείσιμο", "Close"))),
        ('h', ("Στην προεπισκόπηση PDF", "In the PDF preview")),
        ('kb', ("{sc:pv_zoom_in} / {sc:pv_zoom_out}", ("Μεγέθυνση / σμίκρυνση", "Zoom in / out"))),
        ('kb', ("Ctrl + τροχός", ("Μεγέθυνση προς τον δείκτη του ποντικιού", "Zoom towards the mouse pointer"))),
        ('kb', ("{sc:pv_fit}", ("Προσαρμογή στο παράθυρο", "Fit to the window"))),
        ('kb', ("{sc:pv_prev_page} / {sc:pv_next_page}", ("Προηγούμενη / επόμενη σελίδα", "Previous / next page"))),
        ('kb', ("{sc:pv_close}", ("Κλείσιμο της προεπισκόπησης", "Close the preview"))),
        ('h', ("Στον επεξεργαστή διάταξης", "In the layout editor")),
        ('kb', ("{sc:ed_save}", ("Αποθήκευση", "Save"))),
        ('kb', ("{sc:ed_undo} / {sc:ed_redo}", ("Αναίρεση / επανάληψη", "Undo / redo"))),
        ('kb', ("{sc:ed_duplicate}", ("Διπλασιασμός στοιχείου", "Duplicate item"))),
        ('kb', ("{sc:ed_search}", ("Αναζήτηση στη λίστα στοιχείων", "Search the item list"))),
        ('kb', ("{sc:ed_left} / {sc:ed_right} / {sc:ed_up} / {sc:ed_down}", ("Μετακίνηση 0,5 pt (Shift = 5 pt)", "Move 0.5 pt (Shift = 5 pt)"))),
        ('kb', ("Ctrl + τροχός", ("Μεγέθυνση / σμίκρυνση προς τον δείκτη του ποντικιού", "Zoom in / out towards the mouse pointer"))),
        ('kb', ("{sc:ed_pan} + σύρσιμο", ("Μετακίνηση της σελίδας προς κάθε κατεύθυνση (ο δείκτης γίνεται δείκτης μετακίνησης)", "Pan the page in any direction (the pointer becomes the move cursor)"))),
        ('kb', ("{sc:ed_delete}", ("Διαγραφή στοιχείου που προσθέσατε", "Delete an item you added"))),
        ('kb', ("{sc:ed_deselect}", ("Αποεπιλογή", "Deselect"))),
        ('kb', ("{sc:ed_guide}", ("Άνοιγμα αυτού του οδηγού", "Open this guide"))),
        ('tip', ("Κάθε συντόμευση αλλάζει από τις Ρυθμίσεις → Συντομεύσεις πληκτρολογίου.",
                 "Every shortcut can be changed in Settings → Keyboard shortcuts.")),
        ('go', ("Άνοιγμα των Ρυθμίσεων", "Open Settings"), 'tab:settings'),
    ]),
    dict(id='data', title=("Δεδομένα και αντίγραφα ασφαλείας", "Data and backups"), blocks=[
        ('p', ("Όλα τα δεδομένα του προγράμματος βρίσκονται σε έναν φάκελο (δείτε Ρυθμίσεις → Φάκελος δεδομένων):",
               "All of the program's data lives in one folder (see Settings → Data folder):")),
        ('li', ("Τρέχουσα θέση: `[[data_dir]]`", "Current location: `[[data_dir]]`")),
        ('li', ("`hotel_settings.csv` - τα στοιχεία του ξενοδοχείου.", "`hotel_settings.csv` - the hotel details.")),
        ('li', ("`elstat_history.csv` και `elstat_months.csv` - το ιστορικό.", "`elstat_history.csv` and `elstat_months.csv` - the history.")),
        ('li', ("`elstat_preferences.json` - οι ρυθμίσεις.", "`elstat_preferences.json` - the settings.")),
        ('li', ("`elstat_country_mappings.json` - οι δικές σας αντιστοιχίσεις χωρών.", "`elstat_country_mappings.json` - your own country mappings.")),
        ('li', ("`elstat_detected_layout.json` - οι προσαρμογές της Έξυπνης ανίχνευσης πεδίων (υπάρχει μόνο όσο είναι επιλεγμένη).",
                "`elstat_detected_layout.json` - the fine-tuning of Smart Field Detection (exists only while it is selected).")),
        ('li', ("`elstat_manual_layout.json` και `elstat_template_layout.json` - οι διατάξεις του επεξεργαστή και της έξυπνης μετατόπισης.",
                "`elstat_manual_layout.json` and `elstat_template_layout.json` - the layouts of the editor and of Smart Shift.")),
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
        ('p', ("Διαλέξτε το νέο PDF στην καρτέλα Στοιχεία ξενοδοχείου. Αν οι τιμές πέφτουν εκτός θέσης, δοκιμάστε την **Έξυπνη μετατόπιση** ή την **Έξυπνη ανίχνευση πεδίων**, ή τοποθετήστε τα μόνοι σας με τον **επεξεργαστή διάταξης**. Κάντε πάντα δοκιμαστικό δελτίο πριν υποβάλετε.",
               "Choose the new PDF in the Hotel details tab. If values land out of place, try **Smart Shift** or **Smart Field Detection**, or place them yourself with the **layout editor**. Always make a test form before submitting.")),
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
        ('kb', ("Αναφορά", ("Το template όπως ήταν όταν το δελτίο συμπληρωνόταν σωστά - η έξυπνη μετατόπιση μετρά τις αλλαγές σε σχέση με αυτό.", "The template as it was when the form came out right - Smart Shift measures changes against it."))),
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
        ('h', ("Λέει «ενημερωμένο» ενώ ανέβασα νέα έκδοση", "It says “up to date” although I pushed a new version")),
        ('p', ("Περιμένετε περίπου 5 λεπτά και ξαναπροσπαθήστε, γιατί το GitHub κρατά προσωρινά την παλιά έκδοση. Ελέγξτε επίσης ότι ο αριθμός στην πρώτη γραμμή του `main.py` στο GitHub (`APP_VERSION`) είναι μεγαλύτερος από τον τρέχοντα και ότι το **Αποθετήριο GitHub** και ο **Κλάδος** στις Ρυθμίσεις είναι σωστά.",
               "Wait about 5 minutes and try again, because GitHub keeps serving the old version for a short time. Also check that the number in the first line of `main.py` on GitHub (`APP_VERSION`) is higher than the current one, and that the **GitHub repository** and **Branch** in Settings are correct.")),
        ('h', ("Η εγκατάσταση πακέτων στην πρώτη εκκίνηση απέτυχε", "Package installation on first run failed")),
        ('p', ("Ελέγξτε τη σύνδεση στο internet και ανοίξτε ξανά το πρόγραμμα. Αν το σφάλμα συνεχίζεται, ανοίξτε ένα παράθυρο εντολών στον φάκελο του προγράμματος και γράψτε `pip install -r requirements.txt`.",
               "Check your internet connection and open the program again. If the error continues, open a command window in the program's folder and type `pip install -r requirements.txt`.")),
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
                parts += [sc_text(shortcut), _tx(description)]
            elif blk[0] not in ('go', 'live', 'anchor'):
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
        t.tag_configure('toc', font=(FONT, 10), foreground=MUTED, spacing3=14, lmargin1=0, lmargin2=0)
        t.tag_configure('toclabel', font=(FONT, 10, 'bold'), foreground=NAVY)
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

    def show(self, sid, anchor=None):
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
        heads = [b for b in sec['blocks'] if b[0] == 'h']
        if len(heads) >= 4:            # long sections get a clickable index of their headings
            self._ins(T("Σε αυτή τη σελίδα:", "On this page:") + "  ", ('toc', 'toclabel'))
            for k, hb in enumerate(heads):
                if k:
                    t.insert('end', '  \u00b7  ', ('toc',))
                self._link(_tx(hb[1]), lambda n=k: self._goto_anchor(f'hd{n}', flash=True), ('toc',))
            t.insert('end', '\n', ('toc',))
        hd_no = 0
        step_no = 0
        prev = None
        for blk in sec['blocks']:
            kind = blk[0]
            step_no = step_no + 1 if kind == 'step' else 0
            if kind == 'p':
                self._inline(_tx(blk[1]), ('p',))
                t.insert('end', '\n', ('p',))
            elif kind == 'h':
                t.mark_set(f'anc_hd{hd_no}', 'end-1c')
                t.mark_gravity(f'anc_hd{hd_no}', 'left')
                hd_no += 1
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
                self._ins('\u00a0' + sc_text(shortcut) + '\u00a0', ('kb', 'kbd'))
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
            elif kind == 'anchor':              # an invisible bookmark: a settings (i) jumps to it
                t.mark_set('anc_' + blk[1], 'end-1c')
                t.mark_gravity('anc_' + blk[1], 'left')
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
        if anchor:
            self._goto_anchor(anchor)
            self.after(120, lambda a=anchor, s=sid: self._goto_anchor(a, flash=True) if self.current == s else None)

    def _goto_anchor(self, anchor, flash=False):
        """Scrolls the guide to a bookmark and (once) lightly highlights that entry for a moment."""
        t, mark = self.text, 'anc_' + anchor
        try:
            t.update_idletasks()
            t.yview(mark)
            if flash:
                t.tag_add('flash', mark, f"{mark} lineend")
                t.tag_raise('flash')
                self._fade_flash()
        except tk.TclError:
            pass

    def _fade_flash(self, fade_in=150, hold=1100, fade_out=800):
        """The highlight of the entry the guide jumped to fades in, stays a moment, then fades away (only that text)."""
        t = self.text
        self._flash_token = token = getattr(self, '_flash_token', 0) + 1

        def blend(f):
            try:
                a, b = t.winfo_rgb(CARD), t.winfo_rgb(ACCENT_SOFT)
                return '#%02x%02x%02x' % tuple(int((a[i] + (b[i] - a[i]) * f) / 257) for i in range(3))
            except tk.TclError:
                return ACCENT_SOFT

        def paint(f):
            if token != self._flash_token:
                return False
            try:
                t.tag_configure('flash', background=blend(f))
            except tk.TclError:
                return False
            return True

        def step(i, n, direction, ms, then):
            if not paint(i / n if direction > 0 else 1 - i / n):
                return
            if i < n:
                self.after(max(1, ms // n), step, i + 1, n, direction, ms, then)
            else:
                then()

        def done():
            if token == self._flash_token:
                try:
                    t.tag_remove('flash', '1.0', 'end')
                except tk.TclError:
                    pass
        t.tag_configure('flash', background=CARD)
        step(1, 6, 1, fade_in, lambda: self.after(
            hold, lambda: step(1, 16, -1, fade_out, done) if token == self._flash_token else None))

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


class PageImages:
    """The page pictures of a zoomable PDF view (the preview and the layout editor).

    Zooming used to render and encode the whole page again for every notch of the wheel, which is slow at a high
    zoom. Now a page that is already in memory is scaled instantly (Pillow) and the sharp picture is drawn by a
    background thread once the zoom has settled. Without Pillow the old way (one render per zoom) is used."""
    DELAY = 140          # ms of calm before the sharp picture is rendered
    KEEP = 6             # sharp pictures kept in memory

    def __init__(self, widget, path_fn, ready_cb, version_fn=None):
        self.w = widget
        self.path_fn = path_fn
        self.ready_cb = ready_cb                  # ready_cb(page, zoom): the sharp picture of that page / zoom arrived
        self.version_fn = version_fn or (lambda: 0)
        self._img = OrderedDict()                 # (version, page, zoom) -> Pillow image
        self._legacy = {}                         # (version, page, zoom) -> tk.PhotoImage (no Pillow)
        self._pil_ok = PILImage is not None and PILImageTk is not None
        self._job = None
        self._busy = False
        self._want = None
        self._result = None

    def clear(self):
        self._img.clear()
        self._legacy.clear()

    # ---------------------------------------------------------------- what the view asks for
    def photo(self, page, zoom):
        """(picture, sharp) for the page at this zoom. `sharp` is False while a scaled stand-in is shown."""
        zk = round(zoom, 2)
        key = (self.version_fn(), page, zk)
        if not self._pil_ok:
            return self._legacy_photo(key, page, zk), True
        img = self._img.get(key)
        if img is not None:
            self._img.move_to_end(key)
            photo = self._to_photo(img)
            return (photo, True) if photo is not None else (self._legacy_photo(key, page, zk), True)
        source = self._best_source(key)
        if source is None:                        # nothing in memory yet: draw it now (only the first time)
            img = render_pdf_page_image(self.path_fn(), page, zk)
            if img is None:
                return self._legacy_photo(key, page, zk), True
            self._store(key, img)
            photo = self._to_photo(img)
            return (photo, True) if photo is not None else (self._legacy_photo(key, page, zk), True)
        src_zoom, src = source
        ratio = zk / src_zoom
        size = (max(1, int(round(src.width * ratio))), max(1, int(round(src.height * ratio))))
        resample = getattr(getattr(PILImage, 'Resampling', PILImage), 'BILINEAR')
        try:
            small = src.resize(size, resample, reducing_gap=2.0)
        except TypeError:
            small = src.resize(size, resample)
        photo = self._to_photo(small)
        if photo is None:
            return self._legacy_photo(key, page, zk), True
        self._schedule(page, zk)
        return photo, False

    # ---------------------------------------------------------------- internals
    def _best_source(self, key):
        version, page, zk = key
        found = [(z, img) for (v, p, z), img in self._img.items() if v == version and p == page]
        if not found:
            return None
        bigger = [f for f in found if f[0] >= zk]
        return min(bigger, key=lambda f: f[0]) if bigger else max(found, key=lambda f: f[0])

    def _store(self, key, img):
        self._img[key] = img
        self._img.move_to_end(key)
        while len(self._img) > self.KEEP:
            self._img.popitem(last=False)

    def _to_photo(self, img):
        try:
            return PILImageTk.PhotoImage(img, master=self.w)
        except Exception:
            self._pil_ok = False
            return None

    def _legacy_photo(self, key, page, zoom):
        if key in self._legacy:
            return self._legacy[key]
        png = render_pdf_page_png(self.path_fn(), page, zoom)
        photo = None
        if png:
            try:
                photo = tk.PhotoImage(data=base64.b64encode(png).decode('ascii'))
            except tk.TclError:
                photo = None
        if len(self._legacy) > 8:
            self._legacy.clear()
        self._legacy[key] = photo
        return photo

    def _after(self, ms, fn):
        try:
            return self.w.after(ms, fn)
        except tk.TclError:
            return None

    def _schedule(self, page, zk):
        if self._job is not None:
            try:
                self.w.after_cancel(self._job)
            except tk.TclError:
                pass
        self._job = self._after(self.DELAY, lambda: self._fire(page, zk))

    def _fire(self, page, zk):
        self._job = None
        key = (self.version_fn(), page, zk)
        if key in self._img:
            self.ready_cb(page, zk)
            return
        self._want = (key, self.path_fn(), page, zk)
        if not self._busy:
            self._start()

    def _start(self):
        key, path, page, zk = self._want
        self._want = None
        self._busy = True
        self._result = None

        def work():
            try:
                img = render_pdf_page_image(path, page, zk)
            except Exception:
                img = None
            self._result = (key, page, zk, img)

        threading.Thread(target=work, daemon=True).start()
        self._poll()

    def _poll(self):
        if self._result is None:
            self._after(20, self._poll)
            return
        key, page, zk, img = self._result
        self._result = None
        self._busy = False
        try:
            if img is not None:
                self._store(key, img)
                self.ready_cb(page, zk)
        finally:
            if self._want:
                self._start()


class ShortcutDialog(tk.Toplevel):
    """Asks for a new key combination: the person just presses it, sees at once whether it is free and confirms.
    on_save(text) gets the new shortcut ('' = remove it); the program decides what happens with a clash."""

    def __init__(self, parent, sid, on_save):
        super().__init__(parent)
        self.sid = sid
        self.on_save = on_save
        self.info = SHORTCUT_INFO[sid]
        self.candidate = None
        self.title(T("Αλλαγή συντόμευσης", "Change shortcut"))
        self.configure(bg=BG)
        self.transient(parent.winfo_toplevel())
        self.resizable(False, False)
        card = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill='both', expand=True, padx=14, pady=14)
        inner = tk.Frame(card, bg=CARD)
        inner.pack(padx=26, pady=20)
        group_title = dict(SHORTCUT_GROUPS)[self.info['group']]
        tk.Label(inner, text=T(*self.info['label']), bg=CARD, fg=NAVY, font=(FONT, 13, 'bold'),
                 wraplength=380, justify='left').pack(anchor='w')
        tk.Label(inner, text=T(*group_title), bg=CARD, fg=MUTED, font=(FONT, 9)).pack(anchor='w', pady=(0, 4))
        tk.Label(inner, text=T("Τώρα: ", "Now: ") + shortcut_display(sid), bg=CARD, fg=MUTED,
                 font=(FONT, 9)).pack(anchor='w')
        self.chip = tk.Label(inner, text=T("Πατήστε το νέο συνδυασμό πλήκτρων…", "Press the new key combination…"),
                             bg=ACCENT_SOFT, fg=ACCENT, font=(FONT, 14, 'bold'), padx=18, pady=16, width=28)
        self.chip.pack(fill='x', pady=(12, 8))
        self.msg = tk.Label(inner, text=self._hint(), bg=CARD, fg=MUTED, font=(FONT, 9), wraplength=380,
                            justify='left', anchor='w')
        self.msg.pack(fill='x')
        row = tk.Frame(inner, bg=CARD)
        row.pack(fill='x', pady=(16, 0))
        self.save_btn = ttk.Button(row, text=T("Αποθήκευση", "Save"), command=self._save, state='disabled')
        self.save_btn.pack(side='left')
        ttk.Button(row, text=T("Άκυρο", "Cancel"), command=self.destroy).pack(side='left', padx=8)
        ttk.Button(row, text=T("Χωρίς συντόμευση", "No shortcut"), command=self._remove).pack(side='right')
        self.bind('<KeyPress>', self._on_key)
        self.protocol('WM_DELETE_WINDOW', self.destroy)
        self.update_idletasks()
        try:
            top = parent.winfo_toplevel()
            x = top.winfo_rootx() + (top.winfo_width() - self.winfo_width()) // 2
            y = top.winfo_rooty() + (top.winfo_height() - self.winfo_height()) // 3
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass
        self.after(60, self._take_focus)

    def _take_focus(self):
        try:
            self.grab_set()
            self.focus_force()
        except tk.TclError:
            pass

    def _hint(self):
        if 'hold' in self.info['flags']:
            return T("Πατήστε ένα πλήκτρο (χωρίς Ctrl, Alt ή Shift) και μετά «Αποθήκευση».",
                     "Press one key (without Ctrl, Alt or Shift), then “Save”.")
        if 'mod' in self.info['flags']:
            return T("Πατήστε τον συνδυασμό, π.χ. Ctrl+Κάτι, και μετά «Αποθήκευση».",
                     "Press the combination, e.g. Ctrl+something, then “Save”.")
        return T("Πατήστε τον συνδυασμό και μετά «Αποθήκευση». Για να ξαναδιαλέξετε, απλώς πατήστε άλλον.",
                 "Press the combination, then “Save”. To change your mind just press another one.")

    def _say(self, text, colour):
        self.msg.configure(text=text, fg=colour)

    def _on_key(self, event):
        accel = event_to_accel(event)
        if accel is None:
            if event.keysym not in _MODIFIER_KEYSYMS:
                self._say(T("Αυτό το πλήκτρο δεν μπορεί να χρησιμοποιηθεί. Αν γράφετε με ελληνικό πληκτρολόγιο, "
                            "αλλάξτε προσωρινά σε αγγλικό.",
                            "That key cannot be used. If you are typing with a non-English keyboard layout, "
                            "switch to the English one for a moment."), WARN)
            return 'break'
        mods, key = accel_parse(accel)
        self.chip.configure(text=accel_display(accel))
        problem = None
        if 'hold' in self.info['flags'] and mods:
            problem = T("Εδώ χρειάζεται ένα πλήκτρο μόνο του, χωρίς Ctrl, Alt ή Shift.",
                        "This one needs a single key, without Ctrl, Alt or Shift.")
        elif 'mod' in self.info['flags'] and not (mods & {'Control', 'Alt'}):
            problem = T("Πρέπει να περιέχει Ctrl ή Alt, αλλιώς θα εμπόδιζε το γράψιμο.",
                        "It must include Ctrl or Alt, otherwise it would get in the way of typing.")
        if problem:
            self.candidate = None
            self.save_btn.state(['disabled'])
            self._say(problem, BAD)
            return 'break'
        self.candidate = accel
        self.save_btn.state(['!disabled'])
        other = shortcut_conflict(self.sid, accel)
        if other:
            self._say(T(f"Το χρησιμοποιεί ήδη η ενέργεια «{T(*SHORTCUT_INFO[other]['label'])}». Αν αποθηκεύσετε, "
                        f"εκείνη θα μείνει χωρίς συντόμευση.",
                        f"Already used by “{T(*SHORTCUT_INFO[other]['label'])}”. If you save, "
                        f"that one is left without a shortcut."), WARN)
        else:
            self._say(T("Ελεύθερο.", "Free."), GOOD)
        return 'break'

    def _save(self):
        if self.candidate:
            self.on_save(self.candidate)
        self.destroy()

    def _remove(self):
        self.on_save('')
        self.destroy()


class ConfidencePanel(tk.Frame):
    """Side panel of the previews: a bar and a percentage for every section of the form, then the list of positions
    that need checking.  Picking one in the list (or clicking a section) shows it on the page, outlined."""
    WIDTH = 256

    def __init__(self, parent, on_pick, on_toggle=None, list_height=6):
        super().__init__(parent, bg=CARD)
        self.on_pick, self.on_toggle = on_pick, on_toggle
        self._view = None
        self._busy = False
        self.show_var = tk.BooleanVar(value=True)
        tk.Label(self, text=T("Βεβαιότητα αναγνώρισης", "Recognition confidence"), bg=CARD, fg=NAVY,
                 font=(FONT, 10, 'bold')).pack(anchor="w", padx=12, pady=(12, 4))
        self.rows = tk.Frame(self, bg=CARD)
        self.rows.pack(fill="x", padx=12)
        self.hint = tk.Label(self, text="", bg=CARD, fg=MUTED, font=(FONT, 9), justify="left", wraplength=self.WIDTH + 20)
        self.hint.pack(anchor="w", padx=12, pady=(8, 2))
        tk.Checkbutton(self, text=T("Επισήμανση στη σελίδα", "Highlight on the page"), variable=self.show_var,
                       command=self._toggled, bg=CARD, fg=TEXT, activebackground=CARD, activeforeground=TEXT,
                       selectcolor=FIELD, font=(FONT, 9), anchor="w").pack(anchor="w", padx=8)
        box = tk.Frame(self, bg=CARD)
        box.pack(fill="both", expand=True, padx=8, pady=(2, 6))
        self.tree = ttk.Treeview(box, show='tree', selectmode='browse', height=list_height)
        self.tree.column('#0', width=self.WIDTH, minwidth=200, stretch=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure('check', foreground=WARN)
        self.tree.tag_configure('low', foreground=BAD)
        self.tree.bind('<<TreeviewSelect>>', self._selected)
        self.detail = tk.Label(self, text="", bg=CARD, fg=TEXT, font=(FONT, 9), justify="left", wraplength=self.WIDTH + 20)
        self.detail.pack(anchor="w", padx=12, pady=(0, 8))

    @staticmethod
    def level_color(level):
        return {'ok': GOOD, 'check': WARN, 'low': BAD, 'missing': BAD}.get(level, MUTED)

    def update_view(self, view):
        self._view = view
        for child in self.rows.winfo_children():
            child.destroy()
        self._busy = True
        self.tree.delete(*self.tree.get_children())
        self.detail.configure(text="")
        if not view:
            self.hint.configure(text=T("Η διάταξη ορίστηκε χειροκίνητα - δεν υπάρχει αυτόματη αναγνώριση να αξιολογηθεί.",
                                       "The layout was set by hand - there is no automatic recognition to rate."))
            self._busy = False
            return
        for sec in view['sections']:
            color = self.level_color(sec['level'])
            row = tk.Frame(self.rows, bg=CARD, cursor='hand2' if sec['weak'] else '')
            row.pack(fill="x", pady=(0, 5))
            top = tk.Frame(row, bg=CARD)
            top.pack(fill="x")
            tk.Label(top, text=T(*sec['name']), bg=CARD, fg=TEXT, font=(FONT, 9)).pack(side="left")
            weak = ("  ⚠ " + str(len(sec['weak']))) if sec['weak'] else ""
            tk.Label(top, text=f"{sec['pct']}%{weak}", bg=CARD, fg=color, font=(FONT, 9, 'bold')).pack(side="right")
            bar = tk.Canvas(row, width=self.WIDTH, height=6, bg=TRACK, highlightthickness=0)
            bar.pack(fill="x")
            bar.create_rectangle(0, 0, int(self.WIDTH * sec['pct'] / 100.0), 6, fill=color, outline="")
            if sec['weak']:
                for w in (row, top, bar) + tuple(top.winfo_children()):
                    w.bind('<Button-1>', lambda e, s=sec: self.on_pick(s['weak'][0]) or self._select(s['weak'][0]))
        for f in view['flags']:
            icon = '✗' if f['level'] in ('low', 'missing') else '⚠'
            self.tree.insert('', 'end', iid=f['fid'], text=f"{icon}  {f['label']} · {f['pct']}%",
                             tags=('low' if f['level'] in ('low', 'missing') else 'check',))
        self.hint.configure(text=T("Κάντε κλικ σε μια θέση για να τη δείτε στη σελίδα.", "Click a position to see it on the page.")
                            if view['flags'] else T("✓ Τίποτα δεν χρειάζεται έλεγχο.", "✓ Nothing needs checking."))
        self._busy = False

    def _select(self, fid):
        if self.tree.exists(fid):
            self._busy = True
            self.tree.selection_set(fid)
            self.tree.see(fid)
            self._busy = False

    def _selected(self, _event=None):
        sel = self.tree.selection()
        if sel and not self._busy:
            self.on_pick(sel[0])

    def set_detail(self, flag):
        self.detail.configure(text=f"{flag['label']}: {conf_why(flag['why'])}")

    def _toggled(self):
        if self.on_toggle:
            self.on_toggle(self.show_var.get())


class PDFPreviewDialog(tk.Toplevel):
    """In-app preview of the PDF that is about to be made. Nothing has been saved yet:

    * zoom / pan with the same controls as the layout editor (Ctrl + wheel, - / + / Fit, drag to move),
    * "Edit layout..." opens the layout editor for THIS PDF only (the saved settings are never touched),
    * "Save & continue" asks the caller to save the file for real and to show the Result window,
    * "Close preview" throws the preview away without saving anything.

    Callbacks (all given by the application):
      on_continue(dialog) -> True when everything was saved (the dialog then closes), False to stay here
      on_close()          -> called after the preview was closed without saving
      on_regenerate(layout) -> rebuilds the preview PDF with this temporary layout (None = the normal way);
                               raises when it cannot
    """
    PAD = 28
    MIN_ZOOM, MAX_ZOOM = 0.25, 3.0

    def __init__(self, parent, pdf_path, template_path, on_continue, on_close=None, on_regenerate=None, get_live=None,
                 base_layout=None, get_confidence=None):
        super().__init__(parent)
        self._base_layout = base_layout     # -> the layout "Edit layout" starts from (what this PDF was made with)
        self._get_conf = get_confidence     # -> confidence_view(...) of what is shown, or None (no recognition involved)
        self._conf_view = None
        self._conf_sel = None               # the uncertain item picked in the side list
        self._conf_show = True              # draw the highlights
        self.conf_panel = None
        self.app = parent
        self.pdf_path = pdf_path
        self.template = template_path
        self._on_continue = on_continue
        self._on_close = on_close
        self._on_regenerate = on_regenerate
        self._get_live = get_live       # -> {'values': ..., 'rows': ...} of the form being previewed
        self.layout = None              # temporary layout from the editor; None = the normal way of filling
        self._finished = False
        self._busy = False
        self._editor = None
        self.page = 0
        self.zoom = 1.0
        self.ox = self.oy = self.PAD
        self._sw = self._sh = 1
        self._version = 0               # grows every time the preview PDF is rebuilt (drops cached pictures)
        self._pages = PageImages(self, lambda: self.pdf_path, self._sharp_ready, lambda: self._version)
        self._page_ref = None           # the picture on the canvas (must stay referenced or Tk drops it)
        self._zoom_goal = self._zoom_anchor = self._zoom_job = None
        self._auto_fit = True           # keep the page fitted while the window is resized
        self._fit_pending = True
        self._redraw_job = None
        self._last_size = (0, 0)
        self.sizes = []
        self._read_pdf()

        self.title(T("Προεπισκόπηση PDF", "PDF preview"))
        self.configure(bg=BG)
        self.transient(parent)
        self.minsize(780, 580)
        self._build()
        self._update_edit_state()

        self.protocol('WM_DELETE_WINDOW', self._close)
        self._shortcuts = ShortcutBinder(self)         # changeable in Settings -> Keyboard shortcuts
        self._shortcuts.add('pv_close', lambda e: self._close())
        self._shortcuts.add('pv_zoom_in', lambda e: self._set_zoom(self.zoom * 1.2))
        self._shortcuts.add('pv_zoom_out', lambda e: self._set_zoom(self.zoom / 1.2))
        self._shortcuts.add('pv_fit', lambda e: self._fit())
        self._shortcuts.add('pv_prev_page', lambda e: self._step_page(-1))
        self._shortcuts.add('pv_next_page', lambda e: self._step_page(1))
        self._shortcuts.add('pv_up', lambda e: self.cv.yview_scroll(-3, 'units'))
        self._shortcuts.add('pv_down', lambda e: self.cv.yview_scroll(3, 'units'))
        self._shortcuts.add('pv_left', lambda e: self.cv.xview_scroll(-3, 'units'))
        self._shortcuts.add('pv_right', lambda e: self.cv.xview_scroll(3, 'units'))
        self._shortcuts.apply()
        try:
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            w, h = min(1180, sw - 60), min(920, sh - 80)
            x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
            self.geometry(f"{w}x{h}+{max(0, min(x, sw - w))}+{max(0, min(y, sh - h - 40))}")
        except tk.TclError:
            self.geometry("1100x820")
        self.after(120, self._regrab)

    # ------------------------------------------------------------------ setup
    def _read_pdf(self):
        sizes = []
        try:
            for page in pypdf.PdfReader(self.pdf_path).pages:
                box = page.mediabox
                sizes.append((float(box.width), float(box.height)))
        except Exception:
            pass
        self.sizes = sizes or [(595.0, 842.0)]
        self.page = min(self.page, len(self.sizes) - 1)

    def _regrab(self):
        try:
            if self.winfo_exists():
                self.grab_set()
                self.focus_set()
        except tk.TclError:
            pass

    def _build(self):
        head = tk.Frame(self, bg=HEADER_BG)
        head.pack(fill="x")
        title, subtitle = self._header_texts()
        tk.Label(head, text=title, bg=HEADER_BG, fg=HEADER_FG,
                 font=(FONT, 14, 'bold')).pack(side="left", padx=(18, 12), pady=12)
        tk.Label(head, text=subtitle, bg=HEADER_BG, fg=HEADER_MUTED, font=(FONT, 10)).pack(side="left")

        # ---- toolbar: pages, zoom (same controls as the layout editor), editing
        bar = ttk.Frame(self, padding=(14, 10, 14, 6))
        bar.pack(fill="x")
        if len(self.sizes) > 1:
            self.page_seg = Segmented(
                bar, [(i, T(f"Σελίδα {i + 1}", f"Page {i + 1}")) for i in range(len(self.sizes))],
                self._on_page_seg, self.page)
            self.page_seg.pack(side="left")
            ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=14)
        else:
            self.page_seg = None
        ttk.Button(bar, text="−", width=3, command=lambda: self._set_zoom(self.zoom / 1.2)).pack(side="left")
        self.zoom_lbl = ttk.Label(bar, text="100%", width=6, anchor="center")
        self.zoom_lbl.pack(side="left")
        ttk.Button(bar, text="+", width=3, command=lambda: self._set_zoom(self.zoom * 1.2)).pack(side="left")
        ttk.Button(bar, text=T("Προσαρμογή", "Fit"), command=self._fit).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="100%", width=6, command=lambda: self._set_zoom(1.0)).pack(side="left", padx=(8, 0))

        self._build_toolbar_right(bar)

        # ---- bottom: buttons, then the status line above them
        footer = ttk.Frame(self, padding=(14, 8, 14, 12))
        footer.pack(side="bottom", fill="x")
        self._build_footer(footer)

        status = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        status.pack(side="bottom", fill="x")
        ttk.Label(status, text=T("Ctrl + τροχός = ζουμ · σύρετε για μετακίνηση · διπλό κλικ = προσαρμογή/100%",
                                 "Ctrl + wheel = zoom · drag to move · double-click = fit/100%"),
                  style='Status.TLabel').pack(side="left")
        self.render_note = ttk.Label(status, text=T("● Δεν έχει αποθηκευτεί ακόμα", "● Not saved yet"),
                                     style='Status.TLabel')
        self.render_note.pack(side="right")

        self._build_side()

        # ---- the page
        mid = tk.Frame(self, bg=BORDER)
        mid.pack(fill="both", expand=True, padx=14, pady=(4, 10))
        mid_in = tk.Frame(mid, bg=TAB_BG)
        mid_in.pack(fill="both", expand=True, padx=1, pady=1)
        mid_in.rowconfigure(0, weight=1)
        mid_in.columnconfigure(0, weight=1)
        self.cv = tk.Canvas(mid_in, bg=TAB_BG, highlightthickness=0, width=700, height=520, takefocus=True, cursor='fleur')
        vbar = ttk.Scrollbar(mid_in, orient="vertical", command=self.cv.yview)
        hbar = ttk.Scrollbar(mid_in, orient="horizontal", command=self.cv.xview)
        self.cv.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        self.cv.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        self.cv.bind('<Configure>', self._on_canvas_configure)
        self.cv.bind('<ButtonPress-1>', self._on_press)
        self.cv.bind('<B1-Motion>', self._on_drag)
        self.cv.bind('<ButtonRelease-1>', lambda e: self.cv.configure(cursor='fleur'))
        self.cv.bind('<Double-Button-1>', self._on_double_click)
        self.cv.bind('<Enter>', lambda e: self.cv.configure(cursor='fleur'))
        self.cv.bind('<MouseWheel>', self._on_wheel)
        self.cv.bind('<Button-4>', self._on_wheel)
        self.cv.bind('<Button-5>', self._on_wheel)

    # ---- parts that the inspection dialog (Settings) builds differently
    def _header_texts(self):
        return (T("Προεπισκόπηση PDF", "PDF preview"),
                T("Το αρχείο δεν έχει αποθηκευτεί ακόμα · ελέγξτε το, αλλάξτε τη διάταξη αν χρειάζεται "
                  "και μετά πατήστε «Αποθήκευση & συνέχεια»",
                  "The file is not saved yet · check it, change the layout if needed, "
                  "then press “Save & continue”"))

    def _build_toolbar_right(self, bar):
        # Final verification preview: "Edit layout" only. There is deliberately NO "Rerun Auto-Detection" here -
        # the detection of this report is already final; re-detecting belongs to the inspection dialog in Settings.
        right = ttk.Frame(bar)
        right.pack(side="right")
        self.edit_chip = tk.Label(right, text=T("✎ Προσωρινή διάταξη - μόνο για αυτό το PDF",
                                                "✎ Temporary layout - this PDF only"),
                                  bg=WARN_BG, fg=WARN_FG, font=(FONT, 9, 'bold'), padx=10, pady=4)
        self.discard_btn = ttk.Button(right, text=T("Αναίρεση αλλαγών διάταξης", "Undo layout changes"),
                                      command=self._discard_edits)
        self.edit_btn = ttk.Button(right, text=T("Επεξεργασία διάταξης…", "Edit layout…"),
                                   command=self._edit_layout)
        self.edit_btn.pack(side="left")

    def _build_footer(self, footer):
        ttk.Button(footer, text=T("Άνοιγμα PDF", "Open PDF"),
                   command=lambda: open_file(self.pdf_path)).pack(side="left")
        self.continue_btn = ttk.Button(footer, text=T("Αποθήκευση & συνέχεια στο αποτέλεσμα  →",
                                                      "Save & continue to result  →"),
                                       style='Primary.TButton', command=self._continue)
        self.continue_btn.pack(side="right")
        ttk.Button(footer, text=T("Κλείσιμο προεπισκόπησης", "Close preview"),
                   command=self._close).pack(side="right", padx=8)

    def _build_side(self):
        # Recognition confidence per section + the positions that need a look (only when a way of filling
        # that recognises the template was used)
        if not self._get_conf:
            return
        self._conf_view = self._safe_conf()
        if self._conf_view is None:
            return
        side = tk.Frame(self, bg=BORDER)
        side.pack(side="left", fill="y", padx=(14, 0), pady=(4, 10))
        inner = tk.Frame(side, bg=CARD)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        self.conf_panel = ConfidencePanel(inner, self._conf_pick, self._conf_toggle)
        self.conf_panel.pack(fill="both", expand=True)
        self.conf_panel.update_view(self._conf_view)

    # ---- recognition confidence: list, highlights
    def _safe_conf(self):
        if not self._get_conf:
            return None
        try:
            return self._get_conf()
        except Exception:
            return None

    def _refresh_confidence(self):
        self._conf_view = self._safe_conf()
        if self._conf_sel and not any(f['fid'] == self._conf_sel for f in (self._conf_view or {}).get('flags', [])):
            self._conf_sel = None
        if self.conf_panel is not None:
            self.conf_panel.update_view(self._conf_view)

    def _conf_toggle(self, on):
        self._conf_show = bool(on)
        self._draw()

    def _conf_pick(self, fid):
        """An uncertain item was chosen in the list: show it on the page."""
        flag = next((f for f in (self._conf_view or {}).get('flags', []) if f['fid'] == fid), None)
        if not flag:
            return
        self._conf_sel = fid
        if self.conf_panel is not None:
            self.conf_panel.set_detail(flag)
        page = flag.get('page')
        if page is None:                                    # not found at all: nothing to point at
            self._draw()
            return
        if self.page_seg is not None and page != self.page:
            self.page_seg.set(page, True)
        else:
            self._draw()
        self.after(40, lambda f=flag: self._conf_scroll_to(f))

    def _conf_scroll_to(self, flag):
        box = flag.get('box') or flag.get('band')
        try:
            if not box or flag.get('page') != self.page:
                return
            z, ph = self.zoom, self.sizes[self.page][1]
            cx = self.ox + (box[0] + box[2]) / 2.0 * z
            cy = self.oy + (ph - (box[1] + box[3]) / 2.0) * z
            self.cv.xview_moveto(max(0.0, (cx - self.cv.winfo_width() / 2.0) / self._sw))
            self.cv.yview_moveto(max(0.0, (cy - self.cv.winfo_height() / 2.0) / self._sh))
        except tk.TclError:
            pass

    def _draw_confidence(self):
        """Outlines (amber = please check, red = low) with the percentage, around every uncertain position."""
        view = self._conf_view
        if not view or not self._conf_show:
            return
        cv, z, ph = self.cv, self.zoom, self.sizes[self.page][1]
        for f in view['flags']:
            if f.get('page') != self.page:
                continue
            geo = f.get('box') or f.get('band')
            if not geo:
                continue
            color = '#dc2626' if f['level'] in ('low', 'missing') else '#f59e0b'
            x0, y0, x1, y1 = geo
            cx0, cx1 = self.ox + x0 * z - 3, self.ox + x1 * z + 3
            cy0, cy1 = self.oy + (ph - y1) * z - 3, self.oy + (ph - y0) * z + 3
            sel = f['fid'] == self._conf_sel
            cv.create_rectangle(cx0, cy0, cx1, cy1, outline=color, width=4 if sel else 2,
                                dash=() if f.get('box') else (6, 4), tags='conf')
            text = f"{f['pct']}%"
            tw = 10 + 6 * len(text)
            cv.create_rectangle(cx1 - tw, cy0 - 14, cx1, cy0, fill=color, outline=color, tags='conf')
            cv.create_text(cx1 - tw / 2.0, cy0 - 7, text=text, fill='white', font=(FONT, 8, 'bold'), tags='conf')

    def _resolve_base_layout(self):
        if self._base_layout:
            try:
                layout = self._base_layout()
                if layout is not None:
                    return layout
            except Exception:
                pass
        return load_manual_layout()

    def _update_edit_state(self):
        edited = self.layout is not None
        for widget in (self.edit_chip, self.discard_btn):
            widget.pack_forget()
        if edited:
            self.edit_chip.pack(side="left", padx=(0, 8), before=self.edit_btn)
            self.discard_btn.pack(side="left", padx=(0, 8), before=self.edit_btn)

    # ------------------------------------------------------------------ drawing (centred, like the editor)
    def _sharp_ready(self, page, zoom):
        """The background thread finished the sharp picture of a page: swap it in without redrawing everything."""
        try:
            if page != self.page or zoom != round(self.zoom, 2) or not self.cv.find_withtag('pageimg'):
                return
            photo, sharp = self._pages.photo(self.page, self.zoom)
            if photo is not None and sharp:
                self._page_ref = photo
                self.cv.itemconfigure('pageimg', image=photo)
        except tk.TclError:
            pass

    def _draw(self):
        cv, z = self.cv, self.zoom
        pw, ph = self.sizes[self.page]
        w, h = pw * z, ph * z
        cw, ch = max(cv.winfo_width(), 100), max(cv.winfo_height(), 100)
        # the page is centred horizontally AND vertically whenever it is smaller than the window
        self.ox = max(self.PAD, int((cw - w) / 2))
        self.oy = max(self.PAD, int((ch - h) / 2))
        self._sw = max(self.ox * 2 + w, cw)
        self._sh = max(self.oy * 2 + h, ch)
        cv.delete('all')
        cv.configure(scrollregion=(0, 0, self._sw, self._sh))
        cv.create_rectangle(self.ox + 4, self.oy + 4, self.ox + w + 4, self.oy + h + 4, fill=BORDER, outline='')
        photo, _sharp = self._pages.photo(self.page, z)
        if photo:
            self._page_ref = photo
            cv.create_image(self.ox, self.oy, image=photo, anchor='nw', tags='pageimg')
        else:
            cv.create_rectangle(self.ox, self.oy, self.ox + w, self.oy + h, fill='white', outline='')
            cv.create_text(self.ox + w / 2, self.oy + h / 2, width=max(200, w - 60), justify='center', fill='#475569',
                           font=(FONT, 11),
                           text=T("Δεν ήταν δυνατή η εμφάνιση της σελίδας σε αυτόν τον υπολογιστή "
                                  "(εγκαταστήστε το pypdfium2: pip install pypdfium2). "
                                  "Μπορείτε να ανοίξετε το PDF εξωτερικά και μετά να συνεχίσετε.",
                                  "The page cannot be drawn on this computer "
                                  "(install pypdfium2: pip install pypdfium2). "
                                  "You can open the PDF externally and then continue."))
        cv.create_rectangle(self.ox, self.oy, self.ox + w, self.oy + h, outline=ZONE_IDLE_LINE)
        self.zoom_lbl.configure(text=f"{round(z * 100)}%")
        self._draw_confidence()

    def _on_canvas_configure(self, event):
        if self._fit_pending and event.width > 300 and event.height > 200:
            self._fit_pending = False
            self._last_size = (event.width, event.height)
            self.after(10, self._fit)
            return
        size = (event.width, event.height)
        if abs(size[0] - self._last_size[0]) > 2 or abs(size[1] - self._last_size[1]) > 2:
            self._last_size = size
            if self._redraw_job:
                self.after_cancel(self._redraw_job)
            self._redraw_job = self.after(80, self._after_resize)

    def _after_resize(self):
        self._redraw_job = None
        if self._auto_fit:
            self._fit()
        else:
            self._draw()

    # ------------------------------------------------------------------ view controls
    def _zoom_wheel(self, delta, anchor):
        """Ctrl + wheel. A burst of wheel events is merged into ONE zoom step per screen refresh (smooth, no backlog)."""
        steps = max(-3.0, min(3.0, delta / 120.0))
        base = self._zoom_goal if self._zoom_goal is not None else self.zoom
        self._zoom_goal = min(max(base * (1.12 ** steps), self.MIN_ZOOM), self.MAX_ZOOM)
        self._zoom_anchor = anchor
        if self._zoom_job is None:
            self._zoom_job = self.after(15, self._apply_zoom_goal)

    def _apply_zoom_goal(self):
        self._zoom_job = None
        goal, self._zoom_goal = self._zoom_goal, None
        if goal is not None:
            self._set_zoom(goal, self._zoom_anchor)

    def _set_zoom(self, zoom, anchor=None):
        """Zoom around `anchor` (widget coordinates; the middle of the window if omitted)."""
        zoom = round(min(max(zoom, self.MIN_ZOOM), self.MAX_ZOOM), 2)      # whole percents: pictures are reused
        if abs(zoom - self.zoom) < 1e-4:
            return
        cv = self.cv
        ax, ay = anchor if anchor else (cv.winfo_width() / 2, cv.winfo_height() / 2)
        px = (cv.canvasx(ax) - self.ox) / self.zoom          # the spot of the page under the pointer
        py = (cv.canvasy(ay) - self.oy) / self.zoom
        self._auto_fit = False
        self.zoom = zoom
        self._draw()
        cv.xview_moveto(max(0.0, (self.ox + px * zoom - ax) / self._sw))
        cv.yview_moveto(max(0.0, (self.oy + py * zoom - ay) / self._sh))

    def _fit(self):
        pw, ph = self.sizes[self.page]
        cw, ch = max(self.cv.winfo_width(), 300), max(self.cv.winfo_height(), 200)
        zoom = min((cw - 2 * self.PAD) / pw, (ch - 2 * self.PAD) / ph)
        self.zoom = min(max(round(zoom, 2), self.MIN_ZOOM), self.MAX_ZOOM)
        self._auto_fit = True
        self._draw()
        self.cv.xview_moveto(0)
        self.cv.yview_moveto(0)

    def _on_double_click(self, _event):
        if self._auto_fit and abs(self.zoom - 1.0) > 0.02:
            self._set_zoom(1.0, (_event.x, _event.y))
        else:
            self._fit()

    def _on_wheel(self, event):
        num = getattr(event, 'num', None)
        delta = 120 if num == 4 else -120 if num == 5 else getattr(event, 'delta', 0)
        if not delta:
            return 'break'
        if event.state & 0x4:                              # Ctrl + wheel = zoom
            self._zoom_wheel(delta, (event.x, event.y))
        elif event.state & 0x1:                            # Shift + wheel = sideways
            self.cv.xview_scroll(-3 if delta > 0 else 3, 'units')
        else:
            self.cv.yview_scroll(-3 if delta > 0 else 3, 'units')
        return 'break'

    def _on_press(self, event):
        self.cv.focus_set()
        self.cv.scan_mark(event.x, event.y)
        self.cv.configure(cursor='fleur')

    def _on_drag(self, event):
        self.cv.scan_dragto(event.x, event.y, gain=1)

    def _on_page_seg(self, value):
        self.page = value
        if self._auto_fit:
            self._fit()
        else:
            self._draw()
            self.cv.yview_moveto(0)

    def _step_page(self, step):
        if self.page_seg is None:
            return
        new = min(max(self.page + step, 0), len(self.sizes) - 1)
        if new != self.page:
            self.page_seg.set(new, True)

    # ------------------------------------------------------------------ editing the layout (temporary)
    def _edit_layout(self):
        if self._editor is not None:
            try:
                if self._editor.winfo_exists():
                    self._editor.lift()
                    self._editor.focus_force()
                    return
            except tk.TclError:
                pass
        base = self.layout if self.layout is not None else self._resolve_base_layout()
        live = None
        if self._get_live:
            try:
                live = self._get_live()
            except Exception:
                live = None
        editor = LayoutEditor(self.app, self.template, on_save=self._layout_applied,
                              layout=base, temporary=True, parent=self, live=live)
        self._editor = editor
        editor.lift()
        editor.focus_force()
        editor.bind('<Destroy>', lambda e, ed=editor: e.widget is ed and self.after(60, self._regrab))

    def _rebuild(self, layout, failure_title):
        """Rebuilds the preview PDF with `layout` (None = the normal way). True when it worked."""
        if not self._on_regenerate:
            return False
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            self._on_regenerate(layout)
        except PermissionError:
            messagebox.showerror(failure_title, T("Δεν μπορώ να ενημερώσω την προεπισκόπηση. Κλείστε το PDF αν είναι "
                                                  "ανοιχτό σε άλλο πρόγραμμα και δοκιμάστε ξανά.",
                                                  "I can't update the preview. Close the PDF if it is open in "
                                                  "another program and try again."), parent=self)
            return False
        except Exception as e:
            messagebox.showerror(failure_title, T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"),
                                 parent=self)
            return False
        finally:
            self.config(cursor='')
        self.layout = layout
        self._version += 1
        self._pages.clear()
        self._read_pdf()
        self._update_edit_state()
        self._refresh_confidence()
        self._draw()
        return True

    def _layout_applied(self, layout):
        """'Save changes' in the editor: the preview is rebuilt, the saved settings stay as they are."""
        self._rebuild(layout, T("Επεξεργασία διάταξης", "Edit layout"))

    def _discard_edits(self):
        if self.layout is None:
            return
        self._rebuild(None, T("Αναίρεση αλλαγών", "Undo changes"))

    # ------------------------------------------------------------------ leaving
    def _continue(self):
        if self._finished or self._busy:
            return
        self._busy = True
        self.config(cursor='watch')
        try:
            done = self._on_continue(self) if self._on_continue else True
        finally:
            self._busy = False
            try:
                self.config(cursor='')
            except tk.TclError:
                pass
        if done:
            self._finished = True
            self.destroy()

    def _close(self):
        if self._finished or self._busy:
            return
        self._finished = True
        callback = self._on_close
        self.destroy()
        if callback:
            callback()


class DetectInspectDialog(PDFPreviewDialog):
    """Context A - "Configure & Inspect Auto-Detection…" (opened from Settings, never while a report is made).
    Shows what Smart Field Detection found on the blank template and how sample data fits, before any real report.
    Its own actions: Rerun Auto-Detection, and Open in Layout Editor (fine-tuning that is remembered).
    The final verification preview of a report (PDFPreviewDialog) has no Rerun: by then the detection is final."""

    def __init__(self, app, template_path):
        self.template = template_path
        self.work_layout = self.report = self.warning = None
        self.tuned = detect_cache_saved()
        self._marker_fid = None
        self._detect_editor = None
        self.data_mode = 'sample'              # 'sample' (made-up) | 'real' (saved details + latest month)
        self.real_month = None
        self.tmp_dir = tempfile.mkdtemp(prefix='elstat_detect_')
        pdf_path = os.path.join(self.tmp_dir, "ELSTAT_detection_test.pdf")
        try:
            self._render(template_path, pdf_path, None)
        except Exception:
            shutil.rmtree(self.tmp_dir, ignore_errors=True)
            raise
        super().__init__(app, pdf_path, template_path, None, on_close=self._cleanup,
                         get_confidence=self._conf_for_detect)
        self.title(T("Έξυπνη ανίχνευση πεδίων - έλεγχος", "Smart Field Detection - inspect"))
        self._update_note()
        self._fill_list()

    def _update_note(self):
        if self.data_mode == 'real':
            m = self.real_month
            when = f" ({m[1]:02d}/{m[0]})" if m else ""
            self.render_note.configure(text=T(f"● Πραγματικά δεδομένα{when} - δεν δημιουργείται δελτίο",
                                              f"● Real saved data{when} - no report is made"))
        else:
            self.render_note.configure(text=T("● Δείγμα δεδομένων - δεν είναι πραγματικό δελτίο", "● Sample data - not a real report"))

    # ---- the test render (made-up data, or the saved hotel details and latest month)
    def _render(self, template_path, pdf_path, override):
        if self.data_mode == 'real':
            sample, mon, rows, _has = real_report_inputs()
            self.real_month = (int(mon['year']), int(mon['month']))
        else:
            sample, mon = manual_sample_inputs()
            rows = [(name, 1234, 5678) for name in sorted(MAPPED_TARGETS)]
        res = process_elstat_report(None, template_path, pdf_path, sample, mon, rows_override=rows,
                                    save_history=False, pdf_mode='detect', detect_layout=override)
        if res.get('pdf_mode_used') != 'detect':
            raise DetectError(res.get('layout_warning') or T("Η ανίχνευση απέτυχε.", "The detection failed."))
        self.work_layout, self.report, self.warning = res['layout_used'], res['detect_report'], res.get('layout_warning')

    def _cleanup(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def _conf_for_detect(self):
        return confidence_view((self.report or {}).get('confidence'), self.work_layout) if self.report else None

    # ---- parts of the base dialog that differ here
    def _header_texts(self):
        return (T("Έξυπνη ανίχνευση πεδίων - έλεγχος", "Smart Field Detection - inspect"),
                T("Δείτε τι βρέθηκε στο άδειο template και πώς χωράνε δείγματα ή πραγματικά δεδομένα · δεν δημιουργείται πραγματικό δελτίο",
                  "See what was found on the blank template and how sample or real data fits · no real report is made here"))

    def _build_toolbar_right(self, bar):
        right = ttk.Frame(bar)
        right.pack(side="right")
        self.edit_chip = self.discard_btn = self.edit_btn = None
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=14)
        ttk.Label(bar, text=T("Δεδομένα:", "Data:")).pack(side="left", padx=(0, 6))
        self.data_seg = Segmented(bar, (('sample', T("Δείγμα", "Sample")), ('real', T("Πραγματικά", "Real"))),
                                  self._on_data_mode, self.data_mode)
        self.data_seg.pack(side="left")
        self.rerun_btn = ttk.Button(right, text=T("🔄 Επανάληψη ανίχνευσης", "🔄 Rerun Auto-Detection"), command=self._rerun)
        self.rerun_btn.pack(side="left", padx=(0, 8))
        self.open_editor_btn = ttk.Button(right, text=T("Άνοιγμα στον επεξεργαστή διάταξης…", "Open in Layout Editor…"),
                                          style='PrintSm.TButton', command=self._open_editor)
        self.open_editor_btn.pack(side="left")

    def _build_footer(self, footer):
        ttk.Button(footer, text=T("Άνοιγμα PDF", "Open PDF"), command=lambda: open_file(self.pdf_path)).pack(side="left")
        self.continue_btn = None
        ttk.Button(footer, text=T("Κλείσιμο", "Close"), style='Primary.TButton', command=self._close).pack(side="right")

    def _update_edit_state(self):
        pass

    def _build_side(self):
        side = tk.Frame(self, bg=BORDER)
        side.pack(side="left", fill="y", padx=(14, 0), pady=(4, 10))
        inner = ttk.Frame(side, style='Card.TFrame')
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        self.conf_panel = ConfidencePanel(inner, self._conf_pick, self._conf_toggle, list_height=4)
        self.conf_panel.pack(fill="x")
        self.summary_var = tk.StringVar()
        ttk.Label(inner, textvariable=self.summary_var, style='CardHint.TLabel', justify="left",
                  wraplength=290).pack(anchor="w", padx=12, pady=(12, 6))
        box = ttk.Frame(inner, style='Card.TFrame')
        box.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        ttk.Style(self).configure('Detect.Treeview', rowheight=24)
        self.tree = ttk.Treeview(box, show='tree', selectmode='browse', height=8, style='Detect.Treeview')
        self.tree.column('#0', width=300, minwidth=220, stretch=True)
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure('group', font=(FONT, 10, 'bold'), foreground=NAVY)
        self.tree.tag_configure('hidden', foreground=MUTED)
        self.tree.bind('<<TreeviewSelect>>', self._on_pick)
        self.info_var = tk.StringVar(value=T("Επιλέξτε ένα πεδίο για να δείτε πού μπαίνει.", "Pick a field to see where it goes."))
        ttk.Label(inner, textvariable=self.info_var, style='CardHint.TLabel', justify="left",
                  wraplength=290).pack(anchor="w", padx=12, pady=(0, 12))

    def _fill_list(self):
        self._refresh_confidence()
        self.tree.delete(*self.tree.get_children())
        items = self.work_layout['items']
        placed = {}
        for fid, it in items.items():
            if not it.get('parent'):
                placed.setdefault(manual_group(fid, it), []).append((fid, it))
        for gid, (gel, gen) in MANUAL_GROUPS:
            if gid not in placed:
                continue
            self.tree.insert('', 'end', iid='g:' + gid, text=T(gel, gen), open=gid not in ('rows', 'totals'), tags=('group',))
            for fid, it in placed[gid]:
                shown = it.get('visible', True)
                self.tree.insert('g:' + gid, 'end', iid=fid, text=("✓  " if shown else "○  ") + manual_label(fid, it, items),
                                 tags=() if shown else ('hidden',))
        text_ = describe_detect_report(self.report, self.tuned)
        if self.warning:
            text_ += "\n" + self.warning
        text_ += getattr(self, '_rerun_detail', '')
        self.summary_var.set(text_)

    def _on_pick(self, _event=None):
        sel = self.tree.selection()
        fid = sel[0] if sel else None
        it = self.work_layout['items'].get(fid) if fid else None
        if not it:
            return
        self._marker_fid = fid
        align = {'left': T("αριστερά", "left"), 'center': T("κέντρο", "centre"), 'right': T("δεξιά", "right")}.get(it['align'], '')
        self.info_var.set(T(f"Σελίδα {it['page'] + 1} · X {it['x']:.1f} · Y {it['y']:.1f} pt · {align} · μέγεθος {it['size']:g}",
                            f"Page {it['page'] + 1} · X {it['x']:.1f} · Y {it['y']:.1f} pt · {align} · size {it['size']:g}")
                          + ("" if it.get('visible', True) else T("\nΔεν βρέθηκε - κρυμμένο. Τοποθετήστε το στον επεξεργαστή.",
                                                                  "\nNot found - hidden. Place it in the editor.")))
        if self.page_seg is not None and it['page'] != self.page:
            self.page_seg.set(it['page'], True)
        else:
            self._draw()

    def _draw(self):
        super()._draw()
        it = self.work_layout['items'].get(self._marker_fid) if (self.work_layout and self._marker_fid) else None
        if not it or it['page'] != self.page or not it.get('visible', True):
            return
        z, ph = self.zoom, self.sizes[self.page][1]
        x = self.ox + it['x'] * z
        if self._marker_fid.startswith('rows'):
            self.cv.create_line(x, self.oy, x, self.oy + ph * z, fill=ACCENT, width=2, dash=(5, 3), tags='marker')
            return
        y, r = self.oy + (ph - it['y']) * z, 8
        self.cv.create_oval(x - r, y - r, x + r, y + r, outline=ACCENT, width=2, tags='marker')
        self.cv.create_line(x - r - 4, y, x + r + 4, y, fill=ACCENT, tags='marker')
        self.cv.create_line(x, y - r - 4, x, y + r + 4, fill=ACCENT, tags='marker')

    def _refresh_pdf(self):
        self._version += 1
        self._pages.clear()
        self._read_pdf()
        self._fill_list()
        self._draw()

    # ---- the two actions that belong only here
    def _rerun(self):
        title = T("Επανάληψη ανίχνευσης", "Rerun Auto-Detection")
        if detect_cache_saved() and not messagebox.askyesno(
                title, T("Οι προσαρμογές σας από τον επεξεργαστή θα διαγραφούν και η ανίχνευση θα ξεκινήσει από την αρχή.\n\nΣυνέχεια;",
                         "Your fine-tuning from the editor will be discarded and the detection will start from scratch.\n\nContinue?"),
                default='no', parent=self):
            return
        backup = None
        try:
            with open(detect_layout_file(), 'rb') as f:
                backup = f.read()
        except OSError:
            pass
        delete_detect_cache()
        was_deep = detect_deep_saved(self.template)
        set_detect_deep(self.template, True)        # from now on this template is scanned with the deep pass
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            self._render(self.template, self.pdf_path, None)
        except Exception as e:
            if not was_deep:
                set_detect_deep(self.template, False)
            if backup is not None:          # nothing was lost: put the fine-tuning back
                try:
                    with open(detect_layout_file(), 'wb') as f:
                        f.write(backup)
                except OSError:
                    pass
            messagebox.showerror(title, T(f"Η ανίχνευση δεν ολοκληρώθηκε:\n{e}", f"The detection did not finish:\n{e}"), parent=self)
            return
        finally:
            self.config(cursor='')
        self.tuned = False
        self._marker_fid = None
        self._refresh_pdf()
        self._report_rerun(title)

    def _report_rerun(self, title):
        """Quiet feedback: a short note in the status bar that fades by itself - no pop-up to close."""
        deep = (self.report or {}).get('deep')
        if not deep:
            return
        a, b = conf_percent(deep['before']), conf_percent(deep['after'])
        n = len(deep['improved'])
        view = confidence_view(self.report.get('confidence'), self.work_layout) or {'flags': []}
        left = len(view['flags'])
        if n:
            msg = T(f"✓ Εμπιστοσύνη {a}% → {b}% · βελτιώθηκαν {n} πεδία", f"✓ Confidence {a}% → {b}% · {n} field{'s' if n != 1 else ''} improved")
        else:
            msg = T(f"✓ Εμπιστοσύνη {b}% - δεν υπάρχει κάτι ακόμα να βελτιωθεί αυτόματα",
                    f"✓ Confidence {b}% - nothing more to improve automatically")
        if left:
            msg += T(f" · {left} προς έλεγχο", f" · {left} to check")
        # the details stay in the left panel, for anyone who wants them
        detail = [f"{manual_label(f['fid'], {}, None)}: {conf_percent(f['before'])}% → {conf_percent(f['after'])}%"
                  for f in sorted(deep['improved'], key=lambda f: f['before'])[:6]]
        self._rerun_detail = ("\n" + T("Βελτιώθηκαν: ", "Improved: ") + " · ".join(detail)) if detail else ""
        self._fill_list()
        self._toast(msg)

    def _toast(self, text, ms=6000):
        try:
            if getattr(self, '_toast_job', None):
                self.after_cancel(self._toast_job)
            self.render_note.configure(text=text)
            self._toast_job = self.after(ms, self._toast_done)
        except tk.TclError:
            pass

    def _toast_done(self):
        self._toast_job = None
        try:
            self._update_note()
        except tk.TclError:
            pass

    def _open_editor(self):
        ed = self._detect_editor
        try:
            if ed is not None and ed.winfo_exists():
                ed.lift()
                ed.focus_force()
                return
        except tk.TclError:
            pass
        editor = LayoutEditor(self.app, self.template, on_save=self._editor_saved, layout=self.work_layout,
                              parent=self, context='detect')
        self._detect_editor = editor
        editor.lift()
        editor.focus_force()
        editor.bind('<Destroy>', lambda e, ed=editor: e.widget is ed and self.after(60, self._regrab))

    def _on_data_mode(self, value):
        if value == self.data_mode:
            return
        title = T("Πραγματικά δεδομένα", "Real data")
        if value == 'real':
            try:
                has = real_report_inputs()[3]
            except Exception:
                has = False
            if not has:
                self.data_seg.set(self.data_mode)
                messagebox.showinfo(title, T("Δεν υπάρχουν ακόμα αποθηκευμένα δεδομένα (στοιχεία ξενοδοχείου ή μήνας). "
                                             "Θα εμφανιστούν εδώ μόλις δημιουργήσετε ένα δελτίο.",
                                             "There is no saved data yet (hotel details or a month). "
                                             "It will show up here once you have made a report."), parent=self)
                return
        previous = self.data_mode
        self.data_mode = value
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            self._render(self.template, self.pdf_path, None)
            self._refresh_pdf()
        except Exception as e:
            self.data_mode = previous
            self.data_seg.set(previous)
            messagebox.showerror(title, T(f"Δεν ήταν δυνατή η εμφάνιση των δεδομένων:\n{e}",
                                          f"Could not show the data:\n{e}"), parent=self)
        finally:
            self.config(cursor='')
        self._update_note()

    def _editor_saved(self):
        """The editor saved its fine-tuning to the detection cache: draw the sample again with it."""
        self.config(cursor='watch')
        try:
            self._render(self.template, self.pdf_path, None)
            self.tuned = True
            self._refresh_pdf()
        except Exception as e:
            messagebox.showerror(T("Έξυπνη ανίχνευση", "Smart Field Detection"),
                                 T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"), parent=self)
        finally:
            self.config(cursor='')


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
        tree.tag_configure('bold', font=(FONT, 10, 'bold'), background=HEAD_BG)
        tree.tag_configure('plain', background=CARD)
        for label, c, n, bold in result['summary']:
            tree.insert('', 'end', text=label, values=(c, n), tags=('bold' if bold else 'plain',))
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._shortcuts = ShortcutBinder(self)         # changeable in Settings -> Keyboard shortcuts
        self._shortcuts.add('result_print', lambda e: self.print_now())
        self._shortcuts.add('result_close', lambda e: self.destroy())
        self._shortcuts.apply()
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

    def show_option(self, val, visible):
        lbl = self._labels.get(val)
        if lbl is None:
            return
        if visible and not lbl.winfo_manager():
            lbl.pack(side="left", padx=(1, 1), pady=1)
            for v, l in self._labels.items():      # keep the original order
                if l.winfo_manager():
                    l.pack_forget()
                    l.pack(side="left", padx=(1, 1), pady=1)
        elif not visible and lbl.winfo_manager():
            lbl.pack_forget()

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
    MIN_ZOOM, MAX_ZOOM = 0.4, 3.0
    NUDGE_STEPS = ("0.1", "0.5", "1", "2", "5")
    QUICK_COLORS = ('#000000', '#1f2937', '#1d4ed8', '#047857', '#b91c1c', '#7c3aed')

    def __init__(self, app, template_path, on_save=None, layout=None, temporary=False, parent=None, live=None,
                 context='manual'):
        # temporary=True: opened from the PDF preview - "Save changes" hands the layout to the preview and
        # nothing is written to the saved manual layout
        host = parent or app
        super().__init__(host)
        self.app = app
        self.temporary = temporary
        self.context = context       # 'manual' = saved as the manual layout | 'detect' = the fine-tuning of Smart Field Detection
        self._mode_ref = None
        self.template = template_path
        self.on_save = on_save
        self.layout = copy.deepcopy(layout) if layout is not None else load_manual_layout()
        self.page = 0
        self.zoom = 1.0
        self.ox = self.PAD
        self.sel = None
        self.group_ids = []          # items of a whole section picked in the list (moved together)
        self.group_gid = None
        self.picks = []              # what Ctrl+click picked: item ids, 'g:<section>' or 'var:<variable>'
        self._anchor = None          # the row of the list where Shift+click ranges start
        self._last_synced = set()    # what the list was last set to by the program (to recognise its own echo)
        self.dirty = False
        self._loading = False
        self._sync = False
        self._drag = None
        self._space = False
        self._panning = False
        self._pages = PageImages(self, lambda: self.template, self._sharp_ready)
        self._page_ref = None           # the picture on the canvas (must stay referenced or Tk drops it)
        self._zoom_goal = self._zoom_anchor = self._zoom_job = None
        self._fit_pending = True
        self._redraw_job = None
        self._last_width = 0
        self._last_height = 0
        self._auto_fit = True           # the page follows the window size until the person zooms by hand
        self._fallback_texts = None
        self.undo, self.redo = [], []
        self._snap_key, self._snap_time, self._stagger = None, 0.0, 0
        self.q_var = tk.StringVar()
        self.tip_var = tk.StringVar()
        self.sizes = [(595.0, 842.0), (595.0, 842.0)]
        self.row_ys = [{}, {}]
        self._read_template()
        self.sample = manual_sample_values()
        # opened from the preview: show the REAL values and country rows of this report instead of made-up ones
        self.real_values = self.real_rows = None
        self._real_loading = False
        if temporary and live and live.get('values'):
            self.real_values = dict(live['values'])
            self.real_rows = dict(live.get('rows') or {})
        self.families = manual_font_families()
        self.data_mode = tk.StringVar(value='real' if self.real_values is not None else 'sample')
        self.status_var = tk.StringVar()

        self.configure(bg=BG)
        self.minsize(900, 600)
        self._build()
        self._fill_tree()
        self._load_props()
        self._update_title()
        self.protocol('WM_DELETE_WINDOW', self._close)
        sc = self._shortcuts = ShortcutBinder(self)    # changeable in Settings -> Keyboard shortcuts
        sc.add('ed_save', lambda e: self.save())
        sc.add('ed_delete', self._on_delete_key)
        sc.add('ed_undo', lambda e: None if self._typing() else self._undo())
        sc.add('ed_redo', lambda e: None if self._typing() else self._redo())
        sc.add('ed_duplicate', lambda e: None if self._typing() else self._duplicate())
        sc.add('ed_guide', lambda e: self._open_guide())
        sc.add('ed_search', lambda e: (self.search_entry.focus_set(), self.search_entry.select_range(0, 'end'), 'break')[2])
        sc.add('ed_deselect', lambda e: self._select(None))
        sc.add('ed_zoom_in', lambda e: None if self._typing() else self._set_zoom(self.zoom * 1.2))
        sc.add('ed_zoom_out', lambda e: None if self._typing() else self._set_zoom(self.zoom / 1.2))
        sc.add('ed_fit', lambda e: None if self._typing() else self._fit())
        sc.add('ed_pan', self._space_down, release=self._space_up)
        for sid, direction in (('ed_left', (-1, 0)), ('ed_right', (1, 0)), ('ed_up', (0, 1)), ('ed_down', (0, -1))):
            sc.add(sid, lambda e, d=direction: self._on_arrow(e, d))
        sc.apply()
        self.bind('<FocusOut>', lambda e: self._space_up())
        self.bind('<F11>', self._toggle_maximize)
        self.resizable(True, True)
        try:
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            w, h = min(1340, sw - 60), min(880, sh - 90)
            x = host.winfo_rootx() + (host.winfo_width() - w) // 2
            y = host.winfo_rooty() + (host.winfo_height() - h) // 2
            self.geometry(f"{w}x{h}+{max(0, min(x, sw - w))}+{max(0, min(y, sh - h - 40))}")
        except tk.TclError:
            self.geometry("1280x820")
        self.after(150, self._try_grab)

    # ------------------------------------------------------------------ setup
    def _place_sash(self):
        try:
            h = self.vpane.winfo_height()
            if h > 200:
                self.vpane.sashpos(0, int(h * 0.58))
            else:
                self.after(200, self._place_sash)
        except tk.TclError:
            pass

    def _toggle_maximize(self, _event=None):
        """Maximize / restore the editor window (F11, the header button or a double-click on the header)."""
        try:
            self.state('normal' if self.state() == 'zoomed' else 'zoomed')
        except tk.TclError:                         # Linux window managers
            try:
                self.attributes('-zoomed', not bool(self.attributes('-zoomed')))
            except tk.TclError:
                pass
        return 'break'

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
        self.head_title = tk.Label(head, text=self._base_title(), bg=HEADER_BG, fg=HEADER_FG,
                                   font=(FONT, 14, 'bold'))
        self.head_title.pack(side="left", padx=(18, 12), pady=12)
        self.head_note = tk.Label(head, text=(T("Αλλαγές μόνο για αυτό το PDF - οι αποθηκευμένες ρυθμίσεις δεν αλλάζουν · ",
                               "Changes for this PDF only - your saved settings stay unchanged · ")
                             if self.temporary else self._context_note())
                 + T("Σύρετε κάθε στοιχείο εκεί που ανήκει στο έντυπο · βελάκια = λεπτομερής μετακίνηση",
                     "Drag each item to where it belongs on the form · arrow keys = fine movement"),
                 bg=HEADER_BG, fg=HEADER_MUTED, font=(FONT, 10))
        self.head_note.pack_forget()
        self.head_note.configure(justify="left", anchor="w")
        self.head_note.pack(side="left", fill="x", expand=True)

        def fit_head(_e=None):
            room = head.winfo_width() - self.head_title.winfo_reqwidth() - 50
            self.head_note.configure(wraplength=max(220, room))
        head.bind('<Configure>', fit_head)

        # ---- toolbar
        bar_outer = ttk.Frame(self, padding=(14, 10, 14, 6))
        bar_outer.pack(fill="x")
        bar = ttk.Frame(bar_outer)
        bar_r = ttk.Frame(bar_outer)
        bar.grid(row=0, column=0, sticky="w")
        bar_r.grid(row=0, column=1, sticky="e")
        bar_outer.columnconfigure(1, weight=1)
        self._bar_wrapped = False

        def reflow_bar(_e=None):
            need = bar.winfo_reqwidth() + bar_r.winfo_reqwidth() + 40
            wrap = bar_outer.winfo_width() > 1 and bar_outer.winfo_width() < need
            if wrap != self._bar_wrapped:
                self._bar_wrapped = wrap
                if wrap:
                    bar_r.grid(row=1, column=0, columnspan=2, sticky="e", pady=(6, 0))
                else:
                    bar_r.grid(row=0, column=1, columnspan=1, sticky="e", pady=0)
        bar_outer.bind('<Configure>', reflow_bar)
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
        ttk.Label(bar, text=T("Δεδομένα:", "Data:")).pack(side="left", padx=(0, 6))
        self.data_seg = Segmented(bar, (('sample', T("Δείγμα", "Sample")), ('real', T("Πραγματικά", "Real")),
                                        ('names', T("Ονόματα", "Names"))), self._on_data_mode, self.data_mode.get())
        self.data_seg.pack(side="left")
        ttk.Button(bar_r, text=T("Αποθήκευση αλλαγών", "Save changes") if self.temporary else T("Αποθήκευση", "Save"),
                   style='PrintSm.TButton', command=self.save).pack(side="right")
        ttk.Button(bar_r, text=T("Εκκαθάριση όλων", "Clear all"), style='Danger.TButton',
                   command=self._wipe_all).pack(side="right")
        ttk.Button(bar_r, text=T("Αυτόματη τοποθέτηση", "Auto-place items"),
                   command=self._auto_place).pack(side="right", padx=(0, 8))

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
        tk.Label(tipbar, text=T("F11 = μεγιστοποίηση", "F11 = maximize"), bg=ACCENT_SOFT, fg=MUTED,
                 font=(FONT, 9), padx=6).pack(side="right")
        self.tip_lbl = tk.Label(tipbar, textvariable=self.tip_var, bg=ACCENT_SOFT, fg=TEXT, font=(FONT, 9),
                                anchor="w", justify="left", padx=12, pady=6)
        self.tip_lbl.pack(side="left", fill="x", expand=True)
        tipbar.bind('<Configure>', lambda e: self.tip_lbl.configure(wraplength=max(300, e.width - 320)))

        # ---- body: field list | page | properties
        body = ttk.PanedWindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=14, pady=(4, 10))

        left_outer, left = self._card(body)
        self.left_title_var = tk.StringVar()
        ttk.Label(left, textvariable=self.left_title_var, style='CardBold.TLabel').pack(
            anchor="w", padx=14, pady=(12, 0))
        ttk.Label(left, text=T("Μόνο ό,τι τυπώνεται στο PDF. Τα κρυμμένα και όσα δεν χρησιμοποιούνται είναι δεξιά.",
                               "Only what is printed on the PDF. Hidden and unused items are listed on the right."),
                  style='CardHint.TLabel', wraplength=250, justify="left").pack(anchor="w", padx=14, pady=(0, 6))
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
        self.tree = ttk.Treeview(tree_box, show='tree', selectmode='browse', height=8, style='Editor.Treeview')
        self.tree.column('#0', width=250, minwidth=180, stretch=True)
        sb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure('group', font=(FONT, 10, 'bold'), foreground=NAVY)
        self.tree.tag_configure('hidden', foreground=MUTED)
        self.tree.tag_configure('var', foreground=ACCENT)
        self.tree.bind('<<TreeviewSelect>>', self._on_tree_select)
        self.tree.bind('<Control-Button-1>', self._on_tree_ctrl)
        self.tree.bind('<Shift-Button-1>', lambda e: self._on_tree_shift(e))
        self.tree.bind('<Control-Shift-Button-1>', lambda e: self._on_tree_shift(e, add=True))
        lrow = ttk.Frame(left, style='Card.TFrame')
        lrow.pack(fill="x", padx=14, pady=(12, 4), side="bottom", before=tree_box)
        ttk.Button(lrow, text=T("+ Προσθήκη στοιχείου…", "+ Add item…"), style='PrintSm.TButton',
                   command=self._add_item_dialog).pack(side="left")
        ttk.Button(lrow, text=T("+ Δικό μου κείμενο", "+ My own text"), command=self._add_custom).pack(side="left", padx=8)
        lrow2 = ttk.Frame(left, style='Card.TFrame')
        lrow2.pack(fill="x", padx=14, pady=(0, 12), side="bottom", before=lrow)
        self.dup_btn = ttk.Button(lrow2, text=T("Διπλασιασμός", "Duplicate"), command=self._duplicate)
        self.dup_btn.pack(side="left")
        self.hide_btn = ttk.Button(lrow2, text=T("Απόκρυψη", "Hide"), command=self._hide_selected)
        self.hide_btn.pack(side="left", padx=8)
        self.del_btn = ttk.Button(lrow2, text=T("Διαγραφή", "Delete"), command=self._delete_selected)
        self.del_btn.pack(side="left")
        body.add(left_outer, weight=0)

        mid = tk.Frame(body, bg=BORDER)
        mid_in = tk.Frame(mid, bg=TAB_BG)
        mid_in.pack(fill="both", expand=True, padx=1, pady=1)
        mid_in.rowconfigure(0, weight=1)
        mid_in.columnconfigure(0, weight=1)
        self.cv = tk.Canvas(mid_in, bg=TAB_BG, highlightthickness=0, width=320, height=300, takefocus=True)
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
        self.cv.bind('<Enter>', lambda e: None if self._typing() else self.cv.focus_set())
        body.add(mid, weight=1)

        right_col = ttk.Frame(body)
        self.vpane = vpane = ttk.PanedWindow(right_col, orient="vertical")
        vpane.pack(fill="both", expand=True)
        props_outer, props_card = self._card(vpane)
        self.props_scroll = ScrollFrame(props_card)
        self.props_scroll.pack(fill="both", expand=True)
        self.props_scroll.canvas.configure(bg=CARD, width=310, height=200)
        self.props_scroll.body.configure(style='Card.TFrame')
        self._build_props(self.props_scroll.body)
        vpane.add(props_outer, weight=3)
        hidden_outer, hidden_card = self._card(vpane)
        self._build_hidden_panel(hidden_card)
        vpane.add(hidden_outer, weight=2)
        body.add(right_col, weight=0)
        self.after(200, self._place_sash)

    def _build_props(self, parent):
        self.empty_lbl = ttk.Label(
            parent, style='CardHint.TLabel', justify="left", wraplength=250,
            text=T("Επιλέξτε ένα στοιχείο από τη λίστα ή πατήστε πάνω του στη σελίδα.\n\n"
                   "Μετά σύρετέ το εκεί που θέλετε και διαλέξτε γραμματοσειρά, μέγεθος και χρώμα.",
                   "Choose an item from the list or click it on the page.\n\n"
                   "Then drag it where you want it and pick the font, size and colour."))
        self.empty_lbl.pack(anchor="nw", padx=18, pady=18)
        self._empty_default = self.empty_lbl.cget('text')
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

    # ------------------------------------------------------------------ hidden / unused window
    def _build_hidden_panel(self, parent):
        self._h_last = set()
        self._n_hidden = 0
        self.hid_title_var = tk.StringVar()
        self.hq_var = tk.StringVar()
        ttk.Label(parent, textvariable=self.hid_title_var, style='CardBold.TLabel').pack(
            anchor="w", padx=14, pady=(12, 0))
        ttk.Label(parent, text=T("Κρυμμένα = τοποθετημένα αλλά δεν τυπώνονται. Χωρίς χρήση = διαθέσιμα στοιχεία που δεν "
                                 "έχετε βάλει στη σελίδα. Διπλό κλικ για να μπουν στο PDF.",
                                 "Hidden = placed but not printed. Not on the form = available items you have not "
                                 "placed yet. Double-click one to put it on the PDF."),
                  style='CardHint.TLabel', wraplength=280, justify="left").pack(anchor="w", padx=14, pady=(2, 6))
        srow = ttk.Frame(parent, style='Card.TFrame')
        srow.pack(fill="x", padx=14, pady=(0, 6))
        ttk.Label(srow, text=T("Αναζήτηση:", "Search:"), style='CardHint.TLabel').pack(side="left", padx=(0, 6))
        ttk.Entry(srow, textvariable=self.hq_var).pack(side="left", fill="x", expand=True)
        clear = tk.Label(srow, text="\u2715", bg=CARD, fg=MUTED, font=(FONT, 10), cursor='hand2', padx=6)
        clear.pack(side="left")
        clear.bind('<Button-1>', lambda e: self.hq_var.set(''))
        self.hq_var.trace_add('write', lambda *a: self._fill_hidden())
        box = ttk.Frame(parent, style='Card.TFrame')
        box.pack(fill="both", expand=True, padx=8)
        self.htree = ttk.Treeview(box, show='tree', selectmode='extended', height=4, style='Editor.Treeview')
        self.htree.column('#0', width=270, minwidth=200, stretch=True)
        hsb = ttk.Scrollbar(box, orient="vertical", command=self.htree.yview)
        self.htree.configure(yscrollcommand=hsb.set)
        self.htree.pack(side="left", fill="both", expand=True)
        hsb.pack(side="right", fill="y")
        self.htree.tag_configure('group', font=(FONT, 10, 'bold'), foreground=NAVY)
        self.htree.tag_configure('hint', foreground=MUTED, font=(FONT, 9, 'italic'))
        self.htree.bind('<<TreeviewSelect>>', self._on_hidden_select)
        self.htree.bind('<Double-1>', self._on_hidden_double)
        self.htree.bind('<Return>', lambda e: (self._put_selected(), 'break')[1])
        brow = ttk.Frame(parent, style='Card.TFrame')
        brow.pack(fill="x", padx=14, pady=(10, 12), side="bottom", before=box)
        self.put_btn = ttk.Button(brow, text=T("Βάλε στο PDF", "Put on the PDF"), style='PrintSm.TButton',
                                  command=self._put_selected)
        self.put_btn.pack(side="left")
        self.showall_btn = ttk.Button(brow, text=T("Εμφάνιση όλων των κρυμμένων", "Show all hidden"),
                                      command=self._show_all_hidden)
        self.showall_btn.pack(side="left", padx=8)
        self._update_hidden_buttons()

    def _hid_label(self, fid, it):
        base = manual_label(fid, it, self.layout['items'])
        if it.get('country'):
            return base
        return base + T(f"  · Σελ. {it.get('page', 0) + 1}", f"  · Pg {it.get('page', 0) + 1}")

    def _unused_catalog(self):
        """Everything that could be put on the form but is not on it at all (not even hidden)."""
        items = self.layout['items']
        used_src = {i['source'] for i in items.values() if i.get('source')}
        out = []
        for key, label, gid in manual_source_catalog():
            if gid == 'country' or key in items or key in used_src:
                continue
            out.append((key, label, gid))
        for f in MANUAL_FIELDS:
            if f[1] == 'rows' and f[0] not in items:
                out.append((f[0], f[2], 'rows'))
        return out

    def _fill_hidden(self):
        if not hasattr(self, 'htree'):
            return
        tv, items = self.htree, self.layout['items']
        q = _fold(self.hq_var.get().strip())
        was = {}

        def grab(parent):
            for r in tv.get_children(parent):
                if r.startswith(('hg', 'ug')):
                    was[r] = str(tv.item(r, 'open')).lower() in ('1', 'true')
                    grab(r)
        grab('')
        hidden_all = [(f, i) for f, i in items.items() if not i.get('split') and not i.get('visible', True)]
        unused_all = self._unused_catalog()
        kinds = dict(MANUAL_GROUPS)
        kinds['cfield'] = ("Μεταβλητές", "Variables")
        kinds['rows'] = ("Γραμμές χωρών", "Country rows")
        nh, nu = len(hidden_all), len(unused_all)
        self._n_hidden = nh
        self.hid_title_var.set(T(f"Κρυμμένα και αχρησιμοποίητα ({nh + nu})", f"Hidden and not used ({nh + nu})"))
        tv.delete(*tv.get_children())
        tv.insert('', 'end', iid='hg', text=T(f"Κρυμμένα στο PDF ({nh})", f"Hidden on the PDF ({nh})"),
                  open=was.get('hg', True) or bool(q), tags=('group',))
        shown = 0
        for fid, it in hidden_all:
            label = self._hid_label(fid, it)
            if q and q not in _fold(label):
                continue
            tv.insert('hg', 'end', iid='h:' + fid, text=label)
            shown += 1
        if not shown:
            tv.insert('hg', 'end', iid='hnone', tags=('hint',),
                      text=T("(κανένα - όλα εμφανίζονται)" if not nh else "(κανένα αποτέλεσμα)",
                             "(none - everything is shown)" if not nh else "(no matches)"))
        tv.insert('', 'end', iid='ug', text=T(f"Δεν υπάρχουν στο έντυπο ({nu})", f"Not on the form ({nu})"),
                  open=was.get('ug', True) or bool(q), tags=('group',))
        by_group = {}
        for key, label, gid in unused_all:
            text = T(*label)
            if q and q not in _fold(text + ' ' + T(*kinds.get(gid, (gid, gid)))):
                continue
            by_group.setdefault(gid, []).append((key, text))
        order = [g for g, _n in MANUAL_GROUPS] + ['cfield']
        for gid in order + [g for g in by_group if g not in order]:
            if gid not in by_group:
                continue
            gi = 'ug:' + gid
            tv.insert('ug', 'end', iid=gi, text=f"{T(*kinds.get(gid, (gid, gid)))}  ({len(by_group[gid])})",
                      open=was.get(gi, False) or bool(q), tags=('group',))
            for key, text in by_group[gid]:
                tv.insert(gi, 'end', iid='u:' + key, text=text)
        if not by_group:
            tv.insert('ug', 'end', iid='unone', tags=('hint',),
                      text=T("(τίποτα - όλα είναι στη σελίδα)" if not nu else "(κανένα αποτέλεσμα)",
                             "(nothing - everything is on the page)" if not nu else "(no matches)"))
        self._h_last = set()
        self._sync_hidden_selection()

    def _sync_hidden_selection(self):
        """The window follows the selected item (a hidden one is highlighted there)."""
        if not hasattr(self, 'htree'):
            return
        tv = self.htree
        want = 'h:' + self.sel if self.sel else None
        cur = set(tv.selection())
        if want and tv.exists(want):
            if cur != {want}:
                tv.selection_set(want)
                tv.see(want)
        elif cur:
            tv.selection_remove(*cur)
        self._h_last = set(tv.selection())
        self._update_hidden_buttons()

    def _update_hidden_buttons(self):
        if not hasattr(self, 'put_btn'):
            return
        picked = [i for i in self.htree.selection() if i[:2] in ('h:', 'u:')]
        self.put_btn.state(['!disabled'] if picked else ['disabled'])
        self.showall_btn.state(['!disabled'] if self._n_hidden else ['disabled'])

    def _on_hidden_select(self, _event=None):
        chosen = self.htree.selection()
        now = set(chosen)
        if now == self._h_last:             # the program's own echo
            return
        self._h_last = now
        self._update_hidden_buttons()
        items = self.layout['items']
        if len(chosen) == 1 and chosen[0].startswith('h:') and chosen[0][2:] in items:
            if chosen[0][2:] != self.sel:
                self._select(chosen[0][2:], from_tree=True)
                self._sync_tree_selection()
            self.status_var.set(T("Κρυμμένο στοιχείο - ενεργοποιήστε το «Εμφάνιση στο PDF» ή πατήστε «Βάλε στο PDF».",
                                  "Hidden item - tick “Show in the PDF” or press “Put on the PDF”."))
        elif len(chosen) == 1 and chosen[0].startswith('u:'):
            self.status_var.set(T("Δεν υπάρχει ακόμα στο έντυπο - διπλό κλικ ή «Βάλε στο PDF» για να μπει στη σελίδα.",
                                  "Not on the form yet - double-click or press “Put on the PDF” to place it on the page."))

    def _on_hidden_double(self, event):
        iid = self.htree.identify_row(event.y)
        if iid[:2] in ('h:', 'u:'):
            self._put_selected([iid])
            return 'break'

    def _put_selected(self, iids=None):
        iids = [i for i in (iids if iids is not None else self.htree.selection()) if i[:2] in ('h:', 'u:')]
        if not iids:
            return
        items = self.layout['items']
        hidden = [i[2:] for i in iids if i.startswith('h:') and i[2:] in items]
        new = [i[2:] for i in iids if i.startswith('u:')]
        if hidden:
            self._snap(None)
            for fid in hidden:
                items[fid]['visible'] = True
            self._set_dirty()
            self._fill_tree()
            self._draw_items()
        for key in new:
            self._add_source_item(key)
        if hidden and not new:
            self._select(hidden[-1])
            self._reveal()
        n = len(hidden) + len(new)
        self.status_var.set(T(f"{n} στοιχεία μπήκαν στο PDF.", f"{n} item(s) put on the PDF."))

    def _show_all_hidden(self):
        self._put_selected(['h:' + f for f, i in self.layout['items'].items()
                            if not i.get('split') and not i.get('visible', True)])

    def _hide_selected(self):
        items = self.layout['items']
        ids = [self.sel] if self.sel else [f for f in self.group_ids if f in items]
        ids = [f for f in ids if f in items and not items[f].get('split') and items[f].get('visible', True)]
        if not ids:
            return
        self._snap(None)
        for f in ids:
            items[f]['visible'] = False
        self._set_dirty()
        self._fill_tree()
        self._draw_items()
        if self.sel:
            self._load_props()
        self._sync_hidden_selection()
        self.status_var.set(T(f"{len(ids)} στοιχεία κρύφτηκαν - βρίσκονται στο παράθυρο «Κρυμμένα» δεξιά.",
                              f"{len(ids)} item(s) hidden - find them in the “Hidden” window on the right."))

    # ------------------------------------------------------------------ list of items
    def _tree_text(self, fid, it):
        src = str(it.get('source') or '')
        if src.startswith(CF_PREFIX) and fid in self.layout['items']:
            same = [f for f, i in self.layout['items'].items() if i.get('source') == src]
            return ("" if it.get('visible', True) else "○  ") + \
                T(f"Σελίδα {it.get('page', 0) + 1} · θέση {same.index(fid) + 1}",
                  f"Page {it.get('page', 0) + 1} · placement {same.index(fid) + 1}")
        return ("" if it.get('visible', True) else "○  ") + manual_label(fid, it, self.layout['items'])

    def _present(self, fid, it):
        """True when the item is printed on the PDF (a split item counts when any of its pieces is)."""
        if it.get('split'):
            return any(i.get('parent') == fid and i.get('visible', True) for i in self.layout['items'].values())
        return bool(it.get('visible', True))

    def _fill_tree(self):
        self.tree.delete(*self.tree.get_children())
        q = _fold(self.q_var.get().strip())
        items = self.layout['items']
        groups = dict(MANUAL_GROUPS)
        rows = []

        def root_of(it):
            while it.get('parent') in items and not it.get('country'):
                it = items[it['parent']]
            return it
        for fid, it in items.items():
            if not self._present(fid, it):          # hidden items live in the "Hidden & not used" window
                continue
            gid = manual_group(fid, it if not (it.get('parent') in items and not it.get('country')) else root_of(it))
            if q and q not in _fold(manual_label(fid, it, self.layout['items']) + ' ' + T(*groups.get(gid, (gid, gid)))):
                continue
            rows.append((fid, it, gid))
        used = {g for _f, _i, g in rows}
        counts = {}
        for _f, it, gid in rows:
            if not it.get('split'):
                counts[gid] = counts.get(gid, 0) + 1
        cf_names = {CF_PREFIX + c['id']: c['name'] for c in load_custom_fields()}
        placed_vars = {str(root_of(it).get('source')) for _f, it, g in rows if g == 'vars'}
        shown_vars = [k for k in cf_names if k in placed_vars]
        for gid, (gel, gen) in MANUAL_GROUPS:
            if gid in used:
                self.tree.insert('', 'end', iid='g:' + gid, text=f"{T(gel, gen)}  ({counts.get(gid, 0)})",
                                 open=(gid not in ('totals',)) or bool(q), tags=('group',))
        for key in shown_vars:      # a variable is listed once it has a placement; its placements hang under it
            self.tree.insert('g:vars', 'end', iid='var:' + key, text=cf_names[key], open=True, tags=('var',))
        for fid, it, gid in rows:
            parent = 'g:' + gid
            if gid == 'vars':
                vsrc = str(root_of(it).get('source'))
                parent = 'var:' + vsrc
                if not self.tree.exists(parent):        # a variable that was deleted but is still on the form
                    self.tree.insert('g:vars', 'end', iid=parent, text=manual_source_label(vsrc),
                                     open=True, tags=('hidden',))
            if not q and it.get('parent') in items and self.tree.exists(it['parent']):       # pieces of a split item sit under it
                parent = it['parent']
            self.tree.insert(parent, 'end', iid=fid, text=self._tree_text(fid, it))
        n_on = sum(1 for i in items.values() if not i.get('split') and i.get('visible', True))
        self.left_title_var.set(T(f"Στοιχεία στο PDF ({n_on})", f"Items on the PDF ({n_on})"))
        self._sync_tree_selection()
        self._fill_hidden()

    def _sync_tree_selection(self):
        self._sync = True
        try:
            if self.sel and self.tree.exists(self.sel):
                self.tree.selection_set(self.sel)
                self.tree.see(self.sel)
                self._last_synced = {self.sel}
            elif self.picks and any(self.tree.exists(p) for p in self.picks):
                ex = [p for p in self.picks if self.tree.exists(p)]
                self.tree.selection_set(*ex)
                self.tree.see(ex[-1])
                self._last_synced = set(ex)
            else:
                self.tree.selection_remove(*self.tree.selection())
                self._last_synced = set()
        finally:
            self._sync = False

    def _gid_of(self, fid, it):
        """The section (group) an item belongs to - the pieces of a split item belong to their whole's section."""
        items = self.layout['items']
        while it.get('parent') in items and not it.get('country'):
            fid = it['parent']
            it = items[fid]
        return manual_group(fid, it)

    def _select_group(self, gid):
        """A plain click on a section of the list: all its items are picked."""
        self.picks = ['g:' + gid]
        self._apply_picks()

    def _toggle_pick(self, iid):
        """Ctrl+click: add / remove one item, one whole section or one variable to the picked set."""
        items = self.layout['items']
        if not (iid in items or iid.startswith('g:') or iid.startswith('var:')):
            return
        if not self.picks and self.sel:
            self.picks = [self.sel]                  # Ctrl+click adds to the item that is already selected
        if iid in self.picks:
            self.picks.remove(iid)
        else:
            self.picks.append(iid)
        self.sel = None
        if not self.picks:
            self.group_ids, self.group_gid = [], None
            self._draw_selection()
            self._load_props()
            self._sync_tree_selection()
        elif len(self.picks) == 1 and self.picks[0] in items and not items[self.picks[0]].get('split'):
            self._select(self.picks[0])              # only one item left: an ordinary selection again
        else:
            self._apply_picks()

    def _apply_picks(self):
        items = self.layout['items']
        ids = []

        def add(f):
            if f in items and f not in ids:
                ids.append(f)
        for p in self.picks:
            if p.startswith('g:'):
                for f, i in items.items():
                    if not i.get('split') and self._gid_of(f, i) == p[2:]:
                        add(f)
            elif p.startswith('var:'):
                for f, i in items.items():
                    root = items.get(i.get('parent'), i) if i.get('parent') and not i.get('country') else i
                    if root.get('source') == p[4:] and not i.get('split'):
                        add(f)
            elif p in items:
                if items[p].get('split'):
                    for f, i in items.items():
                        if i.get('parent') == p:
                            add(f)
                else:
                    add(p)
        if not ids:
            self.picks, self.group_ids, self.group_gid = [], [], None
            self.status_var.set(T("Δεν υπάρχουν στοιχεία για επιλογή εδώ.", "There are no items to pick here."))
            self._draw_selection()
            self._load_props()
            return
        self.sel = None
        self.group_ids = ids
        self.group_gid = self.picks[0][2:] if len(self.picks) == 1 and self.picks[0].startswith('g:') else 'multi'
        pages = [items[f]['page'] for f in ids]
        if self.page not in pages:
            self.page = pages[0]
            if self.page_seg:
                self.page_seg.set(self.page)
            self._draw_all()
        else:
            self._draw_selection()
        self._load_props()
        self._sync_tree_selection()
        self.status_var.set(T(f"{len(ids)} στοιχεία επιλεγμένα - σύρετε οποιοδήποτε για να μετακινηθούν όλα μαζί "
                              "(ή βελάκια, Shift = 5 pt). Ctrl+κλικ = προσθήκη/αφαίρεση, Shift+κλικ στη λίστα = περιοχή, Esc = καθαρισμός.",
                              f"{len(ids)} items picked - drag any of them to move them all together "
                              "(or arrow keys, Shift = 5 pt). Ctrl+click = add/remove, Shift+click in the list = range, Esc = clear."))
        self.cv.focus_set()

    def _on_tree_ctrl(self, event):
        iid = self.tree.identify_row(event.y)
        if iid and iid != 'varhint':
            self._toggle_pick(iid)
            self._anchor = iid
        return 'break'

    def _visible_rows(self):
        """The rows of the list in the order they are shown (rows inside a closed section are skipped)."""
        out = []

        def walk(parent):
            for r in self.tree.get_children(parent):
                out.append(r)
                if self.tree.item(r, 'open'):
                    walk(r)
        walk('')
        return out

    def _on_tree_shift(self, event, add=False):
        """Shift+click: everything from the last clicked row to this one (like Windows). Ctrl+Shift adds to the picks."""
        iid = self.tree.identify_row(event.y)
        if not iid or iid == 'varhint':
            return 'break'
        items = self.layout['items']
        anchor = self._anchor if self._anchor and self.tree.exists(self._anchor) else (self.sel or iid)
        order = self._visible_rows()
        if anchor not in order or iid not in order:
            anchor = iid
        a, b = sorted((order.index(anchor), order.index(iid)))
        rows = [r for r in order[a:b + 1]
                if r != 'varhint' and (not r.startswith('g:') or r in (anchor, iid))]
        picks = []
        if add:
            picks = list(self.picks) if self.picks else ([self.sel] if self.sel else [])
        for r in rows:
            if r not in picks:
                picks.append(r)
        self.sel = None
        self.picks = picks
        if len(picks) == 1 and picks[0] in items and not items[picks[0]].get('split'):
            keep = self._anchor
            self._select(picks[0])
            self._anchor = keep
        else:
            self._apply_picks()
        return 'break'

    def _on_tree_select(self, _event=None):
        if self._sync:
            return
        chosen = self.tree.selection()
        now = set(chosen)
        if now == self._last_synced:       # the program's own echo of a selection it just made (would loop forever)
            return
        self._last_synced = now
        if len(chosen) > 1:
            return
        if chosen:
            self._anchor = chosen[0]
        if chosen and chosen[0].startswith('g:'):
            self._select_group(chosen[0][2:])
            return
        if chosen and chosen[0].startswith('var:'):
            key = chosen[0][4:]
            placed = [f for f, i in self.layout['items'].items() if i.get('source') == key]
            if placed:
                self._select(placed[0], from_tree=True)
            else:
                self._add_source_item(key)
                self.status_var.set(T(f"Η μεταβλητή «{manual_source_label(key)}» μπήκε στη σελίδα - σύρετέ την όπου θέλετε. "
                                      "(Ctrl+D για ακόμα ένα αντίγραφο)",
                                      f"Variable “{manual_source_label(key)}” was put on the page - drag it where you want. "
                                      "(Ctrl+D for another copy)"))
            return
        if chosen and chosen[0] in self.layout['items'] and chosen[0] != self.sel:
            self._select(chosen[0], from_tree=True)

    # ------------------------------------------------------------------ drawing
    def _px(self, x):
        return self.ox + x * self.zoom

    def _py(self, y):
        return self.PAD + (self.sizes[self.page][1] - y) * self.zoom

    def _sharp_ready(self, page, zoom):
        """The background thread finished the sharp picture of a page: swap it in without redrawing everything."""
        try:
            if page != self.page or zoom != round(self.zoom, 2) or not self.cv.find_withtag('pageimg'):
                return
            photo, sharp = self._pages.photo(self.page, self.zoom)
            if photo is not None and sharp:
                self._page_ref = photo
                self.cv.itemconfigure('pageimg', image=photo)
        except tk.TclError:
            pass

    def _draw_all(self):
        cv, z = self.cv, self.zoom
        pw, ph = self.sizes[self.page]
        w, h = pw * z, ph * z
        cw = max(cv.winfo_width(), 100)
        self.ox = max(self.PAD, int((cw - w) / 2))
        cv.delete('all')
        cv.configure(scrollregion=(0, 0, max(self.ox * 2 + w, cw), h + 2 * self.PAD))
        cv.create_rectangle(self.ox + 4, self.PAD + 4, self.ox + w + 4, self.PAD + h + 4, fill=BORDER, outline='')
        photo, _sharp = self._pages.photo(self.page, z)
        if photo:
            self._page_ref = photo
            cv.create_image(self.ox, self.PAD, image=photo, anchor='nw', tags='pageimg')
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
        mode = self.data_mode.get()
        values = self.real_values if (mode == 'real' and self.real_values is not None) else self.sample
        sample_on = mode == 'sample' or (mode == 'real' and self.real_values is not None)
        live = mode == 'real' and self.real_rows is not None

        def real(target, column):       # the real number of one country ('' = not in this report)
            d = self.real_rows.get(target)
            return str(d[column]) if d else ''

        parent, country, source = it.get('parent'), it.get('country'), it.get('source')
        if country:
            y0 = self.row_ys[it['page']].get(country)
            if y0 is None:
                return
            if not sample_on:
                shown = ''
            elif live:
                shown = real(country, 'clients' if parent.endswith('_c') else 'nights')
            else:
                shown = '12' if parent.endswith('_c') else '34'
            entries = [(it['x'], y0 + it['y'], shown)]
        elif parent:
            entries = [(it['x'], it['y'], manual_piece_text(self.layout['items'], it, values) if sample_on else '')]
        elif fid.startswith('rows'):
            ys = list(self.row_ys[it['page']].values()) or [400.0]
            if sample_on and live:
                column = 'clients' if fid.endswith('_c') else 'nights'
                entries = [(it['x'], y + it['y'], real(c, column))
                           for c, y in self.row_ys[it['page']].items() if c in self.real_rows]
                if not entries:
                    entries = [(it['x'], ys[0] + it['y'], '')]
            elif sample_on:
                shown = '12' if fid.endswith('_c') else '34'
                entries = [(it['x'], y + it['y'], shown) for y in ys]
            else:
                entries = [(it['x'], ys[0] + it['y'], '')]
        elif it.get('custom'):
            entries = [(it['x'], it['y'], it.get('text', ''))]
        elif source:
            if source.startswith('country:'):
                if not sample_on:
                    shown = ''
                elif live:
                    shown = real(source.split(':')[1], 'clients' if source.endswith(':c') else 'nights')
                else:
                    shown = '1234' if source.endswith(':c') else '5678'
            else:
                shown = values.get(source, '') if sample_on else ''
            entries = [(it['x'], it['y'], shown)]
        else:
            entries = [(it['x'], it['y'], values.get(fid, '') if sample_on else '')]
        for x, y, txt in entries:
            if not txt:           # nothing to show: the item's name stands in, so it can still be found
                cv.create_text(self._px(x), self._py(y) + 2, text=manual_label(fid, it, self.layout['items']), anchor='sw', fill=ACCENT,
                               font=(FONT, 9), tags=('item', tag))
            elif step:
                for i, ch in enumerate(txt):
                    cv.create_text(self._px(x + i * step), self._py(y) + drop, text=ch, anchor=anchor, fill=color,
                                   font=(family, -px_size, style), tags=('item', tag))
            else:
                size_px = px_size
                if it.get('maxw'):                      # the same shrinking as in the PDF
                    try:
                        pdf_font = manual_pdf_font(it['font'], it['bold'], it['italic'])
                        pdf_size = it['size']
                        while pdf_size > 6.0 and pdfmetrics.stringWidth(txt, pdf_font, pdf_size) > it['maxw']:
                            pdf_size -= 0.5
                        size_px = max(3, int(round(pdf_size * z)))
                    except Exception:
                        pass
                cv.create_text(self._px(x), self._py(y) + drop, text=txt, anchor=anchor, fill=color,
                               font=(family, -size_px, style), tags=('item', tag))

    def _draw_selection(self):
        self.cv.delete('selbox')
        if self.group_ids and not self.sel:
            boxes = []
            for fid in self.group_ids:
                it = self.layout['items'].get(fid)
                if not it or it['page'] != self.page:
                    continue
                box = self.cv.bbox('it_' + fid)
                if box:
                    boxes.append(box)
                    self.cv.create_rectangle(box[0] - 2, box[1] - 1, box[2] + 2, box[3] + 1, outline=ACCENT, tags='selbox')
            if boxes:
                self.cv.create_rectangle(min(b[0] for b in boxes) - 6, min(b[1] for b in boxes) - 5,
                                         max(b[2] for b in boxes) + 6, max(b[3] for b in boxes) + 5,
                                         outline=ACCENT, width=2, dash=(5, 3), tags='selbox')
            return
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
        if abs(event.width - self._last_width) > 2 or abs(event.height - self._last_height) > 2:
            self._last_width, self._last_height = event.width, event.height
            if self._redraw_job:
                self.after_cancel(self._redraw_job)
            follow = self._auto_fit and event.width > 300 and event.height > 200
            self._redraw_job = self.after(80, self._fit if follow else self._draw_all)

    def _hit(self, event):
        x, y = self.cv.canvasx(event.x), self.cv.canvasy(event.y)
        for item in reversed(self.cv.find_overlapping(x - 4, y - 4, x + 4, y + 4)):
            for tag in self.cv.gettags(item):
                if tag.startswith('it_'):
                    return tag[3:]
        return None

    def _space_down(self, _event=None):
        if self._typing() or isinstance(self.focus_get(), ttk.Treeview):
            return None
        if not self._space:
            self._space = True
            self.cv.configure(cursor='fleur')
        return 'break'

    def _space_up(self, _event=None):
        if self._space:
            self._space = False
            self._panning = False
            try:
                self.cv.configure(cursor='')
            except tk.TclError:
                pass

    def _on_press(self, event):
        self.cv.focus_set()
        if self._space:                             # Space held: drag moves the page instead of the items
            self._panning = True
            self._drag = None
            self.cv.scan_mark(event.x, event.y)
            self.cv.configure(cursor='fleur')
            return
        fid = self._hit(event)
        if event.state & 0x5:                       # Ctrl or Shift: add / remove from the picked set
            if fid:
                self._toggle_pick(fid)
                if fid in self.group_ids:
                    self._drag = [self.cv.canvasx(event.x), self.cv.canvasy(event.y), False]
                    return
            self._drag = None
            return
        if fid and fid in self.group_ids:          # a member of the picked section: drag moves them all
            self._drag = [self.cv.canvasx(event.x), self.cv.canvasy(event.y), False]
            return
        self._select(fid)
        self._drag = [self.cv.canvasx(event.x), self.cv.canvasy(event.y), False] if fid else None

    def _on_drag(self, event):
        if self._panning:
            self.cv.scan_dragto(event.x, event.y, gain=1)
            return
        if not self._drag or not (self.sel or self.group_ids):
            return
        cx, cy = self.cv.canvasx(event.x), self.cv.canvasy(event.y)
        dx, dy = cx - self._drag[0], cy - self._drag[1]
        if not self._drag[2] and abs(dx) + abs(dy) < 3:
            return
        if not self._drag[2]:
            self._snap(None)
        self._drag[0], self._drag[1], self._drag[2] = cx, cy, True
        if self.group_ids and not self.sel:
            for fid in self.group_ids:
                member = self.layout['items'].get(fid)
                if member:
                    member['x'] += dx / self.zoom
                    member['y'] -= dy / self.zoom
                    self.cv.move('it_' + fid, dx, dy)
            self.cv.move('selbox', dx, dy)
            self._set_dirty()
            return
        it = self.layout['items'][self.sel]
        it['x'] += dx / self.zoom
        it['y'] -= dy / self.zoom
        self.cv.move('it_' + self.sel, dx, dy)
        self.cv.move('selbox', dx, dy)
        self._set_dirty()
        self._load_xy()

    def _on_release(self, _event):
        if self._panning:
            self._panning = False
            self.cv.configure(cursor='fleur' if self._space else '')
            return
        moved = bool(self._drag and self._drag[2])
        self._drag = None
        if moved and self.group_ids and not self.sel:
            for fid in self.group_ids:
                member = self.layout['items'].get(fid)
                if member:
                    member['x'], member['y'] = round(member['x'], 1), round(member['y'], 1)
            self._draw_items()
            return
        if moved and self.sel:
            it = self.layout['items'][self.sel]
            it['x'], it['y'] = round(it['x'], 1), round(it['y'], 1)
            self._load_xy()
            self._draw_items()

    def _on_hover(self, event):
        x = (self.cv.canvasx(event.x) - self.ox) / self.zoom
        y = self.sizes[self.page][1] - (self.cv.canvasy(event.y) - self.PAD) / self.zoom
        self.status_var.set(f"X {x:7.1f}   Y {y:7.1f}  pt")
        if self._space:
            self.cv.configure(cursor='fleur')
        else:
            self.cv.configure(cursor='fleur' if self._hit(event) else '')

    def _on_wheel(self, event):
        num = getattr(event, 'num', None)
        delta = 120 if num == 4 else -120 if num == 5 else getattr(event, 'delta', 0)
        if not delta:
            return 'break'
        if event.state & 0x4:                              # Ctrl + wheel = zoom towards the pointer
            self._zoom_wheel(delta, (event.x, event.y))
        elif event.state & 0x1:                            # Shift + wheel = sideways
            self.cv.xview_scroll(-3 if delta > 0 else 3, 'units')
        else:
            self.cv.yview_scroll(-3 if delta > 0 else 3, 'units')
        return 'break'

    def _typing(self):
        return isinstance(self.focus_get(), (tk.Entry, ttk.Entry, ttk.Spinbox, ttk.Combobox, tk.Text))

    def _on_arrow(self, event, direction):
        if not (self.sel or self.group_ids) or self._typing() or isinstance(self.focus_get(), ttk.Treeview):
            return None
        amount = 5.0 if event.state & 0x1 else 0.5
        self._move_by(direction[0] * amount, direction[1] * amount)
        return 'break'

    def _on_delete_key(self, _event):
        if self.sel and not self._typing():
            self._delete_selected()

    # ------------------------------------------------------------------ view
    def _zoom_wheel(self, delta, anchor):
        """Ctrl + wheel. A burst of wheel events is merged into ONE zoom step per screen refresh (smooth, no backlog)."""
        steps = max(-3.0, min(3.0, delta / 120.0))
        base = self._zoom_goal if self._zoom_goal is not None else self.zoom
        self._zoom_goal = min(max(base * (1.12 ** steps), self.MIN_ZOOM), self.MAX_ZOOM)
        self._zoom_anchor = anchor
        if self._zoom_job is None:
            self._zoom_job = self.after(15, self._apply_zoom_goal)

    def _apply_zoom_goal(self):
        self._zoom_job = None
        goal, self._zoom_goal = self._zoom_goal, None
        if goal is not None:
            self._set_zoom(goal, self._zoom_anchor)

    def _set_zoom(self, zoom, anchor=None):
        """Zoom keeping the spot under `anchor` (widget coordinates) fixed, like Photoshop.
        Without an anchor the middle of the window is used."""
        self._auto_fit = False
        zoom = round(min(max(zoom, self.MIN_ZOOM), self.MAX_ZOOM), 2)      # whole percents: pictures are reused
        if abs(zoom - self.zoom) < 1e-4:
            return
        cv = self.cv
        ax, ay = anchor if anchor else (cv.winfo_width() / 2, cv.winfo_height() / 2)
        px = (cv.canvasx(ax) - self.ox) / self.zoom        # the spot of the page under the pointer
        py = (cv.canvasy(ay) - self.PAD) / self.zoom
        self.zoom = zoom
        self._draw_all()
        pw, ph = self.sizes[self.page]
        total_w = max(self.ox * 2 + pw * zoom, max(cv.winfo_width(), 100))
        total_h = ph * zoom + 2 * self.PAD
        cv.xview_moveto(max(0.0, (self.ox + px * zoom - ax) / total_w))
        cv.yview_moveto(max(0.0, (self.PAD + py * zoom - ay) / total_h))

    def _fit(self):
        self._auto_fit = True
        pw, ph = self.sizes[self.page]
        cw, ch = max(self.cv.winfo_width(), 300), max(self.cv.winfo_height(), 200)
        self.zoom = min((cw - 2 * self.PAD) / pw, (ch - 2 * self.PAD) / ph)
        self.zoom = min(max(round(self.zoom, 2), self.MIN_ZOOM), self.MAX_ZOOM)
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
        self.group_ids, self.group_gid, self.picks = [], None, []
        self.sel = fid if fid in self.layout['items'] else None
        if self.sel:
            self._anchor = self.sel
        if self.sel and self.layout['items'][self.sel]['page'] != self.page:
            self.page = self.layout['items'][self.sel]['page']
            self.page_seg.set(self.page)
            self._draw_all()
        else:
            self._draw_selection()
        if not from_tree:
            self._sync_tree_selection()
        self._load_props()
        self._sync_hidden_selection()
        if from_tree:
            self._reveal()

    def _load_props(self):
        if not self.sel:
            self.props.pack_forget()
            if self.group_ids:
                self.empty_lbl.configure(text=T(
                    f"Επιλεγμένα στοιχεία: {len(self.group_ids)}.\n\nΣύρετε οποιοδήποτε από τα στοιχεία της "
                    "στη σελίδα και θα μετακινηθούν όλα μαζί, με την ίδια απόσταση μεταξύ τους. Ctrl+Z = αναίρεση.\n\n"
                    "Ctrl+κλικ σε στοιχείο ή ενότητα (στη σελίδα ή στη λίστα) το προσθέτει ή το αφαιρεί.\n\nΓια ένα μόνο στοιχείο, πατήστε πάνω του.",
                    f"Several items are picked ({len(self.group_ids)}).\n\nDrag any of its items on the page and "
                    "they all move together, keeping their distances. Ctrl+Z = undo.\n\n"
                    "Ctrl+click on an item or section (page or list) adds or removes it.\n\nFor a single item, click it."))
            else:
                self.empty_lbl.configure(text=getattr(self, '_empty_default', self.empty_lbl.cget('text')))
            self.empty_lbl.pack(anchor="nw", padx=18, pady=18)
            self.del_btn.state(['disabled'])
            self.dup_btn.state(['disabled'])
            self.hide_btn.state(['!disabled'] if self.group_ids else ['disabled'])
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
            self.p_title.configure(text=manual_label(self.sel, it, self.layout['items']))
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
            self.hide_btn.state(['!disabled'] if (it.get('visible', True) and not it.get('split')) else ['disabled'])
            self.dup_btn.state(['!disabled'] if (not child and not rows and not it.get('split')) else ['disabled'])
            # individual controls
            tid = it.get('parent') or self.sel
            can_split = tid in items and (tid.startswith('rows') or manual_splittable(items[tid]))
            if can_split:
                is_split = bool(items[tid].get('split'))
                self.split_btn.configure(text=T("Συγχώνευση ξανά σε ένα", "Merge back into one") if is_split
                                         else (T("Χωρισμός σε μεμονωμένα στοιχεία", "Split into individual items")
                                               if tid.startswith('rows') else
                                               T("Χωρισμός σε γράμματα", "Split into letters")))
                if tid.startswith('rows'):
                    hint = T("Κάθε χώρα γίνεται ξεχωριστό στοιχείο που μετακινείται ελεύθερα (το ύψος γραμμής ακολουθεί το template).",
                             "Each country becomes its own item that moves freely (the row height still follows the template).")
                else:
                    hint = (T("Κάθε ψηφίο γίνεται ξεχωριστό στοιχείο.", "Each digit becomes its own item.")
                            if tid in MANUAL_SPLIT_LEN else
                            T("Κάθε γράμμα γίνεται ξεχωριστό στοιχείο που μετακινείται ελεύθερα. Το τελευταίο γράφει και ό,τι "
                              "περισσεύει αν η τιμή είναι μεγαλύτερη. «Συγχώνευση» τα ξαναενώνει.",
                              "Each letter becomes its own item you can move freely. The last one also writes any surplus "
                              "if the value is longer. “Merge” joins them again."))
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
        was_visible = it.get('visible', True)
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
        self.p_title.configure(text=manual_label(self.sel, it, self.layout['items']))
        if it['visible'] != was_visible:            # moves between "Items on the PDF" and "Hidden"
            self._fill_tree()
            self.hide_btn.state(['!disabled'] if it['visible'] else ['disabled'])
        else:
            if self.tree.exists(self.sel):
                self.tree.item(self.sel, text=self._tree_text(self.sel, it))
            if hasattr(self, 'htree') and self.htree.exists('h:' + self.sel):
                self.htree.item('h:' + self.sel, text=self._hid_label(self.sel, it))
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
        if self.group_ids and not self.sel:
            self._snap('gmove:' + str(self.group_gid))
            for fid in self.group_ids:
                member = self.layout['items'].get(fid)
                if member:
                    member['x'] = round(member['x'] + dx, 2)
                    member['y'] = round(member['y'] + dy, 2)
            self._set_dirty()
            self._draw_items()
            return
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
        ref = self._mode_reference()['items'].get(self.sel) or manual_default_layout()['items'][self.sel]
        self.layout['items'][self.sel] = copy.deepcopy(ref)
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
        self._fill_tree()
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
        for cid in [c for c, i in self.layout['items'].items() if i.get('parent') == fid]:
            del self.layout['items'][cid]
        del self.layout['items'][fid]
        self._set_dirty()
        self._fill_tree()
        self._load_props()
        self._draw_items()

    def _current_mode(self):
        return 'detect' if self.context == 'detect' else self.app.prefs.get('pdf_mode', 'overlay')

    def _mode_reference(self):
        """What the selected way of filling the form would place where (computed once; used by 'Reset this item')."""
        if self._mode_ref is None:
            mode = self._current_mode()
            try:
                if mode == 'manual':
                    mode = self.app.prefs.get('manual_base', 'overlay')
                self._mode_ref = (manual_default_layout() if mode in ('manual', 'overlay', 'blank')
                                  else build_mode_layout(mode, self.template)[0])
            except Exception:
                self._mode_ref = manual_default_layout()
        return self._mode_ref

    def _replace_layout(self, layout, status):
        self._snap(None)
        self.layout = layout
        self.sel = None
        self._fill_tree()
        self._set_dirty()
        self._load_props()
        self._draw_items()
        self.status_var.set(status)

    def _wipe_all(self):
        if not messagebox.askyesno(T("Εκκαθάριση όλων", "Clear all"),
                                   T("Όλα τα στοιχεία θα αφαιρεθούν και ο καμβάς θα μείνει εντελώς κενός. Μπορείτε να προσθέσετε "
                                     "ό,τι θέλετε με το «+ Προσθήκη στοιχείου…».\n\nΣυνέχεια;",
                                     "Every item will be removed and the canvas will be completely blank. You can add whatever "
                                     "you want with “+ Add item…”.\n\nContinue?"), default='no', parent=self):
            return
        self._replace_layout(blank_layout(), T("Ο καμβάς είναι κενός.", "The canvas is blank."))

    def _auto_place(self):
        mode = self._current_mode()
        if mode == 'manual':          # "Set manually": use the auto-fill chosen in Settings
            mode = self.app.prefs.get('manual_base', 'overlay')
        names = {'blank': T("κενός καμβάς", "blank canvas"), 'overlay': T("Παλιά μέθοδος (σταθερές συντεταγμένες)", "Legacy (fixed coordinates)"),
                 'smart': T("Έξυπνη μετατόπιση", "Smart Shift"),
                 'detect': T("Έξυπνη ανίχνευση πεδίων", "Smart Field Detection"),
                 'manual': T("τελευταία αποθηκευμένη διάταξη", "last saved layout")}
        if not messagebox.askyesno(T("Αυτόματη τοποθέτηση", "Auto-place items"),
                                   T(f"Όλα τα στοιχεία θα τοποθετηθούν ξανά όπως θα τα έβαζε η επιλεγμένη μέθοδος ({names.get(mode, mode)}). "
                                     "Οι αλλαγές σας και τα δικά σας κείμενα θα χαθούν.\n\nΣυνέχεια;",
                                     f"Every item will be placed again the way the selected method ({names.get(mode, mode)}) would "
                                     "place it. Your changes and your own texts will be lost.\n\nContinue?"),
                                   default='no', parent=self):
            return
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            layout, note = manual_base_layout(mode, self.template)
        except Exception as e:
            messagebox.showerror(T("Αυτόματη τοποθέτηση", "Auto-place items"),
                                 T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"), parent=self)
            return
        finally:
            self.config(cursor='')
        self._mode_ref = copy.deepcopy(layout)
        self._replace_layout(layout, T("✓ Τα στοιχεία τοποθετήθηκαν αυτόματα.", "✓ The items were placed automatically.")
                             + ((" " + note.splitlines()[0]) if note else ""))
        if note:
            messagebox.showwarning(T("Αυτόματη τοποθέτηση", "Auto-place items"), note, parent=self)

    def _context_tip(self):
        """A short tip that fits what is selected right now."""
        if self.group_ids and not self.sel:
            return T("Σύρετε ένα από τα επιλεγμένα για να μετακινηθούν όλα. Ctrl+κλικ = προσθήκη/αφαίρεση. Esc ή κλικ σε άλλο στοιχείο = αφήνετε την επιλογή.",
                     "Drag one of the picked items to move them all. Ctrl+click = add/remove. Esc or a click on another item lets go.")
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
            return T("Ένα γράμμα/ψηφίο: μετακινείται μόνο του. «Ίδια μορφή σε όλα τα αδέλφια» περνά τη μορφή του στα υπόλοιπα.",
                     "One letter/digit: it moves on its own. “Same format for all siblings” copies its format to the others.")
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
        self.group_ids = [f for f in self.group_ids if f in self.layout['items']]
        if not self.group_ids:
            self.picks, self.group_gid = [], None
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
            text = ('0' * MANUAL_SPLIT_LEN[tid]) if tid in MANUAL_SPLIT_LEN else manual_item_text(items, tid, self.sample)
            made = manual_split_item(self.layout, tid, text, [])
            new_sel = made[0] if made else tid
            if not made:
                self.status_var.set(T("Δεν υπάρχει κείμενο για χωρισμό (η τιμή είναι κενή ή έχει μόνο έναν χαρακτήρα).",
                                      "There is no text to split (the value is empty or a single character)."))
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
        if key in MANUAL_FIELD_INDEX and key not in self.layout['items']:      # a field of the program that is not on the canvas
            ref = self._mode_reference()['items'].get(key) or manual_default_layout()['items'][key]
            it = copy.deepcopy(ref)
            it['visible'] = True
            self.layout['items'][key] = it
            self._set_dirty()
            self._fill_tree()
            self._draw_items()
            self._select(key)
            return
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
        kinds['cfield'] = ("Μεταβλητές", "Variables")
        catalog = manual_source_catalog() + [(f[0], f[2], 'rows') for f in MANUAL_FIELDS
                                             if f[1] == 'rows' and f[0] not in self.layout['items']]

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

    def _base_title(self):
        if self.temporary:
            return T("Επεξεργαστής διάταξης (προσωρινά)", "Layout editor (temporary)")
        if self.context == 'detect':
            return T("Επεξεργαστής διάταξης (Έξυπνη ανίχνευση)", "Layout editor (Smart Field Detection)")
        return T("Επεξεργαστής διάταξης", "Layout editor")

    def _context_note(self):
        if self.context == 'detect':
            return T("Οι αλλαγές θυμούνται όσο μένετε στην Έξυπνη ανίχνευση πεδίων · ",
                     "Your changes are remembered while Smart Field Detection stays selected · ")
        return ""

    def _update_title(self):
        self.title(self._base_title() + (" *" if self.dirty else ""))

    def save(self):
        if self.temporary:          # only for the PDF being previewed: hand the layout over, write nothing
            self._set_dirty(False)
            layout, callback = copy.deepcopy(self.layout), self.on_save
            try:
                self.grab_release()
            except tk.TclError:
                pass
            self.destroy()
            if callback:
                callback(layout)
            return True
        try:
            if self.context == 'detect':
                save_detect_cache(self.layout, self.template)
            else:
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
                  "There are changes that are not saved.\n\nSave them?") if not self.temporary else
                T("Υπάρχουν αλλαγές που δεν έχουν εφαρμοστεί στην προεπισκόπηση.\n\nΝα εφαρμοστούν;",
                  "There are changes that are not applied to the preview.\n\nApply them?"), parent=self)
            if answer is None:
                return
            if answer:
                if not self.save():
                    return
                if self.temporary:          # save() already closed the editor
                    return
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()

    def _on_data_mode(self, value):
        self.data_mode.set(value)
        if value == 'real' and self.real_values is None:
            self._load_real()
        self._draw_items()

    def _load_real(self):
        """Builds the 'real data' view from what the user saved: hotel details, own fields and the latest month."""
        if self._real_loading:
            return
        self._real_loading = True
        template = self.template

        def work():
            settings = load_hotel_settings()
            for cf in load_custom_fields():
                settings[CF_PREFIX + cf['id']] = cf['value']
            hist = load_history()
            latest = max(((y, m) for y, ms in hist.items() for m in ms), default=None)
            today = datetime.date.today()
            y, m = latest if latest else (today.year, today.month)
            rows = list(hist.get(y, {}).get(m, [])) if latest else []
            rec = load_month_records().get((y, m), {})
            last = calendar.monthrange(y, m)[1]
            mon = dict(year=str(y), month=str(m), date_to=f"{last:02d}/{m:02d}/{y}",
                       days_of_month=str(rec.get('days') or last),
                       employees=str(rec.get('employees') or settings.get('last_employees') or '0'),
                       sign_date=today.strftime('%d/%m/%Y'))
            path = os.path.join(tempfile.gettempdir(), "ELSTAT_editor_real.pdf")
            return process_elstat_report(None, template, path, settings, mon, rows_override=rows,
                                         save_history=False, pdf_mode='overlay')

        def done(res):
            self._real_loading = False
            self.real_values = res.get('live_values') or {}
            self.real_rows = res.get('pdf_data') or {}
            self.render_note.configure(text="")
            self._draw_items()

        def fail(e):
            self._real_loading = False
            self.data_mode.set('sample')
            self.data_seg.set('sample')
            self.render_note.configure(text="")
            self._draw_items()
            messagebox.showwarning(T("Πραγματικά δεδομένα", "Real data"),
                                   T(f"Δεν ήταν δυνατή η φόρτωση των αποθηκευμένων δεδομένων:\n{e}",
                                     f"Could not load the saved data:\n{e}"), parent=self)
        self.render_note.configure(text=T("Φόρτωση αποθηκευμένων δεδομένων…", "Loading saved data…"))
        self.app.run_async(work, done, fail, text=T("Φόρτωση δεδομένων…", "Loading data…"))

class LoadingBadge(tk.Frame):
    """Small animated 'working...' badge shown at the top right of the window."""
    FRAMES = "\u280b\u2819\u2839\u2838\u283c\u2834\u2826\u2827\u2807\u280f"

    def __init__(self, parent):
        super().__init__(parent, bg=ACCENT)
        self.spin = tk.Label(self, bg=ACCENT, fg='white', font=(FONT, 13, 'bold'))
        self.spin.pack(side="left", padx=(10, 4), pady=4)
        self.msg = tk.Label(self, bg=ACCENT, fg='white', font=(FONT, 9, 'bold'))
        self.msg.pack(side="left", padx=(0, 12))
        self._i = 0
        self._job = None
        self._shown = False

    def start(self, text):
        self.msg.configure(text=text)
        if not self._shown:
            self._shown = True
            self.place(relx=1.0, x=-18, y=74, anchor='ne')
            self._tick()
        self.lift()

    def _tick(self):
        try:
            self.spin.configure(text=self.FRAMES[self._i % len(self.FRAMES)])
            self._i += 1
            self._job = self.after(90, self._tick)
        except tk.TclError:
            pass

    def stop(self):
        if self._job:
            try:
                self.after_cancel(self._job)
            except tk.TclError:
                pass
            self._job = None
        if self._shown:
            self._shown = False
            try:
                self.place_forget()
            except tk.TclError:
                pass


class ElstatApp(TkBase):
    def report_callback_exception(self, exc, val, tb):
        _handle_error(exc, val, tb, self)

    def __init__(self):
        super().__init__()
        self.withdraw()                    # stay hidden while the screen is built, so it does not jump around
        self.geometry("920x840")
        self.minsize(640, 560)
        self.configure(bg=BG)

        self._storage_notices = init_storage()
        load_custom_country_mappings()
        self.prefs = load_prefs()
        load_shortcut_overrides(self.prefs.get('shortcuts'))
        set_language(self.prefs['language'])
        set_theme(self.prefs['theme'])
        self.configure(bg=BG)
        self.current_settings = load_hotel_settings()
        self._rebuilding = False
        self.output_custom = False      # True once the user picks the output file themselves
        self._busy = False
        self._busy_count = 0
        self._busy_text = ''
        self._badge = None
        self._load_token = 0
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
        self._shortcuts = ShortcutBinder(self)         # changeable in Settings -> Keyboard shortcuts
        self._shortcuts.add('open_excel', lambda e: self.browse_excel())
        self._shortcuts.add('help', lambda e: self.open_help())
        self._shortcuts.add('help_search', self._help_find)
        self._shortcuts.add('generate', self._try_generate)
        self._shortcuts.add('generate_print', lambda e: self._try_generate(print_after=True))
        self._shortcuts.apply()
        if self.prefs.get('center_window', True):
            self._center()
        self.deiconify()                   # now it appears once, already in its final place
        if not DND_AVAILABLE and DND_ERROR:
            self._status("Drag & drop unavailable: " + DND_ERROR)
        elif not DND_AVAILABLE:
            self._status(T("Για μεταφορά αρχείων με drag & drop εγκαταστήστε: pip install tkinterdnd2",
                           f"For drag-and-drop support run: \"{sys.executable}\" -m pip install tkinterdnd2"))
        self.after(100, self._apply_titlebar)
        self.after(150, self.excel_btn.focus_set)
        self._boot_time = time.time()
        self.after(2500, self._startup_forced_check)          # silent: only acts if GitHub says FORCE_UPDATE = True
        if load_app_config().get('update_check_on_start'):
            self.after(3000, lambda: self.check_for_updates(silent=True))
        if self._storage_notices:
            self.after(400, lambda: messagebox.showinfo(T("Ρυθμίσεις", "Settings"),
                                                        "\n\n".join(T(*n) for n in self._storage_notices),
                                                        parent=self))

    # ------------------------------------------------------------------ setup
    def _init_vars(self):
        cs = self.current_settings
        self.set_vars = {k: tk.StringVar(value=cs.get(k, '')) for k in SETTING_KEYS}
        self.custom_fields = load_custom_fields()
        self.cf_vars = {cf['id']: tk.StringVar(value=cf['value']) for cf in self.custom_fields}
        if not hasattr(self, '_extra_open'):
            self._extra_open = any(cs.get(k) for k in EXTRA_KEYS) or any(cf['value'] for cf in self.custom_fields)
        self._cf_save_job = None
        for _v in self.cf_vars.values():
            _v.trace_add('write', self._cf_value_changed)
        # month details start EMPTY: the Excel fills most of them, the rest are typed by hand
        self.mon_vars = {k: tk.StringVar() for k in ('year', 'month', 'date_to', 'days_of_month', 'employees', 'sign_date')}
        self.mon_vars['sign_date'].set(datetime.date.today().strftime('%d/%m/%Y'))   # today unless changed
        self.excel_var = tk.StringVar()
        self.out_var = tk.StringVar(value=self._suggest_output())
        saved_tpl = self.prefs.get('template_pdf', '')
        self.pdf_var = tk.StringVar(value=saved_tpl if saved_tpl and os.path.isfile(saved_tpl) else DEFAULT_TEMPLATE)
        self.period_var = tk.StringVar()
        self._period_msg = (lambda: T("Πατήστε «Επιλογή Excel / CSV…» και διαλέξτε το αρχείο «Nationalities Statistics» "
                                      "που βγάζει το Protel (.xls, .xlsx ή .csv).",
                                      "Click “Choose Excel / CSV…” and select the “Nationalities Statistics” file "
                                      "produced by Protel (.xls, .xlsx or .csv)."), 'CardHint.TLabel')
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
                   for k in ('language', 'theme', 'output_mode', 'output_dir', 'filename_pattern', 'excel_dir', 'pdf_mode', 'manual_base')}
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
        sub = tk.Label(texts, text=T("Διαλέξτε το Excel ή CSV από το Protel και το πρόγραμμα συμπληρώνει και "
                                     "εκτυπώνει το δελτίο της ΕΛΣΤΑΤ.",
                                     "Choose the Excel or CSV from Protel and the program fills in and "
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
        self.excel_btn = ttk.Button(c1, text=T("Επιλογή Excel / CSV…", "Choose Excel / CSV…"), command=self.browse_excel)
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
        ttk.Label(bottom, text=T("Συντομεύσεις:  Ctrl+O = επιλογή Excel / CSV  ·  Ctrl+Enter = δημιουργία  ·  "
                                 "Ctrl+P = δημιουργία & εκτύπωση  ·  F1 = οδηγός",
                                 "Shortcuts:  Ctrl+O = choose Excel / CSV  ·  Ctrl+Enter = create  ·  "
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

        self.hotel_sf = sf
        self._build_extra_section(f, r, body)
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

    # ------------------------------------------------------- optional extra details (drop-down)
    def _build_extra_section(self, parent, row, scroll_body):
        """A collapsible block under the hotel details: the optional address fields plus the user's own fields.
        It lives inside the same card, follows the window width (1 or 2 columns) and scrolls with the page."""
        outer = ttk.Frame(parent, style='Card.TFrame')
        outer.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(18, 0))
        outer.columnconfigure(0, weight=1)
        self.extra_outer = outer
        ttk.Separator(outer, orient="horizontal").grid(row=0, column=0, sticky="ew")

        head = tk.Frame(outer, bg=CARD, cursor='hand2')
        head.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        arrow = tk.Label(head, bg=CARD, fg=ACCENT, font=(FONT, 11, 'bold'), width=2, cursor='hand2')
        arrow.pack(side="left")
        title = tk.Label(head, text=T("Πρόσθετα προαιρετικά στοιχεία", "Optional extra details"), bg=CARD, fg=NAVY,
                         font=(FONT, 11, 'bold'), cursor='hand2')
        title.pack(side="left")
        badge = tk.Label(head, bg=ACCENT_SOFT, fg=ACCENT, font=(FONT, 9, 'bold'), padx=8, pady=1, cursor='hand2')
        badge.pack(side="left", padx=10)
        sub = tk.Label(head, text=T("Νομός, οδός, Τ.Κ. και δικά σας πεδία - δεν απαιτούνται",
                                    "Prefecture, street, postal code and your own fields - none required"),
                       bg=CARD, fg=MUTED, font=(FONT, 9), cursor='hand2', anchor='w')
        sub.pack(side="left", fill="x", expand=True)

        content = ttk.Frame(outer, style='Card.TFrame')
        content.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        content.columnconfigure(0, weight=1)
        self.extra_content = content

        grid = ttk.Frame(content, style='Card.TFrame')
        grid.grid(row=0, column=0, sticky="ew")
        self.extra_grid = grid

        labels = {'nomos': T("Νομός", "Prefecture"), 'dimotiki': T("Δημοτική ενότητα", "Municipal unit"),
                  'odos': T("Οδός", "Street"), 'arithmos': T("Αριθμός", "Number"), 'tk': T("Τ.Κ.", "Postal code")}
        self.extra_cells = []
        for key in EXTRA_KEYS:
            self.extra_cells.append(self._extra_cell(grid, labels[key], self.set_vars[key], key))

        ttk.Label(content, text=T("Δικά σας πεδία", "Your own fields"), style='CardBold.TLabel'
                  ).grid(row=1, column=0, sticky="w", pady=(14, 0))
        self.cf_hint = ttk.Label(content, style='CardHint.TLabel', justify="left")
        self.cf_hint.grid(row=2, column=0, sticky="w", pady=(2, 0))
        autowrap(self.cf_hint, scroll_body, margin=150)
        self.cf_grid = ttk.Frame(content, style='Card.TFrame')
        self.cf_grid.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        self.cf_cells = {}
        for cf in self.custom_fields:
            self._cf_make_cell(cf)
        self.cf_add_btn = ttk.Button(content, text=T("＋ Προσθήκη πεδίου", "＋ Add field"), command=self._cf_add)
        self.cf_add_btn.grid(row=4, column=0, sticky="w", pady=(8, 0))

        def toggle(_e=None):
            self._extra_open = not self._extra_open
            self._update_extra_visibility(reveal=self._extra_open)
        for w in (head, arrow, title, badge, sub):
            w.bind('<Button-1>', toggle)
        self._extra_widgets = (arrow, badge)
        grid.bind('<Configure>', lambda e: self._reflow_extra())
        self.cf_grid.bind('<Configure>', lambda e: self._reflow_extra())
        for v in list(self.set_vars[k] for k in EXTRA_KEYS) + list(self.cf_vars.values()):
            v.trace_add('write', lambda *_: self._update_extra_badge())
        self._update_extra_visibility()

    def _extra_cell(self, grid, label, var, key=None):
        """One label + entry pair; positioned later by _reflow_extra."""
        cell = ttk.Frame(grid, style='Card.TFrame')
        cell.columnconfigure(0, weight=1)
        ttk.Label(cell, text=label, style='CardHint.TLabel').grid(row=0, column=0, sticky="w")
        e = ttk.Entry(cell, textvariable=var)
        e.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        if key:
            self.entries[key] = e
        return cell

    def _reflow_extra(self):
        """1 column on a narrow window, 2 columns on a wide one - never wider than the card."""
        for grid, cells in ((self.extra_grid, self.extra_cells), (self.cf_grid, list(self.cf_cells.values()))):
            try:
                width = grid.winfo_width()
            except tk.TclError:
                continue
            cols = 2 if width >= 520 else 1
            if getattr(grid, '_cols', None) == cols and getattr(grid, '_n', None) == len(cells) \
                    and getattr(grid, '_ids', None) == [id(c) for c in cells]:
                continue
            grid._cols, grid._n, grid._ids = cols, len(cells), [id(c) for c in cells]
            for c in range(2):
                grid.columnconfigure(c, weight=1 if c < cols else 0, uniform='x' if c < cols else '')
            for i, cell in enumerate(cells):
                cell.grid(row=i // cols, column=i % cols, sticky="ew",
                          padx=(0, 12) if cols == 2 and i % 2 == 0 else (0, 0), pady=(4, 4))

    def _update_extra_badge(self):
        if self._rebuilding or not hasattr(self, '_extra_widgets'):
            return
        n = sum(1 for k in EXTRA_KEYS if self.set_vars[k].get().strip())
        n += sum(1 for v in self.cf_vars.values() if v.get().strip())
        try:
            self._extra_widgets[1].configure(text=(T(f"{n} συμπληρωμένα", f"{n} filled") if n else T("κενό", "empty")))
        except tk.TclError:
            pass

    def _update_extra_visibility(self, reveal=False):
        """Applies the Settings switch (whole block) and the open/closed state of the drop-down."""
        outer = getattr(self, 'extra_outer', None)
        if outer is None or self._rebuilding:
            return
        try:
            if not self.prefs.get('show_extra_fields', True):
                outer.grid_remove()
                return
            outer.grid()
            self._extra_widgets[0].configure(text="\u25be" if self._extra_open else "\u25b8")
            if self._extra_open:
                self.extra_content.grid()
            else:
                self.extra_content.grid_remove()
            self.cf_hint.configure(text=(T("Προσθέστε ό,τι άλλο θέλετε. Κάθε πεδίο γίνεται «μεταβλητή» που τοποθετείται "
                                           "όπου θέλετε στο δελτίο από τον επεξεργαστή διάταξης PDF (Ρυθμίσεις ή προσωρινά "
                                           "από την προεπισκόπηση).",
                                           "Add anything else you need. Each field becomes a “variable” you can place "
                                           "anywhere on the form in the PDF layout editor (Settings, or temporarily "
                                           "from the preview).")
                                         if self.custom_fields else
                                         T("Δεν έχετε δικά σας πεδία. Πατήστε «Προσθήκη πεδίου» - μετά μπορείτε να τα βάλετε "
                                           "οπουδήποτε στο PDF από τον επεξεργαστή διάταξης.",
                                           "You have no fields of your own yet. Press “Add field” - you can then put "
                                           "them anywhere on the PDF in the layout editor.")))
            self._update_extra_badge()
            self.update_idletasks()
            self._reflow_extra()
            if reveal:
                self.after(30, self._reveal_extra)
        except tk.TclError:
            pass

    def _reveal_extra(self):
        """After opening, scroll just enough so the whole block is visible (nothing is hidden below the fold)."""
        sf = getattr(self, 'hotel_sf', None)
        try:
            self.update_idletasks()
            cv = sf.canvas
            need = (self.extra_outer.winfo_rooty() + self.extra_outer.winfo_height()) - (cv.winfo_rooty() + cv.winfo_height())
            total = max(1, sf.body.winfo_reqheight())
            if need > 0 and sf._needs_scroll():
                cv.yview_moveto(min(1.0, (cv.canvasy(0) + need + 14) / total))
        except (tk.TclError, AttributeError):
            pass

    # ---- the user's own fields
    def _cf_make_cell(self, cf):
        cell = ttk.Frame(self.cf_grid, style='Card.TFrame')
        cell.columnconfigure(0, weight=1)
        top = ttk.Frame(cell, style='Card.TFrame')
        top.grid(row=0, column=0, columnspan=2, sticky="ew")
        top.columnconfigure(0, weight=1)
        lbl = ttk.Label(top, text=cf['name'], style='CardHint.TLabel')
        lbl.grid(row=0, column=0, sticky="w")
        ren = tk.Label(top, text="\u270e", bg=CARD, fg=MUTED, font=(FONT, 10), cursor='hand2', padx=4)
        ren.grid(row=0, column=1)
        ren.bind('<Button-1>', lambda e, i=cf['id']: self._cf_rename(i))
        rem = tk.Label(top, text="\u2715", bg=CARD, fg=MUTED, font=(FONT, 10), cursor='hand2', padx=4)
        rem.grid(row=0, column=2)
        rem.bind('<Button-1>', lambda e, i=cf['id']: self._cf_remove(i))
        rem.bind('<Enter>', lambda e, w=rem: w.configure(fg=BAD))
        rem.bind('<Leave>', lambda e, w=rem: w.configure(fg=MUTED))
        ren.bind('<Enter>', lambda e, w=ren: w.configure(fg=ACCENT))
        ren.bind('<Leave>', lambda e, w=ren: w.configure(fg=MUTED))
        ttk.Entry(cell, textvariable=self.cf_vars[cf['id']]).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 0))
        self.cf_cells[cf['id']] = cell
        cf['_label'] = lbl

    def _cf_persist(self):
        save_custom_fields([{'id': c['id'], 'name': c['name'], 'value': self.cf_vars[c['id']].get().strip()}
                            for c in self.custom_fields])

    def _save_custom_fields_now(self):
        try:
            self._cf_persist()
        except OSError as e:
            self._status(T(f"⚠ Δεν αποθηκεύτηκαν τα δικά σας πεδία: {e}", f"⚠ Could not save your own fields: {e}"))

    def _cf_value_changed(self, *_):
        if self._rebuilding:
            return
        if self._cf_save_job:
            self.after_cancel(self._cf_save_job)
        self._cf_save_job = self.after(700, self._save_custom_fields_now)

    def _cf_ask_name(self, title, initial=""):
        while True:
            name = simpledialog.askstring(title, T("Όνομα πεδίου (π.χ. Όροφος, Fax, Κωδικός ΕΟΤ):",
                                                   "Field name (e.g. Floor, Fax, Licence no.):"),
                                          initialvalue=initial, parent=self)
            if name is None:
                return None
            name = " ".join(name.split())
            if not name:
                return None
            clash = [c for c in self.custom_fields if c['name'].casefold() == name.casefold() and c['name'] != initial]
            if clash:
                messagebox.showwarning(title, T("Υπάρχει ήδη πεδίο με αυτό το όνομα.", "A field with this name already exists."),
                                       parent=self)
                initial = name
                continue
            return name[:60]

    def _cf_add(self):
        name = self._cf_ask_name(T("Νέο πεδίο", "New field"))
        if not name:
            return
        n = 1 + max([int(c['id'][2:]) for c in self.custom_fields] or [0])
        cf = {'id': f'cf{n}', 'name': name, 'value': ''}
        self.custom_fields.append(cf)
        self.cf_vars[cf['id']] = tk.StringVar()
        self.cf_vars[cf['id']].trace_add('write', self._cf_value_changed)
        self.cf_vars[cf['id']].trace_add('write', lambda *_: self._update_extra_badge())
        self._cf_make_cell(cf)
        self._save_custom_fields_now()
        self._update_extra_visibility(reveal=True)
        self._status(T(f"Προστέθηκε το πεδίο «{name}». Βρίσκεται στον επεξεργαστή διάταξης ως μεταβλητή.",
                       f"Field “{name}” added. It is available in the layout editor as a variable."))

    def _cf_rename(self, fid):
        cf = next((c for c in self.custom_fields if c['id'] == fid), None)
        if not cf:
            return
        name = self._cf_ask_name(T("Μετονομασία πεδίου", "Rename field"), cf['name'])
        if name:
            cf['name'] = name
            cf['_label'].configure(text=name)
            self._save_custom_fields_now()

    def _cf_remove(self, fid):
        cf = next((c for c in self.custom_fields if c['id'] == fid), None)
        if not cf:
            return
        if not messagebox.askyesno(T("Διαγραφή πεδίου", "Delete field"),
                                   T(f"Να διαγραφεί το πεδίο «{cf['name']}»;\nΌπου έχει τοποθετηθεί στη διάταξη του PDF "
                                     "δεν θα εμφανίζεται πια τίποτα.",
                                     f"Delete the field “{cf['name']}”?\nWherever it was placed in the PDF layout, "
                                     "nothing will be shown any more."), default='no', parent=self):
            return
        self.custom_fields.remove(cf)
        self.cf_vars.pop(fid, None)
        self.cf_cells.pop(fid).destroy()
        self._save_custom_fields_now()
        self._update_extra_visibility()

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
                                           (T("Στον φάκελο του αρχείου Excel / CSV (από το Protel)",
                                              "In the folder of the Excel / CSV file (from Protel)"), 'excel'),
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
        c3 = Card(body, None, T("Αρχεία Excel / CSV από το Protel", "Excel / CSV files from Protel"))
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
        hint(f3, T("Από ποιον φάκελο ξεκινούν τα «Επιλογή Excel / CSV…» και «Πιο πρόσφατο». "
                   "Κενό = ο φάκελος του τελευταίου αρχείου που χρησιμοποιήσατε.",
                   "The folder where “Choose Excel / CSV…” and “Latest” start. "
                   "Empty = the folder of the last file you used."), 1, columnspan=3)

        # ---- 3b. how the PDF is filled in: 1) who sets the positions  2) which method / starting point
        cp = Card(body, None, T("Τρόπος συμπλήρωσης του PDF", "How the PDF is filled in"))
        cp.grid(row=4, column=0, sticky="ew", padx=4, pady=6)
        polish_settings_card(cp)
        fp = cp.body
        fp.columnconfigure(0, weight=1)
        self._auto_labels = {'overlay': T("Παλιά μέθοδος - σταθερές θέσεις", "Legacy - fixed positions"),
                             'smart': T("Έξυπνη μετατόπιση", "Smart Shift"),
                             'detect': T("Έξυπνη ανίχνευση πεδίων", "Smart Field Detection")}
        self._auto_desc = {
            'overlay': T("Γράφει τις τιμές σε σταθερές θέσεις πάνω στο άδειο PDF. Έτσι δούλευε πάντα το πρόγραμμα.",
                         "Writes the values at fixed positions on the blank PDF. This is how the program has always worked."),
            'smart': T("Μετρά πόσο μετακινήθηκε το έντυπο και μετατοπίζει ανάλογα τις σταθερές θέσεις. "
                       "Χρήσιμο αν η ΕΛΣΤΑΤ αλλάξει λίγο το έντυπο.",
                       "Measures how far the form moved and shifts the fixed positions to match. "
                       "Useful if ELSTAT changes the form slightly."),
            'detect': T("Σαρώνει το άδειο PDF, διαβάζει τις ετικέτες και βρίσκει μόνο του κάθε θέση.",
                        "Scans the blank PDF, reads its labels and works out every position by itself.")}
        self._base_labels = {b: manual_base_name(b) for b in MANUAL_BASES}
        self._base_desc = {
            'overlay': T("Οι σταθερές θέσεις του προγράμματος.", "The program's fixed positions."),
            'smart': T("Οι θέσεις όπως θα τις μετατόπιζε η Έξυπνη μετατόπιση.", "Positions as Smart Shift would shift them."),
            'detect': T("Οι θέσεις όπως θα τις έβρισκε η Έξυπνη ανίχνευση πεδίων.", "Positions as Smart Field Detection would find them."),
            'blank': T("Τίποτα τοποθετημένο - χτίζετε τη διάταξη από το μηδέν.", "Nothing placed - you build the layout from scratch.")}

        ttk.Label(fp, text=T("Ποιος ορίζει τις θέσεις των τιμών;", "Who sets where the values go?"),
                  style='Card.TLabel').grid(row=0, column=0, sticky="w")
        self.kind_seg = Segmented(fp, (('auto', T("Αυτόματα", "Automatic")), ('manual', T("Χειροκίνητα", "Manual"))),
                                  command=self._kind_changed,
                                  value='manual' if self.prefs.get('pdf_mode') == 'manual' else 'auto')
        self.kind_seg.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.kind_desc_var = tk.StringVar()
        kd = ttk.Label(fp, textvariable=self.kind_desc_var, style='CardHint.TLabel', justify="left")
        kd.grid(row=2, column=0, sticky="w", pady=(6, 0))
        autowrap(kd, body, margin=110)

        # ---- automatic: one drop-down instead of three radio buttons with long texts
        self.auto_frame = af = ttk.Frame(fp, style='Card.TFrame')
        af.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        af.columnconfigure(0, weight=1)
        ttk.Label(af, text=T("Μέθοδος", "Method"), style='Card.TLabel').grid(row=0, column=0, sticky="w")
        self.auto_combo = ttk.Combobox(af, state='readonly', width=44, values=list(self._auto_labels.values()))
        self.auto_combo.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.auto_combo.bind('<<ComboboxSelected>>', self._auto_method_changed)
        self.method_desc_var = tk.StringVar()
        md = ttk.Label(af, textvariable=self.method_desc_var, style='CardHint.TLabel', justify="left")
        md.grid(row=2, column=0, sticky="w", pady=(6, 0))
        autowrap(md, body, margin=110)

        # Smart Shift options (only while it is selected)
        self.smart_frame = sm = ttk.Frame(af, style='Card.TFrame')
        sm.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        sm.columnconfigure(0, weight=1)
        self.layout_note = ttk.Label(sm, textvariable=self.layout_note_var, style='CardHint.TLabel', justify="left")
        self.layout_note.grid(row=0, column=0, sticky="w")
        autowrap(self.layout_note, body, margin=110)
        lrow = ttk.Frame(sm, style='Card.TFrame')
        lrow.grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Button(lrow, text=T("Έλεγχος template", "Check template"), command=self.layout_check).pack(side="left")
        ttk.Button(lrow, text=T("Αποθήκευση ως αναφορά…", "Save as reference…"),
                   command=self.layout_save_reference).pack(side="left", padx=8)
        ttk.Button(lrow, text=T("Δοκιμαστικό δελτίο", "Test form"), command=self.layout_test_form).pack(side="left")
        ttk.Button(lrow, text=T("Επαναφορά αναφοράς", "Reset reference"), command=self.layout_reset_reference).pack(side="left", padx=(8, 0))

        # Smart Field Detection options (only while it is selected)
        self.detect_frame = df = ttk.Frame(af, style='Card.TFrame')
        df.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        df.columnconfigure(0, weight=1)
        self.detect_note_var = tk.StringVar()
        self.detect_note = ttk.Label(df, textvariable=self.detect_note_var, style='CardHint.TLabel', justify="left")
        self.detect_note.grid(row=0, column=0, sticky="w")
        autowrap(self.detect_note, body, margin=110)
        ttk.Button(df, text=T("Ρύθμιση & έλεγχος ανίχνευσης…", "Configure & inspect detection…"),
                   style='PrintSm.TButton', command=self.open_detect_inspector).grid(row=1, column=0, sticky="w", pady=(8, 0))

        # editor pre-fill for the automatic methods ("Edit layout…" in the preview)
        self.prefill_frame = pf = ttk.Frame(af, style='Card.TFrame')
        pf.grid(row=5, column=0, sticky="ew", pady=(14, 0))
        ttk.Checkbutton(pf, text=T("Ο επεξεργαστής διάταξης ανοίγει προσυμπληρωμένος από αυτή τη μέθοδο",
                                   "The layout editor opens pre-filled by this method"),
                        variable=pv['auto_populate_layout'], style='Card.TCheckbutton',
                        command=self._pref_changed).grid(row=0, column=0, sticky="w")

        # ---- manual: choose the starting point, then open the editor
        self.manual_note_var = tk.StringVar()
        self.manual_frame = mf = ttk.Frame(fp, style='Card.TFrame')
        mf.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        mf.columnconfigure(0, weight=1)
        ttk.Label(mf, text=T("Ξεκίνα από", "Start from"), style='Card.TLabel').grid(row=0, column=0, sticky="w")
        self.base_combo = ttk.Combobox(mf, state='readonly', width=44, values=list(self._base_labels.values()))
        self.base_combo.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.base_combo.bind('<<ComboboxSelected>>', self._base_changed)
        self.base_desc_var = tk.StringVar()
        bd = ttk.Label(mf, textvariable=self.base_desc_var, style='CardHint.TLabel', justify="left")
        bd.grid(row=2, column=0, sticky="w", pady=(6, 0))
        autowrap(bd, body, margin=110)
        self.manual_note = ttk.Label(mf, textvariable=self.manual_note_var, style='CardHint.TLabel', justify="left")
        self.manual_note.grid(row=3, column=0, sticky="w", pady=(12, 0))
        autowrap(self.manual_note, body, margin=110)
        mrow = ttk.Frame(mf, style='Card.TFrame')
        mrow.grid(row=4, column=0, sticky="w", pady=(10, 0))
        ttk.Button(mrow, text=T("Άνοιγμα επεξεργαστή διάταξης", "Open layout editor"), style='PrintSm.TButton',
                   command=self.open_layout_editor).pack(side="left")
        ttk.Button(mrow, text=T("Δοκιμαστικό δελτίο", "Test form"), command=self.layout_test_form).pack(side="left", padx=8)
        ttk.Button(mrow, text=T("Επαναφορά θέσεων", "Reset positions"), command=self.manual_reset).pack(side="left")
        self._update_smart_panel()
        c2.make_collapsible()
        c3.make_collapsible()

        # ---- 4. behaviour
        c4 = Card(body, None, T("Συμπεριφορά του προγράμματος", "Program behaviour"))
        c4.grid(row=5, column=0, sticky="ew", padx=4, pady=6)
        polish_settings_card(c4)
        f4 = c4.body
        f4.columnconfigure(0, weight=1)
        behaviours = [
            ('center_window', T("Κεντράρισμα του παραθύρου κατά την εκκίνηση", "Centre the window when the program starts"),
             T("Το παράθυρο ανοίγει στο κέντρο της οθόνης. Αν το κλείσετε, το σύστημα επιλέγει τη θέση.",
               "The window opens in the middle of the screen. If you turn it off, the system chooses the position.")),
            ('open_after', T("Άνοιγμα του PDF αμέσως μετά τη δημιουργία",
                             "Open the PDF right after it is created"),
             T("Μόλις φτιαχτεί το δελτίο, ανοίγει στο πρόγραμμα PDF του υπολογιστή σας.",
               "As soon as the form is made it opens in your computer's PDF viewer.")),
            ('confirm_overwrite', T("Ερώτηση πριν αντικατασταθεί PDF που υπάρχει ήδη",
                                    "Ask before replacing a PDF that already exists"),
             T("Αν υπάρχει ήδη αρχείο με το ίδιο όνομα, ρωτάει πριν το αντικαταστήσει.",
               "If a file with the same name already exists, you are asked before it is replaced.")),
            ('keep_archive', T("Αρχειοθέτηση αντιγράφου κάθε δελτίου στον φάκελο δεδομένων",
                               "Archive a copy of every form in the data folder"),
             T("Επιτρέπει να ξανανοίξετε ή να εκτυπώσετε παλιούς μήνες από το Ιστορικό. Αν το κλείσετε, το "
               "δελτίο ξαναφτιάχνεται από τα δεδομένα του ιστορικού όταν το ζητήσετε.",
               "Lets you reopen or print old months from History. If you turn it off, the "
               "form is rebuilt from the history data when you ask for it.")),
            ('show_reminder', T("Υπενθύμιση όταν λείπει το δελτίο του προηγούμενου μήνα",
                                "Remind me when last month's form is missing"),
             T("Εμφανίζει υπενθύμιση στην καρτέλα του δελτίου όταν λείπει ο προηγούμενος μήνας.",
               "Shows a reminder on the form tab when last month's form is missing.")),
            ('prefill_employees', T("Προσυμπλήρωση του «Σύνολο εργαζομένων» με την τελευταία τιμή",
                                    "Pre-fill “Total employees” with the last value"),
             T("Ελέγχετε πάντα τον αριθμό πριν δημιουργήσετε το δελτίο.",
               "Always check the number before creating the form.")),
            ('show_extra_fields', T("Πρόσθετα προαιρετικά στοιχεία στην καρτέλα «Στοιχεία ξενοδοχείου»",
                                    "Optional extra details in the “Hotel details” tab"),
             T("Εμφανίζει ένα αναδιπλούμενο μενού με Νομό, Δημοτική ενότητα, Οδό, Αριθμό, Τ.Κ. και δικά σας πεδία. "
               "Αν το κλείσετε, τίποτα από αυτά δεν γράφεται στο δελτίο (οι τιμές διατηρούνται).",
               "Shows a drop-down with Prefecture, Municipal unit, Street, Number, Postal code and your own fields. "
               "If you turn it off, none of them is written on the form (the values are kept).")),
            ('show_loading', T("Εμφάνιση ενδείξεων φόρτωσης όσο το πρόγραμμα εργάζεται",
                               "Show loading indicators while the program is working"),
             T("Μικρή κινούμενη ένδειξη πάνω δεξιά όταν διαβάζεται το Excel, δημιουργείται το PDF κ.λπ. "
               "Αν την κλείσετε, το πρόγραμμα εργάζεται κανονικά χωρίς ένδειξη.",
               "A small animated badge at the top right while the Excel is read, the PDF is created, etc. "
               "If you turn it off, the program works the same without the badge.")),
            ('preview_before_result', T("Προεπισκόπηση PDF πριν την αποθήκευση και το παράθυρο «Αποτέλεσμα»",
                                        "Preview the PDF before it is saved and the “Result” window"),
             T("Εμφανίζει το δελτίο μέσα στο πρόγραμμα με ζουμ και δυνατότητα προσωρινής αλλαγής της διάταξης. "
               "Το αρχείο αποθηκεύεται (και ρωτάει αν υπάρχει ήδη) μόνο όταν πατήσετε «Αποθήκευση & συνέχεια».",
               "Shows the form inside the application with zoom and a temporary layout editor. The file is "
               "saved (and the “already exists” question asked) only when you press “Save & continue”.")),
        ]
        for r, (key, text, note) in enumerate(behaviours):
            rowf = ttk.Frame(f4, style='Card.TFrame')
            rowf.grid(row=r * 2, column=0, sticky="w", pady=(8 if r else 0, 0))
            ttk.Checkbutton(rowf, text=text, variable=pv[key], style='Card.TCheckbutton',
                            command=self._pref_changed).pack(side="left")
            if note:        # hover the (i) for a moment = the explanation; click it = the guide
                info = tk.Label(rowf, text="\u24d8", bg=CARD, fg=ACCENT, font=(FONT, 12), cursor='hand2')
                info.pack(side="left", padx=(8, 0))
                HoverTip(info, note, T("Πατήστε για περισσότερα στον οδηγό", "Click for more in the guide"))
                info.bind('<Button-1>', lambda e, k=key: self.open_help('settings', 'opt_' + k))

        # ---- 5. keyboard shortcuts
        self._build_shortcuts_card(body, 6, polish_settings_card, hint)

        # ---- 6. online update
        self._build_update_card(body, 7, polish_settings_card, hint)

        # ---- reset
        bottom = ttk.Frame(body)
        bottom.grid(row=8, column=0, sticky="ew", padx=4, pady=(6, 16))
        ttk.Button(bottom, text=T("Επαναφορά ρυθμίσεων", "Reset settings"), style='Danger.TButton',
                   command=self.reset_prefs).pack(side="left")
        ttk.Label(bottom, text=T("Δεν διαγράφει ιστορικό ή στοιχεία ξενοδοχείου.",
                                 "Does not delete history or hotel details."),
                  style='Hint.TLabel').pack(side="left", padx=10)

    # ---- the "How the PDF is filled in" card: Automatic / Manual switch and its two drop-downs
    def _sync_pdf_widgets(self):
        """Makes the switch, the drop-downs and their one-line explanations show the saved choice."""
        mode = self.prefs.get('pdf_mode', 'overlay')
        if mode != 'manual':
            self._last_auto = mode
        auto = getattr(self, '_last_auto', 'overlay')
        base = self.prefs.get('manual_base', 'overlay')
        self.kind_seg.set('manual' if mode == 'manual' else 'auto')
        self.auto_combo.set(self._auto_labels.get(auto, ''))
        self.base_combo.set(self._base_labels.get(base, ''))
        desc = self._auto_desc.get(auto, '')
        if auto in ('smart', 'detect'):
            desc += " " + T("Μπορεί περιστασιακά να κάνει λάθος - ρίξτε μια ματιά στο δελτίο πριν το υποβάλετε.",
                            "It can occasionally get things wrong - give the form a quick look before submitting.")
        self.method_desc_var.set(desc)
        self.base_desc_var.set(self._base_desc.get(base, ''))
        self.kind_desc_var.set(
            T("Το πρόγραμμα βρίσκει μόνο του πού πηγαίνει κάθε τιμή στο έντυπο.",
              "The program works out by itself where each value goes on the form.")
            if mode != 'manual' else
            T("Εσείς τοποθετείτε κάθε τιμή στον επεξεργαστή διάταξης (θέση, γραμματοσειρά, μέγεθος, χρώμα).",
              "You place every value yourself in the layout editor (position, font, size, colour)."))

    def _kind_changed(self, kind):
        self.pv['pdf_mode'].set('manual' if kind == 'manual' else getattr(self, '_last_auto', 'overlay'))
        self._pdf_mode_clicked()

    def _auto_method_changed(self, _event=None):
        chosen = self.auto_combo.get()
        for value, label in self._auto_labels.items():
            if label == chosen:
                self.pv['pdf_mode'].set(value)
                break
        self._pdf_mode_clicked()

    def _base_changed(self, _event=None):
        chosen = self.base_combo.get()
        for value, label in self._base_labels.items():
            if label == chosen:
                self.pv['manual_base'].set(value)
                break
        self._pref_changed()

    # ---- manual layout (settings)
    def _update_manual_panel(self):
        frame = getattr(self, 'manual_frame', None)
        if frame is None:
            return
        try:
            is_manual = self.prefs.get('pdf_mode') == 'manual'
            self._sync_pdf_widgets()
            if is_manual:
                self.auto_frame.grid_remove()
                frame.grid()
            else:
                frame.grid_remove()
                self.auto_frame.grid()
            if manual_layout_saved():
                note = T("✓ Υπάρχει αποθηκευμένη δική σας διάταξη. Πατήστε «Άνοιγμα επεξεργαστή διάταξης» για αλλαγές.",
                         "✓ Your own layout is saved. Press “Open layout editor” to change it.")
            else:
                base = manual_base_name(self.prefs.get('manual_base', 'overlay'))
                note = T(f"⚠ Δεν έχετε αποθηκεύσει ακόμη δική σας διάταξη - χρησιμοποιείται: {base}. "
                         "Ανοίξτε τον επεξεργαστή για να την προσαρμόσετε.",
                         f"⚠ You have not saved your own layout yet - using: {base}. "
                         "Open the editor to adjust it.")
            self.manual_note_var.set(note)
            self.manual_note.configure(style='CardHint.TLabel' if manual_layout_saved() else 'CardWarn.TLabel')
        except tk.TclError:
            pass

    def _editor_start_layout(self, mode, template, result=None):
        """Where the layout editor starts: with "Auto-populate manual layout template" on, from what the selected
        method would place (Legacy / Smart Shift / Smart Field Detection); off, from a blank canvas.
        A manual layout you saved yourself is always what the manual method starts from. -> (layout, note or None)"""
        if mode == 'manual':
            if manual_layout_saved():
                return load_manual_layout(), None
            self.config(cursor='watch')
            self.update_idletasks()
            try:
                return manual_base_layout(self.prefs.get('manual_base', 'overlay'), template)
            finally:
                self.config(cursor='')
        if not self.prefs.get('auto_populate_layout', True):
            return blank_layout(), None
        if mode == 'detect' and result is not None and result.get('layout_used') is not None:
            return copy.deepcopy(result['layout_used']), None      # exactly what this report was made with
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            return build_mode_layout(mode, template)
        finally:
            self.config(cursor='')

    def open_layout_editor(self):
        title = T("Επεξεργαστής διάταξης", "Layout editor")
        mode = self.prefs.get('pdf_mode', 'overlay')
        if mode == 'detect':            # its editor is reached through the inspection dialog (it saves the fine-tuning)
            self.open_detect_inspector()
            return
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
        layout, note = self._editor_start_layout(mode, template)
        self._layout_editor = LayoutEditor(self, template, on_save=self._update_manual_panel, layout=layout)
        if note:
            messagebox.showwarning(title, note)

    def open_detect_inspector(self):
        """Context A: Configure & Inspect Auto-Detection (never opened while a report is being made)."""
        title = T("Έξυπνη ανίχνευση πεδίων", "Smart Field Detection")
        existing = getattr(self, '_detect_dialog', None)
        try:
            if existing is not None and existing.winfo_exists():
                existing.lift()
                return
        except tk.TclError:
            pass
        template = self._template_or_error(title)
        if not template:
            return
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            self._detect_dialog = DetectInspectDialog(self, template)
        except DetectError as e:
            messagebox.showwarning(title, str(e))
        except PermissionError:
            messagebox.showerror(title, T("Δεν μπορώ να γράψω στον προσωρινό φάκελο.", "I can't write to the temporary folder."))
        except Exception as e:
            messagebox.showerror(title, T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"))
        finally:
            self.config(cursor='')

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

    # ---- Smart Shift (settings)
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
        dframe = getattr(self, 'detect_frame', None)
        if dframe is not None:
            if self.prefs.get('pdf_mode') == 'detect':
                dframe.grid()
            else:
                dframe.grid_remove()
            self.detect_note_var.set(
                T("✎ Έχετε αποθηκευμένες προσαρμογές στον επεξεργαστή - θα χρησιμοποιηθούν στα επόμενα δελτία.",
                  "✎ You have fine-tuning saved in the editor - it will be used for the following forms.")
                if detect_cache_saved() else
                T("Η ανίχνευση τρέχει από την αρχή σε κάθε δελτίο πάνω στο τρέχον template.",
                  "The detection runs from scratch on the current template for every form."))
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
        previous, new = self.prefs.get('pdf_mode'), self.pv['pdf_mode'].get()
        if new != previous:
            # fine-tuning of the detected layout only lives while Smart Field Detection stays selected
            try:
                delete_detect_cache()
            except OSError:
                pass
        self._pref_changed()
        if self.prefs['pdf_mode'] == 'smart' and load_layout_reference() is None:
            self._offer_reference(T("Έξυπνη μετατόπιση", "Smart Shift"))

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
                T("Η έξυπνη μετατόπιση χρειάζεται ένα σημείο σύγκρισης: το template όπως είναι όταν το δελτίο "
                  "συμπληρώνεται σωστά.\n\nΝα αποθηκευτεί το τρέχον template ως σημείο σύγκρισης;\n"
                  "(Πατήστε «Ναι» μόνο αν με την προεπιλεγμένη μέθοδο το δελτίο βγαίνει σωστό.)",
                  "Smart Shift needs a point of comparison: the template as it is when the form fills in "
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
        text = describe_report(report)
        if not report['ok']:                    # a problem is worth a dialog
            messagebox.showwarning(title, text)
            return
        # all fine: one quiet line in the note under the method, back to normal after a few seconds
        first = next((ln.strip() for ln in text.splitlines() if ln.strip()), "")
        try:
            if getattr(self, '_check_job', None):
                self.after_cancel(self._check_job)
            self.layout_note_var.set(first)
            self.layout_note.configure(style='CardGood.TLabel')
            self._check_job = self.after(8000, self._update_layout_note)
        except (AttributeError, tk.TclError):
            pass

    def layout_reset_reference(self):
        title = T("Επαναφορά αναφοράς", "Reset reference")
        if load_layout_reference() is None:
            self._update_layout_note()
            return
        if not messagebox.askyesno(
                title, T("Η αποθηκευμένη διάταξη αναφοράς θα διαγραφεί. Η έξυπνη μετατόπιση δεν θα λειτουργεί μέχρι να "
                         "αποθηκεύσετε νέα.\n\nΣυνέχεια;",
                         "The saved reference layout will be deleted. Smart Shift will not work until you save a new one.\n\n"
                         "Continue?"), default='no'):
            return
        delete_layout_reference()
        self._update_layout_note()

    def layout_save_reference(self):
        title = T("Αποθήκευση ως αναφορά", "Save as reference")
        template = self._template_or_error(title)
        if not template:
            return
        if not messagebox.askyesno(
                title,
                T("Να χρησιμοποιηθεί το τρέχον template ως διάταξη αναφοράς;\n\n"
                  "Κάντε το μόνο αν το δελτίο συμπληρώνεται σωστά με την προεπιλεγμένη μέθοδο πάνω σε αυτό το template. "
                  "Η έξυπνη μετατόπιση θα μετράει τις αλλαγές των επόμενων templates σε σχέση με αυτό.",
                  "Use the current template as the reference layout?\n\n"
                  "Only do this if the form is filled in correctly with the default method on this template. "
                  "Smart Shift will measure the changes of later templates against it."), default='no'):
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
        p['manual_base'] = pv['manual_base'].get()
        set_manual_base(p['manual_base'])
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

    def open_help(self, section=None, anchor=None):
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
            self.guide.show(section, anchor)
        self.guide.refresh_back()

    def help_context(self):
        """Live values for the [[name]] placeholders of the guide."""
        hist = self._hist or {}
        last = max(((y, m) for y, ms in hist.items() for m in ms), default=None)
        modes = {'overlay': T("Παλιά μέθοδος (σταθερές συντεταγμένες)", "Legacy (fixed coordinates)"),
                 'smart': T("Έξυπνη μετατόπιση", "Smart Shift"),
                 'detect': T("Έξυπνη ανίχνευση πεδίων", "Smart Field Detection"),
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
        out.append(('ok' if excel_ok else 'warn', T("Excel / CSV του μήνα", "Month's Excel / CSV"),
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
        elif mode == 'detect':
            tuned = detect_cache_saved()
            out.append(('ok', T("Έξυπνη ανίχνευση πεδίων", "Smart Field Detection"),
                        T("υπάρχουν προσαρμογές από τον επεξεργαστή", "fine-tuning from the editor in use") if tuned
                        else T("ανίχνευση από την αρχή σε κάθε δελτίο", "fresh detection for every form"), 'tab:settings'))
        elif mode == 'smart':
            ref = load_layout_reference() is not None
            out.append(('ok' if ref else 'warn', T("Έξυπνη μετατόπιση", "Smart Shift"),
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
                return (T(f"Φτιάξτε το δελτίο {month_name(rem[1])} {rem[0]}: διαλέξτε το Excel / CSV του",
                          f"Make the form for {month_name(rem[1])} {rem[0]}: choose its Excel / CSV"), 'tab:report')
            return T("Διαλέξτε το Excel / CSV του μήνα", "Choose the month's Excel / CSV"), 'tab:report'
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
        self._update_extra_visibility()
        self.refresh_validation()

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
            self.custom_fields = load_custom_fields()
            self.cf_vars = {cf['id']: tk.StringVar(value=cf['value']) for cf in self.custom_fields}
            self._rebuild_screen()
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
        def work():
            ensure_data_dir()
            with tempfile.TemporaryDirectory() as tmp:
                z = shutil.make_archive(os.path.join(tmp, 'backup'), 'zip', DATA_DIR)
                shutil.copy2(z, fp)

        def fail(e):
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Δεν ήταν δυνατή η δημιουργία αντιγράφου:\n{e}",
                                             f"Could not create the backup:\n{e}"))
        self.run_async(work, lambda _r: self._status(T(f"✓ Αντίγραφο ασφαλείας: {fp}", f"✓ Backup: {fp}")),
                       fail, text=T("Δημιουργία αντιγράφου…", "Creating backup…"))

    # ---- online update (GitHub)
    def _ensure_update_vars(self):
        if hasattr(self, 'update_repo_var'):
            return
        cfg = load_app_config()
        self.update_repo_var = tk.StringVar(value=cfg.get('update_repo') or UPDATE_DEFAULT_REPO)
        self.update_branch_var = tk.StringVar(value=cfg.get('update_branch') or UPDATE_DEFAULT_BRANCH)
        self.update_auto_var = tk.BooleanVar(value=bool(cfg.get('update_check_on_start', False)))
        self.update_status_var = tk.StringVar(value='')
        self._update_pending = None

    def _update_save_cfg(self, *_):
        cfg = load_app_config()
        cfg['update_repo'] = self.update_repo_var.get().strip()
        cfg['update_branch'] = self.update_branch_var.get().strip() or UPDATE_DEFAULT_BRANCH
        cfg['update_check_on_start'] = bool(self.update_auto_var.get())
        save_app_config(cfg)

    def _build_update_card(self, body, row, polish, hint):
        self._ensure_update_vars()
        card = Card(body, None, T("Ενημέρωση προγράμματος", "Program update"))
        card.grid(row=row, column=0, sticky="ew", padx=4, pady=6)
        polish(card)
        f = card.body
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text=T(f"Τρέχουσα έκδοση: {APP_VERSION}", f"Current version: {APP_VERSION}"),
                  style='CardBold.TLabel').grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(f, text=T("Αποθετήριο GitHub:", "GitHub repository:"), style='Card.TLabel').grid(
            row=1, column=0, sticky="w", pady=(10, 0))
        repo_entry = ttk.Entry(f, textvariable=self.update_repo_var)
        repo_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(10, 0))
        ttk.Label(f, text=T("Κλάδος:", "Branch:"), style='Card.TLabel').grid(row=2, column=0, sticky="w", pady=(6, 0))
        branch_entry = ttk.Entry(f, textvariable=self.update_branch_var, width=18)
        branch_entry.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(6, 0))
        for e in (repo_entry, branch_entry):
            e.bind('<FocusOut>', self._update_save_cfg)
            e.bind('<Return>', self._update_save_cfg)
        btns = ttk.Frame(f, style='Card.TFrame')
        btns.grid(row=3, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Button(btns, text=T("Έλεγχος για ενημερώσεις", "Check for updates"),
                   command=self.check_for_updates).pack(side="left")
        self.update_install_btn = ttk.Button(btns, text=T("Εγκατάσταση ενημέρωσης", "Install update"),
                                             command=self.install_update,
                                             state='normal' if self._update_pending else 'disabled')
        self.update_install_btn.pack(side="left", padx=8)
        self.update_restore_btn = ttk.Button(btns, text=T("Επαναφορά προηγούμενης έκδοσης", "Restore previous version"),
                                             command=self.restore_previous)
        self.update_restore_btn.pack(side="left")
        self._update_refresh_button()
        lbl = hint(f, "", 4, columnspan=3, pady=(8, 0))
        lbl.configure(textvariable=self.update_status_var)
        ttk.Checkbutton(f, text=T("Έλεγχος για ενημέρωση όταν ανοίγει το πρόγραμμα", "Check for updates when the program starts"),
                        variable=self.update_auto_var, style='Card.TCheckbutton',
                        command=self._update_save_cfg).grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
        hint(f, T("Η ενημέρωση κατεβάζει το main.py από το αποθετήριο (ο κλάδος πρέπει να έχει μεγαλύτερο APP_VERSION). "
                  "Το παλιό αρχείο αποθηκεύεται ως main.py.bak και τα δεδομένα σας δεν αγγίζονται. "
                  "Λειτουργεί μόνο όταν το πρόγραμμα τρέχει ως αρχείο .py.",
                  "The update downloads main.py from the repository (the branch must have a higher APP_VERSION). "
                  "The old file is kept as main.py.bak and your data is not touched. "
                  "It only works when the program runs as a .py file."), 6, columnspan=3)

    def _update_refresh_button(self):
        try:
            self.update_install_btn.configure(state='normal' if self._update_pending else 'disabled')
            has_bak = os.path.isfile(os.path.abspath(__file__) + '.bak') and not getattr(sys, 'frozen', False)
            self.update_restore_btn.configure(state='normal' if has_bak else 'disabled')
        except (tk.TclError, AttributeError):
            pass

    def check_for_updates(self, silent=False):
        self._ensure_update_vars()
        repo = parse_github_repo(self.update_repo_var.get())
        branch = self.update_branch_var.get().strip() or UPDATE_DEFAULT_BRANCH
        if not repo:
            self.update_status_var.set(T("Γράψτε το αποθετήριο GitHub στη μορφή όνομα/αποθετήριο.",
                                         "Enter the GitHub repository as owner/repository."))
            if silent is not True:
                messagebox.showinfo(T("Ενημέρωση", "Update"), self.update_status_var.get(), parent=self)
            return
        self._update_save_cfg()
        self.update_status_var.set(T("Έλεγχος…", "Checking…"))

        def done(info):
            self._update_pending = info if info['newer'] else None
            if info['newer']:
                msg = T(f"Υπάρχει νέα έκδοση: {info['version']} (τρέχουσα {APP_VERSION}).",
                        f"A new version is available: {info['version']} (current {APP_VERSION}).")
            else:
                msg = T(f"Το πρόγραμμα είναι ενημερωμένο (έκδοση {APP_VERSION}).",
                        f"The program is up to date (version {APP_VERSION}).")
            self.update_status_var.set(msg + T(f"  (ελέγχθηκε {time.strftime('%H:%M')})",
                                               f"  (checked at {time.strftime('%H:%M')})"))
            self._update_refresh_button()
            if silent is True:
                if info['newer']:
                    self._status(msg + T(" Δείτε τις Ρυθμίσεις.", " See Settings."))
            else:
                messagebox.showinfo(T("Ενημέρωση", "Update"), msg, parent=self)

        def fail(e):
            if isinstance(e, urllib.error.HTTPError) and e.code == 404:
                msg = T(f"Δεν βρέθηκε το {UPDATE_FILE} στο {repo} (κλάδος {branch}). Αν το αποθετήριο είναι ιδιωτικό, "
                        "δεν υποστηρίζεται.", f"{UPDATE_FILE} was not found in {repo} (branch {branch}). "
                        "Private repositories are not supported.")
            elif isinstance(e, (urllib.error.URLError, OSError)):
                msg = T("Δεν υπάρχει σύνδεση με το GitHub. Ελέγξτε το internet και δοκιμάστε ξανά.",
                        "Could not reach GitHub. Check your internet connection and try again.")
            else:
                msg = T(f"Ο έλεγχος απέτυχε:\n{e}", f"The check failed:\n{e}")
            self.update_status_var.set(msg)
            if silent is not True:
                messagebox.showerror(T("Σφάλμα", "Error"), msg, parent=self)

        self.run_async(lambda: fetch_remote_update(repo, branch), done, fail,
                       text=T("Έλεγχος για ενημερώσεις…", "Checking for updates…"))

    def install_update(self):
        info = getattr(self, '_update_pending', None)
        if not info:
            return
        if getattr(sys, 'frozen', False):
            messagebox.showinfo(T("Ενημέρωση", "Update"),
                                T("Αυτή η έκδοση είναι πακεταρισμένη (.exe) και δεν μπορεί να ενημερωθεί από εδώ. "
                                  "Κατεβάστε τη νέα έκδοση από το GitHub.",
                                  "This is a packaged (.exe) build and cannot update itself. "
                                  "Download the new version from GitHub."), parent=self)
            return
        if not messagebox.askyesno(T("Ενημέρωση", "Update"),
                                   T(f"Να εγκατασταθεί η έκδοση {info['version']};\n\nΤο τρέχον αρχείο θα αποθηκευτεί ως "
                                     "main.py.bak και το πρόγραμμα θα επανεκκινήσει.",
                                     f"Install version {info['version']}?\n\nThe current file will be saved as "
                                     "main.py.bak and the program will restart.")
                                   + (T("\n\nΤι αλλάζει:\n", "\n\nWhat's new:\n") + info['notes'] if info.get('notes') else ''),
                                   default='no', parent=self):
            return

        def done(_backup):
            self._update_pending = None
            self._update_refresh_button()
            self.update_status_var.set(T(f"✓ Εγκαταστάθηκε η έκδοση {info['version']}.",
                                         f"✓ Version {info['version']} installed."))
            if messagebox.askyesno(T("Ενημέρωση", "Update"),
                                   T("Η ενημέρωση ολοκληρώθηκε. Επανεκκίνηση τώρα;", "The update is complete. Restart now?"),
                                   parent=self):
                self._restart()

        def fail(e):
            messagebox.showerror(T("Σφάλμα", "Error"),
                                 T(f"Η ενημέρωση δεν ολοκληρώθηκε (το πρόγραμμα δεν άλλαξε):\n{e}",
                                   f"The update did not complete (the program was not changed):\n{e}"), parent=self)

        self.run_async(lambda: apply_update(info['code']), done, fail,
                       text=T("Εγκατάσταση ενημέρωσης…", "Installing the update…"))

    def restore_previous(self):
        if not messagebox.askyesno(T("Επαναφορά", "Restore"),
                                   T("Να επανέλθει η προηγούμενη έκδοση του προγράμματος (main.py.bak);",
                                     "Go back to the previous version of the program (main.py.bak)?"),
                                   default='no', parent=self):
            return

        def done(_r):
            self._update_refresh_button()
            self.update_status_var.set(T("✓ Επαναφέρθηκε η προηγούμενη έκδοση.", "✓ The previous version was restored."))
            if messagebox.askyesno(T("Επαναφορά", "Restore"), T("Επανεκκίνηση τώρα;", "Restart now?"), parent=self):
                self._restart()

        def fail(e):
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Η επαναφορά απέτυχε:\n{e}", f"The restore failed:\n{e}"),
                                 parent=self)
        self.run_async(restore_backup, done, fail, text=T("Επαναφορά…", "Restoring…"))

    def _restart(self):
        try:
            subprocess.Popen([sys.executable] + sys.argv)
        except Exception as e:
            messagebox.showerror(T("Σφάλμα", "Error"),
                                 T(f"Κλείστε και ανοίξτε ξανά το πρόγραμμα χειροκίνητα.\n{e}",
                                   f"Please close and reopen the program yourself.\n{e}"), parent=self)
            return
        self.destroy()

    def _startup_forced_check(self):
        """Runs at every start, invisibly. If the main.py on GitHub has a higher APP_VERSION AND FORCE_UPDATE = True,
        it is installed right away and the program restarts. Otherwise (or without internet) nothing happens and
        updates stay optional."""
        if getattr(sys, 'frozen', False):
            return
        cfg = load_app_config()
        repo = parse_github_repo(cfg.get('update_repo') or UPDATE_DEFAULT_REPO)
        branch = (cfg.get('update_branch') or UPDATE_DEFAULT_BRANCH).strip() or UPDATE_DEFAULT_BRANCH
        if not repo:
            return
        results = queue.Queue()

        def worker():
            try:
                info = fetch_remote_update(repo, branch)
                if info['force'] and info['newer']:
                    apply_update(info['code'])
                    results.put(info['version'])
                else:
                    results.put(None)
            except BaseException:
                results.put(None)                      # offline, 404, bad file ...: stay silent, nothing was changed
        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                version = results.get_nowait()
            except queue.Empty:
                try:
                    self.after(200, poll)
                except tk.TclError:
                    pass
                return
            if not version:
                return
            try:
                self._update_pending = None
                self._update_refresh_button()
                if time.time() - getattr(self, '_boot_time', 0) < 40:     # still just starting: restart quietly
                    self._restart()
                elif messagebox.askyesno(T("Ενημέρωση", "Update"),
                                         T(f"Εγκαταστάθηκε μια απαραίτητη ενημέρωση (έκδοση {version}). "
                                           "Επανεκκίνηση τώρα;",
                                           f"A required update (version {version}) was installed. Restart now?"),
                                         parent=self):
                    self._restart()
            except tk.TclError:
                pass
        poll()

    # ---- keyboard shortcuts (settings)
    def _build_shortcuts_card(self, body, row, polish, hint):
        """One collapsible block per window; every row is "what it does" + the keys (click to change) + Reset."""
        if not hasattr(self, '_sc_open'):
            self._sc_open = set()                    # every category starts collapsed
        card = Card(body, None, T("Συντομεύσεις πληκτρολογίου", "Keyboard shortcuts"))
        card.grid(row=row, column=0, sticky="ew", padx=4, pady=6)
        polish(card)
        card.make_collapsible()
        f = card.body
        f.columnconfigure(0, weight=1)
        hint(f, T("Πατήστε πάνω σε μια συντόμευση και μετά τον νέο συνδυασμό πλήκτρων. Οι αλλαγές ισχύουν αμέσως, "
                  "ακόμη και σε παράθυρα που είναι ήδη ανοιχτά, και θυμούνται. Αν ο συνδυασμός χρησιμοποιείται ήδη "
                  "στο ίδιο παράθυρο, περνάει σε αυτό που διαλέγετε.",
                  "Click a shortcut, then press the new key combination. Changes apply immediately, even in windows "
                  "that are already open, and are remembered. If the combination is already used in the same window "
                  "it moves to the one you choose."), 0, pady=(0, 0), margin=130)
        bar = ttk.Frame(f, style='Card.TFrame')
        bar.grid(row=1, column=0, sticky="ew", pady=(10, 8))
        ttk.Button(bar, text=T("Επαναφορά όλων των συντομεύσεων", "Reset all shortcuts"),
                   command=self._reset_all_shortcuts).pack(side="left")
        self._sc_summary = ttk.Label(bar, style='CardHint.TLabel')
        self._sc_summary.pack(side="left", padx=12)
        self._sc_rows, self._sc_badges = {}, {}
        for n, (gid, gtitle) in enumerate(SHORTCUT_GROUPS):
            self._build_shortcut_group(f, 2 + n, gid, gtitle)
        self._refresh_shortcut_rows()

    def _build_shortcut_group(self, parent, row, gid, gtitle):
        outer = ttk.Frame(parent, style='Card.TFrame')
        outer.grid(row=row, column=0, sticky="ew")
        outer.columnconfigure(0, weight=1)
        ttk.Separator(outer, orient="horizontal").grid(row=0, column=0, sticky="ew")
        head = tk.Frame(outer, bg=CARD, cursor='hand2')
        head.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        arrow = tk.Label(head, bg=CARD, fg=ACCENT, font=(FONT, 11, 'bold'), width=2, cursor='hand2')
        arrow.pack(side="left")
        title = tk.Label(head, text=T(*gtitle), bg=CARD, fg=NAVY, font=(FONT, 11, 'bold'), cursor='hand2')
        title.pack(side="left")
        badge = tk.Label(head, bg=ACCENT_SOFT, fg=ACCENT, font=(FONT, 9, 'bold'), padx=8, pady=1, cursor='hand2')
        badge.pack(side="left", padx=10)
        self._sc_badges[gid] = badge
        content = ttk.Frame(outer, style='Card.TFrame')
        content.grid(row=2, column=0, sticky="ew", padx=(26, 0), pady=(0, 8))
        content.columnconfigure(0, weight=1)
        for r, (sid, _g, label, _default, _flags) in enumerate(s for s in SHORTCUTS if s[1] == gid):
            ttk.Label(content, text=T(*label), style='Card.TLabel').grid(row=r, column=0, sticky="w", pady=3)
            chip = tk.Label(content, bg=ACCENT_SOFT, fg=ACCENT, font=(FONT, 9, 'bold'), padx=10, pady=3,
                            cursor='hand2', width=16)
            chip.grid(row=r, column=1, padx=(12, 8), pady=3)
            chip.bind('<Button-1>', lambda e, s=sid: self._edit_shortcut(s))
            btn = ttk.Button(content, text=T("Επαναφορά", "Reset"), width=10,
                             command=lambda s=sid: self._reset_shortcut(s))
            btn.grid(row=r, column=2, pady=3)
            self._sc_rows[sid] = (chip, btn)

        def show():
            is_open = gid in self._sc_open
            arrow.configure(text="\u25be" if is_open else "\u25b8")
            if is_open:
                content.grid()
            else:
                content.grid_remove()

        def toggle(_e=None):
            if gid in self._sc_open:
                self._sc_open.discard(gid)
            else:
                self._sc_open.add(gid)
            show()

        for w in (head, arrow, title, badge):
            w.bind('<Button-1>', toggle)
        show()

    def _refresh_shortcut_rows(self):
        rows = getattr(self, '_sc_rows', None)
        if not rows:
            return
        per_group, changed_total = {}, 0
        for sid, (chip, btn) in rows.items():
            default = shortcut_is_default(sid)
            accel = shortcut_accel(sid)
            try:
                chip.configure(text=accel_display(accel), fg=ACCENT if accel else MUTED)
                btn.state(['disabled' if default else '!disabled'])
            except tk.TclError:
                continue
            gid = SHORTCUT_INFO[sid]['group']
            total, changed = per_group.get(gid, (0, 0))
            per_group[gid] = (total + 1, changed + (0 if default else 1))
            changed_total += 0 if default else 1
        try:
            for gid, badge in self._sc_badges.items():
                total, changed = per_group.get(gid, (0, 0))
                badge.configure(text=str(total) + (T(f" · {changed} αλλαγμένες", f" · {changed} changed") if changed else ""))
            self._sc_summary.configure(
                text=T(f"{changed_total} συντομεύσεις έχουν αλλάξει.", f"{changed_total} shortcut(s) changed.")
                if changed_total else T("Όλες οι συντομεύσεις είναι οι προεπιλεγμένες.", "All shortcuts are at their defaults."))
        except tk.TclError:
            pass

    def _edit_shortcut(self, sid):
        ShortcutDialog(self, sid, lambda accel, s=sid: self._apply_shortcut(s, accel))

    def _apply_shortcut(self, sid, accel):
        accel = accel_normalize(accel)
        other = shortcut_conflict(sid, accel)
        if other:
            set_shortcut(other, '')                 # the one that had these keys is left without
        set_shortcut(sid, accel)
        self._shortcuts_changed()
        label = T(*SHORTCUT_INFO[sid]['label'])
        self._status(T(f"Συντόμευση «{label}»: {accel_display(accel)}", f"Shortcut “{label}”: {accel_display(accel)}"))

    def _reset_shortcut(self, sid):
        self._apply_shortcut(sid, SHORTCUT_INFO[sid]['default'])

    def _reset_all_shortcuts(self):
        if not SHORTCUT_OVERRIDES:
            return
        if not messagebox.askyesno(T("Επαναφορά συντομεύσεων", "Reset shortcuts"),
                                   T("Να γυρίσουν όλες οι συντομεύσεις στις προεπιλογές;",
                                     "Reset all keyboard shortcuts to their defaults?"), default='no'):
            return
        load_shortcut_overrides({})
        self._shortcuts_changed()

    def _shortcuts_changed(self):
        self.prefs['shortcuts'] = dict(SHORTCUT_OVERRIDES)
        self._save_prefs()
        ShortcutBinder.refresh_all()
        self._refresh_shortcut_rows()

    def reset_prefs(self):
        if not messagebox.askyesno(T("Επαναφορά ρυθμίσεων", "Reset settings"),
                                   T("Να γυρίσουν όλες οι ρυθμίσεις στις προεπιλογές;\n"
                                     "(Δεν διαγράφονται δεδομένα ή ιστορικό.)",
                                     "Reset all settings to their defaults?\n"
                                     "(No data or history is deleted.)"), default='no'):
            return
        keep_template = self.prefs.get('template_pdf', '')
        try:
            delete_detect_cache()
        except OSError:
            pass
        self.prefs = dict(DEFAULT_PREFS, template_pdf=keep_template, language=self.prefs['language'], shortcuts={})
        load_shortcut_overrides({})
        ShortcutBinder.refresh_all()
        self._refresh_shortcut_rows()
        self._sync_pref_vars()
        self._save_prefs()
        self._after_prefs_change()
        self._status(T("Οι ρυθμίσεις επανήλθαν στις προεπιλογές.", "Settings were reset to their defaults."))

    # ---------------------------------------------------- background work
    def _busy_refresh(self):
        try:
            if self._badge is None or not self._badge.winfo_exists():
                self._badge = LoadingBadge(self)
            if self._busy_count and self.prefs.get('show_loading', True):
                self._badge.start(self._busy_text)
            else:
                self._badge.stop()
            self.config(cursor='watch' if self._busy_count else '')
        except tk.TclError:
            pass

    def _busy_begin(self, text=''):
        self._busy_count += 1
        self._busy_text = text or T("Παρακαλώ περιμένετε…", "Please wait…")
        self._busy_refresh()

    def _busy_end(self):
        self._busy_count = max(0, self._busy_count - 1)
        self._busy_refresh()

    def run_async(self, work, done=None, error=None, text=''):
        """Runs work() in a background thread so the window never freezes; done(result) / error(exception)
        are called back in the main thread (Tk is not thread-safe)."""
        results = queue.Queue()

        def target():
            try:
                results.put(('ok', work()))
            except BaseException as exc:          # reported in the main thread
                results.put(('err', exc))
        self._busy_begin(text)
        threading.Thread(target=target, daemon=True).start()

        def poll():
            try:
                kind, value = results.get_nowait()
            except queue.Empty:
                try:
                    self.after(30, poll)
                except tk.TclError:
                    pass
                return
            self._busy_end()
            try:
                if kind == 'ok':
                    if done:
                        done(value)
                elif error:
                    error(value)
                else:
                    messagebox.showerror(T("Σφάλμα", "Error"), T(f"Κάτι πήγε στραβά:\n{value}",
                                                                 f"Something went wrong:\n{value}"))
            except tk.TclError:
                pass                              # the screen was rebuilt while the work was running
        poll()

    # ---------------------------------------------------- window behaviour
    def _center(self):
        try:
            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            if w <= 1 or h <= 1:            # not shown yet: use the size that was asked for
                m = re.match(r'(\d+)x(\d+)', self.geometry())
                w, h = (int(m.group(1)), int(m.group(2))) if m else (920, 840)
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            w, h = min(w, sw), min(h, sh - 60)
            x = max((sw - w) // 2, 0)
            y = max((sh - h) // 3, 0)
            self.geometry(f"{w}x{h}+{x}+{y}")
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
        s = self._raw_settings()
        if self.prefs.get('show_extra_fields', True):
            for cf in self.custom_fields:
                s[CF_PREFIX + cf['id']] = self.cf_vars[cf['id']].get().strip()
        else:                       # the drop-down is switched off: nothing of it goes on the form
            for k in EXTRA_KEYS:
                s[k] = ''
        m = {k: v.get().strip() for k, v in self.mon_vars.items()}
        return s, m, self.excel_var.get().strip(), self.pdf_var.get().strip(), self.out_var.get().strip()

    def _raw_settings(self):
        return {k: v.get().strip() for k, v in self.set_vars.items()}

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
            zone.set_text(T("Σύρετε εδώ το αρχείο Excel ή CSV", "Drop the Excel or CSV file here"),
                          T("ή πατήστε για να το διαλέξετε από τον υπολογιστή (.xls, .xlsx, .csv)",
                            "or click to choose it from your computer (.xls, .xlsx, .csv)"))
        else:
            zone.set_text(T("Πατήστε για να διαλέξετε το αρχείο Excel ή CSV", "Click to choose the Excel or CSV file"),
                          T("το αρχείο «Nationalities Statistics» από το Protel (.xls, .xlsx ή .csv)",
                            "the “Nationalities Statistics” file from Protel (.xls, .xlsx or .csv)"))

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
            errors = [T("Επιλέξτε πρώτα το αρχείο Excel ή CSV από το Protel (Βήμα 1).",
                        "First choose the Excel or CSV file from Protel (Step 1).")] + validate_hotel(s)
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
        return self._drop_one(event, "Excel / CSV (.xls/.xlsx/.csv)", {'.xls', '.xlsx', '.csv'}, self._accept_dropped_excel)

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
        fp = filedialog.askopenfilename(title=T("Διαλέξτε το Excel ή CSV από το Protel",
                                                "Choose the Excel or CSV file from Protel"), initialdir=self._excel_start_dir(),
                                        filetypes=[("Excel / CSV", "*.xls *.xlsx *.csv"), ("Excel Files", "*.xls *.xlsx"), ("CSV", "*.csv")])
        if fp:
            self._accept_dropped_excel(fp)

    def _excel_typed(self, _event=None):
        fp = self.excel_var.get().strip()
        if fp and fp != self._loaded_excel:
            self.load_period(fp)

    def load_period(self, fp):
        """Reads the month from the Excel header (in the background) and fills in step 2 and the output name."""
        self._loaded_excel = fp
        if not os.path.isfile(fp):
            self._show_period(lambda: T("⚠ Το αρχείο δεν βρέθηκε.", "⚠ The file was not found."),
                              'CardBad.TLabel')
            return
        self._load_token += 1
        token = self._load_token

        def work():
            bad = protel_problems(fp)
            if bad:
                return None, None, bad
            try:
                rows = read_nationality_rows(fp)[0]
            except Exception:
                rows = None
            try:
                info = period_as_month(read_excel_period(fp))
            except Exception:
                info = None
            return rows, info, []

        def done(res):
            if token != self._load_token:
                return
            rows, info, bad = res
            if bad:                                  # not a recognised Protel export: let the user decide
                if ask_use_anyway(bad):
                    PROTEL_OVERRIDES.add(_override_key(fp))
                    self._load_token += 1
                    self.load_period(fp)             # read it again, this time without the check
                    return
                self.excel_var.set('')
                self._loaded_excel = ''
                self.preview_frame.grid_remove()
                self._show_period(
                    lambda: T("✗ Το αρχείο δεν είναι αναφορά Protel και δεν χρησιμοποιήθηκε.",
                              "✗ This is not a Protel report and was not used."),
                    'CardBad.TLabel')
                return
            self._apply_preview(rows)
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
        self._show_period(lambda: T("Ανάγνωση του Excel…", "Reading the Excel…"), 'CardHint.TLabel')
        self.run_async(work, done, text=T("Ανάγνωση Excel…", "Reading Excel…"))

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
                                T(f"Δεν υπάρχει αρχείο Excel ή CSV στον φάκελο:\n{self._excel_start_dir()}\n\n"
                                  "Πατήστε «Επιλογή Excel / CSV…» για να το διαλέξετε.",
                                  f"There is no Excel or CSV file in the folder:\n{self._excel_start_dir()}\n\n"
                                  "Click “Choose Excel / CSV…” to select one."))
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
        """Reads the countries of the Excel in the background and shows them."""
        self._load_token += 1
        token = self._load_token

        def work():
            try:
                return read_nationality_rows(fp)[0]
            except Exception:
                return None

        def done(rows):
            if token == self._load_token:
                self._apply_preview(rows)
        self.run_async(work, done, text=T("Ανάγνωση Excel…", "Reading Excel…"))

    def _apply_preview(self, rows):
        """Shows the countries found in the Excel and flags the ones the program does not know."""
        self.preview_frame.grid_remove()
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
        tab.rowconfigure(2, weight=1)

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

        # ---- view switch: Overview (guests & nights) | Economics
        sw = ttk.Frame(tab)
        sw.grid(row=1, column=0, sticky="w", padx=8, pady=(6, 8))
        self.hist_view_seg = Segmented(sw, (('overview', T("Επισκόπηση", "Overview")),
                                            ('economics', T("Οικονομικά", "Economics"))),
                                       self._hist_view_changed, getattr(self, '_hist_view', 'overview'))
        self.hist_view_seg.pack(side="left")

        content = ttk.Frame(tab)
        content.grid(row=2, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)
        ov = self.hist_overview = ttk.Frame(content)
        ov.grid(row=0, column=0, sticky="nsew")
        ov.columnconfigure(0, weight=1)
        ov.rowconfigure(3, weight=1)
        self.hist_econ = ttk.Frame(content)
        self.hist_econ.grid(row=0, column=0, sticky="nsew")
        self._build_economics(self.hist_econ)

        info_row = ttk.Frame(ov)
        info_row.grid(row=0, column=0, sticky="ew", padx=8, pady=(0, 8))
        info_row.columnconfigure(0, weight=1)
        self.hist_info = tk.StringVar()
        self.hist_info_lbl = ttk.Label(info_row, textvariable=self.hist_info, style='Hint.TLabel', justify="left")
        self.hist_info_lbl.grid(row=0, column=0, sticky="w")
        autowrap(self.hist_info_lbl, tab, margin=260, minimum=200)
        ttk.Button(info_row, text=T("Σύμπτυξη", "Collapse"), command=lambda: self._tree_expand(False)).grid(
            row=0, column=1, padx=(8, 0))
        ttk.Button(info_row, text=T("Ανάπτυξη", "Expand"), command=lambda: self._tree_expand(True)).grid(
            row=0, column=2, padx=(6, 0))

        kp = ttk.Frame(ov)
        kp.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 8))
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

        self.hist_chart = HistoryChart(ov, on_pick=self._chart_pick, height=170)
        self.hist_chart.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 8))

        pane = ttk.PanedWindow(ov, orient="horizontal")
        pane.grid(row=3, column=0, sticky="nsew", padx=6, pady=(0, 8))

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
        self.hist_tree.bind('<Double-1>', self._hist_double)
        self._hist_shortcuts = ShortcutBinder(self.hist_tree)     # changeable in Settings -> Keyboard shortcuts
        self._hist_shortcuts.add('hist_open', lambda e: self.hist_open())
        self._hist_shortcuts.add('hist_delete', lambda e: self.hist_delete())
        self._hist_shortcuts.apply()

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

        act = self.hist_actions = ttk.Frame(tab)
        act.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 12))
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
        self._hist_view_changed(getattr(self, '_hist_view', 'overview'))

    # ---- history: economics view
    def _hist_view_changed(self, view):
        self._hist_view = view
        if view == 'economics':
            self.hist_overview.grid_remove()
            self.hist_econ.grid()
            self.hist_actions.grid_remove()
            self.update_economics()
        else:
            self.hist_econ.grid_remove()
            self.hist_overview.grid()
            self.hist_actions.grid()

    def _build_economics(self, parent):
        """Economics: everything here is computed from the Protel Excel (saved with each month's history)."""
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(3, weight=1)
        self._econ_metric = getattr(self, '_econ_metric', 'revenue')

        # four headline numbers
        kp = self.econ_kp = ttk.Frame(parent)
        kp.grid(row=0, column=0, sticky="ew", padx=6, pady=(0, 8))
        parent.bind('<Configure>', self._econ_relayout)
        self.winfo_toplevel().bind('<Configure>', lambda e: self._econ_fit_height() if e.widget is self.winfo_toplevel() else None, add='+')
        for i in range(4):
            kp.columnconfigure(i, weight=1, uniform='ekpi')
        titles = (T("ΣΥΝΟΛΙΚΑ ΕΣΟΔΑ", "TOTAL REVENUE"), T("ΕΣΟΔΑ ΑΝΑ ΔΙΑΝΥΚΤ. ΔΩΜΑΤΙΟΥ", "REVENUE PER ROOM NIGHT"),
                  T("ΕΣΟΔΑ ΑΝΑ ΔΙΑΝΥΚΤ. ΠΕΛΑΤΗ", "REVENUE PER GUEST NIGHT"), T("ΣΥΝΘΕΣΗ ΕΣΟΔΩΝ", "REVENUE MIX"))
        self.econ_cards = []
        for i, title in enumerate(titles):
            card = StatCard(kp, title)
            card.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 4, 0 if i == 3 else 4))
            self.econ_cards.append(card)

        # one chart, with a switch for what it shows
        box = ttk.Frame(parent)
        box.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 8))
        box.columnconfigure(0, weight=1)
        head = ttk.Frame(box)
        head.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self.econ_seg = Segmented(head, (('revenue', T("Έσοδα", "Revenue")),
                                         ('adr', T("Ανά διανυκτ. δωματίου", "Per room night")),
                                         ('pernight', T("Ανά διανυκτ. πελάτη", "Per guest night"))),
                                  self._econ_metric_changed, self._econ_metric)
        self.econ_seg.pack(side="left")
        self.econ_fill = ttk.Button(head, text=T("Προσθήκη εσόδων παλιών μηνών…", "Add revenue for older months…"),
                                    command=self._econ_add_revenue)
        self.econ_fill.pack(side="right")
        self.econ_chart = HistoryChart(box, height=170)
        self.econ_chart.grid(row=1, column=0, sticky="ew")

        # plain-language summary + month table
        self.econ_insight = tk.StringVar()
        lbl = ttk.Label(parent, textvariable=self.econ_insight, style='Hint.TLabel', justify="left")
        lbl.grid(row=2, column=0, sticky="w", padx=6, pady=(0, 6))
        autowrap(lbl, parent, margin=60, minimum=200)
        lower = ttk.Frame(parent)
        lower.grid(row=3, column=0, sticky="nsew", padx=6, pady=(0, 12))
        lower.columnconfigure(0, weight=1)
        lower.rowconfigure(1, weight=1)
        self._econ_view = getattr(self, '_econ_view', 'months')
        self._econ_sort = getattr(self, '_econ_sort', ('r', True))
        self._econ_rows = []
        vbar = ttk.Frame(lower)
        vbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        self.econ_view_seg = Segmented(vbar, (('months', T("Ανά μήνα", "By month")),
                                              ('countries', T("Ανά χώρα", "By country"))),
                                       self._econ_view_changed, self._econ_view)
        self.econ_view_seg.pack(side="left")
        cols = ('n', 'r', 'rm', 'fb', 'ex', 'pn', 'd')
        self.econ_tree = ttk.Treeview(lower, columns=cols, show='tree headings', height=5)
        self.econ_tree.heading('#0', text=T('Μήνας', 'Month'), anchor='w')
        for key, text in (('n', T('Διανυκτ.', 'Nights')), ('r', T('Έσοδα', 'Revenue')),
                          ('rm', T('Δωμάτια', 'Rooms')), ('fb', 'F&B'), ('ex', T('Έξτρα', 'Extras')),
                          ('pn', T('€/διανυκτ.', '€/night')), ('d', T('έναντι προηγ. έτους', 'vs prev. year'))):
            self.econ_tree.heading(key, text=text)
            self.econ_tree.column(key, width=84, minwidth=64, anchor='e', stretch=True)
        self.econ_tree.column('#0', width=130, minwidth=90, stretch=True)
        self.econ_tree.tag_configure('total', font=(FONT, 10, 'bold'), background=HEAD_BG)
        self.econ_tree.tag_configure('row', background=CARD)
        self.econ_tree.tag_configure('alt', background=ROW_MONTH)
        esb = ttk.Scrollbar(lower, orient="vertical", command=self.econ_tree.yview)
        self.econ_tree.configure(yscrollcommand=esb.set)
        self.econ_tree.grid(row=1, column=0, sticky="nsew")
        esb.grid(row=1, column=1, sticky="ns")
        self.econ_esb = esb
        self._fit_cols(self.econ_tree, cols, ('r', 'n', 'pn', 'd', 'rm', 'fb', 'ex'))

        # by country: the columns of the Protel sheet (revenue per room night / per guest night are its own two ratios)
        ccols = ('g', 'n', 'rm', 'fb', 'ex', 'r', 'sh', 'rn', 'pn')
        self.econ_ctree = ttk.Treeview(lower, columns=ccols, show='tree headings', height=6)
        self.econ_ctree.column('#0', width=150, minwidth=100, stretch=True)
        self._econ_cnames = {'c': T('Χώρα', 'Country'), 'g': T('Πελάτες', 'Guests'), 'n': T('Διανυκτ.', 'Nights'),
                             'rm': T('Δωμάτια', 'Rooms'), 'fb': 'F&B', 'ex': T('Έξτρα', 'Extras'),
                             'r': T('Σύνολο', 'Total'), 'sh': T('% εσόδων', '% share'),
                             'rn': T('€/διανυκτ. δωμ.', '€/room night'), 'pn': T('€/διανυκτ. πελ.', '€/guest night')}
        self.econ_ctree.heading('#0', command=lambda: self._econ_sort_by('c'))
        for key in ccols:
            self.econ_ctree.heading(key, command=lambda k=key: self._econ_sort_by(k))
            self.econ_ctree.column(key, width=84, minwidth=64, anchor='e', stretch=True)
        self.econ_ctree.tag_configure('total', font=(FONT, 10, 'bold'), background=HEAD_BG)
        self.econ_ctree.tag_configure('row', background=CARD)
        self.econ_ctree.tag_configure('alt', background=ROW_MONTH)
        csb = ttk.Scrollbar(lower, orient="vertical", command=self.econ_ctree.yview)
        self.econ_ctree.configure(yscrollcommand=csb.set)
        self.econ_ctree.grid(row=1, column=0, sticky="nsew")
        csb.grid(row=1, column=1, sticky="ns")
        self.econ_csb = csb
        self._fit_cols(self.econ_ctree, ccols, ('r', 'sh', 'g', 'n', 'pn', 'rn', 'rm', 'fb', 'ex'))
        self._econ_view_changed(self._econ_view)

    def _fit_cols(self, tree, cols, priority):
        """Shows as many columns as fit the width (most important first), so the table scales with the window."""
        import tkinter.font as tkfont
        fnt = tkfont.Font(family=FONT, size=9, weight='bold')

        def need(c):
            txt = str(tree.heading(c, 'text')) + "  ▼"
            return max(64, fnt.measure(txt) + 22)

        def fit(_e=None):
            avail = tree.winfo_width() - 150
            usable = [c for c in priority if c in cols and c not in getattr(tree, '_hidden_cols', ())]
            keep, used = [], 0
            for c in usable:
                w = need(c)
                if len(keep) >= 3 or used + w <= avail:
                    keep.append(c)
                    used += w
            for c in keep:
                tree.column(c, width=need(c), minwidth=need(c))
            want = [c for c in cols if c in keep]
            if list(tree.cget('displaycolumns')) != want:
                tree.configure(displaycolumns=want)
        tree._refit = fit
        tree.bind('<Configure>', fit, add='+')

    def _econ_fit_height(self, _e=None):
        """Chart and table share the window's spare height, measured from where they actually sit."""
        try:
            top = self.winfo_toplevel()
            lower = self.econ_tree.master
            if not lower.winfo_ismapped() or top.winfo_height() < 200:
                return
            lower_y = lower.winfo_rooty() - top.winfo_rooty()
            fixed = lower_y - self.econ_chart.winfo_height() + 44 + 70   # stuff above chart + view switch + bottom bar
            avail = top.winfo_height() - fixed
            n_rows = len(self.econ_tree.get_children()) if self._econ_view != 'countries' \
                else len(self.econ_ctree.get_children())
        except tk.TclError:
            return
        chart_h = int(max(130, min(320, avail * 0.38)))
        rows = int((avail - chart_h - 30) // 26)
        rows = max(3, min(rows, max(n_rows, 3), 28))
        if abs(chart_h - int(self.econ_chart.cget('height'))) >= 6:
            self.econ_chart.configure(height=chart_h)
        for t in (self.econ_tree, self.econ_ctree):
            if int(t.cget('height')) != rows:
                t.configure(height=rows)

    def _econ_relayout(self, event):
        """Four headline cards in one row when there is room, two by two when the window is narrow."""
        self._econ_width = event.width
        self._econ_fit_height()
        self._econ_layout_cards()

    def _econ_layout_cards(self):
        vis = [c for i, c in enumerate(self.econ_cards) if i not in getattr(self, '_econ_hidden_cards', ())]
        width = getattr(self, '_econ_width', 900)
        per_row = max(1, min(len(vis), width // 190, 4))
        if len(vis) == 4 and per_row == 3:
            per_row = 2
        key = (per_row, tuple(id(c) for c in vis))
        if key == getattr(self, '_econ_kkey', None):
            return
        self._econ_kkey = key
        for c in self.econ_cards:
            c.grid_remove()
        for i in range(4):
            self.econ_kp.columnconfigure(i, weight=1 if i < per_row else 0, uniform='ekpi' if i < per_row else '')
        rows = -(-len(vis) // per_row) if vis else 0
        for i, card in enumerate(vis):
            card.grid(row=i // per_row, column=i % per_row, sticky="ew", padx=4,
                      pady=(0, 6 if i // per_row < rows - 1 else 0))

    def _econ_view_changed(self, value):
        self._econ_view = value
        months = value != 'countries'
        for w, show in ((self.econ_tree, months), (self.econ_esb, months),
                        (self.econ_ctree, not months), (self.econ_csb, not months)):
            (w.grid if show else w.grid_remove)()

    def _econ_sort_by(self, key):
        cur, desc = self._econ_sort
        self._econ_sort = (key, (not desc) if key == cur else key != 'c')
        self._econ_render_countries()

    def _econ_render_countries(self):
        tree = self.econ_ctree
        tree.delete(*tree.get_children())
        key, desc = self._econ_sort
        for k, name in self._econ_cnames.items():             # arrow on the sorted column
            mark = (" ▼" if desc else " ▲") if k == key else ""
            (tree.heading('#0', text=name + mark) if k == 'c' else tree.heading(k, text=name + mark))
        rows = sorted(self._econ_rows, key=lambda r: r[key] if key != 'c' else r['c'].lower(), reverse=desc)
        dash = "–"
        eur = lambda v: f"{fmt_int(v)} €"
        ratio = lambda v: f"{fmt_dec(v)} €" if v else dash
        tree._hidden_cols = {k for k in ('rn', 'pn', 'fb', 'ex', 'g') if not any(r[k] for r in rows)}
        tree._refit()
        for n_, r in enumerate(rows):
            tree.insert('', 'end', text=r['c'], tags=('alt' if n_ % 2 else 'row',), values=(
                fmt_int(r['g']), fmt_int(r['n']), eur(r['rm']), eur(r['fb']), eur(r['ex']), eur(r['r']),
                f"{fmt_dec(r['sh'])}%", ratio(r['rn']), ratio(r['pn'])))
        if rows:
            sm = lambda k: sum(r[k] for r in rows)
            tree.insert('', 'end', text=T("ΣΥΝΟΛΟ", "TOTAL"), tags=('total',), values=(
                fmt_int(sm('g')), fmt_int(sm('n')), eur(sm('rm')), eur(sm('fb')), eur(sm('ex')), eur(sm('r')), "100%",
                ratio(sm('r') / sm('roomn') if sm('roomn') else 0), ratio(sm('r') / sm('n') if sm('n') else 0)))

    def _econ_add_revenue(self):
        paths = filedialog.askopenfilenames(title=T("Επιλέξτε το «Nationalities Statistics» (ένα ή περισσότερους μήνες)",
                                                    "Choose the “Nationalities Statistics” file(s) (one or more months)"),
                                            filetypes=[("Excel / CSV", "*.xls *.xlsx *.csv"), ("Excel Files", "*.xls *.xlsx"), ("CSV", "*.csv")])
        if not paths:
            return
        done, errors = [], []
        for p in paths:
            try:
                bad = protel_problems(p)
                if bad and ask_use_anyway([os.path.basename(p)] + bad):
                    PROTEL_OVERRIDES.add(_override_key(p))
                y, m, tot = attach_revenue(p)
                done.append(f"{month_name(m)} {y}: {fmt_int(tot)} €")
            except Exception as e:
                errors.append(f"{os.path.basename(p)}: {e}")
        self.update_economics()
        msg = "\n".join(done) if done else ""
        if errors:
            msg += ("\n\n" if msg else "") + "\n".join(errors)
        (messagebox.showinfo if done and not errors else messagebox.showwarning)(T("Έσοδα", "Revenue"), msg)

    def _econ_metric_changed(self, value):
        self._econ_metric = value
        self.update_economics()

    def update_economics(self, *_):
        if not hasattr(self, 'econ_tree'):
            return
        dash = "–"
        try:
            year = int(self.hist_year_var.get())
        except ValueError:
            year = None
        self.econ_tree.delete(*self.econ_tree.get_children())
        self._econ_rows = []
        self._econ_render_countries()
        for c in self.econ_cards:
            c.set(dash, "")
        self._econ_hidden_cards = ()
        self._econ_layout_cards()
        self.econ_seg.show_option('adr', True)
        if year is None or year not in self._hist:
            self.econ_chart.set_data([0] * 12, T("Δεν υπάρχουν δεδομένα ακόμη", "No data yet"))
            self.econ_insight.set("")
            return
        hist, real = self._hist, load_revenue()

        def nights_of(y, m):
            return sum(r[2] for r in hist.get(y, {}).get(m, []))

        def guests_of(y, m):
            return sum(r[1] for r in hist.get(y, {}).get(m, []))

        def rev(y, m):
            r = real.get((y, m))
            return r if r and r['total'] > 0 else None

        months = [m for m in sorted(hist[year])]
        with_rev = [m for m in months if rev(year, m)]
        missing = [m for m in months if not rev(year, m)]
        if not with_rev:
            self.econ_chart.set_data([0] * 12, T("Δεν υπάρχουν έσοδα ακόμη", "No revenue yet"))
            self.econ_insight.set(T("Τα έσοδα αποθηκεύονται αυτόματα όταν δημιουργείτε τη φόρμα. Για παλιούς μήνες "
                                    "πατήστε «Προσθήκη εσόδων παλιών μηνών…».",
                                    "Revenue is saved automatically whenever you make the form. For older months "
                                    "press “Add revenue for older months…”."))
            return

        tot = sum(rev(year, m)['total'] for m in with_rev)
        rooms = sum(rev(year, m)['rooms'] for m in with_rev)
        fb = sum(rev(year, m)['fb'] for m in with_rev)
        extras = sum(rev(year, m)['extras'] for m in with_rev)
        roomn = sum(rev(year, m)['roomn'] for m in with_rev)
        nights = sum(nights_of(year, m) for m in with_rev)
        guests = sum(guests_of(year, m) for m in with_rev)

        # headline cards
        nm = len(with_rev)
        sub, col = (T("1 μήνας", "1 month") if nm == 1 else T(f"{nm} μήνες", f"{nm} months")), None
        both = [m for m in with_rev if rev(year - 1, m)]
        if both:
            a = sum(rev(year, m)['total'] for m in both)
            b = sum(rev(year - 1, m)['total'] for m in both)
            if b:
                pct = (a - b) / b * 100
                sub = f"{'▲' if pct > 0.05 else '▼' if pct < -0.05 else '■'} {fmt_dec(pct, signed=True)}% " + \
                      T(f"έναντι {year - 1}", f"vs {year - 1}")
                col = GOOD if pct > 0.05 else BAD if pct < -0.05 else None
        self.econ_cards[0].set(f"{fmt_int(tot)} €", sub, col)
        self._econ_hidden_cards = () if roomn else (1,)
        self.econ_seg.show_option('adr', bool(roomn))
        if not roomn and getattr(self, '_econ_metric', '') == 'adr':
            self._econ_metric = 'revenue'
            self.econ_seg.set('revenue')
        self._econ_layout_cards()
        self.econ_cards[1].set(f"{fmt_dec(tot / roomn)} €" if roomn else dash,
                               T("συνολικά έσοδα ÷ διανυκτ. δωματίων", "total revenue ÷ room nights"))
        self.econ_cards[2].set(f"{fmt_dec(tot / nights)} €" if nights else dash,
                               T("συνολικά έσοδα ÷ διανυκτ. πελατών", "total revenue ÷ guest nights"))
        self.econ_cards[3].set(f"{fmt_dec(rooms / tot * 100)}% " + T("δωμάτια", "rooms") if tot else dash,
                               f"F&B {fmt_dec(fb / tot * 100)}% · " + T("έξτρα", "extras") + f" {fmt_dec(extras / tot * 100)}%" if tot else "")

        # chart
        metric = getattr(self, '_econ_metric', 'revenue')

        def value(y, m):
            r = rev(y, m)
            if not r:
                return 0
            if metric == 'revenue':
                return r['total']
            if metric == 'adr':
                return r['total'] / r['roomn'] if r['roomn'] else 0
            n = nights_of(y, m)
            return r['total'] / n if n else 0
        names = {'revenue': T("Έσοδα (€)", "Revenue (€)"), 'adr': T("Έσοδα ανά διανυκτ. δωματίου (€)", "Revenue per room night (€)"),
                 'pernight': T("Έσοδα ανά διανυκτ. πελάτη (€)", "Revenue per guest night (€)")}
        prev = [value(year - 1, m) for m in range(1, 13)] if any(rev(year - 1, m) for m in range(1, 13)) else None
        self.econ_chart.set_data([value(year, m) for m in range(1, 13)], f"{names[metric]} · {year}",
                                 color={'revenue': ACCENT, 'adr': TEAL}.get(metric, NAVY),
                                 prev=prev, prev_label=str(year - 1))

        # table
        for m in [x for x in months if rev(year, x)]:
            r = rev(year, m)
            n = nights_of(year, m)
            zb = 'alt' if len(self.econ_tree.get_children()) % 2 else 'row'
            old = rev(year - 1, m)
            delta = f"{fmt_dec((r['total'] - old['total']) / old['total'] * 100, signed=True)}%" if old else ""
            self.econ_tree.insert('', 'end', text=month_name(m), tags=(zb,),
                                  values=(fmt_int(n), f"{fmt_int(r['total'])} €", f"{fmt_int(r['rooms'])} €",
                                          f"{fmt_int(r['fb'])} €", f"{fmt_int(r['extras'])} €",
                                          f"{fmt_dec(r['total'] / n)} €" if n else dash, delta))
        hid = set()
        if not any(rev(year, m)['fb'] for m in with_rev):
            hid.add('fb')
        if not any(rev(year, m)['extras'] for m in with_rev):
            hid.add('ex')
        if not nights:
            hid.update(('n', 'pn'))
        if not any(rev(year - 1, m) for m in with_rev):
            hid.add('d')
        if not any(rev(year, m)['rooms'] for m in with_rev):
            hid.add('rm')
        self.econ_tree._hidden_cols = hid
        self.econ_tree._refit()
        self.econ_tree.insert('', 'end', text=T(f"ΣΥΝΟΛΟ {year}", f"TOTAL {year}"), tags=('total',),
                              values=(fmt_int(nights), f"{fmt_int(tot)} €", f"{fmt_int(rooms)} €", f"{fmt_int(fb)} €",
                                      f"{fmt_int(extras)} €", f"{fmt_dec(tot / nights)} €" if nights else dash, ""))

        # by country: guests and nights of the months that have revenue, next to what each country spent
        gn = {}
        for m in with_rev:
            for c, g, n in hist.get(year, {}).get(m, []):
                d = gn.setdefault(c, [0, 0])
                d[0] += g
                d[1] += n
        for c, d in load_revenue_detail(year).items():
            if d['total'] <= 0:
                continue
            g, n = gn.get(c, (0, 0))
            self._econ_rows.append({'c': c, 'g': g, 'n': n, 'rm': d['rooms'], 'fb': d['fb'], 'ex': d['extras'],
                                    'r': d['total'], 'roomn': d['roomn'], 'sh': d['total'] / tot * 100 if tot else 0.0,
                                    'rn': d['total'] / d['roomn'] if d['roomn'] else 0.0,
                                    'pn': d['total'] / n if n else 0.0})
        self._econ_render_countries()
        self._econ_fit_height()

        # short summary: best month, revenue mix, best-earning markets
        parts = []
        best = max(with_rev, key=lambda m: rev(year, m)['total'])
        parts.append(T(f"Κορυφαίος μήνας: {month_name(best)} ({fmt_dec(rev(year, best)['total'] / tot * 100)}% των εσόδων)",
                       f"Best month: {month_name(best)} ({fmt_dec(rev(year, best)['total'] / tot * 100)}% of revenue)"))
        if tot:
            parts.append(T(f"Δωμάτια {fmt_dec(rooms / tot * 100)}% · F&B {fmt_dec(fb / tot * 100)}% · έξτρα {fmt_dec(extras / tot * 100)}%",
                           f"Rooms {fmt_dec(rooms / tot * 100)}% · F&B {fmt_dec(fb / tot * 100)}% · extras {fmt_dec(extras / tot * 100)}%"))
        bc = sorted(load_revenue_by_country(year).items(), key=lambda kv: -kv[1])[:3]
        if bc and tot:
            parts.append(T("Κορυφαίες αγορές σε έσοδα: ", "Top markets by revenue: ") +
                         ", ".join(f"{c} {fmt_dec(v / tot * 100)}%" for c, v in bc))
        if missing:
            parts.append(T(f"Χωρίς έσοδα: {', '.join(month_name(m) for m in missing)}",
                           f"No revenue saved for: {', '.join(month_name(m) for m in missing)}"))
        self.econ_insight.set("   ·   ".join(parts))

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
        hist = self._hist

        def fail(e):
            if isinstance(e, PermissionError):
                messagebox.showerror(T("Σφάλμα", "Error"), T("Δεν μπορώ να γράψω το αρχείο. Κλείστε το αν είναι ανοιχτό και δοκιμάστε ξανά.",
                                                 "I can't write the file. Close it if it is open and try again."))
            else:
                messagebox.showerror(T("Σφάλμα", "Error"), T(f"Δεν ήταν δυνατή η εξαγωγή:\n{e}", f"Could not export:\n{e}"))

        def done(_r):
            self._status(T(f"✓ Το ιστορικό εξήχθη: {fp}", f"✓ History exported: {fp}"))
            open_file(fp)
        self.run_async(lambda: export_history_excel(hist, fp), done, fail, text=T("Εξαγωγή…", "Exporting…"))

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
        self.update_economics()
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
                       "Διαλέξτε το Excel / CSV του μήνα στο Βήμα 1 για να το δημιουργήσετε.",
                       f"Reminder: there is no form yet for {month_name(r[1])} {r[0]}. "
                       "Choose the month's Excel / CSV in Step 1 to create it."))
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

        preview = bool(self.prefs.get('preview_before_result', False))
        # with the preview on, nothing is written to the output file (and nothing is asked about it) until the
        # user presses "Save & continue" in the preview
        if not preview and not self._confirm_replace(out):
            return

        saved = dict(self._raw_settings(), last_dir=os.path.dirname(excel), last_employees=m['employees'])
        try:
            save_hotel_settings(saved)
        except Exception as e:
            messagebox.showwarning(T("Προσοχή", "Warning"), T(f"Δεν ήταν δυνατή η αποθήκευση των στοιχείων στον φάκελο "
                                                f"δεδομένων:\n{DATA_DIR}\n\n{e}",
                                                f"Could not save the details to the data "
                                                f"folder:\n{DATA_DIR}\n\n{e}"))
        self._save_custom_fields_now()
        self.current_settings = saved
        if template != (self.prefs.get('template_pdf') or DEFAULT_TEMPLATE):
            self._remember_template(template)

        mode = self.prefs.get('pdf_mode', 'overlay')
        preview_dir = None
        try:
            if preview:
                preview_dir = tempfile.mkdtemp(prefix='elstat_preview_')
                target = os.path.join(preview_dir, os.path.basename(out) or "ELSTAT_Report.pdf")
            else:
                os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
                target = out
        except Exception as e:
            if preview_dir:
                shutil.rmtree(preview_dir, ignore_errors=True)
            messagebox.showerror(T("Σφάλμα", "Error"), T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"))
            return
        ctx = (target, preview_dir, preview, mode, excel, template, s, m, out, print_after)

        def work():
            return process_elstat_report(excel, template, target, s, m, pdf_mode=mode, save_history=not preview)

        def fail(e):
            self._set_busy(False)
            if preview_dir:
                shutil.rmtree(preview_dir, ignore_errors=True)
            if isinstance(e, PermissionError):
                messagebox.showerror(T("Σφάλμα", "Error"), T("Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό σε άλλο "
                                                 "πρόγραμμα και δοκιμάστε ξανά.",
                                                 "I can't write the PDF. Close it if it is open in another "
                                                 "program and try again."))
            else:
                messagebox.showerror(T("Σφάλμα", "Error"), T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"))

        def done(result):
            self._set_busy(False)
            self._generate_continue(result, ctx)

        self._set_busy(True)
        self.run_async(work, done, fail, text=T("Δημιουργία PDF…", "Creating the PDF…"))

    def _generate_continue(self, result, ctx):
        """Second half of generate(): runs in the main thread once the PDF has been built."""
        target, preview_dir, preview, mode, excel, template, s, m, out, print_after = ctx
        if not preview:
            self._finish_generation(result, out, s, m, template, print_after)
            return

        # ---- preview first: the PDF lives in a temporary folder until the user proceeds
        state = {'mode': mode, 'result': result}
        state = {'mode': mode, 'result': result}

        def regenerate(layout):
            """The preview was edited: build it again with that temporary layout (None = the normal way)."""
            if layout is None:
                new = process_elstat_report(excel, template, target, s, m, pdf_mode=mode, save_history=False)
                state['mode'] = mode
            else:
                new = process_elstat_report(excel, template, target, s, m, pdf_mode='manual',
                                            manual_layout=layout, save_history=False)
                state['mode'] = 'manual'
            state['result'] = new

        def proceed(dialog):
            """"Save & continue": now the file is really saved, then the Result window follows."""
            if not self._confirm_replace(out, parent=dialog):
                return False
            try:
                os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
                shutil.copyfile(target, out)          # the saved file is exactly the one that was previewed
            except PermissionError:
                messagebox.showerror(T("Σφάλμα", "Error"), T("Δεν μπορώ να γράψω το PDF. Κλείστε το αν είναι ανοιχτό σε "
                                                 "άλλο πρόγραμμα και δοκιμάστε ξανά.",
                                                 "I can't write the PDF. Close it if it is open in another "
                                                 "program and try again."), parent=dialog)
                return False
            except Exception as e:
                messagebox.showerror(T("Σφάλμα", "Error"), T(f"Κάτι πήγε στραβά:\n{e}", f"Something went wrong:\n{e}"),
                                     parent=dialog)
                return False
            shutil.rmtree(preview_dir, ignore_errors=True)
            final_result, final_mode = state['result'], state['mode']
            self.after(60, lambda: self._finish_generation(final_result, out, s, m, template, print_after,
                                                           commit_mode=final_mode))
            return True

        def closed():
            shutil.rmtree(preview_dir, ignore_errors=True)
            self._status(T("Η προεπισκόπηση έκλεισε - το αρχείο δεν αποθηκεύτηκε.",
                           "The preview was closed - the file was not saved."))

        def live_data():
            res = state['result']
            return {'values': res.get('live_values'), 'rows': res.get('pdf_data')}

        def base_layout():
            """What "Edit layout…" starts from: the layout this report was made with (or a blank canvas, see Settings)."""
            return self._editor_start_layout(mode, template, state['result'])[0]

        def confidence():
            """Recognition confidence of what the preview shows (None once the layout was edited by hand)."""
            res = state['result']
            return confidence_view(res.get('confidence'), res.get('layout_used'))

        PDFPreviewDialog(self, target, template, proceed, closed, regenerate, live_data, base_layout, confidence)

    def _confirm_replace(self, out, parent=None):
        """The "file already exists" question (when it is turned on in Settings). True = go ahead."""
        if self.prefs['confirm_overwrite'] and os.path.exists(out):
            return messagebox.askyesno(T("Υπάρχει ήδη αρχείο", "File already exists"),
                                       T(f"Το αρχείο υπάρχει ήδη:\n{out}\n\nΝα αντικατασταθεί;",
                                         f"The file already exists:\n{out}\n\nReplace it?"),
                                       default='no', parent=parent or self)
        return True

    @staticmethod
    def _result_warnings(result):
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
        return warnings_txt

    def _finish_generation(self, result, out, s, m, template, print_after, commit_mode=None):
        """Everything that follows a saved PDF: history, archive, printing, opening, the Result window.
        commit_mode is given when the history was not written yet (the PDF came through the preview)."""
        warnings_txt = self._result_warnings(result)
        if commit_mode is not None:
            try:
                commit_report_history(result['pdf_data'], m, template, commit_mode)
            except Exception as e:
                messagebox.showwarning(T("Προσοχή", "Warning"),
                                       T(f"Το PDF αποθηκεύτηκε, αλλά το ιστορικό δεν ενημερώθηκε:\n{e}",
                                         f"The PDF was saved, but the history could not be updated:\n{e}"))
        extra = ("  ·  " + result['layout_info']) if result.get('layout_info') and commit_mode in (None, 'smart', 'detect') else ""
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
        ResultDialog(self, result, out, warnings_txt, print_status)


if __name__ == "__main__":
    if sys.platform.startswith('win'):      # sharp text on high-DPI screens
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    sys.excepthook = lambda e, v, t: _handle_error(e, v, t)
    if hasattr(threading, 'excepthook'):
        threading.excepthook = lambda a: _handle_error(a.exc_type, a.exc_value, a.exc_traceback)
    app = ElstatApp()
    app.mainloop()
