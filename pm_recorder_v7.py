# -*- coding: utf-8 -*-
"""
PM interaction recorder v7 - full screen map with values, frames and colours.

New in v7 (internal tool): every control of every screen is recorded with its
real value (text of labels, fields, lists), the frame (group box such as
Identification, Contact, Rewards) it sits in, its control ID and position.
Controls are identified by ID and frame, so the map does not depend on screen
size. For the player profile the Identification name (control 1034 in frame
3923) is read on every capture together with its text and background colour
(a green background is flagged), and flagged when it contains "(Loc:" - the runner skips such players.
Passwords and the login window are still never recorded.

--- v6 description ---

New in v6: a screen map. Every distinct PM screen (main window page or popup)
is captured once with the position of each visible control (buttons, fields,
lists, labels, tabs ...) relative to its window, and exported as screens.json
plus screens.html, which draws every screen with a box per control (hover a box
for its type, name, AutomationId and coordinates). Screens are captured when a
popup opens or changes, when the main title / active tab changes, when the main
window becomes usable again, and after each user action; identical layouts are
stored once (values such as player data are not part of the layout).
Also fixed from v5: menu shadow windows are ignored, no "popup changed" after a
popup closed, clicks in other programs are ignored, Ctrl/Alt shortcuts are
recorded as keys, password fields are detected through the focused window, and
popups are rescanned after each action inside them.

--- v5 description ---

Everything v4 recorded is still recorded (mouse clicks, scrolls, navigation /
function keys, shortcuts, full UI snapshots after each user action). Typed
characters, usernames and passwords are still never stored.

New in v5: a popup watcher that runs independently of user input.
  * Polls every PM top-level window (Win32 EnumWindows, filtered by PM process
    ID) every POPUP_POLL_SECONDS and is woken instantly by a WinEvent hook
    (show / hide / destroy / dialog start / dialog end / menu start / menu end).
  * Emits popup_opened / popup_changed / popup_closed events with HWND, title,
    class, owner, kind (dialog, menu, dropdown, login, window) and timestamps.
  * Captures what each popup looks like: native Win32 child controls (class,
    control ID, caption, position relative to the popup, enabled/visible), and
    a UIA tree with control state (radio selected, combo selection, ...),
    taken in a separate thread so a busy PM never stalls the watcher.
  * Tracks the main window: title changes (player profile loaded), enabled /
    disabled (modal popup blocking), and "not responding" (IsHungAppWindow).
  * Links every popup to the user action that preceded it and to the action
    that closed it, with delays in milliseconds.

Output folder pm_recording/:
  pm_recording.json  - everything (events, popups, win_events, snapshots, errors)
  popups.json        - popup lifecycle records only
  actions_only.json  - user actions only (same shape as v4)
  timeline.txt       - human-readable timeline + popup catalog
  screens.json       - distinct screens with control positions (v6)
  screens.html       - drawing of every screen, open it in a browser (v6)
  summary.txt        - counts

Press F8 to stop and export.
Requires: pywin32, pynput, pywinauto (and Pillow only if POPUP_SCREENSHOTS=True).
"""
import ctypes
import hashlib
import html
import json
import os
import queue
import re
import threading
import time
import traceback
from collections import Counter, deque
from ctypes import wintypes
from datetime import datetime
from pathlib import Path

import win32con
import win32gui
import win32process
from pynput import keyboard, mouse
from pywinauto import Desktop

RECORDER_VERSION = 7
TITLE_RE = re.compile(r"Patron Management", re.IGNORECASE)
SENSITIVE_RE = re.compile(
    r"(?:user\s*name|username|user\s*id|login|logon|sign\s*in|password|passcode|\bpin\b|admin)",
    re.IGNORECASE,
)
LOGIN_RE = re.compile(r"logon|login|sign\s*in", re.IGNORECASE)
STOP_KEY = keyboard.Key.f8
OUTPUT_DIR = "pm_recording"
SCAN_DEBOUNCE_SECONDS = 0.7
MAX_DEPTH = 40
DOUBLE_CLICK_SECONDS = 0.55

# Popup watcher settings.
POPUP_POLL_SECONDS = 0.1          # fallback polling interval; WinEvents wake it sooner
POPUP_SETTLE_SECONDS = 0.35       # wait before the UIA scan of a new/changed popup
POPUP_MAX_UIA_SCANS = 12          # per popup
FULL_SNAPSHOTS = False            # v4/v5 full-tree snapshot after every action (replaced by the screen map)
SCREEN_SETTLE_SECONDS = 0.8       # wait before capturing the main window after a change
MAX_SCREENS = 300
RECORD_SCREEN_VALUES = True       # internal tool: store real values (set False to store <value>)
IDENT_FRAME_ID = "3923"           # group box "Identification"
IDENT_NAME_ID = "1034"            # player name inside Identification
SKIP_NAME_RE = re.compile(r"\(\s*Loc\s*:", re.IGNORECASE)
VALUE_TYPES = {"Text", "ListItem", "DataItem", "TreeItem", "Document"}
MAP_CATEGORIES = [                # drawn in screens.html: (category, colour, control types)
    ("action", "#2563eb", {"Button", "SplitButton", "MenuItem", "Hyperlink"}),
    ("input", "#16a34a", {"Edit", "ComboBox", "Spinner", "Slider"}),
    ("choice", "#9333ea", {"RadioButton", "CheckBox"}),
    ("list", "#ea580c", {"List", "ListItem", "Tree", "TreeItem", "DataGrid", "DataItem", "Table",
                         "Header", "HeaderItem"}),
    ("tab", "#0d9488", {"Tab", "TabItem"}),
    ("label", "#6b7280", {"Text"}),
    ("group", "#a1a1aa", {"Group"}),
]
PID_REFRESH_SECONDS = 1.0
CLOSE_LINK_SECONDS = 5.0          # click on the popup at most this long before it closed
HOOK_MATCH_SECONDS = 1.5          # a WinEvent this recent refines the open/close time
MAX_WIN32_CONTROLS = 300
MAX_POPUP_CHANGES = 50
MAX_WIN_EVENTS = 20000
MAX_ERRORS = 500
POPUP_SCREENSHOTS = False         # True saves a PNG of each popup (may show player data)
MAIN_CLASS_HINTS = ("XTPMainFrame",)
IGNORED_CLASS_RE = re.compile(r"tooltip|shadow|^IME$|MSCTFIME|CicMarshal", re.IGNORECASE)
INPUT_CLASS_RE = re.compile(
    r"edit|richedit|ipaddress|syslistview|systreeview|listbox|combobox|grid", re.IGNORECASE
)
TITLEBAR_BUTTONS = {"close", "minimize", "maximize", "restore", "help"}
BUTTON_STYLES = {0: "push", 1: "push", 2: "checkbox", 3: "checkbox", 4: "radio", 5: "checkbox",
                 6: "checkbox", 7: "groupbox", 9: "radio", 11: "owner_draw"}

SAFE_KEYS = {
    keyboard.Key.tab: "TAB", keyboard.Key.enter: "ENTER", keyboard.Key.esc: "ESC",
    keyboard.Key.space: "SPACE", keyboard.Key.backspace: "BACKSPACE",
    keyboard.Key.delete: "DELETE", keyboard.Key.home: "HOME", keyboard.Key.end: "END",
    keyboard.Key.page_up: "PAGE_UP", keyboard.Key.page_down: "PAGE_DOWN",
    keyboard.Key.up: "UP", keyboard.Key.down: "DOWN", keyboard.Key.left: "LEFT",
    keyboard.Key.right: "RIGHT", keyboard.Key.insert: "INSERT",
    keyboard.Key.f1: "F1", keyboard.Key.f2: "F2", keyboard.Key.f3: "F3",
    keyboard.Key.f4: "F4", keyboard.Key.f5: "F5", keyboard.Key.f6: "F6",
    keyboard.Key.f7: "F7", keyboard.Key.f9: "F9", keyboard.Key.f10: "F10",
    keyboard.Key.f11: "F11", keyboard.Key.f12: "F12",
}
MODIFIER_KEYS = {
    keyboard.Key.ctrl, keyboard.Key.ctrl_l, keyboard.Key.ctrl_r,
    keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r,
    keyboard.Key.alt, keyboard.Key.alt_l, keyboard.Key.alt_r,
}
USER_ACTION_EVENTS = {
    "mouse_click", "mouse_double_click", "mouse_scroll", "key_press",
    "text_input", "credential_input", "protected_field_click",
}

# WinEvent constants.
EVENT_NAMES = {
    0x0003: "foreground",
    0x0006: "menu_popup_start",
    0x0007: "menu_popup_end",
    0x0010: "dialog_start",
    0x0011: "dialog_end",
    0x8000: "create",
    0x8001: "destroy",
    0x8002: "show",
    0x8003: "hide",
    0x800A: "state_change",
    0x800C: "name_change",
}
HOOK_RANGES = [(0x0003, 0x0003), (0x0006, 0x0007), (0x0010, 0x0011),
               (0x8000, 0x8003), (0x800A, 0x800A), (0x800C, 0x800C)]
OPEN_HOOK_EVENTS = ("show", "dialog_start", "menu_popup_start")
CLOSE_HOOK_EVENTS = ("hide", "destroy", "dialog_end", "menu_popup_end")
OBJID_WINDOW = 0
CHILDID_SELF = 0
GA_ROOT = 2
WINEVENT_OUTOFCONTEXT = 0x0000
WINEVENT_SKIPOWNPROCESS = 0x0002
WM_QUIT = 0x0012

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
WINEVENTPROC = ctypes.WINFUNCTYPE(
    None, wintypes.HANDLE, wintypes.DWORD, wintypes.HWND,
    wintypes.LONG, wintypes.LONG, wintypes.DWORD, wintypes.DWORD,
)
user32.SetWinEventHook.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.HMODULE,
                                   WINEVENTPROC, wintypes.DWORD, wintypes.DWORD, wintypes.DWORD]
user32.SetWinEventHook.restype = wintypes.HANDLE
user32.UnhookWinEvent.argtypes = [wintypes.HANDLE]
user32.UnhookWinEvent.restype = wintypes.BOOL
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = wintypes.BOOL
user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostThreadMessageW.restype = wintypes.BOOL
user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetAncestor.restype = wintypes.HWND
user32.IsHungAppWindow.argtypes = [wintypes.HWND]
user32.IsHungAppWindow.restype = wintypes.BOOL
user32.InternalGetWindowText.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.InternalGetWindowText.restype = ctypes.c_int


class GUITHREADINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("flags", wintypes.DWORD),
                ("hwndActive", wintypes.HWND), ("hwndFocus", wintypes.HWND),
                ("hwndCapture", wintypes.HWND), ("hwndMenuOwner", wintypes.HWND),
                ("hwndMoveSize", wintypes.HWND), ("hwndCaret", wintypes.HWND),
                ("rcCaret", wintypes.RECT)]


user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.POINTER(GUITHREADINFO)]
user32.GetGUIThreadInfo.restype = wintypes.BOOL
user32.SendMessageTimeoutW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
                                       wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t)]
user32.SendMessageTimeoutW.restype = ctypes.c_size_t
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
for _fn, _args, _res in (
        (user32.GetWindowDC, [wintypes.HWND], ctypes.c_void_p),
        (user32.ReleaseDC, [wintypes.HWND, ctypes.c_void_p], ctypes.c_int),
        (user32.PrintWindow, [wintypes.HWND, ctypes.c_void_p, wintypes.UINT], wintypes.BOOL),
        (gdi32.CreateCompatibleDC, [ctypes.c_void_p], ctypes.c_void_p),
        (gdi32.CreateCompatibleBitmap, [ctypes.c_void_p, ctypes.c_int, ctypes.c_int], ctypes.c_void_p),
        (gdi32.SelectObject, [ctypes.c_void_p, ctypes.c_void_p], ctypes.c_void_p),
        (gdi32.DeleteObject, [ctypes.c_void_p], wintypes.BOOL),
        (gdi32.DeleteDC, [ctypes.c_void_p], wintypes.BOOL),
        (gdi32.GetDIBits, [ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT, wintypes.UINT,
                           ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT], ctypes.c_int)):
    _fn.argtypes, _fn.restype = _args, _res


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", ctypes.c_long), ("biHeight", ctypes.c_long),
                ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
                ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
                ("biXPelsPerMeter", ctypes.c_long), ("biYPelsPerMeter", ctypes.c_long),
                ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]


def wm_gettext(hwnd, limit=2048):
    """Text of a native control (also Edit contents); never blocks on a hung PM."""
    buf = ctypes.create_unicode_buffer(limit)
    result = ctypes.c_size_t()
    ok = user32.SendMessageTimeoutW(hwnd, 0x000D, limit, ctypes.addressof(buf), 0x0002, 500,
                                    ctypes.byref(result))
    return buf.value.replace("\r", " ").replace("\n", " ").strip() if ok else ""


def capture_window(hwnd):
    """Render a window into memory with PrintWindow (works when covered, not when minimized).

    Returns (width, height, BGRA bytes) or None.
    """
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0 or win32gui.IsIconic(hwnd):
        return None
    window_dc = user32.GetWindowDC(hwnd)
    memory_dc = gdi32.CreateCompatibleDC(window_dc)
    bitmap = gdi32.CreateCompatibleBitmap(window_dc, width, height)
    previous = gdi32.SelectObject(memory_dc, bitmap)
    try:
        if not user32.PrintWindow(hwnd, memory_dc, 2):        # PW_RENDERFULLCONTENT
            user32.PrintWindow(hwnd, memory_dc, 0)
        header = BITMAPINFOHEADER(biSize=ctypes.sizeof(BITMAPINFOHEADER), biWidth=width,
                                  biHeight=-height, biPlanes=1, biBitCount=32, biCompression=0)
        buf = ctypes.create_string_buffer(width * height * 4)
        gdi32.GetDIBits(memory_dc, bitmap, 0, height, buf, ctypes.byref(header), 0)
        return width, height, buf.raw
    finally:
        gdi32.SelectObject(memory_dc, previous)
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(memory_dc)
        user32.ReleaseDC(hwnd, window_dc)


def color_name(rgb):
    r, g, b = rgb
    if max(rgb) - min(rgb) < 30:
        return "white" if min(rgb) > 200 else "black" if max(rgb) < 70 else "gray"
    if g >= r + 30 and g >= b + 30:
        return "green"
    if b >= r + 30 and b >= g - 10:
        return "blue"
    if r >= g + 30 and r >= b + 30:
        return "red"
    if r > 150 and g > 150 and b < 110:
        return "yellow"
    return "other"


def region_colors(image, x0, y0, x1, y1):
    """Background (most common) and text (most common clearly different) colour of a region."""
    width, height, raw = image
    counts = Counter()
    for y in range(max(0, y0), min(height, y1)):
        row = y * width * 4
        for x in range(max(0, x0), min(width, x1)):
            i = row + x * 4
            counts[(raw[i + 2] & 0xF8, raw[i + 1] & 0xF8, raw[i] & 0xF8)] += 1
    if not counts:
        return None
    background = counts.most_common(1)[0][0]
    text = next((c for c, _ in counts.most_common()
                 if sum(abs(a - b) for a, b in zip(c, background)) > 90), None)
    hexed = lambda c: "#%02x%02x%02x" % c if c else None
    return {"background_color": hexed(background), "background_color_name": color_name(background),
            "text_color": hexed(text), "text_color_name": color_name(text) if text else None}
kernel32.GetCurrentThreadId.restype = wintypes.DWORD
kernel32.GetTickCount.restype = wintypes.DWORD

STARTED_AT = datetime.now()
MONO_START = time.perf_counter()
EVENTS, SNAPSHOTS, ERRORS, WIN_EVENTS, POPUPS = [], [], [], [], []
POPUP_BY_ID = {}
ACTION_HISTORY = deque(maxlen=500)
HOOK_TIMES = {}          # hwnd -> {win_event_name: (iso, t)}
KNOWN_HWND_PID = {}      # remembers PIDs so destroy events can still be attributed
PM_PIDS = set()
PM_MAIN_HWND = None
LAST_UI_HASH = None
LAST_CLICK = {"time": 0.0, "button": None, "x": None, "y": None}
TEXT_RUN_ACTIVE = False
PRESSED_MODIFIERS = set()
STOP_EVENT = threading.Event()
SCAN_EVENT = threading.Event()
WAKE_EVENT = threading.Event()
DATA_LOCK = threading.RLock()
UIA_JOBS = queue.Queue()
HOOK_STATUS = {"installed": 0, "error": None}
SCREENSHOT_DIR = None
SCREENS = []             # distinct screen layouts
SCREEN_BY_SIG = {}
MAIN_SCREEN_REQUEST = {"due": None, "reasons": []}
TRACKER = None


# ---------------------------------------------------------------- helpers

def now_iso():
    return datetime.now().isoformat(timespec="milliseconds")


def now_t():
    return round(time.perf_counter() - MONO_START, 3)


def stamp():
    """Wall-clock ISO time plus seconds since recorder start (monotonic)."""
    return now_iso(), now_t()


def safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def clean(value):
    if value is None:
        return ""
    return str(value).replace("\x00", "").replace("\r", " ").replace("\n", " ").strip()


def is_sensitive_text(*values):
    return any(SENSITIVE_RE.search(clean(v)) for v in values if clean(v))


def rect_from_tuple(rect):
    try:
        left, top, right, bottom = map(int, rect)
        return {"left": left, "top": top, "right": right, "bottom": bottom,
                "width": max(0, right-left), "height": max(0, bottom-top)}
    except Exception:
        return {k: None for k in ("left", "top", "right", "bottom", "width", "height")}


def rect_from_uia(rect):
    try:
        return rect_from_tuple((rect.left, rect.top, rect.right, rect.bottom))
    except Exception:
        return rect_from_tuple(None)


def record_error(where, exc):
    with DATA_LOCK:
        if len(ERRORS) < MAX_ERRORS:
            iso, t = stamp()
            ERRORS.append({"time": iso, "t": t, "where": where, "error": str(exc),
                           "traceback": traceback.format_exc()})


def append_event(event):
    if "time" not in event or "t" not in event:
        iso, t = stamp()
        event.setdefault("time", iso)
        event.setdefault("t", t)
    with DATA_LOCK:
        EVENTS.append(event)


def hwnd_pid(hwnd):
    return safe(lambda: win32process.GetWindowThreadProcessId(hwnd)[1], None)


def internal_text(hwnd):
    """Read a window caption without sending messages (never blocks on a busy PM)."""
    try:
        buf = ctypes.create_unicode_buffer(512)
        user32.InternalGetWindowText(hwnd, buf, 512)
        return clean(buf.value)
    except Exception:
        return clean(safe(lambda: win32gui.GetWindowText(hwnd), ""))


def is_hung(hwnd):
    return bool(safe(lambda: user32.IsHungAppWindow(hwnd), False)) if hwnd else False


def player_ids_in(title):
    return [x.lstrip("0") or "0" for x in re.findall(r"\((\d+)\)", title or "")]


# ------------------------------------------------------- PM window discovery

def discover_pm_windows():
    global PM_MAIN_HWND, PM_PIDS
    found = []

    def callback(hwnd, _):
        title = internal_text(hwnd)
        if title and TITLE_RE.search(title):
            rect = rect_from_tuple(safe(lambda: win32gui.GetWindowRect(hwnd), None))
            found.append({"hwnd": hwnd, "pid": hwnd_pid(hwnd), "title": title,
                          "class_name": clean(safe(lambda: win32gui.GetClassName(hwnd), "")),
                          "visible": bool(safe(lambda: win32gui.IsWindowVisible(hwnd), False)),
                          "iconic": bool(safe(lambda: win32gui.IsIconic(hwnd), False)),
                          "rectangle": rect})
        return True

    safe(lambda: win32gui.EnumWindows(callback, None))
    PM_PIDS = {x["pid"] for x in found if x["pid"] is not None}

    # Keep the current main window while it is still a valid PM window.
    if PM_MAIN_HWND and any(x["hwnd"] == PM_MAIN_HWND for x in found):
        return found
    hinted = [x for x in found if x["class_name"] in MAIN_CLASS_HINTS]
    pool = hinted or [x for x in found if x["visible"] and not x["iconic"]] or found
    PM_MAIN_HWND = max(
        pool, key=lambda x: (x["rectangle"]["width"] or 0)*(x["rectangle"]["height"] or 0)
    )["hwnd"] if pool else None
    return found


def foreground_context():
    hwnd = safe(win32gui.GetForegroundWindow, 0)
    if not hwnd:
        return None
    pid = hwnd_pid(hwnd)
    if pid not in PM_PIDS:
        discover_pm_windows()
        if pid not in PM_PIDS:
            return None
    title = internal_text(hwnd)
    login_screen = bool(LOGIN_RE.search(title))
    safe_title = "Patron Management Login" if login_screen else title
    return {"hwnd": hwnd, "pid": pid, "title": safe_title, "login_screen": login_screen,
            "class_name": clean(safe(lambda: win32gui.GetClassName(hwnd), "")),
            "rectangle": rect_from_tuple(safe(lambda: win32gui.GetWindowRect(hwnd), None))}


def control_at_point(x, y):
    try:
        ctl = Desktop(backend="uia").from_point(x, y)
        info = ctl.element_info
        name = clean(safe(lambda: info.name, ""))
        aid = clean(safe(lambda: info.automation_id, ""))
        cls = clean(safe(lambda: info.class_name, ""))
        if is_sensitive_text(name, aid, cls):
            return {"redacted": True, "control_type": clean(safe(lambda: info.control_type, ""))}
        return {"name": name, "control_type": clean(safe(lambda: info.control_type, "")),
                "automation_id": aid, "class_name": cls,
                "framework_id": clean(safe(lambda: info.framework_id, "")),
                "process_id": safe(lambda: info.process_id, None),
                "handle": safe(lambda: ctl.handle, None),
                "rectangle": rect_from_uia(safe(lambda: info.rectangle, None))}
    except Exception as exc:
        return {"lookup_error": str(exc)}


def sensitive_control_focused():
    """True when the keyboard focus is a password field or inside a login window."""
    try:
        info = GUITHREADINFO(cbSize=ctypes.sizeof(GUITHREADINFO))
        if not user32.GetGUIThreadInfo(0, ctypes.byref(info)) or not info.hwndFocus:
            return False
        focus = info.hwndFocus
        cls = clean(win32gui.GetClassName(focus)).lower()
        if "edit" in cls and win32gui.GetWindowLong(focus, win32con.GWL_STYLE) & win32con.ES_PASSWORD:
            return True
        root = user32.GetAncestor(focus, GA_ROOT) or focus
        return bool(LOGIN_RE.search(internal_text(root)))
    except Exception:
        return False


# ------------------------------------------------------------ UIA scanning

def root_uia_controls():
    roots, seen = [], set()
    desktop = Desktop(backend="uia")
    for pid in list(PM_PIDS):
        for win in safe(lambda: desktop.windows(process=pid), []) or []:
            key = (safe(lambda: win.handle, None), clean(safe(lambda: win.window_text(), "")))
            if key not in seen and not is_sensitive_text(key[1]):
                seen.add(key); roots.append(win)
    return roots


def control_state(control, ctype):
    """Visible state of selection-type controls. Edit values are never read."""
    state = {}
    if ctype == "RadioButton":
        state["selected"] = safe(lambda: bool(control.is_selected()))
    elif ctype == "CheckBox":
        state["toggle_state"] = safe(lambda: control.get_toggle_state())
    elif ctype == "ComboBox":
        text = clean(safe(lambda: control.selected_text(), ""))
        state["selected_text"] = "<REDACTED>" if is_sensitive_text(text) else text
    elif ctype in ("ListItem", "TabItem", "TreeItem"):
        state["selected"] = safe(lambda: bool(control.is_selected()))
    return {k: v for k, v in state.items() if v not in (None, "")}


def scan_one_tree(root, root_number, with_state=False):
    controls, visited = [], set()

    def walk(control, parent_index=None, depth=0):
        if depth > MAX_DEPTH:
            return
        info = control.element_info
        runtime_id = safe(lambda: tuple(info.runtime_id), None)
        handle = safe(lambda: control.handle, None)
        name = clean(safe(lambda: info.name, "") or safe(lambda: control.window_text(), ""))
        ctype = clean(safe(lambda: info.control_type, ""))
        aid = clean(safe(lambda: info.automation_id, ""))
        cls = clean(safe(lambda: info.class_name, ""))
        identity = (runtime_id, handle, name, ctype)
        if runtime_id is not None and identity in visited:
            return
        visited.add(identity)
        if is_sensitive_text(name, aid, cls) or LOGIN_RE.search(name):
            return
        index = len(controls)
        entry = {"index": index, "parent_index": parent_index, "depth": depth,
                 "root_number": root_number, "name": name, "control_type": ctype,
                 "automation_id": aid, "class_name": cls,
                 "rectangle": rect_from_uia(safe(lambda: info.rectangle, None))}
        if with_state:
            entry["handle"] = handle
            entry["enabled"] = safe(lambda: bool(info.enabled))
            entry["visible"] = safe(lambda: bool(info.visible))
            if handle and ctype in ("Edit", "Text", "ComboBox", "Button", "Group", "RadioButton", "CheckBox"):
                style = safe(lambda: win32gui.GetWindowLong(handle, win32con.GWL_STYLE), 0) or 0
                if "edit" in cls.lower() and style & win32con.ES_PASSWORD:
                    entry["value"] = "<password>"
                else:
                    entry["value"] = wm_gettext(handle)
            state = control_state(control, ctype)
            if state:
                entry["state"] = state
        controls.append(entry)
        for child in safe(lambda: control.children(), []) or []:
            walk(child, index, depth+1)
    walk(root)
    return controls


def request_scan():
    SCAN_EVENT.set()


def scan_worker():
    """Full PM UI snapshot after user actions (same as v4, plus scan timing)."""
    global LAST_UI_HASH
    while not STOP_EVENT.is_set():
        if not SCAN_EVENT.wait(0.2):
            continue
        if not FULL_SNAPSHOTS:
            SCAN_EVENT.clear()
            continue
        SCAN_EVENT.clear(); time.sleep(SCAN_DEBOUNCE_SECONDS)
        try:
            started_iso, started_t = stamp()
            discover_pm_windows()
            roots_payload, all_controls = [], []
            for n, root in enumerate(root_uia_controls()):
                controls = scan_one_tree(root, n)
                roots_payload.append({"root_number": n,
                    "title": clean(safe(lambda: root.window_text(), "")),
                    "handle": safe(lambda: root.handle, None), "control_count": len(controls)})
                all_controls.extend(controls)
            canonical = json.dumps({"roots": roots_payload, "controls": all_controls}, ensure_ascii=False, sort_keys=True)
            digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
            if digest == LAST_UI_HASH:
                continue
            LAST_UI_HASH = digest
            iso, t = stamp()
            with DATA_LOCK:
                sid = len(SNAPSHOTS)
                SNAPSHOTS.append({"snapshot_id": sid, "time": iso, "t": t,
                                  "scan_started_at": started_iso, "scan_started_t": started_t,
                                  "sha256": digest, "roots": roots_payload,
                                  "control_count": len(all_controls), "controls": all_controls})
                EVENTS.append({"time": iso, "t": t, "event": "ui_snapshot", "snapshot_id": sid,
                               "scan_seconds": round(t - started_t, 3)})
        except Exception as exc:
            record_error("scan_worker", exc)


# ------------------------------------------------------------- user actions

def remember_action(event, label):
    fg = event.get("foreground") or {}
    with DATA_LOCK:
        ACTION_HISTORY.append({"t": event["t"], "time": event["time"], "event": event["event"],
                               "label": label, "fg_hwnd": fg.get("hwnd"),
                               "fg_title": fg.get("title", "")})


def target_label(target):
    if not target or target.get("redacted"):
        return "<protected field>"
    if target.get("lookup_error"):
        return "<unknown control>"
    name = target.get("name") or target.get("automation_id") or ""
    return f"\"{name}\" [{target.get('control_type', '')}]"


def on_click(x, y, button, pressed):
    global TEXT_RUN_ACTIVE, LAST_CLICK
    if STOP_EVENT.is_set() or not pressed:
        return
    iso, t = stamp()   # press time, before the UIA lookup below
    fg = foreground_context()
    if fg is None:
        return
    target = control_at_point(x, y)
    if target.get("process_id") is not None and target.get("process_id") not in PM_PIDS:
        return   # click landed in another program (e.g. on a second monitor)
    if target.get("redacted"):
        TEXT_RUN_ACTIVE = False
        event = {"time": iso, "t": t, "event": "protected_field_click",
                 "button": str(button).replace("Button.", ""), "foreground": fg,
                 "target": {"redacted": True, "control_type": target.get("control_type", "")}}
        append_event(event)
        remember_action(event, "CLICK <protected field>")
        return
    TEXT_RUN_ACTIVE = False
    mono = time.monotonic(); b = str(button).replace("Button.", "")
    double = b == LAST_CLICK["button"] and x == LAST_CLICK["x"] and y == LAST_CLICK["y"] and mono-LAST_CLICK["time"] <= DOUBLE_CLICK_SECONDS
    LAST_CLICK = {"time": mono, "button": b, "x": x, "y": y}
    event = {"time": iso, "t": t, "event": "mouse_double_click" if double else "mouse_click",
             "button": b, "screen_x": x, "screen_y": y, "foreground": fg, "target": target}
    append_event(event)
    remember_action(event, f"{'DOUBLE-CLICK' if double else 'CLICK'} {target_label(target)}")
    request_scan()
    after_user_action(fg)
    WAKE_EVENT.set()


def on_scroll(x, y, dx, dy):
    global TEXT_RUN_ACTIVE
    iso, t = stamp()
    fg = foreground_context()
    if STOP_EVENT.is_set() or fg is None:
        return
    target = control_at_point(x, y)
    if target.get("redacted"):
        return
    TEXT_RUN_ACTIVE = False
    event = {"time": iso, "t": t, "event": "mouse_scroll", "screen_x": x, "screen_y": y,
             "dx": dx, "dy": dy, "foreground": fg, "target": target}
    append_event(event)
    remember_action(event, f"SCROLL {target_label(target)}")
    request_scan()


def modifier_name(key):
    return str(key).replace("Key.", "").upper()


def on_press(key):
    global TEXT_RUN_ACTIVE
    if key == STOP_KEY:
        STOP_EVENT.set(); return False
    iso, t = stamp()
    fg = foreground_context()
    if STOP_EVENT.is_set() or fg is None:
        return
    if fg.get("login_screen") or sensitive_control_focused():
        # Record only that an interaction occurred in a protected field.
        # Never store key names, characters, modifiers, username, or password.
        if not TEXT_RUN_ACTIVE:
            TEXT_RUN_ACTIVE = True
            event = {"time": iso, "t": t, "event": "credential_input",
                     "value": "<REDACTED>", "foreground": fg,
                     "note": "Login typing detected; content intentionally not stored."}
            append_event(event)
            remember_action(event, "TYPE <credentials>")
        return
    if key in MODIFIER_KEYS:
        PRESSED_MODIFIERS.add(modifier_name(key)); return
    if key in SAFE_KEYS:
        TEXT_RUN_ACTIVE = False
        mods = sorted(PRESSED_MODIFIERS)
        event = {"time": iso, "t": t, "event": "key_press", "key": SAFE_KEYS[key],
                 "modifiers": mods, "foreground": fg}
        append_event(event)
        remember_action(event, "KEY " + "+".join(mods + [SAFE_KEYS[key]]))
        request_scan(); after_user_action(fg); WAKE_EVENT.set(); return
    command_mods = sorted(m for m in PRESSED_MODIFIERS if m.startswith(("CTRL", "ALT")))
    vk = safe(lambda: key.vk, None)
    if command_mods and vk and (65 <= vk <= 90 or 48 <= vk <= 57):
        # Ctrl/Alt + letter is a command (e.g. Ctrl+F), not text.
        TEXT_RUN_ACTIVE = False
        mods = sorted(PRESSED_MODIFIERS)
        event = {"time": iso, "t": t, "event": "key_press", "key": chr(vk),
                 "modifiers": mods, "foreground": fg, "shortcut": True}
        append_event(event)
        remember_action(event, "KEY " + "+".join(mods + [chr(vk)]))
        request_scan(); after_user_action(fg); WAKE_EVENT.set(); return
    char = safe(lambda: key.char, None)
    if char is not None and not TEXT_RUN_ACTIVE:
        TEXT_RUN_ACTIVE = True
        event = {"time": iso, "t": t, "event": "text_input", "value": "<TEXT_INPUT>",
                 "modifiers": sorted(PRESSED_MODIFIERS), "foreground": fg,
                 "note": "Characters intentionally not stored."}
        append_event(event)
        remember_action(event, "TYPE <text>")


def on_release(key):
    if key in MODIFIER_KEYS:
        PRESSED_MODIFIERS.discard(modifier_name(key))
    if key == STOP_KEY or STOP_EVENT.is_set():
        return False


def last_action_before(t, fg_hwnd=None, max_age=None):
    with DATA_LOCK:
        items = list(ACTION_HISTORY)
    for action in reversed(items):
        if action["t"] > t:
            continue
        if max_age is not None and t - action["t"] > max_age:
            return None
        if fg_hwnd is not None and action["fg_hwnd"] != fg_hwnd:
            continue
        return action
    return None


def action_ref(action, t):
    if not action:
        return None
    return {"action_time": action["time"], "event": action["event"], "label": action["label"],
            "in_window": action["fg_title"], "delay_ms": int(round((t - action["t"]) * 1000))}


# ---------------------------------------------------------- WinEvent hook

def consume_hook_time(hwnd, names, t):
    """Earliest matching WinEvent for hwnd within HOOK_MATCH_SECONDS before t."""
    with DATA_LOCK:
        entry = HOOK_TIMES.get(hwnd) or {}
        hits = [entry[n] for n in names
                if n in entry and t - HOOK_MATCH_SECONDS <= entry[n][1] <= t + 0.05]
        for n in names:
            entry.pop(n, None)
    return min(hits, key=lambda x: x[1]) if hits else None


class WinEventHook(threading.Thread):
    """Receives window show/hide/destroy/dialog/menu events for PM top-level windows."""

    def __init__(self):
        super().__init__(daemon=True)
        self.thread_id = None
        self.hooks = []
        self.ready = threading.Event()
        self._proc = WINEVENTPROC(self._callback)   # keep a reference: ctypes callback

    def run(self):
        try:
            self.thread_id = kernel32.GetCurrentThreadId()
            for low, high in HOOK_RANGES:
                handle = user32.SetWinEventHook(low, high, None, self._proc, 0, 0,
                                                WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS)
                if handle:
                    self.hooks.append(handle)
            HOOK_STATUS["installed"] = len(self.hooks)
        except Exception as exc:
            HOOK_STATUS["error"] = str(exc)
            record_error("win_event_hook", exc)
        finally:
            self.ready.set()
        try:
            msg = wintypes.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        finally:
            for handle in self.hooks:
                safe(lambda: user32.UnhookWinEvent(handle))

    def stop(self):
        if self.thread_id:
            safe(lambda: user32.PostThreadMessageW(self.thread_id, WM_QUIT, 0, 0))

    def _callback(self, _hook, event, hwnd, id_object, id_child, _thread, event_ms):
        try:
            if not hwnd or STOP_EVENT.is_set():
                return
            name = EVENT_NAMES.get(event)
            if name is None:
                return
            if event >= 0x8000 and (id_object != OBJID_WINDOW or id_child != CHILDID_SELF):
                return
            pid = hwnd_pid(hwnd) or KNOWN_HWND_PID.get(hwnd)
            if pid is None or pid not in PM_PIDS:
                return
            root = safe(lambda: user32.GetAncestor(hwnd, GA_ROOT), None)
            if root and root != hwnd:
                return
            cls = clean(safe(lambda: win32gui.GetClassName(hwnd), ""))
            if IGNORED_CLASS_RE.search(cls):
                return
            iso, t = stamp()
            lag = safe(lambda: (kernel32.GetTickCount() - event_ms) & 0xFFFFFFFF, 0) or 0
            if 0 < lag < 10000:
                t = round(t - lag / 1000.0, 3)
            record = {"time": iso, "t": t, "event": name, "hwnd": hwnd,
                      "title": internal_text(hwnd), "class_name": cls}
            with DATA_LOCK:
                if len(WIN_EVENTS) < MAX_WIN_EVENTS:
                    WIN_EVENTS.append(record)
                HOOK_TIMES.setdefault(hwnd, {})[name] = (iso, t)
            WAKE_EVENT.set()
        except Exception:
            pass


# ------------------------------------------------------------ popup watcher

def enum_visible_toplevel(pids):
    out = []

    def callback(hwnd, _):
        try:
            if win32gui.IsWindowVisible(hwnd) and hwnd_pid(hwnd) in pids:
                out.append(hwnd)
        except Exception:
            pass
        return True

    safe(lambda: win32gui.EnumWindows(callback, None))
    return out


def window_props(hwnd):
    title = internal_text(hwnd)
    owner = safe(lambda: win32gui.GetWindow(hwnd, win32con.GW_OWNER), 0) or 0
    return {
        "hwnd": hwnd,
        "pid": hwnd_pid(hwnd),
        "title": title,
        "class_name": clean(safe(lambda: win32gui.GetClassName(hwnd), "")),
        "rectangle": rect_from_tuple(safe(lambda: win32gui.GetWindowRect(hwnd), None)),
        "enabled": bool(safe(lambda: win32gui.IsWindowEnabled(hwnd), True)),
        "owner_hwnd": owner,
        "owner_title": internal_text(owner) if owner else "",
        "sensitive": bool(LOGIN_RE.search(title)),
    }


def popup_kind(props):
    cls = props["class_name"].lower()
    if props["sensitive"]:
        return "login"
    if cls == "#32770":
        return "dialog"
    if cls == "#32768" or "popupbar" in cls or "menu" in cls:
        return "menu"
    if cls == "combolbox":
        return "dropdown"
    if cls.startswith("windowsforms"):
        return "winforms_window"
    return "window"


def is_ignored_window(props):
    rect = props["rectangle"]
    if not rect["width"] or not rect["height"] or rect["width"] <= 8 or rect["height"] <= 8:
        return True   # XTP menu shadows are 4px strips
    return bool(IGNORED_CLASS_RE.search(props["class_name"]))


def win32_controls(hwnd, redact):
    """Native child controls with position relative to the popup."""
    items = []
    parent_rect = safe(lambda: win32gui.GetWindowRect(hwnd), None)

    def callback(child, _):
        if len(items) >= MAX_WIN32_CONTROLS:
            return True
        try:
            cls = clean(win32gui.GetClassName(child))
            r = win32gui.GetWindowRect(child)
            item = {
                "hwnd": child,
                "parent_hwnd": safe(lambda: win32gui.GetParent(child), 0),
                "control_id": safe(lambda: win32gui.GetDlgCtrlID(child), None),
                "class_name": cls,
                "visible": bool(safe(lambda: win32gui.IsWindowVisible(child), False)),
                "enabled": bool(safe(lambda: win32gui.IsWindowEnabled(child), False)),
                "rectangle": rect_from_tuple(r),
            }
            if cls.lower() == "button":
                style = safe(lambda: win32gui.GetWindowLong(child, win32con.GWL_STYLE), 0) or 0
                item["button_style"] = BUTTON_STYLES.get(style & 0x0F, "push")
            if parent_rect:
                item["rel"] = {"x": r[0]-parent_rect[0], "y": r[1]-parent_rect[1],
                               "w": r[2]-r[0], "h": r[3]-r[1]}
            if redact or INPUT_CLASS_RE.search(cls):
                item["text"] = None
                item["text_policy"] = "not_recorded"
            else:
                text = internal_text(child)
                item["text"] = "<REDACTED>" if is_sensitive_text(text) else text
            items.append(item)
        except Exception:
            pass
        return True

    safe(lambda: win32gui.EnumChildWindows(hwnd, callback, None))
    return items


def controls_signature(controls):
    return {c["hwnd"]: (c["control_id"], c["class_name"], c["text"], c["visible"], c["enabled"])
            for c in controls}


def diff_controls(old, new, controls):
    by_hwnd = {c["hwnd"]: c for c in controls}
    changes = []
    for h in new.keys() - old.keys():
        c = by_hwnd.get(h, {})
        changes.append({"type": "control_added", "control_id": c.get("control_id"),
                        "class_name": c.get("class_name"), "text": c.get("text")})
    for h in old.keys() - new.keys():
        cid, cls, text, _, _ = old[h]
        changes.append({"type": "control_removed", "control_id": cid, "class_name": cls, "text": text})
    for h in old.keys() & new.keys():
        if old[h] != new[h]:
            fields = ("control_id", "class_name", "text", "visible", "enabled")
            delta = {f: {"from": a, "to": b} for f, a, b in zip(fields, old[h], new[h]) if a != b}
            changes.append({"type": "control_changed", "control_id": new[h][0],
                            "class_name": new[h][1], "changes": delta})
    return changes


def schedule_uia_scan(popup_id, hwnd, reason, delay=POPUP_SETTLE_SECONDS):
    UIA_JOBS.put((now_t() + delay, popup_id, hwnd, reason, 0))


class PopupTracker:
    """Diffs PM top-level windows on every poll and emits lifecycle events."""

    def __init__(self):
        self.open = {}         # hwnd -> popup record
        self.signatures = {}   # hwnd -> (title, enabled, controls signature)
        self.main = {"hwnd": None, "title": None, "enabled": None, "hung": None}
        self.foreground = None
        self.next_id = 1
        self.last_pid_refresh = -1e9

    def poll(self):
        iso, t = stamp()
        if not PM_PIDS or t - self.last_pid_refresh >= PID_REFRESH_SECONDS:
            discover_pm_windows()
            self.last_pid_refresh = t
        pids, main = set(PM_PIDS), PM_MAIN_HWND
        if not pids or not main:
            return
        seen = {}
        for hwnd in enum_visible_toplevel(pids):
            if hwnd == main:
                continue
            props = window_props(hwnd)
            if not is_ignored_window(props):
                seen[hwnd] = props
        for hwnd, props in seen.items():
            KNOWN_HWND_PID[hwnd] = props["pid"]
            controls = win32_controls(hwnd, redact=props["sensitive"])
            if not safe(lambda: win32gui.IsWindowVisible(hwnd), False):
                continue   # hidden while being read: the next poll reports the close
            if hwnd in self.open:
                self._check_changed(hwnd, props, controls, iso, t)
            else:
                self._opened(hwnd, props, controls, iso, t)
        for hwnd in [h for h in self.open if h not in seen]:
            self._closed(hwnd, iso, t)
        self._check_main(main, iso, t)
        self._check_foreground(pids, iso, t)

    # -- popup lifecycle

    def _opened(self, hwnd, props, controls, iso, t):
        global TEXT_RUN_ACTIVE
        TEXT_RUN_ACTIVE = False
        hook = consume_hook_time(hwnd, OPEN_HOOK_EVENTS, t)
        opened_iso, opened_t = hook if hook else (iso, t)
        kind = popup_kind(props)
        main_rect = rect_from_tuple(safe(lambda: win32gui.GetWindowRect(PM_MAIN_HWND), None))
        rect = props["rectangle"]
        trigger = last_action_before(opened_t)
        record = {
            "popup_id": self.next_id,
            "hwnd": hwnd,
            "pid": props["pid"],
            "kind": kind,
            "title": props["title"],
            "class_name": props["class_name"],
            "owner_hwnd": props["owner_hwnd"],
            "owner_title": props["owner_title"],
            "owner_is_main": props["owner_hwnd"] == PM_MAIN_HWND,
            "sensitive": props["sensitive"],
            "opened_at": opened_iso,
            "opened_t": opened_t,
            "open_time_source": "win_event" if hook else "poll",
            "detected_at": iso,
            "detected_t": t,
            "rectangle": rect,
            "position_in_main": ({"x": rect["left"] - main_rect["left"], "y": rect["top"] - main_rect["top"]}
                                 if main_rect["left"] is not None and rect["left"] is not None else None),
            "enabled_at_open": props["enabled"],
            "main_title_at_open": self.main.get("title"),
            "main_enabled_at_open": self.main.get("enabled"),
            "main_hung_at_open": self.main.get("hung"),
            "modal": False,
            "opened_after_action": action_ref(trigger, opened_t),
            "controls_win32": controls,
            "uia_snapshots": [],
            "screenshot": None,
            "changes": [],
            "closed_at": None,
            "closed_t": None,
            "close_time_source": None,
            "duration_s": None,
            "close_reason": None,
            "closed_by_action": None,
            "last_action_before_close": None,
            "last_rectangle": rect,
            "still_open_at_stop": False,
        }
        self.next_id += 1
        with DATA_LOCK:
            POPUPS.append(record)
            POPUP_BY_ID[record["popup_id"]] = record
        self.open[hwnd] = record
        self.signatures[hwnd] = (props["title"], props["enabled"], controls_signature(controls))
        append_event({"time": opened_iso, "t": opened_t, "event": "popup_opened",
                      "popup_id": record["popup_id"], "hwnd": hwnd, "kind": kind,
                      "title": props["title"], "class_name": props["class_name"],
                      "owner_hwnd": props["owner_hwnd"],
                      "opened_after_action": record["opened_after_action"]})
        if not props["sensitive"]:
            schedule_uia_scan(record["popup_id"], hwnd, "opened")

    def _check_changed(self, hwnd, props, controls, iso, t):
        record = self.open[hwnd]
        record["last_rectangle"] = props["rectangle"]
        old_title, old_enabled, old_sig = self.signatures[hwnd]
        new_sig = controls_signature(controls)
        changes = []
        if props["title"] != old_title:
            changes.append({"type": "title", "from": old_title, "to": props["title"]})
        if props["enabled"] != old_enabled:
            changes.append({"type": "enabled", "from": old_enabled, "to": props["enabled"]})
        if new_sig != old_sig:
            changes.extend(diff_controls(old_sig, new_sig, controls))
        if not changes:
            return
        self.signatures[hwnd] = (props["title"], props["enabled"], new_sig)
        with DATA_LOCK:
            record["title"] = props["title"]
            record["controls_win32_latest"] = controls
            if len(record["changes"]) < MAX_POPUP_CHANGES:
                record["changes"].append({"time": iso, "t": t, "changes": changes})
        append_event({"time": iso, "t": t, "event": "popup_changed",
                      "popup_id": record["popup_id"], "hwnd": hwnd, "changes": changes})
        if not props["sensitive"] and any(c["type"] != "enabled" for c in changes):
            schedule_uia_scan(record["popup_id"], hwnd, "changed")

    def _closed(self, hwnd, iso, t):
        record = self.open.pop(hwnd)
        self.signatures.pop(hwnd, None)
        hook = consume_hook_time(hwnd, CLOSE_HOOK_EVENTS, t)
        closed_iso, closed_t = hook if hook else (iso, t)
        destroyed = not safe(lambda: win32gui.IsWindow(hwnd), False)
        closer = last_action_before(closed_t, fg_hwnd=hwnd, max_age=CLOSE_LINK_SECONDS)
        with DATA_LOCK:
            record.update({
                "closed_at": closed_iso,
                "closed_t": closed_t,
                "close_time_source": "win_event" if hook else "poll",
                "duration_s": round(closed_t - record["opened_t"], 3),
                "close_reason": "destroyed" if destroyed else "hidden",
                "closed_by_action": action_ref(closer, closed_t),
                "last_action_before_close": action_ref(last_action_before(closed_t), closed_t),
            })
        append_event({"time": closed_iso, "t": closed_t, "event": "popup_closed",
                      "popup_id": record["popup_id"], "hwnd": hwnd, "title": record["title"],
                      "duration_s": record["duration_s"], "close_reason": record["close_reason"],
                      "closed_by_action": record["closed_by_action"]})

    # -- main window state

    def _check_main(self, main, iso, t):
        title = internal_text(main)
        enabled = bool(safe(lambda: win32gui.IsWindowEnabled(main), True))
        hung = is_hung(main)
        mdi = mdi_active(main)
        prev = self.main
        if prev["hwnd"] != main:
            self.main = {"hwnd": main, "title": title, "enabled": enabled, "hung": hung, "mdi": mdi}
            append_event({"time": iso, "t": t, "event": "main_window_found", "hwnd": main,
                          "title": title, "enabled": enabled, "hung": hung})
            request_main_screen("main window found")
            return
        if title != prev["title"]:
            request_main_screen("title changed")
        if mdi != prev.get("mdi"):
            request_main_screen("active tab changed")
        if enabled and not prev["enabled"]:
            request_main_screen("main window usable again")
        if title != prev["title"]:
            hook = consume_hook_time(main, ("name_change",), t)
            e_iso, e_t = hook if hook else (iso, t)
            append_event({"time": e_iso, "t": e_t, "event": "main_title_changed",
                          "from": prev["title"], "to": title,
                          "player_ids": player_ids_in(title),
                          "after_action": action_ref(last_action_before(e_t), e_t)})
        if enabled != prev["enabled"]:
            hook = consume_hook_time(main, ("state_change",), t)
            e_iso, e_t = hook if hook else (iso, t)
            open_ids = [r["popup_id"] for r in self.open.values()]
            if not enabled:
                for r in self.open.values():
                    if r["kind"] in ("dialog", "login", "window", "winforms_window"):
                        r["modal"] = True
            append_event({"time": e_iso, "t": e_t, "event": "main_enabled_changed",
                          "enabled": enabled, "open_popup_ids": open_ids})
        if hung != prev["hung"]:
            append_event({"time": iso, "t": t, "event": "main_hung_changed", "hung": hung,
                          "open_popup_ids": [r["popup_id"] for r in self.open.values()]})
        if not enabled:
            for r in self.open.values():
                if r["kind"] in ("dialog", "login", "window", "winforms_window"):
                    r["modal"] = True
        self.main = {"hwnd": main, "title": title, "enabled": enabled, "hung": hung, "mdi": mdi}

    def _check_foreground(self, pids, iso, t):
        global TEXT_RUN_ACTIVE
        fg = safe(win32gui.GetForegroundWindow, 0) or 0
        if fg == self.foreground:
            return
        TEXT_RUN_ACTIVE = False
        old, self.foreground = self.foreground, fg
        fg_is_pm = hwnd_pid(fg) in pids if fg else False
        old_is_pm = old is not None and KNOWN_HWND_PID.get(old) in pids or old == PM_MAIN_HWND
        if fg_is_pm or old_is_pm:
            title = internal_text(fg) if fg_is_pm else "<not PM>"
            append_event({"time": iso, "t": t, "event": "foreground_changed", "hwnd": fg,
                          "is_pm": fg_is_pm,
                          "title": "Patron Management Login" if LOGIN_RE.search(title) else title,
                          "class_name": clean(safe(lambda: win32gui.GetClassName(fg), "")) if fg_is_pm else ""})

    def finalize(self):
        with DATA_LOCK:
            for record in self.open.values():
                record["still_open_at_stop"] = True


def popup_watch_loop(tracker):
    while not STOP_EVENT.is_set():
        WAKE_EVENT.wait(POPUP_POLL_SECONDS)
        WAKE_EVENT.clear()
        try:
            tracker.poll()
        except Exception as exc:
            record_error("popup_watch", exc)


def save_popup_screenshot(record):
    if SCREENSHOT_DIR is None or record["sensitive"]:
        return None
    try:
        from PIL import ImageGrab
        r = win32gui.GetWindowRect(record["hwnd"])
        path = SCREENSHOT_DIR / f"popup_{record['popup_id']:03d}_{record['kind']}.png"
        ImageGrab.grab(bbox=r, all_screens=True).save(path)
        return str(path.relative_to(SCREENSHOT_DIR.parent))
    except Exception as exc:
        record_error("popup_screenshot", exc)
        return None


def popup_uia_worker():
    """UIA scans of individual popups; kept off the watcher thread because UIA can block."""
    while not STOP_EVENT.is_set():
        try:
            due, popup_id, hwnd, reason, attempt = UIA_JOBS.get(timeout=0.2)
        except queue.Empty:
            continue
        wait = due - now_t()
        if wait > 0 and STOP_EVENT.wait(wait):
            break
        record = POPUP_BY_ID.get(popup_id)
        if record is None or len(record["uia_snapshots"]) >= POPUP_MAX_UIA_SCANS:
            continue
        if record["closed_at"] is not None or not safe(lambda: win32gui.IsWindow(hwnd), False):
            with DATA_LOCK:
                record["uia_snapshots"].append({"reason": reason, "skipped": "closed_before_scan"})
            continue
        if is_hung(PM_MAIN_HWND) and attempt < 10:
            UIA_JOBS.put((now_t() + 0.5, popup_id, hwnd, reason, attempt + 1))
            continue
        if POPUP_SCREENSHOTS and record["screenshot"] is None:
            record["screenshot"] = save_popup_screenshot(record)
        started_iso, started_t = stamp()
        try:
            wrapper = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
            controls = scan_one_tree(wrapper, 0, with_state=True)
            iso, t = stamp()
            snap = {"reason": reason, "scan_started_at": started_iso, "scan_started_t": started_t,
                    "time": iso, "t": t, "control_count": len(controls), "controls": controls}
            with DATA_LOCK:
                record["uia_snapshots"].append(snap)
            append_event({"time": iso, "t": t, "event": "popup_uia_snapshot", "popup_id": popup_id,
                          "reason": reason, "control_count": len(controls),
                          "scan_seconds": round(t - started_t, 3)})
            window = {"hwnd": hwnd, "title": record["title"], "class_name": record["class_name"],
                      "kind": record["kind"],
                      "rectangle": rect_from_tuple(safe(lambda: win32gui.GetWindowRect(hwnd), None))}
            register_screen(window, controls, f"popup #{popup_id} {reason}", popup_id)
        except Exception as exc:
            with DATA_LOCK:
                record["uia_snapshots"].append({"reason": reason, "error": str(exc)})
            record_error("popup_uia_worker", exc)


# ------------------------------------------------------------- screen map

def mdi_active(main):
    client = safe(lambda: win32gui.FindWindowEx(main, 0, "MDIClient", None), 0)
    if not client:
        return 0
    return safe(lambda: win32gui.SendMessageTimeout(client, 0x0229, 0, 0,
                                                   win32con.SMTO_ABORTIFHUNG, 200)[1], 0) or 0


def request_main_screen(reason, delay=SCREEN_SETTLE_SECONDS):
    """Capture the main window a moment after it changed (requests are merged)."""
    with DATA_LOCK:
        MAIN_SCREEN_REQUEST["due"] = now_t() + delay
        if reason not in MAIN_SCREEN_REQUEST["reasons"]:
            MAIN_SCREEN_REQUEST["reasons"].append(reason)


def after_user_action(fg):
    """Rescan the popup the user acted in, or the main window."""
    hwnd = (fg or {}).get("hwnd")
    record = TRACKER.open.get(hwnd) if TRACKER else None
    if record is not None:
        if not record["sensitive"]:
            schedule_uia_scan(record["popup_id"], hwnd, "after_action", delay=0.6)
    elif hwnd == PM_MAIN_HWND:
        request_main_screen("after user action")


def title_pattern(title):
    if title.startswith("Patron Management - ") and player_ids_in(title):
        return "Patron Management - <player profile>"
    return re.sub(r"\(\d+\)", "(<id>)", title)


def screen_name(name, ctype):
    if RECORD_SCREEN_VALUES or ctype not in VALUE_TYPES:
        return name
    if ctype == "Text" and name.endswith(":"):
        return name            # a field label
    return "<value>" if name else ""


def map_category(ctype):
    for category, colour, types in MAP_CATEGORIES:
        if ctype in types:
            return category, colour
    return None, None


def find_frames(items):
    """Group boxes of the screen; each control is assigned to the smallest frame around it."""
    frames = [i for i in items if i["type"] == "Group"]
    for item in items:
        best = None
        for frame in frames:
            if frame is item:
                continue
            if (frame["x"] - 2 <= item["x"] and item["x"] + item["w"] <= frame["x"] + frame["w"] + 2
                    and frame["y"] - 2 <= item["y"] and item["y"] + item["h"] <= frame["y"] + frame["h"] + 2):
                if best is None or frame["w"] * frame["h"] < best["w"] * best["h"]:
                    best = frame
        item["frame"] = best["name"] if best else ""
        item["frame_id"] = best["automation_id"] if best else ""
    return [{"name": f["name"], "automation_id": f["automation_id"], "x": f["x"], "y": f["y"],
             "w": f["w"], "h": f["h"], "controls": sum(1 for i in items if i.get("frame_id") == f["automation_id"]
                                                        and i.get("frame") == f["name"])}
            for f in frames]


def register_screen(window, controls, reason, popup_id=None, extra=None):
    """Store the layout of a window once; later identical layouts only add an occurrence."""
    wrect = window["rectangle"]
    if wrect["left"] is None or not wrect["width"] or not wrect["height"]:
        return
    items = []
    for c in controls:
        r = c.get("rectangle") or {}
        if c.get("depth", 0) == 0 or c.get("visible") is False or r.get("left") is None:
            continue
        if not r["width"] or not r["height"]:
            continue
        x, y = r["left"] - wrect["left"], r["top"] - wrect["top"]
        if x >= wrect["width"] or y >= wrect["height"] or x + r["width"] <= 0 or y + r["height"] <= 0:
            continue   # outside the window
        items.append({"type": c["control_type"], "name": screen_name(c["name"], c["control_type"]),
                      "value": c.get("value") if RECORD_SCREEN_VALUES else None,
                      "layout_name": screen_name(c["name"], c["control_type"]) if RECORD_SCREEN_VALUES is False
                      else ("<value>" if c["control_type"] in VALUE_TYPES and c["name"]
                            and not c["name"].endswith(":") else c["name"]),
                      "automation_id": c["automation_id"], "class_name": c["class_name"],
                      "x": x, "y": y, "w": r["width"], "h": r["height"], "depth": c.get("depth"),
                      "enabled": c.get("enabled"), "state": c.get("state") if RECORD_SCREEN_VALUES
                      or c["control_type"] in ("RadioButton", "CheckBox", "TabItem") else None})
    frames = find_frames(items)
    # Layout signature ignores values, so one screen is stored once for every player.
    layout = sorted((i["type"], i["automation_id"], re.sub(r"\d", "#", i.pop("layout_name")),
                     i["x"] // 8, i["y"] // 8, i["w"] // 8, i["h"] // 8) for i in items)
    pattern = title_pattern(window["title"])
    signature = hashlib.sha256(json.dumps([window["kind"], pattern, layout]).encode("utf-8")).hexdigest()[:16]
    iso, t = stamp()
    with DATA_LOCK:
        screen = SCREEN_BY_SIG.get(signature)
        is_new = screen is None
        if is_new:
            if len(SCREENS) >= MAX_SCREENS:
                return
            screen = {"screen_id": len(SCREENS) + 1, "signature": signature, "kind": window["kind"],
                      "title_pattern": pattern, "class_name": window["class_name"],
                      "size": {"w": wrect["width"], "h": wrect["height"]},
                      "first_seen_at": iso, "first_seen_t": t, "popup_id": popup_id,
                      "control_count": len(items), "frames": frames, "controls": items,
                      "occurrences": []}
            SCREENS.append(screen)
            SCREEN_BY_SIG[signature] = screen
        occurrence = {"time": iso, "t": t, "reason": reason, "popup_id": popup_id,
                      "position": {"x": wrect["left"], "y": wrect["top"]},
                      "window_title": window["title"]}
        if RECORD_SCREEN_VALUES:
            occurrence["values"] = {f"{i['frame'] or '-'}/{i['automation_id'] or i['type']}": i.get("value")
                                    for i in items if i.get("value") not in (None, "")}
        if extra:
            occurrence.update(extra)
        screen["occurrences"].append(occurrence)
    append_event({"time": iso, "t": t, "event": "screen_seen", "screen_id": screen["screen_id"],
                  "new": is_new, "kind": window["kind"], "title": pattern, "reason": reason,
                  "control_count": len(items), "frames": [f["name"] for f in frames],
                  "identification": (extra or {}).get("identification")})


def screen_worker():
    """Captures the main window layout when a change was requested and PM is idle."""
    while not STOP_EVENT.is_set():
        if STOP_EVENT.wait(0.1):
            break
        with DATA_LOCK:
            due, reasons = MAIN_SCREEN_REQUEST["due"], list(MAIN_SCREEN_REQUEST["reasons"])
        if due is None or now_t() < due:
            continue
        main = PM_MAIN_HWND
        busy = (not main or is_hung(main) or not safe(lambda: win32gui.IsWindowEnabled(main), False)
                or (TRACKER is not None and TRACKER.open))
        if busy:
            request_main_screen(reasons[0] if reasons else "retry", delay=0.5)
            continue
        with DATA_LOCK:
            MAIN_SCREEN_REQUEST["due"], MAIN_SCREEN_REQUEST["reasons"] = None, []
        try:
            wrapper = Desktop(backend="uia").window(handle=main).wrapper_object()
            controls = scan_one_tree(wrapper, 0, with_state=True)
            window = {"hwnd": main, "title": internal_text(main), "class_name": "XTPMainFrame",
                      "kind": "main",
                      "rectangle": rect_from_tuple(safe(lambda: win32gui.GetWindowRect(main), None))}
            window["class_name"] = clean(safe(lambda: win32gui.GetClassName(main), ""))
            extra = {}
            identification = read_identification(main, controls, window["rectangle"])
            if identification:
                extra["identification"] = identification
            register_screen(window, controls, ", ".join(reasons), extra=extra)
        except Exception as exc:
            record_error("screen_worker", exc)


def read_identification(main, controls, wrect):
    """Player name in the Identification frame, its colours and the (Loc:) flag."""
    visible = [c for c in controls if c.get("visible") is not False and c.get("rectangle", {}).get("left") is not None]
    frame = next((c for c in visible if c["automation_id"] == IDENT_FRAME_ID
                  or (c["control_type"] == "Group" and c["name"] == "Identification")), None)
    if frame is None:
        return None
    fr = frame["rectangle"]
    inside = [c for c in visible if c is not frame and c["control_type"] == "Text"
              and fr["left"] <= c["rectangle"]["left"] and c["rectangle"]["right"] <= fr["right"]
              and fr["top"] <= c["rectangle"]["top"] and c["rectangle"]["bottom"] <= fr["bottom"]]
    name = next((c for c in inside if c["automation_id"] == IDENT_NAME_ID), None)
    source = "control id " + IDENT_NAME_ID
    if name is None and inside:
        name = min(inside, key=lambda c: (c["rectangle"]["top"], c["rectangle"]["left"]))
        source = "top-most label in Identification"
    if name is None:
        return None
    text = name.get("value") or name["name"]
    info = {"text": text, "control_id": name["automation_id"], "frame_id": frame["automation_id"],
            "found_by": source, "has_loc": bool(SKIP_NAME_RE.search(text or ""))}
    r = name["rectangle"]
    try:
        image = capture_window(main)
        if image:
            colors = region_colors(image, r["left"] - wrect["left"], r["top"] - wrect["top"],
                                   r["right"] - wrect["left"], r["bottom"] - wrect["top"])
            if colors:
                info.update(colors)
                info["background_is_green"] = colors["background_color_name"] == "green"
    except Exception as exc:
        record_error("identification_color", exc)
    return info


def build_screens_html():
    with DATA_LOCK:
        screens = [dict(s) for s in SCREENS]
    esc = html.escape
    out = ["<!doctype html><html><head><meta charset='utf-8'>",
           "<meta name='viewport' content='width=device-width, initial-scale=1'>",
           "<title>PM screen map</title><style>",
           ":root{--bg:#fff;--fg:#1f2328;--muted:#6b7280;--line:#d0d7de;--win:#f6f8fa}",
           "@media (prefers-color-scheme: dark){:root{--bg:#0d1117;--fg:#e6edf3;--muted:#9da7b3;"
           "--line:#30363d;--win:#161b22}}",
           "body{font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--fg);margin:16px}",
           "h2{margin-top:40px;border-top:1px solid var(--line);padding-top:16px}",
           ".meta{color:var(--muted);font-size:13px}",
           ".map{overflow-x:auto;margin:12px 0} svg text{font-family:Segoe UI,Arial;pointer-events:none}",
           "table{border-collapse:collapse;font-size:12px} td,th{border:1px solid var(--line);padding:2px 6px}",
           "th{background:var(--win)} .legend span{display:inline-block;margin-right:14px}",
           ".sw{display:inline-block;width:12px;height:12px;margin-right:4px;vertical-align:middle}",
           "details{margin:8px 0}</style></head><body>",
           f"<h1>PM screen map</h1><p class='meta'>Recorded {esc(STARTED_AT.isoformat(timespec='seconds'))}"
           f" &middot; {len(screens)} distinct screen(s). Positions are pixels from the window's top-left "
           "corner. Hover a box for details.</p><p class='legend'>"]
    out += [f"<span><i class='sw' style='background:{c}'></i>{esc(cat)}</span>" for cat, c, _ in MAP_CATEGORIES]
    out.append("</p><ul>")
    out += [f"<li><a href='#s{s['screen_id']}'>#{s['screen_id']} [{esc(s['kind'])}] "
            f"{esc(s['title_pattern'] or '(no title)')}</a> &middot; {len(s['occurrences'])}x</li>" for s in screens]
    out.append("</ul>")
    for s in screens:
        w, h = s["size"]["w"], s["size"]["h"]
        scale = min(1.0, 1100.0 / max(w, 1))
        reasons = sorted({o["reason"] for o in s["occurrences"]})
        out.append(f"<h2 id='s{s['screen_id']}'>Screen #{s['screen_id']} [{esc(s['kind'])}] "
                   f"{esc(s['title_pattern'] or '(no title)')}</h2>")
        out.append(f"<p class='meta'>{w}&times;{h}px &middot; class {esc(s['class_name'])} &middot; "
                   f"first seen {esc(s['first_seen_at'][11:23])} &middot; seen {len(s['occurrences'])}x "
                   f"&middot; {s['control_count']} controls &middot; triggers: {esc('; '.join(reasons)[:300])}</p>")
        out.append(f"<div class='map'><svg viewBox='0 0 {w} {h}' width='{int(w * scale)}' "
                   f"height='{int(h * scale)}' role='img' aria-label='Screen {s['screen_id']}'>")
        out.append(f"<rect x='0' y='0' width='{w}' height='{h}' fill='var(--win)' stroke='var(--line)'/>")
        drawn = [c for c in s["controls"] if map_category(c["type"])[0]]
        for c in sorted(drawn, key=lambda c: -(c["w"] * c["h"])):
            category, colour = map_category(c["type"])
            dash = " stroke-dasharray='6 4'" if category == "group" else ""
            opacity = "0.10" if category in ("group", "list") else "0.18"
            tip = (f"{c['type']} '{c['name']}'" + (f" value='{c['value']}'" if c.get("value") else "")
                   + (f" in frame '{c['frame']}'" if c.get("frame") else "")
                   + f" id={c['automation_id']} at ({c['x']},{c['y']}) "
                   f"{c['w']}x{c['h']}" + ("" if c.get("enabled") is not False else " disabled")
                   + (f" {c['state']}" if c.get("state") else ""))
            out.append(f"<g><title>{esc(tip)}</title><rect x='{c['x']}' y='{c['y']}' width='{c['w']}' "
                       f"height='{c['h']}' fill='{colour}' fill-opacity='{opacity}' stroke='{colour}' "
                       f"stroke-width='{1 / scale:.2f}'{dash}/>")
            label = c["name"] or c["automation_id"]
            size_px = min(11.0, c["h"] * scale - 1)          # shrink the label to fit small boxes
            chars = int(c["w"] * scale / (size_px * 0.55)) if size_px > 0 else 0
            if label and size_px >= 7 and chars >= 3:
                size = size_px / scale
                out.append(f"<text x='{c['x'] + 2 / scale:.1f}' y='{c['y'] + size:.1f}' font-size='{size:.1f}' "
                           f"fill='{colour}'>{esc(label[:chars])}</text>")
            out.append("</g>")
        out.append("</svg></div>")
        if s.get("frames"):
            out.append("<p class='meta'>Frames: " + ", ".join(
                f"{esc(f['name'] or '(no name)')} (id {esc(f['automation_id'])}, {f['controls']} controls)"
                for f in s["frames"]) + "</p>")
        idents = [o for o in s["occurrences"] if o.get("identification")]
        if idents:
            out.append("<details open><summary>Identification name per capture</summary><table><tr>"
                       "<th>Time</th><th>Name</th><th>Text colour</th><th>Background</th><th>Green background</th>"
                       "<th>(Loc:)</th></tr>")
            for o in idents:
                i = o["identification"]
                swatch = lambda c: (f"<i class='sw' style='background:{c}'></i>{esc(c)}" if c else "")
                out.append(f"<tr><td>{esc(o['time'][11:23])}</td><td>{esc(i.get('text') or '')}</td>"
                           f"<td>{swatch(i.get('text_color'))} {esc(i.get('text_color_name') or '')}</td>"
                           f"<td>{swatch(i.get('background_color'))} {esc(i.get('background_color_name') or '')}</td>"
                           f"<td>{'YES' if i.get('background_is_green') else 'no'}</td>"
                           f"<td>{'YES - runner skips' if i.get('has_loc') else 'no'}</td></tr>")
            out.append("</table></details>")
        out.append("<details><summary>Control list</summary><table><tr><th>Frame</th><th>Type</th><th>Name</th>"
                   "<th>Value</th><th>ID</th><th>X</th><th>Y</th><th>W</th><th>H</th><th>Enabled</th>"
                   "<th>State</th></tr>")
        for c in sorted(s["controls"], key=lambda c: (c.get("frame", ""), c["y"], c["x"])):
            out.append(f"<tr><td>{esc(c.get('frame', ''))}</td><td>{esc(c['type'])}</td><td>{esc(c['name'])}</td>"
                       f"<td>{esc(c.get('value') or '')}</td><td>{esc(c['automation_id'])}</td>"
                       f"<td>{c['x']}</td><td>{c['y']}</td><td>{c['w']}</td><td>{c['h']}</td>"
                       f"<td>{'' if c.get('enabled') is None else c['enabled']}</td>"
                       f"<td>{esc(str(c['state'])) if c.get('state') else ''}</td></tr>")
        out.append("</table></details>")
    out.append("</body></html>")
    return "\n".join(out)


# ---------------------------------------------------------------- reports

def fmt_t(t):
    minutes, seconds = divmod(max(0.0, t or 0.0), 60)
    return f"+{int(minutes):02d}:{seconds:06.3f}"


def first_uia_controls(record):
    for snap in record.get("uia_snapshots", []):
        if snap.get("controls"):
            return snap["controls"]
    return []


def popup_buttons(record):
    names = [c["text"] for c in record.get("controls_win32", [])
             if c["class_name"].lower() == "button" and c.get("button_style", "push") in ("push", "owner_draw")
             and c["visible"] and c.get("text")]
    if names:
        return names
    controls = first_uia_controls(record)
    by_index = {c["index"]: c for c in controls}
    out = []
    for c in controls:
        if c["control_type"] not in ("Button", "MenuItem") or not c["name"]:
            continue
        parent = by_index.get(c["parent_index"]) or {}
        if parent.get("control_type") == "TitleBar" and c["name"].lower() in TITLEBAR_BUTTONS:
            continue
        out.append(c["name"])
    return out


def layout_lines(record, indent="      "):
    lines = []
    for c in record.get("controls_win32", []):
        if not c["visible"]:
            continue
        rel = c.get("rel") or {}
        text = "<not recorded>" if c.get("text") is None else f"\"{c['text']}\""
        flags = "" if c["enabled"] else "  [disabled]"
        cls = c["class_name"] + (f"/{c['button_style']}" if c.get("button_style") else "")
        lines.append(f"{indent}{cls[:22]:<22} id={str(c['control_id']):<6} {text:<34} "
                     f"at ({rel.get('x')},{rel.get('y')}) {rel.get('w')}x{rel.get('h')}{flags}")
    interesting = {"Button", "RadioButton", "CheckBox", "ComboBox", "Edit", "MenuItem",
                   "ListItem", "TabItem", "Text", "List", "Tree", "DataItem"}
    uia = [c for c in first_uia_controls(record) if c["control_type"] in interesting]
    if uia:
        lines.append(f"{indent}UIA view:")
        for c in uia[:80]:
            state = c.get("state")
            flags = ("" if c.get("enabled", True) else "  [disabled]") + (f"  {state}" if state else "")
            lines.append(f"{indent}  {c['control_type']:<12} \"{c['name'][:40]}\" aid={c['automation_id']}{flags}")
    return lines


def build_timeline():
    with DATA_LOCK:
        events = sorted(EVENTS, key=lambda e: e.get("t", 0))
        popups = {p["popup_id"]: p for p in POPUPS}
    lines = ["PM RECORDING TIMELINE (v5)", "=" * 100,
             "Columns: seconds since start | wall clock | event", ""]
    for e in events:
        kind = e.get("event")
        head = f"{fmt_t(e.get('t'))}  {str(e.get('time', ''))[11:23]}  "
        fg_title = (e.get("foreground") or {}).get("title", "")
        if kind in ("mouse_click", "mouse_double_click"):
            tgt = e.get("target") or {}
            word = "DOUBLE-CLICK" if kind == "mouse_double_click" else "CLICK"
            lines.append(head + f"{word:<13}{target_label(tgt)} id={tgt.get('automation_id', '')}  in \"{fg_title}\"")
        elif kind == "protected_field_click":
            lines.append(head + f"{'CLICK':<13}<protected field>  in \"{fg_title}\"")
        elif kind == "credential_input":
            lines.append(head + f"{'TYPE':<13}<credentials - not recorded>")
        elif kind == "text_input":
            lines.append(head + f"{'TYPE':<13}<text - not recorded>  in \"{fg_title}\"")
        elif kind == "key_press":
            lines.append(head + f"{'KEY':<13}{'+'.join(e.get('modifiers', []) + [e.get('key', '')])}  in \"{fg_title}\"")
        elif kind == "mouse_scroll":
            lines.append(head + f"{'SCROLL':<13}dy={e.get('dy')} {target_label(e.get('target'))}")
        elif kind == "popup_opened":
            p = popups.get(e["popup_id"], {})
            after = p.get("opened_after_action")
            cause = f"  ({after['delay_ms']/1000:.2f}s after {after['label']})" if after else ""
            lines.append(head + f">>> POPUP OPEN  #{e['popup_id']} [{p.get('kind')}] \"{p.get('title')}\" "
                                f"class={p.get('class_name')} hwnd={p.get('hwnd')} owner={p.get('owner_hwnd')}{cause}")
            rect = p.get("rectangle") or {}
            lines.append(f"{'':34}size {rect.get('width')}x{rect.get('height')} at ({rect.get('left')},{rect.get('top')})"
                         f"  modal={p.get('modal')}  buttons: {', '.join(popup_buttons(p)) or '-'}")
        elif kind == "popup_changed":
            summary = []
            for c in e.get("changes", []):
                if c["type"] in ("title", "enabled"):
                    summary.append(f"{c['type']} {c['from']!r}->{c['to']!r}")
                elif c["type"] == "control_changed":
                    summary.append(f"id={c['control_id']} {c['changes']}")
                else:
                    summary.append(f"{c['type']} id={c.get('control_id')} {c.get('text')!r}")
            lines.append(head + f"~~~ POPUP CHANGE #{e['popup_id']}: " + "; ".join(summary)[:300])
        elif kind == "popup_closed":
            by = e.get("closed_by_action")
            closer = f"  closed by {by['label']} ({by['delay_ms']/1000:.2f}s earlier)" if by else ""
            lines.append(head + f"<<< POPUP CLOSE #{e['popup_id']} \"{e.get('title')}\" after "
                                f"{e.get('duration_s')}s [{e.get('close_reason')}]{closer}")
        elif kind == "main_window_found":
            lines.append(head + f"=== MAIN WINDOW \"{e.get('title')}\" hwnd={e.get('hwnd')}")
        elif kind == "main_title_changed":
            lines.append(head + f"=== MAIN TITLE  \"{e.get('from')}\" -> \"{e.get('to')}\"")
        elif kind == "main_enabled_changed":
            state = "ENABLED" if e.get("enabled") else "DISABLED (blocked by modal popup)"
            lines.append(head + f"=== MAIN {state}  open popups: {e.get('open_popup_ids')}")
        elif kind == "main_hung_changed":
            state = "PM NOT RESPONDING (busy)" if e.get("hung") else "PM RESPONDING AGAIN"
            lines.append(head + f"=== {state}")
        elif kind == "foreground_changed":
            lines.append(head + f"--- FOREGROUND -> \"{e.get('title')}\" {e.get('class_name', '')}")
        elif kind == "ui_snapshot":
            lines.append(head + f"    (full UI snapshot #{e.get('snapshot_id')}, scan {e.get('scan_seconds')}s)")
        elif kind == "screen_seen":
            state = "NEW" if e.get("new") else "same as"
            lines.append(head + f"    [screen #{e.get('screen_id')} {state}] [{e.get('kind')}] "
                                f"\"{e.get('title')}\" {e.get('control_count')} controls ({e.get('reason')})")
            ident = e.get("identification")
            if ident:
                lines.append(f"{'':34}IDENTIFICATION name=\"{ident.get('text')}\" text colour={ident.get('text_color_name')} "
                             f"background={ident.get('background_color_name')} green background={'YES' if ident.get('background_is_green') else 'no'} (Loc:)={'YES' if ident.get('has_loc') else 'no'}")
        elif kind == "popup_uia_snapshot":
            lines.append(head + f"    (popup #{e.get('popup_id')} UIA snapshot: {e.get('control_count')} controls, "
                                f"{e.get('reason')}, scan {e.get('scan_seconds')}s)")
    lines += ["", "", "POPUP CATALOG", "=" * 100]
    groups = {}
    for p in popups.values():
        groups.setdefault((p["kind"], p["class_name"], p["title"] or "(no title)"), []).append(p)
    for (kind, cls, title), items in groups.items():
        durations = [p["duration_s"] for p in items if p["duration_s"] is not None]
        delays = [p["opened_after_action"]["delay_ms"] / 1000 for p in items if p["opened_after_action"]]
        triggers = sorted({p["opened_after_action"]["label"] for p in items if p["opened_after_action"]})
        closers = sorted({p["closed_by_action"]["label"] for p in items if p["closed_by_action"]})
        first = items[0]
        rect = first["rectangle"]
        lines.append(f"\n[{kind}] \"{title}\"  class={cls}  seen {len(items)}x  ids={[p['popup_id'] for p in items]}")
        if durations:
            lines.append(f"    open duration  min {min(durations):.2f}s  avg {sum(durations)/len(durations):.2f}s  max {max(durations):.2f}s")
        if delays:
            lines.append(f"    appears after action  min {min(delays):.2f}s  avg {sum(delays)/len(delays):.2f}s  max {max(delays):.2f}s")
        lines.append(f"    triggered after: {triggers or '-'}")
        lines.append(f"    closed by: {closers or '-'}")
        lines.append(f"    modal (blocks main window): {any(p['modal'] for p in items)}   "
                     f"owner: \"{first['owner_title']}\"   size {rect['width']}x{rect['height']}")
        lines.append(f"    buttons: {', '.join(popup_buttons(first)) or '-'}")
        if not first["sensitive"]:
            lines.append("    layout (first occurrence, positions relative to popup):")
            lines.extend(layout_lines(first, indent="      "))
    return "\n".join(lines) + "\n"


def fix_close_reasons():
    """A popup is hidden a few ms before it is destroyed; use the destroy WinEvent."""
    with DATA_LOCK:
        destroyed = {}
        for e in WIN_EVENTS:
            if e["event"] == "destroy":
                destroyed.setdefault(e["hwnd"], []).append(e["t"])
        for p in POPUPS:
            if p["close_reason"] == "hidden" and p["closed_t"] is not None:
                if any(p["closed_t"] <= t <= p["closed_t"] + 2 for t in destroyed.get(p["hwnd"], [])):
                    p["close_reason"] = "destroyed"


def export_files(tracker):
    tracker.finalize()
    fix_close_reasons()
    stopped = datetime.now()
    folder = Path(__file__).resolve().parent / OUTPUT_DIR
    folder.mkdir(parents=True, exist_ok=True)
    with DATA_LOCK:
        actions = [x for x in EVENTS if x.get("event") in USER_ACTION_EVENTS]
        metadata = {"version": RECORDER_VERSION, "started_at": STARTED_AT.isoformat(timespec="seconds"),
                    "stopped_at": stopped.isoformat(timespec="seconds"),
                    "typed_text_recorded": False, "credentials_recorded": False,
                    "screenshots_recorded": POPUP_SCREENSHOTS, "stop_key": "F8",
                    "time_field_t": "seconds since recorder start (monotonic)",
                    "popup_tracking": {"poll_seconds": POPUP_POLL_SECONDS,
                                       "win_event_hooks_installed": HOOK_STATUS["installed"],
                                       "win_event_hook_error": HOOK_STATUS["error"]}}
        payload = {"metadata": metadata, "events": sorted(EVENTS, key=lambda e: e.get("t", 0)),
                   "popups": POPUPS, "win_events": WIN_EVENTS,
                   "ui_snapshots": SNAPSHOTS, "errors": ERRORS}
        (folder / "pm_recording.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        (folder / "popups.json").write_text(json.dumps({"metadata": metadata, "popups": POPUPS}, ensure_ascii=False, indent=2), encoding="utf-8")
        (folder / "actions_only.json").write_text(json.dumps({"metadata": metadata, "actions": actions}, ensure_ascii=False, indent=2), encoding="utf-8")
        popup_lines = [
            f"  #{p['popup_id']:<3} [{p['kind']}] \"{p['title']}\"  open {fmt_t(p['opened_t'])}  "
            + (f"close {fmt_t(p['closed_t'])}  ({p['duration_s']}s)" if p["closed_t"] is not None else "STILL OPEN AT STOP")
            for p in POPUPS
        ]
        summary_lines = [
            "PM RECORDING", "=" * 60,
            f"Recorder version: {RECORDER_VERSION}",
            f"Started: {metadata['started_at']}",
            f"Stopped: {metadata['stopped_at']}",
            f"Events: {len(EVENTS)}", f"Actions: {len(actions)}",
            f"Popups: {len(POPUPS)}", f"WinEvents: {len(WIN_EVENTS)}",
            f"WinEvent hooks installed: {HOOK_STATUS['installed']}",
            f"UI snapshots: {len(SNAPSHOTS)}", f"Screens (distinct layouts): {len(SCREENS)}",
            f"Errors: {len(ERRORS)}",
            "Typed text: REDACTED", "Credentials: REDACTED", "",
            "Popups:", *popup_lines,
        ]
        (folder / "summary.txt").write_text("\n".join(summary_lines), encoding="utf-8")
        (folder / "screens.json").write_text(json.dumps({"metadata": metadata, "screens": SCREENS},
                                                        ensure_ascii=False, indent=2), encoding="utf-8")
    (folder / "screens.html").write_text(build_screens_html(), encoding="utf-8")
    (folder / "timeline.txt").write_text(build_timeline(), encoding="utf-8")
    print(f"\nExport complete: {folder}")
    try: os.startfile(folder)
    except Exception: pass


def main():
    global SCREENSHOT_DIR, TRACKER
    print("PM INTERACTION RECORDER v7 (popup lifecycle + full screen map with values)")
    print("Login activity is recorded, but username/password content is never stored. Press F8 to stop.\n")
    if not discover_pm_windows():
        print("Open Patron Management first, then run this file again.")
        input("Press Enter to exit..."); return 1
    if POPUP_SCREENSHOTS:
        SCREENSHOT_DIR = Path(__file__).resolve().parent / OUTPUT_DIR / f"popup_images_{STARTED_AT:%Y%m%d_%H%M%S}"
        SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)

    hook = WinEventHook(); hook.start(); hook.ready.wait(2)
    print(f"WinEvent hooks installed: {HOOK_STATUS['installed']}/{len(HOOK_RANGES)}"
          + ("" if HOOK_STATUS["installed"] else "  (falling back to polling only)"))
    tracker = TRACKER = PopupTracker()
    workers = [threading.Thread(target=scan_worker, daemon=True),
               threading.Thread(target=popup_watch_loop, args=(tracker,), daemon=True),
               threading.Thread(target=popup_uia_worker, daemon=True),
               threading.Thread(target=screen_worker, daemon=True)]
    for worker in workers:
        worker.start()
    request_scan(); WAKE_EVENT.set()
    ml = mouse.Listener(on_click=on_click, on_scroll=on_scroll)
    kl = keyboard.Listener(on_press=on_press, on_release=on_release)
    ml.start(); kl.start()
    try:
        while not STOP_EVENT.wait(0.25): pass
    except KeyboardInterrupt:
        STOP_EVENT.set()
    finally:
        STOP_EVENT.set(); WAKE_EVENT.set()
        ml.stop(); kl.stop(); hook.stop()
        for worker in workers:
            worker.join(timeout=3)
        hook.join(timeout=2)
        export_files(tracker)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
