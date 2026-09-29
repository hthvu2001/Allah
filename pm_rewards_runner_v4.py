# -*- coding: utf-8 -*-
"""Patron Management rewards automation - v4: four workflows, background control.

Reads Excel-for-auto.xlsx and runs, in this order (rows in Excel order):
  #1 REBATE_SLOT  sheet REBATE, SLOT REBATE 5% != 0
                  Coupon Redemption (F12) > Competitor Coupon > DAILY REBATE (5%) - 1 > Amount
                  > OK > confirmation "The coupon will reward the player with $X in SLOTS." > OK
  #2 REBATE_BBR   sheet REBATE, BBR REBATE 5% != 0
                  Rewards: BBR > Adjust > Add BBR, Adjustment, Expiration = today + 30 days
                  05:59 AM, Reason "P) Your 5% (Rebate)", Comment "REBATE ON <yesterday>"
  #3 COSMO_SLOT   sheet DAILY REWARDS, ALLOCATION = Slot
                  Coupon Redemption (F12) > Competitor Coupon > COSMO ELITE CIRCUIT - 166 > Amount
                  > OK > confirmation > OK
  #4 COSMO_BBR    sheet DAILY REWARDS, ALLOCATION = BBR
                  Rewards: BBR > Adjust > Add BBR, Adjustment, Expiration = today + 3 days
                  05:59 AM, Reason "P) COSMO ELITE CIRCUIT", Comment "COSMO ELITE CIRCUIT"
  #5 MONTHLY_SLOT sheet MONTHLY BNF, FP ALLOCATION = SLOT (optional sheet)
                  Coupon Redemption (F12) > Competitor Coupon > MBS FP - 7 > Amount > OK > confirmation > OK
  #6 MONTHLY_BBR  sheet MONTHLY BNF, FP ALLOCATION = BBR BUCKET
                  Rewards: BBR > Adjust > Add BBR, Adjustment, Expiration = today + 14 days
                  05:59 AM, Reason "G) MBS FP", Comment "MONTHLY BENEFIT - 30SEP2026" (run on the 1st:
                  last day of the previous month; on the 15th: 15<MON><YYYY>; other days: the user is
                  asked). The comment is shown at the start: Enter keeps it, or type another one.

After the run the Excel check file gets a STATUS sheet (the three sheets side by
side, every row with Status + Note) and a "Loc players to check" sheet. Each
profile tab is closed as soon as its player is finished (Loc players too).

Before touching PM the Excel is checked and a report pm_excel_check_<time>.xlsx
is written (PLAN, NOTES, VIOLATIONS). The run stops if there is any violation:
the same player with both Slot and BBR in REBATE or in DAILY REWARDS, a player
twice in the same part, or invalid data. TEST_PLAYER_IDS (10001) are exempt from
the conflict / duplicate rules. Rows where both REBATE columns are 0 are skipped
and noted.

Every job: pre-check idle PM, Find a Player (Ribbon / command), enter the ID,
OK, wait for the profile, handle System Messages / Player Comment, check the
Identification name (skip "(Loc:"), then the workflow steps, each with a check
before and after. When the check after an action finds it not done (e.g. System
Messages still open after Close), the action is repeated: at most MAX_ATTEMPTS
(3) tries in total, only while PM still shows the same screen; then the run
stops. The final OK is clicked exactly once and never repeated. An error that
Windows reports while sending an action is checked on the control itself (the
value PM will use): taken -> continue, not taken -> a new try. Each failed try
and every stop saves pictures of the PM window (the control framed in red, plus
a close-up of that area) in pm_screenshots/<run time>/, named in the log/CSV.
At start the runner asks whether it may click OK (YES = real
redemption / adjustment, anything else = Cancel test run). A coupon OK is
followed by PM's confirmation; its amount / name are compared with the row and
only noted when they differ (the coupon is confirmed anyway). OK clicks are
written to pm_redeemed_ledger.csv per player and workflow, and a rerun on the
same day skips them. The first error stops the run and leaves PM as it is.

Files of a run go to <yyyymmdd>-automation-log/ (text log, CSV log, Excel check,
screenshots, all named with the run time); after an error they are moved to its
"error" sub-folder, after an Excel violation to "violation", and the end of the
terminal says "ERROR - please check: <folder>" or "NO ERROR". The OK ledger stays
next to the script.

Control: CONTROL_MODE = "background" drives PM with Win32 messages and UIA
Invoke only (no mouse, no keyboard, no focus). Controls are found by their
control IDs, never by screen coordinates.
"""

import csv
import ctypes
import re
import shutil
import struct
import zlib
import sys
import threading
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook

try:
    import win32api
    import win32con
    import win32gui
    import win32process
    from pywinauto import Desktop, handleprops, mouse
    from pywinauto import uia_defines
    from pywinauto.controls.hwndwrapper import HwndWrapper
    from pywinauto.keyboard import send_keys
except ImportError:  # lets the step logic be imported and tested off Windows
    win32con = win32gui = win32process = None

RUNNER_VERSION = "4.9-six-workflows"
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Excel-for-auto.xlsx"
LOG_FILE = BASE_DIR / "pm_automation_log.csv"
TEXT_LOG_FILE = BASE_DIR / "pm_runner_log.txt"
REBATE_SHEET = "REBATE"
REBATE_COLUMNS = ("Player ID", "SLOT REBATE 5%", "BBR REBATE 5%")
DAILY_SHEET = "DAILY REWARDS"
DAILY_COLUMNS = ("PLAYER ID", "FREE PLAY", "ALLOCATION")
MONTHLY_SHEET = "MONTHLY BNF"      # optional sheet
MONTHLY_COLUMNS = ("Player ID", "MONTHLY FP check", "FP ALLOCATION")
MONTHLY_ALLOCATIONS = {"slot": "MONTHLY_SLOT", "bbr bucket": "MONTHLY_BBR", "bbr": "MONTHLY_BBR"}
SOURCE_SHEETS = (REBATE_SHEET, DAILY_SHEET, MONTHLY_SHEET)    # the three tables of the STATUS sheet
TEST_PLAYER_IDS = {"10001"}        # exempt from the conflict / duplicate rules
EXPIRATION_TIME = (5, 59)          # 05:59 AM
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]


def rebate_comment(today):
    yesterday = today - timedelta(days=1)
    return f"REBATE ON {MONTHS[yesterday.month - 1]} {yesterday.day:02d} {yesterday.year}"


def monthly_comment(today):
    """(comment, standard): the 15th -> 15<MON><YYYY>; the 1st -> last day of the previous month.

    Any other day is not a standard run day: the most recent of those two dates is suggested
    and the user is asked (see ask_monthly_comment).
    """
    if today.day == 15:
        stamp, standard = today, True
    elif today.day == 1:
        stamp, standard = today - timedelta(days=1), True
    elif today.day > 15:
        stamp, standard = today.replace(day=15), False
    else:
        stamp, standard = today.replace(day=1) - timedelta(days=1), False
    return f"MONTHLY BENEFIT - {stamp.day:02d}{MONTHS[stamp.month - 1]}{stamp.year}", standard


WORKFLOWS = {
    "REBATE_SLOT": {"no": 1, "label": "Rebate Slot", "kind": "coupon",
                    "competitor": "DAILY REBATE (5%) - 1"},
    "REBATE_BBR": {"no": 2, "label": "Rebate BBR", "kind": "bbr", "expire_days": 30,
                   "reason": "P) Your 5% (Rebate)", "comment": rebate_comment},
    "COSMO_SLOT": {"no": 3, "label": "COSMO ELITE CIRCUIT Slot", "kind": "coupon",
                   "competitor": "COSMO ELITE CIRCUIT - 166"},
    "COSMO_BBR": {"no": 4, "label": "COSMO ELITE CIRCUIT BBR", "kind": "bbr", "expire_days": 3,
                  "reason": "P) COSMO ELITE CIRCUIT", "comment": lambda today: "COSMO ELITE CIRCUIT"},
    "MONTHLY_SLOT": {"no": 5, "label": "Monthly Benefit Slot", "kind": "coupon", "competitor": "MBS FP - 7"},
    "MONTHLY_BBR": {"no": 6, "label": "Monthly Benefit BBR", "kind": "bbr", "expire_days": 14,
                    "reason": "G) MBS FP", "comment": lambda today: monthly_comment(today)[0]},
}
WORKFLOW_ORDER = ["REBATE_SLOT", "REBATE_BBR", "COSMO_SLOT", "COSMO_BBR", "MONTHLY_SLOT", "MONTHLY_BBR"]
STEP_DELAY_SECONDS = 1.0   # pause before every step (checks before/after a step are not shortened)
ACTION_PAUSE_SECONDS = 0.1   # after sending a click
MAX_ATTEMPTS = 3             # tries of one action in total when its result is missing (never the final OK)
RETRY_PAUSE_SECONDS = 1.0    # between two tries
CONTROL_MODE = "background"   # "background": messages + UIA only; "mouse": physical clicks (v2)
# Optional WM_COMMAND ids. When set, the command is posted to the PM main window
# instead of using the Ribbon button / the Options menu.
FIND_PLAYER_COMMAND_ID = None     # None: found automatically from PM's Ctrl+F accelerator
FIND_ACCELERATOR = (ord("F"), "CTRL")
REDEEM_COUPON_COMMAND_ID = None
REDEEM_ACCELERATOR = (0x7B, None)   # F12 opens Coupon Redemption on a profile (recording 09/29)
LEDGER_FILE = BASE_DIR / "pm_redeemed_ledger.csv"      # stays here: it is read by every run of the day
# Each run writes its text log, CSV log, Excel check and screenshots into
# <BASE_DIR>/<yyyymmdd>-automation-log/. After an error they are moved to its "error"
# sub-folder, after an Excel violation to "violation"; the terminal names the folder.
LOG_DIR_SUFFIX = "-automation-log"
ERROR_FOLDER = "error"
VIOLATION_FOLDER = "violation"
LOC_SHEET = "Loc players to check"   # added to pm_excel_check_<time>.xlsx after the run
SCREENSHOTS = True               # save pictures of PM when a try fails and when the run stops
SCREENSHOT_DIR = BASE_DIR / "pm_screenshots"
SCREENSHOT_MARGIN = 40           # pixels around the control in the close-up picture
CLOSE_TABS_AT_END = True
CLOSE_TAB_AFTER_EACH_PLAYER = True   # close each profile tab when the player is finished (also Loc players)
FIND_SHORTCUT = "^f"

# Timeouts (seconds). Measured values from the v5 recordings in comments.
T_LOGIN = 90
T_FIND_OPEN = 3            # 0.21-0.31s
T_FIND_CLOSE = 5           # 0.02-0.16s
T_PROFILE_LOAD = 60        # 3.4-6.6s from OK to the new title; once PM hung 37s (database error)
T_PROFILE_POPUPS = 60      # System Messages + paging through all comments
T_POPUP_CLOSE = 5          # 0.12-0.15s
T_AFTER_OK = 60            # confirmation 0.13-0.16s after OK (once 10.7s); Player Adjustment closes in 0.05-0.7s
T_COMMENT_PAGE = 3         # 0.11-0.2s per Next
T_MENU_OPEN = 3            # 0.15-0.18s
T_COUPON_OPEN = 5          # 0.21-0.22s
T_ADJUST_OPEN = 5          # 0.24-0.27s after Adjust
T_FIELDS_ENABLE = 3        # 0.13-0.24s after Competitor Coupon
T_VERIFY = 3
T_UIA_INVOKE = 10          # UIA call may stay blocked while PM shows the dialog it opened
T_UIA_READ = 4             # reading texts through UIA; never allowed to block a step
T_TAB_CLOSE = 5
QUIET_SECONDS = 1.0        # measured gap between System Messages and Player Comment: 0.04-0.05s
SAME_PLAYER_WAIT_SECONDS = 8.0      # profile load measured 4.1-5.2s
POLL = 0.1
MAX_COMMENT_PAGES = 200

# Native control IDs (identical in every recording).
ID_OK = 1
ID_CANCEL = 2
FIND_PLAYER_ID_EDIT = 1041
SYSMSG_CLOSE = 1
COMMENT_NEXT = 1002
COMMENT_PREVIOUS = 1003
COMMENT_CLOSE = 2
COMMENT_HEADER = 100
OPTIONS_BUTTON = 1074
COUPON_OUR_RADIO = 1481
COUPON_COMPETITOR_RADIO = 1484
COUPON_ID_EDIT = 1477
COUPON_COMPETITOR_COMBO = 1483
COUPON_AMOUNT_EDIT = 1107
COUPON_CLICKABLE_IDS = {COUPON_COMPETITOR_RADIO, COUPON_COMPETITOR_COMBO, COUPON_AMOUNT_EDIT, ID_CANCEL}
REDEEM_MENU_ITEM = "Redeem Coupon..."
REWARDS_FRAME_ID = 3924        # group box "Rewards" on the profile
BBR_RADIO = 2351               # "BBR" in Rewards
ADJUST_BUTTON = 1106           # "Adjust" in Rewards
ADJ_HEADER = 1983              # "BBR Adjustment"
ADJ_ADD_RADIO = 1053           # Add BBR (default)
ADJ_SUBTRACT_RADIO = 1054
ADJ_ZERO_RADIO = 1055          # Set BBR balance to 0
ADJ_CURRENT = 1146             # Current Balance
ADJ_AMOUNT_EDIT = 1147         # Adjustment
ADJ_NEW = 1121                 # New Balance (updates while typing)
ADJ_EXPIRATION = 1031          # date/time picker
ADJ_REASON_COMBO = 1100
ADJ_COMMENT_EDIT = 1004
ADJUST_CLICKABLE_IDS = {ADJ_ADD_RADIO, ID_CANCEL}
IDENT_FRAME_ID = 3923          # group box "Identification" (same in every recording)
IDENT_NAME_ID = 1034           # player name inside Identification
SKIP_NAME_RE = re.compile(r"\(\s*Loc\s*:", re.I)
FIND_RIBBON_BUTTON = "Find Player"

# Popup kinds.
LOGIN, FIND, SYSMSG, COMMENT, COUPON, ADJUST, MENU, DROPDOWN, UNKNOWN = (
    "LOGIN", "FIND", "SYSTEM_MESSAGES", "PLAYER_COMMENT", "COUPON", "ADJUSTMENT", "MENU", "DROPDOWN",
    "UNKNOWN")
# After OK, PM asks again in a second "Coupon Redemption" dialog (owned by the first):
#   1034 "Mr. GEONWOO KIM", 3733 "Redemption Information",
#   1502 "The coupon will reward the player with $645.00 in SLOTS.", OK (1) / Cancel (2)
COUPON_CONFIRM = "COUPON_CONFIRM"
CONFIRM_NAME_ID = 1034
CONFIRM_TEXT_ID = 1502
CONFIRM_TEXT_RE = re.compile(r"reward the player with \$\s*([\d,]+(?:\.\d+)?)\s+in\s+(\w+)", re.I)
CONFIRM_BUCKET = "SLOTS"
DIALOG_TITLES = {
    "find a player": FIND,
    "system messages": SYSMSG,
    "system message": SYSMSG,
    "player comment": COMMENT,
    "coupon redemption": COUPON,
    "player adjustment": ADJUST,
}
LOGIN_TITLE_RE = re.compile(r"^Patron Management\s+(?:Log\s*on|Log\s*in)$", re.I)
MAIN_TITLE_RE = re.compile(r"^Patron Management(?: - .+)?$")
MAIN_CLASSES = {"XTPMainFrame"}
MIN_POPUP_SIZE = 9          # XTP menu shadows are 4px wide windows
IGNORED_CLASS_RE = re.compile(r"tooltip|shadow|PopupBubbleWnd|^IME$|MSCTFIME", re.I)
LOGIN_BUTTON_NAMES = {"login", "log in", "logon", "log on", "sign in", "ok"}
WM_MDIGETACTIVE = 0x0229
BN_CLICKED = 0
BS_TYPEMASK = 0x0F
RADIO_STYLES = {4, 9}         # BS_RADIOBUTTON, BS_AUTORADIOBUTTON
CHECKBOX_STYLES = {2, 3, 5, 6}


def background_mode():
    return CONTROL_MODE == "background"


class StepError(RuntimeError):
    def __init__(self, step, message, hwnd=None):
        super().__init__(message)
        self.step = step
        self.hwnd = hwnd          # control the step was working on (framed in the screenshot)


class PlayerSkipped(Exception):
    """The player must not be processed; logged and the run continues."""

    def __init__(self, status, message):
        super().__init__(message)
        self.status = status


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


def bgra_to_rgb(width, height, raw):
    rgb = bytearray(width * height * 3)
    rgb[0::3], rgb[1::3], rgb[2::3] = raw[2::4], raw[1::4], raw[0::4]
    return rgb


def draw_box(rgb, width, height, box, color=(255, 0, 0), thickness=3):
    """Red frame just outside `box` (left, top, right, bottom in image pixels)."""
    left, top = max(0, box[0] - thickness), max(0, box[1] - thickness)
    right, bottom = min(width, box[2] + thickness), min(height, box[3] + thickness)
    if left >= right or top >= bottom:
        return
    pixel = bytes(color)
    for y in range(top, bottom):
        if y < top + thickness or y >= bottom - thickness:
            xs = range(left, right)
        else:
            xs = list(range(left, min(left + thickness, right))) + list(range(max(right - thickness, left), right))
        for x in xs:
            i = (y * width + x) * 3
            rgb[i:i + 3] = pixel


def crop_rgb(rgb, width, height, box):
    left, top = max(0, box[0]), max(0, box[1])
    right, bottom = min(width, box[2]), min(height, box[3])
    if left >= right or top >= bottom:
        return None
    out = bytearray()
    for y in range(top, bottom):
        start = (y * width + left) * 3
        out += rgb[start:start + (right - left) * 3]
    return right - left, bottom - top, out


def write_png(path, width, height, rgb):
    stride = width * 3
    rows = b"".join(b"\x00" + bytes(rgb[y * stride:(y + 1) * stride]) for y in range(height))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    Path(path).write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
                           + chunk(b"IDAT", zlib.compress(rows, 6)) + chunk(b"IEND", b""))


def region_colors(image, x0, y0, x1, y1):
    """Background (most common) and text (most common clearly different) colour of a region."""
    width, height, raw = image
    counts, rows = {}, []
    for y in range(max(0, y0), min(height, y1)):
        row = y * width * 4
        row_counts = {}
        for x in range(max(0, x0), min(width, x1)):
            i = row + x * 4
            key = (raw[i + 2] & 0xF8, raw[i + 1] & 0xF8, raw[i] & 0xF8)
            row_counts[key] = row_counts.get(key, 0) + 1
        rows.append(row_counts)
        for key, n in row_counts.items():
            counts[key] = counts.get(key, 0) + n
    if not counts:
        return None
    ranked = sorted(counts, key=counts.get, reverse=True)
    background = ranked[0]

    def different(c):
        return sum(abs(a - b) for a, b in zip(c, background)) > 90

    # Text pixels are counted without rows that are one solid line (borders, underlines).
    text_counts = {}
    for row_counts in rows:
        total = sum(row_counts.values())
        line = max(row_counts.values()) >= 0.7 * total and max(row_counts, key=row_counts.get) != background
        if line:
            continue
        for key, n in row_counts.items():
            if different(key):
                text_counts[key] = text_counts.get(key, 0) + n
    text = max(text_counts, key=text_counts.get) if text_counts else None
    return {"text_color": color_name(text) if text else "none", "background_color": color_name(background),
            "text_rgb": "#%02x%02x%02x" % text if text else "", "background_rgb": "#%02x%02x%02x" % background}


RUN_FILES = {}      # dir, stamp, check, problems of the current run (set by start_run_files)


def start_run_files(now=None):
    """Point this run's text log, CSV log and screenshots at <yyyymmdd>-automation-log/."""
    global LOG_FILE, TEXT_LOG_FILE
    now = now or datetime.now()
    day_dir = BASE_DIR / f"{now:%Y%m%d}{LOG_DIR_SUFFIX}"
    day_dir.mkdir(parents=True, exist_ok=True)
    stamp = f"{now:%Y%m%d_%H%M%S}"
    RUN_FILES.clear()
    RUN_FILES.update(dir=day_dir, stamp=stamp, problems=[])
    TEXT_LOG_FILE = day_dir / f"pm_runner_log_{stamp}.txt"
    LOG_FILE = day_dir / f"pm_automation_log_{stamp}.csv"
    _screens["dir"] = day_dir / f"pm_screenshots_{stamp}"
    _screens["count"] = 0
    return day_dir


def note_problem(text):
    """One line for the error / violation summary printed at the end of the run."""
    RUN_FILES.setdefault("problems", []).append(text)


def loc_line(row):
    wf = WORKFLOWS.get(row["workflow"], {})
    details = [f"amount {row['amount']}", str(row["target"])]
    if row.get("expiration"):
        details.append(f"expires {row['expiration']}")
    if row.get("comment"):
        details.append(f"comment '{row['comment']}'")
    return (f"[{row['order']}] #{wf.get('no', '?')} {wf.get('label', row['workflow'])}: player "
            f"{row['player_id']} ({row['sheet']} row {row['excel_row']}), {', '.join(details)} "
            f"- name '{row.get('identification', '')}'")


def add_loc_sheet(path, rows, stopped_at=None):
    """Sheet LOC_SHEET in the Excel check file: players skipped for '(Loc:', to be done by hand."""
    wb = load_workbook(path)
    if LOC_SHEET in wb.sheetnames:
        del wb[LOC_SHEET]
    sheet = wb.create_sheet(LOC_SHEET, 0 if rows else len(wb.sheetnames))
    sheet.append(["Order", "Workflow", "Sheet", "Excel row", "Player ID", "Amount", "Competitor / Reason",
                  "Expiration", "Comment", "Identification name", "Name colour", "Checked at", "Done by hand"])
    for row in rows:
        wf = WORKFLOWS.get(row["workflow"], {})
        sheet.append([row["order"], f"#{wf.get('no', '?')} {wf.get('label', row['workflow'])}", row["sheet"],
                      row["excel_row"], row["player_id"], row["amount"], row["target"], row["expiration"],
                      row["comment"], row.get("identification", ""), row.get("name_color", ""),
                      row["timestamp"], ""])
    if not rows:
        sheet.append(["No player with '(Loc:' in the Identification name in this run."])
    if stopped_at:
        sheet.append([])
        sheet.append([f"The run stopped at #{stopped_at}: players after it were not checked for '(Loc:'."])
    if rows:
        wb.active = 0
    wb.save(path)


def finish_run_files(kind):
    """Print the result; after an error / violation move this run's files into that sub-folder.

    kind: None (no error), ERROR_FOLDER or VIOLATION_FOLDER. Returns the folder to look at.
    """
    global LOG_FILE, TEXT_LOG_FILE
    day_dir = RUN_FILES.get("dir")
    if day_dir is None:
        return None
    problems = RUN_FILES.get("problems") or []
    loc_rows = RUN_FILES.get("loc")
    check = RUN_FILES.get("check")
    bar = "=" * 78
    target = day_dir / kind if kind else day_dir
    if loc_rows is not None and check and Path(check).exists():
        try:
            add_loc_sheet(check, loc_rows, RUN_FILES.get("stopped_at"))
        except Exception as exc:
            log(f"Could not add the sheet '{LOC_SHEET}' to {Path(check).name}: {exc}")
    if "plan" in RUN_FILES and check and Path(check).exists():
        try:
            jobs, notes, violations = RUN_FILES["plan"]
            add_status_sheet(check, RUN_FILES.get("sources", {}),
                             (jobs, notes, violations, RUN_FILES.get("results", [])))
        except Exception as exc:
            log(f"Could not add the sheet '{STATUS_SHEET}' to {Path(check).name}: {exc}")
    log(bar)
    if kind is None:
        log(f"NO ERROR. Logs: {target}")
        for line in problems:
            log(f"  note: {line}")
    else:
        log(f"{'ERROR' if kind == ERROR_FOLDER else 'VIOLATION in the Excel file'} - please check: {target}")
        for line in problems or ["see the log file in that folder"]:
            log(f"  {line}")
    if check and "plan" in RUN_FILES:
        log(f"Status of every Excel row (DONE / NOT DONE + reason): sheet '{STATUS_SHEET}' in "
            f"{target / Path(check).name}")
    if loc_rows:
        log("-" * 78)
        log(f"LOC PLAYERS - NOT PROCESSED, please do them by hand ({len(loc_rows)}): "
            f"sheet '{LOC_SHEET}' in {target / Path(check).name if check else 'the Excel check file'}")
        for row in loc_rows:
            log(f"  {loc_line(row)}")
    log(bar)
    if kind is None:
        return target
    target.mkdir(parents=True, exist_ok=True)
    moved = {}
    for path in (TEXT_LOG_FILE, LOG_FILE, RUN_FILES.get("check"), _screens.get("dir")):
        if not path or not Path(path).exists():
            continue
        try:
            destination = target / Path(path).name
            shutil.move(str(path), str(destination))
            moved[str(path)] = destination
        except OSError as exc:
            print(f"  (could not move {Path(path).name} into {target}: {exc})")
    TEXT_LOG_FILE = moved.get(str(TEXT_LOG_FILE), TEXT_LOG_FILE)
    LOG_FILE = moved.get(str(LOG_FILE), LOG_FILE)
    return target


def log(message):
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
    print(line, flush=True)
    try:
        with TEXT_LOG_FILE.open("a", encoding="utf-8") as file:
            file.write(line + "\n")
    except OSError:
        pass


# ------------------------------------------------------------------ Excel

def original_excel_text(value):
    """Preserve Excel's value representation without padding or recalculation."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else format(value, "g")
    return str(value).strip()


def numeric_nonzero(value):
    if value is None or value == "":
        return False
    try:
        return float(value) != 0
    except (TypeError, ValueError):
        raise ValueError(f"Invalid rebate value: {value!r}")


def sheet_rows(wb, sheet_name, columns):
    """(excel_row, values...) for every non-empty row of a sheet with the given headers."""
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"Sheet not found: {sheet_name}")
    rows = wb[sheet_name].iter_rows(values_only=True)
    headers = next(rows, None) or ()
    normalized = [str(x).strip().casefold() if x is not None else "" for x in headers]
    try:
        indexes = [normalized.index(c.casefold()) for c in columns]
    except ValueError as exc:
        raise ValueError(f"Sheet {sheet_name}: required columns {columns} not found") from exc
    for excel_row, row in enumerate(rows, start=2):
        values = [row[i] if i < len(row) else None for i in indexes]
        if all(v is None or str(v).strip() == "" for v in values):
            continue
        yield excel_row, values


def parse_amount(value):
    if value is None or str(value).strip() == "":
        return 0.0
    return float(value)


def job_details(workflow, today):
    wf = WORKFLOWS[workflow]
    if wf["kind"] == "coupon":
        return {"target": wf["competitor"], "expiration": "", "comment": ""}
    expires = expiration_for(workflow, today)
    return {"target": wf["reason"], "expiration": expires.strftime("%m/%d/%Y %I:%M %p"),
            "comment": wf["comment"](today)}


def expiration_for(workflow, today):
    day = today + timedelta(days=WORKFLOWS[workflow]["expire_days"])
    return datetime(day.year, day.month, day.day, *EXPIRATION_TIME)


def read_plan(today):
    """Jobs in run order, notes (skipped rows) and violations (stop the run)."""
    jobs = {w: [] for w in WORKFLOW_ORDER}
    notes, violations = [], []
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")
    wb = load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    try:
        for row, (pid_value, slot_value, bbr_value) in sheet_rows(wb, REBATE_SHEET, REBATE_COLUMNS):
            pid = original_excel_text(pid_value)
            try:
                slot, bbr = parse_amount(slot_value), parse_amount(bbr_value)
            except (TypeError, ValueError):
                violations.append({"rule": "Invalid amount", "sheet": REBATE_SHEET, "rows": str(row),
                                   "player_id": pid, "detail": f"SLOT={slot_value!r} BBR={bbr_value!r}"})
                continue
            if not pid:
                violations.append({"rule": "Missing Player ID", "sheet": REBATE_SHEET, "rows": str(row),
                                   "player_id": "", "detail": f"SLOT={slot_value!r} BBR={bbr_value!r}"})
                continue
            if slot < 0 or bbr < 0:
                violations.append({"rule": "Negative amount", "sheet": REBATE_SHEET, "rows": str(row),
                                   "player_id": pid, "detail": f"SLOT={slot_value} BBR={bbr_value}"})
                continue
            if slot == 0 and bbr == 0:
                notes.append({"sheet": REBATE_SHEET, "row": row, "player_id": pid,
                              "note": "SLOT and BBR are both 0 - skipped"})
                continue
            if slot:
                jobs["REBATE_SLOT"].append({"workflow": "REBATE_SLOT", "sheet": REBATE_SHEET, "excel_row": row,
                                            "player_id": pid, "amount": original_excel_text(slot_value)})
            if bbr:
                jobs["REBATE_BBR"].append({"workflow": "REBATE_BBR", "sheet": REBATE_SHEET, "excel_row": row,
                                           "player_id": pid, "amount": original_excel_text(bbr_value)})
        for row, (pid_value, amount_value, allocation_value) in sheet_rows(wb, DAILY_SHEET, DAILY_COLUMNS):
            pid = original_excel_text(pid_value)
            allocation = str(allocation_value or "").strip().casefold()
            try:
                amount = parse_amount(amount_value)
            except (TypeError, ValueError):
                violations.append({"rule": "Invalid amount", "sheet": DAILY_SHEET, "rows": str(row),
                                   "player_id": pid, "detail": f"FREE PLAY={amount_value!r}"})
                continue
            if not pid:
                violations.append({"rule": "Missing Player ID", "sheet": DAILY_SHEET, "rows": str(row),
                                   "player_id": "", "detail": f"FREE PLAY={amount_value!r}"})
                continue
            if allocation not in ("slot", "bbr"):
                violations.append({"rule": "Unknown ALLOCATION", "sheet": DAILY_SHEET, "rows": str(row),
                                   "player_id": pid, "detail": f"ALLOCATION={allocation_value!r}"})
                continue
            if amount < 0:
                violations.append({"rule": "Negative amount", "sheet": DAILY_SHEET, "rows": str(row),
                                   "player_id": pid, "detail": f"FREE PLAY={amount_value}"})
                continue
            if amount == 0:
                notes.append({"sheet": DAILY_SHEET, "row": row, "player_id": pid,
                              "note": f"FREE PLAY is 0 ({allocation_value}) - skipped"})
                continue
            workflow = "COSMO_SLOT" if allocation == "slot" else "COSMO_BBR"
            jobs[workflow].append({"workflow": workflow, "sheet": DAILY_SHEET, "excel_row": row,
                                   "player_id": pid, "amount": original_excel_text(amount_value)})
        if MONTHLY_SHEET not in wb.sheetnames:
            notes.append({"sheet": MONTHLY_SHEET, "row": "", "player_id": "",
                          "note": f"sheet '{MONTHLY_SHEET}' not found - nothing to do for Monthly Benefit"})
        else:
            for row, (pid_value, amount_value, allocation_value) in sheet_rows(wb, MONTHLY_SHEET, MONTHLY_COLUMNS):
                pid = original_excel_text(pid_value)
                allocation = " ".join(str(allocation_value or "").split()).casefold()
                try:
                    amount = parse_amount(amount_value)
                except (TypeError, ValueError):
                    violations.append({"rule": "Invalid amount", "sheet": MONTHLY_SHEET, "rows": str(row),
                                       "player_id": pid, "detail": f"MONTHLY FP check={amount_value!r}"})
                    continue
                if not pid:
                    violations.append({"rule": "Missing Player ID", "sheet": MONTHLY_SHEET, "rows": str(row),
                                       "player_id": "", "detail": f"MONTHLY FP check={amount_value!r}"})
                    continue
                if allocation not in MONTHLY_ALLOCATIONS:
                    violations.append({"rule": "Unknown FP ALLOCATION", "sheet": MONTHLY_SHEET, "rows": str(row),
                                       "player_id": pid, "detail": f"FP ALLOCATION={allocation_value!r}"})
                    continue
                if amount < 0:
                    violations.append({"rule": "Negative amount", "sheet": MONTHLY_SHEET, "rows": str(row),
                                       "player_id": pid, "detail": f"MONTHLY FP check={amount_value}"})
                    continue
                if amount == 0:
                    notes.append({"sheet": MONTHLY_SHEET, "row": row, "player_id": pid,
                                  "note": f"MONTHLY FP check is 0 ({allocation_value}) - skipped"})
                    continue
                workflow = MONTHLY_ALLOCATIONS[allocation]
                jobs[workflow].append({"workflow": workflow, "sheet": MONTHLY_SHEET, "excel_row": row,
                                       "player_id": pid, "amount": original_excel_text(amount_value)})
    finally:
        wb.close()

    def rows_of(workflow, pid):
        return [j["excel_row"] for j in jobs[workflow] if j["player_id"] == pid]

    for first, second, sheet in (("REBATE_SLOT", "REBATE_BBR", REBATE_SHEET),
                                 ("COSMO_SLOT", "COSMO_BBR", DAILY_SHEET),
                                 ("MONTHLY_SLOT", "MONTHLY_BBR", MONTHLY_SHEET)):
        for workflow in (first, second):
            seen = {}
            for job in jobs[workflow]:
                seen.setdefault(job["player_id"], []).append(job["excel_row"])
            for pid, rows in seen.items():
                if len(rows) > 1:
                    record = {"sheet": sheet, "rows": ", ".join(map(str, rows)), "player_id": pid,
                              "detail": f"{WORKFLOWS[workflow]['label']} appears {len(rows)} times"}
                    if pid in TEST_PLAYER_IDS:
                        notes.append({"sheet": sheet, "row": record["rows"], "player_id": pid,
                                      "note": record["detail"] + " - allowed for the test player"})
                    else:
                        violations.append(dict(record, rule="Duplicate player"))
        both = sorted({j["player_id"] for j in jobs[first]} & {j["player_id"] for j in jobs[second]})
        for pid in both:
            detail = (f"{WORKFLOWS[first]['label']} rows {rows_of(first, pid)} and "
                      f"{WORKFLOWS[second]['label']} rows {rows_of(second, pid)}")
            if pid in TEST_PLAYER_IDS:
                notes.append({"sheet": sheet, "row": "", "player_id": pid,
                              "note": detail + " - allowed for the test player"})
            else:
                rows = sorted(set(rows_of(first, pid) + rows_of(second, pid)))
                if len(rows) == 1:
                    detail += " (same row)"
                violations.append({"rule": "Slot and BBR for the same player", "sheet": sheet,
                                   "rows": ", ".join(map(str, rows)), "player_id": pid, "detail": detail})
    ordered = [job for workflow in WORKFLOW_ORDER for job in jobs[workflow]]
    for job in ordered:
        job.update(job_details(job["workflow"], today))
    return ordered, notes, violations


def read_source_tables():
    """Every non-empty row of the three sheets, all columns as in Excel: {sheet: (headers, [(row, values)])}."""
    tables = {}
    if not EXCEL_FILE.exists():
        return tables
    wb = load_workbook(EXCEL_FILE, data_only=True, read_only=True)
    try:
        for name in SOURCE_SHEETS:
            if name not in wb.sheetnames:
                continue
            rows = wb[name].iter_rows(values_only=True)
            headers = list(next(rows, None) or ())
            while headers and (headers[-1] is None or str(headers[-1]).strip() == ""):
                headers.pop()
            data = []
            for excel_row, row in enumerate(rows, start=2):
                values = [row[i] if i < len(row) else None for i in range(len(headers))]
                if any(v is not None and str(v).strip() != "" for v in values):
                    data.append((excel_row, values))
            tables[name] = (headers, data)
    finally:
        wb.close()
    return tables


STATUS_SHEET = "STATUS"
STATUS_FILLS = {"DONE": "C6EFCE", "NOT DONE": "FFC7CE", "CHECK IN PM": "FFEB9C", "NOTHING TO DO": "EDEDED"}


def row_status(sheet, excel_row, plan):
    """(status, note) of one Excel row, from the plan and the results of this run."""
    jobs, notes, violations, results = plan
    for v in violations:
        if v["sheet"] == sheet and str(excel_row) in [x.strip() for x in str(v["rows"]).split(",")]:
            return "NOT DONE", f"Excel violation: {v['rule']} - {v['detail']}. The run stopped before PM."
    row_jobs = [j for j in jobs if j["sheet"] == sheet and j["excel_row"] == excel_row]
    if not row_jobs:
        skipped = [n["note"] for n in notes if n["sheet"] == sheet and str(n["row"]) == str(excel_row)]
        return "NOTHING TO DO", "; ".join(skipped) or "nothing to issue for this row"
    parts, statuses = [], []
    for job in row_jobs:
        wf = WORKFLOWS[job["workflow"]]
        what = f"{'Slot coupon' if wf['kind'] == 'coupon' else 'BBR'} {job['amount']}"
        record = next((r for r in results if r["sheet"] == sheet and r["excel_row"] == excel_row
                       and r["workflow"] == job["workflow"]), None)
        status = record["status"] if record else ""
        if status == "DONE":
            check = record.get("confirmation_check", "")
            statuses.append("DONE")
            parts.append(f"{what}: done" + (f" (confirmation note: {check})" if check not in ("", "OK") else ""))
        elif status == "SKIPPED_ALREADY_DONE":
            statuses.append("DONE")
            parts.append(f"{what}: done earlier today (pm_redeemed_ledger.csv)")
        elif status == "TEST_CANCELLED":
            statuses.append("NOT DONE")
            parts.append(f"{what}: test run only - filled and cancelled, no OK")
        elif status == "SKIPPED_LOC":
            statuses.append("NOT DONE")
            parts.append(f"{what}: Loc player '{record.get('identification', '')}' - do it by hand")
        elif status == "ERROR_AFTER_OK":
            statuses.append("CHECK IN PM")
            parts.append(f"{what}: OK was clicked, then an error at {record['failed_step']} - check in PM: "
                         f"{record['message'].split(' | screen:')[0][:200]}")
        elif status.startswith("ERROR"):
            statuses.append("NOT DONE")
            parts.append(f"{what}: error at {record['failed_step']}: {record['message'].split(' | screen:')[0][:200]}")
        else:
            statuses.append("NOT DONE")
            stopped = RUN_FILES.get("stopped_at")
            parts.append(f"{what}: not processed" + (f" - the run stopped at #{stopped}" if stopped
                                                     else " - the run did not reach this row"))
    for status in ("CHECK IN PM", "NOT DONE", "DONE"):
        if status in statuses:
            return status, " | ".join(parts)
    return "NOT DONE", " | ".join(parts)


def add_status_sheet(path, tables, plan):
    """Sheet STATUS: REBATE / DAILY REWARDS / MONTHLY BNF side by side (2 empty columns between them),
    every Excel row with all its columns plus Status (DONE / NOT DONE / CHECK IN PM / NOTHING TO DO) and Note."""
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = load_workbook(path)
    if STATUS_SHEET in wb.sheetnames:
        del wb[STATUS_SHEET]
    sheet = wb.create_sheet(STATUS_SHEET, 0)
    bold = Font(bold=True)
    column = 1
    for name in SOURCE_SHEETS:
        headers, data = tables.get(name, ([], []))
        headers = ["Excel row"] + [str(h) if h is not None else "" for h in headers] + ["Status", "Note"]
        counts = {}
        rows = []
        for excel_row, values in data:
            status, note = row_status(name, excel_row, plan)
            counts[status] = counts.get(status, 0) + 1
            rows.append([excel_row] + list(values) + [status, note])
        title = f"{name}: " + (", ".join(f"{n} {s}" for s, n in counts.items()) if name in tables
                               else "sheet not found")
        sheet.cell(row=1, column=column, value=title).font = bold
        for offset, header in enumerate(headers):
            sheet.cell(row=2, column=column + offset, value=header).font = bold
        for r, values in enumerate(rows, start=3):
            for offset, value in enumerate(values):
                sheet.cell(row=r, column=column + offset, value=value)
            fill = STATUS_FILLS.get(values[-2])
            if fill:
                for offset in (len(values) - 2, len(values) - 1):
                    sheet.cell(row=r, column=column + offset).fill = PatternFill("solid", fgColor=fill)
        for offset, header in enumerate(headers):
            width = 60 if header == "Note" else max(10, min(24, len(header) + 2))
            sheet.column_dimensions[get_column_letter(column + offset)].width = width
        column += len(headers) + 2          # two empty columns between the tables
    sheet.freeze_panes = "A3"
    wb.active = 0
    wb.save(path)


def write_check_file(jobs, notes, violations):
    if RUN_FILES.get("dir"):
        path = RUN_FILES["dir"] / f"pm_excel_check_{RUN_FILES['stamp']}.xlsx"
        RUN_FILES["check"] = path
    else:
        path = BASE_DIR / f"pm_excel_check_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    wb = Workbook()
    plan = wb.active
    plan.title = "PLAN"
    plan.append(["Order", "Workflow", "Sheet", "Excel row", "Player ID", "Amount", "Competitor / Reason",
                 "Expiration", "Comment"])
    for order, job in enumerate(jobs, start=1):
        wf = WORKFLOWS[job["workflow"]]
        plan.append([order, f"#{wf['no']} {wf['label']}", job["sheet"], job["excel_row"], job["player_id"],
                     job["amount"], job["target"], job["expiration"], job["comment"]])
    sheet = wb.create_sheet("NOTES")
    sheet.append(["Sheet", "Excel row", "Player ID", "Note"])
    for note in notes:
        sheet.append([note["sheet"], note["row"], note["player_id"], note["note"]])
    sheet = wb.create_sheet("VIOLATIONS")
    sheet.append(["Rule", "Sheet", "Excel rows", "Player ID", "Detail"])
    for v in violations:
        sheet.append([v["rule"], v["sheet"], v["rows"], v["player_id"], v["detail"]])
    if violations:
        wb.move_sheet("VIOLATIONS", offset=-2)
        wb.active = 0
    wb.save(path)
    return path


# ------------------------------------------------------------ PM state

@dataclass
class Popup:
    hwnd: int
    title: str
    class_name: str
    kind: str

    def label(self):
        return f"{self.kind} '{self.title}' #{self.hwnd}"


@dataclass
class PMState:
    title: str
    main_enabled: bool
    popups: list

    def kinds(self):
        return sorted(p.kind for p in self.popups)

    def get(self, kind):
        for popup in self.popups:
            if popup.kind == kind:
                return popup
        return None

    def summary(self):
        popups = ", ".join(p.label() for p in self.popups) or "none"
        return f"title='{self.title}' main_enabled={self.main_enabled} popups=[{popups}]"


def classify(title, class_name):
    title = (title or "").strip()
    cls = (class_name or "").casefold()
    if LOGIN_TITLE_RE.match(title):
        return LOGIN
    if cls == "#32770":
        return DIALOG_TITLES.get(title.casefold(), UNKNOWN)
    if "popupbar" in cls or cls == "#32768":
        return MENU
    if cls == "combolbox":
        return DROPDOWN
    return UNKNOWN


def ids_in_title(title):
    return [x.lstrip("0") or "0" for x in re.findall(r"\((\d+)\)", title or "")]


def title_has_player(title, player_id):
    expected = str(player_id).lstrip("0") or "0"
    return expected in ids_in_title(title)


def normalized_number_text(value):
    value = str(value).strip().replace(",", "")
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else format(number, "g")
    except (TypeError, ValueError):
        return value


def escape_keys(text):
    return "".join("{%s}" % ch if ch in "{}[]()+^%~" else ch for ch in str(text))


# ------------------------------------------------ Win32 access to PM

class Win32PM:
    """Everything that touches the PM process lives here."""

    def __init__(self):
        if win32gui is None:
            raise SystemExit("This runner needs Windows with pywin32 and pywinauto installed.")
        self.pid = None
        self.main = None
        self._user32 = ctypes.windll.user32
        self._user32.GetAncestor.restype = ctypes.c_void_p
        self._user32.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]

    # -- discovery

    def connect(self):
        found = []

        def callback(hwnd, _):
            try:
                if (win32gui.IsWindowVisible(hwnd)
                        and win32gui.GetClassName(hwnd) in MAIN_CLASSES
                        and MAIN_TITLE_RE.match(win32gui.GetWindowText(hwnd))):
                    found.append(hwnd)
            except win32gui.error:
                pass
            return True

        win32gui.EnumWindows(callback, None)
        if not found:
            raise StepError("connect", "Patron Management main window not found. Open PM first.")
        if len(found) > 1:
            raise StepError("connect", "More than one Patron Management window is open.")
        self.main = found[0]
        self.pid = win32process.GetWindowThreadProcessId(self.main)[1]
        log(f"Connected to PM: main HWND={self.main}, PID={self.pid}")
        self.find_command = FIND_PLAYER_COMMAND_ID
        if not self.find_command:
            try:
                self.find_command = self._accelerator_command(*FIND_ACCELERATOR)
            except Exception as exc:
                log(f"Could not read PM's shortcut table: {exc}")
        log("Find a Player: " + (f"command id {self.find_command} (the Ctrl+F command, sent without the keyboard)"
                                   if self.find_command else "Ribbon button through UIA"))
        self.redeem_command = REDEEM_COUPON_COMMAND_ID
        if not self.redeem_command:
            try:
                self.redeem_command = self._accelerator_command(*REDEEM_ACCELERATOR)
            except Exception as exc:
                log(f"Could not read PM's shortcut table: {exc}")
        log("Redeem Coupon: " + (f"command id {self.redeem_command} (the F12 command, sent without the keyboard)"
                                   if self.redeem_command else f"Options... > {REDEEM_MENU_ITEM}"))

    def _accelerator_command(self, key, modifier):
        """Command ID PM binds to a shortcut such as Ctrl+F, read from its accelerator tables."""
        flags_wanted = 0x01 | ({"CTRL": 0x08, "SHIFT": 0x04, "ALT": 0x10}[modifier] if modifier else 0)
        process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
                                       False, self.pid)
        exe = win32process.GetModuleFileNameEx(process, None)
        folder = str(Path(exe).parent).casefold()
        modules = [exe] + [m for m in (win32process.GetModuleFileNameEx(process, h)
                                        for h in win32process.EnumProcessModules(process))
                           if str(Path(m).parent).casefold() == folder and m != exe]
        commands = {}
        for path in modules:
            try:
                module = win32api.LoadLibraryEx(path, 0, win32con.LOAD_LIBRARY_AS_DATAFILE)
            except win32api.error:
                continue
            try:
                try:
                    names = win32api.EnumResourceNames(module, win32con.RT_ACCELERATOR)
                except win32api.error:
                    names = []
                for name in names:
                    data = win32api.LoadResource(module, win32con.RT_ACCELERATOR, name)
                    for offset in range(0, len(data) - 7, 8):
                        flags, vk, command, _ = struct.unpack_from("<HHHH", data, offset)
                        if flags & 0x1D == flags_wanted and vk == key:
                            commands.setdefault(command, []).append(f"{Path(path).name}#{name}")
                        if flags & 0x80:
                            break
            finally:
                win32api.FreeLibrary(module)
        if len(commands) == 1:
            return next(iter(commands))
        if commands:
            label = f"F{key - 0x6F}" if 0x70 <= key <= 0x87 else chr(key)
            log(f"{modifier + '+' if modifier else ''}{label} maps to several commands {commands}; not used.")
        return None

    def main_title(self):
        return self.window_title(self.main)

    def main_enabled(self):
        return bool(win32gui.IsWindowEnabled(self.main))

    def popups(self):
        found = []

        def callback(hwnd, _):
            try:
                if hwnd == self.main or not win32gui.IsWindowVisible(hwnd):
                    return True
                if win32process.GetWindowThreadProcessId(hwnd)[1] != self.pid:
                    return True
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                if right - left < MIN_POPUP_SIZE or bottom - top < MIN_POPUP_SIZE:
                    return True
                cls = win32gui.GetClassName(hwnd)
                if IGNORED_CLASS_RE.search(cls):
                    return True
                title = win32gui.GetWindowText(hwnd)
                kind = classify(title, cls)
                if kind == COUPON and self.child(hwnd, CONFIRM_TEXT_ID):
                    kind = COUPON_CONFIRM
                found.append(Popup(hwnd, title, cls, kind))
            except win32gui.error:
                pass
            return True

        win32gui.EnumWindows(callback, None)
        return found

    # -- windows and controls

    def exists(self, hwnd):
        return bool(hwnd) and bool(win32gui.IsWindow(hwnd))

    def visible(self, hwnd):
        return self.exists(hwnd) and bool(win32gui.IsWindowVisible(hwnd))

    def enabled(self, hwnd):
        return self.visible(hwnd) and bool(win32gui.IsWindowEnabled(hwnd))

    def window_title(self, hwnd):
        try:
            return win32gui.GetWindowText(hwnd)
        except win32gui.error:
            return ""

    def child(self, parent, control_id):
        found = []

        def callback(hwnd, _):
            try:
                if win32gui.GetDlgCtrlID(hwnd) == control_id:
                    found.append(hwnd)
            except win32gui.error:
                pass
            return True

        try:
            win32gui.EnumChildWindows(parent, callback, None)
        except win32gui.error:
            pass
        visible = [h for h in found if win32gui.IsWindowVisible(h)]
        return (visible or found or [None])[0]

    def text(self, hwnd):
        try:
            return handleprops.text(hwnd) or ""
        except Exception:
            return ""

    def checked(self, hwnd):
        try:
            _, value = win32gui.SendMessageTimeout(
                hwnd, win32con.BM_GETCHECK, 0, 0, win32con.SMTO_ABORTIFHUNG, 2000)
            return value == win32con.BST_CHECKED
        except win32gui.error:
            return False

    def _top(self, hwnd):
        return int(self._user32.GetAncestor(hwnd, 2) or hwnd)

    def _no_input(self, what):
        raise StepError("background", f"{what} needs the mouse/keyboard; background mode never uses them.")

    def ensure_shown(self):
        """Owned dialogs are hidden while PM is minimized: restore it without activating."""
        if win32gui.IsIconic(self.main):
            win32gui.ShowWindow(self.main, win32con.SW_SHOWNOACTIVATE)
            log("PM was minimized: restored without taking focus.")
            time.sleep(0.5)

    def focus_in_dialog(self, dialog, control):
        """Move the dialog's own focus to a control (WM_NEXTDLGCTL), as a real click on it does.

        Only PM's focus inside its dialog changes; the foreground window, the keyboard and the
        mouse of the user do not. Some dialogs treat OK like Enter in a field ("go to the next
        field") unless the OK button has the focus; with it, OK is handled as a real OK.
        """
        if background_mode():
            self._send(dialog, win32con.WM_NEXTDLGCTL, control, 1)

    def focus(self, hwnd):
        if background_mode():
            return
        try:
            HwndWrapper(hwnd).set_focus()
        except Exception:
            pass

    def _style(self, hwnd):
        return win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)

    def _siblings(self, hwnd):
        parent = win32gui.GetParent(hwnd)
        out = []
        child = win32gui.GetWindow(parent, win32con.GW_CHILD)
        while child:
            out.append(child)
            child = win32gui.GetWindow(child, win32con.GW_HWNDNEXT)
        return out

    def _radio_group(self, hwnd):
        """Radio buttons of hwnd's group (a group starts at a control with WS_GROUP)."""
        siblings = self._siblings(hwnd)
        index = siblings.index(hwnd)
        start = index
        while start > 0 and not self._style(siblings[start]) & win32con.WS_GROUP:
            start -= 1
        end = index + 1
        while end < len(siblings) and not self._style(siblings[end]) & win32con.WS_GROUP:
            end += 1
        return [h for h in siblings[start:end]
                if win32gui.GetClassName(h).casefold() == "button"
                and self._style(h) & BS_TYPEMASK in RADIO_STYLES]

    def _send(self, hwnd, message, wparam=0, lparam=0):
        return win32gui.SendMessageTimeout(hwnd, message, wparam, lparam,
                                           win32con.SMTO_ABORTIFHUNG, 3000)[1]

    def click(self, hwnd):
        """Background: tell the dialog the button was clicked (WM_COMMAND/BN_CLICKED).

        Posted, never sent, so a handler that opens a modal dialog cannot block us.
        Mouse mode: physical click in the middle of the control (v2 behaviour).
        """
        if not background_mode():
            self.focus(self._top(hwnd))
            HwndWrapper(hwnd).click_input()
            time.sleep(ACTION_PAUSE_SECONDS)
            return
        cls = win32gui.GetClassName(hwnd).casefold()
        if "button" not in cls:
            self._no_input(f"Clicking a '{cls}' control")
        button_type = self._style(hwnd) & BS_TYPEMASK
        if button_type in RADIO_STYLES:
            for radio in self._radio_group(hwnd):
                self._send(radio, win32con.BM_SETCHECK,
                           win32con.BST_CHECKED if radio == hwnd else win32con.BST_UNCHECKED)
        elif button_type in CHECKBOX_STYLES:
            checked = self.checked(hwnd)
            self._send(hwnd, win32con.BM_SETCHECK,
                       win32con.BST_UNCHECKED if checked else win32con.BST_CHECKED)
        control_id = win32gui.GetDlgCtrlID(hwnd) & 0xFFFF
        win32gui.PostMessage(win32gui.GetParent(hwnd), win32con.WM_COMMAND,
                             (BN_CLICKED << 16) | control_id, hwnd)
        time.sleep(ACTION_PAUSE_SECONDS)

    def set_text(self, hwnd, value):
        HwndWrapper(hwnd).set_edit_text(str(value))    # EM_SETSEL + EM_REPLACESEL, no focus

    def type_text(self, hwnd, value):
        if background_mode():
            self._no_input("Typing into a field")
        self.click(hwnd)
        send_keys("{HOME}+{END}{DEL}")
        send_keys(escape_keys(value), with_spaces=True)

    def post_command(self, command_id, accelerator=False):
        """Post WM_COMMAND to the main window (as a menu, or as the shortcut when accelerator=True)."""
        win32gui.PostMessage(self.main, win32con.WM_COMMAND,
                             ((1 if accelerator else 0) << 16) | (int(command_id) & 0xFFFF), 0)

    def rect(self, hwnd):
        return win32gui.GetWindowRect(hwnd)

    def date_set(self, hwnd, value):
        """Set a date/time picker with DTM_SETSYSTEMTIME (no mouse); the SYSTEMTIME goes through PM's memory.

        No WM_NOTIFY is sent to the dialog: Windows refuses WM_NOTIFY between processes
        ("Access is denied"). The dialog reads the value from the picker when OK is clicked.
        """
        HwndWrapper(hwnd).set_time(year=value.year, month=value.month, day_of_week=value.isoweekday() % 7,
                                   day=value.day, hour=value.hour, minute=value.minute)

    def date_get(self, hwnd):
        st = HwndWrapper(hwnd).get_time()
        return datetime(st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute)

    def _uia_read(self, read, what, timeout=T_UIA_READ):
        """Run a UIA read in a worker thread; give up after `timeout` instead of blocking the run.

        A UIA call can wait behind another UIA call PM has not answered yet (for example an
        Invoke that opened a modal dialog), so a read must never hold up the next step.
        """
        outcome = {}

        def work():
            try:
                import pythoncom
                pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            except Exception:
                pass
            try:
                outcome["value"] = read()
            except Exception as exc:
                outcome["error"] = exc

        worker = threading.Thread(target=work, daemon=True)
        worker.start()
        worker.join(timeout)
        if "value" in outcome:
            return outcome["value"]
        reason = outcome.get("error") or f"no answer within {timeout}s"
        log(f"    ({what} not readable through UIA: {reason}; continuing without it)")
        return None

    def list_items_levels(self, popup_hwnd):
        """(text, left) of each list item, e.g. System Messages; left gives the indent level."""
        items = self._uia_read(lambda: [(i.window_text().strip(), i.rectangle().left)
                                        for i in self._uia(popup_hwnd).descendants(control_type="ListItem")
                                        if i.window_text().strip()], "System Messages text")
        return items or []

    def _capture(self, window):
        """PM's own rendering of a top-level window (PrintWindow: no focus, works when covered).

        Returns (width, height, BGRA rows top-down, screen left, screen top).
        """
        user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
        vp = ctypes.c_void_p
        for fn, args, res in ((user32.GetWindowDC, [vp], vp), (user32.ReleaseDC, [vp, vp], ctypes.c_int),
                              (user32.PrintWindow, [vp, vp, ctypes.c_uint], ctypes.c_int),
                              (gdi32.CreateCompatibleDC, [vp], vp),
                              (gdi32.CreateCompatibleBitmap, [vp, ctypes.c_int, ctypes.c_int], vp),
                              (gdi32.SelectObject, [vp, vp], vp), (gdi32.DeleteObject, [vp], ctypes.c_int),
                              (gdi32.DeleteDC, [vp], ctypes.c_int),
                              (gdi32.GetDIBits, [vp, vp, ctypes.c_uint, ctypes.c_uint, vp, vp, ctypes.c_uint],
                               ctypes.c_int)):
            fn.argtypes, fn.restype = args, res
        left, top, right, bottom = win32gui.GetWindowRect(window)
        width, height = right - left, bottom - top
        window_dc = user32.GetWindowDC(window)
        memory_dc = gdi32.CreateCompatibleDC(window_dc)
        bitmap = gdi32.CreateCompatibleBitmap(window_dc, width, height)
        previous = gdi32.SelectObject(memory_dc, bitmap)
        try:
            if not user32.PrintWindow(window, memory_dc, 2):      # PW_RENDERFULLCONTENT
                user32.PrintWindow(window, memory_dc, 0)
            header = (ctypes.c_int32 * 10)(40, width, -height, 0x00200001, 0, 0, 0, 0, 0, 0)  # 32bpp, planes 1
            buf = ctypes.create_string_buffer(width * height * 4)
            gdi32.GetDIBits(memory_dc, bitmap, 0, height, buf, header, 0)
            raw = buf.raw
        finally:
            gdi32.SelectObject(memory_dc, previous)
            gdi32.DeleteObject(bitmap)
            gdi32.DeleteDC(memory_dc)
            user32.ReleaseDC(window, window_dc)
        return width, height, raw, left, top

    def control_colors(self, hwnd):
        """Colours of a control, read from PM's own rendering."""
        width, height, raw, left, top = self._capture(self.main)
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        return region_colors((width, height, raw), l - left, t - top, r - left, b - top)

    def root(self, hwnd):
        """Top-level window (dialog or PM main window) that holds a control."""
        get_ancestor = ctypes.windll.user32.GetAncestor
        get_ancestor.argtypes, get_ancestor.restype = [ctypes.c_void_p, ctypes.c_uint], ctypes.c_void_p
        return get_ancestor(hwnd, 2) or hwnd       # GA_ROOT

    def snapshot(self, window, mark=None):
        """Picture of one PM window: (width, height, BGRA, box of `mark` inside it or None)."""
        width, height, raw, left, top = self._capture(window)
        box = None
        if mark and mark != window and win32gui.IsWindow(mark):
            l, t, r, b = win32gui.GetWindowRect(mark)
            box = (l - left, t - top, r - left, b - top)
        return width, height, raw, box

    def combo_items(self, hwnd):
        return list(HwndWrapper(hwnd).item_texts())

    def combo_select(self, hwnd, text):
        combo = HwndWrapper(hwnd)
        combo.select(combo.item_texts().index(text))   # exact match, never fuzzy

    def combo_selected(self, hwnd):
        combo = HwndWrapper(hwnd)
        index = combo.selected_index()
        items = combo.item_texts()
        return items[index] if 0 <= index < len(items) else ""

    def describe(self, hwnd):
        texts = []

        def callback(child, _):
            try:
                cls = win32gui.GetClassName(child)
                if cls in ("Static", "Button") and win32gui.IsWindowVisible(child):
                    text = self.text(child).strip()
                    if text:
                        texts.append(text)
            except win32gui.error:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, callback, None)
        except win32gui.error:
            pass
        cls = win32gui.GetClassName(hwnd) if self.exists(hwnd) else "?"
        return f"'{self.window_title(hwnd)}' class={cls} hwnd={hwnd} texts={texts[:12]}"

    # -- UIA-only parts of PM

    def _uia(self, hwnd):
        return Desktop(backend="uia").window(handle=hwnd).wrapper_object()

    def _uia_invoke(self, root_hwnd, name, control_type):
        """Invoke a UIA element in a worker thread (no mouse).

        PM may run the command synchronously and keep the call blocked while the
        dialog it opened is shown, so the worker is left waiting and the runner
        continues; the result is judged by the pre/post checks.
        """
        outcome = {}

        def work():
            try:
                import pythoncom
                pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            except Exception:
                pass
            try:
                target = None
                for element in self._uia(root_hwnd).descendants(control_type=control_type):
                    if element.window_text().strip() == name:
                        target = element
                        break
                if target is None:
                    outcome["missing"] = True
                    return
                outcome["found"] = True
                try:
                    target.invoke()
                except Exception:
                    uia_defines.get_elem_interface(
                        target.element_info.element, "LegacyIAccessible").DoDefaultAction()
                outcome["returned"] = True
            except Exception as exc:
                outcome["error"] = exc

        worker = threading.Thread(target=work, daemon=True)
        worker.start()
        worker.join(T_UIA_INVOKE)
        if "error" in outcome:
            raise outcome["error"]
        if outcome.get("missing"):
            return False
        if not outcome.get("found"):
            raise StepError("uia", f"UIA lookup of '{name}' did not finish within {T_UIA_INVOKE}s.")
        return True

    def click_ribbon_find(self):
        if background_mode():
            if getattr(self, "find_command", None):
                # Same WM_COMMAND the Ctrl+F shortcut sends. Unlike a UIA Invoke it does not
                # leave a call pending inside PM while Find a Player / System Messages are open.
                self.post_command(self.find_command, accelerator=True)
                return True
            return self._uia_invoke(self.main, FIND_RIBBON_BUTTON, "Button")
        self.focus(self.main)
        for button in self._uia(self.main).descendants(control_type="Button"):
            if button.window_text().strip() == FIND_RIBBON_BUTTON:
                button.click_input()
                return True
        return False

    def send_find_shortcut(self):
        if background_mode():
            self._no_input(f"Shortcut {FIND_SHORTCUT}")
        self.focus(self.main)
        send_keys(FIND_SHORTCUT)

    def click_named_item(self, popup_hwnd, name, control_type):
        if background_mode():
            return self._uia_invoke(popup_hwnd, name, control_type)
        for item in self._uia(popup_hwnd).descendants(control_type=control_type):
            if item.window_text().strip() == name:
                item.click_input()
                return True
        return False

    def click_menu_item(self, popup_hwnd, name):
        if self.click_named_item(popup_hwnd, name, "MenuItem"):
            return True
        if background_mode():
            return False
        # Fallback: recorded position of "Redeem Coupon..." in the 244x380 Options menu.
        left, top, right, bottom = win32gui.GetWindowRect(popup_hwnd)
        if name == REDEEM_MENU_ITEM and (right - left, bottom - top) == (244, 380):
            mouse.click(button="left", coords=(left + 122, top + 233))
            return True
        return False

    def list_items(self, popup_hwnd):
        items = self._uia_read(lambda: [t for t in (i.window_text().strip() for i in
                                        self._uia(popup_hwnd).descendants(control_type="ListItem")) if t],
                               "list items")
        return items or []

    def mdi_active(self):
        client = win32gui.FindWindowEx(self.main, 0, "MDIClient", None)
        if not client:
            return 0
        try:
            _, active = win32gui.SendMessageTimeout(
                client, WM_MDIGETACTIVE, 0, 0, win32con.SMTO_ABORTIFHUNG, 2000)
            return int(active or 0)
        except win32gui.error:
            return 0

    def close_tab(self, hwnd):
        win32gui.PostMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_CLOSE, 0)

    def _children(self, hwnd):
        found = []
        try:
            win32gui.EnumChildWindows(hwnd, lambda h, _: found.append(h) or True, None)
        except win32gui.error:
            pass
        return found

    def fill_login(self, hwnd, username, password):
        if background_mode():
            return self._fill_login_background(hwnd, username, password)
        self.focus(hwnd)
        window = self._uia(hwnd)
        edits = [e for e in window.descendants(control_type="Edit") if e.is_visible() and e.is_enabled()]

        def is_password(edit):
            try:
                return bool(edit.element_info.element.CurrentIsPassword)
            except Exception:
                return False

        passwords = [e for e in edits if is_password(e)]
        others = [e for e in edits if not is_password(e)]
        if passwords and others:
            user_ctl, password_ctl = others[0], passwords[0]
        elif len(edits) >= 2:
            user_ctl, password_ctl = edits[0], edits[1]
        else:
            raise StepError("0_login", "Could not identify separate Username and Password fields.")
        for control, value in ((user_ctl, username), (password_ctl, password)):
            control.click_input()
            send_keys("{HOME}+{END}{DEL}")
            send_keys(escape_keys(value), with_spaces=True)
        typed_user = ""
        try:
            typed_user = user_ctl.get_value() or ""
        except Exception:
            typed_user = username
        if typed_user.strip() != username.strip():
            raise StepError("0_login", "Username field does not show the configured username.")
        buttons = [b for b in window.descendants(control_type="Button")
                   if b.window_text().strip().casefold() in LOGIN_BUTTON_NAMES and b.is_enabled()]
        if buttons:
            buttons[0].click_input()
        else:
            send_keys("{ENTER}")

    def _fill_login_background(self, hwnd, username, password):
        edits = [h for h in self._children(hwnd)
                 if "edit" in win32gui.GetClassName(h).casefold() and self.enabled(h)]
        passwords = [h for h in edits if self._style(h) & win32con.ES_PASSWORD]
        others = sorted((h for h in edits if h not in passwords),
                        key=lambda h: win32gui.GetWindowRect(h)[1])
        if not passwords or not others:
            raise StepError("0_login", "Could not identify separate Username and Password fields.")
        user_ctl, password_ctl = others[0], passwords[0]
        HwndWrapper(user_ctl).set_edit_text(username)
        HwndWrapper(password_ctl).set_edit_text(password)
        if self.text(user_ctl).strip() != username.strip():
            raise StepError("0_login", "Username field does not show the configured username.")
        for name in ("Login", "Log In", "Logon", "Log On", "Sign In", "OK"):
            if self._uia_invoke(hwnd, name, "Button"):
                return
        raise StepError("0_login", "Login button not found.")


# ------------------------------------------------------ gates (pre/post)

def read_state(pm):
    ensure_shown = getattr(pm, "ensure_shown", None)
    if ensure_shown:
        ensure_shown()
    return PMState(pm.main_title(), pm.main_enabled(), pm.popups())


def state_problem(state, popups, main_enabled, player_id):
    expected = sorted(popups)
    if state.kinds() != expected:
        return f"expected popups {expected or 'none'}, found {state.kinds() or 'none'}"
    if main_enabled is not None and state.main_enabled != main_enabled:
        return f"expected main window {'enabled' if main_enabled else 'blocked'}"
    if player_id is not None and not title_has_player(state.title, player_id):
        return f"main title does not show player {player_id}: '{state.title}'"
    return None


def raise_on_unknown(pm, step, state, while_doing):
    unknown = [p for p in state.popups if p.kind == UNKNOWN]
    if unknown:
        raise StepError(step, f"unexpected popup while {while_doing}: {pm.describe(unknown[0].hwnd)}")


def check_state(pm, step, popups=(), main_enabled=None, player_id=None):
    """Pre-check: PM must already be in this state, otherwise stop."""
    state = read_state(pm)
    raise_on_unknown(pm, step, state, "checking the state before this step")
    problem = state_problem(state, popups, main_enabled, player_id)
    if problem:
        raise StepError(step, f"pre-check failed: {problem}. State: {state.summary()}")
    log(f"  [{step}] pre-check OK   {state.summary()}")
    return state


def wait_state(pm, step, popups=(), main_enabled=None, player_id=None,
               timeout=T_VERIFY, quiet=0.0, what="the expected state", fail=True):
    """Post-check: wait until PM reaches this state (stable for `quiet` seconds)."""
    deadline = time.time() + timeout
    stable_since = None
    problem = "not checked yet"
    while True:
        state = read_state(pm)
        raise_on_unknown(pm, step, state, f"waiting for {what}")
        problem = state_problem(state, popups, main_enabled, player_id)
        if problem is None:
            stable_since = stable_since or time.time()
            if time.time() - stable_since >= quiet:
                log(f"  [{step}] post-check OK  {state.summary()}")
                return state
        else:
            stable_since = None
        if time.time() > deadline:
            if not fail:
                return None
            raise StepError(step, f"timed out after {timeout}s waiting for {what}: {problem}. "
                                  f"State: {state.summary()}")
        time.sleep(POLL)


def wait_until(predicate, timeout, interval=POLL):
    deadline = time.time() + timeout
    while True:
        try:
            if predicate():
                return True
        except Exception:
            pass
        if time.time() > deadline:
            return False
        time.sleep(interval)


def state_if(pm, popups=(), main_enabled=None, player_id=None):
    """The current state when it matches, else None."""
    state = read_state(pm)
    return state if state_problem(state, popups, main_enabled, player_id) is None else None


def screen_is(popups=(), main_enabled=None):
    """Predicate on a state: exactly these popups (and main window enabled/blocked)."""
    def check(state):
        return state.kinds() == sorted(popups) and (main_enabled is None or state.main_enabled == main_enabled)
    return check


_screens = {"dir": None, "count": 0}


def save_screens(pm, step, reason, target=None, everything=False):
    """Save pictures of PM windows under SCREENSHOT_DIR/<run time>/ and return the file names.

    target: the control the step was working on; its window is saved whole with the control
    framed in red, plus a close-up of the area around it. everything: also every open PM popup
    and the main window. Pictures come from PM's own rendering (PrintWindow): no focus, no
    mouse, never the rest of the desktop, never the login window. Never stops the run.
    """
    if not SCREENSHOTS or not hasattr(pm, "snapshot"):
        return []
    saved = []
    try:
        if _screens["dir"] is None:
            _screens["dir"] = SCREENSHOT_DIR / datetime.now().strftime("%Y%m%d_%H%M%S")
        folder = _screens["dir"]
        folder.mkdir(parents=True, exist_ok=True)
        state = read_state(pm)
        login = {p.hwnd for p in state.popups if p.kind == LOGIN}
        windows = []
        if target and pm.exists(target):
            windows.append((pm.root(target), target))
        if everything or not windows:
            windows += [(p.hwnd, None) for p in state.popups if p.kind != LOGIN]
            windows.append((pm.main, None))
        seen = set()
        for window, mark in windows:
            if not window or window in seen or window in login or (window != pm.main and not pm.visible(window)):
                continue
            seen.add(window)
            width, height, raw, box = pm.snapshot(window, mark)
            if width <= 0 or height <= 0:
                continue
            rgb = bgra_to_rgb(width, height, raw)
            _screens["count"] += 1
            name = re.sub(r"[^A-Za-z0-9_.-]+", "_", f"{_screens['count']:03d}_{step}_{reason}")[:90]
            kind = next((p.kind for p in state.popups if p.hwnd == window), "PM")
            base = folder / f"{name}_{kind}"
            if box:
                area = crop_rgb(rgb, width, height, (box[0] - SCREENSHOT_MARGIN, box[1] - SCREENSHOT_MARGIN,
                                                     box[2] + SCREENSHOT_MARGIN, box[3] + SCREENSHOT_MARGIN))
                draw_box(rgb, width, height, box)
                if area:
                    area_w, area_h, area_rgb = area
                    draw_box(area_rgb, area_w, area_h, (box[0] - max(0, box[0] - SCREENSHOT_MARGIN),
                                                        box[1] - max(0, box[1] - SCREENSHOT_MARGIN),
                                                        box[2] - max(0, box[0] - SCREENSHOT_MARGIN),
                                                        box[3] - max(0, box[1] - SCREENSHOT_MARGIN)))
                    write_png(f"{base}_area.png", area_w, area_h, area_rgb)
                    saved.append(f"{base.name}_area.png")
            write_png(f"{base}.png", width, height, rgb)
            saved.append(f"{base.name}.png")
        if saved:
            log(f"  [{step}] screenshot(s) in {folder.name}: {', '.join(saved)}")
    except Exception as exc:
        log(f"  [{step}] screenshot not saved: {exc}")
    return saved


def attempt_until(pm, step, what, action, done, timeout, same_screen=None, describe=None, target=None):
    """Do `action`, then wait up to `timeout` for `done()`; repeat it while the result is missing.

    At most MAX_ATTEMPTS tries in total. Before a new try PM must still show the screen the
    action was made for (`same_screen(state)`); any other screen or an unknown popup stops the run.
    An error reported by Windows while sending the action does not stop the run by itself: the
    result is read back from the control, and only a missing result counts as a failed try.
    Each failed try saves a screenshot of `target` (the control) and its window.
    """
    errors = {}
    for attempt in range(1, MAX_ATTEMPTS + 1):
        if attempt > 1:
            time.sleep(RETRY_PAUSE_SECONDS)
            state = read_state(pm)
            raise_on_unknown(pm, step, state, what)
            result = _safe_done(done, errors)
            if result:
                log(f"  [{step}] {what}: done (late, no new try needed)")
                return result
            if same_screen is not None and not same_screen(state):
                raise StepError(step, f"{what}: not done and PM shows another screen, so it is not repeated"
                                      f"{_details(describe)}{_errors(errors)}. State: {state.summary()}", target)
            save_screens(pm, step, f"{what}_try{attempt - 1}_not_done", target)
            log(f"  [{step}] {what}: not done yet, trying again (try {attempt}/{MAX_ATTEMPTS})")
        try:
            action(attempt)
        except StepError:
            raise
        except Exception as exc:
            errors["send"] = f"{type(exc).__name__}: {exc}"
            log(f"  [{step}] {what}: Windows reported an error ({errors['send']}); "
                f"checking the control to see whether it took effect")
        deadline = time.time() + timeout
        while True:
            state = read_state(pm)
            raise_on_unknown(pm, step, state, what)
            result = _safe_done(done, errors)
            if result:
                if attempt > 1 or "send" in errors:
                    log(f"  [{step}] {what}: done on try {attempt} (checked on the control)")
                return result
            if time.time() > deadline:
                break
            time.sleep(POLL)
    raise StepError(step, f"{what}: still not done after {MAX_ATTEMPTS} tries"
                          f"{_details(describe)}{_errors(errors)}. State: {read_state(pm).summary()}", target)


def _errors(errors):
    parts = []
    if "send" in errors:
        parts.append(f"last Windows error: {errors['send']}")
    if "read" in errors:
        parts.append(f"last read error: {errors['read']}")
    return f" [{'; '.join(parts)}]" if parts else ""


def _details(describe):
    if describe is None:
        return ""
    try:
        return f" ({describe()})"
    except Exception as exc:
        return f" (details not readable: {exc})"


def _safe_done(done, errors=None):
    """done() with read errors (e.g. a control being destroyed) counted as 'not yet'."""
    try:
        return done()
    except StepError:
        raise
    except Exception as exc:
        if errors is not None:
            errors["read"] = f"{type(exc).__name__}: {exc}"
        return None


def close_popup(pm, step, popup, control_id, what, expect, timeout=None):
    """Click a closing button (Close / Cancel) and retry until the popup is gone."""
    attempt_until(pm, step, what,
                  lambda attempt: click(pm, step, popup, control_id, what, expect=expect),
                  lambda: not pm.visible(popup.hwnd), timeout or T_POPUP_CLOSE,
                  same_screen=lambda state: any(p.hwnd == popup.hwnd for p in state.popups),
                  target=pm.child(popup.hwnd, control_id) or popup.hwnd)
    log(f"  [{step}] {popup.kind} closed")


def wait_popup_closed(pm, step, popup, timeout=T_POPUP_CLOSE):
    deadline = time.time() + timeout
    while pm.visible(popup.hwnd):
        state = read_state(pm)
        raise_on_unknown(pm, step, state, f"waiting for {popup.kind} to close")
        if time.time() > deadline:
            raise StepError(step, f"{popup.label()} did not close within {timeout}s. State: {state.summary()}")
        time.sleep(POLL)
    log(f"  [{step}] {popup.kind} closed")


def caption(pm, hwnd):
    """Button text without the & of the keyboard mnemonic ("&Close" -> "Close")."""
    return pm.text(hwnd).replace("&", "").strip()


def control(pm, step, popup, control_id, what, need_enabled=True, timeout=T_VERIFY, expect=None):
    """Find a control by its ID in a popup; with `expect`, its caption must match too."""
    found = {}

    def ready():
        hwnd = pm.child(popup.hwnd, control_id)
        found["hwnd"] = hwnd
        return hwnd and (pm.enabled(hwnd) if need_enabled else pm.exists(hwnd))

    if not wait_until(ready, timeout):
        state = "missing" if not found.get("hwnd") else "disabled"
        raise StepError(step, f"{what} (id={control_id}) in {popup.label()} is {state}.")
    if expect is not None and caption(pm, found["hwnd"]).casefold() != expect.casefold():
        raise StepError(step, f"control id={control_id} in {popup.label()} reads '{caption(pm, found['hwnd'])}', "
                              f"expected '{expect}'. Nothing was clicked.")
    return found["hwnd"]


def click(pm, step, popup, control_id, what, allow_ok=False, expect=None):
    allowed = {COUPON: COUPON_CLICKABLE_IDS, ADJUST: ADJUST_CLICKABLE_IDS, COUPON_CONFIRM: set()}.get(popup.kind)
    if allowed is not None and control_id not in allowed | ({ID_OK} if allow_ok else set()):
        raise StepError(step, f"Safety stop: refusing to click control id={control_id} in {popup.title}.")
    hwnd = control(pm, step, popup, control_id, what, expect=expect)
    if control_id == ID_OK and allow_ok and hasattr(pm, "focus_in_dialog"):
        try:
            pm.focus_in_dialog(popup.hwnd, hwnd)
            log(f"  [{step}] {what}: the OK button gets the dialog's focus first (as with a real click)")
        except Exception as exc:
            log(f"  [{step}] {what}: could not give the OK button the focus ({exc}); clicking anyway")
    log(f"  [{step}] click {what} (id={control_id}" + (f", '{expect}'" if expect else "") + ")")
    pm.click(hwnd)
    return hwnd


def set_and_verify(pm, step, hwnd, value, what, numeric=False, screen=None):
    """Write a field and read it back; rewrite it (at most MAX_ATTEMPTS tries) until it shows the value."""
    norm = normalized_number_text if numeric else (lambda v: str(v).strip())
    expected = norm(value)

    def action(attempt):
        if attempt == MAX_ATTEMPTS and not background_mode():
            log(f"  [{step}] {what}: typing instead")
            pm.type_text(hwnd, value)
        else:
            pm.set_text(hwnd, value)

    attempt_until(pm, step, f"{what} = {expected}", action, lambda: norm(pm.text(hwnd)) == expected, T_VERIFY,
                  same_screen=screen_is(screen) if screen is not None else None,
                  describe=lambda: f"field shows '{pm.text(hwnd)}'", target=hwnd)
    log(f"  [{step}] checkpoint: {what} = {expected}")


def ensure_logged_in(pm, username, password):
    step = "0_login"
    state = read_state(pm)
    if state.get(LOGIN) is None and state.title.startswith("Patron Management - "):
        log("PM is already logged in.")
        return
    if state.get(LOGIN) is None:
        log("Waiting for the PM login window...")
        state = wait_state(pm, step, popups=(LOGIN,), timeout=30, what="the login window")
    if (not username or not password or username == "ENTER_USERNAME_HERE"
            or password == "ENTER_PASSWORD_HERE"):
        raise StepError(step, "Fill in pm_credentials.py first.")
    login_window = check_state(pm, step, popups=(LOGIN,)).get(LOGIN)
    log(f"  [{step}] typing credentials into {login_window.label()}")
    pm.fill_login(login_window.hwnd, username, password)
    deadline = time.time() + T_LOGIN
    while True:
        state = read_state(pm)
        raise_on_unknown(pm, step, state, "logging in")
        if (state.get(LOGIN) is None and not state.popups and state.main_enabled
                and state.title.startswith("Patron Management - ")):
            break
        if time.time() > deadline:
            raise StepError(step, f"login did not finish within {T_LOGIN}s. State: {state.summary()}")
        time.sleep(POLL)
    wait_state(pm, step, popups=(), main_enabled=True, quiet=QUIET_SECONDS, timeout=10,
               what="PM idle after login")
    log("Login completed.")


def open_find_player(pm, step):
    check_state(pm, step, popups=(), main_enabled=True)

    def action(attempt):
        command = getattr(pm, "find_command", None) if background_mode() else None
        if attempt == MAX_ATTEMPTS and command:
            log(f"  [{step}] the Ctrl+F command {command} did not open Find a Player; using the Ribbon button")
            pm.find_command = command = None
        if attempt == MAX_ATTEMPTS and not background_mode():
            log(f"  [{step}] shortcut {FIND_SHORTCUT}")
            pm.send_find_shortcut()
            return
        log(f"  [{step}] " + (f"send the Ctrl+F command (id {command})" if command
                              else f"click Ribbon '{FIND_RIBBON_BUTTON}'"))
        try:
            pm.click_ribbon_find()
        except Exception as exc:
            log(f"  [{step}] opening Find a Player failed: {exc}")

    state = attempt_until(pm, step, "open Find a Player", action,
                          lambda: state_if(pm, popups=(FIND,), main_enabled=False), T_FIND_OPEN,
                          same_screen=screen_is((), main_enabled=True), target=pm.main)
    log(f"  [{step}] post-check OK  {state.summary()}")
    return state


def wait_profile_title(pm, step, player_id, previous_title, previous_tab):
    """Wait for the new profile; the old title stays for ~4-5s and is not an error.

    When this player was already on screen the title cannot show the reload, so
    wait for a new profile tab, or SAME_PLAYER_WAIT_SECONDS if PM reuses the tab.
    """
    started = time.time()
    deadline = started + T_PROFILE_LOAD
    same_player = title_has_player(previous_title, player_id)
    while True:
        state = read_state(pm)
        raise_on_unknown(pm, step, state, "waiting for the profile to load")
        others = [p for p in state.popups if p.kind not in (SYSMSG, COMMENT)]
        if others:
            raise StepError(step, f"unexpected popup while loading the profile: {others[0].label()}")
        if title_has_player(state.title, player_id):
            tab = pm.mdi_active()
            if not same_player or (tab and tab != previous_tab):
                log(f"  [{step}] post-check OK  profile loaded: '{state.title}'")
                return state
            if time.time() - started >= SAME_PLAYER_WAIT_SECONDS:
                log(f"  [{step}] post-check OK  same player was already open; "
                    f"waited {SAME_PLAYER_WAIT_SECONDS:.0f}s for the reload")
                return state
        if state.title != previous_title and ids_in_title(state.title):
            raise StepError(step, f"Wrong profile opened for {player_id}: '{state.title}'")
        if time.time() > deadline:
            raise StepError(step, f"profile {player_id} did not load within {T_PROFILE_LOAD}s. "
                                  f"State: {state.summary()}")
        time.sleep(POLL)


def summarize_messages(items):
    """System Messages lines as 'Title (description)'; the smallest indent is a group header."""
    if not items:
        return [], []
    levels = sorted({left for _, left in items})
    title_level = levels[1] if len(levels) > 2 else levels[0]
    messages, current = [], None
    for text, left in items:
        if left < title_level:
            continue                      # group header such as "Player Stop Codes"
        if left == title_level:
            current = [text, ""]
            messages.append(current)
        elif current:
            current[1] = (current[1] + " " + text).strip()
    return [f"{t} ({d})" if d else t for t, d in messages], [t for t, _ in messages]


def handle_system_messages(pm, step, popup):
    items = pm.list_items_levels(popup.hwnd)
    lines, titles = summarize_messages(items)
    log(f"  [{step}] System Messages: {' | '.join(lines) or '(no readable items)'} (logged only)")
    close_popup(pm, step, popup, SYSMSG_CLOSE, "System Messages Close", expect="Close")
    return lines, titles


def comment_signature(pm, popup, next_hwnd, close_hwnd):
    header = pm.child(popup.hwnd, COMMENT_HEADER)
    previous = pm.child(popup.hwnd, COMMENT_PREVIOUS)
    return (pm.text(header) if header else "", pm.enabled(next_hwnd),
            pm.enabled(previous) if previous else None, pm.enabled(close_hwnd))


def handle_player_comment(pm, step, popup):
    """Close is disabled until the last comment has been shown: page with Next first."""
    next_hwnd = control(pm, step, popup, COMMENT_NEXT, "Player Comment Next", need_enabled=False, expect="Next")
    close_hwnd = control(pm, step, popup, COMMENT_CLOSE, "Player Comment Close", need_enabled=False,
                         expect="Close")
    page = 1
    while True:
        if pm.enabled(close_hwnd):
            log(f"  [{step}] Player Comment: all {page} page(s) shown, Close is enabled")
            close_popup(pm, step, popup, COMMENT_CLOSE, "Player Comment Close", expect="Close")
            return page
        if page >= MAX_COMMENT_PAGES:
            raise StepError(step, f"Player Comment still not closable after {page} pages.")
        if pm.enabled(next_hwnd):
            before = comment_signature(pm, popup, next_hwnd, close_hwnd)

            def press_next(attempt, page=page):
                log(f"  [{step}] Player Comment: click Next (page {page} -> {page + 1})")
                pm.click(next_hwnd)

            attempt_until(pm, step, "Player Comment Next", press_next,
                          lambda: comment_signature(pm, popup, next_hwnd, close_hwnd) != before, T_COMMENT_PAGE,
                          same_screen=lambda state: any(p.hwnd == popup.hwnd for p in state.popups),
                          target=next_hwnd)
            page += 1
            continue
        # Next can be hidden for ~0.1s while PM redraws it.
        if not wait_until(lambda: pm.enabled(close_hwnd) or pm.enabled(next_hwnd), 2.0):
            raise StepError(step, "Player Comment: Next and Close are both disabled.")


def settle_profile(pm, step, player_id, quiet):
    """Handle System Messages / Player Comment until PM stays idle for `quiet` seconds."""
    deadline = time.time() + T_PROFILE_POPUPS
    quiet_since = None
    messages, codes, pages = [], [], 0
    while True:
        state = read_state(pm)
        raise_on_unknown(pm, step, state, "handling profile popups")
        if not title_has_player(state.title, player_id):
            raise StepError(step, f"profile title changed unexpectedly: '{state.title}'")
        others = [p for p in state.popups if p.kind not in (SYSMSG, COMMENT)]
        if others:
            raise StepError(step, f"unexpected popup on the profile: {others[0].label()}")
        if state.get(SYSMSG):
            lines, titles = handle_system_messages(pm, step, state.get(SYSMSG))
            messages.extend(lines)
            codes.extend(titles)
            quiet_since = None
        elif state.get(COMMENT):
            pages += handle_player_comment(pm, step, state.get(COMMENT))
            quiet_since = None
        elif not state.main_enabled:
            quiet_since = None      # a modal popup is being created (shows 0.03-0.25s later)
        else:
            quiet_since = quiet_since or time.time()
            if time.time() - quiet_since >= quiet:
                log(f"  [{step}] post-check OK  profile idle for {quiet:.1f}s  {state.summary()}")
                return messages, codes, pages
        if time.time() > deadline:
            raise StepError(step, f"profile did not become idle within {T_PROFILE_POPUPS}s. "
                                  f"State: {state.summary()}")
        time.sleep(POLL)


def read_identification(pm, step, tab):
    """Text and colour of the player name in the Identification frame (by control ID)."""
    frame = pm.child(tab, IDENT_FRAME_ID)
    name = pm.child(tab, IDENT_NAME_ID)
    if not frame or not name or not pm.visible(name):
        raise StepError(step, f"Identification name (id {IDENT_NAME_ID}) not found on the profile tab.")
    if pm.text(frame).strip() != "Identification":
        raise StepError(step, f"control {IDENT_FRAME_ID} is '{pm.text(frame)}', expected frame 'Identification'.")
    fl, ft, fr, fb = pm.rect(frame)
    nl, nt, nr, nb = pm.rect(name)
    if not (fl <= nl and nr <= fr and ft <= nt and nb <= fb):
        raise StepError(step, f"control {IDENT_NAME_ID} is not inside the Identification frame.")
    text = pm.text(name).strip()
    if not text:
        raise StepError(step, "Identification name is empty.")
    try:
        colors = pm.control_colors(name) or {}
    except Exception as exc:
        log(f"  [{step}] colour not readable: {exc}")
        colors = {}
    return text, colors


def active_profile_tab(pm, step, player_id):
    tab = pm.mdi_active()
    title = pm.window_title(tab) if tab else ""
    if not tab or not title_has_player(title, player_id):
        raise StepError(step, f"active profile tab is not player {player_id}: '{title}'")
    return tab


def select_combo(pm, step, combo_hwnd, value, what, screen):
    """Select an exact item; reselect (at most MAX_ATTEMPTS tries) until the list shows it."""
    items = pm.combo_items(combo_hwnd)
    if value not in items:
        raise StepError(step, f"'{value}' is not in the {what} list ({len(items)} items).")

    def action(attempt):
        if attempt == MAX_ATTEMPTS and not background_mode():
            log(f"  [{step}] open the {what} list and click '{value}'")
            pm.click(combo_hwnd)
            state = wait_state(pm, step, popups=tuple(screen) + (DROPDOWN,), timeout=T_VERIFY,
                               what=f"the {what} list")
            if not pm.click_named_item(state.get(DROPDOWN).hwnd, value, "ListItem"):
                raise StepError(step, f"'{value}' not found in the open {what} list.")
            return
        log(f"  [{step}] select '{value}' in {what}")
        try:
            pm.combo_select(combo_hwnd, value)
        except Exception as exc:
            log(f"  [{step}] direct select failed: {exc}")

    attempt_until(pm, step, f"{what} = {value}", action, lambda: pm.combo_selected(combo_hwnd) == value,
                  T_VERIFY, same_screen=lambda state: state.kinds() in (sorted(screen), sorted(tuple(screen) + (DROPDOWN,))),
                  describe=lambda: f"list shows '{pm.combo_selected(combo_hwnd)}'", target=combo_hwnd)


def money(text):
    cleaned = re.sub(r"[^0-9.\-]", "", str(text))
    return float(cleaned) if cleaned not in ("", "-", ".") else None


def inside(outer, inner, tolerance=4):
    """inner rectangle within outer (group box borders may be a few pixels off)."""
    return (outer[0] - tolerance <= inner[0] and inner[2] <= outer[2] + tolerance
            and outer[1] - tolerance <= inner[1] and inner[3] <= outer[3] + tolerance)


def begin_step(progress, name):
    """Pause STEP_DELAY_SECONDS before every step, then record the step name."""
    time.sleep(STEP_DELAY_SECONDS)
    progress["step"] = name
    return name


def open_profile(pm, job, progress):
    """Steps 1-6b, common to every workflow: find the player and make sure it may be processed."""
    player_id = job["player_id"]

    step = begin_step(progress, "1_precheck_idle")
    state = check_state(pm, step, popups=(), main_enabled=True)
    if not state.title.startswith("Patron Management - "):
        raise StepError(step, f"PM is not on a logged-in page: '{state.title}'")
    open_tab = pm.mdi_active()
    if open_tab and title_has_player(pm.window_title(open_tab), player_id):
        # Same player as the previous job: close its tab so Find opens a fresh profile.
        log(f"  [{step}] profile of {player_id} is still open, closing that tab first")
        close_tab_until_gone(pm, step, open_tab, f"close the tab of {player_id}")
        state = wait_state(pm, step, popups=(), main_enabled=True, timeout=T_TAB_CLOSE, quiet=0.5,
                           what="PM idle after closing the tab")
    previous_title = state.title
    previous_tab = pm.mdi_active()

    step = begin_step(progress, "2_open_find_player")
    open_find_player(pm, step)

    step = begin_step(progress, "3_enter_player_id")
    find = check_state(pm, step, popups=(FIND,), main_enabled=False).get(FIND)
    field = control(pm, step, find, FIND_PLAYER_ID_EDIT, "Player ID field")
    set_and_verify(pm, step, field, player_id, "Player ID", screen=(FIND,))

    step = begin_step(progress, "4_confirm_find")
    find = check_state(pm, step, popups=(FIND,), main_enabled=False).get(FIND)
    if pm.text(field).strip() != player_id:
        raise StepError(step, f"Player ID field changed to '{pm.text(field)}' before OK.")
    attempt_until(pm, step, "Find a Player OK",
                  lambda attempt: click(pm, step, find, ID_OK, "Find a Player OK", expect="OK"),
                  lambda: not pm.visible(find.hwnd), T_FIND_CLOSE,
                  same_screen=lambda state: state.get(FIND) is not None and pm.text(field).strip() == player_id,
                  target=pm.child(find.hwnd, ID_OK) or find.hwnd)
    log(f"  [{step}] FIND closed")

    step = begin_step(progress, "5_wait_profile_loaded")
    wait_profile_title(pm, step, player_id, previous_title, previous_tab)

    step = begin_step(progress, "6_profile_popups")
    messages, codes, pages = settle_profile(pm, step, player_id, QUIET_SECONDS)
    progress["tab"] = active_profile_tab(pm, step, player_id)
    progress["system_messages"] = " | ".join(messages)
    progress["stop_codes"] = ", ".join(codes)
    progress["comment_pages"] = pages

    step = begin_step(progress, "6b_check_identification")
    check_state(pm, step, popups=(), main_enabled=True, player_id=player_id)
    name, colors = read_identification(pm, step, active_profile_tab(pm, step, player_id))
    green = colors.get("background_color") == "green" if colors else None
    color_text = (f"background {colors.get('background_color')} {colors.get('background_rgb')} "
                  f"(green background: {'YES' if green else 'no'}), "
                  f"text {colors.get('text_color')} {colors.get('text_rgb')}") if colors else "colour unknown"
    progress["identification"] = name
    progress["name_color"] = color_text
    log(f"  [{step}] Identification name: '{name}' ({color_text})")
    if SKIP_NAME_RE.search(name):
        raise PlayerSkipped("SKIPPED_LOC", f"Identification name '{name}' contains '(Loc:' ({color_text}); "
                                           "player not processed.")
    log(f"  [{step}] checkpoint: no '(Loc:' in the name, continue")


def finish_dialog(pm, step, job, progress, dialog, allow_ok):
    """OK (real, recorded in the ledger first) or Cancel (test), then PM must be idle again."""
    player_id = job["player_id"]
    if not allow_ok:
        close_popup(pm, step, dialog, ID_CANCEL, f"{dialog.title} Cancel", expect="Cancel")
        wait_state(pm, step, popups=(), main_enabled=True, player_id=player_id,
                   timeout=T_POPUP_CLOSE, quiet=0.5, what="PM idle after Cancel")
        return
    # Written before the click: if anything goes wrong afterwards this job is
    # still treated as done and is never clicked again by a rerun today.
    ledger_append(job, "OK_CLICKED")
    progress["ok_clicked"] = True
    # Clicked exactly once: a slow PM must never get a second OK (no double redemption).
    click(pm, step, dialog, ID_OK, f"{dialog.title} OK", allow_ok=True, expect="OK")
    wait_popup_closed(pm, step, dialog, T_AFTER_OK)
    wait_state(pm, step, popups=(), main_enabled=True, player_id=player_id,
               timeout=T_AFTER_OK, quiet=1.0, what="PM idle after OK")
    ledger_append(job, "DONE")


def confirmation_notes(text, name, amount, identification):
    """Differences between PM's confirmation and the Excel row (logged only: the coupon is confirmed anyway)."""
    notes = []
    match = CONFIRM_TEXT_RE.search(text)
    if not match:
        notes.append(f"confirmation text not recognised: '{text}'")
    else:
        if abs(float(match.group(1).replace(",", "")) - float(amount)) > 0.005:
            notes.append(f"confirmation says ${match.group(1)}, Excel amount is {amount}")
        if match.group(2).upper() != CONFIRM_BUCKET:
            notes.append(f"confirmation says 'in {match.group(2)}', expected 'in {CONFIRM_BUCKET}'")

    def words(value):
        return set(re.findall(r"[A-Za-z0-9]+", value.upper()))

    if not name:
        notes.append("confirmation shows no player name")
    elif identification and not words(name) <= words(identification):
        notes.append(f"confirmation name '{name}' is not the profile name '{identification}'")
    return notes


def finish_coupon_ok(pm, step, job, progress, coupon):
    """OK once, then PM's confirmation ("The coupon will reward the player with $X in SLOTS.") OK once.

    The confirmation is compared with the Excel row and the profile name; a difference is only
    noted (the coupon is confirmed anyway). Neither OK is ever clicked a second time.
    """
    player_id = job["player_id"]
    ledger_append(job, "OK_CLICKED")
    progress["ok_clicked"] = True
    click(pm, step, coupon, ID_OK, "Coupon Redemption OK", allow_ok=True, expect="OK")
    deadline = time.time() + T_AFTER_OK
    while True:
        state = read_state(pm)
        raise_on_unknown(pm, step, state, "waiting for the coupon confirmation")
        confirm = state.get(COUPON_CONFIRM)
        if confirm is not None or (not pm.visible(coupon.hwnd) and state.get(COUPON) is None):
            break
        if time.time() > deadline:
            raise StepError(step, f"no confirmation {T_AFTER_OK}s after OK and Coupon Redemption is still open. "
                                  f"State: {state.summary()}", coupon.hwnd)
        time.sleep(POLL)
    if confirm is None:
        progress["confirmation_check"] = "no confirmation was shown"
        log(f"  [{step}] NOTE: Coupon Redemption closed without the usual confirmation")
    else:
        step = begin_step(progress, "13_confirm_coupon")
        check_state(pm, step, popups=(COUPON, COUPON_CONFIRM), main_enabled=False, player_id=player_id)
        text_hwnd = control(pm, step, confirm, CONFIRM_TEXT_ID, "confirmation text", need_enabled=False)
        wait_until(lambda: pm.text(text_hwnd).strip(), T_VERIFY)
        text = " ".join(pm.text(text_hwnd).split())
        name_hwnd = pm.child(confirm.hwnd, CONFIRM_NAME_ID)
        name = pm.text(name_hwnd).strip() if name_hwnd else ""
        notes = confirmation_notes(text, name, job["amount"], progress.get("identification", ""))
        progress["confirmation"] = f"{name}: {text}" if name else text
        progress["confirmation_check"] = "; ".join(notes) if notes else "OK"
        if notes:
            log(f"  [{step}] NOTE (logged only, confirming anyway): {'; '.join(notes)}")
        else:
            log(f"  [{step}] checkpoint: confirmation matches: {progress['confirmation']}")
        ledger_append(job, "CONFIRM_CLICKED")
        # Clicked exactly once, like the first OK.
        click(pm, step, confirm, ID_OK, "confirmation OK", allow_ok=True, expect="OK")
        wait_popup_closed(pm, step, confirm, T_AFTER_OK)
    wait_state(pm, step, popups=(), main_enabled=True, player_id=player_id,
               timeout=T_AFTER_OK, quiet=1.0, what="PM idle after the redemption")
    ledger_append(job, "DONE")


def run_coupon(pm, job, progress, allow_ok):
    """Workflows #1 and #3: F12 (or Options > Redeem Coupon...) > Competitor Coupon > competitor > amount."""
    player_id, amount = job["player_id"], job["amount"]
    competitor = WORKFLOWS[job["workflow"]]["competitor"]

    step = begin_step(progress, "7_open_coupon_redemption")
    check_state(pm, step, popups=(), main_enabled=True, player_id=player_id)
    tab = active_profile_tab(pm, step, player_id)
    command = getattr(pm, "redeem_command", None) or REDEEM_COUPON_COMMAND_ID
    options = pm.child(tab, OPTIONS_BUTTON)
    menu_usable = bool(options) and pm.enabled(options) and caption(pm, options) == "Options..."
    if not command and not menu_usable:
        raise StepError(step, "Options... button not found, disabled or renamed on the active profile tab.")

    def via_menu():
        menu = read_state(pm).get(MENU)
        if menu is None:
            log(f"  [{step}] click Options... (id={OPTIONS_BUTTON})")
            pm.click(options)
            state = wait_state(pm, step, popups=(MENU,), main_enabled=True, player_id=player_id,
                               timeout=T_MENU_OPEN, what="the Options menu", fail=False)
            if state is None:
                log(f"  [{step}] the Options menu did not open")
                return
            menu = state.get(MENU)
        log(f"  [{step}] click menu item '{REDEEM_MENU_ITEM}'")
        if not pm.click_menu_item(menu.hwnd, REDEEM_MENU_ITEM):
            raise StepError(step, f"Menu item '{REDEEM_MENU_ITEM}' not found.")

    def open_coupon(attempt):
        if command and (attempt < MAX_ATTEMPTS or not menu_usable):
            log(f"  [{step}] send the F12 command (id {command}): Redeem Coupon without the keyboard")
            pm.post_command(command, accelerator=not REDEEM_COUPON_COMMAND_ID)
            return
        if command:
            log(f"  [{step}] the F12 command did not open Coupon Redemption; using Options... > {REDEEM_MENU_ITEM}")
        via_menu()

    attempt_until(pm, step, "open Coupon Redemption", open_coupon,
                  lambda: state_if(pm, popups=(COUPON,), main_enabled=False, player_id=player_id), T_COUPON_OPEN,
                  same_screen=lambda state: state.kinds() in ([], [MENU]) and state.main_enabled,
                  target=options if options and not command else pm.main)
    log(f"  [{step}] post-check OK  Coupon Redemption is open")

    step = begin_step(progress, "9_select_competitor_coupon")
    coupon = check_state(pm, step, popups=(COUPON,), main_enabled=False, player_id=player_id).get(COUPON)
    radio = control(pm, step, coupon, COUPON_COMPETITOR_RADIO, "Competitor Coupon", expect="Competitor Coupon")
    our_radio = control(pm, step, coupon, COUPON_OUR_RADIO, "Our Coupon", need_enabled=False, expect="Our Coupon")
    combo = control(pm, step, coupon, COUPON_COMPETITOR_COMBO, "Competitor list", need_enabled=False)
    amount_field = control(pm, step, coupon, COUPON_AMOUNT_EDIT, "Amount", need_enabled=False)
    log(f"  [{step}] initial: Competitor list enabled={pm.enabled(combo)}, Amount enabled={pm.enabled(amount_field)}")
    attempt_until(pm, step, "select Competitor Coupon",
                  lambda attempt: click(pm, step, coupon, COUPON_COMPETITOR_RADIO, "Competitor Coupon",
                                        expect="Competitor Coupon"),
                  lambda: pm.checked(radio) and not pm.checked(our_radio)
                  and pm.enabled(combo) and pm.enabled(amount_field), T_FIELDS_ENABLE,
                  same_screen=screen_is((COUPON,), main_enabled=False),
                  describe=lambda: f"selected={pm.checked(radio)}, list enabled={pm.enabled(combo)}, "
                                   f"Amount enabled={pm.enabled(amount_field)}", target=radio)
    wait_state(pm, step, popups=(COUPON,), main_enabled=False, player_id=player_id,
               what="Coupon Redemption only")
    log(f"  [{step}] checkpoint: Competitor Coupon selected, list and Amount enabled")

    step = begin_step(progress, "10_select_competitor")
    check_state(pm, step, popups=(COUPON,), main_enabled=False, player_id=player_id)
    select_combo(pm, step, combo, competitor, "Competitor", screen=(COUPON,))
    wait_state(pm, step, popups=(COUPON,), main_enabled=False, player_id=player_id,
               what="Coupon Redemption only")
    log(f"  [{step}] checkpoint: Competitor = {competitor}")

    step = begin_step(progress, "11_enter_amount")
    check_state(pm, step, popups=(COUPON,), main_enabled=False, player_id=player_id)
    set_and_verify(pm, step, amount_field, amount, "Amount", numeric=True, screen=(COUPON,))
    wait_state(pm, step, popups=(COUPON,), main_enabled=False, player_id=player_id,
               quiet=0.5, what="no popup after entering the amount")

    step = begin_step(progress, "12_ok_coupon" if allow_ok else "12_cancel_coupon")
    check_state(pm, step, popups=(COUPON,), main_enabled=False, player_id=player_id)
    if pm.combo_selected(combo) != competitor or \
            normalized_number_text(pm.text(amount_field)) != normalized_number_text(amount) or \
            not pm.checked(radio) or pm.checked(our_radio) or \
            active_profile_tab(pm, step, player_id) != progress["tab"]:
        raise StepError(step, "Coupon fields or profile changed before the final click. Nothing was clicked.")
    log(f"  [{step}] final check OK: player {player_id}, {competitor}, amount {amount}")
    if allow_ok:
        finish_coupon_ok(pm, step, job, progress, coupon)
    else:
        finish_dialog(pm, step, job, progress, coupon, allow_ok)


def run_bbr(pm, job, progress, allow_ok):
    """Workflows #2 and #4: Rewards BBR > Adjust > Add BBR, amount, expiration, reason, comment."""
    player_id, amount = job["player_id"], job["amount"]
    wf = WORKFLOWS[job["workflow"]]
    expires = expiration_for(job["workflow"], progress["today"])
    comment = wf["comment"](progress["today"])

    step = begin_step(progress, "7_select_bbr")
    check_state(pm, step, popups=(), main_enabled=True, player_id=player_id)
    tab = active_profile_tab(pm, step, player_id)
    frame, bbr, adjust = (pm.child(tab, REWARDS_FRAME_ID), pm.child(tab, BBR_RADIO), pm.child(tab, ADJUST_BUTTON))
    if not frame or not bbr or not adjust:
        raise StepError(step, "Rewards frame, BBR or Adjust not found on the profile tab.")
    if pm.text(frame).strip() != "Rewards" or pm.text(bbr).replace("&", "").strip() != "BBR" \
            or pm.text(adjust).replace("&", "").strip() != "Adjust":
        raise StepError(step, f"unexpected controls: frame '{pm.text(frame)}', radio '{pm.text(bbr)}', "
                              f"button '{pm.text(adjust)}'.")
    if not inside(pm.rect(frame), pm.rect(bbr)) or not inside(pm.rect(frame), pm.rect(adjust)):
        raise StepError(step, "BBR / Adjust are not inside the Rewards frame.")
    def select_bbr(attempt):
        if not pm.checked(bbr):
            log(f"  [{step}] click BBR (id={BBR_RADIO})")
            pm.click(bbr)

    attempt_until(pm, step, "select BBR", select_bbr, lambda: pm.checked(bbr) and pm.enabled(adjust), T_VERIFY,
                  same_screen=screen_is((), main_enabled=True),
                  describe=lambda: f"BBR selected={pm.checked(bbr)}, Adjust enabled={pm.enabled(adjust)}",
                  target=bbr)
    wait_state(pm, step, popups=(), main_enabled=True, player_id=player_id, quiet=0.5,
               what="PM idle with BBR selected")
    log(f"  [{step}] checkpoint: BBR selected")

    step = begin_step(progress, "8_open_adjustment")
    check_state(pm, step, popups=(), main_enabled=True, player_id=player_id)
    if not pm.checked(bbr):
        raise StepError(step, "BBR is no longer selected.")
    def press_adjust(attempt):
        log(f"  [{step}] click Adjust (id={ADJUST_BUTTON})")
        pm.click(adjust)

    dialog = attempt_until(pm, step, "open Player Adjustment", press_adjust,
                           lambda: state_if(pm, popups=(ADJUST,), main_enabled=False, player_id=player_id),
                           T_ADJUST_OPEN,
                           same_screen=lambda state: screen_is((), main_enabled=True)(state) and pm.checked(bbr),
                           target=adjust).get(ADJUST)
    header = control(pm, step, dialog, ADJ_HEADER, "adjustment header", need_enabled=False)
    if pm.text(header).strip() != "BBR Adjustment":
        raise StepError(step, f"Player Adjustment is '{pm.text(header)}', expected 'BBR Adjustment'.")
    log(f"  [{step}] checkpoint: 'BBR Adjustment' is open")

    step = begin_step(progress, "9_add_bbr")
    check_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id)
    add = control(pm, step, dialog, ADJ_ADD_RADIO, "Add BBR", expect="Add BBR")
    others = [control(pm, step, dialog, cid, name, need_enabled=False)
              for cid, name in ((ADJ_SUBTRACT_RADIO, "Subtract BBR"), (ADJ_ZERO_RADIO, "Set BBR to 0"))]
    def select_add(attempt):
        if not pm.checked(add) or any(pm.checked(h) for h in others):
            click(pm, step, dialog, ADJ_ADD_RADIO, "Add BBR", expect="Add BBR")

    attempt_until(pm, step, "select Add BBR", select_add,
                  lambda: pm.checked(add) and not any(pm.checked(h) for h in others), T_VERIFY,
                  same_screen=screen_is((ADJUST,), main_enabled=False), target=add)
    log(f"  [{step}] checkpoint: Add BBR selected")

    step = begin_step(progress, "10_enter_adjustment")
    check_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id)
    amount_field = control(pm, step, dialog, ADJ_AMOUNT_EDIT, "Adjustment")
    current = money(pm.text(control(pm, step, dialog, ADJ_CURRENT, "Current Balance", need_enabled=False)))
    new_label = control(pm, step, dialog, ADJ_NEW, "New Balance", need_enabled=False)
    if current is None:
        raise StepError(step, "Current Balance is not readable.")
    expected_new = round(current + float(amount), 2)

    def new_balance_ok():
        value = money(pm.text(new_label))
        return value is not None and abs(value - expected_new) < 0.005

    # The amount is rewritten if PM did not take it (New Balance must follow it).
    attempt_until(pm, step, f"Adjustment = {amount}", lambda attempt: pm.set_text(amount_field, amount),
                  lambda: normalized_number_text(pm.text(amount_field)) == normalized_number_text(amount)
                  and new_balance_ok(), T_VERIFY,
                  same_screen=screen_is((ADJUST,), main_enabled=False),
                  describe=lambda: f"Adjustment shows '{pm.text(amount_field)}', New Balance shows "
                                   f"'{pm.text(new_label)}', expected {expected_new:.2f} = {current:.2f} + {amount}",
                  target=amount_field)
    log(f"  [{step}] checkpoint: Adjustment = {amount}")
    log(f"  [{step}] checkpoint: New Balance {current:.2f} + {amount} = {expected_new:.2f}")

    step = begin_step(progress, "11_set_expiration")
    check_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id)
    picker = control(pm, step, dialog, ADJ_EXPIRATION, "Expiration")
    def set_date(attempt):
        log(f"  [{step}] set Expiration to {expires:%m/%d/%Y %I:%M %p}")
        pm.date_set(picker, expires)

    attempt_until(pm, step, "Expiration", set_date, lambda: pm.date_get(picker) == expires, T_VERIFY,
                  same_screen=screen_is((ADJUST,), main_enabled=False),
                  describe=lambda: f"Expiration shows {_safe_done(lambda: pm.date_get(picker))}", target=picker)
    wait_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id, quiet=0.3,
               what="Player Adjustment only")
    log(f"  [{step}] checkpoint: Expiration = {expires:%m/%d/%Y %I:%M %p}")

    step = begin_step(progress, "12_select_reason")
    check_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id)
    reason_combo = control(pm, step, dialog, ADJ_REASON_COMBO, "Reason")
    select_combo(pm, step, reason_combo, wf["reason"], "Reason", screen=(ADJUST,))
    wait_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id,
               what="Player Adjustment only")
    log(f"  [{step}] checkpoint: Reason = {wf['reason']}")

    step = begin_step(progress, "13_enter_comment")
    check_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id)
    comment_field = control(pm, step, dialog, ADJ_COMMENT_EDIT, "Comment")
    set_and_verify(pm, step, comment_field, comment, "Comment", screen=(ADJUST,))
    wait_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id, quiet=0.5,
               what="no popup after the comment")

    step = begin_step(progress, "14_ok_adjustment" if allow_ok else "14_cancel_adjustment")
    check_state(pm, step, popups=(ADJUST,), main_enabled=False, player_id=player_id)
    problems = []
    if pm.text(header).strip() != "BBR Adjustment":
        problems.append("header")
    if not pm.checked(add) or any(pm.checked(h) for h in others):
        problems.append("Add BBR")
    if normalized_number_text(pm.text(amount_field)) != normalized_number_text(amount):
        problems.append("Adjustment")
    if money(pm.text(new_label)) is None or abs(money(pm.text(new_label)) - expected_new) >= 0.005:
        problems.append("New Balance")
    if pm.date_get(picker) != expires:
        problems.append("Expiration")
    if pm.combo_selected(reason_combo) != wf["reason"]:
        problems.append("Reason")
    if pm.text(comment_field).strip() != comment:
        problems.append("Comment")
    if active_profile_tab(pm, step, player_id) != progress["tab"]:
        problems.append("profile tab")
    if problems:
        raise StepError(step, f"changed before the final click: {', '.join(problems)}. Nothing was clicked.")
    log(f"  [{step}] final check OK: player {player_id}, Add BBR {amount}, expires "
        f"{expires:%m/%d/%Y %I:%M %p}, {wf['reason']}, '{comment}'")
    finish_dialog(pm, step, job, progress, dialog, allow_ok)


def process_job(pm, job, progress, allow_ok=False):
    open_profile(pm, job, progress)
    if WORKFLOWS[job["workflow"]]["kind"] == "coupon":
        run_coupon(pm, job, progress, allow_ok)
    else:
        run_bbr(pm, job, progress, allow_ok)


def close_tab_until_gone(pm, step, tab, what):
    def close(attempt):
        log(f"  [{step}] {what}")
        pm.close_tab(tab)

    attempt_until(pm, step, what, close, lambda: not pm.exists(tab) or not pm.visible(tab), T_TAB_CLOSE,
                  same_screen=screen_is((), main_enabled=True), target=tab)


def close_job_tab(pm, progress, player_id):
    """Close the player's profile tab right after the player is finished (done, cancelled or Loc)."""
    tab = progress["tab"]
    if not pm.exists(tab):
        return
    step = begin_step(progress, "15_close_profile_tab")
    check_state(pm, step, popups=(), main_enabled=True)
    title = pm.window_title(tab)
    if not title_has_player(title, player_id):
        raise StepError(step, f"tab #{tab} is '{title}', not player {player_id}; not closed.")
    close_tab_until_gone(pm, step, tab, f"close the tab of {player_id}")
    wait_state(pm, step, popups=(), main_enabled=True, timeout=T_TAB_CLOSE, quiet=0.5,
               what="PM idle after closing the tab")


def close_profile_tabs(pm, tabs):
    step = "13_close_profile_tabs"
    log(f"Closing {len(tabs)} profile tab(s) opened by this run.")
    closed = 0
    for tab, player_id in tabs:
        if not pm.exists(tab):
            continue
        title = pm.window_title(tab)
        if not title_has_player(title, player_id):
            log(f"  [{step}] skip tab #{tab}: title '{title}' is not player {player_id}")
            continue
        check_state(pm, step, popups=(), main_enabled=True)
        close_tab_until_gone(pm, step, tab, f"close tab '{title}'")
        wait_state(pm, step, popups=(), main_enabled=True, timeout=T_TAB_CLOSE, quiet=0.5,
                   what="PM idle after closing the tab")
        closed += 1
    log(f"Profile tabs closed: {closed}")


# ------------------------------------------------------------------ run

LEDGER_FIELDS = ["timestamp", "date", "workflow", "player_id", "amount", "target", "expiration", "comment",
                 "status", "sheet", "excel_row"]


def ledger_done_today():
    """(player, workflow) pairs with an OK click recorded today (OK_CLICKED or DONE)."""
    if not LEDGER_FILE.exists():
        return set()
    today = datetime.now().date().isoformat()
    with LEDGER_FILE.open(newline="", encoding="utf-8-sig") as file:
        return {(row["player_id"], row.get("workflow", "")) for row in csv.DictReader(file)
                if row.get("date") == today}


def ledger_append(job, status):
    new_file = not LEDGER_FILE.exists()
    with LEDGER_FILE.open("a", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=LEDGER_FIELDS)
        if new_file:
            writer.writeheader()
        now = datetime.now()
        writer.writerow({"timestamp": now.isoformat(timespec="seconds"), "date": now.date().isoformat(),
                         "workflow": job["workflow"], "player_id": job["player_id"], "amount": job["amount"],
                         "target": job["target"], "expiration": job["expiration"], "comment": job["comment"],
                         "status": status, "sheet": job["sheet"], "excel_row": job["excel_row"]})


def ask_monthly_comment(jobs, today):
    """Show the Monthly Benefit BBR comment; Enter keeps it, or type another one for all those rows."""
    monthly = [j for j in jobs if j["workflow"] == "MONTHLY_BBR"]
    if not monthly:
        return True
    suggested, standard = monthly_comment(today)
    print("\n" + "=" * 78)
    if not standard:
        print(f"WARNING: today ({today:%m/%d/%Y}) is not the 1st or the 15th of the month.")
        print(f"The Monthly Benefit comment would normally be for the 1st or the 15th; suggested: '{suggested}'.")
    print(f"Comment for the {len(monthly)} Monthly Benefit BBR row(s): '{suggested}'")
    answer = input("Press Enter to use it for all of them, or type another comment: ").strip()
    comment = answer or suggested
    if not comment:
        return False
    for job in monthly:
        job["comment"] = comment
    log(f"Monthly Benefit BBR comment: '{comment}'" + (" (typed by the user)" if answer else "")
        + ("" if standard else f" - today is not the 1st or the 15th, suggested was '{suggested}'"))
    return True


def ask_allow_ok(jobs):
    print("\n" + "=" * 78)
    for workflow in WORKFLOW_ORDER:
        part = [j for j in jobs if j["workflow"] == workflow]
        if not part:
            continue
        wf = WORKFLOWS[workflow]
        total = sum(float(j["amount"]) for j in part)
        print(f"#{wf['no']} {wf['label']}: {len(part)} job(s), total {normalized_number_text(total)}"
              f"  [{part[0]['target']}]" + (f" expires {part[0]['expiration']}, comment '{part[0]['comment']}'"
                                            if part[0]["expiration"] else ""))
        for job in part:
            print(f"     {job['sheet']} row {job['excel_row']:>4}  player {job['player_id']:>8}  amount {job['amount']}")
    print("=" * 78)
    answer = input("Allow clicking OK (REAL coupon redemption / BBR adjustment)?\n"
                   "Type YES to click OK, anything else = Cancel only (test run): ")
    return answer.strip() == "YES"


CSV_FIELDS = [
    "timestamp", "order", "workflow", "sheet", "excel_row", "player_id", "amount", "target", "expiration",
    "comment", "test_mode", "status", "failed_step", "message", "identification", "name_color", "stop_codes",
    "system_messages", "comment_pages", "confirmation", "confirmation_check", "duration_s", "screenshots",
]


def write_log(rows):
    with LOG_FILE.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def run(pm, jobs, username, password, allow_ok=False, today=None):
    """Process every job in order; stop at the first error and leave PM untouched."""
    today = today or date.today()
    log(f"Mode: {'OK - REAL REDEMPTION / ADJUSTMENT' if allow_ok else 'Cancel only (test run)'}; "
        f"today = {today:%m/%d/%Y}")
    done_today = ledger_done_today() if allow_ok else set()
    results = []
    RUN_FILES["results"] = results   # for the STATUS sheet
    RUN_FILES["loc"] = []            # SKIPPED_LOC rows, to be done by hand
    RUN_FILES.pop("stopped_at", None)
    tabs = []
    try:
        ensure_logged_in(pm, username, password)
    except Exception as exc:
        log(f"STOPPED at login: {exc}")
        note_problem(f"login: {exc}")
        write_log(results)
        return 1

    for index, job in enumerate(jobs, start=1):
        started = time.time()
        progress = {"step": "1_precheck_idle", "today": today}
        wf = WORKFLOWS[job["workflow"]]
        record = {
            "timestamp": datetime.now().isoformat(timespec="seconds"), "order": index,
            "workflow": job["workflow"], "sheet": job["sheet"], "excel_row": job["excel_row"],
            "player_id": job["player_id"], "amount": job["amount"], "target": job["target"],
            "expiration": job["expiration"], "comment": job["comment"], "test_mode": not allow_ok,
            "status": "", "failed_step": "", "message": "", "identification": "", "name_color": "",
            "stop_codes": "", "system_messages": "", "comment_pages": "", "confirmation": "",
            "confirmation_check": "", "duration_s": "", "screenshots": "",
        }
        log(f"[{index}/{len(jobs)}] #{wf['no']} {wf['label']}: player {job['player_id']}, amount {job['amount']} "
            f"({job['sheet']} row {job['excel_row']})")
        if (job["player_id"], job["workflow"]) in done_today:
            record.update({"status": "SKIPPED_ALREADY_DONE", "duration_s": 0,
                           "message": f"OK was already clicked today for this workflow (see {LEDGER_FILE.name})."})
            results.append(record)
            write_log(results)
            log("  skipped: already done today")
            continue
        try:
            process_job(pm, job, progress, allow_ok)
            record["status"] = "DONE" if allow_ok else "TEST_CANCELLED"
            record["message"] = ("Confirmed with OK." if allow_ok else "All fields filled and checked; cancelled.")
        except PlayerSkipped as skip:
            record["status"] = skip.status
            record["message"] = str(skip)
            log(f"  [{progress['step']}] {skip.status}: {skip}")
            if skip.status == "SKIPPED_LOC":
                RUN_FILES["loc"].append(record)
        except Exception as exc:
            failed_step = getattr(exc, "step", progress["step"])
            try:
                screen = read_state(pm).summary()
            except Exception:
                screen = "unavailable"
            record.update({
                "status": "ERROR_AFTER_OK" if progress.get("ok_clicked") else "ERROR",
                "failed_step": failed_step,
                "message": f"{type(exc).__name__}: {exc} | screen: {screen}",
                "screenshots": " ".join(save_screens(pm, failed_step, "stopped", getattr(exc, "hwnd", None),
                                                     everything=True)),
            })
        for key in ("identification", "name_color", "stop_codes", "system_messages", "comment_pages",
                    "confirmation", "confirmation_check"):
            record[key] = progress.get(key, "")
        record["duration_s"] = round(time.time() - started, 1)
        results.append(record)
        write_log(results)
        if progress.get("tab") and (progress["tab"], job["player_id"]) not in tabs:
            tabs.append((progress["tab"], job["player_id"]))
        if record["status"].startswith("ERROR"):
            log(f"#{wf['no']} player {job['player_id']}: ERROR at {record['failed_step']}: {record['message']}")
            RUN_FILES["stopped_at"] = index
            note_problem(f"#{index} {job['workflow']} player {job['player_id']} ({job['sheet']} row "
                         f"{job['excel_row']}) {record['status']} at {record['failed_step']}: "
                         f"{record['message'].split(' | screen:')[0][:300]}")
            if record["status"] == "ERROR_AFTER_OK":
                note_problem("OK was already clicked for this row: check it in PM before running again "
                             "(a rerun today skips it, see pm_redeemed_ledger.csv)")
            log("STOPPED. PM is left exactly as it is for inspection; no cleanup was done.")
            log(f"Log file: {LOG_FILE}")
            return 1
        log(f"#{wf['no']} player {job['player_id']}: {record['status']} in {record['duration_s']}s")
        if CLOSE_TAB_AFTER_EACH_PLAYER and progress.get("tab"):
            try:
                close_job_tab(pm, progress, job["player_id"])
                tabs.remove((progress["tab"], job["player_id"]))
            except ValueError:
                pass
            except Exception as exc:
                step = getattr(exc, "step", progress["step"])
                log(f"STOPPED while closing the tab of {job['player_id']} at {step}: {exc}")
                RUN_FILES["stopped_at"] = index
                note_problem(f"#{index} player {job['player_id']}: closing the profile tab failed: {exc}")
                save_screens(pm, step, "stopped", getattr(exc, "hwnd", None), everything=True)
                return 1

    if CLOSE_TABS_AT_END:
        try:
            close_profile_tabs(pm, tabs)
        except Exception as exc:
            log(f"STOPPED while closing tabs at {getattr(exc, 'step', 'close_tabs')}: {exc}")
            note_problem(f"closing the profile tabs: {exc}")
            save_screens(pm, getattr(exc, "step", "close_tabs"), "stopped", getattr(exc, "hwnd", None),
                         everything=True)
            return 1
    counts = {}
    for row in results:
        counts[row["status"]] = counts.get(row["status"], 0) + 1
    log(f"Done. {counts}")
    noted = [row for row in results if row.get("confirmation_check") not in ("", "OK")]
    if noted:
        log(f"Coupon confirmation notes on {len(noted)} row(s) (column confirmation_check): "
            + ", ".join(f"#{row['order']} player {row['player_id']}" for row in noted))
        for row in noted:
            note_problem(f"#{row['order']} player {row['player_id']} coupon confirmation: {row['confirmation_check']}")
    log(f"Log file: {LOG_FILE}")
    return 0


def main():
    try:
        from pm_credentials import PM_USERNAME, PM_PASSWORD
    except ImportError as exc:
        raise SystemExit("Missing pm_credentials.py in the same folder.") from exc

    start_run_files()
    try:
        code = run_main(PM_USERNAME, PM_PASSWORD)
    except KeyboardInterrupt:
        log("STOPPED with Ctrl+C. PM is left as it is.")
        note_problem("stopped with Ctrl+C")
        code = 1
    except Exception as exc:
        log(f"[FATAL] {type(exc).__name__}: {exc}")
        note_problem(f"{type(exc).__name__}: {exc}")
        code = 1
    finish_run_files(VIOLATION_FOLDER if code == 2 else ERROR_FOLDER if code else None)
    return code


def run_main(username, password):
    today = date.today()
    log(f"Runner {RUNNER_VERSION}, control mode: {CONTROL_MODE}")
    log(f"Log folder: {RUN_FILES.get('dir', BASE_DIR)}")
    jobs, notes, violations = read_plan(today)
    RUN_FILES["plan"] = (jobs, notes, violations)
    try:
        RUN_FILES["sources"] = read_source_tables()
    except Exception as exc:
        log(f"Could not read the Excel rows for the STATUS sheet: {exc}")
    check_file = write_check_file(jobs, notes, violations)
    log(f"Excel check written to {check_file.name}: {len(jobs)} job(s), {len(notes)} note(s), "
        f"{len(violations)} violation(s)")
    for note in notes:
        log(f"  NOTE {note['sheet']} row {note['row']} player {note['player_id']}: {note['note']}")
    if violations:
        for v in violations:
            log(f"  VIOLATION {v['rule']}: {v['sheet']} rows {v['rows']} player {v['player_id']} - {v['detail']}")
            note_problem(f"{v['rule']}: player {v['player_id']} ({v['sheet']} rows {v['rows']}) - {v['detail']}")
        log("STOPPED before touching PM. Fix the Excel file (see the VIOLATIONS sheet) and run again.")
        return 2
    if not jobs:
        log("Nothing to do.")
        write_log([])
        return 0
    if not ask_monthly_comment(jobs, today):
        log("STOPPED: no comment for the Monthly Benefit BBR rows.")
        note_problem("no comment for the Monthly Benefit BBR rows")
        return 1
    write_check_file(jobs, notes, violations)       # the PLAN shows the comment that will be used
    allow_ok = ask_allow_ok(jobs)
    pm = Win32PM()
    pm.connect()
    return run(pm, jobs, username, password, allow_ok, today)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception as exc:
        print(f"[FATAL] {type(exc).__name__}: {exc}")
        sys.exit(1)
