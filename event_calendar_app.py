# event_calendar_app.py
# -*- coding: utf-8 -*-
"""
Market Events Calendar (MVP)
- KRX(XKRX) + NYSE(XNYS) + NASDAQ(XNAS) exchange closed days using exchange_calendars
- Stores events in SQLite
- Tkinter UI:
  - Calendar(list)
  - Month(view): big cells + per-day coloring + clickable
  - Settings
  - Export CSV

핵심 안정화 포인트:
1) pandas timezone은 ZoneInfo 객체 대신 문자열("UTC","Asia/Seoul")로 통일
2) sessions tz-naive/tz-aware 케이스 모두 방어
3) "생성 전" provider+기간에 해당하는 기존 이벤트를 삭제(구버전 id 중복 제거)

Env:
- Python 3.10/3.11 권장
- pip install -U pandas exchange-calendars
"""

import os
import sys
import json
import csv
import sqlite3
import threading
import platform
from datetime import datetime, timedelta, date
from typing import List, Optional, Dict, Any, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import exchange_calendars as ecals
from exchange_calendars.errors import DateOutOfBounds
from zoneinfo import ZoneInfo

APP_NAME = "Market Events Calendar (MVP)"
DB_PATH = "events.db"
CFG_PATH = "config.json"

# UI/DB 저장 표준은 KST(표시) + UTC(원본)
LOCAL_TZ = ZoneInfo("Asia/Seoul")
UTC_TZ = ZoneInfo("UTC")

# pandas에서는 tz 문자열 사용(중요!)
PANDAS_UTC = "UTC"
PANDAS_KST = "Asia/Seoul"

DEFAULT_CFG = {
    "days_ahead": 60,
    "providers": {
        "krx_holidays": True,
        "nyse_holidays": True,
        "nasdaq_holidays": True
    }
}


# -------------------------
# Config
# -------------------------
def load_cfg() -> Dict[str, Any]:
    if not os.path.exists(CFG_PATH):
        save_cfg(DEFAULT_CFG)
        return json.loads(json.dumps(DEFAULT_CFG))
    with open(CFG_PATH, "r", encoding="utf-8") as f:
        try:
            cfg = json.load(f)
        except Exception:
            cfg = json.loads(json.dumps(DEFAULT_CFG))
    merged = json.loads(json.dumps(DEFAULT_CFG))
    merged.update(cfg)
    merged["providers"].update(cfg.get("providers", {}))
    return merged


def save_cfg(cfg: Dict[str, Any]) -> None:
    with open(CFG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# -------------------------
# DB
# -------------------------
def db_connect():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def db_init():
    conn = db_connect()
    conn.execute("""
    CREATE TABLE IF NOT EXISTS events (
        id TEXT PRIMARY KEY,
        provider TEXT NOT NULL,
        title TEXT NOT NULL,
        country TEXT,
        currency TEXT,
        importance TEXT,
        category TEXT,
        dt_utc TEXT NOT NULL,
        dt_local TEXT NOT NULL,
        source_url TEXT,
        raw_json TEXT,
        updated_at TEXT NOT NULL
    );
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_events_dt_local ON events(dt_local);")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_events_provider ON events(provider);")
    conn.commit()
    conn.close()


def db_upsert_events(rows: List[Dict[str, Any]]) -> int:
    if not rows:
        return 0
    conn = db_connect()
    cur = conn.cursor()
    now = datetime.now(tz=UTC_TZ).isoformat()
    count = 0
    for r in rows:
        cur.execute("""
        INSERT INTO events (id, provider, title, country, currency, importance, category,
                            dt_utc, dt_local, source_url, raw_json, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            title=excluded.title,
            country=excluded.country,
            currency=excluded.currency,
            importance=excluded.importance,
            category=excluded.category,
            dt_utc=excluded.dt_utc,
            dt_local=excluded.dt_local,
            source_url=excluded.source_url,
            raw_json=excluded.raw_json,
            updated_at=excluded.updated_at
        """, (
            r["id"], r["provider"], r["title"], r.get("country"), r.get("currency"),
            r.get("importance"), r.get("category"),
            r["dt_utc"], r["dt_local"], r.get("source_url"),
            json.dumps(r.get("raw", {}), ensure_ascii=False),
            now
        ))
        count += 1
    conn.commit()
    conn.close()
    return count


def db_delete_provider_range(provider: str, start_local: datetime, end_local: datetime) -> int:
    """생성 전: provider + 기간 범위에 해당하는 기존 이벤트 삭제 (구버전 id 중복 제거용)."""
    conn = db_connect()
    cur = conn.cursor()
    cur.execute(
        "DELETE FROM events WHERE provider = ? AND dt_local >= ? AND dt_local <= ?;",
        (provider, start_local.isoformat(), end_local.isoformat())
    )
    n = cur.rowcount
    conn.commit()
    conn.close()
    return n


def db_query_events(start_local: Optional[datetime] = None, end_local: Optional[datetime] = None,
                    keyword: str = "", country: str = "ALL", category: str = "ALL",
                    provider: str = "ALL", importance: str = "ALL") -> List[Dict[str, Any]]:
    conn = db_connect()
    cur = conn.cursor()
    where = []
    params = []

    if start_local:
        where.append("dt_local >= ?")
        params.append(start_local.isoformat())
    if end_local:
        where.append("dt_local <= ?")
        params.append(end_local.isoformat())

    if keyword.strip():
        where.append("(title LIKE ? OR country LIKE ? OR category LIKE ?)")
        k = f"%{keyword.strip()}%"
        params.extend([k, k, k])

    if country != "ALL":
        where.append("country = ?")
        params.append(country)
    if category != "ALL":
        where.append("category = ?")
        params.append(category)
    if provider != "ALL":
        where.append("provider = ?")
        params.append(provider)
    if importance != "ALL":
        where.append("importance = ?")
        params.append(importance)

    sql = "SELECT id, provider, title, country, currency, importance, category, dt_local, dt_utc, source_url, raw_json FROM events"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY dt_local ASC"

    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()

    out = []
    for row in rows:
        out.append({
            "id": row[0],
            "provider": row[1],
            "title": row[2],
            "country": row[3],
            "currency": row[4],
            "importance": row[5],
            "category": row[6],
            "dt_local": row[7],
            "dt_utc": row[8],
            "source_url": row[9],
            "raw_json": row[10],
        })
    return out


def db_distinct(field: str) -> List[str]:
    if field not in ("country", "importance", "category", "provider"):
        return []
    conn = db_connect()
    cur = conn.cursor()
    cur.execute(f"SELECT DISTINCT {field} FROM events WHERE {field} IS NOT NULL AND {field} != '' ORDER BY {field};")
    vals = [r[0] for r in cur.fetchall()]
    conn.close()
    return vals


def db_query_events_by_date_kst(d: date) -> List[Dict[str, Any]]:
    start = datetime(d.year, d.month, d.day, 0, 0, 0, tzinfo=LOCAL_TZ)
    end = datetime(d.year, d.month, d.day, 23, 59, 59, tzinfo=LOCAL_TZ)
    rows = db_query_events(start_local=start, end_local=end)
    filtered = []
    for ev in rows:
        try:
            dt_local = datetime.fromisoformat(ev["dt_local"])
            if dt_local.date() == d:
                filtered.append(ev)
        except Exception:
            pass
    return filtered


def db_query_events_by_month_kst(year: int, month: int) -> List[Dict[str, Any]]:
    first = date(year, month, 1)
    if month == 12:
        last = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        last = date(year, month + 1, 1) - timedelta(days=1)
    start = datetime(first.year, first.month, first.day, 0, 0, 0, tzinfo=LOCAL_TZ)
    end = datetime(last.year, last.month, last.day, 23, 59, 59, tzinfo=LOCAL_TZ)
    return db_query_events(start_local=start, end_local=end)


# -------------------------
# Exchange calendars
# -------------------------
def _calendar_supported_days(cal) -> Tuple[pd.Timestamp, pd.Timestamp]:
    # first_session/last_session are tz-aware (usually UTC). normalize keeps tz.
    # remove tz for comparison with user date range (tz-naive compare)
    first_day = cal.first_session.normalize().tz_localize(None)
    last_day = cal.last_session.normalize().tz_localize(None)
    return first_day, last_day


def _ensure_utc_index(idx: pd.DatetimeIndex) -> pd.DatetimeIndex:
    """sessions가 tz-naive or tz-aware 어느 경우든 UTC tz-aware로 통일."""
    if idx.tz is None:
        return idx.tz_localize(PANDAS_UTC)
    return idx.tz_convert(PANDAS_UTC)


def generate_exchange_closed_days(exchange_code: str, start_d: date, end_d: date) -> Tuple[List[pd.Timestamp], Dict[str, Any]]:
    """
    closed days = calendar days - open sessions
    pandas timezone은 문자열로 통일(중요): "UTC"
    """
    cal = ecals.get_calendar(exchange_code)
    first_day, last_day = _calendar_supported_days(cal)

    req_start = pd.Timestamp(start_d)
    req_end = pd.Timestamp(end_d)

    meta = {
        "exchange": exchange_code,
        "supported_first": first_day.date().isoformat(),
        "supported_last": last_day.date().isoformat(),
        "requested_start": req_start.date().isoformat(),
        "requested_end": req_end.date().isoformat(),
        "clamped_start": None,
        "clamped_end": None,
        "error": None,
    }

    if req_end < first_day or req_start > last_day:
        return [], meta

    start_clamped = max(req_start, first_day)
    end_clamped = min(req_end, last_day)
    meta["clamped_start"] = start_clamped.date().isoformat()
    meta["clamped_end"] = end_clamped.date().isoformat()

    try:
        # ✅ all_days: tz-naive -> tz_localize("UTC")
        all_days_naive = pd.date_range(start=start_clamped, end=end_clamped, freq="D")
        all_days_utc = all_days_naive.tz_localize(PANDAS_UTC)

        # ✅ sessions_in_range: tz-naive inputs
        sessions = cal.sessions_in_range(start_clamped, end_clamped)
        sessions_utc = _ensure_utc_index(sessions)

        # normalize keeps tz; unique -> ndarray, so wrap back to DatetimeIndex
        open_days = pd.DatetimeIndex(sessions_utc.normalize().unique())
        open_days = _ensure_utc_index(open_days)

        # ✅ difference는 tz가 완전히 동일해야 정상
        closed_days = all_days_utc.difference(open_days)

        # 방어: 혹시라도 tz-naive가 섞이면 UTC로 localize
        if closed_days.tz is None:
            closed_days = closed_days.tz_localize(PANDAS_UTC)

        return list(closed_days), meta

    except DateOutOfBounds as e:
        meta["error"] = f"DateOutOfBounds: {repr(e)}"
        return [], meta
    except Exception as e:
        meta["error"] = repr(e)
        return [], meta


def build_exchange_holiday_events(exchange_code: str, provider: str, title: str,
                                 country: str, currency: str,
                                 start_d: date, end_d: date) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    closed_days_utc -> events
    dt_local은 KST(Asia/Seoul)로 저장
    """
    closed_days_utc, meta = generate_exchange_closed_days(exchange_code, start_d, end_d)

    events = []
    for dts_utc in closed_days_utc:
        # ✅ tz-naive 방어 (오류 재발 방지)
        if getattr(dts_utc, "tz", None) is None:
            dts_utc = dts_utc.tz_localize(PANDAS_UTC)

        dts_kst = dts_utc.tz_convert(PANDAS_KST)  # pandas tz string

        wd = dts_kst.weekday()
        sub = "주말" if wd >= 5 else "휴장(공휴일/비거래일)"
        uid = f"{provider}|holiday|{dts_kst.date().isoformat()}"

        events.append({
            "id": uid,
            "provider": provider,
            "title": f"{title} ({sub})",
            "country": country,
            "currency": currency,
            "importance": "HIGH",
            "category": "Market Holiday",
            "dt_utc": dts_utc.to_pydatetime().isoformat(),
            "dt_local": dts_kst.to_pydatetime().isoformat(),
            "source_url": None,
            "raw": {"exchange": exchange_code, "date": dts_kst.date().isoformat(), **meta}
        })

    return events, meta


# -------------------------
# UI App
# -------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)

        # 세로형 UI
        self.geometry("980x920")
        self.minsize(900, 800)

        self.cfg = load_cfg()
        db_init()

        # Month grid cells
        self.month_cells: Dict[Tuple[int, int], tk.Label] = {}
        self.month_cell_dates: Dict[Tuple[int, int], Optional[date]] = {}

        self._build_ui()
        self._refresh_filters()
        self._load_table()
        self._render_month()

    def log(self, msg: str):
        ts = datetime.now(LOCAL_TZ).strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{ts}] {msg}\n")
        self.txt_log.see("end")

    def _build_ui(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        self.tab_calendar = ttk.Frame(self.nb)
        self.tab_month = ttk.Frame(self.nb)
        self.tab_settings = ttk.Frame(self.nb)
        self.tab_log = ttk.Frame(self.nb)

        self.nb.add(self.tab_calendar, text="Calendar")
        self.nb.add(self.tab_month, text="Month")
        self.nb.add(self.tab_settings, text="Settings")
        self.nb.add(self.tab_log, text="Log")

        # -------- Calendar(list) ----------
        top = ttk.Frame(self.tab_calendar)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Start:").pack(side="left")
        self.var_start = tk.StringVar(value=datetime.now(LOCAL_TZ).date().isoformat())
        ttk.Entry(top, textvariable=self.var_start, width=12).pack(side="left", padx=5)

        ttk.Label(top, text="End:").pack(side="left")
        self.var_end = tk.StringVar(value=(datetime.now(LOCAL_TZ).date() + timedelta(days=self.cfg.get("days_ahead", 60))).isoformat())
        ttk.Entry(top, textvariable=self.var_end, width=12).pack(side="left", padx=5)

        ttk.Label(top, text="Keyword:").pack(side="left", padx=(10, 0))
        self.var_keyword = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self.var_keyword, width=16).pack(side="left", padx=5)

        ttk.Button(top, text="Search", command=self._load_table).pack(side="left", padx=(10, 0))
        ttk.Button(top, text="Update (Generate)", command=self._update_generate).pack(side="left", padx=6)
        ttk.Button(top, text="Export CSV", command=self._export_csv).pack(side="left", padx=6)

        filt = ttk.Frame(self.tab_calendar)
        filt.pack(fill="x", padx=10, pady=(0, 8))

        ttk.Label(filt, text="Country:").pack(side="left")
        self.var_country = tk.StringVar(value="ALL")
        self.cmb_country = ttk.Combobox(filt, textvariable=self.var_country, width=16, state="readonly")
        self.cmb_country.pack(side="left", padx=5)

        ttk.Label(filt, text="Category:").pack(side="left", padx=(10, 0))
        self.var_category = tk.StringVar(value="ALL")
        self.cmb_category = ttk.Combobox(filt, textvariable=self.var_category, width=16, state="readonly")
        self.cmb_category.pack(side="left", padx=5)

        ttk.Label(filt, text="Provider:").pack(side="left", padx=(10, 0))
        self.var_provider = tk.StringVar(value="ALL")
        self.cmb_provider = ttk.Combobox(filt, textvariable=self.var_provider, width=16, state="readonly")
        self.cmb_provider.pack(side="left", padx=5)

        mid = ttk.Frame(self.tab_calendar)
        mid.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        cols = ("dt_local", "title", "country", "category", "importance", "provider")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=18)

        for c, t in [
            ("dt_local", "KST DateTime"),
            ("title", "Event"),
            ("country", "Country"),
            ("category", "Category"),
            ("importance", "Importance"),
            ("provider", "Provider"),
        ]:
            self.tree.heading(c, text=t)

        self.tree.column("dt_local", width=160, anchor="w")
        self.tree.column("title", width=380, anchor="w")
        self.tree.column("country", width=120, anchor="w")
        self.tree.column("category", width=140, anchor="w")
        self.tree.column("importance", width=90, anchor="w")
        self.tree.column("provider", width=120, anchor="w")

        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        bottom = ttk.LabelFrame(self.tab_calendar, text="Details")
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        self.txt_detail = tk.Text(bottom, height=6, wrap="word")
        self.txt_detail.pack(fill="x", padx=8, pady=6)

        # -------- Month (grid) ----------
        mtop = ttk.Frame(self.tab_month)
        mtop.pack(fill="x", padx=10, pady=10)

        ttk.Label(mtop, text="Year:").pack(side="left")
        self.var_m_year = tk.IntVar(value=datetime.now(LOCAL_TZ).year)
        ttk.Spinbox(mtop, from_=2000, to=2100, width=6, textvariable=self.var_m_year, command=self._render_month).pack(side="left", padx=5)

        ttk.Label(mtop, text="Month:").pack(side="left", padx=(10, 0))
        self.var_m_month = tk.IntVar(value=datetime.now(LOCAL_TZ).month)
        ttk.Spinbox(mtop, from_=1, to=12, width=4, textvariable=self.var_m_month, command=self._render_month).pack(side="left", padx=5)

        ttk.Button(mtop, text="◀ Prev", command=self._month_prev).pack(side="left", padx=(10, 4))
        ttk.Button(mtop, text="Next ▶", command=self._month_next).pack(side="left", padx=4)
        ttk.Button(mtop, text="Refresh", command=self._render_month).pack(side="left", padx=(10, 0))

        legend = ttk.Frame(self.tab_month)
        legend.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(legend, text="Legend:").pack(side="left")

        self.COL_KRX = "#ffe4b5"
        self.COL_NYSE = "#dbeafe"
        self.COL_NASDAQ = "#dcfce7"
        self.COL_MULTI = "#fca5a5"
        self.COL_NONE = "white"

        ttk.Label(legend, text=" KRX ", background=self.COL_KRX).pack(side="left", padx=6)
        ttk.Label(legend, text=" NYSE ", background=self.COL_NYSE).pack(side="left", padx=6)
        ttk.Label(legend, text=" NASDAQ ", background=self.COL_NASDAQ).pack(side="left", padx=6)
        ttk.Label(legend, text=" Multiple ", background=self.COL_MULTI).pack(side="left", padx=6)

        self.month_frame = ttk.Frame(self.tab_month)
        self.month_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # headers
        for c, name in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
            lbl = ttk.Label(self.month_frame, text=name, anchor="center")
            lbl.grid(row=0, column=c, sticky="nsew", padx=2, pady=2)

        # cells (bigger)
        for r in range(1, 7):
            for c in range(7):
                key = (r, c)
                cell = tk.Label(
                    self.month_frame,
                    text="",
                    anchor="nw",
                    justify="left",
                    bg=self.COL_NONE,
                    bd=1,
                    relief="solid",
                    padx=6,
                    pady=4,
                    font=("Segoe UI", 10),
                    wraplength=120,
                )
                cell.grid(row=r, column=c, sticky="nsew", padx=2, pady=2)
                cell.bind("<Button-1>", lambda e, k=key: self._on_month_cell_click(k))
                self.month_cells[key] = cell
                self.month_cell_dates[key] = None

        for c in range(7):
            self.month_frame.grid_columnconfigure(c, weight=1, minsize=120)
        for r in range(0, 7):
            if r == 0:
                self.month_frame.grid_rowconfigure(r, weight=0)
            else:
                self.month_frame.grid_rowconfigure(r, weight=1, minsize=95)

        self.month_detail = ttk.LabelFrame(self.tab_month, text="Selected Date Details")
        self.month_detail.pack(fill="x", padx=10, pady=(0, 10))
        self.txt_month_detail = tk.Text(self.month_detail, height=7, wrap="word")
        self.txt_month_detail.pack(fill="x", padx=8, pady=6)

        # -------- Settings ----------
        s = ttk.Frame(self.tab_settings)
        s.pack(fill="both", expand=True, padx=12, pady=12)

        ttk.Label(s, text="Days Ahead (default End):").grid(row=0, column=0, sticky="w")
        self.var_days = tk.StringVar(value=str(self.cfg.get("days_ahead", 60)))
        ttk.Entry(s, textvariable=self.var_days, width=10).grid(row=0, column=1, sticky="w", padx=8)

        self.var_p_krx = tk.BooleanVar(value=bool(self.cfg.get("providers", {}).get("krx_holidays", True)))
        self.var_p_nyse = tk.BooleanVar(value=bool(self.cfg.get("providers", {}).get("nyse_holidays", True)))
        self.var_p_nasdaq = tk.BooleanVar(value=bool(self.cfg.get("providers", {}).get("nasdaq_holidays", True)))

        ttk.Checkbutton(s, text="Include KRX Holidays (XKRX)", variable=self.var_p_krx).grid(row=1, column=0, columnspan=2, sticky="w", pady=(12, 0))
        ttk.Checkbutton(s, text="Include NYSE Holidays (XNYS)", variable=self.var_p_nyse).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Checkbutton(s, text="Include NASDAQ Holidays (XNAS)", variable=self.var_p_nasdaq).grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Button(s, text="Save Settings", command=self._save_settings).grid(row=4, column=0, pady=(18, 0), sticky="w")
        ttk.Button(s, text="Update Now", command=self._update_generate).grid(row=4, column=1, pady=(18, 0), sticky="w")

        # -------- Log ----------
        self.txt_log = tk.Text(self.tab_log, wrap="word")
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=10)

        self.log(f"Python: {sys.version.splitlines()[0]}")
        self.log(f"Executable: {sys.executable}")
        self.log(f"Platform: {platform.platform()}")
        try:
            import exchange_calendars
            self.log(f"exchange_calendars: {exchange_calendars.__version__}")
        except Exception:
            pass
        try:
            import pandas
            self.log(f"pandas: {pandas.__version__}")
        except Exception:
            pass
        self.log(f"Script: {os.path.abspath(__file__)}")

    def _refresh_filters(self):
        self.cmb_country["values"] = ["ALL"] + db_distinct("country")
        self.cmb_category["values"] = ["ALL"] + db_distinct("category")
        self.cmb_provider["values"] = ["ALL"] + db_distinct("provider")

        if self.var_country.get() not in self.cmb_country["values"]:
            self.var_country.set("ALL")
        if self.var_category.get() not in self.cmb_category["values"]:
            self.var_category.set("ALL")
        if self.var_provider.get() not in self.cmb_provider["values"]:
            self.var_provider.set("ALL")

    def _parse_date_input(self, s: str) -> Optional[date]:
        s = (s or "").strip()
        if not s:
            return None
        try:
            return date.fromisoformat(s)
        except Exception:
            return None

    def _load_table(self):
        start_d = self._parse_date_input(self.var_start.get())
        end_d = self._parse_date_input(self.var_end.get())

        start_dt = datetime(start_d.year, start_d.month, start_d.day, 0, 0, 0, tzinfo=LOCAL_TZ) if start_d else None
        end_dt = datetime(end_d.year, end_d.month, end_d.day, 23, 59, 59, tzinfo=LOCAL_TZ) if end_d else None

        rows = db_query_events(
            start_local=start_dt,
            end_local=end_dt,
            keyword=self.var_keyword.get(),
            country=self.var_country.get(),
            category=self.var_category.get(),
            provider=self.var_provider.get(),
        )

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        for r in rows:
            try:
                dt_local = datetime.fromisoformat(r["dt_local"])
                dt_str = dt_local.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                dt_str = r["dt_local"]

            self.tree.insert("", "end", iid=r["id"], values=(
                dt_str,
                r["title"],
                r.get("country") or "",
                r.get("category") or "",
                r.get("importance") or "",
                r["provider"]
            ))

        self._refresh_filters()
        self.log(f"Loaded {len(rows)} events.")
        self._render_month()

    def _on_select(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        event_id = sel[0]
        conn = db_connect()
        cur = conn.cursor()
        cur.execute("SELECT title, country, currency, importance, category, dt_local, dt_utc, source_url, raw_json, provider FROM events WHERE id = ?;", (event_id,))
        row = cur.fetchone()
        conn.close()
        if not row:
            return

        title, country, currency, imp, cat, dt_local, dt_utc, url, raw_json, provider = row
        try:
            raw = json.loads(raw_json) if raw_json else {}
        except Exception:
            raw = {}

        self.txt_detail.delete("1.0", "end")
        self.txt_detail.insert("end", f"Title: {title}\n")
        self.txt_detail.insert("end", f"Provider: {provider}\n")
        self.txt_detail.insert("end", f"KST : {dt_local}\nUTC : {dt_utc}\n")
        self.txt_detail.insert("end", f"Country: {country or ''}   Currency: {currency or ''}   Importance: {imp or ''}\n")
        self.txt_detail.insert("end", f"Category: {cat or ''}\n")
        if url:
            self.txt_detail.insert("end", f"Source: {url}\n")
        if raw:
            self.txt_detail.insert("end", "\n--- RAW ---\n")
            self.txt_detail.insert("end", json.dumps(raw, ensure_ascii=False, indent=2))

    def _save_settings(self):
        try:
            self.cfg["days_ahead"] = int(self.var_days.get().strip())
        except Exception:
            self.cfg["days_ahead"] = 60

        self.cfg["providers"]["krx_holidays"] = bool(self.var_p_krx.get())
        self.cfg["providers"]["nyse_holidays"] = bool(self.var_p_nyse.get())
        self.cfg["providers"]["nasdaq_holidays"] = bool(self.var_p_nasdaq.get())
        save_cfg(self.cfg)

        try:
            start_d = date.fromisoformat(self.var_start.get().strip())
        except Exception:
            start_d = datetime.now(LOCAL_TZ).date()
            self.var_start.set(start_d.isoformat())

        self.var_end.set((start_d + timedelta(days=self.cfg["days_ahead"])).isoformat())
        messagebox.showinfo("Saved", "Settings saved.")
        self.log("Settings saved.")

    def _update_generate(self):
        def worker():
            try:
                cfg = load_cfg()
                providers = cfg.get("providers", {})

                start_d = self._parse_date_input(self.var_start.get())
                end_d = self._parse_date_input(self.var_end.get())
                if not start_d or not end_d:
                    self.after(0, lambda: messagebox.showerror("Invalid date", "Start/End 날짜를 YYYY-MM-DD 형식으로 입력해줘."))
                    return
                if end_d < start_d:
                    self.after(0, lambda: messagebox.showerror("Invalid range", "End 날짜가 Start보다 빠를 수 없어."))
                    return

                # 삭제 범위(로컬)
                start_dt = datetime(start_d.year, start_d.month, start_d.day, 0, 0, 0, tzinfo=LOCAL_TZ)
                end_dt = datetime(end_d.year, end_d.month, end_d.day, 23, 59, 59, tzinfo=LOCAL_TZ)

                self.log(f"Generate range: {start_d.isoformat()} ~ {end_d.isoformat()}")

                total = 0

                def run_provider(exchange_code: str, provider: str, title: str, country: str, currency: str) -> int:
                    # ✅ 기존 provider+기간 데이터 삭제(중복 제거)
                    deleted = db_delete_provider_range(provider, start_dt, end_dt)
                    if deleted:
                        self.log(f"Deleted {deleted} old rows for provider={provider} in range.")

                    self.log(f"Generating {provider} ({exchange_code})...")
                    ev, meta = build_exchange_holiday_events(exchange_code, provider, title, country, currency, start_d, end_d)
                    if meta.get("error"):
                        self.log(f"[{exchange_code} ERROR] {meta['error']}")
                        self.after(0, lambda: messagebox.showerror(f"{exchange_code} Error", meta["error"]))
                        return 0
                    n = db_upsert_events(ev)
                    self.log(f"Upserted {n} rows for provider={provider}.")
                    return n

                if providers.get("krx_holidays", True):
                    total += run_provider("XKRX", "krx_holidays", "KRX 휴장", "South Korea", "KRW")
                if providers.get("nyse_holidays", True):
                    total += run_provider("XNYS", "nyse_holidays", "NYSE 휴장", "United States", "USD")
                if providers.get("nasdaq_holidays", True):
                    total += run_provider("XNAS", "nasdaq_holidays", "NASDAQ 휴장", "United States", "USD")

                self.log(f"Done. Total upserted: {total}")
                self.after(0, self._load_table)

            except Exception as e:
                msg = repr(e)
                self.log(f"Error: {msg}")
                self.after(0, lambda: messagebox.showerror("Error", msg))

        threading.Thread(target=worker, daemon=True).start()

    def _export_csv(self):
        rows = []
        for iid in self.tree.get_children():
            rows.append(self.tree.item(iid)["values"])
        if not rows:
            messagebox.showinfo("No data", "내보낼 데이터가 없어.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="events_export.csv"
        )
        if not path:
            return

        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["KST DateTime", "Event", "Country", "Category", "Importance", "Provider"])
            for r in rows:
                w.writerow(r)

        self.log(f"Exported CSV: {path}")
        messagebox.showinfo("Exported", f"CSV 저장 완료:\n{path}")

    # -------- Month view ----------
    def _month_prev(self):
        y = self.var_m_year.get()
        m = self.var_m_month.get() - 1
        if m <= 0:
            m = 12
            y -= 1
        self.var_m_year.set(y)
        self.var_m_month.set(m)
        self._render_month()

    def _month_next(self):
        y = self.var_m_year.get()
        m = self.var_m_month.get() + 1
        if m >= 13:
            m = 1
            y += 1
        self.var_m_year.set(y)
        self.var_m_month.set(m)
        self._render_month()

    def _render_month(self):
        y = int(self.var_m_year.get())
        m = int(self.var_m_month.get())

        first = date(y, m, 1)
        if m == 12:
            last = date(y + 1, 1, 1) - timedelta(days=1)
        else:
            last = date(y, m + 1, 1) - timedelta(days=1)

        events = db_query_events_by_month_kst(y, m)

        dmap: Dict[date, set] = {}
        for ev in events:
            try:
                dt_local = datetime.fromisoformat(ev["dt_local"])
                dmap.setdefault(dt_local.date(), set()).add(ev["provider"])
            except Exception:
                pass

        for key, cell in self.month_cells.items():
            cell.config(text="", bg=self.COL_NONE)
            self.month_cell_dates[key] = None

        start_wd = first.weekday()  # Mon=0
        total_days = (last - first).days + 1

        day_num = 1
        for week in range(1, 7):
            for dow in range(7):
                idx = (week - 1) * 7 + dow
                key = (week, dow)
                cell = self.month_cells[key]

                if idx >= start_wd and day_num <= total_days:
                    dcur = date(y, m, day_num)
                    self.month_cell_dates[key] = dcur

                    provs = dmap.get(dcur, set())
                    marks = []
                    if "krx_holidays" in provs:
                        marks.append("KRX")
                    if "nyse_holidays" in provs:
                        marks.append("NYSE")
                    if "nasdaq_holidays" in provs:
                        marks.append("NASDAQ")

                    if marks:
                        text = f"{day_num}\n" + "\n".join(marks)
                    else:
                        text = f"{day_num}"

                    kinds = 0
                    if "krx_holidays" in provs: kinds += 1
                    if "nyse_holidays" in provs: kinds += 1
                    if "nasdaq_holidays" in provs: kinds += 1

                    if kinds >= 2:
                        bg = self.COL_MULTI
                    elif "krx_holidays" in provs:
                        bg = self.COL_KRX
                    elif "nyse_holidays" in provs:
                        bg = self.COL_NYSE
                    elif "nasdaq_holidays" in provs:
                        bg = self.COL_NASDAQ
                    else:
                        bg = self.COL_NONE

                    cell.config(text=text, bg=bg)
                    day_num += 1

        self.txt_month_detail.delete("1.0", "end")
        self.txt_month_detail.insert("end", f"{y}-{m:02d} month view loaded.\n")
        self.txt_month_detail.insert("end", "Click a date cell to see details.\n")

    def _on_month_cell_click(self, key: Tuple[int, int]):
        d = self.month_cell_dates.get(key)
        if not d:
            return
        rows = db_query_events_by_date_kst(d)

        self.txt_month_detail.delete("1.0", "end")
        self.txt_month_detail.insert("end", f"Date: {d.isoformat()} (KST)\n\n")
        if not rows:
            self.txt_month_detail.insert("end", "No events.\n")
            return
        for ev in rows[:120]:
            self.txt_month_detail.insert("end", f"- [{ev['provider']}] {ev['title']} ({ev.get('country','')})\n")


if __name__ == "__main__":
    db_init()
    app = App()
    app.mainloop()
