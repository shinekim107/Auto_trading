# lw_strategy_gui_tabs_v3.py
# - Email settings input in GUI (optional save to state)
# - Daily candlestick chart (Up=Red / Down=Blue) + Volume bars
# - Up to ~1Y range selectable (3M/6M/1Y) + Zoom/Pan via matplotlib toolbar
# - Top summary line: current price / breakout / signal / auto qty / volume
# - Avg price fix: prefer "매입가" column; fallback to 매입금액/수량
# - Manual buttons simplified: SELL ALL only
# - Next-open sell with holiday handling: pending_sell_date rolls forward if today is non-trading day
# - TR limit 대응: 5초마다 잔고/미체결 번갈아 갱신
# - Enhanced fill email content on full fill (unfilled==0)

import sys
import os
import json
import smtplib
from email.message import EmailMessage
from datetime import datetime, time, timedelta

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QSpinBox, QComboBox, QCheckBox,
    QMessageBox, QTableWidget, QTableWidgetItem, QTabWidget, QSplitter,
    QSizePolicy, QGroupBox
)

import matplotlib
matplotlib.use("Qt5Agg")
import matplotlib.patches as mpatches
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

from pykiwoom.kiwoom import Kiwoom


# -------------------------
# Settings
# -------------------------
STATE_PATH = "state_lw_strategy.json"

MARKET_OPEN = time(9, 0, 0)
MARKET_CLOSE = time(15, 30, 0)
NEXTOPEN_SELL_TIME = time(9, 0, 10)

REFRESH_SEC = 5


# -------------------------
# Utils
# -------------------------
def to_int(x) -> int:
    try:
        return int(str(x).replace(",", "").strip())
    except:
        return 0


def norm_code(x: str) -> str:
    s = str(x).strip()
    return s[1:] if s.startswith("A") else s


def yyyymmdd(dt: datetime = None) -> str:
    dt = dt or datetime.now()
    return dt.strftime("%Y%m%d")


def now_dt() -> datetime:
    return datetime.now()


def is_market_time_clock_only() -> bool:
    """시계 기준 장중(휴장일 여부 미반영)"""
    t = now_dt().time()
    return MARKET_OPEN <= t <= MARKET_CLOSE


def reached_time(target: time) -> bool:
    return now_dt().time() >= target


def load_state() -> dict:
    if os.path.exists(STATE_PATH):
        try:
            with open(STATE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}


def save_state(state: dict) -> None:
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def calc_qty_by_budget(price: int, budget_krw: int) -> int:
    if price <= 0:
        return 0
    return max(0, budget_krw // price)


def find_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


# -------------------------
# Email (GUI input based)
# -------------------------
def send_email_gui(cfg: dict, subject: str, body: str) -> None:
    """
    cfg keys:
      host, port, user, app_password, to
    """
    host = (cfg.get("host") or "").strip() or "smtp.gmail.com"
    port = int(cfg.get("port") or 587)
    user = (cfg.get("user") or "").strip()
    pw = (cfg.get("app_password") or "").strip()
    to_raw = (cfg.get("to") or "").strip()

    if not (user and pw and to_raw):
        return  # 미설정이면 스킵(매매 로직을 막지 않기)

    recipients = [x.strip() for x in to_raw.split(",") if x.strip()]
    if not recipients:
        return

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

    with smtplib.SMTP(host, port, timeout=15) as s:
        s.ehlo()
        s.starttls()
        s.login(user, pw)
        s.send_message(msg)


# -------------------------
# Daily Candle + Volume Chart
# -------------------------
class DailyCandleChart(QWidget):
    """
    - Daily OHLC candlesticks
    - Up candle: red, Down candle: blue
    - Volume bars in background (twin y)
    - Uses matplotlib toolbar for zoom/pan
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.fig = Figure(figsize=(10, 4.6))
        self.canvas = FigureCanvas(self.fig)
        self.toolbar = NavigationToolbar(self.canvas, self)

        lay = QVBoxLayout()
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.toolbar)
        lay.addWidget(self.canvas)
        self.setLayout(lay)

        self.ax = None
        self.axv = None

    def plot(self, candles, breakout_price=None, title=""):
        self.fig.clf()
        self.ax = self.fig.add_subplot(111)
        self.axv = self.ax.twinx()

        if not candles:
            self.ax.set_title(title + " (no data)")
            self.ax.grid(True, alpha=0.3)
            self.canvas.draw()
            return

        xs = list(range(len(candles)))
        vols = [cd.get("v", 0) for cd in candles]

        # Volume bars (subtle)
        self.axv.bar(xs, vols, width=0.6, alpha=0.15)
        self.axv.set_yticks([])
        self.axv.set_ylabel("Volume", rotation=270, labelpad=12)

        # Candles
        for i, cd in enumerate(candles):
            o, h, l, c = cd["o"], cd["h"], cd["l"], cd["c"]
            up = (c >= o)
            col = "red" if up else "blue"

            # wick
            self.ax.vlines(i, l, h, linewidth=1, colors=col)

            # body
            body_low = min(o, c)
            body_high = max(o, c)
            body_h = max(1, body_high - body_low)  # avoid zero height
            self.ax.add_patch(
                mpatches.Rectangle(
                    (i - 0.3, body_low),
                    0.6,
                    body_h,
                    fill=True,
                    edgecolor=col,
                    facecolor=col,
                    linewidth=1,
                    alpha=0.85
                )
            )

        # x labels (sparse)
        labels = [cd["t"].strftime("%m-%d") for cd in candles]
        step = max(1, len(labels) // 10)
        self.ax.set_xticks(xs[::step])
        self.ax.set_xticklabels(labels[::step], rotation=0)

        if breakout_price is not None and breakout_price > 0:
            self.ax.axhline(breakout_price, linestyle="--", linewidth=1)
            self.ax.text(0, breakout_price, f" breakout={breakout_price:,}", va="bottom")

        self.ax.set_title(title)
        self.ax.set_ylabel("Price")
        self.ax.grid(True, alpha=0.25)
        self.fig.tight_layout()
        self.canvas.draw()


# -------------------------
# Main GUI
# -------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LW Strategy GUI (Daily Candles + Volume + Email in GUI)")

        self.kiwoom = Kiwoom()
        self.logged_in = False
        self.account = None

        self.state = load_state()

        self.last_price = None
        self.last_volume = None
        self.breakout_price = None
        self.breakout_ref_date = None
        self.fail_count = 0

        self._refresh_toggle = 0  # 0: unfilled, 1: balance (alternate)
        self._cached_candles = []
        self._cached_candles_code = None
        self._cached_candles_n = None

        # 휴장/거래일 캐시(하루 1회 체크)
        self._trading_day_cache_date = None
        self._trading_day_cache_is_open = None

        # timers
        self.timer = QTimer(self)
        self.timer.setInterval(REFRESH_SEC * 1000)
        self.timer.timeout.connect(self.on_tick)

        # UI root
        root = QWidget()
        self.setCentralWidget(root)
        root_layout = QVBoxLayout()
        root.setLayout(root_layout)

        # Login row
        row0 = QHBoxLayout()
        self.lbl_login = QLabel("Login: (not connected)")
        self.btn_login = QPushButton("Login")
        row0.addWidget(self.lbl_login)
        row0.addStretch(1)
        row0.addWidget(self.btn_login)
        root_layout.addLayout(row0)

        # ✅ Top summary line
        self.lbl_summary = QLabel("요약 | 현재가: - | 돌파가: - | Signal: - | 자동수량: - | 거래량: -")
        self.lbl_summary.setStyleSheet("font-size: 13px; font-weight: 600; color:#111;")
        root_layout.addWidget(self.lbl_summary)

        # Tabs
        self.tabs = QTabWidget()
        root_layout.addWidget(self.tabs)

        self.lbl_status = QLabel("Status: -")
        self.lbl_status.setStyleSheet("font-size: 12px; color: #333;")
        root_layout.addWidget(self.lbl_status)

        self.tab_order = QWidget()
        self.tab_tables = QWidget()
        self.tabs.addTab(self.tab_order, "주문/조회(전략)")
        self.tabs.addTab(self.tab_tables, "잔고/미체결")

        self._build_tab_order()
        self._build_tab_tables_lr()

        # Wiring
        self.btn_login.clicked.connect(self.do_login)
        self.btn_query.clicked.connect(self.query_price_manual)

        self.cmb_k.currentIndexChanged.connect(self.on_params_changed)
        self.chk_strategy.toggled.connect(self.on_params_changed)
        self.chk_real.toggled.connect(self.on_params_changed)
        self.spin_budget.valueChanged.connect(self.on_params_changed)
        self.ed_code.editingFinished.connect(self.on_params_changed)

        self.btn_refresh_chart.clicked.connect(self.reload_daily_candles)
        self.cmb_range.currentIndexChanged.connect(self.reload_daily_candles)
        self.chk_autoscale_y.toggled.connect(self.reload_daily_candles)

        self.btn_balance.clicked.connect(self.refresh_balance)
        self.btn_unfilled.clicked.connect(self.refresh_unfilled)

        self.btn_sell_all.clicked.connect(self.sell_all_position)

        self.btn_test_email.clicked.connect(self.send_test_email)
        self.chk_save_email.toggled.connect(self.on_params_changed)

        self.set_enabled_trading(False)
        self._update_status_labels(initial=True)

    # -------------------------
    # UI build
    # -------------------------
    def _build_tab_order(self):
        v = QVBoxLayout()
        self.tab_order.setLayout(v)

        # row: code + price query + sell all
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("종목코드:"))
        self.ed_code = QLineEdit(self.state.get("code", "122630"))
        self.ed_code.setFixedWidth(120)

        self.btn_query = QPushButton("조회(현재가)")
        self.lbl_price = QLabel("현재가: -")

        self.btn_sell_all = QPushButton("전량 매도(시장가)")
        self.btn_sell_all.setToolTip("현재 종목코드 보유수량 전량 시장가 매도")

        row1.addWidget(self.ed_code)
        row1.addWidget(self.btn_query)
        row1.addSpacing(10)
        row1.addWidget(self.lbl_price)
        row1.addStretch(1)
        row1.addWidget(self.btn_sell_all)
        v.addLayout(row1)

        # Strategy group
        g = QGroupBox("Larry Williams 변동성 돌파 전략")
        gv = QVBoxLayout()
        g.setLayout(gv)

        rowS1 = QHBoxLayout()
        self.chk_strategy = QCheckBox("전략 실행 ON")
        self.chk_strategy.setChecked(bool(self.state.get("strategy_on", False)))
        rowS1.addWidget(self.chk_strategy)

        self.chk_real = QCheckBox("REAL ORDER ON")
        self.chk_real.setChecked(bool(self.state.get("real_on", False)))
        rowS1.addWidget(self.chk_real)

        rowS1.addSpacing(10)
        rowS1.addWidget(QLabel("예산(원):"))
        self.spin_budget = QSpinBox()
        self.spin_budget.setRange(10_000, 200_000_000)
        self.spin_budget.setSingleStep(10_000)
        self.spin_budget.setValue(int(self.state.get("budget", 1_000_000)))
        self.spin_budget.setFixedWidth(140)
        rowS1.addWidget(self.spin_budget)

        rowS1.addSpacing(10)
        rowS1.addWidget(QLabel("k:"))
        self.cmb_k = QComboBox()
        ks = [round(i / 10, 1) for i in range(1, 11)]  # 0.1~1.0
        for k in ks:
            self.cmb_k.addItem(f"{k:.1f}", k)
        rowS1.addWidget(self.cmb_k)

        self.lbl_preset = QLabel("프리셋: 보수적 0.3 / 기본 0.5~0.6 / 공격적 0.8")
        self.lbl_preset.setStyleSheet("color:#666;")
        rowS1.addSpacing(10)
        rowS1.addWidget(self.lbl_preset)

        rowS1.addStretch(1)
        gv.addLayout(rowS1)

        # init k
        self._set_k_combobox(float(self.state.get("k", 0.6)))

        # status row
        rowS2 = QHBoxLayout()
        self.lbl_market = QLabel("장상태: -")
        self.lbl_breakout = QLabel("돌파가: -")
        self.lbl_signal = QLabel("Signal: -")
        self.lbl_qty = QLabel("자동수량: -")
        self.lbl_fail = QLabel("연속실패: 0")
        for lb in [self.lbl_market, self.lbl_breakout, self.lbl_signal, self.lbl_qty, self.lbl_fail]:
            lb.setStyleSheet("font-size: 12px;")
        rowS2.addWidget(self.lbl_market)
        rowS2.addSpacing(8)
        rowS2.addWidget(self.lbl_breakout)
        rowS2.addSpacing(8)
        rowS2.addWidget(self.lbl_signal)
        rowS2.addSpacing(8)
        rowS2.addWidget(self.lbl_qty)
        rowS2.addSpacing(8)
        rowS2.addWidget(self.lbl_fail)
        rowS2.addStretch(1)
        gv.addLayout(rowS2)

        gv.addWidget(QLabel("• 매수: 현재가 ≥ 돌파가 → 시장가 매수(예산 기반 수량)\n• 매도: 다음 영업일 09:00:10 이후 시장가 매도(휴장일이면 다음날로 자동 이월)"))
        v.addWidget(g)

        # Email group
        ge = QGroupBox("체결 알림 이메일 (GUI 입력)")
        gev = QVBoxLayout()
        ge.setLayout(gev)

        rowE1 = QHBoxLayout()
        rowE1.addWidget(QLabel("SMTP Host:"))
        self.ed_smtp_host = QLineEdit(self.state.get("email_host", "smtp.gmail.com"))
        self.ed_smtp_host.setFixedWidth(180)
        rowE1.addWidget(self.ed_smtp_host)

        rowE1.addWidget(QLabel("Port:"))
        self.spin_smtp_port = QSpinBox()
        self.spin_smtp_port.setRange(1, 65535)
        self.spin_smtp_port.setValue(int(self.state.get("email_port", 587)))
        self.spin_smtp_port.setFixedWidth(90)
        rowE1.addWidget(self.spin_smtp_port)

        rowE1.addSpacing(10)
        self.chk_save_email = QCheckBox("이메일 설정 state 저장")
        self.chk_save_email.setChecked(bool(self.state.get("save_email", False)))
        rowE1.addWidget(self.chk_save_email)

        rowE1.addStretch(1)
        gev.addLayout(rowE1)

        rowE2 = QHBoxLayout()
        rowE2.addWidget(QLabel("User:"))
        self.ed_email_user = QLineEdit(self.state.get("email_user", ""))
        self.ed_email_user.setFixedWidth(260)
        rowE2.addWidget(self.ed_email_user)

        rowE2.addWidget(QLabel("App PW:"))
        self.ed_email_pw = QLineEdit(self.state.get("email_pw", ""))
        self.ed_email_pw.setEchoMode(QLineEdit.Password)
        self.ed_email_pw.setFixedWidth(220)
        rowE2.addWidget(self.ed_email_pw)

        rowE2.addWidget(QLabel("To:"))
        self.ed_email_to = QLineEdit(self.state.get("email_to", ""))
        self.ed_email_to.setFixedWidth(260)
        rowE2.addWidget(self.ed_email_to)

        self.btn_test_email = QPushButton("테스트 메일")
        rowE2.addWidget(self.btn_test_email)

        rowE2.addStretch(1)
        gev.addLayout(rowE2)

        v.addWidget(ge)

        # Chart controls
        gc = QGroupBox("차트")
        gcv = QVBoxLayout()
        gc.setLayout(gcv)

        rowC1 = QHBoxLayout()
        rowC1.addWidget(QLabel("기간:"))
        self.cmb_range = QComboBox()
        self.cmb_range.addItems(["3M", "6M", "1Y"])
        self.cmb_range.setCurrentText(self.state.get("chart_range", "1Y"))
        rowC1.addWidget(self.cmb_range)

        self.chk_autoscale_y = QCheckBox("Y축 자동 스케일")
        self.chk_autoscale_y.setChecked(bool(self.state.get("chart_autoscale_y", True)))
        rowC1.addWidget(self.chk_autoscale_y)

        self.btn_refresh_chart = QPushButton("일봉 새로고침")
        rowC1.addWidget(self.btn_refresh_chart)

        rowC1.addStretch(1)
        rowC1.addWidget(QLabel("※ 줌/이동: 차트 상단 툴바 사용"))
        gcv.addLayout(rowC1)

        self.chart = DailyCandleChart()
        gcv.addWidget(self.chart)

        v.addWidget(gc)

        self.lbl_trade_log = QLabel("TradeLog: -")
        self.lbl_trade_log.setStyleSheet("font-size: 12px; color:#222;")
        v.addWidget(self.lbl_trade_log)

        v.addStretch(1)

    def _build_tab_tables_lr(self):
        v = QVBoxLayout()
        self.tab_tables.setLayout(v)

        row = QHBoxLayout()
        self.btn_balance = QPushButton("잔고조회")
        self.btn_unfilled = QPushButton("미체결조회")
        self.lbl_autorefresh = QLabel(f"AutoRefresh: OFF ({REFRESH_SEC}s, alternate)")
        self.lbl_autorefresh.setStyleSheet("color:#666;")
        row.addWidget(self.btn_balance)
        row.addWidget(self.btn_unfilled)
        row.addSpacing(10)
        row.addWidget(self.lbl_autorefresh)
        row.addStretch(1)
        v.addLayout(row)

        splitter = QSplitter(Qt.Horizontal)
        v.addWidget(splitter)

        # Balance
        balance_wrap = QWidget()
        bl = QVBoxLayout()
        bl.setContentsMargins(0, 0, 0, 0)
        balance_wrap.setLayout(bl)
        bl.addWidget(QLabel("잔고 (OPW00018)"))
        self.tbl_balance = QTableWidget(0, 5)
        self.tbl_balance.setHorizontalHeaderLabels(["종목코드", "종목명", "수량", "매입금액", "평단가"])
        self.tbl_balance.horizontalHeader().setStretchLastSection(True)
        self.tbl_balance.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        bl.addWidget(self.tbl_balance)

        # Unfilled
        unfilled_wrap = QWidget()
        ul = QVBoxLayout()
        ul.setContentsMargins(0, 0, 0, 0)
        unfilled_wrap.setLayout(ul)
        ul.addWidget(QLabel("미체결 (OPT10075)"))
        self.tbl_unfilled = QTableWidget(0, 6)
        self.tbl_unfilled.setHorizontalHeaderLabels(["주문번호", "종목", "수량", "미체결", "가격", "구분"])
        self.tbl_unfilled.horizontalHeader().setStretchLastSection(True)
        self.tbl_unfilled.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        ul.addWidget(self.tbl_unfilled)

        splitter.addWidget(balance_wrap)
        splitter.addWidget(unfilled_wrap)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([560, 560])

    # -------------------------
    # Enable/Disable
    # -------------------------
    def set_enabled_trading(self, enabled: bool):
        self.btn_query.setEnabled(enabled)
        self.btn_sell_all.setEnabled(enabled)

        self.chk_strategy.setEnabled(enabled)
        self.chk_real.setEnabled(enabled)
        self.spin_budget.setEnabled(enabled)
        self.cmb_k.setEnabled(enabled)

        self.btn_balance.setEnabled(enabled)
        self.btn_unfilled.setEnabled(enabled)

        self.btn_refresh_chart.setEnabled(enabled)
        self.cmb_range.setEnabled(enabled)
        self.chk_autoscale_y.setEnabled(enabled)

        self.ed_smtp_host.setEnabled(enabled)
        self.spin_smtp_port.setEnabled(enabled)
        self.ed_email_user.setEnabled(enabled)
        self.ed_email_pw.setEnabled(enabled)
        self.ed_email_to.setEnabled(enabled)
        self.btn_test_email.setEnabled(enabled)
        self.chk_save_email.setEnabled(enabled)

    # -------------------------
    # Summary line
    # -------------------------
    def update_summary(self, price=None, breakout=None, signal=None, qty=None, vol=None):
        def fmt(x):
            return "-" if x is None else f"{x:,}"
        s = (
            f"요약 | 현재가: {fmt(price)} | 돌파가: {fmt(breakout)} | "
            f"Signal: {signal or '-'} | 자동수량: {fmt(qty)} | 거래량: {fmt(vol)}"
        )
        self.lbl_summary.setText(s)

    # -------------------------
    # Login + Event hook
    # -------------------------
    def do_login(self):
        try:
            self.kiwoom.CommConnect(block=True)
            self.logged_in = True

            accs_raw = self.kiwoom.GetLoginInfo("ACCNO")
            if isinstance(accs_raw, str):
                accs = [a for a in accs_raw.split(";") if a.strip()]
            else:
                accs = [a for a in accs_raw if str(a).strip()]
            self.account = accs[0] if accs else None
            if not self.account:
                raise RuntimeError("계좌를 찾지 못했습니다. (GetLoginInfo('ACCNO') 결과 없음)")

            # Chejan event hook
            try:
                self.kiwoom.ocx.OnReceiveChejanData.connect(self.on_chejan)
            except Exception as e:
                self.lbl_status.setText(f"Status: 체결 이벤트 연결 실패(ocx). {repr(e)}")

            self.lbl_login.setText(f"Login: OK  |  Account: {self.account}")
            self.set_enabled_trading(True)
            self.lbl_status.setText("Status: 로그인 완료")

            # Initial load
            self.refresh_unfilled()
            self.refresh_balance()
            self.reload_daily_candles()

            # Start timer
            self.timer.start()
            self.lbl_autorefresh.setText(f"AutoRefresh: ON ({REFRESH_SEC}s, alternate)")

        except Exception as e:
            QMessageBox.critical(self, "Login Error", repr(e))

    # -------------------------
    # Params / state
    # -------------------------
    def _set_k_combobox(self, k: float):
        best_idx = 0
        best_diff = 999
        for i in range(self.cmb_k.count()):
            kv = float(self.cmb_k.itemData(i))
            d = abs(kv - k)
            if d < best_diff:
                best_diff = d
                best_idx = i
        self.cmb_k.setCurrentIndex(best_idx)

    def on_params_changed(self):
        self.breakout_price = None  # recalc
        self._save_state()
        # chart might depend on code/range
        if self.logged_in:
            self.reload_daily_candles()

    def _save_state(self):
        st = self.state
        st["strategy_on"] = bool(self.chk_strategy.isChecked())
        st["real_on"] = bool(self.chk_real.isChecked())
        st["budget"] = int(self.spin_budget.value())
        st["k"] = float(self.cmb_k.currentData())
        st["code"] = norm_code(self.ed_code.text())

        st["chart_range"] = self.cmb_range.currentText()
        st["chart_autoscale_y"] = bool(self.chk_autoscale_y.isChecked())

        st["save_email"] = bool(self.chk_save_email.isChecked())
        if st["save_email"]:
            st["email_host"] = self.ed_smtp_host.text().strip()
            st["email_port"] = int(self.spin_smtp_port.value())
            st["email_user"] = self.ed_email_user.text().strip()
            st["email_pw"] = self.ed_email_pw.text()
            st["email_to"] = self.ed_email_to.text().strip()
        else:
            for k in ["email_host", "email_port", "email_user", "email_pw", "email_to"]:
                if k in st:
                    del st[k]

        save_state(st)

    def _email_cfg_from_gui(self) -> dict:
        return {
            "host": self.ed_smtp_host.text().strip(),
            "port": int(self.spin_smtp_port.value()),
            "user": self.ed_email_user.text().strip(),
            "app_password": self.ed_email_pw.text(),
            "to": self.ed_email_to.text().strip(),
        }

    # -------------------------
    # Chart
    # -------------------------
    def _range_to_n(self) -> int:
        r = self.cmb_range.currentText().strip().upper()
        if r == "3M":
            return 63
        if r == "6M":
            return 126
        return 240

    def reload_daily_candles(self):
        try:
            if not self.logged_in:
                return
            code = norm_code(self.ed_code.text())
            if not code.isdigit():
                return

            n = self._range_to_n()
            candles = self.load_daily_candles(code, n)
            self._cached_candles = candles
            self._cached_candles_code = code
            self._cached_candles_n = n

            # If autoscale y is off, keep current y-limits if possible
            ylim = None
            if not self.chk_autoscale_y.isChecked() and self.chart.ax is not None:
                try:
                    ylim = self.chart.ax.get_ylim()
                except:
                    ylim = None

            title = f"{code} Daily Candles ({self.cmb_range.currentText()})"
            self.chart.plot(candles, breakout_price=self.breakout_price, title=title)

            if ylim is not None and self.chart.ax is not None:
                self.chart.ax.set_ylim(ylim)
                self.chart.canvas.draw()

        except Exception as e:
            self.lbl_status.setText(f"Status: chart reload error - {repr(e)}")

    def load_daily_candles(self, code: str, n: int):
        df = self.kiwoom.block_request(
            "OPT10081",
            종목코드=code,
            기준일자=yyyymmdd(),
            수정주가구분="1",
            output="주식일봉차트조회",
            next=0
        )
        if "일자" not in df.columns:
            raise RuntimeError(f"Missing '일자' column: {list(df.columns)}")

        df = df.copy()
        df["일자"] = df["일자"].astype(str)
        for col in ["시가", "고가", "저가", "현재가", "거래량"]:
            if col in df.columns:
                df[col] = df[col].apply(lambda v: abs(to_int(v)))

        df = df.sort_values("일자").reset_index(drop=True).tail(n).reset_index(drop=True)

        candles = []
        has_v = ("거래량" in df.columns)
        for _, r in df.iterrows():
            d = datetime.strptime(str(r["일자"]), "%Y%m%d")
            candles.append({
                "t": d,
                "o": int(r["시가"]),
                "h": int(r["고가"]),
                "l": int(r["저가"]),
                "c": int(r["현재가"]),
                "v": int(r["거래량"]) if has_v else 0,
            })
        return candles

    # -------------------------
    # Price & breakout (LW)
    # -------------------------
    def get_current_price_and_volume(self, code: str):
        df = self.kiwoom.block_request(
            "OPT10001",
            종목코드=code,
            output="주식기본정보",
            next=0
        )
        if len(df) == 0:
            return 0, 0

        # price
        if "현재가" not in df.columns:
            raise RuntimeError(f"Missing '현재가' column: {list(df.columns)}")
        price = abs(to_int(df.iloc[0]["현재가"]))

        # volume (환경에 따라 다름)
        vol = 0
        if "거래량" in df.columns:
            vol = abs(to_int(df.iloc[0]["거래량"]))
        elif "누적거래량" in df.columns:
            vol = abs(to_int(df.iloc[0]["누적거래량"]))

        return price, vol

    def calc_breakout(self, code: str, k: float) -> int:
        df = self.kiwoom.block_request(
            "OPT10081",
            종목코드=code,
            기준일자=yyyymmdd(),
            수정주가구분="1",
            output="주식일봉차트조회",
            next=0
        )
        if "일자" not in df.columns:
            raise RuntimeError(f"Missing '일자' column: {list(df.columns)}")

        df = df.copy()
        df["일자"] = df["일자"].astype(str)
        for col in ["시가", "고가", "저가"]:
            if col in df.columns:
                df[col] = df[col].apply(lambda v: abs(to_int(v)))

        df = df.sort_values("일자").reset_index(drop=True)
        if len(df) < 2:
            raise RuntimeError("Need at least 2 daily bars")

        y = df.iloc[-2]
        t = df.iloc[-1]
        self.breakout_ref_date = str(t["일자"])

        today_open = int(t["시가"])
        y_range = int(y["고가"]) - int(y["저가"])
        return int(round(today_open + k * y_range))

    # -------------------------
    # Trading-day 판단 & pending sell rolling
    # -------------------------
    def is_trading_day_today_cached(self, code: str) -> bool:
        """
        휴장일 판단(캐시):
        - 오늘 날짜 기준으로 1회만 TR로 확인
        """
        today = yyyymmdd()
        if self._trading_day_cache_date == today and self._trading_day_cache_is_open is not None:
            return bool(self._trading_day_cache_is_open)

        try:
            df = self.kiwoom.block_request(
                "OPT10081",
                종목코드=code,
                기준일자=today,
                수정주가구분="1",
                output="주식일봉차트조회",
                next=0
            )
            if "일자" not in df.columns or len(df) == 0:
                is_open = False
            else:
                # 응답 정렬이 역순/정순 섞일 수 있어서 안전하게 max
                latest = max([str(x).strip() for x in df["일자"].tolist()])
                is_open = (latest == today)
        except:
            is_open = False

        self._trading_day_cache_date = today
        self._trading_day_cache_is_open = is_open
        return bool(is_open)

    def roll_pending_sell_if_holiday(self, code: str):
        """
        pending_sell_date==오늘인데 휴장일이면 -> 내일로 이월.
        연휴/연속휴장도 매일 이월되어 결국 첫 거래일에 실행됨.
        """
        st = self.state
        pending = st.get("pending_sell_date")
        if not pending:
            return
        today = yyyymmdd()
        if pending != today:
            return

        if not self.is_trading_day_today_cached(code):
            nd = (now_dt().date() + timedelta(days=1))
            st["pending_sell_date"] = nd.strftime("%Y%m%d")
            st["sell_roll_reason"] = "holiday_or_closed"
            save_state(st)
            self._log_trade(f"[ROLL] pending_sell_date -> {st['pending_sell_date']} (holiday)")
            self.lbl_market.setText("장상태: 휴장(대기)")

    # -------------------------
    # Tables: balance/unfilled
    # -------------------------
    def refresh_balance(self):
        df = self.kiwoom.block_request(
            "OPW00018",
            계좌번호=self.account,
            비밀번호="",
            비밀번호입력매체구분="00",
            조회구분="2",
            output="계좌평가잔고개별합산",
            next=0
        )

        c_code = find_col(df, ["종목번호", "종목코드"])
        c_name = find_col(df, ["종목명"])
        c_qty = find_col(df, ["보유수량"])
        c_buyamt = find_col(df, ["매입금액", "매입금액합", "매입금액합계", "매입금액(합)"])

        # ✅ 평단가 우선: 매입가
        c_avg = find_col(df, [
            "매입가",
            "평균단가", "평균단가 ", "평단가", "매입단가", "평균매입가", "평균매입가 "
        ])

        if not (c_code and c_name and c_qty):
            raise RuntimeError(f"잔고 컬럼을 찾지 못했습니다. columns={list(df.columns)}")

        rows = []
        for _, r in df.iterrows():
            code = norm_code(r[c_code])
            name = str(r[c_name]).strip()
            qty = abs(to_int(r[c_qty]))
            if qty <= 0:
                continue

            buy_amt = abs(to_int(r[c_buyamt])) if c_buyamt else 0
            avg = abs(to_int(r[c_avg])) if c_avg else 0

            # ✅ fallback: 매입금액/수량
            if avg <= 0 and buy_amt > 0 and qty > 0:
                avg = int(round(buy_amt / qty))

            rows.append((code, name, qty, buy_amt, avg))

        self.tbl_balance.setRowCount(len(rows))
        input_code = norm_code(self.ed_code.text())
        for i, (code, name, qty, buy_amt, avg) in enumerate(rows):
            items = [
                QTableWidgetItem(str(code)),
                QTableWidgetItem(str(name)),
                QTableWidgetItem(f"{qty:,}"),
                QTableWidgetItem(f"{buy_amt:,}"),
                QTableWidgetItem(f"{avg:,}" if avg else "-"),
            ]
            for it in items:
                it.setFlags(it.flags() ^ Qt.ItemIsEditable)
            for j, it in enumerate(items):
                self.tbl_balance.setItem(i, j, it)

            if input_code and input_code.isdigit() and code == input_code:
                for j in range(5):
                    self.tbl_balance.item(i, j).setBackground(Qt.yellow)

        if c_avg is None:
            self.lbl_status.setText(f"Status: 평단가 컬럼 미탐지. columns={list(df.columns)}")

    def refresh_unfilled(self):
        df = self.kiwoom.block_request(
            "OPT10075",
            계좌번호=self.account,
            전체종목구분="0",
            매매구분="0",
            종목코드="",
            체결구분="1",
            output="미체결",
            next=0
        )

        c_ordno = find_col(df, ["주문번호", "주문번호 "])
        c_name = find_col(df, ["종목명", "종목명 "])
        c_gubun = find_col(df, ["주문구분", "주문구분 "])
        c_price = find_col(df, ["주문가격", "주문가격 "])
        c_qty = find_col(df, ["주문수량", "주문수량 "])
        c_unfilled = find_col(df, ["미체결수량", "미체결수량 "])

        if not (c_ordno and c_qty and c_unfilled):
            raise RuntimeError(f"미체결 컬럼을 찾지 못했습니다. columns={list(df.columns)}")

        rows = []
        for _, r in df.iterrows():
            ordno = str(r[c_ordno]).strip()
            name = str(r[c_name]).strip() if c_name else ""
            gubun = str(r[c_gubun]).strip() if c_gubun else ""
            gubun = gubun.replace("+", "").strip()

            qty = abs(to_int(r[c_qty])) if c_qty else 0
            unfilled = abs(to_int(r[c_unfilled])) if c_unfilled else 0
            price = abs(to_int(r[c_price])) if c_price else 0
            if unfilled <= 0:
                continue
            rows.append((ordno, name, qty, unfilled, price, gubun))

        self.tbl_unfilled.setRowCount(len(rows))
        for i, (ordno, name, qty, unfilled, price, gubun) in enumerate(rows):
            items = [
                QTableWidgetItem(str(ordno)),
                QTableWidgetItem(str(name)),
                QTableWidgetItem(f"{qty:,}"),
                QTableWidgetItem(f"{unfilled:,}"),
                QTableWidgetItem(f"{price:,}"),
                QTableWidgetItem(str(gubun)),
            ]
            for it in items:
                it.setFlags(it.flags() ^ Qt.ItemIsEditable)
            for j, it in enumerate(items):
                self.tbl_unfilled.setItem(i, j, it)

    # -------------------------
    # Strategy loop (5s) with alternate TR refresh
    # -------------------------
    def on_tick(self):
        if not self.logged_in or not self.account:
            return

        code = norm_code(self.ed_code.text())
        if not code.isdigit():
            return

        try:
            # TR alternate to reduce rate-limit pressure
            if self._refresh_toggle % 2 == 0:
                self.refresh_unfilled()
            else:
                self.refresh_balance()
            self._refresh_toggle += 1

            # 휴장일이면 pending sell 이월
            self.roll_pending_sell_if_holiday(code)

            # 매도 체크(다음날 시초)
            self.try_next_open_sell(code)

            # 전략 실행 / 상태 업데이트
            self.run_strategy_step(code)

            self.fail_count = 0
            self.lbl_fail.setText("연속실패: 0")

        except Exception as e:
            self.fail_count += 1
            self.lbl_fail.setText(f"연속실패: {self.fail_count}")
            self.lbl_status.setText(f"Status: tick error - {repr(e)}")
            # 다음 tick에서 재시도(멈추지 않음)

    def run_strategy_step(self, code: str):
        # current price/volume (even if strategy off, we can update summary)
        price, vol = self.get_current_price_and_volume(code)
        self.last_price = price
        self.last_volume = vol
        if price > 0:
            self.lbl_price.setText(f"현재가: {price:,}")

        # breakout cached (recalc if None)
        k = float(self.cmb_k.currentData())
        if self.breakout_price is None:
            try:
                self.breakout_price = self.calc_breakout(code, k)
            except:
                self.breakout_price = None

        breakout = self.breakout_price
        budget = int(self.spin_budget.value())
        qty = calc_qty_by_budget(price, budget) if price else 0

        # market status (clock-only)
        if is_market_time_clock_only():
            self.lbl_market.setText("장상태: 장중(전략 판단)")
        else:
            self.lbl_market.setText("장상태: 장외(대기)")

        # signal determination (basic)
        if breakout and price:
            sig = "BREAKOUT" if price >= breakout else "WAIT"
        else:
            sig = "-"

        # update per-row labels
        if breakout:
            self.lbl_breakout.setText(f"돌파가: {breakout:,} (k={k:.1f}, ref={self.breakout_ref_date})")
        else:
            self.lbl_breakout.setText("돌파가: -")

        self.lbl_qty.setText(f"자동수량: {qty:,} (예산 {budget:,})" if qty else "자동수량: -")

        # chart refresh (daily candles; not every tick necessary, but safe)
        if self._cached_candles_code != code or self._cached_candles_n != self._range_to_n():
            self.reload_daily_candles()
        else:
            title = f"{code} Daily Candles ({self.cmb_range.currentText()})"
            self.chart.plot(self._cached_candles, breakout_price=breakout, title=title)

        # strategy ON/OFF + market time
        if not self.chk_strategy.isChecked():
            self.lbl_signal.setText("Signal: OFF")
            self.update_summary(price=price, breakout=breakout, signal="OFF", qty=qty, vol=vol)
            return

        if not is_market_time_clock_only():
            self.lbl_signal.setText("Signal: WAIT (장외)")
            self.update_summary(price=price, breakout=breakout, signal="WAIT(장외)", qty=qty, vol=vol)
            return

        # strategy decision
        if sig == "BREAKOUT":
            self.lbl_signal.setText("Signal: BREAKOUT (BUY 조건)")
            self.try_breakout_buy(code, price, qty)
        elif sig == "WAIT":
            self.lbl_signal.setText("Signal: WAIT")
        else:
            self.lbl_signal.setText("Signal: -")

        self.update_summary(price=price, breakout=breakout, signal=sig, qty=qty, vol=vol)

    # -------------------------
    # Buy / Sell logic
    # -------------------------
    def try_breakout_buy(self, code: str, price: int, qty: int):
        st = self.state
        today = yyyymmdd()

        # 하루 1회 매수: "체결 완료" 기준
        if st.get("buy_filled_date") == today:
            return

        if qty <= 0:
            return

        # 이미 오늘 매수 주문 pending이면 중복 방지
        if st.get("buy_sent_date") == today and st.get("buy_order_pending", False):
            return

        # paper
        if not self.chk_real.isChecked():
            st["buy_sent_date"] = today
            st["buy_order_pending"] = False
            st["buy_paper"] = True
            st["buy_snapshot_price"] = price
            st["buy_qty"] = qty
            st["buy_code"] = code
            # 다음날(캘린더) 예약. 휴장일이면 roll_pending_sell_if_holiday가 이월
            st["pending_sell_date"] = (now_dt().date() + timedelta(days=1)).strftime("%Y%m%d")
            save_state(st)
            self._log_trade(f"[PAPER] BUY breakout code={code} qty={qty} price={price:,} t={now_dt()}")
            self.tabs.setCurrentIndex(1)
            return

        ret = self.kiwoom.SendOrder(
            "LW_AUTO_BUY",
            "0201",
            self.account,
            1,
            code,
            qty,
            0,
            "03",
            ""
        )
        st["buy_sent_date"] = today
        st["buy_order_pending"] = True
        st["buy_send_ret"] = ret
        st["buy_snapshot_price"] = price
        st["buy_qty"] = qty
        st["buy_code"] = code
        st["buy_sent_time"] = now_dt().strftime("%H:%M:%S")

        st["pending_sell_date"] = (now_dt().date() + timedelta(days=1)).strftime("%Y%m%d")
        save_state(st)

        self._log_trade(f"[AUTO] BUY sent ret={ret} code={code} qty={qty} price={price:,} t={now_dt()}")
        self.tabs.setCurrentIndex(1)

    def try_next_open_sell(self, code: str):
        st = self.state
        today = yyyymmdd()

        pending = st.get("pending_sell_date")
        if not pending or pending != today:
            return

        # 시계상 장중 & 09:00:10 이후
        if not is_market_time_clock_only():
            return
        if not reached_time(NEXTOPEN_SELL_TIME):
            return

        if st.get("sell_filled_date") == today:
            return
        if st.get("sell_sent_date") == today and st.get("sell_order_pending", False):
            return

        qty = self.get_position_qty(code)
        if qty <= 0:
            st["sell_filled_date"] = today
            st["sell_reason"] = "no_position"
            st["sell_order_pending"] = False
            save_state(st)
            self._log_trade(f"[AUTO] SELL skip (no position) t={now_dt()}")
            return

        if not self.chk_real.isChecked():
            st["sell_sent_date"] = today
            st["sell_order_pending"] = False
            st["sell_paper"] = True
            st["sell_qty"] = qty
            save_state(st)
            self._log_trade(f"[PAPER] SELL next-open code={code} qty={qty} t={now_dt()}")
            self.tabs.setCurrentIndex(1)
            return

        ret = self.kiwoom.SendOrder(
            "LW_AUTO_SELL",
            "0202",
            self.account,
            2,
            code,
            qty,
            0,
            "03",
            ""
        )
        st["sell_sent_date"] = today
        st["sell_order_pending"] = True
        st["sell_send_ret"] = ret
        st["sell_qty"] = qty
        st["sell_code"] = code
        st["sell_sent_time"] = now_dt().strftime("%H:%M:%S")
        save_state(st)

        self._log_trade(f"[AUTO] SELL sent ret={ret} code={code} qty={qty} t={now_dt()}")
        self.tabs.setCurrentIndex(1)

    def get_position_qty(self, code: str) -> int:
        df = self.kiwoom.block_request(
            "OPW00018",
            계좌번호=self.account,
            비밀번호="",
            비밀번호입력매체구분="00",
            조회구분="2",
            output="계좌평가잔고개별합산",
            next=0
        )
        code_col = find_col(df, ["종목번호", "종목코드"])
        qty_col = find_col(df, ["보유수량"])
        if not code_col or not qty_col:
            return 0

        df2 = df.copy()
        df2[code_col] = df2[code_col].apply(norm_code)
        rows = df2[df2[code_col] == norm_code(code)]
        if rows.empty:
            return 0
        return abs(to_int(rows.iloc[0][qty_col]))

    # -------------------------
    # Manual: Sell all only
    # -------------------------
    def sell_all_position(self):
        try:
            if not self.logged_in or not self.account:
                return
            code = norm_code(self.ed_code.text())
            if not code.isdigit():
                return
            qty = self.get_position_qty(code)
            if qty <= 0:
                self._log_trade("[MANUAL] 전량매도: 보유수량 0")
                return

            if not self.chk_real.isChecked():
                self._log_trade(f"[PAPER] SELL ALL skip code={code} qty={qty}")
                return

            ret = self.kiwoom.SendOrder(
                "MAN_SELL_ALL",
                "0112",
                self.account,
                2,
                code,
                qty,
                0,
                "03",
                ""
            )
            self._log_trade(f"[MANUAL] SELL ALL sent ret={ret} code={code} qty={qty} t={now_dt()}")
            self.tabs.setCurrentIndex(1)
        except Exception as e:
            self._log_trade(f"[MANUAL] sell-all error {repr(e)}")

    # -------------------------
    # Chejan: fill-based state confirm + enhanced email
    # -------------------------
    def on_chejan(self, gubun, item_cnt, fid_list):
        try:
            if str(gubun) != "0":
                return

            get = self.kiwoom.GetChejanData

            order_no = str(get(9203)).strip()
            code = norm_code(str(get(9001)).strip())
            name = str(get(302)).strip()
            order_gubun = str(get(905)).strip().replace("+", "").strip()  # 매수/매도
            order_qty = abs(to_int(get(900)))
            unfilled_qty = abs(to_int(get(902)))
            filled_qty = abs(to_int(get(911)))      # 이번 체결량
            filled_price = abs(to_int(get(910)))    # 이번 체결가

            self._log_trade(
                f"[CHEJAN] {order_gubun} {name}({code}) ordno={order_no} "
                f"ord={order_qty} unfilled={unfilled_qty} fill={filled_qty}@{filled_price}"
            )

            # 완전 체결: 미체결수량 == 0
            if order_qty > 0 and unfilled_qty == 0:
                st = self.state
                today = yyyymmdd()

                k = float(self.cmb_k.currentData())
                budget = int(self.spin_budget.value())
                bprice = self.breakout_price or 0
                pending_sell = st.get("pending_sell_date")
                last_price = self.last_price or 0
                last_vol = self.last_volume or 0

                email_subject = ""
                email_body = (
                    f"Time: {now_dt()}\n"
                    f"Type: FILLED\n"
                    f"Gubun: {order_gubun}\n"
                    f"Code: {code}\n"
                    f"Name: {name}\n"
                    f"OrderNo: {order_no}\n"
                    f"OrderQty: {order_qty}\n"
                    f"UnfilledQty: {unfilled_qty}\n"
                    f"FilledQty(last): {filled_qty}\n"
                    f"FilledPrice(last): {filled_price}\n"
                    f"\n[Market Snapshot]\n"
                    f"LastPrice(OPT10001): {last_price:,}\n"
                    f"LastVolume(OPT10001): {last_vol:,}\n"
                    f"\n[Strategy]\n"
                    f"k: {k:.1f}\n"
                    f"Budget: {budget:,}\n"
                    f"BreakoutPrice: {bprice:,}\n"
                    f"BreakoutRefDate: {self.breakout_ref_date}\n"
                    f"PendingSellDate(calendar/rolled): {pending_sell}\n"
                )

                if "매수" in order_gubun:
                    st["buy_filled_date"] = today
                    st["buy_order_pending"] = False
                    st["buy_order_no"] = order_no
                    save_state(st)
                    email_subject = f"[AUTO TRADE] BUY FILLED {code}"

                elif "매도" in order_gubun:
                    st["sell_filled_date"] = today
                    st["sell_order_pending"] = False
                    st["sell_order_no"] = order_no
                    save_state(st)
                    email_subject = f"[AUTO TRADE] SELL FILLED {code}"
                else:
                    return

                try:
                    send_email_gui(self._email_cfg_from_gui(), email_subject, email_body)
                except Exception as e:
                    st["email_error_last"] = repr(e)
                    save_state(st)

        except Exception as e:
            self.lbl_status.setText(f"Status: chejan error - {repr(e)}")

    # -------------------------
    # Manual price query
    # -------------------------
    def query_price_manual(self):
        try:
            code = norm_code(self.ed_code.text())
            if not code.isdigit():
                raise ValueError("종목코드는 숫자여야 합니다.")
            price, vol = self.get_current_price_and_volume(code)
            self.last_price = price
            self.last_volume = vol
            self.lbl_price.setText(f"현재가: {price:,}")
            self.lbl_status.setText(f"Status: 현재가/거래량 조회 성공 ({code})")
            self.update_summary(price=price, breakout=self.breakout_price, signal="-", qty=None, vol=vol)
        except Exception as e:
            self.lbl_status.setText(f"Status: 현재가 조회 실패 - {repr(e)}")

    # -------------------------
    # Test email
    # -------------------------
    def send_test_email(self):
        try:
            cfg = self._email_cfg_from_gui()
            subject = "[TEST] LW Strategy Email"
            body = f"Time: {now_dt()}\nThis is a test email from LW Strategy GUI."
            send_email_gui(cfg, subject, body)
            self.lbl_status.setText("Status: 테스트 메일 전송 시도 완료(오류 없으면 성공)")
        except Exception as e:
            self.lbl_status.setText(f"Status: 테스트 메일 실패 - {repr(e)}")

    # -------------------------
    # UI status init
    # -------------------------
    def _update_status_labels(self, initial=False):
        if initial:
            self.lbl_market.setText("장상태: -")
            self.lbl_breakout.setText("돌파가: -")
            self.lbl_signal.setText("Signal: -")
            self.lbl_qty.setText("자동수량: -")
            self.lbl_fail.setText("연속실패: 0")

    def _log_trade(self, msg: str):
        self.lbl_trade_log.setText(f"TradeLog: {msg}")
        self.lbl_status.setText(f"Status: {msg}")

    def closeEvent(self, event):
        try:
            self._save_state()
        except:
            pass
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(1300, 900)
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
