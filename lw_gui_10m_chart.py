# lw_gui_10m_chart.py

# - Timeframe 선택: 10-Min Candles (Today) / Daily Candles (Last 120 Days)

# - Larry Williams breakout signal

# - Real order ON/OFF toggle

# - Trade budget (KRW) selection

# - Buy: breakout (market buy)

# - Sell: next day open (>=09:00:10) market sell

#

# NOTE:

# - This script must run under 32-bit Python for Kiwoom OpenAPI+ (e.g., py -3.10-32).

# - For next-day open sell to work automatically, keep the app running.

​

import sys

import json

import os

from datetime import datetime

from collections import deque

​

from PyQt5.QtCore import QTimer

from PyQt5.QtWidgets import (

    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,

    QPushButton, QLabel, QDoubleSpinBox, QSpinBox, QMessageBox, QComboBox,

    QCheckBox

)

​

import matplotlib

matplotlib.use("Qt5Agg")

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

from matplotlib.figure import Figure

​

from pykiwoom.kiwoom import Kiwoom

​

​

# =========================

# Settings

# =========================

CODE = "122630"  # KODEX 200 Leverage

​

STATE_PATH = "state_lw_gui_trade.json"

​

# Next-day open sell time (KST)

NEXTOPEN_SELL_HHMMSS = (9, 0, 10)

​

# Market time window (used for reset guards)

MARKET_OPEN_HHMMSS = (9, 0, 0)

MARKET_CLOSE_HHMMSS = (15, 30, 0)

​

​

# =========================

# Utils

# =========================

def to_int(x) -> int:

    try:

        return int(str(x).replace(",", "").strip())

    except:

        return 0

​

​

def yyyymmdd(dt=None) -> str:

    dt = dt or datetime.now()

    return dt.strftime("%Y%m%d")

​

​

def now_dt() -> datetime:

    return datetime.now()

​

​

def reached_hms(hms) -> bool:

    t = now_dt().time()

    target = now_dt().replace(hour=hms[0], minute=hms[1], second=hms[2], microsecond=0).time()

    return t >= target

​

​

def in_time_window(start_hms, end_hms) -> bool:

    t = now_dt().time()

    s = now_dt().replace(hour=start_hms[0], minute=start_hms[1], second=start_hms[2], microsecond=0).time()

    e = now_dt().replace(hour=end_hms[0], minute=end_hms[1], second=end_hms[2], microsecond=0).time()

    return s <= t <= e

​

​

def floor_to_10min(dt: datetime) -> datetime:

    m = (dt.minute // 10) * 10

    return dt.replace(minute=m, second=0, microsecond=0)

​

​

def load_state() -> dict:

    if os.path.exists(STATE_PATH):

        try:

            with open(STATE_PATH, "r", encoding="utf-8") as f:

                return json.load(f)

        except:

            return {}

    return {}

​

​

def save_state(state: dict) -> None:

    with open(STATE_PATH, "w", encoding="utf-8") as f:

        json.dump(state, f, ensure_ascii=False, indent=2)

​

​

def norm_code(x: str) -> str:

    s = str(x).strip()

    return s[1:] if s.startswith("A") else s

​

​

def calc_qty_by_budget(price: int, budget_krw: int) -> int:

    if price <= 0:

        return 0

    return max(0, budget_krw // price)

​

​

# =========================

# Chart widget

# =========================

class CandleChart(QWidget):

    """Simple matplotlib candle rendering (wick + thick body line)."""

    def __init__(self, parent=None):

        super().__init__(parent)

        self.fig = Figure(figsize=(10, 4))

        self.ax = self.fig.add_subplot(111)

        self.canvas = FigureCanvas(self.fig)

​

        lay = QVBoxLayout()

        lay.addWidget(self.canvas)

        self.setLayout(lay)

​

    def plot(self, candles, title: str):

        """

        candles: list of dict {t(datetime), o,h,l,c}

        """

        self.ax.clear()

​

        if not candles:

            self.ax.set_title(title + " (No Data)")

            self.ax.grid(True, alpha=0.3)

            self.canvas.draw()

            return

​

        xs = list(range(len(candles)))

        for i, cd in enumerate(candles):

            o, h, l, c = cd["o"], cd["h"], cd["l"], cd["c"]

            # wick

            self.ax.plot([i, i], [l, h], linewidth=1)

            # body (thick line)

            self.ax.plot([i, i], [o, c], linewidth=6)

​

        # X tick labels

        if "Daily" in title:

            labels = [cd["t"].strftime("%m-%d") for cd in candles]

        else:

            labels = [cd["t"].strftime("%H:%M") for cd in candles]

​

        step = max(1, len(labels) // 10)

        self.ax.set_xticks(xs[::step])

        self.ax.set_xticklabels(labels[::step], rotation=0)

​

        self.ax.set_title(title)

        self.ax.set_ylabel("Price")

        self.ax.grid(True, alpha=0.3)

        self.fig.tight_layout()

        self.canvas.draw()

​

​

# =========================

# Main window

# =========================

class MainWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.setWindowTitle("LW Breakout GUI (122630)")

​

        # Kiwoom

        self.kiwoom = Kiwoom()

        self.logged_in = False

        self.account = None

​

        # State / cache

        self.today = yyyymmdd()

        self.last_price = None

        self.breakout_price = None

        self.breakout_ref_date = None

        self.state = load_state()

​

        # Data buffers

        self.ticks = deque(maxlen=5000)  # (dt, price)

        self.candles_10m = []            # today's 10-min candles

        self.current_candle_start = None

        self.candles_daily = []          # last 120 daily candles

​

        # UI

        central = QWidget()

        self.setCentralWidget(central)

        v = QVBoxLayout()

        central.setLayout(v)

​

        self.lbl_login = QLabel("Login: (not connected)")

        self.lbl_price = QLabel("Last Price: -")

        self.lbl_breakout = QLabel("Breakout: -")

        self.lbl_signal = QLabel("Signal: -")

        self.lbl_trade = QLabel("Trade: -")

​

        for lb in [self.lbl_login, self.lbl_price, self.lbl_breakout, self.lbl_signal, self.lbl_trade]:

            lb.setStyleSheet("font-size: 14px;")

​

        v.addWidget(self.lbl_login)

        v.addWidget(self.lbl_price)

        v.addWidget(self.lbl_breakout)

        v.addWidget(self.lbl_signal)

        v.addWidget(self.lbl_trade)

​

        row = QHBoxLayout()

​

        row.addWidget(QLabel("Timeframe:"))

        self.cmb_tf = QComboBox()

        self.cmb_tf.addItems(["10-Min Candles (Today)", "Daily Candles (Last 120 Days)"])

        row.addWidget(self.cmb_tf)

​

        row.addWidget(QLabel("k:"))

        self.spin_k = QDoubleSpinBox()

        self.spin_k.setDecimals(2)

        self.spin_k.setRange(0.10, 1.50)

        self.spin_k.setSingleStep(0.05)

        self.spin_k.setValue(0.60)

        row.addWidget(self.spin_k)

​

        row.addWidget(QLabel("Poll (sec):"))

        self.spin_poll = QSpinBox()

        self.spin_poll.setRange(1, 60)

        self.spin_poll.setValue(5)

        row.addWidget(self.spin_poll)

​

        self.chk_real = QCheckBox("REAL ORDER ON")

        self.chk_real.setChecked(False)

        row.addWidget(self.chk_real)

​

        row.addWidget(QLabel("Budget (KRW):"))

        self.spin_budget = QSpinBox()

        self.spin_budget.setRange(10_000, 200_000_000)

        self.spin_budget.setSingleStep(10_000)

        self.spin_budget.setValue(1_000_000)

        row.addWidget(self.spin_budget)

​

        self.btn_login = QPushButton("Login")

        self.btn_start = QPushButton("Start")

        self.btn_stop = QPushButton("Stop")

        self.btn_reset = QPushButton("Reset (Today)")

​

        row.addWidget(self.btn_login)

        row.addWidget(self.btn_start)

        row.addWidget(self.btn_stop)

        row.addWidget(self.btn_reset)

​

        row.addStretch(1)

        v.addLayout(row)

​

        self.chart = CandleChart()

        v.addWidget(self.chart)

        self.chart.plot([], "10-Min Candles (Today)")

​

        # Timer

        self.timer = QTimer(self)

        self.timer.timeout.connect(self.on_tick)

​

        # Wiring

        self.btn_login.clicked.connect(self.do_login)

        self.btn_start.clicked.connect(self.start)

        self.btn_stop.clicked.connect(self.stop)

        self.btn_reset.clicked.connect(self.reset_day)

​

        self.cmb_tf.currentIndexChanged.connect(self.on_timeframe_changed)

        self.spin_k.valueChanged.connect(self.on_k_changed)

​

        self.btn_start.setEnabled(False)

        self.btn_stop.setEnabled(False)

​

        self._chart_redraw_counter = 0

        self.render_trade_state()

​

    # -------------------------

    # UI handlers

    # -------------------------

    def on_timeframe_changed(self):

        try:

            if not self.logged_in:

                self.refresh_chart_only(force=True)

                return

​

            if "Daily" in self.cmb_tf.currentText():

                self.load_daily_120()

​

            self.refresh_chart_only(force=True)

        except Exception as e:

            self.lbl_trade.setText(f"Trade: timeframe change error {repr(e)}")

​

    def on_k_changed(self):

        # Recompute breakout next tick

        self.breakout_price = None

​

    def refresh_chart_only(self, force: bool = False):

        tf = self.cmb_tf.currentText()

        if "Daily" in tf:

            title = "Daily Candles (Last 120 Days)"

            candles = self.candles_daily

        else:

            title = "10-Min Candles (Today)"

            candles = self.candles_10m

​

        if force:

            self.chart.plot(candles, title)

            return

​

        self._chart_redraw_counter += 1

        if self._chart_redraw_counter % 2 == 0:

            self.chart.plot(candles, title)

​

    # -------------------------

    # Kiwoom helpers

    # -------------------------

    def do_login(self):

        try:

            self.kiwoom.CommConnect(block=True)

            self.logged_in = True

​

            accs_raw = self.kiwoom.GetLoginInfo("ACCNO")

            if isinstance(accs_raw, str):

                accs = [a for a in accs_raw.split(";") if a.strip()]

            else:

                accs = [a for a in accs_raw if str(a).strip()]

            self.account = accs[0] if accs else None

​

            self.lbl_login.setText(f"Login: OK  |  Account: {self.account}")

            self.btn_start.setEnabled(True)

​

            # Load current timeframe once

            self.on_timeframe_changed()

​

        except Exception as e:

            QMessageBox.critical(self, "Login Error", repr(e))

​

    def get_current_price(self) -> int:

        cur = self.kiwoom.block_request(

            "OPT10001",

            종목코드=CODE,

            output="주식기본정보",

            next=0

        )

        if "현재가" not in cur.columns:

            raise RuntimeError(f"Missing '현재가' column: {list(cur.columns)}")

        return abs(to_int(cur.iloc[0]["현재가"]))

​

    def calc_breakout(self, k: float) -> int:

        df = self.kiwoom.block_request(

            "OPT10081",

            종목코드=CODE,

            기준일자=yyyymmdd(),

            수정주가구분="1",

            output="주식일봉차트조회",

            next=0

        )

        if "일자" not in df.columns:

            raise RuntimeError(f"Missing '일자' column: {list(df.columns)}")

​

        df = df.copy()

        df["일자"] = df["일자"].astype(str)

​

        for col in ["시가", "고가", "저가"]:

            if col in df.columns:

                df[col] = df[col].apply(lambda v: abs(to_int(v)))

​

        df = df.sort_values("일자").reset_index(drop=True)

        if len(df) < 2:

            raise RuntimeError("Need at least 2 daily bars")

​

        y = df.iloc[-2]

        t = df.iloc[-1]

​

        self.breakout_ref_date = str(t["일자"])

        return int(round(int(t["시가"]) + k * (int(y["고가"]) - int(y["저가"]))))

​

    def load_daily_120(self):

        df = self.kiwoom.block_request(

            "OPT10081",

            종목코드=CODE,

            기준일자=yyyymmdd(),

            수정주가구분="1",

            output="주식일봉차트조회",

            next=0

        )

        if "일자" not in df.columns:

            raise RuntimeError(f"Missing '일자' column: {list(df.columns)}")

​

        df = df.copy()

        df["일자"] = df["일자"].astype(str)

        for col in ["시가", "고가", "저가", "현재가"]:

            if col in df.columns:

                df[col] = df[col].apply(lambda v: abs(to_int(v)))

​

        df = df.sort_values("일자").reset_index(drop=True).tail(120).reset_index(drop=True)

​

        candles = []

        for _, r in df.iterrows():

            d = datetime.strptime(str(r["일자"]), "%Y%m%d")

            candles.append({

                "t": d,

                "o": int(r["시가"]),

                "h": int(r["고가"]),

                "l": int(r["저가"]),

                "c": int(r["현재가"]),

            })

        self.candles_daily = candles

​

    def get_position_qty(self) -> int:

        df = self.kiwoom.block_request(

            "OPW00018",

            계좌번호=self.account,

            비밀번호="",

            비밀번호입력매체구분="00",

            조회구분="2",

            output="계좌평가잔고개별합산",

            next=0

        )

​

        code_col = None

        qty_col = None

        for c in df.columns:

            if ("종목번호" in c) or ("종목코드" in c):

                code_col = c

            if "보유수량" in c:

                qty_col = c

​

        if code_col is None or qty_col is None:

            return 0

​

        df2 = df.copy()

        df2[code_col] = df2[code_col].apply(norm_code)

        rows = df2[df2[code_col] == CODE]

        if rows.empty:

            return 0

        return abs(to_int(rows.iloc[0][qty_col]))

​

    def send_market_buy(self, qty: int) -> int:

        return self.kiwoom.SendOrder("LW_BUY_GUI", "0101", self.account, 1, CODE, qty, 0, "03", "")

​

    def send_market_sell(self, qty: int) -> int:

        return self.kiwoom.SendOrder("LW_SELL_GUI", "0102", self.account, 2, CODE, qty, 0, "03", "")

​

    # -------------------------

    # 10-min candles builder (today only)

    # -------------------------

    def reset_day(self):

        self.today = yyyymmdd()

        self.ticks.clear()

        self.candles_10m = []

        self.current_candle_start = None

        self.last_price = None

        self.breakout_price = None

        self.breakout_ref_date = None

        self.lbl_signal.setText("Signal: - (reset)")

        self.refresh_chart_only(force=True)

​

    def update_10m_candles(self, t: datetime, price: int) -> bool:

        start = floor_to_10min(t)

        if self.current_candle_start is None:

            self.current_candle_start = start

            self.candles_10m.append({"t": start, "o": price, "h": price, "l": price, "c": price})

            return True

​

        if start == self.current_candle_start:

            cd = self.candles_10m[-1]

            cd["h"] = max(cd["h"], price)

            cd["l"] = min(cd["l"], price)

            cd["c"] = price

            return False

        else:

            self.current_candle_start = start

            self.candles_10m.append({"t": start, "o": price, "h": price, "l": price, "c": price})

            if len(self.candles_10m) > 80:

                self.candles_10m = self.candles_10m[-80:]

            return True

​

    # -------------------------

    # Trade state label

    # -------------------------

    def render_trade_state(self):

        st = self.state or {}

        bought = st.get("bought_date")

        pending_sell = st.get("pending_sell_date")

        sold = st.get("sold_date")

        msg = f"Trade: bought={bought}, pending_sell={pending_sell}, sold={sold}"

        self.lbl_trade.setText(msg)

​

    # -------------------------

    # Start / Stop

    # -------------------------

    def start(self):

        if not self.logged_in:

            QMessageBox.information(self, "Info", "Please login first.")

            return

​

        poll_ms = int(self.spin_poll.value() * 1000)

        self.timer.start(poll_ms)

​

        self.btn_start.setEnabled(False)

        self.btn_stop.setEnabled(True)

        self.btn_login.setEnabled(False)

​

        # If Daily selected, load once on start

        try:

            if "Daily" in self.cmb_tf.currentText():

                self.load_daily_120()

        except Exception as e:

            self.lbl_trade.setText(f"Trade: daily load error {repr(e)}")

​

        self.on_tick()

​

    def stop(self):

        self.timer.stop()

        self.btn_start.setEnabled(True)

        self.btn_stop.setEnabled(False)

        self.btn_login.setEnabled(True)

​

    # -------------------------

    # Trading logic

    # -------------------------

    def try_next_open_sell(self):

        today = yyyymmdd()

        pending = self.state.get("pending_sell_date")

        if not pending or pending != today:

            return

​

        if self.state.get("sold_date") == today:

            return

​

        if not in_time_window(MARKET_OPEN_HHMMSS, MARKET_CLOSE_HHMMSS):

            return

​

        if not reached_hms(NEXTOPEN_SELL_HHMMSS):

            return

​

        qty = self.get_position_qty()

        if qty <= 0:

            self.state["sold_date"] = today

            self.state["sold_reason"] = "no_position"

            save_state(self.state)

            self.render_trade_state()

            return

​

        if not self.chk_real.isChecked():

            self.state["sold_date"] = today

            self.state["sold_reason"] = "paper_sell"

            save_state(self.state)

            self.render_trade_state()

            return

​

        ret = self.send_market_sell(qty)

        self.state["sold_date"] = today

        self.state["sold_reason"] = "sent_sell"

        self.state["sell_ret"] = ret

        save_state(self.state)

        self.render_trade_state()

​

    def try_breakout_buy(self, current_price: int):

        today = yyyymmdd()

​

        if self.state.get("bought_date") == today:

            return

​

        pos = self.get_position_qty()

        if pos > 0:

            self.state["bought_date"] = today

            self.state["bought_reason"] = "already_holding"

            save_state(self.state)

            self.render_trade_state()

            return

​

        k = float(self.spin_k.value())

        if self.breakout_price is None:

            self.breakout_price = self.calc_breakout(k)

​

        if current_price < self.breakout_price:

            return

​

        budget = int(self.spin_budget.value())

        qty = calc_qty_by_budget(current_price, budget)

        if qty <= 0:

            return

​

        # next day sell schedule

        next_day = datetime.fromordinal(now_dt().date().toordinal() + 1)

        pending_sell_date = yyyymmdd(next_day)

​

        if not self.chk_real.isChecked():

            self.state["bought_date"] = today

            self.state["bought_reason"] = "paper_buy"

            self.state["buy_price_snapshot"] = current_price

            self.state["buy_qty"] = qty

            self.state["pending_sell_date"] = pending_sell_date

            save_state(self.state)

            self.render_trade_state()

            return

​

        ret = self.send_market_buy(qty)

        self.state["bought_date"] = today

        self.state["bought_reason"] = "sent_buy"

        self.state["buy_price_snapshot"] = current_price

        self.state["buy_qty"] = qty

        self.state["buy_ret"] = ret

        self.state["pending_sell_date"] = pending_sell_date

        save_state(self.state)

        self.render_trade_state()

​

    # -------------------------

    # Timer tick

    # -------------------------

    def on_tick(self):

        try:

            # daily reset for 10-min candles

            if yyyymmdd() != self.today:

                self.reset_day()

​

            # 1) next-day open sell

            self.try_next_open_sell()

​

            # 2) current price

            price = self.get_current_price()

            self.last_price = price

            self.lbl_price.setText(f"Last Price: {price:,}")

​

            # 3) breakout (cached)

            k = float(self.spin_k.value())

            if self.breakout_price is None:

                self.breakout_price = self.calc_breakout(k)

​

            self.lbl_breakout.setText(

                f"Breakout: {self.breakout_price:,}  (k={k:.2f}, ref={self.breakout_ref_date})"

            )

​

            # 4) signal

            if price >= self.breakout_price:

                self.lbl_signal.setText("Signal: BREAKOUT (BUY condition met)")

            else:

                self.lbl_signal.setText("Signal: WAIT")

​

            # 5) update 10-min candle (today only)

            t = now_dt()

            self.ticks.append((t, price))

            candle_changed = self.update_10m_candles(t, price)

​

            # 6) breakout buy logic

            self.try_breakout_buy(price)

​

            # 7) chart refresh

            tf = self.cmb_tf.currentText()

            if "Daily" in tf:

                # daily chart is static unless user reloads; just redraw lightly

                self.refresh_chart_only(force=False)

            else:

                if candle_changed:

                    self.refresh_chart_only(force=True)

                else:

                    self.refresh_chart_only(force=False)

​

        except Exception as e:

            self.lbl_signal.setText(f"Error: {repr(e)}")

​

​

def main():

    app = QApplication(sys.argv)

    w = MainWindow()

    w.resize(1200, 760)

    w.show()

    sys.exit(app.exec_())

​

​

if __name__ == "__main__":

    main()
