import tkinter as tk
from tkinter import ttk, messagebox
import math


# =====================
# Formatting / helpers
# =====================
def fnum(x, nd=4):
    try:
        return f"{float(x):,.{nd}f}"
    except Exception:
        return str(x)


def clamp(v, lo=0.0, hi=None):
    v = float(v)
    if hi is None:
        return max(lo, v)
    return max(lo, min(hi, v))


def floor_int(x: float) -> int:
    return int(math.floor(x + 1e-12))


def round_int_half_up(x: float) -> int:
    return int(math.floor(x + 0.5 + 1e-12))


def is_positive_number(s: str) -> bool:
    try:
        return float(s.strip()) > 0
    except Exception:
        return False


def round_price_01(p: float) -> float:
    # US ETF tick size assumed 0.01
    return round(float(p) + 1e-12, 2)


# ==========================================
# IB Tab (Infinite Buy) Calculator - Frame
# ==========================================
class LaoIBCalculatorFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=0)

        # 기본 분할수
        self.default_splits_map = {
            "V2.2": 40,
            "V2.1": 40,
            "V2.1후반": 40,
            "V2.0": 40,
            "V3.0": 20,
            "IBS": 10,
            "쿼터손절": 10,
        }
        self.default_splits_fallback = 40

        # 큰수매수 = 평단 * (1 + K * 매도비율%)
        self.big_buy_k = 0.0052826

        self._build_ui()
        self._bind_events()

        self._apply_auto_defaults()
        self._update_field_states()
        self._refresh_auto_unit_display()
        self._refresh_auto_progress_display()
        self._update_strategy_desc("")

    # ---------- UI ----------
    def _build_ui(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="x")

        # Row 0
        ttk.Label(frm, text="매매방법").grid(row=0, column=0, sticky="w")
        self.strategy = ttk.Combobox(
            frm,
            values=["V2.2", "IBS", "V3.0", "쿼터손절", "V2.1", "V2.1후반", "V2.0", "V1"],
            state="readonly",
            width=12,
        )
        self.strategy.set("V2.2")
        self.strategy.grid(row=0, column=1, padx=6, pady=6, sticky="w")

        ttk.Label(frm, text="종목").grid(row=0, column=2, sticky="w")
        self.symbol = ttk.Combobox(frm, values=["TQQQ", "SOXL", "QLD", "기타"], state="readonly", width=10)
        self.symbol.set("SOXL")
        self.symbol.grid(row=0, column=3, padx=6, pady=6, sticky="w")

        ttk.Label(frm, text="평단가").grid(row=0, column=4, sticky="w")
        self.avg_price_var = tk.StringVar(value="51.4016")
        self.avg_entry = ttk.Entry(frm, textvariable=self.avg_price_var, width=12)
        self.avg_entry.grid(row=0, column=5, padx=6, pady=6, sticky="w")

        ttk.Label(frm, text="현재가").grid(row=0, column=6, sticky="w")
        self.cur_price_var = tk.StringVar(value="53.95")
        self.cur_entry = ttk.Entry(frm, textvariable=self.cur_price_var, width=12)
        self.cur_entry.grid(row=0, column=7, padx=6, pady=6, sticky="w")

        # 전략 요약 (매매방법 아래)
        self.strategy_desc_var = tk.StringVar(value="")
        self.strategy_desc_label = ttk.Label(
            frm,
            textvariable=self.strategy_desc_var,
            font=("Arial", 9),
            foreground="#333",
            wraplength=1120,
            justify="left",
        )
        self.strategy_desc_label.grid(row=1, column=0, columnspan=8, sticky="w", pady=(0, 8))

        # Row 2
        ttk.Label(frm, text="원금(총 투자금)").grid(row=2, column=0, sticky="w")
        self.principal_var = tk.StringVar(value="10000")
        self.principal_entry = ttk.Entry(frm, textvariable=self.principal_var, width=12)
        self.principal_entry.grid(row=2, column=1, padx=6, pady=6, sticky="w")

        ttk.Label(frm, text="분할수 (기본 자동)").grid(row=2, column=2, sticky="w")
        self.splits_var = tk.StringVar(value="")
        self.splits_entry = ttk.Entry(frm, textvariable=self.splits_var, width=12)
        self.splits_entry.grid(row=2, column=3, padx=6, pady=6, sticky="w")

        ttk.Label(frm, text="1회치 금액 (입력 시 최우선)").grid(row=2, column=4, sticky="w")
        self.unit_cash_var = tk.StringVar(value="")
        self.unit_entry = ttk.Entry(frm, textvariable=self.unit_cash_var, width=12)
        self.unit_entry.grid(row=2, column=5, padx=6, pady=6, sticky="w")

        ttk.Label(frm, text="보유수량").grid(row=2, column=6, sticky="w")
        self.hold_qty_var = tk.StringVar(value="6")
        self.hold_entry = ttk.Entry(frm, textvariable=self.hold_qty_var, width=12)
        self.hold_entry.grid(row=2, column=7, padx=6, pady=6, sticky="w")

        # Row 3 (손익률%, 진행률 입력 제거 유지)
        ttk.Label(frm, text="T(IBS)").grid(row=3, column=4, sticky="w")
        self.T_var = tk.StringVar(value="1")
        self.T_entry = ttk.Entry(frm, textvariable=self.T_var, width=12)
        self.T_entry.grid(row=3, column=5, padx=6, pady=6, sticky="w")

        ttk.Label(frm, text="매도비율% (참고값)").grid(row=3, column=6, sticky="w")
        self.sell_pct_var = tk.StringVar(value="13")
        self.sell_pct_entry = ttk.Entry(frm, textvariable=self.sell_pct_var, width=12)
        self.sell_pct_entry.grid(row=3, column=7, padx=6, pady=6, sticky="w")

        ttk.Button(frm, text="계산", command=self.on_calculate).grid(row=4, column=6, padx=6, pady=6, sticky="e")
        ttk.Button(frm, text="초기화", command=self.on_reset).grid(row=4, column=7, padx=6, pady=6, sticky="w")

        # AUTO info
        self.auto_unit_label_var = tk.StringVar(value="")
        self.used_unit_label_var = tk.StringVar(value="")
        self.auto_progress_label_var = tk.StringVar(value="")

        ttk.Label(frm, textvariable=self.auto_unit_label_var, foreground="#444").grid(
            row=4, column=0, columnspan=3, sticky="w", padx=(0, 6), pady=(2, 0)
        )
        ttk.Label(frm, textvariable=self.used_unit_label_var, foreground="#444").grid(
            row=4, column=3, columnspan=3, sticky="w", padx=(0, 6), pady=(2, 0)
        )
        ttk.Label(frm, textvariable=self.auto_progress_label_var, foreground="#444").grid(
            row=5, column=0, columnspan=8, sticky="w", padx=(0, 6), pady=(2, 0)
        )

        # Output
        out = ttk.Frame(self, padding=(12, 0, 12, 12))
        out.pack(fill="both", expand=True)

        ttk.Label(out, text="계산 결과", font=("Arial", 12, "bold")).pack(anchor="w", pady=(8, 6))

        table_frame = ttk.Frame(out)
        table_frame.pack(fill="both", expand=True)

        cols = ("구분", "세션", "태그", "주문가", "수량", "금액")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=18)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, anchor="center")
        self.tree.column("태그", width=360, anchor="w")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.note_var = tk.StringVar(
            value="규칙 계산 전용(주문 전송 없음). 수량은 정수이며, 0주도 '(금액 부족)'으로 표시합니다."
        )
        ttk.Label(out, textvariable=self.note_var, foreground="#555").pack(anchor="w", pady=(6, 0))

    def _bind_events(self):
        self.strategy.bind("<<ComboboxSelected>>", lambda e: self._on_strategy_changed())

        self.principal_var.trace_add("write", lambda *_: self._refresh_auto_unit_display())
        self.splits_var.trace_add("write", lambda *_: self._refresh_auto_unit_display())
        self.unit_cash_var.trace_add("write", lambda *_: self._refresh_auto_unit_display())

        self.avg_price_var.trace_add("write", lambda *_: self._refresh_auto_progress_display())
        self.hold_qty_var.trace_add("write", lambda *_: self._refresh_auto_progress_display())
        self.principal_var.trace_add("write", lambda *_: self._refresh_auto_progress_display())

    def _on_strategy_changed(self):
        self._apply_auto_defaults()
        self._update_field_states()
        self._refresh_auto_unit_display()
        self._refresh_auto_progress_display()
        self._update_strategy_desc("")

    def on_reset(self):
        self.strategy.set("V2.2")
        self.symbol.set("SOXL")
        self.avg_price_var.set("51.4016")
        self.cur_price_var.set("53.95")
        self.principal_var.set("10000")
        self.hold_qty_var.set("6")
        self.T_var.set("1")
        self.sell_pct_var.set("13")
        self.splits_var.set("")
        self.unit_cash_var.set("")

        self._apply_auto_defaults()
        self._clear_table()
        self._update_field_states()
        self._refresh_auto_unit_display()
        self._refresh_auto_progress_display()
        self._update_strategy_desc("")

    def _clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _set_state(self, widget, enabled: bool):
        widget.configure(state=("normal" if enabled else "disabled"))

    def _strategy_default_splits(self, strat: str) -> int:
        return int(self.default_splits_map.get(strat, self.default_splits_fallback))

    def _apply_auto_defaults(self):
        strat = self.strategy.get().strip()
        auto_splits = self._strategy_default_splits(strat)
        if not is_positive_number(self.splits_var.get()):
            self.splits_var.set(str(auto_splits))

    def _update_field_states(self):
        strat = self.strategy.get().strip()

        # always
        self._set_state(self.avg_entry, True)
        self._set_state(self.hold_entry, True)

        # defaults
        self._set_state(self.cur_entry, strat in ("V2.2", "IBS", "V3.0", "V2.1", "V2.1후반", "쿼터손절"))

        buy_enabled = strat in ("V2.2", "IBS", "V3.0", "V2.1", "V2.1후반", "쿼터손절")
        self._set_state(self.principal_entry, buy_enabled)
        self._set_state(self.splits_entry, buy_enabled)
        self._set_state(self.unit_entry, buy_enabled)

        self._set_state(self.T_entry, strat == "IBS")
        self._set_state(self.sell_pct_entry, strat == "V2.2")

    def _resolve_numeric(self, s: str) -> float:
        try:
            return float(s.strip())
        except Exception:
            return 0.0

    def _refresh_auto_unit_display(self):
        strat = self.strategy.get().strip()
        auto_splits_default = self._strategy_default_splits(strat)

        principal = self._resolve_numeric(self.principal_var.get())
        splits_in = self.splits_var.get().strip()
        splits = int(float(splits_in)) if is_positive_number(splits_in) else auto_splits_default

        auto_unit = 0.0
        if principal > 0 and splits > 0:
            auto_unit = principal / float(splits)

        unit_in = self.unit_cash_var.get().strip()
        if is_positive_number(unit_in):
            used_unit = float(unit_in)
            used_src = "MANUAL(1회치)"
        else:
            used_unit = auto_unit
            used_src = "AUTO(원금/분할)"

        self.auto_unit_label_var.set(
            f"자동 분할수 기본값:{auto_splits_default} | 현재 분할수:{splits} | 자동 1회치(AUTO)=원금/분할={fnum(auto_unit,2)}"
        )

        if strat in ("V1", "V2.0"):
            self.used_unit_label_var.set("USED 1회치: (이 전략은 매수 계산 없음)")
        else:
            self.used_unit_label_var.set(f"실제 사용 1회치(USED): {fnum(used_unit,2)} [{used_src}]")

    def _refresh_auto_progress_display(self):
        principal = self._resolve_numeric(self.principal_var.get())
        avg = self._resolve_numeric(self.avg_price_var.get())
        try:
            hold_qty = int(float(self.hold_qty_var.get().strip() or "0"))
        except Exception:
            hold_qty = 0

        invested = max(0.0, float(hold_qty) * float(avg))
        if principal > 0:
            auto_progress = clamp(invested / principal, 0.0, 10.0) * 100.0
            self.auto_progress_label_var.set(
                f"자동 진행률(AUTO)= (보유×평단)/원금 = ({hold_qty}×{fnum(avg,4)})/{fnum(principal,2)} = {fnum(auto_progress,2)}%"
            )
        else:
            self.auto_progress_label_var.set("자동 진행률(AUTO): 원금이 0이어서 계산 불가")

    def _qty_from_cash_int(self, cash, price) -> int:
        if price <= 0:
            return 0
        return floor_int(cash / price)

    def _add_row(self, side, session, tag, price, qty_int: int):
        amount = (price * qty_int) if (price is not None and qty_int is not None) else 0.0
        if qty_int is None:
            qty_int = 0
        self.tree.insert(
            "",
            "end",
            values=(
                side,
                session,
                tag,
                "-" if price is None else fnum(price, 4),
                f"{int(qty_int):,d}",
                fnum(amount, 2),
            ),
        )

    def _update_strategy_desc(self, text: str):
        self.strategy_desc_var.set(text or "")

    def on_calculate(self):
        try:
            strategy = self.strategy.get().strip()
            symbol = self.symbol.get().strip().upper()

            # ✅ 지금까지 맞춘 로직은 V2.2(참고프로그램 맞춤)만 적용
            if strategy != "V2.2":
                raise ValueError("현재 IB 탭은 V2.2(참고프로그램 맞춤)만 계산합니다.")

            avg = float(self.avg_price_var.get().strip())
            hold_qty = int(float(self.hold_qty_var.get().strip()))
            if avg <= 0:
                raise ValueError("평단가는 0보다 커야 합니다.")
            if hold_qty < 0:
                raise ValueError("보유수량은 0 이상이어야 합니다.")

            cur = float(self.cur_price_var.get().strip() or "0")
            if cur <= 0:
                raise ValueError("V2.2 계산에는 '현재가'가 필요합니다.")

            principal = float(self.principal_var.get().strip() or "0")
            splits_in = self.splits_var.get().strip()
            unit_in = self.unit_cash_var.get().strip()

            sell_pct = float(self.sell_pct_var.get().strip() or "0")
            sell_pct = max(0.0, sell_pct)
            if sell_pct <= 0:
                raise ValueError("V2.2에서는 '매도비율%'이 필요합니다. (예: 13)")

            self._clear_table()

            auto_splits = self._strategy_default_splits(strategy)
            splits = int(float(splits_in)) if is_positive_number(splits_in) else auto_splits

            if is_positive_number(unit_in):
                unit_cash = float(unit_in)
                unit_source = "MANUAL(1회치)"
            else:
                if principal <= 0:
                    raise ValueError("원금(총 투자금)이 필요합니다. (자동 1회치 계산용)")
                if splits <= 0:
                    raise ValueError("분할수는 1 이상이어야 합니다.")
                unit_cash = principal / float(splits)
                unit_source = f"AUTO(원금/분할={splits})"

            # AUTO progress (표시 및 전/후반 판단용)
            invested = max(0.0, float(hold_qty) * float(avg))
            progress = (invested / principal) if principal > 0 else 0.0
            front = progress < 0.5

            total_buy_amount = 0.0
            total_sell_amount = 0.0

            def add(side, session, tag, price, qty_int):
                nonlocal total_buy_amount, total_sell_amount
                qty_int = int(qty_int)

                if qty_int <= 0:
                    self._add_row(side, session, tag + " (금액 부족)", price, 0)
                    return

                amt = (price * qty_int) if price is not None else 0.0
                if side == "BUY":
                    total_buy_amount += amt
                else:
                    total_sell_amount += amt
                self._add_row(side, session, tag, price, qty_int)

            # -------------------
            # V2.2 (reference fit)
            # -------------------
            big_buy_price = round_price_01(avg * (1.0 + self.big_buy_k * sell_pct))

            cap = cur * 1.15
            p1 = round_price_01(min(avg, cap))
            p2 = round_price_01(min(big_buy_price, cap))

            if front:
                q1 = self._qty_from_cash_int(unit_cash * 0.5, p1)
                add("BUY", "LOC", "평단LOC 매수 0.5회치", p1, q1)

                q2 = self._qty_from_cash_int(unit_cash * 0.5, p2)
                add("BUY", "LOC", "큰수LOC 매수 0.5회치", p2, q2)
            else:
                q2 = self._qty_from_cash_int(unit_cash * 1.0, p2)
                add("BUY", "LOC", "큰수LOC 매수 1회치", p2, q2)

            q_main = int(clamp(round_int_half_up(hold_qty * 0.75), 0, hold_qty))
            q_split = int(max(0, hold_qty - q_main))

            sell_main_price = round_price_01(avg * (1.0 + sell_pct / 100.0))
            sell_second_price = round_price_01(p2 + 0.01)  # 큰수매수 + 0.01

            add("SELL", "AFTER", f"매도(메인) {q_main}주 (≈75%) [평단+{fnum(sell_pct,2)}%]", sell_main_price, q_main)
            add("SELL", "LOC", f"분할 매도2 잔량 {q_split}주 (큰수+0.01)", sell_second_price, q_split)

            parts = [
                f"전략:{strategy}",
                f"종목:{symbol}",
                f"평단:{fnum(avg,4)}",
                f"보유:{hold_qty:,d}주",
                f"현재가:{fnum(cur,4)}",
                f"진행률:{fnum(progress*100,2)}% [AUTO(보유×평단/원금)]",
                f"분할수:{splits} (기본:{auto_splits})",
                f"1회치:{fnum(unit_cash,2)} [{unit_source}]",
                f"매도비율%:{fnum(sell_pct,2)}%",
                f"매수합:{fnum(total_buy_amount,2)}",
                f"매도합:{fnum(total_sell_amount,2)}",
            ]
            self._update_strategy_desc(" | ".join(parts))

            self._refresh_auto_unit_display()
            self._refresh_auto_progress_display()

        except Exception as e:
            messagebox.showerror("에러", str(e))


# ==========================
# Main App with Tabs (IB/VR)
# ==========================
class LaoMultiCalculatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("라오어 매매규칙 계산기 (주문 없음)")
        self.geometry("1180x820")

        self._build_tabs()

    def _build_tabs(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        # IB tab (기존 프로그램)
        tab_ib = ttk.Frame(nb)
        nb.add(tab_ib, text="IB")
        ib_ui = LaoIBCalculatorFrame(tab_ib)
        ib_ui.pack(fill="both", expand=True)

        # VR tab (공란)
        tab_vr = ttk.Frame(nb)
        nb.add(tab_vr, text="VR")

        placeholder = ttk.Frame(tab_vr, padding=24)
        placeholder.pack(fill="both", expand=True)
        ttk.Label(
            placeholder,
            text="VR(밸류 리밸런싱) 계산기는 여기서 구현합니다.\n(현재는 공란)",
            font=("Arial", 12),
            foreground="#444",
            justify="center",
        ).pack(expand=True)


if __name__ == "__main__":
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = LaoMultiCalculatorApp()
    app.mainloop()
