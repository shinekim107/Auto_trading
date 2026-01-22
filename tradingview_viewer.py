import tkinter as tk
from tkinter import ttk, messagebox
import multiprocessing as mp
import atexit


# -----------------------
# TradingView helpers
# -----------------------
def normalize_symbol(s: str) -> str:
    s = (s or "").strip().upper()
    if not s:
        return ""
    if ":" in s:
        return s
    if s.isdigit():
        return f"KRX:{s}"
    return f"NASDAQ:{s}"


def tv_chart_url(symbol: str) -> str:
    return f"https://www.tradingview.com/chart/?symbol={symbol}"


# -----------------------
# WebView process (runs pywebview in a separate process)
# -----------------------
def webview_process_main(url_queue: mp.Queue, title_queue: mp.Queue):
    import webview
    import threading
    import queue as pyqueue

    # Wait for initial URL
    init_url = url_queue.get()
    init_title = title_queue.get() if not title_queue.empty() else "TradingView"

    window = webview.create_window(
        title=init_title,
        url=init_url,
        width=1200,
        height=800,
        resizable=True
    )

    def worker():
        # This runs after the GUI is initialized (safe place to control window)
        while True:
            try:
                url = url_queue.get(timeout=0.2)
                title = None
                try:
                    title = title_queue.get_nowait()
                except pyqueue.Empty:
                    pass

                if url is None:  # shutdown signal
                    break

                if title:
                    try:
                        window.set_title(title)
                    except Exception:
                        pass

                try:
                    window.load_url(url)
                except Exception:
                    # if load_url fails, ignore (window may be closing)
                    pass

            except pyqueue.Empty:
                continue
            except Exception:
                continue

    webview.start(worker, debug=False)


class TVViewerController:
    def __init__(self):
        self.proc = None
        self.url_q = None
        self.title_q = None

    def ensure_started(self, first_url: str, first_title: str):
        if self.proc is not None and self.proc.is_alive():
            return

        self.url_q = mp.Queue()
        self.title_q = mp.Queue()
        self.proc = mp.Process(
            target=webview_process_main,
            args=(self.url_q, self.title_q),
            daemon=True
        )
        self.proc.start()

        # Send initial payload
        self.url_q.put(first_url)
        self.title_q.put(first_title)

    def open_or_update(self, symbol: str):
        url = tv_chart_url(symbol)
        title = f"TradingView - {symbol}"

        self.ensure_started(url, title)

        # If already started, push updates
        if self.proc is not None and self.proc.is_alive():
            self.url_q.put(url)
            self.title_q.put(title)

    def shutdown(self):
        try:
            if self.proc is not None and self.proc.is_alive():
                self.url_q.put(None)
                self.proc.terminate()
        except Exception:
            pass


# -----------------------
# Tkinter UI
# -----------------------
def main():
    mp.freeze_support()  # for Windows exe packaging / safety

    controller = TVViewerController()
    atexit.register(controller.shutdown)

    root = tk.Tk()
    root.title("TradingView Viewer (Tkinter)")
    root.geometry("520x360")

    # Top input row
    frm_top = ttk.Frame(root, padding=10)
    frm_top.pack(fill="x")

    ttk.Label(frm_top, text="Symbol").pack(side="left")

    symbol_var = tk.StringVar(value="AMEX:TQQQ")
    ent = ttk.Entry(frm_top, textvariable=symbol_var, width=30)
    ent.pack(side="left", padx=8)

    def on_open():
        raw = symbol_var.get()
        symbol = normalize_symbol(raw)
        if not symbol:
            messagebox.showwarning("입력 필요", "예: AMEX:TQQQ / NASDAQ:AAPL / KRX:005930 처럼 입력해줘.")
            return
        controller.open_or_update(symbol)

    def on_enter(event=None):
        on_open()

    ent.bind("<Return>", on_enter)

    btn_open = ttk.Button(frm_top, text="Open / Update", command=on_open)
    btn_open.pack(side="left")

    # Hint
    frm_hint = ttk.Frame(root, padding=(10, 0))
    frm_hint.pack(fill="x")
    ttk.Label(
        frm_hint,
        text="예) AMEX:TQQQ, NASDAQ:AAPL, KRX:005930  (콜론 없이 입력하면 자동 추정)",
        foreground="#555555"
    ).pack(anchor="w")

    # Favorites
    favorites = ["AMEX:TQQQ", "AMEX:SOXL", "NASDAQ:QQQ", "KRX:005930"]

    frm_fav = ttk.LabelFrame(root, text="Favorites", padding=10)
    frm_fav.pack(fill="both", expand=True, padx=10, pady=10)

    fav_listbox = tk.Listbox(frm_fav, height=8)
    fav_listbox.pack(side="left", fill="both", expand=True)

    for s in favorites:
        fav_listbox.insert(tk.END, s)

    scroll = ttk.Scrollbar(frm_fav, orient="vertical", command=fav_listbox.yview)
    scroll.pack(side="right", fill="y")
    fav_listbox.config(yscrollcommand=scroll.set)

    def add_fav():
        raw = symbol_var.get()
        symbol = normalize_symbol(raw)
        if not symbol:
            return
        # avoid duplicates
        current = set(fav_listbox.get(0, tk.END))
        if symbol not in current:
            fav_listbox.insert(tk.END, symbol)

    def on_fav_double_click(event=None):
        sel = fav_listbox.curselection()
        if not sel:
            return
        symbol = fav_listbox.get(sel[0])
        symbol_var.set(symbol)
        controller.open_or_update(symbol)

    fav_listbox.bind("<Double-Button-1>", on_fav_double_click)

    # Bottom buttons
    frm_bottom = ttk.Frame(root, padding=(10, 0, 10, 10))
    frm_bottom.pack(fill="x")

    ttk.Button(frm_bottom, text="Add to Favorites", command=add_fav).pack(side="left")
    ttk.Label(frm_bottom, text="* 즐겨찾기 더블클릭 = 차트 열기/업데이트", foreground="#555").pack(side="right")

    ent.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
