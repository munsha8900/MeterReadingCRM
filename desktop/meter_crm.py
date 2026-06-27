#!/usr/bin/env python3
"""
Meter CRM - a friendly desktop manager for home solar meter readings.
Pure-stdlib UI (Tkinter). Excel import/export is optional (uses openpyxl if present).
Data is stored in meter_data.json next to this program.
"""
import os, sys, json, csv, datetime, calendar
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

APP_NAME = "Home Solar Monitoring"
APP_VER = "1.0"

# ---------- palette ----------
SIDEBAR_BG   = "#13233A"
SIDEBAR_FG   = "#AEB9CC"
SIDEBAR_HOV  = "#1C3252"
SIDEBAR_ACT  = "#1E60D6"
APP_BG       = "#EEF1F6"
CARD_BG      = "#FFFFFF"
LINE         = "#DCE2EC"
TEXT         = "#1B2430"
MUTED        = "#6B7686"
ACCENT       = "#1E60D6"
CH_COLORS    = ["#E8663B", "#2D7DD2", "#2BB673"]   # ch1, ch2, ch3
GOOD         = "#1F9D55"
WARN         = "#C2410C"

FONT       = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_H1    = ("Segoe UI", 18, "bold")
FONT_H2    = ("Segoe UI", 13, "bold")
FONT_BIG   = ("Segoe UI", 22, "bold")
FONT_SMALL = ("Segoe UI", 9)

try:
    import openpyxl
    HAVE_XLSX = True
except Exception:
    HAVE_XLSX = False


# ---------- data layer ----------
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

DATA_FILE = os.path.join(app_dir(), "meter_data.json")


class Store:
    def __init__(self):
        self.data = {"meta": {"version": 1}, "months": []}
        self.load()

    def load(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
            except Exception as e:
                messagebox.showwarning(APP_NAME, f"Could not read data file:\n{e}\nStarting empty.")
                self.data = {"meta": {"version": 1}, "months": []}
        self.data.setdefault("meta", {"version": 1})
        self.data.setdefault("months", [])
        self.data.setdefault("summary", [])

    def save(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2)
            return True
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Could not save:\n{e}")
            return False

    @property
    def months(self):
        return self.data["months"]

    def month(self, name):
        for m in self.months:
            if m["name"] == name:
                return m
        return None

    def add_month(self, name):
        if self.month(name):
            return None
        m = {"name": name, "meter_ref": "", "channels": ["Peak", "Off-peak", "Export"],
             "rates": {}, "readings": []}
        self.months.append(m)
        self.save()
        return m

    def delete_month(self, name):
        self.data["months"] = [m for m in self.months if m["name"] != name]
        self.save()


# ---------- calculations ----------
def compute_units(readings):
    """Return list of rows with computed daily units per channel (today - yesterday)."""
    prev = {"r1": None, "r2": None, "r3": None}
    out = []
    for row in readings:
        u = {}
        for k in ("r1", "r2", "r3"):
            cur, pv = row.get(k), prev[k]
            if isinstance(cur, (int, float)) and isinstance(pv, (int, float)) and cur >= pv:
                u[k] = cur - pv
            else:
                u[k] = None
            if isinstance(cur, (int, float)):
                prev[k] = cur
        out.append({**row, "u1": u["r1"], "u2": u["r2"], "u3": u["r3"]})
    return out


def month_totals(month):
    rows = compute_units(month.get("readings", []))
    t = [0.0, 0.0, 0.0]
    for r in rows:
        for i, k in enumerate(("u1", "u2", "u3")):
            if isinstance(r[k], (int, float)):
                t[i] += r[k]
    return t  # [ch1, ch2, ch3]


def est_savings(month):
    """Rough estimate: export*buyback + offpeak*offpeak_rate - peak*peak_rate.
    Channel meaning varies per month; we map by channel label keywords."""
    rates = month.get("rates", {})
    labels = [c.lower() for c in month.get("channels", ["", "", ""])]
    totals = month_totals(month)
    val = 0.0
    for lab, tot in zip(labels, totals):
        if "export" in lab:
            val += tot * rates.get("buyback", 0)
        elif "off" in lab:
            val += tot * rates.get("offpeak", 0)
        elif "peak" in lab or "import" in lab:
            val -= tot * rates.get("peak", 0)
    return val


MONTH_ABBR = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def quarter_label(ym):
    y, mo = ym.split("-")
    return f"{y}-Q{(int(mo) - 1) // 3 + 1}"


def group_quarters(summary):
    """Group monthly summary rows into calendar quarters (ordered)."""
    out = {}
    order = []
    for s in summary:
        q = quarter_label(s["month"])
        if q not in out:
            out[q] = []
            order.append(q)
        out[q].append(s)
    return [(q, out[q]) for q in order]


def rate_for_month(store, ym):
    """Find the rates dict for a 'YYYY-MM' month by matching a stored month name."""
    y, mo = ym.split("-")
    name = f"{MONTH_ABBR[int(mo)]}{y[2:]}"
    m = store.month(name)
    return m.get("rates", {}) if m else {}


def settlement_charge(peak, offpeak, export, rates):
    """Grid settlement = (Export - Off-peak) x BuyBack  -  Peak x PeakRate.
    Returns (value, credit, peak_cost) or (None, None, None) if rates missing."""
    bb = rates.get("buyback")
    pr = rates.get("peak")
    if bb is None or pr is None:
        return None, None, None
    credit = (export - offpeak) * bb
    peak_cost = peak * pr
    return credit - peak_cost, credit, peak_cost


def actual_saving(store, ym):
    for s in store.data.get("summary", []):
        if s["month"] == ym:
            return s.get("saving")
    return None


def name_to_ym(name):
    """'Jun26' -> '2026-06'. Returns None if not parseable."""
    name = name.strip()
    for i, ab in enumerate(MONTH_ABBR):
        if ab and name.lower().startswith(ab.lower()):
            yy = "".join(ch for ch in name[len(ab):] if ch.isdigit())
            if len(yy) == 2:
                return f"20{yy}-{i:02d}"
    return None


# ---------- UI ----------
class App(tk.Tk):
    def __init__(self, store):
        super().__init__()
        self.store = store
        self.title(f"{APP_NAME}  —  Home Solar Meter Readings")
        self.geometry("1140x740")
        self.minsize(980, 640)
        self.configure(bg=APP_BG)
        self._set_icon()
        self._style()
        self._menu()

        self.current_month = self.months_names()[-1] if self.months_names() else None
        self.active_view = None

        # layout: sidebar + content
        self.sidebar = tk.Frame(self, bg=SIDEBAR_BG, width=210)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self.content = tk.Frame(self, bg=APP_BG)
        self.content.pack(side="left", fill="both", expand=True)

        self._build_sidebar()
        self.show_view("Dashboard")

    def _set_icon(self):
        try:
            base = getattr(sys, "_MEIPASS", app_dir())
            ico = os.path.join(base, "app.ico")
            if os.path.exists(ico):
                self.iconbitmap(ico)
        except Exception:
            pass

    # ----- styling -----
    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("TCombobox", fieldbackground="white", background="white")
        st.configure("Treeview", background="white", fieldbackground="white",
                     foreground=TEXT, rowheight=26, font=FONT, borderwidth=0)
        st.configure("Treeview.Heading", font=FONT_BOLD, background="#E7ECF3",
                     foreground=TEXT, relief="flat")
        st.map("Treeview.Heading", background=[("active", "#DCE3EE")])
        st.map("Treeview", background=[("selected", "#D6E4FF")],
               foreground=[("selected", TEXT)])
        st.configure("Accent.TButton", font=FONT_BOLD, foreground="white",
                     background=ACCENT, borderwidth=0, padding=(14, 7))
        st.map("Accent.TButton", background=[("active", "#174fb0")])
        st.configure("Ghost.TButton", font=FONT, foreground=TEXT,
                     background="#E7ECF3", borderwidth=0, padding=(12, 6))
        st.map("Ghost.TButton", background=[("active", "#D7DEE9")])
        st.configure("Danger.TButton", font=FONT, foreground="white",
                     background="#C0392B", borderwidth=0, padding=(12, 6))
        st.map("Danger.TButton", background=[("active", "#9c2d22")])

    def _menu(self):
        m = tk.Menu(self)
        fm = tk.Menu(m, tearoff=0)
        fm.add_command(label="Import from Excel…", command=self.import_excel)
        fm.add_separator()
        fm.add_command(label="Export current month → CSV…", command=lambda: self.export_csv(False))
        fm.add_command(label="Export ALL months → CSV…", command=lambda: self.export_csv(True))
        fm.add_command(label="Export ALL → Excel…", command=self.export_excel)
        fm.add_separator()
        fm.add_command(label="Save", command=lambda: (self.store.save(), self._flash_saved()))
        fm.add_command(label="Exit", command=self.destroy)
        m.add_cascade(label="File", menu=fm)
        hm = tk.Menu(m, tearoff=0)
        hm.add_command(label="How to use", command=lambda: self.show_view("Help"))
        hm.add_command(label="About", command=self.about)
        m.add_cascade(label="Help", menu=hm)
        self.config(menu=m)

    # ----- helpers -----
    def months_names(self):
        return [m["name"] for m in self.store.months]

    def _flash_saved(self):
        self.title(f"{APP_NAME}  —  Saved ✓")
        self.after(1200, lambda: self.title(f"{APP_NAME}  —  Home Solar Meter Readings"))

    # ----- sidebar -----
    def _build_sidebar(self):
        head = tk.Frame(self.sidebar, bg=SIDEBAR_BG)
        head.pack(fill="x", pady=(22, 18), padx=18)
        tk.Label(head, text="◧  Home Solar", bg=SIDEBAR_BG, fg="white",
                 font=("Segoe UI", 15, "bold")).pack(anchor="w")
        tk.Label(head, text="Monitoring · by MarkFold Digital", bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                 font=FONT_SMALL).pack(anchor="w", pady=(2, 0))

        self.nav_buttons = {}
        for name, icon in [("Dashboard", "▦"), ("Readings", "✎"),
                           ("Quarterly", "▤"),
                           ("Months & Rates", "⚙"), ("Help", "?")]:
            b = tk.Label(self.sidebar, text=f"   {icon}   {name}", bg=SIDEBAR_BG,
                         fg=SIDEBAR_FG, font=FONT, anchor="w", padx=10, pady=11,
                         cursor="hand2")
            b.pack(fill="x", padx=10, pady=2)
            b.bind("<Button-1>", lambda e, n=name: self.show_view(n))
            b.bind("<Enter>", lambda e, w=b: self._nav_hover(w, True))
            b.bind("<Leave>", lambda e, w=b: self._nav_hover(w, False))
            self.nav_buttons[name] = b

        foot = tk.Label(self.sidebar, text=f"v{APP_VER}", bg=SIDEBAR_BG,
                        fg="#4A5C78", font=FONT_SMALL)
        foot.pack(side="bottom", pady=14)

    def _nav_hover(self, w, on):
        if w._is_active if hasattr(w, "_is_active") else False:
            return
        w.configure(bg=SIDEBAR_HOV if on else SIDEBAR_BG)

    def _set_active_nav(self, name):
        for n, w in self.nav_buttons.items():
            active = (n == name)
            w._is_active = active
            w.configure(bg=SIDEBAR_ACT if active else SIDEBAR_BG,
                        fg="white" if active else SIDEBAR_FG,
                        font=FONT_BOLD if active else FONT)

    # ----- view switching -----
    def show_view(self, name):
        self._set_active_nav(name)
        for w in self.content.winfo_children():
            w.destroy()
        if name == "Dashboard":
            self.view_dashboard()
        elif name == "Readings":
            self.view_readings()
        elif name == "Quarterly":
            self.view_quarterly()
        elif name == "Months & Rates":
            self.view_months()
        elif name == "Help":
            self.view_help()

    # ---------- DASHBOARD ----------
    def view_dashboard(self):
        names = self.months_names()
        if not self.current_month or self.current_month not in names:
            self.current_month = names[-1] if names else None

        wrap = tk.Frame(self.content, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=26, pady=20)

        top = tk.Frame(wrap, bg=APP_BG)
        top.pack(fill="x")
        tk.Label(top, text="Dashboard", bg=APP_BG, fg=TEXT, font=FONT_H1).pack(side="left")
        if names:
            sel = tk.StringVar(value=self.current_month)
            cb = ttk.Combobox(top, values=names, textvariable=sel, state="readonly", width=12)
            cb.pack(side="right")
            cb.bind("<<ComboboxSelected>>",
                    lambda e: (setattr(self, "current_month", sel.get()), self.show_view("Dashboard")))
            tk.Label(top, text="Month:", bg=APP_BG, fg=MUTED, font=FONT).pack(side="right", padx=8)

        if not names:
            self._empty_state(wrap, "No data yet. Use File → Import from Excel, or add a month under “Months & Rates”.")
            return

        m = self.store.month(self.current_month)
        ref = m.get("meter_ref", "")
        if ref:
            tk.Label(wrap, text=f"Meter Ref: {ref}", bg=APP_BG, fg=MUTED,
                     font=FONT_SMALL).pack(anchor="w", pady=(4, 0))

        totals = month_totals(m)
        chans = m.get("channels", ["Ch1", "Ch2", "Ch3"])
        ym = name_to_ym(m["name"])
        act = actual_saving(self.store, ym) if ym else None

        cards = tk.Frame(wrap, bg=APP_BG)
        cards.pack(fill="x", pady=16)
        for i in range(3):
            self._card(cards, chans[i], f"{totals[i]:,.0f}", "units this month",
                       CH_COLORS[i], col=i)
        if act is not None:
            self._card(cards, "Saving", f"Rs {act:,.0f}", "this month (actual)", GOOD, col=3)
        else:
            self._card(cards, "Est. Savings", f"Rs {est_savings(m):,.0f}", "this month", GOOD, col=3)
        for i in range(4):
            cards.grid_columnconfigure(i, weight=1, uniform="cards")

        # chart
        cwrap = tk.Frame(wrap, bg=CARD_BG, highlightbackground=LINE, highlightthickness=1)
        cwrap.pack(fill="both", expand=True, pady=(6, 0))
        tk.Label(cwrap, text="Monthly units by channel", bg=CARD_BG, fg=TEXT,
                 font=FONT_H2).pack(anchor="w", padx=18, pady=(14, 0))
        legend = tk.Frame(cwrap, bg=CARD_BG)
        legend.pack(anchor="w", padx=18, pady=(6, 0))
        # legend uses the most-recent month's channel labels
        for i in range(3):
            dot = tk.Label(legend, text="■", bg=CARD_BG, fg=CH_COLORS[i], font=FONT)
            dot.pack(side="left", padx=(0, 2))
            tk.Label(legend, text=chans[i], bg=CARD_BG, fg=MUTED,
                     font=FONT_SMALL).pack(side="left", padx=(0, 14))
        canvas = tk.Canvas(cwrap, bg=CARD_BG, highlightthickness=0, height=300)
        canvas.pack(fill="both", expand=True, padx=10, pady=12)
        self.after(60, lambda: self._draw_chart(canvas))

    def _card(self, parent, title, value, sub, color, col):
        c = tk.Frame(parent, bg=CARD_BG, highlightbackground=LINE, highlightthickness=1)
        c.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 10, 0))
        bar = tk.Frame(c, bg=color, height=4)
        bar.pack(fill="x")
        inner = tk.Frame(c, bg=CARD_BG)
        inner.pack(fill="both", expand=True, padx=16, pady=14)
        tk.Label(inner, text=title.upper(), bg=CARD_BG, fg=MUTED,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(inner, text=value, bg=CARD_BG, fg=TEXT, font=FONT_BIG).pack(anchor="w", pady=(4, 0))
        tk.Label(inner, text=sub, bg=CARD_BG, fg=MUTED, font=FONT_SMALL).pack(anchor="w")

    def _draw_chart(self, canvas):
        if not canvas.winfo_exists():
            return
        canvas.delete("all")
        W = canvas.winfo_width() or 800
        H = canvas.winfo_height() or 300
        pad_l, pad_r, pad_t, pad_b = 46, 16, 14, 46
        months = self.store.months
        data = [(m["name"], month_totals(m)) for m in months]
        if not data:
            return
        peak = max((max(t) for _, t in data), default=0) or 1
        # nice top
        import math
        step = 10 ** max(0, len(str(int(peak))) - 2)
        top = math.ceil(peak / step) * step if step else peak
        top = max(top, 1)
        plot_w = W - pad_l - pad_r
        plot_h = H - pad_t - pad_b
        # gridlines + y labels
        for g in range(5):
            y = pad_t + plot_h * g / 4
            val = top * (1 - g / 4)
            canvas.create_line(pad_l, y, W - pad_r, y, fill="#EEF1F6")
            canvas.create_text(pad_l - 8, y, text=f"{val:,.0f}", anchor="e",
                               fill=MUTED, font=FONT_SMALL)
        n = len(data)
        group_w = plot_w / n
        bar_w = min(13, group_w / 4.2)
        for gi, (name, totals) in enumerate(data):
            gx = pad_l + group_w * gi + group_w / 2
            for i in range(3):
                v = totals[i] if i < len(totals) else 0
                bh = (v / top) * plot_h
                x0 = gx + (i - 1) * (bar_w + 2) - bar_w / 2
                y0 = pad_t + plot_h - bh
                canvas.create_rectangle(x0, y0, x0 + bar_w, pad_t + plot_h,
                                        fill=CH_COLORS[i], outline="")
            canvas.create_text(gx, H - pad_b + 14, text=name.replace("25", "'25").replace("26", "'26"),
                               fill=MUTED, font=("Segoe UI", 8), angle=0)
        canvas.create_line(pad_l, pad_t + plot_h, W - pad_r, pad_t + plot_h, fill=LINE)

    # ---------- READINGS ----------
    def view_readings(self):
        names = self.months_names()
        if not names:
            wrap = tk.Frame(self.content, bg=APP_BG)
            wrap.pack(fill="both", expand=True, padx=26, pady=20)
            tk.Label(wrap, text="Readings", bg=APP_BG, fg=TEXT, font=FONT_H1).pack(anchor="w")
            self._empty_state(wrap, "No months yet. Import from Excel or add one in “Months & Rates”.")
            return
        if not self.current_month or self.current_month not in names:
            self.current_month = names[-1]

        wrap = tk.Frame(self.content, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=26, pady=20)

        top = tk.Frame(wrap, bg=APP_BG)
        top.pack(fill="x")
        tk.Label(top, text="Readings", bg=APP_BG, fg=TEXT, font=FONT_H1).pack(side="left")

        sel = tk.StringVar(value=self.current_month)
        cb = ttk.Combobox(top, values=names, textvariable=sel, state="readonly", width=12)
        cb.pack(side="left", padx=14)
        cb.bind("<<ComboboxSelected>>",
                lambda e: (setattr(self, "current_month", sel.get()), self.show_view("Readings")))

        ttk.Button(top, text="+ Add Day", style="Accent.TButton",
                   command=self.add_day).pack(side="left")
        ttk.Button(top, text="Delete Day", style="Danger.TButton",
                   command=self.delete_day).pack(side="left", padx=8)

        m = self.store.month(self.current_month)
        chans = m.get("channels", ["Ch1", "Ch2", "Ch3"])

        tk.Label(wrap, text="Double-click a Reading or Date cell to edit. Units are calculated automatically.",
                 bg=APP_BG, fg=MUTED, font=FONT_SMALL).pack(anchor="w", pady=(10, 6))

        table = tk.Frame(wrap, bg=CARD_BG, highlightbackground=LINE, highlightthickness=1)
        table.pack(fill="both", expand=True)

        cols = ("date", "r1", "u1", "r2", "u2", "r3", "u3")
        heads = ("Date", f"{chans[0]}  Read", "Units",
                 f"{chans[1]}  Read", "Units", f"{chans[2]}  Read", "Units")
        self.tree = ttk.Treeview(table, columns=cols, show="headings", selectmode="browse")
        for c, h in zip(cols, heads):
            self.tree.heading(c, text=h)
            w = 120 if c == "date" else (110 if c.startswith("r") else 78)
            anchor = "w" if c == "date" else "e"
            self.tree.column(c, width=w, anchor=anchor, stretch=(c != "date"))
        for i in (1, 3, 5):
            self.tree.tag_configure(f"x", )
        self.tree.tag_configure("odd", background="#F7F9FC")
        vs = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._edit_cell)

        self._refresh_tree()

        # footer totals
        tot = month_totals(m)
        foot = tk.Frame(wrap, bg=APP_BG)
        foot.pack(fill="x", pady=(10, 0))
        for i in range(3):
            tk.Label(foot, text=f"{chans[i]} total:", bg=APP_BG, fg=MUTED,
                     font=FONT).pack(side="left", padx=(0, 4))
            tk.Label(foot, text=f"{tot[i]:,.0f} units", bg=APP_BG, fg=CH_COLORS[i],
                     font=FONT_BOLD).pack(side="left", padx=(0, 22))

    def _refresh_tree(self):
        m = self.store.month(self.current_month)
        rows = compute_units(m.get("readings", []))
        for it in self.tree.get_children():
            self.tree.delete(it)
        def fmt(v):
            return "" if v is None else (f"{v:,.0f}" if float(v) == int(v) else f"{v:,.1f}")
        for idx, r in enumerate(rows):
            vals = (r["date"], fmt(r["r1"]), fmt(r["u1"]), fmt(r["r2"]),
                    fmt(r["u2"]), fmt(r["r3"]), fmt(r["u3"]))
            tag = "odd" if idx % 2 else "even"
            self.tree.insert("", "end", iid=str(idx), values=vals, tags=(tag,))

    def _edit_cell(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)   # '#1'..'#7'
        row = self.tree.identify_row(event.y)
        if not row:
            return
        cidx = int(col[1:]) - 1
        if cidx in (2, 4, 6):   # units columns are read-only
            return
        x, y, w, h = self.tree.bbox(row, col)
        cur = self.tree.set(row, self.tree["columns"][cidx])
        edit = tk.Entry(self.tree, font=FONT, justify=("left" if cidx == 0 else "right"),
                        relief="solid", bd=1)
        edit.insert(0, cur.replace(",", ""))
        edit.select_range(0, "end")
        edit.place(x=x, y=y, width=w, height=h)
        edit.focus_set()

        def commit(ev=None):
            val = edit.get().strip()
            edit.destroy()
            self._apply_edit(int(row), cidx, val)

        def cancel(ev=None):
            edit.destroy()

        edit.bind("<Return>", commit)
        edit.bind("<Escape>", cancel)
        edit.bind("<FocusOut>", commit)

    def _apply_edit(self, ridx, cidx, val):
        m = self.store.month(self.current_month)
        reads = m["readings"]
        if ridx >= len(reads):
            return
        if cidx == 0:  # date
            try:
                d = datetime.datetime.strptime(val, "%Y-%m-%d").date()
                reads[ridx]["date"] = d.isoformat()
            except ValueError:
                messagebox.showwarning(APP_NAME, "Use date format YYYY-MM-DD (e.g. 2026-03-15).")
                return
        else:
            key = {1: "r1", 3: "r2", 5: "r3"}[cidx]
            if val == "":
                reads[ridx][key] = None
            else:
                try:
                    reads[ridx][key] = float(val)
                except ValueError:
                    messagebox.showwarning(APP_NAME, "Reading must be a number.")
                    return
        reads.sort(key=lambda r: r["date"])
        self.store.save()
        self.show_view("Readings")

    def add_day(self):
        m = self.store.month(self.current_month)
        reads = m["readings"]
        if reads:
            last = datetime.date.fromisoformat(reads[-1]["date"])
            nd = last + datetime.timedelta(days=1)
            prev = reads[-1]
            new = {"date": nd.isoformat(), "r1": prev.get("r1"),
                   "r2": prev.get("r2"), "r3": prev.get("r3")}
        else:
            new = {"date": datetime.date.today().isoformat(), "r1": None, "r2": None, "r3": None}
        reads.append(new)
        self.store.save()
        self.show_view("Readings")
        kids = self.tree.get_children()
        if kids:
            self.tree.see(kids[-1])
            self.tree.selection_set(kids[-1])

    def delete_day(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Select a day to delete.")
            return
        m = self.store.month(self.current_month)
        ridx = int(sel[0])
        if 0 <= ridx < len(m["readings"]):
            d = m["readings"][ridx]["date"]
            if messagebox.askyesno(APP_NAME, f"Delete reading for {d}?"):
                del m["readings"][ridx]
                self.store.save()
                self.show_view("Readings")

    # ---------- QUARTERLY ----------
    def view_quarterly(self):
        wrap = tk.Frame(self.content, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=26, pady=20)
        top = tk.Frame(wrap, bg=APP_BG)
        top.pack(fill="x")
        tk.Label(top, text="Quarterly", bg=APP_BG, fg=TEXT, font=FONT_H1).pack(side="left")

        quarters = group_quarters(self.store.data.get("summary", []))
        if not quarters:
            self._empty_state(wrap, "No quarterly summary yet. Use File → Import from Excel "
                                    "to load the “DATA HSS” summary sheet.")
            return

        qnames = [q for q, _ in quarters]
        if getattr(self, "current_q", None) not in qnames:
            self.current_q = qnames[-1]
        qsel = tk.StringVar(value=self.current_q)
        cb = ttk.Combobox(top, values=qnames, textvariable=qsel, state="readonly", width=10)
        cb.pack(side="right")
        cb.bind("<<ComboboxSelected>>",
                lambda e: (setattr(self, "current_q", qsel.get()), self.show_view("Quarterly")))
        tk.Label(top, text="Quarter:", bg=APP_BG, fg=MUTED, font=FONT).pack(side="right", padx=8)

        rows = dict(quarters)[self.current_q]
        mlabels = [f"{MONTH_ABBR[int(r['month'].split('-')[1])]} {r['month'][:4]}" for r in rows]

        # per-quarter detail card
        card = tk.Frame(wrap, bg=CARD_BG, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="x", pady=(14, 0))
        tk.Label(card, text=f"{self.current_q}   ·   {'  →  '.join(mlabels)}",
                 bg=CARD_BG, fg=TEXT, font=FONT_H2).pack(anchor="w", padx=18, pady=(14, 4))
        tk.Label(card, text="Net units per channel each month, the quarter total, and the "
                            "meter reading at quarter-end.",
                 bg=CARD_BG, fg=MUTED, font=FONT_SMALL).pack(anchor="w", padx=18)

        cols = ("ch", "m1", "m2", "m3", "tot", "close")
        heads = ("Channel", mlabels[0] if len(mlabels) > 0 else "M1",
                 mlabels[1] if len(mlabels) > 1 else "M2",
                 mlabels[2] if len(mlabels) > 2 else "M3",
                 "Qtr Total", "Closing Mtr")
        tv = ttk.Treeview(card, columns=cols, show="headings", height=4)
        widths = (150, 95, 95, 95, 95, 110)
        for c, h, w in zip(cols, heads, widths):
            tv.heading(c, text=h)
            tv.column(c, width=w, anchor=("w" if c == "ch" else "e"), stretch=False)
        tv.pack(fill="x", padx=14, pady=12)

        disp = {"peak": "Peak (Import)", "offpeak": "Off-peak (Import)", "export": "Export"}
        last_rates = rate_for_month(self.store, rows[-1]["month"])
        totals = {}
        for ch in ("peak", "offpeak", "export"):
            nets = [(r[ch]["net"] or 0) for r in rows]
            while len(nets) < 3:
                nets.append(0)
            tot = sum(nets)
            totals[ch] = tot
            close = rows[-1][ch]["close"]
            tv.insert("", "end", values=(disp[ch], f"{nets[0]:,.0f}", f"{nets[1]:,.0f}",
                                         f"{nets[2]:,.0f}", f"{tot:,.0f}",
                                         f"{close:,.0f}" if close is not None else "—"))

        net_units = totals["export"] - totals["offpeak"] - totals["peak"]
        saving = sum((r.get("saving") or 0) for r in rows)
        settle, credit, peak_cost = settlement_charge(
            totals["peak"], totals["offpeak"], totals["export"], last_rates)

        # settlement breakdown using the user's formula
        calc = tk.Frame(card, bg=CARD_BG)
        calc.pack(fill="x", padx=18, pady=(2, 4))
        tk.Label(calc, text="Est. settlement", bg=CARD_BG, fg=TEXT,
                 font=FONT_BOLD).pack(anchor="w")
        if settle is None:
            tk.Label(calc, text="Set Buyback and Peak rates under “Months & Rates” to calculate.",
                     bg=CARD_BG, fg=WARN, font=FONT_SMALL).pack(anchor="w")
        else:
            bb = last_rates.get("buyback"); pr = last_rates.get("peak")
            formula = (f"(Export {totals['export']:,.0f} − Off-peak {totals['offpeak']:,.0f}) "
                       f"× Rs {bb:g}   −   Peak {totals['peak']:,.0f} × Rs {pr:g}")
            tk.Label(calc, text=formula, bg=CARD_BG, fg=MUTED, font=FONT_SMALL).pack(anchor="w")
            line2 = (f"= Rs {credit:,.0f}  −  Rs {peak_cost:,.0f}  "
                     f"=  Rs {settle:,.0f}")
            col = GOOD if settle >= 0 else WARN
            tk.Label(calc, text=line2, bg=CARD_BG, fg=col, font=FONT_BOLD).pack(anchor="w", pady=(2, 0))
            tk.Label(calc, text="(rates shown are this quarter's latest month)",
                     bg=CARD_BG, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w")

        foot = tk.Frame(card, bg=CARD_BG)
        foot.pack(fill="x", padx=18, pady=(8, 16))
        tk.Label(foot, text="Net  (Export − Off-peak − Peak):", bg=CARD_BG, fg=MUTED,
                 font=FONT).pack(side="left")
        tk.Label(foot, text=f"{net_units:,.0f} units", bg=CARD_BG, fg=TEXT,
                 font=FONT_BOLD).pack(side="left", padx=(6, 26))
        tk.Label(foot, text="Quarter Saving (actual):", bg=CARD_BG, fg=MUTED,
                 font=FONT).pack(side="left")
        tk.Label(foot, text=f"Rs {saving:,.0f}", bg=CARD_BG, fg=GOOD,
                 font=FONT_BOLD).pack(side="left", padx=(6, 0))

        # all-quarters overview
        tk.Label(wrap, text="All quarters", bg=APP_BG, fg=TEXT,
                 font=FONT_H2).pack(anchor="w", pady=(20, 6))
        ov = tk.Frame(wrap, bg=CARD_BG, highlightbackground=LINE, highlightthickness=1)
        ov.pack(fill="both", expand=True)
        ocols = ("q", "pk", "op", "ex", "net", "settle", "sav")
        oheads = ("Quarter", "Peak", "Off-peak", "Export", "Net",
                  "Est. Settle (Rs)", "Saving (Rs)")
        otv = ttk.Treeview(ov, columns=ocols, show="headings")
        for c, h in zip(ocols, oheads):
            otv.heading(c, text=h)
            w = 115 if c == "q" else (130 if c in ("settle", "sav") else 95)
            otv.column(c, width=w, anchor=("w" if c == "q" else "e"),
                       stretch=(c != "q"))
        gpk = gop = gex = gsv = gst = 0
        for q, qr in quarters:
            pk = sum((r["peak"]["net"] or 0) for r in qr)
            op = sum((r["offpeak"]["net"] or 0) for r in qr)
            ex = sum((r["export"]["net"] or 0) for r in qr)
            sv = sum((r.get("saving") or 0) for r in qr)
            st, _, _ = settlement_charge(pk, op, ex, rate_for_month(self.store, qr[-1]["month"]))
            gpk += pk; gop += op; gex += ex; gsv += sv
            stxt = f"{st:,.0f}" if st is not None else "—"
            if st is not None:
                gst += st
            otv.insert("", "end", values=(q, f"{pk:,.0f}", f"{op:,.0f}", f"{ex:,.0f}",
                                          f"{ex-op-pk:,.0f}", stxt, f"{sv:,.0f}"))
        otv.tag_configure("tot", background="#EEF3FB")
        otv.insert("", "end", tags=("tot",),
                   values=("TOTAL", f"{gpk:,.0f}", f"{gop:,.0f}", f"{gex:,.0f}",
                           f"{gex-gop-gpk:,.0f}", f"{gst:,.0f}", f"{gsv:,.0f}"))
        otv.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------- MONTHS & RATES ----------
    def view_months(self):
        wrap = tk.Frame(self.content, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=26, pady=20)
        top = tk.Frame(wrap, bg=APP_BG)
        top.pack(fill="x")
        tk.Label(top, text="Months & Rates", bg=APP_BG, fg=TEXT, font=FONT_H1).pack(side="left")
        ttk.Button(top, text="+ New Month", style="Accent.TButton",
                   command=self.new_month).pack(side="right")

        body = tk.Frame(wrap, bg=APP_BG)
        body.pack(fill="both", expand=True, pady=14)

        left = tk.Frame(body, bg=CARD_BG, highlightbackground=LINE, highlightthickness=1, width=220)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        tk.Label(left, text="MONTHS", bg=CARD_BG, fg=MUTED,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=14, pady=(12, 4))
        self.mlist = tk.Listbox(left, font=FONT, activestyle="none", bd=0,
                                highlightthickness=0, selectbackground="#D6E4FF",
                                selectforeground=TEXT)
        self.mlist.pack(fill="both", expand=True, padx=8, pady=(0, 10))
        for n in self.months_names():
            self.mlist.insert("end", "  " + n)
        self.mlist.bind("<<ListboxSelect>>", self._month_selected)

        self.mdetail = tk.Frame(body, bg=APP_BG)
        self.mdetail.pack(side="left", fill="both", expand=True, padx=(16, 0))
        if self.months_names():
            self.mlist.selection_set(len(self.months_names()) - 1)
            self._month_selected(None)

    def _month_selected(self, _):
        sel = self.mlist.curselection()
        if not sel:
            return
        name = self.months_names()[sel[0]]
        m = self.store.month(name)
        for w in self.mdetail.winfo_children():
            w.destroy()

        card = tk.Frame(self.mdetail, bg=CARD_BG, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="both", expand=True)
        pad = tk.Frame(card, bg=CARD_BG)
        pad.pack(fill="both", expand=True, padx=22, pady=20)

        tk.Label(pad, text=name, bg=CARD_BG, fg=TEXT, font=FONT_H2).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))

        ent = {}

        def field(r, label, value):
            tk.Label(pad, text=label, bg=CARD_BG, fg=MUTED, font=FONT).grid(
                row=r, column=0, sticky="w", pady=6, padx=(0, 14))
            e = tk.Entry(pad, font=FONT, width=30, relief="solid", bd=1)
            e.insert(0, "" if value is None else str(value))
            e.grid(row=r, column=1, sticky="w", pady=6)
            return e

        ent["meter_ref"] = field(1, "Meter Ref", m.get("meter_ref", ""))
        ch = m.get("channels", ["Ch1", "Ch2", "Ch3"])
        ent["c1"] = field(2, "Channel 1 label", ch[0])
        ent["c2"] = field(3, "Channel 2 label", ch[1])
        ent["c3"] = field(4, "Channel 3 label", ch[2])
        rates = m.get("rates", {})
        ent["peak"] = field(5, "Peak rate (Rs/unit)", rates.get("peak", ""))
        ent["offpeak"] = field(6, "Off-peak rate (Rs/unit)", rates.get("offpeak", ""))
        ent["buyback"] = field(7, "Buyback rate (Rs/unit)", rates.get("buyback", ""))

        def save():
            m["meter_ref"] = ent["meter_ref"].get().strip()
            m["channels"] = [ent["c1"].get().strip() or "Ch1",
                             ent["c2"].get().strip() or "Ch2",
                             ent["c3"].get().strip() or "Ch3"]
            r = {}
            for k in ("peak", "offpeak", "buyback"):
                v = ent[k].get().strip()
                if v:
                    try:
                        r[k] = float(v)
                    except ValueError:
                        pass
            m["rates"] = r
            self.store.save()
            self._flash_saved()
            messagebox.showinfo(APP_NAME, "Saved.")

        btns = tk.Frame(pad, bg=CARD_BG)
        btns.grid(row=8, column=0, columnspan=2, sticky="w", pady=(18, 0))
        ttk.Button(btns, text="Save changes", style="Accent.TButton", command=save).pack(side="left")
        ttk.Button(btns, text="Delete month", style="Danger.TButton",
                   command=lambda: self._delete_month(name)).pack(side="left", padx=10)

        info = compute_units(m.get("readings", []))
        tk.Label(pad, text=f"{len(info)} daily readings recorded.",
                 bg=CARD_BG, fg=MUTED, font=FONT_SMALL).grid(
            row=9, column=0, columnspan=2, sticky="w", pady=(14, 0))

    def _delete_month(self, name):
        if messagebox.askyesno(APP_NAME, f"Delete the entire month “{name}” and all its readings?"):
            self.store.delete_month(name)
            if self.current_month == name:
                self.current_month = self.months_names()[-1] if self.months_names() else None
            self.show_view("Months & Rates")

    def new_month(self):
        name = simpledialog.askstring(APP_NAME, "Month name (e.g. Jul26):", parent=self)
        if not name:
            return
        name = name.strip()
        if self.store.month(name):
            messagebox.showwarning(APP_NAME, "That month already exists.")
            return
        self.store.add_month(name)
        self.current_month = name
        self.show_view("Months & Rates")

    # ---------- HELP ----------
    def view_help(self):
        wrap = tk.Frame(self.content, bg=APP_BG)
        wrap.pack(fill="both", expand=True, padx=26, pady=20)
        tk.Label(wrap, text="How to use", bg=APP_BG, fg=TEXT, font=FONT_H1).pack(anchor="w")
        card = tk.Frame(wrap, bg=CARD_BG, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="both", expand=True, pady=14)
        txt = (
            "Dashboard\n"
            "   Pick a month to see total units per channel and an estimated saving, plus a\n"
            "   bar chart comparing every month.\n\n"
            "Readings\n"
            "   Choose a month, then double-click any Reading or Date cell to edit it.\n"
            "   Daily Units are calculated for you (today's reading minus yesterday's).\n"
            "   Use “+ Add Day” to append the next date, “Delete Day” to remove one.\n\n"
            "Quarterly\n"
            "   Each quarter rolls up its three months: net units per channel per month,\n"
            "   the quarter total and the meter reading at quarter-end. The estimated\n"
            "   settlement uses your formula:\n"
            "        (Export − Off-peak) × Buyback rate  −  Peak × Peak rate\n"
            "   The actual saving comes from your DATA HSS summary.\n\n"
            "Months & Rates\n"
            "   Rename channels, set the Peak / Off-peak / Buyback rates used for the saving\n"
            "   estimate, edit the meter reference, or add and delete months.\n\n"
            "Saving\n"
            "   Everything is saved automatically to meter_data.json next to the program.\n"
            "   File → Export lets you write CSV or Excel copies any time.\n\n"
            f"Excel import/export:  {'available' if HAVE_XLSX else 'install openpyxl to enable'}\n"
        )
        tk.Label(card, text=txt, bg=CARD_BG, fg=TEXT, font=FONT, justify="left",
                 anchor="nw").pack(fill="both", expand=True, padx=22, pady=18)

    def _empty_state(self, parent, msg):
        f = tk.Frame(parent, bg=APP_BG)
        f.pack(expand=True)
        tk.Label(f, text="📂", bg=APP_BG, font=("Segoe UI", 40)).pack(pady=(40, 6))
        tk.Label(f, text=msg, bg=APP_BG, fg=MUTED, font=FONT, wraplength=460,
                 justify="center").pack()

    def about(self):
        messagebox.showinfo(
            "About " + APP_NAME,
            f"{APP_NAME} v{APP_VER}\nA friendly manager for home solar meter readings.\n\n"
            f"Data file:\n{DATA_FILE}\n\nExcel support: {'on' if HAVE_XLSX else 'off (no openpyxl)'}")

    # ---------- IMPORT / EXPORT ----------
    def import_excel(self):
        if not HAVE_XLSX:
            messagebox.showinfo(APP_NAME, "Excel support needs the openpyxl package.\n"
                                          "It is included in the .exe build.")
            return
        path = filedialog.askopenfilename(title="Import meter readings",
                                          filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            added = self._parse_excel(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Import failed:\n{e}")
            return
        self.store.save()
        self.current_month = self.months_names()[-1] if self.months_names() else None
        self.show_view("Dashboard")
        messagebox.showinfo(APP_NAME, f"Imported / updated {added} month(s).")

    def _parse_excel(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        skip = {"ShortCuts", "Practice", "DATA HSS"}
        count = 0
        # monthly summary from DATA HSS
        if "DATA HSS" in wb.sheetnames:
            ws = wb["DATA HSS"]
            summ = []
            for r in range(1, ws.max_row + 1):
                d = ws.cell(r, 1).value
                if isinstance(d, (datetime.datetime, datetime.date)):
                    def n(c):
                        v = ws.cell(r, c).value
                        return float(v) if isinstance(v, (int, float)) else None
                    summ.append({"month": d.strftime("%Y-%m"),
                                 "peak": {"open": n(2), "close": n(3), "net": n(4)},
                                 "offpeak": {"open": n(5), "close": n(6), "net": n(7)},
                                 "export": {"open": n(8), "close": n(9), "net": n(10)},
                                 "saving": n(11)})
            if summ:
                self.store.data["summary"] = summ
        for name in wb.sheetnames:
            if name.strip() in {s.strip() for s in skip}:
                continue
            ws = wb[name]
            hdr = None
            for r in range(1, ws.max_row + 1):
                v = ws.cell(r, 1).value
                if isinstance(v, str) and v.strip().lower() == "date":
                    hdr = r
                    break
            if not hdr:
                continue
            labs = [ws.cell(hdr, c).value for c in (2, 4, 6)]
            chans = [(str(x).strip() if x else f"Ch{i+1}") for i, x in enumerate(labs)]
            reads = []
            for r in range(hdr + 2, ws.max_row + 1):
                d = ws.cell(r, 1).value
                if isinstance(d, (datetime.datetime, datetime.date)):
                    def num(c):
                        x = ws.cell(r, c).value
                        return float(x) if isinstance(x, (int, float)) else None
                    reads.append({"date": d.strftime("%Y-%m-%d"),
                                  "r1": num(2), "r2": num(4), "r3": num(6)})
            mname = name.strip()
            existing = self.store.month(mname)
            if existing:
                existing["channels"] = chans
                existing["readings"] = reads
            else:
                self.store.months.append({"name": mname, "meter_ref": "",
                                          "channels": chans, "rates": {}, "readings": reads})
            count += 1
        return count

    def export_csv(self, all_months):
        if all_months:
            path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                initialfile="meter_readings_all.csv",
                                                filetypes=[("CSV", "*.csv")])
            if not path:
                return
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Month", "Date", "Ch1", "Ch1 Read", "Ch1 Units",
                            "Ch2", "Ch2 Read", "Ch2 Units", "Ch3", "Ch3 Read", "Ch3 Units"])
                for m in self.store.months:
                    ch = m.get("channels", ["", "", ""])
                    for r in compute_units(m.get("readings", [])):
                        w.writerow([m["name"], r["date"], ch[0], r["r1"], r["u1"],
                                    ch[1], r["r2"], r["u2"], ch[2], r["r3"], r["u3"]])
        else:
            if not self.current_month:
                return
            m = self.store.month(self.current_month)
            path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                initialfile=f"{self.current_month}.csv",
                                                filetypes=[("CSV", "*.csv")])
            if not path:
                return
            ch = m.get("channels", ["", "", ""])
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Date", f"{ch[0]} Read", f"{ch[0]} Units",
                            f"{ch[1]} Read", f"{ch[1]} Units",
                            f"{ch[2]} Read", f"{ch[2]} Units"])
                for r in compute_units(m.get("readings", [])):
                    w.writerow([r["date"], r["r1"], r["u1"], r["r2"], r["u2"], r["r3"], r["u3"]])
        messagebox.showinfo(APP_NAME, "CSV exported.")

    def export_excel(self):
        if not HAVE_XLSX:
            messagebox.showinfo(APP_NAME, "Excel export needs openpyxl (included in the .exe build).")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialfile="Meter_Readings_export.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        from openpyxl import Workbook
        from openpyxl.styles import Font as XF, PatternFill
        wb = Workbook()
        wb.remove(wb.active)
        for m in self.store.months:
            ws = wb.create_sheet(m["name"][:31])
            ch = m.get("channels", ["Ch1", "Ch2", "Ch3"])
            ws.append([f"Meter Readings {m['name']}"])
            if m.get("meter_ref"):
                ws.append([f"Meter Ref: {m['meter_ref']}"])
            ws.append([])
            head = ["Date", f"{ch[0]} Read", "Units", f"{ch[1]} Read", "Units",
                    f"{ch[2]} Read", "Units"]
            ws.append(head)
            for c in range(1, 8):
                cell = ws.cell(ws.max_row, c)
                cell.font = XF(bold=True)
                cell.fill = PatternFill("solid", start_color="E7ECF3")
            for r in compute_units(m.get("readings", [])):
                ws.append([r["date"], r["r1"], r["u1"], r["r2"], r["u2"], r["r3"], r["u3"]])
            for col in "ABCDEFG":
                ws.column_dimensions[col].width = 13
        wb.save(path)
        messagebox.showinfo(APP_NAME, "Excel file exported.")


def main():
    store = Store()
    app = App(store)
    app.protocol("WM_DELETE_WINDOW", lambda: (store.save(), app.destroy()))
    app.mainloop()


if __name__ == "__main__":
    main()
