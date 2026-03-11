import json
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
from tab4_deep_analysis import TabPhanTichSKChuyenSau

import re
import unicodedata
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SAVE_DIR = os.path.join(BASE_DIR, "KETQUA")
PROFILE_DIR = os.path.join(BASE_DIR, "Ho so luu")
DEFAULT_TEMPLATE_DIR = os.path.join(BASE_DIR, "Mau 1")

os.makedirs(SAVE_DIR, exist_ok=True)
os.makedirs(PROFILE_DIR, exist_ok=True)

WORD_FILES = [
    "1. BBGLK.docx",
    "2. BB Giao nhan tai lieu.docx",
    "3. BB Kiem tra dien thoai.docx",
    "4. BB Niem phong.docx",
    "5. BB Trich xuat SK TKNH.docx",
    "6. Cam doan.docx",
    "7. Giao nhan nguoi.docx",
]

# ================= PLACEHOLDER ENTRY =================
class PlaceholderEntry(ttk.Entry):
    def __init__(self, master=None, placeholder="", color="grey", **kwargs):
        super().__init__(master, **kwargs)
        self.placeholder = placeholder
        self.placeholder_color = color
        self.default_fg = self.cget("foreground")
        self.insert(0, placeholder)
        self.config(foreground=self.placeholder_color)
        self.bind("<FocusIn>", self.clear)
        self.bind("<FocusOut>", self.add)

    def clear(self, e):
        if self.get() == self.placeholder:
            self.delete(0, "end")
            self.config(foreground=self.default_fg)

    def add(self, e):
        if not self.get():
            self.insert(0, self.placeholder)
            self.config(foreground=self.placeholder_color)

    def get_real(self):
        if self.get() == self.placeholder:
            return ""
        return self.get()



# =================================================================================
# TAB 3: PHÂN TÍCH SK TKNH (v13.1)
# =================================================================================
class TabPhanTichSK(ttk.Frame):
    """
    Tab phân tích sao kê TKNH theo cấu hình 1 bank (form giống nhau) và danh sách game.

    - Chọn nhiều file (xlsx/csv) -> gộp, chuẩn hoá, lọc theo keyword game.
    - Tự dò header + cột dữ liệu, hỏi xác nhận trước khi phân tích, cho phép chỉnh sửa.
    - Chỉ thống kê theo Debit/Ghi nợ (tiền đi). Dòng match thiếu Debit vẫn giữ lại (không tính tiền).
    - Xuất Excel (2 sheet: KET_QUA + THONG_KE) và PDF (trang 1 thống kê, trang 2 bảng kết quả) chuẩn A4.

    Nâng cấp (Tab3 - Kết quả):
    - Sort: click vào header cột để sắp xếp tăng/giảm (type-aware cho Ngày/Giờ/Tiền đi).
    - Filter: lọc theo 1 cột hoặc lọc toàn bảng (contains, không phân biệt hoa/thường).
    - Search: tìm tiếp trong bảng (không làm mất dữ liệu, chỉ nhảy/highlight).
    - Copy: Ctrl+C / nút Copy Selected / Copy All (TSV có header).
    """

    BANK_CFG_FILE = os.path.join(BASE_DIR, "bank_config.json")
    GAME_CFG_FILE = os.path.join(BASE_DIR, "game_config.json")

    DEFAULT_BANKS = {
        "VCB": {
            "identify_cols": ["Nội dung", "Nội dung chi tiết", "Diễn giải", "Description", "Transactions in detail"],
            "money_cols": ["Ghi nợ", "Debit", "Số tiền ghi nợ", "Debit amount", "Phát sinh nợ"],
            "date_cols": ["Ngày giao dịch", "Ngày", "Transaction date", "Date", "Thời gian", "Date time"]
        },
        "TCB": {
            "identify_cols": ["Nội dung", "Diễn giải", "Mô tả", "Description"],
            "money_cols": ["Ghi nợ", "Debit", "Số tiền ghi nợ", "Phát sinh nợ"],
            "date_cols": ["Ngày", "Ngày giao dịch", "Date", "Thời gian", "Transaction time", "Transaction date"]
        },
        "VTB": {
            "identify_cols": ["Diễn giải", "Nội dung", "Description"],
            "money_cols": ["Ghi nợ", "Debit", "Số tiền ghi nợ", "Phát sinh nợ"],
            "date_cols": ["Ngày", "Ngày giao dịch", "Date", "Thời gian", "Transaction date"]
        },
        "MB": {
            "identify_cols": ["Diễn giải", "Nội dung", "Description"],
            "money_cols": ["Ghi nợ", "Debit", "Số tiền ghi nợ", "Phát sinh nợ"],
            "date_cols": ["Ngày", "Ngày giao dịch", "Date", "Thời gian"]
        },
        "BIDV": {
            "identify_cols": ["Diễn giải", "Nội dung", "Description"],
            "money_cols": ["Ghi nợ", "Debit", "Số tiền ghi nợ", "Phát sinh nợ"],
            "date_cols": ["Ngày", "Ngày giao dịch", "Date", "Thời gian"]
        },
    }

    DEFAULT_GAMES = [
        {"name": "Hitclub", "keywords": [".XXXXXXXX."], "match_type": "regex"},
        {"name": "78win", "keywords": ["CK XXXXXXXX"], "match_type": "regex"},
    ]

    def __init__(self, parent):
        super().__init__(parent)
        self.selected_files: List[str] = []
        self.df_all: Optional[pd.DataFrame] = None
        self.df_filtered: Optional[pd.DataFrame] = None  # full result after analyze
        self.df_view: Optional[pd.DataFrame] = None      # filtered/sorted view for table
        self.stats: Dict[str, Any] = {}
        self.mapping: Optional[Dict[str, Any]] = None
        self._log_file_path: Optional[str] = None

        # Result UI helpers
        self._sort_state: Dict[str, bool] = {}  # col -> asc?
        self._search_cache: List[str] = []      # list of Tree IIDs matching search
        self._search_pos: int = -1

        self._ensure_configs()
        self.bank_cfg = self._load_json(self.BANK_CFG_FILE, self.DEFAULT_BANKS)
        self.game_cfg = self._load_json(self.GAME_CFG_FILE, self.DEFAULT_GAMES)

        self._pdf_font_name = self._setup_pdf_font()
        self._build_ui()

    # ---------------- CONFIG ----------------
    def _ensure_configs(self):
        if not os.path.exists(self.BANK_CFG_FILE):
            with open(self.BANK_CFG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.DEFAULT_BANKS, f, ensure_ascii=False, indent=4)
        if not os.path.exists(self.GAME_CFG_FILE):
            with open(self.GAME_CFG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.DEFAULT_GAMES, f, ensure_ascii=False, indent=4)

    def _load_json(self, path: str, default):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default

    def _save_json(self, path: str, data):
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.log(f"Không thể lưu cấu hình: {e}", "error")

    def _setup_pdf_font(self) -> str:
        """
        Đăng ký font Unicode cho PDF (ưu tiên DejaVuSans).
        Nếu không có -> fallback Helvetica.
        """
        try:
            candidates = [
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
            ]
            for fp in candidates:
                if os.path.exists(fp):
                    pdfmetrics.registerFont(TTFont("DejaVuSans", fp))
                    return "DejaVuSans"
        except Exception:
            pass
        return "Helvetica"

    # ---------------- UI ----------------
    def _build_ui(self):
        # Top action row
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Button(top, text="Chọn file sao kê", command=self.choose_files).pack(side="left", padx=5)
        ttk.Button(top, text="Dò header/cột", command=self.detect_and_confirm).pack(side="left", padx=5)
        ttk.Button(top, text="Phân tích tất cả", command=self.analyze).pack(side="left", padx=5)
        ttk.Button(top, text="Xuất báo cáo Excel", command=self.export_excel).pack(side="left", padx=5)
        ttk.Button(top, text="Xuất báo cáo PDF", command=self.export_pdf).pack(side="left", padx=5)

        self.lbl_selected = ttk.Label(top, text="Đã chọn: 0 file")
        self.lbl_selected.pack(side="left", padx=15)

        # Config area
        pw = ttk.Panedwindow(self, orient="horizontal")
        pw.pack(fill="x", expand=False, padx=10, pady=6)

        left = ttk.Labelframe(pw, text="Ngân hàng (chỉ chọn một)")
        right = ttk.Labelframe(pw, text="Cổng game (có thể chọn nhiều)")
        pw.add(left, weight=1)
        pw.add(right, weight=1)

        # LEFT: bank tree
        bank_wrap = ttk.Frame(left)
        bank_wrap.pack(fill="both", expand=True, padx=6, pady=6)

        self.bank_tree = ttk.Treeview(bank_wrap, columns=("sel", "name", "identify", "money", "date"), show="headings", height=6)
        b_vsb = ttk.Scrollbar(bank_wrap, orient="vertical", command=self.bank_tree.yview)
        b_hsb = ttk.Scrollbar(bank_wrap, orient="horizontal", command=self.bank_tree.xview)
        self.bank_tree.configure(yscrollcommand=b_vsb.set, xscrollcommand=b_hsb.set)

        self.bank_tree.grid(row=0, column=0, sticky="nsew")
        b_vsb.grid(row=0, column=1, sticky="ns")
        b_hsb.grid(row=1, column=0, sticky="ew")
        bank_wrap.columnconfigure(0, weight=1)
        bank_wrap.rowconfigure(0, weight=1)

        self.bank_tree.heading("sel", text="Chọn")
        self.bank_tree.heading("name", text="Tên ngân hàng")
        self.bank_tree.heading("identify", text="Cột nội dung")
        self.bank_tree.heading("money", text="Cột ghi nợ")
        self.bank_tree.heading("date", text="Cột thời gian")

        self.bank_tree.column("sel", width=55, anchor="center", stretch=False)
        self.bank_tree.column("name", width=160, stretch=False)
        self.bank_tree.column("identify", width=380, stretch=True)
        self.bank_tree.column("money", width=220, stretch=True)
        self.bank_tree.column("date", width=220, stretch=True)

        self.bank_tree.bind("<Button-1>", self._on_bank_click)
        self.bank_tree.bind("<Double-1>", lambda e: self._show_cell_fulltext(self.bank_tree, e, title="Ngân hàng"))

        btns_left = ttk.Frame(left)
        btns_left.pack(pady=4)
        ttk.Button(btns_left, text="Thêm", command=self.add_bank).pack(side="left", padx=6)
        ttk.Button(btns_left, text="Sửa", command=self.edit_bank).pack(side="left", padx=6)
        ttk.Button(btns_left, text="Xoá", command=self.del_bank).pack(side="left", padx=6)

        # RIGHT: game tree
        game_wrap = ttk.Frame(right)
        game_wrap.pack(fill="both", expand=True, padx=6, pady=6)

        self.game_tree = ttk.Treeview(game_wrap, columns=("sel", "name", "keywords", "match"), show="headings", height=6)
        g_vsb = ttk.Scrollbar(game_wrap, orient="vertical", command=self.game_tree.yview)
        g_hsb = ttk.Scrollbar(game_wrap, orient="horizontal", command=self.game_tree.xview)
        self.game_tree.configure(yscrollcommand=g_vsb.set, xscrollcommand=g_hsb.set)

        self.game_tree.grid(row=0, column=0, sticky="nsew")
        g_vsb.grid(row=0, column=1, sticky="ns")
        g_hsb.grid(row=1, column=0, sticky="ew")
        game_wrap.columnconfigure(0, weight=1)
        game_wrap.rowconfigure(0, weight=1)

        self.game_tree.heading("sel", text="Chọn")
        self.game_tree.heading("name", text="Tên game")
        self.game_tree.heading("keywords", text="Từ khoá")
        self.game_tree.heading("match", text="Kiểu match")

        self.game_tree.column("sel", width=55, anchor="center", stretch=False)
        self.game_tree.column("name", width=140, stretch=False)
        self.game_tree.column("keywords", width=420, stretch=True)
        self.game_tree.column("match", width=110, anchor="center", stretch=False)

        self.game_tree.bind("<Button-1>", self._on_game_click)
        self.game_tree.bind("<Double-1>", lambda e: self._show_cell_fulltext(self.game_tree, e, title="Cổng game"))

        btns_right = ttk.Frame(right)
        btns_right.pack(pady=4)
        ttk.Button(btns_right, text="Thêm", command=self.add_game).pack(side="left", padx=6)
        ttk.Button(btns_right, text="Sửa", command=self.edit_game).pack(side="left", padx=6)
        ttk.Button(btns_right, text="Xoá", command=self.del_game).pack(side="left", padx=6)

        # Bottom area
        bottom = ttk.Panedwindow(self, orient="vertical")
        bottom.pack(fill="both", expand=True, padx=10, pady=8)

        stats_frame = ttk.Labelframe(bottom, text="Thống kê (Chỉ theo Debit/Ghi nợ)")
        results_frame = ttk.Labelframe(bottom, text="Kết quả")
        log_frame = ttk.Labelframe(bottom, text="Log hệ thống")

        bottom.add(stats_frame, weight=0)
        bottom.add(results_frame, weight=3)
        bottom.add(log_frame, weight=1)

        # Stats compact layout
        self.stat_vars: Dict[str, tk.StringVar] = {}
        stat_items = [
            ("total_tx", "Tổng số lần GD"),
            ("total_money", "Tổng tiền chuyển đi"),
            ("start_date", "Ngày bắt đầu"),
            ("end_date", "Ngày kết thúc"),
            ("cnt_gt_5m", "Số GD >5M"),
            ("sum_gt_5m", "Tổng tiền >5M"),
            ("max_tx", "GD lớn nhất"),
            ("min_tx", "GD nhỏ nhất"),
            ("day_most_tx", "Ngày nhiều GD nhất"),
            ("day_least_tx", "Ngày ít GD nhất"),
            ("day_max_sum", "Ngày tổng GD lớn nhất"),
            ("day_most_gt_5m", "Ngày nhiều GD >5M nhất"),
            ("k1", "K1 (00-05)"),
            ("k2", "K2 (05-07, 23-00)"),
            ("k3", "K3 (07-08, 22-23)"),
            ("k4", "K4 (08-22)"),
        ]

        for i, (k, label) in enumerate(stat_items):
            self.stat_vars[k] = tk.StringVar(value="0")
            r = i // 4
            c = (i % 4) * 2
            ttk.Label(stats_frame, text=label + ":").grid(row=r, column=c, sticky="e", padx=6, pady=2)
            ttk.Label(stats_frame, textvariable=self.stat_vars[k], foreground="#1f4e79").grid(row=r, column=c+1, sticky="w", padx=6, pady=2)
        for col in range(8):
            stats_frame.columnconfigure(col, weight=1)

        # ---------- Results toolbar (NEW) ----------
        rbar = ttk.Frame(results_frame)
        rbar.pack(fill="x", padx=6, pady=(6, 2))

        ttk.Label(rbar, text="Lọc:").pack(side="left", padx=(0, 4))
        self.cb_filter_col = ttk.Combobox(rbar, state="readonly", width=14, values=["(Tất cả)"])
        self.cb_filter_col.set("(Tất cả)")
        self.cb_filter_col.pack(side="left", padx=(0, 6))

        self.e_filter = ttk.Entry(rbar, width=26)
        self.e_filter.pack(side="left", padx=(0, 6))
        ttk.Button(rbar, text="Áp dụng", command=self.apply_filter).pack(side="left", padx=(0, 6))
        ttk.Button(rbar, text="Reset", command=self.reset_filter).pack(side="left", padx=(0, 10))

        ttk.Separator(rbar, orient="vertical").pack(side="left", fill="y", padx=6)

        ttk.Label(rbar, text="Tìm:").pack(side="left", padx=(6, 4))
        self.e_search = ttk.Entry(rbar, width=24)
        self.e_search.pack(side="left", padx=(0, 6))
        ttk.Button(rbar, text="Tìm tiếp", command=self.search_next).pack(side="left", padx=(0, 10))

        ttk.Button(rbar, text="Copy Selected", command=self.copy_selected).pack(side="right", padx=(6, 0))
        ttk.Button(rbar, text="Copy All", command=self.copy_all_view).pack(side="right", padx=(6, 0))

        # Results table
        cols = ("Ngân hàng", "Game", "Ngày", "Giờ", "Tiền đi", "Nội dung")
        self.result_cols = cols
        self.result_tree = ttk.Treeview(results_frame, columns=cols, show="headings")
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.result_tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.result_tree.pack(fill="both", expand=True, side="left", padx=(6, 0), pady=(0, 6))
        vsb.pack(fill="y", side="right", pady=(0, 6))
        hsb.pack(fill="x", side="bottom")

        for c in cols:
            # click header to sort
            self.result_tree.heading(c, text=c, command=lambda col=c: self.sort_by(col))
        self.result_tree.column("Ngân hàng", width=120, stretch=False)
        self.result_tree.column("Game", width=120, stretch=False)
        self.result_tree.column("Ngày", width=100, anchor="center", stretch=False)
        self.result_tree.column("Giờ", width=70, anchor="center", stretch=False)
        self.result_tree.column("Tiền đi", width=120, anchor="e", stretch=False)
        self.result_tree.column("Nội dung", width=900, stretch=True)

        # copy shortcut
        self.result_tree.bind("<Control-c>", lambda e: self.copy_selected())
        self.result_tree.bind("<Control-C>", lambda e: self.copy_selected())

        # Log widget
        self.log_text = tk.Text(log_frame, height=7, wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)
        self.log_text.tag_config("success", foreground="green")
        self.log_text.tag_config("warning", foreground="#e67e22")
        self.log_text.tag_config("error", foreground="red")
        self.log_text.tag_config("info", foreground="#2980b9")

        # Populate trees
        self._refresh_bank_tree()
        self._refresh_game_tree()

    def _show_cell_fulltext(self, tree: ttk.Treeview, event, title: str = "Chi tiết"):
        row = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not row or not col:
            return
        col_idx = int(col.replace("#", "")) - 1
        vals = tree.item(row, "values")
        if not vals or col_idx >= len(vals):
            return
        txt = str(vals[col_idx])

        w = tk.Toplevel(self)
        w.title(title)
        w.geometry("720x260")
        w.grab_set()
        t = tk.Text(w, wrap="word")
        t.pack(fill="both", expand=True, padx=8, pady=8)
        t.insert("1.0", txt)
        t.config(state="disabled")
        ttk.Button(w, text="Đóng", command=w.destroy).pack(pady=6)

    def _ensure_log_file(self):
        if self._log_file_path:
            return
        out_dir = os.path.join(BASE_DIR, "KETQUA")
        os.makedirs(out_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        self._log_file_path = os.path.join(out_dir, f"log_phan_tich_{ts}.txt")

    def log(self, msg: str, level: str = "info"):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.log_text.insert("end", line, level)
        self.log_text.see("end")
        try:
            self._ensure_log_file()
            with open(self._log_file_path, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass

    # ---------------- BANK UI HANDLERS ----------------
    def _refresh_bank_tree(self):
        for i in self.bank_tree.get_children():
            self.bank_tree.delete(i)
        for name, cfg in self.bank_cfg.items():
            self.bank_tree.insert(
                "", "end",
                values=("☐", name,
                        ", ".join(cfg.get("identify_cols", [])),
                        ", ".join(cfg.get("money_cols", [])),
                        ", ".join(cfg.get("date_cols", [])))
            )

    def _selected_bank_name(self) -> Optional[str]:
        for iid in self.bank_tree.get_children():
            vals = self.bank_tree.item(iid, "values")
            if vals and vals[0] == "☑":
                return vals[1]
        return None

    def _on_bank_click(self, event):
        region = self.bank_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.bank_tree.identify_column(event.x)
        row = self.bank_tree.identify_row(event.y)
        if not row:
            return
        if col == "#1":
            for iid in self.bank_tree.get_children():
                v = list(self.bank_tree.item(iid, "values"))
                v[0] = "☐"
                self.bank_tree.item(iid, values=v)
            v = list(self.bank_tree.item(row, "values"))
            v[0] = "☑"
            self.bank_tree.item(row, values=v)

    def add_bank(self):
        self._bank_popup(title="Thêm ngân hàng")

    def edit_bank(self):
        sel = self.bank_tree.selection()
        if not sel:
            messagebox.showwarning("Thiếu chọn", "Chọn 1 dòng ngân hàng để sửa")
            return
        vals = self.bank_tree.item(sel[0], "values")
        self._bank_popup(title="Sửa ngân hàng", current_name=vals[1])

    def del_bank(self):
        sel = self.bank_tree.selection()
        if not sel:
            messagebox.showwarning("Thiếu chọn", "Chọn 1 dòng ngân hàng để xoá")
            return
        name = self.bank_tree.item(sel[0], "values")[1]
        if messagebox.askyesno("Xác nhận", f"Xoá cấu hình ngân hàng: {name}?"):
            self.bank_cfg.pop(name, None)
            self._save_json(self.BANK_CFG_FILE, self.bank_cfg)
            self._refresh_bank_tree()

    def _bank_popup(self, title: str, current_name: Optional[str] = None):
        w = tk.Toplevel(self)
        w.title(title)
        w.geometry("660x280")
        w.grab_set()

        ttk.Label(w, text="Tên ngân hàng").grid(row=0, column=0, sticky="e", padx=10, pady=8)
        e_name = ttk.Entry(w, width=62)
        e_name.grid(row=0, column=1, padx=10, pady=8)

        ttk.Label(w, text="Identify columns (cách nhau ,)").grid(row=1, column=0, sticky="e", padx=10, pady=8)
        e_id = ttk.Entry(w, width=62)
        e_id.grid(row=1, column=1, padx=10, pady=8)

        ttk.Label(w, text="Debit columns (cách nhau ,)").grid(row=2, column=0, sticky="e", padx=10, pady=8)
        e_money = ttk.Entry(w, width=62)
        e_money.grid(row=2, column=1, padx=10, pady=8)

        ttk.Label(w, text="Date columns (cách nhau ,)").grid(row=3, column=0, sticky="e", padx=10, pady=8)
        e_date = ttk.Entry(w, width=62)
        e_date.grid(row=3, column=1, padx=10, pady=8)

        if current_name and current_name in self.bank_cfg:
            cfg = self.bank_cfg[current_name]
            e_name.insert(0, current_name)
            e_id.insert(0, ", ".join(cfg.get("identify_cols", [])))
            e_money.insert(0, ", ".join(cfg.get("money_cols", [])))
            e_date.insert(0, ", ".join(cfg.get("date_cols", [])))

        def save():
            name = e_name.get().strip()
            if not name:
                messagebox.showwarning("Thiếu dữ liệu", "Tên ngân hàng không được trống")
                return

            def split_csv(s):
                return [x.strip() for x in s.split(",") if x.strip()]

            cfg = {
                "identify_cols": split_csv(e_id.get()),
                "money_cols": split_csv(e_money.get()),
                "date_cols": split_csv(e_date.get())
            }
            if current_name and current_name != name:
                self.bank_cfg.pop(current_name, None)
            self.bank_cfg[name] = cfg
            self._save_json(self.BANK_CFG_FILE, self.bank_cfg)
            self._refresh_bank_tree()
            w.destroy()

        ttk.Button(w, text="Lưu", command=save).grid(row=4, column=0, columnspan=2, pady=18)

    # ---------------- GAME UI HANDLERS ----------------
    def _refresh_game_tree(self):
        for i in self.game_tree.get_children():
            self.game_tree.delete(i)
        for g in self.game_cfg:
            self.game_tree.insert(
                "", "end",
                values=("☑", g.get("name", ""), ", ".join(g.get("keywords", [])), g.get("match_type", "contains"))
            )

    def _on_game_click(self, event):
        region = self.game_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.game_tree.identify_column(event.x)
        row = self.game_tree.identify_row(event.y)
        if not row:
            return
        if col == "#1":
            v = list(self.game_tree.item(row, "values"))
            v[0] = "☑" if v[0] == "☐" else "☐"
            self.game_tree.item(row, values=v)

    def add_game(self):
        self._game_popup("Thêm game")

    def edit_game(self):
        sel = self.game_tree.selection()
        if not sel:
            messagebox.showwarning("Thiếu chọn", "Chọn 1 dòng game để sửa")
            return
        idx = self.game_tree.index(sel[0])
        self._game_popup("Sửa game", index=idx)

    def del_game(self):
        sel = self.game_tree.selection()
        if not sel:
            messagebox.showwarning("Thiếu chọn", "Chọn 1 dòng game để xoá")
            return
        idx = self.game_tree.index(sel[0])
        gname = self.game_cfg[idx].get("name", "")
        if messagebox.askyesno("Xác nhận", f"Xoá game: {gname}?"):
            self.game_cfg.pop(idx)
            self._save_json(self.GAME_CFG_FILE, self.game_cfg)
            self._refresh_game_tree()

    def _game_popup(self, title: str, index: Optional[int] = None):
        w = tk.Toplevel(self)
        w.title(title)
        w.geometry("660x240")
        w.grab_set()

        ttk.Label(w, text="Tên game").grid(row=0, column=0, sticky="e", padx=10, pady=8)
        e_name = ttk.Entry(w, width=62)
        e_name.grid(row=0, column=1, padx=10, pady=8)

        ttk.Label(w, text="Keywords (cách nhau ,)").grid(row=1, column=0, sticky="e", padx=10, pady=8)
        e_kw = ttk.Entry(w, width=62)
        e_kw.grid(row=1, column=1, padx=10, pady=8)

        ttk.Label(w, text="Kiểu match").grid(row=2, column=0, sticky="e", padx=10, pady=8)
        cb = ttk.Combobox(w, values=["contains", "regex", "startswith", "endswith"], state="readonly", width=20)
        cb.grid(row=2, column=1, sticky="w", padx=10, pady=8)
        cb.set("contains")

        if index is not None and 0 <= index < len(self.game_cfg):
            g = self.game_cfg[index]
            e_name.insert(0, g.get("name", ""))
            e_kw.insert(0, ", ".join(g.get("keywords", [])))
            cb.set(g.get("match_type", "contains"))

        def save():
            name = e_name.get().strip()
            if not name:
                messagebox.showwarning("Thiếu dữ liệu", "Tên game không được trống")
                return
            keywords = [x.strip() for x in e_kw.get().split(",") if x.strip()]
            match_type = cb.get()
            item = {"name": name, "keywords": keywords, "match_type": match_type}
            if index is None:
                self.game_cfg.append(item)
            else:
                self.game_cfg[index] = item
            self._save_json(self.GAME_CFG_FILE, self.game_cfg)
            self._refresh_game_tree()
            w.destroy()

        ttk.Button(w, text="Lưu", command=save).grid(row=3, column=0, columnspan=2, pady=18)

    # ---------------- FILE / ANALYSIS ----------------
    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Chọn file sao kê (xlsx/csv)",
            filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")]
        )
        if not paths:
            return
        self.selected_files = list(paths)
        self.lbl_selected.config(text=f"Đã chọn: {len(self.selected_files)} file")
        self.mapping = None
        self.log(f"Đã chọn {len(self.selected_files)} file", "success")

    def _read_preview(self, file_path: str, nrows: int = 40) -> pd.DataFrame:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".csv":
            try:
                return pd.read_csv(file_path, header=None, nrows=nrows, encoding="utf-8")
            except Exception:
                return pd.read_csv(file_path, header=None, nrows=nrows, encoding="latin1")
        else:
            return pd.read_excel(file_path, header=None, nrows=nrows)

    def _detect_header_row(self, preview: pd.DataFrame, bank_cfg: Dict[str, Any]) -> int:
        keys = []
        for k in (bank_cfg.get("identify_cols", []) + bank_cfg.get("money_cols", []) + bank_cfg.get("date_cols", [])):
            if k and isinstance(k, str):
                keys.append(k.strip().lower())
        if not keys:
            return 0

        best_idx, best_score = 0, -1
        for idx in range(len(preview)):
            row_vals = preview.iloc[idx].astype(str).fillna("").tolist()
            row_text = " | ".join([s.lower() for s in row_vals])
            score = sum(1 for k in keys if k in row_text)
            alpha_cells = sum(1 for v in row_vals if any(ch.isalpha() for ch in str(v)))
            score += 1 if alpha_cells >= 3 else 0
            if score > best_score:
                best_idx, best_score = idx, score

        return best_idx if best_score >= 1 else 0

    def _read_with_header(self, file_path: str, header_row: int) -> pd.DataFrame:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".csv":
            try:
                df = pd.read_csv(file_path, header=header_row, encoding="utf-8")
            except Exception:
                df = pd.read_csv(file_path, header=header_row, encoding="latin1")
        else:
            df = pd.read_excel(file_path, header=header_row)
        return df

    def _normalize_df(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame()
        drop_cols = [c for c in df.columns if "unnamed" in str(c).lower() or "nguồn file" in str(c).lower()]
        if drop_cols:
            df = df.drop(columns=drop_cols, errors="ignore")
        df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]
        df = df.dropna(how="all")
        df = df.reset_index(drop=True)

        def row_join(row):
            return " ".join([str(v).lower() for v in row.values if str(v).lower() != "nan"])

        def is_total_row(row):
            joined = row_join(row)
            return any(k in joined for k in ["tổng", "tong", "total", "grand total", "subtotal", "số dư", "so du", "balance"])

        df = df[~df.apply(is_total_row, axis=1)]

        colset = set([str(c).strip().lower() for c in df.columns])

        def is_header_repeat(row):
            vals = [str(v).strip().lower() for v in row.values if str(v).strip().lower() not in ("", "nan")]
            hit = sum(1 for v in vals if v in colset)
            return hit >= max(2, len(colset)//2)

        df = df[~df.apply(is_header_repeat, axis=1)].reset_index(drop=True)
        return df

    def _best_col(self, columns: List[str], candidates: List[str]) -> Optional[str]:
        cols_l = [c.lower() for c in columns]
        for cand in candidates:
            c = cand.lower()
            for orig, low in zip(columns, cols_l):
                if c in low:
                    return orig
        return None

    def _parse_money(self, v) -> Optional[float]:
        if v is None:
            return None
        s = str(v).strip()
        if not s or s.lower() == "nan":
            return None
        s = re.sub(r"[^0-9,.\-]", "", s)
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "")
                s = s.replace(",", ".")
            else:
                s = s.replace(",", "")
        else:
            if "," in s and "." not in s:
                s = s.replace(",", "")
        try:
            return float(s)
        except Exception:
            return None

    def _extract_datetime(self, raw) -> Tuple[Optional[pd.Timestamp], str, str]:
        s = "" if raw is None else str(raw)
        s = s.strip()
        if not s or s.lower() == "nan":
            return None, "", ""

        m = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", s)
        if not m:
            m = re.search(r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})", s)
        date_part = m.group(1) if m else ""
        time_part = ""
        tm = re.search(r"(\d{1,2}:\d{2}(?::\d{2})?)", s)
        if tm:
            time_part = tm.group(1)

        if date_part:
            try:
                dt = pd.to_datetime(f"{date_part} {time_part}".strip(), errors="coerce", dayfirst=True)
            except Exception:
                dt = pd.NaT
            if pd.isna(dt):
                return None, "", ""
            date_str = dt.strftime("%d/%m/%Y")
            time_str = dt.strftime("%H:%M:%S") if time_part else ""
            return dt, date_str, time_str

        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return None, "", ""
        return dt, dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M:%S") if (dt.hour or dt.minute or dt.second) else ""

    def _match_game(self, text: str, game_list: List[Dict[str, Any]]) -> Optional[str]:
        s = (text or "").strip()
        if not s:
            return None
        for g in game_list:
            if not g.get("_enabled", True):
                continue
            name = g.get("name", "")
            match_type = g.get("match_type", "contains")
            for kw in g.get("keywords", []) or []:
                if kw is None:
                    continue
                kw = str(kw)
                if kw == ".XXXXXXXX." and match_type == "regex":
                    for code in re.findall(r"\.([A-Z0-9]{8})\.", s):
                        if re.search(r"[A-Z]", code) and re.search(r"[0-9]", code):
                            return name
                    continue
                if kw == "CK XXXXXXXX" and match_type == "regex":
                    if re.search(r"CK\s+[A-Z0-9]{8}", s, flags=re.IGNORECASE):
                        return name
                    continue

                if match_type == "contains":
                    if kw.lower() in s.lower():
                        return name
                elif match_type == "startswith":
                    if s.lower().startswith(kw.lower()):
                        return name
                elif match_type == "endswith":
                    if s.lower().endswith(kw.lower()):
                        return name
                elif match_type == "regex":
                    try:
                        if re.search(kw, s, flags=re.IGNORECASE):
                            return name
                    except re.error:
                        continue
        return None

    def _time_bucket(self, hour: int) -> str:
        if 0 <= hour < 5:
            return "K1"
        if 5 <= hour < 7 or 23 <= hour < 24:
            return "K2"
        if 7 <= hour < 8 or 22 <= hour < 23:
            return "K3"
        if 8 <= hour < 22:
            return "K4"
        return "K4"

    def detect_and_confirm(self):
        if not self.selected_files:
            messagebox.showwarning("Thiếu file", "Vui lòng chọn file sao kê trước")
            return
        bank_name = self._selected_bank_name()
        if not bank_name:
            messagebox.showwarning("Thiếu bank", "Vui lòng tick 1 ngân hàng để dò")
            return

        bank_cfg = self.bank_cfg.get(bank_name, {})
        first = self.selected_files[0]
        try:
            preview = self._read_preview(first, nrows=40)
            header_row = self._detect_header_row(preview, bank_cfg)
            df_tmp = self._read_with_header(first, header_row)
            df_tmp = self._normalize_df(df_tmp)
        except Exception as e:
            self.log(f"Lỗi đọc file mẫu: {e}", "error")
            messagebox.showerror("Lỗi", f"Không đọc được file: {os.path.basename(first)}")
            return

        cols = list(df_tmp.columns)
        guess_date = self._best_col(cols, bank_cfg.get("date_cols", []))
        guess_money = self._best_col(cols, bank_cfg.get("money_cols", []))
        guess_content = self._best_col(cols, bank_cfg.get("identify_cols", []))

        mapping = self._confirm_mapping_dialog(cols, header_row, guess_date, guess_money, guess_content)
        if not mapping:
            self.log("Huỷ dò/mapping (người dùng không xác nhận)", "warning")
            return

        self.mapping = mapping
        self.log(
            f"Đã lưu mapping chung: header={mapping['header_row']} | date={mapping['date_col']} | "
            f"debit={mapping['money_col']} | content={mapping['content_col']}",
            "success"
        )

    def _confirm_mapping_dialog(
        self,
        df_cols: List[str],
        header_row: int,
        guess_date: Optional[str],
        guess_money: Optional[str],
        guess_content: Optional[str]
    ) -> Optional[Dict[str, Any]]:
        w = tk.Toplevel(self)
        w.title("Xác nhận header & cột dữ liệu")
        w.geometry("760x360")
        w.grab_set()

        ttk.Label(w, text="Header row (0-based):").grid(row=0, column=0, sticky="e", padx=10, pady=10)
        var_header = tk.IntVar(value=header_row)
        sp = ttk.Spinbox(w, from_=0, to=200, textvariable=var_header, width=10)
        sp.grid(row=0, column=1, sticky="w", padx=10, pady=10)

        ttk.Label(w, text="Cột Ngày/Thời gian:").grid(row=1, column=0, sticky="e", padx=10, pady=8)
        cb_date = ttk.Combobox(w, values=df_cols, state="readonly", width=58)
        cb_date.grid(row=1, column=1, padx=10, pady=8, sticky="w")
        if guess_date in df_cols:
            cb_date.set(guess_date)
        elif df_cols:
            cb_date.set(df_cols[0])

        ttk.Label(w, text="Cột Tiền đi (Debit/Ghi nợ):").grid(row=2, column=0, sticky="e", padx=10, pady=8)
        cb_money = ttk.Combobox(w, values=df_cols, state="readonly", width=58)
        cb_money.grid(row=2, column=1, padx=10, pady=8, sticky="w")
        if guess_money in df_cols:
            cb_money.set(guess_money)
        elif df_cols:
            cb_money.set(df_cols[0])

        ttk.Label(w, text="Cột Nội dung:").grid(row=3, column=0, sticky="e", padx=10, pady=8)
        cb_content = ttk.Combobox(w, values=df_cols, state="readonly", width=58)
        cb_content.grid(row=3, column=1, padx=10, pady=8, sticky="w")
        if guess_content in df_cols:
            cb_content.set(guess_content)
        elif df_cols:
            cb_content.set(df_cols[0])

        var_extract = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            w,
            text="Date column có lẫn ký tự thừa: bóc tách ngày/giờ bằng regex (khuyến nghị)",
            variable=var_extract
        ).grid(row=4, column=0, columnspan=2, sticky="w", padx=10, pady=10)

        var_merge_fill = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            w,
            text="Tự bù ô trống do merge (forward-fill) [mặc định OFF]",
            variable=var_merge_fill
        ).grid(row=5, column=0, columnspan=2, sticky="w", padx=10, pady=6)

        result: Dict[str, Any] = {}

        def ok():
            result["header_row"] = int(var_header.get())
            result["date_col"] = cb_date.get()
            result["money_col"] = cb_money.get()
            result["content_col"] = cb_content.get()
            result["use_regex_extract"] = bool(var_extract.get())
            result["merge_fill"] = bool(var_merge_fill.get())
            w.destroy()

        def cancel():
            result.clear()
            w.destroy()

        bf = ttk.Frame(w)
        bf.grid(row=6, column=0, columnspan=2, pady=16)
        ttk.Button(bf, text="OK", command=ok).pack(side="left", padx=10)
        ttk.Button(bf, text="Huỷ", command=cancel).pack(side="left", padx=10)

        w.wait_window()
        return result if result else None

    def analyze(self):
        if not self.selected_files:
            messagebox.showwarning("Thiếu file", "Vui lòng chọn file sao kê trước")
            return
        bank_name = self._selected_bank_name()
        if not bank_name:
            messagebox.showwarning("Thiếu bank", "Vui lòng tick 1 ngân hàng để phân tích")
            return

        enabled_games: List[Dict[str, Any]] = []
        for iid in self.game_tree.get_children():
            v = self.game_tree.item(iid, "values")
            gname = v[1]
            enabled = (v[0] == "☑")
            gcfg = next((g for g in self.game_cfg if g.get("name") == gname), None)
            if not gcfg:
                continue
            g = dict(gcfg)
            g["_enabled"] = enabled
            enabled_games.append(g)
        enabled_games = [g for g in enabled_games if g.get("_enabled")]

        if not enabled_games:
            messagebox.showwarning("Thiếu game", "Vui lòng tick ít nhất 1 cổng game để phân tích")
            return

        self.log(f"Bank: {bank_name} | Game tick: {', '.join([g['name'] for g in enabled_games])}", "info")

        if not self.mapping:
            self.detect_and_confirm()
            if not self.mapping:
                return

        header_row = self.mapping["header_row"]
        date_col = self.mapping["date_col"]
        money_col = self.mapping["money_col"]
        content_col = self.mapping["content_col"]
        use_regex_extract = bool(self.mapping.get("use_regex_extract", True))
        merge_fill = bool(self.mapping.get("merge_fill", False))

        frames = []
        for p in self.selected_files:
            try:
                df = self._read_with_header(p, header_row)
                df = self._normalize_df(df)
                if df.empty:
                    self.log(f"File rỗng: {os.path.basename(p)}", "warning")
                    continue

                if merge_fill:
                    df = df.fillna(method="ffill")

                if content_col not in df.columns:
                    self.log(f"Thiếu cột nội dung '{content_col}' trong {os.path.basename(p)} -> bỏ qua file", "error")
                    continue

                df["_source_file"] = os.path.basename(p)
                frames.append(df)
            except Exception as e:
                self.log(f"Lỗi đọc {os.path.basename(p)}: {e}", "error")

        if not frames:
            messagebox.showinfo("Kết quả", "Không có dữ liệu hợp lệ để phân tích.")
            self.log("Không có dữ liệu hợp lệ", "warning")
            return

        df_all = pd.concat(frames, ignore_index=True)
        self.df_all = df_all

        # Extract date/time
        if date_col in df_all.columns:
            if use_regex_extract:
                parsed = df_all[date_col].apply(self._extract_datetime)
            else:
                dts = pd.to_datetime(df_all[date_col], errors="coerce", dayfirst=True)
                parsed = dts.apply(lambda dt: (
                    dt if not pd.isna(dt) else None,
                    dt.strftime("%d/%m/%Y") if not pd.isna(dt) else "",
                    dt.strftime("%H:%M:%S") if (not pd.isna(dt) and (dt.hour or dt.minute or dt.second)) else ""
                ))
            df_all["_dt"] = parsed.apply(lambda x: x[0])
            df_all["_date"] = parsed.apply(lambda x: x[1])
            df_all["_time"] = parsed.apply(lambda x: x[2])
        else:
            df_all["_dt"] = None
            df_all["_date"] = ""
            df_all["_time"] = ""

        # Parse Debit
        if money_col in df_all.columns:
            df_all["_debit"] = df_all[money_col].apply(self._parse_money)
        else:
            df_all["_debit"] = None
            self.log(f"Không tìm thấy cột Debit '{money_col}' ở một số file. Dòng match sẽ giữ lại nhưng không tính tiền.", "warning")

        df_all["_content"] = df_all[content_col].astype(str)
        df_all["_game"] = df_all["_content"].apply(lambda s: self._match_game(s, enabled_games))

        df_f = df_all[df_all["_game"].notna()].copy()
        if df_f.empty:
            self.df_filtered = pd.DataFrame()
            self._prepare_view_after_analyze()
            self._render_results()
            self._compute_stats()
            self._render_stats()
            self.log("Không tìm thấy giao dịch match game", "warning")
            messagebox.showinfo("Kết quả", "Không tìm thấy giao dịch match theo game.")
            return

        df_f["Ngân hàng"] = bank_name
        df_f["Game"] = df_f["_game"]
        df_f["Ngày"] = df_f["_date"]
        df_f["Giờ"] = df_f["_time"].fillna("")
        df_f["Tiền đi"] = df_f["_debit"]
        df_f["Nội dung"] = df_f["_content"]

        df_f["_dt_sort"] = pd.to_datetime(
            (df_f["Ngày"].fillna("") + " " + df_f["Giờ"].replace("", "00:00:00")).str.strip(),
            errors="coerce", dayfirst=True
        )
        df_f = df_f.sort_values(by=["_dt_sort"], na_position="last").drop(columns=["_dt_sort"])

        self.df_filtered = df_f[["Ngân hàng", "Game", "Ngày", "Giờ", "Tiền đi", "Nội dung", "_source_file"]].reset_index(drop=True)

        # NEW: reset view/filter/sort/search after analyze
        self._prepare_view_after_analyze()
        self._render_results()
        self._compute_stats()
        self._render_stats()

        self.log(f"Hoàn tất phân tích: {len(self.df_filtered)} giao dịch match game", "success")

    # ---------------- RESULTS: sort / filter / search / copy (NEW) ----------------
    def _prepare_view_after_analyze(self):
        # reset view = full results
        self.df_view = None if self.df_filtered is None else self.df_filtered.copy()

        # reset filter UI
        self._search_cache = []
        self._search_pos = -1
        self._sort_state = {}

        # refresh filter columns combobox
        cols = ["(Tất cả)"] + list(self.result_cols)
        try:
            self.cb_filter_col.configure(values=cols)
            if self.cb_filter_col.get() not in cols:
                self.cb_filter_col.set("(Tất cả)")
        except Exception:
            pass

    def apply_filter(self):
        if self.df_filtered is None or self.df_filtered.empty:
            return

        q = (self.e_filter.get() or "").strip()
        col = (self.cb_filter_col.get() or "(Tất cả)").strip()

        # Nếu query rỗng => bỏ lọc
        if not q:
            self.df_view = self.df_filtered.copy()
            self._render_results()
            self.log("Đã bỏ lọc (query rỗng)", "info")
            return

        q_low = q.lower()
        df = self.df_filtered.copy()

        if col == "(Tất cả)":
            # IMPORTANT: pandas cần LIST để chọn nhiều cột, không dùng tuple trực tiếp
            cols = [c for c in list(self.result_cols) if c in df.columns]
            if not cols:
                self.log("Lọc thất bại: không có cột hợp lệ để lọc", "error")
                return

            mask = df[cols].astype(str).apply(
                lambda row: any(q_low in str(v).lower() for v in row.values),
                axis=1
            )
        else:
            if col not in df.columns:
                self.log(f"Lọc thất bại: không có cột {col}", "warning")
                return
            mask = df[col].astype(str).str.lower().str.contains(re.escape(q_low), na=False)

        self.df_view = df.loc[mask].copy().reset_index(drop=True)
        self._render_results()
        self.log(f"Lọc: {col} chứa '{q}' -> {len(self.df_view)} dòng", "success")

    def reset_filter(self):
        if self.df_filtered is None:
            return
        self.e_filter.delete(0, "end")
        self.cb_filter_col.set("(Tất cả)")
        self.df_view = self.df_filtered.copy()
        self._search_cache = []
        self._search_pos = -1
        self._render_results()
        self.log("Đã reset lọc", "info")

    def _dt_key_for_row(self, row: pd.Series):
        """
        Dùng cho sort Ngày/Giờ (dd/mm/yyyy + HH:MM[:SS]).
        """
        d = str(row.get("Ngày", "") or "").strip()
        t = str(row.get("Giờ", "") or "").strip()
        s = (d + " " + (t if t else "00:00:00")).strip()
        try:
            return pd.to_datetime(s, errors="coerce", dayfirst=True)
        except Exception:
            return pd.NaT

    def sort_by(self, col: str):
        if self.df_view is None or self.df_view.empty:
            return
        asc = self._sort_state.get(col, True)
        df = self.df_view.copy()

        try:
            if col in ("Ngày", "Giờ"):
                # sort by combined datetime for stable result
                df["_k"] = df.apply(self._dt_key_for_row, axis=1)
                df = df.sort_values(by=["_k"], ascending=asc, na_position="last").drop(columns=["_k"])
            elif col == "Tiền đi":
                df["_k"] = df["Tiền đi"].apply(lambda x: float(x) if (x is not None and not pd.isna(x)) else float("-inf"))
                df = df.sort_values(by=["_k"], ascending=asc, na_position="last").drop(columns=["_k"])
            else:
                df["_k"] = df[col].astype(str).str.lower()
                df = df.sort_values(by=["_k"], ascending=asc, na_position="last").drop(columns=["_k"])
        except Exception as e:
            self.log(f"Sort lỗi: {e}", "warning")
            return

        self.df_view = df.reset_index(drop=True)
        self._sort_state[col] = not asc
        self._render_results()
        self.log(f"Sort: {col} ({'tăng' if asc else 'giảm'})", "info")

    def search_next(self):
        """
        Search không lọc dữ liệu: chỉ highlight + nhảy đến dòng match tiếp theo trong view hiện tại.
        """
        if self.df_view is None or self.df_view.empty:
            return
        q = (self.e_search.get() or "").strip()
        if not q:
            return
        q_low = q.lower()

        # build cache if query changed or cache empty
        if not self._search_cache:
            # scan current tree items (fast enough for displayed rows)
            self._search_cache = []
            for iid in self.result_tree.get_children():
                vals = self.result_tree.item(iid, "values")
                joined = " | ".join([str(v) for v in vals]).lower()
                if q_low in joined:
                    self._search_cache.append(iid)
            self._search_pos = -1

        if not self._search_cache:
            self.log(f"Không tìm thấy: '{q}'", "warning")
            messagebox.showinfo("Tìm kiếm", f"Không tìm thấy: {q}")
            return

        self._search_pos = (self._search_pos + 1) % len(self._search_cache)
        iid = self._search_cache[self._search_pos]
        try:
            self.result_tree.selection_set(iid)
            self.result_tree.focus(iid)
            self.result_tree.see(iid)
        except Exception:
            pass

    def _to_tsv(self, df: pd.DataFrame, cols: Tuple[str, ...]) -> str:
        header = "\t".join(cols)
        lines = [header]
        for _, r in df.iterrows():
            row = []
            for c in cols:
                v = r.get(c, "")
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    s = ""
                elif c == "Tiền đi":
                    try:
                        s = f"{float(v):,.0f}".replace(",", ".")
                    except Exception:
                        s = str(v)
                else:
                    s = str(v)
                s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
                row.append(s)
            lines.append("\t".join(row))
        return "\n".join(lines)

    def copy_selected(self):
        """
        Copy các dòng đang chọn. Nếu không chọn gì -> copy dòng đang focus (nếu có).
        """
        if self.df_view is None or self.df_view.empty:
            return

        sels = list(self.result_tree.selection() or [])
        if not sels:
            f = self.result_tree.focus()
            if f:
                sels = [f]

        if not sels:
            self.log("Copy: chưa chọn dòng nào", "warning")
            return

        # Map selected rows by visible order (Tree doesn't store index -> rebuild from current displayed rows)
        displayed_iids = list(self.result_tree.get_children())
        idxs = []
        for iid in sels:
            try:
                idxs.append(displayed_iids.index(iid))
            except ValueError:
                continue

        if not idxs:
            self.log("Copy: không lấy được index dòng", "warning")
            return

        df_disp = self._get_display_df()
        df_sel = df_disp.iloc[idxs].copy()

        text = self._to_tsv(df_sel, self.result_cols)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.log(f"Đã copy {len(df_sel)} dòng (Ctrl+V để dán)", "success")

    def copy_all_view(self):
        if self.df_view is None or self.df_view.empty:
            return
        df_disp = self._get_display_df()
        text = self._to_tsv(df_disp, self.result_cols)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.log(f"Đã copy ALL {len(df_disp)} dòng đang hiển thị (Ctrl+V để dán)", "success")

    def _get_display_df(self) -> pd.DataFrame:
        """
        DataFrame tương ứng với những gì UI đang render (có giới hạn 5000 dòng).
        """
        df = self.df_view.copy()
        limit = 5000
        if len(df) > limit:
            df = df.head(limit)
        return df.reset_index(drop=True)

    def _render_results(self):
        for i in self.result_tree.get_children():
            self.result_tree.delete(i)

        if self.df_view is None or self.df_view.empty:
            # still refresh search cache
            self._search_cache = []
            self._search_pos = -1
            return

        df = self._get_display_df()

        # reset search cache (because view changed)
        self._search_cache = []
        self._search_pos = -1

        # Insert rows
        for _, r in df.iterrows():
            money = r.get("Tiền đi", None)
            money_str = "" if (money is None or (isinstance(money, float) and pd.isna(money))) else f"{float(money):,.0f}".replace(",", ".")
            self.result_tree.insert("", "end", values=(
                r.get("Ngân hàng", ""),
                r.get("Game", ""),
                r.get("Ngày", ""),
                r.get("Giờ", ""),
                money_str,
                r.get("Nội dung", ""),
            ))

        # Update filter combobox values if needed
        cols = ["(Tất cả)"] + list(self.result_cols)
        try:
            self.cb_filter_col.configure(values=cols)
        except Exception:
            pass

        # info for large results
        if self.df_view is not None and len(self.df_view) > len(df):
            self.log(f"Kết quả {len(self.df_view)} dòng. UI hiển thị {len(df)} dòng đầu, xem đầy đủ trong Excel/PDF.", "warning")

    # ---------------- STATS ----------------
    def _compute_stats(self):
        df = self.df_filtered
        if df is None or df.empty:
            self.stats = {}
            return

        df_out = df[df["Tiền đi"].notna()].copy()
        df_out = df_out[df_out["Tiền đi"] > 0]

        total_tx = int(len(df_out))
        total_money = float(df_out["Tiền đi"].sum()) if not df_out.empty else 0.0

        gt = df_out[df_out["Tiền đi"] > 5_000_000] if not df_out.empty else df_out
        cnt_gt_5m = int(len(gt))
        sum_gt_5m = float(gt["Tiền đi"].sum()) if not gt.empty else 0.0

        by_day_cnt = df_out.groupby("Ngày").size().sort_values(ascending=False) if not df_out.empty else pd.Series(dtype=int)
        day_most_tx = by_day_cnt.index[0] if not by_day_cnt.empty else ""

        by_day_sum = df_out.groupby("Ngày")["Tiền đi"].sum().sort_values(ascending=False) if not df_out.empty else pd.Series(dtype=float)
        day_max_sum = by_day_sum.index[0] if not by_day_sum.empty else ""

        by_day_gt = gt.groupby("Ngày").size().sort_values(ascending=False) if not gt.empty else pd.Series(dtype=int)
        day_most_gt_5m = by_day_gt.index[0] if not by_day_gt.empty else ""

        max_tx = ""
        min_tx = ""
        if not df_out.empty:
            rmax = df_out.loc[df_out["Tiền đi"].idxmax()]
            rmin = df_out.loc[df_out["Tiền đi"].idxmin()]
            max_tx = f"{float(rmax['Tiền đi']):,.0f}".replace(",", ".") + (f" ({rmax['Ngày']})" if rmax["Ngày"] else "")
            min_tx = f"{float(rmin['Tiền đi']):,.0f}".replace(",", ".") + (f" ({rmin['Ngày']})" if rmin["Ngày"] else "")

        buckets = {"K1": 0, "K2": 0, "K3": 0, "K4": 0}

        def hour_of(t):
            try:
                if not t:
                    return None
                return int(str(t).split(":")[0])
            except Exception:
                return None

        hours = df_out["Giờ"].apply(hour_of) if not df_out.empty else pd.Series(dtype=float)
        for h in hours.dropna():
            b = self._time_bucket(int(h))
            buckets[b] += 1

        start_date = ""
        end_date = ""
        day_least_tx = ""

        try:
            d = pd.to_datetime(df_out["Ngày"], errors="coerce", dayfirst=True)
            d = d.dropna()
            if not d.empty:
                start_date = d.min().strftime("%d/%m/%Y")
                end_date = d.max().strftime("%d/%m/%Y")
        except Exception:
            pass

        try:
            tmp = df_out[df_out["Ngày"].astype(str).str.strip() != ""]
            by_day_cnt_asc = tmp.groupby("Ngày").size().sort_values(ascending=True)
            day_least_tx = by_day_cnt_asc.index[0] if not by_day_cnt_asc.empty else ""
        except Exception:
            pass

        self.stats = dict(
            total_tx=total_tx,
            total_money=total_money,
            cnt_gt_5m=cnt_gt_5m,
            sum_gt_5m=sum_gt_5m,
            max_tx=max_tx,
            min_tx=min_tx,
            day_most_tx=day_most_tx,
            day_max_sum=day_max_sum,
            day_most_gt_5m=day_most_gt_5m,
            k1=buckets["K1"],
            k2=buckets["K2"],
            k3=buckets["K3"],
            k4=buckets["K4"],
            start_date=start_date,
            end_date=end_date,
            day_least_tx=day_least_tx,
        )

    def _render_stats(self):
        s = self.stats or {}
        self.stat_vars["total_tx"].set(str(s.get("total_tx", 0)))
        self.stat_vars["total_money"].set(f"{s.get('total_money', 0):,.0f}".replace(",", "."))
        self.stat_vars["cnt_gt_5m"].set(str(s.get("cnt_gt_5m", 0)))
        self.stat_vars["sum_gt_5m"].set(f"{s.get('sum_gt_5m', 0):,.0f}".replace(",", "."))
        self.stat_vars["max_tx"].set(str(s.get("max_tx", "")))
        self.stat_vars["min_tx"].set(str(s.get("min_tx", "")))
        self.stat_vars["day_most_tx"].set(str(s.get("day_most_tx", "")))
        self.stat_vars["day_max_sum"].set(str(s.get("day_max_sum", "")))
        self.stat_vars["day_most_gt_5m"].set(str(s.get("day_most_gt_5m", "")))
        self.stat_vars["k1"].set(str(s.get("k1", 0)))
        self.stat_vars["k2"].set(str(s.get("k2", 0)))
        self.stat_vars["k3"].set(str(s.get("k3", 0)))
        self.stat_vars["k4"].set(str(s.get("k4", 0)))
        self.stat_vars["start_date"].set(str(s.get("start_date", "")))
        self.stat_vars["end_date"].set(str(s.get("end_date", "")))
        self.stat_vars["day_least_tx"].set(str(s.get("day_least_tx", "")))

    # ---------------- EXPORT ----------------
    def _format_excel_workbook(self, file_path: str):
        try:
            wb = load_workbook(file_path)
            font_style = Font(name="Times New Roman", size=12)
            alignment_style = Alignment(horizontal="left", vertical="top", wrap_text=True)
            thin = Side(border_style="thin", color="000000")
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        cell.font = font_style
                        cell.alignment = alignment_style
                        cell.border = border
                for col in ws.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        v = "" if cell.value is None else str(cell.value)
                        if len(v) > max_len:
                            max_len = len(v)
                    ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 60)

            wb.save(file_path)
        except Exception as e:
            self.log(f"Lỗi format Excel: {e}", "warning")

    def export_excel(self):
        if self.df_filtered is None or self.df_filtered.empty:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng phân tích trước khi xuất Excel")
            return
        out_dir = os.path.join(BASE_DIR, "KETQUA")
        os.makedirs(out_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(out_dir, f"Bao_cao_Phan_tich_SK_TKNH_{ts}.xlsx")
        try:
            df = self.df_filtered.copy()
            s = self.stats or {}
            meta = {
                "Ngày xuất": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "Số file": len(self.selected_files),
                "Bank": self._selected_bank_name() or "",
                "Game tick": ", ".join([self.game_tree.item(i, "values")[1] for i in self.game_tree.get_children() if self.game_tree.item(i, "values")[0] == "☑"])
            }
            stat_rows = [
                ["Cột 1", "", "Cột 2", "", "Cột 3", "", "Cột 4", ""],
                ["Tổng số lần giao dịch", s.get("total_tx", 0),
                 "Tổng số lần giao dịch >5M", s.get("cnt_gt_5m", 0),
                 "Ngày nhiều giao dịch nhất", s.get("day_most_tx", ""),
                 "K1", s.get("k1", 0)],
                ["Tổng số tiền chuyển đi", s.get("total_money", 0),
                 "Tổng số tiền chuyển đi >5M", s.get("sum_gt_5m", 0),
                 "Ngày ít giao dịch nhất", s.get("day_least_tx", ""),
                 "K2", s.get("k2", 0)],
                ["Ngày bắt đầu", s.get("start_date", ""),
                 "Giao dịch lớn nhất", s.get("max_tx", ""),
                 "Ngày có tổng giao dịch lớn nhất", s.get("day_max_sum", ""),
                 "K3", s.get("k3", 0)],
                ["Ngày kết thúc", s.get("end_date", ""),
                 "Giao dịch nhỏ nhất", s.get("min_tx", ""),
                 "Ngày nhiều giao dịch >5M nhất", s.get("day_most_gt_5m", ""),
                 "K4", s.get("k4", 0)],
            ]
            df_stats = pd.DataFrame(stat_rows[1:], columns=stat_rows[0])
            df_meta = pd.DataFrame(list(meta.items()), columns=["Thông tin", "Giá trị"])

            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="KET_QUA", index=False)
                df_meta.to_excel(writer, sheet_name="THONG_KE", index=False, startrow=0)
                df_stats.to_excel(writer, sheet_name="THONG_KE", index=False, startrow=len(df_meta)+2)

            self._format_excel_workbook(out_path)
            self.log(f"Đã xuất Excel: {out_path}", "success")
            messagebox.showinfo("Thành công", f"Đã xuất Excel tại:\n{out_path}")
        except Exception as e:
            self.log(f"Lỗi xuất Excel: {e}", "error")
            messagebox.showerror("Lỗi", str(e))

    def export_pdf(self):
        if self.df_filtered is None or self.df_filtered.empty:
            messagebox.showwarning("Chưa có dữ liệu", "Vui lòng phân tích trước khi xuất PDF")
            return
        out_dir = os.path.join(BASE_DIR, "KETQUA")
        os.makedirs(out_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(out_dir, f"Bao_cao_Phan_tich_SK_TKNH_{ts}.pdf")

        try:
            styles = getSampleStyleSheet()
            base_font = self._pdf_font_name
            style_title = ParagraphStyle("title", parent=styles["Title"], fontName=base_font, fontSize=16, leading=18, alignment=1)
            style_h1 = ParagraphStyle("h1", parent=styles["Heading1"], fontName=base_font, fontSize=12, leading=14)
            style_p = ParagraphStyle("p", parent=styles["BodyText"], fontName=base_font, fontSize=10, leading=12)

            doc = SimpleDocTemplate(
                out_path,
                pagesize=A4,
                leftMargin=3*cm,
                rightMargin=2*cm,
                topMargin=2*cm,
                bottomMargin=2*cm
            )

            story = []
            story.append(Paragraph("BÁO CÁO PHÂN TÍCH SK TKNH", style_title))
            story.append(Spacer(1, 0.4*cm))

            s = self.stats or {}
            meta_lines = [
                f"Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                f"Số file: {len(self.selected_files)}",
                f"Bank: {self._selected_bank_name() or ''}",
            ]
            games_tick = ", ".join([self.game_tree.item(i, "values")[1] for i in self.game_tree.get_children() if self.game_tree.item(i, "values")[0] == "☑"])
            if games_tick:
                meta_lines.append(f"Game tick: {games_tick}")
            story.append(Paragraph("<br/>".join(meta_lines), style_p))
            story.append(Spacer(1, 0.3*cm))

            stat_rows = [
                ["Cột 1", "", "Cột 2", "", "Cột 3", "", "Cột 4", ""],
                ["Tổng số lần giao dịch", str(s.get("total_tx", 0)),
                 "Tổng số lần giao dịch >5M", str(s.get("cnt_gt_5m", 0)),
                 "Ngày nhiều giao dịch nhất", str(s.get("day_most_tx", "")),
                 "K1", str(s.get("k1", 0))],
                ["Tổng số tiền chuyển đi", f"{s.get('total_money', 0):,.0f}".replace(",", "."),
                 "Tổng số tiền chuyển đi >5M", f"{s.get('sum_gt_5m', 0):,.0f}".replace(",", "."),
                 "Ngày ít giao dịch nhất", str(s.get("day_least_tx", "")),
                 "K2", str(s.get("k2", 0))],
                ["Ngày bắt đầu", str(s.get("start_date", "")),
                 "Giao dịch lớn nhất", str(s.get("max_tx", "")),
                 "Ngày có tổng giao dịch lớn nhất", str(s.get("day_max_sum", "")),
                 "K3", str(s.get("k3", 0))],
                ["Ngày kết thúc", str(s.get("end_date", "")),
                 "Giao dịch nhỏ nhất", str(s.get("min_tx", "")),
                 "Ngày nhiều giao dịch >5M nhất", str(s.get("day_most_gt_5m", "")),
                 "K4", str(s.get("k4", 0))],
            ]
            stat_rows_p = [[Paragraph(str(c), style_p) for c in row] for row in stat_rows]

            usable_w = A4[0] - (3*cm) - (2*cm)
            col_w = [usable_w*0.17, usable_w*0.08,
                     usable_w*0.17, usable_w*0.08,
                     usable_w*0.17, usable_w*0.08,
                     usable_w*0.09, usable_w*0.06]
            t = Table(stat_rows_p, colWidths=col_w)
            t.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
                ("GRID", (0,0), (-1,-1), 0.5, colors.black),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("FONTNAME", (0,0), (-1,-1), base_font),
                ("FONTSIZE", (0,0), (-1,-1), 9),
                ("BOTTOMPADDING", (0,0), (-1,-1), 4),
                ("TOPPADDING", (0,0), (-1,-1), 4),
            ]))
            story.append(t)

            story.append(PageBreak())
            story.append(Paragraph("BẢNG KẾT QUẢ", style_h1))
            story.append(Spacer(1, 0.2*cm))

            df = self.df_filtered.copy().fillna("")
            header = ["Bank", "Game", "Ngày", "Giờ", "Tiền đi", "Nội dung"]
            data = [header]

            cell_style = ParagraphStyle("cell", parent=style_p, fontName=base_font, fontSize=7.5, leading=9)
            head_style = ParagraphStyle("head", parent=style_p, fontName=base_font, fontSize=8.2, leading=10)

            for _, r in df.iterrows():
                money = r["Tiền đi"]
                money_str = "" if (money == "" or pd.isna(money)) else f"{float(money):,.0f}".replace(",", ".")
                row = [
                    Paragraph(str(r["Ngân hàng"]), cell_style),
                    Paragraph(str(r["Game"]), cell_style),
                    Paragraph(str(r["Ngày"]), cell_style),
                    Paragraph(str(r["Giờ"]), cell_style),
                    Paragraph(money_str, cell_style),
                    Paragraph(str(r["Nội dung"]), cell_style),
                ]
                data.append(row)

            data[0] = [Paragraph(c, head_style) for c in header]

            col_w = [
                usable_w*0.12,
                usable_w*0.12,
                usable_w*0.10,
                usable_w*0.08,
                usable_w*0.14,
                usable_w*0.44,
            ]
            table = Table(data, colWidths=col_w, repeatRows=1)
            table.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
                ("GRID", (0,0), (-1,-1), 0.25, colors.black),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("FONTNAME", (0,0), (-1,-1), base_font),
                ("LEFTPADDING", (0,0), (-1,-1), 3),
                ("RIGHTPADDING", (0,0), (-1,-1), 3),
                ("TOPPADDING", (0,0), (-1,-1), 2),
                ("BOTTOMPADDING", (0,0), (-1,-1), 2),
            ]))
            story.append(table)

            doc.build(story)

            self.log(f"Đã xuất PDF: {out_path}", "success")
            messagebox.showinfo("Thành công", f"Đã xuất PDF tại:\n{out_path}")
        except Exception as e:
            self.log(f"Lỗi xuất PDF: {e}", "error")
            messagebox.showerror("Lỗi", str(e))
    
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hệ thống hỗ trợ nghiệp vụ v1")
        self.geometry("1300x820")

        self.template_dir = DEFAULT_TEMPLATE_DIR
        self.build_toolbar()
        self.build_ui()

    def build_toolbar(self):

        bar = tk.Frame(self, bg="#1f4e79", height=50)
        bar.pack(fill="x")

        btn_style = {
            "font": ("Segoe UI", 11, "bold"),
            "bd": 0,
            "padx": 15,
            "pady": 8,
            "cursor": "hand2"
        }

        tk.Button(
            bar,
            text="📂 MỞ HỒ SƠ",
            bg="#ffffff",
            fg="#1f4e79",
            command=self.load_profile
        ).pack(side="left", padx=10, pady=5)

        tk.Button(
            bar,
            text="📁 CHỌN MẪU & XUẤT HỐ SƠ",
            bg="#28a745",
            fg="white",
            command=self.choose_template_and_export
        ).pack(side="left", padx=10, pady=5)

        tk.Button(
            bar,
            text="📤 GỬI BÁO CÁO",
            bg="#dc3545",
            fg="white",
            command=self.send_report_placeholder
        ).pack(side="left", padx=10, pady=5)

        self.template_label = tk.Label(
            bar,
            text="Mẫu hiện tại: Mau 1",
            bg="#1f4e79",
            fg="white",
            font=("Segoe UI", 10)
        )
        self.template_label.pack(side="right", padx=20)
    
#    # ================= MENU =================
#    def build_menu(self):
#        m = tk.Menu(self)
#        fm = tk.Menu(m, tearoff=0)
#        fm.add_command(label="📂 Mở hồ sơ", command=self.load_profile)
#        fm.add_command(label="💾 Lưu hồ sơ", command=self.save_profile)
#        fm.add_separator()
#        fm.add_command(label="🖨 Xuất 7 file Word", command=self.export_word)
#        m.add_cascade(label="📁 HỒ SƠ", menu=fm)
#        self.config(menu=m)

    # ================= UI =================
    def build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self.tab_cb = ttk.Frame(nb)
        self.tab_dt = ttk.Frame(nb)
        self.tab_pt = TabPhanTichSK(nb)
        self.tab_ptcs = TabPhanTichSKChuyenSau(nb)
        
        nb.add(self.tab_cb, text="Cán bộ")
        nb.add(self.tab_dt, text="Đối tượng")
        nb.add(self.tab_pt, text="Phân tích SK TKNH")
        nb.add(self.tab_ptcs, text="Phân tích chuyên sâu SK TKNH")
        
        self.build_tab_can_bo()
        self.build_tab_doi_tuong()

    # ================= TAB CÁN BỘ =================
    def build_tab_can_bo(self):
        f = self.tab_cb

        top = ttk.Frame(f)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Ngày").pack(side="left")
        self.e_ngay = PlaceholderEntry(top, placeholder="dd/mm/yyyy")
        self.e_ngay.pack(side="left", fill="x", expand=True, padx=6)

        ttk.Label(top, text="Địa điểm").pack(side="left")
        self.e_dia_diem = PlaceholderEntry(top, placeholder="Địa điểm làm việc")
        self.e_dia_diem.pack(side="left", fill="x", expand=True, padx=6)

        cols = ("Họ và tên", "Chức danh/Chức vụ", "Đơn vị công tác")
        self.tv_cb = ttk.Treeview(f, columns=cols, show="headings")
        self.tv_cb.pack(fill="both", expand=True, padx=10, pady=8)

        for c in cols:
            self.tv_cb.heading(c, text=c)

        bf = ttk.Frame(f)
        bf.pack(pady=6)

        ttk.Button(bf, text="Thêm",
                   command=lambda: self.popup_table(
                       "Cán bộ",
                       ["Họ tên", "Chức vụ", "Đơn vị"],
                       self.tv_cb)).pack(side="left", padx=6)

        ttk.Button(bf, text="Sửa",
                   command=lambda: self.edit_row(
                       self.tv_cb,
                       "Cán bộ",
                       ["Họ tên", "Chức vụ", "Đơn vị"])).pack(side="left", padx=6)

        ttk.Button(bf, text="Xoá",
                   command=lambda: self.del_row(self.tv_cb)).pack(side="left", padx=6)

    # ================= TAB ĐỐI TƯỢNG =================
    def build_tab_doi_tuong(self):
        f = self.tab_dt
        self.entries = {}

        box = ttk.LabelFrame(f, text="Thông tin đối tượng")
        box.pack(fill="x", padx=10, pady=6)

        for i in range(8):
            box.columnconfigure(i, weight=1)

        def add_row(r, items):
            col = 0
            for key, label, default in items:
                ttk.Label(box, text=label).grid(row=r, column=col, sticky="e")
                e = PlaceholderEntry(box, placeholder=label)
                e.grid(row=r, column=col + 1, sticky="ew")
                if default:
                    e.delete(0, "end")
                    e.insert(0, default)
                    e.config(foreground="black")
                self.entries[key] = e
                col += 2

        add_row(0, [
            ("ho_dem", "Họ đệm", ""),
            ("ten", "Tên", ""),
            ("ngay_sinh", "Ngày sinh", ""),
            ("noi_sinh", "Nơi sinh", "")
        ])

        add_row(1, [
            ("gioi_tinh", "Giới tính", ""),
            ("dan_toc", "Dân tộc", "Kinh"),
            ("ton_giao", "Tôn giáo", "Không"),
            ("quoc_tich", "Quốc tịch", "Việt Nam")
        ])

        add_row(2, [
            ("nghe_nghiep", "Nghề nghiệp", ""),
            ("cccd", "CCCD/HC", ""),
            ("noi_cap", "Nơi cấp", "Cục CSQLHC về TTXH"),
            ("sdt", "SĐT", "")
        ])

        for key, label, row in [
            ("thuong_tru", "Thường trú", 3),
            ("tam_tru", "Tạm trú", 4),
            ("noi_o", "Nơi ở hiện nay", 5)
        ]:
            ttk.Label(box, text=label).grid(row=row, column=0, sticky="e")
            e = PlaceholderEntry(box, placeholder=label)
            e.grid(row=row, column=1, columnspan=7, sticky="ew")
            self.entries[key] = e

        self.tv_phone = self.build_sub_table(
            f, "Điện thoại",
            ["Dòng máy", "IMEI1", "IMEI2", "SĐT1", "SĐT2"])

        self.tv_bank = self.build_sub_table(
            f, "Tài khoản",
            ["Số tài khoản", "Chủ tài khoản", "Ngân hàng", "SĐT liên kết"])

    # ================= SUB TABLE =================
    def build_sub_table(self, parent, title, headers):
        box = ttk.LabelFrame(parent, text=title)
        box.pack(fill="both", expand=True, padx=10, pady=6)

        tv = ttk.Treeview(box, columns=headers, show="headings")
        tv.pack(fill="both", expand=True)

        for h in headers:
            tv.heading(h, text=h)

        bf = ttk.Frame(box)
        bf.pack()

        ttk.Button(bf, text="Thêm",
                   command=lambda: self.popup_table(
                       title, headers, tv)).pack(side="left", padx=6)
        ttk.Button(bf, text="Sửa",
                   command=lambda: self.edit_row(
                       tv, title, headers)).pack(side="left", padx=6)
        ttk.Button(bf, text="Xoá",
                   command=lambda: self.del_row(tv)).pack(side="left", padx=6)

        return tv

    # ================= POPUP =================
    def popup_table(self, title, fields, tree, item=None):
        w = tk.Toplevel(self)
        w.title(title)
        w.geometry("600x400")
        w.grab_set()

        entries = []
        for i, f in enumerate(fields):
            ttk.Label(w, text=f).grid(row=i, column=0, padx=10, pady=6)
            e = ttk.Entry(w, width=60)
            e.grid(row=i, column=1, padx=10, pady=6)
            entries.append(e)

        if item:
            vals = tree.item(item, "values")
            for e, v in zip(entries, vals):
                e.insert(0, v)

        def save():
            vals = [e.get() for e in entries]
            if item:
                tree.item(item, values=vals)
            else:
                tree.insert("", "end", values=vals)
            w.destroy()

        ttk.Button(w, text="Lưu", command=save).grid(
            row=len(fields), column=0, columnspan=2, pady=20)

    def edit_row(self, tree, title, fields):
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Thiếu chọn", "Chọn 1 dòng để sửa")
            return
        self.popup_table(title, fields, tree, sel[0])

    def del_row(self, tree):
        sel = tree.selection()
        if sel:
            tree.delete(sel[0])

    # ================= DATA =================
    def collect_data(self):
        return {
            "ngay": self.e_ngay.get_real(),
            "dia_diem": self.e_dia_diem.get_real(),
            "can_bo": [
                self.tv_cb.item(i, "values")
                for i in self.tv_cb.get_children()
            ],
            "doi_tuong": {
                k: (e.get_real() if hasattr(e, "get_real") else e.get())
                for k, e in self.entries.items()
            },
            "dien_thoai": [
                self.tv_phone.item(i, "values")
                for i in self.tv_phone.get_children()
            ],
            "tai_khoan": [
                self.tv_bank.item(i, "values")
                for i in self.tv_bank.get_children()
            ],
        }
    
    # ================= SAVE PROFILE =================
    def save_profile(self):
        data = self.collect_data()

        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=PROFILE_DIR
        )

        if not file_path:
            return

        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        messagebox.showinfo("Thông báo", "Đã lưu hồ sơ thành công.")

    def choose_template_and_export(self):

        folder = filedialog.askdirectory(
            title="Chọn thư mục mẫu",
            initialdir=BASE_DIR
        )

        if not folder:
            return

        self.template_dir = folder
        folder_name = os.path.basename(folder)
        self.template_label.config(text=f"Mẫu hiện tại: {folder_name}")

        # Tự động lưu hồ sơ trước khi xuất
        self.auto_save_profile()

        self.export_word_from_template()
    
    # ================= LOAD PROFILE =================
    def load_profile(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            initialdir=PROFILE_DIR
        )

        if not file_path:
            return

        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Ngày + địa điểm
        self.e_ngay.delete(0, "end")
        self.e_ngay.insert(0, data.get("ngay", ""))

        self.e_dia_diem.delete(0, "end")
        self.e_dia_diem.insert(0, data.get("dia_diem", ""))

        # Cán bộ
        for i in self.tv_cb.get_children():
            self.tv_cb.delete(i)

        for row in data.get("can_bo", []):
            self.tv_cb.insert("", "end", values=row)

        # Đối tượng
        for k, v in data.get("doi_tuong", {}).items():
            if k in self.entries:
                self.entries[k].delete(0, "end")
                self.entries[k].insert(0, v)

        # Điện thoại
        for i in self.tv_phone.get_children():
            self.tv_phone.delete(i)

        for row in data.get("dien_thoai", []):
            self.tv_phone.insert("", "end", values=row)

        # Tài khoản
        for i in self.tv_bank.get_children():
            self.tv_bank.delete(i)

        for row in data.get("tai_khoan", []):
            self.tv_bank.insert("", "end", values=row)

        messagebox.showinfo("Thông báo", "Đã mở hồ sơ thành công.")

    # ================= EXPORT WORD (FIX CHUẨN) =================
    def export_word_from_template(self):

        import unicodedata
        import re
        from datetime import datetime

        data = self.collect_data()

        dt = data.get("doi_tuong", {})
        cb = data.get("can_bo", [])
        phones = data.get("dien_thoai", [])
        banks = data.get("tai_khoan", [])

        # ================= HỌ TÊN =================
        ho_tendem = dt.get("ho_dem", "").strip()
        ten = dt.get("ten", "").strip()
        ho_ten_day_du = f"{ho_tendem} {ten}".strip()

        # ================= NGÀY SINH =================
        ngs = ts = nas = ""
        tuoi = ""
        ngay_sinh = dt.get("ngay_sinh", "").strip()

        if ngay_sinh:
            try:
                d, m, y = ngay_sinh.split("/")
                ngs, ts, nas = d, m, y
                tuoi = datetime.now().year - int(y)
            except:
                messagebox.showerror("Lỗi", "Ngày sinh phải dạng dd/mm/yyyy")
                return

        # ================= DANH XƯNG =================
        dx = ""
        gioi_tinh = dt.get("gioi_tinh", "").strip().lower()

        if gioi_tinh == "nam":
            dx = "Ông" if tuoi and tuoi >= 60 else "Anh"
        elif gioi_tinh in ["nữ", "nu"]:
            dx = "Bà" if tuoi and tuoi >= 60 else "Chị"

        # ================= NGÀY BIÊN BẢN =================
        ngay = thang = nam = ""
        bb_date = data.get("ngay", "").strip()

        if bb_date:
            try:
                d, m, y = bb_date.split("/")
                ngay, thang, nam = d, m, y
            except:
                messagebox.showerror("Lỗi", "Ngày biên bản phải dạng dd/mm/yyyy")
                return

        # ================= CÁN BỘ =================
        cb1 = cb[0] if len(cb) > 0 else ("", "", "")
        cb2 = cb[1] if len(cb) > 1 else ("", "", "")

        # ================= ĐIỆN THOẠI =================
        phone1 = phones[0] if len(phones) > 0 else ("", "", "", "", "")
        imei2 = phone1[2] if phone1[2] else "Không"
        sim2 = phone1[4] if phone1[4] else "Không"

        # ================= NGÂN HÀNG =================
        bank1 = banks[0] if len(banks) > 0 else ("", "", "", "")
        bank2 = banks[1] if len(banks) > 1 else ("", "", "", "")
        bank3 = banks[2] if len(banks) > 2 else ("", "", "", "")

        # ================= BIẾN =================
        variables = {

            "{{Hoten_CB1}}": cb1[0],
            "{{Chucvu_CB1}}": cb1[1],
            "{{Donvi_CB1}}": cb1[2],

            "{{Hoten_CB2}}": cb2[0],
            "{{Chucvu_CB2}}": cb2[1],
            "{{Donvi_CB2}}": cb2[2],

            "{{Ho_tendem_DT}}": ho_tendem,
            "{{Ten_DT}}": ten,
            "{{Ho_ten_day_du}}": ho_ten_day_du,

            "{{Gioitinh_DT}}": dt.get("gioi_tinh", ""),
            "{{QT_DT}}": dt.get("quoc_tich", ""),
            "{{DT_DT}}": dt.get("dan_toc", ""),
            "{{TG_DT}}": dt.get("ton_giao", ""),
            "{{NN_DT}}": dt.get("nghe_nghiep", ""),
            "{{SDT_DT}}": dt.get("sdt", ""),

            "{{NGS_DT}}": ngs,
            "{{TS_DT}}": ts,
            "{{NAS_DT}}": nas,
            "{{Tuoi_DT}}": tuoi,
            "{{NOS_DT}}": dt.get("noi_sinh", ""),

            "{{GT_DT}}": dt.get("cccd", ""),
            "{{NCGT_DT}}": dt.get("noi_cap", ""),
            "{{HK_DT}}": dt.get("thuong_tru", ""),
            "{{TT_DT}}": dt.get("tam_tru", ""),
            "{{NOHT_DT}}": dt.get("noi_o", ""),

            "{{STK1_DT}}": bank1[0],
            "{{CTK1_DT}}": bank1[1],
            "{{NH1_DT}}": bank1[2],
            "{{SDT1_DT}}": bank1[3],

            "{{STK2_DT}}": bank2[0],
            "{{CTK2_DT}}": bank2[1],
            "{{NH2_DT}}": bank2[2],
            "{{SDT2_DT}}": bank2[3],

            "{{STK3_DT}}": bank3[0],
            "{{CTK3_DT}}": bank3[1],
            "{{NH3_DT}}": bank3[2],
            "{{SDT3_DT}}": bank3[3],

            "{{DM_DT_DT}}": phone1[0],
            "{{IMEI1_DT}}": phone1[1],
            "{{IMEI2_DT}}": imei2,
            "{{SIM1_DT}}": phone1[3],
            "{{SIM2_DT}}": sim2,

            "{{DX_DT}}": dx,
            "{{dia_diem}}": data.get("dia_diem", ""),
            "{{ngay}}": ngay,
            "{{thang}}": thang,
            "{{nam}}": nam,
        }

        # ================= ENGINE REPLACE GIỮ FORMAT (BẢN v10 NÂNG CẤP) =================
        def replace_paragraph(paragraph):

            for key, value in variables.items():

                while key in paragraph.text:

                    runs = paragraph.runs
                    full_text = "".join(r.text for r in runs)

                    start = full_text.find(key)
                    if start == -1:
                        break

                    end = start + len(key)

                    char_count = 0
                    start_run = None
                    end_run = None
                    start_offset = 0
                    end_offset = 0

                    for i, run in enumerate(runs):
                        run_len = len(run.text)

                        if start_run is None and char_count + run_len > start:
                            start_run = i
                            start_offset = start - char_count

                        if char_count + run_len >= end:
                            end_run = i
                            end_offset = end - char_count
                            break

                        char_count += run_len

                    if start_run is None or end_run is None:
                        break

                    before = runs[start_run].text[:start_offset]
                    after = runs[end_run].text[end_offset:]

                    for i in range(start_run, end_run + 1):
                        runs[i].text = ""

                    runs[start_run].text = before + str(value) + after

        # ================= LẤY FILE MẪU ĐỘNG =================
        template_files = [
            f for f in os.listdir(self.template_dir)
            if f.lower().endswith(".docx")
        ]

        if not template_files:
            messagebox.showerror("Lỗi", "Thư mục mẫu không có file Word.")
            return

        # ================= TẠO THƯ MỤC LƯU =================
        def bo_dau(text):
            text = unicodedata.normalize('NFD', text)
            return ''.join(c for c in text if unicodedata.category(c) != 'Mn')

        folder_name = re.sub(r'[\\/:*?"<>|]', '', bo_dau(ho_ten_day_du))
        target_dir = os.path.join(SAVE_DIR, folder_name)
        os.makedirs(target_dir, exist_ok=True)

        # ================= XUẤT FILE =================
        for file_name in template_files:

            template_path = os.path.join(self.template_dir, file_name)
            doc = Document(template_path)

            for p in doc.paragraphs:
                replace_paragraph(p)

            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_paragraph(p)

            doc.save(os.path.join(target_dir, file_name))

        messagebox.showinfo("Hoàn tất", f"Đã xuất {len(template_files)} file vào:\n{target_dir}")

    def auto_save_profile(self):

        data = self.collect_data()

        ten = data["doi_tuong"].get("ten", "Ho_so")
        ho = data["doi_tuong"].get("ho_dem", "")

        file_name = f"{ho}_{ten}.json".replace(" ", "_")
        path = os.path.join(PROFILE_DIR, file_name)

        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        
    def send_report_placeholder(self):
        messagebox.showinfo("Thông báo", "Chức năng gửi email sẽ phát triển sau.")
        
# ================= RUN =================
if __name__ == "__main__":
    app = App()
    app.mainloop()
