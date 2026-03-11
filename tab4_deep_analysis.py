
# tab4_deep_analysis.py
from __future__ import annotations
import os, threading, datetime, re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from typing import Any, Dict, List, Optional
import pandas as pd

from deep_engine import (
    MappingConfig, OwnerInfo, CounterpartyConfig, TagRule, SuspiciousConfig,
    BankDetector, StatementLoader, suggest_mapping, DeepPipeline, save_json, load_json, filter_garbage_rows
)

try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CFG_PATH_DEFAULT = os.path.join(BASE_DIR, "tab4_config.json")

def _fmt_money(x: float) -> str:
    try: return f"{float(x):,.0f}"
    except Exception: return str(x)

class ToolTip:
    def __init__(self, widget, text: str):
        self.widget=widget; self.text=text; self.tip=None
        widget.bind("<Enter>", self.show); widget.bind("<Leave>", self.hide)
    def show(self, _=None):
        if self.tip or not self.text: return
        x=self.widget.winfo_rootx()+20; y=self.widget.winfo_rooty()+self.widget.winfo_height()+6
        self.tip=tk.Toplevel(self.widget); self.tip.wm_overrideredirect(True); self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, justify="left", background="#ffffe0", relief="solid", borderwidth=1).pack(ipadx=6, ipady=4)
    def hide(self, _=None):
        if self.tip: self.tip.destroy(); self.tip=None

class TabPhanTichSKChuyenSau(ttk.Frame):
    MAX_ROWS_RAW=8000
    MAX_ROWS_SUM=8000

    def __init__(self, parent, cfg_path: str = CFG_PATH_DEFAULT):
        super().__init__(parent)
        self.cfg_path=cfg_path
        self.selected_files: List[str]=[]
        self.reports: Dict[str,pd.DataFrame]={}

        self.mapping=MappingConfig()
        self.owner=OwnerInfo()
        self.cp_cfg=CounterpartyConfig(min_confidence_for_stats=0.6, max_candidates=3)
        self.susp_cfg=SuspiciousConfig()
        self.tag_rules: List[TagRule]=[]

        self.files_var=tk.StringVar(value="Chưa chọn file")
        self.status_var=tk.StringVar(value="Sẵn sàng")

        self.bank_mode=tk.StringVar(value="AUTO")
        self.bank_other_name=tk.StringVar(value="")
        self.bank_auto_var=tk.StringVar(value="-")
        self.bank_conf_var=tk.StringVar(value="-")

        self.owner_acc_var=tk.StringVar(value="")
        self.owner_name_var=tk.StringVar(value="")
        self.min_conf_var=tk.StringVar(value="0.6")

        self.mode_money=tk.StringVar(value="BOTH")
        self.show_forensic_raw=tk.BooleanVar(value=False)

        self.search_var=tk.StringVar(value="")
        self.filter_var=tk.StringVar(value="")

        self.sum_tx=tk.StringVar(value="0")
        self.sum_in=tk.StringVar(value="0")
        self.sum_out=tk.StringVar(value="0")
        self.sum_net=tk.StringVar(value="0")
        self.sum_cp=tk.StringVar(value="0")
        self.sum_susp=tk.StringVar(value="0")
        self.sum_issues=tk.StringVar(value="0")

        self.sash_px=320

        self._load_cfg()
        self._build_ui()

    def _default_rules(self) -> List[TagRule]:
        return [
            TagRule(name="Lương", priority=10, include=["luong","lương","salary","payroll"], confidence=0.85, direction="CREDIT"),
            TagRule(name="Nhà", priority=20, include=["tien nha","tiền nhà","thue nha","thuê nhà","dat coc","đặt cọc"], confidence=0.8),
            TagRule(name="Ví điện tử", priority=30, include=["momo","zalopay","shopeepay","vnpay","viettel money"], confidence=0.85, require_wallet=True),
            TagRule(name="Topup SĐT", priority=40, include=["nap","nạp","topup","the cao","thẻ cào","viettel","vinaphone","mobifone"], confidence=0.85, require_phone=True),
            TagRule(name="Điện", priority=50, include=["evn","tiền điện","dien luc","điện lực"], confidence=0.85, require_utility=True),
            TagRule(name="Nước", priority=55, include=["tiền nước","cap nuoc","cấp nước"], confidence=0.8, require_utility=True),
            TagRule(name="Internet", priority=60, include=["fpt","vnpt","viettel","internet"], confidence=0.75),
            TagRule(name="Học phí", priority=70, include=["hoc phi","học phí","tuition","truong","trường"], confidence=0.75),
            TagRule(name="Xe cộ", priority=80, include=["xang","xăng","parking","bot","b.o.t","tram thu phi","trạm thu phí"], confidence=0.7),
            TagRule(name="Gia đình", priority=90, include=["vo","vợ","chong","chồng","con","me","mẹ","bo","bố"], confidence=0.65),
        ]

    def _load_cfg(self) -> None:
        if not os.path.exists(self.cfg_path):
            self.tag_rules=self._default_rules()
            self._save_cfg()
            return
        try:
            cfg=load_json(self.cfg_path)
            self.mapping=MappingConfig.from_dict(cfg.get("mapping",{}))
            self.owner=OwnerInfo.from_dict(cfg.get("owner",{}))
            self.cp_cfg=CounterpartyConfig.from_dict(cfg.get("counterparty",{}))
            self.susp_cfg=SuspiciousConfig.from_dict(cfg.get("suspicious",{}))
            rules=cfg.get("tag_rules")
            self.tag_rules=[TagRule.from_dict(x) for x in rules] if isinstance(rules,list) else self._default_rules()
            ui=cfg.get("ui",{})
            self.sash_px=int(ui.get("sash_px",320))
            self.owner_acc_var.set(self.owner.account or "")
            self.owner_name_var.set(self.owner.name or "")
            self.min_conf_var.set(str(self.cp_cfg.min_confidence_for_stats))
        except Exception:
            self.tag_rules=self._default_rules()

    def _save_cfg(self) -> None:
        """Lưu cấu hình (merge-safe) để KHÔNG làm mất dữ liệu khác trong tab4_config.json.
        - Giữ nguyên: profiles, bank_templates và các key lạ khác.
        - Chỉ cập nhật: mapping/owner/counterparty/suspicious/tag_rules/ui.
        """

        # sync UI -> model
        self.owner.account = self.owner_acc_var.get().strip()
        self.owner.name = self.owner_name_var.get().strip()
        try:
            self.cp_cfg.min_confidence_for_stats = float(self.min_conf_var.get().strip())
        except Exception:
            pass

        cfg: Dict[str, Any] = {}
        if os.path.exists(self.cfg_path):
            try:
                cfg = load_json(self.cfg_path)
                if not isinstance(cfg, dict):
                    cfg = {}
            except Exception:
                cfg = {}

        # đảm bảo tồn tại các vùng lưu trữ dài hạn
        if not isinstance(cfg.get("profiles"), dict):
            cfg["profiles"] = {}
        if not isinstance(cfg.get("bank_templates"), dict):
            cfg["bank_templates"] = {}

        # update phần cấu hình tab4
        cfg["mapping"] = self.mapping.to_dict()
        cfg["owner"] = self.owner.to_dict()
        cfg["counterparty"] = self.cp_cfg.to_dict()
        cfg["suspicious"] = self.susp_cfg.to_dict()
        cfg["tag_rules"] = [r.to_dict() for r in self.tag_rules]
        cfg["ui"] = {"sash_px": int(self.sash_px)}

        save_json(self.cfg_path, cfg)



    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        toolbar=ttk.Frame(self)
        toolbar.grid(row=0,column=0,sticky="ew",padx=6,pady=(6,2))
        toolbar.columnconfigure(8,weight=1)

        ttk.Button(toolbar,text="📂 Chọn file",command=self.choose_files).grid(row=0,column=0,padx=2)
        ttk.Button(toolbar,text="🧭 Mapping",command=self.open_mapping_popup).grid(row=0,column=1,padx=2)
        ttk.Button(toolbar,text="🔎 Phân tích",command=self.analyze_thread).grid(row=0,column=2,padx=2)
        ttk.Button(toolbar,text="💾 Xuất Excel",command=self.export_excel).grid(row=0,column=3,padx=2)
        ttk.Button(toolbar,text="🧾 Xuất PDF",command=self.export_pdf).grid(row=0,column=4,padx=2)
        ttk.Label(toolbar,textvariable=self.files_var).grid(row=0,column=9,sticky="e",padx=6)

        status=ttk.Frame(self)
        status.grid(row=1,column=0,sticky="ew",padx=6,pady=4)
        status.columnconfigure(0,weight=1)
        self.progress=ttk.Progressbar(status,mode="determinate",maximum=100)
        self.progress.grid(row=0,column=0,sticky="ew")
        ttk.Label(status,textvariable=self.status_var).grid(row=0,column=1,sticky="e",padx=8)

        self.paned=tk.PanedWindow(self,orient=tk.HORIZONTAL,sashrelief="raised",sashwidth=6,bd=0)
        self.paned.grid(row=2,column=0,sticky="nsew",padx=6,pady=6)

        self.left=ttk.Frame(self.paned)
        self.right=ttk.Frame(self.paned)
        self.paned.add(self.left,minsize=300)
        self.paned.add(self.right,minsize=500)

        self._build_left()
        self._build_right()

        def _init_sash():
            try:
                self.update_idletasks()
                w=max(900,self.winfo_width())
                target=int(min(max(300,self.sash_px), w*0.6))
                self.paned.sash_place(0,target,1)
            except Exception:
                pass
        self.after(150,_init_sash)

        def _poll_sash():
            try:
                x,_=self.paned.sash_coord(0)
                if x!=self.sash_px: self.sash_px=x
            except Exception:
                pass
            self.after(500,_poll_sash)
        self.after(500,_poll_sash)

    def _build_left(self) -> None:
        # Layout ổn định: Notebook (trên) + Tổng quan (dưới)
        # Không dùng Canvas scroll để tránh lỗi "mất panel trái"
        self.left.columnconfigure(0, weight=1)
        self.left.rowconfigure(0, weight=1)
        self.left.rowconfigure(1, weight=0)

        nb = ttk.Notebook(self.left)
        nb.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        tab_chung = ttk.Frame(nb)
        tab_tag = ttk.Frame(nb)
        tab_nghi = ttk.Frame(nb)
        tab_log = ttk.Frame(nb)

        nb.add(tab_chung, text="Chung")
        nb.add(tab_tag, text="Tag")
        nb.add(tab_nghi, text="Nghi vấn")
        nb.add(tab_log, text="Log")

        # ========== TAB CHUNG ==========
        frm = ttk.LabelFrame(tab_chung, text="Cấu hình chung")
        frm.pack(fill="x", padx=6, pady=6)
        frm.columnconfigure(1, weight=1)

        bank_opts = ["AUTO","MB","TCB","VCB","BIDV","VTB","VIB","STB","SHB","EXB","OTHER"]

        ttk.Label(frm, text="Ngân hàng:").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        cb = ttk.Combobox(frm, values=bank_opts, textvariable=self.bank_mode, width=10, state="readonly")
        cb.grid(row=0, column=1, sticky="w", padx=6, pady=3)
        cb.bind("<<ComboboxSelected>>", lambda e: self._toggle_bank_other())

        ttk.Label(frm, text="Tên NH khác:").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        self.bank_other_entry = ttk.Entry(frm, textvariable=self.bank_other_name)
        self.bank_other_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=3)

        ttk.Label(frm, text="Auto NH:").grid(row=2, column=0, sticky="w", padx=6, pady=3)
        ttk.Label(frm, textvariable=self.bank_auto_var, foreground="gray").grid(row=2, column=1, sticky="w", padx=6, pady=3)

        ttk.Label(frm, text="Tin cậy:").grid(row=3, column=0, sticky="w", padx=6, pady=3)
        ttk.Label(frm, textvariable=self.bank_conf_var, foreground="gray").grid(row=3, column=1, sticky="w", padx=6, pady=3)

        ttk.Separator(frm, orient="horizontal").grid(row=4, column=0, columnspan=2, sticky="ew", pady=8)

        ttk.Label(frm, text="STK chủ:").grid(row=5, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(frm, textvariable=self.owner_acc_var).grid(row=5, column=1, sticky="ew", padx=6, pady=3)

        ttk.Label(frm, text="Chủ TK:").grid(row=6, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(frm, textvariable=self.owner_name_var).grid(row=6, column=1, sticky="ew", padx=6, pady=3)

        ttk.Label(frm, text="Min tin cậy đối ứng:").grid(row=7, column=0, sticky="w", padx=6, pady=3)
        e_min = ttk.Entry(frm, textvariable=self.min_conf_var, width=10)
        e_min.grid(row=7, column=1, sticky="w", padx=6, pady=3)
        ToolTip(e_min, "Đối ứng tin cậy < ngưỡng sẽ không lên thống kê Đối ứng (nhưng vẫn thấy ở Giao dịch).")

        mode_box = ttk.LabelFrame(tab_chung, text="Chế độ phân tích theo chiều tiền")
        mode_box.pack(fill="x", padx=6, pady=(0,6))
        ttk.Radiobutton(mode_box, text="Tiền đi (Debit)", variable=self.mode_money, value="DEBIT").pack(anchor="w", padx=8, pady=1)
        ttk.Radiobutton(mode_box, text="Tiền đến (Credit)", variable=self.mode_money, value="CREDIT").pack(anchor="w", padx=8, pady=1)
        ttk.Radiobutton(mode_box, text="2 chiều", variable=self.mode_money, value="BOTH").pack(anchor="w", padx=8, pady=1)

        ttk.Checkbutton(tab_chung, text="Xem RAW forensic (debug)", variable=self.show_forensic_raw, command=self._rerender_raw_only)\
            .pack(anchor="w", padx=10, pady=(0,6))

        map_box = ttk.LabelFrame(tab_chung, text="Mapping hiện tại")
        map_box.pack(fill="x", padx=6, pady=6)

        self.map_summary = tk.StringVar(value="Chưa có mapping. Hãy chọn file → Mapping.")
        ttk.Label(map_box, textvariable=self.map_summary, foreground="gray", justify="left")\
            .pack(anchor="w", padx=8, pady=6)

        btn_row = ttk.Frame(tab_chung)
        btn_row.pack(fill="x", padx=6, pady=(0,6))
        ttk.Button(btn_row, text="💾 Lưu cấu hình", command=self.save_config).pack(side="left")
        ttk.Button(btn_row, text="🧭 Mở Mapping", command=self.open_mapping_popup).pack(side="right")

        # ========== TAB TAG (BẢNG + NÚT Ở DƯỚI, KHÔNG BỊ ĐẨY SANG PHẢI) ==========
        tab_tag.columnconfigure(0, weight=1)
        tab_tag.rowconfigure(1, weight=1)

        hint = ttk.Label(
            tab_tag,
            text=("REQ (điều kiện): chỉ gắn tag khi phát hiện đúng thực thể.\n"
                  "• SĐT • Ví • Biển số • Mã DV | Direction: BOTH/DEBIT/CREDIT"),
            foreground="gray",
            justify="left"
        )
        hint.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))

        # Khung bảng (để grid chuẩn, scrollbar không ăn layout)
        table_frame = ttk.Frame(tab_tag)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        cols = ("priority","enabled","name","type","direction","req","include","exclude","conf")
        self.rule_tv = ttk.Treeview(table_frame, columns=cols, show="headings")

        headers = {
            "priority": "Ưu tiên",
            "enabled": "Bật",
            "name": "Tag",
            "type": "Kiểu",
            "direction": "Chiều",
            "req": "REQ",
            "include": "Include",
            "exclude": "Exclude",
            "conf": "Tin cậy",
        }
        widths = {
            "priority": 62,
            "enabled": 42,
            "name": 120,
            "type": 60,
            "direction": 70,
            "req": 80,
            "include": 260,
            "exclude": 240,
            "conf": 70,
        }

        for c in cols:
            self.rule_tv.heading(c, text=headers[c])
            self.rule_tv.column(c, width=widths[c], stretch=(c in ("include", "exclude")))

        vs = ttk.Scrollbar(table_frame, orient="vertical", command=self.rule_tv.yview)
        self.rule_tv.configure(yscrollcommand=vs.set)

        self.rule_tv.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")

        # Nút Tag chia 2 hàng, nằm DƯỚI BẢNG
        btns = ttk.Frame(tab_tag)
        btns.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 6))
        for i in range(3):
            btns.columnconfigure(i, weight=1)

        ttk.Button(btns, text="Thêm", command=self.add_rule).grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Sửa", command=self.edit_rule).grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Xóa", command=self.delete_rule).grid(row=0, column=2, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Bật/Tắt", command=self.toggle_rule).grid(row=1, column=0, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Lưu", command=self.save_rules).grid(row=1, column=1, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Reset", command=self.reset_rules).grid(row=1, column=2, padx=2, pady=2, sticky="ew")

        self._refresh_rule_tree()

        # ========== TAB NGHI VẤN ==========
        info = ttk.Label(tab_nghi, text=(
            "Nghi vấn giao dịch:\n"
            "• Giờ đêm • Tiền lớn • Z-score • Nhiều GD/ngày • Lặp số tiền/ngày"
        ), foreground="gray", justify="left")
        info.pack(fill="x", padx=6, pady=(6,4))

        form = ttk.LabelFrame(tab_nghi, text="Tham số")
        form.pack(fill="x", padx=6, pady=6)
        form.columnconfigure(1, weight=1)

        # đảm bảo các biến tồn tại
        if not hasattr(self, "big_amt_var"):
            self.big_amt_var = tk.StringVar(value=str(int(self.susp_cfg.big_amount_threshold)))
            self.night_from_var = tk.StringVar(value=str(self.susp_cfg.night_start_hour))
            self.night_to_var = tk.StringVar(value=str(self.susp_cfg.night_end_hour))
            self.z_var = tk.StringVar(value=str(self.susp_cfg.zscore_threshold))
            self.burst_var = tk.StringVar(value=str(self.susp_cfg.burst_txn_count))
            self.split_var = tk.StringVar(value=str(self.susp_cfg.split_repeat_count))

        def row(r, label, var):
            ttk.Label(form, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=3)
            ttk.Entry(form, textvariable=var).grid(row=r, column=1, sticky="ew", padx=6, pady=3)

        row(0, "Ngưỡng tiền lớn:", self.big_amt_var)
        row(1, "Giờ đêm từ:", self.night_from_var)
        row(2, "Giờ đêm đến:", self.night_to_var)
        row(3, "Z-score:", self.z_var)
        row(4, "Nhiều GD/ngày:", self.burst_var)
        row(5, "Lặp số tiền/ngày:", self.split_var)

        ttk.Button(tab_nghi, text="💾 Lưu cấu hình nghi vấn", command=self.save_suspicious).pack(anchor="e", padx=6, pady=(0,6))

        # ========== TAB LOG ==========
        self.log_text = tk.Text(tab_log, height=10, wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

        # ========== TỔNG QUAN (CỐ ĐỊNH DƯỚI) ==========
        summary = ttk.LabelFrame(self.left, text="Tổng quan")
        summary.grid(row=1, column=0, sticky="ew", padx=6, pady=(0,6))
        summary.columnconfigure(1, weight=1)

        def line(r, label, var):
            ttk.Label(summary, text=label, width=10).grid(row=r, column=0, sticky="w", padx=6, pady=2)
            ttk.Label(summary, textvariable=var, font=("Segoe UI", 9, "bold")).grid(row=r, column=1, sticky="w", padx=6, pady=2)

        line(0, "Tổng số GD:", self.sum_tx)
        line(1, "Tổng tiền vào:", self.sum_in)
        line(2, "Tổng tiền ra:", self.sum_out)
        line(3, "NET:", self.sum_net)
        line(4, "Đối ứng:", self.sum_cp)
        line(5, "GD Nghi vấn:", self.sum_susp)
        line(6, "Lỗi dữ liệu:", self.sum_issues)

        self._toggle_bank_other()
        self._update_mapping_summary()

    def _build_right(self) -> None:
        top=ttk.Frame(self.right)
        top.pack(fill="x",padx=6,pady=(0,6))
        ttk.Label(top,text="Tìm:").pack(side="left")
        ent=ttk.Entry(top,textvariable=self.search_var)
        ent.pack(side="left",fill="x",expand=True,padx=6)
        ent.bind("<Return>",lambda e:self.apply_search_filter())
        ttk.Label(top,text="Lọc:").pack(side="left")
        flt=ttk.Entry(top,textvariable=self.filter_var,width=18)
        flt.pack(side="left",padx=6)
        flt.bind("<Return>",lambda e:self.apply_search_filter())
        ToolTip(flt,"Ví dụ: debit | credit | high | medium | tag:TênTag")
        ttk.Button(top,text="Áp dụng",command=self.apply_search_filter).pack(side="left",padx=2)
        ttk.Button(top,text="Xóa",command=self.clear_search_filter).pack(side="left",padx=2)
        ttk.Button(top,text="Copy tất cả",command=self.copy_all_current).pack(side="right",padx=6)

        self.nb=ttk.Notebook(self.right)
        self.nb.pack(fill="both",expand=True,padx=6,pady=6)
        self.tab_map={"RAW":"Giao dịch","COUNTERPARTY":"Đối ứng","TAG":"Theo Tag","TAG_ENTITY":"Tag + Thực thể","SUSPICIOUS":"Nghi vấn","DATA_QUALITY":"Chất lượng dữ liệu"}
        self.tables={}
        for key,title in self.tab_map.items():
            frame=ttk.Frame(self.nb); self.nb.add(frame,text=title)
            tv=ttk.Treeview(frame,show="headings")
            vs=ttk.Scrollbar(frame,orient="vertical",command=tv.yview)
            hs=ttk.Scrollbar(frame,orient="horizontal",command=tv.xview)
            tv.configure(yscrollcommand=vs.set,xscrollcommand=hs.set)
            tv.grid(row=0,column=0,sticky="nsew")
            vs.grid(row=0,column=1,sticky="ns")
            hs.grid(row=1,column=0,sticky="ew")
            frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1)
            tv.bind("<Control-c>", self.copy_selected)
            self.tables[key]=tv

    def _toggle_bank_other(self):
        mode=(self.bank_mode.get() or "AUTO").upper()
        if mode=="OTHER":
            self.bank_other_entry.configure(state="normal")
        else:
            self.bank_other_entry.delete(0,"end")
            self.bank_other_entry.configure(state="disabled")

    def log(self,msg:str)->None:
        try:
            self.log_text.insert("end",msg+"\n"); self.log_text.see("end")
        except Exception:
            pass

    def _set_progress(self,pct:int,msg:str):
        self.progress["value"]=max(0,min(100,int(pct)))
        self.status_var.set(f"{msg} ({int(pct)}%)")

    def save_config(self)->None:
        self._sync_suspicious_from_ui()
        self._save_cfg()
        self.log("Đã lưu cấu hình.")

    def _sync_suspicious_from_ui(self):
        try:
            self.susp_cfg.big_amount_threshold=float(self.big_amt_var.get().strip())
            self.susp_cfg.night_start_hour=int(self.night_from_var.get().strip())
            self.susp_cfg.night_end_hour=int(self.night_to_var.get().strip())
            self.susp_cfg.zscore_threshold=float(self.z_var.get().strip())
            self.susp_cfg.burst_txn_count=int(self.burst_var.get().strip())
            self.susp_cfg.split_repeat_count=int(self.split_var.get().strip())
        except Exception:
            pass

    def save_rules(self)->None:
        self._save_cfg(); self.log("Đã lưu quy tắc Tag.")

    def reset_rules(self)->None:
        if not messagebox.askyesno("Reset","Reset quy tắc Tag về mặc định?"): return
        self.tag_rules=self._default_rules()
        self._refresh_rule_tree()
        self._save_cfg()
        self.log("Đã reset quy tắc Tag.")

    def save_suspicious(self)->None:
        self._sync_suspicious_from_ui()
        self._save_cfg()
        self.log("Đã lưu cấu hình nghi vấn.")

    def choose_files(self)->None:

        paths = filedialog.askopenfilenames(filetypes=[("Excel/CSV","*.xlsx *.xls *.csv")])

        if not paths:

            return

        self.selected_files = list(paths)

        self.files_var.set(f"Đã chọn: {len(self.selected_files)} file")


        first = self.selected_files[0]


        def _try_read_headers(path: str):

            trials = []

            if path.lower().endswith(".csv"):

                for hr in (0,1,2,3):

                    try:

                        df = pd.read_csv(path, header=hr, nrows=20)

                        df.columns = [str(c).strip() for c in df.columns]

                        trials.append(list(df.columns))

                    except Exception:

                        pass

                return trials


            for hr in (0,1,2,3,4,5):

                try:

                    df = pd.read_excel(path, sheet_name=0, header=hr, nrows=20)

                    df.columns = [str(c).strip() for c in df.columns]

                    trials.append(list(df.columns))

                except Exception:

                    pass

            return trials


        try:

            best_bank, best_conf = "OTHER", 0.0

            for cols in _try_read_headers(first):

                if not cols:

                    continue

                bank, conf = BankDetector.detect(cols)

                if conf > best_conf:

                    best_bank, best_conf = bank, conf


            if best_conf <= 0:

                self.bank_auto_var.set("-"); self.bank_conf_var.set("-")

                self.log("Auto-detect NH: không xác định (không đọc được header phù hợp).")

            else:

                self.bank_auto_var.set(best_bank); self.bank_conf_var.set(f"{best_conf:.2f}")

                self.log(f"Auto-detect NH: {best_bank} (tin cậy {best_conf:.2f}).")

        except Exception as e:

            self.bank_auto_var.set("-"); self.bank_conf_var.set("-")

            self.log(f"Auto-detect NH thất bại: {e}")


        self._update_mapping_summary()


    def _read_preview(self,path:str,sheet_name:str,header_row:int,nrows:int=30)->pd.DataFrame:
        hr=int(header_row or 0)
        if path.lower().endswith(".csv"):
            df=pd.read_csv(path,header=hr,nrows=nrows)
        else:
            sn=0 if not sheet_name else sheet_name
            try:
                df=pd.read_excel(path,header=hr,sheet_name=sn,nrows=nrows)
            except Exception:
                df=pd.read_excel(path,header=hr,sheet_name=0,nrows=nrows)
        df.columns=[str(c).strip() for c in df.columns]
        return df

    def _update_mapping_summary(self):
        m=self.mapping
        lines=[
            f"Sheet: {(m.sheet_name if m.sheet_name else '(Sheet đầu tiên)')} | Header row: {m.header_row}",
            f"Lọc dòng rác: {(getattr(m, 'garbage_filter_level', 'CHUAN') or 'CHUAN').upper()}",
            f"Nội dung: {m.content_col or '—'}",
            f"Ngày: {m.date_col or '—'} | Ngày giờ: {m.datetime_col or '—'} | Giờ: {m.time_col or '—'}",
            f"Nợ: {m.debit_col or '—'} | Có: {m.credit_col or '—'} | Số tiền: {m.amount_col or '—'} | Loại: {m.direction_col or '—'}",
            f"Đối ứng STK: {m.cp_account_col or '—'} | Tên: {m.cp_name_col or '—'} | NH: {m.cp_bank_col or '—'}",
            f"TK chuyển: {m.sender_acc_col or '—'} | TK nhận: {m.receiver_acc_col or '—'}",
        ]
        self.map_summary.set("\n".join(lines))

    def open_mapping_popup(self) -> None:
        if not self.selected_files:
            messagebox.showwarning("Thiếu file", "Hãy chọn file trước.")
            return

        w = tk.Toplevel(self)
        w.title("Mapping (Auto → Xác nhận)")
        w.grab_set()
        w.geometry("1200x760")
        w.minsize(1050, 650)

        container = ttk.Frame(w)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)  # preview expand
        container.rowconfigure(3, weight=0)  # mapping fixed-ish (scrollable if needed)

        first = self.selected_files[0]
        sheets = StatementLoader.list_sheets(first)
        sheet_opts = ["(Sheet đầu tiên)"] + sheets if sheets else ["(Sheet đầu tiên)"]

        sheet_var = tk.StringVar(value=(self.mapping.sheet_name if self.mapping.sheet_name else "(Sheet đầu tiên)"))
        header_var = tk.IntVar(value=int(self.mapping.header_row or 0))
        skip_top_var = tk.IntVar(value=0)
        # Mức lọc dòng rác: TAT / NHE / CHUAN / MANH
        filter_level_var = tk.StringVar(value=(getattr(self.mapping,'garbage_filter_level','CHUAN') or 'CHUAN').upper())
        info_var = tk.StringVar(value="")
        template_var = tk.StringVar(value="")


        # =========================
        # 1) CẤU HÌNH (TOP)
        # =========================
        top = ttk.LabelFrame(container, text="Cấu hình đọc file")
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top.columnconfigure(10, weight=1)

        ttk.Label(top, text="File mẫu:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Label(top, text=os.path.basename(first), foreground="gray").grid(row=0, column=1, sticky="w", padx=8, pady=6)

        ttk.Label(top, text="Sheet:").grid(row=0, column=2, sticky="w", padx=8, pady=6)
        ttk.Combobox(top, values=sheet_opts, textvariable=sheet_var, state="readonly", width=18)\
            .grid(row=0, column=3, sticky="w", padx=8, pady=6)

        ttk.Label(top, text="Header row:").grid(row=0, column=4, sticky="w", padx=8, pady=6)
        ttk.Spinbox(top, from_=0, to=120, textvariable=header_var, width=6)\
            .grid(row=0, column=5, sticky="w", padx=8, pady=6)

        ttk.Label(top, text="Bỏ qua dòng đầu:").grid(row=0, column=6, sticky="w", padx=8, pady=6)
        ttk.Spinbox(top, from_=0, to=300, textvariable=skip_top_var, width=6)\
            .grid(row=0, column=7, sticky="w", padx=8, pady=6)

        ttk.Label(top, text="Mức lọc:").grid(row=0, column=8, sticky="w", padx=8, pady=6)
        ttk.Combobox(top, values=["TAT","NHE","CHUAN","MANH"], textvariable=filter_level_var, state="readonly", width=8)\
            .grid(row=0, column=9, sticky="w", padx=8, pady=6)

        ttk.Label(top, textvariable=info_var, foreground="gray")\
            .grid(row=0, column=10, sticky="w", padx=8, pady=6)

        # --- Bank templates (nhiều mẫu / mỗi bank) ---
        ttk.Label(top, text="Mẫu NH:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        cb_tmpl = ttk.Combobox(top, textvariable=template_var, values=[], state="readonly", width=28)
        cb_tmpl.grid(row=1, column=1, sticky="w", padx=8, pady=6)
        btn_apply_tmpl = ttk.Button(top, text="Áp mẫu", command=lambda: None)
        btn_apply_tmpl.grid(row=1, column=2, sticky="w", padx=(8,2), pady=6)
        btn_save_tmpl = ttk.Button(top, text="⭐ Lưu mẫu", command=lambda: None)
        btn_save_tmpl.grid(row=1, column=3, sticky="w", padx=(2,2), pady=6)
        btn_del_tmpl = ttk.Button(top, text="🗑 Xóa mẫu", command=lambda: None)
        btn_del_tmpl.grid(row=1, column=4, sticky="w", padx=(2,2), pady=6)
        ttk.Label(top, text="(Template theo bank: VCB1/VCB2...)" , foreground="gray").grid(row=1, column=5, columnspan=6, sticky="w", padx=8, pady=6)


        # =========================
        # 2) PREVIEW (GIỮA)
        # =========================
        pv_box = ttk.LabelFrame(container, text="Preview (30 dòng)")
        pv_box.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        pv_box.columnconfigure(0, weight=1)
        pv_box.rowconfigure(0, weight=1)

        pv = ttk.Treeview(pv_box, show="headings")
        pvs = ttk.Scrollbar(pv_box, orient="vertical", command=pv.yview)
        pvh = ttk.Scrollbar(pv_box, orient="horizontal", command=pv.xview)
        pv.configure(yscrollcommand=pvs.set, xscrollcommand=pvh.set)

        pv.grid(row=0, column=0, sticky="nsew")
        pvs.grid(row=0, column=1, sticky="ns")
        pvh.grid(row=1, column=0, sticky="ew")

        def render_preview(df: pd.DataFrame):
            pv.delete(*pv.get_children())
            if df is None or df.empty:
                pv["columns"] = []
                return
            cols = list(df.columns)
            pv_cols = ["#"] + cols
            pv["columns"] = pv_cols
            pv["show"] = "headings"

            pv.heading("#", text="STT")
            pv.column("#", width=60, stretch=False, anchor="center")

            for c in cols:
                pv.heading(c, text=c)
                pv.column(c, width=170, stretch=True, anchor="w")

            for i, (_, row) in enumerate(df.head(30).iterrows(), start=1):
                pv.insert("", "end", values=[i] + [row.get(c, "") for c in cols])

        # =========================
        # 3) MAPPING (DƯỚI) - CHIA 2 CỘT TRÁI/PHẢI
        # =========================
        map_outer = ttk.LabelFrame(container, text="Mapping cột (Không chọn = bỏ qua)")
        map_outer.grid(row=3, column=0, sticky="ew")  # mapping không cần ăn height, có scroll riêng
        map_outer.columnconfigure(0, weight=1)
        map_outer.rowconfigure(0, weight=1)

        # scroll area cho mapping (vì nhiều dòng)
        map_canvas = tk.Canvas(map_outer, highlightthickness=0, height=260)
        map_vsb = ttk.Scrollbar(map_outer, orient="vertical", command=map_canvas.yview)
        map_canvas.configure(yscrollcommand=map_vsb.set)

        map_canvas.grid(row=0, column=0, sticky="nsew")
        map_vsb.grid(row=0, column=1, sticky="ns")

        map_inner = ttk.Frame(map_canvas)
        win = map_canvas.create_window((0, 0), window=map_inner, anchor="nw")

        def _on_map_configure(_=None):
            map_canvas.configure(scrollregion=map_canvas.bbox("all"))
            map_canvas.itemconfigure(win, width=map_canvas.winfo_width())
        map_inner.bind("<Configure>", _on_map_configure)

        # 2 cột mapping: (0,label)-(1,combo) và (2,label)-(3,combo)
        map_inner.columnconfigure(1, weight=1)
        map_inner.columnconfigure(3, weight=1)

        combos = {}

        def add_field(row, col_base, label, key):
            ttk.Label(map_inner, text=label, width=16).grid(row=row, column=col_base, sticky="w", padx=(8, 6), pady=4)
            var = tk.StringVar(value="Không chọn")
            cb = ttk.Combobox(map_inner, textvariable=var, values=["Không chọn"], state="readonly", width=28)
            cb.grid(row=row, column=col_base + 1, sticky="ew", padx=(0, 10), pady=4)
            combos[key] = (var, cb)

        # Bố trí hợp lý:
        # LEFT: Nội dung/Ngày/Ngày giờ/Giờ + Debit/Credit/Amount/Direction
        # RIGHT: Đối ứng + sender/receiver
        # Row 0-3: core
        add_field(0, 0, "Nội dung (*)", "content_col")
        add_field(0, 2, "STK đối ứng", "cp_account_col")

        add_field(1, 0, "Ngày (*)", "date_col")
        add_field(1, 2, "Tên đối ứng", "cp_name_col")

        add_field(2, 0, "Ngày giờ (chung)", "datetime_col")
        add_field(2, 2, "NH đối ứng", "cp_bank_col")

        add_field(3, 0, "Giờ", "time_col")
        add_field(3, 2, "TK chuyển", "sender_acc_col")

        ttk.Separator(map_inner, orient="horizontal").grid(row=4, column=0, columnspan=4, sticky="ew", pady=8)

        add_field(5, 0, "Debit (PS Nợ)", "debit_col")
        add_field(5, 2, "TK nhận", "receiver_acc_col")

        add_field(6, 0, "Credit (PS Có)", "credit_col")
        add_field(6, 2, "Loại GD (IN/OUT)", "direction_col")

        add_field(7, 0, "Số tiền (fallback)", "amount_col")
        # cột phải hàng 7 để trống (để cân)
        ttk.Label(map_inner, text="", width=16).grid(row=7, column=2, sticky="w", padx=(8, 6), pady=4)

        ttk.Label(
            map_inner,
            text="(*) bắt buộc. Nếu không có Debit/Credit thì phải có Số tiền hoặc Loại giao dịch.",
            foreground="gray"
        ).grid(row=8, column=0, columnspan=4, sticky="w", padx=8, pady=(6, 10))

        # =========================
        # 4) FOOTER (NÚT CUỐI)
        # =========================
        bottom = ttk.Frame(container)
        bottom.grid(row=4, column=0, sticky="ew", pady=(10, 0))

        # ---------------- STATE ----------------
        state = {"cols": [], "preview": pd.DataFrame(), "sig": "", "suggest": None}

        def set_combo_values(cols: List[str]):
            vals = ["Không chọn"] + cols
            for _, cb in combos.values():
                cb["values"] = vals

        def set_combo_from_mapping(m: MappingConfig):
            def pick(val: str) -> str:
                return val if val and val in state["cols"] else "Không chọn"
            for key in combos.keys():
                combos[key][0].set(pick(getattr(m, key)))

        def effective_bank(cols: List[str]) -> str:
            bank, conf = BankDetector.detect(cols)
            self.bank_auto_var.set(bank)
            self.bank_conf_var.set(f"{conf:.2f}")
            chosen = (self.bank_mode.get() or "AUTO").upper()
            if chosen == "AUTO":
                return bank
            if chosen == "OTHER":
                return (self.bank_other_name.get().strip() or "OTHER")
            return chosen

        def _make_signature(bank: str, sheet_name: str, header_row: int, cols: List[str]) -> str:
            base = "|".join([bank, sheet_name or "", str(int(header_row or 0)), "|".join(cols or [])])
            return str(abs(hash(base)))

        def _load_profiles() -> Dict[str, Dict]:
            try:
                cfg = load_json(self.cfg_path)
                prof = cfg.get("profiles", {})
                return prof if isinstance(prof, dict) else {}
            except Exception:
                return {}

        def _save_profile(sig: str, mapping: MappingConfig) -> None:
            try:
                cfg = load_json(self.cfg_path) if os.path.exists(self.cfg_path) else {}
                if not isinstance(cfg, dict):
                    cfg = {}
                prof = cfg.get("profiles", {})
                if not isinstance(prof, dict):
                    prof = {}
                prof[sig] = mapping.to_dict()
                cfg["profiles"] = prof
                save_json(self.cfg_path, cfg)
            except Exception:
                pass



        # ===== BANK TEMPLATES (mỗi bank nhiều mẫu: VCB1/VCB2...) =====
        def _load_bank_templates_cfg() -> Dict[str, List[Dict[str, Any]]]:
            try:
                cfg = load_json(self.cfg_path) if os.path.exists(self.cfg_path) else {}
                bt = cfg.get("bank_templates", {})
                return bt if isinstance(bt, dict) else {}
            except Exception:
                return {}

        def _save_bank_templates_cfg(bank_templates: Dict[str, List[Dict[str, Any]]]) -> None:
            try:
                cfg = load_json(self.cfg_path) if os.path.exists(self.cfg_path) else {}
                if not isinstance(cfg, dict):
                    cfg = {}
                # merge-safe: giữ profiles và các key khác
                if not isinstance(cfg.get("profiles"), dict):
                    cfg["profiles"] = {}
                cfg["bank_templates"] = bank_templates
                # cũng lưu mapping hiện tại/owner... (để UI chính không lệch)
                cfg.setdefault("mapping", self.mapping.to_dict())
                save_json(self.cfg_path, cfg)
            except Exception:
                pass

        def _norm_cols(cols: List[str]) -> List[str]:
            out = []
            for c in (cols or []):
                s = str(c).replace("\n", " ").strip().upper()
                s = " ".join(s.split())
                out.append(s)
            return out

        def _jaccard(a: List[str], b: List[str]) -> float:
            sa, sb = set(a or []), set(b or [])
            if not sa and not sb:
                return 0.0
            inter = len(sa & sb)
            uni = len(sa | sb) or 1
            return float(inter / uni)

        def _template_score(tpl: Dict[str, Any], bank_eff: str, sheet_name: str, header_row: int, cols_raw: List[str]) -> float:
            """Mức 'vừa': chủ yếu theo similarity cột, sheet/header chỉ là bonus."""
            hint = (tpl or {}).get("hint", {}) or {}
            cols_t = hint.get("cols_norm", []) or []
            sim = _jaccard(_norm_cols(cols_raw), cols_t)

            sheet_bonus = 1.0 if (hint.get("sheet_name", "") or "") == (sheet_name or "") and sheet_name != "" else 0.0
            hr_t = int(hint.get("header_row", -999) or -999)
            hr = int(header_row or 0)
            if hr_t == hr:
                header_bonus = 1.0
            elif abs(hr_t - hr) <= 2:
                header_bonus = 0.6
            else:
                header_bonus = 0.0

            # score tổng (vừa): 75% columns + 15% sheet + 10% header
            return float(0.75 * sim + 0.15 * sheet_bonus + 0.10 * header_bonus)

        def _get_templates_for_bank(bank_eff: str) -> List[Dict[str, Any]]:
            bt = _load_bank_templates_cfg()
            arr = bt.get((bank_eff or "OTHER").upper(), [])
            return arr if isinstance(arr, list) else []

        def _refresh_template_dropdown(bank_eff: str) -> None:
            tpls = _get_templates_for_bank(bank_eff)
            # hiển thị theo name (VCB1/VCB2...), fallback id
            names = []
            for t in tpls:
                nm = (t or {}).get("name") or (t or {}).get("id") or ""
                if nm:
                    names.append(nm)
            cb_tmpl["values"] = names
            if names:
                # giữ giá trị nếu còn tồn tại
                if template_var.get() not in names:
                    template_var.set(names[0])
            else:
                template_var.set("")

        def _find_template_by_name(bank_eff: str, name: str) -> Optional[Dict[str, Any]]:
            for t in _get_templates_for_bank(bank_eff):
                if (t.get("name") or t.get("id")) == name:
                    return t
            return None

        def _apply_template(bank_eff: str, tpl: Dict[str, Any]) -> None:
            if not tpl:
                return
            mp = tpl.get("mapping", {}) or {}
            mcfg = MappingConfig.from_dict(mp)
            set_combo_from_mapping(mcfg)
            self.log(f"Đã áp mẫu ngân hàng: {tpl.get('name') or tpl.get('id')}")

        def _apply_selected_template(bank_eff: str) -> None:
            name = (template_var.get() or "").strip()
            if not name:
                return
            tpl = _find_template_by_name(bank_eff, name)
            if not tpl:
                messagebox.showwarning("Không tìm thấy mẫu", "Mẫu đã chọn không tồn tại (có thể đã bị xóa).")
                _refresh_template_dropdown(bank_eff)
                return
            _apply_template(bank_eff, tpl)

        def _next_template_name(bank_eff: str) -> str:
            # format mặc định: VCB1/VCB2...
            bank = (bank_eff or "OTHER").upper()
            tpls = _get_templates_for_bank(bank)
            used = set()
            for t in tpls:
                nm = (t.get("name") or "")
                m = re.fullmatch(rf"{re.escape(bank)}(\d+)", nm.strip().upper())
                if m:
                    used.add(int(m.group(1)))
            k = 1
            while k in used:
                k += 1
            return f"{bank}{k}"

        def _save_current_as_template(bank_eff: str, cols_raw: List[str], sheet_name: str, header_row: int) -> None:
            bank = (bank_eff or "OTHER").upper()
            # tên mặc định VCB1/VCB2...
            default_name = _next_template_name(bank)
            name = simpledialog.askstring("Lưu mẫu ngân hàng", f"Đặt tên mẫu (ví dụ {default_name}):", initialvalue=default_name, parent=w)
            if not name:
                return
            name = name.strip()
            if not name:
                return

            bt = _load_bank_templates_cfg()
            arr = bt.get(bank, [])
            if not isinstance(arr, list):
                arr = []

            # tránh trùng tên
            for t in arr:
                if (t.get("name") or "").strip().upper() == name.upper():
                    messagebox.showerror("Trùng tên", f"Đã có mẫu '{name}'. Hãy chọn tên khác.")
                    return

            tpl = {
                "id": f"{bank}_{int(datetime.datetime.now().timestamp())}",
                "name": name,
                "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "hint": {
                    "sheet_name": sheet_name or "",
                    "header_row": int(header_row or 0),
                    "cols_norm": _norm_cols(cols_raw),
                },
                "mapping": self.mapping.to_dict(),
            }
            arr.append(tpl)
            bt[bank] = arr
            _save_bank_templates_cfg(bt)
            self.log(f"Đã lưu mẫu ngân hàng: {name}")
            _refresh_template_dropdown(bank)
            template_var.set(name)

        def _delete_selected_template(bank_eff: str) -> None:
            bank = (bank_eff or "OTHER").upper()
            name = (template_var.get() or "").strip()
            if not name:
                return
            if not messagebox.askyesno("Xóa mẫu", f"Xóa mẫu '{name}' của bank {bank}?"):
                return
            bt = _load_bank_templates_cfg()
            arr = bt.get(bank, [])
            if not isinstance(arr, list):
                return
            arr2 = [t for t in arr if (t.get("name") or t.get("id")) != name]
            bt[bank] = arr2
            _save_bank_templates_cfg(bt)
            self.log(f"Đã xóa mẫu: {name}")
            _refresh_template_dropdown(bank)

        def _match_best_template(bank_eff: str, sheet_name: str, header_row: int, cols_raw: List[str]) -> Tuple[Optional[Dict[str, Any]], float]:
            best_tpl = None
            best_sc = 0.0
            for t in _get_templates_for_bank(bank_eff):
                sc = _template_score(t, bank_eff, sheet_name, header_row, cols_raw)
                if sc > best_sc:
                    best_sc, best_tpl = sc, t
            return best_tpl, float(best_sc)

        # wire buttons (we need current bank_eff at runtime, so commands are set inside reload_all)
        btn_apply_tmpl.configure(command=lambda: None)
        btn_save_tmpl.configure(command=lambda: None)
        btn_del_tmpl.configure(command=lambda: None)



        def _auto_drop_rows(dfp: pd.DataFrame, mcfg: MappingConfig) -> Tuple[pd.DataFrame, int]:
            """Áp dụng bỏ qua dòng đầu + lọc dòng rác theo mức đã chọn.
            Trả về (df_sau_loc, so_dong_bi_loai).
            """
            if dfp is None or dfp.empty:
                return dfp, 0

            # 1) skip top N
            nskip = int(skip_top_var.get() or 0)
            if nskip > 0 and len(dfp) > nskip:
                dfp2 = dfp.iloc[nskip:].copy()
            else:
                dfp2 = dfp.copy()

            # 2) lọc theo 4 mức (dựa trên mapping hiện tại/suggest/profile)
            level = (filter_level_var.get() or "CHUAN").upper().strip()
            kept, dropped = filter_garbage_rows(dfp2, mapping=mcfg, level=level)
            return kept, (0 if dropped is None else len(dropped))

        _reload_job = {"id": None}

        
        def reload_all(force: bool = False):
            sheet_sel = sheet_var.get()
            sheet_name = "" if sheet_sel == "(Sheet đầu tiên)" else sheet_sel

            try:
                dfp_raw = self._read_preview(first, sheet_name=sheet_name, header_row=header_var.get(), nrows=220)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không đọc được preview: {e}")
                return

            # bank + signature dựa trên header/cols (không phụ thuộc lọc dòng)
            cols_raw = list(dfp_raw.columns) if dfp_raw is not None else []
            bank_eff = effective_bank(cols_raw)
            _refresh_template_dropdown(bank_eff)
            # set template button actions for current bank
            btn_apply_tmpl.configure(command=lambda b=bank_eff: _apply_selected_template(b))
            btn_save_tmpl.configure(command=lambda b=bank_eff, c=cols_raw, sn=sheet_name, hr=header_var.get(): _save_current_as_template(b, c, sn, hr))
            btn_del_tmpl.configure(command=lambda b=bank_eff: _delete_selected_template(b))

            sig = _make_signature(bank_eff, sheet_name, header_var.get(), cols_raw)
            state["sig"] = sig

            profiles = _load_profiles()

            # 1) PROFILE nếu match signature
            m_profile = MappingConfig.from_dict(profiles[sig]) if sig in profiles else None

            # 2) mapping đã lưu (global) nếu có (ưu tiên hơn suggest để UI ổn định)
            m_saved = self.mapping if getattr(self.mapping, "content_col", "") or getattr(self.mapping, "date_col", "") or getattr(self.mapping, "datetime_col", "") else None

            # 3) suggest luôn được tính để gợi ý, nhưng KHÔNG tự động đè lên lựa chọn người dùng
            m_suggest = suggest_mapping(dfp_raw, bank_eff, header_var.get(), sheet_name=sheet_name)
            state["suggest"] = m_suggest

            if m_profile is not None:
                mcfg_tmp = m_profile
                map_src = "PROFILE"
            else:
                # TEMPLATE theo bank (mức vừa) nếu không có profile strict
                best_tpl, best_sc = _match_best_template(bank_eff, sheet_name, header_var.get(), cols_raw)
                if best_tpl is not None and best_sc >= 0.60:
                    mcfg_tmp = MappingConfig.from_dict(best_tpl.get("mapping", {}))
                    nm = best_tpl.get("name") or best_tpl.get("id") or "TEMPLATE"
                    map_src = f"TEMPLATE:{nm} ({best_sc:.2f})"
                    try:
                        template_var.set(nm)
                    except Exception:
                        pass
                elif m_saved is not None:
                    mcfg_tmp = m_saved
                    map_src = "ĐÃ LƯU"
                else:
                    mcfg_tmp = m_suggest
                    map_src = "SUGGEST"

            # lọc dòng rác theo mapping tạm + mức lọc
            dfp, dropped_n = _auto_drop_rows(dfp_raw, mcfg_tmp)

            state["preview"] = dfp
            state["cols"] = list(dfp.columns) if dfp is not None else []
            set_combo_values(state["cols"])
            set_combo_from_mapping(mcfg_tmp)

            info_var.set(f"Mapping: {map_src} | Bank: {bank_eff} | Sig: {sig[:8]} | Loại: {dropped_n} dòng | (Có gợi ý)")

            render_preview(dfp)

        def schedule_reload():
            if _reload_job["id"]:
                try:
                    w.after_cancel(_reload_job["id"])
                except Exception:
                    pass
            _reload_job["id"] = w.after(150, lambda: reload_all(force=True))

        
        def apply_suggest():
            ms = state.get("suggest")
            if ms is None:
                return
            set_combo_from_mapping(ms)
            self.log("Đã áp gợi ý mapping (chưa lưu).")

        def apply():
            sheet_sel = sheet_var.get()
            self.mapping.sheet_name = "" if sheet_sel == "(Sheet đầu tiên)" else sheet_sel
            self.mapping.header_row = int(header_var.get())
            self.mapping.garbage_filter_level = (filter_level_var.get() or "CHUAN").upper().strip()

            def val(key):
                v = combos[key][0].get()
                return "" if v == "Không chọn" else v

            for key in combos.keys():
                setattr(self.mapping, key, val(key))

            if not self.mapping.content_col:
                messagebox.showerror("Thiếu mapping", "Phải chọn cột Nội dung.")
                return
            if not (self.mapping.date_col or self.mapping.datetime_col):
                messagebox.showerror("Thiếu mapping", "Phải chọn cột Ngày hoặc Ngày giờ.")
                return
            if (not self.mapping.debit_col or not self.mapping.credit_col) and (not self.mapping.amount_col and not self.mapping.direction_col):
                messagebox.showerror("Thiếu mapping", "Nếu không có Debit/Credit thì phải có Số tiền hoặc Loại giao dịch.")
                return

            self._save_cfg()
            self._update_mapping_summary()

            if state.get("sig"):
                _save_profile(state["sig"], self.mapping)

            self.log("Đã lưu mapping (kèm profile signature).")
            w.destroy()

        ttk.Button(bottom, text="🔄 Reload", command=lambda: reload_all(force=True)).pack(side="left")
        ttk.Button(bottom, text="✨ Áp gợi ý", command=lambda: apply_suggest()).pack(side="left", padx=(6, 0))
        ttk.Button(bottom, text="Hủy", command=w.destroy).pack(side="right")
        ttk.Button(bottom, text="✅ Áp dụng", command=apply).pack(side="right", padx=(6, 0))

        # event bindings -> reload
        sheet_var.trace_add("write", lambda *_: schedule_reload())
        header_var.trace_add("write", lambda *_: schedule_reload())
        skip_top_var.trace_add("write", lambda *_: schedule_reload())
        filter_level_var.trace_add("write", lambda *_: schedule_reload())

        reload_all(force=True)

    def _refresh_rule_tree(self)->None:
        for i in self.rule_tv.get_children(): self.rule_tv.delete(i)
        rules=sorted(self.tag_rules,key=lambda r:r.priority)
        for r in rules:
            req="/".join(["SĐT" if r.require_phone else "","Ví" if r.require_wallet else "","Biển số" if r.require_plate else "","Mã DV" if r.require_utility else ""]).replace("//","/").strip("/")
            on="✔" if r.enabled else "✖"
            typ="Contains" if r.match_type.lower()=="contains" else "Regex"
            self.rule_tv.insert("", "end", values=(r.priority,on,r.name,typ,(r.direction or "BOTH").upper(),req,", ".join(r.include or []),", ".join(r.exclude or []),f"{float(r.confidence or 0.8):.2f}"))

    def _selected_rule(self)->Optional[TagRule]:
        sel=self.rule_tv.selection()
        if not sel: return None
        vals=self.rule_tv.item(sel[0]).get("values",[])
        if not vals: return None
        pri=int(vals[0]); name=str(vals[2])
        for r in self.tag_rules:
            if r.priority==pri and r.name==name: return r
        return None

    def _rule_dialog(self,title:str,rule:Optional[TagRule]=None)->Optional[TagRule]:
        w=tk.Toplevel(self); w.title(title); w.grab_set(); w.geometry("820x480")
        name=tk.StringVar(value=(rule.name if rule else ""))
        include=tk.StringVar(value=", ".join(rule.include) if rule else "")
        exclude=tk.StringVar(value=", ".join(rule.exclude) if rule else "")
        pr=tk.StringVar(value=str(rule.priority if rule else 50))
        match=tk.StringVar(value=(rule.match_type if rule else "contains"))
        enabled=tk.BooleanVar(value=(rule.enabled if rule else True))
        conf=tk.StringVar(value=str(rule.confidence if rule else 0.8))
        direction=tk.StringVar(value=(rule.direction if rule else "BOTH"))
        req_phone=tk.BooleanVar(value=(rule.require_phone if rule else False))
        req_wallet=tk.BooleanVar(value=(rule.require_wallet if rule else False))
        req_plate=tk.BooleanVar(value=(rule.require_plate if rule else False))
        req_utility=tk.BooleanVar(value=(rule.require_utility if rule else False))

        frm=ttk.Frame(w); frm.pack(fill="both",expand=True,padx=12,pady=12)
        frm.columnconfigure(1,weight=1)
        def row(r,label,widget):
            ttk.Label(frm,text=label,width=16).grid(row=r,column=0,sticky="w",pady=4)
            widget.grid(row=r,column=1,sticky="ew",pady=4)
        row(0,"Tên tag",ttk.Entry(frm,textvariable=name))
        row(1,"Từ khoá",ttk.Entry(frm,textvariable=include))
        row(2,"Loại trừ",ttk.Entry(frm,textvariable=exclude))
        row(3,"Ưu tiên",ttk.Entry(frm,textvariable=pr,width=10))
        row(4,"Kiểu match",ttk.Combobox(frm,textvariable=match,values=["contains","regex"],state="readonly"))
        row(5,"Chiều tiền",ttk.Combobox(frm,textvariable=direction,values=["BOTH","DEBIT","CREDIT"],state="readonly"))
        row(6,"Tin cậy (0-1)",ttk.Entry(frm,textvariable=conf,width=10))
        ttk.Checkbutton(frm,text="Bật rule",variable=enabled).grid(row=7,column=1,sticky="w",pady=4)
        ttk.Label(frm,text="REQ (điều kiện):").grid(row=8,column=0,sticky="w",pady=4)
        box=ttk.Frame(frm); box.grid(row=8,column=1,sticky="w",pady=4)
        ttk.Checkbutton(box,text="SĐT",variable=req_phone).pack(side="left")
        ttk.Checkbutton(box,text="Ví",variable=req_wallet).pack(side="left")
        ttk.Checkbutton(box,text="Biển số",variable=req_plate).pack(side="left")
        ttk.Checkbutton(box,text="Mã DV",variable=req_utility).pack(side="left")
        ttk.Label(frm,text="Gợi ý: Include/Exclude cách nhau bằng dấu phẩy. Regex dùng khi cần pattern.",foreground="gray").grid(row=9,column=0,columnspan=2,sticky="w",pady=(8,0))

        out={"rule":None}
        def ok():
            nm=name.get().strip()
            if not nm:
                messagebox.showerror("Lỗi","Tên tag không được trống."); return
            inc=[x.strip() for x in include.get().split(",") if x.strip()]
            exc=[x.strip() for x in exclude.get().split(",") if x.strip()]
            try: pri_i=int(pr.get().strip())
            except Exception: pri_i=50
            try: cf=float(conf.get().strip())
            except Exception: cf=0.8
            out["rule"]=TagRule(
                name=nm, match_type=match.get(), include=inc, exclude=exc, enabled=bool(enabled.get()),
                priority=pri_i, confidence=cf,
                require_phone=req_phone.get(), require_wallet=req_wallet.get(), require_plate=req_plate.get(), require_utility=req_utility.get(),
                direction=direction.get()
            )
            w.destroy()
        bottom=ttk.Frame(w); bottom.pack(fill="x",padx=12,pady=12)
        ttk.Button(bottom,text="OK",command=ok).pack(side="left")
        ttk.Button(bottom,text="Cancel",command=w.destroy).pack(side="right")
        w.wait_window()
        return out["rule"]

    def add_rule(self)->None:
        r=self._rule_dialog("Thêm rule")
        if r:
            self.tag_rules.append(r); self._refresh_rule_tree(); self._save_cfg()
    def edit_rule(self)->None:
        r0=self._selected_rule()
        if not r0: return
        r1=self._rule_dialog("Sửa rule", r0)
        if r1:
            for i,rr in enumerate(self.tag_rules):
                if rr.name==r0.name and rr.priority==r0.priority:
                    self.tag_rules[i]=r1; break
            self._refresh_rule_tree(); self._save_cfg()
    def delete_rule(self)->None:
        r0=self._selected_rule()
        if not r0: return
        if not messagebox.askyesno("Xóa", f"Xóa rule '{r0.name}'?"): return
        self.tag_rules=[r for r in self.tag_rules if not (r.name==r0.name and r.priority==r0.priority)]
        self._refresh_rule_tree(); self._save_cfg()
    def toggle_rule(self)->None:
        r0=self._selected_rule()
        if not r0: return
        r0.enabled=not r0.enabled
        self._refresh_rule_tree(); self._save_cfg()

    def _effective_bank(self)->str:
        chosen=(self.bank_mode.get() or "AUTO").upper()
        if chosen=="AUTO": return (self.bank_auto_var.get() or "OTHER").upper()
        if chosen=="OTHER": return (self.bank_other_name.get().strip() or "OTHER")
        return chosen

    def analyze_thread(self)->None:
        if not self.selected_files:
            messagebox.showwarning("Thiếu file","Hãy chọn file."); return
        if not self.mapping.content_col:
            messagebox.showwarning("Thiếu mapping","Hãy vào Mapping để chọn cột."); return
        threading.Thread(target=self._analyze_impl, daemon=True).start()

    def _analyze_impl(self)->None:
        try:
            self.after(0, lambda: self._set_progress(0,"Bắt đầu"))
            self.log("Bắt đầu phân tích...")
            self.owner.account=self.owner_acc_var.get().strip()
            self.owner.name=self.owner_name_var.get().strip()
            self.owner.bank=self._effective_bank()
            try: self.cp_cfg.min_confidence_for_stats=float(self.min_conf_var.get().strip())
            except Exception: pass
            self._sync_suspicious_from_ui()
            self._save_cfg()
            pipeline=DeepPipeline(self.mapping,self.owner,self.cp_cfg,self.tag_rules,self.susp_cfg)
            mode=(self.mode_money.get() or "BOTH").upper()
            def pcb(pct:int,msg:str):
                self.after(0, lambda: self._set_progress(pct,msg))
            self.reports=pipeline.run(self.selected_files, mode=mode, forensic_raw=self.show_forensic_raw.get(), progress_cb=pcb)
            self.after(0,self._update_summary_vars)
            self.after(0,self._render_all)
            self.after(0, lambda: self.status_var.set(f"Hoàn tất ({mode})"))
            self.log("Hoàn tất.")
        except Exception as e:
            self.after(0, lambda: self.status_var.set("Lỗi"))
            self.log(f"ERROR: {e}")
            self.after(0, lambda: messagebox.showerror("Lỗi", str(e)))
        finally:
            self.after(0, lambda: self.progress.configure(value=0))

    def _update_summary_vars(self)->None:
        raw=self.reports.get("RAW_FULL", pd.DataFrame())
        if raw is None or raw.empty:
            for v in [self.sum_tx,self.sum_in,self.sum_out,self.sum_net,self.sum_cp,self.sum_susp,self.sum_issues]: v.set("0")
            return
        total=len(raw)
        sum_in=float(raw["signed_amount"][raw["signed_amount"]>0].sum()) if "signed_amount" in raw.columns else 0.0
        sum_out=float((-raw["signed_amount"][raw["signed_amount"]<0]).sum()) if "signed_amount" in raw.columns else 0.0
        net=float(raw["signed_amount"].sum()) if "signed_amount" in raw.columns else 0.0
        cp_unique=int(raw["cp_key"].nunique()) if "cp_key" in raw.columns else 0
        susp=self.reports.get("SUSPICIOUS", pd.DataFrame())
        issues=self.reports.get("DATA_QUALITY", pd.DataFrame())
        self.sum_tx.set(f"{total:,}")
        self.sum_in.set(_fmt_money(sum_in))
        self.sum_out.set(_fmt_money(sum_out))
        self.sum_net.set(_fmt_money(net))
        self.sum_cp.set(f"{cp_unique:,}")
        self.sum_susp.set(f"{(0 if susp is None else len(susp)):,}")
        self.sum_issues.set(f"{(0 if issues is None else len(issues)):,}")

    def _render_all(self)->None:
        if not self.reports: return
        self._rerender_raw_only()
        for key in ["COUNTERPARTY","TAG","TAG_ENTITY","SUSPICIOUS","DATA_QUALITY"]:
            self._render_tree(self.tables[key], self.reports.get(key, pd.DataFrame()), key)

    def _rerender_raw_only(self)->None:
        if not self.reports: return
        df=self.reports.get("RAW_FORENSIC" if self.show_forensic_raw.get() else "RAW_CLEAN", pd.DataFrame())
        self._render_tree(self.tables["RAW"], df, "RAW")

    def _render_tree(self, tree: ttk.Treeview, df: pd.DataFrame, key: str)->None:
        tree.delete(*tree.get_children())
        if df is None or df.empty:
            tree["columns"]=[]
            return
        limit=self.MAX_ROWS_RAW if key=="RAW" else self.MAX_ROWS_SUM
        view=df.head(limit).copy()
        cols=list(view.columns)
        tree["columns"]=cols; tree["show"]="headings"
        for c in cols:
            tree.heading(c,text=c,command=lambda col=c,t=tree:self._sort_tree(t,col,False))
            tree.column(c,width=150,anchor="w",stretch=True)
        for _,row in view.iterrows():
            tree.insert("", "end", values=[row.get(c,"") for c in cols])

    def _sort_tree(self, tree: ttk.Treeview, col: str, reverse: bool)->None:
        data=[(tree.set(k,col),k) for k in tree.get_children("")]
        def to_num(x):
            try: return float(str(x).replace(",",""))
            except Exception: return None
        nums=[to_num(v) for v,_ in data]
        if data and all(v is not None for v in nums):
            pairs=list(zip(nums,[k for _,k in data])); pairs.sort(reverse=reverse,key=lambda t:t[0])
            data=[(v,k) for v,k in pairs]
        else:
            data.sort(reverse=reverse,key=lambda t:str(t[0]))
        for i,(_,k) in enumerate(data): tree.move(k,"",i)
        tree.heading(col, command=lambda: self._sort_tree(tree,col,not reverse))

    def _current_key(self)->str:
        title=self.nb.tab(self.nb.select(),"text")
        for k,v in self.tab_map.items():
            if v==title: return k
        return "RAW"

    def _get_current_df(self)->pd.DataFrame:
        if not self.reports: return pd.DataFrame()
        k=self._current_key()
        if k=="RAW":
            return self.reports.get("RAW_FORENSIC" if self.show_forensic_raw.get() else "RAW_CLEAN", pd.DataFrame())
        return self.reports.get(k, pd.DataFrame())

    def apply_search_filter(self)->None:
        if not self.reports: return
        k=self._current_key()
        base=self._get_current_df()
        if base is None or base.empty: return
        kw=(self.search_var.get() or "").strip().lower()
        flt=(self.filter_var.get() or "").strip().lower()
        df=base.copy()
        if flt:
            if "debit" in flt and "direction" in df.columns:
                df=df[df["direction"].astype(str).str.upper()=="DEBIT"]
            if "credit" in flt and "direction" in df.columns:
                df=df[df["direction"].astype(str).str.upper()=="CREDIT"]
            if "high" in flt and "Mức" in df.columns:
                df=df[df["Mức"].astype(str).str.upper()=="CAO"]
            if "medium" in flt and "Mức" in df.columns:
                df=df[df["Mức"].astype(str).str.upper().str.contains("TRUNG",na=False)]
            if flt.startswith("tag:"):
                tv=flt.split(":",1)[1].strip()
                for col in ["Tag","primary_tag"]:
                    if col in df.columns:
                        df=df[df[col].astype(str).str.lower().str.contains(tv,na=False)]
                        break
        if kw:
            mask=pd.Series(False,index=df.index)
            for c in df.columns:
                mask |= df[c].astype(str).str.lower().str.contains(kw,na=False)
            df=df[mask]
        self._render_tree(self.tables[k], df, k)

    def clear_search_filter(self)->None:
        self.search_var.set(""); self.filter_var.set("")
        if self.reports: self._render_all()

    def copy_selected(self,event=None)->None:
        tree=event.widget if event else self.tables.get(self._current_key())
        if tree is None: return
        sel=tree.selection()
        if not sel: return
        rows=[]
        for item in sel:
            rows.append("\t".join(map(str, tree.item(item)["values"])))
        text="\n".join(rows)
        self.clipboard_clear(); self.clipboard_append(text)
        self.status_var.set("Đã copy selection")

    def copy_all_current(self)->None:
        df=self._get_current_df()
        if df is None or df.empty: return
        text=df.to_csv(sep="\t", index=False)
        self.clipboard_clear(); self.clipboard_append(text)
        self.status_var.set("Đã copy all (TSV)")

    def export_excel(self)->None:
        if not self.reports:
            messagebox.showwarning("Chưa có dữ liệu","Hãy phân tích trước."); return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if not path: return
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                self.reports.get("RAW_CLEAN", pd.DataFrame()).to_excel(w, sheet_name="Giao_dich", index=False)
                self.reports.get("RAW_FORENSIC", pd.DataFrame()).to_excel(w, sheet_name="Giao_dich_forensic", index=False)
                self.reports.get("COUNTERPARTY", pd.DataFrame()).to_excel(w, sheet_name="Doi_ung", index=False)
                self.reports.get("TAG", pd.DataFrame()).to_excel(w, sheet_name="Theo_tag", index=False)
                self.reports.get("TAG_ENTITY", pd.DataFrame()).to_excel(w, sheet_name="Tag_Thuc_the", index=False)
                self.reports.get("SUSPICIOUS", pd.DataFrame()).to_excel(w, sheet_name="Nghi_van", index=False)
                self.reports.get("DATA_QUALITY", pd.DataFrame()).to_excel(w, sheet_name="Chat_luong_DL", index=False)
            self.status_var.set("Đã xuất Excel")
            self.log(f"Xuất Excel: {path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất", str(e))

    def export_pdf(self)->None:
        if not REPORTLAB_OK:
            messagebox.showinfo("Thiếu thư viện","Chưa có reportlab để xuất PDF."); return
        if not self.reports:
            messagebox.showwarning("Chưa có dữ liệu","Hãy phân tích trước."); return
        path=filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF","*.pdf")])
        if not path: return
        try:
            styles=getSampleStyleSheet()
            story=[]
            story.append(Paragraph("BÁO CÁO PHÂN TÍCH CHUYÊN SÂU SAO KÊ TKNH (TAB4)", styles["Title"]))
            story.append(Spacer(1,10))
            story.append(Paragraph(
                f"TX={self.sum_tx.get()} | IN={self.sum_in.get()} | OUT={self.sum_out.get()} | NET={self.sum_net.get()} | Đối ứng={self.sum_cp.get()} | Nghi vấn={self.sum_susp.get()} | Lỗi DL={self.sum_issues.get()}",
                styles["Normal"]
            ))
            story.append(Spacer(1,10))
            def add_table(title:str, df: pd.DataFrame, n:int, font_size:int=7):
                story.append(Paragraph(title, styles["Heading2"]))
                story.append(Spacer(1,6))
                if df is None or df.empty:
                    story.append(Paragraph("Không có dữ liệu.", styles["Normal"]))
                    story.append(Spacer(1,8))
                    return
                dd=df.head(n)
                t=Table([dd.columns.tolist()] + dd.values.tolist())
                t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("GRID",(0,0),(-1,-1),0.5,colors.grey),("FONTSIZE",(0,0),(-1,-1),font_size)]))
                story.append(t); story.append(Spacer(1,12))
            add_table("1) Đối ứng (Top)", self.reports.get("COUNTERPARTY"), 25, 7)
            add_table("2) Theo Tag (tổng)", self.reports.get("TAG"), 25, 7)
            add_table("3) Tag + Thực thể", self.reports.get("TAG_ENTITY"), 35, 6)
            add_table("4) Nghi vấn", self.reports.get("SUSPICIOUS"), 35, 6)
            add_table("5) Chất lượng dữ liệu", self.reports.get("DATA_QUALITY"), 35, 6)
            doc=SimpleDocTemplate(path, pagesize=landscape(A4))
            doc.build(story)
            self.status_var.set("Đã xuất PDF")
            self.log(f"Xuất PDF: {path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất PDF", str(e))
