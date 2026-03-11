
# deep_engine.py
from __future__ import annotations
import os, re, json
from dataclasses import dataclass, asdict
from typing import Any, Callable, Dict, List, Optional, Tuple
import numpy as np
import pandas as pd

ProgressCB = Optional[Callable[[int, str], None]]

WS_RE = re.compile(r"\s+")
PHONE_RE = re.compile(r"\b0\d{9,10}\b")
WALLET_RE = re.compile(r"\b(MOMO|ZALOPAY|VNPAY|SHOPEEPAY|VIETTEL\s*MONEY)\b", re.I)
PLATE_RE = re.compile(r"\b\d{2}[A-Z]-\d{3,5}\b")
UTILITY_RE = re.compile(r"\b(EVN\d+|KH\d+|HD\d+|ID\d+|FPT\d+|VNPT\d+|VTL\d+)\b", re.I)
ACC_TOKEN_RE = re.compile(r"\b[A-Z0-9][A-Z0-9\-]{5,34}\b", re.I)

KW_ACC = ["STK", "SỐ TK", "SO TK", "TK", "TÀI KHOẢN", "TAI KHOAN", "ACCOUNT", "ACCT", "A/C"]
KW_NAME = ["CTK", "CHỦ TK", "CHU TK", "TÊN", "TEN", "NAME"]
KW_BANK = ["NH", "NGÂN HÀNG", "NGAN HANG", "BANK"]

CREDIT_HINTS = ["nhan","nhận","thu","credit","incoming","chuyen den","chuyển đến","vao","vào","lai","lãi","hoan","hoàn","refund","salary","luong","lương","deposit"]
DEBIT_HINTS  = ["chuyen","chuyển","thanh toan","thanh toán","payment","nap","nạp","rut","rút","phi","phí","mua","purchase","tra","trả","topup","the cao","thẻ cào","withdraw"]

NOISE_PATS = [
    re.compile(r"\bREF\b\s*[:\-]?\s*\w+", re.I),
    re.compile(r"\bTRACE\b\s*[:\-]?\s*\w+", re.I),
    re.compile(r"\bFT\b\s*[:\-]?\s*\w+", re.I),
    re.compile(r"\bID\b\s*[:\-]?\s*\w+", re.I),
    re.compile(r"\bSTT\b\s*[:\-]?\s*\w+", re.I),
]

BANK_ALIASES = {
    "MB": ["MB", "MBBANK", "QUAN DOI", "QUÂN ĐỘI"],
    "TCB": ["TCB", "TECHCOMBANK", "TECHCOM"],
    "VCB": ["VCB", "VIETCOMBANK"],
    "BIDV": ["BIDV"],
    "VTB": ["VTB", "VIETINBANK", "VIETIN"],
    "VIB": ["VIB"],
    "STB": ["STB", "SACOMBANK"],
    "SHB": ["SHB"],
    "EXB": ["EXB", "EXIMBANK"],
}

def _norm_space(s: str) -> str:
    return WS_RE.sub(" ", (s or "").strip())

def _norm_token(s: str) -> str:
    s = (s or "").strip().upper()
    s = re.sub(r"\s+", "", s)
    return s

def _safe_float(x: Any) -> float:
    if x is None:
        return 0.0
    if isinstance(x, (int, float, np.number)):
        if isinstance(x, float) and np.isnan(x):
            return 0.0
        return float(x)
    s = str(x).strip()
    if not s:
        return 0.0
    s = s.replace(",", "").replace(" ", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except Exception:
        return 0.0

def save_json(path: str, data: Dict[str, Any]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

@dataclass
class OwnerInfo:
    bank: str = ""
    account: str = ""
    name: str = ""
    def to_dict(self): return asdict(self)
    @staticmethod
    def from_dict(d): return OwnerInfo(**(d or {}))

@dataclass
class MappingConfig:
    sheet_name: str = ""   # "" => sheet đầu tiên
    header_row: int = 0
    content_col: str = ""
    date_col: str = ""
    datetime_col: str = ""
    time_col: str = ""
    debit_col: str = ""
    credit_col: str = ""
    amount_col: str = ""
    direction_col: str = ""
    cp_account_col: str = ""
    cp_name_col: str = ""
    cp_bank_col: str = ""
    sender_acc_col: str = ""
    receiver_acc_col: str = ""
    # Mức lọc dòng rác: TAT / NHE / CHUAN / MANH
    garbage_filter_level: str = "CHUAN"

    def to_dict(self): return asdict(self)

    @staticmethod
    def from_dict(d):
        d = dict(d or {})
        d.setdefault("sheet_name", ""); d.setdefault("header_row", 0)
        for k in [
            "content_col","date_col","datetime_col","time_col","debit_col","credit_col","amount_col","direction_col",
            "cp_account_col","cp_name_col","cp_bank_col","sender_acc_col","receiver_acc_col",
        ]:
            d.setdefault(k, "")
        d.setdefault("garbage_filter_level", "CHUAN")
        return MappingConfig(**d)

@dataclass
class CounterpartyConfig:

    min_confidence_for_stats: float = 0.6
    max_candidates: int = 3
    def to_dict(self): return asdict(self)
    @staticmethod
    def from_dict(d):
        d = dict(d or {})
        d.setdefault("min_confidence_for_stats",0.6); d.setdefault("max_candidates",3)
        return CounterpartyConfig(**d)

@dataclass
class TagRule:
    name: str
    match_type: str = "contains"
    include: List[str] = None
    exclude: List[str] = None
    enabled: bool = True
    priority: int = 100
    confidence: float = 0.8
    require_phone: bool = False
    require_wallet: bool = False
    require_plate: bool = False
    require_utility: bool = False
    direction: str = "BOTH"
    def __post_init__(self):
        self.include = list(self.include or [])
        self.exclude = list(self.exclude or [])
    def to_dict(self): return asdict(self)
    @staticmethod
    def from_dict(d):
        d = dict(d or {})
        d.setdefault("name",""); d.setdefault("match_type","contains"); d.setdefault("include",[]); d.setdefault("exclude",[])
        d.setdefault("enabled",True); d.setdefault("priority",100); d.setdefault("confidence",0.8)
        d.setdefault("require_phone",False); d.setdefault("require_wallet",False); d.setdefault("require_plate",False); d.setdefault("require_utility",False)
        d.setdefault("direction","BOTH")
        return TagRule(**d)

@dataclass
class SuspiciousConfig:
    night_start_hour: int = 0
    night_end_hour: int = 4
    big_amount_threshold: float = 50_000_000.0
    zscore_threshold: float = 3.0
    burst_txn_count: int = 5
    split_repeat_count: int = 4
    night_score: int = 10
    big_amount_score: int = 20
    zscore_score: int = 20
    burst_score: int = 20
    split_score: int = 20
    low_max: int = 29
    medium_max: int = 59
    def to_dict(self): return asdict(self)
    @staticmethod
    def from_dict(d): return SuspiciousConfig(**(d or {}))

class BankDetector:
    @staticmethod
    def detect(cols: List[str]) -> Tuple[str,float]:
        up = " | ".join([str(c).upper() for c in cols])
        score = {k:0 for k in ["MB","TCB","VCB","BIDV","VTB","VIB","STB","SHB","EXB","OTHER"]}
        if "PHÁT SINH NỢ" in up or "PHAT SINH NO" in up: score["MB"] += 2
        if "PHÁT SINH CÓ" in up or "PHAT SINH CO" in up: score["MB"] += 2
        if "SỐ BÚT TOÁN" in up or "SO BUT TOAN" in up: score["MB"] += 1
        if "TECHCOM" in up or "TCB" in up: score["TCB"] += 2
        if "VIETCOMBANK" in up or "VCB" in up: score["VCB"] += 2
        if "BIDV" in up: score["BIDV"] += 2
        if "VIETINBANK" in up or "VTB" in up or "VIETIN" in up: score["VTB"] += 2
        if "VIB" in up: score["VIB"] += 2
        if "SACOMBANK" in up or "STB" in up: score["STB"] += 2
        if "SHB" in up: score["SHB"] += 2
        if "EXIMBANK" in up or "EXB" in up: score["EXB"] += 2
        b, best = max(score.items(), key=lambda kv: kv[1])
        if best <= 0: return ("OTHER",0.4)
        total = sum(score.values()) or 1
        conf = min(0.95, 0.55 + (best/max(1,total))*0.8)
        return (b,float(conf))

def _col_score(col_name: str, synonyms: List[str]) -> float:
    u = str(col_name).upper()
    s = 0.0
    for syn in synonyms:
        su = syn.upper()
        if su == u: s += 10.0
        elif su in u: s += 4.0
    return s

def _series_date_hit(series: pd.Series) -> float:
    if series is None or len(series)==0: return 0.0
    ss = series.dropna().astype(str).head(80)
    if ss.empty: return 0.0
    return float(ss.str.contains(r"(?:\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\d{8})", regex=True, na=False).mean())

def _series_time_hit(series: pd.Series) -> float:
    if series is None or len(series)==0: return 0.0
    ss = series.dropna().astype(str).head(80)
    if ss.empty: return 0.0
    return float(ss.str.contains(r"\b\d{1,2}:\d{2}(?:\:\d{2})?\b", regex=True, na=False).mean())

def _series_numeric_hit(series: pd.Series) -> float:
    if series is None or len(series)==0: return 0.0
    ss = series.dropna().head(120)
    if ss.empty: return 0.0
    vals = ss.apply(_safe_float)
    return float((vals != 0.0).mean())

def suggest_mapping(df_preview: pd.DataFrame, bank_code: str, header_row: int, sheet_name: str="") -> MappingConfig:
    cols = [str(c).strip() for c in df_preview.columns]

    # ---------- Heuristics ----------
    def _content_quality(series: pd.Series) -> float:
        """0..1: giống nội dung giao dịch (text dài, nhiều chữ, ít số)."""
        if series is None or len(series) == 0:
            return 0.0
        ss = series.dropna().astype(str).head(150)
        if ss.empty:
            return 0.0
        lens = ss.str.len().replace(0, 1)
        mean_len = float(lens.mean())

        alpha_cnt = ss.str.count(r"[A-Za-zÀ-ỹ]")
        digit_cnt = ss.str.count(r"\d")
        alpha_ratio = float((alpha_cnt / lens).mean())
        digit_ratio = float((digit_cnt / lens).mean())

        score = 0.0
        if mean_len >= 35:
            score += 0.45
        elif mean_len >= 20:
            score += 0.30
        elif mean_len >= 12:
            score += 0.18

        if alpha_ratio >= 0.45:
            score += 0.40
        elif alpha_ratio >= 0.30:
            score += 0.28
        elif alpha_ratio >= 0.20:
            score += 0.18

        if digit_ratio >= 0.60:
            score -= 0.25
        elif digit_ratio >= 0.40:
            score -= 0.15
        elif digit_ratio >= 0.25:
            score -= 0.05

        return float(min(1.0, max(0.0, score)))

    def _datetime_quality(series: pd.Series) -> float:
        """
        0..1: giống cột ngày/giờ thật sự.
        Parse theo format tường minh để không sinh warning dateutil.
        """
        if series is None or len(series) == 0:
            return 0.0
        ss = series.dropna().astype(str).head(120)
        if ss.empty:
            return 0.0

        # Extract date (+ optional time) substring trước để loại rác
        s2 = ss.str.extract(
            r"((?:\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\d{8})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)",
            expand=False
        ).fillna("")

        # Nếu extract rỗng nhiều thì fallback raw (nhưng vẫn parse theo nhóm format)
        raw = ss.where(s2.str.len() == 0, s2)

        # Chuẩn hoá whitespace
        raw = raw.str.replace(r"\s+", " ", regex=True).str.strip()

        dt = pd.Series(pd.NaT, index=raw.index)

        # 1) YYYYMMDD [HH:MM(:SS)]
        mask_ymd8 = raw.str.match(r"^\d{8}(\s+\d{1,2}:\d{2}(:\d{2})?)?$", na=False)
        if mask_ymd8.any():
            has_time = raw.loc[mask_ymd8].str.contains(r":", na=False)
            if has_time.any():
                dt.loc[mask_ymd8 & has_time] = pd.to_datetime(raw.loc[mask_ymd8 & has_time], format="%Y%m%d %H:%M", errors="coerce")
                # trường hợp có :SS
                mss = (mask_ymd8 & has_time) & raw.str.match(r"^\d{8}\s+\d{1,2}:\d{2}:\d{2}$", na=False)
                if mss.any():
                    dt.loc[mss] = pd.to_datetime(raw.loc[mss], format="%Y%m%d %H:%M:%S", errors="coerce")
            dt.loc[mask_ymd8 & ~has_time] = pd.to_datetime(raw.loc[mask_ymd8 & ~has_time], format="%Y%m%d", errors="coerce")

        # 2) YYYY-MM-DD [HH:MM(:SS)]
        mask_iso = raw.str.match(r"^\d{4}-\d{2}-\d{2}(\s+\d{1,2}:\d{2}(:\d{2})?)?$", na=False)
        if mask_iso.any():
            has_time = raw.loc[mask_iso].str.contains(r":", na=False)
            if has_time.any():
                dt.loc[mask_iso & has_time] = pd.to_datetime(raw.loc[mask_iso & has_time], format="%Y-%m-%d %H:%M", errors="coerce")
                mss = (mask_iso & has_time) & raw.str.match(r"^\d{4}-\d{2}-\d{2}\s+\d{1,2}:\d{2}:\d{2}$", na=False)
                if mss.any():
                    dt.loc[mss] = pd.to_datetime(raw.loc[mss], format="%Y-%m-%d %H:%M:%S", errors="coerce")
            dt.loc[mask_iso & ~has_time] = pd.to_datetime(raw.loc[mask_iso & ~has_time], format="%Y-%m-%d", errors="coerce")

        # 3) DD/MM/YYYY [HH:MM(:SS)]
        mask_dmy = raw.str.match(r"^\d{2}/\d{2}/\d{4}(\s+\d{1,2}:\d{2}(:\d{2})?)?$", na=False)
        if mask_dmy.any():
            has_time = raw.loc[mask_dmy].str.contains(r":", na=False)
            if has_time.any():
                dt.loc[mask_dmy & has_time] = pd.to_datetime(raw.loc[mask_dmy & has_time], format="%d/%m/%Y %H:%M", errors="coerce")
                mss = (mask_dmy & has_time) & raw.str.match(r"^\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}$", na=False)
                if mss.any():
                    dt.loc[mss] = pd.to_datetime(raw.loc[mss], format="%d/%m/%Y %H:%M:%S", errors="coerce")
            dt.loc[mask_dmy & ~has_time] = pd.to_datetime(raw.loc[mask_dmy & ~has_time], format="%d/%m/%Y", errors="coerce")

        ok = float(dt.notna().mean())
        return ok

    def _accountish_quality(series: pd.Series) -> float:
        """0..1: giống cột STK (alnum) hơn so với cột tiền."""
        if series is None or len(series) == 0:
            return 0.0
        ss = series.dropna().astype(str).head(120)
        if ss.empty:
            return 0.0

        hit = float(ss.str.contains(r"\b[A-Z0-9][A-Z0-9\-]{5,34}\b", case=False, regex=True, na=False).mean())
        vals = ss.apply(_safe_float)
        numeric_ratio = float((vals != 0.0).mean())
        score = hit - 0.30 * numeric_ratio
        return float(min(1.0, max(0.0, score)))

    # ---------- Synonyms (multi-bank) ----------
    syn = {
        "content_col": [
            "Nội dung","Diễn giải","Mô tả","Nội dung giao dịch","Nội dung GD","ND","NDGD","Nội dung GDV",
            "TRAN_RMKS","NARRATIVE","DESCRIPTION","DETAILS","DETAIL","REMARK","REMARKS","NARRATION",
            "Ghi chú","Ghi chu","Ghi chú giao dịch","Notes","Note","Thông tin","Thong tin","Chi tiết","Chi tiet",
            "Nội dung chi tiết/Transaction in detail"
        ],
        "date_col": ["Ngày","Ngày GD","Ngày giao dịch","Transaction Date","Posting Date","Value Date","TRAN_DATE","NGAY_GD","Ngày hạch toán"],
        "datetime_col": ["Ngày giờ","Ngày/giờ","Datetime","Date Time","Ngày giờ giao dịch","Ngày/giờ GD","Ngày giờ GD"],
        "time_col": ["Giờ","Time","Giờ GD","Giờ giao dịch"],
        "debit_col": ["Phát sinh nợ","Ghi nợ","Debit","Nợ","PS Nợ","PHAT_SINH_NO","DEBIT_AMT","OUT_AMT","Chi","Chi tiền"],
        "credit_col": ["Phát sinh có","Ghi có","Credit","Có","PS Có","PHAT_SINH_CO","CREDIT_AMT","IN_AMT","Thu","Thu tiền"],
        "amount_col": ["Số tiền","Amount","Giá trị","So tien","TRAN_AMT","AMOUNT","Số tiền GD","Giá trị GD"],
        "direction_col": ["Loại","Type","IN/OUT","Chi/Thu","D/C","Direction","LOAI_GD","Loại giao dịch"],
        "cp_account_col": ["Tài khoản đối ứng","Số tài khoản đối ứng","TK đối ứng","SO_TK_DOIUNG","COUNTERPARTY ACCOUNT","Account đối ứng","STK đối ứng"],
        "cp_name_col": ["Tên tài khoản đối ứng","Tên đối ứng","Khách hàng đối ứng","TEN_DOIUNG","Counterparty name","Tên KH đối ứng","CTK đối ứng"],
        "cp_bank_col": ["Ngân hàng đối ứng","NH đối ứng","TEN_NH_DOIUNG","Counterparty bank","Tên NH đối ứng"],
        "sender_acc_col": ["Số TK chuyển","SO_TK_CHUYEN","TK chuyển","Sender account","From account","TK gửi","Từ TK"],
        "receiver_acc_col": ["Số TK nhận","SO_TK_NHAN","TK nhận","Receiver account","To account","TK nhận tiền","Đến TK"],
    }

    # Bank-specific boost
    if bank_code == "MB":
        syn["debit_col"] = ["Phát sinh nợ","PS Nợ"] + syn["debit_col"]
        syn["credit_col"] = ["Phát sinh có","PS Có"] + syn["credit_col"]
    if bank_code == "TCB":
        syn["cp_account_col"] = ["Tài khoản đối ứng","TAI KHOAN DOI UNG"] + syn["cp_account_col"]
        syn["cp_bank_col"] = ["Ngân hàng đối ứng","NGAN HANG DOI UNG"] + syn["cp_bank_col"]
    if bank_code == "VCB":
        syn["cp_account_col"] = ["Số tài khoản đối ứng","SO TAI KHOAN DOI UNG"] + syn["cp_account_col"]
    if bank_code == "SHB":
        syn["sender_acc_col"] = ["SO_TK_CHUYEN","Số TK chuyển"] + syn["sender_acc_col"]
        syn["receiver_acc_col"] = ["SO_TK_NHAN","Số TK nhận"] + syn["receiver_acc_col"]

    def best(field: str, prefer: str="") -> str:
        best_c = ""
        best_s = -1.0
        for c in cols:
            s = _col_score(c, syn[field])

            # base signals (existing)
            if prefer == "date":
                s += _series_date_hit(df_preview[c]) * 8.0
            elif prefer == "time":
                s += _series_time_hit(df_preview[c]) * 6.0
            elif prefer == "num":
                s += _series_numeric_hit(df_preview[c]) * 6.0

            # stronger signals
            if field in ("date_col","datetime_col"):
                s += _datetime_quality(df_preview[c]) * 10.0

            if field == "time_col":
                s += _series_time_hit(df_preview[c]) * 4.0

            if field == "content_col":
                cq = _content_quality(df_preview[c])
                s += cq * 12.0
                s -= _series_numeric_hit(df_preview[c]) * 2.5
                s -= _datetime_quality(df_preview[c]) * 2.0

            if field in ("debit_col","credit_col","amount_col"):
                s += _series_numeric_hit(df_preview[c]) * 9.0
                s -= _content_quality(df_preview[c]) * 2.0

            if field == "cp_account_col":
                s += _accountish_quality(df_preview[c]) * 10.0
                s -= _datetime_quality(df_preview[c]) * 2.0

            if s > best_s:
                best_s = s
                best_c = c
        return best_c if best_s >= 6.0 else ""

    m = MappingConfig(sheet_name=sheet_name, header_row=int(header_row or 0))
    m.content_col = best("content_col")
    m.datetime_col = best("datetime_col","date")
    m.date_col = best("date_col","date")
    m.time_col = best("time_col","time")
    m.debit_col = best("debit_col","num")
    m.credit_col = best("credit_col","num")
    m.amount_col = best("amount_col","num")
    m.direction_col = best("direction_col")
    m.cp_account_col = best("cp_account_col")
    m.cp_name_col = best("cp_name_col")
    m.cp_bank_col = best("cp_bank_col")
    m.sender_acc_col = best("sender_acc_col")
    m.receiver_acc_col = best("receiver_acc_col")
    return m

class StatementLoader:
    def __init__(self, mapping: MappingConfig):
        self.m = mapping

    @staticmethod
    def list_sheets(path: str) -> List[str]:
        if path.lower().endswith((".xlsx",".xls")):
            try:
                return list(pd.ExcelFile(path).sheet_names)
            except Exception:
                return []
        return []

    def _read_one(self, path: str) -> pd.DataFrame:
        hr=int(self.m.header_row or 0)
        if path.lower().endswith(".csv"):
            df=pd.read_csv(path, header=hr)
        else:
            sheet = self.m.sheet_name
            sn = 0 if not sheet else sheet
            try:
                df=pd.read_excel(path, header=hr, sheet_name=sn)
            except Exception:
                df=pd.read_excel(path, header=hr, sheet_name=0)
                df["__sheet_fallback__"]=True
            df["__sheet_name__"]=str(sn) if isinstance(sn,str) else ""
        df.columns=[str(c).strip() for c in df.columns]
        df["__source_file__"]=os.path.basename(path)
        return df

    def load(self, paths: List[str]) -> pd.DataFrame:
        frames=[self._read_one(p) for p in paths]
        if not frames: return pd.DataFrame()
        return pd.concat(frames, ignore_index=True)


# =========================
# LỌC DÒNG RÁC (4 mức)
# =========================
GARBAGE_LEVELS = ["TAT", "NHE", "CHUAN", "MANH"]

_BAD_ROW_PAT = re.compile(
    r"SỐ DƯ|SO DU|BALANCE|TỔNG|TONG|CỘNG|CONG|OPENING|CLOSING|KẾT SỐ|KET SO|"
    r"SỐ DƯ CUỐI|SO DU CUOI|SỐ DƯ ĐẦU|SO DU DAU|TRANG\s*\d+|PAGE\s*\d+|NGÀY IN|PRINT DATE",
    re.I
)
_DATE_ANY_PAT = re.compile(
    r"(?:\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\b\d{8}\b)(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?",
    re.I
)
_TIME_ANY_PAT = re.compile(r"\b\d{1,2}:\d{2}(?::\d{2})?\b")

def _norm_colname_for_cmp(s: str) -> str:
    return re.sub(r"\s+","", (s or "").strip().upper())

def filter_garbage_rows(
    df: pd.DataFrame,
    mapping: Optional["MappingConfig"]=None,
    level: str="CHUAN"
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Lọc dòng rác theo 4 mức.

    - TAT  : chỉ bỏ dòng all-NaN
    - NHE  : + bỏ dòng gần như trống, + keyword tổng kết/số dư
    - CHUAN: + bỏ header lặp, + giữ dòng theo rule 2/3 tín hiệu (text/date/money)
    - MANH : như CHUAN nhưng rule 3/3

    Trả về: (df_keep, df_drop)
    """
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame(), pd.DataFrame()

    lv = (level or "CHUAN").strip().upper()
    if lv not in GARBAGE_LEVELS:
        lv = "CHUAN"

    df0 = df.copy()
    # 1) drop all-empty rows
    keep = ~df0.isna().all(axis=1)
    df1 = df0.loc[keep].copy()
    dropped = df0.loc[~keep].copy()

    if lv == "TAT":
        return df1, dropped

    # 2) drop nearly-empty rows
    try:
        empt_ratio = df1.isna().mean(axis=1)
        bad = empt_ratio >= 0.95
        if bad.any():
            dropped = pd.concat([dropped, df1.loc[bad]], ignore_index=True)
            df1 = df1.loc[~bad].copy()
    except Exception:
        pass

    # 3) keyword blacklist on whole row text
    try:
        row_text = df1.astype(str).fillna("").agg(" ".join, axis=1)
        bad = row_text.str.contains(_BAD_ROW_PAT, na=False)
        if bad.any():
            dropped = pd.concat([dropped, df1.loc[bad]], ignore_index=True)
            df1 = df1.loc[~bad].copy()
    except Exception:
        pass

    if lv == "NHE":
        return df1, dropped

    # 4) header-repeat detection (cell equals column name)
    try:
        col_norm = {c: _norm_colname_for_cmp(str(c)) for c in df1.columns}
        # count how many mapped-like columns are equal to their headers
        hit_cnt = pd.Series(0, index=df1.index)
        for c in df1.columns:
            cn = col_norm.get(c, "")
            if not cn:
                continue
            v = df1[c].astype(str).fillna("").map(_norm_colname_for_cmp)
            hit_cnt += (v == cn).astype(int)
        bad = hit_cnt >= 2  # header rows usually match >=2 cols
        if bad.any():
            dropped = pd.concat([dropped, df1.loc[bad]], ignore_index=True)
            df1 = df1.loc[~bad].copy()
    except Exception:
        pass

    # 5) signal-based rule using mapping (or fallback to generic)
    m = mapping or MappingConfig()
    # text signal
    if m.content_col and m.content_col in df1.columns:
        s = df1[m.content_col].astype(str).fillna("").str.strip()
        has_text = (s.str.len() >= 3) & (s.str.contains(r"[A-Za-zÀ-ỹ]", regex=True, na=False) | s.str.contains(r"\d", na=False))
        # NOTE: allow digits-only if content is actually codes, but length>=6
        has_text = has_text | (s.str.len() >= 6)
    else:
        # generic: any column has a long-ish text
        row_text = df1.astype(str).fillna("").agg(" ".join, axis=1)
        has_text = row_text.str.len() >= 8

    # date signal (prefer datetime_col then date_col)
    date_src = None
    if m.datetime_col and m.datetime_col in df1.columns:
        date_src = df1[m.datetime_col]
    elif m.date_col and m.date_col in df1.columns:
        date_src = df1[m.date_col]
    if date_src is not None:
        ds = date_src.astype(str).fillna("")
        has_date = ds.str.contains(_DATE_ANY_PAT, na=False)
    else:
        row_text = df1.astype(str).fillna("").agg(" ".join, axis=1)
        has_date = row_text.str.contains(_DATE_ANY_PAT, na=False)

    # money signal
    has_money = pd.Series(False, index=df1.index)
    if m.debit_col and m.debit_col in df1.columns:
        has_money |= df1[m.debit_col].apply(_safe_float).abs() > 0
    if m.credit_col and m.credit_col in df1.columns:
        has_money |= df1[m.credit_col].apply(_safe_float).abs() > 0
    if (not has_money.any()) and m.amount_col and m.amount_col in df1.columns:
        has_money |= df1[m.amount_col].apply(_safe_float).abs() > 0

    sig_cnt = has_text.astype(int) + has_date.astype(int) + has_money.astype(int)

    if lv == "MANH":
        bad = sig_cnt < 3
    else:  # CHUAN
        bad = sig_cnt < 2

    if bad.any():
        dropped = pd.concat([dropped, df1.loc[bad]], ignore_index=True)
        df1 = df1.loc[~bad].copy()

    return df1, dropped

def drop_garbage_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Giữ lại để tương thích: dùng mức CHUẨN mặc định."""
    kept, _ = filter_garbage_rows(df, mapping=None, level="CHUAN")
    return kept

class DateTimeNormalizer:
    @staticmethod
    def _extract_date_str(s: pd.Series) -> pd.Series:
        ss=s.astype(str).fillna("")
        return ss.str.extract(r"((?:\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\d{8}))", expand=False).fillna("")
    @staticmethod
    def _extract_time_str(s: pd.Series) -> pd.Series:
        ss = s.astype(str).fillna("")
        # Group ngoài để extract lấy được kết quả, group trong là non-capturing để không tạo match group thừa
        return ss.str.extract(r"((?:\b\d{1,2}:\d{2}(?:\:\d{2})?\b))", expand=False).fillna("")
    @staticmethod
    def parse(date_s: pd.Series, time_s: Optional[pd.Series]=None) -> pd.Series:
        if np.issubdtype(date_s.dtype, np.datetime64):
            dt=pd.to_datetime(date_s, errors="coerce")
        else:
            raw=date_s.astype(str).fillna("")
            dpart=DateTimeNormalizer._extract_date_str(raw)
            d_use=pd.Series(np.where(dpart.str.len()>0, dpart, raw), index=date_s.index).astype(str).str.strip()
            dt=pd.Series(pd.NaT, index=date_s.index)
            yyyymmdd=d_use.str.fullmatch(r"\d{8}", na=False)
            if yyyymmdd.any():
                dt.loc[yyyymmdd]=pd.to_datetime(d_use.loc[yyyymmdd], format="%Y%m%d", errors="coerce")
            iso=d_use.str.fullmatch(r"\d{4}-\d{2}-\d{2}", na=False)
            if iso.any():
                dt.loc[iso]=pd.to_datetime(d_use.loc[iso], format="%Y-%m-%d", errors="coerce")
            dmy=d_use.str.fullmatch(r"\d{2}/\d{2}/\d{4}", na=False)
            if dmy.any():
                dt.loc[dmy]=pd.to_datetime(d_use.loc[dmy], format="%d/%m/%Y", errors="coerce")
            rest=~(yyyymmdd|iso|dmy)
            if rest.any():
                dt.loc[rest]=pd.to_datetime(d_use.loc[rest], errors="coerce", dayfirst=True)
        if time_s is not None:
            tpart=DateTimeNormalizer._extract_time_str(time_s.astype(str).fillna(""))
            combo=dt.dt.strftime("%Y-%m-%d")+" "+tpart
            dt2=pd.to_datetime(combo, errors="coerce")
            dt=dt2.fillna(dt)
        return dt

class StatementNormalizer:
    def __init__(self, mapping: MappingConfig, owner: OwnerInfo):
        self.m=mapping; self.owner=owner or OwnerInfo()

    @staticmethod
    def clean_description(s: pd.Series) -> pd.Series:
        out=s.astype(str).fillna("").str.replace(WS_RE," ",regex=True).str.strip()
        for pat in NOISE_PATS:
            out=out.str.replace(pat,"",regex=True)
        return out.str.replace(WS_RE," ",regex=True).str.strip()

    @staticmethod
    def _infer_dir_from_text(desc: pd.Series):
        d=desc.astype(str).str.lower()
        credit=pd.Series(False,index=desc.index)
        debit=pd.Series(False,index=desc.index)
        for kw in CREDIT_HINTS: credit |= d.str.contains(re.escape(kw.lower()), na=False)
        for kw in DEBIT_HINTS: debit  |= d.str.contains(re.escape(kw.lower()), na=False)
        direction=np.where(credit & ~debit,"CREDIT", np.where(debit & ~credit,"DEBIT","UNKNOWN"))
        evidence=np.where(direction=="CREDIT","Keyword nhận/thu", np.where(direction=="DEBIT","Keyword chuyển/chi",""))
        return pd.Series(direction,index=desc.index), pd.Series(evidence,index=desc.index)

    def _infer_dir_from_sender_receiver(self, df: pd.DataFrame) -> Optional[pd.Series]:
        owner=_norm_token(self.owner.account)
        m=self.m
        if not owner or not (m.sender_acc_col and m.receiver_acc_col): return None
        if m.sender_acc_col not in df.columns or m.receiver_acc_col not in df.columns: return None
        send=df[m.sender_acc_col].astype(str).fillna("").map(_norm_token)
        recv=df[m.receiver_acc_col].astype(str).fillna("").map(_norm_token)
        return pd.Series(np.where(send==owner,"DEBIT", np.where(recv==owner,"CREDIT","UNKNOWN")), index=df.index)

    def normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty: return df
        m=self.m
        if not m.content_col or m.content_col not in df.columns:
            raise ValueError("Mapping thiếu cột Nội dung (content_col).")
        if not (m.date_col or m.datetime_col):
            raise ValueError("Mapping thiếu cột Ngày hoặc Ngày giờ.")
        out=df.copy()
        out["description_raw"]=out[m.content_col].astype(str).fillna("")
        out["description_clean"]=self.clean_description(out["description_raw"])

        if m.datetime_col and m.datetime_col in out.columns:
            dt=DateTimeNormalizer.parse(out[m.datetime_col], None)
        else:
            date_src=out[m.date_col] if (m.date_col and m.date_col in out.columns) else out[m.datetime_col]
            time_src=out[m.time_col] if (m.time_col and m.time_col in out.columns) else None
            dt=DateTimeNormalizer.parse(date_src, time_src)
        out["_dt"]=dt
        out["txn_date"]=out["_dt"].dt.strftime("%Y-%m-%d").fillna("")
        out["txn_time"]=out["_dt"].dt.strftime("%H:%M:%S").fillna("")

        has_dc=bool(m.debit_col and m.credit_col and m.debit_col in out.columns and m.credit_col in out.columns)
        out["data_issue_debit_credit_both"]=False
        out["data_issue_amount_missing"]=False

        if has_dc:
            out["debit_amount"]=out[m.debit_col].apply(_safe_float)
            out["credit_amount"]=out[m.credit_col].apply(_safe_float)
            both=(out["debit_amount"]>0) & (out["credit_amount"]>0)
            out.loc[both,"data_issue_debit_credit_both"]=True
            if both.any():
                keep_debit=out.loc[both,"debit_amount"]>=out.loc[both,"credit_amount"]
                out.loc[both & keep_debit,"credit_amount"]=0.0
                out.loc[both & ~keep_debit,"debit_amount"]=0.0
        else:
            out["debit_amount"]=0.0; out["credit_amount"]=0.0

        if out["debit_amount"].sum()==0 and out["credit_amount"].sum()==0:
            if m.amount_col and m.amount_col in out.columns:
                out["amount_raw"]=out[m.amount_col].apply(_safe_float).abs()
            else:
                out["amount_raw"]=0.0
                out["data_issue_amount_missing"]=True
        else:
            out["amount_raw"]=np.where(out["debit_amount"]>0,out["debit_amount"],out["credit_amount"]).astype(float)

        if has_dc and ((out["debit_amount"]>0).any() or (out["credit_amount"]>0).any()):
            out["direction"]=np.where(out["debit_amount"]>0,"DEBIT", np.where(out["credit_amount"]>0,"CREDIT","UNKNOWN"))
            out["direction_evidence"]="Cột Nợ/Có"
        else:
            sr=self._infer_dir_from_sender_receiver(out)
            if sr is not None and (sr!="UNKNOWN").any():
                out["direction"]=sr
                out["direction_evidence"]="TK chuyển/nhận"
            elif m.direction_col and m.direction_col in out.columns:
                s=out[m.direction_col].astype(str).fillna("").str.upper()
                out["direction"]=np.where(s.str.contains("IN")|s.str.contains("CREDIT")|s.str.contains("THU")|s.str.fullmatch("C",na=False),
                                          "CREDIT",
                                          np.where(s.str.contains("OUT")|s.str.contains("DEBIT")|s.str.contains("CHI")|s.str.fullmatch("D",na=False),
                                                   "DEBIT","UNKNOWN"))
                out["direction_evidence"]="Cột loại giao dịch"
            else:
                ddir,evid=self._infer_dir_from_text(out["description_clean"])
                out["direction"]=ddir
                out["direction_evidence"]=evid
            out["debit_amount"]=np.where(out["direction"]=="DEBIT",out["amount_raw"],0.0)
            out["credit_amount"]=np.where(out["direction"]=="CREDIT",out["amount_raw"],0.0)

        out["amount"]=np.where(out["direction"]=="DEBIT",out["debit_amount"],out["credit_amount"]).astype(float)
        out["signed_amount"]=out["credit_amount"]-out["debit_amount"]

        out["cp_account_mapped"]=out[m.cp_account_col].astype(str).fillna("").map(_norm_token) if (m.cp_account_col and m.cp_account_col in out.columns) else ""
        out["cp_name_mapped"]=out[m.cp_name_col].astype(str).fillna("").map(_norm_space) if (m.cp_name_col and m.cp_name_col in out.columns) else ""
        out["cp_bank_mapped"]=out[m.cp_bank_col].astype(str).fillna("").map(_norm_space) if (m.cp_bank_col and m.cp_bank_col in out.columns) else ""
        out["sender_acc_mapped"]=out[m.sender_acc_col].astype(str).fillna("").map(_norm_token) if (m.sender_acc_col and m.sender_acc_col in out.columns) else ""
        out["receiver_acc_mapped"]=out[m.receiver_acc_col].astype(str).fillna("").map(_norm_token) if (m.receiver_acc_col and m.receiver_acc_col in out.columns) else ""

        out["dq_date_parse_fail"]=out["_dt"].isna() | (out["txn_date"].astype(str).str.len()==0)
        out["dq_direction_unknown"]=(out["direction"]=="UNKNOWN")
        return out

class CounterpartyExtractor:
    def __init__(self, owner: OwnerInfo, cfg: CounterpartyConfig):
        self.owner=owner or OwnerInfo()
        self.cfg=cfg or CounterpartyConfig()
        self.owner_acc=_norm_token(self.owner.account)

    @staticmethod
    def _bank_hint(text: str) -> str:
        u=(text or "").upper()
        for code,aliases in BANK_ALIASES.items():
            for a in aliases:
                if a.upper() in u:
                    return code
        return ""

    @staticmethod
    def _near_score(desc_u: str, token: str) -> int:
        idx=desc_u.find(token)
        if idx<0: return 0
        win=desc_u[max(0,idx-50):min(len(desc_u),idx+50)]
        s=0
        for kw in KW_ACC:
            if kw.upper() in win: s+=3
        for kw in KW_BANK:
            if kw.upper() in win: s+=1
        for kw in KW_NAME:
            if kw.upper() in win: s+=1
        return s

    def _extract_candidates(self, desc: str):
        desc_u=(desc or "").upper()
        tokens=ACC_TOKEN_RE.findall(desc_u)
        after_kw=re.findall(r"(?:STK|SO TK|SỐ TK|TK|ACCOUNT|ACCT)\s*[:\-]?\s*([A-Z0-9\- ]{6,40})", desc_u)
        for t in after_kw:
            t2=re.sub(r"\s+","",t).strip().upper()
            if t2: tokens.append(t2)
        seen=set(); uniq=[]
        for t in tokens:
            t2=_norm_token(t)
            if not t2 or t2 in seen: continue
            seen.add(t2); uniq.append(t2)
        out=[]
        for t in uniq:
            if self.owner_acc and t==self.owner_acc: continue
            if PHONE_RE.fullmatch(t): continue
            if UTILITY_RE.fullmatch(t): continue
            if re.fullmatch(r"\d{8}", t): continue
            score=0
            if 8<=len(t)<=18: score+=2
            elif 6<=len(t)<=25: score+=1
            score += self._near_score(desc_u,t)
            bh=self._bank_hint(desc_u)
            if bh: score+=1
            out.append({"token":t,"score":int(score),"bank_hint":bh})
        out.sort(key=lambda x:x["score"], reverse=True)
        return out[: max(1,int(self.cfg.max_candidates or 3))]

    def extract(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty: return df
        out=df.copy()
        cp_acc=out.get("cp_account_mapped", pd.Series([""]*len(out), index=out.index)).astype(str).fillna("").map(_norm_token)
        cp_name=out.get("cp_name_mapped", pd.Series([""]*len(out), index=out.index)).astype(str).fillna("").map(_norm_space)
        cp_bank=out.get("cp_bank_mapped", pd.Series([""]*len(out), index=out.index)).astype(str).fillna("").map(_norm_space)
        sender=out.get("sender_acc_mapped", pd.Series([""]*len(out), index=out.index)).astype(str).fillna("").map(_norm_token)
        receiver=out.get("receiver_acc_mapped", pd.Series([""]*len(out), index=out.index)).astype(str).fillna("").map(_norm_token)
        desc=out["description_clean"].astype(str).fillna("")

        candidates_list=[]; conf_list=[]; chosen_list=[]; bank_list=[]
        for i,txt in enumerate(desc.tolist()):
            if cp_acc.iat[i]:
                chosen=cp_acc.iat[i]
                cands=[{"token":chosen,"score":10,"bank_hint":self._bank_hint(cp_bank.iat[i] or txt)}]
                conf=0.95
                bh=self._bank_hint(cp_bank.iat[i] or txt)
            else:
                chosen=""; conf=0.0; bh=self._bank_hint(txt)
                if self.owner_acc and sender.iat[i] and receiver.iat[i]:
                    if sender.iat[i]==self.owner_acc and receiver.iat[i]!=self.owner_acc:
                        chosen=receiver.iat[i]; conf=0.9
                    elif receiver.iat[i]==self.owner_acc and sender.iat[i]!=self.owner_acc:
                        chosen=sender.iat[i]; conf=0.9
                if not chosen:
                    cands=self._extract_candidates(txt)
                    if not cands:
                        chosen=""; conf=0.0
                    else:
                        chosen=cands[0]["token"]
                        top=cands[0]["score"]
                        conf=min(0.9, 0.35+(top/10.0)*0.65)
                else:
                    cands=[{"token":chosen,"score":9,"bank_hint":bh}]
            candidates_list.append(cands); conf_list.append(float(conf)); chosen_list.append(chosen); bank_list.append(cp_bank.iat[i] or bh)

        out["cp_candidates"]=candidates_list
        out["cp_confidence"]=conf_list
        out["cp_account"]=chosen_list
        out["cp_name"]=cp_name
        out["cp_bank"]=bank_list
        out["cp_key"]=np.where(out["cp_account"].astype(str).str.len()>0,out["cp_account"],out["description_clean"].str.slice(0,40))
        out["dq_cp_missing"]=out["cp_account"].astype(str).str.len().eq(0)
        out["dq_cp_low_conf"]=out["cp_confidence"].astype(float) < float(self.cfg.min_confidence_for_stats or 0.6)
        return out

class EntityExtractor:
    def extract(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty: return df
        out=df.copy()
        s=out["description_clean"].astype(str).fillna("")
        out["phones"]=s.str.findall(PHONE_RE)
        out["wallets"]=s.str.findall(WALLET_RE).apply(lambda xs:[x.upper() for x in xs] if isinstance(xs,list) else [])
        out["plates"]=s.str.findall(PLATE_RE)
        out["utilities"]=s.str.findall(UTILITY_RE).apply(lambda xs:[x.upper() for x in xs] if isinstance(xs,list) else [])
        out["phone_cnt"]=out["phones"].apply(len)
        out["wallet_cnt"]=out["wallets"].apply(len)
        out["plate_cnt"]=out["plates"].apply(len)
        out["utility_cnt"]=out["utilities"].apply(len)
        return out

class TagEngine:
    def __init__(self, rules: List[TagRule]):
        rr=[r for r in (rules or []) if r and r.enabled and r.name]
        self.rules=sorted(rr, key=lambda r:r.priority)

    @staticmethod
    def _contains_any(desc: pd.Series, kws: List[str]) -> pd.Series:
        pats=[re.escape(k.strip()) for k in kws if k and k.strip()]
        if not pats: return pd.Series(False,index=desc.index)
        pattern="(?:"+"|".join(pats)+")"
        return desc.str.contains(pattern, case=False, na=False, regex=True)

    @staticmethod
    def _match_regex(desc: pd.Series, pats: List[str]) -> pd.Series:
        pats=[p.strip() for p in pats if p and p.strip()]
        if not pats: return pd.Series(False,index=desc.index)
        pattern="(?:"+"|".join(pats)+")"
        return desc.str.contains(pattern, case=False, na=False, regex=True)

    def apply(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty: return df
        out=df.copy()
        out["tags"]=[[] for _ in range(len(out))]
        out["tag_confidence"]=0.0
        out["tag_reason"]=""
        desc=out["description_clean"].astype(str).fillna("")
        for rule in self.rules:
            if rule.match_type.lower()=="regex":
                mask=self._match_regex(desc, rule.include); reason="Regex"
            else:
                mask=self._contains_any(desc, rule.include); reason="Từ khóa"
            if rule.exclude:
                mask &= ~self._contains_any(desc, rule.exclude)
            dirf=(rule.direction or "BOTH").upper()
            if dirf in ("DEBIT","CREDIT"):
                mask &= (out["direction"].astype(str).str.upper()==dirf)
            if rule.require_phone: mask &= out.get("phone_cnt",0)>0
            if rule.require_wallet: mask &= out.get("wallet_cnt",0)>0
            if rule.require_plate: mask &= out.get("plate_cnt",0)>0
            if rule.require_utility: mask &= out.get("utility_cnt",0)>0
            if mask.any():
                out.loc[mask,"tags"]=out.loc[mask,"tags"].apply(lambda xs,n=rule.name: xs+[n])
                empty_primary=mask & (out["tag_confidence"]<=0)
                if empty_primary.any():
                    out.loc[empty_primary,"tag_confidence"]=float(rule.confidence or 0.8)
                    out.loc[empty_primary,"tag_reason"]=reason
        out["primary_tag"]=out["tags"].apply(lambda xs: xs[0] if xs else "")
        out["dq_tag_missing"]=out["primary_tag"].astype(str).str.len().eq(0)
        return out

class SuspiciousDetector:
    def __init__(self, cfg: SuspiciousConfig):
        self.cfg=cfg or SuspiciousConfig()
    def detect(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty: return df
        out=df.copy(); cfg=self.cfg
        out["risk_score"]=0
        out["risk_flags"]=[[] for _ in range(len(out))]
        hours=pd.to_datetime(out["_dt"], errors="coerce").dt.hour.fillna(-1).astype(int)
        night=(hours>=cfg.night_start_hour) & (hours<=cfg.night_end_hour)
        if night.any():
            out.loc[night,"risk_score"] += cfg.night_score
            out.loc[night,"risk_flags"]=out.loc[night,"risk_flags"].apply(lambda xs: xs+["Giao dịch ban đêm"])
        big=out["amount"].astype(float).abs()>=float(cfg.big_amount_threshold or 0)
        if big.any():
            out.loc[big,"risk_score"] += cfg.big_amount_score
            out.loc[big,"risk_flags"]=out.loc[big,"risk_flags"].apply(lambda xs: xs+[f"Số tiền lớn (≥{int(cfg.big_amount_threshold):,})"])
        a=out["amount"].astype(float).abs(); mu=a.mean(); sig=a.std(ddof=0)
        if sig and sig>0:
            z=(a-mu)/sig
            out["zscore_abs_amount"]=z
            zmask=z.abs()>=float(cfg.zscore_threshold or 3.0)
            if zmask.any():
                out.loc[zmask,"risk_score"] += cfg.zscore_score
                out.loc[zmask,"risk_flags"]=out.loc[zmask,"risk_flags"].apply(lambda xs: xs+[f"Bất thường z-score ≥ {cfg.zscore_threshold}"])
        else:
            out["zscore_abs_amount"]=0.0
        if "txn_date" in out.columns and "cp_key" in out.columns:
            g=out.groupby(["cp_key","txn_date"]).size().rename("cnt").reset_index()
            hot=g[g["cnt"]>=int(cfg.burst_txn_count or 5)][["cp_key","txn_date"]]
            if not hot.empty:
                hot_set=set(map(tuple, hot.values.tolist()))
                mask=out.apply(lambda r: (r["cp_key"],r["txn_date"]) in hot_set, axis=1).astype(bool)
                out.loc[mask,"risk_score"] += cfg.burst_score
                out.loc[mask,"risk_flags"]=out.loc[mask,"risk_flags"].apply(lambda xs: xs+[f"Nhiều GD cùng ngày (≥{cfg.burst_txn_count})"])
            out["_amt_round1k"]=(out["amount"].astype(float)/1000.0).round()*1000.0
            g2=out.groupby(["cp_key","txn_date","_amt_round1k"]).size().rename("cnt").reset_index()
            rep=g2[g2["cnt"]>=int(cfg.split_repeat_count or 4)][["cp_key","txn_date","_amt_round1k"]]
            if not rep.empty:
                rep_set=set(map(tuple, rep.values.tolist()))
                mask2=out.apply(lambda r: (r["cp_key"],r["txn_date"],r["_amt_round1k"]) in rep_set, axis=1).astype(bool)
                out.loc[mask2,"risk_score"] += cfg.split_score
                out.loc[mask2,"risk_flags"]=out.loc[mask2,"risk_flags"].apply(lambda xs: xs+[f"Nghi vấn lặp/chia nhỏ (≥{cfg.split_repeat_count})"])
        def level(score:int)->str:
            if score<=cfg.low_max: return "THẤP"
            if score<=cfg.medium_max: return "TRUNG BÌNH"
            return "CAO"
        out["risk_level"]=out["risk_score"].apply(level)
        out["risk_flags_text"]=out["risk_flags"].apply(lambda xs: "; ".join(xs) if isinstance(xs,list) else str(xs))
        return out

class ReportBuilder:
    @staticmethod
    def _sum_in(x: pd.Series)->float: return float(x[x>0].sum())
    @staticmethod
    def _sum_out(x: pd.Series)->float: return float((-x[x<0]).sum())
    @staticmethod
    def by_counterparty(df: pd.DataFrame, min_conf: float)->pd.DataFrame:
        if df.empty: return pd.DataFrame()
        dd=df.copy()
        if "cp_confidence" in dd.columns:
            dd=dd[dd["cp_confidence"].astype(float)>=float(min_conf)]
        if dd.empty:
            return pd.DataFrame(columns=["Đối ứng","Ngày bắt đầu","Ngày kết thúc","Số GD","IN","OUT","NET","Max","Ngày Max","Min","Ngày Min"])
        dd["abs_amount"]=dd["amount"].astype(float).abs()
        idx_max=dd.groupby("cp_key")["abs_amount"].idxmax()
        idx_min=dd.groupby("cp_key")["abs_amount"].idxmin()
        maxr=dd.loc[idx_max,["cp_key","abs_amount","txn_date"]].rename(columns={"cp_key":"Đối ứng","abs_amount":"Max","txn_date":"Ngày Max"})
        minr=dd.loc[idx_min,["cp_key","abs_amount","txn_date"]].rename(columns={"cp_key":"Đối ứng","abs_amount":"Min","txn_date":"Ngày Min"})
        agg=dd.groupby("cp_key").agg(**{
            "Ngày bắt đầu":("txn_date","min"),
            "Ngày kết thúc":("txn_date","max"),
            "Số GD":("signed_amount","count"),
            "IN":("signed_amount",ReportBuilder._sum_in),
            "OUT":("signed_amount",ReportBuilder._sum_out),
            "NET":("signed_amount","sum"),
        }).reset_index().rename(columns={"cp_key":"Đối ứng"})
        out=agg.merge(maxr,on="Đối ứng",how="left").merge(minr,on="Đối ứng",how="left")
        return out.sort_values(["Số GD","OUT"],ascending=[False,False]).reset_index(drop=True)
    @staticmethod
    def by_tag(df: pd.DataFrame)->pd.DataFrame:
        if df.empty: return pd.DataFrame()
        expl=df.copy().explode("tags")
        expl["Tag"]=expl["tags"].fillna("").astype(str)
        expl=expl[expl["Tag"]!=""]
        if expl.empty:
            return pd.DataFrame(columns=["Tag","Ngày bắt đầu","Ngày kết thúc","Số GD","IN","OUT","NET"])
        agg=expl.groupby("Tag").agg(**{
            "Ngày bắt đầu":("txn_date","min"),
            "Ngày kết thúc":("txn_date","max"),
            "Số GD":("signed_amount","count"),
            "IN":("signed_amount",ReportBuilder._sum_in),
            "OUT":("signed_amount",ReportBuilder._sum_out),
            "NET":("signed_amount","sum"),
        }).reset_index()
        return agg.sort_values(["OUT","Số GD"],ascending=[False,False]).reset_index(drop=True)
    @staticmethod
    def tag_entity(df: pd.DataFrame)->pd.DataFrame:
        if df.empty: return pd.DataFrame()
        expl=df.copy().explode("tags")
        expl["Tag"]=expl["tags"].fillna("").astype(str)
        expl=expl[expl["Tag"]!=""]
        if expl.empty:
            return pd.DataFrame(columns=["Tag","Loại","Giá trị","Ngày bắt đầu","Ngày kết thúc","Số GD","IN","OUT","NET","Max","Min"])
        rows=[]
        def add(entity_type:str,col:str):
            if col not in expl.columns: return
            tmp=expl[["Tag","txn_date","signed_amount","amount",col]].copy()
            tmp=tmp.explode(col).dropna(subset=[col]).rename(columns={col:"Giá trị"})
            if tmp.empty: return
            g=tmp.groupby(["Tag","Giá trị"]).agg(**{
                "Ngày bắt đầu":("txn_date","min"),
                "Ngày kết thúc":("txn_date","max"),
                "Số GD":("signed_amount","count"),
                "IN":("signed_amount",ReportBuilder._sum_in),
                "OUT":("signed_amount",ReportBuilder._sum_out),
                "NET":("signed_amount","sum"),
                "Max":("amount","max"),
                "Min":("amount","min"),
            }).reset_index()
            g.insert(1,"Loại",entity_type)
            rows.append(g)
        add("SĐT","phones"); add("Ví","wallets"); add("Biển số","plates"); add("Mã DV","utilities")
        if not rows:
            return pd.DataFrame(columns=["Tag","Loại","Giá trị","Ngày bắt đầu","Ngày kết thúc","Số GD","IN","OUT","NET","Max","Min"])
        out=pd.concat(rows,ignore_index=True)
        return out.sort_values(["Tag","Loại","Số GD","OUT"],ascending=[True,True,False,False]).reset_index(drop=True)
    @staticmethod
    def suspicious(df: pd.DataFrame)->pd.DataFrame:
        if df.empty: return pd.DataFrame()
        s=df[df["risk_level"].isin(["CAO","TRUNG BÌNH"])].copy()
        cols=["txn_date","txn_time","direction","amount","cp_account","cp_bank","cp_confidence","risk_level","risk_score","risk_flags_text","primary_tag","tag_reason","description_clean","__source_file__"]
        for c in cols:
            if c not in s.columns: s[c]=""
        rename={"txn_date":"Ngày","txn_time":"Giờ","direction":"Chiều","amount":"Số tiền","cp_account":"Đối ứng","cp_bank":"NH","cp_confidence":"Tin cậy",
                "risk_level":"Mức","risk_score":"Điểm","risk_flags_text":"Cờ nghi vấn","primary_tag":"Tag","tag_reason":"Lý do tag",
                "description_clean":"Nội dung","__source_file__":"File"}
        return s[cols].rename(columns=rename).sort_values(["Điểm","Số tiền"],ascending=[False,False]).reset_index(drop=True)
    @staticmethod
    def data_quality(df: pd.DataFrame)->pd.DataFrame:
        if df.empty: return pd.DataFrame()
        chunks=[]
        def add(mask_col:str, issue:str):
            if mask_col in df.columns:
                bad=df[df[mask_col]].copy()
                if not bad.empty:
                    bad["Vấn đề"]=issue
                    chunks.append(bad)
        add("dq_date_parse_fail","Không đọc được Ngày/Giờ")
        add("dq_direction_unknown","Không xác định được Debit/Credit")
        add("dq_cp_missing","Không bóc tách được đối ứng")
        add("dq_cp_low_conf","Đối ứng độ tin cậy thấp")
        add("data_issue_amount_missing","Thiếu cột số tiền (mapping sai)")
        add("data_issue_debit_credit_both","Nợ/Có cùng có giá trị (dữ liệu bẩn)")
        if not chunks:
            return pd.DataFrame(columns=["Vấn đề","Ngày","Giờ","Chiều","Số tiền","Đối ứng","Tin cậy","Nội dung","File"])
        out=pd.concat(chunks,ignore_index=True)
        cols=["Vấn đề","txn_date","txn_time","direction","amount","cp_account","cp_confidence","description_clean","__source_file__"]
        for c in cols:
            if c not in out.columns: out[c]=""
        rename={"txn_date":"Ngày","txn_time":"Giờ","direction":"Chiều","amount":"Số tiền","cp_account":"Đối ứng","cp_confidence":"Tin cậy",
                "description_clean":"Nội dung","__source_file__":"File"}
        return out[cols].rename(columns=rename).reset_index(drop=True)

RAW_CLEAN_COLS=["txn_date","txn_time","direction","debit_amount","credit_amount","amount","cp_account","cp_name","cp_bank","cp_confidence","description_raw","__source_file__"]
RAW_FORENSIC_COLS=["txn_date","txn_time","direction","amount","signed_amount","direction_evidence","cp_account","cp_bank","cp_confidence","cp_candidates",
                   "primary_tag","tags","tag_confidence","tag_reason","phone_cnt","wallet_cnt","plate_cnt","utility_cnt","risk_level","risk_score","risk_flags_text",
                   "description_clean","__source_file__"]

def build_raw_view(df: pd.DataFrame, forensic: bool=False)->pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    cols=RAW_FORENSIC_COLS if forensic else RAW_CLEAN_COLS
    cols=[c for c in cols if c in df.columns]
    out=df[cols].copy()
    if "tags" in out.columns:
        out["tags"]=out["tags"].apply(lambda xs: ", ".join(xs) if isinstance(xs,list) else str(xs))
    if "cp_candidates" in out.columns:
        out["cp_candidates"]=out["cp_candidates"].apply(lambda xs: json.dumps(xs,ensure_ascii=False) if isinstance(xs,list) else str(xs))
    return out

class DeepPipeline:
    def __init__(self, mapping: MappingConfig, owner: OwnerInfo, cp_cfg: CounterpartyConfig, tag_rules: List[TagRule], suspicious_cfg: SuspiciousConfig):
        self.mapping=mapping
        self.owner=owner or OwnerInfo()
        self.cp_cfg=cp_cfg or CounterpartyConfig()
        self.tag_rules=tag_rules or []
        self.suspicious_cfg=suspicious_cfg or SuspiciousConfig()
    def run(self, paths: List[str], mode: str="BOTH", forensic_raw: bool=False, progress_cb: ProgressCB=None)->Dict[str,pd.DataFrame]:
        def p(pct:int,msg:str):
            if progress_cb:
                try: progress_cb(int(pct),str(msg))
                except Exception: pass
        p(5,"Đọc file")
        raw=StatementLoader(self.mapping).load(paths)
        raw_keep, raw_drop = filter_garbage_rows(raw, mapping=self.mapping, level=getattr(self.mapping,"garbage_filter_level","CHUAN"))
        raw = raw_keep
        p(22,"Chuẩn hoá (Ngày/Giờ/Số tiền/Chiều)")
        norm=StatementNormalizer(self.mapping,self.owner).normalize(raw)
        p(42,"Bóc tách đối ứng")
        norm=CounterpartyExtractor(self.owner,self.cp_cfg).extract(norm)
        p(58,"Bóc tách thực thể")
        norm=EntityExtractor().extract(norm)
        p(72,"Gắn tag")
        if self.tag_rules:
            norm=TagEngine(self.tag_rules).apply(norm)
        else:
            norm["tags"]=[[] for _ in range(len(norm))]; norm["primary_tag"]=""; norm["tag_confidence"]=0.0; norm["tag_reason"]=""; norm["dq_tag_missing"]=True
        p(84,"Tính nghi vấn")
        norm=SuspiciousDetector(self.suspicious_cfg).detect(norm)
        mode_u=(mode or "BOTH").upper()
        if mode_u=="DEBIT": view=norm[norm["direction"]=="DEBIT"].copy()
        elif mode_u=="CREDIT": view=norm[norm["direction"]=="CREDIT"].copy()
        else: view=norm.copy()
        p(92,"Tổng hợp báo cáo")
        reports={
            "RAW_FULL": view,
            "RAW_CLEAN": build_raw_view(view, forensic=False),
            "RAW_FORENSIC": build_raw_view(view, forensic=True),
            "COUNTERPARTY": ReportBuilder.by_counterparty(view, min_conf=float(self.cp_cfg.min_confidence_for_stats)),
            "TAG": ReportBuilder.by_tag(view),
            "TAG_ENTITY": ReportBuilder.tag_entity(view),
            "SUSPICIOUS": ReportBuilder.suspicious(view),
            "DATA_QUALITY": ReportBuilder.data_quality(view),
            "GARBAGE_DROPPED": raw_drop if isinstance(raw_drop, pd.DataFrame) else pd.DataFrame(),
        }
        p(100,"Hoàn tất")
        return reports
