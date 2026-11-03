# -*- coding: utf-8 -*-
import os
import json
import re
import shutil
import tempfile
from contextlib import contextmanager
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from settlement_excel import export_monthly, parse_pay_date, roc_date, display_date, monthly_path, parse_roc_input
from settlement_documents import (export_directory, available_path, create_detail_document,
                                  create_expense_document, create_market_fee_document, market_fee_template, create_collection_document)
import win32com.client as win32
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))
LEGACY_CACHE_FILE = os.path.join(BASE_FOLDER, "pending_batches.json")
CACHE_FILE = os.path.join(BASE_FOLDER, "drafts", "pending_batches.json")
TEMPLATE_FOLDER = os.path.join(BASE_FOLDER, "templates")
EXPORT_FOLDER = os.path.join(os.path.dirname(BASE_FOLDER), "exports")

# 2026 名冊與受款帳戶字典
ACCOUNT_BOOK = {
    # 司機群 (本會/馬光/跨行)
    "張盡固": {"name": "張盡固", "acc": "4270-6", "note": ""},
    "杜虹儀": {"name": "杜虹儀", "acc": "15470-0", "note": ""},
    "洪永原": {"name": "洪永原", "acc": "6096-6", "note": ""},
    "詹啓明": {"name": "土庫鎮農會", "acc": "2941-5", "note": "詹啓明"},
    "林俊賢": {"name": "土庫鎮農會", "acc": "2941-5", "note": "林俊賢"},

    # 豬農本會/馬光專屬帳號
    "張嘉顯": {"name": "張嘉顯", "acc": "10333-9", "note": ""},
    "曾淑美": {"name": "曾淑美", "acc": "1715-0", "note": "張嘉顯汰豬專用"},
    "張政曲": {"name": "張政曲", "acc": "00098-6", "note": ""},
    "陳沛琪": {"name": "陳沛琪", "acc": "50557-5", "note": "張政曲-陳"},
    "張淑玲": {"name": "張淑玲", "acc": "12341-0", "note": ""},
    "謝堎裕": {"name": "謝堎裕", "acc": "15038-2", "note": ""},
    "楊昆量": {"name": "楊昆量", "acc": "18228-6", "note": ""},
    "林冠佑": {"name": "林冠佑", "acc": "15678-6", "note": ""},

    # 非本會用戶 (固定戶名為土庫鎮農會，帳號 2941-5)
    "吳英吉": {"name": "土庫鎮農會", "acc": "2941-5", "note": "吳英吉"},
    "徐麗屏": {"name": "土庫鎮農會", "acc": "2941-5", "note": "吳英吉"},
    "陳尚成": {"name": "土庫鎮農會", "acc": "2941-5", "note": "陳尚成"},
    "張維瑲": {"name": "土庫鎮農會", "acc": "2941-5", "note": "陳尚成-張"},
    "林沅科": {"name": "土庫鎮農會", "acc": "2941-5", "note": "陳尚成-林"},
    "傅淑玲": {"name": "土庫鎮農會", "acc": "2941-5", "note": "陳尚成-傅"},
    "吳櫂閎": {"name": "土庫鎮農會", "acc": "2941-5", "note": "吳櫂閎"},
    "王章憲": {"name": "土庫鎮農會", "acc": "2941-5", "note": "王章憲"},
    "林睿毅": {"name": "土庫鎮農會", "acc": "2941-5", "note": "林睿毅"},
    "吳昌樺": {"name": "土庫鎮農會", "acc": "2941-5", "note": "吳昌樺"},
    "林嘉祈": {"name": "土庫鎮農會", "acc": "2941-5", "note": "林嘉祈"},
    "李籲安": {"name": "土庫鎮農會", "acc": "2941-5", "note": "李籲安"}
}

# 豬農名單 (已剔除吳耀閔，加入張盡固)
FARMER_LIST = [
    "吳英吉", "張嘉顯", "張政曲", "陳尚成", "張盡固", "吳櫂閎",
    "張淑玲", "謝堎裕", "楊昆量", "林冠佑", "王章憲", "林睿毅",
    "吳昌樺", "林嘉祈", "李籲安"
]

# 特殊受款人選單對照
PAYEE_PRESETS = {
    "吳英吉": ["土庫鎮農會 (備註:吳英吉)"],
    "張嘉顯": ["張嘉顯 (正常肉豬)", "曾淑美 (淘汰豬專用)"],
    "張政曲": ["張政曲", "陳沛琪 (備註:張政曲-陳)"],
    "陳尚成": [
        "土庫鎮農會 (備註:陳尚成-張)",
        "土庫鎮農會 (備註:陳尚成-林)",
        "土庫鎮農會 (備註:陳尚成-傅)",
        "土庫鎮農會 (備註:陳尚成)"
    ],
    "張盡固": ["張盡固"]
}

MARKET_OPTIONS = ["(請選擇拍賣市場)", "新北市市場", "桃園縣市場", "南投縣市場", "彰化縣市場", "雲林縣市場"]

def num_to_chinese_digit_upper(n_val):
    """將數字轉換為國字大寫（例如 1 -> 壹, 2 -> 貳, 3 -> 參）"""
    digits_map = {
        "0": "零", "1": "壹", "2": "貳", "3": "參", "4": "肆",
        "5": "伍", "6": "陸", "7": "柒", "8": "捌", "9": "玖", "10": "拾"
    }
    s = str(n_val).strip()
    if s in digits_map:
        return digits_map[s]
    try:
        val = int(s)
        d = ["零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"]
        if val <= 10:
            return digits_map.get(str(val), str(val))
        elif val < 20:
            return "拾" + (d[val % 10] if val % 10 != 0 else "")
        elif val < 100:
            tens = d[val // 10] + "拾"
            units = d[val % 10] if val % 10 != 0 else ""
            return tens + units
        return str(val)
    except:
        return s

def set_cell_format(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, font_name="標楷體", font_size=11):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(str(text))
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    run._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), font_name)

def get_account_info(payee, farmer=None):
    clean = payee.strip()
    
    # 司機優先比對
    if clean in ["張盡固", "杜虹儀", "洪永原", "詹啓明", "林俊賢"]:
        return ACCOUNT_BOOK[clean]
    
    if farmer == "張政曲":
        if "陳沛琪" in clean:
            return {"name": "陳沛琪", "acc": "50557-5", "note": "張政曲-陳"}
        return {"name": "張政曲", "acc": "00098-6", "note": ""}
            
    elif farmer == "陳尚成":
        if "張" in clean or "張維瑲" in clean:
            return {"name": "土庫鎮農會", "acc": "2941-5", "note": "陳尚成-張"}
        if "林" in clean or "林沅科" in clean:
            return {"name": "土庫鎮農會", "acc": "2941-5", "note": "陳尚成-林"}
        if "傅" in clean or "傅淑玲" in clean:
            return {"name": "土庫鎮農會", "acc": "2941-5", "note": "陳尚成-傅"}
        return {"name": "土庫鎮農會", "acc": "2941-5", "note": "陳尚成"}

    elif farmer == "吳英吉":
        return {"name": "土庫鎮農會", "acc": "2941-5", "note": "吳英吉"}

    elif farmer == "張嘉顯":
        if "曾淑美" in clean or "淘汰" in clean:
            return {"name": "曾淑美", "acc": "1715-0", "note": "張嘉顯汰豬專用"}
        return {"name": "張嘉顯", "acc": "10333-9", "note": ""}

    elif farmer == "張盡固":
        return {"name": "張盡固", "acc": "4270-6", "note": ""}

    pure_name = clean.split("(")[0].strip()
    if pure_name in ["張淑玲", "謝堎裕", "楊昆量", "林冠佑"]:
        return ACCOUNT_BOOK[pure_name]

    note_name = farmer if farmer else pure_name
    return {"name": "土庫鎮農會", "acc": "2941-5", "note": note_name}

def num_to_chinese_upper(n):
    units = ["", "拾", "佰", "仟"]
    big_units = ["", "萬", "億"]
    digits = ["零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"]
    if n == 0:
        return "零元整"
    s = str(int(n))
    result = ""
    length = len(s)
    groups = []
    while length > 0:
        groups.append(s[max(0, length - 4):length])
        length -= 4
    for g_idx, grp in enumerate(groups):
        grp_res = ""
        grp_len = len(grp)
        zero_flag = False
        for i, ch in enumerate(grp):
            d = int(ch)
            u_idx = grp_len - i - 1
            if d != 0:
                if zero_flag:
                    grp_res += "零"
                    zero_flag = False
                grp_res += digits[d] + units[u_idx]
            else:
                zero_flag = True
        if grp_res:
            grp_res += big_units[g_idx]
        result = grp_res + result
    return result + "元整"


class PigSettlementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("土庫鎮農會 - 毛豬運銷結算與套印系統")
        self.root.geometry("1040x780")
        self.current_farmers = []
        self.active_draft_id = None
        self.editing_index = None
        self._freight_context = None
        self._selected_farmer = None
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.setup_ui()
        self.load_draft_list()

    def setup_ui(self):
        top_frame = ttk.LabelFrame(self.root, text="批次基本資料", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(top_frame, text="結帳／出帳日:").grid(row=0, column=0, sticky="w")
        self.ent_pay_date = ttk.Entry(top_frame, width=15)
        self.ent_pay_date.insert(0, roc_date(datetime.today()))
        self.ent_pay_date.grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(top_frame, text="民國 yyy-mm-dd，例如 115-11-03；所有統計依出帳日歸檔").grid(row=0, column=2, padx=10)

        farmer_frame = ttk.LabelFrame(self.root, text="新增/修改 豬單明細 (每張豬單可填不同拍賣日期與市場)", padding=10)
        farmer_frame.pack(fill="x", padx=10, pady=5)
        
        self.entries = {}
        ttk.Label(farmer_frame, text="拍賣日期:").grid(row=4, column=0, sticky="w")
        self.ent_auction_date = ttk.Entry(farmer_frame, width=15)
        self.ent_auction_date.insert(0, roc_date(datetime.today()))
        self.ent_auction_date.grid(row=4, column=1, padx=5, pady=2)
        ttk.Label(farmer_frame, text="民國 yyy-mm-dd").grid(row=4, column=2, columnspan=2, sticky="w")

        
        ttk.Label(farmer_frame, text="拍賣市場:").grid(row=0, column=0, sticky="w", pady=2)
        self.cbo_market = ttk.Combobox(farmer_frame, values=MARKET_OPTIONS, width=13, state="readonly")
        self.cbo_market.set(MARKET_OPTIONS[0])
        self.cbo_market.grid(row=0, column=1, padx=5, pady=2)
        self.entries["market"] = self.cbo_market
        self.cbo_market.bind("<<ComboboxSelected>>", self.on_market_change)

        ttk.Label(farmer_frame, text="豬農會員:").grid(row=0, column=2, sticky="w", pady=2)
        self.cbo_farmer = ttk.Combobox(farmer_frame, values=FARMER_LIST, width=12)
        self.cbo_farmer.grid(row=0, column=3, padx=5, pady=2)
        self.entries["farmer"] = self.cbo_farmer
        self.cbo_farmer.bind("<KeyRelease>", self.on_farmer_key_release)
        self.cbo_farmer.bind("<<ComboboxSelected>>", self.on_farmer_selected)
        self.cbo_farmer.bind("<FocusOut>", self.on_farmer_selected)

        ttk.Label(farmer_frame, text="受款帳戶/人:").grid(row=0, column=4, sticky="w", pady=2)
        self.cbo_payee = ttk.Combobox(farmer_frame, values=[], width=24)
        self.cbo_payee.grid(row=0, column=5, padx=5, pady=2)
        self.entries["payee"] = self.cbo_payee

        self.add_field(farmer_frame, "運輸頭數:", "heads", 1, 0)
        self.add_field(farmer_frame, "總價款:", "total", 1, 2)
        self.add_field(farmer_frame, "市場及省縣手續費:", "market_fee", 1, 4)
        
        self.add_field(farmer_frame, "事故款項 (可後補):", "accident", 2, 0)
        self.add_field(farmer_frame, "死亡保險費:", "insurance", 2, 2)
        self.add_field(farmer_frame, "農會手續費(0.1%):", "fa_fee", 2, 4)
        self.entries["fa_fee"].configure(state="readonly")
        
        self.add_field(farmer_frame, "運費:", "freight", 3, 0)
        
        ttk.Label(farmer_frame, text="載運司機:").grid(row=3, column=2, sticky="w", pady=2)
        d_opts = ["(無 / 自運)", "張盡固", "杜虹儀", "洪永原", "詹啓明", "林俊賢"]
        self.cbo_driver = ttk.Combobox(farmer_frame, values=d_opts, width=12, state="readonly")
        self.cbo_driver.set(d_opts[0])
        self.cbo_driver.grid(row=3, column=3, padx=5, pady=2)

        self.add_field(farmer_frame, "實付金額(計算):", "net", 3, 4)
        self.entries["net"].configure(state="readonly")

        self.entries["heads"].bind("<KeyRelease>", self.auto_match_freight)
        for k in ["total", "market_fee", "accident", "insurance", "freight"]:
            self.entries[k].bind("<KeyRelease>", self.calculate_farmer_amounts)

        self.btn_add = ttk.Button(farmer_frame, text="加入/更新此筆豬單", command=self.ask_and_add_farmer)
        self.btn_add.grid(row=5, column=4, columnspan=2, pady=6, sticky="e")
        ttk.Button(farmer_frame, text="取消編輯／清除輸入",
                   command=self.cancel_input).grid(row=5, column=0, columnspan=2)

        self.setup_enter_traversal()

        list_frame = ttk.LabelFrame(self.root, text="本批次明細清單 (雙擊修改/補款，右鍵可單筆或全部刪除)", padding=10)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        cols = ("date", "farmer", "payee", "market", "heads", "total", "market_fee", "slip_net", "accident", "fa_fee", "insurance", "freight", "driver", "net")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=7)
        
        col_defs = [
            ("date", "拍賣日期", 95),
            ("farmer", "豬農會員", 75), ("payee", "受款帳戶", 95), ("market", "拍賣市場", 80),
            ("heads", "頭數", 45), ("total", "總價款", 80),
            ("market_fee", "市場費用", 70), ("slip_net", "豬單實付(核對)", 85),
            ("accident", "事故款", 65), ("fa_fee", "農會手續費", 70),
            ("insurance", "保險費", 65), ("freight", "運費", 65), ("driver", "運費司機", 70), ("net", "農會實付", 85)
        ]
        
        for col_id, col_name, width in col_defs:
            self.tree.heading(col_id, text=col_name)
            self.tree.column(col_id, width=width, anchor="center")
            
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_farmer_select)

        self.menu_popup = tk.Menu(self.root, tearoff=0)
        self.menu_popup.add_command(label="🗑️ 刪除此筆豬單", command=self.delete_selected_farmer)
        self.menu_popup.add_separator()
        self.menu_popup.add_command(label="⚠️ 清空全部明細 (本批重填)", command=self.clear_all_farmers)
        
        self.tree.bind("<Button-3>", self.show_context_menu)

        action_frame = ttk.Frame(self.root, padding=8)
        action_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(action_frame, text="未結案批次:").pack(side="left", padx=2)
        self.cbo_drafts = ttk.Combobox(action_frame, state="readonly", width=28)
        self.cbo_drafts.pack(side="left", padx=5)
        
        btn_load = ttk.Button(action_frame, text="📂 讀取草稿", command=self.load_selected_draft)
        btn_load.pack(side="left", padx=2)

        btn_delete_draft = ttk.Button(action_frame, text="🗑️ 刪除草稿", command=self.delete_selected_draft)
        btn_delete_draft.pack(side="left", padx=2)

        btn_save_draft = ttk.Button(action_frame, text="💾 暫存為草稿", command=self.save_as_draft)
        btn_save_draft.pack(side="left", padx=8)

        btn_finalize = ttk.Button(action_frame, text="✅ 確認結案 (寫入Excel並預覽Word)", command=self.finalize_batch)
        btn_finalize.pack(side="right", padx=5)
        preview_frame = ttk.Frame(self.root, padding=5)
        preview_frame.pack(fill="x", padx=10)
        ttk.Button(preview_frame, text="預覽價款明細表（不結案）",
                   command=self.preview_detail_only).pack(side="left")
        ttk.Button(preview_frame, text="預覽開支請示／手續費",
                   command=self.preview_expenses_only).pack(side="left", padx=8)
        ttk.Button(preview_frame, text="預覽各項費用明細",
                   command=self.preview_collection_only).pack(side="left", padx=8)
        ttk.Button(preview_frame, text="開啟輸出資料夾",
                   command=self.open_export_folder).pack(side="left", padx=8)
        ttk.Button(preview_frame, text="匯入測試草稿",
                   command=self.import_sample_drafts).pack(side="left", padx=8)

    def add_field(self, parent, text, key, r, c):
        ttk.Label(parent, text=text).grid(row=r, column=c, sticky="w", pady=2)
        e = ttk.Entry(parent, width=14)
        e.grid(row=r, column=c+1, padx=5, pady=2)
        self.entries[key] = e

    def on_farmer_key_release(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape"):
            return
        typed = self.cbo_farmer.get().strip()
        if not typed:
            self.cbo_farmer["values"] = FARMER_LIST
        else:
            matches = [name for name in FARMER_LIST if typed in name]
            self.cbo_farmer["values"] = matches
            if matches:
                self.cbo_farmer.event_generate("<Down>")

    def on_farmer_selected(self, event=None):
        farmer = self.cbo_farmer.get().strip()
        if not farmer:
            return
        if farmer == self._selected_farmer:
            return
        self._selected_farmer = farmer
        
        if farmer in PAYEE_PRESETS:
            opts = PAYEE_PRESETS[farmer]
            self.cbo_payee["values"] = opts
            self.cbo_payee.set(opts[0])
        elif farmer in ["張淑玲", "謝堎裕", "楊昆量", "林冠佑", "張盡固"]:
            self.cbo_payee["values"] = [farmer]
            self.cbo_payee.set(farmer)
        else:
            val = f"土庫鎮農會 (備註:{farmer})"
            self.cbo_payee["values"] = [val]
            self.cbo_payee.set(val)

        self.auto_match_freight()

    def on_market_change(self, event=None):
        self.auto_match_freight()

    def auto_match_freight(self, event=None):
        farmer = self.cbo_farmer.get().strip()
        market = self.cbo_market.get().strip()
        context = (farmer, market)
        if context != self._freight_context:
            self.cbo_driver.set("(無 / 自運)")
            self.entries["freight"].delete(0, tk.END)
            self._freight_context = context
        try:
            heads = self._get_int("heads")
        except ValueError:
            self.calculate_farmer_amounts()
            return

        if farmer == "吳英吉":
            self.cbo_driver.set("杜虹儀")
            self.entries["freight"].delete(0, tk.END)
            self.entries["freight"].insert(0, str(heads * 55))
        elif farmer == "張嘉顯" and market == "新北市市場":
            self.cbo_driver.set("洪永原")
            self.entries["freight"].delete(0, tk.END)
            self.entries["freight"].insert(0, str(heads * 120))
        elif farmer == "吳櫂閎":
            self.cbo_driver.set("詹啓明")
            if not self.entries["freight"].get().strip():
                self.entries["freight"].insert(0, "9000")
        
        self.calculate_farmer_amounts()

    def setup_enter_traversal(self):
        seq = [
            self.ent_pay_date, self.ent_auction_date,
            self.cbo_market, self.cbo_farmer, self.cbo_payee,
            self.entries["heads"], self.entries["total"],
            self.entries["market_fee"], self.entries["accident"],
            self.entries["insurance"]
        ]
        for i in range(len(seq) - 1):
            curr_w = seq[i]
            next_w = seq[i + 1]
            curr_w.bind("<Return>", lambda e, nxt=next_w: (nxt.focus_set(), "break")[1])
        
        self.entries["insurance"].bind("<Return>", lambda e: (self.entries["freight"].focus_set(), "break")[1])
        self.entries["freight"].bind("<Return>", self.on_freight_enter)
        self.cbo_driver.bind("<Return>", lambda e: (self.ask_and_add_farmer(), "break")[1])

    def on_freight_enter(self, event=None):
        if not self.calculate_farmer_amounts():
            return "break"
        freight_val = self._get_int("freight")
        driver_val = self.cbo_driver.get().strip()

        if freight_val > 0 and driver_val not in ["(無 / 自運)", ""]:
            self.cbo_driver.focus_set()
        else:
            self.ask_and_add_farmer()
        return "break"

    def check_market_selected(self):
        m = self.cbo_market.get().strip()
        if not m or m == "(請選擇拍賣市場)":
            messagebox.showwarning("提醒", "請先為此筆豬單選擇「拍賣市場」！")
            self.cbo_market.focus_set()
            return False
        return True

    def ask_and_add_farmer(self):
        if not self.check_market_selected():
            return

        farmer = self.cbo_farmer.get().strip()
        payee = self.cbo_payee.get().strip() or farmer
        market = self.cbo_market.get().strip()
        if not farmer:
            messagebox.showwarning("提示", "請輸入豬農姓名！")
            self.cbo_farmer.focus_set()
            return

        try:
            parse_roc_input(self.ent_auction_date.get())
            for key in ["heads", "total", "market_fee", "accident", "insurance", "freight"]:
                self._get_int(key)
        except ValueError as exc:
            messagebox.showwarning("輸入有誤", f"請檢查日期（民國年，如 115-11-03）及數字欄位。\n{exc}")
            return
        self.calculate_farmer_amounts()
        heads = self._get_int("heads")
        net = int(self.entries["net"].get() or 0)
        freight = self._get_int("freight")
        driver = self.cbo_driver.get().strip()
        if freight > 0 and driver in ("", "(無 / 自運)"):
            messagebox.showwarning("請確認運費", "已填運費，請選擇實際受款司機；若沒有運費，請填 0。")
            return

        slip_net = self._get_int("total") - self._get_int("market_fee")
        lines = [
            f"拍賣市場：{market}",
            f"豬農會員(Excel)：{farmer}",
            f"受款帳戶(Word)：{payee}",
            f"頭數：{heads} 頭",
            f"【豬單實付】(核對用)：NT$ {slip_net:,}",
            f"【農會實付】：NT$ {net:,}"
        ]
        if freight > 0:
            lines.append(f"運費：NT$ {freight:,}（司機：{driver}）")

        info = get_account_info(payee, farmer)
        lines.append(f"實際受款：{info['name']}／{info['acc']}／{info['note']}")
        msg_body = ("是否儲存此筆修改？" if self.editing_index is not None else "是否將此筆資料加入本批清單？") + "\n\n" + "\n".join(lines)
        confirm = messagebox.askyesno("確認加入", msg_body)
        if confirm:
            self.add_farmer_to_list()
        else:
            self.cbo_farmer.focus_set()

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.menu_popup.post(event.x_root, event.y_root)

    def delete_selected_farmer(self):
        if not self.require_saved_input():
            return
        selected = self.tree.selection()
        if not selected:
            return
        idx = self.tree.index(selected[0])
        f_name = self.current_farmers[idx]["farmer"]
        confirm = messagebox.askyesno("確認刪除", f"確定要刪除豬農【{f_name}】這筆資料嗎？")
        if confirm:
            self.current_farmers.pop(idx)
            self.refresh_treeview()

    def clear_all_farmers(self):
        if not self.require_saved_input():
            return
        if not self.current_farmers:
            return
        confirm = messagebox.askyesno("確認清空", "確定要清空本批次【所有豬單明細】嗎？")
        if confirm:
            self.current_farmers = []
            self.refresh_treeview()
            self.clear_farmer_inputs()

    def on_closing(self):
        if not self.require_saved_input():
            return
        if self.current_farmers:
            res = messagebox.askyesnocancel(
                "確認離開", 
                "目前清單中尚有未結案資料，是否先【暫存為草稿】再離開？\n\n【是】：儲存草稿並關閉\n【否】：直接離開不儲存\n【取消】：返回視窗"
            )
            if res is True:
                if self.save_as_draft(show_msg=False):
                    self.root.destroy()
            elif res is False:
                self.root.destroy()
            else:
                return
        else:
            self.root.destroy()

    def _get_int(self, key):
        val = self.entries[key].get().strip()
        if not val:
            return 0
        parts = [part.strip() for part in val.split("+")]
        if not all(re.fullmatch(r"(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)", part) for part in parts):
            labels = {"heads": "頭數", "total": "總價款", "market_fee": "市場費用",
                      "accident": "事故款", "insurance": "保險費", "freight": "運費"}
            raise ValueError(f"{labels.get(key, key)}：請輸入整數，可使用千分位或加號。")
        return sum(int(part.replace(",", "")) for part in parts)

    def calculate_farmer_amounts(self, event=None):
        try:
            total = self._get_int("total")
            market_fee = self._get_int("market_fee")
            accident = self._get_int("accident")
            insurance = self._get_int("insurance")
            freight = self._get_int("freight")
        except ValueError:
            for key in ("fa_fee", "net"):
                self.entries[key].configure(state="normal")
                self.entries[key].delete(0, tk.END)
                self.entries[key].configure(state="readonly")
            return False

        fa_fee = round(total * 0.001)
        net_pay = total - (market_fee + accident + fa_fee + insurance + freight)

        self.entries["fa_fee"].configure(state="normal")
        self.entries["fa_fee"].delete(0, tk.END)
        self.entries["fa_fee"].insert(0, str(fa_fee))
        self.entries["fa_fee"].configure(state="readonly")

        self.entries["net"].configure(state="normal")
        self.entries["net"].delete(0, tk.END)
        self.entries["net"].insert(0, str(net_pay))
        self.entries["net"].configure(state="readonly")
        return True

    def add_farmer_to_list(self):
        farmer = self.cbo_farmer.get().strip()
        payee = self.cbo_payee.get().strip() or farmer
        market = self.cbo_market.get().strip()
        record = {
            "farmer": farmer,
            "payee": payee,
            "market": market,
            "date": roc_date(parse_roc_input(self.ent_auction_date.get())),
            "heads": self._get_int("heads"),
            "total": self._get_int("total"),
            "market_fee": self._get_int("market_fee"),
            "accident": self._get_int("accident"),
            "fa_fee": int(self.entries["fa_fee"].get() or 0),
            "insurance": self._get_int("insurance"),
            "freight": self._get_int("freight"),
            "driver": self.cbo_driver.get(),
            "net": int(self.entries["net"].get() or 0)
        }
        
        if self.editing_index is None:
            self.current_farmers.append(record)
        else:
            self.current_farmers[self.editing_index] = record
        self.refresh_treeview()
        self.clear_farmer_inputs(keep_market=True)
        self.cbo_farmer.focus_set()

    def refresh_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for f in self.current_farmers:
            slip_net = f["total"] - f["market_fee"]
            self.tree.insert("", "end", values=(
                display_date(f.get("date", "")), f["farmer"], f.get("payee", f["farmer"]), f.get("market", ""),
                f["heads"], f"{f['total']:,}", f"{f['market_fee']:,}",
                f"{slip_net:,}",
                f"{f['accident']:,}", f"{f['fa_fee']:,}", f"{f['insurance']:,}",
                f"{f['freight']:,}", f["driver"], f"{f['net']:,}"
            ))

    def on_farmer_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        if not self.require_saved_input():
            return
        idx = self.tree.index(selected[0])
        f = self.current_farmers[idx]
        self.editing_index = idx
        self.btn_add.configure(text="儲存此筆修改")
        original_date = self.ent_auction_date.get().strip()
        self.ent_auction_date.delete(0, tk.END)
        self.ent_auction_date.insert(0, display_date(f.get("date", original_date)))
        
        self.cbo_market.set(f.get("market", MARKET_OPTIONS[0]))
        self.cbo_farmer.set(f["farmer"])
        self.on_farmer_selected()
        self.cbo_payee.set(f.get("payee", f["farmer"]))
        for k in ["heads", "total", "market_fee", "accident", "insurance", "freight"]:
            self.entries[k].delete(0, tk.END)
            self.entries[k].insert(0, str(f[k]))
            
        self.cbo_driver.set(f["driver"])
        self.calculate_farmer_amounts()

    def has_pending_input(self):
        return (self.editing_index is not None or bool(self.cbo_farmer.get().strip())
                or bool(self.cbo_payee.get().strip())
                or self.cbo_driver.get() not in ("", "(無 / 自運)")
                or any(self.entries[k].get().strip() for k in
                       ("heads", "total", "market_fee", "accident", "insurance", "freight")))

    def require_saved_input(self):
        if self.has_pending_input():
            messagebox.showwarning("尚有未儲存輸入", "請先按「加入／儲存此筆修改」，或按「取消編輯／清除輸入」，再進行此操作。")
            return False
        return True

    def cancel_input(self):
        if not self.has_pending_input() or messagebox.askyesno("取消輸入", "確定放棄尚未儲存的輸入？原清單紀錄會保留。"):
            self.clear_farmer_inputs(keep_market=True)

    def clear_farmer_inputs(self, keep_market=False):
        self.editing_index = None
        self._selected_farmer = None
        self._freight_context = None
        self.btn_add.configure(text="加入／更新此筆豬單")
        if not keep_market:
            self.cbo_market.set(MARKET_OPTIONS[0])
        self.cbo_farmer.set("")
        self.cbo_farmer["values"] = FARMER_LIST
        self.cbo_payee.set("")
        self.cbo_payee["values"] = []
        for k, ent in self.entries.items():
            if k not in ["farmer", "payee", "market"]:
                ent.configure(state="normal")
                ent.delete(0, tk.END)
        self.entries["fa_fee"].configure(state="readonly")
        self.entries["net"].configure(state="readonly")
        self.cbo_driver.set("(無 / 自運)")

    def save_as_draft(self, show_msg=True):
        if not self.require_saved_input():
            return
        try:
            pay_date = roc_date(parse_roc_input(self.ent_pay_date.get()))
        except ValueError as exc:
            messagebox.showwarning("出帳日期有誤", str(exc))
            return False
        
        if not self.current_farmers:
            if show_msg:
                messagebox.showwarning("提示", "目前清單無資料！")
            return

        draft_id = self.active_draft_id or f"{pay_date}_{int(datetime.now().timestamp())}"
        draft_data = {
            "id": draft_id,
            "pay_date": pay_date, "farmers": self.current_farmers
        }

        drafts = self._read_cache()
        if drafts.get(draft_id, {}).get('name'):
            draft_data['name'] = drafts[draft_id]['name']
        drafts[draft_id] = draft_data
        self._write_cache(drafts)
        self.active_draft_id = draft_id
        if show_msg:
            messagebox.showinfo("成功", f"批次 [出帳:{pay_date}] 已成功暫存！")
        self.load_draft_list()
        return True

    def load_draft_list(self):
        drafts = self._read_cache()
        self.draft_lookup = {}
        options = []
        for d_id, data in drafts.items():
            total_net = sum(f["net"] for f in data["farmers"])
            display = f"出帳:{display_date(data.get('pay_date','未定'))} - {len(data['farmers'])}筆, ${total_net:,}"
            if data.get('name'):
                display = f"{data['name']}｜{display}"
            options.append(display)
            self.draft_lookup[display] = d_id
        self.cbo_drafts["values"] = options
        if options:
            self.cbo_drafts.current(0)
        else:
            self.cbo_drafts.set("")

    def import_sample_drafts(self):
        if not self.require_saved_input():
            return
        try:
            sample_path = os.path.join(os.path.dirname(BASE_FOLDER), 'samples', '測試草稿.json')
            with open(sample_path, encoding='utf-8') as stream:
                samples = json.load(stream)
            drafts = self._read_cache()
            count = 0
            for key, sample in samples.items():
                if key not in drafts:
                    drafts[key] = sample
                    count += 1
            if count:
                self._write_cache(drafts)
            self.load_draft_list()
            messagebox.showinfo('測試草稿', f'已匯入 {count} 組測試草稿，既有草稿未覆蓋。\n請從草稿清單選取【測試】資料讀取。\n測試資料若按結案，仍會寫入月份帳本。')
        except Exception as exc:
            messagebox.showerror('匯入未完成', str(exc))

    def load_selected_draft(self):
        if not self.require_saved_input():
            return
        if self.current_farmers and not messagebox.askyesno("讀取草稿", "讀取將取代目前清單，尚未暫存的變更會遺失。確定繼續？"):
            return
        sel = self.cbo_drafts.get()
        if not sel or sel not in self.draft_lookup:
            messagebox.showinfo("提示", "請先從下拉選單選取要讀取的草稿！")
            return
        d_id = self.draft_lookup[sel]
        draft = self._read_cache().get(d_id)
        if draft:
            self.active_draft_id = d_id
            self.ent_pay_date.delete(0, tk.END); self.ent_pay_date.insert(0, display_date(draft.get("pay_date", "")))
            self.current_farmers = draft["farmers"]
            for record in self.current_farmers:
                record["date"] = display_date(record.get("date", draft.get("auction_date", "")))
                record.setdefault("market", draft.get("market", ""))
            self.refresh_treeview()
            messagebox.showinfo("讀取成功", f"已載入 [出帳:{display_date(draft.get('pay_date', '未定'))}] 草稿！")

    def delete_selected_draft(self):
        if not self.require_saved_input():
            return
        sel = self.cbo_drafts.get()
        if not sel or sel not in self.draft_lookup:
            messagebox.showwarning("提示", "請先選取要刪除的草稿！")
            return
        d_id = self.draft_lookup[sel]
        confirm = messagebox.askyesno("確認刪除", f"確定要徹底刪除草稿：\n【{sel}】嗎？\n（此動作無法復原）")
        if confirm:
            drafts = self._read_cache()
            if d_id in drafts:
                del drafts[d_id]
                self._write_cache(drafts)
            if self.active_draft_id == d_id:
                self.active_draft_id = None
                self.current_farmers = []
                self.refresh_treeview()
                self.clear_farmer_inputs()
            self.load_draft_list()
            messagebox.showinfo("成功", "該草稿已成功刪除！")

    def _read_cache(self):
        if not os.path.exists(CACHE_FILE) and os.path.isfile(LEGACY_CACHE_FILE):
            os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
            os.rename(LEGACY_CACHE_FILE, CACHE_FILE)
        if not os.path.exists(CACHE_FILE):
            return {}
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}

    def _write_cache(self, data):
        os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def finalize_batch(self):
        if not self.require_saved_input():
            return
        if not self.current_farmers:
            messagebox.showwarning("提示", "無資料可結案！")
            return

        pay_date = self.ent_pay_date.get()
        if not pay_date:
            messagebox.showwarning("提示", "請務必填寫「結帳/出帳日」以建立專屬工作表！")
            return
        try:
            pay_date = roc_date(parse_roc_input(pay_date))
        except ValueError as exc:
            messagebox.showwarning("出帳日期有誤", str(exc))
            return

        farmer_net_total = sum(f["net"] for f in self.current_farmers)
        freight_total = sum(f["freight"] for f in self.current_farmers)
        actual_payout = farmer_net_total + freight_total
        fa_fee_total = sum(f["fa_fee"] for f in self.current_farmers)
        grand_total = actual_payout + fa_fee_total
        total_heads = sum(f["heads"] for f in self.current_farmers)

        confirm_msg = (
            f"【結案確認】出帳日：{pay_date}\n"
            f"----------------------------------------\n"
            f"豬農總人數：{len(self.current_farmers)} 人\n"
            f"運銷總頭數：{total_heads} 頭\n\n"
            f"豬農款實付：NT$ {farmer_net_total:,}\n"
            f"司機運費合計：NT$ {freight_total:,}\n"
            f"----------------------------------------\n"
            f"★【實付金額】(含運費)：NT$ {actual_payout:,}\n"
            f"農會手續費(0.1%)：NT$ {fa_fee_total:,}\n"
            f"★【總金額】(實付+手續費)：NT$ {grand_total:,}\n"
            f"----------------------------------------\n"
            f"確定正式寫入 Excel 並產出 Word 預覽嗎？"
        )
        confirm = messagebox.askyesno("確認結案", confirm_msg)
        if not confirm:
            return

        # 提示輸入憑證張數（輸入數字自動轉國字大寫）
        chinese_voucher = self.ask_voucher_count()
        if chinese_voucher is None:
            return

        # 先產生明細表，範本失敗時不寫入月份主帳本。
        try:
            output_dir = export_directory(EXPORT_FOLDER, parse_pay_date(pay_date))
            detail_path = self.create_detail_output(output_dir)
            expense_paths = self.create_expense_outputs(output_dir, pay_date, chinese_voucher)
            expense_paths.append(self.create_collection_output(output_dir, pay_date))
        except Exception as exc:
            messagebox.showerror("Word 文件未完成", f"尚未寫入 Excel，資料已保留。\n{exc}")
            return

        # 1. 寫入符合使用者格式的 Excel
        try:
            excel_path, sheet_name = self.export_to_excel_sheet(pay_date)
        except Exception as exc:
            messagebox.showerror("Excel 未完成", f"資料與草稿已保留。請確認月份檔案未被開啟，並檢查範本。\n{exc}")
            return

        # 2. 套印 Word；Excel 已完成時明確提示，避免誤以為整批未寫入。
        output_errors = []
        word_error = None
        try:
            self.preview_word_documents(pay_date, chinese_voucher, output_dir, detail_path, expense_paths)
        except Exception as exc:
            word_error = str(exc)
            output_errors.append(f"Word 預覽失敗：{word_error}")

        # 3. 結案自動清除草稿
        if self.active_draft_id:
            drafts = self._read_cache()
            if self.active_draft_id in drafts:
                del drafts[self.active_draft_id]
                self._write_cache(drafts)
                self.active_draft_id = None

        if output_errors:
            messagebox.showwarning("月份主帳本已完成，部分輸出失敗",
                f"已寫入 {excel_path.name}／{sheet_name}，請勿重複結案。\n" + "\n".join(output_errors)
                + f"\n已產生的檔案位於：{output_dir}")
        else:
            messagebox.showinfo("完成", f"已寫入 {excel_path.name}／{sheet_name} 並開啟 Word 預覽！\n輸出檔案：{output_dir}")
        self.current_farmers = []
        self.refresh_treeview()
        self.load_draft_list()

    def export_to_excel_sheet(self, pay_date):
        date = parse_pay_date(pay_date)
        legacy = monthly_path(BASE_FOLDER, date)
        target = monthly_path(EXPORT_FOLDER, date)
        if legacy.exists():
            if target.exists():
                raise ValueError('新舊位置都有同月份帳本，請先確認要使用哪份；未覆蓋任何帳本。')
            target.parent.mkdir(parents=True, exist_ok=True)
            os.rename(legacy, target)
        return export_monthly(
            EXPORT_FOLDER, os.path.join(TEMPLATE_FOLDER, "每日工作頁.xlsm"),
            pay_date, self.current_farmers, "")

    def create_detail_output(self, output_dir, preview=False):
        filename = ("預覽_" if preview else "") + "1.毛豬運銷價款明細表.docx"
        return create_detail_document(
            os.path.join(TEMPLATE_FOLDER, "1.毛豬運銷價款明細表.docx"),
            available_path(output_dir, filename), self.current_farmers, get_account_info)

    def open_export_folder(self):
        os.makedirs(EXPORT_FOLDER, exist_ok=True)
        os.startfile(EXPORT_FOLDER)

    @contextmanager
    def preview_directory(self):
        # Word keeps the source open; clean only this preview's temporary files
        # after its window closes, without touching exports or other documents.
        directory = tempfile.mkdtemp(prefix="pig_account_preview_")
        try:
            yield directory
        finally:
            self._clean_preview_directory(directory)

    def _clean_preview_directory(self, directory):
        try:
            for name in os.listdir(directory):
                try:
                    os.remove(os.path.join(directory, name))
                except OSError:
                    pass  # Still open in Word; retry while the app is running.
            os.rmdir(directory)
        except OSError:
            self.root.after(5000, lambda: self._clean_preview_directory(directory))

    def preview_detail_only(self):
        if not self.require_saved_input():
            return
        if not self.current_farmers:
            messagebox.showwarning("提示", "請先加入至少一筆豬單。")
            return
        try:
            with self.preview_directory() as preview_dir:
                path = self.create_detail_output(preview_dir, preview=True)
                word_app = win32.gencache.EnsureDispatch('Word.Application')
                word_app.Visible = True
                doc = word_app.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False)
                doc.Activate()
                doc.ActiveWindow.View.Type = 3  # Print layout.
        except Exception as exc:
            messagebox.showerror("預覽未完成", str(exc))

    def ask_voucher_count(self):
        value = simpledialog.askstring("檢附憑證", "請輸入檢附憑證張數（正整數，自動轉大寫）：", initialvalue="1")
        if value is None:
            return None
        value = value.strip()
        if not value.isascii() or not value.isdigit() or int(value) <= 0:
            messagebox.showwarning("憑證張數有誤", "請輸入大於 0 的整數。")
            return None
        return num_to_chinese_digit_upper(value)

    def create_expense_outputs(self, output_dir, pay_date, chinese_voucher, preview=False):
        prefix = "預覽_" if preview else ""
        name = "2.豬款開支請示.docx"
        paths = [create_expense_document(
            os.path.join(TEMPLATE_FOLDER, name), available_path(output_dir, prefix + name),
            pay_date, self.current_farmers, chinese_voucher, num_to_chinese_upper)]
        fee_name = market_fee_template(self.current_farmers)
        if fee_name:
            paths.append(create_market_fee_document(
                os.path.join(TEMPLATE_FOLDER, fee_name),
                available_path(output_dir, prefix + fee_name), pay_date))
        return paths

    def preview_expenses_only(self):
        if not self.require_saved_input():
            return
        if not self.current_farmers:
            messagebox.showwarning("提示", "請先加入至少一筆豬單。")
            return
        try:
            pay_date = roc_date(parse_roc_input(self.ent_pay_date.get()))
            voucher = self.ask_voucher_count()
            if voucher is None:
                return
            with self.preview_directory() as preview_dir:
                paths = self.create_expense_outputs(preview_dir, pay_date, voucher, preview=True)
                word = win32.gencache.EnsureDispatch('Word.Application')
                word.Visible = True
                docs = []
                for path in paths:
                    doc = word.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False)
                    doc.ActiveWindow.View.Type = 3
                    docs.append(doc)
                docs[0].Activate()
        except Exception as exc:
            messagebox.showerror("預覽未完成", str(exc))

    def create_collection_output(self, output_dir, pay_date, preview=False):
        name = "3.毛豬運銷各項費用收取明細表.docx"
        return create_collection_document(
            os.path.join(TEMPLATE_FOLDER, name),
            available_path(output_dir, ("預覽_" if preview else "") + name),
            pay_date, self.current_farmers)

    def preview_collection_only(self):
        if not self.require_saved_input():
            return
        if not self.current_farmers:
            messagebox.showwarning("提示", "請先加入至少一筆豬單。")
            return
        try:
            pay_date = roc_date(parse_roc_input(self.ent_pay_date.get()))
            with self.preview_directory() as preview_dir:
                path = self.create_collection_output(preview_dir, pay_date, preview=True)
                word = win32.gencache.EnsureDispatch('Word.Application')
                word.Visible = True
                doc = word.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False)
                doc.ActiveWindow.View.Type = 3
                doc.Activate()
        except Exception as exc:
            messagebox.showerror("預覽未完成", str(exc))

    def preview_word_documents(self, pay_date, chinese_voucher, output_dir=None, detail_path=None, expense_paths=None):
        output_dir = output_dir or export_directory(EXPORT_FOLDER, parse_pay_date(pay_date))
        detail_path = detail_path or self.create_detail_output(output_dir)
        if expense_paths is None:
            expense_paths = self.create_expense_outputs(output_dir, pay_date, chinese_voucher)
            expense_paths.append(self.create_collection_output(output_dir, pay_date))
        try:
            word_app = win32.gencache.EnsureDispatch('Word.Application')
        except:
            word_app = win32.Dispatch('Word.Application')
            
        word_app.Visible = True
        
        doc_obj1 = word_app.Documents.Open(str(detail_path.resolve()))
        doc_obj1.ActiveWindow.View.Type = 3
        for path in expense_paths:
            doc = word_app.Documents.Open(str(path.resolve()))
            doc.ActiveWindow.View.Type = 3
        doc_obj1.Activate()

if __name__ == "__main__":
    root = tk.Tk()
    app = PigSettlementApp(root)
    root.mainloop()
