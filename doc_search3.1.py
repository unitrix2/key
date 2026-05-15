import os, subprocess, threading, re, hashlib, time
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from docx import Document
import openpyxl
from pptx import Presentation
from concurrent.futures import ThreadPoolExecutor

# Tesseract Path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ══ TOTP Config ══════════════════════════════════════════
TOTP_SALT   = "SALMAN-TOTP"
TOTP_PERIOD = 300   # 5 minutes

def totp_window():
    return int(time.time()) // TOTP_PERIOD

def totp_alpha():
    """MD5(salt + window) → first 6 chars uppercase"""
    raw = TOTP_SALT + str(totp_window())
    return hashlib.md5(raw.encode()).hexdigest()[:6].upper()

def totp_numeric():
    """SHA1(salt + window) → digits only → first 6"""
    raw = TOTP_SALT + str(totp_window())
    h   = hashlib.sha1(raw.encode()).hexdigest()
    digits = ''.join(c for c in h if c.isdigit())
    return digits[:6]

def totp_valid(entered):
    entered_clean = entered.strip().upper()
    return entered_clean == totp_alpha() or entered_clean == totp_numeric()

# ══ Accent colors ════════════════════════════════════════
CLR_BLUE   = "#3D82F7"
CLR_CYAN   = "#06C0E0"
CLR_PURPLE = "#8455FA"
CLR_GREEN  = "#10CC81"
CLR_RED    = "#F04545"
CLR_ORANGE = "#F5A020"

FILE_CLRS  = {".pdf":"#FF5252",".docx":"#4A9EFF",".xlsx":"#22D66E",
              ".pptx":"#FF9E3A",".txt":"#B09AF0",
              ".jpg":"#F06DB0",".jpeg":"#F06DB0",".png":"#F06DB0",".bmp":"#F06DB0"}
FILE_ICONS = {".pdf":"📕",".docx":"📘",".xlsx":"📗",".pptx":"📙",".txt":"📄",
              ".jpg":"🖼️",".jpeg":"🖼️",".png":"🖼️",".bmp":"🖼️"}

# ══ Theme palettes ════════════════════════════════════════
DARK = {
    "bg_deep" :"#0B0E17", "bg_side" :"#101420", "bg_main" :"#0D1119",
    "bg_card" :"#171C2E", "bg_input":"#1C2236", "txt_a"   :"#EDF0FF",
    "txt_b"   :"#9BAEC8", "txt_c"   :"#4E5E7A", "brd_dim" :"#1E2840",
    "brd_brt" :"#2E4268", "sig_card":"#131828", "sw_track":"#243250",
    "cb_border":"#3D5E98", "btn_fg"  :"#182036", "btn_txt" :"#C8D8F0",
    "stop_fg" :"#220A0A", "snip_bg" :"#171C2E", "snip_fg" :"#9BAEC8",
}

LIGHT = {
    "bg_deep" :"#DDE4F0", "bg_side" :"#C8D2E4", "bg_main" :"#EDF1FA",
    "bg_card" :"#FFFFFF", "bg_input":"#E2EAF8", "txt_a"   :"#05080F",
    "txt_b"   :"#111928", "txt_c"   :"#2C3E5C", "brd_dim" :"#8A9EC0",
    "brd_brt" :"#4060A8", "sig_card":"#B8C8E0", "sw_track":"#5A7AB8",
    "cb_border":"#2848A0", "btn_fg"  :"#D4DCF0", "btn_txt" :"#060C1E",
    "stop_fg" :"#FFE0E0", "snip_bg" :"#FFFFFF", "snip_fg" :"#111928",
}

class UDI_Pro_1(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Doc Search  —  Version 3.1")
        self.geometry("1550x950")

        self.is_dark = True
        self.C = dict(DARK)
        ctk.set_appearance_mode("Dark")
        self.configure(fg_color=self.C["bg_deep"])
        self.withdraw()   

        self.stop_event  = threading.Event()
        self.target_path = ""
        self.hit_data    = []
        self._sw_widgets = []
        self._cb_widgets = []

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ════════════════════ SIDEBAR ════════════════════
        self.sidebar = ctk.CTkScrollableFrame(
            self, width=335, corner_radius=0,
            fg_color=self.C["bg_side"],
            scrollbar_button_color=self.C["brd_dim"],
            scrollbar_button_hover_color=self.C["brd_brt"]
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        logo_f = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_f.pack(fill="x", padx=26, pady=(38, 0))
        ctk.CTkLabel(logo_f, text="DOC", font=("Segoe UI Black", 54, "bold"),
                     text_color=CLR_BLUE).pack(anchor="w")
        self.search_lbl = ctk.CTkLabel(logo_f, text="SEARCH", font=("Segoe UI", 22, "bold"),
                     text_color=self.C["txt_b"])
        self.search_lbl.pack(anchor="w")
        ctk.CTkFrame(self.sidebar, height=2, fg_color=CLR_BLUE,
                     corner_radius=0).pack(fill="x", padx=26, pady=(10, 4))
        self.ver_lbl = ctk.CTkLabel(self.sidebar, text="v3.1  ·  Intelligent File Discovery",
                     font=("Segoe UI", 11), text_color=self.C["txt_b"])
        self.ver_lbl.pack(anchor="w", padx=28, pady=(0, 22))

        self._section("⚡   SEARCH ENGINES")
        self.digital_sw = self.create_sw("Digital Search Engine",   True,  CLR_CYAN)
        self.ocr_sw     = self.create_sw("Optical (OCR) Engine",    False, CLR_ORANGE)
        self.exact_sw   = self.create_sw("Exact Match Only",        False, CLR_RED)
        self.fuzzy_sw   = self.create_sw("Fuzzy Search (Choice B)", False, CLR_PURPLE)

        self._section("🗂️   FILE TYPES")
        self.uncheck_lbl = ctk.CTkLabel(self.sidebar, text="Uncheck to skip — saves scan time",
                     font=("Segoe UI", 11, "italic"),
                     text_color=self.C["txt_b"])
        self.uncheck_lbl.pack(anchor="w", padx=32, pady=(0, 8))
        self.pdf_on   = self.create_cb("📕  PDF Documents  (.pdf)",   True)
        self.word_on  = self.create_cb("📘  Word Files  (.docx)",      True)
        self.excel_on = self.create_cb("📗  Excel Sheets  (.xlsx)",    True)
        self.ppt_on   = self.create_cb("📙  PPT Slides  (.pptx)",      True)
        self.txt_on   = self.create_cb("📄  Text Files  (.txt)",       True)
        self.img_on   = self.create_cb("🖼️  Images / Photos  (OCR)",  True)

        self._section("🚀   PERFORMANCE")
        self.multicore_sw = self.create_sw("Multi-Core Boost", True, CLR_PURPLE)

        ctk.CTkFrame(self.sidebar, height=1, fg_color=self.C["brd_dim"]
                     ).pack(fill="x", padx=20, pady=(30, 0))
        sig = ctk.CTkFrame(self.sidebar, fg_color=self.C["sig_card"],
                           corner_radius=14, border_width=1, border_color=CLR_CYAN)
        sig.pack(pady=20, padx=20, fill="x")
        self.created_lbl = ctk.CTkLabel(sig, text="✦  Created by",
                     font=("Segoe UI", 11), text_color=self.C["txt_b"])
        self.created_lbl.pack(pady=(18, 0))
        ctk.CTkLabel(sig, text="Salman",
                     font=("Lucida Handwriting", 23, "bold"),
                     text_color=CLR_CYAN).pack(pady=(2, 4))
        ctk.CTkLabel(sig, text="North Central Railway",
                     font=("Segoe UI", 11), text_color=self.C["txt_c"]).pack(pady=(0, 18))

        # ════════════════════ MAIN PANEL ════════════════════
        self.main_panel = ctk.CTkFrame(self, fg_color=self.C["bg_main"], corner_radius=0)
        self.main_panel.grid(row=0, column=1, sticky="nsew")
        inner = ctk.CTkFrame(self.main_panel, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=50, pady=45)

        hdr = ctk.CTkFrame(inner, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 18))
        self.hdr_lbl = ctk.CTkLabel(hdr, text="Intelligence Discovery Engine",
                     font=("Segoe UI", 14), text_color=self.C["txt_b"], anchor="w")
        self.hdr_lbl.pack(side="left")
        self.theme_btn = ctk.CTkButton(
            hdr, text="☀️  Light Mode", command=self.toggle_theme,
            width=135, height=34, corner_radius=17,
            fg_color=self.C["brd_dim"], hover_color=self.C["brd_brt"],
            font=("Segoe UI", 12), text_color=self.C["txt_b"],
            border_width=1, border_color=self.C["brd_brt"]
        )
        self.theme_btn.pack(side="right")

        self.search_bar = ctk.CTkEntry(
            inner, placeholder_text="🔍   Enter keyword to search…",
            height=66, font=("Segoe UI", 17),
            fg_color=self.C["bg_input"], border_color=self.C["brd_brt"],
            border_width=2, text_color=self.C["txt_a"],
            placeholder_text_color=self.C["txt_c"], corner_radius=12
        )
        self.search_bar.pack(fill="x", pady=(0, 14))

        btn_row = ctk.CTkFrame(inner, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 10))
        self.folder_btn = ctk.CTkButton(
            btn_row, text="📁  Folder", command=self.pick_folder,
            width=130, height=46, corner_radius=10,
            fg_color=self.C["btn_fg"], hover_color=self.C["brd_brt"],
            font=("Segoe UI", 13), text_color=self.C["btn_txt"],
            border_width=1, border_color=self.C["brd_brt"]
        )
        self.folder_btn.pack(side="left", padx=(0, 8))
        self.file_btn = ctk.CTkButton(
            btn_row, text="📄  File", command=self.pick_file,
            width=118, height=46, corner_radius=10,
            fg_color=self.C["btn_fg"], hover_color=self.C["brd_brt"],
            font=("Segoe UI", 13), text_color=self.C["btn_txt"],
            border_width=1, border_color=self.C["brd_brt"]
        )
        self.file_btn.pack(side="left", padx=(0, 16))
        self.start_btn = ctk.CTkButton(
            btn_row, text="⚡  EXECUTE DISCOVERY", command=self.initiate_search,
            fg_color=CLR_BLUE, hover_color="#2563EB",
            height=46, font=("Segoe UI", 14, "bold"), corner_radius=10
        )
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.stop_btn = ctk.CTkButton(
            btn_row, text="■  STOP", command=self.stop_search,
            fg_color=self.C["stop_fg"], hover_color=CLR_RED,
            width=95, height=46, corner_radius=10,
            font=("Segoe UI", 13, "bold"), text_color=CLR_RED,
            border_width=1, border_color=CLR_RED
        )
        self.stop_btn.pack(side="right")

        self.status_frame = ctk.CTkFrame(inner, fg_color=self.C["bg_card"],
                                          corner_radius=10, border_width=1,
                                          border_color=self.C["brd_dim"])
        self.status_frame.pack(fill="x", pady=(0, 8))
        self.file_label = ctk.CTkLabel(
            self.status_frame, text="●  System Ready",
            font=("Consolas", 12), text_color=CLR_GREEN, anchor="w"
        )
        self.file_label.pack(side="left", padx=18, pady=10)
        self.percent_label = ctk.CTkLabel(
            self.status_frame, text="—",
            font=("Consolas", 13, "bold"), text_color=CLR_CYAN
        )
        self.percent_label.pack(side="right", padx=18)

        self.p_bar = ctk.CTkProgressBar(
            inner, height=6, progress_color=CLR_BLUE,
            fg_color=self.C["brd_dim"], corner_radius=3
        )
        self.p_bar.pack(fill="x", pady=(0, 16))
        self.p_bar.set(0)

        self.results_scroll = ctk.CTkScrollableFrame(
            inner, label_text="   📋  RESULTS",
            label_font=("Segoe UI", 13, "bold"),
            label_fg_color=self.C["bg_card"], label_text_color=self.C["txt_b"],
            fg_color=self.C["bg_deep"], border_width=1, border_color=self.C["brd_dim"],
            corner_radius=12, scrollbar_button_color=self.C["brd_dim"],
            scrollbar_button_hover_color=self.C["brd_brt"],
        )
        self.results_scroll.pack(fill="both", expand=True)

        self.after(0, self._show_login)

    # ══════════════════ LOGIN SCREEN (SPACIOUS UI) ══════════════
    def _show_login(self):
        self._login_win = ctk.CTkToplevel(self)
        lw = self._login_win
        lw.title("Doc Search — Secure Access")
        lw.geometry("480x660") # Increased window width
        lw.resizable(False, False)
        lw.configure(fg_color="#080C16")
        lw.grab_set()
        lw.protocol("WM_DELETE_WINDOW", self.destroy)

        lw.update_idletasks()
        sw = lw.winfo_screenwidth(); sh = lw.winfo_screenheight()
        lw.geometry(f"480x660+{(sw-480)//2}+{(sh-660)//2}")

        outer = tk.Frame(lw, bg="#0F1828", bd=0)
        outer.place(relx=0.5, rely=0.5, anchor="center", width=420, height=610) # Wider outer frame
        tk.Frame(outer, bg="#3D82F7", height=2).pack(fill="x")

        card = tk.Frame(outer, bg="#0C1220", bd=0)
        card.pack(fill="both", expand=True)

        tk.Label(card, text="🔐", font=("Segoe UI", 28), bg="#0C1220", fg="#3D82F7").pack(pady=(28, 0))
        tk.Label(card, text="DOC SEARCH", font=("Segoe UI Black", 26, "bold"), bg="#0C1220", fg="#EDF0FF").pack(pady=(4, 0))
        tk.Label(card, text="S E C U R E   A C C E S S", font=("Segoe UI", 10), bg="#0C1220", fg="#2A3E60").pack()
        tk.Frame(card, bg="#151F35", height=1).pack(fill="x", padx=30, pady=(16, 0))
        tk.Label(card, text="Enter your 6-character access code", font=("Segoe UI", 11), bg="#0C1220", fg="#2E4260").pack(pady=(12, 0))

        # SPACIOUS CODE BOXES
        boxes_frame = tk.Frame(card, bg="#0C1220")
        boxes_frame.pack(pady=(22, 0))

        self._box_vars = []
        self._box_entries = []

        def on_key(event, idx):
            ch = event.char
            if not ch or not ch.isprintable(): return "break"
            ch = ch.upper()
            lw.after_idle(lambda ix=idx, c=ch: _set_and_move(ix, c))
            return "break"

        def _set_and_move(idx, ch):
            self._box_vars[idx].set(ch)
            if idx < 5: self._box_entries[idx+1].focus_set()

        def on_back(event, idx):
            if self._box_vars[idx].get(): self._box_vars[idx].set("")
            elif idx > 0:
                self._box_vars[idx-1].set("")
                self._box_entries[idx-1].focus_set()
            return "break"

        def on_ctrl_v(event, idx):
            try:
                clip = lw.clipboard_get().strip().upper()[:6]
                for i in range(6): self._box_vars[i].set(clip[i] if i < len(clip) else "")
                last = min(len(clip), 5); self._box_entries[last].focus_set()
            except: pass
            return "break"

        for i in range(6):
            var = tk.StringVar(); self._box_vars.append(var)
            # Fixed width to prevent cutting
            e = tk.Entry(boxes_frame, textvariable=var, width=2, font=("Segoe UI Black", 24),
                         justify="center", bg="#111828", fg="#EDF0FF", insertbackground="#3D82F7",
                         relief="flat", highlightthickness=2, highlightbackground="#1E2E50", highlightcolor="#3D82F7")
            e.grid(row=0, column=i, padx=6) # Increased padding between boxes
            e.bind("<KeyPress>", lambda ev, ix=i: on_key(ev, ix))
            e.bind("<BackSpace>", lambda ev, ix=i: on_back(ev, ix))
            e.bind("<Control-v>", lambda ev, ix=i: on_ctrl_v(ev, ix))
            e.bind("<Control-V>", lambda ev, ix=i: on_ctrl_v(ev, ix))
            self._box_entries.append(e)

        self._box_entries[0].focus()
        btn_row = tk.Frame(card, bg="#0C1220"); btn_row.pack(pady=(22, 0))

        def do_paste():
            try:
                clip = lw.clipboard_get().strip().upper()[:6]
                for i in range(6): self._box_vars[i].set(clip[i] if i < len(clip) else "")
                last = min(len(clip), 5); self._box_entries[last].focus_set()
            except: pass

        def do_clear():
            for v in self._box_vars: v.set("")
            self._box_entries[0].focus()
            if hasattr(self, '_login_err_lbl'): self._login_err_lbl.config(text="")

        tk.Button(btn_row, text="⎘  Paste", command=do_paste, font=("Segoe UI", 11), bg="#131C30", fg="#4A7AC8", relief="flat", padx=16, pady=8, cursor="hand2", bd=0).grid(row=0, column=0, padx=(0, 8))
        tk.Button(btn_row, text="✕  Clear", command=do_clear, font=("Segoe UI", 11), bg="#131C30", fg="#4A5A70", relief="flat", padx=16, pady=8, cursor="hand2", bd=0).grid(row=0, column=1, padx=(0, 8))
        tk.Button(btn_row, text="⚡  Authenticate", command=self._do_login, font=("Segoe UI", 12, "bold"), bg="#2A5FCC", fg="#FFFFFF", relief="flat", padx=20, pady=8, cursor="hand2", bd=0).grid(row=0, column=2)

        self._login_err_lbl = tk.Label(card, text="", font=("Segoe UI", 11), bg="#0C1220", fg="#E05050")
        self._login_err_lbl.pack(pady=(12, 0))
        self._timer_lbl_tk = tk.Label(card, text="", font=("Consolas", 10), bg="#0C1220", fg="#1E3050")
        self._timer_lbl_tk.pack(pady=(4, 0))
        tk.Frame(card, bg="#0F1828", height=1).pack(fill="x", padx=30, pady=(18, 0))

        link_row = tk.Frame(card, bg="#0C1220"); link_row.pack(pady=(12, 0))
        tk.Label(link_row, text="Need the code?", font=("Segoe UI", 10), bg="#0C1220", fg="#1E3050").pack(side="left", padx=(0,8))
        tk.Button(link_row, text="Open Code Generator →", command=lambda: subprocess.Popen(['start', '', 'https://cclkey.pages.dev/'], shell=True), font=("Segoe UI", 10, "bold"), bg="#0C1220", fg="#2A5FAA", relief="flat", cursor="hand2", bd=0).pack(side="left")

        self._update_login_timer()

    def _update_login_timer(self):
        if not hasattr(self, '_login_win') or not self._login_win.winfo_exists(): return
        remaining = TOTP_PERIOD - (int(time.time()) % TOTP_PERIOD)
        m, s = divmod(remaining, 60)
        if hasattr(self, '_timer_lbl_tk') and self._timer_lbl_tk.winfo_exists():
            self._timer_lbl_tk.config(text=f"⏱  Refreshes in  {m}:{s:02d}", fg="#1E3A60" if remaining > 60 else "#6B2020")
        self._login_win.after(1000, self._update_login_timer)

    def _do_login(self):
        parts = [v.get().strip() for v in self._box_vars]
        entered = "".join(parts)
        if len(entered) < 6:
            if hasattr(self, '_login_err_lbl'): self._login_err_lbl.config(text="⚠  Fill all 6 boxes")
            return
        if totp_valid(entered):
            self._login_win.destroy(); self.deiconify(); self.focus_force()
        else:
            if hasattr(self, '_login_err_lbl'): self._login_err_lbl.config(text="✗  Wrong code — try again")
            for v in self._box_vars: v.set(""); self._box_entries[0].focus_set()

    # ══════════════════ THEME TOGGLE ══════════════════
    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self.C = dict(DARK if self.is_dark else LIGHT)
        ctk.set_appearance_mode("Dark" if self.is_dark else "Light")
        self.configure(fg_color=self.C["bg_deep"])
        self.sidebar.configure(fg_color=self.C["bg_side"])
        self.main_panel.configure(fg_color=self.C["bg_main"])
        self.status_frame.configure(fg_color=self.C["bg_card"], border_color=self.C["brd_dim"])
        self.search_bar.configure(fg_color=self.C["bg_input"], border_color=self.C["brd_brt"], text_color=self.C["txt_a"], placeholder_text_color=self.C["txt_c"])
        self.p_bar.configure(fg_color=self.C["brd_dim"])
        self.results_scroll.configure(fg_color=self.C["bg_deep"], border_color=self.C["brd_dim"], label_fg_color=self.C["bg_card"], label_text_color=self.C["txt_b"])
        for sw, col in self._sw_widgets: sw.configure(text_color=self.C["txt_a"], fg_color=self.C["sw_track"], progress_color=col)
        for cb in self._cb_widgets: cb.configure(text_color=self.C["txt_b"], border_color=self.C["cb_border"])
        self.folder_btn.configure(fg_color=self.C["btn_fg"], hover_color=self.C["brd_brt"], text_color=self.C["btn_txt"], border_color=self.C["brd_brt"])
        self.file_btn.configure(fg_color=self.C["btn_fg"], hover_color=self.C["brd_brt"], text_color=self.C["btn_txt"], border_color=self.C["brd_brt"])
        self.stop_btn.configure(fg_color=self.C["stop_fg"])
        self.hdr_lbl.configure(text_color=self.C["txt_b"])
        self.uncheck_lbl.configure(text_color=self.C["txt_b"])
        self.ver_lbl.configure(text_color=self.C["txt_b"])
        self.created_lbl.configure(text_color=self.C["txt_b"])
        self.search_lbl.configure(text_color=self.C["txt_b"])
        self.theme_btn.configure(text="☀️  Light Mode" if self.is_dark else "🌙  Dark Mode", fg_color=self.C["brd_dim"], hover_color=self.C["brd_brt"], text_color=self.C["txt_b"], border_color=self.C["brd_brt"])

    # ══════════════════ UI HELPERS ══════════════════
    def _section(self, title):
        f = ctk.CTkFrame(self.sidebar, fg_color="transparent"); f.pack(fill="x", padx=20, pady=(26, 8))
        ctk.CTkLabel(f, text=title, font=("Segoe UI", 11, "bold"), text_color=self.C["txt_b"]).pack(side="left")
        ctk.CTkFrame(f, height=1, fg_color=self.C["brd_dim"]).pack(side="left", fill="x", expand=True, padx=(10, 0), pady=1)

    def create_sw(self, txt, def_v, col):
        sw = ctk.CTkSwitch(self.sidebar, text=txt, progress_color=col, fg_color=self.C["sw_track"], font=("Segoe UI", 13), text_color=self.C["txt_a"])
        sw.select() if def_v else sw.deselect()
        sw.pack(padx=28, anchor="w", pady=6); self._sw_widgets.append((sw, col))
        return sw

    def create_cb(self, txt, def_v):
        cb = ctk.CTkCheckBox(self.sidebar, text=txt, font=("Segoe UI", 13), text_color=self.C["txt_b"], fg_color=CLR_BLUE, border_color=self.C["cb_border"], corner_radius=5)
        cb.select() if def_v else cb.deselect()
        cb.pack(padx=38, anchor="w", pady=4); self._cb_widgets.append(cb)
        return cb

    # ══════════════════ CORE LOGIC ══════════════════
    def pick_folder(self):
        p = filedialog.askdirectory()
        if p: self.target_path = p; self.file_label.configure(text=f"●  Folder: {os.path.basename(p)}")

    def pick_file(self):
        p = filedialog.askopenfilename()
        if p: self.target_path = p; self.file_label.configure(text=f"●  File: {os.path.basename(p)}")

    def stop_search(self): self.stop_event.set()

    def normalize_hindi(self, text):
        replacements = {'क़':'क','ख़':'ख','ग़':'ग','ज़':'ज','ड़':'ड','ढ़':'ढ','फ़':'फ'}
        for k, v in replacements.items(): text = text.replace(k, v)
        return text

    def check_match(self, query, text):
        q_norm = self.normalize_hindi(query.lower())
        t_norm = self.normalize_hindi(text.lower())
        if self.exact_sw.get():
            return bool(re.search(rf'\b{re.escape(q_norm)}\b', t_norm))
        return q_norm in t_norm

    def initiate_search(self):
        query = self.search_bar.get()
        if not self.target_path or not query: messagebox.showwarning("!", "Target aur Keyword dono zaruri hain!"); return
        for w in self.results_scroll.winfo_children(): w.destroy()
        self.stop_event.clear(); self.hit_data = []
        threading.Thread(target=self.engine_core, daemon=True).start()

    def engine_core(self):
        query = self.search_bar.get(); norm_q = self.normalize_hindi(query.lower())
        files = []
        if os.path.isfile(self.target_path): files.append(self.target_path)
        else:
            for r, d, fs in os.walk(self.target_path):
                for f in fs: files.append(os.path.join(r, f))
        
        total = len(files)
        workers = os.cpu_count() if self.multicore_sw.get() else 1
        with ThreadPoolExecutor(max_workers=workers) as executor:
            for i, path in enumerate(files):
                if self.stop_event.is_set(): break
                prog = (i+1)/total
                self.after(0, lambda v=prog, f=os.path.basename(path), pl=int(prog*100): (
                    self.p_bar.set(v), self.file_label.configure(text=f"●  Scanning: {f[:30]}…"), self.percent_label.configure(text=f"{pl}%")))
                ext = path.lower()
                try:
                    if ext.endswith('.pdf') and self.pdf_on.get():
                        doc = fitz.open(path)
                        for pg_n, page in enumerate(doc):
                            txt = page.get_text()
                            if self.digital_sw.get() and self.check_match(query, txt):
                                for m in re.finditer(re.escape(norm_q), self.normalize_hindi(txt.lower())):
                                    s, e = max(0, m.start()-60), min(len(txt), m.end()+100)
                                    self.after(0, lambda p=path, l=f"Page {pg_n+1}", s=txt[s:e].replace('\n',' '): self.add_result_ui(p, l, s))
                            elif self.ocr_sw.get():
                                pix = page.get_pixmap(dpi=150)
                                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                                ocr_t = pytesseract.image_to_string(img, lang='eng+hin', config='--psm 3')
                                if self.check_match(query, ocr_t):
                                    self.after(0, lambda p=path, l=f"OCR Page {pg_n+1}", s="Scanned PDF Match": self.add_result_ui(p, l, s))
                        doc.close()
                    elif ext.endswith('.docx') and self.word_on.get():
                        doc = Document(path)
                        for idx, p in enumerate(doc.paragraphs):
                            if self.check_match(query, p.text): self.after(0, lambda p=path, l=f"Para {idx+1}", s=p.text: self.add_result_ui(p, l, s))
                    elif ext.endswith('.xlsx') and self.excel_on.get():
                        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                        for sheet in wb.worksheets:
                            for r_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                                if not any(row): continue
                                row_txt = " | ".join([str(c) for c in row if c is not None])
                                if self.check_match(query, row_txt):
                                    self.after(0, lambda p=path, l=f"Excel: {sheet.title} R{r_idx+1}", s=row_txt: self.add_result_ui(p, l, s))
                        wb.close()
                    elif ext.endswith('.txt') and self.txt_on.get():
                        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                            for idx, line in enumerate(f):
                                if self.check_match(query, line): self.after(0, lambda p=path, l=f"Line {idx+1}", s=line.strip(): self.add_result_ui(p, l, s))
                except: pass
        self.after(0, lambda: messagebox.showinfo("Done", f"Discovery Complete! Hits: {len(self.hit_data)}"))

    def add_result_ui(self, f_path, loc, snip):
        self.hit_data.append(f_path)
        ext = os.path.splitext(f_path)[1].lower()
        stripe_color = FILE_CLRS.get(ext, "#4455AA")
        file_icon = FILE_ICONS.get(ext, "📄")

        card = ctk.CTkFrame(self.results_scroll, fg_color=self.C["bg_card"], corner_radius=12, border_width=1, border_color=self.C["brd_dim"])
        card.pack(fill="x", pady=5, padx=12)
        inner_row = ctk.CTkFrame(card, fg_color="transparent"); inner_row.pack(fill="both", expand=True)
        ctk.CTkFrame(inner_row, width=5, corner_radius=0, fg_color=stripe_color).pack(side="left", fill="y")
        
        info = ctk.CTkFrame(inner_row, fg_color="transparent"); info.pack(side="left", fill="both", expand=True, padx=16, pady=14)
        name_row = ctk.CTkFrame(info, fg_color="transparent"); name_row.pack(fill="x", anchor="w")
        ctk.CTkLabel(name_row, text=file_icon, font=("Segoe UI", 16)).pack(side="left", padx=(0, 7))
        ctk.CTkLabel(name_row, text=os.path.basename(f_path), font=("Segoe UI", 14, "bold"), text_color=self.C["txt_a"]).pack(side="left")

        badge = ctk.CTkFrame(info, fg_color=self.C["bg_input"], corner_radius=6); badge.pack(anchor="w", pady=(4, 0))
        ctk.CTkLabel(badge, text=f"  📍 {loc}  ", font=("Segoe UI", 11), text_color=CLR_CYAN).pack(padx=2, pady=2)

        snippet_box = tk.Text(info, height=3, bg=self.C["snip_bg"], fg=self.C["snip_fg"], font=("Consolas", 10), borderwidth=0, highlightthickness=0, wrap="word", padx=4, pady=4)
        snippet_box.pack(fill="x", pady=(8, 0)); snippet_box.insert("1.0", f"...{snip}...")
        
        query = self.search_bar.get().lower(); idx = "1.0"
        while True:
            idx = snippet_box.search(query, idx, nocase=True, stopindex="end")
            if not idx: break
            lastidx = f"{idx}+{len(query)}c"; snippet_box.tag_add("hl", idx, lastidx); idx = lastidx
        snippet_box.tag_config("hl", foreground="#00FF94", font=("Consolas", 10, "bold"))
        snippet_box.configure(state="disabled")

        btns = ctk.CTkFrame(inner_row, fg_color="transparent"); btns.pack(side="right", padx=14, pady=14, anchor="n")
        ctk.CTkButton(btns, text="Open", width=74, height=32, fg_color=CLR_BLUE, font=("Segoe UI", 12, "bold"), corner_radius=8, command=lambda: os.startfile(f_path)).pack(pady=(0, 6))
        ctk.CTkButton(btns, text="📁", width=44, height=32, fg_color=self.C["brd_dim"], corner_radius=8, command=lambda: subprocess.run(['explorer', '/select,', os.path.normpath(f_path)])).pack()

if __name__ == "__main__":
    app = UDI_Pro_1(); app.mainloop()