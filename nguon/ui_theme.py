# -*- coding: utf-8 -*-
"""
UI THEME & MODERN WIDGETS FOR PHONG NCCN APPLICATION
====================================================
Bo giao dien hien dai phong cach Enterprise / Fluent Design cho Tkinter.
Cung cap:
  - Bang mau chuyen nghiep, font chu ro rang.
  - Card container bo vien tinh te voi Header badge dep mat.
  - Header app hien dai voi icon va badge trang thai.
  - Dark Slate Terminal voi thanh cong cu (Sao chep, Xoa, phan loai mau log).
  - Nut bam hanh dong noi bat, combobox va entry tinh te.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

# --- BANG MAU CHUAN (PALETTE) ---
COLOR_BG = "#F1F5F9"            # Nen tong the (Slate 100)
COLOR_SURFACE = "#FFFFFF"       # Nen card/khung (White)
COLOR_SURFACE_ALT = "#F8FAFC"   # Nen card phu (Slate 50)
COLOR_BORDER = "#E2E8F0"        # Vien card (Slate 200)
COLOR_BORDER_INPUT = "#CBD5E1"  # Vien o nhap (Slate 300)
COLOR_BORDER_FOCUS = "#3B82F6"  # Vien khi focus (Blue 500)

COLOR_PRIMARY = "#1E40AF"       # Xanh cong nghe chu dao (Blue 800)
COLOR_PRIMARY_HOVER = "#1D4ED8" # Hover (Blue 700)
COLOR_PRIMARY_ACTIVE = "#1E3A8A"# Active (Blue 900)
COLOR_PRIMARY_LIGHT = "#EFF6FF" # Nen nhe (Blue 50)

COLOR_ACCENT = "#0284C7"        # Cyan/Sky 600
COLOR_SUCCESS = "#059669"       # Emerald 600
COLOR_SUCCESS_BG = "#ECFDF5"    # Emerald 50
COLOR_WARNING = "#D97706"       # Amber 600
COLOR_WARNING_BG = "#FFFBEB"    # Amber 50
COLOR_DANGER = "#DC2626"        # Red 600
COLOR_DANGER_BG = "#FEF2F2"     # Red 50

COLOR_TEXT_MAIN = "#0F172A"     # Chu chinh (Slate 900)
COLOR_TEXT_MUTED = "#475569"    # Chu phu (Slate 600)
COLOR_TEXT_HINT = "#64748B"     # Chu goi y (Slate 500)
COLOR_TEXT_LIGHT = "#94A3B8"    # Chu nhe (Slate 400)

# Terminal Console Dark Palette
TERM_BG = "#0B1120"             # Dark navy slate
TERM_HEADER_BG = "#1E293B"      # Header console
TERM_BORDER = "#334155"         # Border console
TERM_FG_NORMAL = "#E2E8F0"      # Chu binh thuong
TERM_FG_SUCCESS = "#34D399"     # Xanh la thanh cong
TERM_FG_WARNING = "#FBBF24"     # Vang canh bao
TERM_FG_ERROR = "#F87171"       # Do loi
TERM_FG_ACCENT = "#38BDF8"      # Xanh troi noi bat
TERM_FG_MUTED = "#64748B"       # Xam ghi chu/gio


def _safe_map(s, name, **kwargs):
    if hasattr(s, "map"):
        try:
            s.map(name, **kwargs)
        except Exception:
            pass


# --- CUON TRANG BANG BANH XE CHUOT ---
# Moi tab la mot Canvas (trang) chua khung noi dung. Banh xe chuot cuon trang
# o MOI cho ben trong - ke ca tren o nhap, nhan, nut. Cach cu (bind_all khi
# <Enter>, unbind_all khi <Leave>) hong: Tk bao <Leave> cho nen trang ngay
# khi chuot di vao mot khung con, nen chi cuon duoc o khe giua cac khung.
#   - Tren bang / o chu (Treeview, Text, Listbox): cuon chinh no; no da het
#     (dau / cuoi) thi cuon tiep ca trang.
#   - Tren combobox: cuon trang, KHONG doi gia tri dang chon (Tk mac dinh
#     lan chuot la doi mac thep / bo cau hinh - de lo tay khi cuon trang).
_TU_CUON = ("Treeview", "Text", "Listbox")


def cuon_bang_chuot(canvas):
    """Danh dau canvas la mot trang cuon. Goi mot lan cho moi trang; goi lai
    chi thay dung binding chung nay (khong chong them)."""
    canvas._trang_cuon = True
    canvas.bind_all("<MouseWheel>", _banh_xe)
    canvas.bind_class("TCombobox", "<MouseWheel>", _banh_xe_combobox)


def _so_buoc(delta):
    """Touchpad gui delta nho (< 120) - van cuon it nhat mot buoc."""
    n = -int(delta / 120)
    return n or (-1 if delta > 0 else 1)


def _trang_chua(w):
    """(trang cuon chua w, bang / o chu gan nhat chua w hoac None)."""
    o = None
    while w is not None:
        if o is None and w.winfo_class() in _TU_CUON:
            o = w
        if getattr(w, "_trang_cuon", False):
            return w, o
        w = w.master
    return None, o


def _cuon_trang(trang, n):
    if trang.yview() != (0.0, 1.0):            # trang ngan hon cua so: dung yen
        trang.yview_scroll(n, "units")


def _banh_xe(e):
    if isinstance(e.widget, str):              # widget rieng cua Tk (vd danh sach tha xuong)
        return
    trang, o = _trang_chua(e.widget)
    if trang is None:
        return
    n = _so_buoc(e.delta)
    if o is not None:
        dau, cuoi = o.yview()
        if (n < 0 and dau > 0) or (n > 0 and cuoi < 1):
            return                             # bang con cuon duoc - lop cua no da cuon
    _cuon_trang(trang, n)


def _banh_xe_combobox(e):
    trang, _o = _trang_chua(e.widget)
    if trang is not None:
        _cuon_trang(trang, _so_buoc(e.delta))
    return "break"


def apply_app_styles(root):
    """Thiet lap Style ttk cho toan bo ung dung."""
    try:
        root.configure(bg=COLOR_BG)
    except Exception:
        pass

    s = ttk.Style()
    try:
        s.theme_use("clam")
    except Exception:
        pass

    # Basic Frame
    s.configure("TFrame", background=COLOR_BG)
    s.configure("Surface.TFrame", background=COLOR_SURFACE)
    s.configure("SurfaceAlt.TFrame", background=COLOR_SURFACE_ALT)

    # Basic Label
    s.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT_MAIN,
                font=("Segoe UI", 9))
    s.configure("Surface.TLabel", background=COLOR_SURFACE,
                foreground=COLOR_TEXT_MAIN, font=("Segoe UI", 9))
    s.configure("Phu.TLabel", background=COLOR_BG, foreground=COLOR_TEXT_HINT,
                font=("Segoe UI", 8))
    s.configure("SurfacePhu.TLabel", background=COLOR_SURFACE,
                foreground=COLOR_TEXT_HINT, font=("Segoe UI", 8))
    s.configure("Bold.TLabel", background=COLOR_SURFACE,
                foreground=COLOR_TEXT_MAIN, font=("Segoe UI", 9, "bold"))
    s.configure("CardTitle.TLabel", background=COLOR_SURFACE_ALT,
                foreground=COLOR_PRIMARY, font=("Segoe UI", 10, "bold"))

    # Notebook & Tabs
    s.configure("TNotebook", background=COLOR_BG, borderwidth=0, tabmargins=[6, 4, 6, 0])
    s.configure("TNotebook.Tab", background="#E2E8F0", foreground=COLOR_TEXT_MUTED,
                font=("Segoe UI", 9, "bold"), padding=[16, 7], borderwidth=0)
    _safe_map(s, "TNotebook.Tab",
              background=[("selected", COLOR_SURFACE), ("active", "#CBD5E1")],
              foreground=[("selected", COLOR_PRIMARY), ("active", COLOR_TEXT_MAIN)],
              expand=[("selected", [1, 2, 1, 0])])

    # Checkbuttons & Radiobuttons
    s.configure("TCheckbutton", background=COLOR_SURFACE, foreground=COLOR_TEXT_MAIN,
                font=("Segoe UI", 9))
    _safe_map(s, "TCheckbutton",
              background=[("active", COLOR_SURFACE)],
              foreground=[("active", COLOR_PRIMARY)])
    s.configure("Bg.TCheckbutton", background=COLOR_BG, foreground=COLOR_TEXT_MAIN,
                font=("Segoe UI", 9))

    s.configure("TRadiobutton", background=COLOR_SURFACE, foreground=COLOR_TEXT_MAIN,
                font=("Segoe UI", 9))
    _safe_map(s, "TRadiobutton",
              background=[("active", COLOR_SURFACE)],
              foreground=[("active", COLOR_PRIMARY)])

    # Entry
    s.configure("TEntry", fieldbackground="#FFFFFF", foreground=COLOR_TEXT_MAIN,
                insertcolor=COLOR_TEXT_MAIN, bordercolor=COLOR_BORDER_INPUT,
                lightcolor=COLOR_BORDER_INPUT, darkcolor=COLOR_BORDER_INPUT,
                padding=5)

    # Combobox
    s.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF",
                foreground=COLOR_TEXT_MAIN, arrowcolor=COLOR_PRIMARY,
                padding=4)
    _safe_map(s, "TCombobox",
              fieldbackground=[("readonly", "#FFFFFF")],
              selectbackground=[("readonly", COLOR_PRIMARY_LIGHT)],
              selectforeground=[("readonly", COLOR_PRIMARY)])

    # Buttons
    s.configure("TButton", font=("Segoe UI", 9), padding=[10, 5],
                background="#FFFFFF", foreground=COLOR_TEXT_MAIN,
                bordercolor=COLOR_BORDER_INPUT)
    _safe_map(s, "TButton",
              background=[("pressed", "#E2E8F0"), ("active", "#F1F5F9")],
              foreground=[("pressed", COLOR_PRIMARY), ("active", COLOR_PRIMARY)])

    # Primary Action Button
    s.configure("Primary.TButton", font=("Segoe UI", 10, "bold"),
                background=COLOR_PRIMARY, foreground="#FFFFFF",
                bordercolor=COLOR_PRIMARY_ACTIVE, padding=[16, 7])
    _safe_map(s, "Primary.TButton",
              background=[("pressed", COLOR_PRIMARY_ACTIVE), ("active", COLOR_PRIMARY_HOVER)],
              foreground=[("pressed", "#FFFFFF"), ("active", "#FFFFFF")])

    # Success Action Button
    s.configure("Success.TButton", font=("Segoe UI", 9, "bold"),
                background=COLOR_SUCCESS, foreground="#FFFFFF",
                bordercolor="#047857", padding=[12, 6])
    _safe_map(s, "Success.TButton",
              background=[("pressed", "#047857"), ("active", "#10B981")],
              foreground=[("pressed", "#FFFFFF"), ("active", "#FFFFFF")])

    # Danger Action Button (Dung)
    s.configure("Danger.TButton", font=("Segoe UI", 9, "bold"),
                background="#EF4444", foreground="#FFFFFF",
                bordercolor="#B91C1C", padding=[10, 5])
    _safe_map(s, "Danger.TButton",
              background=[("pressed", "#B91C1C"), ("active", "#DC2626"), ("disabled", "#FCA5A5")],
              foreground=[("disabled", "#FFFFFF")])

    # Subtle / Browse button
    s.configure("Outline.TButton", font=("Segoe UI", 8),
                background="#FFFFFF", foreground=COLOR_TEXT_MUTED,
                padding=[8, 3], bordercolor=COLOR_BORDER_INPUT)
    _safe_map(s, "Outline.TButton",
              background=[("pressed", "#E2E8F0"), ("active", "#F8FAFC")],
              foreground=[("active", COLOR_PRIMARY)])

    # Progressbar
    s.configure("TProgressbar", thickness=10, background=COLOR_PRIMARY,
                troughcolor="#E2E8F0", bordercolor=COLOR_BG)

    # Separator
    s.configure("TSeparator", background=COLOR_BORDER)


class ModernHeader(tk.Frame):
    """Thanh Header ung dung chuyen nghiep voi Title, Badge, va Subtitle."""
    def __init__(self, parent, title, subtitle="", badge_text="🟢 SẴN SÀNG", **kwargs):
        super(ModernHeader, self).__init__(parent, bg=COLOR_SURFACE, **kwargs)
        self.configure(highlightbackground=COLOR_BORDER, highlightthickness=1)

        bar = tk.Frame(self, bg=COLOR_SURFACE, padx=16, pady=10)
        bar.pack(fill="x")

        # Left: Icon & Title
        left = tk.Frame(bar, bg=COLOR_SURFACE)
        left.pack(side="left", fill="y")

        # App badge / Logo icon box
        icon_box = tk.Label(left, text="🔬", font=("Segoe UI Emoji", 16),
                            bg=COLOR_PRIMARY_LIGHT, fg=COLOR_PRIMARY,
                            padx=8, pady=4)
        icon_box.pack(side="left", padx=(0, 12))

        title_box = tk.Frame(left, bg=COLOR_SURFACE)
        title_box.pack(side="left", fill="y")

        lbl_tieu_de = tk.Label(title_box, text=title,
                               font=("Segoe UI", 12, "bold"),
                               fg=COLOR_TEXT_MAIN, bg=COLOR_SURFACE)
        lbl_tieu_de.pack(anchor="w")

        if subtitle:
            lbl_sub = tk.Label(title_box, text=subtitle,
                               font=("Segoe UI", 8),
                               fg=COLOR_TEXT_HINT, bg=COLOR_SURFACE)
            lbl_sub.pack(anchor="w", pady=(1, 0))

        # Right: Live status badge
        right = tk.Frame(bar, bg=COLOR_SURFACE)
        right.pack(side="right", fill="y", padx=(10, 0))

        self.badge = tk.Label(right, text=badge_text,
                              font=("Segoe UI", 8, "bold"),
                              bg=COLOR_SUCCESS_BG, fg=COLOR_SUCCESS,
                              padx=10, pady=4, relief="solid", borderwidth=1)
        self.badge.configure(highlightbackground=COLOR_SUCCESS, highlightthickness=0)
        self.badge.pack(side="right", pady=4)

    def set_status(self, text, loai="normal"):
        """Cap nhat badge trang thai (normal, success, warning, error, running)."""
        if loai == "success":
            self.badge.configure(text=text, bg=COLOR_SUCCESS_BG, fg=COLOR_SUCCESS)
        elif loai == "warning":
            self.badge.configure(text=text, bg=COLOR_WARNING_BG, fg=COLOR_WARNING)
        elif loai == "error":
            self.badge.configure(text=text, bg=COLOR_DANGER_BG, fg=COLOR_DANGER)
        elif loai == "running":
            self.badge.configure(text=text, bg=COLOR_PRIMARY_LIGHT, fg=COLOR_PRIMARY)
        else:
            self.badge.configure(text=text, bg=COLOR_SUCCESS_BG, fg=COLOR_SUCCESS)


class ModernCard(tk.Frame):
    """Khung Card bo vien tinh te thay the cho Labelframe tho cu."""
    def __init__(self, parent, title="", badge_num=None, **kwargs):
        super(ModernCard, self).__init__(parent, bg=COLOR_SURFACE, **kwargs)
        self.configure(highlightbackground=COLOR_BORDER, highlightthickness=1)

        if title or badge_num is not None:
            header = tk.Frame(self, bg=COLOR_SURFACE_ALT, padx=12, pady=6)
            header.pack(fill="x")
            header.configure(highlightbackground=COLOR_BORDER, highlightthickness=1)

            if badge_num is not None:
                lbl_badge = tk.Label(header, text=str(badge_num),
                                     font=("Segoe UI", 8, "bold"),
                                     bg=COLOR_PRIMARY, fg="#FFFFFF",
                                     padx=6, pady=1)
                lbl_badge.pack(side="left", padx=(0, 8))

            lbl_title = tk.Label(header, text=title,
                                 font=("Segoe UI", 9, "bold"),
                                 fg=COLOR_PRIMARY, bg=COLOR_SURFACE_ALT)
            lbl_title.pack(side="left")

        self.body = tk.Frame(self, bg=COLOR_SURFACE, padx=12, pady=10)
        self.body.pack(fill="both", expand=True)

    def get_body(self):
        return self.body


class ModernTerminal(tk.Frame):
    """Bang dieu khien Terminal/Log voi Header tool bar va phan mau dong."""
    def __init__(self, parent, title="NHẬT KÝ HOẠT ĐỘNG & KẾT QUẢ", height=9, **kwargs):
        super(ModernTerminal, self).__init__(parent, bg=TERM_BG, **kwargs)
        self.configure(highlightbackground=TERM_BORDER, highlightthickness=1)

        # Header bar cua terminal
        thdr = tk.Frame(self, bg=TERM_HEADER_BG, padx=10, pady=4)
        thdr.pack(fill="x")

        # Dot indicator
        self.dot = tk.Label(thdr, text="●", font=("Segoe UI", 10),
                            bg=TERM_HEADER_BG, fg=TERM_FG_SUCCESS)
        self.dot.pack(side="left", padx=(0, 6))

        lbl_t = tk.Label(thdr, text=title, font=("Consolas", 8, "bold"),
                         bg=TERM_HEADER_BG, fg=TERM_FG_NORMAL)
        lbl_t.pack(side="left")

        # Nut Clear log & Copy log
        btn_copy = tk.Button(thdr, text="📋 Sao chép", font=("Segoe UI", 7),
                             bg=TERM_HEADER_BG, fg=TERM_FG_NORMAL,
                             activebackground="#334155", activeforeground="#FFFFFF",
                             relief="flat", cursor="hand2", padx=6, pady=1,
                             command=self.copy_log)
        btn_copy.pack(side="right", padx=(4, 0))

        btn_clear = tk.Button(thdr, text="🧹 Xoá log", font=("Segoe UI", 7),
                              bg=TERM_HEADER_BG, fg=TERM_FG_NORMAL,
                              activebackground="#334155", activeforeground="#FFFFFF",
                              relief="flat", cursor="hand2", padx=6, pady=1,
                              command=self.clear_log)
        btn_clear.pack(side="right")

        # Body ScrolledText
        self.text = scrolledtext.ScrolledText(
            self, height=height, font=("Consolas", 9),
            bg=TERM_BG, fg=TERM_FG_NORMAL,
            insertbackground="#FFFFFF", selectbackground="#1E3A8A",
            selectforeground="#FFFFFF", relief="flat", borderwidth=0,
            padx=10, pady=8, wrap="word"
        )
        self.text.pack(fill="both", expand=True)

        # Setup Tag Colors
        if hasattr(self.text, "tag_config"):
            try:
                self.text.tag_config("normal", foreground=TERM_FG_NORMAL)
                self.text.tag_config("success", foreground=TERM_FG_SUCCESS, font=("Consolas", 9, "bold"))
                self.text.tag_config("warning", foreground=TERM_FG_WARNING, font=("Consolas", 9, "bold"))
                self.text.tag_config("error", foreground=TERM_FG_ERROR, font=("Consolas", 9, "bold"))
                self.text.tag_config("accent", foreground=TERM_FG_ACCENT, font=("Consolas", 9, "bold"))
                self.text.tag_config("muted", foreground=TERM_FG_MUTED)
            except Exception:
                pass

        try:
            self.text.configure(state="disabled")
        except Exception:
            pass

    def ghi(self, s, xoa=False, tag="normal"):
        try:
            self.text.configure(state="normal")
        except Exception:
            pass
        if xoa:
            try:
                self.text.delete("1.0", "end")
            except Exception:
                pass

        # Tu dong phan loai mau neu khong chi dinh tag cu the
        if tag == "normal":
            lower_s = s.lower()
            if any(k in lower_s for k in ["lỗi", "error", "không thành công", "thất bại", "không sinh được", "không ghi được"]):
                tag = "error"
            elif any(k in lower_s for k in ["hoàn tất", "thành công", "đã sinh xong", "khớp 100%", "đạt"]):
                tag = "success"
            elif any(k in lower_s for k in ["cảnh báo", "warning", "chú ý", "lệch"]):
                tag = "warning"
            elif any(k in lower_s for k in ["đang sinh", "bắt đầu", "sẵn sàng", "tiến hành"]):
                tag = "accent"
            elif s.startswith("(") and s.endswith(")\n"):
                tag = "muted"

        try:
            self.text.insert("end", s, tag)
        except Exception:
            try:
                self.text.insert("end", s)
            except Exception:
                pass
        try:
            self.text.see("end")
            self.text.configure(state="disabled")
        except Exception:
            pass

    def clear_log(self):
        try:
            self.text.configure(state="normal")
            self.text.delete("1.0", "end")
            self.text.configure(state="disabled")
        except Exception:
            pass

    def copy_log(self):
        try:
            self.clipboard_clear()
            self.clipboard_append(self.text.get("1.0", "end-1c"))
            messagebox.showinfo("Đã sao chép", "Đã sao chép toàn bộ nhật ký vào Clipboard.")
        except Exception:
            pass
