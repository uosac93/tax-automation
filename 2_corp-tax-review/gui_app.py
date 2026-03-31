"""
Corp Tax_AI - 법인세 신고서 자동 검토 (GUI)
Property Tax AI 디자인 동일 적용
"""
import os
import sys
import re
import threading
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

import ctypes
from ctypes import wintypes

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES

from pdf_parser import parse_all
from tax_reviewer import run_all_reviews, format_won
from report_generator import generate_excel_report
from report_pdf import generate_report_pdf
from report_docx import generate_report_docx

# ── 테마 ──
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# ── Property Tax AI CSS 변수 (파란색 변환) ──
BG = "#191919"
SURFACE = "#232323"
SIDEBAR_BG = "#1E1E1E"
BORDER = "#333333"
TEXT = "#E8E8E8"
SUB = "#999999"
ACCENT = "#5B8DB8"
ACCENT_LIGHT = "rgba(91,141,184,0.06)"  # for result-box
INPUT_BG = "#1E1E1E"
OK = "#5BBC6E"
ERR = "#E55B5B"
WARN = "#D4A27A"
ACCENT_HOVER = "#6DA0CC"
# scrollbar: #30363D (Property Tax AI)
SB_COLOR = "#30363D"
SB_HOVER = "#484F58"

# ── 폰트 ──
FONT = "바탕"
SZ_TITLE = 16    # page-title 15px ≈ 16pt
SZ_BODY = 13     # body 13px
SZ_SIDEBAR = 13  # sidebar-item 13px
SZ_SECTION = 11  # section-title 11px
SZ_SMALL = 12
SZ_TINY = 11
SZ_BIG = 22


def _short_corp_name(name):
    """주식회사이노덴트 → (주)이노덴트"""
    if not name:
        return name
    import re
    m = re.match(r'^주식회사\s*(.+)$', name)
    if m:
        return f"(주){m.group(1)}"
    m = re.match(r'^(.+?)\s*주식회사$', name)
    if m:
        return f"{m.group(1)}(주)"
    return name


class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)
        self.title("Corp Tax_AI")
        self.geometry("1129x750")
        self.minsize(900, 600)
        if getattr(sys, 'frozen', False):
            base = sys._MEIPASS
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        ico = os.path.join(base, "tax_review.ico")
        if os.path.exists(ico):
            self.iconbitmap(ico)
            self.after(200, lambda: self.iconbitmap(ico))
        self.configure(fg_color=BG)
        try:
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(ctypes.c_int(1)), 4)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(ctypes.c_int(0x00000000)), 4)
        except Exception:
            pass
        self.data = None
        self.results = None
        self.pdf_path = None
        self.excel_path = None
        self._build()
        self._page_home()
        # OLE 드래그 & 드롭
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._on_dnd_drop)
        self.dnd_bind('<<DropEnter>>', self._on_dnd_enter)
        self.dnd_bind('<<DropLeave>>', self._on_dnd_leave)

    # ════════════════════════════════════════
    #  드래그 & 드롭 (tkinterdnd2 OLE)
    # ════════════════════════════════════════
    def _on_dnd_drop(self, event):
        files = self.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith('.pdf'):
                self._setpdf(f)
                break

    def _on_dnd_enter(self, event):
        if hasattr(self, 'dropzone'):
            self.dropzone.configure(border_color="#8BC4F0", border_width=2)
        if hasattr(self, 'fl'):
            self.fl.configure(text="파일을 여기에 놓으세요!", text_color=ACCENT)

    def _on_dnd_leave(self, event):
        if hasattr(self, 'dropzone') and not self.pdf_path:
            self.dropzone.configure(border_color=BORDER, border_width=1)
        if hasattr(self, 'fl') and not self.pdf_path:
            self.fl.configure(text="PDF 파일을 여기에 드래그하세요", text_color=SUB)

    # ════════════════════════════════════════
    #  사이드바
    # ════════════════════════════════════════
    def _build(self):
        sb = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color=SIDEBAR_BG, border_width=0)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        self.sb = sb

        ctk.CTkFrame(self, width=1, fg_color=BORDER, corner_radius=0).pack(side="left", fill="y")

        ctk.CTkLabel(sb, text="박양훈 세무사", font=(FONT, SZ_BODY),
                     text_color=TEXT).pack(anchor="w", padx=16, pady=(16, 12))

        ctk.CTkLabel(sb, text="법인세", font=(FONT, SZ_SECTION, "bold"),
                     text_color=ACCENT).pack(anchor="w", padx=16, pady=(6, 4))
        self.b1 = self._sbtn(sb, "신고서 검토", self._page_home)
        self.b2 = self._sbtn(sb, "검토 결과", self._page_results)
        self.b5 = self._sbtn(sb, "보고서", self._page_report)

        ctk.CTkLabel(sb, text="도구", font=(FONT, SZ_SECTION, "bold"),
                     text_color=ACCENT).pack(anchor="w", padx=16, pady=(6, 4))
        self.b3 = self._sbtn(sb, "검토 목록", self._page_checklist)
        self.b4 = self._sbtn(sb, "파싱 데이터", self._page_data)

        footer = ctk.CTkFrame(sb, fg_color="transparent")
        footer.pack(side="bottom", fill="x")
        ctk.CTkFrame(footer, height=1, fg_color=BORDER).pack(fill="x")
        ctk.CTkLabel(footer, text="v1.0 · 2026", font=(FONT, SZ_TINY),
                     text_color=SUB).pack(anchor="w", padx=16, pady=14)

        self.main = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self.main.pack(side="right", fill="both", expand=True)

    def _sbtn(self, parent, text, cmd):
        b = ctk.CTkButton(parent, text=text, command=cmd,
                          font=(FONT, SZ_SIDEBAR), text_color=SUB,
                          fg_color="transparent", hover_color=SURFACE,
                          anchor="w", height=32, corner_radius=6)
        b.pack(fill="x", padx=8, pady=1)
        return b

    def _sel(self, btn):
        for b in [self.b1, self.b2, self.b5, self.b3, self.b4]:
            b.configure(fg_color="transparent", text_color=SUB)
        btn.configure(fg_color="#262626", text_color=TEXT)

    def _clr(self):
        for w in self.main.winfo_children():
            w.destroy()

    def _scroll(self):
        # Property Tax AI: scrollbar 6px, #30363D, 오른쪽 끝에 붙임
        sc = ctk.CTkScrollableFrame(self.main, fg_color=BG,
                                    scrollbar_button_color=SB_COLOR,
                                    scrollbar_button_hover_color=SB_HOVER)
        sc.pack(fill="both", expand=True, padx=(32, 0), pady=0)
        # 스크롤바 폭 줄이기
        try:
            sc._scrollbar.configure(width=12)
        except Exception:
            pass
        return sc

    def _page_title(self, parent, text):
        # page-title: 15px, letter-spacing 2px, border-bottom 1px, margin-bottom 24px
        ctk.CTkLabel(parent, text=text, font=(FONT, SZ_TITLE),
                     text_color=ACCENT).pack(anchor="w", pady=(0, 14))
        ctk.CTkFrame(parent, height=1, fg_color=BORDER).pack(fill="x", pady=(0, 24))

    def _card(self, parent, accent_border=False):
        # card: border-radius 14px, padding 28px, border 1px
        bd = "#3a5f80" if accent_border else BORDER
        c = ctk.CTkFrame(parent, fg_color=SURFACE, corner_radius=14,
                         border_width=1, border_color=bd)
        c.pack(fill="x", pady=(0, 16), padx=(0, 32))
        inner = ctk.CTkFrame(c, fg_color="transparent")
        inner.pack(fill="x", padx=28, pady=24)
        return inner

    def _section_label(self, parent, text):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(anchor="w", pady=(0, 12))
        ctk.CTkFrame(f, fg_color=ACCENT, width=3, height=18, corner_radius=1).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(f, text=text, font=(FONT, SZ_BODY, "bold"), text_color=ACCENT).pack(side="left")

    def _row(self, parent, label, value, color=TEXT, bold=False):
        # rb: flex space-between, padding 6px 0, border-bottom 1px
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x")
        ctk.CTkFrame(f, height=1, fg_color=BORDER).pack(fill="x", side="bottom")
        inner = ctk.CTkFrame(f, fg_color="transparent")
        inner.pack(fill="x", padx=0, pady=6)
        ctk.CTkLabel(inner, text=label, font=(FONT, SZ_BODY), text_color=SUB).pack(side="left")
        ft = (FONT, SZ_BODY, "bold") if bold else (FONT, SZ_BODY)
        ctk.CTkLabel(inner, text=value, font=ft, text_color=color).pack(side="right")

    # ════════════════════════════════════════
    #  홈
    # ════════════════════════════════════════
    def _page_home(self):
        self._clr()
        self._sel(self.b1)
        sc = self._scroll()

        self._page_title(sc, "법인세 신고서 검토")

        # 드롭존 카드 (크게)
        dz_card = ctk.CTkFrame(sc, fg_color=SURFACE, corner_radius=14,
                               border_width=1, border_color=BORDER)
        dz_card.pack(fill="x", pady=(0, 16), padx=(0, 32))

        dz_inner = ctk.CTkFrame(dz_card, fg_color="transparent")
        dz_inner.pack(fill="x", padx=28, pady=20)

        self._section_label(dz_inner, "검토 대상 파일")

        # 드롭존 영역
        self.dropzone = ctk.CTkFrame(dz_inner, fg_color=INPUT_BG, corner_radius=10,
                                     height=100, border_width=1, border_color=BORDER)
        self.dropzone.pack(fill="x", pady=(0, 10))
        self.dropzone.pack_propagate(False)

        dz_center = ctk.CTkFrame(self.dropzone, fg_color="transparent")
        dz_center.place(relx=0.5, rely=0.5, anchor="center")

        self.fl = ctk.CTkLabel(dz_center, text="PDF 파일을 여기에 드래그하세요",
                               font=(FONT, SZ_BODY), text_color=SUB)
        self.fl.pack()

        # 파일 선택 버튼 (드롭존 아래)
        ctk.CTkButton(dz_inner, text="파일 선택", command=self._pick,
                      font=(FONT, SZ_BODY, "bold"), fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      text_color="#fff", width=150, height=40,
                      corner_radius=6).pack(anchor="w")


        # 검토 실행 카드
        card2 = self._card(sc)
        self._section_label(card2, "검토 실행")

        bf = ctk.CTkFrame(card2, fg_color="transparent")
        bf.pack(fill="x")

        self.rbtn = ctk.CTkButton(bf, text="검토 시작", command=self._run,
                                  font=(FONT, SZ_BODY, "bold"), fg_color=ACCENT,
                                  hover_color=ACCENT_HOVER, text_color="#fff",
                                  height=40, corner_radius=6, width=150)
        self.rbtn.pack(side="left")

        self.plbl = ctk.CTkLabel(bf, text="", font=(FONT, SZ_SMALL), text_color=SUB)
        self.plbl.pack(side="left", padx=16)

        self.pbar = ctk.CTkProgressBar(card2, fg_color=INPUT_BG,
                                       progress_color=ACCENT, height=3)
        self.pbar.pack(fill="x", pady=(12, 0))
        self.pbar.set(0)

        self.summary_parent = sc

    def _detect(self, parent):
        dirs = []
        d1 = os.path.expanduser(r"~\Desktop\Corp Tax_review")
        if os.path.isdir(d1):
            dirs.append(d1)
        d2 = os.path.expanduser(r"~\Desktop\Claude Code\법인세 검토")
        if os.path.isdir(d2):
            dirs.append(d2)
        pdfs = []
        for d in dirs:
            for f in os.listdir(d):
                if f.lower().endswith(".pdf"):
                    pdfs.append(os.path.join(d, f))
        if pdfs:
            self._detected_pdfs = {os.path.basename(p): p for p in pdfs}
            names = list(self._detected_pdfs.keys())
            self.pdf_dropdown = ctk.CTkOptionMenu(
                parent, values=names, command=self._on_dropdown_select,
                font=(FONT, SZ_SMALL), dropdown_font=(FONT, SZ_SMALL),
                fg_color=INPUT_BG, button_color="#2a2a2a",
                button_hover_color="#333", dropdown_fg_color=SURFACE,
                dropdown_hover_color="#2a2a2a", text_color=SUB,
                height=32, corner_radius=6)
            self.pdf_dropdown.set("탐색된 PDF 선택")
            self.pdf_dropdown.pack(fill="x", pady=(12, 0))

    def _on_dropdown_select(self, choice):
        path = self._detected_pdfs.get(choice)
        if path:
            self._setpdf(path)

    def _pick(self):
        p = filedialog.askopenfilename(
            title="세무조정계산서 PDF 선택",
            filetypes=[("PDF", "*.pdf")],
            initialdir=os.path.expanduser(r"~\Desktop\Corp Tax_review"))
        if p:
            self._setpdf(p)

    def _setpdf(self, p):
        self.pdf_path = p
        if hasattr(self, 'fl'):
            self.fl.configure(text=os.path.basename(p), text_color=TEXT)
        if hasattr(self, 'dropzone'):
            # 드롭 시 시각 피드백: 잠깐 밝은 테두리 → ACCENT로 전환
            self.dropzone.configure(border_color="#8BC4F0", border_width=2)
            self.after(400, lambda: self.dropzone.configure(
                border_color=ACCENT, border_width=1))

    # ════════════════════════════════════════
    #  검토 실행
    # ════════════════════════════════════════
    def _run(self):
        if not self.pdf_path:
            self.plbl.configure(text="PDF를 먼저 선택하세요", text_color=ERR)
            return
        if not os.path.exists(self.pdf_path):
            self.plbl.configure(text="파일을 찾을 수 없습니다", text_color=ERR)
            return
        self.rbtn.configure(state="disabled", text="검토 중...")
        self.plbl.configure(text="PDF 추출 중...", text_color=SUB)
        self.pbar.set(0.1)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            self.after(0, lambda: (self.pbar.set(0.2),
                                   self.plbl.configure(text="[1/3] PDF 파싱...")))
            data = parse_all(self.pdf_path)
            self.after(0, lambda: (self.pbar.set(0.5),
                                   self.plbl.configure(text="[2/3] 세무 검토...")))
            results = run_all_reviews(data)
            self.after(0, lambda: (self.pbar.set(0.8),
                                   self.plbl.configure(text="[3/3] Excel 생성...")))
            nm = _short_corp_name(data["회사정보"].get("법인명", "검토대상"))
            dt = datetime.now().strftime("%Y%m%d")
            xf = f"법인세검토결과_{nm}_{dt}.xlsx"
            xp = os.path.join(os.path.dirname(self.pdf_path) or ".", xf)
            try:
                generate_excel_report(data, results, xp)
            except Exception:
                xp = xf
                try:
                    generate_excel_report(data, results, xp)
                except Exception:
                    xp = None
            self.data, self.results, self.excel_path = data, results, xp
            self.after(0, self._done)
        except Exception as e:
            self.after(0, lambda: self._err(str(e)))

    def _done(self):
        self.pbar.set(1.0)
        ni = sum(1 for r in self.results if r.상태 == "이슈")
        no = sum(1 for r in self.results if r.상태 == "정상")
        nc = sum(1 for r in self.results if r.상태 == "확인필요")
        self.plbl.configure(
            text=f"완료 · {len(self.results)}건 (정상 {no} / 확인 {nc} / 이슈 {ni})",
            text_color=OK if ni == 0 else ERR)
        self.rbtn.configure(state="normal", text="검토 시작")
        self._show_summary()

    def _err(self, msg):
        self.pbar.set(0)
        self.plbl.configure(text=f"오류: {msg}", text_color=ERR)
        self.rbtn.configure(state="normal", text="검토 시작")

    def _show_summary(self):
        sc = self.summary_parent
        card = self._card(sc, accent_border=True)
        self._section_label(card, "검토 결과 요약")

        info = self.data.get("회사정보", {})
        ctk.CTkLabel(card,
                     text=f"{_short_corp_name(info.get('법인명',''))}  ·  {info.get('사업자등록번호','')}  ·  {info.get('사업연도_시작','')}~{info.get('사업연도_종료','')}",
                     font=(FONT, SZ_BODY), text_color=TEXT).pack(anchor="w", pady=(0, 12))

        ni = sum(1 for r in self.results if r.상태 == "이슈")
        no = sum(1 for r in self.results if r.상태 == "정상")
        nc = sum(1 for r in self.results if r.상태 == "확인필요")

        sf = ctk.CTkFrame(card, fg_color="transparent")
        sf.pack(fill="x", pady=(0, 12))
        for lb, ct, co in [("정상", no, OK), ("확인필요", nc, WARN), ("이슈", ni, ERR)]:
            bx = ctk.CTkFrame(sf, fg_color=INPUT_BG, corner_radius=8, height=58)
            bx.pack(side="left", padx=(0, 8), fill="x", expand=True)
            bx.pack_propagate(False)
            ctk.CTkLabel(bx, text=str(ct), font=(FONT, SZ_BIG, "bold"), text_color=co).pack(pady=(6, 0))
            ctk.CTkLabel(bx, text=lb, font=(FONT, SZ_TINY), text_color=SUB).pack()

        bf = ctk.CTkFrame(card, fg_color="transparent")
        bf.pack(fill="x")
        ctk.CTkButton(bf, text="상세 보기", command=self._page_results,
                      font=(FONT, SZ_SMALL), fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      text_color="#fff", height=32, corner_radius=6,
                      width=110).pack(side="left")
        if self.excel_path:
            ctk.CTkButton(bf, text="Excel 열기",
                          command=lambda: os.startfile(self.excel_path) if self.excel_path and os.path.exists(self.excel_path) else None,
                          font=(FONT, SZ_SMALL), fg_color="#2a2a2a", hover_color="#333",
                          text_color=TEXT, height=32, corner_radius=6,
                          width=110).pack(side="left", padx=(8, 0))

    # ════════════════════════════════════════
    #  결과 화면
    # ════════════════════════════════════════
    def _page_results(self):
        self._clr()
        self._sel(self.b2)
        if not self.results:
            ctk.CTkLabel(self.main, text="먼저 검토를 실행하세요",
                         font=(FONT, SZ_BODY), text_color=SUB).pack(pady=40)
            return
        sc = self._scroll()
        info = self.data.get("회사정보", {})

        self._page_title(sc, "검토 결과")

        # 상단 회사명 + 엑셀 다운로드 버튼
        top_bar = ctk.CTkFrame(sc, fg_color="transparent")
        top_bar.pack(fill="x", padx=(0, 32), pady=(0, 16))

        ctk.CTkLabel(top_bar, text=f"{_short_corp_name(info.get('법인명',''))} · {info.get('사업연도_시작','')}~{info.get('사업연도_종료','')}",
                     font=(FONT, SZ_SMALL), text_color=SUB).pack(side="left")

        ctk.CTkButton(top_bar, text="📥 Excel 다운로드", width=140, height=32,
                      font=(FONT, SZ_SMALL), fg_color="#2B6CB0", hover_color="#2C5282",
                      command=self._export_excel).pack(side="right", padx=(4, 0))

        if self.excel_path and os.path.exists(self.excel_path):
            ctk.CTkButton(top_bar, text="📄 Excel 미리보기", width=140, height=32,
                          font=(FONT, SZ_SMALL), fg_color="#38A169", hover_color="#2F855A",
                          command=lambda: os.startfile(self.excel_path)).pack(side="right", padx=(4, 0))

        tax = self.data.get("세액조정", {})

        # 주요 검토 항목 카테고리 (서식 간 크로스체크)
        주요카테고리_순서 = ["세액감면", "공제감면", "세액공제",
                        "이월결손금", "농어촌특별세"]
        주요카테고리 = set(주요카테고리_순서)
        주요_raw = [r for r in self.results if r.카테고리 in 주요카테고리]
        # 카테고리 순서대로 정렬
        카테고리_순번 = {c: i for i, c in enumerate(주요카테고리_순서)}
        주요 = sorted(주요_raw, key=lambda r: 카테고리_순번.get(r.카테고리, 99))
        기타 = [r for r in self.results if r.카테고리 not in 주요카테고리]

        # 2열 레이아웃: 왼쪽=주요, 오른쪽=기타
        cols = ctk.CTkFrame(sc, fg_color="transparent")
        cols.pack(fill="both", expand=True, padx=(0, 32))
        cols.columnconfigure(0, weight=1, uniform="col")
        cols.columnconfigure(1, weight=1, uniform="col")

        for col_idx, (section_name, section_results) in enumerate([("주요 검토 항목", 주요), ("기타 검토 항목", 기타)]):
            col_frame = ctk.CTkFrame(cols, fg_color="transparent")
            col_frame.grid(row=0, column=col_idx, sticky="nsew", padx=(0, 8 if col_idx == 0 else 0))

            if not section_results:
                continue
            ctk.CTkLabel(col_frame, text=f"  {section_name}",
                         font=(FONT, SZ_BODY, "bold"), text_color=ACCENT
                         ).pack(anchor="w", pady=(16, 4))
            for lbl, sts_list, co, ic in [("확인필요", ["이슈", "확인필요"], ERR, "X"),
                                          ("정상", ["정상"], OK, "O")]:
                grp = [r for r in section_results if r.상태 in sts_list]
                if not grp:
                    continue
                ctk.CTkLabel(col_frame, text=f" {lbl} ({len(grp)}건)",
                             font=(FONT, SZ_BODY, "bold"), text_color=co
                             ).pack(anchor="w", pady=(12, 6))
                for item in grp:
                    self._rcard(col_frame, item, co, ic)

    def _wrap_label(self, parent, text, font_tuple, fg, bg=SURFACE, padx=0, pady=(0,0)):
        """줄바꿈 시 줄간격 넓은 라벨 (tk.Text 기반)"""
        t = tk.Text(parent, wrap="word", borderwidth=0, highlightthickness=0,
                    bg=bg, fg=fg, font=font_tuple, spacing2=6,
                    cursor="arrow", padx=padx, pady=0)
        t.insert("1.0", text)
        t.configure(state="disabled")
        # 높이 자동 계산
        def _fit(event=None):
            t.configure(width=event.width if event else t.winfo_width())
            t.update_idletasks()
            lines = int(t.index("end-1c").split(".")[0])
            t.configure(height=lines)
        t.pack(fill="x", pady=pady)
        t.bind("<Configure>", _fit)
        return t

    def _rcard(self, parent, item, color, icon):
        c = ctk.CTkFrame(parent, fg_color=SURFACE, corner_radius=10,
                         border_width=1, border_color=BORDER)
        c.pack(fill="x", pady=4, padx=(0, 8))
        inn = ctk.CTkFrame(c, fg_color="transparent")
        inn.pack(fill="x", padx=16, pady=16)

        # 제목: [icon] + 항목명
        hd = ctk.CTkFrame(inn, fg_color="transparent")
        hd.pack(fill="x")
        ctk.CTkLabel(hd, text=f"[{icon}]", font=(FONT, SZ_SMALL, "bold"),
                     text_color=color, width=26).pack(side="left", anchor="n", pady=(2,0))
        ctk.CTkLabel(hd, text=item.항목명, font=(FONT, SZ_BODY, "bold"),
                     text_color=TEXT, wraplength=380, justify="left", anchor="w"
                     ).pack(side="left", padx=(6, 0), fill="x", expand=True)

        if item.비고:
            forms = list(dict.fromkeys(re.findall(r'\[([^\]]+)\]', item.비고)))
            if forms:
                ctk.CTkLabel(inn, text="검토서식:",
                             font=(FONT, SZ_TINY), text_color=SUB
                             ).pack(anchor="w", padx=32, pady=(2, 0))
                for f in forms:
                    ctk.CTkLabel(inn, text=f"  · {f}",
                                 font=(FONT, SZ_TINY), text_color=SUB
                                 ).pack(anchor="w", padx=40, pady=0)

        # 멀티라인 비고 (줄바꿈 포함): 각 줄을 그대로 렌더링
        if item.비고 and "\n" in item.비고:
            bigo_lines = [l.strip() for l in item.비고.split("\n") if l.strip()]

            # 검토서식과 비고 사이 여백
            ctk.CTkFrame(inn, fg_color="transparent", height=6).pack()

            for i, line in enumerate(bigo_lines):
                if line.startswith("✓") or line.startswith("✗"):
                    color = OK if line.startswith("✓") else ERR
                    ctk.CTkLabel(inn, text=f"  {line}", font=(FONT, SZ_SMALL, "bold"),
                                 text_color=color, anchor="w"
                                 ).pack(anchor="w", padx=32, pady=(4, 0))
                else:
                    ctk.CTkLabel(inn, text=f"  {line}",
                                 font=(FONT, SZ_TINY), text_color=TEXT,
                                 anchor="w").pack(anchor="w", padx=32, pady=(4 if i == 0 else 1, 0))
        elif item.신고서금액 is not None and item.검증금액 is not None:
                # 비고에서 "서식명: 금액 / 서식명: 금액" 패턴 파싱 (3개 이상 서식 지원)
                detail_parts = [p.strip() for p in item.비고.split(" / ")] if item.비고 and " / " in item.비고 else []
                parsed_details = []
                for p in detail_parts:
                    m = re.match(r'^(.+?):\s*([\d,]+)', p)
                    if m:
                        parsed_details.append((m.group(1).strip(), m.group(2)))
                if len(parsed_details) >= 3:
                    for i, (name, val) in enumerate(parsed_details):
                        ctk.CTkLabel(inn, text=f"  {i+1}. {name}: {val}원",
                                     font=(FONT, SZ_TINY), text_color=TEXT,
                                     anchor="w").pack(anchor="w", padx=32, pady=(4 if i == 0 else 1, 0))
                else:
                    # 2개 서식: 기존 방식
                    forms_in_bigo = re.findall(r'\[([^\]]+)\]', item.비고) if item.비고 else []
                    unique_forms = list(dict.fromkeys(forms_in_bigo))
                    if len(unique_forms) >= 2:
                        lbl1, lbl2 = unique_forms[0], unique_forms[1]
                    else:
                        lbl1, lbl2 = "신고서", "검증"

                    for i, (lb, vl) in enumerate([(lbl1, item.신고서금액), (lbl2, item.검증금액)]):
                        ctk.CTkLabel(inn, text=f"  {i+1}. {lb}: {format_won(vl)}원",
                                     font=(FONT, SZ_TINY), text_color=TEXT,
                                     anchor="w").pack(anchor="w", padx=32, pady=(4 if i == 0 else 1, 0))

                # 일치/불일치 표시 (1원이라도 차이나면 불일치)
                if item.차이금액 is not None and item.차이금액 == 0:
                    ctk.CTkLabel(inn, text="✓ 일치", font=(FONT, SZ_SMALL, "bold"),
                                 text_color=OK).pack(anchor="e", padx=32, pady=(2, 0))
                elif item.차이금액 and item.차이금액 > 0:
                    rf = ctk.CTkFrame(inn, fg_color="transparent")
                    rf.pack(fill="x", padx=32, pady=(2, 0))
                    ctk.CTkLabel(rf, text="✗ 불일치", font=(FONT, SZ_SMALL, "bold"), text_color=ERR).pack(side="left")
                    ctk.CTkLabel(rf, text=f"차이 {format_won(item.차이금액)}원",
                                 font=(FONT, SZ_SMALL, "bold"), text_color=ERR).pack(side="right")
        elif item.신고서금액 is not None:
            ctk.CTkLabel(inn, text=f"금액: {format_won(item.신고서금액)}원",
                         font=(FONT, SZ_SMALL), text_color=TEXT).pack(anchor="w", padx=32, pady=(4, 0))
            if item.비고:
                desc = re.sub(r'\[[^\]]+\]\s*', '', item.비고).strip()
                if desc and len(desc) > 5:
                    ctk.CTkLabel(inn, text=desc, font=(FONT, SZ_TINY), text_color=SUB,
                                 anchor="w").pack(anchor="w", padx=32, pady=(4, 0))

    def _export_excel(self):
        if not self.results:
            return
        from tkinter import filedialog
        info = self.data.get("회사정보", {})
        nm = _short_corp_name(info.get("법인명", "검토대상"))
        dt = datetime.now().strftime("%Y%m%d")
        default_name = f"법인세검토결과_{nm}_{dt}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_name,
            title="검토결과 저장"
        )
        if not path:
            return
        try:
            generate_excel_report(self.data, self.results, path)
            os.startfile(path)
        except Exception as e:
            ctk.CTkLabel(self.main, text=f"저장 실패: {e}",
                         font=(FONT, SZ_SMALL), text_color=ERR).pack(pady=8)

    # ════════════════════════════════════════
    #  보고서
    # ════════════════════════════════════════
    def _page_report(self):
        self._clr()
        self._sel(self.b5)
        if not self.data:
            ctk.CTkLabel(self.main, text="먼저 검토를 실행하세요",
                         font=(FONT, SZ_BODY), text_color=SUB).pack(pady=40)
            return
        sc = self._scroll()
        PAD = (0, 32)  # 오른쪽 여백 통일

        info = self.data.get("회사정보", {})
        tax = self.data.get("세액조정", {})
        fs = self.data.get("재무제표", {})
        is_data = fs.get("손익계산서", {})
        is_prev = fs.get("손익계산서_전기", {})
        if not is_prev:
            is_prev = self.data.get("손익계산서_전기", {})

        corp_name = _short_corp_name(info.get("법인명", ""))
        biz_no = info.get("사업자등록번호", "")
        period_start = info.get("사업연도_시작", "")
        period_end = info.get("사업연도_종료", "")
        cur_year = period_end[:4] if period_end else "당기"
        prev_year = str(int(cur_year) - 1) if cur_year.isdigit() else "전기"
        sales_cur = is_data.get("매출액")
        sales_prev = is_prev.get("매출액")

        # 공통 색상
        TBL_BORDER = "#333333"
        TBL_HDR = "#1a2332"
        TBL_ROW1 = "#1e1e1e"
        TBL_ROW2 = SURFACE
        TBL_ACCENT_ROW = "#1c2530"
        TBL_HDR_TEXT = "#b0c0d0"
        TBL_ACCENT_TEXT = "#8BB8E0"

        # ── 보고서 헤더 ──
        hdr_frame = ctk.CTkFrame(sc, fg_color="transparent")
        hdr_frame.pack(fill="x", padx=PAD, pady=(0, 6))
        ctk.CTkLabel(hdr_frame, text=f"{corp_name} 법인세 신고 사항",
                     font=(FONT, 18, "bold"), text_color=TBL_ACCENT_TEXT).pack(anchor="w")
        sub_text = corp_name
        if biz_no:
            sub_text += f"  |  {biz_no}"
        sub_text += f"  |  {period_start} ~ {period_end}"
        ctk.CTkLabel(hdr_frame, text=sub_text,
                     font=(FONT, SZ_SMALL), text_color=SUB).pack(anchor="w", pady=(4, 0))
        ctk.CTkFrame(sc, height=2, fg_color=TBL_HDR).pack(fill="x", padx=PAD, pady=(6, 16))

        # 내보내기 버튼
        btn_frame = ctk.CTkFrame(sc, fg_color="transparent")
        btn_frame.pack(fill="x", padx=PAD, pady=(0, 16))
        ctk.CTkButton(btn_frame, text="PDF 미리보기", command=self._preview_pdf,
                      font=(FONT, SZ_SMALL, "bold"), fg_color=ACCENT,
                      hover_color=ACCENT_HOVER, text_color="#fff",
                      height=32, corner_radius=6, width=120).pack(side="left")
        ctk.CTkButton(btn_frame, text="PDF 저장", command=self._save_pdf,
                      font=(FONT, SZ_SMALL), fg_color="#2a2a2a",
                      hover_color="#333", text_color=TEXT,
                      height=32, corner_radius=6, width=90).pack(side="left", padx=(8, 0))
        ctk.CTkButton(btn_frame, text="Word 저장", command=self._save_word,
                      font=(FONT, SZ_SMALL), fg_color="#2a2a2a",
                      hover_color="#333", text_color=TEXT,
                      height=32, corner_radius=6, width=90).pack(side="left", padx=(8, 0))
        self.pdf_status = ctk.CTkLabel(btn_frame, text="", font=(FONT, SZ_SMALL), text_color=SUB)
        self.pdf_status.pack(side="left", padx=12)

        # ══════════════════════════════════════
        #  01. 손익계산서 비교
        # ══════════════════════════════════════
        self._report_section_title(sc, "01", "손익계산서 주요 항목 비교", PAD)

        tbl = ctk.CTkFrame(sc, fg_color=TBL_BORDER, corner_radius=0)
        tbl.pack(fill="x", padx=PAD, pady=(0, 24))
        for c, w in enumerate([14, 12, 6, 12, 6, 12, 6]):
            tbl.columnconfigure(c, weight=w)

        # 그룹 헤더
        self._tcell(tbl, 0, 0, "", TBL_HDR, TBL_HDR_TEXT, bold=True)
        self._tcell(tbl, 0, 1, f"{cur_year}년 (당기)", TBL_HDR, TBL_ACCENT_TEXT, bold=True, colspan=2)
        self._tcell(tbl, 0, 3, f"{prev_year}년 (전기)", TBL_HDR, "#8BC4A0", bold=True, colspan=2)
        self._tcell(tbl, 0, 5, "증감", TBL_HDR, "#D4A27A", bold=True, colspan=2)

        # 컬럼 헤더
        for c, txt in enumerate(["계정과목", "금액", "비율", "금액", "비율", "증감액", "증감률"]):
            self._tcell(tbl, 1, c, txt, TBL_HDR, "#778899", bold=True, size=SZ_TINY)

        items = [
            ("매출액", True), ("매출원가", False), ("매출총이익", True),
            ("판관비", False), ("영업이익", True),
            ("영업외수익", False), ("영업외비용", False),
            ("법인세차감전이익", True), ("법인세등", False), ("당기순이익", True),
        ]

        for idx, (key, bold) in enumerate(items):
            r = idx + 2
            cur = is_data.get(key)
            prev = is_prev.get(key)
            bg = TBL_ACCENT_ROW if bold else (TBL_ROW1 if idx % 2 == 0 else TBL_ROW2)
            name_co = TBL_ACCENT_TEXT if bold else TEXT
            self._tcell(tbl, r, 0, key, bg, name_co, bold=bold, anchor="w")
            self._tcell(tbl, r, 1, self._fmt_won(cur), bg, TEXT, bold=bold, anchor="e")
            self._tcell(tbl, r, 2, self._pct(cur, sales_cur), bg, SUB, anchor="e", size=SZ_TINY)
            self._tcell(tbl, r, 3, self._fmt_won(prev), bg, TEXT, bold=bold, anchor="e")
            self._tcell(tbl, r, 4, self._pct(prev, sales_prev), bg, SUB, anchor="e", size=SZ_TINY)
            dv = self._delta(cur, prev)
            dc = OK if "▲" in str(dv) else (ERR if "▽" in str(dv) else TEXT)
            self._tcell(tbl, r, 5, dv, bg, dc, bold=bold, anchor="e")
            dp = self._delta_pct(cur, prev)
            dpc = OK if str(dp).startswith("+") else (ERR if str(dp).startswith("-") and dp != "-" else SUB)
            self._tcell(tbl, r, 6, dp, bg, dpc, anchor="e", size=SZ_TINY)

        # ══════════════════════════════════════
        #  02. 세무조정 내역
        # ══════════════════════════════════════
        self._report_section_title(sc, "02", "주요 세무조정 내역", PAD)

        adj = self.data.get("소득금액조정", {})
        tbl2 = ctk.CTkFrame(sc, fg_color=TBL_BORDER, corner_radius=0)
        tbl2.pack(fill="x", padx=PAD, pady=(0, 24))
        for c, w in enumerate([2, 5, 3, 2]):
            tbl2.columnconfigure(c, weight=w)

        for c, (txt, anc) in enumerate([("구분", "center"), ("조정항목", "w"), ("금액", "e"), ("소득처분", "center")]):
            self._tcell(tbl2, 0, c, txt, TBL_HDR, TBL_HDR_TEXT, bold=True, anchor=anc)

        r = 1
        for prefix, items_key in [("익금산입", "익금산입_항목"), ("손금산입", "손금산입_항목")]:
            for item in adj.get(items_key, []):
                bg = TBL_ROW1 if r % 2 == 0 else TBL_ROW2
                항목 = item.get('항목명', '') or item.get('과목', '')
                self._tcell(tbl2, r, 0, prefix, bg, ACCENT, anchor="center")
                self._tcell(tbl2, r, 1, 항목, bg, TEXT, anchor="w")
                self._tcell(tbl2, r, 2, self._fmt_won(item.get('금액')), bg, TEXT, anchor="e")
                self._tcell(tbl2, r, 3, item.get('처분', ''), bg, SUB, anchor="center")
                r += 1

        for lbl, key in [("익금산입 합계", "익금산입_합계"), ("손금산입 합계", "손금산입_합계")]:
            val = adj.get(key)
            if val:
                self._tcell(tbl2, r, 0, "", TBL_ACCENT_ROW, TEXT)
                self._tcell(tbl2, r, 1, lbl, TBL_ACCENT_ROW, TBL_ACCENT_TEXT, bold=True, anchor="w")
                self._tcell(tbl2, r, 2, self._fmt_won(val), TBL_ACCENT_ROW, TBL_ACCENT_TEXT, bold=True, anchor="e")
                self._tcell(tbl2, r, 3, "", TBL_ACCENT_ROW, TEXT)
                r += 1

        # ══════════════════════════════════════
        #  03. 법인세 산출 내역
        # ══════════════════════════════════════
        self._report_section_title(sc, "03", "법인세 산출 내역", PAD)

        납부 = tax.get("차감납부할세액")
        지방세 = int(납부 * 0.1) if 납부 else None
        합계납부 = (납부 + 지방세) if (납부 and 지방세) else None

        tax_items = [
            ("결산서상 당기순손익", tax.get("결산서상당기순손익"), False),
            ("(+) 익금산입", tax.get("익금산입"), False),
            ("(-) 손금산입", tax.get("손금산입"), False),
            ("각사업연도 소득금액", tax.get("각사업연도소득금액"), True),
            ("(-) 이월결손금 공제", tax.get("이월결손금공제"), False),
            ("과세표준", tax.get("과세표준"), True),
            ("산출세액", tax.get("산출세액"), False),
            ("(-) 공제감면세액", tax.get("최저한세적용대상_공제감면세액"), False),
            ("법인세 차감납부세액", 납부, True),
            ("법인지방소득세 (10%)", 지방세, False),
            ("합계 납부세액", 합계납부, True),
        ]

        tbl3 = ctk.CTkFrame(sc, fg_color=TBL_BORDER, corner_radius=0)
        tbl3.pack(fill="x", padx=PAD, pady=(0, 24))
        tbl3.columnconfigure(0, weight=3)
        tbl3.columnconfigure(1, weight=2)

        self._tcell(tbl3, 0, 0, "구분", TBL_HDR, TBL_HDR_TEXT, bold=True, anchor="center")
        self._tcell(tbl3, 0, 1, "금액 (원)", TBL_HDR, TBL_HDR_TEXT, bold=True, anchor="center")

        r = 1
        for label, val, bold in tax_items:
            if val is None and not bold:
                continue
            bg = TBL_ACCENT_ROW if bold else (TBL_ROW1 if r % 2 == 0 else TBL_ROW2)
            co = TBL_ACCENT_TEXT if bold else TEXT
            self._tcell(tbl3, r, 0, label, bg, co, bold=bold, anchor="w")
            self._tcell(tbl3, r, 1, self._fmt_won(val), bg, co, bold=bold, anchor="e")
            r += 1


    def _report_section_title(self, parent, num, title, padx):
        """보고서 섹션 타이틀 (번호 + 제목)"""
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.pack(fill="x", padx=padx, pady=(0, 8))
        # 번호 뱃지
        badge = ctk.CTkFrame(f, fg_color="#1a2332", width=28, height=22, corner_radius=4)
        badge.pack(side="left", padx=(0, 10))
        badge.pack_propagate(False)
        ctk.CTkLabel(badge, text=num, font=(FONT, SZ_TINY, "bold"),
                     text_color="#8BB8E0").place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(f, text=title, font=(FONT, SZ_BODY, "bold"),
                     text_color=TEXT).pack(side="left")

    def _tcell(self, parent, row, col, text, bg, fg, bold=False, anchor="center",
               size=None, colspan=1):
        """표 셀 생성 (grid 배치)"""
        sz = size or SZ_SMALL
        ft = (FONT, sz, "bold") if bold else (FONT, sz)
        cell = ctk.CTkFrame(parent, fg_color=bg, corner_radius=0)
        cell.grid(row=row, column=col, columnspan=colspan, sticky="nsew",
                  padx=(0, 1), pady=(0, 1))
        ctk.CTkLabel(cell, text=text, font=ft, text_color=fg,
                     anchor=anchor).pack(fill="x", padx=6, pady=4)

    def _row_bg(self, parent, label, value, color=TEXT, bold=False, bg_color="transparent"):
        """배경색 있는 행"""
        f = ctk.CTkFrame(parent, fg_color=bg_color, corner_radius=3)
        f.pack(fill="x")
        ctk.CTkFrame(f, height=1, fg_color="#2a2a2a").pack(fill="x", side="bottom")
        inner = ctk.CTkFrame(f, fg_color="transparent")
        inner.pack(fill="x", padx=0, pady=5)
        ft = (FONT, SZ_BODY, "bold") if bold else (FONT, SZ_BODY)
        ctk.CTkLabel(inner, text=label, font=ft, text_color=color).pack(side="left")
        ctk.CTkLabel(inner, text=value, font=ft, text_color=color).pack(side="right")

    def _fmt_won(self, val):
        if val is None:
            return "-"
        if isinstance(val, int):
            if val < 0:
                return f"△{abs(val):,}"
            return f"{val:,}"
        return str(val)

    def _pct(self, val, total):
        if not total or not val:
            return "-"
        return f"{val / total * 100:.1f}%"

    def _delta(self, cur, prev):
        if cur is None or prev is None:
            return "-"
        d = cur - prev
        if d > 0:
            return f"▲ {d:,}"
        elif d < 0:
            return f"▽ {abs(d):,}"
        return "0"

    def _delta_pct(self, cur, prev):
        if cur is None or prev is None or prev == 0:
            return "-"
        d = (cur - prev) / abs(prev) * 100
        return f"{d:+.1f}%"

    def _preview_pdf(self):
        """임시파일로 PDF 생성 후 바로 열기"""
        if not self.data:
            return
        import tempfile
        info = self.data.get("회사정보", {})
        nm = _short_corp_name(info.get("법인명", "검토대상"))
        dt = datetime.now().strftime("%Y%m%d")
        fname = f"법인세검토보고서_{nm}_{dt}.pdf"
        path = os.path.join(tempfile.gettempdir(), fname)
        try:
            generate_report_pdf(self.data, self.results, path)
            self.pdf_status.configure(text="PDF 열림", text_color=OK)
            os.startfile(path)
        except Exception as e:
            self.pdf_status.configure(text=f"오류: {e}", text_color=ERR)

    def _save_pdf(self):
        """저장 위치 선택 후 PDF 생성"""
        if not self.data:
            return
        info = self.data.get("회사정보", {})
        nm = info.get("법인명", "검토대상")
        dt = datetime.now().strftime("%Y%m%d")
        default_name = f"법인세검토보고서_{nm}_{dt}.pdf"
        path = filedialog.asksaveasfilename(
            title="보고서 PDF 저장",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=default_name,
            initialdir=os.path.expanduser(r"~\Desktop"))
        if not path:
            return
        try:
            generate_report_pdf(self.data, self.results, path)
            self.pdf_status.configure(text=f"PDF 저장 완료", text_color=OK)
        except Exception as e:
            self.pdf_status.configure(text=f"오류: {e}", text_color=ERR)

    def _save_word(self):
        """Word(docx) 저장"""
        if not self.data:
            return
        info = self.data.get("회사정보", {})
        nm = info.get("법인명", "검토대상")
        dt = datetime.now().strftime("%Y%m%d")
        default_name = f"법인세검토보고서_{nm}_{dt}.docx"
        path = filedialog.asksaveasfilename(
            title="보고서 Word 저장",
            defaultextension=".docx",
            filetypes=[("Word", "*.docx")],
            initialfile=default_name,
            initialdir=os.path.expanduser(r"~\Desktop"))
        if not path:
            return
        try:
            generate_report_docx(self.data, self.results, path)
            self.pdf_status.configure(text=f"Word 저장 완료", text_color=OK)
            os.startfile(path)
        except Exception as e:
            self.pdf_status.configure(text=f"오류: {e}", text_color=ERR)

    # ════════════════════════════════════════
    #  검토 목록
    # ════════════════════════════════════════
    def _page_checklist(self):
        self._clr()
        self._sel(self.b3)
        sc = self._scroll()
        self._page_title(sc, "검토 목록")

        CHECKS = [
            ("주요 검토목록 (서식 간 크로스체크)", [
                ("자본금과적립금조정명세서(갑/을)", "전기 이월결손금 유무 확인, 있으면 과세표준 차감 검증"),
                ("세액공제/감면 크로스체크", "최저한세조정계산서 감면세액/세액공제 = 공제감면세액합계표 소계 = 세액공제조정명세서 일치"),
                ("유형자산 vs 감가상각비조정명세서합계표", "재무상태표 유형+무형자산 = 감가상각비조정명세서합계표 기말현재액"),
                ("감가상각비 vs 표준손익계산서", "각 명세서 감가상각비 = 표준손익계산서 인식 감가상각비"),
                ("소득구분계산서 vs 공제감면세액계산서", "감면분 소득금액 = 감면대상소득 일치"),
                ("소득구분계산서 총 매출액", "분모(총수입금액)에 정확히 반영됐는지 검토"),
                ("자본금과적립금조정명세서(갑) 크로스체크", "자본금·잉여금·유보소득이 재무상태표/소득금액조정합계표 등과 일치"),
                ("세액공제조정명세서 → 농어촌특별세", "공제되는 감면세액이 농어촌특별세 과세표준에 반영됐는지"),
            ]),
            ("세무조정계산서", [
                ("당기순손익 → 각사업연도소득", "결산서상당기순손익 + 익금산입 - 손금산입 = 각사업연도소득금액"),
                ("과세표준 산출", "각사업연도소득 - 이월결손금 - 비과세소득 - 소득공제 = 과세표준"),
                ("법인세 산출세액", "과세표준 × 누진세율 (연환산 적용)"),
                ("농어촌특별세", "감면세액 × 20%"),
            ]),
            ("감면·공제", [
                ("중소기업특별감면", "산출세액 × (감면대상소득 / 총소득) × 공제율"),
                ("소득구분계산서", "감면분 + 기타분 = 각사업연도소득, 과세표준 검증"),
                ("세액공제·감면신청서", "신청서 금액 vs 세무조정계산서 공제감면세액"),
            ]),
            ("추가 검증", [
                ("자산총계 일치", "결산보고서 재무상태표 vs 표준재무상태표"),
                ("신설법인 자본금", "표준재무상태표 자본금 계상 확인"),
                ("소득금액조정 처분", "익금산입 항목의 처분 기재 여부"),
                ("인정이자 미수수익", "인정이자 발생시 결산서 미수수익 반영 확인"),
                ("업무용승용차", "감가상각명세서 업무용승용차 한도 확인"),
            ]),
        ]

        for section, items in CHECKS:
            card = self._card(sc)
            self._section_label(card, section)
            for i, (name, desc) in enumerate(items):
                f = ctk.CTkFrame(card, fg_color="transparent")
                f.pack(fill="x")
                if i < len(items) - 1:
                    ctk.CTkFrame(f, height=1, fg_color=BORDER).pack(fill="x", side="bottom")
                inner = ctk.CTkFrame(f, fg_color="transparent")
                inner.pack(fill="x", pady=7)
                ctk.CTkLabel(inner, text=f"{i+1}.", font=(FONT, SZ_SMALL),
                             text_color=ACCENT, width=24).pack(side="left")
                tf = ctk.CTkFrame(inner, fg_color="transparent")
                tf.pack(side="left", fill="x", expand=True)
                ctk.CTkLabel(tf, text=name, font=(FONT, SZ_BODY, "bold"),
                             text_color=TEXT).pack(anchor="w")
                ctk.CTkLabel(tf, text=desc, font=(FONT, SZ_TINY),
                             text_color=SUB).pack(anchor="w")

    # ════════════════════════════════════════
    #  파싱 데이터
    # ════════════════════════════════════════
    def _page_data(self):
        self._clr()
        self._sel(self.b4)
        if not self.data:
            ctk.CTkLabel(self.main, text="먼저 검토를 실행하세요",
                         font=(FONT, SZ_BODY), text_color=SUB).pack(pady=40)
            return
        sc = self._scroll()
        self._page_title(sc, "파싱 데이터")

        # 파싱 데이터에서 숨길 내부용 키
        _hidden_keys = {"갑_all_amounts", "갑_기말잔액_amounts"}

        def _fmt(v):
            if isinstance(v, bool):
                return "O" if v else "X"
            if isinstance(v, (int, float)) and abs(v) > 999:
                return format_won(v)
            return str(v)

        for sn, sd in self.data.items():
            if sn in ("파일경로", "총페이지수"):
                continue
            card = self._card(sc)
            self._section_label(card, sn)
            if isinstance(sd, dict):
                for k, v in sd.items():
                    if k in _hidden_keys:
                        continue
                    if isinstance(v, list):
                        ctk.CTkLabel(card, text=f"{k}: ({len(v)}건)",
                                     font=(FONT, SZ_TINY), text_color=SUB).pack(anchor="w")
                        for it in v[:10]:
                            if isinstance(it, dict):
                                t = " / ".join(f"{a}: {_fmt(b)}" for a, b in it.items())
                                ctk.CTkLabel(card, text=f"  {t}", font=(FONT, SZ_TINY),
                                             text_color=TEXT, wraplength=700,
                                             justify="left").pack(anchor="w")
                    elif isinstance(v, dict):
                        ctk.CTkLabel(card, text=f"{k}:",
                                     font=(FONT, SZ_TINY, "bold"), text_color=SUB).pack(anchor="w", pady=(4,0))
                        for dk, dv in v.items():
                            if isinstance(dv, dict):
                                ctk.CTkLabel(card, text=f"  {dk}:",
                                             font=(FONT, SZ_TINY), text_color=SUB).pack(anchor="w", padx=10)
                                for dk2, dv2 in dv.items():
                                    self._row(card, f"    {dk2}", _fmt(dv2))
                            else:
                                self._row(card, f"  {dk}", _fmt(dv))
                    else:
                        self._row(card, k, _fmt(v))
            else:
                ctk.CTkLabel(card, text=str(sd), font=(FONT, SZ_SMALL), text_color=TEXT).pack(anchor="w")


if __name__ == "__main__":
    app = App()
    app.mainloop()
