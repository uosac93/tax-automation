"""
병의원 매출 자동 회계처리 프로그램
건보공단 지급통보서 PDF → 분개장 엑셀 + 더존 import CSV
"""
import sys
import os
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

# Windows 콘솔 인코딩
if sys.platform == 'win32':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    except Exception:
        pass

try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False

from pdf_parser import parse_pdf_auto
from journal_engine import JournalGenerator, format_won, DEFAULT_ACCOUNTS
from export_excel import generate_wehago_xls

# ─── 색상 상수 ───
BG = '#1A1A1A'
BG2 = '#1E1E1E'
BG3 = '#2A2A2A'
SIDEBAR_BG = '#1E1E1E'
FG = '#E8E4DE'
FG_DIM = '#7A7A7A'
ACCENT = '#F5F0E0'
CARD_BG = '#222222'
BORDER = '#3A3A3A'
ENTRY_BG = '#2A2A2A'


class MedicalRevenueApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Medical Revenue")
        self.root.geometry("1200x750")
        self.root.minsize(1050, 650)
        self.root.configure(bg=BG)

        # 아이콘
        icon_path = os.path.join(os.path.dirname(__file__), 'medical_revenue.ico')
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception:
                pass

        # Windows 타이틀바 다크모드
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(ctypes.c_int(1)), ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

        # 데이터
        self.parsed_data = None
        self.entries = []
        self.summary = {}
        self.generator = JournalGenerator()
        self.accounts = dict(DEFAULT_ACCOUNTS)
        self.cash_rows = []  # 카드/현금 매출 입력 행들
        self.yoyang_rows = []  # 요양급여 PDF 행들
        self.medical_rows = []  # 의료급여 PDF 행들
        self.current_page = None

        self._setup_styles()
        self._build_ui()
        self._show_page('upload')

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('.', background=BG, foreground=FG, font=('바탕', 10))
        style.configure('TFrame', background=BG)

        style.configure('Treeview', rowheight=32, font=('바탕', 10),
                         background=BG2, foreground=FG, fieldbackground=BG2,
                         borderwidth=0)
        style.configure('Treeview.Heading', font=('바탕', 10, 'bold'),
                         background='#2A2A2A', foreground=ACCENT,
                         borderwidth=0, padding=[0, 6])
        style.map('Treeview',
                   background=[('selected', '#333333')],
                   foreground=[('selected', '#FFFFFF')])

        style.configure('Accent.TButton', font=('바탕', 10, 'bold'),
                         background=ACCENT, foreground='#1A1A1A', padding=[20, 8])
        style.map('Accent.TButton',
                   background=[('active', '#D4B48A'), ('pressed', '#B8935F')])

        style.configure('TButton', background=BG3, foreground=FG, padding=[12, 6])
        style.map('TButton', background=[('active', '#333')])

        style.configure('TLabel', background=BG, foreground=FG)
        style.configure('TEntry', fieldbackground=ENTRY_BG, foreground=FG,
                         borderwidth=1, relief='solid')

        style.configure('TScrollbar', background=BG3, troughcolor=BG,
                         borderwidth=0, arrowsize=0)

    # ════════════════════════════════════════
    # UI 구성
    # ════════════════════════════════════════
    def _build_ui(self):
        # 메인 컨테이너
        main = tk.Frame(self.root, bg=BG)
        main.pack(fill='both', expand=True)

        # ── 좌측 사이드바 ──
        sidebar = tk.Frame(main, bg=SIDEBAR_BG, width=200)
        sidebar.pack(side='left', fill='y')
        sidebar.pack_propagate(False)

        # 사이드바 내용
        tk.Label(sidebar, text="", bg=SIDEBAR_BG, height=1).pack()

        # 메뉴 카테고리
        self._sidebar_category(sidebar, "입력")
        self.btn_upload = self._sidebar_item(sidebar, "병의원 매출", 'upload')
        self.btn_preview = self._sidebar_item(sidebar, "파싱 결과", 'preview')

        self._sidebar_category(sidebar, "회계처리")
        self.btn_summary = self._sidebar_item(sidebar, "매출 현황", 'summary')
        self.btn_journal = self._sidebar_item(sidebar, "분개장", 'journal')

        self._sidebar_category(sidebar, "출력")
        self.btn_export = self._sidebar_item(sidebar, "내보내기", 'export')
        self.btn_settings = self._sidebar_item(sidebar, "계정과목 설정", 'settings')

        # 하단 버전
        tk.Label(sidebar, text="v1.0 · 2026", font=('바탕', 8),
                 bg=SIDEBAR_BG, fg=FG_DIM).pack(side='bottom', pady=10)

        # ── 우측 콘텐츠 ──
        self.content = tk.Frame(main, bg=BG)
        self.content.pack(side='left', fill='both', expand=True, padx=0)

        # 페이지 프레임들
        self.pages = {}
        self._build_upload_page()
        self._build_preview_page()
        self._build_summary_page()
        self._build_journal_page()
        self._build_export_page()
        self._build_settings_page()

        # 하단 상태바
        self.status_var = tk.StringVar(value="PDF 파일을 업로드하세요")
        sbar = tk.Frame(self.root, bg=BG3, height=26)
        sbar.pack(fill='x')
        sbar.pack_propagate(False)
        tk.Label(sbar, textvariable=self.status_var, font=('바탕', 8),
                 bg=BG3, fg=FG_DIM).pack(side='left', padx=12)

    # ── 사이드바 헬퍼 ──
    def _sidebar_category(self, parent, text):
        tk.Label(parent, text=text, font=('바탕', 9, 'bold'),
                 bg=SIDEBAR_BG, fg=ACCENT, anchor='w').pack(
            fill='x', padx=15, pady=(18, 4))

    def _sidebar_item(self, parent, text, page_key):
        # 외부 컨테이너 (좌우 마진용)
        outer = tk.Frame(parent, bg=SIDEBAR_BG)
        outer.pack(fill='x', padx=10, pady=1)

        btn = tk.Label(outer, text="  " + text, font=('바탕', 10),
                       bg=SIDEBAR_BG, fg=FG_DIM, anchor='w', cursor='hand2',
                       padx=10, pady=7)
        btn.pack(fill='x')
        btn.bind('<Button-1>', lambda e: self._show_page(page_key))
        btn.bind('<Enter>', lambda e: btn.configure(bg=BG3) if btn != self._active_btn() else None)
        btn.bind('<Leave>', lambda e: btn.configure(bg=SIDEBAR_BG) if btn != self._active_btn() else None)
        btn._page_key = page_key
        return btn

    def _active_btn(self):
        for attr in ['btn_upload', 'btn_preview', 'btn_summary', 'btn_journal', 'btn_export', 'btn_settings']:
            btn = getattr(self, attr, None)
            if btn and hasattr(btn, '_page_key') and btn._page_key == self.current_page:
                return btn
        return None

    def _show_page(self, page_key):
        # 모든 페이지 숨기기
        for page in self.pages.values():
            page.pack_forget()

        # 선택된 페이지 표시
        if page_key in self.pages:
            self.pages[page_key].pack(fill='both', expand=True)

        self.current_page = page_key

        # 사이드바 버튼 스타일 업데이트
        for attr in ['btn_upload', 'btn_preview', 'btn_summary', 'btn_journal', 'btn_export', 'btn_settings']:
            btn = getattr(self, attr, None)
            if btn:
                if btn._page_key == page_key:
                    btn.configure(bg='#3A3A3A', fg=FG)
                else:
                    btn.configure(bg=SIDEBAR_BG, fg=FG_DIM)

    # ── 섹션 헤더 ──
    def _section_header(self, parent, text):
        lbl = tk.Label(parent, text="ㅣ " + text, font=('바탕', 12, 'bold'),
                       bg=BG, fg=ACCENT, anchor='w')
        lbl.pack(fill='x', padx=30, pady=(25, 15))
        return lbl

    # ── 카드 프레임 ──
    def _card(self, parent, **kwargs):
        frame = tk.Frame(parent, bg=CARD_BG, highlightbackground=BORDER,
                         highlightthickness=1, **kwargs)
        return frame

    # ════════════════════════════════════════
    # 페이지 1: 병의원 매출
    # ════════════════════════════════════════
    def _build_upload_page(self):
        page = tk.Frame(self.content, bg=BG)
        self.pages['upload'] = page

        # 스크롤 가능하게
        canvas = tk.Canvas(page, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(page, orient='vertical', command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=BG)

        scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        # 마우스 휠 스크롤
        def _on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all('<MouseWheel>', _on_mousewheel)

        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        # 스크롤 프레임 폭 맞추기
        def _on_canvas_configure(e):
            canvas.itemconfig(canvas.find_all()[0], width=e.width)
        canvas.bind('<Configure>', _on_canvas_configure)

        sf = scroll_frame  # 축약

        tk.Label(sf, text="병의원 매출", font=('바탕', 14, 'bold'),
                 bg=BG, fg=FG).pack(anchor='w', padx=30, pady=(20, 0))

        # ── PDF 비밀번호 (공통) ──
        pw_card = self._card(sf)
        pw_card.pack(fill='x', padx=30, pady=(15, 10))
        pw_row = tk.Frame(pw_card, bg=CARD_BG)
        pw_row.pack(fill='x', padx=20, pady=12)
        tk.Label(pw_row, text="PDF 비밀번호", font=('바탕', 10), bg=CARD_BG, fg=FG).pack(side='left', padx=(0, 10))
        self.pdf_pw_var = tk.StringVar()
        tk.Entry(pw_row, textvariable=self.pdf_pw_var, font=('바탕', 10),
                 bg=ENTRY_BG, fg=FG, insertbackground=FG, relief='solid',
                 bd=1, width=20, highlightthickness=0).pack(side='left', ipady=5)
        tk.Label(pw_row, text="(모든 PDF에 동일 적용)", font=('바탕', 9),
                 bg=CARD_BG, fg=FG_DIM).pack(side='left', padx=(10, 0))

        # ── 요양급여 ──
        self._section_header(sf, "요양급여 지급통보서 (건강보험)")

        card1 = self._card(sf)
        card1.pack(fill='x', padx=30, pady=(0, 10))

        self.yoyang_drop = tk.Label(card1, text="PDF 파일을 여기에 드래그하거나 클릭하세요",
                                     font=('바탕', 10), bg='#2A2A2A', fg=FG_DIM,
                                     relief='solid', bd=1, height=2, cursor='hand2')
        self.yoyang_drop.pack(fill='x', padx=20, pady=(15, 5))
        self.yoyang_drop.bind('<Button-1>', lambda e: self._select_pdf_dialog('yoyang'))

        self.yoyang_container = tk.Frame(card1, bg=CARD_BG)
        self.yoyang_container.pack(fill='x', padx=20, pady=(0, 5))

        yoyang_btn_row = tk.Frame(card1, bg=CARD_BG)
        yoyang_btn_row.pack(fill='x', padx=20, pady=(0, 15))
        self._make_button(yoyang_btn_row, "전체 삭제", lambda: self._clear_pdf_rows('yoyang'))

        # ── 의료급여 ──
        self._section_header(sf, "의료급여 지급통보서 (의료보호)")

        card2 = self._card(sf)
        card2.pack(fill='x', padx=30, pady=(0, 10))

        self.medical_drop = tk.Label(card2, text="PDF 파일을 여기에 드래그하거나 클릭하세요",
                                      font=('바탕', 10), bg='#2A2A2A', fg=FG_DIM,
                                      relief='solid', bd=1, height=2, cursor='hand2')
        self.medical_drop.pack(fill='x', padx=20, pady=(15, 5))
        self.medical_drop.bind('<Button-1>', lambda e: self._select_pdf_dialog('medical'))

        self.medical_container = tk.Frame(card2, bg=CARD_BG)
        self.medical_container.pack(fill='x', padx=20, pady=(0, 5))

        medical_btn_row = tk.Frame(card2, bg=CARD_BG)
        medical_btn_row.pack(fill='x', padx=20, pady=(0, 15))
        self._make_button(medical_btn_row, "전체 삭제", lambda: self._clear_pdf_rows('medical'))

        # ── 카드/현금 매출 ──
        self._section_header(sf, "카드 / 현금영수증 / 현금 매출")

        card3 = self._card(sf)
        card3.pack(fill='x', padx=30, pady=(0, 10))

        tk.Label(card3, text="매입매출전표 기준 월별 매출을 입력하세요.",
                 font=('바탕', 9), bg=CARD_BG, fg=FG_DIM).pack(
            anchor='w', padx=20, pady=(15, 10))

        self.cash_grid = tk.Frame(card3, bg=CARD_BG)
        self.cash_grid.pack(fill='x', padx=20, pady=(0, 5))

        for i in range(4):
            self.cash_grid.columnconfigure(i, weight=1, uniform='cash')

        for col, text in enumerate(['진료월', '카드매출', '현금영수증', '현금매출']):
            tk.Label(self.cash_grid, text=text, font=('바탕', 10, 'bold'),
                     bg='#2A2A2A', fg=ACCENT, anchor='center', pady=6).grid(
                row=0, column=col, sticky='ew', padx=2, pady=(0, 3))

        self.cash_grid_row = 1
        for _ in range(3):
            self._add_cash_row()

        cash_btn_row = tk.Frame(card3, bg=CARD_BG)
        cash_btn_row.pack(fill='x', padx=20, pady=(10, 15))
        self._make_button(cash_btn_row, "+ 행 추가", self._add_cash_row)
        tk.Frame(cash_btn_row, bg=CARD_BG, width=10).pack(side='left')
        self._make_button(cash_btn_row, "- 마지막 행 삭제", self._remove_cash_row)

        # ── 분석 버튼 ──
        btn_frame = tk.Frame(sf, bg=BG)
        btn_frame.pack(fill='x', padx=30, pady=(15, 20))
        self._make_accent_button(btn_frame, "분석 시작", self._start_analysis)

        # ── 드래그앤드롭 설정 (root 윈도우에 한 번만) ──
        if HAS_WINDND:
            try:
                windnd.hook_dropfiles(self.root, func=self._on_drop_files_auto)
            except Exception:
                pass

    # ════════════════════════════════════════
    # 페이지 2: 파싱 결과
    # ════════════════════════════════════════
    def _add_cash_row(self):
        r = self.cash_grid_row

        month_var = tk.StringVar()
        card_var = tk.StringVar()
        receipt_var = tk.StringVar()
        cash_var = tk.StringVar()

        widgets = []
        for col, (var, just) in enumerate([
            (month_var, 'center'), (card_var, 'right'),
            (receipt_var, 'right'), (cash_var, 'right')
        ]):
            e = tk.Entry(self.cash_grid, textvariable=var, font=('바탕', 10),
                         bg=ENTRY_BG, fg=FG, insertbackground=FG, relief='solid',
                         bd=1, justify=just, highlightthickness=0)
            e.grid(row=r, column=col, sticky='ew', padx=2, pady=2, ipady=4)
            widgets.append(e)

        if not self.cash_rows:
            month_var.set('2025-01')

        self.cash_grid_row += 1
        self.cash_rows.append({
            'widgets': widgets,
            'month': month_var,
            'card': card_var,
            'receipt': receipt_var,
            'cash': cash_var,
        })

    def _remove_cash_row(self):
        if len(self.cash_rows) > 1:
            row_data = self.cash_rows.pop()
            for w in row_data['widgets']:
                w.destroy()
            self.cash_grid_row -= 1

    def _get_cash_records(self):
        """카드/현금 매출 입력값을 레코드로 변환"""
        records = []
        for row in self.cash_rows:
            month = row['month'].get().strip()
            if not month:
                continue
            card_amt = self._parse_amount(row['card'].get())
            receipt_amt = self._parse_amount(row['receipt'].get())
            cash_amt = self._parse_amount(row['cash'].get())
            if card_amt > 0 or receipt_amt > 0 or cash_amt > 0:
                records.append({
                    'month': month,
                    'card_amount': card_amt,
                    'receipt_amount': receipt_amt,
                    'cash_amount': cash_amt,
                })
        return records

    @staticmethod
    def _parse_amount(text):
        if not text:
            return 0
        cleaned = text.replace(',', '').replace(' ', '').replace('원', '').strip()
        try:
            return int(cleaned)
        except ValueError:
            return 0

    def _build_preview_page(self):
        page = tk.Frame(self.content, bg=BG)
        self.pages['preview'] = page

        tk.Label(page, text="파싱 결과", font=('바탕', 14, 'bold'),
                 bg=BG, fg=FG).pack(anchor='w', padx=30, pady=(20, 0))

        self._section_header(page, "PDF 추출 데이터")

        tree_frame = tk.Frame(page, bg=CARD_BG, highlightbackground=BORDER, highlightthickness=1)
        tree_frame.pack(fill='both', expand=True, padx=30, pady=(0, 20))

        cols = ('month', 'total', 'insurer', 'patient', 'payment', 'tax', 'pay_date', 'type')
        self.preview_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=15)
        self.preview_tree.heading('month', text='진료월')
        self.preview_tree.heading('total', text='총진료비')
        self.preview_tree.heading('insurer', text='공단부담금')
        self.preview_tree.heading('patient', text='본인부담금')
        self.preview_tree.heading('payment', text='실지급액')
        self.preview_tree.heading('tax', text='원천징수')
        self.preview_tree.heading('pay_date', text='지급일')
        self.preview_tree.heading('type', text='구분')

        self.preview_tree.column('month', width=90, anchor='center')
        self.preview_tree.column('total', width=120, anchor='e')
        self.preview_tree.column('insurer', width=120, anchor='e')
        self.preview_tree.column('patient', width=100, anchor='e')
        self.preview_tree.column('payment', width=120, anchor='e')
        self.preview_tree.column('tax', width=90, anchor='e')
        self.preview_tree.column('pay_date', width=100, anchor='center')
        self.preview_tree.column('type', width=80, anchor='center')

        self.preview_tree.tag_configure('odd', background='#1E1E1E')
        self.preview_tree.tag_configure('even', background='#252525')
        self.preview_tree.tag_configure('advance', background='#2A2222', foreground='#888888')

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=sb.set)
        self.preview_tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

    # ════════════════════════════════════════
    # 페이지 3: 매출 현황
    # ════════════════════════════════════════
    def _build_summary_page(self):
        page = tk.Frame(self.content, bg=BG)
        self.pages['summary'] = page

        tk.Label(page, text="매출 현황", font=('바탕', 14, 'bold'),
                 bg=BG, fg=FG).pack(anchor='w', padx=30, pady=(20, 0))

        # 기관 정보
        info_frame = tk.Frame(page, bg=BG)
        info_frame.pack(fill='x', padx=30, pady=(15, 0))
        self.inst_label = tk.Label(info_frame, text="기관명: -", font=('바탕', 11, 'bold'),
                                    bg=BG, fg=ACCENT)
        self.inst_label.pack(side='left', padx=(0, 30))
        self.period_label = tk.Label(info_frame, text="기간: -", font=('바탕', 10),
                                      bg=BG, fg=FG_DIM)
        self.period_label.pack(side='left')

        # 합계 카드들
        self._section_header(page, "합계")
        cards = tk.Frame(page, bg=BG)
        cards.pack(fill='x', padx=30, pady=(0, 10))

        self.card_labels = {}
        card_items = [
            ('total', '매출 합계'),
            ('insurance', '요양급여수입'),
            ('medical_aid', '의료급여수입'),
            ('ar', '미수금 발생'),
            ('deposit', '입금액'),
        ]
        for key, title in card_items:
            c = tk.Frame(cards, bg=CARD_BG, highlightbackground=BORDER, highlightthickness=1)
            c.pack(side='left', fill='x', expand=True, padx=4)
            tk.Label(c, text=title, font=('바탕', 9), bg=CARD_BG, fg=FG_DIM).pack(padx=12, pady=(12, 3))
            lbl = tk.Label(c, text="0", font=('바탕', 15, 'bold'), bg=CARD_BG, fg=ACCENT)
            lbl.pack(padx=12, pady=(0, 12))
            self.card_labels[key] = lbl

        # 월별 테이블
        self._section_header(page, "월별 매출 상세")
        tree_frame = tk.Frame(page, bg=CARD_BG, highlightbackground=BORDER, highlightthickness=1)
        tree_frame.pack(fill='both', expand=True, padx=30, pady=(0, 20))

        cols = ('month', 'insurance', 'medical_aid', 'card', 'receipt', 'cash_sale', 'deduct', 'total')
        self.summary_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=10)
        self.summary_tree.heading('month', text='진료월')
        self.summary_tree.heading('insurance', text='요양급여수입')
        self.summary_tree.heading('medical_aid', text='의료급여수입')
        self.summary_tree.heading('card', text='카드매출')
        self.summary_tree.heading('receipt', text='현금영수증')
        self.summary_tree.heading('cash_sale', text='현금매출')
        self.summary_tree.heading('deduct', text='본인부담차감')
        self.summary_tree.heading('total', text='매출합계')

        self.summary_tree.column('month', width=80, anchor='center')
        for c in ['insurance', 'medical_aid', 'card', 'receipt', 'cash_sale', 'deduct', 'total']:
            self.summary_tree.column(c, width=115, anchor='e')

        self.summary_tree.tag_configure('odd', background='#1E1E1E')
        self.summary_tree.tag_configure('even', background='#252525')

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.summary_tree.yview)
        self.summary_tree.configure(yscrollcommand=sb.set)
        self.summary_tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

    # ════════════════════════════════════════
    # 페이지 4: 분개장
    # ════════════════════════════════════════
    def _build_journal_page(self):
        page = tk.Frame(self.content, bg=BG)
        self.pages['journal'] = page

        tk.Label(page, text="분개장", font=('바탕', 14, 'bold'),
                 bg=BG, fg=FG).pack(anchor='w', padx=30, pady=(20, 0))

        # 검증 카드
        verify_card = tk.Frame(page, bg=CARD_BG, highlightbackground=BORDER, highlightthickness=1)
        verify_card.pack(fill='x', padx=30, pady=(15, 0))
        self.verify_label = tk.Label(verify_card, text="  분석을 먼저 실행하세요",
                                      font=('바탕', 11), bg=CARD_BG, fg=FG_DIM,
                                      anchor='w')
        self.verify_label.pack(fill='x', padx=15, pady=12)

        self._section_header(page, "분개 목록")
        tree_frame = tk.Frame(page, bg=CARD_BG, highlightbackground=BORDER, highlightthickness=1)
        tree_frame.pack(fill='both', expand=True, padx=30, pady=(0, 20))

        cols = ('no', 'date', 'desc', 'dr_acct', 'dr_amt', 'cr_acct', 'cr_amt',
                'type', 'partner', 'month')
        self.journal_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=18)
        self.journal_tree.heading('no', text='No')
        self.journal_tree.heading('date', text='전표일자')
        self.journal_tree.heading('desc', text='적요')
        self.journal_tree.heading('dr_acct', text='차변계정')
        self.journal_tree.heading('dr_amt', text='차변금액')
        self.journal_tree.heading('cr_acct', text='대변계정')
        self.journal_tree.heading('cr_amt', text='대변금액')
        self.journal_tree.heading('type', text='유형')
        self.journal_tree.heading('partner', text='거래처')
        self.journal_tree.heading('month', text='귀속월')

        self.journal_tree.column('no', width=40, anchor='center')
        self.journal_tree.column('date', width=95, anchor='center')
        self.journal_tree.column('desc', width=200, anchor='w')
        self.journal_tree.column('dr_acct', width=120, anchor='w')
        self.journal_tree.column('dr_amt', width=110, anchor='e')
        self.journal_tree.column('cr_acct', width=120, anchor='w')
        self.journal_tree.column('cr_amt', width=110, anchor='e')
        self.journal_tree.column('type', width=0, stretch=False)
        self.journal_tree.column('partner', width=70, anchor='center')
        self.journal_tree.column('month', width=75, anchor='center')

        # 줄무늬
        self.journal_tree.tag_configure('odd', background='#1E1E1E')
        self.journal_tree.tag_configure('even', background='#252525')
        self.journal_tree.tag_configure('debit', foreground='#8CB8E0')
        self.journal_tree.tag_configure('credit', foreground='#E0A0A0')

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.journal_tree.yview)
        self.journal_tree.configure(yscrollcommand=sb.set)
        self.journal_tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

    # ════════════════════════════════════════
    # 페이지 5: 내보내기
    # ════════════════════════════════════════
    def _build_export_page(self):
        page = tk.Frame(self.content, bg=BG)
        self.pages['export'] = page

        tk.Label(page, text="내보내기", font=('바탕', 14, 'bold'),
                 bg=BG, fg=FG).pack(anchor='w', padx=30, pady=(20, 0))

        # ── 위하고 매입매출전표 ──
        self._section_header(page, "위하고 매입매출전표")

        card1 = self._card(page)
        card1.pack(fill='x', padx=30, pady=(0, 15))

        tk.Label(card1, text="카드/현금 매출 입력 시 매입매출전표 서식으로 내보냅니다.",
                 font=('바탕', 9), bg=CARD_BG, fg=FG_DIM).pack(
            anchor='w', padx=20, pady=(15, 10))

        inner1 = tk.Frame(card1, bg=CARD_BG)
        inner1.pack(fill='x', padx=20, pady=(0, 20))
        self._make_accent_button(inner1, "위하고 매입매출전표 (.xls)", self._export_wehago)

        # ── 분개장 ──
        self._section_header(page, "분개장")

        card2 = self._card(page)
        card2.pack(fill='x', padx=30, pady=(0, 15))

        tk.Label(card2, text="PDF 분석 결과를 분개장 엑셀 또는 더존 CSV로 내보냅니다.",
                 font=('바탕', 9), bg=CARD_BG, fg=FG_DIM).pack(
            anchor='w', padx=20, pady=(15, 10))

        inner2 = tk.Frame(card2, bg=CARD_BG)
        inner2.pack(fill='x', padx=20, pady=(0, 20))
        self._make_accent_button(inner2, "분개장 엑셀 (.xlsx)", self._export_excel)
        tk.Frame(inner2, bg=CARD_BG, width=10).pack(side='left')
        self._make_accent_button(inner2, "더존 Smart A (.csv)", self._export_csv)

    # ════════════════════════════════════════
    # 페이지 6: 계정과목 설정
    # ════════════════════════════════════════
    def _build_settings_page(self):
        page = tk.Frame(self.content, bg=BG)
        self.pages['settings'] = page

        tk.Label(page, text="계정과목 설정", font=('바탕', 14, 'bold'),
                 bg=BG, fg=FG).pack(anchor='w', padx=30, pady=(20, 0))

        self._section_header(page, "계정과목 코드 / 이름")

        card = self._card(page)
        card.pack(fill='x', padx=30, pady=(0, 15))

        tk.Label(card, text="계정과목을 수정하면 분개에 반영됩니다.",
                 font=('바탕', 9), bg=CARD_BG, fg=FG_DIM).pack(
            anchor='w', padx=20, pady=(15, 10))

        self.acct_entries = {}
        acct_items = [
            ('ar', '미수금'), ('cash', '현금'), ('bank', '보통예금'),
            ('prepaid_tax', '선납세금'), ('prepaid_resident', '선납주민세'),
            ('rev_general', '일반진료수입'), ('rev_insurance', '요양급여수입'),
            ('rev_medical_aid', '의료급여수입'),
        ]

        for key, label in acct_items:
            row = tk.Frame(card, bg=CARD_BG)
            row.pack(fill='x', padx=20, pady=3)

            tk.Label(row, text=label, font=('바탕', 10), bg=CARD_BG, fg=FG,
                     width=14, anchor='w').pack(side='left')

            code_var = tk.StringVar(value=self.accounts[key]['code'])
            name_var = tk.StringVar(value=self.accounts[key]['name'])

            tk.Label(row, text="코드", font=('바탕', 9), bg=CARD_BG, fg=FG_DIM).pack(side='left', padx=(10, 5))
            tk.Entry(row, textvariable=code_var, font=('바탕', 10),
                     bg=ENTRY_BG, fg=FG, insertbackground=FG, relief='solid',
                     bd=1, width=6, highlightthickness=0).pack(side='left', ipady=3)

            tk.Label(row, text="계정명", font=('바탕', 9), bg=CARD_BG, fg=FG_DIM).pack(side='left', padx=(15, 5))
            tk.Entry(row, textvariable=name_var, font=('바탕', 10),
                     bg=ENTRY_BG, fg=FG, insertbackground=FG, relief='solid',
                     bd=1, width=20, highlightthickness=0).pack(side='left', ipady=3)

            self.acct_entries[key] = {'code': code_var, 'name': name_var}

        btn_row = tk.Frame(card, bg=CARD_BG)
        btn_row.pack(fill='x', padx=20, pady=(15, 20))
        self._make_accent_button(btn_row, "계정과목 적용 & 분개 재생성", self._apply_accounts)

    # ── 버튼 헬퍼 ──
    def _make_button(self, parent, text, command):
        btn = tk.Label(parent, text=text, font=('바탕', 9),
                       bg=BG3, fg=FG, cursor='hand2', padx=12, pady=5)
        btn.pack(side='left')
        btn.bind('<Button-1>', lambda e: command())
        btn.bind('<Enter>', lambda e: btn.configure(bg='#333'))
        btn.bind('<Leave>', lambda e: btn.configure(bg=BG3))
        return btn

    def _make_accent_button(self, parent, text, command):
        btn = tk.Label(parent, text=text, font=('바탕', 10, 'bold'),
                       bg=ACCENT, fg='#1A1A1A', cursor='hand2', padx=16, pady=7)
        btn.pack(side='left')
        btn.bind('<Button-1>', lambda e: command())
        btn.bind('<Enter>', lambda e: btn.configure(bg='#E8E4D8'))
        btn.bind('<Leave>', lambda e: btn.configure(bg=ACCENT))
        return btn

    # ════════════════════════════════════════
    # 기능 메서드
    # ════════════════════════════════════════
    def _on_drop_files_auto(self, files):
        """드래그앤드롭: PDF 파일을 자동으로 요양급여/의료급여 분류하여 추가"""
        if self.current_page != 'upload':
            return

        pdf_files = []
        for f in files:
            path = f.decode('utf-8') if isinstance(f, bytes) else str(f)
            if path.lower().endswith('.pdf'):
                pdf_files.append(path)

        if not pdf_files:
            return

        # PDF 내용으로 자동 분류 시도
        password = self.pdf_pw_var.get().strip() or None
        yoyang_count = 0
        medical_count = 0

        for path in pdf_files:
            pdf_type = 'yoyang'  # 기본값
            try:
                from pdf_parser import extract_all_text, detect_pdf_type
                pages, _ = extract_all_text(path, password)
                first_text = pages.get(1, '')
                detected = detect_pdf_type(first_text)
                if detected == 'medical_aid':
                    pdf_type = 'medical'
            except Exception:
                pass

            self._add_pdf_row(pdf_type, path)
            if pdf_type == 'yoyang':
                yoyang_count += 1
            else:
                medical_count += 1

        # 결과 표시
        parts = []
        if yoyang_count:
            parts.append(f"요양급여 {yoyang_count}건")
            self.yoyang_drop.config(text=f"{yoyang_count}개 PDF 추가됨", fg=ACCENT)
            self.root.after(2000, lambda: self.yoyang_drop.config(
                text="PDF 파일을 여기에 드래그하거나 클릭하세요", fg=FG_DIM))
        if medical_count:
            parts.append(f"의료급여 {medical_count}건")
            self.medical_drop.config(text=f"{medical_count}개 PDF 추가됨", fg=ACCENT)
            self.root.after(2000, lambda: self.medical_drop.config(
                text="PDF 파일을 여기에 드래그하거나 클릭하세요", fg=FG_DIM))

        self.status_var.set(f"드래그 추가: {', '.join(parts)}")

    def _on_drop_files(self, files, pdf_type):
        """특정 유형으로 PDF 파일 추가 (수동 분류)"""
        pdf_files = []
        for f in files:
            path = f.decode('utf-8') if isinstance(f, bytes) else str(f)
            if path.lower().endswith('.pdf'):
                pdf_files.append(path)

        if not pdf_files:
            return

        for path in pdf_files:
            self._add_pdf_row(pdf_type, path)

        drop_label = self.yoyang_drop if pdf_type == 'yoyang' else self.medical_drop
        drop_label.config(text=f"{len(pdf_files)}개 PDF 추가됨", fg=ACCENT)
        self.root.after(2000, lambda: drop_label.config(
            text="PDF 파일을 여기에 드래그하거나 클릭하세요", fg=FG_DIM))

    def _select_pdf_dialog(self, pdf_type):
        """드롭 영역 클릭 시 파일 다이얼로그로 여러 파일 선택"""
        titles = {'yoyang': '요양급여 지급통보서 PDF', 'medical': '의료급여 지급통보서 PDF'}
        paths = filedialog.askopenfilenames(
            title=f"{titles.get(pdf_type, 'PDF')} 선택",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")]
        )
        if paths:
            for path in paths:
                self._add_pdf_row(pdf_type, path)

    def _add_pdf_row(self, pdf_type, filepath=''):
        rows = self.yoyang_rows if pdf_type == 'yoyang' else self.medical_rows
        container = self.yoyang_container if pdf_type == 'yoyang' else self.medical_container

        row_frame = tk.Frame(container, bg=CARD_BG)
        row_frame.pack(fill='x', pady=(0, 3))

        path_var = tk.StringVar(value=filepath)

        tk.Label(row_frame, text=f"  {len(rows)+1}.", font=('바탕', 9),
                 bg=CARD_BG, fg=FG_DIM, width=4).pack(side='left')
        tk.Entry(row_frame, textvariable=path_var, font=('바탕', 10),
                 bg=ENTRY_BG, fg=FG, insertbackground=FG, relief='solid',
                 bd=1, highlightthickness=0, readonlybackground=ENTRY_BG,
                 state='readonly').pack(side='left', fill='x', expand=True, ipady=5, padx=(0, 5))

        # X 버튼으로 개별 삭제
        def remove_this():
            row_frame.destroy()
            rows.remove(row_data)
        x_btn = tk.Label(row_frame, text=" X ", font=('바탕', 9),
                         bg='#3A3A3A', fg='#FF7070', cursor='hand2', padx=4, pady=2)
        x_btn.pack(side='left')
        x_btn.bind('<Button-1>', lambda e: remove_this())
        x_btn.bind('<Enter>', lambda e: x_btn.configure(bg='#4A3030'))
        x_btn.bind('<Leave>', lambda e: x_btn.configure(bg='#3A3A3A'))

        row_data = {'frame': row_frame, 'path': path_var}
        rows.append(row_data)

    def _clear_pdf_rows(self, pdf_type):
        rows = self.yoyang_rows if pdf_type == 'yoyang' else self.medical_rows
        for row in rows:
            row['frame'].destroy()
        rows.clear()

    def _start_analysis(self):
        pdf_list = []  # [(path, password, claim_type), ...]
        password = self.pdf_pw_var.get().strip() or None

        for row in self.yoyang_rows:
            p = row['path'].get().strip()
            if p:
                if not os.path.exists(p):
                    messagebox.showwarning("경고", f"요양급여 파일을 찾을 수 없습니다:\n{p}")
                    return
                pdf_list.append((p, password, '요양급여'))

        for row in self.medical_rows:
            p = row['path'].get().strip()
            if p:
                if not os.path.exists(p):
                    messagebox.showwarning("경고", f"의료급여 파일을 찾을 수 없습니다:\n{p}")
                    return
                pdf_list.append((p, password, '의료급여'))

        if not pdf_list:
            messagebox.showwarning("경고", "요양급여 또는 의료급여 PDF를 하나 이상 선택하세요.")
            return

        self.status_var.set("PDF 분석 중...")
        self.root.update()

        thread = threading.Thread(target=self._analyze_pdfs, args=(pdf_list,), daemon=True)
        thread.start()

    def _analyze_pdfs(self, pdf_list):
        try:
            all_records = []
            institution = ''
            period = ''

            for pdf_path, password, claim_type in pdf_list:
                data, detected_type = parse_pdf_auto(pdf_path, password)
                for rec in data.get('records', []):
                    rec['claim_type'] = claim_type
                all_records.extend(data.get('records', []))

                if data.get('institution') and not institution:
                    institution = data['institution']
                if data.get('period') and not period:
                    period = data['period']

            self.parsed_data = {
                'institution': institution,
                'period': period,
                'records': sorted(all_records, key=lambda r: r.get('month', '')),
            }

            self.generator = JournalGenerator(self._get_current_accounts())
            # 가지급 제외하고 분개 생성
            normal_records = [r for r in self.parsed_data['records'] if not r.get('is_advance')]
            insurance_entries = self.generator.generate_from_records(normal_records)
            cash_entries = self.generator.generate_cash_entries(self._get_cash_records())
            self.entries = sorted(insurance_entries + cash_entries, key=lambda e: e.date)
            self.summary = self.generator.get_monthly_summary(self.entries)

            total_records = len(self.parsed_data['records'])
            yoyang_cnt = sum(1 for _, _, t in pdf_list if t == '요양급여')
            medical_cnt = sum(1 for _, _, t in pdf_list if t == '의료급여')
            loaded_parts = []
            if yoyang_cnt:
                loaded_parts.append(f"요양급여 {yoyang_cnt}건")
            if medical_cnt:
                loaded_parts.append(f"의료급여 {medical_cnt}건")
            loaded = ', '.join(loaded_parts)

            self.root.after(0, self._update_all_views)
            self.root.after(0, lambda: self._show_page('preview'))
            self.root.after(0, lambda: self.status_var.set(
                f"분석 완료 [{loaded}] — {total_records}건 추출, {len(self.entries)}건 분개 생성"
            ))

        except ValueError as ve:
            msg = str(ve)
            if msg == "PDF_PASSWORD_REQUIRED":
                self.root.after(0, lambda: messagebox.showwarning("비밀번호 필요",
                    "이 PDF는 비밀번호가 설정되어 있습니다.\n비밀번호를 입력 후 다시 시도하세요."))
            elif msg == "PDF_PASSWORD_WRONG":
                self.root.after(0, lambda: messagebox.showerror("비밀번호 오류",
                    "PDF 비밀번호가 틀립니다."))
            else:
                self.root.after(0, lambda: messagebox.showerror("오류", f"PDF 분석 실패:\n{ve}"))
            self.root.after(0, lambda: self.status_var.set("오류 발생"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("오류", f"PDF 분석 실패:\n{e}"))
            self.root.after(0, lambda: self.status_var.set("오류 발생"))

    def _update_all_views(self):
        self._update_preview()
        self._update_summary()
        self._update_journal()

    def _update_preview(self):
        self.preview_tree.delete(*self.preview_tree.get_children())
        if not self.parsed_data:
            return
        for idx, rec in enumerate(self.parsed_data.get('records', [])):
            tax_total = rec.get('income_tax', 0) + rec.get('resident_tax', 0)
            is_adv = rec.get('is_advance', False)
            row_tag = 'advance' if is_adv else ('odd' if idx % 2 == 0 else 'even')
            type_label = rec.get('claim_type', '')
            if is_adv:
                type_label += ' (가지급)'
            self.preview_tree.insert('', 'end', tags=(row_tag,), values=(
                rec.get('month', ''),
                format_won(rec.get('total_charge', 0)),
                format_won(rec.get('insurer_amount', 0)),
                format_won(rec.get('patient_amount', 0)),
                format_won(rec.get('payment_amount', 0)),
                format_won(tax_total),
                rec.get('payment_date', ''),
                type_label
            ))

    def _update_summary(self):
        if not self.parsed_data:
            return

        self.inst_label.config(text=f"기관명: {self.parsed_data.get('institution', '-')}")
        self.period_label.config(text=f"기간: {self.parsed_data.get('period', '-')}")

        total = sum(s.get('total', 0) for s in self.summary.values())
        insurance = sum(s.get('insurance', 0) for s in self.summary.values())
        medical_aid = sum(s.get('medical_aid', 0) for s in self.summary.values())
        ar = sum(s.get('ar_amount', 0) for s in self.summary.values())
        deposit = sum(s.get('deposit', 0) for s in self.summary.values())

        self.card_labels['total'].config(text=format_won(total))
        self.card_labels['insurance'].config(text=format_won(insurance))
        self.card_labels['medical_aid'].config(text=format_won(medical_aid))
        self.card_labels['ar'].config(text=format_won(ar))
        self.card_labels['deposit'].config(text=format_won(deposit))

        self.summary_tree.delete(*self.summary_tree.get_children())
        for idx, (month, data) in enumerate(self.summary.items()):
            row_tag = 'odd' if idx % 2 == 0 else 'even'
            self.summary_tree.insert('', 'end', values=(
                month,
                format_won(data.get('insurance', 0)),
                format_won(data.get('medical_aid', 0)),
                format_won(data.get('card', 0)),
                format_won(data.get('receipt', 0)),
                format_won(data.get('cash_sale', 0)),
                format_won(data.get('general_deduct', 0)),
                format_won(data.get('total', 0))
            ), tags=(row_tag,))

    def _update_journal(self):
        self.journal_tree.delete(*self.journal_tree.get_children())

        for i, entry in enumerate(self.entries, 1):
            dr_label = f"{entry.debit_account} {entry.debit_name}".strip() if entry.debit_account else ''
            cr_label = f"{entry.credit_account} {entry.credit_name}".strip() if entry.credit_account else ''
            dr_amt = format_won(entry.debit_amount) if entry.debit_amount != 0 else ''
            cr_amt = format_won(entry.credit_amount) if entry.credit_amount != 0 else ''

            row_tag = 'odd' if i % 2 == 1 else 'even'
            self.journal_tree.insert('', 'end', values=(
                i, entry.date, entry.description,
                dr_label, dr_amt, cr_label, cr_amt,
                entry.revenue_type, entry.partner_name, entry.month
            ), tags=(row_tag,))

        validation = self.generator.validate_entries(self.entries)
        if validation['balanced']:
            self.verify_label.config(
                text=f"  차대변 일치  |  차변합계: {format_won(validation['total_debit'])}원  |  "
                     f"대변합계: {format_won(validation['total_credit'])}원  |  "
                     f"분개 {validation['entry_count']}건",
                fg='#7DCE82'
            )
        else:
            self.verify_label.config(
                text=f"  차대변 불일치!  |  차이: {format_won(validation['difference'])}원  |  "
                     f"차변: {format_won(validation['total_debit'])}  /  "
                     f"대변: {format_won(validation['total_credit'])}",
                fg='#FF7070'
            )

    def _get_current_accounts(self):
        accounts = {}
        for key, widgets in self.acct_entries.items():
            accounts[key] = {
                'code': widgets['code'].get(),
                'name': widgets['name'].get()
            }
        return accounts

    def _apply_accounts(self):
        if not self.parsed_data:
            messagebox.showinfo("안내", "먼저 PDF를 분석하세요.")
            return
        self.generator = JournalGenerator(self._get_current_accounts())
        normal_records = [r for r in self.parsed_data.get('records', []) if not r.get('is_advance')]
        insurance_entries = self.generator.generate_from_records(normal_records)
        cash_entries = self.generator.generate_cash_entries(self._get_cash_records())
        self.entries = sorted(insurance_entries + cash_entries, key=lambda e: e.date)
        self.summary = self.generator.get_monthly_summary(self.entries)
        self._update_all_views()
        self.status_var.set("계정과목 변경 적용 완료")

    def _export_excel(self):
        if not self.entries:
            messagebox.showinfo("안내", "먼저 PDF를 분석하세요.")
            return
        path = filedialog.asksaveasfilename(
            title="분개장 엑셀 저장", defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile=f"분개장_{self.parsed_data.get('institution', '병의원')}.xlsx"
        )
        if path:
            from export_excel import generate_journal_excel
            generate_journal_excel(self.entries, self.summary, path,
                                    self.parsed_data.get('institution', ''))
            messagebox.showinfo("완료", f"엑셀 저장 완료:\n{path}")
            self.status_var.set(f"엑셀 저장: {os.path.basename(path)}")

    def _export_csv(self):
        if not self.entries:
            messagebox.showinfo("안내", "먼저 PDF를 분석하세요.")
            return
        path = filedialog.asksaveasfilename(
            title="더존 import CSV 저장", defaultextension=".csv",
            filetypes=[("CSV 파일", "*.csv")],
            initialfile=f"더존import_{self.parsed_data.get('institution', '병의원')}.csv"
        )
        if path:
            from export_excel import generate_douzone_csv
            generate_douzone_csv(self.entries, path)
            messagebox.showinfo("완료", f"CSV 저장 완료:\n{path}")
            self.status_var.set(f"CSV 저장: {os.path.basename(path)}")

    def _export_wehago(self):
        cash_records = self._get_cash_records()
        if not cash_records:
            messagebox.showinfo("안내", "카드/현금영수증/현금 매출 데이터를 입력하세요.\n병의원 매출 탭에서 입력 후 다시 시도하세요.")
            return
        inst = self.parsed_data.get('institution', '병의원') if self.parsed_data else '병의원'
        path = filedialog.asksaveasfilename(
            title="위하고 매입매출전표 저장", defaultextension=".xls",
            filetypes=[("Excel 97-2003", "*.xls")],
            initialfile=f"매입매출전표_{inst}.xls"
        )
        if path:
            generate_wehago_xls(cash_records, path)
            messagebox.showinfo("완료", f"위하고 매입매출전표 저장 완료:\n{path}")
            self.status_var.set(f"위하고 저장: {os.path.basename(path)}")



def main():
    root = tk.Tk()
    app = MedicalRevenueApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
