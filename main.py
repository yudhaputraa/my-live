import os
import re
import random
import datetime
import flet as ft
from database import DatabaseManager
from books_calendar_ui import BooksCalendarMixin
from excel_exporter import generate_oxford_excel, export_full_app_report

# =============================================================================
# PALET WARNA SESUAI PERMINTAAN USER
# #0E0B1D (ditulis #OEOB1D), #2A1E4D, #5A3B8C, #8B7CC7, #C9C2E3, #3A3A52
# =============================================================================
C_BG_DEEP = "#0E0B1D"       # Background utama paling gelap
C_CARD_DARK = "#2A1E4D"     # Card / permukaan kontainer utama
C_PRIMARY = "#5A3B8C"       # Aksen ungu utama (tombol aktif, header)
C_SECONDARY = "#8B7CC7"     # Aksen ungu muda (highlight, badge, grafik)
C_TEXT_LIGHT = "#C9C2E3"    # Teks utama terang
C_MUTED = "#3A3A52"         # Border, input background, elemen non-aktif

# Warna indikator status pendukung yang harmonis dengan palet ungu
C_SUCCESS = "#4ADE80"
C_WARNING = "#FBBF24"
C_DANGER = "#F87171"
C_WHITE = "#FFFFFF"

MONTH_NAMES_ID = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

MONTH_SHORT_ID = [
    "", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
    "Jul", "Agu", "Sep", "Okt", "Nov", "Des"
]


def format_rp(amount: float) -> str:
    try:
        val = int(round(float(amount)))
        sign = "-" if val < 0 else ""
        return f"{sign}Rp {abs(val):,}".replace(",", ".")
    except Exception:
        return "Rp 0"


def format_rp_compact(amount: float) -> str:
    try:
        val = float(amount)
        sign = "-" if val < 0 else ""
        aval = abs(val)
        if aval >= 1_000_000_000:
            return f"{sign}Rp {aval / 1_000_000_000:.1f}M"
        if aval >= 1_000_000:
            return f"{sign}Rp {aval / 1_000_000:.1f}Jt"
        if aval >= 1_000:
            return f"{sign}Rp {aval / 1_000:.0f}Rb"
        return f"{sign}Rp {int(aval)}"
    except Exception:
        return "Rp 0"


def parse_penjelasan(penjelasan: str):
    if not penjelasan:
        return "", ""
    parts = re.split(r"\bcontoh(?:\s+kalimat)?\s*:\s*", str(penjelasan), flags=re.IGNORECASE)
    desc = parts[0].strip()
    example = parts[1].strip().strip("'\"") if len(parts) > 1 else ""
    return desc, example


def card_shadow(blur: int = 14, color: str = "#070510") -> ft.BoxShadow:
    return ft.BoxShadow(
        spread_radius=0,
        blur_radius=blur,
        color=color,
        offset=ft.Offset(0, 4),
    )


def hero_gradient() -> ft.LinearGradient:
    return ft.LinearGradient(
        begin=ft.Alignment.TOP_LEFT,
        end=ft.Alignment.BOTTOM_RIGHT,
        colors=[C_PRIMARY, C_CARD_DARK],
    )


def dark_card_gradient() -> ft.LinearGradient:
    return ft.LinearGradient(
        begin=ft.Alignment.TOP_LEFT,
        end=ft.Alignment.BOTTOM_RIGHT,
        colors=[C_CARD_DARK, "#1B1336"],
    )


class MyLifeApp(BooksCalendarMixin):
    def __init__(self, page: ft.Page):
        self.page = page
        self.db = DatabaseManager()

        # State Navigasi Utama & Sub-Tab
        self.main_tab_index = 0
        self.vocab_subtab = 0      # 0: Daftar Kata, 1: Flashcard, 2: Latihan & Tebakan
        self.tracker_subtab = 0    # 0: Tracker Bulanan, 1: Tracker Tahunan
        self.finance_subtab = 0    # 0: Transaksi, 1: Overview Bulanan, 2: Overview Tahunan
        self.setup_subtab = 0      # 0: Set-Up Budget, 1: Tracker Langganan, 2: Saving Goals

        # State Filter Kosakata
        self.vocab_search = ""
        self.vocab_status_filter = "semua"
        self.vocab_level_filter = "semua"
        self.vocab_fav_only = False
        self.vocab_page_offset = 0
        self.vocab_page_size = 25
        self.vocab_results_host = ft.Container(expand=True)

        # State Flashcard
        self.flashcard_list = []
        self.flashcard_index = 0
        self.flashcard_flipped = False
        self.flashcard_filter = "semua"

        # State Quiz / Tebakan
        self.quiz_source = "hafal"         # 'hafal', 'belajar_hafal', 'semua'
        self.quiz_mode = "pilihan_ganda"   # 'pilihan_ganda', 'arti_ke_inggris', 'ketik_kata'
        self.quiz_questions = []
        self.quiz_distractors = []
        self.quiz_current_idx = 0
        self.quiz_correct_count = 0
        self.quiz_answered = False
        self.quiz_selected_opt_id = None
        self.quiz_is_last_correct = False
        self.quiz_feedback_title = ""
        self.quiz_feedback_desc = ""
        self.quiz_feedback_ex = ""
        self.quiz_feedback_color = C_SECONDARY
        self.quiz_active = False

        # State Filter Transaksi Keuangan
        self.tx_tipe_filter = "SEMUA"

        # State Bulan & Tahun untuk Tracker & Keuangan
        today = datetime.date.today()
        self.selected_year = today.year
        self.selected_month = today.month

        self._init_books_calendar_state()
        self.setup_page()
        self.render()

    def setup_page(self):
        self.page.title = "MyLife — Oxford 5000 Vocab & Smart Finance Tracker"
        self.page.bgcolor = C_BG_DEEP
        self.page.padding = 0
        try:
            self.page.window.width = 450
            self.page.window.height = 860
            self.page.window.bgcolor = C_BG_DEEP
        except Exception:
            pass

        self.body_container = ft.Container(
            expand=True,
            bgcolor=C_BG_DEEP,
            padding=ft.Padding.only(left=12, right=12, top=8, bottom=4),
        )

        self.nav_bar = ft.NavigationBar(
            selected_index=self.main_tab_index,
            bgcolor=C_CARD_DARK,
            indicator_color=C_PRIMARY,
            on_change=self.on_main_nav_change,
            destinations=[
                ft.NavigationBarDestination(
                    icon=ft.Icons.HOME_OUTLINED,
                    selected_icon=ft.Icons.HOME_ROUNDED,
                    label="Beranda",
                ),
                ft.NavigationBarDestination(
                    icon=ft.Icons.MENU_BOOK_OUTLINED,
                    selected_icon=ft.Icons.MENU_BOOK_ROUNDED,
                    label="Kosakata",
                ),
                ft.NavigationBarDestination(
                    icon=ft.Icons.INSIGHTS_OUTLINED,
                    selected_icon=ft.Icons.INSIGHTS_ROUNDED,
                    label="Tracker",
                ),
                ft.NavigationBarDestination(
                    icon=ft.Icons.ACCOUNT_BALANCE_WALLET_OUTLINED,
                    selected_icon=ft.Icons.ACCOUNT_BALANCE_WALLET_ROUNDED,
                    label="Keuangan",
                ),
                ft.NavigationBarDestination(
                    icon=ft.Icons.TUNE_OUTLINED,
                    selected_icon=ft.Icons.TUNE_ROUNDED,
                    label="Set-Up",
                ),
                ft.NavigationBarDestination(
                    icon=ft.Icons.AUTO_STORIES_OUTLINED,
                    selected_icon=ft.Icons.AUTO_STORIES_ROUNDED,
                    label="Buku",
                ),
                ft.NavigationBarDestination(
                    icon=ft.Icons.CALENDAR_MONTH_OUTLINED,
                    selected_icon=ft.Icons.CALENDAR_MONTH_ROUNDED,
                    label="Kalender",
                ),
            ],
        )

        self.header_bar = ft.Container(
            gradient=ft.LinearGradient(
                begin=ft.Alignment.CENTER_LEFT,
                end=ft.Alignment.CENTER_RIGHT,
                colors=[C_CARD_DARK, "#1F153C"],
            ),
            padding=ft.Padding.symmetric(horizontal=14, vertical=10),
            border=ft.Border.only(
                bottom=ft.border.BorderSide(1, C_MUTED) if hasattr(ft.border, "BorderSide") else None
            ),
            content=ft.Row(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    ft.Row(
                        spacing=10,
                        controls=[
                            ft.Container(
                                width=38,
                                height=38,
                                border_radius=ft.BorderRadius.all(11),
                                gradient=hero_gradient(),
                                border=ft.Border.all(1, C_SECONDARY),
                                alignment=ft.Alignment.CENTER,
                                content=ft.Icon(ft.Icons.AUTO_AWESOME, color=C_WHITE, size=19),
                            ),
                            ft.Column(
                                spacing=1,
                                controls=[
                                    ft.Text(
                                        "MyLife Vocab & Finance",
                                        size=15,
                                        weight=ft.FontWeight.BOLD,
                                        color=C_WHITE,
                                    ),
                                    ft.Text(
                                        "Oxford 5000™ • 50/30/20 Smart Budget",
                                        size=10.5,
                                        color=C_SECONDARY,
                                    ),
                                ],
                            ),
                        ],
                    ),
                    ft.Container(
                        bgcolor=C_PRIMARY,
                        border_radius=ft.BorderRadius.all(10),
                        border=ft.Border.all(1, C_SECONDARY),
                        padding=ft.Padding.symmetric(horizontal=10, vertical=6),
                        on_click=self.export_all_to_excel,
                        tooltip="Ekspor Excel Kosakata & Laporan Keuangan",
                        content=ft.Row(
                            spacing=4,
                            controls=[
                                ft.Icon(ft.Icons.FILE_DOWNLOAD_OUTLINED, color=C_WHITE, size=15),
                                ft.Text("Excel", size=11, weight=ft.FontWeight.BOLD, color=C_WHITE),
                            ],
                        ),
                    ),
                ],
            ),
        )

        self.page.navigation_bar = self.nav_bar

        self.page.add(
            ft.Column(
                expand=True,
                spacing=0,
                controls=[
                    self.header_bar,
                    self.body_container,
                ],
            )
        )

    def notify(self, message: str, color: str = C_PRIMARY):
        snack = ft.SnackBar(
            content=ft.Text(message, color=C_WHITE, size=13, weight=ft.FontWeight.W_500),
            bgcolor=color,
            duration=2500,
        )
        self.page.show_dialog(snack)

    def on_main_nav_change(self, e):
        self.main_tab_index = int(e.control.selected_index)
        if self.main_tab_index == 1 and self.vocab_subtab == 1:
            self.flashcard_list = []
        self.render()

    def build_subtab_bar(self, labels: list, active_idx: int, on_click_fn) -> ft.Container:
        pills = []
        for idx, text in enumerate(labels):
            is_active = idx == active_idx
            pills.append(
                ft.Container(
                    expand=True,
                    bgcolor=C_PRIMARY if is_active else C_BG_DEEP,
                    padding=ft.Padding.symmetric(horizontal=8, vertical=8),
                    border_radius=ft.BorderRadius.all(10),
                    border=ft.Border.all(1, C_SECONDARY if is_active else C_MUTED),
                    alignment=ft.Alignment.CENTER,
                    on_click=lambda e, i=idx: on_click_fn(i),
                    content=ft.Text(
                        text,
                        size=11.5,
                        weight=ft.FontWeight.BOLD if is_active else ft.FontWeight.W_500,
                        color=C_WHITE if is_active else C_TEXT_LIGHT,
                        text_align=ft.TextAlign.CENTER,
                        max_lines=1,
                        overflow=ft.TextOverflow.ELLIPSIS,
                    ),
                )
            )
        return ft.Container(
            bgcolor=C_CARD_DARK,
            padding=ft.Padding.all(5),
            border_radius=ft.BorderRadius.all(13),
            border=ft.Border.all(1, C_MUTED),
            margin=ft.Margin.only(bottom=8),
            content=ft.Row(spacing=5, controls=pills),
        )

    def build_period_selector(self, show_month: bool = True, on_change_callback=None) -> ft.Container:
        years = [str(y) for y in range(2023, 2030)]
        months = [(str(m), MONTH_NAMES_ID[m]) for m in range(1, 13)]

        def on_year_select(e):
            try:
                self.selected_year = int(e.control.value)
                if on_change_callback:
                    on_change_callback()
                else:
                    self.render()
            except Exception:
                pass

        def on_month_select(e):
            try:
                self.selected_month = int(e.control.value)
                if on_change_callback:
                    on_change_callback()
                else:
                    self.render()
            except Exception:
                pass

        controls = []
        if show_month:
            controls.append(
                ft.Dropdown(
                    value=str(self.selected_month),
                    expand=3,
                    dense=True,
                    color=C_WHITE,
                    bgcolor=C_MUTED,
                    border_color=C_SECONDARY,
                    border_radius=ft.BorderRadius.all(8),
                    content_padding=ft.Padding.symmetric(horizontal=10, vertical=4),
                    text_size=12,
                    options=[ft.dropdown.Option(key=k, text=t) for k, t in months],
                    on_select=on_month_select,
                )
            )

        controls.append(
            ft.Dropdown(
                value=str(self.selected_year),
                expand=2 if show_month else True,
                dense=True,
                color=C_WHITE,
                bgcolor=C_MUTED,
                border_color=C_SECONDARY,
                border_radius=ft.BorderRadius.all(8),
                content_padding=ft.Padding.symmetric(horizontal=10, vertical=4),
                text_size=12,
                options=[ft.dropdown.Option(key=y, text=y) for y in years],
                on_select=on_year_select,
            )
        )

        return ft.Container(
            bgcolor=C_CARD_DARK,
            padding=ft.Padding.symmetric(horizontal=12, vertical=8),
            border_radius=ft.BorderRadius.all(12),
            border=ft.Border.all(1, C_MUTED),
            content=ft.Row(
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=8,
                controls=[
                    ft.Row(
                        spacing=5,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Icon(ft.Icons.CALENDAR_MONTH_ROUNDED, color=C_SECONDARY, size=17),
                            ft.Text("Periode:", size=12, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                        ],
                    ),
                    ft.Row(
                        spacing=6,
                        expand=True,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=controls,
                    ),
                ],
            ),
        )

    def render(self):
        if self.main_tab_index == 0:
            self.body_container.content = self.build_dashboard_view()
        elif self.main_tab_index == 1:
            self.body_container.content = self.build_vocab_view()
        elif self.main_tab_index == 2:
            self.body_container.content = self.build_tracker_view()
        elif self.main_tab_index == 3:
            self.body_container.content = self.build_finance_view()
        elif self.main_tab_index == 4:
            self.body_container.content = self.build_setup_view()
        elif self.main_tab_index == 5:
            self.body_container.content = self.build_books_view()
        elif self.main_tab_index == 6:
            self.body_container.content = self.build_calendar_view()
        self.page.update()

    # =========================================================================
    # TAB 0: DASHBOARD (BERANDA)
    # =========================================================================
    def build_dashboard_view(self) -> ft.Control:
        v_sum = self.db.get_vocab_summary()
        f_ov = self.db.get_monthly_finance_overview(self.selected_year, self.selected_month)
        subs = self.db.get_subscriptions()

        pct_vocab = (v_sum["hafal"] / v_sum["total"] * 100.0) if v_sum["total"] > 0 else 0.0

        # Pilih 1 Kata Hari Ini secara deterministik per tanggal
        today_ordinal = datetime.date.today().toordinal()
        rows, _ = self.db.query_vocab(limit=1, offset=(today_ordinal % max(1, v_sum["total"])))
        word_of_day = rows[0] if rows else {
            "id": 1,
            "word": "Abolish",
            "pos_level": "v. C1",
            "arti": "Menghapuskan, mengakhiri secara resmi",
            "penjelasan": "Menghentikan berlakunya suatu hukum atau sistem secara resmi. Contoh: 'They voted to abolish the old tax.'",
            "status": "belum",
        }

        def quick_memorize_wod(e):
            new_st = "belum" if word_of_day.get("status") == "hafal" else "hafal"
            self.db.update_vocab_status(word_of_day["id"], new_st)
            if new_st == "hafal":
                self.notify(f"Kata '{word_of_day['word']}' ditandai SUDAH HAFAL ✅", C_PRIMARY)
            else:
                self.notify(f"Status '{word_of_day['word']}' di-reset ke Belum Hafal", C_MUTED)
            self.render()

        wod_desc, wod_ex = parse_penjelasan(word_of_day.get("penjelasan", ""))

        return ft.ListView(
            expand=True,
            spacing=12,
            padding=ft.Padding.only(bottom=24),
            controls=[
                # Hero Banner Sambutan
                ft.Container(
                    gradient=hero_gradient(),
                    padding=ft.Padding.all(16),
                    border_radius=ft.BorderRadius.all(16),
                    border=ft.Border.all(1, C_SECONDARY),
                    shadow=card_shadow(),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Column(
                                        spacing=2,
                                        expand=True,
                                        controls=[
                                            ft.Text(
                                                "Selamat Datang di MyLife ✨",
                                                size=17,
                                                weight=ft.FontWeight.BOLD,
                                                color=C_WHITE,
                                            ),
                                            ft.Text(
                                                "Kuasai 2.000 kosakata Oxford 5000™ & kelola budget 50/30/20 Anda.",
                                                size=11.5,
                                                color=C_TEXT_LIGHT,
                                            ),
                                        ],
                                    ),
                                    ft.Container(
                                        bgcolor=C_BG_DEEP,
                                        padding=ft.Padding.symmetric(horizontal=10, vertical=5),
                                        border_radius=ft.BorderRadius.all(20),
                                        border=ft.Border.all(1, C_SECONDARY),
                                        content=ft.Text(
                                            f"{MONTH_SHORT_ID[self.selected_month]} {self.selected_year}",
                                            size=11,
                                            color=C_TEXT_LIGHT,
                                            weight=ft.FontWeight.BOLD,
                                        ),
                                    ),
                                ],
                            ),
                        ],
                    ),
                ),

                # Ringkasan Kosakata Card (2x2 Grid agar tidak sempit di HP)
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(15),
                    border=ft.Border.all(1, C_MUTED),
                    shadow=card_shadow(),
                    content=ft.Column(
                        spacing=10,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Row(
                                        spacing=8,
                                        controls=[
                                            ft.Container(
                                                width=28,
                                                height=28,
                                                bgcolor=C_PRIMARY,
                                                border_radius=ft.BorderRadius.all(8),
                                                alignment=ft.Alignment.CENTER,
                                                content=ft.Icon(ft.Icons.SCHOOL_ROUNDED, color=C_WHITE, size=16),
                                            ),
                                            ft.Text(
                                                "Progres Hafalan Oxford 5000",
                                                size=13.5,
                                                weight=ft.FontWeight.BOLD,
                                                color=C_WHITE,
                                            ),
                                        ],
                                    ),
                                    ft.Container(
                                        bgcolor=C_BG_DEEP,
                                        padding=ft.Padding.symmetric(horizontal=8, vertical=3),
                                        border_radius=ft.BorderRadius.all(8),
                                        border=ft.Border.all(1, C_SECONDARY),
                                        content=ft.Text(
                                            f"{v_sum['hafal']}/{v_sum['total']} ({pct_vocab:.1f}%)",
                                            size=11,
                                            weight=ft.FontWeight.BOLD,
                                            color=C_SECONDARY,
                                        ),
                                    ),
                                ],
                            ),
                            ft.ProgressBar(
                                value=min(1.0, pct_vocab / 100.0),
                                bar_height=8,
                                color=C_SECONDARY,
                                bgcolor=C_MUTED,
                                border_radius=ft.BorderRadius.all(6),
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._stat_mini_box("Sudah Hafal", f"{v_sum['hafal']} Kata", C_SUCCESS, ft.Icons.CHECK_CIRCLE_OUTLINE),
                                    self._stat_mini_box("Sedang Belajar", f"{v_sum['belajar']} Kata", C_WARNING, ft.Icons.BOOKMARK_BORDER),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._stat_mini_box("Hafal Hari Ini", f"+{v_sum['today_memorized']} Kata", C_SECONDARY, ft.Icons.TODAY_ROUNDED),
                                    self._stat_mini_box("Rata2 Skor Kuis", f"{v_sum['avg_quiz_score']}%", C_TEXT_LIGHT, ft.Icons.EMOJI_EVENTS_OUTLINED),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    ft.Button(
                                        content="Daftar Kosakata",
                                        icon=ft.Icons.MENU_BOOK_ROUNDED,
                                        bgcolor=C_PRIMARY,
                                        color=C_WHITE,
                                        expand=True,
                                        on_click=lambda e: self._jump_to(1, 0),
                                    ),
                                    ft.Button(
                                        content="Latihan Tebakan",
                                        icon=ft.Icons.QUIZ_ROUNDED,
                                        bgcolor=C_MUTED,
                                        color=C_TEXT_LIGHT,
                                        expand=True,
                                        on_click=lambda e: self._jump_to(1, 2),
                                    ),
                                ],
                            ),
                        ],
                    ),
                ),

                # Word of the Day Card
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(15),
                    border=ft.Border.all(1, C_PRIMARY),
                    shadow=card_shadow(),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Row(
                                        spacing=6,
                                        controls=[
                                            ft.Icon(ft.Icons.LIGHTBULB_ROUNDED, color=C_WARNING, size=16),
                                            ft.Text("KATA PILIHAN HARI INI", size=11, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                                        ],
                                    ),
                                    ft.Container(
                                        bgcolor=C_PRIMARY,
                                        padding=ft.Padding.symmetric(horizontal=8, vertical=2),
                                        border_radius=ft.BorderRadius.all(6),
                                        content=ft.Text(word_of_day.get("pos_level", ""), size=10.5, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                    ),
                                ],
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Text(word_of_day.get("word", ""), size=22, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                    ft.Text(f"#{word_of_day.get('id', 1)}", size=11, color=C_SECONDARY),
                                ],
                            ),
                            ft.Text(f"Arti: {word_of_day.get('arti', '')}", size=13.5, weight=ft.FontWeight.W_600, color=C_SECONDARY),
                            ft.Text(wod_desc, size=12, color=C_TEXT_LIGHT),
                            *(
                                [
                                    ft.Container(
                                        bgcolor=C_BG_DEEP,
                                        padding=ft.Padding.symmetric(horizontal=10, vertical=7),
                                        border_radius=ft.BorderRadius.all(8),
                                        border=ft.Border.all(1, C_MUTED),
                                        content=ft.Row(
                                            spacing=6,
                                            vertical_alignment=ft.CrossAxisAlignment.START,
                                            controls=[
                                                ft.Text("💬", size=11),
                                                ft.Text(f'"{wod_ex}"', size=11.5, color=C_SECONDARY, italic=True, expand=True),
                                            ],
                                        ),
                                    )
                                ]
                                if wod_ex
                                else []
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.TextButton(
                                        content="Buka Flashcard 🔄",
                                        on_click=lambda e: self._jump_to(1, 1),
                                    ),
                                    ft.Button(
                                        content="Tandai Sudah Hafal ✅" if word_of_day.get("status") != "hafal" else "Sudah Dihafal ✓",
                                        bgcolor=C_PRIMARY if word_of_day.get("status") != "hafal" else C_SUCCESS,
                                        color=C_WHITE if word_of_day.get("status") != "hafal" else C_BG_DEEP,
                                        on_click=quick_memorize_wod,
                                    ),
                                ],
                            ),
                        ],
                    ),
                ),

                # Ringkasan Keuangan Bulan Ini Card
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(15),
                    border=ft.Border.all(1, C_MUTED),
                    shadow=card_shadow(),
                    content=ft.Column(
                        spacing=10,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Row(
                                        spacing=8,
                                        controls=[
                                            ft.Container(
                                                width=28,
                                                height=28,
                                                bgcolor=C_PRIMARY,
                                                border_radius=ft.BorderRadius.all(8),
                                                alignment=ft.Alignment.CENTER,
                                                content=ft.Icon(ft.Icons.ACCOUNT_BALANCE_WALLET_ROUNDED, color=C_WHITE, size=16),
                                            ),
                                            ft.Text(
                                                f"Keuangan {MONTH_NAMES_ID[self.selected_month]} {self.selected_year}",
                                                size=13.5,
                                                weight=ft.FontWeight.BOLD,
                                                color=C_WHITE,
                                            ),
                                        ],
                                    ),
                                    ft.Container(
                                        bgcolor=C_BG_DEEP,
                                        padding=ft.Padding.symmetric(horizontal=8, vertical=3),
                                        border_radius=ft.BorderRadius.all(8),
                                        border=ft.Border.all(1, C_SECONDARY),
                                        content=ft.Text(
                                            f"Rule {f_ov['rule']['mode']}",
                                            size=10.5,
                                            color=C_SECONDARY,
                                            weight=ft.FontWeight.BOLD,
                                        ),
                                    ),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box("Total Pendapatan", format_rp(f_ov["pendapatan"]), C_SUCCESS),
                                    self._finance_box(
                                        "Selisih / Sisa Saldo",
                                        format_rp(f_ov["selisih"]),
                                        C_SUCCESS if f_ov["selisih"] >= 0 else C_DANGER,
                                    ),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box(
                                        f"Kebutuhan ({f_ov['rule']['kebutuhan_actual_pct']}%)",
                                        format_rp(f_ov["kebutuhan"]),
                                        C_TEXT_LIGHT,
                                    ),
                                    self._finance_box(
                                        f"Keinginan ({f_ov['rule']['keinginan_actual_pct']}%)",
                                        format_rp(f_ov["keinginan"]),
                                        C_WARNING,
                                    ),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box(
                                        f"Tabungan ({f_ov['rule']['tabungan_actual_pct']}%)",
                                        format_rp(f_ov["tabungan"]),
                                        C_SECONDARY,
                                    ),
                                    self._finance_box(
                                        "Beban Langganan/Bln",
                                        format_rp(subs["total_bulanan"]),
                                        C_TEXT_LIGHT,
                                    ),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    ft.Button(
                                        content="+ Catat Transaksi",
                                        icon=ft.Icons.ADD_CARD_ROUNDED,
                                        bgcolor=C_PRIMARY,
                                        color=C_WHITE,
                                        expand=True,
                                        on_click=lambda e: self._jump_to(3, 0),
                                    ),
                                    ft.Button(
                                        content="Analisis Budget",
                                        icon=ft.Icons.PIE_CHART_ROUNDED,
                                        bgcolor=C_MUTED,
                                        color=C_TEXT_LIGHT,
                                        expand=True,
                                        on_click=lambda e: self._jump_to(3, 1),
                                    ),
                                ],
                            ),
                        ],
                    ),
                ),
            ],
        )

    def _jump_to(self, main_idx: int, sub_idx: int):
        self.main_tab_index = main_idx
        self.nav_bar.selected_index = main_idx
        if main_idx == 1:
            self.vocab_subtab = sub_idx
            if sub_idx == 1:
                self.flashcard_list = []
        elif main_idx == 2:
            self.tracker_subtab = sub_idx
        elif main_idx == 3:
            self.finance_subtab = sub_idx
        elif main_idx == 4:
            self.setup_subtab = sub_idx
        self.render()

    def _stat_mini_box(self, label: str, val: str, val_color: str, icon_name=None) -> ft.Container:
        top_controls = []
        if icon_name:
            top_controls.append(ft.Icon(icon_name, color=val_color, size=15))
        top_controls.append(
            ft.Text(val, size=13.5, weight=ft.FontWeight.BOLD, color=val_color, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS)
        )
        return ft.Container(
            expand=True,
            bgcolor=C_BG_DEEP,
            padding=ft.Padding.symmetric(horizontal=10, vertical=8),
            border_radius=ft.BorderRadius.all(10),
            border=ft.Border.all(1, C_MUTED),
            content=ft.Column(
                spacing=2,
                controls=[
                    ft.Row(spacing=5, vertical_alignment=ft.CrossAxisAlignment.CENTER, controls=top_controls),
                    ft.Text(label, size=10.5, color=C_TEXT_LIGHT, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                ],
            ),
        )

    def _finance_box(self, label: str, val: str, val_color: str) -> ft.Container:
        return ft.Container(
            expand=True,
            bgcolor=C_BG_DEEP,
            padding=ft.Padding.symmetric(horizontal=10, vertical=8),
            border_radius=ft.BorderRadius.all(10),
            border=ft.Border.all(1, C_MUTED),
            content=ft.Column(
                spacing=2,
                controls=[
                    ft.Text(label, size=10.5, color=C_SECONDARY, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                    ft.Text(val, size=12.5, weight=ft.FontWeight.BOLD, color=val_color, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                ],
            ),
        )

    def export_all_to_excel(self, e=None):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        report_path = os.path.join(base_dir, "Laporan_Lengkap_MyLife.xlsx")
        vocab_rows, _ = self.db.query_vocab(limit=5000, offset=0)
        oxford_path = os.path.join(base_dir, "Oxford_5000_Kosakata.xlsx")
        generate_oxford_excel(vocab_rows, oxford_path)
        export_full_app_report(self.db, report_path)
        self.notify("Berhasil diekspor: Oxford_5000_Kosakata.xlsx & Laporan_Lengkap_MyLife.xlsx!", C_PRIMARY)

    # =========================================================================
    # TAB 1: KOSAKATA & LATIHAN (OXFORD 5000)
    # =========================================================================
    def build_vocab_view(self) -> ft.Control:
        def switch_sub(idx: int):
            self.vocab_subtab = idx
            if idx == 1:
                self.flashcard_list = []
                self.flashcard_flipped = False
            self.render()

        sub_bar = self.build_subtab_bar(
            ["Daftar 2000 Kata", "Flashcard", "Latihan & Tebakan"],
            self.vocab_subtab,
            switch_sub,
        )

        if self.vocab_subtab == 0:
            content = self.build_vocab_list_subview()
        elif self.vocab_subtab == 1:
            content = self.build_flashcard_subview()
        else:
            content = self.build_quiz_subview()

        return ft.Column(expand=True, spacing=4, controls=[sub_bar, content])

    def _populate_vocab_results_host(self):
        """Updates only the vocabulary results container so the search TextField never loses focus."""
        rows, total_matching = self.db.query_vocab(
            search=self.vocab_search,
            status_filter=self.vocab_status_filter,
            level_filter=self.vocab_level_filter,
            fav_only=self.vocab_fav_only,
            limit=self.vocab_page_size,
            offset=self.vocab_page_offset,
        )

        def set_status_filter(val: str):
            self.vocab_status_filter = val
            self.vocab_page_offset = 0
            self._populate_vocab_results_host()
            self.page.update()

        def set_level_filter(val: str):
            self.vocab_level_filter = val
            self.vocab_page_offset = 0
            self._populate_vocab_results_host()
            self.page.update()

        def change_word_status(word_id: int, new_st: str):
            self.db.update_vocab_status(word_id, new_st)
            label_map = {"hafal": "SUDAH HAFAL ✅", "belajar": "SEDANG BELAJAR 📖", "belum": "BELUM HAFAL"}
            self.notify(f"Status kosakata diubah ke {label_map.get(new_st, new_st)}", C_PRIMARY)
            self._populate_vocab_results_host()
            self.page.update()

        def toggle_word_fav(word_id: int):
            self.db.toggle_favorite(word_id)
            self._populate_vocab_results_host()
            self.page.update()

        # Filter Chips Bar
        status_chips = []
        for code, lbl in [("semua", "Semua"), ("belum", "Belum"), ("belajar", "Belajar 📖"), ("hafal", "Hafal ✅")]:
            active = self.vocab_status_filter == code
            status_chips.append(
                ft.Container(
                    bgcolor=C_PRIMARY if active else C_MUTED,
                    padding=ft.Padding.symmetric(horizontal=10, vertical=5),
                    border_radius=ft.BorderRadius.all(16),
                    border=ft.Border.all(1, C_SECONDARY if active else C_MUTED),
                    on_click=lambda e, c=code: set_status_filter(c),
                    content=ft.Text(
                        lbl,
                        size=11,
                        color=C_WHITE if active else C_TEXT_LIGHT,
                        weight=ft.FontWeight.BOLD if active else ft.FontWeight.NORMAL,
                    ),
                )
            )

        for lvl in ["B2", "C1"]:
            active_l = self.vocab_level_filter == lvl
            status_chips.append(
                ft.Container(
                    bgcolor=C_SECONDARY if active_l else C_BG_DEEP,
                    padding=ft.Padding.symmetric(horizontal=9, vertical=5),
                    border_radius=ft.BorderRadius.all(16),
                    border=ft.Border.all(1, C_SECONDARY),
                    on_click=lambda e, l=lvl: set_level_filter("semua" if self.vocab_level_filter == l else l),
                    content=ft.Text(
                        f"Level {lvl}",
                        size=11,
                        color=C_BG_DEEP if active_l else C_SECONDARY,
                        weight=ft.FontWeight.BOLD,
                    ),
                )
            )

        word_cards = []
        for item in rows:
            st = item["status"]
            st_color = C_SUCCESS if st == "hafal" else (C_WARNING if st == "belajar" else C_SECONDARY)
            st_badge = "SUDAH HAFAL" if st == "hafal" else ("SEDANG BELAJAR" if st == "belajar" else "BELUM HAFAL")
            item_desc, item_ex = parse_penjelasan(item["penjelasan"])

            word_cards.append(
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(12),
                    border_radius=ft.BorderRadius.all(13),
                    border=ft.Border.all(1.2 if st == "hafal" else 1, C_PRIMARY if st == "hafal" else C_MUTED),
                    content=ft.Column(
                        spacing=6,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Row(
                                        spacing=6,
                                        expand=True,
                                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                        controls=[
                                            ft.Text(f"#{item['id']}", size=11, color=C_SECONDARY),
                                            ft.Text(
                                                item["word"],
                                                size=15.5,
                                                weight=ft.FontWeight.BOLD,
                                                color=C_WHITE,
                                                max_lines=1,
                                                overflow=ft.TextOverflow.ELLIPSIS,
                                            ),
                                            ft.Container(
                                                bgcolor=C_MUTED,
                                                padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                                                border_radius=ft.BorderRadius.all(6),
                                                content=ft.Text(item["pos_level"] or "", size=10, color=C_TEXT_LIGHT),
                                            ),
                                        ],
                                    ),
                                    ft.Row(
                                        spacing=2,
                                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                        controls=[
                                            ft.Container(
                                                bgcolor=C_BG_DEEP,
                                                padding=ft.Padding.symmetric(horizontal=7, vertical=3),
                                                border_radius=ft.BorderRadius.all(8),
                                                border=ft.Border.all(1, st_color),
                                                content=ft.Text(st_badge, size=8.5, weight=ft.FontWeight.BOLD, color=st_color),
                                            ),
                                            ft.IconButton(
                                                icon=ft.Icons.STAR_ROUNDED if item["is_favorite"] else ft.Icons.STAR_BORDER_ROUNDED,
                                                icon_color=C_WARNING if item["is_favorite"] else C_SECONDARY,
                                                icon_size=18,
                                                width=30,
                                                height=30,
                                                style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                                on_click=lambda e, wid=item["id"]: toggle_word_fav(wid),
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                            ft.Text(f"Arti: {item['arti']}", size=13, weight=ft.FontWeight.W_600, color=C_SECONDARY),
                            ft.Text(item_desc, size=11.5, color=C_TEXT_LIGHT),
                            *(
                                [
                                    ft.Container(
                                        bgcolor=C_BG_DEEP,
                                        padding=ft.Padding.symmetric(horizontal=8, vertical=5),
                                        border_radius=ft.BorderRadius.all(7),
                                        border=ft.Border.all(1, C_MUTED),
                                        content=ft.Row(
                                            spacing=6,
                                            vertical_alignment=ft.CrossAxisAlignment.START,
                                            controls=[
                                                ft.Text("💬", size=10),
                                                ft.Text(f'"{item_ex}"', size=10.5, color=C_SECONDARY, italic=True, expand=True),
                                            ],
                                        ),
                                    )
                                ]
                                if item_ex
                                else []
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Text(
                                        f"Kuis: {item['times_correct']}/{item['times_quizzed']}"
                                        + (f" • {item['memorized_date']}" if item.get("memorized_date") else ""),
                                        size=10,
                                        color=C_SECONDARY,
                                    ),
                                    ft.Row(
                                        spacing=5,
                                        controls=[
                                            ft.Container(
                                                bgcolor=C_MUTED if st != "belum" else C_BG_DEEP,
                                                padding=ft.Padding.symmetric(horizontal=8, vertical=4),
                                                border_radius=ft.BorderRadius.all(6),
                                                on_click=lambda e, wid=item["id"]: change_word_status(wid, "belum"),
                                                content=ft.Text("Reset", size=10, color=C_TEXT_LIGHT),
                                            ),
                                            ft.Container(
                                                bgcolor=C_WARNING if st == "belajar" else C_MUTED,
                                                padding=ft.Padding.symmetric(horizontal=8, vertical=4),
                                                border_radius=ft.BorderRadius.all(6),
                                                on_click=lambda e, wid=item["id"]: change_word_status(wid, "belajar"),
                                                content=ft.Text(
                                                    "Belajar 📖",
                                                    size=10,
                                                    color=C_BG_DEEP if st == "belajar" else C_WHITE,
                                                    weight=ft.FontWeight.BOLD,
                                                ),
                                            ),
                                            ft.Container(
                                                bgcolor=C_PRIMARY if st != "hafal" else C_SUCCESS,
                                                padding=ft.Padding.symmetric(horizontal=10, vertical=4),
                                                border_radius=ft.BorderRadius.all(6),
                                                on_click=lambda e, wid=item["id"]: change_word_status(wid, "hafal"),
                                                content=ft.Text(
                                                    "Hafal ✅",
                                                    size=10,
                                                    color=C_BG_DEEP if st == "hafal" else C_WHITE,
                                                    weight=ft.FontWeight.BOLD,
                                                ),
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),
                )
            )

        start_num = self.vocab_page_offset + 1 if total_matching > 0 else 0
        end_num = min(total_matching, self.vocab_page_offset + self.vocab_page_size)
        total_pages = max(1, (total_matching + self.vocab_page_size - 1) // self.vocab_page_size)
        curr_page = (self.vocab_page_offset // self.vocab_page_size) + 1

        def prev_page(e):
            if self.vocab_page_offset >= self.vocab_page_size:
                self.vocab_page_offset -= self.vocab_page_size
                self._populate_vocab_results_host()
                self.page.update()

        def next_page(e):
            if self.vocab_page_offset + self.vocab_page_size < total_matching:
                self.vocab_page_offset += self.vocab_page_size
                self._populate_vocab_results_host()
                self.page.update()

        pagination_row = ft.Row(
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                ft.Text(
                    f"Hal {curr_page}/{total_pages} ({start_num}–{end_num} dari {total_matching})",
                    size=11,
                    color=C_SECONDARY,
                    weight=ft.FontWeight.BOLD,
                ),
                ft.Row(
                    spacing=6,
                    controls=[
                        ft.Button(
                            content="◀ Prev",
                            bgcolor=C_MUTED,
                            color=C_TEXT_LIGHT,
                            disabled=(self.vocab_page_offset == 0),
                            on_click=prev_page,
                        ),
                        ft.Button(
                            content="Next ▶",
                            bgcolor=C_PRIMARY if (self.vocab_page_offset + self.vocab_page_size < total_matching) else C_MUTED,
                            color=C_WHITE if (self.vocab_page_offset + self.vocab_page_size < total_matching) else C_TEXT_LIGHT,
                            disabled=(self.vocab_page_offset + self.vocab_page_size >= total_matching),
                            on_click=next_page,
                        ),
                    ],
                ),
            ],
        )

        self.vocab_results_host.content = ft.Column(
            expand=True,
            spacing=8,
            controls=[
                ft.Row(spacing=6, wrap=True, controls=status_chips),
                pagination_row,
                ft.ListView(
                    expand=True,
                    spacing=8,
                    padding=ft.Padding.only(bottom=24),
                    controls=(
                        [*word_cards, ft.Container(padding=ft.Padding.symmetric(vertical=8), content=pagination_row)]
                        if word_cards
                        else [
                            ft.Container(
                                bgcolor=C_CARD_DARK,
                                padding=ft.Padding.all(24),
                                border_radius=ft.BorderRadius.all(14),
                                alignment=ft.Alignment.CENTER,
                                content=ft.Text("Tidak ada kosakata yang cocok dengan pencarian/filter ini.", color=C_SECONDARY),
                            )
                        ]
                    ),
                ),
            ],
        )

    def build_vocab_list_subview(self) -> ft.Control:
        def on_search_change(e):
            self.vocab_search = e.control.value or ""
            self.vocab_page_offset = 0
            self._populate_vocab_results_host()
            self.page.update()

        def toggle_fav_filter(e):
            self.vocab_fav_only = not self.vocab_fav_only
            self.vocab_page_offset = 0
            fav_btn.icon = ft.Icons.STAR_ROUNDED if self.vocab_fav_only else ft.Icons.STAR_BORDER_ROUNDED
            fav_btn.icon_color = C_WARNING if self.vocab_fav_only else C_TEXT_LIGHT
            self._populate_vocab_results_host()
            self.page.update()

        def open_add_vocab_dialog(e):
            tf_word = ft.TextField(label="Kata Bahasa Inggris", color=C_TEXT_LIGHT, bgcolor=C_MUTED)
            tf_pos = ft.TextField(label="Kelas Kata & Level (misal: n. B2 / v. C1)", value="n. B2", color=C_TEXT_LIGHT, bgcolor=C_MUTED)
            tf_arti = ft.TextField(label="Arti Bahasa Indonesia", color=C_TEXT_LIGHT, bgcolor=C_MUTED)
            tf_penj = ft.TextField(label="Penjelasan & Contoh Kalimat", multiline=True, min_lines=2, color=C_TEXT_LIGHT, bgcolor=C_MUTED)

            def save_new_word(ev):
                if not tf_word.value.strip() or not tf_arti.value.strip():
                    self.notify("Kata Inggris dan Arti wajib diisi!", C_DANGER)
                    return
                self.db.add_custom_vocab(tf_word.value, tf_pos.value, tf_arti.value, tf_penj.value or tf_arti.value)
                self.page.pop_dialog()
                self.notify(f"Kosakata '{tf_word.value}' berhasil ditambahkan!", C_PRIMARY)
                self._populate_vocab_results_host()
                self.page.update()

            dlg = ft.AlertDialog(
                bgcolor=C_CARD_DARK,
                title=ft.Text("Tambah Kosakata Baru", color=C_TEXT_LIGHT, weight=ft.FontWeight.BOLD),
                content=ft.Column(tight=True, spacing=8, controls=[tf_word, tf_pos, tf_arti, tf_penj]),
                actions=[
                    ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                    ft.Button(content="Simpan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save_new_word),
                ],
            )
            self.page.show_dialog(dlg)

        fav_btn = ft.IconButton(
            icon=ft.Icons.STAR_ROUNDED if self.vocab_fav_only else ft.Icons.STAR_BORDER_ROUNDED,
            icon_color=C_WARNING if self.vocab_fav_only else C_TEXT_LIGHT,
            bgcolor=C_CARD_DARK,
            tooltip="Filter Favorit",
            on_click=toggle_fav_filter,
        )

        search_row = ft.Row(
            spacing=8,
            controls=[
                ft.TextField(
                    hint_text="Cari kata Inggris atau arti Indonesia...",
                    value=self.vocab_search,
                    prefix_icon=ft.Icons.SEARCH,
                    color=C_TEXT_LIGHT,
                    bgcolor=C_CARD_DARK,
                    border_color=C_MUTED,
                    focused_border_color=C_SECONDARY,
                    border_radius=ft.BorderRadius.all(12),
                    dense=True,
                    text_size=13,
                    expand=True,
                    on_change=on_search_change,
                ),
                fav_btn,
                ft.IconButton(
                    icon=ft.Icons.ADD_CIRCLE_OUTLINE,
                    icon_color=C_WHITE,
                    bgcolor=C_PRIMARY,
                    tooltip="Tambah Kosakata Baru",
                    on_click=open_add_vocab_dialog,
                ),
            ],
        )

        self._populate_vocab_results_host()

        return ft.Column(
            expand=True,
            spacing=8,
            controls=[
                search_row,
                self.vocab_results_host,
            ],
        )

    def build_flashcard_subview(self) -> ft.Control:
        if not self.flashcard_list:
            rows, _ = self.db.query_vocab(
                status_filter=self.flashcard_filter if self.flashcard_filter != "semua" else "semua",
                limit=300,
                offset=0,
            )
            self.flashcard_list = rows
            self.flashcard_index = max(0, min(self.flashcard_index, len(self.flashcard_list) - 1))

        filter_chips = ft.Row(
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            controls=[
                ft.Row(
                    spacing=5,
                    wrap=True,
                    controls=[
                        ft.Container(
                            bgcolor=C_PRIMARY if self.flashcard_filter == f_code else C_MUTED,
                            padding=ft.Padding.symmetric(horizontal=10, vertical=5),
                            border_radius=ft.BorderRadius.all(14),
                            border=ft.Border.all(1, C_SECONDARY if self.flashcard_filter == f_code else C_MUTED),
                            on_click=lambda e, fc=f_code: self._set_flashcard_filter(fc),
                            content=ft.Text(f_lbl, size=11, color=C_WHITE, weight=ft.FontWeight.BOLD if self.flashcard_filter == f_code else ft.FontWeight.NORMAL),
                        )
                        for f_code, f_lbl in [("semua", "Semua"), ("belum", "Belum"), ("belajar", "Belajar 📖"), ("hafal", "Hafal ✅")]
                    ],
                ),
                ft.IconButton(
                    icon=ft.Icons.SHUFFLE_ROUNDED,
                    icon_color=C_TEXT_LIGHT,
                    bgcolor=C_CARD_DARK,
                    tooltip="Acak Urutan Kartu",
                    on_click=lambda e: self._shuffle_flashcards(),
                ),
            ],
        )

        if not self.flashcard_list:
            return ft.Column(
                expand=True,
                spacing=12,
                controls=[
                    filter_chips,
                    ft.Container(
                        expand=True,
                        bgcolor=C_CARD_DARK,
                        border_radius=ft.BorderRadius.all(16),
                        border=ft.Border.all(1, C_MUTED),
                        alignment=ft.Alignment.CENTER,
                        padding=ft.Padding.all(24),
                        content=ft.Column(
                            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                            alignment=ft.MainAxisAlignment.CENTER,
                            spacing=10,
                            controls=[
                                ft.Icon(ft.Icons.STYLE_OUTLINED, color=C_SECONDARY, size=48),
                                ft.Text("Belum ada kosakata pada kategori filter ini.", color=C_TEXT_LIGHT, size=14, text_align=ft.TextAlign.CENTER),
                                ft.Button(
                                    content="Tampilkan Semua 2000 Kata",
                                    bgcolor=C_PRIMARY,
                                    color=C_WHITE,
                                    on_click=lambda e: self._set_flashcard_filter("semua"),
                                ),
                            ],
                        ),
                    ),
                ],
            )

        self.flashcard_index = max(0, min(self.flashcard_index, len(self.flashcard_list) - 1))
        card = self.flashcard_list[self.flashcard_index]
        card_desc, card_ex = parse_penjelasan(card.get("penjelasan", ""))
        st = card.get("status", "belum")
        st_color = C_SUCCESS if st == "hafal" else (C_WARNING if st == "belajar" else C_SECONDARY)
        st_label = "SUDAH HAFAL ✅" if st == "hafal" else ("SEDANG BELAJAR 📖" if st == "belajar" else "BELUM HAFAL")

        def flip_card(e):
            self.flashcard_flipped = not self.flashcard_flipped
            self.render()

        def next_card(e):
            if self.flashcard_list:
                self.flashcard_index = (self.flashcard_index + 1) % len(self.flashcard_list)
                self.flashcard_flipped = False
                self.render()

        def prev_card(e):
            if self.flashcard_list:
                self.flashcard_index = (self.flashcard_index - 1) % len(self.flashcard_list)
                self.flashcard_flipped = False
                self.render()

        def toggle_fav_card(e):
            new_fav = self.db.toggle_favorite(card["id"])
            card["is_favorite"] = new_fav
            self.render()

        def mark_card(new_status: str):
            self.db.update_vocab_status(card["id"], new_status)
            card["status"] = new_status
            self.notify(f"'{card['word']}' ditandai sebagai {new_status.upper()}!", C_PRIMARY)
            if self.flashcard_filter != "semua" and self.flashcard_filter != new_status:
                self.flashcard_list.pop(self.flashcard_index)
                if self.flashcard_index >= len(self.flashcard_list):
                    self.flashcard_index = 0
                self.flashcard_flipped = False
                self.render()
            else:
                next_card(None)

        progress_val = (self.flashcard_index + 1) / max(1, len(self.flashcard_list))

        return ft.Column(
            expand=True,
            spacing=10,
            controls=[
                filter_chips,
                ft.Row(
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    controls=[
                        ft.Text(
                            f"Kartu {self.flashcard_index + 1} / {len(self.flashcard_list)}",
                            size=12,
                            weight=ft.FontWeight.BOLD,
                            color=C_SECONDARY,
                        ),
                        ft.Text(
                            "Ketuk kartu untuk membalik (Flip) 🔄",
                            size=11.5,
                            color=C_TEXT_LIGHT,
                        ),
                    ],
                ),
                ft.ProgressBar(
                    value=progress_val,
                    bar_height=6,
                    color=C_SECONDARY,
                    bgcolor=C_MUTED,
                    border_radius=ft.BorderRadius.all(4),
                ),
                ft.Container(
                    expand=True,
                    gradient=hero_gradient() if self.flashcard_flipped else dark_card_gradient(),
                    padding=ft.Padding.all(20),
                    border_radius=ft.BorderRadius.all(20),
                    border=ft.Border.all(2, C_SECONDARY if self.flashcard_flipped else C_PRIMARY),
                    shadow=card_shadow(18),
                    on_click=flip_card,
                    content=ft.Column(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            # Top Row of Card: ID, Status Badge, Favorite button
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Container(
                                        bgcolor=C_BG_DEEP,
                                        padding=ft.Padding.symmetric(horizontal=10, vertical=4),
                                        border_radius=ft.BorderRadius.all(10),
                                        border=ft.Border.all(1, st_color),
                                        content=ft.Text(st_label, size=10, weight=ft.FontWeight.BOLD, color=st_color),
                                    ),
                                    ft.Row(
                                        spacing=4,
                                        controls=[
                                            ft.Container(
                                                bgcolor=C_MUTED,
                                                padding=ft.Padding.symmetric(horizontal=10, vertical=4),
                                                border_radius=ft.BorderRadius.all(10),
                                                content=ft.Text(f"#{card['id']} • {card['pos_level']}", size=11, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                            ),
                                            ft.IconButton(
                                                icon=ft.Icons.STAR_ROUNDED if card.get("is_favorite") else ft.Icons.STAR_BORDER_ROUNDED,
                                                icon_color=C_WARNING if card.get("is_favorite") else C_TEXT_LIGHT,
                                                icon_size=20,
                                                on_click=toggle_fav_card,
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                            # Center Content of Flashcard
                            ft.Column(
                                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                spacing=10,
                                controls=[
                                    ft.Text(
                                        card["word"],
                                        size=32,
                                        weight=ft.FontWeight.BOLD,
                                        color=C_WHITE,
                                        text_align=ft.TextAlign.CENTER,
                                    ),
                                    ft.Divider(color=C_SECONDARY if self.flashcard_flipped else C_MUTED),
                                    ft.Text(
                                        card["arti"] if self.flashcard_flipped else "👆 Ketuk kartu untuk melihat Arti & Penjelasan",
                                        size=18 if self.flashcard_flipped else 13,
                                        weight=ft.FontWeight.BOLD if self.flashcard_flipped else ft.FontWeight.NORMAL,
                                        color=C_WHITE if self.flashcard_flipped else C_SECONDARY,
                                        text_align=ft.TextAlign.CENTER,
                                    ),
                                    *(
                                        [
                                            ft.Text(
                                                card_desc,
                                                size=12.5,
                                                color=C_TEXT_LIGHT,
                                                text_align=ft.TextAlign.CENTER,
                                            )
                                        ]
                                        if self.flashcard_flipped and card_desc
                                        else []
                                    ),
                                    *(
                                        [
                                            ft.Container(
                                                bgcolor=C_BG_DEEP,
                                                padding=ft.Padding.symmetric(horizontal=12, vertical=8),
                                                border_radius=ft.BorderRadius.all(10),
                                                border=ft.Border.all(1, C_SECONDARY),
                                                content=ft.Text(
                                                    f'💬 "{card_ex}"',
                                                    size=11.5,
                                                    italic=True,
                                                    color=C_TEXT_LIGHT,
                                                    text_align=ft.TextAlign.CENTER,
                                                ),
                                            )
                                        ]
                                        if self.flashcard_flipped and card_ex
                                        else []
                                    ),
                                ],
                            ),
                            # Bottom Footer hint
                            ft.Text(
                                "SISI BELAKANG (ARTI & CONTOH)" if self.flashcard_flipped else "SISI DEPAN (KATA INGGRIS)",
                                size=10,
                                weight=ft.FontWeight.BOLD,
                                color=C_TEXT_LIGHT if self.flashcard_flipped else C_SECONDARY,
                            ),
                        ],
                    ),
                ),
                ft.Row(
                    spacing=8,
                    controls=[
                        ft.Button(content="◀ Prev", bgcolor=C_MUTED, color=C_TEXT_LIGHT, expand=True, on_click=prev_card),
                        ft.Button(content="Balik 🔄", bgcolor=C_CARD_DARK, color=C_SECONDARY, expand=True, on_click=flip_card),
                        ft.Button(content="Next ▶", bgcolor=C_MUTED, color=C_TEXT_LIGHT, expand=True, on_click=next_card),
                    ],
                ),
                ft.Row(
                    spacing=8,
                    controls=[
                        ft.Button(
                            content="Masih Belajar 📖",
                            bgcolor=C_WARNING,
                            color=C_BG_DEEP,
                            expand=True,
                            on_click=lambda e: mark_card("belajar"),
                        ),
                        ft.Button(
                            content="Sudah Hafal ✅",
                            bgcolor=C_SUCCESS,
                            color=C_BG_DEEP,
                            expand=True,
                            on_click=lambda e: mark_card("hafal"),
                        ),
                    ],
                ),
            ],
        )

    def _shuffle_flashcards(self):
        if self.flashcard_list:
            random.shuffle(self.flashcard_list)
            self.flashcard_index = 0
            self.flashcard_flipped = False
            self.render()

    def _set_flashcard_filter(self, f_code: str):
        self.flashcard_filter = f_code
        self.flashcard_list = []
        self.flashcard_index = 0
        self.flashcard_flipped = False
        self.render()

    def build_quiz_subview(self) -> ft.Control:
        v_sum = self.db.get_vocab_summary()

        def start_new_quiz(e=None):
            selected, distractors = self.db.get_quiz_words(source_filter=self.quiz_source, count=10)
            if not selected:
                self.notify(
                    "Kategori sumber ini masih kosong! Klik 'Tandai 5 Kata Pertama' atau pilih sumber 'Semua 2000 Kata'.",
                    C_WARNING,
                )
                return
            self.quiz_questions = selected
            self.quiz_distractors = distractors
            self.quiz_current_idx = 0
            self.quiz_correct_count = 0
            self.quiz_answered = False
            self.quiz_selected_opt_id = None
            self.quiz_feedback_title = ""
            self.quiz_feedback_desc = ""
            self.quiz_feedback_ex = ""
            self.quiz_active = True
            self.render()

        def quick_seed_5_memorized(e):
            rows, _ = self.db.query_vocab(status_filter="belum", limit=5, offset=0)
            for r in rows:
                self.db.update_vocab_status(r["id"], "hafal")
            self.notify("5 Kosakata pertama berhasil ditandai Sudah Hafal! Silakan mulai kuis.", C_SUCCESS)
            self.render()

        if not self.quiz_active:
            empty_hafal_banner = []
            if self.quiz_source == "hafal" and v_sum["hafal"] == 0:
                empty_hafal_banner.append(
                    ft.Container(
                        bgcolor=C_BG_DEEP,
                        padding=ft.Padding.all(12),
                        border_radius=ft.BorderRadius.all(10),
                        border=ft.Border.all(1, C_WARNING),
                        content=ft.Column(
                            spacing=6,
                            controls=[
                                ft.Text(
                                    "💡 Anda belum menandai kata sebagai 'Sudah Hafal'.",
                                    size=12,
                                    weight=ft.FontWeight.BOLD,
                                    color=C_WARNING,
                                ),
                                ft.Text(
                                    "Anda dapat menandai 5 kata pertama secara instan untuk mencoba latihan, atau memilih sumber 'Semua 2000 Kata'.",
                                    size=11,
                                    color=C_TEXT_LIGHT,
                                ),
                                ft.Row(
                                    spacing=8,
                                    controls=[
                                        ft.Button(
                                            content="⚡ Tandai 5 Kata Pertama",
                                            bgcolor=C_PRIMARY,
                                            color=C_WHITE,
                                            on_click=quick_seed_5_memorized,
                                        ),
                                        ft.Button(
                                            content="Pakai Semua Kata",
                                            bgcolor=C_MUTED,
                                            color=C_TEXT_LIGHT,
                                            on_click=lambda e: self._set_quiz_opt(source="semua"),
                                        ),
                                    ],
                                ),
                            ],
                        ),
                    )
                )

            return ft.ListView(
                expand=True,
                spacing=12,
                padding=ft.Padding.only(bottom=24),
                controls=[
                    ft.Container(
                        gradient=dark_card_gradient(),
                        padding=ft.Padding.all(16),
                        border_radius=ft.BorderRadius.all(15),
                        border=ft.Border.all(1, C_PRIMARY),
                        shadow=card_shadow(),
                        content=ft.Column(
                            spacing=12,
                            controls=[
                                ft.Row(
                                    spacing=10,
                                    controls=[
                                        ft.Container(
                                            width=36,
                                            height=36,
                                            bgcolor=C_PRIMARY,
                                            border_radius=ft.BorderRadius.all(10),
                                            alignment=ft.Alignment.CENTER,
                                            content=ft.Icon(ft.Icons.QUIZ_ROUNDED, color=C_WHITE, size=20),
                                        ),
                                        ft.Column(
                                            spacing=1,
                                            expand=True,
                                            controls=[
                                                ft.Text("Latihan & Tebakan Kosakata", size=16, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                                ft.Text(
                                                    f"Total Kuis Diselesaikan: {v_sum['total_quizzes']} Sesi • Rata-rata Skor: {v_sum['avg_quiz_score']}%",
                                                    size=11,
                                                    color=C_SECONDARY,
                                                ),
                                            ],
                                        ),
                                    ],
                                ),
                                ft.Text(
                                    "Uji ingatan kosakata yang telah Anda hafal! Setiap sesi tebakan akan tercatat otomatis pada grafik Tracker Bulanan dan Tahunan.",
                                    size=12,
                                    color=C_TEXT_LIGHT,
                                ),
                                *empty_hafal_banner,
                                ft.Divider(color=C_MUTED),
                                ft.Text("1. Pilih Sumber Soal Kosakata:", size=12.5, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                                ft.Row(
                                    spacing=6,
                                    wrap=True,
                                    controls=[
                                        self._choice_chip(
                                            f"Sudah Dihafal ({v_sum['hafal']})",
                                            self.quiz_source == "hafal",
                                            lambda: self._set_quiz_opt(source="hafal"),
                                        ),
                                        self._choice_chip(
                                            f"Belajar + Hafal ({v_sum['hafal'] + v_sum['belajar']})",
                                            self.quiz_source == "belajar_hafal",
                                            lambda: self._set_quiz_opt(source="belajar_hafal"),
                                        ),
                                        self._choice_chip(
                                            "Semua 2000 Kata",
                                            self.quiz_source == "semua",
                                            lambda: self._set_quiz_opt(source="semua"),
                                        ),
                                    ],
                                ),
                                ft.Text("2. Pilih Mode Permainan Tebakan:", size=12.5, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                                ft.Row(
                                    spacing=6,
                                    wrap=True,
                                    controls=[
                                        self._choice_chip(
                                            "Tebak Arti (Inggris ➔ Indo)",
                                            self.quiz_mode == "pilihan_ganda",
                                            lambda: self._set_quiz_opt(mode="pilihan_ganda"),
                                        ),
                                        self._choice_chip(
                                            "Tebak Kata (Indo ➔ Inggris)",
                                            self.quiz_mode == "arti_ke_inggris",
                                            lambda: self._set_quiz_opt(mode="arti_ke_inggris"),
                                        ),
                                        self._choice_chip(
                                            "Ketik Kata (Spelling Test)",
                                            self.quiz_mode == "ketik_kata",
                                            lambda: self._set_quiz_opt(mode="ketik_kata"),
                                        ),
                                    ],
                                ),
                                ft.Button(
                                    content="Mulai Sesi Tebakan 🚀",
                                    bgcolor=C_PRIMARY,
                                    color=C_WHITE,
                                    on_click=start_new_quiz,
                                ),
                            ],
                        ),
                    ),
                ],
            )

        # Jika kuis selesai
        if self.quiz_current_idx >= len(self.quiz_questions):
            total_q = len(self.quiz_questions)
            score_pct = (self.quiz_correct_count / total_q * 100.0) if total_q > 0 else 0.0
            grade_msg = "Luar Biasa! Penguasaan Sangat Baik 🌟" if score_pct >= 80 else ("Bagus! Terus Tingkatkan 💪" if score_pct >= 60 else "Jangan Menyerah! Mari Ulangi Flashcard 📖")
            return ft.Container(
                gradient=dark_card_gradient(),
                padding=ft.Padding.all(24),
                border_radius=ft.BorderRadius.all(18),
                border=ft.Border.all(1.5, C_SECONDARY),
                alignment=ft.Alignment.CENTER,
                content=ft.Column(
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    alignment=ft.MainAxisAlignment.CENTER,
                    spacing=12,
                    controls=[
                        ft.Icon(ft.Icons.EMOJI_EVENTS_ROUNDED, color=C_WARNING, size=60),
                        ft.Text("Sesi Tebakan Selesai!", size=20, weight=ft.FontWeight.BOLD, color=C_WHITE),
                        ft.Text(f"{score_pct:.0f}%", size=38, weight=ft.FontWeight.BOLD, color=C_SUCCESS if score_pct >= 70 else C_WARNING),
                        ft.Text(f"Jawaban Benar: {self.quiz_correct_count} dari {total_q} Soal", size=14, weight=ft.FontWeight.W_600, color=C_TEXT_LIGHT),
                        ft.Text(grade_msg, size=12, color=C_SECONDARY, text_align=ft.TextAlign.CENTER),
                        ft.Divider(color=C_MUTED),
                        ft.Row(
                            spacing=10,
                            alignment=ft.MainAxisAlignment.CENTER,
                            controls=[
                                ft.Button(
                                    content="Main Lagi 🔄",
                                    bgcolor=C_PRIMARY,
                                    color=C_WHITE,
                                    on_click=start_new_quiz,
                                ),
                                ft.Button(
                                    content="Lihat Tracker 📊",
                                    bgcolor=C_MUTED,
                                    color=C_TEXT_LIGHT,
                                    on_click=lambda e: self._jump_to(2, 0),
                                ),
                            ],
                        ),
                    ],
                ),
            )

        q_item = self.quiz_questions[self.quiz_current_idx]
        total_q = len(self.quiz_questions)
        q_desc, q_ex = parse_penjelasan(q_item.get("penjelasan", ""))

        def submit_answer(is_right: bool, selected_opt_id=None):
            if self.quiz_answered:
                return
            self.quiz_answered = True
            self.quiz_selected_opt_id = selected_opt_id
            self.quiz_is_last_correct = is_right
            self.db.record_quiz_answer(q_item["id"], is_right)
            if is_right:
                self.quiz_correct_count += 1
                self.quiz_feedback_title = f"✅ JAWABAN BENAR! — {q_item['word']} ({q_item['pos_level']})"
                self.quiz_feedback_color = C_SUCCESS
            else:
                self.quiz_feedback_title = f"❌ KURANG TEPAT! Jawaban: {q_item['word']} ({q_item['pos_level']})"
                self.quiz_feedback_color = C_DANGER
            self.quiz_feedback_desc = f"Arti: {q_item['arti']}\n{q_desc}"
            self.quiz_feedback_ex = q_ex
            self.render()

        def next_question(e):
            self.quiz_current_idx += 1
            self.quiz_answered = False
            self.quiz_selected_opt_id = None
            if self.quiz_current_idx >= len(self.quiz_questions):
                self.db.record_quiz_session(
                    self.quiz_mode,
                    self.quiz_source,
                    len(self.quiz_questions),
                    self.quiz_correct_count,
                )
            self.render()

        answer_controls = []
        if self.quiz_mode in ("pilihan_ganda", "arti_ke_inggris"):
            rng = random.Random(q_item["id"] + self.quiz_current_idx * 41)
            pool = [d for d in self.quiz_distractors if d["id"] != q_item["id"] and d["arti"] != q_item["arti"]]
            wrong_opts = rng.sample(pool, min(3, len(pool)))
            all_opts = wrong_opts + [q_item]
            rng.shuffle(all_opts)

            letters = ["A", "B", "C", "D"]
            for idx_opt, opt in enumerate(all_opts):
                is_correct_opt = opt["id"] == q_item["id"]
                is_selected_opt = opt["id"] == self.quiz_selected_opt_id
                label_txt = opt["arti"] if self.quiz_mode == "pilihan_ganda" else f"{opt['word']} ({opt['pos_level']})"

                btn_bg = C_CARD_DARK
                border_col = C_MUTED
                badge_bg = C_MUTED
                if self.quiz_answered:
                    if is_correct_opt:
                        btn_bg = "#143828"
                        border_col = C_SUCCESS
                        badge_bg = C_SUCCESS
                    elif is_selected_opt and not is_correct_opt:
                        btn_bg = "#3B1822"
                        border_col = C_DANGER
                        badge_bg = C_DANGER

                letter_str = letters[idx_opt] if idx_opt < len(letters) else "•"
                answer_controls.append(
                    ft.Container(
                        bgcolor=btn_bg,
                        padding=ft.Padding.symmetric(horizontal=12, vertical=11),
                        border_radius=ft.BorderRadius.all(12),
                        border=ft.Border.all(1.5 if (self.quiz_answered and (is_correct_opt or is_selected_opt)) else 1, border_col),
                        on_click=lambda e, right=is_correct_opt, oid=opt["id"]: submit_answer(right, oid),
                        content=ft.Row(
                            spacing=10,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            controls=[
                                ft.Container(
                                    width=26,
                                    height=26,
                                    bgcolor=badge_bg,
                                    border_radius=ft.BorderRadius.all(8),
                                    alignment=ft.Alignment.CENTER,
                                    content=ft.Text(
                                        letter_str,
                                        size=11.5,
                                        weight=ft.FontWeight.BOLD,
                                        color=C_BG_DEEP if (self.quiz_answered and is_correct_opt) else C_WHITE,
                                    ),
                                ),
                                ft.Text(label_txt, size=13, color=C_WHITE, weight=ft.FontWeight.W_500, expand=True),
                            ],
                        ),
                    )
                )
        else:
            target_clean = q_item["word"].split("(")[0].strip()
            tf_guess = ft.TextField(
                label=f"Ketik kata Inggris (Awalan: '{target_clean[0]}', {len(target_clean)} huruf)",
                color=C_WHITE,
                bgcolor=C_MUTED,
                border_color=C_SECONDARY,
                border_radius=ft.BorderRadius.all(10),
                disabled=self.quiz_answered,
            )

            def check_typed(e):
                typed = (tf_guess.value or "").strip().lower()
                if not typed:
                    self.notify("Ketik jawaban Anda terlebih dahulu!", C_WARNING)
                    return
                is_right = typed == target_clean.lower() or typed == q_item["word"].strip().lower()
                submit_answer(is_right, -1)

            tf_guess.on_submit = check_typed

            answer_controls.extend([
                tf_guess,
                ft.Button(
                    content="Periksa Jawaban ✔️",
                    bgcolor=C_PRIMARY,
                    color=C_WHITE,
                    disabled=self.quiz_answered,
                    on_click=check_typed,
                ),
            ])

        prompt_title = (
            f"Apa arti Bahasa Indonesia dari kata:\n\"{q_item['word']}\"?"
            if self.quiz_mode == "pilihan_ganda"
            else f"Kata Bahasa Inggris manakah yang memiliki arti:\n\"{q_item['arti']}\"?"
        )

        return ft.ListView(
            expand=True,
            spacing=10,
            padding=ft.Padding.only(bottom=24),
            controls=[
                ft.Row(
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    controls=[
                        ft.Text(f"Soal {self.quiz_current_idx + 1} / {total_q}", size=13, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                        ft.Container(
                            bgcolor=C_CARD_DARK,
                            padding=ft.Padding.symmetric(horizontal=10, vertical=4),
                            border_radius=ft.BorderRadius.all(10),
                            border=ft.Border.all(1, C_SUCCESS),
                            content=ft.Text(f"Skor: {self.quiz_correct_count} Benar", size=11.5, weight=ft.FontWeight.BOLD, color=C_SUCCESS),
                        ),
                        ft.TextButton(
                            content="Akhiri Kuis ✖",
                            on_click=lambda e: self._stop_quiz(),
                        ),
                    ],
                ),
                ft.ProgressBar(
                    value=(self.quiz_current_idx + 1) / max(1, total_q),
                    bar_height=6,
                    color=C_PRIMARY,
                    bgcolor=C_MUTED,
                    border_radius=ft.BorderRadius.all(4),
                ),
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(16),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_PRIMARY),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text("PERTANYAAN KUIS", size=10.5, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                                    ft.Container(
                                        bgcolor=C_MUTED,
                                        padding=ft.Padding.symmetric(horizontal=8, vertical=2),
                                        border_radius=ft.BorderRadius.all(6),
                                        content=ft.Text(f"Kelas & Level: {q_item['pos_level']}", size=10.5, color=C_TEXT_LIGHT),
                                    ),
                                ],
                            ),
                            ft.Text(prompt_title, size=15.5, weight=ft.FontWeight.BOLD, color=C_WHITE),
                        ],
                    ),
                ),
                ft.Column(spacing=8, controls=answer_controls),
                ft.Container(
                    visible=self.quiz_answered,
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1.5, self.quiz_feedback_color),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Text(self.quiz_feedback_title, size=13.5, weight=ft.FontWeight.BOLD, color=self.quiz_feedback_color),
                            ft.Text(self.quiz_feedback_desc, size=12, color=C_TEXT_LIGHT),
                            *(
                                [
                                    ft.Container(
                                        bgcolor=C_BG_DEEP,
                                        padding=ft.Padding.symmetric(horizontal=10, vertical=6),
                                        border_radius=ft.BorderRadius.all(8),
                                        border=ft.Border.all(1, C_MUTED),
                                        content=ft.Text(f'💬 Contoh: "{self.quiz_feedback_ex}"', size=11, italic=True, color=C_SECONDARY),
                                    )
                                ]
                                if self.quiz_feedback_ex
                                else []
                            ),
                            ft.Button(
                                content="Soal Berikutnya ▶" if self.quiz_current_idx + 1 < total_q else "Lihat Hasil Akhir 🏆",
                                bgcolor=C_PRIMARY,
                                color=C_WHITE,
                                on_click=next_question,
                            ),
                        ],
                    ),
                ),
            ],
        )

    def _choice_chip(self, text: str, active: bool, on_tap) -> ft.Container:
        return ft.Container(
            bgcolor=C_PRIMARY if active else C_MUTED,
            padding=ft.Padding.symmetric(horizontal=12, vertical=8),
            border_radius=ft.BorderRadius.all(10),
            border=ft.Border.all(1, C_SECONDARY if active else C_MUTED),
            on_click=lambda e: on_tap(),
            content=ft.Text(
                text,
                size=11.5,
                color=C_WHITE if active else C_TEXT_LIGHT,
                weight=ft.FontWeight.BOLD if active else ft.FontWeight.NORMAL,
            ),
        )

    def _set_quiz_opt(self, source: str = None, mode: str = None):
        if source:
            self.quiz_source = source
        if mode:
            self.quiz_mode = mode
        self.render()

    def _stop_quiz(self):
        self.quiz_active = False
        self.render()

    # =========================================================================
    # TAB 2: TRACKER HAFALAN (PER BULAN & PER TAHUN)
    # =========================================================================
    def build_tracker_view(self) -> ft.Control:
        def switch_sub(idx: int):
            self.tracker_subtab = idx
            self.render()

        sub_bar = self.build_subtab_bar(
            ["Tracker Bulanan", "Tracker Tahunan"],
            self.tracker_subtab,
            switch_sub,
        )

        if self.tracker_subtab == 0:
            content = self.build_vocab_monthly_tracker_subview()
        else:
            content = self.build_vocab_yearly_tracker_subview()

        return ft.Column(expand=True, spacing=6, controls=[sub_bar, content])

    def build_vocab_monthly_tracker_subview(self) -> ft.Control:
        m_data = self.db.get_vocab_monthly_tracker(self.selected_year, self.selected_month)
        pct_target = (m_data["memorized_count"] / max(1, m_data["monthly_target"])) * 100.0

        def edit_monthly_target(e):
            tf_t = ft.TextField(
                label="Target Hafalan Kosakata per Bulan",
                value=str(m_data["monthly_target"]),
                keyboard_type=ft.KeyboardType.NUMBER,
                color=C_WHITE,
                bgcolor=C_MUTED,
            )

            def save_t(ev):
                try:
                    val = max(1, int(tf_t.value))
                    self.db.set_setting("vocab_monthly_target", str(val))
                    self.page.pop_dialog()
                    self.notify(f"Target bulanan diubah menjadi {val} kata/bulan!", C_PRIMARY)
                    self.render()
                except ValueError:
                    self.notify("Masukkan angka yang valid!", C_DANGER)

            self.page.show_dialog(
                ft.AlertDialog(
                    bgcolor=C_CARD_DARK,
                    title=ft.Text("Atur Target Hafalan Bulanan", color=C_TEXT_LIGHT),
                    content=tf_t,
                    actions=[
                        ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                        ft.Button(content="Simpan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save_t),
                    ],
                )
            )

        daily_rows = []
        if m_data["daily_memorized"]:
            max_daily = max(list(m_data["daily_memorized"].values()) + [5])
            for dt_str, cnt in m_data["daily_memorized"].items():
                daily_rows.append(
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Text(dt_str, size=11.5, color=C_TEXT_LIGHT, width=85),
                            ft.Container(
                                expand=True,
                                margin=ft.Margin.symmetric(horizontal=8),
                                content=ft.ProgressBar(
                                    value=min(1.0, cnt / max_daily),
                                    bar_height=8,
                                    color=C_SECONDARY,
                                    bgcolor=C_MUTED,
                                    border_radius=ft.BorderRadius.all(4),
                                ),
                            ),
                            ft.Text(f"+{cnt} kata", size=11.5, weight=ft.FontWeight.BOLD, color=C_SUCCESS),
                        ],
                    )
                )
        else:
            daily_rows.append(ft.Text("Belum ada kata yang ditandai hafal pada bulan ini.", size=12, color=C_SECONDARY))

        word_items = []
        for w in m_data["words_list"]:
            word_items.append(
                ft.Container(
                    bgcolor=C_BG_DEEP,
                    padding=ft.Padding.symmetric(horizontal=10, vertical=8),
                    border_radius=ft.BorderRadius.all(8),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Column(
                                spacing=2,
                                expand=True,
                                controls=[
                                    ft.Text(f"{w['word']} ({w['pos_level']})", size=13, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                    ft.Text(w["arti"], size=11, color=C_TEXT_LIGHT),
                                ],
                            ),
                            ft.Container(
                                bgcolor=C_CARD_DARK,
                                padding=ft.Padding.symmetric(horizontal=7, vertical=3),
                                border_radius=ft.BorderRadius.all(6),
                                content=ft.Text(w["memorized_date"] or "", size=10, color=C_SECONDARY),
                            ),
                        ],
                    ),
                )
            )

        return ft.ListView(
            expand=True,
            spacing=10,
            padding=ft.Padding.only(bottom=24),
            controls=[
                self.build_period_selector(show_month=True),
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(15),
                    border=ft.Border.all(1, C_PRIMARY),
                    shadow=card_shadow(),
                    content=ft.Column(
                        spacing=10,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Text(
                                        f"Tracker {MONTH_NAMES_ID[self.selected_month]} {self.selected_year}",
                                        size=14.5,
                                        weight=ft.FontWeight.BOLD,
                                        color=C_WHITE,
                                    ),
                                    ft.Container(
                                        bgcolor=C_MUTED,
                                        padding=ft.Padding.symmetric(horizontal=10, vertical=4),
                                        border_radius=ft.BorderRadius.all(8),
                                        on_click=edit_monthly_target,
                                        content=ft.Text(f"Target: {m_data['monthly_target']} ✏️", size=11, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                                    ),
                                ],
                            ),
                            ft.Text(
                                f"{m_data['memorized_count']} / {m_data['monthly_target']} Kata Dihafal ({pct_target:.1f}%)",
                                size=18,
                                weight=ft.FontWeight.BOLD,
                                color=C_SECONDARY,
                            ),
                            ft.ProgressBar(
                                value=min(1.0, pct_target / 100.0),
                                bar_height=10,
                                color=C_PRIMARY,
                                bgcolor=C_MUTED,
                                border_radius=ft.BorderRadius.all(6),
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._stat_mini_box("Sesi Tebakan", f"{m_data['quiz_sessions']} Sesi", C_TEXT_LIGHT, ft.Icons.QUIZ_OUTLINED),
                                    self._stat_mini_box("Soal Dijawab", f"{m_data['quiz_total_q']} Soal", C_SECONDARY, ft.Icons.LIST_ALT_ROUNDED),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._stat_mini_box("Jawaban Benar", f"{m_data['quiz_correct_q']} Benar", C_SUCCESS, ft.Icons.CHECK_CIRCLE_OUTLINE),
                                    self._stat_mini_box("Akurasi Kuis", f"{m_data['quiz_avg_pct']}%", C_WARNING, ft.Icons.INSIGHTS_ROUNDED),
                                ],
                            ),
                        ],
                    ),
                ),
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Text("📅 Grafik Aktivitas Hafalan Harian", size=13.5, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                            ft.Column(spacing=6, controls=daily_rows),
                        ],
                    ),
                ),
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Text(f"✅ Kosakata Dihafal Bulan Ini ({len(word_items)})", size=13.5, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                            ft.Column(
                                spacing=6,
                                controls=word_items if word_items else [ft.Text("Belum ada kata yang dihafal pada periode bulan ini.", size=12, color=C_SECONDARY)],
                            ),
                        ],
                    ),
                ),
            ],
        )

    def build_vocab_yearly_tracker_subview(self) -> ft.Control:
        y_data = self.db.get_vocab_yearly_tracker(self.selected_year)
        max_m = max([m["memorized"] for m in y_data["months"]] + [10])

        # Visual 12-Month Vertical Bar Chart
        bar_columns = []
        for m_info in y_data["months"]:
            m_idx = m_info["month"]
            cnt = m_info["memorized"]
            bar_h = max(6, int(round((cnt / max_m) * 75))) if cnt > 0 else 4
            is_cur_m = m_idx == self.selected_month
            bar_columns.append(
                ft.Column(
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    alignment=ft.MainAxisAlignment.END,
                    spacing=3,
                    expand=True,
                    controls=[
                        ft.Text(str(cnt) if cnt > 0 else "", size=8.5, color=C_SUCCESS, weight=ft.FontWeight.BOLD),
                        ft.Container(
                            width=14,
                            height=bar_h,
                            bgcolor=C_SUCCESS if cnt > 0 else (C_SECONDARY if is_cur_m else C_MUTED),
                            border_radius=ft.BorderRadius.all(4),
                            tooltip=f"{MONTH_NAMES_ID[m_idx]}: {cnt} Kata",
                        ),
                        ft.Text(
                            MONTH_SHORT_ID[m_idx],
                            size=9,
                            weight=ft.FontWeight.BOLD if is_cur_m else ft.FontWeight.NORMAL,
                            color=C_WHITE if is_cur_m else C_SECONDARY,
                        ),
                    ],
                )
            )

        month_rows = []
        for m_info in y_data["months"]:
            m_idx = m_info["month"]
            cnt = m_info["memorized"]
            q_cnt = m_info["quiz_sessions"]
            avg_s = m_info["avg_score"]
            month_rows.append(
                ft.Container(
                    bgcolor=C_BG_DEEP,
                    padding=ft.Padding.symmetric(horizontal=10, vertical=8),
                    border_radius=ft.BorderRadius.all(8),
                    border=ft.Border.all(1, C_PRIMARY if m_idx == self.selected_month else C_MUTED),
                    on_click=lambda e, mi=m_idx: (setattr(self, "selected_month", mi), setattr(self, "tracker_subtab", 0), self.render()),
                    content=ft.Column(
                        spacing=4,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text(MONTH_NAMES_ID[m_idx], size=12, weight=ft.FontWeight.BOLD, color=C_WHITE, width=90),
                                    ft.Text(f"{cnt} Kata Hafal", size=12, weight=ft.FontWeight.BOLD, color=C_SUCCESS if cnt > 0 else C_TEXT_LIGHT),
                                    ft.Text(f"{q_cnt} Kuis ({avg_s}%)", size=11, color=C_SECONDARY),
                                ],
                            ),
                            ft.ProgressBar(
                                value=min(1.0, cnt / max_m),
                                bar_height=5,
                                color=C_PRIMARY if cnt > 0 else C_MUTED,
                                bgcolor=C_MUTED,
                                border_radius=ft.BorderRadius.all(3),
                            ),
                        ],
                    ),
                )
            )

        b2_pct = (y_data["b2_hafal"] / max(1, y_data["b2_total"])) * 100.0
        c1_pct = (y_data["c1_hafal"] / max(1, y_data["c1_total"])) * 100.0

        return ft.ListView(
            expand=True,
            spacing=10,
            padding=ft.Padding.only(bottom=24),
            controls=[
                self.build_period_selector(show_month=False),
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(15),
                    border=ft.Border.all(1, C_PRIMARY),
                    shadow=card_shadow(),
                    content=ft.Column(
                        spacing=10,
                        controls=[
                            ft.Text(f"🏆 Ringkasan Hafalan Tahunan ({self.selected_year})", size=14.5, weight=ft.FontWeight.BOLD, color=C_WHITE),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._stat_mini_box("Total Hafal Tahun Ini", f"{y_data['total_year_memorized']} Kata", C_SUCCESS, ft.Icons.WORKSPACE_PREMIUM_ROUNDED),
                                    self._stat_mini_box("Total Sesi Kuis", f"{y_data['total_year_quizzes']} Sesi", C_SECONDARY, ft.Icons.QUIZ_ROUNDED),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._stat_mini_box("Penguasaan Level B2", f"{y_data['b2_hafal']}/{y_data['b2_total']} ({b2_pct:.1f}%)", C_TEXT_LIGHT),
                                    self._stat_mini_box("Penguasaan Level C1", f"{y_data['c1_hafal']}/{y_data['c1_total']} ({c1_pct:.1f}%)", C_WARNING),
                                ],
                            ),
                            ft.Container(
                                bgcolor=C_BG_DEEP,
                                padding=ft.Padding.only(left=8, right=8, top=12, bottom=8),
                                border_radius=ft.BorderRadius.all(12),
                                border=ft.Border.all(1, C_MUTED),
                                height=120,
                                content=ft.Row(
                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    vertical_alignment=ft.CrossAxisAlignment.END,
                                    controls=bar_columns,
                                ),
                            ),
                        ],
                    ),
                ),
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Text("📊 Rincian Hafalan & Tebakan 12 Bulan (Ketuk Bulan untuk Detail)", size=13, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                            ft.Column(spacing=6, controls=month_rows),
                        ],
                    ),
                ),
            ],
        )

    # =========================================================================
    # TAB 3: KEUANGAN (TRANSAKSI, OVERVIEW BULANAN & OVERVIEW TAHUNAN)
    # =========================================================================
    def build_finance_view(self) -> ft.Control:
        def switch_sub(idx: int):
            self.finance_subtab = idx
            self.render()

        sub_bar = self.build_subtab_bar(
            ["Transaksi", "Overview Bulanan", "Overview Tahunan"],
            self.finance_subtab,
            switch_sub,
        )

        if self.finance_subtab == 0:
            content = self.build_transactions_subview()
        elif self.finance_subtab == 1:
            content = self.build_monthly_finance_subview()
        else:
            content = self.build_yearly_finance_subview()

        return ft.Column(expand=True, spacing=6, controls=[sub_bar, content])

    def build_transactions_subview(self) -> ft.Control:
        tx_list = self.db.get_transactions(
            year=self.selected_year,
            month=self.selected_month,
            tipe_filter=self.tx_tipe_filter,
        )
        all_cats = self.db.get_finance_categories()

        def set_tx_filter(val: str):
            self.tx_tipe_filter = val
            self.render()

        def open_add_tx_dialog(e):
            today_str = datetime.date.today().strftime("%Y-%m-%d")
            tf_date = ft.TextField(label="Tanggal (YYYY-MM-DD)", value=today_str, color=C_WHITE, bgcolor=C_MUTED)

            dd_tipe = ft.Dropdown(
                label="Tipe Transaksi",
                value="KEBUTUHAN",
                color=C_WHITE,
                bgcolor=C_MUTED,
                options=[
                    ft.dropdown.Option(key="PENDAPATAN", text="PENDAPATAN"),
                    ft.dropdown.Option(key="KEBUTUHAN", text="KEBUTUHAN (Needs)"),
                    ft.dropdown.Option(key="KEINGINAN", text="KEINGINAN (Wants)"),
                    ft.dropdown.Option(key="TABUNGAN & INVESTASI", text="TABUNGAN & INVESTASI"),
                ],
            )

            def get_cat_opts(tipe_val: str):
                filtered = [c["kategori"] for c in all_cats if c["tipe"] == tipe_val]
                if not filtered:
                    filtered = ["Umum"]
                return [ft.dropdown.Option(key=k, text=k) for k in filtered]

            dd_kat = ft.Dropdown(
                label="Kategori (Sesuai SET-UP)",
                value=get_cat_opts("KEBUTUHAN")[0].key,
                color=C_WHITE,
                bgcolor=C_MUTED,
                options=get_cat_opts("KEBUTUHAN"),
            )

            def on_tipe_change(ev):
                opts = get_cat_opts(dd_tipe.value)
                dd_kat.options = opts
                dd_kat.value = opts[0].key
                self.page.update()

            dd_tipe.on_select = on_tipe_change

            tf_desc = ft.TextField(label="Deskripsi (misal: Makan siang / Gaji)", color=C_WHITE, bgcolor=C_MUTED)
            tf_total = ft.TextField(label="Nominal Rp (misal: 150000)", keyboard_type=ft.KeyboardType.NUMBER, color=C_WHITE, bgcolor=C_MUTED)
            tf_notes = ft.TextField(label="Notes (Opsional)", color=C_WHITE, bgcolor=C_MUTED)

            def save_tx(ev):
                try:
                    nom = float((tf_total.value or "0").replace(".", "").replace(",", "").strip())
                    if nom <= 0:
                        raise ValueError()
                except ValueError:
                    self.notify("Masukkan nominal Rupiah yang valid!", C_DANGER)
                    return

                self.db.add_transaction(
                    tanggal=tf_date.value.strip() or today_str,
                    tipe=dd_tipe.value,
                    kategori=dd_kat.value or "Umum",
                    deskripsi=tf_desc.value or dd_kat.value,
                    total=nom,
                    notes=tf_notes.value or "",
                )
                self.page.pop_dialog()
                self.notify("Transaksi berhasil disimpan!", C_PRIMARY)
                self.render()

            self.page.show_dialog(
                ft.AlertDialog(
                    bgcolor=C_CARD_DARK,
                    title=ft.Text("Tambah Transaksi Keuangan", color=C_WHITE, weight=ft.FontWeight.BOLD),
                    content=ft.Column(tight=True, spacing=8, controls=[tf_date, dd_tipe, dd_kat, tf_desc, tf_total, tf_notes]),
                    actions=[
                        ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                        ft.Button(content="Simpan Transaksi", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save_tx),
                    ],
                )
            )

        def open_edit_tx_dialog(tx: dict):
            tf_date = ft.TextField(label="Tanggal (YYYY-MM-DD)", value=tx["tanggal"], color=C_WHITE, bgcolor=C_MUTED)

            dd_tipe = ft.Dropdown(
                label="Tipe Transaksi",
                value=tx["tipe"],
                color=C_WHITE,
                bgcolor=C_MUTED,
                options=[
                    ft.dropdown.Option(key="PENDAPATAN", text="PENDAPATAN"),
                    ft.dropdown.Option(key="KEBUTUHAN", text="KEBUTUHAN (Needs)"),
                    ft.dropdown.Option(key="KEINGINAN", text="KEINGINAN (Wants)"),
                    ft.dropdown.Option(key="TABUNGAN & INVESTASI", text="TABUNGAN & INVESTASI"),
                ],
            )

            def get_cat_opts(tipe_val: str):
                filtered = [c["kategori"] for c in all_cats if c["tipe"] == tipe_val]
                if tx["kategori"] and tx["kategori"] not in filtered:
                    filtered.append(tx["kategori"])
                if not filtered:
                    filtered = ["Umum"]
                return [ft.dropdown.Option(key=k, text=k) for k in filtered]

            curr_cat_opts = get_cat_opts(tx["tipe"])
            cat_val = tx["kategori"] if any(o.key == tx["kategori"] for o in curr_cat_opts) else curr_cat_opts[0].key

            dd_kat = ft.Dropdown(
                label="Kategori (Sesuai SET-UP)",
                value=cat_val,
                color=C_WHITE,
                bgcolor=C_MUTED,
                options=curr_cat_opts,
            )

            def on_tipe_change(ev):
                opts = get_cat_opts(dd_tipe.value)
                dd_kat.options = opts
                dd_kat.value = opts[0].key
                self.page.update()

            dd_tipe.on_select = on_tipe_change

            tf_desc = ft.TextField(label="Deskripsi", value=tx["deskripsi"] or "", color=C_WHITE, bgcolor=C_MUTED)
            tf_total = ft.TextField(
                label="Nominal Rp",
                value=str(int(round(float(tx["total"])))),
                keyboard_type=ft.KeyboardType.NUMBER,
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            tf_notes = ft.TextField(label="Notes (Opsional)", value=tx["notes"] or "", color=C_WHITE, bgcolor=C_MUTED)

            def save_updated_tx(ev):
                try:
                    nom = float((tf_total.value or "0").replace(".", "").replace(",", "").strip())
                    if nom <= 0:
                        raise ValueError()
                except ValueError:
                    self.notify("Masukkan nominal Rupiah yang valid!", C_DANGER)
                    return

                self.db.update_transaction(
                    tx_id=tx["id"],
                    tanggal=tf_date.value.strip() or tx["tanggal"],
                    tipe=dd_tipe.value,
                    kategori=dd_kat.value or "Umum",
                    deskripsi=tf_desc.value or dd_kat.value,
                    total=nom,
                    notes=tf_notes.value or "",
                )
                self.page.pop_dialog()
                self.notify("Transaksi berhasil diperbarui!", C_SUCCESS)
                self.render()

            self.page.show_dialog(
                ft.AlertDialog(
                    bgcolor=C_CARD_DARK,
                    title=ft.Text("Edit Transaksi Keuangan", color=C_WHITE, weight=ft.FontWeight.BOLD),
                    content=ft.Column(tight=True, spacing=8, controls=[tf_date, dd_tipe, dd_kat, tf_desc, tf_total, tf_notes]),
                    actions=[
                        ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                        ft.Button(content="Simpan Perubahan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save_updated_tx),
                    ],
                )
            )

        def delete_tx(tx_id: int):
            self.db.delete_transaction(tx_id)
            self.notify("Transaksi dihapus.", C_MUTED)
            self.render()

        filter_chips = []
        for code, lbl in [
            ("SEMUA", "Semua"),
            ("PENDAPATAN", "Pendapatan"),
            ("KEBUTUHAN", "Kebutuhan"),
            ("KEINGINAN", "Keinginan"),
            ("TABUNGAN & INVESTASI", "Tabungan"),
        ]:
            active = self.tx_tipe_filter == code
            filter_chips.append(
                ft.Container(
                    bgcolor=C_PRIMARY if active else C_MUTED,
                    padding=ft.Padding.symmetric(horizontal=9, vertical=4),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_SECONDARY if active else C_MUTED),
                    on_click=lambda e, c=code: set_tx_filter(c),
                    content=ft.Text(lbl, size=10.5, color=C_WHITE if active else C_TEXT_LIGHT, weight=ft.FontWeight.BOLD if active else ft.FontWeight.NORMAL),
                )
            )

        tx_cards = []
        for tx in tx_list:
            t_type = tx["tipe"]
            is_income = t_type == "PENDAPATAN"
            amt_color = C_SUCCESS if is_income else (C_SECONDARY if t_type == "TABUNGAN & INVESTASI" else C_WARNING)
            sign_str = "+" if is_income else "-"

            tx_cards.append(
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(12),
                    border_radius=ft.BorderRadius.all(12),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Container(
                                expand=True,
                                on_click=lambda e, t=tx: open_edit_tx_dialog(t),
                                content=ft.Column(
                                    spacing=3,
                                    controls=[
                                        ft.Row(
                                            spacing=5,
                                            wrap=True,
                                            controls=[
                                                ft.Container(
                                                    bgcolor=C_PRIMARY if is_income else C_MUTED,
                                                    padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                                                    border_radius=ft.BorderRadius.all(6),
                                                    content=ft.Text(t_type, size=8.5, color=C_WHITE, weight=ft.FontWeight.BOLD),
                                                ),
                                                ft.Text(tx["kategori"], size=12, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                                                ft.Text(f"• {tx['tanggal']}", size=10, color=C_TEXT_LIGHT),
                                            ],
                                        ),
                                        ft.Text(tx["deskripsi"] or "-", size=13, weight=ft.FontWeight.W_500, color=C_WHITE),
                                        *(
                                            [ft.Text(tx["notes"], size=10.5, color=C_TEXT_LIGHT, italic=True)]
                                            if tx.get("notes")
                                            else []
                                        ),
                                    ],
                                ),
                            ),
                            ft.Column(
                                horizontal_alignment=ft.CrossAxisAlignment.END,
                                spacing=2,
                                controls=[
                                    ft.Text(f"{sign_str}{format_rp(tx['total'])}", size=13, weight=ft.FontWeight.BOLD, color=amt_color),
                                    ft.Row(
                                        spacing=2,
                                        controls=[
                                            ft.IconButton(
                                                icon=ft.Icons.EDIT_OUTLINED,
                                                icon_color=C_SECONDARY,
                                                icon_size=16,
                                                width=30,
                                                height=30,
                                                style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                                tooltip="Edit Transaksi",
                                                on_click=lambda e, t=tx: open_edit_tx_dialog(t),
                                            ),
                                            ft.IconButton(
                                                icon=ft.Icons.DELETE_OUTLINE,
                                                icon_color=C_DANGER,
                                                icon_size=16,
                                                width=30,
                                                height=30,
                                                style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                                tooltip="Hapus Transaksi",
                                                on_click=lambda e, tid=tx["id"]: delete_tx(tid),
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),
                )
            )

        return ft.Column(
            expand=True,
            spacing=8,
            controls=[
                self.build_period_selector(show_month=True),
                ft.Row(spacing=5, wrap=True, controls=filter_chips),
                ft.Row(
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    controls=[
                        ft.Text(f"Riwayat Transaksi ({len(tx_cards)})", size=13.5, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                        ft.Button(
                            content="+ Tambah Transaksi",
                            bgcolor=C_PRIMARY,
                            color=C_WHITE,
                            on_click=open_add_tx_dialog,
                        ),
                    ],
                ),
                ft.ListView(
                    expand=True,
                    spacing=8,
                    padding=ft.Padding.only(bottom=24),
                    controls=tx_cards if tx_cards else [
                        ft.Container(
                            bgcolor=C_CARD_DARK,
                            padding=ft.Padding.all(20),
                            border_radius=ft.BorderRadius.all(12),
                            border=ft.Border.all(1, C_MUTED),
                            alignment=ft.Alignment.CENTER,
                            content=ft.Text("Belum ada transaksi pada filter/periode ini. Klik '+ Tambah Transaksi'.", color=C_SECONDARY, text_align=ft.TextAlign.CENTER),
                        )
                    ],
                ),
            ],
        )

    def build_monthly_finance_subview(self) -> ft.Control:
        ov = self.db.get_monthly_finance_overview(self.selected_year, self.selected_month)
        rule = ov["rule"]

        income_val = ov["pendapatan"] if ov["pendapatan"] > 0 else ov["target_by_tipe"].get("PENDAPATAN", 0.0)
        base_income = max(1.0, income_val)
        used_pct = (ov["total_alokasi"] / base_income * 100.0) if base_income > 0 else 0.0
        remain_pct = max(0.0, 100.0 - used_pct)

        pct_keb_act = (ov["kebutuhan"] / base_income) * 100.0
        pct_kei_act = (ov["keinginan"] / base_income) * 100.0
        pct_tab_act = (ov["tabungan"] / base_income) * 100.0
        pct_rem_act = max(0.0, 100.0 - (pct_keb_act + pct_kei_act + pct_tab_act))

        if ov["selisih"] < 0:
            health_color = C_DANGER
            health_title = "🚨 OVERBUDGET (DEFISIT)"
            health_desc = f"Pengeluaran melebihi pendapatan sebesar {format_rp(abs(ov['selisih']))}!"
        elif used_pct >= 90:
            health_color = C_WARNING
            health_title = "⚠️ WASPADA BATAS LIMIT"
            health_desc = f"Budget terpakai {used_pct:.1f}%. Sisa uang bebas menipis ({format_rp(ov['selisih'])})."
        else:
            health_color = C_SUCCESS
            health_title = "✅ KEUANGAN SEHAT & SURPLUS"
            health_desc = f"Arus kas positif! Sisa dana bebas tersimpan {format_rp(ov['selisih'])} ({remain_pct:.1f}%)."

        segments = []
        if pct_keb_act > 0:
            segments.append(
                ft.Container(
                    expand=max(1, int(round(pct_keb_act * 10))),
                    bgcolor=C_SECONDARY,
                    height=22,
                    tooltip=f"Kebutuhan: {pct_keb_act:.1f}% ({format_rp(ov['kebutuhan'])})",
                    alignment=ft.Alignment.CENTER,
                    content=ft.Text(f"{pct_keb_act:.0f}%", size=9.5, weight=ft.FontWeight.BOLD, color=C_WHITE) if pct_keb_act >= 8 else None,
                )
            )
        if pct_kei_act > 0:
            segments.append(
                ft.Container(
                    expand=max(1, int(round(pct_kei_act * 10))),
                    bgcolor=C_WARNING,
                    height=22,
                    tooltip=f"Keinginan: {pct_kei_act:.1f}% ({format_rp(ov['keinginan'])})",
                    alignment=ft.Alignment.CENTER,
                    content=ft.Text(f"{pct_kei_act:.0f}%", size=9.5, weight=ft.FontWeight.BOLD, color=C_BG_DEEP) if pct_kei_act >= 8 else None,
                )
            )
        if pct_tab_act > 0:
            segments.append(
                ft.Container(
                    expand=max(1, int(round(pct_tab_act * 10))),
                    bgcolor=C_PRIMARY,
                    height=22,
                    tooltip=f"Tabungan: {pct_tab_act:.1f}% ({format_rp(ov['tabungan'])})",
                    alignment=ft.Alignment.CENTER,
                    content=ft.Text(f"{pct_tab_act:.0f}%", size=9.5, weight=ft.FontWeight.BOLD, color=C_WHITE) if pct_tab_act >= 8 else None,
                )
            )
        if ov["selisih"] >= 0 and pct_rem_act > 0:
            segments.append(
                ft.Container(
                    expand=max(1, int(round(pct_rem_act * 10))),
                    bgcolor=C_SUCCESS,
                    height=22,
                    tooltip=f"Sisa Bebas: {pct_rem_act:.1f}% ({format_rp(ov['selisih'])})",
                    alignment=ft.Alignment.CENTER,
                    content=ft.Text(f"{pct_rem_act:.0f}%", size=9.5, weight=ft.FontWeight.BOLD, color=C_BG_DEEP) if pct_rem_act >= 8 else None,
                )
            )
        elif ov["selisih"] < 0:
            deficit_pct = (abs(ov["selisih"]) / base_income) * 100.0
            segments.append(
                ft.Container(
                    expand=max(1, int(round(deficit_pct * 10))),
                    bgcolor=C_DANGER,
                    height=22,
                    tooltip=f"Defisit: {deficit_pct:.1f}% ({format_rp(abs(ov['selisih']))})",
                    alignment=ft.Alignment.CENTER,
                    content=ft.Text(f"Defisit {deficit_pct:.0f}%", size=9, weight=ft.FontWeight.BOLD, color=C_WHITE),
                )
            )

        if not segments:
            segments.append(
                ft.Container(
                    expand=True,
                    bgcolor=C_MUTED,
                    height=22,
                    alignment=ft.Alignment.CENTER,
                    content=ft.Text("Belum ada alokasi transaksi", size=10, color=C_TEXT_LIGHT),
                )
            )

        # Top 5 Expense Categories ranking
        all_expenses = []
        for t_k in ["KEBUTUHAN", "KEINGINAN"]:
            for item in ov["breakdown_by_tipe"].get(t_k, []):
                if item["actual"] > 0:
                    all_expenses.append({
                        "kategori": item["kategori"],
                        "tipe": t_k,
                        "actual": item["actual"],
                    })
        all_expenses.sort(key=lambda x: x["actual"], reverse=True)
        tot_exp_only = sum(x["actual"] for x in all_expenses)

        top_exp_rows = []
        for rank, exp in enumerate(all_expenses[:5], start=1):
            exp_pct = (exp["actual"] / tot_exp_only * 100.0) if tot_exp_only > 0 else 0.0
            bar_color = C_SECONDARY if exp["tipe"] == "KEBUTUHAN" else C_WARNING
            top_exp_rows.append(
                ft.Container(
                    bgcolor=C_BG_DEEP,
                    padding=ft.Padding.symmetric(horizontal=10, vertical=7),
                    border_radius=ft.BorderRadius.all(8),
                    content=ft.Column(
                        spacing=4,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Row(
                                        spacing=6,
                                        controls=[
                                            ft.Container(
                                                bgcolor=C_MUTED,
                                                padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                                                border_radius=ft.BorderRadius.all(4),
                                                content=ft.Text(f"#{rank}", size=10, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                                            ),
                                            ft.Text(exp["kategori"], size=12, weight=ft.FontWeight.W_600, color=C_WHITE),
                                            ft.Container(
                                                bgcolor=C_PRIMARY if exp["tipe"] == "KEBUTUHAN" else C_MUTED,
                                                padding=ft.Padding.symmetric(horizontal=5, vertical=1),
                                                border_radius=ft.BorderRadius.all(4),
                                                content=ft.Text(exp["tipe"], size=8, color=C_WHITE, weight=ft.FontWeight.BOLD),
                                            ),
                                        ],
                                    ),
                                    ft.Text(f"{format_rp(exp['actual'])} ({exp_pct:.1f}%)", size=11, weight=ft.FontWeight.BOLD, color=bar_color),
                                ],
                            ),
                            ft.ProgressBar(
                                value=min(1.0, exp_pct / 100.0),
                                bar_height=5,
                                color=bar_color,
                                bgcolor=C_MUTED,
                                border_radius=ft.BorderRadius.all(3),
                            ),
                        ],
                    ),
                )
            )

        def build_rule_row(label: str, target_pct: float, actual_pct: float, actual_rp: float, color: str) -> ft.Container:
            over_budget = label in ("KEBUTUHAN", "KEINGINAN") and actual_pct > target_pct
            diff_rp = (target_pct / 100.0 * base_income) - actual_rp
            if label in ("KEBUTUHAN", "KEINGINAN"):
                if diff_rp >= 0:
                    status_badge = f"Hemat {format_rp_compact(diff_rp)} ✅"
                    status_col = C_SUCCESS
                else:
                    status_badge = f"Over {format_rp_compact(abs(diff_rp))} 🚨"
                    status_col = C_DANGER
            else:
                if actual_rp >= (target_pct / 100.0 * base_income):
                    status_badge = "Target Tercapai 🌟"
                    status_col = C_SUCCESS
                else:
                    status_badge = f"Kurang {format_rp_compact(abs(diff_rp))} ⏳"
                    status_col = C_SECONDARY

            return ft.Container(
                bgcolor=C_BG_DEEP,
                padding=ft.Padding.all(10),
                border_radius=ft.BorderRadius.all(10),
                border=ft.Border.all(1, C_DANGER if over_budget else C_MUTED),
                content=ft.Column(
                    spacing=5,
                    controls=[
                        ft.Row(
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            controls=[
                                ft.Row(
                                    spacing=6,
                                    controls=[
                                        ft.Text(label, size=11.5, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                        ft.Container(
                                            bgcolor=C_MUTED,
                                            padding=ft.Padding.symmetric(horizontal=6, vertical=2),
                                            border_radius=ft.BorderRadius.all(6),
                                            content=ft.Text(status_badge, size=9, weight=ft.FontWeight.BOLD, color=status_col),
                                        ),
                                    ],
                                ),
                                ft.Text(
                                    f"Tgt {target_pct:.0f}% • Akt {actual_pct:.1f}%",
                                    size=10.5,
                                    color=C_DANGER if over_budget else C_SECONDARY,
                                    weight=ft.FontWeight.BOLD,
                                ),
                            ],
                        ),
                        ft.ProgressBar(
                            value=min(1.0, actual_pct / max(1.0, target_pct)),
                            bar_height=7,
                            color=C_DANGER if over_budget else color,
                            bgcolor=C_MUTED,
                            border_radius=ft.BorderRadius.all(4),
                        ),
                        ft.Text(f"Realisasi: {format_rp(actual_rp)} dari alokasi {format_rp(target_pct / 100.0 * base_income)}", size=10, color=C_TEXT_LIGHT),
                    ],
                ),
            )

        cat_sections = []
        for tipe_key in ["PENDAPATAN", "KEBUTUHAN", "KEINGINAN", "TABUNGAN & INVESTASI"]:
            items = ov["breakdown_by_tipe"].get(tipe_key, [])
            rows_ui = []
            for it in items:
                tgt = it["target"]
                act = it["actual"]
                pct_cat = (act / tgt) if tgt > 0 else (1.0 if act > 0 else 0.0)
                jt_txt = f" (Jatuh Tempo Tgl {it['jatuh_tempo']})" if it.get("jatuh_tempo") else ""
                rows_ui.append(
                    ft.Container(
                        bgcolor=C_BG_DEEP,
                        padding=ft.Padding.symmetric(horizontal=10, vertical=7),
                        border_radius=ft.BorderRadius.all(8),
                        content=ft.Column(
                            spacing=3,
                            controls=[
                                ft.Row(
                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    controls=[
                                        ft.Text(f"{it['kategori']}{jt_txt}", size=12, color=C_WHITE, weight=ft.FontWeight.W_500, expand=True),
                                        ft.Text(f"{format_rp(act)} / {format_rp(tgt)}", size=11, color=C_SECONDARY),
                                    ],
                                ),
                                ft.ProgressBar(
                                    value=min(1.0, pct_cat),
                                    bar_height=5,
                                    color=C_PRIMARY if act <= tgt or tipe_key in ("PENDAPATAN", "TABUNGAN & INVESTASI") else C_DANGER,
                                    bgcolor=C_MUTED,
                                    border_radius=ft.BorderRadius.all(3),
                                ),
                            ],
                        ),
                    )
                )

            cat_sections.append(
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(12),
                    border_radius=ft.BorderRadius.all(12),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Column(
                        spacing=6,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text(f"📌 {tipe_key}", size=12.5, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                                    ft.Text(f"Target: {format_rp(ov['target_by_tipe'].get(tipe_key, 0))}", size=11, color=C_SECONDARY),
                                ],
                            ),
                            ft.Column(spacing=5, controls=rows_ui),
                        ],
                    ),
                )
            )

        return ft.ListView(
            expand=True,
            spacing=10,
            padding=ft.Padding.only(bottom=24),
            controls=[
                self.build_period_selector(show_month=True),
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(15),
                    border=ft.Border.all(1, C_SECONDARY),
                    shadow=card_shadow(),
                    content=ft.Column(
                        spacing=10,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Row(
                                        spacing=6,
                                        controls=[
                                            ft.Icon(ft.Icons.DONUT_LARGE_ROUNDED, color=C_SECONDARY, size=18),
                                            ft.Text("VISUALISASI ALOKASI UANG", size=13, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                        ],
                                    ),
                                    ft.Container(
                                        bgcolor=C_MUTED,
                                        padding=ft.Padding.symmetric(horizontal=8, vertical=2),
                                        border_radius=ft.BorderRadius.all(6),
                                        content=ft.Text(f"Basis In: {format_rp(base_income)}", size=10, color=C_SUCCESS, weight=ft.FontWeight.BOLD),
                                    ),
                                ],
                            ),
                            ft.Container(
                                border_radius=ft.BorderRadius.all(10),
                                clip_behavior=ft.ClipBehavior.HARD_EDGE,
                                border=ft.Border.all(1, C_MUTED),
                                content=ft.Row(spacing=2, controls=segments),
                            ),
                            ft.Row(
                                wrap=True,
                                spacing=8,
                                controls=[
                                    ft.Row(spacing=4, controls=[ft.Container(width=10, height=10, bgcolor=C_SECONDARY, border_radius=ft.BorderRadius.all(3)), ft.Text(f"Kebutuhan: {pct_keb_act:.1f}%", size=10, color=C_TEXT_LIGHT)]),
                                    ft.Row(spacing=4, controls=[ft.Container(width=10, height=10, bgcolor=C_WARNING, border_radius=ft.BorderRadius.all(3)), ft.Text(f"Keinginan: {pct_kei_act:.1f}%", size=10, color=C_TEXT_LIGHT)]),
                                    ft.Row(spacing=4, controls=[ft.Container(width=10, height=10, bgcolor=C_PRIMARY, border_radius=ft.BorderRadius.all(3)), ft.Text(f"Tabungan: {pct_tab_act:.1f}%", size=10, color=C_TEXT_LIGHT)]),
                                    ft.Row(spacing=4, controls=[ft.Container(width=10, height=10, bgcolor=C_SUCCESS if ov["selisih"] >= 0 else C_DANGER, border_radius=ft.BorderRadius.all(3)), ft.Text(f"Sisa: {remain_pct:.1f}%", size=10, color=C_SUCCESS if ov["selisih"] >= 0 else C_DANGER, weight=ft.FontWeight.BOLD)]),
                                ],
                            ),
                            ft.Divider(color=C_MUTED, height=1),
                            ft.Row(
                                spacing=12,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Stack(
                                        alignment=ft.Alignment.CENTER,
                                        controls=[
                                            ft.ProgressRing(value=1.0, color=C_MUTED, stroke_width=8, width=78, height=78),
                                            ft.ProgressRing(value=min(1.0, used_pct / 100.0), color=health_color, stroke_width=8, width=78, height=78),
                                            ft.Column(
                                                alignment=ft.MainAxisAlignment.CENTER,
                                                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                                                spacing=0,
                                                controls=[
                                                    ft.Text(f"{used_pct:.0f}%", size=14, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                                    ft.Text("Terpakai", size=8.5, color=C_TEXT_LIGHT),
                                                ],
                                            ),
                                        ],
                                    ),
                                    ft.Column(
                                        expand=True,
                                        spacing=3,
                                        controls=[
                                            ft.Container(
                                                bgcolor=C_BG_DEEP,
                                                padding=ft.Padding.symmetric(horizontal=8, vertical=3),
                                                border_radius=ft.BorderRadius.all(6),
                                                border=ft.Border.all(1, health_color),
                                                content=ft.Text(health_title, size=10.5, weight=ft.FontWeight.BOLD, color=health_color),
                                            ),
                                            ft.Text(health_desc, size=11, color=C_TEXT_LIGHT),
                                            ft.Text(f"Sisa Saldo: {format_rp(ov['selisih'])}", size=12, weight=ft.FontWeight.BOLD, color=health_color),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),
                ),
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_PRIMARY),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Text("💰 CASH FLOW SUMMARY BULAN INI", size=13, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box("Total Pendapatan", format_rp(ov["pendapatan"]), C_SUCCESS),
                                    self._finance_box("Selisih (Sisa Uang)", format_rp(ov["selisih"]), C_SUCCESS if ov["selisih"] >= 0 else C_DANGER),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box("Total Kebutuhan", format_rp(ov["kebutuhan"]), C_TEXT_LIGHT),
                                    self._finance_box("Total Keinginan", format_rp(ov["keinginan"]), C_WARNING),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box("Tabungan & Investasi", format_rp(ov["tabungan"]), C_SECONDARY),
                                    self._finance_box("Total Pengeluaran", format_rp(ov["total_alokasi"]), C_WHITE),
                                ],
                            ),
                        ],
                    ),
                ),
                *(
                    [
                        ft.Container(
                            bgcolor=C_CARD_DARK,
                            padding=ft.Padding.all(14),
                            border_radius=ft.BorderRadius.all(14),
                            border=ft.Border.all(1, C_MUTED),
                            content=ft.Column(
                                spacing=6,
                                controls=[
                                    ft.Row(
                                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                        controls=[
                                            ft.Text("🏆 TOP 5 PENGELUARAN TERBESAR", size=12.5, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                                            ft.Text(f"Total: {format_rp(tot_exp_only)}", size=11, color=C_SECONDARY),
                                        ],
                                    ),
                                    ft.Column(spacing=4, controls=top_exp_rows),
                                ],
                            ),
                        )
                    ]
                    if top_exp_rows
                    else []
                ),
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Text(f"📊 BREAKDOWN PERSENTASE ({rule['mode']})", size=13, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                            build_rule_row("KEBUTUHAN", rule["kebutuhan_target_pct"], rule["kebutuhan_actual_pct"], ov["kebutuhan"], C_PRIMARY),
                            build_rule_row("KEINGINAN", rule["keinginan_target_pct"], rule["keinginan_actual_pct"], ov["keinginan"], C_WARNING),
                            build_rule_row("TABUNGAN & INVESTASI", rule["tabungan_target_pct"], rule["tabungan_actual_pct"], ov["tabungan"], C_SUCCESS),
                        ],
                    ),
                ),
                *cat_sections,
            ],
        )

    def build_yearly_finance_subview(self) -> ft.Control:
        y_ov = self.db.get_yearly_finance_overview(self.selected_year)
        tot = y_ov["totals"]
        tab_progress_pct = (tot["TABUNGAN & INVESTASI"] / max(1.0, y_ov["annual_tab_target"])) * 100.0

        max_m_val = max(
            [max(m["pendapatan"], m["kebutuhan"] + m["keinginan"] + m["tabungan"]) for m in y_ov["months"]] + [1000000.0]
        )

        # Visual 12-Month Dual Bar Chart (Income vs Allocation)
        chart_cols = []
        for m in y_ov["months"]:
            m_idx = m["month"]
            in_v = m["pendapatan"]
            out_v = m["kebutuhan"] + m["keinginan"] + m["tabungan"]
            h_in = max(4, int(round((in_v / max_m_val) * 70))) if in_v > 0 else 3
            h_out = max(4, int(round((out_v / max_m_val) * 70))) if out_v > 0 else 3
            is_cur = m_idx == self.selected_month
            chart_cols.append(
                ft.Column(
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    alignment=ft.MainAxisAlignment.END,
                    spacing=3,
                    expand=True,
                    controls=[
                        ft.Row(
                            spacing=2,
                            alignment=ft.MainAxisAlignment.CENTER,
                            vertical_alignment=ft.CrossAxisAlignment.END,
                            controls=[
                                ft.Container(
                                    width=6,
                                    height=h_in,
                                    bgcolor=C_SUCCESS if in_v > 0 else C_MUTED,
                                    border_radius=ft.BorderRadius.all(2),
                                    tooltip=f"{MONTH_NAMES_ID[m_idx]} Pendapatan: {format_rp(in_v)}",
                                ),
                                ft.Container(
                                    width=6,
                                    height=h_out,
                                    bgcolor=C_PRIMARY if out_v > 0 else C_MUTED,
                                    border_radius=ft.BorderRadius.all(2),
                                    tooltip=f"{MONTH_NAMES_ID[m_idx]} Alokasi: {format_rp(out_v)}",
                                ),
                            ],
                        ),
                        ft.Text(
                            MONTH_SHORT_ID[m_idx],
                            size=8.5,
                            weight=ft.FontWeight.BOLD if is_cur else ft.FontWeight.NORMAL,
                            color=C_WHITE if is_cur else C_SECONDARY,
                        ),
                    ],
                )
            )

        month_cards = []
        for m in y_ov["months"]:
            m_idx = m["month"]
            month_cards.append(
                ft.Container(
                    bgcolor=C_BG_DEEP,
                    padding=ft.Padding.all(10),
                    border_radius=ft.BorderRadius.all(8),
                    border=ft.Border.all(1, C_PRIMARY if m_idx == self.selected_month else C_MUTED),
                    on_click=lambda e, mi=m_idx: (setattr(self, "selected_month", mi), setattr(self, "finance_subtab", 1), self.render()),
                    content=ft.Column(
                        spacing=4,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text(f"{MONTH_NAMES_ID[m_idx]} {self.selected_year}", size=12, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                    ft.Text(
                                        f"Selisih: {format_rp(m['selisih'])}",
                                        size=11,
                                        weight=ft.FontWeight.BOLD,
                                        color=C_SUCCESS if m["selisih"] >= 0 else C_DANGER,
                                    ),
                                ],
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                wrap=True,
                                controls=[
                                    ft.Text(f"In: {format_rp_compact(m['pendapatan'])}", size=10.5, color=C_SUCCESS),
                                    ft.Text(f"Keb: {format_rp_compact(m['kebutuhan'])}", size=10.5, color=C_TEXT_LIGHT),
                                    ft.Text(f"Kei: {format_rp_compact(m['keinginan'])}", size=10.5, color=C_WARNING),
                                    ft.Text(f"Tab: {format_rp_compact(m['tabungan'])}", size=10.5, color=C_SECONDARY),
                                ],
                            ),
                        ],
                    ),
                )
            )

        return ft.ListView(
            expand=True,
            spacing=10,
            padding=ft.Padding.only(bottom=24),
            controls=[
                self.build_period_selector(show_month=False),
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(15),
                    border=ft.Border.all(1, C_PRIMARY),
                    shadow=card_shadow(),
                    content=ft.Column(
                        spacing=10,
                        controls=[
                            ft.Text(f"📈 OVERVIEW KEUANGAN TAHUNAN ({self.selected_year})", size=14, weight=ft.FontWeight.BOLD, color=C_WHITE),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box("Pendapatan Setahun", format_rp(tot["PENDAPATAN"]), C_SUCCESS),
                                    self._finance_box("Selisih Setahun", format_rp(y_ov["selisih_tahun"]), C_SECONDARY if y_ov["selisih_tahun"] >= 0 else C_DANGER),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box("Total Kebutuhan", format_rp(tot["KEBUTUHAN"]), C_TEXT_LIGHT),
                                    self._finance_box("Total Keinginan", format_rp(tot["KEINGINAN"]), C_WARNING),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box("Total Tabungan & Invest", format_rp(tot["TABUNGAN & INVESTASI"]), C_SUCCESS),
                                    self._finance_box("Target Tabungan Thn", format_rp(y_ov["annual_tab_target"]), C_SECONDARY),
                                ],
                            ),
                            ft.Text(
                                f"Progres Tabungan & Investasi Tahunan: {tab_progress_pct:.1f}%",
                                size=11,
                                color=C_TEXT_LIGHT,
                            ),
                            ft.ProgressBar(value=min(1.0, tab_progress_pct / 100.0), bar_height=8, color=C_SECONDARY, bgcolor=C_MUTED, border_radius=ft.BorderRadius.all(4)),
                            ft.Container(
                                bgcolor=C_BG_DEEP,
                                padding=ft.Padding.only(left=8, right=8, top=12, bottom=8),
                                border_radius=ft.BorderRadius.all(12),
                                border=ft.Border.all(1, C_MUTED),
                                height=115,
                                content=ft.Row(
                                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                    vertical_alignment=ft.CrossAxisAlignment.END,
                                    controls=chart_cols,
                                ),
                            ),
                        ],
                    ),
                ),
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Column(
                        spacing=6,
                        controls=[
                            ft.Text("🗓️ ARUS KAS 12 BULAN (Ketuk Bulan untuk Detail)", size=13, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                            ft.Column(spacing=6, controls=month_cards),
                        ],
                    ),
                ),
            ],
        )

    # =========================================================================
    # TAB 4: SET-UP BUDGET, TRACKER LANGGANAN & SAVING GOALS
    # =========================================================================
    def build_setup_view(self) -> ft.Control:
        def switch_sub(idx: int):
            self.setup_subtab = idx
            self.render()

        sub_bar = self.build_subtab_bar(
            ["Set-Up Budget", "Langganan", "Saving Goals 🌟"],
            self.setup_subtab,
            switch_sub,
        )

        if self.setup_subtab == 0:
            content = self.build_setup_budget_subview()
        elif self.setup_subtab == 1:
            content = self.build_subscriptions_subview()
        else:
            content = self.build_saving_goals_subview()

        return ft.Column(expand=True, spacing=6, controls=[sub_bar, content])

    def build_setup_budget_subview(self) -> ft.Control:
        cur_mode = self.db.get_setting("budget_mode", "50/30/20")
        cats = self.db.get_finance_categories()

        def apply_budget_rule(mode_str: str, kb: int, ki: int, tb: int):
            self.db.set_setting("budget_mode", mode_str)
            self.db.set_setting("pct_kebutuhan", str(kb))
            self.db.set_setting("pct_keinginan", str(ki))
            self.db.set_setting("pct_tabungan", str(tb))
            self.notify(f"Metode Budget diubah ke {mode_str} ({kb}% / {ki}% / {tb}%)", C_PRIMARY)
            self.render()

        def open_add_cat_dialog(e):
            dd_t = ft.Dropdown(
                label="Tipe Kategori",
                value="KEBUTUHAN",
                color=C_WHITE,
                bgcolor=C_MUTED,
                options=[
                    ft.dropdown.Option(key="PENDAPATAN", text="PENDAPATAN"),
                    ft.dropdown.Option(key="KEBUTUHAN", text="KEBUTUHAN"),
                    ft.dropdown.Option(key="KEINGINAN", text="KEINGINAN"),
                    ft.dropdown.Option(key="TABUNGAN & INVESTASI", text="TABUNGAN & INVESTASI"),
                ],
            )
            tf_k = ft.TextField(label="Nama Kategori", color=C_WHITE, bgcolor=C_MUTED)
            tf_jt = ft.TextField(label="Tanggal Jatuh Tempo (1-31, opsional)", keyboard_type=ft.KeyboardType.NUMBER, color=C_WHITE, bgcolor=C_MUTED)
            tf_tg = ft.TextField(label="Target Budget Bulanan (Rp)", keyboard_type=ft.KeyboardType.NUMBER, color=C_WHITE, bgcolor=C_MUTED)

            def save_c(ev):
                if not tf_k.value.strip():
                    self.notify("Nama kategori wajib diisi!", C_DANGER)
                    return
                jt_val = int(tf_jt.value) if (tf_jt.value and tf_jt.value.strip().isdigit()) else None
                try:
                    tg_val = float((tf_tg.value or "0").replace(".", "").replace(",", "").strip())
                except ValueError:
                    tg_val = 0.0
                self.db.save_finance_category(dd_t.value, tf_k.value, jt_val, tg_val)
                self.page.pop_dialog()
                self.notify("Kategori Set-Up disimpan!", C_PRIMARY)
                self.render()

            self.page.show_dialog(
                ft.AlertDialog(
                    bgcolor=C_CARD_DARK,
                    title=ft.Text("Tambah Kategori Set-Up", color=C_WHITE),
                    content=ft.Column(tight=True, spacing=8, controls=[dd_t, tf_k, tf_jt, tf_tg]),
                    actions=[
                        ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                        ft.Button(content="Simpan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save_c),
                    ],
                )
            )

        def open_edit_cat_dialog(c: dict):
            dd_t = ft.Dropdown(
                label="Tipe Kategori",
                value=c["tipe"],
                color=C_WHITE,
                bgcolor=C_MUTED,
                options=[
                    ft.dropdown.Option(key="PENDAPATAN", text="PENDAPATAN"),
                    ft.dropdown.Option(key="KEBUTUHAN", text="KEBUTUHAN"),
                    ft.dropdown.Option(key="KEINGINAN", text="KEINGINAN"),
                    ft.dropdown.Option(key="TABUNGAN & INVESTASI", text="TABUNGAN & INVESTASI"),
                ],
            )
            tf_k = ft.TextField(label="Nama Kategori", value=c["kategori"], color=C_WHITE, bgcolor=C_MUTED)
            tf_jt = ft.TextField(
                label="Tanggal Jatuh Tempo (1-31, opsional)",
                value=str(c["jatuh_tempo"]) if c.get("jatuh_tempo") else "",
                keyboard_type=ft.KeyboardType.NUMBER,
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            tf_tg = ft.TextField(
                label="Target / Nominal Bulanan (Rp)",
                value=str(int(round(float(c["target"] or 0)))),
                keyboard_type=ft.KeyboardType.NUMBER,
                color=C_WHITE,
                bgcolor=C_MUTED,
            )

            def save_edited_c(ev):
                if not tf_k.value.strip():
                    self.notify("Nama kategori wajib diisi!", C_DANGER)
                    return
                jt_val = int(tf_jt.value) if (tf_jt.value and tf_jt.value.strip().isdigit()) else None
                try:
                    tg_val = float((tf_tg.value or "0").replace(".", "").replace(",", "").strip())
                except ValueError:
                    tg_val = 0.0
                self.db.save_finance_category(dd_t.value, tf_k.value, jt_val, tg_val, cat_id=c["id"])
                self.page.pop_dialog()
                self.notify(f"Kategori '{tf_k.value}' berhasil diperbarui!", C_SUCCESS)
                self.render()

            self.page.show_dialog(
                ft.AlertDialog(
                    bgcolor=C_CARD_DARK,
                    title=ft.Text(f"Edit {c['kategori']}", color=C_WHITE, weight=ft.FontWeight.BOLD),
                    content=ft.Column(tight=True, spacing=8, controls=[dd_t, tf_k, tf_jt, tf_tg]),
                    actions=[
                        ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                        ft.Button(content="Simpan Perubahan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save_edited_c),
                    ],
                )
            )

        grouped_ui = []
        for t_name in ["PENDAPATAN", "KEBUTUHAN", "KEINGINAN", "TABUNGAN & INVESTASI"]:
            t_cats = [c for c in cats if c["tipe"] == t_name]
            tot_t = sum(c["target"] or 0 for c in t_cats)
            c_rows = []
            for c in t_cats:
                jt_str = f" • Tgl {c['jatuh_tempo']}" if c.get("jatuh_tempo") else ""
                c_rows.append(
                    ft.Container(
                        bgcolor=C_BG_DEEP,
                        padding=ft.Padding.symmetric(horizontal=10, vertical=6),
                        border_radius=ft.BorderRadius.all(8),
                        content=ft.Row(
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            vertical_alignment=ft.CrossAxisAlignment.CENTER,
                            controls=[
                                ft.Container(
                                    expand=True,
                                    on_click=lambda e, cat_item=c: open_edit_cat_dialog(cat_item),
                                    content=ft.Text(f"{c['kategori']}{jt_str}", size=12, color=C_WHITE),
                                ),
                                ft.Row(
                                    spacing=2,
                                    controls=[
                                        ft.Text(format_rp(c["target"]), size=12, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                                        ft.IconButton(
                                            icon=ft.Icons.EDIT_OUTLINED,
                                            icon_color=C_SECONDARY,
                                            icon_size=16,
                                            width=30,
                                            height=30,
                                            style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                            tooltip="Edit Budget / Kategori",
                                            on_click=lambda e, cat_item=c: open_edit_cat_dialog(cat_item),
                                        ),
                                        ft.IconButton(
                                            icon=ft.Icons.DELETE_OUTLINE,
                                            icon_color=C_DANGER,
                                            icon_size=16,
                                            width=30,
                                            height=30,
                                            style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                            tooltip="Hapus Kategori",
                                            on_click=lambda e, cid=c["id"]: (self.db.delete_finance_category(cid), self.render()),
                                        ),
                                    ],
                                ),
                            ],
                        ),
                    )
                )

            grouped_ui.append(
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(12),
                    border_radius=ft.BorderRadius.all(12),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Column(
                        spacing=6,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text(t_name, size=13, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                                    ft.Text(f"Total Target: {format_rp(tot_t)}", size=12, weight=ft.FontWeight.BOLD, color=C_SUCCESS),
                                ],
                            ),
                            ft.Column(spacing=4, controls=c_rows),
                        ],
                    ),
                )
            )

        return ft.ListView(
            expand=True,
            spacing=10,
            padding=ft.Padding.only(bottom=24),
            controls=[
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_PRIMARY),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Text("⚙️ METODE BUDGET (50/30/20 atau 70/20/10)", size=13, weight=ft.FontWeight.BOLD, color=C_WHITE),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    ft.Button(
                                        content="50 / 30 / 20 Rule",
                                        bgcolor=C_PRIMARY if cur_mode == "50/30/20" else C_MUTED,
                                        color=C_WHITE,
                                        expand=True,
                                        on_click=lambda e: apply_budget_rule("50/30/20", 50, 30, 20),
                                    ),
                                    ft.Button(
                                        content="70 / 20 / 10 Rule",
                                        bgcolor=C_PRIMARY if cur_mode == "70/20/10" else C_MUTED,
                                        color=C_WHITE,
                                        expand=True,
                                        on_click=lambda e: apply_budget_rule("70/20/10", 70, 20, 10),
                                    ),
                                ],
                            ),
                            ft.Button(
                                content="+ Tambah Kategori & Target Baru",
                                bgcolor=C_SECONDARY,
                                color=C_BG_DEEP,
                                on_click=open_add_cat_dialog,
                            ),
                        ],
                    ),
                ),
                *grouped_ui,
            ],
        )

    def build_subscriptions_subview(self) -> ft.Control:
        sub_data = self.db.get_subscriptions()

        def open_sub_dialog(existing_sub: dict = None):
            is_edit = existing_sub is not None
            tf_nama = ft.TextField(
                label="Nama Langganan (misal: Netflix, Spotify)",
                value=existing_sub["nama"] if is_edit else "",
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            tf_kat = ft.TextField(
                label="Kategori (Entertainment / Productivity / Belanja)",
                value=existing_sub["kategori"] if is_edit else "Entertainment",
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            dd_bill = ft.Dropdown(
                label="Siklus Tagihan (Billing)",
                value=existing_sub["billing"] if is_edit else "Monthly",
                color=C_WHITE,
                bgcolor=C_MUTED,
                options=[
                    ft.dropdown.Option(key="Monthly", text="Monthly (Bulanan)"),
                    ft.dropdown.Option(key="Annually", text="Annually (Tahunan)"),
                ],
            )
            tf_tgl = ft.TextField(
                label="Tanggal Renew (1-31)",
                value=str(existing_sub["tanggal_renew"] or "5") if is_edit else "5",
                keyboard_type=ft.KeyboardType.NUMBER,
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            tf_harga = ft.TextField(
                label="Harga Tagihan (Rp)",
                value=str(int(round(float(existing_sub["harga"])))) if is_edit else "",
                keyboard_type=ft.KeyboardType.NUMBER,
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            sw_auto = ft.Switch(
                label="Auto Renew Aktif",
                value=bool(existing_sub["auto_renew"]) if is_edit else True,
                active_color=C_PRIMARY,
            )

            def save_sub(ev):
                try:
                    h_val = float((tf_harga.value or "0").replace(".", "").replace(",", "").strip())
                    if not tf_nama.value.strip() or h_val <= 0:
                        raise ValueError()
                except ValueError:
                    self.notify("Nama & harga langganan harus valid!", C_DANGER)
                    return
                tgl_r = int(tf_tgl.value) if (tf_tgl.value and tf_tgl.value.strip().isdigit()) else None
                if is_edit:
                    self.db.update_subscription(existing_sub["id"], tf_nama.value, tf_kat.value, dd_bill.value, sw_auto.value, tgl_r, h_val)
                    self.notify("Langganan berhasil diperbarui!", C_SUCCESS)
                else:
                    self.db.add_subscription(tf_nama.value, tf_kat.value, dd_bill.value, sw_auto.value, tgl_r, h_val)
                    self.notify("Langganan berhasil ditambahkan!", C_PRIMARY)
                self.page.pop_dialog()
                self.render()

            self.page.show_dialog(
                ft.AlertDialog(
                    bgcolor=C_CARD_DARK,
                    title=ft.Text("Edit Langganan" if is_edit else "Tambah Tracker Langganan", color=C_WHITE, weight=ft.FontWeight.BOLD),
                    content=ft.Column(tight=True, spacing=8, controls=[tf_nama, tf_kat, dd_bill, tf_tgl, tf_harga, sw_auto]),
                    actions=[
                        ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                        ft.Button(content="Simpan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save_sub),
                    ],
                )
            )

        def record_sub_to_tx(sub_item: dict):
            today_str = datetime.date.today().strftime("%Y-%m-%d")
            self.db.add_transaction(
                tanggal=today_str,
                tipe="KEINGINAN",
                kategori="Langganan",
                deskripsi=sub_item["nama"],
                total=sub_item["harga_bulanan"],
                notes=f"Auto-catat dari Tracker Langganan ({sub_item['billing']})",
            )
            self.notify(f"Langganan '{sub_item['nama']}' dicatat ke Transaksi bulan ini!", C_SUCCESS)

        sub_cards = []
        for s in sub_data["items"]:
            sub_cards.append(
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(12),
                    border_radius=ft.BorderRadius.all(12),
                    border=ft.Border.all(1, C_MUTED),
                    content=ft.Column(
                        spacing=6,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Row(
                                        spacing=8,
                                        controls=[
                                            ft.Text(s["nama"], size=15, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                            ft.Container(
                                                bgcolor=C_PRIMARY,
                                                padding=ft.Padding.symmetric(horizontal=8, vertical=2),
                                                border_radius=ft.BorderRadius.all(8),
                                                content=ft.Text(s["kategori"], size=10, color=C_WHITE),
                                            ),
                                        ],
                                    ),
                                    ft.Row(
                                        spacing=2,
                                        controls=[
                                            ft.IconButton(
                                                icon=ft.Icons.EDIT_OUTLINED,
                                                icon_color=C_SECONDARY,
                                                icon_size=16,
                                                width=30,
                                                height=30,
                                                style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                                tooltip="Edit Langganan",
                                                on_click=lambda e, item=s: open_sub_dialog(item),
                                            ),
                                            ft.IconButton(
                                                icon=ft.Icons.DELETE_OUTLINE,
                                                icon_color=C_DANGER,
                                                icon_size=16,
                                                width=30,
                                                height=30,
                                                style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                                tooltip="Hapus Langganan",
                                                on_click=lambda e, sid=s["id"]: (self.db.delete_subscription(sid), self.render()),
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text(f"Bulanan: {format_rp(s['harga_bulanan'])}", size=12, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                                    ft.Text(f"Tahunan: {format_rp(s['harga_tahunan'])}", size=12, color=C_TEXT_LIGHT),
                                ],
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text(
                                        f"Billing: {s['billing']} • Auto Renew: {'Ya' if s['auto_renew'] else 'Tidak'}"
                                        + (f" • Tgl {s['tanggal_renew']}" if s.get("tanggal_renew") else ""),
                                        size=11,
                                        color=C_TEXT_LIGHT,
                                    ),
                                    ft.Container(
                                        bgcolor=C_PRIMARY,
                                        padding=ft.Padding.symmetric(horizontal=9, vertical=4),
                                        border_radius=ft.BorderRadius.all(6),
                                        on_click=lambda e, item=s: record_sub_to_tx(item),
                                        content=ft.Text("+ Catat Transaksi", size=10, color=C_WHITE, weight=ft.FontWeight.BOLD),
                                    ),
                                ],
                            ),
                        ],
                    ),
                )
            )

        cat_summary_rows = [
            ft.Row(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                controls=[
                    ft.Text(k, size=12, color=C_TEXT_LIGHT),
                    ft.Text(f"{format_rp(v)} / bln", size=12, weight=ft.FontWeight.BOLD, color=C_SECONDARY),
                ],
            )
            for k, v in sub_data["by_category"].items()
        ]

        return ft.ListView(
            expand=True,
            spacing=10,
            padding=ft.Padding.only(bottom=24),
            controls=[
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_PRIMARY),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text("🔄 TRACKER LANGGANAN", size=14, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                    ft.Button(content="+ Langganan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=lambda e: open_sub_dialog(None)),
                                ],
                            ),
                            ft.Row(
                                spacing=8,
                                controls=[
                                    self._finance_box("Total Langganan Bulanan", format_rp(sub_data["total_bulanan"]), C_WARNING),
                                    self._finance_box("Total Langganan Tahunan", format_rp(sub_data["total_tahunan"]), C_SECONDARY),
                                ],
                            ),
                            ft.Divider(color=C_MUTED),
                            ft.Text("Overview Langganan per Kategori:", size=12, weight=ft.FontWeight.BOLD, color=C_TEXT_LIGHT),
                            ft.Column(spacing=4, controls=cat_summary_rows),
                        ],
                    ),
                ),
                *sub_cards,
            ],
        )

    def build_saving_goals_subview(self) -> ft.Control:
        goals = self.db.get_saving_goals()

        def open_goal_dialog(existing_goal: dict = None):
            is_edit = existing_goal is not None
            tf_name = ft.TextField(
                label="Nama Impian / Goal (misal: Dana Darurat, Liburan)",
                value=existing_goal["nama_goal"] if is_edit else "",
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            tf_target = ft.TextField(
                label="Target Nominal (Rp)",
                value=str(int(round(float(existing_goal["target_nominal"])))) if is_edit else "",
                keyboard_type=ft.KeyboardType.NUMBER,
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            tf_curr = ft.TextField(
                label="Nominal Terkumpul Saat Ini (Rp)",
                value=str(int(round(float(existing_goal["terkumpul_nominal"] or 0)))) if is_edit else "0",
                keyboard_type=ft.KeyboardType.NUMBER,
                color=C_WHITE,
                bgcolor=C_MUTED,
            )
            tf_dead = ft.TextField(
                label="Target Tanggal Tercapai (YYYY-MM-DD)",
                value=(existing_goal["deadline"] or "2027-12-31") if is_edit else "2027-12-31",
                color=C_WHITE,
                bgcolor=C_MUTED,
            )

            def save_g(ev):
                try:
                    t_nom = float((tf_target.value or "0").replace(".", "").replace(",", "").strip())
                    c_nom = float((tf_curr.value or "0").replace(".", "").replace(",", "").strip())
                    if not tf_name.value.strip() or t_nom <= 0:
                        raise ValueError()
                except ValueError:
                    self.notify("Masukkan nama dan nominal target yang valid!", C_DANGER)
                    return
                self.db.save_saving_goal(
                    tf_name.value,
                    t_nom,
                    c_nom,
                    tf_dead.value,
                    goal_id=existing_goal["id"] if is_edit else None,
                )
                self.page.pop_dialog()
                self.notify("Saving Goal berhasil disimpan!", C_PRIMARY)
                self.render()

            self.page.show_dialog(
                ft.AlertDialog(
                    bgcolor=C_CARD_DARK,
                    title=ft.Text("Edit Saving Goal" if is_edit else "Tambah Saving Goal Baru", color=C_WHITE, weight=ft.FontWeight.BOLD),
                    content=ft.Column(tight=True, spacing=8, controls=[tf_name, tf_target, tf_curr, tf_dead]),
                    actions=[
                        ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                        ft.Button(content="Simpan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save_g),
                    ],
                )
            )

        def add_funds_to_goal(g_item: dict):
            tf_add = ft.TextField(label="Tambah Saldo Tabungan (Rp)", keyboard_type=ft.KeyboardType.NUMBER, color=C_WHITE, bgcolor=C_MUTED)

            def confirm_add(ev):
                try:
                    add_v = float((tf_add.value or "0").replace(".", "").replace(",", "").strip())
                except ValueError:
                    add_v = 0.0
                if add_v > 0:
                    new_total = float(g_item["terkumpul_nominal"] or 0) + add_v
                    self.db.save_saving_goal(
                        g_item["nama_goal"],
                        g_item["target_nominal"],
                        new_total,
                        g_item["deadline"] or "",
                        g_item["notes"] or "",
                        goal_id=g_item["id"],
                    )
                    self.page.pop_dialog()
                    self.notify(f"Saldo {format_rp(add_v)} ditambahkan ke '{g_item['nama_goal']}'!", C_SUCCESS)
                    self.render()

            self.page.show_dialog(
                ft.AlertDialog(
                    bgcolor=C_CARD_DARK,
                    title=ft.Text(f"Tambah Tabungan: {g_item['nama_goal']}", color=C_WHITE),
                    content=tf_add,
                    actions=[
                        ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                        ft.Button(content="Tambah", bgcolor=C_PRIMARY, color=C_WHITE, on_click=confirm_add),
                    ],
                )
            )

        goal_cards = []
        for g in goals:
            pct = (g["terkumpul_nominal"] / max(1.0, g["target_nominal"])) * 100.0
            sisa_goal = max(0.0, float(g["target_nominal"]) - float(g["terkumpul_nominal"] or 0))
            goal_cards.append(
                ft.Container(
                    bgcolor=C_CARD_DARK,
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_SUCCESS if pct >= 100 else C_MUTED),
                    content=ft.Column(
                        spacing=8,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Column(
                                        spacing=1,
                                        expand=True,
                                        controls=[
                                            ft.Text(g["nama_goal"], size=15, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                            ft.Text(
                                                f"Target Tanggal: {g['deadline'] or '-'}"
                                                + (f" • Kurang {format_rp(sisa_goal)}" if sisa_goal > 0 else " • Tercapai! 🎉"),
                                                size=10.5,
                                                color=C_SECONDARY,
                                            ),
                                        ],
                                    ),
                                    ft.Text(f"{pct:.1f}%", size=15, weight=ft.FontWeight.BOLD, color=C_SUCCESS if pct >= 100 else C_SECONDARY),
                                ],
                            ),
                            ft.ProgressBar(
                                value=min(1.0, pct / 100.0),
                                bar_height=9,
                                color=C_SUCCESS if pct >= 100 else C_PRIMARY,
                                bgcolor=C_MUTED,
                                border_radius=ft.BorderRadius.all(4),
                            ),
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                controls=[
                                    ft.Text(f"{format_rp(g['terkumpul_nominal'])} / {format_rp(g['target_nominal'])}", size=12, weight=ft.FontWeight.W_500, color=C_TEXT_LIGHT),
                                    ft.Row(
                                        spacing=4,
                                        controls=[
                                            ft.Button(
                                                content="+ Nabung",
                                                bgcolor=C_PRIMARY,
                                                color=C_WHITE,
                                                on_click=lambda e, item=g: add_funds_to_goal(item),
                                            ),
                                            ft.IconButton(
                                                icon=ft.Icons.EDIT_OUTLINED,
                                                icon_color=C_SECONDARY,
                                                icon_size=16,
                                                width=30,
                                                height=30,
                                                style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                                tooltip="Edit Target",
                                                on_click=lambda e, item=g: open_goal_dialog(item),
                                            ),
                                            ft.IconButton(
                                                icon=ft.Icons.DELETE_OUTLINE,
                                                icon_color=C_DANGER,
                                                icon_size=16,
                                                width=30,
                                                height=30,
                                                style=ft.ButtonStyle(padding=ft.Padding.all(2)),
                                                tooltip="Hapus Goal",
                                                on_click=lambda e, gid=g["id"]: (self.db.delete_saving_goal(gid), self.render()),
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),
                )
            )

        return ft.ListView(
            expand=True,
            spacing=10,
            padding=ft.Padding.only(bottom=24),
            controls=[
                ft.Container(
                    gradient=dark_card_gradient(),
                    padding=ft.Padding.all(14),
                    border_radius=ft.BorderRadius.all(14),
                    border=ft.Border.all(1, C_PRIMARY),
                    content=ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Column(
                                spacing=2,
                                expand=True,
                                controls=[
                                    ft.Text("🌟 TARGET TABUNGAN IMPIAN", size=14, weight=ft.FontWeight.BOLD, color=C_WHITE),
                                    ft.Text("Pantau dan capai target dana impian Anda", size=11, color=C_SECONDARY),
                                ],
                            ),
                            ft.Button(content="+ Target Baru", bgcolor=C_PRIMARY, color=C_WHITE, on_click=lambda e: open_goal_dialog(None)),
                        ],
                    ),
                ),
                *goal_cards,
            ],
        )


def main(page: ft.Page):
    MyLifeApp(page)


if __name__ == "__main__":
    ft.run(main)
