"""Tab Buku (tracker baca) dan Kalender (aktivitas harian) untuk MyLife."""
import calendar
import datetime
import flet as ft

C_BG_DEEP = "#0E0B1D"
C_CARD_DARK = "#2A1E4D"
C_PRIMARY = "#5A3B8C"
C_SECONDARY = "#8B7CC7"
C_TEXT_LIGHT = "#C9C2E3"
C_MUTED = "#3A3A52"
C_SUCCESS = "#4ADE80"
C_WARNING = "#FBBF24"
C_DANGER = "#F87171"
C_WHITE = "#FFFFFF"
C_BLUE = "#60A5FA"

MONTHS_ID = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
             "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
DAYS_ID = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]
DAY_FULL_ID = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

# warna titik indikator per jenis aktivitas
MARK_COLORS = {
    "keuangan": C_SUCCESS,
    "baca": C_BLUE,
    "kosakata": C_SECONDARY,
    "belajar": C_WARNING,
}
MARK_LABELS = {
    "keuangan": "Keuangan",
    "baca": "Baca buku",
    "kosakata": "Kosakata",
    "belajar": "Belajar/Lainnya",
}
ACTIVITY_CATEGORIES = ["Belajar", "Olahraga", "Kerja", "Ibadah", "Lainnya"]


def _rp(v: float) -> str:
    return "Rp " + f"{int(round(v)):,}".replace(",", ".")


class BooksCalendarMixin:
    """Mixin untuk MyLifeApp. Memakai self.db, self.page, self.render, self.notify."""

    def _init_books_calendar_state(self):
        today = datetime.date.today()
        self.cal_year = today.year
        self.cal_month = today.month
        self.cal_selected = today.strftime("%Y-%m-%d")

    # ------------------------------------------------------------------
    # helper kecil
    # ------------------------------------------------------------------
    def _bc_field(self, label, value="", **kw):
        return ft.TextField(label=label, value=value, color=C_WHITE, bgcolor=C_MUTED, **kw)

    def _bc_stat(self, label, value, color):
        return ft.Container(
            expand=True,
            bgcolor=C_CARD_DARK,
            border=ft.Border.all(1, C_MUTED),
            border_radius=ft.BorderRadius.all(12),
            padding=ft.Padding.symmetric(horizontal=6, vertical=10),
            content=ft.Column(
                spacing=2,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    ft.Text(str(value), size=18, weight=ft.FontWeight.BOLD, color=color),
                    ft.Text(label, size=10, color=C_TEXT_LIGHT, text_align=ft.TextAlign.CENTER),
                ],
            ),
        )

    def _bc_row_item(self, icon, color, title, subtitle="", on_delete=None):
        controls = [
            ft.Icon(icon, color=color, size=20),
            ft.Column(
                expand=True,
                spacing=1,
                controls=[
                    ft.Text(title, size=13, weight=ft.FontWeight.W_600, color=C_WHITE),
                    ft.Text(subtitle, size=11, color=C_TEXT_LIGHT, visible=bool(subtitle)),
                ],
            ),
        ]
        if on_delete:
            controls.append(ft.IconButton(icon=ft.Icons.DELETE_OUTLINE, icon_color=C_DANGER, icon_size=18, on_click=on_delete))
        return ft.Container(
            bgcolor=C_BG_DEEP,
            border=ft.Border.all(1, C_MUTED),
            border_radius=ft.BorderRadius.all(10),
            padding=ft.Padding.symmetric(horizontal=10, vertical=6),
            content=ft.Row(spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER, controls=controls),
        )

    # ------------------------------------------------------------------
    # DIALOG: tambah/edit buku
    # ------------------------------------------------------------------
    def _open_book_dialog(self, book=None):
        tf_judul = self._bc_field("Judul buku", book["judul"] if book else "")
        tf_penulis = self._bc_field("Penulis (opsional)", (book["penulis"] or "") if book else "")
        tf_total = self._bc_field("Total halaman", str(book["total_halaman"]) if book else "", keyboard_type=ft.KeyboardType.NUMBER)

        def save(ev):
            try:
                total = int((tf_total.value or "0").strip())
                if total <= 0:
                    raise ValueError()
            except ValueError:
                self.notify("Total halaman harus angka > 0!", C_DANGER)
                return
            if not (tf_judul.value or "").strip():
                self.notify("Judul buku wajib diisi!", C_DANGER)
                return
            self.db.save_book(tf_judul.value, tf_penulis.value or "", total, book["id"] if book else None)
            self.page.pop_dialog()
            self.notify("Buku disimpan!", C_PRIMARY)
            self.render()

        self.page.show_dialog(ft.AlertDialog(
            bgcolor=C_CARD_DARK,
            title=ft.Text("Edit Buku" if book else "Tambah Buku", color=C_WHITE, weight=ft.FontWeight.BOLD),
            content=ft.Column(tight=True, spacing=8, controls=[tf_judul, tf_penulis, tf_total]),
            actions=[
                ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                ft.Button(content="Simpan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save),
            ],
        ))

    # ------------------------------------------------------------------
    # DIALOG: catat bacaan (sampai halaman berapa)
    # ------------------------------------------------------------------
    def _open_reading_dialog(self, book=None, default_date=None):
        books = self.db.get_books()
        if not books:
            self.notify("Tambahkan buku dulu di tab Buku.", C_WARNING)
            return
        if book is None:
            book = next((b for b in books if b["halaman_sekarang"] < b["total_halaman"]), books[0])
        date_str = default_date or datetime.date.today().strftime("%Y-%m-%d")

        info = ft.Text("", size=12, color=C_SECONDARY)
        dd_book = ft.Dropdown(
            label="Buku", value=str(book["id"]), color=C_WHITE, bgcolor=C_MUTED,
            options=[ft.dropdown.Option(key=str(b["id"]), text=b["judul"]) for b in books],
        )
        tf_date = self._bc_field("Tanggal (YYYY-MM-DD)", date_str)
        tf_page = self._bc_field("Sampai halaman berapa?", "", keyboard_type=ft.KeyboardType.NUMBER)
        tf_notes = self._bc_field("Catatan (opsional)")

        def current_book():
            return next(b for b in books if str(b["id"]) == dd_book.value)

        def refresh_info(ev=None):
            b = current_book()
            info.value = f"Terakhir di halaman {b['halaman_sekarang']} dari {b['total_halaman']}"
            if ev is not None:
                self.page.update()

        dd_book.on_select = refresh_info
        refresh_info()

        def save(ev):
            b = current_book()
            try:
                to_page = int((tf_page.value or "").strip())
                datetime.datetime.strptime(tf_date.value.strip(), "%Y-%m-%d")
            except ValueError:
                self.notify("Halaman harus angka & tanggal format YYYY-MM-DD!", C_DANGER)
                return
            if to_page <= b["halaman_sekarang"]:
                self.notify(f"Halaman harus lebih dari {b['halaman_sekarang']}.", C_DANGER)
                return
            pages = self.db.add_reading_log(b["id"], tf_date.value.strip(), to_page, tf_notes.value or "")
            self.page.pop_dialog()
            self.notify(f"Tercatat +{pages} halaman. Lanjutkan terus!", C_PRIMARY)
            self.render()

        self.page.show_dialog(ft.AlertDialog(
            bgcolor=C_CARD_DARK,
            title=ft.Text("Catat Bacaan", color=C_WHITE, weight=ft.FontWeight.BOLD),
            content=ft.Column(tight=True, spacing=8, controls=[dd_book, info, tf_date, tf_page, tf_notes]),
            actions=[
                ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                ft.Button(content="Simpan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save),
            ],
        ))

    # ------------------------------------------------------------------
    # TAB BUKU
    # ------------------------------------------------------------------
    def build_books_view(self) -> ft.Control:
        stats = self.db.get_reading_stats()
        books = self.db.get_books()

        def del_book(b):
            def go(ev):
                self.db.delete_book(b["id"])
                self.page.pop_dialog()
                self.render()
            self.page.show_dialog(ft.AlertDialog(
                bgcolor=C_CARD_DARK,
                title=ft.Text("Hapus buku?", color=C_WHITE),
                content=ft.Text(f"'{b['judul']}' beserta riwayat bacanya akan dihapus.", color=C_TEXT_LIGHT),
                actions=[
                    ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                    ft.Button(content="Hapus", bgcolor=C_DANGER, color=C_WHITE, on_click=go),
                ],
            ))

        def del_log(log_id):
            self.db.delete_reading_log(log_id)
            self.render()

        cards = []
        for b in books:
            total = max(int(b["total_halaman"]), 1)
            cur = int(b["halaman_sekarang"])
            pct = min(cur / total, 1.0)
            done = cur >= total
            logs = self.db.get_reading_logs(b["id"], limit=3)
            log_rows = [
                ft.Row(
                    spacing=4,
                    controls=[
                        ft.Text(f"{lg['log_date']}  •  hal {lg['from_page']}→{lg['to_page']} (+{lg['pages_read']})",
                                size=11, color=C_TEXT_LIGHT, expand=True),
                        ft.IconButton(icon=ft.Icons.CLOSE, icon_size=14, icon_color=C_MUTED,
                                      on_click=lambda ev, i=lg["id"]: del_log(i)),
                    ],
                ) for lg in logs
            ]
            cards.append(ft.Container(
                bgcolor=C_CARD_DARK,
                border=ft.Border.all(1, C_SUCCESS if done else C_MUTED),
                border_radius=ft.BorderRadius.all(14),
                padding=ft.Padding.all(12),
                margin=ft.Margin.only(bottom=10),
                content=ft.Column(spacing=6, controls=[
                    ft.Row(vertical_alignment=ft.CrossAxisAlignment.START, controls=[
                        ft.Column(expand=True, spacing=1, controls=[
                            ft.Text(b["judul"], size=15, weight=ft.FontWeight.BOLD, color=C_WHITE),
                            ft.Text(b["penulis"] or "—", size=11, color=C_SECONDARY),
                        ]),
                        ft.IconButton(icon=ft.Icons.EDIT_OUTLINED, icon_size=18, icon_color=C_TEXT_LIGHT,
                                      on_click=lambda ev, bb=b: self._open_book_dialog(bb)),
                        ft.IconButton(icon=ft.Icons.DELETE_OUTLINE, icon_size=18, icon_color=C_DANGER,
                                      on_click=lambda ev, bb=b: del_book(bb)),
                    ]),
                    ft.ProgressBar(value=pct, color=C_SUCCESS if done else C_BLUE, bgcolor=C_MUTED, height=8, border_radius=4),
                    ft.Row(alignment=ft.MainAxisAlignment.SPACE_BETWEEN, controls=[
                        ft.Text(f"Halaman {cur} / {b['total_halaman']}", size=12, color=C_WHITE, weight=ft.FontWeight.W_600),
                        ft.Text("SELESAI ✓" if done else f"{int(pct * 100)}% • sisa {b['total_halaman'] - cur} hal",
                                size=12, color=C_SUCCESS if done else C_TEXT_LIGHT),
                    ]),
                    *log_rows,
                    ft.Button(content="Selesai dibaca" if done else "+ Catat Bacaan", bgcolor=C_PRIMARY, color=C_WHITE,
                              disabled=done, on_click=lambda ev, bb=b: self._open_reading_dialog(bb)),
                ]),
            ))

        if not cards:
            cards.append(ft.Container(
                padding=ft.Padding.all(30),
                alignment=ft.Alignment.CENTER,
                content=ft.Text("Belum ada buku. Tekan '+ Buku' untuk mulai melacak bacaanmu.", color=C_TEXT_LIGHT,
                                text_align=ft.TextAlign.CENTER),
            ))

        return ft.Column(
            expand=True, scroll=ft.ScrollMode.AUTO, spacing=10,
            controls=[
                ft.Row(alignment=ft.MainAxisAlignment.SPACE_BETWEEN, controls=[
                    ft.Column(spacing=1, controls=[
                        ft.Text("Tracker Baca Buku", size=17, weight=ft.FontWeight.BOLD, color=C_WHITE),
                        ft.Text("Catat sampai halaman berapa kamu membaca", size=11, color=C_SECONDARY),
                    ]),
                    ft.Button(content="+ Buku", bgcolor=C_PRIMARY, color=C_WHITE, on_click=lambda e: self._open_book_dialog(None)),
                ]),
                ft.Row(spacing=6, controls=[
                    self._bc_stat("Hari ini", stats["today"], C_BLUE),
                    self._bc_stat("Bulan ini", stats["month"], C_SECONDARY),
                    self._bc_stat("Total hal", stats["total"], C_WHITE),
                    self._bc_stat("Streak 🔥", f"{stats['streak']}h", C_WARNING),
                    self._bc_stat("Tamat", stats["finished"], C_SUCCESS),
                ]),
                *cards,
            ],
        )

    # ------------------------------------------------------------------
    # DIALOG: tambah aktivitas (belajar dsb.)
    # ------------------------------------------------------------------
    def _open_activity_dialog(self):
        tf_date = self._bc_field("Tanggal (YYYY-MM-DD)", self.cal_selected)
        dd_kat = ft.Dropdown(label="Kategori", value="Belajar", color=C_WHITE, bgcolor=C_MUTED,
                             options=[ft.dropdown.Option(key=k, text=k) for k in ACTIVITY_CATEGORIES])
        tf_judul = self._bc_field("Kegiatan (misal: Belajar Python)")
        tf_dur = self._bc_field("Durasi menit (opsional)", keyboard_type=ft.KeyboardType.NUMBER)
        tf_notes = self._bc_field("Catatan (opsional)")

        def save(ev):
            try:
                datetime.datetime.strptime(tf_date.value.strip(), "%Y-%m-%d")
                dur = int((tf_dur.value or "0").strip() or 0)
            except ValueError:
                self.notify("Tanggal/durasi tidak valid!", C_DANGER)
                return
            if not (tf_judul.value or "").strip():
                self.notify("Isi nama kegiatan!", C_DANGER)
                return
            self.db.add_calendar_activity(tf_date.value.strip(), dd_kat.value, tf_judul.value, dur, tf_notes.value or "")
            self.cal_selected = tf_date.value.strip()
            self.cal_year, self.cal_month = int(self.cal_selected[:4]), int(self.cal_selected[5:7])
            self.page.pop_dialog()
            self.notify("Kegiatan disimpan!", C_PRIMARY)
            self.render()

        self.page.show_dialog(ft.AlertDialog(
            bgcolor=C_CARD_DARK,
            title=ft.Text("Tambah Kegiatan", color=C_WHITE, weight=ft.FontWeight.BOLD),
            content=ft.Column(tight=True, spacing=8, controls=[tf_date, dd_kat, tf_judul, tf_dur, tf_notes]),
            actions=[
                ft.TextButton(content="Batal", on_click=lambda ev: self.page.pop_dialog()),
                ft.Button(content="Simpan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=save),
            ],
        ))

    # ------------------------------------------------------------------
    # TAB KALENDER
    # ------------------------------------------------------------------
    def build_calendar_view(self) -> ft.Control:
        y, m = self.cal_year, self.cal_month
        marks = self.db.get_month_activity_marks(y, m)
        today_str = datetime.date.today().strftime("%Y-%m-%d")

        def shift_month(delta):
            idx = y * 12 + (m - 1) + delta
            self.cal_year, self.cal_month = idx // 12, idx % 12 + 1
            self.render()

        def select_day(day):
            self.cal_selected = f"{y:04d}-{m:02d}-{day:02d}"
            self.render()

        def go_today(e):
            t = datetime.date.today()
            self.cal_year, self.cal_month = t.year, t.month
            self.cal_selected = t.strftime("%Y-%m-%d")
            self.render()

        header = ft.Container(
            bgcolor=C_CARD_DARK, border=ft.Border.all(1, C_MUTED), border_radius=ft.BorderRadius.all(12),
            padding=ft.Padding.symmetric(horizontal=6, vertical=4),
            content=ft.Row(alignment=ft.MainAxisAlignment.SPACE_BETWEEN, controls=[
                ft.IconButton(icon=ft.Icons.CHEVRON_LEFT, icon_color=C_WHITE, on_click=lambda e: shift_month(-1)),
                ft.Container(on_click=go_today, content=ft.Text(f"{MONTHS_ID[m]} {y}", size=16, weight=ft.FontWeight.BOLD, color=C_WHITE)),
                ft.IconButton(icon=ft.Icons.CHEVRON_RIGHT, icon_color=C_WHITE, on_click=lambda e: shift_month(1)),
            ]),
        )

        dow_row = ft.Row(spacing=2, controls=[
            ft.Container(expand=True, alignment=ft.Alignment.CENTER,
                         content=ft.Text(d, size=11, color=C_DANGER if i == 6 else C_SECONDARY, weight=ft.FontWeight.BOLD))
            for i, d in enumerate(DAYS_ID)
        ])

        week_rows = []
        for week in calendar.Calendar(firstweekday=0).monthdayscalendar(y, m):
            cells = []
            for day in week:
                if day == 0:
                    cells.append(ft.Container(expand=True, height=50))
                    continue
                ds = f"{y:04d}-{m:02d}-{day:02d}"
                is_sel = ds == self.cal_selected
                is_today = ds == today_str
                dots = [ft.Container(width=6, height=6, border_radius=3, bgcolor=MARK_COLORS[k])
                        for k in ("keuangan", "baca", "kosakata", "belajar") if k in marks.get(day, set())]
                cells.append(ft.Container(
                    expand=True, height=50,
                    bgcolor=C_PRIMARY if is_sel else C_CARD_DARK,
                    border=ft.Border.all(2 if is_today else 1, C_WARNING if is_today else (C_SECONDARY if is_sel else C_MUTED)),
                    border_radius=ft.BorderRadius.all(8),
                    on_click=lambda e, d=day: select_day(d),
                    content=ft.Column(
                        spacing=3, alignment=ft.MainAxisAlignment.CENTER,
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Text(str(day), size=13, color=C_WHITE, weight=ft.FontWeight.BOLD if is_sel or is_today else ft.FontWeight.W_500),
                            ft.Row(spacing=2, alignment=ft.MainAxisAlignment.CENTER, controls=dots, height=6),
                        ],
                    ),
                ))
            week_rows.append(ft.Row(spacing=2, controls=cells))

        legend = ft.Row(wrap=True, spacing=10, run_spacing=2, controls=[
            ft.Row(spacing=4, tight=True, controls=[
                ft.Container(width=8, height=8, border_radius=4, bgcolor=MARK_COLORS[k]),
                ft.Text(MARK_LABELS[k], size=10, color=C_TEXT_LIGHT),
            ]) for k in MARK_COLORS
        ])

        return ft.Column(
            expand=True, scroll=ft.ScrollMode.AUTO, spacing=8,
            controls=[header, dow_row, *week_rows, legend, self._build_day_detail()],
        )

    def _build_day_detail(self) -> ft.Control:
        ds = self.cal_selected
        d = datetime.datetime.strptime(ds, "%Y-%m-%d").date()
        data = self.db.get_day_activity(ds)

        def del_tx(tx_id):
            self.db.delete_transaction(tx_id)
            self.render()

        def del_act(act_id):
            self.db.delete_calendar_activity(act_id)
            self.render()

        def del_log(log_id):
            self.db.delete_reading_log(log_id)
            self.render()

        sections = []

        # --- Keuangan
        tx = data["transactions"]
        if tx:
            income = sum(t["total"] for t in tx if t["tipe"] == "PENDAPATAN")
            spend = sum(t["total"] for t in tx if t["tipe"] != "PENDAPATAN")
            sections.append(ft.Text(f"💰 Keuangan  (masuk {_rp(income)} • keluar {_rp(spend)})",
                                    size=12, weight=ft.FontWeight.BOLD, color=C_SUCCESS))
            for t in tx:
                inc = t["tipe"] == "PENDAPATAN"
                sections.append(self._bc_row_item(
                    ft.Icons.ARROW_DOWNWARD if inc else ft.Icons.ARROW_UPWARD,
                    C_SUCCESS if inc else C_DANGER,
                    t["deskripsi"] or t["kategori"],
                    f"{t['kategori']} • {_rp(t['total'])}",
                    on_delete=lambda ev, i=t["id"]: del_tx(i),
                ))

        # --- Baca buku
        reading = data["reading"]
        if reading:
            total_pages = sum(r["pages_read"] for r in reading)
            sections.append(ft.Text(f"📖 Baca buku  (+{total_pages} halaman)", size=12, weight=ft.FontWeight.BOLD, color=C_BLUE))
            for r in reading:
                sections.append(self._bc_row_item(
                    ft.Icons.MENU_BOOK_ROUNDED, C_BLUE, r["judul"],
                    f"Halaman {r['from_page']} → {r['to_page']} (+{r['pages_read']})" + (f" • {r['notes']}" if r["notes"] else ""),
                    on_delete=lambda ev, i=r["id"]: del_log(i),
                ))

        # --- Kosakata
        vocab, quiz = data["vocab"], data["quiz"]
        if vocab or quiz:
            sections.append(ft.Text("🔤 Kosakata", size=12, weight=ft.FontWeight.BOLD, color=C_SECONDARY))
            labels = {"memorized": "kata dihafal", "learning": "kata dipelajari", "flashcard": "flashcard dibuka",
                      "quiz_correct": "jawaban benar", "quiz_wrong": "jawaban salah"}
            for k, c in vocab.items():
                sections.append(self._bc_row_item(ft.Icons.TRANSLATE, C_SECONDARY, f"{c} {labels.get(k, k)}"))
            for q in quiz:
                sections.append(self._bc_row_item(
                    ft.Icons.QUIZ_OUTLINED, C_SECONDARY, f"Latihan {q['quiz_type']}",
                    f"{q['correct_answers']}/{q['total_questions']} benar • {int(q['score_percent'])}%"))

        # --- Kegiatan manual
        acts = data["activities"]
        if acts:
            mins = sum(a["durasi_menit"] or 0 for a in acts)
            sections.append(ft.Text(f"📝 Kegiatan lain" + (f"  ({mins} menit)" if mins else ""),
                                    size=12, weight=ft.FontWeight.BOLD, color=C_WARNING))
            for a in acts:
                sub = a["kategori"] + (f" • {a['durasi_menit']} menit" if a["durasi_menit"] else "") + (f" • {a['notes']}" if a["notes"] else "")
                sections.append(self._bc_row_item(ft.Icons.EDIT_NOTE, C_WARNING, a["judul"], sub,
                                                  on_delete=lambda ev, i=a["id"]: del_act(i)))

        if not sections:
            sections.append(ft.Text("Belum ada aktivitas di hari ini.", size=12, color=C_TEXT_LIGHT))

        return ft.Container(
            bgcolor=C_CARD_DARK, border=ft.Border.all(1, C_MUTED), border_radius=ft.BorderRadius.all(14),
            padding=ft.Padding.all(12), margin=ft.Margin.only(bottom=8),
            content=ft.Column(spacing=6, controls=[
                ft.Text(f"{DAY_FULL_ID[d.weekday()]}, {d.day} {MONTHS_ID[d.month]} {d.year}",
                        size=15, weight=ft.FontWeight.BOLD, color=C_WHITE),
                ft.Row(wrap=True, spacing=6, run_spacing=4, controls=[
                    ft.Button(content="+ Kegiatan", bgcolor=C_PRIMARY, color=C_WHITE, on_click=lambda e: self._open_activity_dialog()),
                    ft.Button(content="+ Baca", bgcolor=C_PRIMARY, color=C_WHITE,
                              on_click=lambda e: self._open_reading_dialog(None, ds)),
                    ft.Button(content="+ Transaksi", bgcolor=C_PRIMARY, color=C_WHITE,
                              on_click=lambda e: self._go_add_transaction()),
                ]),
                *sections,
            ]),
        )

    def _go_add_transaction(self):
        """Pindah ke tab Keuangan > Transaksi agar bisa input transaksi."""
        self.selected_year = self.cal_year
        self.selected_month = self.cal_month
        self.finance_subtab = 0
        self.main_tab_index = 3
        self.nav_bar.selected_index = 3
        self.render()
