import sqlite3
import json
import os
import datetime
from typing import List, Dict, Any, Optional, Tuple

DB_FILENAME = "mylife_app.db"


def get_db_path() -> str:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, DB_FILENAME)


def get_vocab_json_path() -> str:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "oxford_5000_vocab.json")


class DatabaseManager:
    def __init__(self, db_path: Optional[str] = None):
        self.db_path = db_path or get_db_path()
        self.init_db()

    def get_conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def init_db(self):
        conn = self.get_conn()
        cur = conn.cursor()

        # 1. Vocabulary Table
        cur.execute("""
            CREATE TABLE IF NOT EXISTS vocab (
                id INTEGER PRIMARY KEY,
                word TEXT NOT NULL,
                pos_level TEXT,
                arti TEXT NOT NULL,
                penjelasan TEXT NOT NULL,
                status TEXT DEFAULT 'belum', -- 'belum', 'belajar', 'hafal'
                memorized_date TEXT,         -- YYYY-MM-DD when marked as 'hafal'
                last_reviewed TEXT,          -- YYYY-MM-DD
                correct_streak INTEGER DEFAULT 0,
                times_quizzed INTEGER DEFAULT 0,
                times_correct INTEGER DEFAULT 0,
                is_favorite INTEGER DEFAULT 0
            )
        """)

        # 2. Vocabulary Study / Activity Log (for accurate monthly & yearly tracking)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS vocab_activity_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                word_id INTEGER,
                action_type TEXT NOT NULL, -- 'memorized', 'learning', 'quiz_correct', 'quiz_wrong', 'flashcard'
                log_date TEXT NOT NULL,    -- YYYY-MM-DD
                score INTEGER DEFAULT 0,
                FOREIGN KEY(word_id) REFERENCES vocab(id)
            )
        """)

        # 3. Quiz Sessions Table
        cur.execute("""
            CREATE TABLE IF NOT EXISTS quiz_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                quiz_date TEXT NOT NULL,   -- YYYY-MM-DD
                quiz_type TEXT NOT NULL,   -- 'pilihan_ganda', 'tebak_kata', 'arti_ke_inggris'
                source_filter TEXT NOT NULL, -- 'hafal', 'belajar_hafal', 'semua'
                total_questions INTEGER NOT NULL,
                correct_answers INTEGER NOT NULL,
                score_percent REAL NOT NULL
            )
        """)

        # 4. Finance Settings (Budget Ratio 50/30/20 or 70/20/10 & Monthly Vocab Target)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)

        # 5. Finance Categories (SET-UP sheet equivalent)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS finance_categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tipe TEXT NOT NULL,         -- 'PENDAPATAN', 'KEBUTUHAN', 'KEINGINAN', 'TABUNGAN & INVESTASI'
                kategori TEXT NOT NULL,
                jatuh_tempo INTEGER,        -- Day of month (1-31) or NULL
                target REAL DEFAULT 0.0
            )
        """)

        # 6. Finance Transactions (TRANSAKSI sheet equivalent)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS finance_transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal TEXT NOT NULL,      -- YYYY-MM-DD
                tipe TEXT NOT NULL,         -- 'PENDAPATAN', 'KEBUTUHAN', 'KEINGINAN', 'TABUNGAN & INVESTASI'
                kategori TEXT NOT NULL,
                deskripsi TEXT,
                total REAL NOT NULL,
                notes TEXT
            )
        """)

        # 7. Subscriptions (TRACKER LANGGANAN sheet equivalent)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS finance_subscriptions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nama TEXT NOT NULL,
                kategori TEXT NOT NULL,     -- 'Entertainment', 'Productivity', 'Belanja', 'Utilities', etc.
                billing TEXT NOT NULL,      -- 'Monthly' or 'Annually'
                auto_renew INTEGER DEFAULT 1,
                tanggal_renew INTEGER,      -- Day of month (1-31)
                harga REAL NOT NULL,
                notes TEXT
            )
        """)

        # 8. Saving Goals (Extra value-add feature)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS finance_goals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nama_goal TEXT NOT NULL,
                target_nominal REAL NOT NULL,
                terkumpul_nominal REAL DEFAULT 0.0,
                deadline TEXT,
                notes TEXT
            )
        """)

        conn.commit()
        self._seed_defaults(conn)
        conn.close()

    def _seed_defaults(self, conn: sqlite3.Connection):
        cur = conn.cursor()

        # Default Settings
        defaults = {
            "budget_mode": "50/30/20",
            "pct_kebutuhan": "50",
            "pct_keinginan": "30",
            "pct_tabungan": "20",
            "vocab_monthly_target": "100",
            "vocab_daily_target": "5",
        }
        for k, v in defaults.items():
            cur.execute("INSERT OR IGNORE INTO app_settings (key, value) VALUES (?, ?)", (k, v))

        # Seed Finance Categories from template_keuangan.xlsx if empty
        cur.execute("SELECT COUNT(*) FROM finance_categories")
        if cur.fetchone()[0] == 0:
            initial_categories = [
                # PENDAPATAN
                ("PENDAPATAN", "Pendapatan 1 (Gaji Utama)", None, 8000000.0),
                ("PENDAPATAN", "Pendapatan 2 (Sampingan)", None, 500000.0),
                ("PENDAPATAN", "Bonus & Lainnya", None, 0.0),
                # KEBUTUHAN
                ("KEBUTUHAN", "Biaya sewa", 1, 1750000.0),
                ("KEBUTUHAN", "Listrik, Internet", 5, 500000.0),
                ("KEBUTUHAN", "Makan", None, 1500000.0),
                ("KEBUTUHAN", "Transport", None, 300000.0),
                ("KEBUTUHAN", "Kartu Kredit", 20, 200000.0),
                # KEINGINAN
                ("KEINGINAN", "Belanja", None, 750000.0),
                ("KEINGINAN", "Langganan", None, 150000.0),
                ("KEINGINAN", "Staycation", None, 650000.0),
                ("KEINGINAN", "Hobby", None, 200000.0),
                ("KEINGINAN", "Personal Care", None, 300000.0),
                ("KEINGINAN", "Makan di Luar", None, 500000.0),
                # TABUNGAN & INVESTASI
                ("TABUNGAN & INVESTASI", "Emergency Fund", 5, 1200000.0),
                ("TABUNGAN & INVESTASI", "Investments", None, 500000.0),
            ]
            cur.executemany(
                "INSERT INTO finance_categories (tipe, kategori, jatuh_tempo, target) VALUES (?, ?, ?, ?)",
                initial_categories,
            )

        # Seed sample transactions from template_keuangan.xlsx (adjusted to current year/month so charts look alive immediately)
        cur.execute("SELECT COUNT(*) FROM finance_transactions")
        if cur.fetchone()[0] == 0:
            today = datetime.date.today()
            ym = today.strftime("%Y-%m")
            sample_tx = [
                (f"{ym}-01", "PENDAPATAN", "Pendapatan 1 (Gaji Utama)", "Main income", 8000000.0, "Gaji bulanan"),
                (f"{ym}-02", "KEBUTUHAN", "Biaya sewa", "Apartment Rent", 1750000.0, "Sewa tempat tinggal"),
                (f"{ym}-05", "KEINGINAN", "Langganan", "Netflix", 54000.0, "Langganan bulanan"),
                (f"{ym}-06", "TABUNGAN & INVESTASI", "Emergency Fund", "Dana Darurat", 1000000.0, "Alokasi rutin"),
                (f"{ym}-07", "KEINGINAN", "Personal Care", "Skincare, make-up", 250000.0, "Kebutuhan perawatan"),
            ]
            cur.executemany(
                "INSERT INTO finance_transactions (tanggal, tipe, kategori, deskripsi, total, notes) VALUES (?, ?, ?, ?, ?, ?)",
                sample_tx,
            )

        # Seed sample subscriptions from template_keuangan.xlsx
        cur.execute("SELECT COUNT(*) FROM finance_subscriptions")
        if cur.fetchone()[0] == 0:
            sample_subs = [
                ("Netflix", "Entertainment", "Monthly", 1, 5, 54000.0, "Paket Basic"),
                ("Spotify", "Entertainment", "Monthly", 1, 10, 49000.0, "Premium Individual"),
                ("GoodNotes", "Productivity", "Annually", 1, 15, 169000.0, "Langganan Tahunan"),
            ]
            cur.executemany(
                "INSERT INTO finance_subscriptions (nama, kategori, billing, auto_renew, tanggal_renew, harga, notes) VALUES (?, ?, ?, ?, ?, ?, ?)",
                sample_subs,
            )

        # Seed sample saving goals
        cur.execute("SELECT COUNT(*) FROM finance_goals")
        if cur.fetchone()[0] == 0:
            sample_goals = [
                ("Dana Darurat Ideal", 15000000.0, 4500000.0, "2027-06-01", "6x pengeluaran kebutuhan"),
                ("Upgrade Laptop Kerja", 12000000.0, 3000000.0, "2027-03-01", "Tabungan produktivitas"),
            ]
            cur.executemany(
                "INSERT INTO finance_goals (nama_goal, target_nominal, terkumpul_nominal, deadline, notes) VALUES (?, ?, ?, ?, ?)",
                sample_goals,
            )

        conn.commit()
        self.load_vocab_if_needed(conn)

    def load_vocab_if_needed(self, conn: Optional[sqlite3.Connection] = None):
        close_after = False
        if conn is None:
            conn = self.get_conn()
            close_after = True

        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM vocab")
        count = cur.fetchone()[0]

        vocab_json = get_vocab_json_path()
        if count < 2000 and os.path.exists(vocab_json):
            try:
                with open(vocab_json, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if data:
                    cur.executemany(
                        """
                        INSERT OR REPLACE INTO vocab (
                            id, word, pos_level, arti, penjelasan,
                            status, memorized_date, last_reviewed,
                            correct_streak, times_quizzed, times_correct, is_favorite
                        )
                        VALUES (
                            :id, :word, :pos_level, :arti, :penjelasan,
                            COALESCE((SELECT status FROM vocab WHERE id = :id), 'belum'),
                            (SELECT memorized_date FROM vocab WHERE id = :id),
                            (SELECT last_reviewed FROM vocab WHERE id = :id),
                            COALESCE((SELECT correct_streak FROM vocab WHERE id = :id), 0),
                            COALESCE((SELECT times_quizzed FROM vocab WHERE id = :id), 0),
                            COALESCE((SELECT times_correct FROM vocab WHERE id = :id), 0),
                            COALESCE((SELECT is_favorite FROM vocab WHERE id = :id), 0)
                        )
                        """,
                        data,
                    )
                    conn.commit()
            except Exception as e:
                print("Warning loading vocab JSON:", e)

        if close_after:
            conn.close()

    # =========================================================================
    # VOCABULARY METHODS
    # =========================================================================
    def get_vocab_summary(self) -> Dict[str, Any]:
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM vocab")
        total = cur.fetchone()[0]
        cur.execute("SELECT status, COUNT(*) as cnt FROM vocab GROUP BY status")
        by_status = {row["status"]: row["cnt"] for row in cur.fetchall()}
        cur.execute("SELECT COUNT(*) FROM vocab WHERE is_favorite = 1")
        favs = cur.fetchone()[0]

        today = datetime.date.today().strftime("%Y-%m-%d")
        cur.execute("SELECT COUNT(DISTINCT word_id) FROM vocab_activity_log WHERE log_date = ? AND action_type = 'memorized'", (today,))
        today_memorized = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM quiz_sessions")
        total_quizzes = cur.fetchone()[0]
        cur.execute("SELECT AVG(score_percent) FROM quiz_sessions")
        avg_score = cur.fetchone()[0] or 0.0

        conn.close()
        return {
            "total": total,
            "hafal": by_status.get("hafal", 0),
            "belajar": by_status.get("belajar", 0),
            "belum": by_status.get("belum", 0),
            "favorites": favs,
            "today_memorized": today_memorized,
            "total_quizzes": total_quizzes,
            "avg_quiz_score": round(avg_score, 1),
        }

    def query_vocab(
        self,
        search: str = "",
        status_filter: str = "semua",
        level_filter: str = "semua",
        fav_only: bool = False,
        limit: int = 100,
        offset: int = 0,
    ) -> Tuple[List[Dict[str, Any]], int]:
        conn = self.get_conn()
        cur = conn.cursor()

        where_clauses = []
        params = []

        if search.strip():
            q = f"%{search.strip()}%"
            where_clauses.append("(word LIKE ? OR arti LIKE ? OR penjelasan LIKE ?)")
            params.extend([q, q, q])

        if status_filter != "semua":
            where_clauses.append("status = ?")
            params.append(status_filter)

        if level_filter != "semua":
            where_clauses.append("pos_level LIKE ?")
            params.append(f"%{level_filter}%")

        if fav_only:
            where_clauses.append("is_favorite = 1")

        where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""

        cur.execute(f"SELECT COUNT(*) FROM vocab {where_sql}", params)
        total_matching = cur.fetchone()[0]

        cur.execute(
            f"SELECT * FROM vocab {where_sql} ORDER BY id ASC LIMIT ? OFFSET ?",
            params + [limit, offset],
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows, total_matching

    def update_vocab_status(self, word_id: int, new_status: str, custom_date: Optional[str] = None):
        today = custom_date or datetime.date.today().strftime("%Y-%m-%d")
        conn = self.get_conn()
        cur = conn.cursor()
        if new_status == "hafal":
            cur.execute(
                "UPDATE vocab SET status = ?, memorized_date = COALESCE(memorized_date, ?), last_reviewed = ? WHERE id = ?",
                (new_status, today, today, word_id),
            )
            cur.execute(
                "INSERT INTO vocab_activity_log (word_id, action_type, log_date, score) VALUES (?, 'memorized', ?, 100)",
                (word_id, today),
            )
        elif new_status == "belajar":
            cur.execute(
                "UPDATE vocab SET status = ?, last_reviewed = ? WHERE id = ?",
                (new_status, today, word_id),
            )
            cur.execute(
                "INSERT INTO vocab_activity_log (word_id, action_type, log_date, score) VALUES (?, 'learning', ?, 50)",
                (word_id, today),
            )
        else:
            cur.execute(
                "UPDATE vocab SET status = 'belum', memorized_date = NULL, last_reviewed = ? WHERE id = ?",
                (today, word_id),
            )
        conn.commit()
        conn.close()

    def toggle_favorite(self, word_id: int) -> int:
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("SELECT is_favorite FROM vocab WHERE id = ?", (word_id,))
        row = cur.fetchone()
        new_val = 0 if (row and row["is_favorite"]) else 1
        cur.execute("UPDATE vocab SET is_favorite = ? WHERE id = ?", (new_val, word_id))
        conn.commit()
        conn.close()
        return new_val

    def add_custom_vocab(self, word: str, pos_level: str, arti: str, penjelasan: str):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("SELECT COALESCE(MAX(id), 2000) + 1 FROM vocab")
        next_id = cur.fetchone()[0]
        cur.execute(
            "INSERT INTO vocab (id, word, pos_level, arti, penjelasan, status) VALUES (?, ?, ?, ?, ?, 'belum')",
            (next_id, word.strip(), pos_level.strip(), arti.strip(), penjelasan.strip()),
        )
        conn.commit()
        conn.close()

    def get_quiz_words(self, source_filter: str = "hafal", count: int = 10) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """Returns (selected_quiz_words, distractor_pool)"""
        conn = self.get_conn()
        cur = conn.cursor()

        if source_filter == "hafal":
            cur.execute("SELECT * FROM vocab WHERE status = 'hafal' ORDER BY RANDOM() LIMIT ?", (count,))
        elif source_filter == "belajar_hafal":
            cur.execute("SELECT * FROM vocab WHERE status IN ('hafal', 'belajar') ORDER BY RANDOM() LIMIT ?", (count,))
        else:
            cur.execute("SELECT * FROM vocab ORDER BY RANDOM() LIMIT ?", (count,))

        selected = [dict(r) for r in cur.fetchall()]
        cur.execute("SELECT id, word, arti, pos_level FROM vocab ORDER BY RANDOM() LIMIT 100")
        distractors = [dict(r) for r in cur.fetchall()]
        conn.close()
        return selected, distractors

    def record_quiz_answer(self, word_id: int, is_correct: bool):
        today = datetime.date.today().strftime("%Y-%m-%d")
        conn = self.get_conn()
        cur = conn.cursor()
        if is_correct:
            cur.execute(
                """
                UPDATE vocab
                SET times_quizzed = times_quizzed + 1,
                    times_correct = times_correct + 1,
                    correct_streak = correct_streak + 1,
                    last_reviewed = ?
                WHERE id = ?
                """,
                (today, word_id),
            )
            cur.execute(
                "INSERT INTO vocab_activity_log (word_id, action_type, log_date, score) VALUES (?, 'quiz_correct', ?, 100)",
                (word_id, today),
            )
        else:
            cur.execute(
                """
                UPDATE vocab
                SET times_quizzed = times_quizzed + 1,
                    correct_streak = 0,
                    last_reviewed = ?
                WHERE id = ?
                """,
                (today, word_id),
            )
            cur.execute(
                "INSERT INTO vocab_activity_log (word_id, action_type, log_date, score) VALUES (?, 'quiz_wrong', ?, 0)",
                (word_id, today),
            )
        conn.commit()
        conn.close()

    def record_quiz_session(self, quiz_type: str, source_filter: str, total_q: int, correct_q: int):
        today = datetime.date.today().strftime("%Y-%m-%d")
        pct = (correct_q / total_q * 100.0) if total_q > 0 else 0.0
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO quiz_sessions (quiz_date, quiz_type, source_filter, total_questions, correct_answers, score_percent)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (today, quiz_type, source_filter, total_q, correct_q, pct),
        )
        conn.commit()
        conn.close()

    def get_vocab_monthly_tracker(self, year: int, month: int) -> Dict[str, Any]:
        ym = f"{year:04d}-{month:02d}"
        conn = self.get_conn()
        cur = conn.cursor()

        # Words memorized in this month
        cur.execute(
            "SELECT COUNT(*) FROM vocab WHERE status = 'hafal' AND memorized_date LIKE ?",
            (f"{ym}%",),
        )
        memorized_in_month = cur.fetchone()[0]

        # Daily breakdown of memorized words in this month
        cur.execute(
            """
            SELECT memorized_date as dt, COUNT(*) as cnt
            FROM vocab
            WHERE status = 'hafal' AND memorized_date LIKE ?
            GROUP BY memorized_date
            ORDER BY memorized_date ASC
            """,
            (f"{ym}%",),
        )
        daily_memorized = {r["dt"]: r["cnt"] for r in cur.fetchall() if r["dt"]}

        # Quizzes in this month
        cur.execute(
            """
            SELECT COUNT(*) as sessions,
                   COALESCE(SUM(total_questions), 0) as total_q,
                   COALESCE(SUM(correct_answers), 0) as correct_q,
                   COALESCE(AVG(score_percent), 0) as avg_pct
            FROM quiz_sessions
            WHERE quiz_date LIKE ?
            """,
            (f"{ym}%",),
        )
        q_row = dict(cur.fetchone())

        # List of recently memorized words in this month
        cur.execute(
            """
            SELECT id, word, pos_level, arti, memorized_date, times_quizzed, times_correct
            FROM vocab
            WHERE status = 'hafal' AND memorized_date LIKE ?
            ORDER BY memorized_date DESC, id ASC
            LIMIT 50
            """,
            (f"{ym}%",),
        )
        words_list = [dict(r) for r in cur.fetchall()]

        target_str = self.get_setting("vocab_monthly_target", "100")
        try:
            monthly_target = int(target_str)
        except ValueError:
            monthly_target = 100

        conn.close()
        return {
            "year": year,
            "month": month,
            "ym": ym,
            "memorized_count": memorized_in_month,
            "monthly_target": monthly_target,
            "daily_memorized": daily_memorized,
            "quiz_sessions": q_row["sessions"],
            "quiz_total_q": q_row["total_q"],
            "quiz_correct_q": q_row["correct_q"],
            "quiz_avg_pct": round(q_row["avg_pct"], 1),
            "words_list": words_list,
        }

    def get_vocab_yearly_tracker(self, year: int) -> Dict[str, Any]:
        y_str = f"{year:04d}"
        conn = self.get_conn()
        cur = conn.cursor()

        months_data = []
        total_year_memorized = 0
        total_year_quizzes = 0

        for m in range(1, 13):
            ym = f"{year:04d}-{m:02d}"
            cur.execute(
                "SELECT COUNT(*) FROM vocab WHERE status = 'hafal' AND memorized_date LIKE ?",
                (f"{ym}%",),
            )
            mem_cnt = cur.fetchone()[0]

            cur.execute(
                """
                SELECT COUNT(*) as sessions,
                       COALESCE(AVG(score_percent), 0) as avg_pct,
                       COALESCE(SUM(total_questions), 0) as total_q
                FROM quiz_sessions
                WHERE quiz_date LIKE ?
                """,
                (f"{ym}%",),
            )
            qr = cur.fetchone()
            total_year_memorized += mem_cnt
            total_year_quizzes += qr["sessions"]
            months_data.append({
                "month": m,
                "ym": ym,
                "memorized": mem_cnt,
                "quiz_sessions": qr["sessions"],
                "quiz_questions": qr["total_q"],
                "avg_score": round(qr["avg_pct"], 1),
            })

        # Breakdown by CEFR level (B2 vs C1)
        cur.execute("SELECT COUNT(*) FROM vocab WHERE status = 'hafal' AND pos_level LIKE '%B2%'")
        b2_hafal = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM vocab WHERE pos_level LIKE '%B2%'")
        b2_total = cur.fetchone()[0]

        cur.execute("SELECT COUNT(*) FROM vocab WHERE status = 'hafal' AND pos_level LIKE '%C1%'")
        c1_hafal = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM vocab WHERE pos_level LIKE '%C1%'")
        c1_total = cur.fetchone()[0]

        conn.close()
        return {
            "year": year,
            "total_year_memorized": total_year_memorized,
            "total_year_quizzes": total_year_quizzes,
            "months": months_data,
            "b2_hafal": b2_hafal,
            "b2_total": b2_total,
            "c1_hafal": c1_hafal,
            "c1_total": c1_total,
        }

    # =========================================================================
    # SETTINGS & FINANCE METHODS
    # =========================================================================
    def get_setting(self, key: str, default: str = "") -> str:
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("SELECT value FROM app_settings WHERE key = ?", (key,))
        row = cur.fetchone()
        conn.close()
        return row["value"] if row else default

    def set_setting(self, key: str, value: str):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)", (key, str(value)))
        conn.commit()
        conn.close()

    def get_finance_categories(self, tipe: Optional[str] = None) -> List[Dict[str, Any]]:
        conn = self.get_conn()
        cur = conn.cursor()
        if tipe:
            cur.execute("SELECT * FROM finance_categories WHERE tipe = ? ORDER BY id ASC", (tipe,))
        else:
            cur.execute("SELECT * FROM finance_categories ORDER BY tipe, id ASC")
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def save_finance_category(self, tipe: str, kategori: str, jatuh_tempo: Optional[int], target: float, cat_id: Optional[int] = None):
        conn = self.get_conn()
        cur = conn.cursor()
        if cat_id:
            cur.execute(
                "UPDATE finance_categories SET tipe = ?, kategori = ?, jatuh_tempo = ?, target = ? WHERE id = ?",
                (tipe, kategori.strip(), jatuh_tempo, target, cat_id),
            )
        else:
            cur.execute(
                "INSERT INTO finance_categories (tipe, kategori, jatuh_tempo, target) VALUES (?, ?, ?, ?)",
                (tipe, kategori.strip(), jatuh_tempo, target),
            )
        conn.commit()
        conn.close()

    def delete_finance_category(self, cat_id: int):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("DELETE FROM finance_categories WHERE id = ?", (cat_id,))
        conn.commit()
        conn.close()

    def add_transaction(self, tanggal: str, tipe: str, kategori: str, deskripsi: str, total: float, notes: str = ""):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO finance_transactions (tanggal, tipe, kategori, deskripsi, total, notes)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (tanggal, tipe, kategori, deskripsi.strip(), total, notes.strip()),
        )
        conn.commit()
        conn.close()

    def delete_transaction(self, tx_id: int):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("DELETE FROM finance_transactions WHERE id = ?", (tx_id,))
        conn.commit()
        conn.close()

    def update_transaction(self, tx_id: int, tanggal: str, tipe: str, kategori: str, deskripsi: str, total: float, notes: str = ""):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute(
            """
            UPDATE finance_transactions
            SET tanggal = ?, tipe = ?, kategori = ?, deskripsi = ?, total = ?, notes = ?
            WHERE id = ?
            """,
            (tanggal, tipe, kategori, deskripsi.strip(), total, notes.strip(), tx_id),
        )
        conn.commit()
        conn.close()

    def get_transactions(self, year: Optional[int] = None, month: Optional[int] = None, tipe_filter: str = "SEMUA") -> List[Dict[str, Any]]:
        conn = self.get_conn()
        cur = conn.cursor()
        clauses = []
        params = []
        if year and month:
            clauses.append("tanggal LIKE ?")
            params.append(f"{year:04d}-{month:02d}%")
        elif year:
            clauses.append("tanggal LIKE ?")
            params.append(f"{year:04d}%")

        if tipe_filter and tipe_filter != "SEMUA":
            clauses.append("tipe = ?")
            params.append(tipe_filter)

        where_sql = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        cur.execute(f"SELECT * FROM finance_transactions {where_sql} ORDER BY tanggal DESC, id DESC", params)
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def get_monthly_finance_overview(self, year: int, month: int) -> Dict[str, Any]:
        ym = f"{year:04d}-{month:02d}"
        conn = self.get_conn()
        cur = conn.cursor()

        # Actual totals by tipe
        cur.execute(
            """
            SELECT tipe, COALESCE(SUM(total), 0) as actual
            FROM finance_transactions
            WHERE tanggal LIKE ?
            GROUP BY tipe
            """,
            (f"{ym}%",),
        )
        actual_by_tipe = {r["tipe"]: r["actual"] for r in cur.fetchall()}

        # Actual totals by (tipe, kategori)
        cur.execute(
            """
            SELECT tipe, kategori, COALESCE(SUM(total), 0) as actual
            FROM finance_transactions
            WHERE tanggal LIKE ?
            GROUP BY tipe, kategori
            """,
            (f"{ym}%",),
        )
        actual_by_cat = {(r["tipe"], r["kategori"]): r["actual"] for r in cur.fetchall()}

        # Categories & Targets from SET-UP
        cur.execute("SELECT * FROM finance_categories ORDER BY id ASC")
        all_cats = [dict(r) for r in cur.fetchall()]

        breakdown_by_tipe = {
            "PENDAPATAN": [],
            "KEBUTUHAN": [],
            "KEINGINAN": [],
            "TABUNGAN & INVESTASI": [],
        }
        target_by_tipe = {
            "PENDAPATAN": 0.0,
            "KEBUTUHAN": 0.0,
            "KEINGINAN": 0.0,
            "TABUNGAN & INVESTASI": 0.0,
        }

        seen_cats = set()
        for c in all_cats:
            t = c["tipe"]
            k = c["kategori"]
            seen_cats.add((t, k))
            act = actual_by_cat.get((t, k), 0.0)
            tgt = c["target"] or 0.0
            target_by_tipe[t] = target_by_tipe.get(t, 0.0) + tgt
            breakdown_by_tipe.setdefault(t, []).append({
                "id": c["id"],
                "kategori": k,
                "jatuh_tempo": c["jatuh_tempo"],
                "target": tgt,
                "actual": act,
            })

        # Include any transaction category not in SET-UP
        for (t, k), act in actual_by_cat.items():
            if (t, k) not in seen_cats:
                breakdown_by_tipe.setdefault(t, []).append({
                    "id": None,
                    "kategori": k,
                    "jatuh_tempo": None,
                    "target": 0.0,
                    "actual": act,
                })

        tot_pendapatan = actual_by_tipe.get("PENDAPATAN", 0.0)
        tot_kebutuhan = actual_by_tipe.get("KEBUTUHAN", 0.0)
        tot_keinginan = actual_by_tipe.get("KEINGINAN", 0.0)
        tot_tabungan = actual_by_tipe.get("TABUNGAN & INVESTASI", 0.0)
        tot_pengeluaran = tot_kebutuhan + tot_keinginan + tot_tabungan
        selisih = tot_pendapatan - tot_pengeluaran

        # Budget allocation rule percentages (50/30/20 or 70/20/10)
        pct_keb_target = float(self.get_setting("pct_kebutuhan", "50"))
        pct_kei_target = float(self.get_setting("pct_keinginan", "30"))
        pct_tab_target = float(self.get_setting("pct_tabungan", "20"))

        pct_keb_actual = (tot_kebutuhan / tot_pendapatan * 100.0) if tot_pendapatan > 0 else 0.0
        pct_kei_actual = (tot_keinginan / tot_pendapatan * 100.0) if tot_pendapatan > 0 else 0.0
        pct_tab_actual = (tot_tabungan / tot_pendapatan * 100.0) if tot_pendapatan > 0 else 0.0

        conn.close()
        return {
            "year": year,
            "month": month,
            "ym": ym,
            "pendapatan": tot_pendapatan,
            "kebutuhan": tot_kebutuhan,
            "keinginan": tot_keinginan,
            "tabungan": tot_tabungan,
            "total_alokasi": tot_pengeluaran,
            "selisih": selisih,
            "target_by_tipe": target_by_tipe,
            "breakdown_by_tipe": breakdown_by_tipe,
            "rule": {
                "mode": self.get_setting("budget_mode", "50/30/20"),
                "kebutuhan_target_pct": pct_keb_target,
                "keinginan_target_pct": pct_kei_target,
                "tabungan_target_pct": pct_tab_target,
                "kebutuhan_actual_pct": round(pct_keb_actual, 1),
                "keinginan_actual_pct": round(pct_kei_actual, 1),
                "tabungan_actual_pct": round(pct_tab_actual, 1),
            },
        }

    def get_yearly_finance_overview(self, year: int) -> Dict[str, Any]:
        conn = self.get_conn()
        cur = conn.cursor()

        months_summary = []
        totals = {
            "PENDAPATAN": 0.0,
            "KEBUTUHAN": 0.0,
            "KEINGINAN": 0.0,
            "TABUNGAN & INVESTASI": 0.0,
        }

        for m in range(1, 13):
            ym = f"{year:04d}-{m:02d}"
            cur.execute(
                """
                SELECT tipe, COALESCE(SUM(total), 0) as actual
                FROM finance_transactions
                WHERE tanggal LIKE ?
                GROUP BY tipe
                """,
                (f"{ym}%",),
            )
            by_t = {r["tipe"]: r["actual"] for r in cur.fetchall()}
            p = by_t.get("PENDAPATAN", 0.0)
            kb = by_t.get("KEBUTUHAN", 0.0)
            ki = by_t.get("KEINGINAN", 0.0)
            tb = by_t.get("TABUNGAN & INVESTASI", 0.0)
            totals["PENDAPATAN"] += p
            totals["KEBUTUHAN"] += kb
            totals["KEINGINAN"] += ki
            totals["TABUNGAN & INVESTASI"] += tb
            months_summary.append({
                "month": m,
                "ym": ym,
                "pendapatan": p,
                "kebutuhan": kb,
                "keinginan": ki,
                "tabungan": tb,
                "selisih": p - (kb + ki + tb),
            })

        # Annual category totals
        cur.execute(
            """
            SELECT tipe, kategori, COALESCE(SUM(total), 0) as total_year
            FROM finance_transactions
            WHERE tanggal LIKE ?
            GROUP BY tipe, kategori
            ORDER BY tipe, total_year DESC
            """,
            (f"{year:04d}%",),
        )
        cat_totals = [dict(r) for r in cur.fetchall()]

        # Annual target for Tabungan & Investasi (12 * monthly target)
        cur.execute("SELECT COALESCE(SUM(target), 0) FROM finance_categories WHERE tipe = 'TABUNGAN & INVESTASI'")
        monthly_tab_target = cur.fetchone()[0]
        annual_tab_target = monthly_tab_target * 12.0

        conn.close()
        selisih_tahun = totals["PENDAPATAN"] - (totals["KEBUTUHAN"] + totals["KEINGINAN"] + totals["TABUNGAN & INVESTASI"])
        return {
            "year": year,
            "totals": totals,
            "selisih_tahun": selisih_tahun,
            "annual_tab_target": annual_tab_target,
            "months": months_summary,
            "category_totals": cat_totals,
        }

    # =========================================================================
    # SUBSCRIPTION TRACKER & SAVING GOALS
    # =========================================================================
    def get_subscriptions(self) -> Dict[str, Any]:
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("SELECT * FROM finance_subscriptions ORDER BY id ASC")
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()

        total_monthly = 0.0
        total_annually = 0.0
        by_category: Dict[str, float] = {}

        enriched = []
        for r in rows:
            harga = float(r["harga"] or 0.0)
            if r["billing"] == "Annually":
                m_cost = harga / 12.0
                y_cost = harga
            else:
                m_cost = harga
                y_cost = harga * 12.0
            r["harga_bulanan"] = m_cost
            r["harga_tahunan"] = y_cost
            total_monthly += m_cost
            total_annually += y_cost
            by_category[r["kategori"]] = by_category.get(r["kategori"], 0.0) + m_cost
            enriched.append(r)

        return {
            "items": enriched,
            "total_bulanan": total_monthly,
            "total_tahunan": total_annually,
            "by_category": by_category,
        }

    def add_subscription(self, nama: str, kategori: str, billing: str, auto_renew: bool, tanggal_renew: Optional[int], harga: float, notes: str = ""):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO finance_subscriptions (nama, kategori, billing, auto_renew, tanggal_renew, harga, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (nama.strip(), kategori.strip(), billing, 1 if auto_renew else 0, tanggal_renew, harga, notes.strip()),
        )
        conn.commit()
        conn.close()

    def delete_subscription(self, sub_id: int):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("DELETE FROM finance_subscriptions WHERE id = ?", (sub_id,))
        conn.commit()
        conn.close()

    def get_saving_goals(self) -> List[Dict[str, Any]]:
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("SELECT * FROM finance_goals ORDER BY id ASC")
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def save_saving_goal(self, nama_goal: str, target_nominal: float, terkumpul_nominal: float, deadline: str = "", notes: str = "", goal_id: Optional[int] = None):
        conn = self.get_conn()
        cur = conn.cursor()
        if goal_id:
            cur.execute(
                "UPDATE finance_goals SET nama_goal = ?, target_nominal = ?, terkumpul_nominal = ?, deadline = ?, notes = ? WHERE id = ?",
                (nama_goal.strip(), target_nominal, terkumpul_nominal, deadline.strip(), notes.strip(), goal_id),
            )
        else:
            cur.execute(
                "INSERT INTO finance_goals (nama_goal, target_nominal, terkumpul_nominal, deadline, notes) VALUES (?, ?, ?, ?, ?)",
                (nama_goal.strip(), target_nominal, terkumpul_nominal, deadline.strip(), notes.strip()),
            )
        conn.commit()
        conn.close()

    def delete_saving_goal(self, goal_id: int):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("DELETE FROM finance_goals WHERE id = ?", (goal_id,))
        conn.commit()
        conn.close()

