import os
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Optional


# User's requested color palette (without #)
HEX_BG_DARK = "0E0B1D"
HEX_CARD_DARK = "2A1E4D"
HEX_PRIMARY = "5A3B8C"
HEX_SECONDARY = "8B7CC7"
HEX_TEXT_LIGHT = "C9C2E3"
HEX_MUTED = "3A3A52"
HEX_ROW_ALT = "F4F2FA"
HEX_ROW_WHITE = "FFFFFF"


def generate_oxford_excel(vocab_list: List[Dict[str, Any]], output_path: str) -> str:
    """
    Membuat file Excel dari daftar kosakata American Oxford 5000 dengan 3 kolom utama:
    1. Kata Bahasa Inggris
    2. Arti
    3. Penjelasan
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kosakata Oxford 5000"

    # Styles
    header_fill = PatternFill(start_color=HEX_CARD_DARK, end_color=HEX_CARD_DARK, fill_type="solid")
    header_font = Font(name="Segoe UI", size=12, bold=True, color=HEX_TEXT_LIGHT)
    title_fill = PatternFill(start_color=HEX_BG_DARK, end_color=HEX_BG_DARK, fill_type="solid")
    title_font = Font(name="Segoe UI", size=14, bold=True, color=HEX_TEXT_LIGHT)
    subtitle_font = Font(name="Segoe UI", size=10, italic=True, color=HEX_SECONDARY)

    row_fill_even = PatternFill(start_color=HEX_ROW_ALT, end_color=HEX_ROW_ALT, fill_type="solid")
    row_fill_odd = PatternFill(start_color=HEX_ROW_WHITE, end_color=HEX_ROW_WHITE, fill_type="solid")

    word_font = Font(name="Segoe UI", size=11, bold=True, color=HEX_CARD_DARK)
    arti_font = Font(name="Segoe UI", size=11, bold=False, color="1A1A2E")
    penjelasan_font = Font(name="Segoe UI", size=10, bold=False, color="2B2B3D")

    thin_side = Side(border_style="thin", color=HEX_SECONDARY)
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Title Banner (Rows 1-2)
    ws.merge_cells("A1:C1")
    ws["A1"] = "AMERICAN OXFORD 5000™ — DAFTAR KOSAKATA BAHASA INGGRIS, ARTI & PENJELASAN"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:C2")
    ws["A2"] = f"Total {len(vocab_list)} Kosakata Inti (Level B2–C1) | Kolom 1: Kata Bahasa Inggris | Kolom 2: Arti | Kolom 3: Penjelasan"
    ws["A2"].font = subtitle_font
    ws["A2"].fill = title_fill
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # Headers (Row 3)
    headers = ["Kata Bahasa Inggris", "Arti", "Penjelasan"]
    ws.row_dimensions[3].height = 26
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border

    # Populate Data Rows
    for idx, item in enumerate(vocab_list, start=1):
        r = idx + 3
        word_str = item.get("word", "").strip()
        pos_lvl = item.get("pos_level", "").strip()
        full_word_col = f"{word_str} ({pos_lvl})" if pos_lvl else word_str
        arti_col = item.get("arti", "").strip()
        penjelasan_col = item.get("penjelasan", "").strip()

        c1 = ws.cell(row=r, column=1, value=full_word_col)
        c2 = ws.cell(row=r, column=2, value=arti_col)
        c3 = ws.cell(row=r, column=3, value=penjelasan_col)

        row_fill = row_fill_even if idx % 2 == 0 else row_fill_odd

        c1.font = word_font
        c2.font = arti_font
        c3.font = penjelasan_font

        c1.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c2.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c3.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        for c in (c1, c2, c3):
            c.fill = row_fill
            c.border = cell_border

        ws.row_dimensions[r].height = 36

    # Column widths
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 75

    # Freeze Header & AutoFilter
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:C{len(vocab_list) + 3}"

    wb.save(output_path)
    return output_path


def export_full_app_report(db_manager, output_path: str) -> str:
    """
    Mengekspor laporan lengkap dari aplikasi (Progress Hafalan Kosakata + Transaksi & Overview Keuangan)
    ke dalam file Excel bertema ungu.
    """
    wb = openpyxl.Workbook()

    header_fill = PatternFill(start_color=HEX_PRIMARY, end_color=HEX_PRIMARY, fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

    # Sheet 1: Progress Hafalan Kosakata
    ws_vocab = wb.active
    ws_vocab.title = "Progress Kosakata"
    v_headers = ["No", "Kata Bahasa Inggris", "Kelas & Level", "Arti", "Penjelasan", "Status Hafalan", "Tanggal Hafal", "Skor Latihan"]
    for c_i, h in enumerate(v_headers, 1):
        cell = ws_vocab.cell(1, c_i, h)
        cell.fill = header_fill
        cell.font = header_font

    rows, _ = db_manager.query_vocab(limit=5000, offset=0)
    for r_i, item in enumerate(rows, start=2):
        ws_vocab.cell(r_i, 1, item["id"])
        ws_vocab.cell(r_i, 2, item["word"])
        ws_vocab.cell(r_i, 3, item["pos_level"])
        ws_vocab.cell(r_i, 4, item["arti"])
        ws_vocab.cell(r_i, 5, item["penjelasan"])
        ws_vocab.cell(r_i, 6, item["status"].upper())
        ws_vocab.cell(r_i, 7, item["memorized_date"] or "-")
        q_str = f"{item['times_correct']}/{item['times_quizzed']}" if item["times_quizzed"] else "0/0"
        ws_vocab.cell(r_i, 8, q_str)

    ws_vocab.column_dimensions["B"].width = 25
    ws_vocab.column_dimensions["D"].width = 35
    ws_vocab.column_dimensions["E"].width = 60
    ws_vocab.freeze_panes = "A2"

    # Sheet 2: Transaksi Keuangan
    ws_tx = wb.create_sheet("Transaksi Keuangan")
    tx_headers = ["Tanggal", "Jenis Transaksi", "Kategori", "Deskripsi", "Total (Rp)", "Notes"]
    for c_i, h in enumerate(tx_headers, 1):
        cell = ws_tx.cell(1, c_i, h)
        cell.fill = header_fill
        cell.font = header_font

    txs = db_manager.get_transactions()
    for r_i, tx in enumerate(txs, start=2):
        ws_tx.cell(r_i, 1, tx["tanggal"])
        ws_tx.cell(r_i, 2, tx["tipe"])
        ws_tx.cell(r_i, 3, tx["kategori"])
        ws_tx.cell(r_i, 4, tx["deskripsi"])
        c_tot = ws_tx.cell(r_i, 5, tx["total"])
        c_tot.number_format = "#,##0"
        ws_tx.cell(r_i, 6, tx["notes"])

    for col_l, w in [("A", 15), ("B", 24), ("C", 24), ("D", 32), ("E", 18), ("F", 30)]:
        ws_tx.column_dimensions[col_l].width = w
    ws_tx.freeze_panes = "A2"

    # Sheet 3: Tracker Langganan
    ws_sub = wb.create_sheet("Tracker Langganan")
    sub_headers = ["Nama Langganan", "Kategori", "Billing", "Auto Renew", "Tanggal Renew", "Harga Asli", "Harga Bulanan", "Harga Tahunan", "Notes"]
    for c_i, h in enumerate(sub_headers, 1):
        cell = ws_sub.cell(1, c_i, h)
        cell.fill = header_fill
        cell.font = header_font

    subs = db_manager.get_subscriptions()
    for r_i, s in enumerate(subs["items"], start=2):
        ws_sub.cell(r_i, 1, s["nama"])
        ws_sub.cell(r_i, 2, s["kategori"])
        ws_sub.cell(r_i, 3, s["billing"])
        ws_sub.cell(r_i, 4, "Ya" if s["auto_renew"] else "Tidak")
        ws_sub.cell(r_i, 5, s["tanggal_renew"] or "-")
        ws_sub.cell(r_i, 6, s["harga"]).number_format = "#,##0"
        ws_sub.cell(r_i, 7, s["harga_bulanan"]).number_format = "#,##0"
        ws_sub.cell(r_i, 8, s["harga_tahunan"]).number_format = "#,##0"
        ws_sub.cell(r_i, 9, s["notes"] or "")

    wb.save(output_path)
    return output_path

