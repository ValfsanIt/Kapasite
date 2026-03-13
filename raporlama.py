# -*- coding: utf-8 -*-

from datetime import datetime, timedelta
import io
import zipfile
import base64
import re
import pandas as pd
from dash import html, dcc, Input, Output, no_update
import dash_bootstrap_components as dbc
from dash.exceptions import PreventUpdate

from app import app
from run.agent import ag
import kapasite_data

# Veri tipleri: (etiket, tablo, kapasite_tablo, haftalık/aylık)
DATA_TYPES = [
    ("İhtiyaç Miktarı", "VLFCAPFINALPIVOT", "VLFVARDIYASURE", "weekly"),
    ("Sipariş Miktarı", "VLFCAPFINALSIPARIS", "VLFVARDIYASURE", "weekly"),
    ("Öngörü Miktarı", "VLFCAPFINALOY", "VLFVARDIYASUREAY", "monthly"),
]

# Raporlamada işlenecek cost center'lar. Boş liste = tablodaki tüm cost center'lar.
# Eski kısıtlama: ["CNC FREZE", "CNC TORNA", "MONTAJ"] — artık tümü için rapor üretiliyor.
RAPORLAMA_COSTCENTERS = []



def _build_styles():
    """openpyxl stil nesnelerini tek sözlükte döndürür."""
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

    thin  = Side(style="thin",   color="B0BEC5")
    med   = Side(style="medium", color="37474F")

    def _border(l=thin, r=thin, t=thin, b=thin):
        return Border(left=l, right=r, top=t, bottom=b)

    s = {}

   
    s["cc_title_font"]  = Font(name="Arial", bold=True, color="FFFFFF", size=13)
    s["cc_title_fill"]  = PatternFill("solid", fgColor="0D2B5E")
    s["cc_title_align"] = Alignment(horizontal="left", vertical="center", indent=1)
    s["cc_title_border"]= _border(l=med, r=med, t=med, b=med)

    
    s["sec_font"]  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    s["sec_fill"]  = PatternFill("solid", fgColor="1565C0")
    s["sec_align"] = Alignment(horizontal="left", vertical="center", indent=1)
    s["sec_border"]= _border(l=med, r=med, t=med, b=med)

   
    s["hdr_font"]  = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    s["hdr_fill"]  = PatternFill("solid", fgColor="1976D2")
    s["hdr_align"] = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s["hdr_border"]= _border(l=thin, r=thin, t=med, b=med)

    
    s["date_hdr_font"]  = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    s["date_hdr_fill"]  = PatternFill("solid", fgColor="00695C")
    s["date_hdr_align"] = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s["date_hdr_border"]= _border(l=thin, r=thin, t=med, b=med)

   
    s["first_col_font"]  = Font(name="Arial", bold=True, color="0D2B5E", size=9)
    s["first_col_fill"]  = PatternFill("solid", fgColor="E3F2FD")
    s["first_col_align"] = Alignment(horizontal="left", vertical="center", indent=1)
    s["first_col_border"]= _border(l=thin, r=med, t=thin, b=thin)

    
    s["data_font"]  = Font(name="Arial", size=9, color="212121")
    s["data_align"] = Alignment(horizontal="right", vertical="center")
    s["data_border"]= _border()

    
    s["zebra_fill"] = PatternFill("solid", fgColor="F5F5F5")

    
    s["doluluk_font"] = Font(name="Arial", bold=True, color="1A237E", size=9)
    s["doluluk_fill"] = PatternFill("solid", fgColor="E8EAF6")

    
    s["red_fill"] = PatternFill("solid", fgColor="DC2626")
    s["red_font"] = Font(name="Arial", bold=True, color="FFFFFF", size=9)

    
    s["sorunlu_cc_fill"] = PatternFill("solid", fgColor="E8EAF6")
    s["sorunlu_cc_font"] = Font(name="Arial", size=10, color="1A237E")
    s["sorunlu_deger_fill"] = PatternFill("solid", fgColor="B71C1C")
    s["sorunlu_deger_font"] = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    s["sorunlu_kap_fark_fill"] = PatternFill("solid", fgColor="FFE0B2")
    s["sorunlu_kum_fill"] = PatternFill("solid", fgColor="FFF9C4")
    s["sorunlu_doluluk_fill"] = PatternFill("solid", fgColor="F8BBD9")
    s["sorunlu_section_fill"] = PatternFill("solid", fgColor="37474F")
    s["sorunlu_section_font"] = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    s["sorunlu_cc_big_font"] = Font(name="Arial", bold=True, color="FFFFFF", size=15)
    s["sorunlu_cc_big_fill"] = PatternFill("solid", fgColor="0D47A1")
    s["sorunlu_sheet_title_font"] = Font(name="Arial", bold=True, color="FFFFFF", size=16)
    s["sorunlu_sheet_title_fill"] = PatternFill("solid", fgColor="B71C1C")
    s["sorunlu_hdr_font"] = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    s["sorunlu_hdr_fill"] = PatternFill("solid", fgColor="455A64")
    s["sorunlu_empty_font"] = Font(name="Arial", italic=True, color="78909C", size=10)

    
    cc_header_colors = ["0D47A1", "1565C0", "00695C", "2E7D32", "4A148C", "BF360C", "E65100", "1B5E20"]
    s["cc_title_fills"] = [PatternFill("solid", fgColor=c) for c in cc_header_colors]
    section_colors = ["37474F", "455A64", "546E7A", "607D8B", "78909C", "5C6BC0", "00838F", "6A1B9A"]
    s["sorunlu_section_fills"] = [PatternFill("solid", fgColor=c) for c in section_colors]
    s["sec_fills"] = [PatternFill("solid", fgColor=c) for c in section_colors]

    return s


def _sanitize_sheet_name(name, max_len=31):
    """Excel sheet adı: \ / * ? : [ ] kullanılamaz; en fazla 31 karakter."""
    if not name:
        return "Sheet"
    s = str(name).strip()
    for c in r'\/:*?[]':
        s = s.replace(c, "_")
    return s[:max_len] if len(s) > max_len else s


def _unique_sheet_name(base_name, existing_names, max_len=31):
    """Mevcut adlarla çakışmayan sheet adı üretir (gerekirse sonuna _2, _3 ekler)."""
    base = _sanitize_sheet_name(base_name, max_len)
    if base not in existing_names:
        return base
    existing = set(existing_names)
    for i in range(2, 100):
        candidate = f"{base[:max_len - len(str(i)) - 1]}_{i}"
        if candidate not in existing:
            return candidate
    return base + "_1"


def _is_date_col(col_name):
    """Kolon adı tarih/hafta/ay biçimindeyse True."""
    c = str(col_name).strip()
    if re.match(r"^\d{4}[-/W]\d{2}", c):
        return True
    if re.match(r"^W\d{1,2}$", c, re.IGNORECASE):
        return True
    if c == "0":
        return True
    return False


def _auto_col_widths(ws, first_col_width=30, date_col_width=10, other_col_width=14, max_w=40):
    """İlk kolon sabit geniş, tarih kolonları dar, diğerleri içeriğe göre."""
    from openpyxl.utils import get_column_letter
    for i, col_cells in enumerate(ws.columns, start=1):
        col_letter = get_column_letter(i)
        if i == 1:
            ws.column_dimensions[col_letter].width = first_col_width
            continue
        header_val = str(col_cells[0].value or "")
        if _is_date_col(header_val):
            ws.column_dimensions[col_letter].width = date_col_width
        else:
            max_len = max((len(str(cell.value or "")) for cell in col_cells), default=0)
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, other_col_width), max_w)



def _reorder_malzeme_columns_for_report(df, week_cols=None):
    """Malzeme tablosunda MATERIAL kolonunu MACHINE kolonundan hemen önce (sağa) alır; MACHINE, BASEQUAN, MTUNIT sona kalır."""
    if df is None or df.empty:
        return df
    end_cols = [c for c in ["MACHINE", "BASEQUAN", "MTUNIT"] if c in df.columns]
    if not end_cols:
        return df
    # MATERIAL → MACHINE hemen önüne; diğer kolonlar (DRAWNUM, hafta vb.) solda kalsın
    other = [c for c in df.columns if c not in ("MACHINE", "BASEQUAN", "MTUNIT", "MATERIAL")]
    material_col = ["MATERIAL"] if "MATERIAL" in df.columns else []
    return df[other + material_col + end_cols]


def _add_cap_alignment_column(cap_df):
    """Eski hiza için boş kolon ekleme kaldırıldı: Malzeme tablosunda MATERIAL sağa (MACHINE önüne) alınınca hiza aynı kalıyor."""
    return cap_df


def _apply_export_columns(df, hidden_columns=None, table_columns=None, table_name=None):
    """Excel/raporlama: Kolonları Göster/Gizle butonundaki gizli kolonları çıkar; tablo başlıklarını (örn. 0) uygula.
    hidden_columns yoksa (Raporlama sayfasından indirme) varsayılan olarak 'A' ile biten kolonlar (2026-01A vb.) çıkarılır."""
    if df is None or df.empty:
        return df
    df = df.copy()
    hidden_set = set(hidden_columns) if hidden_columns is not None else set()
    # Sadece Raporlama sayfasından indirildiğinde (hidden_columns None) varsayılan: A sütunlarını gizle (2026-01A, 2026-02A ...)
    if hidden_columns is None and not hidden_set:
        hidden_set = {c for c in df.columns if isinstance(c, str) and len(c) >= 6 and c.endswith("A")}
    if hidden_set:
        df = df[[c for c in df.columns if c not in hidden_set]]
    if table_columns:
        rename_map = {c.get("id"): c.get("name") for c in table_columns if c.get("id") and c.get("name") and c["id"] in df.columns}
        if rename_map:
            df = df.rename(columns=rename_map)
    elif table_name in ("VLFCAPFINALPIVOT", "VLFCAPFINALSIPARIS"):
        for c in df.columns:
            if re.match(r"^\d{4}-\d{2}$", str(c)):
                df = df.rename(columns={c: "0"})
                break
    return df


def _format_df_turkish_thousands(df):
    """EK 3/4: Sayısal sütunları binlik nokta ile formatlar (13.600); virgül/ondalık yok.
    Verimlilik kolonu atlanır (yüzde değeri olduğu için 80 vb. aynen yazılır)."""
    if df is None or df.empty:
        return df
    df = df.copy()
    for col in df.columns[1:]:
        if col == "Verimlilik":
            continue
        def _fmt(x):
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return x
            try:
                return f"{float(x):,.0f}".replace(",", ".")
            except Exception:
                return x
        df[col] = df[col].apply(_fmt)
    return df


def _parse_cell_to_float(val):
    """Hücre değerini sayıya çevirir ('13.600', '-2.100' vb.)."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _write_sorunlu_sheet_title(ws, row, s, num_cols=9):
    """Sorunlu Kapasite sayfasının en üstündeki ana başlık satırı."""
    from openpyxl.styles import Alignment
    cell = ws.cell(row=row, column=1, value="Sorunlu Kapasite — Detay için başlığa tıklayın")
    cell.font = s["sorunlu_sheet_title_font"]
    cell.fill = s["sorunlu_sheet_title_fill"]
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = s["sec_border"]
    if num_cols > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    ws.row_dimensions[row].height = 32
    return row + 3


def _write_sorunlu_section_title(ws, row, title, s, section_index=0, num_cols=1, start_col=1):
    """Sorunlu Kapasite sheet'inde bölüm başlığı (Hepsi, WC adı); section_index ile farklı renk."""
    from openpyxl.styles import Alignment
    cell = ws.cell(row=row, column=start_col, value=title)
    cell.font = s["sorunlu_section_font"]
    fills = s.get("sorunlu_section_fills", [s["sorunlu_section_fill"]])
    cell.fill = fills[section_index % len(fills)]
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=2, wrap_text=True)
    cell.border = s["sec_border"]
    if num_cols > 1:
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=start_col + num_cols - 1)
    ws.row_dimensions[row].height = 26
    return row + 2


def _write_sorunlu_cc_big_title(ws, row, costcenter_name, s, cc_index=0, num_cols=1):
    """Sorunlu Kapasite sheet'inde cost center büyük başlık; cc_index ile her CC farklı renk."""
    from openpyxl.styles import Alignment
    cell = ws.cell(row=row, column=1, value=costcenter_name)
    cell.font = s["sorunlu_cc_big_font"]
    fills = s.get("cc_title_fills", [s["sorunlu_cc_big_fill"]])
    cell.fill = fills[cc_index % len(fills)]
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    cell.border = s["sec_border"]
    if num_cols > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
    ws.row_dimensions[row].height = 30
    return row + 2



SORUNLU_DISPLAY_COLS = ["Costcenter", "Satır Tipi", "Hafta/Ay", "Değer"]


def _write_sorunlu_table_to_sheet(ws, start_row, df, s, start_col=1):
    """Sorunlu tablosunu Costcenter pastel, Satır Tipi renkli, Değer kırmızı ile yazar. start_col'dan itibaren (varsayılan 1)."""
    from openpyxl.styles import Alignment
    if df is None or df.empty:
        return start_row + 1
    cols = SORUNLU_DISPLAY_COLS
    # Tablo başlık satırı (Sorunlu'ya özel koyu gri, beyaz yazı)
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=start_row, column=start_col + c_idx - 1, value=col_name)
        cell.font = s.get("sorunlu_hdr_font", s["hdr_font"])
        cell.fill = s.get("sorunlu_hdr_fill", s["hdr_fill"])
        cell.alignment = s["hdr_align"]
        cell.border = s["hdr_border"]
    ws.row_dimensions[start_row].height = 24
    row = start_row + 1
 
    for _, r in df.iterrows():
        satir_tipi = str(r.get("Satır Tipi", "") or "").strip()
        if satir_tipi == "Kapasite Farkı":
            tip_fill = s["sorunlu_kap_fark_fill"]
        elif satir_tipi == "Kümülatif Toplam":
            tip_fill = s["sorunlu_kum_fill"]
        elif satir_tipi == "Doluluk Oranı(%)":
            tip_fill = s["sorunlu_doluluk_fill"]
        else:
            tip_fill = s["zebra_fill"]
        for c_idx, col_name in enumerate(cols, start=1):
            val = r.get(col_name)
            cell = ws.cell(row=row, column=start_col + c_idx - 1, value=val)
            cell.alignment = Alignment(horizontal="right" if c_idx >= 3 else "left", vertical="center")
            cell.border = s["data_border"]
            if c_idx == 1:
                cell.font = s["sorunlu_cc_font"]
                cell.fill = tip_fill
            elif c_idx == 2:
                cell.font = s["data_font"]
                cell.fill = tip_fill
            elif c_idx == 4:
                cell.font = s["sorunlu_deger_font"]
                cell.fill = s["sorunlu_deger_fill"]
            else:
                cell.font = s["data_font"]
                cell.fill = s["zebra_fill"]
        ws.row_dimensions[row].height = 20
        row += 1
    return row + 2


def _write_sorunlu_empty_cell(ws, row, col, s):
    """Sorunlu Kapasite'de 'Sorunlu nokta yok.' mesajı — italic gri."""
    cell = ws.cell(row=row, column=col, value="Sorunlu nokta yok.")
    cell.font = s.get("sorunlu_empty_font", s["data_font"])
    cell.border = s["data_border"]


def _extract_sorunlu_from_cap_df(cap_df, costcenter, workcenter="Tümü"):
    """Kapasite tablosundan sadece Kümülatif Toplam satırındaki kırmızı (negatif) hücreleri çıkarır.
    Başlığı '0' olan ilk kolon dahil edilmez. Dönen liste: [{"Costcenter", "Workcenter", "Satır Tipi", "Hafta/Ay", "Değer"}, ...]"""
    if cap_df is None or cap_df.empty or "STAT" not in cap_df.columns:
        return []
    
    week_cols = [c for c in cap_df.columns if c != "STAT" and str(c).strip() != "0" and _is_date_col(c)]
    out = []
    for _, row in cap_df.iterrows():
        stat_str = str(row.get("STAT") or row.iloc[0]).strip()
        if stat_str != "Kümülatif Toplam":
            continue
        for col in week_cols:
            val = row.get(col)
            num = _parse_cell_to_float(val)
            if num is None:
                continue
            if num < 0:
                out.append({
                    "Costcenter": costcenter,
                    "Workcenter": workcenter,
                    "Satır Tipi": stat_str,
                    "Hafta/Ay": col,
                    "Değer": round(num, 1) if isinstance(num, float) else num,
                })
    return out


def _extract_kumulatif_row_from_cap_df(cap_df, costcenter, workcenter="Tümü"):
    """Kapasite tablosundan Kümülatif Toplam satırının tamamını tek kayıt olarak döndürür.
    Dönen dict: Costcenter, Workcenter, STAT, ve tüm hafta/ay kolonları (0 dahil)."""
    if cap_df is None or cap_df.empty or "STAT" not in cap_df.columns:
        return None
    # 0 kolonu dahil tüm sayısal kolonlar (2. tabloda 0 da gösterilecek)
    week_cols = [c for c in cap_df.columns if c != "STAT"]
    for _, row in cap_df.iterrows():
        stat_str = str(row.get("STAT") or "").strip()
        if stat_str != "Kümülatif Toplam":
            continue
        rec = {"Costcenter": costcenter, "Workcenter": workcenter, "STAT": stat_str}
        for col in week_cols:
            val = row.get(col)
            num = _parse_cell_to_float(val)
            if num is not None:
                rec[col] = int(round(num, 0))
            else:
                rec[col] = val
        return rec
    return None


def _write_kumulatif_table_to_sheet(ws, start_row, df, s):
    """Kümülatif Toplam satırlarını yazar (sadece Workcenter + hafta kolonları). Negatif hücreler kırmızı. Costcenter/STAT yok."""
    from openpyxl.styles import Alignment, PatternFill
    _white_fill = PatternFill("solid", fgColor="FFFFFF")
    if df is None or df.empty:
        return start_row + 1
    cols = list(df.columns)
    # Başlık satırı
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=col_name)
        cell.font = s.get("sorunlu_hdr_font", s["hdr_font"])
        cell.fill = s.get("sorunlu_hdr_fill", s["hdr_fill"])
        cell.alignment = s["hdr_align"]
        cell.border = s["hdr_border"]
    ws.row_dimensions[start_row].height = 24
    row = start_row + 1
    for row_idx, r in enumerate(df.to_dict("records")):
        for c_idx, col_name in enumerate(cols, start=1):
            val = r.get(col_name)
            cell = ws.cell(row=row, column=c_idx, value=val)
            cell.border = s["data_border"]
            if col_name in ("Workcenter", "Costcenter"):
                cell.font = s["data_font"]
                cell.fill = s["sorunlu_kum_fill"]
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                num_val = _parse_cell_to_float(val)
                if num_val is not None:
                    cell.value = int(round(num_val, 0))
                if num_val is not None and num_val < 0:
                    cell.fill = s["sorunlu_deger_fill"]
                    cell.font = s["sorunlu_deger_font"]
                else:
                    cell.font = s["data_font"]
                    cell.fill = s["zebra_fill"] if (row_idx % 2 == 1) else _white_fill
        ws.row_dimensions[row].height = 20
        row += 1
    return row + 2



def _write_table_to_sheet(ws, start_row, df, s, title=None, is_section_title=False, cc_index=None, section_index=None):
    """
    DataFrame'i Excel sayfasına profesyonel stille yazar.
    cc_index: cost center başlığı rengi (her CC farklı renk).
    section_index: Hepsi / WC gibi alt başlık rengi.
    """
    def _title_fill():
        if is_section_title and cc_index is not None and "cc_title_fills" in s:
            return s["cc_title_fills"][cc_index % len(s["cc_title_fills"])]
        if not is_section_title and section_index is not None and "sec_fills" in s:
            return s["sec_fills"][section_index % len(s["sec_fills"])]
        return s["cc_title_fill"] if is_section_title else s["sec_fill"]

    if df is None or df.empty:
        if title:
            row = start_row
            cell = ws.cell(row=row, column=1, value=title)
            cell.font   = s["cc_title_font"]  if is_section_title else s["sec_font"]
            cell.fill   = _title_fill()
            cell.alignment = s["cc_title_align"] if is_section_title else s["sec_align"]
            cell.border = s["cc_title_border"] if is_section_title else s["sec_border"]
            ws.row_dimensions[row].height = 22
            return row + 2
        return start_row + 1

    df = _format_df_turkish_thousands(df)
    is_capacity_table = len(df.columns) > 0 and df.columns[0] == "STAT"
    num_cols = len(df.columns)
    row = start_row

    if title:
        cell = ws.cell(row=row, column=1, value=title)
        cell.font      = s["cc_title_font"]  if is_section_title else s["sec_font"]
        cell.fill      = _title_fill()
        cell.alignment = s["cc_title_align"] if is_section_title else s["sec_align"]
        cell.border    = s["cc_title_border"] if is_section_title else s["sec_border"]
        if num_cols > 1:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=num_cols)
        ws.row_dimensions[row].height = 22
        row += 1

   
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=row, column=col_idx, value=col_name)
        if _is_date_col(col_name) and col_idx > 1:
            cell.font      = s["date_hdr_font"]
            cell.fill      = s["date_hdr_fill"]
            cell.alignment = s["date_hdr_align"]
            cell.border    = s["date_hdr_border"]
        else:
            cell.font      = s["hdr_font"]
            cell.fill      = s["hdr_fill"]
            cell.alignment = s["hdr_align"]
            cell.border    = s["hdr_border"]
    ws.row_dimensions[row].height = 32  # wrap_text için yüksek tut
    row += 1

   
    for row_idx, r in enumerate(df.itertuples(index=False)):
        first_col_val  = r[0] if len(r) else None
        is_doluluk_row = str(first_col_val).strip() == "Doluluk Oranı(%)"
        is_even_row    = (row_idx % 2 == 1)

        for c_idx, val in enumerate(r, start=1):
            num_val = None
            # Kapasite tablosunda sayıları dashboard ile uyumlu tam sayı (veya Doluluk için 1 ondalık) yaz
            write_val = val
            if is_capacity_table and c_idx > 1:
                num_val = _parse_cell_to_float(val)
                if num_val is not None:
                    write_val = int(round(num_val, 0)) if not is_doluluk_row else round(num_val, 1)
            cell = ws.cell(row=row, column=c_idx, value=write_val)

            # Temel stil
            if c_idx == 1:
                cell.font      = s["first_col_font"]
                cell.fill      = s["first_col_fill"]
                cell.alignment = s["first_col_align"]
                cell.border    = s["first_col_border"]
            elif is_doluluk_row:
                cell.font      = s["doluluk_font"]
                cell.fill      = s["doluluk_fill"]
                cell.alignment = s["data_align"]
                cell.border    = s["data_border"]
            elif is_even_row:
                cell.font      = s["data_font"]
                cell.fill      = s["zebra_fill"]
                cell.alignment = s["data_align"]
                cell.border    = s["data_border"]
            else:
                cell.font      = s["data_font"]
                cell.alignment = s["data_align"]
                cell.border    = s["data_border"]

            # ── Kapasite tablosunda negatif veya Doluluk>100 → kırmızı (num_val yukarıda parse edildi)
            if is_capacity_table and c_idx > 1 and num_val is not None and (
                    num_val < 0 or (is_doluluk_row and num_val > 100)):
                cell.fill = s["red_fill"]
                cell.font = s["red_font"]

        ws.row_dimensions[row].height = 17
        row += 1

    return row + 2  



def _build_rapor_excel(costcenters, table_name, capacity_table_name, columns_dict, selected_units,
                       hidden_columns=None, table_columns=None):
    """columns_dict: format1, format2, format3. hidden_columns/table_columns: Kolonları Göster/Gizle ile uyumlu."""
    if not costcenters or ag is None:
        return None
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
    except ImportError:
        return None

    s  = _build_styles()
    wb = Workbook()

    kolon_sum  = columns_dict["format2"]
    kolon_sumb = columns_dict["format1"]
    week_cols  = kapasite_data.get_kolon_list_from_format3(columns_dict.get("format3", ""))

    # ─────────────────────────────────────────────────────────────
    # Sheet 1: 1. Accordion — Cost Center Kapasite Süresi
    # ─────────────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Costcenter Kapasite"
    ws1.sheet_properties.tabColor = "1565C0"
    row1 = 1
    # Kümülatif Toplam satırlarını topla (Sorunlu sayfasında tek tablo için)
    kumulatif_rows = []
    for idx, cc in enumerate(costcenters):
        cap_df = kapasite_data.build_capacity_table_for_cc(
            ag, table_name, capacity_table_name, cc,
            kolon_sum, kolon_sumb, week_cols, selected_units
        )
        cap_df = _apply_export_columns(cap_df, hidden_columns, table_columns, table_name)
        rec = _extract_kumulatif_row_from_cap_df(cap_df, cc, workcenter="Tümü")
        if rec:
            kumulatif_rows.append(rec)
        row1 = _write_table_to_sheet(
            ws1, row1, cap_df, s,
            title=f"◈  {cc}  —  1. Accordion Kapasite Süresi",
            is_section_title=True,
            cc_index=idx,
        )

    if row1 == 1:
        ws1.cell(row=1, column=1, value="Cost center bulunamadı veya veri yok.")

    _auto_col_widths(ws1)

    # Her CC için her Workcenter'ın Kümülatif Toplam satırını ekle
    for cc in costcenters:
        workcenters = kapasite_data.get_workcenters_for_cc(ag, table_name, cc)
        for wc in workcenters:
            cap_wc_df = kapasite_data.build_capacity_table_for_cc_workcenter(
                ag, table_name, capacity_table_name, cc, wc,
                kolon_sum, kolon_sumb, week_cols, selected_units
            )
            cap_wc_df = _apply_export_columns(cap_wc_df, hidden_columns, table_columns, table_name)
            rec = _extract_kumulatif_row_from_cap_df(cap_wc_df, cc, workcenter=wc)
            if rec:
                kumulatif_rows.append(rec)

    # CC → Excel sheet adı (detay sayfalarına link için)
    used_for_sheet = [ws1.title]
    cc_to_sheet = {}
    for cc in costcenters:
        sn = _unique_sheet_name(cc, used_for_sheet)
        cc_to_sheet[cc] = sn
        used_for_sheet.append(sn)

    ws_sorunlu = wb.create_sheet("Sorunlu Kapasite", 1)
    ws_sorunlu.sheet_properties.tabColor = "B71C1C"
    row_sorunlu = 1
    num_sorunlu_cols = 9  # başlık merge genişliği

    row_sorunlu = _write_sorunlu_sheet_title(ws_sorunlu, 1, s, num_cols=num_sorunlu_cols)

    # İki tablo: 1) Tümü olanlar — ilk kolon Costcenter adı (Azot vb.); 2) Makine bazlı — 0 kolonu dahil, sadece kırmızı (negatif) olan satırlar.
    sorunlu_hyperlink_cells = []
    if kumulatif_rows:
        try:
            week_cols_order = [c for c in kumulatif_rows[0].keys() if c not in ("Costcenter", "Workcenter", "STAT")]

            tumu_rows = [r for r in kumulatif_rows if r.get("Workcenter") == "Tümü"]
            diger_rows = [r for r in kumulatif_rows if r.get("Workcenter") != "Tümü"]

            # 2. tablo için: 0 kolonu hariç satırda en az bir negatif (kırmızı) varsa satır gelsin; yoksa gelmesin. Null-only satırlar da gelmesin.
            def _row_has_negative_except_zero(rec):
                for col in week_cols_order:
                    if col == "0":
                        continue
                    v = _parse_cell_to_float(rec.get(col))
                    if v is not None and v < 0:
                        return True
                return False
            diger_rows = [r for r in diger_rows if _row_has_negative_except_zero(r)]

            display_cols_tumu = ["Costcenter"] + week_cols_order   # 1. tablo: Azot, vb. (Costcenter adı)
            display_cols_diger = ["Workcenter"] + week_cols_order   # 2. tablo: 0 dahil, makine adı

            # 1. Tablo: Tümü — ilk kolonda Costcenter adı (Azot, vb.)
            row_sorunlu = _write_sorunlu_section_title(
                ws_sorunlu, row_sorunlu, "1. Kümülatif Toplam — Tümü (Costcenter toplamları)", s,
                section_index=0, num_cols=len(display_cols_tumu), start_col=1
            )
            if tumu_rows:
                df_tumu = pd.DataFrame(tumu_rows)[display_cols_tumu]
                data_start_1 = row_sorunlu + 1
                row_sorunlu = _write_kumulatif_table_to_sheet(ws_sorunlu, row_sorunlu, df_tumu, s)
                for i, r in enumerate(tumu_rows):
                    sorunlu_hyperlink_cells.append((data_start_1 + i, 1, r.get("Costcenter"), r.get("Workcenter")))
            else:
                _write_sorunlu_empty_cell(ws_sorunlu, row_sorunlu, 1, s)
                row_sorunlu += 4

            # 2. Tablo: Makine bazlı — 0 kolonu dahil; sadece 0 dışında en az bir negatif olan satırlar
            row_sorunlu = _write_sorunlu_section_title(
                ws_sorunlu, row_sorunlu, "2. Kümülatif Toplam — Makine bazlı (sadece negatif içeren satırlar)", s,
                section_index=1, num_cols=len(display_cols_diger), start_col=1
            )
            if diger_rows:
                df_diger = pd.DataFrame(diger_rows)[display_cols_diger]
                data_start_2 = row_sorunlu + 1
                row_sorunlu = _write_kumulatif_table_to_sheet(ws_sorunlu, row_sorunlu, df_diger, s)
                for i, r in enumerate(diger_rows):
                    sorunlu_hyperlink_cells.append((data_start_2 + i, 1, r.get("Costcenter"), r.get("Workcenter")))
            else:
                _write_sorunlu_empty_cell(ws_sorunlu, row_sorunlu, 1, s)
        except Exception as ex:
            _write_sorunlu_empty_cell(ws_sorunlu, row_sorunlu, 1, s)
            print(f"[Raporlama] Sorunlu Kapasite tablosu yazılırken hata: {ex}")
    else:
        _write_sorunlu_empty_cell(ws_sorunlu, row_sorunlu, 1, s)

    _auto_col_widths(ws_sorunlu, first_col_width=22, other_col_width=14, date_col_width=12)

    cc_section_rows = {}
    used_sheet_names = [ws1.title, ws_sorunlu.title]
    for idx, cc in enumerate(costcenters):
        sheet_name = _unique_sheet_name(cc, used_sheet_names)
        used_sheet_names.append(sheet_name)
        ws_cc = wb.create_sheet(sheet_name, 2 + idx)
        ws_cc.sheet_properties.tabColor = "00695C"
        row_cc = 1
        cc_section_rows[sheet_name] = {}

        # CC başlığı (her costcenter farklı renk)
        row_cc = _write_table_to_sheet(
            ws_cc, row_cc, pd.DataFrame(), s,
            title=f"◈  {cc}",
            is_section_title=True,
            cc_index=idx,
        )
        # 1) Blok "Hepsi" — tek grup: başlık (Hepsi) + altında Kapasite + Malzeme. Açılışta kapalı, + ile açılır. Outline sembolleri (1,2,3,4) kapalı.
        row_hepsi = row_cc
        row_cc = _write_table_to_sheet(
            ws_cc, row_cc, pd.DataFrame(), s,
            title="  Hepsi",
            is_section_title=False,
            section_index=0,
        )
        cc_section_rows[sheet_name]["Hepsi"] = row_hepsi
        cap_hepsi = kapasite_data.build_capacity_table_for_cc(
            ag, table_name, capacity_table_name, cc,
            kolon_sum, kolon_sumb, week_cols, selected_units
        )
        cap_hepsi = _add_cap_alignment_column(cap_hepsi)
        cap_hepsi = _apply_export_columns(cap_hepsi, hidden_columns, table_columns, table_name)
        row_cc = _write_table_to_sheet(
            ws_cc, row_cc, cap_hepsi, s,
            title="    Kapasite Süresi",
            is_section_title=False,
        )
        malz_hepsi = kapasite_data.build_malzeme_table_for_cc(
            ag, table_name, cc, columns_dict["format2"], selected_units, for_report=True
        )
        malz_hepsi = _reorder_malzeme_columns_for_report(malz_hepsi, week_cols)
        malz_hepsi = _apply_export_columns(malz_hepsi, hidden_columns, table_columns, table_name)
        row_cc = _write_table_to_sheet(
            ws_cc, row_cc, malz_hepsi, s,
            title="    Malzeme Tablosu",
            is_section_title=False,
        )
        ws_cc.row_dimensions[row_hepsi].outlineLevel = 1
        for r in range(row_hepsi + 1, row_cc):
            ws_cc.row_dimensions[r].outlineLevel = 2
            ws_cc.row_dimensions[r].hidden = True

        # 2) Her workcenter için ayrı blok: tek grup (WC başlığı + Kapasite + Malzeme). Açılışta kapalı, + ile açılır.
        workcenters = kapasite_data.get_workcenters_for_cc(ag, table_name, cc)
        for wc_idx, wc in enumerate(workcenters):
            row_wc = row_cc
            row_cc = _write_table_to_sheet(
                ws_cc, row_cc, pd.DataFrame(), s,
                title=f"  {wc}",
                is_section_title=False,
                section_index=wc_idx + 1,
            )
            cc_section_rows[sheet_name][wc] = row_wc
            cap_wc = kapasite_data.build_capacity_table_for_cc_workcenter(
                ag, table_name, capacity_table_name, cc, wc,
                kolon_sum, kolon_sumb, week_cols, selected_units
            )
            cap_wc = _add_cap_alignment_column(cap_wc)
            cap_wc = _apply_export_columns(cap_wc, hidden_columns, table_columns, table_name)
            row_cc = _write_table_to_sheet(
                ws_cc, row_cc, cap_wc, s,
                title="    Kapasite Süresi",
                is_section_title=False,
            )
            malz_wc = kapasite_data.build_malzeme_table_for_cc_workcenter(
                ag, table_name, cc, wc, columns_dict["format2"], selected_units, for_report=True
            )
            malz_wc = _reorder_malzeme_columns_for_report(malz_wc, week_cols)
            malz_wc = _apply_export_columns(malz_wc, hidden_columns, table_columns, table_name)
            row_cc = _write_table_to_sheet(
                ws_cc, row_cc, malz_wc, s,
                title="    Malzeme Tablosu",
                is_section_title=False,
            )
            ws_cc.row_dimensions[row_wc].outlineLevel = 1
            for r in range(row_wc + 1, row_cc):
                ws_cc.row_dimensions[r].outlineLevel = 2
                ws_cc.row_dimensions[r].hidden = True

        # 1, 2, 3, 4 butonları kapalı; sadece her başlığın yanındaki + / - görünür. +/- Excel’de gruplarla birlikte gelir, ayrı kapatılamaz.
        ws_cc.sheet_view.showOutlineSymbols = False
        _auto_col_widths(ws_cc)

    # Sorunlu Kapasite sayfasındaki Costcenter hücresine tıklanınca ilgili detay sayfasına git
    from openpyxl.styles import Font
    for (row, col, cc, workcenter) in sorunlu_hyperlink_cells:
        sheet_name = cc_to_sheet.get(cc)
        if not sheet_name:
            continue
        target_row = 1
        section_key = "Hepsi" if (workcenter == "Tümü" or not workcenter) else workcenter
        if sheet_name in cc_section_rows:
            target_row = cc_section_rows[sheet_name].get(section_key, 1)
        # Excel dahili link: #'Sayfa Adı'!A1
        safe_name = sheet_name.replace("'", "''")
        cell = ws_sorunlu.cell(row=row, column=col)
        cell.hyperlink = f"#'{safe_name}'!A{target_row}"
        old_font = cell.font
        cell.font = Font(
            name=old_font.name if old_font else "Arial",
            bold=old_font.bold if old_font else False,
            size=old_font.size if old_font else 11,
            color="0563C1",
            underline="single",
        )

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
layout = dbc.Container([
    html.Div([
        html.Div([
            html.Span("◈", className="kap-header-icon"),
            html.Div([
                html.H1("KAPASİTE RAPORLAMA", className="kap-header-title"),
                html.P(
                    "İhtiyaç, Sipariş ve Öngörü veri tipleri için ayrı Excel dosyaları — tek ZIP olarak iner. Her tablo bloğunun başında cost center adı yazılır.",
                    className="kap-header-sub",
                ),
                html.P(
                    "Excel'de rapor açıldığında sadece bölüm başlıkları (Hepsi, iş merkezi adları) görünür. Sol taraftaki + işaretine tıklayarak ilgili Malzeme veya Kapasite Süresi tablosunu açar, − ile kapatırsınız.",
                    className="kap-header-sub",
                    style={"fontSize": "0.9em", "opacity": 0.85, "marginTop": "4px"},
                ),
            ]),
        ], className="kap-header-inner"),
    ], className="kap-page-header"),

    dbc.Row([
        dbc.Col([
            html.Label(["Veri Tipi"], className="kap-label"),
            html.Div(
                "İhtiyaç · Sipariş · Öngörü (hepsi ZIP içinde)",
                style={
                    "padding": "8px 12px",
                    "backgroundColor": "#e3f2fd",
                    "borderRadius": "8px",
                    "fontWeight": "600",
                },
            ),
        ], md=4),
        dbc.Col([
            html.Label(["Zaman Birimi"], className="kap-label"),
            html.Div(
                "Saat (sat)",
                style={
                    "padding": "8px 12px",
                    "backgroundColor": "#e8f5e9",
                    "borderRadius": "8px",
                    "fontWeight": "600",
                },
            ),
        ], md=4),
    ], className="mb-4"),

    dbc.Row([
        dbc.Col([
            html.Button(
                [html.Span("⬇ ", style={"marginRight": "6px"}), "Raporu Oluştur ve İndir"],
                id="raporlama-btn-generate",
                n_clicks=0,
                className="kap-btn kap-btn-primary",
            ),
            dcc.Download(id="raporlama-download"),
        ], md=12),
    ]),

    html.Div(id="raporlama-status", className="mt-3"),
], fluid=True, className="kap-control-panel")


# ─────────────────────────────────────────────────────────────────────────────
# Ortak ZIP oluşturma (Raporlama sayfası + Kapasite sayfası butonu birebir aynı)
# ─────────────────────────────────────────────────────────────────────────────
def build_raporlama_zip(ag_instance, hidden_columns=None, table_columns=None):
    """İhtiyaç, Sipariş, Öngörü için ayrı Excel'leri oluşturur, tek ZIP döndürür.
    Kapasite ekranından çağrıldığında sadece tablo kolon filtreleri (Kolonları Göster/Gizle) yansır.
    Birim, cost center vb. şartlar raporlama sayfasının kendi ayarına göre (örn. saat) kalır."""
    if ag_instance is None:
        return None, []
    selected_units = ["hours"]
    zip_buffer = io.BytesIO()
    created = []
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for label, table_name, capacity_table_name, period in DATA_TYPES:
            try:
                df_cc = ag_instance.run_query(f"SELECT DISTINCT STAND FROM [{table_name}] ORDER BY STAND")
                if df_cc is None or df_cc.empty:
                    continue
                all_cc_list = df_cc["STAND"].tolist()
                # RAPORLAMA_COSTCENTERS doluysa sadece onları al; boşsa tablodaki tüm cost center'lar
                if RAPORLAMA_COSTCENTERS:
                    all_cc = set(all_cc_list)
                    costcenters = [cc for cc in RAPORLAMA_COSTCENTERS if cc in all_cc]
                else:
                    costcenters = all_cc_list
                if not costcenters:
                    continue
            except Exception as e:
                import traceback
                print(f"[Raporlama] Sorgu hatası ({table_name}): {e}")
                traceback.print_exc()
                continue
            if period == "weekly":
                columns_dict = kapasite_data.generate_weekly_columns()
            else:
                columns_dict = kapasite_data.generate_monthly_columns_filtered(ag_instance, table_name, capacity_table_name)
            excel_bytes = _build_rapor_excel(
                costcenters, table_name, capacity_table_name, columns_dict, selected_units,
                hidden_columns=hidden_columns, table_columns=table_columns,
            )
            if excel_bytes:
                safe_name = (label
                             .replace(" ", "_")
                             .replace("ı", "i").replace("İ", "I")
                             .replace("ö", "o").replace("Ö", "O")
                             .replace("ü", "u").replace("Ü", "U")
                             .replace("ş", "s").replace("Ş", "S")
                             .replace("ç", "c").replace("Ç", "C")
                             .replace("ğ", "g").replace("Ğ", "G"))
                zf.writestr(f"kapasite_rapor_{safe_name}.xlsx", excel_bytes)
                created.append(f"{label} ({len(costcenters)} CC)")
    zip_buffer.seek(0)
    return zip_buffer.getvalue() if created else None, created


# ─────────────────────────────────────────────────────────────────────────────
# CALLBACK: Rapor oluştur ve indir
# ─────────────────────────────────────────────────────────────────────────────
@app.callback(
    Output("raporlama-download", "data"),
    Output("raporlama-status", "children"),
    Input("raporlama-btn-generate", "n_clicks"),
    prevent_initial_call=True,
)
def raporlama_generate(n_clicks):
    if not n_clicks:
        raise PreventUpdate
    if ag is None:
        return no_update, dbc.Alert("Veritabanı bağlantısı yok. Lütfen bağlantıyı kontrol edin.", color="danger")
    try:
        zip_bytes, created = build_raporlama_zip(ag)
    except Exception as e:
        return no_update, dbc.Alert(
            f"Rapor oluşturulurken hata: Veritabanı bağlantısı kurulamıyor veya sorgu hatası. ({e!r})",
            color="danger",
        )
    if not zip_bytes:
        return no_update, dbc.Alert("Hiçbir veri tipi için rapor oluşturulamadı. Veritabanında veri olmayabilir.", color="warning")
    return (
        dict(
            content=base64.b64encode(zip_bytes).decode("ascii"),
            filename="kapasite_raporlar.zip",
            base64=True,
            type="application/zip",
        ),
        dbc.Alert(f"Raporlar oluşturuldu: {', '.join(created)}. ZIP indiriliyor.", color="success"),
    )