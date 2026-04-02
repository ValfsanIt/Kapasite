# -*- coding: utf-8 -*-

from datetime import datetime, timedelta
import pandas as pd


DISPLAY_MAX_DECIMALS_NON_A = 4


def get_kolon_list_from_format3(format3_str):
    
    if not format3_str:
        return []
    return [x.strip() for x in format3_str.split(",") if x.strip()]


def _sql_bracket_table(name):
    """SQL Server için [tablo] — Dash ve rapor sorgularını birebir hizalar."""
    if name is None:
        return ""
    s = str(name).strip()
    if not s:
        return ""
    if s.startswith("[") and s.endswith("]"):
        return s
    return f"[{s}]"


def _to_numeric_slice(df, start_col_index=1):
    if df is None or df.empty:
        return
    for i in range(start_col_index, len(df.columns)):
        col = df.columns[i]
        df[col] = pd.to_numeric(df[col], errors="coerce")


def _div_safe(df, start_col_index, divisor):
    if df is None or df.empty or start_col_index >= df.shape[1]:
        return
    target_cols = list(df.columns[start_col_index:])
    
    for col in target_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype(float)
    df.loc[:, target_cols] = df.loc[:, target_cols].div(divisor)


def _apply_cap_work_unit_like_dash(sum_df_cap_work, selected_units):
    
    if sum_df_cap_work is None or sum_df_cap_work.empty or sum_df_cap_work.shape[1] < 2:
        return
    su = selected_units or []
    if "hours" in su or "shifts" in su:
        sub = sum_df_cap_work.iloc[:, 1:].apply(pd.to_numeric, errors="coerce").astype("float64")
        if "hours" in su:
            sum_df_cap_work.iloc[:, 1:] = sub / 60.0
        else:
            sum_df_cap_work.iloc[:, 1:] = sub / 510.0
    else:
        _to_numeric_slice(sum_df_cap_work, 0)


def _select_sum_with_unit(kolon_list, selected_units):
    
    if not kolon_list:
        return ""
    if isinstance(kolon_list, str):
        kolon_list = [c.strip() for c in kolon_list.split(",") if c.strip()]
    divisor = None
    if selected_units and "hours" in selected_units:
        divisor = 60
    elif selected_units and "shifts" in selected_units:
        divisor = 510
    parts = []
    for col in kolon_list:
        if not col or not str(col).strip():
            continue
        col = str(col).strip()
        if divisor:
            parts.append(f"SUM([{col}])/{divisor} AS [{col}]")
        else:
            parts.append(f"SUM([{col}]) AS [{col}]")
    return ", ".join(parts)


def _unit_decimals(selected_units):
    if selected_units and "hours" in selected_units:
        return 1
    if selected_units and "shifts" in selected_units:
        return 2
    return 0


def _is_a_col(col_name):
    return str(col_name).strip().endswith("A")


def _format_numeric_cols_by_unit(df, numeric_cols, selected_units):
    
    if df is None or df.empty or not numeric_cols:
        return df
    base_dec = _unit_decimals(selected_units)
    su = selected_units or []
    float_non_a = base_dec > 0 and ("hours" in su or "shifts" in su)
    for col in numeric_cols:
        if col not in df.columns:
            continue
        ser = pd.to_numeric(df[col], errors="coerce")
        if _is_a_col(col):
            df[col] = ser.round(0).astype("Int64")
        elif float_non_a:
            df[col] = ser.round(DISPLAY_MAX_DECIMALS_NON_A)
        else:
            df[col] = ser.round(0).astype("Int64")
    return df


def _finalize_capacity_stat_like_dash(cap_df, doluluk_df, weeks, selected_units):
    
    if cap_df is None or cap_df.empty or doluluk_df is None or doluluk_df.empty or not weeks:
        return cap_df
    _format_numeric_cols_by_unit(cap_df, [c for c in cap_df.columns if c != "STAT"], selected_units)
    out = pd.concat([cap_df, doluluk_df], ignore_index=True)
    _format_numeric_cols_by_unit(out, weeks, selected_units)
    return out


def _id_columns_malzeme(df):
    return {c for c in ("MATERIAL", "DRAWNUM", "MACHINE", "BASEQUAN", "MTUNIT", "CAPWORK", "STAND") if c in df.columns}


def _sql_escape_literal(val):
    if val is None:
        return ""
    return str(val).replace("'", "''")


def _malzeme_where_clause(costcenter, cap_grp=None, workcenter=None):
    parts = [f"STAND = '{_sql_escape_literal(costcenter)}'"]
    if cap_grp and str(cap_grp).strip() and str(cap_grp) != "Kapasite Grubu":
        parts.append(f"CAPGRUP = '{_sql_escape_literal(cap_grp)}'")
    if workcenter and str(workcenter).strip() and str(workcenter) != "Hepsi":
        parts.append(f"CAPWORK = '{_sql_escape_literal(workcenter)}'")
    return "WHERE " + " AND ".join(parts)


def _malzeme_div_start_col_index(df):
    idset = _id_columns_malzeme(df)
    for i, col in enumerate(df.columns):
        if col not in idset:
            return i
    return len(df.columns)


def generate_weekly_columns():
    
    start_date = datetime.now() - timedelta(weeks=1)
    format1, format2, format3, format4 = [], [], [], []
    for i in range(19):
        week_start = start_date + timedelta(weeks=i)
        year, week_num, _ = week_start.isocalendar()
        wk = f"{year}-{str(week_num).zfill(2)}"
        if i == 0:
            format1.append(f"0 AS [{wk}]")
        else:
            format1.append(f"SUM(B.[{wk}]) AS [{wk}]")
        format2.append(f"SUM([{wk}]) AS [{wk}], SUM([{wk}A]) AS [{wk}A]")
        format3.append(f"{wk}, {wk}A")
        if i == 0:
            format4.append(f"0 AS [{wk}]")
        else:
            format4.append(
                f"CAST(CEILING((CAST(SUM(A.[{wk}]) AS DECIMAL(18, 3))/CAST(SUM(B.[{wk}]) AS DECIMAL(18, 3)))*100) AS int) AS [{wk}]"
            )
    return {
        "format1": ", ".join(format1),
        "format2": ", ".join(format2),
        "format3": ", ".join(format3),
        "format4": ", ".join(format4),
    }


def generate_monthly_columns():
    
    today = datetime.today()
    start_year = today.year
    end_year = start_year + 1
    format1, format2, format3, format4 = [], [], [], []
    current_date = datetime(start_year, 1, 1)
    end_date = datetime(end_year, 12, 1)
    while current_date <= end_date:
        ym = f"{current_date.year}-{str(current_date.month).zfill(2)}"
        format1.append(f"SUM(B.[{ym}]) AS [{ym}]")
        format2.append(f"SUM([{ym}]) AS [{ym}], SUM([{ym}A]) AS [{ym}A]")
        format3.append(ym)
        format3.append(f"{ym}A")
        format4.append(
            f"CAST(CEILING((CAST(SUM(A.[{ym}]) AS DECIMAL(18, 3))/"
            f"CASE WHEN CAST(SUM(B.[{ym}]) AS DECIMAL(18, 3)) = 0 "
            f"THEN 1 ELSE CAST(SUM(B.[{ym}]) AS DECIMAL(18, 3)) END)*100) AS int) AS [{ym}]"
        )
        next_month = current_date.month + 1
        next_year = current_date.year + (1 if next_month > 12 else 0)
        next_month = 1 if next_month > 12 else next_month
        current_date = datetime(next_year, next_month, 1)
    return {
        "format1": ", ".join(format1),
        "format2": ", ".join(format2),
        "format3": ", ".join(format3),
        "format4": ", ".join(format4),
    }


def get_table_columns(ag_instance, table_name):
   
    if not ag_instance or not table_name:
        return set()
    try:
        df = ag_instance.run_query(f"SELECT TOP 1 * FROM [{table_name}]")
        if df is None or df.empty:
            return set()
        return set(df.columns.tolist())
    except Exception:
        return set()



VERIMLILIK_SQL = "SELECT WORKCENTER, [VERIMLILIK] AS Verimlilik FROM [VLFVARDIYASURE]"


def get_verimlilik_df(ag_instance):
   
    if not ag_instance:
        return None
    try:
        df = ag_instance.run_query(VERIMLILIK_SQL)
        if df is None or df.empty or "Verimlilik" not in df.columns:
            return None
        return df
    except Exception:
        return None


def generate_monthly_columns_filtered(ag_instance, table_name, capacity_table_name):
    
    raw = generate_monthly_columns()
    table_set = get_table_columns(ag_instance, table_name)
    cap_set = get_table_columns(ag_instance, capacity_table_name)
    if not table_set and not cap_set:
        return raw
    format1, format2, format3 = [], [], []
    seen = set()
    for part in raw["format3"].split(","):
        ym = part.strip()
        if not ym or ym.endswith("A"):
            continue
        if ym in seen:
            continue
        seen.add(ym)
        if ym in table_set and f"{ym}A" in table_set:
            format2.append(f"SUM([{ym}]) AS [{ym}], SUM([{ym}A]) AS [{ym}A]")
            format3.append(ym)
            format3.append(f"{ym}A")
        if ym in cap_set:
            format1.append(f"SUM(B.[{ym}]) AS [{ym}]")
    if not format2 and table_set:
        ay_cols = sorted(
            c for c in table_set
            if len(c) >= 7 and c[4] == "-"
            and (c.replace("-", "")[:6].isdigit() or (c.endswith("A") and c[:-1].replace("-", "")[:6].isdigit()))
        )
        for ym in ay_cols:
            if ym.endswith("A"):
                continue
            if f"{ym}A" in table_set:
                format2.append(f"SUM([{ym}]) AS [{ym}], SUM([{ym}A]) AS [{ym}A]")
                format3.append(ym)
                format3.append(f"{ym}A")
    if not format1 and cap_set:
        ay_cap = sorted(c for c in cap_set if len(c) >= 7 and c[4] == "-" and c.replace("-", "")[:6].isdigit())
        format1 = [f"SUM(B.[{c}]) AS [{c}]" for c in ay_cap]
    return {
        "format1": ", ".join(format1) if format1 else raw["format1"],
        "format2": ", ".join(format2) if format2 else raw["format2"],
        "format3": ", ".join(format3) if format3 else raw["format3"],
    }



def build_capacity_table_for_cc(ag_instance, table_name, capacity_table_name, costcenter,
                                kolon_sum, kolon_sumb, kolon_list, selected_units=("hours",), cap_grp=None):
    
    if not costcenter or not table_name or not kolon_sum:
        return None
    try:
        where_cc = f"STAND = '{costcenter}'"
        if cap_grp and cap_grp != "Kapasite Grubu":
            where_cc = f"CAPGRUP = '{cap_grp}' AND STAND = '{costcenter}'"
            sub_extra = ",CAPGRUP"
            where_extra = f" AND A.CAPGRUP = '{cap_grp}'"
        else:
            sub_extra = ""
            where_extra = ""
        selected_units = selected_units or ["hours"]
        prediv_ihtiyac = _select_sum_with_unit(kolon_list, selected_units)
        ihtiyac_select = prediv_ihtiyac if prediv_ihtiyac else kolon_sum
        tn = _sql_bracket_table(table_name)
        ctn = _sql_bracket_table(capacity_table_name)
        ihtiyac_sql = (
            f"SELECT STAND,{ihtiyac_select} FROM {tn} "
            f"WHERE {where_cc} GROUP BY STAND ORDER BY STAND"
        )
        cap_work_sql = (
            f"SELECT {kolon_sumb} FROM (SELECT STAND,CAPWORK{sub_extra} FROM {tn} "
            f"GROUP BY CAPWORK,STAND{sub_extra}) A LEFT JOIN {ctn} B ON A.CAPWORK = B.WORKCENTER "
            f"WHERE A.STAND = '{costcenter}'{where_extra} GROUP BY A.STAND"
        )
        sum_df = ag_instance.run_query(ihtiyac_sql)
        sum_df_cap_work = ag_instance.run_query(cap_work_sql)
        if sum_df is None or sum_df.empty or sum_df_cap_work is None or sum_df_cap_work.empty:
            return None

        if not prediv_ihtiyac:
            if "hours" in selected_units:
                _div_safe(sum_df, 1, 60)
            elif "shifts" in selected_units:
                _div_safe(sum_df, 1, 510)

        sum_numeric = [c for c in sum_df.columns if c not in ("STAND",)]
        _format_numeric_cols_by_unit(sum_df, sum_numeric, selected_units)

        sum_df["STAT"] = "Kapasite İhtiyacı"
        weeks = [c for c in kolon_list if c in sum_df.columns]
        if not weeks:
            weeks = [c for c in sum_df.columns if c not in ("STAND", "STAT")]
        filtered_sum_df = sum_df[["STAT"] + weeks].copy()

        if sum_df_cap_work.shape[1] > 0:
            _apply_cap_work_unit_like_dash(sum_df_cap_work, selected_units)
            cap_work_numeric = [c for c in sum_df_cap_work.columns if c not in ("STAND",)]
            _format_numeric_cols_by_unit(sum_df_cap_work, cap_work_numeric, selected_units)

        toplam_row = {"STAT": "Toplam Kapasite"}
        toplam_row.update(sum_df_cap_work.iloc[0].to_dict())
        cap_df = pd.concat([filtered_sum_df, pd.DataFrame([toplam_row])], ignore_index=True)

        fark_row = {"STAT": "Kapasite Farkı"}
        su = selected_units or []
        sql_float_fark = _unit_decimals(selected_units) > 0 and ("hours" in su or "shifts" in su)
        for col in weeks:
            try:
                d = float(cap_df.loc[cap_df["STAT"] == "Toplam Kapasite", col].iloc[0]) - float(
                    cap_df.loc[cap_df["STAT"] == "Kapasite İhtiyacı", col].iloc[0]
                )
                if sql_float_fark and not _is_a_col(col):
                    fark_row[col] = round(d, DISPLAY_MAX_DECIMALS_NON_A)
                else:
                    fark_row[col] = round(d, 0)
            except Exception:
                fark_row[col] = 0
        cap_df = pd.concat([cap_df, pd.DataFrame([fark_row])], ignore_index=True)

        numeric_cols = [c for c in cap_df.columns if c != "STAT"]
        cumsum = cap_df.loc[cap_df["STAT"] == "Kapasite Farkı", numeric_cols].cumsum(axis=1)
        cum_row = {"STAT": "Kümülatif Toplam"}
        cum_raw = cumsum.iloc[0].to_dict()
        cum_row.update(cum_raw)
        for k, v in list(cum_row.items()):
            if k == "STAT" or not isinstance(v, (int, float)):
                continue
            if sql_float_fark and not _is_a_col(k):
                cum_row[k] = round(v, DISPLAY_MAX_DECIMALS_NON_A)
            else:
                cum_row[k] = round(v, 0)
        cap_df = pd.concat([cap_df, pd.DataFrame([cum_row])], ignore_index=True)

        ui_row = cap_df[cap_df["STAT"] == "Kapasite İhtiyacı"].iloc[0]
        tk_row = cap_df[cap_df["STAT"] == "Toplam Kapasite"].iloc[0]
        doluluk_vals = []
        for col in weeks:
            try:
                tk = float(tk_row[col])
                ui = float(ui_row[col])
                doluluk_vals.append(round((ui / tk) * 100, 0) if tk != 0 else 0)
            except Exception:
                doluluk_vals.append(0)
        doluluk_df = pd.DataFrame([doluluk_vals], columns=weeks)
        doluluk_df["STAT"] = "Doluluk Oranı(%)"
        return _finalize_capacity_stat_like_dash(cap_df, doluluk_df, weeks, selected_units)
    except Exception as e:
        print(f"kapasite_data build_capacity_table_for_cc hatası: {e}")
        return None


def get_workcenters_for_cc(ag_instance, table_name, costcenter):
    """Verilen cost center için tablodaki tüm CAPWORK (workcenter) listesini döndürür."""
    if not ag_instance or not table_name or not costcenter:
        return []
    try:
        df = ag_instance.run_query(
            f"SELECT DISTINCT CAPWORK FROM [{table_name}] WHERE STAND = '{costcenter}' ORDER BY CAPWORK"
        )
        if df is None or df.empty or "CAPWORK" not in df.columns:
            return []
        return df["CAPWORK"].astype(str).str.strip().tolist()
    except Exception:
        return []


def build_capacity_table_for_cc_workcenter(ag_instance, table_name, capacity_table_name, costcenter, workcenter,
                                           kolon_sum, kolon_sumb, kolon_list, selected_units=("hours",), cap_grp=None):
    """2. accordion: Tek bir (CC, Workcenter) için kapasite süresi tablosu (STAT + hafta/ay)."""
    if not costcenter or not workcenter or not table_name or not kolon_sum:
        return None
    try:
        where_cc_wc = f"STAND = '{costcenter}' AND CAPWORK = '{workcenter}'"
        if cap_grp and cap_grp != "Kapasite Grubu":
            where_cc_wc = f"CAPGRUP = '{cap_grp}' AND STAND = '{costcenter}' AND CAPWORK = '{workcenter}'"
            sub_extra = ",CAPGRUP"
            where_extra = f" AND A.CAPGRUP = '{cap_grp}'"
        else:
            sub_extra = ""
            where_extra = ""
        selected_units = selected_units or ["hours"]
        prediv_ihtiyac = _select_sum_with_unit(kolon_list, selected_units)
        ihtiyac_select = prediv_ihtiyac if prediv_ihtiyac else kolon_sum
        tn = _sql_bracket_table(table_name)
        ctn = _sql_bracket_table(capacity_table_name)
        ihtiyac_sql = (
            f"SELECT CAPWORK,{ihtiyac_select} FROM {tn} "
            f"WHERE {where_cc_wc} GROUP BY CAPWORK ORDER BY CAPWORK"
        )
        cap_work_sql = (
            f"SELECT {kolon_sumb} FROM (SELECT STAND,CAPWORK{sub_extra} FROM {tn} "
            f"WHERE {where_cc_wc} GROUP BY CAPWORK,STAND{sub_extra}) A "
            f"LEFT JOIN {ctn} B ON A.CAPWORK = B.WORKCENTER "
            f"WHERE A.STAND = '{costcenter}' AND A.CAPWORK = '{workcenter}'{where_extra} "
            f"GROUP BY A.STAND,A.CAPWORK"
        )
        sum_df = ag_instance.run_query(ihtiyac_sql)
        sum_df_cap_work = ag_instance.run_query(cap_work_sql)
        if sum_df is None or sum_df.empty or sum_df_cap_work is None or sum_df_cap_work.empty:
            return None

        if not prediv_ihtiyac:
            if "hours" in selected_units:
                _div_safe(sum_df, 1, 60)
            elif "shifts" in selected_units:
                _div_safe(sum_df, 1, 510)

        sum_numeric = [c for c in sum_df.columns if c not in ("CAPWORK",)]
        _format_numeric_cols_by_unit(sum_df, sum_numeric, selected_units)

        sum_df["STAT"] = "Kapasite İhtiyacı"
        weeks = [c for c in kolon_list if c in sum_df.columns]
        if not weeks:
            weeks = [c for c in sum_df.columns if c not in ("CAPWORK", "STAT")]
        filtered_sum_df = sum_df[["STAT"] + weeks].copy()

        if sum_df_cap_work.shape[1] > 0:
            _apply_cap_work_unit_like_dash(sum_df_cap_work, selected_units)
            cap_work_numeric = [c for c in sum_df_cap_work.columns if c not in ("STAND", "CAPWORK")]
            _format_numeric_cols_by_unit(sum_df_cap_work, cap_work_numeric, selected_units)

        toplam_row = {"STAT": "Toplam Kapasite"}
        toplam_row.update(sum_df_cap_work.iloc[0].to_dict())
        cap_df = pd.concat([filtered_sum_df, pd.DataFrame([toplam_row])], ignore_index=True)

        fark_row = {"STAT": "Kapasite Farkı"}
        su = selected_units or []
        sql_float_fark = _unit_decimals(selected_units) > 0 and ("hours" in su or "shifts" in su)
        for col in weeks:
            try:
                d = float(cap_df.loc[cap_df["STAT"] == "Toplam Kapasite", col].iloc[0]) - float(
                    cap_df.loc[cap_df["STAT"] == "Kapasite İhtiyacı", col].iloc[0]
                )
                if sql_float_fark and not _is_a_col(col):
                    fark_row[col] = round(d, DISPLAY_MAX_DECIMALS_NON_A)
                else:
                    fark_row[col] = round(d, 0)
            except Exception:
                fark_row[col] = 0
        cap_df = pd.concat([cap_df, pd.DataFrame([fark_row])], ignore_index=True)

        numeric_cols = [c for c in cap_df.columns if c != "STAT"]
        cumsum = cap_df.loc[cap_df["STAT"] == "Kapasite Farkı", numeric_cols].cumsum(axis=1)
        cum_row = {"STAT": "Kümülatif Toplam"}
        cum_raw = cumsum.iloc[0].to_dict()
        cum_row.update(cum_raw)
        for k, v in list(cum_row.items()):
            if k == "STAT" or not isinstance(v, (int, float)):
                continue
            if sql_float_fark and not _is_a_col(k):
                cum_row[k] = round(v, DISPLAY_MAX_DECIMALS_NON_A)
            else:
                cum_row[k] = round(v, 0)
        cap_df = pd.concat([cap_df, pd.DataFrame([cum_row])], ignore_index=True)

        ui_row = cap_df[cap_df["STAT"] == "Kapasite İhtiyacı"].iloc[0]
        tk_row = cap_df[cap_df["STAT"] == "Toplam Kapasite"].iloc[0]
        doluluk_vals = []
        for col in weeks:
            try:
                tk = float(tk_row[col])
                ui = float(ui_row[col])
                doluluk_vals.append(round((ui / tk) * 100, 0) if tk != 0 else 0)
            except Exception:
                doluluk_vals.append(0)
        doluluk_df = pd.DataFrame([doluluk_vals], columns=weeks)
        doluluk_df["STAT"] = "Doluluk Oranı(%)"
        return _finalize_capacity_stat_like_dash(cap_df, doluluk_df, weeks, selected_units)
    except Exception as e:
        print(f"kapasite_data build_capacity_table_for_cc_workcenter hatası: {e}")
        return None


def build_workcenter_yuk_table_for_cc(ag_instance, table_name, costcenter, kolon_sum_str, selected_units=("hours",), capacity_table_name=None):
   
    if not costcenter or not table_name or not kolon_sum_str:
        return None
    try:
        sql_yuk = (
            f"SELECT CAPWORK, {kolon_sum_str} FROM [{table_name}] "
            f"WHERE STAND = '{costcenter}' GROUP BY CAPWORK ORDER BY CAPWORK"
        )
        df = ag_instance.run_query(sql_yuk)
        if df is None or df.empty:
            return None

        su = selected_units or []
        if "hours" in su:
            _div_safe(df, 1, 60)
        elif "shifts" in su:
            _div_safe(df, 1, 510)
        else:
            try:
                df = df.round(0)
            except Exception:
                pass

        verim_df = get_verimlilik_df(ag_instance)
        if verim_df is not None:
            df = df.merge(verim_df, left_on="CAPWORK", right_on="WORKCENTER", how="left")
            df = df.drop(columns=["WORKCENTER"], errors="ignore")
            df["Verimlilik"] = pd.to_numeric(df["Verimlilik"], errors="coerce")
        else:
            df["Verimlilik"] = None

        wc_numeric_cols = [c for c in df.columns if c != "CAPWORK"]
        if wc_numeric_cols:
            _format_numeric_cols_by_unit(df, wc_numeric_cols, selected_units)
        return df
    except Exception as e:
        print(f"kapasite_data build_workcenter_yuk_table_for_cc hatası: {e}")
        return None


def build_malzeme_table(
    ag_instance,
    table_name,
    costcenter,
    kolon_sum_str,
    selected_units=("hours",),
    *,
    kolon_list=None,
    cap_grp=None,
    workcenter=None,
    for_report=False,
):
    
    if not costcenter or not table_name:
        return None
    if isinstance(kolon_list, str):
        kl = [x.strip() for x in kolon_list.split(",") if x.strip()]
    elif kolon_list:
        kl = list(kolon_list)
    else:
        kl = []
    kss = (kolon_sum_str or "").strip() if isinstance(kolon_sum_str, str) else ""
    sum_select = _select_sum_with_unit(kl, selected_units) if kl else ""
    sql_agg = sum_select.strip() if sum_select else kss
    if not sql_agg:
        return None
    uses_sql_div = bool(sum_select.strip())

    tn = _sql_bracket_table(table_name)
    where = _malzeme_where_clause(costcenter, cap_grp, workcenter)
    if for_report:
        group_sel = "MATERIAL, DRAWNUM, MACHINE, BASEQUAN, MTUNIT"
        sql = (
            f"SELECT {group_sel}, {sql_agg} FROM {tn} {where} "
            f"GROUP BY {group_sel} ORDER BY {group_sel}"
        )
    else:
        sql = (
            f"SELECT MATERIAL, DRAWNUM, {sql_agg} FROM {tn} {where} "
            f"GROUP BY MATERIAL, DRAWNUM ORDER BY MATERIAL, DRAWNUM"
        )
    try:
        df = ag_instance.run_query(sql)
    except Exception as e:
        print(f"kapasite_data build_malzeme_table sorgu hatası: {e}")
        return None
    if df is None or df.empty:
        return None
    if not uses_sql_div:
        start_i = _malzeme_div_start_col_index(df)
        if "hours" in (selected_units or []):
            _div_safe(df, start_i, 60)
        elif "shifts" in (selected_units or []):
            _div_safe(df, start_i, 510)
    num_cols = [c for c in df.columns if c not in _id_columns_malzeme(df)]
    if num_cols:
        _format_numeric_cols_by_unit(df, num_cols, selected_units)
    return df


def build_malzeme_table_for_cc(
    ag_instance,
    table_name,
    costcenter,
    kolon_sum_str,
    selected_units=("hours",),
    for_report=False,
    *,
    kolon_list=None,
    cap_grp=None,
):
    return build_malzeme_table(
        ag_instance,
        table_name,
        costcenter,
        kolon_sum_str,
        selected_units,
        kolon_list=kolon_list,
        cap_grp=cap_grp,
        workcenter=None,
        for_report=for_report,
    )


def build_malzeme_table_for_cc_workcenter(
    ag_instance,
    table_name,
    costcenter,
    workcenter,
    kolon_sum_str,
    selected_units=("hours",),
    for_report=False,
    *,
    kolon_list=None,
    cap_grp=None,
):
    if not workcenter:
        return None
    return build_malzeme_table(
        ag_instance,
        table_name,
        costcenter,
        kolon_sum_str,
        selected_units,
        kolon_list=kolon_list,
        cap_grp=cap_grp,
        workcenter=workcenter,
        for_report=for_report,
    )
