# =============================================================================
# excel_report.py — Generare rapoarte Excel (foi, grafice, formatare)
# =============================================================================
# Exportă:
#   add_charts_and_formatting(excel_path, df, site_id)
# =============================================================================

import os
import calendar
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.worksheet.datavalidation import DataValidation

from config import (
    VEHICLE_ANALYSIS, MIN_LUNI_AN, MIN_LUNI_AN_MAI, MIN_ORE_ZI, BAND_COLORS,
    CENTRAL_FILE_FOLDER, CENTRAL_FILE_NAME,
)

def add_charts_and_formatting(excel_path, df, site_id):
    MIN_ORE_ZI = 22
    MIN_ZILE_LUNA = 15
    MIN_ZILE_SAPT = 7
    # importăm globalele — deja definite la nivel de modul
    _MIN_LUNI_AN = MIN_LUNI_AN
    _MIN_LUNI_AN_MAI = MIN_LUNI_AN_MAI

    C_DARK   = "1F4E79"
    C_MID    = "2E75B6"
    C_LIGHT  = "D6E4F0"
    C_WHITE  = "FFFFFF"
    C_GREEN  = "E2EFDA"
    C_ORANGE = "FCE4D6"
    C_YELLOW = "FFF2CC"

    # Culori benzi
    BAND_COLORS = ["2E75B6", "ED7D31", "70AD47", "FFC000", "5B9BD5", "FF0000"]

    def fill(hex_c):
        return PatternFill('solid', start_color=hex_c)

    thin = Side(style='thin', color='BFBFBF')
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hfont(sz=10):
        return Font(name='Arial', size=sz, bold=True, color=C_WHITE)

    def dfont(sz=10, bold=False, color="1F1F1F"):
        return Font(name='Arial', size=sz, bold=bold, color=color)

    def ctr():
        return Alignment(horizontal='center', vertical='center', wrap_text=True)

    def rgt():
        return Alignment(horizontal='right', vertical='center')

    days_ro = ['Luni', 'Marți', 'Miercuri', 'Joi', 'Vineri', 'Sâmbătă', 'Duminică']

    wb = load_workbook(excel_path)

    df = df.copy()  # defragmentare DataFrame (evită PerformanceWarning)
    df['Data'] = pd.to_datetime(df['Data_Ora'], format='%d.%m.%Y %H:%M')
    df['Zi']   = df['Data'].dt.date
    df['Ora']  = df['Data'].dt.hour
    df['An']   = df['Data'].dt.year
    df['Luna'] = df['Data'].dt.month

    # ── Detectare dinamică număr de benzi ───────────────────────────────────
    # Numărăm câte coloane Total_BX există în DataFrame
    n_bands = sum(1 for b in range(1, 7) if f'Total_B{b}' in df.columns)
    if n_bands not in (2, 4, 6):
        n_bands = 2  # fallback sigur
    band_ids = list(range(1, n_bands + 1))

    # Sens 1 = prima jumătate de benzi, Sens 2 = a doua jumătate
    sens1_bands = band_ids[:n_bands // 2]
    sens2_bands = band_ids[n_bands // 2:]

    # Coloana Total_General = suma tuturor benzilor
    total_band_cols = [f'Total_B{b}' for b in band_ids]
    if 'Total_General' not in df.columns:
        df['Total_General'] = df[total_band_cols].sum(axis=1)

    # Totaluri pe sensuri
    df['Total_Sens1'] = df[[f'Total_B{b}' for b in sens1_bands]].sum(axis=1)
    df['Total_Sens2'] = df[[f'Total_B{b}' for b in sens2_bands]].sum(axis=1)

    agg_dict = {col: 'sum' for col in total_band_cols}
    agg_dict['Total_General'] = 'sum'
    agg_dict['Total_Sens1'] = 'sum'
    agg_dict['Total_Sens2'] = 'sum'

    daily_data = df.groupby('Zi').agg(agg_dict).reset_index()
    for b in band_ids:
        daily_data[f'Peak_B{b}'] = df.groupby('Zi')[f'Total_B{b}'].max().values
    daily_data['Ore'] = df.groupby('Zi').size().values

    hourly_avg_dict = {f'Total_B{b}': 'mean' for b in band_ids}
    hourly_avg_dict['Total_Sens1'] = 'mean'
    hourly_avg_dict['Total_Sens2'] = 'mean'
    hourly_avg = df.groupby('Ora').agg(hourly_avg_dict).reset_index()

    n_days         = len(daily_data)
    latime_dinamica = max(20, n_days * 0.6)

    # ── FOAIE 2: Rezumat Zilnic ──────────────────────────────────────────────
    ws2 = wb.create_sheet("Rezumat Zilnic")
    ws2.freeze_panes = "C3"

    ore_pe_zi = df.groupby('Zi').size().reset_index(name='Ore_pe_zi')

    # Coloane clase existente pentru fiecare bandă
    class_cols_all = {}
    for b in band_ids:
        class_cols_all[b] = [f'B{b}_Clasa_{i}' for i in range(1, 9)] + [f'B{b}_Clasa_15']

    agg_class = {}
    for b in band_ids:
        for col in class_cols_all[b]:
            if col in df.columns:
                agg_class[col] = 'sum'
    for b in band_ids:
        agg_class[f'Total_B{b}'] = 'sum'

    daily_classes = df.groupby('Zi').agg(agg_class).reset_index()
    daily_classes = daily_classes.merge(ore_pe_zi, on='Zi')

    # Suma clase per tip (toate benzile)
    for i in list(range(1, 9)) + [15]:
        cols_ci = [f'B{b}_Clasa_{i}' for b in band_ids if f'B{b}_Clasa_{i}' in daily_classes.columns]
        daily_classes[f'Clasa_{i}'] = daily_classes[cols_ci].sum(axis=1) if cols_ci else 0

    clasa_cols_for_total = [f'Clasa_{i}' for i in list(range(1, 9)) + [15]
                            if f'Clasa_{i}' in daily_classes.columns]
    daily_classes['Total'] = daily_classes[clasa_cols_for_total].sum(axis=1)

    mask_insuficient = daily_classes['Ore_pe_zi'] < 22
    zero_cols = [f'Clasa_{i}' for i in list(range(1, 9)) + [15]] + ['Total']
    for b in band_ids:
        zero_cols += [f'Total_B{b}']
    for col in zero_cols:
        if col in daily_classes.columns:
            daily_classes.loc[mask_insuficient, col] = 0

    for b in band_ids:
        daily_classes[f'Peak_B{b}'] = df.groupby('Zi')[f'Total_B{b}'].max().values
        daily_classes.loc[mask_insuficient, f'Peak_B{b}'] = 0

    first_date = daily_classes['Zi'].min().strftime('%d.%m.%Y')
    last_date  = daily_classes['Zi'].max().strftime('%d.%m.%Y')

    # --- Extragere Localitate din Centralizator ---
    localitate_site = ""
    try:
        path_centralizator = os.path.join(CENTRAL_FILE_FOLDER, CENTRAL_FILE_NAME)
        if os.path.exists(path_centralizator):
            df_contoare = pd.read_excel(path_centralizator, sheet_name='Contoare')
            match = df_contoare[df_contoare.iloc[:, 0].astype(str).str.contains(str(site_id))]
            if not match.empty:
                localitate_site = f" - {match.iloc[0, 3]}"
    except Exception as e:
        print(f"Atenție: Nu s-a putut citi localitatea din centralizator: {e}")

    # ── Construire headers dinamici pentru Rezumat Zilnic ───────────────────
    # Structura: Data | Zi | Clase 1-8 | Clasa 15 | TOTAL | Total_B1..BN | Varf_B1..BN | B1%..BN% | Indicator | Ore
    # 2 benzi: 2 + 9 + 1 + N + N + N + 2 = 2+9+1+2+2+2+2 = 20 col
    # 4 benzi: 2+9+1+4+4+4+2 = 26 col
    # 6 benzi: 2+9+1+6+6+6+2 = 32 col

    headers = ["Data", "Zi săptămână",
               "Clasa 1","Clasa 2","Clasa 3","Clasa 4","Clasa 5",
               "Clasa 6","Clasa 7","Clasa 8","Clasa 15",
               "TOTAL"]
    for b in band_ids:
        headers.append(f"Total Banda {b}")
    for b in band_ids:
        headers.append(f"Vârf Banda {b}")
    for b in band_ids:
        headers.append(f"Banda {b} %")
    headers += ["Indicator", "Ore înregistrate"]

    n_total_cols = len(headers)
    total_col_idx = 12  # coloana L (1-based) = TOTAL
    # Coloanele Total_BX: de la col 13 la 12+n_bands
    total_b_start_col = 13
    # Coloanele Varf_BX: de la col 13+n_bands la 12+2*n_bands
    varf_b_start_col = total_b_start_col + n_bands
    # Coloanele Proc_BX: de la col 13+2*n_bands la 12+3*n_bands
    proc_b_start_col = varf_b_start_col + n_bands
    # Indicator, Ore
    indicator_col = proc_b_start_col + n_bands
    ore_col = indicator_col + 1

    # Titlu dinamic care acoperă toate coloanele
    title_end_col = get_column_letter(n_total_cols)
    ws2.merge_cells(f"A1:{title_end_col}1")
    ws2["A1"] = f"REZUMAT ZILNIC  |  Contor: {site_id}{localitate_site}  |  {first_date} – {last_date}"
    ws2["A1"].font = Font(name='Arial', size=13, bold=True, color=C_WHITE)
    ws2["A1"].fill = fill(C_DARK)
    ws2["A1"].alignment = ctr()
    ws2.row_dimensions[1].height = 28

    for c, h in enumerate(headers, 1):
        cell = ws2.cell(2, c, h)
        cell.font = hfont(9)
        cell.fill = fill(C_MID)
        cell.alignment = ctr()
        cell.border = brd
    ws2.row_dimensions[2].height = 32

    # Calculăm ore cu trafic REAL
    ore_cu_trafic = (
        df[df['Total_General'] > 0]
        .groupby('Zi')
        .size()
        .reset_index(name='Ore_cu_trafic')
    )
    daily_classes = daily_classes.merge(ore_cu_trafic, on='Zi', how='left')
    daily_classes['Ore_cu_trafic'] = daily_classes['Ore_cu_trafic'].fillna(0).astype(int)

    dr = 3
    for i, row in daily_classes.iterrows():
        dt = row['Zi']
        day_name = days_ro[dt.weekday()]
        ore_cu_trafic_zi = int(row['Ore_cu_trafic'])

        if ore_cu_trafic_zi == 0:
            indicator = "Nu există date"
            rf = fill(C_ORANGE)
        elif ore_cu_trafic_zi < MIN_ORE_ZI:
            indicator = "Date parțiale"
            rf = fill(C_YELLOW)
            for col in zero_cols:
                if col in row.index:
                    row[col] = 0
        elif ore_cu_trafic_zi == 24:
            indicator = "Date complete"
            rf = fill(C_LIGHT) if i % 2 == 0 else fill(C_WHITE)
        else:
            indicator = "Date parțiale"
            rf = fill(C_YELLOW)

        # Construire valori rând
        vals = [dt.strftime("%d.%m.%Y"), day_name]
        for ci in list(range(1, 9)) + [15]:
            vals.append(int(row.get(f'Clasa_{ci}', 0)))
        vals.append(int(row.get('Total', 0)))  # TOTAL col 12
        total_col_letter = get_column_letter(total_col_idx)
        for b in band_ids:
            vals.append(int(row.get(f'Total_B{b}', 0)))
        for b in band_ids:
            vals.append(int(row.get(f'Peak_B{b}', 0)))
        for bi, b in enumerate(band_ids):
            b_col = get_column_letter(total_b_start_col + bi)
            vals.append(f"=IFERROR({b_col}{dr}/{total_col_letter}{dr},0)")
        vals.append(indicator)
        vals.append(f"{ore_cu_trafic_zi}/24")

        for c, val in enumerate(vals, 1):
            cell = ws2.cell(dr, c, val)
            cell.border = brd
            cell.alignment = ctr() if c <= 2 or c >= indicator_col else rgt()
            if 3 <= c <= total_col_idx + n_bands * 2:
                cell.number_format = '#,##0'
            if proc_b_start_col <= c <= proc_b_start_col + n_bands - 1:
                cell.number_format = '0.0%'
            if ore_cu_trafic_zi == 0:
                cell.font = dfont(9, bold=True, color="C00000")
                cell.fill = fill(C_ORANGE)
            else:
                cell.font = dfont(9)
                cell.fill = rf
        dr += 1

    last_row = dr - 1
    ws2.merge_cells(f"A{dr}:B{dr}")
    ws2[f"A{dr}"] = "TOTAL PERIOADĂ"
    ws2[f"A{dr}"].font = dfont(bold=True)
    ws2[f"A{dr}"].fill = fill(C_GREEN)
    ws2[f"A{dr}"].alignment = ctr()
    ws2[f"A{dr}"].border = brd
    ws2[f"B{dr}"].fill = fill(C_GREEN)
    ws2[f"B{dr}"].border = brd

    for c in range(3, n_total_cols + 1):
        cl = get_column_letter(c)
        is_varf = varf_b_start_col <= c <= varf_b_start_col + n_bands - 1
        is_proc = proc_b_start_col <= c <= proc_b_start_col + n_bands - 1
        is_text = c >= indicator_col
        if is_varf:
            formula = f"=MAX({cl}3:{cl}{last_row})"
        elif is_proc:
            b_offset = c - proc_b_start_col
            b_col = get_column_letter(total_b_start_col + b_offset)
            formula = f"=IFERROR({b_col}{dr}/{get_column_letter(total_col_idx)}{dr},0)"
        elif is_text:
            formula = ""
        else:
            formula = f"=SUM({cl}3:{cl}{last_row})"
        cell = ws2.cell(dr, c, formula)
        cell.font = dfont(bold=True)
        cell.fill = fill(C_GREEN)
        cell.alignment = rgt() if not is_text else ctr()
        cell.border = brd
        if 3 <= c <= varf_b_start_col + n_bands - 1:
            cell.number_format = '#,##0'
        elif is_proc:
            cell.number_format = '0.0%'

    widths = [12, 13] + [9] * 9 + [11] + [13] * n_bands + [13] * n_bands + [11] * n_bands + [14, 14]
    for c, w in enumerate(widths, 1):
        ws2.column_dimensions[get_column_letter(c)].width = w

    def make_bar_layout():
        return Layout(manualLayout=ManualLayout(x=0.05, y=0.10, h=0.75, w=0.90,
                                                xMode="edge", yMode="edge"))

    def configure_x_axis(ax):
        ax.delete = False
        ax.axPos  = "b"
        ax.tickLblPos  = "nextTo"
        ax.textRotation = -45000
        ax.tickLblSkip  = 1

    # ── Pregătire date pentru Pie Charts — ultimele 30 zile valide ──────────
    dc_valide = daily_classes[~mask_insuficient].copy()
    dc_pie    = dc_valide.tail(30)

    clase_pie = [1, 2, 3, 4, 5, 6, 7, 8, 15]
    label_pie = [f"Clasa {i}" for i in clase_pie]

    # Sume pe clase pentru fiecare bandă
    sume_per_banda = {}
    for b in band_ids:
        sume_b = []
        for i in clase_pie:
            col = f'B{b}_Clasa_{i}'
            sume_b.append(int(dc_pie[col].sum()) if col in dc_pie.columns else 0)
        sume_per_banda[b] = sume_b

    # Plasăm datele auxiliare pentru pie charts
    # Coloana start: după ultimele coloane principale (n_total_cols + 2)
    PIE_AUX_START_COL = n_total_cols + 2  # lăsăm un spațiu
    PIE_START_ROW = 3
    n_clase = len(clase_pie)

    # Header etichete (comune)
    lbl_col = PIE_AUX_START_COL
    ws2.cell(PIE_START_ROW - 1, lbl_col, "Clasa").font = dfont(bold=True)
    for b in band_ids:
        val_col = PIE_AUX_START_COL + b
        ws2.cell(PIE_START_ROW - 1, val_col, f"Banda {b}").font = dfont(bold=True)

    for idx, lbl in enumerate(label_pie):
        r = PIE_START_ROW + idx
        ws2.cell(r, lbl_col, lbl)
        for b in band_ids:
            ws2.cell(r, PIE_AUX_START_COL + b, sume_per_banda[b][idx])

    # Culori fixe pentru clase
    CULORI_PIE = {
        0: "A9CCE3", 1: "27AE60", 2: "F39C12", 3: "8E44AD",
        4: "87CEFA", 5: "1ABC9C", 6: "2E75B6", 7: "F7DC6F", 8: "C0392B",
    }

    def make_pie(title, val_col, n):
        pie = PieChart()
        pie.title  = title
        pie.style  = 10
        pie.width  = 16
        pie.height = 16

        data = Reference(ws2, min_col=val_col, min_row=PIE_START_ROW - 1,
                         max_row=PIE_START_ROW + n - 1)
        cats = Reference(ws2, min_col=lbl_col,  min_row=PIE_START_ROW,
                         max_row=PIE_START_ROW + n - 1)
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)

        serie = pie.series[0]
        for idx_c in range(n):
            pt = DataPoint(idx=idx_c)
            hex_c = CULORI_PIE.get(idx_c, "AAAAAA")
            pt.graphicalProperties.solidFill = hex_c
            pt.graphicalProperties.line.solidFill = "FFFFFF"
            serie.dPt.append(pt)

        serie.dLbls = DataLabelList()
        serie.dLbls.showPercent  = True
        serie.dLbls.showCatName  = False
        serie.dLbls.showVal      = False
        serie.dLbls.showSerName  = False
        serie.dLbls.showLeaderLines = True

        pie.legend.position = 'b'
        pie.legend.overlay  = False

        pie.layout = Layout(manualLayout=ManualLayout(
            x=0.05, y=0.20, h=0.65, w=0.90, xMode="edge", yMode="edge"
        ))
        return pie

    # Poziționare pie charts — pozitii fixe exacte per număr de benzi
    # 2 benzi: Banda1=V3,  Banda2=AH3  (ambele pe rândul 1)
    # 4 benzi: Banda1=AB3, Banda2=AL3, Banda3=AB35, Banda4=AL35
    # 6 benzi: Banda1=AH3, Banda2=AR3, Banda3=BB3,  Banda4=AH35, Banda5=AR35, Banda6=BB35

    if n_bands == 2:
        pie_anchors = ["V3", "AF3"]
        bar_anchor  = "V36"
    elif n_bands == 4:
        pie_anchors = ["AB3", "AL3", "AB35", "AL35"]
        bar_anchor  = "AB67"
    else:  # 6 benzi
        pie_anchors = ["AH3", "AR3", "BB3", "AH35", "AR35", "BB35"]
        bar_anchor  = "AH67"

    for b, anchor in zip(band_ids, pie_anchors):
        val_col   = PIE_AUX_START_COL + b
        pie_chart = make_pie(f"Distribuție clase — Banda {b}  (ult. 30 zile)", val_col, n_clase)
        ws2.add_chart(pie_chart, anchor)

    chart2 = BarChart()
    chart2.type = "col"; chart2.grouping = "percentStacked"; chart2.overlap = 100
    chart2.title = f"Distribuție procentuală {'  /  '.join([f'Banda {b}' for b in band_ids])}"
    chart2.y_axis.title = "Procent (%)"; chart2.style = 10
    chart2.width = latime_dinamica; chart2.height = 12
    chart2.legend.position = 'b'; chart2.legend.overlay = False
    configure_x_axis(chart2.x_axis)
    chart2.layout = make_bar_layout()

    # Date pentru chart: coloanele Total_BX din ws2
    for bi, b in enumerate(band_ids):
        col_idx = total_b_start_col + bi
        ref = Reference(ws2, min_col=col_idx, min_row=2, max_row=2 + n_days)
        chart2.add_data(ref, titles_from_data=True)
        chart2.series[bi].graphicalProperties.solidFill = BAND_COLORS[bi % len(BAND_COLORS)]

    cats_all = Reference(ws2, min_col=1, min_row=3, max_row=2 + n_days)
    chart2.set_categories(cats_all)
    ws2.add_chart(chart2, bar_anchor)

    # ── FOAIE 3: Media Zilnică Lunară ────────────────────────────────────────
    ws_lunar = wb.create_sheet("Media Zilnica Lunara")

    ore_pe_zi2  = df.groupby(['An', 'Luna', 'Zi']).size().reset_index(name='Ore')
    zile_valide = ore_pe_zi2[ore_pe_zi2['Ore'] >= MIN_ORE_ZI][['An', 'Luna', 'Zi']]
    df_valid    = df.merge(zile_valide, on=['An', 'Luna', 'Zi'], how='inner')
    toate_lunile = df.groupby(['An', 'Luna']).size().reset_index()[['An', 'Luna']]

    def check_7_consecutive_days(df_all, an, luna):
        import datetime as dt_mod
        zile_in_luna = calendar.monthrange(an, luna)[1]
        df_luna_ore  = df_all[(df_all['An'] == an) & (df_all['Luna'] == luna)]
        ore_dict     = df_luna_ore.groupby(df_luna_ore['Data'].dt.date).size().to_dict()
        for start_day in range(1, zile_in_luna - 5):
            week_dates = []
            valid_week = True
            for offset in range(7):
                current_date = dt_mod.date(an, luna, start_day + offset)
                if ore_dict.get(current_date, 0) < MIN_ORE_ZI:
                    valid_week = False
                    break
                week_dates.append(current_date)
            if valid_week:
                return True, week_dates
        return False, []

    lunar_data = []
    for _, row in toate_lunile.iterrows():
        an, luna = row['An'], row['Luna']
        df_luna_valida = df_valid[(df_valid['An'] == an) & (df_valid['Luna'] == luna)]
        nr_zile_valide = df_luna_valida['Zi'].nunique() if len(df_luna_valida) > 0 else 0
        zile_in_luna   = calendar.monthrange(an, luna)[1]

        if nr_zile_valide == zile_in_luna:
            indicator, color = "Date complete", "NORMAL"
            df_calcul = df_luna_valida; divisor = zile_in_luna
        elif nr_zile_valide >= MIN_ZILE_LUNA:
            indicator, color = "Date parțiale", "YELLOW"
            df_calcul = df_luna_valida; divisor = nr_zile_valide
        else:
            has_7, week_dates = check_7_consecutive_days(df, an, luna)
            if has_7:
                indicator, color = "Date parțiale - 7 zile", "YELLOW"
                df_calcul = df[df['Zi'].isin(week_dates)]; divisor = 7
            else:
                indicator, color = "Nu există date", "ORANGE"
                df_calcul = pd.DataFrame(); divisor = 0

        if divisor > 0 and not df_calcul.empty:
            sume = {}
            for i in range(1, 9):
                sume[f'Clasa_{i}'] = (df_calcul[f'B1_Clasa_{i}'].sum() +
                                       df_calcul[f'B2_Clasa_{i}'].sum())
            sume['Clasa_15'] = (df_calcul['B1_Clasa_15'].sum() +
                                 df_calcul['B2_Clasa_15'].sum())
            medii = {f'Clasa_{i}': round(sume[f'Clasa_{i}'] / divisor) for i in range(1, 9)}
            medii['Clasa_15'] = round(sume['Clasa_15'] / divisor)
            medii['Total']    = sum(medii.values())
        else:
            medii = {f'Clasa_{i}': 0 for i in range(1, 9)}
            medii['Clasa_15'] = 0; medii['Total'] = 0

        lunar_data.append({
            'Post': site_id, 'An': an, 'Luna': luna,
            **medii, 'Indicator': indicator, 'Color': color,
            'Zile_cu_inregistrari': f"{nr_zile_valide}/{zile_in_luna}",
            'Are_date_valide': (divisor > 0)  # True dacă luna are date suficiente
        })

    ws_lunar.merge_cells("A1:O1")
    ws_lunar["A1"] = (f"MEDIE ZILNICĂ LUNARĂ  |  Contor: {site_id}  |  "
                      f"Minim {MIN_ORE_ZI}h/zi, minim {MIN_ZILE_LUNA} zile/lună")
    ws_lunar["A1"].font = Font(name='Arial', size=13, bold=True, color=C_WHITE)
    ws_lunar["A1"].fill = fill(C_DARK)
    ws_lunar["A1"].alignment = ctr()
    ws_lunar.row_dimensions[1].height = 28

    headers_lunar = ["Post","An","Luna","Clasa 1","Clasa 2","Clasa 3","Clasa 4",
                     "Clasa 5","Clasa 6","Clasa 7","Clasa 8","Clasa 15",
                     "Total","Indicator","Zile cu înregistrări"]
    for c, h in enumerate(headers_lunar, 1):
        cell = ws_lunar.cell(2, c, h)
        cell.font = hfont(9); cell.fill = fill(C_MID)
        cell.alignment = ctr(); cell.border = brd
    ws_lunar.row_dimensions[2].height = 28

    luna_nume = {1:'Ian',2:'Feb',3:'Mar',4:'Apr',5:'Mai',6:'Iun',
                 7:'Iul',8:'Aug',9:'Sep',10:'Oct',11:'Noi',12:'Dec'}

    dr_lunar = 3
    for i, rd in enumerate(lunar_data):
        if rd['Color'] == 'YELLOW':
            rf = fill(C_YELLOW)
        elif rd['Color'] == 'ORANGE':
            rf = fill(C_ORANGE)
        else:
            rf = fill(C_LIGHT) if i % 2 == 0 else fill(C_WHITE)

        vals = [rd['Post'], rd['An'], luna_nume[rd['Luna']],
                rd['Clasa_1'], rd['Clasa_2'], rd['Clasa_3'], rd['Clasa_4'],
                rd['Clasa_5'], rd['Clasa_6'], rd['Clasa_7'], rd['Clasa_8'],
                rd['Clasa_15'], rd['Total'], rd['Indicator'], rd['Zile_cu_inregistrari']]
        for c, val in enumerate(vals, 1):
            cell = ws_lunar.cell(dr_lunar, c, val)
            cell.font = dfont(9); cell.fill = rf; cell.border = brd
            cell.alignment = ctr() if c <= 3 or c >= 14 else rgt()
            if 4 <= c <= 13:
                cell.number_format = '#,##0'
        dr_lunar += 1

    widths_lunar = [10, 8, 8] + [10] * 9 + [12, 14, 16]
    for c, w in enumerate(widths_lunar, 1):
        ws_lunar.column_dimensions[get_column_letter(c)].width = w

    # Grafic lunar
    lunar_df  = pd.DataFrame(lunar_data)
    ani_unici = sorted(lunar_df['An'].unique())
    start_chart_row = dr_lunar + 3

    ws_lunar.cell(start_chart_row, 1, "Luna").font = Font(bold=True)
    for idx_an, an in enumerate(ani_unici, 2):
        cell = ws_lunar.cell(start_chart_row, idx_an, an)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    luna_nome_full = ['Ianuarie','Februarie','Martie','Aprilie','Mai','Iunie',
                      'Iulie','August','Septembrie','Octombrie','Noiembrie','Decembrie']
    for luna_nr in range(1, 13):
        r_idx = start_chart_row + luna_nr
        ws_lunar.cell(r_idx, 1, luna_nome_full[luna_nr - 1])
        for idx_an, an in enumerate(ani_unici, 2):
            row_v   = lunar_df[(lunar_df['An'] == an) & (lunar_df['Luna'] == luna_nr)]
            valoare = int(row_v.iloc[0]['Total']) if len(row_v) > 0 else 0
            ws_lunar.cell(r_idx, idx_an, valoare).number_format = '#,##0'

    cats_l = Reference(ws_lunar, min_col=1, min_row=start_chart_row+1, max_row=start_chart_row+12)
    data_l = Reference(ws_lunar, min_col=2, max_col=1+len(ani_unici),
                       min_row=start_chart_row, max_row=start_chart_row+12)

    chart_lunar = LineChart()
    chart_lunar.title  = "Media Zilnică Lunară - Total vehicule"
    chart_lunar.style  = 10; chart_lunar.width = 25; chart_lunar.height = 14
    chart_lunar.add_data(data_l, titles_from_data=True, from_rows=False)
    chart_lunar.set_categories(cats_l)
    chart_lunar.x_axis.delete = False; chart_lunar.y_axis.delete = False
    chart_lunar.legend.position = 'b'; chart_lunar.legend.overlay = False
    chart_lunar.x_axis.axPos = "b"; chart_lunar.x_axis.tickLblPos = "nextTo"
    chart_lunar.x_axis.textRotation = -45000
    chart_lunar.y_axis.scaling.min = 0
    chart_lunar.y_axis.title = "Număr vehicule"
    chart_lunar.layout = Layout(manualLayout=ManualLayout(x=0.1, y=0.12, h=0.7, w=0.85,
                                                          xMode="edge", yMode="edge"))
    culori_ani = ['2E75B6','ED7D31','70AD47','FFC000','5B9BD5']
    for idx, serie in enumerate(chart_lunar.series):
        serie.graphicalProperties.line.solidFill = culori_ani[idx % len(culori_ani)]
        serie.graphicalProperties.line.width = 25000
        serie.smooth = True; serie.marker.symbol = "circle"; serie.marker.size = 7
        serie.marker.graphicalProperties.solidFill = culori_ani[idx % len(culori_ani)]
        serie.marker.graphicalProperties.line.solidFill = culori_ani[idx % len(culori_ani)]
    ws_lunar.add_chart(chart_lunar, "Q3")

    # ── FOAIE 4: Profil Orar Mediu ───────────────────────────────────────────
    ws3 = wb.create_sheet("Profil Orar Mediu")

    # Coloane: Ora | Media_B1..BN | Media Total | B1%..BN%
    # 2 benzi: 1+2+1+2 = 6 col
    # 4 benzi: 1+4+1+4 = 10 col
    # 6 benzi: 1+6+1+6 = 14 col
    n_cols_ws3 = 1 + n_bands + 1 + n_bands  # Ora + medii + total + procente
    ws3_title_end = get_column_letter(n_cols_ws3)
    ws3.merge_cells(f"A1:{ws3_title_end}1")
    ws3["A1"] = f"PROFIL ORAR MEDIU (toate zilele)  |  Contor: {site_id}"
    ws3["A1"].font = Font(name='Arial', size=13, bold=True, color=C_WHITE)
    ws3["A1"].fill = fill(C_DARK); ws3["A1"].alignment = ctr()
    ws3.row_dimensions[1].height = 28

    h3 = ["Ora"] + [f"Media Banda {b}" for b in band_ids] + ["Media Total"] + \
         [f"Banda {b} %" for b in band_ids]
    for c, h in enumerate(h3, 1):
        cell = ws3.cell(2, c, h)
        cell.font = hfont(); cell.fill = fill(C_MID)
        cell.alignment = ctr(); cell.border = brd
    ws3.row_dimensions[2].height = 28

    # Coloane: Ora=1, Media_B1=2..n_bands+1, Total=n_bands+2, Proc_B1=n_bands+3..
    media_start = 2
    total_col_ws3 = media_start + n_bands  # ex: 4 benzi -> col 6
    proc_start = total_col_ws3 + 1

    # Formula total = SUM(B..E) la ora respectivă
    for row_i in range(24):
        dr3    = row_i + 3
        rf     = fill(C_LIGHT) if row_i % 2 == 0 else fill(C_WHITE)

        vals = [f"{row_i:02d}:00"]
        for b in band_ids:
            hr_data = hourly_avg[hourly_avg['Ora'] == row_i]
            avg_b = int(round(hr_data[f'Total_B{b}'].values[0])) if len(hr_data) > 0 else 0
            vals.append(avg_b)

        # Media Total = formula SUM
        band_cols_letters = [get_column_letter(media_start + bi) for bi in range(n_bands)]
        total_formula = "=" + "+".join([f"{cl}{dr3}" for cl in band_cols_letters])
        vals.append(total_formula)

        total_cl = get_column_letter(total_col_ws3)
        for bi in range(n_bands):
            b_cl = get_column_letter(media_start + bi)
            vals.append(f"=IFERROR({b_cl}{dr3}/{total_cl}{dr3},0)")

        for c, val in enumerate(vals, 1):
            cell = ws3.cell(dr3, c, val)
            cell.font = dfont(); cell.fill = rf; cell.border = brd
            cell.alignment = ctr() if c == 1 else rgt()
            if 2 <= c <= total_col_ws3: cell.number_format = '#,##0'
            if c >= proc_start:         cell.number_format = '0.0%'

    # Ora 50
    ani_pentru_ora50 = sorted(df['An'].unique())
    dr_ora50_start   = 28
    ws3.cell(dr_ora50_start - 1, 1, "Top Trafic - Ora 50 (pe an)").font = Font(bold=True)
    thin2 = Side(style='thin', color='BFBFBF')
    brd2  = Border(left=thin2, right=thin2, top=thin2, bottom=thin2)

    for idx, an_curent in enumerate(ani_pentru_ora50):
        df_an = df[df['An'] == an_curent].copy()
        if len(df_an) < 50: continue
        df_sortat  = df_an.sort_values(by='Total_General', ascending=False).reset_index(drop=True)
        ora_50_row = df_sortat.iloc[49]
        ora_nr     = int(pd.to_datetime(ora_50_row['Data']).hour)
        ora_data_str = pd.to_datetime(ora_50_row['Data']).strftime('%d.%m.%Y %H:00')
        ora_total  = int(ora_50_row['Total_General'])
        rind_curent = dr_ora50_start + idx

        cell_desc = ws3.cell(rind_curent, 1)
        cell_desc.value = f"Ora 50 anul {an_curent}\n({ora_nr:02d}:00 - {ora_data_str})"
        cell_desc.font  = Font(name='Arial', size=9, bold=True, color="FFFFFF")
        cell_desc.fill  = PatternFill('solid', start_color="ED7D31")
        cell_desc.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
        cell_desc.border = brd2
        ws3.row_dimensions[rind_curent].height = 45

        # Valorile pe fiecare bandă + total + procente
        ora_vals_banda = []
        for b in band_ids:
            ora_vals_banda.append(int(ora_50_row.get(f'Total_B{b}', 0)))
        ora_vals_banda.append(ora_total)
        for b_idx in range(n_bands):
            v = ora_vals_banda[b_idx]
            ora_vals_banda.append(v / ora_total if ora_total > 0 else 0)

        for col_idx, valoare in enumerate(ora_vals_banda, 2):
            cell = ws3.cell(rind_curent, col_idx, valoare)
            cell.font = Font(name='Arial', size=9, bold=True)
            cell.border = brd2
            cell.alignment = Alignment(horizontal='right', vertical='center')
            if col_idx == total_col_ws3:  # col Total
                cell.fill = PatternFill('solid', start_color="FCE4D6")
                cell.font = Font(name='Arial', size=9, bold=True, color="C65911")
            is_proc_col = col_idx >= proc_start
            cell.number_format = '0.0%' if is_proc_col else '#,##0'

    # Lățimi coloane ws3
    col_widths_ws3 = [20] + [16] * n_bands + [14] + [12] * n_bands
    for c, w in enumerate(col_widths_ws3, 1):
        ws3.column_dimensions[get_column_letter(c)].width = w

    # Grafic 3: profil orar line — pe SENSURI (sens1 vs sens2)
    chart3 = LineChart()
    sens1_label = " + ".join([f"B{b}" for b in sens1_bands])
    sens2_label = " + ".join([f"B{b}" for b in sens2_bands])
    chart3.title = f"Profil orar mediu — Sens 1 ({sens1_label}) vs Sens 2 ({sens2_label})"
    chart3.y_axis.title = "Vehicule / oră (medie)"; chart3.x_axis.title = "Ora"
    chart3.style = 10; chart3.width = 20; chart3.height = 14
    chart3.legend.position = 'b'; chart3.legend.overlay = False
    chart3.x_axis.delete = False; chart3.y_axis.delete = False
    chart3.x_axis.axPos = "b"; chart3.x_axis.tickLblPos = "nextTo"
    chart3.x_axis.textRotation = -45000; chart3.x_axis.tickLblSkip = 1
    chart3.layout = Layout(manualLayout=ManualLayout(x=0.1,y=0.12,h=0.7,w=0.85,
                                                     xMode="edge",yMode="edge"))

    # Plasăm date auxiliare pentru grafic sensuri în coloane extra ale ws3
    # IMPORTANT: calculăm valorile sens ca suma benzilor individuale deja rotunjite,
    # astfel s1 + s2 = exact suma tuturor benzilor (fără eroare de rotunjire)
    chart3_data_col_start = n_cols_ws3 + 2
    ws3.cell(2, chart3_data_col_start, "Sens 1").font = dfont(bold=True)
    ws3.cell(2, chart3_data_col_start + 1, "Sens 2").font = dfont(bold=True)
    for row_i in range(24):
        dr3 = row_i + 3
        hr_data = hourly_avg[hourly_avg['Ora'] == row_i]
        # Calculăm sens ca suma benzilor individuale rotunjite (nu din Total_Sens1/2)
        s1_val = sum(
            int(round(hr_data[f'Total_B{b}'].values[0])) if len(hr_data) > 0 else 0
            for b in sens1_bands
        )
        s2_val = sum(
            int(round(hr_data[f'Total_B{b}'].values[0])) if len(hr_data) > 0 else 0
            for b in sens2_bands
        )
        ws3.cell(dr3, chart3_data_col_start, s1_val)
        ws3.cell(dr3, chart3_data_col_start + 1, s2_val)

    ref_s1  = Reference(ws3, min_col=chart3_data_col_start,     min_row=2, max_row=26)
    ref_s2  = Reference(ws3, min_col=chart3_data_col_start + 1, min_row=2, max_row=26)
    cats3   = Reference(ws3, min_col=1, min_row=3, max_row=26)
    chart3.add_data(ref_s1, titles_from_data=True)
    chart3.add_data(ref_s2, titles_from_data=True)
    chart3.set_categories(cats3)
    chart3.series[0].graphicalProperties.line.solidFill = "2E75B6"
    chart3.series[0].graphicalProperties.line.width = 20000
    chart3.series[0].marker.symbol = "circle"; chart3.series[0].marker.size = 5
    chart3.series[1].graphicalProperties.line.solidFill = "ED7D31"
    chart3.series[1].graphicalProperties.line.width = 20000
    chart3.series[1].marker.symbol = "circle"; chart3.series[1].marker.size = 5

    chart3_anchor_col = get_column_letter(n_cols_ws3 + 5)
    ws3.add_chart(chart3, f"{chart3_anchor_col}2")

    # Grafic 4: stacked procentual orar — pe BENZI individuale
    chart4 = BarChart()
    chart4.type = "col"; chart4.grouping = "percentStacked"; chart4.overlap = 100
    chart4.title = f"Distribuție procentuală orară — {'  /  '.join([f'Banda {b}' for b in band_ids])}"
    chart4.y_axis.title = "Procent (%)"; chart4.x_axis.title = "Ora"
    chart4.style = 10; chart4.width = 20; chart4.height = 14
    chart4.legend.position = 'b'; chart4.legend.overlay = False
    chart4.x_axis.delete = False; chart4.x_axis.axPos = "b"
    chart4.x_axis.tickLblPos = "nextTo"; chart4.x_axis.textRotation = -45000
    chart4.x_axis.tickLblSkip = 1
    chart4.layout = Layout(manualLayout=ManualLayout(x=0.1,y=0.12,h=0.7,w=0.85,
                                                     xMode="edge",yMode="edge"))

    for bi in range(n_bands):
        col_idx = media_start + bi
        ref_b = Reference(ws3, min_col=col_idx, min_row=2, max_row=26)
        chart4.add_data(ref_b, titles_from_data=True)
        chart4.series[bi].graphicalProperties.solidFill = BAND_COLORS[bi % len(BAND_COLORS)]
    chart4.set_categories(cats3)

    chart4_anchor_col = get_column_letter(n_cols_ws3 + 5)
    ws3.add_chart(chart4, f"{chart4_anchor_col}30")

    # ── FOAIE 5: Profil Zilnic Mediu ─────────────────────────────────────────
    ws3z = wb.create_sheet("Profil Zilnic Mediu")

    # Structura identica cu Profil Orar Mediu, dar pe 7 zile de saptamana
    # Coloane: Zi | Media_B1..BN | Media Total | B1%..BN%
    n_cols_ws3z = 1 + n_bands + 1 + n_bands
    ws3z_title_end = get_column_letter(n_cols_ws3z)
    ws3z.merge_cells(f"A1:{ws3z_title_end}1")
    ws3z["A1"] = f"PROFIL ZILNIC MEDIU (toate săptămânile)  |  Contor: {site_id}"
    ws3z["A1"].font = Font(name='Arial', size=13, bold=True, color=C_WHITE)
    ws3z["A1"].fill = fill(C_DARK); ws3z["A1"].alignment = ctr()
    ws3z.row_dimensions[1].height = 28

    h3z = ["Zi"] + [f"Media Banda {b}" for b in band_ids] + ["Media Total"] + \
          [f"Banda {b} %" for b in band_ids]
    for c, h in enumerate(h3z, 1):
        cell = ws3z.cell(2, c, h)
        cell.font = hfont(); cell.fill = fill(C_MID)
        cell.alignment = ctr(); cell.border = brd
    ws3z.row_dimensions[2].height = 28

    total_col_ws3z = 2 + n_bands   # col Media Total
    proc_start_z   = total_col_ws3z + 1

    # Calculam media pe zi de saptamana din daily_data
    # daily_data are col 'Zi' (date) si Total_Bx
    # Adaugam coloana weekday (0=Luni..6=Duminica)
    dd_z = daily_data.copy()
    dd_z['weekday'] = pd.to_datetime(dd_z['Zi']).dt.weekday  # 0=Luni

    # Media per banda per zi de saptamana (doar zilele cu >= MIN_ORE_ZI ore)
    # Folosim df direct, care are 'Ora' si 'Zi'
    ore_pe_zi_df = df.groupby('Zi').size().reset_index(name='Ore_pe_zi')
    dd_z = dd_z.merge(ore_pe_zi_df, on='Zi', how='left')
    dd_z_valid = dd_z[dd_z['Ore_pe_zi'] >= MIN_ORE_ZI]

    weekly_avg = dd_z_valid.groupby('weekday').agg(
        {f'Total_B{b}': 'mean' for b in band_ids}
    ).reset_index()

    for row_i, day_name in enumerate(days_ro):  # 0=Luni..6=Duminica
        dr3z = row_i + 3
        rf   = fill(C_LIGHT) if row_i % 2 == 0 else fill(C_WHITE)

        vals = [day_name]
        for b in band_ids:
            wd_data = weekly_avg[weekly_avg['weekday'] == row_i]
            avg_b = int(round(wd_data[f'Total_B{b}'].values[0])) \
                    if len(wd_data) > 0 else 0
            vals.append(avg_b)

        # Media Total = formula SUM
        band_cols_z = [get_column_letter(2 + bi) for bi in range(n_bands)]
        total_formula_z = "=" + "+".join([f"{cl}{dr3z}" for cl in band_cols_z])
        vals.append(total_formula_z)

        total_cl_z = get_column_letter(total_col_ws3z)
        for bi in range(n_bands):
            b_cl = get_column_letter(2 + bi)
            vals.append(f"=IFERROR({b_cl}{dr3z}/{total_cl_z}{dr3z},0)")

        for c, val in enumerate(vals, 1):
            cell = ws3z.cell(dr3z, c, val)
            cell.font = dfont(); cell.fill = rf; cell.border = brd
            cell.alignment = ctr() if c == 1 else rgt()
            if 2 <= c <= total_col_ws3z: cell.number_format = '#,##0'
            if c >= proc_start_z:        cell.number_format = '0.0%'

    # Latimi coloane
    col_widths_z = [16] + [16] * n_bands + [14] + [12] * n_bands
    for c, w in enumerate(col_widths_z, 1):
        ws3z.column_dimensions[get_column_letter(c)].width = w

    # ── Grafic Z1: profil zilnic line — Sens 1 vs Sens 2 ─────────────────────
    chartZ1 = LineChart()
    chartZ1.title = f"Profil zilnic mediu — Sens 1 ({sens1_label}) vs Sens 2 ({sens2_label})"
    chartZ1.y_axis.title = "Vehicule / zi (medie)"
    chartZ1.x_axis.title = "Ziua săptămânii"
    chartZ1.style = 10; chartZ1.width = 20; chartZ1.height = 14
    chartZ1.legend.position = 'b'; chartZ1.legend.overlay = False
    chartZ1.x_axis.delete = False; chartZ1.y_axis.delete = False
    chartZ1.x_axis.axPos = "b"; chartZ1.x_axis.tickLblPos = "nextTo"
    chartZ1.layout = Layout(manualLayout=ManualLayout(
        x=0.1, y=0.12, h=0.7, w=0.85, xMode="edge", yMode="edge"))

    # Date auxiliare pentru sensuri in coloane extra
    cz_data_col = n_cols_ws3z + 2
    ws3z.cell(2, cz_data_col,     "Sens 1").font = dfont(bold=True)
    ws3z.cell(2, cz_data_col + 1, "Sens 2").font = dfont(bold=True)
    for row_i in range(7):
        drz = row_i + 3
        wd_data = weekly_avg[weekly_avg['weekday'] == row_i]
        s1_val = sum(
            int(round(wd_data[f'Total_B{b}'].values[0])) if len(wd_data) > 0 else 0
            for b in sens1_bands
        )
        s2_val = sum(
            int(round(wd_data[f'Total_B{b}'].values[0])) if len(wd_data) > 0 else 0
            for b in sens2_bands
        )
        ws3z.cell(drz, cz_data_col,     s1_val)
        ws3z.cell(drz, cz_data_col + 1, s2_val)

    ref_z_s1 = Reference(ws3z, min_col=cz_data_col,     min_row=2, max_row=9)
    ref_z_s2 = Reference(ws3z, min_col=cz_data_col + 1, min_row=2, max_row=9)
    cats_z   = Reference(ws3z, min_col=1, min_row=3, max_row=9)
    chartZ1.add_data(ref_z_s1, titles_from_data=True)
    chartZ1.add_data(ref_z_s2, titles_from_data=True)
    chartZ1.set_categories(cats_z)
    chartZ1.series[0].graphicalProperties.line.solidFill = "2E75B6"
    chartZ1.series[0].graphicalProperties.line.width = 20000
    chartZ1.series[0].marker.symbol = "circle"; chartZ1.series[0].marker.size = 6
    chartZ1.series[1].graphicalProperties.line.solidFill = "ED7D31"
    chartZ1.series[1].graphicalProperties.line.width = 20000
    chartZ1.series[1].marker.symbol = "circle"; chartZ1.series[1].marker.size = 6

    cz_anchor = get_column_letter(n_cols_ws3z + 5)
    ws3z.add_chart(chartZ1, f"{cz_anchor}2")

    # ── Grafic Z2: stacked procentual zilnic — benzi individuale ──────────────
    chartZ2 = BarChart()
    chartZ2.type = "col"; chartZ2.grouping = "percentStacked"; chartZ2.overlap = 100
    chartZ2.title = f"Distribuție procentuală zilnică — " \
                    f"{'  /  '.join([f'Banda {b}' for b in band_ids])}"
    chartZ2.y_axis.title = "Procent (%)"
    chartZ2.x_axis.title = "Ziua săptămânii"
    chartZ2.style = 10; chartZ2.width = 20; chartZ2.height = 14
    chartZ2.legend.position = 'b'; chartZ2.legend.overlay = False
    chartZ2.x_axis.delete = False; chartZ2.x_axis.axPos = "b"
    chartZ2.x_axis.tickLblPos = "nextTo"
    chartZ2.layout = Layout(manualLayout=ManualLayout(
        x=0.1, y=0.12, h=0.7, w=0.85, xMode="edge", yMode="edge"))

    for bi in range(n_bands):
        col_idx_z = 2 + bi
        ref_bz = Reference(ws3z, min_col=col_idx_z, min_row=2, max_row=9)
        chartZ2.add_data(ref_bz, titles_from_data=True)
        chartZ2.series[bi].graphicalProperties.solidFill = BAND_COLORS[bi % len(BAND_COLORS)]
    chartZ2.set_categories(cats_z)

    ws3z.add_chart(chartZ2, f"{cz_anchor}30")

    # ── FOAIE 6: Comparație Benzi ─────────────────────────────────────────────
    ws4 = wb.create_sheet("Comparatie Benzi")

    sens1_lbl = " + ".join([f"Banda {b}" for b in sens1_bands])
    sens2_lbl = " + ".join([f"Banda {b}" for b in sens2_bands])

    n_cols_ws4 = 2 + n_bands + 2 + 1  # Indicator + benzi + Diferenta + Raport
    ws4_title_end = get_column_letter(n_cols_ws4)
    ws4.merge_cells(f"A1:{ws4_title_end}1")
    ws4["A1"] = f"COMPARAȚIE BENZI  |  Contor: {site_id}"
    ws4["A1"].font = Font(name='Arial', size=13, bold=True, color=C_WHITE)
    ws4["A1"].fill = fill(C_DARK); ws4["A1"].alignment = ctr()
    ws4.row_dimensions[1].height = 28

    ws4.merge_cells(f"A2:{ws4_title_end}2")
    ws4["A2"] = f"Sens 1: {sens1_lbl}  |  Sens 2: {sens2_lbl}"
    ws4["A2"].font = dfont(9, color="595959"); ws4["A2"].alignment = ctr()
    ws4.row_dimensions[2].height = 20

    headers4 = ["Indicator"] + [f"Banda {b}" for b in band_ids] + \
               [f"Diferență (Sens1–Sens2)", f"Raport Sens1/Sens2"]
    for c, h in enumerate(headers4, 1):
        cell = ws4.cell(3, c, h)
        cell.font = hfont(); cell.fill = fill(C_MID)
        cell.alignment = ctr(); cell.border = brd
    ws4.row_dimensions[3].height = 28

    totals_per_band = {b: int(df[f'Total_B{b}'].sum()) for b in band_ids}
    peaks_per_band  = {b: int(df[f'Total_B{b}'].max()) for b in band_ids}
    avg_daily_band  = {b: round(totals_per_band[b] / n_days) for b in band_ids}
    avg_hourly_band = {b: round(df[f'Total_B{b}'].mean()) for b in band_ids}

    total_sens1 = sum(totals_per_band[b] for b in sens1_bands)
    total_sens2 = sum(totals_per_band[b] for b in sens2_bands)
    total_all   = total_sens1 + total_sens2

    def sens_diff_ratio(metric_dict):
        s1 = sum(metric_dict[b] for b in sens1_bands)
        s2 = sum(metric_dict[b] for b in sens2_bands)
        return s1 - s2, (s1 / s2 if s2 else 0)

    rows4_data = [
        ("Total vehicule (toată perioada)", totals_per_band),
        ("Medie zilnică (veh/zi)", avg_daily_band),
        ("Medie orară (veh/h)", avg_hourly_band),
        ("Vârf maxim înregistrat (veh/h)", peaks_per_band),
        ("Procent din trafic total", None),
    ]

    dr4 = 4
    for i, (label, data_dict) in enumerate(rows4_data):
        rf = fill(C_LIGHT) if i % 2 == 0 else fill(C_WHITE)
        if label == "Procent din trafic total":
            row_vals = [label]
            for b in band_ids:
                row_vals.append(f"={totals_per_band[b]}/({total_all})")
            row_vals += ["—", "—"]
            fmts = [None] + ['0.0%'] * n_bands + [None, None]
        else:
            diff, ratio = sens_diff_ratio(data_dict)
            row_vals = [label] + [data_dict[b] for b in band_ids] + [diff, f"{ratio:.3f}x"]
            fmts = [None] + ['#,##0'] * n_bands + ['#,##0', None]

        for c, (val, fmt) in enumerate(zip(row_vals, fmts), 1):
            cell = ws4.cell(dr4, c, val)
            cell.font = dfont(bold=(c==1)); cell.fill = rf
            cell.border = brd; cell.alignment = ctr() if c==1 else rgt()
            if fmt: cell.number_format = fmt
            diff_col = 1 + n_bands + 1
            if c == diff_col and isinstance(val, (int, float)):
                cell.font = dfont(bold=True, color=("2E75B6" if val > 0 else "C00000"))
        dr4 += 1

    ws4.column_dimensions['A'].width = 30
    for bi in range(n_bands):
        ws4.column_dimensions[get_column_letter(2 + bi)].width = 15
    ws4.column_dimensions[get_column_letter(2 + n_bands)].width = 20
    ws4.column_dimensions[get_column_letter(3 + n_bands)].width = 16

    # Pie chart — distribuție totală pe benzi
    pie_data_start_col = n_cols_ws4 + 2
    for bi, b in enumerate(band_ids):
        ws4.cell(10 + bi, pie_data_start_col,     f"Banda {b}")
        ws4.cell(10 + bi, pie_data_start_col + 1, totals_per_band[b])

    pie_ws4 = PieChart()
    pct_list = [round(totals_per_band[b] / total_all * 100) for b in band_ids] if total_all else [0]*n_bands
    pie_ws4.title = "Distribuție totală: " + "  |  ".join([f"B{b}({pct_list[bi]}%)" for bi, b in enumerate(band_ids)])
    pie_ws4.style = 10; pie_ws4.width = 14; pie_ws4.height = 12
    pie_data_ref = Reference(ws4, min_col=pie_data_start_col + 1, min_row=10, max_row=10 + n_bands - 1)
    pie_cats_ref = Reference(ws4, min_col=pie_data_start_col,     min_row=10, max_row=10 + n_bands - 1)
    pie_ws4.add_data(pie_data_ref); pie_ws4.set_categories(pie_cats_ref)
    for bi in range(n_bands):
        pt = DataPoint(idx=bi)
        pt.graphicalProperties.solidFill = BAND_COLORS[bi % len(BAND_COLORS)]
        pie_ws4.series[0].dPt.append(pt)
    lbls_ws4 = DataLabelList(); lbls_ws4.showPercent = True; lbls_ws4.showCatName = True
    lbls_ws4.showVal = False; pie_ws4.dLbls = lbls_ws4
    ws4.add_chart(pie_ws4, "J2")

    # Chart trafic zilnic — pe sensuri
    chart_data_start = dr4 + 2
    ws4.cell(chart_data_start, 1, "Data").font = dfont(bold=True)
    ws4.cell(chart_data_start, 2, f"Sens 1 ({sens1_lbl})").font = dfont(bold=True)
    ws4.cell(chart_data_start, 3, f"Sens 2 ({sens2_lbl})").font = dfont(bold=True)
    for ri, row in daily_data.iterrows():
        dr_ch = chart_data_start + 1 + ri
        ws4.cell(dr_ch, 1, row['Zi'].strftime("%d.%m"))
        ws4.cell(dr_ch, 2, int(row['Total_Sens1']))
        ws4.cell(dr_ch, 3, int(row['Total_Sens2']))
    last_chart_row4 = chart_data_start + len(daily_data)

    chart6 = BarChart()
    chart6.type = "col"; chart6.grouping = "clustered"
    chart6.title = f"Trafic zilnic — Sens 1 ({sens1_lbl}) vs Sens 2 ({sens2_lbl})"
    chart6.y_axis.title = "Vehicule / zi"; chart6.style = 10
    chart6.width = latime_dinamica; chart6.height = 12
    chart6.legend.position = 'b'; chart6.legend.overlay = False
    configure_x_axis(chart6.x_axis)
    chart6.y_axis.delete = False
    chart6.layout = make_bar_layout()
    d6s1 = Reference(ws4, min_col=2, min_row=chart_data_start, max_row=last_chart_row4)
    d6s2 = Reference(ws4, min_col=3, min_row=chart_data_start, max_row=last_chart_row4)
    c6ct = Reference(ws4, min_col=1, min_row=chart_data_start+1, max_row=last_chart_row4)
    chart6.add_data(d6s1, titles_from_data=True)
    chart6.add_data(d6s2, titles_from_data=True)
    chart6.set_categories(c6ct)
    chart6.series[0].graphicalProperties.solidFill = "2E75B6"
    chart6.series[1].graphicalProperties.solidFill = "ED7D31"
    ws4.add_chart(chart6, "J26")

    # ── FOAIE 6: Reguli Calcul ───────────────────────────────────────────────
    ws_reguli = wb.create_sheet("Reguli Calcul")
    ws_reguli.merge_cells("A1:D1")
    ws_reguli["A1"] = f"REGULI DE CALCUL  |  Contor: {site_id}"
    ws_reguli["A1"].font = Font(name='Arial', size=13, bold=True, color=C_WHITE)
    ws_reguli["A1"].fill = fill(C_DARK); ws_reguli["A1"].alignment = ctr()
    headers_reguli = ["Regulă","Valoare","Unitate","Descriere"]
    for c, h in enumerate(headers_reguli, 1):
        ws_reguli.cell(2, c, h).font = hfont(10)
        ws_reguli.cell(2, c, h).fill = fill(C_MID)
        ws_reguli.cell(2, c, h).alignment = ctr()
        # Înlocuiește întregul bloc `reguli = [...]` și bucla de scriere cu:
        reguli = [
            ("Minim ore / zi", MIN_ORE_ZI, "ore",
             "Zi validă = minim acest număr de ore înregistrate"),
            ("Minim zile / lună", MIN_ZILE_LUNA, "zile",
             "Medie lunară validă = minim acest număr de zile valide"),
            ("Alternativă săptămână", MIN_ZILE_SAPT, "zile",
             "Sau minim 7 zile consecutive valide din lună"),
            ("Minim luni / an (MZA)", _MIN_LUNI_AN, "luni",
             f"MZA = medie a lunilor valide dacă sunt cel puțin {_MIN_LUNI_AN} luni cu date"),
            ("Fallback MZA — luna Mai", _MIN_LUNI_AN_MAI, "luna nr.",
             f"Dacă sunt sub {_MIN_LUNI_AN} luni valide, MZA preia valoarea lunii Mai (luna {_MIN_LUNI_AN_MAI}); dacă lipsește și Mai → celulă goală"),
            ("Date complete", "toate", "zile/lună",
             "Toate zilele lunii au date valide (ore ≥ MIN_ORE_ZI)"),
            ("Date parțiale", f"≥{MIN_ZILE_LUNA}", "zile/lună",
             "Suficiente zile valide dar nu toate — media pe zilele disponibile"),
            ("Date parțiale - 7 zile", "1", "săptămână",
             "Cel puțin o săptămână completă (7 zile consecutive) validă — media pe 7 zile"),
            ("Nu există date (lună)", "0", "zile/lună",
             "Nicio zi validă în lună — luna exclusă complet din calcule MZL și MZA"),
            ("Nu există date (zi)", "0", "ore/zi",
             "Ziua nu are nicio oră cu vehicule înregistrate — indicator 'Nu există date', rând roșu în Rezumat Zilnic, afișat ca 0/24 la Ore înregistrate"),
        ]
        for i, (reg, val, unit, desc) in enumerate(reguli, 3):
            rf = fill(C_LIGHT) if i % 2 == 0 else fill(C_WHITE)
            for c, v in enumerate([reg, val, unit, desc], 1):
                cell = ws_reguli.cell(i, c, v)
                cell.font = dfont(9)
                cell.fill = rf
                cell.border = brd
                cell.alignment = (ctr() if c <= 3
                                  else Alignment(horizontal='left',
                                                 vertical='center',
                                                 wrap_text=True))
            ws_reguli.row_dimensions[i].height = 30
        ws_reguli.column_dimensions['A'].width = 28
        ws_reguli.column_dimensions['B'].width = 14
        ws_reguli.column_dimensions['C'].width = 12
        ws_reguli.column_dimensions['D'].width = 65

    wb.active = wb.sheetnames.index("Rezumat Zilnic")
    wb.save(excel_path)


# ══════════════════════════════════════════════════════════════════════════════
# PROCESARE COMPLETĂ UN FIȘIER
# ══════════════════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════════════════
# SCANARE RAPIDĂ — extrage site_id și nr. înregistrări fără procesare completă
# ══════════════════════════════════════════════════════════════════════════════
