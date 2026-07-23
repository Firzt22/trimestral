import datetime
import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Configuración de la página Streamlit
st.set_page_config(
    page_title="Informe Trimestral - Mediaciones, Juicios y Sentencias",
    page_icon="📊",
    layout="wide"
)

# ---------------------------------------------------------
# CAMPO DE TEXTO PARA EL TÍTULO DEL INFORME
# ---------------------------------------------------------
custom_title = st.text_input(
    "✏️ Título del Informe (Editable)", 
    value="Informe Trimestral de Gestión Judicial",
    help="Ingresá el título que querés que aparezca en la parte superior de la página y del PDF descargable."
)

st.title(f"📊 {custom_title}")
st.markdown("Análisis comparativo de **Mediaciones**, **Juicios**, **Sentencias** y **Evolución del Stock**.")

# ---------------------------------------------------------
# FUNCIONES AUXILIARES DE FORMATO
# ---------------------------------------------------------
def format_currency_exact(val):
    if pd.isna(val) or val == "" or val is None:
        return ""
    val_str = str(val).strip()
    if '$' in val_str:
        return val_str
    try:
        num = float(val_str.replace('.', '').replace(',', '.')) if ',' in val_str and '.' in val_str else float(val_str)
        num = round(num, 2)
        if num.is_integer():
            formatted = f"{int(num):,}".replace(",", ".")
        else:
            formatted = f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"$ {formatted}"
    except (ValueError, TypeError):
        return val_str


def format_trimestre_label(val):
    if pd.isna(val) or val is None:
        return None
    val_str = str(val).strip()
    if val_str == '' or val_str.upper() in ['NAN', 'NONE']:
        return None

    if not ('00:00:00' in val_str or 'T00:' in val_str) and '-' in val_str and len(val_str) <= 8:
        parts = val_str.split('-')
        if len(parts) == 2 and not parts[0].isdigit():
            return val_str

    try:
        dt = pd.to_datetime(val_str)
        meses_map = {
            1: 'ene', 2: 'feb', 3: 'mar', 4: 'abr', 5: 'may', 6: 'jun',
            7: 'jul', 8: 'ago', 9: 'sept', 10: 'oct', 11: 'nov', 12: 'dic'
        }
        mes_str = meses_map.get(dt.month, str(dt.month))
        year_str = str(dt.year)[-2:]
        return f"{mes_str}-{year_str}"
    except Exception:
        return val_str

# ---------------------------------------------------------
# PARSERS DE DATOS EXCEL
# ---------------------------------------------------------
def parse_stock_sheet(df_sheet):
    stock_val, trim_ant_val = None, None
    auto_val, auto_pct = None, None
    moto_val, moto_pct = None, None
    
    for r in range(len(df_sheet)):
        row_str = [str(cell).strip().upper() for cell in df_sheet.iloc[r].tolist()]
        if 'STOCK' in row_str:
            col_idx = row_str.index('STOCK')
            if r + 1 < len(df_sheet):
                val = df_sheet.iloc[r+1, col_idx]
                if pd.notna(val): stock_val = int(pd.to_numeric(val, errors='coerce'))
                
        if 'AUTO' in row_str and 'MOTO' in row_str:
            col_auto = row_str.index('AUTO')
            col_moto = row_str.index('MOTO')
            if r + 1 < len(df_sheet):
                v_a = df_sheet.iloc[r+1, col_auto]
                v_m = df_sheet.iloc[r+1, col_moto]
                if pd.notna(v_a): auto_val = int(pd.to_numeric(v_a, errors='coerce'))
                if pd.notna(v_m): moto_val = int(pd.to_numeric(v_m, errors='coerce'))
            if r + 2 < len(df_sheet):
                p_a = df_sheet.iloc[r+2, col_auto]
                p_m = df_sheet.iloc[r+2, col_moto]
                if pd.notna(p_a):
                    num_a = float(p_a) * 100 if float(p_a) <= 1 else float(p_a)
                    auto_pct = f"{int(round(num_a))}%"
                if pd.notna(p_m):
                    num_m = float(p_m) * 100 if float(p_m) <= 1 else float(p_m)
                    moto_pct = f"{int(round(num_m))}%"

        for col_idx, cell in enumerate(row_str):
            if 'TRIM. ANT' in cell or 'TRIM' in cell:
                if r + 1 < len(df_sheet):
                    val = df_sheet.iloc[r+1, col_idx]
                    if pd.notna(val): trim_ant_val = int(pd.to_numeric(val, errors='coerce'))

    return {
        'stock': stock_val,
        'trim_ant': trim_ant_val,
        'auto_val': auto_val,
        'auto_pct': auto_pct,
        'moto_val': moto_val,
        'moto_pct': moto_pct
    }


def parse_juicios_extra_sheet(df_sheet):
    presupuesto_list = []
    embargos_list = []
    mode = None
    for r in range(len(df_sheet)):
        row = df_sheet.iloc[r].tolist()
        cell_0 = str(row[0]).strip().upper() if pd.notna(row[0]) else ''
        
        if cell_0 == 'PRESUPUESTO':
            mode = 'PRESUPUESTO'
            continue
        elif cell_0 == 'EMBARGOS':
            mode = 'EMBARGOS'
            continue
            
        if mode == 'PRESUPUESTO' and cell_0 != '' and cell_0 != 'NAN':
            mes = str(row[0]).strip()
            monto = row[1]
            if pd.notna(monto) and str(monto).strip().upper() != 'NAN':
                presupuesto_list.append({'MES': mes, 'MONTO': format_currency_exact(monto)})
                
        elif mode == 'EMBARGOS' and cell_0 != '' and cell_0 != 'NAN':
            mes = str(row[0]).strip()
            cant = row[1]
            monto = row[2]
            if str(cant).strip().upper() == 'CANTIDAD': continue
            if pd.notna(cant) and str(cant).strip().upper() != 'NAN':
                cant_str = str(cant).strip().split('.')[0] if '.' in str(cant) else str(cant).strip()
                embargos_list.append({'MES': mes, 'CANTIDAD': cant_str, 'MONTO': format_currency_exact(monto)})
                
    return pd.DataFrame(presupuesto_list), pd.DataFrame(embargos_list)


def parse_stock_evolution_sheet(df_sheet):
    header_row = -1
    for r in range(len(df_sheet)):
        row_vals = [str(c).strip().upper() for c in df_sheet.iloc[r].tolist()]
        if 'TRIMESTRE' in row_vals and 'CANTIDAD' in row_vals:
            header_row = r
            break
            
    if header_row == -1: return None
        
    row_headers = [str(c).strip().upper() for c in df_sheet.iloc[header_row].tolist()]
    col_tri = row_headers.index('TRIMESTRE')
    col_cant = row_headers.index('CANTIDAD')
    
    col_pct = None
    for idx, c in enumerate(row_headers):
        if '%' in c or 'AUMENTO' in c or 'VARIACION' in c:
            col_pct = idx
            break

    records = []
    for r in range(header_row + 1, len(df_sheet)):
        tri_val = df_sheet.iloc[r, col_tri]
        cant_val = df_sheet.iloc[r, col_cant]
        pct_val = df_sheet.iloc[r, col_pct] if col_pct is not None else None
        
        tri_formatted = format_trimestre_label(tri_val)
        if tri_formatted is not None:
            try:
                cant_clean = str(cant_val).strip().replace('.', '').replace(',', '')
                cant_num = int(float(cant_clean))
            except (ValueError, TypeError): continue
            
            pct_float = None
            pct_str = "-"
            if pd.notna(pct_val):
                p_s = str(pct_val).strip().replace(',', '.').replace('%', '')
                try:
                    pct_float = float(p_s)
                    if abs(pct_float) <= 1 and pct_float != 0: pct_float = pct_float * 100
                    pct_str = f"{pct_float:.1f}%".replace('.', ',') if pct_float != 0 else "0,0%"
                except ValueError:
                    pct_str = str(pct_val).strip()
                    
            records.append({'TRIMESTRE': tri_formatted, 'CANTIDAD': cant_num, 'PCT_NUM': pct_float, 'PCT_STR': pct_str})
            
    df_res = pd.DataFrame(records)
    if not df_res.empty:
        for i in range(len(df_res)):
            if i == 0:
                if pd.isna(df_res.loc[i, 'PCT_NUM']): df_res.loc[i, 'PCT_NUM'] = 0.0
            else:
                prev = df_res.loc[i-1, 'CANTIDAD']
                curr = df_res.loc[i, 'CANTIDAD']
                if pd.isna(df_res.loc[i, 'PCT_NUM']) and prev > 0:
                    diff_pct = ((curr - prev) / prev) * 100
                    df_res.loc[i, 'PCT_NUM'] = diff_pct
                    df_res.loc[i, 'PCT_STR'] = f"{diff_pct:.1f}%".replace('.', ',')
                    
    return df_res


@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    df_raw = pd.read_excel(xls, sheet_name=0, header=None)
    sections = {}
    
    meses_cortos_es = {
        1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
        7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'
    }
    
    i = 0
    while i < len(df_raw):
        val = str(df_raw.iloc[i, 0]).strip().upper()
        if val in ['MEDIACIONES', 'JUICIOS', 'SENTENCIAS']:
            sec_name = val
            raw_dates = df_raw.iloc[i, 1:].values
            dates = pd.to_datetime(raw_dates, errors='coerce')
            
            m1_name = str(df_raw.iloc[i+1, 0]).strip()
            m1_vals = pd.to_numeric(df_raw.iloc[i+1, 1:], errors='coerce').fillna(0)
            
            m2_name = str(df_raw.iloc[i+2, 0]).strip()
            m2_vals = pd.to_numeric(df_raw.iloc[i+2, 1:], errors='coerce').fillna(0)
            
            df_sec = pd.DataFrame({'Fecha': dates, m1_name: m1_vals, m2_name: m2_vals}).dropna(subset=['Fecha']).sort_values('Fecha').reset_index(drop=True)
            df_sec['Año'] = df_sec['Fecha'].dt.year
            df_sec['Mes_Num'] = df_sec['Fecha'].dt.month
            df_sec['Periodo'] = df_sec['Fecha'].apply(lambda d: f"{meses_cortos_es[d.month]}-{d.year}")
            df_sec['Es_Enero'] = df_sec['Mes_Num'] == 1
            
            if sec_name == 'SENTENCIAS':
                df_sec['% Rechazo s/ Condena'] = np.where(df_sec[m1_name] > 0, (df_sec[m2_name] / df_sec[m1_name]) * 100, 0)
            
            sections[sec_name] = {'df': df_sec, 'm1_name': m1_name, 'm2_name': m2_name}
            i += 3
        else:
            i += 1
            
    stock_data, presupuesto_data, embargos_data, stock_evo_df = None, None, None, None
    if len(xls.sheet_names) > 1:
        try: stock_data = parse_stock_sheet(pd.read_excel(xls, sheet_name=1, header=None))
        except Exception: pass

    if len(xls.sheet_names) > 2:
        try: presupuesto_data, embargos_data = parse_juicios_extra_sheet(pd.read_excel(xls, sheet_name=2, header=None, dtype=str))
        except Exception: pass

    if len(xls.sheet_names) > 3:
        try: stock_evo_df = parse_stock_evolution_sheet(pd.read_excel(xls, sheet_name=3, header=None, dtype=str))
        except Exception: pass

    return sections, stock_data, presupuesto_data, embargos_data, stock_evo_df

# ---------------------------------------------------------
# GENERACIÓN DE PDF CON REPORTLAB
# ---------------------------------------------------------
def generate_pdf_report(title, sections, stock_info, df_presupuesto, df_embargos, df_stock_evo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    
    # Estilos personalizados
    style_title = ParagraphStyle('ReportTitle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor('#1A365D'), spaceAfter=10)
    style_subtitle = ParagraphStyle('ReportSubtitle', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#4A5568'), spaceAfter=15)
    style_h1 = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontSize=12, leading=15, textColor=colors.HexColor('#2B6CB0'), spaceBefore=12, spaceAfter=6)
    style_body = ParagraphStyle('ReportBody', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#2D3748'))
    style_th = ParagraphStyle('TableHead', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.white, fontName='Helvetica-Bold')
    style_td = ParagraphStyle('TableData', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#2D3748'))

    story = []
    
    # Encabezado principal
    story.append(Paragraph(f"<b>{title.upper()}</b>", style_title))
    today_str = datetime.datetime.now().strftime("%d/%m/%Y")
    story.append(Paragraph(f"Informe Resumen Ejecutivo • Fecha de generación: {today_str}", style_subtitle))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#2B6CB0'), spaceAfter=12))

    def build_trim_data(sec_key):
        if sec_key not in sections: return None
        df = sections[sec_key]['df']
        m1 = sections[sec_key]['m1_name']
        m2 = sections[sec_key]['m2_name']
        if len(df) < 6: return None
        last_3 = df.iloc[-3:]
        prev_3 = df.iloc[-6:-3]
        p_last = f"{last_3.iloc[0]['Periodo']} a {last_3.iloc[-1]['Periodo']}"
        p_prev = f"{prev_3.iloc[0]['Periodo']} a {prev_3.iloc[-1]['Periodo']}"
        return {
            'p_last': p_last, 'p_prev': p_prev,
            'm1_name': m1, 'm1_last': int(last_3[m1].sum()), 'm1_prev': int(prev_3[m1].sum()),
            'm2_name': m2, 'm2_last': int(last_3[m2].sum()), 'm2_prev': int(prev_3[m2].sum())
        }

    # 1. MEDIACIONES
    story.append(Paragraph("1. MEDIACIONES", style_h1))
    med_data = build_trim_data('MEDIACIONES')
    if med_data:
        d_m1 = med_data['m1_last'] - med_data['m1_prev']
        d_m2 = med_data['m2_last'] - med_data['m2_prev']
        pct_m1 = f"({(d_m1/med_data['m1_prev']*100):+.0f}%)" if med_data['m1_prev'] > 0 else ""
        pct_m2 = f"({(d_m2/med_data['m2_prev']*100):+.0f}%)" if med_data['m2_prev'] > 0 else ""

        table_med = [
            [Paragraph("Concepto", style_th), Paragraph(f"Trim. Actual ({med_data['p_last']})", style_th), Paragraph(f"Trim. Anterior ({med_data['p_prev']})", style_th), Paragraph("Variación", style_th)],
            [Paragraph(med_data['m1_name'], style_td), Paragraph(str(med_data['m1_last']), style_td), Paragraph(str(med_data['m1_prev']), style_td), Paragraph(f"{d_m1:+d} {pct_m1}", style_td)],
            [Paragraph(med_data['m2_name'], style_td), Paragraph(str(med_data['m2_last']), style_td), Paragraph(str(med_data['m2_prev']), style_td), Paragraph(f"{d_m2:+d} {pct_m2}", style_td)]
        ]
        t = Table(table_med, colWidths=[150, 130, 130, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t)
    
    # Stock de Mediaciones
    if stock_info and stock_info['stock'] is not None:
        story.append(Spacer(1, 4))
        auto_str = f"{stock_info['auto_val']:,}".replace(",", ".") if stock_info['auto_val'] else "-"
        moto_str = f"{stock_info['moto_val']:,}".replace(",", ".") if stock_info['moto_val'] else "-"
        trim_ant_str = f"{stock_info['trim_ant']:,}".replace(",", ".") if stock_info['trim_ant'] else "-"
        pct_auto = f" ({stock_info['auto_pct']})" if stock_info['auto_pct'] else ""
        pct_moto = f" ({stock_info['moto_pct']})" if stock_info['moto_pct'] else ""

        stk_txt = f"<b>Stock Actual de Mediaciones:</b> {stock_info['stock']:,}".replace(",", ".") + \
                  f" (Auto: {auto_str}{pct_auto} | Moto: {moto_str}{pct_moto})<br/>" + \
                  f"<b>Stock Trimestre Anterior:</b> {trim_ant_str}"
        story.append(Paragraph(stk_txt, style_body))

    story.append(Spacer(1, 10))

    # 2. JUICIOS
    story.append(Paragraph("2. JUICIOS", style_h1))
    jui_data = build_trim_data('JUICIOS')
    if jui_data:
        d_m1 = jui_data['m1_last'] - jui_data['m1_prev']
        d_m2 = jui_data['m2_last'] - jui_data['m2_prev']
        pct_m1 = f"({(d_m1/jui_data['m1_prev']*100):+.0f}%)" if jui_data['m1_prev'] > 0 else ""
        pct_m2 = f"({(d_m2/jui_data['m2_prev']*100):+.0f}%)" if jui_data['m2_prev'] > 0 else ""

        table_jui = [
            [Paragraph("Concepto", style_th), Paragraph(f"Trim. Actual ({jui_data['p_last']})", style_th), Paragraph(f"Trim. Anterior ({jui_data['p_prev']})", style_th), Paragraph("Variación", style_th)],
            [Paragraph(jui_data['m1_name'], style_td), Paragraph(str(jui_data['m1_last']), style_td), Paragraph(str(jui_data['m1_prev']), style_td), Paragraph(f"{d_m1:+d} {pct_m1}", style_td)],
            [Paragraph(jui_data['m2_name'], style_td), Paragraph(str(jui_data['m2_last']), style_td), Paragraph(str(jui_data['m2_prev']), style_td), Paragraph(f"{d_m2:+d} {pct_m2}", style_td)]
        ]
        t = Table(table_jui, colWidths=[150, 130, 130, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t)

    story.append(Spacer(1, 10))

    # 3. SENTENCIAS
    story.append(Paragraph("3. SENTENCIAS", style_h1))
    sen_data = build_trim_data('SENTENCIAS')
    if sen_data:
        d_m1 = sen_data['m1_last'] - sen_data['m1_prev']
        d_m2 = sen_data['m2_last'] - sen_data['m2_prev']
        pct_m1 = f"({(d_m1/sen_data['m1_prev']*100):+.0f}%)" if sen_data['m1_prev'] > 0 else ""
        pct_m2 = f"({(d_m2/sen_data['m2_prev']*100):+.0f}%)" if sen_data['m2_prev'] > 0 else ""

        table_sen = [
            [Paragraph("Concepto", style_th), Paragraph(f"Trim. Actual ({sen_data['p_last']})", style_th), Paragraph(f"Trim. Anterior ({sen_data['p_prev']})", style_th), Paragraph("Variación", style_th)],
            [Paragraph(sen_data['m1_name'], style_td), Paragraph(str(sen_data['m1_last']), style_td), Paragraph(str(sen_data['m1_prev']), style_td), Paragraph(f"{d_m1:+d} {pct_m1}", style_td)],
            [Paragraph(sen_data['m2_name'], style_td), Paragraph(str(sen_data['m2_last']), style_td), Paragraph(str(sen_data['m2_prev']), style_td), Paragraph(f"{d_m2:+d} {pct_m2}", style_td)]
        ]
        t = Table(table_sen, colWidths=[150, 130, 130, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t)

    story.append(Spacer(1, 10))

    # 4. EVOLUCIÓN STOCK DE JUICIOS
    story.append(Paragraph("4. STOCK DE JUICIOS (EVOLUCIÓN TRIMESTRAL)", style_h1))
    if df_stock_evo is not None and not df_stock_evo.empty:
        last_row = df_stock_evo.iloc[-1]
        prev_row = df_stock_evo.iloc[-2] if len(df_stock_evo) > 1 else last_row
        diff_abs = last_row['CANTIDAD'] - prev_row['CANTIDAD']
        
        stk_txt = f"<b>Stock Trimestre Actual ({last_row['TRIMESTRE']}):</b> {last_row['CANTIDAD']:,}".replace(",", ".") + \
                  f" | <b>Variación:</b> {last_row['PCT_STR']} ({diff_abs:+d} casos vs {prev_row['TRIMESTRE']})"
        story.append(Paragraph(stk_txt, style_body))

    story.append(Spacer(1, 10))

    # 5. PRESUPUESTO Y EMBARGOS (ÚLTIMOS 3 MESES)
    story.append(Paragraph("5. PRESUPUESTO Y EMBARGOS (TRIMESTRAL)", style_h1))
    
    t_p_data = [[Paragraph("Mes", style_th), Paragraph("Presupuesto", style_th), Paragraph("Cant. Embargos", style_th), Paragraph("Monto Embargos", style_th)]]
    
    if df_presupuesto is not None and not df_presupuesto.empty:
        pres_rows = df_presupuesto.tail(3).to_dict('records')
        emb_rows = df_embargos.tail(3).to_dict('records') if (df_embargos is not None and not df_embargos.empty) else []
        
        for idx in range(len(pres_rows)):
            p = pres_rows[idx]
            e = emb_rows[idx] if idx < len(emb_rows) else {'CANTIDAD': '-', 'MONTO': '-'}
            t_p_data.append([
                Paragraph(str(p.get('MES', '')), style_td),
                Paragraph(str(p.get('MONTO', '')), style_td),
                Paragraph(str(e.get('CANTIDAD', '')), style_td),
                Paragraph(str(e.get('MONTO', '')), style_td)
            ])
            
        t_pe = Table(t_p_data, colWidths=[120, 130, 120, 140])
        t_pe.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B6CB0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_pe)

    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# CARGA DE DATOS Y BOTÓN PDF EN SIDEBAR
# ---------------------------------------------------------
st.sidebar.header("📁 Cargar Datos")
uploaded_file = st.sidebar.file_uploader("Seleccioná el archivo Excel (`.xlsx`)", type=["xlsx"])

if uploaded_file is not None:
    data, stock_info, df_presupuesto, df_embargos, df_stock_evo = load_data(uploaded_file)
else:
    try:
        data, stock_info, df_presupuesto, df_embargos, df_stock_evo = load_data('TRIMESTRAL PRUEBA.xlsx')
        st.sidebar.success("Usando archivo local de prueba.")
    except Exception:
        st.info("👋 Por favor cargá un archivo Excel en la barra lateral para continuar.")
        st.stop()

# Botón PDF en la barra lateral
st.sidebar.markdown("---")
st.sidebar.header("📄 Exportar Informe")
try:
    pdf_bytes = generate_pdf_report(custom_title, data, stock_info, df_presupuesto, df_embargos, df_stock_evo)
    st.sidebar.download_button(
        label="📥 Descargar Reporte PDF",
        data=pdf_bytes,
        file_name=f"Informe_Trimestral_{datetime.date.today().strftime('%Y_%m_%d')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )
except Exception as e:
    st.sidebar.error(f"Error generando PDF: {e}")

# ---------------------------------------------------------
# RENDER DE PESTAÑAS Y CONTENIDO INTERACTIVO DE STREAMLIT
# ---------------------------------------------------------
def render_trimestral_comparison(df, m1, m2):
    if len(df) < 6:
        st.warning("Se necesitan al menos 6 meses de datos para realizar la comparación trimestral.")
        return
    
    last_3 = df.iloc[-3:]
    prev_3 = df.iloc[-6:-3]
    
    p_last = f"{last_3.iloc[0]['Periodo']} a {last_3.iloc[-1]['Periodo']}"
    p_prev = f"{prev_3.iloc[0]['Periodo']} a {prev_3.iloc[-1]['Periodo']}"
    
    m1_last, m1_prev = int(last_3[m1].sum()), int(prev_3[m1].sum())
    m2_last, m2_prev = int(last_3[m2].sum()), int(prev_3[m2].sum())
    
    diff_m1 = m1_last - m1_prev
    diff_m2 = m2_last - m2_prev
    
    pct_m1 = int(round((diff_m1 / m1_prev * 100))) if m1_prev > 0 else 0
    pct_m2 = int(round((diff_m2 / m2_prev * 100))) if m2_prev > 0 else 0
    
    st.subheader("📉 Comparativo Trimestral Reciente")
    st.caption(f"Comparando **Último Trimestre** ({p_last}) vs **Trimestre Anterior** ({p_prev})")
    
    c1, c2 = st.columns(2)
    c1.metric(f"Total {m1} (Último Trimestre)", f"{m1_last}", delta=f"{diff_m1:+d} ({pct_m1:+d}%) respecto a {p_prev}")
    c2.metric(f"Total {m2} (Último Trimestre)", f"{m2_last}", delta=f"{diff_m2:+d} ({pct_m2:+d}%) respecto a {p_prev}")

tab_med, tab_jui, tab_sen, tab_stock_evo = st.tabs([
    "🤝 Mediaciones", "⚖️ Juicios", "📜 Sentencias", "📈 Evolución Stock Juicios"
])

# 1. MEDIACIONES
with tab_med:
    st.header("🤝 Análisis de Mediaciones")
    sec = data['MEDIACIONES']
    df = sec['df']
    m1, m2 = sec['m1_name'], sec['m2_name']
    
    df_no_jan = df[~df['Es_Enero']]
    avg_m1, avg_m2 = int(df_no_jan[m1].mean()), int(df_no_jan[m2].mean())
    
    col1, col2, col3 = st.columns(3)
    col1.metric(f"Promedio {m1}", f"{avg_m1}")
    col2.metric(f"Promedio {m2}", f"{avg_m2}")
    col3.metric("Meses Analizados", len(df_no_jan))
    
    st.markdown("---")
    g_col1, g_col2 = st.columns([2, 1])
    
    with g_col1:
        st.subheader("Evolución Mensual")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df['Periodo'], y=df[m1], name=m1, marker_color='#2b5c8f'))
        fig.add_trace(go.Bar(x=df['Periodo'], y=df[m2], name=m2, marker_color='#d9534f'))
        fig.add_hline(y=avg_m1, line_dash="dash", line_color="#2b5c8f", annotation_text=f"Prom. {m1}: {avg_m1}")
        fig.add_hline(y=avg_m2, line_dash="dash", line_color="#d9534f", annotation_text=f"Prom. {m2}: {avg_m2}")
        fig.update_layout(barmode='group', xaxis_title="Período", yaxis_title="Cantidad", legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, use_container_width=True)
        
    with g_col2:
        st.subheader("Promedios Generales")
        fig_avg = go.Figure(data=[go.Bar(name='Promedio', x=[m1, m2], y=[avg_m1, avg_m2], marker_color=['#2b5c8f', '#d9534f'], text=[f"{avg_m1}", f"{avg_m2}"], textposition='auto')])
        fig_avg.update_layout(yaxis_title="Promedio Mensual", showlegend=False)
        st.plotly_chart(fig_avg, use_container_width=True)

    st.subheader("📋 Tabla de Datos Mensuales")
    st.dataframe(df[['Periodo', m1, m2]], use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("📅 Totales Anuales (Ingresos y Bajas)")
    st.dataframe(df.groupby('Año')[[m1, m2]].sum().reset_index(), use_container_width=True, hide_index=True)

    st.markdown("---")
    render_trimestral_comparison(df, m1, m2)

    st.markdown("---")
    st.subheader("📦 Estado e Información de Stock")
    if stock_info and stock_info['stock'] is not None:
        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric("Stock Actual", f"{stock_info['stock']:,}".replace(",", "."))
        sc2.metric("🚗 Auto", f"{stock_info['auto_val']:,}".replace(",", ".") if stock_info['auto_val'] else "-", delta=stock_info['auto_pct'], delta_color="off")
        sc3.metric("🏍️ Moto", f"{stock_info['moto_val']:,}".replace(",", ".") if stock_info['moto_val'] else "-", delta=stock_info['moto_pct'], delta_color="off")
        if stock_info['trim_ant'] is not None:
            trim_ant_val = stock_info['trim_ant']
            diff_stock = stock_info['stock'] - trim_ant_val
            pct_stock = int(round((diff_stock / trim_ant_val) * 100)) if trim_ant_val > 0 else 0
            sc4.metric("Stock Trimestre Anterior", f"{trim_ant_val:,}".replace(",", "."), delta=f"{diff_stock:+d} ({pct_stock:+d}%) vs Actual")
        else: sc4.metric("Stock Trimestre Anterior", "-")

# 2. JUICIOS
with tab_jui:
    st.header("⚖️ Análisis de Juicios")
    sec = data['JUICIOS']
    df = sec['df']
    m1, m2 = sec['m1_name'], sec['m2_name']
    df_no_jan = df[~df['Es_Enero']]
    avg_m1, avg_m2 = int(df_no_jan[m1].mean()), int(df_no_jan[m2].mean())
    
    col1, col2, col3 = st.columns(3)
    col1.metric(f"Promedio {m1}", f"{avg_m1}")
    col2.metric(f"Promedio {m2}", f"{avg_m2}")
    col3.metric("Meses Analizados", len(df_no_jan))
    
    st.markdown("---")
    g_col1, g_col2 = st.columns([2, 1])
    
    with g_col1:
        st.subheader("Evolución Mensual")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df['Periodo'], y=df[m1], name=m1, marker_color='#17a2b8'))
        fig.add_trace(go.Bar(x=df['Periodo'], y=df[m2], name=m2, marker_color='#ffc107'))
        fig.add_hline(y=avg_m1, line_dash="dash", line_color="#17a2b8", annotation_text=f"Prom. {m1}: {avg_m1}")
        fig.add_hline(y=avg_m2, line_dash="dash", line_color="#ffc107", annotation_text=f"Prom. {m2}: {avg_m2}")
        fig.update_layout(barmode='group', xaxis_title="Período", yaxis_title="Cantidad", legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, use_container_width=True)
        
    with g_col2:
        st.subheader("Promedios Generales")
        fig_avg = go.Figure(data=[go.Bar(name='Promedio', x=[m1, m2], y=[avg_m1, avg_m2], marker_color=['#17a2b8', '#ffc107'], text=[f"{avg_m1}", f"{avg_m2}"], textposition='auto')])
        fig_avg.update_layout(yaxis_title="Promedio Mensual", showlegend=False)
        st.plotly_chart(fig_avg, use_container_width=True)

    st.subheader("📋 Tabla de Datos Mensuales")
    st.dataframe(df[['Periodo', m1, m2]], use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("📅 Totales Anuales (Ingresos y Bajas)")
    st.dataframe(df.groupby('Año')[[m1, m2]].sum().reset_index(), use_container_width=True, hide_index=True)

    st.markdown("---")
    render_trimestral_comparison(df, m1, m2)

    st.markdown("---")
    st.subheader("💰 Presupuesto y Embargos")
    j_col1, j_col2 = st.columns(2)
    with j_col1:
        st.markdown("#### 📌 Presupuesto")
        if df_presupuesto is not None and not df_presupuesto.empty: st.dataframe(df_presupuesto, use_container_width=True, hide_index=True)
    with j_col2:
        st.markdown("#### 🔒 Embargos")
        if df_embargos is not None and not df_embargos.empty: st.dataframe(df_embargos, use_container_width=True, hide_index=True)

# 3. SENTENCIAS
with tab_sen:
    st.header("📜 Análisis de Sentencias / Demandas")
    sec = data['SENTENCIAS']
    df = sec['df']
    m1, m2 = sec['m1_name'], sec['m2_name']
    df_no_jan = df[~df['Es_Enero']]
    avg_m1, avg_m2 = int(df_no_jan[m1].mean()), int(df_no_jan[m2].mean())
    total_condena, total_rechazo = df_no_jan[m1].sum(), df_no_jan[m2].sum()
    pct_total_rechazo = int(round(total_rechazo / total_condena * 100)) if total_condena > 0 else 0
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric(f"Promedio {m1}", f"{avg_m1}")
    col2.metric(f"Promedio {m2}", f"{avg_m2}")
    col3.metric("% Total Rechazo / Condena", f"{pct_total_rechazo}%")
    col4.metric("Meses Analizados", len(df_no_jan))
    
    st.markdown("---")
    g_col1, g_col2 = st.columns([2, 1])
    
    with g_col1:
        st.subheader("Evolución Mensual")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df['Periodo'], y=df[m1], name=m1, marker_color='#28a745'))
        fig.add_trace(go.Bar(x=df['Periodo'], y=df[m2], name=m2, marker_color='#dc3545'))
        fig.add_hline(y=avg_m1, line_dash="dash", line_color="#28a745", annotation_text=f"Prom. {m1}: {avg_m1}")
        fig.add_hline(y=avg_m2, line_dash="dash", line_color="#dc3545", annotation_text=f"Prom. {m2}: {avg_m2}")
        fig.update_layout(barmode='group', xaxis_title="Período", yaxis_title="Cantidad", legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, use_container_width=True)
        
    with g_col2:
        st.subheader("Promedios Generales")
        fig_avg = go.Figure(data=[go.Bar(name='Promedio', x=[m1, m2], y=[avg_m1, avg_m2], marker_color=['#28a745', '#dc3545'], text=[f"{avg_m1}", f"{avg_m2}"], textposition='auto')])
        fig_avg.update_layout(yaxis_title="Promedio Mensual", showlegend=False)
        st.plotly_chart(fig_avg, use_container_width=True)

    st.subheader("📈 Porcentaje Mensual de Rechazo respecto a Condena (%)")
    fig_pct = px.bar(
        df, x='Periodo', y='% Rechazo s/ Condena', color='Es_Enero',
        color_discrete_map={True: '#999999', False: '#6f42c1'}, text_auto='.0f'
    )
    fig_pct.add_hline(y=pct_total_rechazo, line_dash="dot", line_color="#e83e8c", annotation_text=f"% General Total: {pct_total_rechazo}%")
    st.plotly_chart(fig_pct, use_container_width=True)

    df_display = df.copy()
    df_display['% Rechazo s/ Condena'] = df_display['% Rechazo s/ Condena'].round().astype(int).astype(str) + '%'
    st.dataframe(df_display[['Periodo', m1, m2, '% Rechazo s/ Condena']], use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("📅 Totales Anuales (Condena y Rechazada)")
    df_annual = df.groupby('Año')[[m1, m2]].sum().reset_index()
    df_annual['% Rechazo s/ Condena'] = np.where(df_annual[m1] > 0, (df_annual[m2] / df_annual[m1] * 100).round().astype(int).astype(str) + '%', '0%')
    st.dataframe(df_annual, use_container_width=True, hide_index=True)

    st.markdown("---")
    render_trimestral_comparison(df, m1, m2)

# 4. EVOLUCIÓN HISTÓRICA DE STOCK (HOJA 4)
with tab_stock_evo:
    st.header("📈 Evolución Histórica de Stock de Juicios")
    
    if df_stock_evo is not None and not df_stock_evo.empty:
        s_col1, s_col2 = st.columns([1, 2])
        
        with s_col1:
            st.subheader("📋 Cuadro Trimestral")
            def highlight_negatives(row):
                pct_val = row['PCT_NUM']
                styles = [''] * len(row)
                if pd.notna(pct_val) and pct_val < 0:
                    styles = ['color: #d9534f; font-weight: bold;'] * len(row)
                return styles

            df_show = df_stock_evo.copy()
            df_show['CANTIDAD'] = df_show['CANTIDAD'].apply(lambda x: f"{x:,}".replace(",", "."))
            df_show = df_show.rename(columns={'PCT_STR': '% DE AUMENTO'})
            styled_df = df_show[['TRIMESTRE', 'CANTIDAD', '% DE AUMENTO', 'PCT_NUM']].style.apply(highlight_negatives, axis=1)
            st.dataframe(styled_df, column_config={'PCT_NUM': None}, use_container_width=True, hide_index=True)

        with s_col2:
            st.subheader("📊 Métrica del Último Trimestre")
            last_row = df_stock_evo.iloc[-1]
            prev_row = df_stock_evo.iloc[-2] if len(df_stock_evo) > 1 else last_row
            diff_abs = last_row['CANTIDAD'] - prev_row['CANTIDAD']
            
            m_col1, m_col2 = st.columns(2)
            m_col1.metric("Stock Actual", f"{last_row['CANTIDAD']:,}".replace(",", "."))
            m_col2.metric("Variación Trimestral", f"{last_row['PCT_STR']}", delta=f"{diff_abs:+d} casos vs {prev_row['TRIMESTRE']}")

        st.markdown("---")
        st.subheader("📉 Variación en Cantidad de Stock")
        colors_line = ['#2b5c8f' if p >= 0 else '#d9534f' for p in df_stock_evo['PCT_NUM']]
        fig_cant = go.Figure()
        fig_cant.add_trace(go.Scatter(
            x=df_stock_evo['TRIMESTRE'], y=df_stock_evo['CANTIDAD'], mode='lines+markers+text',
            text=[f"{v:,}".replace(",", ".") for v in df_stock_evo['CANTIDAD']], textposition="top center",
            line=dict(color='#2b5c8f', width=3), marker=dict(size=8, color=colors_line), name="Cantidad Stock"
        ))
        fig_cant.update_layout(xaxis=dict(title="Trimestre", type='category'), yaxis_title="Cantidad de Casos", hovermode="x unified", height=420)
        st.plotly_chart(fig_cant, use_container_width=True)

        st.markdown("---")
        st.subheader("📊 Variación Porcentual del Stock (%)")
        bar_colors = ['#d9534f' if val < 0 else '#28a745' for val in df_stock_evo['PCT_NUM']]
        fig_pct = go.Figure()
        fig_pct.add_trace(go.Bar(
            x=df_stock_evo['TRIMESTRE'], y=df_stock_evo['PCT_NUM'], marker_color=bar_colors,
            text=df_stock_evo['PCT_STR'], textposition='auto', name="% Variación"
        ))
        fig_pct.add_hline(y=0, line_width=1, line_color="#000000")
        fig_pct.update_layout(xaxis=dict(title="Trimestre", type='category'), yaxis_title="Variación %", hovermode="x unified", height=420)
        st.plotly_chart(fig_pct, use_container_width=True)
