import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# Configuración de la página
st.set_page_config(
    page_title="Informe Trimestral - Mediaciones, Juicios y Sentencias",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Informe Trimestral de Gestión Judicial")
st.markdown("Análisis comparativo de **Mediaciones**, **Juicios** y **Sentencias**.")

# ---------------------------------------------------------
# FUNCIONES AUXILIARES DE FORMATO
# ---------------------------------------------------------
def format_currency(val):
    """Formatea valores numéricos en formato de moneda ($ 1.234.567,89)"""
    if pd.isna(val) or val == "":
        return ""
    try:
        if isinstance(val, str):
            val_clean = val.replace('$', '').replace(' ', '').replace('.', '').replace(',', '.')
            num = float(val_clean)
        else:
            num = float(val)
            
        if num.is_integer():
            return f"$ {int(num):,}".replace(",", ".")
        else:
            return f"$ {num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(val).strip()


# ---------------------------------------------------------
# FUNCIONES DE PROCESAMIENTO DE DATOS
# ---------------------------------------------------------
@st.cache_data
def load_data(file):
    """Parsea el archivo Excel con las 3 secciones principales y las hojas de Stock y Juicios Extra"""
    xls = pd.ExcelFile(file)
    
    # 1. Carga de datos de la Hoja 1
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
            
            df_sec = pd.DataFrame({
                'Fecha': dates,
                m1_name: m1_vals,
                m2_name: m2_vals
            }).dropna(subset=['Fecha']).sort_values('Fecha').reset_index(drop=True)
            
            df_sec['Año'] = df_sec['Fecha'].dt.year
            df_sec['Mes_Num'] = df_sec['Fecha'].dt.month
            df_sec['Periodo'] = df_sec['Fecha'].apply(lambda d: f"{meses_cortos_es[d.month]}-{d.year}")
            df_sec['Es_Enero'] = df_sec['Mes_Num'] == 1
            
            if sec_name == 'SENTENCIAS':
                df_sec['% Rechazo s/ Condena'] = np.where(
                    df_sec[m1_name] > 0,
                    (df_sec[m2_name] / df_sec[m1_name]) * 100,
                    0
                )
            
            sections[sec_name] = {
                'df': df_sec,
                'm1_name': m1_name,
                'm2_name': m2_name
            }
            i += 3
        else:
            i += 1
            
    # 2. Carga de la Hoja 2 (Stock)
    stock_data = None
    if len(xls.sheet_names) > 1:
        try:
            df_stock = pd.read_excel(xls, sheet_name=1, header=None)
            stock_data = parse_stock_sheet(df_stock)
        except Exception:
            stock_data = None

    # 3. Carga de la Hoja 3 (Presupuesto y Embargos)
    presupuesto_data, embargos_data = None, None
    if len(xls.sheet_names) > 2:
        try:
            df_jui_extra = pd.read_excel(xls, sheet_name=2, header=None)
            presupuesto_data, embargos_data = parse_juicios_extra_sheet(df_jui_extra)
        except Exception:
            presupuesto_data, embargos_data = None, None

    return sections, stock_data, presupuesto_data, embargos_data


def parse_stock_sheet(df_sheet):
    """Extrae dinómicamente los datos de Stock de la Hoja 2"""
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
    """Extrae dinómicamente las tablas de Presupuesto y Embargos de la Hoja 3"""
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
            
        if mode == 'PRESUPUESTO' and cell_0 != '':
            mes = str(row[0]).strip()
            monto = row[1]
            if pd.notna(mes) and pd.notna(monto):
                presupuesto_list.append({
                    'MES': mes,
                    'MONTO': format_currency(monto)
                })
                
        elif mode == 'EMBARGOS' and cell_0 != '':
            mes = str(row[0]).strip()
            cant = row[1]
            monto = row[2]
            if pd.notna(mes) and pd.notna(cant):
                cant_str = str(int(cant)) if isinstance(cant, (int, float)) and not pd.isna(cant) else str(cant).strip()
                embargos_list.append({
                    'MES': mes,
                    'CANTIDAD': cant_str,
                    'MONTO': format_currency(monto)
                })
                
    return pd.DataFrame(presupuesto_list), pd.DataFrame(embargos_list)


def render_trimestral_comparison(df, m1, m2):
    """Genera la sección comparativa del último trimestre vs el trimestre anterior"""
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
    
    c1.metric(
        label=f"Total {m1} (Último Trimestre)",
        value=f"{m1_last}",
        delta=f"{diff_m1:+d} ({pct_m1:+d}%) respecto a {p_prev}"
    )
    
    c2.metric(
        label=f"Total {m2} (Último Trimestre)",
        value=f"{m2_last}",
        delta=f"{diff_m2:+d} ({pct_m2:+d}%) respecto a {p_prev}"
    )


# ---------------------------------------------------------
# CARGA DE ARCHIVO EN LA BARRA LATERAL
# ---------------------------------------------------------
st.sidebar.header("📁 Cargar Datos")
uploaded_file = st.sidebar.file_uploader("Seleccioná el archivo Excel (`.xlsx`)", type=["xlsx"])

if uploaded_file is not None:
    data, stock_info, df_presupuesto, df_embargos = load_data(uploaded_file)
else:
    try:
        data, stock_info, df_presupuesto, df_embargos = load_data('TRIMESTRAL PRUEBA.xlsx')
        st.sidebar.success("Usando archivo local de prueba.")
    except Exception:
        st.info("👋 Por favor cargá un archivo Excel en la barra lateral para continuar.")
        st.stop()

# Pestañas de la aplicación
tab_med, tab_jui, tab_sen = st.tabs(["🤝 Mediaciones", "⚖️ Juicios", "📜 Sentencias"])

# ---------------------------------------------------------
# 1. MEDIACIONES
# ---------------------------------------------------------
with tab_med:
    st.header("🤝 Análisis de Mediaciones")
    sec = data['MEDIACIONES']
    df = sec['df']
    m1, m2 = sec['m1_name'], sec['m2_name']
    
    df_no_jan = df[~df['Es_Enero']]
    avg_m1 = int(df_no_jan[m1].mean())
    avg_m2 = int(df_no_jan[m2].mean())
    
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
        fig_avg = go.Figure(data=[
            go.Bar(name='Promedio', x=[m1, m2], y=[avg_m1, avg_m2], marker_color=['#2b5c8f', '#d9534f'], text=[f"{avg_m1}", f"{avg_m2}"], textposition='auto')
        ])
        fig_avg.update_layout(yaxis_title="Promedio Mensual", showlegend=False)
        st.plotly_chart(fig_avg, use_container_width=True)

    st.subheader("📋 Tabla de Datos Mensuales")
    st.dataframe(
        df[['Periodo', m1, m2]],
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("---")
    st.subheader("📅 Totales Anuales (Ingresos y Bajas)")
    df_annual = df.groupby('Año')[[m1, m2]].sum().reset_index()
    st.dataframe(df_annual, use_container_width=True, hide_index=True)

    st.markdown("---")
    render_trimestral_comparison(df, m1, m2)

    st.markdown("---")
    st.subheader("📦 Estado e Información de Stock")
    
    if stock_info and stock_info['stock'] is not None:
        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric("Stock Actual", f"{stock_info['stock']:,}".replace(",", "."))
        
        auto_label = f"{stock_info['auto_val']:,}".replace(",", ".") if stock_info['auto_val'] else "-"
        sc2.metric("🚗 Auto", auto_label, delta=stock_info['auto_pct'], delta_color="off")
        
        moto_label = f"{stock_info['moto_val']:,}".replace(",", ".") if stock_info['moto_val'] else "-"
        sc3.metric("🏍️ Moto", moto_label, delta=stock_info['moto_pct'], delta_color="off")
        
        if stock_info['trim_ant'] is not None:
            trim_ant_val = stock_info['trim_ant']
            diff_stock = stock_info['stock'] - trim_ant_val
            pct_stock = int(round((diff_stock / trim_ant_val) * 100)) if trim_ant_val > 0 else 0
            sc4.metric(
                "Stock Trimestre Anterior",
                f"{trim_ant_val:,}".replace(",", "."),
                delta=f"{diff_stock:+d} ({pct_stock:+d}%) vs Actual"
            )
        else:
            sc4.metric("Stock Trimestre Anterior", "-")
    else:
        st.info("ℹ️ Para visualizar el stock, asegurate de incluir la Hoja 2 en tu archivo Excel.")

# ---------------------------------------------------------
# 2. JUICIOS
# ---------------------------------------------------------
with tab_jui:
    st.header("⚖️ Análisis de Juicios")
    sec = data['JUICIOS']
    df = sec['df']
    m1, m2 = sec['m1_name'], sec['m2_name']
    
    df_no_jan = df[~df['Es_Enero']]
    avg_m1 = int(df_no_jan[m1].mean())
    avg_m2 = int(df_no_jan[m2].mean())
    
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
        fig_avg = go.Figure(data=[
            go.Bar(name='Promedio', x=[m1, m2], y=[avg_m1, avg_m2], marker_color=['#17a2b8', '#ffc107'], text=[f"{avg_m1}", f"{avg_m2}"], textposition='auto')
        ])
        fig_avg.update_layout(yaxis_title="Promedio Mensual", showlegend=False)
        st.plotly_chart(fig_avg, use_container_width=True)

    st.subheader("📋 Tabla de Datos Mensuales")
    st.dataframe(
        df[['Periodo', m1, m2]],
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("---")
    st.subheader("📅 Totales Anuales (Ingresos y Bajas)")
    df_annual = df.groupby('Año')[[m1, m2]].sum().reset_index()
    st.dataframe(df_annual, use_container_width=True, hide_index=True)

    st.markdown("---")
    render_trimestral_comparison(df, m1, m2)

    # SECCIÓN: Presupuesto y Embargos (Hoja 3)
    st.markdown("---")
    st.subheader("💰 Presupuesto y Embargos")
    
    j_col1, j_col2 = st.columns(2)
    
    with j_col1:
        st.markdown("#### 📌 Presupuesto")
        if df_presupuesto is not None and not df_presupuesto.empty:
            st.dataframe(df_presupuesto, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ No se encontraron datos de Presupuesto en la Hoja 3.")
            
    with j_col2:
        st.markdown("#### 🔒 Embargos")
        if df_embargos is not None and not df_embargos.empty:
            st.dataframe(df_embargos, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ No se encontraron datos de Embargos en la Hoja 3.")

# ---------------------------------------------------------
# 3. SENTENCIAS
# ---------------------------------------------------------
with tab_sen:
    st.header("📜 Análisis de Sentencias / Demandas")
    sec = data['SENTENCIAS']
    df = sec['df']
    m1, m2 = sec['m1_name'], sec['m2_name']
    
    df_no_jan = df[~df['Es_Enero']]
    avg_m1 = int(df_no_jan[m1].mean())
    avg_m2 = int(df_no_jan[m2].mean())
    
    total_condena = df_no_jan[m1].sum()
    total_rechazo = df_no_jan[m2].sum()
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
        fig_avg = go.Figure(data=[
            go.Bar(name='Promedio', x=[m1, m2], y=[avg_m1, avg_m2], marker_color=['#28a745', '#dc3545'], text=[f"{avg_m1}", f"{avg_m2}"], textposition='auto')
        ])
        fig_avg.update_layout(yaxis_title="Promedio Mensual", showlegend=False)
        st.plotly_chart(fig_avg, use_container_width=True)

    st.subheader("📈 Porcentaje Mensual de Rechazo respecto a Condena (%)")
    fig_pct = px.bar(
        df,
        x='Periodo',
        y='% Rechazo s/ Condena',
        color='Es_Enero',
        color_discrete_map={True: '#999999', False: '#6f42c1'},
        labels={'% Rechazo s/ Condena': '% Rechazo', 'Es_Enero': 'Mes de Enero'},
        text_auto='.0f'
    )
    fig_pct.add_hline(y=pct_total_rechazo, line_dash="dot", line_color="#e83e8c", annotation_text=f"% General Total: {pct_total_rechazo}%")
    st.plotly_chart(fig_pct, use_container_width=True)

    st.subheader("📋 Tabla de Datos Mensuales con % de Rechazo")
    df_display = df.copy()
    df_display['% Rechazo s/ Condena'] = df_display['% Rechazo s/ Condena'].round().astype(int).astype(str) + '%'
    
    st.dataframe(
        df_display[['Periodo', m1, m2, '% Rechazo s/ Condena']],
        use_container_width=True,
        hide_index=True
    )
    
    st.markdown("---")
    st.subheader("📅 Totales Anuales (Condena y Rechazada)")
    df_annual = df.groupby('Año')[[m1, m2]].sum().reset_index()
    
    df_annual['% Rechazo s/ Condena'] = np.where(
        df_annual[m1] > 0,
        (df_annual[m2] / df_annual[m1] * 100).round().astype(int).astype(str) + '%',
        '0%'
    )
    
    st.dataframe(df_annual, use_container_width=True, hide_index=True)

    st.markdown("---")
    render_trimestral_comparison(df, m1, m2)
