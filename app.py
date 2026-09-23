import streamlit as st
import pandas as pd
import datetime
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Analizador de Tiempos - Laboratorio",
    page_icon="🧪",
    layout="wide"
)

st.title("🧪 Analizador de Tiempos de Operación de Laboratorio")
st.markdown("""
Esta aplicación procesa los archivos de extracción de muestras, calcula los tiempos de proceso:
- **Tiempo de Espera en Sala**: entre Ingreso y Llamado.
- **Tiempo de Punción**: entre Llamado y Extracción.
- **Tiempo de Espera Total**: entre Ingreso y Extracción.

Permite analizar tanto la **Jornada Completa** como el **Horario Punta (Ingresos hasta las 10:30 AM)** para evitar sesgos por caída de flujo.
""")

# Componente de carga de archivos
uploaded_file = st.file_uploader("Carga tu archivo de datos (CSV o Excel)", type=["csv", "xlsx", "xls"])

def procesar_datos(file):
    # Lectura dinámica según el tipo de archivo
    if file.name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
        
    # Identificar nombres de columnas de forma precisa
    col_ingreso = None
    col_llamado = None
    col_extraccion = None
    
    for c in df.columns:
        c_upper = c.strip().upper()
        if 'FECHA' in c_upper and 'INGRESO' in c_upper:
            col_ingreso = c
        elif 'FECHA' in c_upper and 'LLAMADO' in c_upper:
            col_llamado = c
        elif 'FECHA' in c_upper and 'EXTRACC' in c_upper:
            col_extraccion = c

    if not col_ingreso or not col_extraccion:
        st.error("El archivo debe contener al menos las columnas 'FECHA Y HORA DE INGRESO' y 'FECHA Y HORA DE EXTRACCIÓN'.")
        return None, None, None, None, None, None

    df_proc = df.copy()
    df_proc['INGRESO_DT'] = pd.to_datetime(df_proc[col_ingreso], dayfirst=True, errors='coerce', format='mixed')
    df_proc['EXTRACCION_DT'] = pd.to_datetime(df_proc[col_extraccion], dayfirst=True, errors='coerce', format='mixed')
    
    # Tiempo de Espera Total (Ingreso -> Extracción) en minutos
    df_proc['TIEMPO DE ESPERA TOTAL (MINUTOS)'] = (df_proc['EXTRACCION_DT'] - df_proc['INGRESO_DT']).dt.total_seconds() / 60.0
    
    # Si la columna LLAMADO existe, calcular Espera en Sala y Punción
    if col_llamado:
        df_proc['LLAMADO_DT'] = pd.to_datetime(df_proc[col_llamado], dayfirst=True, errors='coerce', format='mixed')
        df_proc['TIEMPO DE ESPERA EN SALA (MINUTOS)'] = (df_proc['LLAMADO_DT'] - df_proc['INGRESO_DT']).dt.total_seconds() / 60.0
        df_proc['TIEMPO DE PUNCIÓN (MINUTOS)'] = (df_proc['EXTRACCION_DT'] - df_proc['LLAMADO_DT']).dt.total_seconds() / 60.0
    else:
        df_proc['LLAMADO_DT'] = pd.NaT
        df_proc['TIEMPO DE ESPERA EN SALA (MINUTOS)'] = None
        df_proc['TIEMPO DE PUNCIÓN (MINUTOS)'] = None

    col_paciente = 'RUT' if 'RUT' in df_proc.columns else ('ORDEN' if 'ORDEN' in df_proc.columns else None)

    # 1. Métricas Generales (Jornada Completa)
    total_muestras = len(df_proc)
    total_pacientes = df_proc[col_paciente].nunique() if col_paciente else total_muestras
    promedio_espera_total_general = df_proc['TIEMPO DE ESPERA TOTAL (MINUTOS)'].mean()
    promedio_espera_sala_general = df_proc['TIEMPO DE ESPERA EN SALA (MINUTOS)'].mean() if col_llamado else None
    promedio_puncion_general = df_proc['TIEMPO DE PUNCIÓN (MINUTOS)'].mean() if col_llamado else None

    metrics_general = {
        'total_m': total_muestras,
        'total_p': total_pacientes,
        'prom_sala': promedio_espera_sala_general,
        'prom_punc': promedio_puncion_general,
        'prom_total': promedio_espera_total_general
    }

    # 2. Métricas Horario Punta (Ingresos hasta las 10:30 AM)
    df_1030 = df_proc[df_proc['INGRESO_DT'].dt.time <= datetime.time(10, 30)]
    total_muestras_1030 = len(df_1030)
    total_pacientes_1030 = df_1030[col_paciente].nunique() if (col_paciente and total_muestras_1030 > 0) else total_muestras_1030
    promedio_espera_total_1030 = df_1030['TIEMPO DE ESPERA TOTAL (MINUTOS)'].mean() if total_muestras_1030 > 0 else 0.0
    promedio_espera_sala_1030 = df_1030['TIEMPO DE ESPERA EN SALA (MINUTOS)'].mean() if (col_llamado and total_muestras_1030 > 0) else None
    promedio_puncion_1030 = df_1030['TIEMPO DE PUNCIÓN (MINUTOS)'].mean() if (col_llamado and total_muestras_1030 > 0) else None

    metrics_1030 = {
        'total_m': total_muestras_1030,
        'total_p': total_pacientes_1030,
        'prom_sala': promedio_espera_sala_1030,
        'prom_punc': promedio_puncion_1030,
        'prom_total': promedio_espera_total_1030
    }
    
    # Identificar columna de usuario extractor
    col_usuario = None
    for c in df_proc.columns:
        if 'EXTRAJO' in c.upper():
            col_usuario = c
            break
            
    # Resumen por usuario extractor (Jornada Completa)
    if col_usuario and col_llamado:
        summary_user = df_proc.groupby(col_usuario).agg(
            cant_muestras=(col_usuario, 'count'),
            cant_pacientes=(col_paciente, 'nunique') if col_paciente else (col_usuario, 'count'),
            tiempo_puncion_prom=('TIEMPO DE PUNCIÓN (MINUTOS)', 'mean')
        ).reset_index()
        summary_user.columns = ['Usuario de Extracción', 'Cantidad de Muestras', 'Cantidad de Pacientes', 'Tiempo Promedio de Punción (Minutos)']
        summary_user = summary_user.sort_values(by='Tiempo Promedio de Punción (Minutos)', ascending=False)
    elif col_usuario:
        summary_user = df_proc.groupby(col_usuario).agg(
            cant_muestras=(col_usuario, 'count'),
            cant_pacientes=(col_paciente, 'nunique') if col_paciente else (col_usuario, 'count')
        ).reset_index()
        summary_user.columns = ['Usuario de Extracción', 'Cantidad de Muestras', 'Cantidad de Pacientes']
    else:
        summary_user = pd.DataFrame()

    # Resumen por usuario extractor (Horario Punta - Hasta 10:30 AM)
    if col_usuario and col_llamado:
        if len(df_1030) > 0:
            summary_user_1030 = df_1030.groupby(col_usuario).agg(
                cant_muestras=(col_usuario, 'count'),
                cant_pacientes=(col_paciente, 'nunique') if col_paciente else (col_usuario, 'count'),
                tiempo_puncion_prom=('TIEMPO DE PUNCIÓN (MINUTOS)', 'mean')
            ).reset_index()
            summary_user_1030.columns = ['Usuario de Extracción', 'Cantidad de Muestras', 'Cantidad de Pacientes', 'Tiempo Promedio de Punción (Minutos)']
            summary_user_1030 = summary_user_1030.sort_values(by='Tiempo Promedio de Punción (Minutos)', ascending=False)
        else:
            summary_user_1030 = pd.DataFrame(columns=['Usuario de Extracción', 'Cantidad de Muestras', 'Cantidad de Pacientes', 'Tiempo Promedio de Punción (Minutos)'])
    elif col_usuario:
        if len(df_1030) > 0:
            summary_user_1030 = df_1030.groupby(col_usuario).agg(
                cant_muestras=(col_usuario, 'count'),
                cant_pacientes=(col_paciente, 'nunique') if col_paciente else (col_usuario, 'count')
            ).reset_index()
            summary_user_1030.columns = ['Usuario de Extracción', 'Cantidad de Muestras', 'Cantidad de Pacientes']
        else:
            summary_user_1030 = pd.DataFrame(columns=['Usuario de Extracción', 'Cantidad de Muestras', 'Cantidad de Pacientes'])
    else:
        summary_user_1030 = pd.DataFrame()

    # Resumen por procedencia
    col_proc = 'PROCEDENCIA' if 'PROCEDENCIA' in df_proc.columns else None
    if col_proc:
        if col_llamado:
            summary_proc = df_proc.groupby(col_proc).agg(
                cant_muestras=(col_proc, 'count'),
                cant_pacientes=(col_paciente, 'nunique') if col_paciente else (col_proc, 'count'),
                tiempo_espera_sala_prom=('TIEMPO DE ESPERA EN SALA (MINUTOS)', 'mean'),
                tiempo_puncion_prom=('TIEMPO DE PUNCIÓN (MINUTOS)', 'mean'),
                tiempo_espera_total_prom=('TIEMPO DE ESPERA TOTAL (MINUTOS)', 'mean')
            ).reset_index()
            summary_proc.columns = [
                'Procedencia / Servicio', 'Cantidad de Muestras', 'Cantidad de Pacientes', 
                'Tiempo Promedio Espera en Sala (Minutos)', 'Tiempo Promedio Punción (Minutos)', 'Tiempo Promedio Espera Total (Minutos)'
            ]
            summary_proc = summary_proc.sort_values(by='Tiempo Promedio Espera Total (Minutos)', ascending=False)
        else:
            summary_proc = df_proc.groupby(col_proc).agg(
                cant_muestras=(col_proc, 'count'),
                cant_pacientes=(col_paciente, 'nunique') if col_paciente else (col_proc, 'count'),
                tiempo_espera_total_prom=('TIEMPO DE ESPERA TOTAL (MINUTOS)', 'mean')
            ).reset_index()
            summary_proc.columns = [
                'Procedencia / Servicio', 'Cantidad de Muestras', 'Cantidad de Pacientes', 
                'Tiempo Promedio Espera Total (Minutos)'
            ]
            summary_proc = summary_proc.sort_values(by='Tiempo Promedio Espera Total (Minutos)', ascending=False)
    else:
        summary_proc = pd.DataFrame()
        
    # Eliminar columnas temporales de datetime antes de exportar
    df_detalles = df_proc.drop(columns=['INGRESO_DT', 'EXTRACCION_DT', 'LLAMADO_DT'], errors='ignore')
    
    return df_detalles, summary_user, summary_user_1030, summary_proc, metrics_general, metrics_1030

def generar_excel_profesional(df_details, summary_user, summary_user_1030, summary_proc, metrics_general, metrics_1030):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # Estilos del reporte
    font_family = "Segoe UI"
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid") # Slate/Navy
    accent_fill = PatternFill(start_color="4A6572", end_color="4A6572", fill_type="solid") # Muted Steel Blue
    zebra_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    kpi_fill_1 = PatternFill(start_color="EAEDED", end_color="EAEDED", fill_type="solid")
    kpi_fill_2 = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid") # Mint suave para 10:30
    
    font_title = Font(name=font_family, size=15, bold=True, color="2C3E50")
    font_subtitle = Font(name=font_family, size=11, bold=True, color="34495E")
    font_section = Font(name=font_family, size=12, bold=True, color="2C3E50")
    font_regular = Font(name=font_family, size=11)
    font_kpi_val = Font(name=font_family, size=16, bold=True, color="2C3E50")
    font_kpi_lbl = Font(name=font_family, size=9, italic=True, color="5D6D7E")
    
    thin_side = Side(border_style="thin", color="BDC3C7")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    
    # ---- HOJA 1: RESUMEN ----
    ws_resumen = wb.active
    ws_resumen.title = "Resumen y Métricas"
    ws_resumen.views.sheetView[0].showGridLines = True
    
    ws_resumen["B2"] = "REPORTE DE TIEMPOS DE OPERACIÓN - LABORATORIO"
    ws_resumen["B2"].font = font_title
    
    # SECCIÓN 1: KPI GENERALES (JORNADA COMPLETA)
    ws_resumen["B4"] = "INDICADORES GENERALES (JORNADA COMPLETA)"
    ws_resumen["B4"].font = font_subtitle
    
    # Fila 5-6: Muestras y Pacientes
    ws_resumen.merge_cells("B5:C5")
    ws_resumen["B5"] = "TOTAL MUESTRAS"
    ws_resumen["B5"].font = font_kpi_lbl; ws_resumen["B5"].alignment = align_center; ws_resumen["B5"].fill = kpi_fill_1
    
    ws_resumen.merge_cells("B6:C6")
    ws_resumen["B6"] = metrics_general['total_m']
    ws_resumen["B6"].font = font_kpi_val; ws_resumen["B6"].alignment = align_center; ws_resumen["B6"].fill = kpi_fill_1; ws_resumen["B6"].number_format = "#,##0"
    
    ws_resumen.merge_cells("D5:E5")
    ws_resumen["D5"] = "TOTAL PACIENTES"
    ws_resumen["D5"].font = font_kpi_lbl; ws_resumen["D5"].alignment = align_center; ws_resumen["D5"].fill = kpi_fill_1
    
    ws_resumen.merge_cells("D6:E6")
    ws_resumen["D6"] = metrics_general['total_p']
    ws_resumen["D6"].font = font_kpi_val; ws_resumen["D6"].alignment = align_center; ws_resumen["D6"].fill = kpi_fill_1; ws_resumen["D6"].number_format = "#,##0"
    
    # Fila 7-8: Tiempos Generales
    if metrics_general['prom_sala'] is not None and metrics_general['prom_punc'] is not None:
        ws_resumen.merge_cells("B7:C7")
        ws_resumen["B7"] = "ESPERA EN SALA PROM."
        ws_resumen["B7"].font = font_kpi_lbl; ws_resumen["B7"].alignment = align_center; ws_resumen["B7"].fill = kpi_fill_1
        
        ws_resumen.merge_cells("B8:C8")
        ws_resumen["B8"] = metrics_general['prom_sala']
        ws_resumen["B8"].font = font_kpi_val; ws_resumen["B8"].alignment = align_center; ws_resumen["B8"].fill = kpi_fill_1; ws_resumen["B8"].number_format = "0.00"
        
        ws_resumen.merge_cells("D7:E7")
        ws_resumen["D7"] = "PUNCIÓN PROMEDIO"
        ws_resumen["D7"].font = font_kpi_lbl; ws_resumen["D7"].alignment = align_center; ws_resumen["D7"].fill = kpi_fill_1
        
        ws_resumen.merge_cells("D8:E8")
        ws_resumen["D8"] = metrics_general['prom_punc']
        ws_resumen["D8"].font = font_kpi_val; ws_resumen["D8"].alignment = align_center; ws_resumen["D8"].fill = kpi_fill_1; ws_resumen["D8"].number_format = "0.00"

        ws_resumen.merge_cells("F7:G7")
        ws_resumen["F7"] = "ESPERA TOTAL PROM."
        ws_resumen["F7"].font = font_kpi_lbl; ws_resumen["F7"].alignment = align_center; ws_resumen["F7"].fill = kpi_fill_1
        
        ws_resumen.merge_cells("F8:G8")
        ws_resumen["F8"] = metrics_general['prom_total']
        ws_resumen["F8"].font = font_kpi_val; ws_resumen["F8"].alignment = align_center; ws_resumen["F8"].fill = kpi_fill_1; ws_resumen["F8"].number_format = "0.00"
    else:
        ws_resumen.merge_cells("B7:C7")
        ws_resumen["B7"] = "ESPERA TOTAL PROM."
        ws_resumen["B7"].font = font_kpi_lbl; ws_resumen["B7"].alignment = align_center; ws_resumen["B7"].fill = kpi_fill_1
        
        ws_resumen.merge_cells("B8:C8")
        ws_resumen["B8"] = metrics_general['prom_total'] if metrics_general['prom_total'] is not None else 0.0
        ws_resumen["B8"].font = font_kpi_val; ws_resumen["B8"].alignment = align_center; ws_resumen["B8"].fill = kpi_fill_1; ws_resumen["B8"].number_format = "0.00"

    for r in [5, 6]:
        for c in [2, 3, 4, 5]:
            ws_resumen.cell(row=r, column=c).border = border_all
    for r in [7, 8]:
        cols = [2, 3, 4, 5, 6, 7] if (metrics_general['prom_sala'] is not None) else [2, 3]
        for c in cols:
            ws_resumen.cell(row=r, column=c).border = border_all

    # SECCIÓN 2: KPI HORARIO PUNTA (HASTA 10:30 AM)
    ws_resumen["B10"] = "INDICADORES HORARIO PUNTA (INGRESOS HASTA LAS 10:30 AM)"
    ws_resumen["B10"].font = font_subtitle
    
    # Fila 11-12: Muestras y Pacientes < 10:30
    ws_resumen.merge_cells("B11:C11")
    ws_resumen["B11"] = "MUESTRAS (HASTA 10:30)"
    ws_resumen["B11"].font = font_kpi_lbl; ws_resumen["B11"].alignment = align_center; ws_resumen["B11"].fill = kpi_fill_2
    
    ws_resumen.merge_cells("B12:C12")
    ws_resumen["B12"] = metrics_1030['total_m']
    ws_resumen["B12"].font = font_kpi_val; ws_resumen["B12"].alignment = align_center; ws_resumen["B12"].fill = kpi_fill_2; ws_resumen["B12"].number_format = "#,##0"
    
    ws_resumen.merge_cells("D11:E11")
    ws_resumen["D11"] = "PACIENTES (HASTA 10:30)"
    ws_resumen["D11"].font = font_kpi_lbl; ws_resumen["D11"].alignment = align_center; ws_resumen["D11"].fill = kpi_fill_2
    
    ws_resumen.merge_cells("D12:E12")
    ws_resumen["D12"] = metrics_1030['total_p']
    ws_resumen["D12"].font = font_kpi_val; ws_resumen["D12"].alignment = align_center; ws_resumen["D12"].fill = kpi_fill_2; ws_resumen["D12"].number_format = "#,##0"
    
    # Fila 13-14: Tiempos < 10:30
    if metrics_1030['prom_sala'] is not None and metrics_1030['prom_punc'] is not None:
        ws_resumen.merge_cells("B13:C13")
        ws_resumen["B13"] = "ESPERA EN SALA (<10:30)"
        ws_resumen["B13"].font = font_kpi_lbl; ws_resumen["B13"].alignment = align_center; ws_resumen["B13"].fill = kpi_fill_2
        
        ws_resumen.merge_cells("B14:C14")
        ws_resumen["B14"] = metrics_1030['prom_sala']
        ws_resumen["B14"].font = font_kpi_val; ws_resumen["B14"].alignment = align_center; ws_resumen["B14"].fill = kpi_fill_2; ws_resumen["B14"].number_format = "0.00"
        
        ws_resumen.merge_cells("D13:E13")
        ws_resumen["D13"] = "PUNCIÓN PROM. (<10:30)"
        ws_resumen["D13"].font = font_kpi_lbl; ws_resumen["D13"].alignment = align_center; ws_resumen["D13"].fill = kpi_fill_2
        
        ws_resumen.merge_cells("D14:E14")
        ws_resumen["D14"] = metrics_1030['prom_punc']
        ws_resumen["D14"].font = font_kpi_val; ws_resumen["D14"].alignment = align_center; ws_resumen["D14"].fill = kpi_fill_2; ws_resumen["D14"].number_format = "0.00"

        ws_resumen.merge_cells("F13:G13")
        ws_resumen["F13"] = "ESPERA TOTAL (<10:30)"
        ws_resumen["F13"].font = font_kpi_lbl; ws_resumen["F13"].alignment = align_center; ws_resumen["F13"].fill = kpi_fill_2
        
        ws_resumen.merge_cells("F14:G14")
        ws_resumen["F14"] = metrics_1030['prom_total']
        ws_resumen["F14"].font = font_kpi_val; ws_resumen["F14"].alignment = align_center; ws_resumen["F14"].fill = kpi_fill_2; ws_resumen["F14"].number_format = "0.00"
    else:
        ws_resumen.merge_cells("B13:C13")
        ws_resumen["B13"] = "ESPERA TOTAL (<10:30)"
        ws_resumen["B13"].font = font_kpi_lbl; ws_resumen["B13"].alignment = align_center; ws_resumen["B13"].fill = kpi_fill_2
        
        ws_resumen.merge_cells("B14:C14")
        ws_resumen["B14"] = metrics_1030['prom_total'] if metrics_1030['prom_total'] is not None else 0.0
        ws_resumen["B14"].font = font_kpi_val; ws_resumen["B14"].alignment = align_center; ws_resumen["B14"].fill = kpi_fill_2; ws_resumen["B14"].number_format = "0.00"

    for r in [11, 12]:
        for c in [2, 3, 4, 5]:
            ws_resumen.cell(row=r, column=c).border = border_all
    for r in [13, 14]:
        cols = [2, 3, 4, 5, 6, 7] if (metrics_1030['prom_sala'] is not None) else [2, 3]
        for c in cols:
            ws_resumen.cell(row=r, column=c).border = border_all
            
    # Tabla Usuarios (Jornada Completa)
    start_r_user = 17
    ws_resumen.cell(row=start_r_user, column=2, value="Promedio de Tiempo por Usuario de Extracción (Jornada Completa)").font = font_section
    
    r_idx = start_r_user + 2
    if not summary_user.empty:
        for c_idx, h in enumerate(summary_user.columns, start=2):
            cell = ws_resumen.cell(row=start_r_user+1, column=c_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = border_all
            
        for _, row in summary_user.iterrows():
            ws_resumen.cell(row=r_idx, column=2, value=row[summary_user.columns[0]]).alignment = align_left
            ws_resumen.cell(row=r_idx, column=3, value=row[summary_user.columns[1]]).number_format = "#,##0"
            ws_resumen.cell(row=r_idx, column=3).alignment = align_right
            ws_resumen.cell(row=r_idx, column=4, value=row[summary_user.columns[2]]).number_format = "#,##0"
            ws_resumen.cell(row=r_idx, column=4).alignment = align_right
            
            if len(summary_user.columns) > 3:
                cell_v = ws_resumen.cell(row=r_idx, column=5, value=row[summary_user.columns[3]])
                cell_v.number_format = "0.00"
                cell_v.alignment = align_right
            
            for col_c in range(2, 2 + len(summary_user.columns)):
                c = ws_resumen.cell(row=r_idx, column=col_c)
                c.font = font_regular
                c.border = border_all
                if r_idx % 2 == 1:
                    c.fill = zebra_fill
            r_idx += 1

    # Tabla Usuarios (Hasta las 10:30 AM) - Justo debajo
    start_r_user_1030 = r_idx + 2
    ws_resumen.cell(row=start_r_user_1030, column=2, value="Promedio de Tiempo por Usuario de Extracción (Hasta las 10:30 AM)").font = font_section
    
    r_idx = start_r_user_1030 + 2
    if not summary_user_1030.empty:
        for c_idx, h in enumerate(summary_user_1030.columns, start=2):
            cell = ws_resumen.cell(row=start_r_user_1030+1, column=c_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = border_all
            
        for _, row in summary_user_1030.iterrows():
            ws_resumen.cell(row=r_idx, column=2, value=row[summary_user_1030.columns[0]]).alignment = align_left
            ws_resumen.cell(row=r_idx, column=3, value=row[summary_user_1030.columns[1]]).number_format = "#,##0"
            ws_resumen.cell(row=r_idx, column=3).alignment = align_right
            ws_resumen.cell(row=r_idx, column=4, value=row[summary_user_1030.columns[2]]).number_format = "#,##0"
            ws_resumen.cell(row=r_idx, column=4).alignment = align_right
            
            if len(summary_user_1030.columns) > 3:
                cell_v = ws_resumen.cell(row=r_idx, column=5, value=row[summary_user_1030.columns[3]])
                cell_v.number_format = "0.00"
                cell_v.alignment = align_right
            
            for col_c in range(2, 2 + len(summary_user_1030.columns)):
                c = ws_resumen.cell(row=r_idx, column=col_c)
                c.font = font_regular
                c.border = border_all
                if r_idx % 2 == 1:
                    c.fill = zebra_fill
            r_idx += 1
            
    # Tabla Procedencia
    start_r_proc = r_idx + 2
    ws_resumen.cell(row=start_r_proc, column=2, value="Promedio de Tiempo por Procedencia / Servicio").font = font_section
    
    r_idx = start_r_proc + 2
    if not summary_proc.empty:
        for c_idx, h in enumerate(summary_proc.columns, start=2):
            cell = ws_resumen.cell(row=start_r_proc+1, column=c_idx, value=h)
            cell.font = header_font
            cell.fill = accent_fill
            cell.alignment = align_center
            cell.border = border_all
            
        for _, row in summary_proc.iterrows():
            ws_resumen.cell(row=r_idx, column=2, value=row[summary_proc.columns[0]]).alignment = align_left
            ws_resumen.cell(row=r_idx, column=3, value=row[summary_proc.columns[1]]).number_format = "#,##0"
            ws_resumen.cell(row=r_idx, column=3).alignment = align_right
            ws_resumen.cell(row=r_idx, column=4, value=row[summary_proc.columns[2]]).number_format = "#,##0"
            ws_resumen.cell(row=r_idx, column=4).alignment = align_right
            
            for c_col in range(3, len(summary_proc.columns)):
                cell_v = ws_resumen.cell(row=r_idx, column=2+c_col, value=row[summary_proc.columns[c_col]])
                cell_v.number_format = "0.00"
                cell_v.alignment = align_right
            
            for col_c in range(2, 2 + len(summary_proc.columns)):
                c = ws_resumen.cell(row=r_idx, column=col_c)
                c.font = font_regular
                c.border = border_all
                if r_idx % 2 == 1:
                    c.fill = zebra_fill
            r_idx += 1

    # ---- HOJA 2: DETALLES ----
    ws_datos = wb.create_sheet(title="Datos Detallados")
    ws_datos.views.sheetView[0].showGridLines = True
    ws_datos.freeze_panes = "A2"
    
    for c_idx, col_name in enumerate(df_details.columns, start=1):
        cell = ws_datos.cell(row=1, column=c_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border_all
        
    for r_i, row in df_details.iterrows():
        for c_i, val in enumerate(row, start=1):
            cell = ws_datos.cell(row=r_i+2, column=c_i)
            col_name = df_details.columns[c_i-1]
            
            if 'MINUTOS' in col_name.upper():
                cell.value = val
                cell.number_format = "0.0"
                cell.alignment = align_right
                cell.fill = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
            elif isinstance(val, (int, float)) and not pd.isna(val):
                cell.value = val
                cell.alignment = align_right if col_name != 'ORDEN' else align_center
                cell.number_format = "#,##0" if col_name == 'ORDEN' else "0.00"
            else:
                cell.value = "" if pd.isna(val) else str(val)
                cell.alignment = align_center if col_name in ['RUT', 'FECHA Y HORA DE INGRESO', 'FECHA Y HORA DE LLAMADO', 'FECHA Y HORA DE EXTRACCIÓN'] else align_left
                
            cell.font = font_regular
            cell.border = border_all
            if (r_i+2) % 2 == 1 and 'MINUTOS' not in col_name.upper():
                cell.fill = zebra_fill

    # Ajuste automático del ancho de columnas
    for ws in [ws_resumen, ws_datos]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value and not (ws == ws_resumen and cell.row in [2, 4, 10, start_r_user, start_r_user_1030, start_r_proc]):
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    wb.save(output)
    return output.getvalue()

if uploaded_file is not None:
    with st.spinner("Procesando datos y calculando métricas..."):
        df_details, summary_user, summary_user_1030, summary_proc, metrics_general, metrics_1030 = procesar_datos(uploaded_file)
        
    if df_details is not None:
        st.success("¡Datos procesados y calculados con éxito a partir de las marcas de tiempo!")
        
        # 1. Sección de Métricas Generales (Jornada Completa)
        st.subheader("📌 Indicadores Generales de Operación (Jornada Completa)")
        if metrics_general['prom_punc'] is not None and metrics_general['prom_sala'] is not None:
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Total Muestras", f"{metrics_general['total_m']:,}")
            c2.metric("Total Pacientes", f"{metrics_general['total_p']:,}")
            c3.metric("Espera en Sala Promedio", f"{metrics_general['prom_sala']:.2f} min", help="Ingreso -> Llamado")
            c4.metric("Punción Promedio", f"{metrics_general['prom_punc']:.2f} min", help="Llamado -> Extracción")
            c5.metric("Espera Total Promedio", f"{metrics_general['prom_total']:.2f} min", help="Ingreso -> Extracción")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Muestras", f"{metrics_general['total_m']:,}")
            c2.metric("Total Pacientes", f"{metrics_general['total_p']:,}")
            c3.metric("Espera Total Promedio", f"{metrics_general['prom_total']:.2f} min", help="Ingreso -> Extracción")
            
        # 2. Sección de Métricas Horario Punta (< 10:30 AM)
        st.subheader("⏰ Indicadores Horario Punta (Ingresos hasta las 10:30 AM)")
        st.caption("Filtro aplicado para excluir la distorsión producida por la caída de flujo de pacientes después de las 10:30 AM.")
        if metrics_1030['prom_punc'] is not None and metrics_1030['prom_sala'] is not None:
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Muestras (< 10:30)", f"{metrics_1030['total_m']:,}")
            k2.metric("Pacientes (< 10:30)", f"{metrics_1030['total_p']:,}")
            k3.metric("Espera en Sala (< 10:30)", f"{metrics_1030['prom_sala']:.2f} min", help="Ingreso -> Llamado (≤ 10:30 AM)")
            k4.metric("Punción Prom. (< 10:30)", f"{metrics_1030['prom_punc']:.2f} min", help="Llamado -> Extracción (≤ 10:30 AM)")
            k5.metric("Espera Total (< 10:30)", f"{metrics_1030['prom_total']:.2f} min", help="Ingreso -> Extracción (≤ 10:30 AM)")
        else:
            k1, k2, k3 = st.columns(3)
            k1.metric("Muestras (< 10:30)", f"{metrics_1030['total_m']:,}")
            k2.metric("Pacientes (< 10:30)", f"{metrics_1030['total_p']:,}")
            k3.metric("Espera Total (< 10:30)", f"{metrics_1030['prom_total']:.2f} min", help="Ingreso -> Extracción (≤ 10:30 AM)")
        
        # Vistas previas en la app
        st.subheader("📊 Vista Previa: Promedio de Tiempo por Usuario de Extracción (Jornada Completa)")
        st.dataframe(summary_user, use_container_width=True)

        st.subheader("⏰ Vista Previa: Promedio de Tiempo por Usuario de Extracción (Hasta las 10:30 AM)")
        st.dataframe(summary_user_1030, use_container_width=True)

        st.subheader("🏥 Vista Previa de Tiempos por Procedencia / Servicio")
        st.dataframe(summary_proc, use_container_width=True)
        
        # Generar el binario del archivo Excel estructurado
        excel_data = generar_excel_profesional(df_details, summary_user, summary_user_1030, summary_proc, metrics_general, metrics_1030)
        
        # Botón para descargar el Excel completo
        st.download_button(
            label="📥 Descargar Reporte Corporativo Excel (.xlsx)",
            data=excel_data,
            file_name="Reporte_Tiempos_Laboratorio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
