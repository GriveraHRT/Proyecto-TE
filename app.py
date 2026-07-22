import streamlit as st
import pandas as pd
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
Esta aplicación procesa los archivos de extracción de muestras, calcula los tiempos del proceso:
- **Tiempo de Espera en Sala**: entre Ingreso y Llamado.
- **Tiempo de Punción**: entre Llamado y Extracción.
- **Tiempo de Espera Total**: entre Ingreso y Extracción.

Además, genera un reporte gerencial estructurado en formato Excel.
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
        return None, None, None, None, None, None, None, None

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

    # Métricas generales
    total_muestras = len(df_proc)
    
    # Conteo de pacientes únicos (por RUT u ORDEN)
    col_paciente = 'RUT' if 'RUT' in df_proc.columns else ('ORDEN' if 'ORDEN' in df_proc.columns else None)
    total_pacientes = df_proc[col_paciente].nunique() if col_paciente else total_muestras

    promedio_espera_total_general = df_proc['TIEMPO DE ESPERA TOTAL (MINUTOS)'].mean()
    promedio_espera_sala_general = df_proc['TIEMPO DE ESPERA EN SALA (MINUTOS)'].mean() if col_llamado else None
    promedio_puncion_general = df_proc['TIEMPO DE PUNCIÓN (MINUTOS)'].mean() if col_llamado else None
    
    # Identificar columna de usuario extractor
    col_usuario = None
    for c in df_proc.columns:
        if 'EXTRAJO' in c.upper():
            col_usuario = c
            break
            
    # Resumen por usuario extractor
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
    
    return df_detalles, summary_user, summary_proc, total_muestras, total_pacientes, promedio_espera_sala_general, promedio_puncion_general, promedio_espera_total_general

def generar_excel_profesional(df_details, summary_user, summary_proc, total_muestras, total_pacientes, promedio_espera_sala_general, promedio_puncion_general, promedio_espera_total_general):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # Estilos del reporte
    font_family = "Segoe UI"
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid") # Slate/Navy
    accent_fill = PatternFill(start_color="4A6572", end_color="4A6572", fill_type="solid") # Muted Steel Blue
    zebra_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    kpi_fill = PatternFill(start_color="EAEDED", end_color="EAEDED", fill_type="solid")
    
    font_title = Font(name=font_family, size=15, bold=True, color="2C3E50")
    font_section = Font(name=font_family, size=12, bold=True, color="2C3E50")
    font_regular = Font(name=font_family, size=11)
    font_kpi_val = Font(name=font_family, size=18, bold=True, color="2C3E50")
    font_kpi_lbl = Font(name=font_family, size=9, italic=True, color="7F8C8D")
    
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
    
    # Fila 1 KPI: Muestras y Pacientes (Columnas B..C y D..E)
    ws_resumen.merge_cells("B4:C4")
    ws_resumen["B4"] = "TOTAL MUESTRAS PROCESADAS"
    ws_resumen["B4"].font = font_kpi_lbl
    ws_resumen["B4"].alignment = align_center
    ws_resumen["B4"].fill = kpi_fill
    
    ws_resumen.merge_cells("B5:C5")
    ws_resumen["B5"] = total_muestras
    ws_resumen["B5"].font = font_kpi_val
    ws_resumen["B5"].alignment = align_center
    ws_resumen["B5"].fill = kpi_fill
    ws_resumen["B5"].number_format = "#,##0"
    
    ws_resumen.merge_cells("D4:E4")
    ws_resumen["D4"] = "TOTAL PACIENTES ATENDIDOS"
    ws_resumen["D4"].font = font_kpi_lbl
    ws_resumen["D4"].alignment = align_center
    ws_resumen["D4"].fill = kpi_fill
    
    ws_resumen.merge_cells("D5:E5")
    ws_resumen["D5"] = total_pacientes
    ws_resumen["D5"].font = font_kpi_val
    ws_resumen["D5"].alignment = align_center
    ws_resumen["D5"].fill = kpi_fill
    ws_resumen["D5"].number_format = "#,##0"
    
    # Fila 2 KPI: Tiempos Promedio Generales (Espera en Sala, Punción y Espera Total)
    if promedio_espera_sala_general is not None and promedio_puncion_general is not None:
        # 3 recuadros en fila 7: B7:C7, D7:E7, F7:G7
        ws_resumen.merge_cells("B7:C7")
        ws_resumen["B7"] = "TIEMPO ESPERA EN SALA PROMEDIO"
        ws_resumen["B7"].font = font_kpi_lbl
        ws_resumen["B7"].alignment = align_center
        ws_resumen["B7"].fill = kpi_fill
        
        ws_resumen.merge_cells("B8:C8")
        ws_resumen["B8"] = promedio_espera_sala_general
        ws_resumen["B8"].font = font_kpi_val
        ws_resumen["B8"].alignment = align_center
        ws_resumen["B8"].fill = kpi_fill
        ws_resumen["B8"].number_format = "0.00"
        
        ws_resumen.merge_cells("D7:E7")
        ws_resumen["D7"] = "TIEMPO PUNCIÓN PROMEDIO GENERAL"
        ws_resumen["D7"].font = font_kpi_lbl
        ws_resumen["D7"].alignment = align_center
        ws_resumen["D7"].fill = kpi_fill
        
        ws_resumen.merge_cells("D8:E8")
        ws_resumen["D8"] = promedio_puncion_general
        ws_resumen["D8"].font = font_kpi_val
        ws_resumen["D8"].alignment = align_center
        ws_resumen["D8"].fill = kpi_fill
        ws_resumen["D8"].number_format = "0.00"

        ws_resumen.merge_cells("F7:G7")
        ws_resumen["F7"] = "TIEMPO ESPERA TOTAL PROMEDIO"
        ws_resumen["F7"].font = font_kpi_lbl
        ws_resumen["F7"].alignment = align_center
        ws_resumen["F7"].fill = kpi_fill
        
        ws_resumen.merge_cells("F8:G8")
        ws_resumen["F8"] = promedio_espera_total_general
        ws_resumen["F8"].font = font_kpi_val
        ws_resumen["F8"].alignment = align_center
        ws_resumen["F8"].fill = kpi_fill
        ws_resumen["F8"].number_format = "0.00"
    else:
        ws_resumen.merge_cells("B7:C7")
        ws_resumen["B7"] = "TIEMPO ESPERA TOTAL PROMEDIO"
        ws_resumen["B7"].font = font_kpi_lbl
        ws_resumen["B7"].alignment = align_center
        ws_resumen["B7"].fill = kpi_fill
        
        ws_resumen.merge_cells("B8:C8")
        ws_resumen["B8"] = promedio_espera_total_general if promedio_espera_total_general is not None else 0.0
        ws_resumen["B8"].font = font_kpi_val
        ws_resumen["B8"].alignment = align_center
        ws_resumen["B8"].fill = kpi_fill
        ws_resumen["B8"].number_format = "0.00"

    # Aplicar bordes a recuadros KPI
    for r in [4, 5]:
        for c in [2, 3, 4, 5]:
            ws_resumen.cell(row=r, column=c).border = border_all
    for r in [7, 8]:
        cols = [2, 3, 4, 5, 6, 7] if (promedio_espera_sala_general is not None) else [2, 3]
        for c in cols:
            ws_resumen.cell(row=r, column=c).border = border_all
            
    # Tabla Usuarios
    ws_resumen["B11"] = "Promedio de Tiempo por Usuario de Extracción"
    ws_resumen["B11"].font = font_section
    
    r_idx = 13
    if not summary_user.empty:
        for c_idx, h in enumerate(summary_user.columns, start=2):
            cell = ws_resumen.cell(row=12, column=c_idx, value=h)
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
            
    # Tabla Procedencia
    start_r_proc = r_idx + 2
    ws_resumen.cell(row=start_r_proc, column=2, value="Promedio de Tiempo por Procedencia / Servicio").font = font_section
    
    if not summary_proc.empty:
        for c_idx, h in enumerate(summary_proc.columns, start=2):
            cell = ws_resumen.cell(row=start_r_proc+1, column=c_idx, value=h)
            cell.font = header_font
            cell.fill = accent_fill
            cell.alignment = align_center
            cell.border = border_all
            
        r_idx = start_r_proc + 2
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
                if cell.value and not (ws == ws_resumen and cell.row == 2):
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    wb.save(output)
    return output.getvalue()

if uploaded_file is not None:
    with st.spinner("Procesando datos y calculando métricas..."):
        df_details, summary_user, summary_proc, total_m, total_p, prom_sala, prom_punc, prom_total = procesar_datos(uploaded_file)
        
    if df_details is not None:
        st.success("¡Datos procesados y calculados con éxito a partir de las marcas de tiempo!")
        
        # Mostrar tarjetas métricas en la interfaz
        st.subheader("📌 Indicadores Generales de Operación")
        if prom_punc is not None and prom_sala is not None:
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Total Muestras Procesadas", f"{total_m:,}")
            c2.metric("Total Pacientes Atendidos", f"{total_p:,}")
            c3.metric("Tiempo Espera en Sala Promedio", f"{prom_sala:.2f} min", help="Ingreso -> Llamado")
            c4.metric("Tiempo Punción Promedio General", f"{prom_punc:.2f} min", help="Llamado -> Extracción")
            c5.metric("Tiempo Espera Total Promedio", f"{prom_total:.2f} min", help="Ingreso -> Extracción")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Muestras Procesadas", f"{total_m:,}")
            c2.metric("Total Pacientes Atendidos", f"{total_p:,}")
            c3.metric("Tiempo Espera Total Promedio", f"{prom_total:.2f} min", help="Ingreso -> Extracción")
        
        # Vistas previas en la app
        st.subheader("📊 Vista Previa de Tiempos por Usuario Extractor")
        st.dataframe(summary_user, use_container_width=True)

        st.subheader("🏥 Vista Previa de Tiempos por Procedencia / Servicio")
        st.dataframe(summary_proc, use_container_width=True)
        
        # Generar el binario del archivo Excel estructurado
        excel_data = generar_excel_profesional(df_details, summary_user, summary_proc, total_m, total_p, prom_sala, prom_punc, prom_total)
        
        # Botón para descargar el Excel completo
        st.download_button(
            label="📥 Descargar Reporte Corporativo Excel (.xlsx)",
            data=excel_data,
            file_name="Reporte_Tiempos_Laboratorio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
