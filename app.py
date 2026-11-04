# camelot_extractor_pro_final_v3.py
"""
Extractor Profesional de PDFs con Camelot + Dashboard Inteligente
VERSIÓN 3.2 - Extracción por coordenadas exactas (pdfplumber), columnas 13/15
corregidas, validación contra el pie de página del PDF, filtros globales y
dashboards de demoras / centros de costo.
"""

import streamlit as st
import pandas as pd
import camelot
import pdfplumber
import hashlib
import tempfile
import os
import re
from datetime import datetime, timedelta
import io
from typing import List, Dict, Tuple, Optional
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import holidays

st.set_page_config(
    page_title="Camelot PDF Extractor Pro v3.2",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# HEADER PROFESIONAL
# ============================================================================

def render_header():
    """Renderiza header profesional de la aplicación"""
    st.markdown("""
        <div style='background: linear-gradient(90deg, #1f77b4 0%, #2ca02c 100%);
                    padding: 20px; border-radius: 10px; margin-bottom: 30px;'>
            <h1 style='text-align: center; color: white; margin: 0;'>
                📄 Camelot PDF Extractor Pro
            </h1>
            <h4 style='text-align: center; color: #f0f0f0; margin: 10px 0 0 0;'>
                Sistema Profesional de Extracción y Análisis de Albaranes v3.2
            </h4>
        </div>
    """, unsafe_allow_html=True)

# ============================================================================
# CONSTANTES Y HELPERS DE DATOS (reporte Alsina "Outstanding count returns")
# ============================================================================

COLUMN_NAMES = [
    'Wh', 'Return_Prefix', 'Return_Slip', 'Return_Date',
    'Jobsite', 'Cost_Center', 'Invoice_Date1', 'Invoice_Date2',
    'Customer', 'Job_Name', 'Definitive', 'Counted_Date',
    'Tablets', 'Total_Tablets', 'Open_Tablets', 'Open_Total',
    'Counting_Delay', 'Validation_Delay'
]

# Fronteras X exactas (x0, x1) de las 18 columnas del reporte de Alsina.
COL_BOUNDS = [
    (7.0, 27.12),     # 0: Wh (Subsidiary/State: FL, TX, DL, etc.)
    (27.12, 47.04),   # 1: Return_Prefix (Warehouse code: 61D, 612D, 28D, etc.)
    (47.04, 99.60),   # 2: Return_Slip (7290000XXXXX)
    (99.60, 146.28),  # 3: Return_Date
    (146.28, 187.32), # 4: Jobsite
    (187.32, 228.00), # 5: Cost_Center
    (228.00, 274.80), # 6: Invoice_Date1 (puede estar vacía)
    (274.80, 321.84), # 7: Invoice_Date2
    (321.84, 411.48), # 8: Customer
    (411.48, 501.36), # 9: Job_Name
    (501.36, 514.56), # 10: Definitive (Yes / No; "Yes" se parte en "Ye" + "s")
    (514.56, 562.20), # 11: Counted_Date (vacía cuando Definitive="No")
    (562.20, 648.96), # 12: Tablets (todas las tablillas del albarán)
    (648.96, 664.08), # 13: Total_Tablets
    (664.08, 750.72), # 14: Open_Tablets (códigos abiertos con sufijo [MALT])
    (750.72, 765.72), # 15: Open_Total (0 cuando Definitive="Yes")
    (765.72, 794.04), # 16: Counting_Delay
    (794.04, 830.00)  # 17: Validation_Delay
]

CAMELOT_STREAM_COLUMNS = [
    '27.12,47.04,99.60,146.28,187.32,228.00,274.80,321.84,411.48,501.36,'
    '514.56,562.20,648.96,664.08,750.72,765.72,794.04'
]

SLIP_PATTERN = re.compile(r'^7290000\d{5}$')
SLIP_SEARCH_PATTERN = re.compile(r'7290000\d{5}')
HEADER_OFFSET = 497          # respaldo: el encabezado termina en page.height - 497
HEADER_LINE_RANGE = (40.0, 200.0)  # rango de 'top' donde buscar la línea bajo el encabezado
ROW_TOP_MARGIN = 2.5         # margen superior de la ventana vertical de cada fila
ROW_HEIGHT = 11.0            # alto base de una fila de 2 líneas
ROW_MAX_HEIGHT = 18.0        # alto máximo si la fila ocupa 3 líneas
ROW_LINE_GAP = 10.0          # separación máxima entre líneas consecutivas de una misma fila
MULTILINE_COLS = (8, 9, 10, 12, 14)  # Customer, Job_Name, Definitive, Tablets, Open_Tablets
CRITICAL_DELAY_DAYS = 20     # Counting_Delay >= 20 días => albarán crítico
EMPTY_MARKERS = ('', 'nan', 'none', 'nat')


def clean_cell(value) -> str:
    """Convierte cualquier valor de celda a texto limpio ('' para vacío/NaN)."""
    if value is None:
        return ''
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return '' if text.lower() in EMPTY_MARKERS else text


def parse_int(value) -> Optional[int]:
    """Entero de una celda ('4', '4.0') o None si no es un entero válido."""
    text = clean_cell(value)
    if re.fullmatch(r'\d+(\.0+)?', text):
        return int(float(text))
    return None


def normalize_definitive(value) -> str:
    """Une 'Ye' + 's' (y variantes YES, Ye s...) en 'Yes'; 'No' en 'No'."""
    text = re.sub(r'\s+', '', clean_cell(value)).lower()
    if text in ('y', 'ye', 'yes'):
        return 'Yes'
    if text in ('n', 'no'):
        return 'No'
    return clean_cell(value)


def split_tablet_codes(codes_str) -> List[str]:
    """Lista de códigos de una celda de tablillas (ignora comas vacías)."""
    return [c for c in re.split(r'[,\s]+', clean_cell(codes_str)) if c]


def count_tablet_codes(codes_str) -> int:
    """Cantidad de códigos de tablillas en una celda (Tablets u Open_Tablets)."""
    return len([c for c in split_tablet_codes(codes_str) if re.search(r'\d', c)])


def sanitize_tablet_codes(tablets_str: str, expected_total: int) -> str:
    """
    Repara códigos concatenados por el ERP (ej. '244, 14901656, , 1666').

    Un token de 8 dígitos se separa en dos códigos de 4 dígitos solo si, tras
    separarlo, la cantidad de códigos coincide con expected_total. Siempre
    elimina comas vacías dobles (', ,') y comas finales.
    """
    tokens = split_tablet_codes(tablets_str)
    expected = parse_int(expected_total)

    if expected is not None and len(tokens) != expected:
        glued = [i for i, t in enumerate(tokens) if re.fullmatch(r'(\d{4})(\d{4})[A-Za-z]?', t)]
        if glued and len(tokens) + len(glued) == expected:
            fixed = []
            for token in tokens:
                match = re.fullmatch(r'(\d{4})(\d{4})([A-Za-z]?)', token)
                if match:
                    fixed.extend([match.group(1), match.group(2) + match.group(3)])
                else:
                    fixed.append(token)
            tokens = fixed

    return ', '.join(tokens)


def to_named_df(df: pd.DataFrame) -> pd.DataFrame:
    """Copia con las primeras 18 columnas renombradas (si existen)."""
    named = df.copy()
    if len(named.columns) >= 18:
        named = named.iloc[:, :18].copy()
        named.columns = COLUMN_NAMES
    return named


def column_text(df: pd.DataFrame, idx: int) -> pd.Series:
    """Columna posicional como texto limpio (vacío para NaN o columna inexistente)."""
    if len(df.columns) <= idx:
        return pd.Series([''] * len(df), index=df.index, dtype=object)
    return df.iloc[:, idx].map(clean_cell)


def column_numeric(df: pd.DataFrame, idx: int) -> pd.Series:
    """Columna posicional convertida a número (NaN si no es numérica o no existe)."""
    if len(df.columns) <= idx:
        return pd.Series([np.nan] * len(df), index=df.index, dtype=float)
    return pd.to_numeric(df.iloc[:, idx], errors='coerce')


class CoordinateTable:
    """Tabla extraída por coordenadas; expone la misma interfaz mínima que Camelot."""

    is_coordinate_exact = True

    def __init__(self, df: pd.DataFrame, page_number: int):
        self.df = df
        self.page = page_number
        self.accuracy = 100.0


# ============================================================================
# CLASE PRINCIPAL: EXTRACTOR
# ============================================================================

class CamelotExtractorPro:
    """
    Extractor especializado - versión profesional con 8 correcciones universales
    
    Versión 3.2 - Motor principal por coordenadas exactas (pdfplumber) con las
    18 columnas fijas del reporte de Alsina; los métodos de Camelot quedan como
    respaldo (stream con fronteras X exactas).
    """

    def __init__(self):
        self.extraction_methods = [
            self.method_coordinate_precision,  # PRIORIDAD 1: coordenadas exactas (pdfplumber)
            self.method_stream_standard,       # PRIORIDAD 2: stream con fronteras X exactas
            self.method_stream_balanced,       # PRIORIDAD 3
            self.method_lattice_standard,
            self.method_stream_aggressive,
            self.method_lattice_detailed,
            self.method_hybrid
        ]

    def extract_with_all_methods(self, pdf_path: str) -> Dict:
        """Prueba todos los métodos y compara resultados"""
        results = {}
        st.session_state.pop('pdf_footer_totals', None)

        for method in self.extraction_methods:
            method_name = method.__name__
            try:
                with st.spinner(f"Probando {method_name}..."):
                    tables = method(pdf_path)
                    if tables:
                        df = self.process_tables(tables)
                        results[method_name] = {
                            'success': True,
                            'tables_found': len(tables),
                            'rows': len(df) if df is not None else 0,
                            'data': df,
                            'accuracy': self.calculate_accuracy(tables)
                        }
                    else:
                        results[method_name] = {'success': False}
            except Exception as e:
                results[method_name] = {'success': False, 'error': str(e)}

        return results

    # ========================================================================
    # CORRECCIONES UNIVERSALES
    # ========================================================================

    def merge_continuation_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Une filas de continuación en DOS columnas:
        - Columna 12 (Tablets): números sin sufijos
        - Columna 14 (Open): números CON sufijos [MALT]
        """
        try:
            merged_rows = []
            skip_next = False

            for idx in range(len(df)):
                if skip_next:
                    skip_next = False
                    continue

                row = df.iloc[idx:idx+1].copy()
                row_text = ' '.join(str(cell) for cell in row.iloc[0].values if pd.notna(cell))

                if re.search(r'7290000\d{5}', row_text):
                    continuation_found = False

                    # CASO 1: TABLETS (col 12)
                    if len(row.columns) > 12:
                        current_tablets = str(row.iloc[0, 12]).strip()

                        if current_tablets.endswith(',') and idx + 1 < len(df):
                            next_row = df.iloc[idx + 1]
                            next_text = ' '.join(str(cell) for cell in next_row.values if pd.notna(cell))

                            if not re.search(r'7290000\d{5}', next_text):
                                found_numbers = re.findall(r'\b\d{2,4}\b', next_text)

                                if found_numbers:
                                    numbers_str = ', '.join(found_numbers)
                                    row.iloc[0, 12] = current_tablets + ' ' + numbers_str
                                    continuation_found = True

                    # CASO 2: OPEN (col 14)
                    if len(row.columns) > 14:
                        current_open = str(row.iloc[0, 14]).strip()

                        if current_open.endswith(',') and idx + 1 < len(df):
                            next_row = df.iloc[idx + 1]
                            next_text = ' '.join(str(cell) for cell in next_row.values if pd.notna(cell))

                            if not re.search(r'7290000\d{5}', next_text):
                                found_codes = re.findall(r'\d{2,4}[MALT]', next_text)

                                if found_codes:
                                    codes_str = ', '.join(found_codes)
                                    row.iloc[0, 14] = current_open + ' ' + codes_str
                                    continuation_found = True

                    if continuation_found:
                        skip_next = True

                    # Limpiar comas finales
                    if len(row.columns) > 12:
                        tablets_value = str(row.iloc[0, 12]).strip()
                        if tablets_value.endswith(','):
                            row.iloc[0, 12] = tablets_value.rstrip(',').strip()

                    if len(row.columns) > 14:
                        open_value = str(row.iloc[0, 14]).strip()
                        if open_value.endswith(','):
                            row.iloc[0, 14] = open_value.rstrip(',').strip()

                merged_rows.append(row)

            if merged_rows:
                return pd.concat(merged_rows, ignore_index=True)
            return df

        except Exception as e:
            st.error(f"Error en merge_continuation_rows: {e}")
            return df

    def fix_missing_invoice_date1(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """
        Corrige filas sin Invoice_Date1 (col 6) en los métodos fallback de Camelot.

        Si el PDF venía sin primera fecha de factura, Camelot deja en la col 6 la
        fecha que corresponde a Invoice_Date2 y en la col 7 el nombre del cliente.
        En ese caso se inserta '' en la col 6 y se desplazan las cols 6..16 una
        posición a la derecha.
        """
        try:
            if len(row_data.columns) < 18:
                return row_data

            date_pattern = r'^\d{1,2}/\d{1,2}/\d{4}$'
            col_6 = clean_cell(row_data.iloc[0, 6])
            col_7 = clean_cell(row_data.iloc[0, 7])

            if re.match(date_pattern, col_6) and col_7 and not re.match(date_pattern, col_7):
                for col_idx in range(17, 6, -1):
                    row_data.iloc[0, col_idx] = row_data.iloc[0, col_idx - 1]
                row_data.iloc[0, 6] = ''

            return row_data
        except Exception:
            return row_data

    def fix_missing_open_column(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """
        Corrige el desplazamiento cuando Open_Tablets (col 14) está vacía.

        Con Definitive = "Yes" la col 14 está vacía en el PDF y Camelot puede
        desplazar [Open_Total, Counting_Delay, Validation_Delay] una posición a
        la izquierda.

        - Col 14 vacía y cols 15..17 ya con valores: NO se desplaza nada.
        - Col 17 vacía y col 14 con un entero (0): se desplazan cols 14..16 a la
          derecha y la col 14 queda vacía.
        """
        try:
            if len(row_data.columns) < 18:
                return row_data

            if normalize_definitive(row_data.iloc[0, 10]) != 'Yes':
                return row_data

            col_14 = clean_cell(row_data.iloc[0, 14])
            col_17 = clean_cell(row_data.iloc[0, 17])

            if col_14 == '' and all(clean_cell(row_data.iloc[0, i]) != '' for i in (15, 16, 17)):
                return row_data

            if col_17 == '' and re.fullmatch(r'\d+', col_14):
                for col_idx in range(17, 14, -1):
                    row_data.iloc[0, col_idx] = row_data.iloc[0, col_idx - 1]
                row_data.iloc[0, 14] = ''

            return row_data
        except Exception:
            return row_data

    def normalize_row(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """
        Normalización final de una fila (aplica a todos los métodos):
        - Wh y Return_Prefix en mayúsculas (612d -> 612D)
        - Definitive: 'Ye' + 's' -> 'Yes', 'No' -> 'No'
        - Totales numéricos limpios y códigos de tablillas saneados
        """
        try:
            if len(row_data.columns) < 18:
                return row_data

            for col_idx in range(len(row_data.columns)):
                row_data.iloc[0, col_idx] = clean_cell(row_data.iloc[0, col_idx])

            row_data.iloc[0, 0] = clean_cell(row_data.iloc[0, 0]).upper()
            row_data.iloc[0, 1] = clean_cell(row_data.iloc[0, 1]).upper()
            row_data.iloc[0, 10] = normalize_definitive(row_data.iloc[0, 10])

            total = parse_int(row_data.iloc[0, 13])
            open_total = parse_int(row_data.iloc[0, 15])
            if total is not None:
                row_data.iloc[0, 13] = str(total)
            if open_total is not None:
                row_data.iloc[0, 15] = str(open_total)

            row_data.iloc[0, 12] = sanitize_tablet_codes(row_data.iloc[0, 12], total)
            row_data.iloc[0, 14] = sanitize_tablet_codes(row_data.iloc[0, 14], open_total)
            return row_data
        except Exception:
            return row_data

    def clean_open_tablets_when_closed(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Limpia Open_Tablets cuando el albarán está cerrado (función legacy, ahora manejada por fix_missing_open_column)"""
        try:
            if len(row_data.columns) < 15:
                return row_data

            definitive = str(row_data.iloc[0, 10]).strip()
            counted_date = str(row_data.iloc[0, 11]).strip()
            open_tablets = str(row_data.iloc[0, 14]).strip()

            if definitive in ['Yes', 'Ye', 'yes', 'ye', 'YES', 'YE']:
                if counted_date and counted_date not in ['', 'nan']:
                    if open_tablets and not re.search(r'[MALT]', open_tablets):
                        if open_tablets.isdigit() and int(open_tablets) <= 5:
                            row_data.iloc[0, 14] = ''

            return row_data
        except:
            return row_data

    def ensure_18_columns(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Asegura 18 columnas"""
        try:
            current_cols = len(row_data.columns)
            if current_cols < 18:
                for i in range(current_cols, 18):
                    row_data[i] = ''
            return row_data
        except:
            return row_data

    def fix_multiline_first_column(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """
        CRÍTICO: Detecta cuando col 0 tiene múltiples valores con saltos de línea
        Patrón: "FL\n612d\n729000018873" → debe separarse en cols 0, 1, 2
        
        VALIDACIÓN UNIVERSAL por LONGITUD:
        - Slip number: 12 dígitos (7290000XXXXX)
        - Warehouse: <= 10 caracteres (puede ser solo números o alfanumérico)
        """
        try:
            if len(row_data.columns) < 3:
                return row_data
            
            first_cell = str(row_data.iloc[0, 0]).strip()
            
            # Detectar si tiene saltos de línea Y contiene slip number
            if '\n' in first_cell and re.search(r'7290000\d{5}', first_cell):
                # Separar por saltos de línea
                lines = [line.strip() for line in first_cell.split('\n') if line.strip()]
                
                # Extraer valores específicos
                fl_value = 'FL'  # Default
                wh_value = ''
                slip_value = ''
                
                for line in lines:
                    # 1. Buscar estado (FL, DL, TX, etc.)
                    if line in ['FL', 'DL', 'TX', 'CA', 'NY', 'GA', 'NC', 'SC', 'VA']:
                        fl_value = line
                        continue
                    
                    # 2. Buscar slip number (SIEMPRE 12 dígitos)
                    slip_match = re.match(r'^(7290000\d{5})$', line)
                    if slip_match:
                        slip_value = slip_match.group(1)
                        continue
                    
                    # 3. Buscar warehouse (NO es slip Y tiene <= 10 caracteres)
                    if len(line) <= 10:
                        # Verificar que sea alfanumérico o tenga formato RO-XX
                        if re.match(r'^(RO-[A-Z]{2}|[\dA-Za-z]+)$', line, re.IGNORECASE):
                            wh_value = line.upper()
                            continue
                
                # Solo proceder si encontramos slip number
                if slip_value:
                    # Guardar TODOS los valores desde col 1 hasta col 17
                    saved_values = []
                    for col_idx in range(1, min(18, len(row_data.columns))):
                        saved_values.append(str(row_data.iloc[0, col_idx]))
                    
                    # Reconstruir fila correctamente
                    row_data.iloc[0, 0] = fl_value
                    row_data.iloc[0, 1] = wh_value if wh_value else '612D'  # Default si no encuentra
                    row_data.iloc[0, 2] = slip_value
                    
                    # Desplazar todos los valores guardados hacia la derecha
                    for i, val in enumerate(saved_values):
                        new_col = 3 + i
                        if new_col < len(row_data.columns):
                            row_data.iloc[0, new_col] = val
            
            return row_data
        
        except Exception as e:
            st.error(f"Error en fix_multiline_first_column: {e}")
            import traceback
            st.error(traceback.format_exc())
            return row_data
    
    
    def clean_warehouse_slip_column(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Separa warehouse code y slip number"""
        try:
            if len(row_data.columns) < 3:
                return row_data

            for col_idx in [1, 2, 3]:
                if col_idx >= len(row_data.columns):
                    continue

                cell_value = str(row_data.iloc[0, col_idx]).strip()
                pattern = r'^(RO-[A-Z]{2}|\d+[A-Za-z]*)\s+(7290000\d{5})'
                match = re.match(pattern, cell_value)

                if match:
                    warehouse_code = match.group(1).upper()
                    slip_number = match.group(2)

                    if col_idx == 1:
                        row_data.iloc[0, 1] = warehouse_code
                        if len(row_data.columns) > 2:
                            row_data.iloc[0, 2] = slip_number
                    elif col_idx == 2:
                        if str(row_data.iloc[0, 1]).strip() in ['', 'nan']:
                            row_data.iloc[0, 1] = warehouse_code
                        row_data.iloc[0, 2] = slip_number
                    elif col_idx == 3:
                        if str(row_data.iloc[0, 1]).strip() in ['', 'nan']:
                            row_data.iloc[0, 1] = warehouse_code
                        if str(row_data.iloc[0, 2]).strip() in ['', 'nan']:
                            row_data.iloc[0, 2] = slip_number
                    return row_data

            for col_idx in [1, 2]:
                if col_idx >= len(row_data.columns):
                    continue
                cell_value = str(row_data.iloc[0, col_idx])
                if re.search(r'(RO-[A-Za-z]{2}|\d+[A-Za-z]+)', cell_value, re.IGNORECASE):
                    row_data.iloc[0, col_idx] = cell_value.upper()

            return row_data
        except:
            return row_data

    def fix_customer_definitive_split(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Separa customer name de definitive"""
        try:
            if len(row_data.columns) < 11:
                return row_data

            for col_idx in [8, 9, 10]:
                if col_idx >= len(row_data.columns):
                    continue

                cell_value = str(row_data.iloc[0, col_idx]).strip()

                double_pattern = r'^(.+?)\s+(No|Yes|Ye)\s+(No|Yes|Ye)\s*$'
                match = re.search(double_pattern, cell_value)

                if match:
                    clean_text = match.group(1).strip()
                    first_definitive = match.group(2)
                    second_definitive = match.group(3)

                    row_data.iloc[0, col_idx] = clean_text + " " + first_definitive

                    definitive_col = 10
                    if definitive_col < len(row_data.columns):
                        definitive_cell = str(row_data.iloc[0, definitive_col]).strip()
                        if definitive_cell in ['', 'nan']:
                            row_data.iloc[0, definitive_col] = second_definitive
                    return row_data

                single_pattern = r'^(.+?)\s+(No|Yes|Ye)\s*$'
                match = re.search(single_pattern, cell_value)

                if match and col_idx in [8, 9]:
                    clean_text = match.group(1).strip()
                    definitive_value = match.group(2)

                    definitive_col = 10
                    if definitive_col < len(row_data.columns):
                        definitive_cell = str(row_data.iloc[0, definitive_col]).strip()
                        if definitive_cell in ['', 'nan']:
                            row_data.iloc[0, col_idx] = clean_text
                            row_data.iloc[0, definitive_col] = definitive_value
                            return row_data

            return row_data
        except:
            return row_data

    def fix_column_shift_after_definitive(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Corrige desplazamiento cuando Definitive=No"""
        try:
            if len(row_data.columns) < 14:
                return row_data

            definitive_cell = str(row_data.iloc[0, 10]).strip()
            counted_date_cell = str(row_data.iloc[0, 11]).strip()

            if definitive_cell in ['No', 'no', 'NO']:
                is_date = re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', counted_date_cell)

                if not is_date and counted_date_cell not in ['', 'nan']:
                    shift_values = []
                    for col_idx in range(11, min(18, len(row_data.columns))):
                        shift_values.append(str(row_data.iloc[0, col_idx]))

                    row_data.iloc[0, 11] = ''

                    for i, val in enumerate(shift_values):
                        new_col = 12 + i
                        if new_col < len(row_data.columns):
                            row_data.iloc[0, new_col] = val

            return row_data
        except:
            return row_data

    def fix_tablets_total_split(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Separa Total de Open"""
        try:
            if len(row_data.columns) < 16:
                return row_data

            total_cell = str(row_data.iloc[0, 13]).strip()
            pattern = r'^(\d+)\s+([\d\s,]+[MALT].*)$'
            match = re.match(pattern, total_cell)

            if match:
                total_number = match.group(1)
                open_tablets = match.group(2).strip()

                saved_values = []
                for col_idx in range(14, min(18, len(row_data.columns))):
                    saved_values.append(str(row_data.iloc[0, col_idx]))

                row_data.iloc[0, 13] = total_number
                row_data.iloc[0, 14] = open_tablets

                for i, val in enumerate(saved_values):
                    new_col = 15 + i
                    if new_col < len(row_data.columns):
                        row_data.iloc[0, new_col] = val

            return row_data
        except:
            return row_data

    # ========================================================================
    # MÉTODOS DE EXTRACCIÓN
    # ========================================================================

    @staticmethod
    def _column_index_for_x(x_center: float) -> int:
        """Columna 0..17 que contiene la coordenada X dada (con saturación en los extremos)."""
        if x_center < COL_BOUNDS[0][0]:
            return 0
        for idx, (x_start, x_end) in enumerate(COL_BOUNDS):
            if x_start <= x_center < x_end:
                return idx
        return len(COL_BOUNDS) - 1

    @staticmethod
    def _detect_header_limit(page) -> float:
        """
        'top' de la línea horizontal bajo el encabezado de la tabla.

        El encabezado cambia de alto entre la primera página y las siguientes,
        por lo que se lee la geometría real de la página; si no hay líneas se usa
        el valor fijo page.height - 497.
        """
        try:
            low, high = HEADER_LINE_RANGE
            tops = [ln['top'] for ln in page.lines
                    if abs(ln['bottom'] - ln['top']) < 1.0 and ln['x0'] < COL_BOUNDS[1][0]
                    and low <= ln['top'] <= high]
            if tops:
                return max(tops)
        except Exception:
            pass
        return page.height - HEADER_OFFSET

    @staticmethod
    def _cluster_lines(words: List[Dict], tolerance: float = 3.0) -> List[List[Dict]]:
        """Agrupa palabras en líneas visuales según su coordenada 'top'."""
        lines: List[List[Dict]] = []
        for w in sorted(words, key=lambda item: item['top']):
            if lines and w['top'] - lines[-1][0]['top'] <= tolerance:
                lines[-1].append(w)
            else:
                lines.append([w])
        return lines

    def _collect_row_words(self, body_words: List[Dict], sy: float,
                           next_sy: Optional[float]) -> List[List[Dict]]:
        """
        Líneas visuales que pertenecen a la fila del albarán que empieza en sy.

        - Ventana base: [sy - 2.5, sy + 11.0] (2 líneas), siempre acotada por la
          fila siguiente (next_sy - 2.5).
        - Si la fila ocupa 3 líneas, la ventana se extiende hasta sy + 18.0 solo
          mientras las líneas sean contiguas (<= 10 pt entre sí) y contengan
          únicamente columnas multilínea; así no se absorbe la fila de totales
          del pie (cols 13 y 15) ni texto ajeno a la fila.
        """
        window_bottom = sy + ROW_MAX_HEIGHT
        if next_sy is not None:
            window_bottom = min(window_bottom, next_sy - ROW_TOP_MARGIN)

        candidates = [w for w in body_words if sy - ROW_TOP_MARGIN <= w['top'] <= window_bottom]
        accepted: List[List[Dict]] = []

        for line in self._cluster_lines(candidates):
            line_top = min(w['top'] for w in line)
            if accepted:
                previous_top = min(w['top'] for w in accepted[-1])
                if line_top - previous_top > ROW_LINE_GAP:
                    break
                if line_top - sy > ROW_HEIGHT:
                    columns = {self._column_index_for_x((w['x0'] + w['x1']) / 2) for w in line}
                    if not columns.issubset(MULTILINE_COLS):
                        break
            accepted.append(line)

        return accepted

    def _extract_page_rows_by_coordinates(self, page, is_last_page: bool) -> Tuple[List[List[str]], Optional[Dict]]:
        """
        Extrae las filas de una página por coordenadas exactas.

        Cada albarán (palabra 7290000XXXXX en la col 2) define una ventana
        vertical que parte en sy - 2.5 y termina de forma dinámica (ver
        _collect_row_words): captura la segunda línea (+7.7 pt) de Customer,
        Job_Name, Definitive ("Ye" + "s"), Tablets y Open_Tablets, y se extiende
        hasta 18 pt cuando esas celdas ocupan 3 líneas.
        """
        words = page.extract_words(x_tolerance=2.5, y_tolerance=1.5, keep_blank_chars=False)

        header_limit = self._detect_header_limit(page)
        body_words = [w for w in words if w['top'] >= header_limit]

        def find_slip_tops(candidates):
            return sorted(
                w['top'] for w in candidates
                if self._column_index_for_x((w['x0'] + w['x1']) / 2) == 2
                and SLIP_PATTERN.match(w['text'])
            )

        slip_tops = find_slip_tops(body_words)
        if not slip_tops:
            # Geometría de página distinta a la esperada: no filtrar por encabezado.
            body_words = words
            header_limit = 0
            slip_tops = find_slip_tops(body_words)

        rows = []
        assigned_ids = set()
        for i, sy in enumerate(slip_tops):
            next_sy = slip_tops[i + 1] if i + 1 < len(slip_tops) else None

            cells = [[] for _ in COL_BOUNDS]
            for line_idx, line in enumerate(self._collect_row_words(body_words, sy, next_sy)):
                for w in line:
                    assigned_ids.add(id(w))
                    col_idx = self._column_index_for_x((w['x0'] + w['x1']) / 2)
                    cells[col_idx].append((line_idx, w['x0'], w['text']))

            row = []
            for col_idx, parts in enumerate(cells):
                parts.sort(key=lambda item: (item[0], item[1]))
                separator = '' if col_idx == 10 else ' '
                row.append(separator.join(text for _, _, text in parts).strip())
            rows.append(row)

        footer = None
        if is_last_page:
            last_sy = slip_tops[-1] if slip_tops else header_limit
            leftovers = [w for w in body_words if w['top'] > last_sy and id(w) not in assigned_ids]
            footer = self._detect_footer_totals(leftovers)

        return rows, footer

    def _detect_footer_totals(self, words: List[Dict]) -> Optional[Dict]:
        """Fila de totales globales del PDF (cols 13 y 15) entre las palabras no asignadas a filas."""
        candidates = sorted(words, key=lambda w: w['top'])

        lines: List[List[Dict]] = []
        for w in candidates:
            if lines and abs(w['top'] - lines[-1][0]['top']) <= 3.0:
                lines[-1].append(w)
            else:
                lines.append([w])

        for line in lines:
            col_13 = col_15 = None
            for w in line:
                if not re.fullmatch(r'\d+', w['text']):
                    continue
                col_idx = self._column_index_for_x((w['x0'] + w['x1']) / 2)
                if col_idx == 13 and col_13 is None:
                    col_13 = int(w['text'])
                elif col_idx == 15 and col_15 is None:
                    col_15 = int(w['text'])
            if col_13 is not None and col_15 is not None:
                return {'total_tablets': col_13, 'total_open': col_15}
        return None

    def method_coordinate_precision(self, pdf_path: str):
        """
        PRIORIDAD 1: extracción por coordenadas exactas con pdfplumber.

        El PDF de Alsina no tiene líneas verticales (lattice falla) y deja
        celdas vacías, por lo que stream sin límites fijos desplaza columnas.
        Aquí cada palabra se asigna a su columna por la geometría fija de las
        18 columnas, de modo que las celdas vacías nunca desplazan datos.
        """
        try:
            tables = []
            footer_totals = None

            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                for page_idx, page in enumerate(pdf.pages):
                    rows, page_footer = self._extract_page_rows_by_coordinates(
                        page, is_last_page=(page_idx == total_pages - 1))
                    if rows:
                        tables.append(CoordinateTable(pd.DataFrame(rows), page_idx + 1))
                    if page_footer:
                        footer_totals = page_footer

            if footer_totals:
                st.session_state['pdf_footer_totals'] = footer_totals

            return tables if tables else None
        except Exception:
            return None

    def _read_stream(self, pdf_path: str, **kwargs):
        """Camelot stream con las fronteras X exactas de las 18 columnas."""
        try:
            return camelot.read_pdf(pdf_path, pages='all', flavor='stream',
                                   columns=CAMELOT_STREAM_COLUMNS, **kwargs)
        except Exception:
            return camelot.read_pdf(pdf_path, pages='all', flavor='stream', **kwargs)

    def method_lattice_standard(self, pdf_path: str):
        try:
            return camelot.read_pdf(pdf_path, pages='all', flavor='lattice',
                                   process_background=True, line_scale=40)
        except:
            return None

    def method_stream_balanced(self, pdf_path: str):
        try:
            return self._read_stream(pdf_path, edge_tol=350, row_tol=12)
        except:
            return None

    def method_stream_standard(self, pdf_path: str):
        try:
            return self._read_stream(pdf_path)
        except:
            return None

    def method_stream_aggressive(self, pdf_path: str):
        try:
            return self._read_stream(pdf_path, edge_tol=500, row_tol=10,
                                     split_text=True, flag_size=True)
        except:
            return None

    def method_lattice_detailed(self, pdf_path: str):
        try:
            return camelot.read_pdf(pdf_path, pages='all', flavor='lattice',
                                   process_background=True, line_scale=40, iterations=2)
        except:
            return None

    def method_hybrid(self, pdf_path: str):
        all_tables = []
        try:
            stream_tables = self._read_stream(pdf_path, edge_tol=500)
            if stream_tables:
                all_tables.extend(stream_tables)
        except:
            pass
        try:
            lattice_tables = camelot.read_pdf(pdf_path, pages='all', flavor='lattice')
            if lattice_tables:
                all_tables.extend(lattice_tables)
        except:
            pass
        return all_tables if all_tables else None

    # ========================================================================
    # PROCESAMIENTO PRINCIPAL
    # ========================================================================

    def process_tables(self, tables) -> pd.DataFrame:
        """Procesa las tablas con todas las correcciones"""
        if not tables:
            return None

        all_data = []
        st.info(f"📄 PDF detectado con {len(tables)} páginas")

        for i, table in enumerate(tables):
            try:
                df = table.df
                st.write(f"📋 Procesando página {i + 1}: {df.shape}")

                if getattr(table, 'is_coordinate_exact', False):
                    for idx in df.index:
                        row_data = df.loc[idx:idx].copy()
                        row_data = self.ensure_18_columns(row_data)
                        row_data = self.normalize_row(row_data)
                        all_data.append(row_data)
                    continue

                df = self.merge_continuation_rows(df)

                for idx in df.index:
                    try:
                        row_text = ' '.join(str(cell) for cell in df.iloc[idx].values if pd.notna(cell))

                        if re.search(r'7290000\d{5}', row_text) and re.search(r'\b[A-Z]{2}\b', row_text):
                            if not any(skip in row_text for skip in ['Outstanding count', 'Page',
                                     'Return packing', 'Customer name', 'Alsina Forms']):
                                row_data = df.iloc[idx:idx+1].copy()

                                row_data = self.ensure_18_columns(row_data)
                                row_data = self.fix_multiline_first_column(row_data)
                                row_data = self.clean_warehouse_slip_column(row_data)
                                row_data = self.fix_missing_invoice_date1(row_data)
                                row_data = self.fix_customer_definitive_split(row_data)
                                row_data = self.fix_column_shift_after_definitive(row_data)
                                row_data = self.fix_tablets_total_split(row_data)
                                row_data = self.fix_missing_open_column(row_data)
                                row_data = self.clean_open_tablets_when_closed(row_data)
                                row_data = self.normalize_row(row_data)
                                all_data.append(row_data)
                    except:
                        continue
            except Exception as e:
                st.error(f"Error procesando página {i + 1}: {e}")
                continue

        if all_data:
            try:
                result = pd.concat(all_data, ignore_index=True)
                self.validate_simple(result)
                return result
            except Exception as e:
                st.error(f"Error combinando datos: {e}")
                return None
        return None

    def validate_simple(self, df: pd.DataFrame):
        """Validación simple"""
        if df is None or df.empty:
            st.error("❌ DataFrame vacío")
            return

        try:
            st.header("🔍 Validación del Sistema")
            total_rows = len(df)
            slip_count = sum(1 for idx in df.index
                           if re.search(r'7290000\d{5}',
                              ' '.join(str(c) for c in df.iloc[idx].values if pd.notna(c))))

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📊 Filas Totales", total_rows)
            with col2:
                st.metric("📊 Slips Válidos", slip_count)
            with col3:
                completeness = (slip_count / total_rows * 100) if total_rows > 0 else 0
                st.metric("📊 Completitud", f"{completeness:.1f}%")

            if completeness >= 95:
                st.success("🎉 **EXTRACCIÓN EXCELENTE**")
            elif completeness >= 80:
                st.info("📊 **EXTRACCIÓN BUENA**")
            else:
                st.warning("⚠️ **EXTRACCIÓN PARCIAL**")
        except Exception as e:
            st.error(f"Error en validación: {e}")

    def calculate_accuracy(self, tables) -> float:
        try:
            if not tables:
                return 0.0
            return sum(getattr(t, 'accuracy', 0) for t in tables) / len(tables)
        except:
            return 0.0

    def validate_extraction(self, df: pd.DataFrame) -> Dict:
        """Validación básica"""
        validation = {
            'total_rows': len(df) if df is not None else 0,
            'has_fl_column': False,
            'has_slip_numbers': False,
            'data_quality': 'unknown'
        }

        try:
            if df is None or df.empty:
                return validation

            if len(df.columns) > 0:
                first_col = df.iloc[:, 0].astype(str)
                validation['has_fl_column'] = any('FL' in val for val in first_col)

            if len(df.columns) > 2:
                for col_idx in range(min(5, len(df.columns))):
                    col_data = df.iloc[:, col_idx].astype(str)
                    if any(re.search(r'7290000\d{5}', val) for val in col_data):
                        validation['has_slip_numbers'] = True
                        break

            if validation['has_fl_column'] and validation['has_slip_numbers']:
                validation['data_quality'] = 'good'
            elif validation['has_slip_numbers']:
                validation['data_quality'] = 'acceptable'
            else:
                validation['data_quality'] = 'poor'
        except:
            validation['data_quality'] = 'error'

        return validation


# ============================================================================
# ANALIZADOR DE NEGOCIO
# ============================================================================

class BusinessAnalyzer:
    """Analizador de métricas de negocio"""

    def __init__(self):
        self.us_holidays = holidays.US(years=range(2024, 2027))

    def calculate_business_days(self, start_date_str: str, end_date_str: str) -> int:
        """Calcula días hábiles"""
        try:
            if not start_date_str or not end_date_str:
                return 0

            start_date = datetime.strptime(start_date_str, '%m/%d/%Y')
            end_date = datetime.strptime(end_date_str, '%m/%d/%Y')

            business_days = 0
            current_date = start_date

            while current_date <= end_date:
                if current_date.weekday() < 5:
                    if current_date.date() not in self.us_holidays:
                        business_days += 1
                current_date += timedelta(days=1)

            return business_days
        except:
            return 0

    def parse_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Procesa DataFrame para análisis (columnas con nombres y métricas derivadas)"""
        try:
            analysis_df = to_named_df(df).reset_index(drop=True)

            if 'Return_Slip' not in analysis_df.columns:
                return analysis_df

            analysis_df['Definitive'] = analysis_df['Definitive'].map(normalize_definitive)
            analysis_df['Return_Prefix'] = analysis_df['Return_Prefix'].map(clean_cell).str.upper()

            totals = row_tablet_totals(df.reset_index(drop=True))
            analysis_df['Total_Tablets'] = totals['total']
            analysis_df['Open_Total'] = totals['abiertas']
            analysis_df['Counting_Delay'] = pd.to_numeric(analysis_df['Counting_Delay'], errors='coerce')
            analysis_df['Validation_Delay'] = pd.to_numeric(analysis_df['Validation_Delay'], errors='coerce')

            def extract_slip(value):
                match = SLIP_SEARCH_PATTERN.search(clean_cell(value))
                return match.group(0) if match else ''

            analysis_df['slip_number'] = analysis_df['Return_Slip'].map(extract_slip)
            analysis_df['is_closed'] = (analysis_df['Definitive'] == 'Yes') & (analysis_df['Open_Total'] == 0)

            def extract_warehouse(value):
                match = re.search(r'(RO-[A-Z]{2}|\d+[A-Za-z]*)', clean_cell(value), re.IGNORECASE)
                return match.group(1).upper() if match else 'UNKNOWN'

            analysis_df['warehouse'] = analysis_df['Return_Prefix'].map(extract_warehouse)
            analysis_df['customer_name'] = analysis_df['Customer'].map(
                lambda v: clean_cell(v)[:50] if clean_cell(v) else 'Unknown')

            business_days = []
            for idx in analysis_df.index:
                counted_date = clean_cell(analysis_df.at[idx, 'Counted_Date'])
                return_date = clean_cell(analysis_df.at[idx, 'Return_Date'])
                if analysis_df.at[idx, 'is_closed'] and return_date and counted_date:
                    business_days.append(self.calculate_business_days(return_date, counted_date))
                else:
                    business_days.append(None)
            analysis_df['business_days_to_close'] = pd.to_numeric(
                pd.Series(business_days, index=analysis_df.index), errors='coerce')

            return analysis_df
        except Exception as e:
            st.error(f"Error procesando DataFrame: {e}")
            return df


# ============================================================================
# FUNCIONES DE ANÁLISIS DE TABLILLAS
# ============================================================================

def row_tablet_totals(df: pd.DataFrame) -> pd.DataFrame:
    """
    Totales por fila: 'total' (col 13 Total_Tablets) y 'abiertas' (col 15 Open_Total).
    Si la celda numérica no es válida se cuenta la lista de códigos (col 12 / col 14).
    """
    total = column_numeric(df, 13)
    total = total.where(total.notna(), column_text(df, 12).map(count_tablet_codes))
    abiertas = column_numeric(df, 15)
    abiertas = abiertas.where(abiertas.notna(), column_text(df, 14).map(count_tablet_codes))
    return pd.DataFrame({
        'total': total.fillna(0).astype(int),
        'abiertas': abiertas.fillna(0).astype(int)
    }, index=df.index)


def calculate_tablets_metrics(df: pd.DataFrame) -> Dict:
    """Calcula métricas globales de tablillas (Total_Tablets col 13, Open_Total col 15)"""
    try:
        totals = row_tablet_totals(df)
        total_tablillas = int(totals['total'].sum())
        tablillas_abiertas = int(totals['abiertas'].sum())
        tablillas_cerradas = total_tablillas - tablillas_abiertas

        return {
            'total': total_tablillas,
            'cerradas': tablillas_cerradas,
            'abiertas': tablillas_abiertas,
            'tasa_cierre': (tablillas_cerradas / total_tablillas * 100) if total_tablillas > 0 else 0
        }
    except Exception as e:
        st.error(f"Error en calculate_tablets_metrics: {e}")
        return {'total': 0, 'cerradas': 0, 'abiertas': 0, 'tasa_cierre': 0}


def _group_tablets(df: pd.DataFrame, keys: pd.Series, label: str) -> pd.DataFrame:
    """Agrupa totales/abiertas de tablillas por una clave (warehouse, cliente...)"""
    totals = row_tablet_totals(df)
    work = pd.DataFrame({label: keys.values, 'total': totals['total'].values,
                         'abiertas': totals['abiertas'].values})
    work = work[(work[label] != '') & (work['total'] > 0)]
    if work.empty:
        return pd.DataFrame()
    grouped = work.groupby(label, as_index=False)[['total', 'abiertas']].sum()
    grouped['cerradas'] = grouped['total'] - grouped['abiertas']
    grouped['tasa'] = (grouped['cerradas'] / grouped['total'] * 100).round(2)
    return grouped


def create_tablets_breakdown_by_warehouse(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown de tablillas por warehouse (Return_Prefix normalizado a mayúsculas)"""
    try:
        grouped = _group_tablets(df, column_text(df, 1).str.upper(), 'Warehouse')
        if grouped.empty:
            return pd.DataFrame()

        result = pd.DataFrame({
            'Warehouse': grouped['Warehouse'],
            'Total_Tablillas': grouped['total'],
            'Cerradas': grouped['cerradas'],
            'Abiertas': grouped['abiertas'],
            'Tasa_Cierre_%': grouped['tasa']
        })
        return result.sort_values('Total_Tablillas', ascending=False).reset_index(drop=True)
    except Exception as e:
        st.error(f"Error en create_tablets_breakdown_by_warehouse: {e}")
        return pd.DataFrame()


def create_tablets_by_customer(df: pd.DataFrame) -> pd.DataFrame:
    """Análisis de tablillas por cliente (top 10 por tablillas abiertas)"""
    try:
        grouped = _group_tablets(df, column_text(df, 8).str[:50], 'Cliente')
        if grouped.empty:
            return pd.DataFrame()

        result = pd.DataFrame({
            'Cliente': grouped['Cliente'],
            'Total': grouped['total'],
            'Abiertas': grouped['abiertas'],
            'Cerradas': grouped['cerradas'],
            'Tasa_Cierre_%': grouped['tasa']
        })
        return result.sort_values(['Abiertas', 'Total'], ascending=False).head(10).reset_index(drop=True)
    except Exception as e:
        st.error(f"Error en create_tablets_by_customer: {e}")
        return pd.DataFrame()


def build_cost_center_summary(analysis_df: pd.DataFrame) -> pd.DataFrame:
    """
    Resumen por Cost_Center a partir del DataFrame de BusinessAnalyzer.parse_dataframe:
    albaranes totales/cerrados/pendientes, tablillas totales/abiertas y demora promedio.
    """
    try:
        if analysis_df is None or analysis_df.empty or 'Cost_Center' not in analysis_df.columns:
            return pd.DataFrame()

        work = pd.DataFrame({
            'Cost Center': analysis_df['Cost_Center'].map(clean_cell).replace('', '(Sin CC)'),
            'closed': analysis_df['is_closed'].astype(bool),
            'Total Tablets': analysis_df['Total_Tablets'],
            'Open Tablets': analysis_df['Open_Total'],
            'Counting_Delay': analysis_df['Counting_Delay']
        })

        grouped = work.groupby('Cost Center').agg(
            **{'Albaranes Totales': ('closed', 'size'),
               'Cerrados': ('closed', 'sum'),
               'Total Tablets': ('Total Tablets', 'sum'),
               'Open Tablets': ('Open Tablets', 'sum'),
               'Promedio Counting Delay': ('Counting_Delay', 'mean')}
        ).reset_index()
        grouped['Pendientes'] = grouped['Albaranes Totales'] - grouped['Cerrados']
        grouped['Promedio Counting Delay'] = grouped['Promedio Counting Delay'].round(1)

        grouped = grouped[['Cost Center', 'Albaranes Totales', 'Cerrados', 'Pendientes',
                           'Total Tablets', 'Open Tablets', 'Promedio Counting Delay']]
        for col in ['Albaranes Totales', 'Cerrados', 'Pendientes', 'Total Tablets', 'Open Tablets']:
            grouped[col] = grouped[col].astype(int)
        return grouped.sort_values('Albaranes Totales', ascending=False).reset_index(drop=True)
    except Exception as e:
        st.error(f"Error en build_cost_center_summary: {e}")
        return pd.DataFrame()


def check_pdf_footer_totals(df: pd.DataFrame) -> Optional[Dict]:
    """Compara sumas de Total_Tablets (col 13) y Open_Total (col 15) contra el pie del PDF."""
    footer = st.session_state.get('pdf_footer_totals')
    if not footer or df is None:
        return None

    sum_total = int(column_numeric(df, 13).fillna(0).astype(int).sum())
    sum_open = int(column_numeric(df, 15).fillna(0).astype(int).sum())
    return {
        'pdf_total': footer['total_tablets'],
        'pdf_open': footer['total_open'],
        'extracted_total': sum_total,
        'extracted_open': sum_open,
        'match': sum_total == footer['total_tablets'] and sum_open == footer['total_open']
    }


def render_footer_validation(df: pd.DataFrame):
    """Banner de cuadratura contra el pie de página oficial del PDF (si fue detectado)."""
    check = check_pdf_footer_totals(df)
    if check is None:
        return

    if check['match']:
        st.success(f"✅ Cuadratura 100% exacta contra el pie de página del PDF: "
                   f"{check['pdf_total']:,} Total Tablets | {check['pdf_open']:,} Open Tablets")
    else:
        st.error(
            f"❌ La extracción no cuadra con el pie de página del PDF: "
            f"Total Tablets {check['extracted_total']:,} (PDF: {check['pdf_total']:,}) | "
            f"Open Tablets {check['extracted_open']:,} (PDF: {check['pdf_open']:,})")


def validate_tablets_integrity(df: pd.DataFrame, show_ui: bool = False,
                               full_df: Optional[pd.DataFrame] = None) -> List[Dict]:
    """
    Valida dos reglas de integridad por fila:
    - Tablillas Totales: códigos en Tablets (col 12) == Total_Tablets (col 13)
    - Tablillas Abiertas: códigos en Open_Tablets (col 14) == Open_Total (col 15)

    Con show_ui=True muestra además el banner de cuadratura contra el pie del PDF
    (sobre full_df si se entrega, ya que el pie corresponde al PDF completo).
    """
    discrepancies = []
    try:
        rules = [
            ('Tablillas Totales', 12, 13),
            ('Tablillas Abiertas', 14, 15),
        ]
        slips = column_text(df, 2)

        for rule_name, codes_idx, total_idx in rules:
            codes_count = column_text(df, codes_idx).map(count_tablet_codes)
            expected = column_text(df, total_idx).map(parse_int)

            for pos in range(len(df)):
                expected_value = expected.iloc[pos]
                found = int(codes_count.iloc[pos])

                if expected_value is None:
                    if found == 0:
                        continue
                elif expected_value == found:
                    continue

                discrepancies.append({
                    'Slip': slips.iloc[pos],
                    'Regla': rule_name,
                    'Esperado': expected_value,
                    'Encontrado': found,
                    'Diferencia': abs(found - (expected_value or 0))
                })
    except Exception:
        discrepancies = []

    if show_ui:
        render_footer_validation(full_df if full_df is not None else df)

    return discrepancies


# ============================================================================
# EXPORTACIÓN EXCEL PROFESIONAL CON MÚLTIPLES HOJAS
# ============================================================================

def _format_excel_sheet(ws, totals_row: bool = False, autofilter: bool = False):
    """Encabezados con estilo, anchos automáticos y (opcional) fila de totales resaltada."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    if ws.max_row < 1 or ws.max_column < 1:
        return

    header_fill = PatternFill('solid', start_color='1F4E78', end_color='1F4E78')
    header_font = Font(bold=True, color='FFFFFF')
    thin = Side(style='thin', color='BFBFBF')

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)

    for col_cells in ws.iter_cols(min_row=1, max_row=ws.max_row):
        longest = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max(longest + 2, 10), 60)

    ws.freeze_panes = 'A2'
    if autofilter and ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    if totals_row and ws.max_row > 1:
        total_fill = PatternFill('solid', start_color='D9E1F2', end_color='D9E1F2')
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)
            cell.fill = total_fill
            cell.border = Border(top=Side(style='medium', color='1F4E78'), bottom=thin)


def _with_totals_row(summary_df: pd.DataFrame, label_col: str, sum_cols: List[str],
                     overrides: Optional[Dict] = None) -> pd.DataFrame:
    """Agrega una fila 'TOTAL' con la suma de sum_cols y valores explícitos en overrides."""
    totals = {col: '' for col in summary_df.columns}
    totals[label_col] = 'TOTAL'
    for col in sum_cols:
        totals[col] = int(summary_df[col].sum())
    totals.update(overrides or {})
    return pd.concat([summary_df, pd.DataFrame([totals])], ignore_index=True)


def export_to_professional_excel(df: pd.DataFrame) -> io.BytesIO:
    """
    Crea Excel profesional con múltiples hojas:
    - Metadata
    - Datos_Principales (18 columnas; la lee create_historical_dashboard)
    - Resumen_Ejecutivo
    - Tablillas_Por_Warehouse
    - Top_Clientes_Tablillas
    - Por_Cost_Center
    - Discrepancias
    """
    try:
        buffer = io.BytesIO()
        totals_sheets = set()

        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:

            metadata = pd.DataFrame([{
                'Sistema': 'Camelot PDF Extractor Pro',
                'Versión': '3.2',
                'Fecha_Generación': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Total_Albaranes': len(df),
                'Empresa': 'Alsina Forms Co., Inc.'
            }])
            metadata.to_excel(writer, sheet_name='Metadata', index=False)

            # Sin fila de totales: el análisis histórico lee esta hoja fila a fila.
            export_df = df.copy()
            if len(export_df.columns) >= 18:
                export_df.columns = COLUMN_NAMES + list(export_df.columns[18:])
                for numeric_col in ['Total_Tablets', 'Open_Total', 'Counting_Delay', 'Validation_Delay']:
                    converted = pd.to_numeric(export_df[numeric_col], errors='coerce')
                    export_df[numeric_col] = converted.where(converted.notna(), export_df[numeric_col])
            export_df.to_excel(writer, sheet_name='Datos_Principales', index=False)

            metrics = calculate_tablets_metrics(df)
            summary_df = pd.DataFrame([{
                'Total_Tablillas': metrics['total'],
                'Tablillas_Cerradas': metrics['cerradas'],
                'Tablillas_Abiertas': metrics['abiertas'],
                'Tasa_Cierre_%': round(metrics['tasa_cierre'], 2),
                'Estado': 'EXCELENTE' if metrics['tasa_cierre'] >= 80 else 'BUENO' if metrics['tasa_cierre'] >= 70 else 'REQUIERE_ATENCION'
            }])
            summary_df.to_excel(writer, sheet_name='Resumen_Ejecutivo', index=False)

            warehouse_df = create_tablets_breakdown_by_warehouse(df)
            if not warehouse_df.empty:
                total_row = int(warehouse_df['Total_Tablillas'].sum())
                open_row = int(warehouse_df['Abiertas'].sum())
                warehouse_out = _with_totals_row(
                    warehouse_df, 'Warehouse', ['Total_Tablillas', 'Cerradas', 'Abiertas'],
                    {'Tasa_Cierre_%': round((total_row - open_row) / total_row * 100, 2) if total_row else 0})
                warehouse_out.to_excel(writer, sheet_name='Tablillas_Por_Warehouse', index=False)
                totals_sheets.add('Tablillas_Por_Warehouse')

            customer_df = create_tablets_by_customer(df)
            if not customer_df.empty:
                customer_df.to_excel(writer, sheet_name='Top_Clientes_Tablillas', index=False)

            analysis_df = BusinessAnalyzer().parse_dataframe(df)
            cost_center_df = build_cost_center_summary(analysis_df)
            if not cost_center_df.empty:
                overall_delay = pd.to_numeric(analysis_df['Counting_Delay'], errors='coerce').mean()
                cost_out = _with_totals_row(
                    cost_center_df, 'Cost Center',
                    ['Albaranes Totales', 'Cerrados', 'Pendientes', 'Total Tablets', 'Open Tablets'],
                    {'Promedio Counting Delay': round(float(overall_delay), 1) if pd.notna(overall_delay) else ''})
                cost_out.to_excel(writer, sheet_name='Por_Cost_Center', index=False)
                totals_sheets.add('Por_Cost_Center')

            discrepancies = validate_tablets_integrity(df)
            if discrepancies:
                disc_df = pd.DataFrame(discrepancies)
                disc_df.to_excel(writer, sheet_name='Discrepancias', index=False)

            for sheet_name, worksheet in writer.sheets.items():
                _format_excel_sheet(worksheet,
                                    totals_row=sheet_name in totals_sheets,
                                    autofilter=sheet_name == 'Datos_Principales')

        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"Error creando Excel profesional: {e}")
        return None


# ============================================================================
# DASHBOARD INTELIGENTE DE TABLILLAS
# ============================================================================

def create_tablets_dashboard(df: pd.DataFrame, full_df: Optional[pd.DataFrame] = None):
    """Dashboard completo e inteligente de tablillas (df ya filtrado; full_df = PDF completo)"""
    st.header("📦 Dashboard Inteligente de Tablillas")
    st.markdown("**Análisis completo del inventario de tablillas**")

    try:
        metrics = calculate_tablets_metrics(df)

        st.subheader("🎯 Métricas Globales")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Total Tablillas", f"{metrics['total']:,}",
                     help="Todas las tablillas en el sistema")
        with col2:
            st.metric("Tablillas Cerradas", f"{metrics['cerradas']:,}",
                     delta=f"{metrics['tasa_cierre']:.1f}%",
                     delta_color="normal",
                     help="Tablillas que ya fueron cerradas")
        with col3:
            st.metric("Tablillas Abiertas", f"{metrics['abiertas']:,}",
                     delta=f"{100-metrics['tasa_cierre']:.1f}%",
                     delta_color="inverse",
                     help="Tablillas pendientes de cierre")
        with col4:
            if metrics['tasa_cierre'] >= 80:
                st.metric("Estado", "✅ EXCELENTE",
                         help="Tasa de cierre superior al 80%")
            elif metrics['tasa_cierre'] >= 70:
                st.metric("Estado", "⚠️ BUENO",
                         help="Tasa de cierre entre 70-80%")
            else:
                st.metric("Estado", "🔴 ATENCIÓN",
                         help="Tasa de cierre inferior al 70%")

        st.subheader("📊 Distribución Visual")
        col1, col2 = st.columns(2)

        with col1:
            fig_pie = go.Figure(data=[go.Pie(
                labels=['Cerradas', 'Abiertas'],
                values=[metrics['cerradas'], metrics['abiertas']],
                hole=0.4,
                marker=dict(colors=['#28a745', '#dc3545']),
                textinfo='label+percent',
                textposition='auto'
            )])
            fig_pie.update_layout(
                title="Estado de Tablillas",
                annotations=[dict(
                    text=f"{metrics['total']:,}<br>Total",
                    x=0.5, y=0.5,
                    font_size=20,
                    showarrow=False
                )],
                showlegend=True
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        with col2:
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=metrics['tasa_cierre'],
                title={'text': "Tasa de Cierre (%)"},
                delta={'reference': 80, 'increasing': {'color': "green"}},
                gauge={
                    'axis': {'range': [None, 100]},
                    'bar': {'color': "#1f77b4"},
                    'steps': [
                        {'range': [0, 70], 'color': "#ffcccc"},
                        {'range': [70, 80], 'color': "#fff4cc"},
                        {'range': [80, 100], 'color': "#ccffcc"}
                    ],
                    'threshold': {
                        'line': {'color': "red", 'width': 4},
                        'thickness': 0.75,
                        'value': 80
                    }
                }
            ))
            st.plotly_chart(fig_gauge, use_container_width=True)

        st.subheader("🏭 Análisis por Warehouse")
        warehouse_df = create_tablets_breakdown_by_warehouse(df)

        if not warehouse_df.empty:
            col1, col2 = st.columns([2, 1])

            with col1:
                fig_bar = go.Figure()
                fig_bar.add_trace(go.Bar(
                    name='Cerradas',
                    x=warehouse_df['Warehouse'],
                    y=warehouse_df['Cerradas'],
                    marker_color='#28a745',
                    text=warehouse_df['Cerradas'],
                    textposition='inside'
                ))
                fig_bar.add_trace(go.Bar(
                    name='Abiertas',
                    x=warehouse_df['Warehouse'],
                    y=warehouse_df['Abiertas'],
                    marker_color='#dc3545',
                    text=warehouse_df['Abiertas'],
                    textposition='inside'
                ))
                fig_bar.update_layout(
                    title="Distribución por Warehouse",
                    barmode='stack',
                    xaxis_title="Warehouse",
                    yaxis_title="Cantidad de Tablillas",
                    showlegend=True,
                    hovermode='x unified'
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            with col2:
                st.markdown("**📋 Detalle por Warehouse**")
                st.dataframe(warehouse_df, use_container_width=True, height=300)

        st.subheader("🏢 Top Clientes con Tablillas Abiertas")
        customer_df = create_tablets_by_customer(df)

        if not customer_df.empty:
            fig_customers = px.bar(
                customer_df,
                x='Abiertas',
                y='Cliente',
                orientation='h',
                title="Top 10 Clientes - Tablillas Pendientes",
                color='Abiertas',
                color_continuous_scale='Reds',
                text='Abiertas'
            )
            fig_customers.update_traces(textposition='outside')
            fig_customers.update_layout(
                yaxis={'categoryorder': 'total ascending'},
                showlegend=False,
                height=400
            )
            st.plotly_chart(fig_customers, use_container_width=True)

            with st.expander("📊 Ver tabla detallada"):
                st.dataframe(customer_df, use_container_width=True)

        st.subheader("🔍 Validación de Integridad")
        discrepancies = validate_tablets_integrity(df, show_ui=True, full_df=full_df)

        if discrepancies:
            st.warning(f"⚠️ Se encontraron **{len(discrepancies)}** discrepancias de integridad")

            disc_df = pd.DataFrame(discrepancies)

            col1, col2 = st.columns([2, 1])
            with col1:
                st.dataframe(disc_df, use_container_width=True, height=300)
            with col2:
                st.info("""
                **Reglas de integridad por albarán:**

                - **Tablillas Totales**: cantidad de códigos en
                  `Tablets` = `Total_Tablets`
                - **Tablillas Abiertas**: cantidad de códigos en
                  `Open_Tablets` = `Open_Total`

                Puede indicar:
                - Códigos concatenados o truncados por el ERP
                - Errores de extracción del PDF
                """)
        else:
            st.success("✅ **Integridad perfecta**: los códigos de Tablets y Open_Tablets coinciden con Total_Tablets y Open_Total")

        st.subheader("⚠️ Alertas Inteligentes")

        alerts = []

        if metrics['tasa_cierre'] < 70:
            alerts.append({
                'tipo': '🔴 CRÍTICO',
                'mensaje': f"Tasa de cierre global muy baja: {metrics['tasa_cierre']:.1f}%",
                'acción': "Revisar procesos de cierre de tablillas"
            })
        elif metrics['tasa_cierre'] < 80:
            alerts.append({
                'tipo': '🟡 ADVERTENCIA',
                'mensaje': f"Tasa de cierre por debajo del objetivo: {metrics['tasa_cierre']:.1f}%",
                'acción': "Objetivo recomendado: >80%"
            })

        if not warehouse_df.empty:
            for _, row in warehouse_df.iterrows():
                tasa = float(row['Tasa_Cierre_%'])
                if tasa < 70:
                    alerts.append({
                        'tipo': '🔴 WAREHOUSE',
                        'mensaje': f"{row['Warehouse']}: Tasa de cierre {tasa:.1f}%",
                        'acción': f"Revisar {row['Abiertas']} tablillas abiertas"
                    })

        if not customer_df.empty:
            top_cliente = customer_df.iloc[0]
            if top_cliente['Abiertas'] > 10:
                alerts.append({
                    'tipo': '🟡 CLIENTE',
                    'mensaje': f"{top_cliente['Cliente']}: {top_cliente['Abiertas']} tablillas abiertas",
                    'acción': "Contactar para coordinación de cierre"
                })

        if alerts:
            for alert in alerts:
                st.warning(f"**{alert['tipo']}**: {alert['mensaje']}  \n💡 *{alert['acción']}*")
        else:
            st.success("✅ **Sin alertas**: Todos los indicadores dentro de parámetros normales")

    except Exception as e:
        st.error(f"Error creando dashboard de tablillas: {e}")
        import traceback
        st.error(traceback.format_exc())


# ============================================================================
# DASHBOARD DE ALBARANES
# ============================================================================

def create_analysis_dashboard(df: pd.DataFrame):
    """Dashboard de análisis de albaranes"""
    st.header("📊 Dashboard Ejecutivo - Análisis de Albaranes")

    try:
        analyzer = BusinessAnalyzer()

        with st.spinner("Procesando datos para análisis..."):
            analysis_df = analyzer.parse_dataframe(df)

        total_albaranes = len(analysis_df)
        closed_albaranes = len(analysis_df[analysis_df['is_closed'] == True])
        pending_albaranes = total_albaranes - closed_albaranes
        closure_rate = (closed_albaranes / total_albaranes * 100) if total_albaranes > 0 else 0

        st.subheader("🎯 KPIs Principales")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Total Albaranes", total_albaranes)
        with col2:
            st.metric("Albaranes Cerrados", closed_albaranes)
        with col3:
            st.metric("Albaranes Pendientes", pending_albaranes)
        with col4:
            st.metric("Tasa de Cierre", f"{closure_rate:.1f}%")

        avg_counting_delay = pd.to_numeric(analysis_df['Counting_Delay'], errors='coerce').mean()
        avg_validation_delay = pd.to_numeric(analysis_df['Validation_Delay'], errors='coerce').mean()
        critical_albaranes = int((pd.to_numeric(analysis_df['Counting_Delay'], errors='coerce')
                                  >= CRITICAL_DELAY_DAYS).sum())

        st.subheader("⏳ Demoras de Conteo y Validación")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Promedio Counting Delay (días)",
                      f"{avg_counting_delay:.1f}" if pd.notna(avg_counting_delay) else "N/D")
        with col2:
            st.metric("Promedio Validation Delay (días)",
                      f"{avg_validation_delay:.1f}" if pd.notna(avg_validation_delay) else "N/D")
        with col3:
            st.metric("Albaranes Críticos", critical_albaranes,
                      help=f"Albaranes con Counting Delay ≥ {CRITICAL_DELAY_DAYS} días")

        st.subheader("🏭 Performance por Warehouse")

        warehouse_stats = []
        for wh in analysis_df['warehouse'].unique():
            if wh != 'UNKNOWN':
                wh_data = analysis_df[analysis_df['warehouse'] == wh]
                wh_total = len(wh_data)
                wh_closed = len(wh_data[wh_data['is_closed'] == True])
                wh_rate = (wh_closed / wh_total * 100) if wh_total > 0 else 0

                closed_data = wh_data[wh_data['is_closed'] == True]
                avg_days = closed_data['business_days_to_close'].mean() if len(closed_data) > 0 else 0

                warehouse_stats.append({
                    'Warehouse': wh,
                    'Total': wh_total,
                    'Cerrados': wh_closed,
                    'Pendientes': wh_total - wh_closed,
                    'Tasa Cierre': f"{wh_rate:.1f}%",
                    'Días Promedio': f"{avg_days:.1f}" if avg_days > 0 else "N/A"
                })

        if warehouse_stats:
            warehouse_df = pd.DataFrame(warehouse_stats)
            st.dataframe(warehouse_df, use_container_width=True)

            fig = px.bar(
                warehouse_df,
                x='Warehouse',
                y=['Cerrados', 'Pendientes'],
                title="Distribución de Albaranes por Warehouse",
                color_discrete_map={'Cerrados': '#28a745', 'Pendientes': '#dc3545'}
            )
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("🏗️ Performance por Cost Center")
        cost_center_df = build_cost_center_summary(analysis_df)
        if not cost_center_df.empty:
            st.dataframe(
                cost_center_df, use_container_width=True, hide_index=True,
                column_config={
                    'Promedio Counting Delay': st.column_config.NumberColumn(
                        'Promedio Counting Delay', format="%.1f")
                })

            fig_cc = make_subplots(specs=[[{"secondary_y": True}]])
            bar_series = [('Albaranes Totales', '#1f77b4'), ('Cerrados', '#28a745'),
                          ('Pendientes', '#dc3545'), ('Total Tablets', '#9467bd'),
                          ('Open Tablets', '#ff7f0e')]
            for column, color in bar_series:
                fig_cc.add_trace(go.Bar(name=column, x=cost_center_df['Cost Center'],
                                        y=cost_center_df[column], marker_color=color),
                                 secondary_y=False)
            fig_cc.add_trace(go.Scatter(
                name='Promedio Counting Delay (días)', x=cost_center_df['Cost Center'],
                y=cost_center_df['Promedio Counting Delay'], mode='lines+markers',
                line=dict(color='#333333', width=2), marker=dict(size=8)),
                secondary_y=True)
            fig_cc.update_layout(title="Albaranes, Tablillas y Demora por Cost Center",
                                 barmode='group', hovermode='x unified', height=500)
            fig_cc.update_yaxes(title_text="Cantidad", secondary_y=False)
            fig_cc.update_yaxes(title_text="Counting Delay (días)", secondary_y=True)
            st.plotly_chart(fig_cc, use_container_width=True)
        else:
            st.info("No hay datos de Cost Center para mostrar")

        st.subheader("⏱️ Análisis de Tiempos de Cierre")

        closed_data = analysis_df[analysis_df['is_closed'] == True]
        if len(closed_data) > 0:
            valid_days = closed_data.dropna(subset=['business_days_to_close'])

            if len(valid_days) > 0:
                avg_days = valid_days['business_days_to_close'].mean()
                median_days = valid_days['business_days_to_close'].median()
                max_days = valid_days['business_days_to_close'].max()

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Días Promedio", f"{avg_days:.1f}")
                with col2:
                    st.metric("Días Mediana", f"{median_days:.1f}")
                with col3:
                    st.metric("Máximo Días", f"{max_days:.0f}")

                fig = px.histogram(
                    valid_days,
                    x='business_days_to_close',
                    nbins=15,
                    title="Distribución de Días Hábiles para Cierre",
                    labels={'business_days_to_close': 'Días Hábiles', 'count': 'Cantidad'}
                )
                st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"Error creando dashboard: {e}")


# ============================================================================
# DASHBOARD HISTÓRICO MEJORADO
# ============================================================================

def create_historical_dashboard():
    """Dashboard de análisis histórico MEJORADO con tablillas"""
    st.header("📈 Dashboard Histórico - Análisis Comparativo")
    st.info("📁 Carga múltiples archivos Excel para análisis de tendencias")

    uploaded_files = st.file_uploader(
        "Selecciona archivos Excel",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Archivos Excel generados por la app"
    )

    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} archivos cargados")

        all_data = []
        file_info = []

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name='Datos_Principales')
                filename = uploaded_file.name
                date_match = re.search(r'(\d{8})', filename)

                if date_match:
                    file_date = datetime.strptime(date_match.group(1), '%Y%m%d')
                else:
                    file_date = datetime.now()

                df['fecha_archivo'] = file_date
                df['nombre_archivo'] = filename
                all_data.append(df)

                file_info.append({
                    'Archivo': filename,
                    'Fecha': file_date.strftime('%Y-%m-%d'),
                    'Filas': len(df)
                })
            except Exception as e:
                st.error(f"Error leyendo {uploaded_file.name}: {e}")

        if all_data:
            combined_df = pd.concat(all_data, ignore_index=True)

            # ============================================================
            # ANÁLISIS DE TABLILLAS HISTÓRICO
            # ============================================================

            st.subheader("📦 Evolución de Tablillas en el Tiempo")

            tablets_daily = []
            for fecha in sorted(combined_df['fecha_archivo'].unique()):
                df_fecha = combined_df[combined_df['fecha_archivo'] == fecha]

                # Asegurar que usamos solo las primeras 18 columnas
                df_fecha_original = df_fecha.iloc[:, :18].copy()
                tablets_metrics = calculate_tablets_metrics(df_fecha_original)

                tablets_daily.append({
                    'Fecha': fecha.strftime('%Y-%m-%d'),
                    'Total': tablets_metrics['total'],
                    'Cerradas': tablets_metrics['cerradas'],
                    'Abiertas': tablets_metrics['abiertas'],
                    'Tasa_Cierre': tablets_metrics['tasa_cierre']
                })

            tablets_df = pd.DataFrame(tablets_daily)

            # Gráfico de evolución de tablillas
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=tablets_df['Fecha'],
                y=tablets_df['Total'],
                mode='lines+markers',
                name='Total',
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=8)
            ))
            fig.add_trace(go.Scatter(
                x=tablets_df['Fecha'],
                y=tablets_df['Cerradas'],
                mode='lines+markers',
                name='Cerradas',
                line=dict(color='#28a745', width=3),
                marker=dict(size=8),
                fill='tonexty'
            ))
            fig.add_trace(go.Scatter(
                x=tablets_df['Fecha'],
                y=tablets_df['Abiertas'],
                mode='lines+markers',
                name='Abiertas',
                line=dict(color='#dc3545', width=3),
                marker=dict(size=8)
            ))
            fig.update_layout(
                title="Evolución de Tablillas - Total, Cerradas y Abiertas",
                xaxis_title="Fecha",
                yaxis_title="Cantidad de Tablillas",
                hovermode='x unified',
                height=500
            )
            st.plotly_chart(fig, use_container_width=True)

            # Gráfico de tasa de cierre
            fig_tasa = go.Figure()
            fig_tasa.add_trace(go.Scatter(
                x=tablets_df['Fecha'],
                y=tablets_df['Tasa_Cierre'],
                mode='lines+markers',
                name='Tasa de Cierre',
                line=dict(color='#ff7f0e', width=3),
                marker=dict(size=8),
                fill='tozeroy'
            ))
            fig_tasa.add_hline(
                y=80,
                line_dash="dash",
                line_color="green",
                annotation_text="Objetivo: 80%",
                annotation_position="right"
            )
            fig_tasa.update_layout(
                title="Tasa de Cierre de Tablillas Histórica (%)",
                xaxis_title="Fecha",
                yaxis_title="Tasa de Cierre (%)",
                hovermode='x unified',
                height=400
            )
            st.plotly_chart(fig_tasa, use_container_width=True)

            # ============================================================
            # ANÁLISIS POR WAREHOUSE HISTÓRICO
            # ============================================================

            st.subheader("🏭 Evolución por Warehouse")

            warehouse_daily = {}
            for fecha in sorted(combined_df['fecha_archivo'].unique()):
                df_fecha = combined_df[combined_df['fecha_archivo'] == fecha]
                df_fecha_original = df_fecha.iloc[:, :18].copy()

                warehouse_breakdown = create_tablets_breakdown_by_warehouse(df_fecha_original)

                for _, row in warehouse_breakdown.iterrows():
                    wh = row['Warehouse']
                    if wh not in warehouse_daily:
                        warehouse_daily[wh] = []

                    warehouse_daily[wh].append({
                        'Fecha': fecha.strftime('%Y-%m-%d'),
                        'Total': row['Total_Tablillas'],
                        'Abiertas': row['Abiertas'],
                        'Cerradas': row['Cerradas']
                    })

            # Crear gráfico por warehouse
            if warehouse_daily:
                fig_wh = go.Figure()

                for wh, data in warehouse_daily.items():
                    wh_df = pd.DataFrame(data)
                    fig_wh.add_trace(go.Scatter(
                        x=wh_df['Fecha'],
                        y=wh_df['Abiertas'],
                        mode='lines+markers',
                        name=wh,
                        line=dict(width=2),
                        marker=dict(size=6)
                    ))

                fig_wh.update_layout(
                    title="Evolución de Tablillas Abiertas por Warehouse",
                    xaxis_title="Fecha",
                    yaxis_title="Tablillas Abiertas",
                    hovermode='x unified',
                    height=500
                )
                st.plotly_chart(fig_wh, use_container_width=True)

            # ============================================================
            # TABLA RESUMEN
            # ============================================================

            st.subheader("📋 Resumen Histórico")
            st.dataframe(tablets_df, use_container_width=True)

            # ============================================================
            # EXPORTACIÓN
            # ============================================================

            st.subheader("💾 Exportar Consolidado")

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Hoja 1: Datos consolidados
                combined_df.to_excel(writer, sheet_name='Datos_Consolidados', index=False)

                # Hoja 2: Evolución tablillas
                tablets_df.to_excel(writer, sheet_name='Evolucion_Tablillas', index=False)

                # Hoja 3: Por warehouse (último período)
                if not combined_df.empty:
                    last_date = combined_df['fecha_archivo'].max()
                    last_df = combined_df[combined_df['fecha_archivo'] == last_date].iloc[:, :18]
                    last_warehouse = create_tablets_breakdown_by_warehouse(last_df)
                    if not last_warehouse.empty:
                        last_warehouse.to_excel(writer, sheet_name='Ultimo_Por_Warehouse', index=False)

            st.download_button(
                "📊 Descargar Análisis Histórico Consolidado",
                buffer.getvalue(),
                f"historico_consolidado_{datetime.now().strftime('%Y%m%d')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("👆 Sube archivos Excel para comenzar")


# ============================================================================
# FILTROS GLOBALES Y TABLA INTERACTIVA
# ============================================================================

FILTER_KEYS = {
    'Subsidiary (Wh)': 'filter_subsidiary',
    'Warehouse': 'filter_warehouse',
    'Cost Center': 'filter_cost_center',
    'Estado (Definitive)': 'filter_definitive',
}
DEFINITIVE_ALL = 'Todos'


def reset_global_filters():
    """Callback del botón 'Limpiar filtros'."""
    for key in FILTER_KEYS.values():
        st.session_state[key] = DEFINITIVE_ALL if key == 'filter_definitive' else []


def render_global_filters(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filtros globales en el sidebar (Subsidiary, Warehouse, Cost_Center y Definitive).
    Un multiselect vacío equivale a 'todos'. Devuelve el DataFrame filtrado.
    """
    filter_columns = [
        ('Subsidiary (Wh)', 0, True),
        ('Warehouse', 1, True),
        ('Cost Center', 5, False),
    ]

    mask = pd.Series(True, index=df.index)

    with st.sidebar:
        st.divider()
        st.header("🔎 Filtros Globales")

        for label, col_idx, to_upper in filter_columns:
            values = column_text(df, col_idx)
            if to_upper:
                values = values.str.upper()
            options = sorted(v for v in values.unique() if v)

            key = FILTER_KEYS[label]
            st.session_state[key] = [v for v in st.session_state.get(key, []) if v in options]
            selected = st.multiselect(label, options, key=key,
                                      placeholder="Todos")
            if selected:
                mask &= values.isin(selected)

        definitive_key = FILTER_KEYS['Estado (Definitive)']
        if st.session_state.get(definitive_key) not in (DEFINITIVE_ALL, 'Yes', 'No'):
            st.session_state[definitive_key] = DEFINITIVE_ALL
        definitive_choice = st.radio("Estado (Definitive)", [DEFINITIVE_ALL, 'Yes', 'No'],
                                     key=definitive_key, horizontal=True)
        if definitive_choice != DEFINITIVE_ALL:
            mask &= column_text(df, 10).map(normalize_definitive) == definitive_choice

        st.button("🧹 Limpiar filtros", on_click=reset_global_filters)

        filtered = df[mask]
        st.caption(f"Mostrando **{len(filtered)}** de **{len(df)}** albaranes")

    return filtered


def render_interactive_table(df: pd.DataFrame):
    """Tabla interactiva con column_config y resaltado por Styler."""
    st.subheader("🗂️ Tabla Interactiva de Albaranes")

    if df is None or df.empty:
        st.info("No hay albaranes que coincidan con los filtros seleccionados")
        return

    display = to_named_df(df).reset_index(drop=True)
    if 'Open_Total' not in display.columns:
        st.dataframe(display, use_container_width=True, height=450)
        return

    numeric_columns = ['Total_Tablets', 'Open_Total', 'Counting_Delay', 'Validation_Delay']
    for column in numeric_columns:
        display[column] = pd.to_numeric(display[column], errors='coerce')

    def highlight_rows(row):
        is_pending = row['Definitive'] == 'No' or (pd.notna(row['Open_Total']) and row['Open_Total'] > 0)
        is_critical = pd.notna(row['Counting_Delay']) and row['Counting_Delay'] >= CRITICAL_DELAY_DAYS

        if is_pending:
            style = 'background-color: #f8d7da; color: #212529'
        elif is_critical:
            style = 'background-color: #fff3cd; color: #212529'
        else:
            style = ''

        styles = [style] * len(row)
        if is_critical:
            styles[list(row.index).index('Counting_Delay')] = 'background-color: #ffe08a; color: #212529'
        return styles

    st.dataframe(
        display.style.apply(highlight_rows, axis=1),
        use_container_width=True,
        height=450,
        hide_index=True,
        column_config={
            'Total_Tablets': st.column_config.NumberColumn(
                'Total_Tablets', help="Cantidad total de tablillas del albarán", format="%d"),
            'Open_Total': st.column_config.NumberColumn(
                'Open_Total', help="Cantidad de tablillas abiertas (0 si el albarán está cerrado)", format="%d"),
            'Counting_Delay': st.column_config.NumberColumn(
                'Counting_Delay', help="Días de demora de conteo", format="%d"),
            'Validation_Delay': st.column_config.NumberColumn(
                'Validation_Delay', help="Días de demora de validación", format="%d"),
        }
    )
    st.caption(f"🟥 Definitive = No o Open_Total > 0  |  🟨 Counting_Delay ≥ {CRITICAL_DELAY_DAYS} días")


# ============================================================================
# APLICACIÓN PRINCIPAL
# ============================================================================

def run_extraction_cached(uploaded_file, extractor: 'CamelotExtractorPro') -> Dict:
    """Extrae el PDF una sola vez por archivo (los filtros provocan reruns de Streamlit)."""
    file_bytes = uploaded_file.getvalue()
    file_key = hashlib.md5(file_bytes).hexdigest()

    if st.session_state.get('extraction_key') != file_key:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(file_bytes)
            tmp_path = tmp_file.name

        try:
            st.header("📄 Ejecutando Extracción")
            results = extractor.extract_with_all_methods(tmp_path)
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

        st.session_state['extraction_results'] = results
        st.session_state['extraction_key'] = file_key

    return st.session_state['extraction_results']


def main():
    render_header()

    main_tabs = st.tabs([
        "📄 Extracción PDF",
        "📊 Análisis de Albaranes",
        "📦 Dashboard de Tablillas",
        "📈 Análisis Histórico"
    ])

    with main_tabs[0]:
        with st.sidebar:
            st.header("📊 Información del Sistema")
            st.info("""
            **Versión 3.2**

            ✨ **Características**:
            - ✅ Extracción por coordenadas exactas (18 columnas)
            - ✅ Cuadratura contra el pie de página del PDF
            - ✅ Validación de integridad Total/Open
            - ✅ Filtros globales en vivo
            - ✅ Demoras y Cost Center
            - ✅ Excel con múltiples hojas
            - ✅ Análisis histórico avanzado
            """)

            st.divider()

            st.markdown("**🔧 Opciones**")
            show_debug = st.checkbox("Modo Debug", value=False)

        uploaded_file = st.file_uploader(
            "📂 Selecciona el PDF",
            type=['pdf'],
            help="Reporte Outstanding Count Returns"
        )

        if uploaded_file:
            extractor = CamelotExtractorPro()
            results = run_extraction_cached(uploaded_file, extractor)

            st.header("📊 Resultados de Extracción")
            method_names = list(results.keys())

            if method_names:
                tabs = st.tabs(method_names)
                best_method = None
                best_score = 0

                for tab, method_name in zip(tabs, method_names):
                    with tab:
                        result = results[method_name]

                        if result['success']:
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Tablas", result.get('tables_found', 0))
                            with col2:
                                st.metric("Filas", result.get('rows', 0))
                            with col3:
                                st.metric("Precisión", f"{result.get('accuracy', 0):.1f}%")

                            if result.get('data') is not None and len(result['data']) > 0:
                                validation = extractor.validate_extraction(result['data'])

                                score = validation['total_rows']
                                if validation['has_fl_column']:
                                    score += 10
                                if validation['has_slip_numbers']:
                                    score += 10
                                if method_name == 'method_coordinate_precision':
                                    # Geometría exacta: gana sobre los respaldos de Camelot
                                    # (hybrid puede duplicar filas y sumar más puntaje).
                                    score += 100000

                                if score > best_score:
                                    best_score = score
                                    best_method = method_name

                                st.dataframe(result['data'], use_container_width=True, height=400)
                        elif result.get('error') and show_debug:
                            st.error(result['error'])

                if best_method:
                    st.header("🏆 Mejor Método de Extracción")
                    st.success(f"**{best_method}**")

                    best_data = results[best_method]['data']
                    st.session_state['extracted_data'] = best_data

                    render_footer_validation(best_data)

                    st.subheader("💾 Exportar Datos")
                    col1, col2 = st.columns(2)

                    with col1:
                        try:
                            csv = best_data.to_csv(index=False)
                            st.download_button(
                                "📄 Descargar CSV Simple",
                                csv,
                                f"data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                                "text/csv",
                                help="Archivo CSV básico con datos extraídos"
                            )
                        except Exception as e:
                            st.error(f"Error generando CSV: {e}")

                    with col2:
                        try:
                            excel_buffer = export_to_professional_excel(best_data)
                            if excel_buffer:
                                st.download_button(
                                    "📊 Descargar Excel Profesional",
                                    excel_buffer.getvalue(),
                                    f"analisis_completo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    help="Excel con múltiples hojas de análisis"
                                )
                        except Exception as e:
                            st.error(f"Error generando Excel: {e}")

    extracted_data = st.session_state.get('extracted_data')
    has_data = extracted_data is not None and len(extracted_data) > 0
    filtered_data = render_global_filters(extracted_data) if has_data else None

    if has_data:
        with main_tabs[0]:
            render_interactive_table(filtered_data)

    with main_tabs[1]:
        if has_data:
            if filtered_data.empty:
                st.warning("No hay albaranes que coincidan con los filtros seleccionados")
            else:
                create_analysis_dashboard(filtered_data)
        else:
            st.info("💡 Primero extrae datos del PDF en la pestaña 'Extracción PDF'")

    with main_tabs[2]:
        if has_data:
            if filtered_data.empty:
                st.warning("No hay albaranes que coincidan con los filtros seleccionados")
            else:
                create_tablets_dashboard(filtered_data, full_df=extracted_data)
        else:
            st.info("💡 Primero extrae datos del PDF en la pestaña 'Extracción PDF'")

    with main_tabs[3]:
        create_historical_dashboard()


if __name__ == "__main__":
    main()
