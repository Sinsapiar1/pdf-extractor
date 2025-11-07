# camelot_extractor_pro_final_v3.py
"""
Extractor Profesional de PDFs con Camelot + Dashboard Inteligente
VERSIÓN 3.1 - Corrección para columna Open vacía (último día de cierre)
"""

import streamlit as st
import pandas as pd
import camelot
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
    page_title="Camelot PDF Extractor Pro v3.0",
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
                Sistema Profesional de Extracción y Análisis de Albaranes v3.0
            </h4>
        </div>
    """, unsafe_allow_html=True)

# ============================================================================
# CLASE PRINCIPAL: EXTRACTOR
# ============================================================================

class CamelotExtractorPro:
    """
    Extractor especializado - versión profesional con 8 correcciones universales
    
    Versión 3.1 - Incluye corrección crítica para columna Open vacía cuando 
    todas las tablillas están cerradas (último día de cierre de mes)
    """

    def __init__(self):
        self.extraction_methods = [
            self.method_stream_standard,       # PRIORIDAD 1: Funciona mejor con tablillas cerradas
            self.method_stream_balanced,       # PRIORIDAD 2
            self.method_lattice_standard,      # PRIORIDAD 3
            self.method_stream_aggressive,
            self.method_lattice_detailed,
            self.method_hybrid
        ]

    def extract_with_all_methods(self, pdf_path: str) -> Dict:
        """Prueba todos los métodos y compara resultados"""
        results = {}

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

    def fix_missing_open_column(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """
        CRÍTICO: Corrige desplazamiento cuando columna Open está completamente vacía.
        
        Problema: Cuando todas las tablillas están cerradas, la columna Open (col 14) 
        está vacía en el PDF. Camelot NO detecta columnas vacías, causando que todas 
        las columnas posteriores se desplacen una posición a la izquierda.
        
        Detección:
        - Definitive = "Yes" o "Ye" (albarán cerrado)
        - Counted_Date existe y no está vacía
        - Col 14 NO tiene códigos con sufijos [MALT]
        - Col 14 parece ser un número simple (probablemente Tablets_Total desplazado)
        
        Solución:
        - Insertar columna vacía en posición 14 (Open)
        - Desplazar todo desde col 14 hasta col 17 hacia la derecha
        """
        try:
            if len(row_data.columns) < 18:
                return row_data

            definitive = str(row_data.iloc[0, 10]).strip()
            counted_date = str(row_data.iloc[0, 11]).strip()
            tablets = str(row_data.iloc[0, 12]).strip()
            total = str(row_data.iloc[0, 13]).strip()
            col_14 = str(row_data.iloc[0, 14]).strip()
            
            # Verificar si el albarán está cerrado
            is_closed = definitive in ['Yes', 'Ye', 'yes', 'ye', 'YES', 'YE'] and \
                       counted_date and counted_date not in ['', 'nan', 'None']
            
            if is_closed:
                # Verificar si col 14 NO tiene códigos [MALT] (indicador de Open)
                has_malt_codes = bool(re.search(r'\d+[MALT]', col_14))
                
                # Verificar si col 14 parece ser un número simple (probablemente Tablets_Total desplazado)
                is_simple_number = col_14.isdigit() and len(col_14) <= 3
                
                # Si col 14 NO tiene códigos [MALT] Y parece ser un número simple
                if not has_malt_codes and is_simple_number:
                    # Guardar valores desde col 14 hasta col 17
                    saved_values = []
                    for col_idx in range(14, min(18, len(row_data.columns))):
                        saved_values.append(str(row_data.iloc[0, col_idx]))
                    
                    # Col 14 (Open) debe estar vacía cuando está cerrado
                    row_data.iloc[0, 14] = ''
                    
                    # Desplazar valores guardados hacia la derecha
                    for i, val in enumerate(saved_values):
                        new_col = 15 + i
                        if new_col < len(row_data.columns):
                            row_data.iloc[0, new_col] = val
                
                # Caso adicional: Si col 14 tiene un número pequeño sin [MALT], limpiarlo
                elif col_14 and not has_malt_codes:
                    if col_14.isdigit() and int(col_14) <= 5:
                        row_data.iloc[0, 14] = ''

            return row_data
        except Exception as e:
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

    def method_lattice_standard(self, pdf_path: str):
        try:
            return camelot.read_pdf(pdf_path, pages='all', flavor='lattice',
                                   process_background=True, line_scale=40)
        except:
            return None

    def method_stream_balanced(self, pdf_path: str):
        try:
            return camelot.read_pdf(pdf_path, pages='all', flavor='stream',
                                   edge_tol=350, row_tol=12, column_tol=5)
        except:
            return None

    def method_stream_standard(self, pdf_path: str):
        try:
            return camelot.read_pdf(pdf_path, pages='all', flavor='stream')
        except:
            return None

    def method_stream_aggressive(self, pdf_path: str):
        try:
            return camelot.read_pdf(pdf_path, pages='all', flavor='stream',
                                   edge_tol=500, row_tol=10, column_tol=0,
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
            stream_tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream', edge_tol=500)
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
                                row_data = self.fix_customer_definitive_split(row_data)
                                row_data = self.fix_column_shift_after_definitive(row_data)
                                row_data = self.fix_tablets_total_split(row_data)
                                row_data = self.fix_missing_open_column(row_data)  # NUEVA: Corrige desplazamiento cuando Open vacía
                                row_data = self.clean_open_tablets_when_closed(row_data)
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
        """Procesa DataFrame para análisis"""
        try:
            analysis_df = df.copy()

            if len(analysis_df.columns) >= 18:
                base_df = analysis_df.iloc[:, :18].copy()
                base_df.columns = [
                    'Wh', 'Return_Prefix', 'Return_Slip', 'Return_Date',
                    'Jobsite', 'Cost_Center', 'Invoice_Date1', 'Invoice_Date2',
                    'Customer', 'Job_Name', 'Definitive', 'Counted_Date',
                    'Tablets', 'Total', 'Open', 'Tablets_Total',
                    'Counting_Delay', 'Validation_Delay'
                ]
                analysis_df = base_df

            for idx in analysis_df.index:
                slip_text = str(analysis_df.loc[idx, 'Return_Slip'])
                slip_match = re.search(r'(7290000\d{5})', slip_text)
                analysis_df.at[idx, 'slip_number'] = slip_match.group(1) if slip_match else ''

                open_tablets = str(analysis_df.loc[idx, 'Open'])
                counted_date = str(analysis_df.loc[idx, 'Counted_Date'])

                is_closed = (not open_tablets or open_tablets in ['', 'nan', '0']) and \
                           counted_date and counted_date not in ['', 'nan']
                analysis_df.at[idx, 'is_closed'] = is_closed

                wh_text = str(analysis_df.loc[idx, 'Return_Prefix'])
                wh_match = re.search(r'(RO-[A-Z]{2}|\d+[A-Za-z]*)', wh_text, re.IGNORECASE)
                analysis_df.at[idx, 'warehouse'] = wh_match.group(1).upper() if wh_match else 'UNKNOWN'

                customer_text = str(analysis_df.loc[idx, 'Customer'])
                analysis_df.at[idx, 'customer_name'] = customer_text[:50] if \
                                                       customer_text not in ['', 'nan'] else 'Unknown'

                return_date = str(analysis_df.loc[idx, 'Return_Date'])

                if is_closed and return_date and counted_date:
                    business_days = self.calculate_business_days(return_date, counted_date)
                    analysis_df.at[idx, 'business_days_to_close'] = business_days
                else:
                    analysis_df.at[idx, 'business_days_to_close'] = None

            return analysis_df
        except Exception as e:
            st.error(f"Error procesando DataFrame: {e}")
            return df


# ============================================================================
# FUNCIONES DE ANÁLISIS DE TABLILLAS
# ============================================================================

def calculate_tablets_metrics(df: pd.DataFrame) -> Dict:
    """Calcula métricas globales de tablillas"""
    try:
        total_tablillas = 0
        tablillas_cerradas = 0
        tablillas_abiertas = 0

        for i in range(len(df)):
            tablets_str = str(df.iloc[i, 12])
            open_tablets_str = str(df.iloc[i, 14])

            if tablets_str not in ['', 'nan', 'None']:
                tablets_list = [x.strip() for x in tablets_str.split(',')
                               if x.strip() and x.strip() != '0']
                total = len(tablets_list)

                if total > 0:
                    total_tablillas += total

                    if open_tablets_str not in ['', 'nan', 'None', '0']:
                        open_list = [x.strip() for x in open_tablets_str.split(',') if x.strip()]
                        abiertas = len(open_list)
                        tablillas_abiertas += abiertas
                        tablillas_cerradas += (total - abiertas)
                    else:
                        tablillas_cerradas += total

        return {
            'total': total_tablillas,
            'cerradas': tablillas_cerradas,
            'abiertas': tablillas_abiertas,
            'tasa_cierre': (tablillas_cerradas / total_tablillas * 100) if total_tablillas > 0 else 0
        }
    except Exception as e:
        st.error(f"Error en calculate_tablets_metrics: {e}")
        return {'total': 0, 'cerradas': 0, 'abiertas': 0, 'tasa_cierre': 0}


def create_tablets_breakdown_by_warehouse(df: pd.DataFrame) -> pd.DataFrame:
    """Breakdown de tablillas por warehouse"""
    try:
        warehouse_data = {}

        for i in range(len(df)):
            try:
                wh = str(df.iloc[i, 1])
                tablets_str = str(df.iloc[i, 12])
                open_tablets_str = str(df.iloc[i, 14])

                if tablets_str not in ['', 'nan', 'None']:
                    tablets_list = [x.strip() for x in tablets_str.split(',')
                                   if x.strip() and x.strip() != '0']
                    total = len(tablets_list)

                    if total > 0:
                        if wh not in warehouse_data:
                            warehouse_data[wh] = {'total': 0, 'cerradas': 0, 'abiertas': 0}

                        warehouse_data[wh]['total'] += total

                        if open_tablets_str not in ['', 'nan', 'None', '0']:
                            open_list = [x.strip() for x in open_tablets_str.split(',') if x.strip()]
                            abiertas = len(open_list)
                            warehouse_data[wh]['abiertas'] += abiertas
                            warehouse_data[wh]['cerradas'] += (total - abiertas)
                        else:
                            warehouse_data[wh]['cerradas'] += total
            except:
                continue

        result = []
        for wh, data in warehouse_data.items():
            result.append({
                'Warehouse': wh,
                'Total_Tablillas': data['total'],
                'Cerradas': data['cerradas'],
                'Abiertas': data['abiertas'],
                'Tasa_Cierre_%': round((data['cerradas']/data['total']*100), 2)
            })

        if result:
            return pd.DataFrame(result).sort_values('Total_Tablillas', ascending=False)
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error en create_tablets_breakdown_by_warehouse: {e}")
        return pd.DataFrame()


def create_tablets_by_customer(df: pd.DataFrame) -> pd.DataFrame:
    """Análisis de tablillas por cliente"""
    try:
        customer_data = {}

        for i in range(len(df)):
            try:
                customer = str(df.iloc[i, 8])[:50]
                tablets_str = str(df.iloc[i, 12])
                open_tablets_str = str(df.iloc[i, 14])

                if tablets_str not in ['', 'nan', 'None'] and customer not in ['', 'nan']:
                    tablets_list = [x.strip() for x in tablets_str.split(',')
                                   if x.strip() and x.strip() != '0']
                    total = len(tablets_list)

                    if total > 0:
                        if customer not in customer_data:
                            customer_data[customer] = {'total': 0, 'cerradas': 0, 'abiertas': 0}

                        customer_data[customer]['total'] += total

                        if open_tablets_str not in ['', 'nan', 'None', '0']:
                            open_list = [x.strip() for x in open_tablets_str.split(',') if x.strip()]
                            abiertas = len(open_list)
                            customer_data[customer]['abiertas'] += abiertas
                            customer_data[customer]['cerradas'] += (total - abiertas)
                        else:
                            customer_data[customer]['cerradas'] += total
            except:
                continue

        result = []
        for customer, data in customer_data.items():
            result.append({
                'Cliente': customer,
                'Total': data['total'],
                'Abiertas': data['abiertas'],
                'Cerradas': data['cerradas'],
                'Tasa_Cierre_%': round((data['cerradas']/data['total']*100), 2)
            })

        if result:
            return pd.DataFrame(result).sort_values('Abiertas', ascending=False).head(10)
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error en create_tablets_by_customer: {e}")
        return pd.DataFrame()


def validate_tablets_integrity(df: pd.DataFrame) -> List[Dict]:
    """Valida integridad: Total vs Open count"""
    try:
        discrepancies = []

        for i in range(len(df)):
            try:
                slip = str(df.iloc[i, 2])
                total_str = str(df.iloc[i, 13])
                open_str = str(df.iloc[i, 14])

                if total_str.isdigit() and open_str not in ['', 'nan', 'None']:
                    expected = int(total_str)
                    actual = len(re.findall(r'\d{2,4}[MALT]', open_str))

                    if expected != actual:
                        discrepancies.append({
                            'Slip': slip,
                            'Esperado': expected,
                            'Encontrado': actual,
                            'Diferencia': abs(expected - actual)
                        })
            except:
                continue

        return discrepancies
    except:
        return []


# ============================================================================
# EXPORTACIÓN EXCEL PROFESIONAL CON MÚLTIPLES HOJAS
# ============================================================================

def export_to_professional_excel(df: pd.DataFrame) -> io.BytesIO:
    """
    Crea Excel profesional con múltiples hojas:
    - Metadata
    - Datos principales
    - Resumen ejecutivo
    - Tablillas por warehouse
    - Top clientes
    - Discrepancias
    """
    try:
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:

            # HOJA 1: METADATA
            metadata = pd.DataFrame([{
                'Sistema': 'Camelot PDF Extractor Pro',
                'Versión': '3.0',
                'Fecha_Generación': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Total_Albaranes': len(df),
                'Empresa': 'Alsina Forms Co., Inc.'
            }])
            metadata.to_excel(writer, sheet_name='Metadata', index=False)

            # HOJA 2: DATOS PRINCIPALES (con nombres de columnas)
            export_df = df.copy()
            if len(export_df.columns) >= 18:
                export_df.columns = [
                    'Wh', 'Return_Prefix', 'Return_Slip', 'Return_Date',
                    'Jobsite', 'Cost_Center', 'Invoice_Date1', 'Invoice_Date2',
                    'Customer', 'Job_Name', 'Definitive', 'Counted_Date',
                    'Tablets', 'Total', 'Open', 'Tablets_Total',
                    'Counting_Delay', 'Validation_Delay'
                ] + list(export_df.columns[18:])
            export_df.to_excel(writer, sheet_name='Datos_Principales', index=False)

            # HOJA 3: RESUMEN EJECUTIVO TABLILLAS
            metrics = calculate_tablets_metrics(df)
            summary_df = pd.DataFrame([{
                'Total_Tablillas': metrics['total'],
                'Tablillas_Cerradas': metrics['cerradas'],
                'Tablillas_Abiertas': metrics['abiertas'],
                'Tasa_Cierre_%': round(metrics['tasa_cierre'], 2),
                'Estado': 'EXCELENTE' if metrics['tasa_cierre'] >= 80 else 'BUENO' if metrics['tasa_cierre'] >= 70 else 'REQUIERE_ATENCION'
            }])
            summary_df.to_excel(writer, sheet_name='Resumen_Ejecutivo', index=False)

            # HOJA 4: TABLILLAS POR WAREHOUSE
            warehouse_df = create_tablets_breakdown_by_warehouse(df)
            if not warehouse_df.empty:
                warehouse_df.to_excel(writer, sheet_name='Tablillas_Por_Warehouse', index=False)

            # HOJA 5: TOP CLIENTES
            customer_df = create_tablets_by_customer(df)
            if not customer_df.empty:
                customer_df.to_excel(writer, sheet_name='Top_Clientes_Tablillas', index=False)

            # HOJA 6: DISCREPANCIAS (si existen)
            discrepancies = validate_tablets_integrity(df)
            if discrepancies:
                disc_df = pd.DataFrame(discrepancies)
                disc_df.to_excel(writer, sheet_name='Discrepancias', index=False)

        buffer.seek(0)
        return buffer

    except Exception as e:
        st.error(f"Error creando Excel profesional: {e}")
        return None


# ============================================================================
# DASHBOARD INTELIGENTE DE TABLILLAS
# ============================================================================

def create_tablets_dashboard(df: pd.DataFrame):
    """Dashboard completo e inteligente de tablillas"""
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
        discrepancies = validate_tablets_integrity(df)

        if discrepancies:
            st.warning(f"⚠️ Se encontraron **{len(discrepancies)}** discrepancias entre Total y Open")

            disc_df = pd.DataFrame(discrepancies)

            col1, col2 = st.columns([2, 1])
            with col1:
                st.dataframe(disc_df, use_container_width=True, height=300)
            with col2:
                st.info("""
                **¿Qué significa esto?**

                Estas filas tienen diferencias entre:
                - **Esperado**: Valor en columna Total
                - **Encontrado**: Códigos en columna Open

                Puede indicar:
                - Tablillas cerradas recientemente
                - Errores de captura del PDF
                """)
        else:
            st.success("✅ **Integridad perfecta**: Todos los totales coinciden con los códigos en Open")

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
# APLICACIÓN PRINCIPAL
# ============================================================================

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
            **Versión 3.1**

            ✨ **Características**:
            - ✅ Saltos de línea (Tablets + Open)
            - ✅ 8 correcciones universales
            - ✅ Corrección columna Open vacía 🆕
            - ✅ Validación inteligente
            - ✅ Excel con múltiples hojas
            - ✅ Análisis histórico avanzado
            - ✅ Headers profesionales
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
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                tmp_file.write(uploaded_file.read())
                tmp_path = tmp_file.name

            extractor = CamelotExtractorPro()
            st.header("📄 Ejecutando Extracción")
            results = extractor.extract_with_all_methods(tmp_path)

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
                                acc = result.get('accuracy', 0) * 100
                                st.metric("Precisión", f"{acc:.1f}%")

                            if result.get('data') is not None and len(result['data']) > 0:
                                validation = extractor.validate_extraction(result['data'])

                                score = validation['total_rows']
                                if validation['has_fl_column']:
                                    score += 10
                                if validation['has_slip_numbers']:
                                    score += 10

                                if score > best_score:
                                    best_score = score
                                    best_method = method_name

                                st.dataframe(result['data'], use_container_width=True, height=400)

                if best_method:
                    st.header("🏆 Mejor Método de Extracción")
                    st.success(f"**{best_method}**")

                    best_data = results[best_method]['data']
                    st.session_state['extracted_data'] = best_data

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

            try:
                os.unlink(tmp_path)
            except:
                pass

    with main_tabs[1]:
        if 'extracted_data' in st.session_state and st.session_state['extracted_data'] is not None:
            create_analysis_dashboard(st.session_state['extracted_data'])
        else:
            st.info("💡 Primero extrae datos del PDF en la pestaña 'Extracción PDF'")

    with main_tabs[2]:
        if 'extracted_data' in st.session_state and st.session_state['extracted_data'] is not None:
            create_tablets_dashboard(st.session_state['extracted_data'])
        else:
            st.info("💡 Primero extrae datos del PDF en la pestaña 'Extracción PDF'")

    with main_tabs[3]:
        create_historical_dashboard()


if __name__ == "__main__":
    main()
