
"""
Extractor Profesional de PDFs con Camelot
UNIVERSAL - Soporta todos los warehouses con correcciones automáticas
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
    page_title="Camelot PDF Extractor Pro",
    page_icon="📄",
    layout="wide"
)

class CamelotExtractorPro:
    """Extractor especializado - versión profesional con correcciones universales"""

    def __init__(self):
        self.extraction_methods = [
            self.method_lattice_standard,
            self.method_stream_balanced,
            self.method_stream_standard,
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

    def method_stream_with_columns(self, pdf_path: str):
        """Método stream con coordenadas de columnas FORZADAS"""
        try:
            return camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='stream',
                columns=['60,120,180,240,300,360,420,480,540,600,660,720,780,840,900,960,1020'],
                split_text=True
            )
        except:
            return None

    def method_lattice_standard(self, pdf_path: str):
        """Método lattice estándar"""
        try:
            return camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='lattice',
                process_background=True,
                line_scale=40
            )
        except:
            return None

    def ensure_18_columns(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Asegura que el DataFrame tenga exactamente 18 columnas - UNIVERSAL"""
        try:
            current_cols = len(row_data.columns)
            if current_cols < 18:
                for i in range(current_cols, 18):
                    row_data[i] = ''
            return row_data
        except Exception as e:
            return row_data

    def separate_merged_row(self, merged_text: str) -> Optional[pd.DataFrame]:
        """Separa una fila que está toda junta en una sola celda - UNIVERSAL"""
        try:
            parts = []
            parts.append('FL')
            
            wh_match = re.search(r'(RO-[A-Z]{2}|\d+[A-Za-z]*)', merged_text)
            parts.append(wh_match.group(1).upper() if wh_match else '612D')
            
            slip_match = re.search(r'(7290000\d{5})', merged_text)
            parts.append(slip_match.group(1) if slip_match else '')
            
            dates = re.findall(r'(\d{1,2}/\d{1,2}/\d{4})', merged_text)
            parts.extend(dates[:4])
            while len(parts) < 7:
                parts.append('')
            remaining = merged_text
            for part in parts:
                remaining = remaining.replace(str(part), '', 1)
            remaining_parts = remaining.split()[:11]
            parts.extend(remaining_parts)
            while len(parts) < 18:
                parts.append('')
            return pd.DataFrame([parts[:18]])
        except Exception as e:
            return None

    def clean_warehouse_slip_column(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Separa warehouse code y slip number si están juntos - UNIVERSAL"""
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

        except Exception as e:
            return row_data

    def fix_customer_definitive_split(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Separa customer name de definitive cuando están unidos - UNIVERSAL"""
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
            
        except Exception as e:
            return row_data

    def fix_column_shift_after_definitive(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Corrige desplazamiento cuando Definitive=No y Counted date tiene tablets - UNIVERSAL"""
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
            
        except Exception as e:
            return row_data

    def fix_tablets_total_split(self, row_data: pd.DataFrame) -> pd.DataFrame:
        """Separa Total de Open preservando todas las columnas - UNIVERSAL Y PROFESIONAL"""
        try:
            if len(row_data.columns) < 16:
                return row_data
            
            total_cell = str(row_data.iloc[0, 13]).strip()
            
            pattern = r'^(\d+)\s+([\d\s,]+[MALT].*)$'
            match = re.match(pattern, total_cell)
            
            if match:
                total_number = match.group(1)
                open_tablets = match.group(2).strip()
                
                # Guardar TODOS los valores desde col 14 hasta el final
                saved_values = []
                for col_idx in range(14, min(18, len(row_data.columns))):
                    saved_values.append(str(row_data.iloc[0, col_idx]))
                
                # Reconstruir correctamente
                row_data.iloc[0, 13] = total_number        # Total
                row_data.iloc[0, 14] = open_tablets        # Open Tablets
                
                # Desplazar los valores guardados: lo que estaba en 14 va a 15, etc.
                for i, val in enumerate(saved_values):
                    new_col = 15 + i
                    if new_col < len(row_data.columns):
                        row_data.iloc[0, new_col] = val
            
            return row_data
            
        except Exception as e:
            return row_data

    def method_stream_balanced(self, pdf_path: str):
        """Método stream balanceado"""
        try:
            return camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='stream',
                edge_tol=350,
                row_tol=12,
                column_tol=5
            )
        except:
            return None

    def method_stream_aggressive(self, pdf_path: str):
        """Stream con configuración agresiva"""
        try:
            return camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='stream',
                edge_tol=500,
                row_tol=10,
                column_tol=0,
                split_text=True,
                flag_size=True
            )
        except:
            return None

    def method_stream_standard(self, pdf_path: str):
        """Stream estándar"""
        try:
            return camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='stream'
            )
        except:
            return None

    def method_lattice_detailed(self, pdf_path: str):
        """Lattice con configuración detallada"""
        try:
            return camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='lattice',
                process_background=True,
                line_scale=40,
                iterations=2
            )
        except:
            return None

    def method_hybrid(self, pdf_path: str):
        """Método híbrido: combina Stream y Lattice"""
        all_tables = []

        try:
            stream_tables = camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='stream',
                edge_tol=500
            )
            if stream_tables:
                all_tables.extend(stream_tables)
        except:
            pass

        try:
            lattice_tables = camelot.read_pdf(
                pdf_path,
                pages='all',
                flavor='lattice'
            )
            if lattice_tables:
                all_tables.extend(lattice_tables)
        except:
            pass

        return all_tables if all_tables else None

    def process_tables(self, tables) -> pd.DataFrame:
        """Procesa las tablas con correcciones automáticas universales"""
        if not tables:
            return None

        all_data = []
        total_pages = len(tables)

        st.info(f"📄 PDF detectado con {total_pages} páginas")

        for i, table in enumerate(tables):
            try:
                df = table.df
                page_num = i + 1

                st.write(f"📋 Procesando página {page_num}: {df.shape}")

                for idx in df.index:
                    try:
                        row_text = ' '.join(str(cell) for cell in df.iloc[idx].values if pd.notna(cell))

                        if re.search(r'7290000\d{5}', row_text) and re.search(r'\b[A-Z]{2}\b', row_text):
                            if not any(skip in row_text for skip in ['Outstanding count', 'Page', 'Return packing', 'Customer name', 'Alsina Forms']):
                                row_data = df.iloc[idx:idx+1].copy()

                                first_cell = str(row_data.iloc[0, 0])
                                if len(first_cell) > 100 and re.search(r'7290000\d{5}', first_cell):
                                    separated_row = self.separate_merged_row(first_cell)
                                    if separated_row:
                                        all_data.append(separated_row)
                                        continue

                                # PIPELINE DE CORRECCIONES UNIVERSALES
                                row_data = self.ensure_18_columns(row_data)  # PRIMERO: Asegurar 18 columnas
                                row_data = self.clean_warehouse_slip_column(row_data)
                                row_data = self.fix_customer_definitive_split(row_data)
                                row_data = self.fix_column_shift_after_definitive(row_data)
                                row_data = self.fix_tablets_total_split(row_data)
                                all_data.append(row_data)

                    except Exception as e:
                        continue

            except Exception as e:
                st.error(f"Error procesando página {page_num}: {e}")
                continue

        if all_data:
            try:
                result = pd.concat(all_data, ignore_index=True)
                result = self.fix_merged_rows(result)
                self.validate_simple(result)
                return result
            except Exception as e:
                st.error(f"Error combinando datos: {e}")
                return None

        return None

    def fix_merged_rows(self, df: pd.DataFrame) -> pd.DataFrame:
        """Detecta y separa filas que están todas juntas en una celda"""
        fixed_rows = []

        for idx in df.index:
            row = df.iloc[idx]
            merged_detected = False
            merged_cell_text = None

            for col_idx in range(min(3, len(row))):
                cell_text = str(row.iloc[col_idx])

                if len(cell_text) > 80 and re.search(r'7290000\d{5}', cell_text):
                    merged_detected = True
                    merged_cell_text = cell_text
                    st.info(f"⚠️ Fila unida detectada en índice {idx}, columna {col_idx}")
                    break

                if col_idx in [1, 2] and re.search(r'7290000\d{5}', cell_text) and len(cell_text) > 30:
                    merged_detected = True
                    merged_cell_text = cell_text
                    st.info(f"⚠️ Fila con datos unidos en columna {col_idx}, índice {idx}")
                    break

            if merged_detected and merged_cell_text:
                separated = self.separate_unified_row(merged_cell_text)
                if separated is not None:
                    st.success(f"✅ Fila separada exitosamente")
                    fixed_rows.append(separated)
                else:
                    st.warning(f"⚠️ No se pudo separar automáticamente, manteniendo original")
                    fixed_rows.append(row.to_frame().T)
            else:
                fixed_rows.append(row.to_frame().T)

        if fixed_rows:
            return pd.concat(fixed_rows, ignore_index=True)
        return df

    def separate_unified_row(self, text: str) -> Optional[pd.DataFrame]:
        """Separa una fila que está toda unida usando patrones - UNIVERSAL"""
        try:
            parts = []
            parts.append('FL')
            
            wh_match = re.search(r'(RO-[A-Z]{2}|\d+[A-Za-z]*)\s', text)
            wh_code = wh_match.group(1).upper() if wh_match else '612D'
            parts.append(wh_code)
            
            slip_match = re.search(r'(7290000\d{5})', text)
            parts.append(slip_match.group(1) if slip_match else '')
            dates = re.findall(r'(\d{1,2}/\d{1,2}/\d{4})', text)
            parts.append(dates[0] if len(dates) > 0 else '')
            jobsite_match = re.search(r'(4\d{7})', text)
            parts.append(jobsite_match.group(1) if jobsite_match else '')
            cost_match = re.search(r'(FL\d{3})', text)
            parts.append(cost_match.group(1) if cost_match else '')
            parts.append(dates[1] if len(dates) > 1 else '')
            parts.append(dates[2] if len(dates) > 2 else '')
            customers = ['Thales Builders Corp', 'Caribbean Building Corp', 'Phorcys Builders Corp',
                        'JGR Construction', 'O&R Construction', 'American Concrete shell']
            customer = ''
            for cust in customers:
                if cust in text:
                    customer = cust
                    break
            parts.append(customer)
            if customer:
                job_pattern = f'{re.escape(customer)}\\s+([^\\n]+?)\\s+(Yes|No|Ye)'
                job_match = re.search(job_pattern, text)
                parts.append(job_match.group(1).strip() if job_match else '')
            else:
                parts.append('')
            def_match = re.search(r'\b(Yes|No|Ye)\b', text)
            parts.append(def_match.group(1) if def_match else 'No')
            parts.append(dates[3] if len(dates) > 3 else '')
            tablet_codes = re.findall(r'(\d{2,4}[MALT])', text)
            parts.append(', '.join(tablet_codes[:4]) if tablet_codes else '')
            final_numbers = re.findall(r'\b(\d{1,2})\b', text[-80:])
            parts.append(final_numbers[-5] if len(final_numbers) >= 5 else '1')
            parts.append(', '.join(tablet_codes[-2:]) if len(tablet_codes) > 4 else '')
            parts.append(final_numbers[-4] if len(final_numbers) >= 4 else '1')
            parts.append(final_numbers[-2] if len(final_numbers) >= 2 else '0')
            parts.append(final_numbers[-1] if len(final_numbers) >= 1 else '0')
            while len(parts) < 18:
                parts.append('')
            return pd.DataFrame([parts[:18]])
        except Exception as e:
            st.warning(f"No se pudo separar fila automáticamente: {e}")
            return None

    def validate_simple(self, df: pd.DataFrame):
        """Validación simple y efectiva"""
        if df is None or df.empty:
            st.error("❌ DataFrame vacío")
            return

        try:
            st.header("🔍 Validación del Sistema")

            total_rows = len(df)
            slip_count = 0
            valid_slips = []

            for idx in df.index:
                row_text = ' '.join(str(cell) for cell in df.iloc[idx].values if pd.notna(cell))
                if re.search(r'7290000\d{5}', row_text):
                    slip_match = re.search(r'(7290000\d{5})', row_text)
                    if slip_match:
                        slip_count += 1
                        valid_slips.append(slip_match.group(1))

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📊 Filas Totales", total_rows)
            with col2:
                st.metric("📊 Slips Válidos", slip_count)
            with col3:
                completeness = (slip_count / total_rows * 100) if total_rows > 0 else 0
                st.metric("📊 Completitud", f"{completeness:.1f}%")

            if len(valid_slips) > 1:
                slip_numbers = [int(slip[-3:]) for slip in valid_slips]
                slip_numbers.sort()
                first_slip = slip_numbers[0]
                last_slip = slip_numbers[-1]
                expected_count = last_slip - first_slip + 1

                if len(valid_slips) == expected_count:
                    st.success(f"✅ Secuencia completa: {first_slip} a {last_slip} ({len(valid_slips)} slips)")
                else:
                    missing = expected_count - len(valid_slips)
                    st.warning(f"⚠️ Secuencia incompleta: Faltan {missing} slips")

            total_pattern = self.find_pdf_totals(df)
            if total_pattern:
                st.success(f"🎯 Totales detectados: {total_pattern}")

            if completeness >= 95:
                st.success("🎉 **EXTRACCIÓN EXCELENTE** - Datos completos")
            elif completeness >= 80:
                st.info("📊 **EXTRACCIÓN BUENA** - Datos casi completos")
            else:
                st.warning("⚠️ **EXTRACCIÓN PARCIAL** - Revisar configuración")

        except Exception as e:
            st.error(f"Error en validación: {e}")

    def find_pdf_totals(self, df: pd.DataFrame):
        """Busca totales en el PDF de manera simple"""
        try:
            last_rows_text = ""
            for idx in df.tail(3).index:
                row_text = ' '.join(str(cell) for cell in df.iloc[idx].values if pd.notna(cell))
                last_rows_text += row_text + " "

            matches = re.findall(r'\b(\d{2,3})\s+(\d{2,3})\b', last_rows_text)
            if matches:
                for match in matches:
                    num1, num2 = int(match[0]), int(match[1])
                    if 50 <= num1 <= 500 and 20 <= num2 <= 200:
                        return f"{num1} / {num2}"

            return None
        except:
            return None

    def calculate_accuracy(self, tables) -> float:
        """Calcula un score de precisión simple"""
        try:
            if not tables:
                return 0.0
            total_accuracy = sum(getattr(table, 'accuracy', 0) for table in tables)
            return total_accuracy / len(tables)
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

        except Exception as e:
            validation['data_quality'] = 'error'

        return validation


class BusinessAnalyzer:
    """Analizador de métricas de negocio para albaranes"""

    def __init__(self):
        self.us_holidays = holidays.US(years=range(2024, 2027))

    def calculate_business_days(self, start_date_str: str, end_date_str: str) -> int:
        """Calcula días hábiles entre dos fechas"""
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
        """Procesa el DataFrame para análisis de negocio"""
        try:
            analysis_df = df.copy()

            extra_columns = {}
            if len(analysis_df.columns) > 18:
                for i in range(18, len(analysis_df.columns)):
                    col_name = analysis_df.columns[i]
                    extra_columns[col_name] = analysis_df[col_name].copy()

            if len(analysis_df.columns) >= 18:
                base_df = analysis_df.iloc[:, :18].copy()
                base_df.columns = [
                    'Wh', 'Return_Prefix', 'Return_Slip', 'Return_Date',
                    'Jobsite', 'Cost_Center', 'Invoice_Date1', 'Invoice_Date2',
                    'Customer', 'Job_Name', 'Definitive', 'Counted_Date',
                    'Tablets', 'Total', 'Open', 'Tablets_Total',
                    'Counting_Delay', 'Validation_Delay'
                ]
                
                for col_name, col_data in extra_columns.items():
                    base_df[col_name] = col_data
                
                analysis_df = base_df

            for idx in analysis_df.index:
                slip_text = str(analysis_df.loc[idx, 'Return_Slip'])
                slip_match = re.search(r'(7290000\d{5})', slip_text)
                analysis_df.at[idx, 'slip_number'] = slip_match.group(1) if slip_match else ''

                open_tablets = str(analysis_df.loc[idx, 'Open'])
                counted_date = str(analysis_df.loc[idx, 'Counted_Date'])

                is_closed = (not open_tablets or open_tablets in ['', 'nan', '0']) and counted_date and counted_date not in ['', 'nan']
                analysis_df.at[idx, 'is_closed'] = is_closed

                wh_text = str(analysis_df.loc[idx, 'Return_Prefix'])
                wh_match = re.search(r'(RO-[A-Z]{2}|\d+[A-Za-z]*)', wh_text, re.IGNORECASE)
                analysis_df.at[idx, 'warehouse'] = wh_match.group(1).upper() if wh_match else 'UNKNOWN'

                customer_text = str(analysis_df.loc[idx, 'Customer'])
                analysis_df.at[idx, 'customer_name'] = customer_text[:50] if customer_text not in ['', 'nan'] else 'Unknown'

                return_date = str(analysis_df.loc[idx, 'Return_Date'])

                if is_closed and return_date and counted_date:
                    business_days = self.calculate_business_days(return_date, counted_date)
                    analysis_df.at[idx, 'business_days_to_close'] = business_days
                else:
                    analysis_df.at[idx, 'business_days_to_close'] = None

            return analysis_df

        except Exception as e:
            st.error(f"Error procesando DataFrame: {e}")
            import traceback
            st.error(traceback.format_exc())
            return df

def create_analysis_dashboard(df: pd.DataFrame):
    """Dashboard de análisis profesional"""
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
                    labels={'business_days_to_close': 'Días Hábiles', 'count': 'Cantidad de Albaranes'}
                )
                st.plotly_chart(fig, use_container_width=True)

        st.subheader("🏢 Top Customers por Volumen")

        customer_stats = analysis_df.groupby('customer_name').agg({
            'is_closed': ['count', 'sum'],
            'business_days_to_close': 'mean'
        }).round(1)

        customer_stats.columns = ['Total_Albaranes', 'Cerrados', 'Dias_Promedio']
        customer_stats = customer_stats.sort_values('Total_Albaranes', ascending=False).head(10)

        if len(customer_stats) > 0:
            st.dataframe(customer_stats, use_container_width=True)

            fig = px.bar(
                customer_stats.reset_index().head(8),
                x='customer_name',
                y='Total_Albaranes',
                title="Top 8 Customers por Volumen de Albaranes"
            )
            fig.update_layout(xaxis_tickangle=45)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("⚠️ Alertas y Recomendaciones")

        alerts = []

        if closure_rate < 70:
            alerts.append(f"🔴 **Tasa de cierre baja:** {closure_rate:.1f}% (objetivo: >80%)")

        if len(closed_data) > 0:
            valid_days = closed_data.dropna(subset=['business_days_to_close'])
            if len(valid_days) > 0:
                if valid_days['business_days_to_close'].mean() > 7:
                    alerts.append(f"🔴 **Días promedio alto:** {valid_days['business_days_to_close'].mean():.1f} días (objetivo: <7 días)")

                old_albaranes = len(valid_days[valid_days['business_days_to_close'] > 15])
                if old_albaranes > 0:
                    alerts.append(f"🟡 **Albaranes antiguos:** {old_albaranes} albaranes con >15 días")

        if alerts:
            for alert in alerts:
                st.warning(alert)
        else:
            st.success("✅ **Excelente rendimiento:** Todos los KPIs dentro del rango objetivo")

    except Exception as e:
        st.error(f"Error creando dashboard: {e}")

def create_historical_dashboard():
    """Dashboard de análisis histórico"""
    st.header("📈 Dashboard Histórico - Análisis Comparativo")

    st.info("📁 Carga múltiples archivos Excel generados día a día para análisis de tendencias")

    uploaded_files = st.file_uploader(
        "Selecciona archivos Excel (puedes seleccionar múltiples archivos)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Carga los archivos Excel que la app ha generado en diferentes fechas"
    )

    if uploaded_files and len(uploaded_files) > 0:
        st.success(f"✅ {len(uploaded_files)} archivos cargados")

        all_data = []
        file_info = []

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file)

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
            st.subheader("📋 Archivos Cargados")
            st.dataframe(pd.DataFrame(file_info), use_container_width=True)

            combined_df = pd.concat(all_data, ignore_index=True)
            analyzer = BusinessAnalyzer()

            with st.spinner("Procesando datos históricos..."):
                if len(combined_df.columns) >= 18:
                    combined_df.columns = list(combined_df.columns[:18]) + list(combined_df.columns[18:])
                    combined_df = analyzer.parse_dataframe(combined_df)

            st.subheader("📊 Evolución de KPIs en el Tiempo")

            daily_stats = []
            for fecha in sorted(combined_df['fecha_archivo'].unique()):
                df_fecha = combined_df[combined_df['fecha_archivo'] == fecha]

                total = len(df_fecha)
                cerrados = len(df_fecha[df_fecha['is_closed'] == True]) if 'is_closed' in df_fecha.columns else 0
                tasa_cierre = (cerrados / total * 100) if total > 0 else 0

                daily_stats.append({
                    'Fecha': fecha.strftime('%Y-%m-%d'),
                    'Total_Albaranes': total,
                    'Cerrados': cerrados,
                    'Pendientes': total - cerrados,
                    'Tasa_Cierre': tasa_cierre
                })

            daily_df = pd.DataFrame(daily_stats)

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=daily_df['Fecha'],
                y=daily_df['Total_Albaranes'],
                mode='lines+markers',
                name='Total Albaranes',
                line=dict(color='#1f77b4', width=3)
            ))
            fig.add_trace(go.Scatter(
                x=daily_df['Fecha'],
                y=daily_df['Cerrados'],
                mode='lines+markers',
                name='Cerrados',
                line=dict(color='#2ca02c', width=3)
            ))
            fig.update_layout(
                title="Evolución de Albaranes en el Tiempo",
                xaxis_title="Fecha",
                yaxis_title="Cantidad",
                hovermode='x unified'
            )
            st.plotly_chart(fig, use_container_width=True)

            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(
                x=daily_df['Fecha'],
                y=daily_df['Tasa_Cierre'],
                mode='lines+markers',
                name='Tasa de Cierre',
                line=dict(color='#ff7f0e', width=3),
                fill='tozeroy'
            ))
            fig2.add_hline(y=80, line_dash="dash", line_color="green", annotation_text="Objetivo: 80%")
            fig2.update_layout(
                title="Tasa de Cierre Histórica (%)",
                xaxis_title="Fecha",
                yaxis_title="Tasa de Cierre (%)",
                hovermode='x unified'
            )
            st.plotly_chart(fig2, use_container_width=True)

            st.subheader("🏭 Comparativo por Warehouse")

            if 'warehouse' in combined_df.columns:
                warehouse_comparison = combined_df.groupby(['fecha_archivo', 'warehouse']).size().reset_index(name='count')
                warehouse_comparison['Fecha'] = warehouse_comparison['fecha_archivo'].dt.strftime('%Y-%m-%d')

                fig3 = px.bar(
                    warehouse_comparison,
                    x='Fecha',
                    y='count',
                    color='warehouse',
                    title="Distribución por Warehouse a lo Largo del Tiempo",
                    labels={'count': 'Cantidad de Albaranes'}
                )
                st.plotly_chart(fig3, use_container_width=True)

            st.subheader("📋 Resumen Comparativo")
            st.dataframe(daily_df, use_container_width=True)

            st.subheader("💾 Exportar Consolidado")

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Datos_Consolidados', index=False)
                daily_df.to_excel(writer, sheet_name='Resumen_Diario', index=False)

            st.download_button(
                "📊 Descargar Análisis Histórico Consolidado",
                buffer.getvalue(),
                f"historico_consolidado_{datetime.now().strftime('%Y%m%d')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.warning("👆 Sube al menos un archivo Excel para comenzar el análisis histórico")

def main():
    st.title("📄 Camelot PDF Extractor Pro")
    st.markdown("**Versión PROFESIONAL con correcciones automáticas universales**")

    main_tabs = st.tabs(["📄 Extracción PDF", "📊 Análisis Profesional", "📈 Análisis Histórico"])

    with main_tabs[0]:
        with st.sidebar:
            st.header("📊 Información")
            st.info("""
            Versión profesional con:
            - Soporte UNIVERSAL para todos los warehouses
            - 4 funciones de autocorrección
            - Garantía de 18 columnas
            - Slip numbers 7290000XXXXX
            """)

            show_debug = st.checkbox("Mostrar información de debug", value=False)
            show_raw_data = st.checkbox("Mostrar datos crudos", value=False)

        uploaded_file = st.file_uploader(
            "Selecciona el PDF",
            type=['pdf'],
            help="Sube el reporte Outstanding Count Returns"
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
                                st.metric("Tablas encontradas", result.get('tables_found', 0))
                            with col2:
                                st.metric("Filas extraídas", result.get('rows', 0))
                            with col3:
                                accuracy = result.get('accuracy', 0) * 100
                                st.metric("Precisión", f"{accuracy:.1f}%")

                            if result.get('data') is not None and len(result['data']) > 0:
                                validation = extractor.validate_extraction(result['data'])

                                st.subheader("✅ Validación")

                                col1, col2 = st.columns(2)
                                with col1:
                                    st.write(f"✓ Columna FL: {'✅' if validation['has_fl_column'] else '❌'}")
                                    st.write(f"✓ Números slip: {'✅' if validation['has_slip_numbers'] else '❌'}")
                                with col2:
                                    quality = validation['data_quality']
                                    color = "🟢" if quality == 'good' else "🟡" if quality == 'acceptable' else "🔴"
                                    st.write(f"✓ Calidad: {color} {quality}")

                                score = validation['total_rows']
                                if validation['has_fl_column']:
                                    score += 10
                                if validation['has_slip_numbers']:
                                    score += 10

                                if score > best_score:
                                    best_score = score
                                    best_method = method_name

                                if show_raw_data:
                                    st.subheader("📋 Datos Extraídos")
                                    st.dataframe(result['data'])

                                st.subheader("📋 Tabla de Datos Extraídos")
                                st.dataframe(result['data'], use_container_width=True, height=400)

                                st.subheader("👁️ Vista Previa")
                                st.write("Primeras 5 filas:")
                                st.dataframe(result['data'].head())
                                st.write("Últimas 5 filas:")
                                st.dataframe(result['data'].tail())

                                if show_debug:
                                    with st.expander("🔧 Debug Info"):
                                        st.write(f"Shape: {result['data'].shape}")
                                        st.write(f"Columnas: {list(result['data'].columns)}")
                            else:
                                st.warning("No hay datos para mostrar en este método")

                if best_method:
                    st.header("🏆 Mejor Método")
                    st.success(f"**{best_method}**")

                    best_data = results[best_method]['data']

                    st.session_state['extracted_data'] = best_data

                    st.subheader("💾 Exportar Datos")

                    col1, col2 = st.columns(2)

                    with col1:
                        try:
                            csv = best_data.to_csv(index=False)
                            st.download_button(
                                "📄 Descargar CSV",
                                csv,
                                f"extracted_data_{datetime.now().strftime('%Y%m%d')}.csv",
                                "text/csv"
                            )
                        except Exception as e:
                            st.error(f"Error CSV: {e}")

                    with col2:
                        try:
                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                best_data.to_excel(writer, sheet_name='Extracted', index=False)

                            st.download_button(
                                "📊 Descargar Excel",
                                buffer.getvalue(),
                                f"extracted_data_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as e:
                            st.error(f"Error Excel: {e}")

            try:
                os.unlink(tmp_path)
            except:
                pass

    with main_tabs[1]:
        if 'extracted_data' in st.session_state and st.session_state['extracted_data'] is not None:
            create_analysis_dashboard(st.session_state['extracted_data'])
        else:
            st.info("💡 Primero extrae datos de un PDF en la pestaña 'Extracción PDF'")

    with main_tabs[2]:
        create_historical_dashboard()

if __name__ == "__main__":
    main()