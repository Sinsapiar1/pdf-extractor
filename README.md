# 📄 Camelot PDF Extractor Pro v3.0

Extractor profesional de PDFs con Camelot - Sistema completo de análisis de albaranes y tablillas

![Version](https://img.shields.io/badge/version-3.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![Status](https://img.shields.io/badge/status-production-green.svg)

## 🌟 Características Principales

### 🔧 Sistema de Extracción Universal
- ✅ Soporte para **TODOS** los warehouses (RO-XX, 61D, 612D, 298T, etc.)
- ✅ Detección automática de slip numbers (`7290000XXXXX`)
- ✅ **6 métodos de extracción** con selección automática del mejor
- ✅ **6 funciones de autocorrección** que se ejecutan en pipeline
- ✅ **Unión de saltos de línea** - Detecta y une códigos en filas siguientes (Tablets + Open)
- ✅ **Respeta tablillas cerradas** - NO inventa códigos, solo extrae lo que existe

### 📊 Dashboards Profesionales

#### 1. Dashboard de Albaranes
- KPIs principales (Total, Cerrados, Pendientes, Tasa de Cierre)
- Análisis por Warehouse
- Tiempos de cierre (días hábiles)
- Top Customers por volumen

#### 2. Dashboard Inteligente de Tablillas ⭐ NUEVO
- Métricas globales de tablillas (Total, Abiertas, Cerradas)
- Visualización con Pie Chart y Gauge Chart
- Breakdown por Warehouse y Cliente
- **Validación de integridad** (Total vs Open count)
- **Alertas inteligentes** automáticas

#### 3. Dashboard Histórico Mejorado ⭐ NUEVO
- Análisis comparativo entre fechas
- Evolución temporal de tablillas
- Análisis por warehouse histórico
- Tasa de cierre histórica
- Exportación consolidada con múltiples hojas

### 📊 Exportación Excel Profesional ⭐ NUEVO
El sistema genera Excel con **6 hojas**:
1. **Metadata** - Info del sistema, fecha, empresa
2. **Datos_Principales** - Data completa con 18 columnas
3. **Resumen_Ejecutivo** - KPIs globales de tablillas
4. **Tablillas_Por_Warehouse** - Breakdown detallado
5. **Top_Clientes_Tablillas** - Top 10 clientes
6. **Discrepancias** - Validación de integridad (si existen)

## 🚀 Instalación Local
```bash
# Clonar repositorio
git clone https://github.com/tu-usuario/pdf-extractor.git
cd pdf-extractor

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar aplicación
streamlit run app.py
📋 Estructura de Datos
Columnas Esperadas (18 columnas)
#ColumnaDescripciónEjemplo0WhEstado (FL, DL, TX, CA, NY)FL1Return_PrefixWarehouse code61D, 612D, RO-FL2Return_SlipSlip number7290000188223Return_DateFecha de retorno10/1/20254JobsiteCódigo de obra400366455Cost_CenterCentro de costoFL0526Invoice_Date1Fecha factura 18/31/20257Invoice_Date2Fecha factura 29/30/20258CustomerNombre del clienteThales Builders Corp9Job_NameNombre del proyectoResidences at Martin10DefinitiveDefinitivo (Yes/No)No11Counted_DateFecha de conteo10/5/202512TabletsCódigos de tablillas1321, 1656, 166113TotalTotal tablillas ABIERTAS314OpenCódigos tablillas abiertas1656T, 1661A, 1665T15Tablets_TotalTotal de tablillas416Counting_DelayDías de retraso conteo517Validation_DelayDías retraso validación0
🔧 Sistema de Correcciones Automáticas
El sistema ejecuta 6 funciones de autocorrección en secuencia:
1. merge_continuation_rows() 🆕 MEJORADO
Problema: Códigos de tablillas en salto de línea
Fila actual:  Tablets = "1703, 1707, 1710, 1728,"
              Open = "84A, 1651A, 1657T, 1666A,"
Fila siguiente:        "1733, 1736"
                       "1759A"
Solución: Une automáticamente AMBAS columnas
Tablets = "1703, 1707, 1710, 1728, 1733, 1736"
Open = "84A, 1651A, 1657T, 1666A, 1759A"
2. ensure_18_columns()
Garantiza que todas las filas tengan exactamente 18 columnas
3. clean_warehouse_slip_column()
Separa warehouse code y slip number cuando están juntos
4. fix_customer_definitive_split()
Separa customer name cuando termina en "No" o "Yes"
5. fix_column_shift_after_definitive()
Corrige desplazamiento cuando Definitive="No"
6. fix_tablets_total_split()
Separa Total y Open cuando están mezclados
🎯 Casos de Uso
Caso 1: Extracción de PDF
bash1. Subir PDF en Tab "Extracción PDF"
2. Sistema prueba 6 métodos automáticamente
3. Selecciona el mejor método
4. Aplica las 6 correcciones en pipeline
5. Muestra validación con completitud %
6. Exportar a Excel profesional con 6 hojas
Caso 2: Análisis de Tablillas
bash1. Extraer datos del PDF
2. Ir a Tab "Dashboard de Tablillas"
3. Ver métricas globales (Total, Abiertas, Cerradas)
4. Analizar por Warehouse y Cliente
5. Revisar discrepancias de integridad
6. Actuar sobre alertas inteligentes
Caso 3: Análisis Histórico
bash1. Ir a Tab "Análisis Histórico"
2. Cargar múltiples archivos Excel (generados por la app)
3. Ver evolución temporal de tablillas
4. Analizar tendencias por warehouse
5. Comparar tasas de cierre entre fechas
6. Exportar consolidado con 3 hojas
🛠️ Métodos de Extracción
El sistema prueba 6 métodos y selecciona automáticamente el mejor:
MétodoDescripciónMejor paramethod_lattice_standardLattice estándarPDFs con tablas definidasmethod_stream_balancedStream balanceadoPDFs mixtosmethod_stream_standardStream estándarPDFs simplesmethod_stream_aggressiveStream agresivoPDFs complejosmethod_lattice_detailedLattice detalladoPDFs con muchas líneasmethod_hybridStream + LatticePDFs difíciles
📈 Métricas y KPIs
Albaranes

Total de albaranes
Albaranes cerrados/pendientes
Tasa de cierre (%)
Días hábiles promedio para cierre
Performance por warehouse

Tablillas ⭐ NUEVO

Total de tablillas en inventario
Tablillas abiertas/cerradas
Tasa de cierre de tablillas (%)
Distribución por warehouse
Top clientes con tablillas abiertas
Validación de integridad automática

🐛 Troubleshooting
PDF no extrae correctamente

Verificar que el PDF tenga la estructura esperada
Probar diferentes métodos en las pestañas
Revisar el debug info para ver detalles

Columnas desalineadas

Las 6 funciones de autocorrección deberían resolver esto
Si persiste, verificar que el PDF tenga 18 columnas

Discrepancias en Total vs Open

Esto es NORMAL cuando tablillas se cierran entre extracciones
El sistema solo reporta, NO inventa códigos

Análisis histórico no funciona

Asegurarse de usar archivos Excel generados por esta app (v3.0)
Los archivos deben tener la hoja "Datos_Principales"

📚 Documentación Adicional
Ver HANDOFF.md para:

Arquitectura detallada del sistema
Explicación técnica de cada función
Diagramas de flujo
Casos edge detallados
Guía de desarrollo

🔐 Notas Técnicas

Sistema 100% universal, sin hardcoding
Todas las correcciones basadas en patrones regex
Pipeline de correcciones garantiza 18 columnas
Respeta estado real de tablillas (NO inventa datos)
Diseñado para PDFs de "Outstanding Count Returns"
Excel compatible con análisis histórico

📦 Versión
v3.0 Final - Sistema completo con Dashboard Inteligente de Tablillas y Excel Profesional
Changelog v3.0

✅ Dashboard Inteligente de Tablillas
✅ Validación de integridad (Total vs Open)
✅ Alertas inteligentes automáticas
✅ Unión de saltos de línea mejorada (Tablets + Open)
✅ Excel profesional con 6 hojas
✅ Análisis histórico mejorado con gráficos por warehouse
✅ Headers profesionales con branding

Changelog v2.0

Sistema de 6 correcciones automáticas
Garantía de 18 columnas
Soporte universal para warehouses

👥 Contribuir
Para reportar bugs o sugerir mejoras, crear un issue en GitHub.
📄 Licencia
MIT License - Ver LICENSE para más detalles

Desarrollado con ❤️ usando Streamlit + Camelot + Plotly
