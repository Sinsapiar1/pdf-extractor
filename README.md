# 📄 Camelot PDF Extractor Pro v3.1

Extractor profesional de PDFs con Camelot - Sistema completo de análisis de albaranes y tablillas

![Version](https://img.shields.io/badge/version-3.1-blue.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)
![Status](https://img.shields.io/badge/status-production-green.svg)

## 🆕 Nueva Actualización v3.1 - Corrección Crítica para Cierre de Mes

### ✅ Problema Solucionado: Columna "Open Tablets" Vacía

**¿Qué pasaba antes?**
- En el último día de cierre de mes, cuando **todas las tablillas están cerradas**, la columna "Open Tablets" queda completamente vacía en el PDF
- Camelot **no detecta columnas vacías**, causando que todas las columnas posteriores se desplacen una posición a la izquierda
- Esto provocaba que los datos se extrajeran incorrectamente: `Tablets_Total` aparecía en la posición de `Open`, causando confusión en los reportes

**¿Qué hace ahora la v3.1?**
- ✅ **Detecta automáticamente** cuando la columna Open está vacía (albaranes cerrados)
- ✅ **Corrige el desplazamiento** insertando la columna vacía en la posición correcta
- ✅ **Reubica todas las columnas** automáticamente sin intervención manual
- ✅ **Trabaja silenciosamente** sin spam de logs en la interfaz
- ✅ **100% automático** - no requiere configuración adicional

**Casos de uso:**
- 📅 Reportes de fin de mes (cuando todo está cerrado)
- 📊 Análisis de albaranes completados
- 💾 Exportaciones Excel con datos precisos

## 🌟 Características Principales

### 🔧 Sistema de Extracción Universal
- ✅ Soporte para **TODOS** los warehouses (RO-XX, 61D, 612D, 298T, etc.)
- ✅ Detección automática de slip numbers (`7290000XXXXX`)
- ✅ **6 métodos de extracción** con selección automática del mejor
- ✅ **8 funciones de autocorrección** que se ejecutan en pipeline
- ✅ **Unión de saltos de línea** - Detecta y une códigos en filas siguientes (Tablets + Open)
- ✅ **Corrección de columna Open vacía** 🆕 - Detecta y corrige desplazamientos cuando todas las tablillas están cerradas (**perfecto para cierre de mes**)
- ✅ **Priorización inteligente** - Método `stream_standard` optimizado para PDFs con columnas vacías
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
El sistema ejecuta 8 funciones de autocorrección en secuencia:
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
7. fix_missing_open_column() 🆕 **CRÍTICO PARA CIERRE DE MES**
**Problema:** En el último día del mes, cuando todas las tablillas están cerradas, la columna "Open" está completamente vacía en el PDF. Camelot no detecta columnas vacías, causando que todas las columnas posteriores (Tablets_Total, Counting_Delay, Validation_Delay) se desplacen una posición a la izquierda.

**Detección Inteligente:**
- ✅ Verifica si el albarán está cerrado (Definitive="Yes" + Counted_Date existe)
- ✅ Detecta si col 14 NO tiene códigos con sufijos [MALT] (indicador de Open vacía)
- ✅ Identifica si col 14 contiene un número simple (probablemente Tablets_Total desplazado)

**Solución Automática:**
- Inserta columna Open vacía en posición 14
- Desplaza todas las columnas posteriores a la derecha
- Mantiene la integridad de los datos sin modificar valores

**Ejemplo:**
```
❌ ANTES (columnas desplazadas):
Col 13: Total=4, Col 14: 5 (¿Open o Tablets_Total?), Col 15: 3

✅ DESPUÉS (corrección automática):
Col 13: Total=4, Col 14: (vacío - correcto), Col 15: 5, Col 16: 3
```

8. clean_open_tablets_when_closed()
Limpia basura residual en Open cuando el albarán está definitivamente cerrado
🎯 Casos de Uso

### Caso 1: Extracción de PDF (Último Día de Cierre) 🆕
**Escenario:** Fin de mes, todas las tablillas están cerradas, columna "Open" vacía en el PDF

```bash
1. Subir PDF en Tab "Extracción PDF"
2. Sistema detecta automáticamente:
   ✓ Método stream_standard (priorizado)
   ✓ Albaranes con Definitive="Yes"
   ✓ Columna Open vacía
3. Aplica corrección fix_missing_open_column():
   ✓ Inserta columna Open vacía
   ✓ Reubica Tablets_Total, Counting_Delay, Validation_Delay
4. Muestra datos correctos en tabla
5. Exportar a Excel profesional con 6 hojas
   ✓ Todas las columnas en posición correcta
   ✓ Sin desplazamientos
   ✓ Datos precisos para reportes
```

### Caso 2: Extracción Normal (Con Tablillas Abiertas)
```bash
1. Subir PDF en Tab "Extracción PDF"
2. Sistema prueba 6 métodos automáticamente
3. Selecciona el mejor método
4. Aplica las 8 correcciones en pipeline
5. Muestra validación con completitud %
6. Exportar a Excel profesional con 6 hojas
```
### Caso 3: Análisis de Tablillas
```bash
1. Extraer datos del PDF
2. Ir a Tab "Dashboard de Tablillas"
3. Ver métricas globales (Total, Abiertas, Cerradas)
4. Analizar por Warehouse y Cliente
5. Revisar discrepancias de integridad
6. Actuar sobre alertas inteligentes
```
### Caso 4: Análisis Histórico
```bash
1. Ir a Tab "Análisis Histórico"
2. Cargar múltiples archivos Excel (generados por la app)
3. Ver evolución temporal de tablillas
4. Analizar tendencias por warehouse
5. Comparar tasas de cierre entre fechas
6. Exportar consolidado con 3 hojas
```
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

### PDF no extrae correctamente
✅ **Solución:** Verificar que el PDF tenga la estructura esperada
- Probar diferentes métodos en las pestañas (stream_standard está priorizado)
- La v3.1 maneja automáticamente columnas vacías

### Columnas desalineadas (Especialmente en cierre de mes)
✅ **Solución v3.1:** Las 8 funciones de autocorrección ahora incluyen `fix_missing_open_column()`
- Detecta y corrige automáticamente cuando Open está vacía
- Reubica Tablets_Total, Counting_Delay, Validation_Delay
- Trabaja silenciosamente sin intervención manual

### Columna "Open" aparece con números sin códigos [MALT]
✅ **Solución:** Probablemente es desplazamiento de Tablets_Total
- La v3.1 detecta automáticamente este caso
- Verifica si Definitive="Yes" y Open no tiene sufijos [MALT]
- Aplica corrección automática

### Discrepancias en Total vs Open
ℹ️ **Normal:** Esto es ESPERADO cuando tablillas se cierran entre extracciones
- El sistema solo reporta, NO inventa códigos
- Discrepancias indican tablillas cerradas recientemente
- Revisar Dashboard de Tablillas para detalles

### Análisis histórico no funciona
✅ **Solución:** Asegurarse de usar archivos Excel generados por esta app (v3.1)
- Los archivos deben tener la hoja "Datos_Principales"
- Archivos de v3.0 son compatibles

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
v3.1 - Corrección Crítica para Cierre de Mes
Changelog v3.1 (Noviembre 2025)

### 🎯 Mejoras Críticas
✅ **Nueva corrección: fix_missing_open_column()** - Soluciona el problema #1 reportado en cierre de mes
✅ **Detecta columna Open completamente vacía** - Cuando todas las tablillas están cerradas
✅ **Corrige desplazamiento automático** - Reubica columnas sin intervención manual
✅ **Prioriza method_stream_standard** - Optimizado para manejar columnas vacías
✅ **Interfaz limpia** - Correcciones trabajan silenciosamente en segundo plano
✅ **100% automático** - Sin configuración adicional requerida

### 🔧 Mejoras de UX
- Removidos logs excesivos de debug
- Correcciones trabajan sin spam en interfaz
- Mejor priorización de métodos de extracción
- Mensajes de validación solo cuando son importantes

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
