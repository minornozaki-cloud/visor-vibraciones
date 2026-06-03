# Proyecto 730 — Diagnóstico y Medición de Vibraciones (CMP)
## Plataforma de acero elevada con máquinas centrífugas + aisladores elastoméricos

---

## Descripción del proyecto

Análisis dinámico de una plataforma de acero elevada que soporta **tres máquinas centrífugas** (9.8 ton c/u) operando a **2900 RPM → f_op = 48.33 Hz**. Cada máquina tiene cuatro puntos de apoyo con aisladores elastoméricos **Trelleborg Novibra RAEM2500** (custom, durómetro 70, caucho butil).

**Objetivo**: Caracterizar el desempeño de la aislación vibratoria y asegurar un montaje seguro y efectivo.

**Unidades en todo el modelo SAP2000**: `kgf / cm`

---

## Hallazgo central (crítico)

Existe un **problema grave de razón de rigideces** (criterio de Hutchinson):

```
RF = K_estructura_local / K_aislador_dinámico
RF = 0.159 kN/mm / 5.6 kN/mm = 0.028   <<< debe ser ≥ 10
```

La estructura es ~35 veces más flexible que cada aislador. Consecuencias:
- Los aisladores no pueden cumplir su función de diseño
- La rigidez efectiva del sistema es dominada por la estructura (`K_ef ≈ 0.155 kN/mm`)
- La aislación a velocidad de operación es accidental, no de diseño
- El cruce de resonancia en arranque/parada representa un riesgo crítico

---

## Parámetros clave del sistema

| Parámetro | Valor |
|---|---|
| Masa por máquina | 9,800 kg |
| Velocidad operación | 2900 RPM |
| Frecuencia operación (f_op) | 48.33 Hz |
| Puntos de apoyo por máquina | 4 |
| Total nodos de apoyo | 12 (C11–C14, C21–C24, C31–C34) |
| Aislador | Trelleborg Novibra RAEM2500 |
| K_static aislador | 2.8 kN/mm |
| K_dynamic aislador | 5.6 kN/mm |
| K_estructura_local (unit load SAP2000) | 0.159 kN/mm (U3=0.626 cm con 1 ton) |
| RF (razón de rigideces) | 0.028 — CRÍTICO |

**Nota metodológica importante**: Cada aislador ve únicamente su propia rigidez estructural *local* (no rigideces en paralelo), ya que los aisladores no están modelados en SAP2000 y cada uno actúa como dispositivo independiente.

---

## Archivos de datos del proyecto

### `DEFORMACIONES_MODAL.xlsx`
- **Contenido**: Desplazamientos absolutos de juntas — análisis modal (`MODAL-1`)
- **Hoja**: `Joint Displacements`
- **Juntas monitoreadas (12 nodos de apoyo)**:
  - Máquina 1: `C11`, `C12`, `C13`, `C14`
  - Máquina 2: `C21`, `C22`, `C23`, `C24`
  - Máquina 3: `C31`, `C32`, `C33`, `C34`
- **Columnas**: `Joint, OutputCase, CaseType, StepType, StepNum, U1[cm], U2[cm], U3[cm], R1[rad], R2[rad], R3[rad]`
- **Modos**: 1 a 800 (shape: 9602 filas × 11 cols, header en fila 3)
- **Uso principal**: Identificar qué modo maximiza `U3` en los nodos de apoyo → frecuencia natural local

### `MASS_RATIOS.xlsx`
- **Contenido**: Razones de masa modal participante
- **Hoja**: `Modal Participating Mass Ratios`
- **Columnas**: `OutputCase, StepType, StepNum, Period[s], UX, UY, UZ, SumUX, SumUY, SumUZ, RX, RY, RZ, SumRX, SumRY, SumRZ`
- **Modos**: 150 modos (shape: 152 filas × 16 cols, header en fila 3)
- **Rango de frecuencias**: 2.35 Hz (modo 150) a 11.04 Hz (modo 1) — rangos globales
- **Uso principal**: Verificar participación de masa modal; complemento a DEFORMACIONES_MODAL

### `STDS_COMPLETO.xlsx`
- **Contenido**: Análisis Steady-State (respuesta armónica forzada) en 3 direcciones
- **Hoja**: `Joint Displacements`
- **Casos**: `STST_X`, `STST_Y`, `STST_Z` (barrido en frecuencia por dirección)
- **Rango de frecuencias**: **30 a 80 Hz**, 151 puntos de frecuencia (∆f ≈ 0.333 Hz)
- **Mismas 12 juntas de apoyo**: C11–C34
- **Columnas**: igual que DEFORMACIONES_MODAL
- **Uso principal**: FRF (Función de Respuesta en Frecuencia) — identificar resonancias, calcular amplitudes en f_op, calcular velocidades de vibración

---

## Lectura estándar de los archivos Excel

```python
import pandas as pd

COL_NAMES = ['Joint','OutputCase','CaseType','StepType','StepNum',
             'U1','U2','U3','R1','R2','R3']

def read_sap_table(filepath, sheet=0):
    """Lee tablas exportadas de SAP2000 (header en fila 3, datos desde fila 4)."""
    df = pd.read_excel(filepath, sheet_name=sheet, header=2)
    df.columns = COL_NAMES
    df = df.dropna(subset=['Joint'])
    df['StepNum'] = pd.to_numeric(df['StepNum'], errors='coerce')
    for col in ['U1','U2','U3','R1','R2','R3']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

# Ejemplos de uso
modal   = read_sap_table('DEFORMACIONES_MODAL.xlsx')
stds    = read_sap_table('STDS_COMPLETO.xlsx')
# MASS_RATIOS tiene columnas diferentes — leer por separado
```

---

## Análisis pendientes / próximos pasos

### 1. Extracción de frecuencia natural local (PENDIENTE)
Metodología definida:
1. Para cada modo, encontrar el máximo `|U3|` entre los 12 nodos de apoyo
2. El modo con mayor `U3` en los nodos de apoyo → su frecuencia = `f_natural_local`
3. Comparar `f_natural_local` con `f_op = 48.33 Hz` para evaluar margen de separación

```python
# Lógica básica
support_nodes = ['C11','C12','C13','C14','C21','C22','C23','C24','C31','C32','C33','C34']
modal_support = modal[modal['Joint'].isin(support_nodes)].copy()
modal_support['U3_abs'] = modal_support['U3'].abs()
idx_max = modal_support.groupby('StepNum')['U3_abs'].max().idxmax()
modo_critico = modal_support[modal_support['StepNum'] == idx_max]
```

### 2. FRF desde STDS_COMPLETO
- `FRF [mm/ton] = |U3| * 10` (cm → mm, por 1 ton aplicada)
- Resonancia confirmada si: pico en zona de exclusión + `|fase − 90°| < 15°` + `FRF > 0.05 mm/ton`
- Velocidad: `v_peak = 2π × f × A`; `v_RMS = v_peak / √2`
- Comparar contra **ISO 10816 / ISO 20816**

### 3. Datos pendientes de Trelleborg
- Razón de amortiguamiento `ξ` para RAEM2500 custom
- Confirmar dependencia frecuencial de K_static vs K_dynamic

### 4. Análisis armónico con Link elements en SAP2000
- Requiere: datos de desbalance de máquina (`m·e`)
- Modelar aisladores como Link elements
- Barrido de frecuencia 20–80 Hz (actual empieza en 30 Hz — revisar si faltan modos entre 20–30 Hz)

### 5. Integración con mediciones de campo
- Comparar resultados teóricos vs mediciones ISO 10816 / ISO 20816 ya disponibles

---

## Metodología unit load (rigidez estructural local)

En SAP2000, se aplica **1 ton en U3** en cada nodo de apoyo individualmente y se lee la deflexión `U3` en ese mismo nodo:

```
K_local = F / U3 = 1 ton / 0.626 cm = 1.597 ton/cm = 15.97 kN/cm = 0.159 kN/mm
```

**Error conceptual a evitar**: NO sumar rigideces en paralelo. Cada aislador ve su propia rigidez local, no la suma de los cuatro apoyos.

---

## App Streamlit (en desarrollo)

- **Ruta local**: `c:\Users\mnozaki\OneDrive - MACROSTEEL INGENIERIA SPA\0.- PROYECTOS\730.- Diagnóstico y medición de vibraciones – CMP\PLANILLAS DE CÁLCULO\GRÁFICOS.py`
- **Dependencias**: `streamlit`, `pandas`, `numpy`, `matplotlib`, `plotly`, `python-docx`, `openpyxl`
- **Ejecución**: `streamlit run GRÁFICOS.py` (desde terminal, no F5)
- **Entradas esperadas**: archivos de SAP2000 (Joint Displacements en frecuencia — Mag/Real/Imag)

---

## Principios guía del análisis

| Concepto | Regla |
|---|---|
| Criterio de Hutchinson | `RF = K_estructura_local / K_aislador_dinámico ≥ 10` |
| Rigidez local vs global | Usar rigidez puntual por nodo, no combinada en paralelo |
| Modo dominante en apoyos | El modo que maximiza `U3` en los 12 nodos de apoyo |
| Unidades SAP2000 | Siempre `kgf / cm` |
| Benchmarks de vibración | ISO 10816 / ISO 20816 |
