# Especificación funcional — visor-vibraciones v2

> Documento de referencia compartido. Recoge el modelo acordado para la segunda
> versión del *Analizador Dinámico de Plataformas de Soporte para Equipos Rotativos*.
> Estado: **borrador de diseño** (previo a implementación).

---

## 1. Motivación

La v1 (`graficos.py`) clasifica la severidad vibratoria (Richart / Blake / ISO 20816-3)
de los puntos de apoyo de una plataforma, a partir de un análisis Steady-State de
SAP2000. Al revisarla surgieron hallazgos que invalidan el tratamiento del
**run-down (parada)** y que conviene corregir antes de ampliar la herramienta.

---

## 2. Modelo conceptual acordado

El equipo tiene dos regímenes distintos:

| Condición | Naturaleza | Frecuencia | Fuerza | Cómo se evalúa |
|-----------|-----------|-----------|--------|----------------|
| **Operación** | Estacionario (steady-state) fijo | `f_op` = 2900 rpm (48.33 Hz), **máxima** del equipo | `F_op` | FRF leída en `f_op` |
| **Transitorio** (partida / parada) | Barrido de frecuencia 0 ↔ `f_op` | Crítica en el **cruce resonante del aislador** ≈ 3–4 Hz | `F_rd` | FRF a lo largo del barrido; respuesta = FRF(f) × F(f) |

### Por qué el transitorio es a ~3–4 Hz y no en el peak estructural

- 3–4 Hz es el **modo de cuerpo rígido del equipo sobre los aisladores** (Novibra,
  K_din = 5600 N/mm por pata). Verificado: f ≈ 3.5 Hz ⇒ ~46 ton de equipo; f ≈ 4 Hz ⇒ ~35 ton.
- `F_rd` (5000 N V / 2500 N H) **no** es el desbalance estacionario escalado por ω²
  (eso daría ~2 N a 3.5 Hz). Es la **fuerza transmitida durante el cruce resonante**
  del aislador en la parada/partida.
- Por eso "el run-down está declarado a 3–4 Hz": ahí ocurre la amplificación, no en
  el peak estructural de ~48 Hz.

### Nota sobre conservadurismo (paso por resonancia)

Al cruzar **rápido** la resonancia de 3–4 Hz (la operación variable llega a régimen
"en poco tiempo"), la amplitud real es **menor** que la de resonancia estacionaria
(efecto *sweep-rate*, Lewis 1932). Evaluar `F_rd` como estacionario en `f_rd` es por
tanto **conservador** — aceptable para diseño.

---

## 3. Defectos de la v1 a corregir

1. **Run-down evaluado en el peak estructural** (`graficos.py:484-490`): usa `frf_pk`
   y `f_pk` (~48 Hz) en vez de la FRF a 3–4 Hz. → amplitud y velocidad equivocadas.
2. **Velocidad ISO inflada ~14×**: como `v = A·2π·f`, usar `f_pk≈48 Hz` en lugar de
   3.5 Hz sobreestima la velocidad del run-down por ~14×.
3. **Máscara fija 40–80 Hz** (`graficos.py:258`): descarta justamente la banda
   transitoria de baja frecuencia.
4. **Punto run-down graficado en ~2900 cpm** (Richart/Blake, `:628`, `:708`): debería
   ir en ~210 cpm (3.5 Hz). En log-log esto cambia la zona de severidad y hace ver el
   run-down más crítico de lo que es.

---

## 4. Requisitos funcionales v2

### 4.1 Condición de operación
- Sin cambios conceptuales: FRF leída en `f_op` (fijo), con `F_op`.

### 4.2 Condición transitoria (partida / parada)
- **Frecuencia `f_rd`**: calculada desde el aislador
  `f_rd = (1/2π)·√(K_din·g / W_pata)`, **editable** por el usuario.
- **Fuerza**, dos modos seleccionables:
  - **Valor fijo**: un `F_rd` constante (el del fabricante).
  - **Curva F(f)**: tabla fuerza–frecuencia, o modelo de desbalance
    `F(f) = m_u·e·(2πf)²`.
- **Evaluación**: idealmente la **envolvente** `respuesta(f) = FRF(f) × F(f)` sobre el
  barrido completo 0 → `f_op`, reportando el peor caso y marcando `f_rd`. Si solo se
  dispone de valor fijo, se evalúa la FRF en `f_rd` (punto).
- **Partida y parada**: mismas fuerzas por defecto, **editables por separado** para
  evaluarlas distinto.

### 4.3 Generalización ("que sirva para otros análisis a futuro")
Des-hardcodear:
- Banda de detección de peak (hoy 40–80 Hz fija).
- Mapeo de dirección por nombre de caso `_X/_Y/_Z`.
- Unidades / factor de conversión SAP2000 y `F_REF_N`.
- Zonas de clasificación Richart / Blake / ISO (parametrizables).

---

## 5. Requisitos de datos

- **Barrido SAP2000 de rango completo**: el Steady-State debe incluir **0 → ~50 Hz**
  (no solo 40–80 Hz), para que exista FRF en la banda transitoria.
  → `data/sap2000/`.
- **Peso del equipo** (total o por pata) → para `f_rd`. → `docs/memoria_calculo/`.
- **Aisladores**: modelo, cantidad, `K_din`. → `docs/datasheets/`.
- **(Opcional) Curva de fuerza dinámica** F(f) o desbalance `m_u·e`, si se usa el modo
  curva.

---

## 6. Pendientes / inputs por confirmar

- [ ] Peso del equipo (para fijar `f_rd`).
- [ ] Cantidad y modelo de aisladores.
- [ ] Nuevo barrido SAP2000 con baja frecuencia.
- [ ] ¿Se modelará la curva F(f), o solo valor fijo de `F_rd` en esta etapa?

---

## 7. Referencias

- ACI 351.3R — *Report on Foundations for Dynamic Equipment*
- Arya, O'Neill & Pincus (1979) — *Design of Structures and Foundations for Vibrating Machines*
- Richart, Hall & Woods (1970) — *Vibrations of Soils and Foundations*
- Bachmann et al. (1995) — *Vibration Problems in Structures*
- Lewis, F.M. (1932) — *Vibration during acceleration through a critical speed*, Trans. ASME
- Den Hartog — *Mechanical Vibrations*
- ISO 20816-3 / ISO 10816 — Evaluación de vibración mecánica
- Blake (1964) — Criterios de severidad vibratoria
