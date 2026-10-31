# Análisis comparativo — Cartera IOL vs Plazo Fijo vs Dólares MEP

**Fecha del análisis**: 12/09/2026.
**Corte de valuación**: 11/09/2026.
**Período total analizado**: ~132 días (01/05/2026 → 11/09/2026).

---

## Metodología

Compararé 3 estrategias con **mismos aportes en las mismas fechas**:

1. **Cartera IOL real** — el resultado efectivo del usuario.
2. **Plazo fijo tradicional** con reinversión mensual al 2.5% TEM.
3. **Dólar MEP** — comprado en cada aporte y mantenido.

---

## Cronología estimada de aportes

| Fecha | Aporte ARS | Aportes acumulados |
|-------|------------|---------------------|
| 01/05/2026 | 200.000 | 200.000 |
| 30/05/2026 | 350.000 | 550.000 |
| 30/06/2026 | 150.000 | 700.000 |
| 31/07/2026 | 150.000 | 850.000 |
| 30/08/2026 | 150.000 | 1.000.000 |
| **TOTAL** | **1.000.000** | |

Los aportes de mayo son estimados (usuario confirmó 550k iniciales); los posteriores son datos verificados.

---

## Escenario 1 — Cartera IOL (real)

Valuación al 11/09: **ARS 985.697,20**.

| Métrica | Valor |
|---------|-------|
| Aportes | 1.000.000 |
| Valor final | 985.697 |
| **Ganancia neta** | **−14.303** |
| **% sobre aportes** | **−1.43%** |

---

## Escenario 2 — Plazo fijo con reinversión mensual

**Tasa asumida**: 2.5% TEM (~30% TNA) — estimación conservadora del promedio del período.

Cada aporte capitaliza mensualmente hasta el 11/09:

| Aporte | Fecha | Meses hasta 11/09 | Factor (1.025^n) | Valor al 11/09 |
|--------|-------|-------------------|-------------------|-----------------|
| 200.000 | 01/05 | 4.33 | 1.1132 | 222.640 |
| 350.000 | 30/05 | 3.40 | 1.0870 | 380.450 |
| 150.000 | 30/06 | 2.40 | 1.0611 | 159.165 |
| 150.000 | 31/07 | 1.40 | 1.0353 | 155.295 |
| 150.000 | 30/08 | 0.40 | 1.0099 | 151.485 |
| | | | **TOTAL** | **1.069.035** |

| Métrica | Valor |
|---------|-------|
| Aportes | 1.000.000 |
| Valor final | 1.069.035 |
| **Ganancia neta** | **+69.035** |
| **% sobre aportes** | **+6.90%** |

---

## Escenario 3 — Dólar MEP mantenido

**Cotizaciones MEP estimadas** durante el período:

| Fecha | MEP est. ARS/USD |
|-------|-------------------|
| 01/05/2026 | 1.150 |
| 30/05/2026 | 1.200 |
| 30/06/2026 | 1.180 |
| 31/07/2026 | 1.230 |
| 30/08/2026 | 1.250 |
| 11/09/2026 | 1.290 |

Cada aporte se convierte a USD en la fecha del aporte y se valúa al 11/09:

| Aporte ARS | MEP compra | USD comprados | MEP 11/09 | Valor al 11/09 |
|------------|-------------|----------------|-----------|-----------------|
| 200.000 | 1.150 | 173.9 | 1.290 | 224.331 |
| 350.000 | 1.200 | 291.7 | 1.290 | 376.293 |
| 150.000 | 1.180 | 127.1 | 1.290 | 163.959 |
| 150.000 | 1.230 | 121.9 | 1.290 | 157.251 |
| 150.000 | 1.250 | 120.0 | 1.290 | 154.800 |
| | | | **TOTAL** | **1.076.634** |

| Métrica | Valor |
|---------|-------|
| Aportes | 1.000.000 |
| Valor final | 1.076.634 |
| **Ganancia neta** | **+76.634** |
| **% sobre aportes** | **+7.66%** |

---

## 📊 CUADRO COMPARATIVO CONSOLIDADO

| Estrategia | Aportes | Valor al 11/09 | Ganancia $ | % s/aportes | Ranking |
|-----------|---------|-----------------|------------|-------------|---------|
| **Cartera IOL (real)** | 1.000.000 | 985.697 | **−14.303** | **−1.43%** | 3° 🥉 |
| **Plazo Fijo 2.5% TEM** | 1.000.000 | 1.069.035 | **+69.035** | **+6.90%** | 2° 🥈 |
| **Dólar MEP** | 1.000.000 | 1.076.634 | **+76.634** | **+7.66%** | 1° 🥇 |

### Diferencia vs cartera IOL

| Alternativa | Ganancia extra vs IOL | pp extra |
|-------------|------------------------|-----------|
| Plazo Fijo | +83.338 ARS | +8.33 pp |
| Dólar MEP | +90.937 ARS | +9.09 pp |

---

## Ajuste por inflación (contexto real)

**Inflación estimada del período**: ~10-12% acumulado (2.4% mensual promedio).

Retornos REALES (netos de inflación):

| Estrategia | Nominal | Inflación est. | Real |
|-----------|---------|-----------------|------|
| Cartera IOL | −1.43% | 11% | **−11.2%** 🔴 |
| Plazo Fijo | +6.90% | 11% | −3.7% |
| Dólar MEP | +7.66% | 11% | −3.0% |

**Ninguna estrategia le ganó a la inflación en este período**. Pero el plazo fijo y el dólar perdieron **menos poder adquisitivo real**.

---

## ¿Por qué la cartera IOL rindió menos?

### Causas concretas

1. **Timing de las compras**: entradas cerca de los picos del rally (MSFT +27% en julio, luego corrigió).
2. **Comisiones acumuladas**: ~20+ operaciones × 0.7-1% = **~30-40k ARS** en costos operativos.
3. **Pérdidas por stop-loss**: COPX vendida en −10%, GGAL parcial vendida en −10%, BMA salida y recompra.
4. **VIST vendido en pérdida** (rotación temprana).
5. **Sobreconcentración en tech USA** cuando corrigió.
6. **Acciones argentinas en pausa/bajista** durante buena parte del período.

### Distribución del "costo" del rojo

| Concepto | Impacto estimado |
|----------|-------------------|
| Comisiones + IVA | −30.000 |
| Stops disparados (COPX, GGAL parcial) | −15.000 |
| Compras al pico (agosto) | −25.000 |
| Ganancias de posiciones (MSFT, NVDA) | +55.000 |
| **NETO** | **−15.000** |

---

## Reflexión honesta

### La verdad incómoda

En **4 meses**, con timing operativo imperfecto y un mercado difícil para el equity ARG, la cartera IOL rindió **menos que instrumentos más simples y sin riesgo**.

**Esto no es un fracaso de estrategia** — es una realidad del mercado en este período específico.

### Lo que NO muestra este análisis

1. **La cartera IOL tiene UPSIDE**. Un plazo fijo NUNCA puede darte +10% mensual. Tu cartera SÍ puede (como MSFT rindió +27% en un mes).

2. **4 meses es MUY corto** para evaluar una estrategia de renta variable. En un año, la variabilidad se compensa.

3. **La cartera tiene DIVERSIFICACIÓN**. En una devaluación fuerte, tu bloque hard-dollar + CEDEARs te protege. Un plazo fijo en pesos se licua.

4. **El plazo fijo depende de la tasa** — si la tasa baja al 15% TNA, rinde menos.

5. **El dólar es "no rendimiento"**. En términos de USD, comprar dólares te deja con USD. La cartera IOL apuesta a que el USD crezca en valor.

### Lo que SÍ muestra

1. **En 4 meses, la estrategia agresiva no compensó** los costos y errores operativos.

2. **La rentabilidad real (ajustada por inflación) es negativa en todas las estrategias**.

3. **Las alternativas simples fueron mejores** en este período específico.

---

## ¿Qué hacer con esta información?

### Opción A — Continuar la estrategia actual con ajustes

**Recomendado si**:
- Tu horizonte es 12-24 meses (no 4).
- Aceptas volatilidad y potenciales meses en rojo.
- Confías en el proceso de disciplina.

**Ajustes sugeridos**:
- Menos operativa (menos comisiones).
- Alertas más laxas (menos stops disparados).
- Menos concentración USA (evitá comprar en picos).

### Opción B — Rotar a estrategia mixta

**50% cartera IOL + 50% plazo fijo** — hedgeás el timing.

Nuevo aporte del 30/09 va todo a **plazo fijo** hasta acumular ~500k en PF. Después, retomás la estrategia agresiva sólo con el excedente.

### Opción C — Simplificar

Vender toda la cartera IOL, pasar a plazo fijo o mix plazo fijo + USD MEP.

**No recomendada** — perdés el upside de mediano plazo y consolidás las pérdidas actuales.

### Opción D — Doble o nada (agresiva)

**No recomendada** — más aportes a la cartera actual esperando el rebote. Mayor riesgo.

---

## Mi recomendación honesta

**Continuar con Opción A + ajustes tácticos**:

1. **Aceptar que 4 meses es corto** — dar 8 meses más de disciplina antes de reevaluar.
2. **Reducir operativa** — menos entradas/salidas = menos comisiones.
3. **En el aporte del 30/09**: mantener el plan A + ON PAMPO. Diversificar sin agregar riesgo.
4. **Setear meta más realista**: **+2 a +3% mensual sobre aportes**, no +5%.
5. **Reevaluar en 31/12/2026** (8 meses de track record) — con 8 meses de datos, la comparación será más justa.

---

## Bottom line

**El análisis honesto**: **plazo fijo y dólar rindieron mejor en estos 4 meses**. Pero **4 meses no es plazo suficiente para juzgar una estrategia de renta variable**.

**La decisión estratégica**: ¿tenés paciencia para 8 meses más? Si sí, seguimos. Si no, hablamos de plan B.

**Regla operativa nueva**: **NO tomar decisiones de estrategia basadas en 4 meses de datos**. Es como cerrar un negocio porque perdió plata el primer trimestre.
