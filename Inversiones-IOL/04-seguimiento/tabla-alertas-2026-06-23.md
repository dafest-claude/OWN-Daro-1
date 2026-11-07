# Tabla de alertas y reglas — 2026-06-23

**Objetivo**: dejar concentrado en una sola tabla qué alertas hay que
crear, actualizar o eliminar en IOL, con los gatillos calculados.

**Importante**: los PMC son **estimados** a partir del histórico de
movimientos. Antes de crear las alertas, verificá el PMC real en
**IOL → Mi Cartera → Detalle de tenencia**. Si difiere en >2%
del estimado, usá el PMC real (no el estimado).

---

## 1. Tabla maestra de alertas

| # | Pos. | Precio 23/06 | PMC estimado | Δ vs PMC | Regla stop | Regla TP | **Stop** | **TP** | Estado en IOL |
|---|------|--------------|--------------|----------|------------|----------|----------|--------|----------------|
| 1 | **BMA** | 14.640 | 13.120 | +11.6% | −15% | +25% | **11.152** | **16.400** | Crear o verificar |
| 2 | **GGAL** | 7.970 | 7.350 | +8.4% | −15% | +25% | **6.248** | **9.188** | Crear o verificar |
| 3 | **TGSU2** | 9.415 | 9.461 | −0.5% | −15% | +25% | **8.042** | **11.826** | **ACTUALIZAR** |
| 4 | **GD30** | 99.320 | 96.069 | +3.4% | −8% | +15% | **88.384** | **110.479** | **ACTUALIZAR** |
| 5 | **AO27** | 154.900 | 149.480 | +3.6% | −8% | +15% | **137.522** | **171.902** | **Crear (nueva)** |
| 6 | **BPOC7** | 154.100 | 147.500 | +4.5% | −8% | +15% | **135.700** | **169.625** | **Crear (nueva)** |
| 7 | **NVDA** | 12.950 | 13.281 | −2.5% | −12% | +20% | **11.687** | **15.937** | **ACTUALIZAR** |
| 8 | **COPX** | 8.805 | 9.770 | −9.9% | −12% | +20% | **8.598** | **11.724** | Crear o vender (ver §3) |
| 9 | **IOLPORA** | 1.544 | 1.526 | +1.2% | −10% | n/a | **1.373** | n/a | **Crear (nueva)** |

> Reglas por tipo de instrumento ya definidas en informes previos:
> - **Acciones líderes ARG**: stop −15% / TP +25%.
> - **Bonos hard-dollar**: stop −8% / TP +15%.
> - **CEDEARs**: stop −12% / TP +20%.
> - **IOLPORA (cuasi RF)**: stop −10% / sin TP (es acumulación).

---

## 2. Posiciones SIN alerta (no necesitan)

| Pos. | Razón |
|------|-------|
| **PRCPPEB** | FCI Money Market — no tiene precio variable significativo, no requiere stop |
| **TZX26** | Vence el 30/06 — solo esperar |
| **VIST** | **Vendida el lunes 22/06** — si tenía alertas, **ELIMINAR** |

---

## 3. Caso especial — COPX

| Variable | Valor |
|----------|-------|
| Precio actual | 8.805 |
| PMC estimado | 9.770 |
| Pérdida actual | **−9.9%** |
| Distancia al stop (−12%) | solo **−2.1 pp** |

**COPX está a punto de tocar stop**. Una caída adicional del 2.1%
dispara la alerta.

### Tres opciones para COPX (decidir esta semana)

**Opción A — Vender ahora** (recomendada si no tenés tesis del cobre).
- Producido: ~ARS 17.610.
- Reasignar: 100% a PRCPPEB.
- Razón: la posición no tiene tesis documentada y está cerca del stop
  natural. Mejor vender en orden que esperar a que toque el stop.

**Opción B — Esperar al 30/06** y decidir junto con el aporte
mensual. Solo válido si **NO querés tomar acción esta semana**. Si
cae otro 2%, el stop disparará igual.

**Opción C — Mantener y documentar tesis** (cobre + transición
energética). Requiere:
- Documentar tesis en `01-estrategia/tesis-COPX.md`.
- Subir el stop a −15% (PMC × 0.85 = 8.305) para dar más margen.
- Aceptar drawdown adicional.

**Mi recomendación**: **Opción A**. Sin tesis, no vale el espacio
mental.

---

## 4. Verificación de PMC — paso obligatorio antes de crear alertas

Los PMC de la tabla son estimados. Para los 8 instrumentos que
requieren alerta, validar el PMC real así:

**Ruta en IOL**: *Mi Cartera → Detalle de tenencia* → ver columna
"PMC" o "Precio promedio de compra".

Si tu PMC real **difiere en más del 2% del estimado**, recalculá los
gatillos con esta fórmula:

- **Stop acciones ARG**: PMC real × 0.85
- **Stop bonos hard-dollar**: PMC real × 0.92
- **Stop CEDEARs**: PMC real × 0.88
- **Stop IOLPORA**: PMC real × 0.90
- **TP acciones**: PMC × 1.25
- **TP bonos**: PMC × 1.15
- **TP CEDEARs**: PMC × 1.20

---

## 5. Resumen de acciones — qué hacer en IOL

### Lista corta (orden de ejecución)

1. **Eliminar alertas de VIST** (si quedaron creadas, ya no aplican).
2. **Verificar PMC real** en IOL para las 8 posiciones de §1.
3. **Crear o actualizar las alertas** según la tabla:

#### Alertas a CREAR desde cero (probablemente nuevas)
- AO27 — Stop 137.522 + TP 171.902
- BPOC7 — Stop 135.700 + TP 169.625
- IOLPORA — Stop 1.373

#### Alertas a ACTUALIZAR (PMC cambió por compras del rebalanceo)
- GD30 — borrar viejas, crear con Stop 88.384 + TP 110.479
- NVDA — borrar viejas, crear con Stop 11.687 + TP 15.937
- TGSU2 — borrar viejas, crear con Stop 8.042 + TP 11.826

#### Alertas a VERIFICAR (no cambiaron PMC, pueden estar OK)
- BMA — Stop 11.152 + TP 16.400 — confirmar que existen
- GGAL — Stop 6.248 + TP 9.188 — confirmar que existen

**Total operación en IOL**: **17 alertas activas finales** (8 stops +
7 TPs + 2 verificaciones de BMA y GGAL).

### Tiempo estimado: 20–25 minutos

---

## 6. Cómo afecta esto al objetivo del 30/06

- **Las alertas no cambian el objetivo** (+7–8% sobre aportes para el
  30/06).
- **Sí cambian la gestión de riesgo**: si llega otro día como el 23/06
  con −0.5% y NVDA toca 11.687, la alerta dispara y se vende
  automáticamente, **protegiendo el resto del portfolio**.
- **Si no configurás las alertas**, una caída sostenida de 3 días
  como hoy nos lleva fácil a +3-4% sobre aportes (debajo de meta).

---

## 7. Posición por posición — riesgo actual

| Pos. | Distancia al stop | Distancia al TP | Riesgo |
|------|-------------------|-----------------|--------|
| BMA | −23.8% | +12.0% | Lejos de ambos |
| GGAL | −21.6% | +15.3% | Lejos de ambos |
| TGSU2 | −14.6% | +25.6% | Cerca de stop |
| GD30 | −11.0% | +11.2% | Equidistante |
| AO27 | −11.2% | +11.0% | Equidistante |
| BPOC7 | −11.9% | +10.1% | Equidistante |
| NVDA | −9.8% | +23.1% | Cerca de stop |
| **COPX** | **−2.3%** | +33.2% | **CRÍTICO** |
| IOLPORA | −11.1% | n/a | Lejos |

**Lectura**:
- **COPX**: crítico, distancia al stop −2.3%. **Decidir esta semana**.
- **NVDA, TGSU2**: cerca del stop (−10% y −14.6%). **Vigilar** si hay
  otro día rojo.
- **El resto**: tiene margen razonable.

---

## 8. Bottom line de las alertas

**Operación más urgente del jueves 25/06**:
1. Verificar PMC reales (5 min).
2. Crear/actualizar las 17 alertas según tabla (20 min).
3. **Decidir COPX** (vender hoy o documentar tesis).

**Una vez listo esto**, la gestión de riesgo queda automatizada y
podés dormir tranquilo hasta el 30/06.
