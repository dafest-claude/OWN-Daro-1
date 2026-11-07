# Plan agresivo realista — 3 meses — meta en ARS

## Parámetros confirmados

| Parámetro                | Valor                                |
|--------------------------|--------------------------------------|
| Capital inicial          | ARS 250.000 (rango 200–300k)         |
| Aporte mensual           | ARS 150.000 (rango 100–200k)         |
| Horizonte                | 3 meses                              |
| Moneda de medición       | ARS                                  |
| Retorno objetivo total   | +50% a +80% (a 3 meses)              |
| Retorno mensual implícito| ~15–22% compuesto                    |
| Drawdown tolerable       | 20–50%                               |
| Frecuencia operativa     | Diaria                               |
| Plataforma               | IOL                                  |

## Filosofía

Como la meta se mide en ARS, no necesito dolarizar todo. Pero **mantengo
~30% dolarizado como cobertura** contra un evento de devaluación que
licúe los retornos en pesos. El resto se asigna a activos argentinos de
alta beta (acciones líderes, bonos en pesos) y CEDEARs líquidos.

No uso opciones ni apalancamiento con caución en V1 — el objetivo
50–80% en 3 meses es alcanzable sin esos vehículos y el ratio
riesgo/recompensa de meterlos no compensa cuando el drawdown tolerable
es 50%, no 100%.

## Cartera modelo — asignación inicial (ARS 250.000)

| # | Bloque                          | %   | Monto ARS | Instrumentos (ticker IOL)        |
|---|---------------------------------|-----|-----------|----------------------------------|
| 1 | Acciones líderes Merval         | 40% | 100.000   | GGAL 15% · YPFD 10% · PAMP 8% · BMA 7% |
| 2 | Bonos soberanos hard-dollar     | 25% |  62.500   | GD30 15% · AL30 10%              |
| 3 | CEDEARs alta beta               | 20% |  50.000   | NVDA 8% · TSLA 6% · QQQ 6%       |
| 4 | Renta fija pesos corta          | 10% |  25.000   | LECAP corta o FCI Money Market   |
| 5 | Caución colocadora (liquidez)   |  5% |  12.500   | Caución 1–7 días                 |

**Por qué cada bloque:**

- **Acciones líderes (40%)**: motor principal. En ciclos pro-mercado el
  Merval rinde 8–15% mensual en USD; en pesos puede dar más por
  efecto MEP. GGAL/BMA capturan compresión de riesgo país; YPFD/PAMP
  capturan suba de energía y dolarización tarifaria.
- **Bonos hard-dollar (25%)**: cobertura cambiaria + upside por
  compresión de spread soberano. GD30/AL30 son los más líquidos.
- **CEDEARs alta beta (20%)**: descorrelación con Argentina. NVDA y
  TSLA tienen volatilidad >40% anual; QQQ es el ancla.
- **Renta fija pesos (10%)**: colchón con tasa positiva en términos
  reales esperada. Sirve para promediar en caídas.
- **Caución (5%)**: liquidez para oportunidades intradiarias.

## Reglas de operación

### Entradas
- Comprar en **2 tramos**: 50% el día de armado, 50% repartido en los
  primeros 5 días hábiles (DCA corto). Reduce timing risk.
- No entrar después de gaps de +5% diario en una acción — esperar
  retroceso.

### Stops y toma de ganancia
| Posición          | Stop loss | TP parcial (vender 1/3) | TP final |
|-------------------|-----------|-------------------------|----------|
| Acciones líderes  | −15%      | +25%                    | +50%     |
| CEDEARs           | −12%      | +20%                    | +40%     |
| Bonos hard-dollar | −8%       | +15%                    | +25%     |
| Renta fija pesos  | —         | —                       | rolar    |

Al hacer TP parcial: mover stop a punto cero (break-even) en el resto.

### Rebalanceo
- **Cada lunes** (revisión semanal).
- Si una posición pasa el 50% de su bloque (ej. GGAL supera 20% del
  total cuando le tocaba 15%), vender el exceso y reasignar al bloque
  rezagado.

### Aportes mensuales
- 60% al bloque más rezagado (comprar abajo).
- 40% al bloque que mantiene momentum (no cortar ganadoras).

## Calendario de seguimiento

| Día            | Acción                                                |
|----------------|-------------------------------------------------------|
| Lunes          | Rebalanceo semanal, revisión de stops                 |
| Miércoles      | Lectura de licitación Tesoro (si la hay)              |
| Jueves         | Revisión de cartera vs. Merval (alpha check)          |
| Cierre de mes  | Cálculo de retorno, ajuste de plan, aporte mensual    |

## Eventos macro a monitorear

- **IPC INDEC** (mensual, ~13 hábil): inflación.
- **Licitaciones Tesoro**: rollover de deuda pesos, impacta LECAP.
- **Comunicados BCRA**: tasa de política, restricciones cambiarias.
- **Brecha MEP/CCL vs. oficial**: si se abre >40%, alerta de evento
  cambiario.
- **Datos de actividad**: EMAE, recaudación, balanza comercial.

## Salidas de emergencia (kill-switches)

Salir a caución/MM y reevaluar si pasa cualquiera de estos:

1. Drawdown total de la cartera **>30%**.
2. Brecha MEP/oficial sube **>15% en 5 días**.
3. Suba de tasa BCRA **>10 puntos en una sola decisión**.
4. Movimiento del MEP **>10% en un día**.

## KPI de éxito por mes

| Mes | Capital esperado | Acumulado vs inicial | Aporte ese mes |
|-----|------------------|----------------------|----------------|
| 0   | 250.000          | 0%                   | —              |
| 1   | 287.500–305.000  | +15% a +22%          | +150.000       |
| 2   | 503.000–555.000  | +30% a +44%¹         | +150.000       |
| 3   | 730.000–845.000  | +50% a +80%¹         | —              |

¹ Calculado **sobre el aporte total acumulado** (250k + 150k + 150k = 550k),
no sobre el inicial. El % refleja el rendimiento real del capital invertido.

## Próximos archivos

- `02-instrumentos/` — ficha técnica por cada ticker.
- `03-operativa-IOL/` — paso a paso de cómo cargar cada orden.
- `04-seguimiento/registro-diario.md` — planilla de tracking.
- `05-escenarios/simulacion.md` — Monte Carlo simple (bull/base/bear).
