# Simulación de escenarios — 3 meses

Tres escenarios sobre la cartera V1 con capital inicial ARS 250.000
y aportes mensuales de ARS 150.000 al cierre de los meses 1 y 2.

## Hipótesis de retorno mensual por bloque

| Bloque                  | %   | Bull mensual | Base mensual | Bear mensual |
|-------------------------|-----|--------------|--------------|--------------|
| Acciones líderes        | 40% | +28%         | +15%         | −12%         |
| Bonos hard-dollar       | 25% | +18%         | +8%          | −10%         |
| CEDEARs alta beta       | 20% | +20%         | +10%         | −18%         |
| Renta fija pesos corta  | 10% | +6%          | +5%          | +4%          |
| Caución                 |  5% | +5%          | +4%          | +3%          |

**Retorno mensual ponderado:**
- Bull: ~20.4%
- Base: ~10.7%
- Bear: ~−10.6%

## Escenario BULL (mercado favorable)

| Mes | Capital inicio | Retorno mes | Capital fin | Aporte | Total fin |
|-----|----------------|-------------|-------------|--------|-----------|
| 1   | 250.000        | +20.4%      | 301.000     | 150.000| 451.000   |
| 2   | 451.000        | +20.4%      | 542.964     | 150.000| 692.964   |
| 3   | 692.964        | +20.4%      | 834.328     | —      | **834.328** |

- Capital aportado total: 550.000
- Resultado neto: +51.7% sobre aportes ⇒ **dentro de la banda 50–80%** ✅

## Escenario BASE (mercado neutro/positivo moderado)

| Mes | Capital inicio | Retorno mes | Capital fin | Aporte | Total fin |
|-----|----------------|-------------|-------------|--------|-----------|
| 1   | 250.000        | +10.7%      | 276.750     | 150.000| 426.750   |
| 2   | 426.750        | +10.7%      | 472.412     | 150.000| 622.412   |
| 3   | 622.412        | +10.7%      | 689.011     | —      | **689.011** |

- Capital aportado total: 550.000
- Resultado neto: +25.3% sobre aportes (debajo de la meta)
- **Acción**: en este escenario considerar agregar lanzamientos cubiertos
  sobre las posiciones de acciones líderes para sumar 1.5–3% mensual
  extra (V2 del plan).

## Escenario BEAR (mercado adverso)

| Mes | Capital inicio | Retorno mes | Capital fin | Aporte | Total fin |
|-----|----------------|-------------|-------------|--------|-----------|
| 1   | 250.000        | −10.6%      | 223.500     | 150.000| 373.500   |
| 2   | 373.500        | −10.6%      | 333.909     | 150.000| 483.909   |
| 3   | 483.909        | −10.6%      | 432.614     | —      | **432.614** |

- Capital aportado total: 550.000
- Resultado neto: **−21.3%** sobre aportes
- Drawdown peak-to-trough: ~25–30% en algún punto del recorrido
- **Sigue dentro del rango tolerable (20–50%)** ✅
- Acción: si en el mes 1 hay drawdown >20%, ejecutar kill-switch del
  plan (rotar 50% de acciones a caución) y reevaluar.

## Tabla resumen

| Escenario | Capital fin 3M | % sobre aporte | Probabilidad subjetiva |
|-----------|----------------|----------------|------------------------|
| Bull      | 834.328        | +51.7%         | 25%                    |
| Base      | 689.011        | +25.3%         | 50%                    |
| Bear      | 432.614        | −21.3%         | 25%                    |

**Esperanza matemática** (probabilidad × resultado):
0.25 × 834k + 0.50 × 689k + 0.25 × 432.6k = **661k** ≈ +20% sobre aportes.

## Lectura honesta

El objetivo declarado (50–80%) **solo se cumple en el escenario bull**.
En el escenario base —que es el más probable a juicio del plan— el
resultado es +25%, abajo de la meta. Para acercar la esperanza
matemática a la meta, las opciones son:

1. **Subir riesgo**: agregar venta de puts cubiertos sobre CEDEARs
   (suma ~2–4%/mes en bull/base, pierde en bear) → plan V2.
2. **Apalancar con caución tomadora**: comprar bonos con plata
   prestada al 100–110% del capital (multiplica ganancias y pérdidas).
3. **Aceptar que la meta es ambiciosa** y ejecutar V1 con disciplina,
   sabiendo que el caso base te deja en +25%.

Mi recomendación es **arrancar con V1 disciplinado, evaluar al mes 1**,
y recién entonces decidir si pasás a V2 con opciones.
