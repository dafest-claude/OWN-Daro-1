# Cartera Starter (V0) — 4 instrumentos en IOL

Versión simplificada para arrancar. Cada instrumento cumple un rol
distinto y no hay redundancia. Si V0 funciona después de 3–4 semanas,
escalamos a V1 (10 instrumentos del plan original).

## Por qué solo 4 instrumentos

- **Menos comisiones de armado** (4 órdenes en vez de 10).
- **Foco**: podés seguir cada posición en serio todos los días.
- **Aprendizaje**: en 2–3 semanas vas a ver claro cómo se comporta cada
  uno contra tu tolerancia real al riesgo (que muchas veces solo se
  conoce cuando estás adentro).
- **Trade-off**: menos diversificación. Por eso los stops se respetan
  más estrictamente que en V1.

## Asignación inicial — capital ARS 250.000

| # | Ticker | Bloque que cubre        | %   | Monto ARS | Rol en la cartera |
|---|--------|-------------------------|-----|-----------|-------------------|
| 1 | GGAL   | Acciones líderes ARG    | 30% |  75.000   | Motor argentino — beta alta vs Merval |
| 2 | GD30   | Bonos soberanos USD     | 30% |  75.000   | Cobertura cambiaria + upside spread |
| 3 | NVDA   | CEDEAR tech USA         | 25% |  62.500   | Descorrelación + dólar implícito |
| 4 | FCI MM | Cash con tasa           | 15% |  37.500   | Liquidez para promediar y reserva |

**Total invertido**: 100%
**FCI sugerido**: IOL Ahorro Plus (rescate T+0).

## Por qué estos 4 y no otros

### GGAL (no GGAL+YPFD+PAMP+BMA juntos)
Es la acción más líquida del panel y la que tiene beta más limpia
contra el ciclo macro argentino. Concentrar en una sola te ahorra 3
órdenes y mantiene exposición al bloque "ARG equity". Si Merval sube
20%, GGAL típicamente sube 22–28%.

### GD30 (no GD30+AL30)
GD30 es ley NY y tiene más liquidez que AL30. El diferencial AL30/GD30
es un trade de nicho que no aporta nada en V0.

### NVDA (no NVDA+TSLA+QQQ)
NVDA es el CEDEAR de mayor volumen y el más correlacionado con el ciclo
IA, que es el driver dominante de Nasdaq hoy. TSLA suma volatilidad sin
mejor relación riesgo/recompensa. QQQ es más ancla que motor.

### FCI Money Market (no LECAP)
El FCI rescata T+0; la LECAP te ata al vencimiento o te obliga a vender
en mercado secundario con riesgo de tasa. Para una reserva de 15% que
puede necesitar moverse rápido, el FCI gana.

## Reglas de entrada

- **Día 0**: comprar 50% de cada posición.
- **Días 1–5**: completar el otro 50% en 3 tramos.
- Si en alguno de esos días el activo sube >5%, esperar retroceso.
- El FCI se suscribe el día 0 con el 100% (no necesita DCA).

### Órdenes concretas (precios a confirmar en pantalla el día de la operación)

| Ticker | Tramo 1 (Día 0) | Tramo 2 (días 1–5) |
|--------|-----------------|---------------------|
| GGAL   | ARS 37.500      | ARS 37.500          |
| GD30   | ARS 37.500      | ARS 37.500          |
| NVDA   | ARS 31.250      | ARS 31.250          |
| FCI MM | ARS 37.500      | —                   |

## Reglas de salida (stops y TP)

| Posición | Stop loss | TP parcial (vender 1/3) | TP final |
|----------|-----------|-------------------------|----------|
| GGAL     | −15%      | +25% (mover stop a 0)   | +50%     |
| GD30     | −8%       | +15%                    | +25%     |
| NVDA     | −12%      | +20%                    | +40%     |
| FCI MM   | —         | —                       | —        |

> IOL no tiene stop loss automático. **Operativa manual**: dejar orden
> limitada de venta a precio stop con vigencia día y renovarla cada
> mañana, o configurar alerta por precio en la app.

## Rebalanceo semanal (lunes)

1. Mirar % real de cada posición sobre el total.
2. Si alguna pasó +5pp de su peso objetivo (ej. GGAL al 35%), vender el
   exceso y pasarlo al FCI.
3. Si alguna cayó −5pp y no rompió stop, comprar desde el FCI para
   volverla al peso objetivo (promediar a la baja **solo si no rompió
   stop**).

## Kill-switch (corte de emergencia)

Vender todo a FCI/caución y reevaluar si:

- Drawdown total de cartera **>25%** (en V0 es más estricto que V1
  porque hay menos diversificación).
- Movimiento del MEP **>10% en un día**.
- Cualquier posición individual rompe stop y al revisar la siguiente
  rueda, una segunda posición también rompe stop. Eso significa que
  el régimen del mercado cambió.

## KPI para decidir el salto a V1

Después de 3–4 semanas con V0, si pasan **2 de estos 3**, escalamos a
V1 (suma YPFD/PAMP/BMA, AL30, TSLA/QQQ, LECAP):

1. Rendimiento acumulado **≥ +10%**.
2. Ninguna posición rompió stop por error operativo (que el stop se
   ejecutó cuando correspondía).
3. Pudiste seguir la cartera todos los días sin estrés.

Si **2 de los 3 fallan**, no escalamos: revisamos qué falló y ajustamos
V0 antes de sumar complejidad.

## Esperanza matemática V0 a 3 meses

Con la misma metodología que el escenario base del plan completo
(`05-escenarios/simulacion.md`), pero ponderado a 4 activos:

| Escenario | Retorno mensual ponderado | Capital fin 3M (sin aportes) |
|-----------|---------------------------|-------------------------------|
| Bull      | ~18%                      | 410.000                       |
| Base      | ~10%                      | 333.000                       |
| Bear      | ~−10%                     | 182.000                       |

**Con aportes de ARS 150.000 al cierre de mes 1 y 2** (total aporte
acumulado 550.000):

| Escenario | Capital fin 3M | % sobre aporte total |
|-----------|----------------|----------------------|
| Bull      | 791.000        | +43.8%               |
| Base      | 645.000        | +17.3%               |
| Bear      | 405.000        | −26.4%               |

V0 sacrifica ~5–8 puntos de upside en bull respecto a V1, a cambio de
simplicidad operativa. Es un trade-off consciente para empezar.

## Checklist primer día

- [ ] Verificar saldo disponible en IOL ≥ ARS 250.000.
- [ ] Confirmar que el plazo 48h (T+2) está habilitado para acciones,
      bonos y CEDEARs.
- [ ] Tener anotado: precio último de GGAL, GD30, NVDA al abrir rueda.
- [ ] Cargar las 3 órdenes limitadas (tramo 1) entre 11:00 y 12:00
      ARG, cuando el spread suele ser más ajustado.
- [ ] Suscribir IOL Ahorro Plus por ARS 37.500.
- [ ] Anotar todo en `04-seguimiento/registro-diario.md`.
