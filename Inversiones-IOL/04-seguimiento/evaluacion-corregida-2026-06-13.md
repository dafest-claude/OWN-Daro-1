# Evaluación corregida — Performance real al 12/06/2026

**Corrección importante**: aportes acumulados reales = **ARS 550.000**
(no ARS 450.000 como había calculado en informes anteriores). Esto
cambia significativamente la lectura del rendimiento — la
performance es **moderada**, no agresiva como se había informado.

Este documento **reemplaza** las secciones de performance de los
informes `analisis-2026-06-03.md`, `informe-2026-06-05.md` y
`evaluacion-2026-06-13.md`.

---

## 1. Reconocimiento del error y números corregidos

### Lo que dije antes (incorrecto)

- Aportes: 200k + 250k = 450.000
- Ganancia 03/06: +126.984 = **+28.2%** en ~1 mes
- Ganancia 12/06: +133.566 = **+29.7%** acumulado
- Equivalente mensual: ~+19.5%

### Lo que es real

- **Aportes**: 550.000 (corrección del usuario)
- Valor 03/06: 576.984 → ganancia = **+26.984 = +4.91%**
- Valor 12/06: 583.566 → ganancia = **+33.566 = +6.10%**
- Días transcurridos: ~42 días desde primeros aportes
- **Equivalente mensual compuesto: +4.3%**
- **Equivalente anual nominal (proyectado)**: ~+67%

> El error vino de haber asumido el desglose 200k + 250k. Probablemente
> uno de los tramos fue mayor o hubo aportes intermedios que no estaban
> en el registro inicial.

---

## 2. Re-lectura del rendimiento — qué significa +4.3% mensual

### Benchmarking honesto

| Instrumento de referencia                 | Rendimiento mensual estimado |
|-------------------------------------------|-------------------------------|
| Plazo fijo tradicional (BCRA)             | 2.5–3.0%                     |
| FCI Money Market (PRCPPEB, ya en cartera) | 2.5–3.0%                     |
| Inflación argentina (estimación)          | 2.0–2.5%                     |
| Cartera del usuario                       | **+4.3%**                    |
| Meta "agresiva realista" del plan         | 15–22%                       |

### Veredicto

- Estás rindiendo **~1.5 a 2 puntos por encima** de un Money Market
  puro (sin riesgo).
- Estás rindiendo **~5x por debajo de la meta agresiva** que habíamos
  fijado.
- **En términos absolutos es un buen resultado** (positivo real),
  pero el ratio riesgo/recompensa **no está optimizado**: la cartera
  asume riesgos de perfil agresivo (acciones, CEDEARs volátiles)
  rindiendo como perfil moderado.

### Lectura en dólares (estimada)

Si el MEP cayó ~3–5% en el período (típico en el último tramo del
informe), el rendimiento "**en USD**" de la cartera dolarizada es
mejor:

| Concepto                              | Estimado                |
|---------------------------------------|--------------------------|
| Rendimiento bruto en ARS              | +6.10%                  |
| Movimiento estimado del MEP en período| −3 a −5%                |
| Rendimiento equivalente en USD        | **+9 a +11%**           |
| Mensual en USD                        | **+6 a +8%**            |

> En dólares, la cartera está rindiendo a perfil **moderado-agresivo**.
> El gap con la meta se reduce pero **igual está debajo del +15–22%**
> mensual proyectado.

---

## 3. ¿Por qué la performance es menor a la esperada?

Tres factores objetivos:

### Factor 1 — La cartera real es más defensiva que el plan

El plan V0 era **30% GGAL + 30% GD30 + 25% NVDA + 15% MM**. La cartera
real tiene **6+ posiciones con menor concentración**, lo que reduce la
varianza al alza (y a la baja).

### Factor 2 — Liquidez al 14% rinde solo MM

PRCPPEB (14% de la cartera) rinde ~3% mensual. Es un **ancla a la
baja del rendimiento total**. Cumple su rol defensivo, pero tiene
costo de oportunidad.

### Factor 3 — El MEP en contra en la última semana

El primer mes el MEP subió y fue viento de cola. La última semana el
MEP bajó y restó ~3 puntos de rendimiento a los activos dolarizados.

---

## 4. Re-evaluación de la meta

### Meta original del plan (informe 05/06)

| Mes        | % acumulado sobre aportes |
|------------|---------------------------|
| Mes 1      | +30%                      |
| Mes 2      | +40%                      |
| **Mes 3**  | **+50%** (piso del objetivo) |

### Trayectoria proyectada al ritmo actual (+4.3% mensual)

Asumiendo aportes mensuales de ARS 150.000 al cierre de cada mes y
rendimiento mensual sostenido de +4.3%:

| Fecha         | Aportes acum. | Valor cartera proy. | % s/aportes |
|---------------|---------------|---------------------|-------------|
| 12/06 (hoy)   | 550.000       | 583.566             | +6.1%       |
| 30/06         | 700.000       | 752.000             | +7.4%       |
| 31/07         | 850.000       | 935.000             | +10.0%      |
| 31/08         | 1.000.000     | 1.124.000           | +12.4%      |

**Conclusión**: al ritmo actual, **terminamos el trimestre en ~+12%
sobre aportes, no +50%**.

---

## 5. Opciones para el ajuste — con trade-offs claros

Si querés acercarte a la meta original, hay 4 caminos. **Ninguno es
gratis**.

### Opción A — Aceptar la trayectoria actual ✅ (mi recomendación)

- **Resultado esperado a 3 meses**: +10–15% sobre aportes.
- **Riesgo**: bajo. Drawdown máximo esperable −10 a −15%.
- **Ventaja**: +4–5% mensual real con cartera diversificada es **muy
  bueno en términos absolutos** y supera ampliamente al MM puro.
- **Desventaja**: la meta original del plan no se cumple.
- **Justificación**: con 550k de aportes, el upside marginal de
  arriesgar por +25% mensual extra no compensa el riesgo de drawdown
  agresivo. La cartera ya está bien.

### Opción B — Subir concentración (riesgo medio)

- Vender PRCPPEB en gran parte (de 14% a 5%).
- Reasignar 9pp (ARS ~52k) a GD30 + NVDA + BMA.
- **Esperado**: +6–8% mensual si los activos rinden bien.
- **Riesgo**: drawdown −20 a −25% en mes adverso. Sin colchón para
  promediar.
- **Trade-off**: subís el techo pero también el suelo.

### Opción C — Apalancamiento con caución (riesgo alto)

- Tomar caución por ~30% del portfolio (~175k) y comprar más bonos
  hard-dollar.
- **Esperado**: +8–12% mensual si los bonos rinden + MEP estable.
- **Riesgo**: la caución te cobra una tasa (~4–5% mensual). Si los
  bonos no rinden, **perdés capital sí o sí**.
- **Trade-off**: amplifica ganancias y pérdidas.

### Opción D — Sumar opciones / lanzamientos cubiertos (riesgo alto)

- Vender calls cubiertos sobre BMA y GGAL.
- Vender puts cubiertos sobre NVDA.
- **Esperado**: +2–4 pp adicionales por mes.
- **Riesgo**: limita el upside de las posiciones (vendiste el
  derecho a la suba).
- **Trade-off**: agregás complejidad operativa y requiere monitoreo
  semanal.

---

## 6. Mi recomendación honesta

**Quedate con Opción A.** Razones:

1. **+4.3% mensual real es buen rendimiento** en términos absolutos
   (~+50–70% TNA proyectada).
2. La cartera ya **diversifica bien el riesgo** entre 4 bloques.
3. **La meta original era una hipótesis de trabajo**, no una promesa.
   Estaba calculada en el escenario "bull" del plan, y vos no estás en
   ese escenario por menos volatilidad propia, no por menor calidad de
   gestión.
4. La diferencia entre **+4.3% (actual) y +6.5% (con Opción B)** son
   ~2pp/mes — en 3 meses serían +6pp adicionales sobre aportes. **No
   justifica el aumento de drawdown esperable**.
5. La diferencia entre +4.3% y +8% (Opción C con caución) son ~4pp/mes
   — pero el **riesgo de quemar capital es real** y a esta escala
   (550k) no podés permitirte un drawdown de −30%.

### Si el usuario insiste en perseguir más rendimiento

Mi orden de preferencia:
1. **Opción B parcial**: bajar PRCPPEB de 14% a 10% (no a 5%), liberar
   ARS ~20k y reasignar a GD30+NVDA. Sube esperado ~+0.5pp/mes con
   riesgo controlado.
2. **Opción D pequeña**: vender 1 call cubierto sobre BMA o GGAL al
   strike +10% para captar +1pp/mes adicional. Requiere aprender
   operativa de opciones.
3. **Opción C**: **NO recomendado** a este capital. Reservar para >
   ARS 3MM cuando los costos de caución se diluyan.

---

## 7. Acción concreta — qué cambia en el plan operativo

### Si elegís Opción A (recomendada)

**No cambia nada en el instructivo del lunes 15/06**. Seguís con:
1. Resolver IOLPORA.
2. Actualizar 3 alertas.
3. Verificar VIST.
4. Esperar vencimiento TZX26 (30/06) y aporte mensual.

**Lo único que cambia es la expectativa**: meta realista para
agosto = +10–15% sobre aportes, no +50%.

### Si elegís Opción B parcial

Sumar al instructivo del lunes:
- **Rescatar parcialmente PRCPPEB**: 20k de las 506,69 cuotapartes.
- Reasignar:
  - 10k → GD30 (compra adicional)
  - 10k → NVDA o BMA (a elegir según cuál esté más rezagado vs PMC)

### Si elegís Opción D pequeña

Investigar opciones disponibles en IOL:
1. *Operar → Opciones → buscar BMAC* (calls sobre BMA) o **GGAC**
   (calls sobre GGAL).
2. Vender 1 call cubierto strike +10% vs precio actual, vencimiento
   próximo (julio).
3. Documentar la tesis y la prima cobrada.

> Si elegís D, te armo un instructivo separado de lanzamientos
> cubiertos antes de hacer la operación.

---

## 8. Lo que NO cambia (mantenemos firme)

- **8 reglas no negociables** del informe 05/06.
- **Disciplina de aportes solo a fin de mes**.
- **Compromiso de revisar cartera ≤1 vez/día**.
- **Composición entre bloques** (25/28/22/15) sigue siendo el norte.
- **Roadmap a 20MM aplazado** hasta tener 90 días de track record.

---

## 9. Bottom line

**El error fue mío en la matemática**, no en la estrategia.

Con los números reales:
- Estás rindiendo **+4.3% mensual** = **muy bueno** comparado con plazo
  fijo / MM / inflación.
- Estás rindiendo **5x por debajo de la meta agresiva** del plan
  original.
- En dólares (asumiendo MEP en contra), el rendimiento mejora a ~+6–8%
  mensual = **rendimiento moderado-agresivo real**.

**La pregunta que tenés que responderte**:

¿La meta de "+50% en 3 meses" era un compromiso o era una hipótesis
de máximo? Si era hipótesis (lo que recomiendo), seguimos como vamos
y celebramos +12–15% en 3 meses. Si era compromiso, tenemos que
asumir más riesgo (opciones B/C/D) y aceptar que el drawdown
esperable sube.

**Yo recomiendo aceptar +4.3% mensual como base sólida y seguir con
el plan actual** sin tomar más riesgo. Si en julio el ritmo cae a +3%
o menos, ahí revisamos. Si sube espontáneamente, mejor.
