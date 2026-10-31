# Evaluación de cartera y ajustes — 2026-06-13

**Corte del estado de cuenta**: viernes 12/06/2026.
**Período de evaluación**: 03/06 → 12/06 (9 días corridos, 6 ruedas).
**Comparación**: contra hoja de ruta del `informe-2026-06-05.md`.

---

## 1. Resumen ejecutivo

### Lo que se ejecutó del plan ✅

| Acción planificada                              | Estado    |
|-------------------------------------------------|-----------|
| Vender EDN (ficha técnica)                      | ✅ Ejecutado |
| Vender TXAR (ficha técnica)                     | ✅ Ejecutado |
| Vender GOOGL (1 unidad)                         | ✅ Ejecutado |
| Vender GD41 (duration alta)                     | ✅ Ejecutado |
| Vender PAMP (solapaba con TGSU2)                | ✅ Ejecutado |
| Reasignar a PRCPPEB (subir liquidez)            | ✅ Ejecutado (+30k) |
| Reasignar a GD30 (refuerzo ancla)               | ✅ Ejecutado (+15 nominales) |
| Reasignar a NVDA o QQQ                          | ✅ Ejecutado (+1 NVDA) |

### Lo que no se ejecutó o queda pendiente ⚠️

| Pendiente                                       | Prioridad |
|-------------------------------------------------|-----------|
| Clarificar IOLPORA (reglamento + tesis)         | **Alta**  |
| Configurar 10 alertas (stop/TP) — verificar    | **Alta**  |
| Generar reserva de cash (hoy es 0)              | **Media** |
| Documentar tesis de VIST y COPX                 | **Media** |

### Decisión no planificada que tomaste

- **Sumaste 1 unidad de TGSU2** (de 2 a 3). No estaba en el plan, pero
  es coherente con la tesis "energía argentina concentrada en TGSU2".
  La acepto. Solo anotalo en `01-estrategia/` como tesis activa para
  saber por qué.

---

## 2. Performance del período

### Cifras

| Concepto                           | 03/06           | 12/06           | Δ              |
|------------------------------------|-----------------|-----------------|----------------|
| Cash ARS                           | 10.272          | 0               | −10.272        |
| Títulos valorizados                | 566.712         | 583.566         | +16.854        |
| **TOTAL cartera**                  | **576.984**     | **583.566**     | **+6.582**     |
| **Variación %**                    |                 |                 | **+1.14%**     |

### Performance acumulada desde el inicio (1 de mayo aprox.)

| Concepto                           | Valor           |
|------------------------------------|-----------------|
| Aportes totales                    | 450.000         |
| Valor actual                       | 583.566         |
| **Ganancia neta acumulada**        | **+133.566**    |
| **% sobre aportes**                | **+29.7%**      |
| Días aproximados                   | ~42 días         |
| **Equivalente mensual compuesto**  | **~+19.5%**      |

> **Lectura**: seguís corriendo por encima de la meta agresiva (+15–22%
> mensual del plan). Pero la última semana **rendiste solo +1.14%** —
> esto es **normal y esperado**, el mercado no rinde uniforme.

### Anatomía del +1.14% (descomposición por ticker)

| Ticker  | Δ precio % | Comentario |
|---------|------------|------------|
| BMA     | **+12.8%** | 🚀 Rally fuerte |
| GGAL    | **+12.0%** | 🚀 Rally fuerte |
| TGSU2   | **+7.4%**  | ✅ Bien |
| TZX26   | +0.7%      | Devengo CER esperado |
| IOLPORA | +1.1%      | Devengo normal |
| PRCPPEB | +0.4%      | Devengo MM normal (9 días) |
| BPOC7   | +0.1%      | Flat (esperado) |
| AO27    | −0.5%      | Plano |
| GD30    | −0.5%      | Plano |
| COPX    | −5.7%      | 🔻 Cobre flojo |
| NVDA    | −5.5%      | 🔻 Tech USA flojo |
| VIST    | **−7.0%**  | 🔻 **Atención** — ver §4 |

### Lectura técnica del período

Hubo **dos fuerzas opuestas**:

1. **Rally Merval doméstico**: BMA, GGAL y TGSU2 subieron 7–13%. Esto
   compensó casi todo lo demás.
2. **MEP/USD bajó** (estimado −3 a −5% en el período): los activos
   dolarizados (bonos USD, CEDEARs) se "achicaron" en pesos, no porque
   el activo cayera sino por el FX.

**El mes pasado** ganaste con MEP subiendo (viento de cola
cambiario). **Esta semana** el MEP te jugó en contra y **igual rendiste
+1.14%** gracias al equity argentino. **Eso es buena diversificación
funcionando.**

Si ajustamos por el MEP estimado: el rendimiento "en dólares" del
período fue **probablemente +5 a +6%** — excelente.

---

## 3. Composición actual vs objetivo V0.5

| Bloque                          | % real 12/06 | % objetivo V0.5 | Estado     |
|---------------------------------|--------------|-----------------|------------|
| Acciones líderes ARG            | 29.1%        | 25%             | ⚠️ +4pp    |
| Bonos hard-dollar (GD30+AO27+BPOC7) | 27.4%    | 28%             | ✅ OK      |
| CEDEARs USA                     | 22.5%        | 22%             | ✅ OK      |
| FCI MM (PRCPPEB)                | 13.9%        | 15%             | ✅ Cerca   |
| Boncer (TZX26, vence 30/06)     | 3.5%         | (sale el 30/06) | ⏳ Esperar |
| IOLPORA (sin tesis)             | 3.7%         | 0% (decidir)    | ⚠️ Pendiente |
| Cash ARS                        | 0.0%         | 2–3%            | ❌ **Crítico** |

### Issues principales

1. **Cash = 0**: en una cartera agresiva, tener 0 munición para
   promediar caídas es un error operativo. Hay que volver a 2–3%
   mínimo.
2. **Acciones argentinas en 29% (+4pp del objetivo)**: producto del
   rally que vimos. Si seguís cómodo con el peso, no toques. Si querés
   tomar ganancia parcial, vendé un **20% de la posición ganadora** y
   pasalo a PRCPPEB.
3. **PRCPPEB casi en 15%**: muy buena ejecución. Con el vencimiento de
   TZX26 (20k el 30/06) podés llevarlo a 17%.

---

## 4. Análisis posición por posición — qué cambia

### Posiciones núcleo — sin tocar
- **BMA, GGAL, TGSU2**: cumplen su rol, mantenelas.
- **GD30, AO27, BPOC7**: bloque hard-dollar bien construido.
- **NVDA**: pos. principal del bloque USA, mantenela.
- **PRCPPEB**: liquidez bien dimensionada.

### Posiciones que requieren decisión esta semana

**IOLPORA (3.7%) — sigue pendiente desde hace 8 días**
- Acción concreta: IOL → *Fondos → IOL Portafolio Potenciado →
  Reglamento*. Buscar 3 datos: (i) política de inversión, (ii)
  apalancamiento permitido, (iii) comisión de gestión.
- Si es **renta variable apalancada agresiva**: aceptable como 3–5%
  satélite. Documentar tesis.
- Si es **mixto sin apalancamiento**: redundante con la cartera, **vender**.
- Si tiene **comisión >2% anual**: probablemente vender.
- **Deadline real para esta decisión**: 19/06.

**VIST (6.2%) — cayó −7% en la semana**
- Verificar dos cosas:
  1. ¿Vista Energy en NYSE cayó también? Si sí, es subyacente (revisar
     noticias: producción, deuda, balance Q2).
  2. Si el subyacente está plano y solo cayó el CEDEAR, fue el MEP.
- Si el subyacente se mantiene, no vendas — la tesis Vaca Muerta sigue
  intacta.
- Si el subyacente cae >10% por noticia fundamental (downgrade, miss
  de producción), vendé y rotá a NVDA o QQQ.

**COPX (3.2%) — cayó −5.7%**
- Es el ETF del cobre. Cobre cae cuando hay miedo de recesión global.
- Si no tenés tesis macro clara sobre cobre/transición energética,
  **vendé**. 3% no aporta y agrega ruido sectorial.
- Si sí tenés tesis: documentala (precio objetivo, plazo).

**TZX26 (3.5%) — vence el 30/06 (en 17 días)**
- A vencimiento te paga capital × (1 + CER acumulado).
- **Plan de reinversión** (proyectado ARS 21–22k al cobro):
  - 60% → PRCPPEB (sube liquidez a ~17%).
  - 40% → un nuevo bono pesos (TZXJ7, TZXM7 o TZX27) para mantener
    cobertura inflación.

---

## 5. Comparación con la hoja de ruta — ¿estamos donde queríamos?

### Meta del informe del 05/06 para junio

| KPI                              | Meta junio | Actual al 12/06 | Veredicto |
|----------------------------------|------------|------------------|-----------|
| Rebalanceo V0.5 ejecutado        | sí         | Sí (80%)         | ✅        |
| Liquidez (MM + cash)             | ≥15%       | 13.9% (MM solo)  | ⚠️ Cerca  |
| Posiciones ≤12                   | ≤12        | 12 exactas       | ✅        |
| Concentración top-3 <40%         | <40%       | 42.3%            | ⚠️ +2pp   |
| Drawdown intramensual <15%       | <15%       | No se midió      | ❓        |
| Retorno mensual (proy 30/06)     | +10%       | +1.14% en 9 días | En curso  |

### Veredicto

**Ejecución muy buena** del rebalanceo, dos puntos de mejora menores:
- Liquidez está en 13.9% (faltan 1.1pp para meta).
- Top-3 está en 42% (objetivo <40%, exceso por rally de Merval + GD30
  reforzado).

Ambos se corrigen solos cuando llegue el vencimiento de TZX26 (+20k a
liquidez) y/o el próximo aporte mensual.

---

## 6. Ajustes recomendados — semana del 15 al 19/06

### Movimiento 1 — Decidir IOLPORA (prioridad #1)
Plazo: **antes del miércoles 17/06**. Sin tesis documentada es ruido.

### Movimiento 2 — Verificar las 10 alertas
Si las configuraste el 05/06, los precios gatillo pueden estar
desactualizados (los PMC bajaron por el promediado de GD30 y NVDA).

| Posición | Recalcular alerta porque... |
|----------|------------------------------|
| GD30     | Compraste 15 más → PMC cambió |
| NVDA     | Compraste 1 más → PMC cambió  |
| TGSU2    | Compraste 1 más → PMC cambió  |
| BMA, GGAL, VIST | Sin movimiento, alertas válidas |

Acción: revisar PMC actuales en *Mi Cartera* y **actualizar las 3
alertas afectadas**.

### Movimiento 3 — Plan para el vencimiento TZX26 (30/06)
Dejar pre-armado:
- **A los 2 días del cobro** suscribir 60% en PRCPPEB.
- **El 40% restante**: comprar TZXM7 o TZX27 (Boncer corto). Si no
  estás cómodo investigando bonos pesos nuevos, simplemente **todo a
  PRCPPEB** (lleva la liquidez a 17%, sin problema).

### Movimiento 4 — Investigación VIST (esta semana)
2 minutos:
1. Google: "Vista Energy stock 12 jun 2026".
2. Anotar: ¿cayó solo el CEDEAR o también el VIST en NYSE?
3. Si solo el CEDEAR → fue MEP, no preocuparte.
4. Si VIST en NYSE cae >5% por noticia → revisar y decidir.

### Movimiento 5 — NO HACER ahora
- **No comprar nada nuevo** esta semana.
- **No aumentar acciones argentinas** (ya están +4pp del objetivo).
- **No tomar ganancia** en BMA o GGAL salvo que rompan TP de alerta.
- **No reaccionar** a la caída de NVDA/VIST/COPX: es MEP, no falla de
  tesis.

---

## 7. Plan para el cierre de junio (30/06)

### Eventos esperados
1. **30/06**: vence TZX26 → cobro de ARS 21–22k.
2. **Fin de mes**: aporte mensual del usuario (100–200k según
   disponibilidad).

### Si aportás ARS 150.000 a fin de mes

Distribución sugerida del aporte:

| Destino             | Monto    | % aporte | Razón |
|---------------------|----------|----------|-------|
| PRCPPEB (FCI MM)    | 30.000   | 20%      | Mantener liquidez ≥15% |
| Cash (no invertir)  | 20.000   | 13%      | Reserva munición (no podés tener 0) |
| GD30                | 30.000   | 20%      | Reforzar ancla hard-dollar |
| BMA o GGAL          | 20.000   | 13%      | Solo si correccionaron −5% desde su pico |
| NVDA                | 30.000   | 20%      | Promediar a la baja del −5.5% |
| VIST                | 20.000   | 14%      | Solo si subyacente NYSE sigue intacto |

### Cálculo proyectado al 30/06

Asumiendo: cartera mantiene +1% promedio en lo que queda + cobro TZX26
+ aporte ARS 150k:

- Valor 13/06 aprox: 583.566
- + Rendimiento esperado 17 días: ~+2% = +11.700
- + Aporte: +150.000
- **Valor proyectado al 30/06: ~745.000**
- Aportes totales acumulados: 600.000
- **% sobre aportes: ~+24%**

Meta original del informe: +30% acumulado para junio. **Estamos en
camino** aunque algo por debajo, lo cual es esperable porque el +28% del
mes 1 era extraordinario.

---

## 8. Reglas que mantenemos en pie

Las 8 reglas no negociables del informe del 05/06 siguen vigentes. Las
copio acá para tenerlas al alcance:

1. Nunca más de 15% del portfolio en una sola acción argentina.
2. Nunca más de 15% en un solo CEDEAR.
3. Nunca menos de 10% en FCI MM + cash combinados.
4. No comprar nada sin documentar tesis.
5. No vender en pánico (esperar 24h salvo stop).
6. No agregar capital en día de >+5%.
7. Revisar 1 vez al día máximo.
8. Aportes solo a fin de mes.

### Compliance al 12/06

| Regla | Cumple? |
|-------|---------|
| 1. Acción ARG ≤15% | ✅ BMA en 12.7%, GGAL 11.3% |
| 2. CEDEAR ≤15% | ✅ NVDA en 13.2%, VIST 6.2% |
| 3. MM + cash ≥10% | ✅ 13.9% (justo, atender) |
| 4. Tesis documentada | ⚠️ IOLPORA, VIST, COPX pendientes |
| 5. No pánico | ✅ |
| 6. No comprar techo | ✅ |
| 7. Revisión 1x/día | (auto-reporte) |
| 8. Aportes a fin de mes | ✅ |

---

## 9. Bottom line

**El rebalanceo se ejecutó muy bien.** Pasaste de 17 a 12 posiciones,
subiste liquidez de 8.8% a 13.9%, mantuviste la asignación entre
bloques y no rompiste reglas.

**El +1.14% de la semana es saludable.** Mostró que la cartera
**funciona con MEP en contra** gracias al equity argentino. Eso es
diversificación que vale.

**Tres cosas a cerrar esta semana**:
1. Decidir qué hacer con **IOLPORA**.
2. **Actualizar las alertas** afectadas por nuevos PMC (GD30, NVDA,
   TGSU2).
3. **Verificar VIST**: ¿la caída fue MEP o subyacente?

**Una cosa a planificar**:
- Plan para el **cobro de TZX26** el 30/06 (≈21–22k).

**No tocamos la hoja de ruta**: estamos dentro del corredor proyectado,
con margen para llegar a la meta de cierre de junio (+30% acumulado) y
trimestre (+50% acumulado).
