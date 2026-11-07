# ARRANQUE — Miércoles 2026-05-27

Este es el único documento que necesitás abrir mañana. Reúne en orden la
calculadora de órdenes, el árbol de decisiones para imprevistos y el
link al registro del día.

Si querés profundizar en algo, los archivos de referencia están al final.

---

## 1. CHECKLIST — antes de dormir esta noche (martes 26/05)

- [ ] Saldo IOL en pesos ≥ ARS 250.000 confirmado.
- [ ] Si falta plata, transferencia hecha antes de las 22:00.
- [ ] Perfil de inversor activado para bonos y CEDEARs.
- [ ] App IOL abierta y logueada — clave operativa a mano.
- [ ] App con notificaciones push habilitadas (para alertas de precio).
- [ ] Despertador a las 09:00 ARG.

---

## 2. CALCULADORA DE ÓRDENES — completar entre 09:30 y 10:45

> Buscar cada ticker en IOL, anotar el último precio visible y completar
> la columna *Cantidad*. Redondear siempre hacia abajo en nominales.

### GGAL — Tramo 1: ARS 37.500

| Dato                   | Valor                  |
|------------------------|------------------------|
| Último precio (ARS)    | ____________ (anotar)  |
| Precio límite (+0.4%)  | último × 1.004 = _____ |
| Cantidad a comprar     | piso(37.500 / último) = _____ acciones |
| Monto real estimado    | cantidad × precio límite = _____ |

**Ejemplo numérico**: si último = 5.200 → límite 5.221 → cantidad 7 (37.500/5.200=7.21, piso 7) → monto 36.547. La diferencia con 37.500 queda como cash residual.

### GD30 — Tramo 1: ARS 37.500

> **OJO**: GD30 cotiza **por cada 100 VN**. La fórmula cambia.

| Dato                       | Valor                  |
|----------------------------|------------------------|
| Último precio (cada 100 VN) | ____________ (anotar) |
| Precio límite (+0.3%)      | último × 1.003 = _____ |
| Cantidad a comprar (en VN) | piso(37.500 / último × 100) = _____ VN |
| Monto real estimado        | (cantidad/100) × precio límite = _____ |

**Ejemplo numérico**: si último = 55.000 (por 100 VN) → cantidad = piso(37.500/55.000 × 100) = piso(68.18) = **68 VN** → monto 37.400. Confirmar que la operación se carga en múltiplos válidos (suele ser 1 VN en GD30).

### NVDA — Tramo 1: ARS 31.250

| Dato                   | Valor                  |
|------------------------|------------------------|
| Último precio (ARS)    | ____________ (anotar)  |
| Premarket NVDA en USA  | ______ %               |
| Precio límite (+0.5%)  | último × 1.005 = _____ |
| Cantidad a comprar     | piso(31.250 / último) = _____ CEDEARs |
| Monto real estimado    | cantidad × precio límite = _____ |

> Si premarket está peor a −3%, **saltear NVDA al mediodía**. Cargarlo después de 15:45 (apertura USA confirmada).

### FCI Money Market — Tramo único: ARS 37.500

No hay precio que anotar. **Elegir el fondo en pantalla aplicando estos 4 filtros**:

1. Categoría: **Money Market** o "Liquidez".
2. Moneda: **pesos**.
3. Rescate: **T+0** (mismo día).
4. Comisión de suscripción/rescate: **0%**.

Ruta: *Inversiones → Fondos → filtrar Money Market → ordenar por TIR 30d desc → elegir el primero que cumpla los 4 criterios y tenga mayor patrimonio*.

Candidatos típicos en IOL (verificar disponibilidad): Allaria Ahorro Pesos, Cocos Ahorro Pesos, Compass Renta Fija, Consultatio Ahorro Plus, Delta Pesos, Galileo Ahorro, Pellegrini Renta Pesos, Premier Renta Corto Plazo, Schroder Renta Plus, ST Ahorro Pesos.

Una vez elegido: *Suscribir → ARS 37.500 → Confirmar con clave*. Anotar en el registro qué fondo terminaste eligiendo.

### Suma de control

- Monto teórico Tramo 1: 37.500 + 37.500 + 31.250 + 37.500 = **143.750**
- Monto real estimado (después de redondeos): _____________
- Cash residual esperado: 143.750 − monto real = _____________ (irá al FCI más tarde)

---

## 3. ORDEN DE OPERACIÓN — miércoles 27/05

| Hora       | Acción                                                  |
|------------|---------------------------------------------------------|
| 09:30–10:45| Completar la **Calculadora** de arriba                  |
| 11:00–11:15| Observar apertura — no operar                           |
| 11:15      | **Orden 1**: FCI Money Market elegido, ARS 37.500       |
| 11:30      | **Orden 2**: GGAL — limitada                            |
| 12:15      | **Orden 3**: GD30 — limitada (verificar GD30, no GD30D) |
| 12:45      | **Orden 4**: NVDA — limitada (si premarket OK)          |
| 13:15–17:00| Vigilar ejecución cada 30 min                           |
| 15:30      | Apertura Wall Street — verificar NVDA                   |
| 17:00      | Cierre + configurar alertas + completar registro        |

---

## 4. ÁRBOL DE DECISIONES — qué hacer si pasa X

### A. GGAL abre con gap >+3% sobre el cierre de ayer
→ **Esperar retroceso.** No cargar a las 11:30. Mirar a las 12:30: si bajó al menos 1.5%, cargar limitada. Si sigue volando, **saltar el Tramo 1 de GGAL hoy** y cargarlo mañana junto con el Tramo 2A.

### B. GD30 cae >−2% en la apertura
→ **Oportunidad.** Cargar el Tramo 1 inmediatamente con limitada al **último visible (sin +0.3%)**, no pagar premium si el mercado está vendedor.

### C. NVDA en premarket USA está a −3% o peor
→ **Saltear orden al mediodía.** Esperar a las 15:45 (USA ya abierta y digerida la apertura). Si en USA NVDA recuperó algo, cargar. Si sigue cayendo, **cancelar la orden hoy** y reintentar mañana.

### D. Una orden no se ejecuta después de 1 hora
→ Subir el precio limitado un **+0.3% adicional**. Si en 2 horas tampoco, **pasar a mercado solo si el spread bid-ask es <1%**. Si el spread está más ancho, cancelar y reintentar mañana — no pagues spread caro por urgencia artificial.

### E. Una orden se ejecuta solo parcialmente
→ Dejar el parcial como está. **No reorder por el remanente** salvo que sea <50% del objetivo. Si fue >50% ejecutado, completás el resto en el Tramo 2A mañana.

### F. Merval cae más de −3% entre las 11:00 y las 13:00
→ Cargar GGAL y GD30 igual (entran a mejor precio). **No tocar NVDA hasta que abra USA.** Si al cierre Merval terminó >−5%, **no cargar Tramo 2A mañana** — esperar un día de estabilización.

### G. MEP sube más de +3% en el día
→ Señal de stress cambiario. Cargar GD30 sin descuento (entrás a mejor precio porque la suba del MEP infla la cotización en ARS). **No vender nada por ahora**, pero anotar el evento. Si al día siguiente el MEP sigue subiendo a tasa alta, activar kill-switch (ver `01-estrategia/cartera-starter-V0.md`).

### H. Stop disparado en el Día 1 (improbable, pero por las dudas)
→ Vender la posición. **No promediar a la baja en el primer día.** El stop existe para protegerte de la rueda; respetarlo el día 1 es la prueba de fuego.

### I. Plataforma IOL caída o lenta
→ Probar la web (https://iol.invertironline.com) si era la app, o viceversa. Si las dos están caídas, usar el teléfono al 0810 de IOL para confirmar estado. **No operar a través de canales informales** (mail, WhatsApp).

### J. No tenés ganas / sentís miedo / dudás
→ Cargar solo el FCI y GD30 hoy (los dos más conservadores del set). Dejar GGAL y NVDA para mañana o pasado. **La cartera completa puede armarse en 5 días**, no es obligatorio que sea hoy.

---

## 5. CIERRE DEL DÍA — checklist 17:00

- [ ] Las 3 órdenes (GGAL, GD30, NVDA) figuran como **ejecutadas** en *Operaciones del día*.
- [ ] FCI Money Market aparece en *Mis tenencias* con la suscripción del día (anotar nombre exacto del fondo elegido).
- [ ] PMC de cada posición anotado.
- [ ] Cantidad nominal de cada uno anotada.
- [ ] **Alertas configuradas** (6 en total):

| # | Ticker | Tipo  | Precio gatillo            |
|---|--------|-------|---------------------------|
| 1 | GGAL   | Stop  | PMC × 0.85                |
| 2 | GGAL   | TP    | PMC × 1.25                |
| 3 | GD30   | Stop  | PMC × 0.92                |
| 4 | GD30   | TP    | PMC × 1.15                |
| 5 | NVDA   | Stop  | PMC × 0.88                |
| 6 | NVDA   | TP    | PMC × 1.20                |

- [ ] Registro del día completado: `04-seguimiento/registro-2026-05-27.md`.

---

## 6. ARCHIVOS DE REFERENCIA (no es necesario abrirlos mañana salvo duda puntual)

| Tema                           | Archivo                                              |
|--------------------------------|------------------------------------------------------|
| Cartera y reglas               | `01-estrategia/cartera-starter-V0.md`                |
| Detalle de cada instrumento    | `02-instrumentos/fichas-tickers.md`                  |
| Operativa general en IOL       | `03-operativa-IOL/paso-a-paso.md`                    |
| Cronograma del Día 1 ampliado  | `03-operativa-IOL/instructivo-dia-1.md`              |
| Template de registro diario    | `04-seguimiento/registro-diario.md`                  |
| Escenarios bull/base/bear      | `05-escenarios/simulacion.md`                        |
| Plan V1 (para después)         | `01-estrategia/plan-agresivo-3M.md`                  |

---

## 7. UN ÚLTIMO RECORDATORIO

El **objetivo del Día 1 NO es ganar plata**. Es ejecutar el armado
limpio, sin slippage caro, con alertas configuradas, y dejar el registro
prolijo. La rentabilidad se mide al cierre del mes 1, no a las 17:00 de
mañana.

Si mañana al cierre el portfolio cae −2% por slippage de ejecución y
movimiento normal de rueda, eso **no es un error** ni un mal arranque —
es ruido esperable.

Suerte.
