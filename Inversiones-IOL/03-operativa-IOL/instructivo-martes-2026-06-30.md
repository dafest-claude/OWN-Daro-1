# Instructivo — Martes 2026-06-30 (cierre mes 1)

**Día crítico**: vencimiento TZX26 + aporte mensual + venta COPX +
compra QQQ + rebalanceo.

**Total operaciones**: 1 venta + 7 compras.
**Tiempo estimado**: 45 minutos.

---

## 1. Recursos que se van a mover

| Fuente | Monto ARS | Timing |
|--------|-----------|--------|
| Cobro TZX26 (vence hoy) | ~20.400 | Automático durante el día |
| Aporte del usuario | 150.000 | Al hacer transferencia desde banco |
| Venta COPX (3 CEDEARs) | ~24.970 neto | Al ejecutar la orden |
| Cash actual | 1.449 | ya disponible |
| **TOTAL disponible** | **~196.800** | |

---

## 2. Antes de las 11:00 — preparación

### Chequeo previo
- [ ] Verificar que el aporte de 150k está en camino desde tu banco a IOL.
- [ ] Anotar precios de cierre del lunes 29/06 (los tenés en el estado
  de cuenta).
- [ ] Confirmar en IOL que el TZX26 aparece marcado como "vencimiento
  hoy" o similar.

### Precios de referencia (del 29/06)

| Ticker | Precio 29/06 | Cantidad actual |
|--------|--------------|------------------|
| BMA | 14.400 | 5 |
| GGAL | 7.875 | 8 |
| TGSU2 | 9.275 | 3 |
| NVDA | 12.590 | 8 |
| **COPX** | **8.390** | **3 (a VENDER)** |
| PRCPPEB (cuotaparte) | 160.831 | 506.69 |

### Cálculo aproximado de cantidades a comprar

Con precios del 29/06 (los reales pueden variar mañana):

| Compra | Monto | Precio est. | Cantidad estimada |
|--------|-------|-------------|--------------------|
| BMA | 25.000 | 14.400 | **1 acción** (25k/14.4k = 1.7 → 1) |
| GGAL | 20.000 | 7.875 | **2 acciones** (20k/7.9k = 2.5 → 2) |
| TGSU2 | 15.000 | 9.275 | **1 acción** (15k/9.3k = 1.6 → 1) |
| NVDA | 30.000 | 12.590 | **2 CEDEARs** (30k/12.6k = 2.4 → 2) |
| QQQ | 60.000 | ~35.000–40.000 (mirar pantalla) | **1 o 2 CEDEARs** |

> Redondear siempre hacia abajo en cantidades.

---

## 3. Cronograma del día

| Hora | Acción |
|------|--------|
| 09:30 – 10:45 | Chequeos + verificar aporte llegó a IOL + anotar precios de apertura |
| 11:00 – 11:15 | Apertura BYMA — no operar, observar |
| **11:15** | **Venta 1**: COPX (3 CEDEARs) |
| 11:30 | Verificar acreditación TZX26 (aparece en Movimientos) |
| **11:45** | **Compra 2**: PRCPPEB (suscribir 30k) — T+0 |
| **12:00** | **Compra 3**: BMA (1 acción) |
| **12:15** | **Compra 4**: GGAL (2 acciones) |
| **12:30** | **Compra 5**: TGSU2 (1 acción) |
| **12:45** | **Compra 6**: NVDA (2 CEDEARs) |
| **13:00** | **Compra 7**: QQQ (1–2 CEDEARs) |
| 13:15–17:00 | Vigilar ejecuciones |
| 15:30 | Apertura Wall Street — confirmar QQQ, NVDA |
| 17:00 | Cierre + crear alerta QQQ + registro del día |

---

## 4. Paso a paso operación por operación

### 🔴 Venta 1 (11:15) — COPX 100%

1. IOL → *Operar → CEDEARs → buscar COPX*.
2. *Vender*.
3. Plazo: **48 horas (T+2)**.
4. Tipo: **Limitada**.
5. Precio: último visible **× 0.997** (bajar 0.3% para asegurar fill).
6. Cantidad: **3** (toda la posición).
7. Vigencia: día.
8. Confirmar.

Producido esperado: **~ARS 24.970 netos** (a acreditar T+2 = jueves 02/07).

> ⚠️ **Atención**: el producido de la venta acredita **T+2** (jueves).
> Las compras que hacés hoy deben usar el aporte y el TZX26. La venta
> COPX se **materializa el jueves** — es cash prometido, no inmediato.

### 🟢 Compra 2 (11:45) — PRCPPEB 30k

1. IOL → *Inversiones → Fondos → PRCPPEB → Suscribir*.
2. Monto: **ARS 30.000**.
3. Confirmar con clave.

Acreditación cuotaparte: al cierre del día.

### 🟢 Compra 3 (12:00) — BMA 25k

1. IOL → *Operar → Acciones → BMA → Comprar*.
2. Plazo: **48 horas**.
3. Tipo: **Limitada**.
4. Precio: último × **1.004**.
5. Cantidad: **1 acción** (verificar en pantalla).
6. Vigencia: día.
7. Confirmar.

### 🟢 Compra 4 (12:15) — GGAL 20k

1. IOL → *Operar → Acciones → GGAL → Comprar*.
2. Plazo: 48h. Tipo Limitada.
3. Precio: último × **1.004**.
4. Cantidad: **2 acciones**.
5. Confirmar.

### 🟢 Compra 5 (12:30) — TGSU2 15k

1. IOL → *Operar → Acciones → TGSU2 → Comprar*.
2. Plazo: 48h. Limitada.
3. Precio: último × **1.004**.
4. Cantidad: **1 acción**.
5. Confirmar.

### 🟢 Compra 6 (12:45) — NVDA 30k

1. IOL → *Operar → CEDEARs → NVDA → Comprar*.
2. Plazo: 48h. Limitada.
3. Precio: último × **1.005** (CEDEARs tienen spread más ancho).
4. Cantidad: **2 CEDEARs**.
5. Confirmar.

### ⭐ Compra 7 (13:00) — QQQ 60k [POSICIÓN NUEVA]

1. IOL → *Operar → CEDEARs → **buscar QQQ***.
2. **Verificar** que aparece "Invesco QQQ Trust" y ratio 16:1 (o el
   que corresponda ese día).
3. *Comprar*.
4. Plazo: **48 horas**.
5. Tipo: **Limitada**.
6. Precio: último × **1.005**.
7. Cantidad: **redondear hacia abajo** el resultado de 60.000 /
   precio unitario. Ej. si precio = 38.000 → 60k/38k = 1.58 →
   **1 CEDEAR**. Si precio = 30.000 → 60k/30k = 2 → **2 CEDEARs**.
8. Vigencia: día.
9. Confirmar.

> Con 60k probablemente compres **1 o 2 CEDEARs**. Si te queda saldo
> residual >15k y podés comprar 1 CEDEAR más, dale. Si es <15k,
> dejalo en cash para el próximo mes.

---

## 5. Al cierre 17:00 — 3 tareas

### A. Verificar todas las ejecuciones

- [ ] Venta COPX: ejecutada / pendiente / cancelada
- [ ] PRCPPEB: suscripción confirmada (aparece en Mi Cartera)
- [ ] BMA: cantidad ejecutada ______
- [ ] GGAL: cantidad ejecutada ______
- [ ] TGSU2: cantidad ejecutada ______
- [ ] NVDA: cantidad ejecutada ______
- [ ] QQQ: cantidad ejecutada ______
- [ ] Cash residual final: ARS ______

### B. Crear alerta para QQQ ⭐

Regla nueva para QQQ (CEDEAR):
- Stop: PMC × **0.88** (−12%)
- TP: PMC × **1.20** (+20%)

**Cálculo con el precio de compra final** (ej. PMC 38.500):
- Stop: 38.500 × 0.88 = **33.880**
- TP: 38.500 × 1.20 = **46.200**

Pasos en IOL para configurar:
1. Ficha del CEDEAR QQQ → 🔔 Crear alerta.
2. Alerta 1 (Stop): condición *Menor o igual a* + valor stop.
3. Alerta 2 (TP): condición *Mayor o igual a* + valor TP.
4. Canal: push + email.

### C. Registro del día

Crear archivo `04-seguimiento/registro-2026-06-30.md` con:

```markdown
# Registro 30/06 — Cierre mes 1

## Recursos
- Aporte: 150.000
- TZX26 cobrado: __________
- Venta COPX (bruto): __________
- Comisión venta: __________
- Producido neto COPX: __________

## Operaciones
| Ticker | Cantidad | PMC final | Total ARS |
|--------|----------|-----------|-----------|
| PRCPPEB (nuevas cuotapartes) |  |  |  |
| BMA (adicional) |  |  |  |
| GGAL (adicional) |  |  |  |
| TGSU2 (adicional) |  |  |  |
| NVDA (adicional) |  |  |  |
| QQQ (nueva posición) |  |  |  |

## Cartera al cierre mes 1
- Valor total: ARS __________
- Aportes acumulados: 700.000
- % sobre aportes: ______ %
- Cumplimiento meta ajustada (+5.5-6.5%): ______

## Alertas nuevas
- QQQ Stop @ __________ ✅/❌
- QQQ TP @ __________ ✅/❌

## Notas
```

---

## 6. Qué esperar mañana 01/07 y jueves 02/07

### Miércoles 01/07
- Nada especial. Verificar que todas las operaciones se acreditaron
  correctamente.

### Jueves 02/07
- **Acredita el efectivo de la venta COPX** (~24.970).
- Ese cash queda en pesos disponibles.
- **Decisión**: llevarlo a PRCPPEB (subir liquidez a ~17-18%) o
  dejarlo en cash para "munición" en el mes.
- Recomendación: **50/50** — 12k a PRCPPEB, 12k queda en cash.

### Viernes 03/07
- Cálculo del retorno cerrado del mes 1.
- Comparar con la meta ajustada (+5.5–6.5%).
- Documentar en `04-seguimiento/cierre-mes-1-2026-06.md`.

---

## 7. Reglas y prohibiciones para hoy

### NO HACER
- ❌ Cambiar el orden de compras.
- ❌ Comprar **bonos USD** (GD30, AO27, BPOC7). Están en máximos.
- ❌ Comprar más **COPX** (lo estamos vendiendo).
- ❌ Comprar **IOLPORA** (ya está en su rango objetivo).
- ❌ Aumentar aporte por encima de 150k (no cambiar el plan
  ejecutando más de lo definido).

### SÍ HACER
- ✅ Ejecutar en orden.
- ✅ Verificar cantidades antes de confirmar.
- ✅ Anotar todo al cierre.
- ✅ Crear alerta QQQ nueva.

---

## 8. Bottom line del día

**Objetivo del 30/06**: dejar la cartera limpia, con QQQ como nueva
diversificación tech, sin COPX (ruido eliminado), promediando las
posiciones argentinas y NVDA en descuento, y con liquidez sana.

**No es un día para hacer más de lo planificado**. Ejecutar los 7
movimientos limpios, anotar todo, crear alerta QQQ, y descansar. El
mes 2 empieza el miércoles.
