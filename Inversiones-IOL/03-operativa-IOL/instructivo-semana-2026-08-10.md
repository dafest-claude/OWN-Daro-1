# Instructivo — Semana del 10 al 14/08/2026

**2 tareas concretas**:
1. Configurar 8 alertas nuevas (30 min).
2. Comprar +2 BMA con el cash disponible (10 min).

**Tiempo total estimado**: 40 minutos.

---

## 1. COMPRA DE BMA (10 min)

### Datos operativos

| Concepto | Valor |
|----------|-------|
| Cash disponible | 29.680 |
| Precio BMA (06/08) | 13.650 |
| Precio límite (+0.4%) | **13.705** |
| Cantidad a comprar | **2 acciones** |
| Monto estimado | 27.410 |
| Cash residual esperado | ~2.270 |

### Pasos en IOL

1. *Operar → Acciones → buscar BMA → Comprar*.
2. Plazo: **48 horas (T+2)**.
3. Tipo: **Limitada**.
4. Precio: **13.705** (o último × 1.004 si el precio cambió).
5. Cantidad: **2 acciones**.
6. Vigencia: día.
7. Confirmar con clave.

### Después de ejecutar

- Posición BMA nueva: **9 acciones** (7 + 2).
- PMC mixto estimado: ~**13.545**.
- **Actualizar alerta BMA** con los gatillos nuevos:
  - Stop: 13.545 × 0.85 = **11.513**
  - TP: 13.545 × 1.25 = **16.931**

---

## 2. CONFIGURAR 8 ALERTAS (30 min)

### Nuevas alertas — MSFT, AAPL, QQQ, SPY (4 tickers × 2 alertas = 8)

| # | Ticker | Tipo | Condición | Precio |
|---|--------|------|-----------|--------|
| 1 | **MSFT** | Stop | Menor o igual a | **20.680** |
| 2 | **MSFT** | TP | Mayor o igual a | **28.200** |
| 3 | **AAPL** | Stop | Menor o igual a | **21.736** |
| 4 | **AAPL** | TP | Mayor o igual a | **29.640** |
| 5 | **QQQ** | Stop | Menor o igual a | **49.698** |
| 6 | **QQQ** | TP | Mayor o igual a | **67.770** |
| 7 | **SPY** | Stop | Menor o igual a | **17.776** |
| 8 | **SPY** | TP | Mayor o igual a | **24.240** |

### Pasos en IOL (repetir 8 veces)

1. *Operar → CEDEARs → buscar ticker → ficha del CEDEAR*.
2. 🔔 **Crear alerta**.
3. Tipo: **Precio**.
4. Condición: la de la tabla (*Menor o igual* para stop, *Mayor o igual* para TP).
5. Precio gatillo: el valor de la tabla.
6. Canal: **push (app) + email**.
7. Guardar.

### Los PMC son ESTIMADOS

Los gatillos están calculados sobre los **precios de compra estimados** (el ratio y PMC exactos podrían diferir). Si al ver el PMC real en *Mi Cartera* difiere en >5%, avisame y recalculo.

---

## 3. Cronograma sugerido

| Día | Acción |
|-----|--------|
| **Lunes 10/08** | Comprar +2 BMA (11:30 apertura BYMA calma) |
| **Martes 11/08** | Verificar ejecución BMA + configurar 4 alertas (MSFT + AAPL) |
| **Miércoles 12/08** | Configurar 4 alertas restantes (QQQ + SPY) |
| **Jueves 13/08** | Verificar todas las alertas activas en la app |
| **Viernes 14/08** | Revisión rápida de cierre de semana |

---

## 4. Después de esta semana — total alertas activas

Cuando termines el jueves 13/08, tendrás cubiertas todas las posiciones que necesitan alerta:

| Bloque | Alertas activas | Sin alerta |
|--------|-----------------|-------------|
| Acciones ARG | BMA + GGAL + TGSU2 (a actualizar cuando tengas PMC real) | — |
| Bonos hard-dollar | AO27, GD30, BPOC7 (a configurar cuando tengas PMC real) | — |
| CEDEARs USA | ✅ NVDA, MSFT, AAPL, QQQ, SPY | — |
| FCI MM | — | PRCPPEB (no requiere) |
| IOLPORA | ✅ Configurada 15/06 | — |

**Cobertura de gestión de riesgo**: ~70% del portfolio con alerta
formal. Suficiente para operar tranquilo hasta el 31/08.

---

## 5. Reglas de acción cuando dispara una alerta

### Si dispara STOP → vender 100%

1. Vender toda la posición (limitada al precio actual × 0.997).
2. Producido a PRCPPEB.
3. Anotar en registro y avisarme.

### Si dispara TP → vender 1/3 (con condiciones)

| Ticker | Total | Vender 1/3 (redondeado) |
|--------|-------|--------------------------|
| MSFT | 2 | **1** (50% real) |
| AAPL | 2 | **1** (50% real) |
| QQQ | 1 | ❌ no aplica (solo 1) |
| SPY | 1 | ❌ no aplica (solo 1) |
| NVDA | 11 | **4** (36% real) |
| BMA (post-compra) | 9 | **3** (33% real) |
| GGAL | 11 | **4** (36% real) |
| TGSU2 | 4 | **1** (25% real) |

**Para QQQ y SPY**: como solo tenés 1 CEDEAR, el TP parcial no funciona. Alternativas:
- Vender el 100% cuando llegue al TP.
- Mantener y esperar más upside (subir alerta a +30%).

Mi sugerencia: **cuando toque TP, vender el 100% de QQQ o SPY** y reasignar a otro instrumento.

---

## 6. Lo que NO hacés esta semana

- ❌ No comprar más CEDEARs (ya en 37.7%).
- ❌ No vender nada.
- ❌ No tocar bonos hard-dollar (esperan aporte 31/08).
- ❌ No aumentar posición en QQQ o SPY hasta que tengan track record.

---

## 7. Checklist rápido

- [ ] Lunes: comprar +2 BMA a 13.705
- [ ] Martes: verificar BMA ejecutado + alertas MSFT (2) + AAPL (2)
- [ ] Miércoles: alertas QQQ (2) + SPY (2)
- [ ] Jueves: actualizar alerta BMA con nuevo PMC (11.513 / 16.931)
- [ ] Viernes: revisión semanal + registro

Total: 40 minutos distribuidos en 5 días.
