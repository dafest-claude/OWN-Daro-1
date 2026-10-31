# Cómo ejecutar stops y TP en IOL

**Pregunta original**: ¿cómo configuro stop con venta y TP de 1/3 en IOL?

**Respuesta corta**: IOL no tiene stop-loss automático garantizado en la
versión estándar. Hay 3 caminos según las capacidades de tu cuenta.

---

## Los 3 caminos

### Camino 1 — Alerta + venta manual ✅ (recomendado para esta cartera)

1. Configurás alerta de precio (lo que ya armamos en la tabla del 23/06).
2. Cuando suena → entrás a IOL.
3. Cargás la orden de venta limitada.
4. Confirmás con clave.

**Pro**: simple, control total.
**Contra**: necesitás reaccionar.

### Camino 2 — Stop Loss "real" (si tu cuenta lo tiene)

Verificar en la app:
1. *Operar → buscar ticker → Vender*.
2. En "Tipo de orden" mirar opciones disponibles.
3. Si aparece "**Stop Loss**" o "**Stop Limit**", usalo:
   - **Precio de disparo**: el del stop.
   - **Cantidad**: total de la posición.
   - **Vigencia**: GTC (hasta cancelar) si disponible, o día.

**Pro**: ejecuta sola.
**Contra**: depende de si tu cuenta lo soporta.

### Camino 3 — Orden limitada renovada diaria (workaround)

1. Cargás venta limitada al precio del stop.
2. Vigencia: día.
3. Si no se ejecuta, IOL la cancela al cierre.
4. La renovás cada mañana.

**Pro**: protección pasiva durante el día.
**Contra**: hay que cargarla todos los días.

---

## Cantidades exactas para vender 1/3 (TP parcial)

| Posición | Total | **Vender 1/3** | Resto |
|----------|-------|----------------|-------|
| BMA | 5 acciones | **2** (40% real) | 3 |
| GGAL | 8 acciones | **3** (37.5% real) | 5 |
| TGSU2 | 3 acciones | **1** (33% real) | 2 |
| GD30 | 93 nominales | **31** | 62 |
| AO27 | 28 nominales | **9** (32% real) | 19 |
| BPOC7 | 20 nominales | **7** (35% real) | 13 |
| NVDA | 8 CEDEARs | **3** (37.5% real) | 5 |
| COPX | **2 CEDEARs** | ⚠️ no se puede (mín. 1 = 50%) | — |

> El 1/3 redondea al entero más cercano. En posiciones de 2 unidades
> (COPX) el TP parcial **no funciona** — se vende 1 (50%) o 2 (100%).

---

## Pasos para vender cuando dispara una alerta

### Venta 100% (al disparar STOP)

1. *Operar → Acciones/CEDEARs/Bonos → buscar ticker → Vender*.
2. Plazo: **48 horas (T+2)**.
3. Tipo: **Limitada**.
4. Precio: último visible × **0.997** (cerca del libro para asegurar fill).
5. Cantidad: **toda la posición**.
6. Vigencia: día.
7. Confirmar con clave.
8. Producido en pesos en T+2 → suscribir **PRCPPEB**.

### Venta 1/3 (al disparar TP parcial)

1. Mismo flujo, pero cantidad = columna "Vender 1/3".
2. Después de ejecutar: **subir el stop al PMC original** (break-even
   sobre el resto de la posición).

---

## Cuál camino usar según perfil

| Mirás la app cada día? | Camino sugerido |
|------------------------|-----------------|
| Sí | **Camino 1** (alertas + venta manual) |
| Sí, doble protección | **Camino 1 + 3** combinados |
| No, viajás | **Camino 2** si está disponible, sino **3** |

Para esta cartera y semana (hasta el 30/06), **Camino 1 alcanza**.

---

## Notas operativas

- **No hay stop loss "garantizado a precio"** en IOL. En días de gap
  fuerte (apertura con −10%), la ejecución puede ser peor que el
  precio del stop.
- **Las comisiones** se aplican igual: 0.6% + IVA + mercado por venta.
  En posiciones chicas (COPX) puede haber mínimo de operación.
- **Plazo T+2**: el dinero está "atrapado" en boleto por 48hs antes de
  que puedas reasignarlo.
