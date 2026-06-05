# Instructivo — Ventas del viernes 2026-06-06

**Objetivo del día**: vender 3 posiciones chicas (EDN, TXAR, GOOGL) y
configurar las 10 alertas sobre las posiciones principales.

**Posiciones a tocar**:

| Ticker | Cantidad | Cotización 03/06 | Valor estimado | Tipo de orden |
|--------|----------|------------------|----------------|---------------|
| EDN    | 4        | ARS 2.035        | ARS 8.140      | A mercado o limitada −0.5% |
| TXAR   | 10       | ARS 676.5        | ARS 6.765      | A mercado o limitada −0.5% |
| GOOGL  | 1        | ARS 9.375        | ARS 9.375      | Limitada −0.3% |

**Total a liberar estimado**: ARS 24.280 (acredita T+2 = martes 10/06).

---

## 1. CHEQUEO CRÍTICO antes de las 11:00 — comisión mínima de IOL

Las 3 ventas son chicas. Si IOL aplica una **comisión mínima por
operación**, podría comerse un % alto del producido.

### Cómo verificar (5 minutos):

1. IOL → *Mi Cuenta → Costos / Comisiones* (o sección "Aranceles").
2. Buscar el dato: **"Comisión mínima por operación"** en Renta
   Variable y en CEDEARs.
3. Anotar acá: ARS __________ por orden.

### Regla de decisión:

| Comisión mínima | Decisión                                                |
|-----------------|----------------------------------------------------------|
| ≤ ARS 200       | **Vender las 3 sin problema**. Comisión total <2.5%.    |
| ARS 200–500     | **Vender, pero verificar que la comisión combinada <3%**. Aceptable como "limpieza". |
| ARS 500–1.000   | **Vender solo GOOGL** (la más grande). EDN+TXAR la comisión se come ~10–15%. |
| > ARS 1.000     | **No vender ninguna esta semana**. Esperar a poder consolidarlas o vender en bloque mayor. |

> Si caés en el caso ≥500, avisame antes de ejecutar y replanteamos.

### Cálculo orientativo (con 0.6% + IVA = 0.726%):

| Ticker | Producido | Comisión teórica 0.726% | ¿Aplica mínimo? |
|--------|-----------|--------------------------|------------------|
| EDN    | 8.140     | 59                       | Si mínimo > 59, sí |
| TXAR   | 6.765     | 49                       | Si mínimo > 49, sí |
| GOOGL  | 9.375     | 68                       | Si mínimo > 68, sí |

En estas operaciones chicas casi seguro aplica mínimo. Por eso el
chequeo importa.

---

## 2. ALERTAS — antes de cualquier venta

> Si solo vas a hacer una cosa hoy, que sea esto.

### Pasos en IOL para crear alerta:
1. Ficha del ticker (Operar → Acciones / Bonos / CEDEARs → buscar).
2. Ícono de campana 🔔 → *Crear alerta*.
3. Tipo: **Precio**.
4. Condición: **Mayor o igual / Menor o igual** según corresponda.
5. Precio gatillo: el del cuadro de abajo.
6. Canal: **Push (app)** + email.
7. Guardar.

### Las 10 alertas a configurar HOY:

Necesitás el **PMC (precio medio de compra)** de cada posición. Lo
encontrás en IOL → *Mi Cartera → Detalle de tenencia*.

| # | Ticker | Tipo  | Precio gatillo            | PMC × |
|---|--------|-------|---------------------------|-------|
| 1 | BMA    | STOP  | _____ (PMC × 0.85)        | 0.85  |
| 2 | BMA    | TP    | _____ (PMC × 1.25)        | 1.25  |
| 3 | GGAL   | STOP  | _____ (PMC × 0.85)        | 0.85  |
| 4 | GGAL   | TP    | _____ (PMC × 1.25)        | 1.25  |
| 5 | GD30   | STOP  | _____ (PMC × 0.92)        | 0.92  |
| 6 | GD30   | TP    | _____ (PMC × 1.15)        | 1.15  |
| 7 | NVDA   | STOP  | _____ (PMC × 0.88)        | 0.88  |
| 8 | NVDA   | TP    | _____ (PMC × 1.20)        | 1.20  |
| 9 | VIST   | STOP  | _____ (PMC × 0.85)        | 0.85  |
| 10| VIST   | TP    | _____ (PMC × 1.25)        | 1.25  |

**Anotar los PMC en `04-seguimiento/informe-2026-06-05.md`** (sección
plantilla mensual) para tenerlos a la vista.

---

## 3. CRONOGRAMA — viernes 2026-06-06

| Hora           | Acción                                                       |
|----------------|--------------------------------------------------------------|
| 09:30 – 10:45  | Chequeo de comisión mínima (§1) y PMC de las 5 principales (§2) |
| 10:45 – 11:00  | **Configurar las 10 alertas** (§2)                           |
| 11:00 – 11:15  | Apertura BYMA — **no operar**, observar 15 min               |
| 11:15 – 11:30  | Verificar libro de ofertas de EDN, TXAR, GOOGL               |
| 11:30          | **Venta 1**: GOOGL (la más simple, ratio claro, USA premarket OK) |
| 11:45          | **Venta 2**: EDN                                             |
| 12:00          | **Venta 3**: TXAR                                            |
| 13:00          | Verificar ejecuciones. Si quedaron pendientes, ajustar precio |
| 15:30          | Apertura USA — verificar VIST y NVDA (no operar, solo observar) |
| 17:00          | Cierre — registrar producido, validar que alertas siguen activas |

---

## 4. PASO A PASO DE CADA VENTA

### Venta 1 — GOOGL (11:30)

1. IOL → *Operar → CEDEARs → buscar GOOGL*.
2. *Vender* (no comprar).
3. **Plazo**: 48 horas (T+2).
4. **Tipo de orden**: Limitada.
5. **Cantidad**: **1** (vender la unidad completa).
6. **Precio limitado**: último visible **× 0.997** (3 décimas por
   debajo para asegurar fill sin regalar). Ej.: si último = 9.380 →
   limitada 9.351.
7. **Vigencia**: día.
8. Confirmar con clave operativa.

### Venta 2 — EDN (11:45)

1. IOL → *Operar → Acciones → buscar EDN*.
2. *Vender*.
3. **Plazo**: 48 horas.
4. **Tipo**: Limitada.
5. **Cantidad**: **4**.
6. **Precio limitado**: último × 0.995.
7. **Vigencia**: día.
8. Confirmar.

### Venta 3 — TXAR (12:00)

1. IOL → *Operar → Acciones → buscar TXAR*.
2. *Vender*.
3. **Plazo**: 48 horas.
4. **Tipo**: Limitada.
5. **Cantidad**: **10**.
6. **Precio limitado**: último × 0.995.
7. **Vigencia**: día.
8. Confirmar.

---

## 5. QUÉ HACER SI NO SE EJECUTA

### A los 60 minutos sin ejecución
- Mirar el spread bid-ask. Si está ancho (>2% entre puntas), la
  operación va a costar caro de cualquier manera.
- **Bajar el precio limitado 0.5% adicional**.

### A las 14:00 sin ejecución
- **Pasar a orden "A mercado"** SOLO si el spread es <1.5%.
- Si el spread está más ancho, **cancelar y reintentar el lunes**.

### Cierre del día sin ejecución
- Cancelar todas las órdenes pendientes (IOL las cancela automáticamente
  al cierre pero verificar).
- Anotar en el registro qué falló: ¿sin liquidez? ¿spread ancho?
  ¿precio mal puesto?
- Reintentar el lunes a primera hora con la orden ajustada.

> **No frustrarse si TXAR no se ejecuta el viernes**. Son lots chicos,
> es normal que tarden 1–2 días en colocarse. El plan no se rompe por
> 48hs de demora.

---

## 6. CIERRE DEL DÍA — 17:00

### Checklist

- [ ] Las **10 alertas** están configuradas y aparecen en *Mis
      alertas*.
- [ ] Venta GOOGL: ejecutada / pendiente / cancelada → producido ARS
      _________.
- [ ] Venta EDN: ejecutada / pendiente / cancelada → producido ARS
      _________.
- [ ] Venta TXAR: ejecutada / pendiente / cancelada → producido ARS
      _________.
- [ ] Producido total del día: ARS _________ (estimado bruto).
- [ ] Comisiones cobradas (revisar en *Mi Cuenta → Operaciones del
      día*): ARS _________.
- [ ] **Producido neto**: ARS _________ (= bruto − comisiones).
- [ ] Saldo disponible para operar: aparece T+2 = martes 10/06.

### Anotar en el registro del día

Crear archivo `04-seguimiento/registro-2026-06-06.md` con:
- Operaciones realizadas (las 3 ventas con producido neto).
- Comisión total pagada.
- PMC de las 5 posiciones principales (BMA, GGAL, GD30, NVDA, VIST).
- Alertas configuradas (sí/no).
- Plan para el lunes 09/06 (vender GD41 y PAMP).

---

## 7. QUE SIGUE — adelanto del lunes 09/06

Lunes vamos a vender GD41 (18 nominales) y PAMP (3 acciones).

| Ticker | Cantidad | Valor estimado |
|--------|----------|----------------|
| GD41   | 18       | ARS 19.980     |
| PAMP   | 3        | ARS 15.345     |
| **Total** | | **ARS 35.325** |

Acreditación: T+2 = miércoles 11/06.

Y el martes 10/06 (cuando acrediten las ventas del viernes) empezamos
a reinvertir los ARS 24.280 en PRCPPEB + GD30 + NVDA/QQQ según el
informe.

---

## 8. UN ÚLTIMO RECORDATORIO

El producido de estas 3 ventas (~24.000) es **el 4% del portfolio**.
**No vas a hacerte rico o pobre con esta operación**. El objetivo es
**limpieza y disciplina**: liberar capital atrapado en posiciones que
no mueven, configurar gestión de riesgo, y dejar la cartera lista para
los próximos 90 días.

Si una de las 3 ventas no se ejecuta hoy o el lunes, **no es un
problema**. El plan tiene margen.

**Las únicas cosas que SÍ son críticas hoy**:
1. Configurar las 10 alertas (gestión de riesgo).
2. Anotar el PMC de las 5 posiciones principales (para futuro
   tracking).
