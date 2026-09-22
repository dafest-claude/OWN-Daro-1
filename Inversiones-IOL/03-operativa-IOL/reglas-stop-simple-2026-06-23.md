# Reglas STOP — versión simple

**Fecha**: 2026-06-23. Versión limpia para cargar en IOL sin
ambigüedad. Si querés más detalle, ver
`04-seguimiento/tabla-alertas-2026-06-23.md`.

---

## Las 2 tablas que necesitás

### Tabla 1 — Regla por tipo (% fijo)

| Posición | Tipo | **% STOP** | **% TP parcial** |
|----------|------|-----------|--------------------|
| BMA | Acción ARG | **−15%** | +25% |
| GGAL | Acción ARG | **−15%** | +25% |
| TGSU2 | Acción ARG | **−15%** | +25% |
| GD30 | Bono USD | **−8%** | +15% |
| AO27 | Bono USD | **−8%** | +15% |
| BPOC7 | Bono USD | **−8%** | +15% |
| NVDA | CEDEAR | **−12%** | +20% |
| COPX | CEDEAR | **−12%** | +20% |
| IOLPORA | FCI cuasi RF | **−10%** | — (no aplica) |

### Tabla 2 — Precio absoluto para cargar en IOL (desde precio 23/06)

| Posición | Precio 23/06 | **STOP en IOL** | **TP en IOL** |
|----------|--------------|-----------------|----------------|
| BMA | 14.640 | **12.444** | 18.300 |
| GGAL | 7.970 | **6.775** | 9.963 |
| TGSU2 | 9.415 | **8.003** | 11.769 |
| GD30 | 99.320 | **91.374** | 114.218 |
| AO27 | 154.900 | **142.508** | 178.135 |
| BPOC7 | 154.100 | **141.772** | 177.215 |
| NVDA | 12.950 | **11.396** | 15.540 |
| COPX | 8.805 | **7.748** | 10.566 |
| IOLPORA | 1.544 | **1.390** | — |

**Total alertas a crear**: 17 (9 stops + 8 TPs).

---

## Pasos en IOL (por cada posición)

1. *Operar → buscar ticker*.
2. Ficha del activo → 🔔 *Crear alerta*.
3. **STOP**: *"Menor o igual a"* + valor columna STOP.
4. **TP**: *"Mayor o igual a"* + valor columna TP.
5. Canal: push + email.
6. Guardar.

---

## Qué hacer cuando dispara

| Si dispara | Acción |
|------------|--------|
| **STOP** | Vender **100%** de la posición. Producido → PRCPPEB. |
| **TP** | Vender **1/3** de la posición. Mover el stop al precio de compra (break-even). |

> Las alertas IOL **no venden automáticamente**. Solo notifican. Vos
> confirmás la venta en la app.

---

## Caso COPX — atención

Precio hoy 8.805. Stop calculado: 7.748.

Pero COPX ya está en **−9.9% sobre PMC (9.770)**. Eso significa que
**si el stop fuera sobre PMC** (regla original), el gatillo es:
- PMC × 0.88 = 9.770 × 0.88 = **8.598**.
- Y hoy COPX cotiza 8.805 → falta solo −2.3% para tocar.

**Decisión esta semana**:
- **A** Vender ya (recomendado): producido ~17.610 → PRCPPEB.
- **B** Cargar alerta en **7.748** (regla nueva, desde precio hoy).
- **C** Cargar alerta en **8.598** (regla original, desde PMC).

Si no tenés tesis sobre el cobre, **Opción A**.

---

## Diferencia entre "stop desde PMC" y "stop desde precio actual"

| Método | Cómo funciona | Cuándo conviene |
|--------|---------------|------------------|
| Stop desde **PMC** | Se calcula sobre lo que pagaste. Stop fijo desde la compra. | Si querés disciplina rígida desde la entrada |
| Stop desde **precio actual** | Se recalcula con el precio del día. Se "mueve" si el activo sube. | Si querés "trailing stop" manual para proteger ganancias |

**Tabla 2 usa precio actual (más simple)**. Si preferís PMC, usá la
`tabla-alertas-2026-06-23.md` que tiene los gatillos desde PMC.

Para esta cartera y momento, **cualquiera de las dos sirve**. Lo
importante es **tenerlas configuradas**, no cuál método elegís.
