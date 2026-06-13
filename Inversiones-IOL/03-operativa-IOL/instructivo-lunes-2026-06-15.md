# Instructivo — Lunes 2026-06-15

**Objetivo del día**: cerrar los 3 pendientes operativos antes del 19/06.

| # | Tarea                                | Tiempo | Prioridad |
|---|--------------------------------------|--------|-----------|
| 1 | Verificación rápida de VIST (NYSE)   | 5 min  | Media     |
| 2 | Investigar IOLPORA y decidir         | 20 min | **Alta**  |
| 3 | Actualizar 3 alertas afectadas       | 15 min | **Alta**  |
| 4 | Verificar las otras 7 alertas        | 5 min  | Media     |

Total estimado: **45 minutos** distribuidos durante el día.

---

## 1. VIST — verificación rápida (antes de la apertura, ~9:00)

### Qué hacer
1. Google → "Vista Energy stock NYSE" o ir a Yahoo Finance / Google
   Finance.
2. Anotar:
   - Precio VIST en NYSE al **03/06/2026** (cierre): USD __________
   - Precio VIST en NYSE al **12/06/2026** (cierre): USD __________
   - **Variación %**: __________

### Regla de decisión

| Variación NYSE en el período | Conclusión                          | Acción                       |
|------------------------------|-------------------------------------|------------------------------|
| Caída <2%                    | La pérdida del CEDEAR fue **MEP**   | **No hacer nada**, mantener  |
| Caída 2–7%                   | Mezcla MEP + subyacente             | Mantener, anotar en registro |
| Caída >7%                    | Subyacente cayó fuerte              | **Investigar noticias** (production miss, downgrade, balance) antes de decidir |
| Caída >10% con noticia mala  | Tesis quebrada                      | **Vender** el CEDEAR, rotar a NVDA o QQQ |

### Documentar
Anotar el resultado en `04-seguimiento/registro-2026-06-15.md` (lo
creamos al cierre).

---

## 2. IOLPORA — investigación y decisión (durante la mañana, ~11:00)

### Paso 1 — Encontrar el reglamento (5 min)

**Ruta en IOL**:
1. *Inversiones → Fondos*.
2. Buscar **IOL Portafolio Potenciado** (símbolo IOLPORA).
3. Hacer clic en el fondo → debería abrir la ficha completa.
4. Buscar pestaña/link: **"Reglamento"**, **"Información del fondo"**
   o **"Documentos"**.

Si no aparece en IOL directo:
1. Google: "IOL Portafolio Potenciado reglamento" o "IOLPORA CNV".
2. Buscar en el sitio de la sociedad gerente (puede ser **Schroder**,
   **Galileo**, **AdCap** u otra — IOL distribuye fondos de varias).

### Paso 2 — Extraer 5 datos clave (10 min)

Anotar en una hoja o directo en el registro:

| Dato                                    | Valor |
|-----------------------------------------|-------|
| Sociedad gerente                        | _______________ |
| Categoría CNV (RF / Mixto / RV / Otros) | _______________ |
| Política de inversión (resumen 2 líneas) | _______________ |
| Comisión de administración anual %      | _______________ |
| Comisión de éxito (si tiene)            | _______________ |
| Plazo de rescate                        | T+__ |
| ¿Permite apalancamiento? S/N            | _______________ |
| Patrimonio bajo administración (ARS)    | _______________ |
| Rendimiento últimos 12 meses %          | _______________ |

### Paso 3 — Decisión según hallazgos

Aplicar este árbol en orden:

```
¿La comisión total (gestión + éxito) supera 3% anual?
├── SÍ → VENDER (no compensa)
└── NO → seguir
    │
    ¿La política de inversión usa apalancamiento real (leverage > 1)?
    ├── SÍ → MANTENER como satélite agresivo (3-5% máx) y
    │        documentar tesis. Poner alerta de stop:
    │        PMC × 0.80 (cae 20%) → vender.
    │
    └── NO → ¿Es un fondo mixto/RV que no agrega vs cartera?
             ├── SÍ (redundante) → VENDER, rotar a PRCPPEB
             └── NO (estrategia distinta y única) → mantener,
                 documentar tesis con plazo y objetivo
```

### Paso 4 — Si la decisión es VENDER

**Cómo rescatar un FCI en IOL**:
1. *Inversiones → Fondos → Mis tenencias*.
2. IOLPORA → botón **Rescatar**.
3. **Tipo de rescate**: total (todas las cuotapartes).
4. Confirmar con clave.
5. Acreditación: T+__ según reglamento (típicamente T+0 a T+2).

**Reasignación del producido (~ARS 21.400)**:
- 100% a **PRCPPEB**.
- Esto lleva la liquidez de 13.9% a ~17.5%, **por encima del objetivo
  del 15%** — perfecto para tener munición.

### Paso 5 — Si la decisión es MANTENER

Documentar en `01-estrategia/tesis-IOLPORA.md`:

```
# Tesis IOLPORA

Sociedad gerente: ___
Categoría: ___
Por qué lo mantengo: ___
% objetivo en cartera: 3-5%
Precio de entrada (cotización inicial PMC): ___
Stop loss: PMC × 0.80 = ___
Take profit parcial: PMC × 1.30 = ___
Plazo de revisión: 30 días (próximo: 15/07/2026)
```

Y agregar la alerta de stop en IOL.

---

## 3. Actualización de alertas — 3 que cambiaron (~14:00)

Compraste más unidades de GD30, NVDA y TGSU2 el 09-10/06. El **PMC se
movió**. Las alertas viejas (que se calculaban sobre el PMC anterior)
**quedaron mal calibradas**.

### Paso 1 — Obtener PMC actualizado (5 min)

**Ruta en IOL**: *Mi Cartera → Detalle de tenencia*. Anotar el PMC de
cada una:

| Ticker | PMC anterior (antes del rebalanceo) | PMC actual | Nuevos gatillos |
|--------|--------------------------------------|------------|------------------|
| GD30   | _______ | _______ | Stop: PMC × 0.92 = _______ <br> TP: PMC × 1.15 = _______ |
| NVDA   | _______ | _______ | Stop: PMC × 0.88 = _______ <br> TP: PMC × 1.20 = _______ |
| TGSU2  | _______ | _______ | Stop: PMC × 0.85 = _______ <br> TP: PMC × 1.25 = _______ |

### Paso 2 — Borrar alertas viejas (5 min)

**Ruta en IOL**: *Mi cuenta → Alertas* (o ícono de campana en el menú
lateral).

Para cada una de las 6 alertas viejas (Stop y TP de GD30, NVDA, TGSU2):
1. Seleccionar la alerta.
2. **Eliminar** o **Desactivar**.
3. Confirmar.

### Paso 3 — Crear las 6 alertas nuevas (5 min)

Para cada ticker (GD30, NVDA, TGSU2):
1. Ir a la ficha del ticker.
2. Botón campana 🔔 → *Crear alerta*.
3. Tipo: **Precio**.
4. **Stop**: Condición *Menor o igual* + precio gatillo (col. derecha de
   la tabla de arriba).
5. **TP**: Condición *Mayor o igual* + precio gatillo.
6. Canal: Push (app) + email.
7. Guardar.

---

## 4. Verificación de las otras 7 alertas (~15:00)

Las alertas de **BMA, GGAL, VIST** no necesitan recalcular (no
agregaste posición). Pero conviene verificar que **siguen activas** —
IOL a veces las cancela tras un período o cambio de versión.

### Checklist rápido

- [ ] BMA Stop activa
- [ ] BMA TP activa
- [ ] GGAL Stop activa
- [ ] GGAL TP activa
- [ ] VIST Stop activa
- [ ] VIST TP activa
- [ ] (si IOLPORA se mantiene) IOLPORA Stop activa

Si alguna no aparece en *Mis alertas*, recrearla con los gatillos del
informe del 05/06.

---

## 5. Registro del día — al cierre (17:00)

Crear archivo `04-seguimiento/registro-2026-06-15.md` con:

```markdown
# Registro 2026-06-15 — Cierre pendientes operativos

## VIST — verificación NYSE
- Precio NYSE 03/06: USD ___
- Precio NYSE 12/06: USD ___
- Variación: ___%
- Conclusión: ___ (MEP / subyacente / mezcla)
- Acción: ___

## IOLPORA — investigación
- Sociedad gerente: ___
- Categoría CNV: ___
- Comisión anual: ___%
- Apalancamiento: ___
- Decisión: MANTENER / VENDER
- Si vendí: producido ARS ___ → reasignado a ___
- Si mantengo: tesis documentada en `01-estrategia/tesis-IOLPORA.md`

## Alertas actualizadas
- GD30: nuevo Stop ___ / TP ___
- NVDA: nuevo Stop ___ / TP ___
- TGSU2: nuevo Stop ___ / TP ___

## Verificación alertas existentes
- BMA, GGAL, VIST: todas activas ✅/❌

## Estado de cartera al cierre
- Valor total: ARS ___
- Liquidez (PRCPPEB + cash): ___%
- Cash residual: ARS ___
```

---

## 6. Resumen de prioridades

Si solo tenés 30 minutos hoy, hacé en este orden:

1. **Actualizar las 3 alertas** (15 min) — gestión de riesgo desactualizada es peor que sin alertas.
2. **Investigar IOLPORA** y decidir (15 min) — pendiente desde hace 10 días.

Si tenés más tiempo, sumar:

3. Verificación NYSE de VIST (5 min).
4. Verificación de las otras 7 alertas (5 min).

---

## 7. Lo que NO hay que hacer hoy

- **No comprar nada nuevo.** El próximo movimiento de compra es el
  aporte de fin de mes (30/06).
- **No vender BMA, GGAL ni TGSU2** aunque venga otra suba grande hoy.
  Si toca TP en alerta, ahí sí (vender 1/3 y mover stop a break-even).
- **No tomar decisiones reactivas** por movimientos diarios. La cartera
  no necesita ajuste estratégico, solo cerrar los pendientes operativos.

---

## 8. Próximos hitos

| Fecha    | Hito                                       |
|----------|---------------------------------------------|
| 19/06    | Deadline informal IOLPORA decidido y documentado |
| 26/06    | Revisión semanal del lunes (compliance reglas) |
| 30/06    | Vencimiento TZX26 + aporte mensual         |
| 03/07    | Cierre del mes 1 — calcular retorno real   |
