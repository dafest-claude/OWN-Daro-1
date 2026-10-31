# Tesis — IOL Portafolio Potenciado (IOLPORA)

**Fecha de documentación**: 2026-06-15
**Estado**: ACTIVA — posición a mantener
**Origen**: investigación de reglamento ejecutada según instructivo
del 15/06.

---

## 1. Datos del fondo (relevados por el usuario)

| Dato                        | Valor                              |
|-----------------------------|-------------------------------------|
| Sociedad gerente            | _por completar_                    |
| Símbolo                     | IOLPORA                            |
| Categoría / perfil          | Cuasi renta fija                   |
| Composición de cartera      | **~41% títulos públicos** + resto en otros activos (por confirmar detalle: ¿caución? ¿ON corporativas? ¿FCI subyacentes?) |
| Comisión de administración  | _por confirmar (% anual)_          |
| **Comisión de éxito**       | **0% (sin comisión de éxito)** ✅  |
| Plazo de rescate            | _por confirmar_                    |
| Apalancamiento              | _por confirmar_                    |

> Pendiente menor: completar comisión de administración anual y plazo
> de rescate cuando estés en la app. No bloquea la decisión.

---

## 2. Tesis de inversión

**Por qué se mantiene en cartera**:

1. **No es redundante** con PRCPPEB (que es Money Market puro).
   IOLPORA aporta exposición a **deuda soberana argentina con gestión
   activa**, con un perfil de volatilidad media.
2. **Costo controlado**: sin comisión de éxito → no hay efecto
   "high-water mark" que premie al gestor en años buenos sin
   penalizarlo en malos.
3. **Diversificación dentro del bloque renta fija**: PRCPPEB rinde
   tasa MM (~3% mensual), IOLPORA puede capturar movimientos de
   precio de los títulos públicos (compresión de spread, ajuste CER,
   etc.).
4. **Tamaño adecuado como satélite**: 3.7% del portfolio = pequeño
   pero relevante. No domina la cartera ni la diluye.

---

## 3. Reglas de gestión

### Tamaño objetivo
- **% objetivo en cartera**: 3–5%.
- **Tope superior**: 7% (si supera, vender el exceso).
- **Tope inferior**: 2% (si baja, no agregar más — dejar correr).

### Stop loss (ajustado al perfil real)

> Corrección sobre el instructivo previo (PMC × 0.80 era para un
> equity apalancado; al ser cuasi renta fija, el stop apropiado es
> más ajustado).

- **Stop**: caída del valor de cuotaparte de **−10% desde PMC**.
- **Razón**: si un fondo cuasi renta fija con 41% títulos públicos cae
  10%, hay un evento estructural (default soberano, reperfilamiento,
  default técnico, salto de tasa BCRA brutal). En esos escenarios
  conviene rescatar y reasignar.

### Take profit
- **No aplica TP parcial** para este perfil. Es una posición de
  acumulación, no de trade direccional.
- **Sí aplica revisión de rebalanceo**: si por crecimiento de
  cuotaparte la posición supera 7% del portfolio, vender el exceso.

### Configuración en IOL
- Crear alerta de precio: cuotaparte **menor o igual a PMC × 0.90**
  → notificación push + email.
- PMC actual: ARS **1.526** (al 03/06 y mantenido al 12/06).
- **Precio gatillo**: 1.526 × 0.90 = **ARS 1.373**.

---

## 4. Plazo y revisión

- **Plazo de tesis**: 90 días desde la documentación (próxima
  revisión obligatoria: **15/09/2026**).
- **Triggers de revisión anticipada**:
  - Cambio de sociedad gerente del fondo.
  - Modificación de la composición de cartera publicada (si pasa de
    41% títulos públicos a, por ejemplo, 60% acciones, el perfil
    cambia y la tesis deja de ser válida).
  - Aumento de comisión de administración por encima de 2.5% anual.
  - Cualquier comisión de éxito que se incorpore al reglamento.

---

## 5. Cómo se relaciona con el resto de la cartera

| Bloque                    | Instrumento principal | Rol                          |
|---------------------------|------------------------|------------------------------|
| Liquidez con tasa         | PRCPPEB                | MM puro, T+0                |
| **Renta fija mixta**      | **IOLPORA**            | **Deuda soberana gestionada**|
| Hard-dollar directo       | GD30 + AO27 + BPOC7    | Bonos USD propios            |
| Inflación pesos           | TZX26 (vence 30/06)    | CER puntual                  |

Con IOLPORA, **la cartera de renta fija queda en 3 capas**: MM puro
(reserva), gestión soberana indirecta (IOLPORA), bonos USD directos
(GD30/AO27/BPOC7). Cada una cumple un rol distinto.

---

## 6. Compromisos del usuario

- [ ] Completar los 3 datos pendientes (sociedad gerente, comisión
      administración, plazo rescate) cuando estés operando.
- [ ] Configurar alerta de stop en ARS 1.373.
- [ ] Revisar la tesis al 15/09/2026.
- [ ] Si cualquiera de los 4 triggers de §4 se dispara, abrir
      revisión inmediata.
