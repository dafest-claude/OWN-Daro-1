# Instructivo Día 1 — Armado de cartera V0 en IOL

**Fecha objetivo: miércoles 2026-05-27**
Capital a invertir: ARS 250.000 (ajustar al monto real que tengas).

---

## CHEAT SHEET — para tener a mano durante la operación

### Tramo 1 (Día 1 — miércoles 27/05) — invertir HOY

| Orden | Tipo      | Ticker            | Monto ARS | Plazo | Tipo de orden |
|-------|-----------|-------------------|-----------|-------|----------------|
| 1     | FCI       | IOL Ahorro Plus   | 37.500    | T+0   | Suscripción    |
| 2     | Acción    | **GGAL**          | 37.500    | T+2   | Limitada       |
| 3     | Bono      | **GD30**          | 37.500    | T+2   | Limitada       |
| 4     | CEDEAR    | **NVDA**          | 31.250    | T+2   | Limitada       |

**Total Día 1: ARS 143.750** (57.5% del capital)
**Queda para Tramos 2A/2B/2C: ARS 106.250** (42.5%, repartido en 3 días)

### Tramos 2A / 2B / 2C (Días 2 a 5)

| Día            | GGAL    | GD30    | NVDA    | Total/día |
|----------------|---------|---------|---------|-----------|
| 2A — jue 28/05 | 12.500  | 12.500  | 10.400  | 35.400    |
| 2B — vie 29/05 | 12.500  | 12.500  | 10.400  | 35.400    |
| 2C — lun 01/06 | 12.500  | 12.500  | 10.450  | 35.450    |

---

## ESTA NOCHE (martes 26/05) — Preparación

1. **Verificar saldo en IOL**: entrar a la app o web, ir a *Mi cuenta* → *Saldos*. Tiene que haber al menos ARS 250.000 disponibles en pesos.
2. Si el dinero está en otra cuenta bancaria, **transferir esta noche** vía CBU/Alias de IOL para que acredite mañana antes de las 11:00. Las transferencias inmediatas suelen acreditar al instante, pero hay corte nocturno en algunos bancos: hacerlo antes de las 22:00.
3. **Activar la operatoria en cuenta comitente** si no la usás hace tiempo: a veces IOL pide confirmar perfil de inversor antes de operar bonos o CEDEARs. Hacerlo esta noche para no perder tiempo mañana.
4. **Habilitar alertas por precio** en la app: vas a usarlas como stop loss manual.

---

## MAÑANA — Cronograma del Día 1 (miércoles 27/05)

> Horario de BYMA: **11:00 a 17:00 ARG**. NO operar en los primeros 15 minutos (apertura volátil con spreads anchos).

### 09:30 – 10:45 — Pre-mercado

1. **Anotar precios de referencia** del cierre de ayer (martes 26):
   - GGAL — último ARS: ______
   - GD30 — último ARS (paridad): ______
   - NVDA — último ARS: ______

   Estos precios los ves en la pantalla de cada ticker en IOL, en *Acciones*, *Bonos* y *CEDEARs* respectivamente.

2. **Mirar premarket de Wall Street** (importa para NVDA):
   - En Google: "NVDA premarket". Si NVDA cotiza −3% o más en premarket, **demorar la carga de NVDA** hasta después de la apertura de Wall Street (15:30 ARG).
   - Si está flat o positivo, cargar NVDA dentro de la ventana normal.

3. **Calcular cantidad nominal** de cada activo con los precios de ayer:
   - GGAL: 37.500 ÷ precio = ____ acciones (redondear hacia abajo).
   - GD30: 37.500 ÷ (precio/100) ÷ 100 = ____ VN (nominales). Atención: GD30 cotiza por cada 100 VN. Si precio = 55.000, entonces 100 VN cuestan 55.000, y 37.500 te alcanzan para ~68 VN.
   - NVDA: 31.250 ÷ precio = ____ CEDEARs (redondear hacia abajo).

### 11:00 — Apertura

**NO OPERAR.** Mirar la rueda 15 minutos. Anotar:
- ¿GGAL abrió con gap >+3%? Si sí, esperar retroceso.
- ¿GD30 abrió con caída >−2%? Si sí, oportunidad.
- ¿Hay alguna noticia macro del día (IPC, BCRA, licitación)? Si sí, esperar a las 12:00.

### 11:15 – 11:30 — Suscribir FCI

Es lo primero porque es lo más simple y deja el cash trabajando desde el día 1.

**Pasos en IOL**:
1. Menú → *Fondos*.
2. Buscar **"IOL Ahorro Plus"**.
3. Botón *Suscribir*.
4. Monto: **ARS 37.500**.
5. Confirmar con clave.
6. Verificar en *Mis tenencias* que aparece la cuotaparte (puede demorar hasta el cierre del día).

### 11:30 – 12:15 — Cargar GGAL (Tramo 1)

**Pasos en IOL**:
1. Menú → *Operar* → *Acciones*.
2. Buscar **`GGAL`**.
3. *Comprar*.
4. **Plazo**: 48 horas (T+2).
5. **Tipo de orden**: Limitada.
6. **Precio**: el último visible **+ 0.4%** (para asegurar fill sin sobrepagar). Ej.: si último = ARS 5.200, poner ARS 5.220.
7. **Cantidad**: nominales calculados en pre-mercado.
8. **Vigencia**: día.
9. Confirmar con clave.

> Si en 30 minutos no se ejecutó, **subir el precio limitado un 0.3% más**. Si en una hora no se ejecuta, pasar a *A mercado* — pero solo si el spread bid-ask es <1%.

### 12:15 – 12:45 — Cargar GD30 (Tramo 1)

**Pasos en IOL**:
1. Menú → *Operar* → *Bonos*.
2. Buscar **`GD30`** (no `GD30C` ni `GD30D` — esos son las especies en USD).
3. *Comprar*.
4. **Plazo**: 48 horas (T+2).
5. **Tipo**: Limitada.
6. **Precio**: último + 0.3% (los bonos tienen spread más ajustado).
7. **Cantidad**: en VN (valor nominal). Recordá que cotiza por cada 100 VN.
8. **Vigencia**: día.
9. Confirmar.

### 12:45 – 13:15 — Cargar NVDA (Tramo 1)

> Si NVDA estaba >−3% en premarket, **saltear este paso ahora** y volver a las 15:45 (post-apertura Wall Street).

**Pasos en IOL**:
1. Menú → *Operar* → *CEDEARs*.
2. Buscar **`NVDA`**.
3. *Comprar*.
4. **Plazo**: 48 horas.
5. **Tipo**: Limitada.
6. **Precio**: último + 0.5% (CEDEARs tienen spread más ancho hasta apertura USA).
7. **Cantidad**: en CEDEARs (no en acciones reales — recordá ratio 5:1).
8. **Vigencia**: día.
9. Confirmar.

### 13:15 – 17:00 — Vigilar ejecución

- Cada 30 minutos: revisar *Operaciones del día* en IOL para confirmar que las 3 órdenes se ejecutaron.
- Si alguna sigue pendiente cerca del cierre (16:30):
  - Si es GGAL o GD30: subir precio limitado un 0.5% adicional o pasar a mercado.
  - Si es NVDA y la apertura de USA fue mala: cancelar la orden y reintentar mañana.

### 15:30 — Apertura Wall Street

- Mirar cómo abre NVDA en USA. Si abre verde, tu CEDEAR en IOL va a tener volumen y se ejecuta solo (si lo tenías pendiente).
- Si NVDA en USA cae >2% en la primera hora y tu orden todavía no se ejecutó, **cancelar** y reintentar mañana a mejor precio.

### 17:00 — Cierre y registro

1. Verificar que las 3 órdenes están **ejecutadas** (no pendientes).
2. **Configurar alertas de precio** para cada posición (proxy de stop loss):

   | Ticker | Alerta de venta (stop) | Alerta de TP parcial |
   |--------|------------------------|----------------------|
   | GGAL   | PMC × 0.85 (−15%)      | PMC × 1.25 (+25%)    |
   | GD30   | PMC × 0.92 (−8%)       | PMC × 1.15 (+15%)    |
   | NVDA   | PMC × 0.88 (−12%)      | PMC × 1.20 (+20%)    |

   En IOL: ficha del ticker → ícono de campana → *Crear alerta* → precio objetivo + canal (push/email).

3. **Anotar en `04-seguimiento/registro-diario.md`**:
   - Capital total al cierre.
   - PMC (precio medio de compra) de cada posición.
   - Cantidad nominal de cada una.
   - % real vs % objetivo de cada bloque.

---

## DÍAS 2 a 5 — Tramos 2A, 2B, 2C

### Regla general antes de cargar cada tramo

Mirar al abrir la rueda. Si la posición **subió >3% desde el PMC**, **saltear el tramo de ese día** para esa posición y meter la plata al FCI. Volver a intentar al día siguiente.

Si **cayó >5% desde el PMC sin romper stop**, podés **doblar el tramo de ese día** para esa posición (promediar a la baja agresivo), pero **solo una vez** en los 5 días.

### Jueves 28/05 — Tramo 2A

Repetir el procedimiento del Día 1 con los montos del cheat sheet:
- GGAL: 12.500
- GD30: 12.500
- NVDA: 10.400

### Viernes 29/05 — Tramo 2B

Igual. Misma asignación.

### Lunes 01/06 — Tramo 2C + Primer rebalanceo

1. Cargar el último tramo (mismos montos).
2. **Primer rebalanceo semanal** según `01-estrategia/cartera-starter-V0.md` sección "Rebalanceo semanal".
3. Calcular retorno acumulado de la primera semana. Anotarlo.

---

## Reglas durante toda la primera semana

1. **Revisar la app 3 veces al día**: 11:30, 14:00, 17:00.
2. Si una alerta de stop se dispara: **vender ese día**, no esperar al siguiente. Pasar el efectivo al FCI.
3. Si una alerta de TP parcial se dispara: **vender 1/3 de la posición** y mover la alerta de stop al PMC original (break-even).
4. NO comprar nuevos instrumentos fuera de los 4 elegidos.
5. NO aumentar el capital esta semana — el primer aporte va al cierre del mes 1 (fines de junio).

---

## Errores típicos a evitar el Día 1

| Error                                    | Cómo evitarlo                                |
|------------------------------------------|----------------------------------------------|
| Operar en los primeros 15 min            | Esperar a 11:15                              |
| Orden "a mercado" en CEDEAR              | Usar limitada siempre, especialmente CEDEARs |
| Confundir GD30 (ARS) con GD30D (USD MEP) | Verificar dos veces antes de confirmar       |
| Comprar NVDA con USA en rojo fuerte      | Mirar premarket y abrir USA antes de cargar  |
| No configurar alertas al cierre          | Cierre = anotar PMC + crear alertas, sin excepción |
| Olvidar plazo y poner T+0 en acciones    | Siempre T+2 (48 horas) en acciones, bonos y CEDEARs |

---

## Si algo no se ejecuta el Día 1

No es una urgencia. Cancelás la orden al cierre, anotás qué pasó (sin
precio, no había liquidez, etc.) y reintentás al día siguiente a primera
hora. Mantener disciplina es más importante que ejecutar al 100% el
primer día.

---

## Después del cierre — chequeo de salud del plan

Al final del miércoles, tenés que poder responder estas 5 preguntas con
la pantalla de IOL abierta:

1. ¿Cuál es mi PMC en GGAL, GD30 y NVDA?
2. ¿Cuántos nominales tengo de cada uno?
3. ¿Cuánto cash queda en pesos vs FCI?
4. ¿Las 3 alertas de stop y las 3 de TP parcial están configuradas?
5. ¿El total invertido cierra con los 143.750 esperados (±2% por slippage)?

Si las 5 están OK, **el Día 1 fue exitoso**. El próximo paso es el Día 2.
