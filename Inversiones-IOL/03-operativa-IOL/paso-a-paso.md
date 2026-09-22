# Operativa en IOL — paso a paso

Esta guía asume que ya tenés cuenta en IOL y que está validada para
operar pesos, MEP y CEDEARs.

## 1. Comprar una acción líder (ej. GGAL)

1. App o web IOL → **Operar** → **Acciones**.
2. Buscar ticker `GGAL`.
3. Plazo: **48 horas (T+2)** salvo que necesites liquidación inmediata.
4. Tipo de orden:
   - **Precio limitado**: poné el precio máximo al que estás dispuesto a
     comprar. Recomendado.
   - **A mercado**: ejecuta al mejor postor. Riesgo de slippage.
5. Cantidad: en nominales (acciones).
6. Confirmar. La orden queda en libro hasta ejecución o cierre de
   rueda.

> **Tip**: en acciones líquidas como GGAL, poné el precio limitado
> 0.3–0.5% por encima del último para asegurar ejecución sin pagar
> demasiado spread.

## 2. Comprar GD30 o AL30 en pesos

1. **Operar** → **Bonos** → buscar `GD30` o `AL30`.
2. Plazo: **48 horas en ARS**.
3. Cotización: por cada 100 nominales. Si el precio aparece como 53.000
   y querés invertir ARS 30.000 → comprás ~56 nominales.
4. Orden limitada por encima del último para asegurar fill.

## 3. Comprar CEDEAR (ej. NVDA)

1. **Operar** → **CEDEARs** → buscar `NVDA`.
2. Plazo: **48 horas**.
3. Cantidad: en CEDEARs (no en acciones reales). Recordá el ratio (NVDA
   5:1 → 5 CEDEARs equivalen a 1 NVDA real).
4. Orden limitada.

## 4. Suscribir un FCI (IOL Ahorro Plus)

1. **Fondos** → **IOL Ahorro Plus**.
2. **Suscribir** → poner el monto en ARS.
3. Confirmar. La cuotaparte se acredita al cierre del día.
4. Rescate: misma vía, **Rescatar** → llega T+0 o T+1 según fondo.

## 5. Colocar en caución

1. **Operar** → **Caución** → **Colocadora**.
2. Elegir plazo (1, 7, 14, 30 días).
3. Monto y tasa mínima aceptable (mirá la curva del día, suele estar
   2–4 puntos arriba de un FCI MM).
4. Confirmar. Al vencimiento, capital + intereses vuelven a saldo en
   pesos.

## 6. Comprar dólar MEP (si querés salir de pesos puntualmente)

1. **Dólar MEP** → IOL hace el rulo GD30/AL30 automáticamente.
2. Ingresar monto en ARS.
3. Tiempo total: 1 día hábil (compra D, venta D al día siguiente, o
   inmediato si usás la operativa "MEP en el día").
4. El USD MEP queda en tu cuenta IOL en USD.

## 7. Vender con stop loss

IOL no tiene stop loss automático en todas las especies. **Workaround**:

- **Para acciones líquidas**: dejar orden limitada de venta a precio
  stop, vigencia día. Revisar al cierre y renovar al día siguiente.
- **Para CEDEARs**: igual, orden limitada por debajo del precio.
- **Alternativa**: alerta por precio (configurable en la app) y vender
  manualmente cuando dispara.

> **Importante**: no existe stop loss garantizado en IOL como en
> brokers de USA. Operativa manual disciplinada.

## 8. Comisiones a tener presentes

| Operación             | Comisión IOL típica | Mercado | Total aprox |
|-----------------------|---------------------|---------|-------------|
| Acciones (renta var.) | 0.6%                | 0.07%   | ~0.7%       |
| Bonos                 | 0.3%                | 0.04%   | ~0.4%       |
| CEDEARs               | 0.6%                | 0.07%   | ~0.7%       |
| FCI                   | 0%                  | —       | 0%          |
| Caución colocadora    | 25% de los intereses| —       | depende     |
| Dólar MEP             | 0.5% c/lado         | —       | ~1%         |

**Costo round-trip estimado**: 1.4% en acciones/CEDEARs, 0.8% en
bonos. Tenerlo presente al fijar TP — no vender por +1% porque después
de comisiones el resultado es nulo.
