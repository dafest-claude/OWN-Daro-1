# Inversiones IOL — Plan de trabajo

Carpeta de trabajo para diseñar y operar una estrategia de inversión en
Argentina vía **InvertirOnline (IOL)** con un capital inicial de
**ARS 200.000 – 300.000**.

---

## 1. Premisas del usuario (declaradas)

- Capital inicial: **ARS 200.000 a 300.000**.
- Plataforma operativa: **IOL (InvertirOnline)** — cuenta ya abierta.
- Objetivo declarado: **duplicar el capital cada mes**.
- Mercado: Argentina (pesos y dólar MEP/CCL vía IOL).

## 2. Premisas técnicas (a validar con el usuario)

> Este bloque existe porque el objetivo declarado (100% mensual sostenido)
> no es alcanzable de forma replicable en ningún mercado. Antes de operar
> hay que fijar un objetivo y un nivel de riesgo realistas.

| Perfil           | Retorno mensual objetivo | Drawdown esperable | Instrumentos típicos |
|------------------|--------------------------|--------------------|----------------------|
| Conservador      | 3 – 6 %                  | < 5 %              | FCI money market, cauciones, LECAP, ONs cortas |
| Moderado         | 6 – 12 %                 | 10 – 20 %          | Bonos soberanos ARS/USD, CEDEARs líderes, dólar MEP |
| Agresivo         | 12 – 25 %                | 30 – 50 %          | CEDEARs alta beta, acciones líderes argentinas, bonos largos |
| Especulativo     | 25 – 60 % (no sostenible)| 50 – 100 %         | Opciones (calls/puts), apalancamiento con caución, rulos MEP |

**Importante:** ningún perfil garantiza el retorno. La columna
"drawdown" indica pérdidas posibles en un mes adverso.

## 3. Instrumentos disponibles en IOL

### Renta fija en pesos
- **Cauciones bursátiles** (1 a 120 días) — tasa segura, liquidez diaria.
- **FCI Money Market** (ej. IOL Ahorro Plus) — rescate inmediato.
- **LECAP / LECER / BONCER** — letras y bonos del Tesoro en pesos.
- **Obligaciones Negociables (ON) en pesos**.

### Renta fija dólar
- **Bonos soberanos hard-dollar**: AL29, AL30, AL35, GD30, GD35, GD41.
- **ON corporativas USD**: YPF, Pampa, TGS, Vista, etc.
- **BOPREAL** (BCRA).

### Dólar financiero
- **MEP** vía AL30/GD30 (paridad D y CI).
- **CCL** vía GD30/CEDEARs.

### Renta variable
- **Acciones argentinas líderes (Merval)**: GGAL, YPFD, PAMP, BMA, ALUA, TXAR, COME, TGSU2, MIRG, CRES.
- **CEDEARs**: AAPL, MSFT, NVDA, TSLA, KO, JPM, SPY, QQQ, etc. (ratio + dólar implícito).

### Derivados
- **Opciones sobre acciones líderes** (calls y puts, lanzamiento cubierto / venta de puts).
- **Plazo firme vs. caución**: arbitraje de tasa.

## 4. Información que necesito del usuario para avanzar

1. **Horizonte real**: ¿1 mes, 3 meses, 12 meses?
2. **Tolerancia a pérdida**: ¿cuánto del capital estás dispuesto a perder en el peor mes? (5 %, 20 %, 50 %, 100 %).
3. **Disponibilidad operativa**: ¿podés operar diariamente o solo semanalmente?
4. **Experiencia previa**: ¿operaste opciones, bonos, MEP, CEDEARs?
5. **Origen del capital**: ¿es ahorro disponible o tiene un uso comprometido a corto plazo?
6. **Aporte mensual adicional**: ¿podés sumar pesos cada mes o el capital es fijo?
7. **Objetivo final en pesos o en dólares**: ¿pensás en ARS o querés dolarizar?

## 5. Estructura de la carpeta (a completar en pasos siguientes)

```
Inversiones-IOL/
├── README.md                 ← este archivo
├── 01-estrategia/            ← plan por perfil elegido
├── 02-instrumentos/          ← fichas técnicas de cada instrumento
├── 03-operativa-IOL/         ← cómo ejecutar cada operación en la app
├── 04-seguimiento/           ← planillas y registro mensual
└── 05-escenarios/            ← simulaciones de retorno y drawdown
```

## 6. Estado actual del plan

| Versión | Archivo                                       | Estado     |
|---------|-----------------------------------------------|------------|
| V0      | `01-estrategia/cartera-starter-V0.md`         | **ACTIVA** |
| V1      | `01-estrategia/plan-agresivo-3M.md`           | en reserva |

**Arrancamos con V0**: 4 instrumentos (GGAL, GD30, NVDA, FCI MM) sobre
IOL. Después de 3–4 semanas, si los KPI se cumplen, escalamos a V1.
