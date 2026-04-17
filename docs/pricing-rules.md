# Pricing Rules — v3.1

## Core Formula

All suggested prices are derived from production cost using a full cost-plus formula:

```
preco = custo_prod / (1 - custo_total% - margem%)
```

Where:
- `custo_prod` = production cost from BOM cascade, including OLUC 12% = 11% oil + 1% packaging (column G)
- `custo_total%` = sum of all variable cost percentages (column L):
  - Freight (H) — by customer state
  - Commission (I) — by rep full name / CNPJ / UF
  - Taxes (J) — by customer state
  - Financial cost (K) — monthly rate × average payment term (v3.1: dynamic per customer)
- `margem%` = from the margin policy 3D hierarchy (columns M or N)

Two prices are calculated per SKU:
- **PRECO MINIMO** (col O): uses `margem_min%` — the floor; analyst should not go below this
- **PRECO SUGERIDO** (col P): uses `margem_alvo%` — the commercial target

> **v3.1 column shift note:** column E is now `CLASSIFICACAO ESTRATEGICA` (strategic quadrant). All columns from F onward shifted one position to the right versus v3.0.

---

## Variable Cost Breakdown (v3.1 columns)

| Component | Source | Column |
|---|---|---|
| Freight | EP_PARAMETROS_MARGEM / FRETE tab, by customer UF | H |
| Commission | EP_PARAMETROS_MARGEM / COMISSAO tab (by rep full name; see commission engine below) | I |
| Taxes | EP_PARAMETROS_MARGEM / IMPOSTOS tab, by customer UF | J |
| Financial cost | EP_PARAMETROS_MARGEM / CUSTOS_ADICIONAIS — `monthly_rate × avg_term_months` (0 if term = CASH) | K |
| **Total** | H + I + J + K | L |

---

## Margin Hierarchy — 3D lookup (v3.1)

Margins (min% and target%) are looked up with the following priority. Most specific wins — strategic class **overrides** classification, which overrides family, which overrides the default.

### Priority 1 — Strategic Quadrant (SKU-level)

The CLASSIF_ESTRATEGICA tab in EP_PARAMETROS_MARGEM maps individual SKUs to one of four strategic quadrants:

- **PROFIT DRIVER** — high volume, high margin SKUs; defend margin aggressively
- **HIDDEN STAR** — high margin but low volume; protect and grow
- **CASH COW** — high volume, low margin; keep stable
- **DEAD WEIGHT** — low volume, low margin; candidate for discontinuation or price increase

Each quadrant has its own min/target margins defined in MARGEM_POLITICA. If a MARGEM_POLITICA row exists with this quadrant and blank family/classification, it wins the lookup.

### Priority 2 — Product Classification

MINERAL, SYNTHETIC, SEMI-SYNTHETIC, GREASE, AQUEOUS, BASE OILS — mapped to SKUs via the CLASSIFICACAO tab in EP_MOTOR.

### Priority 3 — Product Family

HYDRAULICS, MOTORS, TRACTORS, etc. (column C of TABELA_NOVA).

### Priority 4 — Default

If no match at any level: `min = 15%`, `alvo = 25%`.

> **Note on "most specific wins":** The current implementation is *override* — the strategic class fully replaces the family/classification margin. An alternative under study is an *additive delta* model (`PROFIT DRIVER = base + 5pp`). Business decision pending.

---

## Commission Engine (v3.1)

Lookup for each customer + product combination — **by rep full name**, not by code.

| Priority | Rule | Source |
|---|---|---|
| 1 (highest) | CNPJ exception | `params.comissaoExcecao[cnpj]` |
| 2 | Per-rep default (by full name) | `params.comissaoPadrao[rep_full_name]` |
| 3 | Special regional rule (e.g., `ALL_SP` adds +0.3% for specific reps) | `params.comissaoEspecial[]` |
| 4 (fallback) | Hardcoded default | 5% |

**v3.1 change:** the lookup key for rep defaults is the full name (e.g., "João Silva"), not a code. Exact match is required — trailing spaces or accent differences block the lookup. See [troubleshooting.md §6](troubleshooting.md) for the common mismatch pattern.

Use case for CNPJ exception: key accounts negotiated a specific commission rate with their representative. Use case for regional rules: state-level commercial uplifts (e.g., São Paulo special reps).

---

## Price Base Logic

The **PRECO BASE** (column T in v3.1) is the starting point before the adjustment % is applied.

```
if TEM_HIST = TRUE:
    PRECO_BASE = ULTIMO_PRECO  (last invoiced price from EP_BASE_VENDAS)
else:
    PRECO_BASE = PRECO_REF     (benchmark from EP_TABELAS_REF)
```

### Reference Table Selection

The benchmark reference depends on the customer's `TABELA_REF` field in EP_CLIENTES:

| TABELA_REF value | Tab used in EP_TABELAS_REF | Note |
|---|---|---|
| PADRAO | **LEME** | Uses LEME tab (commercial benchmark), not PADRAO |
| RJ | RJ | Regional pricing for Rio de Janeiro |
| CONSUMO | CONSUMO | Consumption-based pricing |
| VAREJO | VAREJO | Retail pricing with 4-tier layout |

This PADRAO → LEME mapping is intentional: the LEME tab reflects the current commercial reference price and is more appropriate as a base for standard-contract customers.

---

## DAGDA vs Historical Source Distinction

The FONTE column (AC in v3.1) flags the origin of each price:

| FONTE value | Meaning |
|---|---|
| DAGDA | Sale recorded via the current ERP (post-December 2025 migration); price adjusted by PCT_FATURAMENTO |
| HISTORICO | Sale from the pre-migration period; price recorded as-is |
| SEM HIST | No sales history found for this SKU + customer combination |

`PCT_FATURAMENTO` is stored in EP_CLIENTES and represents the billing fraction (e.g., 0.85 means the customer pays 85% of list price). This field is stored for reference but is **NOT applied to the price shown in the table**. The table always shows the full invoice price (preco NF cheio).

---

## Adjustment Percentages

Adjustments are read from the REAJUSTES tab and applied additively:

```
pct_reajuste = global_pct
             + classif_pct   (if this SKU's classification has an entry)
             + familia_pct   (if this SKU's family has an entry)

NOVO_PRECO = PRECO_BASE * (1 + pct_reajuste / 100)
```

The REAJUSTES tab uses three row types (column A):
- `GLOBAL` — applies to every SKU
- `CLASSIF` — applies to SKUs with the named classification
- `FAMILIA` — applies to SKUs in the named family

All three levels accumulate. A SKU can receive all three adjustments simultaneously.

---

## Alert System (v3.1 — tighter OK band)

The ALERTA column (Y in v3.1) signals how the proposed price compares to the margin policy:

| Alert | Condition | Recommended action |
|---|---|---|
| SEM CUSTO | `custo_prod = 0` | Resolve missing BOM before approving price |
| ABAIXO MINIMO | `NOVO_PRECO < PRECO_MINIMO` | Raise price or approve exception with reason |
| ABAIXO ALVO | `PRECO_MINIMO ≤ NOVO_PRECO < PRECO_SUGERIDO` | Review — may be acceptable for strategic accounts |
| OK | `PRECO_SUGERIDO ≤ NOVO_PRECO ≤ 1.15 × PRECO_SUGERIDO` | No action needed |
| PREMIUM | `NOVO_PRECO > 1.15 × PRECO_SUGERIDO` | Verify — may indicate outdated base price or aggressive adjustment |

**v3.1 change:** the PREMIUM threshold was tightened from 1.20× to **1.15×** the suggested price. The OK band is now narrower — more prices now trigger the PREMIUM flag for review. The multiplier is centralized in the `ALERT_PREMIUM_MULTIPLIER` constant at the top of `main.gs`.

---

## Export Layouts

### Wholesale (Standard)

Used for most customers. Single price column (NOVO_PRECO). SKU hidden.

Exported columns: PRODUTO, UNID, PRECO, FAMILIA (as separator rows)

### VAREJO (Retail)

Used for PADRAO VAREJO and retail accounts. Four price columns. SKU hidden. QTD and VALOR TOTAL columns are omitted.

| Column | Calculation |
|---|---|
| VAREJO | NOVO_PRECO (base) |
| Acima 10 Vol | NOVO_PRECO × 1.10 |
| Acima 20 Vol | NOVO_PRECO × 1.20 |
| Acima 50 Vol | NOVO_PRECO × 1.50 |

### Exported File Header

All exports include a 3-row header:
- Row 1: "Tabela de Precos [Company] Lubrificantes" — 16pt bold (brand color)
- Row 2: Customer name in title case — 13pt bold (brand color)
- Row 3: DATA_VIGENCIA from CONFIG!B5 — centered (brand color)

---

## Safety Notes

1. The `corrigirFormulas()` function ("Recalcular Margens") only recomputes margin/cost/alert columns. It does not reload prices or history from external sheets. Use it after changing margin policy or adjustment percentages.

2. There is no automatic audit trail. Recommended practice: before each cycle, record the DATA_VIGENCIA and take a copy of the REAJUSTES tab.

3. The Rhino runtime deprecation may affect scheduled or triggered executions. Monitor Google's deprecation timeline and plan V8 migration.
