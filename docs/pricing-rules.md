# Pricing Rules — v3.0

## Core Formula

All suggested prices are derived from production cost using a full cost-plus formula:

```
preco = custo_prod / (1 - custo_total% - margem%)
```

Where:
- `custo_prod` = production cost from BOM cascade, including OLUC 12% (column F)
- `custo_total%` = sum of all variable cost percentages (column K):
  - Freight (G) — by customer state
  - Commission (H) — by rep / CNPJ / UF
  - Taxes (I) — by customer state
  - Financial cost (J) — monthly rate × average payment term
- `margem%` = from the margin policy hierarchy (columns L or M)

Two prices are calculated per SKU:
- **PRECO MINIMO** (col N): uses `margem_min%` — the floor; analyst should not go below this
- **PRECO SUGERIDO** (col O): uses `margem_alvo%` — the commercial target

---

## Variable Cost Breakdown

| Component | Source | Column |
|---|---|---|
| Freight | EP_PARAMETROS_MARGEM / FRETE tab, by customer UF | G |
| Commission | EP_PARAMETROS_MARGEM / COMISSAO tab (see commission engine below) | H |
| Taxes | EP_PARAMETROS_MARGEM / IMPOSTOS tab, by customer UF | I |
| Financial cost | EP_PARAMETROS_MARGEM / CUSTOS_ADICIONAIS — monthly rate | J |
| **Total** | G + H + I + J | K |

---

## Margin Hierarchy

Margins (min% and target%) are looked up with the following priority:

### Priority 1 — Strategic Quadrant (SKU-level)

The CLASSIF_ESTRATEGICA tab in EP_PARAMETROS_MARGEM maps individual SKUs to a strategic quadrant (e.g., ESTRELA, VACA_LEITEIRA, QUESTAO, ABACAXI or similar BCG-style classification).

Each quadrant has its own min/target margins defined in MARGEM_POLITICA.

If the SKU has a quadrant assignment and that quadrant has a policy row with blank family and classification, that policy is used.

### Priority 2 — Commercial Classification

The CLASSIFICACAO tab in EP_MOTOR maps SKUs to commercial tiers (e.g., PREMIUM, STANDARD, ECONOMY). If a MARGEM_POLITICA row exists for the SKU's classification, it is used.

### Priority 3 — Product Family

If no classification match, MARGEM_POLITICA is checked for the SKU's family (column C of TABELA_NOVA).

### Priority 4 — Default

If no match at any level: `min = 15%`, `alvo = 25%`.

---

## Commission Engine

Three-level lookup for each customer + product combination:

| Priority | Rule | Source |
|---|---|---|
| 1 (highest) | CNPJ exception | `params.comissaoExcecoes[cnpj]` |
| 2 | State (UF) rule | `params.comissaoUF[uf]` |
| 3 | Rep default | `params.comissao[repCode]` |
| 4 (fallback) | Hardcoded default | 5% |

Use case for CNPJ exception: key accounts negotiated a specific commission rate with their representative. Use case for UF rule: very remote states where the commercial cost structure differs.

---

## Price Base Logic

The **PRECO BASE** (column S) is the starting point before the adjustment % is applied.

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

The FONTE column (AB) flags the origin of each price:

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

## Alert System

The ALERTA column (X) signals how the proposed price compares to the margin policy:

| Alert | Condition | Recommended action |
|---|---|---|
| SEM CUSTO | `custo_prod = 0` | Resolve missing BOM before approving price |
| ABAIXO MINIMO | `NOVO_PRECO < PRECO_MINIMO` | Raise price or approve exception with reason |
| ABAIXO ALVO | `PRECO_MINIMO ≤ NOVO_PRECO < PRECO_SUGERIDO` | Review — may be acceptable for strategic accounts |
| OK | Within range | No action needed |
| PREMIUM | `NOVO_PRECO > 1.20 × PRECO_SUGERIDO` | Verify — may indicate outdated base price or aggressive adjustment |

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
