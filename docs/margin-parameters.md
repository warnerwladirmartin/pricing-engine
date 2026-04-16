# Margin Parameters — EP_PARAMETROS_MARGEM

The EP_PARAMETROS_MARGEM spreadsheet is the single source of truth for all variable cost rates and margin policies used by the pricing engine. It has 6 tabs.

---

## Tab 1 — FRETE

Freight cost as a percentage of the invoice price, by customer state (UF).

| Column | Description |
|---|---|
| A | UF (state code, e.g., SP, RJ, MG) |
| B | Freight % (e.g., 2.5 for 2.5%) |

These rates reflect the average freight cost to deliver to each state. They are applied uniformly to all SKUs for customers in that state.

---

## Tab 2 — COMISSAO

Sales commission rates at three levels (processed in priority order by the engine):

| Column A (type) | Column B (key) | Column C (%) | Priority |
|---|---|---|---|
| `REP` | Representative code | Default % for that rep | Lowest |
| `UF` | State code | Override for that state | Medium |
| `CNPJ` | Customer CNPJ (digits only) | Override for that customer | Highest |

The `calcComissao_()` function checks CNPJ exceptions first, then UF rules, then the rep default, and finally falls back to 5%.

Use case for CNPJ exceptions: key accounts where a specific commission structure was negotiated.

Use case for UF rules: remote or high-logistics-cost states where the rep receives a different commission rate.

---

## Tab 3 — IMPOSTOS

Total tax burden as a percentage of the invoice price, by state.

| Column | Description |
|---|---|
| A | UF (state code) |
| B | Total tax burden % (all applicable taxes combined) |

This is the effective rate that reduces net revenue after taxes. It accounts for ICMS, PIS/COFINS, and other applicable levies for each state.

---

## Tab 4 — MARGEM_POLITICA

Defines minimum and target margin percentages by product hierarchy level.

| Column | Description |
|---|---|
| A | FAMILIA — product family (or blank for quadrant-level rows) |
| B | CLASSIFICACAO — commercial classification (or blank) |
| C | QUADRANTE_ESTRATEGICO — strategic quadrant code (or blank) |
| D | MARGEM_MIN% — minimum acceptable margin |
| E | MARGEM_ALVO% — commercial target margin |

The `buscarMargem_()` function matches this table with the following priority:
1. Quadrant match (col C non-blank, col A and B blank)
2. Classification match (col B non-blank)
3. Family match (col A non-blank)
4. Default: 15% / 25%

Rows can be specific (e.g., family=MOTOR OILS, classif=PREMIUM) or broad (quadrant=ESTRELA only). The first match wins.

Example structure:

| FAMILIA | CLASSIFICACAO | QUADRANTE | MIN% | ALVO% |
|---|---|---|---|---|
| | | ESTRELA | 22 | 35 |
| | PREMIUM | | 20 | 30 |
| MOTOR OILS | | | 18 | 28 |
| GREASES | ECONOMY | | 12 | 20 |
| | | | 15 | 25 |  ← catch-all default row |

---

## Tab 5 — CUSTOS_ADICIONAIS

Financial and administrative costs that are not covered by the BOM or the state-specific tabs.

| Column A (label) | Column B (value) | Description |
|---|---|---|
| CUSTO_FINANCEIRO | e.g., 1.5 | Monthly financial carrying cost % |
| CUSTO_ADMIN | e.g., 3.0 | Administrative overhead % |

The engine reads these by matching the label string (accent-insensitive). The financial cost is used in the margin formula as-is; it is not multiplied by payment term in the current version (the `calcPrazoMedio_` utility exists for future use).

---

## Tab 6 — CLASSIF_ESTRATEGICA

Maps individual SKUs to their strategic classification quadrant.

| Column | Description |
|---|---|
| A | SKU (e.g., PA00001) |
| B | Strategic quadrant code (e.g., ESTRELA, VACA_LEITEIRA, QUESTAO, ABACAXI) |

This quadrant drives the highest-priority margin lookup in `buscarMargem_()`. SKUs not listed here fall through to the classification or family level.

The quadrant codes should match the values in MARGEM_POLITICA column C.

---

## How Parameters Are Loaded

All 6 tabs are read in a single `lerParametros_()` call at the start of `carregarCliente()`. The function returns a structured object:

```javascript
{
  frete:             { UF: pct },
  comissao:          { repCode: pct },
  comissaoExcecoes:  { cnpj: pct },
  comissaoUF:        { UF: pct },
  impostos:          { UF: pct },
  margemPolitica:    [ { familia, classif, quadrante, min, alvo }, ... ],
  custosAdicionais:  { custoFin: pct, custoAdmin: pct },
  classifEstrategica: { sku: quadrante }
}
```

If the EP_PARAMETROS_MARGEM spreadsheet is unavailable, the function returns default values: all rates = 0 except commission fallback = 5%, and margins default to 15%/25%.

---

## Maintenance Notes

- Update FRETE and IMPOSTOS whenever the company's logistics or tax structure changes.
- Update COMISSAO whenever a rep contract changes or a new CNPJ exception is negotiated.
- Update MARGEM_POLITICA at the start of each pricing cycle or when commercial strategy changes.
- Update CLASSIF_ESTRATEGICA when new products are launched or the strategic portfolio review occurs (typically quarterly).
- CUSTOS_ADICIONAIS rarely changes but should be reviewed when the Selic rate moves significantly (it affects the financial carrying cost).
