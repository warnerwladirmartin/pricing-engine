# Margin Parameters — EP_PARAMETROS_MARGEM (v3.1)

The EP_PARAMETROS_MARGEM spreadsheet is the single source of truth for all variable cost rates and margin policies used by the pricing engine. It has 6 tabs.

**v3.1 updates in this file:** commission lookup is now by **rep full name** (not code), financial cost is now dynamic per customer (`monthly_rate × avg_term_months`), and the strategic-quadrant dimension now overrides the product classification / family (3D margin lookup).

---

## Tab 1 — FRETE

Freight cost as a percentage of the invoice price, by customer state (UF).

| Column | Description |
|---|---|
| A | UF (state code, e.g., SP, RJ, MG) |
| B | Freight % (e.g., 2.5 for 2.5%) |

These rates reflect the average freight cost to deliver to each state. They are applied uniformly to all SKUs for customers in that state.

---

## Tab 2 — COMISSAO (v3.1: by full name)

Sales commission rates at three levels (processed in priority order by the engine). **v3.1 change:** rep default is keyed by **full name** (e.g., "João Silva"), not by rep code.

| Column A (type) | Column B (key) | Column C (%) | Priority |
|---|---|---|---|
| `REP` | Rep full name (exact — accents, casing, spaces matter) | Default % for that rep | Lowest |
| `UF` / `ESPECIAL` | State code or special rule (e.g., `ALL_SP` adds +0.3% for specific reps) | Override / uplift | Medium |
| `CNPJ` | Customer CNPJ (digits only) | Override for that customer | Highest |

The commission engine checks CNPJ exceptions first, then special rules, then the rep default, and finally falls back to 5%.

**Common pitfall:** a trailing space or missing accent in the rep name (between EP_CLIENTES and COMISSAO) blocks the lookup and the customer gets the 5% fallback. See [troubleshooting.md §6](troubleshooting.md).

Use case for CNPJ exceptions: key accounts where a specific commission structure was negotiated. Use case for special rules: regional commercial uplifts (e.g., São Paulo special reps).

---

## Tab 3 — IMPOSTOS

Total tax burden as a percentage of the invoice price, by state.

| Column | Description |
|---|---|
| A | UF (state code) |
| B | Total tax burden % (all applicable taxes combined) |

This is the effective rate that reduces net revenue after taxes. It accounts for ICMS, PIS/COFINS, and other applicable levies for each state.

---

## Tab 4 — MARGEM_POLITICA (3D lookup)

Defines minimum and target margin percentages by product hierarchy level.

| Column | Description |
|---|---|
| A | FAMILIA — product family (or blank for higher-level rows) |
| B | CLASSIFICACAO — product classification (MINERAL, SYNTHETIC, ...) or blank |
| C | QUADRANTE_ESTRATEGICO — strategic quadrant (PROFIT DRIVER / HIDDEN STAR / CASH COW / DEAD WEIGHT) or blank |
| D | MARGEM_MIN% — minimum acceptable margin |
| E | MARGEM_ALVO% — commercial target margin |

The `buscarMargem_()` function matches this table with the following priority (**most specific wins — v3.1 3D lookup**):
1. Strategic quadrant (SKU-level via CLASSIF_ESTRATEGICA) — col C non-blank, cols A and B blank
2. Product classification — col B non-blank
3. Product family — col A non-blank
4. Default: 15% / 25%

Rows can be specific (e.g., family=MOTOR OILS, classif=SYNTHETIC) or broad (quadrant=PROFIT DRIVER only). The first match wins.

Example structure:

| FAMILIA | CLASSIFICACAO | QUADRANTE | MIN% | ALVO% |
|---|---|---|---|---|
| | | PROFIT DRIVER | 22 | 35 |
| | | HIDDEN STAR | 25 | 40 |
| | SYNTHETIC | | 20 | 30 |
| MOTOR OILS | | | 18 | 28 |
| GREASES | MINERAL | | 12 | 20 |
| | | | 15 | 25 | ← catch-all default row |

---

## Tab 5 — CUSTOS_ADICIONAIS (v3.1: dynamic financial cost)

Financial and administrative costs that are not covered by the BOM or the state-specific tabs.

| Column A (label) | Column B (value) | Description |
|---|---|---|
| CUSTO_FINANCEIRO | e.g., 1.5 | **Monthly** financial rate % |
| CUSTO_ADMIN | e.g., 3.0 | Administrative overhead % |

**v3.1 change:** the financial cost applied to each customer is `monthly_rate × avg_term_months` (from EP_CLIENTES `PRAZO`). Customers with term = `CASH` get 0%. Previously (v3.0) the rate was applied flat, regardless of term. This makes longer-term customers correctly carry more financial cost in the suggested price.

---

## Tab 6 — CLASSIF_ESTRATEGICA (v3.1 — written to TABELA_NOVA col E)

Maps individual SKUs to their strategic classification quadrant.

| Column | Description |
|---|---|
| A | SKU (e.g., PA00001) |
| B | Strategic quadrant (one of `PROFIT DRIVER`, `HIDDEN STAR`, `CASH COW`, `DEAD WEIGHT`) |

This quadrant drives the highest-priority margin lookup in `buscarMargem_()`. SKUs not listed here fall through to the classification or family level.

**v3.1 change:** the engine now writes this value into column E (`CLASSIF_ESTRATEGICA`) of TABELA_NOVA on every `carregarCliente` and `corrigirFormulas` run, so the analyst can see the quadrant alongside the price and alert without cross-referencing another tab.

**Known pending:** only 1 test SKU is currently classified in CLASSIF_ESTRATEGICA. Full classification of the SKU portfolio is a pending business task. The 3D lookup gracefully falls back to product classification / family until this is populated.

The quadrant values should match the strings in MARGEM_POLITICA column C (exact spelling and case).

---

## How Parameters Are Loaded

All 6 tabs are read in a single `lerParametros_()` call at the start of `carregarCliente()`. The function returns a structured object:

```javascript
{
  frete:              { UF: pct },                       // by state
  comissao:           { repFullName: pct },              // v3.1: keyed by full name
  comissaoExcecoes:   { cnpj: pct },                     // CNPJ-level override
  comissaoUF:         { UF: pct },                       // regional uplift (e.g., ALL_SP)
  impostos:           { UF: pct },
  margemPolitica:     [ { familia, classif, quadrante, min, alvo }, ... ],
  custosAdicionais:   { custoFin: pct, custoAdmin: pct },  // custoFin = monthly rate
  classifEstrategica: { sku: quadrante }                  // PROFIT DRIVER / HIDDEN STAR / ...
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
