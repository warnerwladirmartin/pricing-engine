# Pricing & Margin Engine v3.0

Google Sheets + Apps Script engine for a lubricant manufacturer. Calculates production costs, applies a hierarchical margin policy, generates formatted price tables for 200+ customers, and exports them to Google Drive — replacing a fully manual process.

## What changed in v3.0 ("Motor de Margem")

| Area | v2 | v3.0 |
|---|---|---|
| TABELA_NOVA columns | ~18 | **28** (cost, margin, alert added) |
| Margin calculation | Not implemented | custo / (1 - custos% - margem%) |
| Production cost | Not integrated | Reads BOM cascade (MP→PE→PA + OLUC 12%) |
| Variable cost breakdown | Not present | Freight + commission + taxes + financial |
| Alerts | Not present | ABAIXO MINIMO / ABAIXO ALVO / OK / PREMIUM |
| Margin parameters | Not present | EP_PARAMETROS_MARGEM (6 tabs) |
| Commission engine | Not present | Per-rep → per-CNPJ exception → UF rules |
| Adjustment levels | Family only | Global + classification + family |
| New tabs | — | INSTRUCOES, CLASSIFICACAO, MODELO_REAJUSTES |

---

## TABELA_NOVA — 28-Column Layout

| Col | Letter | Field | Description |
|---|---|---|---|
| 1 | A | SKU | Internal ERP product code (hidden in export) |
| 2 | B | PRODUTO | Product description |
| 3 | C | FAMILIA | Product family grouping |
| 4 | D | CLASSIFICACAO | Commercial classification (PREMIUM, STANDARD, etc.) |
| 5 | E | UNID | Unit of measure (BD = pail, TB = drum, CX = case) |
| 6 | F | CUSTO PROD | Production cost from BOM cascade, incl. OLUC 12% |
| 7 | G | FRETE% | Freight cost % by customer state |
| 8 | H | COMISSAO% | Sales commission % (rep/CNPJ/UF logic) |
| 9 | I | IMPOSTOS% | Total tax burden % by customer state |
| 10 | J | CUSTO FINANCEIRO% | Financial carrying cost % (monthly rate) |
| 11 | K | CUSTO TOTAL% | Sum G+H+I+J |
| 12 | L | MARGEM MIN% | Minimum acceptable margin % |
| 13 | M | MARGEM ALVO% | Target margin % from margin policy |
| 14 | N | PRECO MINIMO | Floor price: custo / (1 - K% - L%) |
| 15 | O | PRECO SUGERIDO | Suggested price: custo / (1 - K% - M%) |
| 16 | P | PRECO REF | Benchmark price from EP_TABELAS_REF |
| 17 | Q | ULTIMO PRECO | Last invoiced price from sales history |
| 18 | R | TEM HIST | TRUE if customer has sales history for this SKU |
| 19 | S | PRECO BASE | Base price: last price if history exists, else ref price |
| 20 | T | REAJUSTE% | Adjustment % from REAJUSTES (global + classif + family) |
| 21 | U | NOVO PRECO | Final proposed price after adjustment |
| 22 | V | MARGEM REAL% | Actual margin % at NOVO PRECO |
| 23 | W | MARGEM REAL R$ | Actual margin R$ at NOVO PRECO |
| 24 | X | ALERTA | ABAIXO MINIMO / ABAIXO ALVO / OK / PREMIUM |
| 25 | Y | DT ULT VENDA | Date of last sale |
| 26 | Z | QTD HIST | Total volume in history period |
| 27 | AA | N VENDAS | Number of invoices in history period |
| 28 | AB | FONTE | DAGDA / HISTORICO / SEM HIST |

---

## Margin Calculation

```
preco = custo_prod / (1 - custo_total% - margem%)
```

Where `custo_total%` = freight% + commission% + taxes% + financial%

This ensures all variable costs and the desired margin are fully covered at the proposed price. Two prices are calculated per SKU: floor (minimum margin) and suggested (target margin).

---

## Architecture: 6 External Data Sources

```
EP_PARAMETROS_MARGEM
  (6 tabs: FRETE, COMISSAO, IMPOSTOS, MARGEM_POLITICA, CUSTOS_ADICIONAIS, CLASSIF_ESTRATEGICA)
        │
EP_BASE_CUSTOS / CostStructure_v2
  (BOM cascade: MP → PE → PA → RESULTADO with OLUC 12%)
        │
EP_BASE_VENDAS
  (sales history by customer CNPJ and SKU)
        │
EP_CLIENTES
  (customer master: CNPJ, PCT_FATURAMENTO, TABELA_REF, ESTADO, PRAZO, REPRESENTANTE)
        │
EP_TABELAS_REF
  (benchmark prices: tabs PADRAO, RJ, LEME, CONSUMO, VAREJO)
        │
        ▼
EP_MOTOR_PRECIFICACAO
  ├── CONFIG         (IDs, dates, customer selector B2)
  ├── REAJUSTES      (global + classification + family adjustments)
  ├── CLASSIFICACAO  (SKU → commercial classification)
  ├── TABELA_NOVA    (28-column working table)
  └── Export → Google Drive folder with fully formatted tables
```

---

## Hierarchical Margin Policy

Lookup priority (highest to lowest):

1. **Strategic quadrant** — SKU-level assignment from CLASSIF_ESTRATEGICA tab; each quadrant has its own min/target margins
2. **Product classification** — commercial tier (e.g., PREMIUM, STANDARD, ECONOMY)
3. **Product family** — family-level default (e.g., Motor Oils, Gear Oils, Greases)
4. **Global default** — 15% minimum / 25% target

---

## Commission Engine

For each customer/product:

1. Check `params.comissaoExcecoes[cnpj]` — CNPJ-level exception (highest priority)
2. Check `params.comissaoUF[uf]` — special UF rule (e.g., remote states)
3. Check `params.comissao[repCode]` — representative default
4. Fallback: 5%

---

## Cost Integration

Production costs are read from the RESULTADO tab of CostStructure_v2, column H (`custo_c_oluc`). The BOM cascade ensures any raw material price change flows automatically:

```
CUSTOS_MP (raw material costs + reajuste)
    → ESTRUTURA_PEs (semi-finished goods = sum of MPs × qty)
    → PE_RESUMO (total cost per PE = SUMIFS)
    → ESTRUTURA_PAs (finished goods = granel + packaging + aux MPs)
    → RESULTADO (production cost + OLUC if applicable)
```

OLUC (mandatory used-oil collection levy, ANP) is 12% and is already included in column H.

Of the ~612 active finished goods: 537 have a calculated cost, 75 do not (see Known Issues).

---

## Alert System

| Alert | Condition |
|---|---|
| SEM CUSTO | No production cost on file for this SKU |
| ABAIXO MINIMO | NOVO PRECO < PRECO MINIMO (below minimum margin) |
| ABAIXO ALVO | PRECO MINIMO ≤ NOVO PRECO < PRECO SUGERIDO |
| OK | NOVO PRECO within range of PRECO SUGERIDO |
| PREMIUM | NOVO PRECO > 120% of PRECO SUGERIDO |

---

## Price Base Logic

- If the customer has sales history for the SKU: `PRECO BASE = ULTIMO PRECO`
- If no history: `PRECO BASE = PRECO REF` (benchmark from EP_TABELAS_REF)
- For customers with `TABELA_REF = PADRAO`: the LEME benchmark tab is used (not PADRAO), which reflects the current commercial reference

FONTE flag distinguishes:
- `DAGDA`: sales from December 2025 onwards (post-DAGDA migration), adjusted by `PCT_FATURAMENTO`
- `HISTORICO`: sales from the pre-migration period
- `SEM HIST`: no sales history found

---

## Batch Export

`exportarTodasTabelas()` generates:
- 4 standard tables: PADRAO GERAL, PADRAO RJ, PADRAO CONSUMO, PADRAO VAREJO
- One file per customer from EP_CLIENTES

Each exported file is a standalone Google Sheets with professional formatting (brand colors, zebra rows, family separators, logo, hidden SKU column).

### VAREJO Layout

The VAREJO table uses a 4-column pricing layout:

| Column | Description |
|---|---|
| VAREJO | Base price |
| Acima 10 Vol | +10% volume tier |
| Acima 20 Vol | +20% volume tier |
| Acima 50 Vol | +50% volume tier |

QTD and VALOR TOTAL columns are omitted from the VAREJO export.

---

## Professional Formatting

| Element | Style |
|---|---|
| Title row | 16pt bold, white text on #CA4F24 |
| Customer name | 13pt bold, white text on #CA4F24 |
| Validity date | 10pt centered, white text on #CA4F24 |
| Column headers | Bold, white on #222221 |
| Even data rows | White (#FFFFFF) |
| Odd data rows | Light warm (#FFF3EB) |
| Family separators | Bold white on #CA4F24, SOLID_MEDIUM border |
| Gridlines | Hidden |
| SKU column | Always hidden |
| Company logo | Embedded top-right from Drive URL |

---

## Repository Structure

```
pricing-engine/
├── src/
│   └── main.gs                    # Full Apps Script source, sanitized
├── docs/
│   ├── architecture.md            # Flow diagram and module descriptions
│   ├── pricing-rules.md           # Business rules, formulas, edge cases
│   ├── cost-integration.md        # How the BOM cascade feeds the engine
│   └── margin-parameters.md       # EP_PARAMETROS_MARGEM structure
├── examples/
│   └── sample-output.md           # Mock data showing all 28 columns
└── README.md
```

---

## Technology Stack

| Layer | Technology |
|---|---|
| Scripting | Google Apps Script |
| Spreadsheet | Google Sheets |
| Output | Formatted Sheets file in Google Drive |
| Number parsing | Custom PT-BR locale parser (parsePreco_) |
| Accent handling | NFD normalization (removerAcentos_) |

---

## Setup

1. Copy `src/main.gs` into the Apps Script editor of your EP_MOTOR_PRECIFICACAO spreadsheet.
2. Replace all `<PLACEHOLDER>` constants at the top with the actual Google Sheets IDs.
3. Run `atualizarDropdownB2_()` once to populate the customer selector in CONFIG!B2.
4. Use the custom menu **EP Precificacao** to load customers and export tables.

---

## Known Issues

| # | Issue | Priority |
|---|---|---|
| KI-1 | **Rhino runtime deprecated** — the script runs on Google's legacy Rhino engine, which has been flagged for deprecation. Migration to V8 is required. No code changes have been made for V8 yet. | High |
| KI-2 | **75 PAs without cost structure** — 34 have historical cost but no BOM, 17 are base/bulk oils needing direct cost input, 24 are new products. All return `CUSTO_PROD = 0` and receive a `SEM CUSTO` alert. | High |
| KI-3 | **DATA_VIGENCIA may be stale** — the validity date in CONFIG!B5 should be updated at the start of each pricing cycle. As of this writing it was approximately 3 weeks behind. | High |
| KI-4 | **v3.0 code not in local backup** — the local file at `apps_script_motor.js` still contains v2. The source of truth is the Google Apps Script editor. | Medium |
| KI-5 | **ALERTA formula uses English strings** — the alert values (ABAIXO MINIMO, etc.) are written in Portuguese but the spreadsheet UI may show English strings in conditional formatting. | Low |
