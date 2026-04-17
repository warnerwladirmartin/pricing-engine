# Pricing & Margin Engine v3.1

> Google Sheets + Apps Script engine for a lubricant manufacturer. Calculates production costs, applies a hierarchical margin policy, generates formatted price tables, and exports them to Google Drive — replacing a fully manual process. **v3.1** adds a strategic-classification dimension, sharper alert thresholds, and full operational documentation.

## What's new in v3.1 (vs v3.0)

| Aspect | v3.0 | v3.1 |
|---|---|---|
| Columns in TABELA_NOVA | 28 | **29** (added `CLASSIFICACAO ESTRATEGICA` as column E) |
| Margin lookup | 2D (classification + family) | **3D** — Strategic Class > Product Class > Family |
| Alert `PREMIUM` threshold | ≥ 1.20× suggested | **≥ 1.15× suggested** (tighter OK band) |
| Commission lookup | by rep code | **by rep full name** (26 reps + CNPJ exceptions + special rules like `ALL_SP`) |
| Financial cost | flat monthly rate | **dynamic**: `monthly_rate × average_term_months` (respects "CASH" = 0%) |
| OLUC detail | "12%" mention | **11% oil + 1% packaging = 12%** (ANP breakdown) |
| Troubleshooting | not documented | **9 common issues** with root cause + fix ([docs/troubleshooting.md](docs/troubleshooting.md)) |
| Daily workflows | not documented | **5 scenarios A–E** step-by-step ([docs/daily-workflow.md](docs/daily-workflow.md)) |
| Formula memory | partial | **numeric walkthrough** for every key formula ([docs/MANUAL_v3.1.md §10](docs/MANUAL_v3.1.md)) |
| PAs without BOM | 75 | **604** (awaiting BOM coordinator meeting) |
| Complete manual | — | **14-section operational manual** ([docs/MANUAL_v3.1.md](docs/MANUAL_v3.1.md)) |

## Overview

Each customer receives a price table derived from one of the reference tables (PADRAO, RJ, LEME, CONSUMO, VAREJO). A pricing analyst selects the customer in cell B2, the engine loads the customer context (state, incoterm, term, salesperson, reference table, billing %), pulls live production cost from a 5-level BOM cascade, applies the margin policy from a rules workbook, and generates:

- **Internal view (29 columns):** cost, variable-cost breakdown, minimum price, suggested price, real margin, alert status.
- **External export:** clean customer-facing Google Sheets with only price, family, unit — **no margin, no cost, no alerts**.

## Key capabilities

- **Real-cost-aware pricing.** Production cost comes from a 5-tab BOM cascade (MP → PE → PA → RESULTADO) including the ANP regulatory OLUC surcharge (12% = 11% oil + 1% packaging) when applicable.
- **Multiplicative adjustment policy.** 4 independent levels (global, classification, family, table/customer) compose as `(1+g)(1+c)(1+f)(1+t) - 1` — captures compound effect, not linear sum.
- **Minimum and suggested price formula.** `price = cost / (1 - variable_costs% - margin%)` — variable costs and margin are percentages of sale price, not cost, so the formula isolates the price algebraically.
- **3D margin lookup with specificity priority.** `buscarMargem_(family, productClass, strategicClass)` walks from most specific to most general; the most specific hit wins.
- **State-aware freight and taxes.** Freight % and tax burden are looked up by the customer's state (with abbreviation conversion for legacy full-name entries).
- **Commission engine with exceptions.** Default by salesperson → override by CNPJ → special regional rules (e.g., `ALL_SP` adds +0.3% for specific reps).
- **Alert system.** Every NEW PRICE is classified as BELOW MIN (red), BELOW TARGET (yellow), OK (green), or PREMIUM (blue, > 1.15× suggested).
- **Batch export.** One click generates all 4 standard tables + every registered customer, writes to a timestamped Google Drive folder.
- **Brand-consistent formatting.** Embedded logo, brand colors, zebra striping, family separators, hidden auxiliary columns.

## TABELA_NOVA — 29-column layout

| Block | Cols | Contents |
|---|---|---|
| 1. Identification | A–F | SKU, product name, family, classification, **strategic class** (v3.1), unit |
| 2. Cost & margin | G–N | Production cost, variable-cost breakdown (freight, commission, taxes, financial), total cost %, min margin, target margin |
| 3. Calculated prices | O–P | Minimum price, suggested price |
| 4. History & reference | Q–S | Reference price, last actual price, "has history" flag |
| 5. Adjustment & new price | T–V | Base price, accumulated adjustment %, new price |
| 6. Real margin & alert | W–Y | Real margin %, real margin R$, alert status |
| 7. Sales statistics | Z–AC | Last sale date, historical qty, invoice count, source system (SAP/DAGDA) |

See [docs/MANUAL_v3.1.md §4](docs/MANUAL_v3.1.md) for column-by-column detail.

## Core formulas

```
TOTAL COST %   = FREIGHT % + COMMISSION % + TAXES % + FINANCIAL COST %
MIN PRICE      = PROD COST / (1 - TOTAL COST % - MIN MARGIN %)
SUGGESTED      = PROD COST / (1 - TOTAL COST % - TARGET MARGIN %)
ADJUSTMENT %   = (1+GLOBAL) × (1+CLASSIF) × (1+FAMILY) × (1+TABLE) - 1
NEW PRICE      = BASE PRICE × (1 + ADJUSTMENT %)
REAL MARGIN %  = (NEW PRICE - PROD COST) / NEW PRICE - TOTAL COST %
```

**Numeric example:** Cost = R$ 100, Variable costs = 30%, Min margin = 15% → Min price = 100 / 0.55 = **R$ 181.82**. Check: 181.82 × 30% = 54.55 variable + 100 cost = 154.55; remainder 27.27 = 15% of 181.82 ✅

## Alert thresholds

| Condition | Alert | Color |
|---|---|---|
| V < O | BELOW MIN | 🔴 Red |
| O ≤ V < P | BELOW TARGET | 🟡 Yellow |
| P ≤ V ≤ P × 1.15 | OK | 🟢 Green |
| V > P × 1.15 | PREMIUM | 🔵 Blue |

## Architecture: 6 external data sources

```
EP_MOTOR_PRECIFICACAO (the brain)
    │
    ├── EP_PARAMETROS_MARGEM      — freight, commission, taxes, margin policy, strategic class
    ├── EP_TABELAS_REF             — standard reference prices (PADRAO, RJ, LEME, CONSUMO, VAREJO)
    ├── EP_CLIENTES                — lean customer registry (state, incoterm, term, rep, table ref)
    ├── EP_BASE_VENDAS             — sales history (SAP + DAGDA) for last-price and stats
    └── CostStructure_v2           — BOM cascade producing per-PA cost with OLUC
```

The script reads all via `SpreadsheetApp.openById()` — never via `IMPORTRANGE` in columns it writes to.

## Hierarchical margin policy (3D)

Lookup priority (most specific wins):

1. **Strategic quadrant** (col E, new in v3.1) — PROFIT DRIVER / HIDDEN STAR / CASH COW / DEAD WEIGHT
2. **Product classification** — MINERAL, SYNTHETIC, SEMI-SYNTHETIC, GREASE, AQUEOUS, BASE OILS
3. **Product family** — HYDRAULICS, MOTORS, TRACTORS, etc.
4. **Global default** — 15% minimum / 25% target

## Commission engine

For each customer:

1. Check `params.comissaoExcecao[cnpj]` — CNPJ-level exception (highest priority)
2. Check `params.comissaoPadrao[rep_full_name]` — per-rep default (by full name, not code)
3. Apply `params.comissaoEspecial[]` rules — regional uplifts (e.g., `ALL_SP` adds +0.3% for specific reps)
4. Fallback: 5%

See [docs/MANUAL_v3.1.md §10.4](docs/MANUAL_v3.1.md) for the full commission algorithm.

## Cost integration

Production costs are read from the RESULTADO tab of CostStructure_v2, column H (`custo_c_oluc`). The BOM cascade ensures any raw-material price change flows automatically:

```
CUSTOS_MP (raw material costs + reajuste)
    → ESTRUTURA_PEs (semi-finished = sum of MPs × qty)
    → PE_RESUMO (total cost per PE = SUMIFS)
    → ESTRUTURA_PAs (finished = bulk + packaging + aux MPs)
    → RESULTADO (production cost + OLUC if applicable)
```

OLUC (mandatory used-oil collection levy, ANP regulatory) = **12% = 11% oil + 1% packaging**, applied only to SKUs with `ies_disp_coleta = 'N'` in the ERP (i.e., excluding greases, 2-stroke, chainsaw, agricultural oils).

## Known issues / open items

| # | Issue | Priority |
|---|---|---|
| KI-1 | **Rhino runtime deprecated** — script runs on Google's legacy Rhino engine; migration to V8 pending. No breaking changes yet, but should be scheduled. | High |
| KI-2 | **604 PAs without BOM** — these return `CUSTO_PROD = 0` and cannot have their minimum/suggested price calculated. Awaiting BOM coordinator meeting. | High |
| KI-3 | **DATA_VIGENCIA may be stale** — validity date in CONFIG B10 must be updated each pricing cycle; no automatic rollover. | Medium |
| KI-4 | **CLASSIF_ESTRATEGICA tab has only 1 test SKU** — full classification of the portfolio into the 4 strategic quadrants is a pending business task. 3D lookup gracefully falls back to product class / family. | Medium |
| KI-5 | **3D margin lookup uses "override"** — most specific wins fully. Alternative under study: strategic class as additive delta (`PROFIT DRIVER = base + 5pp`). Business decision pending. | Low |

## Repository structure

```
pricing-engine/
├── src/
│   └── main.gs                    # Apps Script source (sanitized)
├── docs/
│   ├── MANUAL_v3.1.md             # Complete operational manual (14 sections)
│   ├── architecture.md            # Flow diagram and module descriptions
│   ├── pricing-rules.md           # Business rules, formulas, edge cases
│   ├── cost-integration.md        # How the BOM cascade feeds the engine
│   ├── margin-parameters.md       # EP_PARAMETROS_MARGEM structure
│   ├── troubleshooting.md         # 9 common errors + fixes
│   └── daily-workflow.md          # 5 step-by-step operational scenarios
├── examples/
│   └── sample-output.md           # Mock 29-column output row
├── .gitignore
└── README.md
```

## Technology stack

| Layer | Technology |
|---|---|
| Scripting | Google Apps Script (Rhino, migration to V8 pending) |
| Spreadsheet | Google Sheets (6 linked workbooks) |
| Output | Formatted Google Sheets in Drive + PDF export |
| Data sources | SAP sales history, internal ERP (DAGDA / PostgreSQL), internal cost structure workbook |
| Number parsing | Custom PT-BR locale parser (`parsePreco_`) |
| Accent handling | NFD normalization (`removerAcentos_`) |

## Setup

1. Create the 6 Google Sheets workbooks following the structure in [docs/architecture.md](docs/architecture.md).
2. Copy `src/main.gs` into the Apps Script editor of EP_MOTOR_PRECIFICACAO.
3. Replace all `<PLACEHOLDER>` constants with real Sheet IDs.
4. Fill the CONFIG tab with the 5 sheet IDs and parameters (logo URL, history cutoff, validity date, monthly financial rate).
5. Populate CLASSIFICACAO tab with SKU → classification mapping.
6. Populate REAJUSTES tab with initial policy (at least global % in B5).
7. Run `atualizarDropdownB2_()` once via the custom menu to populate the customer selector.
8. Select a customer in B2 → **Load Customer / Table** → **Recalculate Margins** → review alert column → **Export current table (Drive)**.

## Privacy

All real Sheet IDs, customer names, financial values, and internal personnel names have been replaced with sanitized placeholders (`<SHEET_ID_...>`, "the company", "the lubricant brand", etc.) before publication to this public portfolio.
