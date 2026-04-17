# Pricing & Margin Engine — Complete Manual v3.1

> Full operational manual for the Pricing & Margin Engine.
> Originally written 2026-04-17 by the internal analyst who owns the pricing domain.
> Sanitized for portfolio publication.
>
> For source code, see: [src/main.gs](../src/main.gs)
> For architecture summary, see: [architecture.md](architecture.md)
> For troubleshooting guide, see: [troubleshooting.md](troubleshooting.md)
> For daily workflows, see: [daily-workflow.md](daily-workflow.md)

**Version:** v3.1 (Motor with Margin + Strategic Classification)
**Last update:** 2026-04-17

---

## Summary

1. Project overview
2. Architecture — how the spreadsheets talk
3. Main spreadsheet: EP_MOTOR_PRECIFICACAO
4. TABELA_NOVA tab — column by column (29 columns)
5. CONFIG tab — general parameters
6. REAJUSTES tab — adjustment policy
7. CLASSIFICACAO tab — product nature
8. Supporting spreadsheets (linked)
9. "Pricing Engine" menu — what each button does
10. Calculation memory — formulas explained
11. Day-to-day flow (step-by-step)
12. Key architectural decisions
13. Troubleshooting — common errors and fixes
14. Glossary

---

## 1. Project Overview

### What it is

The **Pricing Engine** is a Google Sheets + Apps Script solution that centralizes the entire pricing policy of the lubricant brand. It replaces the old manual routine of "open Excel, copy table, apply adjustment with a formula, export" with a single, auditable system that **knows the real cost** of every product.

### Why it exists

Previously, the price table was a static spreadsheet the sales team adjusted by "feeling" or by a flat-percentage uplift. Nobody knew, at the moment of accepting an order, whether the agreed price covered margin or generated **loss**. The engine was built to answer three questions with precision:

1. What is the **minimum price** (minimum acceptable margin) at which this product can be sold to this customer?
2. What is the **suggested price** (target margin per commercial policy) for this product/customer?
3. Is the current price table within policy? If not, which SKUs are below the minimum?

### Two universes, one engine

The engine operates on two fronts:

- **External front (customer):** generates the formatted price tables that go to end customers. Clean price, no internal data, no visible margin.
- **Internal front (commercial/executive):** calculates real margin, triggers alerts, feeds the Margin Monitoring dashboard and the Order Simulator (both separate projects, in construction).

**Non-negotiable principle:** margin data **never** leaves for the customer. The exported spreadsheet shows only full price + family + unit.

---

## 2. Architecture — how the spreadsheets talk

```
┌─────────────────────────────────────────────────────────────────┐
│                  EP_MOTOR_PRECIFICACAO (the brain)              │
│   TABELA_NOVA tab — where the magic happens (29 columns)        │
│   CONFIG tab      — parameters and linked spreadsheet IDs       │
│   REAJUSTES tab   — adjustment % per level                      │
│   CLASSIFICACAO tab — SKU → MINERAL / SYNTHETIC / etc.          │
└───────┬──────────┬──────────┬──────────┬──────────┬─────────────┘
        │          │          │          │          │
        ▼          ▼          ▼          ▼          ▼
   ┌────────┐ ┌────────┐ ┌────────┐ ┌────────┐ ┌────────────────┐
   │EP_PARAM│ │EP_TAB  │ │EP_CLI  │ │EP_BASE │ │CostStructure   │
   │_MARGEM │ │ELAS_REF│ │ENTES   │ │_VENDAS │ │_v2             │
   │        │ │        │ │        │ │        │ │                │
   │Freight,│ │Standard│ │Customer│ │Sales   │ │Production cost │
   │comm.,  │ │prices  │ │master  │ │history │ │per PA          │
   │tax,    │ │per SKU │ │(state, │ │(SAP +  │ │(MP→PE→PA→OLUC) │
   │margin, │ │        │ │rep,    │ │DAGDA)  │ │                │
   │strateg.│ │        │ │incoterm│ │        │ │                │
   │        │ │        │ │term)   │ │        │ │                │
   └────────┘ └────────┘ └────────┘ └────────┘ └────────────────┘
```

**How the spreadsheets communicate:** the Apps Script reads all of them via `SpreadsheetApp.openById()`. IDs are stored in the `CONFIG` tab. We do NOT use `IMPORTRANGE` in columns that the script writes to — it conflicts with `setValues()`. The script reads and writes directly.

---

## 3. Main spreadsheet: EP_MOTOR_PRECIFICACAO

**ID:** `<SHEET_ID_EP_MOTOR>`

This is where you work 95% of the time. It has 4 operational tabs:

| Tab | Purpose |
|---|---|
| **TABELA_NOVA** | The working table. Each row is a SKU. 29 columns show everything: cost, minimum margin, suggested price, new price, alert. This is where you analyze and decide. |
| **CONFIG** | General parameters: IDs of linked spreadsheets, validity date, history cutoff date, logo URL. |
| **REAJUSTES** | Price adjustment policy in 4 levels: global, by classification, by family, by table/customer. Accumulate multiplicatively. |
| **CLASSIFICACAO** | Simple SKU → classification mapping (MINERAL, SYNTHETIC, SEMI-SYNTHETIC, GREASE, etc.). 458 items. |

---

## 4. TABELA_NOVA tab — column by column

29 columns. Each one explained below: **source of data**, **meaning**, and whether you can **edit manually** or if it is **calculated/repopulated by the script**.

### Block 1 — Product identification (A–F)

| Col | Header | Source | Edit? | What it is |
|---|---|---|---|---|
| A | SKU | Fixed registry | No | Product code (e.g., PA00400). Always hidden on export. |
| B | PRODUCT | Fixed registry | No | Description (e.g., "HYDRA 68 HL BUCKET 20L"). |
| C | FAMILY | Fixed registry | No | Commercial family (e.g., HYDRAULICS, MOTORS, TRACTORS). Used for adjustment and margin by family. |
| D | CLASSIFICATION | CLASSIFICACAO tab | No | MINERAL, SYNTHETIC, SEMI-SYNTHETIC, GREASE, AQUEOUS PRODUCT, MINERAL BASE, SYNTHETIC BASE. |
| **E** | **STRATEGIC CLASSIFICATION** ⭐ v3.1 | EP_PARAMETROS_MARGEM, CLASSIF_ESTRATEGICA tab | No | PROFIT DRIVER / HIDDEN STAR / CASH COW / DEAD WEIGHT. Margin × volume quadrant. Currently only 1 test SKU — pending population. |
| F | UNIT | CostStructure_v2, RESULTADO tab, col C | No | Commercial unit: BD (bucket), TB (drum), CX (case). Repopulated every time you run "Recalculate Margins". |

### Block 2 — Cost and margin (G–N)

This is the heart of the engine. Everything here is calculated by the script.

| Col | Header | Formula / Origin | Consideration |
|---|---|---|---|
| G | PROD COST | Read from CostStructure_v2 (RESULTADO tab, col H) | Total production cost: raw material + packaging + OLUC (if applicable). "Factory gate" cost. |
| H | FREIGHT % | EP_PARAMETROS_MARGEM, FRETE tab, lookup by customer state | Zero if Incoterm = FOB (customer pays freight). If CIF, applies the state's %. |
| I | COMMISSION % | EP_PARAMETROS_MARGEM, COMISSAO tab, lookup by salesperson full name | Lookup by **full name** (not code). Has 3 sections: default per rep, exceptions by customer, special rules (e.g., ALL_SP adds +0.3% for a specific rep). |
| J | TAXES % | EP_PARAMETROS_MARGEM, IMPOSTOS tab, lookup by state | ICMS + PIS/COFINS per state. Single column with the sum. |
| K | FIN. COST % | Calculated: monthly_rate × average_term_in_months | Zero if term = "CASH". Example: term "30/60/90" → average 60 days = 2 months × 2% = 4%. |
| L | TOTAL COST % | = H + I + J + K | Sum of all variable costs that hit the **sale price** (not the production cost). |
| M | MIN MARGIN % | EP_PARAMETROS_MARGEM, MARGEM_POLITICA tab, 3D lookup | Minimum acceptable margin. Lookup priority: **Strategic Class > Product Class > Family**. |
| N | TARGET MARGIN % | EP_PARAMETROS_MARGEM, MARGEM_POLITICA tab, same logic | Target margin (commercial policy). |

### Block 3 — Calculated prices (O–P)

| Col | Header | Formula | Meaning |
|---|---|---|---|
| O | MIN PRICE | = G / (1 - L - M) | Minimum price that covers cost + variable costs + minimum margin. Selling below is a policy loss. |
| P | SUGGESTED PRICE | = G / (1 - L - N) | Price that delivers the target margin. The "price we'd like to sell at". |

**Why divide instead of adding?** Because commission, freight, tax and margin are calculated **on the sale price**, not on the cost. The correct formula isolates the price: if 30% of the price disappears in variable costs and 20% needs to become margin, then 50% remains to cover production cost → price = cost / 0.50.

### Block 4 — History and reference (Q–S)

| Col | Header | Source | Meaning |
|---|---|---|---|
| Q | REF PRICE | EP_TABELAS_REF (PADRAO/RJ/CONSUMO/VAREJO/LEME tab per customer's TABELA_REF) | Current standard table price for the product. PADRAO customers use the **LEME** tab (business decision). |
| R | LAST PRICE | EP_BASE_VENDAS, latest invoice of this SKU for this customer | Price actually charged in the last sale. SAP = full price. DAGDA = reduced price (divided by billing %). |
| S | HAS HIST | YES/NO | If YES, the adjustment base price comes from history (R); if NO, from the reference table (Q). |

### Block 5 — Adjustment and new price (T–V)

| Col | Header | Formula | Meaning |
|---|---|---|---|
| T | BASE PRICE | =IF(S="YES"; R; Q) | Adjustment starting price. |
| U | ADJUSTMENT % | Accumulated from REAJUSTES tab: (1+global) × (1+classif) × (1+family) × (1+table) - 1. **Multiplicative, not additive.** | Policy-driven price uplift. |
| V | NEW PRICE | = T × (1 + U) | Final adjusted price that goes to the exported table. |

### Block 6 — Real margin and alert (W–Y)

This block validates whether the NEW PRICE makes sense.

| Col | Header | Formula | Meaning |
|---|---|---|---|
| W | REAL MARGIN % | = (V - G) / V - L | Margin remaining after paying production cost and all variable costs. |
| X | REAL MARGIN R$ | = V × W | Margin in currency per unit. |
| Y | ALERT | Compare V vs O and P | **BELOW MIN** (V < O, red), **BELOW TARGET** (O ≤ V < P, yellow), **OK** (P ≤ V ≤ P × 1.15, green), **PREMIUM** (V > P × 1.15, blue). |

**⚠️ v3.1 change:** PREMIUM threshold is now `1.15×` (was `1.20×` in v3.0). OK band is now tighter.

### Block 7 — Sales statistics (Z–AC)

Helps decision-making for the SKU.

| Col | Header | Source | Meaning |
|---|---|---|---|
| Z | LAST SALE DT | EP_BASE_VENDAS | When this customer last bought this SKU. |
| AA | HIST QTY | EP_BASE_VENDAS | Total quantity purchased in history. |
| AB | N INVOICES | EP_BASE_VENDAS | Number of invoices with this SKU. |
| AC | SOURCE | SAP / DAGDA | Which system the history came from (affects price reading in column R). |

---

## 5. CONFIG tab — general parameters

| Cell | Parameter | Example | Purpose |
|---|---|---|---|
| B2 | Selected customer/table | (dropdown) | Defines which customer/table to load. Fed by EP_CLIENTES + 4 standard tables. |
| B3 | (free) | — | — |
| B4 | EP_BASE_VENDAS ID | `<SHEET_ID_...>` | Sales history. |
| B5 | CostStructure ID | `<SHEET_ID_EP_BASE_CUSTOS>` | **Critical:** if the ID is wrong, cost comes back empty without any error. |
| B6 | EP_CLIENTES ID | `<SHEET_ID_...>` | Customer registry. |
| B8 | LOGO_URL | `https://drive.google.com/uc?...` | Brand logo embedded in exported tables. |
| B9 | HISTORY_CUTOFF_DATE | 2026-02-28 | Sales before this date are ignored in history. |
| B10 | VALIDITY_DATE | 2026-03-26 | Date shown in the exported table header. |
| B11 | EP_PARAMETROS_MARGEM ID | `<SHEET_ID_EP_PARAMETROS_MARGEM>` | Where all margin rules live. |
| B12 | MONTHLY_FIN_RATE | 0.02 (2%) | Cost of capital per month. Multiplied by the customer's average term. |

---

## 6. REAJUSTES tab — adjustment policy

The tab has **4 sections** (4 adjustment levels). All accumulate (they are multiplicative), giving flexibility to handle exceptions without touching other levels.

| Range | Level | Example |
|---|---|---|
| B5 | Global | 8% — applies to everything |
| A9:B15 | By Classification | MINERAL +2%, SYNTHETIC 0%, GREASE -1% |
| A19:B43 | By Family | HYDRAULICS +3%, MOTORS 0%, TRACTORS +5% |
| A47:B200 | By Table/Customer | LEME -2%, RJ +1%, customer XYZ +0.5% |

**How it accumulates:** a MINERAL SKU of the HYDRAULICS family sold to LEME with Global 8%, Classif 2%, Family 3%, Table -2%:

```
Final adjustment = (1 + 0.08) × (1 + 0.02) × (1 + 0.03) × (1 - 0.02) - 1
                 = 1.08 × 1.02 × 1.03 × 0.98 - 1
                 = 0.1116 ≈ 11.16%
```

**Why multiplicative, not additive?** Because each level is an independent decision. Summing would give 11% (8+2+3-2). Multiplying gives 11.16% — capturing the correct composite effect, like compound interest.

---

## 7. CLASSIFICACAO tab — product nature

Simple SKU → classification map. 458 items imported from an internal sales-lead spreadsheet.

| SKU | Classification |
|---|---|
| PA00001 | MINERAL |
| PA00100 | SYNTHETIC |
| PA00250 | GREASE |
| ... | ... |

**What it's for:**

1. Feeds column D of TABELA_NOVA.
2. Feeds the "Classification" level of the adjustment.
3. Feeds the min/target margin lookup (columns M and N).

---

## 8. Supporting spreadsheets (linked)

### 8.1 EP_PARAMETROS_MARGEM

**ID:** `<SHEET_ID_EP_PARAMETROS_MARGEM>`

**Function:** the engine's "rulebook". Everything that is **not customer-specific** lives here.

| Tab | Contents | Owner |
|---|---|---|
| FRETE | % per state (all 27 states) | Logistics |
| COMISSAO | 26 actual salespeople + exceptions per customer + special rules (ALL_SP) | Commercial |
| IMPOSTOS | ICMS + PIS/COFINS per state | Accounting |
| CUSTOS_ADICIONAIS | OLUC $0.11/L, Logistics $0.01/L, Admin 3%, Financial cost 2%/mo | Finance |
| MARGEM_POLITICA | Min/target margin in 3 dimensions: Family × Classification × Strategic Classification | Executive/Commercial |
| **CLASSIF_ESTRATEGICA** ⭐ v3.1 | SKU → Quadrant (Profit Driver / Hidden Star / Cash Cow / Dead Weight) | Pending population |

**Why separate from the customer?** Allows calculating margin for **any** customer, even those not in the main registry. Rules apply per state/rep/product, not per individual customer.

### 8.2 EP_TABELAS_REF

**ID:** `<SHEET_ID_EP_TABELAS_REF>`

**Function:** reference prices per standard table.

5 tabs: PADRAO, RJ, LEME, CONSUMO, VAREJO. Each lists full price per SKU.

**Important detail:** customers with `TABELA_REF = "PADRAO"` use the **LEME** tab, not PADRAO. This is a business decision — the LEME table is the actual operational reference.

### 8.3 EP_CLIENTES

**Function:** lean customer registry.

| Column | Contents |
|---|---|
| CNPJ | Tax ID |
| RAZAO_SOCIAL | Legal name |
| ESTADO | Full state name ("MATO GROSSO") — script converts to 2-letter abbrev |
| INCOTERM | CIF or FOB (affects freight) |
| VENDEDOR | Salesperson full name (direct lookup into COMISSAO) |
| PRAZO | "30/60/90" or "CASH" (affects financial cost) |
| TABELA_REF | Which standard table to use |
| PCT_FATURAMENTO | % of full price the customer actually pays |

**Legacy fields** (commission, ICMS, PIS/COFINS) from this spreadsheet are no longer read — they come from EP_PARAMETROS_MARGEM now.

### 8.4 EP_BASE_VENDAS

**Function:** sales history to feed columns R, Z, AA, AB, AC of TABELA_NOVA.

Two source systems:
- **SAP:** full price (read directly)
- **DAGDA:** reduced price (the script divides by the billing % to normalize)

### 8.5 CostStructure_v2

**ID:** `<SHEET_ID_EP_BASE_CUSTOS>`

**Function:** calculates the real production cost of every PA (Finished Product).

8 tabs in cascade:

```
CUSTOS_MP (721 MPs with manual adjustment)
    ↓
ESTRUTURA_PEs (2,161 lines — bulk = Σ MP × qty)
    ↓
PE_RESUMO (400 PEs — total cost per semi-finished)
    ↓
ESTRUTURA_PAs (2,806 lines — PA = bulk + packaging + auxiliaries)
    ↓
RESULTADO (537 PAs — final cost + OLUC if applicable) ← engine reads from here
```

**OLUC** (Used Oil Collection Obligation): 12% on the price (11% oil + 1% containers). Applies only to products with `ies_disp_coleta = 'N'` in the ERP. Exempt: greases, 2-stroke, chainsaw, agricultural.

**Pending:** 604 PAs without BOM — awaiting meeting with the BOM coordinator.

---

## 9. "Pricing Engine" menu — what each button does

When you open the spreadsheet, the **Pricing Engine** menu appears in the bar. Each button has a clear function:

### 🟦 Load Customer / Table

**When to use:** when you change the customer/table in B2 and want to populate TABELA_NOVA with that data.

**Under the hood:**
1. Reads the customer selected in B2.
2. Looks up in EP_CLIENTES: state, salesperson, incoterm, term, reference table, billing %.
3. Loads SKUs with the reference price (Q) from the appropriate table.
4. Loads sales history (R, Z, AA, AB, AC) from EP_BASE_VENDAS.
5. **Does not calculate margin yet** — only populates data.

### 🟧 Clear Data

**When to use:** before loading a new customer, or to start from scratch.

**What it does:** clears the data rows of TABELA_NOVA (keeps header and formatting).

### 🟩 Recalculate Margins

**When to use:** **always** after loading the customer, or when any parameter changes (adjustment, margin policy, MP cost).

**Under the hood (the most important part):**
1. Reads updated costs from CostStructure_v2 (RESULTADO col H) → fills G.
2. Repopulates UNIT (F) via `unidMap[sku]` from the RESULTADO tab col C.
3. Reads EP_PARAMETROS_MARGEM and calculates H, I, J, K → sum into L.
4. Runs 3D lookup on MARGEM_POLITICA → fills M and N.
5. Calculates O and P (minimum and suggested price).
6. Looks up strategic classification (E).
7. Applies accumulated adjustment (U) → calculates NEW PRICE (V).
8. Calculates real margin (W, X) and fires alert (Y).

### 🟨 Export current table (Drive)

**When to use:** when you finish working on a table and want to generate the customer file.

**What it does:**
1. Creates a new Google Sheets in Drive.
2. Applies styling: logo, header ("Price Table — Lubricant Brand"), validity date, zebra, borders, family separators.
3. Includes only customer-facing columns: SKU (hidden), Product, Family, Unit, **Invoice Price** (full, no margin).
4. Adds Quantity (empty, initialized to 0) and Total Value (`=D × E`) columns for the salesperson to fill at order time.
5. Total row with `=SUM()` directly.

**VAREJO export is different:** 4 price columns (Retail, Over 10 Vol, Over 20 Vol, Over 50 Vol) + Family. No Quantity/Total.

### 🟨 Export ALL tables (Drive)

**When to use:** after a general adjustment, to generate all tables at once.

**What it does:** runs the export in a loop for:
- 4 standard tables (PADRAO GERAL, PADRAO RJ, PADRAO CONSUMO, PADRAO VAREJO)
- All customers registered in EP_CLIENTES

### 🟪 Refresh customer list (B2)

**When to use:** when you registered a new customer in EP_CLIENTES and it doesn't appear in the B2 dropdown.

**What it does:** re-reads EP_CLIENTES and regenerates the B2 data validation with the updated list (customers + 4 standard tables).

**Why a separate button?** The automatic `onOpen()` has limited permissions and uses a fallback. This button uses full permissions to read EP_CLIENTES directly.

---

## 10. Calculation memory — formulas explained

### 10.1 Total Cost %

```
TOTAL COST % = FREIGHT % + COMMISSION % + TAXES % + FINANCIAL COST %
```

All calculated **on the sale price**, not on the production cost. It is the share of the price that "evaporates" before becoming margin.

### 10.2 Minimum Price and Suggested Price

```
MIN PRICE       = PROD COST / (1 - TOTAL COST % - MIN MARGIN %)
SUGGESTED PRICE = PROD COST / (1 - TOTAL COST % - TARGET MARGIN %)
```

**Step-by-step logic (numeric example):**

Imagine: Production cost = R$ 100, Total Cost = 30% (freight 5 + commission 5 + taxes 18 + financial 2), Minimum margin = 15%.

- Of the sale price, 30% disappears in variable costs and 15% needs to become minimum margin.
- What remains: `100% - 30% - 15% = 55%` to cover the production cost.
- So: `55% × price = R$ 100 → price = R$ 100 / 0.55 = R$ 181.82`.
- Check: `price 181.82 × 30% (costs) = 54.55`. `price - cost - variable costs = 181.82 - 100 - 54.55 = 27.27 = 15% of price` ✅

### 10.3 Accumulated adjustment

```
ADJUSTMENT % = (1 + GLOBAL) × (1 + CLASSIF) × (1 + FAMILY) × (1 + TABLE) - 1
```

Multiplicative. Each level is an independent decision that composes.

### 10.4 Financial Cost

```
FIN COST % = MONTHLY_FIN_RATE × AVERAGE_TERM_MONTHS
```

Where:
- `MONTHLY_FIN_RATE` = CONFIG B12 (today 2%)
- `AVERAGE_TERM_MONTHS` = arithmetic mean of customer's term / 30
- Customer "30/60/90" → (30+60+90)/3 = 60 days = 2 months → 2 × 2% = 4%
- Customer "CASH" → 0%

### 10.5 Real Margin

```
REAL MARGIN % = (NEW PRICE - PROD COST) / NEW PRICE - TOTAL COST %
REAL MARGIN R$ = NEW PRICE × REAL MARGIN %
```

Subtracting TOTAL COST % at the end is what separates **gross margin** (only covers production cost) from **real margin** (also covers variable costs on revenue).

### 10.6 Alert (column Y)

| Condition | Alert | Color |
|---|---|---|
| V < O | BELOW MIN | Red |
| O ≤ V < P | BELOW TARGET | Yellow |
| P ≤ V ≤ P × 1.15 | OK | Green |
| V > P × 1.15 | PREMIUM | Blue |

### 10.7 3D Lookup of Margin Policy ⭐ v3.1

The function `buscarMargem_(family, productClass, strategicClass)` searches in order:

1. **Most specific:** strategicClass (e.g., PROFIT DRIVER) → if found, uses this.
2. **Medium:** productClass (e.g., MINERAL) → if found, uses this.
3. **Most general:** family (e.g., HYDRAULICS) → final fallback.

**Open question (business decision):** today this priority is "override" — the most specific level wins completely. Alternative under study: turning Strategic Classification into a **delta** (e.g., PROFIT DRIVER = base + 5pp). Business decision; small code impact (lines 319–345 of `buscarMargem_`).

---

## 11. Day-to-day flow (step-by-step)

See [daily-workflow.md](daily-workflow.md) for the full 5 scenarios (A–E).

---

## 12. Key architectural decisions

These are the "structural" decisions of the project. Knowing them prevents someone from accidentally breaking the system.

### D1 — The exported table goes to the customer, without internal data

Margin, alert, cost: **never** appear in the exported table. The commercial team may see it in TABELA_NOVA (internal), but the final file is a clean price only.

**Why:** commercial confidentiality. If margin leaks, the customer understands the markup and renegotiates.

**Consequence:** Margin Monitoring and Order Simulator became **separate internal projects**.

### D2 — EP_PARAMETROS_MARGEM is the single source of rules

Commission, freight, tax, minimum margin: **everything** comes from here, never from the customer registry. The customer registry is lean (state, salesperson, incoterm, term).

**Why:** allows analyzing the margin of **any** customer, not just the main ones. Scales to "view the whole company's margin".

### D3 — The script reads costs via SpreadsheetApp, not IMPORTRANGE

Column G (cost) does not have an `IMPORTRANGE` formula. The script reads directly via `SpreadsheetApp.openById()` and writes with `setValues()`.

**Why:** IMPORTRANGE + ARRAYFORMULA conflicts with the script's `setValues()` ("the array result was not expanded because it would overwrite data in F7"). Reading via script is more robust.

### D4 — UNIT is repopulated each run, not preserved

Column F (UNIT) is rewritten on every "Recalculate Margins" via `unidMap[sku]` read from RESULTADO col C.

**Why:** wrong numeric values inherited from old imports would "stick" forever if the script only preserved the existing value.

### D5 — Adjustment accumulates multiplicatively

The 4 adjustment levels compose by multiplication, not by sum.

**Why:** captures the correct composite effect. Each level is an independent decision.

### D6 — Margin Policy has a 3D lookup with specificity priority

**Strategic Class > Product Class > Family.** The most specific level wins (today, by override).

**Under review:** see open question in section 10.7.

---

## 13. Troubleshooting — common errors and fixes

See [troubleshooting.md](troubleshooting.md) for the full reference.

---

## 14. Glossary

| Acronym | Meaning |
|---|---|
| PA | Finished Product (e.g., PA00400 = HYDRA 68 HL BUCKET 20L) |
| PE | Product in Elaboration (bulk, semi-finished) |
| MP | Raw Material (base oil, additive, packaging) |
| BOM | Bill of Materials — component list of a product |
| OLUC | Used Oil Collection Obligation (ANP regulatory fee, 12%) |
| DAGDA | The company's legacy ERP (Progress/OpenEdge) |
| SAP | New system, current source of sales history |
| SKU | Product code |
| NF | Invoice — full invoiced price |
| PCT_FATURAMENTO | % of full price the customer actually pays |
| CIF | Freight paid by the manufacturer (enters the calculation) |
| FOB | Freight paid by the customer (freight = 0 in the calculation) |
| Average term | Arithmetic mean of installment terms ("30/60/90" → 60 days) |
| BD / TB / CX / IBC | Bucket 20L / Drum 60-200L / Case / Container 1000L |
| Strategic quadrant | Profit Driver / Hidden Star / Cash Cow / Dead Weight (margin × volume) |

---

**End of manual.**

In case of doubt about a calculation, start with section 10. About operation, section 11. About architecture/decisions, section 12.
