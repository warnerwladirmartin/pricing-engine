# Cost Integration — v3.0

How production costs flow from the BOM spreadsheet into the pricing engine.

---

## Source Spreadsheet

- **Name**: CostStructure_v2
- **Referred to in code as**: EP_BASE_CUSTOS (same Google Sheets ID)
- **Tab consumed by motor**: `RESULTADO`
- **Column consumed**: H — `custo_c_oluc` (production cost including OLUC levy)

The motor reads this tab via `lerCustos_()` and builds a `{ sku: cost }` map loaded once per `carregarCliente()` call.

---

## BOM Cascade — 8-Tab Structure

The cost flows through a cascade of formula-linked tabs. Changing a raw material price in CUSTOS_MP recalculates all downstream tabs automatically.

```
CUSTOS_MP
  721 raw materials
  col H = adjusted cost: F × (1 + G)
    │
    ▼
ESTRUTURA_PEs
  2,161 lines — composition of semi-finished goods (PE = produto em elaboracao)
  col F = cost lookup via XLOOKUP from CUSTOS_MP
  col G = component cost: E × F
    │
    ▼
PE_RESUMO
  ~400 PEs — one row per semi-finished good
  col B = total PE cost: SUMIFS(ESTRUTURA_PEs)
    │
    ▼
ESTRUTURA_PAs
  2,806 lines — composition of finished goods (PA = produto acabado)
  col G = cost lookup via XLOOKUP (from PE_RESUMO or CUSTOS_MP)
  col H = component cost: F × G
    │
    ▼
OLUC
  735 items — ANP mapping: which products owe the collection levy
  Flag: N = nao dispensado (with OLUC) | S = dispensado (without OLUC)
    │
    ▼
RESULTADO
  537 PAs with calculated costs
  col D = base production cost: SUMIFS(ESTRUTURA_PAs)
  col G = OLUC amount (12% of price, applied if flag = N)
  col H = custo_c_oluc = D + G  ← this is what the motor reads
```

---

## OLUC — Mandatory Used-Oil Collection Levy

OLUC (Obrigacao de Coleta de Oleo Usado/Contaminado) is a levy mandated by the ANP (Brazilian petroleum regulator).

- **Total rate**: 12% of the finished-goods price
- **Composition**: 11% lubricant oil component + 1% container/packaging
- **Determination**: based on the `ies_disp_coleta` field in the DAGDA ERP table `obr_produto_anp`
  - `N` = nao dispensado = OLUC applies (most lubricating oils)
  - `S` = dispensado = no OLUC (greases, 2-stroke oils, chain oils, extenders, etc.)

The OLUC tab contains the SKU-to-dispensation mapping. For OLUC-eligible products, column G of RESULTADO adds the 12% to the base production cost.

---

## Coverage: 537 with cost, 75 without

Of approximately 612 active finished goods (PAs):

| Group | Count | Status | Impact in motor |
|---|---|---|---|
| With BOM cost | 537 | RESULTADO col H populated | `CUSTO_PROD` filled; margins calculated normally |
| Group 1: historical cost, no BOM | ~34 | No ESTRUTURA_PAs entries | `CUSTO_PROD = 0`; ALERTA = SEM CUSTO |
| Group 2: base/bulk oils | ~17 | Direct cost input needed | `CUSTO_PROD = 0`; ALERTA = SEM CUSTO |
| Group 3: new products | ~24 | No history, no BOM | `CUSTO_PROD = 0`; ALERTA = SEM CUSTO |

The 75 without cost structure will show `SEM CUSTO` in column X of TABELA_NOVA. Prices can still be proposed manually (by editing PRECO_BASE or NOVO_PRECO directly), but margin calculations will be incorrect.

**Resolving Group 1**: query DAGDA table `obr_produto_anp` for BOM data or use historical cost as a temporary placeholder.

**Resolving Group 2**: energy base oils and bulk products have a direct purchase cost; these should be entered in CUSTOS_MP and linked in ESTRUTURA_PAs.

**Resolving Group 3**: new products need a BOM before pricing can be cost-based.

---

## Additional Cost Parameters (not in BOM)

The following variable costs are NOT part of the BOM cascade. They are read from EP_PARAMETROS_MARGEM and applied per customer:

| Cost | Source | Applied to |
|---|---|---|
| Freight | FRETE tab — by customer UF | All SKUs for that customer |
| Sales commission | COMISSAO tab — rep/CNPJ/UF | All SKUs for that customer |
| Tax burden | IMPOSTOS tab — by customer UF | All SKUs for that customer |
| Financial carrying cost | CUSTOS_ADICIONAIS tab — monthly rate | All SKUs; scales with payment term |

These are combined into `CUSTO_TOTAL%` (column K) and fed into the margin formula.

---

## CONFIG Reference

In EP_BASE_CUSTOS / CostStructure_v2:

| Tab | Key cell | Value |
|---|---|---|
| CONFIG | OLUC_RATE | 0.12 (12%) |
| CONFIG | DATA | snapshot date |
| CONFIG | VERSAO | v2 |

The motor does not read the CONFIG tab directly — it reads only the RESULTADO tab. OLUC is already baked into column H values.
