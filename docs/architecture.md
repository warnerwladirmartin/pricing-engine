# Architecture — Pricing & Margin Engine v3.0

## Data Flow

```
EP_PARAMETROS_MARGEM (6 tabs)
  ├── FRETE              by UF (%)
  ├── COMISSAO           rep default / CNPJ exception / UF rule
  ├── IMPOSTOS           total tax burden by UF (%)
  ├── MARGEM_POLITICA    family × classification × strategic quadrant → min% / target%
  ├── CUSTOS_ADICIONAIS  financial rate (monthly %), admin cost %
  └── CLASSIF_ESTRATEGICA  SKU → strategic quadrant mapping
        │
        │  lerParametros_()
        ▼
EP_BASE_CUSTOS / CostStructure_v2
  RESULTADO tab, column H = custo_c_oluc (537 PAs with BOM cost)
        │
        │  lerCustos_()
        ▼
EP_BASE_VENDAS
  sales history by customer CNPJ and SKU
  (DAGDA source post Dec/2025, HISTORICO source pre-migration)
        │
        │  (loaded inside carregarCliente)
        ▼
EP_CLIENTES
  customer master: NOME, CNPJ, ESTADO, REPRESENTANTE, PCT_FATURAMENTO,
                   TABELA_REF, PRAZO
        │
        │  atualizarDropdownB2_ → user selects in CONFIG!B2
        │  carregarCliente()
        ▼
EP_TABELAS_REF
  tabs: PADRAO, RJ, LEME, CONSUMO, VAREJO
  (customers with TABELA_REF=PADRAO use LEME as source — see note below)
        │
        │  (loaded inside carregarCliente)
        ▼
EP_MOTOR_PRECIFICACAO
  ├── CONFIG             IDs, DATA_CORTE_HISTORICO, DATA_VIGENCIA, LOGO_URL
  ├── REAJUSTES          global + classification + family adjustment %s
  ├── CLASSIFICACAO      SKU → commercial classification lookup
  ├── TABELA_NOVA        28-column working table (see below)
  └── Export
        │
        │  exportarTabelaAtual() / exportarTodasTabelas()
        ▼
Google Drive folder
  └── [Customer/Table name] — fully formatted Sheets file per customer
```

> **LEME note**: Customers with `TABELA_REF = PADRAO` in EP_CLIENTES are mapped to
> the LEME tab in EP_TABELAS_REF (not the PADRAO tab). LEME reflects the current
> commercial benchmark and is more appropriate as a base for standard-table customers.

---

## TABELA_NOVA — Column Map

28 columns, 1-indexed. The `COL` object in `main.gs` maps names to indices.

| Range | Purpose |
|---|---|
| A–E (1–5) | Product identity: SKU, name, family, classification, unit |
| F–K (6–11) | Cost breakdown: production + freight + commission + taxes + financial + total |
| L–O (12–15) | Margin policy: min%, target%, floor price, suggested price |
| P–R (16–18) | Reference data: benchmark price, last invoiced price, history flag |
| S–U (19–21) | Pricing decision: base price, adjustment %, new price |
| V–X (22–24) | Margin check: real margin %, real margin R$, alert |
| Y–AB (25–28) | Sales intelligence: last sale date, volume, invoice count, source flag |

---

## Module Map

### [B] Lifecycle Hooks

| Function | Trigger | Description |
|---|---|---|
| `onOpen()` | Spreadsheet open | Creates "EP Precificacao" menu; refreshes B2 dropdown |
| `onEdit(e)` | Cell edit | Detects B2 change in CONFIG tab; fires `carregarCliente()` |
| `atualizarDropdownB2_()` | Manual / onOpen | Reads EP_CLIENTES for customer list; prepends 4 standard tables; falls back to CONFIG col S |

### [C] Utility Functions

| Function | Description |
|---|---|
| `parsePreco_(val)` | Converts PT-BR price strings ("R$ 1.234,56") to JS float |
| `removerAcentos_(str)` | NFD normalization for accent-insensitive comparisons |
| `lerClassificacoes_()` | Reads CLASSIFICACAO tab → `{ sku: classification }` map |
| `lerUnidades_(wsCfg)` | Reads EP_BASE_CUSTOS first sheet → `{ sku: unit }` map |
| `calcPrazoMedio_(historico)` | Weighted average payment term from sales history (used for financial cost estimation) |

### [D] Data Loaders

| Function | Source | Returns |
|---|---|---|
| `lerCustos_()` | EP_BASE_CUSTOS / RESULTADO tab, col H | `{ sku: custo_c_oluc }` |
| `lerParametros_()` | EP_PARAMETROS_MARGEM (all 6 tabs) | Structured params object |
| `buscarMargem_(sku, familia, classif, params)` | params.margemPolitica + classifEstrategica | `{ min: %, alvo: % }` |
| `calcComissao_(cnpj, uf, rep, params)` | params.comissao/Excecoes/UF | commission % (number) |
| `lerReajustes_(wsRea)` | REAJUSTES tab | `{ global, classif: {}, familia: {} }` |

### [E] Main Workflows

| Function | Description |
|---|---|
| `carregarCliente()` | Main orchestrator: reads all sources, fills all 28 columns for the selected customer |
| `carregarTabelaPadrao_(nome)` | Loads a standard table (PADRAO GERAL / RJ / CONSUMO / VAREJO) from EP_TABELAS_REF |
| `corrigirFormulas()` | "Recalcular Margens": recomputes cost/margin/alert columns without reloading history or prices |

### [F] Export

| Function | Description |
|---|---|
| `exportarTabelaAtual()` | Exports the currently loaded table to Drive (single customer) |
| `exportarTodasTabelas()` | Batch export: 4 standard tables + all EP_CLIENTES customers |
| `montarItens_(wsNova)` | Builds item array for wholesale layout (1 price column) |
| `montarItensVarejo_(wsNova)` | Builds item array for VAREJO layout (4 price tier columns) |
| `criarArquivoTabela_(...)` | Creates formatted Sheets file with brand styling and logo |

---

## EP_MOTOR Tabs

| Tab | Purpose | Edited by |
|---|---|---|
| CONFIG | Sheet IDs (B3–B8), dates, customer selector (B2) | Analyst |
| REAJUSTES | Adjustment %s: global, by classification, by family | Analyst |
| CLASSIFICACAO | SKU → commercial classification mapping | Analyst |
| TABELA_NOVA | 28-column working table | Script (auto) |
| INSTRUCOES | User documentation (new in v3.0) | Analyst |
| MODELO_REAJUSTES | Template for filling REAJUSTES tab (new in v3.0) | Reference |

---

## Fallback Mechanisms

| Situation | Fallback |
|---|---|
| EP_CLIENTES unavailable | Reads customer list from CONFIG col S |
| EP_BASE_CUSTOS unavailable | `lerCustos_()` returns empty map; CUSTO_PROD = 0; alert = SEM CUSTO |
| EP_PARAMETROS_MARGEM unavailable | Default margins: 15% min / 25% target; 5% commission |
| EP_BASE_VENDAS unavailable | TEM_HIST = FALSE; PRECO_BASE = PRECO_REF |
| Logo URL unreachable | Logo embed skipped (logged), file still created |
| CNPJ not in EP_CLIENTES | Script alerts user and aborts load |
