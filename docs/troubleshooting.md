# Troubleshooting — Common errors and fixes

Reference guide for the 9 most common issues in the Pricing & Margin Engine v3.1. Follow symptom → probable cause → solution.

## Quick reference table

| # | Symptom | Probable cause | Quick fix |
|---|---|---|---|
| 1 | Column G (PROD COST) all empty | Wrong Sheet ID in CONFIG B5 | Confirm B5 points to `<SHEET_ID_EP_BASE_CUSTOS>` |
| 2 | Column F (UNIT) showing numbers (e.g., 342.29) | Old "PA CUSTO FINAL" tab reference | Already fixed (reads from RESULTADO). If recurs, check whether the cost sheet changed structure again |
| 3 | `#ERROR!` in export totalizers E4/F4 | SUBTOTAL+IFERROR + PT-BR locale | Already fixed (SUM direct) |
| 4 | B2 dropdown showing numbers instead of names | Old fallback reading column S | Click menu → Refresh customer list (B2) |
| 5 | `#REF!` in IMPORTRANGE | Unauthorized access | Type the formula in an empty cell, click "Allow access", delete |
| 6 | Salesperson not found, commission = 0 | Name mismatch between EP_CLIENTES and COMISSAO | Check exact spelling (accents, spaces). Lookup is by full name |
| 7 | Cost OK but margin still 0 | Missing combination in MARGEM_POLITICA | Verify the SKU's family/classification/strategic class exists in at least one of the 3 dimensions |
| 8 | Column E (Strategic Class) all empty | CLASSIF_ESTRATEGICA tab unpopulated | **Known pending** — only 1 test SKU populated currently |
| 9 | Freight = 0 for CIF customer | State name mismatch (full vs abbrev) | EP_CLIENTES uses full name ("MATO GROSSO"), FRETE tab uses abbreviation (MT). Script converts, but spelling must be exact |

---

## Detailed breakdown

### 1. Column G (PROD COST) all empty

**Symptom:** After running "Recalculate Margins", the PROD COST column is completely empty for every SKU, including SKUs you know have cost in the cost structure.

**Why it happens:** The CONFIG tab cell B5 holds the Sheet ID of the cost structure workbook. If the ID is wrong, mistyped, or the workbook was moved/duplicated, the script's `lerCustos_()` returns an empty map — silently. No error message is shown because the script handles "sheet not found" gracefully.

**Fix:**
1. Open the CONFIG tab.
2. Check cell B5 — it should be `<SHEET_ID_EP_BASE_CUSTOS>`.
3. If it's a placeholder (`<...>`), ask the system owner for the correct ID.
4. If it seems right but still fails, open that Sheet ID in a new browser tab to verify the workbook exists and has a RESULTADO tab with data in column H.

---

### 2. Column F (UNIT) showing numbers

**Symptom:** The UNIT column shows numeric values like `342.29` instead of text codes like `BD`, `TB`, `CX`.

**Why it happens:** In older versions of the code, this column read from a tab called "PA CUSTO FINAL" that has since been renamed to "RESULTADO". The script briefly pulled numeric cost values instead of unit codes.

**Fix:** Already fixed in v3.x — the script reads from RESULTADO col C. If this recurs, the cost structure workbook may have been restructured again. Investigate the RESULTADO tab's column C and confirm it still contains unit codes (BD/TB/CX/IBC).

---

### 3. `#ERROR!` in export totalizers E4/F4

**Symptom:** The exported customer-facing spreadsheet shows `#ERROR!` in the E4 (quantity total) and F4 (total value) cells.

**Why it happens:** Earlier versions used `SUBTOTAL(109, ...)` wrapped in `IFERROR(...)`. In Brazilian Portuguese locale, the comma/semicolon argument separator conflicted with the function parsing.

**Fix:** Already fixed — the engine now uses direct `=SUM(E6:E{lastRow})` and `=SUM(F6:F{lastRow})` formulas, which work consistently regardless of locale.

---

### 4. B2 dropdown showing numbers instead of customer names

**Symptom:** The B2 dropdown shows a list of numbers (looks like SKU codes or random IDs) instead of customer names.

**Why it happens:** The `onOpen()` automatic trigger has limited permissions and falls back to reading column S of the local sheet when it can't access EP_CLIENTES. If column S contains unrelated data, the dropdown is populated with garbage.

**Fix:** Click the menu **Pricing Engine → Refresh customer list (B2)**. This button uses full permissions to read EP_CLIENTES directly and rebuild the dropdown with real customer names + 4 standard tables.

---

### 5. `#REF!` in IMPORTRANGE

**Symptom:** A cell using `IMPORTRANGE(...)` shows `#REF!` with the tooltip "You need to connect these sheets".

**Why it happens:** Google Sheets requires explicit authorization the first time a workbook imports from another workbook. This is a one-time authorization per source.

**Fix:**
1. Click the cell showing `#REF!`.
2. A "Allow access" button appears in the popup.
3. Click "Allow access".
4. If the formula is not needed (since the engine uses SpreadsheetApp natively), delete the IMPORTRANGE formula entirely.

**Tip:** The engine **intentionally avoids** IMPORTRANGE in columns that the script writes to. IMPORTRANGE + ARRAYFORMULA conflicts with `setValues()`.

---

### 6. Salesperson not found, commission = 0

**Symptom:** For a specific customer, after running "Recalculate Margins", column I (COMMISSION %) shows 0% for all rows — unexpected because other customers' commissions work.

**Why it happens:** The commission lookup is **by full name, not by code**. The salesperson name in EP_CLIENTES (column VENDEDOR) must match exactly the name in EP_PARAMETROS_MARGEM → COMISSAO tab. Common mismatches:
- Trailing/leading spaces
- Accented characters: `SÉRGIO` vs `SERGIO`
- Casing: `joão silva` vs `João Silva`

**Fix:**
1. Open EP_CLIENTES and note the exact salesperson name (including accents).
2. Open EP_PARAMETROS_MARGEM → COMISSAO tab.
3. Compare character by character — any whitespace or accent difference blocks the match.
4. Fix the mismatch in whichever spreadsheet is incorrect.
5. Rerun "Recalculate Margins".

---

### 7. Cost OK but margin still 0

**Symptom:** Column G (PROD COST) is populated correctly, but columns M (MIN MARGIN %) and N (TARGET MARGIN %) remain 0%.

**Why it happens:** The 3D lookup `buscarMargem_(family, productClass, strategicClass)` did not find a matching rule in any of the three dimensions.

**Fix:**
1. Open EP_PARAMETROS_MARGEM → MARGEM_POLITICA tab.
2. For the problematic SKU, identify its FAMILY (col C of TABELA_NOVA), CLASSIFICATION (col D), and STRATEGIC CLASSIFICATION (col E).
3. Confirm that at least one combination exists in MARGEM_POLITICA:
   - Exact match on strategic class (strongest)
   - Match on product classification (medium)
   - Match on family (fallback)
4. Add a row to MARGEM_POLITICA covering this SKU if missing.
5. Rerun "Recalculate Margins".

---

### 8. Column E (Strategic Class) all empty

**Symptom:** Column E (CLASSIFICATION ESTRATEGICA) is empty for nearly all SKUs.

**Why it happens:** This is a **known pending state**. The CLASSIF_ESTRATEGICA tab in EP_PARAMETROS_MARGEM currently has only 1 test SKU populated. Full classification of the SKU portfolio into the four quadrants (PROFIT DRIVER / HIDDEN STAR / CASH COW / DEAD WEIGHT) is a pending business task.

**Status:** Not a bug. The 3D margin lookup gracefully falls back to product classification or family when strategic class is missing.

---

### 9. Freight = 0 for CIF customer

**Symptom:** A customer with Incoterm = CIF has column H (FREIGHT %) = 0, but should have a freight rate by state.

**Why it happens:** The FRETE tab in EP_PARAMETROS_MARGEM uses 2-letter state abbreviations (MT, SP, RJ), while EP_CLIENTES column ESTADO uses full state names ("MATO GROSSO", "SÃO PAULO"). The script converts full-name → abbreviation via an internal map. If the full name in EP_CLIENTES has a typo or unusual spelling, the conversion fails and freight returns 0.

**Fix:**
1. Open EP_CLIENTES for the problematic customer.
2. Check the ESTADO cell — should be the full state name exactly as Brazilian Portuguese official naming (e.g., "MATO GROSSO", not "MATOGROSSO" or "MT").
3. If ambiguous, copy-paste a state name from another customer row that works correctly.
4. Rerun "Recalculate Margins".

---

## Escalation

If the symptom doesn't match any of the above, open the Apps Script editor (Extensions → Apps Script) and check the execution log for stack traces. Most silent failures leave a trace. If the trace points to a function not in this guide, document the new pattern in this file.
