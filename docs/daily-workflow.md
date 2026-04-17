# Daily Workflow — 5 operational scenarios

Reference guide for the most common day-to-day tasks in the Pricing & Margin Engine v3.1.

---

## Scenario A — Load and analyze margin for a specific customer

**When to use:** you want to review a particular customer's profitability, or prepare a quote for them.

**Steps:**

1. Open EP_MOTOR_PRECIFICACAO.
2. In cell B2, pick the customer from the dropdown.
   - If the customer does not appear, click **Pricing Engine → Refresh customer list (B2)**.
3. Click **Pricing Engine → Load Customer / Table**. TABELA_NOVA fills with SKUs, reference prices, and sales history.
4. Click **Pricing Engine → Recalculate Margins**. Columns G–Y are populated/recalculated (cost, variable costs, margins, alerts).
5. Filter column Y (Alert) by `BELOW MIN` → see where there is a policy loss.
6. Filter column Y by `BELOW TARGET` → see where there is an opportunity to raise the price.
7. Decide the action: renegotiate, maintain, or change the customer's reference table.

**Expected outcome:** clear visibility of which SKUs are under-priced vs. the policy for this customer.

---

## Scenario B — Apply a table-wide price adjustment

**When to use:** company-wide adjustment (e.g., supplier cost went up, inflationary uplift, strategic repricing).

**Steps:**

1. Decide the percentages per level (global, classification, family, table).
2. Open the REAJUSTES tab and fill the 4 levels. Remember: they **accumulate multiplicatively**.
3. Go back to TABELA_NOVA. Select the customer (or standard table) in B2.
4. **Load Customer / Table** → **Recalculate Margins**.
5. Review column V (NEW PRICE) and Y (Alert). If anything ended up `BELOW MIN`, revisit — the adjustment may have been insufficient for that SKU.
6. **Export current table (Drive)** to generate the customer file.
7. Repeat for each customer — or use **Export ALL tables** if the adjustment is general.

**Expected outcome:** a consistent price increase applied across the portfolio, with visibility into any SKUs that still violate policy.

**Tip:** Always eyeball the "NEW PRICE / OLD PRICE" ratio for a handful of top-selling SKUs before exporting. Large deviations may indicate wrong input percentages.

---

## Scenario C — Register a new customer

**When to use:** onboarding a new customer who will receive a price table.

**Steps:**

1. Open EP_CLIENTES and add a row with:
   - CNPJ
   - Legal name (RAZAO_SOCIAL)
   - State (full name, e.g., "MATO GROSSO" — not "MT")
   - Salesperson (full name, must match EP_PARAMETROS_MARGEM → COMISSAO exactly)
   - Incoterm (CIF or FOB)
   - Term (e.g., "30/60/90" or "CASH")
   - Reference table (PADRAO, RJ, CONSUMO, VAREJO)
   - PCT_FATURAMENTO (% of full price the customer actually pays)
2. Go back to EP_MOTOR_PRECIFICACAO.
3. Click **Pricing Engine → Refresh customer list (B2)**.
4. Select the new customer in B2 and follow Scenario A.

**Expected outcome:** the new customer appears in the B2 dropdown and can be loaded, recalculated, and exported.

**⚠️ Common issue:** if commission comes back 0% for this customer (even after recalc), check the salesperson name spelling against EP_PARAMETROS_MARGEM → COMISSAO. See [troubleshooting.md §6](troubleshooting.md).

---

## Scenario D — Raw material cost changed

**When to use:** supplier quoted a new price for a raw material; the impact needs to cascade through all finished goods.

**Steps:**

1. Open CostStructure_v2.
2. In the CUSTOS_MP tab, update column G (% reajuste) or column F (direct cost) for the MP that changed.
3. The cascade recalculates automatically:

   ```
   CUSTOS_MP → ESTRUTURA_PEs → PE_RESUMO → ESTRUTURA_PAs → RESULTADO
   ```

4. Go back to EP_MOTOR_PRECIFICACAO and run **Recalculate Margins** on every customer/table that uses products containing that MP.

**Expected outcome:** column G (PROD COST) updates to the new cost, and all downstream columns (margin, alerts) reflect the change.

**Tip:** For systemic reviews (e.g., base-oil price cycle), run **Export ALL tables** after recalculating — this regenerates every customer's file with the new cost reality.

---

## Scenario E — Adjust margin policy

**When to use:** executive decision to change minimum/target margin (e.g., during commercial plan review, or when repositioning a product family).

**Steps:**

1. Open EP_PARAMETROS_MARGEM → MARGEM_POLITICA tab.
2. Update the min/target margin values in the correct dimension:
   - **Family:** affects all SKUs of that family
   - **Classification:** affects all MINERAL/SYNTHETIC/etc. SKUs
   - **Strategic Classification:** affects specific SKUs (PROFIT DRIVER / HIDDEN STAR / CASH COW / DEAD WEIGHT)
3. Remember the lookup priority: **Strategic > Classification > Family** (most specific wins).
4. Go back to EP_MOTOR_PRECIFICACAO and run **Recalculate Margins** for any/all customers.
5. Compare column Y (Alert) before and after — how many SKUs changed status? (e.g., how many SKUs moved from "OK" to "BELOW MIN"?)

**Expected outcome:** a before/after snapshot showing how many SKUs now violate the new, stricter policy — or how many SKUs are now compliant under a relaxed policy.

**Tip:** This scenario is heavy. Consider exporting TABELA_NOVA to a versioned snapshot before running recalc, so you can compare Y (Alert) value by value.

---

## Operational cadence (recommended)

| Frequency | Action |
|---|---|
| Daily | Monitor column Y (Alert) for new BELOW MIN / BELOW TARGET items during quote processing |
| Weekly | Run Scenario A for the top-20 customers; review margin evolution |
| Monthly | Review EP_PARAMETROS_MARGEM → COMISSAO tab for new salespeople or reassignments |
| Quarterly | Full reprice cycle (Scenario B) tied to published inflation / FX shifts |
| On demand | Scenarios C, D, E when events trigger them |
