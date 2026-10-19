---
name: inventory-clearing-recon
description: Reconcile an inventory clearing/transfer account that should net to zero (every in has a matching out), find the items that didn't clear, and build a formula-driven reconciliation on the same tab without changing the source data.
---

## When to use
- The user has an inventory movement / location transfer / clearing account export and says the beginning balance is 0 (or another given figure) and all inventory should move in and out.
- They ask to "reconcile the inventory account", "find which items didn't clear", or "find what's making up the balance".
- Typical export columns: Div., Product Line, GL Department, Type, Number, Post.Date, Per.Date, Debit, Credit, Narrative, Amount, T.Balance. The Narrative holds the item, like `ITEMCODE x -1 @ 113.45`.

## Ground rules
- **Never edit, sort, delete or re-type the source data columns.** Only add helper columns to the right, apply formatting, add conditional formatting and freeze panes.
- Every number in the reconciliation must be a formula that points back to the data. The only typed inputs are the beginning balance and the expected ending balance (0), shown in blue font.
- Sheet names may have trailing spaces (e.g. "Inventory movement "). Use the exact name from the workbook metadata.

## Steps

### 1. Profile the data
- Use get_cell_ranges to read the header plus about 10 rows. Read the last 20 rows to get the last row number and the final running balance.
- Use execute_office_js to read the full range and return only a summary:
  - the count of each value in Type
  - a check that Debit − Credit = Amount on every row
  - the total of Amount
  - a dry run of the item extraction (text before `" x "` in Narrative), with any rows where it fails
  - the items whose Amount nets to something other than 0 after rounding to 2 decimals, and the sum of those items
- If the Narrative format is different (no `" x "`), stop and ask the user how to identify the item.
- **Done when:** the sum of the non-zero items equals (final balance − beginning balance).

### 2. Add helper columns to the right of the data
Put these in the first three empty columns (M:O when the data is A:L). Use the real last row for `N` below.
- **Item:** `=LEFT(J2,FIND(" x ",J2)-1)`
- **Item Net Amount:** `=ROUND(SUMIFS($K$2:$K$N,$M$2:$M$N,M2),2)`
- **Status:** `=IF(N2=0,"Cleared","Open")`

Write row 2, then fill down with autoFill. Copy the header formatting from the existing header row and apply the number format `#,##0.00;(#,##0.00);-`. For large sheets, switch calculation to manual during the write and set it back to automatic afterwards.
- **Done when:** the number of Open rows matches the dry run from step 1.

### 3. Build the reconciliation (leave one blank spacer column, e.g. P, and start at Q2)
- **Q2:** Inventory Account Reconciliation (bold, 14pt). **Q3:** the period text (italic, grey).
- **Q5:R14:**
  - Beginning balance | 0 (blue input)
  - Add: Inventory in (Debits) | `=SUM(Debit col)`
  - Less: Inventory out (Credits) | `=-SUM(Credit col)`
  - Calculated ending balance | `=ROUND(SUM(R5:R7),2)` (bold, top border)
  - Ending balance per ledger (T.Balance) | `=L<last row>`
  - Difference | `=ROUND(R8-R9,2)` (bold)
  - Ending balance should be (all in = out) | 0 (blue input)
  - Unreconciled balance to explain | `=R9-R11`
  - Less: Total uncleared items (below) | `=-SUM(V18#)`
  - Unexplained variance | `=R12+R13` (bold, top border, double bottom border)
- **Q16:** Uncleared Items (no offsetting in/out entry)
- **Q17:V17:** the headers Item, Doc #, Post Date, Per., Narrative, Amount (dark navy fill #1F3864, white bold font)
- **Q18:V18:** one FILTER per column so the list updates itself:
  - `=FILTER(M2:M<N>,O2:O<N>="Open","None")`
  - the same pattern for Number, Post.Date, Per.Date, Narrative and Amount (use `0` as the empty value for Amount)
- **Formats:** dates m/d/yyyy, amounts `#,##0.00;(#,##0.00);-`. Set sensible column widths (Q ≈ 230pt, U ≈ 160pt), keep the spacer column narrow, and keep row 1 at standard height with wrap off.
- If FILTER returns #NAME? (older Excel), write the open items as static rows pulled with INDEX/MATCH or SUMIFS, and tell the user.

### 4. Make the open items easy to find
- Add a conditional format on the data plus helper columns, using the custom formula `=$O2="Open"`, with fill #FFF2CC and font #C00000.
- Freeze row 1.

### 5. Verify before reporting
- Re-read Q5:V(last spill row) as **text** and confirm:
  - Calculated balance = ledger balance
  - Difference shows "-"
  - Unexplained variance shows "-"
  - the FILTER spill has the expected number of rows
- Check that freeze panes are set (`freezePanes.getLocationOrNullObject`) and that the conditional format exists.
- Optionally render an image of the reconciliation block (`getImage`) to check the layout.
- **Done when:** all of these checks pass.

## Report to the user
- A short table of the uncleared items (Item, Doc #, Date, Amount) and their total.
- Where the reconciliation lives and confirmation that the unexplained variance is 0.
- What was added (helper columns, highlighting, freeze panes) and that the source columns were not changed.
- Which ranges you re-read to verify.

## Edge cases
- **Item partly cleared** (e.g. in 2, out 1): it shows as Open and every row for that item is listed. The Amount column still sums to the item's net.
- **Zero-cost rows** (`x 1 @` with no price): they net to 0 and show as Cleared.
- **Unexplained variance not 0:** look for rows where the item couldn't be extracted, or where Debit − Credit ≠ Amount, and report them.
