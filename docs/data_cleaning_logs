# Data Cleaning & Validation Log
### Online Retail II Dataset — UCI Machine Learning Repository

**Tool:** Microsoft Excel (Tables, formulas, PivotTables, Power Query)
**Dataset:** ~1M rows of UK-based online retail transactions, Dec 2009–Dec 2011
**Purpose of this log:** document not just *what* was done, but the reasoning behind each decision — what was investigated, what evidence was checked before acting, and why each call was made. Intended as a reference for the eventual README and as a demonstration of the decision-making process behind the cleaning.

---

## 1. Setup

Converted the raw pasted data into an Excel **Table** to enable structured referencing (`Table1[ColumnName]`), auto-expanding ranges, and reliable filtering at scale. Removed stray blank rows/columns that had been pasted in around the actual data range, since these can interfere with sorting, filtering, and PivotTable source ranges later.

---

## 2. Description Column

**Observation:** A meaningful number of rows had a blank Description.

**Investigation:** Rather than assuming these were simple data-entry gaps, cross-checked what else was true about these rows. Found that all but one also had **Price = 0** and **no CustomerID** — a consistent pattern suggesting these weren't real product sales at all.

**Judgment call:** The one exception (non-zero price, blank description) was treated differently from the rest — flagged as worth checking individually rather than folding it into the same bucket, since a real sale with a missing description is a different situation from a non-sale row.

**Action taken:**
- Filled blank Description cells with `"Missing"` (using filter + Alt+Enter to fill all visible filtered cells at once — worked through a couple of selection issues first, since `Ctrl+Shift+End` was selecting the whole sheet rather than just the filtered column, and the Table's column-header-arrow selection method was also inconsistent between grabbing just data vs. data+header; ultimately used click-first-cell + `Ctrl+Shift+Down` to reliably select only the visible filtered rows)
- Confirmed the Price = 0 rows were non-product/write-off entries, copied them to an `Excluded_Rows` sheet, then removed from the main table
- **Note on mechanics:** the first deletion attempt used the `Delete` key, which only clears cell contents and leaves blank rows behind rather than actually removing them — corrected by using **Home → Delete → Delete Sheet Rows**, which properly removes rows and shifts data up

**Decision:** Retain the "Missing" label rather than deleting rows outright, to preserve the ability to reconcile total transaction counts and revenue later. Exclude the confirmed zero-price/no-description/no-customer-ID rows as non-sales entries — moved, not deleted, so the reasoning stays auditable.

---

## 3. StockCode Column

**Observation:** StockCodes are not uniformly 5-digit numbers — some carry a trailing letter (e.g., `48173C`).

**Initial instinct (corrected through reasoning):** First thought was that since the dataset documentation calls StockCode "nominal," the trailing letters should be stripped to standardize the codes. On reflection, recognized this was a misreading — "nominal" means *categorical* (not to be used in arithmetic), not "must be purely numeric." Trailing letters represent legitimate product variants (size/color), and stripping them would merge distinct products into one code — a fabricated error masquerading as cleaning. **Decision reversed: letters retained.**

**Real issue identified instead:** A separate category of StockCodes that aren't product codes at all — administrative/fee entries like `POST`, `D`, `M`, `BANK CHARGES`.

**Action taken:** Built an `Is_Product` flag to distinguish real product codes from administrative ones. Went through a few iterations to get the logic right:
- Initial approach (`ISNUMBER(VALUE(LEFT(...)))`) worked in Excel formulas
- When migrating this logic to Power Query for speed (dataset is too large for smooth filter/copy operations in the worksheet), the direct equivalent broke — Power Query's auto-detected column type had already converted problem values into errors *before* the custom formula ran, so `try...otherwise` wasn't catching what was expected
- Diagnosed and corrected the type-detection issue, then further simplified the logic to check only whether the first character is a digit (`List.Contains({"0".."9"}, Text.Start([StockCode],1))`), which is both simpler and avoids number-parsing ambiguity entirely

**Decision:** Flag non-product codes via `Is_Product`; defer actual removal to a single consolidated exclusion pass at the end of cleaning rather than removing incrementally.

---

## 4. Invoice Column — Cancellations & Adjustments

**Observation:** Invoice numbers aren't purely numeric — some start with letters.

**Process:** Initially identified "C" (Cancellation) as the only special case and built a `Cancelled_or_Not` TRUE/FALSE flag. Later discovered a second prefix, "A," associated with 3 rows. Rather than patching the existing flag with a second special case, recognized the better approach was to first establish the **complete** set of prefixes in the data before finalizing any categorization — avoiding a formula that would need to be re-patched every time a new letter turned up.
- Built `Invoice_Prefix` to extract and label every leading character across the full dataset, confirming only three categories exist: Numeric, "C", "A"
- Replaced the earlier binary flag with a complete `Sale_or_Not` classification: Normal Sale / Cancellation / Adjustment

**Investigation of the 3 "Adjustment" rows:** Checked what else was true about them — all tied to StockCode "B" (Bad debt adjustment, already flagged as non-product), no CustomerID, negative Price. Confirmed these are accounting write-offs, not real transactions.

**Decision:** Classify via `Sale_or_Not`; include in the same consolidated exclusion pass as other non-product rows, rather than treating "Adjustment" as a separate special case requiring its own removal logic.

---

## 5. Quantity Column

**Observation/hypothesis:** Cancelled invoices appear to consistently have negative Quantity — raised the question of whether anything further needed to be done with this column at all, or whether the existing `Cancelled_or_Not` flag already captured everything relevant.

**Validation (rather than assuming the hypothesis was correct):** Built a `Quantity_Sign` column (Negative/Non-negative) and cross-tabbed it against `Cancelled_or_Not` in a PivotTable to test the assumption directly, rather than taking it on faith.
- First PivotTable attempt put both fields in **Values**, which produced matching counts that looked like confirmation but were actually just counting the same rows twice from two angles — not a real cross-tab. Corrected by moving `Quantity_Sign` into **Columns**, which produced the actual 2×2 breakdown needed.

**Finding:** One exception — a cancelled invoice with a *non-negative* Quantity.

**Investigation of the exception:** Rather than treating this as a new, separate anomaly requiring its own fix, examined the full row: StockCode = "M" (Manual entry — already flagged as non-product), Price = 373.57, no CustomerID.

**Decision:** This is not a unique data-quality bug — it's a manual/administrative entry that happens to also carry a cancellation-style invoice prefix. It was already going to be excluded via the `Is_Product` flag regardless of its invoice prefix or quantity sign. No separate correction, no special-case note beyond documenting the overlap — avoids overstating a normal category-intersection as an inconsistency.

---

## 6. InvoiceDate Column

**Checks performed:**
- Confirmed the column is stored as a true date/time value, not text (important for correct sorting, filtering, and later date arithmetic)
- Checked for blanks — none found
- Verified MIN/MAX dates fell within the documented dataset range (01/12/2009–09/12/2011), confirming no stray out-of-range entries

**Action taken:** Added `Invoice_Date_Only` (`=INT([@InvoiceDate])`, formatted as Short Date) to separate the date from the time-of-day component, needed for calculating Recency later in RFM analysis. Kept the original `InvoiceDate` (with time) unchanged, in case time-of-day purchase patterns are worth exploring later — a deliberate choice not to discard information that isn't currently needed but might be later.

---

## 7. Price Column

**Observation:** 3 rows with negative Price.

**Investigation:** Checked what else was true of these rows — all 3 already matched the "Adjustment" invoices and non-product "B" StockCode identified in Section 4.

**Decision:** No separate action required — already fully captured by existing `Sale_or_Not` and `Is_Product` flags. Recognized this as the same underlying rows already investigated, rather than treating it as a new issue.

**Outstanding:** Large unit-price outliers (aside from the negative-price rows) noted as worth a closer look before cleaning is considered fully complete — not yet investigated.

---

## 8. CustomerID Column

**Observation:** A substantial proportion of rows have no CustomerID.

**Reasoning:** Two competing interpretations considered — (a) these are guest/non-account purchases, still real transactions; (b) these rows are simply unusable and should be excluded. Resolved by recognizing these aren't mutually exclusive: the rows are real and should be kept for revenue-level analysis, but they're structurally incompatible with customer-level analysis (RFM), which requires an identifiable customer per row.

**Action taken:** Added `Has_CustomerID` (TRUE/FALSE) rather than deleting or imputing a value — deliberately avoided inventing a customer identity for these rows.

**Decision:** Retain all rows in the cleaned dataset. Filter on `Has_CustomerID` only at the point of customer-level analysis, not during general cleaning — keeps the full dataset intact for revenue reconciliation while still making the customer-level subset easy to isolate later.

---

## 9. Country Column

**Checks performed:**
- Confirmed no blanks
- Built a PivotTable of unique Country values with counts to see the full picture in one glance, rather than scrolling
- Investigated **"Channel Islands"** specifically, since it wasn't an obviously recognizable "country" — confirmed it's a legitimate British Crown Dependency, distinct from the UK for shipping/customs purposes, not a data error
- Confirmed **"EIRE"** is the dataset's standard (if old-fashioned) label for Ireland, not a typo — decided to leave as-is rather than "correcting" something that wasn't wrong
- Found and corrected one genuine inconsistency: `"france"` (lowercase) vs. `"France"`

**Discrepancy investigated:** An initial `SUBTOTAL(COUNT,...)` check on the "Unspecified" entries returned 306, while the PivotTable returned 310. Rather than accepting either number blindly, investigated the mismatch and identified the cause: `SUBTOTAL` only counts currently visible (non-filtered) rows, and a stray filter elsewhere in the sheet was silently excluding 4 rows from that calculation. The PivotTable, drawing from the full source regardless of filter state, gave the correct count. Noted as a takeaway: **trust PivotTable counts over SUBTOTAL when filter state isn't explicitly verified.**

**Decision:** "Unspecified" (310 rows) treated as a legitimate placeholder category — deliberately entered by whoever recorded the transaction, not a missing value — retained as-is. No lookup table was ultimately needed, since the only real inconsistency found (France casing) was resolved directly, and everything else in the list checked out as legitimate.

---

## 10. Duplicate Invoice Line Items

**Framing the actual question first:** Recognized early that a naive duplicate check on Invoice alone would be meaningless, since one invoice legitimately contains many product line items. Defined the real question as: does the same **Invoice + StockCode combination** ever appear more than once?

**Method:** Given the dataset's size, filtering/copying in the worksheet had already proven slow for earlier steps, so this check was done in Power Query from the start.
- Built a `Duplicate_Check` query: grouped by Invoice + StockCode, counted rows as `Dup_Count`
- Extended the grouping to also capture Min/Max of Quantity, Price, and InvoiceDate per group — reasoning: if Min equals Max across a group, every value in that group must be identical (a mathematical shortcut to test "are all values the same" using only aggregate functions)
- Combined these into a `Fully_Identical` flag (`Min = Max` for all three metrics, joined with `and`, meaning **all three** must match for a group to count as truly identical — one mismatch is enough to disqualify it)
- Merged `Dup_Count` back into the main table via a Left Outer merge (after an initial attempt merged in the wrong direction, bringing Table1's fields into the Duplicate_Check query instead of the reverse — corrected by re-running the merge from Table1's side)

**Reasoned conclusion on how to treat each case, worked out independently rather than following a prescriptive rule:**
- If duplicate rows differ in Price or InvoiceDate, they're unambiguously two separate real transactions — not something to "fix"
- If duplicate rows differ only in Quantity, there's no reliable way to determine which entry (if either) is erroneous — arbitrarily deleting one risks discarding real data
- Therefore: the only safe, non-arbitrary action is limited to groups where **everything matches exactly** (`Fully_Identical = TRUE`) — only then is it safe to say the repeated row is truly redundant rather than a distinct event

**Finding:** Filtered to `Dup_Count > 1 AND Fully_Identical = TRUE` — zero results.

**Decision:** No rows removed. Every duplicated Invoice+StockCode combination in the dataset represents a genuinely distinct transaction (differing in Price, Quantity, or Date), not an accidental repeat entry. No further action needed on duplicates.

---

## Summary — Flag/Helper Columns Created

| Column | Purpose |
|---|---|
| `Is_Product` | Distinguishes real product StockCodes from administrative/fee codes |
| `Sale_or_Not` | Classifies each Invoice as Normal Sale / Cancellation / Adjustment |
| `Quantity_Sign` | Negative vs. Non-negative Quantity (used to validate the cancellation hypothesis) |
| `Invoice_Date_Only` | Date without time component, for Recency calculations |
| `Has_CustomerID` | Flags rows with no CustomerID (excluded only at customer-level analysis stage) |

## Rows Identified for Exclusion (pending consolidated removal pass)
- Blank Description + Price = 0 (non-product/write-off entries) — already moved to `Excluded_Rows`
- Non-product StockCodes (`Is_Product = FALSE`)
- Adjustment invoices (`Sale_or_Not = "Adjustment"`)

## Deliberate Non-Actions (investigated, and correctly left alone)
- StockCode trailing letters — legitimate product variants, not errors
- "EIRE," "Channel Islands," "Unspecified" country values — all legitimate, not inconsistencies
- Cancelled invoice with non-negative Quantity — already covered by non-product exclusion, not a separate anomaly
- Duplicate Invoice+StockCode groups with differing Price/Quantity/Date — genuinely separate transactions, not errors

## Outstanding Before Cleaning Is Considered Complete
1. Review Price outliers (unusually large unit prices)
2. Review Quantity outliers (unusually large quantities)
3. Execute the consolidated exclusion pass — move all remaining flagged rows to the excluded sheet(s) with documented reasons
4. Proceed to analysis phase (RFM segmentation via PivotTables)
