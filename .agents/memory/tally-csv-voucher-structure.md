---
name: Tally 7.2 CSV voucher export structure
description: How to parse RSU/Tally 7.2 Voucher + Master CSV exports when no item-level data exists, and how to classify ledger lines correctly.
---

## Shape of the data
- `RSUTallyVoucher*.csv`: one row per **accounting ledger line**, not per item. Header:
  `Sno,Voucher Date,Voucher Type,Voucher No,Transaction ref No,Account Name,Account Code,Amount,Transaction ref No,Amount,Narration,Tax Percentage,Tax Type`.
  There is **no item name, quantity, rate, or HSN anywhere** in this export — only party/GST-rate-bucket/tax/round-off/discount ledger lines per bill.
- `RSUTallyMaster*.csv`: pure ledger master (`Type,AccountName,AccountCode,Account Groups Name,...`), one row per account (Debtor/Creditor/Bank/etc). **No broker field.**
- Only `Voucher Type` in `{Sales, Credit Note, Purchase}` map to Sale/Sale Return/Purchase outputs — `Receipt`/`Journal` are payment/settlement entries with irregular embedded-comma narration (multiple bill refs concatenated) that shifts column counts; they must be filtered out *before* doing any positional CSV parsing, not handled by it.
- Confirmed empirically: every row of the 3 target voucher types has exactly 13 comma-separated fields (no quoting needed) — safe to `split(",")` directly once Receipt/Journal are excluded.

## Classifying ledger lines within a bill (group by Voucher No)
Classify each line **by Account Code pattern**, not by row position — position/order of lines within a bill is not fixed:
- Revenue/purchase (GST-rate bucket) line: Account Code contains `RSUSALES` or `RSUPURCHASE` (case-insensitive). A single bill can have 2+ of these (one per GST% rate) — this is expected and normal, not an error.
- Tax lines: Account Code exactly `CGST`/`SGST`/`IGST` — informational only; the target import template only needs GST% + taxable value per line, it computes tax itself, so these don't need to be reproduced.
- Round off: Account Code or Account Name contains "round off".
- Party (customer/supplier) line: look up Account Code in the Master file to get `Account Groups Name`; party group is `Sundry Debtor` (sales/returns) or `Sundry Creditor` (purchases). Voucher Account Codes are often composite (`DPBC_1619_P1538`) where the Master's own AccountCode is the **last underscore-delimited segment** (`P1538`) — try the full code first, then fall back to the last segment.
- Everything else in this dataset was just two known discount codes (`SCH01` = scheme discount, `OTH01` = other discount) — every bill has ≤2 "other" lines, so they map cleanly to a template's two Add/less slots without needing an open-ended list.
- Every target-type bill had exactly one party line — if that ever isn't true, treat missing-party bills as a data quality signal, not a silent default.

**Why:** discovered by writing an explore-subagent analysis script over the real CSVs rather than guessing from a sample — the row-count histograms (party lines, GST-bucket lines, "other" lines) confirmed the classification rule covers 100% of rows with no ambiguity.

## Party name cleanup
Master/voucher party names often carry a trailing garbage suffix from the sync tool, e.g. `HUNNY GENERAL-D-D-D-D-D-D`, `KANAIYA MED STO*KATARGAM*-P-D-D`. Strip only the trailing run of `-<single-letter>` tokens (regex `/(-[A-Za-z])+$/`) — keep `*AREA*` segments since those often encode meaningful locality info, don't discard them wholesale.

## Reconciling item-level templates with item-less source data
When the target Excel template expects item-level rows (Item Name, Qty, Rate, GST%) but the source is bill/ledger-level only: emit one row per (bill, GST%-bucket) with Qty=1 and Rate=taxable value, and do NOT feed these synthetic rows into any Item Master merge step — they'd pollute a real item catalog with fake generic "item" entries.
