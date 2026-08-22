---
name: LeverEdge Sales Register column names
description: Exact column names and positions in LeverEdge Sales Register Excel files; critical for parseSalesRegister() bill number detection.
---

## Key facts

- Header row is at **index 12** (0-based) in the Sales Register sheet
- Bill number column: **col 3 = "BillRefNo"** — normalizes to `"billrefno"` (NOT "bill no" with space)
- Adjustments column: **col 21 = "Adjustments"** — added to BTRP DIS discount amount
- BTPR column: **col 37 = "BTPR SchDisc"**
- Ushop column: **col 42/43 = "Ushop Redemption"**
- Shikhar column: **col 46/47 = "Shikhar Scheme"**
- Bill value column: **col 34 = "BillValue"** — normalizes to `"billvalue"` (NOT "bill value" with space)

## Discount mapping rules
- `BTRP DIS` = `BTPR SchDisc` + `Adjustments`
- `Add / less 1` (`USHOP DIS`) = `Ushop Redemption` + `BTRP DIS` (includes Adjustments)
- `Add / less 2` (`SHIKHAR DIS`) = `Shikhar Scheme`

## Why this matters

The original `parseSalesRegister()` searched for `"bill no"` (with space) using `includes()`. The actual column name `"BillRefNo"` normalizes to `"billrefno"` which does NOT include the substring `"bill no"` (has a space). This caused `billNoCol = -1`, all rows were skipped, and the discount map was always empty.

**Fix:** Add `"billrefno"` to the `findCol()` alias list for the bill number column.

## How to apply

In any future change to `parseSalesRegister()` or similar Sales Register parsers, always include `"billrefno"` as an alias for bill number detection.
