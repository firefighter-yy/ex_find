# Search normalization specification

TISSUE-004 uses one deterministic representation for both indexed values and user keywords:

- Unicode text is normalized with NFKC, case-folded, trimmed, and runs of whitespace are collapsed to one space.
- Control, zero-width, bidi, soft-hyphen, and BOM characters are removed.
- Boolean values become `TRUE` or `FALSE`; dates and times use ISO-8601 text.
- Numeric values use decimal notation without insignificant trailing zeroes. String leading zeroes remain significant.
- Excel errors are searched as their visible text (for example, `#N/A`). Formula searches use the cached/calculated value by default; `FormulaMode.TEXT` searches formula text, and `BOTH` searches both.
- Excel COM `Value2` does not reliably contain number-format metadata. Therefore a numeric `0.15` cannot be silently treated as `15%`, and an Excel date serial cannot be silently converted to a date. Such display-format searches require future ingestion metadata.

Search returns one result per physical row, ordered by file selection order, worksheet order, and row number. Empty keywords are ignored and duplicate normalized keywords are collapsed.
