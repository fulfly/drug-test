This repository contains a simple utility for extracting excipient names from
the provided `input.xlsx` spreadsheet.

To run the extraction and generate a cleaned CSV of product names and
excipients:

```
python extract_excipients.py input.xlsx drug_excipients.csv
```

The script reads the Excel file, locates excipient sections labelled with
phrases such as “Inactive Ingredients,” “Other Ingredients,” or
“Preservatives,” and also captures excipients described in diluent
composition lines, “micro‑encapsulated in” phrases, or statements like
“each vial contains ...” (including cases that list numeric concentrations
with decimals). It converts long dash separators to commas, removes
unrelated text such as concentration units or words like “equivalent,”
filters out packaging or device descriptors, deduplicates the excipient
names, and writes the result to `drug_excipients.csv`. When trimming out
product names, the parser preserves common ionic words (e.g. “sodium,”
“chloride,” “hydroxide”) so salts remain intact in the final output.

Subheadings like “tablet core:” or “film coating:” are split so their
ingredients are captured, while simple enumeration numbers are stripped and
packaging or formulation descriptors (e.g. “film” or “capsule shell”) are
discarded.

The generated CSV has two columns:

- `product` – product name
- `excipients` – deduplicated list of excipient names

### Reviewing the output

Run the review helper to highlight rows that still contain suspicious text or
missing data without altering the source CSV:

```
python review_excipients.py drug_excipients.csv excipients_issues.csv
```

The script scans the `excipients` column for placeholder values, lingering
units, packaging/device terminology (e.g. “vial,” “needle,” “single-use”),
chemical-structure references, and other phrases that typically indicate the
row needs manual cleanup. Each problematic entry is written to
`excipients_issues.csv` together with the 1-based row number from the original
dataset and the list of detected issues so the original spreadsheet can be
reviewed safely before any edits are made.
