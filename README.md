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
“each vial contains ...”. It converts long dash separators to commas,
removes unrelated text such as concentration units or words like
“equivalent,” filters out packaging or device descriptors, deduplicates
the excipient names, and writes the result to `drug_excipients.csv`.

Subheadings like “tablet core:” or “film coating:” are split so their
ingredients are captured, while simple enumeration numbers are stripped and
packaging or formulation descriptors (e.g. “film” or “capsule shell”) are
discarded.

The generated CSV has two columns:

- `product` – product name
- `excipients` – deduplicated list of excipient names
