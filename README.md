This repository contains a simple utility for extracting excipient names from
the provided `input.xlsx` spreadsheet.

To run the extraction and generate a cleaned CSV of product names and
excipients:

```
python extract_excipients.py input.xlsx drug_excipients.csv
```

The script reads the Excel file, locates excipient sections labelled as
“Preservative,” “Inactives,” or “Inactive Ingredients,” removes unrelated text
such as concentration units and special characters, deduplicates the excipient
names, and writes the result to `drug_excipients.csv`.

The generated CSV has three columns:

- `product` – product name
- `excipients` – deduplicated list of excipient names
- `notes` – descriptive phrases that are not excipients (e.g. formulation
  details like "each capsule is imprinted with …" or "inactive granules are
  constituted with water to form a suspension")
