import csv
import re
import sys
from typing import List, Dict
from openpyxl import load_workbook

# capture excipients following the "Inactive Ingredients" label
EXCIPIENT_PAT = re.compile(r"inactive ingredients[^:]*:\s*(.*)", re.I | re.S)
# remove numbers followed by common concentration units (mg, g, %, etc.)
UNIT_PAT = re.compile(r"\b\d+(?:\.\d+)?\s*(mg|g|kg|mcg|ug|µg|ml|l|%)\b", re.I)

# tokens containing any of these keywords will be moved to the notes column
NOTE_KEYWORDS = {
    "capsule",
    "capsules",
    "tablet",
    "tablets",
    "structural",
    "formula",
    "vial",
    "bottle",
    "oral",
    "imprinted",
    "available",
    "granules",
    "reconstituted",
    "diluted",
    "administration",
    "inactive",
    "constituted",
    "form",
    "serotonin",
    "reuptake",
    "inhibitor",
}


def read_excel(path: str) -> List[Dict[str, str]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h) if h is not None else f"col{idx}" for idx, h in enumerate(rows[0])]
    data: List[Dict[str, str]] = []
    for row in rows[1:]:
        row_dict: Dict[str, str] = {}
        for h, v in zip(header, row):
            row_dict[h] = str(v) if v is not None else ""
        data.append(row_dict)
    return data


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r"\([^)]*\)", "", text)
    text = text.replace("\n", " ")
    text = text.replace(".", ";")
    text = re.sub(r"\s+", " ", text)
    text = UNIT_PAT.sub("", text)
    text = re.sub(r"\b\d+\b", " ", text)
    text = re.sub(r"[^a-z0-9,;\s-]", " ", text.lower())
    text = text.replace(" and ", "; ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_from_description(desc: str) -> str:
    if not desc:
        return ""
    m = EXCIPIENT_PAT.search(desc)
    if not m:
        return ""
    text = m.group(1)
    text = re.sub(r"(structural formula|chemical structure).*", "", text, flags=re.I)
    return text.strip()


def split_excipients_notes(text: str):
    tokens = [p.strip() for p in re.split(r"[;,]", text) if p.strip()]
    excipients = []
    notes = []
    for token in tokens:
        if "contains" in token:
            before, after = token.split("contains", 1)
            if before.strip():
                notes.append(before.strip())
            token = after.strip()
        if any(k in token for k in NOTE_KEYWORDS) or (
            "suspension" in token and len(token.split()) > 2
        ) or re.search(r"oral\s+solution", token):
            notes.append(token)
        else:
            excipients.append(token)
    # dedupe while preserving order
    def dedupe(items):
        seen = set()
        result = []
        for item in items:
            if item not in seen:
                seen.add(item)
                result.append(item)
        return result

    return dedupe(excipients), dedupe(notes)


def main():
    infile = sys.argv[1] if len(sys.argv) > 1 else "input.xlsx"
    outfile = sys.argv[2] if len(sys.argv) > 2 else "drug_excipients.csv"
    rows = read_excel(infile)
    results = []
    for row in rows:
        candidates = [
            row.get("English Product Name", ""),
            row.get("English Common Name", ""),
            row.get("English Drug Name", ""),
        ]
        product = ""
        for cand in candidates:
            if cand and re.search(r"[A-Za-z]", cand):
                product = cand.strip()
                break
        if not product:
            product = candidates[0]

        desc = row.get("Drug Description", "") or row.get("English Description", "")
        excip = parse_from_description(desc)
        if not excip:
            excip = row.get("Excipients", "")
        cleaned = clean_text(excip)
        if not cleaned:
            results.append((product, "", ""))
            continue
        excipients, notes = split_excipients_notes(cleaned)
        results.append((product, "; ".join(excipients), "; ".join(notes)))

    with open(outfile, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["product", "excipients", "notes"])
        for product, excips, notes in results:
            writer.writerow([product, excips, notes])


if __name__ == "__main__":
    main()
