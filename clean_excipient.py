import pandas as pd
import re
import sys

EXCIPIENT_PAT = re.compile(r"inactive ingredients[^:]*:\s*(.*)", re.I | re.S)

UNITS = r"mg|g|kg|mcg|ug|µg|ml|l|%"

CLEAN_UNITS = re.compile(r"\b\d+(?:\.\d+)?\s*(" + UNITS + r")\b", re.I)
SPECIAL_CHARS = re.compile(r"[^A-Za-z0-9,;\s.-]")


def clean_text(text: str) -> str:
    if not isinstance(text, str) or not text:
        return ""
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    text = CLEAN_UNITS.sub("", text)
    # remove standalone numbers that may remain
    text = re.sub(r"\b\d+\b", "", text)
    text = SPECIAL_CHARS.sub(" ", text)
    text = re.sub(r"\s+", " ", text)
    return text.lower().strip()


def parse_from_description(desc: str) -> str:
    if not desc:
        return ""
    m = EXCIPIENT_PAT.search(desc)
    if not m:
        return ""
    text = m.group(1)
    text = re.sub(r"\([^\)]*\)", "", text)
    return text.strip()


def get_product(row: dict) -> str:
    candidates = [
        row.get("English Product Name", ""),
        row.get("English Common Name", ""),
        row.get("English Drug Name", ""),
    ]
    for c in candidates:
        if isinstance(c, str) and re.search(r"[A-Za-z]", c):
            return c
    return candidates[0] if candidates else ""


def process_file(infile: str, outfile: str):
    df = pd.read_excel(infile)
    rows = []
    for _, row in df.iterrows():
        row_dict = row.to_dict()
        product = get_product(row_dict)
        desc = row_dict.get("Drug Description") or row_dict.get("English Description") or ""
        excip = parse_from_description(desc)
        if not excip:
            excip = row_dict.get("Excipients", "")
        cleaned = clean_text(excip)
        names = [p.strip() for p in re.split(r"[;,]", cleaned) if p.strip()]
        # deduplicate names
        dedup = []
        seen = set()
        for n in names:
            if n not in seen:
                dedup.append(n)
                seen.add(n)
        rows.append({"product": product, "excipients": "; ".join(dedup)})
    out_df = pd.DataFrame(rows)
    out_df = out_df.drop_duplicates(subset=["product"])
    out_df.to_csv(outfile, index=False)


def main():
    infile = sys.argv[1] if len(sys.argv) > 1 else "test 2.xlsx"
    outfile = sys.argv[2] if len(sys.argv) > 2 else "drug_excipients.csv"
    process_file(infile, outfile)


if __name__ == "__main__":
    main()
