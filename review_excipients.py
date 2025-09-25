"""Flag potential issues in the generated excipient CSV file.

This script scans the extracted excipient dataset and records rows that
deserve manual review.  Each flagged row is written to a secondary CSV file
without modifying the original dataset so that reviewers can inspect and fix
issues safely.

Usage::

    python review_excipients.py [input_csv] [output_csv]

`input_csv` defaults to ``drug_excipients.csv`` and `output_csv` defaults to
``excipients_issues.csv`` when not provided.
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple


# Rows containing only these placeholder values in the excipient column are
# certainly incorrect and should be surfaced for review.
PLACEHOLDER_VALUES = {
    "",
    "either",
    "na",
    "n/a",
    "none",
    "not available",
    "unknown",
    "--",
}


# Regex patterns paired with descriptive labels to identify suspicious text
# fragments within the excipient column.  The expressions focus on phrases that
# usually describe packaging, chemical structure information, or other
# non-excipient details that previously slipped into the dataset.
PATTERN_FLAGS: Sequence[Tuple[str, re.Pattern[str]]] = (
    (
        "chemical_structure_reference",
        re.compile(
            r"\b(chemical structure|structural formula|molecular weight|"
            r"empirical formula|chiral)\b"
        ),
    ),
    (
        "regulatory_note",
        re.compile(r"\bno longer marketed\b"),
    ),
    (
        "packaging_reference",
        re.compile(
            r"\b(single[- ]use|dose pack|carton|container|kit|package|"
            r"provided as a|supplied as)\b"
        ),
    ),
    (
        "delivery_device",
        re.compile(
            r"\b(vial|syringe|needle|adapter|stopper|plunger|"
            r"safety device|applicator|needleless|plunger rod)\b"
        ),
    ),
    (
        "non_exipient_description",
        re.compile(
            r"\b(inactive granules|constituted with water|capsule is imprinted|"
            r"is a (?:sterile|white|clear|colorless)|osmolality|"
            r"available in dosage strengths|each vial contains)\b"
        ),
    ),
    (
        "image_reference",
        re.compile(r"\bimage of\b"),
    ),
)


# Units should have been removed during extraction.  If they remain, flag them
# as a potential issue.
UNIT_PATTERN = re.compile(r"\b\d+(?:\.\d+)?\s*(?:mg|g|kg|mcg|ug|µg|ml|l|%)\b")


def normalize(text: str) -> str:
    """Collapse whitespace and lowercase the provided text."""

    return re.sub(r"\s+", " ", text.strip().lower())


def collect_issues(row: Dict[str, str]) -> List[str]:
    """Inspect a CSV row and return a list of issue labels."""

    issues: List[str] = []
    excipients = row.get("excipients", "")
    normalized = normalize(excipients)

    if normalized in PLACEHOLDER_VALUES:
        issues.append("placeholder_value")
        # No need to continue evaluating placeholder rows; they are already
        # unusable.
        return issues

    if not normalized:
        issues.append("missing_excipients")
        return issues

    if UNIT_PATTERN.search(normalized):
        issues.append("contains_units")

    for label, pattern in PATTERN_FLAGS:
        match = pattern.search(normalized)
        if match:
            issues.append(f"{label}: '{match.group(0)}'")

    # Detect extremely short entries that are unlikely to be valid excipients.
    # Entries with five or fewer alphabetic characters and no delimiter often
    # point to truncated parsing.
    if ";" not in normalized and "," not in normalized:
        alphabetic = re.sub(r"[^a-z]", "", normalized)
        if len(alphabetic) <= 5:
            issues.append("very_short_entry")

    return issues


def flag_duplicates(rows: Sequence[Tuple[int, Dict[str, str], List[str]]]) -> None:
    """Append a duplicate-product issue when the same product appears twice."""

    product_map: Dict[str, List[List[str]]] = {}
    for _, row, issues in rows:
        product = normalize(row.get("product", ""))
        if not product:
            continue
        product_map.setdefault(product, []).append(issues)

    for issue_lists in product_map.values():
        if len(issue_lists) > 1:
            for issues in issue_lists:
                issues.append("duplicate_product")


def read_rows(path: Path) -> List[Tuple[int, Dict[str, str], List[str]]]:
    """Load the CSV rows with their 1-based line numbers."""

    with path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or "excipients" not in reader.fieldnames:
            raise ValueError("Input CSV must contain an 'excipients' column")

        rows: List[Tuple[int, Dict[str, str], List[str]]] = []
        for line_number, row in enumerate(reader, start=2):
            issues = collect_issues(row)
            rows.append((line_number, row, issues))
    return rows


def write_issues(path: Path, rows: Iterable[Tuple[int, Dict[str, str], List[str]]]) -> int:
    """Write flagged rows to ``path`` and return the number of records."""

    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["product", "row_number", "issues"])
        count = 0
        for line_number, row, issues in rows:
            if not issues:
                continue
            writer.writerow([row.get("product", ""), line_number, "; ".join(issues)])
            count += 1
    return count


def main(argv: Sequence[str] | None = None) -> int:
    argv = list(argv or sys.argv[1:])
    input_path = Path(argv[0]) if argv else Path("drug_excipients.csv")
    output_path = Path(argv[1]) if len(argv) > 1 else Path("excipients_issues.csv")

    rows = read_rows(input_path)
    flag_duplicates(rows)

    issue_count = write_issues(output_path, rows)
    print(f"Flagged {issue_count} potential issue(s). Output written to {output_path}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
