import zipfile
import xml.etree.ElementTree as ET
import re
import csv
import sys

EXCIPIENT_PAT = re.compile(
    r"inactive ingredients[^:]*:\s*(.*)", re.I | re.S
)

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        strings = []
        with z.open('xl/sharedStrings.xml') as f:
            tree = ET.parse(f)
            strings = [t.text for t in tree.findall('.//m:t', NS)]
        with z.open('xl/worksheets/sheet1.xml') as f:
            sheet = ET.parse(f)
        data = []
        for row in sheet.findall('.//m:sheetData/m:row', NS):
            row_data = []
            for c in row.findall('m:c', NS):
                t = c.get('t')
                v = c.find('m:v', NS)
                val = v.text if v is not None else ''
                if t == 's':
                    val = strings[int(val)]
                row_data.append(val)
            data.append(row_data)
    return data


def clean_text(text: str) -> str:
    if not text:
        return ''
    text = re.sub(r'\s+', ' ', text.replace('\n', ' '))
    text = re.sub(r'\b\d+(?:\.\d+)?\s*(mg|g|kg|mcg|ug|µg|ml|l|%)\b', '', text, flags=re.I)
    text = re.sub(r'[^A-Za-z0-9,;\s.-]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.lower().strip()


def parse_from_description(desc: str) -> str:
    if not desc:
        return ''
    m = EXCIPIENT_PAT.search(desc)
    if not m:
        return ''
    text = m.group(1)
    text = re.sub(r"\([^\)]*\)", "", text)
    return text.strip()


def main():
    infile = sys.argv[1] if len(sys.argv) > 1 else 'test.xlsx'
    outfile = sys.argv[2] if len(sys.argv) > 2 else 'drug_excipients.csv'

    data = read_xlsx(infile)
    header = data[0]
    rows = []
    for r in data[1:]:
        if len(r) > len(header):
            extra = [x or '' for x in r[len(header)-1:]]
            r = r[:len(header)-1] + [' '.join(extra)]
        if len(r) < len(header):
            r += [''] * (len(header) - len(r))
        rows.append(dict(zip(header, r)))

    results = []
    for row in rows:
        candidates = [
            row.get('English Product Name', ''),
            row.get('English Common Name', ''),
            row.get('English Drug Name', ''),
        ]
        product = ''
        for cand in candidates:
            if cand and re.search(r'[A-Za-z]', cand):
                product = cand
                break
        if not product:
            product = candidates[0]
        excip = parse_from_description(row.get('Drug Description', '') or row.get('English Description', ''))
        if not excip:
            excip = row.get('Excipients', '')
        cleaned = clean_text(excip)
        names = [p.strip() for p in re.split(r'[;,]', cleaned) if p.strip()]
        results.append((product, '; '.join(names)))

    with open(outfile, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['product', 'excipients'])
        for product, excips in results:
            writer.writerow([product, excips])

if __name__ == '__main__':
    main()
