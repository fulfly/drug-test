import zipfile
import xml.etree.ElementTree as ET
import re
import csv

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


def main():
    data = read_xlsx('test.xlsx')
    header = data[0]
    rows = [dict(zip(header, r)) for r in data[1:]]

    results = []
    for row in rows:
        product = (
            row.get('English Product Name')
            or row.get('English Common Name')
            or row.get('English Drug Name')
            or ''
        )
        excip = row.get('Excipients', '')
        cleaned = clean_text(excip)
        names = [p.strip() for p in re.split(r'[;,]', cleaned) if p.strip()]
        results.append((product, '; '.join(names)))

    with open('drug_excipients.csv', 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['product', 'excipients'])
        for product, excips in results:
            writer.writerow([product, excips])

if __name__ == '__main__':
    main()
