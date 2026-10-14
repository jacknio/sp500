r"""Build the DATA QCEW tab from extracted BLS 2025 annual-by-industry CSVs.
Place the extracted CSV files anywhere under H:\\labor data (subfolders are fine).
Run: python build_qcew_data_tab.py
Optional: python build_qcew_data_tab.py "H:\labor data"
Python standard library only. Produces DATA_QCEW_2025.csv and QCEW_QA_2025.csv.
"""
from pathlib import Path
import csv
import sys

DATA_DIR = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(r'H:\labor data')
OUTPUT_DIR = DATA_DIR
if not DATA_DIR.is_dir():
    raise FileNotFoundError(f'Folder not found: {DATA_DIR}')
SOURCE_URL = 'https://www.bls.gov/cew/downloadable-data-files.htm'

# Exact 19 non-overlapping NAICS sectors matching the existing model.
SECTORS = [
    ('11', 'Agriculture, forestry, fishing and hunting'),
    ('21', 'Mining, quarrying, and oil and gas extraction'),
    ('22', 'Utilities'),
    ('23', 'Construction'),
    ('31-33', 'Manufacturing'),
    ('42', 'Wholesale trade'),
    ('44-45', 'Retail trade'),
    ('48-49', 'Transportation and warehousing'),
    ('51', 'Information'),
    ('52', 'Finance and insurance'),
    ('53', 'Real estate and rental and leasing'),
    ('54', 'Professional, scientific, and technical services'),
    ('55', 'Management of companies and enterprises'),
    ('56', 'Administrative and support and waste management'),
    ('61', 'Educational services'),
    ('62', 'Health care and social assistance'),
    ('71', 'Arts, entertainment, and recreation'),
    ('72', 'Accommodation and food services'),
    ('81', 'Other services (except public administration)'),
]

# IMPORTANT: read sector-level files only; do not sum nested 3/4/5/6-digit industries.
def find_file(folder, code):
    matches = [p for p in folder.rglob('*.csv')
               if p.name.startswith('2025.annual ' + code + ' ')]
    if len(matches) != 1:
        raise ValueError(f'Expected one CSV for NAICS {code}; found {len(matches)}: {matches[:5]}')
    return matches[0]

def get_us_private_record(path, expected_code):
    with path.open('r', newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        found = [r for r in reader if
                 r['area_fips'] == 'US000'
                 and r['own_code'] == '5'
                 and r['industry_code'] == expected_code
                 and r['year'] == '2025'
                 and r['qtr'] == 'A'
                 and r['size_code'] == '0']
    if len(found) != 1:
        raise ValueError(f'Expected one US/private/annual/all-sizes row for NAICS {expected_code}; found {len(found)} in {path}')
    return found[0]

rows = []
for code, name in SECTORS:
    original = find_file(DATA_DIR, code)
    r = get_us_private_record(original, code)
    if r['disclosure_code']:
        raise ValueError(f'Unexpected disclosure at {code}: {r["disclosure_code"]}')
    rows.append([code, name, int(r['total_annual_wages']), SOURCE_URL])

aggregate = get_us_private_record(find_file(DATA_DIR, '10'), '10')
unclassified = get_us_private_record(find_file(DATA_DIR, '99'), '99')
national = int(aggregate['total_annual_wages'])
unclassified_wages = int(unclassified['total_annual_wages'])
sector_total = sum(r[2] for r in rows)
difference = national - sector_total
if difference != unclassified_wages:
    raise ValueError(f'QA FAIL: Difference {difference:,} != NAICS 99 {unclassified_wages:,}')

# Matches DATA QCEW existing workbook layout A:D.
out = OUTPUT_DIR / 'DATA_QCEW_2025.csv'
with out.open('w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow(['NAICS', 'Industry', 'QCEW 2025 private annual wages ($)', 'QCEW source'])
    writer.writerows(rows)

qa = OUTPUT_DIR / 'QCEW_QA_2025.csv'
with qa.open('w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow(['Check', 'Value ($)', 'Interpretation'])
    writer.writerow(['19-sector annual wages', sector_total, 'Sum of the 19 industry rows'])
    writer.writerow(['US private all industries', national, 'US000 / own_code=5 / NAICS 10'])
    writer.writerow(['NAICS 99 unclassified', unclassified_wages, 'Excluded from the 19 sectors'])
    writer.writerow(['Difference', difference, 'Must equal NAICS 99'])
    writer.writerow(['QA pass', difference == unclassified_wages, 'True means industry coverage reconciles'])

print('Saved:', out)
print('Saved:', qa)
print(f'19 industries: ${sector_total:,.0f}')
print(f'US private total: ${national:,.0f}')
print(f'Unclassified (NAICS 99): ${unclassified_wages:,.0f}')
print('QA PASS: sector sum + unclassified = national total')
