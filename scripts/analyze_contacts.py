#!/usr/bin/env python3
import csv, re, sys
from collections import defaultdict, Counter

IN = 'tmp_contacts.tsv'
OUT_PREFIX = 'analysis_contacts'

def is_email(s):
    return bool(re.search(r"@", s))

rows = []
with open(IN, newline='', encoding='utf-8') as f:
    for raw in f:
        # strip only trailing \n, keep possible internal tabs
        line = raw.rstrip('\n')
        cols = line.split('\t')
        rows.append(cols)

if not rows:
    print('No data')
    sys.exit(0)

header = rows[0]
data = rows[1:]

# Stats
total_rows = len(data)
col_counts = Counter(len(r) for r in rows)

# Detect rows where contactName looks like email (column 2 expected)
email_in_contact = []
numeric_facilityid = []
malformed_rows = []
dup_counter = Counter()

a_index = 0
# determine header indexes if header looks right
h_map = {}
for i,h in enumerate(header):
    h_map[h] = i

for i,r in enumerate(data, start=2):
    # r is list of columns for this line (1-based line i)
    if len(r) < 3:
        malformed_rows.append((i, r))
        continue
    facilityId = r[1].strip()
    contactName = r[2].strip()
    # facilityId numeric-only
    if re.fullmatch(r"\d+", facilityId):
        numeric_facilityid.append((i, facilityId, contactName))
    if is_email(contactName):
        email_in_contact.append((i, facilityId, contactName))
    # duplicates by facilityId + contactName
    dup_counter[(facilityId, contactName)] += 1
    # detect if createdAt column contains an email (shifted createdBy)
    # createdAt expected near last column; check any later columns for @
    for extra in r[6:]:
        if is_email(extra):
            # probable createdBy in extra column
            pass
    # malformed if odd createdAt format (very simple check)
    if len(r) < 7:
        malformed_rows.append((i, r))

# prepare summary
duplicates = [(k,c) for k,c in dup_counter.items() if c > 1]

# write outputs
with open(OUT_PREFIX + '_summary.txt', 'w', encoding='utf-8') as f:
    f.write(f'Total rows (excluding header): {total_rows}\n')
    f.write('Column counts distribution (including header line):\n')
    for cnt, num in sorted(col_counts.items()):
        f.write(f'  columns={cnt}: lines={num}\n')
    f.write('\n')
    f.write(f'Rows where contactName appears to be an email: {len(email_in_contact)}\n')
    for i, fid, name in email_in_contact[:50]:
        f.write(f'  line {i}: facilityId={fid} contactName={name}\n')
    f.write('\n')
    f.write(f'Rows where facilityId is numeric only: {len(numeric_facilityid)}\n')
    for i,fid,name in numeric_facilityid[:50]:
        f.write(f'  line {i}: facilityId={fid} contactName={name}\n')
    f.write('\n')
    f.write(f'Malformed / short rows: {len(malformed_rows)}\n')
    for i,r in malformed_rows[:20]:
        f.write(f'  line {i}: cols={len(r)} raw={r}\n')
    f.write('\n')
    f.write(f'Duplicate (facilityId, contactName) groups: {len(duplicates)}\n')
    for (fid,name),c in duplicates[:50]:
        f.write(f'  {fid} | {name} : {c} occurrences\n')

# write CSVs for inspection
with open(OUT_PREFIX + '_email_in_contact.csv', 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f)
    w.writerow(['line','facilityId','contactName'])
    for i,fid,name in email_in_contact:
        w.writerow([i,fid,name])

with open(OUT_PREFIX + '_numeric_facilityid.csv', 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f)
    w.writerow(['line','facilityId','contactName'])
    for i,fid,name in numeric_facilityid:
        w.writerow([i,fid,name])

with open(OUT_PREFIX + '_duplicates.csv', 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f)
    w.writerow(['facilityId','contactName','count'])
    for (fid,name),c in dup_counter.items():
        if c > 1:
            w.writerow([fid,name,c])

with open(OUT_PREFIX + '_malformed_rows.csv', 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f)
    w.writerow(['line','cols','raw'])
    for i,r in malformed_rows:
        w.writerow([i,len(r),'|'.join(r)])

print('Analysis complete. Summary written to', OUT_PREFIX + '_summary.txt')
print('CSV outputs:', OUT_PREFIX + '_email_in_contact.csv', OUT_PREFIX + '_numeric_facilityid.csv', OUT_PREFIX + '_duplicates.csv', OUT_PREFIX + '_malformed_rows.csv')
