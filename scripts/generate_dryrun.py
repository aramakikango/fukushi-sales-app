#!/usr/bin/env python3
import csv, re
IN = 'tmp_contacts.tsv'
OUT_BEFORE_AFTER = 'analysis_contacts_dryrun_before_after.csv'
OUT_PROPOSED = 'analysis_contacts_dryrun_proposed.csv'

email_re = re.compile(r"\S+@\S+")

rows = []
with open(IN, newline='', encoding='utf-8') as f:
    for raw in f:
        line = raw.rstrip('\n')
        cols = line.split('\t')
        rows.append(cols)

if not rows:
    print('no data')
    raise SystemExit(1)

header = rows[0]
data = rows[1:]

# build new header for proposed: add contactEmail and needsReview
proposed_header = header.copy()
# ensure contactEmail exists
if 'contactEmail' not in proposed_header:
    # insert after contactName (if present), else append
    if 'contactName' in proposed_header:
        idx = proposed_header.index('contactName')
        proposed_header.insert(idx+1, 'contactEmail')
    else:
        proposed_header.append('contactEmail')
# add needsReview
if 'needsReview' not in proposed_header:
    proposed_header.append('needsReview')

changed_rows = []
proposed_all = []

for i, r in enumerate(data, start=2):
    # normalize length to header length
    row = r.copy()
    # pad to header length
    if len(row) < len(header):
        row += [''] * (len(header) - len(row))
    # map by header names
    row_map = {h: (row[idx] if idx < len(row) else '') for idx, h in enumerate(header)}
    proposed = row.copy()
    # ensure proposed has place for contactEmail
    if 'contactEmail' in proposed_header and 'contactEmail' not in row_map:
        # find index of contactName in proposed_header
        if 'contactName' in header:
            # compute position
            pass
    # detect email in contactName
    contactName = row_map.get('contactName','').strip()
    needs = []
    contactEmail = ''
    if contactName and email_re.search(contactName):
        contactEmail = contactName
        # blank the contactName (proposal)
        row_map['contactName'] = ''
        needs.append('movedEmailFromName')
    # facilityId numeric-only
    facilityId = row_map.get('facilityId','').strip()
    if facilityId and facilityId.isdigit():
        needs.append('numericFacilityId')
    # assemble proposed row according to proposed_header
    out_proposed = []
    for h in proposed_header:
        if h == 'contactEmail':
            val = contactEmail or ''
        elif h == 'needsReview':
            val = ';'.join(needs) if needs else ''
        else:
            val = row_map.get(h,'')
        out_proposed.append(val)
    proposed_all.append(out_proposed)
    if needs:
        # before row as original (pad to header length)
        before = [row[idx] if idx < len(row) else '' for idx in range(len(header))]
        after = out_proposed
        changed_rows.append((i,before,after))

# write before/after CSV (only changed rows)
with open(OUT_BEFORE_AFTER, 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f)
    # header: line + original headers + proposed headers
    w.writerow(['line'] + ['orig_'+h for h in header] + ['prop_'+h for h in proposed_header])
    for line,before,after in changed_rows:
        w.writerow([line] + before + after)

# write proposed full CSV
with open(OUT_PROPOSED, 'w', encoding='utf-8', newline='') as f:
    w = csv.writer(f)
    w.writerow(proposed_header)
    for row in proposed_all:
        w.writerow(row)

print('Dryrun generated:', OUT_BEFORE_AFTER, OUT_PROPOSED)
