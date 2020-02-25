from docx import Document
from unidecode import unidecode
import json
import timeit


doc = Document('./ts-0004.docx')

# meta-info of short name tables in TS-0004 temporal v4.0.0 (as of 2020-Aug-25)
version = 'v4.0.0'
short_name_tables = {
    'first_idx': 370,
    'last_idx': 380,
    'column_info': [
        {'shortName': 3, 'longName': 0, 'occursIn': 2, 'category': 'primitive parameter'}, #[370] Table 8.2.2 1: Primitive parameter short names
        {'shortName': 2, 'longName': 0, 'occursIn': 1, 'category': 'Primitive root element'}, #[371] Table 8.2.2-2: Primitive root element short names
        {'shortName': 2, 'longName': 0, 'occursIn': 1, 'category': 'Resource attribute'}, #[372] Table 8.2.3-1: Resource attribute short names (1/6)
        {'shortName': 2, 'longName': 0, 'occursIn': 1, 'category': 'Resource attribute'}, #[373] Table 8.2.3-2: Resource attribute short names (2/6)
        {'shortName': 2, 'longName': 0, 'occursIn': 1, 'category': 'Resource attribute'}, #[374] Table 8.2.3-3: Resource attribute short names (3/6)
        {'shortName': 2, 'longName': 0, 'occursIn': 1, 'category': 'Resource attribute'}, #[375] Table 8.2.3-4: Resource attribute short names (4/6)
        {'shortName': 2, 'longName': 0, 'occursIn': 1, 'category': 'Resource attribute'}, #[376] Table 8.2.3-5: Resource attribute short names (5/6)
        {'shortName': 2, 'longName': 0, 'occursIn': 1, 'category': 'Resource attribute'}, #[377] Table 8.2.3-6: Resource attribute short names (6/6)
        {'shortName': 1, 'longName': 0, 'occursIn': -1, 'category': 'Resource and specialization type'}, #[378] Table 8.2.4-1: Resource and specialization type short names
        {'shortName': 2, 'longName': 0, 'occursIn': 1, 'category': 'Complex data type member'}, #[379] Table 8.2.5-1: Complex data type member short names
        {'shortName': 1, 'longName': 0, 'occursIn': -1, 'category': 'Trigger payload field'} #[380] Table 8.2.6 1: Trigger payload field short names
    ]
}


start = timeit.default_timer()


short_name_defs = []
for t_idx, table in enumerate(doc.tables):
    if t_idx < short_name_tables['first_idx'] or t_idx > short_name_tables['last_idx']:
        continue

    hasShortName = 0
    shortNameCol = 0
    occursInCol = 0
    for r_idx, row in enumerate(table.rows):
        if r_idx == 0:
            for c_idx, cell in enumerate(row.cells):
                if -1 != cell.text.lower().find('short name'):
                    hasShortName = 1
                    shortNameCol = c_idx
                    print('table[{}] has short names'.format(t_idx))
                if -1 != cell.text.lower().find('occurs in'):
                    occursInCol = c_idx
        if 0 == hasShortName:
            # go to next row in the same table
            break
        if r_idx > 0:
            # get short name, long name, occurs in and category info
            short_name_col = short_name_tables['column_info'][t_idx - short_name_tables['first_idx']]['shortName']
            long_name_col = short_name_tables['column_info'][t_idx - short_name_tables['first_idx']]['longName']
            occurs_in_col = short_name_tables['column_info'][t_idx - short_name_tables['first_idx']]['occursIn']
            category = short_name_tables['column_info'][t_idx - short_name_tables['first_idx']]['category']

            long_name = unidecode(row.cells[long_name_col].text).strip()
            # sometimes last row has a note, which needs to be skipped
            if long_name.find('NOTE') == 0:
                break

            # if shortName contains '*' for annotation, then remove it
            # if shortName mistakenly has upper cases, then make them lower cases
            short_name = unidecode(row.cells[short_name_col].text.replace("*", "").lower())

            if occurs_in_col == -1:
                occurs_in = '(n/a)'
            else:
                occurs_in = unidecode(row.cells[occurs_in_col].text)

            # print(shortName, longName, occursIn)
            short_name_defs.append({
                'shortName': short_name,
                'longName': long_name,
                'occursIn': occurs_in,
                'category': category
            })
    if 0 == hasShortName:
        # if there's no short name column in the first row, then go to the next table
        continue


stop = timeit.default_timer()
print('elapsed time: {}'.format(stop - start))

with open("TS-0004_" + version + "_short_name_defs.json", "w") as json_file:
    json.dump(short_name_defs, json_file, indent=4)
