__doc__ = """
## Generating AIMMS manuscript reports

There is a two stage process to create the AIMMS paper report

### Generate a custom report in PURE
You need to create an Excel report in PURE. To do this you need elevated user rights to the AIMMS contex and the correct report format. Get this from  Brett).

### Parse data and update database
This code use a sql database to store publication reports. There are two things that are important here, the first being the normal run and the second database maintenance.

#### Normal operation
This is what you do on a regular basis, first copy the Excel spreadsheet you generated from PURE into the `\data` directory. Next add a line of code that loads that file.

```python
data_file = 'AIMMS_research_2021-8_02_21.xls'
```

Save the file and run this script from a terminal
```bash
python aimms_monthly_articles.py
````

This will update the database and JSON index files.

#### Database update (each new year)
The database *aimmsDB.sqlite* is set up to use the table *publications* for the current years results. At the begining of a calendar year simply rename
the *publications* table to the previous year, for example, 2020. In this way a history of AIMMS publications is maintained.

### Report generation

Next you can generate the report using

```bash
python generate_publication_report.py
```
This will result in two Word documents:

```bash
AIMMS_publication_report-2021-02-08.docx
AIMMS_publications_for_newsletter-2021-02-08.docx
```
A detailed list of new publications and a shortened version suitable for publication in the newsletter.


### Author
Author: Brett G. Olivier (b.g.olivier@vu.nl)
Licence: GNU GPL v3

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>

"""
__version__ = 1.5

# data_file = 'AIMMS_research_2020-7_08_20.xls'
# data_file = 'AIMMS_research_2020-31_08_20.xls'
# data_file = 'AIMMS_research_2020-14_09_20.xls'
# data_file = 'AIMMS_research_2020-28_09_20.xls'
# data_file = 'AIMMS_research_2020-12_10_20.xls'
# data_file = 'AIMMS_research_2020-26_10_20.xls'
# data_file = 'AIMMS_research_2020-9_11_20.xls'
# data_file = 'AIMMS_research_2020-29_11_20.xls'
# data_file = 'AIMMS_research_2021-11_01_21.xls'
# data_file = 'AIMMS_research_2021-25_01_21.xls'
# data_file = 'AIMMS_research_2021-8_02_21.xls'
# data_file = 'AIMMS_research_2021-22_02_21.xls'
# data_file = 'AIMMS_research_2021-8_03_21.xls'
# data_file = 'AIMMS_research_2021-22_03_21.xls'
# data_file = 'AIMMS_research_2021-31_05_21.xls'
# data_file = 'AIMMS_research_2021-14_06_21.xls'
# data_file = 'AIMMS_research_2021-28_06_21.xls'
# data_file = 'AIMMS_research_2021-13_09_21.xls'
#data_file = 'AIMMS_research_2021-20_10_21.xls'
#data_file = 'AIMMS_research_2021-1_12_21.xls'
data_file = 'AIMMS_research_2022-19_01_22.xls'

import os
import time
import pprint
import re
import json
import CBNetDB
import xlrd

assert os.environ['CONDA_DEFAULT_ENV'] == 'sandbox', 'Guess what ...'

# this sets the current year and month to search until
CURRENT_YEAR = int(time.strftime('%Y'))
CURRENT_MONTH = int(time.strftime('%m'))
# manual override for specific month/year
# CURRENT_YEAR = 2022
# CURRENT_MONTH = 12

# set up current env.
ctime = time.strftime('%Y-%m-%d')
cdir = os.path.dirname(os.path.abspath(os.sys.argv[0]))
data_dir = os.path.join(cdir, 'data')

# set up database env.
# default DB name: 'aimmsDB.sqlite'
DB_FILE_NAME = 'aimmsDB.sqlite'
# default table name: 'publications'
DB_ACTIVE_TABLE = 'Y2020'
DB_ACTIVE_TABLE = 'publications'

# using excel file directly
data = {}
Bk = xlrd.open_workbook(os.path.join(data_dir, data_file))
St = Bk.sheet_by_index(0)
month = St.col(0)
paper = St.col(1)
for r in range(len(paper)):
    try:
        if str(int(month[r].value)) in data:
            data[str(int(month[r].value))].append(str(paper[r].value))
        else:
            data[str(int(month[r].value))] = [str(paper[r].value)]
    except:
        pass

## using exported tsv file
# data_file = 'AIMMS_research_2020-3_08_20.txt'
# data = {}
# with open(os.path.join(data_dir, data_file), 'r') as F:
##csvread = csv.reader(F, dialect="excel-tab")
# for r in F:
# l = r.split('\t')
# if len(l) < 2:
# continue
# if l[0] in data:
# data[l[0]].append(l[1])
# else:
# data[l[0]] = [l[1]]
# pprint.pprint(data)

re_pub_data = re.compile(r"Publication date:(.*?)Handle")
re_early_date = re.compile(r"Early online date:(.*?)Publication")
re_contrib = re.compile(r"Contributors:(.*?)(?:Number|Publication)")
re_corresp = re.compile(r"Corresponding author:(.*?)Contributors")
re_journal = re.compile(r"Journal:(.*?)(Volume|ISSN|Original)")
# need to make the URL parsing deal with PURE duplicate DOI's.
re_doi = re.compile(r"DOIs:(.*?)URLs")
re_orgs = re.compile(r"Organisations:(.*?)(?:Contributors|Corresponding)")
re_title = re.compile(r"^(.*?[a-z\?])[A-Z]")
re_descript = re.compile(r"^(.*?)General information")

parsed_data = {}
parsed_data_nodoi = {}
data_keys = []
for m in [str(a + 1) for a in range(CURRENT_MONTH)]:
    cntr = 1
    for p in data[m]:
        NODOI = False
        pclean = p.replace('\xa0', '')
        pub_date_match = re_pub_data.search(pclean)
        early_date_match = re_early_date.search(pclean)
        contrib_match = re_contrib.search(pclean)
        corresp_match = re_corresp.search(pclean)
        journal_match = re_journal.search(pclean)
        doi_match = re_doi.search(pclean)
        orgs_match = re_orgs.search(pclean)
        title_match = re_title.search(pclean)
        descript_match = re_descript.search(pclean)

        print('\nProcessing record: {} ...'.format(cntr))

        if journal_match is None and 'Research output: PhD Thesis' in p:
            print('PHD thesis --> NODOI')
            NODOI = True
        elif journal_match is None:
            print('No journal: \"{}\"'.format(p[:40]))

        if doi_match is not None:
            key = doi_match.groups()[0].strip()
            if 'https://doi.org/' in key:
                print('URLkey: {}'.format(key))
                if key.count(key[-4:]) > 1:
                    key = key.split(key[-4:])[0] + key[-4:]
                    print('FIXkey', key)
            else:
                key = 'https://doi.org/' + key
                print('DOIkey: {}'.format(key))
        else:
            key = str(time.time())
            print('NOkey', key)
            NODOI = True

        parsed_data[key] = {}
        parsed_data[key]['month'] = m

        if pub_date_match is not None:
            # print(m, cntr, pub_date_match.groups()[0])
            parsed_data[key]['pub_date'] = pub_date_match.groups()[0]
            parsed_data[key]['year'] = int(parsed_data[key]['pub_date'].split(' ')[-1])
        else:
            parsed_data[key]['year'] = CURRENT_YEAR
        if early_date_match is not None:
            # print(m, cntr, title_match.groups()[0])
            parsed_data[key]['early_online'] = early_date_match.groups()[0]
        else:
            parsed_data[key]['early_online'] = None
        if title_match is not None:
            # print(m, cntr, title_match.groups()[0])
            parsed_data[key]['title'] = title_match.groups()[0]
        else:
            parsed_data[key]['title'] = None
        if contrib_match is not None:
            # print(m, cntr, contrib_match.groups()[0])
            parsed_data[key]['contributors'] = contrib_match.groups()[0]
        else:
            parsed_data[key]['contributors'] = None
        if corresp_match is not None:
            # print(m, cntr, corresp_match.groups()[0])
            parsed_data[key]['corresponding'] = corresp_match.groups()[0]
        else:
            parsed_data[key]['corresponding'] = None
        if orgs_match is not None:
            # print(m, cntr, orgs_match.groups()[0])
            parsed_data[key]['organisations'] = orgs_match.groups()[0]
        else:
            parsed_data[key]['organisations'] = None
        if journal_match is not None:
            # print(m, cntr, journal_match.groups()[0])
            parsed_data[key]['journal'] = journal_match.groups()[0]
        else:
            parsed_data[key]['journal'] = None
        if descript_match is not None:
            # print(m, cntr, descript_match.groups()[0])
            parsed_data[key]['description'] = descript_match.groups()[0]
        else:
            parsed_data[key]['description'] = None
        data_keys = tuple(parsed_data[key].keys())
        cntr += 1
        if NODOI:
            parsed_data_nodoi[key] = parsed_data.pop(key)

# pprint.pprint(parsed_data)

# os.sys.exit(1)

pprint.pprint(parsed_data_nodoi)
for k in parsed_data:
    print(k)

with open(os.path.join(data_dir, 'current_doi_dataset.json'), 'w') as F:
    json.dump(parsed_data, F, indent=1)
with open(os.path.join(data_dir, 'current_nodoi_dataset.json'), 'w') as F:
    json.dump(parsed_data_nodoi, F, indent=1)

# setup/connect/load database
table_cols = ['doi TEXT PRIMARY KEY', 'newdata TEXT']
_ = [table_cols.append('{} TEXT'.format(a)) for a in data_keys]
print(table_cols)

aimmsDB = CBNetDB.DBTools()
aimmsDB.connectSQLiteDB(DB_FILE_NAME, work_dir=cdir)
test = aimmsDB.getTable(DB_ACTIVE_TABLE)
if aimmsDB.getTable(DB_ACTIVE_TABLE) is None:
    aimmsDB.createDBTable(DB_ACTIVE_TABLE, table_cols)
table_pub = aimmsDB.getTable(DB_ACTIVE_TABLE)

new_data = []
for doi in parsed_data:
    if len(aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', doi)) == 0:
        dta = parsed_data[doi].copy()
        dta['doi'] = doi
        dta['newdata'] = 'True'
        aimmsDB.insertData(DB_ACTIVE_TABLE, dta, commit=False)
        new_data.append(doi)
        print('New data inserted for DOI: {}'.format(doi))
    else:
        aimmsDB.updateData(
            DB_ACTIVE_TABLE, 'doi', doi, {'newdata': 'False'}, commit=False
        )
aimmsDB.commitDB()
aimmsDB.closeDB()
