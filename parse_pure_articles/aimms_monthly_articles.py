# data_file = 'AIMMS_research_2020-7_08_20.xls'
# data_file = 'AIMMS_research_2020-31_08_20.xls'
# data_file = 'AIMMS_research_2020-14_09_20.xls'
data_file = 'AIMMS_research_2020-28_09_20.xls'

import os
import time
import pprint
import re
import json
import CBNetDB
import xlrd

ctime = time.strftime('%Y-%m-%d')
cdir = os.path.dirname(os.path.abspath(os.sys.argv[0]))
data_dir = os.path.join(cdir, 'data')

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
re_journal = re.compile(r"Journal:(.*?)Volume")
re_doi = re.compile(r"DOIs:(.*?)URLs")
re_orgs = re.compile(r"Organisations:(.*?)(?:Contributors|Corresponding)")
re_title = re.compile(r"^(.*?[a-z\?])[A-Z]")
re_descript = re.compile(r"^(.*?)General information")

parsed_data = {}
parsed_data_nodoi = {}
data_keys = []
for m in [str(a + 1) for a in range(int(time.strftime('%m')))]:
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

        print('Processing record: {} ...'.format(cntr))
        if doi_match is not None:
            # print(m, cntr, 'https://doi.org/'+doi_match.groups()[0])
            key = 'https://doi.org/' + doi_match.groups()[0]
        else:
            key = str(time.monotonic()).split('.')[0]
            NODOI = True

        parsed_data[key] = {}
        parsed_data[key]['month'] = m

        if pub_date_match is not None:
            # print(m, cntr, pub_date_match.groups()[0])
            parsed_data[key]['pub_date'] = pub_date_match.groups()[0]
            parsed_data[key]['year'] = int(parsed_data[key]['pub_date'].split(' ')[-1])
        else:
            parsed_data[key]['year'] = time.strftime('%Y')
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

pprint.pprint(parsed_data)
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
aimmsDB.connectSQLiteDB('aimmsDB.sqlite', work_dir=cdir)
test = aimmsDB.getTable('publications')
if aimmsDB.getTable('publications') is None:
    aimmsDB.createDBTable('publications', table_cols)
table_pub = aimmsDB.getTable('publications')

new_data = []
for doi in parsed_data:
    if len(aimmsDB.getRow('publications', 'doi', doi)) == 0:
        dta = parsed_data[doi].copy()
        dta['doi'] = doi
        dta['newdata'] = 'True'
        aimmsDB.insertData('publications', dta, commit=False)
        new_data.append(doi)
        print('New data inserted for DOI: {}'.format(doi))
    else:
        aimmsDB.updateData(
            'publications', 'doi', doi, {'newdata': 'False'}, commit=False
        )
aimmsDB.commitDB()
aimmsDB.closeDB()
