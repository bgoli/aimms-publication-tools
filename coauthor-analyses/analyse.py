__doc__ = """
This is a docstring.

"""
__version__ = 0.1

import os
import time

import CBNetDB
import json
import xlsxwriter


# import numpy

assert os.environ['CONDA_DEFAULT_ENV'] == 'sandbox', 'Guess what ...'

# set current month
CURRENT_MONTH = int(time.strftime('%m'))
CURRENT_YEAR = int(time.strftime('%Y'))

ctime = time.strftime('%Y-%m-%d')
cdir = os.path.dirname(os.path.abspath(os.sys.argv[0]))

# set up database env.
DB_FILE_NAME = 'aimmsDB.sqlite'
# default table name: 'publications'
DB_ACTIVE_TABLE = 'Y2020'

# Grab data from database and create a json dict and dump to file.
aimmsDB = CBNetDB.DBTools()
aimmsDB.connectSQLiteDB(DB_FILE_NAME, work_dir=cdir)

sldata = aimmsDB.getColumns(
    DB_ACTIVE_TABLE, ['doi', 'year', 'title', 'contributors', 'organisations', 'corresponding']
)

# print(sldata)

data = {}
print('SLdata', len(sldata[0]))
for r_ in range(len(sldata[0])):
    data[aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][0]] = {
        'year' : aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][4].strip(),
        'title' : aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][6].strip(),
        'contributors' : [a.strip() for a in aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][7].replace('.,', '.|').split('|')],
        'corresponding' : aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][8].strip(),
        'organisations' : [a.strip() for a in aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][9].split(',')],
    }
aimmsDB.closeDB()

with open('data0.json', 'w') as F:
    json.dump(data, F, indent=1)


# Lets play with our new data
import copy
import collections
import wordcloud
import matplotlib.pyplot as plt

analysis_results = xlsxwriter.Workbook('analysis_results.xlsx')
orgsheet = analysis_results.add_worksheet()
orgsheet.name = 'organisations'

# extract all author organisations into sheet 1
all_author_organisations = []

for p in data:
    for a in data[p]['organisations']:
        all_author_organisations.append(a)
all_author_organisations.sort()
all_author_organisations_freq = collections.Counter(all_author_organisations)

orgsheet.write(0, 0, 'organisation')
orgsheet.write(0, 1, 'count')
row = 1
for i in all_author_organisations_freq:
    orgsheet.write(row, 0, i)
    orgsheet.write(row, 1, all_author_organisations_freq[i])
    row += 1



def MakeWordcloud(fname, kword_dict):
    wcloud = wordcloud.WordCloud(height = 1024, width = 1280).generate_from_frequencies(kword_dict)
    plt.figure(figsize=(30, 16))
    plt.imshow(wcloud)
    plt.axis("off")
    plt.savefig(fname, bbox_inches='tight')
    plt.close()

# create a wordcloud from organisation list
min_count_allowed = 2
exclude_list = ['AIMMS', 'VU University', 'The Netherlands', 'The Netherlands.', 'Amsterdam', 'Vrije Universiteit Amsterdam', 'Utrecht', 'Spain', 'Spain.',
                'Inc.', 'UK.', 'Denmark.', 'Germany.', 'Sweden.', 'Barcelona', 'Amsterdam Institute for Molecules']

uni_dict = {}
kword_dict = copy.deepcopy(all_author_organisations_freq)
for a in tuple(kword_dict.keys()):
    if 'University' in a:
        uni_dict[a] = kword_dict.pop(a)
    elif kword_dict[a] < min_count_allowed or a in exclude_list:
        kword_dict.pop(a)

MakeWordcloud('author_organisations_all.png', kword_dict)
orgsheet.insert_image('D2', 'author_organisations_all.png')

MakeWordcloud('author_organisations_uni.png', uni_dict)
orgsheet.insert_image('D62', 'author_organisations_uni.png')


analysis_results.close()