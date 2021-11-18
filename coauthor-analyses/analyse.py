__doc__ = """
This is a docstring.

"""
__version__ = 0.1

import os
import time

import CBNetDB
import json
import xlsxwriter

import copy
import collections
import wordcloud
import matplotlib.pyplot as plt


# import numpy

assert os.environ['CONDA_DEFAULT_ENV'] == 'sandbox', 'Guess what ...'

# set current month
CURRENT_MONTH = int(time.strftime('%m'))
CURRENT_YEAR = int(time.strftime('%Y'))

ctime = time.strftime('%Y-%m-%d')
cdir = os.path.dirname(os.path.abspath(os.sys.argv[0]))

# Test database
DB_FILE_NAME = 'aimmsDBtest.sqlite'
DB_ACTIVE_TABLE = 'Y2020'
# Production DB
# DB_FILE_NAME = 'aimmsDB6yrs.sqlite'
# DB_ACTIVE_TABLE = 'publications'


# Grab data from database and create a json dict and dump to file.
aimmsDB = CBNetDB.DBTools()
aimmsDB.connectSQLiteDB(DB_FILE_NAME, work_dir=cdir)

sldata = aimmsDB.getColumns(
    DB_ACTIVE_TABLE,
    ['doi', 'year', 'title', 'contributors', 'organisations', 'corresponding'],
)

# print(sldata)

data = {}
print('SLdata', len(sldata[0]))
for r_ in range(len(sldata[0])):
    data[aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][0]] = {
        'year': aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][4].strip(),
        'title': aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][6].strip(),
        'contributors': [a.strip() for a in aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][7].replace('.,', '.|').split('|')],
        'corresponding': aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][8].strip(),
        'organisations': [a.strip() for a in aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][9].split(',')],
    }
aimmsDB.closeDB()

with open('data0.json', 'w') as F:
    json.dump(data, F, indent=1)


# Here we could implement a more intelligent junk filter
# IAMAJUNKFILTER

# Lets play with our new data


def makeWordcloud(fname, kword_dict, size=(30, 16), height=1024, width=1280):
    wcloud = wordcloud.WordCloud(height=height, width=width).generate_from_frequencies(kword_dict)
    plt.figure(figsize=size)
    plt.imshow(wcloud)
    plt.axis("off")
    plt.savefig(fname, bbox_inches='tight')
    plt.close()


def createWordcloudSheet(
    wbook,
    sheet_name,
    thelist,
    theheader,
    include_list,
    exclude_list,
    min_count_allowed,
    apply_filter=True,
    create_wordcloud=True,
):

    thelist.sort()
    thelist_freq = collections.Counter(thelist)

    # filter, this probably needs some work
    out_dict_included = {}
    thelist_freq_filtered = copy.deepcopy(thelist_freq)
    if apply_filter:
        for a in tuple(thelist_freq_filtered.keys()):
            for k_ in include_list:
                if k_ in a and a in thelist_freq_filtered:
                    out_dict_included[a] = thelist_freq_filtered.pop(a)
            if (a in exclude_list or thelist_freq_filtered[a] < min_count_allowed) and a in thelist_freq_filtered:
                thelist_freq_filtered.pop(a)
        if len(include_list) > 0:
            thelist_freq_filtered = out_dict_included

    # write filtered data to sheet
    orgsheet = wbook.add_worksheet()
    orgsheet.name = sheet_name

    orgsheet.write(0, 0, theheader)
    orgsheet.write(0, 1, 'count')
    row = 1
    for i in thelist_freq_filtered:
        orgsheet.write(row, 0, i)
        orgsheet.write(row, 1, thelist_freq_filtered[i])
        row += 1

    if create_wordcloud:
        makeWordcloud(sheet_name + '.png', thelist_freq_filtered)
        orgsheet.insert_image('D2', sheet_name + '.png')


# Generate reports
analysis_results = xlsxwriter.Workbook('analysis_results.xlsx')

include_list = []
exclude_list = []

# create organisation report

all_author_organisations = []
for p in data:
    for a in data[p]['organisations']:
        all_author_organisations.append(a)

include_list = ['University']

exclude_list = [
    'AIMMS',
    'VU University',
    'The Netherlands',
    'The Netherlands.',
    'Amsterdam',
    'Vrije Universiteit Amsterdam',
    'Utrecht',
    'Spain',
    'Spain.',
    'Inc.',
    'UK.',
    'Denmark.',
    'Germany.',
    'Sweden.',
    'Barcelona',
    'Amsterdam Institute for Molecules',
]


createWordcloudSheet(
    analysis_results,
    'author_organisations_all',
    all_author_organisations,
    'organisation',
    [],
    [],
    0,
    apply_filter=False,
    create_wordcloud=True,
)

createWordcloudSheet(
    analysis_results,
    'author_organisations_cleaned',
    all_author_organisations,
    'organisation',
    [],
    exclude_list,
    2,
    apply_filter=True,
    create_wordcloud=True,
)


createWordcloudSheet(
    analysis_results,
    'author_organisations_uni',
    all_author_organisations,
    'organisation',
    include_list,
    exclude_list,
    0,
    apply_filter=True,
    create_wordcloud=True,
)

analysis_results.close()
