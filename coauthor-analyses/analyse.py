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

import datafilters as FLT

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
DB_FILE_NAME = 'aimmsDB6yrs.sqlite'
DB_ACTIVE_TABLE = 'publications'


# Grab data from database and create a json dict and dump to file.
aimmsDB = CBNetDB.DBTools()
aimmsDB.connectSQLiteDB(DB_FILE_NAME, work_dir=cdir)

sldata = aimmsDB.getColumns(DB_ACTIVE_TABLE, ['doi', 'year', 'title', 'contributors', 'organisations', 'corresponding'],)

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

def makeWordcloud(fname, kword_dict, size=(20, 10), height=1024, width=1280):
    wcloud = wordcloud.WordCloud(height=height, width=width).generate_from_frequencies(kword_dict)
    plt.figure(figsize=size)
    plt.imshow(wcloud)
    plt.axis("off")
    plt.savefig(fname, bbox_inches='tight')
    plt.close()

def writeFreqToSheet(wbook, sheet_name, col_header, thelist):
    sheet = wbook.add_worksheet()
    sheet.name = sheet_name

    sheet.write(0, 0, col_header)
    sheet.write(0, 1, 'count')
    row = 1
    for i in thelist:
        sheet.write(row, 0, i)
        sheet.write(row, 1, thelist[i])
        row += 1
    return sheet

def filterFreq(thelist_freq, min_count_allowed, exclude_list, include_list):
    """
    If include list is not empty then exclude list is used as a filter
    """
    out_dict_included = {}
    thelist_freq_filtered = copy.deepcopy(thelist_freq)
    include_mode = False
    if len(include_list) > 0:
        include_mode = True

    for a in tuple(thelist_freq_filtered.keys()):
        if include_mode:
            for k_ in include_list:
                if k_ in a and a in thelist_freq_filtered and a not in exclude_list:
                    out_dict_included[a] = thelist_freq_filtered.pop(a)
        else:
            if (a in exclude_list or thelist_freq_filtered[a] < min_count_allowed) and a in thelist_freq_filtered:
                thelist_freq_filtered.pop(a)
    if include_mode:
        thelist_freq_filtered = out_dict_included
    return thelist_freq_filtered

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
    if apply_filter:
        thelist_freq = filterFreq(thelist_freq, min_count_allowed, exclude_list, include_list)

    # write filtered data to sheet
    orgsheet = writeFreqToSheet(wbook, sheet_name, theheader, thelist_freq)
    if create_wordcloud:
        makeWordcloud(sheet_name + '.png', thelist_freq)
        orgsheet.insert_image('D2', sheet_name + '.png')

# ###################
# Generate reports
# ###################

# create organisation reports
analysis_results = xlsxwriter.Workbook('organisation_analysis.xlsx')
all_author_organisations = []
for p in data:
    for a in data[p]['organisations']:
        all_author_organisations.append(a)

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
    FLT.org_exclude_list,
    2,
    apply_filter=True,
    create_wordcloud=True,
)


createWordcloudSheet(
    analysis_results,
    'author_organisations_uni',
    all_author_organisations,
    'organisation',
    FLT.org_uni_list,
    ['VU University', 'Vrije Universiteit', 'Vrije Universiteit Amsterdam'],
    0,
    apply_filter=True,
    create_wordcloud=True,
)

createWordcloudSheet(
    analysis_results,
    'author_organisations_groups',
    all_author_organisations,
    'organisation',
    FLT.org_group_list2,
    FLT.org_group_exclude_list,
    0,
    apply_filter=True,
    create_wordcloud=True,
)



# playing around with multi-refernces
import pprint

multigroup = {}
mapped_depts = []
cross_dept_frequency_list = []
for paper in data:
    groups = []
    groups_nomap = []
    for grp in FLT.org_group_list:
        if grp in data[paper]['organisations']:
            groups_nomap.append(grp)
            groups.append(FLT.org_group_map_dept2[grp])
            cross_dept_frequency_list.append(FLT.org_group_map_dept2[grp])
    if len(groups) > 1:
        multigroup[paper] = {'groups' : list(set([FLT.org_group_map_dept2[g] for g in groups])),
                             'groups0' : groups_nomap,
                             'groups1' : groups,
                             'contributors' : data[paper]['contributors'],
                             'title' : data[paper]['title'],
                             'year' : data[paper]['year']
                             }
cross_dept_data = {}
for paper in multigroup:
    if len(multigroup[paper]['groups']) > 1:
        cross_dept_data[paper] = {'groups' : multigroup[paper]['groups'],
                                  'groups0' : multigroup[paper]['groups0'],
                                  'groups1' : multigroup[paper]['groups1'],
                                  'contributors' : multigroup[paper]['contributors'],
                                  'title' : multigroup[paper]['title'],
                                  'year' : multigroup[paper]['year']
                                  }

with open('data_multigroup_raw.json', 'w') as F:
    json.dump(multigroup, F, indent=1)

with open('data_multigroup.json', 'w') as F:
    json.dump(cross_dept_data, F, indent=1)

spread_out = []
dept_combi_freq = []
for p in cross_dept_data:
    rowdat = [p, cross_dept_data[p]['year'], cross_dept_data[p]['title']]
    grps = cross_dept_data[p]['groups']
    grps.sort()
    dept_combi_freq.append(','.join(grps))
    spread_out.append(rowdat+groups)

#print(spread_out)
print(dept_combi_freq)

createWordcloudSheet(
    analysis_results,
    'author_organisations_dept_freq',
    cross_dept_frequency_list,
    'organisation',
    FLT.org_group_list,
    FLT.org_group_exclude_list,
    0,
    apply_filter=True,
    create_wordcloud=True,
)

#createWordcloudSheet(
    #analysis_results,
    #'dept_combi_freq',
    #dept_combi_freq,
    #'organisation',
    #[],
    #[],
    #0,
    #apply_filter=True,
    #create_wordcloud=False,
#)


# write combinations to sheet
dept_combi_freq.sort()
dept_combi_freq = collections.Counter(dept_combi_freq)
dept_combi_sheet = analysis_results.add_worksheet()
dept_combi_sheet.name = 'dept_combi_freq'

rcntr = 0
for i in dept_combi_freq:
    print(i)
    grpl = i.split(',')
    print(grpl)
    for j in range(len(grpl)):
        dept_combi_sheet.write(rcntr, j, grpl[j])
    dept_combi_sheet.write(rcntr, 3, dept_combi_freq[i])
    rcntr += 1

analysis_results.close()

#pprint.pprint(cross_dept_data)

#import numpy
#import networkx as nx
#import matplotlib.pyplot as plt

#dept_interactions = [[0,14,6,0],
                     #[14,0,1,10],
                     #[6,1,0,0],
                     #[0,10,0,0]]

#x = [0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3]
#y = [0,1,2,3,0,1,2,3,0,1,2,3,0,1,2,3]
#s = [0,14,6,0,14,0,1,10,6,1,0,0,0,10,0,0]
#colors = numpy.random.rand(len(x))


#plt.scatter(x, y, s=[a*50 for a in s], c=colors, alpha=0.5)
#plt.show()
#plt.savefig('dept_interaction_graph.png')

