__doc__ = """
This is a docstring.

"""
__version__ = 0.1

import os
import time
import operator
import numpy

import CBNetDB
import json
import xlsxwriter
import pprint

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

def makeWordcloud(fname, kword_dict, size=(12, 12), height=600, width=800):
    wcloud = wordcloud.WordCloud(height=height, width=width).generate_from_frequencies(kword_dict)
    plt.figure(figsize=size)
    plt.imshow(wcloud)
    plt.axis("off")
    plt.savefig(fname, bbox_inches='tight')
    plt.close()

def makeBarChart(wbook, wsheet, sheet_name, thetuplelist, descending=True):
    #chart1 = wbook.add_chart({'type': 'bar'})
    chart1 = wbook.add_chart({'type': 'radar', 'subtype': 'with_markers'})

    # Configure the first series.
    chart1.add_series({
        'name':       '={}!$B$1'.format(sheet_name),
        'categories': '={}!$A$2:$A${}'.format(sheet_name, len(thetuplelist)+1),
        'values':     '={}!$B$2:$B${}'.format(sheet_name, len(thetuplelist)+1),
    })

    # Add a chart title and some axis labels.
    chart1.set_title ({'name': 'Top {} {}'.format(len(thetuplelist), sheet_name)})
    chart1.set_x_axis({'name': 'Papers'})
    chart1.set_y_axis({'name': 'Organisation'})

    # Set an Excel chart style.
    chart1.set_style(11)
    chart1.set_size({'x_scale': 3, 'y_scale': 5})

    return chart1


def writeFreqToSheet(wbook, sheet_name, col_header, thedict, addbarchart):
    sheet = wbook.add_worksheet()
    sheet.name = sheet_name
    sheet.set_column('A:A', 50)

    thelist = sorted(thedict.items(), key=operator.itemgetter(1))

    sheet.write(0, 0, col_header)
    sheet.write(0, 1, 'count')
    row = 1
    outlist = []
    for i in range(len(thelist)-1, -1, -1):
        sheet.write(row, 0, thelist[i][0])
        sheet.write(row, 1, thelist[i][1])
        outlist.append((thelist[i][0], thelist[i][1]))
        row += 1
    if addbarchart > 0:
        chart1 = makeBarChart(wbook, sheet, sheet_name, outlist[:addbarchart+1])
        sheet.insert_chart('D2', chart1, {'x_offset': 25, 'y_offset': 10})

    return sheet

def filterFreq(thelist_freq, min_count_allowed, exclude_list, include_list, nasty_exclude):
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
        elif nasty_exclude:
            for k_ in exclude_list:
                if k_ in a and a in thelist_freq_filtered:
                    thelist_freq_filtered.pop(a)
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
    nasty_exclude=False,
    create_barchart=0
):

    thelist.sort()
    thelist_freq = collections.Counter(thelist)

    # filter, this probably needs some work
    if apply_filter:
        thelist_freq = filterFreq(thelist_freq, min_count_allowed, exclude_list, include_list, nasty_exclude)

    # write filtered data to sheet
    orgsheet = writeFreqToSheet(wbook, sheet_name, theheader, thelist_freq, addbarchart=create_barchart)
    if create_wordcloud:
        makeWordcloud(sheet_name + '.png', thelist_freq)
        orgsheet.insert_image('D77', sheet_name + '.png', {'x_offset': 25, 'y_offset': 10})


# ###################
# Generate reports
# ###################

max_barchart_items = 50

# create organisation reports
analysis_results = xlsxwriter.Workbook('coauthor_organisation_analysis.xlsx')
all_author_organisations = []
for p in data:
    for a in data[p]['organisations']:
        all_author_organisations.append(a)

createWordcloudSheet(
    analysis_results,
    'organisations_raw',
    all_author_organisations,
    'organisation',
    [],
    [],
    0,
    apply_filter=False,
    create_wordcloud=True,
    nasty_exclude=False,
    create_barchart=max_barchart_items
)

createWordcloudSheet(
    analysis_results,
    'organisations_cleaner',
    all_author_organisations,
    'organisation',
    [],
    FLT.org_exclude_list,
    2,
    apply_filter=True,
    create_wordcloud=True,
    nasty_exclude=False,
    create_barchart=max_barchart_items
)

createWordcloudSheet(
    analysis_results,
    'organisations_noedu',
    all_author_organisations,
    'organisation',
    [],
    FLT.org_uni_list+FLT.org_exclude_list+FLT.org_nouni_exclude_list+FLT.org_group_list,
    0,
    apply_filter=True,
    create_wordcloud=True,
    nasty_exclude=True,
    create_barchart=max_barchart_items
)

createWordcloudSheet(
    analysis_results,
    'aimms_research_groups',
    all_author_organisations,
    'group',
    FLT.org_group_list,
    FLT.org_group_exclude_list,
    0,
    apply_filter=True,
    create_wordcloud=True,
    nasty_exclude=False,
    create_barchart=max_barchart_items
)



# playing around with multi-refernces


multigroup = {}
mapped_depts = []
cross_dept_frequency_list = []
for paper in data:
    groups = []
    groups_nomap = []
    for grp in FLT.org_group_list:
        if grp in data[paper]['organisations']:
            groups_nomap.append(grp)
            groups.append(FLT.org_group_map_dept[grp])
            cross_dept_frequency_list.append(FLT.org_group_map_dept[grp])
    groups_nomap.sort()
    if len(groups) > 1:
        multigroup[paper] = {'groups' : list(set([FLT.org_group_map_dept[g] for g in groups])),
                             'groups0' : groups_nomap,
                             'groups1' : groups,
                             'contributors' : data[paper]['contributors'],
                             'title' : data[paper]['title'],
                             'year' : data[paper]['year']
                             }
cross_dept_data = {}
cross_group_data = {}
for paper in multigroup:
    if len(multigroup[paper]['groups']) > 1:
        cross_dept_data[paper] = {'groups' : multigroup[paper]['groups'],
                                  'groups0' : multigroup[paper]['groups0'],
                                  'groups1' : multigroup[paper]['groups1'],
                                  'contributors' : multigroup[paper]['contributors'],
                                  'title' : multigroup[paper]['title'],
                                  'year' : multigroup[paper]['year']
                                  }
    if len(multigroup[paper]['groups0']) > 1:
        cross_group_data[paper] = {'groups' : multigroup[paper]['groups'],
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

with open('data_multigroup_group.json', 'w') as F:
    json.dump(cross_group_data, F, indent=1)

spread_out = []
dept_combi_freq = []
grp_combi_freq = []

out_all = {'dois' : []}
out_cap_mcb = {'dois' : []}
out_eah_mcb = {'dois' : []}
out_cap_eah = {'dois' : []}
cap = 'Chemistry and Pharmaceutical Sciences'
eah = 'Environment and Health'
mcb = 'Molecular Cell Biology'

for p in cross_dept_data:
    rowdat = [p, cross_dept_data[p]['year'], cross_dept_data[p]['title']]
    grps = cross_dept_data[p]['groups']
    grps.sort()
    print(grps)
    if len(grps) == 3:
        out_all[p] = cross_dept_data[p]
        out_all['dois'].append(p)
    elif cap in grps and mcb in grps:
        out_cap_mcb[p] = cross_dept_data[p]
        out_cap_mcb['dois'].append(p)
    elif cap in grps and eah in grps:
        out_cap_eah[p] = cross_dept_data[p]
        out_cap_eah['dois'].append(p)
    elif mcb in grps and eah in grps:
        out_eah_mcb[p] = cross_dept_data[p]
        out_eah_mcb['dois'].append(p)

    dept_combi_freq.append(','.join(grps))
    spread_out.append(rowdat+groups)

with open('out_all.json', 'w') as F:
    json.dump(out_all, F, indent=1)
with open('out_cap_mcb.json', 'w') as F:
    json.dump(out_cap_mcb, F, indent=1)
with open('out_cap_eah.json', 'w') as F:
    json.dump(out_cap_eah, F, indent=1)
with open('out_eah_mcb.json', 'w') as F:
    json.dump(out_eah_mcb, F, indent=1)

for p in cross_group_data:
    rowdat = [p, cross_group_data[p]['year'], cross_group_data[p]['title']]
    grps = cross_group_data[p]['groups0']
    grps.sort()
    multi_group_group = []
    for g in grps:
        if g in FLT.org_multigroup_list:
            multi_group_group.append(g)
    multi_group_group.sort()
    if len(multi_group_group) > 1:
        grp_combi_freq.append(','.join(multi_group_group))

createWordcloudSheet(
    analysis_results,
    'aimms_multi_groups',
    grp_combi_freq,
    'group',
    FLT.org_group_list,
    FLT.org_group_exclude_list,
    0,
    apply_filter=True,
    create_wordcloud=True,
    nasty_exclude=False,
    create_barchart=max_barchart_items
)

createWordcloudSheet(
    analysis_results,
    'aimms_departments',
    cross_dept_frequency_list,
    'department',
    FLT.org_group_list,
    FLT.org_group_exclude_list,
    0,
    apply_filter=True,
    create_wordcloud=True,
    nasty_exclude=False,
    create_barchart=max_barchart_items
)




# write combinations to sheet
dept_combi_freq.sort()
dept_combi_freq = collections.Counter(dept_combi_freq)
dept_combi_sheet = analysis_results.add_worksheet()
dept_combi_sheet.name = 'aimms_multi_department'

rcntr = 0
for i in dept_combi_freq:
    grpl = i.split(',')
    for j in range(len(grpl)):
        dept_combi_sheet.write(rcntr, j, grpl[j])
    dept_combi_sheet.write(rcntr, 3, dept_combi_freq[i])
    rcntr += 1

# draw combinations as chord graph
#print(spread_out)
#print(dept_combi_freq)

from mne.viz import plot_connectivity_circle

N = 4
node_names = ['IBIVU', 'C+PS', 'MCB', 'E+H']
con = numpy.array([[numpy.nan,	numpy.nan,	3,	numpy.nan],
                    [numpy.nan,	numpy.nan,	19,	8],
                    [3,	19,	numpy.nan,	3],
                    [numpy.nan,	8,	3,	numpy.nan]])
fig, axes = plot_connectivity_circle(con, node_names, vmin=0, vmax=20, fontsize_names=9, title='Joint publications', show=False)
fig.savefig('org_multi_dept_chord.png')
dept_combi_sheet.insert_image('F1', 'org_multi_dept_chord.png', {'x_offset': 25, 'y_offset': 10})

analysis_results.close()

