__doc__ = """
This is a docstring.

"""
__version__ = 0.1

import os
import time
#import operator
import numpy

import CBNetDB
import json
import openpyxl
import pprint

#import copy
#import collections
#import wordcloud
#import matplotlib.pyplot as plt

#import datafilters as FLT

# import numpy

assert os.environ['CONDA_DEFAULT_ENV'] == 'sandbox10', 'Guess what ...'

# set current month
CURRENT_MONTH = int(time.strftime('%m'))
CURRENT_YEAR = int(time.strftime('%Y'))

ctime = time.strftime('%Y-%m-%d')
cdir = os.path.dirname(os.path.abspath(os.sys.argv[0]))

# Test database
DB_FILE_NAME = 'aimms_equipment.sqlite'
DB_ACTIVE_TABLE = 'equipment'
# Production DB
DB_FILE_NAME = 'aimms_equipment-test.sqlite'
DB_ACTIVE_TABLE = 'equipment'


# OPen database
aimmsDB = CBNetDB.DBTools()
aimmsDB.connectSQLiteDB(DB_FILE_NAME, work_dir=cdir)

#sldata = aimmsDB.getColumns(DB_ACTIVE_TABLE, ['doi', 'year', 'title', 'contributors', 'organisations', 'corresponding'],)

#data = {}
#print('SLdata', len(sldata[0]))
#for r_ in range(len(sldata[0])):
    #data[aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][0]] = {
        #'year': aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][4].strip(),
        #'title': aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][6].strip(),
        #'contributors': [a.strip() for a in aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][7].replace('.,', '.|').split('|')],
        #'corresponding': aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][8].strip(),
        #'organisations': [a.strip() for a in aimmsDB.getRow(DB_ACTIVE_TABLE, 'doi', sldata[0][r_])[0][9].split(',')],
    #}
#aimmsDB.closeDB()

#with open('data0.json', 'w') as F:
    #json.dump(data, F, indent=1)

