__doc__ = """
## Generating a poster program from formdesk output

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
__version__ = 0.2

data_file = 'AIMMSannualmeeting2021_exportforBrett.xlsx'

DB_FILE_NAME = 'aimmsday_posters.sqlite'
# DATA_COLUMNS = [
# "_fd_id",
# "_fd_seqno",
# "_fd_source",
# "_fd_status",
# "_fd_add",
# "_fd_edit",
# "first_name",
# "prefix",
# "last_name",
# "aimms_research_group",
# "job_title_1",
# "job_title1",
# "organization",
# "strees_address",
# "street_address2",
# "city",
# "zip_code",
# "country",
# "e_mail",
# "do_you_want_to_present_a_",
# "title",
# "authors",
# "short_abstract",
# "i_would_like_to_receive_a",
# ]

import os
import time
import pprint
import json
import CBNetDB
import openpyxl
import docx

assert os.environ['CONDA_DEFAULT_ENV'] == 'sandbox', 'Guess what ...'

# this sets the current year and month
CURRENT_YEAR = int(time.strftime('%Y'))
CURRENT_MONTH = int(time.strftime('%m'))

# set up current env.
ctime = time.strftime('%Y-%m-%d')
cdir = os.path.dirname(os.path.abspath(os.sys.argv[0]))
data_dir = os.path.join(cdir, '..', 'noupload', 'parse_formdesk_output', 'data')
out_dir = os.path.join(cdir, '..', 'noupload', 'parse_formdesk_output')
data_file = os.path.join(data_dir, data_file)


# set up database env.
DB_FILE_NAME = os.path.join(out_dir, DB_FILE_NAME)
DB_ACTIVE_TABLE = '\"{}\"'.format(CURRENT_YEAR)
DB_COLS = [
    'pid INT PRIMARY KEY',
    'status INT',
    'fname TEXT',
    'mname TEXT',
    'lname TEXT',
    'grp TEXT',
    'job1 TEXT',
    'job2 TEXT',
    'email TEXT',
    'poster INT',
    'title TEXT',
    'author TEXT',
    'abstract TEXT',
    'borrel INT',
    'pnumber TEXT',
    'participant INT',
]

posterDB = None


"""
json_db_file = os.path.join(out_dir, json_db_file)
# configure persistent data
if os.path.exists(json_db_file):
    F = open(json_db_file, 'r')
    DB_DATA = json.load(F)
else:
    DB_DATA = {}
    F = open(json_db_file, 'w')
    DB_DATA = json.dump(DB_DATA, F)
F.close()
"""

# open database
posterDB = CBNetDB.DBTools()
posterDB.connectSQLiteDB(DB_FILE_NAME, work_dir=out_dir)
if posterDB.getTable(DB_ACTIVE_TABLE) is None:
    posterDB.createDBTable(DB_ACTIVE_TABLE, DB_COLS)

# use excel file directly
print(data_file)
exl_wb = openpyxl.load_workbook(filename=data_file)
exl_sh = exl_wb[exl_wb.sheetnames[0]]

# get max numbers
pRegMax = pNumMax = len(posterDB.getColumns(DB_ACTIVE_TABLE, ['pid'])[0])
pNumMax = posterDB.getColumns(DB_ACTIVE_TABLE, ['poster'])[0].count('1')
print(pRegMax, pNumMax)
for row in range(2, exl_sh.max_row + 1):
    pid = int(exl_sh['B{}'.format(row)].value)
    if not posterDB.checkEntryInColumn(DB_ACTIVE_TABLE, 'pid', pid):
        pRegMax += 1
        dta = {
            'pid': pid,
            'status': int(
                True if exl_sh['D{}'.format(row)].value == 'Completed' else False
            ),
            'fname': exl_sh['G{}'.format(row)].value,
            'mname': exl_sh['H{}'.format(row)].value,
            'lname': exl_sh['I{}'.format(row)].value,
            'grp': exl_sh['J{}'.format(row)].value,
            'job1': exl_sh['K{}'.format(row)].value,
            'job2': exl_sh['L{}'.format(row)].value,
            'email': exl_sh['S{}'.format(row)].value,
            'poster': int(exl_sh['T{}'.format(row)].value),
            'title': exl_sh['U{}'.format(row)].value,
            'author': exl_sh['V{}'.format(row)].value,
            'abstract': exl_sh['W{}'.format(row)].value,
            'borrel': int(exl_sh['X{}'.format(row)].value),
            'participant': pRegMax,
        }
        if dta['poster']:
            pNumMax += 1
            dta['pnumber'] = 'P{0:03d}'.format(pNumMax)
        posterDB.insertData(DB_ACTIVE_TABLE, dta, commit=False)
        print('Adding row with ID \"{}\".'.format(pid))
        # pprint.pprint(dta)
    else:
        print('Skipping existing ID \"{}\".'.format(pid))


posterDB.commitDB()
posterDB.closeDB()
