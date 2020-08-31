import os
import time

# import pprint
import CBNetDB
import docx
import json

# import numpy

DO_ALL = False

ctime = time.strftime('%Y-%m-%d')
cdir = os.path.dirname(os.path.abspath(os.sys.argv[0]))
data_dir = os.path.join(cdir, 'data')

sqlite_file = 'aimmsDB.sqlite'
json_file = os.path.join(data_dir, 'current_doi_dataset.json')

with open(json_file, 'r') as F:
    parsed_data = json.load(F)

aimmsDB = CBNetDB.DBTools()
aimmsDB.connectSQLiteDB('aimmsDB.sqlite', work_dir=cdir)

sldata = aimmsDB.getColumns('publications', ['doi', 'month', 'newdata'])

new_papers = []
month_papers = []
print(len(sldata[0]))
for d in range(len(sldata[0])):
    print(sldata[0][d], sldata[1][d], sldata[2][d])
    if (
        int(sldata[1][d]) in range(1, int(time.strftime('%m')) + 1)
        and eval(sldata[2][d])
    ) or DO_ALL:
        new_papers.append(aimmsDB.getRow('publications', 'doi', sldata[0][d])[0])
    if int(sldata[1][d]) in (
        int(time.strftime('%m')),
        # int(time.strftime('%m')) - 1,
    ) and eval(sldata[2][d]):
        month_papers.append(aimmsDB.getRow('publications', 'doi', sldata[0][d])[0])

del sldata

# new_papers = []
# darray = numpy.array(sldata)
# darray[[1,0,2], :] = darray.copy()
# darray = darray.transpose()
# darray.sort()
# darray= darray.transpose()
# print(darray)
# print(darray.shape)

# print(darray[0])
# print(darray[1])
# print(darray[2])
# for d in range(darray.shape[1]):
# if (int(darray[0][d]) == int(time.strftime('%m')) and bool(darray[1][d])) or DO_ALL:
# new_papers.append(aimmsDB.getRow('publications', 'doi', darray[2][d])[0])
# del darray

aimmsDB.closeDB()

# pprint.pprint(new_papers)

Dx = docx.Document()
Dx2news = docx.Document()

_ = Dx.add_heading('AIMMS publication report for: {}'.format(ctime), level=1)
_ = Dx2news.add_heading('AIMMS publication report for: {}'.format(ctime), level=1)

cntr = 0
curr_month = None
for d in new_papers:
    # print(d)
    p0 = Dx.add_paragraph(
        '{} ({}-{})'.format(d[6][:60], d[4], d[2]), style='List Number'
    )
    if cntr == 0:
        curr_month = int(d[2])
        cntr += 1

for d in month_papers:
    # print(d)
    p0 = Dx.add_paragraph(style='List Number')
    p0.add_run('{} ({}-{})'.format(d[6][:60], d[4], d[2])).italic = True


def add_detail_list(doc, D):
    x = doc.add_paragraph(D[7], style='List Bullet')
    x = doc.add_paragraph(D[9], style='List Bullet')
    x = doc.add_paragraph(D[10], style='List Bullet')
    x = doc.add_paragraph(D[0], style='List Bullet')
    x = doc.add_paragraph('Corresponding author: {}'.format(D[8]), style='List Bullet')
    x = doc.add_paragraph(
        'Published {} (early online {})'.format(D[3], D[5]), style='List Bullet'
    )
    x = doc.add_paragraph('Processed: {}-{}'.format(D[4], D[2]), style='List Bullet')
    del x


def add_newsletter_item(doc, D):
    p = doc.add_paragraph()
    p.add_run(D[7] + ' ')
    p.add_run(D[6] + ' ').bold = True
    p.add_run('({}, {})[{}]'.format(D[10], D[3], D[0]))


_ = Dx2news.add_heading(level=3).add_run(
    'New papers: {}-{}'.format(time.strftime('%Y'), time.strftime('%m'))
)

for d in new_papers:
    # detailed doc
    h = Dx.add_heading(level=3)
    h.add_run('{}) {}'.format(cntr, d[6]))
    add_detail_list(Dx, d)
    if DO_ALL:
        p = Dx.add_paragraph(d[11].replace(d[6], '')[:301] + ' ...')
    else:
        p = Dx.add_paragraph(d[11].replace(d[6], ''))
    p.add_run(' ')

    # newsletter
    add_newsletter_item(Dx2news, d)

    cntr += 1

_ = Dx2news.add_heading(level=3).add_run('New papers: {}'.format(time.strftime('%Y')))

for d in month_papers:
    # detailed doc
    h = Dx.add_heading(level=3)
    h.add_run('{}) {}'.format(cntr, d[6])).italic = True
    add_detail_list(Dx, d)
    p = Dx.add_paragraph(d[11].replace(d[6], '')[:200] + ' ...')
    p.add_run(' ')

    # newsletter
    add_newsletter_item(Dx2news, d)

    cntr += 1

Dx.save('AIMMS_publication_report-{}.docx'.format(ctime))
Dx2news.save('AIMMS_publications_for_newsletter-{}.docx'.format(ctime))
time.sleep(2)
