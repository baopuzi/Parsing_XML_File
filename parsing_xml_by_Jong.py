# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 2019
@author: Jong from Malaysia
"""

import sys, os, glob, itertools, csv, time, openpyxl, mmap
from openpyxl import Workbook

start = time.time()
fpath = os.path.expanduser(os.sep.join(["~", "Desktop\DV"]))
os.chdir(fpath)


# print 'new current path=',os.getcwd()

def readNext(keyWord, string):
    for i, j in enumerate(string):
        if j == keyWord:
            return string[i + 1]


i = 0
lista = []
listb = []
listc = []
listd = []
liste = []
listf = []
listg = []
listTitle = ['Subnet ID', 'NE ID', 'Version', 'Parameter', 'Value']
for files in glob.glob("*.xml"):
    with open(files) as f:
        for setLine in itertools.islice(f, 0, 500000):
            if 'ucQosMaxSelNumAdjustOffset' in setLine:
                with open(files) as f2:
                    line = f2.readlines(12)
                    if len(line[1]) == 1:
                        version = line[2]
                    else:
                        version = line[1]
                    versionValue = version[31:48]
                    # versionValue=version
                    # print versionValue
                    parameter = 'ucQosMaxSelNumAdjustOffset'
                    newSetLine = setLine.split("\"")
                    dValue = readNext(' value=', newSetLine)
                    subnet = files.split('_')[0]
                    y = subnet.replace("subnet", "")
                    reTrieveNEID = files[-16:-11]
                    temp = reTrieveNEID
                    # testing1=y+'_'+temp
                    # testing2=testing1+'_'+versionValue+'_'+dValue
                    liste.append(parameter)
                    listd.append(y)
                    lista.append(temp)
                    # listf.append(testing1)
                    listb.append(versionValue)
                    listc.append(dValue)
                    # listg.append(testing2)
                    # print (versionValue+" "+dValue)
                    i += 1
                    print i
print 'Total file(s) : %d' % i
print 'took ' + str(time.time() - start)

timeFormat = time.strftime("%Y%m%d-%H%M%S")
filename = 'result_' + timeFormat
xl = openpyxl.Workbook()
xs = xl.active

with open(filename + '.csv', 'wb') as cf:
    filewriter = csv.writer(cf, quoting=csv.QUOTE_ALL)
    filewriter.writerow(listTitle)
    # filewriter.writerows(zip(listd,lista,listf,listb,liste,listc,listg))
    # filewriter.writerows(zip(listf,listg))
    filewriter.writerows(zip(listd, lista, listb, liste, listc))
with open(filename + '.csv', 'rb') as rf:
    filereader = csv.reader(rf, quoting=csv.QUOTE_NONNUMERIC)
    for row in filereader:
        xs.append(row)
xl.save(filename + '.xlsx')

print filename + ' is created.'
raw_input('Press any key to exit...')