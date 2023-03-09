import os
import sys
from datetime import datetime
import datefinder
import mysql.connector
from datetime import datetime
import dateutil.parser as dparser
from dateutil.relativedelta import relativedelta  


#LIMS
lims = mysql.connector.connect(host='192.168.21.240', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
db = lims.cursor()

##db.execute('SELECT * FROM avancelims.samples WHERE datecollected is not null and tubelabel is not null')
db.execute('SELECT * FROM avancelims.samples WHERE tubelabel is not null')
l_samples = db.fetchall()

dates = 0
s_count = 1


for sample in l_samples:

    date=''

    label = str(sample[6].lower().replace(':', ' '))
    label_space_split = label.split(' ')

    now = datetime.now()
    
    month_list = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec','january','february','march','april','june','july','august','september','october','november','december']
    month_detected = ''

    for x in month_list:
        if x in label:

##            label = label.replace(' '+month_detected, month_detected)
##            label = label.replace(month_detected+' ', month_detected)
##            label = label.replace(' '+month_detected, month_detected)
##            label = label.replace(month_detected+' ', month_detected)
##            label = label.replace(' '+month_detected, month_detected)
##            label = label.replace(month_detected+' ', month_detected)
            
            label_space_split = label.split(' ')
            label = label.replace('-',' ')

            label_stripped = label.replace(' ', '')
            m_index = label_stripped.index(x)


            #Datefinder has trouble when month abbreviations are used
            #ddyyyy extraction
            if label_stripped[m_index-2:m_index].isnumeric() and label_stripped[m_index+len(x):m_index+len(x)+4].isnumeric():
                matches = list(datefinder.find_dates(label_stripped[m_index-2:(m_index+len(x)+4)], strict=True))
                if len(matches) > 0:
                    date = matches[0]
                    print(str(s_count) + ':\t' + label + '\t\t\t\t' + str(date) + '\t\t\t' + str(matches) + 'did it')
                    dates+=1

            #ddyy extraction
            elif label_stripped[m_index-2:m_index].isnumeric() and label_stripped[m_index+len(x):m_index+len(x)+2].isnumeric():
                matches = list(datefinder.find_dates(label_stripped[m_index-2:(m_index+ len(x)+2)], strict=True))
                if len(matches) > 0:
                    date = matches[0]
                    print(str(s_count) + ':\t' + label + '\t\t\t\t' + str(date) + '\t\t\t' + str(matches) + 'didd it')
                    dates+=1

            #yyyydd extraction
            elif label_stripped[m_index-4:m_index].startswith('20'):
                if label_stripped[m_index+len(x):m_index+len(x)+2].isnumeric():
                    print('lookinn!')
                    matches = list(datefinder.find_dates(label_stripped[m_index+len(x):m_index+len(x)+2]+x+label_stripped[m_index-4], strict=True))
                    if len(matches) > 0:
                        date = matches[0]
                        print(str(s_count) + ':\t' + label + '\t\t\t\t' + str(date) + '\t\t\t' + str(matches) + 'didd it')
                        dates+=1                
                    
            if date=='':
                for i in label_space_split:
                    if x in i:
                        try:
                            date = datetime.strptime(i, '%d%b')
                            date = date.replace(year=strftime('%Y'))
                            if date > now:
                                date = date - relativedelta(years = 1)
                                dates+=1
                            print(str(s_count) + ':\t' + label + '\t\t\t\t' + str(date) + ' yahoo')
                        except:
                            donothing=1


                
    if date=='':
        for i in label_space_split:
            if len(i)==8 and i.isnumeric() and '202' in i:
                try:
                    date = datetime.strptime(i,'%m%d%Y')
                    print(str(s_count) + ':\t' + label + '\t\t\t\t' + str(date) + '\t\t\t' + str(matches) + ' wow')
                    dates+=1
                except:
                    date = datetime.strptime(i,'%Y%m%d')
                    print(str(s_count) + ':\t' + label + '\t\t\t\t' + str(date) + '\t\t\t' + str(matches) + ' nuu')
                    dates+=1
                finally:
                    donothing=1
            elif date=='':
                matches = list(datefinder.find_dates(i, strict=True))
                if len(matches) > 0:
                    date = matches[0]
                    print(str(s_count) + ':\t' + label + '\t\t\t\t' + str(date) + '\t\t\t' + str(matches) + ' printo')
                    dates+=1               

    if date=='':
        matches = list(datefinder.find_dates(label, strict=True))
        if len(matches) > 0:
            date = matches[0]
            print(str(s_count) + ':\t' + label + '\t\t\t\t' + str(date) + '\t\t\t' + str(matches) + ' plan z')
            dates+=1


    if date=='':
        print(str(s_count) + ':\t' + label + '\t\t\t\t' + 'No dates found')

    print(str(dates) + '\n')

    if s_count == 1000:
        print('\n' + str(dates) + ' dates detected')
        print('compared to ' + str(s_count) + ' total')
        break

    s_count+=1




