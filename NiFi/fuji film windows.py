import os
import sys
import datetime
from pathlib import Path
import mysql.connector
from datetime import datetime



try:
	path = sys.argv[1].rstrip('\\')
	filename = sys.argv[2]
##	path = r"C:\Users\efu\Desktop\AEI203664"
##	filename = 'aei203664_gel 1_9days_021821_SIZING__smh-[phosphor].gel'
	
except Exception as e:
	#print(e)
	print("No arguments supplied, try:")
	print("Fujifilm.py file-path file-name")
	sys.exit()




	
# Database Connection (Pointed at Dev)
lims_db = mysql.connector.connect(host='192.168.21.240', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor()



# Set Filepath:
##og_filepath = '{0}{1}'.format(path, filename)
og_filepath = path + '\\' + filename



# File Owner / CTime
import win32security
def GetOwner(filename):
    f = win32security.GetFileSecurity(filename, win32security.OWNER_SECURITY_INFORMATION)
    (username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
    return username
owner = GetOwner(og_filepath)
cdate = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')




# Project ID
projid = 'N/A'
for x in og_filepath.split('\\'):
    if len(x) == 9 and x[:3].isalpha() == True and x[3:].isnumeric() == True:               
        projid = x
        break

    for z in x.split('-'):
        if len(z) == 9 and z[:3].isalpha() == True and z[3:].isnumeric() == True:               
            projid = z
            break

    for z in x.split('_'):
        if len(z) == 9 and z[:3].isalpha() == True and z[3:].isnumeric() == True:               
            projid = z
            break

    for z in x.split(' '):
        if len(z) == 9 and z[:3].isalpha() == True and z[3:].isnumeric() == True:               
            projid = z
            break




# Update Audit Trail
db_insert = "INSERT INTO t5_nifi_fujifilm (File_Creation_Time, User, filepath, Project_ID) VALUES (%s,%s,%s,%s)"
db_insert_vals = (cdate, owner, og_filepath, projid.upper())
lims_db_cursor.execute(db_insert, db_insert_vals)
lims_db.commit()
