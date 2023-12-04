import os
import sys
import shutil
import mysql.connector
from datetime import datetime
# sys.path.append("/usr/local/lib/python3.10/dist-packages")


#Issues
#hardcoded db creds, dest path



# Set up variables passed from NiFi:

##try: #Python test mode
##	path = 
##	filename = 
##
##	print(path)

        
try: #Production mode
	path = sys.argv[1].rstrip('\\') + '\\'
	filename = sys.argv[2]

	print('\npath: ' + path)
	print('filename: ' + filename)
	
except Exception as e:
	#print(e)
	print("No arguments supplied, try:")
	print("7900.py file-path file-name")
	sys.exit()
	


# Database Connection (Pointed at Dev)
lims_db = mysql.connector.connect(host='192.168.21.240', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor()



# Set Filepath:
og_filepath = '{0}{1}'.format(path, filename)
print('filepath: ' + og_filepath)


# Metadata user
import win32security
def GetOwner(filename):
    f = win32security.GetFileSecurity(filename, win32security.OWNER_SECURITY_INFORMATION)
    (username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
    return username
metadata_user = GetOwner(og_filepath)
cdate = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')




# Move Folder
dest_path = ''
projid = ''





# Assets:
# 7900HT-6-1813		191.168.1.204		no nifi server / efu access to 7900HT-1812 Data				    D:\Applied Biosystems\SDS Documents\Q0420P1EC.sds
# 7900HT-4-1396		191.168.1.231		no nifi server / efu access to 7900HT-1395 Data	/ SDS docs		D:\Applied Biosystems\SDS Documents\SDS?
# 7900HT-2118		191.168.1.110		no nifi server / efu access to 7900HT-2117 Data				        D:\Applied Biosystems\SDS Documents\Q010423P5EC.sds


if path.split('\\')[2] == '191.168.1.204': 
    asset_id = '7900HT-6-1813'

elif path.split('\\')[2] == '191.168.1.231':
    asset_id = '7900HT-4-1396'

elif path.split('\\')[2] == '191.168.1.110':
    asset_id = '7900HT-2118'

else:
    asset_id = 'N/A'

print('asset id: ' + asset_id)



# File path                                                                                          
IP_filepath_converted_to_localpath = r'D:\Applied Biosystems' + '\\' + '\\'.join(og_filepath.split('\\')[2:]) 



# Update DB
doccy = ''
projid = ''


db_insert = "INSERT INTO t5_nifi_7900_sds (sds_metadata_timestamp, sds_metadata_user, Asset_ID, original_local_filepath, file, DOC_contents, Project_ID) VALUES (%s,%s,%s,%s,%s,%s,%s)"
db_insert_vals = (cdate, metadata_user, asset_id, IP_filepath_converted_to_localpath, filename, doccy, projid)
lims_db_cursor.execute(db_insert, db_insert_vals)
lims_db.commit()
print('database updated')





# MOVE SDS FILES

##############
