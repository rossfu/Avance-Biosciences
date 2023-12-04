import os
import sys
import shutil
import mysql.connector
from datetime import datetime
# sys.path.append("/usr/local/lib/python3.10/dist-packages")




# Set up variables passed from NiFi:

##try: #Python test mode
##    path = 
##    filename = 
##
##    print(path)

        
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


# Metadata timestamp
cdate = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')




# Assets:
# 7900HT-6-1813        191.168.1.204        no nifi server / efu access to 7900HT-1812 Data                    D:\Applied Biosystems\SDS Documents\Q0420P1EC.sds
# 7900HT-4-1396        191.168.1.231        no nifi server / efu access to 7900HT-1395 Data    / SDS docs        D:\Applied Biosystems\SDS Documents\SDS?
# 7900HT-2118        191.168.1.110        no nifi server / efu access to 7900HT-2117 Data                        D:\Applied Biosystems\SDS Documents\Q010423P5EC.sds


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
IP_filepath_converted_to_localpath = r'D:\Applied Biosystems\SDS2.4\log' + '\\' + filename 




# Read File
run = 0
loggy = ''
logaroni = ''
sds_path = ''

with open(og_filepath, 'r', encoding='ISO-8859-1') as doc:
    doc_content = doc.readlines()

    for line in doc_content:
        if line.startswith('-----'):
            run += 1
            print('Run: ' + str(run))
            timestamp = line.split('Log ')[1].split('-----')[0]
            loggy += 'Run ' + str(run)
            loggy += '\n'
            loggy += line

        if 'StartRun' in line:
            loggy += line

        if 'StartPlate' in line:
            loggy += line
            sds_path += line.split('"')[1]
            sds_path += '\n'
        if 'scan completed' in line:
            loggy += line
            
        if 'Uninitializing COM port' in line:
            loggy += line + '\n'


        logaroni += line




# Update DB
db_insert = "INSERT INTO t5_nifi_7900_doc (log_metadata_timestamp, Asset_ID, original_local_filepath, file, num_sds, SDS_files_associated, short_log, DOC_contents) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
db_insert_vals = (cdate, asset_id, IP_filepath_converted_to_localpath, filename, run, sds_path, loggy, logaroni)
lims_db_cursor.execute(db_insert, db_insert_vals)
lims_db.commit()
print('database updated')