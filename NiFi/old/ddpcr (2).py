import os
import sys
import shutil
import mysql.connector
from datetime import datetime
sys.path.append("/usr/local/lib/python3.10/dist-packages")


#Issues
#This script can't work beyond the year 2029.
#hardcoded db creds, dest path






# Set up variables passed from NiFi:


##try: #Python test mode
##	path = r'C:\Users\efu\Desktop\qlps' + '\\'
##	filename = 'D120122P2_CTB301130.qlps'
##
##	print(path)

        
try: #Production mode
	path = sys.argv[1]
	filename = sys.argv[2]

	print('path: ' + path)
	print('filename: ' + filename)
	
except Exception as e:
	#print(e)
	print("No arguments supplied, try:")
	print("ddpct_test.py file-path file-name")
	sys.exit()
	


	
# Database Connection (Pointed at Dev)
lims_db = mysql.connector.connect(host='192.168.21.240', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor()




# Set Filepath:
og_filepath = '{0}{1}'.format(path, filename)
##loggy = 'From filepath: ' + og_filepath + '\n\n202'



# GET AUDIT TRAIL
z = 0
with open(og_filepath, 'r', encoding='ISO-8859-1') as file_in:

    write_file = False
    file_contents = file_in.readlines()
    for line_num in range(len(file_contents)):
        line_content = file_contents[line_num]

        # insert these lines into the DB
        if write_file == True:
            z+=1
            if z == 1:
                timestamp = ':'.join(line_content.split(':')[:3])
            if line_num == len(file_contents) - 1:
                print('end of file')
                write_file = False


            loggy += line_content

                

        # if Run Start detected we write the rest of the file to log:
        if 'Run started' in line_content:
            if "201" in line_content:
                loggy = '201'
                loggy += line_content.split('201')[-1]
            elif "202" in line_content:
                loggy = '202'
                loggy += line_content.split('202')[-1]
            elif "203" in line_content:
                loggy = '203'
                loggy += line_content.split('203')[-1]
            else:
                print('year not detected')

            write_file = True
            user = line_content.split(':')[3]
            
file_in.close()



# Metadata user
from os import stat
from pwd import getpwuid

metadata_user = getpwuid(stat(og_filepath).st_uid).pw_name




# Move Folder
dest_path = r'/mnt/cifs/hometree/DataBackup/ddPCR'
projid = ''




# Asset ID
    # quantalife 1 = QX200-DR-1537
    # quantalife 2 = QX200-DR-1464
    # quantalife 3 = QX200-DR-2297

    # local qlps path:                              "C:\QuantaLife\Data\D120722P1_ddPCR PQ_2022-12-07-17-18\D120722P1_ddPCR PQ.qlps"
    # linux nifi server mounted qlps path:  "\mnt\cifs\quantalife3\Data\D120722P1_ddPCR PQ_2022-12-07-17-18\D120722P1_ddPCR PQ.qlps"



if path.split('/')[3][-1] == '1': # '/' because accessing paths mounted to a linux machine
    asset_id = 'QX200-DR-1537'

elif path.split('/')[3][-1] == '2':
    asset_id = 'QX200-DR-1464'

elif path.split('/')[3][-1] == '3':
    asset_id = 'QX200-DR-2297'

else:
    asset_id = 'N/A'


    


# File path
mounted_filepath_converted_to_localpath = r'C:\QuantaLife' + '\\' + '\\'.join(og_filepath.split('/')[4:]) # converting linux '/' slashes to windows '\' slashes
    






# Project ID
for x in filename[:-5].split('_'):

    if len(x) == 9 and x[:3].isalpha() == True and x[3:].isnumeric() == True:                 #FILENAME HAS PROJID, GOOD

        # Database
        projid = x
        db_insert = "INSERT INTO t5_nifi_ddpcr (Time, qlps_contents_User, qlps_metadata_User, Asset_ID, filepath, Log_Contents, Project_ID, Database_Entry_Time) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
        db_insert_vals = (timestamp, user, metadata_user, asset_id, mounted_filepath_converted_to_localpath, loggy, projid, datetime.now())
        lims_db_cursor.execute(db_insert, db_insert_vals)
        lims_db.commit()

        # Folder Move
        shutil.copytree(path, dest_path + '/' + projid, dirs_exist_ok = True)
        break


if len(projid) != 9:

    # Database
    print('Problem: Qlps file doesnt have projid: ' + projid)
    db_insert = "INSERT INTO t5_nifi_ddpcr (Time, qlps_contents_User, qlps_metadata_User, Asset_ID, filepath, Log_Contents, Project_ID, Database_Entry_Time) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
    db_insert_vals = (timestamp, user, metadata_user, asset_id, mounted_filepath_converted_to_localpath, loggy, 'PROJID NOT FOUND IN QLPS FILENAME', datetime.now())
    lims_db_cursor.execute(db_insert, db_insert_vals)
    lims_db.commit()

    # Folder Move
    shutil.copytree(path, dest_path + '/' + 'UnspecifiedPROJID' + '/' + filename, dirs_exist_ok = True)














