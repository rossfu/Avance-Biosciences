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

	print('\npath: ' + path)
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
print('filepath: ' + og_filepath)


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

            if line_num == len(file_contents) - 1:
                print('End of file')
                write_file = False


            loggy += line_content

                

        # if Run Start detected we write the rest of the file to log:
        if 'Run started' in line_content:
            print('')
            print('Run Started string located')


            if "201" in line_content:
                loggy = '201'
                loggy += line_content.split('201')[-1]
                user = line_content.split('201')[-1].split(':')[3]
                timestamp = '201' + ':'.join(line_content.split('201')[-1].split(':')[:3])

            elif "202" in line_content:
                loggy = '202'
                loggy += line_content.split('202')[-1]
                user = line_content.split('202')[-1].split(':')[3]
                timestamp = '202' + ':'.join(line_content.split('202')[-1].split(':')[:3])

            elif "203" in line_content:
                loggy = '203'
                loggy += line_content.split('203')[-1]
                user = line_content.split('203')[-1].split(':')[3]
                timestamp = '203' + ':'.join(line_content.split('203')[-1].split(':')[:3])

            else:
                print('year not detected')

            write_file = True


file_in.close()



# Metadata user
import win32security
def GetOwner(filename):
    f = win32security.GetFileSecurity(filename, win32security.OWNER_SECURITY_INFORMATION)
    (username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
    return username
metadata_user = GetOwner(og_filepath)
cdate = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')


# Move Folder
dest_path = r'\\45drives\hometree\DataBackup\ddPCR'
projid = ''




# Asset ID
    # quantalife 1 = QX200-DR-1537
    # quantalife 2 = QX200-DR-1464
    # quantalife 3 = QX200-DR-2297

    # local qlps path:                              "C:\QuantaLife\Data\D120722P1_ddPCR PQ_2022-12-07-17-18\D120722P1_ddPCR PQ.qlps"
    # windows qlps path:               "\\191.168.1.209\QuantaLife\Data\D120722P1_ddPCR PQ_2022-12-07-17-18\D120722P1_ddPCR PQ.qlps"
    # linux nifi server mounted qlps path:   "\mnt\cifs\quantalife3\Data\D120722P1_ddPCR PQ_2022-12-07-17-18\D120722P1_ddPCR PQ.qlps"


if path.split('\\')[2] == '191.168.1.209': 
    asset_id = 'QX200-DR-1537'

elif path.split('\\')[2] == '191.168.1.210':
    asset_id = 'QX200-DR-1464'

elif path.split('\\')[2] == '191.168.1.56':
    asset_id = 'QX200-DR-2297'

else:
    asset_id = 'N/A'

print('asset id: ' + asset_id)

    


# File path                                                                                           #C:\QuantaLife +
IP_filepath_converted_to_localpath = r'C:\QuantaLife' + '\\' + '\\'.join(og_filepath.split('\\')[4:]) #Data\D120722P1_ddPCR PQ_2022-12-07-17-18\D120722P1_ddPCR PQ.qlps"
    






# Project ID
for x in filename[:-5].split('_'):

    if len(x) == 9 and x[:3].isalpha() == True and x[3:].isnumeric() == True:                 #FILENAME HAS PROJID, GOOD

        # Database
        projid = x
        print('projid is: ' + projid)
        db_insert = "INSERT INTO t5_nifi_ddpcr (qlps_contents_timestamp, qlps_metadata_timestamp, qlps_contents_User, qlps_metadata_User, Asset_ID, filepath, file, Log_Contents, Project_ID,  Archived_filepath, Database_Entry_Time) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        db_insert_vals = (timestamp, cdate, user, metadata_user, asset_id, IP_filepath_converted_to_localpath, filename, loggy, projid, dest_path + '\\' + projid + '\\' + path.split('\\')[-2] + '\\' + filename, datetime.now())
        lims_db_cursor.execute(db_insert, db_insert_vals)
        lims_db.commit()
        print('database updated')

        # Folder Move
        shutil.copytree(path, dest_path + '\\' + projid + '\\' + path.split('\\')[-2], dirs_exist_ok = True)
        print('folder copied')
        break


if len(projid) != 9:

    # Database
    print('No projid in qlps file title')
    db_insert = "INSERT INTO t5_nifi_ddpcr (qlps_contents_timestamp, qlps_metadata_timestamp, qlps_contents_User, qlps_metadata_User, Asset_ID, filepath, file, Log_Contents, Project_ID, Archived_filepath, Database_Entry_Time) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
    db_insert_vals = (timestamp, cdate, user, metadata_user, asset_id, IP_filepath_converted_to_localpath, filename, loggy, 'PROJID NOT FOUND IN QLPS FILENAME', dest_path + '\\' + 'UnspecifiedPROJID' + '\\' + filename + '\\' + filename, datetime.now())
    lims_db_cursor.execute(db_insert, db_insert_vals)
    lims_db.commit()
    print('database updated')

    # Folder Move
    shutil.copytree(path, dest_path + '\\' + 'UnspecifiedPROJID' + '\\' + filename, dirs_exist_ok = True)
    print('folder copied')


print('Done')











