import os
import sys
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
loggy = 'From filepath: ' + og_filepath + '\n\n202'


proj_id = filename.split('_')[-1].split('.qlps')[0]
print('projid from filename:' + proj_id)


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

##                db_insert = "INSERT INTO t5_nifi_ddpcr (Time, User, Project_ID, Log_Contents) VALUES (%s,%s,%s,%s)"
##                db_insert_vals = (timestamp, user, proj_id, loggy)
##                lims_db_cursor.execute(db_insert, db_insert_vals)
##                lims_db.commit()

            loggy += line_content

                

        # if Run Start detected we write the rest of the file to log:
        if 'Run started' in line_content and "202" in line_content:

            write_file = True
            user = line_content.split(':')[3]
            loggy += line_content.split('202')[-1]
            
file_in.close()







# Move Folder
filename_projid = filename.split('_')[-1].split('.qlps')[0]      
dest_path = r'/mnt/cifs/hometree/DataBackup/ddPCR'
##dest_path = r'\\45drives\hometree\DataBackup\ddPCR'
print(dest_path)

if len(filename_projid) == 9 and filename_projid[:3].isalpha() == True and filename_projid[3:].isnumeric() == True:                 #FILENAME HAS PROJID, GOOD
    db_insert = "INSERT INTO t5_nifi_ddpcr (Time, User, Project_ID, Log_Contents, Database_Entry_Time) VALUES (%s,%s,%s,%s,%s)"
    db_insert_vals = (timestamp, user, proj_id, loggy, datetime.now())
    lims_db_cursor.execute(db_insert, db_insert_vals)
    lims_db.commit()

##    if os.path.exists(dest_path + '\\' + filename_projid) == False:
##        os.mkdir(dest_path + '\\' + filename_projid)
        
    shutil.copytree(path, dest_path + '\\' + filename_projid)

    
else:
    print('Problem: Qlps file doesnt have projid: ' + filename_projid)
    db_insert = "INSERT INTO t5_nifi_ddpcr (Time, User, Project_ID, Log_Contents, Database_Entry_Time) VALUES (%s,%s,%s,%s,%s)"
    db_insert_vals = (timestamp, user, 'PROJID NOT FOUND IN QLPS FILENAME, FOLDER NOT TRANSFERRED TO BACKUP', loggy, datetime.now())
    lims_db_cursor.execute(db_insert, db_insert_vals)
    lims_db.commit()


##    os.mkdir(r'\\45drives\hometree\DataBackup\ddPCR\UnspecifiedPROJID' + '\\' + filename)
    shutil.copytree(path, r'\\45drives\hometree\DataBackup\ddPCR\UnspecifiedPROJID' + '\\' + filename)












