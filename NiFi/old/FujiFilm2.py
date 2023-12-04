import os
import sys
import datetime
from pathlib import Path
import mysql.connector



try:
##	path = sys.argv[1]
##	filename = sys.argv[2]
	path = r"C:\Users\efu\Desktop\AEI203664\aei203664_gel 1_9days_021821_SIZING__smh-[phosphor].gel"
	filename = 'aei203664_gel 1_9days_021821_SIZING__smh-[phosphor].gel'
	
except Exception as e:
	#print(e)
	print("No arguments supplied, try:")
	print("Fujifilm.py file-path file-name")
	sys.exit()




	
# Database Connection (Pointed at Dev)
lims_db = mysql.connector.connect(host='192.168.21.240', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor()



# Set Filepath:
og_filepath = '{0}{1}'.format(path, filename)


### Get File Owner
##path_obj_of_ogfilepath = Path(og_filepath)
##owner = path_obj_of_ogfilepath.owner()
##
##
### File Timestamp
##from os import stat
##from pwd import getpwuid
##dt_c = getpwuid(stat(og_filepath).st_uid).pw_name
##
##
##
##
### Update Audit Trail
##db_insert = "INSERT INTO t5_nifi_fujifilm (Time, User, filepath) VALUES (%s,%s,%s)"
##db_insert_vals = (dt_c, owner, og_filepath)
##lims_db_cursor.execute(db_insert, db_insert_vals)
##lims_db.commit()



# File move
parent_folder_of_qlps_file = '\\'.join(og_filepath.split('\\')[:-1])
projid_of_qlps_file = og_filepath.split('\\')[-2]
print(projid_of_qlps_file)
