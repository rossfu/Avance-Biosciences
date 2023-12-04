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



def get_sds_files_pid():        
	try: #Production mode
		path = sys.argv[1].rstrip('\\') + '\\'
		filename = sys.argv[2]

		print('\npath: ' + path)
		print('filename: ' + filename)
		
	except Exception as e:
		#print(e)
		print("No arguments supplied, try:")
		print("qpcr_pid.py file-path file-name")
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
	# cdate = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')




	# Projid (From folder path on birch, nested folders observed)
	projid = 'N/A'

	for folder_name in og_filepath.split('\\')[6:]:
		if len(folder_name) == 9 and folder_name[:3].isalpha() == True and folder_name[3:].isnumeric() == True: #nested projid
			projid = folder_name





	# CHECK PROJ DB
	lims_db_cursor.execute("SELECT * FROM t5_nifi_qpcr_pid WHERE sds_file = '" + filename.split('.txt')[0] + "'")
	record = lims_db_cursor.fetchone()



	# WRITE TO Project_ID DB
	if record is None:
		db_insert = "INSERT INTO t5_nifi_qpcr_pid (ProjectID, sds_analysis_metadata_user, sds_analysis_filepath, sds_analysis_file, Last_Updated) VALUES (%s,%s,%s,%s,%s)"
		db_insert_vals = (projid, metadata_user, og_filepath.replace('\\',r'\\'), filename.split('.txt')[0], str(datetime.now()))
		lims_db_cursor.execute(db_insert, db_insert_vals)
		lims_db.commit()

	else:
		lims_db_cursor.execute("UPDATE t5_nifi_qpcr_pid SET ProjectID = '" + projid + "', sds_analysis_metadata_user = '" + metadata_user + "', sds_analysis_filepath = '" + og_filepath + "', sds_analysis_file = '" + filename.split('.txt')[0] + "', Last_Updated = '" + str(datetime.now()) + "' WHERE sds_file = '" + filename.split('.txt')[0] + "'")
		lims_db.commit()





try:
	get_sds_files_pid()
except BaseException as e:
	import traceback
	exc_type, exc_value, exc_traceback = sys.exc_info()
	print('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))


	# force error to categorize as non zero status
	print(1 + 'a')