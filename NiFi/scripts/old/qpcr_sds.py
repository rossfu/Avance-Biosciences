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
##    path = 
##    filename = 
##
##    print(path)


def manage_sds_files():  
	try: #Production mode
		path = '\\' + sys.argv[1].replace('\\\\', '\\')
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




	#4th instrument, sds move to 45 drives

	# Assets:
	# 7900HT-6-1813        191.168.1.204        no nifi server / efu access to 7900HT-1812 Data                    D:\Applied Biosystems\SDS Documents\Q0420P1EC.sds
	# 7900HT-4-1396        191.168.1.231        no nifi server / efu access to 7900HT-1395 Data    / SDS docs        D:\Applied Biosystems\SDS Documents\SDS?
	# 7900HT-2118        191.168.1.110        no nifi server / efu access to 7900HT-2117 Data                    D:\Applied Biosystems\SDS Documents\Q010423P5EC.sds
	# 7900HT-1400        191.168.1.78        no nifi server / efu access to 7900HT-1036 Data                    D:\Applied Biosystems\SDS Documents\Q010423P5EC.sds


	if path.split('\\')[2] == '191.168.1.204': 
		asset_id = '7900HT-6-1813'
		dest = "\\\\191.168.1.204\\7900HT-1812 Data" + '\\' + filename
		if os.path.exists(dest) == False:
			dest = '\\' + dest.replace('\\\\', '\\')
			shutil.copy(og_filepath, dest)
			print('file transferred to protected folder')

	elif path.split('\\')[2] == '191.168.1.231':
		asset_id = '7900HT-4-1396'
		dest = "\\\\191.168.1.231\\7900HT-1395 Data" + '\\' + filename
		if os.path.exists(dest) == False:
			dest = '\\' + dest.replace('\\\\', '\\')
			shutil.copy(og_filepath, dest)        
			print('file transferred to protected folder')

	elif path.split('\\')[2] == '191.168.1.110':
		asset_id = '7900HT-2118'
		dest = "\\\\191.168.1.110\\7900HT-2117Data" + '\\' + filename
		if os.path.exists(dest) == False:
			dest = '\\' + dest.replace('\\\\', '\\')
			shutil.copy(og_filepath, dest)
			print('file transferred to protected folder')

	elif path.split('\\')[2] == '191.168.1.78':
		asset_id = '7900HT-1400'
		dest = "\\\\191.168.1.78\\7900HT-1036 Data" + '\\' + filename
		if os.path.exists(dest) == False:
			dest = '\\' + dest.replace('\\\\', '\\')
			shutil.copy(og_filepath, dest)        
			print('file transferred to protected folder')

	elif path.split('\\')[2] == '191.168.1.81': 
		asset_id = '7500-2'
		dest = "\\\\191.168.1.81\\7500-1771 Data" + '\\' + filename
		if os.path.exists(dest) == False:
			dest = '\\' + dest.replace('\\\\', '\\')
			shutil.copy(og_filepath, dest)
			print('file transferred to protected folder')

	else:
		asset_id = 'N/A'


	print('asset id: ' + asset_id)





	# File path                                                                                          
	IP_filepath_converted_to_localpath = r'D:\Applied Biosystems' + '\\' + '\\'.join(og_filepath.split('\\')[3:]) 



	# Update DB
	db_insert = "INSERT INTO t5_nifi_qpcr_sds (sds_metadata_timestamp, sds_metadata_user, Asset_ID, original_local_filepath, file) VALUES (%s,%s,%s,%s,%s)"
	db_insert_vals = (cdate, metadata_user, asset_id, IP_filepath_converted_to_localpath, filename.split('.sds')[0])
	lims_db_cursor.execute(db_insert, db_insert_vals)
	lims_db.commit()
	print('sds database updated')




	# CHECK PROJ DB
	lims_db_cursor.execute("SELECT * FROM t5_nifi_qpcr_pid WHERE SDS_analysis_filepath = '" + filename.split('.sds')[0] + '.txt' + "'") #compare plate names b/w 2 files
	record = lims_db_cursor.fetchone()



	# WRITE TO Project_ID DB
	if record is None:
		db_insert = "INSERT INTO t5_nifi_qpcr_pid (Asset_ID, sds_file_metadata_user, original_local_sds_filepath, sds_file, Last_Updated) VALUES (%s,%s,%s,%s,%s)"
		db_insert_vals = (asset_id, metadata_user, IP_filepath_converted_to_localpath, filename.split('.sds')[0], str(datetime.now()))
		lims_db_cursor.execute(db_insert, db_insert_vals)
		lims_db.commit()
		print('pid database updated')
	else:
		lims_db_cursor.execute("UPDATE t5_nifi_qpcr_pid SET sds_file_metadata_user = '" + metadata_user + "', Asset_ID = '" + asset_id + "', original_local_sds_filepath = '" + IP_filepath_converted_to_localpath + "', sds_file = '" + filename.split('.sds')[0] + "', Last_Updated = '" + str(datetime.now()) + "' WHERE SDS_analysis_file = '" + filename.split('.sds')[0] + '.txt' +  "'")
		lims_db.commit()
		print('pid database updated')


try:
	manage_sds_files()
except BaseException as e:
	import traceback
	exc_type, exc_value, exc_traceback = sys.exc_info()
	print('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))


	# force error to categorize as non zero status
	print(1 + 'a')