import os
import sys
import shutil
import mysql.connector
from datetime import datetime
sys.path.append("/usr/local/lib/python3.10/dist-packages")


#Issues
#hardcoded db creds, dest path



# Set up variables passed from NiFi:

##try: #Python test mode
##    path = 
##    filename = 
##
##    print(path)


def manage_files():  
	try: #Production mode
		path = '\\' + sys.argv[1].replace('\\\\', '\\')
		filename = sys.argv[2]
		extension = filename.split('.')[-1]

		print('\npath: ' + path)
		print('filename: ' + filename)
		
	except Exception as e:
		#print(e)
		print("No arguments supplied, try:")
		print("qpcr_datafiles.py file-path file-name")
		sys.exit()
		




	# Database Connection (Pointed at Dev)
	lims_db = mysql.connector.connect(host='192.168.21.240', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
	lims_db_cursor = lims_db.cursor(buffered = True)



	# Set Filepath:
	og_filepath = '{0}{1}'.format(path, filename)
	print('filepath: ' + og_filepath)


	# Metadata user
	import win32security
	def GetOwner(filename):
		try:
			f = win32security.GetFileSecurity(filename, win32security.OWNER_SECURITY_INFORMATION)
			(username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
			return username
		except:
			return 'N/A'

	metadata_user = GetOwner(og_filepath)

		
	ctime = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')
	mtime = datetime.fromtimestamp(os.path.getmtime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')




	#4th instrument, sds move to 45 drives

	# Assets:
	# 7900HT-6-1813	191.168.1.204	no nifi server / efu access to 7900HT-1812 Data		    D:\Applied Biosystems\SDS Documents\Q0420P1EC.sds
	# 7900HT-4-1396	191.168.1.231	no nifi server / efu access to 7900HT-1395 Data    / SDS docs	D:\Applied Biosystems\SDS Documents\SDS?
	# 7900HT-2118	191.168.1.110	no nifi server / efu access to 7900HT-2117 Data		    D:\Applied Biosystems\SDS Documents\Q010423P5EC.sds
	# 7900HT-1400	191.168.1.78	no nifi server / efu access to 7900HT-1036 Data		    D:\Applied Biosystems\SDS Documents\Q010423P5EC.sds


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



	#Quant Studio archives files automatically.
	elif path.split('\\')[2] == '191.168.1.98' or path.split('\\')[2] == '191.168.1.63' or path.split('\\')[2] == '191.168.1.16' or path.split('\\')[2] == '191.168.1.175': #1 of 4 QS7 Pros
		asset_id = 'QS 7 Pro'


	elif path.split('\\')[2] == '191.168.1.60' or path.split('\\')[2] == '191.168.1.52': #1 of 2 QS 12k Flex
		asset_id = 'QS 12k Flex'



	else:
		asset_id = 'N/A'


	print('asset id: ' + asset_id)
	print('mysql filepath = ' + r'\\' + og_filepath.replace('\\',r'\\'))
	print('filename no ext: ' + filename.split('.' + extension)[0])


	# PLATE ID
	plate_id = 'N/A'
	startpos = 0
	endpos = 0
	platehasid = False




	# CHECK FOR AN EXISTING RECORD OF THIS FILEPATH
	lims_db_cursor.execute("SELECT * FROM t5_nifi_qpcr_file_audit_trail WHERE filepath = '" + r'\\' + og_filepath.replace('\\',r'\\') + "'")  #Because multiple files with same name but different file extensions are produced for an analyzed plate, lets search the db to make sure we do not create duplicate records for the plate we are looking at.
	duplicate_check_for_different_file_extension = lims_db_cursor.fetchone()

	if duplicate_check_for_different_file_extension is None: #this is a new file



		# PLATE ID
		for i in range(len(filename)):
			if filename[i] == 'Q' and filename[i+1].isnumeric() and filename[i+2].isnumeric() and filename[i+3].isnumeric() and filename[i+4].isnumeric() and filename[i+5].isnumeric() and filename[i+6].isnumeric(): #Plate_ID
				startpos = i
				print('start position located')
				for z in range(startpos, len(filename)):
					if filename[z] == '.' or filename[z] == '_' or filename[z] == '-' or filename[z] == ' ':
						endpos = z
						print('end position located')
						platehasid = True
						break
				break
			if platehasid:
				plate_id = filename[startpos:endpos]





		# INSERT NEW ROW INTO DB
		db_insert = "INSERT INTO t5_nifi_qpcr_file_audit_trail (datafile_metadata_timestamp, datafile_metadata_user, Asset_ID, filepath, file, Database_Entry_timestamp, mtime, PlateID) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
		db_insert_vals = (ctime, metadata_user, asset_id, og_filepath, filename.split('.' + extension)[0], datetime.now(), mtime, plate_id)
		lims_db_cursor.execute(db_insert, db_insert_vals)
		lims_db.commit()
		print('qpcr file database updated')



	

		# Is there a file on BIRCH with the same name as this raw data file? (ideal)
		lims_db_cursor.execute("SELECT * FROM t5_nifi_qpcr_proj_ids WHERE analysis_file LIKE = '%" + filename.split('.' + extension)[0] + "%'") #compare plate names b/w 2 files
		file_match = lims_db_cursor.fetchone()




		# 3 options for updating the DB:
		if file_match is None:
			lims_db_cursor.execute("SELECT * FROM t5_nifi_qpcr_proj_ids WHERE PlateID = '" + plate_id + "'")
			plate_match = lims_db_cursor.fetchone()
			
			if plate_match is None: #no match on file or plate, insert new row into db
				db_insert = "INSERT INTO t5_nifi_qpcr_proj_ids (Asset_ID, datafile_metadata_user, filepath, file, PlateID, Last_Updated) VALUES (%s,%s,%s,%s,%s,%s)"
				db_insert_vals = (asset_id, metadata_user, og_filepath, filename.split('.' + extension)[0], plate_id, str(datetime.now()))
				lims_db_cursor.execute(db_insert, db_insert_vals)
				lims_db.commit()
				print('qpcr pid database updated')

				
			else:		   #match on plate, not file
				lims_db_cursor.execute("SELECT file, filepath FROM t5_nifi_qpcr_proj_ids WHERE PlateID = '" + plate_id + "'")
				existing_db_record_forthis_plateid = lims_db_cursor.fetchone()



				# The PROJ ID DB is designed so that plateIDs do not have duplicate entries. In reality, there are many files with same plateID. They are now appended multiple entries into one row to ensure 1-to-1 matching.
				old_filename = existing_db_record_forthis_plateid[0]
				old_filepath = existing_db_record_forthis_plateid[1] # We retrieve the old entries so that we can re-enter them (with additional info) as opposed to overwriting them.
				

				lims_db_cursor.execute("UPDATE t5_nifi_qpcr_proj_ids SET PlateID = '" + plate_id + "', ProjectID = '" + projid + "', datafile_metadata_user = '" + metadata_user + "', filepath = '" + old_filepath + ';' + og_filepath + "', file = '" + old_filename + ';' + filename.replace('.' + extension, '') + "', Last_Updated = '" + str(datetime.now()) + "' WHERE PlateID = '" + plate_id + "'")
				lims_db.commit()


				
		else:			   #match on filename, good
				lims_db_cursor.execute("UPDATE t5_nifi_qpcr_proj_ids SET PlateID = '" + plate_id + "', ProjectID = '" + projid + "', datafile_metadata_user = '" + metadata_user + "', filepath = '" + old_filepath + ';' + og_filepath + "', file = '" + old_filename + ';' + filename.replace('.' + extension, '') + "', Last_Updated = '" + str(datetime.now()) + "' WHERE file LIKE '%" + filename.replace('.' + extension, '') + "%'")
				lims_db.commit()




	else:
		pass

		

try:
	manage_files()
except BaseException as e:
	import traceback
	exc_type, exc_value, exc_traceback = sys.exc_info()
	print('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))


	# force error to categorize as non zero status
	print(1 + 'a')
