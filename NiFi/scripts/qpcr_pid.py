import os
import sys
import shutil
import win32security
import mysql.connector
from datetime import datetime
sys.path.append("/usr/local/lib/python3.10/dist-packages")


#Issues
#hardcoded db creds, dest path



# Set up variables passed from NiFi:

##try: #Python test mode
##	path = 
##	filename = 
##
##	print(path)



def get_datafiles_pid():

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
	def GetOwner(filename):
		try:
			f = win32security.GetFileSecurity(filename, win32security.OWNER_SECURITY_INFORMATION)
			(username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
			return username
		except: username = 'N/A'
	metadata_user = GetOwner(og_filepath)



	ctime = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')
	mtime = datetime.fromtimestamp(os.path.getmtime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')



	# WHICH FILE EXTENSION?
	extension = filename.split('.')[-1]
	print('extension: ' + extension)
	print('filename no ext: ' + filename.replace('.' + extension, ''))



	# CHECK FOR AN EXISTING RECORD OF THIS FILENAME (not including file extension)
	lims_db_cursor.execute("SELECT * FROM t5_nifi_qpcr_proj_ids WHERE analysis_file = '" + filename.replace('.' + extension, '') + "'")  #Because multiple files with same name but different file extensions are produced for an analyzed plate, lets search the db to make sure we do not create duplicate records for the plate we are looking at.
	duplicate_check_for_different_file_extension = lims_db_cursor.fetchone()





	if duplicate_check_for_different_file_extension is None: #new file, proceed

		# Projid (From folder path on birch, nested folders observed)
		projid = 'N/A'
		plate_id = 'N/A'
		startpos = 0
		endpos = 0
		platehasid = False

		for folder_name in og_filepath.split('\\')[6:]:
			if len(folder_name) == 9 and folder_name[:3].isalpha() == True and folder_name[3:].isnumeric() == True: #nested projid
				projid = folder_name

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




		# CHECK PROJ DB for a match (prioritize matching filename then plate ID) (this script looks for first match, datafiles script also writes to the database and will update on all matches)			
		lims_db_cursor.execute("SELECT * FROM t5_nifi_qpcr_proj_ids WHERE file LIKE '%" + filename.replace('.' + extension, '') + "%'")
		file_match = lims_db_cursor.fetchone()


		
		# 3 options for updating the DB:
		if file_match is None:
			lims_db_cursor.execute("SELECT * FROM t5_nifi_qpcr_proj_ids WHERE PlateID = '" + plate_id + "'")
			plate_match = lims_db_cursor.fetchone()
			
			if plate_match is None: #no match on file or plate, insert new row into db
				db_insert = "INSERT INTO t5_nifi_qpcr_proj_ids (PlateID, ProjectID, analysisfile_metadata_user, analysis_filepath, analysis_file, Last_Updated, analysis_file_ctime, analysis_file_mtime) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
				db_insert_vals = (plate_id, projid, metadata_user, og_filepath.replace('\\',r'\\'), filename.replace('.' + extension, ''), str(datetime.now()), ctime, mtime)
				lims_db_cursor.execute(db_insert, db_insert_vals)
				lims_db.commit()


				
			else:		   #match on plate, not file
				lims_db_cursor.execute("SELECT analysis_file, analysis_filepath FROM t5_nifi_qpcr_proj_ids WHERE PlateID = '" + plate_id + "'")
				existing_db_record_forthis_plateid = lims_db_cursor.fetchone()



				# The PROJ ID DB is designed so that plateIDs do not have duplicate entries. In reality, there are many files with same plateID. They are now appended multiple entries into one row to ensure 1-to-1 matching.
				old_filename = existing_db_record_forthis_plateid[0]
				old_filepath = existing_db_record_forthis_plateid[1] # We retrieve the old entries so that we can re-enter them (with additional info) as opposed to overwriting them.
				

				lims_db_cursor.execute("UPDATE t5_nifi_qpcr_proj_ids SET PlateID = '" + plate_id + "', ProjectID = '" + projid + "', analysisfile_metadata_user = '" + metadata_user + "', analysis_filepath = '" + old_filepath + ';' + og_filepath.replace('\\',r'\\') + "', analysis_file = '" + old_filename + ';' + filename.replace('.' + extension, '') + "', Last_Updated = '" + str(datetime.now()) + "', analysis_file_ctime = '" + ctime + "', analysis_file_mtime = '" + mtime + "' WHERE PlateID = '" + plate_id + "'")
				lims_db.commit()


				
		else:			   #match on filename, good
				lims_db_cursor.execute("UPDATE t5_nifi_qpcr_proj_ids SET PlateID = '" + plate_id + "', ProjectID = '" + projid + "', analysisfile_metadata_user = '" + metadata_user + "', analysis_filepath = '" + old_filepath + ';' + og_filepath.replace('\\',r'\\') + "', analysis_file = '" + old_filename + ';' + filename.replace('.' + extension, '') + "', Last_Updated = '" + str(datetime.now()) + "', analysis_file_ctime = '" + ctime + "', analysis_file_mtime = '" + mtime + "' WHERE file = '" + filename.replace('.' + extension, '') + "'")
				lims_db.commit()










	else: #projid in db already, this plate has been encountered just as a different file extension, no action needed as far as getting this birch file into the DB
		print('pass')
		pass





try:
	get_datafiles_pid()


except BaseException as e:
	import traceback
	exc_type, exc_value, exc_traceback = sys.exc_info()
	print('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))


	# force error to categorize as non zero status
	print(1 + 'a')
