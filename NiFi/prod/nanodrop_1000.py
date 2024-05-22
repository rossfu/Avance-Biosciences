import re
import os
import sys
import csv
import datetime
import pymysql.cursors
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import numpy as np
import math
sys.path.append("/usr/local/lib/python3.10/dist-packages")

# ProjectID is not reliable and messed up.. take a look.
# filename + fill filepath now. (we know where the files are going)

# Set up variables passed from NiFi:
try:
	path = sys.argv[1].rstrip('\\')
	filename = sys.argv[2]
except Exception as e:
	#print(e)
	print("No arguments supplied, try:")
	print("nanodrop_1000.py file-path file-name")
	sys.exit()
#print("path {0}".format(path))
#print("filename {0}".format(filename))


# Database Connection (Pointed at Dev)
# Connect to the database:
lims_db = pymysql.connect(host='192.168.21.69',
	user=os.environ['LIMS_USER'],
	password=os.environ['LIMS_PASS'],
	database='avancelims',
	port=3306,
	cursorclass=pymysql.cursors.DictCursor)
lims_db_cursor = lims_db.cursor()


# Set Filepath:
og_filepath = '{0}\{1}'.format(path, filename)
filename = os.path.basename(og_filepath)

print('filepath', og_filepath)
print('filename', filename)


# set year AUTOmatically here:
# current_year = datetime.datetime.now().year
# filepath_45 = rf"\\\\45drives\\hometree\\DataArchive\\{current_year}\\NanoDrop1000\\{filename}"
# print(filepath_45)

# GRAB ProjectID from filename!
match = re.search(r'[A-Z]{3}\d{6}', og_filepath)

if match:
	_project_id = match.group(0)
	# _project_id = _project_id[3:]
else:
	_project_id = ""
#----------------------------

with open(og_filepath, 'r', encoding='ISO-8859-1') as f:
	lines = f.readlines()
	first_lines = lines[:4]
	tsv_lines = lines[4:] # extract the main TSV:

	# extract the first 4 lines
	_module = first_lines[0].split("\t")[1].strip()
	_path = first_lines[1].split("\t")[1].strip()
	_software = first_lines[2].split("\t")[1].strip()


	tsv_reader = csv.reader(tsv_lines, delimiter='\t')

	# Read the first row as headers
	headers = next(tsv_reader) # take CNT out? does this skip past the headers for us?

	'''
	# different cols for cell culture:
	header_row = headers[16:]
	for row in tsv_reader:
		# grab only plotable cols:
		data_row = row[16:]

		# convert the x and y data to arrays of floats
		header_row = np.array([float(val) for val in header_row])
		data_row = np.array([float(val) for val in data_row])

		plot_data = { 'x': header_row, 'y': data_row}
		#print(plot_data)
		df = pd.DataFrame(plot_data)

		# Create a line plot using Seaborn
		sns.lineplot(x='x', y='y', data=df)

		# Set the x-axis tick labels to display only every 10th value
		plt.xticks(df['x'][::10])

		#[If the scales are different (0.01 rather than 0.1) we change the ylim and yticks??]

		# Set the y-axis range to start at the minimum data value rounded down to the nearest 0.1, and end at the maximum data value rounded up to the nearest 0.1
		plt.ylim(math.floor(min(data_row) / 0.1) * 0.1, math.ceil(max(data_row) / 0.1) * 0.1)

		# Set the y-axis tick marks to display every 0.1 unit
		plt.yticks(np.arange(math.floor(min(data_row) / 0.1) * 0.1, math.ceil(max(data_row) / 0.1) * 0.1, 0.1))

		# Display the plot
		plt.show()
		print("___________________")

		# Save the plot as a PNG image
		plt.savefig('my_plot.png')
	'''

	try:
		# Start by selecting the current last # from DB:
		sql = f"SELECT max(insert_id) AS last_number FROM t6_nanodrop"
		lims_db_cursor.execute(sql)
		result = lims_db_cursor.fetchone()
	except Exception as e:
		pass

	try:
		insert_id = result['last_number']+1
	except Exception as e:
		insert_id = 1

	#'''
	_first_val = None
	for row in tsv_reader:
		print(_module)
		if "Nucleic Acid" in _module:
			# strip whitespace:
			try:
				good_cols = row[:15] + [row[26]]
				good_cols = [col.strip() if isinstance(col, str) else col for col in good_cols]

				# Format fields:
				if _first_val is None:
					sampleid = good_cols[0]
				else:
					sampleid = _first_val
					_first_val = None
				userid = good_cols[1]
				date = good_cols[2]
				time = good_cols[3]

				# Convert the date string to a datetime object:
				date_obj = datetime.datetime.strptime(date, '%m/%d/%Y')
				# Convert the datetime object to a string in the correct format:
				date_str = date_obj.strftime('%Y-%m-%d')

				# Convert the time string to a datetime object:
				time_obj = datetime.datetime.strptime(time, '%I:%M %p')
				# Convert the datetime object to a string in the correct format:
				time_str = time_obj.strftime('%H:%M:%S')

				conc = good_cols[4]
				a260 = good_cols[5]
				a280 = good_cols[6]
				a260_280 = good_cols[7]
				a260_230 = good_cols[8]
				a230 = good_cols[9]
				serial = good_cols[14]
			except:
				_first_val= row[0]
				print("skipping")
				continue

			try:
				print("inserting")
				# Insert into DB:
				db_insert = f"INSERT INTO t6_nanodrop (`NanodropType`,`Serial`, `ProjectID`, `SampleID`, `UserID`,  `Date`, `Time`, `Conc`, `Units`, `A260`, `A280`, `260_280`, `260_230`, `230`, `module`, `path`, `software`, `filepath`, `insert_id`) VALUES ('1000', '{serial}', '{_project_id}', '{sampleid}','{userid}','{date_str}','{time_str}',{conc},'ng/ul',{a260},{a280},{a260_280},{a260_230},{a230},'{_module}','{_path}','{_software}', '{og_filepath.replace('\\',r'\\')}', '{insert_id}')"
				#print(db_insert)
				lims_db_cursor.execute(db_insert)
				lims_db.commit()
			except Exception as e:
				print(e)
				sys.exit()
		elif "Cell Culture" in _module:
			try:
				# strip whitespace:
				good_cols = row[:10] + [row[31]] # 280
				good_cols = [col.strip() if isinstance(col, str) else col for col in good_cols]
	
				# Format fields:
				if _first_val is None:
					sampleid = good_cols[0]
				else:
					sampleid = _first_val
					_first_val = None
				userid = good_cols[1]
				date = good_cols[2]
				time = good_cols[3]
				_baseline = good_cols[4]
				_cell_600 = good_cols[5]
				serial = good_cols[9] 
				_cell_280 = good_cols[-1]
	
				# Convert the date string to a datetime object:
				date_obj = datetime.datetime.strptime(date, '%m/%d/%Y')
				# Convert the datetime object to a string in the correct format:
				date_str = date_obj.strftime('%Y-%m-%d')
	
				# Convert the time string to a datetime object:
				time_obj = datetime.datetime.strptime(time, '%I:%M %p')
				# Convert the datetime object to a string in the correct format:
				time_str = time_obj.strftime('%H:%M:%S')
			except:
				_first_val= row[0]
				print("skipping")
				continue

			try:
				# Insert into DB:
				db_insert = f"INSERT INTO t6_nanodrop (`NanodropType`, `Serial`, `ProjectID`, `SampleID`, `UserID`,  `Date`, `Time`, `cell_baseline`, `cell_600`, `cell_280`,`module`, `path`, `software`, `filepath`, `insert_id`) VALUES ('1000', '{serial}', '{_project_id}', '{sampleid}','{userid}','{date_str}','{time_str}',{_baseline}, {_cell_600}, {_cell_280}, '{_module}','{_path}','{_software}','{og_filepath.replace('\\',r'\\')}','{insert_id}')"
				#print(db_insert)
				lims_db_cursor.execute(db_insert)
				lims_db.commit()
			except Exception as e:
				print(e)
				sys.exit()
		else: 
			# strip whitespace:
			good_cols = row[:15] + [row[26]]
			good_cols = [col.strip() if isinstance(col, str) else col for col in good_cols]

			print("[UNKNOWN TYPE]:")
			print(good_cols)
			print("-------")

			try:
				# Insert into DB:
				db_insert = f"INSERT INTO t6_nanodrop (`module`) VALUES ('{_module}')"
				#print(db_insert)
				lims_db_cursor.execute(db_insert)
				lims_db.commit()
			except Exception as e:
				print(e)
				sys.exit()
	#'''
