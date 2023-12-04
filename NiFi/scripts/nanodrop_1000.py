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

# /home/avance/data-in/1000
# in:  /mnt/cifs/nanodrop1000/
# out: /mnt/cifs/dataarchive/2023/NanoDrop1000/

# ProjectID is not reliable and messed up.. take a look.
# filename + fill filepath now. (we know where the files are going)

# Set up variables passed from NiFi:
try:
	# comma in file name, real strings?
	path = sys.argv[1]
	path = path.replace(r"\\", "/")
	path = path.replace('"', '')
	filename = sys.argv[2]
	filename = filename.replace(r"\\", "/")
	filename = filename.replace('"', '')
except Exception as e:
	#print(e)
	print("No arguments supplied, try:")
	print("nanodrop_1000.py file-path file-name")
	sys.exit("no args")
print("path {0}".format(path))
print("filename {0}".format(filename))


# Database Connection (Pointed at Dev)
# Connect to the database:
lims_db = pymysql.connect(host='192.168.21.240',
	user='valuserall',
	password='ValUserAll1!',
	database='avancelims',
	port=3306,
	cursorclass=pymysql.cursors.DictCursor)
lims_db_cursor = lims_db.cursor()


# Set Filepath:
og_filepath = os.path.join(path, filename)
filename = os.path.basename(og_filepath)
print(f"og_filepath: {og_filepath}")

#og_filepath = og_filepath.replace("\\", "\")

# set year AUTOmatically here:
filepath_45 = r"\\45drives\hometree\DataArchive\2023\NanoDrop1000\{0}".format(filename)
print(f"filepath_45: {filepath_45}")

# GRAB ProjectID from filename!
match = re.search(r'[A-Z]{3}\d{6}', filename)

if match:
	_project_full = match.group(0)
	_project_id = _project_full[3:]
else:
	_project_id = "unspecified"
	_project_full = "unspecified"
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

	#'''
	for row in tsv_reader:
		if "Nucleic Acid" in _module:
			# strip whitespace:
			good_cols = row[:15] + [row[26]]
			good_cols = [col.strip() if isinstance(col, str) else col for col in good_cols]

			# Format fields:
			sampleid = good_cols[0]
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

			try:
				# Insert into DB:
				db_insert = f"INSERT INTO t6_nanodrop (`NanodropType`,`Serial`, `ProjectID`, `SampleID`, `UserID`,  `Date`, `Time`, `Conc`, `Units`, `A260`, `A280`, `260_280`, `260_230`, `230`, `module`, `path`, `software`, `filepath`) VALUES ('1000', '{serial}', '{_project_id}', '{sampleid}','{userid}','{date_str}','{time_str}',{conc},'ng/ul',{a260},{a280},{a260_280},{a260_230},{a230},'{_module}','{_path}','{_software}', '{filepath_45}')"
				#print(db_insert)
				lims_db_cursor.execute(db_insert)
				lims_db.commit()
			except Exception as e:
				print(e)
				sys.exit("insert1")
		elif "Cell Culture" in _module:
			# strip whitespace:
			good_cols = row[:10] + [row[31]] # 280
			good_cols = [col.strip() if isinstance(col, str) else col for col in good_cols]

			# Format fields:
			sampleid = good_cols[0]
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

			try:
				# Insert into DB:
				db_insert = f"INSERT INTO t6_nanodrop (`NanodropType`, `Serial`, `ProjectID`, `SampleID`, `UserID`,  `Date`, `Time`, `cell_baseline`, `cell_600`, `cell_280`,`module`, `path`, `software`, `filepath`) VALUES ('1000', '{serial}', '{_project_id}', '{sampleid}','{userid}','{date_str}','{time_str}',{_baseline}, {_cell_600}, {_cell_280}, '{_module}','{_path}','{_software}','{filepath_45}')"
				#print(db_insert)
				lims_db_cursor.execute(db_insert)
				lims_db.commit()
			except Exception as e:
				print(e)
				sys.exit("insert2")
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
				sys.exit("insert3")

# create shortcuts:
import subprocess
if _project_full != "unspecified":
	command = ['python', 'C:/nifi-1.21.0/scripts/create_shortcut.py', f'H:/DataBackup/2023/NanoDrop1000/{_project_full}', _project_full]
	subprocess.run(command, shell=True)

#"H:\DataBackup\2023\NanoDrop1000\XL1 Blue sonicated gDNA_7303_101420_EN.ndv"