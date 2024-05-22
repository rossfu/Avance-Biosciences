import re
import os
import sys
import csv
import win32security
import pymysql.cursors
from datetime import datetime
sys.path.append("/usr/local/lib/python3.10/dist-packages")

# Set up variables passed from NiFi:
try:
    path = sys.argv[1].rstrip('\\')
    filename = sys.argv[2]
except Exception as e:
    #print(e)
    print("No arguments supplied, try:")
    print("tecan.py file-path file-name")
    sys.exit()
#print("path {0}".format(path))
#print("filename {0}".format(filename))

# only run code for CSV files:
if ".csv" in filename:
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

    # GRAB ProjectID from filepath!
    match = re.search(r'\d{6}', og_filepath)

    if match:
        _project_id = match.group(0)

        # Start by selecting the current last # from DB:
        sql = f"SELECT * from get_full_idproject WHERE idprojects = '{_project_id}';"
        lims_db_cursor.execute(sql)
        result = lims_db_cursor.fetchone()

        try:
            _project_id = result['full_idprojects']
        except:
            # should keep _project_id from before lookup (ie, 6digit projectID)
            pass
    else:
        # No ProjectID found
        _project_id = ""



    # Metadata user
    def GetOwner(filename):
        try:
            f = win32security.GetFileSecurity(filename, win32security.OWNER_SECURITY_INFORMATION)
            (username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
            return username
        except: username = 'N/A'
    metadata_user = GetOwner(og_filepath)
    print("metadata_user: {0}".format(metadata_user))


    ctime = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')
    mtime = datetime.fromtimestamp(os.path.getmtime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')
    print("ctime: {0}".format(ctime))
    print("mtime: {0}".format(mtime))
    #----------------------------

    try:
        # Start by selecting the current last # from DB:
        sql = f"SELECT max(insert_id) AS last_number FROM t6_tecan"
        lims_db_cursor.execute(sql)
        result = lims_db_cursor.fetchone()
    except Exception as e:
        pass

    try:
        insert_id = result['last_number']+1
    except Exception as e:
        insert_id = 1

    print(f"insert_id: {insert_id}")

    #----------------------------

    # Open file and read into CSV reader:
    #errors='ignore'
    try:
        with open(og_filepath, mode='r', encoding='utf-8', errors='ignore') as file:
            csv_reader = csv.reader(file, delimiter=',') 

            # Read the first row as headers, skip headers they are in DB cols
            headers = next(csv_reader)

            # Prepare all rows for insertion
            rows_to_insert = [(insert_id,_project_id,og_filepath,metadata_user,ctime,mtime,) + tuple(row) for row in csv_reader]
            #print(rows_to_insert)

            # Construct the SQL INSERT statement
            insert_stmt = "INSERT INTO t6_tecan (`insert_id`, `project_id`, `filename`, `user`, `ctime`, `mtime`, `Position`, `tubeID`, `ScanError`, `SRCRack`, `SRCPos`, `SRCTubeID`, `Volume`, `Error`, `SRCRackID`, `GridPos`, `SiteOnGrid`, `TipNumber`, `DetectVol`, `Time`) VALUES (%s,%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"

            # Execute a bulk insert
            try:
                lims_db_cursor.executemany(insert_stmt, rows_to_insert)
                lims_db.commit()  # Committing the transaction to the database
                print(f"Successfully inserted {len(rows_to_insert)} rows.")
            except Exception as e:
                print("Error inserting data: ", e)
                lims_db.rollback()  # Rollback in case of error
    except: 
        #skipping if empty file
        pass
    # close the cursor and connection
    lims_db_cursor.close()
    lims_db.close()