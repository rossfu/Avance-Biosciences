import os
import sys
import win32security
import mysql.connector
from datetime import datetime
sys.path.append("/usr/local/lib/python3.10/dist-packages")


def get_audit_trail():
    try: #Production mode
        path = sys.argv[1].rstrip('\\') + '\\'
        filename = sys.argv[2]

        print('\npath: ' + path)
        print('filename: ' + filename)
            
    except Exception as e:
        #print(e)
        print("No arguments supplied, try:")
        print("artel.py file-path file-name")
        sys.exit()

    # Database Connection (Pointed at Prod)
    lims_db = mysql.connector.connect(host='192.168.21.69', user=os.environ['LIMS_USER'], password=os.environ['LIMS_PASS'], database='avancelims', auth_plugin='mysql_native_password')
    lims_db_cursor = lims_db.cursor(buffered = True)

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

    # Grab file content
    file_text = ""

    with open(og_filepath, "r") as config_file:
        for line in config_file:
            file_text += line               

    # CHECK FOR AN EXISTING RECORD OF THIS FILENAME (not including file extension)
    lims_db_cursor.execute("SELECT * FROM t6_config_files WHERE filename = '" + filename + "' AND modification_date = '" + mtime + "'")  
    exact_dupe = lims_db_cursor.fetchone()

    if exact_dupe is not None:
        print("Dupe")
        return

    # CHECK FOR AN EXISTING RECORD OF THIS FILENAME (not including file extension)
    lims_db_cursor.execute("SELECT * FROM t6_config_files WHERE filename = '" + filename + "'")  
    exists = lims_db_cursor.fetchone()

    if not exists:
        comment = "Initial creation"
    else:
        comment = None

    db_insert = "INSERT INTO t6_config_files (location, filename, creation_date, modification_date, user, content, comment) VALUES (%s,%s,%s,%s,%s,%s,%s)"
    db_insert_vals = (path, filename, ctime, mtime, metadata_user, file_text, comment)
    lims_db_cursor.execute(db_insert, db_insert_vals)
    lims_db.commit()



try:
    get_audit_trail()
except BaseException as e:
    import traceback
    exc_type, exc_value, exc_traceback = sys.exc_info()
    print('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))

    # force error to categorize as non zero status
    print(1 + 'a')
