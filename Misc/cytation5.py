import os
import sys
import win32security
import mysql.connector
from bs4 import BeautifulSoup
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
        print("qpcr_pid.py file-path file-name")
        sys.exit()


    path = r"C:\Users\efu\Desktop\mac\cytation5\Y040224P1.xls"
    filename = Y040224P1.xls
    

    # Database Connection (Pointed at Dev)
    lims_db = mysql.connector.connect(host='192.168.21.240', user=os.environ['LIMS_USER'], password=os.environ['LIMS_PASS'], database='avancelims', auth_plugin='mysql_native_password')
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


                    


    # HTML SCRAPE for audit trail info
    audit_trail = ''
    
                        


    # CHECK FOR AN EXISTING RECORD OF THIS FILENAME (not including file extension)
    lims_db_cursor.execute("SELECT * FROM t5_nifi_cytation WHERE file = '" + filename + "' AND mtime = '" + mtime + "'")  #Because multiple files with same name but different file extensions are produced for an analyzed plate, lets search the db to make sure we do not create duplicate records for the plate we are looking at.
    duplicate_records_of_this_file = lims_db_cursor.fetchone()
    if duplicate_records_of_this_file is None:



        # Update DB
        db_insert = "INSERT INTO t5_nifi_cytation (filepath, file, file_owner, ctime, mtime, audit_trail) VALUES (%s,%s,%s,%s,%s,%s)"
        db_insert_vals = (og_filepath, filename, metadata_user, ctime, mtime, audit_trail)
        # removed several columns that were not in the DB table -JH 1/23/24
        lims_db_cursor.execute(db_insert, db_insert_vals)
        lims_db.commit()


            
    else:                                                                                                                                                                                               #no action, this is a duplicate
        print('duplicate, no action')
        pass





try:
    get_audit_trail()


except BaseException as e:
    
    import traceback
    exc_type, exc_value, exc_traceback = sys.exc_info()
    print('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))


    # force error to categorize as non zero status
    print(1 + 'a')
