import os
import re
import sys
import win32security
import mysql.connector
from bs4 import BeautifulSoup
from datetime import datetime

sys.path.append("/usr/local/lib/python3.10/dist-packages")

        
        


def get_audit_trail():
        
    try: #Production mode
        path = '\\' + sys.argv[1].replace('\\\\', '\\')
        filename = sys.argv[2]
        extension = filename.split('.')[-1]
        print('\npath: ' + path)
        print('filename: ' + filename)
            
    except Exception as e:
        #print(e)
        print("No arguments supplied, try:")
        print("qpcr_pid.py file-path file-name")
        sys.exit()
            

    # Database Connection (Pointed at Dev)
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


                    

    # HTML SCRAPE for audit trail info
    audit_trail = ''
    username = ''
    starttime = ''
    stoptime = ''
    runtime = ''
    assc_file = ''
    robotnum = ''
    project_id = ''



    # GRAB ProjectID from filepath!
    match = re.search(r'[A-Z]{3}\d{6}', og_filepath)

    if match:
        project_id = match.group(0)



    # Audit Trail
    with open(og_filepath, 'r', encoding = 'ISO-8859-1') as htm_file:

        contents = htm_file.read()
        beautifulsouptext = BeautifulSoup(contents, 'lxml')
            
        audit_trail = str(beautifulsouptext.get_text())


        z = audit_trail.split('\n')



        for i in range(len(z)):
            if z[i].startswith('        Username : '):
                username = z[i].split('        Username : ')[-1]
                
            if 'Duration :  ' in z[i]:
                runtime = z[i].split('Duration :  ')[-1]

            if 'Duration   : ' in z[i]:
                runtime = z[i].split('Duration   : ')[-1]
                
            if 'Start time : ' in z[i]:
                starttime = z[i].split('Start time : ')[-1]
                
            if 'Stop Time  : ' in z[i]:
                stoptime = z[i].split('Stop Time  : ')[-1]
                
            if z[i].startswith('Filename : '):
                assc_file = z[i].split('Filename : ')[-1]

            if z[i].startswith('Robot serial number : '):
                robotnum = z[i].split('Robot serial number : ')[-1]






    # CHECK FOR AN EXISTING RECORD OF THIS FILENAME (not including file extension)
    lims_db_cursor.execute(f'SELECT * FROM t5_nifi_qiagility WHERE file = "{filename}" AND mtime = "{mtime}"')  #Because multiple files with same name but different file extensions are produced for an analyzed plate, lets search the db to make sure we do not create duplicate records for the plate we are looking at.
    duplicate_records_of_this_file = lims_db_cursor.fetchone()
    if duplicate_records_of_this_file is None:



        # Update DB
        db_insert = "INSERT INTO t5_nifi_qiagility (device_id, filepath, file, file_owner, ctime, mtime, assc_file, operator, runtime, start_time, stop_time, proj, audit_trail) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        db_insert_vals = (robotnum, og_filepath, filename, metadata_user, ctime, mtime, assc_file, username, runtime, starttime, stoptime, project_id, audit_trail)
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
