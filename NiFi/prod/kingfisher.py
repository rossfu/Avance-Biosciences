import os
import sys
import shutil
import pdfplumber
import mysql.connector
from datetime import datetime
sys.path.append("/usr/local/lib/python3.10/dist-packages")


#Issues
#datafile table does not have updated mtimes


try: #Production mode
    global filename
    global extension
    path = '\\' + sys.argv[1].replace('\\\\', '\\')
    filename = sys.argv[2]


    print('\npath: ' + path)
    print('filename: ' + filename)
    
except Exception as e:
    #print(e)
    print("No arguments supplied, try:")
    print("qpcr_datafiles.py file-path file-name")
    sys.exit()




def manage_files():  


    # Database Connection (Pointed at Dev)
    lims_db = mysql.connector.connect(host='192.168.21.69', user=os.environ['LIMS_USER'], password=os.environ['LIMS_PASS'], database='avancelims', auth_plugin='mysql_native_password')
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


    print('mysql filepath = ' + r'\\' + og_filepath.replace('\\',r'\\'))
    print('filename: ' + filename)






    # CHECK FOR AN EXISTING RECORD OF THIS FILEPATH
    lims_db_cursor.execute("SELECT * FROM t5_nifi_kingfisher_audit_trail WHERE filepath = '" + og_filepath.replace('\\',r'\\') + "' and mtime = '" + mtime + "'")  #Because multiple files with same name but different file extensions are produced for an analyzed plate, lets search the db to make sure we do not create duplicate records for the plate we are looking at.
    duplicate_check = lims_db_cursor.fetchone()

    if duplicate_check is None: #this is a new file

        #Projid
        projid = ''
        for folder_name in og_filepath.split('\\')[6:]:
            if len(folder_name) == 9 and folder_name[:3].isalpha() == True and folder_name[3:].isnumeric() == True:
                projid = folder_name


        # SCRAPE PDF FOR TEXT CONTENT
        text = ''


        try:
            with pdfplumber.open(og_filepath) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if isinstance(page_text, str):
                        text += page_text




                instrument_name = ''
                execution_started = ''
                executing_user = ''
                modified_by = ''





                for line in text.split('\n'):

                    if line.startswith('Instrument name '):
                        instrument_name = line.split('Instrument name ')[-1]
                        
                    if line.startswith('Execution started '):
                        execution_started = line.split('Execution started ')[-1]
                        
                    if line.startswith('Executing user '):
                        executing_user = line.split('Executing user ')[-1]

                    if line.startswith('Modified by '):
                        modified_by = line.split('Modified by ')[-1]



                

            # INSERT NEW ROW INTO DATA FILE AUDIT TRAIL DB
            db_insert = "INSERT INTO t5_nifi_kingfisher_audit_trail (filepath, file, file_owner, ctime, mtime, run_log, proj, instrument_name, execution_started, executing_user, modified_by) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            db_insert_vals = (og_filepath, filename, metadata_user, ctime, mtime, text, projid, instrument_name, execution_started, executing_user, modified_by)
            lims_db_cursor.execute(db_insert, db_insert_vals)
            lims_db.commit()
            print('kingfisher database updated')

        except:
            # INSERT NEW ROW INTO DATA FILE AUDIT TRAIL DB
            db_insert = "INSERT INTO t5_nifi_kingfisher_audit_trail (filepath, file, file_owner, ctime, mtime, run_log, proj, instrument_name, execution_started, executing_user, modified_by) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            db_insert_vals = (og_filepath, filename, metadata_user, ctime, mtime, 'File could not be opened', projid, '', '', '', '')
            lims_db_cursor.execute(db_insert, db_insert_vals)
            lims_db.commit()
            print('kingfisher database updated')



    else: #duplicate entry, skip
        print('duplicate, no action')
        pass










try:
    if filename.split('.')[-1] == 'pdf':
        manage_files()
except BaseException as e:
    import traceback
    exc_type, exc_value, exc_traceback = sys.exc_info()
    print('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))


    # force error to categorize as non zero status
    print(1 + 'a')
