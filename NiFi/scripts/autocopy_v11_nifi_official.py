# Avance Data Backup
# Developed by Eric Ross Fu in 2021 and 2023 and 2023 again
# ericrossfu@yahoo.com

#NiFi

import os
import sys
import shutil
import pandas as pd
import win32security
import mysql.connector
from hashlib import sha256
from datetime import datetime






# Database
lims_db = mysql.connector.connect(host='192.168.21.240', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor()

            


        

# Copy Function
def Auto_Copy():


        # Skip locked files
        if os.access(file_path, os.W_OK) == False:
            log('File: ' + file_path.split('\\')[-1] + ' is locked atm')

            db_insert = "INSERT INTO t5_nifi_autocopy (time, source, dest, hash, summary, log, owner, ctime, root) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            db_insert_vals = (datetime.now(), file_path, '', '', 'ERROR - file locked', loggy, '', '', root)
            lims_db_cursor.execute(db_insert, db_insert_vals)
            lims_db.commit()
            continue

        #Get Orig Hash
        orig_hash_process = sha256()
        try: 
            with open(file_path, 'rb') as original_file:
                for byte_block in iter(lambda: original_file.read(128000),b""):
                    orig_hash_process.update(byte_block)
            original_file.close()

        except PermissionError:
            log('Permissions Error')

            db_insert = "INSERT INTO t5_nifi_autocopy (time, source, dest, hash, summary, log, owner, ctime, root) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            db_insert_vals = (datetime.now(), file_path, '', '', 'ERROR - permissions', loggy, '', '', root)
            lims_db_cursor.execute(db_insert, db_insert_vals)
            lims_db.commit()
            continue
        



        # Audit Trail (user / ctime)
        global ctime
        global username
        username = GetOwner(file_path)
        ctime = datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S')


        # SHA 256 Hash
        global orig_hash
        orig_hash = orig_hash_process.hexdigest() # Now we can check the database for existing records of both filepaths and their sha256 hash





# Assessing this file's state of back up --------> initializing action paths
###############################################################################################################################################################################################################################################

        # Check if the path and hash is already in database 
        # (would mean current version of file was already backed up)
        lims_db_cursor.execute("SELECT source, hash FROM t5_nifi_autocopy WHERE source = '" + file_path + "' AND  hash = '" + orig_hash + "'")
        records_that_would_indicate_file_backup_is_uptodate = lims_db_cursor.fetchall()



        if len(records_that_would_indicate_file_backup_is_uptodate) != 0:   #no action needed
            
            old_file()


        else: # file isnt backed up yet, continue assessment

            # Check if brand new file
            lims_db_cursor.execute("SELECT source, hash FROM t5_nifi_autocopy WHERE source = '" + file_path + "'")
            records_of_this_source_file = lims_db_cursor.fetchall()



            if len(records_of_this_source_file) == 0: #new file, must backup and hash

                new_file()


            else: # the source file path is registered but the hash of this file is not. This means the file has been updated and must be backed up without tampering with the back up of the previous version of this file.

                modified_file()
            


#################################################################################################################################################################################################################################################







def new_file():
    copy_and_hash_check(file_path, dest_folder + filename, 'new', 1)
    





def modified_file():
    lims_db_cursor.execute("SELECT source, hash FROM t5_nifi_autocopy WHERE source = '" + file_path + "' and summary != 'No Action Needed'")
    previous_versions = lims_db_cursor.fetchall()

    num_previous_versions = len(previous_versions)
    current_version = num_previous_versions + 1
    log('version ' + str(current_version))   

    copy_and_hash_check(file_path, dest_folder + 'v' + str(current_version) + ' --- ' + file_name, 'modified', current_version)








def copy_and_hash_check(source, dest, summary, version):

    # Copy
    shutil.copy2(source, dest)
    log('dest path: ' + dest)

    # Hash
    copyfile_hash_process = sha256()
    with open(dest, 'rb') as copy_file:
        for byte_block in iter(lambda: copy_file.read(128000),b""):
            copyfile_hash_process.update(byte_block)
    copy_file.close()
    copy_file_hash = copyfile_hash_process.hexdigest()

    
    log('source file hash: ' + orig_hash)
    log('dest file hash:   ' + copy_file_hash)
    

    #Match Check
    if orig_hash != copy_file_hash: #Error
        log('Hash mismatch')
        db_insert = "INSERT INTO t5_nifi_autocopy (time, source, dest, hash, summary, log, owner, ctime, root) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        db_insert_vals = (datetime.now(), file_path, dest, orig_hash, 'ERROR - hash', loggy, '', '', root)
        lims_db_cursor.execute(db_insert, db_insert_vals)
        lims_db.commit()

    else: # Transfer Successful
        db_insert = "INSERT INTO t5_nifi_autocopy (time, source, dest, hash, summary, log, owner, ctime, root) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        db_insert_vals = (datetime.now(), file_path, dest, orig_hash, summary, loggy, username, ctime, root)
        lims_db_cursor.execute(db_insert, db_insert_vals)
        lims_db.commit()









def old_file():
    pass







def log(msg):
    global loggy
    loggy += msg + '\n'
    print(msg)





def GetOwner(filename):
    try:
        f = win32security.GetFileSecurity(filename, win32security.OWNER_SECURITY_INFORMATION)
        (username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
        return username

    except BaseException as e:
        
        return 'N/A'




def exception_handler():

    import traceback
    exc_type, exc_value, exc_traceback = sys.exc_info()
    log('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))

    db_insert = "INSERT INTO t5_nifi_autocopy (time, source, dest, hash, summary, log, owner, ctime, root) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"
    db_insert_vals = (datetime.now(), file_path, dest_folder + '\\' + filename, '', 'ERROR - program', loggy, '', '', '')
    lims_db_cursor.execute(db_insert, db_insert_vals)
    lims_db.commit()

    return 'N/A'






# globalize these variables, less to track as information moves from function to function
global source_folder
global dest_folder
global file_path
global filename
global loggy
global root




# Program Execution Manager


try:
    
    source_folder = '\\' + sys.argv[1].replace(r'\\', '\\')
    filename = sys.argv[2]
    loggy = ''
    
    file_path = source_folder + filename


    dest_folder = r"\\45drives\hometree\TempInstrumentBackupFolder\NiFiAutoCopy" + '\\'.join(source_folder.split('\\')[1:])
    if os.path.exists(dest_folder) == False:
        os.makedirs(dest_folder)

    if source_folder.split('\\')[2].lower() == 'birch':
        root = source_folder

    else:
        root = source_folder.split('\\')[2]


        
    log('For ' + root + ' Back Up')
    print('')
    log('path: ' + file_path)




    
    Auto_Copy() # Lets back this file up



except BaseException as e:
    
    exception_handler()

