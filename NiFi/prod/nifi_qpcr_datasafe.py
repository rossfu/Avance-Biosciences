import os
import re
import sys
import shutil
import hashlib
import pandas as pd
import win32security
import mysql.connector
from datetime import datetime


try:

    path = sys.argv[1].rstrip('\\')
    filename = sys.argv[2]
    home_path = sys.argv[3]
    instrument = sys.argv[4].split('\\')[0]
    asset = sys.argv[4].split('\\')[1]

except Exception as e:
    #print(e)
    #print("No arguments supplied, try:")
    #print("versioning.py file-path file-name instrument")
    import traceback
    exc_type, exc_value, exc_traceback = sys.exc_info()
    with open('C:\\nifi-1.24.0\\log.txt', 'a') as file: file.write('\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))


def GetOwner(filepath):
    try:
        f = win32security.GetFileSecurity(filepath, win32security.OWNER_SECURITY_INFORMATION)
        (username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
    except:
        username = 'N/A'
    return username



# Database
lims_db = mysql.connector.connect(host='192.168.21.69', user=os.environ['LIMS_USER'], password=os.environ['LIMS_PASS'], database='avancelims', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor()


# Log
call_str = '\n\n' + str(datetime.now()) + f"\n\npath: {path}\nfilename: {filename}\ninstrument: {instrument}\nasset: {asset}\nhome_path: {home_path}\n\n"
print(call_str)
with open('C:\\nifi-1.24.0\\log.txt', 'a') as file: file.write(call_str)


# Set Filepath Vars:
og_filepath = '{0}\{1}'.format(path, filename)
print('full filepath: ' + og_filepath)
year = datetime.now().strftime('%Y')


# Metadata (user / ctime)
username = GetOwner(og_filepath)
ctime = datetime.fromtimestamp(os.path.getctime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')
mtime = datetime.fromtimestamp(os.path.getmtime(og_filepath)).strftime('%Y-%m-%d %H:%M:%S')


#Comprehensive PROJID Scrape
idproj = ''
project_folder_exist = False
match = re.findall(r'[A-Z]{3}\d{6}', og_filepath)
if match:
    idproj = match[-1]
    print(idproj)
    for folder_name in og_filepath.split('\\')[3:]: #folderpath scrape
        if folder_name == idproj:
            idproj = folder_name
            project_folder_exist = True



# Determine DEST root paths -- depending on the need for re-organization of folder
if project_folder_exist or idproj == '': #copy as is IF there is no projid OR if there is a projectID folder (which is what the archive wants to have)
    workfolder_root_path = rf"\\45drives\hometree\DataWorkFolder\{year}\{instrument}\{asset}"
    archive_root_path = rf"\\45drives\FDABackup\{year}\{instrument}\{asset}"
elif project_folder_exist == False and idproj != '': #create projid folder in archive IF the filename has projid but it is not in a projid folder
    workfolder_root_path = rf"\\45drives\hometree\DataWorkFolder\{year}\{instrument}\{asset}\{idproj}"
    archive_root_path = rf"\\45drives\FDABackup\{year}\{instrument}\{asset}\{idproj}"


# Establish dest paths
print('workfolder_root_path: ' + workfolder_root_path)
print('archive_root_path: ' + archive_root_path)

og_filepath_stripped = og_filepath.replace(home_path, '')
print('stripped source folderpath: ' + og_filepath_stripped + '\n\n')

work_folder_dest_path = "{0}\{1}".format(workfolder_root_path.rstrip('\\'), og_filepath_stripped.lstrip('\\'))
archive_dest_path = "{0}\{1}".format(archive_root_path.rstrip('\\'), og_filepath_stripped.lstrip('\\'))

with open('C:\\nifi-1.24.0\\log.txt', 'a') as file: file.write(f'og_filepath: {og_filepath}\nog_filepath_stripped: {og_filepath_stripped}\nworkfolder_dest_path: {work_folder_dest_path}\narchive_dest_path: {archive_dest_path}\n')
print('work_folder_dest_path: ' + work_folder_dest_path)
print('archive_dest_path: ' + archive_dest_path + '\n\n')



#hash
from hashlib import sha256
orig_hash_process = sha256()
with open(og_filepath, 'rb') as original_file:
    for byte_block in iter(lambda: original_file.read(128000),b""):
        orig_hash_process.update(byte_block)
original_file.close()
orig_hash = orig_hash_process.hexdigest()




#load records
lims_db_cursor.execute("SELECT source, hash FROM t5_nifi_datasafe")
history = lims_db_cursor.fetchall()
db_records = pd.DataFrame(history, columns = ('source', 'hash'))




def modified_file(file_path, dest_path, orig_hash, ctime, mtime, username, instrument, asset, proj):
    lims_db_cursor.execute("SELECT source, hash FROM t5_nifi_datasafe WHERE source = '" + file_path.replace('\\',r'\\') + "'")
    previous_versions = lims_db_cursor.fetchall()
    num_previous_versions = len(previous_versions)
    current_version = num_previous_versions + 1
    print(str(current_version) + ' versions of this file. ' + str(current_version - 1) + ' are in DataWorkFolder. Working on getting this most current version into Data Archive.\n\n')

    if os.path.exists(dest_path) == True:
        try:
            shutil.copy2(dest_path, '\\'.join(dest_path.split('\\')[:-1]) + '\\' + 'old' + '\\' + 'v' + str(current_version - 1) + ' --- ' + filename)
        except:
            os.makedirs('\\'.join(dest_path.split('\\')[:-1]) + '\\' + 'old')
            print('folder created for old: ' + '\\'.join(dest_path.split('\\')[:-1]) + '\\' + 'old')
            shutil.copy2(dest_path, '\\'.join(dest_path.split('\\')[:-1]) + '\\' + 'old' + '\\' + 'v' + str(current_version - 1) + ' --- ' + filename)
    

        # FILE DIFF
        import difflib
        def compare_text_files(file1_path, file2_path):
            # Try opening files with UTF-8 encoding
            try:
                with open(file1_path, 'r', encoding='utf-8') as file1:
                    file1_lines = file1.readlines()
                with open(file2_path, 'r', encoding='utf-8') as file2:
                    file2_lines = file2.readlines()
            # If encoding error occurs, try opening files with latin-1 encoding
            except UnicodeDecodeError:
                with open(file1_path, 'r', encoding='latin-1') as file1:
                    file1_lines = file1.readlines()
                with open(file2_path, 'r', encoding='latin-1') as file2:
                    file2_lines = file2.readlines()

            # Perform the comparison
            differ = difflib.Differ()
            diff = differ.compare(file1_lines, file2_lines)

            # Print out the differences
            differences = []
            differences.append("Differences between " + file1_path + " and " + file2_path + ":\n")
            for line in diff:
                if line.startswith('- ') or line.startswith('+ '):
                    differences.append(line)

            return '\n'.join(differences)


        # Compare the text files
        global filediff

        try:
            filediff = compare_text_files(file_path, dest_path)[:30000]
        except Exception as e:
            import traceback
            exc_type, exc_value, exc_traceback = sys.exc_info()
            filediff = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))



        print('formerly newest versioned backup of this file relocated to: ' + '\\'.join(dest_path.split('\\')[:-1]) + '\\' + 'old' + '\\' + 'v' + str(current_version- 1) + ' --- ' + filename)
        print('\n\nexisting backup file moved to old, attempting to copy the latest modified version of that file.\n\n')
        os.remove(dest_path)
        with open('C:\\nifi-1.24.0\\log.txt', 'a') as file: file.write('modified file, old file sent to old folder at: ' + '\\'.join(dest_path.split('\\')[:-1]) + '\\' + 'old' + '\\' + 'v' + str(current_version - 1) + ' --- ' + filename)

    return copy_and_hash_check(file_path, dest_path, orig_hash, ctime, mtime, username, instrument, asset, proj, 'modified')





def copy_and_hash_check(source, dest, orig_hash, ctime, mtime, username, instrument, asset, proj, summary):

    # Copy
    try:
        shutil.copy2(source, dest)
    except:
        os.makedirs('\\'.join(dest.split('\\')[:-1]), exist_ok=True)
        shutil.copy2(source, dest)
    print('copied')


    # Making sure copied file has same hash:
    copyfile_hash_process = sha256()
    with open(dest, 'rb') as copy_file:
        for byte_block in iter(lambda: copy_file.read(128000),b""):
            copyfile_hash_process.update(byte_block)
        copy_file_hash = copyfile_hash_process.hexdigest()
        copy_file.close()


    #Match Check
    if orig_hash != copy_file_hash: #Error
        os.remove(dest)
        with open('C:\\nifi-1.24.0\\log.txt', 'a') as file: file.write(f'hash mismatch!\n')
        print('hash mismatch, force non-zero' + 1)
    
    # Transfer Successful
    else:        
        return True




def update_db(source, dest, archive_dest, orig_hash, ctime, owner, mtime, instrument, asset, proj, filediff, summary):
    # Update DB
    db_insert = "INSERT INTO t5_nifi_datasafe (source, dest, archive_dest, hash, ctime, owner, mtime, instrument, asset, proj, modifications, summary) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
    db_insert_vals = (source, dest, archive_dest, orig_hash, ctime, owner, mtime, instrument, asset, proj, filediff, summary)
    lims_db_cursor.execute(db_insert, db_insert_vals)
    lims_db.commit()

    with open('C:\\nifi-1.24.0\\log.txt', 'a') as file: file.write(f'success summary: source: {source} | dest: {dest}')
    print('done\n\nSummary:\n')
    print('source file: ' + source)
    print('workfolder path: ' + dest)
    print('archive path: ' + archive_dest)



# Status Vars
work_folder_backup_status = False
archive_backup_status = False
modified_status = False
no_action_needed = False
filediff = ''

##########################################################################
with open(r"C:\scripts\prod\qpcr_date_filter.txt") as file:
    lines = file.readlines()
    year = int(lines[0])
    month = int(lines[1])
    day = int(lines[2])
    qpcr_timestamp = datetime(year, month, day, 0, 0, 0)                  # qPCR specific file date filter
    print(datetime.fromtimestamp(os.path.getctime(og_filepath)))
    print(qpcr_timestamp)
##########################################################################

# qPCR filter
if datetime.fromtimestamp(os.path.getctime(og_filepath)) > qpcr_timestamp:


    # Classify
    if og_filepath not in set(db_records['source']): #unseen filepath = new file
        print('new file')
        work_folder_backup_status = copy_and_hash_check(og_filepath, work_folder_dest_path, orig_hash, ctime, mtime, username, instrument, asset, idproj, 'new')
        archive_backup_status = copy_and_hash_check(og_filepath, archive_dest_path, orig_hash, ctime, mtime, username, instrument, asset, idproj, 'new')


    elif orig_hash in set(db_records['hash']): #filepath and hash previously seen = old file
        print('No action needed\n')
        no_action_needed = True


    else:
        print('modified file')
        modified_status = modified_file(og_filepath, work_folder_dest_path, orig_hash, ctime, mtime, username, instrument, asset, idproj)




    # Update DB if success
    if work_folder_backup_status == True and archive_backup_status == True:
        update_db(og_filepath, work_folder_dest_path, archive_dest_path, orig_hash, ctime, username, mtime, instrument, asset, idproj, filediff, 'new')

    elif modified_status == True:
        update_db(og_filepath, work_folder_dest_path, '', orig_hash, ctime, username, mtime, instrument, asset, idproj, filediff, 'modified')

    elif no_action_needed:
        pass

    else:
        print('db not updated, force non-zero' + 1)


else:
    print('older than cutoff date')