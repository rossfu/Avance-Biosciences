import re #shortcuts script creates shortcuts from 'AllCurrentProjects' ---> 'DataWorkFolder' only
import os
import sys
import glob
import ctypes
import shutil
import hashlib
import datetime
import win32com.client
from swinlnk.swinlnk import SWinLnk

# helper function for creating shortcuts
def create_network_shortcut(target_path, shortcut_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target_path
    shortcut.Save()



# helper function for creating shortcuts
def set_folder_hidden_status(path, hide=True):
    FILE_ATTRIBUTE_HIDDEN = 0x02
    FILE_ATTRIBUTE_NORMAL = 0x80
    try:
        if hide:
            ctypes.windll.kernel32.SetFileAttributesW(path, FILE_ATTRIBUTE_HIDDEN)
            #print(f"The folder at '{path}' is now hidden.")
        else:
            ctypes.windll.kernel32.SetFileAttributesW(path, FILE_ATTRIBUTE_NORMAL)
            #print(f"The folder at '{path}' is now visible.")
    except Exception as e:
        print(f"Error modifying folder attributes: {str(e)}")




def create_shortcut(dest_path, old_path, instrument): 
    # FIND PROJECT ID:
    if instrument == "Tecan":
        fn = os.path.basename(old_path)
        match = re.search(r'\d{6}', fn)

        # Look up 9 character project ID from 6 digits:
        try:
            project = match.group(0)

            import pymysql.cursors
            # Connect to the database:
            lims_db = pymysql.connect(host='192.168.21.69',
            	user=os.environ['LIMS_USER'],
            	password=os.environ['LIMS_PASS'],
            	database='avancelims',
            	port=3306,
            	cursorclass=pymysql.cursors.DictCursor)
            lims_db_cursor = lims_db.cursor()

            print("pID:", project)

            # Start by selecting the current last # from DB:
            sql = f"SELECT * from get_full_idproject WHERE idprojects = '{project}';"
            lims_db_cursor.execute(sql)
            result = lims_db_cursor.fetchone()

            try:
                print(result['full_idprojects'])
                project = result['full_idprojects']
                unspecified = False
            except:
                print("full id not found ") # TODO: delete
                unspecified = True
                project = "NA"

        except:
            print("project id from 6 digits not found. ") # TODO: delete
            unspecified = True
            project = "NA"

    else:
        print("instrument not Tecan. ") # TODO: delete
        match = re.search(r'[A-Z]{3}\d{6}', dest_path)

        if match:
            project = match.group(0)
            print("pID:", project)
            unspecified = False
        else:
            unspecified = True
            project = "NA"

    with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write(f"project: {project}\n")
    print("opened log. ") # TODO: delete

    # If ProjectID is specified, check if it exists in the AllCurrentProjects directory:
    if unspecified == False and not os.path.isdir("//softtest/Quality/hometree/Operations/Projects/AllCurrentProjects/{0}".format(project)):
        print("ProjectID specified. ") # TODO: delete

        with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write("create folder in AllCurrentProject: //softtest/Quality/hometree/Operations/Projects/AllCurrentProjects/{0}\n".format(project))

        # [no proj. dir. we need to copy template into place]:
        try:
            shutil.copytree("//softtest/Quality/hometree/Operations/Projects/AllCurrentProjects/_ProjectTemplate", "//softtest/Quality/hometree/Operations/Projects/AllCurrentProjects/{0}".format(project))
        except:
            # Sometimes we get errors that the directory already exists.. despite there being a direct check for this above.
            pass

        # Get folders to hide:
        folder_path = "//softtest/Quality/hometree/Operations/Projects/AllCurrentProjects/{0}/DataLinks".format(project)
        folders = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
        print("got folders to hide. ") # TODO: delete

        # Iterate over the folders, hiding each:
        for folder in folders:
            full_path = os.path.join(folder_path, folder)
            set_folder_hidden_status(full_path, hide=True)
        print("hid files. ") # TODO: delete
    
    #create unspecified folder if it doesn't exist:
    if unspecified:
        if not os.path.isdir("//softtest/Quality/hometree/Operations/Projects/AllCurrentProjects/unspecified"):
            os.mkdir("//softtest/Quality/hometree/Operations/Projects/AllCurrentProjects/unspecified")
            print("created unspec folder. ") # TODO: delete
    
    # try to create shortcut to files:
    try:
        with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write(f"create shortcut: {project} | {dest_path}\n")
        current_year = datetime.datetime.now().year

        if unspecified:
            with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write(f"unspecified\n")
            target_path  = fr"\\softtest\Quality\hometree\DataWorkFolder\{current_year}\{instrument}\{asset}\{old_path}" #assets are in target path
            shortcut_path= fr"\\softtest\Quality\hometree\Operations\Projects\AllCurrentProjects\unspecified\{instrument}\{asset}\{old_path}.lnk" #organize unspecified
            print(f"target_path: {target_path}")
            print(f"shortcut_path: {shortcut_path}")
            #shortcut_path= fr"\\softtest\Quality\hometree\Operations\Projects\AllCurrentProjects\unspecified\{filename}.lnk"
            os.makedirs(os.path.dirname(shortcut_path), exist_ok=True)

            with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write(f"target_path: {target_path}\n shortcut_path: {shortcut_path}\n old_path: {old_path}\n")
            
            # CHECK IF SHORTCUT EXISTS, create:
            if not os.path.isfile(shortcut_path):
                create_network_shortcut(target_path, shortcut_path)
        else:
            with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write(f"project specified\n")
            target_path  = fr"\\softtest\Quality\hometree\DataWorkFolder\{current_year}\{instrument}\{asset}\{old_path}"
            shortcut_path= fr"\\softtest\Quality\hometree\Operations\Projects\AllCurrentProjects\{project}\DataLinks\{instrument}\{filename}.lnk" #fileame ????? old_path #assets probably not needed here
            print(f"target_path: {target_path}")
            print(f"shortcut_path: {shortcut_path}")
            os.makedirs(os.path.dirname(shortcut_path), exist_ok=True)

            with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write(f"target_path: {target_path}\n shortcut_path: {shortcut_path}\n old_path: {old_path}\n")
            
            # CHECK IF SHORTCUT EXISTS, create:
            if not os.path.isfile(shortcut_path):
                create_network_shortcut(target_path, shortcut_path)
                set_folder_hidden_status(f"//softtest/Quality/hometree/Operations/Projects/AllCurrentProjects/{project}/DataLinks/{instrument}", hide=False)
    except Exception as e:
        exit(e)




# Pass in files 1 at a time from NiFi:
try:
    path = sys.argv[1].rstrip('\\')
    filename = sys.argv[2]
    home_path = sys.argv[3]
    instrument = sys.argv[4].split('\\')[0]
    asset = sys.argv[4].split('\\')[1]
except Exception as e:
    print(e)
    #print("No arguments supplied, try:")
    with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write('No arguments supplied\n')
    sys.exit()


# REMOVE home_path from full_path:

# Input Filepath:
og_filepath = '{0}\{1}'.format(path, filename)

# Work folder root path:
current_year = datetime.datetime.now().year
workfolder_path = rf"\\softtest\Quality\hometree\DataWorkFolder\{current_year}\{instrument}\{asset}"

# Destination path:
og_filepath_stripped = og_filepath.replace(home_path, '')
dest_path = "{0}\{1}".format(workfolder_path, og_filepath_stripped)
with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write(f'og_filepath_stripped: {og_filepath_stripped}\n')
with open('C:\\nifi-1.24.0\\log_shortcut.txt', 'a') as file: file.write(f'dest_path: {dest_path}\n')

# CREATE SHORTCUT for each file:
create_shortcut(dest_path, og_filepath_stripped, instrument)