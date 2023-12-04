import re
import os
import sys
import ctypes
import shutil
from datetime import datetime

from swinlnk.swinlnk import SWinLnk
import win32com.client


def create_network_shortcut(target_path, shortcut_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target_path
    shortcut.Save()

# create_shortcut.py "NanoDrop1000" "AEM301321"

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


# Set up variables passed from NiFi:
try:
	instrument = sys.argv[1]
	passed_path = sys.argv[2]
	filename = sys.argv[3]
	if passed_path[-1] == '\\':
		passed_path = passed_path[:-1]
	project = os.path.basename(passed_path)
except Exception as e:
	#print(e)
	print("No arguments supplied, try:")
	print("create_shortcuts.py filepath instrument project")
	sys.exit()
with open("C:/nifi-1.21.0/scripts/pylog.txt", "a+") as f:
	f.write(f"passed_path: {passed_path} | ")
	f.write(f"instrument: {instrument} | ")
	f.write(f"project: {project}\n")
	f.write(f"filename: {filename}\n")


# Check if  the file matches ProjectID pattern:
if re.match(r'^[A-Z]{3}\d{6}$', project):
    print(f"{project} matches the pattern.")
    unspecified =  False
else:
    print(f"{project} does not match the pattern.")
    unspecified = True


# If ProjectID is specified, check if it exists in the AllCurrentProjects directory:
if unspecified == False and not os.path.isdir("//45drives/hometree/Operations/Projects/AllCurrentProjects/{0}".format(project)):
    print("create folder in AllCurrentProject: {0}".format("//45drives/hometree/Operations/Projects/AllCurrentProjects/{0}".format(project)))
    # [no proj. dir. we need to copy template into place]:
    shutil.copytree("//45drives/hometree/Operations/Projects/AllCurrentProjects/_ProjectTemplate", "//45drives/hometree/Operations/Projects/AllCurrentProjects/{0}".format(project))
    
    # Get folders to hide:
    folder_path = "//45drives/hometree/Operations/Projects/AllCurrentProjects/{0}/DataLinks".format(project)
    folders = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]

    # Iterate over the folders, hiding each:
    for folder in folders:
        print("hide folders")
        full_path = os.path.join(folder_path, folder)
        set_folder_hidden_status(full_path, hide=True)


#create unspecified folder if it doesn't exist:
if unspecified:
    if not os.path.isdir("//45drives/hometree/Operations/Projects/AllCurrentProjects/unspecified"):
        os.mkdir("//45drives/hometree/Operations/Projects/AllCurrentProjects/unspecified")


# try to create shortcut to files:
try:
    print("create shortcut")
    current_year = datetime.now().year

    #print(f"target_path: {target_path}")
    #print(f"shortcut_path: {shortcut_path}")

    if unspecified:
        #target_path  = fr"\\45drives\hometree\DataBackup\{instrument}\{current_year}\{project}"
        target_path  = fr"\\45drives\hometree\DataArchive\{current_year}\{instrument}\{project}"
        shortcut_path= fr"\\45drives\hometree\Operations\Projects\AllCurrentProjects\unspecified\{project}.lnk"

        # if target_path doesnt exist, project must not be a folder, link to file instead:
        if not os.path.isdir(target_path):
            target_path  = fr"\\45drives\hometree\DataBackup\{instrument}\{current_year}\{filename}"
            shortcut_path= fr"\\45drives\hometree\Operations\Projects\AllCurrentProjects\unspecified\{filename}.lnk"

        create_network_shortcut(target_path, shortcut_path)
    else:
        #target_path  = fr"\\45drives\hometree\DataBackup\{instrument}\{current_year}\{project}"
        target_path  = fr"\\45drives\hometree\DataArchive\{current_year}\{instrument}\{project}"
        shortcut_path= fr"\\45drives\hometree\Operations\Projects\AllCurrentProjects\{project}\DataLinks\{instrument}\{project}.lnk"
        create_network_shortcut(target_path, shortcut_path)
        set_folder_hidden_status(f"//45drives/hometree/Operations/Projects/AllCurrentProjects/{project}/DataLinks/{instrument}", hide=False)
except Exception as e:
    exit(e)