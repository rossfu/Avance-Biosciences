import re
import os
import sys
import shutil
import win32security #pywin32 
import pymysql.cursors
from datetime import datetime
import xml.etree.ElementTree as ET
sys.path.append("/usr/local/lib/python3.10/dist-packages")

# \\GELDOC-XR\GelDoc XR-1526 Data\AAV301746
# AAV301746_6284-C2_10.8_Plate 1_Induced_071123_BV.scn

try:
    path = sys.argv[1].rstrip('\\') #+ '\\'
    filename = sys.argv[2]
    print('\npath: ' + path)
    print('filename: ' + filename)
except Exception as e:
    #print(e)
    print("No arguments supplied, try:")
    print("geldoc.py file-path file-name")
    sys.exit()




# only run code for SCN and SSCN files:
if ".scn" in filename or ".sscn" in filename:
    # Set Filepath:
    og_filepath = '{0}\{1}'.format(path, filename)
    print('filepath: ' + og_filepath)


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

    # WHICH FILE EXTENSION?
    extension = filename.split('.')[-1]

    print("extension: {0}".format(extension))

    # FIND PROJECT ID:
    match = re.search(r'[A-Z]{3}\d{6}', filename)
    if match:
        project_id = match.group(0)
        print(project_id)
    else:
        project_id = "N/A"

    # pull audit data from .scn files:

    # start parsing XML at <root>
    # cut file at </root>
    cnt = 0
    xml_str = ""
    with open(og_filepath, 'r', errors='ignore') as file:
        for line in file:
            if cnt > 12:
                xml_str+=line

                if "</root>" in line:
                    break
            else:
                cnt=cnt+1

    #process XML:
    print(xml_str)
    # put all XML into a blob?

    # insert into DB:
    try:
        # Connect to the database:
        connection = pymysql.connect(host='192.168.21.69',
        user=os.environ['LIMS_USER'],
        password= os.environ['LIMS_PASS'],
        #user='nconnector',
        #password= os.environ.get('ADB'),
        database='avancelims',
        port=3306,
        cursorclass=pymysql.cursors.DictCursor)

        with connection.cursor() as cursor:
            sql = "INSERT INTO `t6_geldoc` (`User`, `Filepath`, `Project_ID`, `ctime`, `mtime`, `xml_blob`) VALUES (%s, %s, %s, %s, %s, %s)"
            cursor.execute(sql, (metadata_user, og_filepath, project_id, ctime, mtime, xml_str))
            connection.commit()
    except Exception as e:
        print(e)
        print("boo")