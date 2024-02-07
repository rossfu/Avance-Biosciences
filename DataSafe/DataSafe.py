# Avance Data Backup
# Developed by Eric Ross Fu in 2021 and 2023 and 2023 again
# ericrossfu@yahoo.com

import os
import sys
import time
import shutil
import pandas as pd
import win32security
import mysql.connector
from hashlib import sha256
from datetime import datetime

import warnings
warnings.filterwarnings("ignore")

from pyspark.sql import SparkSession
from pyspark.sql.functions import col





# Spark Session
spark = SparkSession.builder.appName("DataSafe").config("spark.jars", "C:\spark\spark-3.4.1-bin-hadoop3\jars\mysql-connector-j-8.1.0.jar").getOrCreate()





# Database
lims_db = mysql.connector.connect(host='192.168.21.240', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor()

            





def Auto_Copy(): # Main function


    print('Backing Up The ' + root + '\n\n')

    
    # Set Up Folders in Dest, Gather Paths for every File
    file_path_list = []
    
    for source_root_path, dirs, files in os.walk(source_folder):

        for names in files:

            if '.' + names.split('.')[-1] in set(dtypes): #The file has one of the important file extensions

                file_path_list.append(os.path.join(source_root_path,names))

                
                if not os.path.exists(source_root_path.replace(source_folder,dest_folder)):

                    try:
                        os.makedirs(source_root_path.replace(source_folder,dest_folder))
                    except:
                        continue





    # Vars
    global old_files
    global new_files
    global modified_files
    global error_files
    global orig_hash
    global username
    global ctime
    global mtime
    old_files = 0
    new_files = 0
    modified_files = 0
    error_files = 0

    

    # Process Every File
    for i in range(len(file_path_list)):


        file_path = file_path_list[i]
        dest_path = file_path.replace(source_folder, dest_folder)
        filename = file_path.split('\\')[-1]
        
        global loggy
        loggy = ''
        
        log('File ' + str(i+1) + ' of ' + str(len(file_path_list)) + ' for ' + root + ':                                        ' + file_path)
        
        

        #Comprehensive? PROJID Scrape
        global projid
        projid = ''
        for folder_name in file_path.split('\\')[3:]: #folderpath scrape
            if len(folder_name) == 9 and folder_name[:3].isalpha() == True and folder_name[3:].isnumeric() == True:
                projid = folder_name


        if projid == '':
            for z in range(len(filename)-9): #filename letter by letter scrape
                if filename[z:z+2].isalpha == True and filename[z+2:].isnumeric() == True:
                    projid = filename[z:z+9]





        # Skip locked files
        if os.access(file_path, os.W_OK) == False:
            log('File: ' + file_path.split('\\')[-1] + ' is locked atm\n')
            File_Copying_Records_LIST.append([datetime.now(), file_path, '', '', 'ERROR -- file locked', loggy, '', '', '', root, projid])
            error_files += 1
            continue



        #Get Orig Hash
        orig_hash_process = sha256()
        try: 
            with open(file_path_list[i], 'rb') as original_file:
                for byte_block in iter(lambda: original_file.read(128000),b""):
                    orig_hash_process.update(byte_block)
            original_file.close()

            
        except PermissionError:
            log('Permissions Error')
            File_Copying_Records_LIST.append([datetime.now(), file_path, '', '', 'ERROR -- permission', loggy, '', '', '', root, projid])
            error_files += 1
            continue
        



        # Audit Trail (user / ctime)
        username = GetOwner(file_path)
        ctime = datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
        mtime = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')


        # Hash
        orig_hash = orig_hash_process.hexdigest() # Now we can check the database for existing records of both filepaths and their sha256 hash









# Assessing this file's state of back up --------> initializing action paths
###############################################################################################################################################################################################################################################


        if file_path not in set(db_records['source']): #unseen filepath = new file
            new_file(file_path)

            
        elif orig_hash in set(db_records['hash']): #filepath and hash previously seen = old file
            old_files += 1
            print('No action needed\n')         

        else: 
            modified_file(file_path)


#################################################################################################################################################################################################################################################






                

    # Done Processing All Files in Instrument -- consolidate results (this is for speed/efficiency)
    global File_Copying_Records_DF
    file_count_summary = str(len(file_path_list)) + ' Total Files,\n' + str(old_files) + ' Old Files,\n' + str(new_files) + ' New Files,\n' + str(modified_files) + ' Modified Files,\n' + str(error_files) + ' Error Files'

    
    if len(File_Copying_Records_LIST) == 0:
        print('No new records')
        File_Copying_Records_DF = spark.createDataFrame([[datetime.now(), '','','','complete', 'No New Files Found \n' + file_count_summary , '','','', root, '']], ['time','source','dest','hash','summary','log','owner','ctime', 'mtime', 'root', 'proj'])
        File_Copying_Records_DF.show()
        pyspark_to_mysql()
        
    else:
        File_Copying_Records_LIST.append([datetime.now(), '', '', '', 'complete', file_count_summary, '', '', '', root, ''])
        rdd = spark.sparkContext.parallelize(File_Copying_Records_LIST)

        File_Copying_Records_DF = spark.createDataFrame(rdd, ['time','source','dest','hash','summary','log','owner','ctime', 'mtime', 'root', 'proj'])
        File_Copying_Records_DF.show()
        pyspark_to_mysql()
        

    print('Database Updated')





    # Report Results    
    print(root + ' is Done.\n')
    print('Old Files: ' + str(old_files) + '\nNew Files: ' + str(new_files) + '\nModified Files: ' + str(modified_files) + '\nError Files: ' + str(error_files) + '\n\n\n')









def new_file(file_path):
    
    global new_files
    new_files += 1

    copy_and_hash_check(file_path, dest_folder + file_path.replace(source_folder,''), 'new')
    





def modified_file(file_path):



    
    lims_db_cursor.execute("SELECT source, hash FROM t5_datasafe WHERE source = '" + file_path.replace('\\',r'\\') + "'")
    previous_versions = lims_db_cursor.fetchall()


    

    num_previous_versions = len(previous_versions)
    current_version = num_previous_versions + 1
    log('version ' + str(current_version))

    global modified_files
    modified_files += 1

    copy_and_hash_check(file_path, dest_folder + '\\'.join(file_path.split('\\')[:-1]).replace(source_folder,'') + '\\' + 'v' + str(current_version) + ' --- ' + file_path.split('\\')[-1], 'modified')






def copy_and_hash_check(source, dest, summary):
    

    # Copy
    try:
        shutil.copy2(source, dest)
        log('file copied to: ' + dest)

    except:

        try:
            log('error with copying')

            if os.path.exists(dest) == False:
                os.mkdirs(dest)

            shutil.copy2(source, dest)

        except:

            log('error with copying2')
            return




    # Making sure copied file has same hash:
    copyfile_hash_process = sha256()
    with open(dest, 'rb') as copy_file:
        for byte_block in iter(lambda: copy_file.read(128000),b""):
            copyfile_hash_process.update(byte_block)
        copy_file_hash = copyfile_hash_process.hexdigest()
        copy_file.close()

    
    log('source file hash: '+ orig_hash)
    log('dest file hash:   '+ copy_file_hash + '\n\n')



    #Match Check
    if orig_hash != copy_file_hash: #Error

        log('Hash mismatch: ' + copy_file_hash)
        os.remove(dest)
        log('File was removed from back up')

        
        File_Copying_Records_LIST.append([datetime.now(), source, dest, orig_hash, 'ERROR -- hash mismatch', loggy, username, ctime, mtime, root, projid])

        global error_files
        error_files += 1

        
    # Transfer Successful
    else: 
        File_Copying_Records_LIST.append([datetime.now(), source, dest, orig_hash, summary, loggy, username, ctime, mtime, root, projid])






def log(msg):
    global loggy
    loggy += msg + '\n'
    print(msg)




def GetOwner(filepath):
    try:
        f = win32security.GetFileSecurity(filepath, win32security.OWNER_SECURITY_INFORMATION)
        (username, domain, sid_name_use) =  win32security.LookupAccountSid(None, f.GetSecurityDescriptorOwner())
    except:
        username = 'N/A'
    return username






def pyspark_to_mysql():

    global File_Copying_Records_DF
    File_Copying_Records_DF.sort('time').write \
    .mode("append") \
    .format("jdbc") \
    .option("driver","com.mysql.cj.jdbc.Driver") \
    .option("url", "jdbc:mysql://192.168.21.240:3306/avancelims") \
    .option("dbtable", "t5_datasafe") \
    .option("user", "efu") \
    .option("password", "M0nkey!!") \
    .save()



def load_all_previous_records():

    # Load all previous records
    global db_records


    
    lims_db_cursor.execute("SELECT source, hash FROM t5_datasafe")
    history = lims_db_cursor.fetchall()
    db_records = pd.DataFrame(history, columns = ('source', 'hash'))


##    db_records = spark.read \
##        .format("jdbc") \
##        .option("driver","com.mysql.cj.jdbc.Driver") \
##        .option("url", "jdbc:mysql://191.168.1.240:3306/avancelims") \
##        .option("query", "SELECT source,hash FROM t5_pyspark_autocopy_records") \
##        .option("user", "efu") \
##        .option("password", "M0nkey!!") \
##        .load()
##
##
##    db_records.show()
##
##    import time
##
##    print(db_records.dtypes)
##
##
##    print(db_records.filter(db_records.source == '\\GELDOCXR-1985\Bio-Rad\Image Lab\Data\Agarose Gel\Qualification\Gel Doc 1986\Seongman Kim 2022-02-08_15h54m51s.scn').count())
##    print(db_records.filter(col("source") == '\\\\GELDOCXR-1985\\Bio-Rad\\Image Lab\\Data\\Agarose Gel\\Qualification\\Gel Doc 1986\\Seongman Kim 2022-02-08_15h54m51s.scn').count())
##
##    print('is in')
##    listy = ['\\GELDOCXR-1985\Bio-Rad\Image Lab\Data\Agarose Gel\Qualification\Gel Doc 1986\Seongman Kim 2022-02-08_15h54m51s.scn']
##    listyy = ['\\\\GELDOCXR-1985\\Bio-Rad\\Image Lab\\Data\\Agarose Gel\\Qualification\\Gel Doc 1986\\Seongman Kim 2022-02-08_15h54m51s.scn']
##    print(db_records.filter(db_records.source.isin(listy)).count())
##    print(db_records.filter(col("source").isin(listyy)).count())
##    time.sleep(2000)


# Get history of file back ups at start of run
load_all_previous_records()







# Email Request
db_insert = "INSERT INTO auto_email (receiver, subject, body, parent_process) VALUES (%s,%s,%s,%s)"
db_insert_vals = ('eric.fu@avancebio.com', 'DataSafe is Running', '', 'DataSafe')
lims_db_cursor.execute(db_insert, db_insert_vals)
lims_db.commit()



### Send Email
##import smtplib
##from email.mime.multipart import MIMEMultipart
##from email.mime.text import MIMEText
##
### Email configuration
##sender_email = "nconnector@avancebio.com"
##receiver_email = 'eric.fu@avancebio.com'
##subject = 'DataSafe'
##body = 'no need for auto email'
##
##
### SMTP server configuration (for Gmail)
##smtp_server = "smtp.office365.com"
##smtp_port = 587
##smtp_password = ">_1d#2qP9M3]}92m"
##
##
### Create the MIME object
##message = MIMEMultipart()
##message["From"] = sender_email
##message["To"] = receiver_email
##message["Subject"] = subject
##
##
### Attach the body
##message.attach(MIMEText(body, "plain"))
##
##
### Establish a connection with the SMTP server
##with smtplib.SMTP(smtp_server, smtp_port) as server:
##    # Start the TLS encryption
##    server.starttls()
##    
##    # Login to the email account
##    server.login(sender_email, smtp_password)
##    
##    # Send the email
##    server.sendmail(sender_email, receiver_email, message.as_string())
##









# Config
##instrument_sources = spark.read \
##    .format("jdbc") \
##    .option("driver","com.mysql.cj.jdbc.Driver") \
##    .option("url", "jdbc:mysql://191.168.1.240:3306/avancelims") \
##    .option("dbtable", "t5_autocopy_config__v9") \
##    .option("user", "efu") \
##    .option("password", "M0nkey!!") \
##    .load()


lims_db_cursor.execute("SELECT * FROM gmp_data_sources")
instrument_sources = lims_db_cursor.fetchall()

print('Task List:')
for i in range(len(instrument_sources)):
    print(instrument_sources[i][0], instrument_sources[i][1])
print('\n\n\n')
time.sleep(5)





# globalize these variables, less to track as information moves from function to function
global File_Copying_Records_LIST
global File_Copying_Records_DF
global source_folder
global dest_folder
global db_records
global orig_hash
global dtypes
global root
task_number = 0

starting_task = task_number





# Create 'Program Start' Record
File_Copying_Records_DF = spark.createDataFrame([[datetime.now(), '','','','start', '' , '','','', '','']], ['time','source','dest','hash','summary','log','owner','ctime', 'mtime', 'root', 'proj'])
pyspark_to_mysql()





# Program Execution Manager
while True:



    # End Program Condition
    if task_number == len(instrument_sources) - 1: # No Miseq
        task_number = 0
        File_Copying_Records_DF = spark.createDataFrame([[datetime.now(), '','','','end', '' , '','','', '','']], ['time','source','dest','hash','summary','log','owner','ctime', 'mtime', 'root', 'proj'])
        pyspark_to_mysql()
        os._exit(0)


##    if instrument_sources[task_number][6] == 'Miseq': #skip Miseq
##        task_number += 1
        

    try:
        # Instrument Data Path
        source_folder = instrument_sources[task_number][2]


        
        # Create a new list to store records for this instrument
        File_Copying_Records_LIST = []



        # Root
        root = instrument_sources[task_number][6]


        


        # DESTINATION FOLDER FOR BACKUPS
        if root == 'DataSafeTestFolder':
            dest_folder = r'\\45drives\hometree\TempInstrumentBackupFolder\DataSafe' + '\\' + root + '\\' + '\\'.join(instrument_sources[task_number][2].split('\\')[4:])
            
        else:
            dest_folder = r'\\45drives\hometree\TempInstrumentBackupFolder\DataSafe' + '\\' + root + '\\' + '\\'.join(instrument_sources[task_number][2].split('\\')[3:])




        # Which File Extensions matter?
        dtypes = instrument_sources[task_number][5].split('; ')

            
            
        # Run Program
        Auto_Copy() # Lets analyze the source parent folder
        task_number += 1






    except BaseException as e:
        
        import traceback
        exc_type, exc_value, exc_traceback = sys.exc_info()

        log('\n\n' + str(datetime.now()) + '\n\n')
        log('DEBUG INFO:')
        log(source_folder)
        log(dest_folder)

        log('\n\n\n### ERROR ###                                    ' + str(datetime.now()))
        log('source folder: ' + source_folder)
        log('dest folder: ' + dest_folder)
        log('\n\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))

        try:
            File_Copying_Records_LIST.append([datetime.now(),'','','','end', loggy, '','','', root,''])
            rdd = spark.sparkContext.parallelize(File_Copying_Records_LIST)
            File_Copying_Records_DF = spark.createDataFrame(rdd, ['time','source','dest','hash','summary','log','owner','ctime', 'mtime', 'root', 'proj'])

            pyspark_to_mysql()

            print('Stopping')
            break

        except:
            print('error log failed')
