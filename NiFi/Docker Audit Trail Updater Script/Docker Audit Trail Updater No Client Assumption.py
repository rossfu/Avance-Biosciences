# Developed by Ross Fu  Jan2023
# For the purpose of creating an audit trail for docker activity in the LIMS DB.
# Only intended to run once in its original form.

import os
import sys
import pandas as pd
import mysql.connector
from datetime import datetime


# Get Audit Trail Folders
file_path_list = []
for root,dirs,files in os.walk(r'\\birch\Operations\Projects'): #Hard Coded Path expected to contain all Docker Audit Trail Files
    for names in files:
            
        if names.endswith('analysis_run_info.txt'):
            file_path_list.append(os.path.join(root,names))
            file_path_list.sort(key=os.path.getmtime)
            print('Script has found ' + str(len(file_path_list)) + ' audit trail files')

num_files = len(file_path_list)
num = 1

print('##############################')
print('#    ' + str(num_files) + ' Docker logs found   #')
print('##############################')







# Establish DB Connection
lims_db = mysql.connector.connect(host='192.168.21.96', user='teracopy', password='T3RA_copy!', database='ngspipelinedb', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor(buffered=True)

lims_db_cursor.execute("SELECT * FROM ngspipelinedb.ngs_analysis_pipeline_activity")
docker_records = lims_db_cursor.fetchall()

for x in docker_records:
    print(x)


client_matches = 0
client_non_matches = 0

proj_id_matches = 0
proj_id_non_matches = 0




date_matches = 0
date_non_matches = 0

machine_matches = 0
machine_non_matches = 0

container_matches = 0
container_non_matches = 0

pipeline_matches = 0
pipeline_non_matches = 0

user_matches = 0
user_non_matches = 0


clients = 0
projects = 0
machines = 0
containers = 0
pipelines = 0
users = 0

print('\nProcessing Files...\n')



    

# Processing the Audit Trail Files (from \\birch\Projects)
for data_file in file_path_list:

    User = None
    Client = None
    Project_ID = None

    
    print('for data file #' + str(num) + '\n')
    print('File Path: ' + data_file)

        


    # Extract File Contents
    with open(data_file, 'r') as file:
        file_contents = file.readlines()

        for line in file_contents:

            if line.startswith('MachineName'):
                machines += 1
                machine = line.split(': ')[-1].strip()


            if line.startswith('ContainerID'):
                containers += 1
                container = line.split(': ')[-1].strip()


                lims_db_cursor.execute("SELECT * FROM ngspipelinedb.ngs_analysis_pipeline_activity WHERE ContainerID = '" + container + "'")
                container_matched_result = lims_db_cursor.fetchone()
                
                if container_matched_result == None:
                    container_non_matches+=1
                else:
                    container_matches+=1


                    
            if line.startswith('PipelineName'):
                pipelines += 1
                pipeline = line.split(': ')[-1].strip()
                
                lims_db_cursor.execute("SELECT * FROM ngspipelinedb.ngs_analysis_pipeline_activity WHERE PipelineName = '" + pipeline + "'")
                pipeline_matched_result = lims_db_cursor.fetchone()
                
                if pipeline_matched_result == None:
                    pipeline_non_matches+=1
                else:
                    pipeline_matches+=1



                    

            if line.startswith('UserName') or line.startswith('AnalyzerName'):
                users += 1
                user = line.split(': ')[-1].strip()

                lims_db_cursor.execute("SELECT * FROM ngspipelinedb.ngs_analysis_pipeline_activity WHERE UserID = '" + user + "'")
                user_matched_result = lims_db_cursor.fetchone()
                
                if user_matched_result == None:
                    user_non_matches+=1
                    print('\n$$$$$$$data file has no user match in LIMS.$$$$$$$n') # No DB Match
                else:
                    user_matches+=1               

                
            if line.startswith('DateTime'):

                date = line.split(': ')[-1]
                date = date[:date.find(next(filter(str.isalpha, date)))-1]
                date = date.strip()
                
                num+=1
                




                    
    # Client
    
    path = data_file.split('\\')
    print('Path:' + str(path) + '\n')
    
    for i in range(5, len(path)): # It is beyond \\birch\operations\project and in some subdirectory which might contain information


        print(path[i])


                
        if 'test' in path[i] == False \
        and 'log' in path[i] == False \
        and 'plasmid' in path[i] == False \
        and 'Project' in path[i] == False \
        and path[i].startswith('GenericID') == False \
        and path[i].startswith('analysis_run_info.txt') == False \
        and 'run' in path[i] == False \
        and path[i][0].isnumeric() == False \
        and 'virus' in path[i] == False \
        and 'ref' in path[i] == False:

            print('%%%')
            Client = path[i]


        if path[i] == 'Sanofi' or path[i] == 'Sangamo':
            Client = path[i]

        if 'istari' in path[i].lower():
            Client = 'Istari'
            
        if 'beam' in path[i].lower() or 'bean' in path[i].lower():
            Client = 'Beam'
            
        if 'eurogentec' in path[i].lower():
            Client = 'Eurogentec'

        if 'graphitebio' in path[i].lower():
            Client = 'GraphiteBio'

        if 'avrobio' in path[i].lower():
            Client = 'AvroBio'



        # Project ID
        if len(path[i]) == 6:
            if path[i].isnumeric():
                Project_ID = path[i]
        if len(path[i]) == 9:
            if path[i][0:3].isnumeric() == False and path[i][3:].isnumeric():
                Project_ID = path[i]
                    






    if Client != None:
        print('\nClient:' + Client)
        client_matches += 1
        clients += 1
    else:
        print('\nClient:N/A')
        client_non_matches += 1
        
    if Project_ID != None:
        print('Project ID:' + Project_ID)
        proj_id_matches += 1
        projects += 1
        
    else:
        print('Project ID:N/A')
        proj_id_non_matches += 1





    # Print outs / DB Updates
    print('Machine:' + machine)
    print('Container:' + container)
    print('Pipeline:' + pipeline)
    print('User:' + user)
    try:
        datetime_date = datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
        print('Date:' + str(datetime_date) + '\n')
    except:
        datetime_date = datetime.strptime(date, '%Y-%m-%d %H:%M')
        print('Date:' + str(datetime_date) + ' (no seconds)\n')




        
    # Does this time stamp exist in our existing DB Records? (Match Files to existing records)
    lims_db_cursor.execute("SELECT * FROM ngspipelinedb.ngs_analysis_pipeline_activity WHERE Date = '" + str(datetime_date) + "'")
    date_matched_result = lims_db_cursor.fetchone()
                
    if date_matched_result == None:
        date_non_matches+=1
        print('\n@@@@@@@data file has no date match in LIMS.@@@@@@\n') # No DB Match
        print('Client Matches = ' + str(client_matches) + '\tClient Non-matches = ' + str(client_non_matches))
        print('Project ID Matches = ' + str(proj_id_matches) + '\tProject ID Non-matches = ' + str(proj_id_non_matches))
        print('User Matches = ' + str(user_matches) + '\tUser Non-matches = ' + str(user_non_matches))
        print('Date Matches = ' + str(date_matches) + '\tDate Non-matches = ' + str(date_non_matches))
        print('Container Matches = ' + str(container_matches) + '\tContainer Non-matches = ' + str(container_non_matches))
        print('Pipeline Matches = ' + str(pipeline_matches) + '\tPipeline Non-matches = ' + str(pipeline_non_matches))
        print('\nMachines = ' + str(machines) + '\tContainers = ' + str(containers) + '\tPipelines = ' + str(pipelines) + '\t\tUsers = ' + str(users) + '\tClients = ' + str(clients) + '\tProjectIDs = ' + str(projects))

        new_record_insert = "INSERT INTO ngspipelinedb.docker_audit_trail (UserID, Date, Project_ID, Client, ComputerID, PipelineName, ContainerID, Files_Path) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
        new_record_insert_vals = (user, datetime_date, Project_ID, Client, machine, pipeline, container, data_file.replace('\\',r'\\'))
        lims_db_cursor.execute(new_record_insert, new_record_insert_vals)
        lims_db.commit()
        
    else:
        date_matches+=1
        matching_result_index = date_matched_result[0] # File matches a record in the DB, get DB index
        print('Client Matches = ' + str(client_matches) + '\tClient Non-matches = ' + str(client_non_matches))
        print('Project ID Matches = ' + str(proj_id_matches) + '\tProject ID Non-matches = ' + str(proj_id_non_matches))
        print('User Matches = ' + str(user_matches) + '\tUser Non-matches = ' + str(user_non_matches))
        print('Date Matches = ' + str(date_matches) + '\tDate Non-matches = ' + str(date_non_matches))
        print('Container Matches = ' + str(container_matches) + '\tContainer Non-matches = ' + str(container_non_matches))
        print('Pipeline Matches = ' + str(pipeline_matches) + '\tPipeline Non-matches = ' + str(pipeline_non_matches))
        print('\nMachines = ' + str(machines) + '\tContainers = ' + str(containers) + '\tPipelines = ' + str(pipelines) + '\t\tUsers = ' + str(users) + '\tClients = ' + str(clients) + '\tProjectIDs = ' + str(projects))
        
        lims_db_cursor.execute("UPDATE ngspipelinedb.docker_audit_trail SET Project_ID = '" + str(Project_ID) + "', Client = '" + str(Client) + "', Files_Path = '" + str(data_file).replace('\\',r'\\') + "' WHERE Date = '" + str(datetime_date) + "'") # Database
        lims_db.commit()
        
    print('-------------------------------------------------------------------------------------------------------------')
                
