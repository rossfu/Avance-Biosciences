import os
import sys
import csv
import time
import math
import shutil
import statistics
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side



print(time.asctime())

#Set the parameters for the Program


Excel_File_Name = 'AXH210921_SCS 19.01.xlsm'        
Data_Directory_Path = r'C:\Users\efu\Desktop\Test\AXH219821'
Project_Output_Path = r'C:\Users\efu\Desktop\Test\Project_Output'
Excel_Template_Path = r'C:\Users\efu\Desktop\Test' + '\\' + Excel_File_Name
SlopeMin = -3.7
SlopeMax = -3.1
LOD = 2.5
LOQ = 25
R2 = .998
#Real_Project_Output_Path = r'O:\5 GlobalShare\qPCR\Ross\Project_Output'



#SlopeMin = float(sys.argv[3])
#SlopeMax = float(sys.argv[4])
#LOD = float(sys.argv[5])
#R2 = float(sys.argv[6])

#SlopeMin = float(SlopeMin)
#SlopeMax = float(SlopeMax)
#LOD = float(LOD)
#R2 = float(R2)




#############################################################################################################
###################################     Table Of Contents    ################################################
#############################################################################################################
#                                                                                                           #
#   Set the Parameters                                                                                      #
#   Table Of Contents                                                                                       #
#   Form10 Functions                                                                                        #
#   Find and Sort Raw Data Files (.txt)                                                                     #
#   Set Up Informational Files (.txt)                                                                       #
#   Initialize variables outside the data file loop                                                         #
#   Set Up Excel                                                                                            #
#   Eliminate the 1 row off issue and Welcome the next batch                                                #
#   Iterate over every file                                                                                 #
#   Classify the Data file                                                                                  #
#                                                                                                           #
###################################   NOT QUANT STUDIO FILE   ###############################################
#                                                                                                           #
#   Clean Raw Data                                                                                          #
#   Write Raw Data to Excel                                                                                 #
#   Form 10 Table                                                                                           #
#   Style Formatting                                                                                        #
#   Plate Report                                                                                            #
#   Plate Report Table                                                                                      #
#   ReRun Status (Plates)                                                                                   #
#   Re_Run_Samples Report                                                                                   #
#   AC Report                                                                                               #
#                                                                                                           #
#####################################   QUANT STUDIO FILE   #################################################
#                                                                                                           #
#   Find the actual data                                                                                    #
#   Process File                                                                                            #
#   Create Modified Quant Studio File                                                                       #
#   Write Data to Excel                                                                                     #
#   Form 10 Table                                                                                           #
#   Style Formatting                                                                                        #
#   Plate Report                                                                                            #
#   Plate Report Table                                                                                      #
#   ReRun Status (Plates)                                                                                   #
#   Re_Run_Samples Report                                                                                   #
#   AC Report                                                                                               #
#                                                                                                           #
#############################################################################################################
#                                                                                                           #
#   Write Re Run Samples                                                                                    #
#   Set Column Width for Readability                                                                        #
#   Save!                                                                                                   #
#   Run the Program                                                                                         #
#                                                                                                           #
#############################################################################################################
#############################################################################################################
#############################################################################################################








########################################  FORM 10 FUNCTIONS

def Get_Sample_Names(form10):
    Sample_Name_List = []
    for x in range(1,len(form10)):
        if (form10['Sample Name'][x]) == (form10['Sample Name'][x-1]):
            Sample_Name_List.append(form10['Sample Name'][x])
    return Sample_Name_List


def QtyMean_Dataframe(form10):
    Qty_Means = pd.DataFrame()
    QtyMean_List = []
    Sample_names_list = []
    for x in range(1,len(form10)-1):
        if (form10['Sample Name'][x]) == (form10['Sample Name'][x-1]) and str(form10['Sample Name'][x]).startswith('S') == False and str(form10['Sample Name'][x]).startswith('N') == False:
            QtyMean_List.append(float(form10['Quantity'][x])+float(form10['Quantity'][x-1])/ 2)
            Sample_names_list.append(form10['Sample Name'][x])
        if (form10['Sample Name'][x]) == 'NTC' and (form10['Sample Name'][x-1]) == 'NTC':
            if form10['Quantity'][x] == 'Undetermined' or math.isnan(form10['Quantity'][x]):
                Qty_Means['NTC'] = 'Undetermined'
            else:
                Qty_Means['NTC'] = (form10['Quantity'][x]+form10['Quantity'][x-1])/2
    Qty_Means['Sample Names'] = Sample_names_list
    Qty_Means['Qty Mean'] = QtyMean_List
    return Qty_Means


def SPK_Ct_dataframe(form10):
    SPK_Cts = pd.DataFrame()
    SPK_names = []
    SPK_values = []
    SPK_value_differences = []
    Inhibition_true_false = []
    SPK_control_values = []
    #SPK Control Average
    for x in range(1, len(form10)):
        if form10['Sample Name'][x] == 'SPK CONTROL' and form10['Sample Name'][x-1] == 'SPK CONTROL':
            if form10['Ct'][x] == 'Undetermined' or form10['Ct'][x-1] == 'Undetermined':
                SPK_avg = 'Undetermined'
            else:
                #SPK_control_values = list(SPK_control_values)
                SPK_control_values.append(form10['Ct'][x-1])
                SPK_control_values.append(form10['Ct'][x])
                SPK_control_values = pd.to_numeric(SPK_control_values)
                SPK_avg = sum(SPK_control_values) / len(SPK_control_values)
    #Extracting +SPK Information
    for x in range(0,len(form10)): 
        if form10['Sample Name'][x].endswith('spk') or form10['Sample Name'][x].endswith('SPK'):
            SPK_names.append(form10['Sample Name'][x])
            SPK_values.append(form10['Ct'][x])
            if SPK_avg == 'Undetermined':
                SPK_Cts['SPK Value Differences'] = 'Undetermined'
                Inhibition_true_false.append('True')   
            else:
                SPK_value_difference_calculation = abs(SPK_avg - float(form10['Ct'][x]))
                SPK_value_differences.append(SPK_value_difference_calculation)
                if SPK_value_difference_calculation > 1:
                    Inhibition_true_false.append('True')   
                else:
                    Inhibition_true_false.append('False') 
    #Displaying Results in DataFrame
    SPK_Cts['Sample Name'] = SPK_names
    SPK_Cts['SPK Values'] = SPK_values
    SPK_Cts['SPK Control Value'] = SPK_avg
    SPK_Cts['Inhibition?'] = Inhibition_true_false
    if SPK_avg != 'Undetermined':
        SPK_Cts['SPK Value Differences'] = SPK_value_differences
    return SPK_Cts


def Replicate_Diff_Dataframe(form10, LOQ):
    Replicate_Diff_Dataframe = pd.DataFrame()
    Ct_diff_list = []
    Sample_list = []
    for x in range(1,len(form10)):  
        if (form10['Sample Name'][x]) == (form10['Sample Name'][x-1]):
            Sample_list.append(form10['Sample Name'][x])  
            #Diff type and above LOQ
            if form10['Ct'][x].isalpha() != form10['Ct'][x-1].isalpha():
                if form10['Quantity'][x] > LOQ or form10['Quantity'][x-1] > LOQ:  
                    Ct_diff_list.append('PROBLEM: 1 Undetermined & 1 Ct and Quantity > LOQ')
            #Diff type and below LOQ
                if form10['Quantity'][x] <= LOQ or form10['Quantity'][x-1] <= LOQ:
                    Ct_diff_list.append('1 Undetermined and 1 Ct but Quantity < LOQ')
            #2 undetermined
            elif form10['Ct'][x].isalpha() == True and form10['Ct'][x-1].isalpha() == True:
                Ct_diff_list.append('Undetermined')
            #Large diff and above LOQ
            elif abs(float(form10['Ct'][x]) - float(form10['Ct'][x-1])) > 1 and form10['Qty Mean'][x] > LOQ:
                Ct_diff_list.append('PROBLEM: LARGE DIFFERENCE')                             
            else:
                Ct_diff_list.append(float(form10['Ct'][x]) - float(form10['Ct'][x-1]))   
    Replicate_Diff_Dataframe['Sample Name'] = Sample_list
    Replicate_Diff_Dataframe['Ct diff'] = Ct_diff_list                                 
    return Replicate_Diff_Dataframe


def AC_report(form10): #Returns the 2 AC's in a plate
    AC_report_df = pd.DataFrame()
    for i in range(1, len(form10)):
        if form10['Sample Name'][i] == 'AC' and form10['Sample Name'][i-1] == 'AC':   
            AC_report_df = AC_report_df.append({'Ct1':form10['Ct'][i-1], 'Ct2':form10['Ct'][i], 'Qty1':form10['Quantity'][i-1], 'Qty2':form10['Quantity'][i], 'QtyMean': ((form10['Quantity'][i-1]+form10['Quantity'][i])/2)}, ignore_index = True)
    return AC_report_df
    

def healthy_samples_DF(form10):
    return 'in development'

def STD_info_DF(form10):

    std_df = pd.DataFrame()
    std_range = []
    std_ct1 = []
    std_ct2 = []
    statement = ''
    
    for sample in range(1, len(form10)):
        if form10['Sample Name'][sample].startswith('STD'):
            if form10['Sample Name'][sample] == form10['Sample Name'][sample-1]:
                std_range.append(form10['Quantity'][sample])
                std_ct1.append(form10['Ct'][sample-1])
                std_ct2.append(form10['Ct'][sample])
            
    
    for i in range(len(std_range)):     #counts over each standard point
        tmp_x = []                      #a tmp list that will hold the x values for slope. The x value will be the log base 10 of the quantity (see below)
        tmp_y = []                      #a tmp list that will hold the y values
        for x in range(len(std_range)): #determines which point will be removed
            if x!=i:
                tmp_x.append(np.log10(std_range[x]))
                tmp_x.append(np.log10(std_range[x]))
                tmp_y.append(std_ct1[x])
                tmp_y.append(std_ct2[x])

        tmp_y1 = []
        for y in tmp_y:
            tmp_y1.append(float(y))
                
        slope, intercept = np.polyfit(tmp_x, tmp_y1, 1)
        
        if SlopeMin <= slope <= SlopeMax:
            statement = "Remove STD" + str(i+1) + ": " + str(slope)


        std_df['std_range'] = std_range
        std_df['ct1'] = std_ct1
        std_df['ct2'] = std_ct2
        std_df['statement'] = statement

    return std_df


############################################################################################################################################################################









def execute(Data_Directory_Path, Excel_File_Name, Project_Output_Path, SlopeMin, SlopeMax, LOD, LOQ, R2):

#Find and Sort .TXT Raw Data Files
        
    excel_file_exists = False
    backup_file_exists = False
    Need_modified_quant_folder = True
    Need_Project_Output_folder = True
    Plate_Info_Exists = False
    Re_Run_Sample_Info_Exists = False
    
    file_path_list = []


    #Data Directory Folder
    for root,dirs,files in os.walk(Data_Directory_Path):

        for names in files:
                
            if names.endswith('.txt'):
                file_path_list.append(os.path.join(root,names))
                file_path_list.sort(key=os.path.getmtime)


    #Project Output Folder
    for root, dirs, files in os.walk('\\'.join(Project_Output_Path.split('\\')[:-1])):

        if Project_Output_Path.split('\\')[-1] in dirs:
            Need_Project_Output_folder = False
                
    if Need_Project_Output_folder == True:
        os.mkdir(Project_Output_Path)


    for root, dirs, files in os.walk(Project_Output_Path):

        if names == Excel_File_Name:
            excel_file_exists = True

        if 'Modified_Quant_data_files' in dirs:
            Need_modified_quant_folder = False


        for names in files:
            
            if names.startswith('Form10_BackUp'):
                Backup_Filename = os.path.join(root,names)
                backup_file_exists = True

            if names.split('\\')[-1] == 'Plate_Information.txt':
                Plate_Info_Exists = True
                
            if names.split('\\')[-1] == 'Re_Run_Sample_Info.txt':
                Re_Run_Sample_Info_Exists = True

    if Need_modified_quant_folder == True:
        os.mkdir(Project_Output_Path + '\\Modified_Quant_data_files')


    
    file_count = 0 


#Set Up Informational .txt Files
    

    if Plate_Info_Exists == False: #Create Plate_Information.txt
        with open(Project_Output_Path + '\\' + 'Plate_Information.txt', 'w', newline='') as Plate_Information:
            Plate_Information.write('Plate\tDate_Processed\tFile_Length\tPlate_Failed\tSamples_Failed\tAC_qty1\tAC_qty2')
        Plate_Information.close()

    Plate_Info_DF = pd.read_csv(Project_Output_Path + '\\' + 'Plate_Information.txt', sep = '\t', lineterminator = '\r', header = 0)
    
    files_before = len(Plate_Info_DF['Plate'])
    

    #Read File into Re_Run_Sample_Info_DATAFRAME, append dataframe, save as file:
    
    if Re_Run_Sample_Info_Exists == False: #Create Re_Run_Sample_Info.txt
        with open(Project_Output_Path + '\\' + 'Re_Run_Sample_Info.txt', 'w') as Re_Run_Sample_Info:
            Re_Run_Sample_Info.write('Need_Re_Run_Plate\tNeed_Re_Run_Samples\tReason_Re_Run\tRe_Run_Plate\tRe_Run_Samples')
        Re_Run_Sample_Info.close()

    Re_Run_Sample_Info_Dataframe = pd.read_csv(Project_Output_Path + '\\' + 'Re_Run_Sample_Info.txt', sep = '\t', lineterminator = '\n', header = 0)    
    

    
#Initialize variables outside the data file loop
    
    row_where_file_ends = len(Plate_Info_DF.File_Length) * 109
    currentRow = 1 + row_where_file_ends
    
    plates_passed = 0 #For AC report

    AC_qty_means = []

    for plate in Plate_Info_DF['Plate_Failed']:
        if plate == False:
            plates_passed += 1

    #healthy_samples_DF = pd.DataFrame(columns = ['Plate', 'Sample Names', 'Qty Mean'])



    #Only Recent Files get processed
            
    file_path_list = [i for i in file_path_list if i.split('\\')[-1][:-4] not in list(Plate_Info_DF['Plate'])]
    file_path_list = [i for i in file_path_list if i.split('\\')[-1] != 'Plate_Information.txt']
    file_path_list = [i for i in file_path_list if i.split('\\')[-1] != 'Re_Run_Sample_Info.txt']
    file_path_list = [i for i in file_path_list if i.split('\\')[-1].startswith('MODIFIED_QUANTSTUDIO_FILE') == False]

    
#Set Up Excel
        
    workbook_path = Project_Output_Path + '\\' + Excel_File_Name

    if excel_file_exists == True:
        wb = load_workbook(workbook_path, keep_vba=True)

    else:
        #SCS TEMPLATE LOCATION (Not DATA_DIRECTORY_PATH)
        wb = load_workbook(Excel_Template_Path, keep_vba=True)


    if 'Form10_Raw Data' in wb.sheetnames:
        ws = wb['Form10_Raw Data']
    else:
        ws = wb.create_sheet('Form10_Raw Data')

    if 'Plate Report' in wb.sheetnames:
        del wb['Plate Report']

    plate_report_ws = wb.create_sheet('Plate Report')

    if 'Re Run Samples' in wb.sheetnames:
        re_run_samples_ws = wb['Re Run Samples']
    else:
        re_run_samples_ws = wb.create_sheet('Re Run Samples')

    if 'AC Report' in wb.sheetnames:
        ac_report_ws = wb['AC Report']
    else:
        ac_report_ws = wb.create_sheet('AC Report')


    
#Create Back-Up Project File -- too mach backups
    
    now = datetime.now()

#    if excel_file_exists == True: #Preserve Current File
#        if backup_file_exists == True:
#            os.rename(Backup_Filename, Data_Directory_Path + '\\' + 'Form10_BackUp_' + Excel_File_Name[:-5] + '_' + str(now.strftime('%m-%d-%Y_time_%H_%M')) + '.xlsm')
#            shutil.copyfile(Data_Directory_Path + '\\' + Excel_File_Name, Data_Directory_Path + '\\' + 'Form10_BackUp_' + Excel_File_Name[:-5] + '_' + str(now.strftime('%m-%d-%Y_time_%H_%M')) + '.xlsm')
#        else:
#            shutil.copyfile(Data_Directory_Path + '\\' + Excel_File_Name, Data_Directory_Path + '\\' + 'Form10_BackUp_' + Excel_File_Name[:-5] + '_' + str(now.strftime('%m-%d-%Y_time_%H_%M')) + '.xlsm')
      

#Eliminate the 1 row off issue and Welcome the next batch
    
    if row_where_file_ends > 0:
        ws['A' + str(row_where_file_ends)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
        ws['B' + str(row_where_file_ends)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
        ws['C' + str(row_where_file_ends)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
        ws['D' + str(row_where_file_ends)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
        ws['E' + str(row_where_file_ends)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
        ws['F' + str(row_where_file_ends)] = 'Files Processed ' + str(now.strftime('%Y-%m-%d %H:%M')) + ' and after'

#Iterate over every file


    for data_file in file_path_list:



              
#Classify the Data file
        
        test = pd.read_csv(data_file, sep = '\t', skip_blank_lines = False, usecols = [0])




        if test.columns[0] == 'SDS 2.4': #NOT QUANT

            print('')
            print('Processing File: ' + data_file.split('\\')[-1])

            form10 = pd.read_csv(data_file, sep = '\t', lineterminator='\n', header = 10)
            form10_raw_extra = pd.read_csv(data_file, sep = '\t', lineterminator='\n', nrows = 9)

        
#Clean Raw Data
        
            form10.drop(['\r','FOS','HMD','LME','EW','BPR','NAW','HNS','HRN','EAF','BAF','TAF','CAF'], axis = 1, inplace = True)


#Write Raw Data to Excel

            open_data_file = open(data_file, 'r')

            currentRow = (file_count + files_before) * 109 + 1
            currentCell = None
            currentCol = 2

            while True:
            
                line = open_data_file.readline()

                if not line:
                    break

                if line in ['\n', '\r', '\r\n'] or line.startswith('NAP'):
                    continue

                fields = line.split("\t") #separate line by tabs into up to 34 cells
                for fieldIndex in range(len(fields)): #copy up to 21 cells in each line
                
                    currentCell = ws.cell(row=currentRow, column=currentCol)

                    if fields[fieldIndex]: #only try to copy field if it has a value
                        if (fields[fieldIndex][0].isdigit() or fields[fieldIndex].startswith("-")): #set cell value as a float if it begins with a number or a "-"

                            try: #copy value as a float
                                currentCell.value = float(fields[fieldIndex])

                            except: #copy value as text
                                currentCell.value = fields[fieldIndex]

                        else: #copy value as text
                            currentCell.value = fields[fieldIndex]

                    currentCol += 1

                    if (currentCol == 23 or fieldIndex == len(fields)-1): #move to the next row in sheet after 21 columns or at the end of the row
                        currentCol = 2
                        currentRow += 1
                        break

            open_data_file.close()

            print(data_file + " is written")
            

#Form 10 Table
            
        #Manage Paste Offset

            Paste_row = row_where_file_ends + 20

            Sample_Name_List = Get_Sample_Names(form10)

            AC_exist = False
            for name in Sample_Name_List:
                if name.startswith('AC'):
                    Paste_row -= 1
                    AC_exist = True
                    break
            

            #Paste Calculation Header
            
            ws['X' + str(Paste_row-1)] = 'Sample Name'
            
            ws['Y' + str(Paste_row-1)] = 'Qty Mean'
            
            ws['Z' + str(Paste_row-1)] = 'Replicate Difference'
            
            ws['AA' + str(Paste_row-1)] = '(+) SPK Value'
            ws['AB' + str(Paste_row-1)] = 'SPK Difference'
            ws['AC' + str(Paste_row-1)] = 'Inhibition?'
            
            #Paste Sample Names
            
            #Sample_Name_List = form10func.Get_Sample_Names(form10)
            for name in Sample_Name_List:
                ws['X' + str(Paste_row)] = name
                Paste_row += 1
            

            #QuantityMean Calculation and write
                
            Paste_row = row_where_file_ends + 20 #Reset Paste_Row position

            if AC_exist == True:
                Paste_row -= 1
             
            QuantityMeans = QtyMean_Dataframe(form10)
            for Qty_mean_value in QuantityMeans['Qty Mean']:
                ws['Y' + str(Paste_row)] = Qty_mean_value
                Paste_row += 1
        

            #Replicate Diff
            Paste_row = row_where_file_ends + 20 #Reset Paste_Row position

            if AC_exist == True:
                Paste_row -= 1
                
            Replicate_Differences = Replicate_Diff_Dataframe(form10, LOQ)
            for value in Replicate_Differences['Ct diff']:
                ws['Z' + str(Paste_row)] = value
                Paste_row+=1
                

            #SPK Ct
            Paste_row = row_where_file_ends + 20 #Reset Paste_Row position

            if AC_exist == True:
                Paste_row -= 1
                
            if Sample_Name_List[0].startswith('A'):
                Paste_row += 1
            
            SPK_Ct = SPK_Ct_dataframe(form10) #SPK_Ct will be used for a list of C_Samples
            for x in range(len(SPK_Ct['SPK Values'])):
                ws['AA' + str(Paste_row)] = SPK_Ct['SPK Values'][x]
                ws['AB' + str(Paste_row)] = SPK_Ct['SPK Value Differences'][x]
                ws['AC' + str(Paste_row)] = SPK_Ct['Inhibition?'][x]
                Paste_row += 1


            #SPK Control Information
            Paste_row += 2
            ws['AA' + str(Paste_row)] = SPK_Ct['SPK Control Value'][0]
            ws['AA' + str(Paste_row)].font = Font(size = 13, bold = True)

#Style Formatting
            
            #ws['U' + str(row_where_file_ends+1)] = 1 + file_count + files_before
            #ws['U' + str(row_where_file_ends+1)].font = Font(size = 16, bold = True)
            #ws['U' + str(row_where_file_ends+1)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
            #ws['U' + str(row_where_file_ends+1)].alignment = Alignment(horizontal = 'center')

            row_where_file_ends = row_where_file_ends + 109

###############################################################################################################################################################################

            
#Plate Report

            Paste_row = file_count * 40 + 1 #Reset Paste_Row position
            
            plate_report_ws['A'+ str(Paste_row)] = 'Plate ' + str(file_count+1)
            plate_report_ws['B'+ str(Paste_row)] = 'Re-Run Status'
            plate_report_ws['C'+ str(Paste_row)] = 'Samples'
            plate_report_ws['D'+ str(Paste_row)] = 'Reason'
            plate_report_ws['F'+ str(Paste_row)] = 'Table'

            plate_report_ws['A' + str(Paste_row+1)] = data_file.split('\\')[-1][:-4]
            plate_report_ws['A'+ str(Paste_row+1)].font = Font(bold = True)
            plate_report_ws['A' + str(Paste_row+1)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')

            plate_report_ws['A'+ str(Paste_row)].font = Font(bold = True)

#Plate Report Table
            
            Paste_row = file_count * 40 + 1 #Reset Paste_Row position
            
            plate_report_ws['F' + str(Paste_row+1)] = 'Sample Name'
            
            plate_report_ws['G' + str(Paste_row+1)] = 'Qty Mean'
            
            plate_report_ws['H' + str(Paste_row+1)] = 'Replicate Difference'
            
            plate_report_ws['I' + str(Paste_row+1)] = '(+) SPK Value'
            plate_report_ws['J' + str(Paste_row+1)] = 'SPK Difference'
            plate_report_ws['K' + str(Paste_row+1)] = 'Inhibition?'
            

            for name in Sample_Name_List:
                plate_report_ws['F' + str(Paste_row+2)] = name
                Paste_row += 1

            Paste_row = file_count * 40 + 1 #Reset Paste_Row position

            for Qty_mean_value in QuantityMeans['Qty Mean']:
                plate_report_ws['G' + str(Paste_row+2)] = Qty_mean_value
                Paste_row += 1

            Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                
            for value in Replicate_Differences['Ct diff']:
                plate_report_ws['H' + str(Paste_row+2)] = value
                Paste_row+=1

            Paste_row = file_count * 40 + 1 #Reset Paste_Row position

            if Sample_Name_List[0].startswith('AC'):
                Paste_row += 1
            
            for x in range(len(SPK_Ct['SPK Values'])):
                plate_report_ws['I' + str(Paste_row+2)] = SPK_Ct['SPK Values'][x]
                plate_report_ws['J' + str(Paste_row+2)] = SPK_Ct['SPK Value Differences'][x]
                plate_report_ws['K' + str(Paste_row+2)] = SPK_Ct['Inhibition?'][x]
                Paste_row += 1

            Paste_row += 2
            plate_report_ws['I' + str(Paste_row+2)] = SPK_Ct['SPK Control Value'][0]
            plate_report_ws['I' + str(Paste_row+2)].font = Font(size = 12, bold = True)

            Paste_row = file_count * 40 + 1 #Reset Paste_Row position
            

#ReRun Status (Plates)

            re_run_status = 'green'
            re_run_samples_replicates = 'not yet'
            re_run_samples_inhib = 'not yet'
            
            Plate_Fail = False #Lets track if the plate failed
            Samples_Failed = False #and if samples failed

            #Status Color Indicator Default is Green
            plate_report_ws['B'+ str(Paste_row+1)].fill = PatternFill(start_color='0000FF00', end_color='0000FF00', fill_type='solid')

            #STDs 1-7
            STD_additional_comment = []
            
            for x in range(0,len(Replicate_Differences['Sample Name'])):
                if Replicate_Differences['Sample Name'][x].startswith('STD'):
                    if Replicate_Differences['Sample Name'][x] != 'STD8':
                        if '.' in str(Replicate_Differences['Ct diff'][x]): #Numbers only
                            if abs(Replicate_Differences['Ct diff'][x]) > 1: #Fails
                                plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                                plate_report_ws['D'+ str(Paste_row+2)] = Replicate_Differences['Sample Name'][x] + ' Failed' + ' '.join(STD_additional_comment)

                                if len(STD_additional_comment) == 0:
                                    STD_additional_comment.append(Replicate_Differences['Sample Name'][x] + 'Failed') #STD 1-7 message
                                else:
                                    STD_additional_comment.append('and' + Replicate_Differences['Sample Name'][x] + 'Failed')
                                    
                                re_run_status = 'red'
                                Samples_Failed = True
                               
                            
            Paste_row = file_count * 40 + 1 #Reset Paste_Row position

            #Only STD8 fails plate
            for x in range(0,len(Replicate_Differences['Sample Name'])):
                if Replicate_Differences['Sample Name'][x] == 'STD8':
                    if '.' in str(Replicate_Differences['Ct diff'][x]): #Numbers only
                        if abs(Replicate_Differences['Ct diff'][x]) > 1:
                            plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                            plate_report_ws['D'+ str(Paste_row+2)] = 'STD 8 Failed' + ' '.join(STD_additional_comment)

                            re_run_status = 'red'
                            Plate_Fail = True
                            Samples_Failed = True


            #Grabbing all samples needing re run due to failed STD, can replace above 2 blocks if modified

            plate_fail_reason = []
            sample_fail_reason = []
                        
            for x in range(0,len(Replicate_Differences['Sample Name'])):
                if Replicate_Differences['Sample Name'][x].startswith('STD'):

                    #Ask about this case (diff above 1 but qtymean < 25 ; do need re run?)
                    
                    #if '.' in str(Replicate_Differences['Ct diff'][x]):
                    #    if abs(Replicate_Differences['Ct diff'][x]) > 1:
                    #        plate_fail_reason.append(Replicate_Differences['Sample Name'][x])
                    
                    if str(Replicate_Differences['Ct diff'][x]).startswith('PROBLEM'):
                        plate_fail_reason.append(Replicate_Differences['Sample Name'][x])
                        Samples_Failed = True

                    
            Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                                                 

            #Slope, R2, LOD : The plate requirements

            plate_report_ws['M' + str(Paste_row+1)] = 'Slope'
            plate_report_ws['N' + str(Paste_row+1)] = 'R^2'
            plate_report_ws['O' + str(Paste_row+1)] = 'NTC'
            plate_report_ws['P' + str(Paste_row+1)] = 'LOD'
            plate_report_ws['M' + str(Paste_row+2)] = str(form10.iloc[-8,1]) + ' cycles/log decade'
            plate_report_ws['N' + str(Paste_row+2)] = form10.iloc[-6,2][:-1]
                
            plate_report_ws['P' + str(Paste_row+2)] = LOD

            if QuantityMeans['NTC'][0] > 0:
                plate_report_ws['O' + str(Paste_row+2)] = QuantityMeans['NTC'][0]
            else:
                plate_report_ws['O' + str(Paste_row+2)] = 'Undetermined'

            if float(form10.iloc[-8,1]) < float(SlopeMin):
                plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                plate_report_ws['D'+ str(Paste_row+3)] = 'Slope Requirement'
                plate_fail_reason.append('Slope Requirement')
                re_run_status = 'red'
                Plate_Fail = True
                Samples_Failed = True
              
            if float(form10.iloc[-8,1]) > float(SlopeMax):
                plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                plate_report_ws['D'+ str(Paste_row+3)] = 'Slope Requirement'
                plate_fail_reason.append('Slope Requirement')
                re_run_status = 'red'
                Plate_Fail = True
                Samples_Failed = True

            if isinstance(form10.iloc[-6,2][:-1], str) == False:
                if float(form10.iloc[-6,2][:-1]) < R2:
                    plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                    plate_report_ws['D'+ str(Paste_row+4)] = 'R^2 Requirement'
                    plate_fail_reason.append('R2 Requirement')
                    re_run_status = 'red'
                    Plate_Fail = True
                    Samples_Failed = True

            if QuantityMeans['NTC'][0] < float(LOD):
                plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                plate_report_ws['D'+ str(Paste_row+5)] = 'NTC vs LOD Requirement'
                plate_fail_reason.append('NTC vs LOD Requirement')
                re_run_status = 'red'
                Plate_Fail = True
                Samples_Failed = True




            #Is Qty Mean within STD curve range? Did Slope Fail? Which STD to kick?
            
            STD_DF = STD_info_DF(form10)


            if 'Slope Requirement' in plate_fail_reason:
                plate_report_ws['C' + str(Paste_row+1)] = STD_DF['statement'][0]
                

            below_std_curve_range_list = []
            above_std_curve_range_list = []
            min_STD_qty = 25
            max_STD_qty = 10000000



            min_STD_qty = min(STD_DF['std_range'])
            max_STD_qty = max(STD_DF['std_range'])


            Paste_row = file_count * 40 + 1 #Reset Paste_Row position

            
            for i in range(len(form10['Qty Mean'])):

                multiple_re_run_sample_reason = False

                if form10['Qty Mean'][i] < min_STD_qty: #Report dont retest
                    plate_report_ws['B' + str(Paste_row+31)] = 'Fine, Just Reporting'
                    plate_report_ws['D' + str(Paste_row+31)] = 'Qty Mean below ' + str(min_STD_qty)

                    if form10['Sample Name'][i] not in below_std_curve_range_list:
                        below_std_curve_range_list.append(form10['Sample Name'][i])
                    
                if form10['Qty Mean'][i] > max_STD_qty: #Dilute, Rerun
                    plate_report_ws['B' + str(Paste_row+32)] = 'Re-run Samples'
                    plate_report_ws['D' + str(Paste_row+32)] = 'Qty Mean above ' + str(max_STD_qty)

                    if form10['Sample Name'][i] not in above_std_curve_range_list:
                        above_std_curve_range_list.append(form10['Sample Name'][i])

                    Samples_Failed = True


                    #Update Re-Run Samples: above STD curve range
                    
                    if form10['Qty Mean'][i-1] > max_STD_qty and form10['Sample Name'][i].startswith('C') and form10['Sample Name'][i].endswith('spk') == False and form10['Sample Name'][i].endswith('SPK') == False: #No duplicate entry for Qty Mean > LOQ

                        for z in range(len(Re_Run_Sample_Info_Dataframe)): #is sample in re run df already?
                            
                            if Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][z] == form10['Sample Name'][i]:
                                first_reason = Re_Run_Sample_Info_Dataframe.iloc[z, 2]
                                Re_Run_Sample_Info_Dataframe.iloc[z, 2] = first_reason + 'and Qty Mean above STD curve range'
                                multiple_re_run_sample_reason = True
                                
                if multiple_re_run_sample_reason == True:
                    
                    Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples':form10['Sample Name'][i],'Reason_Re_Run':'Qty Mean above STD curve range'}, ignore_index=True)

                            
            plate_report_ws['C' + str(Paste_row+31)] = ' ,'.join(below_std_curve_range_list)
            plate_report_ws['C' + str(Paste_row+32)] = ' ,'.join(above_std_curve_range_list)

            
######################################################################################################################################################################

#Re_Run_Samples Report

            
                #Connect Samples that were re_run with their former selves.
                for C_name in SPK_Ct['Sample Name']:
                    C_name_with_no_plus_SPK = C_name.split('+')[0]
                    C_num = C_name.split('_')[0]
                    for Re_Run_Sample_Info_Index in range(0,len(Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'])):
                        if C_num == Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][Re_Run_Sample_Info_Index].split('_')[0]:
                            

                            #Robust dumb method FOR PASTING RE RAN SAMPLES into DF, relies on date sort for accuracy
                            Re_Run_Sample_Info_Dataframe.iloc[Re_Run_Sample_Info_Index, 3] = data_file.split('\\')[-1][:-4]
                            Re_Run_Sample_Info_Dataframe.iloc[Re_Run_Sample_Info_Index, 4] = C_name_with_no_plus_SPK
                            

                #Samples needing re-run from Plate Failiures
                if len(plate_fail_reason) > 0:
                    for i in range(0,len(SPK_Ct)): #First sample failiure error added
                        Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples':SPK_Ct['Sample Name'][i].split('+')[0],'Reason_Re_Run': ', '.join(plate_fail_reason) + ' Failed'}, ignore_index=True)


                           
                #Samples needing re-run from Replicate Difference 

                replicate_diff_issue_samples = []

                
                for Replicate_value_index in range(0,len(Replicate_Differences)):
                    
                    multiple_re_run_sample_reason = False #check if sample had other reasons to fail
                    

                    if str(Replicate_Differences['Ct diff'][Replicate_value_index]).startswith('PROBLEM'):

                        plate_report_ws['D'+ str(Paste_row+6)] = 'Replicate Difference'

                        if Replicate_Differences['Sample Name'][Replicate_value_index].startswith('C'):
                            replicate_diff_issue_samples.append(Replicate_Differences['Sample Name'][Replicate_value_index])
                            
                        re_run_samples_replicates = 'true'
                        re_run_status = 'red'
                        Samples_Failed = True
                        Paste_row += 1

                        if Replicate_Differences['Sample Name'][Replicate_value_index].startswith('AC'):
                            plate_fail_reason.append('AC diff')
                                               

                        #Update Re_Run_Samples with replicate diff fails
                        for z in range(len(Re_Run_Sample_Info_Dataframe)):         #is it in the re run sample df?
                            if Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][z] == Replicate_Differences['Sample Name'][Replicate_value_index]: #its in the df

                                multiple_re_run_sample_reason = True
                                Re_Run_Sample_Info_Dataframe.iloc[z, 2] = Re_Run_Sample_Info_Dataframe.iloc[z, 2] + ' and Replicate Difference'
                                
                        if multiple_re_run_sample_reason == False and Replicate_Differences['Sample Name'][Replicate_value_index].startswith('C') == True: #if the sample starts with problem and no other fail reason
                            Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples': Replicate_Differences['Sample Name'][Replicate_value_index].split('+')[0],'Reason_Re_Run':'Replicate Difference'}, ignore_index=True)


                plate_report_ws['C' + str(Paste_row+6)] = ', '.join(replicate_diff_issue_samples) #paste into 1 cell



                #Samples needing re-run from Inhibition

                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                inhibited_samples = []
                
                for Inhibited_Sample_Index in range(0,len(SPK_Ct['Inhibition?'])):

                    multiple_re_run_sample_reason = False #check if sample had other reasons to fail


                    if SPK_Ct['Inhibition?'][Inhibited_Sample_Index] == 'True':

                        plate_report_ws['C' + str(Paste_row+7)] = SPK_Ct['Sample Name'][Inhibited_Sample_Index]
                        plate_report_ws['D'+ str(Paste_row+7)] = 'Inhibited'

                        if SPK_Ct['Sample Name'][Inhibited_Sample_Index].startswith('C'):
                            inhibited_samples.append(SPK_Ct['Inhibition?'][Inhibited_Sample_Index])
                        
                        re_run_samples_inhib = 'true'
                        re_run_status = 'red'
                        Samples_Failed = True
                        Paste_row += 1


                        #Update Re_Run_Samples with inhibited
                        for z in range(len(Re_Run_Sample_Info_Dataframe)):         #is it in the re run sample df?
                            if Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][z] == SPK_Ct['Sample Name'][Inhibited_Sample_Index]: #it is in the df

                                Re_Run_Sample_Info_Dataframe.iloc[z, 2] = Re_Run_Sample_Info_Dataframe.iloc[z, 2] + ' and Inhibition'
                                multiple_re_run_sample_reason = True
                                
                        if multiple_re_run_sample_reason == False and SPK_Ct['Sample Name'][Inhibited_Sample_Index].startswith('C') == True: #if the sample starts with problem and no other fail reason
                            Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples':SPK_Ct['Sample Name'][Inhibited_Sample_Index].split('+')[0],'Reason_Re_Run':'Inhibited'}, ignore_index=True)

           
                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                

                #Some Re Run Status Reasons
                
                if re_run_samples_replicates == 'true':
                        plate_report_ws['B'+ str(Paste_row+6)] = 'Re-run Samples'

                if re_run_samples_inhib == 'true' and re_run_samples_replicates == 'not yet':
                        plate_report_ws['B'+ str(Paste_row+7)] = 'Re-run Samples'

                if re_run_status == 'red':
                        plate_report_ws['B'+ str(Paste_row+1)].fill = PatternFill(start_color='00FF0000',end_color='00FF0000',fill_type='solid')

                for x in Sample_Name_List[-9:]:
                        if x.startswith('S') == False:
                                plate_report_ws['D' + str(Paste_row+1)] = 'Re-Run with a STD removed'



#######


                #Is Qty Mean within STD curve range? Did Slope Fail? Which STD to kick? + Reruns
                
                STD_DF = STD_info_DF(form10)


                if 'Slope Requirement' in plate_fail_reason:
                    plate_report_ws['C' + str(Paste_row+1)] = STD_DF['statement'][0]
                    
                
                below_std_curve_range_list = []
                above_std_curve_range_list = []
                min_STD_qty = 25
                max_STD_qty = 10000000            



                min_STD_qty = min(STD_DF['std_range'])
                max_STD_qty = max(STD_DF['std_range'])


                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                
                for i in range(len(form10['Qty Mean'])):

                    multiple_re_run_sample_reason = False

                    if form10['Qty Mean'][i] < min_STD_qty: #Report dont retest
                        plate_report_ws['B' + str(Paste_row+31)] = 'Fine, Just Reporting'
                        plate_report_ws['D' + str(Paste_row+31)] = 'Qty Mean below ' + str(min_STD_qty)

                        if form10['Sample Name'][i] not in below_std_curve_range_list:
                            below_std_curve_range_list.append(form10['Sample Name'][i]) #Duplicate issue?
                        
                    if form10['Qty Mean'][i] > max_STD_qty and form10['Sample Name'][i+1] == form10['Sample Name'][i]: #Dilute, Rerun
                        plate_report_ws['B' + str(Paste_row+32)] = 'Re-run Samples'
                        plate_report_ws['D' + str(Paste_row+32)] = 'Qty Mean above ' + str(max_STD_qty)

                        if form10['Sample Name'][i] not in above_std_curve_range_list:
                            above_std_curve_range_list.append(form10['Sample Name'][i])

                        Samples_Failed = True


                        #Update Re-Run Samples: above STD curve range
                        
                        for z in range(len(Re_Run_Sample_Info_Dataframe)): #is sample in re run df already?
                            if Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][z] == form10['Sample Name'][i]: #it is in the df
                                
                                Re_Run_Sample_Info_Dataframe.iloc[z, 2] = Re_Run_Sample_Info_Dataframe.iloc[z, 2] + ' and Qty Mean above STD curve range'
                                multiple_re_run_sample_reason = True
                                    
                        if multiple_re_run_sample_reason == False and form10['Sample Name'][i].endswith('SPK') == False and form10['Sample Name'][i].endswith('spk') == False:
                            Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples':form10['Sample Name'][i],'Reason_Re_Run':'Qty Mean above STD curve range'}, ignore_index=True)

                            
                plate_report_ws['C' + str(Paste_row+31)] = ' ,'.join(below_std_curve_range_list)
                plate_report_ws['C' + str(Paste_row+32)] = ' ,'.join(above_std_curve_range_list)

#######################################################################################################################################################################

#AC Report

            AC_report_df = AC_report(form10)
            AC_qty_list = []
            

            #Add File Information to PlateInformation.txt; Info not contained in Plate_Info_DF however
            
            with open(Project_Output_Path + '\\' + 'Plate_Information.txt', 'a', newline = '\r') as Plate_Information:
                Plate_Information.write('\n'+data_file.split('\\')[-1][:-4]+'\t'+time.asctime()+'\t'+str(len(form10)+len(form10_raw_extra)+2)+'\t'+str(Plate_Fail)+'\t'+str(Samples_Failed)+'\t'+str(AC_report_df['Qty1'][0])+'\t'+str(AC_report_df['Qty2'][0]))
            Plate_Information.close()

            Plate_Info_DF = pd.read_csv(Project_Output_Path + '\\' + 'Plate_Information.txt', sep = '\t', lineterminator = '\r', header = 0)


            ac_report_ws['A1'] = 'Plate'
            ac_report_ws['B1'] = 'Sample Name'
            ac_report_ws['C1'] = 'Ct'
            ac_report_ws['D1'] = 'Qty'
            ac_report_ws['E1'] = 'Qty Mean'
            ac_report_ws['F1'] = 'Std Dev of Qty Mean'
            ac_report_ws['G1'] = 'Avg of Qty Mean'
            ac_report_ws['H1'] = 'Percent CV' #std/avg
            ac_report_ws['I1'] = '# of Plates'
            ac_report_ws['I2'] = list(Plate_Info_DF['Plate_Failed']).count(False)

         
            
            if Plate_Fail == False:
                if len(AC_report_df) > 0: #Not every plate has AC
                    
                    ac_report_ws['A' + str(plates_passed * 2 + 2)] = data_file.split('\\')[-1][:-4]
                    
                    ac_report_ws['B' + str(plates_passed * 2 + 2)] = 'AC'
                    ac_report_ws['B' + str(plates_passed * 2 + 3)] = 'AC'

                    ac_report_ws['C' + str(plates_passed * 2 + 2)] = AC_report_df['Ct1'][0]
                    ac_report_ws['C' + str(plates_passed * 2 + 3)] = AC_report_df['Ct2'][0]

                    ac_report_ws['D' + str(plates_passed * 2 + 2)] = AC_report_df['Qty1'][0]
                    ac_report_ws['D' + str(plates_passed * 2 + 3)] = AC_report_df['Qty2'][0]

                    ac_report_ws['E' + str(plates_passed * 2 + 2)] = AC_report_df['QtyMean'][0]
                   

                    plates_passed += 1
                

            for i in Plate_Info_DF['AC_qty1']:
                AC_qty_list.append(i)
            for i in Plate_Info_DF['AC_qty2']:
                AC_qty_list.append(i)
                
            ac_report_ws['F2'] = statistics.stdev(AC_qty_list)
            ac_report_ws['G2'] = ((sum(AC_qty_list))/len(AC_qty_list))
            ac_report_ws['H2'] = statistics.stdev(AC_qty_list) / ((sum(AC_qty_list))/len(AC_qty_list))
            ac_report_ws['I3'] = plates_passed



                                               
#######################################################################################################################################################################
                                    
#Locate Area to write next set of Calcultions
            file_count += 1
            currentRow = (file_count + files_before) * 109 + 1

#Save Re Run Sample Info
                            
            Re_Run_Sample_Info_Dataframe.to_csv(Project_Output_Path + '\\' + 'Re_Run_Sample_Info.txt', sep='\t', index = False, header = True)





########################################################################################################################################################################
########################################################################################################################################################################
#####################################     FOR QUANT STUDIO FILES     ###################################################################################################
########################################################################################################################################################################
########################################################################################################################################################################


            
        else:                              #Hello Quant Studio


            print('')
            print('Processing Quant Studio File: ' + data_file.split('\\')[-1])

            temp_form10 = pd.read_csv(data_file, sep = '\t', skip_blank_lines = False, usecols = [0])
            Index_where_data_begins = 0
            Data_exists = False


#Find the actual data
            
            for i in range(0, len(temp_form10)):
                if temp_form10.iloc[:,0][i] == '[Results]':
                    Data_exists = True
                    Index_where_data_begins = i+2 #+2 because Row 0 is the header and Row 1 is index 0

            if Data_exists == False:
                
                print("Theres no Data in QS File" + data_file.split('\\')[-1] + ", Plate " + str(file_count+files_before+1) + " will be empty")
                print('')
                
                if row_where_file_ends == 0:
                    ws['U' + str(row_where_file_ends+1)] = str(1 + file_count + files_before)
                    ws['U' + str(row_where_file_ends+1)].border = Border(bottom = Side(style = 'medium'))
                    ws['U' + str(row_where_file_ends+1)].font = Font(size = 16, bold = True)
                    ws['U' + str(row_where_file_ends+1)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
                    ws['U' + str(row_where_file_ends+1)].alignment = Alignment(horizontal = 'center')
                    ws['V' + str(row_where_file_ends+1)] = 'Empty Quant Studio File'
                    ws['V' + str(row_where_file_ends+1)].font = Font(bold = True)

                    row_where_file_ends += 1 #Avoid error of pasting into row 0

                    with open(Data_Directory_Path + '\\' + 'Plate_Information.txt', 'a', newline = '\r') as Plate_Information:
                        Plate_Information.write('\n'+data_file.split('\\')[-1][:-4]+'\t'+str(now.strftime('%Y-%m-%d %H:%M'))+'\t'+str(1))
                    Plate_Information.close()

                    #Plate Report

                    Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                
                    plate_report_ws['A'+ str(Paste_row)] = 'Plate ' + str(file_count+1)
                    plate_report_ws['A' + str(Paste_row+1)] = data_file.split('\\')[-1][:-4]
                    plate_report_ws['A'+ str(Paste_row+1)].font = Font(bold = True)
                    plate_report_ws['A'+ str(Paste_row)].font = Font(bold = True)
                    plate_report_ws['B'+ str(Paste_row)] = 'Re-Run Status'
                    plate_report_ws['B'+ str(Paste_row+1)].fill = PatternFill(start_color='00FF0000',end_color='00FF0000',fill_type='solid')
                    plate_report_ws['D'+ str(Paste_row)] = 'Reason'
                    plate_report_ws['D'+ str(Paste_row+1)] = 'Empty Quant Studio File'

                else:
                    ws['U' + str(row_where_file_ends)] = str(1 + file_count + files_before)
                    ws['U' + str(row_where_file_ends)].border = Border(bottom = Side(style = 'medium'))
                    ws['U' + str(row_where_file_ends)].font = Font(size = 16, bold = True)
                    ws['U' + str(row_where_file_ends)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
                    ws['U' + str(row_where_file_ends)].alignment = Alignment(horizontal = 'center')
                    ws['V' + str(row_where_file_ends)] = 'Empty Quant Studio File'

                    #Plate Report

                    Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                
                    plate_report_ws['A'+ str(Paste_row)] = 'Plate ' + str(file_count+1)
                    plate_report_ws['A' + str(Paste_row+1)] = data_file.split('\\')[-1][:-4]
                    plate_report_ws['A'+ str(Paste_row+1)].font = Font(bold = True)
                    plate_report_ws['A'+ str(Paste_row)].font = Font(bold = True)
                    plate_report_ws['B'+ str(Paste_row)] = 'Re-Run Status'
                    plate_report_ws['B'+ str(Paste_row+1)].fill = PatternFill(start_color='00FF0000',end_color='00FF0000',fill_type='solid')
                    plate_report_ws['D'+ str(Paste_row)] = 'Reason'
                    plate_report_ws['D'+ str(Paste_row+1)] = 'Empty Quant Studio File'

                wb.save(workbook_path)


#Process Quant File
                
            else:
                
                form10_unsorted = pd.read_csv(data_file, sep = '\t', lineterminator = '\n', header = Index_where_data_begins, usecols = ['Well', 'Sample Name', 'Target Name','Reporter','Task','CT','Quantity','Quantity Mean', 'Slope','R(superscript 2)', 'Y-Intercept'])


                #Sort Samples
                form10_full = form10_unsorted.sort_values(by=['Sample Name'], ignore_index = True)


                #Rename SPK
                spike_end = 0
                Titer_info_exists = False
                
                for i in range(len(form10_full)):

                    if form10_full['Sample Name'][i].startswith('Tit'):
                        Titer_info_exists = True
                        
                    if form10_full['Sample Name'][i] == 'SPK CONTROL' and form10_full['Sample Name'][i-1] == 'SPK CONTROL':
                        spike_end = i
                    
                    if form10_full['Sample Name'][i] == 'Spike Control' and form10_full['Sample Name'][i-1] == 'Spike Control':
                        spike_end = i

                        form10_full.iloc[spike_end,1] = 'SPK CONTROL'
                        form10_full.iloc[spike_end-1,1] = 'SPK CONTROL'

                        
                #Rename Columns
                form10_full.rename(columns={"CT": "Ct", "Quantity Mean": "Qty Mean", "R(superscript 2)":"R^2"}, inplace = True)
                

                #Take commas out of columns
                form10_full['Quantity'] = form10_full['Quantity'].str.replace(',', '').astype('float')
                form10_full['Qty Mean'] = form10_full['Qty Mean'].str.replace(',', '').astype('float')


                #Chop off Titer Information
                
                if Titer_info_exists == True:
                    
                    titer_idx = 0
                    titers = []
                    for i in range(len(form10_full['Sample Name'])):
                        if form10_full['Sample Name'][i].startswith('Titer'):
                            titer_idx = i
                            break
                    for z in range(titer_idx, len(form10_full)):
                        titers.append(z)
                        
                    form10 = form10_full.drop(titers)

                else:
                    form10 = form10_full

                #Add a Header/Footer
                Header = pd.DataFrame()
                Header = Header.append(['Quant Studio Results',' ','For File:',data_file.split('\\')[-1][:-4],' ',' ',' ',' ', 'Sample Information',' '])
                Footer = pd.DataFrame()
                Footer = Footer.append([' ',' '])

#Create Modified Quant Studio File

                Header.to_csv(Project_Output_Path + '\\Modified_Quant_data_files\\MODIFIED_QUANTSTUDIO_FILE_' + data_file.split('\\')[-1], sep = '\t', mode='w',index=False, quoting=csv.QUOTE_NONE, quotechar='',  escapechar='$')
                form10.to_csv(Project_Output_Path + '\\Modified_Quant_data_files\\MODIFIED_QUANTSTUDIO_FILE_' + data_file.split('\\')[-1], sep = '\t', mode='a',index=False, quoting=csv.QUOTE_NONE, quotechar='',  escapechar='$')
                Footer.to_csv(Project_Output_Path + '\\Modified_Quant_data_files\\MODIFIED_QUANTSTUDIO_FILE_' + data_file.split('\\')[-1], sep = '\t', mode='a',index=False, quoting=csv.QUOTE_NONE, quotechar='',  escapechar='$')
                

#Write QS Data to Excel

                open_data_file = open(Project_Output_Path + '\\Modified_Quant_data_files\\MODIFIED_QUANTSTUDIO_FILE_' + data_file.split('\\')[-1], 'r')
                
                currentRow = (file_count + files_before) * 109 + 1
                currentCell = None
           
                currentCol = 2

                while True:
                    
                    line = open_data_file.readline()

                    if not line:
                        break

                    if line in ['\n', '\r', '\r\n'] or line.startswith('NAP'):
                        continue

                    fields = line.split("\t") #separate line by tabs into up to 34 cells
                    for fieldIndex in range(len(fields)): #copy up to 21 cells in each line
                    
                        currentCell = ws.cell(row=currentRow, column=currentCol)

                        if fields[fieldIndex]: #only try to copy field if it has a value
                            if (fields[fieldIndex][0].isdigit() or fields[fieldIndex].startswith("-")): #set cell value as a float if it begins with a number or a "-"

                                try: #copy value as a float
                                    currentCell.value = float(fields[fieldIndex])

                                except: #copy value as text
                                    currentCell.value = fields[fieldIndex]

                            else: #copy value as text
                                currentCell.value = fields[fieldIndex]

                        currentCol += 1

                        if (currentCol == 23 or fieldIndex == len(fields)-1): #move to the next row in sheet after 21 columns or at the end of the row
                            currentCol = 2
                            currentRow += 1
                            break


                open_data_file.close()
                    

                print('Quant Studio file ' + data_file.split('\\')[-1] + ' is written')
                print('')
                

#QS Form 10 Table
                
                
                #Manage Paste Offset
                Paste_row = row_where_file_ends + 20

                Sample_Name_List = Get_Sample_Names(form10)

                AC_exist = False
                for name in Sample_Name_List:
                    
                    if name.startswith('AC'):
                        Paste_row -= 1
                        AC_exist = True
                        break

                #Paste Calculation Header
                ws['X' + str(Paste_row-1)] = 'Sample Name'
                
                ws['Y' + str(Paste_row-1)] = 'Qty Mean'
                
                ws['Z' + str(Paste_row-1)] = 'Replicate Difference'
                
                ws['AA' + str(Paste_row-1)] = '(+) SPK Value'
                ws['AB' + str(Paste_row-1)] = 'SPK Difference'
                ws['AC' + str(Paste_row-1)] = 'Inhibition?'


                #Paste Sample Names

                #Sample_Name_List = form10func.Get_Sample_Names(form10)
                for name in Sample_Name_List:
                    ws['X' + str(Paste_row)] = name
                    Paste_row += 1


                #QuantityMean Calculation and write
                    
                Paste_row = row_where_file_ends + 20 #Reset Paste_Row position

                if AC_exist == True:
                    Paste_row -= 1
                    
                QuantityMeans = QtyMean_Dataframe(form10)
                for Qty_mean_value in QuantityMeans['Qty Mean']:
                    ws['Y' + str(Paste_row)] = Qty_mean_value
                    Paste_row += 1

            
                #Replicate Diff
                    
                Paste_row = row_where_file_ends + 20 #Reset Paste_Row position
                if AC_exist == True:
                    Paste_row -= 1
                    
                Replicate_Differences = Replicate_Diff_Dataframe(form10, LOQ)
                for value in Replicate_Differences['Ct diff']:
                    ws['Z' + str(Paste_row)] = value
                    Paste_row+=1

                                          
                #SPK Ct
                    
                Paste_row = row_where_file_ends + 20 #Reset Paste_Row position
                if AC_exist == True:
                    Paste_row -= 1
                    
                if Sample_Name_List[0].startswith('A'):
                    Paste_row += 1
                
                SPK_Ct = SPK_Ct_dataframe(form10) #SPK_Ct will be used for a list of C_Samples
                for x in range(len(SPK_Ct['SPK Values'])):
                    ws['AA' + str(Paste_row)] = SPK_Ct['SPK Values'][x]
                    ws['AB' + str(Paste_row)] = SPK_Ct['SPK Value Differences'][x]
                    ws['AC' + str(Paste_row)] = SPK_Ct['Inhibition?'][x]
                    Paste_row += 1

                #SPK Control Information
                Paste_row += 2
                ws['AA' + str(Paste_row)] = SPK_Ct['SPK Control Value'][0]
                ws['AA' + str(Paste_row)].font = Font(size = 13, bold = True)

#QS Style Formatting

                
                #ws['U' + str(row_where_file_ends+1)] = 1 + file_count + files_before
                #ws['U' + str(row_where_file_ends+1)].font = Font(size = 16, bold = True)
                #ws['U' + str(row_where_file_ends+1)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
                #ws['U' + str(row_where_file_ends+1)].alignment = Alignment(horizontal = 'center')

                row_where_file_ends = row_where_file_ends + 109

##############################################################################################################################################################################

                
#QS Plate Report

                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                
                plate_report_ws['A'+ str(Paste_row)] = 'Plate ' + str(file_count+1)
                plate_report_ws['B'+ str(Paste_row)] = 'Re-Run Status'
                plate_report_ws['C'+ str(Paste_row)] = 'Samples'
                plate_report_ws['D'+ str(Paste_row)] = 'Reason'
                plate_report_ws['F'+ str(Paste_row)] = 'Table'

                plate_report_ws['A' + str(Paste_row+1)] = data_file.split('\\')[-1][:-4]
                plate_report_ws['A'+ str(Paste_row+1)].font = Font(bold = True)
                plate_report_ws['A' + str(Paste_row+1)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')

                plate_report_ws['A'+ str(Paste_row)].font = Font(bold = True)

#QS Plate Report Table
                
                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                
                plate_report_ws['F' + str(Paste_row+1)] = 'Sample Name'
                
                plate_report_ws['G' + str(Paste_row+1)] = 'Qty Mean'
                
                plate_report_ws['H' + str(Paste_row+1)] = 'Replicate Difference'
                
                plate_report_ws['I' + str(Paste_row+1)] = '(+) SPK Value'
                plate_report_ws['J' + str(Paste_row+1)] = 'SPK Difference'
                plate_report_ws['K' + str(Paste_row+1)] = 'Inhibition?'
                

                for name in Sample_Name_List:
                    plate_report_ws['F' + str(Paste_row+2)] = name
                    Paste_row += 1

                Paste_row = file_count * 40 + 1 #Reset Paste_Row position

                for Qty_mean_value in QuantityMeans['Qty Mean']:
                    plate_report_ws['G' + str(Paste_row+2)] = Qty_mean_value
                    Paste_row += 1

                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                    
                for value in Replicate_Differences['Ct diff']:
                    plate_report_ws['H' + str(Paste_row+2)] = value
                    Paste_row+=1

                Paste_row = file_count * 40 + 1 #Reset Paste_Row position

                if Sample_Name_List[0].startswith('AC'):
                    Paste_row += 1
                
                for x in range(len(SPK_Ct['SPK Values'])):
                    plate_report_ws['I' + str(Paste_row+2)] = SPK_Ct['SPK Values'][x]
                    plate_report_ws['J' + str(Paste_row+2)] = SPK_Ct['SPK Value Differences'][x]
                    plate_report_ws['K' + str(Paste_row+2)] = SPK_Ct['Inhibition?'][x]
                    Paste_row += 1

                Paste_row += 2
                plate_report_ws['I' + str(Paste_row+2)] = SPK_Ct['SPK Control Value'][0]
                plate_report_ws['I' + str(Paste_row+2)].font = Font(size = 12, bold = True)

                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                

#QS ReRun Status (Plates)

                re_run_status = 'green'
                re_run_samples_replicates = 'not yet'
                re_run_samples_inhib = 'not yet'

                Plate_Fail = False #Lets track if the plate failed
                Samples_Failed = False #and if samples failed

                #Status Color Indicator Default is Green
                plate_report_ws['B'+ str(Paste_row+1)].fill = PatternFill(start_color='0000FF00', end_color='0000FF00', fill_type='solid')

                #STDs 1-7
                STD_additional_comment = []
                
                for x in range(0,len(Replicate_Differences['Sample Name'])):
                    if Replicate_Differences['Sample Name'][x].startswith('STD'):
                        if Replicate_Differences['Sample Name'][x] != 'STD8':
                            if '.' in str(Replicate_Differences['Ct diff'][x]): #Numbers only
                                if abs(Replicate_Differences['Ct diff'][x]) > 1: #Fails
                                    plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                                    plate_report_ws['D'+ str(Paste_row+2)] = Replicate_Differences['Sample Name'][x] + ' Failed' + ' '.join(STD_additional_comment)

                                    if len(STD_additional_comment) == 0:
                                        STD_additional_comment.append(Replicate_Differences['Sample Name'][x] + 'Failed') #STD 1-7 message
                                    else:
                                        STD_additional_comment.append('and' + Replicate_Differences['Sample Name'][x] + 'Failed')


                                    re_run_status = 'red'
                                    Samples_Failed = True
                                
                Paste_row = file_count * 40 + 1 #Reset Paste_Row position

                #Only STD8 fails plate
                for x in range(0,len(Replicate_Differences['Sample Name'])):
                    if Replicate_Differences['Sample Name'][x] == 'STD8':
                        if '.' in str(Replicate_Differences['Ct diff'][x]): #Numbers only
                            if abs(Replicate_Differences['Ct diff'][x]) > 1:
                                plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                                plate_report_ws['D'+ str(Paste_row+2)] = 'STD 8 Failed' + ' '.join(STD_additional_comment)

                                re_run_status = 'red'
                                Plate_Fail = True
                                Samples_Failed = True


                #Grabbing all samples needing re run due to failed STD, can replace above 2 blocks if modified

                plate_fail_reason = []
                sample_fail_reason = []
                            
                for x in range(0,len(Replicate_Differences['Sample Name'])):
                    if Replicate_Differences['Sample Name'][x].startswith('STD'):

                        #if '.' in str(Replicate_Differences['Ct diff'][x]):
                        #    if abs(Replicate_Differences['Ct diff'][x]) > 1:
                        #        plate_fail_reason.append(Replicate_Differences['Sample Name'][x])

                        if str(Replicate_Differences['Ct diff'][x]).startswith('PROBLEM'):
                            plate_fail_reason.append(Replicate_Differences['Sample Name'][x])
                            Samples_Failed = True
                            
                        
                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                                                     

                #Slope, R2, LOD

                plate_report_ws['M' + str(Paste_row+1)] = 'Slope'
                plate_report_ws['N' + str(Paste_row+1)] = 'R^2'
                plate_report_ws['O' + str(Paste_row+1)] = 'NTC'
                plate_report_ws['P' + str(Paste_row+1)] = 'LOD'
                plate_report_ws['M' + str(Paste_row+2)] = str(form10['Slope'][0]) + ' cycles/log decade'
                plate_report_ws['N' + str(Paste_row+2)] = form10['R^2'][0]
                    
                plate_report_ws['P' + str(Paste_row+2)] = LOD

                if QuantityMeans['NTC'][0] > 0:
                    plate_report_ws['O' + str(Paste_row+2)] = QuantityMeans['NTC'][0]
                else:
                    plate_report_ws['O' + str(Paste_row+2)] = 'Undetermined'

                if float(form10['Slope'][0]) < SlopeMin:
                    plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                    plate_report_ws['D'+ str(Paste_row+3)] = 'Slope Requirement'
                    plate_fail_reason.append('Slope Requirement')
                    re_run_status = 'red'
                    Plate_Fail = True
                    Samples_Failed = True

                if float(form10['Slope'][0]) > SlopeMax:
                    plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                    plate_report_ws['D'+ str(Paste_row+3)] = 'Slope Requirement'
                    plate_fail_reason.append('Slope Requirement')
                    re_run_status = 'red'
                    Plate_Fail = True
                    Samples_Failed = True
                    
                if isinstance(form10['R^2'][0], str) == False:
                    if float(form10['R^2'][0]) < R2:
                        plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                        plate_report_ws['D'+ str(Paste_row+4)] = 'R^2 Requirement'
                        plate_fail_reason.append('R2 Requirement')
                        re_run_status = 'red'
                        Plate_Fail = True
                        
                if QuantityMeans['NTC'][0] < LOD:
                    plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                    plate_report_ws['D'+ str(Paste_row+5)] = 'NTC vs LOD Requirement'
                    plate_fail_reason.append('NTC vs LOD Requirement')
                    re_run_status = 'red'
                    Plate_Fail = True
                    Samples_Failed = True
                



                
                


                
#############################################################################################################################

#QS Re_Run_Samples Report

            
                #Connect Samples that were re_run with their former selves.
                for C_name in SPK_Ct['Sample Name']:
                    C_name_with_no_plus_SPK = C_name.split('+')[0]
                    C_num = C_name.split('_')[0]
                    for Re_Run_Sample_Info_Index in range(0,len(Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'])):
                        if C_num == Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][Re_Run_Sample_Info_Index].split('_')[0]:
                            

                            #Robust dumb method FOR PASTING RE RAN SAMPLES into DF, relies on date sort for accuracy
                            Re_Run_Sample_Info_Dataframe.iloc[Re_Run_Sample_Info_Index, 3] = data_file.split('\\')[-1][:-4]
                            Re_Run_Sample_Info_Dataframe.iloc[Re_Run_Sample_Info_Index, 4] = C_name_with_no_plus_SPK
                            

                #Samples needing re-run from Plate Failiures
                if len(plate_fail_reason) > 0:
                    for i in range(0,len(SPK_Ct)): #First sample failiure error added
                        Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples':SPK_Ct['Sample Name'][i].split('+')[0],'Reason_Re_Run': ', '.join(plate_fail_reason) + ' Failed'}, ignore_index=True)


                           
                #Samples needing re-run from Replicate Difference 

                replicate_diff_issue_samples = []

                
                for Replicate_value_index in range(0,len(Replicate_Differences)):
                    
                    multiple_re_run_sample_reason = False #check if sample had other reasons to fail
                    

                    if str(Replicate_Differences['Ct diff'][Replicate_value_index]).startswith('PROBLEM'):

                        plate_report_ws['D'+ str(Paste_row+6)] = 'Replicate Difference'

                        if Replicate_Differences['Sample Name'][Replicate_value_index].startswith('C'):
                            replicate_diff_issue_samples.append(Replicate_Differences['Sample Name'][Replicate_value_index])
                            
                        re_run_samples_replicates = 'true'
                        re_run_status = 'red'
                        Samples_Failed = True
                        Paste_row += 1

                        if Replicate_Differences['Sample Name'][Replicate_value_index].startswith('AC'):
                            plate_fail_reason.append('AC diff')
                                               

                        #Update Re_Run_Samples with replicate diff fails
                        for z in range(len(Re_Run_Sample_Info_Dataframe)):         #is it in the re run sample df?
                            if Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][z] == Replicate_Differences['Sample Name'][Replicate_value_index]: #its in the df

                                multiple_re_run_sample_reason = True
                                Re_Run_Sample_Info_Dataframe.iloc[z, 2] = Re_Run_Sample_Info_Dataframe.iloc[z, 2] + ' and Replicate Difference'
                                
                        if multiple_re_run_sample_reason == False and Replicate_Differences['Sample Name'][Replicate_value_index].startswith('C') == True: #if the sample starts with problem and no other fail reason
                            Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples': Replicate_Differences['Sample Name'][Replicate_value_index].split('+')[0],'Reason_Re_Run':'Replicate Difference'}, ignore_index=True)


                plate_report_ws['C' + str(Paste_row+6)] = ', '.join(replicate_diff_issue_samples) #paste into 1 cell



                #Samples needing re-run from Inhibition

                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                inhibited_samples = []
                
                for Inhibited_Sample_Index in range(0,len(SPK_Ct['Inhibition?'])):

                    multiple_re_run_sample_reason = False #check if sample had other reasons to fail


                    if SPK_Ct['Inhibition?'][Inhibited_Sample_Index] == 'True':

                        plate_report_ws['C' + str(Paste_row+7)] = SPK_Ct['Sample Name'][Inhibited_Sample_Index]
                        plate_report_ws['D'+ str(Paste_row+7)] = 'Inhibited'

                        if SPK_Ct['Sample Name'][Inhibited_Sample_Index].startswith('C'):
                            inhibited_samples.append(SPK_Ct['Inhibition?'][Inhibited_Sample_Index])
                        
                        re_run_samples_inhib = 'true'
                        re_run_status = 'red'
                        Samples_Failed = True
                        Paste_row += 1


                        #Update Re_Run_Samples with inhibited
                        for z in range(len(Re_Run_Sample_Info_Dataframe)):         #is it in the re run sample df?
                            if Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][z] == SPK_Ct['Sample Name'][Inhibited_Sample_Index]: #it is in the df

                                Re_Run_Sample_Info_Dataframe.iloc[z, 2] = Re_Run_Sample_Info_Dataframe.iloc[z, 2] + ' and Inhibition'
                                multiple_re_run_sample_reason = True
                                
                        if multiple_re_run_sample_reason == False and SPK_Ct['Sample Name'][Inhibited_Sample_Index].startswith('C') == True: #if the sample starts with problem and no other fail reason
                            Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples':SPK_Ct['Sample Name'][Inhibited_Sample_Index].split('+')[0],'Reason_Re_Run':'Inhibited'}, ignore_index=True)

           
                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                

                #Some Re Run Status Reasons
                
                if re_run_samples_replicates == 'true':
                        plate_report_ws['B'+ str(Paste_row+6)] = 'Re-run Samples'

                if re_run_samples_inhib == 'true' and re_run_samples_replicates == 'not yet':
                        plate_report_ws['B'+ str(Paste_row+7)] = 'Re-run Samples'

                if re_run_status == 'red':
                        plate_report_ws['B'+ str(Paste_row+1)].fill = PatternFill(start_color='00FF0000',end_color='00FF0000',fill_type='solid')

                for x in Sample_Name_List[-9:]:
                        if x.startswith('S') == False:
                                plate_report_ws['D' + str(Paste_row+1)] = 'Re-Run with a STD removed'



#######


                #Is Qty Mean within STD curve range? Did Slope Fail? Which STD to kick? + Reruns
                
                STD_DF = STD_info_DF(form10)


                if 'Slope Requirement' in plate_fail_reason:
                    plate_report_ws['C' + str(Paste_row+1)] = STD_DF['statement'][0]
                    
                
                below_std_curve_range_list = []
                above_std_curve_range_list = []
                min_STD_qty = 25
                max_STD_qty = 10000000            



                min_STD_qty = min(STD_DF['std_range'])
                max_STD_qty = max(STD_DF['std_range'])


                Paste_row = file_count * 40 + 1 #Reset Paste_Row position
                
                for i in range(len(form10['Qty Mean'])):

                    multiple_re_run_sample_reason = False

                    if form10['Qty Mean'][i] < min_STD_qty: #Report dont retest
                        plate_report_ws['B' + str(Paste_row+31)] = 'Fine, Just Reporting'
                        plate_report_ws['D' + str(Paste_row+31)] = 'Qty Mean below ' + str(min_STD_qty)

                        if form10['Sample Name'][i] not in below_std_curve_range_list:
                            below_std_curve_range_list.append(form10['Sample Name'][i]) #Duplicate issue?
                        
                    if form10['Qty Mean'][i] > max_STD_qty and form10['Sample Name'][i+1] == form10['Sample Name'][i]: #Dilute, Rerun
                        plate_report_ws['B' + str(Paste_row+32)] = 'Re-run Samples'
                        plate_report_ws['D' + str(Paste_row+32)] = 'Qty Mean above ' + str(max_STD_qty)

                        if form10['Sample Name'][i] not in above_std_curve_range_list:
                            above_std_curve_range_list.append(form10['Sample Name'][i])

                        Samples_Failed = True


                        #Update Re-Run Samples: above STD curve range
                        
                        for z in range(len(Re_Run_Sample_Info_Dataframe)): #is sample in re run df already?
                            if Re_Run_Sample_Info_Dataframe['Need_Re_Run_Samples'][z] == form10['Sample Name'][i]: #it is in the df
                                
                                Re_Run_Sample_Info_Dataframe.iloc[z, 2] = Re_Run_Sample_Info_Dataframe.iloc[z, 2] + ' and Qty Mean above STD curve range'
                                multiple_re_run_sample_reason = True
                                    
                        if multiple_re_run_sample_reason == False and form10['Sample Name'][i].endswith('SPK') == False and form10['Sample Name'][i].endswith('spk') == False:
                            Re_Run_Sample_Info_Dataframe = Re_Run_Sample_Info_Dataframe.append({'Need_Re_Run_Plate':data_file.split('\\')[-1][:-4],'Need_Re_Run_Samples':form10['Sample Name'][i],'Reason_Re_Run':'Qty Mean above STD curve range'}, ignore_index=True)

                            
                plate_report_ws['C' + str(Paste_row+31)] = ' ,'.join(below_std_curve_range_list)
                plate_report_ws['C' + str(Paste_row+32)] = ' ,'.join(above_std_curve_range_list)
                                
######################################################################################################################################################################

#QS AC Report


                AC_report_df = AC_report(form10)
                AC_qty_list = []
                

                #Add File Information to PlateInformation.txt; Have to reload Plate_Info_DF
                
                with open(Project_Output_Path + '\\' + 'Plate_Information.txt', 'a', newline = '\r') as Plate_Information:
                    Plate_Information.write('\n'+data_file.split('\\')[-1][:-4]+'\t'+time.asctime()+'\t'+str(len(form10)+12)+'\t'+str(Plate_Fail)+'\t'+str(Samples_Failed)+'\t'+str(AC_report_df['Qty1'][0])+'\t'+str(AC_report_df['Qty2'][0]))
                Plate_Information.close()
                
                Plate_Info_DF = pd.read_csv(Project_Output_Path + '\\' + 'Plate_Information.txt', sep = '\t', lineterminator = '\r', header = 0)


                ac_report_ws['A1'] = 'Plate'
                ac_report_ws['B1'] = 'Sample Name'
                ac_report_ws['C1'] = 'Ct'
                ac_report_ws['D1'] = 'Qty'
                ac_report_ws['E1'] = 'Qty Mean'
                ac_report_ws['F1'] = 'Std Dev of Qty Mean'
                ac_report_ws['G1'] = 'Avg of Qty Mean'
                ac_report_ws['H1'] = 'Percent CV' #std/avg
                ac_report_ws['I1'] = '# of Plates'
                ac_report_ws['I2'] = list(Plate_Info_DF['Plate_Failed']).count(False)

             
                
                if Plate_Fail == False:
                    if len(AC_report_df) > 0: #Not every plate has AC
                        
                        ac_report_ws['A' + str(plates_passed * 2 + 2)] = data_file.split('\\')[-1][:-4]
                        
                        ac_report_ws['B' + str(plates_passed * 2 + 2)] = 'AC'
                        ac_report_ws['B' + str(plates_passed * 2 + 3)] = 'AC'

                        ac_report_ws['C' + str(plates_passed * 2 + 2)] = AC_report_df['Ct1'][0]
                        ac_report_ws['C' + str(plates_passed * 2 + 3)] = AC_report_df['Ct2'][0]

                        ac_report_ws['D' + str(plates_passed * 2 + 2)] = AC_report_df['Qty1'][0]
                        ac_report_ws['D' + str(plates_passed * 2 + 3)] = AC_report_df['Qty2'][0]

                        ac_report_ws['E' + str(plates_passed * 2 + 2)] = AC_report_df['QtyMean'][0]
                       

                        plates_passed += 1
                    

                for i in Plate_Info_DF['AC_qty1']:
                    AC_qty_list.append(i)
                for i in Plate_Info_DF['AC_qty2']:
                    AC_qty_list.append(i)

                if len(file_path_list) - file_count == 1:    
                    ac_report_ws['F2'] = statistics.stdev(AC_qty_list)
                    ac_report_ws['G2'] = ((sum(AC_qty_list))/len(AC_qty_list))
                    ac_report_ws['H2'] = statistics.stdev(AC_qty_list) / ((sum(AC_qty_list))/len(AC_qty_list))
                    ac_report_ws['I3'] = plates_passed



                file_count += 1



##############################################################################################################

#Write Re Run Samples

            Re_Run_Sample_Info_Dataframe.to_csv(Project_Output_Path + '\\' + 'Re_Run_Sample_Info.txt', sep='\t', index = False, header = True)
            
    if files_before == 0:
        for r in dataframe_to_rows(Re_Run_Sample_Info_Dataframe, index=False, header=True):
            re_run_samples_ws.append(r)
    else:
        
        del wb['Re Run Samples']
        re_run_samples_ws = wb.create_sheet('Re Run Samples')
        
        for r in dataframe_to_rows(Re_Run_Sample_Info_Dataframe, index=False, header=True):
            re_run_samples_ws.append(r)
            


#Set Column Width for Readability

    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['X'].width = 17

    plate_report_ws.column_dimensions['A'].width = 11
    plate_report_ws.column_dimensions['B'].width = 16
    plate_report_ws.column_dimensions['C'].width = 21
    plate_report_ws.column_dimensions['D'].width = 22
    plate_report_ws.column_dimensions['F'].width = 18
    #plate_report_ws.column_dimensions['G'].width = 8
    plate_report_ws.column_dimensions['H'].width = 17
    plate_report_ws.column_dimensions['I'].width = 12
    plate_report_ws.column_dimensions['J'].width = 13
    plate_report_ws.column_dimensions['K'].width = 10
    plate_report_ws.column_dimensions['M'].width = 10
    plate_report_ws.column_dimensions['N'].width = 10
    plate_report_ws.column_dimensions['O'].width = 13
    plate_report_ws.column_dimensions['P'].width = 5


    re_run_samples_ws.column_dimensions['A'].width = 40
    re_run_samples_ws.column_dimensions['B'].width = 25
    re_run_samples_ws.column_dimensions['C'].width = 60
    re_run_samples_ws.column_dimensions['D'].width = 40
    re_run_samples_ws.column_dimensions['E'].width = 25

    ac_report_ws.column_dimensions['A'].width = 36
    ac_report_ws.column_dimensions['B'].width = 13
    ac_report_ws.column_dimensions['C'].width = 12
    ac_report_ws.column_dimensions['D'].width = 11
    ac_report_ws.column_dimensions['E'].width = 12
    ac_report_ws.column_dimensions['F'].width = 18
    ac_report_ws.column_dimensions['G'].width = 16
    ac_report_ws.column_dimensions['H'].width = 11
    ac_report_ws.column_dimensions['I'].width = 10


    
        
#Save!

    wb.save(workbook_path)



#Save Back Up File
    
    if backup_file_exists == True:
        os.rename(Backup_Filename, Project_Output_Path + '\\' + 'Form10_BackUp_' + Excel_File_Name[:-5] + '_' + str(now.strftime('%m-%d-%Y_time_%H_%M')) + '.xlsm')
        shutil.copyfile(Project_Output_Path + '\\' + Excel_File_Name, Project_Output_Path + '\\' + 'Form10_BackUp_' + Excel_File_Name[:-5] + '_' + str(now.strftime('%m-%d-%Y_time_%H_%M')) + '.xlsm')
    else:
        shutil.copyfile(Project_Output_Path + '\\' + Excel_File_Name, Project_Output_Path + '\\' + 'Form10_BackUp_' + Excel_File_Name[:-5] + '_' + str(now.strftime('%m-%d-%Y_time_%H_%M')) + '.xlsm')

    wb.save(workbook_path)
    
#Run the Program

#CMD Line: python C:\Users\efu\Desktop\Python\SCS19AutoForm10_Excelver.py "1" "2" "3" "4" "5" "6"

#if __name__ == '__main__':
#    execute(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], sys.argv[6])

    #Data_Directory_Path, Excel_File_Name, Slope_Min, Slope_Max, LOD, R2


execute(Data_Directory_Path, Excel_File_Name, Project_Output_Path, SlopeMin, SlopeMax, LOD, LOQ, R2)


print(time.asctime())

print('Done')
