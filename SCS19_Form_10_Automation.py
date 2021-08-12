from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd
import os
import shutil
from pathlib import Path
from datetime import datetime
import form10func



def execute(Data_Directory_Path, Excel_File_Name, SlopeMin, SlopeMax, LOD, R2):

#Find and Sort .TXT Raw Data Files
        
    excel_file_exists = False
    
    file_path_list = []
    
    for root,dirs,files in os.walk(Data_Directory_Path):
        for names in files:
            if names == Excel_File_Name:
                excel_file_exists = True
                
            if names.endswith('.txt'):
                file_path_list.append(os.path.join(root,names))
                file_path_list.sort(key=os.path.getmtime)
        
    #file_path_list.reverse() #newest files are on top for report
              
    
    file_count = 0
    

#Set Up PlateInformation.txt File

    Plate_Info_Exists = False
    Re_Run_Sample_Info_Exists = False

    for file in file_path_list:
        if file.split('\\')[-1] == 'Plate_Information.txt':
            Plate_Info_Exists = True
        if file.split('\\')[-1] == 'Re_Run_Sample_Information.txt':
            Re_Run_Sample_Info_Exists = True

    if Plate_Info_Exists == False: #Create Plate_Information.txt
        with open(Data_Directory_Path + '\\' + 'Plate_Information.txt', 'w', newline='') as Plate_Information:
            Plate_Information.write('Plate\tDate_Processed\tFile_Length')
        Plate_Information.close()

#Idea: Read File into Re_Run_Sample_Info_DATAFRAME, append dataframe, save as file:
        
    if Re_Run_Sample_Info_Exists == False: #Create Re_Run_Sample_Info.txt
        with open(Data_Directory_Path + '\\' + 'Re_Run_Sample_Info.txt', 'w', newline='') as Re_Run_Sample_Info:
            Re_Run_Sample_Info.write('Need_Re_Run_Plate\tNeed_Re_Run_Samples\tReason_Re_Run\tRe_Run_Plate\tRe_Run_Samples')
        Re_Run_Sample_Info.close()

    Plate_Info_DF = pd.read_csv(Data_Directory_Path + '\\' + 'Plate_Information.txt', sep = '\t', lineterminator = '\r', header = 0)    
    
#Initialize starting row variable outside the data file loop
    
    row_where_file_ends = Plate_Info_DF.File_Length.sum()


#Only Recent Files get processed

            
    file_path_list = [i for i in file_path_list if i.split('\\')[-1][:-4] not in list(Plate_Info_DF['Plate'])]
    file_path_list = [i for i in file_path_list if i.split('\\')[-1] != 'Plate_Information.txt']
    file_path_list = [i for i in file_path_list if i.split('\\')[-1] != 'Re_Run_Sample_Info.txt']
    
#Set Up Excel
        
    workbook_path = Data_Directory_Path + '\\' + Excel_File_Name

    if excel_file_exists == True:
        wb = load_workbook(workbook_path)
    else:
        wb = Workbook()
                
    if 'Form10_Raw Data' in wb.sheetnames:
        ws = wb['Form10_Raw Data']
    else:
        ws = wb.create_sheet('Form10_Raw Data')

    if 'Plate Report' in wb.sheetnames:
        del wb['Plate Report']

    plate_report_ws = wb.create_sheet('Plate Report')

    
#Create Back-Up Project File
    
    now = datetime.now()

    if excel_file_exists == True: #Preserve Current File
        shutil.copyfile(Data_Directory_Path + '\\' + Excel_File_Name, Data_Directory_Path + '\\' + 'Form10_BackUp_' + Excel_File_Name[:-5] + '_' + str(now.strftime('%m-%d-%Y_time_%H_%M')) + '.xlsx')

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
        

#Load Raw Data
        form10 = pd.read_csv(data_file, sep = '\t', lineterminator='\n', header = 10)
        form10_raw_extra = pd.read_csv(data_file, sep = '\t', lineterminator='\n', nrows = 9)

#Clean Raw Data
        form10.drop(['\r','FOS','HMD','LME','EW','BPR','NAW','HNS','HRN','EAF','BAF','TAF','CAF'], axis = 1, inplace = True)

        
#Write Raw Data to Excel
        for r in dataframe_to_rows(form10_raw_extra, index=False, header=True):
            ws.append(r)
            
        for r in dataframe_to_rows(form10, index=False, header=True):
            ws.append(r)
        
        print(data_file + " is written")

#Add File Information to PlateInformation.txt; Info not contained in Plate_Info_DF however
        
        with open(Data_Directory_Path + '\\' + 'Plate_Information.txt', 'a', newline = '\r') as Plate_Information:
            Plate_Information.write('\n'+data_file.split('\\')[-1][:-4]+'\t'+str(now.strftime('%Y-%m-%d %H:%M'))+'\t'+str(len(form10)+len(form10_raw_extra)+2))
        Plate_Information.close()

#Manage Paste Offset

        Paste_row = row_where_file_ends + 12

#Paste Calculation Header
        
        ws['X' + str(Paste_row-1)] = 'Sample Name'
        
        ws['Y' + str(Paste_row-1)] = 'Qty Mean'
        
        ws['Z' + str(Paste_row-1)] = 'Replicate Difference'
        
        ws['AA' + str(Paste_row-1)] = '(+) SPK Value'
        ws['AB' + str(Paste_row-1)] = 'SPK Difference'
        ws['AC' + str(Paste_row-1)] = 'Inhibition?'
        
#Paste Sample Names
        Sample_Name_List = form10func.Get_Sample_Names(form10)
        for name in Sample_Name_List:
            ws['X' + str(Paste_row)] = name
            Paste_row += 1
        
#QuantityMean Calculation and write
        
        Paste_row = row_where_file_ends + 12 #Reset Paste_Row position
            
        QuantityMeans = form10func.QtyMean_Dataframe(form10)
        for Qty_mean_value in QuantityMeans['Qty Mean']:
            ws['Y' + str(Paste_row)] = Qty_mean_value
            Paste_row += 1
    
#Replicate Diff
        Paste_row = row_where_file_ends + 12 #Reset Paste_Row position
            
        Replicate_Differences = form10func.Replicate_Diff_Dataframe(form10)
        for value in Replicate_Differences['Ct diff']:
            ws['Z' + str(Paste_row)] = value
            Paste_row+=1
#SPK Ct
        Paste_row = row_where_file_ends + 12 #Reset Paste_Row position

        if Sample_Name_List[0].startswith('A'):
            Paste_row += 1
        
        SPK_Ct = form10func.SPK_Ct_dataframe(form10)
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
        ws['U' + str(row_where_file_ends+1)] = 1 + len(Plate_Info_DF['Plate']) + file_count
        ws['U' + str(row_where_file_ends+1)].font = Font(size = 16, bold = True)
        ws['U' + str(row_where_file_ends+1)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
        ws['U' + str(row_where_file_ends+1)].alignment = Alignment(horizontal = 'center')

        row_where_file_ends = row_where_file_ends + len(form10) + len(form10_raw_extra) + 2

##############################################################################################################################

        
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

        if Sample_Name_List[0].startswith('A'):
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
        

#ReRun Status

        re_run_status = 'green'
        re_run_samples = 'not yet'
        re_run_samples_inhib = 'not yet'

        #Status Color Indicator
        plate_report_ws['B'+ str(Paste_row+1)].fill = PatternFill(start_color='0000FF00',end_color='0000FF00',fill_type='solid') #Default Green

        #STDs
        STD_additional_comment = []
        
        for x in range(0,len(Replicate_Differences['Sample Name'])):
            if Replicate_Differences['Sample Name'][x].startswith('STD'):
                if Replicate_Differences['Sample Name'][x] != 'STD8':
                    if '.' in str(Replicate_Differences['Ct diff'][x]): #Test Numbers only
                        if abs(Replicate_Differences['Ct diff'][x]) > 1:
                            plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                            plate_report_ws['D'+ str(Paste_row+2)] = Replicate_Differences['Sample Name'][x] + ' Failed' + ' '.join(STD_additional_comment)
                            STD_additional_comment.append('and' + Replicate_Differences['Sample Name'][x] + 'Failed')
                            re_run_status = 'red'
                        
        Paste_row = file_count * 40 + 1 #Reset Paste_Row position

        #Only STD8 fails plate
        for x in range(0,len(Replicate_Differences['Sample Name'])):
            if Replicate_Differences['Sample Name'][x] == 'STD8':
                if abs(Replicate_Differences['Ct diff'][x]) > 1:
                    plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
                    plate_report_ws['D'+ str(Paste_row+2)] = 'STD 8 Failed' + ' '.join(STD_additional_comment)
                    re_run_status = 'red'

                
        Paste_row = file_count * 40 + 1 #Reset Paste_Row position

        #Slope, R2, LOD

        plate_report_ws['M' + str(Paste_row+1)] = 'Slope'
        plate_report_ws['N' + str(Paste_row+1)] = 'R^2'
        plate_report_ws['O' + str(Paste_row+1)] = 'NTC'
        plate_report_ws['P' + str(Paste_row+1)] = 'LOD'
        
        plate_report_ws['M' + str(Paste_row+2)] = (form10.iloc[-8,1] + ' cycles/log decade')
        plate_report_ws['N' + str(Paste_row+2)] = form10.iloc[-6,2][:-1]
        plate_report_ws['P' + str(Paste_row+2)] = LOD

        if QuantityMeans['NTC'][0] > 0:
            plate_report_ws['O' + str(Paste_row+2)] = QuantityMeans['NTC'][0]
        else:
            plate_report_ws['O' + str(Paste_row+2)] = 'Undetermined'

        if float(form10.iloc[-8,1]) < SlopeMin:
            plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
            plate_report_ws['D'+ str(Paste_row+3)] = 'Slope Requirement'
            re_run_status = 'red'
        if float(form10.iloc[-8,1]) > SlopeMax:
            plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
            plate_report_ws['D'+ str(Paste_row+3)] = 'Slope Requirement'
            re_run_status = 'red'
        if float(form10.iloc[-6,2][:-1]) < R2:
            plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
            plate_report_ws['D'+ str(Paste_row+4)] = 'R^2 Requirement'
            re_run_status = 'red'
        if QuantityMeans['NTC'][0] < LOD:
            plate_report_ws['B'+ str(Paste_row+2)] = 'Re-run Plate'
            plate_report_ws['D'+ str(Paste_row+5)] = 'NTC vs LOD Requirement'
            re_run_status = 'red'

        Paste_row = file_count * 40 + 1 #Reset Paste_Row position

        Samples_with_replicate_diff_issue = []


        #Re_Run_Samples_Info Text Information
        #open(Data_Directory_Path + '\\' + 'Re_Run_Sample_Info.txt', 'w', newline='') as Re_Run_Sample_Info:


        
        for Replicate_value_index in range(0,len(Replicate_Differences)):
            if str(Replicate_Differences['Ct diff'][Replicate_value_index]).startswith('PROBLEM'):
                    
                Samples_with_replicate_diff_issue.append(Replicate_Differences['Sample Name'][Replicate_value_index])
                plate_report_ws['C' + str(Paste_row+6)] = ', '.join(Samples_with_replicate_diff_issue)
                plate_report_ws['D'+ str(Paste_row+6)] = 'Replicate Difference'
                
                re_run_samples = 'true'
                re_run_status = 'red'
                Paste_row += 1

   #             if str(Replicate_Differences['Sample Name'][Replicate_value_index].startswith('S')) == False and str(Replicate_Differences['Sample Name'][Replicate_value_index].startswith('N')) == False and str(Replicate_Differences['Sample Name'][Replicate_value_index].startswith('A')) == False:
   #                 print('Put something in Re_Run_Samples_DF')


        Paste_row = file_count * 40 + 1 #Reset Paste_Row position
        
        
        for SampleIndex in range(0,len(SPK_Ct['Inhibition?'])):        

            if SPK_Ct['Inhibition?'][SampleIndex] == 'True':

                plate_report_ws['C' + str(Paste_row+7)] = SPK_Ct['SPK Names'][SampleIndex]
                plate_report_ws['D'+ str(Paste_row+7)] = 'Inhibited'
                
                re_run_samples_inhib = 'true'
                re_run_status = 'red'
                Paste_row += 1
       
        Paste_row = file_count * 40 + 1 #Reset Paste_Row position
        

        if re_run_samples == 'true':
                plate_report_ws['B'+ str(Paste_row+6)] = 'Re-run Samples'

        if re_run_samples_inhib == 'true' and re_run_samples == 'not yet':
                plate_report_ws['B'+ str(Paste_row+7)] = 'Re-run Samples'

        if re_run_status == 'red':
                plate_report_ws['B'+ str(Paste_row+1)].fill = PatternFill(start_color='00FF0000',end_color='00FF0000',fill_type='solid')

        for x in Sample_Name_List[-9:]:
                if x.startswith('S') == False:
                        plate_report_ws['D' + str(Paste_row+1)] = 'Re-Run with a STD removed'

                        
#Locate Area to write next set of Calcultions
        file_count += 1


#Set Column Width for Readability

    plate_report_ws.column_dimensions['A'].width = 11
    plate_report_ws.column_dimensions['B'].width = 14
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
        
#Save!
    
    wb.save(workbook_path)
    
    if excel_file_exists == False: #First-Result Back Up
        shutil.copyfile(Data_Directory_Path + '\\' + Excel_File_Name, Data_Directory_Path + '\\' + 'Form10_BackUp_' + Excel_File_Name[:-5] + '_' + str(now.strftime('%m-%d-%Y_time_%H_%M')) + '.xlsx')

#Set the parameters for the Program
        
Data_Directory_Path = r"C:\Users\efu\Desktop\TestData"
Excel_File_Name = 'Book1.xlsx'
SlopeMin = -3.7
SlopeMax = -3.1
LOD = 35
R2 = .998

execute(Data_Directory_Path, Excel_File_Name, SlopeMin, SlopeMax, LOD, R2)
print('Done')
