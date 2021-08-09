from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd
import os
from pathlib import Path
import form10func


def execute(data_directory_path, workbook):

#Find and Sort .TXT Raw Data Files
    file_name_list = []
    
    for root,dirs,files in os.walk(data_directory_path):
        for names in files:
            if names.endswith('.txt'):
                file_name_list.append(os.path.join(root,names))
                file_name_list.sort(key=os.path.getmtime)
    
#Iterate over every file
    file_count = 0
    file_length = 115
    
#Set Up Excel
        
    workbook_path = data_directory_path + '\\' + workbook
    wb = load_workbook(workbook_path)
    ws = wb['Yur']
    
    for data_file in file_name_list:

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

#Paste Calculation Header

        Paste_row = file_count * file_length + 12
        
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
        Paste_row = file_count * file_length + 12 #Reset Paste_Row position
            
        QuantityMeans = form10func.QuantityMean(form10)
        for key,value in QuantityMeans.items():
            ws['Y' + str(Paste_row)] = value
            Paste_row += 1
    
#Replicate Diff
        Paste_row = file_count * file_length + 12 #Reset Paste_Row position
            
        Replicate_Differences = form10func.Replicate_Diff(form10)
        for key,value in Replicate_Differences.items():
            ws['Z' + str(Paste_row)] = value
            Paste_row+=1
#SPK Ct
        Paste_row = file_count * file_length + 12 #Reset Paste_Row position

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
        ws['U' + str(file_count*file_length+1)] = file_count+1
        ws['U' + str(file_count*file_length+1)].font = Font(size = 16, bold = True)
        ws['U' + str(file_count*file_length+1)].fill = PatternFill(start_color='00FFFF00',end_color='00FFFF00',fill_type='solid')
        ws['U' + str(file_count*file_length+1)].alignment = Alignment(horizontal = 'center')
 
#Locate Area to write next set of Calcultions
        file_count += 1

#Save!
        wb.save(workbook_path)

execute(r"C:\Users\efu\Desktop\Test\Ruzz", 'Book1.xlsx')
print('Done')
