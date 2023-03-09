import os
import csv
import shutil
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import simpledialog, filedialog


print('\nWelcome to the Quant Studio 7 Pro data reformatting application')
print('Choose a folder of QS 7 Pro Data. (CSV Format)')
print('Then choose a folder to store reformatted files')
print('The program can convert QS 7 Pro Data to match 7900 data format')
print('Which allows the QACAR program (for qPCR analysis) to process QS 7 Pro Data.\n')


root = tk.Tk()
root.withdraw()

tk.messagebox.showinfo('','Select a folder that only has your Quant Studio 7 Pro Data.')
original_data_dir = tk.filedialog.askdirectory()

tk.messagebox.showinfo('','Now select a folder to hold your re-formatted files')
reformatted_data_dir = tk.filedialog.askdirectory()


##original_data_dir = r'\\Birch\Operations\5 GlobalShare\Python Scripts\AutoForm10\QACAR\ross qacar test\data\TEST'
##reformatted_data_dir = r'\\Birch\Operations\5 GlobalShare\Python Scripts\AutoForm10\QACAR\ross qacar test\data\ConvertedQS7Pro'




file_path_list = []



#Data Directory Folder
for root,dirs,files in os.walk(original_data_dir):
    for names in files:
            
        if names.endswith('.txt') or names.endswith('.csv') or names.endswith('.xlsx'):
            file_path_list.append(os.path.join(root,names))
            file_path_list.sort(key=os.path.getmtime)

file_path_list = [i for i in file_path_list if i.split('\\')[-1] != 'Plate_Information.txt']
file_path_list = [i for i in file_path_list if i.split('\\')[-1] != 'Re_Run_Sample_Info.txt']
file_path_list = [i for i in file_path_list if i.split('\\')[-1].startswith('MODIFIED_QUANTSTUDIO_FILE') == False]
file_path_list = [i for i in file_path_list if i.split('\\')[-1].startswith('MODIFIED_7500_FILE') == False]
file_path_list = [i for i in file_path_list if i.split('\\')[-1].startswith('RE_ENCODED_7500_FILE') == False]
file_path_list = [i for i in file_path_list if i.split('\\')[-1].endswith('log.txt') == False]
file_path_list = [i for i in file_path_list if i.split('\\')[-1].startswith('Form1_Imported_Samples') == False]
file_path_list = [i for i in file_path_list if i.split('\\')[-1].endswith('_Form4_Imported_ndv_list.txt') == False]

num_files = len(file_path_list)
num = 0



#Reformat
for data_file in file_path_list:
    num += 1
    print('Converting file ' + str(num) + ' out of ' + str(num_files) + ': ' + data_file)
    
    with open(data_file, 'r', encoding='utf-8', errors='ignore') as infile, \
         open(reformatted_data_dir + '\\' + 'Converted_QS_Pro7_' + data_file.split('\\')[-1][:-4] + '.txt', 'w', encoding='utf-8') as outfile: #[:-4] slices off '.csv' from file name

    
        for index, row in enumerate(infile.readlines()):
            line = row.replace(',','\t').replace('"','').replace("'","")

            if index == 0:
                line = 'SDS 2.4\t (This is a Quant Studio 7 Pro data file that was reformatted like a 7900, to allow qPCR analysis via QACAR.)\n'
                
            elif index == 1 or index == 3 or index == 5 or index == 7 or index == 8 or index == 9 or index == 13 or index == 14 or index == 15 or index == 18 or index == 19 or index == 20 or index == 21:
                line = ''


            elif line.startswith('Well'):

                
                df = pd.read_csv(data_file, sep = ',', header = index, dtype={'Quantity': np.float64}, usecols = ['Well', 'Sample', 'Task', 'Cq', 'Cq Mean', 'Quantity', 'Quantity Mean', 'Slope','R2','Y-Intercept'])
                df[['FOS','HMD','LME','EW','BPR','NAW','HNS','HRN','EAF','BAF','TAF','CAF','Reporter','Tm Value','Tm Type','Qty StdDev','Ct Mean', 'Ct Median','Ct StdDev','Ct Type','Template Name','Baseline Type','Baseline Start','Baseline Stop','Threshold Type','Threshold']] = '0'
                df.rename(columns={'Sample':'Sample Name', 'Cq':'Ct'}, inplace = True)

                slope = df['Slope'][0]
                r2 = df['R2'][0]
                y_int = df['Y-Intercept'][0]

                df.drop(['Slope','R2','Y-Intercept'], axis = 1, inplace = True)

                # Must Re order columns
                df = df[['Well', 'Sample Name', 'Task', 'Quantity', 'Quantity Mean', 'Ct', 'Cq Mean',\
                         'FOS','HMD','LME','EW','BPR','NAW','HNS','HRN','EAF','BAF','TAF','CAF','Reporter','Tm Value','Tm Type','Qty StdDev','Ct Mean', \
                         'Ct Median','Ct StdDev','Ct Type','Template Name','Baseline Type','Baseline Start','Baseline Stop','Threshold Type','Threshold']]

                
                outfile.write('\t'.join(df.columns)) # Write in column header
                outfile.write('\t') # The problematic and required '\r' column

                for i in range(len(df['Sample Name'])):
                    if str(df['Sample Name'][i]).upper().startswith('SPK'):
                        df.iloc[i,1] = 'SPK CONTROL'
                        print('spk control reformatted')
                        
                        
                for z in range(len(df)): # Write in DF rows
                    outfile.write('\n') # Establish \r

                    for item in df.iloc[z]: 
                        outfile.write(str(item)) 
                        outfile.write('\t')
                        
                    outfile.write('0') #\r column data


                # Slope Stuff
                outfile.write('\nSlope\t' + str(slope) + '\tcycles/log decade\nY-Intercept\t\t' + str(y_int) + '\nR^2\t\t' + str(r2) + '\r\n')
                outfile.write('\n\n\n\nyahoo')


                break

            outfile.write(line)
            
        infile.close()
        outfile.close()
        

