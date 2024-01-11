#For Automating Analysis & COA for AvroBio

#Developed by: Ross Fu
#eric.fu@avancebio.com


#These can be updated from the database without updating this program:
#SOP, Assay Title, Testing Facility, Quality Statement, Pipeline, Vendor Version




import os
import json
import tkinter as tk
from datetime import datetime
from tkinter import simpledialog, filedialog
import mysql.connector
import xlsxwriter



# User Interface
print('Time to get your data')

root = tk.Tk()
root.withdraw()


q_statement = tk.messagebox.askyesno('','Include Quality Statement?')
    
tk.messagebox.showinfo('','Select a folder that has your data')
original_data_dir = tk.filedialog.askdirectory()





# Search Data Directory Folder
file_path_list = []
for root,dirs,files in os.walk(original_data_dir):
    for names in files:
            
        if names.endswith('.json'):
            file_path_list.append(os.path.join(root,names))
            file_path_list.sort(key=os.path.getmtime)

            
file_path_list = [x for x in file_path_list if '-' in x.split('\\')[-1] and '_' in x.split('\\')[-1]]



# Number of Files?
if len(file_path_list) == 0:
    print('No JSON files found with the expected format (containing - and _ characters)')
    print('Program terminating')
    os._exit(0)
else:
    num_files = len(file_path_list)
    num = 1





# Get Config Information from DB (SOP title etc.)
lims_db = mysql.connector.connect(host='192.168.21.239', user='efu', password='M0nkey!!', database='avancelims', auth_plugin='mysql_native_password')
lims_db_cursor = lims_db.cursor(buffered=True)

lims_db_cursor.execute("SELECT * FROM t5_coa_scripts_config WHERE Client = 'AvroBio'")
AvroBio_coa_config_settings = lims_db_cursor.fetchone()


sop = AvroBio_coa_config_settings[2]    
assay_title = AvroBio_coa_config_settings[3]
pipeline = AvroBio_coa_config_settings[4]
vendor_version = AvroBio_coa_config_settings[5]
quality_statement = AvroBio_coa_config_settings[6]
testing_facility = AvroBio_coa_config_settings[7]
save_path = AvroBio_coa_config_settings[8]





# Project ID
def getprojid():
    
    global projid_from_user_input
    
    projid_from_user_input = tk.simpledialog.askstring('','\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n                                                                                       What is your 9 digit Project ID?                                                                                       \n\n\n\n\n\n\n\n\n\n\n\n\n\n\n')

    if len(projid_from_user_input) == 9:
        pass
##        coa.merge_range('E17:H20',projid_from_user_input.upper(),normal_border_format)
    else:
        tk.messagebox.showinfo('','Enter 9 digit project ID')
        getprojid()
getprojid()






# Set up workbook
now = datetime.now().replace(microsecond=0)

if os.path.exists(save_path + '\\' + projid_from_user_input) == False:
    os.mkdir(save_path + '\\' + projid_from_user_input)
    os.mkdir('\\'.join(save_path.split('\\')[:-1]) + '\\' + 'Final_Signed_COA_PDFs' + '\\' + projid_from_user_input)
    
wb_name = save_path + '\\' + projid_from_user_input + '\\' + os.environ['USERNAME'] + '_AvroBio_COA_ ' + str(now).replace(':',';') + ' .xlsx'

workbook = xlsxwriter.Workbook(wb_name)

ws = workbook.add_worksheet('Raw Data')
finalreport_ws = workbook.add_worksheet('Final Reporting')
coa = workbook.add_worksheet('COA')




# Set Up Xlsxwriter Formats
lil_format = workbook.add_format({'font_size':8, 'font_name':'Times New Roman'})#, 'locked':True})
normal_format = workbook.add_format({'font_size':12, 'font_name':'Times New Roman'})#, 'locked':True})
normal_text_wrapped_format = workbook.add_format({'font_size':12, 'font_name':'Times New Roman'})#, 'locked':True})
normal_text_wrapped_format.set_text_wrap()
bold_format = workbook.add_format({'bold': True, 'font_size':12, 'font_name':'Times New Roman'})#, 'locked':True})

Fancy_bold_format = workbook.add_format({'bold': True, 'font_size':14, 'font_name':'Times New Roman', 'align':'center'})#, 'locked':True})
Fancy_bold_format.set_underline(1)

bold_border_format = workbook.add_format({'bold': True, 'border':1,'font_size':12, 'font_name':'Times New Roman', 'align':'left', 'valign':'vcenter'})#, 'locked':True})
normal_border_format = workbook.add_format({'border':1, 'font_size':11, 'font_name':'Times New Roman', 'align':'left', 'valign':'vcenter'})#, 'locked':True})

unlocked_border_format = workbook.add_format({'border':1, 'font_size':11, 'font_name':'Times New Roman', 'align':'left', 'valign':'vcenter'})#, 'locked':False})
##unlocked_border_format = workbook.add_format({'border':1, 'font_size':11, 'font_name':'Times New Roman', 'align':'left', 'valign':'vcenter', 'locked':False, 'bg_color':'#FFFF00'})

normal_border_format.set_text_wrap()




# Locking was necessary when the sheet was being protected. Now the strategy is to use write once directories to load and save final coas.
##coa.protect(options = {'insert_rows':True, 'format_rows':True})
##coa.protect(options = {'insert_rows':True, 'format_cells':True})





header = '&C&18&"Times New Roman, bold"\nCERTIFICATE OF ANALYSIS'
footer = '&L&G&C&12&"Times New Roman"\nPage &P of &N&R&12&"Times New Roman"\n\n\n\nCONFIDENTIAL\n\n\n\n\n\n\n\n\n'




# Page Setup
coa.set_header(header)
coa.set_footer(footer, {'image_left': 'AvanceLogo.png'})
coa.hide_gridlines(2)
coa.set_margins(top=1.2, bottom=.75)


ws.write('A1', 'Run ID', bold_format)
ws.write('B1', 'Transduction %', bold_format)
ws.write('C1', 'Read Quality (Q30)', bold_format)
ws.write('D1', '% Reads Mapped to Genome', bold_format)
ws.write('E1', '% Reads Mapped to Target', bold_format)
ws.write('F1', '% Reads Mapped to Cell', bold_format)


finalreport_ws.write('A1', 'Sample Description', bold_format)
finalreport_ws.write('B1', 'Transduction % (Rounded to whole number)', bold_format)

ws_col_widths = [9,15,19,29,29,25]
FR_ws_col_widths = [40,45]
coa_col_widths = 5


for i in range(len(ws_col_widths)):
    ws.set_column(i, i, ws_col_widths[i])

for i in range(26):
    coa.set_column(i, i, 4.4)
    
finalreport_ws.set_column(0, 0, 40)
finalreport_ws.set_column(1, 1, 45)







# Content
coa.merge_range('A1:Q3', ' Assay ID: ' + sop, bold_border_format)
coa.merge_range('A4:Q6', ' Assay Title: ' + assay_title, bold_border_format)
coa.merge_range('A7:D11',' Customer Contact', bold_border_format)
coa.merge_range('A12:D16',' Customer Address', bold_border_format)
coa.merge_range('A17:D20',' Project ID', bold_border_format)
coa.merge_range('E17:H20',projid_from_user_input.upper(),normal_border_format)

coa.merge_range('I7:L11','Testing Facility', bold_border_format)
coa.merge_range('I12:L16','Sample Receipt Date', bold_border_format)
##coa.merge_range('I17:L18','Assay Initiation Date', bold_border_format)
coa.merge_range('I17:L20','Assay Issue Date', bold_border_format)
coa.merge_range('M7:Q11', testing_facility, normal_border_format)

coa.merge_range('E12:H16','', normal_border_format)
##coa.merge_range('M17:Q18','', normal_border_format)
coa.merge_range('M17:Q20','', unlocked_border_format) #user enters assay issue date

coa.write('A23', 'All Test and Control Samples', bold_format)
coa.merge_range('A24:B25',' Identity', bold_border_format)
coa.merge_range('C24:O25','Sample Tube Label / Control Sample Description', bold_border_format)
coa.merge_range('P24:Q25','AvanceID', bold_border_format)
##coa.merge_range('E24:G25','Sample Name', bold_border_format)  #sample name (removed)



#################################################################
coa.write('A' + str((num_files * 2) + 28), 'Test Result', bold_format)
coa.write('A' + str((num_files * 2) + 29), 'A percent transduction assay was performed for samples submitted by AVROBIO, Inc.', normal_format)
coa.write('A' + str((num_files * 2) + 30), 'Percent transduction for test sample and control is listed below.', normal_format)
coa.write('A' + str((num_files * 2) + 31), 'Note: Data analysis was performed using a validated pipeline:('+pipeline+', vendor version:'+vendor_version+')', lil_format)


coa.merge_range('C' + str((num_files * 2) + 34) + ':K' + str((num_files * 2) + 35),'Sample', bold_border_format)
coa.merge_range('L' + str((num_files * 2) + 34) + ':O' + str((num_files * 2) + 35), 'Transduction %', bold_border_format)
#################################################################








# Extract
for data_file in file_path_list:
    print('\nExtracting data from file ' + str(num) + ' out of ' + str(num_files) + ': ' + data_file)

    with open(data_file) as f:
        data = json.load(f)

    # Data zz
    ws.write('A' + str(num+1), data['metadata']['sample_name'], normal_format)
    ws.write('B' + str(num+1), str(round(data['summary']['transduction_rate'],2)) + str('%'), normal_format)
    
    coa.merge_range('L' + str((num_files * 2) + (num * 2) + 34) + ':O' + str((num_files * 2) + (num * 2) + 35), str(round(data['summary']['transduction_rate'])) + str('%'), normal_border_format)

    finalreport_ws.write('B' + str(num+1), str(round(data['summary']['transduction_rate'])) + str('%'), normal_format)
    ws.write('C' + str(num+1), str(round(data['sequencing']['percentage_read_quality'],2)) + str('%'), normal_format)
    ws.write('D' + str(num+1), str(round(data['mapping']['percentage_reads_mapped_to_genome'],2)) + str('%'), normal_format)
    ws.write('E' + str(num+1), str(round(data['mapping']['percentage_reads_mapped_to_target'],2)) + str('%'), normal_format)
    ws.write('F' + str(num+1), str(round(data['cell_calling']['percentage_reads_mapped_to_cells'],2)) + str('%'), normal_format)



    order_id = data_file.split('\\')[-1]
    order_id = order_id.split('-')[0]

    cnum = data_file.split('\\')[-1]
    cnum = cnum.split('-')[-1]
    cnum = cnum.split('_')[0]

    avance_id = data_file.split('\\')[-1]
    avance_id = avance_id.split('_')[0]
    


    print('order id is: ' + str(order_id))
    print('cnum is: ' + str(cnum))
    

    lims_db_cursor.execute("SELECT samplename, tubelabel FROM samples WHERE idorders = '" + order_id + "' AND cnumber = '" + cnum + "'")
    sample_name_stuff = lims_db_cursor.fetchone()



    if lims_db_cursor.rowcount == 1:
    
        sample_name = sample_name_stuff[0]
        tube_label = sample_name_stuff[1]

        print('avance id: ' + str(avance_id))
        print('samplename: ' + str(sample_name))
        print('tubelabel: ' + str(tube_label))

            
        finalreport_ws.write('A' + str(num+1), sample_name, normal_format)
        
        coa.merge_range('C' + str((num_files * 2) + (num * 2) + 34) + ':K' + str((num_files * 2) + (num * 2) + 35), sample_name, normal_border_format) #transduction table
        
        coa.merge_range('A' + str(num * 2 + 24) + ':B' + str(num * 2 + 25), '', unlocked_border_format) #identity (filled in by user)
        coa.merge_range('C' + str(num * 2 + 24) + ':O' + str(num * 2 + 25), tube_label, normal_border_format)
        coa.merge_range('P' + str(num * 2 + 24) + ':Q' + str(num * 2 + 25), avance_id, normal_border_format)

        
        






    elif lims_db_cursor.rowcount == 0:
        coa.write('F' + str(num * 2 + 24),'?', normal_format)
        coa.merge_range('F' + str(num * 2 + 24) + ':J' + str(num * 2 + 25), tube_label, normal_border_format)
        finalreport_ws.write('A' + str(num+1), '?', normal_format)
        print('No Database results for this Order # / C#')
        tk.messagebox.showinfo('','One of your samples was not found in the database: \n\n' + data_file)
        tk.messagebox.showinfo('','COA contents will not be correct for this sample: \n\n' + data_file)


    num += 1











########


# Customer Contact Table
print('\n\nSearching Database for details on Order ID #' + str(order_id))



    # Find Order Date
lims_db_cursor.execute("SELECT samplecheckin.checkindate FROM samplecheckin INNER JOIN samples ON samplecheckin.idsamplecheckin = samples.idsamplecheckin WHERE samples.idorders = '" + str(order_id) + "'")
order_info = lims_db_cursor.fetchone()

if order_info == None:
    print('Order Date Not Found')
    coa.merge_range('M12:Q16','', normal_border_format)

else:
    order_date = order_info[0]
    formatted_date = order_date.strftime('%d %b %Y')
    
    coa.merge_range('M12:Q16',formatted_date, normal_border_format)
    print('Order Date Found')



    # Find Customer Contact
lims_db_cursor.execute("SELECT sponsors.idaccounts, sponsors.firstname, sponsors.lastname, sponsors.shippingaddress1 FROM sponsors INNER JOIN orders ON sponsors.idsponsors = orders.idsponsors WHERE orders.idorders = '" + str(order_id) + "'")
sponsor_info = lims_db_cursor.fetchone()

    # Customer Address
if sponsor_info[3] == None:
    print('Customer Address Not Found')
else:
    coa.write('E12',sponsor_info[3],normal_border_format)
    print('Customer Address Found')


    
    # Sponsor Name
firstname = sponsor_info[1]
lastname = sponsor_info[2]


    
if firstname != None and lastname != None:
    coa.merge_range('E7:H11',firstname + ' ' + lastname,normal_border_format)
    print('Sponsor info found')
    
else:
    print('Sponsor info Not Found')









# Quality Statement
if q_statement == True:
    coa.merge_range('A' + str(num_files * 4 + 38) + ':Q' + str(num_files * 4 + 38),'Quality Statement', Fancy_bold_format)

    coa.merge_range('A' + str(num_files * 4 + 40) + ':Q' + str(num_files * 4 + 55), quality_statement, normal_text_wrapped_format)

    coa.write('A' + str(num_files * 4 + 64), '_______________________________________', normal_format)
    coa.write('A' + str(num_files * 4 + 65), 'Project Manager', normal_format)
    coa.write('G' + str(num_files * 4 + 65), 'Date', normal_format)
    coa.write('A' + str(num_files * 4 + 66), 'Avance Biosciences Inc.', normal_format)

    coa.write('J' + str(num_files * 4 + 64), '______________________________________', normal_format)
    coa.write('J' + str(num_files * 4 + 65), 'Quality Unit Representative', normal_format)
    coa.write('P' + str(num_files * 4 + 65), 'Date', normal_format)
    coa.write('J' + str(num_files * 4 + 66), 'Avance Biosciences Inc.', normal_format)

else:
    coa.write('A' + str(num_files * 4 + 48), '_______________________________________', normal_format)
    coa.write('A' + str(num_files * 4 + 49), 'Project Manager', normal_format)

    coa.write('K' + str(num_files * 4 + 48), '_________________________________', normal_format)
    coa.write('K' + str(num_files * 4 + 49), 'Date', normal_format)



######








            
# Save
print('\nLoading your results')
workbook.close()
os.startfile(wb_name)

