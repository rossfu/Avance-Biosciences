import os
import sys
import cv2 as cv
import time
import shutil
import atexit
import imutils
import traceback
import threading
import pytesseract
import numpy as np
##import xlwings as xw
import tkinter as tk
##from functools import partial
from datetime import datetime, timedelta

import mysql.connector

from tkinter.messagebox import askyesno
from tkinter import simpledialog, filedialog

from PIL import ImageTk, Image
from pylibdmtx.pylibdmtx import decode



#Considerations
#######################################
#Pytesseract exe should be on network?
#LIMS connections not hardcoded (IP and credentials)
#Output Dir Vars are Hardcoded
#logo




class scnr:
    
    def __init__(self): # init self
        self.make_GUI()


    def exit_handler(self):
        print('eroor')
        exc_type, exc_value, exc_traceback = sys.exc_info()
        print(''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))
        self.run_log('\n\n######################################################\n' + ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback)))
        

    def LIMSlogin(self):
        
        self.limsloginwindowcount += 1
        if self.limsloginwindowcount > 1:
            try:
                self.login_window.destroy()
            except:
                self.limsloginwindowcount -= 1 #It was already closed by user

        # Locate txt file that contains config settings (db connection, path for imgs)
        config_path = self.config_and_log_path_on_45drives + '\\' + 'SCNR_Config.txt'
        if os.path.exists(config_path) == False:
            tk.showerror.messagebox('Error', 'Config File is unavailable, Contact IT department')
            os._exit(0)


                #Example of config file structure commented below
######################################################################
##                with open(config_path, 'w') as config:
##                    config.write('Configuration Settings for SCNR\n')
##                    config.write('avancelims,scnr_log,192.168.21.240' + ',' + '\n')
##                    config.write('avancelims,samples,192.168.21.240' + ',' + '\n')
##                    config.write('avancelims,scnr_config,192.168.21.240' + ',' + '\n') #Config table is assumed to be in same database as 'scnr_log' and 'samples'
#######################################################################
                    
            
        with open(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Config.txt', 'r') as c: # Read Config.txt, get SCNR settings
            config = c.readlines()
      
            self.scnr_log_db = config[1].split(',')[0] #AvanceLIMS - prepare to write to scnr log table
            self.scnr_log_table = config[1].split(',')[1]
            self.scnr_log_IPaddress = config[1].split(',')[2]

            self.config_tbl = config[3].split(',')[1]
            
        print('You are logging in to the ' + self.scnr_log_db + ' database at IP Address: ' + self.scnr_log_IPaddress + '\n')


        # Log in window
        self.login_window = tk.Tk()
        self.login_window.title("Database Log In")
        self.login_window.geometry("400x150")
        self.login_window.columnconfigure(0,weight=1)
        self.login_window.columnconfigure(1,weight=2)
        self.login_window.configure(bg='white')
        self.login_window.lift()
        
        self.db_user = tk.Entry(self.login_window, width = 40)
        self.db_user.grid(column=1,row=0, sticky='EW', padx = 30)
        user_label = tk.Label(self.login_window, text = 'Username', font=('Arial', 10), bg='white')
        user_label.grid(column=0,row=0)

        self.db_pw = tk.Entry(self.login_window, width = 40, show='*')
        self.db_pw.grid(column=1,row=1, sticky='EW', padx = 30)
        
        pw_label = tk.Label(self.login_window, text = 'Password', font=('Arial', 10), bg='white')
        pw_label.grid(column=0,row=1)

            
        self.login_button = tk.Button(self.login_window, text='Log In', command=lambda: self.try_login(self.db_user.get(), self.db_pw.get(), self.scnr_log_db, self.scnr_log_table, self.scnr_log_IPaddress, self.config_tbl), font=('Arial', 9), bg='white')
        self.login_button.grid(column=1,row=2, columnspan = 3)
        self.login_button.focus_set()

        self.login_forme = tk.Button(self.login_window, text='Fill Credentials', command=self.auto_login, font=('Arial', 9), bg='white')
        self.login_forme.grid(column=1,row=3, columnspan = 3)




    def try_login(self, user, pw, db, LIMS_Table, LIMS_IP, config_tbl):
        
        # LIMS Database connection(s)
        try:
            self.lims_db = mysql.connector.connect(host=LIMS_IP, user=user, password=pw, database=db, auth_plugin='mysql_native_password')
            self.lims_db_cursor = self.lims_db.cursor()
            db_insert = "INSERT INTO " + LIMS_Table + "(Time_Stamp, User, Computer_ID, Date, Time, State, Last_update_time) VALUES (%s,%s,%s,%s,%s,%s,%s)"
            db_insert_vals = (self.runtime, os.environ['USERNAME'], os.environ['COMPUTERNAME'], self.rundate, self.run_time_of_day, 'Start', datetime.now())
            self.lims_db_cursor.execute(db_insert, db_insert_vals)
            self.lims_db.commit()
            self.login_success1 = True

            print('\nDatabase credentials approved')
##            run_log('\nDatabase credentials approved')
            
            goodlogin = tk.messagebox.showinfo('Notice', 'Database credentials approved')
            self.login_window.destroy()
            self.limsloginbutton.destroy()
            
        except:
            badlogin = tk.messagebox.showinfo('Problem', 'Please check your database credentials and try again')

        if self.login_success1:# Proceed

            # Config Settings!!!
            self.lims_db_cursor.execute("SELECT * FROM " + self.config_tbl)
            config_settings = self.lims_db_cursor.fetchone()

            print('\n')
            print(config_settings)

            self.temp_scan_img_save_dir_path = config_settings[0]
            self.final_scan_img_save_dir_path = config_settings[1]
            self.pytesseract_exe_path = config_settings[2]
            
            print('\n')
            print(config_settings[0])
            
            self.mysql_ver_temp_img_save_path = self.temp_scan_img_save_dir_path.replace('\\', r'\\')
            self.mysql_ver_final_img_save_path = self.final_scan_img_save_dir_path.replace('\\', r'\\')

            self.login_success = True
            
            print('\nConfig Settings Acquired\n')
            

            # Params button
            self.param_window_count=0
            self.main_window_set_param_button = tk.Button(text='Set Parameters', command=self.set_parameters, font=('Arial', 10), bg='white')
            self.main_window_set_param_button.grid(column=1,row=14)

            # Run Log and Reset Check
            self.create_runlog_and_check_for_records_to_reset()

            # Set Params
            if self.need_reset == False:
                self.set_parameters()
            
        
    def auto_login(self):
        self.db_user.insert(0,'efu')
        self.db_pw.insert(0,'M0nkey!!')

        
        
    def make_GUI(self):
        
        # Init vars:
        self.scanning = -1
        self.current_frame = 0
        self.init_video = 0
        self.rescan1 = False
        self.rescan2 = False
        
        # Output Directory Vars
        self.local_scnr_path = r'C:\users' + '\\' + os.environ['USERNAME'] + '\\' + 'Documents' + '\\' + 'Local_SCNR_data' # My Documents
        self.config_and_log_path_on_45drives = r'\\45drives\hometree\IT\Software\Internal\AvanceLIMS\ConfigurationFiles\ClientLabelDocumentationSystem' # 45 drives
        if os.path.exists(self.config_and_log_path_on_45drives) == False:
            tk.showerror.messagebox('Error', 'Config File is unavailable, Contact IT department')
            os._exit(0)
            

        # Time Vars        
        self.runtime = datetime.now()
        self.rundate = datetime.now().strftime("%d %b %Y")
        self.run_time_of_day = datetime.now().strftime("%H:%M:%S.%f")
        

        # Set up window:
        self.window = tk.Tk()
        self.window.title("SCNR")
        self.window.geometry("1536x950")
        self.window.configure(bg='white')
        self.window.columnconfigure(0,weight=3)
        self.window.columnconfigure(1,weight=2)
        self.window.columnconfigure(2,weight=3)

        # Exit Handler
        atexit.register(self.exit_handler)

        # Use preview size to scale image depending on video resolution
        self.preview_scale = .1 # percent

        # Initialize camera properties
        self.slitwidth = 12
        self.focus = 450
        self.exposure = -8
        self.zoom = 50
        self.num_frames_to_scan = 90
        
        # Initialize checkpoint variables
        self.next_action = 'FIRST SCAN'
        self.second_scan_checkpoint = False
        self.attempt_ocr = False

        # Set Up GUI Header
        Header_Label = tk.Label(self.window, bg='white')
        Header_Label.grid(column=1,row=0)
        Header_Image = Image.open('GUIHeader.png')
        Header = ImageTk.PhotoImage(Header_Image)
        Header_Label.configure(image=Header)
        Header_Label.image = Header

        # Set up Live Preview
        live_preview_Header = tk.Label(self.window, text = 'Live Preview', font=('Arial', 20), bg='white')
        live_preview_Header.grid(column=1, row=2)
        
        self.live_preview = tk.Label(self.window, bg='white')
        self.live_preview.grid(column=1, row=3)
        
        # Set up Scan 1 preview
        Scan_Result_Header = tk.Label(self.window, text = 'Client Label', font=('Arial', 24), bg='white')
        Scan_Result_Header.grid(column=0, row=2, sticky='EW')
        
        self.scan_result = tk.Label(self.window, bg='white')
        self.scan_result.grid(column=0,row=3, sticky='EW')

        # Set up Scan 2 preview
        Scan_Result_Header2 = tk.Label(self.window, text = 'C-Number Label', font=('Arial', 24), bg='white')
        Scan_Result_Header2.grid(column=2, row=2, sticky='EW')
        
        self.scan_result2 = tk.Label(self.window, bg='white')
        self.scan_result2.grid(column=2,row=3, sticky='EW')

        # Create Start button:       
        self.scan1_button_text = tk.StringVar()
        self.scan1_button_text.set('Scan')
        self.scan1_button = tk.Button(textvariable=self.scan1_button_text, command=self.main_scan_button, font=('Arial', 10), bg='white')
        self.scan1_button.grid(column=0,row=5)

        # Instructions Label
        Instructions_header = tk.Label(self.window, text = 'Instructions:', font=('Arial', 16), bg='white')
        Instructions_header.grid(column=1, row=4)
        self.instructions = tk.Label(self.window, text = 'Select the Vial Type from the drop down list below, select the \ncorrect vial adapter and put the vial and adapter on the turntable. \nEnsure the focus is good, the vial is centered, and then click \nthe blue Scan button to capture the Client Label image.', font=('Arial', 10), bg='white')
        self.instructions.grid(column=1,row=5)

        # Vial Drop Down Menu
        Vial_Drop_Down_Menu_Header = tk.Label(self.window, text = 'Vial Type to Scan:', font=('Arial', 14), bg='white')
        Vial_Drop_Down_Menu_Header.grid(column=1, row=6)

        self.vial_types = ['1.5ml Screw Cap Vial', '1.5ml Flip Cap Vial', '2ml Screw Cap Vial', '2ml Flip Cap Vial', '2ml Cryo Vial', '5ml Screw Cap Vial', '15ml Screw Cap Vial', '50ml Screw Cap Vial']
        self.vial_type_option_var = tk.StringVar()
        self.vial_menu = tk.OptionMenu(self.window, self.vial_type_option_var, *self.vial_types)
        self.vial_menu.grid(column=1, row=7)


        # Extra Settings button
        self.settingswindowcount=0
        self.settingsbutton = tk.Button(text='Settings', command=self.settings, font=('Arial', 10), bg='white')
        self.settingsbutton.grid(column=1,row=16)

        # LIMS Login button
        self.limsloginwindowcount=0
        self.limsloginbutton = tk.Button(text='LIMS Log In', command=self.LIMSlogin, font=('Arial', 10), bg='white')
        self.limsloginbutton.grid(column=1,row=15)
        
        # Database Log In
        self.login_success = False
        self.LIMSlogin()
                
        # Start video thread:
        self.window.after(0, self.init_videoloop())
        
                
        # Main event loop:
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.window.mainloop()

        

    def settings(self):
        self.settingswindowcount +=1
        if self.settingswindowcount > 1:
            try:
                self.settings_window.destroy()
            except:
                self.settingswindowcount -= 1 #It was already closed by user

        self.settings_window = tk.Tk()
        self.settings_window.title("SCNR Settings")
        self.settings_window.geometry('400x600')
        self.settings_window.lift()


        # Options
        # Slit Width label and slider
        label = tk.Label(self.settings_window,text="Slit Width:")
        label.pack()
        self.slitwidth_scale = tk.Scale(self.settings_window, from_=1, to=100, resolution=2, orient=tk.HORIZONTAL)
        self.slitwidth_scale.set(self.slitwidth)
        self.slitwidth_scale.pack()

        # Exposure label and slider
        label = tk.Label(self.settings_window,text="Exposure:")
        label.pack()
        self.exposure_scale = tk.Scale(self.settings_window, from_=-12, to=-1, resolution=1, orient=tk.HORIZONTAL)
        self.exposure_scale.set(self.exposure)
        self.exposure_scale.pack()

        # Zoom label and slider
        label = tk.Label(self.settings_window,text="Zoom:")
        label.pack()
        self.zoom_scale = tk.Scale(self.settings_window, from_=6, to=50, resolution=2, orient=tk.HORIZONTAL)
        self.zoom_scale.set(self.zoom)
        self.zoom_scale.pack()

        # Frame Count label and slider
        label = tk.Label(self.settings_window,text="Num of frames to capture:")
        label.pack()
        self.frame_scale = tk.Scale(self.settings_window, from_=50, to=220, resolution=10, orient=tk.HORIZONTAL)
        self.frame_scale.set(self.num_frames_to_scan)
        self.frame_scale.pack()
        
        # Apply Settings
        self.apply_settings_button = tk.Button(self.settings_window, text='Apply Settings', command=self.apply_settings, font=('Arial', 10), bg='white')
        self.apply_settings_button.pack()


    def apply_settings(self):

        self.slitwidth = self.slitwidth_scale.get()
        self.exposure = self.exposure_scale.get()
        self.zoom = self.zoom_scale.get()
        self.num_frames_to_scan = self.frame_scale.get()

        print('Applied new webcam settings\n')


    def set_parameters(self):

        self.param_window_count += 1

        if self.param_window_count > 1:
            try:
                self.enter_params_window.destroy()
            except:
                self.param_window_count -= 1 #It was already closed by user


        self.enter_params_window = tk.Toplevel(self.window, bg='white')
        self.enter_params_window.title("Enter Parameters")
        self.enter_params_window.geometry("400x600")
        self.enter_params_window.lift()

        # Param Headers
        self.pciemployee_header = tk.Label(self.enter_params_window, text = 'PCI Employee:', font=('Arial', 10), bg='white')
        self.ordernumber_header = tk.Label(self.enter_params_window, text = 'Order Number:', font=('Arial', 10), bg='white')
        self.boxnumber_header = tk.Label(self.enter_params_window, text = 'Box Number:', font=('Arial', 10), bg='white')
        self.clabels_to_print_header = tk.Label(self.enter_params_window, text = 'C# Labels to Print:', font=('Arial', 10), bg='white')
        self.sampletype_header = tk.Label(self.enter_params_window, text = 'Sample Type:', font=('Arial', 10), bg='white')
        self.containertype_header = tk.Label(self.enter_params_window, text = 'Container Type:', font=('Arial', 10), bg='white')
        self.storagecond_header = tk.Label(self.enter_params_window, text = 'Storage Condition:', font=('Arial', 10), bg='white')
        self.arrivecond_header = tk.Label(self.enter_params_window, text = 'Arriving Condition:', font=('Arial', 10), bg='white')
        self.sampleretention_header = tk.Label(self.enter_params_window, text = 'Sample Retention:', font=('Arial', 10), bg='white')
        self.tissuetype_header = tk.Label(self.enter_params_window, text = 'Tissue Type:', font=('Arial', 10), bg='white')

        # Param Entries
        self.pci_employee = tk.Label(self.enter_params_window, text = os.environ['USERNAME'], font=('Arial', 10), bg='white')
##        if self.pci_employee != '' and self.clabels_to_print_entry != '' and self.boxnumber_entry != '' and self.sampletype_entry != '' and self.tissuetype_option_var.get() != '' and self.containertype_entry != '' and self.storage_cond_option_var.get() != '' and self.arrive_cond_option_var.get() != '' and self.sample_retention_option_var.get() != '':


        try:
            # Temp vars
            self.cLprint_store = self.clabels_to_print_entry.get()
            self.ordernum_store = self.ordernumber_entry.get()
            self.boxnum_store = self.boxnumber_entry.get()
            self.stype_store = self.sampletype_entry.get()
            self.ctype_store = self.containertype_entry.get()
            self.ttype_store = self.tissuetype_option_var.get()
            self.scond_store = self.storage_cond_option_var.get()
            self.acond_store = self.arrive_cond_option_var.get()
            self.sret_store = self.sample_retention_option_var.get()
        except:
            donothing=1

        try:
            # Copy existing values into entry fields
            self.ordernumber_entry = tk.Entry(self.enter_params_window, width = 20, textvariable = self.ordernum_store)
            self.boxnumber_entry = tk.Entry(self.enter_params_window, width = 20, textvariable = self.boxnum_store)
            self.clabels_to_print_entry = tk.Entry(self.enter_params_window, width = 20, textvariable = self.cLprint_store)
            self.sampletype_entry = tk.Entry(self.enter_params_window, width = 20, textvariable = self.stype_store)
            self.containertype_entry = tk.Entry(self.enter_params_window, width = 20, textvariable = self.ctype_store)

            self.tissuetype_options = ['Cells', 'Control', 'DNA', 'Feces', 'Other', 'Protein', 'Reagent', 'RNA', 'Saliva', 'Standard', 'Tissue/Blood', 'Urine', 'Virus']
            self.tissuetype_option_var = tk.StringVar()
            self.tissuetype_drop_down = tk.OptionMenu(self.enter_params_window, self.tissuetype_option_var, *self.tissuetype_options)
            self.tissuetype_option_var.set(self.ttype_store)

            self.storage_cond_options = ['-80 degrees Celsius', '-20 degrees Celsius', '4 degrees Celsius', 'RT']
            self.storage_cond_option_var = tk.StringVar()
            self.storage_cond_drop_down = tk.OptionMenu(self.enter_params_window, self.storage_cond_option_var, *self.storage_cond_options)
            self.storage_cond_option_var.set(self.scond_store)
            
            self.arrive_cond_options = ['RT', 'Dry Ice', 'Credo/Cube Cooler', 'Liquid Nitrogen', 'Cold Pack/Ice']
            self.arrive_cond_option_var = tk.StringVar()
            self.arrive_cond_drop_down = tk.OptionMenu(self.enter_params_window, self.arrive_cond_option_var, *self.arrive_cond_options)
            self.arrive_cond_option_var.set(self.acond_store)
            
            self.sample_retention_options = ['90 Days']
            self.sample_retention_option_var = tk.StringVar()
            self.sample_retention_drop_down = tk.OptionMenu(self.enter_params_window, self.sample_retention_option_var, *self.sample_retention_options)
            self.sample_retention_option_var.set(self.sret_store)

            
        except: #fresh run, no established values
            self.ordernumber_entry = tk.Entry(self.enter_params_window, width = 20)
            self.boxnumber_entry = tk.Entry(self.enter_params_window, width = 20)
            self.clabels_to_print_entry = tk.Entry(self.enter_params_window, width = 20)
            self.sampletype_entry = tk.Entry(self.enter_params_window, width = 20)
            self.containertype_entry = tk.Entry(self.enter_params_window, width = 20)

            self.tissuetype_options = ['Cells', 'Control', 'DNA', 'Feces', 'Other', 'Protein', 'Reagent', 'RNA', 'Saliva', 'Standard', 'Tissue/Blood', 'Urine', 'Virus']
            self.tissuetype_option_var = tk.StringVar()
            self.tissuetype_drop_down = tk.OptionMenu(self.enter_params_window, self.tissuetype_option_var, *self.tissuetype_options)

            self.storage_cond_options = ['-80 degrees Celsius', '-20 degrees Celsius', '4 degrees Celsius', 'RT']
            self.storage_cond_option_var = tk.StringVar()
            self.storage_cond_drop_down = tk.OptionMenu(self.enter_params_window, self.storage_cond_option_var, *self.storage_cond_options)

            self.arrive_cond_options = ['RT', 'Dry Ice', 'Credo/Cube Cooler', 'Liquid Nitrogen', 'Cold Pack/Ice']
            self.arrive_cond_option_var = tk.StringVar()
            self.arrive_cond_drop_down = tk.OptionMenu(self.enter_params_window, self.arrive_cond_option_var, *self.arrive_cond_options)

            self.sample_retention_options = ['90 Days']
            self.sample_retention_option_var = tk.StringVar()
            self.sample_retention_drop_down = tk.OptionMenu(self.enter_params_window, self.sample_retention_option_var, *self.sample_retention_options)


        self.pciemployee_header.pack()
        self.pci_employee.pack()
        self.ordernumber_header.pack()
        self.ordernumber_entry.pack()
        self.boxnumber_header.pack()
        self.boxnumber_entry.pack()
        self.clabels_to_print_header.pack()
        self.clabels_to_print_entry.pack()
        self.sampletype_header.pack()
        self.sampletype_entry.pack()
        self.tissuetype_header.pack()
        self.tissuetype_drop_down.pack()
        self.containertype_header.pack()
        self.containertype_entry.pack()
        self.storagecond_header.pack()
        self.storage_cond_drop_down.pack()
        self.arrivecond_header.pack()
        self.arrive_cond_drop_down.pack()
        self.sampleretention_header.pack()
        self.sample_retention_drop_down.pack()

        self.set_param_button = tk.Button(self.enter_params_window, text='Set Parameters', command=lambda: self.transfer_params(False), font=('Arial', 10), bg='white')
        self.set_param_button.pack()
        fill_params = tk.Button(self.enter_params_window, text='Fill Parameters', command=self.fill_params, font=('Arial', 10), bg='white')
        fill_params.pack()

        

    def fill_params(self): #Lazy dev tool
        self.ordernumber_entry.insert(0,'6')
        self.boxnumber_entry.insert(0,'5000')
        self.clabels_to_print_entry.insert(0,'100')
        self.sampletype_entry.insert(0,'Fake')
        self.containertype_entry.insert(0,'None')
        self.tissuetype_option_var.set('Feces')
        self.storage_cond_option_var.set('RT')
        self.arrive_cond_option_var.set('RT')
        self.sample_retention_option_var.set('90 Days')


    def transfer_params(self,reset):
        if reset==True: #Transfer params from LIMS
            self.previous_params_query = self.lims_db_cursor.execute("SELECT * FROM avancelims.scnr_log WHERE Time_Stamp = '" + self.unfinished_scnr_run_records[self.record_idx][0] + "'")
            self.previous_params = self.lims_db_cursor.fetchone()

            text =  'PCI Employee: ' + os.environ['USERNAME'] + '\n' +\
                    'Order Number: ' + str(self.unfinished_scnr_run_records[self.record_idx][20]) + '\n' +\
                    'Box Number: ' + str(self.unfinished_scnr_run_records[self.record_idx][21]) + '\n' +\
                    'C# Labels to Print: ' + str(self.unfinished_scnr_run_records[self.record_idx][22]) + '\n' +\
                    'Sample Type: ' + str(self.unfinished_scnr_run_records[self.record_idx][23]) + '\n' +\
                    'Tissue Type: ' + str(self.unfinished_scnr_run_records[self.record_idx][24]) + '\n' +\
                    'Container Type: ' + str(self.unfinished_scnr_run_records[self.record_idx][25]) + '\n' +\
                    'Storage Condition: ' + str(self.unfinished_scnr_run_records[self.record_idx][26]) + '\n' +\
                    'Arriving Condition: ' + str(self.unfinished_scnr_run_records[self.record_idx][27]) + '\n' +\
                    'Sample Retention: ' + str(self.unfinished_scnr_run_records[self.record_idx][28])


            self.ordernum_store = tk.StringVar(value=str(self.unfinished_scnr_run_records[self.record_idx][20]))
            self.boxnum_store = tk.StringVar(value=str(self.unfinished_scnr_run_records[self.record_idx][21]))
            self.cLprint_store = tk.StringVar(value=str(self.unfinished_scnr_run_records[self.record_idx][22]))
            self.stype_store = tk.StringVar(value=str(self.unfinished_scnr_run_records[self.record_idx][23]))
            self.ctype_store = tk.StringVar(value=str(self.unfinished_scnr_run_records[self.record_idx][25]))

            self.ttype_store = str(self.unfinished_scnr_run_records[self.record_idx][23])
            self.scond_store = str(self.unfinished_scnr_run_records[self.record_idx][25])
            self.acond_store = str(self.unfinished_scnr_run_records[self.record_idx][26])
            self.sret_store = str(self.unfinished_scnr_run_records[self.record_idx][27])


            param_header = tk.Label(self.window, text = 'Parameters:', font=('Arial', 14), bg='white')
            param_header.grid(column=1, row=12)
            self.main_window_param_list = tk.Label(self.window, text = text, font=('Arial', 8), bg='white')
            self.main_window_param_list.grid(column=1, row=13)
            self.change_param_button = tk.Button(self.window, text='Change Parameters', command=self.set_parameters, font=('Arial', 10), bg='white')

            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Order_Number = '" + str(self.unfinished_scnr_run_records[self.record_idx][20]) + \
                                        "', Box_Number = '" + str(self.unfinished_scnr_run_records[self.record_idx][21]) + \
                                        "', CLabels_to_Print = '"+str(self.unfinished_scnr_run_records[self.record_idx][22])+\
                                        "', Sample_Type = '" + str(self.unfinished_scnr_run_records[self.record_idx][23]) + \
                                        "', Tissue_Type = '" + str(self.unfinished_scnr_run_records[self.record_idx][24]) + \
                                        "', Container_Type = '" + str(self.unfinished_scnr_run_records[self.record_idx][25]) + \
                                        "', Storage_Condition = '" + str(self.unfinished_scnr_run_records[self.record_idx][26]) + \
                                        "', Arrival_Condition = '" + str(self.unfinished_scnr_run_records[self.record_idx][27]) + \
                                        "', Sample_Retention = '" + str(self.unfinished_scnr_run_records[self.record_idx][28]) + \
                                        "', Last_update_time = '" + str(datetime.now()) + "' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database

            
        elif self.ordernumber_entry.get() == '' or self.boxnumber_entry.get() == '' or self.clabels_to_print_entry.get() == '' or self.sampletype_entry.get() == '' or self.tissuetype_option_var.get() == '' or self.containertype_entry.get() == '' or self.storage_cond_option_var.get() == '' or self.arrive_cond_option_var.get() == '' or self.sample_retention_option_var.get() == '':
            tk.messagebox.showwarning('Problem', 'Fill in all of the Parameters.')
        else:
            text =  'PCI Employee: ' + os.environ['USERNAME'] + '\n' +\
                    'Order Number: ' + str(self.ordernumber_entry.get()) + '\n' +\
                    'Box Number: ' + str(self.boxnumber_entry.get()) + '\n' +\
                    'C# Labels to Print: ' + str(self.clabels_to_print_entry.get()) + '\n' +\
                    'Sample Type: ' + str(self.sampletype_entry.get()) + '\n' +\
                    'Tissue Type: ' + str(self.tissuetype_option_var.get()) + '\n' +\
                    'Container Type: ' + str(self.containertype_entry.get()) + '\n' +\
                    'Storage Condition: ' + str(self.storage_cond_option_var.get()) + '\n' +\
                    'Arriving Condition: ' + str(self.arrive_cond_option_var.get()) + '\n' +\
                    'Sample Retention: ' + str(self.sample_retention_option_var.get())


            param_header = tk.Label(self.window, text = 'Parameters:', font=('Arial', 14), bg='white')
            param_header.grid(column=1, row=12)

            try:
                self.main_window_param_list.destroy()
            except:
                donothing=1
                
            self.main_window_param_list = tk.Label(self.window, text = text, font=('Arial', 8), bg='white')
            self.main_window_param_list.grid(column=1, row=13)
            self.change_param_button = tk.Button(self.window, text='Change Parameters', command=self.set_parameters, font=('Arial', 10), bg='white')
            
            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Order_Number = '" + str(self.ordernumber_entry.get()) + '\n' +\
                                        "', Box_Number = '" + str(self.boxnumber_entry.get()) + \
                                        "', CLabels_to_Print = '" + str(self.clabels_to_print_entry.get()) + \
                                        "', Sample_Type = '" + str(self.sampletype_entry.get()) + \
                                        "', Tissue_Type = '" + str(self.tissuetype_option_var.get()) + \
                                        "', Container_Type = '" + str(self.containertype_entry.get()) + \
                                        "', Storage_Condition = '" + str(self.storage_cond_option_var.get()) + \
                                        "', Arrival_Condition = '" + str(self.arrive_cond_option_var.get()) + \
                                        "', Sample_Retention = '" + str(self.sample_retention_option_var.get()) + \
                                        "', Last_update_time = '" + str(datetime.now()) + "' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database

            self.enter_params_window.destroy()



    def create_runlog_and_check_for_records_to_reset(self):
        # Run Log
        if os.path.exists(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt'):
            os.remove(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt')
            
        with open(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt', 'w') as z:
            z.write('User: ' + os.environ['USERNAME'] + '\n')
            z.write('Computer Name: ' + os.environ['COMPUTERNAME'] + '\n\n')
            z.write(str(str(self.runtime)) + '\n\n') # Log and DB have same run time

        # Reset Code
        self.need_reset = False
        resetting = False
        self.preview_scale = .4 #hasnt gotten set yet
        check_records = self.lims_db_cursor.execute("SELECT * FROM avancelims.scnr_log WHERE State <> 'Done' AND State <> 'Start' AND State <> 'Reset Complete' AND Reset_attempt IS NULL AND Reset_Abort_Reason IS NULL")
        self.unfinished_scnr_run_records = self.lims_db_cursor.fetchall()
        self.record_idx = 0 # oldest index is 0 newest at high numbers
        five_mins_ago = datetime.now() - \
                        timedelta(minutes=5)
        
        if len(self.unfinished_scnr_run_records) > 0:# Find a record by the same user or return the oldest one if its x amount of time ago
            for i in range(len(self.unfinished_scnr_run_records)): 
                if os.environ['USERNAME'] in self.unfinished_scnr_run_records[i]:
                    self.record_idx = i
                    self.reset_previous_scan()
                    resetting = True
            if resetting == False and datetime.strptime(self.unfinished_scnr_run_records[self.record_idx][19], "%Y-%m-%d %H:%M:%S.%f") < five_mins_ago: # Was the last update more than 5 mins ago? Prevent back to back resets
                self.reset_previous_scan()

                
        
    def run_log(self, message):
        with open(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt', 'a') as z:
            z.write(message)
        with open(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt', 'r') as r: # Read Log, write to database
            log = r.read()

        self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Log = '" + log + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
        self.lims_db.commit()
            


    def init_videoloop(self):
        self.cam = cv.VideoCapture(1) # May need to change webcam source

        # Webcam check
##        ret,frame = self.cam.read()
##        try:
##            self.zoom_height,self.zoom_width,self.zoom_channels = frame.shape
##            self.cam.release()
##            self.cam = cv.VideoCapture(1)
##        except:
##            no_cam = tk.messagebox.showinfo('Problem', 'Plug in your webcam')
##            if no_cam == 'ok':
##                self.exit_handler()

                
        # Find image width/height
        print('Initializing webcam:\n')
        self.width  = int(self.cam.get(cv.CAP_PROP_FRAME_WIDTH))
        self.height = int(self.cam.get(cv.CAP_PROP_FRAME_HEIGHT))
        print(f'w: {self.width} | h: {self.height}')
        print('fps is: ' + str(self.cam.get(cv.CAP_PROP_FPS)))

                
        # Slitwidth Stuff
##        self.currentX = self.width*20 # start at right side
        self.currentX = 0 # start at left side
        self.slitpoint = (self.width//2) # Centered
        

        # Rough sketch on auto-scaling based on resolution:
        if self.init_video == 0:
            if self.height > 3500:
                self.preview_scale = .1
            elif self.height > 1900:
                self.preview_scale = .15
            elif self.height > 720:
                self.preview_scale = .2 # (?)
            elif self.height <= 720: # Webcam settings (?)
                self.preview_scale = .5
            else:
                print('woops... need to clean this up.')
            self.init_video = 1

        print(f'Preview Scale: {self.preview_scale}')
        print(f'Slit Width: {self.slitwidth}')

        
        # start a thread that constantly pools the video sensor for the most recently read frame
        print('Active Threads (before opening new thread(s)): ' + str(threading.active_count()))
        self.stopEvent = threading.Event()

##        self.cam_settings_thread = threading.Thread(target=self.cam_settings, args=())
##        self.cam_settings_thread.start()
        
        self.thread = threading.Thread(target=self.video_loop, args=())
        self.thread.start()
        
        print('Active Threads (after opening a new thread): ' + str(threading.active_count()))




    # callback for start scan button:
    # self.scanning =  1 SCANNING AND PREVIEW?
    # self.scanning =  0 DONE
    # self.scanning = -1 PREVIEW ONLY
    # self.scanning = -2 OFF
    def start_scan(self):

        # Set width at multiple of the image width (this is why the scan was not catching the repeats):
        # Make incredibly large width to account for different resolutions.. we will cut extra "black" image out:
        self.img = np.zeros((self.height, self.width*40, 3), dtype='uint8')
        
        self.current_frame = 0
        self.start_time = time.time()
 
        self.scanning = 1


        

    # Threaded version of 
    def video_loop(self):
        
        self.preview = True
        self.current_vial_type_selection = self.vial_type_option_var.get()

        try:        
            while not self.stopEvent.is_set():


                ret,frame = self.cam.read()

                
                while self.vial_type_option_var.get() != self.current_vial_type_selection:
                    if self.vial_type_option_var.get() == '1.5ml Screw Cap Vial':
                        self.current_vial_type_selection = '1.5ml Screw Cap Vial'
                        self.slitwidth=12
                        self.zoom = 22
                    elif self.vial_type_option_var.get() == '1.5ml Flip Cap Vial':
                        self.current_vial_type_selection = '1.5ml Flip Cap Vial'
                        self.slitwidth=12
                        self.zoom = 10
                    elif self.vial_type_option_var.get() == '2ml Screw Cap Vial':
                        self.current_vial_type_selection = '2ml Screw Cap Vial'
                        self.slitwidth=12
                        self.zoom = 20
                    elif self.vial_type_option_var.get() == '2ml Flip Cap Vial':
                        self.current_vial_type_selection = '2ml Flip Cap Vial'
                        self.slitwidth=12
                        self.zoom = 20
                    elif self.vial_type_option_var.get() == '2ml Cryo Vial':
                        self.current_vial_type_selection = '2ml Cryo Vial'
                        self.slitwidth=12
                        self.zoom = 20
                    elif self.vial_type_option_var.get() == '5ml Screw Cap Vial':
                        self.current_vial_type_selection = '5ml Screw Cap Vial'
                        self.slitwidth=12
                        self.zoom = 24
                    elif self.vial_type_option_var.get() == '15ml Screw Cap Vial':
                        self.current_vial_type_selection = '15ml Screw Cap Vial'
                        self.slitwidth=12
                        self.zoom = 40
                    elif self.vial_type_option_var.get() == '50ml Screw Cap Vial':
                        self.current_vial_type_selection = '50ml Screw Cap Vial'
                        self.slitwidth=12
                        self.zoom = 50


                # Focus set up
                self.cam.set(cv.CAP_PROP_FOCUS, self.focus)

                # Exposure set up
                self.cam.set(cv.CAP_PROP_EXPOSURE, self.exposure)

                # Zoom set up
                self.zoom_height,self.zoom_width,self.zoom_channels = frame.shape
                centerX,centerY = int(self.zoom_height/2),int(self.zoom_width/2)
                radiusX,radiusY=int(self.zoom*self.zoom_height/100),int(self.zoom_width/100 * self.zoom)
                minX,maxX=centerX-radiusX, centerX+radiusX
                minY,maxY=centerY-radiusY, centerY+radiusY
                cropped = frame[minX:maxX, minY:maxY]
                self.zoomed_frame = cv.resize(cropped, (self.zoom_width,self.zoom_height))

                # Draw Slitwidth Lines # Displayed +2 width to remove from final image
                cv.line(img=self.zoomed_frame, pt1=(((self.width//2)-(self.slitwidth//2)-1),self.height), pt2=(((self.width//2)-(self.slitwidth//2)-1),0), color=(255,0,0), thickness=1, lineType=8, shift=0)
                cv.line(img=self.zoomed_frame, pt1=(((self.width//2)+(self.slitwidth//2)+1),self.height), pt2=(((self.width//2)+(self.slitwidth//2)+1),0), color=(255,0,0), thickness=1, lineType=8, shift=0)

            
                if ret:

                    if self.scanning == -1:
                        
                        # Video Preview:
                        self.frame = imutils.resize(self.zoomed_frame, height=300)
                        image = cv.cvtColor(self.frame, cv.COLOR_BGR2RGB)
                        image = Image.fromarray(image)
                        image = ImageTk.PhotoImage(image)
                        self.live_preview.configure(image=image)
                        self.live_preview.image = image

                            
                    # Grab slit from photo:
                    if self.scanning == 1 and self.preview == True:
                        print (f'Processing Frame# {self.current_frame}')

                        self.frame = imutils.resize(self.zoomed_frame, height=300)
                        image = cv.cvtColor(self.frame, cv.COLOR_BGR2RGB)
                        image = Image.fromarray(image)
                        image = ImageTk.PhotoImage(image)
                        self.live_preview.configure(image=image)
                        self.live_preview.image = image

                        
                        # Build image from left to right
                        self.img[:,self.currentX:self.currentX + self.slitwidth,:] = self.zoomed_frame[:,self.slitpoint - (self.slitwidth // 2):self.slitpoint + (self.slitwidth//2),:]

                        self.currentX += self.slitwidth
                        self.current_frame += 1



                    # Stop after 200 frames:
                    if self.next_action != 'Scan2_and_display': 
                        if self.scanning == 1 and self.current_frame > self.num_frames_to_scan: # Scan1 frames
                            self.scanning = 0
                    else:
                        if self.scanning == 1 and self.current_frame > self.num_frames_to_scan: # Scan2 frames
                            self.scanning = 0



                            
                else: # end of video
                    self.scanning = 0
                    

                if self.scanning == 0: # Done Scanning
                    
                    # Get Execution time:
                    execution_time = (time.time() - self.start_time)
                    print('Execution time in seconds: ' + str(execution_time) + '\n')

                    # Crop unused black pixels from numpy array:
                    crop = self.crop_unused()

                    # Save image:
                    self.img_rgb = cv.cvtColor(crop, cv.COLOR_BGR2RGB)


                    # FIRST SCAN
                    if self.next_action == 'FIRST SCAN':

                        if os.path.exists(self.temp_scan_img_save_dir_path) == False:
                            os.mkdir(self.temp_scan_img_save_dir_path)

                        self.scan1_temp_img_path = self.temp_scan_img_save_dir_path + '\\' + os.environ['USERNAME'] + '_Scan1_' + str(self.runtime).replace(':', ';') + '.jpg'
                        Image.fromarray(self.img_rgb).save(self.scan1_temp_img_path) # Temp 45 drives

                        self.next_action = 'JUDGE SCAN 1'

                        # Update DB
                        self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET State = 'Judge Scan 1', Client_Label_Image_Path = '" + self.mysql_ver_temp_img_save_path + r'\\' + os.environ['USERNAME'] + '_Scan1_' + str(self.runtime).replace(':', ';') + '.jpg' + "', Last_update_time = '"+str(datetime.now())+"'WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
                        self.lims_db.commit()

                        # OCR
                        self.OCR(self.scan1_temp_img_path, 1)

                        # Update Instructions
                        self.instructions.configure(text='Review the Client Label image and if it is clearly legible then click \nthe green Accept button. If not, click the scan button until a \ngood image of the label is obtained.')
                        
                        # Accept button
                        self.Approve_Scan1_button = tk.Button(self.window, text='Accept', command=self.scan_approved, font=('Arial', 10), bg='green')
                        self.Approve_Scan1_button.grid(column=0,row=6)

                        # Update preview with final image:
                        
                        img = (Image.open(self.scan1_temp_img_path))
                        self.scan_img_width, self.scan_img_height = img.size
                        resized_image= img.resize((int(self.scan_img_width*self.preview_scale),int(self.scan_img_height*self.preview_scale)))
                        self.final_image= ImageTk.PhotoImage(resized_image)

                        self.scan_result.configure(image=self.final_image)
                        self.scan_result.image = self.final_image

                        self.scan1_button_text.set("Re-Scan Client Label")
                        self.scan1_button.configure(state='normal', bg='red')
                        
                    # SECOND Scan
                    else:
                        self.scan2_process()
                        
                    # Back to neutral preview mode
                    self.scanning = -1 
                    
        except Exception as e:
            print('[fin].')
            import traceback
            exc_type, exc_value, exc_traceback = sys.exc_info()
            errmsg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
            errmsg = errmsg.replace("'", '')
            print(errmsg)
            self.run_log('\n\n######################################################\n' + errmsg)



    def crop_unused(self):
        # Grayscale image, try to crop out unused pixels:
        gray = cv.cvtColor(self.img,cv.COLOR_BGR2GRAY)
        _,thresh = cv.threshold(gray,1,255,cv.THRESH_BINARY)

        # Find contours to help trim extra pixels from the image:
        contours,hierarchy = cv.findContours(thresh,cv.RETR_EXTERNAL,cv.CHAIN_APPROX_SIMPLE)
        cnt = contours[0]
        x,y,w,h = cv.boundingRect(cnt)

        # write cropped image
        crop = self.img[y:y+h,x:x+w]

        return crop



    def main_scan_button(self):

        if self.next_action == 'FIRST SCAN':            

            if self.login_success == False:
                tk.messagebox.showerror('Error', 'Log In to LIMS')
                self.LIMSlogin()

            elif self.pci_employee == '' or self.clabels_to_print_entry == '' or self.boxnumber_entry == '' or self.sampletype_entry == '' or self.tissuetype_option_var.get() == '' or self.containertype_entry == '' or self.storage_cond_option_var.get() == '' or self.arrive_cond_option_var.get() == '' or self.sample_retention_option_var.get() == '':
                tk.messagebox.showerror('Error', 'Set the parameters')
                self.set_parameters()
                
            elif self.vial_type_option_var.get() == '':
                tk.messagebox.showerror('Error', 'Set the vial type')

            # Scan 1
            else:
                self.main_window_set_param_button.destroy()
                
                print('option var: ' + self.vial_type_option_var.get() + '\n')
                if self.rescan1 == False: # Do not state vial type or action during rescan 1
                    self.run_log(self.vial_type_option_var.get() + '\n\n' + str(time.asctime()) + '. ' + self.next_action)

                self.vial_menu.configure(state='disabled')

                self.settingsbutton.configure(state='disabled') # disable settings stuff
                
                try:
                    self.settings_window.destroy()
                except:
                    do_nothing=1

                
                self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Vial_Type = '" + self.vial_type_option_var.get() + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
                self.lims_db.commit()

                self.scan1_button_text.set("Scanning..")
                self.scan1_button.configure(state='disabled')
                self.start_scan()
           
        elif self.next_action == 'JUDGE SCAN 1':
            self.scan_result.configure(image='')
            self.scan_result.image = None
            self.scan1_button_text.set('Scan')
            self.scan1_button.configure(bg='white')
            self.Approve_Scan1_button.destroy()
            self.next_action = 'FIRST SCAN' # reset scan 1

            self.settingsbutton.configure(state='normal')

            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET State = 'Start', Vial_Type = '" + self.vial_type_option_var.get() + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
            self.lims_db.commit()
                
            self.rescan1 = True
            
            self.run_log('\n\n' + time.asctime() + '. Rescan 1')


        elif self.next_action == 'SCAN 2':
            if self.rescan2 == False: # Do not log next action during rescan 2
                self.run_log('\n\n' + time.asctime() + '. ' + self.next_action)

            self.scan2_button_text.set("Scanning..")
            self.scan2_button.configure(state='disabled')

            self.settingsbutton.configure(state='disabled') # disable settings stuff

            try:
                self.settings_window.destroy()
            except:
                do_nothing=1
                
            self.start_scan()


        elif self.next_action == 'JUDGE SCAN 2':
            self.scan_result2.configure(image='')
            self.scan_result2.image = None
            self.scan2_button_text.set('Scan')
            self.scan2_button.configure(bg='white')
            self.settingsbutton.configure(state='normal')

            
            if self.one_read_correct == True:
                self.Approve_Scan2_button.destroy()

            self.next_action = 'SCAN 2' # reset scan 2

            self.rescan2 = True
            self.run_log('\n\n' + time.asctime() + '. Rescan 2')



            
    def scan_approved(self): # APPROVE RESULT
            
        if self.next_action == 'JUDGE SCAN 1': # accepted scan 1
            # Update Instructions
            self.instructions.configure(text='Now put on the C-number label on the vial, scan the c-number into \nthe box below, and place the vial back on the turntable. \nClick the Scan button once the vial is in focus and \ncentered in the Live View screen.')

            # Delete buttons
            self.scan1_button.destroy()
            self.Approve_Scan1_button.destroy()

            # C Number Scan loop (separate scanner)
            C_Number_Scan_Header = tk.Label(self.window, text = 'C Number Scan:', font = ('Arial', 15), bg='white')
            C_Number_Scan_Header.grid(column=1, row=8)
            external_scan_yesno_prompt_answer = ''
            while external_scan_yesno_prompt_answer != True:
                self.cnum_input_from_external_scanner = tk.simpledialog.askstring('C Number Label Scan', 'Please use the external scanner to read the cnum label')
                if self.cnum_input_from_external_scanner != None:
                    external_scan_yesno_prompt_answer = tk.messagebox.askyesno('Confirm', 'You scanned in C-number: ' + "'" + self.cnum_input_from_external_scanner + "'" +  ', is this the label you are going to place on the Client Label shown on the left?')
            self.Cnumber_scan = tk.Label(self.window, bg='white', text=self.cnum_input_from_external_scanner)
            self.Cnumber_scan.grid(column=1, row=9)

            try:
                self.order_num = self.cnum_input_from_external_scanner.split('-')[0]
                self.c_num = self.cnum_input_from_external_scanner.split('-')[1]
            except:
                self.order_num = self.cnum_input_from_external_scanner.split('_')[0]
                self.c_num = self.cnum_input_from_external_scanner.split('_')[1]

                
            # Save Final Scan 1 Image
            if os.path.exists(self.final_scan_img_save_dir_path + '\\' + self.order_num) == False:
                os.mkdir(self.final_scan_img_save_dir_path + '\\' + self.order_num)
            self.scan1_final_img_path = self.final_scan_img_save_dir_path + '\\' + self.order_num + '\\' + self.order_num + '_' + self.c_num + '_' + os.environ['USERNAME'] + '_Client_Label_Scan.jpg' # Windows Path
            self.mysql_client_label_path = self.mysql_ver_final_img_save_path + r'\\' + self.order_num + r'\\' + self.order_num + '_' + self.c_num + '_' + os.environ['USERNAME'] + '_Client_Label_Scan.jpg' # MySQL Path
            if self.need_reset == False:
                Image.fromarray(self.img_rgb).save(self.scan1_final_img_path) # Save to final path
            else:
                shutil.copy(self.unfinished_scnr_run_records[self.record_idx][9], self.scan1_final_img_path) # Copy file to final path


            # Database update: Cnum Label Scan, Client Label Img Path, Cnum, OrderNum, Last update time
            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET State = 'Scan 2', C_Number_Scan = '"+self.cnum_input_from_external_scanner+"', Client_Label_Image_Path = '"+self.mysql_client_label_path+"', Cnum_Label_Order_Number = "+self.order_num+", Cnum_Label_C_Number = '"+self.c_num+"', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '"+str(self.runtime)+"'") # Database
            self.lims_db.commit()

            # Next action
            self.next_action = 'SCAN 2'
            self.rescan1 = False
            
            # Create scan 2 button
            self.scan2_button_text = tk.StringVar()
            self.scan2_button_text.set('Scan')
            self.scan2_button = tk.Button(textvariable=self.scan2_button_text, command=self.main_scan_button, font=('Arial', 10))
            self.scan2_button.grid(column=2,row=5)
                
            # Log
            self.run_log('\n\n' + time.asctime() + '. Scan 1 Approved, Img 1 saved. Starting Scan 2.')


            
        elif self.next_action == 'JUDGE SCAN 2': # accepted scan 2

            # Update Instructions
            self.instructions.configure(text='The images were saved to H:\Operations\PCI\C-Labels\<OrderID>. \nClick one of the buttons to the right to continue scanning \nor close the program.')

            # Process complete for tube, reset for next tube scan.
            self.scan2_button.destroy()
            self.Approve_Scan2_button.destroy()

            self.settingsbutton.configure(state='normal') # disable settings stuff

            # New Vial/Same Vial
            self.newvialtype_button = tk.Button(text='New Vial Type', command=lambda:self.next_vial('new_vial'), font=('Arial', 10))
            self.newvialtype_button.grid(column=2, row=4, sticky='NW')

            self.samevialtype_button = tk.Button(text='Same Vial Type', command=lambda:self.next_vial('same_vial'), font=('Arial', 10))
            self.samevialtype_button.grid(column=2, row=5, sticky='NW')

            # Save Scan Image 2 and Database update
            self.scan2_final_img_path = self.final_scan_img_save_dir_path + '\\' + self.order_num + '\\' + self.order_num + '_' + self.c_num + '_' + os.environ['USERNAME'] + '_C_Number_Label_Scan.jpg' # Windows Path
            self.mysql_cnum_label_path = self.mysql_ver_final_img_save_path + r'\\' + self.order_num + r'\\' + self.order_num + '_' + self.c_num + '_' + os.environ['USERNAME'] + '_C_Number_Label_Scan.jpg' # MySQL Path

            if self.need_reset == False:
                Image.fromarray(self.img_rgb).save(self.scan2_final_img_path) # Save to final path
            else:
                Image.open(self.scan2_temp_img_path).save(self.scan2_final_img_path) # Save to final path
                
            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET CNumber_Label_Image_Path = '" + self.mysql_cnum_label_path + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
            self.lims_db.commit()
            
            # Log
            self.run_log('\n\n' + time.asctime() + '. Scan 2 Approved, Img 2 saved')


            # Transfer to LIMS
##        db_insert = "INSERT INTO scnr_log (Time_Stamp, User, Computer_ID, Date, Time, State, Last_update_time) VALUES (%s,%s,%s,%s,%s,%s,%s)"
##        db_insert_vals = (self.runtime, os.environ['USERNAME'], os.environ['COMPUTERNAME'], self.rundate, self.run_time_of_day, 'Start', datetime.now())
##        self.lims_db_cursor.execute(db_insert, db_insert_vals)
##        self.lims_db.commit()
##        



#No current plan to interact
            
            #idsamples
            #idsamplecheckin
            #idprojects
            #checkindate
            #clientprogramid
            #animalspecies
            #animalgroup
            #animalid
            #sex
            
            #datecollected
            #daycollected
            #avancedbID

#Figure out these LIMS columns
            #samplename
            
            current_record_query = self.lims_db_cursor.execute("SELECT * FROM avancelims.scnr_log WHERE Time_Stamp = '" + str(self.runtime) + "'")
            current_record = self.lims_db_cursor.fetchone()

            lims_db_insert = "INSERT INTO samples (idorders, cnumber, tubelabel, Check_In_Timestamp, employee, Log, CLabels_to_print, boxnumber, sampletype, tissuetype, containertype, storagecondition, `condition`, retentiontime) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            lims_db_insert_vals = (current_record[20].replace('\n',''), current_record[16].replace('\n',''), current_record[13], self.runtime, os.environ['USERNAME'], current_record[18], current_record[22], current_record[21], current_record[23], current_record[24], current_record[25], current_record[26], current_record[27], current_record[28])

            print(lims_db_insert, lims_db_insert_vals)
     
            self.lims_db_cursor.execute(lims_db_insert, lims_db_insert_vals)
            self.lims_db.commit()

            
            

            # DB Mark completion
            if self.need_reset == True:
                self.scan1_temp_img_path = self.temp_scan_img_save_dir_path + '\\' + os.environ['USERNAME'] + '_Scan1_' + str(self.unfinished_scnr_run_records[self.record_idx][0]).replace(':', ';') + '.jpg'
                self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Reset_attempt = 'True - Completed', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
##                self.lims_db.commit()
                self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Reset_attempt = 'Completed', Time_stamp_of_associated_reset_record = '" + str(self.runtime) + "' WHERE Time_Stamp = '" + str(self.unfinished_scnr_run_records[self.record_idx][0]) + "'") # Mark the old need reset record as completed
                self.lims_db.commit()


            # Mark Scan as Done in avancelims.scnr_log
            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET State = 'Done', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
            self.lims_db.commit()

            # Delete old images
            os.remove(self.scan1_temp_img_path)
            os.remove(self.scan2_temp_img_path)
            print('Temp output files removed')
            self.run_log('\n\n' + str(time.asctime()) + '. Temp output files removed\n\n')
            
            self.run_log(str(time.asctime()) + '. SCAN PROCESS COMPLETED\n')
            print('\nScan Process Complete. You can begin the next scan')
            
            # Reset.. the reset variable
            self.need_reset = False
            self.rescan2 = False


    def reset_prompts(self):
        reset_choice = tk.messagebox.askquestion('Incomplete Scan Detected', str(len(self.unfinished_scnr_run_records)) + ' SCNR runs are incomplete. Please complete the scan(s).\n\n' + str(self.time_ago) + ' ago ' + os.environ['USERNAME'] + ' was working on this scan from ' + os.environ['COMPUTERNAME'] + '.\n\nIf this scan is not completed the program will continue prompting for its completion.\n\nDo You intend to complete this unfinished scan now?')
        if reset_choice == 'no':

            self.need_reset = False
            reset_abort_choice = tk.messagebox.askquestion('Reset', 'Will anybody complete this scan in the future?')
            
            if reset_abort_choice == 'yes':
                tk.messagebox.showinfo('The scan will be completed in the future', 'Since you did not provide a reason, next time the program runs, it will once again ask for this scan to be completed.')
                self.un_reset()
            elif reset_abort_choice == 'no':
                reset_abort_reason = None

                while reset_abort_reason == None or len(reset_abort_reason) < 5:#tk input box: Provide a reason for permanently aborting this incomplete scan
                    reset_abort_reason = tk.simpledialog.askstring('SCNR will abort the reset process if provided a reason', 'Please type an explanation for why this incomplete scan does not need to be completed.')
                    if reset_abort_reason == None:
                        tk.messagebox.showinfo('The scan is still marked as needing completion', 'Since you did not provide a reason, next time the program runs, it will once again ask for this scan to be completed.')
                        self.un_reset()
                        break
                    elif reset_abort_reason != None and len(reset_abort_reason) < 5:
                        tk.messagebox.showinfo('Provide a longer reason', 'The length of your explanation should exceed 5 characters.')
                    elif reset_abort_reason != None and len(reset_abort_reason) > 5:
                        #database
                        self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Reset_Abort_Reason = '" + reset_abort_reason + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.unfinished_scnr_run_records[self.record_idx][0]) + "'") # Update the incomplete record in the Database
                        self.lims_db.commit()
                        self.un_reset()
                        break
        else: # Update DB based on past record

            try:
                self.enter_params_window.destroy()
            except:
                donothing=1
                
            if 'Judge Scan 1' == self.unfinished_scnr_run_records[self.record_idx][5]:
                self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Client_Label_Image_Path = '" + self.unfinished_scnr_run_records[self.record_idx][9] + "', Vial_Type = '" + self.unfinished_scnr_run_records[self.record_idx][11] + "', Last_update_time = '"+str(datetime.now()) + "' WHERE Time_Stamp = '" + str(self.runtime) + "'")
                self.lims_db.commit()
                                            
            elif 'Scan 2' == self.unfinished_scnr_run_records[self.record_idx][5]:
                self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Client_Label_Image_Path = '" + self.unfinished_scnr_run_records[self.record_idx][9] + "', Vial_Type = '" + self.unfinished_scnr_run_records[self.record_idx][11] + "', C_Number_Scan = '" + self.unfinished_scnr_run_records[self.record_idx][14] + "', Cnum_Label_Order_Number = " + self.unfinished_scnr_run_records[self.record_idx][15] + ", Cnum_Label_C_Number = " + self.unfinished_scnr_run_records[self.record_idx][16] + ", Last_update_time = '"+str(datetime.now()) + "' WHERE Time_Stamp = '" + str(self.runtime) + "'")
                self.lims_db.commit()

            elif 'Judge Scan 2' == self.unfinished_scnr_run_records[self.record_idx][5]:
                self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Client_Label_Image_Path = '" + self.unfinished_scnr_run_records[self.record_idx][9] + "', CNumber_Label_Image_Path = '" + self.unfinished_scnr_run_records[self.record_idx][10] + "', Vial_Type = '" + self.unfinished_scnr_run_records[self.record_idx][11] + "', C_Number_Scan = '" + self.unfinished_scnr_run_records[self.record_idx][14] + "', Cnum_Label_Order_Number = " + self.unfinished_scnr_run_records[self.record_idx][15] + ", Cnum_Label_C_Number = " + self.unfinished_scnr_run_records[self.record_idx][16] + ", Cnum_Label_OCR = '" + self.unfinished_scnr_run_records[self.record_idx][17] + "', Last_update_time = '" + str(datetime.now()) + "' WHERE Time_Stamp = '" + str(self.runtime) + "'")
                self.lims_db.commit()



    def un_reset(self):

        self.instructions.configure(text='Select the Vial Type from the drop down list below, select the \ncorrect vial adapter and put the vial and adapter on the turntable. \nEnsure the focus is good, the vial is centered, and then click \nthe blue Scan button to capture the Client Label image.')
        self.next_action = 'FIRST SCAN'
        self.zoom = 50
       
        if 'Judge Scan 1' == self.unfinished_scnr_run_records[self.record_idx][5]:

            print('un resetting judge scan 1\n')
            
            # Vial Type
            self.vial_menu.destroy()
            self.vial_types = ['1.5ml Screw Cap Vial', '1.5ml Flip Cap Vial', '2ml Screw Cap Vial', '2ml Flip Cap Vial', '2ml Cryo Vial', '5ml Screw Cap Vial', '15ml Screw Cap Vial', '50ml Screw Cap Vial']
            self.vial_type_option_var = tk.StringVar()
            self.vial_menu = tk.OptionMenu(self.window, self.vial_type_option_var, self.vial_types[0], *self.vial_types)
            self.vial_menu.grid(column=1, row=7)

            # Scan 1 button
            self.scan1_button_text.set("Scan")
            self.scan1_button.configure(bg='white')

            # Accept button
            self.Approve_Scan1_button.destroy()

            # Update preview with final image:
            self.scan_result.configure(image='')
            self.scan_result.image = None

            # Reset Params
            self.main_window_param_list.destroy()
            
            
        elif 'Scan 2' == self.unfinished_scnr_run_records[self.record_idx][5]:

            # Add Abort Prompt
            print('un resetting scan 2\n')

            # Scan 1 button
            self.scan1_button_text.set('Scan')
            self.scan1_button = tk.Button(textvariable=self.scan1_button_text, command=self.main_scan_button, font=('Arial', 10), bg='white')
            self.scan1_button.grid(column=0,row=5)
            
            # Delete button
            self.scan2_button.destroy()

            # Vial Type
            self.vial_menu.destroy()
            self.vial_types = ['1.5ml Screw Cap Vial', '1.5ml Flip Cap Vial', '2ml Screw Cap Vial', '2ml Flip Cap Vial', '2ml Cryo Vial', '5ml Screw Cap Vial', '15ml Screw Cap Vial', '50ml Screw Cap Vial']
            self.vial_type_option_var = tk.StringVar()
            self.vial_menu = tk.OptionMenu(self.window, self.vial_type_option_var, self.vial_types[0], *self.vial_types)
            self.vial_menu.grid(column=1, row=7)
                
            # Delete C Scan Stuff
            self.Cnumber_scan.destroy()
            self.Program_Cnumber_scan.destroy()

            # Load Scan 1 preview
            self.scan_result.configure(image='')
            self.scan_result.image = None

            # Reset Params
            self.main_window_param_list.destroy()

            
        elif 'Judge Scan 2' == self.unfinished_scnr_run_records[self.record_idx][5]:
            
            # Add Abort Prompt
            print('un resetting judge scan 2\n')

            # Vial Type
            self.vial_menu.destroy()
            self.vial_types = ['1.5ml Screw Cap Vial', '1.5ml Flip Cap Vial', '2ml Screw Cap Vial', '2ml Flip Cap Vial', '2ml Cryo Vial', '5ml Screw Cap Vial', '15ml Screw Cap Vial', '50ml Screw Cap Vial']
            self.vial_type_option_var = tk.StringVar()
            self.vial_menu = tk.OptionMenu(self.window, self.vial_type_option_var, self.vial_types[0], *self.vial_types)
            self.vial_menu.grid(column=1, row=7)

            # Scan 1 button
            self.scan1_button_text.set('Scan')
            self.scan1_button = tk.Button(textvariable=self.scan1_button_text, command=self.main_scan_button, font=('Arial', 10), bg='white')
            self.scan1_button.grid(column=0,row=5)
            
            # Delete button
            self.scan2_button.destroy()
            self.Approve_Scan2_button.destroy()

            # Delete C Scan Stuff
            self.Cnumber_scan.destroy()
            self.Program_Cnumber_scan.destroy()
            
            # Load Scan 1
            self.scan_result.configure(image='')
            self.scan_result.image = None

            # Load Scan 2
            self.scan_result2.configure(image='')
            self.scan_result2.image = None

            # Reset Params
            self.main_window_param_list.destroy()

        self.current_vial_type_selection = ''
        tk.messagebox.showinfo('Notice', 'Loading the default state. You may begin a new scan.')
        


    def reset_previous_scan(self):

        # Indexes from unfinished scnr run records
        
        # 0 = Time stamp
        # 1 = User
        # 2 = Computer ID
        # 3 = Date
        # 4 = Time
        # 5 = State
        # 6 = Reset Attempt
        # 7 = Time Stamp of Associated Reset Record
        # 8 = Reset Abort Reason
        # 9 = Client Label Image Path
        # 10 = Cnum Label Image Path
        # 11 = Vial Type
        # 12 = Excel File
        # 13 = Client Label OCR
        # 14 = C Number Scan
        # 15 = Order Num
        # 16 = C Num
        # 17 = C Number Label OCR
        # 18 = Log
        # 19 = Last Update Time
        # 20 = Order Number
        # 21 = CLabels to Print
        # 22 = Box Number
        # 23 = Sample Type
        # 24 = Tissue Type
        # 25 = Container Type
        # 26 = Storage Condition
        # 27 = Arrival Condition
        # 28 = Sample Retention
        
        print('\n')
        print(self.unfinished_scnr_run_records[self.record_idx])


        # When was last update
        last_update = datetime.strptime(self.unfinished_scnr_run_records[self.record_idx][19], "%Y-%m-%d %H:%M:%S.%f").replace(microsecond=0)
        self.time_ago = datetime.now().replace(microsecond=0) - \
                   last_update

        # Reset Notice
        reset_notice = tk.messagebox.showwarning('Notice', 'An incomplete scan was detected.')
        self.need_reset = True

        # Run Log, Add information from previous incomplete record
        
##        self.lims_db_cursor.execute("SELECT Log FROM avancelims.scnr_log WHERE Time_Stamp = '" + str(self.unfinished_scnr_run_records[self.record_idx][0]) + "'")
##        Log_Contents_from_previous_incomplete_run = self.lims_db_cursor.fetchone()
##        self.run_log('\nAn incomplete record was detected. ' + os.environ['USERNAME'] + ' will attempt to complete a record which ' + self.unfinished_scnr_run_records[self.record_idx][1] + ' began at: ' + str(self.unfinished_scnr_run_records[self.record_idx][0]) + '\n\nDocumenting the resulting attempt to complete this record.\n\nHere is a copy of the log of the incomplete record:\n############################################\n' + str(Log_Contents_from_previous_incomplete_run[0]) + '\n\n\n############################################\nThis log will now document the continuation of that record with VIALTYPE: ' + self.unfinished_scnr_run_records[self.record_idx][11] + ' & Starting from STATE: ' + self.unfinished_scnr_run_records[self.record_idx][5])
        self.run_log('\nAn incomplete record was detected. ' + os.environ['USERNAME'] + ' will attempt to complete a record which ' + self.unfinished_scnr_run_records[self.record_idx][1] + ' began at: ' + str(self.unfinished_scnr_run_records[self.record_idx][0]) + '\n\nDocumenting the resulting attempt to complete this record.\n\nHere is a copy of the log of the incomplete record:\n############################################\n' + str(self.unfinished_scnr_run_records[self.record_idx][18]) + '\n\n\n############################################\nThis log will now document the continuation of that record with VIALTYPE: ' + self.unfinished_scnr_run_records[self.record_idx][11] + ' & Starting from STATE: ' + self.unfinished_scnr_run_records[self.record_idx][5])
        
        
        # Database
        self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Reset_attempt = 'True', Time_stamp_of_associated_reset_record = '" + self.unfinished_scnr_run_records[self.record_idx][0] + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
        self.lims_db.commit()
        
        if 'Judge Scan 1' == self.unfinished_scnr_run_records[self.record_idx][5]:

            # Add Abort Prompt
            print('\nresetting judge scan 1\n')

            # Update Instructions
            self.instructions.configure(text='Review the Client Label image and if it is clearly legible then click \nthe green Accept button. If not, click the scan button until a \ngood image of the label is obtained.')

            # Next Action
            self.next_action = 'JUDGE SCAN 1'
            
            # Vial Type
            self.vial_type_option_var.set(self.unfinished_scnr_run_records[self.record_idx][11])
            self.vial_menu.configure(state='disabled')

            # Accept button
            self.Approve_Scan1_button = tk.Button(self.window, text='Accept', command=self.scan_approved, font=('Arial', 10), bg='green')
            self.Approve_Scan1_button.grid(column=0,row=6)

            # Update preview with final image:
            img = (Image.open(self.unfinished_scnr_run_records[self.record_idx][9]))
            self.scan_img_width, self.scan_img_height = img.size
            resized_image = img.resize((int(self.scan_img_width*self.preview_scale),int(self.scan_img_height*self.preview_scale)))
            self.final_image= ImageTk.PhotoImage(resized_image)

            self.scan_result.configure(image=self.final_image)
            self.scan_result.image = self.final_image

            self.scan1_button_text.set("Re-Scan Client Label")
            self.scan1_button.configure(bg='red')

            # Params
            self.transfer_params(True)

            # Paths
            self.scan1_temp_img_path = self.unfinished_scnr_run_records[self.record_idx][9]

            # OCR
            self.OCR(self.scan1_temp_img_path, 1)

            # Reset Prompts
            self.reset_prompts()
            

            
        elif 'Scan 2' == self.unfinished_scnr_run_records[self.record_idx][5]:

            # Add Abort Prompt
            print('\nresetting scan 2\n')

            # Update Instructions
            self.instructions.configure(text='Now put on the C-number label on the vial, scan the c-number into \nthe box below, and place the vial back on the turntable. \nClick the Scan button once the vial is in focus and \ncentered in the Live View screen.')

            # Delete button
            self.scan1_button.destroy()

            # Next Action
            self.next_action = 'SCAN 2'

            # Vial Type
            self.vial_type_option_var.set(self.unfinished_scnr_run_records[self.record_idx][11])
            self.vial_menu.configure(state='disabled')

            # C Number Scan
            C_Number_Scan_Header = tk.Label(self.window, text = 'C Number Scan:', font = ('Arial', 15), bg='white')
            C_Number_Scan_Header.grid(column=1, row=8)

            self.cnum_input_from_external_scanner = self.unfinished_scnr_run_records[self.record_idx][14]
            self.Cnumber_scan = tk.Label(self.window, bg='white', text=self.cnum_input_from_external_scanner)
            self.Cnumber_scan.grid(column=1, row=9)
            self.order_num = self.unfinished_scnr_run_records[self.record_idx][15]
            self.c_num = self.unfinished_scnr_run_records[self.record_idx][16]

            # Load Scan 1 preview
            img = (Image.open(self.unfinished_scnr_run_records[self.record_idx][9]))
            self.scan_img_width, self.scan_img_height = img.size
            resized_image= img.resize((int(self.scan_img_width*self.preview_scale),int(self.scan_img_height*self.preview_scale)))
            self.final_image= ImageTk.PhotoImage(resized_image)

            self.scan_result.configure(image=self.final_image)
            self.scan_result.image = self.final_image

            # Create scan 2 button
            self.scan2_button_text = tk.StringVar()
            self.scan2_button_text.set('Scan')
            self.scan2_button = tk.Button(textvariable=self.scan2_button_text, command=self.main_scan_button, font=('Arial', 10))
            self.scan2_button.grid(column=2,row=5)

            # Params
            self.transfer_params(True)

            # Reset Prompts
            self.reset_prompts()


            
        elif 'Judge Scan 2' == self.unfinished_scnr_run_records[self.record_idx][5]:
            
            # Add Abort Prompt
            print('\nresetting judge scan 2\n')

            
            # Update Instructions
            self.instructions.configure(text='Review the Client Label image and if it is clearly legible then click \nthe green Accept button. If not, click the scan button until a \ngood image of the label is obtained.')

            # Delete Scan 1 Button
            self.scan1_button.destroy()

            # Paths
            self.scan1_temp_img_path = self.unfinished_scnr_run_records[self.record_idx][9]
            self.scan2_temp_img_path = self.unfinished_scnr_run_records[self.record_idx][10]
            
            # Vial Type
            self.vial_type_option_var.set(self.unfinished_scnr_run_records[self.record_idx][11])
            self.vial_menu.configure(state='disabled')
            
            # C Number Scan loop (separate scanner)
            self.cnum_input_from_external_scanner = self.unfinished_scnr_run_records[self.record_idx][14] # reset c num scan variable
            C_Number_Scan_Header = tk.Label(self.window, text = 'C Number Scan:', font = ('Arial', 15), bg='white')
            C_Number_Scan_Header.grid(column=1, row=8)
            self.Cnumber_scan = tk.Label(self.window, bg='white', text=self.cnum_input_from_external_scanner)
            self.Cnumber_scan.grid(column=1, row=9)
            self.order_num = self.unfinished_scnr_run_records[self.record_idx][15]
            self.c_num = self.unfinished_scnr_run_records[self.record_idx][16]

            # Load Scan 1
            img = (Image.open(self.scan1_temp_img_path))
            self.scan_img_width, self.scan_img_height = img.size
            resized_image= img.resize((int(self.scan_img_width*self.preview_scale),int(self.scan_img_height*self.preview_scale)))
            self.final_image= ImageTk.PhotoImage(resized_image)
            self.scan_result.configure(image=self.final_image)
            self.scan_result.image = self.final_image

            # Load Scan 2
            img = (Image.open(self.scan2_temp_img_path))
            self.scan_img_width, self.scan_img_height = img.size
            resized_image= img.resize((int(self.scan_img_width*self.preview_scale),int(self.scan_img_height*self.preview_scale)))
            self.final_image= ImageTk.PhotoImage(resized_image)
            self.scan_result2.configure(image=self.final_image)
            self.scan_result2.image = self.final_image
            
            # Create scan 2 button
            self.scan2_button_text = tk.StringVar()
            self.scan2_button_text.set('Scan')
            self.scan2_button = tk.Button(textvariable=self.scan2_button_text, command=self.main_scan_button, font=('Arial', 10))
            self.scan2_button.grid(column=2,row=5)

            # Scan 2 stuff
            self.scan2_process()

            # Params
            self.transfer_params(True)

            # Reset Prompts
            self.reset_prompts()


    def grayscale(self, scan_num):
        if scan_num == 1:
            return cv.cvtColor(cv.imread(self.scan1_temp_img_path), cv.COLOR_BGR2GRAY)
        if scan_num == 2:
            return cv.cvtColor(cv.imread(self.scan2_temp_img_path), cv.COLOR_BGR2GRAY)


    def decode_dmatrix(self, scan_num):
        print('\n################################################################')
        print('#  Decoding Data Matrix')
        print('################################################################\n')

        # Original Img
        self.dmatrix_read = decode(cv.imread(self.scan2_temp_img_path))
        print('Un binarized Data Matrix reads: ')
        print(self.dmatrix_read)
        
##        # Binarized Image
##        gray_image = self.grayscale(2)
##        thresh, im_bw = cv.threshold(gray_image, 125, 255, cv.THRESH_BINARY) # blocky, shadowy        
##        self.dmatrix_read_binarized = decode(im_bw)
##        print('\nBinarized Data Matrix reads: ')
##        print(self.dmatrix_read_binarized)
## 
##        # Rotate 90 Img
##        img_90 = Image.open(self.scan2_temp_img_path)
##        self.dmatrix_rotate_90 = decode(img_90.rotate(90))
##        print('\n+90 Degrees Rotated Data Matrix reads: ')
##        print(self.dmatrix_rotate_90)
## 
##        # Rotate 270 Img
##        img_270 = Image.open(self.scan2_temp_img_path)
##        self.dmatrix_rotate_270 = decode(img_270.rotate(270))
##        print('\n-90 Degrees Rotated Data Matrix reads: ')
##        print(self.dmatrix_rotate_270)



    def OCR(self, path, scan_num):
        pytesseract.pytesseract.tesseract_cmd = r'C:\Users\efu\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
        
        print('\n\n################################################################')
        print('#  Reading text from Image:')
        print('################################################################')

        self.ocr = pytesseract.image_to_string(Image.open(path), config='--psm 1')# --oem 4')        
        self.ocr = self.ocr.replace("'", "")
        
        print('-----------------------------------------------------------------')
        print(self.ocr)
        print('-----------------------------------------------------------------')

        if scan_num == 1:

            #idsamples          done(autogenerated)
            #idsamplecheckin    ???
            #idprojects         ???
            #checkindate        done
            #animalspecies      done
            #animalgroup        obsolete
            #animalid           obsolete
            #sex                done
            #datecollected      need from tube?
            #daycollected       need from tube?

            gray_image = self.grayscale(1)
            thresh, im_bw = cv.threshold(gray_image, 125, 255, cv.THRESH_BINARY) # blocky, shadowy        
            self.binarized_ocr = pytesseract.image_to_string(im_bw, config='--psm 1')# --oem 4')
            
            print('\nBinarized OCR\n-----------------------------------------------------------------')
            print(self.binarized_ocr)
            print('-----------------------------------------------------------------')
            
            # Reg ex vars
            search_phrase_found,sample_type,checkin_date,animalspecies,animalgroup,animalid,sex,tissue_type,tissuesegmentnumber,datecollected,daycollected=False,'','','','','','','','','',''

            # Check the search phrase table
            self.lims_db_cursor.execute("SELECT * FROM determine_sample_characteristics")
            extract_sample_info_table = self.lims_db_cursor.fetchall()

            # Init vars
            checkin_date = datetime.now().strftime('%Y-%m-%d')
            tissuesegmentnumber = 'O'

            # Animal Species
            if 'male' in str(self.ocr).lower() or ' m ' in str(self.ocr).lower():
                sex = 'male'
            elif 'female' in str(self.ocr).lower() or ' f ' in str(self.ocr).lower():
                sex = 'female'

            # Gender
            if 'human' in str(self.ocr).lower():
                animalspecies='Human'
            elif 'mouse' in str(self.ocr).lower():
                animalspecies='Mouse'
            elif 'rat' in str(self.ocr).lower():
                animalspecies='Rat'
            elif 'monkey' in str(self.ocr).lower():
                animalspecies='Monkey'
            elif 'swine' in str(self.ocr).lower():
                animalspecies='Swine'
            elif 'rabbit' in str(self.ocr).lower():
                animalspecies='Rabbit'
            elif 'human' in str(self.ocr).lower():
                animalspecies='Human'

            # Day Collected
##            for char in range(len(str(self.ocr).lower()):
##                if char
##
##
##            if 'week' in str(self.ocr).lower():
##                    
##            elif 'day' in str(self.ocr).lower():


            # Locate search phrase
            for i in range(len(extract_sample_info_table)):
                if extract_sample_info_table[i][0] in str(self.ocr):
                    
                    search_phrase_found = True

                    tissuesegmentnumber = extract_sample_info_table[i][1]
                    sample_type = extract_sample_info_table[i][2]
                    tissue_type = extract_sample_info_table[i][3]
                    
                    break

            if search_phrase_found == False: #Just put OCR into LIMS
                sample_type = "Other"
                tissue_type = "Other"
                print('\nNo search phrases found')
            
##            else: #Put other fields into LIMS too
##                self.lims_db.execute("UPDATE avancelims.samples SET sample_type 
                
            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Client_Label_OCR = '" + self.ocr + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
            
##            self.user_confirmed_tube_label = tk.simpledialog.askstring("Tube Label", "Please correct or confirm the Tube Label for this sample")
            
        elif scan_num == 2:
            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Cnum_Label_OCR = '" + self.ocr + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
        self.lims_db.commit()


    def next_vial(self, vial_type):
        
        # Delete old results.
        self.scan_result.configure(image='')
        self.scan_result.image = None
        
        self.scan_result2.configure(image='')
        self.scan_result2.image = None

        self.newvialtype_button.destroy()
        self.samevialtype_button.destroy()

        self.Cnumber_scan.destroy()
        self.Program_Cnumber_scan.destroy()
                # Reset instructions
        self.instructions.configure(text='Select the Vial Type from the drop down list below, select the \ncorrect vial adapter and put the vial and adapter on the turntable. \nEnsure the focus is good, the vial is centered, and then click \nthe blue Scan button to capture the Client Label image.')            

        # Reset Scan button
##        self.scan1_button_text = tk.StringVar()
        self.scan1_button_text.set('Scan')
        self.scan1_button = tk.Button(textvariable=self.scan1_button_text, command=self.main_scan_button, font=('Arial', 10))
        self.scan1_button.grid(column=0,row=5)

        # Vial Drop Down Menu
        if vial_type == 'new_vial':
            self.vial_menu.destroy()
##            Vial_Drop_Down_Menu_Header = tk.Label(self.window, text = 'Vial Type to Scan:', font=('Arial', 14), bg='white')
##            Vial_Drop_Down_Menu_Header.grid(column=1, row=6)
            vial_types = ['1.5ml Screw Cap Vial', '1.5ml Flip Cap Vial', '2ml Screw Cap Vial', '2ml Flip Cap Vial', '2ml Cryo Vial', '5ml Screw Cap Vial', '15ml Screw Cap Vial', '50ml Screw Cap Vial']
            self.vial_type_option_var = tk.StringVar()
            self.vial_menu = tk.OptionMenu(self.window, self.vial_type_option_var, vial_types[0], *vial_types)
            self.vial_menu.grid(column=1, row=7)
            self.current_vial_type_selection = ''


        # Reset Time Vars
        self.runtime = datetime.now()
        self.rundate = datetime.now().strftime("%d %b %Y")
        self.run_time_of_day = datetime.now().strftime("%H:%M:%S.%f")

        # Set Up Next row in DB
        db_insert = "INSERT INTO scnr_log (Time_Stamp, User, Computer_ID, Date, Time, State, Last_update_time) VALUES (%s,%s,%s,%s,%s,%s,%s)"
        db_insert_vals = (self.runtime, os.environ['USERNAME'], os.environ['COMPUTERNAME'], self.rundate, self.run_time_of_day, 'Start', datetime.now())

        self.lims_db_cursor.execute(db_insert, db_insert_vals)
        self.lims_db.commit()

        # Vars
        self.next_action = "FIRST SCAN"
        self.scanning = -1
        
        # Yes
        self.create_runlog_and_check_for_records_to_reset()
        

        
    def scan2_process(self):
        
        self.next_action = 'JUDGE SCAN 2'

        if self.need_reset == False: # Already exists when resetting judge scan 2
            self.scan2_temp_img_path = self.temp_scan_img_save_dir_path + '\\' + os.environ['USERNAME'] + '_Scan2_' + str(self.runtime).replace(':', ';') + '.jpg'
            Image.fromarray(self.img_rgb).save(self.scan2_temp_img_path)
        elif self.need_reset == True:
            if self.unfinished_scnr_run_records[self.record_idx][5] != 'Judge Scan 2':
                self.scan2_temp_img_path = self.temp_scan_img_save_dir_path + '\\' + os.environ['USERNAME'] + '_Scan2_' + str(self.runtime).replace(':', ';') + '.jpg'
                Image.fromarray(self.img_rgb).save(self.scan2_temp_img_path)

            
        # Update preview with final 2 images:
        img2 = (Image.open(self.scan2_temp_img_path))
        self.scan_img2_width, self.scan_img2_height = img2.size
        resized_image2= img2.resize((int(self.scan_img2_width*self.preview_scale),int(self.scan_img2_height*self.preview_scale)))
        self.final_image2= ImageTk.PhotoImage(resized_image2)

        self.scan_result2.configure(image=self.final_image2)
        self.scan_result2.image = self.final_image2

        # Data Matrix
        self.decode_dmatrix(2)
        
        # Data Matrix Scan Comparison
        Program_C_Number_Scan_Header = tk.Label(self.window, text = 'C Number Read by CLDS:', font = ('Arial', 15), bg='white')
        Program_C_Number_Scan_Header.grid(column=1, row=10)

        dmatrix_was_read = False
        binarized_dmatrix_was_read = False
        dmatrix_rotate_90_was_read = False
        dmatrix_rotate_270_was_read = False
        self.one_read_correct = False
        
        if len(self.dmatrix_read) > 0:
            dmatrix_was_read = True

##        if len(self.dmatrix_read_binarized) > 0:
##            binarized_dmatrix_was_read = True
            
        if dmatrix_was_read:
            for i in range(len(self.dmatrix_read)):
                if str(self.dmatrix_read[i].data) == ("b'"+self.cnum_input_from_external_scanner+"'") and self.one_read_correct != True:                                    
                    self.Program_Cnumber_scan = tk.Label(self.window, bg='white', text=self.cnum_input_from_external_scanner) # scan read matched, copy value
                    self.Program_Cnumber_scan.grid(column=1, row=11)                    
                    self.one_read_correct = True

##                        elif binarized_dmatrix_was_read and self.one_read_correct != True:
##                            for i in range(len(self.dmatrix_read_binarized)):
##                                if str(self.dmatrix_read_binarized[i].data) == ("b'"+self.cnum_input_from_external_scanner+"'") and self.one_read_correct != True:                                    
##                                    self.Program_Cnumber_scan = tk.Label(self.window, bg='white', text=self.cnum_input_from_external_scanner) # scan read matched, copy value
##                                    self.Program_Cnumber_scan.grid(column=1, row=11)                                                                               
##                                    self.one_read_correct = True
## 
##                        elif dmatrix_rotate_90_was_read and self.one_read_correct != True:
##                            for i in range(len(self.dmatrix_rotate_90)):
##                                if str(self.dmatrix_rotate_90[i].data) == ("b'"+self.cnum_input_from_external_scanner+"'") and self.one_read_correct != True:                                    
##                                    self.Program_Cnumber_scan = tk.Label(self.window, bg='white', text=self.cnum_input_from_external_scanner) # scan read matched, copy value
##                                    self.Program_Cnumber_scan.grid(column=1, row=11)                                 
##                                    self.one_read_correct = True
## 
##                        elif dmatrix_rotate_270_was_read and self.one_read_correct != True:
##                            for i in range(len(self.dmatrix_rotate_90)):
##                                if str(self.dmatrix_rotate_270[i].data) == ("b'"+self.cnum_input_from_external_scanner+"'") and self.one_read_correct != True:                                    
##                                    self.Program_Cnumber_scan = tk.Label(self.window, bg='white', text=self.cnum_input_from_external_scanner) # scan read matched, copy value
##                                    self.Program_Cnumber_scan.grid(column=1, row=11)                                                                               
##                                    self.one_read_correct = True
                    

        if self.one_read_correct == False:
            print('\nThe programs data matrix reads did not match the original read. Please rescan the Cnumber label')

        else:
            self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET State = 'Judge Scan 2', CNumber_Label_Image_Path = '" + self.mysql_ver_temp_img_save_path + r'\\' + os.environ['USERNAME'] + '_Scan2_' + str(self.runtime).replace(':', ';') + '.jpg' + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
            self.lims_db.commit()

            # Update Instructions
            self.instructions.configure(text='Review the C-Number Label image and if it is clearly legible \nthen click the green Accept button. If not, click the scan \nbutton until a good image of the label is obtained.')
            
            # OCR
            self.OCR(self.scan2_temp_img_path, 2)
            
            # Accept button
            self.Approve_Scan2_button = tk.Button(self.window, text='Accept', command=self.scan_approved, font=('Arial', 10), bg='green')
            self.Approve_Scan2_button.grid(column=2,row=6)

        # Rescan 2
        self.scan2_button_text.set("Re-Scan C-Number Label")
        self.scan2_button.configure(state='normal', bg='red')


        

            
    # Clean up before exit:
    def on_closing(self):
        print('on closing')

        self.window.after(1, self.stopEvent.set())
        self.cam.release()
        self.window.after(20, self.window.destroy)

        self.run_log('\n\n' + str(time.asctime()) + '. User closed the window')
        
        print('Active Threads (at very end of program): ' + str(threading.active_count()))
        print('Done')

##
##        
##        import traceback
##        exc_type, exc_value, exc_traceback = sys.exc_info()
##        error_string = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
##
##        scnr_db_connect = mysql.connector.connect(host="192.168.21.96",user="teracopy",password="T3RA_copy!",database="avancelims", auth_plugin='mysql_native_password')
##        db_cursor = scnr_db_connect.cursor()
##        db_cursor.execute("UPDATE avancelims.scnr_log SET Reset_attempt = 'True but error prevented reset completion' WHERE Time_Stamp = '" + str(scnr().runtime) + "'") # Database
##        with open(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt', 'a') as z:
##            z.write('\n\n################################\n' + error_string + '\n\n################################')
##        with open(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt', 'r') as r: # Read Log, write to database
##            log = r.read()
##        db_cursor.execute("UPDATE avancelims.scnr_log SET Log = '" + log + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
##        scnr_db_connect.commit()
##        os._exit(0)
##
##
##
##    def exit_handler(self, err_msg):
##        print('exiting')
##        exc_type, exc_value, exc_traceback = sys.exc_info()
##        err_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
##        
##        print(err_msg)
##        self.run_log('\n\n############################################\n' + err_msg)
##
##        self.lims_db_cursor.execute("UPDATE avancelims.scnr_log SET Scan_In_Progress = 'False' WHERE Time_Stamp = '"+str(self.runtime)+"'") # Database
##        self.lims_db.commit()
        


# MAIN:
if __name__ == '__main__':
##    try:


    
    #Validation
    import platform
    import pkg_resources
    

    if platform.system() != 'Windows':
        print('')
        print('This program is only validated on Windows OS, You are using ' + str(platform.system()))
        run_log('This program is only validated on Windows OS, You are using ' + str(platform.system()))
        print('')

    if platform.python_version() != '3.9.6':
        print('')
        print('This program is only validated on Python ver 3.9.6, You are using ' + str(platform.python_version()))
        run_log('This program is only validated on Python ver 3.9.6, You are using ' + str(platform.python_version()))
        print('')

##    if xw.__version__ != '0.25.3':
##        print('')
##        print('This program is only validated on XLwings ver 0.25.3, You are using ' + str(xw.__version__))
##        run_log('This program is only validated on XLwings ver 0.25.3, You are using ' + str(xw.__version__))
##        print('')

    if pkg_resources.get_distribution("DateTime").version != '4.3':
        print('')
        print('This program is only validated on DateTime version 4.3, You are using ' + str(pkg_resources.get_distribution("DateTime").version))
        run_log('This program is only validated on DateTime version 4.3, You are using ' + str(pkg_resources.get_distribution("DateTime").version))
        print('')

    scnr()





##    except BaseException as e:
##        print('exception')
##        import traceback
##        exc_type, exc_value, exc_traceback = sys.exc_info()
##        error_string = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
##
##        print('eroooooor')
##        print(error_string)
##
##        scnr_db_connect = mysql.connector.connect(host="192.168.21.96",user="teracopy",password="T3RA_copy!",database="avancelims", auth_plugin='mysql_native_password')
##        db_cursor = scnr_db_connect.cursor()
##        db_cursor.execute("UPDATE avancelims.scnr_log SET Reset_attempt = 'True but error prevented reset completion' WHERE Time_Stamp = '" + str(scnr().runtime) + "'") # Database
##        with open(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt', 'a') as z:
##            z.write('\n\n################################\n' + error_string + '\n\n################################')
##        with open(self.config_and_log_path_on_45drives + '\\' + 'SCNR_Log.txt', 'r') as r: # Read Log, write to database
##            log = r.read()
##        db_cursor.execute("UPDATE avancelims.scnr_log SET Log = '" + log + "', Last_update_time = '"+str(datetime.now())+"' WHERE Time_Stamp = '" + str(self.runtime) + "'") # Database
##        scnr_db_connect.commit()
