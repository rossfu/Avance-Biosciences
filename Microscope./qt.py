import os, sys, swiftcam
from PyQt5.QtCore import pyqtSignal, pyqtSlot, QTimer, QSignalBlocker, Qt
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtWidgets import QLabel, QApplication, QWidget, QDesktopWidget, QCheckBox, QMessageBox, QMainWindow, QPushButton, QComboBox, QSlider, QGroupBox, QGridLayout, QBoxLayout, QHBoxLayout, QVBoxLayout, QMenu, QAction, QLineEdit

import tkinter as tk
import mysql.connector
from datetime import datetime





   
def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)
    
    

class LoginWindow(QWidget):
    
    def __init__(self):
        super().__init__()
        
        #Config
        config_path = r'\\45drives\hometree\IT\Software\Internal\AvanceLIMS\ConfigurationFiles\MicroscopeAuditTrailProgram\Microscope_AT_Config.txt'
        if os.path.exists(config_path) == False:
            print('Config file does not exist. Contact IT.')
            os._exit(0)
            
####      Current Contents of Config File: \\45drives\hometree\IT\Software\Internal\AvanceLIMS\ConfigurationFiles\MicroscopeAuditTrailProgram
####        Configuration Settings for Microscope
####        avancelims,microscope_cam_audit_trail,192.168.21.240,
####        avancelims,microscope_cam_config,192.168.21.240,
        

        # Read Config
        with open(config_path, 'r') as c: # Read Config.txt, get SCNR settings
            config = c.readlines()
            
            # Audit Trail table
            self.microscope_audit_trail_db = config[1].split(',')[0]
            self.microscope_audit_trail_table = config[1].split(',')[1]
            self.microscope_audit_trail_IPaddress = config[1].split(',')[2]

            # Config table
            self.config_db = config[2].split(',')[0]
            self.config_table = config[2].split(',')[1]
            self.config_IPaddress = config[2].split(',')[2]

        
        self.setWindowTitle('Login Form')
        self.resize(500,120)

        layout = QGridLayout()
        label_name = QLabel('<font size = "4"> Username </font>')
        self.lineEdit_username = QLineEdit()
        self.lineEdit_username.setPlaceholderText('Please enter your LIMS username')
        layout.addWidget(label_name, 0, 0)
        layout.addWidget(self.lineEdit_username, 0, 1)

        label_password = QLabel('<font size = "4"> Password </font>')
        self.lineEdit_password = QLineEdit()
        self.lineEdit_password.setEchoMode(QLineEdit.Password)
        self.lineEdit_password.setPlaceholderText('Please enter your LIMS password')
        layout.addWidget(label_password, 1, 0)
        layout.addWidget(self.lineEdit_password, 1, 1)

        label_projid = QLabel('<font size = "4"> Project ID </font>')
        self.lineEdit_projid = QLineEdit()
        self.lineEdit_projid.setPlaceholderText('Please enter your Project ID')
        layout.addWidget(label_projid, 2, 0)
        layout.addWidget(self.lineEdit_projid, 2, 1)


        # Changing Database will require re-validation
##        self.db = 'avancelims'
##        self.table = 'microscope_cam_audit_trail'
##        self.IP = 'limsnewprod.avancebio.com'
        

        button_login = QPushButton('Login')
        button_login.clicked.connect(lambda: self.try_login(self.lineEdit_username.text(), self.lineEdit_password.text(), self.lineEdit_projid.text(), self.microscope_audit_trail_db, self.microscope_audit_trail_table, self.microscope_audit_trail_IPaddress)) #Config and audit trail assumed to be in same db
        layout.addWidget(button_login, 3, 0, 1, 2)
        layout.setRowMinimumHeight(3, 75)
        self.setLayout(layout)





    def try_login(self, user, pw, proj_id, db, LIMS_Table, LIMS_IP):

        
        # LIMS Database connection(s)
        try:
            global lims_db
            global lims_db_cursor
            global runtime
            global projid

            projid = proj_id
            runtime = datetime.now()
        
            lims_db = mysql.connector.connect(host=LIMS_IP, user=user, password=pw, database=db, auth_plugin='mysql_native_password')
            lims_db_cursor = lims_db.cursor()
            db_insert = "INSERT INTO " + LIMS_Table + "(Time_Stamp, User, Computer, Project_ID) VALUES (%s,%s,%s,%s)"
            db_insert_vals = (runtime, os.environ['USERNAME'], os.environ['COMPUTERNAME'], projid)
            lims_db_cursor.execute(db_insert, db_insert_vals)
            lims_db.commit()
            login_success = True
            
            print('Database credentials approved')

            self.close()
            self.LoginWindow = None
            
        except:
            print('Please check your database credentials and try again')
            login_success = False
            

        if login_success:# Proceed

            # Config Settings!!!
            lims_db_cursor.execute("SELECT * FROM " + self.config_table)
            self.config_settings = lims_db_cursor.fetchone()

            global microscope_img_path
            global mysql_ver_microscope_img_path
            
            microscope_img_path = self.config_settings[0]
            mysql_ver_microscope_img_path = microscope_img_path.replace('\\', r'\\')

            btn_snap.setEnabled(True)



    
class MainWindow(QMainWindow):
    evtCallback = pyqtSignal(int)

    @staticmethod
    def makeLayout(lbl_1, sli_1, val_1, lbl_2, sli_2, val_2):
        hlyt_1 = QHBoxLayout()
        hlyt_1.addWidget(lbl_1)
        hlyt_1.addStretch()
        hlyt_1.addWidget(val_1)
        hlyt_2 = QHBoxLayout()
        hlyt_2.addWidget(lbl_2)
        hlyt_2.addStretch()
        hlyt_2.addWidget(val_2)
        vlyt = QVBoxLayout()
        vlyt.addLayout(hlyt_1)
        vlyt.addWidget(sli_1)
        vlyt.addLayout(hlyt_2)
        vlyt.addWidget(sli_2)
        return vlyt

    def __init__(self):
        super().__init__()
        self.setMinimumSize(1024, 768)
        self.hcam = None
        self.timer = QTimer(self)
        self.imgWidth = 0
        self.imgHeight = 0
        self.pData = None
        self.res = 0
        self.temp = swiftcam.SWIFTCAM_TEMP_DEF
        self.tint = swiftcam.SWIFTCAM_TINT_DEF
        self.count = 0

        gbox_res = QGroupBox("Resolution")
        self.cmb_res = QComboBox()
        self.cmb_res.setEnabled(False)
        vlyt_res = QVBoxLayout()
        vlyt_res.addWidget(self.cmb_res)
        gbox_res.setLayout(vlyt_res)
        self.cmb_res.currentIndexChanged.connect(self.onResolutionChanged)

        gbox_exp = QGroupBox("Exposure")
        self.cbox_auto = QCheckBox()
        self.cbox_auto.setEnabled(False)
        lbl_auto = QLabel("Auto exposure")
        hlyt_auto = QHBoxLayout()
        hlyt_auto.addWidget(self.cbox_auto)
        hlyt_auto.addWidget(lbl_auto)
        hlyt_auto.addStretch()
        lbl_time = QLabel("Time(us):")
        lbl_gain = QLabel("Gain(%):")
        self.lbl_expoTime = QLabel("0")
        self.lbl_expoGain = QLabel("0")
        self.slider_expoTime = QSlider(Qt.Horizontal)
        self.slider_expoGain = QSlider(Qt.Horizontal)
        self.slider_expoTime.setEnabled(False)
        self.slider_expoGain.setEnabled(False)
        vlyt_exp = QVBoxLayout()
        vlyt_exp.addLayout(hlyt_auto)
        vlyt_exp.addLayout(self.makeLayout(lbl_time, self.slider_expoTime, self.lbl_expoTime, lbl_gain, self.slider_expoGain, self.lbl_expoGain))
        gbox_exp.setLayout(vlyt_exp)
        self.cbox_auto.stateChanged.connect(self.onAutoExpo)
        self.slider_expoTime.valueChanged.connect(self.onExpoTime)
        self.slider_expoGain.valueChanged.connect(self.onExpoGain)

        gbox_wb = QGroupBox("White balance")
        self.btn_autoWB = QPushButton("White balance")
        self.btn_autoWB.setEnabled(False)
        self.btn_autoWB.clicked.connect(self.onAutoWB)
        lbl_temp = QLabel("Temperature:")
        lbl_tint = QLabel("Tint:")
        self.lbl_temp = QLabel(str(swiftcam.SWIFTCAM_TEMP_DEF))
        self.lbl_tint = QLabel(str(swiftcam.SWIFTCAM_TINT_DEF))
        self.slider_temp = QSlider(Qt.Horizontal)
        self.slider_tint = QSlider(Qt.Horizontal)
        self.slider_temp.setRange(swiftcam.SWIFTCAM_TEMP_MIN, swiftcam.SWIFTCAM_TEMP_MAX)
        self.slider_temp.setValue(swiftcam.SWIFTCAM_TEMP_DEF)
        self.slider_tint.setRange(swiftcam.SWIFTCAM_TINT_MIN, swiftcam.SWIFTCAM_TINT_MAX)
        self.slider_tint.setValue(swiftcam.SWIFTCAM_TINT_DEF)
        self.slider_temp.setEnabled(False)
        self.slider_tint.setEnabled(False)
        vlyt_wb = QVBoxLayout()
        vlyt_wb.addLayout(self.makeLayout(lbl_temp, self.slider_temp, self.lbl_temp, lbl_tint, self.slider_tint, self.lbl_tint))
        vlyt_wb.addWidget(self.btn_autoWB)
        gbox_wb.setLayout(vlyt_wb)
        self.slider_temp.valueChanged.connect(self.onWBTemp)
        self.slider_tint.valueChanged.connect(self.onWBTint)

        self.btn_open = QPushButton("Open")
        self.btn_open.clicked.connect(self.onBtnOpen)
        
        global btn_snap
        btn_snap = QPushButton("Snap")
        btn_snap.setEnabled(False)
        btn_snap.clicked.connect(self.onBtnSnap)

        vlyt_ctrl = QVBoxLayout()
        vlyt_ctrl.addWidget(gbox_res)
        vlyt_ctrl.addWidget(gbox_exp)
        vlyt_ctrl.addWidget(gbox_wb)
        vlyt_ctrl.addWidget(self.btn_open)
        vlyt_ctrl.addWidget(btn_snap)
        vlyt_ctrl.addStretch()
        wg_ctrl = QWidget()
        wg_ctrl.setLayout(vlyt_ctrl)

        self.lbl_frame = QLabel()
        self.lbl_video = QLabel()
        vlyt_show = QVBoxLayout()
        vlyt_show.addWidget(self.lbl_video, 1)
        vlyt_show.addWidget(self.lbl_frame)
        wg_show = QWidget()
        wg_show.setLayout(vlyt_show)

        grid_main = QGridLayout()
        grid_main.setColumnStretch(0, 1)
        grid_main.setColumnStretch(1, 4)
        grid_main.addWidget(wg_ctrl)
        grid_main.addWidget(wg_show)
        w_main = QWidget()
        w_main.setLayout(grid_main)
        self.setCentralWidget(w_main)

        self.timer.timeout.connect(self.onTimer)
        self.evtCallback.connect(self.onevtCallback)

    def onTimer(self):
        if self.hcam:
            nFrame, nTime, nTotalFrame = self.hcam.get_FrameRate()
            self.lbl_frame.setText("{}, fps = {:.1f}".format(nTotalFrame, nFrame * 1000.0 / nTime))

    def closeCamera(self):
        if self.hcam:
            self.hcam.Close()
        self.hcam = None
        self.pData = None

        self.btn_open.setText("Open")
        self.timer.stop()
        self.lbl_frame.clear()
        self.cbox_auto.setEnabled(False)
        self.slider_expoGain.setEnabled(False)
        self.slider_expoTime.setEnabled(False)
        self.btn_autoWB.setEnabled(False)
        self.slider_temp.setEnabled(False)
        self.slider_tint.setEnabled(False)
        btn_snap.setEnabled(False)
        self.cmb_res.setEnabled(False)
        self.cmb_res.clear()

    def closeEvent(self, event):
        self.closeCamera()

    def onResolutionChanged(self, index):
        if self.hcam: #step 1: stop camera
            self.hcam.Stop()

        self.res = index
        self.imgWidth = self.cur.model.res[index].width
        self.imgHeight = self.cur.model.res[index].height

        if self.hcam: #step 2: restart camera
            self.hcam.put_eSize(self.res)
            self.startCamera()

    def onAutoExpo(self, state):
        if self.hcam:
            self.hcam.put_AutoExpoEnable(1 if state else 0)
            self.slider_expoTime.setEnabled(not state)
            self.slider_expoGain.setEnabled(not state)

    def onExpoTime(self, value):
        if self.hcam:
            self.lbl_expoTime.setText(str(value))
            if not self.cbox_auto.isChecked():
                self.hcam.put_ExpoTime(value)

    def onExpoGain(self, value):
        if self.hcam:
            self.lbl_expoGain.setText(str(value))
            if not self.cbox_auto.isChecked():
                self.hcam.put_ExpoAGain(value)

    def onAutoWB(self):
        if self.hcam:
            self.hcam.AwbOnce()

    def wbCallback(nTemp, nTint, self):
        self.slider_temp.setValue(nTemp)
        self.slider_tint.setValue(nTint)

    def onWBTemp(self, value):
        if self.hcam:
            self.temp = value
            self.hcam.put_TempTint(self.temp, self.tint)
            self.lbl_temp.setText(str(value))

    def onWBTint(self, value):
        if self.hcam:
            self.tint = value
            self.hcam.put_TempTint(self.temp, self.tint)
            self.lbl_tint.setText(str(value))

    def startCamera(self):
        self.pData = bytes(swiftcam.TDIBWIDTHBYTES(self.imgWidth * 24) * self.imgHeight)
        uimin, uimax, uidef = self.hcam.get_ExpTimeRange()
        self.slider_expoTime.setRange(uimin, uimax)
        self.slider_expoTime.setValue(uidef)
        usmin, usmax, usdef = self.hcam.get_ExpoAGainRange()
        self.slider_expoGain.setRange(usmin, usmax)
        self.slider_expoGain.setValue(usdef)
        self.handleExpoEvent()
        if self.cur.model.flag & swiftcam.SWIFTCAM_FLAG_MONO == 0:
            self.handleTempTintEvent()
        try:
            self.hcam.StartPullModeWithCallback(self.eventCallBack, self)
        except swiftcam.HRESULTException:
            self.closeCamera()
            QMessageBox.warning(self, "Warning", "Failed to start camera.")
        else:
            self.cmb_res.setEnabled(True)
            self.cbox_auto.setEnabled(True)
            self.btn_autoWB.setEnabled(True)
            self.slider_temp.setEnabled(self.cur.model.flag & swiftcam.SWIFTCAM_FLAG_MONO == 0)
            self.slider_tint.setEnabled(self.cur.model.flag & swiftcam.SWIFTCAM_FLAG_MONO == 0)
            self.btn_open.setText("Close")
##            btn_snap.setEnabled(True)
            bAuto = self.hcam.get_AutoExpoEnable()
            self.cbox_auto.setChecked(1 == bAuto)
            self.timer.start(1000)
            

    def openCamera(self):

        # LIMS Login (Qt)
        self.loginwindow = LoginWindow()
        self.loginwindow.show()


        # Open Camera
        self.hcam = swiftcam.Swiftcam.Open(self.cur.id)
        if self.hcam:
            self.res = self.hcam.get_eSize()
            self.imgWidth = self.cur.model.res[self.res].width
            self.imgHeight = self.cur.model.res[self.res].height
            with QSignalBlocker(self.cmb_res):
                self.cmb_res.clear()
                for i in range(0, self.cur.model.preview):
                    self.cmb_res.addItem("{}*{}".format(self.cur.model.res[i].width, self.cur.model.res[i].height))
                self.cmb_res.setCurrentIndex(self.res)
                self.cmb_res.setEnabled(True)                     
            self.hcam.put_Option(swiftcam.SWIFTCAM_OPTION_BYTEORDER, 0) #Qimage use RGB byte order
            self.hcam.put_AutoExpoEnable(1)
            self.startCamera()

            

    def onBtnOpen(self):
        if self.hcam:
            self.closeCamera()
        else:
            arr = swiftcam.Swiftcam.EnumV2()
            if 0 == len(arr):
                QMessageBox.warning(self, "Warning", "No camera found.")
            elif 1 == len(arr):
                self.cur = arr[0]
                self.openCamera()
            else:
                menu = QMenu()
                for i in range(0, len(arr)):
                    action = QAction(arr[i].displayname, self)
                    action.setData(i)
                    menu.addAction(action)
                action = menu.exec(self.mapToGlobal(self.btn_open.pos()))
                if action:
                    self.cur = arr[action.data()]
                    self.openCamera()


    def onBtnSnap(self):
        if self.hcam:
            if 0 == self.cur.model.still:    # not support still image capture
                if self.pData is not None:
                    image = QImage(self.pData, self.imgWidth, self.imgHeight, QImage.Format_RGB888)
                    self.count += 1
                    image.save("pyqt{}.jpg".format(self.count))
            else:
                menu = QMenu()
                for i in range(0, self.cur.model.still):
                    action = QAction("{}*{}".format(self.cur.model.res[i].width, self.cur.model.res[i].height), self)
                    action.setData(i)
                    menu.addAction(action)
                action = menu.exec(self.mapToGlobal(btn_snap.pos()))
                self.hcam.Snap(action.data())
                

    @staticmethod
    def eventCallBack(nEvent, self):
        '''callbacks come from swiftcam.dll/so internal threads, so we use qt signal to post this event to the UI thread'''
        self.evtCallback.emit(nEvent)

    def onevtCallback(self, nEvent):
        '''this run in the UI thread'''
        if self.hcam:
            if swiftcam.SWIFTCAM_EVENT_IMAGE == nEvent:
                self.handleImageEvent()
            elif swiftcam.SWIFTCAM_EVENT_EXPOSURE == nEvent:
                self.handleExpoEvent()
            elif swiftcam.SWIFTCAM_EVENT_TEMPTINT == nEvent:
                self.handleTempTintEvent()
            elif swiftcam.SWIFTCAM_EVENT_STILLIMAGE == nEvent:
                self.handleStillImageEvent()
            elif swiftcam.SWIFTCAM_EVENT_ERROR == nEvent:
                self.closeCamera()
                QMessageBox.warning(self, "Warning", "Generic Error.")
            elif swiftcam.SWIFTCAM_EVENT_STILLIMAGE == nEvent:
                self.closeCamera()
                QMessageBox.warning(self, "Warning", "Camera disconnect.")

    def handleImageEvent(self):
        try:
            self.hcam.PullImageV3(self.pData, 0, 24, 0, None)
        except swiftcam.HRESULTException:
            pass
        else:
            image = QImage(self.pData, self.imgWidth, self.imgHeight, QImage.Format_RGB888)
            newimage = image.scaled(self.lbl_video.width(), self.lbl_video.height(), Qt.KeepAspectRatio, Qt.FastTransformation)
            self.lbl_video.setPixmap(QPixmap.fromImage(newimage))

    def handleExpoEvent(self):
        time = self.hcam.get_ExpoTime()
        gain = self.hcam.get_ExpoAGain()
        with QSignalBlocker(self.slider_expoTime):
            self.slider_expoTime.setValue(time)
        with QSignalBlocker(self.slider_expoGain):
            self.slider_expoGain.setValue(gain)
        self.lbl_expoTime.setText(str(time))
        self.lbl_expoGain.setText(str(gain))

    def handleTempTintEvent(self):
        nTemp, nTint = self.hcam.get_TempTint()
        with QSignalBlocker(self.slider_temp):
            self.slider_temp.setValue(nTemp)
        with QSignalBlocker(self.slider_tint):
            self.slider_tint.setValue(nTint)
        self.lbl_temp.setText(str(nTemp))
        self.lbl_tint.setText(str(nTint))

    def handleStillImageEvent(self):
        info = swiftcam.SwiftcamFrameInfoV3()
        try:
            self.hcam.PullImageV3(None, 1, 24, 0, info) # peek
        except swiftcam.HRESULTException:
            pass
        else:
            if info.width > 0 and info.height > 0:
                buf = bytes(swiftcam.TDIBWIDTHBYTES(info.width * 24) * info.height)
                try:
                    self.hcam.PullImageV3(buf, 1, 24, 0, info)
                except swiftcam.HRESULTException:
                    pass
                else:
                    now = datetime.now()
                    image = QImage(buf, info.width, info.height, QImage.Format_RGB888)
                    self.count += 1
                    image.save(microscope_img_path + '\\' + projid + '\\' + os.environ['USERNAME'] + '_' + str(now).replace(':',';') + '_capture{}.jpg'.format(self.count))

                    # Update Audit Trail Table
                    lims_db_cursor.execute("UPDATE avancelims.microscope_cam_audit_trail SET Path_to_IMG = '" + mysql_ver_microscope_img_path + '\\\\' + projid + '\\\\' + os.environ['USERNAME'] + '_' + str(now).replace(':',';') + '_capture{}.jpg'.format(self.count) + "' WHERE Time_Stamp = '" + str(runtime) + "'") # Database
                    lims_db.commit()

                    print('\nImg Saved to ' + microscope_img_path + '\\' + projid)




                    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())
##    app.exec_()


