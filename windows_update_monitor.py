# C:\Windows\SoftwareDistribution\Download

import os
import sys
import time
import subprocess
from pathlib import Path
import PyQt5.QtCore
from PyQt5 import QtWidgets, QtCore, QtGui

tray_icon = []


class SystemTrayIcon(QtWidgets.QSystemTrayIcon):

    def __init__(self, icon, parent=None):

        QtWidgets.QSystemTrayIcon.__init__(self, icon, parent)

        self.context_menu_style = """QMenu 
                                   {background-color: rgb(33, 33, 33);
                                   color: rgb(255, 255, 255);
                                   border-top:2px solid rgb(0, 0, 0);
                                   border-bottom:2px solid rgb(0, 0, 0);
                                   border-right:2px solid rgb(0, 0, 0);
                                   border-left:2px solid rgb(0, 0, 0);
                                   }
                                   QMenu::item::selected
                                   {background-color : rgb(33, 33, 33);
                                   color: rgb(33, 33, 255);
                                   }
                                   """

        # Initiate QMenu and set context menu style
        menu = QtWidgets.QMenu(parent)
        menu.setStyleSheet(self.context_menu_style)

        # Initiate context menu items
        menu.addAction(QtGui.QIcon("./img_standard.ico"), "Windows Update Monitor")
        menu.addSeparator()
        update_windows = menu.addAction(QtGui.QIcon("./img_downloading.ico"), "Update Windows")
        menu.addSeparator()
        pause_updates_0 = menu.addAction(QtGui.QIcon("./img_paused.ico"), "Pause Windows Updates (1 Month)")
        pause_updates_1 = menu.addAction(QtGui.QIcon("./img_paused.ico"), "Pause Windows Updates (3 Months)")
        pause_updates_2 = menu.addAction(QtGui.QIcon("./img_paused.ico"), "Pause Windows Updates (6 Months)")
        pause_updates_3 = menu.addAction(QtGui.QIcon("./img_paused.ico"), "Pause Windows Updates (1 Year)")
        pause_updates_4 = menu.addAction(QtGui.QIcon("./img_paused.ico"), "Pause Windows Updates (100 Years)")
        pause_updates_5 = menu.addAction(QtGui.QIcon("./img_paused.ico"), "Pause Windows Updates (Max)")
        menu.addSeparator()
        resume_updates = menu.addAction(QtGui.QIcon("./img_downloading.ico"), "Resume Windows Updates")
        menu.addSeparator()
        exit_action = menu.addAction(QtGui.QIcon("./img_exit.ico"), "Exit")

        # Set context menu
        self.setContextMenu(menu)

        # Plug in all the things
        exit_action.triggered.connect(self.exit)

        pause_updates_0.triggered.connect(self.pause_updates_function_pause_updates_0)
        pause_updates_1.triggered.connect(self.pause_updates_function_pause_updates_1)
        pause_updates_2.triggered.connect(self.pause_updates_function_pause_updates_2)
        pause_updates_3.triggered.connect(self.pause_updates_function_pause_updates_3)
        pause_updates_4.triggered.connect(self.pause_updates_function_pause_updates_4)
        pause_updates_5.triggered.connect(self.pause_updates_function_pause_updates_5)

        resume_updates.triggered.connect(self.resume_updates_function)
        update_windows.triggered.connect(self.update_function)

        # Indicator
        self.timer_gui = PyQt5.QtCore.QTimer(self)
        self.timer_gui.setInterval(3000)
        self.timer_gui.timeout.connect(self.timer_gui_function)
        self.timer_gui_start_function()

        # Updates Downloading
        self.timer_0 = PyQt5.QtCore.QTimer(self)
        self.timer_0.setInterval(3000)
        self.timer_0.timeout.connect(self.timer_0_function)
        self.timer_0_start_function()

        # Updates Paused
        self.timer_1 = PyQt5.QtCore.QTimer(self)
        self.timer_1.setInterval(4000)
        self.timer_1.timeout.connect(self.timer_1_function)
        self.timer_1_start_function()

        # Locals
        self.win_update_dir_sz_prev = 0
        self.downloading_updates = False
        self.updates_paused = False
        self.pause_script = ['powershell "./ps1_WindowsUpdates_Pause1Months.ps1"',
                             'powershell "./ps1_WindowsUpdates_Pause3Months.ps1"',
                             'powershell "./ps1_WindowsUpdates_Pause6Months.ps1"',
                             'powershell "./ps1_WindowsUpdates_Pause1Years.ps1"',
                             'powershell "./ps1_WindowsUpdates_Pause100Years.ps1"',
                             'powershell "./ps1_WindowsUpdates_PauseMaxExpiryTime.ps1"']

    @QtCore.pyqtSlot()
    def exit(self):
        QtCore.QCoreApplication.exit()

    @QtCore.pyqtSlot()
    def pause_updates_function_pause_updates_0(self):
        script = self.pause_script[0]
        self.pause_updates_function(script=script)

    @QtCore.pyqtSlot()
    def pause_updates_function_pause_updates_1(self):
        script = self.pause_script[1]
        self.pause_updates_function(script=script)

    @QtCore.pyqtSlot()
    def pause_updates_function_pause_updates_2(self):
        script = self.pause_script[2]
        self.pause_updates_function(script=script)

    @QtCore.pyqtSlot()
    def pause_updates_function_pause_updates_3(self):
        script = self.pause_script[3]
        self.pause_updates_function(script=script)

    @QtCore.pyqtSlot()
    def pause_updates_function_pause_updates_4(self):
        script = self.pause_script[4]
        self.pause_updates_function(script=script)

    @QtCore.pyqtSlot()
    def pause_updates_function_pause_updates_5(self):
        script = self.pause_script[5]
        self.pause_updates_function(script=script)

    @QtCore.pyqtSlot()
    def pause_updates_function(self, script):
        subprocess.Popen(script,
                         shell=True,
                         stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT)

    @QtCore.pyqtSlot()
    def resume_updates_function(self):
        subprocess.Popen('powershell "./ps1_WindowsUpdates_Resume.ps1"',
                         shell=True,
                         stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT)

    @QtCore.pyqtSlot()
    def update_function(self):
        xcmd = subprocess.Popen('powershell "./ps1_WindowsUpdates_Update.ps1"',
                                shell=True,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT)
        while True:
            output = xcmd.stdout.readline()
            if output == '' and xcmd.poll() is not None:
                break
            if output:
                pass
            else:
                break
            tray_icon.setIcon(QtGui.QIcon("./img_downloading.ico"))
            time.sleep(0.2)
            tray_icon.setIcon(QtGui.QIcon("./img_standard.ico"))
            time.sleep(0.2)

    @QtCore.pyqtSlot()
    def timer_gui_start_function(self):
        self.timer_gui.start()

    @QtCore.pyqtSlot()
    def timer_gui_stop_function(self):
        self.timer_gui.stop()

    @QtCore.pyqtSlot()
    def timer_0_start_function(self):
        self.timer_0.start()

    @QtCore.pyqtSlot()
    def timer_0_stop_function(self):
        self.timer_0.stop()

    @QtCore.pyqtSlot()
    def timer_1_start_function(self):
        self.timer_1.start()

    @QtCore.pyqtSlot()
    def timer_1_stop_function(self):
        self.timer_1.stop()

    @QtCore.pyqtSlot()
    def timer_gui_function(self):
        """ Set Indicators """

        # Standard
        if self.downloading_updates is False and self.updates_paused is False:
            tray_icon.setIcon(QtGui.QIcon("./img_standard.ico"))

        # Paused (Secondary)
        if self.updates_paused is True:
            tray_icon.setIcon(QtGui.QIcon("./img_paused.ico"))

        # Activity (Primary, overrides display paused indicator, displaying activity whether updates are paused or not)
        if self.downloading_updates is True:
            tray_icon.setIcon(QtGui.QIcon("./img_downloading.ico"))

    @QtCore.pyqtSlot()
    def timer_0_function(self):
        """ Monitor For When Windows Is Downloading Updates """

        if os.path.exists('C:\\Windows\\SoftwareDistribution'):
            win_update_dir_sz = sum(file.stat().st_size for file in Path('C:\\Windows\\SoftwareDistribution').rglob('*'))
            if win_update_dir_sz != self.win_update_dir_sz_prev:
                self.win_update_dir_sz_prev = win_update_dir_sz
                self.downloading_updates = True
            else:
                self.downloading_updates = False

    @QtCore.pyqtSlot()
    def timer_1_function(self):
        """ Monitors For When Windows Updates Are Paused """

        output_idx = 0
        xcmd = subprocess.Popen('powershell ./ps1_WindowsUpdates_PauseExpiryTimeGet.ps1',
                                shell=True,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.STDOUT)
        while True:
            output = xcmd.stdout.readline()
            if output == '' and xcmd.poll() is not None:
                break
            if output:
                output = output.decode("utf-8").strip()
                if output_idx == 3:
                    if len(output) >= 10:
                        self.updates_paused = True
                    else:
                        self.updates_paused = False
            else:
                break
            output_idx += 1


def main(_image):
    global tray_icon
    app = QtWidgets.QApplication(sys.argv)
    w = QtWidgets.QWidget()
    tray_icon = SystemTrayIcon(QtGui.QIcon(_image), w)
    tray_icon.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    image = './img_standard.ico'
    main(image)
