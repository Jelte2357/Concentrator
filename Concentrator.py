import pygetwindow as gw
from win11toast import toast
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QCheckBox, QPushButton, QLineEdit, QLabel, QFrame
from PyQt5.QtGui import QIcon, QKeySequence
from PyQt5.QtCore import QTimer 
import ctypes
from ctypes import windll, create_string_buffer 
import random
import string
import os
import threading

def generateRandomString(length=30): # for the quit text and alter text (remove "iIl| " to prevent confusion)
    return ''.join(random.choice(''.join(set(string.ascii_letters + string.digits + string.punctuation) - set("iIl| "))) for _ in range(length))

def notify(*args, **kwargs): # notify in a thread, as it would otherwise block the program
    threading.Thread(target=toast, args=args, kwargs=kwargs).start()
    
#-----Setup Vars----------
AllowedApps = [] # All allowed apps
AllowedAppParts = [] # All allowed app parts
BlockByParts = True # if false, block by full app
ChangeApps = True # if false, grey out checkboxes
quittext = generateRandomString()
altertext = generateRandomString()
selected = {"apps": [], "parts": []} # the selected apps and parts, needed for switching between viewing parts and apps
EdgeCaseApps = ['', 'realtek audio console', 'task manager', 'concentrator', 'translate this page?', 'mainwindow', 'settings', 'windows input experience', 'program manager', 'popuphost']
# the edge case apps are apps that are open sometimes that the program should not block (like concentrator (itself)), or that can not be closed using the program (like task manager)
#-------------------------

# originally I used win32gui, but that would not work with py2exe, so I found this somewhere (forgot to note down where)
def get_class_name(hwnd):
    buf_size = 256
    buffer = create_string_buffer(buf_size)
    windll.user32.GetClassNameA(hwnd, buffer, buf_size)
    return buffer.value.decode("utf-8")

def getAllWindows():
    global EdgeCaseApps, BlockByParts 
    windows = gw.getAllWindows()
    WindowsFiltered = []

    for window in windows:
        windowT = window.title.replace('\u200B', '') # microsoft products really like to add this 0 width character
        
        # remove all the weird edge cases
        if windowT.lower() not in EdgeCaseApps and not window.title.endswith('.ini'):
            if not BlockByParts:
                if windowT.startswith("command prompt"): # command prompt is a special case, as has command prompt at the start, not the end
                    WindowsFiltered.append("Command Prompt")
                
                # from https://stackoverflow.com/questions/72108062/check-if-windows-file-explorer-is-already-opened-in-python
                elif "CabinetWClass" in get_class_name(window._hWnd):
                    WindowsFiltered.append("File Explorer")
                    
                else: # remove even more edge cases
                    WindowsFiltered.append(windowT.replace("—", "-").replace("–", "-").split("-")[-1].strip())
                    
            else:
                
                if "CabinetWClass" in get_class_name(window._hWnd):
                    WindowsFiltered.append(windowT + " - File Explorer")
                else:
                    WindowsFiltered.append(window.title.replace("—", "-").replace("–", "-"))        
    
    return WindowsFiltered

def closeWindows():
    global AllowedApps, AllowedAppParts, EdgeCaseApps, ChangeApps
    if ChangeApps: # if the user is currently changing the allowed apps, dont close any apps
        return

    for window in gw.getAllWindows():
        if window.title.lower() not in EdgeCaseApps and not window.title.endswith('.ini'):
            # parts
            if ("CabinetWClass" in get_class_name(window._hWnd)) and (window.title + " - File Explorer" in AllowedAppParts):
                continue
                
            elif window.title.replace("—", "-").replace("–", "-") in AllowedAppParts:
                continue
            
            # apps
            if ("CabinetWClass" in get_class_name(window._hWnd)) and ("File Explorer" in AllowedApps):
                continue
                
            elif window.title.replace("—", "-").replace("–", "-").split("-")[-1].strip() in AllowedApps:
                continue
            
            # close the window
            window.close()
            notify("Concentrate", f'You are not allowed to open the app "{window.title}"')

# copy of QLineEdit, not allowing copy and paste, edited to also disable the right click menu
#found on https://stackoverflow.com/questions/65834966/how-to-block-copy-and-paste-key-in-pyqt5-python
class NOCVLineEdit(QLineEdit):
    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy) or event.matches(QKeySequence.Paste):
            return
        super().keyPressEvent(event)
    # set the context menu to be disabled, so you cant paste using the right click menu
    def contextMenuEvent(self, event):
        pass
        
class Concentrator(QWidget):
    global AllowedApps
        
    def __init__(self):
        super().__init__() # boilerplate code
        self.init_ui()
        self.checkboxes = [] # the checkboxes
        self.allowedapps = [] # the allowed apps
        self.timer = QTimer() # set the timer to run BGRunner every 250ms
        self.timer.timeout.connect(self.BGRunner)
        self.timer.start(250)

    def init_ui(self):

        self.setWindowTitle('Concentrator')
        self.setMinimumWidth(500) # ensure the name of the window is visible
        self.setMaximumWidth(2000) # larger than this and it looks weird
        try: # this code is needed to get the icon to work when compiled with py2exe (forgot to note down where I found this)
            IconPath = os.path.join(sys._MEIPASS, r"Blocker.ico")
        except:
            IconPath = os.path.abspath(r"./Blocker.ico")
        self.setWindowIcon(QIcon(IconPath))
        # this ctypes line is needed to set the taskbar icon in windows (7-11), no idea why this works
        # found on https://stackoverflow.com/questions/1551605/how-to-set-applications-taskbar-icon-in-windows-7/1552105#1552105
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('Concentrator')

        layout = QVBoxLayout()
        
        # these mostly are just labels and buttons, so I wont explain them
        
        self.textbox = QLabel("Note you can't quit this program using conventional methods, like alt+f4 or the x button. This is to make sure you concentrate.", wordWrap=True)
        layout.addWidget(self.textbox)
        
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        layout.addWidget(line)
        
        self.AlterTbx = QLabel("Type this text to alter the allowed apps: " + 
                              "-"
                              , wordWrap=True)
        layout.addWidget(self.AlterTbx)
        
        # no copy and paste textbox, see the definition of the class
        self.AlterInp = NOCVLineEdit(self)
        self.AlterInp.setPlaceholderText("Type here to confirm")
        self.AlterInp.setEnabled(False)
        layout.addWidget(self.AlterInp)
        
        self.AlterBtn = QPushButton('Alter apps', self)
        self.AlterBtn.clicked.connect(self.enableOptions)
        self.AlterBtn.setEnabled(False)
        layout.addWidget(self.AlterBtn)
        
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        layout.addWidget(line)
        
        # in here the checkboxes are added
        
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        layout.addWidget(line)
        
        self.ViewBtn = QPushButton('View Apps / App Parts', self)
        self.ViewBtn.clicked.connect(self.changeOA)
        layout.addWidget(self.ViewBtn)
        
        self.UpdateBtn = QPushButton('Update Allowed Apps', self)
        self.UpdateBtn.clicked.connect(self.updateAllowedApps)
        layout.addWidget(self.UpdateBtn)
        
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        layout.addWidget(line)
        
        self.AApps = QLabel("Allowed apps: \n", wordWrap=True)
        layout.addWidget(self.AApps)
        
        self.AParts = QLabel("Allowed app parts: \n", wordWrap=True)
        layout.addWidget(self.AParts)
        
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        layout.addWidget(line)
        
        self.textbox = QLabel("Type this text to quit: " + 
                              quittext
                              , wordWrap=True)
        layout.addWidget(self.textbox)
        
        #make a textbox here, in which you cant paste, but can type
        #see how I did this at the definition of the function
        self.textbox = NOCVLineEdit(self)
        self.textbox.setPlaceholderText("Type here to confirm")
        layout.addWidget(self.textbox)

        # quit button
        self.btn = QPushButton('Quit', self)
        self.btn.clicked.connect(self.quitter)
        layout.addWidget(self.btn)
        
        # start
        self.setLayout(layout)
        self.show()

    def get_checked_boxes(self): # get the checked and unchecked checkboxes
        checked_boxes = [checkbox.text() for checkbox in self.checkboxes if checkbox.isChecked()]
        unchecked_boxes = [checkbox.text() for checkbox in self.checkboxes if not checkbox.isChecked()]
        return checked_boxes, unchecked_boxes

    def quitter(self): # actually quit
        if self.textbox.text() != quittext:
            notify("Concentrate", "Your text is incorrect, you may not quit.")
        else:
            app.quit()
    
    def changeAlterText(self): # no real need to explain this one
        global altertext
        altertext = generateRandomString()
        self.AlterTbx.setText("Type this text to alter the allowed apps: " +
                                altertext)
    
    def enableOptions(self):
        # set the buttons to be enabled
        global ChangeApps
        # check the text in the textbox with altertext
        if self.AlterInp.text() != altertext:
            notify("Concentrate", "Your text is incorrect, you may not alter the allowed apps.")
            return
        ChangeApps = True
        # set the text to "Type this text to alter the allowed apps: -"
        self.AlterTbx.setText("Type this text to alter the allowed apps: -")
        # disable the textbox and button
        self.AlterInp.setEnabled(False)
        self.AlterInp.setText("")
        self.AlterBtn.setEnabled(False)
        self.updateCheckboxes(force=True)
        # enable the update allowed apps button
        self.UpdateBtn.setEnabled(True)
        self.ViewBtn.setEnabled(True)
        
    def updateAllowedApps(self):
        global AllowedApps, AllowedAppParts, ChangeApps
        # get the selected apps and app parts
        AllowedAppParts = selected["parts"]
        AllowedApps = selected["apps"]
        # update the checkboxes and set a new alter text
        self.updateCheckboxes(force=True)
        self.changeAlterText()
        # reset a few things
        self.AlterInp.setEnabled(True)
        self.AlterBtn.setEnabled(True)
        self.AlterInp.setText("")
        self.AlterInp.setFocus()
        # set the Update allowed apps button to be disabled
        self.UpdateBtn.setEnabled(False)
        self.ViewBtn.setEnabled(False)
        # set the buttons to disabled
        ChangeApps = False
        self.updateCheckboxes(force=True)
        
        # display the allowed apps and app parts
        self.AApps.setText(r"Allowed apps: " + (("\n - "+"\n - ".join(AllowedApps).removesuffix(" \n - ")) if AllowedApps else "\n"))
        self.AParts.setText(r"Allowed app parts: " + (("\n - "+"\n - ".join(AllowedAppParts).removesuffix("\n - ")) if AllowedAppParts else "\n"))
        
    def closeEvent(self, event): # this is run when the user tries to close the window (by pressing the x button or alt+f4)
        notify("Concentrate", "You can only close the concentration application by inputting the confirm text.")
        event.ignore()
        
    def changeOA(self): # change the view between apps and app parts and update the checkboxes
        global BlockByParts
        BlockByParts = not BlockByParts
        self.updateCheckboxes(force=True)

    def updateCheckboxes(self, force=False): # update the checkboxes
        self.windows =  getAllWindows()
        if (self.windows == [checkbox.text() for checkbox in self.checkboxes]) and not force: # if it is the same as last time, dont update, but if it is forced, update
            return
        if BlockByParts: # show the parts
            for checkbox in self.checkboxes:
                self.layout().removeWidget(checkbox)
            self.checkboxes = []
            for window in self.windows:
                self.checkboxes.append(QCheckBox(window))
                if window in AllowedAppParts:
                    self.checkboxes[-1].setChecked(True)
            for checkbox in self.checkboxes:
                # insert at above the location of update blocked apps
                self.layout().insertWidget(6, checkbox)
                checkbox.setEnabled(ChangeApps)
                if checkbox.text() in selected["parts"]:
                    checkbox.setChecked(True)
                
            self.adjustSize()
        else: # show the apps
            for checkbox in self.checkboxes:
                self.layout().removeWidget(checkbox)
            self.checkboxes = []
            for window in self.windows:
                self.checkboxes.append(QCheckBox(window))
                if window in AllowedApps:
                    self.checkboxes[-1].setChecked(True)
            for checkbox in self.checkboxes:
                # insert at above the location of update blocked apps
                self.layout().insertWidget(6, checkbox)
                checkbox.setEnabled(ChangeApps)
                if checkbox.text() in selected["apps"]:
                    checkbox.setChecked(True)
                
            self.adjustSize()
    
    def BGRunner(self): # this is run every 250ms (kind of like a while loop, but it doesnt block the program)
        global selected
        self.updateCheckboxes() # update the checkboxes
        self.currentCB, _ = self.get_checked_boxes()
        if BlockByParts: # set selected
            selected["parts"] = self.currentCB
        else:
            selected["apps"] = self.currentCB
        closeWindows() # close the windows that are not allowed to be open

# standard code to run the app
def run_app():
    global app, window
    app = QApplication(sys.argv)
    window = Concentrator()

    sys.exit(app.exec_())
    
# start the program
if __name__ == '__main__':
    run_app()
