import shutil
import tempfile

from PyQt6.QtWidgets import *
from PyQt6.QtCore import QThread,pyqtSignal
from PyQt6 import uic
import os
from time import sleep
import undetected_chromedriver.v2 as uc
import pandas as pd
from selenium import webdriver
_count = 0
_countSucess = 0
_countExcel = 0

class UI(QMainWindow):
    def __init__(self):
        super(UI, self).__init__()
        uic.loadUi("Data/FormGG.ui",self)
        self.pushButton = self.findChild(QPushButton,'pushButton')
        self.pushButton.setEnabled(False)
        self.pushButton.clicked.connect(self.start)
        self.pushButton_2 = self.findChild(QPushButton,'pushButton_2')
        self.pushButton_2.clicked.connect(self.diaLogExcelFile)
        self.label_2 = self.findChild(QLabel,'label_2')
        self.spinBox = self.findChild(QSpinBox,'spinBox')
        self.checkBox = self.findChild(QCheckBox,'checkBox')
        self.label_3 = self.findChild(QLabel,'label_3')
        self.threadHandel = {}
        self.show()

    def start(self):
        for i in range(0,int(self.spinBox.value())):
            self.threadHandel[i] = HandelThread(i)
            self.threadHandel[i].excelFiles = self.fileName
            self.threadHandel[i].labelSucess.connect(self.labelSucess)
            self.threadHandel[i].labelStatus.connect(self.labelStatus)
            self.threadHandel[i].checkBox = self.checkBox
            self.threadHandel[i].start()
    def diaLogExcelFile(self):
        try:
            self.fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()")
            self.pushButton.setEnabled(True)
        except:
            pass
    def labelSucess(self,index):
        self.label_2.setText(f'Thành Công: {index}')
    def labelStatus(self,text):
        self.label_3.setText(f'Trạng Thái: {text}')
class HandelThread(QThread):
    labelSucess = pyqtSignal(str)
    labelStatus = pyqtSignal(str)
    def __init__(self,index = 0):
        super(HandelThread,self).__init__()
        self.index = index
    def run(self):
        try:
            self.handel()
        except:
            self.labelStatus.emit('Lỗi,Vui lòng liên hệ dev để kiểm tra lỗi !')
    def setBrowser(self):
        self.temp = os.path.normpath(tempfile.mkdtemp())
        opts = uc.ChromeOptions()
        opts.add_argument(f"--window-position={self.index * 200},0")
        opts.add_argument("--window-size=800,880")
        self.labelStatus.emit('Khởi tạo trình duyệt !')
        args = ["hide_console", ]
        opts.add_argument(f'--user-data-dir=Profile {self.temp}')
        opts.add_argument("--disable-popup-blocking")
        opts.add_argument("--disable-gpu")
        if self.checkBox.isChecked():
            opts.headless = True
        self.browser = uc.Chrome(executable_path=os.getcwd()+'/chromedriver',options=opts)
        with self.browser:self.browser.get('https://gmail.com')
    def handel(self):
        # Create Browser
        self.setBrowser()
        # Open Email.txt And Get Info Email.If not Email in Email.txt.Show Debug
        try:
            with open('email.txt', 'r') as openEmail:
                openEmail = openEmail.readlines()[0]
            email = openEmail.split('|')[0]
            password = openEmail.split('|')[1]
        except:
            print('No Find Email !')
        try:
            # Send Email text
            self.browser.find_element_by_xpath('//*[@id="identifierId"]').send_keys(email)
            sleep(1)
            self.labelStatus.emit('Đang sử lý đăng nhập gmail !')
            # Try Except Click Next .
            try:
                self.browser.find_element_by_xpath('//*[@id="identifierNext"]/div/button').click()
            except:
                pass
            # Check Loading Page
            check = self.waitBrowser(self.browser, 'input[type="password"]')
            if check == 0:
                sleep(3)
                # If check continue handel
                # Send password text
                self.browser.find_element_by_css_selector('input[type="password"]').send_keys(password)
                sleep(1)
            # Click Continue
            try:
                self.browser.find_element_by_xpath('//*[@id="passwordNext"]/div/button').click()
            except:
                pass

            sleep(5)
        except:pass
        # Connect To Form Google
        while True:
            try:
                with self.browser:
                    self.browser.get(
                        'https://docs.google.com/forms/d/e/1FAIpQLSdnOuji6opZn7wSoErzkSfsbE21iu-hUnsM0DisaY7DniYRBQ/viewform')
                sleep(1)
                self.labelStatus.emit('Đang sử lý điền form !')
                # Clear Form
                self.browser.execute_script("""document.querySelectorAll('span[class="l4V7wb Fxmcue"]')[1].click()""")
                sleep(2)
                self.browser.execute_script("""document.querySelectorAll('div[data-id="EBS5u"]')[1].click()""")
                sleep(2)

                # Handel Form
                # Open Excel Files
                global _countExcel
                file_errors_location = self.excelFiles
                xlsxData = pd.read_excel(file_errors_location)
                loop = xlsxData.shape[0]
                if _countExcel >= int(loop):break
                dataExcel = xlsxData.iloc[_countExcel]
                _countExcel += 1
                # Send Key
                self.browser.find_element_by_css_selector("input.whsOnd.zHQkBf").send_keys(dataExcel[0])
                sleep(1)
                self.browser.find_elements_by_css_selector("input.whsOnd.zHQkBf")[1].send_keys(str(dataExcel[1]))
                sleep(1)
                self.browser.find_element_by_css_selector("textarea.KHxj8b.tL9Q4c").send_keys(str(dataExcel[2]))
                sleep(1)
                self.browser.find_elements_by_css_selector("input.whsOnd.zHQkBf")[2].send_keys(str(dataExcel[3]))
                sleep(1)
                self.browser.find_elements_by_css_selector("input.whsOnd.zHQkBf")[3].send_keys(str(dataExcel[4]))
                sleep(1)
                self.browser.find_elements_by_css_selector("input.whsOnd.zHQkBf")[4].send_keys(str(dataExcel[5]))
                sleep(1)
                self.browser.find_elements_by_css_selector("input.whsOnd.zHQkBf")[5].send_keys(str(dataExcel[6]))
                sleep(1)
                self.browser.find_elements_by_css_selector("input.whsOnd.zHQkBf")[6].send_keys(str(dataExcel[7]))
                sleep(1)
                self.browser.find_elements_by_css_selector("input.whsOnd.zHQkBf")[7].send_keys(str(dataExcel[8]))
                sleep(1)
                self.browser.find_elements_by_css_selector("textarea.KHxj8b.tL9Q4c")[1].send_keys(str(dataExcel[9]))
                sleep(1)
                self.browser.find_elements_by_css_selector("input.whsOnd.zHQkBf")[8].send_keys(str(dataExcel[10]))
                sleep(1)
                self.browser.find_elements_by_css_selector("textarea.KHxj8b.tL9Q4c")[2].send_keys(str(dataExcel[11]))
                sleep(1)
                #CLick Button
                self.browser.execute_script("""document.querySelector("span.l4V7wb.Fxmcue").click()""")
                global _countSucess
                _countSucess += 1
                self.labelSucess.emit(f'{_countSucess}')
                self.labelStatus.emit('Điền form thành công !')
                sleep(3)
            except:pass
        self.closeBrowser()

    def closeBrowser(self):
        if self.browser:
            self.browser.close()
            self.browser.quit()
            try:
                shutil.rmtree(r'{}'.format(self.temp))
            except:
                pass

    def waitBrowser(self,browser, options=None, options1=None, options2=None, options3=None, options4=None,
                    options5=None,
                    time_out=15, index=0, index1=0, index2=0, index3=0, index4=0, index5=0):
        check = None
        count = 0
        indexcheck = None
        while not check and count < time_out:
            try:
                check = browser.find_elements_by_css_selector(options)[index]
                indexcheck = 0
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options1)[index1]
                indexcheck = 1
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options2)[index2]
                indexcheck = 2
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options3)[index3]
                indexcheck = 3
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options4)[index4]
                indexcheck = 4
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options5)[index5]
                indexcheck = 5
            except:
                pass
            sleep(1)
            count += 1

        return indexcheck
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    UIwindow = UI()
    app.exec()


