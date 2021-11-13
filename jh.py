import sys
import os
import datetime
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("jh.ui")[0]
## python실행파일 디렉토리
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
form_class = uic.loadUiType(BASE_DIR + r'/jh.ui')[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.vesselCode = "1JM"
        self.date = QDate.currentDate()
        self.cur_date = self.date.toString("yyyy-MM-dd")
        self.dateEdit.setDate(self.date)
        self.timerVar = QTimer()
        self.interval = 60000


        self.btn_update.clicked.connect(self.crawl)
        self.dateEdit.dateChanged.connect(self.dateChanged)
        self.btn_start.clicked.connect(self.setTimer)
        self.btn_stop.clicked.connect(self.stopTimer)
        self.cmb_interval.currentIndexChanged.connect(self.changeInterval)
        self.crawl()

    def changeInterval(self):
        text = self.cmb_interval.currentText()
        if text == "1분":
            self.interval = 1000 * 60
        elif text == "10분":
            self.interval = 1000 * 60 * 10
        elif text == "30분":
            self.interval = 1000 * 60 * 30
        elif text == "1시간":
            self.interval = 1000 * 60 * 60
        elif text == "10시간":
            self.interval = 1000 * 60 * 60 * 10


    def stopTimer(self):
        self.timerVar.stop()
        self.txt_status.setText("OFF")

    def setTimer(self):
        self.timerVar.setInterval(self.interval)
        self.timerVar.timeout.connect(self.crawl)
        self.timerVar.start()
        self.txt_status.setText("ON")

    def dateChanged(self):
        self.date = self.dateEdit.date()
        self.cur_date = self.date.toString("yyyy-MM-dd")



    def crawl(self):
        browser = self.txt_browser
        browser.clear()
        url = f"http://www.maersk.com/schedules/vesselSchedules?vesselCode={self.vesselCode}&fromDate={self.cur_date}"
        driver = webdriver.Chrome(ChromeDriverManager().install())
        driver.get(url)
        driver.implicitly_wait(10)
        btn_cookies = driver.find_element_by_class_name("coi-banner__accept").click()
        results = driver.find_elements_by_class_name("ptp-results__transport-plan--item")
        for result in results:
            port,terminal = result.find_element_by_class_name("location").find_elements_by_tag_name("div")
            arrival = result.find_element_by_class_name("transport-label")
            if len(port.text.rstrip()) != 0:
                text = port.text +"\n" + terminal.text + "\n" + arrival.text + "\n"
            else:
                text = arrival.text + "\n"
            browser.append(text)

        self.lbl_date.setText(f"Updated to {self.cur_date}")
        self.lbl_vesselName.setText("Vessel Name : MARIE MAERSK")
        self.txt_time.setText(datetime.datetime.now().strftime("%y/%m/%d %H.%M.%S"))
        driver.quit()

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()