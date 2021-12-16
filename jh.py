import sys
import os
import datetime
import openpyxl
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from selenium import webdriver
#from selenium.webdriver.support.relative_locator import locate_with
from selenium.webdriver.common.by import By
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
        self.url = f"http://www.maersk.com/schedules/vesselSchedules?vesselCode={self.vesselCode}&fromDate={self.cur_date}"
        self.driver = webdriver.Chrome(ChromeDriverManager().install())
        self.timerVar = QTimer()
        self.interval = 60000
        self.cnt = 0


        self.btn_update.clicked.connect(self.crawl)
        self.dateEdit.dateChanged.connect(self.dateChanged)
        self.btn_start.clicked.connect(self.setTimer)
        self.btn_stop.clicked.connect(self.stopTimer)
        self.cmb_interval.currentIndexChanged.connect(self.changeInterval)
        self.crawl()
        self.btn_refresh.clicked.connect(self.initTable)
        self.btn_updateValue.clicked.connect(self.updateValue)
        self.initTable()

    def updateValue(self):

        """Update Value Modified by User"""

        wb = openpyxl.load_workbook(filename="schedule.xlsx",read_only=False, data_only=True)
        ws = wb["Vessel schedule"]
        target = chr(int(self.c) + 65) + str(self.r + 2) # (0,0) -> A1, Z열까지만 가능
        newValue = self.valueEdit.text()
        ws[target] = newValue
        wb.save("schedule.xlsx")
        self.initTable()


        pass


    def resetUpdateLayout(self):

        """ Clear All Widgets in Update Layout """

        for i in reversed(range(self.updateLayout.count())):
            self.updateLayout.itemAt(i).widget().setParent(None)

    def showClickedLabel(self, r, c, t1 = True):

        """Show Clicked Label"""

        self.resetUpdateLayout()

        clickedColumnValue = self.df.iloc[r,c]

        layout = self.updateLayout

        e = QLineEdit()
        e.setText(str(clickedColumnValue))
        layout.addWidget(QLabel("Values: "))
        layout.addWidget(e)

        self.r, self.c, self.valueEdit = r, c, e



    def setLabel1(self, row, column):

        """Change Index Table1 to Original DataFrame"""

        r,c = row+8,column+1
        self.showClickedLabel(r,c,t1=True)

    def setLabel2(self, row, column):

        """Change Index Table2 to Original DataFrame"""

        r, c = row + 24, column + 1
        print(self.df.iloc[r, c])
        self.showClickedLabel(r, c, t1 = False)

    def initTable(self):

        """Initialize Table"""

        table1 = self.table1
        table2 = self.table2
        df = pd.read_excel("schedule.xlsx")
        df1 = df.iloc[8:15,1:]
        df2 = df.iloc[24:,1:8]
        df1.columns = ["Vessel Name(AE10)","Planned_ETA(B)","Current_ETA1(B)","Current_ETA2(B)","Delay(B)","Planned_ETA(G)","Current_ETA(G)","Delay(G)","Vessel Location","ETA Change"]
        df2.columns = ["Vessel Name(AE05)","Planned_ETA(B)","Current_ETA1(B)","Current_ETA2(B)","Delay(B)","Vessel Location","ETA Change"]
        self.df,self.df1,self.df2 = df,df1,df2

        table1.setColumnCount(len(df1.columns))
        table1.setRowCount(len(df1))
        for i in range(len(df1)):
            for j in range(len(df1.columns)):
                table1.setItem(i, j, QTableWidgetItem(str(df1.iloc[i, j])))
        table1.setHorizontalHeaderLabels(df1.columns)  # 컬럼 헤더 입력
        table1.resizeColumnsToContents()
        table1.cellClicked.connect(self.setLabel1)

        table2.setColumnCount(len(df2.columns))
        table2.setRowCount(len(df2))
        for i in range(len(df2)):
            for j in range(len(df2.columns)):
                table2.setItem(i, j, QTableWidgetItem(str(df2.iloc[i, j])))
        table2.setHorizontalHeaderLabels(df2.columns)  # 컬럼 헤더 입력
        table2.resizeColumnsToContents()
        table2.cellClicked.connect(self.setLabel2)


    def changeInterval(self):

        """Set Auto-Crawling Interval"""

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

        """Stop Timer"""

        self.timerVar.stop()
        self.txt_status.setText("OFF")

    def setTimer(self):

        """Set Timer"""

        self.timerVar.setInterval(self.interval)
        self.timerVar.timeout.connect(self.crawl)
        self.timerVar.start()
        self.txt_status.setText("ON")

    def dateChanged(self):

        """Set Changed Date"""

        self.date = self.dateEdit.date()
        self.cur_date = self.date.toString("yyyy-MM-dd")



    def crawl(self):

        """Start Crawling"""

        self.url = f"http://www.maersk.com/schedules/vesselSchedules?vesselCode={self.vesselCode}&fromDate={self.cur_date}"
        self.driver.get(self.url)
        self.driver.implicitly_wait(10)
        try:
            btn_cookies = self.driver.find_element(By.CLASS_NAME,"coi-banner__accept")
            btn_cookies.click()
        except:
            pass

        results = self.driver.find_elements(By.CLASS_NAME,"ptp-results__transport-plan--item")
        final_element = self.driver.find_elements(By.CLASS_NAME,"ptp-results__transport-plan--item-final")
        port_info = [] # port, arrival, department
        arr = True
        for i, result in enumerate(results + final_element): # port와 date 정보 추출
            port,_ = result.find_element(By.CLASS_NAME,"location").find_elements(By.TAG_NAME,"div")
            _,date = result.find_element(By.CLASS_NAME,"transport-label").find_elements(By.TAG_NAME,"div")
            p,d = port.text, date.text

            if (len(port.text)) == 0: # port 문자열이 비어있다면 date는 arrival 날짜 else department
                arr = False
            else:
                arr = True

            if arr:
                info = [p,d]
                port_info.append(info)
            else:
                port_info[-1].append(d)
        print(port_info)


        self.lbl_date.setText(f"Last Updated {self.cur_date}")
        self.lbl_vesselName.setText("Vessel Name : MARIE MAERSK")
        self.txt_time.setText(datetime.datetime.now().strftime("%y/%m/%d %H.%M.%S"))


if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()
    myWindow.driver.quit()