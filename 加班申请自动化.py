#pyinstaller -F -w
from selenium import webdriver
import json
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from time import time, sleep
from chinese_calendar import is_holiday,is_workday
import datetime
from PyQt6.QtCore import QDate,Qt,QThread
from PyQt6 import QtCore

# This Python file uses the following encoding: utf-8
import sys
from PyQt6.QtWidgets import QApplication,QMainWindow,QTableWidget,QTableWidgetItem,QAbstractItemView
from 加班申请 import Ui_Form

# 静态载入2
class mainwindow(Ui_Form,QMainWindow):
    def __init__(self, parent=None):
        super(mainwindow, self).__init__(parent)
        self.setupUi(self)
        #self.ui=Ui_Form()
        #self.ui.setupUi(self)
        self.dateEdit.setDate(QDate.currentDate())
        
        self.tableWidget.setHorizontalHeaderLabels(['工号','姓名','状态'])
        #gonghaos =json.load(open("加班列表.json", "rb"))
        with open("加班列表.json", 'r', encoding='utf-8') as fw:
            gonghaos = json.load(fw)
        x = 0
        for key, value in gonghaos.items():
            row_count = self.tableWidget.rowCount()  # 返回当前行数(尾部)
            self.tableWidget.insertRow(row_count)  # 尾部插入一行
            
            self.tableWidget.setItem(x,0,QTableWidgetItem(key))
            self.tableWidget.setItem(x,1,QTableWidgetItem(value.replace(value[-3:],'')))
            self.tableWidget.setItem(x,2,QTableWidgetItem(value[-3:].strip()))
            x = x + 1
        self.tableWidget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tableWidget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        # 点击事件获取所选内容、行列
        self.tableWidget.cellPressed.connect(self.getPosContent)
        self.pushButton_3.clicked.connect(self.clicked_button_3)
        self.pushButton_2.clicked.connect(self.clicked_button_2)
        self.pushButton_4.clicked.connect(self.clicked_button_4)
        self.pushButton.clicked.connect(self.clicked_button_1)
        self.show()
        
    # 获取选中行列、内容
    def getPosContent(self,row,col):
        gh = self.tableWidget.item(row,0).text()
        xm = self.tableWidget.item(row,1).text()
        zt = self.tableWidget.item(row,2).text()
        self.lineEdit.setText(gh)
        self.lineEdit_2.setText(xm)
        self.lineEdit_3.setText(zt)
        
    def clicked_button_3(self):
        
        r=self.tableWidget.selectedItems()[0].row()
        
        self.tableWidget.setItem(r,0,QTableWidgetItem(self.lineEdit.text()))
        self.tableWidget.setItem(r,1,QTableWidgetItem(self.lineEdit_2.text()))
        self.tableWidget.setItem(r,2,QTableWidgetItem(self.lineEdit_3.text()))
        self.save()

    def clicked_button_4(self):
        r=self.tableWidget.selectedItems()[0].row()
        self.tableWidget.removeRow(r)
        self.save()

    def clicked_button_1(self):
        h=self.lineEdit.text()
        m=self.lineEdit_2.text()
        t=self.lineEdit_3.text()
        if h !='' and m !='' and t!='':
            row_count = self.tableWidget.rowCount()  # 返回当前行数(尾部)
            self.tableWidget.insertRow(row_count)  # 尾部插入一行
            self.tableWidget.setItem(row_count,0,QTableWidgetItem(h))
            self.tableWidget.setItem(row_count,1,QTableWidgetItem(m))
            self.tableWidget.setItem(row_count,2,QTableWidgetItem(t))
            self.save()
        
    def save(self):
        row_count = self.tableWidget.rowCount()
        
        dic = {}
        for i in range(0,row_count):
            #print(self.tableWidget.item(i,0).text()+':'+ self.tableWidget.item(i,1).text()+' '+self.tableWidget.item(i,2).text())
            dict2 ={self.tableWidget.item(i,0).text(): self.tableWidget.item(i,1).text()+' '+self.tableWidget.item(i,2).text()}
            dic.update(dict2)
        #j = json.dumps(dic)
        with open("./加班列表.json", 'w', encoding='utf-8') as fw:
            json.dump(dic, fw, indent=4, ensure_ascii=False)
            print("加载入文件完成...")
    
    def clicked_button_2(self):
        riqi = self.dateEdit.date().toString('yyyy-MM-dd')
        # 新建对象，传入参数
        self.my_thread = Mythread()
        self.my_thread.setPath(riqi)
        self.my_thread.start()


        
# 打开网址
   
def is_weekends(t):
    is_weekend=t.weekday() + 1
    if is_weekend==6 or is_weekend==7:
        return True
    else:
        return False
        
def dayadd(t):   
    delta = datetime.timedelta(days=1)
    n_days = t + delta
    return(n_days)


class Mythread(QThread):
    #声明一个信号，同时返回一个int，什么都可以返回，参数是发送信号时附带参数的数据类型

    def __init__(self):
        super(Mythread, self).__init__()
        self.filepath = ''
    
     
    def setPath(self,path):
        self.filepath = path
        '''
    def is_weekends(t):
        is_weekend=t.weekday() + 1
        if is_weekend==6 or is_weekend==7:
            return True
        else:
            return False
        
    def dayadd(t):   
        delta = datetime.timedelta(days=1)
        n_days = t + delta
        return(n_days)
'''
    def run(self):
        riqi=self.filepath
        
        #riqi = input("请输入加班申请日期格式如2022-10-10：")
        date=datetime.datetime.strptime(riqi, "%Y-%m-%d")
        
        is_weekend = is_weekends(date)
        print(is_weekend)
        #wb = webdriver.Chrome()
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option('detach', True)
        chrome_options.add_experimental_option("excludeSwitches",["enable-logging"])
        chrome_options.add_argument('--ignore-certificate-errors')
        wb = webdriver.Chrome(options=chrome_options)
        
        wb.maximize_window()
        wb.get('http://218.4.20.46:8081/seeyon/main.do')
        # 隐式地等待
        wb.implicitly_wait(10)
        try:
            username = wb.find_element(By.ID,'login_username')
            password = wb.find_element(By.ID,'login_password')
            username.send_keys('CX009')
            password.send_keys('CX0009')
            wb.find_element(By.ID,'login_button').click()
        except Exception as e:
            print(e)

        wb.implicitly_wait(10)
        wb.find_element(By.XPATH,"//*[@id='magnet_621082647729357284']/div/div[2]").click()
        wb.implicitly_wait(10)
        wb.switch_to.frame('mainIframe')

        wb.find_element(By.XPATH,'//*[@id="tree_7_span"]').click()

        wb.find_element(By.XPATH,'//*[@id="templateDatasTab"]/tbody/tr[6]/td[1]/a').click()
        wb.switch_to.window(wb.window_handles[-1])

        wb.switch_to.frame(wb.find_element(By.XPATH,'//*[@id="zwIframe"]'))

        wb.find_element(By.XPATH,'//*[@id="field0001"]').send_keys(date.strftime('%Y-%m-%d'))
        with open("加班列表.json", 'r', encoding='utf-8') as fw:
            gonghaos = json.load(fw)
        for key, value in gonghaos.items():
    
    
            if "请假" in value:
                continue
            if is_weekend:
                if "白班" in value:
                    start=date.strftime('%Y-%m-%d')+' 08:00'
                    end=date.strftime('%Y-%m-%d')+' 20:00'
                    total='11.00'
                else:
                    start=date.strftime('%Y-%m-%d')+' 20:00'
                    end=dayadd(date).strftime('%Y-%m-%d')+' 08:00'
                    total='11.50'
            else:
                if "白班" in value:
                    start=date.strftime('%Y-%m-%d')+' 17:00'
                    end=date.strftime('%Y-%m-%d')+' 20:00'
                    total='3.00'
                else:
                    start=dayadd(date).strftime('%Y-%m-%d')+' 04:30'
                    end=dayadd(date).strftime('%Y-%m-%d')+' 08:00'
                    total='3.50'
            wb.implicitly_wait(10)
            wb.find_element(By.XPATH,'//*[@id="field0005"]').click()
            wb.find_element(By.XPATH,'//*[@id="field0005_span"]/span[2]').click()
            wb.switch_to.window(wb.window_handles[-1])
            iframe=wb.find_element(By.XPATH,'//*[contains(@id,"layui-layer-iframe")]')
            wb.switch_to.frame(iframe)
    
            # 定位收藏栏
            collect  = wb.find_element(By.XPATH,'//*[@class="common_search common_search_condition clearfix"]')

            # 悬停至收藏标签处
            ActionChains(wb).move_to_element(collect).perform()

            wb.find_element(By.XPATH,'//*[@class="common_drop_list_content common_drop_list_content_action"]/a[3]').click()
            #wb.switch_to.frame('layui-layer-iframe1')
            gh  = wb.find_element(By.XPATH,'//*[@id="field0001"]')
            gh.send_keys(key)
            gh.send_keys(Keys.ENTER)
            sleep(1)
            wb.find_element(By.XPATH,'//*[contains(@id,"row")]/td[1]/div/input').click()
    
            wb.switch_to.window(wb.window_handles[-1])
            wb.find_element(By.XPATH,'//*[@class="layui-layer-btn0 margin_r_10 common_button common_button_emphasize  "]').click()
    
            wb.switch_to.frame('zwIframe')
    
            sleep(0.5)
            t1=wb.find_element(By.XPATH,'//*[@id="field0008"]')
            t1.clear()
            t1.send_keys(start)
            t2=wb.find_element(By.XPATH,'//*[@id="field0009"]')
            t2.clear()
            t2.send_keys(end)
            t3=wb.find_element(By.XPATH,'//*[@id="field0010"]')
            t3.clear()
            t3.send_keys(total)
            t4=wb.find_element(By.XPATH,'//*[@id="field0011"]')
            t4.clear()
            t4.send_keys('生产需要')
            sleep(0.5)
            print(value+"加班申请成功！")
            wb.find_element(By.XPATH,'//*[@id="field0005"]').click()
    
            wb.find_element(By.XPATH,'//*[@id="addImg"]').click()

if __name__=="__main__":
    app=QApplication(sys.argv)
    window=mainwindow() 
    sys.exit(app.exec())
    

   



