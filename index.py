from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
import mysql.connector
from PyQt5.uic import loadUiType
import datetime
from xlrd import *
from xlsxwriter import *

ui,_ = loadUiType('library.ui')
login,_ = loadUiType('login.ui')

class Login(QWidget , login):
    def __init__(self):
        QWidget.__init__(self)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.Handel_Login)
        style = open('Themes/darkorange.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    
class MainApp(QMainWindow , ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Handel_Ui_Changes()
        self.Handel_Buttons()

        self.Show_Category()
        self.Show_Author()
        self.Show_Publisher()

        self.Show_Category_Combobox()
        self.Show_Author_Combobox()
        self.Show_Publisher_Combobox()

        self.Show_All_Client()
        self.Show_All_Book()

        self.Show_Handel_Day_Operation()

      def main():
          app = QApplication(sys.argv)
          window = Login()
          window.show()
          app.exec_()

      if __name__ == '__main__':
          main()
    
    def Handel_Ui_Changes(self):
        self.Hiding_Themes()
        self.tabWidget.tabBar().setVisible(False)


    def Handel_Buttons(self):
        self.pushButton_5.clicked.connect(self.Show_Themes)
        self.pushButton_8.clicked.connect(self.Hiding_Themes)

        self.pushButton.clicked.connect(self.Open_Day_To_Day_Tab)
        self.pushButton_2.clicked.connect(self.Open_Books_Tab)
        self.pushButton_25.clicked.connect(self.Open_Clients_Tab)
        self.pushButton_3.clicked.connect(self.Open_Users_Tab)
        self.pushButton_4.clicked.connect(self.Open_Settings_Tab)

        self.pushButton_6.clicked.connect(self.Handel_Day_Operations)

        self.pushButton_9.clicked.connect(self.Add_New_Book)
        self.pushButton_10.clicked.connect(self.Search_Books)
        self.pushButton_7.clicked.connect(self.Edit_Books)

        self.pushButton_15.clicked.connect(self.Add_Category)
        self.pushButton_16.clicked.connect(self.Add_Author)
        self.pushButton_17.clicked.connect(self.Add_Publisher)
        
        self.pushButton_33.clicked.connect(self.Delete_Day_Operation)
        self.pushButton_29.clicked.connect(self.Delete_Category)
        self.pushButton_30.clicked.connect(self.Delete_Author)
        self.pushButton_31.clicked.connect(self.Delete_Publisher)
        self.pushButton_11.clicked.connect(self.Delete_Books)
        self.pushButton_32.clicked.connect(self.Delete_User)
        self.pushButton_24.clicked.connect(self.Delete_Client)

        self.btn_add_user.clicked.connect(self.Add_New_User)
        self.pushButton_14.clicked.connect(self.Login)
        self.pushButton_13.clicked.connect(self.Edit_User)

        self.pushButton_22.clicked.connect(self.Dark_Orange_Theme)
        self.pushButton_23.clicked.connect(self.Dark_Gray_Theme)
        self.pushButton_21.clicked.connect(self.Dark_Blue_Theme)
        self.pushButton_20.clicked.connect(self.QDark_Theme)

        self.pushButton_12.clicked.connect(self.Add_New_Client)
        self.pushButton_19.clicked.connect(self.Search_Client)
        self.pushButton_18.clicked.connect(self.Edit_Client)

        self.pushButton_28.clicked.connect(self.Export_Day_Operations)
        self.pushButton_26.clicked.connect(self.Export_Books)
        self.pushButton_27.clicked.connect(self.Export_Clients)

        
    def Show_Themes(self):
        self.groupBox_3.show()


    def Hiding_Themes(self):
        self.groupBox_3.hide()


    #################################### OPENING TABS ################################
    

    def Open_Day_To_Day_Tab(self):
        self.tabWidget.setCurrentIndex(0)


    def Open_Books_Tab(self):
        self.tabWidget.setCurrentIndex(1)


    def Open_Clients_Tab(self):
        self.tabWidget.setCurrentIndex(2)


    def Open_Users_Tab(self):
        self.tabWidget.setCurrentIndex(3)


    def Open_Settings_Tab(self):
        self.tabWidget.setCurrentIndex(4)
    
    
   
    ############################################################
    ################# Show Settings Data In Ui #################


    def Show_Category_Combobox(self):
        mydb = mysql.connector.connect(host='remotemysql.com', user='kD9aDA144X', passwd='vdhr8AqoVB', database='kD9aDA144X')
        mycursor = mydb.cursor()

        queryString = 'Select category_name FROM category'
        mycursor.execute(queryString)
        
        data = mycursor.fetchall()

        self.comboBox_9.clear()
        self.comboBox_3.clear()
        for category in data :
            self.comboBox_9.addItem(category[0])
            self.comboBox_3.addItem(category[0])
        mydb.commit()
        mydb.close()


    def Show_Author_Combobox(self):
        mydb = mysql.connector.connect(host='remotemysql.com', user='kD9aDA144X', passwd='vdhr8AqoVB', database='kD9aDA144X')
        mycursor = mydb.cursor()

        queryString = 'Select outhor_name FROM outhor'
        mycursor.execute(queryString)
        
        data = mycursor.fetchall()

        self.comboBox_11.clear()
        self.comboBox_5.clear()
        for outhor in data :
           # print(outhor)
            self.comboBox_11.addItem(outhor[0])
            self.comboBox_5.addItem(outhor[0])
        mydb.commit()
        mydb.close()


    def Show_Publisher_Combobox(self):
        mydb = mysql.connector.connect(host='remotemysql.com', user='kD9aDA144X', passwd='vdhr8AqoVB', database='kD9aDA144X')
        mycursor = mydb.cursor()

        queryString = 'Select publisher_name FROM publisher'
        mycursor.execute(queryString)
        
        data = mycursor.fetchall()

        self.comboBox_10.clear()
        self.comboBox_4.clear()
        for publisher in data :
            self.comboBox_10.addItem(publisher[0])
            self.comboBox_4.addItem(publisher[0])
        mydb.commit()
        mydb.close()

################################################
   

