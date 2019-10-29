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
        self.connectionString = mysql.connector.connect(option_files='my.conf')

    def Handel_Login(self):
        cur = self.connectionString.cursor()

        username = self.lineEdit.text()
        password = self.lineEdit_2.text()
        #requires all usernames to be unique
        login_query = ('''SELECT * FROM User WHERE Username= ? AND Password= ?''',(username,password,))

        cur.execute(login_query)
        data = cur.fetchall()
        if str(data) == "[]":
            self.label.setText('Make Sure You Enterd Your User Name And Password Correctly.')
        else:
            print("User Match")
            self.close()
            self.window2.show()
                
        cur.close()



class MainApp(QMainWindow , ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.connectionString = mysql.connector.connect(option_files='my.conf')
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
    
    
    ################################################
    ################# Day Operations ###############


    def Handel_Day_Operations(self):
        book_title = self.lineEdit.text()
        client_name = self.lineEdit_5.text()
        types = self.comboBox.currentText()
        days = int(self.comboBox_2.currentText())
        date_days = datetime.date.today()
        to_days = date_days + datetime.timedelta(days=days)

        print(date_days)
        print(to_days)

        cur = self.connectionString.cursor()

        add_day_operation_query = ("INSERT INTO dayoperations "
                        "(book_name, client, type, days, date_days, to_days) "
                        "VALUES (%s, %s, %s, %s, %s, %s)")

        add_day_operation_query_data = (book_title, client_name, types, days, date_days, to_days)

        cur.execute(add_day_operation_query, add_day_operation_query_data)

        self.connectionString.commit()
        cur.close()
        self.statusBar().showMessage('Add New Opperation ' + str(add_day_operation_query_data))
        self.Show_Handel_Day_Operation()


    def Show_Handel_Day_Operation(self):
        cur = self.connectionString.cursor()

        show_handel_day_operation_query = ("SELECT book_name, client, type, date_days, to_days,  days FROM dayoperations")
        cur.execute(show_handel_day_operation_query)
        data = cur.fetchall()

        self.tableWidget.setRowCount(0)
        self.tableWidget.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate (form):
                self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_position = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_position)
        cur.close()

    def Delete_Day_Operation(self):
        cur = self.connectionString.cursor()

        operation_name = self.lineEdit_7.text()

        warning = QMessageBox.warning(self, 'Delete dayoperations', "Are You Sure You Want To Delete This dayoperation", QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes :
            delete_dayoperations_query = ("delete from dayoperations where client=%s")
            cur.execute(delete_dayoperations_query, [(operation_name)])
            self.connectionString.commit()
            cur.close()
            self.statusBar().showMessage('dayoperation Deleted')
            self.Show_Handel_Day_Operation()

    #########################################
    ################# Books #################


    def Show_All_Book(self):
        cur = self.connectionString.cursor()

        Show_All_Client_query = ("Select book_code, book_name, book_description, book_category, book_author, book_publisher, book_price From book ")
        cur.execute(Show_All_Client_query)
        data = cur.fetchall()

        self.tableWidget_5.setRowCount(0)
        self.tableWidget_5.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate (form):
                self.tableWidget_5.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_position = self.tableWidget_5.rowCount()
            self.tableWidget_5.insertRow(row_position)
        self.connectionString.commit()
        cur.close()


    def Add_New_Book(self):
        cur = self.connectionString.cursor()

        book_title = self.lineEdit_8.text()
        book_description = self.textEdit.toPlainText()
        book_code = self.lineEdit_10.text()
        book_category = self.comboBox_9.currentText()
        book_author = self.comboBox_11.currentText()
        book_publisher = self.comboBox_10.currentText()
        book_price = self.lineEdit_9.text()

        add_book_sql_query = ("INSERT INTO book "
                        "(book_name, book_description, book_code, book_category, book_author, book_publisher, book_price)"
                        "VALUES (%s, %s, %s, %s, %s, %s, %s)")

        add_book_sql_query_data = (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price)

        cur.execute(add_book_sql_query, add_book_sql_query_data)

        self.connectionString.commit()
        cur.close()
        self.statusBar().showMessage('New Book Addedd Successfuly'+ str(add_book_sql_query_data))
        
        self.lineEdit_8.setText('')
        self.textEdit.setPlainText('')
        self.lineEdit_10.setText('')
        self.comboBox_9.setCurrentIndex(0)
        self.comboBox_11.setCurrentIndex(0)
        self.comboBox_10.setCurrentIndex(0)
        self.lineEdit_9.setText('')

        self.Show_All_Book()


    def Search_Books(self):
        cur = self.connectionString.cursor()

        book_title = self.lineEdit_3.text()

        sql = 'SELECT * FROM book WHERE book_name = %s'
        cur.execute(sql , [(book_title)])
        data = cur.fetchone()

        if (data):
            self.book_id.setText(str(data[0]))
            self.lineEdit_11.setText(data[1])
            self.textEdit_2.setPlainText(data[2])
            self.lineEdit_2.setText(data[3])
            self.comboBox_3.setCurrentText(data[4])
            self.comboBox_5.setCurrentText(data[5])
            self.comboBox_4.setCurrentText(data[6])
            self.lineEdit_4.setText(data[7])
        else:
            self.statusBar().showMessage('No Record found')
        self.connectionString.commit()
        cur.close()


    def Edit_Books(self):
        cur = self.connectionString.cursor()

        book_id = self.book_id.text()
        book_title = self.lineEdit_11.text()
        book_description = self.textEdit_2.toPlainText()
        book_code = self.lineEdit_2.text()
        book_category = self.comboBox_3.currentText()
        book_author = self.comboBox_5.currentText()
        book_publisher = self.comboBox_4.currentText()
        book_price = self.lineEdit_4.text()


        Search_Book_title = self.lineEdit_3.text()
        
        edit_book_query = ("update book set book_name=%s, book_description=%s, book_code=%s, book_category=%s,"
                        " book_author=%s, book_publisher=%s, book_price=%s "
                        "Where id=%s")

        edit_book_query_data = (book_title, book_description, book_code, book_category, book_author, book_publisher, book_price, int(book_id))
        
        cur.execute(edit_book_query, edit_book_query_data)

        self.connectionString.commit()
        cur.close()

        self.statusBar().showMessage('Edit Book Successfuly'+ str(edit_book_query_data))
        self.Show_All_Book()


    def Delete_Books(self):
        cur = self.connectionString.cursor()

        book_title = self.lineEdit_3.text()

        warning = QMessageBox.warning(self, 'Delete Book', "Are You Sure You Want To Delete This Book", QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes :
            delete_book_query = ("delete from book where book_name=%s")
            cur.execute(delete_book_query, [(book_title)])
            self.connectionString.commit()
            self.statusBar().showMessage('Book Deleted')
            self.Show_All_Book()
        cur.close()


    #########################################
    ################# Clients ###############


    def Show_All_Client(self):
        cur = self.connectionString.cursor()

        Show_All_Client_query = ("Select client_name, client_email, client_nationalid From clients")
        cur.execute(Show_All_Client_query)
        data = cur.fetchall()

        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate (form):
                self.tableWidget_6.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)
        self.connectionString.commit()
        cur.close()

    def Add_New_Client(self):
        cur = self.connectionString.cursor()

        client_name = self.lineEdit_12.text()
        client_email = self.lineEdit_14.text()
        client_nationalid = self.lineEdit_13.text()

        Add_New_Client_query = ("Insert into clients "
                                "(client_name, client_email, client_nationalid)"
                                "VALUES (%s, %s, %s)")

        Add_New_Client_data = (client_name, client_email, client_nationalid)
        cur.execute(Add_New_Client_query, Add_New_Client_data)
        self.connectionString.commit()
        self.statusBar().showMessage('New Client Addedd Successfuly'+ str(Add_New_Client_data))
        self.Show_All_Client()
        cur.close()

    def Search_Client(self):
        cur = self.connectionString.cursor()

        client_nationalid = self.lineEdit_6.text()

        search_client_query = "SELECT * FROM clients WHERE client_nationalid = %s"
        search_client_data = [client_nationalid]
   
        cur.execute(search_client_query, search_client_data)
        data = cur.fetchone()

        if (data):
            self.id_client.setText(str(data[0]))
            self.lineEdit_15.setText(data[1])
            self.lineEdit_26.setText(data[2])
            self.lineEdit_27.setText(data[3])
        else:
            self.statusBar().showMessage('No Record found')
        self.connectionString.commit()
        cur.close()

    def Edit_Client(self):
        cur = self.connectionStringcursor().cursor()

        client_id = self.id_client.text()
        client_name =  self.lineEdit_15.text()
        client_email =  self.lineEdit_26.text()
        client_nationalid =  self.lineEdit_27.text()

        client_nationalid = self.lineEdit_6.text()
        
        edit_client_query = ("update clients set client_name=%s, client_email=%s, client_nationalid=%s "
                        "Where idclients=%s")

        edit_client_query_data = (client_name, client_email, client_nationalid, client_id)
        
        cur.execute(edit_client_query, edit_client_query_data)

        self.connectionString.commit()

        self.statusBar().showMessage('Edit Client Successfuly'+ str(edit_client_query_data))
        self.Show_All_Client()
        cur.close()

    def Delete_Client(self):
        cur = self.connectionStringcursor().cursor()

        client_id = self.id_client.text()

        client_nationalid = self.lineEdit_6.text()
        
        warning = QMessageBox.warning(self, 'Delete Client', "Are You Sure You Want To Delete This Book", QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes :
            delete_book_query = ("delete from clients where idclients=%s")
            cur.execute(delete_book_query, [(client_id)])
            self.connectionString.commit()
            self.statusBar().showMessage('Client Deleted')
            self.Show_All_Client()
        cur.close()

    #########################################
    ################# Users #################


    def Add_New_User(self):
        cur = self.connectionString.cursor()

        user_name= self.user_name.text()
        user_email = self.user_email.text()
        user_password = self.user_password.text()
        user_again_password = self.user_again_password.text()

        if user_password != user_again_password:
            self.statusBar().showMessage("Password and Again Password did not matched. Please enter password again")
            self.user_password.setText('')
            self.user_again_password.setText('')
        else:
            add_new_user_query = "INSERT INTO users (user_name, user_email, user_password) VALUES (%s, %s, SHA1(%s))"
            Add_New_User_data = (user_name, user_email, user_password)
            
            cur.execute(add_new_user_query, Add_New_User_data)
            
            self.connectionString.commit()
            self.statusBar().showMessage('New user addedd successfully')
        cur.close()


    def Login(self):
        cur = self.connectionString.cursor()

        user_name = self.lineEdit_21.text()
        password = self.lineEdit_20.text()

        login_query = ("Select * From users")

        cur.execute(login_query)
        data = cur.fetchall()
        for row in data :
            if user_name == row[1] and password == row[3]:
                self.statusBar().showMessage('Valid Username & Password')
                self.groupBox_4.setEnabled(True)
                self.user_key.setText(str(row[0]))
                self.lineEdit_18.setText(row[1])
                self.lineEdit_17.setText(row[2])
                self.lineEdit_16.setText(row[3])
        cur.close()


    def Edit_User(self):
        user_id = self.user_key.text()
        user_name = self.lineEdit_18.text()
        user_email = self.lineEdit_17.text()
        user_password = self.lineEdit_16.text()
        user_again_password = self.lineEdit_19.text()


        if user_password == user_again_password :
            cur = self.connectionString.cursor()

            edit_user_query = ("update users set user_name=%s, user_email=%s, user_password=%s Where id=%s")

            edit_user_query_data = (user_name, user_email, user_password, user_id)
        
            cur.execute(edit_user_query, edit_user_query_data)

            self.connectionString.commit()
            self.statusBar().showMessage('User Data Update Successfuly')
        else:
            self.statusBar().showMessage("Make sure you entered your password correctly")
        cur.close()
    
    def Delete_User(self):
        cur = self.connectionString.cursor()

        user_Name = self.lineEdit_18.text()
        
        warning = QMessageBox.warning(self, 'Delete users', "Are You Sure You Want To Delete This User", QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes :
            delete_user_query = ("delete from users where user_name=%s")
            cur.execute(delete_user_query, [(user_Name)])
            self.connectionString.commit()

            self.statusBar().showMessage('User Deleted')
        cur.close()

    ############################################
    ################# Settings #################


    def Add_Category(self):
        cur = self.connectionString.cursor()

        category_name = self.lineEdit_22.text()

        queryString = 'INSERT INTO category (category_name) VALUES ("' + category_name + '")'

        cur.execute(queryString)

        self.connectionString.commit()
        self.statusBar().showMessage('New Category Addedd : ' + category_name)
        self.lineEdit_22.setText('')
        self.Show_Category()
        self.Show_Category_Combobox()
        cur.close()


    def Show_Category(self):
        cur = self.connectionString.cursor()
        queryString = 'Select category_name from category order by category_name asc;'
        cur.execute(queryString)
        data = cur.fetchall()
        
        if data:
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.insertRow(0)
            for row, form in enumerate(data):
                for column , item in enumerate(form) :
                    self.tableWidget_2.setItem(row , column , QTableWidgetItem(str(item)))
                    column += 1

                row_position = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_position)
        cur.close()

    def Delete_Category(self):
        cur = self.connectionString.cursor()

        category_Name = self.lineEdit_28.text()
        
        warning = QMessageBox.warning(self, 'Delete category', "Are You Sure You Want To Delete This Book", QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes :
            delete_category_query = ("delete from category where category_Name=%s")
            cur.execute(delete_category_query, [(category_Name)])
            self.connectionString.commit()
            self.statusBar().showMessage('category Deleted')
            self.Show_Category()
        cur.close()

    def Add_Author(self):
        cur = self.connectionString.cursor()

        author_name = self.lineEdit_23.text()

        queryString = 'INSERT INTO author (author_name) VALUES ("' + author_name + '")'
        cur.execute(queryString)
        self.connectionString.commit()
        cur.close()
        self.statusBar().showMessage('New Author Addedd : ' + author_name)
        self.lineEdit_23.setText('')
        self.Show_Author()
        self.Show_Author_Combobox()


    def Show_Author(self):
        cur = self.connectionString.cursor()
        queryString = 'Select author_name from author order by author_name asc'

        cur.execute(queryString)
        data = cur.fetchall()
        
        if data :
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.insertRow(0)
            for row , form in enumerate(data):
                for column , item in enumerate(form) :
                    self.tableWidget_3.setItem(row , column , QTableWidgetItem(str(item)))
                    column += 1

                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)
        cur.close()

    def Delete_Author(self):
        cur = self.connectionString.cursor()

        author_Name = self.lineEdit_29.text()
        
        warning = QMessageBox.warning(self, 'Delete author', "Are You Sure You Want To Delete This author", QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes :
            delete_author_query = ("delete from author where author_Name=%s")
            cur.execute(delete_author_query, [(author_Name)])
            self.connectionString.commit()
            self.statusBar().showMessage('author Deleted')
            self.Show_Publisher()
        cur.close()


    def Add_Publisher(self):
        cur = self.connectionString.cursor()

        publisher_name = self.lineEdit_24.text()

        queryString = 'INSERT INTO publisher (publisher_name) VALUES ("' + publisher_name + '")'

        print(queryString)

        cur.execute(queryString)
        self.connectionString.commit()
        cur.close()
        self.statusBar().showMessage('New Publisher Addedd : ' + publisher_name)
        self.lineEdit_24.setText('')
        print('New Publisher Addedd : ' + publisher_name)
        self.Show_Publisher()
        self.Show_Publisher_Combobox()


    def Show_Publisher(self):
        cur = self.connectionString.cursor()

        queryString = 'Select publisher_name from publisher order by publisher_name asc'

        cur.execute(queryString)
        data = cur.fetchall()
        
        if data :
            self.tableWidget_4.setRowCount(0)
            self.tableWidget_4.insertRow(0)
            for row , form in enumerate(data):
                for column , item in enumerate(form) :
                    self.tableWidget_4.setItem(row , column , QTableWidgetItem(str(item)))
                    column += 1

                row_position = self.tableWidget_4.rowCount()
                self.tableWidget_4.insertRow(row_position)

        cur.close()


    def Delete_Publisher(self):
        cur = self.connectionString.cursor()


        publisher_Name = self.lineEdit_30.text()

        warning = QMessageBox.warning(self, 'Delete publisher', "Are You Sure You Want To Delete This Publisher", QMessageBox.Yes | QMessageBox.No)
        if warning == QMessageBox.Yes :
            delete_category_query = ("delete from publisher where publisher_Name=%s")
            cur.execute(delete_category_query, [(publisher_Name)])
            self.connectionString.commit()
            self.statusBar().showMessage('Publisher Deleted')
            self.Show_Publisher()
        cur.close()


    ############################################################
    ################# Show Settings Data In Ui #################


    def Show_Category_Combobox(self):
        cur = self.connectionString.cursor()

        queryString = 'Select category_name FROM category'
        cur.execute(queryString)
        
        data = cur.fetchall()

        self.comboBox_9.clear()
        self.comboBox_3.clear()
        for category in data :
            self.comboBox_9.addItem(category[0])
            self.comboBox_3.addItem(category[0])
        self.connectionString.commit()
        cur.close()


    def Show_Author_Combobox(self):
        cur = self.connectionString.cursor()

        queryString = 'Select author_name FROM author'
        cur.execute(queryString)
        
        data = cur.fetchall()

        self.comboBox_11.clear()
        self.comboBox_5.clear()
        for author in data :
           # print(author)
            self.comboBox_11.addItem(author[0])
            self.comboBox_5.addItem(author[0])
        self.connectionString.commit()
        cur.close()


    def Show_Publisher_Combobox(self):
        cur = self.connectionString.cursor()

        queryString = 'Select publisher_name FROM publisher'
        cur.execute(queryString)
        
        data = cur.fetchall()

        self.comboBox_10.clear()
        self.comboBox_4.clear()
        for publisher in data :
            self.comboBox_10.addItem(publisher[0])
            self.comboBox_4.addItem(publisher[0])
        self.connectionString.commit()
        cur.close()


    #############################################
    ################# Export Data #################


    def Export_Day_Operations(self):
        cur = self.connectionString.cursor()


        show_handel_day_operation_query = ("SELECT book_name, client, type, date_days, to_days,  days FROM dayoperations")
        cur.execute(show_handel_day_operation_query)
        data = cur.fetchall()

        wb = Workbook('day_operations.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0, 'Book Title')
        sheet1.write(0,1, 'Client Name')
        sheet1.write(0,2, 'Type')
        sheet1.write(0,3, 'from - date')
        sheet1.write(0,4, 'to - date')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        cur.close()
        self.statusBar().showMessage('Report Created Successfully')


    def Export_Clients(self):
        cur = self.connectionString.cursor()


        Show_All_Client_query = ("Select client_name, client_email, client_nationalid From clients")
        cur.execute(Show_All_Client_query)
        data = cur.fetchall()

        wb = Workbook('All_Clients.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0, 'Client Name')
        sheet1.write(0,1, 'Client Email')
        sheet1.write(0,2, 'Client National Id')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        cur.close
        self.statusBar().showMessage('Report All Clients Successfully')        


    def Export_Books(self):
        cur = self.connectionString.cursor()


        Show_All_Client_query = ("Select book_code, book_name, book_description, book_category, book_author, book_publisher, book_price From book ")
        cur.execute(Show_All_Client_query)
        data = cur.fetchall()

        wb = Workbook('All Books.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0, 'Book Name')
        sheet1.write(0,1, 'Book Description')
        sheet1.write(0,2, 'Book Category')
        sheet1.write(0,3, 'Book author')
        sheet1.write(0,4, 'Book Publisher')
        sheet1.write(0,5, 'Book Price')

        row_number = 1
        for row in data:
            column_number = 0
            for item in row:
                sheet1.write(row_number, column_number, str(item))
                column_number += 1
            row_number += 1

        wb.close()
        cur.close
        self.statusBar().showMessage('Report All Books Successfully')


    #############################################
    ################# Ui Themes #################


    def Dark_Blue_Theme(self):
        style = open('Themes/darkblue.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Dark_Gray_Theme(self):
        style = open('Themes/darkgray.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Dark_Orange_Theme(self):
        style = open('Themes/darkorange.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    def QDark_Theme(self):
        style = open('Themes/qdark.css', 'r')
        style = style.read()
        self.setStyleSheet(style)


def main():
    app = QApplication(sys.argv)
    window = Login()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
