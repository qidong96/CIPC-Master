#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# !/usr/bin/env python
# coding: utf-8


from re import match
from sys import argv, exit

import pymysql
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import *
from numpy import nan
from pandas import DataFrame, read_excel, concat
from pypinyin import Style, pinyin


class Ui_MainWindow(object):

    def __init__(self):
        self.regulated_table = '业务数据字段命名及设计规范.xlsx'
        self.table_name = ''
        self.file_path = ''
        self.database_name = ''
        self.skiprows = 0
        self.password = ''
        self.user = ''
        self.IP = ''
        self.char = ''
        self.flag = False
        self.dbport = 3306
        self.exec_code_ls = []

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.textEdit_2 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_2.setGeometry(QtCore.QRect(210, 97, 149, 41))
        self.textEdit_2.setObjectName("textEdit_2")
        self.textEdit_3 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_3.setGeometry(QtCore.QRect(210, 152, 149, 41))
        self.textEdit_3.setObjectName("textEdit_3")
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setGeometry(QtCore.QRect(430, 20, 321, 511))
        self.textBrowser.setObjectName("textBrowser")
        self.textEdit_4 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_4.setGeometry(QtCore.QRect(210, 207, 149, 41))
        self.textEdit_4.setObjectName("textEdit_4")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(240, 490, 93, 28))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(70, 490, 93, 28))
        self.pushButton.setObjectName("pushButton")


        self.textEdit_5 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_5.setGeometry(QtCore.QRect(210, 262, 149, 41))
        self.textEdit_5.setObjectName("textEdit_5")
        self.textEdit_6 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_6.setGeometry(QtCore.QRect(210, 317, 149, 41))
        self.textEdit_6.setObjectName("textEdit_6")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(210, 440, 87, 22))
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.textEdit_7 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_7.setGeometry(QtCore.QRect(210, 372, 149, 41))
        self.textEdit_7.setObjectName("textEdit_7")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setGeometry(QtCore.QRect(51, 32, 152, 441))
        self.widget.setObjectName("widget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.widget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(self.widget)
        self.label.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.label_2 = QtWidgets.QLabel(self.widget)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2)
        self.label_3 = QtWidgets.QLabel(self.widget)
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3)
        self.label_4 = QtWidgets.QLabel(self.widget)
        self.label_4.setObjectName("label_4")
        self.verticalLayout.addWidget(self.label_4)
        self.label_5 = QtWidgets.QLabel(self.widget)
        self.label_5.setObjectName("label_5")
        self.verticalLayout.addWidget(self.label_5)
        self.label_6 = QtWidgets.QLabel(self.widget)
        self.label_6.setObjectName("label_6")
        self.verticalLayout.addWidget(self.label_6)
        self.label_7 = QtWidgets.QLabel(self.widget)
        self.label_7.setObjectName("label_7")
        self.verticalLayout.addWidget(self.label_7)
        self.label_8 = QtWidgets.QLabel(self.widget)
        self.label_8.setObjectName("label_8")
        self.verticalLayout.addWidget(self.label_8)
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(210, 41, 149, 41))
        self.textEdit.setObjectName("textEdit")
        self.textEdit.raise_()
        self.textEdit_2.raise_()
        self.textEdit_3.raise_()
        self.textBrowser.raise_()
        self.textEdit_4.raise_()
        self.pushButton_2.raise_()
        self.pushButton.raise_()
        self.textEdit_5.raise_()
        self.textEdit_6.raise_()
        self.comboBox.raise_()
        self.label_8.raise_()
        self.textEdit_7.raise_()
        self.textEdit.raise_()

        self.pushButton.clicked.connect(self.openfile)
        self.pushButton_2.clicked.connect(self.processing)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton_2.setText(_translate("MainWindow", "开始转换"))
        self.pushButton.setText(_translate("MainWindow", "打开文件"))
        self.comboBox.setItemText(0, _translate("MainWindow", "gbk"))
        self.comboBox.setItemText(1, _translate("MainWindow", "utf8"))
        self.label.setText(_translate("MainWindow", "表格名"))
        self.label_2.setText(_translate("MainWindow", "读取表格时跳过的行数"))
        self.label_3.setText(_translate("MainWindow", "数据库名"))
        self.label_4.setText(_translate("MainWindow", "数据库密码"))
        self.label_5.setText(_translate("MainWindow", "用户名"))
        self.label_6.setText(_translate("MainWindow", "数据库IP地址"))
        self.label_7.setText(_translate("MainWindow", "端口"))
        self.label_8.setText(_translate("MainWindow", "文件编码方式"))

    def openfile(self):
        file_path = QFileDialog.getOpenFileName(None, '打开文件', '', 'Excel files(*.xlsx , *.xls);;CSV Files(*.csv)')
        self.file_path = file_path[0]
        self.printf('成功选择：' + self.file_path)

    def printf(self, mes):  # 输出函数
        self.textBrowser.append(str(mes))  # 在指定的区域显示提示信息,一定得是字符型
        cursot = self.textBrowser.textCursor()
        self.textBrowser.moveCursor(cursot.End)  # 光标移到最后，这样就会自动显示出来
        QtWidgets.QApplication.processEvents()  # 一定加上这个功能，不然有卡顿

    def dataDate_mode(self, file_name):
        '''
        判断文件统计口径是时间点or时间段
        '''
        file_name = file_name.split('/')[-1]
        file_name = file_name[:file_name.find('.')]
        file_name_ls = file_name.split('_')

        TF_list = []
        # 正则匹配yyyyMMdd格式
        pattern = '((\d{3}[1-9]|\d{2}[1-9]\d|\d[1-9]\d{2}|[1-9]\d{3})(((0[13578]|1[02])(0[1-9]|[12]\d|3[01]))|((0[469]|11)(0[1-9]|[12]\d|30))|(02(0[1-9]|[1]\d|2[0-8]))))|(((\d{2})(0[48]|[2468][048]|[13579][26])|((0[48]|[2468][048]|[3579][26])00))0229)'
        for param in file_name_ls:
            if bool(match(pattern, param, flags=0)):
                TF_list.append(param)

        return TF_list

    def bopomofo_joint(self, a):
        '''
        连接中文拼音
        '''
        x = ''
        for i in a:
            x += i[0]
        return x.upper()

    def isNull(self, values):
        '''
        填数SQL语句空值处理方法
        '''
        error_type = ['nan', 'NaN', 'NaT']
        if str(values) in error_type:
            output = 'NULL,'
        else:
            output = '\'' + str(values) + '\','
        return output

        # 时点表创建

    def create_table_timePoint(self, file_path, label, label_2, regulated_table, cursor):
        file_name = file_path
        file_name = file_name.split('/')[-1]
        file_name = file_name[:file_name.find('.')]
        df = read_excel(file_path, skiprows=int(label_2))
        print(df.head())
        # 读取<规范表>
        df_address = read_excel(regulated_table)
        df_address.columns = ['中文名', '英文名', 'SQL语句']
        dict_address = dict(zip(df_address['中文名'], df_address['SQL语句']))
        # 根据表头建库
        sqltxt = ['CREATE TABLE ' + label + '(\n', "    ID INT NOT NULL AUTO_INCREMENT  COMMENT 'ID' ,\n",
                  "    DATA_DATE DATE    COMMENT '数据时间' ,\n", "    REVISION INT    COMMENT '乐观锁' ,\n",
                  "    CREATED_BY VARCHAR(32)    COMMENT '创建人' ,\n", "    CREATED_TIME DATETIME    COMMENT '创建时间' ,\n",
                  "    UPDATED_BY VARCHAR(32)    COMMENT '更新人' ,\n", "    UPDATED_TIME DATETIME    COMMENT '更新时间' ,\n"]
        for col_name in df.columns:

            if col_name in dict_address.keys():
                sqltxt.append(dict_address[col_name])
            else:
                self.printf("新增字段名: "+col_name)
                type_name = 'VARCHAR(255)'
                sentence = '    ' + self.bopomofo_joint(
                    pinyin(col_name, style=Style.NORMAL)) + "   " + type_name + "     COMMENT '" + col_name + "' ,\n"
                sqltxt.append(sentence)
                dict_address[col_name] = sentence

        sqltxt.append("    PRIMARY KEY (ID)\n) COMMENT = '" + file_name.split('.')[0] + " ';;")
        # 创建表
        creatTable_sql = "drop table if exists " + label
        cursor.execute(creatTable_sql)
        # list 转 str
        sql = "".join(sqltxt)
        # 执行sql
        cursor.execute(sql)
        self.printf('创建数据库成功!')

    # 时段表创建
    def create_table_timeQuantum(self, file_path, label, label_2, regulated_table, cursor):
        file_name = file_path
        file_name = file_name.split('/')[-1]
        file_name = file_name[:file_name.find('.')]
        df = read_excel(file_path, skiprows=int(label_2))
        # 读取<规范表>
        try:
            df_address = read_excel(regulated_table)
        except:
            self.printf('规范表不存在')

        df_address.columns = ['中文名', '英文名', 'SQL语句']
        dict_address = dict(zip(df_address['中文名'], df_address['SQL语句']))
        # 根据表头建库
        sqltxt = ['CREATE TABLE ' + label + '(\n', "    ID INT NOT NULL AUTO_INCREMENT  COMMENT 'ID' ,\n",
                  "    DATA_START_DATE DATE    COMMENT '数据开始时间' ,\n", "    REVISION INT    COMMENT '乐观锁' ,\n",
                  "    CREATED_BY VARCHAR(32)    COMMENT '创建人' ,\n", "    CREATED_TIME DATETIME    COMMENT '创建时间' ,\n",
                  "    UPDATED_BY VARCHAR(32)    COMMENT '更新人' ,\n", "    UPDATED_TIME DATETIME    COMMENT '更新时间' ,\n"]
        for col_name in df.columns:

            if col_name in dict_address.keys():
                sqltxt.append(dict_address[col_name])
            else:
                self.printf("新增字段名: "+col_name)
                type_name = 'VARCHAR(255)'
                sentence = '    ' + self.bopomofo_joint(
                    pinyin(col_name, style=Style.NORMAL)) + "   " + type_name + "     COMMENT '" + col_name + "' ,\n"
                sqltxt.append(sentence)
                dict_address[col_name] = sentence

        sqltxt.append("    DATA_END_DATE DATE    COMMENT '数据结束时间' ,\n")
        sqltxt.append("    PRIMARY KEY (ID)\n) COMMENT = '" + file_name.split('.')[0] + " ';;")
        # 创建表
        creatTable_sql = "drop table if exists " + label
        cursor.execute(creatTable_sql)
        # list 转 str
        sql = "".join(sqltxt)
        # 执行sql
        cursor.execute(sql)
        self.printf('创建数据库成功!')

    def insert_to_sql(self, exec_code_ls, cursor, conn):
        for i in range(len(exec_code_ls)):
            sql = exec_code_ls[i]
            try:
                # 执行sql语句
                cursor.execute(sql)
                # 提交到数据库执行
                conn.commit()

            except:
                # 如果发生错误则回滚
                conn.rollback()

        conn.close()

    def excel_to_sql_timePoint(self, file_path, TF_list, label, label_2):
        '''
        按时间点的模版插入SQL语句
        '''
        Data_Date = TF_list[0]
        df = read_excel(file_path, skiprows=int(label_2))
        Table_Header = DataFrame({'ID': nan * len(df), 'DATA_DATE': [Data_Date] * len(df)})
        df = concat([Table_Header, df], axis=1)
        cols_len = len(df.iloc[0])
        exec_code_ls = []

        for j in range(len(df)):
            col_values = 'INSERT INTO ' + label + ' values('
            # 将"乐观锁",“创建人”,“创建时间”,“更新人”,“更新时间列”设置为空值
            for i in range(cols_len):
                if i == 2:
                    col_values += 'NULL,' * 5

                col_values += self.isNull(df.iloc[j][i])

            # 去除最后一个逗号
            col_values = col_values[:-1]

            # 加上后缀
            col_values += ');'
            exec_code_ls.append(col_values)

        return exec_code_ls

    def excel_to_sql_timeQuantum(self, file_path, TF_list, label, label_2):
        '''
        按时间段的模版插入SQL语句
        '''
        Data_Start_Date = min(TF_list)
        Data_End_Date = max(TF_list)
        df = read_excel(file_path, skiprows=int(label_2))
        Table_Header = DataFrame({'ID': nan * len(df), 'DATA_START_DATE': [Data_Start_Date] * len(df)})
        df = concat([Table_Header, df], axis=1)
        end_date = DataFrame({'DATA_END_DATE': [Data_End_Date] * len(df)})
        df = concat([df, end_date], axis=1)
        cols_len = len(df.iloc[0])
        exec_code_ls = []

        for j in range(len(df)):
            col_values = 'INSERT INTO ' + label + ' values('
            # 将"乐观锁",“创建人”,“创建时间”,“更新人”,“更新时间列”设置为空值
            for i in range(cols_len):
                if i == 2:
                    col_values += 'NULL,' * 5

                col_values += self.isNull(df.iloc[j][i])

            # 去除最后一个逗号
            col_values = col_values[:-1]

            # 加上后缀
            col_values += ');'
            exec_code_ls.append(col_values)

        return exec_code_ls

    def connect_database(self, db, pw, user, IP, char, por):
        # 配置需连接的数据库
        conn = pymysql.connect(host=IP, user=user, passwd=pw, db=db, charset=char, port=por)
        cursor = conn.cursor()
        return conn, cursor

    def excel_to_sql_mode(self):
        # 输出数据入库SQL语句于exec_code_ls表
        TF_list = self.dataDate_mode(self.file_path)
        # 统计口径为时间点
        if len(TF_list) == 1:
            self.exec_code_ls = self.excel_to_sql_timePoint(self.file_path, TF_list, self.table_name, self.skiprows)
            # 统计口径为时间段
        elif len(TF_list) == 2:
            self.exec_code_ls = self.excel_to_sql_timeQuantum(self.file_path, TF_list, self.table_name, self.skiprows)
        else:
            self.printf("文件名异常!")

        try:
            conn, cursor = self.connect_database(self.database_name, self.password, self.user, self.IP, self.char,
                                                 self.dbport)
        except:
            self.printf('连接数据库失败')


        # 下判断数据库中要插入的表是否存在: 若不存在，则新建该表
        test_sql = 'show tables;'
        cursor.execute(test_sql)
        data = cursor.fetchall()

        tables = []
        for i in range(len(data)):
            tables.append(data[i][0])

        # 若判断为表存在:
        if self.label in tables:
            self.insert_to_sql(self.exec_code_ls, cursor, conn)

        else:
            if len(TF_list) == 1:
                self.create_table_timePoint(self.file_path, self.table_name, self.skiprows, self.regulated_table,
                                            cursor)
                self.insert_to_sql(self.exec_code_ls, cursor, conn)

            elif len(TF_list) == 2:
                self.create_table_timeQuantum(self.file_path, self.table_name, self.skiprows, self.regulated_table,
                                              cursor)
                self.insert_to_sql(self.exec_code_ls, cursor, conn)

            else:
                pass


    def processing(self):  # 处理函数

        self.table_name = self.textEdit.toPlainText()  # 数据表的名字

        if self.textEdit_2.toPlainText().strip():
            self.skiprows = int(self.textEdit_2.toPlainText())  # 读取表格时跳过的行数
        else:
            self.skiprows = 0

        self.database_name = self.textEdit_3.toPlainText()  # 数据库的名字

        self.password = self.textEdit_4.toPlainText()  # 数据库的密码

        if self.textEdit_5.toPlainText().strip():
            self.user = self.textEdit_5.toPlainText()
        else:
            self.user = 'root'

        if self.textEdit_6.toPlainText().strip():
            self.IP = self.textEdit_6.toPlainText()
        else:
            self.IP = 'localhost'

        self.char = self.comboBox.currentText()

        if self.textEdit_7.toPlainText().strip():
            self.dbport = int(self.textEdit_7.toPlainText().strip())
        else:
            self.dbport = 3306

        '''
        self.file_path#路径名
        填充代码完成入库操作
        '''
        self.excel_to_sql_mode()


    def main(self):
        application = QApplication(argv)
        mainWindow = QMainWindow()
        userInterface = Ui_MainWindow()
        userInterface.setupUi(mainWindow)
        mainWindow.show()
        exit(application.exec_())


if __name__ == '__main__':
    Ui = Ui_MainWindow()
    Ui.main()
