#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019/2/26 21:31
# @Author  : joestww@163.com
# @Site    : 
# @File    : GUI.py
# @Software: PyCharm

import sys
import os
from PyQt4 import QtGui,QtCore
from ExcelDiff import *
from functools import partial
import time
from threading import Thread

class MainUi(QtGui.QMainWindow):
    def __init__(self, parent=None):
        super(MainUi, self).__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.setBaseSize(960,700)
        self.setWindowTitle("Excel Diff")
        self.setWindowOpacity(0.97)  # 设置窗口透明度
        #self.setAttribute(QtCore.Qt.WA_TranslucentBackground)  # 设置窗口背景透明
        #窗口主部件
        self.main_widget=QtGui.QWidget()
        self.main_layout=QtGui.QGridLayout()
        self.main_widget.setLayout(self.main_layout)

        #上侧部件
        self.top_widget=QtGui.QWidget()
        self.top_widget.setObjectName('top_widget')
        self.top_layout=QtGui.QGridLayout()
        self.top_widget.setLayout(self.top_layout)
        self.top_widget.setStyleSheet('''
                    QWidget#top_widget{
                        background:#F0FFF0;
                        border-top:1px solid white;
                        border-bottom:1px solid white;
                        border-left:1px solid white;
                        border-top-left-radius:10px;
                        border-bottom-left-radius:10px;
                        border-top-right-radius:10px;
                        border-bottom-right-radius:10px;
                    }
                    QLabel#left_label{
                        border:none;
                        font-size:20px;
                        font-weight:700;
                        font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
                    }
                ''')

        #下侧部件
        self.bottom_widget = QtGui.QWidget()
        self.bottom_widget.setObjectName('bottom_widget')
        self.bottom_layout = QtGui.QGridLayout()
        self.bottom_widget.setLayout(self.bottom_layout)

        #布局管理
        self.main_layout.addWidget(self.top_widget,0,0,2,14)
        self.main_layout.addWidget(self.bottom_widget,2,0,19,14)
        self.main_layout.setSpacing(0)
        self.setCentralWidget(self.main_widget)

        self.make_top()

    def make_top(self):
        # 文件选择按钮
        self.file_text1 = QtGui.QLineEdit()
        self.file_button1 = QtGui.QPushButton(u'请选择文件1: ')
        self.file_button1.clicked.connect(partial(self.select_file,self.file_text1))
        self.file_text2 = QtGui.QLineEdit()
        self.file_button2 = QtGui.QPushButton(u'请选择文件2: ')
        self.file_button2.clicked.connect(partial(self.select_file, self.file_text2))

        # diff动作按钮
        self.diff_button = QtGui.QPushButton(u'diff')
        self.diffrst=None
        self.diff_button.clicked.connect(self.diff)
        self.top_layout.addWidget(self.file_button1,0,0,1,1)
        self.top_layout.addWidget(self.file_text1,0,1,1,2)
        self.top_layout.addWidget(self.file_button2,0,4,1,1)
        self.top_layout.addWidget(self.file_text2,0,5,1,2)
        self.top_layout.addWidget(self.diff_button,0,13,1,1)

        #进度条
        self.process_bar=QtGui.QProgressBar()
        self.process_bar.setValue(0)
        self.timer=QtCore.QBasicTimer()
        self.timer_step=0
        self.process_bar.setFixedHeight(10)
        self.top_layout.addWidget(self.process_bar,1,1,1,10)

        #构建下部区域
        #self.make_bottom()

    #选择文件动作
    def select_file(self,txt):
        filepath=QtGui.QFileDialog.getOpenFileName(self,u'选择文件','./','Excel(*.xls*);;All Files(*)')
        txt.setText(filepath)

    #diff动作
    def diff(self):
        self.timer_step=0
        if not os.path.isfile(self.file_text1.text().__str__()):
            self.file_text1.setText('')
            self.file_text1.setPlaceholderText(u"请选择文件")
        else:
            if not os.path.isfile(self.file_text2.text().__str__()):
                self.file_text2.setText('')
                self.file_text2.setPlaceholderText(u"请选择文件")
            else:
                for i in reversed(range(0, self.bottom_layout.count())):
                    self.bottom_layout.itemAt(i).widget().setParent(None)
                self.process_bar.reset()
                timeStart=time.time()
                diffthread = Thread(target=self.diffThread)
                prossbarthread=Thread(target=self.processBarThread)
                diffthread.start()
                prossbarthread.start()
                diffthread.join()
                prossbarthread.join()
                timeEnd=time.time()
                # print u"time consumed %.3fseconds"%(timeEnd-timeStart)
                self.position=self.positionFind()
                self.make_bottom()
                self.process_bar.setValue(100)
                # print u'Diff successed'

    def diffThread(self):
        excelDiff = ExcelDiff()
        self.diffrst = excelDiff.main(self.file_text1.text(), self.file_text2.text())
        self.timer_step=99
    def processBarThread(self):
        while self.timer_step<99:
            self.timer_step+=1
            self.process_bar.setValue(self.timer_step)
            time.sleep(0.5)

    def timerEvent(self, *args, **kwargs):
        if self.timer_step>=99:
            self.timer.stop()
            return
        self.timer_step+=1
        self.process_bar.setValue(self.timer_step)

    def make_bottom(self):
        for i in reversed(range(0, self.bottom_layout.count())):
            self.bottom_layout.itemAt(i).widget().setParent(None)
        #左侧部件
        self.left_widget=QtGui.QWidget()
        self.left_widget.setObjectName('left_widget')
        self.left_layout=QtGui.QGridLayout()
        self.left_widget.setLayout(self.left_layout)

        #右侧部件
        self.right_widget=QtGui.QWidget()
        self.right_layout=QtGui.QGridLayout()
        self.right_widget.setLayout(self.right_layout)

        #显示两个sheet的部件
        self.tabWidget1 = QtGui.QTableWidget()
        self.tabWidget2 = QtGui.QTableWidget()
        #增删改的按钮
        self.row_button=QtGui.QPushButton(u'行增删')
        self.col_button=QtGui.QPushButton(u'列增删')
        self.cell_button=QtGui.QPushButton(u'单元格改动')
        self.clearHighlight_button=QtGui.QPushButton(u'清除高亮')
        self.detail=QtGui.QTableWidget()
        self.right_layout.addWidget(self.tabWidget1,0,0,15,6)
        self.right_layout.addWidget(self.tabWidget2,0,6,15,6)
        self.right_layout.addWidget(self.row_button,15,0,1,2)
        self.right_layout.addWidget(self.col_button,15,2,1,2)
        self.right_layout.addWidget(self.cell_button,15,4,1,2)
        self.right_layout.addWidget(self.clearHighlight_button,15,6,1,2)
        self.right_layout.addWidget(self.detail,16,0,3,12)

        #主布局管理
        self.bottom_layout.addWidget(self.left_widget,0,0,19,1)
        self.bottom_layout.addWidget(self.right_widget,0,1,19,13)

        self.make_left()

    def make_left(self):
        #构建左侧sheet显示按钮
        for i in reversed(range(0, self.left_layout.count())):
            self.left_layout.itemAt(i).widget().setParent(None)
        label = QtGui.QLabel('sheetnames')
        label.setObjectName('left_label')
        self.left_layout.addWidget(label,0,0,1,2,QtCore.Qt.AlignTop)
        self.sheet_button_list=[]
        for i in range(len(self.diffrst)):
            sheet_button = QtGui.QPushButton()
            sheet_button.setText(self.diffrst[i][0])
            sheet_button.clicked.connect(partial(self.show_sheets,i))
            self.left_layout.addWidget(sheet_button,i+1,0,1,1,QtCore.Qt.AlignTop)
            self.sheet_button_list.append(sheet_button)
        self.left_layout.setAlignment(QtCore.Qt.AlignTop)
        self.click_sheet=0
        self.show_sheets(self.click_sheet)

        self.left_widget.setStyleSheet('''
            QWidget#left_widget{
                background:gray;
                border-top:1px solid white;
                border-bottom:1px solid white;
                border-left:1px solid white;
                border-top-left-radius:10px;
                border-bottom-left-radius:10px;
                border-top-right-radius:10px;
                border-bottom-right-radius:10px;
            }
            QLabel#left_label{
                border:none;
                font-size:20px;
                font-weight:700;
                font-family: "Helvetica Neue", Helvetica, Arial, sans-serif;
            }
        ''')

    def show_sheets(self,i):
        self.detail.clearContents()
        self.sheet_button_list[self.click_sheet].setStyleSheet('background-color:rgb(255,255,255)')
        self.click_sheet=i
        self.sheet_button_list[i].setStyleSheet('background-color:rgb(100,100,150)')
        diff=self.diffrst[i]
        self.set_sheets(self.tabWidget1,diff[1],1,i)
        self.set_sheets(self.tabWidget2,diff[2],2,i)
        self.show_detail(diff[3],i)

    def set_sheets(self,tw,diff,num,sn):
        lrow=len(diff)
        tw.setRowCount(lrow)
        if lrow>0:
            lcol=max([len(x)for x in diff])
        else:
            lcol=0
        tw.setColumnCount(lcol)
        vheadlable=[]

        if num >1:
            for i in range(lrow):
                if self.position[sn][1][1][i]<0:
                    vheadlable.append(' ')
                else:
                    vheadlable.append(str(self.position[sn][1][1][i]+1))
        else:
            for i in range(lrow):
                if self.position[sn][1][0][i]<0:
                    vheadlable.append(' ')
                else:
                    vheadlable.append(str(self.position[sn][1][0][i]+1))
        tw.setVerticalHeaderLabels(vheadlable)
        hheadlable = []
        if num > 1:
            for i in range(lcol):
                if self.position[sn][1][3][i]<0:
                    hheadlable.append(' ')
                else:
                    hheadlable.append(self.colToString(self.position[sn][1][3][i]))
        else:
            for i in range(lcol):
                if self.position[sn][1][2][i]<0:
                    hheadlable.append(' ')
                else:
                    hheadlable.append(self.colToString(self.position[sn][1][2][i]))
        tw.setHorizontalHeaderLabels(hheadlable)

        for i in range(len(diff)):
            for j in range(len(diff[i])):
                dtxt=diff[i][j][1]
                if dtxt or dtxt==0:
                    txt = self.mtoString(dtxt)#str(dtxt)
                else:
                    txt=' '
                twi=QtGui.QTableWidgetItem(txt)
                if diff[i][j][0]:
                    if diff[i][j][0]<0:
                        twi.setBackground(QtGui.QColor(233,150,122))#红色
                    elif diff[i][j][0]>1:
                        twi.setBackground(QtGui.QColor(255,255,51))#橘色
                    else:
                        twi.setBackground(QtGui.QColor(135,206,235))#蓝色
                else:
                    twi.setBackground(QtGui.QColor(255,255,255))
                tw.setItem(i, j,twi)

    #列号转化为字母标识
    def colToString(self,n):
        ss = ''
        m = n // 26
        r = n % 26
        if m:
            ss += self.colToString(m)
        ss += chr(65 + r)
        return ss


    #数字转化为字符
    def mtoString(self,input):
        try:
            # 因为使用float有一个例外是'NaN'
            if input == 'NaN':
                return 'NaN'
            float(input)
            return str(input)
        except ValueError:
            return input

    # 设置展示动作区域
    def show_detail(self,merge,num):
        if merge[0]:
            if merge[0]<0:
                rowE=[-1,[],[]]
                colE=[-1,[],[]]
                cellE=[-1,[],[]]
            elif merge[0]>1:
                row1=[]
                row2=[]
                for i in range(len(merge[1])):
                    if merge[1][i]:
                        if merge[1][i] <1:
                            row1.append(i)
                        elif merge[1][i] >1:
                            pass
                        else:
                            row2.append(i)
                rowE=[0,row1,row2]
                row1 = []
                row2 = []
                for i in range(len(merge[2])):
                    if merge[2][i]:
                        if merge[2][i] < 1:
                            row1.append(i)
                        elif merge[2][i]> 1:
                            pass
                        else:
                            row2.append(i)
                colE=[0,row1,row2]
                cellE=[0,merge[3]]
            else:
                rowE = [1, [], []]
                colE = [1, [], []]
                cellE = [1, [], []]
        else:
            rowE = [0, [], []]
            colE = [0, [], []]
            cellE = [0, [], []]
        self.row_button.clicked.connect(partial(self.diff_comment, 1, rowE,num))
        self.col_button.clicked.connect(partial(self.diff_comment, -1, colE,num))
        self.cell_button.clicked.connect(partial(self.diff_comment_cell,cellE,num))
        self.clearHighlight_button.clicked.connect(partial(self.show_sheets, num))
        self.diff_comment(1, rowE,num)

    #选择文件动作
    def select_file1(self):
        self.file1=QtGui.QFileDialog.getOpenFileName(self,u'选择文件','./','Excel(*.xls*);;All Files(*)')
        if self.file1:
            self.file_button1.setText(self.file1)
    def select_file2(self):
        self.file2=QtGui.QFileDialog.getOpenFileName(self,u'选择文件','./','Excel(*.xls*);;All Files(*)')
        if self.file2:
            self.file_button2.setText(self.file2)



    #因为diff把两个表合在了一起，所以有个对照
    def positionFind(self):
        def match(rowmerge ):
            rowpre = []
            rowpreR = []
            rowsuf = []
            rowsufR = []
            pre = 0
            suf = 0
            for i in range(len(rowmerge)):
                if rowmerge[i]:
                    if rowmerge[i] < 1:
                        rowpre.append(i)
                        rowpreR.append(pre)
                        rowsufR.append(-1)
                        pre += 1
                    elif rowmerge[i] > 1:
                        rowpre.append(i)
                        rowpreR.append(pre)
                        pre += 1
                        rowsuf.append(i)
                        rowsufR.append(suf)
                        suf += 1
                    else:
                        rowsuf.append(i)
                        rowsufR.append(suf)
                        rowpreR.append(-1)
                        suf += 1
                else:
                    rowpre.append(i)
                    rowpreR.append(pre)
                    rowsuf.append(i)
                    rowsufR.append(suf)
                    pre += 1
                    suf += 1
            return rowpre,rowsuf,rowpreR,rowsufR

        position=[]
        for itr in self.diffrst:
            rowpre, rowsuf, rowpreR, rowsufR=match(itr[3][1])
            colpre, colsuf, colpreR, colsufR=match(itr[3][2])
            #原表到展示表，展示表到原表
            position.append(([rowpre,rowsuf,colpre,colsuf],[rowpreR,rowsufR,colpreR,colsufR]))
        return position


    #显示行列动作的详情
    #输入（行或者列标识号，[表的改变，删除的行号，新增加的行号]）
    def diff_comment(self,sig,diff,num):
        self.row_button.setStyleSheet('background-color:rgb(255,255,255)')
        self.col_button.setStyleSheet('background-color:rgb(255,255,255)')
        self.cell_button.setStyleSheet('background-color:rgb(255,255,255)')
        self.detail.clearContents()
        self.detail.setRowCount(2)
        self.detail.setColumnCount(2)
        if sig>0:
            self.row_button.setStyleSheet('background-color:rgb(100,100,150)')
            hzh=u'行号'
        elif sig<0:
            self.col_button.setStyleSheet('background-color:rgb(100,100,150)')
            hzh =u'列号'
        # infolabel = QtGui.QTableWidgetItem(self.diffrst[num][0])
        # infolabel.setBackground(QtGui.QColor(100,100,150))
        # self.detail.setItem(0,0,infolabel)
        # lbadd=QtGui.QTableWidgetItem(u'新增')
        # self.detail.setItem(1,0,lbadd)
        # lbrm=QtGui.QTableWidgetItem(u'删除')
        # self.detail.setItem(2,0,lbrm)
        # lbhzh=QtGui.QTableWidgetItem(hzh)
        # self.detail.setItem(0, 1,lbhzh)
        vheadlable = [u'新增', u'删除']
        self.detail.setVerticalHeaderLabels(vheadlable)
        hheadlable=[]
        if diff[0] < 0:
            lbqb=QtGui.QTableWidgetItem(u'全部')
            self.detail.setItem(1, 0,lbqb)
        elif diff[0]>0:
            lbqb =QtGui.QTableWidgetItem(u'全部')
            self.detail.setItem(0, 0,lbqb)
        else:
            allrow=list(set(diff[1]).union(diff[2]))
            allrow.sort()
            self.detail.setColumnCount(len(allrow))
            for i in range(len(allrow)):
                # lbhzh = QtGui.QTableWidgetItem(hzh)
                # self.detail.setItem(0,i+1,lbhzh)
                hheadlable.append(hzh)
                if allrow[i] in diff[1]:
                    delbut=QtGui.QPushButton()
                    #delbut.setFixedSize(50,20)
                    if sig>0:
                        txt=self.position[num][1][0][allrow[i]]+1
                        delbut.setText(str(txt))
                    else:
                        txt = self.position[num][1][2][allrow[i]]+1
                        txt=self.colToString(txt-1)
                        delbut.setText(txt)
                    delbut.clicked.connect(partial(self.highlight,sig,allrow[i]))
                    self.detail.setCellWidget(1, i,delbut)
                if allrow[i] in diff[2]:
                    addbut=QtGui.QPushButton()
                    #addbut.setFixedSize(50,20)
                    if sig < 0:
                        txt = self.position[num][1][3][allrow[i]]+1
                        txt = self.colToString(txt-1)
                        addbut.setText(txt)
                    else:
                        txt = self.position[num][1][1][allrow[i]]+1
                        addbut.setText(str(txt))
                    addbut.clicked.connect(partial(self.highlight,sig,allrow[i],num))
                    self.detail.setCellWidget(0, i,addbut)
        self.detail.setHorizontalHeaderLabels(hheadlable)

    def diff_comment_cell(self,cell,num):
        self.row_button.setStyleSheet('background-color:rgb(255,255,255)')
        self.col_button.setStyleSheet('background-color:rgb(255,255,255)')
        self.cell_button.setStyleSheet('background-color:rgb(100, 100, 150)')
        self.detail.clearContents()
        self.detail.setRowCount(1)
        self.detail.setColumnCount(len(cell[1]))
        headlable=[u'改动']*len(cell[1])
        self.detail.setHorizontalHeaderLabels(headlable)
        self.detail.setVerticalHeaderLabels([u'单元格'])
        # infolabel = QtGui.QTableWidgetItem(self.diffrst[num][0])
        # infolabel.setBackground(QtGui.QColor(100, 100, 150))
        # self.detail.setItem(0, 0, infolabel)
        m=0
        for itr in cell[1]:
            # txt=itr.__str__().replace('), (',')->(')
            # txt=txt[1:-1]
            txt='('+str(itr[0][0]+1)+','+str(itr[0][1]+1)+')->('+str(itr[1][0]+1)+','+str(itr[1][1]+1)+')'
            inf = QtGui.QTableWidgetItem(txt)
            self.detail.setItem(0,m,inf)
            m+=1
    def highlight(self,sig,n,num):
        if sig>0:
            for i in range(self.tabWidget1.columnCount()):
                self.tabWidget1.item(n,i).setBackground(QtGui.QColor(100,100,150))
            for i in range(self.tabWidget2.columnCount()):
                self.tabWidget2.item(n,i).setBackground(QtGui.QColor(100,100,150))
            self.tabWidget1.verticalScrollBar().setSliderPosition(n)
            self.tabWidget2.verticalScrollBar().setSliderPosition(n)
        else:
            for i in range(self.tabWidget1.rowCount()):
                self.tabWidget1.item(i,n).setBackground(QtGui.QColor(100,100,150))
            for i in range(self.tabWidget2.rowCount()):
                self.tabWidget2.item(i,n).setBackground(QtGui.QColor(100,100,150))
            self.tabWidget1.horizontalScrollBar().setSliderPosition(n)
            self.tabWidget2.horizontalScrollBar().setSliderPosition(n)




def main():
    app = QtGui.QApplication(sys.argv)
    mainwindow=MainUi()
    mainwindow.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
