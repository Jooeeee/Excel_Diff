#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019/2/27 10:33
# @Author  : joestww@163.com
# @Site    :
# @File    : ExcelDiff.py
# @Software: PyCharm

import xlrd
from diff import Diff

class ExcelDiff:
    # print u"------欢迎使用Excel diff 工具------"
    DIFF_DELETE = -1
    DIFF_INSERT = 1
    DIFF_EQUAL = 0
    DIFF_EDIT = 2
    def main(self,file1path,file2path):
        xlsPre = xlrd.open_workbook(file1path.__str__())
        xlsSuf = xlrd.open_workbook(file2path.__str__())
        #存储tabel转换为list后的结果，按照行的格式，形如[(sheetname,[sheetPre],[sheetSuf]),]
        sheets=[]
        #存储diff后的结果，形如[(表名，改动情况，行的diff，列的diff)]，-1：删除，0：未改动，1：新增，2：改动
        #增删改的结构为，[(表的增删改，行的改动，列的改动，单元格的改动）]
        #行列的改动的结构为，[（增删，行号）,]
        #单元格的改动的结构为，[(行号，列号)，]
        diffrst=[]
        sheetnamesPre=xlsPre.sheet_names()
        sheetnamesSuf=xlsSuf.sheet_names()
        # print u"There are %s sheets."%(len(sheetnamesSuf))
        diffmain=Diff()
        for sheetname in sheetnamesPre:
            if sheetname in sheetnamesSuf:
                # print u"Start processing: ",sheetname
                rowsPre=self.sheet2ListByRow(xlsPre.sheet_by_name(sheetname))
                rowsSuf=self.sheet2ListByRow(xlsSuf.sheet_by_name(sheetname))
                sheets.append((sheetname,rowsPre,rowsSuf))
                if rowsPre==rowsSuf:
                    diffrst.append((sheetname, 0, rowsPre, rowsSuf))
                else:
                    # print u"old sheet has %s rows. " % len(rowsPre), u"new sheet has %s rows" % len(rowsSuf)
                    diffrows=diffmain.diff_main(rowsPre,rowsSuf,True)
                    # print u"finish rows diff"
                    colsPre=self.sheet2ListByCol(xlsPre.sheet_by_name(sheetname))
                    colsSuf=self.sheet2ListByCol(xlsSuf.sheet_by_name(sheetname))
                    # print u"old sheet has %s columns. " % len(colsSuf), u"new sheet has %s columns" % len(colsSuf)
                    diffcols=diffmain.diff_main(colsPre,colsSuf,True)
                    # print u"finish columns diff"
                    #diffmeg=self.diffMerge(diffrows,diffcols)
                    diffrst.append((sheetname,2,diffrows,diffcols))
            else:
                # print u"Start processing: ", sheetname
                rowsPre = self.sheet2ListByRow(xlsPre.sheet_by_name(sheetname))
                sheets.append((sheetname, rowsPre, []))
                diffrst.append((sheetname,-1,rowsPre,[]))
        for sheetname in sheetnamesSuf:
            if sheetname not in sheetnamesPre:
                # print u"Start processing: ", sheetname
                rowsPre = self.sheet2ListByRow(xlsSuf.sheet_by_name(sheetname))
                sheets.append((sheetname, [], rowsPre))
                diffrst.append((sheetname,1, [],rowsPre))
        diffmerge=self.diffMerge(diffrst)
        tables=self.makeTable(sheets, diffmerge)
        return tables

    #形成表格模式
    def makeTable(self,sheets, diffmerge):
        # print u"Convert data format for display"
        tables = []
        for i in range(len(diffmerge)):
            table1 = []
            table2 = []
            celledit=[]
            if diffmerge[i][0]:  # 表变了
                if diffmerge[i][0] < 0:  # 表被删除
                    for itr in sheets[i][1]:
                        table1.append([(-1, x) for x in itr])
                        table2.append([(-1, ' ') for x in range(len(itr))])
                elif diffmerge[i][0] > 1:  # 表改变

                    pm = -1
                    sm = -1
                    for itr in diffmerge[i][1]:  # 对行循环
                        if itr:  # 行变了
                            if itr < 0:  # 行删除
                                pm += 1
                                oneline1, oneline2, cellline = self.make_row(sheets[i][1][pm], [' ' for x in range(len(diffmerge[i][2]))],
                                                                             diffmerge[i][2])
                                for iitr in oneline1:
                                    iitr[0]=-1
                                for iitr in oneline2:
                                    iitr[0]=-1
                                table1.append(oneline1)
                                table2.append(oneline2)
                                #table1.append([(-1, x) for x in sheets[i][1][pm]])
                                #table2.append([(-1, ' ') for x in range(len(sheets[i][1][pm]))])
                            elif itr > 1:  # 这一行有了改动
                                pm += 1
                                sm += 1
                                oneline1,oneline2,cellline=self.make_row(sheets[i][1][pm],sheets[i][2][sm],diffmerge[i][2])
                                celledit.extend(((pm,x),(sm,y)) for x,y in cellline)
                                '''                                
                                pn = -1
                                sn = -1
                                oneline1 = []
                                oneline2 = []
                                for jtr in diffmerge[i][2]:  # 对列循环
                                    if jtr:  # 列改动了
                                        if jtr < 0:  # 列被删除
                                            pn += 1
                                            oneline1.append((-1, sheets[i][1][pm][pn]))
                                            oneline2.append((-1, ' '))
                                        elif jtr > 1:  # 列被修改
                                            pn += 1
                                            sn += 1

                                            if sheets[i][1][pm][pn] != sheets[i][2][sm][sn]:  # 单元格改动
                                                oneline1.append((2, sheets[i][1][pm][pn]))
                                                oneline2.append((2, sheets[i][2][sm][sn]))
                                                celledit.append((len(table1),len(oneline1)))
                                            else:  # 单元格没有被修改
                                                oneline1.append((0, sheets[i][1][pm][pn]))
                                                oneline2.append((0, sheets[i][2][sm][sn]))
                                        else:  # 列增加
                                            sn += 1
                                            oneline1.append((1, ' '))
                                            oneline2.append((1, sheets[i][2][pm][sn]))
                                    else:  # 列没有改动
                                        pn += 1
                                        sn += 1
                                        oneline1.append((0, sheets[i][1][pm][pn]))
                                        oneline2.append((0, sheets[i][2][sm][sn]))
                                '''
                                table1.append(oneline1)
                                table2.append(oneline2)

                            else:  # 行增加
                                sm += 1
                                oneline1, oneline2, cellline = self.make_row([' ' for x in range(len(diffmerge[i][2]))],
                                                                             sheets[i][2][sm],diffmerge[i][2])
                                for iitr in oneline1:
                                    iitr[0] = 1
                                for iitr in oneline2:
                                    iitr[0] = 1
                                table1.append(oneline1)
                                table2.append(oneline2)
                                #table1.append([(1, ' ') for x in range(len(sheets[i][2][sm]))])
                                #table2.append([(1, x) for x in sheets[i][2][sm]])
                        else:  # 行没变
                            pm += 1
                            sm += 1
                            table1.append([(0, x) for x in sheets[i][1][pm]])
                            table2.append([(0, x) for x in sheets[i][1][pm]])
                else:  # 表增加
                    for itr in sheets[i][2]:
                        table1.append([(1, ' ') for x in range(len(itr))])
                        table2.append([(1, x) for x in itr])
            else:  # 表没变
                for itr in sheets[i][1]:
                    table1.append([(0, x) for x in itr])
                    table2.append([(0, x) for x in itr])
            merge=(diffmerge[i][0],diffmerge[i][1],diffmerge[i][2],celledit)
            tables.append([sheets[i][0],table1, table2,merge])
        return tables

    def make_row(self,row1,row2,mergecol):
        pn = 0
        sn = 0
        oneline1 = []
        oneline2 = []
        celledit=[]
        for jtr in mergecol:  # 对列循环
            if jtr:  # 列改动了
                if jtr < 0:  # 列被删除
                    oneline1.append([-1, row1[pn]])
                    oneline2.append([-1, ' '])
                    pn += 1
                elif jtr > 1:  # 列被修改
                    if row1[pn] != row2[sn]:  # 单元格改动
                        oneline1.append([2, row1[pn]])
                        oneline2.append([2, row2[sn]])
                        celledit.append([pn,sn])
                    else:  # 单元格没有被修改
                        oneline1.append([0, row1[pn]])
                        oneline2.append([0, row2[sn]])
                    pn += 1
                    sn += 1
                else:  # 列增加
                    oneline1.append([1, ' '])
                    oneline2.append([1, row2[sn]])
                    sn += 1
            else:  # 列没有改动
                oneline1.append([0, row1[pn]])
                oneline2.append([0, row2[sn]])
                pn += 1
                sn += 1
        return  oneline1,oneline2,celledit

    #将diff结果聚合
    def diffMerge(self,diffrst):
        # print u"statistical changes by rows or columns"
        change = []
        for itr in diffrst:
            if itr[1]:
                if itr[1] < 1:
                    if len(itr[2]):
                        change.append((-1, range(len(itr[2])), range(len(itr[2][0]))))
                    else:
                        change.append((-1, [], []))
                elif itr[1]>1:
                    rowchange = []
                    for i in itr[2]:
                        rowchange.append(i[0])
                    colchange = []
                    for i in itr[3]:
                        colchange.append(i[0])
                    change.append((2, rowchange, colchange))
                else:
                    if len(itr[3]):
                        change.append((1, range(len(itr[3])), range(len(itr[3][0]))))
                    else:
                        change.append((1, [],[]))
            else:
                if len(itr[2]):
                    change.append((0, range(len(itr[2])), range(len(itr[2][0]))))
                else:
                    change.append((0, [], []))
        return change
    '''
    #按照行的形式比较两份表格（sheet）的不同
    def rowDiff(self,pre_sheet,tar_sheet):
        diff=Diff()
        return diff.diff_main(pre_sheet,tar_sheet)

    #按照列的形式比较两份表格（sheet）的不同
    def colDiff(self,pre_sheet,tar_sheet):
        return

    '''

    def sheet2ListByRow(self,sheet):
        rst=[]
        for i in range(sheet.nrows):
            rst.append(sheet.row_values(i))
        return rst

    def sheet2ListByCol(self, sheet):
        rst = []
        for i in range(sheet.ncols):
            rst.append(sheet.col_values(i))
        return rst