#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019/3/1 10:45
# @Author  : joestww@163.com
# @Site    :
# @File    : diff.py
# @Software: PyCharm


class Diff:
    #  DIFF FUNCTIONS

    # The data structure representing a diff is an array of tuples:
    # [(DIFF_DELETE, "Hello"), (DIFF_INSERT, "Goodbye"), (DIFF_EQUAL, " world.")]
    # which means: delete "Hello", add "Goodbye" and keep " world."

    DIFF_DELETE = -1
    DIFF_INSERT = 1
    DIFF_EQUAL = 0
    DIFF_EDIT = 2

    def diff_main(self,list1,list2,linemode):
        #错误输入
        if list1==None or list2==None:
            raise ValueError("Null inputs. (diff_main)")

        #完全相同
        if list1==list2:
            if list1:
                return [(self.DIFF_EQUAL,list1)]
            return []

        #存储结果
        diffPre=[]
        diffSuf=[]
        #公共前缀
        commonPrefix=self.diffCommonPrefix(list1,list2)
        if commonPrefix:
            for itr in commonPrefix:
                diffPre.append((self.DIFF_EQUAL,itr))
            list1=list1[len(commonPrefix):]
            list2=list2[len(commonPrefix):]
            if not list1:
                for itr in list2:
                    diffPre.append((self.DIFF_INSERT, itr))
                return diffPre
            if not list2:
                for itr in list1:
                    diffPre.append( (self.DIFF_DELETE, itr))
                return diffPre
        #公共后缀
        commonSuffix=self.diffCommonSuffix(list1,list2)
        if commonSuffix:
            for itr in commonSuffix:
                diffSuf.append((self.DIFF_EQUAL, itr))
            list1=list1[:-len(commonSuffix)]
            list2=list2[:-len(commonSuffix)]
            if not list1:
                for itr in list2:
                    diffPre.append((self.DIFF_INSERT,itr))
                diffPre.extend(diffSuf)
                return diffPre
                '''
                if commonPrefix:
                    return [(self.DIFF_EQUAL,commonPrefix),(self.DIFF_INSERT,list2),(self.DIFF_EQUAL, commonSuffix)]
                else:
                    return [(self.DIFF_INSERT, list2), (self.DIFF_EQUAL, commonSuffix)]
                '''
            if not list2:
                for itr in list1:
                    diffPre.append((self.DIFF_DELETE, itr))
                diffPre.extend(diffSuf)
                return diffPre
                '''
                if commonPrefix:
                    return [(self.DIFF_EQUAL, commonPrefix), (self.DIFF_DELETE, list1),(self.DIFF_EQUAL, commonSuffix)]
                else:
                    return [(self.DIFF_DELETE, list1), (self.DIFF_EQUAL, commonSuffix)]
                '''
        #包含关系
        singular=self.diffSingular(list1,list2)
        if singular:
            diffPre.extend(singular)
            diffPre.extend(diffSuf)
            return diffPre

        #两步编辑
        twoEdit=self.diffTwoEdit(list1,list2)
        if twoEdit:
            (list1Pre, list1Suf, list2Pre, list2Suf, commonInner)=twoEdit
            listPre=self.diffCompute(list1Pre,list2Pre,linemode)
            listSuf=self.diffCompute(list1Suf,list2Suf,linemode)
            diffPre.extend(listPre)
            diffPre.extend([(self.DIFF_EQUAL,itrt) for itrt in commonInner])
            diffPre.extend(listSuf)
            diffPre.extend(diffSuf)
            return diffPre

        #没有任何可以取巧的办法后，只能暴力解决
        diffcomp=self.diffCompute(list1,list2,linemode)
        diffPre.extend(diffcomp)
        diffPre.extend(diffSuf)
        return diffPre




    '''
    def diffEqual(self,list1,list2):
        if list1==list2:
            if list1:
                return [self.DIFF_EQUAL,list1]
            return []
    '''
    #Common Prefix/Suffix
    def diffCommonPrefix(self,list1,list2):
        commonPrefix=[]
        cnt=min(len(list1),len(list2))
        for i in range(cnt):
            if list1[i]==list2[i]:
                commonPrefix.append(list1[i])
            else:
                break
        return commonPrefix

    def diffCommonSuffix(self,list1,list2):
        commonSuffix=[]
        cnt=min(len(list1),len(list2))
        for i in range(1,cnt+1):
            if list1[-i]==list2[-i]:
                commonSuffix.append(list1[-i])
            else:
                break
        commonSuffix.reverse()
        return commonSuffix

    def diffSingular(self,list1,list2):
        diffrst=[]
        if len(list1)>=len(list2):
            pos=self.listFind(list1,list2)
            if pos!=-1:
                diffrst.extend([(self.DIFF_DELETE, itrt )for itrt in list1[:pos]])
                diffrst.extend([(self.DIFF_EQUAL, itrt) for itrt in list2])
                diffrst.extend([(self.DIFF_DELETE, itrt) for itrt in list1[pos+len(list2):]])
                return diffrst
        else:
            pos = self.listFind(list2, list1)
            if pos != -1:
                diffrst.extend([(self.DIFF_INSERT, itrt) for itrt in list2[:pos]])
                diffrst.extend([(self.DIFF_EQUAL, itrt) for itrt in list1])
                diffrst.extend([(self.DIFF_INSERT, itrt) for itrt in list2[pos + len(list1) :]])
                return diffrst

    def diffTwoEdit(self,list1,list2):
        l1=len(list1)
        l2=len(list2)
        if l1>=l2:
            hmp=self.diffHalfMatch(list1,list2,(l1+3)//4)
            if hmp:
                return hmp
            else:
                hmp=self.diffHalfMatch(list1,list2,(l1+1)//2)
                if hmp:
                    return hmp
        else:
            hmp=self.diffHalfMatch(list2,list1,(l2+3)//4)
            if hmp:
                (p1,p2,t1,t2,cm)=hmp
                return (t1,t2,p1,p2,cm)
            else:
                hmp = self.diffHalfMatch(list2, list1, (l2 + 1) // 2)
                if hmp:
                    (p1, p2, t1, t2, cm) = hmp
                    return (t1, t2, p1, p2, cm)


    #每个list中元素是list，也就是行比较模式
    def diffCompute(self,list1,list2,linemode):
        diffrst=[]
        diffline=self.diffLCS(list1,list2)
        if not linemode:
            return diffline
        editPre=[]
        editSuf=[]
        for itr in diffline:
            if itr[0]:
                if itr[0]>0:
                    editSuf.append(itr[1])
                else:
                    editPre.append(itr[1])
            else:
                if editPre and editSuf:
                    editmatch=self.diffEditMatch(editPre,editSuf)
                    # put result into diffrst
                    diffrst.extend(editmatch)
                else:
                    for ei in editPre:
                        diffrst.append((-1,ei))
                    for ei in editSuf:
                        diffrst.append((1, ei))
                editPre = []
                editSuf = []
                diffrst.append(itr)
        if editPre and editSuf:
            editmatch = self.diffEditMatch(editPre, editSuf)
            diffrst.extend(editmatch)
        else:
            for ei in editPre:
                diffrst.append((-1, ei))
            for ei in editSuf:
                diffrst.append((1, ei))
        return diffrst

    def diffEditMatch(self,editPre,editSuf):
        # print u'match by rows or columns'
        diffcost=[]
        dpall=[[0]]*(len(editPre)*len(editSuf))
        hasmatched = set()
        for itrP in range(len(editPre)):
            linecost=[]
            isbreak=False
            for itrS in range(len(editSuf)):
                if itrS in hasmatched:
                    linecost.append(1)
                    continue
                dp=self.diff_main(editPre[itrP],editSuf[itrS],False)
                dpall[itrP*len(editSuf)+itrS]=dp
                dpcost=1- float(self.dpCost(dp))/(len(editPre[itrP]))
                if dpcost<0.2:
                    # print u'0.8 match',(itrP,itrS)
                    diffcost.append((itrP,itrS))
                    hasmatched.add(itrS)
                    isbreak=True
                    break
                linecost.append(dpcost)
            if isbreak:
                continue
            linemincost=min(linecost)
            if linemincost<0.5:
                for i in range(len(linecost)):
                    if linemincost==linecost[i]:
                        # print u'best matche: ',(itrP,i)
                        diffcost.append((itrP,i))
                        hasmatched.add(i)
                        break
        lineMatchPerfect=[]

        for itr in range(len(diffcost)):
            linematch = []
            linematch.append(diffcost[itr])
            pos=itr
            for nextitr in range(itr+1,len(diffcost)):
                if diffcost[nextitr][0]>diffcost[pos][0] and diffcost[nextitr][1]>diffcost[pos][1]:
                    linematch.append(diffcost[nextitr])
                    pos = nextitr
            if len(linematch)>len(lineMatchPerfect):
                lineMatchPerfect=linematch


        diffrst=[]
        st=0
        sd=0

        for match in lineMatchPerfect:
            while st<match[0]:
                diffrst.append((self.DIFF_DELETE,editPre[st]))
                st+=1
            while sd<match[1]:
                diffrst.append((self.DIFF_INSERT,editSuf[sd]))
                sd+=1
            diffrst.append((self.DIFF_EDIT,dpall[match[0]*len(editSuf)+match[1]]))
            st += 1
            sd += 1

        while st<len(editPre):
            diffrst.append((self.DIFF_DELETE,editPre[st]))
            st+=1
        while sd<len(editSuf):
            diffrst.append((self.DIFF_INSERT,editSuf[sd]))
            sd+=1
        return diffrst

    ##LCS算法
    def diffLCS(self,list1,list2):
        len1 = len(list1) + 1
        len2 = len(list2) + 1
        dp = [(0, 0)] * (len1 * len2)
        for i in range(1, len1):
            dp[i * len2] = (-1, 0)
        for i in range(1, len2):
            dp[i] = (1, 0)
        for i in range(1, len1):
            for j in range(1, len2):
                if list1[i - 1] == list2[j - 1]:
                    dp[i * len2 + j] = (self.DIFF_EQUAL, dp[(i - 1) * len2 + j - 1][1] + 1)
                else:
                    if dp[(i - 1) * len2 + j][1] > dp[i * len2 + j - 1][1]:
                        dp[i * len2 + j] = (self.DIFF_DELETE, dp[(i - 1) * len2 + j][1])
                    else:
                        dp[i * len2 + j] = (self.DIFF_INSERT, dp[i * len2 + j - 1][1])
        si = len1 - 1
        sj = len2 - 1
        rst = []
        while si or sj:
            ck = dp[si * len2 + sj]
            if ck[0] > 0:
                rst.append((1,list2[sj - 1]))
                sj -= 1
            elif ck[0]<0:
                rst.append((-1,list1[si - 1]))
                si -= 1
            else:
                rst.append((0,list1[si - 1]))
                si -= 1
                sj -= 1
        rst.reverse()
        return rst

        '''以下部分是讲同类项合并，后来改用每个元素都配上改动的形式，所以舍弃
        si = len1 - 1
        sj = len2 - 1
        rst = []
        sig = dp[len1 * len2 - 1][0]
        li = []
        while si or sj:
            ck = dp[si * len2 + sj]
            if ck[0] > 0:
                if sig != ck[0]:
                    li.reverse()
                    rst.append((sig, li))
                    sig = ck[0]
                    li = []
                li.append(list2[sj - 1])
                sj -= 1
            elif ck[0] < 0:
                if sig != ck[0]:
                    li.reverse()
                    rst.append((sig, li))
                    sig = ck[0]
                    li = []
                li.append(list1[si - 1])
                si -= 1
            else:
                if sig != ck[0]:
                    li.reverse()
                    rst.append((sig, li))
                    sig = ck[0]
                    li = []
                li.append(list1[si - 1])
                si -= 1
                sj -= 1
        if li:
            li.reverse()
            rst.append((sig, li))
        rst.reverse()
        return rst
        '''

    def dpCost(self,dp):
        cost=0
        for itr in dp:
            if not itr[0]:
                cost+=1
        return cost

    #SED算法
    '''
    def diffCompute(self,list1,list2):
        len1=len(list1)
        len2=len(list2)
        dp=[0]*(len1*len2)
        for i in range(len1):
            dp[i*len2]=i
        for i in range(len2):
            dp[i]=i
        for i in range(1,len1):
            for j in range(1,len2):
                dp[i*len2+j]=dpComp(i,j)
        def fEdit(i,j):
            if list1[i]==list2[j]:
                return 0
            else:
                return 2
        def dpComp(i,j):
            lf=dp[(i-1)*len2+j]+1
            tp=dp[i*len2+j-1]+1
            sp=dp[(i-1)*len2+j-1]+fEdit(i,j)
            return min(lf,tp,sp)
    '''

    def diffHalfMatch(self,longlist,shortlist,startl1):
        l1=len(longlist)
        l2=len(shortlist)
        if l1>2*l2:
            return None
        endl1=startl1+l1//4
        pos=self.listFind(shortlist,longlist[startl1:endl1])
        posend=pos+endl1-startl1
        if pos!=-1:
            while startl1>0 and pos>0 and longlist[startl1-1]==shortlist[pos-1]:
                startl1-=1
                pos-=1
            while endl1<l1 and posend<l2 and longlist[endl1]==shortlist[posend]:
                endl1+=1
                posend+=1
            if (endl1-startl1)*2>l1:
                return (longlist[:startl1],longlist[endl1:],shortlist[:pos],
                    shortlist[posend:],longlist[startl1:endl1])
        else:
            return None


    def listFind(self,list1,list2):
        l1=len(list1)
        l2=len(list2)
        for i in range(l1-l2+1):
            sig=True
            for j in range(l2):
                if list2[j]!=list1[i+j]:
                    sig=False
                    break
            if sig:
                return i
        return -1