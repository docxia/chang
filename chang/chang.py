#适用于老版本的病理查询系统
#!/usr/bin/env python
# -*- coding: utf-8 -*-
import xlrd
import xlwt
from xlutils.copy import copy
from xlwt import Style
import openpyxl    # 只能用于xlsx文件
from win32com.client import Dispatch     
import win32com.client as win32
workBook = xlrd.open_workbook("D:/chang/肠癌.xls")
workBook2= copy(workBook)  
print(workBook)
# 1.获取sheet的名字
# 1.1 获取所有sheet的名字(list类型)
sheetname= workBook.sheet_names()
print(sheetname)
# 1.2 按索引号获取sheet的名字（string类型）
sheet1Name = workBook.sheet_names()[0];
print(sheet1Name);
# 1.3 编写一个写入文档
# 2. 获取sheet内容
sheet1 = workBook.sheets()[0]
sheet2 = workBook2.get_sheet(0)
print(sheet1)
print(sheet2)
#   获取整行和整列的值（数组）
a=sheet1.row_values(0)
b=sheet1.col_values(0)
print(a,b)
# 3. 获取行数和列数
nrows = sheet1.nrows #行数
ncols = sheet1.ncols #列数
print(nrows,ncols)
# 4. 获取整行和整列的值（数组）
row1 = sheet1.row_values(0);   # 获取第1行内容
print(row1)
colnamber=0
col=[None]*ncols #设置连续变量名
while colnamber<ncols:
    col[colnamber+1] =  sheet1.col_values(colnamber);
    print(col[colnamber+1])
    colnamber=colnamber+1
col[3]
len(col[3])
col[3][1]
i=1  #标记门诊病理
menzhen=[]
while i < nrows:
    if  col[3][i]=="":    #此处只能用""，不能用None来表示
        menzhen.append(i);
    i=i+1
print(menzhen)
menzhennamber=len(menzhen)
print("门诊病例数为：",menzhennamber)
#确认组织来源是结直肠
i1=1
feijiezhichangai=[]
while i1<nrows:    
    if "结肠" in col[11][i1]:
        print(col[11][i1])
    elif "直肠" in col[11][i1]:
        print(col[11][i1])
    elif "腹痛查因" in col[11][i1]:
        print(col[11][i1])
    elif "盆腔包块" in col[11][i1]:
        print(col[11][i1])
    elif "卵巢癌" in col[11][i1]:
        feijiezhichangai.append(i1)
    elif "卵巢肿" in col[11][i1]:
        feijiezhichangai.append(i1)
    elif "胃癌" in col[11][i1]:
        feijiezhichangai.append(i1)
    elif "胃底贲门肿瘤" in col[11][i1]:
        feijiezhichangai.append(i1)
    elif "胃肿瘤" in col[11][i1]:
        feijiezhichangai.append(i1)
    elif "子宫内膜" in col[11][i1]:
        feijiezhichangai.append(i1)
    elif "胰" in col[11][i1]:
        feijiezhichangai.append(i1) 
    elif "十二指肠" in col[11][i1]:
        feijiezhichangai.append(i1)
    elif "阑尾" in col[11][i1]:
        feijiezhichangai.append(i1)
    elif "结肠" in col[10][i1]:
        print(col[11][i1])
    elif "直肠" in col[10][i1]:
        print(col[11][i1])
    else:
        feijiezhichangai.append(i1)
    i1=i1+1 
feijiezhichangai
#如果还有其他的筛选条件，可在此处添加
'''
i1=1 
 while i1<nrows: 
     if "" in col[][i1]:
'''
len(feijiezhichangai)
paichu=[]
paichu=feijiezhichangai+menzhen
paichu=list(set(paichu))#去重复
paichu.sort(reverse = False)#排列，从小到大
paichu
len(paichu)
workBook2.save("D:/chang/new.xls")
workBook2.save("D:/chang/new.xlsx")
#接下来是删除任务
excel= win32.Dispatch('Excel.Application')
vb=excel.Workbooks.Open('D:/chang/new.xls')
sht = vb.Worksheets(1)
arryx=paichu[0]
arryx
xx=0
for arry in paichu:
     
     if arry > arryx:
         arry=arry-xx     #删除不需要的样本，一定要加入循环数
     xx=xx+1
     arryy=arry+1
     sht.Rows(arryy).Delete()
     print(xx)
     print(arryy)
vb.Save() #保存修改的文件
vb.Close() # 做完后一定要关闭
#################################
####老文件规范文本所必备的#######
#################################
####新系统只需要使用下面的文本###
#################################
####必须对应表格名称#############
#TNM分期中的T分期
workBook = xlrd.open_workbook("D:/chang/new.xls")
workBook2= copy(workBook)  
print(workBook)
# 1.获取sheet的名字
# 1.1 获取所有sheet的名字(list类型)
sheetname= workBook.sheet_names()
print(sheetname)
# 1.2 按索引号获取sheet的名字（string类型）
sheet1Name = workBook.sheet_names()[0];
print(sheet1Name);
# 1.3 编写一个写入文档
# 2. 获取sheet内容
sheet1 = workBook.sheets()[0]
sheet2 = workBook2.get_sheet(0)
print(sheet1)
print(sheet2)
#   获取整行和整列的值（数组）
a=sheet1.row_values(0)
b=sheet1.col_values(0)
print(a,b)
# 3. 获取行数和列数
nrows = sheet1.nrows #行数
ncols = sheet1.ncols #列数
print(nrows,ncols)
# 4. 获取整行和整列的值（数组）
row1 = sheet1.row_values(0);   # 获取第1行内容
print(row1)
colnamber=0
col=[None]*ncols #设置连续变量名
while colnamber<ncols:
    col[colnamber+1] =  sheet1.col_values(colnamber);
    print(col[colnamber+1])
    colnamber=colnamber+1   #设置连续变量名结束
col[13] #查看第十三行变量
#T浸润深度的分期
#T4
T4=1
T4X=[]
while T4<nrows:
    if "T4" in col[13][T4]:
        T4X.append(T4);
    elif "腹膜结节" in col[15][T4]:
        T4X.append(T4);
    elif "侵犯腹膜组织" in col[15][T4]:
        T4X.append(T4);
    elif "系膜结节" in col[15][T4]:
        T4X.append(T4);
    elif "穿透肠壁并累及" in col[15][T4]:
        T4X.append(T4);
    elif "）见癌侵犯" in col[15][T4]:
        T4X.append(T4);
    elif "）浆液性囊肿，并见低分化腺癌侵犯" in col[15][T4]:
        T4X.append(T4);    
    if "全层" in col[13][T4]:
        print(T4)
        if "并突破浆膜层" in col[15][T4]:
            T4X.append(T4);
        elif "未突破浆膜层" in col[15][T4]:
            print(T4);        
        elif "突破浆膜层" in col[15][T4]:
            T4X.append(T4);
        elif "肿物向下累及齿状线"  in col[15][T4]:
            T4X.append(T4);
        elif "神经见癌侵犯" in col[15][T4]:
            T4X.append(T4);
        elif "切缘未见癌累及" in col[15][T4]:
            print(T4);
    T4=T4+1
T4X
len(T3X)
#T3
T3=1
T3X=[]
while T3<nrows:
    if T3 in T4X:
        print(T3)
    elif "T3N" in col[15][T3]:
        T3X.append(T3)
    elif "浸润至浆膜下层" in col[15][T3]:
        T3X.append(T3)
    elif "浸润至浆膜层" in col[15][T3]:
        T3X.append(T3)
    elif "浸润肠壁全层" in col[15][T3]:
        T3X.append(T3)
    elif "浸润肠壁全层至浆膜下层" in col[15][T3]:
        T3X.append(T3)    
    elif "癌组织浸润至肠壁浆膜层" in col[15][T3]:
        T3X.append(T3)    
    T3=T3+1
T3X
len(T3X)
#T2
T2=1
T2X=[]
while T2<nrows:
    if T2 in T4X:
        print(T2)
    elif T2 in T3X:
        print(T2)
    elif "浸润至浅肌层" in col[15][T2]:
        T2X.append(T2)
    elif "浸润至深肌层" in col[15][T2]:
        T2X.append(T2)
    elif "浸润肠壁深肌层" in col[15][T2]:
        T2X.append(T2)
    elif "浸润肠壁深肌层" in col[15][T2]:
        T2X.append(T2)
    elif "浸润至外纵肌层" in col[15][T2]:
        T2X.append(T2)
    elif "浸润至内环肌层" in col[15][T2]:
        T2X.append(T2)
    elif "浸润至肌层" in col[15][T2]:
        T2X.append(T2)
    elif "浸润肠壁浅肌层" in col[15][T2]:
        T2X.append(T2)
    T2=T2+1 
T2X
len(T2X)
#T1
T1=1
T1X=[]
while T1<nrows:
    if T1 in T4X:
        print(T1)
    elif T1 in T3X:
        print(T1)
    elif T1 in T2X:
        print(T1)
    elif "T1N" in col[15][T1]:
        T1X.append(T1) 
    elif "浸润至粘膜下层" in col[15][T1]:
        T1X.append(T1) 
    elif "癌组织主要位于粘膜层" in col[15][T1]:
        T1X.append(T1)  
    elif "浸润至粘膜下层" in col[15][T1]:
        T1X.append(T1)
    T1=T1+1
T1X
len(T1X)  
#Tis
Tis=1
TisX=[]
while Tis<nrows:
    if Tis in T4X:
        print(Tis)
    elif Tis in T3X:
        print(Tis)
    elif Tis in T2X:
        print(Tis)
    elif Tis in T1X:
        print(Tis)
    elif "局限于粘膜层" in col[15][Tis]:
        TisX.append(Tis)
    elif "局限于粘膜下层" in col[15][Tis]:
        TisX.append(Tis)
    elif "粘膜内癌" in col[15][Tis]:
        TisX.append(Tis)
    elif "粘液腺癌" in col[15][Tis]:
        TisX.append(Tis)
    Tis=Tis+1
TisX
len(TisX)
#T4b
T4bX=[]
T4b=1
while T4b<nrows:
    if "T4b" in col[15][T4b]:
        T4bX.append(T4b)
    elif "局部浸润至" in col[15][T4b]:
        T4bX.append(T4b)
    elif "侵至" in col[15][T4b]:
        T4bX.append(T4b)
    elif "系膜内见癌" in col[15][T4b]:
        T4bX.append(T4b)
    elif "全层并部分累" in col[15][T4b]:
        T4bX.append(T4b)
    T4b=T4b+1 
len(T4bX)
#T4a
T4aX=list(set(T4X).difference(set(T4bX))) # 取差集（前者为主集合，后者为排除的集合）
len(T4aX)

#N肿瘤淋巴结的分期
#淋巴结计数模块
import re 
N=0
NX=[]
while N<nrows:
    ff1=re.findall(r'\d+/',col[15][N])
    k=len(ff1)
    i=0
    ff2=[]
    while i<k:
        x=re.findall(r'\d+',ff1[i])
        ff2.append(x[0])
        i+=1
    ff2
    ff2=list(map(int,ff2))
    ele=0
    total=0
    while(ele<len(ff2)):
        total=total+ff2[ele]
        ele+=1
    total
    NX.append(total)
    N+=1
NX #为每个样本的淋巴转移之和的集
len(NX)
NX[1]
col[15][1]
#副程序
tx=1
#N2b
N2b=[]
while tx<len(NX):
    if "N2b" in col[15][tx]:
        N2b.append(tx)
    if (NX[tx] > 6):
        N2b.append(tx)
    tx=tx+1
tx=1
#N2a
N2a=[]
while tx<nrows:
    z=NX[tx]
    if "N2a" in col[15][tx]:
        N2a.append(tx)
    elif z in range(4, 7):
        N2a.append(tx)
    tx+=1
#N1a
N1a=[]
tx=1
while t<nrows:
    z=NX[tx]
    if "N1a" in col[15][tx]:
        N1a.append(tx)
    elif z==1:
        N1a.append(tx)
    tx+=1
#N1c
N1c=[]
tx=1
while tx<nrows:
    z=NX[tx]
    if "N1c" in col[15][tx]:
        N1c.append(tx)
    elif z==1:
        print(z)
        if "直肠周围软组织内卫星肿瘤结节" in col[15][tx]:
            N1c.append(tx)       
    tx+=1
#N1b
tx=1
N1b=[]
while tx<nrows:
    z=NX[tx]
    if "N1b" in col[15][tx]:
        N1b.append(tx)
    elif z==1:
        print(z)
        if z not in N1c:
            N1b.append(tx)
    tx+=1
#N0
N0=[]
tx=1
while tx<nrows:
    z=NX[tx]
    if "N0" in col[15][tx]:
        N1a.append(tx)
    elif tx==0:
        N1a.append(tx)
    tx+=1 

#M分期
#M1开始
#M1
M1cX=[]
tx=1
while tx<nrows:
    if "M1c" in col[15][tx]:
        M1cX.append(tx)
    elif "系膜淋巴结见癌" in col[15][tx]:
        M1cX.append(tx)
    elif "大网膜）见中"  in col[15][tx]:
        M1cX.append(tx)
    elif "大网膜见癌"  in col[15][tx]:
        M1cX.append(tx)
    elif "大网膜）见多个癌结节"  in col[15][tx]:
        M1cX.append(tx)
    elif "小肠系膜内见" in col[15][tx]:
        M1cX.append(tx)
    elif "大肠系膜内见" in col[15][tx]:
        M1cX.append(tx)
    tx+=1
#M1b
#以累计值的方式算评价危险系数

M1bX=[]
tx=1
M1aX=[]
while tx<nrows:
    if tx in M1cX:
        print(tx)
    elif "M1b" in col[15][tx]:
        M1bX.append(tx)
    elif "M1a" in col[15][tx]:
        M1aX.append(tx)
        k=0
        if "癌组织转移至（" in col[15][tx]:
            k+=2
        if "均见癌" in col[15][tx]:
            k+=2   
        if ")均见" in col[15][tx]:
            k+=2
        if "）见浆液性腺癌" in col[15][tx]:
            k+=2
        if "）见癌累及" in col[15][tx]:
            k+=2
        if "）见癌转移" in col[15][tx]:
            k+=2
        if "卵巢）中分" in col[15][tx]:
            k+=1
        if "小肠）中分" in col[15][tx]:
            k+=1
        if "盆壁肿物）中分" in col[15][tx]:
            k+=1
        if ("结肠" in col[15][tx]) & ("直肠" in col[15][tx]):
            k+=1
        if "浸润至宫颈" in col[15][tx]:
            k+=1
        if "）转移性腺癌" in col[15][tx]:
            k+=2
        if "子宫浆膜面见癌" in col[15][tx]:
            k+=1
        if "及宫颈管见癌" in col[15][tx]:
            k+=2
        if "）送检组织均见癌" in col[15][tx]:
            k+=2
        if "卵巢内见癌累" in col[15][tx]:
            k+=1
        if "乙状结肠与子宫间肿物）中分化腺癌" in col[15][tx]:
            k+=1
        if "子宫）中分化" in col[15][tx]:
            k+=1
        if "子宫旁结节）送检组织见癌" in col[15][tx]:
            k+=1
        if "小肠）缝线处见癌浸润" in col[15][tx]:
            k+=1
        
        if "内见癌转移" in col[15][tx]:
            k+=1    
        if "段)转移性" in col[15][tx]:
            k+=1  
        if "肝转移腺癌" in col[15][tx]:
            k+=1
        if "（肝）见癌转移" in col[15][tx]:
            k+=1
        if "肝曲)中分化" in col[15][tx]:
            k+=1
        if "（肝曲、直肠）中分化腺癌" in col[15][tx]:
            k+=1
        if "肝脏肿物）中分" in col[15][tx]:
            k+=1
        if "肝曲)低分化" in col[15][tx]:
            k+=1 
        if "肝肿物）转移性" in col[15][tx]:
            k+=1
        if "段）符合" in col[15][tx]:
            k+=1
        if "及肝曲结肠）中分化腺癌" in col[15][tx]:
            k+=1
        if "肝脏结节）见腺癌" in col[15][tx]:
            k+=1
        if "肝肿物）中分化" in col[15][tx]:
            k+=1
        if k==1:
            M1aX.append(tx)
        if k>1:
            M1bX.append(tx)
    tx+=1
M0X=[]
tx=1
while tx<nrows:
    if tx in M1cX:
        M0X.append(tx)
    tx=tx+1
#癌栓有无


# 写入
workBook = xlrd.open_workbook("D:/chang/new.xls")
workBook2= copy(workBook)
sheet1 = workBook.sheets()[0]
sheet2 = workBook2.get_sheet(0)
#写入表格
sheet2.write(0,17,"T")
sheet2.write(0,18,"N")
sheet2.write(0,19,"M")
sheet2.write(0,20,"癌栓有无")
sheet2.write(0,21,"神经侵犯")
#T写入
i=1
while i<nrows:
    if i in T4bX:
        sheet2.write(i,17,"T4b")
    elif i in T4aX:
        sheet2.write(i,17,"T4a")
    elif i in T3X:
        sheet2.write(i,17,"T3")
    elif i in T2X:
        sheet2.write(i,17,"T2")
    elif i in T1X:
        sheet2.write(i,17,"T1")
    i+=1
#N写入
i=1
while i<nrows:
    if i in N2b:
        sheet2.write(i,18,"N2b")
    elif i in N2a:
        sheet2.write(i,18,"N2a")
    elif i in N1a:
        sheet2.write(i,18,"N1a")
    elif i in N1b:
        sheet2.write(i,18,"N1b") 
    elif i in N1c:
        sheet2.write(i,18,"N1c")
    elif i in N0:
        sheet2.write(i,18,"N0")
    i+=1
#M写入
i=1
while i<nrows:
    if i in M1bX:
        sheet2.write(i,19,"M1b")
    elif i in M1cX:
        sheet2.write(i,19,"M1c")
    elif i in M1aX:
        sheet2.write(i,19,"M1a")
    i+=1    
#癌栓转移有无
col15=sheet1.col_values(14)
#癌栓是否有
i2=1
while i2<nrows:
    aishuan=col15[i2]
    if "癌栓" not in aishuan:
        sheet2.write(i2,20,"不明")
    elif "未见癌栓" in aishuan:
        sheet2.write(i2,20,"无")
    elif "未见明确癌栓" in aishuan:
        sheet2.write(i2,20,"无")
    else:
        sheet2.write(i2,20,"有")    
    i2=i2+1
#神经侵犯
i4=1
while i4<nrows:
    shenjin=col15[i4]
    if "神经未见癌侵犯" in shenjin:
        sheet2.write(i4,21,"否")
    elif "未见神经侵犯" in shenjin:
        sheet2.write(i4,21,"否")
    elif "神经未见明确" in shenjin:
        sheet2.write(i4,21,"否")
    elif "神经未查见" in shenjin:
        sheet2.write(i4,21,"否")
    elif "肿瘤侵犯神经" in shenjin:
        sheet2.write(i4,21,"是")
    else:
        sheet2.write(i4,21,"不明")
    i4=i4+1
workBook2.save("D:/chang/newd.xls")
    
    
    
