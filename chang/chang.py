#适用于老版本的病理查询系统
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
#T4
T4=1
T4X=[]
while T4<nrows:
    if "腹膜结节" in col[13][T4]:
        T4X.append(T4);
    elif "侵犯腹膜组织" in col[13][T4]:
        T4X.append(T4);
    elif "系膜结节" in col[13][T4]:
        T4X.append(T4);
    elif "穿透肠壁并累及" in col[13][T4]:
        T4X.append(T4);
    elif "T4" in col[13][T4]:
        T4X.append(T4);
    elif "全层" in col[13][T4]:
        print(T4)
        if "并突破浆膜层" in col[13][T4]:
            T4X.append(T4);
        elif "未突破浆膜层" in col[13][T4]:
            print(T4);        
        elif "突破浆膜层" in col[13][T4]:
            T4X.append(T4);
        elif "肿物向下累及齿状线"  in col[13][T4]:
            T4X.append(T4);
        elif "神经见癌侵犯" in col[13][T4]:
            T4X.append(T4);
        elif "切缘未见癌累及" in col[13][T4]:
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
    Tis=Tis+1
Tis=Tis+1
TisX
len(TisX)
#T4a

#T4b

