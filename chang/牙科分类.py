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
workBook = xlrd.open_workbook("D:/新建文件夹/口腔科.xls")
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
while colnamber<ncols-1:
    col[colnamber+1] =  sheet1.col_values(colnamber);
    print(col[colnamber+1])
    colnamber=colnamber+1
col[3]
sex=col[4]
name=col[17]
old=col[5]
i=1
man=[]
woman=[]
while i<nrows:
    if sex[i]=="男":
        man.append(i)
    if sex[i]=="女":
        woman.append(i)
    i+=1
#连续创造多个空集
for e in range(1,7):
    exec( 'x%s = []' % e)
i=1
while i<nrows:
    if "鳞状细胞癌" in name[i]:
        x1.append(i)
    elif "Warthin瘤" in name[i]:
        x2.append(i)
    elif "腺样囊性癌" in name[i]:
        x3.append(i) 
    elif "成釉细胞瘤" in name[i]:
        x4.append(i)
    elif "含牙囊肿" in name[i]:
        x5.append(i)
    elif "鳃裂囊肿" in name[i]:
        x6.append(i)
    i+=1
z1=[]
z2=[]
z3=[]
z4=[]
z5=[]
z6=[]
i=1
xx=0
while i<nrows:
    if old[i] != "":
        xx=int(old[i])
    if xx>0 & xx<7:
        z1.append(i)
    elif xx>6 & xx<13:
        z2.append(i)    
    elif xx>12 & xx<18:
        z3.append(i)
    elif xx>19 & xx<46:
        z4.append(i)
    elif xx>45 & xx<70:
        z5.append(i) 
    elif xx>69:
        z6.append(i)
    i+=1
workBook2.save("D:/新建文件夹/new.xls")
workBook2.save("D:/新建文件夹/new.xlsx")
workBook = xlrd.open_workbook("D:/新建文件夹/new.xls")
workBook2= copy(workBook)
sheet1 = workBook.sheets()[0]
sheet2 = workBook2.get_sheet(0)
sheet2.write(0,23,"疾病")
sheet2.write(0,24,"年龄段")

#图形绘制
# -*- coding: utf-8 -*-
#显示所有疾病的数量分布
import matplotlib.pyplot as plt 
plt.figure(figsize=(10, 10), dpi=200)
name_list=["鳞状细胞癌","Warthin瘤","腺样囊性癌",\
    "成釉细胞瘤","含牙囊肿","鳃裂囊肿"]
num_list=[len(x1),len(x2),len(x3),len(x4),\
    len(x5),len(x6)]
plt.xlabel("肿瘤/疾病名称")
plt.ylabel('患者数量')
plt.title('牙科疾病统计')
width = 0.65
plt.rcParams['font.sans-serif'] = ['SimHei'] # 步骤一（替换sans-serif字体）
plt.rcParams['axes.unicode_minus'] = False   # 步骤二（解决坐标轴负数的负号显示问题）
p1=plt.bar(name_list,num_list,width,color="#87CEFA")
plt.show()

#显示所有疾病的数量分布性别比
import matplotlib.pyplot as plt
import numpy as np
plt.figure(figsize=(10, 10), dpi=200)
plt.rcParams['font.sans-serif'] = ['SimHei'] # 步骤一（替换sans-serif字体）
plt.rcParams['axes.unicode_minus'] = False   # 步骤二（解决坐标轴负数的负号显示问题）
i=1
x1m=[]
x1w=[]
x2m=[]
x2w=[]
x3m=[]
x3w=[]
x4m=[]
x4w=[]
x5m=[]
x5w=[]
x6m=[]
x6w=[]
while i<nrows:
    if (i in man) & (i in x1):
          x1m.append(i)
    elif (i in woman) & (i in x1):
          x1w.append(i)
    elif (i in man) & (i in x2):
          x2m.append(i)
    elif (i in woman) & (i in x2):
          x2w.append(i)
    elif (i in man) & (i in x3):
          x3m.append(i)
    elif (i in woman) & (i in x3):
          x3w.append(i) 
    elif (i in man) & (i in x4):
          x4m.append(i)
    elif (i in woman) & (i in x4):
          x4w.append(i)
    elif (i in man) & (i in x5):
          x5m.append(i)
    elif (i in woman) & (i in x5):
          x5w.append(i)
    elif (i in man) & (i in x6):
          x6m.append(i)
    elif (i in woman) & (i in x6):
          x6w.append(i) 
    i+=1
inc=np.arange(6)
sc=[len(x1m),len(x2m),len(x3m),len(x4m),len(x5m),len(x6m)]
ss=[len(x1w),len(x2w),len(x3w),len(x4w),len(x5w),len(x6w)]
plt.yticks(np.arange(0, 76, 15)) #0到76 间隔15
plt.ylabel('number')
plt.xticks(inc,("鳞状细胞癌","Warthin瘤","腺样囊性癌",\
    "成釉细胞瘤","含牙囊肿","鳃裂囊肿"))
plt.bar(inc, sc, label='male',fc = 'b')  
plt.bar(inc, ss , label='female',bottom=sc,fc = 'r')
plt.legend()  # 给图像加上图例
plt.show()

#显示所有疾病的数量分布年龄比利
import matplotlib.pyplot as plt
import numpy as np
plt.figure(figsize=(10, 10), dpi=200)
plt.rcParams['font.sans-serif'] = ['SimHei'] # 步骤一（替换sans-serif字体）
plt.rcParams['axes.unicode_minus'] = False   # 步骤二（解决坐标轴负数的负号显示问题）
for i0 in range(1,7):
    exec( 'zx1%s = []' % i0)

for i1 in x1:
    if i1 in z1:
        zx11.append(i1)
    elif i1 in z2:
        zx12.append(i1)
    elif i1 in z3:
        zx13.append(i1)
    elif i1 in z4:
        zx14.append(i1)
    elif i1 in z5:
        zx15.append(i1)
    elif i1 in z6:
        zx16.append(i1)

for i0 in range(1,7):
    exec( 'zx2%s = []' % i0)
for i1 in x2:
    if i1 in z1:
        zx21.append(i1)
    elif i1 in z2:
        zx22.append(i1)
    elif i1 in z3:
        zx23.append(i1)
    elif i1 in z4:
        zx24.append(i1)
    elif i1 in z5:
        zx25.append(i1)
    elif i1 in z6:
        zx26.append(i1)        
for i0 in range(1,7):
    exec( 'zx3%s = []' % i0)
for i1 in x3:
    if i1 in z1:
        zx31.append(i1)
    elif i1 in z2:
        zx32.append(i1)
    elif i1 in z3:
        zx33.append(i1)
    elif i1 in z4:
        zx34.append(i1)
    elif i1 in z5:
        zx35.append(i1)
    elif i1 in z6:
        zx36.append(i1)    
for i0 in range(1,7):
    exec( 'zx4%s = []' % i0) 
for i1 in x4:
    if i1 in z1:
        zx41.append(i1)
    elif i1 in z2:
        zx42.append(i1)
    elif i1 in z3:
        zx43.append(i1)
    elif i1 in z4:
        zx44.append(i1)
    elif i1 in z5:
        zx45.append(i1)
    elif i1 in z6:
        zx46.append(i1)
for i0 in range(1,7):
    exec( 'zx5%s = []' % i0) 
for i1 in x5:
    if i1 in z1:
        zx51.append(i1)
    elif i1 in z2:
        zx52.append(i1)
    elif i1 in z3:
        zx53.append(i1)
    elif i1 in z4:
        zx54.append(i1)
    elif i1 in z5:
        zx55.append(i1)
    elif i1 in z6:
        zx56.append(i1)
for i0 in range(1,7):
    exec( 'zx4%s = []' % i0) 
for i1 in x6:
    if i1 in z1:
        zx61.append(i1)
    elif i1 in z2:
        zx62.append(i1)
    elif i1 in z3:
        zx63.append(i1)
    elif i1 in z4:
        zx64.append(i1)
    elif i1 in z5:
        zx65.append(i1)
    elif i1 in z6:
        zx66.append(i1)        
inc=np.arange(6)
s1=[len(zx11),len(zx21),len(zx31),len(zx41),len(zx51),len(zx61)]
s2=[len(zx12),len(zx22),len(zx32),len(zx42),len(zx52),len(zx62)]
s3=[len(zx13),len(zx23),len(zx33),len(zx43),len(zx53),len(zx63)]
s4=[len(zx14),len(zx24),len(zx34),len(zx44),len(zx54),len(zx64)]
s5=[len(zx15),len(zx25),len(zx35),len(zx45),len(zx55),len(zx65)]
s6=[len(zx16),len(zx26),len(zx36),len(zx46),len(zx56),len(zx66)]
plt.yticks(np.arange(0, 76, 15)) #0到76 间隔15
plt.ylabel('number')
plt.xticks(inc,("鳞状细胞癌","Warthin瘤","腺样囊性癌",\
    "成釉细胞瘤","含牙囊肿","鳃裂囊肿"))