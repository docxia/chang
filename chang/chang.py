
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
while colnamber<17:
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
    if "结肠" in col[10][i1]:
        print(col[10][i1])
    elif "直肠" in col[10][i1]:
        print(col[10][i1])
    elif "结肠" in col[11][i1]:
        print(col[11][i1])
    elif "直肠" in col[11][i1]:
        print(col[11][i1])
    else:
        feijiezhichangai.append(i1)
    i1=i1+1 
feijiezhichangai
len(feijiezhichangai)
paichu=[]
paichu=feijiezhichangai+menzhen
paichu=list(set(paichu))#去重复
paichu.sort(reverse = False)#排列，从小到大
paichu
len(paichu)
len(paichu)
#接下来是删除任务

