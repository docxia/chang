hhg="子宫，卵巢，小肠， 肺，肝"
k=0
if "子宫" in hhg:
    k=k+1
    print(k)
if "卵巢" in hhg:
    k=k+1
    print(k)
k

if  ("子宫" in hhg) & ("卵巢" in hhg):
    print("yes")
import re
hhu='（乙状结肠）中分化腺癌，肿物大小4.7×3.5×1.5cm，浸润至肠壁外脂肪组织；（近端、远端切缘）均未见癌；脉管内未见癌栓；肠系膜淋巴结见癌转移（5/17），肠系膜内见癌结节4枚。（建议加做KRAS、NRAS、PI3KCA、BRAF基因检测，指导个体化治疗。）'
N=0
nrows=1
while N<nrows:
    if "/" not in hhu:
          NX.append("")    
    else: 
         ff1=re.findall(r'\d+/',hhu)
         k=len(ff1)
         ff2=[]
         i=0
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

ff1=['3/','4/']
ff2=[]
i=0
k=len(ff1)
         while i<k:
             x=re.findall(r'\d+',ff1[i])
             ff2.append(x[0])
             i+=1
         ff2
         ff2=list(map(int,ff2))


x = []*6 #设置连续变量名
col=1
while col<7:
    x[col] = '';
    print(x[col])
    col+=1
x1=[]
gh="x"+str(1)
class test(object):
t = test()

for i in range(1, 11):

setattr(t, "a" + str(i), [])

print t.__dict__

print t.a1
gh
i=1
while i<4:

    x[i]=
    i=i+1

n = 3
for i in range(1, n+1):
    exec("lst%s =[[] for _ in range(1)]"% (i, n))
for i in range(1,11):
    exec( 'a%s = []' % i)

x=[1,2,3,4,5,56]
z=[2,4,0,9,89,6]
for i in x:
    if i in z:
        print(i)
ks=[]
len(ks)

for i4 in range(1,7):
    exec( 'zx3%s = []' % i4)

for i4 in range(7,13):
    exec( 'zx4%s = []' % i4)

import matplotlib.pyplot as plt
import numpy as np
plt.figure(figsize=(10, 10), dpi=200)
plt.rcParams['font.sans-serif'] = ['SimHei'] # 步骤一（替换sans-serif字体）
plt.rcParams['axes.unicode_minus'] = False   # 步骤二（解决坐标轴负数的负号显示问题）
x1=[4,5,6]
x3=[3,4,5]
x2=[1,2,3]
plt.yticks(np.arange(0, 76, 15)) #0到76 间隔15
plt.ylabel('number')
plt.xticks(inc,("鳞状细胞癌","Warthin瘤","腺样囊性癌"))


