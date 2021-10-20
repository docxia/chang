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