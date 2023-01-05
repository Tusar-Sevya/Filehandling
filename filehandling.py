import pandas as pd
import os
import json
import numpy as np
from itertools import combinations, permutations
#Read data for all 3 sheet
x=pd.read_excel('charac_data.xlsx','INV',header=None).dropna(how='all')
x.dropna(how='all',axis=1,inplace=True)
x1=x.fillna('')
y=pd.read_excel('charac_data.xlsx','ND2',header=None).dropna(how='all')
y.dropna(how='all',axis=1,inplace=True)
y1=y.fillna('')
z=pd.read_excel('charac_data.xlsx','AN2',header=None).dropna(how='all')
z.dropna(how='all',axis=1,inplace=True)
z1=z.fillna('')

#find perticular index of INV sheet
l1=[i for i,j in enumerate(x1[0]) if j=='ARC']
l2 = []
for i in range(len(l1)):
  if i == len(l1)-1:
    p=(l1[i],x1.shape[0])
    l2.append(p)
  else:
    p = (l1[i], l1[i+1])
    l2.append(p)
n1=['A_R_Z_F_cell_delay','A_R_Z_F_fall_transition','A_F_Z_R_cell_delay','A_F_Z_R_rise_transition']
n2=['A_R_Z_F','A_F_Z_R','B_R_Z_F','B_F_Z_R']
n3=['A_R_Z_R','A_F_Z_F','B_R_Z_R','B_F_Z_F']

#find perticular index of ND2 sheet
l3=[i for i,j in enumerate(y1[0]) if j=='ARC']
l4=[]
for i in range(len(l3)):
    if i ==len(l3)-1:
        q=(l3[i],y1.shape[0])
        l4.append(q)
    else:
        q=(l3[i],l3[i+1])
        l4.append(q)
#find perticular index of AN2 sheet
l5=[i for i,j in enumerate(z1[0]) if j=='ARC']
l6=[]
for i in range(len(l5)):
    if i ==len(l5)-1:
        r=(l5[i],z1.shape[0])
        l6.append(r)
    else:
        r=(l5[i],l5[i+1])
        l6.append(r)
        
        
def INV(y1, y2):
    dict1={}
    x2=list(x.reset_index().drop('index', axis = 1).iloc[y1+1:y2,0:len(x.columns)][1].dropna())
    x3=x.reset_index().drop('index', axis = 1).iloc[y1+1:y2,0:len(x.columns)][1:2].dropna(axis=1).values
    x5 = x.reset_index().drop('index', axis = 1).iloc[y1+2:y2,0:len(x.columns)]
    x4=x3.tolist()
    for data in x4:
        dict1['slope']=data
    dict1['cap']=x2
    x5=x.iloc[y1+1:y2,0+1:len(x.columns)].dropna().values
    x6=x5.tolist()
    for data in x6:
        data.remove(data[0])
    x7=np.array(x6)
    dict1['values']=x7
    return dict1
def ND2(y1,y2):
    dict2={}
    x2=list(y.reset_index().drop('index', axis = 1).iloc[y1+1:y2,0:len(y.columns)][1].dropna())
    x3=y.reset_index().drop('index', axis = 1).iloc[y1+1:y2,0:len(y.columns)][1:2].dropna(axis=1).values
    x5 = y.reset_index().drop('index', axis = 1).iloc[y1+2:y2,0:len(y.columns)]
    x4=x3.tolist()
    for data in x4:
        dict2['slope']=data
    dict2['cap']=x2
    x5=y.iloc[y1+1:y2,0+1:len(y.columns)].dropna().values
    x6=x5.tolist()
    for data in x6:
        data.remove(data[0])
    x7=np.array(x6)
    dict2['values']=x7
    return dict2
def AN2(y1,y2):
    dict3={}
    x2=list(z.reset_index().drop('index', axis = 1).iloc[y1+1:y2,0:len(z.columns)][1].dropna())
    x3=z.reset_index().drop('index', axis = 1).iloc[y1+1:y2,0:len(z.columns)][1:2].dropna(axis=1).values
    x5 = z.reset_index().drop('index', axis = 1).iloc[y1+2:y2,0:len(z.columns)]
    x4=x3.tolist()
    for data in x4:
        dict3['slope']=data
    dict3['cap']=x2
    x5=z.iloc[y1+1:y2,0+1:len(z.columns)].dropna().values
    x6=x5.tolist()
    for data in x6:
        data.remove(data[0])
    x7=np.array(x6)
    dict3['values']=x7
    return dict3
  
def main():
    while True:
        print('1.INV')
        print('2.ND2')
        print('3.AN2')
        opt=input('Enter your option for which sheet?:')
        if opt=='1':
            with open('Inv.txt','w')as f:
                temp=0
                for i in range(len(l2)):
                    k,l=l2[i]
                    result=INV(k,l)
                    head=n1[temp]
                    f.write(head)
                    f.write('\n')
                    temp+=1
                    for i,j in result.items():
                        print(f'{i}:{j}',end='/\n',file=f)   
                
            os.startfile('Inv.txt')
        elif opt=='2':
          with open('ND2.txt','w')as f:
            temp=0
            for i in range(len(l4)):
              k,l=l4[i]
              result=ND2(k,l)
              
              head=n2[temp]
              f.write(head)
              f.write('\n')
              temp+=1
              for i,j in result.items():
                print(f'{i}:{j}',end='/\n',file=f)   
                
          os.startfile('ND2.txt')
        elif opt=='3':
          with open('AN2.txt','w')as f:
            temp=0
            for i in range(len(l6)):
              k,l=l6[i]
              result=AN2(k,l)
              head=n3[temp]
              f.write(head)
              f.write('\n')
              temp+=1
              for i,j in result.items():
                print(f'{i}:{j}',end='/\n',file=f)
                          
                
          os.startfile('AN2.txt')
        else:
            print('oops! you have choose wrong option please try again!')
        q=input("Do you want to continue?:")
        if q=='yes':
            continue
        else:
            break
     
main()
