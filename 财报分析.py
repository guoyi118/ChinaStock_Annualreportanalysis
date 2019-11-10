#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Nov  2 20:09:32 2019

@author: guoyi
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Nov  2 16:40:16 2019

@author: guoyi
"""

# -*- coding: utf-8 -*-full_fill
"""
Spyder Editor

This is a temporary script full_fill.
"""
import pandas as pd
import numpy as np 

import matplotlib.pyplot as plt
import matplotlib.ticker as mtick  
from pylab import mpl

#%%read fill

lrb=pd.read_excel('利润表.xls' )
xjllb=pd.read_excel('现金流量表.xls' )
zcfzb=pd.read_excel('资产负债表.xls' )

frames = [lrb, xjllb, zcfzb]
full_fill = pd.concat(frames)
full_fill = full_fill.drop_duplicates()
#    
#
full_fill=full_fill.fillna(0)
years=full_fill.columns
years=years.tolist()
full_fill=np.array(full_fill)



def add2(sheet_name,new_column_name,parameter1,parameter2):
    
    try:
        sheet_name.loc[new_column_name]=(sheet_name.loc[parameter1]+sheet_name.loc[parameter2])
    except Exception as e:    
        print(repr(e))
        print('不存在')
        pass

def min2(sheet_name,new_column_name,parameter1,parameter2):
    
    try:
        sheet_name.loc[new_column_name]=(sheet_name.loc[parameter1]-sheet_name.loc[parameter2])
    except Exception as e:    
        print(repr(e))
        print('不存在')
        pass

def add3(sheet_name,new_column_name,parameter1,parameter2,parameter3):
    
    try:
        sheet_name.loc[new_column_name]=(sheet_name.loc[parameter1]+sheet_name.loc[parameter2]+sheet_name.loc[parameter3])
    except Exception as e:    
        print(repr(e))
        print('不存在')
        pass

def divide(sheet_name,new_column_name,parameter1,parameter2):
    
    try:
        sheet_name.loc[new_column_name]=(sheet_name.loc[parameter1]/sheet_name.loc[parameter2])*100
    except Exception as e:    
        print(repr(e))
        print('不存在')

def divide3(sheet_name,new_column_name,parameter1,parameter2,parameter3):
    
    try:
        sheet_name.loc[new_column_name]=(sheet_name.loc[parameter1]+sheet_name.loc[parameter2])/sheet_name.loc[parameter2]*100
    except Exception as e:    
        print(repr(e))
        print('不存在')

def autolabel(rects):
    for rect in rects:
        height = rect.get_height()
        plt.text(rect.get_x()+rect.get_width()/2.-0.15, 1.03*height, '%.2f' % float(height),fontsize=3)

def set_ch():
    mpl.rcParams['font.sans-serif'] = ['SimHei']


def average(sheet_name,new_column_name,parameter1):
    try:
        sheet_name.loc[new_column_name]=sheet_name.loc[parameter1]
        for i in range(0,len(sheet_name.loc[parameter1])):
            if i <len(sheet_name.loc[parameter1])-1:
                sheet_name.loc[new_column_name][i] = (sheet_name.loc[parameter1][i] +sheet_name.loc[parameter1][i+1])/2
            elif i ==len(sheet_name.loc[parameter1])-1:
                sheet_name.loc[new_column_name][i]=0
    except Exception as e:    
        print(repr(e))
        print('不存在')

def increaserate(sheet_name,new_column_name,parameter1):
    try:
        sheet_name.loc[new_column_name]=sheet_name.loc[parameter1]
        for i in range(0,len(sheet_name.loc[parameter1])):
            if i <len(sheet_name.loc[parameter1])-1:
                sheet_name.loc[new_column_name][i] = (sheet_name.loc[parameter1][i] - sheet_name.loc[parameter1][i+1])/sheet_name.loc[parameter1][i+1] *100
            elif i ==len(sheet_name.loc[parameter1])-1:
                sheet_name.loc[new_column_name][i]=0
    except Exception as e:    
        print(repr(e))
        print('不存在')

def CAGR1(first, last, periods):
    
    return (last/first)**(1/periods)-1

def CAGR(sheet_name,new_column_name,parameter1):
    
    try:
        sheet_name.loc[new_column_name]=sheet_name.loc[parameter1]
        for i in range(0,len(sheet_name.loc[parameter1])):
            sheet_name.loc[new_column_name][i] = CAGR1(sheet_name.loc[parameter1][0],sheet_name.loc[parameter1][i],i+1)        
    except Exception as e:    
        print(repr(e))
        print('不存在')
        pass

def index1(sheet_name,new_column_name,parameter1):
    try:
        sheet_name.loc[new_column_name]=sheet_name.loc[parameter1]
        for i in range(0,len(sheet_name.loc[parameter1])):
            if i <len(sheet_name.loc[parameter1])-1:
                sheet_name.loc[new_column_name][i] = sheet_name.loc[parameter1][i]/sheet_name.loc[parameter1][i+1]
            elif i ==len(sheet_name.loc[parameter1])-1:
                sheet_name.loc[new_column_name][i]=0
    except Exception as e:    
        print(repr(e))
        print('不存在')


def increase(sheet_name,new_column_name,parameter1):
    try:
        sheet_name.loc[new_column_name]=sheet_name.loc[parameter1]
        for i in range(0,len(sheet_name.loc[parameter1])):
            if i <len(sheet_name.loc[parameter1])-1:
                sheet_name.loc[new_column_name][i] = sheet_name.loc[parameter1][i] - sheet_name.loc[parameter1][i+1]
            elif i ==len(sheet_name.loc[parameter1])-1:
                sheet_name.loc[new_column_name][i]=0
    except Exception as e:    
        print(repr(e))
        print('不存在')

def divide1(sheet_name,new_column_name,parameter1,parameter2):
    
    try:
        sheet_name.loc[new_column_name]=(sheet_name.loc[parameter1]/sheet_name.loc[parameter2])
    except Exception as e:    
        print(repr(e))
        print('不存在')



#%%

ylnl=[]
for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):     
        if full_fill[i,j]=='六、净利润(亿元)':
            ylnl.append(full_fill[i])
        elif full_fill[i,j]=='五、利润总额(亿元)':
            ylnl.append(full_fill[i])
        elif full_fill[i,j]=='四、营业利润(亿元)':
             ylnl.append(full_fill[i])
             
ylnl=pd.DataFrame(ylnl,columns=years)

index=ylnl.pop(ylnl.columns[0])


ylnl.index=index

grouped =  ylnl.groupby(level=0)

ylnl = grouped.first()




divide(ylnl,'营业利润占利润总额比','四、营业利润(亿元)','五、利润总额(亿元)')
#%%

writer = pd.ExcelWriter('年报分析.xlsx')

ylnl=ylnl.fillna(0)
ylnl.replace(np.inf,0,inplace=True)
ylnl.replace(-np.inf,0,inplace=True)

ylnl.to_excel(writer,'盈利能力')


name_years=years[1:]

jlr=ylnl.loc['六、净利润(亿元)'].tolist()
lrze=ylnl.loc['五、利润总额(亿元)'].tolist()
zyywlr=ylnl.loc['四、营业利润(亿元)'].tolist() 
zyywlrzeb=ylnl.loc['营业利润占利润总额比'].tolist() 





#%%
for i in range(len(name_years)):
    name_years[i]=name_years[i]

x=list(range(len(jlr)))
y=list(range(len(jlr)))

total_width, n = 0.8, 3
width = total_width / n

fig1 = plt.figure(figsize=(12,4))

ax1 = fig1.add_subplot(211)

bar1=ax1.bar(x, jlr, width=width, label='净利润',fc = 'y')

for i in range(len(x)):
    x[i] = x[i] + width
    y[i]=x[i]+ width
  
bar2=ax1.bar(x, lrze, width=width, label='利润总额(亿元)',tick_label = name_years,fc = 'r')
bar3=ax1.bar(y,zyywlr, width=width,label='营业利润(亿元)',tick_label = name_years,fc = 'b')
plt.tick_params(axis='x', labelsize=5) 


autolabel(bar1)
autolabel(bar2)
autolabel(bar3)

plt.legend()

#for x1,y1 in zip(x,lrze):
#    plt.text(x,y,'%.2f' %y, ha='center',va='top',fontsize=3)
#    
#for x2,y2 in zip(x,zyywlr):
#    plt.text(x,y,'%.2f' %y, ha='center',va='top',fontsize=3)
#    
#    
fmt='%.2f%%'
yticks = mtick.FormatStrFormatter(fmt) 


ax2 = fig1.add_subplot(212)  
ax2.plot(name_years,zyywlrzeb ,'or-',label='营业利润占利润总额比');
ax2.yaxis.set_major_formatter(yticks)

for a,b in zip(name_years,zyywlrzeb):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=7)

plt.rcParams['savefig.dpi'] = 300 
plt.rcParams['figure.dpi'] = 300 

plt.tick_params(axis='x', labelsize=4) 



plt.legend()

plt.savefig('盈利能力分析.png')
plt.show()



#%%
ylzl=[]

for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):  
        if full_fill[i,j]=='一、营业总收入(亿元)':
            ylzl.append(full_fill[i])   
        elif full_fill[i,j]=='四、营业利润(亿元)':
            ylzl.append(full_fill[i])
        elif full_fill[i,j]=='资产总计(亿元)':
            ylzl.append(full_fill[i])
        elif full_fill[i,j]=='六、净利润(亿元)':
             ylzl.append(full_fill[i])
        elif full_fill[i,j]=='五、利润总额(亿元)':
             ylzl.append(full_fill[i])                    
        elif full_fill[i,j]=='其中:利息费用(亿元)':
             ylzl.append(full_fill[i])
        elif full_fill[i,j]=='经营活动现金流入小计(亿元)':
             ylzl.append(full_fill[i])        
        elif full_fill[i,j]=='经营活动产生的现金流量净额(亿元)':
             ylzl.append(full_fill[i])                    


        
          
ylzl=pd.DataFrame(ylzl,columns=years)

index=ylzl.pop(ylzl.columns[0])


ylzl.index=index

grouped =  ylzl.groupby(level=0)

ylzl = grouped.first()

        
average(ylzl,'资产平均总额','资产总计(亿元)')
divide(ylzl,'营业利润占利润总额(亿元)比','四、营业利润(亿元)','五、利润总额(亿元)')
divide(ylzl,'全部资产现金回收率','经营活动产生的现金流量净额(亿元)','资产平均总额')
divide(ylzl,'盈利现金比率','经营活动产生的现金流量净额(亿元)','六、净利润(亿元)')
divide(ylzl,'销售收现比率','经营活动现金流入小计(亿元)','一、营业总收入(亿元)')
divide(ylzl,'利润现金保证倍数','经营活动产生的现金流量净额(亿元)','六、净利润(亿元)')





ylzl=ylzl.fillna(0)
ylzl.replace(np.inf,0,inplace=True)
ylzl.replace(-np.inf,0,inplace=True)




ylzl.to_excel(writer,'盈利质量')

jyhdxjjll=ylzl.loc['经营活动产生的现金流量净额(亿元)'].tolist()
jlr=ylzl.loc['六、净利润(亿元)'].tolist()
lrxjbzbs=ylzl.loc['利润现金保证倍数'].tolist() 
xssxbl=ylzl.loc['销售收现比率'].tolist() 
qbzcxjhsl=ylzl.loc['全部资产现金回收率'].tolist() 

total_width, n = 0.8, 2
width = total_width / n
fig2 = plt.figure(figsize=(12,4))
ax3 = fig2.add_subplot(211)
bar1=ax3.bar(x, jlr, width=width, label='净利润',fc = 'y')
for i in range(len(x)):
    x[i] = x[i] + width    
bar2=ax3.bar(x, jyhdxjjll, width=width, label='经营活动产生的现金流量净额(亿元)',tick_label = name_years,fc = 'r')
plt.tick_params(axis='x', labelsize=5) 

autolabel(bar1)
autolabel(bar2)


plt.legend()
fmt='%.2f%%'
yticks = mtick.FormatStrFormatter(fmt) 
   
ax4 = fig2.add_subplot(212)  
ax4.plot(name_years,lrxjbzbs ,color='green',label='利润现金保证倍数');
ax4.plot(name_years,xssxbl ,color='red',label='销售收现比率');
ax4.plot(name_years,qbzcxjhsl ,color='skyblue',label='全部资产现金回收率');

ax4.yaxis.set_major_formatter(yticks)

for a,b in zip(name_years,lrxjbzbs):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,xssxbl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,qbzcxjhsl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

plt.legend()
plt.savefig('盈利质量分析.png')
plt.tick_params(axis='x', labelsize=4) 
plt.show()



#%%
czxfx=[]
for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):  
        if full_fill[i,j]=='一、营业总收入(亿元)':
            czxfx.append(full_fill[i])   
        elif full_fill[i,j]=='资产总计(亿元)':
            czxfx.append(full_fill[i])
        elif full_fill[i,j]=='六、净利润(亿元)':
             czxfx.append(full_fill[i])
        elif full_fill[i,j]=='股东权益合计(亿元)':
             czxfx.append(full_fill[i])                    
        elif full_fill[i,j]=='一、营业总收入(亿元)增长率':
             czxfx.append(full_fill[i])                           
             
    
                       
czxfx=pd.DataFrame(czxfx,columns=years)

index=czxfx.pop(czxfx.columns[0])


czxfx.index=index


grouped =  czxfx.groupby(level=0)

czxfx = grouped.first()


increaserate(czxfx,'净利润增长率','六、净利润(亿元)')
increaserate(czxfx,'资产总计(亿元)增长率','资产总计(亿元)')
increaserate(czxfx,'股东权益合计(亿元)增长率','股东权益合计(亿元)')



CAGR(czxfx,'营业总收入(亿元)复合增长率','一、营业总收入(亿元)')
CAGR(czxfx,'净利润复合增长率·','六、净利润(亿元)')
CAGR(czxfx,'资产总计(亿元)复合增长率','资产总计(亿元)')
CAGR(czxfx,'股东权益合计(亿元)复合增长率','股东权益合计(亿元)')

czxfx=czxfx.fillna(0)
czxfx.replace(np.inf,0,inplace=True)
czxfx.replace(-np.inf,0,inplace=True)


czxfx.to_excel(writer,'成长性分析')


jlr=czxfx.loc['六、净利润(亿元)'].tolist()
zyywsr=czxfx.loc['一、营业总收入(亿元)'].tolist()
zzc=czxfx.loc['资产总计(亿元)'].tolist() 

jlrzzl=czxfx.loc['净利润增长率'].tolist()
zzczzl=czxfx.loc['资产总计(亿元)增长率'].tolist()
gdqyzzl=czxfx.loc['股东权益合计(亿元)增长率'].tolist() 


fig3 = plt.figure(figsize=(12,4))

ax5 = fig3.add_subplot(211)
ax5.plot(name_years,jlr ,color='green',label='六、净利润(亿元)');
ax5.plot(name_years,zyywsr ,color='red',label='一、营业总收入(亿元)');
ax5.plot(name_years,zzc,color='blue',label='资产总计(亿元)');

for a,b in zip(name_years,jlr):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,zyywsr):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,zzc):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

plt.tick_params(axis='x', labelsize=5) 





plt.legend()
#plt.savefig('盈利能力分析1.png')

fmt='%.2f%%'
yticks = mtick.FormatStrFormatter(fmt) 


ax6 = fig3.add_subplot(212)  
ax6.plot(name_years,jlrzzl ,color='green',label='净利润增长率');
ax6.plot(name_years,zzczzl ,color='red',label='资产总计(亿元)增长率');
ax6.plot(name_years,gdqyzzl,color='blue',label='股东权益合计(亿元)增长率');
ax6.yaxis.set_major_formatter(yticks)


for a,b in zip(name_years,jlrzzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,zzczzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,gdqyzzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

plt.tick_params(axis='x', labelsize=5) 

plt.legend()
plt.savefig('成长性分析.png')
plt.show()



#%%

jylnfx=[]
for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):  
        if full_fill[i,j]=='一、营业总收入(亿元)':
            jylnfx.append(full_fill[i])   
        elif full_fill[i,j]=='四、营业利润(亿元)':
            jylnfx.append(full_fill[i])
        elif full_fill[i,j]=='加:投资收益(亿元)':
             jylnfx.append(full_fill[i])
        elif full_fill[i,j]=='加:营业外收入(亿元)':
             jylnfx.append(full_fill[i])                    
        elif full_fill[i,j]=='营业成本(亿元)':
             jylnfx.append(full_fill[i])                           
        elif full_fill[i,j]=='销售费用(亿元)':
            jylnfx.append(full_fill[i])
        elif full_fill[i,j]=='管理费用(亿元)':
             jylnfx.append(full_fill[i])
        elif full_fill[i,j]=='财务费用(亿元)':
             jylnfx.append(full_fill[i])                    
        elif full_fill[i,j]=='研发费用(亿元)':
             jylnfx.append(full_fill[i])                
         


    
               
jylnfx=pd.DataFrame(jylnfx,columns=years)

index=jylnfx.pop(jylnfx.columns[0])


jylnfx.index=index

grouped =jylnfx.groupby(level=0)

jylnfx = grouped.first()

    
add3(jylnfx,'收入总额','一、营业总收入(亿元)','加:投资收益(亿元)','加:营业外收入(亿元)')
divide(jylnfx,'营业总收入占比','一、营业总收入(亿元)','收入总额')
increaserate(jylnfx,'营业总收入(亿元)增长率','一、营业总收入(亿元)')


try:
    jylnfx.loc['成本费用总额']=jylnfx.loc['营业成本(亿元)']+jylnfx.loc['销售费用(亿元)']+jylnfx.loc['管理费用(亿元)']+jylnfx.loc['财务费用(亿元)']+jylnfx.loc['研发费用(亿元)'] 
except Exception as e:    
    print(repr(e))
    print('不存在')
    pass

increaserate(jylnfx,'销售费用(亿元)增长率','销售费用(亿元)')
increaserate(jylnfx,'研发费用(亿元)增长率','研发费用(亿元)')



jylnfx=jylnfx.fillna(0)
jylnfx.replace(np.inf,0,inplace=True)
jylnfx.replace(-np.inf,0,inplace=True)


jylnfx.to_excel(writer,'经营理念分析')


xsfy=jylnfx.loc['销售费用(亿元)'].tolist()
yffy=jylnfx.loc['研发费用(亿元)'].tolist()

zyywsrzzl=jylnfx.loc['营业总收入(亿元)增长率'].tolist() 
xsfyzzl=jylnfx.loc['销售费用(亿元)增长率'].tolist() 
yffyzzl=jylnfx.loc['研发费用(亿元)增长率'].tolist() 



total_width, n = 0.8, 2
width = total_width / n

fig4 = plt.figure(figsize=(12,4))

ax7 = fig4.add_subplot(211)

bar1=ax7.bar(x, xsfy, width=width, label='销售费用(亿元)',fc = 'y')
for i in range(len(x)):
    x[i] = x[i] + width
    
bar2=ax7.bar(x, yffy, width=width, label='研发费用(亿元)',tick_label = name_years,fc = 'r')

autolabel(bar1)
autolabel(bar2)

plt.tick_params(axis='x', labelsize=5)

plt.legend()

fmt='%.2f%%'
yticks = mtick.FormatStrFormatter(fmt) 
   
ax8 = fig4.add_subplot(212)  
ax8.plot(name_years,zyywsrzzl ,color='green',label='营业总收入(亿元)增长率');
ax8.plot(name_years,xsfyzzl ,color='red',label='销售费用(亿元)增长率');
ax8.plot(name_years,yffyzzl ,color='skyblue',label='研发费用(亿元)增长率');


for a,b in zip(name_years,zyywsrzzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,xsfyzzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,yffyzzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

plt.tick_params(axis='x', labelsize=5) 

ax8.yaxis.set_major_formatter(yticks)
plt.legend()
plt.savefig('经营理念分析.png')
plt.show()
#%%

scdw=[]
for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):  
        if full_fill[i,j]=='一、营业总收入(亿元)':
            scdw.append(full_fill[i])   
        elif full_fill[i,j]=='营业成本(亿元)':
            scdw.append(full_fill[i])            
    
scdw=pd.DataFrame(scdw,columns=years)

index=scdw.pop(scdw.columns[0])


scdw.index=index         


min2(scdw,'毛利','一、营业总收入(亿元)','营业成本(亿元)')
divide(scdw,'毛利率','毛利','一、营业总收入(亿元)')

scdw=scdw.fillna(0)
scdw.replace(np.inf,0,inplace=True)
scdw.replace(-np.inf,0,inplace=True)


scdw.to_excel(writer,'市场地位')

mll=scdw.loc['毛利率'].tolist()


plt.plot(name_years,mll,color='blue',label='毛利率')

for a,b in zip(name_years,mll):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

  
plt.tick_params(axis='x', labelsize=5)

plt.legend()

plt.savefig('市场定位.png')

plt.show()  

#%%
glcfx=[]
for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):  
        if full_fill[i,j]=='一、营业总收入(亿元)':
            glcfx.append(full_fill[i])   
        elif full_fill[i,j]=='资产总计(亿元)':
            glcfx.append(full_fill[i])
        elif full_fill[i,j]=='营业成本(亿元)':
             glcfx.append(full_fill[i])                    
        elif full_fill[i,j]=='存货(亿元)':
             glcfx.append(full_fill[i])                           
        elif full_fill[i,j]=='应收账款(亿元)':
            glcfx.append(full_fill[i])
              
glcfx=pd.DataFrame(glcfx,columns=years)

index=glcfx.pop(glcfx.columns[0])


glcfx.index=index    



average(glcfx,'资产平均总额','资产总计(亿元)')
average(glcfx,'平均存货(亿元)','存货(亿元)')
divide(glcfx,'存货(亿元)周转率','营业成本(亿元)','平均存货(亿元)') 
average(glcfx,'应收账款(亿元)平均余额','应收账款(亿元)')
divide(glcfx,'应收账款(亿元)周转率','一、营业总收入(亿元)','应收账款(亿元)平均余额') 
divide(glcfx,'资产周转率','一、营业总收入(亿元)','资产平均总额')

glcfx=glcfx.fillna(0)
glcfx.replace(np.inf,0,inplace=True)
glcfx.replace(-np.inf,0,inplace=True)

glcfx.to_excel(writer,'管理层分析')


chzzl=glcfx.loc['存货(亿元)周转率'].tolist() 
yszkzzl=glcfx.loc['应收账款(亿元)周转率'].tolist() 
zczzl=glcfx.loc['资产周转率'].tolist() 





plt.plot(name_years,chzzl ,color='green',label='存货(亿元)周转率');
plt.plot(name_years,yszkzzl ,color='red',label='应收账款(亿元)周转率');
plt.plot(name_years,zczzl,color='skyblue',label='资产周转率');


for a,b in zip(name_years,chzzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)
#
for a,b in zip(name_years,yszkzzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)
#
for a,b in zip(name_years,zczzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

plt.tick_params(axis='x', labelsize=4) 


plt.legend()
plt.savefig('管理层分析.png')
plt.show()

#%%

zcfxfx=[]

for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):     
        if full_fill[i,j]=='应收账款(亿元)':
            zcfxfx.append(full_fill[i])
        elif full_fill[i,j]=='存货(亿元)':
            zcfxfx.append(full_fill[i])
        elif full_fill[i,j]=='长期待摊费用(亿元)':
             zcfxfx.append(full_fill[i])
        elif full_fill[i,j]=='商誉(亿元)':
            zcfxfx.append(full_fill[i])
        elif full_fill[i,j]=='资产总计(亿元)':
             zcfxfx.append(full_fill[i])
           

zcfxfx=pd.DataFrame(zcfxfx,columns=years)

index=zcfxfx.pop(zcfxfx.columns[0])

zcfxfx.index=index

divide(zcfxfx,'应收账款占比','应收账款(亿元)','资产总计(亿元)')
divide(zcfxfx,'存货占比','存货(亿元)','资产总计(亿元)')
divide(zcfxfx,'长期待摊费用占比','长期待摊费用(亿元)','资产总计(亿元)')
divide(zcfxfx,'商誉占比','商誉(亿元)','资产总计(亿元)')


zcfxfx=zcfxfx.fillna(0)
zcfxfx.replace(np.inf,0,inplace=True)
zcfxfx.replace(-np.inf,0,inplace=True)

zcfxfx.to_excel(writer,'资产风险分析')


yszk=zcfxfx.loc['应收账款(亿元)'].tolist()
ch=zcfxfx.loc['存货(亿元)'].tolist()
cqdtfy=zcfxfx.loc['长期待摊费用(亿元)'].tolist() 
sy=zcfxfx.loc['商誉(亿元)'].tolist() 
zczj=zcfxfx.loc['资产总计(亿元)'].tolist()
yszkzb=zcfxfx.loc['应收账款占比'].tolist()
chzb=zcfxfx.loc['存货占比'].tolist() 
cqdtfyzb=zcfxfx.loc['长期待摊费用占比'].tolist() 
syzb=zcfxfx.loc['商誉占比'].tolist() 


for i in range(len(name_years)):
    name_years[i]=name_years[i]

x=list(range(len(yszk)))
y=list(range(len(yszk)))
z=list(range(len(yszk)))

total_width, n = 0.8, 4
width = total_width / n

fig1 = plt.figure(figsize=(12,4))

ax1 = fig1.add_subplot(211)

bar1=ax1.bar(x, yszk, width=width, label='应收账款(亿元)',fc = 'y')

for i in range(len(x)):
    x[i] = x[i] + width
    y[i]=x[i]+ width
    z[i]=y[i]+width
    
bar2=ax1.bar(x, ch, width=width, label='存货(亿元)',tick_label = name_years,fc = 'r')
bar3=ax1.bar(y,sy, width=width,label='商誉(亿元)',tick_label = name_years,fc = 'b')
bar4=ax1.bar(z,zczj, width=width,label='资产总计(亿元)',tick_label = name_years,fc = 'g')
plt.tick_params(axis='x', labelsize=4) 


autolabel(bar1)
autolabel(bar2)
autolabel(bar3)
autolabel(bar4)

plt.legend()
#for x1,y1 in zip(x,lrze):
#    plt.text(x,y,'%.2f' %y, ha='center',va='top',fontsize=3)
#    
#for x2,y2 in zip(x,zyywlr):
#    plt.text(x,y,'%.2f' %y, ha='center',va='top',fontsize=3)
#    
#    
fmt='%.2f%%'
yticks = mtick.FormatStrFormatter(fmt) 


ax2 = fig1.add_subplot(212)  
ax2.plot(name_years,yszkzb ,'or-',label='应收账款占比',c = 'red');
ax2.plot(name_years,chzb ,'or-',label='存货占比',c='green');
ax2.plot(name_years,syzb ,'or-',label='商誉占比',c='blue');


ax2.yaxis.set_major_formatter(yticks)

for a,b in zip(name_years,yszkzb):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,chzb):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

for a,b in zip(name_years,syzb):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

plt.rcParams['savefig.dpi'] = 300 
plt.rcParams['figure.dpi'] = 300 

plt.tick_params(axis='x', labelsize=4) 



plt.legend()

plt.savefig('资产风险分析.png')
plt.show()



#%%
zcfzfx=[]

for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):     
        if full_fill[i,j]=='负债合计(亿元)':
            zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='经营活动产生的现金流量净额(亿元)':
            zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='资产总计(亿元)':
             zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='股东权益合计(亿元)':
            zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='流动资产合计(亿元)':
             zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='货币资金(亿元)':
            zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='交易性金融资产(亿元)':
             zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='存货(亿元)':
            zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='五、利润总额(亿元)':
             zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='其中:利息费用(亿元)':
             zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='流动资产合计(亿元)':
            zcfzfx.append(full_fill[i])
        elif full_fill[i,j]=='流动负债合计(亿元)':
             zcfzfx.append(full_fill[i])
           

zcfzfx=pd.DataFrame(zcfzfx,columns=years)

index=zcfzfx.pop(zcfzfx.columns[0])

zcfzfx.index=index

grouped = zcfzfx.groupby(level=0)

zcfzfx = grouped.first()

add2(zcfzfx,'息税前利润','五、利润总额(亿元)','其中:利息费用(亿元)')
min2(zcfzfx,'速动资产','流动资产合计(亿元)','存货(亿元)')
min2(zcfzfx,'净运营资本','流动资产合计(亿元)','流动负债合计(亿元)')
divide(zcfzfx,'负债率','负债合计(亿元)','资产总计(亿元)')
divide(zcfzfx,'流动比率','流动资产合计(亿元)','流动负债合计(亿元)')
divide3(zcfzfx,'现金比率','货币资金(亿元)','交易性金融资产(亿元)','流动负债合计(亿元)')
divide(zcfzfx,'现金流量比率','经营活动产生的现金流量净额(亿元)','流动负债合计(亿元)')
divide(zcfzfx,'速动比率','速动资产','流动负债合计(亿元)')
divide(zcfzfx,'产权比率','负债合计(亿元)','股东权益合计(亿元)')
divide(zcfzfx,'权益比率','股东权益合计(亿元)','资产总计(亿元)')
divide(zcfzfx,'权益乘数','资产总计(亿元)','股东权益合计(亿元)')
divide(zcfzfx,'现金流量利息保障倍数','经营活动产生的现金流量净额(亿元)','其中:利息费用(亿元)')
divide(zcfzfx,'经营现金流量债务比','经营活动产生的现金流量净额(亿元)','负债合计(亿元)')




zcfzfx=zcfzfx.fillna(0)
zcfzfx.replace(np.inf,0,inplace=True)
zcfzfx.replace(-np.inf,0,inplace=True)


zcfzfx.to_excel(writer,'资产负债分析')


fzhj=zcfzfx.loc['负债合计(亿元)'].tolist()
zchj=zcfzfx.loc['资产总计(亿元)'].tolist()
fzl=zcfzfx.loc['负债率'].tolist() 
ldbl=zcfzfx.loc['流动比率'].tolist()
sdbl=zcfzfx.loc['速动比率'].tolist()
cqbl=zcfzfx.loc['产权比率'].tolist() 
qycs=zcfzfx.loc['权益乘数'].tolist() 




for i in range(len(name_years)):
    name_years[i]=name_years[i]

x=list(range(len(yszk)))
y=list(range(len(yszk)))
z=list(range(len(yszk)))

total_width, n = 0.8, 4
width = total_width / n

fig1 = plt.figure(figsize=(12,4))

ax1 = fig1.add_subplot(311)

bar1=ax1.bar(x, ldbl, width=width, label='流动比率',fc = 'y')

for i in range(len(x)):
    x[i] = x[i] + width
    y[i]=x[i]+ width
    z[i]=y[i]+width
        
bar2=ax1.bar(x, sdbl, width=width, label='速动比率',tick_label = name_years,fc = 'r')
bar3=ax1.bar(y,cqbl, width=width,label='产权比率',tick_label = name_years,fc = 'b')
bar4=ax1.bar(z,qycs, width=width,label='权益乘数',tick_label = name_years,fc = 'g')
plt.tick_params(axis='x', labelsize=4) 


autolabel(bar1)
autolabel(bar2)
autolabel(bar3)
autolabel(bar4)

plt.legend()
#for x1,y1 in zip(x,lrze):
#    plt.text(x,y,'%.2f' %y, ha='center',va='top',fontsize=3)
#    
#for x2,y2 in zip(x,zyywlr):
#    plt.text(x,y,'%.2f' %y, ha='center',va='top',fontsize=3)
#    
# 



   
fmt='%.2f%%'
yticks = mtick.FormatStrFormatter(fmt) 


ax2 = fig1.add_subplot(312)  
ax2.plot(name_years,fzl ,'or-',label='应收账款占比',c = 'red');

ax2.yaxis.set_major_formatter(yticks)


for a,b in zip(name_years,fzl):
    plt.text(a, b+0.05, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

plt.rcParams['savefig.dpi'] = 300 
plt.rcParams['figure.dpi'] = 300 

plt.tick_params(axis='x', labelsize=4) 



plt.legend()

ax3 = fig1.add_subplot(313)

bar1=ax3.bar(x, fzhj, width=width, label='流动比率',fc = 'y')

for i in range(len(x)):
    x[i] = x[i] + width
    
bar2=ax3.bar(x, zchj, width=width, label='速动比率',tick_label = name_years,fc = 'r')
plt.tick_params(axis='x', labelsize=4) 


autolabel(bar1)
autolabel(bar2)

plt.legend()


plt.savefig('资产负债分析.png')
plt.show()

#%%

fznlfx=[]
for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):     
        if full_fill[i,j]=='六、净利润(亿元)':
            fznlfx.append(full_fill[i])
        elif full_fill[i,j]=='资产总计(亿元)':
             fznlfx.append(full_fill[i])
        elif full_fill[i,j]=='股东权益合计(亿元)':
            fznlfx.append(full_fill[i])
        elif full_fill[i,j]=='四、营业利润(亿元)':
             fznlfx.append(full_fill[i])
 
fznlfx=pd.DataFrame(fznlfx,columns=years)

index=fznlfx.pop(fznlfx.columns[0])

fznlfx.index=index

grouped = fznlfx.groupby(level=0)

fznlfx = grouped.first()

fznlfx.loc['毛利']=scdw.loc['毛利']
average(fznlfx,'股东权益平均总额','股东权益合计(亿元)')
increaserate(fznlfx,'总资产增长率','资产总计(亿元)')
increaserate(fznlfx,'股东权益增长率','股东权益合计(亿元)')
increaserate(fznlfx,'主营业务收入增长率','四、营业利润(亿元)')
increaserate(fznlfx,'净利润增长率','六、净利润(亿元)')
increaserate(fznlfx,'毛利增长率','毛利')
divide(fznlfx,'净资产收益率','六、净利润(亿元)','股东权益平均总额')



fznlfx=fznlfx.fillna(0)
fznlfx.replace(np.inf,0,inplace=True)
fznlfx.replace(-np.inf,0,inplace=True)


fznlfx.to_excel(writer,'发展能力分析')

#writer.save()

gdqyhj=fznlfx.loc['股东权益合计(亿元)'].tolist()
jlr=fznlfx.loc['六、净利润(亿元)'].tolist()
jzcsyl=fznlfx.loc['净资产收益率'].tolist() 
zzczzl=fznlfx.loc['总资产增长率'].tolist()
zyywsrzzl=fznlfx.loc['主营业务收入增长率'].tolist()
gdqyzzl=fznlfx.loc['股东权益增长率'].tolist() 
jlrzzl=fznlfx.loc['净利润增长率'].tolist() 




for i in range(len(name_years)):
    name_years[i]=name_years[i]

x=list(range(len(yszk)))
y=list(range(len(yszk)))
z=list(range(len(yszk)))

total_width, n = 0.8, 4
width = total_width / n

fig6 = plt.figure(figsize=(12,4))

ax1 = fig6.add_subplot(311)

bar1=ax1.bar(x, jlrzzl, width=width, label='净利润增长率',fc = 'y')

for i in range(len(x)):
    x[i] = x[i] + width
    y[i]=x[i]+ width
    z[i]=y[i]+width
        
bar2=ax1.bar(x, zzczzl, width=width, label='总资产增长率',tick_label = name_years,fc = 'r')
bar3=ax1.bar(y,zyywsrzzl, width=width,label='主营业务收入增长率',tick_label = name_years,fc = 'b')
bar4=ax1.bar(z,gdqyzzl, width=width,label='股东权益增长率',tick_label = name_years,fc = 'g')
plt.tick_params(axis='x', labelsize=4) 


autolabel(bar1)
autolabel(bar2)
autolabel(bar3)
autolabel(bar4)

plt.legend()



ax2 = fig6.add_subplot(312)  
bar2=ax2.bar(x, gdqyhj, width=width, label='股东权益合计',tick_label = name_years,fc = 'r')
bar3=ax2.bar(y,jlr, width=width,label='净利润',tick_label = name_years,fc = 'b')

plt.rcParams['savefig.dpi'] = 300 
plt.rcParams['figure.dpi'] = 300 

plt.tick_params(axis='x', labelsize=4) 

autolabel(bar2)
autolabel(bar3)

plt.legend()


fmt='%.2f%%'
yticks = mtick.FormatStrFormatter(fmt) 


ax3 = fig6.add_subplot(313)  
ax3.plot(name_years, jzcsyl ,'or-',label='净利润收益率',c = 'red');

ax3.yaxis.set_major_formatter(yticks)

for a,b in zip(name_years,jzcsyl):
    plt.text(a, b, '%.2f' % b, ha='center', va= 'bottom',fontsize=3)

plt.rcParams['savefig.dpi'] = 300 
plt.rcParams['figure.dpi'] = 300 

plt.tick_params(axis='x', labelsize=4) 


plt.savefig('发展能力分析.png')
plt.show()
 #%%
M=[]
for i in range(len(full_fill)):
    oneline=full_fill[i]
    for j in range(len(oneline)):     
        if full_fill[i,j]=='货币资金(亿元)':
            M.append(full_fill[i])
        elif full_fill[i,j]=='流动资产合计(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='流动负债合计(亿元)':
            M.append(full_fill[i])
        elif full_fill[i,j]=='一年内到期的非流动负债(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='应交税费(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='应收账款(亿元)':
            M.append(full_fill[i])
        elif full_fill[i,j]=='营业收入(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='营业成本(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='流动负债合计(亿元)':
            M.append(full_fill[i])
        elif full_fill[i,j]=='非流动资产合计(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='固定资产(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='在建工程(亿元)':
            M.append(full_fill[i])
        elif full_fill[i,j]=='工程物资(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='资产总计(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='销售费用(亿元)':
            M.append(full_fill[i])
        elif full_fill[i,j]=='管理费用(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='负债合计(亿元)':
             M.append(full_fill[i])
        elif full_fill[i,j]=='固定资产和投资性房地产折旧(亿元)':
            M.append(full_fill[i])

M=pd.DataFrame(M,columns=years)

index=M.pop(M.columns[0])

M.index=index

grouped = M.groupby(level=0)

M = grouped.first()

M.loc['毛利']=scdw.loc['毛利']

try:
    M.loc['非实物资产']=M.loc['非流动资产合计(亿元)']-M.loc['固定资产(亿元)']-M.loc['在建工程(亿元)']-M.loc['工程物资(亿元)']
except Exception as e:    
    print(repr(e))
    print('不存在')
    pass

add2(M,'销售管理费用','销售费用(亿元)','管理费用(亿元)')
add2(M,'固定资产原值','固定资产(亿元)','固定资产和投资性房地产折旧(亿元)')
divide(M,'应收账款占营业收入比例','应收账款(亿元)','营业收入(亿元)')
divide(M,'毛利率','毛利','营业收入(亿元)')
index1(M,'毛利率指数','毛利')
index1(M,'应收账款指数','应收账款占营业收入比例')
index1(M,'资产质量指数','非实物资产')
index1(M,'营业收入指数','营业收入(亿元)')
divide(M,'折旧率','固定资产和投资性房地产折旧(亿元)','固定资产原值')
index1(M,'折旧率指数','折旧率')
index1(M,'销售管理费用指数','销售管理费用')
divide(M,'资产负债率','负债合计(亿元)','资产总计(亿元)')
index1(M,'财务杠杆指数','资产负债率')
increase(M,'流动资产改变量','流动资产合计(亿元)')
increase(M,'货币资金改变量','货币资金(亿元)')
increase(M,'流动负债改变量','流动负债合计(亿元)')
increase(M,'一年内到期的长期负债改变量','一年内到期的非流动负债(亿元)')
increase(M,'应交税费改变量','应交税费(亿元)')

try:
    M.loc['应计项']=(M.loc['流动资产改变量']-M.loc['货币资金改变量'])-(M.loc['流动负债改变量']-M.loc['一年内到期的长期负债改变量']-M.loc['应交税费改变量'])-M.loc['固定资产和投资性房地产折旧(亿元)']
except Exception as e:    
    print(repr(e))
    print('不存在')
    pass

divide1(M,'总应计项','应计项','资产总计(亿元)')

try:
    M.loc['m-score']=M.loc['总应计项']
    for i in range(0,len(M.loc['总应计项'])):
            if i <len(M.loc['总应计项'])-1:
                M.loc['m-score'][i] = -4.84+0.92*M.loc['应收账款指数'][i] +0.528*M.loc['毛利率指数'][i] +0.404* M.loc['资产质量指数'][i]
                +0.892*M.loc['营业收入指数'][i]+0.115*M.loc['折旧率指数'][i]
                -0.172*M.loc['销售管理费用指数'][i]-0.327*M.loc['财务杠杆指数'][i]+4.697*M.loc['总应计项'][i]
            elif i ==len(M.loc['总应计项'])-1:
                M.loc['m-score'][i]=0
except Exception as e:    
    print(repr(e))
    print('不存在')




M=M.fillna(0)
M.replace(np.inf,0,inplace=True)
M.replace(-np.inf,0,inplace=True)


M.to_excel(writer,'M-score')

writer.save()
