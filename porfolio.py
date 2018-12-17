import pandas as pd
from dateutil.relativedelta import *

stock_price=pd.read_excel('portfolio.xlsx',sheet_name='stock_price')
sales_value=pd.read_excel('portfolio.xlsx',sheet_name='sales_value')
sales_unit=pd.read_excel('portfolio.xlsx',sheet_name='sales_unit')
value_set=stock_price.copy()

def date(df_):
    df_= pd.to_datetime(df_)
    return df_

stock_price['date']=date(stock_price['date'])
sales_value['date']=date(sales_value['date'])
sales_unit['date']=date(sales_unit['date'])


def locatedate_calcu_value(c_,lead,lst_top,lst_bot):
    period=[]    
    #if date_>=(c_) and date_<(date_+relativedelta(months=+1)):
    date_=c_+relativedelta(day=lead)
    index_=pd.Index(stock_price['date']).get_loc(date_,method='nearest')
    period.append(index_)
    date_=date_+relativedelta(months=+1)
    index_=pd.Index(stock_price['date']).get_loc(date_,method='nearest')
    period.append(index_)
    
    ini_s=period[0]
    ini_e=period[-1]
    
    for z in range(ini_s,ini_e):
        for w in range(1,len(stock_price.columns)):
            value_set.iloc[z,w]=0
        
    for i in range(1,len(period)):
        s=period[i-1]
        e=period[i]
        n=len(lst_top)  
        for i in range(s,e):
            for x in range(n):
                value_set.iloc[i,lst_top[x]]=(stock_price.iloc[i,lst_top[x]]/stock_price.iloc[s,lst_top[x]])*1000    
                value_set.iloc[i,lst_bot[x]]=(stock_price.iloc[i,lst_bot[x]]/stock_price.iloc[s,lst_bot[x]])*(-1000)

def mom(n,pre,ratio,lead_day): # n+pre<24
    dic={}
    if(pre.month+n)-13>0:
        c_m=pre.month+n-12
        c_y=pre.year+1
    else:
        c_m=pre.month+n
        c_y=pre.year
    c_date=pd.to_datetime(str(c_y)+'-'+str(c_m))
    p_date=pre
    # calculate
    for i in range(1,len(sales_value[sales_value.index==c_date].columns)):
        c_value=int(sales_value.loc[sales_value['date']==c_date,sales_value.columns[i]])
        p_value=int(sales_value.loc[sales_value['date']==p_date,sales_value.columns[i]])
        dic.update({i:c_value/p_value})
    lst=sorted(dic.values())
    number=int(ratio*(len(sales_value[sales_value.index==c_date].columns)-1))
    lst_top=[]
    lst_bot=[]
    for num in range(number):
        for key, value in dic.items():
            if value== lst[num]:
                lst_bot.append(key)
        for key, value in dic.items():
            if value== lst[-num-1]:
                lst_top.append(key)
    locatedate_calcu_value(c_date,lead_day,lst_top,lst_bot)

mom(1,pd.to_datetime('2016-3'),0.25,5)
value_set
