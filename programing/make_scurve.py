import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt
import numpy as np

def convert_to_date(df,col_name):
    df[col_name] = pd.to_datetime(df[col_name],)
    #df[col_name]=df[col_name].dt.date
    return df
    
def list_to_str(l):
    listToStr = []
    for item in l:
        listToStr.append(str(item))
    return listToStr

def setup_df(df):
    df['%man_hours']=round(df['MAN__HOUR']/df.MAN__HOUR.sum()*100,3)
    df['duration']=(df['Finish_Date']-df['Start_Date'])
    
    df.loc[df['duration'].dt.days == 0, ['duration']] = timedelta(days=1)
    df['Start_Date']=df['Start_Date'].dt.strftime('%Y-%m-%d')
    df['Start_Date']=pd.to_datetime(df['Start_Date'],)
    df['Finish_Date']=df['Start_Date']+df['duration']-timedelta(days=1)
    df['day_weigth']=df['%man_hours']/(df['duration'].dt.days)
    return df

def make_master_df(df):
    df = convert_to_date(df, 'Start_Date')
    df = convert_to_date(df, 'Finish_Date')
    df=df.reset_index(drop=True)
    df= setup_df(df)
    return df

def make_master_progress():
    xl = pd.ExcelFile('daily-progress.xlsx')
    sheets=np.array(xl.sheet_names).reshape(-1)
    plt_data = {}
    for sheet in sheets:
        df=pd.read_excel('daily-progress.xlsx',sheet_name=sheet)
        df=filter_by_w_order(df,'W.R.NO')
        df['%man_hours']=round((df['MAN- HOUR']/(df['MAN- HOUR'].sum()))*100,3)
        plt_data[sheet]=((df['ACTUAL%']*df['%man_hours']).sum())/100
    
    
    
    return plt_data

def add_daily_chart(df):
    lists=pd.date_range(df.Start_Date.min(),df.Finish_Date.max(),freq='d').strftime('%Y-%m-%d')
    for l in lists:
        df[l]=0.0
        
    for index, row in df.iterrows():
        date_list = pd.date_range(row['Start_Date'],row['Finish_Date'],freq='d').strftime('%Y-%m-%d')
        for l in date_list:
            df.loc[index,l]=row.day_weigth
            
    return df

def make_plotdata(df):
    lists=pd.date_range(df.Start_Date.min(),df.Finish_Date.max(),freq='d').strftime('%Y-%m-%d')
    dic={}
    for l in lists:
        dic[l]=df[l].sum()
        
        
    plt_df=pd.DataFrame.from_dict(dic,orient='index')
    plt_df['cumsum']=plt_df.cumsum()
    return plt_df


def plot_df(plt_data,plt2):
    fix, ax = plt.subplots(figsize=(10,10),dpi=300)
    ax.plot(plt_data['cumsum'],label='Plan')
    ax.plot(plt2.values(),label='Actual')
    ax.set_xticks(range(0,32,1))
    ax.set_yticks(range(0,105,5))
    plt.xticks(rotation=90)
    plt.legend()
    plt.title('Cold Shutdown 1401',fontsize=30)
    plt.tight_layout()
    plt.grid()
    plt.savefig('foo.png')
    plt.show()
    

def make_prog_plotdata(df):
    plt_data = {}
    dates = pd.date_range('2022-06-22','2022-07-22',freq='d').strftime('%Y-%m-%d')

    for col_names in df.columns:
        if col_names in dates:
            plt_data[col_names]=((df[col_names]*df['%man_hours']).sum())/100
        
    return plt_data


def filter_by_w_order(df,col):
    dic=[]
    w_orders= pd.read_excel('w_orders.xlsx')
    mylist = w_orders['w_orders'].tolist()
    mylist = list_to_str(mylist)
    for item in df[col].to_list():
        cc='no'
        for l in mylist:
            if (l in str(item)):
                cc='yes'
        if cc == 'yes':
            dic.append(True)
        if cc == 'no':
            dic.append(False)
    return df[dic]



df= pd.read_excel('data.xlsx')
df=filter_by_w_order(df,'WRNO')
df = make_master_df(df)
df = add_daily_chart(df)
df.to_excel('out.xlsx')


plt_data=make_plotdata(df)





plt2 = make_master_progress()

#prog.to_excel('out.xlsx')



plot_df(plt_data,plt2)





