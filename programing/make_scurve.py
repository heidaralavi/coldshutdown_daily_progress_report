import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt

def convert_to_date(df,col_name):
    df[col_name] = pd.to_datetime(df[col_name],)
    df[col_name]=df[col_name].dt.date
    return df
    
def list_to_str(l):
    listToStr = []
    for item in l:
        listToStr.append(str(item))
    return listToStr

def setup_df(df):
    df['%man_hours']=round(df['MAN__HOUR']/df.MAN__HOUR.sum()*100,3)
    df['duration']=df['Finish_Date']-df['Start_Date']+timedelta(days=1)
    df['day_weigth']=df['%man_hours']/(df['duration'].dt.days)
    return df



def make_master_df(df):
    w_orders= pd.read_excel('w_orders.xlsx')
    mylist = w_orders['w_orders'].tolist()
    mylist = list_to_str(mylist)
    df = convert_to_date(df, 'Start_Date')
    df = convert_to_date(df, 'Finish_Date')
    df = df[df['WRNO'].isin(mylist)]
    df=df.reset_index(drop=True)
    df= setup_df(df)
    return df


def make_master_progress(df):
    w_orders= pd.read_excel('w_orders.xlsx')
    mylist = w_orders['w_orders'].tolist()
    #mylist = list_to_str(mylist)
    df = df[df['WRNO'].isin(mylist)]
    df=df.reset_index(drop=True)
    return df




    

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


def plot_df(plt_data):
    fix, ax = plt.subplots(figsize=(10,10),dpi=300)
    ax.plot(plt_data['cumsum'])
    ax.set_xticks(range(0,32,1))
    plt.xticks(rotation=90)
    
    plt.tight_layout()
    plt.savefig('foo.png')
    plt.show()
    


df= pd.read_excel('data.xlsx')
df = make_master_df(df)
df = add_daily_chart(df)
plt_data=make_plotdata(df)
plot_df(plt_data)
prog = pd.read_excel('daily-progress.xlsx')
prog = make_master_progress(prog)

print(prog.head())
prog.to_excel('out.xlsx')



