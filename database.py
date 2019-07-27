import pandas as pd
import datetime
import timeit
import openpyxl
from datetime import date, timedelta
import numpy as np
import matplotlib.pyplot as plt
import os


def Export(df,filename):
    df.to_excel("out_"+str(filename),encoding="utf-8-sig")
    return

def Import():
    files = [f for f in os.listdir('.') if (os.path.isfile(f) and f[-4:] == 'xlsx')]
    for filename in files:
        try:
            df = pd.read_excel(filename,index_col=0)
            df['Profit Rate'] = df['Profit']/df['Sell Price']
            df[['Sell Price','Profit','Profit Rate']] = df[['Sell Price','Profit','Profit Rate']].fillna(float(0))
            df[['Year','Month','Day']] = df[['Year','Month','Day']].fillna(method='ffill')
            df[['Year']]=df[['Year']].astype(int).astype(str).apply(lambda x: x.str.zfill(4))
            df[['Month','Day']]=df[['Month','Day']].astype(int).astype(str).apply(lambda x: x.str.zfill(2))
            print("Reading file: "+filename)
        except:
            pass
    return df.reset_index(inplace = False,drop = True),filename

def daily_profit(df,day):
    try:
        df = df[df['Year'] == day[0:4]]
        df = df[df['Month'] == day[5:7]]
        df = df[df['Day'] == day[8:10]]
        return sum(df['Profit'])
    except:
        print("Enter format yyyy-mm-dd: ")
        return

def daily_volume(df,day):
    try:
        df = df[df['Year'] == day[0:4]]
        df = df[df['Month'] == day[5:7]]
        df = df[df['Day'] == day[8:10]]
        return sum(df['Sell Price'])
    except:
        print("Enter format yyyy-mm-dd: ")
        return

def daily_profit_rate(df,day):
    try:
        df = df[df['Year'] == day[0:4]]
        df = df[df['Month'] == day[5:7]]
        df = df[df['Day'] == day[8:10]]
        return sum(df['Profit Rate'])
    except:
        print("Enter format yyyy-mm-dd: ")
        return

def daily_summary(df,plot_type):
    summary = []
    day = []
    day_df = pd.DataFrame()
    day_df['Date'] = pd.to_datetime(df[['Year','Month','Day']], format="%Y%m%d")
    length = len(day_df)-1
    start_date = day_df.loc[0]['Date']  # start date
    end_date = day_df.loc[length]['Date'] # end date
    delta = end_date - start_date         # timedelta

    for i in range(delta.days + 1):
        day.append((start_date + timedelta(days=i)).strftime('%Y-%m-%d'))
        if plot_type == 'Profit':
            summary.append((daily_profit(df,(start_date + timedelta(days=i)).strftime('%Y-%m-%d'))))
        elif plot_type == 'Volume':
            summary.append((daily_volume(df,(start_date + timedelta(days=i)).strftime('%Y-%m-%d'))))
        elif plot_type == 'Profit Rate':
            summary.append((daily_profit_rate(df,(start_date + timedelta(days=i)).strftime('%Y-%m-%d'))))
    get_plot(day,summary,plot_type)

def get_plot(day,data,plot_type):
    fig = plt.figure(dpi=200)
    plt.title("Daily "+plot_type)
    plt.xlabel('Date')
    plt.ylabel(plot_type)
    plt.plot(day, data, '--o', alpha = 0.7)
    for x, y in zip(day, data):
        plt.text(x, y, str("%.2f" % y), color="#0abab9", alpha = 0.5, fontsize=10)
    plt.xticks(rotation=30)
    plt.savefig('Daily_'+plot_type+'.png',quality=100,bbox_inches = 'tight')

def main():
    df,filename = Import()
    Export(df,filename)

    daily_summary(df,"Profit")
    daily_summary(df,"Volume")
    daily_summary(df,"Profit Rate")
    input("生成图表完毕，请按任意键退出。")

if __name__ == "__main__":
    main()

