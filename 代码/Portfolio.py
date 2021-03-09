# -*- coding: utf-8 -*-
"""
Created on Fri Jan  8 14:32:12 2021

@author: Muffler
"""
import pandas as pd
import numpy as np
from chinese_calendar import is_workday, is_holiday
from datetime import timedelta
import matplotlib.pyplot as plt
import warnings

warnings.filterwarnings('ignore')

class Portfolio:
    def __init__(self, ann, start, end, fee_rate, rf, input_path, file,
                 output_path, asset_file, single):
        self.ann = ann
        self.start = start
        self.end = end
        self.fee_rate = fee_rate
        self.rf = rf
        self.input_path = input_path
        self.file = file
        self.output_path = output_path
        self.asset_file = asset_file
        self.single = single
        
    #读取excel
    def sheet_read(self):    
        data = pd.read_excel(self.input_path+self.file, sheet_name = '数据', 
                             index_col = 0)
        weight = pd.read_excel(self.input_path+self.file, sheet_name = '权重', 
                               index_col = 0)
        return data, weight     
    
    #根据要求截取数据
    def data_process(self, data, weight):

        data = data[weight.columns]
        data = data.dropna()
        
        data = data[data.index >= self.start]
        data = data[data.index <= self.end]
        return data
    
    #生成日收益率、净值、years列
    def data_clean(self, data, net):

        networth = net
        data['daily'] = networth.diff()/networth.shift(1)
        data['net_worth'] = networth
        #生成years列，便于后续使用groupby
        years = data.index.year
        years = pd.DataFrame(years)
        years.index = data.index
        data['years'] = data.index.year
  
        return data
    
    #交易日处理    
    def time_process(self, weight):
        #计算策略期中的交易日，返回为list
        l = []
        temp = self.start
        while True:
            if temp > self.end:
                break
            if is_workday(temp)==False or is_holiday(temp)==True \
                                       or temp.weekday()>4:
                l.append(temp)
            temp += timedelta(days=1)
                    
        workday1 = weight[~weight.index.isin(l)]
        #调整非交易日的日期
        otherday = weight[weight.index.isin(l)]
        o = []
        for i in otherday.index:
            while is_workday(i)==False or is_holiday(i)==True or i.weekday()>4:
                i +=timedelta(days=1)
                o.append(i)
        o = pd.Series(o)
        workday2 = o[~o.isin(l)]
        workday = pd.concat([pd.Series(workday1.index), workday2])
        workday = workday.sort_values(ascending=True).reset_index(drop = True)
        for i in range(len(workday)):
            if workday[i] < self.start and workday[i+1] > self.start:
                workday[i] = self.start
            elif workday[i]< self.start:
                workday[i] = np.nan
            elif workday[i] > self.end:
                workday[i] = np.nan
        
        return workday
        
    
    #计算日收益率    
    def dailyR(self, data):    
        daily = data.diff()/data.shift(1)
        #对数收益率
        #daily = np.log(data/data.shift(1))
        return daily
    
    #计算策略总收益
    def wholeR(self, data):
        p = data['net_worth']
        wr = (p.iloc[-1]-p.iloc[0])/p.iloc[0]
        return wr
    
    #计算年化收益率
    def annualR(self, data, gb, period):
        """period = 'by year' or 'all period' """
        if period == 'by year':
            r_p = gb['net_worth'].apply(lambda x:(x.iloc[-1]-x.iloc[0])/x.iloc[0])
        elif period == 'all period':
            r = self.wholeR(data)
            r_p = ((1+r)**(self.ann/data['daily'].count())-1)
        return r_p
    
    #计算年化波动率
    def annualV(self, data, gb, period):
        """period = 'by year' or 'all period' """
        if period == 'by year':
            s = gb['daily'].apply(np.nanstd)
            V = np.sqrt(self.ann)*s
        elif period == 'all period':
            s = np.nansum((data['daily'].diff())**2)
            V = np.sqrt(self.ann*s/(data['daily'].count()-1))
        return V
    
    #计算历史最大回撤，回撤发生时间和结束时间
    def max_dd(self, returns):
        #计算每天的累计收益
        r = (returns+1).cumprod()
        #r.cummax()计算出累计收益的最大值，再用每天的累计收益除以这个最大值，算出收益率
        dd = r.div(r.cummax()).sub(1)
        #取最小
        mdd = dd.min()
        end = dd.idxmin()
        end = end.date()
        start = r.loc[:end].idxmax()
        start = start.date()
        return -mdd, start, end
    
    #计算夏普比率
    def sharpe(self, data, gb, period):
        """period = 'by year' or 'all period' """
        r_p = self.annualR(data, gb, period)
        V = self.annualV(data, gb, period)
        sharpe =(r_p - self.rf)/V
        if period == 'by year':
            sharpe = pd.DataFrame(sharpe, columns=['夏普比率'])
        return sharpe
    
    #计算卡玛比率
    def carmar(self, data, gb, period):
        """period = 'by year' or 'all period' """
        r_p = self.annualR(data, gb, period)
        if period == 'by year':
            mdd, start, end = zip(*gb['daily'].apply(self.max_dd))
            carmar = r_p/mdd
            carmar = pd.DataFrame(carmar)
            carmar.rename(columns={'net_worth':'carmar比率'},inplace = True)
        elif period == 'all period':
            mdd, start, end = self.max_dd(data['daily'])
            carmar = r_p/mdd
        return carmar
    
    
    #计算每个产品的换手率
    def turnR(self, weight, period):
        """period = 'by year' or 'all period' """
        #修改列名
        columns = []
        for i in weight.columns:
            columns.append(i+'换手率')
        weight.columns = columns
        
        if period =='by year':        
            weight = abs(weight.diff())
            weight['years'] = weight.index.year
            weight = weight.groupby('years').apply(sum)
            weight.drop('years',axis = 1,inplace = True)
        elif period == 'all period':
            weight = abs(weight.diff())
            weight = weight.sum()
            weight = pd.DataFrame(weight)
            weight = weight.T
        return weight
    
    #计算交易费用
    def trade_fee(self, volume):
        fee = volume * self.fee_rate
        return fee
    
    #计算组合净值    
    def net_worth(self, data, weight, workday):
        #构建标的净值dataframe
        price = pd.DataFrame.copy(data, deep = True)
        #向上填充，得到每日持仓权重
        time = pd.DataFrame(data.index, columns = ['日期'], index = data.index)
        w = pd.merge(time, weight, how='outer', left_on = '日期',
                      right_index = True).drop(['日期'], axis = 1)
        w = w.sort_index(ascending=True)
        w = w.fillna(method = 'ffill')
        w = w.dropna()
        #price = price/price.iloc[0]
               
        #计算产品净值
        data['net'] = np.nan
        asset_share = pd.DataFrame(index = data.index,columns =weight.columns)
        asset_weight = pd.DataFrame(index = data.index,columns =weight.columns)

        for j in range(len(data)):
            #建仓的净值为1，所以份额等于cash
            if j ==0:
                data['net'][j] = 1.0
                asset_share.iloc[j] = w.iloc[j]/price.iloc[j]
                asset_weight.iloc[j] = w.iloc[j]
                continue
            elif data.index[j] in list(workday):
                now_value = sum(asset_share.iloc[j-1]*price.iloc[j])
                pre_position = asset_share.iloc[j-1]*price.iloc[j-1]
                now_position = now_value * w.iloc[j]
                trade_volume = now_position - pre_position
                total_volume = sum(abs(trade_volume))
                fee = self.trade_fee(total_volume)
                data['net'][j] = now_value - fee
                asset_share.iloc[j] = data['net'][j]*w.iloc[j]/price.iloc[j]
                asset_weight.iloc[j] = w.iloc[j]
            else:
                asset_share.iloc[j] = asset_share.iloc[j-1]
                data['net'][j] = sum(asset_share.iloc[j]*price.iloc[j])
                asset_weight.iloc[j] = (asset_share.iloc[j]
                                        *price.iloc[j]/data['net'][j])       
        
        net = data['net']
        #计算每日组合内部权重(百分比)
        columns = []
        for i in asset_weight.columns:
            columns.append(i+'每日权重')
        asset_weight.columns = columns
        asset_weight.index = asset_weight.index.date
        asset_weight = asset_weight.applymap(lambda x: '%.2f%%' % (x*100))
        return net, asset_weight, asset_share
    
    #输出策略分年度、整体表现，绘制净值图          
    def output(self, data, net, weight, df_year, df_all):
        networth = net
        networth.name = '资产组合净值'
        price = data
        price = price/price.iloc[0]
        price = pd.concat([price, networth], axis =1)
        price.index = price.index.date
        df = pd.concat([networth, price[self.single]], axis=1)

        plt.rcParams['font.sans-serif']=['SimHei']
        plt.rcParams['axes.unicode_minus'] = False
        color = 'tab:red'
        writer = pd.ExcelWriter(self.output_path + self.asset_file,
                                engine='xlsxwriter')
        df_year.to_excel(writer, sheet_name = '资产组合分年度表现')
        df_all.to_excel(writer, sheet_name = '资产组合整体表现')
        weight.to_excel(writer, sheet_name = '资产组合内每日权重')
        price['资产组合净值'].to_excel(writer, sheet_name = '资产组合每日净值')
        book=writer.book
        
        for i in range(len(price.columns)):
            p = price.iloc[:,i]
            plt.clf()
            p.plot(figsize=(12,6),color = color, linewidth=1.0)
            plt.title(p.name+'净值走势图',size=15)
            plt.savefig(self.output_path+p.name+'净值走势图.png', dpi=300)
            
            sheet=book.add_worksheet(p.name+'净值走势图')
            sheet.insert_image('A1', self.output_path+p.name+'净值走势图.png')
        
        #绘制资产组合和标的对比走势图
        plt.clf()
        #绘制资产组合走势图
        fig, ax = plt.subplots()
        
        df.plot(figsize=(12,6),color = ['tab:red','tab:blue'],rot = 0,
                linewidth=1.0)
        plt.title('资产组合和标的净值对比走势图',size=15)
        ax.set_ylabel('净值',size=15)
        ax.set_xlabel('日期',size=15)
        plt.legend(loc='upper left')        
        
        plt.savefig(self.output_path+'资产组合和标的净值对比走势图.png', dpi=300)
        #输出到excel的sheet中
        sheet=book.add_worksheet('资产组合和标的对比走势图')
        sheet.insert_image('A1', self.output_path+'资产组合和标的净值对比走势图.png')
        
        writer.save()

        
    
    def portfolio_backtest(self):
        #读取数据
        data, weight = self.sheet_read()
        data = self.data_process(data, weight)
        #处理调仓日非交易日，生成workday序列
        #根据策略起始时间对data数据切片
        workday = self.time_process(weight)
        weight.index = workday
        weight = weight[~weight.index.isnull()]

        net, asset_weight, asset_share= self.net_worth(data, weight, workday)
        d = pd.DataFrame.copy(data, deep = True)
        d = self.data_clean(d, net)
        gb = d.groupby('years')
        
        #计算分年度的表现，并输出excel
        year_annualR = self.annualR(d, gb, 'by year')
        year_annualR = year_annualR.apply(lambda x: '%.2f%%' % (x*100))
        year_annualR = pd.DataFrame(year_annualR)
        year_annualR.rename(columns={'net_worth':'年化收益率'},inplace = True)
        
        year_annualV = self.annualV(d, gb, 'by year')
        year_annualV = year_annualV.apply(lambda x: '%.2f%%' % (x*100))
        year_annualV = pd.DataFrame(year_annualV)
        year_annualV.rename(columns={'daily':'年化波动率'},inplace = True)
        
        year_md, year_start, year_end = zip(*gb['daily'].apply(self.max_dd))
        year_md = pd.DataFrame(year_md, index = year_annualR.index,
                               columns = ['最大回撤'])
        year_md = year_md.applymap(lambda x: '%.2f%%' % (x*100))
        year_start = pd.DataFrame(year_start, index = year_annualR.index,
                                  columns = ['最大回撤开始时间'])
        year_end = pd.DataFrame(year_end, index = year_annualR.index,
                                columns = ['最大回撤结束时间'])
        
        year_sharpe = round(self.sharpe(d, gb, 'by year'),2)
        year_carmar = round(self.carmar(d, gb, 'by year'),2)
        year_turnover = self.turnR(weight, 'by year')
        year_turnover = year_turnover.applymap(lambda x: '%.2f%%' % (x*100))
        
        df_year = pd.concat([year_annualR, year_annualV, year_md, year_start, 
                             year_end, year_sharpe, year_carmar, year_turnover
                             ], axis = 1)
        
        #计算整体表现，并输出excel
        all_wholeR = self.wholeR(d)
        all_wholeR = "%.2f%%" % (all_wholeR * 100)
        all_annualR = self.annualR(d, gb, 'all period')
        all_annualR = "%.2f%%" % (all_annualR * 100)
        all_annualV = self.annualV(d, gb, 'all period')
        all_annualV = "%.2f%%" % (all_annualV * 100)
        
        all_md, all_start, all_end = self.max_dd(d['daily'])
        all_md = "%.2f%%" % (all_md * 100)
        
        all_sharpe = self.sharpe(d, gb, 'all period')
        all_carmar = self.carmar(d, gb, 'all period')
        all_turnover = self.turnR(weight, 'all period')
        all_turnover = all_turnover.applymap(lambda x: '%.2f%%' % (x*100))
        
        df_all = pd.DataFrame({'策略总收益' : [all_wholeR], '年化收益率' : 
                               [all_annualR], '年化波动率' : [all_annualV], 
                               '最大回撤' : [all_md], '最大回撤起始时间' : 
                               [all_start], '最大回撤结束时间': [all_end], 
                               '夏普比率' : [all_sharpe], 'carmar比率' : 
                               [all_carmar]})
        df_all = pd.concat([df_all, all_turnover], axis = 1)
        
        df_all.index = ['整体表现']
        
        #绘制净值图
        self.output(data, net, asset_weight, df_year, df_all)

