import pandas as pd
from scipy.optimize import newton
from .Single_Asset import Single_Asset
# Relative import used here because Portfolio is imported in main.py
# If running Portfolio.py independently, remove the relative import dot above
import matplotlib.pyplot as plt
import warnings

warnings.filterwarnings('ignore')


class Portfolio:
    def __init__(self, ann: int, rf: float, data=None, weight=None):
        """
        Initialize a backtester for a portfolio
        :param int ann: number of days used to annualize statistics, e.g. 250 or 252
        :param float rf: risk-free rate
        :param pd.DataFrame data: closing price series, so that you can also use this backtester after some other
        Python programs
        without loading a local Excel file
        :param pd.DataFrame weight: asset weight series, so that you can also use this backtester after some other
        Python programs
        without loading a local Excel file
        """
        self.ann = ann
        self.rf = rf

        self.input_path = self.output_path = None
        self.data = data
        self.weight = weight

        self.high_risk_name_list = self.high_risk_fee_rate = self.low_risk_name_list = self.low_risk_fee_rate = None

        self.backtest_results = dict()

    def load_sheets_from_file(self, input_path: str, data_sheet_name='数据', weight_sheet_name='权重'):
        """
        Load closing price series and weight series data from a local Excel file
        :param str input_path: file path of the Excel file
        :param str data_sheet_name: name of the sheet containing closing price series, default 数据
        :param str data_sheet_name: name of the sheet containing weight series, default 权重
        """
        self.input_path = input_path
        self.data = pd.read_excel(self.input_path, sheet_name=data_sheet_name, index_col=0)
        self.data.sort_index(ascending=True, inplace=True)
        self.weight = pd.read_excel(self.input_path, sheet_name=weight_sheet_name, index_col=0)
        self.weight.sort_index(ascending=True, inplace=True)

        self.data = self.data[self.weight.columns]
        # only keep columns in self.data that has weight information; hence if not loading closing price and weight
        # series from an local Excel file, the user will need to do this explicitly before initializing a backtester
        # now columns in self.data also aligns with those in self.weight

    def load_fee_rates(self, high_risk_name_list=None, high_risk_fee_rate=None, low_risk_name_list=None,
                       low_risk_fee_rate=None):
        """
        Specify high and low risk assets as well as applicable fee rates respectively
        :param list high_risk_name_list: list of names of high risk assets
        :param float high_risk_fee_rate: fee rate for high risk assets
        :param list low_risk_name_list: list of names of low risk assets (excluding cash)
        :param float low_risk_fee_rate: fee rate for low risk assets (excluding cash)
        """

        duplicate_assets = list(set(high_risk_name_list) & set(low_risk_name_list))
        if len(duplicate_assets) > 0:
            raise ValueError('assets found in both high and low risk asset name lists: %s.' %
                             ', '.join(duplicate_assets))

        unspecified_assets = list(set(self.weight.columns) - (set(high_risk_name_list) | set(low_risk_name_list)))
        if len(unspecified_assets) > 0:
            raise ValueError('risk level unspecified for assets: %s.' % ', '.join(unspecified_assets))

        self.high_risk_name_list, self.high_risk_fee_rate = high_risk_name_list, high_risk_fee_rate
        self.low_risk_name_list, self.low_risk_fee_rate = low_risk_name_list, low_risk_fee_rate

    def slice(self, start_date=None, end_date=None):
        """
        Slice the closing price series data based on desired start and end, and process both closing price and weight
        dataframes
        It actually does more than merely slicing, but the method is still named "slice" to mirror that in Single_Asset
        :param str start_date: desired start date
        :param str end_date: desired end date
        """
        # --------------------------------------------------------------------------------------------------------------
        # Step 1: general slicing

        self.data = self.data.loc[pd.to_datetime(start_date):pd.to_datetime(end_date)]
        # Remaining slice of NAV series falls between start date and end date inclusive, does NOT necessarily contain
        # start date or end date depending on its original input

        self.weight = self.weight.loc[self.data.index[0]: self.data.index[-1]]
        # any weight info outside of the time span of NAV is useless

        # --------------------------------------------------------------------------------------------------------------
        # Step 2: adjust the weight dataframe
        # Backtesting is essentially imposing some sort of weighting structure on those NAV series, hence we adjust
        # the weight dataframe s.t. its index is a subset of that of the NAV dataframe

        date_adjusted_weight = pd.DataFrame(columns=self.weight.columns)

        for date in self.weight.index:  # iterate all dates in the weight dataframe

            if date in self.data.index:    # if date exists in the NAV dataframe, it can be safely kept
                date_adjusted_weight.loc[date] = self.weight.loc[date]

            else:   # if date does NOT exist in the NAV dataframe...
                nearest_closing_date = pd.Series(self.data.index, index=self.data.index).loc[:date].iloc[-1]

                if nearest_closing_date in self.weight.index:
                    # but weight information for the nearest closing date is already specified...
                    continue    # then we simply discard it

                else:
                    # and weight information for the nearest closing date is NOT specified...
                    date_adjusted_weight.loc[nearest_closing_date] = self.weight.loc[date]
                    # then we move it to the nearest closing date, based on consideration that this is desired for
                    # scenarios like monthly backtesting where weight series are indexed by last calendar dates of each
                    # month instead of last trading date

        # --------------------------------------------------------------------------------------------------------------
        # Step 3: final processing
        # Now that the weight dataframe's index is a subset of the NAV dataframe's index, we slice the NAV dataframe
        # again s.t. their indices agree on the initial start date, but not necessarily the end date because you can
        # still backtest your portfolio for as long as NAV series allow after the last rebalancing
        date_adjusted_weight.sort_index(ascending=True, inplace=True)
        self.weight = date_adjusted_weight
        self.data = self.data.loc[self.weight.index[0]:]

    def calculate_fee(self, sb: pd.Series, sa: pd.Series, f: pd.Series, pa: pd.Series) -> float:
        """
        Calculate fee incurred for a given rebalancing
        :param pd.Series sb: shares right before rebalancing
        :param pd.Series sa: shares right after rebalancing
        :param pd.Series f: fee rate vector
        :param pd.Series pa: closing prices right after rebalancing, i.e. the prices you use to rebalance
        :return float: fee incurred for the given rebalancing
        """
        return sum(abs(sb - sa) * f * pa)

    def generate_nav(self):
        """
        Generate NAV series of the portfolio based on closing prices and
        """
        # --------------------------------------------------------------------------------------------------------------
        # Step 1: generate a bunch of dataframes and series to store various results
        data = self.data / self.data.iloc[0]
        # normalize intial prices to 1 for the concern of floating point number precision
        input_weight = self.weight

        portfolio_stats = pd.DataFrame(index=data.index)
        shares = pd.DataFrame(index=data.index, columns=data.columns)
        actual_weight = pd.DataFrame(index=data.index, columns=data.columns)
        turnover_ratio = pd.DataFrame(index=data.index, columns=data.columns)
        fee = pd.Series(index=data.index)
        fee_rate = pd.Series(index=data.columns)
        nav = pd.Series(index=data.index)

        # --------------------------------------------------------------------------------------------------------------
        # Step 2: fill in the fee rate vector, then define a function for root solving NAV after rebalancing
        for asset in fee_rate.index:
            if asset in self.high_risk_name_list:
                fee_rate.loc[asset] = self.high_risk_fee_rate
            elif asset in self.low_risk_name_list:
                fee_rate.loc[asset] = self.low_risk_fee_rate

        def nav_equation(x: float, sb: pd.Series, wa: pd.Series, pa: pd.Series, f: pd.Series, navb: float):
            """
            Solve the NAV right after rebalancing using a fundamental relationship that the change in NAV should equal
            fee incurred; total fee is simply the sum over all assets; and for each asset its fee is calculated as
            its fee rate times the price you rebalance at times the absolute change of shares
            :param pd.Series sb: shares right before rebalancing
            :param pd.Series wa: weight right after rebalancing
            :param pd.Series pa: closing prices right after rebalancing, i.e. the prices you use to rebalance
            :param pd.Series f: fee rate vector
            :param float navb: NAV right before rebalancing
            :return float: return the value of the root solving function, i.e. x is the solution when this function
            returns 0
            """
            return self.calculate_fee(sb=sb, sa=(wa/pa)*x, f=f, pa=pa) - navb + x

        # --------------------------------------------------------------------------------------------------------------
        # Step 3: backtest the portfolio, calculate detailed statistics
        for idx, date in enumerate(data.index):

            if idx == 0:    # for the start date...
                actual_weight.loc[date] = input_weight.loc[date]    # actual weight equals input weight
                nav.loc[date] = 1.0    # normalize NAV to 1
                shares.loc[date] = actual_weight.loc[date] / data.loc[date]
                # value over price equals amt, nav of 1 omitted
                fee.loc[date] = self.calculate_fee(sb=pd.Series([0] * len(data.columns), index=data.columns),
                                                   sa=shares.loc[date], f=fee_rate, pa=data.loc[date])
                turnover_ratio.loc[date] = abs(input_weight.loc[date])

            else:   # for the remaining dates...

                if date not in input_weight.index:    # and if it's not a rebalancing date...
                    shares.loc[date] = shares.loc[data.index[idx - 1]]
                    # number of shares unchanged, i.e. NOT rebalanced
                    fee.loc[date] = 0   # obvsly no fee incurred, since no rebalancing
                    turnover_ratio.loc[date] = 0    # obvsly turnover is 0, since no rebalancing
                    nav.loc[date] = nav.loc[data.index[idx - 1]] + sum(shares.loc[date] *
                                                                       (data.loc[date] - data.loc[data.index[idx - 1]]))
                    # NAV's increase comes from sum of price increment times shares held over assets
                    actual_weight.loc[date] = shares.loc[date] * data.loc[date] / nav.loc[date]
                    # actual weight by definition is value of each asset over NAV

                else:   # and it is a rebalancing date...
                    shares_before = shares.loc[data.index[idx - 1]]
                    # shares right before rebalancing equals previous date
                    nav_before = nav.loc[data.index[idx - 1]] + sum(shares_before * (data.loc[date] -
                                                                                     data.loc[data.index[idx - 1]]))
                    # NAV's increase comes from sum of price increment times shares held over assets
                    nav_after = newton(nav_equation, x0=nav_before, args=(shares_before, input_weight.loc[date],
                                                                          data.loc[date], fee_rate, nav_before),
                                       x1=nav_before*max(self.high_risk_fee_rate, self.low_risk_fee_rate))
                    # Use the solver to find NAV after rebalancing s.t. weight after rebalancing is as wanted and fee is
                    # subtracted from NAV
                    fee.loc[date] = nav_before - nav_after
                    nav.loc[date] = nav_after
                    actual_weight.loc[date] = input_weight.loc[date]
                    shares.loc[date] = actual_weight.loc[date] * nav_after / data.loc[date]
                    turnover_ratio.loc[date] = abs(input_weight.loc[date] -
                                                   actual_weight.loc[actual_weight.index[idx - 1]])

        # --------------------------------------------------------------------------------------------------------------
        # Step 4: organize the results
        portfolio_stats['组合净值'] = nav
        portfolio_stats['交易费用'] = fee
        self.backtest_results['回测结果汇总'] = None
        self.backtest_results['组合净值和交易费用'] = portfolio_stats
        self.backtest_results['归一化资产价格'] = data
        self.backtest_results['资产持有股数（对应归一化资产价格）'] = shares
        self.backtest_results['资产权重'] = actual_weight
        self.backtest_results['资产调仓目标'] = input_weight
        self.backtest_results['资产换手率'] = turnover_ratio
    
    def backtest(self):
        nav_backtest = Single_Asset(ann=self.ann, rf=self.rf,
                                    data=pd.DataFrame(self.backtest_results['组合净值和交易费用']['组合净值']))
        nav_backtest.backtest('组合净值')
        self.backtest_results['回测结果汇总'] = nav_backtest.backtest_results['组合净值']

        # --------------------------------------------------------------------------------------------------------------
        # Calculate turnover ratio
        years = list(set(self.backtest_results['资产换手率'].index.year))
        years.sort()
        self.backtest_results['回测结果汇总']['组合换手率'] = 0
        for asset in self.backtest_results['资产换手率'].columns:
            self.backtest_results['回测结果汇总'].loc['整体表现', asset + '换手率'] =\
                sum(self.backtest_results['资产换手率'][asset])
            self.backtest_results['回测结果汇总'].loc['整体表现', '组合换手率'] +=\
                self.backtest_results['回测结果汇总'].loc['整体表现', asset + '换手率']
            for year in years:
                self.backtest_results['回测结果汇总'].loc[year, asset + '换手率'] =\
                    sum(self.backtest_results['资产换手率'].loc[self.backtest_results['资产换手率'].index.year == year,
                                                           asset])
                self.backtest_results['回测结果汇总'].loc[year, '组合换手率'] +=\
                    self.backtest_results['回测结果汇总'].loc[year, asset + '换手率']

    def output(self, output_path: str):
        """
        Save results as an Excel file
        :param str output_path: desired file path of the Excel file containing backtest results
        """
        writer = pd.ExcelWriter(path=output_path)
        for key in self.backtest_results.keys():
            temp = self.backtest_results[key]
            if key != '回测结果汇总':
                temp.index = temp.index.date
            temp.to_excel(writer, sheet_name=key)
        writer.save()

    """
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
        # 读取数据
        data, weight = self.sheet_read()
        data = self.data_process(data, weight)
        # 处理调仓日非交易日，生成workday序列
        # 根据策略起始时间对data数据切片
        workday = self.time_process(weight)
        weight.index = workday
        weight = weight[~weight.index.isnull()]

        net, asset_weight, asset_share= self.net_worth(data, weight, workday)
        d = pd.DataFrame.copy(data, deep = True)
        d = self.data_clean(d, net)
        gb = d.groupby('years')
        
        # 计算分年度的表现，并输出excel
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
        
        # 计算整体表现，并输出excel
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
        
        # 绘制净值图
        self.output(data, net, asset_weight, df_year, df_all)
    """


if __name__ == "__main__":
    ann = 250
    rf = 0
    input_path = r'..\测试\05带杠杆和做空\data.xlsx'
    output_path = r'..\测试\05带杠杆和做空\组合回测结果.xlsx'
    high_risk_name_list = ['沪深300', '中证500', '创业板指', '南华商品指数']
    high_risk_fee_rate = 0.0003
    low_risk_name_list = ['中债-总财富(总值)指数', '中债-信用债总财富(总值)指数']
    low_risk_fee_rate = 0.0002

    pb = Portfolio(ann=ann, rf=rf)
    pb.load_sheets_from_file(input_path=input_path, data_sheet_name='数据', weight_sheet_name='权重')
    pb.load_fee_rates(high_risk_name_list=high_risk_name_list, high_risk_fee_rate=high_risk_fee_rate,
                      low_risk_name_list=low_risk_name_list, low_risk_fee_rate=low_risk_fee_rate)
    start_date = None
    end_date = None
    pb.slice(start_date, end_date)
    pb.generate_nav()
    pb.backtest()
    pb.output(output_path=output_path)
