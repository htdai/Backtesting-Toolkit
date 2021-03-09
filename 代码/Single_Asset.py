# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np


class Single_Asset:
    def __init__(self, ann: int, rf: float, input_path: str, input_file: str, output_path: str):
        """
        Initialize a backtester for one single asset
        Think of it as sth that takes in a NAV series and spits out several stats
        Can't further impose positions on the NAV series
        :param int ann: number of days used to annualize statistics, e.g. 250 or 252
        :param float rf: risk-free rate
        :param str input_path: path of the folder where NAV (or closing) spreadsheet is stored
        :param str input_file: file name of the NAV (or closing) spreadsheet
        :param str output_path: path of the output folder
        """
        self.ann = ann
        self.rf = rf
        self.input_path = input_path
        self.input_file = input_file
        self.output_path = output_path
        self.data = None

    def read_sheet(self, sheet_name: str):
        """
        Read the Excel file
        :param str sheet_name: sheet name of NAV series
        """
        self.data = pd.read_excel(self.input_path+self.input_file, sheet_name=sheet_name, index_col=0)

    def backtest_whole_period(self, nav_series: pd.Series):
        """
        Backtest the given NAV series, regardless of its length
        So that this method can be used to both backtest the entire period, as well as backtest by year as long as nav_series is properly sliced
        :param pd.Series nav_series: nav series used to calculate stats
        :return pd.DataFrame: a DataFrame of stats
        """
        ret_series = nav_series.pct_change()

        # --------------------------------------------------------------------------------------------------------------
        # Calculate return and stdev
        holding_period_return = nav_series.iloc[-1] / nav_series.iloc[0] - 1
        annualized_return = (holding_period_return + 1) ** (self.ann / (len(ret_series) - 1)) - 1
        # should take ((len(ret_series) - 1) / self.ann)-th root of hpr, i.e. raise it to the (self.ann / (len(ret_series) - 1))-th power
        # len(ret_series) needs to minus 1 because of the initial np.nan as a result of pct_change()
        annualized_stdev = np.nanstd(ret_series, ddof=1) * np.sqrt(self.ann)

        # --------------------------------------------------------------------------------------------------------------
        # Calculate mdd and ratios
        mdd, mdd_start, mdd_formation = self.mdd(nav_series)

        sharpe, calmar = (annualized_return - self.rf) / annualized_stdev, (annualized_return - self.rf) / mdd

        # --------------------------------------------------------------------------------------------------------------
        # Store results to the dataframe
        df = pd.DataFrame(
            {'标的总收益': [holding_period_return], '年化收益率': [annualized_return],
             '年化波动率': [annualized_stdev], '夏普比率': [sharpe], '卡玛比率': [calmar],
             '最大回撤': [mdd],
             '最大回撤起始时间': [mdd_start], '最大回撤形成时间': [mdd_formation], '最大回撤恢复时间': [np.nan]
             })

        return df

    def backtest(self, asset_name: str):
        """
        Run backtest on the selected asset
        :param str asset_name: name of the asset, used to locate the column among many assets
        """
        if asset_name not in self.data.columns:
            raise ValueError('invalid asset name')
        nav_series = self.data[asset_name].dropna()
        df_list = []

        # --------------------------------------------------------------------------------------------------------------
        # Backtest the entire period
        df_all = self.backtest_whole_period(nav_series)
        try:
            df_all['最大回撤恢复时间'] = nav_series.loc[(nav_series >= nav_series.loc[df_all['最大回撤起始时间']][0]) & (pd.to_datetime(nav_series.index) > pd.to_datetime(df_all['最大回撤起始时间'][0]))].index[0].date()
        except:
            df_all['最大回撤恢复时间'] = '尚未恢复'
        df_all.index = ['整体表现']
        df_list.append(df_all)

        # --------------------------------------------------------------------------------------------------------------
        # Backtest by year
        years = list(set(nav_series.index.year))
        years.sort()
        for idx, year in enumerate(years):
            nav_series_by_year = nav_series.loc[nav_series.index.year == year]
            if idx == 0:
                if len(nav_series_by_year) == 1:
                    continue
            else:
                last_year_close = pd.Series([nav_series.loc[nav_series.index.year == years[idx - 1]].iloc[-1]])
                last_year_close.index = [nav_series.loc[nav_series.index.year == years[idx - 1]].index[-1]]
                last_year_close.name = nav_series_by_year.name
                nav_series_by_year = last_year_close.append(nav_series_by_year)
            df_by_year = self.backtest_whole_period(nav_series_by_year)
            try:
                df_by_year['最大回撤恢复时间'] = nav_series.loc[
                    (nav_series >= nav_series.loc[df_by_year['最大回撤起始时间']][0]) & (pd.to_datetime(nav_series.index) > pd.to_datetime(df_by_year['最大回撤起始时间'][0]))].index[0].date()
            except:
                df_by_year['最大回撤恢复时间'] = '尚未恢复'
            df_by_year.index = [year]
            df_by_year['年化收益率'] = df_by_year['标的总收益']
            df_list.append(df_by_year)

        # --------------------------------------------------------------------------------------------------------------
        # Concatenate results to get one holistic DataFrame
        df = pd.concat(df_list)
        self.output(df, asset_name)

    def mdd(self, nav_series):
        """
        Calculate maximum drawdown using the given NAV series
        :param pd.Series nav_series: NAV series used to calculate mdd stats
        :return : stats
        """
        dd = nav_series.div(nav_series.cummax()).sub(1)
        # NAV divided by its cumulative maximum then subtracted by 1 gives the drawdown series
        mdd, formation = dd.min(), dd.idxmin()
        formation = formation.date()
        start = nav_series.loc[:formation].idxmax()
        start = start.date()
        return -mdd, start, formation

    def output(self, df, asset_name):
        """
        Save results as an Excel file
        :param pd.DataFrame df: result sheet for both the entire period and by year
        :param str asset_name: name of the asset
        """

        writer = pd.ExcelWriter(self.output_path + '%s回测结果.xlsx' % asset_name,
                                engine='xlsxwriter')
        df.to_excel(writer, sheet_name='%s回测结果' % asset_name)
        writer.save()


if __name__ == '__main__':
    a = Single_Asset(ann=250, rf=0, input_path=r'E:/College/Gap/Huatai/Misc/20210308回测框架v5.0/数据/', input_file='data.xlsx',
                     output_path=r'E:/College/Gap/Huatai/Misc/20210308回测框架v5.0/输出/')


