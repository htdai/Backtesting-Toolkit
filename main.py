# -*- coding: utf-8 -*-
from Codes.Portfolio import Portfolio
from Codes.Single_Asset import Single
import pandas as pd


def main():
    """
    ann : 年化时考虑的天数
    start, end : 策略开始和结束时间
    fee_rate : 交易费率
    rf : 无风险利率
    input_path : 输入文件夹地址
    file : 输入数据文件名
    output_path : 输出文件夹地址
    asset_file : 资产组合输出文件名
    single_file : 单一资产输出文件名
    single : 单一资产名字 
    """
    ann = 250
    start = pd.datetime(2010,7,29)
    end = pd.datetime(2021,2,3)
    fee_rate = 0#0.0003
    rf = 0
    input_path = r'Data/'
    file = r'data.xlsx'
    output_path = r'Output/'
    asset_file = r'资产组合表现.xlsx'
    single_file = r'单一资产表现.xlsx'
    single = r'沪深300'
    
    pb = Portfolio(ann, start, end, fee_rate, rf, input_path, file,
                   output_path, asset_file, single)
    pb.portfolio_backtest()
    
    #sb = Single(ann, single, rf, input_path, file, output_path, single_file)
    #sb.single_backtest()


if __name__ == "__main__":
    main()
