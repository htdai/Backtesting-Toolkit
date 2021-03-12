# -*- coding: utf-8 -*-
from Codes.Portfolio import Portfolio
from Codes.Single_Asset import Single_Asset
import pandas as pd


if __name__ == "__main__":
    # ------------------------------------------------------------------------------------------------------------------
    # Initialize parameters
    ann = 250
    rf = 0
    start_date = None
    end_date = None
    input_path = r'Data\data.xlsx'
    output_path_single_asset = r'Output\单资产回测结果.xlsx'
    output_path_portfolio = r'Output\组合回测结果.xlsx'

    high_risk_name_list = ['沪深300', '中证500', '创业板指', '南华商品指数']
    high_risk_fee_rate = 0.0003
    low_risk_name_list = ['中债-总财富(总值)指数', '中债-信用债总财富(总值)指数']
    low_risk_fee_rate = 0.0002

    # ------------------------------------------------------------------------------------------------------------------
    # Single asset backtesting
    single_asset = Single_Asset(ann=ann, rf=rf)
    single_asset.load_sheet_from_file(input_path=input_path, sheet_name='数据')
    single_asset.slice(start_date, end_date)
    for asset in single_asset.data.columns:
        single_asset.backtest(asset)
    single_asset.output(output_path=output_path_single_asset, asset_name_list=list(single_asset.data.columns))

    # ------------------------------------------------------------------------------------------------------------------
    # Portfolio backtesting
    portfolio = Portfolio(ann=ann, rf=rf)
    portfolio.load_sheets_from_file(input_path=input_path, data_sheet_name='数据', weight_sheet_name='权重')
    portfolio.load_fee_rates(high_risk_name_list=high_risk_name_list, high_risk_fee_rate=high_risk_fee_rate,
                             low_risk_name_list=low_risk_name_list, low_risk_fee_rate=low_risk_fee_rate)
    portfolio.slice(start_date, end_date)
    portfolio.generate_nav()
    portfolio.backtest()
    portfolio.output(output_path=output_path_portfolio)
