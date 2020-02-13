import os
import re
import sys
import csv
import pandas as pd
from os.path import join as pjoin

sys.path.append("../tools")
from misc import get_logger

class Report:

    def __init__(self, data_path="data", res_path="res"):
        self.logger = get_logger()
        self.data_path = data_path
        self.res_file = pjoin(res_path, "fnguide_report.csv")
    
    def get_report(self, report_file):
        """
        dateframe columns description
        
        columns
            Market
                시장, KS: kospi, KQ: kosdaq
            Code
                종목코드, A + 6자리
            Company Name
                회사명
            Accounting Standard
                회계기준
            Unit
                단위, 천원
            Total Assets
                자산총계
            Current Assets
                유동자산
            Non-current Assets
                비유동자산
            Other Financial InstitutionsAssets
                기타금융업자산
            Total Liabilities
                부채총계
            Current Liabilities
                유동부채
            Non-Current Liabilites
                비유동부채
            Other Financial Institutions Liabilities
                기타금융업부채
            Total Stockholder's Equity
                자본총계
            Owners of Parent Equity
                지배기업주주지분
            Non-Controlling Interests Equity
                비지배주주지분
            Net Sales
                매출액
            Cost of Sales
                매출원가
            Gross Profit
                매출총이익
            Other Operating Income
                기타영업수익
            Distribution Costs and Administrative Expenses
                판관비
            Other Operating Expenses
                기타영업비용
            Operating Income(Reported)
                영업이익(보고서기재)
            Operating Income(Cal.)
                영업어익(계산수치)
            Other Non-operating Income
                기타영업외수익
            Other Non-operating Expenses
                기타영업외비용
            Financial Income
                금융수익
            Financial Costs
                금융비용
            Gains(Losses) in Subsidiaries, Joint Ventures, Associates
                종속기업등관련이익
            Income Before Income Taxes Expenses
                법인세비용차감전이익
            Income Taxes Expenses
                법인세비용
            Net Profit from Subsidies before Acquisition
                종속회사 매수일전 순손익
            Net Income(Net Loss) from Disposed Subsidies
                처분된 종속회사 순손익
            Ongoing Operating Income
                계속사업이익
            Discontinued Operating Income
                중단사업이익
            Profit
                당기순이익
            Other Comprehensive Income
                기타포괄수익
            Total Comprehensive Income
                총포괄수익
            (Net Income(Net Loss) for The Year Attribute to)Owners of Parent Equity
                (당기순이익귀속)지배기업주주지분
            (Net Income(Net Loss) for The Year Attribute to)Non-Controlling Interests Equity
                (당기순이익귀속)비지배주주기준
            (Total Comprehensive Income Attribute to)Owners of Parent Equity
                (총포괄이익귀속)지배기업주주지분
            (Total Comprehensive Income Attribute to)Non-Controlling Interests Equity
                (총포관이익귀속)비지배주주지분
            Cash Flows from Operatings
                영업활동으로인한현금흐름
            Cash Flows from Investing
                투자활동으로인한현금흐름
            Cash Flows from Financing
                재무활동으로인한현금흐름
        """
        self.logger.info("Start extract excel file: {}".format(pjoin(self.data_path, report_file)))
        regex = re.compile("\d")
        df = pd.read_excel(pjoin(self.data_path, report_file), header = 1)
        columns = [c for c in df.columns if not bool(regex.search(c))]
        df = df[columns]
        new_columns = [column.replace("\n", "").replace(",", ".") for column in columns]
        df.rename(columns = {c: new_c for c, new_c in zip(columns, new_columns)},
                  inplace=True)
        df = df.drop(0, 0)
        df["period"] = report_file
        return df

    def load(self):
        data_file_list = os.listdir(self.data_path)
        df = None
        for data_file in data_file_list:
            if not data_file.isdigit():
                continue
            temp_df = self.get_report(data_file)
            df = temp_df if df is None else df.append(temp_df)
        self.logger.info("Toal Dataframe shape: {}".format(df.shape))
        return df
        
    def save(self, df):
        self.logger.info("save file to {}".format(self.res_file))
        df.to_csv(self.res_file)

    def run(self):
        df = self.load()
        self.save(df)

if __name__ == "__main__":
    report = Report()
    report.run()
