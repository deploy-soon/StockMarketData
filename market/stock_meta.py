import os
import sys
import time
import json
import requests
import datetime
import win32com.client
from os.path import join as pjoin

sys.path.append("../")
from config import Config
sys.path.append("../tools")
from misc import get_logger
from login import Status


class StockMeta(Status):
    """
    Stock Meta Info structure
    {
        "code" : {
            "name": stock name,
            "mememin": ,
            "marketkind": ,
            "sectionkind": ,
            "listeddate": ,
        }, ...
    }
    :key mememin: 주식 매매 거래 단위 주식수
    :key marketkind: 소속부
        0: 구분없음
        1: 거래소
        2: 코스닥
        3: K-OTC
        4: KRX
        5: KONEX
    :key sectionkind:
        deprecated
    :key listeddate: 상장일
    """

    def __init__(self, **kwargs):
        Status.__init__(self, conf={}, **kwargs)
        self.res_path = "./res"
        self.logger = get_logger()
        self.code = None

    def get_dispatch(self):
        self.assert_disconnect()
        self.code = win32com.client.Dispatch("CpUtil.CpCodeMgr")

    def save(self, info):
        with open(pjoin(self.res_path, "stock_meta.json"), 'w', newline='') as fout:
            json.dump(info, fout)

    def run(self):
        self.get_dispatch()
        kospi = self.code.GetStockListByMarket(1)
        kosdaq = self.code.GetStockListByMarket(2)
        self.logger.info("get {} stocks".format(len(kospi) + len(kosdaq)))

        info = {}
        for stock in kospi + kosdaq:
            info[stock] = dict(
                name=self.code.CodeToName(stock),
                mememin=self.code.GetStockMemeMin(stock),
                marketkind=self.code.GetStockMarketKind(stock),
                sectionkind=self.code.GetStockSectionKind(stock),
                listeddate=self.code.GetStockListedDate(stock),
            )
        self.save(info)

if __name__ == "__main__":
    sm = StockMeta()
    sm.run()