import os
import sys
import csv
import fire
import time
import json
import tqdm
import h5py
import win32com.client
from os.path import join as pjoin

sys.path.append("../tools")
from misc import get_logger
from login import Status


class Minute(Status):

    def __init__(self, conf="./config/stock_minute.json", **kwargs):
        Status.__init__(self, conf=conf, **kwargs)
        self.logger = get_logger()
        if not os.path.isdir(self.opt.export_to):
            os.mkdir(self.opt.export_to)

        self.stock_chart = None
        self.stock_code = None

    def get_dispatch(self):
        self.assert_disconnect()
        self.stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")
        self.stock_code = win32com.client.Dispatch("CpUtil.CpCodeMgr")

    def get_stockcode(self):
        codeList = self.stock_code.GetStockListByMarket(1)
        codeList2 = self.stock_code.GetStockListByMarket(2)
        stockcodes = codeList + codeList2
        # self.stocknames = {}
        # for i, code in enumerate(codeList + codeList2):
        #     # section = self.stockcode.GetStockSectionKind(code)
        #     # if section != 1:
        #     #     # not corporation such as ETF, futures, ...
        #     #     continue
        #     name = self.stock_code.CodeToName(code)
        #     stockcodes.append(code)
        #     self.stocknames[code] = name
        self.logger.info("Get stock list : {}".format(len(stockcodes)))
        return stockcodes

    def log_request(self):
        code = self.stock_chart.GetDibStatus()
        message = self.stock_chart.GetDibMsg1()
        if code != 0:
            self.logger.warning("code : {}, message : {}".format(code, message))
            raise
        if self.verbose:
            self.logger.info("code : {}, message : {}".format(code, message))

    def get_data(self, stock_code):
        self.stock_chart.SetInputValue(0, stock_code)
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(4, 100000)
        self.stock_chart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
        self.stock_chart.SetInputValue(6, ord("m"))
        self.stock_chart.SetInputValue(9, ord('1'))
        self.stock_chart.SetInputValue(10, ord('3'))
        self.stock_chart.BlockRequest()
        length = self.stock_chart.GetHeaderValue(3)
        dates, minutes, opens, highs, lows, closes, volumes = [], [], [], [], [], [], []
        date_set = set()
        while self.stock_chart.Continue:
            for i in range(length):
                if self.opt.todate > self.stock_chart.GetDataValue(0, i):
                    break
                _date = self.stock_chart.GetDataValue(0, i)
                dates.append(self.stock_chart.GetDataValue(0, i))
                minutes.append(self.stock_chart.GetDataValue(1, i))
                opens.append(self.stock_chart.GetDataValue(2, i))
                highs.append(self.stock_chart.GetDataValue(3, i))
                lows.append(self.stock_chart.GetDataValue(4, i))
                closes.append(self.stock_chart.GetDataValue(5, i))
                volumes.append(self.stock_chart.GetDataValue(6, i))
                date_set.add(_date)
            if self.status.getLimitRemainCount(1) < 2:
                time.sleep(15.0)
            self.stock_chart.BlockRequest()
            self.log_request()
            length = self.stock_chart.GetHeaderValue(3)
        #self.logger.info("get date from {} to {}".format(min(date_set), max(date_set)))
        # for time forward
        for list_data in [dates, minutes, opens, highs, lows, closes, volumes]:
            list_data.reverse()
        return {"dates": dates, "minutes": minutes, "opens": opens, "highs": highs,
                "lows": lows, "closes": closes, "volumes": volumes}

    def run(self):
        self.get_dispatch()
        stock_codes = self.get_stockcode()
        with h5py.File(pjoin(self.opt.export_to, "minute_data.h5"), "w") as fout:
            for stock_code in tqdm.tqdm(stock_codes):
                minute_data = self.get_data(stock_code)
                if not minute_data:
                    continue
                stockgroup = fout.create_group(stock_code)
                stockgroup.create_dataset("dates", data=minute_data["dates"])
                stockgroup.create_dataset("minutes", data=minute_data["minutes"])
                stockgroup.create_dataset("opens", data=minute_data["opens"])
                stockgroup.create_dataset("highs", data=minute_data["highs"])
                stockgroup.create_dataset("lows", data=minute_data["lows"])
                stockgroup.create_dataset("closes", data=minute_data["closes"])
                stockgroup.create_dataset("volumes", data=minute_data["volumes"])
        with open(pjoin(self.opt.export_to, "minute_data.keys"), "w") as fout:
            fout.write("\n".join(stock_codes))

def load():
    with h5py.File(pjoin("minute_to", "minute_data.h5"), "r") as fin:
        for k in fin.keys():
            print(k)
            print(fin[k]["dates"])

if __name__ == "__main__":
    minute = Minute()
    minute.run()
    #load()
