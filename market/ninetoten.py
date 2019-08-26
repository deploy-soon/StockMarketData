import os
import sys
sys.path.append("../tools")
from os.path import join as pjoin
import fire
import tqdm
import h5py
import win32com.client
from misc import get_logger
from login import Status
import time


class NinetoTen(Status):

    def get_dispatch(self):
        self.assert_disconnect()
        self.stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")
        self.stockcode = win32com.client.Dispatch("CpUtil.CpCodeMgr")

    def get_stockcode(self):
        codeList = self.stockcode.GetStockListByMarket(1)
        codeList2 = self.stockcode.GetStockListByMarket(2)
        stockcodes = []
        self.stocknames = {}
        for i, code in enumerate(codeList + codeList2):
            name = self.stockcode.CodeToName(code)
            stockcodes.append(code)
            self.stocknames[code] = name
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

    def get_minute_data(self, stockcode, pivots):
        if not isinstance(pivots, list) or not pivots:
            return {}
        self.stock_chart.SetInputValue(0, stockcode)
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(4, 100000)
        self.stock_chart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
        self.stock_chart.SetInputValue(6, ord("m"))
        self.stock_chart.SetInputValue(9, ord('1'))
        self.stock_chart.BlockRequest()

        candle = {}
        length = self.stock_chart.GetHeaderValue(3)
        while self.stock_chart.Continue:
            for i in range(length):
                if self.stock_chart.GetDataValue(0, i) not in pivots:
                    continue
                minute = self.stock_chart.GetDataValue(1, i)
                if 901 <= minute and minute < 1100:
                    candle.setdefault("dates", []).append(self.stock_chart.GetDataValue(0, i))
                    candle.setdefault("minutes", []).append(self.stock_chart.GetDataValue(1, i))
                    candle.setdefault("opens", []).append(self.stock_chart.GetDataValue(2, i))
                    candle.setdefault("highs", []).append(self.stock_chart.GetDataValue(3, i))
                    candle.setdefault("lows", []).append(self.stock_chart.GetDataValue(4, i))
                    candle.setdefault("closes", []).append(self.stock_chart.GetDataValue(5, i))
                    candle.setdefault("volumes", []).append(self.stock_chart.GetDataValue(6, i))
            if pivots[0] > self.stock_chart.GetDataValue(0, 0):
                break
            if self.opt.startdate > self.stock_chart.GetDataValue(0, 0):
                break
            if self.status.getLimitRemainCount(1) < 2:
                time.sleep(15.0)
            self.stock_chart.BlockRequest()
            self.log_request()
            length = self.stock_chart.GetHeaderValue(3)
        return candle

    def get_tick_data(self, stockcode, pivots):
        """deprecated"""
        if not isinstance(pivots, list) or not pivots:
            return [], [], [], []
        self.stock_chart.SetInputValue(0, stockcode)
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(4, 100000)
        self.stock_chart.SetInputValue(5, [0, 1, 2, 8])
        self.stock_chart.SetInputValue(6, ord("T"))
        self.stock_chart.SetInputValue(9, ord('1'))
        self.stock_chart.BlockRequest()
        length = self.stock_chart.GetHeaderValue(3)
        dates, minutes, prices, volumes = [], [], [], []
        while self.stock_chart.Continue:
            for i in range(length):
                if self.stock_chart.GetDataValue(0, i) not in pivots:
                    continue
                minute = self.stock_chart.GetDataValue(1, i)
                if 901 <= minute and minute < 1000:
                    dates.append(self.stock_chart.GetDataValue(0, i))
                    minutes.append(self.stock_chart.GetDataValue(1, i))
                    prices.append(self.stock_chart.GetDataValue(2, i))
                    volumes.append(self.stock_chart.GetDataValue(3, i))
            if pivots[0] > self.stock_chart.GetDataValue(0, i):
                break
            if self.status.getLimitRemainCount(1) < 2:
                time.sleep(15.0)
            self.stock_chart.BlockRequest()
            self.log_request()
            length = self.stock_chart.GetHeaderValue(3)
        return dates, minutes, prices, volumes

    def get_volume(self, stockcode):
        self.stock_chart.SetInputValue(0, stockcode)
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(4, 100000)
        self.stock_chart.SetInputValue(5, [0, 1, 8])
        self.stock_chart.SetInputValue(6, ord("m"))
        self.stock_chart.SetInputValue(9, ord('1'))
        self.stock_chart.BlockRequest()
        length = self.stock_chart.GetHeaderValue(3)
        dates, volumes = [], []
        while self.stock_chart.Continue:
            for i in range(length):
                if self.opt.startdate > self.stock_chart.GetDataValue(0, i):
                    break
                if self.stock_chart.GetDataValue(1, i) == 901:
                    dates.append(self.stock_chart.GetDataValue(0, i))
                    volumes.append(self.stock_chart.GetDataValue(2, i))
            if self.opt.startdate > self.stock_chart.GetDataValue(0, 0):
                break
            if self.status.getLimitRemainCount(1) < 2:
                time.sleep(15.0)
            self.stock_chart.BlockRequest()
            self.log_request()
            length = self.stock_chart.GetHeaderValue(3)
        return dates, volumes

    def get_high_volume(self):
        stockcodes = self.get_stockcode()
        stockdata, dateset = dict(), set()
        for stockcode in tqdm.tqdm(stockcodes):
            dates, volumes = self.get_volume(stockcode)
            stockdata[stockcode] = {date: volume for date, volume in zip(dates, volumes)}
            dateset |= set(dates)
        high_volume = {}
        for date in dateset:
            price = [(stockcode, series[date]) for stockcode, series in stockdata.items() if date in series]
            price = sorted(price, key=lambda x:x[1], reverse=True)
            high_volume[date] = [stockcode for stockcode, p in price[:self.opt.top_volume]]
        return high_volume

    def save(self, stock_map):
        stockcodes = []
        self.logger.info("Extract minute data : {} stocks".format(len(stock_map.keys())))
        with h5py.File(pjoin(self.opt.export_to, "ninetoten.h5"), "w") as fout:
            for key, value in tqdm.tqdm(stock_map.items()):
                candle = self.get_minute_data(key, value)
                if not candle:
                    continue
                stockcodes.append(key)
                stockgroup = fout.create_group(key)
                stockgroup.create_dataset("dates", data=candle["dates"])
                stockgroup.create_dataset("minutes", data=candle["minutes"])
                stockgroup.create_dataset("opens", data=candle["opens"])
                stockgroup.create_dataset("highs", data=candle["highs"])
                stockgroup.create_dataset("lows", data=candle["lows"])
                stockgroup.create_dataset("closes", data=candle["closes"])
                stockgroup.create_dataset("volumes", data=candle["volumes"])
        with open(pjoin(self.opt.export_to, "ninetoten.keys"), "w") as fout:
            fout.write("\n".join(stockcodes))

    def run(self):
        self.get_dispatch()
        volume = self.get_high_volume()
        stock_map = {}
        for pivot, stockcodes in volume.items():
            for stockcode in stockcodes:
                stock_map.setdefault(stockcode, []).append(pivot)
        stock_map = {k: sorted(v) for k, v in stock_map.items()}
        self.save(stock_map)


if __name__ == "__main__":
    fire.Fire(NinetoTen)
