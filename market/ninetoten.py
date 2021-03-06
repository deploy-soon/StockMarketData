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
            section = self.stockcode.GetStockSectionKind(code)
            if section != 1:
                # not corporation such as ETF, futures, ...
                continue
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
        """
        get minute data of target stock and datetimes
        :param stockcode: kospi, kosdaq stockcode which startswith "A"
        :param pivots: list of target datetimes
        :return: dict of dates, minutes, opens, highs, lows, closes, volumes
        """
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
                if 901 <= minute and minute < 1130:
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

    def get_volume(self, stockcode, time_mode="D"):
        assert time_mode in ["D", "W", "M", "m", "T"], "invalid time mode"

        self.stock_chart.SetInputValue(0, stockcode)
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(4, 1000)
        self.stock_chart.SetInputValue(5, [0, 8])
        self.stock_chart.SetInputValue(6, ord(time_mode))
        self.stock_chart.SetInputValue(9, ord('1'))
        self.stock_chart.BlockRequest()
        length = self.stock_chart.GetHeaderValue(3)
        dates, volumes = [], []
        if length > 0:
            for i in range(length):
                if self.opt.startdate > self.stock_chart.GetDataValue(0, i):
                    break
                dates.append(self.stock_chart.GetDataValue(0, i))
                volumes.append(self.stock_chart.GetDataValue(1, i))

        while self.stock_chart.Continue:
            for i in range(length):
                if self.opt.startdate > self.stock_chart.GetDataValue(0, i):
                    break
                dates.append(self.stock_chart.GetDataValue(0, i))
                volumes.append(self.stock_chart.GetDataValue(1, i))
            if self.opt.startdate > self.stock_chart.GetDataValue(0, 0):
                break
            if self.status.getLimitRemainCount(1) < 2:
                time.sleep(15.0)
            self.stock_chart.BlockRequest()
            self.log_request()
            length = self.stock_chart.GetHeaderValue(3)

        assert len(dates) == len(volumes)
        return dates, volumes

    def get_high_volume(self):
        """
        get high volume stocks between opt.fromdate to opt.todate
        :return: {
            "20200101": [code1, code2, ...],
            ...
        }
        """
        stockcodes = self.get_stockcode()
        stockdata, dateset = dict(), set()
        for stockcode in tqdm.tqdm(stockcodes):
            dates, volumes = self.get_volume(stockcode)
            stockdata[stockcode] = {int(date): volume for date, volume in zip(dates, volumes)}
            dateset |= set(dates)
        self.logger.info("get dates from {} to {}".format(min(dateset), max(dateset)))

        high_volume = {}
        for date in dateset:
            if date < self.opt.fromdate or self.opt.todate < date:
                continue
            price = [(stockcode, series[date]) for stockcode, series in stockdata.items() if date in series]
            price = sorted(price, key=lambda x: x[1], reverse=True)
            high_volume[date] = [stockcode for stockcode, p in price[:self.opt.top_volume]]

        return high_volume

    def save(self, stock_map):
        stockcodes = []
        self.logger.info("Extract minute data : {} stocks".format(len(stock_map.keys())))
        with h5py.File(pjoin(self.opt.export_to, "ninetoten.h5"), "w") as fout:
            for stockcode, datelist in tqdm.tqdm(stock_map.items()):
                candle = self.get_minute_data(stockcode, datelist)
                if not candle:
                    continue
                stockcodes.append(stockcode)
                stockgroup = fout.create_group(stockcode)
                stockgroup.create_dataset("dates", data=candle["dates"])
                stockgroup.create_dataset("minutes", data=candle["minutes"])
                stockgroup.create_dataset("opens", data=candle["opens"])
                stockgroup.create_dataset("highs", data=candle["highs"])
                stockgroup.create_dataset("lows", data=candle["lows"])
                stockgroup.create_dataset("closes", data=candle["closes"])
                stockgroup.create_dataset("volumes", data=candle["volumes"])
        with open(pjoin(self.opt.export_to, "ninetoten.keys"), "w") as fout:
            fout.write("\n".join(stockcodes))

    def test_save(self):
        with h5py.File(pjoin(self.opt.export_to, "ninetoten.h5"), "r") as fin:
            for k in fin.keys():
                print(k)
                print(fin[k]["dates"])

    def test(self):
        self.get_dispatch()
        self.get_stockcode()
        # print(self.get_volume("A000660"))

    def test_high_volume(self):
        self.get_dispatch()
        self.get_high_volume()

    def test_volume(self):
        self.get_dispatch()
        dates, volumes = self.get_volume("A093230")
        print(dates[0], volumes[0])

    def run(self):
        self.get_dispatch()
        volume = self.get_high_volume()
        stock_map = {}
        for date, stockcodes in volume.items():
            for stockcode in stockcodes:
                stock_map.setdefault(stockcode, []).append(date)
        stock_map = {k: sorted(v) for k, v in stock_map.items()}
        self.save(stock_map)


if __name__ == "__main__":
    fire.Fire(NinetoTen)
