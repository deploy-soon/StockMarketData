import os
import sys
import csv
import fire
import time
import json
import win32com.client
from os.path import join as pjoin

sys.path.append("../tools")
from misc import get_logger
from login import Status


class DartMinute(Status):

    def __init__(self, stock_code, conf="./config/stock_minute.json", **kwargs):
        Status.__init__(self, conf=conf, **kwargs)
        self.logger = get_logger()
        self.model_path = "./minute_data"
        stock_code = str(stock_code)
        if not stock_code.startswith("A"):
            stock_code = "A" + stock_code
        self.stock_code = stock_code
        self.logger.info("target stock: {}".format(self.stock_code))
        if not os.path.isdir(self.opt.export_to):
            os.mkdir(self.opt.export_to)
        if not os.path.isdir(pjoin(self.opt.export_to, "minute")):
            os.mkdir(pjoin(self.opt.export_to, "minute"))

        self.stock_chart = None

    def get_dispatch(self):
        self.assert_disconnect()
        self.stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")

    def log_request(self):
        code = self.stock_chart.GetDibStatus()
        message = self.stock_chart.GetDibMsg1()
        if code != 0:
            self.logger.warning("code : {}, message : {}".format(code, message))
            raise
        if self.verbose:
            self.logger.info("code : {}, message : {}".format(code, message))

    def get_data(self):
        self.stock_chart.SetInputValue(0, self.stock_code)
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(4, 100000)
        self.stock_chart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
        self.stock_chart.SetInputValue(6, ord("m"))
        self.stock_chart.SetInputValue(9, ord('1'))
        self.stock_chart.SetInputValue(10, ord('3'))
        self.stock_chart.BlockRequest()
        length = self.stock_chart.GetHeaderValue(3)
        minute_data = list()
        date_data = set()
        while self.stock_chart.Continue:
            for i in range(length):
                if self.opt.todate > self.stock_chart.GetDataValue(0, i):
                    break
                _date = self.stock_chart.GetDataValue(0, i)
                minute_data.append([
                    self.stock_chart.GetDataValue(0, i),
                    self.stock_chart.GetDataValue(1, i),
                    self.stock_chart.GetDataValue(2, i),
                    self.stock_chart.GetDataValue(3, i),
                    self.stock_chart.GetDataValue(4, i),
                    self.stock_chart.GetDataValue(5, i),
                    self.stock_chart.GetDataValue(6, i),
                ])
                date_data.add(_date)
            if self.status.getLimitRemainCount(1) < 2:
                time.sleep(15.0)
            self.stock_chart.BlockRequest()
            self.log_request()
            length = self.stock_chart.GetHeaderValue(3)
        self.logger.info("get date from {} to {}".format(min(date_data), max(date_data)))
        minute_data.reverse()
        return minute_data

    def save(self, minute_data):
        with open(pjoin(self.opt.export_to, "minute", self.stock_code), "w", newline='', encoding='utf-8') as fout:
            writer = csv.writer(fout)
            writer.writerow(["date", "minute", "open", "high", "low", "close", "volume"])
            for data in minute_data:
                writer.writerow(data)

    def run(self):
        self.get_dispatch()
        minute_data = self.get_data()
        self.save(minute_data)


def test():
    import pprint
    with open(pjoin("res", "dart_report.json"), "r", newline='', encoding='utf-8') as fin:
        dart_reports = json.load(fin)
    for key, value in dart_reports.items():
        count = 0
        print(key, len(value))
        for v in value:
            if "prices" not in v:
                continue
            pprint.pprint(v)
            count = count + 1
            if count > 2:
                break

if __name__ == "__main__":
    fire.Fire(DartMinute)
    # test()
