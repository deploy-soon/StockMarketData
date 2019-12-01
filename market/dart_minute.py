import os
import sys
import csv
import fire
import tqdm
import h5py
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


class StockCodeCache(dict):

    def __init__(self, cache_path=pjoin("config", "stock_code.json")):
        try:
            arg = json.loads(open(cache_path).read())
        except:
            arg = {}
        for key, value in arg.items():
            self[key] = value
        self["__cache_path"] = cache_path

    def get_stock_code(self, dart_internal_code):
        stock_code = None
        try:
            url = "http://dart.fss.or.kr/api/company.json"
            params = {
                "auth": Config.DartAPIKey,
                "crp_cd": dart_internal_code
            }
            r = requests.get(url, params=params)
            stock_code = r.json().get("stock_cd")
        except:
            print("Get Stock Code Error", dart_internal_code)
        time.sleep(2)
        return stock_code

    def save(self):
        with open(self["__cache_path"], "w") as fout:
            json.dump({key: value for key, value in self.items()}, fout, indent=4)

    def __getitem__(self, key):
        if key not in self:
            value = self.get_stock_code(key)
            self.__setitem__(key, value)
            return value
        return super(StockCodeCache, self).__getitem__(key)

    def __setitem__(self, key, value):
        super(StockCodeCache, self).__setitem__(key, value)
        self.__dict__.update({key: value})

    def __delitem__(self, key):
        super(StockCodeCache, self).__delitem__(key)
        del self.__dict__[key]

class DartMinute(Status):

    def __init__(self, conf="./config/dart_minute.json", **kwargs):
        Status.__init__(self, conf=conf, **kwargs)
        self.model_path = "./model"
        self.stockCodeCache = StockCodeCache()
        self.logger = get_logger()

    def load(self):
        dart_reports = dict()
        file_list = os.listdir(self.model_path)
        for report in file_list:
            with open(pjoin(self.model_path, report), "r", newline='', encoding='utf-8') as fin:
                reader = csv.DictReader(fin, delimiter='\t')
                dart_reports[report] = list(reader)
        return dart_reports

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

    def add_report_info(self, report_book):
        for stock_code, reports in tqdm.tqdm(report_book.items()):
            self.get_data(stock_code, reports)

    def get_data(self, stock_code, reports):
        pivots = [int(r["year"] + r["month"] + r["day"]) for r in reports]
        min_date = min(pivots)

        self.stock_chart.SetInputValue(0, "A{}".format(stock_code))
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(4, 100000)
        self.stock_chart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])
        self.stock_chart.SetInputValue(6, ord("m"))
        self.stock_chart.SetInputValue(9, ord('1'))
        self.stock_chart.BlockRequest()
        length = self.stock_chart.GetHeaderValue(3)
        data = {}
        while self.stock_chart.Continue:
            for i in range(length):
                if min_date > self.stock_chart.GetDataValue(0, i):
                    break
                if self.stock_chart.GetDataValue(0, i) not in pivots:
                    continue
                _date = self.stock_chart.GetDataValue(0, i)
                data.setdefault(_date, []).append({
                    "date": self.stock_chart.GetDataValue(0, i),
                    "minute": self.stock_chart.GetDataValue(1, i),
                    "open": self.stock_chart.GetDataValue(2, i),
                    "high": self.stock_chart.GetDataValue(3, i),
                    "low": self.stock_chart.GetDataValue(4, i),
                    "close": self.stock_chart.GetDataValue(5, i),
                    "volume": self.stock_chart.GetDataValue(6, i),
                })
            if self.status.getLimitRemainCount(1) < 2:
                time.sleep(15.0)
            self.stock_chart.BlockRequest()
            self.log_request()
            length = self.stock_chart.GetHeaderValue(3)

        self.stock_chart.SetInputValue(0, "A{}".format(stock_code))
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(4, 100000)
        self.stock_chart.SetInputValue(5, [0, 12, 13, 17, 25, 26])
        self.stock_chart.SetInputValue(6, ord("D"))
        self.stock_chart.SetInputValue(9, ord('1'))
        self.stock_chart.BlockRequest()
        length = self.stock_chart.GetHeaderValue(3)
        day_data = {}
        while self.stock_chart.Continue:
            for i in range(length):
                if min_date > self.stock_chart.GetDataValue(0, i):
                    break
                if self.stock_chart.GetDataValue(0, i) not in pivots:
                    continue
                _date = self.stock_chart.GetDataValue(0, i)
                day_data[_date] = {
                    "stocks": self.stock_chart.GetDataValue(1, i),
                    "marketcap": self.stock_chart.GetDataValue(2, i),
                    "foreign": self.stock_chart.GetDataValue(3, i),
                    "turnover_ratio": self.stock_chart.GetDataValue(4, i),
                    "transation_ratio": self.stock_chart.GetDataValue(5, i),
                }
            if self.status.getLimitRemainCount(1) < 2:
                time.sleep(15.0)
            self.stock_chart.BlockRequest()
            self.log_request()
            length = self.stock_chart.GetHeaderValue(3)

        self._update_report(reports, data, day_data)

    def _update_report(self, reports, data, day_data):
        for report in reports:
            _data = data.get(int(report["year"] + report["month"] + report["day"]), [])
            _day_data = day_data.get(int(report["year"] + report["month"] + report["day"]), [])
            report_datetime = datetime.datetime(year=int(report["year"]), month=int(report["month"]),
                                                day=int(report["day"]), hour=int(report["hout"]),
                                                minute=int(report["minute"]))
            for d in _data:
                d_datetime = datetime.datetime(year=int(d["date"] / 10000), month=int(d["date"] / 100) % 100,
                                               day=d["date"] % 100, hour=int(d["minute"] / 100),
                                               minute=d["minute"] % 100)
                if (report_datetime <= d_datetime) and (d_datetime - report_datetime <= datetime.timedelta(minutes=7)):
                    new_data = d.copy()
                    new_data.update(_day_data)
                    report.setdefault("prices", []).append(new_data)

    def _report_intime(self, report):
        if int(report["hout"]) >= 16:
            return True
        if int(report["hout"]) >= 15 and int(report["minute"]) >= 25:
            return True
        elif int(report["hout"]) < 9:
            return True
        else:
            today = datetime.datetime.now()
            report_datetime = datetime.datetime(year=int(report["year"]), month=int(report["month"]),
                                                day=int(report["day"]))
            if today - report_datetime > datetime.timedelta(days=365 * 2):
                return True
            return False

    def save(self, dart_reports):
        with open(pjoin(self.opt.export_to, "dart_report.json"), "w", newline='', encoding='utf-8') as fout:
            json.dump(dart_reports, fout)

    def run(self):
        self.get_dispatch()
        dart_reports = self.load()

        report_book = {}
        for report_type, reports in dart_reports.items():
            count = 0
            for report in reports:
                if self._report_intime(report):
                    continue
                stock_code = self.stockCodeCache[report["company_id"]]
                self.stockCodeCache.save()
                if not stock_code:
                    continue
                report_book.setdefault(stock_code, []).append(report)
                count += 1
            self.logger.info("{} get {} reports".format(report_type, count))

        self.add_report_info(report_book)
        self.save(dart_reports)
        self.stockCodeCache.save()


def test():
    import pprint
    with open(pjoin("res", "dart_report.json"), "r", newline='', encoding='utf-8') as fin:
        dart_reports = json.load(fin)
    for key, value in dart_reports.items():
        print(key)
        count = 0
        for v in value:
            if "prices" not in v:
                continue
            pprint.pprint(v)
            count = count + 1
            if count > 5:
                break

if __name__ == "__main__":
    fire.Fire(DartMinute)
    # test()
