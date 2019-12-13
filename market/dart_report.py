import os
import csv
import sys
from os.path import join as pjoin
import fire
import time
import tqdm
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import requests
import re

sys.path.append("../tools")
from misc import get_logger


class Report:

    def __init__(self, res_path="./res"):
        self.logger = get_logger()
        self.root = "http://dart.fss.or.kr"
        self.res_path = res_path
        self.res_file = pjoin(res_path, "report.csv")

    def load_reports(self):
        with open(pjoin(self.res_path, "reports.csv"), "r") as fin:
            reader = csv.DictReader(fin, delimiter='\t')
            header = reader
            for row in reader:
                yield row

    def check_report_datetime(self, row):
        if row.get("hout") and int(row.get("hout")) >= 16:
            return False
        if row.get("hout") and int(row.get("hout")) <= 8:
            return False
        today = datetime.now()
        report_datetime = datetime(year=int(row["year"]),
                                   month=int(row["month"]),
                                   day=int(row["day"]))
        if today - report_datetime > timedelta(days=365 * 2):
            return False
        return True

    def check_report_valid(self, title):
        raise NotImplemented

    def get_document(self, params):
        url = self.root + "/report/viewer.do"
        params = {
            "rcpNo": params[0],
            "dcmNo": params[1],
            "eleId": 0 if params[2] == 'null' else params[2],
            "offset": 0 if params[3] == 'null' else params[3],
            "length": 0 if params[4] == 'null' else params[4],
            "dtd": params[5],
        }
        r = requests.get(url, params=params)
        self.logger.debug(r.url)
        if r.status_code != 200:
            self.logger.warning("GET error: {}".format(r.url))
            return None
        data = self.parse_document(r.content)

        return data

    def parse_document(self, content):
        raise NotImplemented

    def _clean_data(self, data):
        return data.strip().replace("\\", "").replace("'", "")

    def _get_report_document(self, content):
        p = re.compile(r"viewDoc[(]+[^(^)]+[)]+")
        for r in p.finditer(str(content)):
            methodline = r.group()
            arguments = methodline[8:-1]
            params = arguments.split(",")
            if len(params) < 6:
                continue
            params = [self._clean_data(value) for value in params]
            if params[0].isdigit() and params[1].isdigit():
                return params
        return []

    def get_report(self, href):
        url = self.root + href
        r = requests.get(url)
        self.logger.debug(r.url)
        if r.status_code != 200:
            self.logger.warning("GET error: {}".format(r.url))
            return None
        params = self._get_report_document(r.text)
        if not params:
            return None

        return self.get_document(params)

    def save(self, results):
        if not isinstance(results, list) or len(results) == 0:
            self.logger.info("no results")
            return
        print("SAVE TO {}".format(self.res_file))
        with open(self.res_file, 'w', newline='') as fout:
            fieldnames = results[0].keys()
            writer = csv.DictWriter(fout,
                                    fieldnames=fieldnames, delimiter='\t')
            writer.writeheader()
            for result in results:
                writer.writerow(result)

    def run(self):
        results = []
        print("START CRAWLING TO {}".format(self.res_file))
        for row in tqdm.tqdm(self.load_reports()):
            if not self.check_report_valid(row.get("title")):
                continue
            if not self.check_report_datetime(row):
                continue
            data = self.get_report(row.get("href"))
            row.update(data)
            results.append(row)
            time.sleep(2)
        self.save(results)


class Danil(Report):

    def __init__(self):
        Report.__init__(self)
        self.res_file = pjoin(self.res_path, "danil.csv")

    def check_report_valid(self, title):
        if not title:
            return False
        return "단일판매" in title

    def parse_document(self, content):
        soup = BeautifulSoup(content, "html.parser")
        trs = soup.find_all('tr')
        data = {
            "total_payment": 0,
            "recent_profit": 0,
            "profit_ratio": 0.0,
            "big_deal": "",
        }
        for tr in trs:
            tds = tr.find_all("td")

            key = None
            value = None
            for td in tds:
                row = td.text.strip() or ""
                if "계약금액" in row:
                    key = "total_payment"
                elif "최근" in row and "매출액" in row:
                    key = "recent_profit"
                elif "매출액" in row and "대비" in row:
                    key = "profit_ratio"
                elif "대규모법인" in row:
                    key = "big_deal"
                elif key == "big_deal":
                    value = row
                elif row.replace(",", "").isdigit():
                    value = row.replace(",", "")
                    if value == "0":
                        value = None
                elif row.replace(".", "").isdigit():
                    value = row
            if key is not None and value is not None:
                data[key] = value
        return data

class Usang(Report):

    def __init__(self):
        Report.__init__(self)
        self.res_file = pjoin(self.res_path, "usang.csv")

    def check_report_valid(self, title):
        if not title:
            return False
        return "유상증자결정" in title

    def parse_document(self, content):
        soup = BeautifulSoup(content, "html.parser")
        trs = soup.find_all('tr')
        data = {
            "facility_fund": 0,
            "operation_fund": 0,
            "acquisition_fund": 0,
            "guitar_fund": 0,
        }
        for tr in trs:
            tds = tr.find_all("td")
            key, value = None, None
            for td in tds:
                row = td.text.strip() or ""
                if "시설자금" in row:
                    key = "facility_fund"
                elif "운영자금" in row:
                    key = "operation_fund"
                elif "취득자금" in row:
                    key = "acquisition_fund"
                elif "기타자금" in row:
                    key = "guitar_fund"
                elif row.replace(",", "").isdigit():
                    value = row.replace(",", "")
                    if value == "0":
                        value = None
                    if len(tds) > 2 and "시설자금" not in row:
                        value = None
            if key is not None and value is not None:
                data[key] = value
        return data


class TreasuryStock(Report):

    def __init__(self):
        Report.__init__(self)
        self.res_file = pjoin(self.res_path, "treasury.csv")

    def check_report_valid(self, title):
        if not title:
            return False
        return "자기주식취득결정" in title

    def parse_document(self, content):
        soup = BeautifulSoup(content, "html.parser")
        trs = soup.find_all('tr')
        data = {
            "buy_stock": 0,
            "buy_amount": 0,
        }
        for tr in trs:
            tds = tr.find_all("td")
            key, value = None, None
            for td in tds:
                row = td.text.strip() or ""
                if "취득예정주식" in row:
                    key = "buy_stock"
                elif "취득예정금액" in row:
                    key = "buy_amount"
                elif row.replace(",", "").isdigit():
                    value = row.replace(",", "")
                    if value == "0":
                        value = None
            if key is not None and value is not None:
                data[key] = value
        return data


class CB(Report):
    r"""
    Convertible Bond Report

    params
    ------
    href: link of cb report

    returns
    -------
    cb_amount: 사채의 권면총액(원)
    facility_fund: 자금조달의 목적(시설자금)
    operation_fund: 자금조달의 목적(운영자금)
    acquisition_fund: 자금조달의 목적(타법인 증권 취득자금)
    guitar_fund: 자금조달의 목적(기타자금)
    coupon_rate: 표면이자율(%)
    maturity_rate: 만기이자율(%)
    maturity_date: 사채만기일
    amortization_method: 사채발행방법
    stock_ratio: 전환에 따라 발행할 주식 - 주식총수대비(%)
    """

    def __init__(self):
        Report.__init__(self)
        self.res_file = pjoin(self.res_path, "cb.csv")

    def check_report_valid(self, title):
        if not title:
            return False
        return "전환사채권발행결정" in title

    def _list_contain(self, str_list, query):
        for token in str_list:
            if query in token:
                return True
        return False

    def _get_numeric_value(self, str_list):
        for token in str_list:
            if token.replace(",", "").isdigit():
                return token.replace(",", "")
        return ""

    def parse_document(self, content):
        soup = BeautifulSoup(content, "html.parser")
        trs = soup.find_all('tr')
        data = {
            "cb_amount": 0,
            "facility_fund": 0,
            "operation_fund": 0,
            "acquisition_fund": 0,
            "guitar_fund": 0,
            "coupon_rate": 0.0,
            "maturity_rate": 0.0,
            "maturity_date": "",
            "amortization_method": "",
            "stock_ratio": 0.0,
        }
        for tr in trs:
            tds = tr.find_all("td")
            tds_text = [td.text.strip() for td in tds if td.text.strip()]
            if self._list_contain(tds_text, "사채의 권면총액"):
                data["cb_amount"] = self._get_numeric_value(tds_text)
            elif self._list_contain(tds_text, "시설자금"):
                data["facility_fund"] = self._get_numeric_value(tds_text)
            elif self._list_contain(tds_text, "운영자금"):
                data["operation_fund"] = self._get_numeric_value(tds_text)
            elif self._list_contain(tds_text, "취득자금"):
                data["acquisition_fund"] = self._get_numeric_value(tds_text)
            elif self._list_contain(tds_text, "기타자금"):
                data["guitar_fund"] = self._get_numeric_value(tds_text)
            elif self._list_contain(tds_text, "표면이자율"):
                if len(tds_text[-1]) > 20:
                    continue
                data["coupon_rate"] = tds_text[-1]
            elif self._list_contain(tds_text, "만기이자율"):
                if len(tds_text[-1]) > 20:
                    continue
                data["maturity_rate"] = tds_text[-1]
            elif self._list_contain(tds_text, "사채만기일"):
                if len(tds_text[-1]) > 20:
                    continue
                data["maturity_date"] = tds_text[-1]
            elif self._list_contain(tds_text, "사채발행방법"):
                if len(tds_text[-1]) > 20:
                    continue
                data["amorization_method"] = tds_text[-1]
            elif self._list_contain(tds_text, "주식총수대비"):
                if len(tds_text[-1]) > 20:
                    continue
                data["stock_ratio"] = tds_text[-1]
        return data


if __name__ == "__main__":
    fire.Fire({
        "Danil": Danil,
        "Usang": Usang,
        "Treasury": TreasuryStock,
        "CB": CB,
    })
