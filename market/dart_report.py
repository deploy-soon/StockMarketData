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

        with open(self.res_file, 'w', newline='') as fout:
            fieldnames = results[0].keys()
            writer = csv.DictWriter(fout,
                                    fieldnames=fieldnames, delimiter='\t')
            writer.writeheader()
            for result in results:
                writer.writerow(result)

    def run(self):
        results = []
        for row in self.load_reports():
            if not self.check_report_valid(row.get("title")):
                continue
            data = self.get_report(row.get("href"))
            print(data)
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


if __name__ == "__main__":
    fire.Fire({
        "Danil": Danil,
        "Usang": Usang,
    })
