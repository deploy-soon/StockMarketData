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

sys.path.append("../tools")
from misc import get_logger


SATURDAY = 6
SUNDAY = 7

class DART:

    def __init__(self, days):
        self.logger = get_logger()
        self.root = "http://dart.fss.or.kr"
        self.from_date = datetime.now()
        self.to_date = datetime.now() - timedelta(days=days)

    def _check_page_valid(self, trs):
        try:
            if len(trs) < 2:
                self.logger.info("INVALID TABLE")
                return False
            if trs[1].find("td", class_="no_data"):
                self.logger.info("NO DATA PAGE")
                return False
            return True
        except:
            self.logger.info("INVALID PAGE")
            return False

    def _extract_tuple(self, tr):
        try:
            tds = tr.find_all("td")
            time = tds[0].text.split(":")
            hour = int(time[0])
            minute = int(time[1])
            company_a = tds[1].find('a')
            company_href = company_a.get("href")
            company_id = company_href.split("=")[-1]
            company_name = company_a.text.strip()
            report = tds[2].find('a')
            report_href = report.get('href')
            report_title = report.text.strip()
            report_date = tds[4].text.split(".")
            year = int(report_date[0])
            month = int(report_date[1])
            day = int(report_date[2])
            return {
                "title": report_title.replace("\t", "").replace("\r\n", ""),
                "href": report_href,
                "company_name": company_name,
                "company_id": company_id,
                "datetime": datetime(year=year, month=month, day=day,
                                     hour=hour, minute=minute)
            }

        except:
            self.logger.info("INVALID TUPLE")
            return {}

    def get_main_page(self):
        report_list = []

        url = "{}/dsac001/mainAll.do".format(self.root)

        pivot_date = self.from_date
        while pivot_date >= self.to_date:
            if pivot_date.weekday() == SATURDAY or pivot_date.weekday() == SUNDAY:
                pivot_date = pivot_date - timedelta(days=1)
                continue

            self.logger.info("START GET {}".format(pivot_date.strftime("%Y.%m.%d")))
            params = {
                "selectDate": pivot_date.strftime("%Y%m%d"),
                "maxResults": 500,
                "currentPage": 1
            }
            while True:
                r = requests.get(url, params=params)
                self.logger.debug(r.url)
                if r.status_code != 200:
                    break
                soup = BeautifulSoup(r.content, "html.parser")
                trs = soup.find_all('tr')
                if not self._check_page_valid(trs):
                    break
                for tr in trs[1:]:
                    report_list.append(self._extract_tuple(tr))
                params["currentPage"] = params["currentPage"] + 1
                time.sleep(3)
            pivot_date = pivot_date - timedelta(days=1)

        self.logger.info("report num: {}".format(len(report_list)))
        return report_list

    def save(self, results):
        with open(pjoin('res', 'reports.csv'), 'w', encoding='utf-8', newline='') as fout:
            wr = csv.writer(fout, delimiter='\t')
            wr.writerow(["title", "href", "company", "company_id",
                         "year", "month", "day", "hout", "minute"])
            for result in results:
                wr.writerow([result["title"], result["href"],
                             result["company_name"], result["company_id"],
                             result["datetime"].year, result["datetime"].month,
                             result["datetime"].day, result["datetime"].hour,
                             result["datetime"].minute])

    def run(self):
        result = self.get_main_page()
        self.save(result)

if __name__ == "__main__":
    fire.Fire(DART)
