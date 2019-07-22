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


class TImeSeries(Status):

    def get_dispatch(self):
        self.assert_disconnect()
        self.stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")

    def _set(self):
        feature = sorted(list(map(int,list(self.opt.format.keys()))))
        self.idmap = {key: num for num, key in enumerate(feature)}
        self.stock_chart.SetInputValue(0, self.opt.stockcode)
        self.stock_chart.SetInputValue(1, ord('2'))
        self.stock_chart.SetInputValue(2, self.opt.startdate)
        self.stock_chart.SetInputValue(4, self.opt.datalen)
        self.stock_chart.SetInputValue(5, feature)
        self.stock_chart.SetInputValue(6, ord(self.opt.datatype))
        self.stock_chart.SetInputValue(9, ord('1'))

    def _get_tuple(self, i):
        res = {value : self.stock_chart.GetDataValue(self.idmap[int(key)], i)
                for key, value in self.opt.format.items()}
        return res

    def log_request(self):
        code = self.stock_chart.GetDibStatus()
        message = self.stock_chart.GetDibMsg1()
        if code != 0:
            self.logger.warning("code : {}, message : {}".format(code, message))
            raise
        if self.verbose:
            self.logger.info("code : {}, message : {}".format(code, message))

    def _get_data_len(self):
        return self.stock_chart.GetHeaderValue(3)

    def _block_request(self, offset=0, len=0, is_first=True):
        self.stock_chart.BlockRequest()
        self.log_request()
        _offset = 0 if is_first else offset + len
        _len = self._get_data_len()
        return _offset, _len

    def consume(self):

        if not os.path.exists(self.opt.export_to):
            os.makedirs(self.opt.export_to)
        self.logger.info("load request format")
        self._set()

        res = list()
        offset, _len = self._block_request()
        for i in tqdm.tqdm(range(self.opt.datalen)):
            if i == _len + offset:

                if not self.stock_chart.Continue:
                    break
                offset, _len = self._block_request(offset=offset,
                                                   len=_len,
                                                   is_first=False)
                if self.status.getLimitRemainCount(1) < 2:
                    time.sleep(15.0)

            res.append(self._get_tuple(i-offset))

        self.save(res)

    def save(self, res):
        res = {key: [item[key] for item in res]
               for key in self.opt.format.values()}
        with h5py.File(pjoin(self.opt.export_to, "{}.h5".format(self.opt.stockcode)), "w") as f:
            for key, value in res.items():
                f.create_dataset(key, data=value)

    def run(self):
        self.get_dispatch()
        self.consume()

if __name__ == "__main__":
    fire.Fire(TImeSeries)