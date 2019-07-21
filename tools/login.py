import win32com.client
from misc import get_logger, Option

class Status:
    def __init__(self, conf_path, verbose=False):
        self.logger = get_logger()
        self.verbose = verbose
        self.status = CpCybos.get_instance()
        self.opt = Option(conf_path)

    def assert_disconnect(self):
        assert self.status.getIsConnect()

    def get_dispatch(self):
        raise NotImplementedError

class CpCybos:

    def __init__(self):
        self.logger = get_logger()

    def OnDisconnect(self):
        # Event Handler that is called when server is disconnected
        self.logger.info("Server disconnected")

    def getIsConnect(self):
        return self.IsConnect

    def getServerType(self):
        """
        :return:
        > 0 : disconnect
        > 1: cybos server
        > 2: HTS server
        """
        return self.ServerType

    def getLimitRequestRemainTime(self):
        assert self.IsConnect
        return self.LimitRequestRemainTime

    def getLimitRemainCount(self, limitType):
        """
        :param limitType:
        > LT_TRADE_REQUEST : 0
        > LT_NONTRADE_REQUEST : 1
        > LT_SUBSCRIBE : 2
        :return: remained request num
        """
        assert self.IsConnect
        return self.GetLimitRemainCount(limitType)

    def disconnect(self):
        self.PlusDisconnect()

    @classmethod
    def get_instance(cls):
        return win32com.client.DispatchWithEvents("CpUtil.CpCybos", cls)
