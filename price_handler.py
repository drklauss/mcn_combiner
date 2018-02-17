import openpyxl
import logging
import os


# shortNumberBook = openpyxl.load_workbook(filename='short_numbers.xlsx')
# sheet = shortNumberBook.active


# logging.info("ffffffffffff")
# logging.critical("erwerwerer")


# for i in sheet.rows:
#     logging.info(i[1].value)
#     print(i)


class Combiner:
    STATISTICS_PRICE_SEARCH_PHRASE = 'statistics'
    CALLS_PRICE_SEARCH_PHRASE = 'calls'
    _statistics_price_name = _calls_price_name = None
    _calls = []
    _statistics = []
    _result_sheet = []

    def init_files(self) -> None:
        """Поиск файлов прайсов"""
        files = os.listdir('.')
        xlsx_files = tuple(filter(lambda x: '.xlsx' in x, files))
        if 2 > len(xlsx_files) or len(xlsx_files) > 2:
            raise IndexError
        for file in xlsx_files:
            if self.STATISTICS_PRICE_SEARCH_PHRASE in file:
                self._statistics_price_name = file
            if self.CALLS_PRICE_SEARCH_PHRASE in file:
                self._calls_price_name = file
        if self._statistics_price_name is None or self._calls_price_name is None:
            raise FileNotFoundError
        logging.info('Found prices: {}, {}'.format(self._statistics_price_name, self._calls_price_name))

    @staticmethod
    def init_logger() -> None:
        """Инициализация логгера"""
        log_format = u'%(asctime)s %(levelname)s %(filename)s:%(lineno)d %(message)s'
        logging.basicConfig(filename='price_combiner.log', format=log_format, level=logging.DEBUG)

    def read_statistic(self):
        pass

    def read_calls(self):
        calls_sheet = openpyxl.load_workbook(filename=self._calls_price_name).active
        for i in calls_sheet.rows:
            print(*i)
            row = filter(lambda x: x is not None, map(lambda x: x.value, i))
            self._calls.append(tuple(row))

    def run(self):
        """Запуск """
        self.init_logger()
        self.init_files()
        self.read_calls()
        print(self._calls)
