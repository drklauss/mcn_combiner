import openpyxl
import datetime
import logging
import os


class Combiner:
    STATISTICS_PRICE_SEARCH_PHRASE = 'statistics'
    STATISTICS_COLUMNS_COUNT = 7
    CALLS_PRICE_SEARCH_PHRASE = 'calls'
    CALLS_COLUMNS_COUNT = 8
    _statistics_price_name = _calls_price_name = None
    _statistics = []  # Собеседник, Время разговора, Время звонка, Стоимость c НДС
    _calls = []  # Абонент, Собеседник, Время звонка
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
        """Вычитывает файл статистики"""
        statistics_sheet = openpyxl.load_workbook(filename=self._statistics_price_name).active
        rows_total = 0
        unread_rows = []
        logging.info('Start reading statistics file')
        for i in statistics_sheet.rows:
            rows_total += 1
            row = tuple(filter(lambda x: x is not None, map(lambda x: x.value, i)))
            if len(row) == self.STATISTICS_COLUMNS_COUNT:
                try:
                    call_time = datetime.datetime.strptime(row[0], '%d-%m-%Y %H:%M:%S').timestamp()
                    call_duration = datetime.datetime.strptime(row[4], '%H:%M:%S').time()
                    amount_with_nds = float(str.replace(row[6], ',', '.', -1)) * 1.18
                    self._statistics.append(tuple((int(row[3]), call_duration, int(call_time), amount_with_nds)))
                except ValueError:
                    unread_rows.append(row)
                    continue
        unread_rows_count = rows_total - len(unread_rows)
        if unread_rows_count != rows_total:
            logging.warning('Rows total: {}, rows read: {}'.format(rows_total, unread_rows_count))
            logging.warning('Unread rows: {}'.format(unread_rows))

    def read_calls(self):
        """Вычитывает файл звонков"""
        calls_sheet = openpyxl.load_workbook(filename=self._calls_price_name).active
        rows_total = 0
        unread_rows = []
        logging.info('Start reading calls file')
        for i in calls_sheet.rows:
            rows_total += 1
            row = tuple(filter(lambda x: x is not None, map(lambda x: x.value, i)))
            if len(tuple(row)) == self.CALLS_COLUMNS_COUNT:
                try:
                    self._calls.append(tuple((int(row[1]), int(row[3]), int(row[6].timestamp()))))
                except ValueError:
                    unread_rows.append(row)
                    continue
        unread_rows_count = rows_total - len(unread_rows)
        if unread_rows_count != rows_total:
            logging.warning('Rows total: {}, rows read: {}'.format(rows_total, unread_rows_count))
            logging.warning('Unread rows: {}'.format(unread_rows))

    def run(self):
        """Запуск """
        self.init_logger()
        self.init_files()
        self.read_statistic()
        self.read_calls()
        print(self._statistics)
        print(self._calls)
