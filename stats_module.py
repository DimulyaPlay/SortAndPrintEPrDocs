import pandas as pd
import os
from datetime import datetime


class stat_reader:

    def __init__(self, statfile_path):
        columns = ['Дата и время', 'Номер', 'Кол-во док-ов', 'Кол-во страниц', 'Напечатано док-ов', 'Затрата без эко (листов)',  'Затрачено листов', 'Сэкономлено листов', 'Всего сэкономлено листов']
        self.statfile_path = statfile_path
        if os.path.exists(self.statfile_path):
            self.statfile = pd.read_excel(self.statfile_path)
            print('statfile read')
        else:
            self.statfile = pd.DataFrame(columns=columns)
            self.savestat()
            print('statfile created')

    def savestat(self):
        self.statfile.to_excel(self.statfile_path)

    def addstats(self, statrow):
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        self.statfile.append(dt_string+statrow)
