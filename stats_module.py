import os

import pandas as pd


class stat_reader:

	def __init__(self, statfile_path):
		"""
		Создание или чтение файла статистики
		:param statfile_path: путь к файлу статистики
		"""
		self.columns = ['Дата и время', 'Номер', 'Кол-во док-ов', 'Кол-во страниц в документах',
						'Кол-во листов в документах', 'Напечатано док-ов', 'Затрата без эко была бы',
						'Затрачено листов', 'Сэкономлено листов']
		self.statfile_path = statfile_path
		if os.path.exists(self.statfile_path):
			self.statfile = pd.read_excel(self.statfile_path)
			print('statfile read')
		else:
			self.statfile = pd.DataFrame(columns = self.columns)
			self.savestat()
			print('statfile created')

	def addstats(self, statrow):
		"""
		Добавление в файл статистики очередной строки
		:param statrow: лист со статистическими данными по количеству столбцов
		"""
		row_for_concat = {i:j for i, j in zip(self.columns, statrow)}
		df_for_concat = pd.DataFrame([row_for_concat])
		self.statfile = pd.concat([self.statfile, df_for_concat])

	def savestat(self):
		"""
		Сохранение файла статистики
		"""
		self.statfile.to_excel(self.statfile_path, index = False)
