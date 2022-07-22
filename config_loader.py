import os
from configparser import ConfigParser

import win32print


class config_file:
	def __init__(self, config_filepaths):
		self.config_path = config_filepaths[0]
		self.default_config = {'delete_zip':'no', 'paper_eco_mode':'yes', 'print_directly':'yes', 'save_stat':'yes',
							   'default_printer':win32print.GetDefaultPrinter(), 'PDF_PRINT_PATH':config_filepaths[1],
							   'opacity':60}
		self.readcreateconfig()
		self.read_vars_from_config()

	def readcreateconfig(self):
		"""
		Чтение конфига или создание конфига по-умолчанию, если не обнаружен
		"""
		self.current_config = ConfigParser()
		if not os.path.exists(self.config_path):
			self.current_config['DEFAULT'] = self.default_config
			with open(self.config_path, 'w') as configfile:
				self.current_config.write(configfile)  # print('default config created')
		else:
			self.current_config.read(self.config_path)  # print('config read')

	def write_config_to_file(self):
		"""
		Запись в файл параметров конфигурации
		"""
		self.current_config['DEFAULT']['delete_zip'] = self.deletezip  # Удалять ли архив
		self.current_config['DEFAULT']['paper_eco_mode'] = self.paperecomode  # Режим экономии бумаги
		self.current_config['DEFAULT']['print_directly'] = self.print_directly  # Прямая печать на принтер
		self.current_config['DEFAULT']['save_stat'] = self.save_stat  # Сохранение статистики в файл
		self.current_config['DEFAULT']['default_printer'] = self.default_printer  # Принтер по умолчанию для программы
		self.current_config['DEFAULT']["opacity"] = self.gui_opacity  # Прозрачность основного окна
		with open(self.config_path, 'w') as configfile:
			self.current_config.write(configfile)

	# print('saved')

	def read_vars_from_config(self):
		"""
		Создание атрибутов класса из конфига для удобства
		"""
		self.deletezip = self.current_config.get('DEFAULT', 'delete_zip')  # Удалять ли архив
		self.paperecomode = self.current_config.get('DEFAULT', 'paper_eco_mode')  # Режим экономии бумаги
		self.print_directly = self.current_config.get('DEFAULT', 'print_directly')  # Прямая печать на принтер
		self.save_stat = self.current_config.get('DEFAULT', 'save_stat')  # Сохранение статистики в файл
		self.default_printer = self.current_config.get('DEFAULT',
													   'default_printer')  # Принтер по умолчанию для программы
		self.PDF_PRINT_FILE = self.current_config.get('DEFAULT', 'PDF_PRINT_PATH')  # Путь для программы PDFPrint
		self.gui_opacity = self.current_config.get('DEFAULT', "opacity")  # Прозрачность основного окна
