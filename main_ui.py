import os.path
import time
from functools import partial
from tkinter import *
from tkinter import messagebox

import win32print
from tkinterdnd2 import *

from config_loader import config_file
from msg_printer import Message_handler
from scrollable_frame import VerticalScrolledFrame
from sorter_class import *
from stats_module import stat_loader

ver = '3.4.4'
curdate = '2022/08/01'

if getattr(sys, 'frozen', False):
	application_path = os.path.dirname(sys.executable)
elif __file__:
	application_path = os.path.dirname(__file__)

license_name = 'PDFTron_license_key.txt'  # название файла с лицензией
config_name = 'config.ini'  # название файла конфигурации
stats_name = 'statistics.xlsx'  # название файла статистики
PDF_PRINT_NAME = 'PDFtoPrinter.exe'  # название файла программы для печати
printer_list = [i[2] for i in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]  # список принтеров в системе
statfile_path = os.path.join(application_path, stats_name)  # полный путь файла статистики
config_path = os.path.join(application_path, config_name)  # полный путь файла конфигурации
PDF_PRINT_FILE = os.path.join(application_path, PDF_PRINT_NAME)  # полный путь программы для печати
license_path = os.path.join(application_path, license_name)
config_paths = [config_path, PDF_PRINT_FILE]

current_config = config_file(config_paths)

with open(license_path, 'r') as lic:
	LICENSE_KEY = lic.readline()
initPDFTron(LICENSE_KEY)

if current_config.save_stat == 'yes':
	stat_writer = stat_loader(statfile_path)
	sorterClass = main_sorter(current_config, stat = stat_writer)
else:
	sorterClass = main_sorter(current_config)
msg_handler = Message_handler()


def main_drop(event):
	if '{' in event.data:
		path = event.data[1:-1]
	else:
		path = event.data
	if path[-4:] != '.zip':
		msgnames = parse_names(event.data)
		msg_handler.handle_messages(msgnames)
		msg_handler.print_dialog_msg(root)
	else:
		sorterClass.agregate_file(path)
		if current_config.print_directly == "yes":
			print_dialog()
	dropzone.configure(text = '+', foreground = 'black')
	root.attributes('-alpha', (int(current_config.gui_opacity) / 100))


def move_app(e):
	root.geometry(f'+{e.x_root}+{e.y_root}')


def quitter(e):
	root.quit()
	root.destroy()


def apply(e = current_config):
	# Set main class vars from checkbuttons
	current_config.deletezip = opt1DelZip.get()
	current_config.paperecomode = opt2EcoMode.get()
	current_config.print_directly = opt3Print.get()
	current_config.default_printer = opt4DefPrinter.get()
	current_config.save_stat = opt5SaveStat.get()
	current_config.gui_opacity = opt6Opacity.get()
	if current_config.save_stat == 'yes':
		stat_writer = stat_loader(statfile_path)
	root.attributes('-alpha', (int(current_config.gui_opacity) / 100))
	current_config.write_config_to_file()


def show_settings(e):
	settings = Toplevel(root)
	settings.title("Параметры")
	Checkbutton(settings, text = "Удалить Zip", variable = opt1DelZip, onvalue = 'yes', offvalue = 'no',
				command = apply).pack(anchor = W)

	Checkbutton(settings, text = "Эко режим", variable = opt2EcoMode, onvalue = 'yes', offvalue = 'no',
				command = apply).pack(anchor = W)

	Checkbutton(settings, text = "Печать на принтер", variable = opt3Print, onvalue = 'yes', offvalue = 'no',
				command = apply).pack(anchor = W)
	Checkbutton(settings, text = "Сохранять статистику", variable = opt5SaveStat, onvalue = 'yes', offvalue = 'no',
				command = apply).pack(anchor = W)
	Scale(settings, from_ = 10, to = 100, orient = HORIZONTAL, variable = opt6Opacity, command = apply).pack(anchor = W,
																											 fill = X)
	Label(settings, text = 'Прозрачность интерфейса').pack(anchor = W, fill = X, pady = 5)
	OptionMenu(settings, opt4DefPrinter, *printer_list, command = apply).pack(anchor = W)
	showcredits = Label(settings, text = "  Автор  ", borderwidth = 2, relief = "groove")
	showcredits.pack(anchor = S, padx = 2, pady = 2, fill = X)
	showcredits.bind("<Button-1>", info_show_credits)
	opengh = Label(settings, text = " GitHub ", borderwidth = 2, relief = "groove")
	opengh.pack(anchor = S, padx = 2, pady = 2, fill = X)
	opengh.bind("<Button-1>", lambda e:os.startfile('https://github.com/DimulyaPlay/SortAndPrintEPrDocs'))
	opengstat = Label(settings, text = "Просмотр статистики", borderwidth = 2, relief = "groove")
	opengstat.pack(anchor = S, padx = 2, pady = 2, fill = X)
	opengstat.bind("<Button-1>", lambda e:os.startfile(statfile_path))
	opengstat = Label(settings, text = "Просмотр конфига", borderwidth = 2, relief = "groove")
	opengstat.pack(anchor = S, padx = 2, pady = 2, fill = X)
	opengstat.bind("<Button-1>", lambda e:os.startfile(config_path))


def print_dialog():
	dialog = Toplevel(root)
	dialog.title(f'Файлов на печать {len(sorterClass.files_for_print)}')
	dialog.attributes('-topmost', True)
	dialog.resizable(False, False)

	def apply_print(e):
		print_button.unbind("<Button-1>")
		print_button.config(relief = SUNKEN)
		print_button.update()
		for i, j in printcbVariables.items():
			if j.get():
				print_file(i, rbVariables[i].get(), current_config.default_printer)
		if current_config.save_stat == 'yes' and statsaver.get():
			print('saving to stats')
			stat_writer.statdict['Напечатано док-ов'] = num_docs_for_print.get()
			stat_writer.statdict[
				'Затрата без эко была бы'] = full_dupl_len_for_print_var.get() + eco_protocols_var.get()
			stat_writer.statdict['Затрачено листов'] = eco_dupl_len_for_print_var.get()
			stat_writer.statdict[
				'Сэкономлено листов'] = full_dupl_len_for_print_var.get() - eco_dupl_len_for_print_var.get() + eco_protocols_var.get()
			stat_writer.add_and_save_stats()
		info_show_printed()
		print_button.config(relief = RAISED)
		print_button.bind("<Button-1>", apply_print)

	def open_folder(e):
		open_folder_b.unbind("<Button-1>")
		open_folder_b.config(relief = SUNKEN)
		open_folder_b.update()
		subprocess.Popen(f'explorer {os.path.dirname(filepathsForPrint[0])}')
		time.sleep(0.1)
		open_folder_b.config(relief = RAISED)
		open_folder_b.update()
		open_folder_b.bind("<Button-1>", open_folder)

	def update_num_pages():
		full_len_pages = sum([sorterClass.num_pages[filepathsForPrint[i]][0] for i in range(len(filepathsForPrint)) if
							  printcbVariables[filepathsForPrint[i]].get()])
		eco_dupl_len_for_print = sum(
			[int(sorterClass.num_pages[filepathsForPrint[i]][0] / 2 / (rbVariables[filepathsForPrint[i]].get()) + 0.9)
			 for i in range(len(filepathsForPrint)) if printcbVariables[filepathsForPrint[i]].get()])
		full_dupl_len_for_print = sum(
			[int(sorterClass.num_pages[filepathsForPrint[i]][0] / 2 + 0.9) for i in range(len(filepathsForPrint)) if
			 printcbVariables[filepathsForPrint[i]].get()])
		eco_protocols = sum(
			[sorterClass.num_protocols_eco[filepathsForPrint[i]] for i in range(len(filepathsForPrint)) if
			 printcbVariables[filepathsForPrint[i]].get()])
		string_num_docs = sum(
			[1 for i in range(len(filepathsForPrint)) if printcbVariables[filepathsForPrint[i]].get()])
		string_pages_papers = f"Всего для печати страниц: {full_len_pages}, листов: {eco_dupl_len_for_print}"
		num_docs_for_print.set(string_num_docs)
		eco_dupl_len_for_print_var.set(eco_dupl_len_for_print)
		full_dupl_len_for_print_var.set(full_dupl_len_for_print)
		eco_protocols_var.set(eco_protocols)
		dialog.title(f'Документов на печать {num_docs_for_print.get()}')
		len_pages.set(string_pages_papers)

	container = VerticalScrolledFrame(dialog, height = 550 if len(sorterClass.files_for_print) > 20 else (
																												 len(sorterClass.files_for_print) + 1) * 25,
									  width = 600)
	container.pack()
	filepathsForPrint = sorterClass.files_for_print
	filenamesForPrint = [os.path.basename(i) if len(os.path.basename(i)) < 58 else os.path.basename(i)[:55] + '...' for
						 i in filepathsForPrint]
	printcbVariables = {}
	rbVariables = {}
	Label(container, text = 'Название документа').grid(column = 1, row = 0)
	Label(container, text = 'Страниц').grid(column = 2, row = 0)
	Label(container, text = '1').grid(column = 3, row = 0)
	Label(container, text = '2').grid(column = 4, row = 0)
	Label(container, text = '4').grid(column = 5, row = 0)
	for i in range(len(filepathsForPrint)):
		printcbVariables[filepathsForPrint[i]] = BooleanVar()
		printcbVariables[filepathsForPrint[i]].set(1)
		prntchb = Checkbutton(container, variable = printcbVariables[filepathsForPrint[i]], command = update_num_pages)
		prntchb.grid(column = 0, row = i + 1, sticky = W)
		lb1 = Label(container, text = filenamesForPrint[i], font = 'TkFixedFont')
		lb1.grid(column = 1, row = i + 1, sticky = W)
		lb1.bind('<Double-Button-1>', lambda event, a = filepathsForPrint[i]:os.startfile(a))
		lb2 = Label(container, text = str(sorterClass.num_pages[filepathsForPrint[i]][0]), padx = 2)
		lb2.grid(column = 2, row = i + 1)
		rbVariables[filepathsForPrint[i]] = IntVar()
		rbVariables[filepathsForPrint[i]].set(1)
		rb1 = Radiobutton(container, variable = rbVariables[filepathsForPrint[i]], value = 1,
						  command = update_num_pages)
		rb1.grid(column = 3, row = i + 1, sticky = W)
		rb2 = Radiobutton(container, variable = rbVariables[filepathsForPrint[i]], value = 2,
						  command = update_num_pages)
		rb2.grid(column = 4, row = i + 1, sticky = W)
		rb4 = Radiobutton(container, variable = rbVariables[filepathsForPrint[i]], value = 4,
						  command = update_num_pages)
		rb4.grid(column = 5, row = i + 1, sticky = W)
	bottom_actions = Frame(dialog)
	bottom_actions.pack()
	len_pages = StringVar()
	num_docs_for_print = IntVar()
	full_len_pages_for_print_var = IntVar()
	eco_dupl_len_for_print_var = IntVar()
	full_dupl_len_for_print_var = IntVar()
	eco_protocols_var = IntVar()
	update_num_pages()
	if current_config.save_stat == "yes":
		statsaver = BooleanVar()
		statsaver.set(1)
		save_to_stat_chkbtn = Checkbutton(bottom_actions, variable = statsaver, text = 'Добавить в статистику',
										  command = lambda:print(statsaver.get()))
		save_to_stat_chkbtn.grid(column = 0, row = 0, sticky = S, padx = 5, pady = 2)
	open_folder_b = Label(bottom_actions, text = " Открыть папку ", borderwidth = 2, relief = RAISED)
	open_folder_b.grid(column = 1, row = 0, sticky = S, padx = 5, pady = 2)
	open_folder_b.bind("<Button-1>", open_folder)
	print_button = Label(bottom_actions, text = " Печать ", borderwidth = 2, relief = RAISED)
	print_button.grid(column = 2, row = 0, sticky = S, padx = 5, pady = 2)
	print_button.bind("<Button-1>", apply_print)
	sum_pages = Label(bottom_actions, textvariable = len_pages)
	sum_pages.grid(column = 3, row = 0, sticky = S, padx = 5, pady = 2)


def info_show_printed():
	messagebox.showinfo("Готово", "Документы отправлены в очередь принтера.")


def info_show_credits(e):
	messagebox.showinfo("Кредитс",
						message = f"Сортировка документов с сайта Электронное провосудие.\nАвтор: консультант Краснокамского гс "
								  f"Соснин Дмитрий.\nВерсия {ver} от {curdate}")


def color_config_enter(widget, color, event):
	widget.configure(foreground = color)
	root.attributes('-alpha', 1)


def color_config_leave(widget, color, event):
	widget.configure(foreground = color)
	root.attributes('-alpha', (int(current_config.gui_opacity) / 100))


# main window
root = Tk()
root.geometry('38x50')
root.overrideredirect(True)
root.attributes('-topmost', True)
root.attributes('-alpha', (int(current_config.gui_opacity) / 100))

title_bar = Frame(root, bd = 0)
title_bar.pack(expand = 0, fill = X)
title_bar.bind("<B1-Motion>", move_app)
title_bar.bind("<Double-Button-1>", show_settings)

close_button = Label(title_bar, text = 'X', font = ('Arial', 7))
close_button.pack(side = RIGHT)
close_button.bind("<Button-1>", quitter)

settings_button = Label(title_bar, text = 'S', font = ('Arial', 7))
settings_button.pack(side = RIGHT)
settings_button.bind("<Button-1>", show_settings)

dropzone = Label(root, text = '+', relief = "ridge", font = ('Arial', 20))
dropzone.pack(fill = X)
dropzone.drop_target_register(DND_FILES)
dropzone.dnd_bind('<<DropEnter>>', partial(color_config_enter, dropzone, "green"))
dropzone.dnd_bind('<<DropLeave>>', partial(color_config_leave, dropzone, "black"))
dropzone.dnd_bind('<<Drop>>', main_drop)

opt1DelZip = StringVar()
opt1DelZip.set(current_config.deletezip)
opt2EcoMode = StringVar()
opt2EcoMode.set(current_config.paperecomode)
opt3Print = StringVar()
opt3Print.set(current_config.print_directly)
opt4DefPrinter = StringVar()
opt4DefPrinter.set(current_config.default_printer)
opt5SaveStat = StringVar()
opt5SaveStat.set(current_config.save_stat)
opt6Opacity = StringVar()
opt6Opacity.set(current_config.gui_opacity)
root.mainloop()
