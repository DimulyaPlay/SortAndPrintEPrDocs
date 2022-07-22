import os.path
import sys
from tkinter import *
from tkinter import messagebox

from tkinterdnd2 import *

from config_loader import config_file
from scrollable_frame import VerticalScrolledFrame
from sorter_class import *
from stats_module import stat_loader

ver = '3.4.3'
curdate = '2022/07/22'

if getattr(sys, 'frozen', False):
	application_path = os.path.dirname(sys.executable)
elif __file__:
	application_path = os.path.dirname(__file__)

config_name = 'config.ini'  # название файла конфигурации
stats_name = 'statistics.xlsx'  # название файла статистики
PDF_PRINT_NAME = 'PDFtoPrinter.exe'  # название файла программы для печати
printer_list = [i[2] for i in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]  # список принтеров в системе
statfile_path = os.path.join(application_path, stats_name)  # полный путь файла статистики
config_path = os.path.join(application_path, config_name)  # полный путь файла конфигурации
PDF_PRINT_FILE = os.path.join(application_path, PDF_PRINT_NAME)  # полный путь программы для печати
config_paths = [config_path, PDF_PRINT_FILE]
current_config = config_file(config_paths)
if current_config.save_stat == 'yes':
	stat_writer = stat_loader(statfile_path)
	sorterClass = main_sorter(current_config, stat = stat_writer)
else:
	sorterClass = main_sorter(current_config)


def main_drop(event):
	if '{' in event.data:
		path = event.data[1:-1]
	else:
		path = event.data
	if path[-4:] != '.zip':
		warning_not_zip()
		return
	sorterClass.agregate_file(path)
	if current_config.print_directly == "yes":
		print_dialog()


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
	current_config.gui_opacity = opt5Opacity.get()
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
	Scale(settings, from_ = 10, to = 100, orient = HORIZONTAL, variable = opt5Opacity, command = apply).pack(anchor = W,
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
		for i, j in printcbVariables.items():
			if j.get():
				print_file(multiplePagesPerSheet(i, rbVariables[i].get()), PDF_PRINT_FILE,
						   current_config.default_printer)
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
		print(eco_protocols)
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
									  width = 731)
	container.pack()
	filepathsForPrint = sorterClass.files_for_print
	filenamesForPrint = [os.path.basename(i) if len(os.path.basename(i)) < 58 else os.path.basename(i)[:55] + '...' for
						 i in filepathsForPrint]
	printcbVariables = {}
	printcb = {}
	filenames = {}
	numpages = {}
	rbVariables = {}
	rbuttons1perPage = {}
	rbuttons2perPage = {}
	rbuttons4perPage = {}
	previewbtns = {}
	Label(container, text = 'Название документа').grid(column = 1, row = 0)
	Label(container, text = 'Страниц').grid(column = 2, row = 0)
	Label(container, text = '1').grid(column = 3, row = 0)
	Label(container, text = '2').grid(column = 4, row = 0)
	Label(container, text = '4').grid(column = 5, row = 0)
	for i in range(len(filepathsForPrint)):
		printcbVariables[filepathsForPrint[i]] = BooleanVar()
		printcbVariables[filepathsForPrint[i]].set(1)
		printcb[i] = Checkbutton(container, variable = printcbVariables[filepathsForPrint[i]],
								 command = update_num_pages)
		printcb[i].grid(column = 0, row = i + 1, sticky = W)
		filenames[i] = Label(container, text = filenamesForPrint[i], font = 'TkFixedFont')
		filenames[i].grid(column = 1, row = i + 1, sticky = W)
		filenames[i].bind('<Double-Button-1>', lambda event, a = filepathsForPrint[i]:os.startfile(a))
		numpages[i] = Label(container, text = str(sorterClass.num_pages[filepathsForPrint[i]][0]), padx = 2)
		numpages[i].grid(column = 2, row = i + 1)
		rbVariables[filepathsForPrint[i]] = IntVar()
		rbVariables[filepathsForPrint[i]].set(1)
		rbuttons1perPage[i] = Radiobutton(container, variable = rbVariables[filepathsForPrint[i]], value = 1,
										  command = update_num_pages)
		rbuttons1perPage[i].grid(column = 3, row = i + 1, sticky = W)
		rbuttons2perPage[i] = Radiobutton(container, variable = rbVariables[filepathsForPrint[i]], value = 2,
										  command = update_num_pages)
		rbuttons2perPage[i].grid(column = 4, row = i + 1, sticky = W)
		rbuttons4perPage[i] = Radiobutton(container, variable = rbVariables[filepathsForPrint[i]], value = 4,
										  command = update_num_pages)
		rbuttons4perPage[i].grid(column = 5, row = i + 1, sticky = W)
		previewbtns[i] = Label(container, text = 'Предпросмотр', padx = 2)
		previewbtns[i].grid(column = 6, row = i + 1)
		previewbtns[i].bind('<Button-1>', lambda event, a = filepathsForPrint[i]:os.startfile(
			multiplePagesPerSheet(a, rbVariables[a].get())))
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
	open_folder_b = Label(bottom_actions, text = " Открыть папку ", borderwidth = 2, relief = "groove")
	open_folder_b.grid(column = 1, row = 0, sticky = S, padx = 5, pady = 2)
	open_folder_b.bind("<Button-1>",
					   lambda event, a = os.path.dirname(filepathsForPrint[0]):subprocess.Popen(f'explorer {a}'))
	print_button = Label(bottom_actions, text = " Печать ", borderwidth = 2, relief = "groove")
	print_button.grid(column = 2, row = 0, sticky = S, padx = 5, pady = 2)
	print_button.bind("<Button-1>", apply_print)
	sum_pages = Label(bottom_actions, textvariable = len_pages)
	sum_pages.grid(column = 3, row = 0, sticky = S, padx = 5, pady = 2)


def info_show_printed():
	messagebox.showinfo("Готово", "Документы отправлены в очередь принтера.")


def warning_not_zip():
	messagebox.showwarning("Варнинг", "Загружен не Zip архив.")


def info_show_credits(e):
	messagebox.showinfo("Кредитс",
						message = f"Сортировка документов с сайта Электронное провосудие.\nАвтор: консультант Краснокамского гс "
								  f"Соснин Дмитрий.\nВерсия {ver} от {curdate}")


def color_config(widget, color, event):
	widget.configure(foreground = color)


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
opt5Opacity = StringVar()
opt5Opacity.set(current_config.gui_opacity)
root.mainloop()
