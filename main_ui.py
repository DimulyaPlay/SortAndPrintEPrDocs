import os.path
import time
import glob
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

# ver = '3.4.4'
# ver = '1.0.10_TRON'
ver = '1.01_JPrinter'
curdate = '2022/08/30'

if getattr(sys, 'frozen', False):
	application_path = os.path.dirname(sys.executable)
elif __file__:
	application_path = os.path.dirname(__file__)

os.startfile(glob.glob(application_path + '//*.jar')[0])
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

if os.path.exists(license_path):
	with open(license_path, 'r') as lic:
		LICENSE_KEY = lic.readline()
else:
	LICENSE_KEY = 'demo:1651643691881:7bbe6e960300000000f44976dbcbd47a0cb10a4317da3b8120ca6a1ff8'
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
		msg_handler.print_dialog_msg(root, current_config)
	else:
		sorterClass.agregate_file(path)
		if current_config.print_directly == "yes":
			print_dialog()
	dropzone.configure(text = '+', foreground = 'black')
	root.attributes('-alpha', (int(current_config.gui_opacity) / 100))


def move_app(e):
	root.geometry(f'+{e.x_root}+{e.y_root}')


def quitter(e):
	os.system('taskkill /f /im javaw.exe')
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
		for fp, prntcbvar in printcbVariables.items():
			if prntcbvar.get():
				to_queue_time = print_file(fp, rbVariables[fp].get(), current_config.default_printer,
										   convertVars[fp].get(), os.path.basename(fp))
				stat_writer.statdict['Постановка в очередь заняла'] = to_queue_time
				prntcbvar.set(0)
				lb1[fp].config(background = 'green1')
				lb1[fp].update()
		if current_config.save_stat == 'yes' and statsaver.get():
			print('saving to stats')
			stat_writer.statdict['Напечатано док-ов'] = num_docs_for_print.get()
			stat_writer.statdict[
				'Затрата без эко была бы'] = full_dupl_len_for_print_var.get() + eco_protocols_var.get()
			stat_writer.statdict['Затрачено листов'] = eco_dupl_len_for_print_var.get()
			stat_writer.statdict[
				'Сэкономлено листов'] = full_dupl_len_for_print_var.get() - eco_dupl_len_for_print_var.get() + eco_protocols_var.get()
			stat_writer.add_and_save_stats()
		# info_show_printed()
		update_num_pages()
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
		full_len_pages = 0
		eco_dupl_len_for_print = 0
		full_dupl_len_for_print = 0
		eco_protocols = 0
		string_num_docs = 0
		for fp in filepathsForPrint:
			if printcbVariables[fp].get():
				full_len_pages += sorterClass.num_pages[fp][0]
				eco_dupl_len_for_print += int(sorterClass.num_pages[fp][0] / 2 / (rbVariables[fp].get()) + 0.9)
				full_dupl_len_for_print += int(sorterClass.num_pages[fp][0] / 2 + 0.9)
				eco_protocols += sorterClass.num_protocols_eco[fp]
				string_num_docs += 1
		eco_dupl_len_for_print_var.set(eco_dupl_len_for_print)
		full_dupl_len_for_print_var.set(full_dupl_len_for_print)
		eco_protocols_var.set(eco_protocols)
		num_docs_for_print.set(string_num_docs)
		dialog.title(f'Документов на печать {num_docs_for_print.get()}')
		len_pages.set(f"Всего для печати страниц: {full_len_pages}, листов: {eco_dupl_len_for_print}")

	def check_all_chbtns():
		if prntchballvar.get():
			for chbtn in printcbVariables.values():
				chbtn.set(1)
		else:
			for chbtn in printcbVariables.values():
				chbtn.set(0)
		update_num_pages()

	container = VerticalScrolledFrame(dialog, height = 550 if len(sorterClass.files_for_print) > 20 else (
																												 len(sorterClass.files_for_print) + 1) * 25,
									  width = 640)
	container.pack()
	filepathsForPrint = sorterClass.files_for_print
	printcbVariables = {}
	rbVariables = {}
	lb1 = {}
	convertVars = {}
	prntchballvar = BooleanVar()
	prntchballvar.set(1)
	prntchball = Checkbutton(container, variable = prntchballvar, command = check_all_chbtns)
	prntchball.grid(column = 0, row = 0)
	Label(container, text = 'Название документа').grid(column = 1, row = 0)
	Label(container, text = 'Страниц').grid(column = 2, row = 0)
	Label(container, text = '1').grid(column = 3, row = 0)
	Label(container, text = '2').grid(column = 4, row = 0)
	Label(container, text = '4').grid(column = 5, row = 0)
	current_row = 1
	for fp in filepathsForPrint:
		fn = os.path.basename(fp)
		fn = fn if len(fn) < 58 else fn[:55] + '...'
		printcbVariables[fp] = BooleanVar()
		printcbVariables[fp].set(1)
		prntchb = Checkbutton(container, variable = printcbVariables[fp], command = update_num_pages)
		prntchb.grid(column = 0, row = current_row, sticky = W)
		lb1[fp] = Label(container, text = fn, font = 'TkFixedFont')
		lb1[fp].grid(column = 1, row = current_row, sticky = W)
		lb1[fp].bind('<Double-Button-1>', lambda event, a = fp:os.startfile(a))
		lb2 = Label(container, text = str(sorterClass.num_pages[fp][0]), padx = 2)
		lb2.grid(column = 2, row = current_row)
		rbVariables[fp] = IntVar()
		rbVariables[fp].set(1)
		rb1 = Radiobutton(container, variable = rbVariables[fp], value = 1, command = update_num_pages)
		rb1.grid(column = 3, row = current_row, sticky = W)
		rb2 = Radiobutton(container, variable = rbVariables[fp], value = 2, command = update_num_pages)
		rb2.grid(column = 4, row = current_row, sticky = W)
		rb4 = Radiobutton(container, variable = rbVariables[fp], value = 4, command = update_num_pages)
		rb4.grid(column = 5, row = current_row, sticky = W)
		convertVars[fp] = BooleanVar()
		convertVars[fp].set(0)
		convertchb = Checkbutton(container, variable = convertVars[fp], command = update_num_pages)
		convertchb.grid(column = 6, row = current_row, sticky = W)
		current_row += 1
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
dropzone.dnd_bind('<<DropEnter>>', partial(color_config_enter, dropzone, "green1"))
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
