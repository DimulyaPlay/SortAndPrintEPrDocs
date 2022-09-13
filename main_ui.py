import os.path
import subprocess
import sys
from functools import partial
from tkinter import *
from tkinter import messagebox

import win32print
from tkinterdnd2 import *

from config_loader import config_file
from msg_printer import MessageHandler
from epr_printer import print_dialog
from scrollable_frame import VerticalScrolledFrame
from sorter_class import *
from stats_module import stat_loader

# ver = '3.4.4'
# ver = '1.0.10_TRON'
ver = '1.0_5_JPrinterVer, 0.5_JavaUtils'
curdate = '2022/09/13'

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)
try:
    os.startfile(glob.glob(application_path + '//*.jar')[0])
except IndexError as e:
    raise IndexError('Не обнаружен файл для печати с расширением .jar') from e

os.startfile(glob.glob(application_path + '//*.jar')[0])
config_name = 'config.ini'  # название файла конфигурации
stats_name = 'statistics.xlsx'  # название файла статистики
PDF_PRINT_NAME = 'PDFtoPrinter.exe'  # название файла программы для печати
iconname = 'scales.ico'
printer_list = [i[2] for i in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]  # список принтеров в системе
statfile_path = os.path.join(application_path, stats_name)  # полный путь файла статистики
config_path = os.path.join(application_path, config_name)  # полный путь файла конфигурации
PDF_PRINT_FILE = os.path.join(application_path, PDF_PRINT_NAME)  # полный путь программы для печати
iconpath = os.path.join(application_path, iconname)
config_paths = [config_path, PDF_PRINT_FILE]

current_config = config_file(config_paths)

if current_config.save_stat == 'yes':
    stat_writer = stat_loader(statfile_path)
    sorterClass = main_sorter(current_config, stat=stat_writer)
else:
    sorterClass = main_sorter(current_config)
try:
    msg_handler = MessageHandler()
    outlook_connected = True
except:
    messagebox.showwarning("Ошибка", 'Не удалось соединиться с Outlook. Работа только с ЭПр')
    outlook_connected = False
    pass


def main_drop(event):
    if '{' in event.data:
        path = event.data[1:-1]
    else:
        path = event.data
    if path[-4:] != '.zip' and outlook_connected:
        msgnames = parse_names(event.data)
        msg_handler.handle_messages(msgnames)
        msg_handler.print_dialog_msg(root, current_config, iconpath)
    else:
        sorterClass.agregate_file(path)
        if current_config.print_directly == "yes":
            print_dialog(root, current_config, sorterClass, stat_writer, iconpath)
    dropzone.configure(text='+', foreground='black')
    root.attributes('-alpha', (int(current_config.gui_opacity) / 100))


def move_app(e):
    root.geometry(f'+{e.x_root}+{e.y_root}')


def quitter(e):
    os.system('taskkill /f /im javaw.exe')
    root.quit()
    root.destroy()


def apply(e=current_config):
    # Set main class vars from checkbuttons
    current_config.deletezip = opt1DelZip.get()
    current_config.paperecomode = opt2EcoMode.get()
    current_config.print_directly = opt3Print.get()
    current_config.default_printer = opt4DefPrinter.get()
    current_config.save_stat = opt5SaveStat.get()
    current_config.gui_opacity = opt6Opacity.get()
    current_config.no_protocols = opt7noProtocols.get()
    if current_config.save_stat == 'yes':
        stat_writer = stat_loader(statfile_path)
    root.attributes('-alpha', (int(current_config.gui_opacity) / 100))
    current_config.write_config_to_file()


def show_settings(e):
    settings = Toplevel(root)
    settings.iconbitmap(iconpath)
    settings.title("Параметры")
    Checkbutton(settings, text="Удалить Zip", variable=opt1DelZip, onvalue='yes', offvalue='no',
                command=apply).pack(anchor=W)

    Checkbutton(settings, text="Эко режим", variable=opt2EcoMode, onvalue='yes', offvalue='no',
                command=apply).pack(anchor=W)

    Checkbutton(settings, text="Печать на принтер", variable=opt3Print, onvalue='yes', offvalue='no',
                command=apply).pack(anchor=W)
    Checkbutton(settings, text="Сохранять статистику", variable=opt5SaveStat, onvalue='yes', offvalue='no',
                command=apply).pack(anchor=W)
    Checkbutton(settings, text="Без протоколов", variable=opt7noProtocols, onvalue='yes', offvalue='no',
                command=apply).pack(anchor=W)
    Scale(settings, from_=10, to=100, orient=HORIZONTAL, variable=opt6Opacity, command=apply).pack(anchor=W,
                                                                                                   fill=X)
    Label(settings, text='Прозрачность интерфейса').pack(anchor=W, fill=X, pady=5)
    OptionMenu(settings, opt4DefPrinter, *printer_list, command=apply).pack(anchor=W)
    showcredits = Label(settings, text="  Автор  ", borderwidth=2, relief="groove")
    showcredits.pack(anchor=S, padx=2, pady=2, fill=X)
    showcredits.bind("<Button-1>", info_show_credits)
    opengh = Label(settings, text=" GitHub ", borderwidth=2, relief="groove")
    opengh.pack(anchor=S, padx=2, pady=2, fill=X)
    opengh.bind("<Button-1>", lambda e: os.startfile('https://github.com/DimulyaPlay/SortAndPrintEPrDocs'))
    opengstat = Label(settings, text="Просмотр статистики", borderwidth=2, relief="groove")
    opengstat.pack(anchor=S, padx=2, pady=2, fill=X)
    opengstat.bind("<Button-1>", lambda e: os.startfile(statfile_path))
    opengstat = Label(settings, text="Просмотр конфига", borderwidth=2, relief="groove")
    opengstat.pack(anchor=S, padx=2, pady=2, fill=X)
    opengstat.bind("<Button-1>", lambda e: os.startfile(config_path))


def info_show_printed():
    messagebox.showinfo("Готово", "Документы отправлены в очередь принтера.")


def info_show_credits(e):
    messagebox.showinfo("Кредитс",
                        message=f"Сортировка документов с сайта Электронное провосудие.\nАвтор: консультант Краснокамского гс "
                                f"Соснин Дмитрий.\nВерсия {ver} от {curdate}")


def color_config_enter(widget, color, event):
    widget.configure(foreground=color)
    root.attributes('-alpha', 1)


def color_config_leave(widget, color, event):
    widget.configure(foreground=color)
    root.attributes('-alpha', (int(current_config.gui_opacity) / 100))


# main window
root = Tk()
root.geometry('38x50')
root.overrideredirect(True)
root.attributes('-topmost', True)
root.attributes('-alpha', (int(current_config.gui_opacity) / 100))

title_bar = Frame(root, bd=0)
title_bar.pack(expand=0, fill=X)
title_bar.bind("<B1-Motion>", move_app)
title_bar.bind("<Double-Button-1>", show_settings)

close_button = Label(title_bar, text='X', font=('Arial', 7))
close_button.pack(side=RIGHT)
close_button.bind("<Button-1>", quitter)

settings_button = Label(title_bar, text='S', font=('Arial', 7))
settings_button.pack(side=RIGHT)
settings_button.bind("<Button-1>", show_settings)

dropzone = Label(root, text='+', relief="ridge", font=('Arial', 20))
dropzone.pack(fill=X)
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
opt7noProtocols = StringVar()
opt7noProtocols.set(current_config.no_protocols)
root.mainloop()
