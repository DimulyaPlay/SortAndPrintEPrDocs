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
from options import open_settings
from scrollable_frame import VerticalScrolledFrame
from sorter_class import *
from stats_module import stat_loader

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)
try:
    os.startfile(glob.glob(application_path + '//*.jar')[0])
except IndexError as e:
    raise IndexError('Не обнаружен файл для печати с расширением .jar') from e
documents_path = os.path.expanduser('~/Documents')
config_name = 'EPr_print_config.ini'  # название файла конфигурации
stats_name = 'EPr_print_statistics.xlsx'  # название файла статистики
iconname = 'scales.ico'
donaters_file_name = 'donaters.txt'
printer_list = [i[2] for i in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]  # список принтеров в системе
statfile_path = os.path.join(documents_path, stats_name)  # полный путь файла статистики
config_path = os.path.join(documents_path, config_name)  # полный путь файла конфигурации
iconpath = os.path.join(application_path, iconname)
encoded_file = os.path.join(application_path, donaters_file_name)

key = b' '
try:
    donaters_lst, exp_date = get_donaters_lst_from_encoded(encoded_file, key)
except:
    donaters_lst, exp_date = [], '2000/01/01'

try:
    current_config = config_file(config_path, donaters_lst, exp_date)
except:
    if os.path.exists(config_path):
        os.remove(config_path)
        current_config = config_file(config_path, donaters_lst, exp_date)
        messagebox.showinfo('Внимание',
                            'Версия конфигурационного файла устарела и была сброшена до настроек по умолчанию.')
    else:
        config_path = os.path.join(application_path, config_name)
        current_config = config_file(config_path, donaters_lst, exp_date)
        statfile_path = os.path.join(application_path, stats_name)
        messagebox.showinfo('Внимание',
                            'Невозможно создать конфиг в папке Документы, конфиг создан в папке с программой.')
    pass
if current_config.save_stat == 'yes':
    stat_writer = stat_loader(statfile_path)
    sorterClass = main_sorter(current_config, stat=stat_writer)
else:
    sorterClass = main_sorter(current_config)

try:
    msg_handler = MessageHandler()
    outlook_connected = True
except:
    outlook_connected = False
    pass


def main_drop(event):
    if '{' in event.data:
        path = event.data[1:-1]
    else:
        path = event.data
    print(path)
    if path.endswith('.msg'):
        if not outlook_connected:
            messagebox.showwarning("Ошибка", 'Не удалось соединиться с Outlook. Работа только с ЭПр')
        else:
            msgnames = parse_names(event.data)
            msgnames = [i if i.endswith('.msg') and i.startswith('C:') else '' for i in msgnames]
            try:
                msg_handler.handle_messages(msgnames)
                msg_handler.print_dialog_msg(root, current_config, iconpath)
            except Exception as e:
                print(e)
                messagebox.showwarning('Упс', f'Не удалось обработать сообщения\n{e}')
    elif path.endswith('.zip'):
        try:
            sorterClass.agregate_file(path)
            if current_config.print_directly == "yes":
                print_dialog(root, current_config, sorterClass, stat_writer, iconpath)
        except Exception as e:
            print(e)
            messagebox.showwarning('Упс', f'Не удалось обработать архив\n{e}')
    else:
        messagebox.showwarning('Упс', 'Данный тип файлов не поддерживается')
    dropzone.configure(text='+', foreground='black')
    root.attributes('-alpha', (int(current_config.gui_opacity) / 100))


def move_app(e):
    root.geometry(f'+{e.x_root}+{e.y_root}')


def quitter(e):
    os.system('taskkill /f /im javaw.exe')
    root.quit()
    root.destroy()


def show_settings(e):
    open_settings(root, current_config, statfile_path, iconpath, stat_loader, config_path)


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

root.mainloop()
