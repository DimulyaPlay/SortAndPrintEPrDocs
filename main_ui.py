import datetime
import os.path
import sys
from tkinter import *
import win32print
from tkinterdnd2 import *

import stats_module
from sorter_class import *
import configparser
from scrollable_frame import VerticalScrolledFrame

ver = '3.1'
curdate = '16/06/2022'

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

# пути к скрипту и конфигу
config_name = 'config.ini'
stats_name = 'statistics.xlsx'
PDF_PRINT_NAME = 'PDFtoPrinter.exe'
printer_list = [i[2] for i in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
statfile_path = os.path.join(application_path, stats_name)
config_path = os.path.join(application_path, config_name)
PDF_PRINT_FILE = os.path.join(application_path, PDF_PRINT_NAME)

default_config = {'delete_zip': 'no',
                  'paper_eco_mode': 'yes',
                  'print_directly': 'no',
                  'save_stat': 'no',
                  'default_printer': win32print.GetDefaultPrinter(),
                  'PDF_PRINT_PATH': PDF_PRINT_FILE,
                  }


def readcreateconfig(default_config, config_path):
    # создание конфига или открыие имеющегося
    config = configparser.ConfigParser()
    if not os.path.exists(config_path):
        config['DEFAULT'] = default_config
        with open(config_path, 'w') as configfile:
            config.write(configfile)
        print('default config created')
    else:
        config.read(config_path)
        print('config read')
    return config


def write_config_to_file(class_obj, config_obj):
    # записать переменные в конфиг
    config_obj['DEFAULT']['delete_zip'] = class_obj.deletezip
    config_obj['DEFAULT']['paper_eco_mode'] = class_obj.paperecomode
    config_obj['DEFAULT']['print_directly'] = class_obj.print_directly
    config_obj['DEFAULT']['save_stat'] = class_obj.save_stat
    config_obj['DEFAULT']['default_printer'] = class_obj.default_printer
    config_obj['DEFAULT']['PDF_PRINT_PATH'] = os.path.join(os.path.dirname(config_path),
                                                           'PDFtoPrinter.exe')  # Установить место хранения программы
    print('saved')
    with open(config_path, 'w') as configfile:
        config_obj.write(configfile)


sorter = main_sorter(config=readcreateconfig(default_config, config_path), config_path=config_path)
stat_counter = stats_module.stat_reader(statfile_path)


def main_drop(event):
    if ' ' in event.data:
        path = event.data[1:-1]
    else:
        path = event.data
    if path[-4:] != '.zip':
        not_zip()
        return
    sorter.agregate_file(path)
    if sorter.print_directly == "yes":
        print_dialog()


def move_app(e):
    root.geometry(f'+{e.x_root}+{e.y_root}')


def quitter(e):
    root.quit()
    root.destroy()


def apply(e=sorter):
    # Set main class vars from checkbuttons
    sorter.deletezip = opt1.get()
    sorter.paperecomode = opt2.get()
    sorter.print_directly = opt3.get()
    sorter.default_printer = opt4.get()
    sorter.save_stat = opt5.get()
    write_config_to_file(sorter, sorter.config_obj)


def show_settings(e):
    settings = Toplevel(root)
    settings.title("Параметры")
    Checkbutton(settings, text="Удалить Zip",
                variable=opt1,
                onvalue='yes', offvalue='no', command=apply).pack(anchor=W)

    Checkbutton(settings, text="Эко режим",
                variable=opt2,
                onvalue='yes', offvalue='no', command=apply).pack(anchor=W)

    Checkbutton(settings, text="Печать на принтер",
                variable=opt3,
                onvalue='yes', offvalue='no', command=apply).pack(anchor=W)
    Checkbutton(settings, text="Сохранять статистику",
                variable=opt5,
                onvalue='yes', offvalue='no', command=apply).pack(anchor=W)
    OptionMenu(settings, opt4, *printer_list, command=apply).pack(anchor=W)
    showcredits = Label(settings, text="  Автор  ", borderwidth=2, relief="groove")
    showcredits.pack(anchor=S, padx=2, pady=2)
    showcredits.bind("<Button-1>", show_credits)
    opengh = Label(settings, text=" GitHub ", borderwidth=2, relief="groove")
    opengh.pack(anchor=S, padx=2, pady=2)
    opengh.bind("<Button-1>", lambda e: os.startfile('https://github.com/DimulyaPlay/SortAndPrintEPrDocs'))


def print_dialog():
    dialog = Toplevel(root)
    dialog.title(f'Файлов на печать {len(sorter.files_for_print)}')
    dialog.attributes('-topmost', True)
    dialog.resizable(False, False)

    def apply_print(e):
        for i, j in printcbVariables.items():
            if j.get():
                print_file(multiplePagesPerSheet(i, rbVariables[i].get()), PDF_PRINT_FILE, sorter.default_printer)
        show_printed()

    def update_num_pages():
        full_len = sum([sorter.num_pages[winSt[i]][0] for i in range(len(winSt)) if printcbVariables[winSt[i]].get()])
        eco_len = sum(
            [int(sorter.num_pages[winSt[i]][0] / 2 / (rbVariables[winSt[i]].get()) + 0.9) for i in range(len(winSt)) if printcbVariables[winSt[i]].get()])
        string = f"Всего для печати страниц: {full_len}, листов: {eco_len}"
        len_pages.set(string)

    container = VerticalScrolledFrame(dialog, height=550 if len(sorter.files_for_print) > 20 else (len(sorter.files_for_print)+1)*25, width=731)
    container.pack()
    winSt = sorter.files_for_print
    winSt_names = [os.path.basename(i) if len(os.path.basename(i)) < 58 else os.path.basename(i)[:55]+'...' for i in winSt]
    printcbVariables = {}
    printcb = {}
    filenames = {}
    numpages = {}
    rbVariables = {}
    rbuttons1 = {}
    rbuttons2 = {}
    rbuttons4 = {}
    previewbtns = {}
    Label(container, text='Название документа').grid(column=1, row=0)
    Label(container, text='Страниц').grid(column=2, row=0)
    Label(container, text='1').grid(column=3, row=0)
    Label(container, text='2').grid(column=4, row=0)
    Label(container, text='4').grid(column=5, row=0)
    for i in range(len(winSt)):
        printcbVariables[winSt[i]] = BooleanVar()
        printcbVariables[winSt[i]].set(1)
        printcb[i] = Checkbutton(container, variable=printcbVariables[winSt[i]], command=update_num_pages)
        printcb[i].grid(column=0, row=i + 1, sticky=W)
        filenames[i] = Label(container, text=winSt_names[i], font='TkFixedFont')
        filenames[i].grid(column=1, row=i + 1, sticky=W)
        filenames[i].bind('<Double-Button-1>', lambda event, a=winSt[i]: os.startfile(a))
        numpages[i] = Label(container, text=str(sorter.num_pages[winSt[i]][0]), padx=2)
        numpages[i].grid(column=2, row=i + 1)
        rbVariables[winSt[i]] = IntVar()
        rbVariables[winSt[i]].set(1)
        rbuttons1[i] = Radiobutton(container, variable=rbVariables[winSt[i]], value=1, command=update_num_pages)
        rbuttons1[i].grid(column=3, row=i + 1, sticky=W)
        rbuttons2[i] = Radiobutton(container, variable=rbVariables[winSt[i]], value=2, command=update_num_pages)
        rbuttons2[i].grid(column=4, row=i + 1, sticky=W)
        rbuttons4[i] = Radiobutton(container, variable=rbVariables[winSt[i]], value=4, command=update_num_pages)
        rbuttons4[i].grid(column=5, row=i + 1, sticky=W)
        previewbtns[i] = Label(container, text='Предпросмотр', padx=2)
        previewbtns[i].grid(column=6, row=i + 1)
        previewbtns[i].bind('<Button-1>', lambda event, a=winSt[i]: os.startfile(multiplePagesPerSheet(a, rbVariables[a].get())))
    bottom_actions = Frame(dialog)
    bottom_actions.pack()
    len_pages = StringVar()
    update_num_pages()
    if sorter.save_stat == "yes":
        statsaver = BooleanVar()
        statsaver.set(1)
        save_to_stat_chkbtn = Checkbutton(bottom_actions, variable=statsaver, text='Добавить в статистику', command=lambda: print(statsaver.get()))
        save_to_stat_chkbtn.grid(column=0, row=0, sticky=S, padx=5, pady=2)
    open_folder_b = Label(bottom_actions, text=" Открыть папку ", borderwidth=2, relief="groove")
    open_folder_b.grid(column=1, row=0, sticky=S, padx=5, pady=2)
    open_folder_b.bind("<Button-1>", lambda event, a=os.path.dirname(winSt[0]): subprocess.Popen(f'explorer {a}'))
    print_b = Label(bottom_actions, text=" Печать ", borderwidth=2, relief="groove")
    print_b.grid(column=2, row=0, sticky=S, padx=5, pady=2)
    print_b.bind("<Button-1>", apply_print)
    sum_pages = Label(bottom_actions, textvariable=len_pages)
    sum_pages.grid(column=3, row=0, sticky=S, padx=5, pady=2)


def show_printed():
    messagebox.showinfo("Готово", "Документы отправлены в очередь принтера.")


def not_zip():
    messagebox.showwarning("Варнинг", "Загружен не Zip архив.")


def show_credits(e):
    messagebox.showinfo("Кредитс", message=f"Сортировка документов с сайта Электронное провосудие.\nАвтор: консультант Краснокамского гс "
                        f"Соснин Дмитрий.\nВерсия {ver} от {curdate}")


# main window
root = Tk()
root.geometry('35x45')
root.overrideredirect(True)
root.attributes('-topmost', True)

title_bar = Frame(root, bd=0)
title_bar.pack(expand=0, fill=X)
title_bar.bind("<B1-Motion>", move_app)
title_bar.bind("<Double-Button-1>", show_settings)

close_label = Label(title_bar, text=' X ')
close_label.pack(side=RIGHT)
close_label.bind("<Button-1>", quitter)

entry = Label(root)
entry.pack(fill=X)
entry.drop_target_register(DND_FILES)
entry.dnd_bind('<<Drop>>', main_drop)

opt1 = StringVar()
opt1.set(sorter.deletezip)
opt2 = StringVar()
opt2.set(sorter.paperecomode)
opt3 = StringVar()
opt3.set(sorter.print_directly)
opt4 = StringVar()
opt4.set(sorter.default_printer)
opt5 = StringVar()
opt5.set(sorter.save_stat)

root.mainloop()
