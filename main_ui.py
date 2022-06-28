import os.path
import sys
from tkinter import *
from tkinter import messagebox
import win32print
from tkinterdnd2 import *

import stats_module
from sorter_class import *
import configparser
from scrollable_frame import VerticalScrolledFrame

ver = '3.3'
curdate = '28/06/2022'

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


sorterClass = main_sorter(config=readcreateconfig(default_config, config_path), config_path=config_path)
if sorterClass.save_stat == 'yes':
    stat_writer = stats_module.stat_reader(statfile_path)


def main_drop(event):
    if '{' in event.data:
        path = event.data[1:-1]
    else:
        path = event.data
    if path[-4:] != '.zip':
        warning_not_zip()
        return
    sorterClass.agregate_file(path)
    if sorterClass.print_directly == "yes":
        print_dialog()


def move_app(e):
    root.geometry(f'+{e.x_root}+{e.y_root}')


def quitter(e):
    root.quit()
    root.destroy()


def apply(e=sorterClass):
    # Set main class vars from checkbuttons
    sorterClass.deletezip = opt1DelZip.get()
    sorterClass.paperecomode = opt2EcoMode.get()
    sorterClass.print_directly = opt3Print.get()
    sorterClass.default_printer = opt4DefPrinter.get()
    sorterClass.save_stat = opt5SaveStat.get()
    write_config_to_file(sorterClass, sorterClass.config_obj)


def show_settings(e):
    settings = Toplevel(root)
    settings.title("Параметры")
    Checkbutton(settings, text="Удалить Zip",
                variable=opt1DelZip,
                onvalue='yes', offvalue='no', command=apply).pack(anchor=W)

    Checkbutton(settings, text="Эко режим",
                variable=opt2EcoMode,
                onvalue='yes', offvalue='no', command=apply).pack(anchor=W)

    Checkbutton(settings, text="Печать на принтер",
                variable=opt3Print,
                onvalue='yes', offvalue='no', command=apply).pack(anchor=W)
    Checkbutton(settings, text="Сохранять статистику",
                variable=opt5SaveStat,
                onvalue='yes', offvalue='no', command=apply).pack(anchor=W)
    OptionMenu(settings, opt4DefPrinter, *printer_list, command=apply).pack(anchor=W)
    showcredits = Label(settings, text="  Автор  ", borderwidth=2, relief="groove")
    showcredits.pack(anchor=S, padx=2, pady=2)
    showcredits.bind("<Button-1>", info_show_credits)
    opengh = Label(settings, text=" GitHub ", borderwidth=2, relief="groove")
    opengh.pack(anchor=S, padx=2, pady=2)
    opengh.bind("<Button-1>", lambda e: os.startfile('https://github.com/DimulyaPlay/SortAndPrintEPrDocs'))


def print_dialog():
    dialog = Toplevel(root)
    dialog.title(f'Файлов на печать {len(sorterClass.files_for_print)}')
    dialog.attributes('-topmost', True)
    dialog.resizable(False, False)

    def apply_print(e):
        for i, j in printcbVariables.items():
            if j.get():
                print_file(multiplePagesPerSheet(i, rbVariables[i].get()), PDF_PRINT_FILE, sorterClass.default_printer)
        if sorterClass.save_stat == 'yes' and statsaver.get():
            print('saving to stats')
            sorterClass.stats_list.append(num_docs_for_print.get())
            sorterClass.stats_list.append(full_dupl_len_for_print_var.get())
            sorterClass.stats_list.append(eco_dupl_len_for_print_var.get())
            sorterClass.stats_list.append(full_dupl_len_for_print_var.get()-eco_dupl_len_for_print_var.get())
            stat_writer.addstats(sorterClass.stats_list)
            stat_writer.savestat()
        info_show_printed()

    def update_num_pages():
        full_len_pages = sum([sorterClass.num_pages[filepathsForPrint[i]][0] for i in range(len(filepathsForPrint)) if printcbVariables[filepathsForPrint[i]].get()])
        eco_dupl_len_for_print = sum([int(sorterClass.num_pages[filepathsForPrint[i]][0] / 2 / (rbVariables[filepathsForPrint[i]].get()) + 0.9) for i in range(len(filepathsForPrint)) if printcbVariables[filepathsForPrint[i]].get()])
        full_dupl_len_for_print = sum([int(sorterClass.num_pages[filepathsForPrint[i]][0] / 2 + 0.9) for i in range(len(filepathsForPrint)) if printcbVariables[filepathsForPrint[i]].get()])
        string_num_docs = sum([1 for i in range(len(filepathsForPrint)) if printcbVariables[filepathsForPrint[i]].get()])
        string_pages_papers = f"Всего для печати страниц: {full_len_pages}, листов: {eco_dupl_len_for_print}"
        num_docs_for_print.set(string_num_docs)
        eco_dupl_len_for_print_var.set(eco_dupl_len_for_print)
        full_dupl_len_for_print_var.set(full_dupl_len_for_print)
        dialog.title(f'Документов на печать {num_docs_for_print.get()}')
        len_pages.set(string_pages_papers)

    container = VerticalScrolledFrame(dialog, height=550 if len(sorterClass.files_for_print) > 20 else (len(sorterClass.files_for_print) + 1) * 25, width=731)
    container.pack()
    filepathsForPrint = sorterClass.files_for_print
    filenamesForPrint = [os.path.basename(i) if len(os.path.basename(i)) < 58 else os.path.basename(i)[:55] + '...' for i in filepathsForPrint]
    printcbVariables = {}
    printcb = {}
    filenames = {}
    numpages = {}
    rbVariables = {}
    rbuttons1perPage = {}
    rbuttons2perPage = {}
    rbuttons4perPage = {}
    previewbtns = {}
    Label(container, text='Название документа').grid(column=1, row=0)
    Label(container, text='Страниц').grid(column=2, row=0)
    Label(container, text='1').grid(column=3, row=0)
    Label(container, text='2').grid(column=4, row=0)
    Label(container, text='4').grid(column=5, row=0)
    for i in range(len(filepathsForPrint)):
        printcbVariables[filepathsForPrint[i]] = BooleanVar()
        printcbVariables[filepathsForPrint[i]].set(1)
        printcb[i] = Checkbutton(container, variable=printcbVariables[filepathsForPrint[i]], command=update_num_pages)
        printcb[i].grid(column=0, row=i + 1, sticky=W)
        filenames[i] = Label(container, text=filenamesForPrint[i], font='TkFixedFont')
        filenames[i].grid(column=1, row=i + 1, sticky=W)
        filenames[i].bind('<Double-Button-1>', lambda event, a=filepathsForPrint[i]: os.startfile(a))
        numpages[i] = Label(container, text=str(sorterClass.num_pages[filepathsForPrint[i]][0]), padx=2)
        numpages[i].grid(column=2, row=i + 1)
        rbVariables[filepathsForPrint[i]] = IntVar()
        rbVariables[filepathsForPrint[i]].set(1)
        rbuttons1perPage[i] = Radiobutton(container, variable=rbVariables[filepathsForPrint[i]], value=1, command=update_num_pages)
        rbuttons1perPage[i].grid(column=3, row=i + 1, sticky=W)
        rbuttons2perPage[i] = Radiobutton(container, variable=rbVariables[filepathsForPrint[i]], value=2, command=update_num_pages)
        rbuttons2perPage[i].grid(column=4, row=i + 1, sticky=W)
        rbuttons4perPage[i] = Radiobutton(container, variable=rbVariables[filepathsForPrint[i]], value=4, command=update_num_pages)
        rbuttons4perPage[i].grid(column=5, row=i + 1, sticky=W)
        previewbtns[i] = Label(container, text='Предпросмотр', padx=2)
        previewbtns[i].grid(column=6, row=i + 1)
        previewbtns[i].bind('<Button-1>', lambda event, a=filepathsForPrint[i]: os.startfile(multiplePagesPerSheet(a, rbVariables[a].get())))
    bottom_actions = Frame(dialog)
    bottom_actions.pack()
    len_pages = StringVar()
    num_docs_for_print = IntVar()
    full_len_pages_for_print_var = IntVar()
    eco_dupl_len_for_print_var = IntVar()
    full_dupl_len_for_print_var = IntVar()
    update_num_pages()
    if sorterClass.save_stat == "yes":
        statsaver = BooleanVar()
        statsaver.set(1)
        save_to_stat_chkbtn = Checkbutton(bottom_actions, variable=statsaver, text='Добавить в статистику', command=lambda: print(statsaver.get()))
        save_to_stat_chkbtn.grid(column=0, row=0, sticky=S, padx=5, pady=2)
    open_folder_b = Label(bottom_actions, text=" Открыть папку ", borderwidth=2, relief="groove")
    open_folder_b.grid(column=1, row=0, sticky=S, padx=5, pady=2)
    open_folder_b.bind("<Button-1>", lambda event, a=os.path.dirname(filepathsForPrint[0]): subprocess.Popen(f'explorer {a}'))
    print_button = Label(bottom_actions, text=" Печать ", borderwidth=2, relief="groove")
    print_button.grid(column=2, row=0, sticky=S, padx=5, pady=2)
    print_button.bind("<Button-1>", apply_print)
    sum_pages = Label(bottom_actions, textvariable=len_pages)
    sum_pages.grid(column=3, row=0, sticky=S, padx=5, pady=2)


def info_show_printed():
    messagebox.showinfo("Готово", "Документы отправлены в очередь принтера.")


def warning_not_zip():
    messagebox.showwarning("Варнинг", "Загружен не Zip архив.")


def info_show_credits(e):
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

close_button = Label(title_bar, text=' X ')
close_button.pack(side=RIGHT)
close_button.bind("<Button-1>", quitter)

dropzone = Label(root)
dropzone.pack(fill=X)
dropzone.drop_target_register(DND_FILES)
dropzone.dnd_bind('<<Drop>>', main_drop)

opt1DelZip = StringVar()
opt1DelZip.set(sorterClass.deletezip)
opt2EcoMode = StringVar()
opt2EcoMode.set(sorterClass.paperecomode)
opt3Print = StringVar()
opt3Print.set(sorterClass.print_directly)
opt4DefPrinter = StringVar()
opt4DefPrinter.set(sorterClass.default_printer)
opt5SaveStat = StringVar()
opt5SaveStat.set(sorterClass.save_stat)

root.mainloop()
