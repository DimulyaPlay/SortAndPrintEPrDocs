import sys
from tkinter import *
from tkinterdnd2 import *
from sorter_class import *
import configparser

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

# пути к скрипту и конфигу
config_name = 'config.ini'
stats_name = 'epr_stats.ini'
PDF_PRINT_NAME = 'PDFtoPrinter.exe'

stats_path = os.path.join(application_path, stats_name)
config_path = os.path.join(application_path, config_name)
PDF_PRINT_FILE = os.path.join(application_path, PDF_PRINT_NAME)

default_config = {'delete_zip': 'no',
                  'paper_eco_mode': 'yes',
                  'print_directly': 'no',
                  'PDF_PRINT_PATH': PDF_PRINT_FILE,
                  }


def readcreateconfig(default_config, config_path):
    # создание конфига или открыие имеющегося
    config = configparser.ConfigParser()
    if not os.path.exists(config_path):
        config['DEFAULT'] = default_config
        with open(config_path, 'w') as configfile:
            config.write(configfile)
    else:
        config.read(config_path)
    return config


config = readcreateconfig(default_config, config_path)
sorter = main_sorter(config=config, config_path=config_path)


def main_drop(event):
    if ' ' in event.data:
        path = event.data[1:-1]
    else:
        path = event.data
    if path[-4:] != '.zip':
        not_zip()
        return
    sorter.agregate_file(path)
    if sorter.print_directly:
        print_dialog()


def move_app(e):
    root.geometry(f'+{e.x_root}+{e.y_root}')


def quitter(e):
    root.quit()
    root.destroy()


def apply(sorter_obj=sorter):
    # Set main class vars from checkbuttons
    sorter_obj.deletezip = opt1.get()
    sorter_obj.paperecomode = opt2.get()
    sorter_obj.print_directly = opt3.get()
    sorter_obj.write_config_to_file()


def show_settings(e):
    newWindow = Toplevel(root)
    newWindow.title("Параметры")
    newWindow.geometry('140x100')
    c1 = Checkbutton(newWindow, text="Удалить Zip",
                     variable=opt1,
                     onvalue=1, offvalue=0, command=apply)
    c1.pack(anchor=W)
    c2 = Checkbutton(newWindow, text="Эко режим",
                     variable=opt2,
                     onvalue=1, offvalue=0, command=apply)
    c2.pack(anchor=W)
    c3 = Checkbutton(newWindow, text="Печать на принтер",
                     variable=opt3,
                     onvalue=1, offvalue=0, command=apply)
    c3.pack(anchor=W)
    label = Label(newWindow, text=" Кредитс ", borderwidth=2, relief="groove")
    label.pack(anchor=S)
    label.bind("<Button-1>", show_credits)


def print_dialog():
    dialog = Toplevel(root)
    dialog.title(f'Файлов на печать {len(sorter.files_for_print)}')
    dialog.attributes('-topmost', True)

    def apply_print(e):
        for i, j in cbVariables.items():
            if j.get():
                print_file(i, PDF_PRINT_FILE)
        show_printed()

    def update_num_pages():
        len_pages.set(sum([sorter.num_pages[winSt[i]] for i in range(len(winSt)) if cbVariables[winSt[i]].get()]))

    winSt = sorter.files_for_print
    winSt_names = [os.path.basename(i) for i in winSt]
    cbVariables = {}
    cb = {}
    lb = {}
    lbstr = {}
    lb_pages = Label(dialog, text='Страниц')
    lb_pages.grid(column=2, row=0)
    lb_names = Label(dialog, text='Название документа')
    lb_names.grid(column=1, row=0)
    len_pages = IntVar()
    for i in range(len(winSt)):
        cbVariables[winSt[i]] = BooleanVar()
        cbVariables[winSt[i]].set(1)
        cb[i] = Checkbutton(dialog, variable=cbVariables[winSt[i]], command=update_num_pages)
        cb[i].grid(column=0, row=i+1, sticky=W)
        lb[i] = Label(dialog, text=winSt_names[i])
        lb[i].grid(column=1, row=i+1, sticky=W)
        lb[i].bind('<Double-Button-1>', lambda event, a=winSt[i]: os.startfile(a))
        lbstr[i] = Label(dialog, text=str(sorter.num_pages[winSt[i]]), padx=2)
        lbstr[i].grid(column=2, row=i+1)
    len_pages.set(sum([sorter.num_pages[winSt[i]] for i in range(len(winSt)) if cbVariables[winSt[i]].get()]))
    print_b = Label(dialog, text=" Печать ", borderwidth=2, relief="groove")
    print_b.grid(column=1, row=i+2, sticky=S)
    print_b.bind("<Button-1>", apply_print)
    sum_pages = Label(dialog, textvariable=len_pages)
    sum_pages.grid(column=2, row=i+2, sticky=S)


def show_printed():
    messagebox.showinfo("Готово", "Документы отправлены в очередь принтера.")


def not_zip(e):
    messagebox.showinfo("Варнинг", "Загружен не Zip архив.")


def show_credits(e):
    messagebox.showinfo("Кредитс",
                        "Сортировка документов с сайта Электронное провосудие.\nАвтор: консультант Краснокамского гс "
                        "Соснин Дмитрий.\nВерсия 2.1")


root = Tk()
root.geometry('35x45')
root.overrideredirect(True)
root.attributes('-topmost', True)


opt1 = BooleanVar()
opt1.set(sorter.deletezip)
opt2 = BooleanVar()
opt2.set(sorter.paperecomode)
opt3 = BooleanVar()
opt3.set(sorter.print_directly)

title_bar = Frame(root, bd=1)
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
root.mainloop()
