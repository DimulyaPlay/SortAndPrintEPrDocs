from tkinter import *
from scrollable_frame import VerticalScrolledFrame
from sort_utils import *
import subprocess


def print_dialog(root, current_config, sorterClass, stat_writer, iconpath):
    dialog = Toplevel(root)
    dialog.iconbitmap(iconpath)
    dialog.title(f'Файлов на печать {len(sorterClass.files_for_print)}')
    dialog.attributes('-topmost', True)
    dialog.resizable(False, False)

    def apply_print(e):
        print_button.unbind("<Button-1>")
        print_button.config(relief=SUNKEN)
        print_button.update()
        stat_writer.statdict['Постановка в очередь заняла'] = 0
        for fp, prntcbvar in printcbVariables.items():
            if prntcbvar.get():
                to_queue_time = print_file(fp, rbVariables[fp].get(), current_config.default_printer,
                                           int(sorterClass.num_pages[fp][0] / rbVariables[fp].get()),
                                           entryCopyVariables[fp].get(), os.path.basename(fp))
                stat_writer.statdict['Постановка в очередь заняла'] += to_queue_time
                prntcbvar.set(0)
                lb1[fp].config(background='green1')
                lb1[fp].update()
        if current_config.save_stat == 'yes' and statsaver.get():
            stat_writer.statdict['Напечатано док-ов'] = num_docs_for_print.get()
            stat_writer.statdict[
                'Затрата без эко была бы'] = full_dupl_len_for_print_var.get() + eco_protocols_var.get()
            stat_writer.statdict['Затрачено листов'] = eco_dupl_len_for_print_var.get()
            stat_writer.statdict[
                'Сэкономлено листов'] = full_dupl_len_for_print_var.get() - eco_dupl_len_for_print_var.get() + eco_protocols_var.get()
            stat_writer.add_and_save_stats()
        update_num_pages()
        print_button.config(relief=RAISED)
        print_button.bind("<Button-1>", apply_print)

    def open_folder(e):
        open_folder_b.unbind("<Button-1>")
        open_folder_b.config(relief=SUNKEN)
        open_folder_b.update()
        os.startfile(os.path.dirname(filepathsForPrint[0]))
        time.sleep(0.1)
        open_folder_b.config(relief=RAISED)
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

    container = VerticalScrolledFrame(dialog, height=550 if len(sorterClass.files_for_print) > 20 else (
                                                                                                               len(sorterClass.files_for_print) + 1) * 25,
                                      width=675)
    container.pack()
    filepathsForPrint = sorterClass.files_for_print
    printcbVariables = {}
    rbVariables = {}
    entryCopyVariables = {}
    lb1 = {}
    prntchballvar = BooleanVar()
    prntchballvar.set(1)
    prntchball = Checkbutton(container, variable=prntchballvar, command=check_all_chbtns)
    prntchball.grid(column=0, row=0)
    Label(container, text='Название документа').grid(column=1, row=0)
    Label(container, text='Страниц').grid(column=2, row=0)
    Label(container, text='1').grid(column=3, row=0)
    Label(container, text='2').grid(column=4, row=0)
    Label(container, text='4').grid(column=5, row=0)
    Label(container, text='Коп').grid(column=6, row=0)
    current_row = 1
    for fp in filepathsForPrint:
        fn = os.path.basename(fp)
        fn = fn if len(fn) < 58 else fn[:55] + '...'
        printcbVariables[fp] = BooleanVar()
        printcbVariables[fp].set(1)
        prntchb = Checkbutton(container, variable=printcbVariables[fp], command=update_num_pages)
        prntchb.grid(column=0, row=current_row, sticky=W)
        lb1[fp] = Label(container, text=fn, font='TkFixedFont')
        lb1[fp].grid(column=1, row=current_row, sticky=W)
        lb1[fp].bind('<Double-Button-1>', lambda event, a=fp: os.startfile(a))
        lb2 = Label(container, text=str(sorterClass.num_pages[fp][0]), padx=2)
        lb2.grid(column=2, row=current_row)
        rbVariables[fp] = IntVar()
        rbVariables[fp].set(1)
        rb1 = Radiobutton(container, variable=rbVariables[fp], value=1, command=update_num_pages)
        rb1.grid(column=3, row=current_row, sticky=W)
        rb2 = Radiobutton(container, variable=rbVariables[fp], value=2, command=update_num_pages)
        rb2.grid(column=4, row=current_row, sticky=W)
        rb4 = Radiobutton(container, variable=rbVariables[fp], value=4, command=update_num_pages)
        rb4.grid(column=5, row=current_row, sticky=W)
        entryCopyVariables[fp] = IntVar()
        entryCopyVariables[fp].set(1)
        entryCopies = Entry(container, textvariable=entryCopyVariables[fp], width=5)
        entryCopies.grid(column=6, row=current_row, sticky=W)
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
        save_to_stat_chkbtn = Checkbutton(bottom_actions, variable=statsaver, text='Добавить в статистику',
                                          command=lambda: print(statsaver.get()))
        save_to_stat_chkbtn.grid(column=0, row=0, sticky=S, padx=5, pady=2)
    open_folder_b = Label(bottom_actions, text=" Открыть папку ", borderwidth=2, relief=RAISED)
    open_folder_b.grid(column=1, row=0, sticky=S, padx=5, pady=2)
    open_folder_b.bind("<Button-1>", open_folder)
    print_button = Label(bottom_actions, text=" Печать ", borderwidth=2, relief=RAISED)
    print_button.grid(column=2, row=0, sticky=S, padx=5, pady=2)
    print_button.bind("<Button-1>", apply_print)
    sum_pages = Label(bottom_actions, textvariable=len_pages)
    sum_pages.grid(column=3, row=0, sticky=S, padx=5, pady=2)
