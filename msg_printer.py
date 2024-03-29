import shutil
from tkinter import *
import win32com.client
from scrollable_frame import VerticalScrolledFrame
from sort_utils import *
import copy


class MessageHandler:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.allowed_ext = ['.doc', '.docx', '.pdf', '.jpg', '.jpeg', '.tif', '.tiff', '.png', '.gif', '.rtf', '.ods',
                            '.odt', '.xlsx', '.xls']
        self.allowed_ext_img = ['.jpg', '.jpeg', '.tif', '.tiff', '.png', '.gif']
        self.allowed_ext_docs = ['.doc', '.docx', '.rtf', '.odt', '.ods', '.xlsx', '.xls']
        self.allowed_ext_archives = ['.7z', '.rar', '.zip']

    @staticmethod
    def get_file_from_attach(att):
        """
        Извлечение вложения, сохранение в temp
        :param att: outlook attachment object
        :return: путь к тепмфайлу, расшриение файла
        """
        ext = '.' + att.FileName.rsplit('.', 1)[1].lower()
        fd, outpath = tempfile.mkstemp(ext)
        os.close(fd)
        att.SaveAsFile(outpath)
        return outpath, ext

    def get_files_from_msg(self, msg):
        """
        Извлечение всех вложений из msg файла включая вложенные архивы
        :param msg: outlook message object
        :return: лист из путей ко всем извлеченным файлам
        """
        list_files = []
        for att in msg.attachments:
            filename = att.FileName
            outpath, ext = self.get_file_from_attach(att)
            if ext in self.allowed_ext_archives:
                filepaths, filenames = unpack_archieved_files(outpath, ext)
                for fp, fn in zip(filepaths, filenames):
                    if not os.path.isdir(fp):
                        list_files.append([fp, fn])
            list_files.append([outpath, filename])
        return list_files

    def convert_attachments(self, list_fp_fn):
        """
        Ковертация разрешенных форматов в пдф

        :param list_fp_fn: лист путей и имен файлов
        :return: лист формата [[путь, название, кол-во страниц, кол-во листов, возможна ли печать],[...]]
        """
        new_list_fp_fn_pgs_pps_isprnt = []
        for fp, fn in list_fp_fn:
            ext = '.' + fn.rsplit('.', 1)[1].lower()
            if ext in self.allowed_ext:
                if ext != '.pdf':
                    if ext in self.allowed_ext_img:
                        fp = convertImageWithJava(fp)
                        num_pgs, num_pps = check_num_pages(fp)
                        new_list_fp_fn_pgs_pps_isprnt.append([fp, fn, num_pgs, num_pps, 1])
                    if ext in self.allowed_ext_docs:
                        fp = office2pdf(fp)
                        num_pgs, num_pps = check_num_pages(fp)
                        new_list_fp_fn_pgs_pps_isprnt.append([fp, fn, num_pgs, num_pps, 1])
                if ext == '.pdf':
                    num_pgs, num_pps = check_num_pages(fp)
                    new_list_fp_fn_pgs_pps_isprnt.append([fp, fn, num_pgs, num_pps, 1])
            else:
                new_list_fp_fn_pgs_pps_isprnt.append([fp, fn, '?', '?', 0])
        return new_list_fp_fn_pgs_pps_isprnt

    def handle_messages(self, msgs):
        """
        Объединенный метод, формирующий:
        лист из ключей(пути к мсг файлам)
        словари {ключ: msg}, {ключ:вложения},{ключ:путь к оригинальному msg}
        :param msgs: list из путей к msg файлам
        """
        self.handle_keys = []
        self.handled_messages = {}
        self.handled_attachments = {}
        self.orig_messages = {}
        for i in msgs:
            if i == '':
                continue
            fd, outpath = tempfile.mkstemp('.msg')
            os.close(fd)
            shutil.copy2(i, outpath)
            self.orig_messages[outpath] = i
            msg = self.outlook.OpenSharedItem(fr'{outpath}')
            attachment_files = self.get_files_from_msg(msg)
            attachment_files = self.convert_attachments(attachment_files)
            self.handle_keys.append(outpath)
            self.handled_messages[outpath] = msg
            self.handled_attachments[outpath] = attachment_files

    def print_dialog_msg(self, root, current_config, iconpath):
        """
        Функция окна печати
        :param iconpath: путь к иконке окна
        :param root: родительское окно
        :param current_config: конфигурация
        """
        dialog = Toplevel(root)
        dialog.title(f'Файлов на печать ')
        dialog.iconbitmap(iconpath)
        dialog.attributes('-topmost', True)
        dialog.resizable(False, False)

        def update_num_pages():
            total_pages = 0
            total_papers = 0
            docs_for_print = 0
            for msg in self.handle_keys:
                if printcbVariables[msg].get():
                    docs_for_print += 1
                    total_pages += 1
                    total_papers += 1
                for att in self.handled_attachments[msg]:
                    try:
                        is_checked = printcbVariables[att[0]].get()
                    except:
                        is_checked = 0
                    if is_checked:
                        docs_for_print += 1
                        total_pages += att[2]
                        total_papers += int(att[3] / rbVariables[att[0]].get() + 0.9)
            string_pages_papers = f"Всего для печати страниц: {total_pages}, листов: {total_papers}"
            len_pages.set(string_pages_papers)
            dialog.title(f'Документов на печать {docs_for_print}')

        def apply_print(e):
            print_button.unbind("<Button-1>")
            print_button.config(relief=SUNKEN)
            print_button.update()
            group_cycle_real = group_cycle[1:]
            group_cycle_print_list_sorted = copy.deepcopy(group_cycle[1:])  # other methods can't make indipendent copy
            for msg in self.handle_keys:
                if printcbVariables[msg].get():
                    try:
                        copies = int(entryCopyVariables[msg].get())
                    except:
                        copies = 1
                    for i in range(copies):  # кол-во копий печатать
                        self.handled_messages[msg].PrintOut()
                    printcbVariables[msg].set(False)
                    lb1[msg].config(background='green1')
                    lb1[msg].update()
                for att in self.handled_attachments[msg]:
                    try:
                        printcbVariables[att[0]].get()
                    except:
                        continue
                    if printcbVariables[att[0]].get():
                        doc_n_group = None
                        for n, group in enumerate(group_cycle_real):
                            if att[0] in group:
                                doc_n_group = n
                                break
                        if doc_n_group is not None:
                            group_cycle_real[doc_n_group].remove(att[0])
                            if not group_cycle_real[doc_n_group]:
                                groupped_file = concat_pdfs(group_cycle_print_list_sorted[doc_n_group], False)
                                print_file(groupped_file, rbVariables[att[0]].get(), current_config.default_printer,
                                           int(att[2] / rbVariables[att[0]].get()), entryCopyVariables[att[0]].get(),
                                           att[1])
                        else:
                            print_file(att[0], rbVariables[att[0]].get(), current_config.default_printer,
                                       int(att[2] / rbVariables[att[0]].get()), entryCopyVariables[att[0]].get(),
                                       att[1])
                        printcbVariables[att[0]].set(0)
                        lb1[att[0]].config(background='green1')
                        lb1[att[0]].update()
            update_num_pages()
            print_button.config(relief=RAISED)
            print_button.bind("<Button-1>", apply_print)

        def check_all_chbtns():
            if prntchballvar.get():
                for chbtn in printcbVariables.values():
                    chbtn.set(1)
            else:
                for chbtn in printcbVariables.values():
                    chbtn.set(0)
            update_num_pages()

        def change_concat_category(e):
            widget = e.widget
            forward = False if e.num == 3 else True
            keys = list(but1.keys())
            values = list(but1.values())
            found_index = values.index(widget)
            doc_for_concat = keys[found_index]
            doc_n_group = None
            for n, group in enumerate(group_cycle):
                if doc_for_concat in group:
                    doc_n_group = n
                    break
            if doc_n_group is None:
                group_cycle[0].append(doc_for_concat)
            for n, group in enumerate(group_cycle):
                if doc_for_concat in group:
                    group.remove(doc_for_concat)
                    if n == 9 and forward:
                        group_cycle[0].append(doc_for_concat)
                        widget.config(background=group_colors[0])
                        widget.update()
                        break
                    if forward and 9 != n:
                        group_cycle[n + 1].append(doc_for_concat)
                        widget.config(background=group_colors[n + 1])
                        widget.update()
                    if not forward:
                        group_cycle[n - 1].append(doc_for_concat)
                        widget.config(background=group_colors[n - 1])
                        widget.update()
                    break

        MAXHEIGHT = 650
        height = 1
        width = 0
        for msg in self.handle_keys:
            height += 1
            if width < 68:
                msg_len = len(self.handled_messages[msg].subject)
                if width < msg_len:
                    width = msg_len
            for att in self.handled_attachments[msg]:
                if width < 68:
                    att_len = len(att[1])
                    if width < att_len:
                        width = att_len
                height += 1
        height = (height * 25) + 10
        if width < 26:
            width = 26
        if width > 68:
            width = 68
        width = (width * 8) + 230
        if height > MAXHEIGHT:
            height = MAXHEIGHT
        container = VerticalScrolledFrame(dialog, height=height, width=width)
        container.pack()
        rbVariables = {}
        entryCopyVariables = {}
        but1 = {}
        lb1 = {}
        printcbVariables = {}
        group0_print_list = []
        group1_print_list = []
        group2_print_list = []
        group3_print_list = []
        group4_print_list = []
        group5_print_list = []
        group6_print_list = []
        group7_print_list = []
        group8_print_list = []
        group9_print_list = []
        group_cycle = [group0_print_list, group1_print_list, group2_print_list, group3_print_list, group4_print_list,
                       group5_print_list, group6_print_list, group7_print_list, group8_print_list, group9_print_list]
        group_colors = ['SystemButtonFace', 'aquamarine1', 'brown2', 'deep sky blue', 'indian red', 'dark slate gray',
                        'cyan3', 'SeaGreen1', 'firebrick1', 'DarkOrchid1', 'DodgerBlue2']
        prntchballvar = BooleanVar()
        prntchballvar.set(1)
        prntchball = Checkbutton(container, variable=prntchballvar, command=check_all_chbtns)
        prntchball.grid(column=0, row=0)
        Label(container, text='Название документа/тема').grid(column=2, row=0)
        Label(container, text='Страниц').grid(column=3, row=0)
        Label(container, text='1').grid(column=4, row=0)
        Label(container, text='2').grid(column=5, row=0)
        Label(container, text='4').grid(column=6, row=0)
        Label(container, text='Коп').grid(column=7, row=0)
        currentrow = 1
        for filepath in self.handle_keys:
            subject = self.handled_messages[filepath].subject if self.handled_messages[
                                                                     filepath].subject != '' else 'Пустая тема'
            subject = subject if len(subject) < 68 else subject[:65] + "..."
            printcbVariables[filepath] = BooleanVar()
            printcbVariables[filepath].set(1)
            prntchb = Checkbutton(container, variable=printcbVariables[filepath], command=update_num_pages)
            prntchb.var = printcbVariables[filepath]
            prntchb.grid(column=0, row=currentrow, sticky=W)
            lb1[filepath] = Label(container, text=subject, font='TkFixedFont', fg='blue')
            lb1[filepath].grid(column=2, row=currentrow, sticky=W)
            lb1[filepath].bind('<Double-Button-1>', lambda event, a=self.orig_messages[filepath]: os.startfile(a))
            entryCopyVariables[filepath] = StringVar()
            entryCopyVariables[filepath].set(1)
            entryCopies = Entry(container, textvariable=entryCopyVariables[filepath], width=5)
            entryCopies.grid(column=7, row=currentrow, sticky=W)
            currentrow += 1
            for att in self.handled_attachments[filepath]:
                current_key = att[0]
                current_name = att[1]
                current_pages = att[2]
                current_papers = att[3]
                current_printable = att[4]
                current_name = current_name if len(current_name) < 58 else current_name[:55] + "..."
                if current_printable:
                    printcbVariables[current_key] = BooleanVar()
                    printcbVariables[current_key].set(1)
                    prntchb = Checkbutton(container, variable=printcbVariables[current_key], command=update_num_pages)
                    prntchb.var = printcbVariables[current_key]
                else:
                    prntchb = Checkbutton(container, state=DISABLED)
                prntchb.grid(column=0, row=currentrow, sticky=W)
                if not current_printable:
                    but1[current_key] = Button(container, width=1, height=1, state=DISABLED)
                    but1[current_key].grid(column=1, row=currentrow, sticky=E)
                else:
                    but1[current_key] = Button(container, width=1, height=1)
                    but1[current_key].grid(column=1, row=currentrow, sticky=E)
                    but1[current_key].bind('<Button>', change_concat_category)
                lb1[current_key] = Label(container, text=current_name, font='TkFixedFont')
                lb1[current_key].grid(column=2, row=currentrow, sticky=W, padx=10)
                lb1[current_key].bind('<Double-Button-1>', lambda event, a=current_key: os.startfile(a))
                lb2 = Label(container, text=current_pages)
                lb2.grid(column=3, row=currentrow, sticky=N)
                rbVariables[current_key] = IntVar()
                rbVariables[current_key].set(1)
                rb1 = Radiobutton(container, variable=rbVariables[current_key], value=1, command=update_num_pages)
                rb1.var = rbVariables[current_key]
                rb1.grid(column=4, row=currentrow, sticky=W)
                rb2 = Radiobutton(container, variable=rbVariables[current_key], value=2, command=update_num_pages)
                rb2.var = rbVariables[current_key]
                rb2.grid(column=5, row=currentrow, sticky=W)
                rb4 = Radiobutton(container, variable=rbVariables[current_key], value=4, command=update_num_pages)
                rb4.var = rbVariables[current_key]
                rb4.grid(column=6, row=currentrow, sticky=W)
                entryCopyVariables[current_key] = StringVar()
                entryCopyVariables[current_key].set(1)
                entryCopies = Entry(container, textvariable=entryCopyVariables[current_key], width=5)
                entryCopies.grid(column=7, row=currentrow, sticky=W)
                currentrow += 1
        bottom_actions = Frame(dialog)
        bottom_actions.pack()
        len_pages = StringVar()
        print_button = Label(bottom_actions, text=" Печать ", borderwidth=2, relief=RAISED)
        print_button.grid(column=2, row=0, sticky=S, padx=5, pady=2)
        print_button.bind("<Button-1>", apply_print)
        sum_pages = Label(bottom_actions, textvariable=len_pages)
        sum_pages.grid(column=3, row=0, sticky=S, padx=5, pady=2)
        update_num_pages()
