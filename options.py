from tkinter import *
import win32print

# ver = '3.4.4'
# ver = '1.0.10_TRON'
ver = '1.0_5_JPrinterVer, 0.5_JavaUtils'
curdate = '2022/09/15'


def open_settings(root, current_config, statfile_path, iconpath, stat_loader):
    printer_list = [i[2] for i in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]  # список принтеров в системе
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


def info_show_credits(e):
    messagebox.showinfo("Кредитс",
                        message=f"Сортировка документов с сайта Электронное провосудие.\nАвтор: консультант Краснокамского гс "
                                f"Соснин Дмитрий.\nВерсия {ver} от {curdate}")
