import time

import win32api
import win32print
from PyPDF2 import PdfFileReader, PdfFileWriter
from difflib import SequenceMatcher
import pdfplumber
import os
import win32com.client


def concat_pdfs(main_pdf_filepath, slave_pdf_filepath):
    # присоединение второго пдф файла к первому
    file_writer = PdfFileWriter()
    broken = False
    outpath = main_pdf_filepath
    try:
        file_main = PdfFileReader(main_pdf_filepath, strict=False)
    except:
        broken = True
    if not broken:
        file_slave = PdfFileReader(slave_pdf_filepath, strict=False)
        file_writer.appendPagesFromReader(file_main)
        file_writer.appendPagesFromReader(file_slave)
        outpath = f"{main_pdf_filepath[:-4]}+protocol.pdf"
        with open(outpath, 'wb') as out:
            file_writer.write(out)
    return outpath, broken


def similar(a: str, b: str) -> float:
    # принимает две строки и возвращет коэффициент схожести
    return SequenceMatcher(None, a, b).ratio()


def extracttext(path):
    # принимает строку - путь к файлу пдф, возвращает извлеченный текст
    with pdfplumber.open(path[0]) as pdf:
        all_text = ''
        for i in pdf.pages:
            all_text += i.extract_text()
    return all_text.replace('\n', '')


def check_num_pages(path):
    # принимает строку - путь к файлу пдф, возвращает кол-во страниц
    with pdfplumber.open(path) as pdf:
        pages = len(pdf.pages)
    return pages


def wordpdf(origfile):
    # конвертация word в pdf открывает копию, и сохраняет в ориг
    convfile = f'{origfile}.pdf'
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(origfile)
    doc.SaveAs(convfile, FileFormat=17)
    doc.Close()
    word.Quit()
    os.remove(origfile)
    neworigfile = f'{origfile.rsplit(".")[0]}.pdf'
    os.rename(convfile, neworigfile)
    return neworigfile


def print_file(filepath, exe_path):
    # Печать файла через консольную утилиту.
    # Принимает строки - путь к пдф и путь к утилите. Открывает утилиту, печатает и дожидается
    # пока документ не будет напечатан
    currentprinter = win32print.GetDefaultPrinter()  # можно будет задать свой принтер для печати
    win32api.ShellExecute(0, 'open', exe_path,
                          '/s ' + filepath,
                          '.', 0)
    jobs = [1]
    time.sleep(1.5)
    while jobs:
        time.sleep(0.01)
        jobs = []
        phandle = win32print.OpenPrinter(currentprinter)
        print_jobs = win32print.EnumJobs(phandle, 0, -1, 1,)
        print(print_jobs)
        if print_jobs:
            # for job in print_jobs:
            #     if job['Status'] != 148:
                    jobs.extend(list(print_jobs))
                    for j in print_jobs:
                        print(j['Status'])
        win32print.ClosePrinter(phandle)



