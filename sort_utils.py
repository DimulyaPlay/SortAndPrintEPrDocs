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
        for i in range(len(file_main.pages)):
            page = file_main.getPage(i)
            # print(page.mediaBox)
            if page.mediaBox[2] > page.mediaBox[3]:
                file_writer.addPage(page.rotateClockwise(90))
            else:
                file_writer.addPage(page)
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
    jobs = [0, 0, 0, 0, 0]
    while sum(jobs) < 3:
        time.sleep(0.005)
        phandle = win32print.OpenPrinter(currentprinter)
        print_jobs = win32print.EnumJobs(phandle, 0, -1, 1, )
        docs_in_queue = {job['pDocument']: job['Status'] for job in print_jobs}
        # print(docs_in_queue)
        file_printing = os.path.basename(filepath)
        if file_printing in docs_in_queue.keys() and jobs[0] != 1:
            jobs[0] = 1
        if file_printing in docs_in_queue.keys() and jobs[1] != 1:
            if docs_in_queue[file_printing] == 8:
                jobs[1] = 1
        if file_printing in docs_in_queue.keys() and jobs[2] != 1:
            if docs_in_queue[file_printing] == 0:
                jobs[2] = 1
        if file_printing in docs_in_queue.keys() and jobs[3] != 1:
            if docs_in_queue[file_printing] == 8208:
                jobs[3] = 1
        if file_printing not in docs_in_queue.keys() and jobs[0] == 1:
            jobs[4] = 1
        print(jobs)
        win32print.ClosePrinter(phandle)
    time.sleep(0.5)
    os.system("taskkill /im pdftoprinter.exe")
