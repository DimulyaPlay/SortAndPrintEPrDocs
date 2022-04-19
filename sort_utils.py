import time

import win32api
import win32print
from PyPDF2 import PdfFileReader, PdfFileWriter
import PyPDF2
from difflib import SequenceMatcher
import pdfplumber
import os
import win32com.client
a4orig = [595.0, 842.0]
a4 = [566.0, 800.0]


def concat_pdfs(main_pdf_filepath, slave_pdf_filepath, print_directly):
    # присоединение второго пдф файла к первому
    file_writer = PdfFileWriter()
    broken = False
    outpath = main_pdf_filepath
    try:
        file_main = PdfFileReader(main_pdf_filepath, strict=False)
        print(main_pdf_filepath, 'opened')
    except:
        broken = True
        print('broken')
    if not broken:
        file_slave = PdfFileReader(slave_pdf_filepath, strict=False)
        print(slave_pdf_filepath, 'opened')
        if print_directly == 'yes':
            for i in range(len(file_main.pages)):
                page = file_main.getPage(i)
                mb = page.mediaBox[2:]
                if mb[0] > mb[1]:
                    page = page.rotateClockwise(270)
                mb = page.mediaBox[2:]
                if mb[0] > a4orig[0] or mb[1] > a4orig[1]:
                    hor_koef = a4[0] / float(mb[0])
                    ver_koef = a4[1] / float(mb[1])
                    min_koef = min([hor_koef, ver_koef])
                    page.scaleBy(min_koef)
                    oldpage = page
                    page = PyPDF2.pdf.PageObject.createBlankPage(width=595.2, height=842.88)
                    padx = oldpage.mediaBox[2]/2
                    pady = oldpage.mediaBox[3]/2
                    page.mergeTranslatedPage(oldpage, 300-padx, 420-pady)
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
    try:
        with pdfplumber.open(path) as pdf:
            pages = len(pdf.pages)
    except:
        print("num pages error, set to 2", path)
        pages = 2
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


def print_file(filepath, exe_path, currentprinter):
    # Печать файла через консольную утилиту.
    # Принимает строки - путь к пдф и путь к утилите. Открывает утилиту, печатает и дожидается
    # пока документ не будет напечатан
    win32api.ShellExecute(0, 'open', exe_path,
                          '/s ' + filepath + ' "'+currentprinter+'" ',
                          '.', 0)
    jobs = [0, 0, 0, 0, 0]
    while sum(jobs) < 3:
        time.sleep(0.01)
        phandle = win32print.OpenPrinter(currentprinter)
        print_jobs = win32print.EnumJobs(phandle, 0, -1, 1, )
        docs_in_queue = {job['pDocument']: job['Status'] for job in print_jobs}
        # print(docs_in_queue)
        file_printing = os.path.basename(filepath)
        if file_printing in docs_in_queue.keys() and jobs[0] != 1:  # "в списке" flag
            jobs[0] = 1
        if file_printing in docs_in_queue.keys() and jobs[1] != 1:  # "постановка в очередь" flag
            if docs_in_queue[file_printing] == 8:
                jobs[1] = 1
        if file_printing in docs_in_queue.keys() and jobs[2] != 1:  # "в очереди" flag
            if docs_in_queue[file_printing] == 0:
                jobs[2] = 1
        if file_printing in docs_in_queue.keys() and jobs[3] != 1:  # "печатается" flag
            if docs_in_queue[file_printing] == 8208:
                jobs[3] = 1
        if file_printing not in docs_in_queue.keys() and jobs[0] == 1:  # "напечатан" flag
            jobs[4] = 1
        # print(jobs)
        win32print.ClosePrinter(phandle)
    os.system("taskkill /im pdftoprinter.exe")
