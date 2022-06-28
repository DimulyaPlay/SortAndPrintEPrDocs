import time
import win32api
import win32print
from PyPDF2 import PdfFileReader, PdfFileWriter
import PyPDF2
from pikepdf import Pdf
from difflib import SequenceMatcher
import pdfplumber
import os
import subprocess
import win32com.client
import tempfile
from PIL import Image

a4orig = [612.1, 842.0]
a4small = [i * 0.95 for i in a4orig]


def concat_pdfs(main_pdf_filepath, slave_pdf_filepath):
    # присоединение второго пдф файла к первому
    file_main = Pdf.open(main_pdf_filepath)
    file_slave = Pdf.open(slave_pdf_filepath)
    file_main.pages.extend(file_slave.pages)
    outpath = f"{main_pdf_filepath[:-4]}+protocol.pdf"
    file_main.save(outpath)
    file_main.close()
    file_slave.close()
    return outpath


def fitPdfInA4(pdfpath):
    pdf = PdfFileReader(pdfpath)
    new_pdf = PdfFileWriter()
    for page in pdf.pages:
        page_width = page.mediaBox.getWidth()
        page_height = page.mediaBox.getHeight()
        if page_width > page_height:
            page.rotateClockwise(270)
        page_width = page.mediaBox.getWidth()
        page_height = page.mediaBox.getHeight()
        if page_width > a4small[0] or page_height > a4small[1]:
            hor_koef = a4small[0] / float(page_width)
            ver_koef = a4small[1] / float(page_height)
            min_koef = min([hor_koef, ver_koef])
            page.scaleBy(min_koef)
            oldpage = page
            page = PyPDF2.pdf.PageObject.createBlankPage(width=612.1, height=842.0)
            padx = oldpage.mediaBox.getWidth() / 2
            pady = oldpage.mediaBox.getHeight() / 2
            page.mergeTranslatedPage(oldpage, 306 - padx, 421 - pady)
        new_pdf.addPage(page)
    fd, outpath = tempfile.mkstemp('.pdf')
    os.close(fd)
    with open(outpath, mode='wb') as export:
        new_pdf.write(export)
    return outpath


def multiplePagesPerSheet(filepath, mode):
    if mode == 1:
        return fitPdfInA4(filepath)
    merged_file = PdfFileWriter()
    if mode == 2:
        rotated_pdf = fitPdfInA4(filepath)
        orig_file = PdfFileReader(rotated_pdf, strict=False)
        n_pages = len(orig_file.pages)
        for i in range(0, n_pages, 2):
            big_page = PyPDF2.pdf.PageObject.createBlankPage(width=595.2, height=842.88)
            big_page.mergeRotatedScaledTranslatedPage(orig_file.pages[i],
                                                      rotation=90,
                                                      scale=0.7,
                                                      tx=585.2,
                                                      ty=10)
            try:
                big_page.mergeRotatedScaledTranslatedPage(orig_file.pages[i + 1],
                                                          rotation=90,
                                                          scale=0.7,
                                                          tx=585.2,
                                                          ty=420)
            except:
                pass
            merged_file.addPage(big_page)
    if mode == 4:
        rotated_pdf = fitPdfInA4(filepath)
        orig_file = PdfFileReader(rotated_pdf, strict=False)
        n_pages = len(orig_file.pages)
        for i in range(0, n_pages, 4):
            big_page = PyPDF2.pdf.PageObject.createBlankPage(width=595.2, height=842.88)
            big_page.mergeScaledTranslatedPage(orig_file.pages[i],
                                               scale=0.48,
                                               tx=10,
                                               ty=411.44)
            try:
                big_page.mergeScaledTranslatedPage(orig_file.pages[i + 1],
                                                   scale=0.48,
                                                   tx=288,
                                                   ty=411.44)
            except:
                pass

            try:
                big_page.mergeScaledTranslatedPage(orig_file.pages[i + 2],
                                                   scale=0.48,
                                                   tx=10,
                                                   ty=10)
            except:
                pass
            try:
                big_page.mergeScaledTranslatedPage(orig_file.pages[i + 3],
                                                   scale=0.48,
                                                   tx=283,
                                                   ty=10)
            except:
                pass
            merged_file.addPage(big_page)
    fd, outpath = tempfile.mkstemp('.pdf')
    os.close(fd)
    with open(outpath, 'wb') as out:
        merged_file.write(out)
    return outpath


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
    pdf = Pdf.open(path)
    pages = len(pdf.pages)
    papers = int(pages / 2 + 0.9)
    pdf.close()
    return [pages, papers]


def wordpdf(origfile):
    # конвертация word в pdf открывает копию, и сохраняет в ориг
    convfile = f'{origfile}.pdf'
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(origfile)
    doc.SaveAs(convfile, FileFormat=17)
    doc.Close()
    word.Quit()
    os.remove(origfile)
    neworigfile = f'{origfile.rsplit(".",1)[0]}.pdf'
    try:
        os.rename(convfile, neworigfile)
    except:
        neworigfile = f'{origfile.rsplit(".",1)[0]}..pdf'
        os.rename(convfile, neworigfile)
    return neworigfile


def imagepdf(origfile):
    # конвертация картинку в pdf открывает копию, и сохраняет в ориг
    convfile = f'{origfile}.pdf'
    workbook = load_workbook(origfile, guess_types=True, data_only=True)
    worksheet = workbook.active
    image = Image.open(origfile)
    image.convert('RGB')
    image.save(convfile)
    os.remove(origfile)
    neworigfile = f'{origfile.rsplit(".",1)[0]}.pdf'
    os.rename(convfile, neworigfile)
    return neworigfile


def print_file(filepath, exe_path, currentprinter):
    # Печать файла через консольную утилиту.
    # Принимает строки - путь к пдф и путь к утилите. Открывает утилиту, печатает и дожидается
    # пока документ не будет напечатан
    win32api.ShellExecute(0, 'open', exe_path,
                          '/s ' + '"' + filepath + '"' + ' "' + currentprinter + '" ',
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
        win32print.ClosePrinter(phandle)
    subprocess.call("taskkill.exe /im pdftoprinter.exe /f", shell=True)
