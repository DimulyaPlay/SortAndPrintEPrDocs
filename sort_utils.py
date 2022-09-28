import glob
import os
import tempfile
import time
from difflib import SequenceMatcher
import zipfile
import rarfile
import py7zr
import PyPDF2
import pdfplumber
import win32com
from PyPDF2 import PdfFileReader, PdfFileWriter
from py4j.java_gateway import JavaGateway
import py4j.java_collections

gateway = JavaGateway()
a4orig = [612.1, 842.0]  # оригинальный формат А4
a4small = [i * 0.95 for i in a4orig]  # размер для масштабирования под область печати


def similar(a: str, b: str) -> float:
    """
    Сравенение двух последовательностей для сравнения имен файлов
    :param a: Первая последовательность
    :param b: Вторая последовательность
    :return: Коэффициент схожести последовательностей
    """
    return SequenceMatcher(None, a, b).ratio()


def extracttext(path):
    """
    Извлечение текста из pdf
    :param path: путь к файлу
    :return: текст файла
    """
    all_text = gateway.entry_point.extractTextFromPdf(path)
    return all_text


def check_num_pages(path):
    """
    Рассчет количества страниц и листов в документе

    :param path: путь к файлу
    :return: лист - страниц, листов
    """
    try:
        doc = pdfplumber.open(path)
    except:
        print('cant open non pdf for check pages')
        return [0, 0]
    n_pages = len(doc.pages)
    if n_pages == 0:
        n_pages += 1
    doc.close()
    pages = n_pages
    papers = int(pages / 2 + 0.9)
    return [pages, papers]


def splitBy10(filepath, n_pages):
    reader = PdfFileReader(filepath)
    splitter = round((n_pages / 10) + 0.5)
    filepaths = []
    for i in range(splitter):
        writer = PdfFileWriter()
        for j in range(i * 10, (i + 1) * 10):
            if j < n_pages:
                writer.addPage(reader.pages[j])
        outpath = f"{filepath}_{i}th.pdf"
        with open(outpath, mode='wb') as export:
            writer.write(export)
        filepaths.append(outpath)
    return filepaths


def concat_pdfs(list_of_filepaths, is_del):
    """
    Конкатенация пдф файлов

    :return: str путь к обьединенному пдф
    """
    for i in list_of_filepaths:
        if not i.endswith('.pdf'):
            print('non pdf found')
            return list_of_filepaths[0]
    object_class = gateway.jvm.java.lang.String
    MyJavaArray = gateway.new_array(object_class, len(list_of_filepaths))
    for i in range(len(list_of_filepaths)):
        MyJavaArray[i] = list_of_filepaths[i]
    out = gateway.entry_point.concatenatePdfs(MyJavaArray, bool(is_del))
    return out


def print_file(filepath, mode, currentprinter, n_pages, copies, fileName='Empty', preview=False):
    """
    Отправка документа в очередь печати с заданными параметрами

    :param preview: print on not print
    :param copies: копий
    :param n_pages: количество страниц в пдф
    :param fileName: имя, которое сохранится в логе spooler
    :param filepath: путь к файлу
    :param mode: расположение страниц на листе
    :param currentprinter: название принтера
    """
    if not filepath.endswith('.pdf'):
        print('cant print non pdf file(print_file error)')
        return 0
    if mode == 1:
        try:
            filepath = fitPdfInA4(filepath)
        except:
            print('Fitting error with: ', fileName)
            pass
    if mode == 2:
        filepath = twoUP(filepath)
    if mode == 4:
        filepath = fourUP(filepath)
    starttime = time.time()
    if n_pages > 10:
        filepaths = splitBy10(filepath, n_pages)
    else:
        filepaths = [filepath]
    try:
        copies = int(copies)
    except:
        copies = 1
    if not preview:
        for i in range(copies):
            for filepath in filepaths:
                gateway.entry_point.printToPrinter(filepath, currentprinter, fileName)
    deltatime = time.time() - starttime
    return deltatime, filepaths


def parse_names(names: str):
    """
    Разбор входящей строкии на имена мсг файлов
    :param names: входящая строка

    :return: лист из путей к файлам
    """
    namesstart = 0
    nameslist = []
    while namesstart != -1:
        namesstart = names.find('C:/')
        namesend = names.find('.msg')
        foundname = names[namesstart:namesend + 4]
        if len(foundname) > 5:
            nameslist.append(foundname)
            names = names[namesend + 4:]
    return nameslist


def unpack_archieved_files(path, ext):
    """
    Рекурсивная распаковка архивов

    :param ext: расширение файла
    :param path: путь к архиву
    :return: список путей к распакованным файлам, список названий файлов
    """
    tempdir = tempfile.mkdtemp()
    total_files = []
    total_names = []

    if ext == '.zip':
        zf = zipfile.ZipFile(path)
        try:
            zf.testzip()
        except:
            return [path], [os.path.basename(path) + '.lockedArchive']
        zf.extractall(tempdir)

    if ext == '.rar':
        rf = rarfile.RarFile(path)
        try:
            rf.testrar()
        except:
            return [path], [os.path.basename(path) + '.lockedArchive']
        rf.extractall(tempdir)
    if ext == '.7z':
        z7f = py7zr.SevenZipFile(path)
        if z7f.password_protected:
            return [path], [os.path.basename(path) + '.lockedArchive']
        z7f.extractall(tempdir)

    # patoolib.extract_archive(path, outdir=tempdir)
    extracted_files = glob.glob(tempdir + '/**/*', recursive=True)
    total_files.extend(extracted_files)
    for ex_file in extracted_files:
        if ex_file.lower().endswith(('.zip', '.7z', '.rar')):
            files, names = unpack_archieved_files(ex_file, '.' + ex_file.rsplit('.')[1])
            total_files.extend(files)
    total_names = [os.path.basename(i) for i in total_files]
    return total_files, total_names


def fitPdfInA4(pdfpath):
    """
    Автоматический поворот в вертикальную ориентацию и вписывание документа в А4
    :param pdfpath: путь к файлу
    :return: outpath - путь к сформированному temp-файлу
    """
    try:
        writer = PdfFileWriter()
        pdf = PdfFileReader(pdfpath, strict=False)
    except Exception as e:
        print(e)
        print('cant open given path (fitPdfInA4 error)')
        return pdfpath
    for page in pdf.pages:
        writer.addPage(page)
    fd, tempoutpath = tempfile.mkstemp('.pdf')
    os.close(fd)
    with open(tempoutpath, "wb") as fp:
        writer.write(fp)
    pdf = PdfFileReader(tempoutpath)
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
            page = PyPDF2.pdf.PageObject.createBlankPage(width=a4orig[0], height=a4orig[1])
            padx = oldpage.mediaBox.getWidth() / 2
            pady = oldpage.mediaBox.getHeight() / 2
            page.mergeTranslatedPage(oldpage, 306 - padx, 421 - pady)
        new_pdf.addPage(page)
    fd, outpath = tempfile.mkstemp('.pdf')
    os.close(fd)
    with open(outpath, mode='wb') as export:
        new_pdf.write(export)
    return outpath


def twoUP(filepath):
    """
    Реализация функции печати n страниц на 1 листе бумаги.
    :param filepath: Путь к файлу
    :return: путь к сформированному temp-файлу
    """
    try:
        rotated_pdf = fitPdfInA4(filepath)
    except:
        rotated_pdf = filepath
        print('Fitting error with: ', filepath)
        pass
    merged_file = PdfFileWriter()
    orig_file = PdfFileReader(rotated_pdf, strict=False)
    n_pages = len(orig_file.pages)
    for i in range(0, n_pages, 2):
        big_page = PyPDF2.pdf.PageObject.createBlankPage(width=595.2, height=842.88)
        big_page.mergeRotatedScaledTranslatedPage(orig_file.pages[i], rotation=90, scale=0.7, tx=585.2, ty=10)
        try:
            big_page.mergeRotatedScaledTranslatedPage(orig_file.pages[i + 1], rotation=90, scale=0.7, tx=585.2, ty=420)
        except:
            pass
        merged_file.addPage(big_page)
    fd, outpath = tempfile.mkstemp('.pdf')
    os.close(fd)
    with open(outpath, 'wb') as out:
        merged_file.write(out)
    return outpath


def fourUP(filepath):
    rotated_pdf = fitPdfInA4(filepath)
    orig_file = PdfFileReader(rotated_pdf, strict=False)
    merged_file = PdfFileWriter()
    n_pages = len(orig_file.pages)
    for i in range(0, n_pages, 4):
        big_page = PyPDF2.pdf.PageObject.createBlankPage(width=595.2, height=842.88)
        big_page.mergeScaledTranslatedPage(orig_file.pages[i], scale=0.48, tx=10, ty=411.44)
        try:
            big_page.mergeScaledTranslatedPage(orig_file.pages[i + 1], scale=0.48, tx=288, ty=411.44)
        except:
            pass

        try:
            big_page.mergeScaledTranslatedPage(orig_file.pages[i + 2], scale=0.48, tx=10, ty=10)
        except:
            pass
        try:
            big_page.mergeScaledTranslatedPage(orig_file.pages[i + 3], scale=0.48, tx=283, ty=10)
        except:
            pass
        merged_file.addPage(big_page)
    fd, outpath = tempfile.mkstemp('.pdf')
    os.close(fd)
    with open(outpath, 'wb') as out:
        merged_file.write(out)
    return outpath


def office2pdf(origfile):
    """
    Конвертация из word в pdf с помощью API офиса, создает файл в той же директории
    :param origfile: путь к файлу
    :return: путь к сконвертированному файлу
    """
    ext = '.' + origfile.rsplit('.', 1)[1].lower()
    convfile = f'{origfile}.pdf'
    wordext = ['.odt', '.rtf', '.doc', '.docx']
    excelext = ['.ods', '.xls', '.xlsx']
    try:
        if ext in wordext:
            word = win32com.client.Dispatch('Word.Application')
            doc = word.Documents.Open(origfile)
            doc.SaveAs(convfile, FileFormat=17)
            doc.Close()
            doc = None
            word.Quit()
            word = None
        if ext in excelext:
            excel = win32com.client.Dispatch("Excel.Application")
            book = excel.Workbooks.Open(Filename=origfile)
            book.ExportAsFixedFormat(0, convfile)
            book = None
            excel.Quit()
            excel = None
    except Exception as e:
        print(e)
        print('cant convert office file ', origfile)
        return origfile
    try:
        os.remove(origfile)
    except:
        print('cant remove ', origfile)
        pass
    neworigfile = f'{origfile.rsplit(".", 1)[0]}.pdf'
    try:
        os.rename(convfile, neworigfile)
    except:
        neworigfile = f'{origfile.rsplit(".", 1)[0]}..pdf'
        os.rename(convfile, neworigfile)
    return neworigfile


def convertImageWithJava(fp):
    return gateway.entry_point.generatePDFFromImage(fp)
