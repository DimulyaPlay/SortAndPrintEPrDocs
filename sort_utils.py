import glob
import os
import subprocess
import tempfile
import time
from difflib import SequenceMatcher

import PyPDF2
import patoolib
import pdfplumber
import win32api
import win32com.client
import win32print
from PIL import Image
from PyPDF2 import PdfFileReader, PdfFileWriter
from pikepdf import Pdf

a4orig = [612.1, 842.0]  # оригинальный формат А4
a4small = [i * 0.95 for i in a4orig]  # размер для масштабирования под область печати


def concat_pdfs(main_pdf_filepath, slave_pdf_filepath):
	"""
	:param main_pdf_filepath:
	путь к основному документу
	:param slave_pdf_filepath:
	путь к присоединяемому документу
	:return:
	outpath - путь к объединенному файлу
	is_paper_eco - bool - сэкономлен ли лист бумаги при конкатенации для статистики
	"""
	file_main = Pdf.open(main_pdf_filepath)
	is_paper_eco = int(len(file_main.pages) % 2)
	file_slave = Pdf.open(slave_pdf_filepath)
	file_main.pages.extend(file_slave.pages)
	outpath = f"{main_pdf_filepath[:-4]}+protocol.pdf"
	file_main.save(outpath)
	file_main.close()
	file_slave.close()
	return outpath, is_paper_eco


def fitPdfInA4(pdfpath):
	"""
	Автоматический поворот в вертикальную ориентацию и вписывание документа в А4
	:param pdfpath: путь к файлу
	:return: outpath - путь к сформированному temp-файлу
	"""
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
			page = PyPDF2.pdf.PageObject.createBlankPage(width = 612.1, height = 842.0)
			padx = oldpage.mediaBox.getWidth() / 2
			pady = oldpage.mediaBox.getHeight() / 2
			page.mergeTranslatedPage(oldpage, 306 - padx, 421 - pady)
		new_pdf.addPage(page)
	fd, outpath = tempfile.mkstemp('.pdf')
	os.close(fd)
	with open(outpath, mode = 'wb') as export:
		new_pdf.write(export)
	return outpath


def multiplePagesPerSheet(filepath, mode):
	"""
	Реализация функции печати n страниц на 1 листе бумаги.
	:param filepath: Путь к файлу
	:param mode: 1,2 или 4 - варианты размещения
	:return: путь к сформированному temp-файлу
	"""
	if os.path.basename(filepath)[3:].startswith('Kvitantsiya_ob_otpravke['):
		return filepath
	if mode == 1:
		return fitPdfInA4(filepath)
	merged_file = PdfFileWriter()
	if mode == 2:
		rotated_pdf = fitPdfInA4(filepath)
		orig_file = PdfFileReader(rotated_pdf, strict = False)
		n_pages = len(orig_file.pages)
		for i in range(0, n_pages, 2):
			big_page = PyPDF2.pdf.PageObject.createBlankPage(width = 595.2, height = 842.88)
			big_page.mergeRotatedScaledTranslatedPage(orig_file.pages[i], rotation = 90, scale = 0.7, tx = 585.2,
													  ty = 10)
			try:
				big_page.mergeRotatedScaledTranslatedPage(orig_file.pages[i + 1], rotation = 90, scale = 0.7,
														  tx = 585.2, ty = 420)
			except:
				pass
			merged_file.addPage(big_page)
	if mode == 4:
		rotated_pdf = fitPdfInA4(filepath)
		orig_file = PdfFileReader(rotated_pdf, strict = False)
		n_pages = len(orig_file.pages)
		for i in range(0, n_pages, 4):
			big_page = PyPDF2.pdf.PageObject.createBlankPage(width = 595.2, height = 842.88)
			big_page.mergeScaledTranslatedPage(orig_file.pages[i], scale = 0.48, tx = 10, ty = 411.44)
			try:
				big_page.mergeScaledTranslatedPage(orig_file.pages[i + 1], scale = 0.48, tx = 288, ty = 411.44)
			except:
				pass

			try:
				big_page.mergeScaledTranslatedPage(orig_file.pages[i + 2], scale = 0.48, tx = 10, ty = 10)
			except:
				pass
			try:
				big_page.mergeScaledTranslatedPage(orig_file.pages[i + 3], scale = 0.48, tx = 283, ty = 10)
			except:
				pass
			merged_file.addPage(big_page)
	fd, outpath = tempfile.mkstemp('.pdf')
	os.close(fd)
	with open(outpath, 'wb') as out:
		merged_file.write(out)
	return outpath


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
	with pdfplumber.open(path[0]) as pdf:
		all_text = ''
		for i in pdf.pages:
			all_text += i.extract_text()
	return all_text.replace('\n', '')


def check_num_pages(path):
	"""
	Рассчет количества страниц и листов в документе
	:param path: путь к файлу
	:return: лист - страниц, листов
	"""
	pdf = Pdf.open(path)
	pages = len(pdf.pages)
	papers = int(pages / 2 + 0.9)
	pdf.close()
	return [pages, papers]


def wordpdf(origfile):
	"""
	Конвертация из word в pdf с помощью API офиса, создает файл в той же директории
	:param origfile: путь к файлу
	:return: путь к сконвертированному файлу
	"""
	convfile = f'{origfile}.pdf'
	word = win32com.client.Dispatch('Word.Application')
	doc = word.Documents.Open(origfile)
	doc.SaveAs(convfile, FileFormat = 17)
	doc.Close()
	word.Quit()
	os.remove(origfile)
	neworigfile = f'{origfile.rsplit(".", 1)[0]}.pdf'
	try:
		os.rename(convfile, neworigfile)
	except:
		neworigfile = f'{origfile.rsplit(".", 1)[0]}..pdf'
		os.rename(convfile, neworigfile)
	return neworigfile


def imagepdf(origfile):
	"""
	Конвертация из .jpg', '.jpeg', '.png', '.tif в .pdf
	:param origfile: путь к файлу
	:return: путь к файлу
	"""
	convfile = f'{origfile}.pdf'
	if origfile.endswith('.png'):
		image = Image.open(origfile)
		if image.width > image.height:
			print('rotating')
			image.rotate(270)
		image.convert('RGB')
		image.save(convfile)
		os.remove(origfile)
		neworigfile = f'{origfile.rsplit(".", 1)[0]}.pdf'
		os.rename(convfile, neworigfile)
		return neworigfile
	else:
		convfile = f'{origfile}.png'
		image = Image.open(origfile)
		image.convert('RGB')
		image.save(convfile)
		neworigfile = imagepdf(convfile)
	return neworigfile


def print_file(filepath, exe_path, currentprinter):
	"""
	Помещение файла в очередь печати принтера. Документ передается в программу для печати, далее в цикле проверяется статус
	этого документа до тех пор, пока он не пройдет все стадии постановки в очередь, затем программа завершается.
	:param filepath: путь к файлу
	:param exe_path: путь к программе PDFtoprinter
	:param currentprinter: название принтера
	"""
	win32api.ShellExecute(0, 'open', exe_path, '/s ' + '"' + filepath + '"' + ' "' + currentprinter + '" ', '.', 0)
	jobs = [0, 0, 0, 0, 0]
	while sum(jobs) < 3:
		time.sleep(0.01)
		phandle = win32print.OpenPrinter(currentprinter)
		print_jobs = win32print.EnumJobs(phandle, 0, -1, 1, )
		docs_in_queue = {job['pDocument']:job['Status'] for job in print_jobs}
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
	subprocess.call("taskkill.exe /im pdftoprinter.exe /f", shell = True)


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


def unpack_archieved_files(path):
	tempdir = tempfile.mkdtemp()
	total_files = []
	total_names = []
	patoolib.extract_archive(path, outdir = tempdir)
	extracted_files = glob.glob(tempdir + '/**/*', recursive = True)
	total_files.extend(extracted_files)
	for ex_file in extracted_files:
		if ex_file.lower().endswith(('.zip', '7z', 'rar')):
			files, names = unpack_archieved_files(ex_file)
			total_files.extend(files)
	total_names = [os.path.basename(i) for i in total_files]
	return total_files, total_names
