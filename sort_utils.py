import glob
import os
import tempfile
from difflib import SequenceMatcher
import patoolib
from PDFNetPython3 import *
from PyPDF2 import PdfFileReader, PdfFileWriter
import PyPDF2
# import fitz
import win32com

a4orig = [612.1, 842.0]  # оригинальный формат А4
a4small = [i * 0.99 for i in a4orig]  # размер для масштабирования под область печати


def initPDFTron(lc):
	PDFNet.Initialize(lc)


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
	doc = PDFDoc(path[0])
	# doc.InitSecurityHandler()
	n_pages = doc.GetPageCount()
	txt = TextExtractor()
	all_text = ''
	for i in range(n_pages):
		page = doc.GetPage(i + 1)
		txt.Begin(page)
		text = txt.GetAsText()
		all_text += text
	doc.Close()
	return all_text.replace('\n', '')


def check_num_pages(path):
	"""
	Рассчет количества страниц и листов в документе

	:param path: путь к файлу
	:return: лист - страниц, листов
	"""
	doc = PDFDoc(path)
	n_pages = doc.GetPageCount()
	pages = n_pages
	papers = int(pages / 2 + 0.9)
	doc.Close()
	return [pages, papers]


def concat_pdfs(master, wingman):
	"""
	Конкатенация пдф файлов
	:param master: путь к первому пдф

	:param wingman: путь ко второму пдф

	:return: str путь к обьединенному пдф, bool сохранен ли лист
	"""
	doc1 = PDFDoc(master)
	doc2 = PDFDoc(wingman)
	page_num = doc1.GetPageCount()
	doc1.InsertPages(page_num + 1, doc2, 1, doc2.GetPageCount(), PDFDoc.e_none)
	doc1.Save(master, SDFDoc.e_remove_unused)
	doc1.Close()
	doc2.Close()
	is_paper_eco = page_num % 2
	return master, is_paper_eco


def print_file(filepath, mode, currentprinter, convert = False):
	"""
	Отправка документа в очередь печати с заданными параметрами

	:param filepath: путь к файлу
	:param mode: расположение страниц на листе

	:param currentprinter: название принтера

	:param convert: принудительная конвертация
	"""
	doc = PDFDoc(filepath)
	printerMode = PrinterMode()
	printerMode.SetAutoCenter(True)
	printerMode.SetAutoRotate(True)
	printerMode.SetScaleType(PrinterMode.e_ScaleType_ReduceToOutputPage)
	printerMode.SetNUp(1, 1)
	if convert:
		Convert.ToTiff(doc, filepath + '.tiff')
	if mode == 2:
		doc.Close()
		doc = PDFDoc(twoUP(filepath))
		if not convert:
			Convert.ToTiff(doc, filepath + '.tiff')
	if mode == 4:
		printerMode.SetNUp(2, 2, PrinterMode.e_PageOrder_LeftToRightThenTopToBottom)
	Print.StartPrintJob(doc, currentprinter, doc.GetFileName(), "", None, printerMode, None)
	doc.Close()


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
	"""
	Рекурсивная распаковка архивов

	:param path: путь к архиву

	:return: список путей к распакованным файлам, список названий файлов
	"""
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
			page = PyPDF2.pdf.PageObject.createBlankPage(width = a4orig[0], height = a4orig[1])
			padx = oldpage.mediaBox.getWidth() / 2
			pady = oldpage.mediaBox.getHeight() / 2
			page.mergeTranslatedPage(oldpage, 306 - padx, 421 - pady)
		new_pdf.addPage(page)
	fd, outpath = tempfile.mkstemp('.pdf')
	os.close(fd)
	with open(outpath, mode = 'wb') as export:
		new_pdf.write(export)
	return outpath


def twoUP(filepath):
	"""
	Реализация функции печати n страниц на 1 листе бумаги.
	:param filepath: Путь к файлу
	:return: путь к сформированному temp-файлу
	"""
	merged_file = PdfFileWriter()
	rotated_pdf = fitPdfInA4(filepath)
	orig_file = PdfFileReader(rotated_pdf, strict = False)
	n_pages = len(orig_file.pages)
	for i in range(0, n_pages, 2):
		big_page = PyPDF2.pdf.PageObject.createBlankPage(width = 595.2, height = 842.88)
		big_page.mergeRotatedScaledTranslatedPage(orig_file.pages[i], rotation = 90, scale = 0.7, tx = 585.2, ty = 10)
		try:
			big_page.mergeRotatedScaledTranslatedPage(orig_file.pages[i + 1], rotation = 90, scale = 0.7, tx = 585.2,
													  ty = 420)
		except:
			pass
		merged_file.addPage(big_page)
	fd, outpath = tempfile.mkstemp('.pdf')
	os.close(fd)
	with open(outpath, 'wb') as out:
		merged_file.write(out)
	return outpath


def word2pdf(origfile):
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

# def pdf2images(origfile):
# 	list_files = []
# 	zoom_x = 2.0  # horizontal zoom
# 	zoom_y = 2.0  # vertical zoom
# 	mat = fitz.Matrix(zoom_x, zoom_y)  # zoom factor 2 in each dimension
# 	doc = fitz.open(origfile)  # open document
# 	for page in doc:  # iterate through the pages
# 		pix = page.get_pixmap(matrix = mat)  # render page to an image
# 		outpath = origfile + f'_{page.number:02}.png'
# 		pix.save(outpath)  # store image as a PNG
# 		list_files.append(outpath)
# 	return list_files


# def images2pdf(list_files):
# 	doc = fitz.open()  # PDF with the pictures
# 	for i in list_files:
# 		img = fitz.open(i)  # open pic as document
# 		rect = img[0].rect  # pic dimension
# 		pdfbytes = img.convert_to_pdf()  # make a PDF stream
# 		img.close()  # no longer needed
# 		imgPDF = fitz.open("pdf", pdfbytes)  # open stream as PDF
# 		page = doc.new_page(width = rect.width,  # new page with ...
# 							height = rect.height)  # pic dimension
# 		page.show_pdf_page(rect, imgPDF, 0)  # image fills the page
# 	doc.save(list_files[0] + '.pdf')
# 	return list_files[0] + '.pdf'
