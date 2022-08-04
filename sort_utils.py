import glob
import tempfile
from difflib import SequenceMatcher
import patoolib
from PDFNetPython3 import *


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
	# doc.InitSecurityHandler()
	n_pages = doc.GetPageCount()
	pages = n_pages
	papers = int(pages / 2 + 0.9)
	doc.Close()
	return [pages, papers]


def concat_pdfs(master, wingman):
	doc1 = PDFDoc(master)
	doc2 = PDFDoc(wingman)
	page_num = doc1.GetPageCount()
	doc1.InitSecurityHandler()
	doc2.InitSecurityHandler()
	doc1.InsertPages(page_num + 1, doc2, 1, doc2.GetPageCount(), PDFDoc.e_none)
	doc1.Save(master, SDFDoc.e_remove_unused)
	doc1.Close()
	doc2.Close()
	is_paper_eco = page_num % 2
	return master, is_paper_eco


def print_file(filepath, mode, currentprinter):
	doc = PDFDoc(filepath)
	doc.InitSecurityHandler()
	printerMode = PrinterMode()
	printerMode.SetAutoCenter(True)
	printerMode.SetAutoRotate(True)
	printerMode.SetScaleType(PrinterMode.e_ScaleType_ReduceToOutputPage)
	printerMode.SetNUp(1, 1)
	if mode == 2:
		printerMode.SetNUp(1, 2, PrinterMode.e_PageOrder_BottomToTopThenLeftToRight)
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
