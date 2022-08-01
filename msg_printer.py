import shutil
from tkinter import *

import win32com.client

from scrollable_frame import VerticalScrolledFrame
from sort_utils import *


class Message_handler:
	def __init__(self):
		self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
		self.allowed_ext = ['.doc', '.docx', '.pdf', '.jpg', '.jpeg', '.tif', '.tiff', '.png', '.gif', '.heic']
		self.allowed_ext_img = ['.jpg', '.jpeg', '.tif', '.tiff', '.png', '.gif', '.heic']
		self.allowed_ext_docs = ['.doc', '.docx']
		self.allowed_ext_archives = ['.7z', '.rar', '.zip']

	@staticmethod
	def get_file_from_attach(att):
		"""
		Коневртирует аттачмент, считает количество листов, позвращает аттачмент ссылку и статут возможности печати
		:param att: outlook attachment object
		:return: путь к тепмфайлу, готовность к печати файла, количество страниц, количество листов
		"""
		ext = '.' + att.FileName.rsplit('.', 1)[1].lower()
		fd, outpath = tempfile.mkstemp(ext)
		os.close(fd)
		att.SaveAsFile(outpath)
		return outpath, ext

	def get_files_from_msg(self, msg):
		list_files = []
		for att in msg.attachments:
			filename = att.FileName
			outpath, ext = self.get_file_from_attach(att)
			if ext in self.allowed_ext_archives:
				filepaths, filenames = unpack_archieved_files(outpath)
				for fp, fn in zip(filepaths, filenames):
					list_files.append([fp, fn])
			if ext in self.allowed_ext:
				list_files.append([outpath, filename])
		return list_files

	def convert_attachments(self, list_fp_fn):
		new_list_fp_fn_pgs_pps = []
		for fp, fn in list_fp_fn:
			ext = '.' + fn.rsplit('.', 1)[1].lower()
			if ext in self.allowed_ext:
				if ext in self.allowed_ext_img:
					new_fp = imagepdf(fp)
					num_pgs, num_pps = check_num_pages(new_fp)
					new_list_fp_fn_pgs_pps.append([new_fp, fn, num_pgs, num_pps, 1])
				if ext in self.allowed_ext_docs:
					new_fp = wordpdf(fp)
					num_pgs, num_pps = check_num_pages(new_fp)
					new_list_fp_fn_pgs_pps.append([new_fp, fn, num_pgs, num_pps, 1])
				if ext == '.pdf':
					num_pgs, num_pps = check_num_pages(fp)
					new_list_fp_fn_pgs_pps.append([fp, fn, num_pgs, num_pps, 1])
			else:
				new_list_fp_fn_pgs_pps.append([fp, fn, '?', '?', 0])
		return new_list_fp_fn_pgs_pps

	def handle_messages(self, msgs):
		self.handle_keys = []
		self.handled_messages = {}
		self.handled_attachments = {}
		for i in msgs:
			fd, outpath = tempfile.mkstemp('.msg')
			os.close(fd)
			shutil.copy2(i, outpath)
			msg = self.outlook.OpenSharedItem(fr'{outpath}')
			attachment_files = self.get_files_from_msg(msg)
			attachment_files = self.convert_attachments(attachment_files)
			self.handle_keys.append(outpath)
			self.handled_messages[outpath] = msg
			self.handled_attachments[outpath] = attachment_files

	def print_dialog_msg(self, root):
		dialog = Toplevel(root)
		dialog.title(f'Файлов на печать ')
		dialog.attributes('-topmost', True)
		dialog.resizable(False, False)

		def update_num_pages():
			string_pages_papers = f"Всего для печати страниц: , листов: "
			dialog.title(f'Документов на печать ')

		if sum([len(item) for item in self.handled_attachments.items()]) + len(self.handled_messages.keys()) > 20:
			height = 550
		else:
			height = (sum([len(item) for item in self.handled_attachments.items()]) + len(
				self.handled_messages.keys()) + 1) * 25

		container = VerticalScrolledFrame(dialog, height = height, width = 731)
		container.pack()
		rbVariables = {}
		printcbVariables = {}
		Label(container, text = 'Название документа/тема').grid(column = 1, row = 0)
		Label(container, text = 'Страниц').grid(column = 2, row = 0)
		Label(container, text = '1').grid(column = 3, row = 0)
		Label(container, text = '2').grid(column = 4, row = 0)
		Label(container, text = '4').grid(column = 5, row = 0)
		currentrow = 1
		for n, filepath in enumerate(self.handle_keys):
			subject = self.handled_messages[filepath].subject if self.handled_messages[
																	 filepath].subject != '' else 'Пустая тема'
			subject = subject if len(subject) < 48 else subject[:45] + "..."
			printcbVariables[filepath] = BooleanVar()
			printcbVariables[filepath].set(1)
			prntchb = Checkbutton(container, variable = printcbVariables[filepath], command = update_num_pages)
			prntchb.var = printcbVariables[filepath]
			prntchb.grid(column = 0, row = currentrow, sticky = W)
			Label(container, text = subject, font = 'TkFixedFont 10 bold').grid(column = 1, row = currentrow,
																				sticky = W)
			currentrow += 1
			for j, att in enumerate(self.handled_attachments[filepath]):
				current_key = att[0]
				current_name = att[1]
				current_pages = att[2]
				current_name = current_name if len(current_name) < 48 else current_name[:45] + "..."
				padx = 10
				printcbVariables[current_key] = BooleanVar()
				printcbVariables[current_key].set(1)
				prntchb = Checkbutton(container, variable = printcbVariables[current_key], command = update_num_pages)
				prntchb.var = printcbVariables[current_key]
				prntchb.grid(column = 0, row = currentrow, sticky = W)
				lb = Label(container, text = current_name, font = 'TkFixedFont')
				lb.grid(column = 1, row = currentrow, sticky = W, padx = padx)
				lb.bind('<Double-Button-1>', lambda event, a = current_key:os.startfile(a))
				lb1 = Label(container, text = current_pages)
				lb1.grid(column = 2, row = currentrow, sticky = W, padx = padx)
				rbVariables[current_key] = IntVar()
				rbVariables[current_key].set(1)
				rb1 = Radiobutton(container, variable = rbVariables[current_key], value = 1, command = update_num_pages)
				rb1.grid(column = 3, row = currentrow, sticky = W)
				rb2 = Radiobutton(container, variable = rbVariables[current_key], value = 2, command = update_num_pages)
				rb2.grid(column = 4, row = currentrow, sticky = W)
				rb4 = Radiobutton(container, variable = rbVariables[current_key], value = 4, command = update_num_pages)
				rb4.grid(column = 5, row = currentrow, sticky = W)
				prvbtn = Label(container, text = 'Предпросмотр', padx = 2)
				prvbtn.grid(column = 6, row = currentrow)
				prvbtn.bind('<Button-1>',
							lambda event, a = current_key:os.startfile(multiplePagesPerSheet(a, rbVariables[a].get())))
				currentrow += 1
