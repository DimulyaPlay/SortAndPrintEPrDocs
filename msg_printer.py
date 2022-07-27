import shutil
from tkinter import *

import win32com.client

from scrollable_frame import VerticalScrolledFrame
from sort_utils import *


class Message_handler:
	def __init__(self):
		self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
		self.allowed_ext = ['.doc', '.docx', '.pdf', '.jpg', '.jpeg', '.tif', '.tiff', '.png', '.gif']
		self.allowed_ext_img = ['.jpg', '.jpeg', '.tif', '.tiff', '.png', '.gif']
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
			else:
				if ext in self.allowed_ext:
					list_files.append([outpath, filename])

		return list_files

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

		# if len(self.handled_attachments.keys()) + len(self.handled_messages.keys()) > 20:
		height = 550
		# else:
		# 	height = (len(*self.handled_attachments.items()) + len(self.handled_messages.keys()) + 1) * 25

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
			subject = subject if len(subject) < 58 else subject[:55] + "..."
			printcbVariables[filepath] = BooleanVar()
			printcbVariables[filepath].set(1)
			prntchb = Checkbutton(container, variable = printcbVariables[filepath], command = update_num_pages)
			prntchb.var = printcbVariables[filepath]
			prntchb.grid(column = 0, row = currentrow, sticky = W)
			Label(container, text = subject, font = 'TkFixedFont 10 bold').grid(column = 1, row = currentrow,
																				sticky = W)
			currentrow += 1
			for j, att in enumerate(self.handled_attachments[filepath]):
				padx = 10
				printcbVariables[att[0]] = BooleanVar()
				printcbVariables[att[0]].set(1)
				prntchb = Checkbutton(container, variable = printcbVariables[att[0]], command = update_num_pages)
				prntchb.var = printcbVariables[att[0]]
				prntchb.grid(column = 0, row = currentrow, sticky = W)
				lb = Label(container, text = att[1], font = 'TkFixedFont')
				lb.grid(column = 1, row = currentrow, sticky = W, padx = padx)
				lb.bind('<Double-Button-1>', lambda event, a = att[0]:os.startfile(a))
				lb1 = Label(container, text = '-')
				lb1.grid(column = 2, row = currentrow, sticky = W, padx = padx)
				rbVariables[att[0]] = IntVar()
				rbVariables[att[0]].set(1)
				rb1 = Radiobutton(container, variable = rbVariables[att[0]], value = 1, command = update_num_pages)
				rb1.grid(column = 3, row = currentrow, sticky = W)
				rb2 = Radiobutton(container, variable = rbVariables[att[0]], value = 2, command = update_num_pages)
				rb2.grid(column = 4, row = currentrow, sticky = W)
				rb4 = Radiobutton(container, variable = rbVariables[att[0]], value = 4, command = update_num_pages)
				rb4.grid(column = 5, row = currentrow, sticky = W)
				prvbtn = Label(container, text = 'Предпросмотр', padx = 2)
				prvbtn.grid(column = 6, row = currentrow)
				prvbtn.bind('<Button-1>',
							lambda event, a = att[0]:os.startfile(multiplePagesPerSheet(a, rbVariables[att[0]].get())))
				currentrow += 1
