import shutil
from tkinter import *
import win32com.client
from scrollable_frame import VerticalScrolledFrame
from sort_utils import *
from tkinterhtml import HtmlFrame


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
			list_files.append([outpath, filename])
		return list_files

	def convert_attachments(self, list_fp_fn):
		new_list_fp_fn_pgs_pps = []
		for fp, fn in list_fp_fn:
			ext = '.' + fn.rsplit('.', 1)[1].lower()
			if ext in self.allowed_ext:
				if ext != '.pdf':
					pdfdoc = PDFDoc()
					Convert.ToPdf(pdfdoc, fp)
					pdfdoc.Save(fp + '.pdf', SDFDoc.e_compatibility)
					pdfdoc.Close()
					num_pgs, num_pps = check_num_pages(fp + '.pdf')
					new_list_fp_fn_pgs_pps.append([fp + '.pdf', fn, num_pgs, num_pps, 1])
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

	def print_dialog_msg(self, root, current_config):
		dialog = Toplevel(root)
		dialog.title(f'Файлов на печать ')
		dialog.attributes('-topmost', True)
		dialog.resizable(False, False)

		def view_body(msg):
			viewer = Toplevel(root)
			frame = HtmlFrame(viewer, horizontal_scrollbar = "auto")
			frame.set_content(msg.HTMLBody)
			frame.pack()

		def update_num_pages():
			total_pages = 0
			total_papers = 0
			docs_for_print = 0
			for msg in self.handle_keys:
				if printcbVariables[msg].get():
					docs_for_print += 1
					total_pages += 1
					total_papers += 1
				for att in self.handled_attachments[msg]:
					try:
						is_checked = printcbVariables[att[0]].get()
					except:
						is_checked = 0
					if is_checked:
						docs_for_print += 1
						total_pages += att[2]
						total_papers += att[3]
			string_pages_papers = f"Всего для печати страниц: {total_pages}, листов: {total_papers}"
			len_pages.set(string_pages_papers)
			dialog.title(f'Документов на печать {docs_for_print}')

		def apply_print(e):
			print_button.unbind("<Button-1>")
			print_button.config(relief = SUNKEN)
			print_button.update()
			for msg in self.handle_keys:
				if printcbVariables[msg].get():
					printcbVariables[msg].set(0)
					self.handled_messages[msg].PrintOut()
				for att in self.handled_attachments[msg]:
					if printcbVariables[att[0]].get():
						printcbVariables[att[0]].set(0)
						print_file(att[0], rbVariables[att[0]].get(), current_config.default_printer)
			info_show_printed()
			print_button.config(relief = RAISED)
			print_button.bind("<Button-1>", apply_print)

		def check_all_chbtns():
			if prntchballvar.get():
				for chbtn in printcbVariables.values():
					chbtn.set(1)
			else:
				for chbtn in printcbVariables.values():
					chbtn.set(0)
			update_num_pages()

		MAXHEIGHT = 650
		height = 1
		for msg in self.handle_keys:
			height += 1
			for att in self.handled_attachments[msg]:
				height += 1
		height = height * 25
		if height > MAXHEIGHT:
			height = MAXHEIGHT

		container = VerticalScrolledFrame(dialog, height = height, width = 620)
		container.pack()
		rbVariables = {}
		printcbVariables = {}
		prntchballvar = BooleanVar()
		prntchballvar.set(1)
		prntchball = Checkbutton(container, variable = prntchballvar, command = check_all_chbtns)
		prntchball.grid(column = 0, row = 0)
		Label(container, text = 'Название документа/тема').grid(column = 1, row = 0)
		Label(container, text = 'Страниц').grid(column = 2, row = 0)
		Label(container, text = '1').grid(column = 3, row = 0)
		Label(container, text = '2').grid(column = 4, row = 0)
		Label(container, text = '4').grid(column = 5, row = 0)
		currentrow = 1

		for filepath in self.handle_keys:
			subject = self.handled_messages[filepath].subject if self.handled_messages[
																	 filepath].subject != '' else 'Пустая тема'
			subject = subject if len(subject) < 68 else subject[:65] + "..."
			printcbVariables[filepath] = BooleanVar()
			printcbVariables[filepath].set(1)
			prntchb = Checkbutton(container, variable = printcbVariables[filepath], command = update_num_pages)
			prntchb.var = printcbVariables[filepath]
			prntchb.grid(column = 0, row = currentrow, sticky = W)
			msglb = Label(container, text = subject, font = 'TkFixedFont 9 bold')
			msglb.grid(column = 1, row = currentrow, sticky = W)
			msglb.bind('<Double-Button-1>', lambda event, a = self.handled_messages[filepath]:view_body(a))
			currentrow += 1
			for j, att in enumerate(self.handled_attachments[filepath]):
				current_key = att[0]
				current_name = att[1]
				current_pages = att[2]
				current_papers = att[3]
				current_printable = att[4]
				current_name = current_name if len(current_name) < 58 else current_name[:55] + "..."
				padx = 10
				if current_printable:
					printcbVariables[current_key] = BooleanVar()
					printcbVariables[current_key].set(1)
					prntchb = Checkbutton(container, variable = printcbVariables[current_key],
										  command = update_num_pages)
					prntchb.var = printcbVariables[current_key]
				else:
					prntchb = Checkbutton(container, state = DISABLED)
				prntchb.grid(column = 0, row = currentrow, sticky = W, padx = padx / 2)
				lb = Label(container, text = current_name, font = 'TkFixedFont')
				lb.grid(column = 1, row = currentrow, sticky = W, padx = padx)
				lb.bind('<Double-Button-1>', lambda event, a = current_key:os.startfile(a))
				lb1 = Label(container, text = current_pages)
				lb1.grid(column = 2, row = currentrow, sticky = W, padx = padx)
				rbVariables[current_key] = IntVar()
				rbVariables[current_key].set(1)
				rb1 = Radiobutton(container, variable = rbVariables[current_key], value = 1, command = update_num_pages)
				rb1.var = rbVariables[current_key]
				rb1.grid(column = 3, row = currentrow, sticky = W)
				rb2 = Radiobutton(container, variable = rbVariables[current_key], value = 2, command = update_num_pages)
				rb2.var = rbVariables[current_key]
				rb2.grid(column = 4, row = currentrow, sticky = W)
				rb4 = Radiobutton(container, variable = rbVariables[current_key], value = 4, command = update_num_pages)
				rb4.var = rbVariables[current_key]
				rb4.grid(column = 5, row = currentrow, sticky = W)
				currentrow += 1
		bottom_actions = Frame(dialog)
		bottom_actions.pack()
		len_pages = StringVar()
		print_button = Label(bottom_actions, text = " Печать ", borderwidth = 2, relief = RAISED)
		print_button.grid(column = 2, row = 0, sticky = S, padx = 5, pady = 2)
		print_button.bind("<Button-1>", apply_print)
		sum_pages = Label(bottom_actions, textvariable = len_pages)
		sum_pages.grid(column = 3, row = 0, sticky = S, padx = 5, pady = 2)
		update_num_pages()


def info_show_printed():
	messagebox.showinfo("Готово", "Документы отправлены в очередь принтера.")
