import glob
import random
import subprocess
import time
from tkinter import messagebox
from zipfile import ZipFile
from sort_utils import *


class main_sorter:
    def __init__(self, config, config_path):
        self.config_path = config_path
        self.config_obj = config
        self.read_vars_from_config()

    def read_vars_from_config(self):
        # получить переменные из конфига
        self.deletezip = self.config_obj.get('DEFAULT', 'delete_zip')
        self.paperecomode = self.config_obj.get('DEFAULT', 'paper_eco_mode')
        self.print_directly = self.config_obj.get('DEFAULT', 'print_directly')
        self.default_printer = self.config_obj.get('DEFAULT', 'default_printer')
        self.PDF_PRINT_FILE = self.config_obj.get('DEFAULT', 'PDF_PRINT_PATH')

    def agregate_file(self, givenpath):
        self.num_pages = {}
        givenpath = givenpath.replace('/', '\\')
        foldername = givenpath[:-4]
        if os.path.exists(foldername):
            foldername = foldername + str(random.randint(1, 999))
        with ZipFile(givenpath, 'r') as zipObj:
            zipObj.extractall(foldername)
        if self.deletezip == 'yes':
            os.remove(givenpath)
        siglist = glob.glob("{0}{1}*.sig".format(foldername, os.sep))
        for i in siglist:
            os.remove(i)
        abspathlist = glob.glob(foldername + os.sep + "*")
        basedoclist = []
        for i in abspathlist:
            if not os.path.basename(i).startswith('Kvitantsiya_ob_otpravke[') and not os.path.basename(i).startswith(
                    'Protokol_proverki_fayla_'):
                basedoclist.append(i)
        doclist = [wordpdf(i) if i.endswith(('.doc', '.docx')) else i for i in basedoclist]
        protlist = [i for i in abspathlist if os.path.basename(i).startswith('Protokol_proverki_fayla_')]
        kvitanciya = [i for i in abspathlist if os.path.basename(i).startswith('Kvitantsiya_ob_otpravke')]

        if not kvitanciya:
            no_kvitancii()
            return
        doc_list = extracttext(kvitanciya)
        queue = {}
        queue[int('-2')] = os.path.basename(kvitanciya[0])
        for i in doclist:
            filename = os.path.basename(i)
            # print(filename)
            file_id = doc_list.find(filename[:-4])
            if file_id != -1:
                queue[file_id] = filename
            prots_similarity = {}
            if not protlist:
                continue
            for protpath in protlist:
                protname = os.path.basename(protpath)
                if protname.find(filename.rsplit('_na_', 1)[0][:76]) != -1:
                    similarity = similar(filename.rsplit('_na_', 1)[0][:76], protname[24:-4])
                    prots_similarity[protpath] = similarity
            maxsimilarity = max(zip(prots_similarity.values(), prots_similarity.keys()))[1]
            queue[file_id + 1] = os.path.basename(maxsimilarity)
            protlist.remove(maxsimilarity)

        queue_files = []
        queue_num_files = []
        counter = 0
        all_keys = sorted(queue.keys())
        for i in sorted(queue.keys()):
            if self.paperecomode == "no":
                queue_files.append('{0}\\{1}'.format(foldername, queue[i]))
                queue_num_files.append(foldername + '\\' + f'{counter:02}_' + queue[i])
                if os.path.exists('{0}\\{1}'.format(foldername, queue[i])):
                    counter += 1
            else:
                if i + 1 in all_keys:
                    if queue[i + 1].startswith('Protokol_proverki_fayla_'):
                        # print("Протокол есть для файла: ", queue[i])
                        merged_file = concat_pdfs('{0}\\{1}'.format(foldername, queue[i]),
                                                          '{0}\\{1}'.format(foldername, queue[i + 1]),
                                                          self.print_directly)
                        broken = 0
                        if not broken:
                            os.remove('{0}\\{1}'.format(foldername, queue[i]))
                            os.remove('{0}\\{1}'.format(foldername, queue[i + 1]))
                        queue_files.append(merged_file)
                        queue_num_files.append(foldername + '\\' + f'{counter:02}_' + queue[i])
                        # if broken:
                        #     queue_files.append('{0}\\{1}'.format(foldername, queue[i + 1]))
                        #     counter += 1
                        #     queue_num_files.append(foldername + '\\' + f'{counter:02}_' + queue[i + 1])
                        counter += 1
                else:
                    if not queue[i].startswith('Protokol_proverki_fayla_'):
                        # print("Протокола нет для файла: ", queue[i])
                        queue_files.append('{0}\\{1}'.format(foldername, queue[i]))
                        queue_num_files.append(foldername + '\\' + f'{counter:02}_' + queue[i])
                        counter += 1
        self.files_for_print = []

        for i, j in zip(queue_files, queue_num_files):
            if os.path.exists(i):
                os.replace(i, j)
                self.num_pages[j] = check_num_pages(j)
                self.files_for_print.append(j)
        if self.print_directly == 'no':
            subprocess.Popen(f'explorer {foldername}')


def no_kvitancii():
    messagebox.showinfo("Варнинг", "В архиве не обнаружена квитанция.")
