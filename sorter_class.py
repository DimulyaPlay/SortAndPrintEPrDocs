import random
from datetime import datetime
from zipfile import ZipFile
from sort_utils import *


class main_sorter:
    def __init__(self, config, stat=False):
        """
        :argument
        config - экземпляр объекта корфигурации

        Создает атрибуты касса из объекта конфигурации
        Функции из других модулей используют эти атрибуты
        """
        self.config = config
        self.stat = stat if stat else None

    def agregate_file(self, givenpath):
        """
        :argument
        givenpath - путь к архиву

        Выполняет распаковку документов, конвертацию из .doc, .docx, .jpg', '.jpeg', '.png', '.tif в .pdf, сортировку
        относительно положения документов в квитанции, сохраняет в словарь внутри экземпляра.

        Производит действия согласно заданой конфигурации:
        deletezip - удаление Zip архива после распаковки при включенной опции
        paperecomode - объединение документа и протокола проверки в один файл при включенной опции
        print_directly - открытие папку с отсортированными файлами, если опция выключена
        save_stat - формирует первоначальную статистику по документам, сохраняет в лист внутри экземпляра
        """
        self.num_pages = {}
        givenpath = givenpath.replace('/', '\\')
        foldername = givenpath[:-4]
        if os.path.exists(foldername):
            foldername = foldername + str(random.randint(1, 999))
        with ZipFile(givenpath, 'r') as zipObj:
            zipObj.extractall(foldername)
        if self.config.deletezip == 'yes':
            os.remove(givenpath)
        siglist = glob.glob(foldername + os.sep + "*.sig")
        [os.remove(i) for i in siglist]
        abspathlist = glob.glob(foldername + os.sep + "*")
        basedoclist = []
        num_appeal = os.path.basename(givenpath).split('_all_files')[0]
        for i in abspathlist:
            if not os.path.basename(i).startswith(f'Kvitantsiya_ob_otpravke[{num_appeal}]') and not os.path.basename(
                    i).startswith('Protokol_proverki_fayla_'):
                basedoclist.append(i)
        doclist = [office2pdf(i) if i.endswith(('.doc', '.docx', '.rtf', '.odt', '.ods', '.xls', '.xlsx')) else i for i
                   in basedoclist]
        doclist = [convertImageWithJava(i) if i.endswith(('.jpg', '.jpeg', '.png', '.tif')) else i for i in doclist]
        protlist = [i for i in abspathlist if os.path.basename(i).startswith('Protokol_proverki_fayla_')]
        if self.config.no_protocols == 'yes':
            protlist = []
        kvitanciya = [i for i in abspathlist if
                      os.path.basename(i).startswith(f'Kvitantsiya_ob_otpravke[{num_appeal}]')]
        if not kvitanciya:
            return
        doc_list = extracttext(kvitanciya[0])
        queue = {}
        queue[int('-2')] = os.path.basename(kvitanciya[0])
        for i in doclist:
            filename = os.path.basename(i)
            file_id = doc_list.find(filename[:-4])
            if file_id != -1:
                queue[file_id] = filename
            prots_similarity = {}
            if not protlist:
                continue
            for protpath in protlist:
                protname = os.path.basename(protpath)
                if protname.find(filename.rsplit('_na_', 1)[0][:76]) != -1 or protname.find(
                        filename.rsplit('[', 1)[0][:76]) != -1:
                    similarity = similar(filename.rsplit('_na_', 1)[0][:76], protname[24:-4])
                    prots_similarity[protpath] = similarity
            maxsimilarity = max(zip(prots_similarity.values(), prots_similarity.keys()))[1]
            queue[file_id + 1] = os.path.basename(maxsimilarity)
            protlist.remove(maxsimilarity)

        queue_files = []
        queue_num_files = []
        protocols_for_concat = []
        self.num_protocols_eco = {}
        counter = 0
        # self.config.add_stamp == 'yes'
        all_keys = sorted(queue.keys())
        for i in sorted(queue.keys()):
            if self.config.paperecomode == "no" and self.config.concat_protocols == 'no':  # Если не экомод просто нумеруем файлы и протоколы, подставляя переменную count
                queue_files.append('{0}\\{1}'.format(foldername, queue[i]))
                self.num_protocols_eco[foldername + '\\' + f'{counter:02}_' + queue[i]] = 0
                queue_num_files.append(foldername + '\\' + f'{counter:02}_' + queue[i])
                if os.path.exists('{0}\\{1}'.format(foldername, queue[i])):
                    counter += 1
            else:
                if i + 1 in all_keys:  # Если экомод, то проверяем есть ли протокол для файла
                    if queue[i + 1].startswith(
                            'Protokol_proverki_fayla_'):  # Если следующий протокол, то склеиваем с текущим, если нет, то хз??
                        if self.config.concat_protocols == 'no':
                            merged_file = concat_pdfs(['{0}\\{1}'.format(foldername, queue[i]),
                                                       '{0}\\{1}'.format(foldername, queue[i + 1])], True)
                        else:
                            merged_file = '{0}\\{1}'.format(foldername, queue[i])
                            protocols_for_concat.append('{0}\\{1}'.format(foldername, queue[i + 1]))
                        queue_files.append(merged_file)
                        numered_file = foldername + '\\' + f'{counter:02}_' + queue[i]
                        queue_num_files.append(numered_file)
                        counter += 1
                else:
                    if not queue[i].startswith('Protokol_proverki_fayla_'):
                        queue_files.append('{0}\\{1}'.format(foldername, queue[i]))
                        numered_file = foldername + '\\' + f'{counter:02}_' + queue[i]
                        queue_num_files.append(numered_file)
                        self.num_protocols_eco[numered_file] = 0
                        counter += 1
        self.files_for_print = []
        self.files_for_print_stamps = {}
        for i, j in zip(queue_files, queue_num_files):
            if os.path.exists(i):
                os.replace(i, j)
                self.num_pages[j] = check_num_pages(j)
                self.num_protocols_eco[j] = int(self.num_pages[j][0] % 2 != 1)
                self.files_for_print.append(j)
        if self.config.add_stamp == 'yes':
            counter_stamp = -1
            for i in range(len(self.files_for_print)):
                if counter_stamp < 0:
                    num_doc = 'Квитанция'
                elif counter_stamp == 0:
                    num_doc = 'Суть заявления'
                elif counter_stamp > 0:
                    num_doc = 'Приложение ' + str(counter_stamp)
                if not os.path.basename(self.files_for_print[i])[3:].startswith('Protokol_proverki_fayla_'):
                    self.files_for_print_stamps[self.files_for_print[i]] = [num_appeal, num_doc]
                    counter_stamp += 1
        if self.config.concat_protocols == 'yes' and self.config.no_protocols == 'no':
            protocols_concatenated = concat_pdfs(protocols_for_concat, True)
            new_name = foldername + '\\Protocols.pdf'
            try:
                os.rename(protocols_concatenated, new_name)
            except:
                new_name = protocols_concatenated
                pass
            self.files_for_print.append(new_name)
            self.num_pages[new_name] = check_num_pages(new_name)
            self.num_protocols_eco[new_name] = int(self.num_pages[new_name][0] % 2 != 1)

        if self.config.save_stat == 'yes':
            docnumber = os.path.basename(givenpath).split('_', 1)[0]
            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
            self.stat.statdict['Дата и время'] = dt_string
            self.stat.statdict['Номер'] = docnumber
            self.stat.statdict['Кол-во док-ов'] = (len(self.files_for_print))
            self.stat.statdict['Кол-во страниц в документах'] = (sum(i[0] for i in self.num_pages.values()))
            self.stat.statdict['Кол-во листов в документах'] = (sum(i[1] for i in self.num_pages.values()))
        if self.config.print_directly == 'no':
            os.startfile(foldername)
