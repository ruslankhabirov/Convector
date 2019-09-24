from tkinter import *
import tkinter.scrolledtext as st
import tkinter.ttk as ttk
import winsound as ws
from collections import OrderedDict
from functools import partial
import threading

from openpyxl import load_workbook

"""Кортеж имен столбцов excel-таблицы"""
LETTERS = ('M', 'N', 'O', 'P', 'Q', 'R')

"""Словарь с текстовыми значениями возможных ошибок"""
text_dictionary = {"value_error": "Введены некорректные значения",
                   "excess_of_the_value": "Доступен обсчет не более 6 таблеток",
                   "negative_value_error": "Отрицательное или нулевое значение",
                   "not_txt_error": "В названии текстового файла \n упущено расширение .txt",
                   "not_excel_error": "В названии excel файла упущено \n расширение .xls, .xlsm или .xlsx",
                   "file_not_found_error": "Файл с таким именем не найден \n в текущей директории",
                   "key_error": "Лист с таким именем не найден",
                   "index_error": "txt-файл не содержит значений или\nвведено неверное количество точек",
                   "saving_error": "Excel-файл уже открыт"}


def extract_data(file_name: str):

    matrix = []

    with open(file_name, "r") as file:
        data = float(file.readline().strip().split()[1])
        buffer_list = []
        while data:
            if data < 0:
                matrix.append((buffer_list[-3], buffer_list[-2], buffer_list[-1]))
                buffer_list.clear()
            else:
                buffer_list.append(data)
            try:
                data = float(file.readline().strip().split()[1])
            except IndexError:
                matrix.append((buffer_list[-3], buffer_list[-2], buffer_list[-1]))
                break
    return matrix


def create_dictionary(matrix: list, letters: tuple, tablets: int, targets: int) -> dict:
    matrix_dictionary = OrderedDict({})
    step = 0
    for targets_counter in range(targets):
        for tablets_counter in range(tablets):
            matrix_dictionary[letters[tablets_counter] + str(7 + 3 * targets_counter)] = matrix[step][0]
            matrix_dictionary[letters[tablets_counter] + str(8 + 3 * targets_counter)] = matrix[step][1]
            matrix_dictionary[letters[tablets_counter] + str(9 + 3 * targets_counter)] = matrix[step][2]
            step += 1
    return matrix_dictionary


def add_in_table(work_book, dict_key: str, dict_value: float):
    work_book[dict_key].value = dict_value


def help_btn():
    ChildWindows(650, 690).help_request(650, 450, 650, 790)


class MainWindowClass(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.pack(fill=BOTH, expand=1)

        """Подписываем главное окно"""
        self.parent.title("BIOCAD Сonvector")

        """Указываем размеры главного окна"""
        w = 470
        h = 190

        """Получаем текущий размер экрана пользователя"""
        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()

        """При открытии файла, размещаем главное окно по центру пользовательского экрана"""
        x = (sw - w) / 2
        y = (sh - h) / 2
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))

        """Запрещаем изменять размер окна"""
        self.parent.resizable(False, False)

        """Явным образом указываем допустимые значения для полей ввода"""
        self.input_file_name = StringVar()
        self.input_table_name = StringVar()
        self.input_page_name = StringVar()
        self.input_tablets = IntVar()
        self.input_targets = IntVar()

        """Создаем поля ввода значений и размещаем их на главном окне"""
        self.input_file_name_row = Entry(self.parent, width=45, textvariable=self.input_file_name)
        self.input_file_name_row.place(x=180, y=12)
        self.input_table_name_row = Entry(self.parent, width=45, textvariable=self.input_table_name)
        self.input_table_name_row.place(x=180, y=42)
        self.input_page_name_row = Entry(self.parent, width=45, textvariable=self.input_page_name)
        self.input_page_name_row.place(x=180, y=72)
        self.input_tablets_row = Entry(self.parent, width=45, textvariable=self.input_tablets)
        self.input_tablets_row.place(x=180, y=102)
        self.input_targets_row = Entry(self.parent, width=45, textvariable=self.input_targets)
        self.input_targets_row.place(x=180, y=132)

        """Подписываем поля ввода и размещаем их на главном окне"""
        Label(self.parent, text="Название текстового файла:").place(x=12, y=12)
        Label(self.parent, text="Название таблицы:").place(x=12, y=42)
        Label(self.parent, text="Название листа таблицы:").place(x=12, y=72)
        Label(self.parent, text="Количество таблеток:").place(x=12, y=102)
        Label(self.parent, text="Количество точек:").place(x=12, y=132)

        """Кнопка запуска"""
        Button(self.parent, text="OK", width=20, command=self.click_btn, bd=2).place(relx=0.34, rely=0.85)

        """Кнопка открытия окна помощи"""
        Button(self.parent, text="Help", width=20, command=help_btn, bd=2).place(relx=0.01, rely=0.85)

        """Кнопка выхода из программы"""
        Button(self.parent, text="Close", width=20, command=self.parent.destroy, bd=2).place(relx=0.67, rely=0.85)

    def int_check(self):
        """Проверяем на наличие неподходящих символов внутри числовых строк ввода"""
        try:
            int(self.input_tablets_row.get())
            int(self.input_targets_row.get())
        except ValueError:
            global text_dictionary
            ChildWindows(230, 130).create_window(20, 33, 40, 63, text_dictionary["value_error"])
            return 0
        else:
            return 1

    def tabl_check(self):
        """Проверяем, не введено ли больше, чем 6 таблеток"""
        tabl = int(self.input_tablets_row.get())
        if tabl > 6:
            global text_dictionary
            ChildWindows(230, 120).create_window(12, 33, 40, 63, text_dictionary["excess_of_the_value"])
            return 0
        return 1

    def targ_check(self):
        """Проверка на отрицательные или нулевые значения"""
        targ = int(self.input_targets_row.get())
        if targ <= 0:
            ChildWindows(230, 120).create_window(6, 33, 40, 63, text_dictionary["negative_value_error"])
            return 0
        return 1

    def tabl_negative_check(self):
        """Проверка на отрицательные или нулевые значения"""
        tabl = int(self.input_tablets_row.get())
        if tabl <= 0:
            ChildWindows(230, 120).create_window(6, 33, 40, 63, text_dictionary["negative_value_error"])
            return 0
        return 1

    def text_txt_check(self):
        """Проверка на наличие txt-расширения у текстового файла"""
        file_name = str(self.input_file_name_row.get())
        if not file_name.endswith(".txt"):
            ChildWindows(230, 120).create_window(26, 23, 40, 73, text_dictionary["not_txt_error"])
            return 0
        return 1

    def excel_name_check(self):
        """Проверка на наличие поддерживаемых библиотекой расширений у excel-файла"""
        excel_table_name = str(self.input_table_name_row.get())
        if (not excel_table_name.endswith(".xls")) and \
                (not excel_table_name.endswith(".xlsm")) and \
                (not excel_table_name.endswith(".xlsx")):
            ChildWindows(230, 120).create_window(15, 23, 40, 73, text_dictionary["not_excel_error"])
            return 0
        return 1

    def excel_input_check(self):
        targ = int(self.input_targets_row.get())
        tabl = int(self.input_tablets_row.get())
        file_name = str(self.input_file_name_row.get())
        excel_table_name = str(self.input_table_name_row.get())
        excel_page_name = str(self.input_page_name_row.get())
        try:
            work_book = load_workbook(excel_table_name, data_only=False)
            print("1 Таблица загружена")
            page_value = work_book[excel_page_name]
            print("2 Лист загружен")
            matrix_dictionary = create_dictionary(extract_data(file_name), LETTERS, tabl, targ)
            print("3 Матрица создана")
            partial_add_in_table = partial(add_in_table, page_value)
            print("4 Partial-функция создана")

            for i, j in list(zip([x for x in matrix_dictionary], [matrix_dictionary[x] for x in matrix_dictionary])):
                partial_add_in_table(i, j)
            print("5 функция выполнена")

            work_book.save(excel_table_name)
            print("6 Книга сохранена")
            work_book.close()
            print("7 Книга закрыта")

        except FileNotFoundError:
            """Проверка на наличие указанных файлов в текущей директории"""
            ChildWindows(230, 120).create_window(18, 23, 40, 73, text_dictionary["file_not_found_error"])
            return 0
        except KeyError:
            """Проверка на наличие внутри книги листа с указанным названием"""
            ChildWindows(230, 120).create_window(22, 23, 40, 63, text_dictionary["key_error"])
            return 0
        except IndexError:
            """Проверка на наличие значений в txt-файле или на корректность ввода количества таблеток"""
            ChildWindows(230, 120).create_window(10, 23, 40, 73, text_dictionary["index_error"])
            return 0
        except PermissionError:
            """Проверка на возможность сохранения изменений в книге"""
            ChildWindows(230, 120).create_window(47, 23, 40, 63, text_dictionary["saving_error"])
            return 0
        return 1

    function_names = [int_check, tabl_check, targ_check, tabl_negative_check, text_txt_check, excel_name_check,
                      excel_input_check]

    def click_btn(self):
        """Функция выполняет основное действие программы - поиск и помещение внутрь файла данных, предварительно
        проверяя поля ввода на корректность"""

        progress = ChildWindows(300, 100)
        progress.title("Прогресс выполнения")

        pb = ttk.Progressbar(progress, length=280, mode="determinate")
        pb.pack(expand=1)

        Label(progress, text="Выполнение программы").place(relx=0.25, rely=0.15)

        k = 0

        def local_gen():

            pb['value'] = 0

            for i in self.function_names:
                if i(self) == 0:
                    progress.destroy()
                    break
                else:
                    nonlocal k
                    k += 1
                    pb['value'] = int((k / 7) * 100)
            progress.destroy()

        thread1 = threading.Thread(target=local_gen)
        thread1.start()


class ChildWindows(Toplevel):
    """Класс, наследуемый от Toplevel - отвечает за создание всплывающих
    окон с ошибками и нотациями"""
    def __init__(self, w: int, h: int):
        Toplevel.__init__(self)

        """Размещаем дочерний экран по центру, используя прежний алгоритм"""
        self.w = w
        self.h = h

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) / 2
        y = (sh - h) / 2

        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

        """Запрещаем работу с главным окном, пока не будет закрыто дочернее окно"""
        self.grab_set()
        self.focus_set()

    def create_window(self, x_label: int, y_label: int, x_button: int, y_button: int, label_text: str):
        self.title("Ошибка!")

        """Размещаем кнопку и Label с нужным текстом"""
        Label(self, text=label_text).place(x=x_label, y=y_label)
        Button(self, text="OK", width=20, command=self.destroy).place(x=x_button, y=y_button)

        """Запрещаем изменять размер экрана"""
        self.resizable(False, False)

        """Воспроизводим звук системной ошибки каждый раз, когда запускается экземпляр класса с данным атрибутом"""
        ws.PlaySound("*", ws.SND_ASYNC)

    def help_request(self, min_x: int, min_y: int, max_x: int, max_y: int):
        self.title("Окно помощи")

        """Явным образом указываем минимальный и максимальный размер кона"""
        self.minsize(min_x, min_y)
        self.maxsize(max_x, max_y)

        """Создаем окно вывода текста со Scrollbar-ом"""
        console = st.ScrolledText(self, state='disable')

        """Вводим необходимый текст"""
        text_to_insert = "Поле «Название текстового файла»: сюда вносится название текстового файла,\n" \
                         "внутри которого содержатся исходные данные.\n" \
                         "Внутри текстового файла каждый отдельный образец должен быть отделён\n" \
                         "обнулением по воздуху (отрицательными значениями), иначе ничего работать\n" \
                         "не будет. Также в названии текстового файла не следует забывать ставить\n" \
                         "его расширение: «.txt».\n\n" \
                         "Поле «Название таблицы»: сюда запишите название таблицы, в которую\n" \
                         "вы хотите внести данные. Таблица должна быть взята из шаблона, который\n" \
                         "расположен на SharePoint Группы Скрининга малых молекул. Обращаем\n" \
                         "ваше внимание на то, что программа вносит данные только для одной среды.\n" \
                         "Если вы хотите внести в одном txt-файле сразу несколько сред, то следует\n" \
                         "разделить их по разным файлам и вносить поочерёдно.\n" \
                         "Для корректной работы приложения файл Excel должен быть закрыт, иначе\n" \
                         "программа не сможет зайти внутрь книги и отредактировать её.\n" \
                         "Важно ставить в конец этой строки расширение вашей таблицы\n" \
                         "(.xlsx, .xls, либо .xlsm).\n\n" \
                         "Поле «Название листа таблицы»: сюда записывается название листа таблицы\n" \
                         "excel, в которую пользователь хочет перенести данные из txt-файла.\n" \
                         "Название листа excel написано в левом нижнем углу книги. Скопируйте его и\n" \
                         "вставьте в точности. Важно знать, что для того, чтобы вставить в поле ввода\n" \
                         "программы скопированные с помощью «Ctrl + C» данные, необходимо находиться в\n" \
                         "режиме англоязычной раскладки.\n\n" \
                         "Поле «Количество таблеток»: сюда пользователь вносит своё количество таблеток.\n" \
                         "Программа поддерживает конвертацию не более, чем шести таблеток. В данное поле\n" \
                         "ввода следует вносить целое положительное число от 1 до 6 включительно.\n\n" \
                         "Поле «Количество точек»: сюда записывается необходимое пользователю\n" \
                         "количество точек. За количество точек принимается любое целое положительное\n" \
                         "число\n\n" \
                         "Все данные вносятся в лист excel ровно в одно место - в первую таблицу\n" \
                         "(в столбцы M-R, в строки с номерами от 7 до 7+(N*3-1), где N – количество\n" \
                         "точек). Чтобы выполнить перенос данных для нескольких сред, необходимо\n" \
                         "скопировать таблицу и перенести её ниже, а данные в первой таблице удалить.\n" \
                         "(скопируйте данные из первой таблицы, перенесите их ниже, закройте\n" \
                         "книгу Excel и ещё раз запустите программу для нового txt-файла)\n\n" \
                         "Крайне важно: для того, чтобы программа смогла найти конвертируемые файлы,\n" \
                         "все они должны лежать с ней в одной папке.\n\n" \
                         "Внутри текстового файла .txt не должно быть лишних данных: все избыточные\n" \
                         "строки со значениями стандартов и заголовками файла .bsr должны быть\n" \
                         "предварительно убраны пользователем вручную."

        console.configure(state='normal')
        console.insert(END, text_to_insert)

        """Помещаем зрение пользователя в начало текста"""
        console.yview()

        """Запрещаем вводить пользователю свои значения в текстовое поле"""
        console.configure(state='disabled')

        console.pack(fill=Y, expand=True)

        """Кнопка выхода из окна справки"""
        Button(self, text="Close", width=20, command=self.destroy, bd=2).pack(pady=10)


def main():
    root_element = Tk()
    MainWindowClass(root_element)
    root_element.mainloop()


if __name__ == "__main__":
    main()
