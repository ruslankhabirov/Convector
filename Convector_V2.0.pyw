from openpyxl import load_workbook
from tkinter import *
import winsound as ws
import tkinter.scrolledtext as st

"""Кортеж имен столбцов excel-таблицы"""
LETTERS = ('M', 'N', 'O', 'P', 'Q', 'R')


def append_into_table(file_name: str, table_name: str, page_name: str, tablets: int, targets: int):
    """Функция использется для создания матрицы из элементов текстового файла
    и дальнейшей распаковки элементов матрицы в нужные ячейки excel-таблицы.
    В качестве разделителя используются отрицательные значения из текстового файла."""
    with open(file_name) as file:
        array = [row.strip() for row in file]

    for elements in range(len(array)):
        array[elements] = array[elements].split()

    for elements in range(len(array)):
        array[elements] = float(array[elements][1])

    matrix = []

    for number in range(len(array)):
        """Создаем список кортежей: матрицу изъятых из текстового файла элементов"""
        if array[number] < 0:
            matrix.append((array[number-3], array[number-2], array[number-1]))
    matrix.append((array[-3], array[-2], array[-1]))

    work_book = load_workbook(table_name, data_only=False)
    page_value = work_book[page_name]

    step = 0

    for targets_counter in range(targets):
        """Изымаем из матрицы и распаковываем в нужные ячейки таблицы числовые значения"""
        for tablets_counter in range(tablets):
            page_value[LETTERS[tablets_counter] + str(7 + 3 * targets_counter)].value = matrix[step][0]
            page_value[LETTERS[tablets_counter] + str(8 + 3 * targets_counter)].value = matrix[step][1]
            page_value[LETTERS[tablets_counter] + str(9 + 3 * targets_counter)].value = matrix[step][2]
            work_book.save(table_name)
            step += 1


def help_btn():
    HelpWindow().text_creator()


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

    def click_btn(self):
        """Функция-кнопка, проверяющая вводимые пользователем значения"""
        try:
            """Проверяем на наличие неподходящих символов внутри числовых строк ввода"""
            int(self.input_tablets_row.get())
            int(self.input_targets_row.get())
        except ValueError:
            ChildWindows().value_error()

        tabl = int(self.input_tablets_row.get())
        targ = int(self.input_targets_row.get())
        file_name = str(self.input_file_name_row.get())
        excel_table_name = str(self.input_table_name_row.get())

        if tabl > 6:
            """Проверяем, не введено ли больше, чем 6 таблеток"""
            ChildWindows().excess_of_the_value()
            assert (tabl <= 6)
        elif tabl <= 0:
            """Проверка на отрицательные или нулевые значения"""
            ChildWindows().negative_value_error()
            assert (tabl > 0)
        elif targ <= 0:
            """Проверка на отрицательные или нулевые значения"""
            ChildWindows().negative_value_error()
            assert (targ > 0)
        elif file_name[-1:-5:-1] != "txt.":
            """Проверка на наличие txt-расширения у текстового файла"""
            ChildWindows().not_txt_error()
            assert (file_name[-1:-5:-1] == "txt.")
        elif ("slx." not in excel_table_name[-1:-5:-1]) and \
                ("mslx." not in excel_table_name[-1:-6:-1]) and \
                ("xslx." not in excel_table_name[-1:-6:-1]):
            """Проверка на наличие поддерживаемых библиотекой расширений у excel-файла"""
            ChildWindows().not_excel_error()
            assert (("slx." in excel_table_name[-1:-5:-1]) or (
                    "mslx." in excel_table_name[-1:-6:-1]) or (
                    "xslx." in excel_table_name[-1:-6:-1]))
        try:
            append_into_table(str(self.input_file_name_row.get()),
                              str(self.input_table_name_row.get()),
                              str(self.input_page_name_row.get()),
                              int(self.input_tablets_row.get()),
                              int(self.input_targets_row.get()))
        except FileNotFoundError:
            """Проверка на наличие указанных файлов в текущей директории"""
            ChildWindows().file_not_found_error()
        except KeyError:
            """Проверка на наличие внутри книги листа с указанным названием"""
            ChildWindows().key_error()
        except IndexError:
            """Проверка на наличие значений в txt-файле или на корректность ввода количества таблеток"""
            ChildWindows().index_error()


class ChildWindows(Toplevel):
    """Класс, наследуемый от Toplevel - отвечает за создание всплывающих
    окон с ошибками"""
    def __init__(self):
        Toplevel.__init__(self)

        """Размещаем дочерний экран по центру, используя прежний алгоритм"""
        w = 230
        h = 130

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) / 2
        y = (sh - h) / 2

        self.title("Ошибка!")
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

        """Запрещаем изменять размер экрана"""
        self.resizable(False, False)

        """Запрещаем работу с главным окном, пока не будет закрыто дочернее окно"""
        self.grab_set()
        self.focus_set()

        """Воспроизводим звук системной ошибки каждый раз, когда запускается экземпляр класса"""
        ws.PlaySound("*", ws.SND_ASYNC)

    def value_error(self):
        """Функция, отвечающая за ошибки ввода строк в числовые поля"""
        Label(self, text="Введены некорректные значения").place(x=20, y=33)
        Button(self, text="OK", width=20, command=self.destroy).place(x=40, y=63)

    def excess_of_the_value(self):
        """Функция, отвечающая за ошибку введения более, чем 6-ти таблеток"""
        Label(self, text="Доступен обсчет не более 6 таблеток").place(x=12, y=33)
        Button(self, text="OK", width=20, command=self.destroy).place(x=40, y=63)

    def negative_value_error(self):
        """Функция, отвечающая за ошибки ввода отрицательных или нулевых значений"""
        Label(self, text="Отрицательное или нулевое значение").place(x=6, y=33)
        Button(self, text="OK", width=20, command=self.destroy).place(x=40, y=63)

    def not_txt_error(self):
        """Функция, отвечающая за проверку наличия .txt-расширения у файла"""
        Label(self, text="В названии текстового файла \n упущено расширение .txt").place(x=26, y=23)
        Button(self, text="OK", width=20, command=self.destroy).place(x=40, y=73)

    def not_excel_error(self):
        """Функция, отвечающая за проверку наличия расширений для Excel-файлов"""
        Label(self, text="В названии excel файла упущено \n расширение .xls, .xlsm или .xlsx").place(x=15, y=23)
        Button(self, text="OK", width=20, command=self.destroy).place(x=40, y=73)

    def file_not_found_error(self):
        """Функция, отвечающая за проверку наличия указаных файлов в текущей директории"""
        Label(self, text="Файл с таким именем не найден \n в текущей директории").place(x=18, y=23)
        Button(self, text="OK", width=20, command=self.destroy).place(x=40, y=73)

    def key_error(self):
        """Функция, отвечающая за проверку наличия указанного листа в Excel-файле"""
        Label(self, text="Лист с таким именем не найден").place(x=22, y=23)
        Button(self, text="OK", width=20, command=self.destroy).place(x=40, y=63)

    def index_error(self):
        """Функция, отвечающая за проверку наличия в txt-файле значений, а также
        за корректность ввода пользователем количества таблеток и точек"""
        Label(self, text="txt-файл не содержит значений или\nвведено неверное количество точек").place(x=10, y=23)
        Button(self, text="OK", width=20, command=self.destroy).place(x=40, y=73)


class HelpWindow(Toplevel):
    """Класс, наследуемый от Toplevel - отвечает за инициализацию
    окна помощи для пользователя"""
    def __init__(self):
        Toplevel.__init__(self)

        """Размещаем экран помощи по центру. Алгоритм без изменений"""
        w = 650
        h = 450

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) / 2
        y = (sh - h) / 2

        self.title("Справка")
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

        """Запрещаем изменять размер экрана"""
        self.resizable(False, True)

        """Запрещаем работу с главным окном, пока не будет закрыто дочернее окно"""
        self.grab_set()
        self.focus_set()

        self.minsize(650, 450)
        self.maxsize(650, 790)

    def text_creator(self):
        """Функция, поомещающая текс в окно вывода и создающая кнопку
        для выхода из 'окна справки' """

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
