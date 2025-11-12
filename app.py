#

import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from class_tel_spr import load_xls


class App:
    def __init__(self):
        self.name = "app"
        self.version = "0.0.1"
        self.author = "Eske Alexey"
        self.description = "Телефонный справочник"


class BaseTab:
    """Базовый класс для вкладок"""

    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.original_data = []
        self.create_widgets()

    def create_widgets(self):
        """Метод для создания виджетов вкладки (должен быть переопределен)"""
        pass

    def load_data(self, file_path):
        """Метод для загрузки данных (должен быть переопределен)"""
        pass

    def clear_table(self):
        """Очистка таблицы"""
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                self.tree.delete(item)


class Osfr(BaseTab):
    """Класс для вкладки ОСФР"""

    def __init__(self, parent):
        self.podrazdel_list = ["--- Выберите подразделение ---", ]
        super().__init__(parent)

    def create_widgets(self):
        """Создает виджеты для вкладки ОСФР"""
        # Основной контейнер для фильтров
        filters_container = tk.Frame(self.frame)
        filters_container.pack(padx=10, pady=10, fill='x')

        # ПЕРВАЯ СТРОКА: Поиск по ФИО и телефонам
        search_frame = tk.Frame(filters_container)
        search_frame.pack(fill='x', pady=(0, 10))

        label_search = tk.Label(search_frame, text="Поиск (ФИО/телефон):", font=('Arial', 10))
        label_search.pack(side='left', padx=5)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=('Arial', 10),
            width=40
        )
        self.search_entry.pack(side='left', padx=5)
        self.search_entry.bind("<KeyRelease>", self.search_data)

        # Кнопка сброса поиска
        btn_clear_search = tk.Button(
            search_frame,
            text="Очистить поиск",
            command=self.clear_search,
            font=('Arial', 9)
        )
        btn_clear_search.pack(side='left', padx=5)

        # ВТОРАЯ СТРОКА: Фильтр по подразделениям
        dept_frame = tk.Frame(filters_container)
        dept_frame.pack(fill='x', pady=(0, 10))

        label_dept = tk.Label(dept_frame, text="Фильтр по подразделению:", font=('Arial', 10))
        label_dept.pack(side='left', padx=5)

        self.podrazdel_var = tk.StringVar()
        self.podrazdel_combobox = ttk.Combobox(
            dept_frame,
            textvariable=self.podrazdel_var,
            state="normal",
            font=('Arial', 10),
            width=100
        )
        self.podrazdel_combobox.pack(side='left', padx=5)
        self.podrazdel_combobox.bind("<KeyRelease>", self.podrazdel_filter)
        self.podrazdel_combobox.bind("<<ComboboxSelected>>", self.on_select)

        # Кнопка сброса фильтра подразделения
        btn_reset_dept = tk.Button(
            dept_frame,
            text="Сбросить фильтр",
            command=self.reset_filter,
            font=('Arial', 9)
        )
        btn_reset_dept.pack(side='left', padx=5)

        # Создаем таблицу
        self.create_table()

    def create_table(self):
        """Создает таблицу для отображения данных"""
        # Создаем и настраиваем стиль
        style = ttk.Style()
        style.configure("Custom.Treeview.Heading",
                        padding=(0, 5, 0, 25),
                        font=('Arial', 10, 'bold'))

        columns = ("city_phone", "internal_phone", "room", "last_name", "first_name", "patronymic", "position",
                   "department", "location")

        # Многострочные заголовки
        headings = {
            "city_phone": "Городской\nномер",
            "internal_phone": "Корп.\nтелефон",
            "room": "Номер\nкомнаты",
            "last_name": "Фамилия",
            "first_name": "Имя",
            "patronymic": "Отчество",
            "position": "Должность",
            "department": "Отдел",
            "location": "Место\nрасположения"
        }

        self.tree = ttk.Treeview(self.frame, columns=columns, show="headings",
                                 height=15, style="Custom.Treeview")

        # Настраиваем заголовки и колонки
        for col in columns:
            self.tree.heading(col, text=headings[col], anchor="center")
            self.tree.column(col, width=100, anchor="w")

        # Индивидуальные ширины колонок
        self.tree.column("last_name", width=120)
        self.tree.column("first_name", width=100)
        self.tree.column("patronymic", width=120)
        self.tree.column("position", width=150)
        self.tree.column("department", width=120)
        self.tree.column("city_phone", width=100)
        self.tree.column("internal_phone", width=100)
        self.tree.column("room", width=80)
        self.tree.column("location", width=200)

        # Scrollbar для таблицы
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Позиционируем виджеты
        self.tree.pack(fill='both', expand=True, side='left', padx=5, pady=5)
        scrollbar.pack(side='right', fill='y')

    def load_data(self, file_path):
        """Загрузка и отображение JSON данных"""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)

            # Сохраняем оригинальные данные для фильтрации
            self.original_data = data

            # Очищаем существующие данные в таблице
            self.clear_table()

            # Добавляем данные в таблицу
            for item in data:
                self.tree.insert("", "end", values=(
                    item.get("Городской номер", ""),
                    item.get("Кор. тел.", ""),
                    item.get("№ комн.", ""),
                    item.get("ФАМИЛИЯ", ""),
                    item.get("ИМЯ", ""),
                    item.get("ОТЧЕСТВО", ""),
                    item.get("ДОЛЖНОСТЬ", ""),
                    item.get("Отдел", ""),
                    item.get("Место расположения", "")
                ))

            self.podrazdel_list = self.create_podrazdel_list(data)
            self.update_combobox()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Нет данных для отображения:\nЗагрузите файл с номерами телефонов")
            print(str(e))

    def create_podrazdel_list(self, data):
        """Создает список подразделений"""
        lst = []
        for item in data:
            if item.get("Отдел") not in lst and item.get("Отдел") != "":
                lst.append(item.get("Отдел"))
        lst = ["--- Выберите подразделение ---", ] + lst
        return lst

    def podrazdel_filter(self, event=None):
        """Фильтрация списка подразделений по введенному тексту"""
        current_text = self.podrazdel_var.get()
        if current_text:
            filtered_list = [item for item in self.podrazdel_list if current_text.lower() in item.lower()]
            self.podrazdel_combobox['values'] = filtered_list
        else:
            self.podrazdel_combobox['values'] = self.podrazdel_list

    def search_data(self, event=None):
        """Поиск данных по ФИО и номерам телефонов"""
        self.apply_filters()

    def apply_filters(self):
        """Применяет все активные фильтры"""
        selected_dept = self.podrazdel_var.get()
        search_text = self.search_var.get().lower().strip()

        # Очищаем таблицу
        self.clear_table()

        # Начинаем с оригинальных данных
        filtered_data = self.original_data

        # Применяем фильтр по подразделению
        if selected_dept and selected_dept != "--- Выберите подразделение ---":
            filtered_data = [item for item in filtered_data if item.get("Отдел") == selected_dept]

        # Применяем поисковый фильтр
        if search_text:
            filtered_data = [item for item in filtered_data if (
                    search_text in item.get("ФАМИЛИЯ", "").lower() or
                    search_text in item.get("ИМЯ", "").lower() or
                    search_text in item.get("ОТЧЕСТВО", "").lower() or
                    search_text in item.get("Городской номер", "").lower() or
                    search_text in item.get("Кор. тел.", "").lower()
            )]

        # Добавляем отфильтрованные данные в таблицу
        for item in filtered_data:
            self.tree.insert("", "end", values=(
                item.get("Городской номер", ""),
                item.get("Кор. тел.", ""),
                item.get("№ комн.", ""),
                item.get("ФАМИЛИЯ", ""),
                item.get("ИМЯ", ""),
                item.get("ОТЧЕСТВО", ""),
                item.get("ДОЛЖНОСТЬ", ""),
                item.get("Отдел", ""),
                item.get("Место расположения", "")
            ))

    def on_select(self, event=None):
        """Обработчик выбора подразделения"""
        self.apply_filters()

    def clear_search(self):
        """Очистка поля поиска"""
        self.search_var.set("")
        self.apply_filters()

    def reset_filter(self):
        """Сброс фильтра подразделения"""
        self.podrazdel_var.set("")
        self.apply_filters()

    def update_combobox(self):
        """Обновление значений combobox"""
        self.podrazdel_combobox['values'] = self.podrazdel_list


class ClientService(BaseTab):
    """Класс для вкладки Клиентские службы"""

    def __init__(self, parent):
        super().__init__(parent)

    def create_widgets(self):
        """Создает виджеты для вкладки Клиентские службы"""
        # Основной контейнер для фильтров
        filters_container = tk.Frame(self.frame)
        filters_container.pack(padx=10, pady=10, fill='x')

        # ПЕРВАЯ СТРОКА: Поиск по ФИО и телефонам
        search_frame = tk.Frame(filters_container)
        search_frame.pack(fill='x', pady=(0, 10))
        #
        label_search = tk.Label(search_frame, text="Поиск (ФИО/телефон):", font=('Arial', 10))
        label_search.pack(side='left', padx=5)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=('Arial', 10),
            width=40
        )
        self.search_entry.pack(side='left', padx=5)
        self.search_entry.bind("<KeyRelease>", self.search_data)

        # Кнопка сброса поиска
        btn_clear_search = tk.Button(
            search_frame,
            text="Очистить поиск",
            command=self.clear_search,
            font=('Arial', 9)
        )
        btn_clear_search.pack(side='left', padx=5)

        # # ВТОРАЯ СТРОКА: Фильтр по подразделениям
        dept_frame = tk.Frame(filters_container)
        dept_frame.pack(fill='x', pady=(0, 10))

        label_dept = tk.Label(dept_frame, text="Фильтр по подразделению:", font=('Arial', 10))
        label_dept.pack(side='left', padx=5)

        self.podrazdel_var = tk.StringVar()
        self.podrazdel_combobox = ttk.Combobox(
            dept_frame,
            textvariable=self.podrazdel_var,
            state="normal",
            font=('Arial', 10),
            width=100
        )
        self.podrazdel_combobox.pack(side='left', padx=5)
        self.podrazdel_combobox.bind("<KeyRelease>", self.podrazdel_filter)
        self.podrazdel_combobox.bind("<<ComboboxSelected>>", self.on_select)

        # Кнопка сброса фильтра подразделения
        btn_reset_dept = tk.Button(
            dept_frame,
            text="Сбросить фильтр",
            command=self.reset_filter,
            font=('Arial', 9)
        )
        btn_reset_dept.pack(side='left', padx=5)

        # Создаем таблицу
        self.create_table()

    def create_podrazdel_list(self, data):
        """Создает список подразделений"""
        lst = []
        for item in data:
            if item.get("отдел") not in lst and item.get("отдел") != "":
                lst.append(item.get("отдел"))
        lst = ["--- Выберите Клиентскую службу ---", ] + lst
        return lst

    def search_data(self, event=None):
        """Поиск данных по ФИО и номерам телефонов"""
        self.apply_filters()

    def update_combobox(self):
        """Обновление значений combobox"""
        self.podrazdel_combobox['values'] = self.podrazdel_list

    def podrazdel_filter(self, event=None):
        """Фильтрация списка подразделений по введенному тексту"""
        current_text = self.podrazdel_var.get()
        if current_text:
            filtered_list = [item for item in self.podrazdel_list if current_text.lower() in item.lower()]
            self.podrazdel_combobox['values'] = filtered_list
        else:
            self.podrazdel_combobox['values'] = self.podrazdel_list

    def create_table(self):
        """Создает таблицу для отображения данных"""
        # Создаем и настраиваем стиль
        style = ttk.Style()
        style.configure("Custom.Treeview.Heading",
                        padding=(0, 5, 0, 25),
                        font=('Arial', 10, 'bold'))

        columns = ("city_phone", "internal_phone", "last_name", "first_name", "patronymic", "position",
                   "department", "location")

        # Многострочные заголовки
        headings = {
            "city_phone": "Городской\nномер",
            "internal_phone": "Корп.\nтелефон",
            "last_name": "Фамилия",
            "first_name": "Имя",
            "patronymic": "Отчество",
            "position": "Должность",
            "department": "Клиентская\nслужба",
            "location": "Место\nрасположения"
        }

        self.tree = ttk.Treeview(self.frame, columns=columns, show="headings",
                                 height=15, style="Custom.Treeview")

        # Настраиваем заголовки и колонки
        for col in columns:
            self.tree.heading(col, text=headings[col], anchor="center")
            self.tree.column(col, width=100, anchor="w")

        # Индивидуальные ширины колонок
        self.tree.column("last_name", width=120)
        self.tree.column("first_name", width=100)
        self.tree.column("patronymic", width=120)
        self.tree.column("position", width=150)
        self.tree.column("department", width=120)
        self.tree.column("city_phone", width=100)
        self.tree.column("internal_phone", width=100)
        self.tree.column("location", width=200)

        # Scrollbar для таблицы
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Позиционируем виджеты
        self.tree.pack(fill='both', expand=True, side='left', padx=5, pady=5)
        scrollbar.pack(side='right', fill='y')



        # Здесь можно добавить специфичные для клиентских служб виджеты
        # Например, другую таблицу или фильтры

    def load_data(self, file_path='JSON/ks.json'):
        """Загрузка данных для клиентских служб"""
        # Реализация загрузки данных для клиентских служб

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)

            # Сохраняем оригинальные данные для фильтрации
            self.original_data = data

            # Очищаем существующие данные в таблице
            self.clear_table()

            # Добавляем данные в таблицу
            for item in data:
                self.tree.insert("", "end", values=(
                    item.get("город", ""),
                    item.get("кспд", ""),
                    item.get("Фамилия", ""),
                    item.get("Имя", ""),
                    item.get("Отчество", ""),
                    item.get("Должность", ""),
                    item.get("отдел", ""),
                    item.get("Место расположения", "")
                ))

            self.podrazdel_list = self.create_podrazdel_list(data)
            self.update_combobox()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Нет данных для отображения:\nЗагрузите файл с номерами телефонов")
            print(str(e))

    def apply_filters(self):
        """Применяет все активные фильтры"""
        selected_dept = self.podrazdel_var.get()
        search_text = self.search_var.get().lower().strip()

        # Очищаем таблицу
        self.clear_table()

        # Начинаем с оригинальных данных
        filtered_data = self.original_data

        # Применяем фильтр по подразделению
        if selected_dept and selected_dept != "--- Выберите Клиентскую службу ---":
            filtered_data = [item for item in filtered_data if item.get("отдел") == selected_dept]

        # Применяем поисковый фильтр
        if search_text:
            filtered_data = [item for item in filtered_data if (
                    search_text in item.get("Фамилия", "").lower() or
                    search_text in item.get("Имя", "").lower() or
                    search_text in item.get("Отчество", "").lower() or
                    search_text in item.get("город", "").lower() or
                    search_text in item.get("кспд", "").lower()
            )]

        # Добавляем отфильтрованные данные в таблицу
        for item in filtered_data:
            self.tree.insert("", "end", values=(
                item.get("город", ""),
                item.get("кспд", ""),
                item.get("Фамилия", ""),
                item.get("Имя", ""),
                item.get("Отчество", ""),
                item.get("Должность", ""),
                item.get("отдел", ""),
                item.get("Место расположения", "")
            ))

    def on_select(self, event=None):
        """Обработчик выбора подразделения"""
        self.apply_filters()

    def clear_search(self):
        """Очистка поля поиска"""
        self.search_var.set("")
        self.apply_filters()

    def reset_filter(self):
        """Сброс фильтра подразделения"""
        self.podrazdel_var.set("")
        self.apply_filters()


class Window:
    """
    Класс окна приложения
    """

    def __init__(self, root):
        self.app = App()
        self.root = root
        self.root.geometry("1200x500")
        self.root.title(f'{self.app.description} v.{self.app.version}')
        self.root.resizable(True, True)

        # Создаем экземпляры вкладок
        self.osfr_tab = None
        self.client_service_tab = None

        self.create_menu()
        self.create_widgets()

    def create_menu(self):
        """Создает меню для приложения"""
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Открыть", command=self.open_file_xls)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.destroy)
        menubar.add_cascade(label="Файл", menu=file_menu)
        self.root.config(menu=menubar)

    def open_file_xls(self):
        """Открывает диалоговое окно для выбора файла"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            load_xls(file_path)
            if self.osfr_tab:
                self.osfr_tab.load_data('JSON/osfr.json')
            if self.client_service_tab:
                self.client_service_tab.load_data('JSON/ks.json')
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def create_widgets(self):
        """Создает виджеты для окна приложения"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Создаем вкладки
        self.osfr_tab = Osfr(self.notebook)
        self.client_service_tab = ClientService(self.notebook)

        # Добавляем вкладки в notebook
        self.notebook.add(self.osfr_tab.frame, text="ОСФР")
        self.notebook.add(self.client_service_tab.frame, text="Клиентские службы")

        # Загружаем начальные данные
        try:
            self.osfr_tab.load_data('JSON/osfr.json')
            self.client_service_tab.load_data('JSON/ks.json')
        except:
            pass  # Если файла нет, просто пропускаем