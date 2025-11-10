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

class Window:
    def __init__(self, root):
        self.app = App()  # Композиция вместо наследования
        self.root = root
        self.root.geometry("1027x500")
        self.root.title(f'{self.app.description} v.{self.app.version}')
        self.root.resizable(True, True)
        self.create_menu()
        self.create_widgets()
        self.load_json_osfr('JSON/osfr.json')

    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Открыть", command=self.open_file_xls)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit)
        menubar.add_cascade(label="Файл", menu=file_menu)
        self.root.config(menu=menubar)

    def open_file_xls(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            load_xls(file_path)
            self.load_json_osfr('JSON/osfr.json')
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def create_widgets(self):
        frame_top = tk.Frame(self.root)
        frame_top.pack(padx=10, pady=10, fill='x')

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Вкладка 1: телефоны ОСФР
        self.frame_osfr= ttk.Frame(self.notebook)
        self.notebook.add(self.frame_osfr, text="ОСФР")

        # Вкладка 2: телефоны клиентских служб
        self.frame_ks = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_ks, text="Клиентские службы")

        # Создаем виджеты для вкладки "ОСФР"
        self.create_osfr_tab()
        self.load_json_osfr('JSON/osfr.json')

    def load_json_osfr(self, file_path):
        """Загрузка и отображение JSON данных"""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                data = json.load(file)

            # Очищаем существующие данные в таблице
            for item in self.tree.get_children():
                self.tree.delete(item)

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
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def create_osfr_tab(self):
        # Верхняя панель
        frame_top = tk.Frame(self.frame_osfr)
        frame_top.pack(padx=10, pady=10, fill='x')

        # Создаем и настраиваем стиль
        style = ttk.Style()
        style.configure("Custom.Treeview.Heading",
                        padding=(0, 5, 0, 25),
                        # rowheight=40,  # Высота строки заголовка
                        font=('Arial', 10, 'bold'))  # Можно уменьшить шрифт если нужно

        columns = ("city_phone", "internal_phone", "room", "last_name", "first_name", "patronymic", "position",
                   "department", "location")

        # Многострочные заголовки
        headings = {
            "city_phone": "Городской\nномер",
            "internal_phone": "Корпоративный\nтелефон",
            "room": "Номер\nкомнаты",
            "last_name": "Фамилия",
            "first_name": "Имя",
            "patronymic": "Отчество",
            "position": "Должность",
            "department": "Отдел",
            "location": "Место\nрасположения"
        }
        # self.tree = ttk.Treeview(self.frame_osfr, columns=columns, show="headings", height=15)
        self.tree = ttk.Treeview(self.frame_osfr, columns=columns, show="headings",
                                 height=15, style="Custom.Treeview")
        # Сначала настраиваем все заголовки
        for col in columns:
            self.tree.heading(col, text=headings[col], anchor="center")
            self.tree.column(col, width=100, anchor="w")

        # Затем настраиваем индивидуальные ширины (ВНЕ цикла)
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
        scrollbar = ttk.Scrollbar(self.frame_osfr, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Позиционируем виджеты
        self.tree.pack(fill='both', expand=True, side='left', padx=5, pady=5)
        scrollbar.pack(side='right', fill='y')

