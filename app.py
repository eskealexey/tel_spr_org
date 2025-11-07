import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from class_tel_spr import load_xls


class App:
    def __init__(self):
        self.name = "app"
        self.version = "0.0.1"
        self.author = "Eske Alexey"
        self.description = "Телефонный справочник"

class Window(App):
    def __init__(self, root):
        self.app = App()  # Композиция вместо наследования
        self.root = root
        self.root.geometry("900x500")
        self.root.title(f'{self.app.description} v.{self.app.version}')
        self.root.resizable(True, True)
        self.create_menu()
        self.create_widgets()

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
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def create_widgets(self):
        frame_top = tk.Frame(self.root)
        frame_top.pack(padx=10, pady=10, fill='x')

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Вкладка 1: Контакты
        self.frame_osfr= ttk.Frame(self.notebook)
        self.notebook.add(self.frame_osfr, text="ОСФР")

        # Вкладка 2: Отделы
        self.frame_ks = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_ks, text="Клиентские службы")

        # Создаем виджеты для вкладки "ОСФР"
        self.create_osfr_tab()




    def create_osfr_tab(self):
        frame_top = tk.Frame(self.frame_osfr)
        frame_top.pack(padx=10, pady=10, fill='x')
        # tk.Label(frame_top, text="МОЛ:", font=('Arial', 10)).pack(side='left')
        # # Создаём Treeview и Scrollbar
        # self.tree = ttk.Treeview(self.root)
        # scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        #
        # # Позиционируем виджеты
        # self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=5, pady=5)
        # scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        #
        # # Связываем Treeview с Scrollbar
        # self.tree.config(yscrollcommand=scrollbar.set)

