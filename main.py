import tkinter as tk
from app import Window
from class_tel_spr import load_xls

if __name__ == '__main__':
    root = tk.Tk()
    app = Window(root)
    # app.load_json('JSON/osfr.json')
    root.mainloop()
    # load_xls("telef.xls")
