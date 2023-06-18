import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import subprocess
import time
import os
from pywinauto import Application
from pywinauto import mouse
import pyautogui
import win32gui
import threading

class AutoClicker:
    def __init__(self, root):
        self.root = root
        self.root.title("Bot")
        pyautogui.FAILSAFE = False
        self.exe_file_path = None
        self.window_in_focus = False #флаг, что окно в фокусе
        # создаем кнопку для выбора exe-файла
        self.exe_file_button = tk.Button(
            self.root, text="Выбрать exe-файл", bg='lightyellow', command=self.choose_exe_file)
        self.exe_file_button.pack(fill='both')

        # создаем кнопку для добавления вкладок
        self.add_tab_btn = tk.Button(self.root, text="добавить вкладку", bg='lightblue', command=self.add_tab)
        self.add_tab_btn.pack(fill='both')

        #создаём панель вкладок
        self.notebook_style = ttk.Style()
        self.notebook_style.theme_create('rounded', parent='alt', settings={
            # Задаем свойство border-radius для закругления углов
            "TNotebook": {"configure": {"tabmargins": [2, 5, 2, 0], "border-radius": 10}},
            "TNotebook.Tab": {"configure": {"padding": [20, 5], "background": "#fff", "foreground": "#333"}},
            "TNotebook.Tab": {"map": {"background": [("selected", "#fff")], "foreground": [("selected", "#333")]}},
        })
        self.notebook_style.theme_use('rounded')
        self.notebook = ttk.Notebook() #панель вкладок
        self.notebook.pack()
        #
        self.frames = [] #фреймы вкладок
        self.bots = [] #объекты ботов
        for i in range(4):
            self.add_tab()

        # создаем кнопку для запуска окон приложений
        self.start_button = tk.Button(
            self.root, text="открыть окна приложений", bg='lightgreen',command=self.run_apps)
        self.start_button.pack(fill='both')

    def close_tab(self):
        if len(self.frames) < 2:
            return
        current_tab_id = self.notebook.tabs().index(self.notebook.select())
        self.notebook.forget(self.notebook.index(current_tab_id)) #закрываю вкладку
        del self.frames[current_tab_id]
        del self.bots[current_tab_id]
        for i in range(len(self.notebook.tabs())):
            self.notebook.tab(i, text="окно %s "%str(i+1))
            self.bots[i].bot_id = i

    def add_tab(self):
        panel_id=len(self.frames)
        if panel_id >= 5:
            return
        self.frames.append(tk.Frame())
        self.notebook.enable_traversal()
        self.notebook.add(self.frames[-1], text="окно %s " %str(panel_id+1))
        self.bots.append(Bot(self, bot_id=panel_id))

    def choose_exe_file(self):
        # вызываем диалоговое окно для выбора exe-файла
        self.exe_file_path = filedialog.askopenfilename(filetypes=[("EXE files", "*.exe")])
        if self.exe_file_path != '' and self.exe_file_path != None:
            self.exe_file_button.configure(text=os.path.basename(self.exe_file_path))

    def run_apps(self):
        '''метод для запуска окон приложений'''
        #проверка, какие вкладки нужно закрыть
        if self.exe_file_path == None:
            return
        checked = False
        while checked == False:
            for i in range(len(self.bots)):
                #если осталась всего 1 вкладка, и она без прописанных клавиш - выхожу из метода
                if len(self.bots) == 1:
                    if len(self.bots[0].keys) == 0:
                        self.notebook.tab(0, text="окно 1 ")
                        return
                    else:
                        checked = True
                        break
                self.bots[i].step()
                if len(self.bots[i].keys) == 0: #проверяю, есть-ли список кнопок. Если нет - удаляю вкладку и из списков
                    self.notebook.forget(self.notebook.index(i)) #закрываю вкладку
                    del self.frames[i]
                    del self.bots[i]
                    break
                if i == len(self.bots)-1:
                    checked = True
        #обновляю названия вкладок
        for n in range(len(self.notebook.tabs())):
            self.notebook.tab(n, text="окно %s "%str(n+1))
            self.bots[n].bot_id = n

        # запускаем exe-файл нужное количество раз
        self.apps = [] #список запущенных экземпляров
        self.main_windows = [] #список с гл. окнами экземпляров
        for i in range(len(self.bots)):
            self.apps.append(Application().start(self.exe_file_path))
            time.sleep(1)
            # Получаем идентификатор окна, которое сейчас активно
            hwnd = win32gui.GetForegroundWindow()
            # Получаем class_name этого окна
            class_name = win32gui.GetClassName(hwnd)
            self.main_windows.append(self.apps[-1].window(class_name=class_name))

        #тут нужно поменять кнопку на СТАРТ-----------------------------------------------------------------------------
        self.start_button.destroy()
        # создаем кнопку для запуска автокликера
        self.start_button = tk.Button(
            self.root, text="запустить бота!", bg='lightgreen',command=self.start_threads)
        self.start_button.pack(fill='both')

    def start_threads(self):
        #запускаем потоки ботов            
        for i in range(len(self.bots)):
            thread = threading.Thread(target=self.bots[i].actions_thread)
            thread.daemon = True
            thread.start()

class Bot():
    def __init__(self, parrent, bot_id):
        entry_var_key = tk.StringVar(value=0.1)
        entry_var_click = tk.StringVar(value=0.1)
        self.keys = [] #список кнопок, которые будут прожиматься
        self.bot_id = bot_id
        self.parrent = parrent
        self.key_interval = 0.5
        self.click_interval = 0.5
        self.delete_tab_btn = tk.Button(self.parrent.frames[-1], bg='pink', text="удалить вкладку", command=self.parrent.close_tab)
        self.delete_tab_btn.pack(fill='both')
        # создаем поле для ввода интервала между нажатиями клавиш
        self.key_interval_label = tk.Label(
            self.parrent.frames[-1], text="Задержка между нажатиями клавиш (в секундах):")
        self.key_interval_label.pack()
        self.key_interval_entry = tk.Entry(self.parrent.frames[-1], textvariable=entry_var_key, justify="center")
        self.key_interval_entry.pack()

        # создаем поле для ввода интервала между кликами правой кнопкой мыши
        self.click_interval_label = tk.Label(
            self.parrent.frames[-1], text="Задержка между нажатиями мыши (в секундах):")
        self.click_interval_label.pack()
        self.click_interval_entry = tk.Entry(self.parrent.frames[-1], textvariable=entry_var_click, justify="center")
        self.click_interval_entry.pack()

        # создаем поле для ввода последовательности клавиш для каждого окна
        self.buttons_frame = tk.Frame(self.parrent.frames[-1])
        self.buttons_frame.pack(fill='both')
        self.add_key_btn = tk.Button(self.buttons_frame, text="+нажатие клавиши+", command=self.add_action, width=20)
        self.add_key_btn.pack(fill='both', side='left')
        self.add_mouse_btn = tk.Button(self.buttons_frame, text="+клик мышкой+", command=self.add_mouse_action, width=20)
        self.add_mouse_btn.pack(fill='both')
        self.del_key_btn = tk.Button(self.parrent.frames[-1], text="-", width=3, command=self.del_action)
        self.del_key_btn.pack(side='right', fill='both')
        self.keys_list = tk.Listbox(self.parrent.frames[-1], height=5)
        self.keys_list.pack(fill='both')

    def add_mouse_action(self):
        self.open_mouse_modal()
        self.refresh_keys_list()

    def del_action(self):
        if self.keys_list.curselection() != ():
            del self.keys[self.keys_list.curselection()[0]]
        self.refresh_keys_list()

    def add_action(self):
        '''добавление кнопки в список (через нажатие и модальное окно)'''
        self.open_modal()
        self.refresh_keys_list()

    def refresh_keys_list(self):
        self.keys_list.delete(0, self.keys_list.size())
        for i in self.keys:
            if type(i) == str:
                self.keys_list.insert(self.keys_list.size(), "press key: [%s]"%i)
            elif type(i) == list:
                btn = ''
                if i[0] == 1:
                    btn = 'left'
                elif i[0] == 2:
                    btn = 'middle'
                elif i[0] == 3:
                    btn = 'right'
                self.keys_list.insert(self.keys_list.size(), "mouse click: btn=[%s], x=[%s], y=[%s]"%(btn, i[1], i[2]))
        for i in range(self.keys_list.size()):
            if i%2 == 0:
                self.keys_list.itemconfig(i, background='lightgray')

    def add_button(self, event):
        #добавление нажатой кнопки
        key_mapping = {
            'Return': 'enter',
            'Escape': 'escape',
            'BackSpace': 'backspace',
            'Tab': 'tab',
            'Delete': 'delete',
            'comma': '.',
            'period': ',',
            'slash': '/',
        }

        keysym = event.keysym
        if keysym in key_mapping:
            keysym = key_mapping[keysym]
        elif keysym == '??':
            return
        self.keys.append(keysym)
        self.close_modal_window()


    def open_modal(self):
        '''модальное окно для добавления кнопки'''
        self.modal_window = tk.Toplevel(self.parrent.root)
        width = 320
        height = 120
        self.modal_window.bind('<Key>', self.add_button)
        self.modal_window.title("Добавление клавиши")
        self.label = tk.Label(self.modal_window, text='нажмите клавишу', font=('system', 22))
        self.label.pack(anchor="center")
        self.modal_window.transient(self.parrent.root)
        self.modal_window.focus()
        self.modal_window.grab_set()

        # Размещение окна по центру родительского окна
        x = self.parrent.root.winfo_x() + (self.parrent.root.winfo_width()//2 - width//2)
        y = self.parrent.root.winfo_y() + (self.parrent.root.winfo_height() - self.modal_window.winfo_height()) // 2
        self.modal_window.geometry("{}x{}+{}+{}".format(width, height, x, y))

        self.modal_window.update_idletasks()
        self.modal_window.wait_window()

    def open_mouse_modal(self):
        self.modal_window = tk.Toplevel(self.parrent.root)
        self.modal_window.attributes('-fullscreen', True)
        self.modal_window.attributes('-alpha', 0.5)
        self.modal_window.bind('<Key>', self.close_modal_window)
        self.modal_window.bind('<Button>', self.add_mouse)
        self.modal_window.title("Добавление клика мыши")
        self.label = tk.Label(self.modal_window, text='кликните в нужное Вам место на экране', font=('system', 22))
        self.label.pack(anchor="center")
        self.modal_window.transient(self.parrent.root)
        self.modal_window.focus()
        self.modal_window.grab_set()
        self.modal_window.update_idletasks()
        self.modal_window.wait_window()

    def close_modal_window(self, *args):
        self.modal_window.destroy()

    def add_mouse(self, event):
        x=event.x
        y=event.y
        button=event.num
        self.keys.append([button, x, y])
        self.close_modal_window()

    def step(self):
        # получаем значения всех полей
        try:
            self.key_interval = float(self.key_interval_entry.get())
            self.click_interval = float(self.click_interval_entry.get())
        except:
            self.key_interval = 0.5
            self.click_interval = 0.5

    def actions_thread(self):
        #поток действий бота
        while True:
            # вводим заданную последовательность клавиш
            for key in self.keys:
                while self.parrent.window_in_focus == True:
                    time.sleep(0.03)
                self.parrent.window_in_focus = True
                self.parrent.main_windows[self.bot_id].set_focus()

                if type(key) == list:
                    x=key[1]
                    y=key[2]
                    # кликаем
                    time.sleep(self.click_interval)
                    if key[0] == 1:
                        mouse.click(coords=(x, y), button='left')
                    elif key[0] == 2:
                        mouse.click(coords=(x, y), button='middle')
                    elif key[0] == 3:
                        mouse.click(coords=(x, y), button='right')
                else:
                    pyautogui.press(key)
                self.parrent.window_in_focus = False
                time.sleep(self.key_interval)
            

# создаем графический интерфейс
root = tk.Tk()
auto_clicker = AutoClicker(root)
root.mainloop()