import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import time
import os
import sys
from pywinauto import Application
from pywinauto import mouse
import pyautogui
import win32gui
import threading
import keyboard

STOP_KEY = 'end' #кнопка для останова бота

class AutoClicker:
    def __init__(self, root, exe_path=None, bots_data=None):
        self.root = root
        self.root.title("AutoBot")
        #пытаюсь установить иконку
        try:
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
            self.root.wm_iconbitmap(os.path.join(base_path, "main.ico"))
        except:
            pass
        self.exe_path = exe_path
        self.bots_data = bots_data 
        pyautogui.FAILSAFE = False
        if self.exe_path != None:
            exe_name = os.path.basename(self.exe_path)
            self.exe_file_path = self.exe_path
        else:
            exe_name = "Выбрать exe-файл"
            self.exe_file_path = None
        self.window_in_focus = False #флаг, что окно в фокусе
        # создаем кнопку для выбора exe-файла
        self.exe_file_button = tk.Button(
            self.root, text=exe_name, bg='orange', command=self.choose_exe_file, font=('Times', 12))
        self.exe_file_button.pack(fill='both')

        #создаём панель вкладок
        self.notebook_style = ttk.Style()
        self.notebook_style.theme_create('rounded', parent='alt', settings={
            # Задаем свойство border-radius для закругления углов
            "TNotebook": {"configure": {"tabmargins": [2, 5, 2, 0], "border-radius": 10, "background": "black"}},
            "TNotebook.Tab": {"configure": {"padding": [20, 5], "background": "#fff", "foreground": "#333"}},
            "TNotebook.Tab": {"map": {"background": [("selected", "orange")], "foreground": [("selected", "#333")]}},
        })
        self.notebook_style.theme_use('rounded')
        self.notebook = ttk.Notebook() #панель вкладок
        self.notebook.pack()
        #
        self.frames = [] #фреймы вкладок
        self.bots = [] #объекты ботов
        if self.bots_data != None:
            try:
                for i in range(len(self.bots_data)):
                    self.add_tab()
                    #меняем значения задержек
                    self.bots[-1].entry_var_key.set(float(self.bots_data[i][0]))
                    self.bots[-1].entry_var_click.set(float(self.bots_data[i][1]))
                    #прописываем кнопки и клики
                    for key in self.bots_data[i][2]:
                        if "mouse=" in key:
                            mouse = key[6:].split()
                            mouse = [int(mouse[0]), int(mouse[1]), int(mouse[2])]
                            self.bots[-1].keys.append(mouse)
                        else:
                            self.bots[-1].keys.append(key)
                    self.bots[-1].refresh_keys_list()
            except:
                for i in range(4):
                    self.add_tab()
        else:
            for i in range(4):
                self.add_tab()

        # создаем кнопку для запуска окон приложений
        self.start_button = tk.Button(
            self.root, text="открыть окна приложений", bg='orange',command=self.run_apps, font=('Times', 12))
        self.start_button.pack(fill='both')
        if self.exe_file_path == None:
            self.start_button.config(state='disabled')

        #поток отслеживания нажатия клавиш
        thread = threading.Thread(target=self.key_checker_thread)
        thread.daemon = True
        thread.start()

    def save_state(self):
        '''метод сохраняет в conf.ini настройки'''
        try:
            file = open('conf.ini', 'w')
            file.write("exe=%s\n"%self.exe_file_path) #путь к exe-файлу
            for i in range(len(self.bots)):
                file.write("bot_tab kd=%s md=%s\n" %(self.bots[i].key_interval, self.bots[i].click_interval))
                for key in self.bots[i].keys:
                    if type(key) == list:
                        file.write("mouse=")
                        for n in key:
                            file.write("%s " %n)
                        file.write('\n')
                    else:
                        file.write("%s\n"%(key))
            file.close()
        except:
            pass

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
        self.frames.append(tk.Frame(bg='black', relief='flat', border=0))
        self.notebook.enable_traversal()
        self.notebook.add(self.frames[-1], text="окно %s " %str(panel_id+1))
        self.bots.append(Bot(self, bot_id=panel_id))

    def choose_exe_file(self):
        # вызываем диалоговое окно для выбора exe-файла
        self.exe_file_path = tk.filedialog.askopenfilename(filetypes=[("EXE files", "*.exe")])
        if self.exe_file_path != '' and self.exe_file_path != None:
            self.exe_file_button.configure(text=os.path.basename(self.exe_file_path))
            self.start_button.config(state='normal')

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
        #удаляю кнопку выбора EXE-файла
        self.exe_file_button.destroy()
        #обновляю названия вкладок
        for n in range(len(self.notebook.tabs())):
            self.notebook.tab(n, text="окно %s "%str(n+1))
            self.bots[n].bot_id = n
            #удаляю кнопки для управления вкладками
            self.bots[n].delete_tab_btn.destroy()
            self.bots[n].add_tab_btn.destroy()
            self.bots[n].stop_label = tk.Label(self.bots[n].tab_buttons_frame, text='для остановки бота нажмите клавишу: %s' %STOP_KEY, bg='black', fg='red', font=('Times', 12))
            self.bots[n].stop_label.pack(fill='both')

        # запускаем exe-файл нужное количество раз
        self.apps = [] #список запущенных экземпляров
        self.main_windows = [] #список с гл. окнами экземпляров
        windows_processes = []
        for i in range(len(self.bots)):
            self.apps.append(Application().start(self.exe_file_path))
            while True:
                # Получаем идентификатор окна, которое сейчас активно
                time.sleep(1)
                hwnd = win32gui.GetForegroundWindow()
                if win32gui.GetClassName(hwnd) != 'TkTopLevel':
                    if hwnd not in windows_processes:
                        windows_processes.append(hwnd)
                        # Получаем class_name этого окна
                        class_name = win32gui.GetClassName(hwnd)
                        self.main_windows.append(self.apps[-1].window(class_name=class_name))
                        break
                time.sleep(1)

        self.start_button.destroy()
        # создаем кнопку для запуска автокликера
        self.start_button = tk.Button(
            self.root, text="запустить бота!", bg='orange',command=self.start_threads, font=('Times', 12))
        self.start_button.pack(fill='both')

    def key_checker_thread(self):
        '''поток перехватывает клавишу, что
        бы остановить действия ботов'''
        keyboard.on_press(self.on_key_event)

    def on_key_event(self, event):
        '''смотрит, какая клавиша была нажата'''
        # print(f"Нажата клавиша {event.name}")
        #если нажата определённая клавиша - завершаю потоки ботов
        if event.name == STOP_KEY:
            self.root.deiconify()
            for i in self.bots:
                i.started = False

    def start_threads(self):
        #сохраняю настройки
        self.save_state()
        #запускаем потоки ботов   
        self.root.withdraw() #скрываем окно
        for i in range(len(self.bots)):
            self.bots[i].started = True
            thread = threading.Thread(target=self.bots[i].actions_thread)
            thread.daemon = True
            thread.start()

class Bot():
    def __init__(self, parrent, bot_id):
        self.started = False
        self.entry_var_key = tk.StringVar(value=0.1)
        self.entry_var_click = tk.StringVar(value=0.1)
        self.keys = [] #список кнопок, которые будут прожиматься
        self.bot_id = bot_id
        self.parrent = parrent
        self.key_interval = 0.5
        self.click_interval = 0.5
        self.tab_buttons_frame = tk.Frame(self.parrent.frames[-1], bg='black')
        self.tab_buttons_frame.pack(fill='both', ipady=2, ipadx=2)
        self.delete_tab_btn = tk.Button(self.tab_buttons_frame, bg='red', text=" x ", command=self.parrent.close_tab, font='system', relief='solid', border=1, width=3)
        self.delete_tab_btn.pack(side='right')
        # создаем кнопку для добавления вкладок
        self.add_tab_btn = tk.Button(self.tab_buttons_frame, text=" + ", bg='lightgreen', command=self.parrent.add_tab, font='system', relief='solid', border=1, width=3)
        self.add_tab_btn.pack(side='right')

        self.entrys_frame = tk.Frame(self.parrent.frames[-1], bg='black')
        self.entrys_frame.pack(fill='both', pady=10)
        # создаем поле для ввода интервала между нажатиями клавиш
        self.key_interval_label = tk.Label(
            self.entrys_frame, text="Задержка между нажатиями клавиш (в секундах):", bg='black', fg='white', font=('Times', 11))
        self.key_interval_label.pack()
        self.key_interval_entry = tk.Entry(self.entrys_frame, textvariable=self.entry_var_key, justify="center")
        self.key_interval_entry.pack(fill='both', padx=10)

        # создаем поле для ввода интервала между кликами правой кнопкой мыши
        self.click_interval_label = tk.Label(
            self.entrys_frame, text="Задержка между нажатиями мыши (в секундах):", bg='black', fg='white', font=('Times', 11))
        self.click_interval_label.pack()
        self.click_interval_entry = tk.Entry(self.entrys_frame, textvariable=self.entry_var_click, justify="center")
        self.click_interval_entry.pack(fill='both', padx=10)

        # создаем поле для ввода последовательности клавиш для каждого окна
        self.buttons_frame = tk.Frame(self.parrent.frames[-1], bg='black')
        self.buttons_frame.pack(fill='both')
        self.add_key_btn = tk.Button(self.buttons_frame, text="+нажатие клавиши+", command=self.add_action, width=20, bg='orange', font=('Courier', 10))
        self.add_key_btn.pack(fill='both', side='left')
        self.add_mouse_btn = tk.Button(self.buttons_frame, text="+клик мышкой+", command=self.add_mouse_action, width=20, bg='orange', font=('Courier', 10))
        self.add_mouse_btn.pack(fill='both')
        self.del_key_btn = tk.Button(self.parrent.frames[-1], text="у\nд\nа\nл\nи\nт\nь", width=3, command=self.del_action, bg='orange', font=('Courier', 10))
        self.del_key_btn.pack(side='right', fill='both')
        self.keys_list = tk.Listbox(self.parrent.frames[-1], height=8, bg='black', fg='orange', font='system', relief='flat', border=0)
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
                self.keys_list.itemconfig(i, background='darkgray', foreground='black')

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
            'Control_L': 'ctrlleft',
            'Down': 'down',
            'Up': 'up',
            'Left': 'left',
            'Right': 'right'
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
        while self.started == True:
            # вводим заданную последовательность клавиш
            for key in self.keys:
                if self.started == False:
                    self.parrent.window_in_focus = False
                    return False
                while self.parrent.window_in_focus == True:
                    if self.started == False:
                        self.parrent.window_in_focus = False
                        return False
                    time.sleep(0.03)
                self.parrent.window_in_focus = True
                try:
                    self.parrent.main_windows[self.bot_id].set_focus()
                except:
                    self.parrent.window_in_focus = False
                    return False

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
                    # keyboard.send(key) #---------------------------------------------------------------------------------------------------------------------------------------------------------------
                    # keyboard.press(key)

                self.parrent.window_in_focus = False
                time.sleep(self.key_interval)

def load_config():
    '''парсим файл конфигурации'''
    if os.path.exists('conf.ini'):
        file = open('conf.ini', 'r')
        config = file.read()
        file.close()
        bots = []
        config = config.split('\n')
        for i in config: 
            if i[:4] == 'exe=':
                exe_path = i[4:]
                if not os.path.exists(exe_path):
                    exe_path = None
            elif i[:8] == 'bot_tab ':
                kd = float(i[8:].split()[0][3:])
                md = float(i[8:].split()[1][3:])
                bots.append([kd, md, []])
            else:
                key = i
                if key != '':
                    bots[-1][-1].append(key)
        return exe_path, bots
    else:
        return None, None
            

if __name__ == "__main__":

    #проверяем сохраненные конфиги
    exe_path, bots = load_config()
    # создаем графический интерфейс
    root = tk.Tk()
    auto_clicker = AutoClicker(root, exe_path, bots)
    root.mainloop()