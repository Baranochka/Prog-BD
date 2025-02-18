# Программа для работы с базой данных "Название" и последующего вывода в специальной форме в Exele
# P.S. Извините за кринж

from datetime import datetime
import subprocess
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import openpyxl
from PIL import Image, ImageTk
from configparser import ConfigParser
from backend import MSSQL
import os
from pathlib import Path
import sys
from docx import Document
import requests
from docx.shared import Pt

version = "v0.2" # Надо менять версию после каждого изменения
latest_version = None

data = []
# Создание объекта для работы с базой данных
db = None 
# Переменная для хранения строки таблицы, на которую нажали
row_click = 0

path = Path(__file__).parent
def get_resource_path(filename):
    """Возвращает корректный путь к файлу, работает и в `.exe`, и в обычном запуске Python"""
    if getattr(sys, '_MEIPASS', False):  # Проверяем, запущен ли скрипт в режиме PyInstaller
        return os.path.join(sys._MEIPASS, filename)
    p = os.path.join(path, f"config\\{filename}")
    return p


jpg_path = get_resource_path("image.jpg")
ico_path = get_resource_path("icon.ico")
excle_path = get_resource_path("sample.xlsx")
config_path = get_resource_path("config.ini")
doc_path = get_resource_path("sample.docx")

config = ConfigParser()
config.read(config_path)

DRIVER        = config["SQL Server"]["DRIVER"] 
SERVER_NAME   = config["SQL Server"]["SERVER_NAME"]   
DATABASE_NAME = config["SQL Server"]["DATABASE_NAME"] 
USERNAME      = None
PASSWORD      = None


# DRIVER = config["Test"]["DRIVER"]
# SERVER_NAME = config["Test"]["SERVER_NAME"]
# DATABASE_NAME = config["Test"]["DATABASE_NAME"]
# USERNAME      = None
# PASSWORD      = None


# DRIVER        = config["Test_Maks"]["DRIVER"] 
# SERVER_NAME   = config["Test_Maks"]["SERVER_NAME"]   
# DATABASE_NAME = config["Test_Maks"]["DATABASE_NAME"] 
# USERNAME      = None
# PASSWORD      = None




class WindowAuthorizationConnectionBD(tk.Toplevel):  
    def __init__(self, parent):
        super().__init__(parent)

        self.parent = parent
        self.attributes("-topmost", True)
        self.title("Авторизация")
        center_window(self, 300, 200)
        self.config(bg="white")
        self.iconbitmap(ico_path)
        
        self.frame_auth = tk.Frame(self, bg="white")
        self.frame_auth.pack(expand=True, fill=tk.BOTH)
        tk.Label(self.frame_auth, text="Авторизация", bg="white",font=("Arial", 12)).place(x=100, y=10)
        tk.Label(self.frame_auth, text="Логин", bg="white").place(x=130, y=40)
        self.entry_login = tk.Entry(self.frame_auth, width=30,borderwidth=2)
        self.entry_login.place(x=60, y=60)
        tk.Label(self.frame_auth, text="Пароль", bg="white").place(x=125, y=80)
        self.entry_password = tk.Entry(self.frame_auth, width=30 ,borderwidth=2)
        self.entry_password.place(x=60, y=100)
        self.entry_password.config(show="*") 

        self.frame_incorrect = tk.Frame(self, bg="white")
           
        tk.Label(self.frame_incorrect, text="Соединение не установлено", bg="white", fg="red").place(x=75, y=75)
        self.frame_incorrect.pack_forget()

        self.frame_correct = tk.Frame(self, bg="white")
           
        tk.Label(self.frame_correct, text="Соединение установлено", bg="white", fg="green").place(x=75, y=75)
        self.frame_correct.pack_forget()

        self.entry_login.bind("<FocusIn>", lambda e: self.entry_login.delete(0, tk.END))
        self.entry_password.bind("<FocusIn>", lambda e: self.entry_password.delete(0, tk.END))

        tk.Button(self.frame_auth, text="Подключиться", command=self.connect).place(x=100, y=130)
        # Перехватываем нажатие на крестик (закрытие окна)
        self.protocol("WM_DELETE_WINDOW", self.close_app)

    def close_app(self):
        # Завершает весь Tkinter
        self.parent.quit()  
        
    
    def connect(self):
        self.frame_auth.pack_forget()
        global DRIVER, SERVER_NAME, DATABASE_NAME, USERNAME, PASSWORD
        USERNAME = self.entry_login.get()
        PASSWORD = self.entry_password.get()
        global db
        db = MSSQL(DRIVER, SERVER_NAME, DATABASE_NAME, USERNAME, PASSWORD)
        print(db.is_connected)
        if db.is_connected == True:
            self.frame_correct.pack(expand=True, fill=tk.BOTH)
            self.after(1000, self.hide_correct_frame)
            
        else:
            self.frame_incorrect.pack(expand=True, fill=tk.BOTH)
            self.after(1000, self.hide_incorrect_frame)
    
    def hide_incorrect_frame(self):
        self.frame_incorrect.pack_forget()
        self.frame_auth.pack(expand=True, fill=tk.BOTH)
        self.entry_login.delete(0, tk.END)
        self.entry_password.delete(0, tk.END)
    
    def hide_correct_frame(self):
        
        self.destroy()
        self.parent.deiconify()

# Описание класса графической части программы
class WindowProgramm(tk.Tk):
    # Инициализация окна программы и откликов в программе
    def __init__(window):
        super().__init__()
        # Описание окна
        window.title("Программа для работы с БД")
        center_window(window, 1000, 600)
        window.config(bg="white")
        # Устанавливаем иконку (Только .ico!)
        window.iconbitmap(ico_path)
        # Загружаем изображение
        window.image = Image.open(jpg_path)
        window.image = window.image.resize((150, 100))  
        window.photo = ImageTk.PhotoImage(window.image)

        # Создаём Label с изображением
        tk.Label(window, image=window.photo, borderwidth=0, highlightthickness=0).pack(pady=20,anchor="w")


        """--------------------------------------------------ФИО---------------------------------------------------"""

        # Создание контейнера для ФИО
        window.frame_fio = tk.Frame(bg="white")
        window.frame_fio.pack(fill=tk.X)

        # Создание ярлыка и текстового поля для ввода фамилии
        tk.Label(master=window.frame_fio, text="Фамилия:", bg="white").grid(row=1, column=0, sticky="e")
        window.entry_familia = tk.Entry(master=window.frame_fio, width=50,borderwidth=2)
        window.entry_familia.grid(row=1, column=1)

        # Создание ярлыка и текстового поля для ввода имени
        tk.Label(master=window.frame_fio, text="Имя:", bg="white").grid(row=2, column=0, sticky="e")
        window.entry_name = tk.Entry(master=window.frame_fio, width=50,borderwidth=2)
        window.entry_name.grid(row=2, column=1)

        # Создание ярлыка и текстового поля для ввода отчества
        tk.Label(master=window.frame_fio, text="Отчество:", bg="white").grid(row=3, column=0, sticky="e")
        window.entry_otchestvo = tk.Entry(master=window.frame_fio, width=50,borderwidth=2)
        window.entry_otchestvo.grid(row=3, column=1)

        """---------------------------------------------------ДАТА--------------------------------------------------"""

        # Создание контейнера для даты
        window.frame_data = tk.Frame(window,bg="white")
        window.frame_data.pack(fill=tk.X)

        # Создание ярлыка даты рождения
        tk.Label(window.frame_data, text="Дата рождения:",bg="white").grid(row=0, column=0, padx=5)

        # Переменные для хранения значений даты рождения
        window.day_var = tk.StringVar(value="1")
        window.month_var = tk.StringVar(value="1")
        window.year_var = tk.StringVar(value="1900")
        
        # Spinbox для дня от 1 до 31
        window.day_spinbox = tk.Spinbox(window.frame_data, from_=1, to=31, width=5, textvariable=window.day_var)
        window.day_spinbox.grid(row=0, column=1, padx=5)
        
        # Spinbox для месяца от 1 до 12
        window.month_spinbox = tk.Spinbox(window.frame_data, from_=1, to=12, width=5, textvariable=window.month_var)
        window.month_spinbox.grid(row=0, column=2, padx=5)
        
        # Spinbox для года от 1900 до 2025
        window.year_spinbox = tk.Spinbox(window.frame_data, from_=1900, to=2025, width=7, textvariable=window.year_var)
        window.year_spinbox.grid(row=0, column=3, padx=5)

        """------------------------------------------------------ГАЛОЧКИ-----------------------------------------------------------"""

        # # Переменные для хранения значений состояния галочек
        # window.var_check_fam = tk.BooleanVar()
        # window.var_check_name = tk.BooleanVar()
        # window.var_check_otch = tk.BooleanVar()
        # window.var_check_data = tk.BooleanVar()
        # window.var_check_all = tk.BooleanVar()

        # # Создание галочек
        # tk.Checkbutton(window.frame_fio, variable=window.var_check_fam, bg="white").grid(row=1, column=2, padx=5)
        # tk.Checkbutton(window.frame_fio, variable=window.var_check_name, bg="white").grid(row=2, column=2, padx=5)
        # tk.Checkbutton(window.frame_fio, variable=window.var_check_otch, bg="white").grid(row=3, column=2, padx=5)
        # tk.Checkbutton(window.frame_data, variable=window.var_check_data, bg="white").grid(row=0, column=4, padx=5)
        
        # # Создание контейнера для галочки ALL
        # window.frame_all = tk.Frame(bg="white")
        # window.frame_all.pack(fill=tk.X)

        # # Создание кнопки "Выделить всё"
        # tk.Checkbutton(window.frame_all,text="Выделить всё", variable=window.var_check_all, bg="white").pack(side="left")
        # window.var_check_all.trace_add("write", window.sync_checkboxes)
        
        """---------------------------------------------------КНОПКА-ПОИСКА--------------------------------------------------------------"""

        # Создание кнопки "Поиск"
        window.search_button = tk.Button(window, text="Поиск", command=window.click_find)
        window.search_button.pack(pady=10)

        """------------------------------------------------------ТАБЛИЦА-------------------------------------------------------------"""

        # Создание контейнера для таблицы (изначально скрыт)
        window.frame_table = tk.Frame(bg="white")
        
        # Отключение видимости таблицы
        window.frame_table.pack_forget()  

        # Первая строка с названиями столбцов
        window.columns = ( "Фамилия",                               #0      fru
                          "Familia",                                #1      last_lat
                          "Имя",                                    #2      name_rus
                          "Imya"                                    #3      nla
                        #   "Отчество",                               #4      och
                        #   "Otchestvo",                              #5      oche
                        #   "Гражданство",                            #6      pob
                        #   "Дата рождения",                          #7      dob
                        #   "Пол",                                    #8      sex
                        #   "Государство рождения",                   #9      pob
                        #   "Город рождения",                         #10     cob
                        #   "Серия паспорта",                         #11     pas_ser
                        #   "Номер паспорта",                         #12     pas_num
                        #   "Дата выдачи паспорта",                   #13     pds
                        #   "Срок действия паспорта",                 #14     pde
                        #   "Признак наличия визы",                   #15     visa_priz
                        #   "Виза серия",                             #16     vis_ser
                        #   "Виза номер",                             #17     vis_num
                        #   "Дата выдачи визы",                       #18     d_poluch
                        #   "Срок действия визы",                     #19     vis_end
                        #   "Телефон",                                #20     tel_nom
                        #   "Дата въезда",                            #21     d_enter
                        #   "Срок пребывания до",                     #22     date_okon
                        #   "Миграционная карта серия",               #23     mcs
                        #   "Миграционная карта номер",               #24     mcn
                        #   "Номер корпуса общежития",                #25     obsch
                        #   "Дата начала договора с общагой",         #26     dnd
                        #   "Номер договора",                         #27     dog_obsh
                        #   "Наличие ВНЖ, РВПО",                      #28     rf
                        #   "Дата выдачи ВНЖ, РВПО",                  #29     rfd
                        #   "Срок действия ВНЖ, РВПО",                #30     mot
                        #   "Серия ВНЖ, РВПО",                        #31     ser
                        #   "Номер ВНЖ, РВПО",                        #32     nmr 
                        #    "Кратность визы",                        #33     vis_krat
                        #    "Индентификатор визы",                   #34     vis_id
                        #    "Направление",                           #35     gos_nap
                        #    "Дата выдачи контракта",                 #36     gos_start
                        #    "Номер контракта",                       #37     kontrakt
                        #    "Срок обучения с",                       #38     kont_start
                        #    "Срок обучения по",                      #39     kont_end
        )               
        
        # Создание таблицы
        window.tree = ttk.Treeview(window.frame_table, columns=window.columns, show='headings')

        # Определение размера столбцов
        for col in window.columns :
            window.tree.heading(col, text=col)
            window.tree.column(col, anchor="center")
        

        # Создание пролистывания по таблице
        scrollbar_y = tk.Scrollbar(window.frame_table, orient="vertical", command=window.tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        window.tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = tk.Scrollbar(window.frame_table, orient="horizontal", command=window.tree.xview)
        scrollbar_x.pack(side="bottom", fill="x")
        window.tree.configure(xscrollcommand=scrollbar_x.set)
        
        window.tree.pack(expand=True, fill=tk.BOTH)



        """--------------------------------------------------------ОШИБКА-ОТСУТСТВИЯ-ДАННЫХ-----------------------------------------------------------"""

        # Создание контейнера для ошибки при отсутствии вводных данных
        window.frame_false = tk.Frame(bg="white")
        
        # Отключение видимости ошибки
        window.frame_false.pack_forget() 
        
        # Создание ярлыка отсутствия галочек
        tk.Label(master=window.frame_false, text="Введите данные", bg="white",fg="red").pack(side="top")

        """--------------------------------------------------------ОШИБКА-ОТСУТСТВИЯ-СТУДЕНТОВ-----------------------------------------------------------"""

        # Создание контейнера для ошибки при отсутствии студентов
        window.frame_false_stud = tk.Frame(bg="white")

        # Отключение видимости ошибки
        window.frame_false_stud.pack_forget()

        # Создание ярлыка отсутствия галочек
        tk.Label(master=window.frame_false_stud, text="Студент(ы) не найден(ы)", bg="white", fg="red").pack(side="top")

        """-------------------------------------------------------ОТКЛИК-НА-ТАБЛИЦУ----------------------------------------------------------"""

        # Возможность отклика на таблицу
        window.tree.bind("<Double-1>", window.click_on_table)



    # Реализация кнопки поика
    def click_find(window):
        
        if window.entry_familia.get() == "": 
            if window.entry_name.get() == "":
                if window.entry_otchestvo.get() == "":
                    if window.day_var.get() == "1":
                        if window.month_var.get() == "1":
                            if window.year_var.get() == "1900":
                                
                                window.frame_table.pack_forget() 
                                window.frame_false.pack(expand=True, fill=tk.BOTH)
                                return

        window.frame_false_stud.pack_forget()
        window.frame_false.pack_forget()
        
        window.find()    
        
    # Функция поиска
    def find(window):
        window.tree.delete(*window.tree.get_children())


        surname = window.entry_familia.get().upper() if window.entry_familia.get() else None
        name = window.entry_name.get().upper() if window.entry_name.get() else None
        patronymic = window.entry_otchestvo.get().upper() if window.entry_otchestvo.get() else None
        day = window.day_var.get()
        month = window.month_var.get()
        year = window.year_var.get()
        birthdate = datetime(int(year), int(month), int(day)) # datetime(1990 01 02)
        global data
        global db

        data = db.get_person(surname, name, patronymic, birthdate)

        if not data:
            window.frame_table.pack_forget()
            window.frame_false_stud.pack(expand=True, fill=tk.BOTH)
        else:
            for i, var in enumerate(data):
                window.tree.insert("", tk.END, iid=f"I00{i}", values=var)
            window.frame_table.pack(expand=True, fill=tk.BOTH)
            
    # Реализация изменения в таблице     
    def click_on_table(window, event):
        row = window.tree.identify_row(event.y)
        global row_click
        row_click = int(row[1:]) #I000
        WindowInformation(window)
        #print(row_click)
        # # Определение координат строки для изменения
        # item = window.tree.selection()[0]
        # col = window.tree.identify_column(event.x)
        
        # col_index = int(col[1:]) - 1
        # x, y, width, height = window.tree.bbox(item, col_index)

        # # Создание строки изменения
        # entry = tk.Entry(window.tree)
        # entry.place(x=x, y=y, width=width, height=height)
        # entry.insert(0, window.tree.item(item, "values")[col_index])
        # entry.focus()

        # # Функция обработки нажатия на Enter
        # def on_enter(event):
        #     new_value = entry.get()
        #     values = list(window.tree.item(item, "values"))
        #     values[col_index] = new_value
        #     data[0][col_index] = entry.get()
        #     window.tree.item(item, values=values)
        #     entry.destroy()

        # # Обработки нажатия на Enter
        # entry.bind("<Return>", on_enter)
        # # Обработка нажатия вне поля
        # entry.bind("<FocusOut>", lambda e: entry.destroy())

    
    # Функция изменяет состояние галочек при нажатии на 'Выделить всё'
    def sync_checkboxes(window, *args): 
        window.var_check_fam.set(window.var_check_all.get())
        window.var_check_name.set(window.var_check_all.get())
        window.var_check_otch.set(window.var_check_all.get())
        window.var_check_data.set(window.var_check_all.get())


class WindowInformation(tk.Toplevel):

    def __init__(window_inf, parent):
        super().__init__(parent)
        # Описание окна
        window_inf.title("Информация о студенте")
        center_window(window_inf, 900, 600)
        window_inf.config(bg="white")
        window_inf.iconbitmap(ico_path)
        bg_blue = "#6699CC"
        wh_color = "white"
        bl_color = "black"
        RoundedFrame(window_inf, width=890, height=90, radius=30, color="#6699CC").place(x=5, y=5)
        window_inf.do_lable( "Фамилия:", 10, 10, bg_blue,wh_color)
        window_inf.do_lable_value( data[row_click][0], 100, 9) if data[row_click][0] else window_inf.do_lable_value(  "Не указано", 100, 9)
        window_inf.do_lable( "Имя:", 10, 40, bg_blue, wh_color)
        window_inf.do_lable_value( data[row_click][2], 100, 39) if data[row_click][2] else window_inf.do_lable_value(  "Не указано", 100, 39)
        window_inf.do_lable(  "Отчество:", 10, 70, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][4], 100, 69) if data[row_click][4] else window_inf.do_lable_value(  "Не указано", 100, 69)
        
        window_inf.do_lable(  "Фамилия (лат):", 500, 10, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][1], 600, 9) if data[row_click][1] else window_inf.do_lable_value(  "Не указано", 600, 9)
        window_inf.do_lable(  "Имя (лат):", 500, 40, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][3], 600, 39) if data[row_click][3] else window_inf.do_lable_value(  "Не указано", 600, 39)
        window_inf.do_lable(  "Отчество (лат):", 500, 70, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][5], 600, 69) if data[row_click][5] else window_inf.do_lable_value(  "Не указано", 600, 69)
      
        window_inf.do_lable(  "Пол:", 10, 100, wh_color, bl_color)
        window_inf.do_lable_value(  data[row_click][8], 45, 99) if data[row_click][8] else window_inf.do_lable_value(  "Не указано", 45, 99)  
        window_inf.do_lable(  "Дата рождения:", 110, 100, wh_color, bl_color)
        window_inf.do_lable_value(  data[row_click][7], 205, 99) if data[row_click][7] else window_inf.do_lable_value(  "Не указано", 205, 99)        
        window_inf.do_lable(  "Телефон:", 270, 100, wh_color, bl_color)
        window_inf.do_lable_value(  f"+7{data[row_click][20]}", 330, 99) if data[row_click][20] else window_inf.do_lable_value(  "Не указано", 330, 99)

        RoundedFrame(window_inf, width=400, height=95, radius=30, color=bg_blue).place(x=5, y=125)
        window_inf.do_lable(  "Гражданство:", 10, 130, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][6], 95, 129) if data[row_click][6] else window_inf.do_lable_value(  "Не указано", 95, 129)
        window_inf.do_lable(  "Государство рождения:", 10, 160, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][9], 150, 159) if data[row_click][9] else window_inf.do_lable_value(  "Не указано", 150, 159)
        window_inf.do_lable(  "Город рождения:", 10, 190, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][10], 120, 189) if data[row_click][10] else window_inf.do_lable_value(  "Не указано", 120, 189)
 
        RoundedFrame(window_inf, width=400, height=90, radius=30, color=bg_blue).place(x=5, y=230)
        tk.Label(master=window_inf, text="Паспортные данные", bg=bg_blue, fg=wh_color, font=("Arial", 12)).place(x=10, y=235)
        
        window_inf.do_lable(  "Серия", 10, 265, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][11], 13, 285) if data[row_click][11] else window_inf.do_lable_value(  "Не указано", 13, 285)
        window_inf.do_lable(  "Номер", 110, 265, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][12], 113, 285) if data[row_click][12] else window_inf.do_lable_value(  "Не указано", 113, 285)
        window_inf.do_lable(  "Дата выдачи", 210, 265, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][13], 213, 285) if data[row_click][13] else window_inf.do_lable_value(  "Не указано", 213, 285)
        window_inf.do_lable(  "Срок действия", 300, 265, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][14], 303, 285) if data[row_click][14] else window_inf.do_lable_value(  "Не указано", 303, 285)

        if data[row_click][28] == "нет":
            RoundedFrame(window_inf, width=400, height=90, radius=30, color=bg_blue).place(x=425, y=230)
            tk.Label(master=window_inf, text="Виза", bg=bg_blue, fg=wh_color, font=("Arial", 12)).place(x=430, y=235)
            window_inf.do_lable(  "Серия", 430, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][16], 433, 285) if data[row_click][16] else window_inf.do_lable_value(  "Не указано", 433, 285)
            window_inf.do_lable(  "Номер", 530, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][17], 533, 285) if data[row_click][17] else window_inf.do_lable_value(  "Не указано", 533, 285)
            window_inf.do_lable(  "Дата выдачи", 630, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][18], 633, 285) if data[row_click][18] else window_inf.do_lable_value(  "Не указано", 633, 285)
            window_inf.do_lable(  "Срок действия", 730, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][19], 733, 285) if data[row_click][19] else window_inf.do_lable_value(  "Не указано", 733, 285)
        elif data[row_click][28] == "ВНЖ":
            RoundedFrame(window_inf, width=400, height=90, radius=30, color=bg_blue).place(x=425, y=230)
            tk.Label(master=window_inf, text="ВНЖ", bg="white", font=("Arial", 12)).place(x=430, y=235)
            window_inf.do_lable(  "Серия", 430, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][31], 433, 285) if data[row_click][31] else window_inf.do_lable_value(  "Не указано", 433, 285)
            window_inf.do_lable(  "Номер", 530, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][32], 533, 285) if data[row_click][32] else window_inf.do_lable_value(  "Не указано", 533, 285)
            window_inf.do_lable(  "Дата выдачи", 630, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][29], 633, 285) if data[row_click][29] else window_inf.do_lable_value(  "Не указано", 633, 285)
            window_inf.do_lable(  "Срок действия", 730, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][30], 733, 285) if data[row_click][30] else window_inf.do_lable_value(  "Не указано", 733, 285)
        elif data[row_click][28] == "РВПО":
            RoundedFrame(window_inf, width=400, height=90, radius=30, color=bg_blue).place(x=425, y=230)
            tk.Label(master=window_inf, text="РВПО", bg="white", font=("Arial", 12)).place(x=430, y=235)
            window_inf.do_lable(  "Серия", 430, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][31], 433, 285) if data[row_click][31] else window_inf.do_lable_value(  "Не указано", 433, 285)
            window_inf.do_lable(  "Номер", 530, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][32], 533, 285) if data[row_click][32] else window_inf.do_lable_value(  "Не указано", 533, 285)
            window_inf.do_lable(  "Дата выдачи", 630, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][29], 633, 285) if data[row_click][29] else window_inf.do_lable_value(  "Не указано", 633, 285)
            window_inf.do_lable(  "Срок действия", 730, 265, bg_blue, wh_color)
            window_inf.do_lable_value(  data[row_click][30], 733, 285) if data[row_click][30] else window_inf.do_lable_value(  "Не указано", 733, 285)

        window_inf.do_lable(  "Дата въезда", 10, 325, wh_color, bl_color)
        window_inf.do_lable_value(  data[row_click][21], 13, 345) if data[row_click][21] else window_inf.do_lable_value(  "Не указано", 13, 345)
        window_inf.do_lable(  "Срок пребывания до", 110, 325, wh_color, bl_color)
        window_inf.do_lable_value(  data[row_click][22], 139, 345) if data[row_click][22] else window_inf.do_lable_value(  "Не указано", 139, 345)
        
        
        RoundedFrame(window_inf, width=400, height=90, radius=30, color=bg_blue).place(x=5, y=375)
        tk.Label(master=window_inf, text="Миграционная карта", bg=bg_blue, fg=wh_color, font=("Arial", 12)).place(x=10, y=380)
        window_inf.do_lable(  "Серия", 10, 410, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][23], 13, 430) if data[row_click][23] else window_inf.do_lable_value(  "Не указано", 13, 410)
        window_inf.do_lable(  "Номер", 100, 410, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][24], 103, 430) if data[row_click][24] else window_inf.do_lable_value(  "Не указано", 103, 410)
        
        RoundedFrame(window_inf, width=400, height=90, radius=30, color=bg_blue).place(x=425, y=375)
        tk.Label(master=window_inf, text="Общежитие", bg=bg_blue, fg=wh_color, font=("Arial", 12)).place(x=430, y=380)
        num = data[row_click][25]
        window_inf.do_lable(  "Номер корпуса", 435, 410, bg_blue, wh_color)
        window_inf.do_lable_value( num[0], 470, 430) if data[row_click][25] else window_inf.do_lable_value(  "Не указано", 470, 430)
        window_inf.do_lable(  "Дата начала договора", 535, 410, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][26], 565, 430) if data[row_click][26] else window_inf.do_lable_value(  "Не указано", 565, 430)
        window_inf.do_lable(  "Номер договора", 670, 410, bg_blue, wh_color)
        window_inf.do_lable_value(  data[row_click][27], 673, 430) if data[row_click][27] else window_inf.do_lable_value(  "Не указано", 673, 430)

        # Создание кнопки "Сформировать"
        tk.Button(window_inf, text="Сформировать в Excel",bg=bg_blue,fg=wh_color, command=window_inf.click_form_excel).place(x=750, y=500)
        tk.Button(window_inf, text="Сформировать уведомление",bg=bg_blue,fg=wh_color, command=window_inf.click_form_Word).place(x=550, y=500)
        
        # Фиксация окна
        window_inf.grab_set()  
        window_inf.focus_set()  
        window_inf.wait_window()

    def do_lable(window_inf, text, x, y, bg_color, fg_color):
        tk.Label(master=window_inf, text=text, bg=bg_color, fg=fg_color).place(x=x, y=y)
    def do_lable_value(window_inf, text, x, y):
        tk.Label(master=window_inf, text=text, bg="white", relief="ridge" , borderwidth=3).place(x=x, y=y)

    def click_form_excel(window_inf):
        WindowSaveExcel(window_inf)
        
    def click_form_Word(window_inf):
        WindowSaveWord(window_inf)


class WindowSaveExcel(tk.Toplevel):
    def __init__(window_save, parent): 
        super().__init__(parent)
        # Описание окна
        window_save.title("Сохранение в Exel")
        center_window(window_save, 300, 100)
        window_save.config(bg="white")
        window_save.iconbitmap(ico_path)
        # Создание ярлыка сохранения
        tk.Label(master=window_save, text="Сохранение", bg="white", font=("Arial", 12)).pack(side="top",pady=7)
        
        #Создание контейнера для названия файла
        window_save.frame_name_file = tk.Frame(window_save,bg="white")
        window_save.frame_name_file.pack(fill=tk.X)

        # Создание ярлыка и текстового поля для ввода названия файла
        tk.Label(master=window_save.frame_name_file, text="Название файла:", bg="white").grid(row=0,column=0)
        window_save.entry_name_file = tk.Entry(master=window_save.frame_name_file, width=25, relief="groove", borderwidth=5)
        window_save.entry_name_file.insert(0, f"{data[row_click][1]} {data[row_click][3]}")

        window_save.entry_name_file.grid(row=0,column=1)

        # Создание кнопки "Сохранить"
        window_save.search_button = tk.Button(window_save, text="Сохранить", command=window_save.click_save)
        window_save.search_button.pack(side="right", padx=5)

        # Обработка нажатия внутри области
        window_save.entry_name_file.bind("<FocusIn>", lambda e: window_save.entry_name_file.delete(0,len(window_save.entry_name_file.get())+1))

        # Фиксация окна
        window_save.grab_set()  
        window_save.focus_set()  
        window_save.wait_window()
    
    def click_save(window_save):
        CompletionExcel(window_save.entry_name_file.get())
        window_save.destroy()


class WindowSaveWord(tk.Toplevel):
    def __init__(window_save, parent): 
        super().__init__(parent)
        # Описание окна
        window_save.title("Сохранение в Exel")
        center_window(window_save, 500, 250)
        window_save.config(bg="white")
        window_save.iconbitmap(ico_path)
        # Создание ярлыка сохранения
        tk.Label(master=window_save, text="Сохранение", bg="white", font=("Arial", 12)).place(x=210,y=10)
        
        window_save.var_check_1 = tk.BooleanVar()
        window_save.var_check_2 = tk.BooleanVar()
        window_save.var_check_3 = tk.BooleanVar()
        
        # # Создание галочек
        tk.Checkbutton(window_save, variable=window_save.var_check_1, bg="white").place(x=10,y=50)
        tk.Label(master=window_save, text="о предоставлении иностранному гражданину (лицу без гражданства)\n" \
                "академического отпуска образовательной/научной организации;", bg="white").place(x=30, y=50)
        tk.Checkbutton(window_save, variable=window_save.var_check_2, bg="white").place(x=10,y=90)
        tk.Label(master=window_save, text="о досрочном прекращении обучения иностранного гражданина\n" \
                "(лица без гражданства) в образовательной/научной организации;", bg="white").place(x=30, y=90)
        tk.Checkbutton(window_save, variable=window_save.var_check_3, bg="white").place(x=10,y=130)
        tk.Label(master=window_save, text="о завершении обучения иностранного гражданина (лица без гражданства)\n" \
                "(лица без гражданства) в образовательной/научной организации;", bg="white").place(x=30, y=130)
               

        # Создание ярлыка и текстового поля для ввода названия файла
        tk.Label(master=window_save, text="Название файла:", bg="white").place(x=30,y=180)
        window_save.entry_name_file = tk.Entry(master=window_save, width=50, relief="groove", borderwidth=5)
        window_save.entry_name_file.insert(0, f"Уведомление {data[row_click][1]} {data[row_click][3]}")
        window_save.entry_name_file.place(x=30,y=200)

        # Создание кнопки "Сохранить"
        window_save.search_button = tk.Button(window_save, text="Сохранить", command=window_save.click_save)
        window_save.search_button.place(x=350,y=199)

        # Обработка нажатия внутри области
        window_save.entry_name_file.bind("<FocusIn>", lambda e: window_save.entry_name_file.delete(0,len(window_save.entry_name_file.get())+1))

        # Фиксация окна
        window_save.grab_set()  
        window_save.focus_set()  
        window_save.wait_window()
    
    def click_save(window_save):
        ComplectionWord(window_save.entry_name_file.get(), window_save)
        window_save.destroy()


class RoundedFrame(tk.Canvas):
    def __init__(self, parent, width, height, radius=25, color="lightblue"):
        super().__init__(parent, width=width, height=height, bg="white", highlightthickness=0)

        # Рисуем скругленный прямоугольник
        self.create_rounded_rect(0, 0, width, height, radius, fill=color, outline=color)

    def create_rounded_rect(self, x1, y1, x2, y2, radius, **kwargs):
        """Создаёт прямоугольник с закруглёнными углами"""
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1
        ]
        return self.create_polygon(points, smooth=True, **kwargs)

# Функция центрирования окна
def center_window(window, width, height):
    # Получаем размеры экрана
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Вычисляем координаты верхнего левого угла
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размеры и положение окна
    window.geometry(f"{width}x{height}+{x}+{y}")

# Функция Заполнения Excel-таблицы
def change_sheet(sheet, row, col, step, value, max_col):
    col_cur = col
    row_cur = row
    for val in value:
        if col_cur <= max_col:
            sheet.cell(row=row_cur, column=col_cur, value = val)
            col_cur += step

# Функция сохранения Excel-таблицы
def CompletionExcel(file_out):
    #Путь к файлу

    # Загрузка книги
    wb = openpyxl.load_workbook(excle_path)
    sheet = wb.active  # Работаем с активным листом (если нужен другой, уточните)

    # Начало заполнения таблицы
    # Фамилия на русском
    change_sheet(sheet, 10, 14, 4, data[row_click][0],  122)
    change_sheet(sheet, 176, 14, 4, data[row_click][0],  122)
    # Фамилия на латинице
    change_sheet(sheet, 12, 38, 4, data[row_click][1],  122)
    # Имя на русском
    change_sheet(sheet, 14, 14, 4, data[row_click][2],  122)
    change_sheet(sheet, 178, 14, 4, data[row_click][2],  122)
    # Имя на латинице
    change_sheet(sheet, 16, 34, 4, data[row_click][3],  122)
    # Отчество на русском
    change_sheet(sheet, 18, 30, 4, data[row_click][4],  122)
    change_sheet(sheet, 180, 34, 4, data[row_click][4],  122)
    # Отчество на латинице
    change_sheet(sheet, 20, 38, 4, data[row_click][5],  122)
    # Гражданство
    change_sheet(sheet, 24, 22, 4, data[row_click][6],  122)
    change_sheet(sheet, 182, 18, 4, data[row_click][6],  122)
    #Дата рождения
    date = data[row_click][7] 
    # Число рождения
    change_sheet(sheet, 27, 30, 0, date[0],  30)
    change_sheet(sheet, 27, 34, 0, date[1],  34)
    change_sheet(sheet, 185, 30, 0, date[0],  30)
    change_sheet(sheet, 185, 34, 0, date[1],  34)
    # Месяц рождения
    change_sheet(sheet, 27, 46, 0, date[3],  46)
    change_sheet(sheet, 27, 50, 0, date[4],  50)
    change_sheet(sheet, 185, 50, 0, date[3],  50)
    change_sheet(sheet, 185, 54, 0, date[4],  54)
    # Год рождения
    change_sheet(sheet, 27, 58, 0, date[6],  58)
    change_sheet(sheet, 27, 62, 0, date[7],  62)
    change_sheet(sheet, 27, 66, 0, date[8],  66)
    change_sheet(sheet, 27, 70, 0, date[9],  70)
    change_sheet(sheet, 185, 66, 0, date[6],  66)
    change_sheet(sheet, 185, 70, 0, date[7],  70)
    change_sheet(sheet, 185, 74, 0, date[8],  74)
    change_sheet(sheet, 185, 78, 0, date[9],  78)
    # Пол
    if data[row_click][8] == "Мужской":
        change_sheet(sheet, 27, 90, 0, "X",  90)
        change_sheet(sheet, 185, 102, 0, "X",  102)
    elif data[row_click][8] == "Женский":
        change_sheet(sheet, 27, 106, 0, "X",  106)
        change_sheet(sheet, 185, 118, 0, "X",  118)
    # Государство рождения
    change_sheet(sheet, 29, 26, 4, data[row_click][9],  122)
    # Город рождения
    change_sheet(sheet, 31, 2, 4, data[row_click][10], 122)
    # Серия паспорта
    change_sheet(sheet, 37, 10, 4, data[row_click][11],  30)
    change_sheet(sheet, 187, 58, 4, data[row_click][11],  78)
    # Номер паспорта
    change_sheet(sheet, 37, 42, 4, data[row_click][12],  122)
    change_sheet(sheet, 187, 86, 4, data[row_click][12],  122)
    # Дата выдачи паспорта
    date = data[row_click][13]
    # Число
    change_sheet(sheet, 39, 9, 0, date[0],  9)
    change_sheet(sheet, 39, 13, 0, date[1],  13)
    change_sheet(sheet, 189, 9, 0, date[0],  9)
    change_sheet(sheet, 189, 13, 0, date[1],  13)
    # Месяц выдачи
    change_sheet(sheet, 39, 26, 0, date[3],  26)
    change_sheet(sheet, 39, 30, 0, date[4],  30)
    change_sheet(sheet, 189, 26, 0, date[3],  26)
    change_sheet(sheet, 189, 30, 0, date[4],  30)
    # Год выдачи
    change_sheet(sheet, 39, 38, 0, date[6],  38)
    change_sheet(sheet, 39, 42, 0, date[7],  42)
    change_sheet(sheet, 39, 46, 0, date[8],  46)
    change_sheet(sheet, 39, 50, 0, date[9],  50)
    change_sheet(sheet, 189, 38, 0, date[6],  38)
    change_sheet(sheet, 189, 42, 0, date[7],  42)
    change_sheet(sheet, 189, 46, 0, date[8],  46)
    change_sheet(sheet, 189, 50, 0, date[9],  50)
    # Дата срока действия паспорта
    date = data[row_click][14]
    # Число 
    change_sheet(sheet, 39, 66, 0, date[0],  66)
    change_sheet(sheet, 39, 70, 0, date[1],  70)
    change_sheet(sheet, 189, 66, 0, date[0],  66)
    change_sheet(sheet, 189, 70, 0, date[1],  70)
    # Месяц 
    change_sheet(sheet, 39, 82, 0, date[3],  82)
    change_sheet(sheet, 39, 86, 0, date[4],  86)
    change_sheet(sheet, 189, 82, 0, date[3],  82)
    change_sheet(sheet, 189, 86, 0, date[4],  86)
    # Год 
    change_sheet(sheet, 39, 94, 0, date[6],  94)
    change_sheet(sheet, 39, 98, 0, date[7],  98)
    change_sheet(sheet, 39, 102, 0, date[8],  102)
    change_sheet(sheet, 39, 106, 0, date[9],  106)
    change_sheet(sheet, 189, 94, 0, date[6],  94)
    change_sheet(sheet, 189, 98, 0, date[7],  98)
    change_sheet(sheet, 189, 102, 0, date[8],  102)
    change_sheet(sheet, 189, 106, 0, date[9],  106)
    # Проверка на наличие визы, ВНЖ, РВПО
    if data[row_click][28] == "нет":
        # Наличие визы
        if data[row_click][15] == "1":
            change_sheet(sheet, 53, 8, 0, "X",  8)
        # Серия визы
        change_sheet(sheet, 56, 10, 4, data[row_click][16],  30)
        # Номер визы
        change_sheet(sheet, 56, 42, 4, data[row_click][17],  122)
        # Дата выдачи визы
        date = data[row_click][18]
        # Число 
        change_sheet(sheet, 58, 9, 0, date[0],  9)
        change_sheet(sheet, 58, 13, 0, date[1],  13)
        # Месяц 
        change_sheet(sheet, 58, 26, 0, date[3],  26)
        change_sheet(sheet, 58, 30, 0, date[4],  30)
        # Год 
        change_sheet(sheet, 58, 38, 0, date[6],  38)
        change_sheet(sheet, 58, 42, 0, date[7],  42)
        change_sheet(sheet, 58, 46, 0, date[8],  46)
        change_sheet(sheet, 58, 50, 0, date[9],  50)
        # Дата срока визы
        date=data[row_click][19]
        # Число 
        change_sheet(sheet, 58, 66, 0, date[0],  66)
        change_sheet(sheet, 58, 70, 0, date[1],  70)
        # Месяц 
        change_sheet(sheet, 58, 82, 0, date[3],  82)
        change_sheet(sheet, 58, 86, 0, date[4],  86)
        # Год 
        change_sheet(sheet, 58, 94, 0, date[6],  94)
        change_sheet(sheet, 58, 98, 0, date[7],  98)
        change_sheet(sheet, 58, 102, 0, date[8],  102)
        change_sheet(sheet, 58, 106, 0, date[9],  106)
    elif data[row_click][28] == "РВПО":
        change_sheet(sheet, 53, 108, 0, "X",  108)
        # Серия РВПО
        change_sheet(sheet, 56, 10, 4, data[row_click][31],  30)
        # Номер РВПО
        change_sheet(sheet, 56, 42, 4, data[row_click][32],  122)
        # Дата выдачи РВПО
        date = data[row_click][29]
        # Число
        change_sheet(sheet, 58, 9, 0, date[0],  9)
        change_sheet(sheet, 58, 13, 0, date[1],  13)
        # Месяц 
        change_sheet(sheet, 58, 26, 0, date[3],  26)
        change_sheet(sheet, 58, 30, 0, date[4],  30)
        # Год 
        change_sheet(sheet, 58, 38, 0, date[6],  38)
        change_sheet(sheet, 58, 42, 0, date[7],  42)
        change_sheet(sheet, 58, 46, 0, date[8],  46)
        change_sheet(sheet, 58, 50, 0, date[9],  50)
        # Дата срока РВПО
        date = data[row_click][30]
        # Число 
        change_sheet(sheet, 58, 66, 0, date[0],  66)
        change_sheet(sheet, 58, 70, 0, date[1],  70)
        # Месяц 
        change_sheet(sheet, 58, 82, 0, date[3],  82)
        change_sheet(sheet, 58, 86, 0, date[4],  86)
        # Год 
        change_sheet(sheet, 58, 94, 0, date[6],  94)
        change_sheet(sheet, 58, 98, 0, date[7],  98)
        change_sheet(sheet, 58, 102, 0, date[8],  102)
        change_sheet(sheet, 58, 106, 0, date[9],  106)
    elif data[row_click][28] == "ВНЖ":
        change_sheet(sheet, 53, 36, 0, "X",  36)
        # Серия ВНЖ
        change_sheet(sheet, 56, 10, 4, data[row_click][31],  30)
        # Номер ВНЖ
        change_sheet(sheet, 56, 42, 4, data[row_click][32],  122)
        # Дата выдачи ВНЖ
        date = data[row_click][29]
        # Число
        change_sheet(sheet, 58, 9, 0, date[0],  9)
        change_sheet(sheet, 58, 13, 0, date[1],  13)
        # Месяц 
        change_sheet(sheet, 58, 26, 0, date[3],  26)
        change_sheet(sheet, 58, 30, 0, date[4],  30)
        # Год 
        change_sheet(sheet, 58, 38, 0, date[6],  38)
        change_sheet(sheet, 58, 42, 0, date[7],  42)
        change_sheet(sheet, 58, 46, 0, date[8],  46)
        change_sheet(sheet, 58, 50, 0, date[9],  50)
        # Дата срока ВНЖ
        date = data[row_click][30]
        # Число 
        change_sheet(sheet, 58, 66, 0, date[0],  66)
        change_sheet(sheet, 58, 70, 0, date[1],  70)
        # Месяц 
        change_sheet(sheet, 58, 82, 0, date[3],  82)
        change_sheet(sheet, 58, 86, 0, date[4],  86)
        # Год 
        change_sheet(sheet, 58, 94, 0, date[6],  94)
        change_sheet(sheet, 58, 98, 0, date[7],  98)
        change_sheet(sheet, 58, 102, 0, date[8],  102)
        change_sheet(sheet, 58, 106, 0, date[9],  106)
    # Телефон
    # change_sheet(sheet, 62, 82, 4, data[row_click][20],  118)
    # Дата въезда
    date = data[row_click][21]
    # Число 
    change_sheet(sheet, 66, 9, 0, date[0],  9)
    change_sheet(sheet, 66, 13, 0, date[1],  13)
    # Месяц 
    change_sheet(sheet, 66, 26, 0, date[3],  26)
    change_sheet(sheet, 66, 30, 0, date[4],  30)
    # Год 
    change_sheet(sheet, 66, 38, 0, date[6],  38)
    change_sheet(sheet, 66, 42, 0, date[7],  42)
    change_sheet(sheet, 66, 46, 0, date[8],  46)
    change_sheet(sheet, 66, 50, 0, date[9],  50)
    # Дата срок пребывания
    date = data[row_click][22]
    # Число
    change_sheet(sheet, 66, 66, 0, date[0],  66)
    change_sheet(sheet, 66, 70, 0, date[1],  70)
    change_sheet(sheet, 222, 49, 0, date[0],  49)
    change_sheet(sheet, 222, 53, 0, date[1],  53)
    change_sheet(sheet, 266, 38, 0, date[0],  38)
    change_sheet(sheet, 266, 42, 0, date[1],  42)
    # Месяц 
    change_sheet(sheet, 66, 82, 0, date[3],  82)
    change_sheet(sheet, 66, 86, 0, date[4],  86)
    change_sheet(sheet, 222, 65, 0, date[3],  65)
    change_sheet(sheet, 222, 69, 0, date[4],  69)
    change_sheet(sheet, 266, 54, 0, date[3],  54)
    change_sheet(sheet, 266, 58, 0, date[4],  58)
    # Год 
    change_sheet(sheet, 66, 94, 0, date[6],  94)
    change_sheet(sheet, 66, 98, 0, date[7],  98)
    change_sheet(sheet, 66, 102, 0, date[8],  102)
    change_sheet(sheet, 66, 106, 0, date[9],  106)
    change_sheet(sheet, 222, 79, 0, date[6],  79)
    change_sheet(sheet, 222, 83, 0, date[7],  83)
    change_sheet(sheet, 222, 87, 0, date[8],  87)
    change_sheet(sheet, 222, 91, 0, date[9],  91)
    change_sheet(sheet, 266, 68, 0, date[6],  68)
    change_sheet(sheet, 266, 72, 0, date[7],  72)
    change_sheet(sheet, 266, 76, 0, date[8],  76)
    change_sheet(sheet, 266, 80, 0, date[9],  80)
    # Миграционная карта серия
    change_sheet(sheet, 68, 34, 4, data[row_click][23],  46)
    # Миграционная карта номер
    change_sheet(sheet, 68, 54, 4, data[row_click][24],  102)

    # Номер корпус
    if data[row_click][25]:
        num = data[row_click][25]
        sheet.cell(row=104, column=46, value = f"КОРПУС {num[0]}")
        sheet.cell(row=200, column=46, value = f"КОРПУС {num[0]}")
    # Дата начала найма
    date_naym = data[row_click][26]
    if date_naym != "":
        # Число найма
        change_sheet(sheet, 112, 90, 0, date_naym[0],  90)
        change_sheet(sheet, 112, 94, 0, date_naym[1],  94)
        # Месяц найма
        change_sheet(sheet, 112, 102, 0, date_naym[3],  102)
        change_sheet(sheet, 112, 106, 0, date_naym[4],  106)
        # Год найма
        change_sheet(sheet, 112, 114, 0, date_naym[8],  114)
        change_sheet(sheet, 112, 118, 0, date_naym[9],  118)
        # Номер договора
    
    change_sheet(sheet, 114, 46, 4, data[row_click][27],  90)

    # Сохранение изменений
    if os.path.isdir(".\\out"):
        pass
    else:
        os.mkdir(f".\\out")

    output_path = f".\\out\\{file_out}.xlsx"
    wb.save(output_path)
    os.startfile(output_path)


def ComplectionWord(file_out, window_save):
    print(doc_path)
    doc = Document(doc_path)
    if window_save.var_check_1.get() == True:
        UpdateWord(doc, 1, 12, 1, "X")
    if window_save.var_check_2.get() == True:
        UpdateWord(doc, 1, 14, 1, "X")
    if window_save.var_check_3.get() == True:
        UpdateWord(doc, 1, 16, 1, "X")
    UpdateWord(doc, 1, 19, 3, data[row_click][0]) # Фамилия на русском
    UpdateWord(doc, 1, 19, 15, data[row_click][1]) # Фамилия на латинице
    UpdateWord(doc, 1, 21, 4, data[row_click][2]) # Имя на русском
    UpdateWord(doc, 1, 21, 15, data[row_click][3]) # Имя на латинице
    UpdateWord(doc, 1, 23, 7, data[row_click][4]) # Отчество на русском
    UpdateWord(doc, 1, 23, 15, data[row_click][5]) # Отчество на латинице
    UpdateWord(doc, 1, 25, 5, data[row_click][7]) # Дата рождения
    UpdateWord(doc, 1, 25, 13, f"{data[row_click][9]}, {data[row_click][10]}", 11) # Стаана и город рождения
    UpdateWord(doc, 1, 27, 8, data[row_click][6]) # Гражданство
    if data[row_click][8] == "Мужской":
        UpdateWord(doc,1, 27, 20, "X") # Пол
    elif data[row_click][8] == "Женский":
        UpdateWord(doc,1, 27, 22, "X") # Пол
    UpdateWord(doc, 1, 29, 2, data[row_click][11]) # Серия паспорта
    UpdateWord(doc, 1, 29, 6, data[row_click][12]) # Номер паспорта
    UpdateWord(doc, 1, 29, 15, data[row_click][13], 11) # Дата выдачи паспорта
    UpdateWord(doc, 1, 29, 19, data[row_click][14], 11) # Дата срока действия паспорта
    if data[row_click][25]:
        num = data[row_click][25]
        UpdateWord(doc, 2, 2, 0, f"Г. МОСКВА, КОЧНОВСКИЙ ПР., Д.7, КОРПУС {num[0]}") 
    UpdateWord(doc, 2, 3, 14, data[row_click][21]) # Дата въезда
    UpdateWord(doc, 2, 3, 23, data[row_click][22]) # Срок пребывания
    if data[row_click][28] == "нет":
        UpdateWord(doc, 2, 6, 4, data[row_click][33], 12, False) # Кратност визы
        UpdateWord(doc, 2, 6, 28, "учебная",12, False) # Категория визы
        UpdateWord(doc, 2, 7, 3, "учеба", 12, False) # Цель визы
        UpdateWord(doc, 2, 7, 19, data[row_click][16]) # Серия визы
        UpdateWord(doc, 2, 7, 28, data[row_click][17]) # Номер визы
        UpdateWord(doc, 2, 8, 8, data[row_click][34]) # Инденитификационный номер визы
        UpdateWord(doc, 2, 8, 24, data[row_click][18]) # Дата выдачи визы
        UpdateWord(doc, 2, 8, 31, data[row_click][19]) # Дата срока визы
    if data[row_click][35] != None:
        UpdateWord(doc, 2, 12, 10, "гос.направление", 11, False) # Направление
        UpdateWord(doc, 2, 12, 21, data[row_click][36], 11) # Дата выдачи контракта
    UpdateWord(doc, 2, 12, 28, data[row_click][37], 11) # Номер контракта
    UpdateWord(doc, 2, 14, 5, data[row_click][38], 11) # Срок обучения с
    UpdateWord(doc, 2, 14, 15, data[row_click][39], 11) # Срок обучения по
    
    # Сохранение изменений
    if os.path.isdir(".\\out"):
        pass
    else:
        os.mkdir(f".\\out")

    output_path = f".\\out\\{file_out}.docx"
    doc.save(output_path)
    os.startfile(output_path)


def UpdateWord(doc, table_index, row_index, cell_index, new_text, font_size=12, uppercase= True):
    """ Обновляет текст в ячейке таблицы, добавляя жирность и изменение регистра """
    try:
        table = doc.tables[table_index]
        if row_index < len(table.rows) and cell_index < len(table.rows[row_index].cells):
            cell = table.rows[row_index].cells[cell_index]

            # Определяем окончательный текст
            formatted_text = new_text.upper() if uppercase else new_text.lower()

            # Если в ячейке нет параграфов, создаем новый
            if not cell.paragraphs:
                paragraph = cell.add_paragraph()
            else:
                paragraph = cell.paragraphs[0]

            # Если нет runs, создаем новый run
            if not paragraph.runs:
                run = paragraph.add_run(formatted_text)
            else:
                run = paragraph.runs[0]
                run.text = formatted_text

            # Применяем жирный стиль
            run.bold = True
            run.font.name = "Times New Roman"  # Устанавливаем шрифт
            run.font.size = Pt(font_size)  # Устанавливаем размер шрифта
        else:
            print(f"Invalid row or cell index: ({row_index}, {cell_index})")
    except IndexError as e:
        print(f"Error accessing table: {e}")       


def CheckUpdate():
    url = "https://api.github.com/repos/Baranochka/Prog-BD/releases/latest"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            latest_release = response.json()
            global latest_version
            latest_version = latest_release['tag_name']
            if latest_version != version:
                return True
            else:
                return False
    except:
        return False


def Update():
        url = f"https://github.com/Baranochka/Prog-BD/releases/download/{latest_version}/Program.BD.exe"
        cur_path_exe = os.path.dirname(sys.executable)
        save_path = os.path.join(cur_path_exe, 'update\\Program BD.exe')
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            if os.path.isdir(".\\update"):
                pass
            else:
                os.mkdir(f".\\update")
            with open(save_path, "wb") as file:
                for chunk in response.iter_content(1024):
                    file.write(chunk)
            if os.path.isfile(save_path):
                UpdateSelf(save_path, cur_path_exe)
        else:
            return False


def UpdateSelf(new_exe_path, cur_path_exe):
    
    
    updater_script = "updater.bat"
    
    with open(updater_script, "w") as f:
        f.write(f"""
        @echo off
        setlocal
        set EXE_NAME=Program BD.exe
        set UPDATE_FOLDER=.\\update\\
        set TARGET_FOLDER=.\\

        :: Убиваем процесс, если он запущен
        taskkill /F /IM %EXE_NAME% > nul 2>&1
        
        timeout /t 2 /nobreak > nul

        copy "%UPDATE_FOLDER%%EXE_NAME%" 

        rmdir /Q /S "%UPDATE_FOLDER%"

        :: start "" "%TARGET_FOLDER%%EXE_NAME%"

        cmd /c del "%~f0"
        """)
    try:
        subprocess.Popen(updater_script, shell=True)
        sys.exit()
    except PermissionError as e:
        print(f"PermissionError: {e}")
        messagebox.showerror("Ошибка", f"Отказано в доступе: {e}")


# Начало работы программы
if __name__ == "__main__":
    if CheckUpdate() == True:
        if messagebox.askyesno("Обновление", "Доступно новое обновление. Хотите его установить?"):
            Update()
        
    app = WindowProgramm()
    app.withdraw()  # Скрываем главное окно
    WindowAuthorizationConnectionBD(app)  # Показываем окно перед основным
    app.mainloop()
    del app
