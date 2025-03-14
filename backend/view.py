"""
Описание классов Окон приложения. 

"""
import customtkinter as ctk
from PIL import Image
from tkinter import ttk
from datetime import datetime
from backend.update_app import UpdateApp

__row_click__ = 0   # Переменная для хранения строки таблицы, на которую нажали

class View():
    def __init__(self, model, version: str) -> None:
        app = WindowProgramm(model)
        app.withdraw()
        UpdateApp(version)
        WindowAuthorizationConnectionBD(app, model)
        app.mainloop()
        del app

class WindowAuthorizationConnectionBD(ctk.CTkToplevel):
    def __init__(self, parent: ctk.CTk, model):
        super().__init__(parent)
        self.model = model
        self.parent = parent
        # self.attributes("-topmost", True)
        self.title("Авторизация")
        center_window(self, 300, 200)
        self.configure(fg_color="white")
        self.iconbitmap(self.model.path_ico)

        self.frame_auth = ctk.CTkFrame(self, fg_color="white")
        self.frame_auth.pack(expand=True, fill=ctk.BOTH)
        ctk.CTkLabel(self.frame_auth, text="Авторизация",
                      bg_color="white", font=("Arial", 17)).place(x=100, y=10)
        ctk.CTkLabel(self.frame_auth, text="Логин",
                      bg_color="white", height=15).place(x=130, y=40)
        self.entry_login = ctk.CTkEntry(
            self.frame_auth, width=180, justify="center")
        self.entry_login.insert(0,"sa")

        self.entry_login.place(x=60, y=60)
        ctk.CTkLabel(self.frame_auth, text="Пароль",
                      bg_color="white", height=10).place(x=126, y=90)
        self.entry_password = ctk.CTkEntry(
            self.frame_auth, width=180, justify="center")
        self.entry_password.insert(0,"121270")

        self.entry_password.place(x=60, y=110)
        self.entry_password.configure(show="*")

        self.frame_incorrect = ctk.CTkFrame(self, fg_color="white")

        ctk.CTkLabel(self.frame_incorrect, text="Соединение не установлено",
                      fg_color="white", text_color="red").place(x=75, y=75)
        self.frame_incorrect.pack_forget()

        self.frame_correct = ctk.CTkFrame(self, fg_color="white")

        ctk.CTkLabel(self.frame_correct, text="Соединение установлено",
                      fg_color="white", text_color="green").place(x=75, y=75)
        self.frame_correct.pack_forget()

        self.entry_login.bind(
            "<FocusIn>", lambda e: self.entry_login.delete(0, ctk.END))
        self.entry_password.bind(
            "<FocusIn>", lambda e: self.entry_password.delete(0, ctk.END))

        ctk.CTkButton(self.frame_auth, text="Подключиться",
                       command=self.connect).place(x=80, y=150)
        # Перехватываем нажатие на крестик (закрытие окна)
        self.protocol("WM_DELETE_WINDOW", self.close_app)
        
        self.connect()

    def close_app(self):
        # Завершает весь Tkinter
        self.parent.quit()

    def connect(self):
        self.frame_auth.pack_forget()
        self.model.USERNAME = self.entry_login.get()
        self.model.PASSWORD = self.entry_password.get()
        self.model.connect_db()
        if self.model.is_connect():
            self.frame_correct.pack(expand=True, fill=ctk.BOTH)
            self.after(1000, self.hide_correct_frame)
        else:
            self.frame_incorrect.pack(expand=True, fill=ctk.BOTH)
            self.after(1000, self.hide_incorrect_frame)

    def hide_incorrect_frame(self):
        self.frame_incorrect.pack_forget()
        self.frame_auth.pack(expand=True, fill=ctk.BOTH)
        self.entry_login.delete(0, ctk.END)
        self.entry_password.delete(0, ctk.END)

    def hide_correct_frame(self):
        self.destroy()
        self.parent.deiconify()

class WindowProgramm(ctk.CTk):
    # Инициализация окна программы и откликов в программе
    def __init__(self, model):
        super().__init__()
        self.model = model
        # Описание окна
        self.title("Программа для работы с БД")
        center_window(self, 1000, 600)
        self.configure(fg_color="white")
        # Устанавливаем иконку (Только .ico!)
        self.iconbitmap(self.model.path_ico)

        # Создаём CTkLabel с изображением
        my_image = ctk.CTkImage(
            light_image=Image.open(self.model.path_jpg), size=(150, 100))

        ctk.CTkLabel(self, image=my_image, text="").place(x=10, y=10)

        """--------------------------------------------------ФИО---------------------------------------------------"""

        # Создание ярлыка и текстового поля для ввода фамилии
        ctk.CTkLabel(master=self, text="Фамилия:",
                      bg_color="white").place(x=20, y=150)
        self.entry_familia = ctk.CTkEntry(master=self, width=250)
        self.entry_familia.place(x=90, y=150)
        self.entry_familia.bind("<Return>", command=self.click_find)
        self.entry_familia.insert(0,"ahmed")
        
        # # Создание ярлыка и текстового поля для ввода имени
        ctk.CTkLabel(master=self, text="Имя:",
                      bg_color="white").place(x=20, y=180)
        self.entry_name = ctk.CTkEntry(master=self, width=250)
        self.entry_name.place(x=90, y=180)
        self.entry_name.bind("<Return>", command=self.click_find)

        # # Создание ярлыка и текстового поля для ввода отчества
        ctk.CTkLabel(master=self, text="Отчество:",
                      bg_color="white").place(x=20, y=210)
        self.entry_otchestvo = ctk.CTkEntry(master=self, width=250)
        self.entry_otchestvo.place(x=90, y=210)
        self.entry_otchestvo.bind("<Return>", command=self.click_find)

        """---------------------------------------------------ДАТА--------------------------------------------------"""

        # Создание ярлыка даты рождения
        ctk.CTkLabel(self, text="Дата рождения:",
                      bg_color="white").place(x=20, y=240)

        # Переменные для хранения значений даты рождения
        self.day_var = ctk.StringVar(value="1")
        self.month_var = ctk.StringVar(value="1")
        self.year_var = ctk.StringVar(value="1900")

        # CTkSpinbox для дня от 1 до 31
        ctk.CTkEntry(master=self, textvariable=self.day_var,
                      width=20).place(x=130, y=240)

        # CTkSpinbox для месяца от 1 до 12
        ctk.CTkEntry(master=self, textvariable=self.month_var,
                      width=20).place(x=160, y=240)

        # CTkSpinbox для года от 1900 до 2025
        ctk.CTkEntry(master=self, textvariable=self.year_var,
                      width=45).place(x=190, y=240)

        """---------------------------------------------------КНОПКА-ПОИСКА--------------------------------------------------------------"""

        # Создание кнопки "Поиск"
        ctk.CTkButton(self, text="Поиск",
                       command=self.click_find).place(x=450, y=240)
        
        """---------------------------------------------------КНОПКА-ПРОВЕРКИ--------------------------------------------------------------"""

        # Создание кнопки "Поиск"
        ctk.CTkButton(self, text="Проверка в реестре контролируемых лиц",
                       command=self.click_check_registry).place(x=650, y=240)

        """------------------------------------------------------ТАБЛИЦА-------------------------------------------------------------"""

        # Создание контейнера для таблицы (изначально скрыт)
        self.frame_table = ctk.CTkFrame(
            self, fg_color="white", width=1000, height=310)

        # Отключение видимости таблицы
        self.frame_table.pack_forget()

        # Первая строка с названиями столбцов
        self.columns = ("Фамилия",  # 0      fru
                          "Familia",  # 1      last_lat
                          "Имя",  # 2      name_rus
                          "Imya"  # 3      nla
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
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Cascadia Code", 14))  # Устанавливаем шрифт заголовков
        style.configure("Custom.Treeview", font=("Cascadia Code", 14))  # Устанавливаем шрифт заголовков
        style.configure("Treeview", rowheight=30)
        # Создание таблицы
        self.tree = ttk.Treeview(self.frame_table, columns=self.columns, show='headings', style="Custom.Treeview")

        # Определение размера столбцов
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", stretch=True, width=300)

        # Создание пролистывания по таблице
        scrollbar_y = ctk.CTkScrollbar(
            self.frame_table, width=15, height=50)
        scrollbar_y.place(x=985, y=0)
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        # scrollbar_x = ctk.CTkScrollbar(self.frame_table, width=15, height=15)
        # scrollbar_x.place(x=0, y=295)
        # self.tree.configure(xscrollcommand=scrollbar_x.set)

        self.tree.place(x=0, y=0, relwidth=1, relheight=1)

        """--------------------------------------------------------ОШИБКА-ОТСУТСТВИЯ-ДАННЫХ-----------------------------------------------------------"""

        # Создание контейнера для ошибки при отсутствии вводных данных
        self.frame_false = ctk.CTkFrame(self, fg_color="white")

        # Отключение видимости ошибки
        self.frame_false.pack_forget()

        # Создание ярлыка отсутствия галочек
        ctk.CTkLabel(master=self.frame_false, text="Введите данные",
                      fg_color="white", text_color="red").pack(side="top")

        """--------------------------------------------------------ОШИБКА-ОТСУТСТВИЯ-СТУДЕНТОВ-----------------------------------------------------------"""

        # Создание контейнера для ошибки при отсутствии студентов
        self.frame_false_stud = ctk.CTkFrame(self, fg_color="white")

        # Отключение видимости ошибки
        self.frame_false_stud.pack_forget()

        # Создание ярлыка отсутствия галочек
        ctk.CTkLabel(master=self.frame_false_stud, text="Студент(ы) не найден(ы)",
                      fg_color="white", text_color="red").pack(side="top")

        """-------------------------------------------------------ОТКЛИК-НА-ТАБЛИЦУ----------------------------------------------------------"""

        # Возможность отклика на таблицу
        self.tree.bind("<Double-1>", self.click_on_table)
        

  
    def click_find(self, event = None):

        if self.entry_familia.get() == "":
            if self.entry_name.get() == "":
                if self.entry_otchestvo.get() == "":
                    if self.day_var.get() == "1":
                        if self.month_var.get() == "1":
                            if self.year_var.get() == "1900":

                                self.frame_table.place_forget()
                                self.frame_false.place(x=440, y=290)
                                return

        self.frame_false_stud.place_forget()
        self.frame_false.place_forget()
        self.find()

    # Функция поиска
    def find(self):
        self.tree.delete(*self.tree.get_children())

        surname = self.entry_familia.get().upper() if self.entry_familia.get() else None
        name = self.entry_name.get().upper() if self.entry_name.get() else None
        och = self.entry_otchestvo.get().upper() if self.entry_otchestvo.get() else None
        day = self.day_var.get()
        month = self.month_var.get()
        year = self.year_var.get()
        birthdate = datetime(int(year), int(month), int(day))  # datetime(1990 01 02)

        self.model.find_in_db(surname, name, och, birthdate)

        if self.model.is_data:
            for i, var in enumerate(self.model.data):
                self.tree.insert("", ctk.END, iid=f"I00{i}", values=var)
            self.frame_table.place(x=0, y=290) 
        else:
            self.frame_table.place_forget()
            self.frame_false_stud.place(x=440, y=290)
            

    # Реализация изменения в таблице
    def click_on_table(self, event):
        row = self.tree.identify_row(event.y)
        global __row_click__
        __row_click__ = int(row[1:])  # I000
        WindowInformation(self, self.model)
        
    def click_check_registry(self):
        WindowCheckGosuslugi(self, self.model)

class WindowInformation(ctk.CTkToplevel):

    def __init__(self, parent, model):
        super().__init__(parent)
        self.model = model
        self.parent = parent
        # Описание окна
        self.title("Информация о студенте")
        # center_window(self, 900, 700)
        self.state("zoomed")

        
        self.configure(fg_color="white")
        self.iconbitmap(self.model.path_ico)
        blue = "#6699CC"
        white = "white"
        black = "black"
        
        
        # tabview = ctk.CTkTabview(
        #     self, 
        #     fg_color=white, 
        #     anchor="nw",
        # )
        # tabview.pack(fill="both", expand=True, padx=10, pady=10)
        # tab1 = tabview.add("Информация")
        # tab2 = tabview.add("Редактирование")
        
        # self.frame_view_inf = ctk.CTkFrame(tab1, width=865, height=520, fg_color=white)
        # self.frame_view_inf.place(x=0, y=5)
        
        self.frame_editing = ctk.CTkFrame(self, width=865, height=520, fg_color=blue)
        self.frame_editing.pack(fill="both", expand=True)
        self.scroll_frame_editing = ctk.CTkScrollableFrame(self.frame_editing, width=865, height=300, fg_color=blue)
        self.scroll_frame_editing.pack(fill="both", expand=True)
        
        # self.view_inf()
        self.view_editing()

        # Фиксация окна
        # self.grab_set()
        # self.focus_set()
        # self.wait_window()
        

    def view_inf(self):
        blue = "#6699CC"
        white = "white"
        black = "black"
        my_font = ctk.CTkFont(family="Arial", size=16)

        ctk.CTkFrame(master = self.frame_view_inf, width=855, height=90,
                      fg_color=blue).place(x=5, y=5)
        self.do_lable("Фамилия:", 10, 10, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][0], 80, 10, blue, blue, white) 
        self.do_lable("Имя:", 10, 40, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][2], 80, 40, blue, blue, white) 
        self.do_lable("Отчество:", 10, 70, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][4], 80, 70, blue, blue, white) 

        self.do_lable("Фамилия (лат):", 500, 10, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][1], 600, 10, blue, blue, white) 
        self.do_lable("Имя (лат):", 500, 40, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][3], 600, 40, blue, blue, white) 
        self.do_lable("Отчество (лат):", 500, 70, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][5], 600, 70, blue, blue, white) 

        self.do_lable("Пол:", 10, 100, white, white, black)
        self.do_lable(self.model.data[__row_click__][8], 45, 100, white, white, black) 
        self.do_lable("Дата рождения:", 110, 100, white, white, black)
        self.do_lable(self.model.data[__row_click__][7], 215, 100, white, white, black) 
        self.do_lable("Телефон:", 290, 100, white, white, black)
        self.do_lable(f"+7{self.model.data[__row_click__][20]}", 350, 100, white, white, black) 

        ctk.CTkFrame(self.frame_view_inf, width=400, height=95,
                      fg_color=blue).place(x=5, y=125)
        self.do_lable("Гражданство:", 10, 130, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][6], 100, 130, blue, blue, white) 
        self.do_lable("Государство рождения:",10, 160, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][9], 160, 160, blue, blue, white) 
        self.do_lable("Город рождения:", 10, 190, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][10], 120, 190, blue, blue, white)

        ctk.CTkFrame(self.frame_view_inf, width=400, height=90,
                      fg_color=blue).place(x=5, y=230)
        ctk.CTkLabel(master=self.frame_view_inf, text="Паспортные данные", bg_color=blue,
                      fg_color=blue, text_color=white, font=my_font).place(x=10, y=235)

        self.do_lable("Серия", 10, 265, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][11], 13, 285, blue, white, blue) 
        self.do_lable("Номер", 110, 265, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][12], 113, 285, blue, white, blue) 
        self.do_lable("Дата выдачи", 210, 265, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][13], 213, 285, blue, white, blue) 
        self.do_lable("Срок действия", 300, 265, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][14], 303, 285, blue, white, blue) 

        if self.model.data[__row_click__][28] == "нет":
            ctk.CTkFrame(self.frame_view_inf, width=400, height=90,
                          fg_color=blue).place(x=425, y=230)
            ctk.CTkLabel(master=self.frame_view_inf, text="Виза", bg_color=blue,
                          fg_color=blue, text_color=white, font=my_font).place(x=430, y=235)
            self.do_lable("Серия", 430, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][16], 433, 285, blue, white, blue) 
            self.do_lable("Номер", 530, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][17], 533, 285, blue, white, blue) 
            self.do_lable("Дата выдачи", 630, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][18], 633, 285, blue, white, blue) 
            self.do_lable("Срок действия", 730, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][19], 733, 285, blue, white, blue) 
        elif self.model.data[__row_click__][28] == "ВНЖ":
            ctk.CTkFrame(self.frame_view_inf, width=400, height=90,
                          fg_color=blue).place(x=425, y=230)
            ctk.CTkLabel(master=self.frame_view_inf, text="ВНЖ", bg_color=blue,
                          fg_color=blue, text_color=white, font=my_font).place(x=430, y=235)
            self.do_lable("Серия", 430, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][31], 433, 285, blue, white, blue) 
            self.do_lable("Номер", 530, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][32], 533, 285, blue, white, blue) 
            self.do_lable("Дата выдачи", 630, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][29], 633, 285, blue, white, blue) 
            self.do_lable("Срок действия", 730, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][30], 733, 285, blue, white, blue) 
        elif self.model.data[__row_click__][28] == "РВПО":
            ctk.CTkFrame(self.frame_view_inf, width=400, height=90,
                          fg_color=blue).place(x=425, y=230)
            ctk.CTkLabel(master=self.frame_view_inf, text="РВПО", bg_color=blue,
                          fg_color=blue, text_color=white, font=my_font).place(x=430, y=235)
            self.do_lable("Серия", 430, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][31], 433, 285, blue, white, blue) 
            self.do_lable("Номер", 530, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][32], 533, 285, blue, white, blue) 
            self.do_lable("Дата выдачи", 630, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][29], 633, 285, blue, white, blue) 
            self.do_lable("Срок действия", 730, 265, blue, blue, white)
            self.do_lable(self.model.data[__row_click__][30], 733, 285, blue, white, blue) 

        self.do_lable("Дата въезда", 10, 325, white, white, black)
        self.do_lable(self.model.data[__row_click__][21], 13, 345, white, white, black) 
        self.do_lable("Срок пребывания до", 110,
                            325, white, white, black)
        self.do_lable(self.model.data[__row_click__][22], 139, 345, white, white, black) 

        ctk.CTkFrame(self.frame_view_inf, width=400, height=90,
                      fg_color=blue).place(x=5, y=375)
        ctk.CTkLabel(master=self.frame_view_inf, text="Миграционная карта", bg_color=blue,
                      fg_color=blue, text_color=white, font=my_font).place(x=10, y=380)
        self.do_lable("Серия", 10, 410, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][23], 13, 430, blue, white, blue) 
        self.do_lable("Номер", 100, 410, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][24], 103, 430, blue, white, blue) 

        ctk.CTkFrame(self.frame_view_inf, width=400, height=90,
                      fg_color=blue).place(x=425, y=375)
        ctk.CTkLabel(master=self.frame_view_inf, text="Общежитие", bg_color=blue,
                      fg_color=blue, text_color=white, font=my_font).place(x=430, y=380)
        num = self.model.data[__row_click__][25]
        self.do_lable("Номер корпуса", 435, 410, blue, blue, white)
        self.do_lable(num[0], 470, 430, blue, white, blue) 
        self.do_lable("Дата начала договора", 535, 410, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][26], 565, 430, blue, white, blue) 
        self.do_lable("Номер договора", 680, 410, blue, blue, white)
        self.do_lable(self.model.data[__row_click__][27], 683, 430, blue, white, blue) 

        # Создание кнопки "Сформировать"
        ctk.CTkButton(self.frame_view_inf, text="Сформировать в Excel", text_color=white,
                       fg_color=blue, command=self.click_form_excel).place(x=710, y=490)
        ctk.CTkButton(self.frame_view_inf, text="Сформировать уведомление", text_color=white,
                       fg_color=blue, command=self.click_form_Word).place(x=510, y=490)

    def view_editing(self):
        blue = "#6699CC"
        white = "white"
        black = "black"
        frame_FIO = ctk.CTkFrame(master=self.scroll_frame_editing, height=120, fg_color=white)
        frame_FIO.pack(padx=5, pady=[5,4], anchor="nw", fill="x")
        
        
        ctk.CTkLabel(master=frame_FIO, text="Фамилия: ", fg_color=white, text_color=blue, height=30, anchor = "center").place(x=5, y=5)
        self.fru = ctk.CTkEntry(master=frame_FIO, width=250, height=30)
        self.fru.insert(0, self.model.data[__row_click__][0])
        self.fru.place(x=70, y=5)
        
        ctk.CTkLabel(master=frame_FIO, text="Имя: ", fg_color=white, text_color=blue, height=30, anchor = "center").place(x=5,y=45)
        self.name_rus = ctk.CTkEntry(master=frame_FIO, width=250, height=30)
        self.name_rus.insert(0, self.model.data[__row_click__][2])
        self.name_rus.place(x=70, y=45)
    
        ctk.CTkLabel(master=frame_FIO, text="Отчество: ", fg_color=white, text_color=blue, height=30, anchor = "center").place(x=5,y=85)
        self.och = ctk.CTkEntry(master=frame_FIO, width=250, height=30)
        self.och.insert(0, self.model.data[__row_click__][4])
        self.och.place(x=70, y=85)
        
        
        ctk.CTkLabel(master=frame_FIO, text="Фамилия (лат.): ", fg_color=white, text_color=blue, height=30, anchor = "center").place(x=350, y=5)
        self.last_lat = ctk.CTkEntry(master=frame_FIO, width=250, height=30)
        self.last_lat.insert(0, self.model.data[__row_click__][1])
        self.last_lat.place(x=451, y=5)
        
        ctk.CTkLabel(master=frame_FIO, text="Имя (лат.): ", fg_color=white, text_color=blue, height=30, anchor = "center").place(x=350,y=45)
        self.nla = ctk.CTkEntry(master=frame_FIO, width=250, height=30)
        self.nla.insert(0, self.model.data[__row_click__][3])
        self.nla.place(x=451, y=45)
     
        ctk.CTkLabel(master=frame_FIO, text="Отчество (лат.): ", fg_color=white, text_color=blue, height=30, anchor = "center").place(x=350,y=85)
        self.oche = ctk.CTkEntry(master=frame_FIO, width=250, height=30)
        self.oche.insert(0, self.model.data[__row_click__][5])
        self.oche.place(x=451, y=85)

        frame_2 = ctk.CTkFrame(master=self.scroll_frame_editing, height=105, fg_color=white) #55
        frame_2.pack(padx=5, pady=[1,4], anchor="nw", fill="x")

        ctk.CTkLabel(master=frame_2, text="Пол", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=28,y=5)
        self.sex = ctk.CTkComboBox(master=frame_2, values=["Мужской", "Женский"], width=90, height=20)
        self.sex.configure(variable=ctk.StringVar(value=self.model.data[__row_click__][8]))
        self.sex.place(x=5, y=25)
        
        ctk.CTkLabel(master=frame_2, text="Дата рождения", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=95,y=5)
        self.dob = ctk.CTkEntry(master=frame_2, width=80, height=20)
        if self.model.data[__row_click__][7] != '':
            self.dob.insert(0, self.model.data[__row_click__][7])
        self.dob.place(x=100, y=25)
        
        ctk.CTkLabel(master=frame_2, text="Страна рождения", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=237,y=5)
        self.pob = ctk.CTkEntry(master=frame_2, width=160, height=20)
        self.pob.insert(0, self.model.data[__row_click__][9])
        self.pob.place(x=210, y=25)
        
        ctk.CTkLabel(master=frame_2, text="Город рождения", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=430,y=5)
        self.cob = ctk.CTkEntry(master=frame_2, width=160, height=20)
        self.cob.insert(0, self.model.data[__row_click__][10])
        self.cob.place(x=400, y=25)
        
        ctk.CTkLabel(master=frame_2, text="Гражданство", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=627,y=5)
        self.ctz1 = ctk.CTkEntry(master=frame_2, width=160, height=20)
        self.ctz1.insert(0, self.model.data[__row_click__][6])
        self.ctz1.place(x=590, y=25)
        
        ctk.CTkLabel(master=frame_2, text="Адрес в стране постоянного проживания", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=50)
        self.proz = ctk.CTkEntry(master=frame_2, width=280, height=20)
        self.proz.insert(0, self.model.data[__row_click__][36])
        self.proz.place(x=5, y=70)
        
        frame_3 = ctk.CTkFrame(master=self.scroll_frame_editing, height=80, fg_color=white)
        frame_3.pack(padx=5, pady=[1,4], anchor="nw", fill="x")

        ctk.CTkLabel(master=frame_3, text="ПАСПОРТ", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=5)
        
        ctk.CTkLabel(master=frame_3, text="Серия", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=25)
        self.pas_ser = ctk.CTkEntry(master=frame_3, width=70, height=20)
        self.pas_ser.insert(0, self.model.data[__row_click__][11])
        self.pas_ser.place(x=5, y=45)
        
        ctk.CTkLabel(master=frame_3, text="Номер", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=80,y=25)
        self.pas_num = ctk.CTkEntry(master=frame_3, width=85, height=20)
        self.pas_num.insert(0, self.model.data[__row_click__][12])
        self.pas_num.place(x=80, y=45)
        
        ctk.CTkLabel(master=frame_3, text="Дата выдачи", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=176, y=25)
        self.pds = ctk.CTkEntry(master=frame_3, width=80, height=20)
        if self.model.data[__row_click__][13] != '':
            self.pds.insert(0, self.model.data[__row_click__][13])
        self.pds.place(x=176, y=45)
        
        ctk.CTkLabel(master=frame_3, text="Дата окончания", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=270,y=25)
        self.pde = ctk.CTkEntry(master=frame_3,  width=80, height=20)
        if self.model.data[__row_click__][14] != '':
            self.pde.insert(0, self.model.data[__row_click__][14])
        self.pde.place(x=270, y=45)
        
        frame_4 = ctk.CTkFrame(master=self.scroll_frame_editing, height=130, fg_color=white)
        frame_4.pack(padx=5, pady=[1,4], anchor="nw", fill="x")
        
        ctk.CTkLabel(master=frame_4, text="ВИЗА", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=5)
        
        ctk.CTkLabel(master=frame_4, text="Кратность", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=25)
        self.vis_krat = ctk.CTkEntry(master=frame_4, width=103, height=20)
        self.vis_krat.insert(0, self.model.data[__row_click__][33])
        self.vis_krat.place(x=10, y=45)
        
        ctk.CTkLabel(master=frame_4, text="Дата выдачи", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=130,y=25)
        self.d_poluch = ctk.CTkEntry(master=frame_4, width=80, height=20)
        if self.model.data[__row_click__][18] != '':
            self.d_poluch.insert(0, self.model.data[__row_click__][18])
        self.d_poluch.place(x=130, y=45)
        
        ctk.CTkLabel(master=frame_4, text="Страна выдачи", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=230,y=25)
        self.str_poluch = ctk.CTkEntry(master=frame_4, width=80, height=20)
        self.str_poluch.insert(0, self.model.data[__row_click__][44])
        self.str_poluch.place(x=230, y=45)
        
        ctk.CTkLabel(master=frame_4, text="Город выдачи", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=340,y=25)
        self.city_poluch = ctk.CTkEntry(master=frame_4, width=80, height=20)
        self.city_poluch.insert(0, self.model.data[__row_click__][45])
        self.city_poluch.place(x=340, y=45)
        
        ctk.CTkLabel(master=frame_4, text="Индентификационный номер", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=450,y=25)
        self.vis_id = ctk.CTkEntry(master=frame_4, width=150, height=20)
        self.vis_id.insert(0, self.model.data[__row_click__][34])
        self.vis_id.place(x=450, y=45)
              
        ctk.CTkLabel(master=frame_4, text="Номер приглашения", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=650,y=25)
        self.num_prig = ctk.CTkEntry(master=frame_4, width=150, height=20)
        self.num_prig.insert(0, self.model.data[__row_click__][46])
        self.num_prig.place(x=650, y=45)
        
        ctk.CTkLabel(master=frame_4, text="Серия", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=75)
        self.vis_ser = ctk.CTkEntry(master=frame_4, width=80, height=20)
        self.vis_ser.insert(0, self.model.data[__row_click__][16])
        self.vis_ser.place(x=10, y=95)
              
        ctk.CTkLabel(master=frame_4, text="Номер", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=100,y=75)
        self.vis_num = ctk.CTkEntry(master=frame_4, width=80, height=20)
        self.vis_num.insert(0, self.model.data[__row_click__][17])
        self.vis_num.place(x=100, y=95)
        
        ctk.CTkLabel(master=frame_4, text="Срок действия: c", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=200,y=95)
        self.vis_start = ctk.CTkEntry(master=frame_4, width=80, height=20)
        if self.model.data[__row_click__][18] != '':
            self.vis_start.insert(0, self.model.data[__row_click__][18])
        self.vis_start.place(x=308, y=95)
        ctk.CTkLabel(master=frame_4, text="по", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=396,y=95)
        self.vis_end = ctk.CTkEntry(master=frame_4, width=80, height=20)
        if self.model.data[__row_click__][19] != '':
            self.vis_end.insert(0, self.model.data[__row_click__][19])
        self.vis_end.place(x=420, y=95)
        
        frame_5 = ctk.CTkFrame(master=self.scroll_frame_editing, height=80, fg_color=white)
        frame_5.pack(padx=5, pady=[1,4], anchor="nw", fill="x")
        
        ctk.CTkLabel(master=frame_5, text="МИГРАЦИОННАЯ КАРТА", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=5)

        ctk.CTkLabel(master=frame_5, text="Серия", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=25)
        self.mcs = ctk.CTkEntry(master=frame_5, width=70, height=20)
        self.mcs.insert(0, self.model.data[__row_click__][23])
        self.mcs.place(x=5, y=45)
        
        ctk.CTkLabel(master=frame_5, text="Номер", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=80,y=25)
        self.mcn = ctk.CTkEntry(master=frame_5, width=85, height=20)
        self.mcn.insert(0, self.model.data[__row_click__][24])
        self.mcn.place(x=80, y=45)
        
        ctk.CTkLabel(master=frame_5, text="КПП", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=176, y=25)
        self.kpp = ctk.CTkEntry(master=frame_5, width=150, height=20)
        self.kpp.insert(0, self.model.data[__row_click__][47])
        self.kpp.place(x=176, y=45)
        
        ctk.CTkLabel(master=frame_5, text="Дата въезда в РФ", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=335,y=25)
        self.d_enter = ctk.CTkEntry(master=frame_5,  width=80, height=20)
        if self.model.data[__row_click__][21] != '':
            self.d_enter.insert(0, self.model.data[__row_click__][21])
        self.d_enter.place(x=335, y=45)

        frame_6 = ctk.CTkFrame(master=self.scroll_frame_editing, height=105, fg_color=white)
        frame_6.pack(padx=5, pady=[1,4], anchor="nw", fill="x")
        
        ctk.CTkLabel(master=frame_6, text="ГРАЖДАНСТВО РФ", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=5)

        ctk.CTkLabel(master=frame_6, text="Вид:", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=25)
        self.rf = ctk.CTkComboBox(master=frame_6, values=["ВНЖ", "РВПО", "нет"], width=110, height=20)
        self.rf.configure(variable=ctk.StringVar(value=self.model.data[__row_click__][28]))
        self.rf.place(x=50, y=25)
        
        ctk.CTkLabel(master=frame_6, text="Серия", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=50)
        self.ser = ctk.CTkEntry(master=frame_6, placeholder_text="12", width=70, height=20)
        self.ser.insert(0, self.model.data[__row_click__][31])
        self.ser.place(x=5, y=70)
        
        ctk.CTkLabel(master=frame_6, text="Номер", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=80,y=50)
        self.nmr = ctk.CTkEntry(master=frame_6, width=85, height=20)
        self.nmr.insert(0, self.model.data[__row_click__][32])
        self.nmr.place(x=80, y=70)
        
        ctk.CTkLabel(master=frame_6, text="Дата выдачи", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=176, y=50)
        self.rfd = ctk.CTkEntry(master=frame_6, width=80, height=20)
        self.rfd.insert(0, self.model.data[__row_click__][29])
        self.rfd.place(x=176, y=70)
        
        ctk.CTkLabel(master=frame_6, text="Дата окончания", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=270,y=50)
        self.mot = ctk.CTkEntry(master=frame_6,  width=80, height=20)
        self.mot.insert(0, self.model.data[__row_click__][30])
        self.mot.place(x=270, y=70)

        frame_7 = ctk.CTkFrame(master=self.scroll_frame_editing, height=140, fg_color=white)
        frame_7.pack(padx=5, pady=[1,4], anchor="nw", fill="x")
        
        ctk.CTkLabel(master=frame_7, text="ОБУЧЕНИЕ", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=5)

        ctk.CTkLabel(master=frame_7, text="Статус", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=25)
        self.uch_st_st = ctk.CTkComboBox(master=frame_7, values=["Студент", "Аспирант", "Отчислен"], width=90, height=20)
        self.uch_st_st.configure(variable=ctk.StringVar(value=self.model.data[__row_click__][50]))
        self.uch_st_st.place(x=10, y=45)
        
        ctk.CTkLabel(master=frame_7, text="Программа", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=105,y=25)
        self.pr = ctk.CTkComboBox(master=frame_7, values=["бак", "маг", "спец"], width=80, height=20)
        self.pr.configure(variable=ctk.StringVar(value=self.model.data[__row_click__][51]))
        self.pr.place(x=105, y=45)
        
        ctk.CTkLabel(master=frame_7, text="Форма", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=205,y=25)
        self.fr = ctk.CTkEntry(master=frame_7, width=96, height=20)
        self.fr.insert(0, self.model.data[__row_click__][52])
        self.fr.place(x=205, y=45)

        ctk.CTkLabel(master=frame_7, text="Направление", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=310,y=25)
        self.gos_nap = ctk.CTkComboBox(master=frame_7, values=["гос. направление", "контракт", "приказ (с РФ)"], width=140, height=20)
        self.gos_nap.configure(variable=ctk.StringVar(value=self.model.data[__row_click__][35]))
        self.gos_nap.place(x=310, y=45)
        
        ctk.CTkLabel(master=frame_7, text="Номер договора", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=470,y=25)
        self.kontrakt = ctk.CTkEntry(master=frame_7, placeholder_text="HTI-10082/24", width=96, height=20)
        self.kontrakt.insert(0, self.model.data[__row_click__][37])
        self.kontrakt.place(x=470, y=45)
        
        ctk.CTkLabel(master=frame_7, text="Специальность", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=590,y=25)
        self.o_p = ctk.CTkEntry(master=frame_7, placeholder_text="Менеджмент", width=200, height=20)
        self.o_p.insert(0, self.model.data[__row_click__][53])
        self.o_p.place(x=590, y=45)

        ctk.CTkLabel(master=frame_7, text="Срок обучения: c", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=75)
        self.kont_start = ctk.CTkEntry(master=frame_7,  width=80, height=20)
        if self.model.data[__row_click__][38] != '':
            self.kont_start.insert(0, self.model.data[__row_click__][38])
        self.kont_start.place(x=120, y=75)
        
        ctk.CTkLabel(master=frame_7, text="по", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=205,y=75)
        self.kont_end = ctk.CTkEntry(master=frame_7,  width=80, height=20)
        if self.model.data[__row_click__][39] != '':
            self.kont_end.insert(0, self.model.data[__row_click__][39])
        self.kont_end.place(x=225, y=75)

        ctk.CTkLabel(master=frame_7, text="Приказ №", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=105)
        self.prikaz = ctk.CTkEntry(master=frame_7,  width=80, height=20)
        self.prikaz.insert(0, self.model.data[__row_click__][54])
        self.prikaz.place(x=80, y=105)
        ctk.CTkLabel(master=frame_7, text="от", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=165,y=105)
        self.prik_start = ctk.CTkEntry(master=frame_7,  width=80, height=20)
        if self.model.data[__row_click__][55] != '':
            self.prik_start.insert(0, self.model.data[__row_click__][55])
        self.prik_start.place(x=185, y=105)
        
        ctk.CTkLabel(master=frame_7, text="статус", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=275,y=105)
        self.o_s = ctk.CTkEntry(master=frame_7,  width=150, height=20)
        self.o_s.insert(0, self.model.data[__row_click__][56])
        self.o_s.place(x=320, y=105)
        
        frame_8 = ctk.CTkFrame(master=self.scroll_frame_editing, height=135, fg_color=white)
        frame_8.pack(padx=5, pady=[1,4], anchor="nw", fill="x")
        
        ctk.CTkLabel(master=frame_8, text="ОБЩЕЖИТИЕ", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=5)
        
        ctk.CTkLabel(master=frame_8, text="Догвор найма №", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=35)
        self.dog_obsh = ctk.CTkEntry(master=frame_8,  width=100, height=20)
        self.dog_obsh.insert(0, self.model.data[__row_click__][27])
        self.dog_obsh.place(x=115, y=35)
        ctk.CTkLabel(master=frame_8, text="с", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=220,y=35)
        self.dnd = ctk.CTkEntry(master=frame_8,  width=80, height=20)
        self.dnd.insert(0, self.model.data[__row_click__][26])
        self.dnd.place(x=235, y=35)
        ctk.CTkLabel(master=frame_8, text="по", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=320,y=35)
        self.d_naym = ctk.CTkEntry(master=frame_8,  width=80, height=20)
        if self.model.data[__row_click__][57] != '':
            self.d_naym.insert(0, self.model.data[__row_click__][57])
        self.d_naym.place(x=340, y=35)
        
        ctk.CTkLabel(master=frame_8, text="Адрес:", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=65)
        self.addres = ctk.CTkEntry(master=frame_8,  width=300, height=20)
        self.addres.insert(0,"Г. МОСКВА, КОЧНОВСКИЙ ПР., Д.7")
        self.addres.place(x=55, y=65)
        
        ctk.CTkLabel(master=frame_8, text="Комната:", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=380,y=65)
        self.k = ctk.CTkEntry(master=frame_8,  width=50, height=20)
        self.k.insert(0, self.model.data[__row_click__][25])
        self.k.place(x=440, y=65)
        
        ctk.CTkLabel(master=frame_8, text="Прежний адрес:", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=95)
        self.star = ctk.CTkEntry(master=frame_8,  width=400, height=20)
        self.star.insert(0, self.model.data[__row_click__][42])
        self.star.place(x=110, y=95)
        
        frame_9 = ctk.CTkFrame(master=self.scroll_frame_editing, height=150, fg_color=white)
        frame_9.pack(padx=5, pady=[1,4], anchor="nw", fill="x")
        
        ctk.CTkLabel(master=frame_9, text="ПРИМЕЧАНИЕ", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=5)
        
        ctk.CTkLabel(master=frame_9, text="e-mail:", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=35)
        self.email = ctk.CTkEntry(master=frame_9,  width=300, height=20)
        self.email.insert(0, self.model.data[__row_click__][58])
        self.email.place(x=50, y=35)
        
        ctk.CTkLabel(master=frame_9, text="Телефон:", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=370,y=35)
        self.tel_nom = ctk.CTkEntry(master=frame_9,  width=110, height=20)
        self.tel_nom.insert(0, self.model.data[__row_click__][20])
        self.tel_nom.place(x=430, y=35)
        
        ctk.CTkLabel(master=frame_9, text="Примечание:", fg_color=white, text_color=blue, height=20, anchor = "center").place(x=10,y=65)
        self.prim = ctk.CTkEntry(master=frame_9,  width=1000, height=60)
        self.prim.insert(0, self.model.data[__row_click__][59])
        self.prim.place(x=10, y=85)
        
        
        ctk.CTkButton(self.frame_editing, text="Сохранить изменения", text_color=blue, 
                      fg_color=white, width=50, height=30, command=self.click_on_save).pack(anchor="w",padx=10, pady=5)
        ctk.CTkButton(self.frame_editing, text="Сформировать в Excel", text_color=blue,
                       fg_color=white, command=self.click_form_excel).pack(anchor="w",padx=10, pady=5)
        ctk.CTkButton(self.frame_editing, text="Сформировать уведомление", text_color=blue,
                       fg_color=white, command=self.click_form_Word).pack(anchor="w",padx=10, pady=5)



    def do_lable(self, text, x, y, bg_color, fg_color, text_color):
        if text == "" or text.isspace():
            ctk.CTkLabel(master=self.frame_view_inf, text="Отсутствует", height=10, fg_color=fg_color,
                      bg_color=bg_color, text_color=text_color).place(x=x, y=y)
        else: 
            ctk.CTkLabel(master=self.frame_view_inf, text=text, height=10, fg_color=fg_color,
                      bg_color=bg_color, text_color=text_color).place(x=x, y=y)

    def click_form_excel(self):
        WindowSaveExcel(self, self.model)

    def click_form_Word(self):
        WindowSaveWord(self, self.model)

    def click_on_save(self):
        self.model.process_update_database(self, __row_click__)
        self.parent.frame_table.place_forget()
        self.destroy()

class WindowCheckGosuslugi(ctk.CTkToplevel):

    def __init__(self, parent, model):
        super().__init__(parent)
        self.model = model
        # Описание окна
        self.title("Проверка в реестре контролируемых лиц")
        center_window(self, 1000, 600)
        self.configure(fg_color="white")
        """--------------------------------------------------ФИО---------------------------------------------------"""

        # Создание ярлыка и текстового поля для ввода фамилии
        ctk.CTkLabel(master=self, text="Фамилия:",
                      bg_color="white").place(x=20, y=150)
        self.entry_familia = ctk.CTkEntry(master=self, width=250)
        self.entry_familia.place(x=90, y=150)
        # self.entry_familia.bind("<Return>", command=self.click_find)
        
        # # Создание ярлыка и текстового поля для ввода имени
        ctk.CTkLabel(master=self, text="Имя:",
                      bg_color="white").place(x=20, y=180)
        self.entry_name = ctk.CTkEntry(master=self, width=250)
        self.entry_name.place(x=90, y=180)
        # self.entry_name.bind("<Return>", command=self.click_find)

        # # Создание ярлыка и текстового поля для ввода отчества
        ctk.CTkLabel(master=self, text="Отчество:",
                      bg_color="white").place(x=20, y=210)
        self.entry_otchestvo = ctk.CTkEntry(master=self, width=250)
        self.entry_otchestvo.place(x=90, y=210)
        # self.entry_otchestvo.bind("<Return>", command=self.click_find)

        """---------------------------------------------------ДАТА--------------------------------------------------"""

        # Создание ярлыка даты рождения
        ctk.CTkLabel(self, text="Дата рождения:",
                      bg_color="white").place(x=20, y=240)

        # Переменные для хранения значений даты рождения
        self.day_var = ctk.StringVar(value="1")
        self.month_var = ctk.StringVar(value="1")
        self.year_var = ctk.StringVar(value="1900")

        # CTkSpinbox для дня от 1 до 31
        ctk.CTkEntry(master=self, textvariable=self.day_var,
                      width=20).place(x=130, y=240)

        # CTkSpinbox для месяца от 1 до 12
        ctk.CTkEntry(master=self, textvariable=self.month_var,
                      width=20).place(x=160, y=240)

        # CTkSpinbox для года от 1900 до 2025
        ctk.CTkEntry(master=self, textvariable=self.year_var,
                      width=45).place(x=190, y=240)

        """---------------------------------------------------КНОПКА-ПОИСКА--------------------------------------------------------------"""

        # Создание кнопки "Поиск"
        ctk.CTkButton(self, text="Поиск",
                       command=self.click_find).place(x=450, y=240)
        """---------------------------------------------------КНОПКА-ПОКАЗАТЬ-ВСЕХ--------------------------------------------------------------"""

        # Создание кнопки "Поиск"
        ctk.CTkButton(self, text="Показать всех",
                       command=self.click_find_all).place(x=650, y=240)


        """------------------------------------------------------ТАБЛИЦА-------------------------------------------------------------"""

        # Создание контейнера для таблицы (изначально скрыт)
        self.frame_table = ctk.CTkFrame(
            self, fg_color="white", width=1000, height=310)

        # Отключение видимости таблицы
        self.frame_table.pack_forget()

        # Первая строка с названиями столбцов
        self.columns = ("Фамилия",  # 0      fru
                          "Familia",  # 1      last_lat
                          "Имя",  # 2      name_rus
                          "Imya"  # 3      nla
                        )
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Cascadia Code", 14))  # Устанавливаем шрифт заголовков
        style.configure("Custom.Treeview", font=("Cascadia Code", 14))  # Устанавливаем шрифт заголовков
        style.configure("Treeview", rowheight=30)
        # Создание таблицы
        self.tree = ttk.Treeview(self.frame_table, columns=self.columns, show='headings', style="Custom.Treeview")

        # Определение размера столбцов
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", stretch=True, width=300)

        # Создание пролистывания по таблице
        scrollbar_y = ctk.CTkScrollbar(
            self.frame_table, width=15, height=50)
        scrollbar_y.place(x=985, y=0)
        self.tree.configure(yscrollcommand=scrollbar_y.set)

        # scrollbar_x = ctk.CTkScrollbar(self.frame_table, width=15, height=15)
        # scrollbar_x.place(x=0, y=295)
        # self.tree.configure(xscrollcommand=scrollbar_x.set)

        self.tree.place(x=0, y=0, relwidth=1, relheight=1)

        """--------------------------------------------------------ОШИБКА-ОТСУТСТВИЯ-ДАННЫХ-----------------------------------------------------------"""

        # Создание контейнера для ошибки при отсутствии вводных данных
        self.frame_false = ctk.CTkFrame(self, fg_color="white")

        # Отключение видимости ошибки
        self.frame_false.pack_forget()

        # Создание ярлыка отсутствия галочек
        ctk.CTkLabel(master=self.frame_false, text="Введите данные",
                      fg_color="white", text_color="red").pack(side="top")

        """--------------------------------------------------------ОШИБКА-ОТСУТСТВИЯ-СТУДЕНТОВ-----------------------------------------------------------"""

        # Создание контейнера для ошибки при отсутствии студентов
        self.frame_false_stud = ctk.CTkFrame(self, fg_color="white")

        # Отключение видимости ошибки
        self.frame_false_stud.pack_forget()

        # Создание ярлыка отсутствия галочек
        ctk.CTkLabel(master=self.frame_false_stud, text="Студент(ы) не найден(ы)",
                      fg_color="white", text_color="red").pack(side="top")

        """-------------------------------------------------------ОТКЛИК-НА-ТАБЛИЦУ----------------------------------------------------------"""

        # Возможность отклика на таблицу
        self.tree.bind("<Double-1>", self.click_on_table)
        
                # Фиксация окна
                
        self.protocol("WM_DELETE_WINDOW", self.close_app)
                
        self.grab_set()
        self.focus_set()
        self.wait_window()
        
    def click_find_all(self, event = None):

        self.frame_false_stud.place_forget()
        self.frame_false.place_forget()
        self.find(2)

    def click_find(self, event = None):

        if self.entry_familia.get() == "":
            if self.entry_name.get() == "":
                if self.entry_otchestvo.get() == "":
                    if self.day_var.get() == "1":
                        if self.month_var.get() == "1":
                            if self.year_var.get() == "1900":

                                self.frame_table.place_forget()
                                self.frame_false.place(x=440, y=290)
                                return

        self.frame_false_stud.place_forget()
        self.frame_false.place_forget()
        self.find(1)

    # Функция поиска
    def find(self, check):
        self.tree.delete(*self.tree.get_children())

        surname = self.entry_familia.get().upper() if self.entry_familia.get() else None
        name = self.entry_name.get().upper() if self.entry_name.get() else None
        och = self.entry_otchestvo.get().upper() if self.entry_otchestvo.get() else None
        day = self.day_var.get()
        month = self.month_var.get()
        year = self.year_var.get()
        birthdate = datetime(int(year), int(month), int(day))  # datetime(1990 01 02)
        if check == 1:
            self.model.find_in_db_for_check(surname, name, och, birthdate)
        elif check == 2:
            self.model.find_in_db_all_rows_for_check()

        if self.model.is_data:
            for i, var in enumerate(self.model.data):
                self.tree.insert("", ctk.END, iid=f"I00{i}", values=var)
            self.frame_table.place(x=0, y=290) 
        else:
            self.frame_table.place_forget()
            self.frame_false_stud.place(x=440, y=290)

    def click_on_table(self, event):
        row = self.tree.identify_row(event.y)
        global __row_click__
        __row_click__ = int(row[1:])  # I000
        self.model.check_on_gosuslugi(__row_click__)

    def close_app(self):
        if self.model.check_open_browser():
            self.model.close_browser()
        self.destroy()

class WindowSaveExcel(ctk.CTkToplevel):
    def __init__(self, parent, model):
        super().__init__(parent)
        self.model = model
        # Описание окна
        self.title("Сохранение в Exel")
        center_window(self, 300, 100)
        self.configure(fg_color="white")
        self.iconbitmap(self.model.path_ico)
        # Создание ярлыка сохранения
        ctk.CTkLabel(master=self, text="Сохранение", fg_color="white",
                      text_color="black", font=("Arial", 16)).pack(side="top", padx=7)

        # Создание контейнера для названия файла
        self.frame_name_file = ctk.CTkFrame(
            self, fg_color="white")
        self.frame_name_file.pack(fill=ctk.X)

        # Создание ярлыка и текстового поля для ввода названия файла
        ctk.CTkLabel(master=self.frame_name_file, text="Название файла:",
                      fg_color="white", text_color="black").grid(row=0, column=0, padx=7)
        self.entry_name_file = ctk.CTkEntry(
            master=self.frame_name_file, width=150)
        self.entry_name_file.insert(
            0, f"{self.model.data[__row_click__][1]} {self.model.data[__row_click__][3]}")

        self.entry_name_file.grid(row=0, column=1, pady=7)

        # Создание кнопки "Сохранить"
        self.search_button = ctk.CTkButton(
            self, text="Сохранить", command=self.click_save)
        self.search_button.pack(side="right", padx=5)

        # Обработка нажатия внутри области
        self.entry_name_file.bind("<FocusIn>", lambda e: self.entry_name_file.delete(
            0, len(self.entry_name_file.get())+1))

        # Фиксация окна
        self.grab_set()
        self.focus_set()
        self.wait_window()

    def click_save(self):
        self.model.completion_excel(self.entry_name_file.get(), __row_click__)
        self.destroy()

class WindowSaveWord(ctk.CTkToplevel):
    def __init__(self, parent, model):
        super().__init__(parent)
        self.model = model
        # Описание окна
        self.title("Сохранение в Exel")
        center_window(self, 500, 250)
        self.configure(fg_color="white")
        self.iconbitmap(self.model.path_ico)
        # Создание ярлыка сохранения
        ctk.CTkLabel(master=self, text="Сохранение", fg_color="white",
                      text_color="white", font=("Arial", 16)).place(x=210, y=10)

        self.var_check_1 = ctk.BooleanVar()
        self.var_check_2 = ctk.BooleanVar()
        self.var_check_3 = ctk.BooleanVar()

        # # Создание галочек
        ctk.CTkCheckBox(self, variable=self.var_check_1,
                         bg_color="white").place(x=10, y=50)
        ctk.CTkLabel(master=self, text="о предоставлении иностранному гражданину (лицу без гражданства)\n"
                      "академического отпуска образовательной/научной организации;", bg_color="white").place(x=40, y=50)
        ctk.CTkCheckBox(self, variable=self.var_check_2,
                         bg_color="white").place(x=10, y=90)
        ctk.CTkLabel(master=self, text="о досрочном прекращении обучения иностранного гражданина\n"
                      "(лица без гражданства) в образовательной/научной организации;", bg_color="white").place(x=40, y=90)
        ctk.CTkCheckBox(self, variable=self.var_check_3,
                         bg_color="white").place(x=10, y=130)
        ctk.CTkLabel(master=self, text="о завершении обучения иностранного гражданина (лица без гражданства)\n"
                      "(лица без гражданства) в образовательной/научной организации;", bg_color="white").place(x=40, y=130)

        # Создание ярлыка и текстового поля для ввода названия файла
        ctk.CTkLabel(master=self, text="Название файла:",
                      bg_color="white").place(x=30, y=180)
        self.entry_name_file = ctk.CTkEntry(
            master=self, width=300)
        self.entry_name_file.insert(
            0, f"Уведомление {self.model.data[__row_click__][1]} {self.model.data[__row_click__][3]}")
        self.entry_name_file.place(x=30, y=200)

        # Создание кнопки "Сохранить"
        self.search_button = ctk.CTkButton(
            self, text="Сохранить", command=self.click_save)
        self.search_button.place(x=350, y=199)

        # Обработка нажатия внутри области
        self.entry_name_file.bind("<FocusIn>", lambda e: self.entry_name_file.delete(
            0, len(self.entry_name_file.get())+1))

        # Фиксация окна
        self.grab_set()
        self.focus_set()
        self.wait_window()

    def click_save(self):
        check = None
        if self.var_check_1.get() == True:
            check = 1
        if self.var_check_2.get() == True:
            check = 2
        if self.var_check_3.get() == True:
            check = 3
        self.model.complection_word(self.entry_name_file.get(), __row_click__, check)
        self.destroy()

def center_window(
                self, 
                width: int, 
                height: int
) -> None:

    # Получаем размеры экрана
    screen_width = self.winfo_screenwidth()
    screen_height = self.winfo_screenheight()

    # Вычисляем координаты верхнего левого угла
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размеры и положение окна
    self.geometry(f"{width}x{height}+{x}+{y}")

def full_screen(self):
    # Получаем размеры экрана
    screen_width = self.winfo_screenwidth()
    screen_height = self.winfo_screenheight()
    self.geometry(f"{screen_width}x{screen_height}+0+0")
    