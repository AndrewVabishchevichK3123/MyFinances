from datetime import date, datetime
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
import sqlite3, copy, csv, sys, os, openpyxl

class App:
    def __init__(self, root):

        self.root = root
        self.root.title("Мои финансы")
        self.root.geometry('1280x780')
        self.root.minsize(1280, 780)


        # Устанавливаем цвет фона для корневого виджета
        self.root.configure(background="#319158")

        self.flag=False
        self.load=[]

        self.journal = [[]]
        self.num = 1
        self.collections=['Супермаркеты','Рестораны и кафе','Здоровье и красота','Транспорт','Развлечения и хобби','Переводы']

        self.filtered=[]

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.connect_to_db()

        # Создание элементов управления
        self.create_widgets()

    def create_widgets(self):
        # Надпись "Мои финансы"
        self.lbl1 = Label(root, text="Мои финансы", font=("Calibri", 30, "bold"), foreground="white",
                          background="#319158")
        self.lbl1.place(relx=0.5, rely=0.5, anchor='center')

        frame = Frame(self.root, bg="#003200", width=1280, height=40)
        frame.place(relx=0.5, rely=0.6, anchor='center')

        # Надпись "Введите имя пользователя"
        self.lbl2 = Label(root, text="Введите имя пользователя:", font=("Calibri", 16), foreground="white",
                          background="#003200")
        self.lbl2.place(relx=0.4, rely=0.6, anchor='center')

        # Поле ввода имени пользователя
        self.chk_state = StringVar()
        self.txt = Entry(root, width=43)
        self.txt.place(relx=0.6, rely=0.6, anchor='center')

        # Надпись "Вход"
        self.btn = Button(self.root, text="Вход", font=("Calibri", 16), background="#003200", foreground="white",
                          command=self.begin_page)

        self.btn.place(relx=0.5, rely=0.7, anchor='center')

    def begin_page(self):
        self.username = self.txt.get()
        if len(self.username) > 20 or len(self.username) == 0:
            self.clear()
            self.create_widgets()
            self.lbl4 = Label(root,
                              text="Ошибка! (Имя пользователя не должно быть пустым и должно состоять не более чем из 20-ти символов)",
                              font=("Calibri", 12), foreground="red",
                              background="#319158")
            self.lbl4.place(relx=0.5, rely=0.8, anchor='center')

        else:
            self.clear()
            self.menu_page()

    def clear(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def menu_page(self):
        print(self.journal)

        self.root.geometry('1200x700')
        self.root.minsize(1200, 700)
        self.root.maxsize(1300, 700)
        self.root.title("Мои финансы")

        menu = Menu(self.root)
        self.root.config(menu=menu)

        new_item = Menu(menu)
        new_item.add_command(label='Сохранить данные в файл', command=self.save)
        new_item.add_command(label='Завершить работу', command=self.exit)
        menu.add_cascade(label='Программа', menu=new_item)

        new_item2 = Menu(menu)
        new_item2.add_command(label='О программе "Мои финансы"', command=self.info)
        menu.add_cascade(label='Справка', menu=new_item2)

        self.tor = ['№', 'Название', 'Цена', 'Дата', 'Коллекция']

        self.data = self.load_from_db()

        if self.data == []:
            if self.journal == [[]] or self.journal == []:
                n = '-' * 46
                self.journal = [[n, n, n, n, n]]
        else:
            self.journal = []
            for el in self.data:
                self.journal.append([int(el[0]), el[1], el[2], el[3], el[4]])
                self.load=copy.deepcopy(self.journal)
                if el[4] not in self.collections:
                    self.collections.append(el[4])

            if self.journal[0][0]!=1:
                cnt=1
                for i in range(len(self.journal)):
                    self.journal[i][0]=cnt+i
            self.num = self.journal[-1][0] + 1

        self.delete_from_db(self.username)
        self.tree = ttk.Treeview(columns=self.tor, show="headings", style="mystyle.Treeview")
        self.tree.pack(fill=BOTH, expand=1)

        self.style = ttk.Style()
        self.style.theme_use("default")
        self.style.configure("mystyle.Treeview", background="#319158", foreground="white", fieldbackground="#319158",
                             font=("Calibri", 12))
        self.style.map("mystyle.Treeview", background=[('selected', '#003200')])

        self.tree.heading("№", text="№")
        self.tree.heading('Название', text='Название')
        self.tree.heading('Цена', text='Цена')
        self.tree.heading('Дата', text='Дата')
        self.tree.heading('Коллекция', text='Коллекция')

        for el in self.journal:
            self.tree.insert("", END, values=el)

        self.tree.pack(expand=True, fill='both')

        # Привязываем функцию сортировки к щелчку на заголовке столбца
        for col in self.tree['columns']:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_column(self.tree, _col, False))

        # Привязываем контекстное меню к Treeview
        self.tree.bind("<Button-3>", self.popup)
        self.popup_menu = Menu(self.root, tearoff=0)
        self.popup_menu.add_command(label="Удалить", command=self.delete_item)

        self.btn2 = Button(self.root, text="Добавить продукт", font=("Calibri", 16), background="#003200",
                           foreground="white", command=self.add_product)
        self.btn2.pack()

        self.btn0 = Button(self.root, text="Очистить все", font=("Calibri", 16), background="#003200",
                           foreground="white", command=self.clear_table)
        self.btn0.pack()

        self.lbl5 = Label(root, text="Сумма: "+str(self.summa()), font=("Calibri", 16), foreground="white",
                          background="#003200")
        self.lbl5.place(relx=0.1, rely=0.9, anchor='center')

        self.btn_apply_filter = Button(self.root, text="Применить фильтр", font=("Calibri", 16), background="#003200",
                                       foreground="white", command=self.apply_filter)
        self.btn_apply_filter.place(relx=0.85, rely=0.9, anchor='center')

    def apply_filter(self):
        n = '-' * 46
        if self.journal == [[n, n, n, n, n]] or self.journal == [[]] or self.journal == []:
            messagebox.showinfo("Уведомление", "Таблица пустая")

        else:
            self.filter = Toplevel(self.root)
            self.filter.title("Фильтр")
            self.filter.geometry("450x680")
            self.filter.minsize(450, 680)
            self.filter.maxsize(550, 780)
            self.filter.protocol("WM_DELETE_WINDOW", self.on_closing2)

            self.lbl_min_price = Label(self.filter, text="Минимальная цена:", font=("Calibri", 12), foreground="white",
                                       background="#003200")
            self.lbl_min_price.place(relx=0.1, rely=0.1)

            self.entry_min_price = Entry(self.filter, width=10)
            self.entry_min_price.place(relx=0.6, rely=0.104)

            self.lbl_max_price = Label(self.filter, text="Максимальная цена:", font=("Calibri", 12), foreground="white",
                                       background="#003200")
            self.lbl_max_price.place(relx=0.1, rely=0.2)

            self.entry_max_price = Entry(self.filter, width=10)
            self.entry_max_price.place(relx=0.6, rely=0.204)

            self.filter.configure(background="#003200")
            self.chk_state2 = IntVar(value=1)

            self.rbtime = Radiobutton(self.filter, text="За Период", font=("Calibri", 12),
                                    value=1, variable=self.chk_state2, command=self.show_date_widgets)
            self.rbtime.place(relx=0.1, rely=0.3)

            self.rbtime2 = Radiobutton(self.filter, text="За все время", font=("Calibri", 12),
                                      value=2, variable=self.chk_state2, command=self.hide_date_widgets)
            self.rbtime2.place(relx=0.5, rely=0.3)


            self.lbl_min_date = Label(self.filter, text="Начальная дата (ГГГГ-ММ-ДД):", font=("Calibri", 12),
                                      foreground="white",
                                      background="#003200")

            self.entry_min_date = DateEntry(self.filter, width=10, background='#319158', foreground='white',
                                            borderwidth=2,
                                            date_pattern="yyyy-mm-dd")

            self.lbl_max_date = Label(self.filter, text="Конечная дата (ГГГГ-ММ-ДД):", font=("Calibri", 12),
                                      foreground="white",
                                      background="#003200")

            self.entry_max_date = DateEntry(self.filter, width=10, background='#319158', foreground='white',
                                            borderwidth=2,
                                            date_pattern="yyyy-mm-dd")

            self.show_date_widgets()

            self.col11 = IntVar(value=1)
            self.col22 = IntVar(value=1)
            self.col33 = IntVar(value=1)
            self.col44 = IntVar(value=1)
            self.col55 = IntVar(value=1)
            self.col66 = IntVar(value=1)
            self.col77 = IntVar(value=1)
            self.col88 = IntVar(value=1)

            self.col1 = Checkbutton(self.filter, text="Супермаркеты", font=("Calibri", 10),
                               variable=self.col11)
            self.col1.place(relx=0.1, rely=0.55)

            self.col2 = Checkbutton(self.filter, text="Рестораны и кафе", font=("Calibri", 10),
                                    variable=self.col22)
            self.col2.place(relx=0.1, rely=0.6)

            self.col3 = Checkbutton(self.filter, text="Здоровье и красота", font=("Calibri", 10),
                                    variable=self.col33)
            self.col3.place(relx=0.1, rely=0.65)

            self.col4 = Checkbutton(self.filter, text="Транспорт", font=("Calibri", 10),
                                    variable=self.col44)
            self.col4.place(relx=0.1, rely=0.7)

            self.col5 = Checkbutton(self.filter, text="Развлечения и хобби", font=("Calibri", 10),
                                    variable=self.col55)
            self.col5.place(relx=0.6, rely=0.55)

            self.col6 = Checkbutton(self.filter, text="Переводы", font=("Calibri", 10),
                                    variable=self.col66)
            self.col6.place(relx=0.6, rely=0.6)

            if len(self.collections)==7:
                self.col7 = Checkbutton(self.filter, text=self.collections[-1], font=("Calibri", 10),
                                        variable=self.col77)
                self.col7.place(relx=0.6, rely=0.65)

            elif len(self.collections)==8:
                self.col7 = Checkbutton(self.filter, text=self.collections[-2], font=("Calibri", 10),
                                        variable=self.col77)
                self.col7.place(relx=0.6, rely=0.65)

                self.col8 = Checkbutton(self.filter, text=self.collections[-1], font=("Calibri", 10),
                                        variable=self.col88)
                self.col8.place(relx=0.6, rely=0.7)

            self.btn_apply = Button(self.filter, text="Применить", font=("Calibri", 12), background="#319158",
                               foreground="white", command=self.apply_changes_in_filter)
            self.btn_apply.place(relx=0.2, rely=0.8, anchor='center')

            self.btn_cancel = Button(self.filter, text="Очистить", font=("Calibri", 12), background="#319158",
                                    foreground="white", command=self.clear_filter)
            self.btn_cancel.place(relx=0.8, rely=0.8, anchor='center')

            self.btn_close = Button(self.filter, text="Закрыть", font=("Calibri", 12), background="#319158",
                                     foreground="white", command=self.close_filter)
            self.btn_close.place(relx=0.5, rely=0.9, anchor='center')

    def show_date_widgets(self):
        self.lbl_min_date.place(relx=0.1, rely=0.4)
        self.entry_min_date.place(relx=0.6, rely=0.404)
        self.lbl_max_date.place(relx=0.1, rely=0.5)
        self.entry_max_date.place(relx=0.6, rely=0.504)

    def hide_date_widgets(self):
        self.lbl_min_date.place_forget()
        self.entry_min_date.place_forget()
        self.lbl_max_date.place_forget()
        self.entry_max_date.place_forget()

    def apply_changes_in_filter(self):
        # Получаем значения диапазонов цены и даты
        self.min_price = self.entry_min_price.get()
        self.max_price = self.entry_max_price.get()

        print(self.collections)

        self.col1_state = self.col11.get()
        self.col2_state = self.col22.get()
        self.col3_state = self.col33.get()
        self.col4_state = self.col44.get()
        self.col5_state = self.col55.get()
        self.col6_state = self.col66.get()
        self.col7_state = 0
        self.col8_state = 0
        if len(self.collections)==7:
            self.col7_state = self.col77.get()
        elif len(self.collections)==8:
            self.col7_state = self.col77.get()
            self.col8_state = self.col88.get()

        self.onduty=[self.col1_state, self.col2_state,self.col3_state,self.col4_state,self.col5_state,
                     self.col6_state,self.col7_state,self.col8_state]

        error = " "

        print(self.onduty)

        if sum(self.onduty)==0:
            error+="Выберите одну из категорий!"+"\n"

        if len(self.min_price) == 0 and len(self.max_price) != 0 or len(self.min_price) != 0 and len(self.max_price) == 0:
            error+="Либо оба поля цен заполнены, либо оба пустые"+"\n"

        if len(self.min_price) == 0 and len(self.max_price) == 0:
            pass
        elif len(self.min_price) != 0 and len(self.max_price) != 0:
            if not self.checkf(self.min_price) or float(self.min_price) <= 0:
                error += "Цена 1 должна быть числом целым или вещественным и большим нуля" + "\n"
            if len(self.min_price) > 15:
                error += "Цена 1 слишком большая" + "\n"

            if not self.checkf(self.max_price) or float(self.max_price) <= 0:
                error += "Цена 2 должна быть числом целым или вещественным и большим нуля" + "\n"
            if len(self.max_price) > 15:
                error += "Цена 2 слишком большая" + "\n"

            if error == " ":
                self.min_price = float(self.min_price)
                self.max_price = float(self.max_price)
                if self.min_price>self.max_price:
                    self.min_price, self.max_price = self.max_price, self.min_price

        self.pressed_period=self.chk_state2.get()

        if self.pressed_period==1:
            self.min_date = self.entry_min_date.get()
            self.max_date = self.entry_max_date.get()

            dt = True
            try:
                datetime.strptime(self.min_date, "%Y-%m-%d")
                datestr1 = self.max_date.split('-')
                if len(datestr1[-2]) < 2 and len(datestr1[-1]) < 2:
                    dt = False
                else:
                    dt = True
            except ValueError:
                dt = False
            if dt:
                if int(self.min_date[0:4]) < 1991 or int(self.min_date[0:4]) > 2056:
                    dt = False
                else:
                    dt = True
            if not dt:
                error += "Неправильно указана дата 1 (минимальный год - 1991, максимальный - 2056)" + "\n"

            dt = True
            try:
                datetime.strptime(self.max_date, "%Y-%m-%d")
                datestr2 = self.max_date.split('-')
                if len(datestr2[-2]) < 2 and len(datestr2[-1]) < 2:
                    dt = False
                else:
                    dt = True
            except ValueError:
                dt = False
            if dt:
                if int(self.max_date[0:4]) < 1991 or int(self.max_date[0:4]) > 2056:
                    dt = False
                else:
                    dt = True
            if not dt:
                error += "Неправильно указана дата 2 (минимальный год - 1991, максимальный - 2056)" + "\n"

            if self.min_date>self.max_date:
                self.min_date, self.max_date = self.max_date, self.min_date

        elif self.pressed_period==2:

            self.min_date='1991-01-01'
            self.max_date='2056-12-31'

        if error == " " and len(str(self.min_price)) != 0 and len(str(self.max_price)) != 0:
            filtered_journal = []

            kt=0
            for entry in self.journal:
                price = float(entry[2])
                date = entry[3]
                collection = entry[4]
                if self.min_price <= price <= self.max_price and self.min_date <= date <= self.max_date and self.onduty[self.collections.index(collection)]!=0:
                    filtered_journal.append(entry)
                    kt += price

            self.lbl5.destroy()
            self.lbl5 = Label(root, text="Сумма: " + str(kt), font=("Calibri", 16), foreground="white",
                              background="#003200")
            self.lbl5.place(relx=0.1, rely=0.9, anchor='center')

            # Очищаем текущее содержимое Treeview
            for item in self.tree.get_children():
                self.tree.delete(item)

            # Заполняем Treeview отфильтрованными данными
            for entry in filtered_journal:
                self.tree.insert("", END, values=entry)

        elif error == " " and len(self.min_price) == 0 and len(self.max_price)==0:
            filtered_journal = []
            kt=0
            for entry in self.journal:
                date = entry[3]
                price = float(entry[2])
                collection = entry[4]
                if self.min_date <= date <= self.max_date and self.onduty[self.collections.index(collection)]!=0:
                    filtered_journal.append(entry)
                    kt+=price

            # Очищаем текущее содержимое Treeview
            for item in self.tree.get_children():
               self.tree.delete(item)

            # Заполняем Treeview отфильтрованными данными
            for entry in filtered_journal:
                self.tree.insert("", END, values=entry)

            self.lbl5.destroy()
            self.lbl5 = Label(root, text="Сумма: " + str(kt), font=("Calibri", 16), foreground="white",
                              background="#003200")
            self.lbl5.place(relx=0.1, rely=0.9, anchor='center')

        else:
            messagebox.showinfo("Возникла ошибка:(",error)
            '''self.filter.destroy()
            for widget in self.root.winfo_children():
                widget.destroy()
            self.menu_page()'''

        print(self.journal)
        '''selected_item = self.tree.selection()
        print(selected_item)
        if selected_item:
            confirm = messagebox.askyesno("Подтверждение удаления", "Вы уверены, что хотите удалить выбранную позицию?")
            if confirm:
                for item in selected_item:
                    self.tree.delete(item)'''

    def clear_filter(self):
        self.entry_min_price.delete(0, END)
        self.entry_max_price.delete(0, END)
        self.col11.set(0)
        self.col22.set(0)
        self.col33.set(0)
        self.col44.set(0)
        self.col55.set(0)
        self.col66.set(0)
        self.col77.set(0)
        self.col88.set(0)
        '''for widget in self.root.winfo_children():
            widget.destroy()
        self.menu_page()'''

    def close_filter(self):
        self.filter.destroy()
        for widget in self.root.winfo_children():
            widget.destroy()
        self.menu_page()

    def summa(self):
        cnt2=0
        n= '-' * 46
        if self.journal == [[n, n, n, n, n]] or self.journal == [[]] or self.journal == []:
            pass
        else:
            for k in range(len(self.journal)):
                cnt2+=float(self.journal[k][2])
        return cnt2

    def sort_column(self, tree, col, reverse):
        data=[]
        for child in tree.get_children(''):
            if col == 'Цена':
                data.append((float(tree.set(child, col)), child))
            elif col == '№':
                data.append((int(tree.set(child, col)), child))
            else:
                data.append((tree.set(child, col), child))
        #print(data)

        data.sort(reverse=reverse, )
        #print(data)

        for index, item in enumerate(data):
            tree.move(item[1], '', index)

        # Меняем направление сортировки при следующем щелчке
        tree.heading(col, command=lambda _col=col: self.sort_column(tree, _col, not reverse))

    def popup(self, event):
        self.popup_menu.post(event.x_root, event.y_root)

    def delete_item(self):
        selected_item = self.tree.selection()
        n = '-' *46
        if self.journal != [[n,n,n,n,n]] and self.journal != [] and self.journal != [[]]:
            if selected_item:
                confirm = messagebox.askyesno("Подтверждение удаления", "Вы уверены, что хотите удалить выбранную позицию?")
                if confirm:
                    try:
                        print(selected_item)
                        ind = int(selected_item[0].replace('I',''), 16) - 1
                        deleted = self.journal.pop(ind) #ind
                        self.tree.delete(selected_item)
                        for widget in self.root.winfo_children():
                            widget.destroy()
                        if self.journal != []:
                            q = 1
                            for i in range(len(self.journal)):
                                self.journal[i][0] = int(q + i)
                            self.num = int(self.journal[-1][0]) + 1
                            print(self.num)

                        find = False
                        for el in self.journal:
                            if deleted[-1] == el[-1]:
                                find = True
                        if not find:
                            self.collections.remove(deleted[-1])
                    except IndexError:# and ValueError:
                        messagebox.showinfo("Уведомление", "К сожалению, функция удаления элементов в отфильтрованной таблице находится в разработке")
                    '''except ValueError:
                        messagebox.showinfo("Уведомление", "К сожалению, функция удаления элементов в отфильтрованной таблице находится в разработке")'''

                    for widget in self.root.winfo_children():
                        widget.destroy()
                    self.menu_page()
                else:
                    pass
        else:
            messagebox.showinfo("Уведомление",
                                "Таблица пустая!")

    def clear_table(self):
        confirm = True
        n = '-' * 46
        if self.journal == [[n, n, n, n, n]] or self.journal == [[]] or self.journal == []:
            confirm = False
        if confirm:
            confirm = messagebox.askyesno("Удаление данных",
                                          "Вы уверены, что хотите очистить форму? Это приведет к удалению всех "
                                          "данных (Данные при следующем запуске не сохранятся).")
            if confirm:
                for widget in self.root.winfo_children():
                    widget.destroy()
                self.journal=[[]]
                self.num=1
                self.menu_page()

    def add_product(self):
        self.top = Toplevel(self.root)
        self.top.title("Добавление продукта")
        self.top.geometry("450x680")
        self.top.minsize(450, 680)
        self.top.maxsize(550, 780)

        self.top.configure(background="#003200")

        self.lbl_product = Label(self.top, text="Название продукта:", font=("Calibri", 14), foreground="white",
                                 background="#003200")
        self.lbl_product.place(relx=0.1, rely=0.1)

        self.entry_product = Entry(self.top, width=28)
        self.entry_product.place(relx=0.5, rely=0.11)

        self.lbl_product2 = Label(self.top, text="Цена:", font=("Calibri", 14), foreground="white",
                                  background="#003200")
        self.lbl_product2.place(relx=0.1, rely=0.2)

        self.entry_product2 = Entry(self.top, width=28)
        self.entry_product2.place(relx=0.5, rely=0.21)

        self.lbl_product2 = Label(self.top, text="Цена:", font=("Calibri", 14), foreground="white",
                                  background="#003200")
        self.lbl_product2.place(relx=0.1, rely=0.2)

        self.lbl_product3 = Label(self.top, text="Дата покупки:", font=("Calibri", 14), foreground="white",
                                  background="#003200")
        self.lbl_product3.place(relx=0.1, rely=0.3)

        today = date.today()
        self.cal = DateEntry(self.top, width=28, background='#319158', foreground='white', borderwidth=2,
                             date_pattern="yyyy-mm-dd")
        self.cal.place(relx=0.5, rely=0.31)
        self.cal.set_date(today)

        self.lbl_product4 = Label(self.top, text="Выберите коллекцию", font=("Calibri", 14), foreground="white",
                                  background="#003200")
        self.lbl_product4.place(relx=0.5, rely=0.4, anchor='center')

        self.chk_state = IntVar(value=0)

        self.chk = Radiobutton(self.top, text='Супермаркеты', variable=self.chk_state, value=1, font=("Calibri", 12))
        self.chk.place(relx=0.1, rely=0.45)

        self.chk2 = Radiobutton(self.top, text='Рестораны и кафе', variable=self.chk_state, value=2,
                                font=("Calibri", 12))
        self.chk2.place(relx=0.6, rely=0.45)

        self.chk3 = Radiobutton(self.top, text='Здоровье и красота', variable=self.chk_state, value=3,
                                font=("Calibri", 12))
        self.chk3.place(relx=0.1, rely=0.5)

        self.chk4 = Radiobutton(self.top, text='Транспорт', variable=self.chk_state, value=4,
                                font=("Calibri", 12))
        self.chk4.place(relx=0.6, rely=0.5)

        self.chk5 = Radiobutton(self.top, text='Развлечения и хобби', variable=self.chk_state, value=5,
                                font=("Calibri", 12))
        self.chk5.place(relx=0.1, rely=0.55)

        self.chk6 = Radiobutton(self.top, text='Переводы', variable=self.chk_state, value=6,
                                font=("Calibri", 12))
        self.chk6.place(relx=0.6, rely=0.55)

        self.chk7 = Radiobutton(self.top, text='Другое:', font=("Calibri", 12), variable=self.chk_state, value=7)
        self.chk7.place(relx=0.1, rely=0.65)

        self.warn = Label(self.top, text="Внимание, вы можете добавить не более 2-х коллекций!", font=("Calibri", 10), foreground="white",
                                  background="#003200")
        self.warn.place(relx=0.1, rely=0.7)

        self.entry_category = Entry(self.top, width=40)
        self.entry_category.place(relx=0.3, rely=0.657)

        self.btn3 = Button(self.top, text="Сохранить", font=("Calibri", 16), background="#319158", foreground="white",
                           command=self.save_product)
        self.btn3.place(relx=0.5, rely=0.8, anchor="center")

    def checkf(self, k):  # проверка на вещественное число
        flag = True
        try:
            k = float(k)
        except ValueError:
            flag = False
        return 1 if flag else 0

    def save_product(self):
        self.product_name = self.entry_product.get()
        self.product_cost = self.entry_product2.get()
        self.product_date = self.cal.get()
        self.product_category = self.chk_state.get()

        error = " "

        if not self.checkf(self.product_cost) or float(self.product_cost) <= 0:
            error += "Цена должна быть числом целым или вещественным и большим нуля" + "\n"
        if len(self.product_cost) > 15:
            error += "Цена слишком большая" + "\n"

        if error == " ":
            self.product_cost = float(self.product_cost)

        if len(self.product_name) == 0 or len(self.product_name) > 28:
            error += "Название товара не должно быть пустым и должно состоять не более чем из 28 символов" + "\n"
        else:
            self.product_name=self.product_name[0].upper()+self.product_name[1::]

        dt = True
        try:
            datetime.strptime(self.product_date, "%Y-%m-%d")
            datestr=self.product_date.split('-')
            if len(datestr[-2]) < 2 and len(datestr[-1]) < 2:
                dt=False
            else:
                dt = True
        except ValueError:
            dt = False
        if dt:
            if int(self.product_date[0:4]) < 1991 or int(self.product_date[0:4]) > 2056:
                dt = False
            else:
                dt = True
        if not dt:
            error += "Неправильно указана дата (минимальный год - 1991, максимальный - 2056)" + "\n"

        if self.product_category == 1:
            self.product_category = 'Супермаркеты'
        if self.product_category == 2:
            self.product_category = 'Рестораны и кафе'
        if self.product_category == 3:
            self.product_category = 'Здоровье и красота'
        if self.product_category == 4:
            self.product_category = 'Транспорт'
        if self.product_category == 5:
            self.product_category = 'Развлечения и хобби'
        if self.product_category == 6:
            self.product_category = 'Переводы'
        if self.product_category == 7:
            self.product_category = self.entry_category.get()
            if len(self.product_category) == 0 or len(self.product_category) > 28:
                error += "Название коллекции не должно быть пустым и должно состоять не более чем из 28 символов" + "\n"
            else:
                self.product_category = self.product_category[0].upper() + self.product_category[1::]
                if self.product_category in self.collections:
                    pass
                else:
                    self.collections.append(self.product_category)

        if self.product_category == 0:
            error+='Выберите коллекцию (нажать кнопку)'+"\n"
        if len(self.collections)>8:
            error+="Превышен лимит коллекций!"+"\n"
            self.collections.pop(-1)

        if error != " ":
            messagebox.showinfo("Возникла ошибка:(", error)
            self.top.destroy()

        else:
            self.add_to_table()

    def add_to_table(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        n = '-' * 46
        if self.journal == [[n, n, n, n, n]]:
            self.journal.pop(0)
        self.journal.append([self.num, self.product_name, self.product_cost, self.product_date, self.product_category])
        self.num += 1
        self.menu_page()

    def connect_to_db(self):
        conn = sqlite3.connect('finance.db')
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS finance_list
                         (id INTEGER PRIMARY KEY AUTOINCREMENT,
                          name TEXT,
                          price REAL,
                          purchase_date TEXT,
                          category TEXT,
                          username TEXT,
                          session TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
        conn.commit()
        conn.close()

    def save_into_db(self):
        conn = sqlite3.connect('finance.db')
        c = conn.cursor()
        for el in self.journal:
            c.execute('''INSERT INTO finance_list (name, price, purchase_date, category, username) VALUES (?, ?, ?, 
            ?, ?)''',
                      (el[1], el[2], el[3], el[4], self.username))
        conn.commit()
        conn.close()

    def exit(self):
        n = '-' * 46
        if self.flag or self.journal == [[n, n, n, n, n]] or self.journal==[[]] or self.journal==[]:
            pass
        else:
            confirm = messagebox.askyesno("Подтверждение выхода", "У вас есть несохраненные данные. Сохранить перед выходом?")
            if confirm:
                self.save()
            elif not confirm and (self.journal != [[n, n, n, n, n]] or self.journal!=[[]] or self.journal!=[]):
                self.journal=self.load
                self.save_into_db()
        sys.exit(0)

    def info(self):
        messagebox.showinfo("Информация", "Эта программа предназначена для учета личных финансов.")

    def save(self):
        n = '-' * 46
        if self.journal == [[n, n, n, n, n]] or self.journal==[[]] or self.journal==[]:
            messagebox.showinfo("Уведомление", "Данные не сохранены, таблица пустая.")
        else:
            # Запись данных в CSV файл
            filename = os.path.join(os.path.expanduser("~"), "Desktop", "finances.csv")
            with open(filename, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerows(self.journal)

            # Сохранение в Excel (.xlsx)
            filename_xlsx = os.path.join(os.path.expanduser("~"), "Desktop", "finances.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(self.tor)
            for row in self.journal:
                ws.append(row)
            wb.save(filename_xlsx)

            messagebox.showinfo("Сохранение", f"Данные успешно сохранены в файл: {filename}")
            self.flag=True
            self.save_into_db()

    def load_from_db(self):
        conn = sqlite3.connect('finance.db')
        c = conn.cursor()
        c.execute('''SELECT * FROM finance_list WHERE username=?''', (self.username,))
        data = c.fetchall()
        conn.commit()
        conn.close()
        return data

    def delete_from_db(self, user):
        conn = sqlite3.connect('finance.db')
        c = conn.cursor()
        c.execute('''DELETE FROM finance_list WHERE username=?''', (user,))
        c.execute('''DELETE FROM sqlite_sequence WHERE name="finance_list"''')
        conn.commit()
        conn.close()
        self.reorder_ids()

    def reorder_ids(self):
        conn = sqlite3.connect('finance.db')
        c = conn.cursor()
        # Получаем все записи, отсортированные по id
        c.execute('''SELECT id FROM finance_list ORDER BY id''')
        ids = c.fetchall()
        # Перезаписываем id, начиная с 1
        for index, (old_id,) in enumerate(ids, start=1):
            c.execute('''UPDATE finance_list SET id=? WHERE id=?''', (index, old_id))
        conn.commit()
        conn.close()

    def on_closing(self):
        self.exit()

    def on_closing2(self):
        self.filter.destroy()
        for widget in self.root.winfo_children():
            widget.destroy()
        self.menu_page()

# Создание основного окна приложения и его запуск
root = tk.Tk()
app = App(root)
root.mainloop()
