import tkinter as tk
from person import Person
from tkcalendar import DateEntry
import datetime
import re
from query import create_worker, check_person, deleteEmployee, search_worker_prof, search_all_worker_prof, change_department_person
from tkinter import messagebox, ttk, END
import mysql.connector
from docx import Document

profession = []


def main():
    root = tk.Tk()
    root.title("Dister")
    root.geometry("650x400")

    tab_control = ttk.Notebook(root)
    tab1 = tk.Frame(tab_control)
    tab2 = tk.Frame(tab_control)
    tab3 = tk.Frame(tab_control)

    #Создание вкладок в окне
    tab_control.add(tab1, text='Добавление сотрудника')
    tab_control.add(tab2, text='Удаление')
    tab_control.add(tab3, text="Таблица")
    tab_control.pack(expand=1, fill='both')



    def updateTable():
        mydb = mysql.connector.connect(
            host="localhost",
            user="root",
            password="root",
            database="practica",
            port='3306'
        )

        mycursor = mydb.cursor()

        mycursor.execute("SELECT * FROM person")
        updateResult = mycursor.fetchall()

        for i in tree.get_children():
            tree.delete(i)
        for row in updateResult:
            tree.insert("", END, values=(row[1], row[2], row[3], row[4], row[5], row[6], row[7]))

        mydb.close()
        mycursor.close()

    def create_profession_fields():
        def destroy_popup():
            popup.destroy()

        def comands_retrieve_button():
            retrieve_data()
            destroy_popup()

        popup = tk.Toplevel(root)
        popup.title("Popup Window")
        popup.geometry("450x200")
        num_professions = int(selected_professions.get())

        profession_frames = tk.Frame(popup)
        profession_frames.pack()

        entry_lists.clear()
        for i in range(num_professions):
            profession_frame = tk.Frame(profession_frames)
            profession_frame.pack(fill=tk.X, padx=5, pady=5,)

            label = tk.Label(profession_frame, text=f"Доп. проф. {i + 1}:")
            label.pack(side=tk.LEFT)

            entry_prof = tk.Entry(profession_frame, width=30, validate="key", validatecommand=(validate_entry40, "%P"))
            entry_prof.pack(side=tk.LEFT, padx=5)

            label_exp = tk.Label(profession_frame, text=f"Стаж {i + 1}:")
            label_exp.pack(side=tk.LEFT)

            entry_exp = tk.Entry(profession_frame, width=30, validate="key", validatecommand=(validate_entry2, "%P"))
            entry_exp.pack(side=tk.LEFT, padx=5)

            entry_lists.append(entry_prof)
            entry_lists.append(entry_exp)

        retrieve_button = tk.Button(popup, text="Получить данные", command=comands_retrieve_button)
        retrieve_button.pack()
        popup.grab_set()


    def retrieve_data():
        for i, entry in enumerate(entry_lists):
            profession_info = entry.get()
            profession.append(profession_info)

        return profession


    def validate_entry_40(P):
        return len(P) <= 40 and (bool(re.match("^[а-яА-ЯёЁ ]+$", P)) or P == "")


    def validate_entry_20(P):
        return len(P) <= 20 and (bool(re.match("^[а-яА-ЯёЁ ]+$", P)) or P == "")


    def validate_entry_2(P):
        return (len(P) <= 2 and P.isdigit()) or P == ""

    def change_department():
        ok = True
        department = entry_department.get()
        full_name = entry_FIO.get()

        if len(department) == 0 or len(full_name) == 0:
            ok = False

        if ok == True:
            if (check_person(full_name)):
                change_department_person(full_name, department)
                correct_label1.config(text="Отдел изменен", fg="blue")
                updateTable()
            else:
                correct_label1.config(text="Такого сотрудника нет", fg="red")



    def on_button_click():
        ok = True
        worker = Person
        worker.full_name = entry_FIO.get()
        worker.Date = entry_birthday.get()
        worker.departament = entry_department.get()
        worker.family_status = selected_family_status.get()
        worker.position = entry_position.get()
        worker.work_exp = entry_experience.get()

        if (len(worker.full_name) < 1) or (len(worker.departament) < 1) or (len(worker.family_status) < 1) or (len(worker.position) < 1) or (len(worker.work_exp) < 1):
            ok = False

        if ok == True:
            if not (check_person(worker.full_name)):
                entry_FIO.delete(0, tk.END)
                time_now = datetime.date.today()
                eighteen_years_ago = time_now - datetime.timedelta(days=18 * 365 + 4)
                entry_birthday.set_date(eighteen_years_ago)
                entry_department.delete(0, tk.END)
                entry_position.delete(0, tk.END)
                entry_experience.delete(0, tk.END)
                correct_label1.config(text="Сотрудник добавлен", fg="blue")
                create_worker(worker, profession)
                profession.clear()
                updateTable()
            else:
                correct_label1.config(text="Такой сотрудник уже есть", fg="red")

    validate_entry40 = tab1.register(validate_entry_40)
    validate_entry20 = tab1.register(validate_entry_20)
    validate_entry2 = tab1.register(validate_entry_2)

    #фио
    label_FIO = tk.Label(tab1, text="ФИО:", anchor="w")
    label_FIO.grid(row=0, column=0, padx=5, pady=7)

    entry_FIO = tk.Entry(tab1, width=40, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_FIO.grid(row=0, column=1, padx=5, pady=7)

    #др
    label_birthday = tk.Label(tab1, text="Дата рождения:", anchor="w")
    label_birthday.grid(row=1, column=0, padx=5, pady=7)

    time_now = datetime.date.today().year

    entry_birthday = DateEntry(tab1, width=35, date_pattern='yyyy.mm.dd', locale='ru_RU', year=time_now-18, state='readonly')
    entry_birthday.grid(row=1, column=1, padx=5, pady=7)

    #отдел
    label_department = tk.Label(tab1, text="Наименование отдела:", anchor="w")
    label_department.grid(row=2, column=0, padx=5, pady=7)

    entry_department = tk.Entry(tab1, width=35, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_department.grid(row=2, column=1, padx=5, pady=7)

    #семейный статус
    selected_family_status = tk.StringVar(tab1)
    selected_family_status.set("Холост/Не замужем")

    label_family_status = tk.Label(tab1, text="Семеный статус:", anchor="w")
    label_family_status.grid(row=3, column=0, padx=5, pady=7)

    option_menu_family_status = tk.OptionMenu(tab1, selected_family_status, "Женат/Замужем", "Холост/Не замужем")
    option_menu_family_status.grid(row=3, column=1, padx=5, pady=7)

    #профессия
    label_position = tk.Label(tab1, text="Действующая професия:", anchor="w")
    label_position.grid(row=4, column=0, padx=5, pady=7)

    entry_position = tk.Entry(tab1, width=35, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_position.grid(row=4, column=1, padx=5, pady=7)

    #стаж
    label_experience = tk.Label(tab1, text="Стаж работы:", anchor="w")
    label_experience.grid(row=6, column=0, padx=5, pady=7)

    entry_experience = tk.Entry(tab1, width=35, validate="key", validatecommand=(validate_entry2, "%P"))
    entry_experience.grid(row=6, column=1, padx=5, pady=7)

    correct_label1 = tk.Label(tab1)
    correct_label1.grid(row=8, column=0, columnspan=2, pady=5)

    ##################################33
    selected_professions = tk.StringVar(tab1)
    selected_professions.set("1")

    label_info_prof = tk.Label(tab1, text="Выберите количество доп профессий (от 1 до 4):")
    label_info_prof.grid(row=5, column=0, padx=5, pady=7)

    option_menu = tk.OptionMenu(tab1, selected_professions, "1", "2", "3", "4")
    option_menu.grid(row=5, column=2, padx=5, pady=7)

    submit_button = tk.Button(tab1, text="Добавить профессии", command=create_profession_fields)
    submit_button.grid(row=5, column=1, padx=5, pady=5)

    entry_lists = []  # Список для хранения ссылок на Entry для каждой профессии
    ########################################3

    button = tk.Button(tab1, text="Добавить сотрудника", command=on_button_click)
    button.grid(row=7, columnspan=2, padx=5, pady=5)

    button_change_department = tk.Button(tab1, text="Измененение отдела у сотрудника", command=change_department)
    button_change_department.place(x=350, y=269.5)


    ########################## удаление
    def button_delete():
        full_name = deleteEntry.get()
        check = check_person(full_name)
        if check == True:
            deleteEntry.delete(0, tk.END)
            deleteEmployee(full_name)
            correct_label.config(text="Сотрудник удалён", fg="blue")
            updateTable()
        else:
            correct_label.config(text="Сотрудник не найден",fg="red")

    deleteLabel = tk.Label(tab2, text="Введите ФИО сотрудника которого вы хотите уволить")
    deleteLabel.grid(row=0, column=0)
    deleteEntry = tk.Entry(tab2)
    deleteEntry.grid(row=0, column=1, padx=15, pady=5)


    deleteButton = tk.Button(tab2, text="Уволить сотрудника",command=button_delete)
    deleteButton.grid(row=1, column=0, ipadx=10, ipady=7, pady=10)

    correct_label = tk.Label(tab2)
    correct_label.grid(row=2, column=0, columnspan=2, pady=5)

    ################## Таблица

    def print_worker_prof():
        popup = tk.Toplevel(root)
        popup.title("Popup Window")
        popup.geometry("450x200")

        tree_worker_prof = ttk.Treeview(popup)
        tree_worker_prof.pack(fill='both', expand=True)

        data = search_worker_prof()

        tree_worker_prof["columns"] = ("ФИО", "Количество профессий")
        tree_worker_prof.column("#0", width=0, stretch=tk.NO)  # Пустая колонка для удобства

        tree_worker_prof.heading("ФИО", text="ФИО")
        tree_worker_prof.heading("Количество профессий", text="Количество профессий")

        for row in data:
            tree_worker_prof.insert('', tk.END, values=row)

        makeFile(data)


    def print_family_status():
        family_status = selected_family_status_table.get()
        position = entry_prof.get()

        popup = tk.Toplevel(root)
        popup.title("Popup Window")
        popup.geometry("450x200")

        tree_worker_prof = ttk.Treeview(popup)
        tree_worker_prof.pack(fill='both', expand=True)

        data = search_all_worker_prof(family_status, position)

        tree_worker_prof["columns"] = ("ФИО", "Стаж")
        tree_worker_prof.column("#0", width=0, stretch=tk.NO)

        tree_worker_prof.heading("ФИО", text="ФИО")
        tree_worker_prof.heading("Стаж", text="Стаж")

        for row in data:
            tree_worker_prof.insert('', tk.END, values=row)


    def makeFile(data):
        doc = Document()

        # Создаем таблицу с двумя столбцами
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'  # Устанавливаем стиль таблицы

        # Добавляем заголовки столбцов
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'FIO'
        hdr_cells[1].text = 'experience'

        # Добавляем данные в таблицу
        for row in data:
            new_row = table.add_row().cells
            new_row[0].text = str(row[0]) if len(row) > 0 else ''  # Предполагается, что данные хранятся в первом столбце
            new_row[1].text = str(row[1]) if len(row) > 1 else ''  # Предполагается, что данные хранятся во втором столбце

        # Сохраняем документ
        doc.save('WorkersTable.docx')
        messagebox.showinfo('Успех', "Файл WorkersTable.docx успешно сформирован")



    mydb = mysql.connector.connect(
        host="localhost",
        user="root",
        password="root",
        database="practica",
        port='3306'
    )

    mycursor = mydb.cursor()

    mycursor.execute("SELECT * FROM person")

    result = mycursor.fetchall()

    tree = ttk.Treeview(tab3)
    tree["columns"] = ("1", "2", "3", "4", "5", "6", "7")
    tree["show"] = "headings"

    tree.heading("1", text="ФИО")
    tree.heading("2", text="Дата рождения")
    tree.heading("3", text="Наименование отдела")
    tree.heading("4", text="Семейное положение")
    tree.heading("5", text="Текущая профессия")
    tree.heading("6", text="Стаж работы")
    tree.heading("7", text="Количество профессий")

    tree.column("1", width=70)
    tree.column("2", width=70)
    tree.column("3", width=70)
    tree.column("4", width=70)
    tree.column("5", width=70)
    tree.column("6", width=70)
    tree.column("7", width=70)

    for row in result:
        tree.insert("", END, values=(row[1], row[2], row[3], row[4], row[5], row[6], row[7]))

    tree.pack()
    mycursor.close()
    mydb.close()

    print_button_1 = tk.Button(tab3, text="Вывод", command=print_family_status)
    print_button_1.pack(side=tk.LEFT, padx=10, pady=20)

    selected_family_status_table = tk.StringVar(tab3)
    selected_family_status_table.set("Холост/Не замужем")

    option_menu_family_status = tk.OptionMenu(tab3, selected_family_status_table, "Женат/Замужем", "Холост/Не замужем")
    option_menu_family_status.pack(side=tk.LEFT, padx=10, pady=20)

    deleteLabel = tk.Label(tab3, text="Введите профеcсию сотрудников")
    deleteLabel.pack(side=tk.LEFT, padx=10, pady=20)

    entry_prof = tk.Entry(tab3)
    entry_prof.pack(side=tk.LEFT, padx=10, pady=20)


    print_button_2 = tk.Button(tab3, text="Работник/Количество проффесий",command=print_worker_prof)
    print_button_2.place(x=10, y=325)


    root.resizable(0,0)
    root.mainloop()