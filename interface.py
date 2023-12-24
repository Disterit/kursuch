import tkinter as tk
from person import Person
from tkcalendar import DateEntry
import datetime
import re
from query import *
from tkinter import messagebox, ttk, END
from docx import Document

profession = []

def main_interface():
    root = tk.Tk()
    root.title("Dister")
    root.geometry("600x350")

    tab_control = ttk.Notebook(root)
    tab1 = ttk.Frame(tab_control)
    tab2 = ttk.Frame(tab_control)
    tab3 = ttk.Frame(tab_control)

    tab_control.add(tab1, text='Добавление сотрудника')
    tab_control.add(tab2, text='Удаление')
    tab_control.add(tab3, text="Таблица")
    tab_control.pack(expand=1, fill='both')

    tab_append(tab1, root, tab_control)
    tab_delete(tab2)
    tab_table(tab3, root)

    root.resizable(0, 0)
    root.mainloop()


def tab_table(tab3, root):
    validate_entry40 = tab3.register(validate_entry_40)

    table = create_table()

    global tree

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

    for row in table:
        tree.insert("", END, values=(row[1], row[2], row[3], row[4], row[5], row[6], row[7]))

    tree.pack(pady=5)

    print_button_1 = ttk.Button(tab3, text="Вывод",
                                command=lambda: print_family_status(root, entry_prof, selected_family_status_table))
    print_button_1.place(x=35, y=250)

    selected_family_status_table = tk.StringVar(tab3)
    selected_family_status_table.set("Холост/Не замужем")

    option_menu_family_status = ttk.OptionMenu(tab3, selected_family_status_table, "Женат/Замужем", "Женат/Замужем",
                                               "Холост/Не замужем", style='Select.TButton')
    option_menu_family_status.place(x=125, y=250)

    deleteLabel = ttk.Label(tab3, text="Введите профеcсию сотрудников")
    deleteLabel.place(x=235, y=252)

    entry_prof = ttk.Entry(tab3, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_prof.place(x=430, y=252)

    print_button_2 = ttk.Button(tab3, text="Работник/Количество проффесий", command=lambda: print_worker_prof(root),
                                style='Select.TButton')
    print_button_2.place(x=35, y=285)

    professions_label = ttk.Button(tab3, text="Узнать профессии сотрудника",
                                   command=lambda: find_prof_person(root, entry_prof_2), style='Select.TButton')
    professions_label.place(x=245, y=285)

    entry_prof_2 = ttk.Entry(tab3, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_prof_2.place(x=430, y=287)


def destroy_popup(popup):
        popup.destroy()


def create_profession_fields(entry_lists, root, selected_professions, validate_entry40, validate_entry2):

    def comands_retrieve_button():
        retrieve_data(entry_lists)
        destroy_popup(popup)

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

        label = ttk.Label(profession_frame, text=f"Доп. проф. {i + 1}:")
        label.pack(side=tk.LEFT)

        entry_prof = ttk.Entry(profession_frame, width=30, validate="key", validatecommand=(validate_entry40, "%P"))
        entry_prof.pack(side=tk.LEFT, padx=5)

        label_exp = ttk.Label(profession_frame, text=f"Стаж {i + 1}:")
        label_exp.pack(side=tk.LEFT)

        entry_exp = ttk.Entry(profession_frame, width=30, validate="key", validatecommand=(validate_entry2, "%P"))
        entry_exp.pack(side=tk.LEFT, padx=5)

        entry_lists.append(entry_prof)
        entry_lists.append(entry_exp)

    retrieve_button = ttk.Button(popup, text="Получить данные", command=comands_retrieve_button)
    retrieve_button.pack()
    popup.grab_set()


def retrieve_data(entry_lists):
    for i, entry in enumerate(entry_lists):
        profession_info = entry.get()
        profession.append(profession_info)

    return profession


def validate_entry_40(P):
    return len(P) <= 40 and (bool(re.match("^[а-яА-ЯёЁ ]+$", P)) or P == "")


def validate_entry_2(P):
    return (len(P) <= 2 and P.isdigit()) or P == ""


def change_department(root):
    popup = tk.Toplevel(root)
    popup.title("Popup Window")
    popup.geometry("200x150")

    validate_entry40 = popup.register(validate_entry_40)

    label_FIO = ttk.Label(popup, text="Введите ФИО:")
    label_FIO.pack()

    entry_FIO = ttk.Entry(popup, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_FIO.pack()

    label_department = ttk.Label(popup, text="Введите отдел:")
    label_department.pack()

    entry_department = ttk.Entry(popup, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_department.pack()

    submit_button = ttk.Button(popup, text="Изменить отдел", command=lambda: change_department_button(entry_department, entry_FIO, popup))
    submit_button.pack(pady=5)
    popup.grab_set()


def change_department_button(entry_department, entry_FIO, popup):
    ok = True
    department = entry_department.get()
    full_name = entry_FIO.get()

    if len(department) == 0 or len(full_name) == 0:
        ok = False

    if ok == True:
        if (check_person(full_name)):
            change_department_person(full_name, department)
            messagebox.showinfo("Success", "Отдел изменен")
            destroy_popup(popup)
            updateTable(tree)
        else:
            messagebox.showerror("Error", "Такого сотрудника нет")


def on_button_click(entry_FIO, entry_birthday, entry_department, selected_family_status, entry_position, entry_experience):
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
            messagebox.showinfo("Success", "Сотрудник добавлен")
            create_worker(worker, profession)
            profession.clear()
            updateTable(tree)
        else:
            messagebox.showerror("Error", "Такой сотрудник уже есть")


def tab_append(tab1, root, tab_control):

    validate_entry40 = tab1.register(validate_entry_40)
    validate_entry2 = tab1.register(validate_entry_2)

    #фио
    label_FIO = tk.Label(tab1, text="ФИО:", anchor="w")
    label_FIO.grid(row=0, column=0, padx=5, pady=7)

    entry_FIO = ttk.Entry(tab1, width=35, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_FIO.grid(row=0, column=1, padx=5, pady=7)

    #др
    label_birthday = tk.Label(tab1, text="Дата рождения:", anchor="w")
    label_birthday.grid(row=1, column=0, padx=5, pady=7)

    time_now = datetime.date.today().year

    entry_birthday = DateEntry(tab1, width=33, date_pattern='yyyy.mm.dd', locale='ru_RU', year=time_now-18, state='readonly')
    entry_birthday.grid(row=1, column=1, padx=5, pady=7)

    #отдел
    label_department = tk.Label(tab1, text="Наименование отдела:", anchor="w")
    label_department.grid(row=2, column=0, padx=5, pady=7)

    entry_department = ttk.Entry(tab1, width=35, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_department.grid(row=2, column=1, padx=5, pady=7)

    #семейный статус
    selected_family_status = tk.StringVar(tab1)
    selected_family_status.set("Холост/Не замужем")

    label_family_status = tk.Label(tab1, text="Семеный статус:", anchor="w")
    label_family_status.grid(row=3, column=0, padx=5, pady=7)

    option_menu_family_status = ttk.OptionMenu(tab1, selected_family_status, "Женат/Замужем", "Холост/Не замужем", style='Select.TButton')
    option_menu_family_status.grid(row=3, column=1, padx=5, pady=7)

    #профессия
    label_position = tk.Label(tab1, text="Действующая професия:", anchor="w")
    label_position.grid(row=4, column=0, padx=5, pady=7)

    entry_position = ttk.Entry(tab1, width=35, validate="key", validatecommand=(validate_entry40, "%P"))
    entry_position.grid(row=4, column=1, padx=5, pady=7)

    #стаж
    label_experience = tk.Label(tab1, text="Стаж работы:", anchor="w")
    label_experience.grid(row=6, column=0, padx=5, pady=7)

    entry_experience = ttk.Entry(tab1, width=35, validate="key", validatecommand=(validate_entry2, "%P"))
    entry_experience.grid(row=6, column=1, padx=5, pady=7)

    correct_label1 = tk.Label(tab1)
    correct_label1.grid(row=8, column=0, columnspan=2, pady=5)

    ##################################
    # Создание стиля
    custom_style = ttk.Style()

    # Установка цвета обводки
    custom_style.map('Select.TButton',
                     bordercolor=[('active', 'black')],
                     highlightcolor=[('focus', 'black')])

    selected_professions = tk.StringVar(tab1)
    selected_professions.set("1")

    label_info_prof = ttk.Label(tab1, text="Выберите количество доп профессий (от 1 до 4):")
    label_info_prof.grid(row=5, column=0, padx=5, pady=7)

    option_menu = ttk.OptionMenu(tab1, selected_professions, "1", "1","2", "3", "4", style='Select.TButton')
    option_menu.grid(row=5, column=2, padx=5, pady=7)

    submit_button = ttk.Button(tab1, text="Добавить профессии", command=lambda: create_profession_fields(entry_lists, root, selected_professions, validate_entry40, validate_entry2))
    submit_button.grid(row=5, column=1, padx=5, pady=5)

    entry_lists = []
    ########################################

    tab_control.add(tab1, text='Добавление сотрудника')
    tab_control.pack(expand=1, fill='both')


    button = ttk.Button(tab1, text="Добавить сотрудника", command=lambda:on_button_click(entry_FIO, entry_birthday, entry_department, selected_family_status, entry_position, entry_experience))
    button.grid(row=7, columnspan=2, padx=5, pady=5)

    button_change_department = ttk.Button(tab1, text="Изменение отдела у сотрудника", command=lambda: change_department(root))
    button_change_department.place(x=350, y=258)


def button_delete(deleteEntry):
    full_name = deleteEntry.get()
    check = check_person(full_name)
    if check == True:
        deleteEntry.delete(0, tk.END)
        deleteEmployee(full_name)
        messagebox.showinfo("Success", "Сотрудник удалён")
        updateTable(tree)
    else:
        messagebox.showerror("Error", "Сотрудник не найден")


def delete_all_department(deleteDepartment):
    department = deleteDepartment.get()
    list_full_name = select_full_name_where_department(department)
    if len(list_full_name) > 0 and len(department) > 0:
        messagebox.showinfo("Success", "Отдел успешно удалён")
        delete_full_name_where_department(list_full_name)
        updateTable(tree)
    else:
        messagebox.showerror("Error", "Такого отдела нет")


def tab_delete(tab2):

    validate_entry40 = tab2.register(validate_entry_40)

    deleteLabel_person = ttk.Label(tab2, text="Введите ФИО сотрудника которого вы хотите уволить")
    deleteLabel_person.grid(row=0, column=0)
    deleteEntry = ttk.Entry(tab2, validate="key", validatecommand=(validate_entry40, "%P"))
    deleteEntry.grid(row=0, column=1, padx=15, pady=5)


    deleteLabel_person = ttk.Button(tab2, text="Уволить сотрудника",command=lambda: button_delete(deleteEntry))
    deleteLabel_person.grid(row=1, column=0, pady=10)

    correct_label = ttk.Label(tab2)
    correct_label.grid(row=2, column=0, columnspan=2, pady=5)

    #удаление отдела
    deleteLabel = ttk.Label(tab2, text="Введите отдел который хотите уволить")
    deleteLabel.grid(row=3, column=0)
    deleteDepartment = ttk.Entry(tab2, validate="key", validatecommand=(validate_entry40, "%P"))
    deleteDepartment.grid(row=3, column=1, padx=15, pady=5)


    deleteButton = ttk.Button(tab2, text="Уволить отдел", command= lambda: delete_all_department(deleteDepartment))
    deleteButton.grid(row=4, column=0, pady=10)

    correct_label = ttk.Label(tab2)
    correct_label.grid(row=5, column=0, columnspan=2, pady=5)

    ######### таблица на странице удаления



    treeview1 = ttk.Treeview(tab2, show='tree')
    treeview1.column('#0', width=258, stretch=False)
    treeview1.heading('#0', text='ФИО')

    list_full_name = search_all_full_name()

    for i in list_full_name:
        treeview1.insert('', 'end', text=f'{i[0]}')

    treeview2 = ttk.Treeview(tab2, show='tree')
    treeview2.column('#0', width=258, stretch=False)
    treeview2.heading('#0', text='Отдел')

    list_department = search_all_department()

    for i in list_department:
        treeview2.insert('', 'end', text=f'{i[0]}')

    # Инициализация кнопок для переключения между таблицами
    switch_table1_button = ttk.Button(tab2, text="ФИО", command=lambda: switch_table1(treeview2, treeview1))
    switch_table1_button.place(x=35, y=196, width=130)

    switch_table2_button = ttk.Button(tab2, text="Отдел", command=lambda: switch_table2(treeview2, treeview1))
    switch_table2_button.place(x=163, y=196, width=130)

    # Отображение первой таблицы по умолчанию
    treeview1.place(x=35, y=220, width=258, height=100)


def switch_table1(treeview2, treeview1):
    treeview2.place_forget()
    treeview1.place(x=35, y=220, width=258, height=100)


def switch_table2(treeview2, treeview1):
    treeview1.place_forget()
    treeview2.place(x=35, y=220, width=258, height=100)


def print_worker_prof(root):
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
    popup.grab_set()


def print_family_status(root, entry_prof, selected_family_status_table):
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

    popup.grab_set()


def makeFile(data):
    doc = Document()

    # Создаем таблицу с двумя столбцами
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'  # Устанавливаем стиль таблицы

    # Добавляем заголовки столбцов
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'ФИО'
    hdr_cells[1].text = 'Опыт работы'

    # Добавляем данные в таблицу
    for row in data:
        new_row = table.add_row().cells
        new_row[0].text = str(row[0]) if len(row) > 0 else ''
        new_row[1].text = str(row[1]) if len(row) > 1 else ''

    # Сохраняем документ
    doc.save('WorkersTable.docx')
    messagebox.showinfo('Успех', "Файл WorkersTable.docx успешно сформирован")

def print_worker_prof(root):
    popup = tk.Toplevel(root)
    popup.title("Popup Window")
    popup.geometry("450x200")

    tree_worker_prof = ttk.Treeview(popup)
    tree_worker_prof.pack(fill='both', expand=True)

    data = search_worker_prof()

    tree_worker_prof["columns"] = ("ФИО", "Количество профессий")
    tree_worker_prof.column("#0", width=0, stretch=tk.NO)

    tree_worker_prof.heading("ФИО", text="ФИО")
    tree_worker_prof.heading("Количество профессий", text="Количество профессий")

    for row in data:
        tree_worker_prof.insert('', tk.END, values=row)

    makeFile(data)
    popup.grab_set()


def find_prof_person(root, entry_prof_2):
    popup = tk.Toplevel(root)
    popup.title("Popup Window")
    popup.geometry("200x150")

    tree_worker_prof = ttk.Treeview(popup)
    tree_worker_prof.pack(fill='both', expand=True)

    full_name = entry_prof_2.get()

    data = find_all_profession_person(full_name)

    tree_worker_prof["columns"] = ("Профессия")
    tree_worker_prof.column("#0", width=0, stretch=tk.NO)

    tree_worker_prof.heading("Профессия", text="Профессия")

    for row in data:
        tree_worker_prof.insert('', tk.END, values=row[1])

    popup.grab_set()


