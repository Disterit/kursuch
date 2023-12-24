import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from interface import main_interface
import hashlib

def close_window():
    root.destroy()

def check_password():
    entered_password = password_entry.get()
    entered_password = hashlib.md5(entered_password.encode()).hexdigest()
    expected_password = "c4ca4238a0b923820dcc509a6f75849b"

    #fbcb523b7a21baf0fe0264247af92821

    if entered_password == expected_password:
        messagebox.showinfo("Success", "Добро пожаловать!")
        close_window()
        main_interface()
    else:
        messagebox.showerror("Ошибка", "Неверный пароль")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Вход")
    root.geometry("200x120")  # Увеличил высоту для более корректного отображения виджетов

    frame = ttk.Frame(root, padding="10")
    frame.pack(fill=tk.BOTH, expand=True)

    password_label = ttk.Label(frame, text="Введите пароль:")
    password_label.pack()

    password_entry = ttk.Entry(frame, show="*")
    password_entry.pack()

    login_button = ttk.Button(frame, text="Войти", command=check_password)
    login_button.pack(pady=5)

    root.mainloop()
