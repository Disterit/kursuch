import tkinter as tk
from tkinter import messagebox
from interface import main
import hashlib

def close_window():
    root.destroy()


def check_password():
    entered_password = password_entry.get()

    entered_password = hashlib.md5(entered_password.encode()).hexdigest()

    expected_password = "fbcb523b7a21baf0fe0264247af92821"

    if entered_password == expected_password:
        messagebox.showinfo("Success", "Добро пожаловать!")
        close_window()
        main()
    else:
        messagebox.showerror("Ошибка", "Неверный пароль")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Вход")
    root.geometry("200x75")

    password_label = tk.Label(root, text="Введите пароль:")
    password_label.pack()

    password_entry = tk.Entry(root, show="*")
    password_entry.pack()

    login_button = tk.Button(root, text="Войти", command=check_password)
    login_button.pack()

    root.mainloop()