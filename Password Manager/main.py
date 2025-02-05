import tkinter as tk
from tkinter import messagebox as ms
from random import choice, randint, shuffle
import pyperclip as pc
import pandas as pd

# ---------------------------- PASSWORD GENERATOR ------------------------------- #

def generate_password():

    password_entry.delete(0,tk.END)

    letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    symbols = ['!', '#', '$', '%', '&', '(', ')', '*', '+']

    password_letters = [choice(letters) for _ in range(randint(8, 10))]
    password_symbols = [choice(symbols) for _ in range(randint(2, 4))]
    password_numbers = [choice(numbers) for _ in range(randint(2, 4))]

    password_list = password_letters + password_symbols + password_numbers
    shuffle(password_list)

    password = "".join(password_list)

    password_entry.insert(0,password)
    pc.copy(password)
    

# ---------------------------- SAVE PASSWORD ------------------------------- #

def save():
    
    website = website_entry.get()
    email = email_entry.get()
    password = password_entry.get()

    if len(website) == 0 or len(password) == 0:

        ms.showinfo(title="Oops",message="Please make sure you haven't left any fields empty")
        
    else:
        is_ok = ms.askokcancel(title=website, message=f"These are the details entered: \nEmail: {email} \nPassword: {password} \nIs it okay to save?")

        if is_ok:

            new_data = {
                "Website":[website],
                "Email/Username": [email],
                "Password": [password]
            }

            new_df = pd.DataFrame(new_data)

            try:
                # Try to read existing data
                data_file = pd.read_excel(r"C:\Users\nil.monreal\Desktop\Nil\Passwords.xlsx")
                data_file = pd.concat([data_file, new_df], ignore_index=True)
            except FileNotFoundError:
                data_file = new_df
                # If file does not exist, create a new one

            data_file.to_excel(r"C:\Users\nil.monreal\Desktop\Nil\Passwords.xlsx", index=False)

            website_entry.delete(0,tk.END)
            password_entry.delete(0,tk.END)

        else:
            pass

# ---------------------------- UI SETUP ------------------------------- #

window = tk.Tk()
window.title("Password Manager")
window.config(padx=50, pady=50, bg="white")


canvas = tk.Canvas(width=200, height=200, bg="white", highlightthickness=0)
my_pas_image = tk.PhotoImage(file=r"C:\Users\nil.monreal\Desktop\Nil\Personal\Studies\Python\100 Days of coding\Intermediate\Day 29 - Password Manager\logo.png")
canvas.grid(column=1, row=0)

#Labels
website_label = tk.Label(text="Website:",bg="white")
website_label.grid(column=0,row=1)

email_label = tk.Label(text="Email/Username:",bg="white")
email_label.grid(column=0,row=2)

password_label = tk.Label(text="Password:",bg="white")
password_label.grid(column=0,row=3)

#Entries
website_entry = tk.Entry(width=35,justify="left")
website_entry.focus()
website_entry.grid(column=1,row=1,columnspan=2)

email_entry = tk.Entry(width=35,justify="left")
email_entry.insert(0,"nilmonreal@gmail.com")
email_entry.grid(column=1,row=2,columnspan=2)

password_entry = tk.Entry(width=20)
password_entry.grid(column=1,row=3)

#Buttons

gen_password = tk.Button(text="Gen. Password",command=generate_password)
gen_password.grid(column=2,row=3)

add_password = tk.Button(text="Add",command=save,width=35)
add_password.grid(column=1,row=4,columnspan=2)

window.mainloop()
