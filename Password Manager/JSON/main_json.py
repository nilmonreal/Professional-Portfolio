import tkinter as tk
from tkinter import messagebox as mb
from random import choice, randint, shuffle
import pyperclip as pc
import json as js

# ---------------------------- PASSWORD GENERATOR ------------------------------- #

#Password Generator Project
def generate_password():
    letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    symbols = ['!', '#', '$', '%', '&', '(', ')', '*', '+']

    password_letters = [choice(letters) for _ in range(randint(8, 10))]
    password_symbols = [choice(symbols) for _ in range(randint(2, 4))]
    password_numbers = [choice(numbers) for _ in range(randint(2, 4))]

    password_list = password_letters + password_symbols + password_numbers
    shuffle(password_list)

    password = "".join(password_list)
    password_entry.insert(0, password)
    pc.copy(password)

# ---------------------------- SAVE PASSWORD ------------------------------- #
def save():

    website = website_entry.get()
    email = email_entry.get()
    password = password_entry.get()
    new_data = {
        website:{
            "email": email,
            "password": password,
        }
    }

    if len(website) == 0 or len(password) == 0:
        mb.showinfo(title="Oops", message="Please make sure you haven't left any fields empty.")
        
    else:
        try:
            with open(r"C:\Users\nil.monreal\Desktop\Nil\Personal\Studies\Python\100 Days of coding\Intermediate\Day 29 - Password Manager\JSON\data.json", "r") as data_file:
                #Reading old data
                data = js.load(data_file)
        
        except FileNotFoundError:
            with open(r"C:\Users\nil.monreal\Desktop\Nil\Personal\Studies\Python\100 Days of coding\Intermediate\Day 29 - Password Manager\JSON\data.json", "w") as data_file:
                #Saving updated data
                js.dump(new_data, data_file, indent=4)

        else:
            #Updating old data with new data
            data.update(new_data)

            with open(r"C:\Users\nil.monreal\Desktop\Nil\Personal\Studies\Python\100 Days of coding\Intermediate\Day 29 - Password Manager\JSON\data.json", "w") as data_file:
                #Saving updated data
                js.dump(new_data, data_file, indent=4)
                
        finally:
            website_entry.delete(0, tk.END)
            password_entry.delete(0, tk.END)
# ---------------------------- FIND PASSWORD ------------------------------- #

def find_password():
    website = website_entry.get()
    with open(r"C:\Users\nil.monreal\Desktop\Nil\Personal\Studies\Python\100 Days of coding\Intermediate\Day 29 - Password Manager\JSON\data.json", "w") as data_file:
        #Saving updated data
        data = js.load(data_file)
        if website in data:
            email = data[website]["email"]
            password = data[website["password"]]
            mb.showinfo(title=website, message=f"Email: {email}\nPassword: {password}")

           

# ---------------------------- UI SETUP ------------------------------- #

window = tk.Tk()
window.title("Password Manager")
window.config(padx=50, pady=50)

canvas = tk.Canvas(height=200, width=200)
logo_img = tk.PhotoImage(file=r"100 Days of coding\Intermediate\Day 29 - Password Manager\JSON\logo.png")
canvas.create_image(100, 100, image=logo_img)
canvas.grid(row=0, column=1)

#Labels
website_label = tk.Label(text="Website:")
website_label.grid(column=0,row=1)

email_label = tk.Label(text="Email/Username:")
email_label.grid(column=0,row=2)

password_label = tk.Label(text="Password:")
password_label.grid(column=0,row=3)

#Entries
website_entry = tk.Entry(width=30)
website_entry.focus()
website_entry.grid(column=1,row=1)

email_entry = tk.Entry(width=30)
email_entry.insert(0,"nilmonreal@gmail.com")
email_entry.grid(column=1,row=2,columnspan=2)

password_entry = tk.Entry(width=21)
password_entry.grid(column=1,row=3)

#Buttons

search_button = tk.Button(text="Search",width=13, command=find_password)
search_button.grid(column=2,row=1)

gen_password = tk.Button(text="Gen. Password",command=generate_password)
gen_password.grid(column=2,row=3)

add_password = tk.Button(text="Add",command=save,width=35)
add_password.grid(column=1,row=4,columnspan=2)

window.mainloop()