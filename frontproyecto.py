import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry

from mails import leer_correos_outlook_y_guardar_en_excel, get_inbox_folders, connect_outlook



# Ejecutar la función

selected_button = None  # Global variable to store the selected button value

folders_dic = get_inbox_folders()


def on_submit():
    start_date = start_date_entry.get_date()
    print("start_date", start_date)
    end_date = end_date_entry.get_date()
    print("end_date", end_date)
    print("folders", folders_dic)
    inbox_path = folders_dic[selected_button]
    print("inbox", inbox_path)
    folders_list = inbox_path.split('/')
    print("Folders separated:", folders_list )
    outlook = connect_outlook()
    i= 0
    for folder in folders_list:
        print(i)
        print(folder)
        if i == 0:
            print("helloooo")
            # continue
        else:
            outlook= outlook.Folders(folder)
        i= i+1
    
    leer_correos_outlook_y_guardar_en_excel(start_date, end_date, outlook)

def on_button_click(label):
    global selected_button
    selected_button = label
   

def create_buttons(frame):
    for i, label in enumerate(folders_dic):
        button = ttk.Button(frame, text=label, command=lambda l=label: on_button_click(l))
        button.grid(row=3, column=i, padx=5, pady=5)

root = tk.Tk()
root.title("Create ")


# Create a frame
frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky="nsew")

# Start Date Label and Entry
start_date_label = ttk.Label(frame, text="Start Date:")
start_date_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
start_date_entry = DateEntry(frame, width=12, background='darkblue', foreground='white', borderwidth=2)
start_date_entry.grid(row=0, column=1, padx=5, pady=5)

# End Date Label and Entry
end_date_label = ttk.Label(frame, text="End Date:")
end_date_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
end_date_entry = DateEntry(frame, width=12, background='darkblue', foreground='white', borderwidth=2)
end_date_entry.grid(row=1, column=1, padx=5, pady=5)




# Button labels
create_buttons(frame)


# Button to save data
get_dates_button = ttk.Button(frame, text="Search emails and process data", command=on_submit)
get_dates_button.grid(row=2, columnspan=2, pady=10)


root.mainloop()