import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from PIL import Image, ImageTk
import os
from mails import leer_correos_outlook_y_guardar_en_excel, get_inbox_folders, connect_outlook

# Ejecutar la función
selected_folder = None  # Global variable to store the selected folder value

folders_dic = get_inbox_folders()

def on_submit():
    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()
    inbox_path = folders_dic[selected_folder.get()]
    folders_list = inbox_path.split('/')
    outlook = connect_outlook()
    i = 0
    for folder in folders_list:
        if i != 0:
            outlook = outlook.Folders(folder)
        i += 1
    leer_correos_outlook_y_guardar_en_excel(start_date, end_date, outlook)

root = tk.Tk()
root.title("Outlook Mail Processor")

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

# Combobox for folder selection
folder_label = ttk.Label(frame, text="Select Folder:")
folder_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
selected_folder = ttk.Combobox(frame, values=list(folders_dic.keys()), width=40)
selected_folder.grid(row=2, column=1, padx=5, pady=5)
selected_folder.set("Select a folder")

# Button to save data
get_dates_button = ttk.Button(frame, text="Search emails and process data", command=on_submit)
get_dates_button.grid(row=3, columnspan=2, pady=10)



basedir = os.path.dirname(__file__)

# resized_image = Image.open("images.png")
resized_image = Image.open(os.path.join(basedir, "images", "images.png"))
# Add an image below the buttons
# resized_image = Image.open("images/images.png")
resized_image = resized_image.resize((400, 150))
resized_photo = ImageTk.PhotoImage(resized_image)
image_label = ttk.Label(frame, image=resized_photo)
image_label.grid(row=4, columnspan=3)

root.mainloop()