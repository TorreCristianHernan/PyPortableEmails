import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from mails import leer_correos_outlook_y_guardar_en_excel


# Ejecutar la función

def get_dates():
    start_date = start_date_entry.get_date()
    end_date = end_date_entry.get_date()
    print("Start Date:", start_date)
    print("End Date:", end_date)
    leer_correos_outlook_y_guardar_en_excel(start_date, end_date)

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

# Button to get dates
get_dates_button = ttk.Button(frame, text="Search emails and process data", command=get_dates)
get_dates_button.grid(row=2, columnspan=2, pady=10)




root.mainloop()