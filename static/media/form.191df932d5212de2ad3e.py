import customtkinter as ctk
import tkinter.filedialog as fd
import tkinter.messagebox as mbox
import os
import graph as gh

app = ctk.CTk()
app.title('Organizational Chart')
app.geometry('800x600')

ctk.set_appearance_mode('system')
ctk.set_default_color_theme('blue')

label = ctk.CTkLabel(app, text='Organizational Chart')
label.pack(pady=20)

# File selection
file_path_var = ctk.StringVar()

def select_file():
    file_path = fd.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")],
        title="Select Excel File"
    )
    if file_path:
        file_path_var.set(file_path)

file_frame = ctk.CTkFrame(app)
file_frame.pack(pady=10)
file_label = ctk.CTkLabel(file_frame, text="Excel File:")
file_label.pack(side="left", padx=5)
file_entry = ctk.CTkEntry(file_frame, textvariable=file_path_var, width=400)
file_entry.pack(side="left", padx=5)
file_button = ctk.CTkButton(file_frame, text="Browse", command=select_file)
file_button.pack(side="left", padx=5)

# Manager and Month/Year fields (in case you want to keep them)
manager_name_label = ctk.CTkLabel(app, text='Manager Name')
manager_name_label.pack(pady=5)
manager_name_entry = ctk.CTkEntry(app)
manager_name_entry.pack(pady=5)

month_year_label = ctk.CTkLabel(app, text='Month and Year')
month_year_label.pack(pady=5)
month_year_entry = ctk.CTkEntry(app)
month_year_entry.pack(pady=5)

# Checkboxes for options
show_location_var = ctk.BooleanVar()
show_level_var = ctk.BooleanVar()
checkbox_frame = ctk.CTkFrame(app)
checkbox_frame.pack(pady=10)
show_location_cb = ctk.CTkCheckBox(checkbox_frame, text="Show Location", variable=show_location_var)
show_location_cb.pack(side="left", padx=10)
show_level_cb = ctk.CTkCheckBox(checkbox_frame, text="Show Level", variable=show_level_var)
show_level_cb.pack(side="left", padx=10)

def button_function():
    file_path = file_path_var.get()
    manager = manager_name_entry.get()
    month_year = month_year_entry.get()
    show_location = show_location_var.get()
    show_level = show_level_var.get()
    if not file_path or not os.path.exists(file_path):
        mbox.showerror("Error", "Please select a valid Excel file.")
        return
    if not manager or not month_year:
        mbox.showerror("Error", "Please enter both Manager Name and Month/Year.")
        return
    try:
        # Pass the selected file_path to load_data
        datasheet = gh.load_data(manager, month_year, file_path=file_path)
        df = gh.save_df(datasheet, manager, month_year)
        gh.generateGraph(df, manager, month_year, show_location=show_location, show_level=show_level)
        mbox.showinfo("Success", "Org chart generated successfully!")
    except Exception as e:
        mbox.showerror("Error", f"Failed to generate org chart:\n{e}")

button = ctk.CTkButton(app, text='Generate Graph', command=button_function)
button.pack(pady=20)

app.mainloop()


