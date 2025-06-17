import tkinter as tk
from tkinter import messagebox
from logic import (
    validate_inputs, is_pass, clear_all_data, next_entry_data, 
    start_entry_process, add_measurement_data, save_data_to_excel
)

def launch_app():
    root = tk.Tk()
    root.title("Industrial Test Logger")
    root.geometry("650x500")
    root.configure(bg="#f0f0f0")
    style = {'font': ('Arial', 11), 'bg': '#f0f0f0'}
    entry_style = {'font': ('Arial', 11)}

    input_frame = tk.Frame(root, bg="#f0f0f0")
    input_frame.pack(pady=10)

    def create_input_row(row, label_text):
        label = tk.Label(input_frame, text=label_text, **style)
        label.grid(row=row, column=0, sticky="e", padx=10, pady=5)
        entry = tk.Entry(input_frame, **entry_style, width=30)
        entry.grid(row=row, column=1, padx=10, pady=5)
        return entry

    global PN, SN, NUM, SS, TP, MV, TESTER
    PN = create_input_row(0, "Part Number")
    SN = create_input_row(1, "Serial Number")
    NUM = create_input_row(2, "Number of Testing Points")
    SS = create_input_row(3, "Source Supply")
    TP = create_input_row(4, "Test Point")
    MV = create_input_row(5, "Measured Value")
    TESTER = create_input_row(6, "Tested By")

    button_frame = tk.Frame(root, bg="#f0f0f0")
    button_frame.pack(pady=20)

    tk.Button(button_frame, text="Start Entry",
            command=lambda: start_entry_process(NUM),
            font=('Arial', 11), bg='#007acc', fg='white', width=18).grid(row=0, column=0, padx=10, pady=5)

    tk.Button(button_frame, text="Add Measurement",
            command=lambda: add_measurement_data(SS, TP, MV),
            font=('Arial', 11), bg='#28a745', fg='white', width=18).grid(row=0, column=1, padx=10, pady=5)

    tk.Button(button_frame, text="Finish Serial Entry",
            command=lambda: save_data_to_excel(PN, SN, TESTER),
            font=('Arial', 11), bg='#ffc107', fg='black', width=18).grid(row=1, column=0, padx=10, pady=5)

    tk.Button(button_frame, text="Clear Fields (New SN)",
            command=lambda: next_entry_data(SS, TP, MV),
            font=('Arial', 11), bg='#17a2b8', fg='white', width=18).grid(row=1, column=1, padx=10, pady=5)

    tk.Button(button_frame, text="Create New File",
            command=lambda: clear_all_data([PN, SN, SS, TP, MV, TESTER, NUM]),
            font=('Arial', 11), bg='#dc3545', fg='white', width=38).grid(row=2, column=0, columnspan=2, pady=10)

    root.mainloop()
