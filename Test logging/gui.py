import tkinter as tk
from tkinter import messagebox
from logic import (
    validate_inputs, is_pass, clear_all_data, next_entry_data, 
    start_entry_process, add_measurement_data, save_data_to_excel
)

def launch_app():
    root = tk.Tk()
    root.title("Industrial Test Logger")
    root.geometry("650x520")
    root.configure(bg="#f7f6f9") 

    style = {'font': ('Segoe UI', 11), 'bg': '#f7f6f9', 'fg': '#333'}
    entry_style = {'font': ('Segoe UI', 11), 'bg': '#ffffff', 'fg': '#333'}

    input_frame = tk.Frame(root, bg="#f7f6f9")
    input_frame.pack(pady=10)

    def create_input_row(row, label_text):
        label = tk.Label(input_frame, text=label_text, **style)
        label.grid(row=row, column=0, sticky="e", padx=12, pady=6)
        entry = tk.Entry(input_frame, **entry_style, width=32, relief="flat", highlightthickness=1, highlightcolor="#9370DB", highlightbackground="#ccc")
        entry.grid(row=row, column=1, padx=12, pady=6)
        return entry

    global PN, SN, NUM, SS, TP, MV, TESTER
    PN = create_input_row(0, "Part Number")
    SN = create_input_row(1, "Serial Number")
    NUM = create_input_row(3, "Number of Testing Points")
    SS = create_input_row(4, "Source Supply")
    TP = create_input_row(5, "Test Point")
    MV = create_input_row(6, "Measured Value")
    TESTER = create_input_row(7, "Tested By")

    button_frame = tk.Frame(root, bg="#f7f6f9")
    button_frame.pack(pady=25)

    def fancy_button(master, text, command, bg, fg):
        return tk.Button(
            master, text=text, command=command,
            font=('Segoe UI Semibold', 10), bg=bg, fg=fg,
            width=22, height=2, relief="flat", activebackground=bg, activeforeground=fg, cursor="hand2"
        )

    fancy_button(button_frame, "Start Entry", lambda: start_entry_process(NUM), "#6A5ACD", "white")\
        .grid(row=0, column=0, padx=10, pady=8)

    fancy_button(button_frame, "Add Measurement", lambda: add_measurement_data(SS, TP, MV), "#20B2AA", "white")\
        .grid(row=0, column=1, padx=10, pady=8)

    fancy_button(button_frame, "Finish Serial Entry", lambda: save_data_to_excel(PN, SN, TESTER), "#FFD700", "#333")\
        .grid(row=1, column=0, padx=10, pady=8)

    fancy_button(button_frame, "Clear Fields (New SN)", lambda: next_entry_data(SS, TP, MV), "#87CEFA", "#333")\
        .grid(row=1, column=1, padx=10, pady=8)

    fancy_button(button_frame, "Create New File", lambda: clear_all_data([PN, SN, SS, TP, MV, TESTER, NUM]), "#FF6B6B", "white")\
        .grid(row=2, column=0, columnspan=2, pady=10)

    root.mainloop()
