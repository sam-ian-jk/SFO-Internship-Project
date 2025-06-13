import tkinter as tk
from openpyxl import Workbook,load_workbook
import os
from datetime import datetime 

measurement_entries=[]

def add_fields():
    try:
        count = int(number.get())
        for i in range(count):
            tk.Label(frame1, text=f"Measurement {i+1}:").pack()
            entry = tk.Entry(frame1)
            entry.pack()
            measurement_entries.append(entry)
    except ValueError:
        print("Please enter a valid number")
        
def clear():
    for widget in frame1.winfo_children():
        widget.destroy()
    measurement_entries.clear()

def save_to_excel():
    uid = ID.get()
    mac = MAC.get()
    measurements = [entry.get() for entry in measurement_entries]
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")
    file_path = "Test logging/data.xlsx"
    if os.path.exists(file_path):
        wb = load_workbook(file_path)
        ws = wb.active
        if ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None:
            ws.append(["Date", "Time", "Unique ID", "MAC"] + [f"Measurement {i+1}" for i in range(len(measurements))])
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Date", "Time", "Unique ID", "MAC"] + [f"Measurement {i+1}" for i in range(len(measurements))])
    ws.append([date_str, time_str, uid, mac] + measurements)
    wb.save(file_path)
    tk.Label(frame1, text="Saved to Excel ✅", fg="green").pack()
    
root = tk.Tk()
root.title("Test Logging")
root.geometry("640x320")

tk.Label(root, text="Enter Unique ID").pack()
ID = tk.Entry(root)
ID.pack()
tk.Label(root, text="Enter MAC number").pack()
MAC = tk.Entry(root)
MAC.pack()
tk.Label(root, text="Enter number of measurements").pack()
number = tk.Entry(root)
number.pack()
tk.Button(root, text="Add Fields", command=add_fields).pack()
frame1=tk.Frame(root)
frame1.pack()
tk.Button(root, text="Save to Excel", command=save_to_excel).pack()
tk.Button(root, text="Next entry", command=clear).pack()
root.mainloop()