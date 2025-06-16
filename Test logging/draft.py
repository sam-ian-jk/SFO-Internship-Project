import tkinter as tk
from openpyxl import Workbook,load_workbook
import os
from datetime import datetime  
root = tk.Tk()
root.title("Test Logging")
root.geometry("1080x720")

def save_to_excel():
    pno = PN.get()
    sno = SN.get()
    ss=supp.get()
    tp=point.get()
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")
    file_path=f"{pno}.xlsx"
    if os.path.exists(file_path):
        wb = load_workbook(file_path)
        ws = wb.active
        if ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value is None:
            ws.append(["Date", "Time","Source Supply","Test Point", "impedence",] + [f"{sno}"])
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Date", "Time", "Source Supply","Test Point"] + [f"{sno}"])
    ws.append([date_str, time_str, ss,tp ] + measured_value)
    wb.save(file_path)
    tk.Label(ts, text="Saved to Excel ✅", fg="green").pack()

def clear():
    for widget in ts.winfo_children():
        widget.destroy()

    
tk.Label(root, text="Enter Part Number").pack()
PN= tk.Entry(root)
PN.pack()
tk.Label(root, text="Enter Serial number").pack()
SN = tk.Entry(root)
SN.pack()
tk.Label(root, text="Source Supply").pack()
supp = tk.Entry(root)
supp.pack()
tk.Label(root, text="Test Point").pack()
point=tk.Entry(root)
point.pack()
tk.Label(root, text="Measured Value").pack()
measured_value=tk.Entry(root)
measured_value.pack()
tk.Button(root, text="Save to Excel", command=save_to_excel).pack()
tk.Button(root, text="Next entry", command=clear).pack()
root.mainloop()
