import tkinter as tk
measurement_entries=[]
def add_fields():
    try:
        count = int(number.get())
        for i in range(count):
            tk.Label(ts, text=f"Measurement {i+1}:").pack()
            entry = tk.Entry(ts)
            entry.pack()
            measurement_entries.append(entry)
    except ValueError:
        print("Please enter a valid number")
def clear():
    for widget in ts.winfo_children():
        widget.destroy()
    measurement_entries.clear()
    
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
ts=tk.Frame(root)
ts.pack()
tk.Button(root, text="Next entry", command=clear).pack()
root.mainloop()