import tkinter as tk

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
root.mainloop()