import sys, traceback, os

def excepthook(exctype, value, tb):
    with open(os.path.join(os.path.dirname(__file__), "error_log.txt"), "w", encoding="utf-8") as f:
        f.write("=== Uncaught Exception ===\n")
        traceback.print_exception(exctype, value, tb, file=f)
        f.write("\n==========================\n")

sys.excepthook = excepthook
import tkinter as tk
from tkinter import ttk, filedialog
from main import Controller

ACCENT = "#FF6B6B"   # light red
BG = "#1E1E1E"       # #blue
FG = "white"         
def toggle_entries(entries, var):
    state = "normal"
    for entry in entries:
        entry.config(state=state)

def browse_csv():
    path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
    if path:
        csv_path_var.set(path)

def browse_dest():
    path = filedialog.askopenfilename(
        title="Select an existing Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if path:
        dest_path_var.set(path)

def browse_new_folder():
    """Open a directory dialog for selecting a folder for new Excel file creation."""
    path = filedialog.askdirectory(title="Select Folder for New Excel File")
    if path:
        new_folder_var.set(path)




root = tk.Tk()
root.title("CSV Processing GUI")
root.configure(bg=BG)

# --- Style ---
style = ttk.Style()
style.theme_use("clam")

style.configure("TFrame", background=BG)
style.configure("TLabel", foreground=FG, background=BG, font=("Segoe UI", 11))
style.configure("Header.TLabel", foreground=ACCENT, background=BG, font=("Segoe UI", 13, "bold"))

style.configure("Rounded.TEntry",
                fieldbackground="#2B2B2B",
                foreground=FG,
                bordercolor=ACCENT,
                padding=8,
                insertcolor="white",
                relief="flat")

style.configure("Accent.TButton",
                foreground="white", background=ACCENT,
                font=("Segoe UI", 11, "bold"),
                padding=8)
style.map("Accent.TButton",
          background=[("active", "#FF8787")],
          relief=[("pressed", "sunken")])

style.configure("Custom.TCheckbutton",
                foreground=FG,
                background=BG,
                font=("Segoe UI", 10),
                padding=5)
style.map("Custom.TCheckbutton",
          foreground=[("active", ACCENT)],
          background=[("active", BG)])

# Variables
csv_path_var = tk.StringVar()
dest_path_var = tk.StringVar()
new_folder_var = tk.StringVar()
new_name_var = tk.StringVar()
start_date_var = tk.StringVar()
start_hour_var = tk.StringVar()
threshold1_var = tk.StringVar()
threshold2_var = tk.StringVar()
use_dest_var = tk.BooleanVar(value=False)
use_create_var = tk.BooleanVar(value=False)
end_date_var = tk.StringVar()
end_hour_var = tk.StringVar()

# Layout Frame
frm = ttk.Frame(root, padding=20, style="TFrame")
frm.pack(fill="both", expand=True)

row = 0

# --- File Inputs ---
ttk.Label(frm, text="File Settings", style="Header.TLabel").grid(row=row, column=0, sticky="w", pady=(0, 10))
row += 1

# CSV path
ttk.Label(frm, text="CSV File Path(Test Data):").grid(row=row, column=0, sticky="w", pady=5)
csv_entry = ttk.Entry(frm, textvariable=csv_path_var, width=45, style="Rounded.TEntry")
csv_entry.grid(row=row, column=1, sticky="ew", pady=5)
ttk.Button(frm, text="Browse", style="Accent.TButton", command=browse_csv).grid(row=row, column=2, padx=5)
row += 1


# Destination label
ttk.Label(frm, text="Destination", style="Header.TLabel").grid(row=row, column=0, sticky="w", pady=(20, 5))
row += 1

# checkboxes
ttk.Checkbutton(frm, text="Use Destination Path", variable=use_dest_var,
                style="Custom.TCheckbutton",
                command=lambda: toggle_entries([dest_entry], use_dest_var)).grid(row=row, column=0, sticky="w")
ttk.Checkbutton(frm, text="Create new file in existing folder", variable=use_create_var,
                style="Custom.TCheckbutton",
                command=lambda: toggle_entries([new_folder_entry, new_name_entry], use_create_var)).grid(row=row, column=1, sticky="w")
row += 1

# Destination path
ttk.Label(frm, text="Destination Path:").grid(row=row, column=0, sticky="w", pady=5)
dest_entry = ttk.Entry(frm, textvariable=dest_path_var, width=45, style="Rounded.TEntry")
dest_entry.grid(row=row, column=1, sticky="ew", pady=5)
ttk.Button(frm, text="Browse", style="Accent.TButton", command=browse_dest).grid(row=row, column=2, padx=5)
row += 1

# New folder
ttk.Label(frm, text="Choose Folder:").grid(row=row, column=0, sticky="w", pady=5)
new_folder_entry = ttk.Entry(frm, textvariable=new_folder_var, width=45, style="Rounded.TEntry")
new_folder_entry.grid(row=row, column=1, sticky="ew", pady=5)
ttk.Button(frm, text="Browse", style="Accent.TButton", command=browse_new_folder).grid(row=row, column=2, padx=5)
row += 1

# New file name
ttk.Label(frm, text="New File Name:").grid(row=row, column=0, sticky="w", pady=5)
new_name_entry = ttk.Entry(frm, textvariable=new_name_var, width=45, style="Rounded.TEntry")
new_name_entry.grid(row=row, column=1, sticky="ew", pady=5)
row += 1

# --- Date & Time ---
ttk.Label(frm, text="Run Settings", style="Header.TLabel").grid(row=row, column=0, sticky="w", pady=(20, 10))
row += 1

# Start Date
ttk.Label(frm, text="Start Date (DD/MM/YY):").grid(row=row, column=0, sticky="w", pady=5)
ttk.Entry(frm, textvariable=start_date_var, width=40, style="Rounded.TEntry").grid(row=row, column=1, sticky="ew", pady=5)
row += 1

# Start Hour
ttk.Label(frm, text="Start Hour (HH:MM):").grid(row=row, column=0, sticky="w", pady=5)
ttk.Entry(frm, textvariable=start_hour_var, width=40, style="Rounded.TEntry").grid(row=row, column=1, sticky="ew", pady=5)
row += 1

#End Date (optional)
end_date_var = tk.StringVar()
ttk.Label(frm, text="End Date (YY/MM/DD) (optional):").grid(row=row, column=0, sticky="w", pady=5)
ttk.Entry(frm, textvariable=end_date_var, width=40, style="Rounded.TEntry").grid(row=row, column=1, sticky="ew", pady=5)
row += 1

# End Hour (optional)
end_hour_var = tk.StringVar()
ttk.Label(frm, text="End Hour (HH:MM) (optional):").grid(row=row, column=0, sticky="w", pady=5)
ttk.Entry(frm, textvariable=end_hour_var, width=40, style="Rounded.TEntry").grid(row=row, column=1, sticky="ew", pady=5)
row += 1


# --- Thresholds ---
ttk.Label(frm, text="Thresholds", style="Header.TLabel").grid(row=row, column=0, sticky="w", pady=(20, 10))
row += 1

ttk.Label(frm, text="Lower Threshold:").grid(row=row, column=0, sticky="w", pady=5)
ttk.Entry(frm, textvariable=threshold1_var, width=20, style="Rounded.TEntry").grid(row=row, column=1, sticky="w", pady=5)
row += 1

ttk.Label(frm, text="Higher Threshold:").grid(row=row, column=0, sticky="w", pady=5)
ttk.Entry(frm, textvariable=threshold2_var, width=20, style="Rounded.TEntry").grid(row=row, column=1, sticky="w", pady=5)
row += 1

# --- Submit ---
generatebutton=ttk.Button(frm, text="Generate", style="Accent.TButton")
generatebutton.grid(row=row, column=0, columnspan=3, pady=30)





def on_generate():
    print("in on_generate")
    data = {
            "csv_path": csv_path_var.get(),
            "use_dest": use_dest_var.get(),
            "dest_path": dest_path_var.get(),
            "use_create": use_create_var.get(),
            "new_folder": new_folder_var.get(),
            "new_name": new_name_var.get(),
            "start_date": start_date_var.get(),
            "start_hour": start_hour_var.get(),
            "threshold1": threshold1_var.get(),
            "threshold2": threshold2_var.get(),
            "end_date": end_date_var.get(),
            "end_hour": end_hour_var.get(),

        }

    control=Controller(
        data=data
    )
    control.run()

generatebutton.config(command=on_generate)




# Resize handling
for i in range(3):
    frm.columnconfigure(i, weight=1)

#Prevent cutting off bottom widgets
root.update_idletasks()  
min_w = root.winfo_width()
min_h = root.winfo_height()
root.minsize(min_w, min_h)

root.mainloop()
