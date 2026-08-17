import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import os
import pyodbc
import threading

from compare_ddr_aconex import main


# ============================================================
# SQL CONFIGURATION
# ============================================================

SQL_SERVER = "ES-MSSQL-01"
SQL_DATABASE = "ACONEX Reporting Data"


# ============================================================
# MAIN WINDOW
# ============================================================

root = tk.Tk()

# ============================================================
# DARK THEME
# ============================================================

BG = "#1E1E1E"
CARD = "#252526"
TEXT = "#FFFFFF"
ACCENT = "#0078D4"
SUCCESS = "#107C10"
WARNING = "#FFB900"
ENTRY_BG = "#2D2D30"

root.configure(bg=BG)
root.title("DDR vs ACONEX Validator")
root.geometry("950x650")
root.resizable(False, False)

ddr_file = tk.StringVar()
plip_file = tk.StringVar()
output_folder = tk.StringVar()

status_text = tk.StringVar(value="Ready")


# ============================================================
# LOGO
# ============================================================

try:
    img = Image.open("enppi_logo.png")
    img = img.resize((220, 80))
    logo = ImageTk.PhotoImage(img)

    logo_label = tk.Label(
        root,
        image=logo,
        bg=BG
    )
    logo_label.pack(pady=(10, 5))

except Exception:
    pass


# ============================================================
# TITLE
# ============================================================

title_label = tk.Label(
    root,
    text="DDR vs ACONEX VALIDATOR",
    font=("Segoe UI", 22, "bold"),
    bg=BG,
    fg=TEXT
)

title_label.pack(pady=(10, 20))


# ============================================================
# FUNCTIONS
# ============================================================

def create_button(parent, text, command, color):

    btn = tk.Button(
        parent,
        text=text,
        command=command,
        bg=color,
        fg="white",
        relief="flat",
        cursor="hand2",
        activebackground=color,
        activeforeground="white",
        font=("Segoe UI", 9, "bold")
    )

    return btn

def browse_ddr():
    path = filedialog.askopenfilename(
        title="Select DDR File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if path:
        ddr_file.set(path)


def browse_plip():
    path = filedialog.askopenfilename(
        title="Select PLIP File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if path:
        plip_file.set(path)


def browse_output():
    path = filedialog.askdirectory(
        title="Select Output Folder"
    )

    if path:
        output_folder.set(path)


def test_connection():

    try:

        conn_string = (
            "DRIVER={ODBC Driver 18 for SQL Server};"
            f"SERVER={SQL_SERVER};"
            f"DATABASE={SQL_DATABASE};"
            "Trusted_Connection=yes;"
            "TrustServerCertificate=yes;"
        )

        conn = pyodbc.connect(conn_string)

        cursor = conn.cursor()
        cursor.execute("SELECT 1")

        conn.close()

        messagebox.showinfo(
            "Success",
            "SQL Connection Successful."
        )

    except Exception as e:

        messagebox.showerror(
            "Connection Failed",
            str(e)
        )


def validation_worker():

    try:

        output_file = os.path.join(
            output_folder.get(),
            "DDR_Comparison_Result.xlsx"
        )

        root.after(
            0,
            lambda: status_text.set("Running Validation...")
        )

        from compare_ddr_aconex import main

        main(
            ddr_path=ddr_file.get(),
            plip_path=plip_file.get(),
            out_path=output_file
        )

        root.after(
            0,
            lambda: status_text.set("Completed Successfully")
        )
        
        root.after(0, progress.stop)
        
        root.after(
            0,
            lambda: messagebox.showinfo(
                "Success",
                f"Output generated successfully:\n\n{output_file}"
            )
        )

        if os.path.exists(output_file):
            os.startfile(output_file)

    except Exception as e:

        error_msg = str(e)

        root.after(
            0,
            lambda: status_text.set("Failed")
        )

        root.after(
            0,
            lambda msg=error_msg: messagebox.showerror(
                "Error",
                msg
            )
        )

def run_validation():

    if not ddr_file.get():
        messagebox.showerror(
            "Missing File",
            "Please select a DDR file."
        )
        return

    if not plip_file.get():
        messagebox.showerror(
            "Missing File",
            "Please select a PLIP file."
        )
        return

    if not output_folder.get():
        messagebox.showerror(
            "Missing Folder",
            "Please select an output folder."
        )
        return
    
    progress.start(10)

    threading.Thread(
        target=validation_worker,
        daemon=True
    ).start()

def add_hover(widget, normal, hover):

    widget.bind(
        "<Enter>",
        lambda e: widget.config(bg=hover)
    )

    widget.bind(
        "<Leave>",
        lambda e: widget.config(bg=normal)
    )
# ============================================================
# INPUT FRAME
# ============================================================

main_frame = tk.Frame(
    root,
    bg=CARD,
    padx=15,
    pady=15,
    highlightbackground="#3F3F46",
    highlightthickness=1
)

main_frame.pack(
    fill="x",
    padx=20,
    pady=10
)

# DDR

tk.Label(
    main_frame,
    text="DDR File",
    font=("Segoe UI", 10),
    bg=CARD,
    fg=TEXT
).grid(row=0, column=0, sticky="w", pady=10)

tk.Entry(
    main_frame,
    textvariable=ddr_file,
    width=70,
    bg=ENTRY_BG,
    fg="white",
    insertbackground="white",
    relief="flat"
).grid(row=0, column=1, padx=10)

create_button(
    main_frame,
    "Browse",
    browse_ddr,
    ACCENT
).grid(row=0, column=2)

# PLIP

tk.Label(
    main_frame,
    text="PLIP File",
    font=("Segoe UI", 10),
    bg=CARD,
    fg=TEXT
).grid(row=1, column=0, sticky="w", pady=10)

tk.Entry(
main_frame,
    textvariable=plip_file,
    width=70,
    bg=ENTRY_BG,
    fg="white",
    insertbackground="white",
    relief="flat"
).grid(row=1, column=1, padx=10)

create_button(
    main_frame,
    "Browse",
    browse_plip,
    ACCENT
).grid(row=1, column=2)

# OUTPUT

tk.Label(
    main_frame,
    text="Output Folder",
    font=("Segoe UI", 10),
    bg=CARD,
    fg=TEXT
).grid(row=2, column=0, sticky="w", pady=10)

tk.Entry(
    main_frame,
    textvariable=output_folder,
    width=70,
    bg=ENTRY_BG,
    fg="white",
    insertbackground="white",
    relief="flat"
).grid(row=2, column=1, padx=10)

create_button(
    main_frame,
    "Browse",
    browse_output,
    ACCENT
).grid(row=2, column=2)


# ============================================================
# SQL BUTTON
# ============================================================

sql_button = tk.Button(
    root,
    text="Test SQL Connection",
    command=test_connection,
    bg=ACCENT,
    fg="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 10, "bold"),
    padx=15,
    pady=8
)

sql_button.pack(pady=15)

style = ttk.Style()
style.theme_use("default")

style.configure(
    "blue.Horizontal.TProgressbar",
    troughcolor=BG,
    background=ACCENT
)

progress = ttk.Progressbar(
    root,
    style="blue.Horizontal.TProgressbar",
    mode="indeterminate",
    length=450
)

progress.pack(pady=10)
# ============================================================
# STATUS
# ============================================================

status_frame = tk.Frame(
    root,
    bg=BG
)

status_frame.pack(pady=10)

tk.Label(
    status_frame,
    text="Status:",
    bg=BG,
    fg=TEXT,
    font=("Segoe UI", 10, "bold")
).pack(side="left")

tk.Label(
    status_frame,
    textvariable=status_text,
    bg=BG,
    fg="#00D7FF",
    font=("Segoe UI", 10)
).pack(side="left", padx=5)

# ============================================================
# RUN BUTTON
# ============================================================

run_button = tk.Button(
    root,
    text="RUN VALIDATION",
    command=run_validation,
    bg=SUCCESS,
    fg="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 12, "bold"),
    width=25,
    height=2
)

run_button.pack(pady=25)

add_hover(run_button, SUCCESS, "#16A34A")
add_hover(sql_button, ACCENT, "#1A8CFF")

# ============================================================
# FOOTER
# ============================================================

footer = tk.Label(
    root,
    text="DDR / PLIP / ACONEX Validation Tool",
    bg=BG,
    fg="#808080",
    font=("Segoe UI", 9)
)

footer.pack(side="bottom", pady=10)
# ============================================================
# START APPLICATION
# ============================================================

root.mainloop()