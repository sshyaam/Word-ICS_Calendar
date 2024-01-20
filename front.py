import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkcalendar import Calendar, DateEntry
from tkinter.ttk import Checkbutton
import webbrowser

import datetime
from datetime import datetime
import pytz
import tt_extract
from tt_extract import final

def generate_timezone_dict():
    timezone_dict = {}
    for tz in pytz.all_timezones:
        timezone = pytz.timezone(tz)
        offset_seconds = timezone.utcoffset(datetime.utcnow()).total_seconds()
        offset_hours = offset_seconds // 3600
        offset_minutes = (offset_seconds % 3600) // 60
        offset_str = f"{offset_hours:+03.0f}:{offset_minutes:02.0f}"
        timezone_dict[tz] = offset_str
    return timezone_dict

all_timezones_dict = generate_timezone_dict()

def jumbledate(date):
    date_object = datetime.strptime(date, "%m/%d/%y")
    return date_object.strftime("%d/%m/%Y")

def convert_to_ics(file_path, start_date, end_date, timezone, repetition_check):
    reply = final(file_path, start_date, end_date, all_timezones_dict[timezone], repetition_check)
    if "Error: " not in reply:
        logs_text.config(state=tk.NORMAL)
        logs_text.insert(tk.END, f"{reply}\n\n")
        logs_text.insert(tk.END, "Success: .ics file generated successfully!\n")
        logs_text.config(state=tk.DISABLED)
    else:
        logs_text.config(state=tk.NORMAL)
        logs_text.insert(tk.END, f"{reply}\n")
        logs_text.config(state=tk.DISABLED)

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def on_start_date_change(event):
    end_date_cal['mindate'] = start_date_cal.get_date()

def show_tooltip(widget, message):
    x, y, _, _ = widget.bbox("insert")
    x += widget.winfo_rootx() + 25
    y += widget.winfo_rooty() + 25
    tooltip_geometry = f"+{x}+{y}"
    
    tooltip = tk.Toplevel(widget)
    tooltip.wm_overrideredirect(True)
    tooltip.wm_geometry(tooltip_geometry)
    
    label = tk.Label(tooltip, text=message, background="#ffffe0", relief="solid", borderwidth=1)
    label.pack(ipadx=1)

    def close_tooltip(event=None):
        tooltip.destroy()
    
    widget.bind("<Leave>", lambda e: close_tooltip())
    tooltip.bind("<Button-1>", lambda e: close_tooltip())

def download_ics():
    file_path = file_entry.get()
    start_date = jumbledate(start_date_cal.get())
    end_date = jumbledate(end_date_cal.get())
    timezone = timezone_options.get()
    repetition_check = repetition_var.get()
    
    tooltip_message = (
        "8:30AM - 9:00AM : Subject 1\n"
        "9:00AM - 10:00AM : Subject 1\n"
        "-> 8:30AM - 10:00AM : Subject 1"
    )
    
    show_tooltip(repetition_check_label, tooltip_message)
    convert_to_ics(file_path, start_date, end_date, timezone, repetition_check)

root = tk.Tk()
root.title("WORD TO ICS CONVERTER")
root.geometry("380x600")
root.configure(bg="sky blue")

title_label = tk.Label(root, text="WORD TO ICS CONVERTER", font=("Helvetica", 16), bg="sky blue")
divider_line = ttk.Separator(root, orient=tk.HORIZONTAL)
file_label = tk.Label(root, text="Select Word File:", bg="sky blue")
file_entry = tk.Entry(root, width=30)
browse_button = tk.Button(root, text="Browse", command=browse_file)
timezone_label = tk.Label(root, text="Select Timezone:", bg="sky blue")
timezone_options = ttk.Combobox(root, values=list(all_timezones_dict.keys()), state="readonly")
timezone_options.set("Asia/Kolkata")
start_date_label = tk.Label(root, text="Start Date:", bg="sky blue")
start_date_cal = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
end_date_label = tk.Label(root, text="End Date:", bg="sky blue")
end_date_cal = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
end_date_cal.bind("<<DateEntrySelected>>", on_start_date_change)

repetition_var = tk.BooleanVar()
repetition_check_label = tk.Label(root, text="Repetition Check", bg="sky blue")
repetition_tooltip_label = tk.Label(root, text="?", font=("Helvetica", 8), bg="grey", fg="white")
repetition_tooltip_label.bind("<Enter>", lambda e: show_tooltip(repetition_tooltip_label, "8:30AM - 9:00AM : Subject 1\n9:00AM - 10:00AM : Subject 1\n-> 8:30AM - 10:00AM : Subject 1"))
repetition_check = Checkbutton(root, variable=repetition_var, onvalue=True, offvalue=False)
repetition_check.grid(row=7, column=1, pady=5, sticky='w', padx=(5, 10), columnspan=2)
repetition_check_label.grid(row=7, column=0, pady=5, padx=(10, 5), sticky='w')
repetition_tooltip_label.grid(row=7, column=1, pady=5, sticky='w', padx=(28, 10))

download_button = tk.Button(root, text="Download", command=download_ics)

title_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))
divider_line.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky="ew")
file_label.grid(row=2, column=0, pady=(10, 5))
file_entry.grid(row=2, column=1, pady=(10, 5))
browse_button.grid(row=2, column=2, pady=(10, 5))
timezone_label.grid(row=3, column=0, pady=5)
timezone_options.grid(row=3, column=1, pady=5)
start_date_label.grid(row=4, column=0, pady=5)
start_date_cal.grid(row=4, column=1, pady=5)
end_date_label.grid(row=5, column=0, pady=5)
end_date_cal.grid(row=5, column=1, pady=5)

logs_separator = ttk.Separator(root, orient=tk.HORIZONTAL)
logs_separator.grid(row=9, column=0, columnspan=3, pady=(10, 5), sticky="ew")

logs_label = tk.Label(root, text="Logs", font=("Helvetica", 14), bg="sky blue")
logs_text = tk.Text(root, height=10, state='disabled', wrap='word')
logs_text.grid(row=10, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="nsew")
logs_label.grid(row=9, column=0, columnspan=3, pady=(0, 10))

logs_scrollbar = tk.Scrollbar(root, command=logs_text.yview)
logs_scrollbar.grid(row=10, column=3, sticky='nsew')
logs_text['yscrollcommand'] = logs_scrollbar.set

root.grid_rowconfigure(10, weight=1)
root.grid_columnconfigure(0, weight=1)

download_button.grid(row=8, column=0, columnspan=3, pady=(20, 10))

footer_frame = tk.Frame(root, bg="skyblue")
footer_label = tk.Label(footer_frame, text="Made by shyaaaaaaam", fg="skyblue", cursor="hand2", bg="white")
footer_label.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/shyaaaaaaam/Word-ICS_Calendar"))
footer_label.pack(pady=5)

footer_frame.grid(row=11, column=0, columnspan=3, sticky="ew")

root.grid_rowconfigure(11, weight=0)

root.resizable(False,False)

root.mainloop()
