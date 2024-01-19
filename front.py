import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkcalendar import Calendar, DateEntry
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

def convert_to_ics(file_path, start_date, end_date, timezone):
    reply = final(file_path, start_date, end_date, all_timezones_dict[timezone])
    if (reply == "1"):
        messagebox.showinfo("Success", ".icsfile Generated Successfully!")
    else:
        messagebox.showerror("Error", reply)

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def on_start_date_change(event):
    end_date_cal['mindate'] = start_date_cal.get_date()

def download_ics():
    file_path = file_entry.get()
    start_date = start_date_cal.get()
    end_date = end_date_cal.get()
    timezone = timezone_options.get()
    if not file_path:
        messagebox.showerror("Error", "Please select a Word file.")
        return
    convert_to_ics(file_path, jumbledate(start_date), jumbledate(end_date), timezone)

root = tk.Tk()
root.title("WORD TO ICS CONVERTER")
root.geometry("400x400")
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
download_button.grid(row=6, column=0, columnspan=3, pady=(20, 10))

root.mainloop()