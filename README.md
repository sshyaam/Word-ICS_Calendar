# WORD-ICS_Calnedar
This is a Python programmed Tool to Read Data in a .docx File Cell by Cell and append it to a .ics File.<br />

You May Directly Run The Program If Your Computer Has Python `3.0` and Above along with `Pip` and `python-docx`. However if you're unaware of how to run the program or don't have the above dependencies, you may install the .exe provided in the releases section and run the executable file directly!<br /><br />

The Program Takes the .docx File As Input And Provides a .ics File As Output.<br />
- - - -

# Working: <br />

| File                | Description                                                                                        |
|---------------------|----------------------------------------------------------------------------------------------------|
| `front.py`          | Consists of Tkinter-based Front End.                                                              |
| `tt_extract.py`     | Consists of the functions called by the tkinter-based `front.py`|
| Functions Inside tt_extract.py                                               |                                                                                                    |
|----------------------------------------------------------------------------------|------------------------------------------------------------------------------------|
| `addevent(date, start_time, end_time, event_name, timezone)`                  | Adds an event to the calendar with specified date, start and end times, event name, and timezone.    |
| `split_time(time)`                                                           | Splits a time range string into start and end times, handling different separators.                   |
| `getdate(start_date, end_date, target_day)`                                 | Returns a list of dates between start_date and end_date that match the target_day.                    |
| `table_extract(file_path)`                                                   | Extracts data from a Word document table, returning times and a dictionary of days with events.       |
| `calendar_create(timelist, daylist, sd, ed, timezone)`                     | Creates events based on the extracted table data for each specified day and time range.              |
| `final(document, sd, ed, timezone)`                                         | Finalizes the process, extracting table data, creating events, and writing to the .ics file.          |


Please Make Sure You Have `ics`, `datetime`, `docx`, `pytz`, `tkinter`, `tkcalendar` Installed via `pip` If You're Running The Source Code.


- - - -

# Note:
Currently The Program Has Difficulty Reading Merged Cells, So If The Header Has `N` Rows, Please Make Sure That Each Coloumn Has `N` Rows As Well.
Also Please Make Sure That The Time Range In The Header Field Does Not Contain Space Between The `Raw Time Str[8:30]` and `am/pm` `Ex: 8:30am - 9:30am or 8:30am - 9:30pm`

- - - -
