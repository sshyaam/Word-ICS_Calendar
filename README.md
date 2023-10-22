# WORD-ICS_Calnedar
This is a Python programmed Tool to Read Data in a .docx File Cell by Cell and append it to a .ics File.<br />

You May Directly Run The Program If Your Computer Has Python `3.0` and Above along with `Pip` and `python-docx`. However if you're unaware of how to run the program or don't have the above dependencies, you may install the .exe provided in the releases section and run the executable file directly!<br /><br />

The Program Takes the .docx File As Input And Provides a .ics File As Output.<br />
- - - -

# Behind The Scenes: <br />

The program takes the `path to .docx`, `start_date [YYYY-MM-DD]` and `end_date [YYYY-MM-DD]` as inputs and provides a `.ics` file as output.
The program consists of the following functions:
  splittimerange(time_range: str) -> This function is used to convert two time ranges in the given header such as `8:30AM - 9:30AM` into two return values `8:30am` and `9:30am`.
  splittime(time_str: str) -> This function has the capability to convert the `raw string date` into a `datetime object`.
  create_event(day, time_range, summary) -> This function creates the event along with its parameters so it can be appended into the `.ics` file.
  main() -> It goes through the table in the word file, strips and splits the cells and creates the relevent events to be appended into the `.ics` file.
- - - -

# Note:
Currently The Program Has Difficulty Reading Merged Cells, So If The Header Has `N` Rows, Please Make Sure That Each Coloumn Has `N` Rows As Well.
Also Please Make Sure That The Time Range In The Header Field Does Not Contain Space Between The `Raw Time Str[8:30]` and `am/pm` `Ex: 8:30am - 9:30am or 8:30am - 9:30pm`

- - - -
