import docx
import ics
from ics import Calendar, Event
import datetime
from datetime import datetime, timedelta
import re

calendar = Calendar()

def addevent(date, start_time, end_time, event_name, timezone):
    start_datetime = datetime.fromisoformat(datetime.strptime(f"{date} {start_time}", "%d/%m/%Y %I:%M%p").strftime(f"%Y-%m-%dT%H:%M:%S{timezone}"))
    end_datetime = datetime.fromisoformat(datetime.strptime(f"{date} {end_time}", "%d/%m/%Y %I:%M%p").strftime(f"%Y-%m-%dT%H:%M:%S{timezone}"))

    event = Event()
    event.name = event_name
    event.begin = start_datetime
    event.end = end_datetime

    calendar.events.add(event)

def split_time(time):
    time = time.replace(" ", "")
    if ("–" in time):
        return time.upper().split("–")
    else:
        return time.upper().split("-")
    
def getdate(start_date, end_date, target_day):
    start_datetime = datetime.strptime(start_date, "%d/%m/%Y")
    end_datetime = datetime.strptime(end_date, "%d/%m/%Y")
    delta = end_datetime - start_datetime
    result_dates = [start_datetime + timedelta(days=i) for i in range(delta.days + 1) if (start_datetime + timedelta(days=i)).strftime("%A") == target_day]
    result_dates = [date.strftime("%d/%m/%Y") for date in result_dates]
    return result_dates

def table_extract(file_path, repeat):
    if (repeat == False):
        doc = docx.Document(file_path)
        table = doc.tables[0]

        times = []
        days = {}
        for num, row in enumerate(table.rows):
            if (num == 0):
                for count, cell in enumerate(row.cells):
                    if (count == 2):
                        pass
                    else:
                        times.append(cell.text)
            else:
                global day
                subs = []
                for count, cell in enumerate(row.cells):
                    day = row.cells[0].text
                    if (count == 2):
                        pass
                    else:
                        subs.append(cell.text)
                days[day] = subs
        return times, days
    else:
        doc = docx.Document(file_path)
        table = doc.tables[0]

        times = []
        days = {}
        for num, row in enumerate(table.rows):
            if (num == 0):
                for count, cell in enumerate(row.cells):
                    if (count == 2):
                        pass
                    else:
                        times.append(cell.text)
            else:
                subs = []
                prev_cell_value = None
                repetition_count = 1
                for count, cell in enumerate(row.cells):
                    cell_value = cell.text
                    if count == 0:
                        day = cell_value
                    elif count == 2:
                        pass
                    else:
                        if cell_value == prev_cell_value:
                            repetition_count += 1
                        else:
                            if repetition_count > 1:
                                subs.append(f"{prev_cell_value}~x{repetition_count}~")
                            else:
                                subs.append(prev_cell_value)
                            repetition_count = 1
                        prev_cell_value = cell_value

                if repetition_count > 1:
                    subs.append(f"{prev_cell_value}~x{repetition_count}~")
                else:
                    subs.append(prev_cell_value)

                days[day] = subs

        return times, days

def calendar_create(timelist, daylist, sd, ed, timezone):
    answer = ""
    for key, value in daylist.items():
        for i, event in enumerate(list(value)):
            if (i == 0):
                pass
            elif (event == ''):
                pass
            else:
                time = split_time(timelist[i])
                dates = getdate(sd, ed, key)
                for date in dates:
                    match = re.search(r"~x(\d+)~", event)
                    if (match):
                        number = int(match.group(1))
                        name = re.sub(r"~x(\d+)~", "", event)
                        if (str(name).strip() == ''):
                            pass
                        else:
                            endtime = split_time(timelist[i + (number - 1)])
                            answer += (f"{date} | {time[0]} | {endtime[1]} | {name} | {timezone}\n")
                            addevent(date=date, start_time=time[0], end_time=endtime[1], event_name=name, timezone=timezone)
                    else:
                        answer +=(f"{date} | {time[0]} | {time[1]} | {event} | {timezone}\n")
                        addevent(date=date, start_time=time[0], end_time=time[1], event_name=event, timezone=timezone)
    return answer

def final(document, sd, ed, timezone, repeat=False):
    try:
        times, days = table_extract(document, repeat)
        res = calendar_create(times, days, sd, ed, timezone)
        with open('my_calendar.ics', 'w') as f:
            f.writelines(calendar.serialize())
        return res
    except Exception as e:
        return f"Error: {e}"