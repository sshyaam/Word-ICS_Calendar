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

def find_all_consecutive_strings(arr):
    result = []
    start_idx = -1

    for i in range(len(arr)):
        if i < len(arr) - 1 and arr[i] == arr[i + 1]:
            if start_idx == -1:
                start_idx = i
        else:
            if start_idx != -1:
                result.append({'start': start_idx, 'end': i})
                start_idx = -1
            else:
                result.append({'start': i, 'end': i})

    return result

def table_extract(file_path):
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

def calendar_create(timelist, daylist, sd, ed, timezone, repeat):
    timelist = timelist[1:]
    if (repeat == True):
        answer = ""
        for key, value in daylist.items():
            value = list(value)[1:]
            consecutives = find_all_consecutive_strings(value)
            for dict in consecutives:
                event = value[dict["start"]]
                startindex = dict["start"]
                endindex = dict["end"]
                starttime = split_time(timelist[startindex])[0]
                stoptime = split_time(timelist[endindex])[1]
                dates = getdate(sd, ed, key)
                for date in dates:
                    answer +=(f"{date} | {starttime} | {stoptime} | {event} | {timezone}\n")
                    addevent(date=date, start_time=starttime, end_time=stoptime, event_name=event, timezone=timezone)
            
        return answer
    else:
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
                        answer += (f"{date} | {time[0]} | {time[1]} | {event} | {timezone}\n")
                        addevent(date=date, start_time=time[0], end_time=time[1], event_name=event, timezone=timezone)
        return answer

def final(document, sd, ed, timezone, repeat=False):
    try:
        times, days = table_extract(document)
        res = calendar_create(times, days, sd, ed, timezone, repeat)
        with open('my_calendar.ics', 'w') as f:
            f.writelines(calendar.serialize())
        return res
    except Exception as e:
        return f"Error: {e}"