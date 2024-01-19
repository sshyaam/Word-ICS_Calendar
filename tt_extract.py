import docx
import ics
from ics import Calendar, Event
import datetime
from datetime import datetime, timedelta

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

def calendar_create(timelist, daylist, sd, ed, timezone):
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
                    addevent(date=date, start_time=time[0], end_time=time[1], event_name=event, timezone=timezone)

def final(document, sd, ed, timezone):
    try:
        times, days = table_extract(document)
        calendar_create(times, days, sd, ed, timezone)
        with open('my_calendar.ics', 'w') as f:
            f.writelines(calendar.serialize())
        return "1"
    except Exception as e:
        return e