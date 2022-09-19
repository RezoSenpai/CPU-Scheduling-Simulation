# imports
import logging
import time

import openpyxl

# Increase this to make the log move slower. Nice for following along in the terminal.
timetosleep = 0

# Open Excel
workbook = openpyxl.load_workbook(filename="cpu-scheduling.xlsx")
sheet = workbook.active
# --------------------------------
# Create a list of all the values inside the Excel sheet. Put them into a dictionary so everything is easily accessible
thequeue = []
for i, x, z, c in sheet.iter_rows(min_row=2, values_only=True):
    d = {
        "Process ID": i,
        "Arrival Time": x,
        "Instruction Load": z,
        "Priority": c,
        "Wait": 0,
    }
    thequeue.append(d)
# -----------------------------
# List for processes that are waiting
thewaiters = []
theexecutor = []


def sortque(que):
    return sorted(que, key=lambda d: (d["Priority"], d["Process ID"]))


def queue_output(que):
    que_msg = ""
    if len(que) == 0:
        que_msg = "no queue."
        return que_msg
    elif len(que) >= 1:
        for i in range(0, len(que)):
            que_msg += f"[PID {que[i]['Process ID']} wait={que[i]['Wait']}] "
        return que_msg


def timeunit(message):
    logging.info("Timeunit %d:" % timeunit.counter + message)
    timeunit.counter += 1


timeunit.counter = 1


def check_arrivals():
    for _ in range(0, len(thequeue)):
        for i in thequeue:
            if i["Arrival Time"] == 1:
                thewaiters.append(i)
                thequeue.remove(i)
    sortedlist = sortque(thewaiters)
    for i in sortedlist:
        if i["Arrival Time"] == 1 and len(theexecutor) == 0:
            theexecutor.append(i)
            thewaiters.remove(i)


def waiting(listofprocess):
    for i in range(0, len(listofprocess)):
        listofprocess[i]["Wait"] += 1
    for i in thequeue:
        i["Arrival Time"] -= 1


def worker(pid):
    pid[0]["Instruction Load"] -= 1
    timeunit(
        f"PID {pid[0]['Process ID']} executes. {pid[0]['Instruction Load']} instructions left. {queue_output(thewaiters)}"
    )
    time.sleep(timetosleep)


if __name__ == "__main__":
    format = f"%(message)s"
    logging.basicConfig(format=format, level=logging.INFO)


working = True
while working:
    check_arrivals()
    waiting(thewaiters)
    if len(theexecutor) > 0:
        if theexecutor[0]["Instruction Load"] >= 1:
            worker(theexecutor)
        elif theexecutor[0]["Instruction Load"] == 0 and len(thewaiters) > 0:
            theexecutor.pop()
            timeunit("Context Switch")
        elif theexecutor[0]["Instruction Load"] == 0 and len(thewaiters) == 0:
            theexecutor.pop(0)
    elif len(theexecutor) == 0:
        logging.info("All processes have executed. End of simulation.")
        working = False
