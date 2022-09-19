# Imports
# This is for output
import logging

# This is not really needed, but its used with timetosleep variable for slower output in the terminal.
import time

# This is needed for reading the Excel file.
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

# This is for the queue output. This is added to the timeunit output that is called by worker.
def queue_output(que):
    que_msg = ""
    if len(que) == 1:
        que_msg = "no queue."
        return que_msg
    elif len(que) > 1:
        for i in range(1, len(que)):
            que_msg += f"PID {que[i]['Process ID']} wait={que[i]['Wait']} "
        return que_msg


# Keeps track of the timeunit and the logger.
def timeunit(message):
    logging.info("Timeunit %d:" % timeunit.counter + message)
    timeunit.counter += 1


timeunit.counter = 1

# Check if any process is ready to be added to the queue, it also adds the next process to be executed.
def check_arrivals():
    for _ in range(0, len(thequeue)):
        for i in thequeue:
            if i["Arrival Time"] == 1:
                thewaiters.append(i)
                thequeue.remove(i)


# This adds +1 to all processes that are waiting in the queue(thewaiters). It also removes 1 from arrival time for the ones that are yet to arrive. (those in the first list called thequeue)
def waiting(listofprocess):
    for i in range(0, len(listofprocess)):
        listofprocess[i]["Wait"] += 1
    for i in thequeue:
        i["Arrival Time"] -= 1


# This is what removes instruction load and prints at every time unit.
def worker(pid):
    pid[0]["Instruction Load"] -= 1
    timeunit(
        f"PID {pid[0]['Process ID']} executes. {pid[0]['Instruction Load']} instructions left. {queue_output(pid)}"
    )
    time.sleep(timetosleep)


# This is here to enable the logging. Could also specify a filename and get the logs in a file instead of the terminal
if __name__ == "__main__":
    format = f"%(message)s"
    logging.basicConfig(format=format, level=logging.INFO)

# Just a flag to disable the while loop.
working = True
while working:
    # Check if any processess are ready to be added to thewaiters and theexecutor.
    check_arrivals()
    waiting(thewaiters)
    if len(thewaiters) > 0:
        if thewaiters[0]["Instruction Load"] >= 1:
            worker(thewaiters)
        elif thewaiters[0]["Instruction Load"] == 0 and len(thewaiters) > 1:
            thewaiters.pop(0)
            timeunit("Context Switch")
        elif thewaiters[0]["Instruction Load"] == 0:
            thewaiters.pop(0)
    elif len(thewaiters) == 0:
        logging.info("All processes have executed. End of simulation.")
        working = False
