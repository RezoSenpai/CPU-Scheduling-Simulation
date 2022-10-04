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

# Create a list of all the values inside the Excel sheet. Put them into a dictionary so everything is easily accessible
thequeue = []
for i, x, z, c in sheet.iter_rows(min_row=2, values_only=True):
    d = {
        "Process ID": i,
        "Arrival Time": x,
        "Instruction Load": z,
        "Priority": c,
        "Wait": 0,
        "Executed": 0,
    }
    thequeue.append(d)

# List for processes that are waiting
thewaiters = []
# List for the process that is executing.
theexecutor = []


# Sort the queue
# I am sorting it after executions and arrival time. The shortest job is selected first.
# If a process has already been executed it is skipped until every processs has been executed once.
# If 2 processes have the same instruction load, the one that has waited the longest will execute first.
def sortque(que):
    return sorted(que, key=lambda d: (d["Executed"], d["Instruction Load"], -d["Wait"]))


# This is for the queue output. This is added to the timeunit output that is called by worker.
def queue_output(que):
    que_msg = ""
    if len(que) == 0:
        que_msg = "No queue."
        return que_msg
    elif len(que) >= 1:
        for i in range(0, len(que)):
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
    sortedlist = sortque(thewaiters)
    for i in sortedlist:
        if i["Arrival Time"] == 1 and len(theexecutor) == 0:
            theexecutor.append(i)
            thewaiters.remove(i)


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
        f"PID {pid[0]['Process ID']} executes. {pid[0]['Instruction Load']} instructions left. Q={quantum}.\n{queue_output(thewaiters)}"
    )
    time.sleep(timetosleep)


# This is here to enable the logging. Could also specify a filename and get the logs in a file instead of the terminal
if __name__ == "__main__":
    format = f"%(message)s"
    logging.basicConfig(format=format, level=logging.INFO)


# Time units a process should be executed before moving to a new one.
quantum = 4

# Just a flag to disable the while loop.
working = True
while working:
    # Check if any processess are ready to be added to thewaiters and theexecutor.
    check_arrivals()
    # If there is anything in the executor
    if len(theexecutor) > 0 and quantum > 0:
        # Checks wether or not the process actually has any instructions left.
        # If it does it works on it.
        if theexecutor[0]["Instruction Load"] >= 1:
            quantum -= 1
            waiting(thewaiters)
            worker(theexecutor)
        # If it doesn't have instructions it checks if there is anything waiting in the queue, and if not it is done.
        # This exists because in some cases the process could be finished before quantum hits 0. This basically removes 1 extra context switch.
        elif theexecutor[0]["Instruction Load"] == 0 and len(thewaiters) == 0:
            logging.info("All processes have executed. End of simulation.")
            working = False
        # If there are no instructions left, it removes the process completely and +1 is added to the timeunit.
        elif theexecutor[0]["Instruction Load"] == 0:
            theexecutor.pop()
            waiting(thewaiters)
            timeunit(f"Context Switch\n{queue_output(thewaiters)}")
            quantum = 4
    # This is another check for if it has finished all processes.
    elif len(theexecutor) == 0 and len(thewaiters) == 0:
        logging.info("All processes have executed. End of simulation.")
        working = False
    # If the process is done and there are still more in queue. Remove it and context switch.
    elif theexecutor[0]["Instruction Load"] == 0 and len(thewaiters) > 0:
        theexecutor.pop()
        waiting(thewaiters)
        timeunit(f"Context Switch\n{queue_output(thewaiters)}")
        quantum = 4
    # Remove process from execution and put it back into queue because it is not finished yet.
    elif len(theexecutor) > 0 and quantum == 0:
        theexecutor[0]["Executed"] += 1
        thewaiters.append(theexecutor.pop())
        waiting(thewaiters)
        timeunit(f"Context Switch\n{queue_output(thewaiters)}")
        quantum = 4
