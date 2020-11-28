from win10toast import ToastNotifier
import xlrd
import datetime
path = 'schedule.xlsx'

inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)
print("Running...")
while True:
    work = []
    timing = []
    toast = ToastNotifier()
    for y in range(1,inputWorksheet.nrows):
        work.append(inputWorksheet.cell_value(y,0))
        timing.append(inputWorksheet.cell_value(y,1))
        

    now = datetime.datetime.now()

    CurrentDate = str(now.day)+"/"+str(now.month)+"/"+str(now.year)+" "+str(now.hour)+":"+str(now.minute)
    CurrentDate1 = datetime.datetime.strptime(CurrentDate, "%d/%m/%Y %H:%M")
    
    for i in range(len(timing)):
        if timing[i] == CurrentDate:
            toast.show_toast("Notification",work[i],duration=20,icon_path="countdown_icon_161455.ico")
        else:
            continue



