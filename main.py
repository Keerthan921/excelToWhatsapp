import openpyxl as xl
import pywhatkit as pk

sourceFilename = "Notification List.xlsx"

wb1 = xl.load_workbook(sourceFilename)
ws1 = wb1.worksheets[0]
mr = ws1.max_row

for i in range(2, mr + 1):
    PhoneNo = ws1.cell(row=i, column=3).value
    Message = str(ws1.cell(row=i, column=2).value) + ", Thank you for attending " + str(ws1.cell(row=i, column=1).value)
    pk.sendwhatmsg_instantly("+65" + str(PhoneNo), Message, 30, True, 10)


