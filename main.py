import openpyxl as xl
import pywhatkit as pk
from datetime import datetime
sourceFilename = "Notification List.xlsx"

loadedExcel = xl.load_workbook(sourceFilename)
NameListSheet = loadedExcel.worksheets[0]
TemplateSheet = loadedExcel.worksheets[1]


mr = NameListSheet.max_row

for i in range(2, mr + 1):
    if NameListSheet.cell(row=i, column=7).value == 'Y' or 'y':

            PhoneNo = NameListSheet.cell(row=i, column=4).value
    
            templateMessage = TemplateSheet.cell(row=2, column = 2).value

            date_str = str(NameListSheet.cell(row=i, column=5).value)
            date_str = date_str.strip(' 00:00:00')
            date_str = datetime.strptime(date_str, '%Y-%m-%d').strftime('%d/%m/%y')

            time_str = str(NameListSheet.cell(row=i, column=6).value)
            time_str = datetime.strptime(time_str, '%H:%M:%S').strftime('%I:%M %p')

            print(time_str)
    
            dataToFillTemplate = {'[EventType]' : str(NameListSheet.cell(row=i, column=1).value),
                                  '[EventName]' : str(NameListSheet.cell(row=i, column=2).value),
                                  '[Date]'      : date_str,
                                  '[Time]'      : time_str }
   
            for key,value in dataToFillTemplate.items():        
               templateMessage = templateMessage.replace(key,value)

            introMessage = 'Dear SSPT devotee,\nWe have embarked on a new initiative to inform you via WhatsApp on key SSPT specific events and your ubayams & services at the temple.\nThis is one our efforts to improve our engagement with you. We seek you understanding while we finetune the initiative.\nLooking forward to your support and constructive feedback.\nPlease do not reply to this message.\nFor any queries, please call the Temple Office at 6298 5771'
            print(introMessage)
            pk.sendwhatmsg_instantly("+65" + str(PhoneNo), templateMessage, 15, True, 10)              
            now = datetime.now()
            TimeNow = now.strftime("%d/%m/%Y %H:%M")
            NameListSheet.cell(row=i ,column=8).value = TimeNow

   
loadedExcel.save('Notification List.xlsx')
        
 





  

