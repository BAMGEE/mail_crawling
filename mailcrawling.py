import os
import openpyxl as op
import win32com.client
import arrow

# -------------- Outlook Variations
# Beginning of Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# The 'exture3' inbox in the total inbox(index: 6) 
inbox = outlook.GetDefaultFolder(6).Folders["exture3"]
# all messages in the 'exture3'
messages = inbox.Items
# all messages' count
msg_count = messages.count
# var for messages numbering
i = 1

# -------------- Excel Variations
# Beginning of Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
# Set the new workbook
wb = op.Workbook()
# Set the new sheet
ws = wb.active

# -------------- Mail Crawling
for mail in messages:
    # Viewing Progress
    os.system('cls')
    print("In progress... (" + str(i) + "/" + str(msg_count)+ ")")
    
    # The arrow object to get the received time.
    arrowobj = arrow.get(mail.ReceivedTime)
    
    # Col A: Numbering
    ws[f'A{i}'] = i
    # Col B: Received datetime(set 'tzinfo' to 'None')
    ws[f'B{i}'] = arrowobj.datetime.replace(tzinfo=None)
    # Col C: Sender Name(Sender Email Address)
    if mail.Class == 43 and mail.SenderEmailType == 'EX' and mail.Sender.GetExchangeUser() != None:
        ws[f'C{i}'] = mail.SenderName + "(" + mail.Sender.GetExchangeUser().PrimarySmtpAddress + ")"
    else:
        ws[f'C{i}'] = mail.SenderName + "(" + mail.SenderEmailAddress + ")"
    # Col D: receiving email addresse
    ws[f'D{i}'] = mail.To
    # Col E: Mail Title
    ws[f'E{i}'] = mail.Subject
    # Col F: Mail Body
    ws[f'F{i}'] = mail.Body
    
    # Get the mail attachments
    attachments = mail.Attachments
    # Mail attachments' count
    r = attachments.count
    # Loop to save attachments (File Name: MailNumbering_AttachmentCount_AttachmentTitle)
    for j in range(1, r+1):
        attachment = attachments.Item(j)
        attachment.SaveASFile("C:\\Users\\gee\\Desktop\\attachments\\" + str(i) + "_" + str(j) + "_" + str(attachment)) # File Name
    i += 1

# -------------- Save Excel File
wb.save("C:\\Users\\gee\\Desktop\\output.xlsx")