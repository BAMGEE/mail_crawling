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
# var for skip messages numbering
skip = 1

# -------------- Excel Variations
# Beginning of Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
# Set the new workbook
wb = op.Workbook()
# Active the sheet
ws = wb.active

# -------------- Class
class Site:
    idx = 1
    list = []
    ws = None

# -------------- Site
kb = Site()
kb.idx = 1
kb.ws = wb.create_sheet('KB', 0)
kb.list = ['KB', 'kb']

bnk = Site()
bnk.idx = 1
bnk.ws = wb.create_sheet('BNK', 1)
bnk.list = ['BNK', 'bnk']

hana = Site()
hana.idx = 1
hana.ws = wb.create_sheet('하나', 2)
hana.list = ['하나금융', '하나투자']

yt = Site()
yt.idx = 1
yt.ws = wb.create_sheet('유안타', 3)
yt.list = ['유안타', 'yuanta']

eb = Site()
eb.idx = 1
eb.ws = wb.create_sheet('이베스트', 4)
eb.list = ['이베스트', 'ebst']

# -------------- File Path
SavePath = "C:\\Users\\gee\\Desktop\\"
AttachPath = "C:\\Users\\gee\\Desktop\\attachments\\"

# -------------- Mail Crawling Function
def MailCrawling(ws, i):
    # Viewing Progress
    ViewingProgress(i)
    
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
        attachment.SaveASFile(AttachPath + str(i) + "_" + str(j) + "_" + str(attachment)) # File Name
    return i

# -------------- Viewing Progress Function
def ViewingProgress(i):
    os.system('cls')
    print("In progress... (" + str((kb.idx-1)+(bnk.idx-1)+(hana.idx-1)+(yt.idx-1)+(eb.idx-1)+skip) + "/" + str(msg_count)+ ")")

# -------------- MAIN
if __name__ == '__main__':
    # -------------- Mail Crawling Filtering Kewords
    for mail in messages:
        if any(str in mail.Subject + mail.Body for str in kb.list):
            kb.idx = MailCrawling(kb.ws, kb.idx)
            kb.idx += 1
        elif any(str in mail.Subject + mail.Body for str in bnk.list):
            bnk.idx = MailCrawling(bnk.ws, bnk.idx)
            bnk.idx += 1
        elif any(str in mail.Subject + mail.Body for str in hana.list):
            hana.idx = MailCrawling(hana.ws, hana.idx)
            hana.idx += 1
        elif any(str in mail.Subject + mail.Body for str in yt.list):
            yt.idx = MailCrawling(yt.ws, yt.idx)
            yt.idx += 1
        elif any(str in mail.Subject + mail.Body for str in eb.list):
            eb.idx = MailCrawling(eb.ws, eb.idx)
            eb.idx += 1
        else:
            ViewingProgress(skip)
            skip += 1
            
    # -------------- Save Excel File
    print("KB: "+str(kb.idx-1))
    print("BNK: "+str(bnk.idx-1))
    print("HANA: "+str(hana.idx-1))
    print("YUANTA: "+str(yt.idx-1))
    print("EBEST: "+str(eb.idx-1))
    wb.save(SavePath + "mailcrawling.xlsx")