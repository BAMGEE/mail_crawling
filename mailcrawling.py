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
KbIdx = 1
BnkIdx = 1
HanaIdx = 1
YtIdx = 1
EbIdx = 1

# -------------- Excel Variations
# Beginning of Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
# Set the new workbook
wb = op.Workbook()
# Active the sheet
ws = wb.active
# Set the new sheet
kbws = wb.create_sheet('KB', 0)
bnkws = wb.create_sheet('BNK', 1)
hanaws = wb.create_sheet('하나', 2)
ytws = wb.create_sheet('유안타', 3)
ebws = wb.create_sheet('이베스트', 4)

# -------------- Keywords List
# declude list
declude = ['DB', 'db', 'OMS', 'oms', '다올', '유지보수', '점검']
# solting include list
kblist = ['KB', 'kb']
bnklist = ['BNK', 'bnk']
hanalist = ['하나금융', '하나투자']
ytlist = ['유안타', 'yuanta']
eblist = ['이베스트', 'ebst']

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
    print("In progress... (" + str(KbIdx+BnkIdx+HanaIdx+YtIdx+EbIdx) + "/" + str(msg_count)+ ")")

# -------------- MAIN
if __name__ == '__main__':
    # -------------- Mail Crawling Filtering Kewords
    for mail in messages:
        if any(str in mail.Subject + mail.Body for str in declude):
            print(mail.Subject + "해당없음")
        elif any(str in mail.Subject + mail.Body for str in kblist):
            KbIdx = MailCrawling(kbws, KbIdx)
            KbIdx += 1
        elif any(str in mail.Subject + mail.Body for str in bnklist):
            BnkIdx = MailCrawling(bnkws, BnkIdx)
            BnkIdx += 1
        elif any(str in mail.Subject + mail.Body for str in hanalist):
            HanaIdx = MailCrawling(hanaws, HanaIdx)
            HanaIdx += 1
        elif any(str in mail.Subject + mail.Body for str in ytlist):
            YtIdx = MailCrawling(ytws, YtIdx)
            YtIdx += 1
        elif any(str in mail.Subject + mail.Body for str in eblist):
            EbIdx = MailCrawling(ebws, EbIdx)
            EbIdx += 1

    # -------------- Save Excel File
    wb.save(SavePath + "mailcrawling.xlsx")