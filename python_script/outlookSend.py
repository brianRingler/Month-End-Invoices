import win32com.client as win32

def send_email_outlook(to, sub, body, path=""):
    '''
    Data will imported from Excel file into Pandas to populate outgoing emails. The Sender will be the Outlook account on the local machine running the Invoice Program
    * Using Outlook to send emails. 
    * to: The individual receiving the email
    * sub: Is the subject of the email 
    * body: The body of the email (HTML can be added)
    * path: Path to the attachment plus file name will end path
    '''

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = sub
    mail.Body = body
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    attachment = path
    mail.Attachments.Add(attachment)

    mail.Send()