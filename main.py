import os
import datetime as dt
import win32com.client as win32
from icecream import ic
from zipfile import ZipFile
from exchangelib import (Account, Credentials, FileAttachment, HTMLBody,
                         Mailbox, Message)


today = dt.datetime.now().date() 

save_path = 'C:\\Users\\itanalyst\\Desktop\\ResizeExcel\\'
   
credentials = Credentials(
    username='DOMAIN\\USERNAME', 
    password='PASSWORD'
)

a = Account('username@domain.com', credentials=credentials, autodiscover=True)

def getFile(ext):
    '''
    function to get filename by its extension
    '''
    for file in os.listdir(save_path):
        filetime = dt.datetime.fromtimestamp(
                os.path.getctime(save_path + file))
        extension = os.path.splitext(file)[1]
        if filetime.date() == today and extension == ext:
            return str(os.path.splitext(file)[0])

def convertXLSB():
    '''
    function to convert xlsx to xlsb
    '''
    excel= win32.Dispatch("Excel.Application") 
    doc = excel.Workbooks.Open(os.path.join(save_path, str(getFile('.xlsx'))))
    doc.SaveAs(os.path.join(save_path, str(getFile('.xlsx')) + '.xlsb'), 50)
    doc = excel.Workbooks.Close()

def convertToZip():
    # Check if file exists
	if (getFile('.xlsb')) != "":
            zipObj = ZipFile(str(getFile('.xlsb')) + '.zip', 'w')
        # Add files to the zip
            zipObj.write(str(getFile('.xlsb'))+ '.xlsb')
            zipObj.close()


def sendEmail():
    '''
    function to send email with the zip file as attachment
    '''
    m = Message(
        account=a,
        subject='AUTOMATED DISCOUNT OVERLAP REPORT '  + str(today),
        body = HTMLBody("Dear ALL, <br/><br/> Please find attached report. <br/><br/>The report is also accessible at the following address: <br/><a style='color:blue; text-decoration:underline'>link here</a> "),
                to_recipients=[
            Mailbox(email_address='test@gmail.com')
        ],
    cc_recipients=['test2@gmail.com'],  # Simple strings work, too
   
    # bcc_recipients=[
        #  Mailbox(email_address='erik@example.com'),
        #    'felicity@example.com',
        #],  # Or a mix of both
    )

    attachments=[]
    with open(str(getFile('.zip')) + '.zip', 'rb') as f:
        content = f.read()
        attachments.append((str(getFile('.zip')) + '.zip', content))

    for attachment_name, attachment_content in attachments:
        file = FileAttachment(name=attachment_name, content=attachment_content)
        m.attach(file)
        
    m.send_and_save()

def deleteFiles():
    '''
    function to delete the excel files 
    '''
    os.remove(str(getFile('.xlsx')) + '.xlsx')
    os.remove(str(getFile('.xlsb')) + '.xlsb')

def main():
    '''using icecream to debug'''
    ic(convertXLSB())
    ic(convertToZip())
    ic(sendEmail())
    ic(deleteFiles())
    ic('DONE!')

if __name__== "__main__":
    	  main()

