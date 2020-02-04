from apiclient import discovery
from apiclient import errors
from httplib2 import Http
from oauth2client import file, client, tools
import base64
import re
from datetime import date, timedelta
import datetime
import tkinter as tk
from tkinter import filedialog
import os
root = tk.Tk()
root.withdraw()


#https://developers.google.com/gmail/api/quickstart/python
# https://support.google.com/mail/answer/7190

# Creating a storage.JSON file with authentication details
def GmailAuthenticationService():
    SCOPES = 'https://www.googleapis.com/auth/gmail.modify' # we are using modify and not readonly, as we will be marking the messages Read
    store = file.Storage('storage.json') 
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)

    GMAIL = discovery.build('gmail', 'v1', http=creds.authorize(Http()))
    return GMAIL

def ListMessagesMatchingQuery(service, user_id, query=''):

    try :
        response = service.users().messages().list(userId=user_id,q=query).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])

        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id, q=query,
                                                pageToken=page_token).execute()
            messages.extend(response['messages'])

        return (messages)

    except(errors.HttpError):
        print('An error occurred:')

def GetAttachments(service, user_id, msg_id, store_dir):
    """Get and store attachment from Message with given id.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        msg_id: ID of Message containing attachment.
        store_dir: The directory used to store attachments.
    """
    try:
        id = (msg_id)
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        # print(message)
        # for printing subject and dateTime of that message
        for hr in message['payload']['headers']:
            if hr['name']=='Subject':
                sub = (hr['value'])
            if hr['name']=='Date':
                date = (hr['value'])
                
        for part in message['payload']['parts']:
            if '.csv' in part['filename']:
                print(part['filename'], sub, id, date)
                if 'data' in part['body']:
                    data = part['body']['data']
                else:
                    att_id = part['body']['attachmentId']
                    att = service.users().messages().attachments().get(userId=user_id, messageId=msg_id,id=att_id).execute()
                    
                    data = att['data']

                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                filename = re.sub(r'\d', '', part['filename'])
                filename = re.sub(r'-', '', filename)
                path = '/'.join([store_dir,filename])
                print(file_data)
                with open(path,'w',encoding='UTF-8') as f :
                    f.write(file_data.decode("UTF-8"))
                    f.close()
                print('\n\n------ Saved File ....\n\n')
                break

    except(errors.HttpError):
        print('An error occurred:')



# ------- mail code starts from here ----


user_id =  'me'
label_id_one = 'inbox'

service = GmailAuthenticationService()

print("\n\n ----------- Enter Filter Details ------------- \n\n")

Subject_key_search = 'Data Bucket'
# Subject_key_search = input('Subject_key_search- Ex Data Bucket :')

From_key_search = 'donotreply@appfolio.com'
# From_key_search = input('From_key_search- Ex donotreply@appfolio.com :')

DateInput = '2019/07/18'
# DateInput = input('Date Input- Ex 2019/07/18 :')

StartDateInput = datetime.datetime.strptime(DateInput, '%Y/%m/%d') 
EndDateInput = StartDateInput + timedelta(1)
# print(StartDateInput,EndDateInput)



q = "has:attachment ,filename:csv ,subject:("+Subject_key_search+") ,from:"+From_key_search+" ,in:"+label_id_one+" ,after: "+StartDateInput.strftime('%Y/%m/%d')+" ,before: "+EndDateInput.strftime('%Y/%m/%d')+" "
mssg_list = (ListMessagesMatchingQuery(service, user_id, q))
print(mssg_list)
print("\n\n ----------- loading attachments ------------- \n\n")

if(len(mssg_list) >= 1):
    store_dir = filedialog.askdirectory(title = 'Save file into', initialdir=os.getcwd)
    for mssg in mssg_list:
        GetAttachments(service,user_id, mssg['id'], store_dir)
else:
    print('No files to Save for given filter')

    
# Removing Credentials given from user once done
# os.remove("storage.json")