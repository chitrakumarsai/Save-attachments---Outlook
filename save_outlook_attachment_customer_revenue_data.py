#!/usr/bin/env python
# coding: utf-8

# In[6]:


import datetime
import os
import win32com.client
import glob
import win32api, sys, os
files = glob.glob("path\\*")
for f in files:
    if f.endswith(".png") or f.endswith(".jpg") or f.endswith(".htm") or f.endswith("a.xlsx"):
        os.remove(f)
path = os.path.expanduser('~') + 'path/'
today = datetime.date.today()

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) 
messages = inbox.Items

def saveattachments(subject):
    for message in messages:
        if message.Subject == subject and message.Senton.date() == today:
            # body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                if message.Subject == subject and message.Unread:
                    message.Unread = False
                break


# In[7]:


subject = "Customer Revenue Data"
saveattachments(subject)


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




