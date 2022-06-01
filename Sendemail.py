
import os
import win32com.client as win32

# construct Outlook application instance
olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

# construct the email item object
mailItem = olApp.CreateItem(0)
mailItem.Subject = 'Test '  #can be any subject
mailItem.BodyFormat = 1
mailItem.Body = "Attachment of Consolidate View"  #can be any body
mailItem.To = 'email' 

# mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('<email@gmail.com'))) [NOTHING JUST IGNORE FOR NOW! DONT DELETE ]

mailItem.Attachments.Add(os.path.join(os.getcwd(), r'directory'))
# mailItem.Attachments.Add(os.path.join(os.getcwd(), r'C:\Users\Kamarul_Syamil\Desktop\Dell\Project\Test2.csv')) <*sample*>

mailItem.Display()
mailItem.Send()
