import win32com.client
#pip install pywin32
#have to have office->outlook->version win32 

#CONSTANTS
#BE VERY CAREFUL WITH CONSTANTS CAN AND WILL SEND TO YOUR
SENDER_NAME = 'Shamoun Yousuf'
SUBJECT = "hello"
SEND_MESSAGE = "PLACEHOLDER"
LOOPS = 1 #for fun

#if more emails or replys specify exact email with this
#TIME_RECEIVED = ""
# format: 2024-12-20 22:52:15.643000+00:00 


# Initialize the Outlook application and get the inbox folder
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 represents the Inbox folder
messages = inbox.Items  

# Loop through the messages to find the one with the subject "hello"
for message in messages:
    if ((SENDER_NAME.lower() in message.SenderName.lower()) and (SUBJECT in message.Subject)): #and (TIME_RECEIVED in message.ReceivedTime)):
        print(message.ReceivedTime)
        for i in range(LOOPS):
            reply = message.Reply()  
            reply.Body = SEND_MESSAGE
            reply.Send()  
            print("Reply sent successfully.")
            print(f"Sender Name: {message.SenderName}")
            print(f"Full Subject Line: {message.Subject}")
        break  
else:
    print("No email found with parameters")