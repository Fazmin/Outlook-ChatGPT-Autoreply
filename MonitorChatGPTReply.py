import openai
import requests
import os
import win32com.client
from bs4 import BeautifulSoup
import pythoncom
import time

# pip install openai requests bs4 pywin32 time

class OutlookHandler:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.inbox = self.outlook.GetDefaultFolder(6)  # 6 is the inbox folder

    def new_email_received(self, subject, sender):
        print(f"New email received from {sender}: {subject}")
        check_for_last_email()

class NewMailHandler:
    def OnNewMail(self, *args, **kwargs):
        outlook_handler = OutlookHandler()
        messages = outlook_handler.inbox.Items
        message = messages.GetLast()
        outlook_handler.new_email_received(message.Subject, message.SenderName)

def main():
    outlook_handler = OutlookHandler()

    event_handler = win32com.client.WithEvents(
        win32com.client.Dispatch("Outlook.Application"), NewMailHandler
    )

    print("Monitoring Outlook inbox...")

    while True:
        pythoncom.PumpWaitingMessages()
        time.sleep(10)

def should_reply_to(sender, allowed_senders):
    sender_lower = sender.lower()
    allowed_parts = [email.lower().split('@')[0] for email in allowed_senders]
    return any(part in sender_lower for part in allowed_parts)

# Leverage ChatGPT API to generate a response - Im using text-davinci-003 you can use ChatGPT4 if you have access.
def generate_response(outprompt):
    # submit OpenAI prompt
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=outprompt,
        temperature=0.8,
        max_tokens=400,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    # print response
    print(f"Reply generated: {response.choices[0].text}")
    return response.choices[0].text


def check_for_last_email():
    # Set up OpenAI API key and endpoint
    openai.api_key = "xxxx"

    # Access Outlook Application
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 refers to the inbox folder

    # Get the latest email
    messages = inbox.Items
    message = messages.GetLast()

    # Extract email subject, body, and sender
    subject = message.Subject
    body = message.Body
    sender = message.SenderEmailAddress
    print (f"Sent by: {sender}")

    # Remove HTML tags from the email body if needed
    if message.HTMLBody:
        soup = BeautifulSoup(message.HTMLBody, "html.parser")
        body = soup.get_text()
        # Microsoft Exchange emails come with a comprehensive Exchange tag. Therefore, merely specifying an email address may not suffice.
        allowed_senders = ["email1xxxxx@gmail.com", "email2xxxxx@gmail.com"]

    if should_reply_to(sender, allowed_senders):
        #This is the ChatGPT prompt - You are free to customize the response as desired.
        outprompt = f"Compose a reply to this email with subject '{subject}' and body: '{body}' keep it short but be nice and always greet."
        print (f"Feeding message: {outprompt}")
        response_text = generate_response(outprompt)

        # Open the reply message in the Outlook window for review
        reply = message.Reply()
        reply.Body = response_text

        # Send Reply
        reply.Send()
        os.system('cls')
        print("Continue monitoring Outlook 365 inbox...")

    else:
        print("Sender not in the allowed_senders list. No reply generated.")


if __name__ == "__main__":
    main()