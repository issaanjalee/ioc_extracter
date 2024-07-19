from exchangelib import Credentials, Account
# Input your credentials here
email = 'testingforpy@outlook.com'
password = 'testing@2023'
# Connect to the Outlook account
credentials = Credentials(email, password)
account = Account(email, credentials=credentials, autodiscover=True)
# Get the desired email by subject
email_subject = ''
email_items = account.inbox.filter(subject__contains=email_subject)
if email_items:
    # Get the latest email with the specified subject
    latest_email = email_items.order_by('-datetime_received')[0]
    
    # Save the email body to a text file
    with open('email_text.txt', 'w', encoding='utf-8') as file:
        file.write(latest_email.text_body)
    print(f"Email downloaded and saved as 'email_text.txt'")
else:
    print(f"No emails found with the subject: {email_subject}")