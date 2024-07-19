IOCs Extractor
This Python script is designed to extract Indicators of Compromise (IOCs) such as IP addresses, domains, URLs, and file names from an email and save them to a text file. Additionally, it provides the option to create a zip file containing these extracted IOCs for easier distribution and analysis.

Usage
Prerequisites:

Make sure you have Python 3.x installed on your system.
Install the required Python packages using the following command: pip install exchangelib requests xlwt xlrd
Configuration:

Open the script and set your Outlook email address and password in the email and password variables, respectively.
Run the Script:

Execute the script to connect to your Outlook account and search for emails with a specific subject (defined by email_subject variable).
Extraction:

The script will download the latest email matching the subject and extract IOCs (IP addresses, domains, URLs, and file names) from the email's text body.
IOCs Text File:

The extracted IOCs will be saved in a file named email_text.txt.
Zip File:

The script will create a zip file named IOCs.zip containing the extracted IOCs.
Example
```python

Configuration
email = 'your_email@example.com' password = 'your_password' email_subject = 'as per choice'

How to run: python IOC_extract.py
