
pip install pywin32

import win32com.client as win32
from datetime import datetime
#import os

# Open the Outlook
outlook = win32.Dispatch('outlook.application')

# Create the email
mail = outlook.CreateItem(0)

# Set the email subject
mail.Subject = 'Test' + datetime.now().strftime('%#d %b %Y %H:%M')

# Set the receiver email
mail.To = "ldsantana@sefaz.al.gov.br"

# Add the image
# attachment = mail.Attachments.Add(os.getcwd() + "\\Currencies.png")
# attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "currency_img")

# Write the email content
mail.HTMLBody = r"""
Dear Larissa,<br><br>
Testing if it works:<br><br>"""

'''<img src="cid:currency_img"><br><br>
For more details, you can check the table in the Excel file attached.<br><br>
Best regards,<br>
Yeung
"""'''

# Add the Excel attachement
# mail.Attachments.Add(os.getcwd() + "\\Currencies.xlsx")

# Send the email
mail.Send()
print('The Email is sent.')