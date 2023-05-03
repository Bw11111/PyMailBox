from email.message import EmailMessage
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



import markdown
class OutlookMessage:
  def __init__(self, sender, recipient, message, password, subject):
    return self
def SendMailOutlook(sender, recipient, message, password, subject):
  """
  PyMail.SendMailOutlook(sender, recipient, message, password)
  sender = Person sending message
  recipient = Person getting message
  message = message
  password = sender outlook password
  subject = email subject
  """
  #email = EmailMessage()
  email=MIMEMultipart('alternative')
  text = "Open this message in a client that supports html messages!"
  html = f"""\
  <html>
    <head></head>
    <body>
      {message}
    </body>
  </html>
  <center><strong style="font-size:5px;">This message was sent using bw1111's PyMailbox python library</strong></center>
  """
  
  # Record the MIME types of both parts - text/plain and text/html.
  part1 = MIMEText(text, 'plain')
  part2 = MIMEText(html, 'html')

  email.attach(part1)
  email.attach(part2)
  email["From"] = sender
  email["To"] = recipient
  email["Subject"] = subject

  smtp = smtplib.SMTP("smtp-mail.outlook.com", port=587)
  smtp.starttls()
  smtp.login(sender, password)
  smtp.sendmail(sender, recipient, email.as_string())
  smtp.quit()


