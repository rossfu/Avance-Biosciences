import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Email configuration
sender_email = "nconnector@avancebio.com"
##receiver_email = "michellelee2242@gmail.com"
receiver_email = "ericrossfu93@gmail.com"
subject = "Go Morning Hun <3"
body = "Hanzaaaaaaaaaaaaaaa\nYippeeeeee!!!!!!\nLuv yuuz,\n-Ross"


# SMTP server configuration (for Gmail)
smtp_server = "smtp.office365.com"
smtp_port = 587
smtp_password = ">_1d#2qP9M3]}92m"


# Create the MIME object
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject

# Attach the body
message.attach(MIMEText(body, "plain"))

# Establish a connection with the SMTP server
with smtplib.SMTP(smtp_server, smtp_port) as server:
    # Start the TLS encryption
    server.starttls()
    
    # Login to the email account
    server.login(sender_email, smtp_password)
    
    # Send the email
    server.sendmail(sender_email, receiver_email, message.as_string())

print("Email sent successfully.")
