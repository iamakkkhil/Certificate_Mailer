from PIL import Image, ImageFont, ImageDraw 
import os
import pandas as pd

# import comtypes.client
import smtplib, ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv

load_dotenv()


def delete_file(filename):
    try:
        os.remove(filename)
    except Exception as e:
        print(e)


def delete_files(names):
    for name in names:
        try:
            os.remove(f"output/{name}_certificate.png")
        except Exception as e:
            print(e)
            pass


def read_csv(filename):
    df = pd.read_csv(filename)
    emails = df.iloc[:, 1]
    names = df.iloc[:, 0]
    return names, emails


def add_name_to_certificate(file_path, names):
    """
    Add a name to a PDF file.
    """
    filepath = os.path.abspath(file_path)
    for name in names:
        image = Image.open(filepath)
        caveat_font = ImageFont.truetype('Caveat/caveat.ttf', 50)
        title_text = name
        image_editable = ImageDraw.Draw(image)
        image_editable.text((60,285), title_text, (0, 0, 0), font=caveat_font)
        image.save(f"output/{name}_certificate.png")


def read_body(filename):
    with open(filename, "r") as f:
        body = f.read()
    return body


def create_email_body(name, email, body):
    subject = "[GDSC I²IT] Congratulations🎉 #30DaysofGoogleCloud "
    body = body
    sender_email = os.environ.get("EMAIL")
    receiver_email = email

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = "manastole.01@gmail.com"  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    filename = f"output/{name}_certificate.png"  # In same directory as script
    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {name} Certificate.png",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()
    return text


def send_mail(names, emails):
    sender_email = os.environ.get("EMAIL")
    password = os.environ.get("PASSWORD")

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        print("Logging In...")
        server.login(sender_email, password)
        print("Logged In...")
        print("Sending mail...")

        body = read_body("assets/email_body.txt")
        for i in range(len(emails)):
            text = create_email_body(names[i], emails[i], body)
            receiver_email = emails[i]
            server.sendmail(sender_email, receiver_email, text)
            print(f"Mail sent successfully to {i} - {names[i]} - {emails[i]}.")


if __name__ == "__main__":
    file_path = "assets/Both_tracks.png"
    user_data_file_path = os.path.abspath("User_details/Both_Track_Winners_Data.csv")
    names, emails = read_csv(user_data_file_path)
    add_name_to_certificate(file_path, names)
    send_mail(names, emails)
    delete_files(names)
