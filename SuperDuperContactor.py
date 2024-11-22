import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pdf2image import convert_from_path
import base64
import os
import re
import time
from dotenv import load_dotenv
import openai
Yourname = "Vincent Kara" #TODO
excel_file_path = r"C:\Users\Vince-Pc\Desktop\code\emailbot\STAGE_INFO.xlsx" #TODO
file_to_attach = r"C:\Users\Vince-Pc\Desktop\code\emailbot\CVVincentKara.pdf" #TODO



def send_email_with_attachment(to_address, subject, body, file_path):
    from_address = os.getenv('EMAIL_USER')   #TODO
    password = os.getenv('EMAIL_PASSWORD')   #TODO
    


    message = MIMEMultipart()
    message["From"] = from_address
    message["To"] = to_address
    message["Subject"] = subject

    message.attach(MIMEText(body, "plain"))

    if os.path.isfile(file_path):
        with open(file_path, "rb") as attachment:
            mime_base = MIMEBase("application", "octet-stream")
            mime_base.set_payload(attachment.read())
        encoders.encode_base64(mime_base)
        mime_base.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
        message.attach(mime_base)
    else:
        print(f"The file {file_path} does not exist.")
        return

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(from_address, password)
        server.sendmail(from_address, to_address, message.as_string())
        print(f"Email sent successfully to {to_address}!")
    except Exception as e:
        print(f"An error occurred while sending to {to_address}: {e}")
    finally:
        server.quit()


def convert_pdf_to_png(pdf_path):
    poppler_path = r"C:\ProgramData\chocolatey\lib\poppler\tools\Library\bin" #TODO
    output_folder = os.path.dirname(pdf_path) 
    output_filename = os.path.splitext(os.path.basename(pdf_path))[0] + ".png"
    output_path = os.path.join(output_folder, output_filename)


    images = convert_from_path(pdf_path, poppler_path=poppler_path)


    images[0].save(output_path, "PNG")
    return output_path


def encode_png_to_base64(png_path):
    with open(png_path, "rb") as png_file:
        base64_string = base64.b64encode(png_file.read()).decode("utf-8")
    return base64_string

def read_emails_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df['Email'].tolist()

def is_valid_email(email):
    if isinstance(email, str):
        email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(email_regex, email) is not None
    return False

def gptsubjet(company_name,base64):
    load_dotenv()
    openai.api_key = os.getenv('OPENAI_API_KEY')#TODO

    try:
        response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
    messages=[
    {
      "role": "user",
      "content": [
        {
          "type": "text",
          "text": f"J'ai besoin le contenu dun email pour soumettre ma canditature a la compagnie {company_name}. Je suis a ma derniere session du cegep qui moblige a faire un stage, il faut presicer que le stage doit etre au alentour de montreal. tu peut dinformer sur mon cv pour ecrire un text. le cv va aussi etre attacher a lemail. Veuillier debuter le sujet de lemail directement sans suject, parler aussi un peu de la compagnie. Seulement mettre mon nom, numero de telephone et courriel a la fin. et joindre le message exact en en anglais ci-dessous. Ne dite pas de message pour lusager mais seulement le body du email Veuillez débuter le body de l'email directement sans objet, parler aussi un peu de la compagnie. Seulement mettre mon nom, numéro de téléphone et courriel à la fin.NE PAS METTRE MON ADDRESSE"
        },{
          "type": "image_url",
          "image_url": {
            "url": f"data:image/png;base64,{base64}"}
        }
      ]
    }
  ],
            temperature=1,
            max_tokens=3999,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        
        return response['choices'][0]['message']['content']
    except Exception as e:
        print("An error occurred:", e)
        return ""





emails = read_emails_from_excel(excel_file_path)
company_names = pd.read_excel(excel_file_path)['Company_Name'].tolist()
png_path = convert_pdf_to_png(file_to_attach)
base64_png = encode_png_to_base64(png_path)

for email, company_name in zip(emails, company_names):
    if pd.notna(email) and is_valid_email(email):
        subject = f"Demande de stage / Internship Application – {Yourname}"
        body = gptsubjet(company_name,base64_png)  
        send_email_with_attachment(email, subject, body, file_to_attach)
        print(f"DONE WITH THIS EMAIL:{email}")
        time.sleep(50000)  # PLS REDUCE - THIS DELAY IS FOR TESTING #TODO
    else:
        print(f"Skipping invalid email: {email}")
        time.sleep(50000)#TODO
