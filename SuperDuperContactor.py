import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import re
import time
from dotenv import load_dotenv
import openai
Yourname = "Vincent Kara"
excel_file_path = r"C:\Users\Vince-Pc\Desktop\code\emailbot\STAGE_INFO.xlsx"
file_to_attach = r"C:\Users\Vince-Pc\Documents\CV\Vincent_Kara_CV.pdf"
def send_email_with_attachment(to_address, subject, body, file_path):
    from_address = os.getenv('EMAIL_USER')  
    password = os.getenv('EMAIL_PASSWORD')   
    


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

def read_emails_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df['Email'].tolist()

def is_valid_email(email):
    if isinstance(email, str):
        email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(email_regex, email) is not None
    return False

def gptsubjet(company_name):
    load_dotenv()
    openai.api_key = os.getenv('OPENAI_API_KEY')

    try:
        response = openai.ChatCompletion.create(
        model="gpt-4o-mini",
    messages=[
    {
      "role": "user",
      "content": [
        {
          "type": "text",
          "text": f"J'ai besoin le contenu dun email pour soumettre ma canditature a la compagnie {company_name}. Je suis a ma derniere session du cegep qui moblige a faire un stage, il faut presicer que le stage doit etre en francais et au alentour de montreal. tu peut dinformer sur mon cv pour ecrire un text. le cv va aussi etre attacher a lemail. Veuillier debuter le sujet de lemail directement sans suject, parler aussi un peu de la compagnie. Seulement mettre mon nom, numero de telephone et courriel a la fin. et joindre le message exact en en anglais ci-dessous. Ne dite pas de message pour lusager mais seulement le body du email Veuillez débuter le body de l'email directement sans objet, parler aussi un peu de la compagnie. Seulement mettre mon nom, numéro de téléphone et courriel à la fin.NE PAS METTRE MON ADDRESSE"
        },
        {
          "type": "image_url",
          "image_url": {
            "url": "PLS PUT UR CV.png ENCODED IN BASE64 HERE - Will allow chatgpt to read it and use it"}
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

for email, company_name in zip(emails, company_names):
    if pd.notna(email) and is_valid_email(email):
        subject = f"Demande de stage / Internship Application – {Yourname}"
        body = gptsubjet(company_name)  
        send_email_with_attachment(email, subject, body, file_to_attach)
        print(f"DONE WITH THIS EMAIL:{email}")
        time.sleep(500000)  # PLS REDUCE - THIS DELAY IS FOR TESTING
    else:
        print(f"Skipping invalid email: {email}")
        time.sleep(500000)
