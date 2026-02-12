from flask import Flask, render_template, request, send_file
import pandas as pd
import numpy as np
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)

def topsis(input_file, weights, impacts, result_file):
    data = pd.read_excel(input_file)
    # Basic TOPSIS logic (simplified for flow)
    # Ensure your specific TOPSIS math logic is inserted here
    # For now, we save the processed file
    data.to_excel(result_file, index=False)

def send_email(receiver_email, file_path):
    sender_email = os.environ.get('EMAIL_SENDER')
    password = os.environ.get('EMAIL_PASSWORD')

    if not sender_email or not password:
        print("CRITICAL: Environment Variables are missing!")
        return False

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = "TOPSIS Result File"

    try:
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(file_path)}")
            msg.attach(part)

        # Using Port 465 (SSL) for better compatibility with Render
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        # This will now definitely show up in your Render Logs!
        print(f"SMTP ERROR: {e}")
        return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        weights = request.form['weights']
        impacts = request.form['impacts']
        email = request.form['email']

        if file and email:
            input_path = "data.xlsx"
            output_path = "result.xlsx"
            file.save(input_path)

            try:
                topsis(input_path, weights, impacts, output_path)
                if send_email(email, output_path):
                    return "Success! Check your email."
                else:
                    # This message helps us know it's a mail issue, not a math issue
                    return "Calculation done, but email failed. PLEASE CHECK RENDER LOGS NOW."
            except Exception as e:
                return f"Math Error: {e}"

    return render_template('index.html')

if __name__ == '__main__':
    # Render requires host 0.0.0.0
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)