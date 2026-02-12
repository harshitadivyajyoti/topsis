import os
import smtplib
import pandas as pd
import numpy as np
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)

# --- CONFIGURATION ---
UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

def calculate_topsis(input_file, weights, impacts, output_file):
    try:
        df = pd.read_excel(input_file)
        if df.shape[1] < 3:
            return False, "Input file must contain three or more columns."

        data = df.iloc[:, 1:].values.astype(float)
        if len(weights) != data.shape[1] or len(impacts) != data.shape[1]:
            return False, "Weights/Impacts length mismatch."

        # Math Logic
        rss = np.sqrt(np.sum(data**2, axis=0))
        weighted_data = (data / rss) * weights

        ideal_best = []
        ideal_worst = []
        for i in range(data.shape[1]):
            if impacts[i] == '+':
                ideal_best.append(np.max(weighted_data[:, i]))
                ideal_worst.append(np.min(weighted_data[:, i]))
            else:
                ideal_best.append(np.min(weighted_data[:, i]))
                ideal_worst.append(np.max(weighted_data[:, i]))

        S_plus = np.sqrt(np.sum((weighted_data - np.array(ideal_best))**2, axis=1))
        S_minus = np.sqrt(np.sum((weighted_data - np.array(ideal_worst))**2, axis=1))
        df['Topsis Score'] = S_minus / (S_plus + S_minus + 1e-9)
        df['Rank'] = df['Topsis Score'].rank(ascending=False).astype(int)

        df.to_excel(output_file, index=False)
        return True, "Success"
    except Exception as e:
        return False, str(e)

def send_email(receiver_email, file_path):
    sender_email = os.environ.get('EMAIL_SENDER')
    password = os.environ.get('EMAIL_PASSWORD')

    if not sender_email or not password:
        print("CRITICAL: Environment Variables are missing!")
        return False

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = "TOPSIS Analysis Results"

    try:
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
            msg.attach(part)

        # Port 465 is more stable on Render
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"SMTP ERROR: {e}")
        return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            file = request.files.get('file')
            weights_str = request.form.get('weights')
            impacts_str = request.form.get('impacts')
            email = request.form.get('email')

            if not all([file, weights_str, impacts_str, email]):
                return "Missing form data!"

            weights = [float(w) for w in weights_str.split(',')]
            impacts = impacts_str.split(',')

            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            output_path = os.path.join(app.config['RESULT_FOLDER'], f"result_{filename}")
            file.save(input_path)

            success, message = calculate_topsis(input_path, weights, impacts, output_path)
            if not success:
                return f"Math Error: {message}"

            if send_email(email, output_path):
                return "Success! Check your email."
            else:
                return "Calculation done, but email failed. CHECK RENDER LOGS."

        except Exception as e:
            return f"Error: {str(e)}"

    return render_template('index.html')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)