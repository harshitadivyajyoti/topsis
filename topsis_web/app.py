import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import numpy as np
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename

app = Flask(__name__)

# --- CONFIGURATION ---
UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"

# 1. Create folders if they don't exist (Fixes the GitHub issue)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

def calculate_topsis(input_file, weights, impacts, output_file):
    try:
        # Read the Excel file
        df = pd.read_excel(input_file)
        
        # Check sufficient columns
        if df.shape[1] < 3:
            raise ValueError("Input file must contain three or more columns.")

        # Extract numeric data (assuming first column is Fund Name/ID)
        data = df.iloc[:, 1:].values.astype(float)
        
        # Check if weights match columns
        if len(weights) != data.shape[1]:
            raise ValueError(f"Number of weights ({len(weights)}) does not match number of columns ({data.shape[1]}).")
        if len(impacts) != data.shape[1]:
            raise ValueError(f"Number of impacts ({len(impacts)}) does not match number of columns ({data.shape[1]}).")

        # Vector Normalization
        rss = np.sqrt(np.sum(data**2, axis=0))
        normalized_data = data / rss

        # Weighted Normalization
        weighted_data = normalized_data * weights

        # Ideal Best and Worst
        ideal_best = []
        ideal_worst = []

        for i in range(data.shape[1]):
            if impacts[i] == '+':
                ideal_best.append(np.max(weighted_data[:, i]))
                ideal_worst.append(np.min(weighted_data[:, i]))
            else: # Impact is '-'
                ideal_best.append(np.min(weighted_data[:, i]))
                ideal_worst.append(np.max(weighted_data[:, i]))
        
        ideal_best = np.array(ideal_best)
        ideal_worst = np.array(ideal_worst)

        # Euclidean Distance
        S_plus = np.sqrt(np.sum((weighted_data - ideal_best)**2, axis=1))
        S_minus = np.sqrt(np.sum((weighted_data - ideal_worst)**2, axis=1))

        # Topsis Score
        # Avoid division by zero
        score = S_minus / (S_plus + S_minus + 1e-9)

        df['Topsis Score'] = score
        df['Rank'] = df['Topsis Score'].rank(ascending=False).astype(int)

        # Save result
        df.to_excel(output_file, index=False)
        return True, "Success"

    except Exception as e:
        return False, str(e)

def send_email(receiver_email, file_path):
    # 2. GET CREDENTIALS SECURELY FROM ENVIRONMENT VARIABLES
    sender_email = os.getenv("EMAIL_SENDER")  # Must be set in Render Env Vars
    sender_password = os.getenv("EMAIL_PASSWORD") # Must be set in Render Env Vars

    if not sender_email or not sender_password:
        print("Error: Email credentials are missing in Environment Variables.")
        return False

    msg = EmailMessage()
    msg['Subject'] = "Your TOPSIS Analysis Result"
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg.set_content("Hello,\n\nPlease find attached the result of your TOPSIS analysis.\n\nBest Regards,\nTopsis Web App")

    # Attach the file
    with open(file_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(file_path)

    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    # Send Email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"Failed to send email: {e}")
        return False

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            # 1. Get Form Data
            if 'file' not in request.files:
                return "No file part"
            
            file = request.files["file"]
            if file.filename == '':
                return "No selected file"

            weights_str = request.form["weights"]
            impacts_str = request.form["impacts"]
            email = request.form["email"]

            # 2. Parse Weights and Impacts
            weights = list(map(float, weights_str.split(",")))
            impacts = impacts_str.split(",")

            # 3. Save Uploaded File
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            output_path = os.path.join(app.config['RESULT_FOLDER'], f"result_{filename}")
            
            file.save(input_path)

            # 4. Run Topsis
            success, message = calculate_topsis(input_path, weights, impacts, output_path)
            
            if not success:
                return f"Error in calculation: {message}"

            # 5. Send Email
            email_sent = send_email(email, output_path)

            if email_sent:
                return "Success! The result has been sent to your email."
            else:
                return "Calculation done, but failed to send email. Check server logs."

        except Exception as e:
            return f"An error occurred: {str(e)}"

    return render_template("index.html")

if __name__ == "__main__":
    # 3. FIX HOST FOR RENDER
    app.run(host="0.0.0.0", port=5000)