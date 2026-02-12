from flask import Flask, render_template, request
import os
import smtplib
from email.message import EmailMessage
import pandas as pd
import numpy as np

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)


def topsis(input_file, weights, impacts, output_file):
    df = pd.read_excel(input_file)

    if df.shape[1] < 3:
        raise ValueError("Input file must contain three or more columns.")

    data = df.iloc[:, 1:].values

    if not np.issubdtype(data.dtype, np.number):
        raise ValueError("All criteria columns must be numeric.")

    weights = np.array(weights)
    impacts = np.array(impacts)

    # Normalization
    norm = data / np.sqrt((data ** 2).sum(axis=0))

    # Weighted normalization
    weighted = norm * weights

    # Ideal best and worst
    ideal_best = np.max(weighted, axis=0)
    ideal_worst = np.min(weighted, axis=0)

    for i in range(len(impacts)):
        if impacts[i] == '-':
            ideal_best[i], ideal_worst[i] = ideal_worst[i], ideal_best[i]

    # Distance
    dist_best = np.sqrt(((weighted - ideal_best) ** 2).sum(axis=1))
    dist_worst = np.sqrt(((weighted - ideal_worst) ** 2).sum(axis=1))

    score = dist_worst / (dist_best + dist_worst)
    df["Topsis Score"] = score
    df["Rank"] = df["Topsis Score"].rank(ascending=False)

    df.to_excel(output_file, index=False)


def send_email(receiver_email, file_path):
    sender_email = "harshitadivyajyoti@gmail.com"
    sender_password = "rjabpmlqnselfenf"

    msg = EmailMessage()
    msg["Subject"] = "TOPSIS Result"
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.set_content("Please find attached the TOPSIS result file.")

    with open(file_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(file_path)

    msg.add_attachment(file_data, maintype="application",
                       subtype="octet-stream", filename=file_name)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        weights = request.form["weights"]
        impacts = request.form["impacts"]
        email = request.form["email"]

        weights = list(map(float, weights.split(",")))
        impacts = impacts.split(",")

        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        output_path = os.path.join(RESULT_FOLDER, "result.xlsx")

        file.save(input_path)

        topsis(input_path, weights, impacts, output_path)
        send_email(email, output_path)

        return "Result sent to your email!"

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)