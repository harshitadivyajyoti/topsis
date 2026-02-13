# TOPSIS Implementation and Web Service

Author: Harshita  
Roll Number: 102317208  

Project Overview:
This project implements the TOPSIS (Technique for Order Preference by Similarity to Ideal Solution) method
for multi-criteria decision making. It includes command line tools, a PyPI package, and a web service.

Part I – Command Line TOPSIS

Usage:
python topsis.py <InputFile> <Weights> <Impacts> <OutputFile>
python topsis.py data.xlsx "1,1,1,2" "+,+,-,+" result.xlsx

Input Rules:
Minimum Columns: 3 (First column is Alternatives/Names)
Numeric Data: All columns from the 2nd onwards must be numeric
Weights/Impacts: Must match the number of numeric columns
Impacts: Must be either + (Beneficial) or - (Non-beneficial)
Format: Comma-separated strings for weights and impacts

Part II – PyPI Package

Package Name: topsis-harshita-102317208
Install & Run:
pip install topsis-harshita-102317208
topsis data.xlsx "1,1,1,2" "+,+,-,+" result.xlsx
PyPI Link: https://pypi.org/project/topsis-harshita-102317208/

Part III – Web Service

Live Demo: https://topsis-faef.onrender.com

Features:
Excel Upload: Supports .xlsx and .csv file processing
Parameter Validation: Ensures weights and impacts are correctly formatted before processing
Automated Results: Calculates scores and ranks, then attempts to email the result file

Technical Note on Email Delivery:
In the cloud production environment (Render), Gmail SMTP may be restricted due to "Suspicious Login" flags
Validation: If the app displays "Calculation done, but failed to send email", the TOPSIS logic executed correctly

Run Locally:
git clone <your-repo-link>
cd topsis_web
pip install -r requirements.txt
python app.py
Open your browser at http://127.0.0.1:5000/

Technologies Used:
Language: Python
Data Analysis: Pandas, NumPy
Web Framework: Flask
Deployment: Render
DevOps: Git, PyPI, SMTP

Outcome:
A comprehensive end-to-end decision-making tool that demonstrates:
Advanced mathematical modeling in Python
Building and maintaining public developer tools (PyPI)
Developing and deploying full-stack web applications with cloud integration

Screenshots:
![image](https://github.com/user-attachments/assets/f4b19494-b464-4640-b9d0-5e0fe3d4b1b3)
![image](https://github.com/user-attachments/assets/ce018773-4aec-44a1-a99a-823bd13afabf)
