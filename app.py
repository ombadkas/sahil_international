from flask import Flask, render_template, request
import openpyxl
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import threading

app = Flask(__name__)

def send_email_background(name, phone, user_email, message):
    try:
        sender_email = 'ombadkas03@gmail.com'
        sender_password = 'hfga qapm vckn nxhy'
        recipient_email = 'ombadkas03@gmail.com'
        subject = "New Enquiry from Website Contact Form"

        body = f"""
You have received a new enquiry:

Name: {name}
Phone: {phone}
Email: {user_email}
Message: {message}
        """

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()

        print("📨 Email sent successfully.")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/run-python', methods=['POST'])
def run_python():
    try:
        data = request.get_json()

        name = data.get('name')
        phone = data.get('phone')
        user_email = data.get('email')
        message = data.get('message')

        # Save to Excel
        file_path = 'contacts.xlsx'
        if not os.path.exists(file_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Name", "Phone", "Email", "Message"])
        else:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active

        ws.append([name, phone, user_email, message])
        wb.save(file_path)

        # Start email thread
        threading.Thread(
            target=send_email_background,
            args=(name, phone, user_email, message)
        ).start()


        return "✅ Your request was saved! We'll contact you shortly.", 200

    except Exception as e:
        return "Something went wrong!", 200


if __name__ == '__main__':
    app.run(debug=True)
