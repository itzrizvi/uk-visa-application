import os
from flask import Flask, request, render_template, redirect,session, url_for
from weasyprint import HTML
import ssl
import smtplib
import base64
import argparse
import os.path
import shutil
import subprocess
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Access environment variables
EMAIL_SENDER = os.getenv('EMAIL_SENDER')
EMAIL_APP_PASSWORD = os.getenv('EMAIL_APP_PASSWORD')
EMAIL_RECEIVER = os.getenv('EMAIL_RECEIVER')

app = Flask(__name__)

# Retrieve the secret key from the environment variable
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY')

# Ensure that a secret key is set
if not app.config['SECRET_KEY']:
    raise ValueError("No SECRET_KEY set for Flask application. Set the SECRET_KEY environment variable.")


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/second-form', methods=['GET','POST'])
def second_form():
    if request.method == 'POST':
        # Process the form data if POST request is received
        return redirect(url_for('submit_second_form'))
    else:
        form_data = request.args.to_dict()
        session["data_from_index"] = form_data
        # Render the second_form.html template for GET requests
        return render_template('second_form.html')
    

@app.route('/submit-form', methods=['POST'])
def submit_form():
    second_form_data = request.form
    previously_merged_form_data = session.get("data_from_index",{})
    whole_form_data = {**second_form_data, **previously_merged_form_data}
    pdf_path = generate_pdf(whole_form_data)  # Call the function to generate the PDF
    send_email_with_attachment(pdf_path)  # Call the function to send email with PDF attachment
    return 'Form submitted successfully!'


def generate_pdf(form_data):
    # Convert the image to a base64-encoded string
    with open('UKVisas.png', 'rb') as f:
        image_data = base64.b64encode(f.read()).decode('utf-8')
    html_template = render_template('pdf_template.html', form_data=form_data, image_data=image_data)
    pdf_path = 'temporary-output.pdf'
    # Generate PDF using WeasyPrint
    HTML(string=html_template).write_pdf(pdf_path)
    compress(pdf_path, "uk-visa-sponsorship-application.pdf", power=3)
    return "uk-visa-sponsorship-application.pdf"

def send_email_with_attachment(pdf_path):
    sender_email = EMAIL_SENDER
    receiver_email = EMAIL_RECEIVER
    subject = 'UK Visa Application Form'
    body = 'Please find attached the UK visa application form.'

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))
    context = ssl.create_default_context();

    with open(pdf_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {pdf_path}')
        msg.attach(part)

    smtp_server = 'smtp.gmail.com'
    port = 465
    email_sender = EMAIL_SENDER
    email_password = EMAIL_APP_PASSWORD

    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(email_sender, email_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())

    print('Email sent successfully!')


def compress(input_file_path, output_file_path, power=0):
    """Function to compress PDF via Ghostscript command line interface"""
    quality = {
        0: "/default",
        1: "/prepress",
        2: "/printer",
        3: "/ebook",
        4: "/screen"
    }
    # Additional options for fine-tuning compression
    additional_options = {
        0: "",
        1: "-dColorImageDownsampleType=/Bicubic",
        2: "-dColorImageDownsampleType=/Bicubic",
        3: "-dColorImageDownsampleType=/Bicubic",
        4: "-dColorImageDownsampleType=/Bicubic"
    }

    # Basic controls
    # Check if valid path
    if not os.path.isfile(input_file_path):
        print("Error: invalid path for input PDF file.", input_file_path)
        sys.exit(1)

    # Check compression level
    if power < 0 or power > len(quality) - 1:
        print("Error: invalid compression level, run pdfc -h for options.", power)
        sys.exit(1)

    # Check if file is a PDF by extension
    if input_file_path.split('.')[-1].lower() != 'pdf':
        print(f"Error: input file is not a PDF.", input_file_path)
        sys.exit(1)

    gs = get_ghostscript_path()
    print("Compress PDF...")
    initial_size = os.path.getsize(input_file_path)
    subprocess.call(
        [
            gs,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS={}".format(quality[power]),
            additional_options[power],
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            "-sOutputFile={}".format(output_file_path),
            input_file_path,
        ]
    )
    final_size = os.path.getsize(output_file_path)
    ratio = 1 - (final_size / initial_size)
    print("Compression by {0:.0%}.".format(ratio))
    print("Final file size is {0:.2f}KB".format(final_size / 1024))
    print("Done.")


def get_ghostscript_path():
    gs_names = ["gs", "gswin32", "gswin64"]
    for name in gs_names:
        if shutil.which(name):
            return shutil.which(name)
    raise FileNotFoundError(
        f"No GhostScript executable was found on path ({'/'.join(gs_names)})"
    )



if __name__ == '__main__':
    app.run(debug=True, port=8000)
